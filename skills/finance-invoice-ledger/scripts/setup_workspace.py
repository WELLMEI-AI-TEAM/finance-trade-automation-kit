"""Portable, network-free finance workspace initialization and local binding."""
from __future__ import annotations

import argparse
import json
import sys
import uuid
from pathlib import Path
from urllib.parse import urlsplit
from workspace_runtime import prepare_runtime

BINDING = Path(__file__).resolve().parents[1]/'binding.local.json'
FOLDERS = ('待处理', '发票', '台账文件', '异常未处理发票', '备份', '.运行记录')


class SetupRequired(ValueError):
    def __init__(self, status, root=None):
        self.status = status
        self.root = root
        super().__init__('尚未绑定工作区' if status == 'workspace_setup_required' else '工作区尚未配置财务账户')


def setup_message(status='workspace_setup_required', root=None):
    message = {'completed': False, 'status': status, 'invoice_processing_started': False}
    if status == 'workspace_setup_required':
        message['question'] = '请问要在哪个文件夹下建立“台账处理”？请提供完整路径；我会在其中创建待处理、发票、台账文件、异常未处理发票、备份等目录。'
    else:
        message.update(workspace_root=str(root), config=str(Path(root)/'.运行记录/财务配置.json'),
            question='目录已建立。请提供此电脑使用的财务台账服务地址、财务用户名及凭据文件路径，或指定已有的本机财务配置。')
    return message


def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name+'.'+uuid.uuid4().hex+'.tmp')
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    temp.replace(path)


def load_binding(binding=BINDING):
    binding = Path(binding)
    if not binding.is_file(): raise SetupRequired('workspace_setup_required')
    data = json.loads(binding.read_text(encoding='utf-8-sig'))
    value = data.get('workspace_root')
    if not value: raise SetupRequired('workspace_setup_required')
    root = Path(value)
    if not root.is_absolute(): raise ValueError('工作区绑定必须是完整路径')
    return root.resolve()


def require_account(config, root):
    if not config.get('base_url') or not config.get('username'):
        raise SetupRequired('account_setup_required', root)


def initialize(*, parent=None, root=None, binding=BINDING, base_url=None, username=None,
               credentials_file=None, password_env=None):
    if (parent is None) == (root is None): raise ValueError('请指定父目录或已有工作区根目录，二选一')
    selected = Path(parent if parent is not None else root)
    if not selected.is_absolute() or not selected.is_dir():
        raise ValueError('请选择已经存在的文件夹完整路径')
    selected = selected.resolve()
    root = selected/'台账处理' if parent is not None else selected
    root = root.resolve()
    if parent is not None and (root == selected or not root.is_relative_to(selected)):
        raise ValueError('台账处理目录指向所选父目录之外；请明确选择实际根目录')
    binding = Path(binding)
    if binding.exists() and load_binding(binding) != root:
        raise ValueError('此技能已绑定其他工作区；请显式使用该工作区，不自动替换已有绑定')
    target = root/'.运行记录/财务配置.json'
    config = None
    if target.exists():
        config = json.loads(target.read_text(encoding='utf-8-sig'))
        if config.get('expected_department') != 'finance': raise ValueError('已有配置不是财务账户配置')
        value = Path(config.get('workspace_root', ''))
        if not value.is_absolute() or value.resolve() != root:
            raise ValueError('已有配置与工作区不匹配，保留原件')
    for folder in FOLDERS:
        path = root/folder
        if path.exists() and (not path.is_dir() or not path.resolve().is_relative_to(root)):
            raise ValueError('业务子目录被文件占用或指向工作区外，未初始化')
    provided = any(value is not None for value in (base_url, username, credentials_file, password_env))
    changes = {}
    if provided:
        if not base_url or not username or bool(credentials_file) == bool(password_env):
            raise ValueError('配置账户需要服务地址、用户名，以及凭据文件或密码环境变量名（二选一）')
        url = urlsplit(base_url)
        if url.scheme not in ('http', 'https') or not url.hostname or url.username or url.password:
            raise ValueError('服务地址必须是无内嵌凭据的HTTP或HTTPS地址')
        if url.query or url.fragment: raise ValueError('服务地址不能包含查询参数或片段')
        changes = {'base_url': base_url.rstrip('/'), 'username': username}
        if credentials_file:
            credential = Path(credentials_file)
            if not credential.is_absolute() or not credential.is_file(): raise ValueError('凭据文件必须是本机已有文件的完整路径')
            changes['credentials_file'] = str(credential.resolve())
        else:
            if not password_env.isidentifier(): raise ValueError('密码环境变量名无效')
            changes['password_env'] = password_env
        if config and config.get('username') and (
                config.get('username') != username or config.get('base_url') != changes['base_url']):
            raise ValueError('工作区已有账户；不能用初始化操作更换账户或服务')
    new_config = config is None
    if new_config:
        config = {'workspace_root': str(root), 'expected_department': 'finance',
                  'base_url': '', 'username': '', 'archive_month_layout': 'year.month',
                  'wait_seconds': 900, 'poll_seconds': 2}
    if changes:
        config.pop('credentials_file', None); config.pop('password_env', None)
        config.update(changes)
    # Keep this setup script stdlib-only; the workflow validates its full layout.
    for folder in FOLDERS: (root/folder).mkdir(parents=True, exist_ok=True)
    prepare_runtime(root)
    if new_config or changes: write_json(target, config)
    if not binding.exists(): write_json(binding, {'workspace_root': str(root)})
    return {'status': 'configured' if config.get('base_url') and config.get('username') else 'account_setup_required',
            'workspace_root': str(root), 'config': str(target),
            'created_directories': [str(root/name) for name in FOLDERS],
            'invoice_processing_started': False, 'service_connection_tested': False}


def main(argv=None):
    parser = argparse.ArgumentParser(description='按用户选择建立财务台账处理目录；不处理发票、不创建账户')
    paths = parser.add_mutually_exclusive_group(required=True)
    paths.add_argument('--parent'); paths.add_argument('--root')
    parser.add_argument('--service-url'); parser.add_argument('--username')
    credentials = parser.add_mutually_exclusive_group()
    credentials.add_argument('--credentials-file'); credentials.add_argument('--password-env')
    args = parser.parse_args(argv)
    result = initialize(parent=args.parent, root=args.root, base_url=args.service_url,
        username=args.username, credentials_file=args.credentials_file, password_env=args.password_env)
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'): stream.reconfigure(encoding='utf-8')
    try:
        raise SystemExit(main())
    except (ValueError, OSError):
        print(json.dumps({'completed': False, 'status': 'setup_failed', 'invoice_processing_started': False,
            'reason': '无法初始化：请检查路径、已有绑定或账户配置；不会覆盖已有财务文件。'}, ensure_ascii=False))
        raise SystemExit(1)
