import importlib.util
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).parents[1]
SCRIPTS = ROOT/'skills/finance-invoice-ledger/scripts'


def setup_module_file():
    path = SCRIPTS/'setup_workspace.py'
    assert path.exists(), 'portable workspace initializer is missing'
    spec = importlib.util.spec_from_file_location('finance_setup_under_test', path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_selected_parent_creates_workspace_without_using_another_computer_account(tmp_path):
    mod = setup_module_file()
    binding = tmp_path/'skill/binding.local.json'
    result = mod.initialize(parent=tmp_path, binding=binding)
    root = tmp_path/'台账处理'
    assert result['workspace_root'] == str(root)
    assert result['status'] == 'account_setup_required'
    assert not result['invoice_processing_started']
    assert all((root/name).is_dir() for name in ['待处理','发票','台账文件','异常未处理发票','备份','.运行记录'])
    cfg = json.loads((root/'.运行记录/财务配置.json').read_text(encoding='utf-8'))
    assert cfg['workspace_root'] == str(root)
    assert cfg['archive_month_layout'] == 'year.month'
    assert not cfg.get('username') and not cfg.get('credentials_file')
    assert json.loads(binding.read_text(encoding='utf-8'))['workspace_root'] == str(root)


def test_repeat_setup_preserves_existing_configuration_and_all_financial_files(tmp_path):
    mod = setup_module_file(); binding = tmp_path/'binding.local.json'
    mod.initialize(parent=tmp_path, binding=binding)
    root = tmp_path/'台账处理'; cfg = root/'.运行记录/财务配置.json'
    data = json.loads(cfg.read_text(encoding='utf-8'))
    data.update(base_url='http://127.0.0.1:18081', username='own_finance', password_env='TEST_FINANCE_PASSWORD')
    cfg.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
    invoice = root/'发票/2026.09/原件.pdf'; invoice.parent.mkdir(); invoice.write_bytes(b'original')
    before = cfg.read_bytes()
    result = mod.initialize(parent=tmp_path, binding=binding)
    assert result['status'] == 'configured'
    assert cfg.read_bytes() == before and invoice.read_bytes() == b'original'
    assert not (root/'台账处理').exists()


def test_explicit_existing_root_is_not_nested_again(tmp_path):
    mod = setup_module_file(); root = tmp_path/'已有工作区'; root.mkdir()
    result = mod.initialize(root=root, binding=tmp_path/'binding.json')
    assert result['workspace_root'] == str(root) and not (root/'台账处理').exists()


def test_service_account_binding_uses_only_explicit_credentials_location(tmp_path):
    mod = setup_module_file(); credential = tmp_path/'my-credentials.txt'; credential.write_text('finance secret')
    result = mod.initialize(parent=tmp_path, binding=tmp_path/'binding.json',
        base_url='http://127.0.0.1:18081', username='finance', credentials_file=credential)
    cfg = json.loads(Path(result['config']).read_text(encoding='utf-8'))
    assert cfg['username'] == 'finance' and cfg['credentials_file'] == str(credential)
    assert 'secret' not in json.dumps(result)
    assert result['status'] == 'configured'


def test_other_binding_is_protected_before_creating_a_new_root(tmp_path):
    mod = setup_module_file(); binding = tmp_path/'binding.json'
    binding.write_text(json.dumps({'workspace_root': str(tmp_path/'old')}))
    with pytest.raises(ValueError, match='绑定'):
        mod.initialize(parent=tmp_path, binding=binding)
    assert not (tmp_path/'台账处理').exists()


def test_foreign_department_config_is_not_adopted_or_overwritten(tmp_path):
    mod = setup_module_file(); root = tmp_path/'台账处理'; folder = root/'.运行记录'; folder.mkdir(parents=True)
    cfg = folder/'财务配置.json'; cfg.write_text(json.dumps({'workspace_root': str(root), 'expected_department':'cashier'}))
    before = cfg.read_bytes()
    with pytest.raises(ValueError, match='财务'):
        mod.initialize(parent=tmp_path, binding=tmp_path/'binding.json')
    assert cfg.read_bytes() == before and not (tmp_path/'binding.json').exists()


@pytest.mark.parametrize('parent_kind', ['relative', 'file', 'missing'])
def test_bad_parent_is_rejected_without_creating_folders(tmp_path, parent_kind):
    mod = setup_module_file(); parent = tmp_path/'not-a-directory'
    if parent_kind == 'relative': parent = Path('relative-parent')
    elif parent_kind == 'file': parent.write_bytes(b'keep')
    with pytest.raises(ValueError): mod.initialize(parent=parent, binding=tmp_path/'binding.json')


@pytest.mark.skipif(sys.platform != 'win32', reason='Windows directory junction regression')
def test_parent_cannot_redirect_workspace_through_existing_directory_junction(tmp_path):
    import _winapi
    mod = setup_module_file(); selected = tmp_path/'selected'; outside = tmp_path/'other'
    selected.mkdir(); outside.mkdir()
    _winapi.CreateJunction(str(outside), str(selected/'台账处理'))
    with pytest.raises(ValueError, match='目录'):
        mod.initialize(parent=selected, binding=tmp_path/'binding.json')
    assert not (outside/'.运行记录').exists() and not (tmp_path/'binding.json').exists()


@pytest.mark.parametrize('package,entry', [('skills','run_finance.py')])
def test_clean_install_requests_directory_instead_of_processing_some_other_workspace(tmp_path, package, entry):
    target = tmp_path/'finance-invoice-ledger/scripts'
    shutil.copytree(ROOT/package/'finance-invoice-ledger/scripts', target, ignore=shutil.ignore_patterns('__pycache__'))
    process = subprocess.run([sys.executable, str(target/entry)], capture_output=True, text=True, encoding='utf-8')
    assert process.returncode == 3, process.stdout + process.stderr
    result = json.loads(process.stdout)
    assert result['status'] == 'workspace_setup_required'
    assert result['question'] and not result['invoice_processing_started']
    assert 'workspace_root' not in result


@pytest.mark.parametrize('package,entry', [('skills','run_finance.py')])
def test_initialized_without_account_reports_account_setup_without_network(tmp_path, package, entry):
    target = tmp_path/'finance-invoice-ledger/scripts'
    shutil.copytree(ROOT/package/'finance-invoice-ledger/scripts', target, ignore=shutil.ignore_patterns('__pycache__'))
    setup = subprocess.run([sys.executable, str(target/'setup_workspace.py'), '--parent', str(tmp_path)],
        capture_output=True, text=True, encoding='utf-8')
    assert setup.returncode == 0, setup.stdout + setup.stderr
    process = subprocess.run([sys.executable, str(target/entry)], capture_output=True, text=True, encoding='utf-8')
    assert process.returncode == 3, process.stdout + process.stderr
    data = json.loads(process.stdout)
    assert data['status'] == 'account_setup_required' and not data['invoice_processing_started']
