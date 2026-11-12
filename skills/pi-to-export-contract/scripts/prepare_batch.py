"""Create one export-contract batch and verify a complete copy of its inputs.

This helper does not classify documents or generate contracts. The caller must
finish document checks before publishing any file under the contract directory.
"""
import argparse
import hashlib
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path


def sha256(path):
    result = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            result.update(chunk)
    return result.hexdigest()


def linked(path):
    return path.is_symlink() or getattr(path, 'is_junction', lambda: False)()


def contained(path, parent):
    return path == parent or parent in path.parents


def write_json(path, value):
    with path.open('x', encoding='utf-8') as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2)


def fail_walk(error):
    raise error


def prepare_batch(root, inputs=None, stamp=None):
    root = Path(root).absolute()
    if linked(root):
        raise ValueError('工作根目录不能是链接或目录联接')
    root = root.resolve()
    managed = [root / '处理结果', root / '备份']
    pending_root = root / '待核对'
    if any(linked(path) for path in managed + [pending_root]):
        raise ValueError('处理结果、备份和待核对目录不能是链接或目录联接')
    sources = [Path(path).absolute() for path in inputs] if inputs else [root / '待处理']
    files, seen = [], set()

    def add_file(path, destination):
        if linked(path):
            raise ValueError(f'输入文件是链接，需明确提供实际文件：{path}')
        actual = path.resolve()
        if actual not in seen:
            if not actual.is_file():
                raise ValueError(f'输入不是普通文件：{path}')
            seen.add(actual)
            files.append((actual, destination))

    for index, source in enumerate(sources, 1):
        if linked(source):
            raise ValueError(f'输入是链接或目录联接：{source}')
        source = source.resolve(strict=True)
        if any(contained(source, parent) for parent in managed):
            raise ValueError(f'不能将历史结果或备份作为当前原件：{source}')
        # An explicit retry may read the pending directory. Root scans exclude it.
        excluded = managed + ([] if contained(source, pending_root) else [pending_root])
        label = Path(f'{index:02d}_{source.name}')
        if source.is_file():
            add_file(source, label / source.name)
            continue
        for directory, names, filenames in os.walk(source, followlinks=False, onerror=fail_walk):
            directory = Path(directory)
            retained = []
            for name in sorted(names):
                child = directory / name
                if any(contained(child.resolve(), parent) for parent in excluded):
                    continue
                if linked(child):
                    raise ValueError(f'输入目录含链接或目录联接：{child}')
                retained.append(name)
            names[:] = retained
            for filename in sorted(filenames):
                path = directory / filename
                add_file(path, label / path.relative_to(source))
    if not files:
        raise ValueError('本次输入没有文件，不建立空批次')
    if stamp is None:
        stamp = datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d_%H%M%S')
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}_\d{6}', stamp):
        raise ValueError('批次时间格式应为 YYYY-MM-DD_HHMMSS')
    for parent in managed:
        parent.mkdir(parents=True, exist_ok=True)
    for sequence in range(1, 10000):
        batch = stamp if sequence == 1 else f'{stamp}_{sequence:02d}'
        result_dir, backup_dir = [parent / batch for parent in managed]
        if backup_dir.exists() or (pending_root / batch).exists():
            continue
        try:
            result_dir.mkdir()
        except FileExistsError:
            continue
        try:
            backup_dir.mkdir()
        except FileExistsError:
            result_dir.rmdir()  # Only the empty directory reserved by this call.
            continue
        break
    else:
        raise ValueError('无法分配唯一批次目录')

    manifest = {'batch': batch, 'root': str(root), 'result_dir': str(result_dir),
                'backup_dir': str(backup_dir), 'pending_dir': str(pending_root / batch), 'backup_status': 'pending',
                'processing_status': 'not_started', 'files': []}
    try:
        for name in ['出口合同', 'PI', '购货合同']:
            (result_dir / name).mkdir()
        for source, relative in files:
            target = backup_dir / '原文件' / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            copied = hashlib.sha256()
            with source.open('rb') as incoming, target.open('xb') as outgoing:
                for chunk in iter(lambda: incoming.read(1024 * 1024), b''):
                    outgoing.write(chunk)
                    copied.update(chunk)
            expected = copied.hexdigest()
            if sha256(source) != expected or sha256(target) != expected:
                raise ValueError(f'备份校验失败或原文件在复制期间变化：{source}')
            manifest['files'].append({'source': str(source), 'backup': str(target),
                                      'sha256': expected, 'bytes': target.stat().st_size})
        # Catch changes to earlier inputs while later files were being copied.
        for record in manifest['files']:
            if sha256(record['source']) != record['sha256']:
                raise ValueError(f'备份期间原文件变化：{record["source"]}')
        manifest['backup_status'] = 'verified'
        write_json(backup_dir / '备份清单.json', manifest)
        write_json(result_dir / '批次记录.json', manifest)
    except Exception as error:
        manifest['backup_status'] = 'failed'
        manifest['error'] = str(error)
        write_json(backup_dir / '备份失败.json', manifest)
        raise
    return manifest


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--request', required=True, help='UTF-8 JSON: root, optional inputs (absolute paths)')
    args = parser.parse_args()
    request = json.loads(Path(args.request).read_text(encoding='utf-8-sig'))
    if not Path(request['root']).is_absolute():
        raise ValueError('root 必须为绝对路径')
    if 'inputs' in request and (not isinstance(request['inputs'], list) or not request['inputs']):
        raise ValueError('inputs 必须为非空路径列表；默认待处理模式请省略该字段')
    if any(not Path(path).is_absolute() for path in request.get('inputs', [])):
        raise ValueError('输入必须为绝对路径')
    print(json.dumps(prepare_batch(request['root'], request.get('inputs')), ensure_ascii=False, indent=2))


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f'批次准备失败：{exc}', file=sys.stderr)
        sys.exit(1)
