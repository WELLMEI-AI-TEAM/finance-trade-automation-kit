"""Time-grouped exception originals with resumable legacy relocation."""
from datetime import datetime, timedelta, timezone
from pathlib import Path
import json
import re

from run_invoice_workflow import atomic_write, digest

FOLDER_PATTERN = r'\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}(?:_\d+)?'


def exception_folder(paths):
    if '_exception_folder' not in paths:
        root = paths['exception_dir']
        stamp = datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d_%H-%M-%S')
        number = 1
        while True:
            folder = root/(stamp if number == 1 else f'{stamp}_{number:02d}')
            try:
                folder.mkdir()
                paths['_exception_folder'] = folder
                break
            except FileExistsError: number += 1
    return paths['_exception_folder']


def exception_target(record, paths, filename):
    if record.get('exception_path'):
        target = Path(record['exception_path']).resolve()
        if not target.is_relative_to(paths['exception_dir']) or target == paths['exception_dir']:
            raise ValueError('异常副本路径超出指定异常目录，保留原件')
        return target
    return exception_folder(paths)/filename


def organize_legacy_exceptions(paths, state, save):
    """Copy first, persist paths, update queues, then remove verified old copies."""
    root = paths['exception_dir']
    jobs = state.get('exception_relocations', [])
    planned = {job['sha256'] for job in jobs}
    candidates = []
    for sha, record in state['sources'].items():
        if record.get('status') != 'exception' or not record.get('exception_path') or sha in planned: continue
        old = Path(record['exception_path']).resolve()
        if old.parent != root: continue
        if not old.is_file() or digest(old.read_bytes()) != sha:
            raise RuntimeError('历史异常副本缺失或已被人工修改，保留原件，停止自动整理')
        note = old.with_suffix(old.suffix+'.原因.txt')
        candidates.append((sha, old, digest(note.read_bytes()) if note.exists() else None))
    if candidates:
        folder = exception_folder(paths)
        jobs += [{'sha256':sha, 'old':str(old), 'new':str(folder/old.name), 'note_sha256':note_sha}
                 for sha, old, note_sha in candidates]
        state['exception_relocations'] = jobs
        try:save()
        except Exception:
            try:folder.rmdir()
            except OSError:pass
            raise
    if not jobs: return 0

    def checked_copy(old, new, expected):
        content=None
        if old.exists():
            if not old.is_file():raise RuntimeError('整理期间异常原件不再是文件，停止整理')
            content=old.read_bytes()
            if digest(content)!=expected:raise RuntimeError('整理期间异常文件被修改，保留所有副本并停止')
        if new.exists():
            if not new.is_file() or digest(new.read_bytes()) != expected:
                raise RuntimeError('时间文件夹中存在内容不同的同名文件，停止覆盖')
        elif content is not None: atomic_write(new, content)
        else: raise RuntimeError('异常原件及整理副本均缺失，停止自动整理')

    mapping = {}; removals = []
    for job in jobs:
        old = Path(job['old']).resolve(); new = Path(job['new']).resolve()
        if (old.parent != root or new.parent.parent != root
                or not re.fullmatch(FOLDER_PATTERN,new.parent.name)):
            raise ValueError('异常目录整理记录路径不合法，未删除任何文件')
        checked_copy(old,new,job['sha256'])
        mapping[str(old)] = str(new); removals.append((old,job['sha256']))
        if job['note_sha256'] is not None:
            old_note=old.with_suffix(old.suffix+'.原因.txt');new_note=new.with_suffix(new.suffix+'.原因.txt')
            checked_copy(old_note,new_note,job['note_sha256'])
            mapping[str(old_note)]=str(new_note);removals.append((old_note,job['note_sha256']))
        record=state['sources'][job['sha256']]
        record['exception_path']=str(new)
        if record.get('path')==str(old):record['path']=str(new)
    save()

    def update(value):
        if isinstance(value,str):return mapping.get(value,value)
        if isinstance(value,list):return [update(item) for item in value]
        if isinstance(value,dict):return {key:update(item) for key,item in value.items()}
        return value
    for name in ['待复核队列.json','最近结果.json']:
        path=paths['state_path'].parent/name
        if not path.exists():continue
        before=json.loads(path.read_text(encoding='utf-8-sig'));after=update(before)
        if before!=after:atomic_write(path,json.dumps(after,ensure_ascii=False,indent=2).encode('utf-8'))
    # Never delete an old file until every required copy and persisted pointer
    # is safe. The persisted plan makes all interruption stages resumable.
    for old,expected in removals:
        if old.exists() and digest(old.read_bytes())!=expected:
            raise RuntimeError('整理期间旧异常文件被修改，保留原件')
    for old,_ in removals:old.unlink(missing_ok=True)
    state.pop('exception_relocations',None);save()
    return len(jobs)
