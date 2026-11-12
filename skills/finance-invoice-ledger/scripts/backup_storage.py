"""Managed time folders containing source originals and the previous workbook."""
from datetime import datetime, timedelta, timezone
from pathlib import Path
import re

from run_invoice_workflow import atomic_write, digest

OLD = '旧台账'
INVOICES = '本次处理的发票'
TZ = timezone(timedelta(hours=8))


def reserve_group(root, stamp=None):
    stamp = stamp or datetime.now(TZ).strftime('%Y-%m-%d_%H-%M-%S')
    index = 1
    while True:
        folder = root/(stamp if index == 1 else f'{stamp}_{index:02d}')
        try:
            folder.mkdir(); break
        except FileExistsError: index += 1
    for name in [INVOICES,OLD]: (folder/name).mkdir()
    return folder


def validate_backup(entry, paths):
    root = paths['backup_dir']
    if not entry.get('directory'):
        old = Path(entry['path']).resolve()
        if old.parent != root: raise RuntimeError('备份索引路径异常，停止替换台账')
        if old.exists() and digest(old.read_bytes()) != entry['sha256']:
            raise RuntimeError('最早备份被人工修改，保留该文件，停止自动轮换')
        return
    folder = Path(entry['directory']).resolve()
    if folder.parent != root: raise RuntimeError('备份文件夹超出指定备份目录，停止轮换')
    expected = {}
    for row in entry['files']:
        relative = Path(row['relative_path']); path = (folder/relative).resolve()
        if relative.is_absolute() or len(relative.parts)!=2 or relative.parts[0] not in {OLD,INVOICES} or path.parent.parent!=folder:
            raise RuntimeError('备份原件路径异常，停止轮换')
        expected[path] = row['sha256']
    if folder.exists():
        for item in folder.rglob('*'):
            path=item.resolve()
            if item.is_dir():
                if path not in {folder/OLD,folder/INVOICES}: raise RuntimeError('备份中新增人工目录，保留该备份，停止轮换')
            elif path not in expected: raise RuntimeError('备份中新增人工文件，保留该备份，停止轮换')
        for path,sha in expected.items():
            if path.exists() and digest(path.read_bytes())!=sha:
                raise RuntimeError('备份原件被人工修改，保留该备份，停止轮换')


def remove_backup(entry, paths):
    validate_backup(entry,paths)
    if not entry.get('directory'):
        Path(entry['path']).unlink(missing_ok=True);return
    folder=Path(entry['directory'])
    for row in entry['files']:(folder/row['relative_path']).unlink(missing_ok=True)
    for child in [folder/INVOICES,folder/OLD]:
        if child.exists():child.rmdir()
    if folder.exists():folder.rmdir()


def create_backup(target, before, sources, paths):
    # Prepare all immutable bytes before creating any visible backup group.
    contents=[];names=set();hashes=set()
    if before is not None:contents.append((Path(OLD)/target.name,before))
    for item in sources:
        if digest(item['content'])!=item['sha256']:raise RuntimeError('本次发票备份哈希不一致，未替换台账')
        if item['sha256'] in hashes:continue
        hashes.add(item['sha256']);original=Path(item['name']).name;name=original;counter=1
        while name.casefold() in names:
            suffix='' if counter==1 else f'_{counter:02d}'
            name=Path(original).stem+'__'+item['sha256'][:16]+suffix+Path(original).suffix
            counter+=1
        names.add(name.casefold());contents.append((Path(INVOICES)/name,item['content']))
    folder=reserve_group(paths['backup_dir'])
    entry={'directory':str(folder),'path':str(folder/OLD/target.name) if before is not None else None,
        'sha256':digest(before) if before is not None else None,
        'files':[{'relative_path':relative.as_posix(),'sha256':digest(content)} for relative,content in contents]}
    try:
        for relative,content in contents:atomic_write(folder/relative,content)
    except Exception:
        remove_backup(entry,paths);raise
    return entry


def preflight_rotation(paths,state):
    retained=state.get('backups',[])
    for entry in retained[:max(0,len(retained)-9)]:validate_backup(entry,paths)


def rotate_backups(paths,state):
    retained=state.setdefault('backups',[])
    while len(retained)>10:
        remove_backup(retained[0],paths);retained.pop(0)


def organize_legacy_backups(paths,state,save):
    moved=0
    for index,entry in enumerate(state.get('backups',[])):
        if not entry.get('directory'):
            validate_backup(entry,paths);old=Path(entry['path'])
            if not old.exists():continue
            content=old.read_bytes()
            if digest(content)!=entry['sha256']:raise RuntimeError('历史备份被人工修改，保留原件')
            match=re.search(r'_(\d{8})_(\d{6})_',old.name)
            stamp=datetime.strptime(''.join(match.groups()),'%Y%m%d%H%M%S').strftime('%Y-%m-%d_%H-%M-%S') if match else None
            folder=reserve_group(paths['backup_dir'],stamp)
            new=folder/OLD/old.name
            replacement={'directory':str(folder),'path':str(new),'sha256':entry['sha256'],
                'files':[{'relative_path':(Path(OLD)/old.name).as_posix(),'sha256':entry['sha256']}],
                'historical_inputs_unrecorded':True,'legacy_path':str(old)}
            try:
                atomic_write(new,content)
                state['backups'][index]=replacement;save()
            except Exception:
                state['backups'][index]=entry;remove_backup(replacement,paths);raise
            entry=replacement
        if entry.get('legacy_path'):
            validate_backup(entry,paths);old=Path(entry['legacy_path']).resolve()
            if old.parent!=paths['backup_dir']:raise RuntimeError('历史备份迁移路径异常，未删除任何文件')
            if old.exists() and digest(old.read_bytes())!=entry['sha256']:
                raise RuntimeError('历史备份被人工修改，保留原件')
            old.unlink(missing_ok=True);entry.pop('legacy_path');save();moved+=1
    return moved
