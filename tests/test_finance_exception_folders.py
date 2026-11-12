import json
import re
from pathlib import Path

import pytest

from test_finance_workspace import module, config, source, Backend


def broken(path, content=b'broken pdf'):
    path.parent.mkdir(parents=True,exist_ok=True);path.write_bytes(content);return path


def test_one_run_groups_exceptions_by_detailed_time_and_repeat_does_not_duplicate(tmp_path):
    mod=module();backend=Backend();cfg=config(tmp_path)
    files=[broken(tmp_path/'待处理/a.pdf'),broken(tmp_path/'待处理/b.pdf',b'another broken pdf')]
    first=mod.run_workspace(cfg,files,backend)
    folder=tmp_path/'异常未处理发票';groups=list(folder.iterdir())
    assert len(groups)==1 and groups[0].is_dir()
    assert re.fullmatch(r'\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}(?:_\d+)?',groups[0].name)
    assert len(list(groups[0].glob('*.pdf')))==2 and len(list(groups[0].glob('*.原因.txt')))==2
    paths=[r['local_path'] for r in first['exceptions']]
    repeat=mod.run_workspace(cfg,files,backend)
    assert [r['local_path'] for r in repeat['exceptions']]==paths
    assert list(folder.iterdir())==groups and backend.calls==0
    newer=mod.run_workspace(cfg,[broken(tmp_path/'待处理/c.pdf',b'third broken pdf')],backend)
    assert newer['pending_review_count']==3 and len(list(folder.iterdir()))==2


def test_resolved_exception_removes_only_its_copy_and_empty_timestamp_folder(tmp_path):
    mod=module();backend=Backend();backend.blocked.add('bad.pdf');cfg=config(tmp_path)
    incoming=source(tmp_path/'待处理/bad.pdf');first=mod.run_workspace(cfg,[incoming],backend)
    item=first['exceptions'][0];folder=Path(item['local_path']).parent
    assert folder.parent==tmp_path/'异常未处理发票'
    mod.resolve_exception(cfg,{'action':'confirm','review_id':item['review_ids'][0],
        'source_sha256':item['source_sha256'],'candidate':{'invoice_type':'vat_special'},
        'reason':'核对票面','pages':[1]},backend)
    result=mod.run_workspace(cfg,[incoming],backend)
    assert result['pending_review_count']==0 and not folder.exists() and incoming.exists()


def seed_legacy(tmp_path):
    mod=module();cfg=config(tmp_path);backend=Backend();incoming=broken(tmp_path/'待处理/old.pdf')
    result=mod.run_workspace(cfg,[incoming],backend);nested=Path(result['exceptions'][0]['local_path'])
    flat=tmp_path/'异常未处理发票'/nested.name
    if nested!=flat:
        nested.rename(flat);nested.with_suffix('.pdf.原因.txt').rename(flat.with_suffix('.pdf.原因.txt'))
        nested.parent.rmdir()
    state_path=Path(cfg['state_path']);state=json.loads(state_path.read_text(encoding='utf-8'))
    next(iter(state['sources'].values()))['exception_path']=str(flat)
    state_path.write_text(json.dumps(state,ensure_ascii=False),encoding='utf-8')
    for name in ['待复核队列.json','最近结果.json']:
        (state_path.parent/name).write_text(json.dumps({'items':[{'local_path':str(flat)}]}),encoding='utf-8')
    return mod,cfg,backend,incoming,flat


def test_flat_legacy_exceptions_are_grouped_with_state_and_queue_paths_updated(tmp_path):
    mod,cfg,backend,incoming,flat=seed_legacy(tmp_path);before=flat.read_bytes()
    result=mod.run_workspace(cfg,[],backend)
    moved=Path(result['exceptions'][0]['local_path'])
    assert moved.parent!=flat.parent and not flat.exists() and moved.read_bytes()==before
    assert not list(flat.parent.glob('*.*'))
    state=json.loads(Path(cfg['state_path']).read_text(encoding='utf-8'))
    assert next(iter(state['sources'].values()))['exception_path']==str(moved)
    assert str(flat) not in (Path(cfg['state_path']).parent/'最近结果.json').read_text(encoding='utf-8')
    assert backend.calls==0 and incoming.exists()


def test_changed_legacy_exception_is_preserved_and_reported(tmp_path):
    mod,cfg,backend,incoming,flat=seed_legacy(tmp_path);flat.write_bytes(b'manually changed')
    result=mod.run_workspace(cfg,[],backend)
    assert result['errors'] and flat.read_bytes()==b'manually changed' and backend.calls==0


@pytest.mark.parametrize('failure_at',[1,2,3])
def test_legacy_relocation_resumes_after_state_save_interruption(tmp_path,failure_at):
    mod,cfg,backend,incoming,flat=seed_legacy(tmp_path)
    import exception_storage
    paths=mod.layout(cfg);state_path=Path(cfg['state_path']);state=json.loads(state_path.read_text(encoding='utf-8'))
    count=0
    def save():
        nonlocal count
        count+=1
        if count==failure_at:raise OSError('simulated save interruption')
        mod.atomic_write(state_path,json.dumps(state,ensure_ascii=False).encode('utf-8'))
    with pytest.raises(OSError):exception_storage.organize_legacy_exceptions(paths,state,save)
    assert incoming.exists() and (flat.exists() or list((tmp_path/'异常未处理发票').rglob('*.pdf')))
    result=mod.run_workspace(cfg,[],backend)
    moved=Path(result['exceptions'][0]['local_path'])
    assert not result['errors'] and moved.parent!=flat.parent and not flat.exists()
    assert len(list((tmp_path/'异常未处理发票').rglob('*.pdf')))==1 and backend.calls==0
    assert all(any(folder.iterdir()) for folder in (tmp_path/'异常未处理发票').iterdir())


def test_same_second_folder_names_are_unique(tmp_path,monkeypatch):
    import exception_storage
    from datetime import datetime,timezone,timedelta
    class FixedClock:
        @staticmethod
        def now(tz):return datetime(2026,10,29,12,34,56,tzinfo=timezone(timedelta(hours=8)))
    monkeypatch.setattr(exception_storage,'datetime',FixedClock)
    root=tmp_path/'异常未处理发票';root.mkdir()
    first=exception_storage.exception_folder({'exception_dir':root})
    second=exception_storage.exception_folder({'exception_dir':root})
    assert first.name=='2026-10-29_12-34-56' and second.name==first.name+'_02'
