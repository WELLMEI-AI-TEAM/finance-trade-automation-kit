import json
import re
from pathlib import Path

import pytest

from test_finance_workspace import module, config, source, Backend, book


def test_backup_group_contains_processed_originals_and_previous_ledger(tmp_path):
    mod=module();cfg=config(tmp_path);backend=Backend()
    ledger=tmp_path/'台账文件/已有账本.xlsx';ledger.parent.mkdir();before=book(['old']);ledger.write_bytes(before)
    incoming=source(tmp_path/'待处理/new.pdf')
    result=mod.run_workspace(cfg,[incoming],backend)
    assert not result['errors']
    groups=list((tmp_path/'备份').iterdir());assert len(groups)==1 and groups[0].is_dir()
    group=groups[0];assert re.fullmatch(r'\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}(?:_\d+)?',group.name)
    assert {p.name for p in group.iterdir()}=={'本次处理的发票','旧台账'}
    assert (group/'旧台账/已有账本.xlsx').read_bytes()==before
    assert [p.read_bytes() for p in (group/'本次处理的发票').iterdir()]==[incoming.read_bytes()]
    calls=(backend.calls,backend.exports,backend.downloads)
    repeat=mod.run_workspace(cfg,[incoming],backend)
    assert not repeat['errors'] and list((tmp_path/'备份').iterdir())==groups
    assert calls==(backend.calls,backend.exports,backend.downloads)


def test_first_ledger_creation_still_backs_up_invoices_with_empty_old_ledger_folder(tmp_path):
    mod=module();incoming=source(tmp_path/'待处理/first.pdf')
    result=mod.run_workspace(config(tmp_path),[incoming],Backend());assert not result['errors']
    group=next((tmp_path/'备份').iterdir())
    assert (group/'旧台账').is_dir() and not list((group/'旧台账').iterdir())
    assert len(list((group/'本次处理的发票').iterdir()))==1


def test_backup_invoice_names_do_not_overwrite_different_sources(tmp_path):
    mod=module();backend=Backend()
    a=source(tmp_path/'待处理/one/invoice.pdf','one');b=source(tmp_path/'待处理/two/invoice.pdf','two')
    result=mod.run_workspace(config(tmp_path),[a,b],backend);assert not result['errors']
    group=next((tmp_path/'备份').iterdir())
    assert {p.read_bytes() for p in (group/'本次处理的发票').iterdir()}=={a.read_bytes(),b.read_bytes()}


def test_backup_keeps_all_current_inputs_including_preflight_exception(tmp_path):
    mod=module();backend=Backend();good=source(tmp_path/'待处理/good.pdf')
    bad=tmp_path/'待处理/bad.pdf';bad.write_bytes(b'broken invoice')
    result=mod.run_workspace(config(tmp_path),[good,bad],backend)
    assert not result['errors'] and result['pending_review_count']==1
    group=next((tmp_path/'备份').iterdir())
    assert {p.read_bytes() for p in (group/'本次处理的发票').iterdir()}=={good.read_bytes(),bad.read_bytes()}


def test_modified_managed_backup_invoice_prevents_rotation(tmp_path):
    mod=module();backend=Backend();cfg=config(tmp_path)
    for n in range(10):assert not mod.run_workspace(cfg,[source(tmp_path/f'待处理/{n}.pdf',str(n))],backend)['errors']
    state=json.loads(Path(cfg['state_path']).read_text(encoding='utf-8'))
    assert len(state['backups'])==10
    oldest=Path(state['backups'][0]['directory']);invoice=next((oldest/'本次处理的发票').iterdir())
    invoice.write_bytes(b'manually changed backup')
    ledger=tmp_path/'台账文件/财务发票台账.xlsx';before=ledger.read_bytes()
    result=mod.run_workspace(cfg,[source(tmp_path/'待处理/next.pdf','next')],backend)
    assert result['errors'] and ledger.read_bytes()==before and invoice.read_bytes()==b'manually changed backup'
    assert len(list((tmp_path/'备份').iterdir()))==10


def test_old_flat_backups_are_grouped_without_guessing_historical_inputs(tmp_path):
    mod=module();cfg=config(tmp_path);paths=mod.layout(cfg)
    for path in [paths['backup_dir'],paths['state_path'].parent]:path.mkdir(parents=True,exist_ok=True)
    old=paths['backup_dir']/'财务台账_20261028_123456_123456_abcd.xlsx';content=book(['old']);old.write_bytes(content)
    state={'backups':[{'path':str(old),'sha256':mod.digest(content)}]}
    def save():mod.atomic_write(paths['state_path'],json.dumps(state,ensure_ascii=False).encode('utf-8'))
    assert hasattr(mod,'organize_legacy_backups')
    mod.organize_legacy_backups(paths,state,save)
    group=Path(state['backups'][0]['directory'])
    assert group.name.startswith('2026-10-28_12-34-56') and not old.exists()
    assert Path(state['backups'][0]['path']).read_bytes()==content
    assert (group/'本次处理的发票').is_dir() and not list((group/'本次处理的发票').iterdir())


def test_batch_backup_restores_original_from_service_when_local_input_is_missing(tmp_path):
    mod=module();cfg=config(tmp_path);backend=Backend();backend.fail_export=True
    incoming=source(tmp_path/'待处理/later.pdf');content=incoming.read_bytes()
    first=mod.run_workspace(cfg,[incoming],backend);assert first['errors']
    incoming.unlink();backend.fail_export=False
    result=mod.run_workspace(cfg,[],backend)
    assert not result['errors'] and backend.calls==1
    group=next((tmp_path/'备份').iterdir())
    assert [p.read_bytes() for p in (group/'本次处理的发票').iterdir()]==[content]


def test_user_ledger_edit_during_backup_creation_is_preserved(tmp_path,monkeypatch):
    mod=module();paths=mod.layout(config(tmp_path));paths['backup_dir'].mkdir(parents=True)
    target=tmp_path/'台账文件/book.xlsx';target.parent.mkdir();target.write_bytes(b'old')
    original=mod.create_backup
    def create(*args):
        entry=original(*args);target.write_bytes(b'user edit');return entry
    monkeypatch.setattr(mod,'create_backup',create)
    with pytest.raises(RuntimeError,match='修改'):
        mod.replace_ledger(target,b'old',b'new',paths,{'backups':[]})
    assert target.read_bytes()==b'user edit' and not list(paths['backup_dir'].iterdir())


def test_generated_collision_name_does_not_overwrite_another_original(tmp_path):
    mod=module();paths=mod.layout(config(tmp_path));paths['backup_dir'].mkdir(parents=True)
    a,b,c=b'A',b'B',b'C'
    sources=[{'name':name,'content':data,'sha256':mod.digest(data)} for name,data in [
        ('invoice.pdf',a),('invoice__'+mod.digest(b)[:16]+'.pdf',c),('invoice.pdf',b)]]
    entry=mod.create_backup(tmp_path/'book.xlsx',None,sources,paths)
    assert {p.read_bytes() for p in (Path(entry['directory'])/'本次处理的发票').iterdir()}=={a,b,c}


def test_backup_names_differing_only_in_case_preserve_both_sources(tmp_path):
    mod=module();paths=mod.layout(config(tmp_path));paths['backup_dir'].mkdir(parents=True)
    sources=[{'name':name,'content':data,'sha256':mod.digest(data)} for name,data in [
        ('Invoice.pdf',b'A'),('invoice.pdf',b'B')]]
    entry=mod.create_backup(tmp_path/'book.xlsx',None,sources,paths)
    assert {p.read_bytes() for p in (Path(entry['directory'])/'本次处理的发票').iterdir()}=={b'A',b'B'}


@pytest.mark.parametrize('failure_at',[1,2])
def test_old_backup_migration_recovers_from_interrupted_state_save(tmp_path,failure_at):
    mod=module();cfg=config(tmp_path);paths=mod.layout(cfg)
    paths['backup_dir'].mkdir(parents=True);paths['state_path'].parent.mkdir(parents=True)
    old=paths['backup_dir']/'old_20261028_123456_000000_abcd.xlsx';old.write_bytes(b'old workbook')
    state={'backups':[{'path':str(old),'sha256':mod.digest(old.read_bytes())}]}
    paths['state_path'].write_text(json.dumps(state),encoding='utf-8');count=0
    def save():
        nonlocal count
        count+=1
        if count==failure_at:raise OSError('interrupted state save')
        paths['state_path'].write_text(json.dumps(state),encoding='utf-8')
    with pytest.raises(OSError):mod.organize_legacy_backups(paths,state,save)
    state=json.loads(paths['state_path'].read_text(encoding='utf-8'))
    def resume_save():paths['state_path'].write_text(json.dumps(state),encoding='utf-8')
    mod.organize_legacy_backups(paths,state,resume_save)
    assert not old.exists() and len(list(paths['backup_dir'].iterdir()))==1
    assert Path(state['backups'][0]['path']).read_bytes()==b'old workbook'
