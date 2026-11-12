import json
import sys
from pathlib import Path

import pytest

from test_finance_workspace import module, config, source, Backend


def invoke(monkeypatch,root,backend,*inputs):
    mod=module();cfg=config(root);runtime=Path(cfg['state_path']).parent;runtime.mkdir(parents=True,exist_ok=True)
    cfg_path=runtime/'config.json';cfg_path.write_text(json.dumps(cfg),encoding='utf-8')
    monkeypatch.setattr(mod,'WorkspaceBackend',lambda cfg:backend)
    monkeypatch.setattr(backend,'close',lambda:None,raising=False)
    result_path=runtime/'result.json'
    args=['--config',str(cfg_path),'--result',str(result_path)]
    if inputs:args+=['--input',*map(str,inputs)]
    exit_code=mod.main(args)
    return exit_code,json.loads(result_path.read_text(encoding='utf-8'))


def test_entry_cleans_normal_inputs_only_after_ledger_archive_backup_and_result_are_saved(tmp_path,monkeypatch):
    incoming=source(tmp_path/'待处理/nested/good.pdf');content=incoming.read_bytes()
    code,result=invoke(monkeypatch,tmp_path,Backend())
    assert code==0 and not result['errors']
    assert not incoming.exists() and (tmp_path/'待处理').is_dir()
    assert result['cleaned_input_files']==1
    assert any(p.read_bytes()==content for p in (tmp_path/'发票').rglob('*.pdf'))
    assert any(p.read_bytes()==content for p in (tmp_path/'备份').rglob('*.pdf'))
    assert (tmp_path/'台账文件/财务发票台账.xlsx').is_file()


def test_entry_cleans_same_byte_duplicates_and_external_attachments_are_retained(tmp_path,monkeypatch):
    backend=Backend();a=source(tmp_path/'待处理/a.pdf');content=a.read_bytes();b=a.with_name('copy.pdf');b.write_bytes(content)
    external=source(tmp_path/'attachment.pdf','attachment')
    code,result=invoke(monkeypatch,tmp_path,backend,a,b,external)
    assert code==0 and result['cleaned_input_files']==2
    assert not a.exists() and not b.exists() and external.is_file()
    assert len(list((tmp_path/'发票').rglob('*.pdf')))==2 and backend.calls==1


def test_entry_moves_exception_out_of_inbox_without_changing_reviewability(tmp_path,monkeypatch):
    backend=Backend();backend.blocked.add('bad.pdf');incoming=source(tmp_path/'待处理/bad.pdf');content=incoming.read_bytes()
    code,result=invoke(monkeypatch,tmp_path,backend)
    assert code==2 and not result['errors'] and not incoming.exists()
    assert result['pending_review_count']==1 and result['cleaned_input_files']==1
    exception=Path(result['exceptions'][0]['local_path'])
    assert exception.read_bytes()==content and exception.with_suffix('.pdf.原因.txt').is_file()
    assert not list((tmp_path/'发票').rglob('*.pdf'))
    calls=backend.calls
    code,repeat=invoke(monkeypatch,tmp_path,backend)
    assert code==2 and not repeat['errors'] and repeat['cleaned_input_files']==0
    assert repeat['exceptions'][0]['local_path']==str(exception) and backend.calls==calls


def test_entry_preserves_inputs_after_export_failure(tmp_path,monkeypatch):
    backend=Backend();backend.fail_export=True;incoming=source(tmp_path/'待处理/good.pdf');content=incoming.read_bytes()
    code,result=invoke(monkeypatch,tmp_path,backend)
    assert code==2 and result['errors'] and incoming.read_bytes()==content
    assert result['cleaned_input_files']==0


def test_entry_preserves_all_inputs_if_initial_result_cannot_be_saved(tmp_path,monkeypatch):
    mod=module();backend=Backend();incoming=source(tmp_path/'待处理/good.pdf');content=incoming.read_bytes();cfg=config(tmp_path)
    runtime=Path(cfg['state_path']).parent;runtime.mkdir();cfg_path=runtime/'config.json';cfg_path.write_text(json.dumps(cfg),encoding='utf-8')
    result_path=runtime/'result.json';original=mod.atomic_write
    def write(path,data):
        if Path(path)==result_path:raise OSError('result save interrupted')
        return original(path,data)
    monkeypatch.setattr(mod,'atomic_write',write);monkeypatch.setattr(mod,'WorkspaceBackend',lambda cfg:backend)
    monkeypatch.setattr(backend,'close',lambda:None,raising=False)
    with pytest.raises(OSError):mod.main(['--config',str(cfg_path),'--result',str(result_path)])
    assert incoming.read_bytes()==content


def test_input_changed_after_processing_is_retained(tmp_path,monkeypatch):
    backend=Backend();incoming=source(tmp_path/'待处理/good.pdf');original_close=lambda:incoming.write_bytes(b'user changed original')
    mod=module();cfg=config(tmp_path);runtime=Path(cfg['state_path']).parent;runtime.mkdir()
    cfg_path=runtime/'config.json';cfg_path.write_text(json.dumps(cfg),encoding='utf-8');result_path=runtime/'result.json'
    monkeypatch.setattr(mod,'WorkspaceBackend',lambda cfg:backend);monkeypatch.setattr(backend,'close',original_close,raising=False)
    assert mod.main(['--config',str(cfg_path),'--result',str(result_path)])==2
    result=json.loads(result_path.read_text(encoding='utf-8'))
    assert incoming.read_bytes()==b'user changed original' and result['cleaned_input_files']==0 and result['errors']


def test_missing_archive_after_processing_prevents_input_deletion(tmp_path,monkeypatch):
    backend=Backend();incoming=source(tmp_path/'待处理/good.pdf');content=incoming.read_bytes();mod=module();cfg=config(tmp_path)
    runtime=Path(cfg['state_path']).parent;runtime.mkdir();cfg_path=runtime/'config.json';cfg_path.write_text(json.dumps(cfg),encoding='utf-8');result_path=runtime/'result.json'
    def close():next((tmp_path/'发票').rglob('*.pdf')).unlink()
    monkeypatch.setattr(mod,'WorkspaceBackend',lambda cfg:backend);monkeypatch.setattr(backend,'close',close,raising=False)
    assert mod.main(['--config',str(cfg_path),'--result',str(result_path)])==2
    result=json.loads(result_path.read_text(encoding='utf-8'))
    assert incoming.read_bytes()==content and result['cleaned_input_files']==0 and result['errors']


@pytest.mark.skipif(sys.platform!='win32',reason='Windows junction')
def test_linked_inbox_directory_cannot_delete_external_invoice(tmp_path,monkeypatch):
    import _winapi
    outside=tmp_path/'external';outside.mkdir();incoming=source(outside/'invoice.pdf')
    inbox=tmp_path/'待处理';inbox.mkdir();_winapi.CreateJunction(str(outside),str(inbox/'linked'))
    code,result=invoke(monkeypatch,tmp_path,Backend())
    assert code==0 and incoming.is_file() and result['cleaned_input_files']==0


def test_known_different_byte_duplicate_is_cleaned_only_after_identity_reuse(tmp_path,monkeypatch):
    from test_finance_history_index import SameInvoiceBackend
    backend=SameInvoiceBackend();first=source(tmp_path/'待处理/first.pdf');content=first.read_bytes()
    assert invoke(monkeypatch,tmp_path,backend)[0]==0
    second=tmp_path/'待处理/second.pdf';second.write_bytes(content+b'\n% different encoding\n')
    code,result=invoke(monkeypatch,tmp_path,backend)
    assert code==0 and result['archive_identity_reuses']==1 and result['cleaned_input_files']==1
    assert not second.exists() and len(list((tmp_path/'发票').rglob('*.pdf')))==1


def test_failed_unlink_can_resume_cleanup_without_recognition_or_export(tmp_path,monkeypatch):
    backend=Backend();incoming=source(tmp_path/'待处理/good.pdf');original=Path.unlink
    def unlink(path,*args,**kwargs):
        if path==incoming:raise PermissionError('inbox file in use')
        return original(path,*args,**kwargs)
    with monkeypatch.context() as patch:
        patch.setattr(Path,'unlink',unlink)
        code,result=invoke(patch,tmp_path,backend)
        assert code==2 and result['errors'] and incoming.exists() and result['cleaned_input_files']==0
    calls=(backend.calls,backend.exports,backend.downloads)
    code,result=invoke(monkeypatch,tmp_path,backend)
    assert code==0 and not incoming.exists() and result['cleaned_input_files']==1
    assert calls==(backend.calls,backend.exports,backend.downloads)


def test_new_arrival_during_processing_is_not_deleted(tmp_path,monkeypatch):
    mod=module();backend=Backend();incoming=source(tmp_path/'待处理/good.pdf');new=tmp_path/'待处理/new.pdf';cfg=config(tmp_path)
    runtime=Path(cfg['state_path']).parent;runtime.mkdir();cfg_path=runtime/'config.json';cfg_path.write_text(json.dumps(cfg),encoding='utf-8')
    monkeypatch.setattr(mod,'WorkspaceBackend',lambda cfg:backend);monkeypatch.setattr(backend,'close',lambda:source(new,'new arrival'),raising=False)
    assert mod.main(['--config',str(cfg_path),'--result',str(runtime/'result.json')])==0
    assert not incoming.exists() and new.is_file()


def test_missing_exception_reason_prevents_cleanup(tmp_path,monkeypatch):
    mod=module();backend=Backend();backend.blocked.add('bad.pdf');incoming=source(tmp_path/'待处理/bad.pdf');cfg=config(tmp_path)
    runtime=Path(cfg['state_path']).parent;runtime.mkdir();cfg_path=runtime/'config.json';cfg_path.write_text(json.dumps(cfg),encoding='utf-8')
    monkeypatch.setattr(mod,'WorkspaceBackend',lambda cfg:backend)
    monkeypatch.setattr(backend,'close',lambda:next((tmp_path/'异常未处理发票').rglob('*.原因.txt')).unlink(),raising=False)
    assert mod.main(['--config',str(cfg_path),'--result',str(runtime/'result.json')])==2
    assert incoming.is_file() and list((tmp_path/'异常未处理发票').rglob('*.pdf'))
