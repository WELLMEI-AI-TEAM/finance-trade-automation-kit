import json
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0,str(Path(__file__).parent))
from test_finance_workspace import module,config,source,Backend
from test_finance_archive_index import metadata


class SameInvoiceBackend(Backend):
    total = '100.00'

    def list_archives(self, batch_id):
        return [dict(row, invoice_id='same-business-invoice') for row in super().list_archives(batch_id)]

    def archive_manifest(self, batch_id):
        self.manifest_calls += 1
        good = [item for item in self.groups[batch_id] if item['original_filename'] not in self.blocked]
        return {'archive_identities': {'same-business-invoice': metadata(total=self.total)} if good else {},
            'source_documents': {'same-business-invoice': [{'document_id':item['document_id'],
                'source_sha256':module().digest(item['content'])} for item in good]} if good else {}}


class PartialHistoryBackend(Backend):
    def archive_manifest(self,batch_id):
        self.manifest_calls += 1
        docs=self.groups[batch_id];both=next(d for d in docs if d['original_filename']=='both.pdf')
        return {'archive_identities':{'X':metadata('X'),'Y':metadata('Y')},
            'source_documents':{'X':[{'document_id':d['document_id'],'source_sha256':module().digest(d['content'])} for d in docs],
            'Y':[{'document_id':both['document_id'],'source_sha256':module().digest(both['content'])}]}}

    def list_archives(self,batch_id):
        docs=self.groups[batch_id];both=next(d for d in docs if d['original_filename']=='both.pdf')
        one=next(d for d in docs if d['original_filename']=='one.pdf')
        return [{'archive_id':key,'invoice_id':key,'archive_filename':'20261001_seller_'+key+'.pdf',
            'source_sha256':module().digest(doc['content']),'archive_sha256':module().digest(doc['content'])}
            for key,doc in [('X',one),('Y',both)]]


def two_versions(tmp_path):
    original=source(tmp_path/'new.pdf');old=tmp_path/'发票/2026.10/历史人工名称.pdf';old.parent.mkdir(parents=True)
    old.write_bytes(original.read_bytes()+b'\n% other bytes, same invoice\n')
    return original,old


def test_history_different_bytes_is_indexed_once_and_no_duplicate_is_archived(tmp_path):
    mod=module();cfg=config(tmp_path);backend=SameInvoiceBackend();incoming,old=two_versions(tmp_path)
    before=old.read_bytes()
    history=mod.index_history(cfg,backend,limit=20)
    assert history['processed_files']==1 and history['archive_index']['unverified_files']==0
    assert backend.exports==backend.downloads==0
    result=mod.run_workspace(cfg,[incoming],backend)
    assert not result['errors'] and result['pending_review_count']==0
    assert result['archive_writes']==0 and result['archive_identity_reuses']==1
    assert list((tmp_path/'发票').rglob('*.pdf'))==[old] and old.read_bytes()==before
    calls=(backend.calls,backend.exports,backend.downloads,backend.manifest_calls)
    repeat=mod.run_workspace(cfg,[incoming],backend)
    assert calls==(backend.calls,backend.exports,backend.downloads,backend.manifest_calls)
    assert repeat['archive_index']['hashes_read']==0
    assert repeat['processed_files']==repeat['archive_writes']==repeat['ledger_exports']==0


def test_ordinary_processing_blocks_unknown_history_before_new_recognition(tmp_path):
    mod=module();cfg=config(tmp_path);backend=SameInvoiceBackend();incoming,old=two_versions(tmp_path)
    result=mod.run_workspace(cfg,[incoming],backend)
    assert result['pending_review_count']==1 and result['archive_index']['unverified_files']==1
    assert backend.calls==backend.exports==backend.downloads==0
    assert '历史' in result['exceptions'][0]['reason']
    mod.index_history(cfg,backend,limit=20)
    done=mod.run_workspace(cfg,[incoming],backend)
    assert not done['errors'] and done['pending_review_count']==0 and done['archive_writes']==0


def test_normal_processing_seeds_index_and_accepts_confirmed_different_byte_alias(tmp_path):
    mod=module();cfg=config(tmp_path);backend=SameInvoiceBackend()
    incoming=source(tmp_path/'one.pdf');first=mod.run_workspace(cfg,[incoming],backend)
    assert not first['errors'] and first['archive_writes']==1
    variant=tmp_path/'two.pdf';variant.write_bytes(incoming.read_bytes()+b'\n% variant\n')
    result=mod.run_workspace(cfg,[variant],backend)
    assert not result['errors'] and result['pending_review_count']==0 and result['archive_writes']==0
    assert result['archive_identity_reuses']==1
    assert len(list((tmp_path/'发票').rglob('*.pdf')))==1
    book=load_workbook(result['ledger_path']);assert book.active.max_row==2;book.close()


def test_same_identity_conflict_is_quarantined_before_ledger_export(tmp_path):
    mod=module();cfg=config(tmp_path);backend=SameInvoiceBackend()
    original=source(tmp_path/'one.pdf');mod.run_workspace(cfg,[original],backend)
    ledger=tmp_path/'台账文件/财务发票台账.xlsx';before=ledger.read_bytes();exports=backend.exports
    variant=tmp_path/'two.pdf';variant.write_bytes(original.read_bytes()+b'\n% variant\n');backend.total='999'
    result=mod.run_workspace(cfg,[variant],backend)
    assert result['pending_review_count']==1 and result['archive_writes']==0
    assert '冲突' in result['exceptions'][0]['reason'] and backend.exports==exports
    assert ledger.read_bytes()==before and len(list((tmp_path/'发票').rglob('*.pdf')))==1


def test_history_limit_and_repeated_command_do_not_repeat_recognition(tmp_path):
    mod=module();cfg=config(tmp_path);backend=Backend()
    for n in range(3):source(tmp_path/f'发票/2026.10/{n}.pdf',str(n))
    first=mod.index_history(cfg,backend,limit=2)
    assert first['processed_files']==2 and first['archive_index']['unverified_files']==1
    second=mod.index_history(cfg,backend,limit=2)
    assert second['processed_files']==1 and second['archive_index']['unverified_files']==0
    repeat=mod.index_history(cfg,backend,limit=2)
    assert repeat['processed_files']==0 and backend.calls==2
    assert backend.exports==backend.downloads==0 and not list((tmp_path/'台账文件').glob('*.xlsx'))


def test_old_managed_state_rebuilds_deleted_index_without_recognition(tmp_path):
    mod=module();cfg=config(tmp_path);backend=Backend();incoming=source(tmp_path/'one.pdf')
    mod.run_workspace(cfg,[incoming],backend)
    index=tmp_path/'.运行记录/归档身份索引.json';assert index.exists();index.unlink()
    state_file=Path(cfg['state_path']);state=json.loads(state_file.read_text(encoding='utf-8'))
    for batch in state['batches'].values():batch.pop('identity_manifest',None);batch.pop('identity_revision',None)
    state_file.write_text(json.dumps(state),encoding='utf-8')
    count=(backend.calls,backend.exports,backend.downloads)
    result=mod.index_history(cfg,backend,limit=20)
    assert result['processed_files']==0 and result['archive_index']['verified_files']==1
    assert count==(backend.calls,backend.exports,backend.downloads)


def test_identical_archive_input_can_build_index_from_its_one_normal_processing_result(tmp_path):
    mod=module();cfg=config(tmp_path);backend=Backend()
    archived=source(tmp_path/'发票/2026.09/renamed.pdf');incoming=tmp_path/'input.pdf';incoming.write_bytes(archived.read_bytes())
    result=mod.run_workspace(cfg,[incoming],backend)
    assert not result['errors'] and result['archive_writes']==0
    assert result['archive_index']['unverified_files']==0 and backend.calls==1


def test_known_history_bytes_reuse_receipt_for_ledger_without_uploading_again(tmp_path):
    mod=module();cfg=config(tmp_path);backend=Backend()
    archived=source(tmp_path/'发票/2026.10/old.pdf');incoming=tmp_path/'input.pdf';incoming.write_bytes(archived.read_bytes())
    mod.index_history(cfg,backend,limit=20)
    calls=backend.calls
    result=mod.run_workspace(cfg,[incoming],backend)
    assert not result['errors'] and result['processed_files']==0 and backend.calls==calls
    assert result['ledger_exports']==1 and result['archive_writes']==0


def test_missing_identity_metadata_contract_blocks_publication(tmp_path):
    mod=module();cfg=config(tmp_path);backend=Backend()
    backend.archive_manifest=lambda batch_id:{'source_documents':{}}
    result=mod.run_workspace(cfg,[source(tmp_path/'input.pdf')],backend)
    assert result['errors'] and backend.exports==backend.downloads==0


def test_unsupported_or_damaged_history_is_reported_without_model_call(tmp_path):
    mod=module();cfg=config(tmp_path);backend=Backend();folder=tmp_path/'发票';folder.mkdir()
    (folder/'broken.pdf').write_bytes(b'broken');(folder/'unsupported.ofd').write_bytes(b'unsupported')
    result=mod.index_history(cfg,backend,limit=20)
    assert result['archive_index']['unverified_files']==2 and backend.calls==0
    assert len(result['unresolved_history'])==2


def test_model_output_is_short_but_full_exception_evidence_stays_in_result(tmp_path):
    mod=module();data={'processed_files':1,'pending_review_count':100,'errors':[],
        'exceptions':[{'source':f'{n}.pdf','reason':'reason-'+str(n)} for n in range(100)],
        'batch_ids':[str(n) for n in range(30)]}
    summary=mod.output_summary(data,tmp_path/'result.json')
    assert len(json.dumps(summary))<2500
    assert summary['pending_review_count']==100 and len(summary['exception_samples'])<=3
    assert len(data['exceptions'])==100 and summary['result_file']==str(tmp_path/'result.json')


def test_partially_overlapping_historical_pdf_sets_are_not_reported_fully_indexed(tmp_path):
    mod=module();backend=PartialHistoryBackend()
    source(tmp_path/'发票/both.pdf','X,Y');source(tmp_path/'发票/one.pdf','X')
    result=mod.index_history(config(tmp_path),backend)
    assert result['status']=='partial' and result['archive_index']['unverified_files']>=1
    assert result['unresolved_history'] and len(list((tmp_path/'发票').glob('*.pdf')))==2


def test_managed_rebuild_uses_manifest_complete_set_instead_of_canonical_archive_ids(tmp_path):
    mod=module();backend=PartialHistoryBackend()
    both=source(tmp_path/'发票/both.pdf','X,Y');one=source(tmp_path/'发票/one.pdf','X')
    result=backend.process_many([both,one],'old');sha_both=mod.digest(both.read_bytes());sha_one=mod.digest(one.read_bytes())
    state={'archives':{'X':{'path':str(one),'sha256':sha_one},'Y':{'path':str(both),'sha256':sha_both}},
        'batches':{result['batch_id']:{'archive_items':backend.list_archives(result['batch_id']),'result':result}}}
    index=mod.ArchiveIndex(tmp_path/'发票',tmp_path/'index.json','scope');inventory=index.scan()
    mod.sync_managed_index(state,index,inventory,backend)
    row=index.files['both.pdf']
    assert row['status']!='verified' or len(row['identities'])==2
    with pytest.raises(mod.ArchiveConflict):index.choose([metadata('Y')],mod.digest(b'new Y file'))


def test_history_timeout_does_not_wait_twice_in_one_invocation(tmp_path):
    class TimeoutBackend(Backend):
        attempts = 0
        def process_many(self, paths, key, on_batch=None, batch_id=None, **kwargs):
            self.attempts += 1
            if on_batch: on_batch('pending-history')
            raise TimeoutError('waiting for history')
    mod=module();backend=TimeoutBackend();source(tmp_path/'发票/old.pdf')
    result=mod.index_history(config(tmp_path),backend)
    assert backend.attempts==1 and result['errors']
    mod.index_history(config(tmp_path),backend)
    assert backend.attempts==2


def test_same_bytes_history_copies_rebuild_without_new_recognition(tmp_path):
    mod=module();backend=Backend();cfg=config(tmp_path);incoming=source(tmp_path/'input.pdf')
    mod.run_workspace(cfg,[incoming],backend)
    original=next((tmp_path/'发票').rglob('*.pdf'));copy=tmp_path/'发票/copy.pdf'
    copy.write_bytes(original.read_bytes());(tmp_path/'.运行记录/归档身份索引.json').unlink()
    calls=backend.calls
    result=mod.run_workspace(cfg,[incoming],backend)
    assert result['archive_index']['verified_files']==2
    history=mod.index_history(cfg,backend)
    assert history['status']=='complete' and backend.calls==calls


@pytest.mark.skipif(sys.platform!='win32',reason='Windows change-time regression')
def test_fast_repeat_detects_same_length_archive_edit_with_restored_mtime(tmp_path):
    import os
    mod=module();backend=Backend();cfg=config(tmp_path);incoming=source(tmp_path/'input.pdf')
    mod.run_workspace(cfg,[incoming],backend)
    original=next((tmp_path/'发票').rglob('*.pdf'));stat=original.stat();content=original.read_bytes()
    original.write_bytes(content.replace(b'PDF',b'PXX',1));os.utime(original,ns=(stat.st_atime_ns,stat.st_mtime_ns))
    result=mod.run_workspace(cfg,[incoming],backend)
    assert result['archive_index']['unverified_files']==1
    assert result['errors'] or result['pending_review_count']


def test_archive_created_during_download_is_never_overwritten(tmp_path):
    mod=module();cfg=config(tmp_path);paths=mod.layout(cfg)
    target=tmp_path/'发票/2026/10/20261001_seller.pdf'
    class ConcurrentWriter(Backend):
        def download_archive(self, archive_id):
            target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(b'other original')
            return b'incoming original'
    index=mod.ArchiveIndex(paths['archive_dir'],tmp_path/'index.json','scope')
    group={'sha256':mod.digest(b'incoming original'),'choice':{'action':'new'},'target':target,
        'items':[{'archive_id':'a','invoice_id':'i','archive_filename':target.name}],
        'metadata':[metadata()]}
    with pytest.raises(mod.ArchiveConflict):
        mod.publish_indexed_group(group,paths,{'archives':{}},index,{},ConcurrentWriter(),'b','year/month',[])
    assert target.read_bytes()==b'other original'


def test_shared_archive_tree_is_locked_across_separate_state_paths(tmp_path):
    mod=module();cfg=config(tmp_path);archive=tmp_path/'发票';archive.mkdir()
    with mod.exclusive(archive/'.finance-archive.lock'):
        with pytest.raises((RuntimeError,OSError)):
            mod.run_workspace(cfg,[source(tmp_path/'input.pdf')],Backend())
