import importlib.util
import io
import json
import sys
from pathlib import Path

import fitz
import pytest
from openpyxl import Workbook, load_workbook

SCRIPTS = Path(__file__).parents[1] / 'skills/finance-invoice-ledger/scripts'
sys.path.insert(0, str(SCRIPTS))


def module():
    path = SCRIPTS / 'run_workspace.py'
    assert path.exists(), 'desktop workspace workflow is missing'
    spec = importlib.util.spec_from_file_location('workspace_under_test', path)
    mod = importlib.util.module_from_spec(spec); spec.loader.exec_module(mod)
    return mod


def book(ids):
    workbook = Workbook(); sheet = workbook.active; sheet.title = '普通发票'
    sheet.append(['日期', '供应商', '发票号码', '备注'])
    for number in ids: sheet.append(['2026-10-28', '测试供应商', number, '人工备注'])
    target = io.BytesIO(); workbook.save(target); return target.getvalue()


def source(path, text='invoice'):
    doc = fitz.open(); page = doc.new_page(); page.insert_text((30, 30), text)
    path.parent.mkdir(parents=True, exist_ok=True); path.write_bytes(doc.tobytes()); doc.close()
    return path


class Backend:
    def __init__(self):
        self.groups = {}; self.files = {}; self.calls = self.exports = self.downloads = self.confirms = 0
        self.manifest_calls = 0
        self.blocked = set(); self.fail_export = False; self.change_ledger = None

    def identity(self): return {'username': 'finance', 'user_id': 'owner', 'department': 'finance'}

    def process_many(self, paths, key, batch_id=None, on_batch=None, expected_hashes=None):
        if batch_id is None:
            self.calls += 1; batch_id = 'batch-' + str(self.calls)
            self.groups[batch_id] = []
            for path in paths:
                number = str(len(self.files) + 1); content = path.read_bytes()
                item = {'document_id': 'doc-'+number, 'original_filename': path.name,
                        'byte_length': len(content), 'invoice_id': 'invoice-'+number,
                        'content': content, 'date': '20260911' if 'old' in path.name else '20261001'}
                self.files[item['document_id']] = item; self.groups[batch_id].append(item)
        if on_batch: on_batch(batch_id)
        return self.get_batch(batch_id)

    def get_batch(self, batch_id):
        items = self.groups[batch_id]; failed = [v for v in items if v['original_filename'] in self.blocked]
        return {'batch_id': batch_id, 'status': 'completed_with_errors' if failed else 'completed',
            'accepted_invoice_count': len(items)-len(failed), 'skipped_invoice_count': 0,
            'documents': [{k:v for k,v in item.items() if k in {'document_id','original_filename','byte_length'}} for item in items],
            'manual_reviews': [{'review_id': 'review-'+item['document_id'], 'status': 'pending',
                'document_id': item['document_id'], 'issues': [{'code': 'invoice_type_evidence_missing', 'message': '票种标题不清晰'}]} for item in failed],
            'rejections': []}

    def list_archives(self, batch_id):
        mod = module()
        return [{'archive_id': item['invoice_id'], 'invoice_id': item['invoice_id'],
            'archive_filename': item['date']+'_seller_'+item['invoice_id']+'.pdf',
            'source_sha256': mod.digest(item['content']), 'archive_sha256': mod.digest(item['content'])}
            for item in self.groups[batch_id] if item['original_filename'] not in self.blocked]

    def download_archive(self, archive_id):
        self.downloads += 1
        return next(item['content'] for item in self.files.values() if item['invoice_id'] == archive_id)

    def original_bytes(self, batch_id, doc_id): return self.files[doc_id]['content']
    def invoice_sources(self, invoice_id, batch_id): return []

    def archive_manifest(self, batch_id):
        from test_finance_archive_index import metadata
        self.manifest_calls += 1
        sources = {}; values = {}
        for item in self.list_archives(batch_id):
            row = next((r for r in self.groups[batch_id]
                        if module().digest(r['content']) == item['archive_sha256']), self.groups[batch_id][0])
            values[item['invoice_id']] = metadata(number=item['invoice_id'])
            sources.setdefault(item['invoice_id'], []).append({'document_id': row['document_id'],
                'source_sha256': module().digest(row['content'])})
        return {'source_documents': sources, 'archive_identities': values}

    def export_selected(self, batch_id, ids, existing):
        self.exports += 1
        if self.change_ledger: self.change_ledger()
        if self.fail_export: raise RuntimeError('existing_ledger_incompatible')
        old = []
        if existing:
            workbook = load_workbook(io.BytesIO(existing)); old = [row[2] for row in list(workbook.active.values)[1:]]; workbook.close()
        return book(old + [value for value in ids if value not in old])

    def get_review(self, review_id):
        return {'review_id': review_id, 'status': 'pending', 'original_candidate': {'invoice_type': 'vat_ordinary'}}

    def confirm_review(self, review_id, candidate, reason):
        self.confirms += 1
        item = self.files[review_id.removeprefix('review-')]
        self.blocked.remove(item['original_filename'])
        return {'status': 'accepted'}


def config(root):
    return {'workspace_root': str(root), 'base_url': 'http://test', 'username': 'finance',
        'input_dir': str(root/'待处理'), 'archive_dir': str(root/'发票'),
        'ledger_dir': str(root/'台账文件'), 'exception_dir': str(root/'异常未处理发票'),
        'backup_dir': str(root/'备份'), 'state_path': str(root/'.运行记录/state.json')}


def test_year_month_archive_and_idempotent_repetition(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend()
    first = source(tmp_path/'待处理/old.pdf'); second = source(tmp_path/'待处理/new.pdf', 'other')
    result = mod.run_workspace(cfg, [first, second], backend)
    assert not result['errors'] and result['archive_writes'] == 2
    assert len(list((tmp_path/'发票/2026/09').glob('*.pdf'))) == 1
    assert len(list((tmp_path/'发票/2026/10').glob('*.pdf'))) == 1
    assert backend.calls == backend.exports == 1
    before = (backend.calls, backend.exports, backend.downloads)
    assert not mod.run_workspace(cfg, [first, second], backend)['errors']
    assert (backend.calls, backend.exports, backend.downloads) == before
    assert not list((tmp_path/'备份').rglob('*.xlsx'))


def test_new_workspace_uses_dotted_months_and_routes_by_each_invoice_date(tmp_path):
    mod = module(); cfg = config(tmp_path); cfg['archive_month_layout'] = 'year.month'
    backend = Backend()
    result = mod.run_workspace(cfg, [source(tmp_path/'old.pdf'), source(tmp_path/'new.pdf', 'other')], backend)
    assert not result['errors']
    assert len(list((tmp_path/'发票/2026.09').glob('*.pdf'))) == 1
    assert len(list((tmp_path/'发票/2026.10').glob('*.pdf'))) == 1
    assert not (tmp_path/'发票/2026').exists()


def test_same_bytes_anywhere_in_archive_are_reused_and_wrong_month_is_reported(tmp_path):
    mod = module(); cfg = config(tmp_path); cfg['archive_month_layout'] = 'year.month'
    backend = Backend(); incoming = source(tmp_path/'new.pdf')
    archived = tmp_path/'发票/2026.08/历史人工名称.pdf'
    archived.parent.mkdir(parents=True); archived.write_bytes(incoming.read_bytes())
    result = mod.run_workspace(cfg, [incoming], backend)
    assert not result['errors'] and result['archive_writes'] == 0
    assert archived.exists() and list((tmp_path/'发票').rglob('*.pdf')) == [archived]
    assert result['archive_location_warnings']
    assert backend.downloads == 0


def test_same_content_renamed_and_missing_archive_do_not_repeat_recognition(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend()
    first = source(tmp_path/'待处理/first.pdf'); duplicate = first.with_name('renamed.pdf'); duplicate.write_bytes(first.read_bytes())
    assert mod.run_workspace(cfg, [first, duplicate], backend)['duplicate_files'] == 1
    next((tmp_path/'发票').rglob('*.pdf')).unlink()
    result = mod.run_workspace(cfg, [duplicate], backend)
    assert result['archive_writes'] == 1
    assert backend.calls == backend.exports == 1


def test_mixed_batch_publishes_normal_file_and_only_quarantines_exception(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.blocked.add('bad.pdf')
    good = source(tmp_path/'待处理/good.pdf'); bad = source(tmp_path/'待处理/bad.pdf', 'other')
    result = mod.run_workspace(cfg, [good, bad], backend)
    assert not result['errors'] and result['pending_review_count'] == 1
    assert result['archive_writes'] == 1 and backend.exports == 1
    assert len(list((tmp_path/'异常未处理发票').rglob('*.pdf'))) == 1
    assert all(p.read_bytes() != bad.read_bytes() for p in (tmp_path/'发票').rglob('*.pdf'))
    assert bad.exists()
    repeat = mod.run_workspace(cfg, [good, bad], backend)
    assert repeat['pending_review_count'] == 1 and backend.calls == backend.exports == 1


def test_review_correction_resumes_same_batch_and_cleans_only_owned_exception(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.blocked.add('bad.pdf')
    bad = source(tmp_path/'待处理/bad.pdf')
    result = mod.run_workspace(cfg, [bad], backend)
    item = result['exceptions'][0]
    decision = {'action': 'confirm', 'review_id': item['review_ids'][0], 'source_sha256': mod.digest(bad.read_bytes()),
                'candidate': {'invoice_type': 'vat_special'}, 'reason': '已核对票面完整标题为增值税专用发票', 'pages': [1]}
    mod.resolve_exception(cfg, decision, backend)
    result = mod.run_workspace(cfg, [bad], backend)
    assert result['pending_review_count'] == 0 and result['archive_writes'] == 1
    assert backend.calls == backend.confirms == 1
    assert not list((tmp_path/'异常未处理发票').rglob('*.pdf'))


def test_backup_retention_ten_and_original_bytes_are_preserved(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); originals = []
    for n in range(13):
        item = source(tmp_path/'待处理'/f'{n}.pdf', str(n))
        ledger = tmp_path/'台账文件/财务发票台账.xlsx'
        if ledger.exists(): originals.append(ledger.read_bytes())
        assert not mod.run_workspace(cfg, [item], backend)['errors']
    backups = sorted((tmp_path/'备份').rglob('*.xlsx'))
    assert len(backups) == 10
    assert [p.read_bytes() for p in backups] == originals[-10:]
    assert not list((tmp_path/'台账文件').glob('*backup*'))


def test_existing_single_ledger_name_is_reused_and_backed_up(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend()
    directory = tmp_path/'台账文件'; directory.mkdir(); target = directory/'已有财务台账.xlsx'; previous = book(['old'])
    target.write_bytes(previous)
    result = mod.run_workspace(cfg, [source(tmp_path/'input.pdf')], backend)
    assert not result['errors'] and result['ledger_path'] == str(target)
    assert next((tmp_path/'备份').rglob('*.xlsx')).read_bytes() == previous
    workbook = load_workbook(target); assert workbook.active['D2'].value == '人工备注'; workbook.close()


def test_multiple_ledgers_fail_without_processing_or_overwriting(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend()
    directory = tmp_path/'台账文件'; directory.mkdir()
    for name in ['first.xlsx', 'second.xlsx']: (directory/name).write_bytes(book([name]))
    with pytest.raises(ValueError, match='多个台账'):
        mod.run_workspace(cfg, [source(tmp_path/'input.pdf')], backend)
    assert backend.calls == 0


def test_export_failure_preserves_old_ledger_and_does_not_publish_archive(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.fail_export = True
    directory = tmp_path/'台账文件'; directory.mkdir(); ledger = directory/'old.xlsx'; original = book(['old']); ledger.write_bytes(original)
    result = mod.run_workspace(cfg, [source(tmp_path/'input.pdf')], backend)
    assert result['errors'] and ledger.read_bytes() == original
    assert not list((tmp_path/'发票').rglob('*.pdf')) and not list((tmp_path/'备份').rglob('*.xlsx'))


def test_bad_or_unsupported_document_is_quarantined_without_provider_call(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend()
    bad = tmp_path/'broken.pdf'; bad.write_bytes(b'not a pdf')
    other = tmp_path/'invoice.zip'; other.write_bytes(b'unsupported')
    result = mod.run_workspace(cfg, [bad, other], backend)
    assert result['pending_review_count'] == 2 and backend.calls == 0
    assert len(result['exceptions']) == 2 and all(item['reason'] for item in result['exceptions'])
    assert not list((tmp_path/'发票').rglob('*.pdf'))


def test_wrong_department_refused_before_output_creation(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend()
    backend.identity = lambda: {'username': 'cashier', 'department': 'cashier'}
    with pytest.raises(ValueError, match='finance'):
        mod.run_workspace(cfg, [], backend)
    assert not (tmp_path/'发票').exists()


def test_secondary_review_rejects_hash_mismatch_and_does_not_call_confirmation(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.blocked.add('bad.pdf')
    bad = source(tmp_path/'bad.pdf'); item = mod.run_workspace(cfg, [bad], backend)['exceptions'][0]
    with pytest.raises(ValueError, match='hash'):
        mod.resolve_exception(cfg, {'action': 'confirm', 'review_id': item['review_ids'][0],
            'source_sha256': 'wrong', 'reason': 'verified', 'pages': [1], 'candidate': {}}, backend)
    assert backend.confirms == 0


def test_deferred_reason_survives_refresh_and_is_not_reviewed_twice(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.blocked.add('bad.pdf')
    bad = source(tmp_path/'bad.pdf'); item = mod.run_workspace(cfg, [bad], backend)['exceptions'][0]
    mod.resolve_exception(cfg, {'action': 'defer', 'source_sha256': item['source_sha256'],
        'reason': '票面模糊，无法核实专票或普票，请提供清晰原件'}, backend)
    repeat = mod.run_workspace(cfg, [], backend)
    assert repeat['exceptions'][0]['reason'].startswith('票面模糊')
    assert not repeat['exceptions'][0]['can_secondary_review'] and backend.confirms == 0


def test_partial_pdf_never_exports_or_archives_until_all_reviews_are_resolved(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.blocked.add('mixed.pdf')
    original_batch = backend.get_batch; original_archives = backend.list_archives
    def batch(batch_id):
        result = original_batch(batch_id); result['accepted_invoice_count'] = 1; return result
    def archives(batch_id):
        backend.blocked.clear()
        items = original_archives(batch_id); backend.blocked.add('mixed.pdf'); return items
    backend.get_batch = batch; backend.list_archives = archives
    result = mod.run_workspace(cfg, [source(tmp_path/'mixed.pdf')], backend)
    assert not result['errors'] and result['pending_review_count'] == 1
    assert backend.exports == backend.downloads == 0
    assert not list((tmp_path/'发票').rglob('*.pdf'))


def test_partial_batch_correction_adds_only_missing_invoice(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.blocked.add('bad.pdf')
    good = source(tmp_path/'good.pdf'); bad = source(tmp_path/'bad.pdf', 'other')
    result = mod.run_workspace(cfg, [good, bad], backend); item = result['exceptions'][0]
    mod.resolve_exception(cfg, {'action':'confirm', 'source_sha256':item['source_sha256'],
        'review_id':item['review_ids'][0], 'reason':'核对完整专票标题', 'pages':[1], 'candidate':{}}, backend)
    result = mod.run_workspace(cfg, [], backend)
    assert not result['errors'] and result['archive_writes'] == 1
    assert backend.calls == 1 and backend.exports == 2
    workbook = load_workbook(result['ledger_path']); assert workbook.active.max_row == 3; workbook.close()
    assert len(list((tmp_path/'备份').rglob('*.xlsx'))) == 1


def test_timeout_does_not_poll_twice_in_same_run_and_resumes_next_run(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); process = backend.process_many; attempts = []
    def timeout(*args, **kwargs):
        result = process(*args, **kwargs); attempts.append(result['batch_id']); raise TimeoutError('still running')
    backend.process_many = timeout
    result = mod.run_workspace(cfg, [source(tmp_path/'file.pdf')], backend)
    assert len(attempts) == 1 and result['errors']
    backend.process_many = process
    assert not mod.run_workspace(cfg, [], backend)['errors'] and backend.calls == 1


def test_failure_before_batch_id_can_retry_without_new_input(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); process = backend.process_many
    backend.process_many = lambda *a, **k: (_ for _ in ()).throw(RuntimeError('service unavailable'))
    assert mod.run_workspace(cfg, [source(tmp_path/'file.pdf')], backend)['errors']
    backend.process_many = process
    result = mod.run_workspace(cfg, [], backend)
    assert result['archive_writes'] == 1 and backend.calls == 1


def test_same_archive_bytes_with_two_identities_have_one_stable_path(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); original = backend.list_archives
    def archives(batch_id):
        items = original(batch_id); alias = dict(items[0], invoice_id='alias', archive_filename='20261001_other_alias.pdf')
        return items + [alias]
    backend.list_archives = archives
    item = source(tmp_path/'file.pdf'); result = mod.run_workspace(cfg, [item], backend)
    assert not result['errors']
    saved = list((tmp_path/'发票').rglob('*.pdf')); assert len(saved) == 1
    mod.run_workspace(cfg, [], backend)
    assert saved == list((tmp_path/'发票').rglob('*.pdf')) and backend.downloads == 1


def test_archive_for_unmapped_source_is_not_assumed_to_belong_to_all_inputs(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); original = backend.list_archives
    def archives(batch_id):
        return [dict(item, source_sha256='foreign') for item in original(batch_id)]
    backend.list_archives = archives
    result = mod.run_workspace(cfg, [source(tmp_path/'file.pdf')], backend)
    assert result['errors'] and backend.exports == backend.downloads == 0


def test_cross_month_invoices_in_same_original_require_split_without_publication(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); original = backend.list_archives
    def archives(batch_id):
        items = original(batch_id)
        return items + [dict(items[0], invoice_id='other-month', archive_filename='20260901_seller_other.pdf')]
    backend.list_archives = archives
    result = mod.run_workspace(cfg, [source(tmp_path/'file.pdf')], backend)
    assert result['pending_review_count'] == 1 and backend.exports == backend.downloads == 0
    assert '月份' in result['exceptions'][0]['reason']


def test_concurrent_ledger_edit_is_preserved_and_archive_is_not_published(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend()
    target = tmp_path/'台账文件/财务发票台账.xlsx'; target.parent.mkdir(); target.write_bytes(book(['old']))
    edited = book(['old', 'manual']); backend.change_ledger = lambda: target.write_bytes(edited)
    result = mod.run_workspace(cfg, [source(tmp_path/'file.pdf')], backend)
    assert result['errors'] and target.read_bytes() == edited and backend.downloads == 0


def test_same_name_same_size_documents_map_by_original_hash(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.blocked.add('same.pdf')
    first = source(tmp_path/'a/same.pdf', 'one'); second = tmp_path/'b/same.pdf'; second.parent.mkdir()
    # Equal length valid documents with different bytes.
    second.write_bytes(first.read_bytes().replace(b'/Producer', b'/ProducEr', 1))
    assert len(first.read_bytes()) == len(second.read_bytes()) and first.read_bytes() != second.read_bytes()
    result = mod.run_workspace(cfg, [first, second], backend)
    assert result['pending_review_count'] == 2 and not result['errors']
    assert len({v['source_sha256'] for v in result['exceptions']}) == 2


def test_inspect_review_is_scoped_and_reads_no_other_review(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.blocked.add('bad.pdf')
    item = mod.run_workspace(cfg, [source(tmp_path/'bad.pdf')], backend)['exceptions'][0]
    result = mod.inspect_review(cfg, item['review_ids'][0], backend)
    assert result['source_sha256'] == item['source_sha256'] and result['detail']['status'] == 'pending'
    with pytest.raises(ValueError): mod.inspect_review(cfg, 'foreign-review', backend)


def test_finance_entry_routes_desktop_config_and_retains_legacy_config(tmp_path, monkeypatch):
    entry = SCRIPTS/'run_finance.py'
    spec = importlib.util.spec_from_file_location('entry', entry); mod = importlib.util.module_from_spec(spec); spec.loader.exec_module(mod)
    seen = []
    import run_workspace, run_invoice_workflow
    monkeypatch.setattr(run_workspace, 'main', lambda args: seen.append(('workspace', args)) or 0)
    monkeypatch.setattr(run_invoice_workflow, 'main', lambda **kwargs: seen.append(('legacy', sys.argv[1:], kwargs)) or 0)
    target = tmp_path/'config.json'; target.write_text(json.dumps(config(tmp_path)))
    mod.main(['--config', str(target)]); assert seen[-1][0] == 'workspace'
    target.write_text(json.dumps({'ledger_path':'old.xlsx'}))
    mod.main(['--config', str(target)]); assert seen[-1][0] == 'legacy' and seen[-1][2]['expected_department'] == 'finance'


def test_silent_loss_of_historical_comment_refuses_ledger_replacement(tmp_path):
    from openpyxl.comments import Comment
    mod = module(); cfg = config(tmp_path); backend = Backend()
    target = tmp_path/'台账文件/old.xlsx'; target.parent.mkdir(); target.write_bytes(book(['old']))
    workbook = load_workbook(target); workbook.active['D2'].comment = Comment('不可丢失的付款批注','tester'); workbook.save(target); workbook.close()
    before = target.read_bytes()
    result = mod.run_workspace(cfg, [source(tmp_path/'file.pdf')], backend)
    assert result['errors'] and target.read_bytes() == before and backend.downloads == 0


def test_missing_ledger_rebuilds_all_previous_batches_without_recognition(tmp_path):
    mod = module(); cfg = config(tmp_path); backend = Backend()
    mod.run_workspace(cfg, [source(tmp_path/'one.pdf')], backend)
    first = mod.run_workspace(cfg, [source(tmp_path/'two.pdf', 'another')], backend)
    target = Path(first['ledger_path']); target.unlink()
    result = mod.run_workspace(cfg, [], backend)
    assert not result['errors'] and backend.calls == 2
    workbook = load_workbook(target); assert workbook.active.max_row == 3; workbook.close()


def test_modified_oldest_backup_prevents_replacement_and_is_never_deleted(tmp_path):
    mod = module(); cfg = config(tmp_path); paths = mod.layout(cfg)
    paths['backup_dir'].mkdir(); target = tmp_path/'ledger.xlsx'; target.write_bytes(b'original')
    state = {'backups':[]}
    for n in range(10):
        path = paths['backup_dir']/f'{n}.xlsx'; path.write_bytes(str(n).encode())
        state['backups'].append({'path':str(path),'sha256':mod.digest(path.read_bytes())})
    oldest = Path(state['backups'][0]['path']); oldest.write_bytes(b'manually changed')
    with pytest.raises(RuntimeError, match='人工修改'):
        mod.replace_ledger(target,b'original',b'new',paths,state)
    assert target.read_bytes() == b'original' and oldest.read_bytes() == b'manually changed'
    assert len(list(paths['backup_dir'].glob('*.xlsx'))) == 10


@pytest.mark.parametrize('archive_from_bad', [False, True])
def test_duplicate_occurrence_in_mixed_exception_pdf_never_silently_loses_clean_original(tmp_path, archive_from_bad):
    mod = module(); cfg = config(tmp_path); backend = Backend(); backend.blocked.add('mixed.pdf')
    original_archives = backend.list_archives
    def archives(batch_id):
        items = original_archives(batch_id)
        if archive_from_bad:
            mixed = next(f for f in backend.groups[batch_id] if f['original_filename'] == 'mixed.pdf')
            items[0]['source_sha256'] = items[0]['archive_sha256'] = mod.digest(mixed['content'])
        return items
    def invoice_sources(invoice_id, batch_id):
        return [{'document_id':f['document_id'], 'source_sha256':mod.digest(f['content'])} for f in backend.groups[batch_id]]
    backend.list_archives = archives; backend.invoice_sources = invoice_sources
    good = source(tmp_path/'good.pdf'); bad = source(tmp_path/'mixed.pdf', 'mixed')
    result = mod.run_workspace(cfg, [good, bad], backend)
    assert not result['errors']
    if archive_from_bad:
        assert result['pending_review_count'] == 2 and backend.exports == backend.downloads == 0
        assert any(item['source'] == 'good.pdf' and '归档原件' in item['reason'] for item in result['exceptions'])
    else:
        assert result['archive_writes'] == backend.exports == 1 and result['pending_review_count'] == 1
        assert next((tmp_path/'发票').rglob('*.pdf')).read_bytes() == good.read_bytes()
