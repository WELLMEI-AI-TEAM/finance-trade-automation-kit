import hashlib
import importlib.util
import json
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).parents[1]/'skills/finance-invoice-ledger/scripts'
sys.path.insert(0, str(SCRIPTS))


def module():
    path = SCRIPTS/'archive_index.py'
    assert path.exists(), 'local archive identity index missing'
    spec = importlib.util.spec_from_file_location('archive_index_under_test', path)
    result = importlib.util.module_from_spec(spec); spec.loader.exec_module(result)
    return result


def metadata(number='25332000000053136090', total='100.00', **changes):
    number = str(number).upper()
    kind = changes.get('invoice_type', 'vat_ordinary'); code = changes.get('invoice_code') or ''
    return {'identity_scheme':'cn_vat_v1',
        'identity_key':hashlib.sha256(f'cn_vat_v1|{kind}|{code}|{number}'.encode()).hexdigest(),
        'invoice_type':kind, 'invoice_code':code or None, 'invoice_number_full':number,
        'issue_date':'2026-10-01', 'buyer_name':'Buyer', 'buyer_tax_id':'BUYER-TAX',
        'seller_name':'Seller', 'seller_tax_id':'SELLER-TAX', 'amount_ex_tax':'100.00',
        'tax_amount':'0.00', 'total_amount':total, 'currency_code':'CNY', 'is_red_letter':False, **changes}


def setup(tmp_path):
    mod=module(); root=tmp_path/'发票'; root.mkdir(exist_ok=True)
    index=mod.ArchiveIndex(root,tmp_path/'归档身份索引.json','owner-scope')
    return mod,root,index


def test_same_identity_different_bytes_reuses_existing_file(tmp_path):
    mod,root,index=setup(tmp_path); old=root/'2026.09/old.pdf';old.parent.mkdir();old.write_bytes(b'original')
    index.scan(); index.bind(old,mod.digest(b'original'),[metadata()],'history-batch');index.save()
    choice=index.choose([metadata(total='100')],mod.digest(b'another encoding'))
    assert choice['action']=='reuse' and choice['path']==old and choice['sha256']==mod.digest(b'original')
    index.record_alias(mod.digest(b'another encoding'),[metadata()],old)
    index.save()
    again=mod.ArchiveIndex(root,tmp_path/'归档身份索引.json','owner-scope');again.scan()
    assert again.source(mod.digest(b'another encoding'))['path']==old
    assert len(list(root.rglob('*.pdf')))==1


def test_same_identity_different_financial_fields_is_a_conflict(tmp_path):
    mod,root,index=setup(tmp_path);old=root/'old.pdf';old.write_bytes(b'original')
    index.scan();index.bind(old,mod.digest(b'original'),[metadata()],'history')
    with pytest.raises(mod.ArchiveConflict,match='冲突'):
        index.choose([metadata(total='999')],mod.digest(b'changed'))


def test_unverified_history_prevents_claiming_a_new_invoice_is_unique(tmp_path):
    mod,root,index=setup(tmp_path);(root/'unknown.pdf').write_bytes(b'unverified')
    index.scan()
    assert index.report()['unverified_files']==1
    with pytest.raises(mod.ArchiveConflict,match='历史'):
        index.choose([metadata()],mod.digest(b'new source'))


def test_uncertain_historical_image_is_not_verified_from_filename(tmp_path):
    mod,root,index=setup(tmp_path);(root/'25332000000053136090.jpg').write_bytes(b'not an invoice image')
    index.scan();index.save()
    assert index.report()['verified_files']==0
    assert index.report()['unverified_files']==1


def test_scan_caches_unchanged_files_and_relocates_by_hash(tmp_path):
    mod,root,index=setup(tmp_path);old=root/'old.pdf';old.write_bytes(b'original')
    index.scan();index.bind(old,mod.digest(b'original'),[metadata()],'history');index.save()
    fresh=mod.ArchiveIndex(root,tmp_path/'归档身份索引.json','owner-scope');fresh.scan()
    assert fresh.statistics['hashes_read']==0 and fresh.statistics['cached_files']==1
    moved=root/'2026.10/renamed.pdf';moved.parent.mkdir();old.rename(moved)
    fresh.scan(); assert fresh.choose([metadata()],mod.digest(b'other'))['path']==moved


def test_changed_original_is_not_treated_as_missing_and_replaced(tmp_path):
    mod,root,index=setup(tmp_path);old=root/'old.pdf';old.write_bytes(b'original')
    index.scan();index.bind(old,mod.digest(b'original'),[metadata()],'history');index.save()
    old.write_bytes(b'manually modified')
    index.scan()
    with pytest.raises(mod.ArchiveConflict): index.choose([metadata()],mod.digest(b'new source'))
    assert old.read_bytes()==b'manually modified'


def test_reuse_verifies_bytes_even_if_stat_cache_was_unchanged(tmp_path):
    import os
    mod,root,index=setup(tmp_path);old=root/'old.pdf';old.write_bytes(b'original')
    index.scan();index.bind(old,mod.digest(b'original'),[metadata()],'history');index.save()
    stat=old.stat();old.write_bytes(b'changed!');os.utime(old,ns=(stat.st_atime_ns,stat.st_mtime_ns))
    with pytest.raises(mod.ArchiveConflict): index.choose([metadata()],mod.digest(b'new source'))


def test_multi_invoice_source_must_have_exact_identity_set_to_reuse(tmp_path):
    mod,root,index=setup(tmp_path);old=root/'old.pdf';old.write_bytes(b'original')
    index.scan();index.bind(old,mod.digest(b'original'),[metadata()],'history')
    with pytest.raises(mod.ArchiveConflict,match='部分|集合|拆分'):
        index.choose([metadata(),metadata('25332000000053136091')],mod.digest(b'mixed'))


def test_missing_original_allows_recovery_but_modified_original_does_not(tmp_path):
    mod,root,index=setup(tmp_path);old=root/'old.pdf';old.write_bytes(b'original')
    index.scan();index.bind(old,mod.digest(b'original'),[metadata()],'history');index.save()
    old.unlink();index.scan()
    assert index.choose([metadata()],mod.digest(b'new source'))['action']=='new'


def test_index_paths_are_relative_and_scope_isolation_is_enforced(tmp_path):
    mod,root,index=setup(tmp_path);old=root/'old.pdf';old.write_bytes(b'original')
    index.scan();index.bind(old,mod.digest(b'original'),[metadata()],'history');index.save()
    data=json.loads((tmp_path/'归档身份索引.json').read_text(encoding='utf-8'))
    assert 'old.pdf' in data['files'] and str(root) not in json.dumps(data,ensure_ascii=False)
    with pytest.raises(ValueError,match='范围|账户'):
        mod.ArchiveIndex(root,tmp_path/'归档身份索引.json','other-owner')


def test_invalid_backend_identity_key_cannot_seed_local_index(tmp_path):
    mod,root,index=setup(tmp_path);old=root/'old.pdf';old.write_bytes(b'original');index.scan()
    with pytest.raises(ValueError,match='身份'):
        index.bind(old,mod.digest(b'original'),[metadata(identity_key='0'*64)],'history')


def test_preexisting_duplicate_archives_are_reported_and_preserved(tmp_path):
    mod,root,index=setup(tmp_path)
    for name,content in [('one.pdf',b'one'),('two.pdf',b'two')]:
        path=root/name;path.write_bytes(content);index.scan();index.bind(path,mod.digest(content),[metadata()],'history')
    report=index.report()
    assert report['duplicate_identity_groups']==1
    assert len(list(root.glob('*.pdf')))==2


def test_cached_unknown_exact_copy_in_old_index_adopts_verified_evidence(tmp_path):
    mod,root,index=setup(tmp_path);original=root/'one.pdf';copy=root/'two.pdf'
    original.write_bytes(b'original');copy.write_bytes(b'original');index.scan()
    index.bind(original,mod.digest(b'original'),[metadata()],'history')
    # A previous release could leave an identical path without business evidence.
    index.files['two.pdf'].update(status='unverified',identities={});index.save()
    fresh=mod.ArchiveIndex(root,tmp_path/'归档身份索引.json','owner-scope');fresh.scan()
    assert fresh.report()['verified_files']==2
