"""Financial desktop workspace; extraction, identities and Excel formatting stay in the service."""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import uuid
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path

import fitz
from openpyxl import load_workbook

from run_invoice_workflow import HttpBackend, atomic_write, checked, digest, exclusive, workbook_rows, detected_format
from archive_index import ArchiveIndex, ArchiveConflict, identities
from workspace_runtime import prepare_runtime, runtime_exclusive
from exception_storage import exception_target, organize_legacy_exceptions, FOLDER_PATTERN
from backup_storage import create_backup, remove_backup, preflight_rotation, rotate_backups, organize_legacy_backups
from input_cleanup import capture_inputs, cleanup_completed_inputs

SUPPORTED = {'.pdf', '.png', '.jpg', '.jpeg'}
TZ = timezone(timedelta(hours=8))


def workbook_notes(data):
    if not data: return {}
    book = load_workbook(io.BytesIO(data), data_only=False)
    try:
        return {(sheet.title, cell.coordinate): cell.comment.text for sheet in book
                for row in sheet for cell in row if cell.comment and cell.comment.text}
    finally: book.close()


class WorkspaceBackend(HttpBackend):
    def get_batch(self, batch_id):
        return checked(self.client.get('/batches/' + batch_id)).json()

    def original_bytes(self, batch_id, document_id):
        return checked(self.client.get(f'/batches/{batch_id}/documents/{document_id}/original')).content

    def invoice_sources(self, invoice_id, batch_id):
        return self.archive_manifest(batch_id)['source_documents'].get(invoice_id, [])

    def archive_manifest(self, batch_id):
        if not hasattr(self, '_archive_manifests'): self._archive_manifests = {}
        if batch_id not in self._archive_manifests:
            manifest = checked(self.client.get(f'/batches/{batch_id}/ledger/manifest',
                params={'include_sources':'true'})).json()
            if 'source_documents' not in manifest or 'archive_identities' not in manifest:
                raise RuntimeError('后台缺少归档身份元数据能力，停止发布；请升级原服务')
            self._archive_manifests[batch_id] = manifest
        return self._archive_manifests[batch_id]

    def list_archives(self, batch_id):
        manifest = self.archive_manifest(batch_id)
        invoice_ids = set(manifest['archive_identities'])
        if set(manifest.get('invoice_ids', invoice_ids)) != invoice_ids:
            raise RuntimeError('归档身份元数据与本批发票集合不一致')
        rows = checked(self.client.get('/archives')).json()['items']
        selected = {}
        for item in rows:
            if item['invoice_id'] in invoice_ids and item['status'] == 'succeeded':
                selected.setdefault(item['invoice_id'], item)
        for invoice_id in sorted(invoice_ids - set(selected)):
            selected[invoice_id] = self.restore_archive(batch_id, invoice_id)
        return list(selected.values())

    def get_review(self, review_id):
        return checked(self.client.get('/manual-reviews/' + review_id)).json()

    def confirm_review(self, review_id, candidate, reason):
        return checked(self.client.post('/manual-reviews/' + review_id + '/confirm',
            json={'candidate': candidate, 'correction_reason': reason})).json()

    def export_selected(self, batch_id, invoice_ids, existing):
        files = {'existing_ledger': ('existing.xlsx', existing,
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')} if existing else None
        response = checked(self.client.post(f'/batches/{batch_id}/ledger', files=files,
            data={'selected_invoice_ids': json.dumps(invoice_ids)},
            params={'unique_invoices': 'true', 'normalize_template_styles': 'true'}))
        if (response.headers.get('X-Ledger-Unique-Invoices') != 'true'
                or response.headers.get('X-Ledger-Template-Normalized') != 'true'
                or response.headers.get('X-Ledger-Invoice-Selection') != 'explicit'):
            raise RuntimeError('后台缺少安全筛选与模板样式统一能力，未替换台账')
        self.last_export_report = {key: int(response.headers.get(header, 0)) for key, header in (
            ('historical_duplicate_groups', 'X-Historical-Duplicate-Groups'),
            ('skipped_duplicate_invoices', 'X-Skipped-Duplicate-Count'),
            ('batch_duplicate_occurrences', 'X-Batch-Duplicate-Occurrences'))}
        return response.content


def layout(config):
    root = Path(config['workspace_root']).resolve()
    defaults = {'input_dir': '待处理', 'archive_dir': '发票', 'ledger_dir': '台账文件',
                'exception_dir': '异常未处理发票', 'backup_dir': '备份', 'state_path': '.运行记录/财务状态.json'}
    result = {'root': root}
    for key, suffix in defaults.items():
        path = Path(config.get(key, root / suffix)).resolve()
        if not path.is_relative_to(root) or path == root:
            raise ValueError('输出及状态路径必须位于台账处理根目录的固定子目录')
        result[key] = path
    directories = [result[key] for key in ('input_dir', 'archive_dir', 'ledger_dir', 'exception_dir', 'backup_dir')]
    if len(set(directories)) != len(directories) or any(a.is_relative_to(b) for a in directories for b in directories if a != b):
        raise ValueError('输入、发票、台账、异常和备份目录不能重叠')
    return result


def select_ledger(config, paths):
    if config.get('ledger_path'):
        target = Path(config['ledger_path']).resolve()
        if target.parent != paths['ledger_dir'] or target.suffix.lower() != '.xlsx':
            raise ValueError('台账文件必须是台账文件目录中的xlsx')
        return target
    candidates = sorted(p for p in paths['ledger_dir'].glob('*.xlsx') if not p.name.startswith('~$'))
    if len(candidates) > 1:
        raise ValueError('发现多个台账，请指定需要追加的台账文件，不会自动覆盖任何一个')
    return candidates[0] if candidates else paths['ledger_dir'] / '财务发票台账.xlsx'


def scope_for(config, backend, paths, ledger):
    user = backend.identity()
    if user.get('department') != 'finance':
        raise ValueError('desktop workspace requires finance department')
    return digest(json.dumps([config['base_url'], user.get('user_id') or user.get('id') or user['username'],
                             str(paths['root']), str(ledger)], ensure_ascii=False).encode())


def preflight_source(path):
    if path.suffix.lower() not in SUPPORTED:
        return '暂不支持此文件格式，请提供PDF、PNG或JPG发票', False
    try:
        content = path.read_bytes()
        actual = detected_format(content)
        if not actual:
            return '文件内容签名不是受支持的PDF、PNG或JPEG，不能仅凭扩展名识别，请重新导出原件', False
        if actual == 'pdf':
            with fitz.open(stream=content, filetype='pdf') as document:
                if document.needs_pass: return 'PDF已加密，无法读取票面，请提供未加密文件', False
                if document.page_count < 1: return 'PDF没有可读取页面', False
                if document.page_count > 20: return 'PDF页数超过系统20页限制，请按发票拆分后提供', False
        else:
            fitz.Pixmap(content)
    except Exception:
        return '文件损坏或无法读取，不能可靠识别发票内容，请重新提供清晰原件', False
    return None, True


def safe_name(name):
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip(' .')
    if not value or value.split('.')[0].upper() in {'CON','PRN','AUX','NUL','COM1','LPT1'}:
        value = '发票_' + value
    return value[:140]


def find_source(record):
    for key in ('path', 'exception_path'):
        path = Path(record.get(key) or '__missing__')
        if path.is_file() and digest(path.read_bytes()) == record['sha256']:
            return path
    return None


def quarantine(record, reason, review_ids, paths):
    if record.get('secondary_result', {}).get('status') == 'deferred':
        reason = record['secondary_reason']
    target = exception_target(record, paths, safe_name(Path(record['name']).stem) + '__' + record['sha256'][:16] + Path(record['name']).suffix.lower())
    source = find_source(record)
    if target.exists() and digest(target.read_bytes()) != record['sha256']:
        raise RuntimeError('异常副本已被人工修改，保留原文件，停止覆盖')
    if not target.exists():
        if source is None: raise RuntimeError('原文件及异常副本均不可用，无法补存异常发票')
        atomic_write(target, source.read_bytes())
    record.update(exception_path=str(target), reason=reason, review_ids=review_ids, status='exception')
    note = target.with_suffix(target.suffix + '.原因.txt')
    contents = (reason + '\n').encode('utf-8')
    if not note.exists() or note.read_bytes() != contents: atomic_write(note, contents)
    return {'source_sha256': record['sha256'], 'source': record['name'], 'local_path': str(target),
            'reason': reason, 'review_ids': review_ids,
            'secondary_attempted': bool(record.get('secondary_attempted')),
            'can_secondary_review': bool(review_ids) and not record.get('secondary_attempted', False)}


def clear_quarantine(record, paths):
    target = Path(record.get('exception_path') or '__missing__')
    if target.is_file():
        if digest(target.read_bytes()) != record['sha256']:
            raise RuntimeError('已解决异常的副本被人工修改，保留副本，请人工核对')
        target.unlink()
        target.with_suffix(target.suffix + '.原因.txt').unlink(missing_ok=True)
        if target.parent.parent == paths['exception_dir'] and re.fullmatch(FOLDER_PATTERN,target.parent.name):
            try:target.parent.rmdir()
            except OSError:pass  # Preserve a group containing another exception or user file.
    for key in ('exception_path', 'reason', 'review_ids'):
        record.pop(key, None)
    record['status'] = 'complete'


def replace_ledger(target, before, output, paths, state, sources=()):
    if (target.read_bytes() if target.exists() else None) != before:
        raise RuntimeError('台账在处理期间被修改，已保留用户新版本，请重试')
    if output == before: return
    preflight_rotation(paths,state)
    backup = create_backup(target,before,sources,paths)
    try:
        if (target.read_bytes() if target.exists() else None) != before:
            raise RuntimeError('备份期间台账被人工修改，已保留用户版本，请重试')
        atomic_write(target, output)
    except Exception:
        remove_backup(backup,paths)
        raise
    state.setdefault('backups', []).append(backup)
    rotate_backups(paths,state)


def backup_batch_sources(records, batch_state, inventory, backend, batch_id):
    sources=[]
    for record in records:
        sha=record['sha256'];path=find_source(record) or inventory.get(sha)
        if path and path.is_file():content=path.read_bytes()
        else:
            document_id=next((doc for doc,value in batch_state.get('document_sources',{}).items() if value==sha),None)
            if not document_id:raise RuntimeError('本次处理的原发票缺失，无法完成备份，未替换台账')
            content=backend.original_bytes(batch_id,document_id)
        if digest(content)!=sha:raise RuntimeError('本次处理的发票已变化，无法完成原件备份，未替换台账')
        name=safe_name(Path(record['name']).stem)+Path(record['name']).suffix.lower()
        sources.append({'name':name,'sha256':sha,'content':content})
    return sources


def map_documents(batch, records, backend):
    mapping = {}
    for doc in batch.get('documents', []):
        # Filenames (including normalized upload suffixes) and lengths are not
        # identities. Reuse the upload receipt, or verify server original bytes once.
        sha = batch.get('upload_source_hashes', {}).get(doc['document_id'])
        if not sha: sha = digest(backend.original_bytes(batch['batch_id'], doc['document_id']))
        if sha not in {r['sha256'] for r in records}:
            raise RuntimeError('批次原文件与当前源哈希不一致，未发布归档或台账')
        mapping[doc['document_id']] = sha
    return mapping


def manifest_for(backend, batch_id, batch, revision=None):
    revision = revision if revision is not None else sorted(
        [row['review_id'], row['status']] for row in batch.get('result', {}).get('manual_reviews', []))
    if batch.get('identity_manifest') is None or batch.get('identity_revision') != revision:
        manifest = backend.archive_manifest(batch_id)
        if 'archive_identities' not in manifest or 'source_documents' not in manifest:
            raise RuntimeError('后台缺少归档身份元数据能力，停止发布；请升级原服务')
        batch.update(identity_manifest=manifest, identity_revision=revision)
    return batch['identity_manifest']


def sync_managed_index(state, index, inventory, backend):
    """Rebuild known historical metadata from stored receipts, without extraction."""
    groups = {}
    for invoice_id, saved in state.get('archives', {}).items():
        path = inventory.get(saved['sha256'])
        if path is None: continue
        row = index.files[index.relative(path)]
        if row['status'] == 'modified': continue
        saved['path'] = str(path)
        if row['status'] == 'verified': continue
        groups.setdefault(saved['sha256'], {'path': path, 'ids': set()})['ids'].add(invoice_id)
    for sha, group in groups.items():
        values = {}; evidence_batch = None
        for batch_id, batch in state.get('batches', {}).items():
            relevant = group['ids'] & {item['invoice_id'] for item in batch.get('archive_items', [])}
            if not relevant: continue
            manifest = manifest_for(backend, batch_id, batch)
            # Canonical archives can point X at a single-ticket PDF and Y at
            # another PDF containing both X and Y. Restore the latter's entire
            # validated source set, never merely its canonical archive item IDs.
            source_ids = {invoice_id for invoice_id, sources in manifest['source_documents'].items()
                          if any(row['source_sha256'] == sha for row in sources)}
            for invoice_id in source_ids:
                if invoice_id not in manifest['archive_identities']: raise RuntimeError('历史归档缺少已核验身份元数据')
                values[invoice_id] = manifest['archive_identities'][invoice_id]
            if source_ids: evidence_batch = batch_id
        if group['ids'] <= set(values):
            try:
                index.bind(group['path'], sha, list(values.values()), evidence_batch)
                index.record_alias(sha, list(values.values()), group['path'], evidence_batch)
            except ArchiveConflict as error:
                index.files[index.relative(group['path'])]['history_error'] = str(error)


def source_metadata(manifest, source_sha):
    invoice_ids = [invoice_id for invoice_id, sources in manifest['source_documents'].items()
                   if any(row['source_sha256'] == source_sha for row in sources)]
    if not invoice_ids: raise RuntimeError('缺少原件完整发票集合，暂停归档')
    values = [manifest['archive_identities'][invoice_id] for invoice_id in invoice_ids]
    identities(values)
    return values


def adopt_history_receipt(sha, record, state, index):
    known = index.source(sha)
    if not known: return False
    batch_id = known.get('batch_id'); job = state.get('history_batches', {}).get(batch_id)
    if not job: return False
    manifest = job.get('identity_manifest')
    if not manifest: return False
    matched_ids = {invoice_id for invoice_id, rows in manifest['source_documents'].items()
                   if any(row['source_sha256'] == sha for row in rows)}
    if not matched_ids: return False
    batch = state['batches'].setdefault(batch_id, {'source_hashes': [], 'key': job['key'],
        'result': job['result'], 'identity_manifest': manifest, 'identity_revision': job['identity_revision']})
    if sha not in batch['source_hashes']: batch['source_hashes'].append(sha)
    mapping = {row['document_id']: row['source_sha256'] for rows in manifest['source_documents'].values()
               for row in rows if row['source_sha256'] in batch['source_hashes']}
    selected = {invoice_id for invoice_id, rows in manifest['source_documents'].items()
                if any(row['source_sha256'] in batch['source_hashes'] for row in rows)}
    batch.update(document_sources=mapping,
        archive_items=[item for item in job['archive_items'] if item['invoice_id'] in selected],
        review_revision=job['identity_revision'])
    record['batch_id'] = batch_id
    return True


def index_history(config, backend, limit=20):
    """Bounded historical recognition only; never export a ledger or copy originals."""
    if not 1 <= limit <= 500: raise ValueError('历史建档每批数量应为1到500')
    paths = layout(config); ledger = select_ledger(config, paths); scope = scope_for(config, backend, paths, ledger)
    for key, path in paths.items():
        if key != 'root': (path.parent if key == 'state_path' else path).mkdir(parents=True, exist_ok=True)
    target = paths['state_path'].parent/'归档身份索引.json'
    result = {'operation': 'index_history', 'processed_files': 0, 'archive_writes': 0,
              'ledger_exports': 0, 'errors': [], 'unresolved_history': [], 'batch_ids': []}
    prepare_runtime(paths['root'])
    with runtime_exclusive(paths['archive_dir']/'.finance-archive.lock'), runtime_exclusive(paths['state_path'].with_suffix('.lock')), runtime_exclusive(target.with_suffix('.lock')):
        state = json.loads(paths['state_path'].read_text(encoding='utf-8')) if paths['state_path'].exists() else {
            'scope': scope, 'sources': {}, 'batches': {}, 'archives': {}, 'backups': []}
        if state['scope'] != scope: raise ValueError('账户、根目录或台账已改变，不能复用另一套账的状态')
        index = ArchiveIndex(paths['archive_dir'], target, scope); inventory = index.scan()
        def save():
            index.save(); atomic_write(paths['state_path'], json.dumps(state, ensure_ascii=False, indent=2).encode('utf-8'))
        sync_managed_index(state, index, inventory, backend)
        jobs = state.setdefault('history_batches', {})
        assigned = {sha for job in jobs.values() for sha in job['source_hashes']}
        candidates = []
        for name, row in index.files.items():
            if row['status'] != 'unverified' or row['sha256'] in assigned or row.get('history_error'): continue
            reason, _ = preflight_source(index.file_path(name))
            if reason: row['history_error'] = reason; continue
            if row['sha256'] not in {sha for _, sha in candidates} and len(candidates) < limit:
                candidates.append((index.file_path(name), row['sha256']))
        failed_this_run = set()
        if candidates:
            key = 'finance-history-' + uuid.uuid4().hex
            pending = {'source_hashes': [sha for _, sha in candidates], 'key': key}
            def on_batch(batch_id):
                jobs.setdefault(batch_id, pending); save()
            try:
                outcome = backend.process_many([path for path, _ in candidates], key, on_batch=on_batch,
                    expected_hashes={str(path): sha for path, sha in candidates})
                on_batch(outcome['batch_id']); jobs[outcome['batch_id']]['result'] = outcome
                result['processed_files'] += len(candidates) - len(outcome.get('upload_rejections', [])); save()
            except Exception as error:
                failed_this_run.update(batch_id for batch_id, job in jobs.items() if job['key'] == key)
                result['errors'].append({'error': str(error)}); save()
        for batch_id, job in jobs.items():
            result['batch_ids'].append(batch_id)
            if batch_id in failed_this_run: continue
            if job.get('indexed'):
                # Restoring a deleted index can use this persisted, validated result.
                if all(not any(row['sha256'] == sha and row['status'] != 'verified' for row in index.files.values())
                       for sha in job['source_hashes']): continue
            try:
                outcome = job.get('result')
                if not outcome or outcome['status'] not in {'completed', 'completed_with_errors', 'upload_rejected'}:
                    paths_for_job = [inventory[sha] for sha in job['source_hashes']]
                    outcome = backend.process_many(paths_for_job, job['key'], batch_id=batch_id,
                        expected_hashes={str(path): digest(path.read_bytes()) for path in paths_for_job})
                    result['processed_files'] += len(paths_for_job)
                elif job.get('pending_reviews'):
                    outcome = backend.get_batch(batch_id)
                job['result'] = outcome
                if outcome['status'] == 'upload_rejected':
                    for row in outcome.get('upload_rejections', []):
                        for file in index.files.values():
                            if file['sha256'] == row['source_sha256']: file['history_error'] = '系统拒绝该历史文件：'+row['code']
                    continue
                if outcome['status'] not in {'completed', 'completed_with_errors'}:
                    raise RuntimeError('历史核验批次未完成，保留原件及批次，稍后续跑')
                records = [{'sha256': sha, 'path': str(inventory[sha])} for sha in job['source_hashes'] if sha in inventory]
                mapping = job.get('document_sources') or map_documents(outcome, records, backend)
                job['document_sources'] = mapping
                blocked = {}
                pending_reviews = []
                for review in outcome.get('manual_reviews', []):
                    if review['status'] not in {'accepted', 'duplicate', 'skipped'}:
                        sha = mapping.get(review['document_id'])
                        if sha:
                            blocked[sha] = '历史原件需要系统审核后才能建立身份索引'
                            pending_reviews.append({'source_sha256': sha, 'review_id': review['review_id']})
                job['pending_reviews'] = pending_reviews
                for rejection in outcome.get('rejections', []):
                    blocked[mapping.get(rejection['document_id'])] = rejection.get('business_message') or '系统未能核验历史原件'
                accepted = outcome.get('accepted_invoice_count', 0) + outcome.get('skipped_invoice_count', 0)
                if accepted:
                    manifest = manifest_for(backend, batch_id, job)
                    if job.get('archive_items') is None or job.get('review_revision') != job['identity_revision']:
                        job['archive_items'] = backend.list_archives(batch_id)
                        job['review_revision'] = job['identity_revision']
                    for name, row in index.files.items():
                        sha = row['sha256']
                        if sha not in job['source_hashes'] or row['status'] in {'verified', 'modified'}: continue
                        if sha in blocked: row['history_error'] = blocked[sha]; continue
                        try:
                            values = source_metadata(manifest, sha)
                            index.bind(index.file_path(name), sha, values, batch_id)
                            index.record_alias(sha, values, index.file_path(name), batch_id)
                        except (ValueError, RuntimeError) as error: row['history_error'] = str(error)
                for row in index.files.values():
                    if row['sha256'] in blocked: row['history_error'] = blocked[row['sha256']]
                job['indexed'] = not pending_reviews
                save()
            except Exception as error: result['errors'].append({'batch_id': batch_id, 'error': str(error)}); save()
        result['archive_index'] = index.report()
        result['unresolved_history'] = [{'path': str(index.file_path(name)),
            'reason': row.get('history_error') or ('原件被修改，需人工核对' if row['status'] == 'modified' else '尚未建立身份索引')}
            for name, row in index.files.items() if row['status'] != 'verified']
        result['history_reviews'] = [review for job in jobs.values() for review in job.get('pending_reviews', [])]
        result['status'] = 'complete' if not result['unresolved_history'] and not result['errors'] else 'partial'
        save()
    return result


def archive_target(item, paths, month_layout='year/month'):
    name = item['archive_filename']
    if Path(name).name != name or re.search(r'[\\/]', name):
        raise ValueError('系统归档文件名无效')
    match = re.match(r'^(\d{8})_', name)
    if not match: raise ValueError('系统归档文件名缺少开票日期，不能猜测年月')
    date = datetime.strptime(match.group(1), '%Y%m%d')
    if month_layout == 'year.month':
        return paths['archive_dir'] / date.strftime('%Y.%m') / name
    if month_layout != 'year/month': raise ValueError('不支持的归档月份目录格式')
    return paths['archive_dir'] / f'{date.year:04d}' / f'{date.month:02d}' / name


def plan_archive_groups(normal, manifest, index, inventory, paths, month_layout):
    groups = {}
    for item in normal:
        groups.setdefault(item['archive_sha256'], {'items': []})['items'].append(item)
    planned = []; failures = {}
    for sha, group in groups.items():
        try:
            try:
                values = source_metadata(manifest, sha)
                complete_archive_source = True
            except RuntimeError:
                # A service archive can point at an older exact-duplicate source.
                # Only reuse a local verified complete set; do not publish old
                # bytes whose entire invoice set is absent from this manifest.
                values = [manifest['archive_identities'][item['invoice_id']] for item in group['items']]
                complete_archive_source = False
            normalized = identities(values)
            path = inventory.get(sha)
            if path and complete_archive_source and index.files[index.relative(path)]['status'] == 'unverified':
                index.bind(path, sha, values, group['items'][0].get('batch_id'))
            choice = index.choose(values, sha)
            target = archive_target(group['items'][0], paths, month_layout)
            if choice['action'] == 'new':
                if not complete_archive_source:
                    raise ArchiveConflict('系统归档原件的完整发票集合无法核对，暂缓发布并保留当前来源')
                for previous in planned:
                    overlap = set(previous['identities']) & set(normalized)
                    if not overlap: continue
                    if previous['identities'] != normalized:
                        raise ArchiveConflict('本批原件发票集合部分重复或关键字段冲突，请拆分后处理')
                    choice = {'action': 'reuse', 'path': previous['target'], 'sha256': previous['sha256']}
                    break
                if target.exists() and digest(target.read_bytes()) != sha:
                    raise ArchiveConflict('归档文件同名但内容不同，保留已有原件')
            group.update(sha256=sha, metadata=values, identities=normalized, choice=choice,
                         target=choice.get('path', target))
            planned.append(group)
        except (ArchiveConflict, ValueError) as error:
            failures[sha] = str(error)
    return planned, failures


def publish_indexed_group(group, paths, state, index, inventory, backend, batch_id, month_layout, warnings):
    expected = group['sha256']; choice = group['choice']; target = group['target']; written = 0
    if choice['action'] == 'reuse':
        if not target.is_file() or digest(target.read_bytes()) != choice['sha256']:
            raise ArchiveConflict('待复用归档在发布前发生变化，保留原件并停止发布')
        physical_sha = choice['sha256']
    else:
        if target.exists() and digest(target.read_bytes()) != expected:
            raise ArchiveConflict('归档文件同名但内容不同，保留原件')
        if not target.exists():
            content = backend.download_archive(group['items'][0]['archive_id'])
            if digest(content) != expected: raise RuntimeError('系统归档下载内容哈希不一致，未保存')
            written = publish_original(target, content, expected)
        physical_sha = expected
    index.bind(target, physical_sha, group['metadata'], batch_id)
    index.record_alias(expected, group['metadata'], target, batch_id)
    inventory[physical_sha] = target
    for item in group['items']:
        state['archives'][item['invoice_id']] = {'path': str(target), 'sha256': physical_sha}
    day = datetime.strptime(group['items'][0]['archive_filename'][:8], '%Y%m%d')
    if target.parent.relative_to(paths['archive_dir']) not in {Path(day.strftime('%Y.%m')), Path(day.strftime('%Y/%m'))}:
        warning = {'path': str(target), 'expected_directory': str(archive_target(group['items'][0], paths, month_layout).parent),
                   'reason': '已有同票原件的月份位置与票面日期不符；已复用，未移动或新增副本'}
        if warning not in warnings: warnings.append(warning)
    return written


def publish_original(target, content, expected):
    """Publish complete bytes atomically without replacing any existing original."""
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        with temporary.open('xb') as stream:
            stream.write(content); stream.flush(); os.fsync(stream.fileno())
        try:
            os.link(temporary, target)
            return 1
        except FileExistsError:
            if not target.is_file() or digest(target.read_bytes()) != expected:
                raise ArchiveConflict('归档发布时已有同名原件，内容不同，保留已有原件') from None
            return 0
    finally:
        temporary.unlink(missing_ok=True)


def run_workspace(config, inputs, backend):
    paths = layout(config); ledger = select_ledger(config, paths)
    month_layout = config.get('archive_month_layout', 'year/month')
    if month_layout not in {'year/month', 'year.month'}: raise ValueError('不支持的归档月份目录格式')
    scope = scope_for(config, backend, paths, ledger)
    # A malformed existing ledger must fail before submitting new documents.
    if ledger.exists(): workbook_rows(ledger.read_bytes())
    for key, path in paths.items():
        if key != 'root': (path.parent if key == 'state_path' else path).mkdir(parents=True, exist_ok=True)
    summary = {'processed_files': 0, 'duplicate_files': 0, 'resumed_files': 0, 'archive_writes': 0, 'ledger_exports': 0,
               'pending_review_count': 0, 'exceptions': [], 'errors': [], 'batch_ids': [], 'ledger_path': str(ledger),
               'archive_dir': str(paths['archive_dir']), 'archive_location_warnings': [], 'archive_identity_reuses': 0}
    index_path = paths['state_path'].parent/'归档身份索引.json'
    prepare_runtime(paths['root'])
    with runtime_exclusive(paths['archive_dir']/'.finance-archive.lock'), runtime_exclusive(paths['state_path'].with_suffix('.lock')), runtime_exclusive(ledger.with_suffix('.lock')), runtime_exclusive(index_path.with_suffix('.lock')):
        state = json.loads(paths['state_path'].read_text(encoding='utf-8')) if paths['state_path'].exists() else {
            'scope': scope, 'sources': {}, 'batches': {}, 'archives': {}, 'backups': []}
        if state['scope'] != scope: raise ValueError('账户、根目录或台账已改变，不能复用另一套账的状态')
        ledger_changed = state.get('ledger_sha256') != (digest(ledger.read_bytes()) if ledger.exists() else None)
        def save():
            index.save()
            atomic_write(paths['state_path'], json.dumps(state, ensure_ascii=False, indent=2).encode('utf-8'))
        index = ArchiveIndex(paths['archive_dir'], index_path, scope); inventory = index.scan()
        try:
            organize_legacy_backups(paths,state,save)
            organize_legacy_exceptions(paths,state,save)
            sync_managed_index(state, index, inventory, backend)
        except Exception as error:
            summary['errors'].append({'error': str(error)})
            summary['archive_index'] = index.report(); save()
            return summary
        unknown_history = index.report()['unverified_files']
        if not unknown_history:
            for record in state['sources'].values():
                if record.pop('archive_blocked', False):
                    record['status'] = 'new'
        seen = set()
        for raw in inputs:
            source = Path(raw).resolve()
            if not source.is_file():
                summary['errors'].append({'source': source.name, 'error': '输入文件不存在'}); continue
            content = source.read_bytes(); sha = digest(content)
            if sha in seen: summary['duplicate_files'] += 1; continue
            seen.add(sha)
            if sha in state['sources']:
                state['sources'][sha]['path'] = str(source)
                summary['resumed_files' if state['sources'][sha]['status'] == 'new' else 'duplicate_files'] += 1
                continue
            reason, readable = preflight_source(source)
            record = {'sha256': sha, 'path': str(source), 'name': source.name,
                      'byte_length': len(content), 'status': 'new'}
            state['sources'][sha] = record
            if reason: quarantine(record, reason, [], paths)
            elif adopt_history_receipt(sha, record, state, index):
                summary['duplicate_files'] += 1
            elif unknown_history and sha not in inventory:
                record['archive_blocked'] = True
                quarantine(record, '历史归档尚有未核验文件，请先执行历史身份建档；未重复调用识别', [], paths)
        new = [sha for sha, r in state['sources'].items() if r['status'] == 'new' and not r.get('batch_id')]
        save()
        failed_this_run = set()
        # Persist a batch id before upload/processing, so interruptions resume it.
        for offset in range(0, len(new), 500):
            hashes = new[offset:offset+500]; key = 'finance-workspace-' + uuid.uuid4().hex
            pending = {'source_hashes': hashes, 'key': key}
            def on_batch(batch_id, pending=pending):
                state['batches'].setdefault(batch_id, pending)
                for sha in pending['source_hashes']: state['sources'][sha]['batch_id'] = batch_id
                save()
            try:
                result = backend.process_many([Path(state['sources'][sha]['path']) for sha in hashes], key,
                    on_batch=on_batch, expected_hashes={state['sources'][sha]['path']: sha for sha in hashes})
                on_batch(result['batch_id']); state['batches'][result['batch_id']]['result'] = result
                summary['processed_files'] += len(hashes) - len(result.get('upload_rejections', [])); save()
            except Exception as error:
                failed_this_run.update(state['sources'][s]['batch_id'] for s in hashes if state['sources'][s].get('batch_id'))
                summary['errors'].append({'error': str(error), 'source': ', '.join(state['sources'][s]['name'] for s in hashes)})
        for batch_id, batch_state in state['batches'].items():
            summary['batch_ids'].append(batch_id)
            if batch_id in failed_this_run: continue
            records = [state['sources'][s] for s in batch_state['source_hashes']]
            try:
                if (not ledger_changed and batch_state.get('exported_ids')
                    and all(record['status'] == 'complete' for record in records)
                    and all(state['archives'].get(invoice_id, {}).get('sha256') in inventory
                            and index.files[index.relative(inventory[state['archives'][invoice_id]['sha256']])]['status'] == 'verified'
                            for invoice_id in batch_state['exported_ids'])):
                    continue
                result = batch_state.get('result')
                if result is None or result['status'] not in {'completed','completed_with_errors','failed','cancelled','upload_rejected'}:
                    result = backend.process_many([find_source(r) or Path(r['path']) for r in records], batch_state['key'],
                        batch_id=batch_id, on_batch=lambda _: None,
                        expected_hashes={str(find_source(r) or Path(r['path'])): r['sha256'] for r in records})
                    summary['processed_files'] += len(records) - len(result.get('upload_rejections', []))
                elif any(r.get('review_ids') for r in records):
                    result = backend.get_batch(batch_id)
                batch_state['result'] = result
                if 'upload_rejections' in result:
                    batch_state['upload_rejections'] = result['upload_rejections']
                upload_issues = {}
                for row in batch_state.get('upload_rejections', []):
                    sha = row['source_sha256']
                    if sha not in batch_state['source_hashes']:
                        raise RuntimeError('上传拒绝记录与批次源文件不匹配，暂停发布')
                    code = row['code']
                    reason = {'pdf_page_limit':'PDF页数超过系统限制，请按发票拆分后提供',
                        'file_too_large':'文件大小超过系统限制，请压缩或拆分后提供',
                        'pdf_encrypted':'PDF已加密，请提供未加密原件',
                        'pdf_structural_error':'PDF结构损坏，请重新导出原件',
                        'extension_signature_mismatch':'文件扩展名与内容签名不匹配，系统拒绝读取，请重新导出原件',
                        'mime_type_mismatch':'文件真实格式与上传类型不匹配，请重新导出原件',
                        'unsupported_file_type':'系统不支持此文件格式，请提供PDF、PNG或JPEG'}.get(code, '系统拒绝上传此文件，请重新提供原件')
                    upload_issues[sha] = [reason]
                    quarantine(state['sources'][sha], reason, [], paths)
                if result['status'] == 'upload_rejected':
                    save(); continue
                if result['status'] not in {'completed','completed_with_errors'}:
                    raise RuntimeError('批次未正常完成，保留原文件及处理状态')
                mapping = batch_state.get('document_sources') or map_documents(result, records, backend)
                batch_state['document_sources'] = mapping
                issues = upload_issues; reviews = {}
                for row in result.get('manual_reviews', []):
                    if row['status'] not in {'accepted', 'duplicate', 'skipped'}:
                        sha = mapping.get(row['document_id'])
                        if sha:
                            issues.setdefault(sha, []).extend(i.get('message') or i.get('code') for i in row.get('issues', []))
                            reviews.setdefault(sha, []).append(row['review_id'])
                for row in result.get('rejections', []):
                    sha = mapping.get(row['document_id'])
                    if sha: issues.setdefault(sha, []).append(row.get('business_message') or '系统无法可靠识别该文件，请提供清晰原件')
                accepted = result.get('accepted_invoice_count', 0) + result.get('skipped_invoice_count', 0)
                archive_items = batch_state.get('archive_items')
                revision = sorted([r['review_id'], r['status']] for r in result.get('manual_reviews', []))
                if archive_items is None or batch_state.get('review_revision') != revision:
                    archive_items = backend.list_archives(batch_id) if accepted else []
                    batch_state.update(archive_items=archive_items, review_revision=revision)
                manifest = manifest_for(backend, batch_id, batch_state, revision) if accepted else {
                    'archive_identities': {}, 'source_documents': {}}
                normal = []; invoice_sources = {}
                for item in archive_items:
                    source_hash = item.get('source_sha256')
                    matched = {source_hash} & set(batch_state['source_hashes'])
                    cached_sources = batch_state.setdefault('invoice_sources', {})
                    if cached_sources.get(item['invoice_id'], {}).get('revision') != revision:
                        sources = backend.invoice_sources(item['invoice_id'], batch_id)
                        cached_sources[item['invoice_id']] = {'revision': revision, 'items': sources}
                    sources = cached_sources[item['invoice_id']]['items']
                    matched |= {mapping[row['document_id']] for row in sources if row['document_id'] in mapping
                                and row['source_sha256'] == mapping[row['document_id']]}
                    if not matched:
                        raise RuntimeError('无法将归档发票绑定到正常源文件，暂停发布')
                    invoice_sources[item['invoice_id']] = matched
                months_by_content = {}
                for item in archive_items:
                    months_by_content.setdefault(item['archive_sha256'], set()).add(str(archive_target(item, paths, month_layout).parent))
                for item in archive_items:
                    if len(months_by_content[item['archive_sha256']]) > 1:
                        for sha in invoice_sources[item['invoice_id']]:
                            issues.setdefault(sha, []).append('同一原件包含不同月份发票，无法同时按月归档且避免重复，请按发票拆分原件后重新提供')
                # A duplicate occurrence inside a mixed PDF must not taint a
                # separate clean original. However the actual archived bytes
                # themselves must never come from an unresolved source.
                while True:
                    held = set(issues)
                    for item in archive_items:
                        if item.get('source_sha256') in issues:
                            for sha in invoice_sources[item['invoice_id']] - set(issues):
                                issues[sha] = ['系统当前归档原件来自仍有异常的合并文件，不能发布该原件；已保留此正常来源，待原系统提供正常归档原件后续作']
                    if set(issues) == held: break
                for record in records:
                    sha = record['sha256']
                    if sha not in issues and not any(sha in values for values in invoice_sources.values()):
                        issues[sha] = ['系统未识别到可入账发票，请核对是否为发票及票面是否完整']
                    if sha in issues:
                        quarantine(record, '；'.join(dict.fromkeys(issues[sha])), reviews.get(sha, []), paths)
                normal = [item for item in archive_items if invoice_sources[item['invoice_id']] - set(issues)
                          and item.get('source_sha256') not in issues]
                groups, archive_failures = plan_archive_groups(normal, manifest, index, inventory, paths, month_layout)
                for sha, reason in archive_failures.items():
                    for item in normal:
                        if item['archive_sha256'] == sha:
                            for source_sha in invoice_sources[item['invoice_id']]:
                                issues.setdefault(source_sha, []).append(reason)
                for record in records:
                    if record['sha256'] in issues:
                        quarantine(record, '；'.join(dict.fromkeys(issues[record['sha256']])),
                                   reviews.get(record['sha256'], []), paths)
                normal = [item for item in normal if invoice_sources[item['invoice_id']] - set(issues)
                          and item.get('source_sha256') not in issues]
                normal_ids = {item['invoice_id'] for item in normal}
                groups = [group for group in groups if all(item['invoice_id'] in normal_ids for item in group['items'])]
                ids = sorted(item['invoice_id'] for item in normal)
                if ids:
                    before = ledger.read_bytes() if ledger.exists() else None
                    needs_export = ledger_changed or batch_state.get('exported_ids') != ids or state.get('ledger_sha256') != (digest(before) if before else None)
                    if needs_export:
                        output = backend.export_selected(batch_id, ids, before)
                        before_rows = workbook_rows(before) if before else Counter()
                        after_rows = workbook_rows(output)
                        if any(after_rows[row] < n for row, n in before_rows.items()):
                            raise RuntimeError('系统输出改变了历史内容或人工备注，未替换原台账')
                        after_notes = workbook_notes(output)
                        if any(after_notes.get(key) != text for key, text in workbook_notes(before).items()):
                            raise RuntimeError('系统输出改变了历史批注，未替换原台账')
                        backup_records = {record['sha256']:record for record in records}
                        backup_records.update({sha:state['sources'][sha] for sha in sorted(seen)})
                        backup_sources = backup_batch_sources(backup_records.values(),batch_state,inventory,backend,batch_id) if output != before else []
                        replace_ledger(ledger, before, output, paths, state,backup_sources)
                        state['ledger_sha256'] = digest(output); batch_state['exported_ids'] = ids
                        summary['ledger_exports'] += 1; save()
                        batch_state['export_report'] = getattr(backend, 'last_export_report', {})
                    # Store only original-byte, verified backend archives after ledger validation.
                    for group in groups:
                        summary['archive_writes'] += publish_indexed_group(group, paths, state, index, inventory, backend,
                            batch_id, month_layout, summary['archive_location_warnings'])
                        if group['choice']['action'] == 'reuse' and group['choice']['sha256'] != group['sha256']:
                            summary['archive_identity_reuses'] += 1
                        for source_sha in batch_state['source_hashes']:
                            if source_sha in issues: continue
                            try:
                                values = source_metadata(manifest, source_sha)
                                if identities(values) == group['identities']:
                                    index.record_alias(source_sha, values, group['target'], batch_id)
                            except RuntimeError: pass
                        save()
                    for record in records:
                        if record['sha256'] not in issues: clear_quarantine(record,paths)
                save()
            except Exception as error:
                summary['errors'].append({'batch_id': batch_id, 'error': str(error)})
                save()
        for record in state['sources'].values():
            if record['status'] == 'exception':
                summary['exceptions'].append(quarantine(record, record['reason'], record.get('review_ids', []), paths))
        summary['pending_review_count'] = len(summary['exceptions'])
        summary['historical_duplicate_groups'] = max((b.get('export_report', {}).get('historical_duplicate_groups', 0)
            for b in state['batches'].values()), default=0)
        atomic_write(paths['state_path'].parent / '待复核队列.json', json.dumps(
            {'scope': scope, 'items': summary['exceptions']}, ensure_ascii=False, indent=2).encode('utf-8'))
        summary['review_queue'] = str(paths['state_path'].parent / '待复核队列.json')
        summary['archive_index'] = index.report()
        save()
    return summary


def inspect_review(config, review_id, backend):
    paths = layout(config); ledger = select_ledger(config, paths)
    scope = scope_for(config, backend, paths, ledger)
    prepare_runtime(paths['root'])
    with runtime_exclusive(paths['state_path'].with_suffix('.lock')):
        state = json.loads(paths['state_path'].read_text(encoding='utf-8'))
        if state['scope'] != scope: raise ValueError('审核账户与工作区不一致')
        record = next((r for r in state['sources'].values() if r['status'] == 'exception'
            and review_id in r.get('review_ids', [])), None)
        if not record or find_source(record) is None: raise ValueError('审核记录与可用的异常原件不匹配')
        if record.get('secondary_attempted'): raise ValueError('同一异常已完成二次判断，不自动重复读取')
        return {'source_sha256': record['sha256'], 'local_path': str(find_source(record)),
                'detail': backend.get_review(review_id)}


def resolve_exception(config, decision, backend):
    paths = layout(config); ledger = select_ledger(config, paths)
    scope = scope_for(config, backend, paths, ledger)
    prepare_runtime(paths['root'])
    with runtime_exclusive(paths['state_path'].with_suffix('.lock')):
        state = json.loads(paths['state_path'].read_text(encoding='utf-8'))
        if state['scope'] != scope: raise ValueError('审核账户与工作区不一致')
        sha = decision.get('source_sha256'); record = state['sources'].get(sha)
        if not record or record['status'] != 'exception' or find_source(record) is None:
            raise ValueError('source hash does not match an unchanged pending exception')
        reason = str(decision.get('reason') or '').strip()
        if not reason: raise ValueError('二次处理必须记录票面证据或无法处理原因')
        if record.get('secondary_attempted'): raise ValueError('同一异常已进行过二次处理，不会自动重复调用')
        if decision.get('action') == 'confirm':
            review_id = decision.get('review_id')
            if review_id not in record.get('review_ids', []): raise ValueError('审核记录与此源文件不匹配')
            pages = decision.get('pages')
            if not pages or not all(isinstance(p, int) and p >= 1 for p in pages):
                raise ValueError('必须提供已核对的票面页码')
            source = find_source(record)
            if source.suffix.lower() == '.pdf':
                with fitz.open(source) as doc:
                    if max(pages) > doc.page_count: raise ValueError('票面证据页码超出原文件')
            elif pages != [1]: raise ValueError('图片证据只能为第1页')
            detail = backend.get_review(review_id)
            if detail['status'] != 'pending': raise ValueError('审核记录已经处理，请按原配置续跑刷新')
            result = backend.confirm_review(review_id, decision['candidate'], reason)
        elif decision.get('action') == 'defer':
            result = {'status': 'deferred'}; record['reason'] = reason
        else: raise ValueError('仅支持confirm或defer二次处理动作')
        record.update(secondary_attempted=True, secondary_reason=reason, secondary_result=result)
        atomic_write(paths['state_path'], json.dumps(state, ensure_ascii=False, indent=2).encode('utf-8'))
        return result


def output_summary(result, result_file):
    """Keep model context bounded; full evidence is persisted locally."""
    summary = {key: value for key, value in result.items()
               if key not in {'exceptions', 'unresolved_history', 'history_reviews', 'batch_ids', 'errors'}}
    summary['result_file'] = str(result_file)
    summary['errors'] = result.get('errors', [])[:5]
    summary['error_count'] = len(result.get('errors', []))
    summary['batch_count'] = len(result.get('batch_ids', []))
    summary['batch_ids'] = result.get('batch_ids', [])[-3:]
    summary['exception_samples'] = [{key: str(row.get(key, ''))[:200] for key in ('source', 'reason')}
                                    for row in result.get('exceptions', [])[:3]]
    if 'unresolved_history' in result:
        summary['unresolved_history_count'] = len(result['unresolved_history'])
        summary['unresolved_history_samples'] = result['unresolved_history'][:3]
        summary['history_review_count'] = len(result.get('history_reviews', []))
    return summary


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument('--config', required=True); parser.add_argument('--input', nargs='*')
    parser.add_argument('--result')
    action = parser.add_mutually_exclusive_group()
    action.add_argument('--decision'); action.add_argument('--inspect-review')
    action.add_argument('--index-history', action='store_true')
    parser.add_argument('--history-limit', type=int, default=20)
    args = parser.parse_args(argv)
    config = json.loads(Path(args.config).read_text(encoding='utf-8-sig'))
    backend = WorkspaceBackend(config)
    cleanup_receipts = None
    try:
        if args.index_history:
            if args.input: parser.error('历史建档只核验归档目录，不能同时指定新输入')
            result = index_history(config, backend, args.history_limit)
        elif args.inspect_review:
            result = inspect_review(config, args.inspect_review, backend)
        elif args.decision:
            result = resolve_exception(config, json.loads(Path(args.decision).read_text(encoding='utf-8')), backend)
        else:
            paths = layout(config)
            inputs = [Path(p) for p in args.input] if args.input else sorted(
                [p for p in paths['input_dir'].rglob('*') if p.is_file() and not p.name.startswith('.')]
                + [p for p in paths['root'].iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED])
            cleanup_receipts = capture_inputs(inputs, paths['input_dir'])
            result = run_workspace(config, inputs, backend)
    finally:
        backend.close()
    target = Path(args.result) if args.result else layout(config)['state_path'].parent / (
        '历史建档结果.json' if args.index_history else '最近结果.json')
    atomic_write(target, json.dumps(result, ensure_ascii=False, indent=2).encode('utf-8'))
    if cleanup_receipts is not None:
        # Publication and its result must be durable before any inbox deletion.
        result.update(cleaned_input_files=0,retained_input_files=len(cleanup_receipts))
        if not result.get('errors'):
            try:
                cleanup=cleanup_completed_inputs(paths,Path(result['ledger_path']),cleanup_receipts)
                result['errors'].extend(cleanup.pop('errors'));result.update(cleanup)
            except Exception as error:
                result['errors'].append({'error':'待处理清理未完成，保留剩余原件：'+str(error)})
        atomic_write(target,json.dumps(result,ensure_ascii=False,indent=2).encode('utf-8'))
    summary = result if args.inspect_review or args.decision else output_summary(result, target)
    print(json.dumps(summary, ensure_ascii=False))
    return 2 if result.get('errors') or result.get('pending_review_count') or result.get('status') == 'partial' else 0


if __name__ == '__main__':
    raise SystemExit(main())
