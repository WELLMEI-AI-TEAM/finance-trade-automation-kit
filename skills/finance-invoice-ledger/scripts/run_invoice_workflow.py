"""Deterministic HTTP adapter; invoice extraction and writing stay in the backend."""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import mimetypes
import os
import re
import time
import uuid
from collections import Counter
from contextlib import contextmanager
from pathlib import Path

import httpx
from openpyxl import load_workbook


def digest(data):
    return hashlib.sha256(data).hexdigest()


def atomic_write(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        with temporary.open('xb') as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


@contextmanager
def exclusive(path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    # Kernel locks are released after a crash; never unlink the shared guard inode.
    descriptor = os.open(str(path) + '.guard', os.O_CREAT | os.O_RDWR, 0o600)
    locked = False
    try:
        if os.fstat(descriptor).st_size == 0:
            os.write(descriptor, b'0')
        os.lseek(descriptor, 0, os.SEEK_SET)
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(descriptor, msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
        locked = True
        if path.exists():
            # Safely migrate a legacy PID lock. Unknown owners remain protected.
            pid = int(path.read_text().strip())
            if pid <= 0 or process_alive(pid):
                raise RuntimeError('legacy lock owner is active or unknown; retry after it exits')
            path.unlink()
        yield
    finally:
        if locked:
            os.lseek(descriptor, 0, os.SEEK_SET)
            if os.name == 'nt':
                msvcrt.locking(descriptor, msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(descriptor, fcntl.LOCK_UN)
        os.close(descriptor)


def process_alive(pid):
    if os.name == 'nt':
        import ctypes
        from ctypes import wintypes
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel.OpenProcess.restype = wintypes.HANDLE
        kernel.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        handle = kernel.OpenProcess(0x1000, False, pid)  # query only
        if not handle:
            return ctypes.get_last_error() != 87  # invalid PID is known dead
        try:
            code = wintypes.DWORD()
            return not kernel.GetExitCodeProcess(handle, ctypes.byref(code)) or code.value == 259
        finally:
            kernel.CloseHandle(handle)
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True


def unresolved_reviews(result):
    return any(row.get('status') not in {'accepted', 'duplicate', 'skipped'}
               for row in result.get('manual_reviews', []))


def workbook_rows(data):
    """Read-only fingerprints, not an alternative invoice identity implementation."""
    book = load_workbook(io.BytesIO(data), data_only=False)
    rows = Counter()
    for sheet in book:
        for values in sheet.iter_rows(min_row=2, values_only=True):
            normalized = tuple('' if value is None else str(value) for value in values)
            while normalized and normalized[-1] == '':
                normalized = normalized[:-1]
            if any(normalized):
                rows[(sheet.title, normalized)] += 1
    book.close()
    return rows


def summarize_usage(rows):
    unique = {row['invocation_id']: row for row in rows}
    complete = all(all(row.get(key) is not None for key in
                       ('input_tokens', 'output_tokens', 'total_tokens')) for row in unique.values())
    return {'model_calls': len(unique), **{
        key: sum(row[key] for row in unique.values()) if complete else None
        for key in ('input_tokens', 'output_tokens', 'total_tokens')}, 'usage_complete': complete}


class BackendHTTPError(RuntimeError):
    def __init__(self, status_code, code):
        self.status_code = status_code
        self.code = code
        super().__init__(f'{status_code}: {code}')


UPLOAD_REJECTION_CODES = {'unsupported_file_type', 'file_too_large', 'extension_signature_mismatch',
    'pdf_structural_error', 'pdf_encrypted', 'pdf_page_limit', 'mime_type_mismatch'}


def detected_format(content):
    if content.startswith(b'%PDF-'): return 'pdf'
    if content.startswith(b'\x89PNG\r\n\x1a\n'): return 'png'
    if content.startswith(b'\xff\xd8\xff'): return 'jpeg'
    return None


def upload_metadata(path, content):
    """Correct a supported file's upload name/MIME, without rewriting its bytes."""
    name = path.name
    actual = detected_format(content)
    expected = {'.pdf':'pdf', '.png':'png', '.jpg':'jpeg', '.jpeg':'jpeg'}.get(path.suffix.lower())
    if actual and expected:
        if actual != expected:
            name = path.stem + {'pdf':'.pdf', 'png':'.png', 'jpeg':'.jpg'}[actual]
        return name, {'pdf':'application/pdf', 'png':'image/png', 'jpeg':'image/jpeg'}[actual]
    return name, mimetypes.guess_type(name)[0] or 'application/octet-stream'


def checked(response):
    if not response.is_success:
        # Never echo request bodies or authentication material.
        try:
            payload = response.json()
            detail = payload.get('error') or payload.get('detail', {}) if isinstance(payload, dict) else {}
            code = detail.get('code', 'http_error') if isinstance(detail, dict) else 'http_error'
        except ValueError:
            code = 'http_error'
        raise BackendHTTPError(response.status_code, code)
    return response


class HttpBackend:
    def __init__(self, config):
        self.config = config
        self.client = httpx.Client(base_url=config['base_url'].rstrip('/') + '/api/v1',
                                   trust_env=False, timeout=180)
        username = config['username']
        if config.get('credentials_file'):
            raw = Path(config['credentials_file']).read_bytes()
            try:
                contents = raw.decode('utf-8-sig')
            except UnicodeDecodeError:
                contents = raw.decode('gb18030')
            parts = next(line.split() for line in contents.splitlines() if username in line.split())
            password = parts[parts.index(username) + 1]
        else:
            password = os.environ[config.get('password_env', 'INVOICE_SKILL_PASSWORD')]
        self.user = checked(self.client.post('/auth/login', json={
            'username': username, 'password': password})).json()['user']

    def identity(self):
        return self.user

    def close(self):
        self.client.post('/auth/logout')
        self.client.close()

    def process(self, path, key, batch_id=None, on_batch=None):
        return self.process_many([path], key, batch_id, on_batch)

    def process_many(self, paths, key, batch_id=None, on_batch=None, expected_hashes=None):
        upload_rejections = []
        upload_source_hashes = {}
        if not batch_id:
            request = {'client_session_id': 'invoice-skill-' + key[:32]}
            batch = checked(self.client.post('/batches', json=request,
                            headers={'Idempotency-Key': key})).json()
            batch_id = batch['batch_id']
            if on_batch: on_batch(batch_id)
        batch = checked(self.client.get('/batches/' + batch_id)).json()
        if batch['status'] in {'created', 'draft'}:
            for path in paths:
                content = path.read_bytes()
                if expected_hashes and digest(content) != expected_hashes[str(path)]:
                    raise RuntimeError('source changed since batch creation; original batch retained')
                name, mime = upload_metadata(path, content)
                try:
                    uploaded = checked(self.client.post(f'/batches/{batch_id}/files', headers={
                        'Idempotency-Key': key + '-' + digest(content)}, files={'file': (name, content, mime)})).json()
                    upload_source_hashes[uploaded['document_id']] = digest(content)
                except BackendHTTPError as error:
                    if error.status_code not in {413, 415, 422} or error.code not in UPLOAD_REJECTION_CODES:
                        raise
                    upload_rejections.append({'source_sha256': digest(content), 'code': error.code})
            # A local terminal outcome: keep the empty server batch for audit and
            # do not submit an empty job. Transport/auth/conflict failures still stop.
            if upload_rejections:
                batch = checked(self.client.get('/batches/' + batch_id)).json()
                if not batch.get('documents'):
                    return {**batch, 'status':'upload_rejected', 'upload_rejections':upload_rejections}
            checked(self.client.post(f'/batches/{batch_id}/submit',
                                    headers={'Idempotency-Key': key + '-submit'}))
        deadline = time.monotonic() + self.config.get('wait_seconds', 900)
        while True:
            batch = checked(self.client.get('/batches/' + batch_id)).json()
            if batch['status'] in {'completed', 'completed_with_errors', 'failed', 'cancelled'}:
                if upload_rejections: batch['upload_rejections'] = upload_rejections
                if upload_source_hashes: batch['upload_source_hashes'] = upload_source_hashes
                return batch
            if time.monotonic() >= deadline:
                raise TimeoutError('batch still running; rerun the same configuration to resume')
            time.sleep(self.config.get('poll_seconds', 3))

    def list_archives(self, batch_id):
        manifest = checked(self.client.get(f'/batches/{batch_id}/ledger/manifest')).json()
        invoice_ids = set(manifest['invoice_ids'])
        items = checked(self.client.get('/archives')).json()['items']
        selected = {}
        for item in items:
            if item['invoice_id'] in invoice_ids and item['status'] == 'succeeded':
                selected.setdefault(item['invoice_id'], item)
        for invoice_id in sorted(invoice_ids - set(selected)):
            selected[invoice_id] = self.restore_archive(batch_id, invoice_id)
        return list(selected.values())

    def restore_archive(self, batch_id, invoice_id):
        item = checked(self.client.post('/archives/restore', json={
            'batch_id': batch_id, 'invoice_id': invoice_id})).json()
        if item.get('status') != 'succeeded':
            raise RuntimeError('backend could not restore validated archive; recognition was not repeated')
        return item

    def download_archive(self, archive_id):
        return checked(self.client.get('/archives/' + archive_id + '/download')).content

    def export_ledger(self, batch_id, existing):
        files = {'existing_ledger': ('existing.xlsx', existing,
                 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')} if existing else None
        response = checked(self.client.post('/batches/' + batch_id + '/ledger', files=files,
                                           params={'unique_invoices': 'true'}))
        if response.headers.get('X-Ledger-Unique-Invoices') != 'true':
            raise RuntimeError('backend lacks unique export support; upgrade backend before publishing ledger')
        self.last_export_report = {
            'historical_duplicate_groups': int(response.headers.get('X-Historical-Duplicate-Groups', 0)),
            'skipped_duplicate_invoices': int(response.headers.get('X-Skipped-Duplicate-Count', 0)),
            'batch_duplicate_occurrences': int(response.headers.get('X-Batch-Duplicate-Occurrences', 0)),
        }
        return response.content


def run_workflow(config, paths, backend):
    ledger = Path(config['ledger_path']).resolve()
    archive = Path(config['archive_dir']).resolve()
    state_path = Path(config['state_path']).resolve()
    if ledger.suffix.lower() != '.xlsx':
        raise ValueError('ledger_path must be an .xlsx file')
    if state_path == ledger or archive == ledger.parent and ledger.name.endswith('.pdf'):
        raise ValueError('overlapping output configuration')
    identity = backend.identity()
    expected_department = config.get('expected_department')
    if expected_department and identity.get('department') != expected_department:
        raise ValueError(f'department mismatch: expected {expected_department}, got {identity.get("department")}')
    scope = digest(json.dumps([config['base_url'], identity.get('user_id') or identity.get('id')
                    or identity['username'], identity.get('department'), str(ledger), str(archive)],
                    ensure_ascii=False).encode())
    with exclusive(state_path.with_suffix('.lock')), exclusive(archive / '.invoice-skill.lock'):
        state = json.loads(state_path.read_text(encoding='utf-8')) if state_path.exists() else {
            'scope': scope, 'sources': {}, 'archives': {}}
        if state['scope'] != scope:
            raise ValueError('configuration/account changed; use a separate state_path')
        state.setdefault('batches', {})
        for record in state['sources'].values():
            if record.get('batch_id'):
                batch_state = state['batches'].setdefault(record['batch_id'], {})
                for field in ('result', 'archive_items'):
                    if field in record:
                        batch_state.setdefault(field, record.pop(field))

        def result_for(record):
            return state['batches'].get(record.get('batch_id'), {}).get('result')
        index_path = archive / '.invoice-archive-index.json'
        archive_scope = digest(json.dumps([config['base_url'], identity.get('user_id') or identity.get('id')
                              or identity['username'], identity.get('department')]).encode())
        index = json.loads(index_path.read_text(encoding='utf-8')) if index_path.exists() else {
            'scope': archive_scope, 'archives': {}}
        if index['scope'] != archive_scope:
            raise ValueError('archive directory belongs to another backend/account; select a separate directory')
        state['archives'] = {**state.get('archives', {}), **index['archives']}
        inventory = {digest(path.read_bytes()): path for path in archive.iterdir()
                     if path.is_file() and path.suffix.lower() in {'.pdf', '.png', '.jpg', '.jpeg'}}
        initial = ledger.read_bytes() if ledger.exists() else None
        initial_rows = workbook_rows(initial) if initial else Counter()
        if state.get('ledger_sha256') != (digest(initial) if initial else None):
            for record in state['sources'].values():
                record['ledger_done'] = False
        summary = {'input_files': len(paths), 'duplicate_files': 0, 'already_complete': 0,
                   'processed_files': 0, 'archive_writes': 0, 'ledger_exports': 0,
                   'historical_duplicate_groups': sum(n > 1 for n in initial_rows.values()),
                   'errors': [], 'batch_ids': [], 'archive_dir': str(archive), 'ledger_path': str(ledger)}
        known_paths = {str(Path(item['path']).resolve()) for item in state['archives'].values()}
        summary['unindexed_archive_files'] = sorted(path.name for path in inventory.values() if str(path.resolve()) not in known_paths)

        def save():
            atomic_write(state_path, json.dumps(state, ensure_ascii=False, indent=2).encode('utf-8'))
            atomic_write(index_path, json.dumps({'scope': archive_scope, 'archives': state['archives']},
                                               ensure_ascii=False, indent=2).encode('utf-8'))

        seen = set()
        sources = []
        for raw_path in paths:
            path = Path(raw_path).resolve()
            source_hash = digest(path.read_bytes())
            if source_hash in seen:
                summary['duplicate_files'] += 1
                continue
            seen.add(source_hash)
            record = state['sources'].setdefault(source_hash, {'path': str(path),
                       'key': 'invoice-skill-' + uuid.uuid4().hex, 'ledger_done': False})
            record['path'] = str(path)
            sources.append((path, source_hash, record))
        for source_hash, record in state['sources'].items():
            if source_hash not in seen and record.get('batch_id') and (
                    not record.get('ledger_done') or result_for(record) is None):
                sources.append((Path(record['path']), source_hash, record))
        initially_complete = {source_hash for _, source_hash, record in sources if record.get('ledger_done')}
        failed_sources = set()
        if hasattr(backend, 'process_many'):
            groups = {}
            for path, source_hash, record in sources:
                if result_for(record) is None or unresolved_reviews(result_for(record)):
                    groups.setdefault(record.get('batch_id'), []).append((path, source_hash, record))
            for batch_id, group in groups.items():
                if batch_id:
                    group = [(Path(record['path']), source_hash, record)
                             for source_hash, record in state['sources'].items() if record.get('batch_id') == batch_id]
                chunks = [group] if batch_id else [group[i:i + 500] for i in range(0, len(group), 500)]
                for chunk in chunks:
                    chunk = sorted(chunk, key=lambda item: item[1])
                    key = 'invoice-skill-batch-' + digest('|'.join(record['key'] for _, _, record in chunk).encode())
                    def on_batch(value, chunk=chunk):
                        for _, _, record in chunk:
                            record['batch_id'] = value
                        save()
                    try:
                        save()
                        result = backend.process_many([path for path, _, _ in chunk], key, batch_id, on_batch,
                                                      expected_hashes={str(path): source_hash for path, source_hash, _ in chunk})
                        for _, _, record in chunk:
                            record['batch_id'] = result['batch_id']
                        state['batches'].setdefault(result['batch_id'], {})['result'] = result
                        summary['processed_files'] += len(chunk)
                        save()
                    except Exception as error:
                        failed_sources.update(source_hash for _, source_hash, _ in chunk)
                        summary['errors'].append({'source': ', '.join(path.name for path, _, _ in chunk),
                                                  'error': str(error), 'batch_id': chunk[0][2].get('batch_id')})
                        save()
        handled_batches = set()
        for path, source_hash, record in sources:
            if source_hash in failed_sources:
                continue
            try:
                if result_for(record) is None:
                    def on_batch(batch_id):
                        record['batch_id'] = batch_id
                        save()
                    save()
                    result = backend.process(path, record['key'], record.get('batch_id'), on_batch)
                    record['batch_id'] = result['batch_id']
                    state['batches'].setdefault(result['batch_id'], {})['result'] = result
                    save()
                    summary['processed_files'] += 1
                result = result_for(record)
                batch_id = record['batch_id']
                if batch_id in handled_batches:
                    if source_hash in initially_complete and record.get('ledger_done'):
                        summary['already_complete'] += 1
                    continue
                handled_batches.add(batch_id)
                batch_state = state['batches'][batch_id]
                if batch_id not in summary['batch_ids']:
                    summary['batch_ids'].append(batch_id)
                if result['status'] not in {'completed', 'completed_with_errors'}:
                    raise RuntimeError('batch did not complete successfully')
                if unresolved_reviews(result) or result.get('rejected_invoice_count', 0):
                    raise RuntimeError('backend requires review or rejected a document; inspect batch')
                if not result.get('accepted_invoice_count', 0) and not result.get('skipped_invoice_count', 0):
                    raise RuntimeError('batch contains no validated invoice')
                if 'archive_items' not in batch_state:
                    batch_state['archive_items'] = backend.list_archives(batch_id)
                    save()
                if not batch_state['archive_items']:
                    raise RuntimeError('backend archive unavailable; repair archive without recognition')
                archive_changed = False
                for item in batch_state['archive_items']:
                    key = item['invoice_id']
                    existing = state['archives'].get(key)
                    if existing:
                        target = Path(existing['path'])
                        if target.exists() and digest(target.read_bytes()) == existing['sha256']:
                            continue
                        if not target.exists() and existing['sha256'] in inventory:
                            existing['path'] = str(inventory[existing['sha256']])
                            save()
                            continue
                    else:
                        filename = item['archive_filename']
                        if Path(filename).name != filename or re.search(r'[\\/]', filename):
                            raise ValueError('invalid backend archive filename')
                        target = archive / filename
                    try:
                        content = backend.download_archive(item['archive_id'])
                    except RuntimeError:
                        expected = item.get('archive_sha256')
                        candidates = [Path(source['path']) for source in state['sources'].values()]
                        content = next((candidate.read_bytes() for candidate in candidates
                                        if candidate.is_file() and expected and digest(candidate.read_bytes()) == expected.lower()), None)
                        if content is None:
                            if not hasattr(backend, 'restore_archive'):
                                raise
                            restored = backend.restore_archive(batch_id, item['invoice_id'])
                            content = backend.download_archive(restored['archive_id'])
                            item.update(restored)
                    content_hash = digest(content)
                    expected = item.get('archive_sha256')
                    if expected and content_hash != expected.lower():
                        raise ValueError('archive integrity mismatch')
                    if content_hash in inventory:
                        target = inventory[content_hash]
                    elif not existing and any(str(path.resolve()) not in {
                            str(Path(row['path']).resolve()) for row in state['archives'].values()}
                            for path in inventory.values()):
                        raise RuntimeError('unindexed historical archives require identity reconciliation before adding files; existing documents retained')
                    if target.exists() and digest(target.read_bytes()) != content_hash:
                        raise RuntimeError('archive filename/content conflict; existing file retained')
                    if not target.exists():
                        atomic_write(target, content)
                        summary['archive_writes'] += 1
                        archive_changed = True
                    state['archives'][key] = {'path': str(target), 'sha256': content_hash}
                    inventory[content_hash] = target
                    save()
                if record.get('ledger_done') and ledger.exists():
                    if not archive_changed:
                        summary['already_complete'] += 1
                    continue
                with exclusive(ledger.with_suffix('.xlsx.lock')):
                    before = ledger.read_bytes() if ledger.exists() else None
                    before_rows = workbook_rows(before) if before else Counter()
                    output = backend.export_ledger(batch_id, before)
                    summary['ledger_exports'] += 1
                    after_rows = workbook_rows(output)
                    if any(after_rows[row] < count for row, count in before_rows.items()):
                        raise RuntimeError('backend output changed existing values or manual notes')
                    export_report = getattr(backend, 'last_export_report', {})
                    summary['historical_duplicate_groups'] = max(summary['historical_duplicate_groups'],
                                                                 export_report.get('historical_duplicate_groups', 0))
                    summary.setdefault('export_reports', []).append({'batch_id': batch_id, **export_report})
                    now = ledger.read_bytes() if ledger.exists() else None
                    if now != before:
                        raise RuntimeError('ledger changed during export; retry against latest file')
                    if output != before:
                        if before:
                            backup = ledger.with_name(ledger.stem + '.backup-' + digest(before)[:12] + '.xlsx')
                            if not backup.exists():
                                atomic_write(backup, before)
                        atomic_write(ledger, output)
                    state['ledger_sha256'] = digest(output)
                    for source in state['sources'].values():
                        if source.get('batch_id') == batch_id:
                            source['ledger_done'] = True
                    save()
            except Exception as error:
                summary['errors'].append({'source': path.name, 'error': str(error),
                                          'batch_id': record.get('batch_id')})
                save()
        known_paths = {str(Path(item['path']).resolve()) for item in state['archives'].values()}
        summary['unindexed_archive_files'] = sorted(path.name for path in inventory.values()
            if str(path.resolve()) not in known_paths)
        return summary


def main(expected_department=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', required=True)
    parser.add_argument('--input', nargs='*')
    parser.add_argument('--result')
    args = parser.parse_args()
    config = json.loads(Path(args.config).read_text(encoding='utf-8-sig'))
    if expected_department:
        if config.get('expected_department', expected_department) != expected_department:
            raise ValueError('department mismatch between skill entrypoint and configuration')
        config['expected_department'] = expected_department
    paths = [Path(path) for path in args.input] if args.input else sorted(
        path for path in Path(config['input_dir']).iterdir() if path.suffix.lower() in
        {'.pdf', '.png', '.jpg', '.jpeg'})
    backend = HttpBackend(config)
    started = time.monotonic()
    try:
        result = run_workflow(config, paths, backend)
    finally:
        backend.close()
    result['elapsed_seconds'] = round(time.monotonic() - started, 3)
    if args.result:
        atomic_write(Path(args.result), json.dumps(result, ensure_ascii=False, indent=2).encode('utf-8'))
    print(json.dumps(result, ensure_ascii=False, separators=(',', ':')))
    return 2 if result['errors'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
