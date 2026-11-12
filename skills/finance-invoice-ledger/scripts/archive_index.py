"""Local archive inventory and verified business identities; no recognition calls."""
from __future__ import annotations

import json
import os
import re
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

from run_invoice_workflow import atomic_write, digest


def windows_change_time(path):
    """NTFS ChangeTime changes even when an application restores LastWriteTime."""
    import ctypes
    from ctypes import wintypes
    class BasicInfo(ctypes.Structure):
        _fields_ = [(name, ctypes.c_longlong) for name in
                    ('CreationTime', 'LastAccessTime', 'LastWriteTime', 'ChangeTime')] + [
                    ('FileAttributes', wintypes.DWORD)]
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    create = kernel.CreateFileW
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p,
                       wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    create.restype = wintypes.HANDLE
    get_info = kernel.GetFileInformationByHandleEx
    get_info.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD]
    close = kernel.CloseHandle; close.argtypes = [wintypes.HANDLE]
    handle = create(str(path), 0x80, 7, None, 3, 0, None)
    if handle == wintypes.HANDLE(-1).value: return None
    try:
        info = BasicInfo()
        return info.ChangeTime if get_info(handle, 0, ctypes.byref(info), ctypes.sizeof(info)) else None
    finally: close(handle)

ARCHIVE_SUFFIXES = {'.pdf', '.png', '.jpg', '.jpeg', '.ofd', '.xml', '.tif', '.tiff', '.bmp', '.webp'}
FIELDS = ('invoice_type', 'invoice_code', 'invoice_number_full', 'issue_date', 'buyer_name',
          'buyer_tax_id', 'seller_name', 'seller_tax_id', 'amount_ex_tax', 'tax_amount',
          'total_amount', 'currency_code', 'is_red_letter')


class ArchiveConflict(RuntimeError):
    """A local archive cannot safely be reused or supplemented."""


def compact(value):
    return re.sub(r'\s+', '', str(value or ''))


def identities(values):
    result = {}
    for raw in values:
        if not isinstance(raw, dict) or any(key not in raw for key in FIELDS):
            raise ValueError('系统归档身份元数据不完整')
        scheme = raw.get('identity_scheme'); kind = compact(raw['invoice_type']).lower()
        code = compact(raw['invoice_code']).upper(); number = compact(raw['invoice_number_full']).upper()
        if scheme != 'cn_vat_v1' or kind not in {'vat_ordinary', 'vat_special', 'railway_ticket'} or not number:
            raise ValueError('系统归档身份规则尚不支持，不能自行推断')
        key = digest(f'{scheme}|{kind}|{code}|{number}'.encode())
        if raw.get('identity_key') != key: raise ValueError('系统归档身份校验不一致')
        row = {key: raw[key] for key in FIELDS}
        row.update(identity_scheme=scheme, identity_key=key, invoice_type=kind,
                   invoice_code=code, invoice_number_full=number)
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', str(row['issue_date'])):
            raise ValueError('系统归档身份日期不完整')
        for field in ('amount_ex_tax', 'tax_amount', 'total_amount'):
            try:
                value = Decimal(str(row[field]))
                if not value.is_finite(): raise ValueError()
            except (InvalidOperation, ValueError):
                raise ValueError('系统归档身份金额不完整') from None
            row[field] = str(value.normalize())
        for field in ('buyer_name', 'seller_name', 'buyer_tax_id', 'seller_tax_id', 'currency_code'):
            row[field] = compact(row[field])
        if not row['buyer_name'] or not row['seller_name'] or not row['currency_code']:
            raise ValueError('系统归档身份缺少交易双方或币种')
        if not isinstance(row['is_red_letter'], bool): raise ValueError('系统归档身份红字状态不完整')
        if key in result and result[key] != row: raise ArchiveConflict('同一发票身份的关键字段冲突')
        result[key] = row
    if not result: raise ValueError('缺少已核验的归档发票身份')
    return result


class ArchiveIndex:
    def __init__(self, root, path, scope):
        self.root = Path(root).resolve(); self.path = Path(path)
        self.data = json.loads(self.path.read_text(encoding='utf-8-sig')) if self.path.exists() else {
            'version': 1, 'scope': scope, 'files': {}, 'aliases': {}}
        if self.data.get('version') != 1 or self.data.get('scope') != scope:
            raise ValueError('归档索引范围或账户不一致，保留原索引')
        self.files = self.data['files']
        self.statistics = {'hashes_read': 0, 'cached_files': 0}
        for relative in self.files: self.file_path(relative)

    def file_path(self, relative):
        value = Path(relative)
        target = (self.root/value).resolve()
        if value.is_absolute() or target == self.root or not target.is_relative_to(self.root):
            raise ValueError('归档索引路径超出指定目录范围')
        return target

    def relative(self, path):
        path = Path(path).resolve()
        if not path.is_relative_to(self.root) or path == self.root:
            raise ValueError('归档文件不在指定目录范围')
        return path.relative_to(self.root).as_posix()

    @staticmethod
    def stamp(path):
        stat = path.stat()
        return {'size': stat.st_size, 'mtime_ns': stat.st_mtime_ns,
                'ctime_ns': stat.st_ctime_ns, 'file_id': stat.st_ino,
                'change_time': windows_change_time(path) if os.name == 'nt' else stat.st_ctime_ns}

    def scan(self):
        prior = self.files; by_hash = defaultdict(list)
        for record in prior.values(): by_hash[record['sha256']].append(record)
        current = {}; self.statistics = {'hashes_read': 0, 'cached_files': 0}
        for path in sorted(self.root.rglob('*')):
            if not path.is_file() or path.suffix.lower() not in ARCHIVE_SUFFIXES: continue
            relative = self.relative(path); stamp = self.stamp(path); old = prior.get(relative)
            if old and stamp['change_time'] is not None and all(old.get(key) == value for key, value in stamp.items()):
                row = dict(old); self.statistics['cached_files'] += 1
            else:
                before = stamp; sha = digest(path.read_bytes()); self.statistics['hashes_read'] += 1
                if self.stamp(path) != before: raise ArchiveConflict('归档原件在扫描期间发生变化，请重试')
                matching = [r for r in by_hash.get(sha, []) if r['status'] == 'verified']
                row = dict(matching[0]) if matching else {'status': 'unverified', 'identities': {}}
                row.update(sha256=sha, **stamp)
                if old and old['sha256'] != sha and old['status'] == 'verified':
                    row.update(status='modified', previous_sha256=old['sha256'],
                               previous_identities=old['identities'])
            current[relative] = row
        self.files = self.data['files'] = current
        verified_by_hash = {row['sha256']: row for row in current.values() if row['status'] == 'verified'}
        for name, row in list(current.items()):
            evidence = verified_by_hash.get(row['sha256'])
            if row['status'] != 'unverified' or not evidence: continue
            self.verify_file(name, row)
            current[name] = {**evidence, **self.stamp(self.file_path(name))}
        return {row['sha256']: self.file_path(name) for name, row in current.items()}

    def bind(self, path, expected_sha, metadata, batch_id):
        relative = self.relative(path); path = self.file_path(relative)
        verified = identities(metadata)
        if not path.is_file() or digest(path.read_bytes()) != expected_sha:
            raise ArchiveConflict('归档原件已变化，不能登记为已核验')
        current = self.files.get(relative, {})
        if current.get('status') == 'modified':
            raise ArchiveConflict('已有归档被修改，需要先人工核对，不能自动覆盖核验记录')
        if current.get('status') == 'verified' and current['identities'] != verified:
            raise ArchiveConflict('归档原件的完整发票身份集合冲突')
        for name, other in self.files.items():
            if name == relative or other['status'] != 'verified': continue
            if set(verified) & set(other['identities']) and verified != other['identities']:
                raise ArchiveConflict('历史原件完整发票集合部分重复或关键字段冲突，保留原件待核对')
        self.files[relative] = {'sha256': expected_sha, 'status': 'verified', 'identities': verified,
                                'batch_id': batch_id, **self.stamp(path)}
        # One validated byte sequence proves every unchanged exact copy, without
        # asking the recognition service to inspect the same original again.
        for name, row in list(self.files.items()):
            if name == relative or row['sha256'] != expected_sha or row['status'] != 'unverified': continue
            self.verify_file(name, row)
            self.files[name] = {**self.files[relative], **self.stamp(self.file_path(name))}

    def verify_file(self, name, row):
        path = self.file_path(name)
        if not path.is_file() or digest(path.read_bytes()) != row['sha256']:
            raise ArchiveConflict('已有归档原件缺失或内容被修改，请刷新索引后核对')
        return path

    def choose(self, metadata, expected_sha):
        incoming = identities(metadata); keys = set(incoming); matches = []
        for name, row in self.files.items():
            if row['status'] != 'verified': continue
            old = row['identities']; overlap = keys & set(old)
            if not overlap: continue
            if any(incoming[key] != old[key] for key in overlap):
                raise ArchiveConflict('同一发票身份的关键字段冲突，保留原件并暂缓归档')
            if keys != set(old):
                raise ArchiveConflict('原件发票集合存在部分重复，不能跳过整份或重复归档，请拆分后处理')
            matches.append((name, row))
        if matches:
            for name, row in matches: self.verify_file(name, row)
            name, row = matches[0]
            return {'action': 'reuse', 'path': self.file_path(name), 'sha256': row['sha256']}
        if any(row['status'] != 'verified' for row in self.files.values()):
            raise ArchiveConflict('历史归档尚有未核验或被修改的文件，请先分批建立历史身份索引')
        return {'action': 'new', 'sha256': expected_sha}

    def record_alias(self, source_sha, metadata, canonical_path, batch_id=None):
        relative = self.relative(canonical_path); row = self.files.get(relative)
        verified = identities(metadata)
        if not row or row['status'] != 'verified' or row['identities'] != verified:
            raise ArchiveConflict('来源与已核验原件的完整发票身份集合不匹配')
        self.data['aliases'][source_sha] = {'archive_sha256': row['sha256'],
            'identities': verified, 'batch_id': batch_id or row.get('batch_id')}

    def source(self, sha):
        alias = self.data['aliases'].get(sha)
        for name, row in self.files.items():
            if row['status'] != 'verified': continue
            if row['sha256'] == sha or (alias and alias['archive_sha256'] == row['sha256']
                                       and alias['identities'] == row['identities']):
                return {'path': self.verify_file(name, row), 'sha256': row['sha256'],
                        'identities': row['identities'],
                        'batch_id': (alias or {}).get('batch_id') or row.get('batch_id')}
        return None

    def report(self):
        counts = Counter(key for row in self.files.values() if row['status'] == 'verified' for key in row['identities'])
        return {'total_files': len(self.files),
                'verified_files': sum(r['status'] == 'verified' for r in self.files.values()),
                'unverified_files': sum(r['status'] != 'verified' for r in self.files.values()),
                'duplicate_identity_groups': sum(n > 1 for n in counts.values()),
                **self.statistics}

    def save(self):
        atomic_write(self.path, json.dumps(self.data, ensure_ascii=False, indent=2).encode('utf-8'))
