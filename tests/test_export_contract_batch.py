import importlib.util
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPT = Path(__file__).resolve().parents[1] / 'skills/pi-to-export-contract/scripts/prepare_batch.py'


class BatchTests(unittest.TestCase):
    def setUp(self):
        self.assertTrue(SCRIPT.is_file(), '批次准备与全量备份功能尚未实现')
        spec = importlib.util.spec_from_file_location('contract_batch', SCRIPT)
        self.module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.module)
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / '出关合同'
        self.root.mkdir()

    def test_default_pending_backup_and_layout(self):
        pending = self.root / '待处理'
        pending.mkdir()
        (pending / 'PI.xlsx').write_bytes(b'original PI')
        (pending / '合同.zip').write_bytes(b'original archive')
        result = self.module.prepare_batch(self.root, stamp='2026-10-28_102030')
        output = Path(result['result_dir'])
        backup = Path(result['backup_dir'])
        self.assertEqual(output.name, backup.name)
        self.assertTrue(all((output / name).is_dir() for name in ['出口合同', 'PI', '购货合同']))
        self.assertEqual(len(result['files']), 2)
        for record in result['files']:
            self.assertEqual(Path(record['source']).read_bytes(), Path(record['backup']).read_bytes())
        self.assertEqual(json.loads((backup / '备份清单.json').read_text(encoding='utf-8'))['backup_status'], 'verified')

    def test_root_scan_excludes_history_and_repeated_batches(self):
        (self.root / 'PI.xls').write_bytes(b'PI')
        for name in ['备份', '处理结果', '待核对']:
            (self.root / name).mkdir()
            (self.root / name / 'old.xlsx').write_bytes(b'old')
        a = self.module.prepare_batch(self.root, [self.root], stamp='2026-10-28_102030')
        b = self.module.prepare_batch(self.root, [self.root], stamp='2026-10-28_102030')
        self.assertEqual(len(a['files']), 1)
        self.assertEqual(len(b['files']), 1)
        self.assertNotEqual(a['batch'], b['batch'])
        self.assertEqual(Path(a['files'][0]['backup']).read_bytes(), b'PI')
        self.assertEqual(Path(a['pending_dir']), self.root / '待核对' / a['batch'])

    def test_explicit_pending_retry_is_allowed(self):
        pending = self.root / '待核对/old'
        pending.mkdir(parents=True)
        (pending / 'PI.xlsx').write_bytes(b'pending original')
        result = self.module.prepare_batch(self.root, [self.root / '待核对'], stamp='2026-10-28_102030')
        self.assertEqual(len(result['files']), 1)
        self.assertEqual(Path(result['files'][0]['backup']).read_bytes(), b'pending original')

    def test_attachments_duplicate_names_and_overlapping_inputs(self):
        incoming = Path(self.temp.name) / 'attachments'
        for name in ['a', 'b']:
            (incoming / name).mkdir(parents=True)
            (incoming / name / 'PI.xlsx').write_bytes(name.encode())
        result = self.module.prepare_batch(self.root, [incoming, incoming / 'a/PI.xlsx'], stamp='2026-10-28_102030')
        self.assertEqual(len(result['files']), 2)
        self.assertEqual({Path(x['backup']).read_bytes() for x in result['files']}, {b'a', b'b'})

    def test_empty_or_missing_inputs_do_not_create_batch(self):
        (self.root / '待处理').mkdir()
        with self.assertRaises(ValueError):
            self.module.prepare_batch(self.root)
        self.assertFalse((self.root / '处理结果').exists())
        with self.assertRaises(FileNotFoundError):
            self.module.prepare_batch(self.root, [self.root / 'missing.pdf'])

    def test_explicit_history_source_is_rejected(self):
        history = self.root / '处理结果'
        history.mkdir()
        previous = history / 'old.xlsx'
        previous.write_bytes(b'old')
        with self.assertRaises(ValueError):
            self.module.prepare_batch(self.root, [previous])

    def test_backup_failure_never_reports_ready(self):
        source = self.root / 'PI.xlsx'
        source.write_bytes(b'original')
        real_hash = self.module.sha256

        def concurrent_update(path):
            if Path(path) == source:
                source.write_bytes(b'changed during backup')
            return real_hash(path)

        with patch.object(self.module, 'sha256', concurrent_update):
            with self.assertRaises(ValueError):
                self.module.prepare_batch(self.root, [source], stamp='2026-10-28_102030')
        backup = self.root / '备份/2026-10-28_102030'
        output = self.root / '处理结果/2026-10-28_102030'
        self.assertFalse((output / '批次记录.json').exists())
        self.assertFalse((backup / '备份清单.json').exists())
        failed = json.loads((backup / '备份失败.json').read_text(encoding='utf-8'))
        self.assertEqual(failed['backup_status'], 'failed')
        self.assertFalse(list((output / '出口合同').iterdir()))


if __name__ == '__main__':
    unittest.main()
