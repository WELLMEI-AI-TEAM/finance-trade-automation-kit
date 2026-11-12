import importlib.util
import json
import tempfile
import unittest
import zipfile
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / 'skills/pi-to-export-contract/scripts/apply_ooxml_plan.py'


class ContractPackageTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / 'PI.xlsx'
        self.order = self.root / 'order.txt'
        self.purchase = self.root / 'purchase.txt'
        self.order.write_text('WMDS1 -> ABC1 -> PO1')
        self.purchase.write_text('WMDS1 signing 2026-05-04')
        self.parts = {
            'xl/workbook.xml': b'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="PI" sheetId="1" r:id="rId1"/></sheets></workbook>',
            'xl/_rels/workbook.xml.rels': b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/></Relationships>',
            'xl/sharedStrings.xml': b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>Proforma Invoice</t></si><si><t>PI No ABC1</t></si></sst>',
            'xl/worksheets/sheet1.xml': b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1" ht="26.25"><c r="A1" s="3" t="s"><v>0</v></c></row><row r="2"><c r="A2" s="1" t="s"><v>1</v></c></row><row r="3"><c r="D3" s="7"><v>46147</v></c></row><row r="5"><c r="D5"><f>12*34.5</f><v>414</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>',
            'xl/styles.xml': b'original-style-bytes',
            'xl/media/image1.png': b'original-image-bytes',
            'xl/vbaProject.bin': b'original-macro-bytes',
        }
        self.write_source()
        self.module = None

    def tearDown(self):
        self.temp.cleanup()

    def write_source(self):
        with zipfile.ZipFile(self.source, 'w', zipfile.ZIP_DEFLATED) as z:
            for name, data in self.parts.items():
                z.writestr(name, data)

    def load_module(self):
        self.assertTrue(SCRIPT.exists(), 'Exact OOXML contract editor is not implemented')
        spec = importlib.util.spec_from_file_location('ooxml_contract', SCRIPT)
        self.module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.module)

    def plan(self):
        import hashlib
        return {
            'source': str(self.source), 'output': str(self.root/'CONTRACT.xlsx'),
            'source_sha256': hashlib.sha256(self.source.read_bytes()).hexdigest(),
            'association': {'status': 'confirmed',
                'order_confirmation': {'path': str(self.order), 'locator': 'line1', 'evidence': 'WMDS1 -> ABC1 -> PO1'},
                'purchase_contract': {'path': str(self.purchase), 'locator': 'line1', 'evidence': 'WMDS1 signing date', 'signing_date': '2026-05-04'}},
            'title': {'sheet': 'PI', 'cell': 'A1', 'expected': 'Proforma Invoice'},
            'date': {'sheet': 'PI', 'cell': 'D3', 'expected': 46147, 'pi_date': '2026-05-05'}
        }

    def test_exact_preservation_and_later_date(self):
        self.load_module()
        before = self.source.read_bytes()
        audit = self.module.apply_plan(self.plan())
        self.assertEqual(self.source.read_bytes(), before)
        self.assertTrue(audit['all_other_package_parts_identical'])
        with zipfile.ZipFile(self.root/'CONTRACT.xlsx') as z:
            for name in self.parts:
                if name not in audit['modified_parts']:
                    self.assertEqual(z.read(name), self.parts[name], name)
            sheet = z.read('xl/worksheets/sheet1.xml')
            self.assertIn(b'<v>46146</v>', sheet)
            self.assertIn(b'ht="26.25"', sheet)
            self.assertIn(b'<f>12*34.5</f><v>414</v>', sheet)
        self.assertEqual(audit['date_result'], '2026-05-04')

    def test_earlier_equal_and_no_title(self):
        self.load_module()
        import re
        original_sheet = re.sub(rb'<c r="A1"[^>]*>.*?</c>', b'', self.parts['xl/worksheets/sheet1.xml'])
        for serial, day in [(46145, '2026-05-03'), (46146, '2026-05-04')]:
            self.parts['xl/worksheets/sheet1.xml'] = original_sheet.replace(b'<v>46147</v>',f'<v>{serial}</v>'.encode())
            self.write_source()
            plan=self.plan();plan['title']=None;plan['date'].update(expected=serial,pi_date=day)
            plan['output']=str(self.root/f'{day}.xlsx')
            audit=self.module.apply_plan(plan)
            self.assertFalse(audit['date_changed'])
            self.assertEqual(Path(plan['output']).read_bytes(),self.source.read_bytes())

    def test_text_date_pattern(self):
        self.load_module()
        self.parts['xl/worksheets/sheet1.xml']=self.parts['xl/worksheets/sheet1.xml'].replace(b'<c r="D3" s="7"><v>46147</v></c>',b'<c r="D3" s="7" t="inlineStr"><is><t>Date: May 05, 2026</t></is></c>')
        self.write_source();plan=self.plan()
        plan['date'].update(expected='Date: May 05, 2026',format='MMMM dd, yyyy',prefix='Date: ',suffix='',text_case='none')
        audit=self.module.apply_plan(plan)
        with zipfile.ZipFile(plan['output']) as z:self.assertIn(b'Date: May 04, 2026',z.read('xl/worksheets/sheet1.xml'))
        self.assertEqual(audit['date_result'],'2026-05-04')

    def test_1904_dates_and_time_fraction(self):
        self.load_module()
        self.parts['xl/workbook.xml']=self.parts['xl/workbook.xml'].replace(b'<sheets>',b'<workbookPr date1904="1"/><sheets>')
        self.parts['xl/worksheets/sheet1.xml']=self.parts['xl/worksheets/sheet1.xml'].replace(b'<v>46147</v>',b'<v>44685.5</v>')
        self.write_source();plan=self.plan();plan['date']['expected']=44685.5
        self.module.apply_plan(plan)
        with zipfile.ZipFile(plan['output']) as z:self.assertIn(b'<v>44684.5</v>',z.read('xl/worksheets/sheet1.xml'))

    def test_shared_title_does_not_change_other_occurrence(self):
        self.load_module()
        self.parts['xl/worksheets/sheet1.xml']=self.parts['xl/worksheets/sheet1.xml'].replace(b'</sheetData>',b'<row r="8"><c r="A8" t="s"><v>0</v></c></row></sheetData>')
        self.write_source();plan=self.plan();self.module.apply_plan(plan)
        with zipfile.ZipFile(plan['output']) as z:
            self.assertEqual(z.read('xl/sharedStrings.xml'),self.parts['xl/sharedStrings.xml'])
            self.assertIn(b'<c r="A8" t="s"><v>0</v></c>',z.read('xl/worksheets/sheet1.xml'))

    def test_invoice_label_preserves_number_and_other_occurrences(self):
        self.load_module()
        extra = b'<row r="9"><c r="V9" s="3" t="inlineStr"><is><t>  Invoice No.: ABC1  </t></is></c><c r="W9" t="inlineStr"><is><t>Invoice No.</t></is></c></row>'
        self.parts['xl/worksheets/sheet1.xml'] = self.parts['xl/worksheets/sheet1.xml'].replace(b'</sheetData>', extra+b'</sheetData>')
        self.write_source()
        plan = self.plan()
        plan['invoice_label'] = {'sheet': 'PI', 'cell': 'V9', 'expected': '  Invoice No.: ABC1  '}
        audit = self.module.apply_plan(plan)
        with zipfile.ZipFile(plan['output']) as archive:
            sheet = archive.read('xl/worksheets/sheet1.xml')
            self.assertIn(b'  No.: ABC1  ', sheet)
            self.assertIn(b'<c r="W9" t="inlineStr"><is><t>Invoice No.</t></is></c>', sheet)
            self.assertIn(b'<c r="A2" s="1" t="s"><v>1</v></c>', sheet)
        self.assertIn('invoice_label', [change['kind'] for change in audit['changes']])

    def test_english_ordinal_date_changes_suffix_with_day(self):
        self.load_module()
        self.parts['xl/worksheets/sheet1.xml'] = self.parts['xl/worksheets/sheet1.xml'].replace(
            b'<c r="D3" s="7"><v>46147</v></c>',
            b'<c r="D3" s="7" t="inlineStr"><is><t>Date: May 1st, 2026</t></is></c>')
        self.write_source()
        plan = self.plan()
        plan['date'].update(expected='Date: May 1st, 2026', pi_date='2026-05-01',
                            format='MMMM d{ordinal}, yyyy', prefix='Date: ')
        plan['association']['purchase_contract']['signing_date'] = '2026-04-30'
        self.module.apply_plan(plan)
        with zipfile.ZipFile(plan['output']) as archive:
            self.assertIn(b'Date: April 30th, 2026', archive.read('xl/worksheets/sheet1.xml'))

    def test_invoice_label_rejects_wrong_field_and_colliding_targets(self):
        self.load_module()
        plan = self.plan()
        plan['invoice_label'] = {'sheet': 'PI', 'cell': 'A2', 'expected': 'PI No ABC1'}
        with self.assertRaises(ValueError):
            self.module.apply_plan(plan)
        plan['invoice_label'] = dict(plan['title'])
        with self.assertRaises(ValueError):
            self.module.apply_plan(plan)

    def test_refuse_conflict_formula_and_overwrite(self):
        self.load_module()
        plan=self.plan();plan['association']['status']='pending'
        with self.assertRaises(ValueError):self.module.apply_plan(plan)
        plan=self.plan();plan['output']=plan['source']
        with self.assertRaises(ValueError):self.module.apply_plan(plan)
        self.parts['xl/worksheets/sheet1.xml']=self.parts['xl/worksheets/sheet1.xml'].replace(b'<v>46147</v>',b'<f>DATE(2026,5,5)</f><v>46147</v>')
        self.write_source()
        with self.assertRaises(ValueError):self.module.apply_plan(self.plan())


if __name__ == '__main__':
    unittest.main()
