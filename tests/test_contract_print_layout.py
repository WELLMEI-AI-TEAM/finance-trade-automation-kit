import importlib.util
from pathlib import Path
import sys

import fitz
import pytest
from openpyxl import Workbook

SCRIPT = Path(__file__).resolve().parents[1] / 'skills/contract-print-layout/scripts/print_contracts.py'
spec = importlib.util.spec_from_file_location('contract_print_layout', SCRIPT)
layout = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = layout
spec.loader.exec_module(layout)


def pdf(path, texts, size=10):
    doc = fitz.open()
    for text in texts:
        page = doc.new_page(width=595, height=842)
        if text:
            page.insert_text((40, 60), text, fontsize=size)
    doc.save(path)
    return path


def test_pdf_rejects_tiny_text_and_empty_pages(tmp_path):
    small = layout.pdf_metrics(pdf(tmp_path/'small.pdf', ['CONTRACT CP26015 amount 123.45'], 5))
    assert not layout.readable(small)
    blank = layout.pdf_metrics(pdf(tmp_path/'blank.pdf', ['CONTRACT amount 123.45', '']))
    assert not layout.readable(blank)


def test_pdf_detects_missing_invoice_numbers(tmp_path):
    source = pdf(tmp_path/'source.pdf', ['CONTRACT CP26015 amount 123.45'])
    target = pdf(tmp_path/'target.pdf', ['CONTRACT amount 123.45'])
    assert layout.missing_tokens(layout.pdf_metrics(source), layout.pdf_metrics(target)) == ['CP26015']


def test_best_candidate_prefers_readable_fewer_pages():
    candidates = [
        {'id':'one', 'eligible':False, 'metrics':{'pages':1,'median_font':5}},
        {'id':'three', 'eligible':True, 'metrics':{'pages':3,'median_font':10}},
        {'id':'two', 'eligible':True, 'metrics':{'pages':2,'median_font':8}},
    ]
    assert layout.choose_candidate(candidates)['id'] == 'two'
    assert layout.choose_candidate(candidates[:1]) is None


def test_fingerprint_ignores_print_changes_but_detects_content_and_font(tmp_path):
    path = tmp_path/'source.xlsx'
    book = Workbook(); sheet = book.active
    sheet['A1'] = 'CONTRACT'; sheet['A2'] = '=1+1'; sheet.merge_cells('A4:B5')
    book.save(path)
    baseline = layout.fingerprint(path)
    sheet.page_setup.orientation = 'landscape'; sheet.page_setup.scale=70
    sheet.print_area='A1:C7'; book.save(path)
    assert baseline == layout.fingerprint(path)
    sheet['A1']='CHANGED'; book.save(path)
    assert baseline != layout.fingerprint(path)
    sheet['A1']='CONTRACT'
    from copy import copy
    font=copy(sheet['A1'].font); font.sz=5; sheet['A1'].font=font; book.save(path)
    assert baseline != layout.fingerprint(path)


def test_inputs_are_explicit_and_deduplicated(tmp_path):
    a=tmp_path/'a.xlsx'; a.write_bytes(b'same')
    b=tmp_path/'b.xlsx'; b.write_bytes(b'same')
    (tmp_path/'ignore.txt').write_text('ignore')
    files=layout.collect_inputs([str(tmp_path)])
    assert len(files)==1
    with pytest.raises(ValueError): layout.collect_inputs([str(tmp_path/'missing')])


def test_finish_refuses_unreviewed_or_modified_drafts(tmp_path):
    import json
    work=tmp_path/'work';work.mkdir()
    draft=work/'draft.xlsx'; draft.write_bytes(b'original')
    target=pdf(work/'draft.pdf', ['CONTRACT CP26015 amount 123.45'])
    manifest=work/'batch.json'
    item={'id':'a','status':'awaiting_visual_review','excel':str(draft),'pdf':str(target),
          'excel_sha256':layout.sha256(draft),'pdf_sha256':layout.sha256(target),'name':'a.xlsx'}
    manifest.write_text(json.dumps({'work_dir':str(work),'output_dir':str(tmp_path/'results'),'items':[item]}))
    with pytest.raises(ValueError): layout.finish(manifest, {'a':{'status':'passed','pages':[]}})
    draft.write_bytes(b'changed')
    with pytest.raises(ValueError): layout.finish(manifest, {'a':{'status':'passed','pages':[1],'notes':'readable'}})


def test_description_stays_with_first_priced_line():
    values={(1,2):'Description',(1,3):'Quantity',(1,5):'Amount',
            (3,2):'Pillow',(4,2):'Fabric: cotton',(6,2):'King',(6,3):48,(6,5):90,
            (7,2):'Queen',(7,3):24,(7,5):80,
            (9,2):'Blanket',(10,2):'Fabric: wool',(12,3):10,(12,5):20}
    assert layout.description_blocks(values)==[[3,6],[9,12]]


def test_bank_information_starts_footer_at_numeric_total():
    values={(50,3):30,(50,5):90,(52,2):'Bank information:',(53,2):'SWIFT CODE: ABC'}
    assert layout.footer_start(values)==50
    assert layout.footer_start({(10,2):'Pillow'})==0
    assert layout.footer_start({(10,2):'Pillow',(10,3):3,(10,5):30,(12,3):3,(12,5):30})==10


def test_dimension_continuation_stays_with_its_priced_line():
    values={(1,2):'Description',(1,3):'Quantity',(1,5):'Amount',
            (3,2):'Queen Duvet',(3,3):10,(3,5):30,(4,2):'210*210cm',
            (5,2):'King Duvet',(5,3):12,(5,5):40}
    assert layout.description_blocks(values)==[[3,5]]


def test_print_transplant_keeps_original_package_and_namespace(tmp_path):
    import zipfile
    from lxml import etree as ET
    from print_patch import patch_print
    source=tmp_path/'source.xlsx';native=tmp_path/'native.xlsx';out=tmp_path/'out.xlsx'
    book=Workbook();s=book.active;s['A1']='CONTRACT';s['B2']='=5+6';s.print_area='A1:C8';book.save(source)
    s['A1']='Native accidentally changed content';s.page_setup.orientation='landscape';s.page_setup.scale=65;book.save(native)
    with zipfile.ZipFile(source,'a') as z:z.writestr('customXml/business.xml',b'<original/>')
    patch_print(source,native,out)
    assert layout.fingerprint(source)==layout.fingerprint(out)
    with zipfile.ZipFile(source) as a,zipfile.ZipFile(out) as b:
        assert set(a.namelist())==set(b.namelist())
        for n in a.namelist():
            if n not in ('xl/workbook.xml','xl/worksheets/sheet1.xml'):assert a.read(n)==b.read(n)
        assert b'landscape' in b.read('xl/worksheets/sheet1.xml')
    from openpyxl import load_workbook
    result=load_workbook(out);assert result.active['A1'].value=='CONTRACT';assert result.active.page_setup.scale==65


def test_output_collisions_are_preflighted(tmp_path):
    import json
    work=tmp_path/'work';work.mkdir();out=tmp_path/'results';(out/'PDF打印版').mkdir(parents=True)
    draft=work/'a.xlsx';draft.write_bytes(b'keep')
    target=pdf(work/'a.pdf',['CONTRACT amount 10'])
    items=[]
    for k in ('1','2'):
        items.append({'id':k,'status':'awaiting_visual_review','excel':str(draft),'pdf':str(target),
          'excel_sha256':layout.sha256(draft),'pdf_sha256':layout.sha256(target),'name':k+'.xlsx'})
    (out/'PDF打印版/2.pdf').write_bytes(b'existing')
    manifest=work/'batch.json';manifest.write_text(json.dumps({'work_dir':str(work),'output_dir':str(out),'items':items}))
    with pytest.raises(ValueError):layout.finish(manifest,{k:{'status':'passed','pages':[1],'notes':'Observed complete'} for k in ('1','2')})
    assert not (out/'Excel打印版').exists()
    assert list((out/'PDF打印版').iterdir())==[out/'PDF打印版/2.pdf']
