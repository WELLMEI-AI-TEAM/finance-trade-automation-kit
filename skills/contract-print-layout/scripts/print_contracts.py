"""Create verified Excel/PDF printing drafts using native Excel or WPS.

Never sends a physical print job. All mutations are confined to isolated copies.
"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import struct
import subprocess
import sys
import unicodedata
import uuid

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0,str(SCRIPT_DIR))
from print_patch import patch_print
SUFFIXES = {'.xls', '.xlsx', '.xlsm'}


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def save_json(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2, default=str), encoding='utf-8')


def collect_inputs(inputs):
    result = []; seen = set()
    for raw in inputs:
        root = Path(raw).resolve()
        if not root.exists():
            raise ValueError('Input does not exist: ' + str(root))
        paths = sorted(root.rglob('*')) if root.is_dir() else [root]
        for path in paths:
            if not path.is_file() or path.is_symlink() or path.suffix.lower() not in SUFFIXES or path.name.startswith('~$'):
                continue
            if root.is_dir() and any(p in {'处理结果','备份','待核对','Excel打印版','PDF打印版'} for p in path.relative_to(root).parts[:-1]):
                continue
            key = sha256(path)
            if key not in seen:
                seen.add(key); result.append(path)
    return result


def _style_xml(value):
    from xml.etree.ElementTree import tostring
    return tostring(value.to_tree(), encoding='unicode') if value is not None else None


def fingerprint(path):
    """Logical cell/style/layout fingerprint; print parameters intentionally excluded."""
    path = Path(path)
    if path.suffix.lower() == '.xls':
        import xlrd
        book = xlrd.open_workbook(path, formatting_info=True, on_demand=True)
        def font(index):
            f=book.font_list[index]
            return {k:getattr(f,k) for k in ('name','height','colour_index','weight','escapement','underline_type','family','character_set','bold','italic','underlined','struck_out','outline','shadow')}
        def style(index):
            xf=book.xf_list[index]
            return {'font':font(xf.font_index),'format':book.format_map[xf.format_key].format_str,
                    'align':vars(xf.alignment),'border':vars(xf.border),'fill':vars(xf.background),'protection':vars(xf.protection)}
        data={'date_system':book.datemode,'palette':book.colour_map,'sheets':[]}
        try:
            for s in book.sheets():
                cells=[]
                for row in range(s.nrows):
                    for col in range(s.ncols):
                        c=s.cell(row,col)
                        if c.ctype == 0: continue
                        cells.append((row,col,c.ctype,c.value,style(c.xf_index)))
                data['sheets'].append({'name':s.name,'visibility':s.visibility,'cells':cells,'merges':sorted(s.merged_cells),
                    'rows':{r:{k:getattr(v,k,None) for k in ('height','hidden','outline_level','height_mismatch')} for r,v in s.rowinfo_map.items()},
                    'cols':{c:{k:getattr(v,k,None) for k in ('width','hidden','outline_level','collapsed')} for c,v in s.colinfo_map.items()},
                    'rich':{str(k):[(off,font(idx)) for off,idx in v] for k,v in s.rich_text_runlist_map.items()}})
            # Independent formula tokens and embedded drawing bytes.
            records=[];pos=book.base;end=book.base+book.stream_len;continuation=False
            while pos+4<=end:
                code,size=struct.unpack_from('<HH',book.mem,pos);pos+=4
                raw=bytes(book.mem[pos:pos+size]);pos+=size
                if code==6 and len(raw)>=22:
                    n=struct.unpack_from('<H',raw,20)[0];records.append((code,raw[:4].hex(),raw[22:22+n].hex()))
                if code in (0xEB,0xEC) or (code==0x3C and continuation):
                    records.append((code,raw.hex()));continuation=True
                else: continuation=False
            data['formula_drawings']=records
        finally: book.release_resources()
    else:
        from openpyxl import load_workbook
        book=load_workbook(path,data_only=False,keep_links=True,keep_vba=path.suffix.lower()=='.xlsm')
        data={'date_system':str(book.epoch),'sheets':[]}
        try:
            for s in book:
                cells=[]
                for row in s:
                    for c in row:
                        if c.value is None and not c.has_style:continue
                        cells.append((c.coordinate,c.data_type,str(c.value),c.number_format,
                                      *[_style_xml(x) for x in (c.font,c.fill,c.border,c.alignment,c.protection)]))
                data['sheets'].append({'name':s.title,'visibility':s.sheet_state,'cells':cells,
                    'merges':sorted(str(x) for x in s.merged_cells.ranges),
                    'rows':{k:{a:getattr(v,a) for a in ('height','hidden','outlineLevel','collapsed')} for k,v in s.row_dimensions.items()},
                    'cols':{k:{a:getattr(v,a) for a in ('width','hidden','outlineLevel','collapsed','min','max')} for k,v in s.column_dimensions.items()},
                    'images':[hashlib.sha256(im._data()).hexdigest() for im in s._images]})
            import zipfile
            with zipfile.ZipFile(path) as archive:
                data['objects']={n:hashlib.sha256(archive.read(n)).hexdigest() for n in archive.namelist()
                                 if n.startswith(('xl/media/','xl/embeddings/')) or n.endswith('vbaProject.bin')}
        finally:book.close()
    return hashlib.sha256(json.dumps(data,sort_keys=True,ensure_ascii=False,default=str).encode('utf-8')).hexdigest()


def _tokens(text):
    return re.findall(r'[A-Za-z0-9]+(?:[.,/-][A-Za-z0-9]+)*|[\u4e00-\u9fff]', unicodedata.normalize('NFKC',text))


def pdf_metrics(path):
    import fitz
    sizes=[];texts=[];blank=[];page_sizes=[]
    with fitz.open(path) as doc:
        for i,page in enumerate(doc):
            text=page.get_text();texts.append(text);page_sizes.append([round(page.rect.width,1),round(page.rect.height,1)])
            if not text.strip() and not page.get_images() and not page.get_drawings():blank.append(i+1)
            for block in page.get_text('dict')['blocks']:
                for line in block.get('lines',[]):
                    for span in line['spans']:
                        sizes.extend([round(span['size'],3)]*len(span['text'].strip()))
    sizes.sort();counts=Counter(_tokens('\n'.join(texts)))
    return {'pages':len(texts),'page_sizes':page_sizes,'blank_pages':blank,
            'median_font':sizes[len(sizes)//2] if sizes else 0,
            'p10_font':sizes[int(len(sizes)*.1)] if sizes else 0,
            'tokens':sorted(counts),'token_counts':dict(counts), 'text_chars':sum(map(len,texts))}


def readable(metrics, min_median=8, min_p10=7):
    return bool(metrics['pages'] and not metrics['blank_pages'] and metrics['median_font']>=min_median and metrics['p10_font']>=min_p10)


def missing_tokens(before, after):
    return sorted(set(before['tokens'])-set(after['tokens']))


def choose_candidate(candidates):
    valid=[c for c in candidates if c.get('eligible')]
    portrait=[c for c in valid if c.get('orientation',1)==1]
    return min(portrait or valid,key=lambda c:(c['metrics']['pages'], -c['metrics']['median_font'])) if valid else None


def _powershell():
    bound=local_binding().get('powershell')
    if bound and Path(bound).is_file():return bound
    return shutil.which('pwsh') or shutil.which('powershell') or str(Path(os.environ.get('SystemRoot','C:/Windows'))/'System32/WindowsPowerShell/v1.0/powershell.exe')


def local_binding():
    path=SCRIPT_DIR.parent/'binding.local.json'
    return json.loads(path.read_text(encoding='utf-8-sig')) if path.exists() else {}


def run_native(request,folder,engine,timeout):
    try:
        return subprocess.run([_powershell(),'-NoProfile','-File',str(SCRIPT_DIR/'native_print.ps1'),'-Request',str(request),'-Engine',engine],
                              capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=timeout)
    except subprocess.TimeoutExpired:
        session=folder/'session.json'
        if session.exists():
            cleanup=subprocess.run([_powershell(),'-NoProfile','-File',str(SCRIPT_DIR/'cleanup_native.ps1'),'-Session',str(session)],
                                   capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=20)
            (folder/'timeout_cleanup.log').write_text(cleanup.stdout+'\n'+cleanup.stderr,encoding='utf-8')
            if cleanup.returncode:raise ValueError('Native Office timed out; owned-process cleanup could not be verified.')
        return subprocess.CompletedProcess([],124,'','Native Office timed out; any verified owned process was cleaned up.')


def description_blocks(values):
    """Keep a multirow description with its first priced line, when headers identify it."""
    desc=quantity=price=None;header=0
    for (r,c),v in values.items():
        text=str(v).strip().lower().replace(' ','')
        if text in ('description','descriptionofgoods','商品描述'):desc=c;header=max(header,r)
        if text in ('quantity','qty','数量'):quantity=c;header=max(header,r)
        if text in ('amount','金额'):price=c;header=max(header,r)
    if not all((desc,quantity,price)):return []
    blocks=[];start=None;previous_priced=None
    for r in range(header+1,max((r for r,c in values),default=header)+1):
        value=values.get((r,desc))
        priced=isinstance(values.get((r,quantity)),(int,float)) and isinstance(values.get((r,price)),(int,float))
        if priced:
            if start is not None and r>start:blocks.append([start,r])
            start=None;previous_priced=r
        elif value is not None and str(value).strip():
            if re.match(r'^\s*(?:GRAND\s+)?TOTAL\b|^\s*合计',str(value),re.I):start=None
            elif start is None:
                continuation=re.match(r'^(?:\d+(?:\.\d+)?\s*[xX*×]|Flat\s*Sheet|Fitted\s*Sheet|Pillowcase)',str(value).strip(),re.I)
                start=previous_priced if continuation and previous_priced is not None and r-previous_priced<=2 else r
    return blocks


def footer_start(values,default=0):
    footer=[];numeric=[];priced=[]
    for (r,c),v in values.items():
        text=re.sub(r'\s+','',str(v)).lower()
        if any(t in text for t in ('bankinformation','bankdetails','bankingdetails','kindlyfindourbank','beneficiary:')):footer.append(r)
    last=max((r for r,c in values),default=0)
    start=min(footer) if footer else last+1
    for r in sorted(set(r for r,c in values)):
        row=[v for (rr,c),v in values.items() if rr==r]
        if r<start and r>=last*.5 and len(row)>=2 and all(isinstance(v,(int,float)) for v in row):numeric.append(r)
        if sum(isinstance(v,(int,float)) for v in row)>=2 and any(isinstance(v,str) for v in row):priced.append(r)
    if numeric:
        start=max(numeric)
        # Keep the final amount with at least its preceding priced line. An
        # overlapping description block then also protects that description.
        preceding=[r for r in priced if r<start]
        if preceding:start=max(preceding)
    elif not footer:return default
    return min(default,start) if default else start


def sheet_metadata(path):
    """Bulk cell/merge discovery avoids slow COM cell-by-cell roundtrips."""
    result={}
    if Path(path).suffix.lower()=='.xls':
        import xlrd
        book=xlrd.open_workbook(path,formatting_info=True)
        for s in book.sheets():
            cells=[];blocks=[];ends=[];total=0;values={}
            for r in range(s.nrows):
                if getattr(s.rowinfo_map.get(r),'hidden',False):continue
                for c in range(s.ncols):
                    if getattr(s.colinfo_map.get(c),'hidden',False):continue
                    value=str(s.cell_value(r,c)).strip()
                    if not value:continue
                    cells.append([r+1,c+1])
                    values[r+1,c+1]=s.cell_value(r,c)
                    if re.match(r'^\s*(?:GRAND\s+)?TOTAL\b|^\s*合计',value,re.I):total=max(total,r+1)
            populated=set(map(tuple,cells))
            for r0,r1,c0,c1 in s.merged_cells:
                if (r0+1,c0+1) in populated:
                    ends.append([r1,c1])
                    if r1-r0>1:blocks.append([r0+1,r1])
            blocks+=description_blocks(values)
            total=footer_start(values,total)
            result[s.name]={'cells':cells,'blocks':blocks,'merge_ends':ends,'total_row':total}
        book.release_resources()
    else:
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        book=load_workbook(path,data_only=False)
        for s in book:
            cells=[];blocks=[];ends=[];total=0;values={}
            hidden_cols=set()
            for k,d in s.column_dimensions.items():
                if d.hidden:hidden_cols.update(range(d.min or 1,(d.max or d.min or 1)+1))
            for row in s:
                for c in row:
                    if c.value is None or not str(c.value).strip() or s.row_dimensions[c.row].hidden or c.column in hidden_cols:continue
                    cells.append([c.row,c.column])
                    values[c.row,c.column]=c.value
                    if re.match(r'^\s*(?:GRAND\s+)?TOTAL\b|^\s*合计',str(c.value),re.I):total=max(total,c.row)
            populated=set(map(tuple,cells))
            for m in s.merged_cells.ranges:
                if (m.min_row,m.min_col) in populated:
                    ends.append([m.max_row,m.max_col])
                    if m.max_row>m.min_row:blocks.append([m.min_row,m.max_row])
            blocks+=description_blocks(values)
            total=footer_start(values,total)
            result[s.title]={'cells':cells,'blocks':blocks,'merge_ends':ends,'total_row':total}
        book.close()
    return result


def prepare(inputs, output, work_root=None, engine='auto', min_median=8, min_p10=7, timeout=240):
    paths=collect_inputs(inputs)
    if not paths:raise ValueError('No Excel input files.')
    output=Path(output).resolve()
    default_work=Path(os.environ.get('LOCALAPPDATA',Path.home()))/'ContractPrintLayout/工作区'
    work_root=Path(work_root or local_binding().get('work_root') or default_work).resolve()
    if output==work_root or output.is_relative_to(work_root) or work_root.is_relative_to(output):
        raise ValueError('Keep the work directory separate from printing results.')
    for p in paths:
        if output==p.parent or work_root==p.parent:raise ValueError('Use separate result and work directories.')
    for raw in inputs:
        root=Path(raw).resolve()
        if root.is_dir() and (output.is_relative_to(root) or work_root.is_relative_to(root)):
            raise ValueError('Put work and results outside recursively processed input directories.')
    batch=datetime.now().strftime('%Y-%m-%d_%H%M%S_%f')+'_'+uuid.uuid4().hex[:6]
    work=work_root/batch;work.mkdir(parents=True,exist_ok=False)
    target=output/batch
    manifest={'version':1,'work_dir':str(work),'output_dir':str(target),'items':[]}
    record=work/'batch.json'
    for index,source in enumerate(paths,1):
        key=sha256(source)[:12];folder=work/(f'{index:03d}_'+key);folder.mkdir()
        item={'id':str(index),'name':source.name,'source':str(source),'source_sha256':sha256(source),'status':'pending'}
        manifest['items'].append(item)
        print(json.dumps({'processing':index,'total':len(paths),'file':source.name},ensure_ascii=False),flush=True)
        try:
            isolated=folder/('source'+source.suffix.lower());shutil.copy2(source,isolated)
            baseline_fingerprint=fingerprint(isolated)
            meta=sheet_metadata(isolated)
            save_json(folder/'request.json',{'source':str(isolated),'work_dir':str(folder),'sheet_meta':meta})
            attempts=[]
            for selected in (['Excel','WPS'] if engine=='auto' else [engine]):
                proc=run_native(folder/'request.json',folder,selected,timeout)
                (folder/(selected+'.log')).write_text(proc.stdout+'\n'+proc.stderr,encoding='utf-8')
                attempts.append({'engine':selected,'exit_code':proc.returncode})
                if proc.returncode==0:break
                # A fallback uses a fresh folder; no collisions with failed drafts.
                if selected=='Excel' and engine=='auto':
                    folder=folder/'wps';folder.mkdir();isolated2=folder/isolated.name;shutil.copy2(isolated,isolated2);isolated=isolated2
                    save_json(folder/'request.json',{'source':str(isolated),'work_dir':str(folder),'sheet_meta':meta})
            if proc.returncode:raise ValueError('Native Office failed; see work logs: '+str(attempts))
            native=json.loads((folder/'native.json').read_text(encoding='utf-8-sig'))
            before=pdf_metrics(folder/'baseline.pdf'); candidates=[]
            for candidate in native['candidates']:
                if candidate.get('error'):
                    candidate['eligible']=False;candidates.append(candidate);continue
                try:
                    patched=folder/(candidate['id']+'_print'+isolated.suffix)
                    patch_print(isolated,candidate['excel'],patched)
                    candidate['native_excel']=candidate['excel'];candidate['excel']=str(patched)
                    intact=fingerprint(patched)==baseline_fingerprint
                except Exception as exc:
                    candidate.update(error=str(exc),eligible=False);candidates.append(candidate);continue
                metrics=pdf_metrics(candidate['pdf'])
                missing=missing_tokens(before,metrics)
                candidate.update(metrics=metrics,structure_preserved=intact,missing_tokens=missing,
                                 reduced_occurrences={t:[n,metrics['token_counts'].get(t,0)] for t,n in before.get('token_counts',{}).items() if metrics['token_counts'].get(t,0)<n},
                                 orientation=candidate['settings'][0]['orientation'])
                candidate['eligible']=intact and not missing and not candidate['unsafe'] and readable(metrics,min_median,min_p10)
                candidates.append(candidate)
            save_json(folder/'checks.json',{'before':before,'candidates':candidates})
            chosen=choose_candidate(candidates)
            exclusions={s['name']:s['excluded'] for s in native['sheets'] if s.get('excluded')}
            item.update(engine=native['engine'],checks=str(folder/'checks.json'),before_pages=before['pages'],range_exclusions=exclusions)
            if chosen:
                # The delivered file must render correctly too: native SaveCopyAs
                # normalizes unrelated document records, so publish only a print
                # metadata transplant into the original, then export it afresh.
                final_pdf=folder/'final.pdf'
                save_json(folder/'export.json',{'source':chosen['excel'],'work_dir':str(folder),'export_only':True,'pdf':str(final_pdf)})
                rendered=run_native(folder/'export.json',folder,native['engine'],timeout)
                (folder/'export.log').write_text(rendered.stdout+'\n'+rendered.stderr,encoding='utf-8')
                if rendered.returncode:raise ValueError('Patched printing file failed native export.')
                final_metrics=pdf_metrics(final_pdf)
                if (missing_tokens(before,final_metrics) or not readable(final_metrics,min_median,min_p10)
                    or final_metrics['page_sizes']!=chosen['metrics']['page_sizes']
                    or abs(final_metrics['median_font']-chosen['metrics']['median_font'])>.1
                    or abs(final_metrics['p10_font']-chosen['metrics']['p10_font'])>.1):
                    raise ValueError('Final printing-file PDF differs from approved candidate checks.')
                chosen['pdf']=str(final_pdf);chosen['metrics']=final_metrics
                item.update(status='awaiting_visual_review',excel=chosen['excel'],pdf=chosen['pdf'],
                            excel_sha256=sha256(chosen['excel']),pdf_sha256=sha256(chosen['pdf']),
                            pages=chosen['metrics']['pages'],metrics=chosen['metrics'],settings=chosen['settings'])
                import fitz
                with fitz.open(chosen['pdf']) as doc:
                    previews=[]
                    for n,page in enumerate(doc,1):
                        png=folder/f'preview-{n}.png';page.get_pixmap(matrix=fitz.Matrix(1.4,1.4)).save(png);previews.append(str(png))
                item['previews']=previews
            else:item['reason']='No candidate passed readability, text preservation, structure and safe-break checks.'
        except Exception as exc:item['reason']=str(exc)
        if sha256(source)!=item['source_sha256']:
            item.update(status='pending',reason='Source changed during processing; no result published.')
        save_json(record,manifest)
    return {'record':str(record),'output_dir':str(target),'drafts':sum(i['status']=='awaiting_visual_review' for i in manifest['items']),
            'pending':sum(i['status']=='pending' for i in manifest['items'])}


def finish(record, reviews):
    record=Path(record).resolve();batch=json.loads(record.read_text(encoding='utf-8-sig'))
    work=Path(batch['work_dir']).resolve();output=Path(batch['output_dir']).resolve()
    if record.parent!=work:raise ValueError('Manifest is outside its work directory.')
    ready=[]
    for item in batch['items']:
        if item['status']!='awaiting_visual_review':continue
        review=reviews.get(item['id'])
        if not review or review.get('status')!='passed':continue
        pdf=Path(item['pdf']).resolve();excel=Path(item['excel']).resolve()
        if not pdf.is_relative_to(work) or not excel.is_relative_to(work):raise ValueError('Draft is outside work directory.')
        if sha256(pdf)!=item['pdf_sha256'] or sha256(excel)!=item['excel_sha256']:raise ValueError('Draft changed after checks.')
        metrics=pdf_metrics(pdf)
        if sorted(set(review.get('pages',[])))!=list(range(1,metrics['pages']+1)) or not review.get('notes','').strip():
            raise ValueError('Review must cover all pages and provide observed findings.')
        if item.get('range_exclusions') and not review.get('range_review','').strip():
            raise ValueError('Existing print area exclusions must be reviewed before publishing.')
        if item.get('source') and sha256(item['source'])!=item['source_sha256']:raise ValueError('Source changed after checks.')
        ready.append((item,review))
    used=set();planned=[]
    for item,review in ready:
        name=Path(item['name']).name
        if name!=item['name']:raise ValueError('Invalid output name.')
        stem=Path(name).stem
        if stem.casefold() in used:name=f'{stem}_{item["id"]}{Path(name).suffix}';stem=Path(name).stem
        while stem.casefold() in used:stem+='_'+item['id'];name=stem+Path(name).suffix
        used.add(stem.casefold())
        exceldir=output/'Excel打印版';pdfdir=output/'PDF打印版'
        destinations=[(Path(item['excel']),exceldir/name),(Path(item['pdf']),pdfdir/(stem+'.pdf'))]
        for src,dst in destinations:
            if dst.exists():raise ValueError('Refusing to overwrite existing result: '+str(dst))
        planned.append((item,review,destinations))
    created=[]
    try:
        for item,review,destinations in planned:
            for src,dst in destinations:
                dst.parent.mkdir(parents=True,exist_ok=True)
                with dst.open('xb') as stream:stream.write(src.read_bytes())
                created.append(dst)
                if sha256(src)!=sha256(dst):raise ValueError('Copy verification failed.')
    except Exception:
        for dst in created:dst.unlink()
        raise
    for item,review,destinations in planned:
        item.update(status='published',visual_review=review,output_excel=str(destinations[0][1]),output_pdf=str(destinations[1][1]))
    save_json(record,batch)
    return {'published':sum(i['status']=='published' for i in batch['items']),
            'pending':sum(i['status']!='published' for i in batch['items']),'output_dir':str(output)}


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    sub=parser.add_subparsers(dest='action',required=True)
    p=sub.add_parser('prepare');p.add_argument('--input',nargs='+',required=True);p.add_argument('--output',required=True)
    p.add_argument('--work-root');p.add_argument('--engine',choices=['auto','Excel','WPS'],default='auto')
    p.add_argument('--min-median',type=float,default=8);p.add_argument('--min-p10',type=float,default=7)
    p.add_argument('--timeout',type=int,default=240)
    p=sub.add_parser('finish');p.add_argument('--record',required=True);p.add_argument('--reviews',required=True)
    sub.add_parser('check')
    args=parser.parse_args(argv)
    if args.action=='prepare':result=prepare(args.input,args.output,args.work_root,args.engine,args.min_median,args.min_p10,args.timeout)
    elif args.action=='finish':result=finish(args.record,json.loads(Path(args.reviews).read_text(encoding='utf-8-sig')))
    else:
        import fitz,openpyxl,xlrd
        result={'python':sys.executable,'powershell':_powershell(),'native_script':str(SCRIPT_DIR/'native_print.ps1'),
                'office':'Excel/WPS COM is checked when creating an isolated batch.'}
    print(json.dumps(result,ensure_ascii=False,indent=2))


if __name__=='__main__':main()
