"""Independent read-only BIFF audit after native XLS editing. Requires xlrd."""
import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import struct
import xlrd
from apply_ooxml_plan import date_text, invoice_label_text


def font(book, index):
    f=book.font_list[index]
    return {k:getattr(f,k) for k in ('name','height','colour_index','weight','escapement','underline_type','family','character_set','bold','italic','underlined','struck_out','outline','shadow')}


def style(book, index):
    xf=book.xf_list[index]
    return {'font':font(book,xf.font_index), 'format':book.format_map[xf.format_key].format_str,
            'alignment':vars(xf.alignment),'border':vars(xf.border),
            'background':vars(xf.background),'protection':vars(xf.protection),
            'lotus_123_prefix':xf.lotus_123_prefix}


def biff_controls(book):
    # Preserve formula tokens, print setup and drawing records independently of
    # the application model's normalized row/column/font representation.
    records=[];pos=book.base;end=book.base+book.stream_len;sheet=-1;continuation=False
    keep={0x14,0x15,0x83,0x84,0x26,0x27,0x28,0x29,0xA1,0x2A,0x2B,0x82,0x1B,0x1A,0xEB,0xEC}
    while pos+4<=end:
        code,length=struct.unpack_from('<HH',book.mem,pos);pos+=4
        data=bytes(book.mem[pos:pos+length]);pos+=length
        if code==0x809 and len(data)>=4 and struct.unpack_from('<H',data,2)[0] in (0x10,0x20,0x40):sheet+=1
        if code==0x06 and len(data)>=22:
            row,col=struct.unpack_from('<HH',data)
            size=struct.unpack_from('<H',data,20)[0]
            records.append((sheet,'formula',row,col,data[22:22+size].hex()))
        if code in keep or (code==0x3C and continuation):
            records.append((sheet,code,data.hex()))
            continuation=code in (0xEB,0xEC,0x3C)
        else:continuation=False
    return records


def audit(plan):
    source=Path(plan['source']);output=Path(plan['output'])
    if source.suffix.lower()!='.xls' or output.suffix.lower()!='.xls':raise ValueError('BIFF audit requires xls inputs.')
    a=xlrd.open_workbook(source,formatting_info=True,on_demand=True)
    b=xlrd.open_workbook(output,formatting_info=True,on_demand=True)
    errors=[];changes=[]
    try:
        if a.sheet_names()!=b.sheet_names():errors.append('sheet names/order')
        if a.datemode!=b.datemode:errors.append('date system')
        if a.colour_map!=b.colour_map:errors.append('colour palette')
        expected={}
        date=plan['date'];pi=dt.date.fromisoformat(date['pi_date']);purchase=dt.date.fromisoformat(plan['association']['purchase_contract']['signing_date'])
        target=min(pi,purchase)
        if isinstance(date['expected'],str):newdate=date_text(date,target)
        else:newdate=date['expected']+(target-pi).days
        expected[(date['sheet'],date['cell'])]=newdate
        title=plan.get('title')
        if title:
            old=title['expected'];new=old[:len(old)-len(old.lstrip())]+'CONTRACT'+old[len(old.rstrip()):]
            expected[(title['sheet'],title['cell'])]=new
        label=plan.get('invoice_label')
        if label:
            expected[(label['sheet'],label['cell'])]=invoice_label_text(label['expected'])
        for sa,sb in zip(a.sheets(),b.sheets()):
            if (sa.nrows,sa.ncols)!=(sb.nrows,sb.ncols):errors.append(sa.name+': used dimensions')
            if sa.merged_cells!=sb.merged_cells:errors.append(sa.name+': merges')
            for r in range(max(sa.nrows,sb.nrows)):
                for c in range(max(sa.ncols,sb.ncols)):
                    key=xlrd.formula.colname(c)+str(r+1)
                    if r>=min(sa.nrows,sb.nrows) or c>=min(sa.ncols,sb.ncols):continue
                    ca=sa.cell(r,c);cb=sb.cell(r,c)
                    want=expected.get((sa.name,key),ca.value)
                    if cb.value!=want:errors.append(f'{sa.name}!{key}: value')
                    if ca.value!=cb.value:changes.append(f'{sa.name}!{key}')
                    if style(a,ca.xf_index)!=style(b,cb.xf_index):errors.append(f'{sa.name}!{key}: style')
                    if ca.ctype!=cb.ctype:errors.append(f'{sa.name}!{key}: type')
            for row in set(sa.rowinfo_map)|set(sb.rowinfo_map):
                x=sa.rowinfo_map.get(row);y=sb.rowinfo_map.get(row)
                if x is None or y is None:errors.append(f'{sa.name}: row {row+1} metadata');continue
                attrs=('height','has_default_height','height_mismatch','outline_level','outline_group_starts_ends','hidden','additional_space_above','additional_space_below')
                if any(getattr(x,k,None)!=getattr(y,k,None) for k in attrs):errors.append(f'{sa.name}: row {row+1} dimensions/flags')
            for col in set(sa.colinfo_map)|set(sb.colinfo_map):
                x=sa.colinfo_map.get(col);y=sb.colinfo_map.get(col)
                if x is None or y is None:errors.append(f'{sa.name}: column {col+1} metadata');continue
                if any(getattr(x,k,None)!=getattr(y,k,None) for k in ('width','hidden','outline_level','collapsed')):errors.append(f'{sa.name}: column {col+1} dimensions/flags')
                if style(a,x.xf_index)!=style(b,y.xf_index):errors.append(f'{sa.name}: column {col+1} style')
            for coords in set(sa.rich_text_runlist_map)|set(sb.rich_text_runlist_map):
                x=[(offset,font(a,index)) for offset,index in sa.rich_text_runlist_map.get(coords,[])]
                y=[(offset,font(b,index)) for offset,index in sb.rich_text_runlist_map.get(coords,[])]
                if x!=y:errors.append(f'{sa.name}: rich text {coords}')
        names=lambda book: sorted((n.name,n.scope,getattr(n.result,'text',str(n.raw_formula))) for n in book.name_obj_list)
        if names(a)!=names(b):errors.append('defined names/print areas')
        if biff_controls(a)!=biff_controls(b):errors.append('BIFF formula/print/drawing records')
        return {'status':'passed' if not errors else 'failed','source':str(source),'output':str(output),
                'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'output_sha256':hashlib.sha256(output.read_bytes()).hexdigest(),
                'cell_changes':changes,'errors':errors,
                'scope':'BIFF values, formats/fonts, rich text, row/column layout, merges, palette, names, formula tokens and print/drawing records; visual review and unsupported OLE objects remain separate.'}
    finally:
        a.release_resources();b.release_resources()


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--plan',type=Path,required=True);args=parser.parse_args()
    plan=json.loads(args.plan.read_text(encoding='utf-8-sig'));result=audit(plan)
    report=Path(str(plan['output'])+'.preservation.json')
    with report.open('x',encoding='utf-8') as f:json.dump(result,f,ensure_ascii=False,indent=2)
    print(json.dumps({'status':result['status'],'errors':len(result['errors']),'report':str(report)},ensure_ascii=False))
    raise SystemExit(0 if result['status']=='passed' else 1)
