"""Transplant only native print records; leave contract content and objects intact."""
from pathlib import Path
import struct
import zipfile
from lxml import etree as ET


def patch_xlsx(source, native, output):
    ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    tag=lambda name:'{'+ns+'}'+name
    targets={'pageSetup','pageMargins','printOptions','rowBreaks','colBreaks'}
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(native) as b, zipfile.ZipFile(output,'x',zipfile.ZIP_DEFLATED) as out:
        for item in a.infolist():
            raw=a.read(item.filename)
            if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml') and item.filename in b.namelist():
                original=ET.fromstring(raw);edited=ET.fromstring(b.read(item.filename))
                # Insert using source/native schema order, preserving every cell,
                # drawing reference, merge and unrelated worksheet property.
                for name in targets:
                    old=original.find(tag(name));new=edited.find(tag(name))
                    if old is not None:
                        index=list(original).index(old);original.remove(old)
                    else:
                        order=['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst']
                        rank=order.index(name)
                        index=next((i for i,c in enumerate(original) if c.tag.split('}')[-1] in order and order.index(c.tag.split('}')[-1])>rank),len(original))
                    if new is not None:
                        new.attrib.pop('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id',None)
                        original.insert(index,new)
                new_pr=edited.find(tag('sheetPr')+'/'+tag('pageSetUpPr'))
                if new_pr is not None:
                    prop=original.find(tag('sheetPr'))
                    if prop is None:prop=ET.Element(tag('sheetPr'));original.insert(0,prop)
                    old=prop.find(tag('pageSetUpPr'))
                    if old is not None:prop.remove(old)
                    prop.append(new_pr)
                raw=ET.tostring(original,encoding='utf-8',xml_declaration=True)
            elif item.filename=='xl/workbook.xml':
                original=ET.fromstring(raw);edited=ET.fromstring(b.read(item.filename))
                definitions=original.find(tag('definedNames'))
                newdefs=edited.find(tag('definedNames'))
                if newdefs is not None:
                    if definitions is None:
                        definitions=ET.Element(tag('definedNames'))
                        sheets=original.find(tag('sheets'));original.insert(list(original).index(sheets)+1,definitions)
                    for old in list(definitions):
                        if old.get('name')=='_xlnm.Print_Area':definitions.remove(old)
                    for new in newdefs:
                        if new.get('name')=='_xlnm.Print_Area':definitions.append(new)
                    raw=ET.tostring(original,encoding='utf-8',xml_declaration=True)
            out.writestr(item,raw)


def _chain(doc, first):
    result=[];seen=set();sid=first
    while sid>=0:
        if sid in seen or sid>=len(doc.SAT):raise ValueError('Invalid OLE chain.')
        result.append(sid);seen.add(sid);sid=doc.SAT[sid]
    if sid!=-2:raise ValueError('Invalid OLE chain terminator.')
    return result


def _stream(path):
    from xlrd.compdoc import CompDoc
    raw=Path(path).read_bytes();doc=CompDoc(raw)
    if doc.sec_size!=512:raise ValueError('Unsupported OLE sector size.')
    items=[x for x in doc.dirlist if x.name in ('Workbook','Book') and x.etype==2]
    if len(items)!=1 or items[0].tot_size<4096:raise ValueError('Unsupported Workbook stream.')
    item=items[0];sectors=_chain(doc,item.first_SID)
    data=b''.join(raw[512+s*512:512+(s+1)*512] for s in sectors)[:item.tot_size]
    return raw,doc,item,sectors,data


def _records(data):
    out=[];p=0
    while p+4<=len(data):
        code,size=struct.unpack_from('<HH',data,p)
        if code==size==0 and not any(data[p:]):break
        if p+4+size>len(data):raise ValueError('Truncated BIFF record.')
        out.append((p,code,data[p+4:p+4+size]));p+=4+size
    if any(data[p:]):raise ValueError('Unsupported BIFF tail.')
    return out


def patch_xls(source, native, output):
    raw,doc,item,sectors,data=_stream(source)
    records=_records(data);newrecords=_records(_stream(native)[4])
    if any(c in (0x2f,0xd3) for _,c,_ in records):raise ValueError('Encrypted or legacy BIFF layout.')
    def sheets(recs):
        return {struct.unpack_from('<I',d)[0]:i for i,(_,c,d) in enumerate(x for x in recs if x[1]==0x85)}
    offsets=sheets(records);native_offsets=sheets(newrecords)
    native_print={};sheet=None
    codes={0xA1,0x26,0x27,0x28,0x29,0x83,0x84,0x1B,0x1A}
    native_wsbool={}
    for p,c,d in newrecords:
        if p in native_offsets:sheet=native_offsets[p]
        if sheet is not None and c in codes:native_print[(sheet,c)]=d
        if sheet is not None and c==0x81:native_wsbool[sheet]=d
    existing=set();sheet=None
    for p,c,d in records:
        if p in offsets:sheet=offsets[p]
        if sheet is not None and c in codes:existing.add((sheet,c))
    chunks=[];sheet=None
    for p,c,d in records:
        if p in offsets:sheet=offsets[p]
        if sheet is not None and c==0xA1:
            for missing in sorted(codes):
                if (sheet,missing) not in existing and (sheet,missing) in native_print:
                    chunks.append((None,missing,native_print[(sheet,missing)]))
        if sheet is not None and c in codes:
            if (sheet,c) not in native_print:continue
            d=native_print[(sheet,c)]
        if sheet is not None and c==0x81 and sheet in native_wsbool:
            flags=struct.unpack_from('<H',d)[0]
            newflag=struct.unpack_from('<H',native_wsbool[sheet])[0]&0x100
            d=struct.pack('<H',(flags&~0x100)|newflag)+d[2:]
        chunks.append((p,c,d))
    mapping={};pos=0
    for old,c,d in chunks:
        if old is not None:mapping[old]=pos
        pos+=4+len(d)
    result=bytearray()
    for old,c,d in chunks:
        if c==0x85:
            pointer=struct.unpack_from('<I',d)[0];d=struct.pack('<I',mapping[pointer])+d[4:]
        elif c==0x20b:
            d=bytearray(d)
            for at in range(16,len(d),4):
                pointer=struct.unpack_from('<I',d,at)[0]
                if pointer:
                    if pointer not in mapping:raise ValueError('Unsupported BIFF INDEX pointer.')
                    struct.pack_into('<I',d,at,mapping[pointer])
        result.extend(struct.pack('<HH',c,len(d)));result.extend(d)
    if len(result)>len(sectors)*512:raise ValueError('Original OLE stream has insufficient print-record capacity.')
    result.extend(b'\0'*max(0,len(data)-len(result)))
    patched=bytearray(raw)
    for n,sector in enumerate(sectors):
        block=result[n*512:(n+1)*512];at=512+sector*512
        patched[at:at+len(block)]=block
    directories=_chain(doc,doc.dir_first_sec_sid);logical=item.DID*128+120
    at=512+directories[logical//512]*512+logical%512
    struct.pack_into('<Q',patched,at,len(result))
    Path(output).write_bytes(patched)


def patch_print(source, native, output):
    if Path(output).exists():raise ValueError('Output already exists.')
    if Path(source).suffix.lower()=='.xls':patch_xls(source,native,output)
    else:patch_xlsx(source,native,output)
