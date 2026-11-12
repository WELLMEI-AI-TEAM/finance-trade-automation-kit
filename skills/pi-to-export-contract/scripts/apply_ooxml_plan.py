"""Edit only approved PI cell payloads; preserve every other OOXML package byte.

Uses Python standard library only. Input is the same evidence-based JSON plan as
the native Excel/WPS helper. No workbook import/export or formula recalculation.
"""
import argparse
import datetime as dt
from decimal import Decimal
import hashlib
import html
import json
import os
from pathlib import Path
import posixpath
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
MONTHS = ('January February March April May June July August September October November December').split()


def require(condition, message):
    if not condition:
        raise ValueError(message)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def invoice_label_text(value):
    require(isinstance(value, str), 'Invoice number label must be text.')
    match = re.match(r'^(\s*)(?:Invoice[ \t]+)?No\.', value, re.I)
    require(match is not None, 'Target is not an Invoice No. / No. label.')
    return match.group(1) + 'No.' + value[match.end():]


def date_text(spec, day):
    pattern = spec.get('format', '')
    require(bool(pattern), 'Text date requires original format.')
    mapping = {'yyyy': f'{day.year:04}', 'yy': f'{day.year % 100:02}',
               'MMMM': MONTHS[day.month-1], 'MMM': MONTHS[day.month-1][:3],
               'MM': f'{day.month:02}', 'M': str(day.month), 'dd': f'{day.day:02}', 'd': str(day.day)}
    result = ''
    i = 0
    while i < len(pattern):
        if pattern.startswith('{ordinal}', i):
            suffix = 'th' if 11 <= day.day % 100 <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day.day % 10, 'th')
            ordinal_case = spec.get('ordinal_case', 'lower')
            require(ordinal_case in ('lower', 'upper'), 'Unknown ordinal_case.')
            result += suffix.upper() if ordinal_case == 'upper' else suffix
            i += len('{ordinal}')
            continue
        if pattern[i] == "'":
            end = pattern.find("'", i+1)
            require(end >= 0, 'Unclosed literal in date format.')
            result += pattern[i+1:end]
            i = end+1
            continue
        token = next((t for t in mapping if pattern.startswith(t, i)), None)
        if token:
            result += mapping[token]
            i += len(token)
        else:
            require(not pattern[i].isascii() or not pattern[i].isalpha(), 'Unsupported date token; preserve via separate handling.')
            result += pattern[i]
            i += 1
    case = spec.get('text_case', 'none')
    require(case in ('none', 'upper', 'lower'), 'Unknown text_case.')
    if case != 'none':
        result = getattr(result, case)()
    return spec.get('prefix', '') + result + spec.get('suffix', '')


def sheet_paths(parts):
    book = ET.fromstring(parts['xl/workbook.xml'])
    rels = ET.fromstring(parts['xl/_rels/workbook.xml.rels'])
    targets = {}
    for rel in rels:
        if rel.get('TargetMode') != 'External':
            target = rel.get('Target', '')
            targets[rel.get('Id')] = (target.lstrip('/') if target.startswith('/')
                                     else posixpath.normpath(posixpath.join('xl', target)))
    result = {}
    for sheet in book.findall(f'{{{NS}}}sheets/{{{NS}}}sheet'):
        result[sheet.get('name')] = targets[sheet.get(f'{{{REL}}}id')]
    return book, result


def cell_value(parts, paths, spec):
    require(re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]*', spec['cell']) is not None, 'Use one A1 cell address.')
    require(spec['sheet'] in paths, 'PI worksheet not found.')
    part = paths[spec['sheet']]
    sheet = ET.fromstring(parts[part])
    matches = [c for c in sheet.iter(f'{{{NS}}}c') if c.get('r') == spec['cell']]
    require(len(matches) == 1, 'Target must be one existing cell.')
    cell = matches[0]
    require(cell.find(f'{{{NS}}}f') is None, 'Refusing formula overwrite.')
    kind = cell.get('t', 'n')
    v = cell.find(f'{{{NS}}}v')
    if kind == 's':
        shared = ET.fromstring(parts['xl/sharedStrings.xml'])
        item = shared.findall(f'{{{NS}}}si')[int(v.text)]
        require(item.find(f'{{{NS}}}r') is None, 'Rich text requires separate preservation handling.')
        value = ''.join(t.text or '' for t in item.iter(f'{{{NS}}}t'))
    elif kind == 'inlineStr':
        item = cell.find(f'{{{NS}}}is')
        require(item is not None and item.find(f'{{{NS}}}r') is None, 'Rich text requires separate preservation handling.')
        value = ''.join(t.text or '' for t in item.iter(f'{{{NS}}}t'))
    elif kind == 'str':
        value = '' if v is None else v.text or ''
    elif kind in ('n', ''):
        require(v is not None and v.text, 'Numeric date cell is empty.')
        value = Decimal(v.text)
    else:
        raise ValueError(f'Unsupported target cell type: {kind}')
    expected = spec['expected']
    if isinstance(value, Decimal):
        require(not isinstance(expected, (str, bool)) and value == Decimal(str(expected)), 'Numeric target differs from expected original.')
    else:
        require(isinstance(expected, str) and value == expected, 'Text target differs from expected original.')
    return part, cell, value


def replace_cell(data, address, new_value):
    text = data.decode('utf-8')
    pattern = re.compile(r'<(?P<tag>(?:[\w.-]+:)?c)\b[^>]*\br\s*=\s*["\']' + re.escape(address)
                         + r'["\'][^>]*(?:/>|>[\s\S]*?</(?P=tag)\s*>)')
    matches = list(pattern.finditer(text))
    require(len(matches) == 1, 'Cannot identify unique raw XML target.')
    match = matches[0]
    original = match.group()
    opening = original[:original.index('>')+1]
    require(not opening.endswith('/>'), 'Target is empty.')
    tag = match.group('tag')
    prefix = tag[:-1]
    # Preserve all cell attributes other than the payload type. Styles,
    # coordinates, row/column geometry and the surrounding XML stay untouched.
    opening = re.sub(r'\s+t\s*=\s*["\'][^"\']*["\']', '', opening)
    if isinstance(new_value, str):
        opening = opening[:-1] + ' t="inlineStr">'
        content = f'<{prefix}is><{prefix}t xml:space="preserve">{html.escape(new_value, quote=False)}</{prefix}t></{prefix}is>'
    else:
        content = f'<{prefix}v>{new_value}</{prefix}v>'
    # All existing non-payload child nodes must remain (metadata/extLst, etc.).
    # Formula targets were rejected during inspection above.
    interior = original[original.index('>')+1:original.rfind('</')]
    payload = re.compile(r'<(?:[\w.-]+:)?(?:v|is)\b[^>]*(?:/>|>[\s\S]*?</(?:[\w.-]+:)?(?:v|is)\s*>)')
    found = list(payload.finditer(interior))
    require(len(found) == 1, 'Expected exactly one existing cell payload.')
    payload_match = found[0]
    interior = interior[:payload_match.start()] + content + interior[payload_match.end():]
    replacement = opening + interior + f'</{tag}>'
    return (text[:match.start()] + replacement + text[match.end():]).encode('utf-8')


def apply_plan(plan):
    source = Path(plan['source'])
    output = Path(plan['output'])
    audit_path = Path(str(output)+'.audit.json')
    require(source.is_absolute() and output.is_absolute(), 'Absolute source and output paths required.')
    require(source.resolve() != output.resolve(), 'Cannot overwrite original PI.')
    require(source.suffix.lower() in ('.xlsx', '.xlsm'), 'Exact package editor supports xlsx/xlsm only.')
    require(source.suffix.lower() == output.suffix.lower(), 'Keep the original extension.')
    require(not output.exists() and not audit_path.exists(), 'Output exists; choose a new name.')
    require(sha(source).lower() == plan['source_sha256'].lower(), 'Source hash differs from inspected PI.')
    association = plan['association']
    require(association['status'] == 'confirmed', 'Association is not confirmed.')
    originals = [{'role': 'PI', 'path': str(source), 'sha256': sha(source)}]
    for role in ('order_confirmation', 'purchase_contract'):
        record = association[role]
        path = Path(record['path'])
        require(path.is_absolute() and path.is_file() and record.get('locator') and record.get('evidence'), 'Actual input and association evidence required.')
        originals.append({'role': role, 'path': str(path), 'sha256': sha(path)})
    spec = plan.get('date')
    require(spec is not None, 'Verified PI document date required; otherwise mark pending.')
    targets = [item for item in (plan.get('title'), plan.get('invoice_label'), spec) if item is not None]
    require(len({(item['sheet'], item['cell']) for item in targets}) == len(targets),
            'Title, invoice label and date must use different cells; combined fields require separate handling.')
    pi_date = dt.date.fromisoformat(spec['pi_date'])
    purchase_date = dt.date.fromisoformat(association['purchase_contract']['signing_date'])
    result_date = min(pi_date, purchase_date)
    with zipfile.ZipFile(source) as z:
        entries = z.infolist()
        require(len({i.filename for i in entries}) == len(entries), 'Duplicate package entries require inspection.')
        parts = {i.filename: z.read(i.filename) for i in entries}
        zip_comment = z.comment
    require(not any(n.startswith('_xmlsignatures/') for n in parts), 'Digitally signed PI requires separate handling.')
    original_parts = dict(parts)
    book, paths = sheet_paths(parts)
    part, cell, original_date = cell_value(parts, paths, spec)
    if isinstance(original_date, str):
        require(date_text(spec, pi_date) == original_date, 'Original date format cannot reproduce the text exactly.')
        target_date = date_text(spec, result_date)
    else:
        props = book.find(f'{{{NS}}}workbookPr')
        date1904 = props is not None and props.get('date1904', '').lower() in ('1', 'true')
        epoch = dt.date(1904, 1, 1) if date1904 else dt.date(1899, 12, 30)
        require(epoch + dt.timedelta(days=int(original_date)) == pi_date, 'Numeric value does not match verified PI date.')
        target_date = original_date + Decimal((result_date-pi_date).days)
    changes = []
    title = plan.get('title')
    if title is not None:
        require((title['sheet'], title['cell']) != (spec['sheet'], spec['cell']), 'Title and date are different fields.')
        title_part, title_cell, original_title = cell_value(parts, paths, title)
        require(isinstance(original_title, str) and re.fullmatch(r'\s*(?:PRO\s*[-–]?\s*FORMA\s*(?:INVOICE)?|INVOICE|PI|CONTRACT)\s*', original_title, re.I), 'Target is not a standalone document title.')
        leading = original_title[:len(original_title)-len(original_title.lstrip())]
        trailing = original_title[len(original_title.rstrip()):]
        target_title = leading+'CONTRACT'+trailing
        if target_title != original_title:
            parts[title_part] = replace_cell(parts[title_part], title['cell'], target_title)
            changes.append({'kind': 'title', 'sheet': title['sheet'], 'cell': title['cell'], 'before': original_title, 'after': target_title})
    if result_date != pi_date:
        parts[part] = replace_cell(parts[part], spec['cell'], target_date)
        changes.append({'kind': 'date', 'sheet': spec['sheet'], 'cell': spec['cell'], 'before': str(original_date), 'after': str(target_date)})
    label = plan.get('invoice_label')
    if label is not None:
        label_part, _, original_label = cell_value(parts, paths, label)
        target_label = invoice_label_text(original_label)
        if target_label != original_label:
            parts[label_part] = replace_cell(parts[label_part], label['cell'], target_label)
            changes.append({'kind': 'invoice_label', 'sheet': label['sheet'], 'cell': label['cell'],
                            'before': original_label, 'after': target_label})
    modified = [n for n in parts if parts[n] != original_parts[n]]
    allowed_parts = {paths[spec['sheet']]}
    if title is not None:
        allowed_parts.add(paths[title['sheet']])
    if label is not None:
        allowed_parts.add(paths[label['sheet']])
    require(set(modified) <= allowed_parts, 'Unexpected package mutation.')
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix='.contract-', suffix=output.suffix, dir=output.parent)
    os.close(fd)
    temp = Path(name)
    temp_audit = Path(name+'.audit.json')
    published = False
    try:
        if not modified:
            shutil.copyfile(source, temp)
        else:
            with zipfile.ZipFile(temp, 'w') as z:
                z.comment = zip_comment
                for entry in entries:
                    z.writestr(entry, parts[entry.filename])
        with zipfile.ZipFile(temp) as z:
            require(z.namelist() == [e.filename for e in entries], 'Package entries changed.')
            for entry in entries:
                require(z.read(entry.filename) == parts[entry.filename], 'Saved package differs from verified patch.')
        after_spec = dict(spec, expected=(str(target_date) if isinstance(target_date, str) else float(target_date)))
        _, _, actual = cell_value(parts, paths, after_spec)
        require(actual == target_date, 'Saved date does not match planned result.')
        if title is not None:
            cell_value(parts, paths, dict(title, expected=target_title))
        if label is not None:
            cell_value(parts, paths, dict(label, expected=target_label))
        for item in originals:
            require(sha(item['path']) == item['sha256'], 'Original input changed during execution.')
        audit = {'status': 'verified', 'engine': 'exact-ooxml-package-patch',
                 'inputs': originals, 'association': association, 'output': str(output), 'output_sha256': sha(temp),
                 'pi_date': pi_date.isoformat(), 'purchase_signing_date': purchase_date.isoformat(),
                 'date_result': result_date.isoformat(), 'date_changed': result_date != pi_date,
                 'changes': changes, 'modified_parts': modified, 'all_other_package_parts_identical': True,
                 'original_inputs_unchanged': True, 'visual_review': 'not_performed_by_script',
                 'preservation': 'Only planned cell payloads/type attributes changed; all styles, dimensions, objects, formula caches and other package parts retain their original bytes.'}
        temp_audit.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding='utf-8')
        require(json.loads(temp_audit.read_text(encoding='utf-8'))['output_sha256'] == sha(temp), 'Audit readback failed.')
        # Hard links provide exclusive publication on Windows and POSIX. Never
        # replace an existing file if another run created it meanwhile.
        os.link(temp, output)
        published = True
        try:
            os.link(temp_audit, audit_path)
        except BaseException:
            output.unlink()
            published = False
            raise
        return audit
    finally:
        temp.unlink(missing_ok=True)
        temp_audit.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--plan', required=True, type=Path)
    args = parser.parse_args()
    plan = json.loads(args.plan.read_text(encoding='utf-8-sig'))
    audit = apply_plan(plan)
    print(json.dumps({'output': audit['output'], 'date_changed': audit['date_changed'], 'modified_parts': audit['modified_parts']}, ensure_ascii=False))


if __name__ == '__main__':
    main()
