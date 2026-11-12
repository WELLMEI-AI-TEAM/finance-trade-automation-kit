"""Build independent skill packages from explicit resource directories."""
from pathlib import Path
import hashlib
import json
import re
import zipfile

ROOT = Path(__file__).resolve().parents[1]
ALLOWED = {'.md', '.py', '.ps1', '.yaml', '.yml', '.txt'}


def resources(skill):
    selected = [skill/'SKILL.md', skill/'requirements.txt']
    for folder in ('agents', 'references', 'scripts'):
        if (skill/folder).is_dir():
            selected += [p for p in (skill/folder).rglob('*') if p.is_file()
                         and p.suffix in ALLOWED
                         and not any(x.startswith('.') or x == '__pycache__'
                                     for x in p.relative_to(skill).parts)]
    assert all(p.is_file() and skill.resolve() in p.resolve().parents for p in selected)
    paths = {p.resolve() for p in selected}
    for p in selected:
        if p.suffix == '.md':
            for target in re.findall(r'\]\(([^)]+)\)', p.read_text('utf-8-sig')):
                if '://' in target or target.startswith('#'):
                    continue
                assert (p.parent/target.split('#')[0]).resolve() in paths, (p, target)
    return sorted(selected)


def main():
    meta = json.loads((ROOT/'suite.json').read_text('utf-8'))
    dest = ROOT/'dist/upload'
    dest.mkdir(parents=True, exist_ok=True)
    sums = []
    for name in meta['skills']:
        skill = ROOT/'skills'/name
        data = {p.relative_to(skill).as_posix():p.read_bytes() for p in resources(skill)}
        assert not data['SKILL.md'].startswith(b'\xef\xbb\xbf'), 'SKILL.md must be UTF-8 without BOM'
        assert f'name: {name}' in data['SKILL.md'].decode('utf-8')
        data['LICENSE.txt'] = (ROOT/'LICENSE').read_bytes()
        archive = dest/f'{name}-v{meta["version"]}.zip'
        with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
            for path, content in sorted(data.items()):
                info = zipfile.ZipInfo(path, (2026,11,12,0,0,0))
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, content)
        manifest = {'skill':name, 'version':meta['version'], 'archive':archive.name,
                    'sha256':hashlib.sha256(archive.read_bytes()).hexdigest(),
                    'files':{p:hashlib.sha256(c).hexdigest() for p,c in sorted(data.items())}}
        (dest/(archive.stem+'.manifest.json')).write_text(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n','utf-8')
        sums.append(manifest['sha256']+'  '+archive.name)
        print(f'{name}: {len(data)} files; {archive.stat().st_size} bytes')
    (dest/'SHA256SUMS.txt').write_text('\n'.join(sums)+'\n','utf-8')


if __name__ == '__main__':
    main()
