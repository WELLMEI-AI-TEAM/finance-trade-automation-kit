"""Independently compare ZIP entries, manifest hashes and repository sources."""
from pathlib import Path
import hashlib
import json
import re
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def digest(data):
    return hashlib.sha256(data).hexdigest()


def main():
    meta = json.loads((ROOT/'suite.json').read_text('utf-8'))
    dest = ROOT/'dist/upload'
    sums = dict(line.split('  ',1)[::-1] for line in (dest/'SHA256SUMS.txt').read_text().splitlines())
    for skill in meta['skills']:
        archive = dest/f'{skill}-v{meta["version"]}.zip'
        manifest = json.loads((dest/(archive.stem+'.manifest.json')).read_text('utf-8'))
        assert digest(archive.read_bytes()) == manifest['sha256'] == sums[archive.name]
        assert manifest['skill'] == skill and manifest['version'] == meta['version']
        with zipfile.ZipFile(archive) as z:
            assert z.testzip() is None
            assert len(z.namelist()) == len(set(z.namelist()))
            assert set(z.namelist()) == set(manifest['files'])
            assert 'SKILL.md' in z.namelist()
            for name in z.namelist():
                p = Path(name)
                assert not p.is_absolute() and '..' not in p.parts
                assert not any(x.startswith('.') or x in {'profiles','__pycache__'} for x in p.parts)
                assert p.suffix in {'.md','.py','.ps1','.yaml','.yml','.txt'}
                assert digest(z.read(name)) == manifest['files'][name]
                source = ROOT/'LICENSE' if name == 'LICENSE.txt' else ROOT/'skills'/skill/name
                assert source.read_bytes() == z.read(name)
                if p.suffix == '.md':
                    for target in re.findall(r'\]\(([^)]+)\)',z.read(name).decode('utf-8-sig')):
                        if '://' not in target and not target.startswith('#'):
                            import posixpath
                            resolved = posixpath.normpath(posixpath.join(p.parent.as_posix(),target.split('#')[0]))
                            assert resolved in z.namelist(), (name,target)
        print(f'PASS {archive.name}: {len(manifest["files"])} entries verified')


if __name__ == '__main__':
    main()
