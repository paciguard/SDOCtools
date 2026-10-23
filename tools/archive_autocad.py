"""Retire the superseded add-in only after end-to-end standalone checks pass."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import shutil

root = Path(__file__).resolve().parents[1]
assert 'STANDALONE_API_OK' in (root/'tests/standalone-api.log').read_text(encoding='utf-8-sig'), 'Standalone integration checks must pass first'
targets = [root/'ZTools.AutoCAD', root/'deploy', root/'dist/ZTools.bundle']
files = []
for target in targets:
    assert target.resolve().is_relative_to(root) and target.is_dir(), target
    files.extend(p for p in target.rglob('*') if p.is_file())
extra = [root/'ZTools.sln', root/'tests/WorkbookSmoke.cs', root/'tests/WorkbookSmoke.csproj']
extra.extend((root/'tests').glob('*.scr'))
extra.extend(p for p in (root/'.tools/archive-during-migration').rglob('*') if p.is_file())
files.extend(extra)
complete = root/'.tools/archived-autocad-complete-20261023.zip'
source = root/'migration/archived-autocad-addin-20261023.zip'
assert not complete.exists() and not source.exists(), 'Do not overwrite an existing recovery archive'
with ZipFile(complete, 'w', ZIP_DEFLATED) as z:
    for p in files:z.write(p, p.relative_to(root).as_posix())
with ZipFile(complete) as z:
    assert z.testzip() is None
    for p in files:
        assert hashlib.sha256(z.read(p.relative_to(root).as_posix())).digest() == hashlib.sha256(p.read_bytes()).digest(), p
with ZipFile(source, 'w', ZIP_DEFLATED) as z:
    for p in files:
        rel = p.relative_to(root)
        if not {'bin','obj','dist'} & set(rel.parts):
            z.write(p, rel.as_posix())
with ZipFile(source) as z:assert z.testzip() is None
print(f'Complete recovery archive verified against all {len(files)} original files: {complete}')
print(f'Portable source recovery archive: {source}')
# Targets are explicit, resolved, and verified inside the workspace. Every file
# in each target has already been byte-checked in the complete recovery archive.
for target in targets:shutil.rmtree(target)
print('Removed the superseded active add-in, deployment manifest, and old bundle.')
