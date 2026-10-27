"""Verify all original reference files against the selected source snapshot."""
import hashlib
import argparse
import json
from pathlib import Path

root = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser()
current = 'migration/source-manifest-20261023-embedded.json'
parser.add_argument('--manifest', default=current if (root/current).is_file() else 'migration/source-manifest.json', help='Reference snapshot relative to the project root')
args = parser.parse_args()
print('Reference snapshot:', args.manifest)
manifest = json.loads((root / args.manifest).read_text(encoding='utf-8'))
for entry in manifest:
    path = root / entry['path']
    assert path.is_file(), f'Missing original: {path}'
    assert path.stat().st_size == entry['bytes'], f'Original size changed: {path}'
    assert hashlib.sha256(path.read_bytes()).hexdigest() == entry['sha256'], f'Original hash changed: {path}'
print(f'PASS: all {len(manifest)} original reference files are unchanged byte for byte.')
split = root/'migration/workbook-split-20261026.json'
if split.is_file():
    provenance = json.loads(split.read_text(encoding='utf-8'))
    for key, hash_key in [('source','sourceSha256'),('embedded','embeddedSha256'),('inputTemplate','inputTemplateSha256')]:
        assert hashlib.sha256((root/provenance[key]).read_bytes()).hexdigest() == provenance[hash_key]
    print('PASS: newer uploaded workbook and derived embedded model match their recorded hashes.')
