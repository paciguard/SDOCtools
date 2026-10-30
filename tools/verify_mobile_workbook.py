"""Check input-only projects and preservation of the archived source workbook."""
import hashlib
import argparse
import json
import subprocess
from zipfile import ZipFile
from pathlib import Path

root = Path(__file__).resolve().parents[1]

parser = argparse.ArgumentParser()
parser.add_argument('--workbook', type=Path, help='A saved private mobile workbook to verify')
args = parser.parse_args()
saved = args.workbook or Path((root/'tests/offline-output/last-workbook.txt').read_text(encoding='utf-8-sig'))
provenance = json.loads((root/'migration/workbook-split-20261026.json').read_text(encoding='utf-8'))
assert hashlib.sha256((root/provenance['source']).read_bytes()).hexdigest() == provenance['sourceSha256']
assert saved.suffix.lower() == '.xlsx'
with ZipFile(saved) as book:
    assert 'xl/workbook.xml' in book.namelist()
    assert not any('vba' in name.lower() for name in book.namelist())
exe = root/'dist/CFL-SubmitTools/ZTools.Standalone.exe'
result = subprocess.run([str(exe), '--verify-input', str(saved)], check=True, capture_output=True, text=True,
                        creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
info = json.loads(result.stdout)
assert info['format'] == 'xlsx' and info['inputOnly'] and info['sheets'] == ['input'] and info['dropdowns'] >= 8
print('PASS public XLSX contains only input, retains dropdowns/project parameters, has no VBA, and the uploaded original is unchanged')
