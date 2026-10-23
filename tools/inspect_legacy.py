"""Read-only extraction of the supplied VBA projects and workbook schema."""
import hashlib
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '.tools/python'))
from oletools.olevba import VBA_Parser
import xlrd

out = ROOT / 'migration'
out.mkdir(exist_ok=True)
inventory = []
for path in sorted((ROOT / 'Referrings').rglob('*')):
    if not path.is_file():
        continue
    inventory.append({'path': path.relative_to(ROOT).as_posix(), 'bytes': path.stat().st_size,
                      'sha256': hashlib.sha256(path.read_bytes()).hexdigest()})
    if path.suffix.lower() not in ('.xls', '.dvb'):
        continue
    parser = VBA_Parser(str(path))
    dest = out / 'legacy' / path.stem
    dest.mkdir(parents=True, exist_ok=True)
    count = 0
    for _, stream, name, code in parser.extract_macros():
        if isinstance(code, bytes):
            code = code.decode('cp1252', errors='replace')
        (dest / name).write_text(code.replace('\r\n', '\n').replace('\r', '\n'), encoding='utf-8', newline='\n')
        count += 1
    parser.close()
    print(path.name, 'modules:', count)
(out / 'source-manifest.json').write_text(json.dumps(inventory, indent=2), encoding='utf-8')
book = xlrd.open_workbook(str(ROOT / 'Referrings/Ztools/ZTools.xls'), formatting_info=True)
schema = []
for sheet in book.sheets():
    rows = [[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in range(min(sheet.nrows, 12))]
    schema.append({'name': sheet.name, 'rows': sheet.nrows, 'columns': sheet.ncols,
                   'visibility': sheet.visibility, 'sample': rows})
(out / 'workbook-schema.json').write_text(json.dumps(schema, ensure_ascii=False, indent=2), encoding='utf-8')
print('Sheets:', [(s.name, s.nrows, s.ncols) for s in book.sheets()])
