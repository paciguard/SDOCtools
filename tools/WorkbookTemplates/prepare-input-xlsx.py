"""Keep the input sheet intact while removing unused legacy names Excel rejects."""
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

root = Path(__file__).resolve().parents[2]
source = root / 'Referrings/Updates/CFL-input-template-remote-20261027.xlsx'
destination = root / 'Resources/Workbooks/Input.xlsx'
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with ZipFile(source) as original:
    contents = {name: original.read(name) for name in original.namelist()}
workbook = contents['xl/workbook.xml'].decode('utf-8')
tree = ET.fromstring(workbook)
assert [s.attrib['name'] for s in tree.findall('m:sheets/m:sheet', ns)] == ['input']
sheet = ET.fromstring(contents['xl/worksheets/sheet1.xml'])
formulas = '\n'.join(e.text or '' for tag in ('f', 'formula1', 'formula2') for e in sheet.findall('.//m:' + tag, ns))
names = {e.attrib['name']: e.text or '' for e in tree.findall('m:definedNames/m:definedName', ns)}
used = {name for name in names if name.startswith('_xlnm.Print_')}
while True:
    text = formulas + '\n' + '\n'.join(names[name] for name in used)
    found = {name for name in names if re.search(r'(?<![\w.])' + re.escape(name) + r'(?![\w.])', text, re.IGNORECASE)} | used
    if found == used:
        break
    used = found
removed = sorted(set(names) - used)

def keep(match):
    name = ET.fromstring(match[0]).attrib['name']
    return match[0] if name in used else ''

cleaned = re.sub(r'<definedName\b[^>]*>.*?</definedName>', keep, workbook, flags=re.S)
cleaned = cleaned.replace('<definedNames></definedNames>', '')
contents['xl/workbook.xml'] = cleaned.encode('utf-8')
with ZipFile(destination, 'w', ZIP_DEFLATED) as output:
    for name, content in contents.items():
        output.writestr(name, content)
# Every cell, rich-text string, style, validation and other package part is intact.
with ZipFile(source) as original, ZipFile(destination) as embedded:
    for name in original.namelist():
        if name != 'xl/workbook.xml':
            assert original.read(name) == embedded.read(name), name

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

record = dict(source='Referrings/Updates/CFL-input-template-remote-20261027.xlsx', sourceSha256=sha(source),
              embedded='Resources/Workbooks/Input.xlsx', embeddedSha256=sha(destination),
              format='xlsx', sheets=['input'], removedUnusedNames=removed,
              unchanged='All package parts except unused defined names in xl/workbook.xml')
(root / 'migration/input-xlsx-20261027.json').write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
print(f'Prepared embedded XLSX; removed {len(removed)} unused legacy names; input sheet/styles/dropdowns unchanged.')
