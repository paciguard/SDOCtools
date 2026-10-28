"""Remove legacy external links from a COPY of the CFL input template.

Use ZIP-part edits because spreadsheet import/export does not preserve all legacy
ActiveX/VML controls. No Excel process, macros, or external sources are executed.
Unrelated package entries retain their exact uncompressed bytes.
"""
from pathlib import Path
import hashlib
import re
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'Referrings/CFL资料工具输入模板.xlsx'
OUTPUT = ROOT / 'outputs/2026-10-28-template-cleanup/CFL资料工具输入模板_无外链.xlsx'
EXTERNAL = re.compile(rb'\[[1-9][0-9]*\](?:[^<"\r\n]*!)')


def clean(source=SOURCE, output=OUTPUT):
    if source.resolve() == output.resolve():
        raise ValueError('The original reference must not be overwritten.')
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    with zipfile.ZipFile(source) as archive:
        parts = {item.filename: archive.read(item) for item in archive.infolist()}
        result = dict(parts)
        # The external sources occur in nine obsolete names and three hidden
        # combo-box bindings (each binding has both an OOXML and a VML copy).
        # Keep all cell values/formulas and local listFillRange/FmlaRange values.
        workbook = result['xl/workbook.xml']
        names = re.findall(rb'<definedName\b[^>]*>[^<]*\[[1-9][0-9]*\][^<]*</definedName>', workbook)
        if len(names) != 9:
            raise ValueError('Unexpected template names. Inspect before editing.')
        for name in names:
            workbook = workbook.replace(name, b'')
        workbook, count = re.subn(rb'<externalReferences>.*?</externalReferences>', b'', workbook)
        assert count == 1
        result['xl/workbook.xml'] = workbook
        bindings = 0
        for name in list(result):
            if name.startswith('xl/externalLinks/'):
                del result[name]
            elif name == 'xl/_rels/workbook.xml.rels':
                result[name], count = re.subn(rb'<Relationship\b(?=[^>]*Type="[^"]*/externalLink")[^>]*/>', b'', result[name])
                assert count == 3
            elif name == '[Content_Types].xml':
                result[name], count = re.subn(rb'<Override\b(?=[^>]*PartName="/xl/externalLinks/)[^>]*/>', b'', result[name])
                assert count == 3
            elif name.endswith(('.xml', '.vml')):
                result[name], count1 = re.subn(rb' linkedCell="\[3\]Sheet1!\$C\$1[234]"', b'', result[name])
                result[name], count2 = re.subn(rb'<x:FmlaLink>\[3\]Sheet1!\$C\$1[234]</x:FmlaLink>', b'', result[name])
                bindings += count1 + count2
        assert bindings == 6
        # Fail closed rather than leave an unhandled external formula/connection.
        for name, content in result.items():
            if name.endswith(('.xml', '.rels', '.vml')):
                ET.fromstring(content)
                assert not EXTERNAL.search(content), name
                assert b'TargetMode="External"' not in content, name
                assert b'/externalLink' not in content, name
        assert 'xl/connections.xml' not in result
        allowed = {'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml',
                   'xl/worksheets/sheet1.xml', 'xl/drawings/vmlDrawing1.vml'}
        changed = {name for name in result if result[name] != parts[name]}
        assert changed == allowed, changed
        # The worksheet differs ONLY in the three linkedCell attributes.
        assert result['xl/worksheets/sheet1.xml'] == re.sub(
            rb' linkedCell="\[3\]Sheet1!\$C\$1[234]"', b'', parts['xl/worksheets/sheet1.xml'])
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, 'w') as target:
            for item in archive.infolist():
                if item.filename in result:
                    target.writestr(item, result[item.filename])
    with zipfile.ZipFile(output) as saved:
        assert saved.testzip() is None
        assert {name: saved.read(name) for name in saved.namelist()} == result
    assert hashlib.sha256(source.read_bytes()).hexdigest() == original_hash
    print('Original SHA256:', original_hash)
    print('Removed: 3 sources, 9 external names, 3 external control bindings (6 representations).')
    print('Verified: no external workbook links/relationships; all cell data, formulas, styles,')
    print('validations, images, controls and other package parts preserved except the listed bindings.')
    print('Saved:', output)


if __name__ == '__main__':
    clean()
