"""Check saved mobile XLS VBA/OLE streams against the untouched source template."""
import hashlib
import sys
from pathlib import Path

root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(root / '.tools/python'))
import olefile

source = root/'Referrings/Ztools/ZTools.xls'
saved = Path((root/'tests/offline-output/last-workbook.txt').read_text(encoding='utf-8-sig'))
with olefile.OleFileIO(source) as before, olefile.OleFileIO(saved) as after:
    preserved = 0
    for stream in before.listdir():
        # BIFF values/formula caches and document summary properties may change;
        # all other original streams (VBA, forms, controls, embedded data) must stay.
        if stream[0] in {'Workbook', '\x05SummaryInformation', '\x05DocumentSummaryInformation'}:
            continue
        assert after.exists(stream), f'Missing original OLE stream: {stream}'
        assert hashlib.sha256(before.openstream(stream).read()).digest() == hashlib.sha256(after.openstream(stream).read()).digest(), stream
        preserved += 1
    assert preserved > 0
print(f'PASS {preserved} original VBA/form/OLE streams retained byte for byte after offline edits and generation')
