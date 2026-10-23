from pathlib import Path
from pypdf import PdfReader
import pypdfium2 as pdfium

root = Path(__file__).resolve().parents[1]
for path in sorted((root / 'tests/plot-output').glob('*.pdf')):
    reader = PdfReader(path)
    assert len(reader.pages) == 1, f'{path.name}: unexpected page count'
    page = reader.pages[0]
    width = float(page.mediabox.width) * 25.4 / 72
    height = float(page.mediabox.height) * 25.4 / 72
    print(path.name, 'page size mm:', round(width,1), round(height,1), 'rotation:', page.get('/Rotate',0))
    pdf = pdfium.PdfDocument(str(path))
    pdf[0].render(scale=1.3).to_pil().save(path.with_suffix('.png'))
    text = page.extract_text() or ''
    path.with_suffix('.txt').write_text(text,encoding='utf-8')
    print('Text characters:',len(text))
