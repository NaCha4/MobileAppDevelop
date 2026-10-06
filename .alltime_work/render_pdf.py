from pathlib import Path
import pypdfium2 as pdfium
from pypdf import PdfReader
root=Path(__file__).resolve().parent/'render'
pdf=root/'alltime.pdf'
r=PdfReader(pdf)
for i,p in enumerate(r.pages):
    text=p.extract_text() or ''
    print(f'{i+1}: {len(text)} chars | {text[:75].replace(chr(10)," ")} | END {text[-80:].replace(chr(10)," ")}')
d=pdfium.PdfDocument(str(pdf))
for i in range(len(d)):
    d[i].render(scale=1.5).to_pil().save(root/f'page-{i+1}.png')
