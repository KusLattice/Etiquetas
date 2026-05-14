import pdfplumber
import os

pdf_path = r'C:\Users\kus\OneDrive\Escritorio\HUB Claro Maipú\Visio\FiberConnectionDiagram_MAIPU VTR.pdf'
with pdfplumber.open(pdf_path) as pdf:
    page = pdf.pages[1] # Page 2
    words = page.extract_words()
    print("Labels matching ODF or suspicious on Page 2:")
    for w in words:
        txt = w['text'].upper()
        if 'ODF' in txt or 'MAIPU' in txt or 'LOUT' in txt or 'LIN' in txt:
            print(f"  '{w['text']}' @ ({w['x0']:.1f}, {w['top']:.1f})")
