import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
sys.stdout.reconfigure(encoding='utf-8')
from pdf_parser import parse_pdf_connections
from core_parser import generate_excel

PDF = r'C:\Users\kus\OneDrive\Escritorio\HUB Claro Maipú\Visio\FiberConnectionDiagram_MAIPU VTR.pdf'
OUT = r'C:\Users\kus\OneDrive\Escritorio\HUB Claro Maipú\Visio\MAIPU_VTR_PDF_TEST.xlsx'

REFERENCE = [
    ("F:RD_MAIP_VTR_1B(04)-WSMD9-OUT",  "T:RD_MAIP_VTR_1B(01)-DAP-VI_2"),
    ("F:RD_MAIP_VTR_1B(12)-OH9S-TX2",   "T:RD_MAIP_VTR_1B(12)-OH9S-RM1"),
    ("F:RD_MAIP_VTR_1A(12)-OH9S-LOUT",  "T:ODF_MAIPU_VTR_B"),
    ("F:RD_MAIP_VTR_1B(01)-DAP-OUT_2",  "T:RD_MAIP_VTR_1B(02)-DWSS20-IN(P)"),
    ("F:ODF_MAIPU_VTR_B",                "T:RD_MAIP_VTR_1A(12)-OH9S-LIN"),
]

conns = parse_pdf_connections(PDF)

print(f"\n{'='*60}")
print(f"Primeras 25 conexiones:")
print(f"{'='*60}")
for i, c in enumerate(conns[:25], 1):
    print(f"  {i:02d}. {c['FROM:']:50s} | {c['TO:']}")

print(f"\n{'='*60}")
print("VALIDACION:")
print(f"{'='*60}")
found = {(c['FROM:'], c['TO:']) for c in conns}
hits = 0
for f, t in REFERENCE:
    ok = (f, t) in found
    print(f"  {'[OK]' if ok else '[--]'} {f} -> {t}")
    if ok: hits += 1
print(f"\n  Matches: {hits}/{len(REFERENCE)}")

generate_excel(conns, OUT)
