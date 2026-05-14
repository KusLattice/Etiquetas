"""
debug_pdf.py — Análisis forense de estructura PDF para diagramas DWDM Huawei.
Examina texto, líneas y rectángulos con coordenadas.
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import pdfplumber

PDF_PATH = r'C:\Users\kus\OneDrive\Escritorio\HUB Claro Maipú\Visio\FiberConnectionDiagram_MAIPU VTR.pdf'

print("=" * 70)
print(f"  Análisis PDF: {PDF_PATH.split(chr(92))[-1]}")
print("=" * 70)

with pdfplumber.open(PDF_PATH) as pdf:
    print(f"\n[INFO] Total de páginas: {len(pdf.pages)}")

    for page_num, page in enumerate(pdf.pages, 1):
        print(f"\n{'='*60}")
        print(f"  PÁGINA {page_num}  — Size: {page.width:.1f} x {page.height:.1f} pts")
        print(f"{'='*60}")

        # ── Texto con coordenadas ─────────────────────────────────
        words = page.extract_words(
            x_tolerance=2,
            y_tolerance=2,
            keep_blank_chars=False,
            use_text_flow=False
        )
        print(f"\n[TEXTO] {len(words)} palabras encontradas. Primeras 40:")
        for w in words[:40]:
            print(f"  '{w['text']:30s}' x0={w['x0']:6.1f} y0={w['top']:6.1f} x1={w['x1']:6.1f} y1={w['bottom']:6.1f}")

        # ── Líneas ───────────────────────────────────────────────
        lines = page.lines
        print(f"\n[LINEAS] {len(lines)} líneas encontradas. Primeras 20:")
        for ln in lines[:20]:
            print(f"  ({ln['x0']:6.1f},{ln['top']:6.1f}) -> ({ln['x1']:6.1f},{ln['bottom']:6.1f})  linewidth={ln.get('linewidth', '?'):.1f}")

        # ── Rectángulos (pueden ser shapes de boards) ─────────────
        rects = page.rects
        print(f"\n[RECTS] {len(rects)} rectángulos. Primeros 15:")
        for r in rects[:15]:
            w_r = r['x1'] - r['x0']
            h_r = r['bottom'] - r['top']
            print(f"  x0={r['x0']:6.1f} y0={r['top']:6.1f}  W={w_r:.1f} H={h_r:.1f}  fill={r.get('non_stroking_color','none')}")

        # ── Curvas/paths ──────────────────────────────────────────
        curves = page.curves
        print(f"\n[CURVES/PATHS] {len(curves)} paths.")

        # Si hay muchas líneas, mostrar estadísticas de longitud
        if lines:
            horiz = [l for l in lines if abs(l['top'] - l['bottom']) < 1.0]
            vert  = [l for l in lines if abs(l['x0'] - l['x1']) < 1.0]
            print(f"\n[STATS] Horizontales: {len(horiz)}, Verticales: {len(vert)}, Diagonales: {len(lines)-len(horiz)-len(vert)}")

        # Solo página 1 en detalle para no saturar output
        if page_num >= 2:
            print(f"\n  [+] Páginas restantes — solo conteo:")
            for pn, pg in enumerate(pdf.pages[page_num:], page_num + 1):
                ws = pg.extract_words()
                print(f"    Pág {pn}: {len(ws)} palabras, {len(pg.lines)} líneas, {len(pg.rects)} rects")
            break

print("\n[DONE] Análisis completado.")
