"""
run_parser.py — Runner de diagnóstico y validación.
Uso: python -X utf8 run_parser.py [ruta_al_vsdx]
"""
import sys, os

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

VSDX_DEFAULT = r'C:\Users\kus\OneDrive\Escritorio\HUB Claro Maipú\Visio\FiberConnectionDiagram_MAIPU VTR.vsdx'
OUTPUT_PATH  = r'C:\Users\kus\OneDrive\Escritorio\HUB Claro Maipú\Visio\MAIPU_VTR_ETIQUETAS.xlsx'

REFERENCE = [
    ("F:RD_MAIP_VTR_1B(04)-WSMD9-OUT",  "T:RD_MAIP_VTR_1B(01)-DAP-VI_2"),
    ("F:RD_MAIP_VTR_1B(12)-OH9S-TX2",   "T:RD_MAIP_VTR_1B(12)-OH9S-RM1"),
    ("F:RD_MAIP_VTR_1A(12)-OH9S-LOUT",  "T:ODF_MAIPU_VTR_B"),
    ("F:RD_MAIP_VTR_1B(01)-DAP-OUT_2",  "T:RD_MAIP_VTR_1B(02)-DWSS20-IN(P)"),
    ("F:RD_MAIP_VTR_1B(01)-DAP-VO_1",   "T:RD_MAIP_VTR_1B(01)-DAP-IN_1"),
    ("F:RD_MAIP_VTR_1B(12)-OH9S-TM2",   "T:RD_MAIP_VTR_1B(12)-OH9S-RX1"),
    ("F:RD_MAIP_VTR_1A(12)-OH9S-TX1",   "T:RD_MAIP_VTR_1A(12)-OH9S-RM2"),
    ("F:ODF_MAIPU_VTR_B",                "T:RD_MAIP_VTR_1A(12)-OH9S-LIN"),
]


def main():
    vsdx = sys.argv[1] if len(sys.argv) > 1 else VSDX_DEFAULT
    out  = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_PATH

    if not os.path.exists(vsdx):
        print(f"[ERROR] Archivo no encontrado: {vsdx}")
        sys.exit(1)

    print("=" * 70)
    print(f"  DWDM Parser — {os.path.basename(vsdx)}")
    print("=" * 70)

    from core_parser import parse_vsdx_connections, generate_excel

    # ── Intento con threshold 0.5 ──────────────────────────────────────────
    conns = parse_vsdx_connections(vsdx, label_threshold=0.5)

    # Si da 0 o muy pocas, escalar threshold automáticamente
    if len(conns) < 10:
        print("\n[RETRY] Menos de 10 conexiones con threshold=0.5. Intentando 1.0...")
        conns = parse_vsdx_connections(vsdx, label_threshold=1.0)

    if len(conns) < 10:
        print("\n[RETRY] Menos de 10 conexiones con threshold=1.0. Intentando 2.0...")
        conns = parse_vsdx_connections(vsdx, label_threshold=2.0)

    # ── Excel ──────────────────────────────────────────────────────────────
    count = generate_excel(conns, out)

    # ── Validación de referencia ───────────────────────────────────────────
    found_set = {(c['FROM:'], c['TO:']) for c in conns}

    print("\n" + "=" * 70)
    print("  VALIDACION DE CONEXIONES DE REFERENCIA")
    print("=" * 70)
    hits = 0
    for f_ref, t_ref in REFERENCE:
        ok = (f_ref, t_ref) in found_set
        status = "[OK]" if ok else "[--]"
        print(f"  {status} {f_ref}  ->  {t_ref}")
        if ok:
            hits += 1

    print(f"\n  Matches: {hits}/{len(REFERENCE)}")

    # Si 0 hits, mostrar lo que SÍ generamos para los boards de referencia
    if hits == 0:
        print("\n[DELTA] Endpoints generados para boards de referencia (WSMD9, DAP, OH9S):")
        keywords = ['WSMD9', 'DAP', 'OH9S', 'DWSS20', 'ODF']
        shown = set()
        for c in conns:
            for kw in keywords:
                key = (c['FROM:'], c['TO:'])
                if kw in c['FROM:'] or kw in c['TO:']:
                    if key not in shown:
                        shown.add(key)
                        print(f"  {c['FROM:']:50s} | {c['TO:']}")

    print("\n" + "=" * 70)
    print(f"  RESUMEN: {count} conexiones unicas en {out}")
    print("=" * 70)


if __name__ == '__main__':
    main()
