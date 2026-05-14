"""
Script de diagnóstico y extracción para FiberConnectionDiagram_MAIPU_VTR.vsdx
Ejecutar: python run_parser.py <ruta_al_vsdx>
"""
import sys
import os
import io

# Forzar UTF-8 en la terminal (Windows puede usar CP1252 por defecto)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ─── Ruta por defecto (ajusta si es necesario) ───────────────────────────────
DEFAULT_VSDX = r'C:\Users\kus\OneDrive\Escritorio\HUB Claro Maipú\Visio\FiberConnectionDiagram_MAIPU VTR.vsdx'

def main():
    vsdx_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_VSDX

    if not os.path.exists(vsdx_path):
        print(f"[ERROR] Archivo no encontrado: {vsdx_path}")
        print("Uso: python run_parser.py <ruta_al_vsdx>")
        sys.exit(1)

    print("=" * 60)
    print(f"Procesando: {os.path.basename(vsdx_path)}")
    print("=" * 60)

    from core_parser import parse_vsdx_connections, generate_excel

    connections = parse_vsdx_connections(vsdx_path, telco_name="ClaroVTR")

    print("\n" + "=" * 60)
    print(f"PRIMERAS 20 CONEXIONES EXTRAÍDAS:")
    print("=" * 60)
    for i, c in enumerate(connections[:20], 1):
        print(f"  {i:02d}. {c['FROM:']:45s} → {c['TO:']}")

    # ─── Validación contra referencias esperadas ─────────────────────────────
    EXPECTED = [
        ("F:RD_MAIP_VTR_1B(04)-WSMD9-OUT",    "T:RD_MAIP_VTR_1B(01)-DAP-VI_2"),
        ("F:RD_MAIP_VTR_1B(01)-DAP-OUT_2",    "T:RD_MAIP_VTR_1B(02)-DWSS20-IN(P)"),
        ("F:RD_MAIP_VTR_1B(01)-DAP-VO_1",     "T:RD_MAIP_VTR_1B(01)-DAP-IN_1"),
        ("F:RD_MAIP_VTR_1B(12)-OH9S-TX2",     "T:RD_MAIP_VTR_1B(12)-OH9S-RM1"),
        ("F:ODF_MAIPU_VTR_B",                  "T:RD_MAIP_VTR_1A(12)-OH9S-LIN"),
    ]

    found_set = {(c['FROM:'], c['TO:']) for c in connections}

    print("\n" + "=" * 60)
    print("VALIDACIÓN DE CONEXIONES DE REFERENCIA:")
    print("=" * 60)
    hits = 0
    for exp_f, exp_t in EXPECTED:
        found = (exp_f, exp_t) in found_set
        status = "[OK]" if found else "[--]"
        print(f"  {status} {exp_f} → {exp_t}")
        if found:
            hits += 1

    print(f"\n  Matches: {hits}/{len(EXPECTED)}")

    # ─── Generar Excel ───────────────────────────────────────────────────────
    output_dir  = os.path.dirname(vsdx_path)
    output_file = os.path.join(output_dir, "MAIPU_VTR_ETIQUETAS.xlsx")
    count = generate_excel(connections, output_file)

    print("\n" + "=" * 60)
    print(f"RESUMEN FINAL:")
    print(f"  Total conexiones únicas : {count}")
    print(f"  Matches con referencia   : {hits}/{len(EXPECTED)}")
    print(f"  Excel generado en        : {output_file}")
    print("=" * 60)


if __name__ == "__main__":
    main()
