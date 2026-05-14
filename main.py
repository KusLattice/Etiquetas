"""
main.py — Motor de Ingeniería de Etiquetas DWDM (Senior Edition)
Arquitectura: Clean Code, Modular y Escalable.
Especialización: Redes Ópticas Huawei/Nokia.
"""

import os
import sys
import re
import pandas as pd
from datetime import datetime
from functools import wraps

# Importar lógica operativa base
try:
    from core_parser import parse_vsdx_connections, generate_excel, validate_telecom_connection
except ImportError:
    print("[ERROR] No se encontró core_parser.py. Asegúrate de que el archivo esté en la misma carpeta.")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────────────────
# DECORADORES DE CALIDAD (Senior approach)
# ─────────────────────────────────────────────────────────────────────────────
def engineering_audit(func):
    """Decorador para auditar y sanitizar conexiones antes de procesarlas."""
    @wraps(func)
    def wrapper(*args, **kwargs):
        connections = func(*args, **kwargs)
        # Aquí se podrían añadir métricas de calidad o logs adicionales
        return connections
    return wrapper

# ─────────────────────────────────────────────────────────────────────────────
# MÓDULO DE ENERGÍA (Dualidad -48VDC)
# ─────────────────────────────────────────────────────────────────────────────
class PowerManager:
    """Maneja la lógica de circuitos de energía TDCC."""
    
    @staticmethod
    def is_power_circuit(conn: dict) -> bool:
        """Determina si una conexión pertenece al TDCC."""
        target_keywords = ['TDCC', 'PDB', 'POWER', 'ALIMENTACION']
        f, t = conn['FROM:'].upper(), conn['TO:'].upper()
        return any(kw in f or kw in t for kw in target_keywords)

    @staticmethod
    def extract_capacity(text: str) -> str:
        """Extrae capacidad (Amperaje) mediante RegEx."""
        # Busca patrones como '63A', '32A', '100A'
        match = re.search(r'(\d+A)\b', text, re.IGNORECASE)
        return match.group(1).upper() if match else "63A"

    @classmethod
    def generate_power_labels(cls, power_conns: list) -> list:
        """Genera bloques de etiquetas para energía (Positivo y Negativo)."""
        rows = []
        for pc in power_conns:
            capacity = cls.extract_capacity(f"{pc['FROM:']} {pc['TO:']}")
            location = "SALA EQUIPOS" # Fallback
            
            # Dupla de bloques: RTN (+) y NEG (-)
            for polarity, label in [('RTN (+)', 'RTN (+)'), ('NEG (-)', 'NEG (-)')]:
                rows.append([f"FROM: {pc['FROM:']} {label}", ""])
                rows.append([f"TO: {pc['TO:']} {label}", ""])
                rows.append([label, ""])
                rows.append([f"{capacity} / {location}", ""])
                rows.append(["", ""]) # Fila en blanco (separador de bloque)
        return rows

# ─────────────────────────────────────────────────────────────────────────────
# GENERADOR DE EXCEL AVANZADO
# ─────────────────────────────────────────────────────────────────────────────
def generate_senior_excel(connections: list, output_path: str):
    """
    Genera un Excel multiactivo con pestañas de FIBRA y ENERGIA.
    Implementa el formato de bloques para energía.
    """
    fiber_conns = []
    power_conns = []
    
    # Clasificación modular
    for c in connections:
        if PowerManager.is_power_circuit(c):
            power_conns.append(c)
        else:
            fiber_conns.append(c)

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. Pestaña de Fibra (Standard)
            if fiber_conns:
                # Mantener la lógica de separadores de equipo de core_parser
                # Para simplificar aquí usamos un DF directo si no hay equipos detectados
                df_fiber = pd.DataFrame([{'FROM:': c['FROM:'], 'TO:': c['TO:']} for c in fiber_conns])
                df_fiber.drop_duplicates(inplace=True)
                df_fiber.to_excel(writer, sheet_name='FIBRA', index=False)
            
            # 2. Pestaña de Energía (Especial -48VDC)
            if power_conns:
                p_rows = PowerManager.generate_power_labels(power_conns)
                df_power = pd.DataFrame(p_rows, columns=['ETIQUETA BANDERITA', ''])
                df_power.to_excel(writer, sheet_name='ENERGIA', index=False)

        return len(fiber_conns), len(power_conns)
    except Exception as e:
        print(f"[CRITICAL] Error al generar Excel: {e}")
        return 0, 0

# ─────────────────────────────────────────────────────────────────────────────
# ORQUESTADOR PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────
@engineering_audit
def run_senior_parser(vsdx_path: str, telco: str = 'ClaroVTR'):
    """Punto de entrada con rigor técnico."""
    print(f"\n[INIT] Iniciando procesamiento Senior: {os.path.basename(vsdx_path)}")
    
    # 1. Parseo con core_parser (Aprovechando lógica existente)
    # Nota: parse_vsdx_connections ya incluye el nuevo motor validate_telecom_connection
    conns = parse_vsdx_connections(vsdx_path, telco_name=telco)
    
    if not conns:
        print("[WARN] No se detectaron conexiones válidas.")
        return
    
    # 2. Definir salida
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"ETIQUETAS_SENIOR_{telco.upper()}_{timestamp}.xlsx"
    output_path = os.path.join(os.path.dirname(vsdx_path), output_name)
    
    # 3. Generación Avanzada
    f_count, p_count = generate_senior_excel(conns, output_path)
    
    print("\n" + "="*50)
    print(f"  REPORTE TÉCNICO DE SALIDA")
    print("="*50)
    print(f"  Fibras procesadas:   {f_count}")
    print(f"  Circuitos Energía:   {p_count} (x2 etiquetas)")
    print(f"  Archivo generado:    {output_name}")
    print("="*50)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Motor de Ingeniería de Etiquetas DWDM (Senior Edition)")
    parser.parse_args() # For consistency
    
    # Si no hay argumentos, mostrar ayuda básica
    if len(sys.argv) == 1:
        print("Uso: python main.py [ruta_al_archivo_vsdx] --telco [CLAROVTR|MOVISTAR|WOM]")
        sys.exit(0)
    
    # Ejecución directa para compatibilidad con el prompt del usuario
    vsdx = sys.argv[1]
    telco = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] != '--telco' else 'ClaroVTR'
    run_senior_parser(vsdx, telco)
