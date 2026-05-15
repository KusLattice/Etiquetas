import os
import sys
import re
import math
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime

"""
main.py — Motor de Ingeniería de Etiquetas DWDM (Senior Edition)
Arquitectura: Clean Code, Modular y Escalable.
Especialización: Redes Ópticas Huawei/Nokia y Energía TDCC.
"""

NS = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}
BOARD_EXCEPTIONS = ["ODF", "DDF", "PATCH PANEL", "CABLE", "FIBRA", "SUB-RACK", "OPM", "WSMD", "D40", "M40", "SLA", "OAU"]

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS DE GEOMETRÍA Y XML
# ─────────────────────────────────────────────────────────────────────────────

def _cell(shape, name, default=0.0):
    c = shape.find(f'./v:Cell[@N="{name}"]', NS)
    if c is not None:
        try: return float(c.attrib.get('V', default))
        except: return default
    return default

def _get_text(shape):
    t_elem = shape.find('./v:Text', NS)
    if t_elem is not None:
        text = "".join(t_elem.itertext()).strip()
        # Clean extra spaces between letters if they appear
        if re.search(r'^[A-Z0-9](\s[A-Z0-9])+$', text):
            text = text.replace(" ", "")
        text = re.sub(r'([A-Z0-9])\s(?=[A-Z0-9]\s|[A-Z0-9]$)', r'\1', text)
        return text.strip()
    return "PORT_UNK"

def clean_name(text):
    if not text or text == "PORT_UNK": return "PORT_UNK"
    # Eliminar paréntesis residuales
    t = re.sub(r'\(.*?\)', '', text)
    t = re.sub(r'\[.*?\]', '', t)
    return t.strip()

# ─────────────────────────────────────────────────────────────────────────────
# FILTRO DE INGENIERÍA (is_physical_fiber)
# ─────────────────────────────────────────────────────────────────────────────

def clean_telecom_data(df):
    # 1. PURGA RADICAL (Blacklist extendida para nodos virtuales de Visio)
    blacklist = ['customer', 'prepared by', 'modified by', 'huawei', 'claro', 'version', 'reviewed by', 'created on', 'technologies', 'freeodf', 'occupiedodf', 'virtualboard', 'unknown']
    
    for col in ['Equipo_A', 'Equipo_B']:
        if col in df.columns:
            df = df[~df[col].astype(str).str.contains('|'.join(blacklist), case=False, na=False)]

    # 2. NORMALIZACIÓN DE HARDWARE (Limpiar basura como 191.650000THz o paréntesis de Visio)
    def normalize_board(board_str):
        board_str = str(board_str).upper()
        # Mapeo directo para tarjetas conocidas
        if 'DAP' in board_str: return 'DAP'
        if 'WSMD9' in board_str: return 'WSMD9'
        if 'DWSS20' in board_str: return 'DWSS20'
        if 'OPM8' in board_str: return 'OPM8'
        if 'M808' in board_str: return 'M808SA'
        if 'OH9S' in board_str: return 'OH9S'
        
        # Limpieza genérica si es otra tarjeta
        clean = re.sub(r'\(.*?\)', '', board_str) # Elimina datos entre paréntesis
        clean = re.sub(r'\d+\.\d+THZ\+-\d+\.\d+GHZ', '', clean) # Elimina frecuencias ópticas
        clean = re.sub(r'^G\d+', '', clean) # Elimina prefijos como G1, G2
        return clean.strip('- ')

    if 'Board_A' in df.columns and 'Board_B' in df.columns:
        df['Board_A'] = df['Board_A'].apply(normalize_board)
        df['Board_B'] = df['Board_B'].apply(normalize_board)

    # 3. EXCEPCIONES DE TERRENO (Whitelist para loops físicos obligatorios)
    whitelist_ports = ['VO', 'VI', 'IN', 'OUT', 'EXP', 'DM', 'AM', 'RM', 'TM']
    
    def is_valid_connection(row):
        is_loop = (str(row.get('Rack_A', '')) == str(row.get('Rack_B', ''))) and \
                  (str(row.get('Slot_A', '')) == str(row.get('Slot_B', ''))) and \
                  (str(row.get('Board_A', '')) == str(row.get('Board_B', '')))
        
        port_a, port_b = str(row.get('Port_A', '')).upper(), str(row.get('Port_B', '')).upper()
        has_valid_port = any(p in port_a or p in port_b for p in whitelist_ports)
        
        if is_loop and not has_valid_port:
            return False # Descartar: loop lógico de software
        return True # Mantener: conexión física válida
        
    return df[df.apply(is_valid_connection, axis=1)].copy()

# ─────────────────────────────────────────────────────────────────────────────
# MÓDULO DE ENERGÍA Y FORMATEO
# ─────────────────────────────────────────────────────────────────────────────

class PowerManager:
    @staticmethod
    def is_power_circuit(board_a, board_b):
        target_keywords = ['TDCC', 'PDB', 'POWER', 'ALIMENTACION']
        combined = f"{board_a} {board_b}".upper()
        return any(kw in combined for kw in target_keywords)

    @staticmethod
    def extract_capacity(text: str) -> str:
        match = re.search(r'(\d+A)\b', text, re.IGNORECASE)
        return match.group(1).upper() if match else "63A"

def format_fiber_label(board, port):
    cb = clean_name(board)
    cp = clean_name(port)
    return f"RD_{cb}_{cp}"

# ─────────────────────────────────────────────────────────────────────────────
# PARSER VSDX
# ─────────────────────────────────────────────────────────────────────────────

def resolve_endpoint(point, text_items):
    px, py = point
    candidates = []
    
    for tx, ty, tw, th, ttext in text_items:
        dist = math.sqrt((px - tx)**2 + (py - ty)**2)
        if dist < 3.0: 
            candidates.append({'dist': dist, 'text': ttext, 'item': (tx, ty, tw, th, ttext)})
            
    if not candidates:
        return "PORT_UNK", "PORT_UNK", None

    candidates.sort(key=lambda x: x['dist'])
    
    best_port = "PORT_UNK"
    port_item = None
    
    for c in candidates:
        txt = c['text'].upper()
        if any(kw in txt for kw in ["IN", "OUT", "MON", "VI", "VO", "TX", "RX", "SIG", "EXP", "COM", "LINE"]) and len(txt) < 15:
            best_port = c['text']
            port_item = c['item']
            break
            
    if best_port == "PORT_UNK":
        for c in candidates:
            if len(c['text']) < 10:
                best_port = c['text']
                port_item = c['item']
                break

    best_board = "PORT_UNK"
    
    for c in candidates:
        if c['item'] == port_item: continue
        txt = c['text'].upper()
        if any(exc in txt for exc in BOARD_EXCEPTIONS) or re.search(r'\(\d+_[A-Z]\d+\)', txt) or len(txt) > 5:
            best_board = c['text']
            break
                
    return best_board, best_port, port_item

def parse_vsdx(filepath):
    fiber_conns = []
    power_conns = []
    
    print(f"\n[INIT] Iniciando procesamiento Senior: {os.path.basename(filepath)}")

    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            pages = [f for f in z.namelist() if f.startswith('visio/pages/page')]
            for page_path in sorted(pages):
                root = ET.fromstring(z.read(page_path))
                
                text_items = []
                lines = []
                
                for s in root.findall('.//v:Shape', NS):
                    txt = _get_text(s)
                    x = _cell(s, 'PinX', None)
                    y = _cell(s, 'PinY', None)
                    w = _cell(s, 'Width', 0.1)
                    h = _cell(s, 'Height', 0.1)
                    
                    if x is not None and y is not None:
                        if txt and txt != "PORT_UNK":
                            text_items.append((x, y, w, h, txt))
                        if (w < 0.05 or h < 0.05) and (w > 0 or h > 0):
                            lines.append(((x - w/2, y - h/2), (x + w/2, y + h/2)))
                    
                    bx, by = _cell(s, 'BeginX', None), _cell(s, 'BeginY', None)
                    ex, ey = _cell(s, 'EndX', None), _cell(s, 'EndY', None)
                    if bx is not None and by is not None:
                        if abs(bx-ex) > 0.01 or abs(by-ey) > 0.01:
                            lines.append(((bx, by), (ex, ey)))

                for start_pt, end_pt in lines:
                    b_start, l_start, _ = resolve_endpoint(start_pt, text_items)
                    b_end, l_end, _ = resolve_endpoint(end_pt, text_items)
                    
                    from_label = format_fiber_label(b_start, l_start)
                    to_label = format_fiber_label(b_end, l_end)
                    
                    print(f"Procesando: {from_label} -> {to_label}")
                    
                    if PowerManager.is_power_circuit(b_start, b_end):
                        power_conns.append({'FROM:': b_start, 'TO:': b_end})
                        continue

                    # Guardar raw data para pandas
                    fiber_conns.append({
                        'Equipo_A': b_start, 'Rack_A': '', 'Slot_A': '', 'Board_A': b_start, 'Port_A': l_start,
                        'Equipo_B': b_end, 'Rack_B': '', 'Slot_B': '', 'Board_B': b_end, 'Port_B': l_end,
                        'FROM:': from_label,
                        'TO:': to_label
                    })

    except Exception as e:
        print(f"[CRITICAL] Error parsing VSDX: {e}")

    # Filtrado estricto con Pandas
    if fiber_conns:
        df_fibers = pd.DataFrame(fiber_conns)
        df_clean = clean_telecom_data(df_fibers)
    else:
        df_clean = pd.DataFrame()
        
    return df_clean, power_conns

# ─────────────────────────────────────────────────────────────────────────────
# GENERADOR DE EXCEL
# ─────────────────────────────────────────────────────────────────────────────

class CircuitoTDCC:
    def __init__(self, origen, destino, capacidad, ubicacion):
        self.origen = origen
        self.destino = destino
        self.capacidad = capacidad
        self.ubicacion = ubicacion

def generate_senior_excel(df, power_conns, output_path):
    f_count = 0
    p_count = 0
    
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. Pestaña de Fibra
            if not df.empty:
                # Generar dataframe final limpio
                df_fibra = pd.DataFrame({
                    'FROM:': 'F:RD_' + df['Rack_A'].astype(str) + '(' + df['Slot_A'].astype(str) + ')-' + df['Board_A'].astype(str) + '-' + df['Port_A'].astype(str),
                    'TO:': 'T:RD_' + df['Rack_B'].astype(str) + '(' + df['Slot_B'].astype(str) + ')-' + df['Board_B'].astype(str) + '-' + df['Port_B'].astype(str)
                })
                df_fibra.drop_duplicates(inplace=True)
                df_fibra.to_excel(writer, sheet_name='FIBRA', index=False)
                f_count = len(df_fibra)
            
            # 2. Pestaña de Energía
            if power_conns:
                circuitos_tdcc = []
                for pc in power_conns:
                    cap = PowerManager.extract_capacity(f"{pc['FROM:']} {pc['TO:']}")
                    circuitos_tdcc.append(CircuitoTDCC(pc['FROM:'], pc['TO:'], cap, "SALA EQUIPOS"))

                ws_energia = []
                
                # Iterar sobre circuitos para generar bloques de 4 filas + 1 de corte
                for circ in circuitos_tdcc:
                    # Bloque NEGATIVO
                    ws_energia.append([f"FROM: {circ.origen} NEG (-)"])
                    ws_energia.append([f"TO: {circ.destino} NEG (-)"])
                    ws_energia.append(["NEG (-)"])
                    ws_energia.append([f"{circ.capacidad} / {circ.ubicacion}"])
                    ws_energia.append([""]) # Guía de corte
                    
                    # Bloque RETORNO
                    ws_energia.append([f"FROM: {circ.origen} RTN (+)"])
                    ws_energia.append([f"TO: {circ.destino} RTN (+)"])
                    ws_energia.append(["RTN (+)"])
                    ws_energia.append([f"{circ.capacidad} / {circ.ubicacion}"])
                    ws_energia.append([""]) # Guía de corte
                        
                df_power = pd.DataFrame(ws_energia, columns=['ETIQUETA BANDERITA'])
                df_power.to_excel(writer, sheet_name='ENERGIA', index=False)
                p_count = len(power_conns)

        return f_count, p_count
    except Exception as e:
        print(f"[CRITICAL] Error al generar Excel: {e}")
        return 0, 0

def run_senior_parser(vsdx_path: str):
    df_clean, power_conns = parse_vsdx(vsdx_path)
    
    if df_clean.empty and not power_conns:
        print("[WARN] No se detectaron conexiones válidas.")
        return
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"ETIQUETAS_SENIOR_OUTPUT_{timestamp}.xlsx"
    output_path = os.path.join(os.path.dirname(vsdx_path), output_name)
    
    f_count, p_count = generate_senior_excel(df_clean, power_conns, output_path)
    
    print("\n" + "="*50)
    print(f"  REPORTE TÉCNICO DE SALIDA")
    print("="*50)
    print(f"  Fibras procesadas:   {f_count}")
    print(f"  Circuitos Energía:   {p_count} (x2 etiquetas)")
    print(f"  Archivo generado:    {output_name}")
    print("="*50)

if __name__ == "__main__":
    if len(sys.argv) == 1:
        print("Uso: python main.py [ruta_al_archivo_vsdx]")
        sys.exit(0)
    
    vsdx = sys.argv[1]
    run_senior_parser(vsdx)
