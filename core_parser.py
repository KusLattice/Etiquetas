import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os
import math

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN Y CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────

NS = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}

# Palabras clave para identificar Puertos (prioridad alta)
PORT_KEYWORDS = ["IN", "OUT", "MON", "VI", "VO", "TX", "RX", "SIG", "EXP", "COM", "LINE"]

# Excepciones de hardware (Boards)
BOARD_EXCEPTIONS = ["ODF", "DDF", "PATCH PANEL", "CABLE", "FIBRA", "SUB-RACK", "OPM", "WSMD", "D40", "M40", "SLA", "OAU"]

# 1. PURIFICACIÓN DE DATOS (BLACKLIST ADMINISTRATIVA)
ADMIN_BLACKLIST = ['CUSTOMER', 'PREPARED BY', 'MODIFIED BY', 'HUAWEI', 'CLARO 2025', 'VERSION']

# VALIDACIÓN DE HARDWARE REAL
HW_WHITELIST = ['OSN', 'DAP', 'WSMD', 'DWSS', 'M808', 'OPM', 'OH9']

# 2. INTELIGENCIA DE TERRENO (WHITELIST DE PUERTOS FÍSICOS)
PORT_WHITELIST = ['VO', 'VI', 'IN', 'OUT', 'EXP', 'DM', 'AM', 'RM', 'TM']

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS DE LIMPIEZA
# ─────────────────────────────────────────────────────────────────────────────

def _clean_string(text):
    """Limpia ruidos de Visio, como espacios entre cada letra."""
    if not text: return ""
    if re.search(r'^[A-Z0-9](\s[A-Z0-9])+$', text):
        text = text.replace(" ", "")
    text = re.sub(r'([A-Z0-9])\s(?=[A-Z0-9]\s|[A-Z0-9]$)', r'\1', text)
    return text.strip()

def _cell(shape, name, default=None):
    c = shape.find(f'./v:Cell[@N="{name}"]', NS)
    if c is not None:
        try: return float(c.attrib.get('V', default or 0.0))
        except: return default or 0.0
    return default

def _get_text(shape):
    t_elem = shape.find('./v:Text', NS)
    if t_elem is not None:
        text = "".join(t_elem.itertext()).strip()
        return _clean_string(text)
    return ""

def is_annotation(text):
    if not text: return True
    t = text.upper()
    if len(text) > 50: return True
    if any(x in t for x in ["NOTA", "VER ", "REV", "PROYECTO", "DWDM", "FIBRA"]): return True
    # Aplicar Blacklist administrativa
    if any(x in t for x in ADMIN_BLACKLIST): return True
    return False

def is_valid_hardware(text):
    t = text.upper()
    return any(hw in t for hw in HW_WHITELIST) or any(exc in t for exc in BOARD_EXCEPTIONS)

# ─────────────────────────────────────────────────────────────────────────────
# LÓGICA DE RESOLUCIÓN ESPACIAL
# ─────────────────────────────────────────────────────────────────────────────

def resolve_endpoint(point, text_items):
    px, py = point
    candidates = []
    
    for item in text_items:
        tx, ty, tw, th, ttext = item
        if is_annotation(ttext): continue # Ignora metadatos
        
        dist = math.sqrt((px - tx)**2 + (py - ty)**2)
        if dist < 3.0: 
            candidates.append({'dist': dist, 'text': ttext, 'item': item})
            
    if not candidates:
        return "Unknown", "Unknown", None

    candidates.sort(key=lambda x: x['dist'])
    
    best_port = "Unknown"
    port_item = None
    
    for c in candidates:
        txt = c['text'].upper()
        if any(kw in txt for kw in PORT_KEYWORDS) and len(txt) < 15:
            best_port = c['text']
            port_item = c['item']
            break
            
    if best_port == "Unknown":
        for c in candidates:
            if not is_annotation(c['text']) and len(c['text']) < 10:
                best_port = c['text']
                port_item = c['item']
                break

    best_board = "Unknown"
    board_item = None
    
    for c in candidates:
        if c['item'] == port_item: continue
        txt = c['text'].upper()
        
        if is_valid_hardware(txt) or re.search(r'\(\d+_[A-Z]\d+\)', txt):
            best_board = c['text']
            board_item = c['item']
            break
            
    if best_board == "Unknown":
        for c in candidates:
            if c['item'] == port_item: continue
            if not is_annotation(c['text']):
                best_board = c['text']
                board_item = c['item']
                break
                
    return best_board, best_port, (board_item, port_item)

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

# ─────────────────────────────────────────────────────────────────────────────
# PARSER VSDX
# ─────────────────────────────────────────────────────────────────────────────

def parse_vsdx_connections(filepath, telco_name="ClaroVTR", project_name="Core"):
    fiber_conns = []
    power_conns = []
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            pages = [f for f in z.namelist() if f.startswith('visio/pages/page')]
            for page_path in sorted(pages):
                root = ET.fromstring(z.read(page_path))
                
                text_items = []
                lines = []
                
                for s in root.findall('.//v:Shape', NS):
                    txt = _get_text(s)
                    x = _cell(s, 'PinX')
                    y = _cell(s, 'PinY')
                    w = _cell(s, 'Width', 0.1)
                    h = _cell(s, 'Height', 0.1)
                    
                    if x is not None and y is not None:
                        if txt:
                            text_items.append((x, y, w, h, txt))
                        if (w < 0.05 or h < 0.05) and (w > 0 or h > 0):
                            lines.append(((x - w/2, y - h/2), (x + w/2, y + h/2)))
                    
                    bx, by = _cell(s, 'BeginX'), _cell(s, 'BeginY')
                    ex, ey = _cell(s, 'EndX'), _cell(s, 'EndY')
                    if bx is not None and by is not None:
                        if abs(bx-ex) > 0.01 or abs(by-ey) > 0.01:
                            lines.append(((bx, by), (ex, ey)))

                for start_pt, end_pt in lines:
                    b_start, l_start, _ = resolve_endpoint(start_pt, text_items)
                    b_end, l_end, _ = resolve_endpoint(end_pt, text_items)
                    
                    if b_start == "Unknown" and b_end == "Unknown":
                        continue
                        
                    if PowerManager.is_power_circuit(b_start, b_end):
                        power_conns.append({'FROM:': b_start, 'TO:': b_end})
                        continue

                    # Validación de Hardware: Ignorar si no coincide con los patrones reales
                    if not (is_valid_hardware(b_start) or is_valid_hardware(b_end)):
                        continue

                    # Filtro de Terreno (Anti-Loopbacks con Whitelist)
                    if b_start == b_end:
                        # Rack, Slot y Board idénticos. Chequeamos la whitelist.
                        port_kw_a = any(w in l_start.upper() for w in PORT_WHITELIST)
                        port_kw_b = any(w in l_end.upper() for w in PORT_WHITELIST)
                        if not (port_kw_a or port_kw_b):
                            continue # Se descarta el loopback lógico

                    fiber_conns.append({
                        'A_Board': b_start, 'A_Port': l_start,
                        'B_Board': b_end, 'B_Port': l_end
                    })
                    
    except Exception as e:
        print(f"Error parsing VSDX: {e}")
        
    return {'fiber': fiber_conns, 'power': power_conns}

def format_fiber_label(board, port):
    # Formato: F:RD_[Rack]([Slot])-[Board]-[Puerto]
    # Asume que el nombre del board ya contiene parte de la info y lo estructuramos.
    board_clean = board.replace(" ", "")
    port_clean = port.replace(" ", "")
    return f"RD_{board_clean}-{port_clean}"

def generate_excel(data_dict, output_path):
    if not isinstance(data_dict, dict):
        return 0
    
    fiber_conns = data_dict.get('fiber', [])
    power_conns = data_dict.get('power', [])
    
    count = 0
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 3. FORMATO DE SALIDA (EXCEL P-TOUCH)
            if fiber_conns:
                formatted_fibers = []
                for c in fiber_conns:
                    formatted_fibers.append({
                        'FROM:': f"F:{format_fiber_label(c['A_Board'], c['A_Port'])}",
                        'TO:': f"T:{format_fiber_label(c['B_Board'], c['B_Port'])}"
                    })
                df_fiber = pd.DataFrame(formatted_fibers)
                df_fiber.drop_duplicates(inplace=True)
                df_fiber.to_excel(writer, sheet_name='FIBRA', index=False)
                count += len(df_fiber)
            
            if power_conns:
                p_rows = []
                for pc in power_conns:
                    capacity = PowerManager.extract_capacity(f"{pc['FROM:']} {pc['TO:']}")
                    location = "SALA EQUIPOS"
                    
                    # Generar bloques de 4 filas con una fila vacía intermedia
                    for polarity in ['NEG (-)', 'RTN (+)']:
                        p_rows.append([f"FROM: {pc['FROM:']}"])
                        p_rows.append([f"TO: {pc['TO:']}"])
                        p_rows.append([polarity])
                        p_rows.append([f"{capacity} / {location}"])
                        p_rows.append([""]) # Fila vacía para guía de corte físico
                        
                df_power = pd.DataFrame(p_rows, columns=['ETIQUETA ENERGIA'])
                df_power.to_excel(writer, sheet_name='ENERGIA', index=False)
                count += len(power_conns)
    except Exception as e:
        print(f"Error generando Excel: {e}")
        
    return count

# ─────────────────────────────────────────────────────────────────────────────
# COMPATIBILIDAD CON EXTENSIONES (PDF Parser)
# ─────────────────────────────────────────────────────────────────────────────
def normalize_board(text):
    if not text or text == "Unknown": return text
    t = text.upper().strip()
    t = re.sub(r'RACK[- ]?\d+', '', t)
    t = re.sub(r'SLOT[- ]?\d+', '', t)
    return t.strip() or text

def parse_board_shape(shape):
    return {'text': _get_text(shape), 'x': _cell(shape, 'PinX'), 'y': _cell(shape, 'PinY'),
            'w': _cell(shape, 'Width'), 'h': _cell(shape, 'Height')}
def normalize_site(text, ope=""): return text
def build_endpoint(board, port, site="Unknown"):
    return {"Site": site, "Board": board, "Port": port}
def is_odf_label(text): return "ODF" in (text or "").upper()
def extract_odf_name(text, site=""): return text
def detect_site(text): return "Unknown"
def _clean_port(text): return text
