"""
pdf_parser.py — Parser DWDM Huawei para diagramas en PDF vectorial.
Usa pdfplumber para extraer texto con coordenadas y líneas geométricas.
Aplica la misma lógica de normalización que core_parser.py.

Arquitectura PDF vs VSDX:
  - Boards: rectángulos verticales (W~10 pts, H>>W, fill=1.0) con texto superpuesto
  - Labels: palabras sueltas cerca de los extremos de línea
  - Lines: segmentos linewidth~0.7 (las fibras ópticas del diagrama)
  - Coordenadas en puntos PDF (pts), origen en esquina superior izquierda
"""

import pdfplumber
import pandas as pd
import re
import os

# ─────────────────────────────────────────────────────────────────────────────
# Importar normalización desde core_parser (reutilizar 100%)
# ─────────────────────────────────────────────────────────────────────────────
from core_parser import (
    normalize_board,
    normalize_site,
    parse_board_shape,
    build_endpoint,
    is_annotation,
    is_odf_label,
    extract_odf_name,
    detect_site,
    _clean_port,
    BOARD_EXCEPTIONS,
)

# ─────────────────────────────────────────────────────────────────────────────
# Constantes PDF
# ─────────────────────────────────────────────────────────────────────────────
# Los boards en PDF son rectángulos con fill=1.0 y relación H/W > 5
BOARD_MIN_H_W_RATIO = 5.0
BOARD_MIN_H_PTS     = 30.0   # mínimo de altura para ser board
BOARD_MAX_W_PTS     = 25.0   # máximo de ancho para ser board

# Las fibras son líneas con linewidth ~ 0.5-1.0
FIBER_MAX_LINEWIDTH = 2.0

# Thresholds de proximidad en pts (1 inch = 72 pts)
LABEL_THRESHOLD_PTS  = 20.0   # radio para buscar labels en extremos de línea
ODF_THRESHOLD_PTS    = 70.0   # radio ampliado para labels ODF

# ─────────────────────────────────────────────────────────────────────────────
# Agrupación de palabras PDF en "shapes" (reconstruir texto de board)
# ─────────────────────────────────────────────────────────────────────────────
def cluster_words_to_board(words, rect):
    """
    Agrupa palabras dentro del rect para encontrar el ID del board.
    Capura tanto el nombre de placa como el slot info (1_B04).
    """
    x0, y0, x1, y1 = rect['x0'], rect['top'], rect['x1'], rect['bottom']
    
    # Capturar palabras que estén dentro o muy cerca horizontalmente del rect
    inside = [
        w for w in words
        if (x0 - 5.0) <= w['x0'] <= (x1 + 5.0) 
        and (y0 - 2.0) <= w['top'] <= (y1 + 2.0)
    ]
    
    if not inside:
        return ''
        
    # Ordenar por Y para que el nombre de placa (arriba) venga antes que el slot (abajo)
    inside.sort(key=lambda w: (w['top'], w['x0']))
    
    # Filtrar ruidos: ignorar labels de puerto muy comunes si hay otras palabras
    ports = {'IN', 'OUT', 'VO', 'VI', 'MON', 'IN1', 'IN2', 'OUT1', 'OUT2'}
    filtered = [w['text'] for w in inside if w['text'].upper() not in ports or len(inside) < 3]
    
    return "".join(filtered)


# ─────────────────────────────────────────────────────────────────────────────
# Detección de boards desde rectángulos PDF
# ─────────────────────────────────────────────────────────────────────────────
def detect_boards_from_rects(page, words):
    """
    Identifica boards como rectángulos verticales (H >> W, fill=1.0).
    Retorna lista de dicts con coordenadas centrales y texto colapsado.
    """
    boards = []
    seen_positions = set()

    for rect in page.rects:
        w = rect['x1'] - rect['x0']
        h = rect['bottom'] - rect['top']

        # Filtros geométricos
        if h < BOARD_MIN_H_PTS:
            continue
        if w > BOARD_MAX_W_PTS:
            continue
        if w <= 0 or (h / w) < BOARD_MIN_H_W_RATIO:
            continue
        if rect.get('non_stroking_color') == 0:  # color negro = borde exterior
            continue

        # Deduplicar (Visio exporta rects duplicados en PDF)
        pos_key = (round(rect['x0'], 1), round(rect['top'], 1))
        if pos_key in seen_positions:
            continue
        seen_positions.add(pos_key)

        # Centro del rectángulo
        cx = (rect['x0'] + rect['x1']) / 2
        cy = (rect['top'] + rect['bottom']) / 2

        # Texto del board (buscar en zona adyacente al rect)
        text = cluster_words_to_board(words, rect)

        boards.append({
            'x': cx, 'y': cy,
            'x0': rect['x0'], 'y0': rect['top'],
            'x1': rect['x1'], 'y1': rect['bottom'],
            'w': w, 'h': h,
            'text': text,
            'page': page.page_number
        })

    return boards


# ─────────────────────────────────────────────────────────────────────────────
# Extracción de labels (palabras que no pertenecen a boards)
# ─────────────────────────────────────────────────────────────────────────────
def extract_labels(page, words, boards):
    """
    Retorna palabras agrupadas en clusters lógicos que no están dentro de boards.
    Usa una distancia mínima para unir fragmentos como 'OUT' + '1' o 'ODF' + 'MAIPU'.
    """
    # 1. Filtrar palabras que están dentro de boards
    non_board_words = []
    for w in words:
        wx = (w['x0'] + w['x1']) / 2
        wy = (w['top'] + w['bottom']) / 2
        in_board = any(
            (b['x0'] - 2) <= wx <= (b['x1'] + 2) and
            (b['y0'] - 2) <= wy <= (b['y1'] + 2)
            for b in boards
        )
        if not in_board:
            non_board_words.append(w)
            
    if not non_board_words:
        return []

    # 2. Agrupar palabras cercanas (clustering simple por proximidad)
    # Ordenar por Y y luego X
    non_board_words.sort(key=lambda w: (w['top'], w['x0']))
    
    clusters = []
    for w in non_board_words:
        assigned = False
        for c in clusters:
            # Si está en la misma línea (Y similar) y cerca en X
            if abs(c['y'] - (w['top'] + w['bottom'])/2) < 3.0:
                if (w['x0'] - c['x1']) < 8.0: # margen para unir 'ODF' + 'MAIPU'
                    c['text'] += "_" + w['text'] if "_" not in w['text'] and "_" not in c['text'][-1:] else w['text']
                    c['x1'] = w['x1']
                    c['x'] = (c['x0'] + c['x1']) / 2
                    assigned = True
                    break
        if not assigned:
            clusters.append({
                'x0': w['x0'], 'x1': w['x1'],
                'x': (w['x0'] + w['x1']) / 2,
                'y': (w['top'] + w['bottom']) / 2,
                'text': w['text'],
                'page': page.page_number
            })
            
    return clusters


# ─────────────────────────────────────────────────────────────────────────────
# Extracción de líneas (fibras ópticas)
# ─────────────────────────────────────────────────────────────────────────────
def extract_lines(page):
    """Extrae segmentos de línea que representan fibras."""
    lines = []
    for ln in page.lines:
        lw = ln.get('linewidth', 1.0)
        if lw > FIBER_MAX_LINEWIDTH:
            continue
        lines.append({
            'bx': ln['x0'], 'by': ln['top'],
            'ex': ln['x1'], 'ey': ln['bottom'],
            'page': page.page_number
        })
    return lines


# ─────────────────────────────────────────────────────────────────────────────
# Parseo de texto de board en PDF
# En PDF el texto queda como "G2WSMD901(1_B04)MAIPUVTR" (colapsado)
# ─────────────────────────────────────────────────────────────────────────────
def parse_board_text_pdf(text: str):
    """
    Limpia y parsea el ID del board de forma robusta.
    """
    # Limpiar caracteres de control pero preservar letras, números y paréntesis
    clean_txt = "".join(c for c in text.upper() if c.isalnum() or c in '()_-')
    
    # Intentar capturar la placa que está justo antes del paréntesis
    # Ejemplo: DM2DM3G2WSMD901(1_B04) -> queremos G2WSMD901
    m = re.search(r'([A-Z0-9]+)\((\d+)[_\-]([A-Z])(\d+)\)', clean_txt)
    if m:
        board_raw, rack, subrack, slot_raw = m.groups()
        
        # Si la placa capturada es muy larga (>12), probablemente trae ruido de etiquetas internas (DM1, DM2...)
        # Intentamos buscar el último 'G2' o 'TNG2' o placa conocida
        if len(board_raw) > 10:
            for p in ['TNG2', 'G2', 'M808', 'DAP', 'WSMD', 'OH9S', 'DWSS', 'OPM8', 'OACU']:
                idx = board_raw.rfind(p)
                if idx != -1:
                    board_raw = board_raw[idx:]
                    break
        
        return (normalize_board(board_raw), rack, subrack.upper(), str(int(slot_raw)).zfill(2))
    
    # Fallback: buscar solo la parte del slot (Rack_Slot)
    m_slot = re.search(r'\((\d+)[_\-]([A-Z])(\d+)\)', clean_txt)
    if m_slot:
        rack, subrack, slot = m_slot.groups()
        plate = "UNK"
        for p in ['WSMD', 'DAP', 'DWSS', 'OH9S', 'OPM8', 'OACU', 'VA4', 'M808SA']:
            if p in clean_txt:
                plate = p
                break
        return (normalize_board(plate), rack, subrack.upper(), str(int(slot)).zfill(2))
        
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Funciones de proximidad para coordenadas PDF
# ─────────────────────────────────────────────────────────────────────────────
def find_board_at_endpoint(x, y, page_num, boards, y_tolerance=None, x_limit=50.0):
    """
    Encuentra el board cuyo rectángulo abarca el punto Y (con tolerancia)
    y es el más cercano en X. Mismo concepto que VSDX pero en pts PDF.
    """
    best, min_dist = None, x_limit
    for b in boards:
        if b['page'] != page_num:
            continue
        half_h = b['h'] / 2
        tol = half_h * 0.5 + 5.0  # 50% del board + 5pts fijo
        if abs(b['y'] - y) <= half_h + tol:
            dist = abs(b['x'] - x)
            if dist < min_dist:
                min_dist = dist
                best = b
    return best


def find_label_at_endpoint(x, y, page_num, labels, threshold=LABEL_THRESHOLD_PTS):
    """
    Distancia Euclidiana para labels. 
    Prioriza labels que parecen puertos (IN, OUT, TX, RX) sobre labels genéricos.
    """
    best, min_dist = None, float('inf')
    
    # Puertos estándar Huawei
    PORT_KEYWORDS = {'IN', 'OUT', 'VI', 'VO', 'TX', 'RX', 'MON', 'RM', 'TM', 'LOUT', 'LIN'}

    for l in labels:
        if l['page'] != page_num:
            continue
        dist = ((l['x'] - x) ** 2 + (l['y'] - y) ** 2) ** 0.5
        if dist < threshold:
            # Prioridad virtual para puertos conocidos
            txt = l['text'].upper()
            virtual_dist = dist
            if any(k in txt for k in PORT_KEYWORDS):
                virtual_dist -= 5.0 # Prioridad
                
            if virtual_dist < min_dist:
                min_dist = virtual_dist
                best = l
    return best


def find_odf_label_at_endpoint(x, y, page_num, labels, threshold=ODF_THRESHOLD_PTS):
    """Búsqueda ODF con radio ampliado."""
    best, min_dist = None, float('inf')
    for l in labels:
        if l['page'] != page_num:
            continue
        if not is_odf_label(l['text']):
            continue
        dist = ((l['x'] - x) ** 2 + (l['y'] - y) ** 2) ** 0.5
        if dist < min_dist:
            min_dist = dist
            best = l
    return best if min_dist < threshold else None


# ─────────────────────────────────────────────────────────────────────────────
# Detección de sitio para PDF
# ─────────────────────────────────────────────────────────────────────────────
_PDF_SITE_RE = re.compile(
    r'([A-Z]{3,8})\s*(VTR|MOVISTAR|WOM|ENTEL|CLARO|FON)',
    re.IGNORECASE
)


def detect_site_pdf(words) -> str:
    """Detecta sitio ignorando palabras que parecen puertos."""
    all_text = ' '.join(w['text'] for w in words).upper()
    
    # Excluir palabras cortas que suelen ser puertos (LOUT, LIN, TM1...)
    blacklist = {'LOUT', 'LIN', 'TM1', 'TM2', 'RM1', 'RM2', 'MON'}
    
    hits = []
    for m in _PDF_SITE_RE.finditer(all_text):
        city, ope = m.group(1), m.group(2)
        if city not in blacklist and len(city) >= 4:
            hits.append((city, ope))
            
    if not hits:
        # Fallback agresivo: buscar MAIPU o nombres conocidos
        for name in ['MAIPU', 'LA CISTERNA', 'SANTIAGO', 'PUDAHUEL']:
            if name in all_text:
                return normalize_site(name.replace(' ', ''), 'VTR')
        return 'SITE_UNK'
        
    vtr = [(c, o) for c, o in hits if o == 'VTR']
    city, ope = vtr[0] if vtr else hits[0]
    return normalize_site(city, ope)


# ─────────────────────────────────────────────────────────────────────────────
# Motor principal PDF
# ─────────────────────────────────────────────────────────────────────────────
def parse_pdf_connections(file_path: str, label_threshold: float = LABEL_THRESHOLD_PTS) -> list:
    """
    Parsea un PDF vectorial DWDM Huawei y retorna lista de {'FROM:', 'TO:'}.
    """
    raw_connections = []

    print(f"\n[PDF] Procesando: {os.path.basename(file_path)}")

    try:
        with pdfplumber.open(file_path) as pdf:
            print(f"[PDF] {len(pdf.pages)} páginas")

            for page in pdf.pages:
                pn = page.page_number

                # Extraer palabras con tolerancia para texto vertical fragmentado
                words = page.extract_words(
                    x_tolerance=3,
                    y_tolerance=3,
                    keep_blank_chars=False,
                    use_text_flow=False
                )

                # Detectar sitio
                site = detect_site_pdf(words)

                # Clasificar shapes
                boards = detect_boards_from_rects(page, words)
                labels = extract_labels(page, words, boards)
                lines  = extract_lines(page)

                print(f"  [Pág {pn}] Boards={len(boards)}, Labels={len(labels)}, Lines={len(lines)}, Sitio={site}")

                # Debug primeros boards
                for b in boards[:4]:
                    print(f"    BOARD '{b['text'][:40]}' @ ({b['x']:.1f},{b['y']:.1f}) H={b['h']:.1f}")

                # Procesar líneas
                for line in lines:
                    # Extremo FROM (BeginX/Y en PDF)
                    bx, by = line['bx'], line['by']
                    # Extremo TO
                    ex, ey = line['ex'], line['ey']

                    # Boards en extremos
                    b_f = find_board_at_endpoint(bx, by, pn, boards)
                    b_t = find_board_at_endpoint(ex, ey, pn, boards)

                    # Labels en extremos
                    l_f = find_label_at_endpoint(bx, by, pn, labels, label_threshold)
                    l_t = find_label_at_endpoint(ex, ey, pn, labels, label_threshold)

                    # ── Resolver FROM ──
                    from_ep = None
                    if b_f:
                        parsed = parse_board_text_pdf(b_f['text'])
                        if parsed:
                            port_txt = ''
                            if l_f and not is_annotation(l_f['text']) and not is_odf_label(l_f['text']):
                                port_txt = l_f['text']
                            from_ep = build_endpoint("F:", parsed, port_txt, site)
                    else:
                        # Buscar ODF o "From/To" label
                        odf_l = find_label_at_endpoint(bx, by, pn, labels, ODF_THRESHOLD_PTS)
                        if odf_l:
                            txt = odf_l['text'].upper()
                            if is_odf_label(txt) or any(k in txt for k in ['FROM', 'TO', 'LIN', 'LOUT']):
                                odf = extract_odf_name(txt, site)
                                from_ep = odf if odf.startswith('F:') else f"F:{odf}"

                    # ── Resolver TO ──
                    to_ep = None
                    if b_t:
                        parsed = parse_board_text_pdf(b_t['text'])
                        if parsed:
                            port_txt = ''
                            if l_t and not is_annotation(l_t['text']) and not is_odf_label(l_t['text']):
                                port_txt = l_t['text']
                            to_ep = build_endpoint("T:", parsed, port_txt, site)
                    else:
                        # Buscar ODF o "From/To" label
                        odf_l = find_label_at_endpoint(ex, ey, pn, labels, ODF_THRESHOLD_PTS)
                        if odf_l:
                            txt = odf_l['text'].upper()
                            if is_odf_label(txt) or any(k in txt for k in ['FROM', 'TO', 'LIN', 'LOUT']):
                                odf = extract_odf_name(txt, site)
                                to_ep = odf if odf.startswith('T:') else f"T:{odf}"

                    if from_ep and to_ep and from_ep != to_ep:
                        # Limpieza final de ruidos comunes en PDF (labels de una sola letra sin contexto)
                        if len(from_ep.split('-')[-1]) == 1 and from_ep.split('-')[-1] not in '1234567890':
                           if not any(k in from_ep for k in ['DAP', 'OH9S']): # DAP y OH9S sí pueden tener puertos cortos
                               pass # skip suspicious single-letter ports
                        
                        raw_connections.append({'FROM:': from_ep, 'TO:': to_ep})

    except Exception as e:
        import traceback
        print(f"[ERROR PDF] {e}")
        traceback.print_exc()
        return []

    # Deduplicar
    seen = set()
    unique = []
    for c in raw_connections:
        k = (c['FROM:'], c['TO:'])
        if k not in seen:
            seen.add(k)
            unique.append(c)

    print(f"\n[PDF] {len(unique)} conexiones únicas extraídas.")
    return unique
