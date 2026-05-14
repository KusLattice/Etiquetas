"""
core_parser.py — Parser DWDM Huawei para diagramas Visio (.vsdx)
Genera Excel con columnas FROM: y TO: para impresora Brother P-touch.

Arquitectura:
  - Shapes clasificados en 3 tipos:
      boards: h > 1.0 y w < 1.0  → info de hardware (placa, rack, slot)
      labels: resto con texto     → info de puerto
      lines: tienen BeginX/EndX  → conexiones a resolver
  - Para cada extremo de línea se busca por separado:
      find_closest_board() → board (hardware)
      find_closest_label() → label  (puerto)
  - Normalización completa de placas, sitio, subrack/slot y puertos
"""

import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os

# ─────────────────────────────────────────────────────────────────────────────
# Namespace Visio
# ─────────────────────────────────────────────────────────────────────────────
NS = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}

# ─────────────────────────────────────────────────────────────────────────────
# Tabla de normalización de placas (excepciones explícitas primero)
# ─────────────────────────────────────────────────────────────────────────────
BOARD_EXCEPTIONS = {
    'DWSS2001': 'DWSS20',
    'DWSS20':   'DWSS20',
    'OH9S01':   'OH9S',
    'OH9S':     'OH9S',
    'WSMD901':  'WSMD9',
    'WSMD9':    'WSMD9',
    'OHN801':   'OHN8',
    'OHN8':     'OHN8',
    'OPM8':     'OPM8',
    'DAP':      'DAP',
}

# Sufijos que indican número de revisión y deben eliminarse
_SUFFIX_RE = re.compile(r'(S\d+|0[01])$', re.IGNORECASE)
# Prefijos a eliminar (incluyendo posibles dígitos de rack/subrack al inicio en PDF)
_PREFIX_RE = re.compile(r'^(\d+)?(TNG2|G2)', re.IGNORECASE)
# M808SA: quitar sufijo Fxx y números de frecuencia
_M808_RE   = re.compile(r'^M808SA', re.IGNORECASE)
# P2ON32S: quitar 01 final
_P2ON_RE   = re.compile(r'^P2ON32S', re.IGNORECASE)

# Ruido técnico a eliminar de los puertos (PDF)
_TECHNICAL_NOISE = [
    r'\d+(\.\d+)?(G|M)?Hz',  # 75.0GHz, 50GHz, etc.
    r'[-+]?\d+(\.\d+)?dBm?', # 0.5dBm, -10dBm, etc.
    r'\(\d+_[A-Z]\d+\)',    # (1_C01), (2_B04) - PDF coordinate tags
    r'^[A-Z]\d{1,2}(?=[A-Z])', # Prefix noise like 'G2' in 'G2WSMD9'
]

# ─────────────────────────────────────────────────────────────────────────────
# Lógica de Negocio de Telecomunicaciones (Senior Level)
# ─────────────────────────────────────────────────────────────────────────────

# Whitelist de puertos que requieren parcheo físico (excepciones a la regla anti-loopback)
PHYSICAL_PATCH_WHITELIST = {
    'DAP': ['VO_1', 'VO_2', 'IN_1', 'IN_2', 'VI_1', 'VI_2', 'OUT_1', 'OUT_2'],
    'WSMD9': ['EXP', 'COM'],
    'DWSS20': ['EXP', 'COM'],
    'M808SA': ['DM01', 'AM01', 'IN(P)', 'OUT(P)'],
    'OPM8': [f'IN{i}' for i in range(1, 9)] + ['MON', 'MONT', 'MONR'],
    'AUX': ['ETH', 'NM', 'EXT', 'RJ45'],
    'CTU': ['ETH', 'NM', 'EXT', 'RJ45'],
}

def is_internal_loop(f_ep: str, t_ep: str) -> bool:
    """Verifica si la conexión es un bucle interno (mismo rack/slot/board)."""
    def get_body(ep):
        # Quitar prefijo F: o T:
        ep_no_pref = ep[2:] if ep.startswith(('F:', 'T:')) else ep
        parts = ep_no_pref.split('-')
        return '-'.join(parts[:-1]) if len(parts) > 1 else ep_no_pref
    
    f_body = get_body(f_ep)
    t_body = get_body(t_ep)
    return f_body == t_body and f_body != f_ep

def is_bridge_exception(ep: str) -> bool:
    """Verifica si el puerto pertenece a la lista blanca de parcheo obligatorio."""
    parts = ep.split('-')
    if len(parts) < 3: return False
    board = parts[-2].upper()
    port = parts[-1].upper()
    
    allowed = PHYSICAL_PATCH_WHITELIST.get(board, [])
    return port in allowed

def validate_telecom_connection(f_ep: str, t_ep: str) -> bool:
    """
    Motor de Filtrado de Ingeniería (is_physical_fiber).
    Aplica la regla Anti-Loopback con excepciones de la WhiteList.
    """
    if f_ep == t_ep: return False
    
    if is_internal_loop(f_ep, t_ep):
        # Es loopback. Permitir solo si alguno de los puertos es una excepción técnica.
        if is_bridge_exception(f_ep) or is_bridge_exception(t_ep):
            return True
        return False # Bloquear loopback puro
    return True


def normalize_board(raw: str) -> str:
    """Normaliza el nombre de una placa Huawei DWDM."""
    s = raw.strip().upper()
    s = _PREFIX_RE.sub('', s)          # quitar G2 / TNG2

    # Excepciones directas (antes de quitar sufijos)
    if s in BOARD_EXCEPTIONS:
        return BOARD_EXCEPTIONS[s]

    # M808SAFxx → M808SA
    if _M808_RE.match(s):
        return 'M808SA'

    # P2ON32Sxx → P2ON32S
    if _P2ON_RE.match(s):
        return 'P2ON32S'

    # Regla general: quitar sufijo de revisión (S01, 01, 00)
    s = _SUFFIX_RE.sub('', s)

    # Segunda pasada de excepciones (post-strip)
    return BOARD_EXCEPTIONS.get(s, s)


# ─────────────────────────────────────────────────────────────────────────────
# Normalización de sitio
# ─────────────────────────────────────────────────────────────────────────────
def normalize_site(city: str, ope: str) -> str:
    """'MAIPU', 'VTR' → 'MAIP_VTR'"""
    return f"{city.strip()[:4].upper()}_{ope.strip().upper()}"


# ─────────────────────────────────────────────────────────────────────────────
# Extracción de texto Visio (chars fragmentados → string limpio)
# ─────────────────────────────────────────────────────────────────────────────
def get_text(shape) -> str:
    """Extrae el texto completo de un shape, colapsando fragmentación char-a-char."""
    elem = shape.find('.//v:Text', NS)
    if elem is None:
        return ''
    raw = ''.join(elem.itertext())
    # Quitar whitespace entre caracteres individuales (artefacto de formateo Visio)
    cleaned = re.sub(r'(?<=\S)\s+(?=\S)', '', raw)
    return cleaned.strip()


def get_num(shape, cell_name, default=0.0):
    cell = shape.find(f'.//v:Cell[@N="{cell_name}"]', NS)
    if cell is not None:
        try:
            return float(cell.attrib.get('V', default))
        except (ValueError, TypeError):
            pass
    return default


def get_num_or_none(shape, cell_name):
    cell = shape.find(f'.//v:Cell[@N="{cell_name}"]', NS)
    if cell is not None:
        try:
            return float(cell.attrib.get('V', 0))
        except (ValueError, TypeError):
            pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Parseo de texto de board → (board_norm, rack, subrack, slot)
# ─────────────────────────────────────────────────────────────────────────────
# Formato en Visio (tras colapsar whitespace): "G2WSMD901(1_B04)" o "OPM8(1_B06)"
# Pattern: BOARDNAME(RACK_SUBRACKSLOT)   donde SUBRACK es letra, SLOT son dígitos
_BOARD_PARSE_RE = re.compile(
    r'^([A-Z0-9]+)'        # nombre de placa (con prefijo G2, etc.)
    r'\((\d+)'             # (rack número
    r'[_\-]'
    r'([A-Za-z])'          # subrack letra (A, B, C...)
    r'(\d+)\)',            # slot número)
    re.IGNORECASE
)


def parse_board_shape(text: str):
    """
    Parsea el texto compacto de un shape de board.
    Retorna (board_norm, rack_str, subrack_str, slot_zfill2) o None.
    Ej: 'G2WSMD901(1_B04)' → ('WSMD9', '1', 'B', '04')
    """
    m = _BOARD_PARSE_RE.match(text.strip().upper())
    if m:
        board_raw, rack, subrack, slot_raw = m.groups()
        return (
            normalize_board(board_raw),
            rack,
            subrack.upper(),
            str(int(slot_raw)).zfill(2)
        )
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Filtros de labels de anotación (no son puertos)
# ─────────────────────────────────────────────────────────────────────────────
_ANNOTATION_PATTERNS = [
    re.compile(r'^\d+(\.\d+)?\s*(dBm|THz|km|dB)$', re.IGNORECASE),  # valores físicos
    re.compile(r'^\d+km[\-\s]?v\d', re.IGNORECASE),                   # '150km-V3.1'
    re.compile(r'^v\d+(\.\d+)?$', re.IGNORECASE),                     # versiones V3.1
    re.compile(r'^\d+$'),                                              # solo dígitos
    re.compile(r'^\d+\)$'),                                            # números aislados tipo '1)'
    re.compile(r'^(APD|PIN)$', re.IGNORECASE),                        # detectores
    re.compile(r'^[RT]X\d{4,}$', re.IGNORECASE),                     # frecuencias RX21491, TX11511
    re.compile(r'150\s*km', re.IGNORECASE),
    re.compile(r'proyectado|existente|software|reserva', re.IGNORECASE),
    re.compile(r'^(?:From|To)\s+[A-Z]', re.IGNORECASE),              # refs externas largas
    # Labels de leyenda / notas largas del diagrama
    re.compile(r'(fiber|rack|board|connected|outside|legend|virtual|backplane|customer|atenuator|fixed)', re.IGNORECASE),
    re.compile(r'[RT]X_?\d{4,}', re.IGNORECASE),                      # Frecuencias con/sin underscore
    re.compile(r'\d{4,}\s*(THz|nm)', re.IGNORECASE),                 # Wavelengths
]

def is_annotation(text: str) -> bool:
    """True si el texto es una anotación técnica, no un nombre de puerto."""
    if not text or len(text.strip()) < 1:
        return True
    t = text.strip()
    # Strings muy largos NO son puertos (pero podrían ser ODF — se evalúan antes)
    if len(t) > 25 and not is_odf_label(t):
        return True
    for pat in _ANNOTATION_PATTERNS:
        if pat.search(t):
            return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Detección de labels ODF / conexiones externas
# ─────────────────────────────────────────────────────────────────────────────
_ODF_FROM_RE = re.compile(
    r'From\s+([A-Za-z0-9\s]+?)\s*_\s*[A-Z0-9]+\([^)]+\)_L(?:OUT|IN)',
    re.IGNORECASE
)
_ODF_TO_RE = re.compile(
    r'To\s+([A-Za-z0-9\s]+?)\s*_\s*[A-Z0-9]+\([^)]+\)_L(?:OUT|IN)',
    re.IGNORECASE
)
# También el shape puede tener texto genérico tipo "ODF_MAIPU_VTR_B"
_ODF_DIRECT_RE = re.compile(r'^ODF(_|\s|$)', re.IGNORECASE)

# Determinar si un extremo de línea NO tiene board pero llega a un label ODF
def is_odf_label(text: str) -> bool:
    t = text.strip()
    return bool(
        _ODF_FROM_RE.search(t) or
        _ODF_TO_RE.search(t)   or
        _ODF_DIRECT_RE.match(t)
    )


def extract_odf_name(text: str, local_site: str) -> str:
    """
    Extrae el nombre ODF de un label de conexión externa.
    'From MAIPU VTR_G2OH9S01(1_A12)_LOUT (MAIPU VTR...)' → 'ODF_MAIPU_VTR_B'

    Cuando el texto contiene 'From [SITIO]...' significa que la fibra viene de ese
    sitio externo → el endpoint ODF en el diagrama local es el "lado B" del ODF.
    Por convención usamos el nombre de sitio local + sufijo del cable.
    """
    t = text.strip()

    # Caso directo: "ODF_MAIPU_VTR_B"
    if _ODF_DIRECT_RE.match(t):
        return t.upper().replace(' ', '_')

    # Extraer sitio del "From ..." o "To ..."
    m = _ODF_FROM_RE.search(t) or _ODF_TO_RE.search(t)
    if m:
        site_raw = m.group(1).strip()
        # "MAIPU VTR" → "MAIPU_VTR"
        site_clean = re.sub(r'\s+', '_', site_raw).upper()
        return f"ODF_{site_clean}_B"

    # Fallback
    return f"ODF_{local_site.replace(' ', '_').upper()}_B"


# ─────────────────────────────────────────────────────────────────────────────
# Detección de sitio en la página
# ─────────────────────────────────────────────────────────────────────────────
_SITE_PATTERNS = [
    # "MAIPU VTR" (con espacio)
    re.compile(r'\b([A-Z]{3,8})\s+(VTR|MOVISTAR|WOM|ENTEL|CLARO|FON)\b', re.IGNORECASE),
    # "MAIPUVTR" (colapsado, sin espacio)
    re.compile(r'\b([A-Z]{3,8})(VTR|MOVISTAR|WOM|ENTEL|CLARO|FON)\b', re.IGNORECASE),
    # "From MAIPU VTR_..." en labels ODF
    re.compile(r'(?:From|To)\s+([A-Z]{3,8})\s*(VTR|MOVISTAR|WOM|ENTEL|CLARO|FON)', re.IGNORECASE),
]
_SITE_BLACKLIST = {'OPM', 'G2OH', 'G2DAP', 'G2WS', 'FON'}


def detect_site(labels, boards) -> str:
    """
    Detecta (city, ope) del sitio local; prioriza VTR sobre FON.
    Acepta tanto listas como dicts {sid: shape_info}.
    """
    hits = []
    _labels = labels.values() if isinstance(labels, dict) else labels
    _boards = boards.values() if isinstance(boards, dict) else boards
    candidates = [l['text'] for l in _labels] + [b['text'] for b in _boards]
    for text in candidates:
        tu = text.upper()
        for pat in _SITE_PATTERNS:
            m = pat.search(tu)
            if m:
                city, ope = m.group(1).upper(), m.group(2).upper()
                if city not in _SITE_BLACKLIST:
                    hits.append((city, ope))
    if not hits:
        return 'SITE_UNK'
    # Priorizar VTR
    vtr = [(c, o) for c, o in hits if o == 'VTR']
    city, ope = vtr[0] if vtr else hits[0]
    return normalize_site(city, ope)


# ─────────────────────────────────────────────────────────────────────────────
# Funciones de proximidad geométrica
# ─────────────────────────────────────────────────────────────────────────────
def find_closest_board(x, y, page, boards, y_tolerance=0.5, x_limit=3.0):
    """
    Encuentra el board más cercano al punto (x, y) en la misma página.
    Criterio: overlap vertical (con tolerancia) + menor distancia horizontal.
    boards puede ser dict {sid: info} o lista.
    """
    best, min_dist = None, x_limit
    items = boards.values() if isinstance(boards, dict) else boards
    for b in items:
        if b['page'] != page:
            continue
        if abs(b['y'] - y) <= b['h'] / 2 + y_tolerance:
            dist = abs(b['x'] - x)
            if dist < min_dist:
                min_dist = dist
                best = b
    return best


def find_closest_label(x, y, page, labels, threshold=0.5):
    """
    Distancia Euclidiana al label más cercano. Retorna None si supera el threshold.
    labels puede ser dict {sid: info} o lista.
    """
    best, min_dist = None, float('inf')
    items = labels.values() if isinstance(labels, dict) else labels
    for l in items:
        if l['page'] != page:
            continue
        dist = ((l['x'] - x) ** 2 + (l['y'] - y) ** 2) ** 0.5
        if dist < min_dist:
            min_dist = dist
            best = l
    return best if min_dist < threshold else None


# ─────────────────────────────────────────────────────────────────────────────
# Constructor de endpoint
# ─────────────────────────────────────────────────────────────────────────────
def _clean_port(raw: str) -> str:
    """Limpia el nombre del puerto de ruido técnico y anotaciones del PDF."""
    s = raw.strip()
    
    # 1. Quitar frecuencias/wavelengths y todo lo que sigue (ej: _1491_CISTERNA -> '')
    # Buscamos patrones de 4 dígitos (típico de lambda/frecuencia)
    s = re.sub(r'_?\d{4,}.*$', '', s)

    # 2. Quitar ruido técnico específico
    for pattern in _TECHNICAL_NOISE:
        s = re.sub(pattern, '', s, flags=re.IGNORECASE)

    # 3. Quitar sufijos comunes de anotación que ensucian el puerto
    s = re.sub(r'_(FROM|TO|ATENUATOR|FIXED|LA|RM|TM|EI|AM|EO|DM).*$', '', s, flags=re.IGNORECASE)

    # 4. Limpiar caracteres especiales al inicio/fin
    s = s.strip('-_ ')
    
    # 5. Normalizar (P) y (W)
    s = s.replace('(p)', '(P)').replace('(w)', '(W)')

    # 6. Balancear paréntesis
    opens  = s.count('(')
    closes = s.count(')')
    if closes > opens:
        s = s[:-(closes - opens)]
        
    return s.strip().upper() or 'P'


def build_endpoint(prefix: str, board_parsed: tuple, port_text: str, site: str, project: str = 'Core', board_raw: str = '') -> str:
    """
    Construye el endpoint según el proyecto.
    Genérico: F:NOMBRE_PLACA - PUERTO
    Estándar: F:RD_MAIP_VTR_1B(04)-WSMD9-OUT
    """
    port = _clean_port(port_text) if port_text else 'P'
    
    if project == 'Genérico':
        # Limpiar el texto original de la placa (quitar saltos de línea y espacios extra)
        b_name = board_raw.split('(')[0].strip() if board_raw else board_parsed[0]
        return f"{prefix}{b_name} - {port}"

    board_norm, rack, subrack, slot = board_parsed
    return f"{prefix}RD_{site}_{rack}{subrack}({slot})-{board_norm}-{port}"

def find_closest_odf_label(x, y, page, labels, threshold=0.6):
    """
    Búsqueda de labels ODF con radio mayor (los ODF suelen estar alejados del extremo).
    labels puede ser dict {sid: info} o lista.
    """
    best, min_dist = None, float('inf')
    items = labels.values() if isinstance(labels, dict) else labels
    for l in items:
        if l['page'] != page:
            continue
        if not is_odf_label(l['text']):
            continue
        dist = ((l['x'] - x) ** 2 + (l['y'] - y) ** 2) ** 0.5
        if dist < min_dist:
            min_dist = dist
            best = l
    return best if min_dist < threshold else None

# ─────────────────────────────────────────────────────────────────────────────
def parse_vsdx_connections(file_path: str, telco_name: str = 'ClaroVTR',
                           label_threshold: float = 0.8, project_name: str = 'Core') -> list:
    """
    Parsea un VSDX Huawei y retorna lista de dicts {'FROM:': str, 'TO:': str}.

    Clasificación de shapes:
      boards {sid: info} — h > 1.0 y w < 1.0  → info de hardware (placa/rack/slot)
      labels {sid: info} — resto con texto      → info de puerto
      lines  [...]       — tienen BeginX/EndX   → conexiones a resolver
    """
    boards = {}   # sid → {'x','y','w','h','text','page'}  — shapes de hardware
    labels = {}   # sid → {'x','y','text','page'}          — shapes de puerto/anotación
    lines_all = []

    # ── 1. Leer y clasificar todos los shapes ─────────────────────────────────
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            pages = sorted([
                f for f in z.namelist()
                if 'visio/pages/page' in f and f.endswith('.xml') and 'rels' not in f
            ])

            print(f"[INFO] {len(pages)} páginas encontradas en {os.path.basename(file_path)}")

            for page_name in pages:
                root = ET.fromstring(z.read(page_name))
                page_boards, page_labels, page_lines = 0, 0, 0

                for shape in root.findall('.//v:Shape', NS):
                    sid  = shape.attrib.get('ID', '')
                    text = get_text(shape)
                    pinx = get_num(shape, 'PinX')
                    piny = get_num(shape, 'PinY')
                    w    = get_num(shape, 'Width')
                    h    = get_num(shape, 'Height')
                    bx   = get_num_or_none(shape, 'BeginX')
                    ex   = get_num_or_none(shape, 'EndX')

                    if bx is not None and ex is not None:
                        # ── Línea / conector ──────────────────────────────────
                        by = get_num(shape, 'BeginY')
                        ey = get_num(shape, 'EndY')
                        lines_all.append({
                            'bx': bx, 'by': by, 'ex': ex, 'ey': ey,
                            'page': page_name
                        })
                        page_lines += 1

                    elif h > 1.0 and w < 1.0 and text:
                        # ── Board (hardware): alto y angosto ──────────────────
                        # h > 1.0: altura mínima de 1 unidad Visio (~2.54 cm)
                        # w < 1.0: ancho menor a 1 unidad (shapes de placa típicos)
                        boards[f"{page_name}_{sid}"] = {
                            'x': pinx, 'y': piny, 'w': w, 'h': h,
                            'text': text, 'page': page_name
                        }
                        page_boards += 1

                    elif text:
                        # ── Label (puerto / anotación): el resto con texto ────
                        labels[f"{page_name}_{sid}"] = {
                            'x': pinx, 'y': piny,
                            'text': text, 'page': page_name
                        }
                        page_labels += 1

                print(f"  [{page_name}] Boards={page_boards}, Labels={page_labels}, Lines={page_lines}")

    except zipfile.BadZipFile:
        print(f"[ERROR] No es un ZIP válido: {file_path}")
        return []
    except Exception as e:
        import traceback
        print(f"[ERROR] Leyendo VSDX: {e}")
        traceback.print_exc()
        return []

    print(f"\n[TOTAL] Boards={len(boards)}, Labels={len(labels)}, Lines={len(lines_all)}")

    # ── Debug: primeros shapes ─────────────────────────────────────────────────
    print("\n[DEBUG] Primeros 15 boards (hardware):")
    for b in list(boards.values())[:15]:
        print(f"  BOARD '{b['text'][:50]}' @ ({b['x']:.3f}, {b['y']:.3f}) W={b['w']:.3f} H={b['h']:.3f} [{b['page']}]")
    print("\n[DEBUG] Primeros 15 labels (puertos):")
    for l in list(labels.values())[:15]:
        print(f"  LABEL '{l['text'][:50]}' @ ({l['x']:.3f}, {l['y']:.3f}) [{l['page']}]")
    if lines_all:
        print("\n[DEBUG] Primeras 5 líneas:")
        for ln in lines_all[:5]:
            print(f"  LINE  bx={ln['bx']:.3f} by={ln['by']:.3f} ex={ln['ex']:.3f} ey={ln['ey']:.3f} [{ln['page']}]")

    # ── 2. Detectar sitio por página ──────────────────────────────────────────
    pages_seen = list(dict.fromkeys(
        info['page'] for info in list(boards.values()) + list(labels.values())
    ))
    site_map = {}
    for pg in pages_seen:
        pg_boards = {k: v for k, v in boards.items() if v['page'] == pg}
        pg_labels = {k: v for k, v in labels.items() if v['page'] == pg}
        site_map[pg] = detect_site(pg_labels, pg_boards)
        print(f"  [{pg}] Sitio detectado: {site_map[pg]}")

    # ── 3. Procesar líneas → conexiones ───────────────────────────────────────
    raw_connections = []

    # Debug de primera línea
    if lines_all:
        ln = lines_all[0]
        pg = ln['page']
        print(f"\n[DEBUG] Primera línea: bx={ln['bx']:.3f} by={ln['by']:.3f} ex={ln['ex']:.3f} ey={ln['ey']:.3f}")
        b_f = find_closest_board(ln['bx'], ln['by'], pg, boards)
        b_t = find_closest_board(ln['ex'], ln['ey'], pg, boards)
        l_f = find_closest_label(ln['bx'], ln['by'], pg, labels, label_threshold)
        l_t = find_closest_label(ln['ex'], ln['ey'], pg, labels, label_threshold)
        print(f"  FROM board: {b_f['text'][:30] if b_f else 'None'}")
        print(f"  FROM label: {l_f['text'][:30] if l_f else 'None'}")
        print(f"  TO   board: {b_t['text'][:30] if b_t else 'None'}")
        print(f"  TO   label: {l_t['text'][:30] if l_t else 'None'}")

    for line in lines_all:
        pg   = line['page']
        site = site_map.get(pg, 'SITE_UNK')

        # ── Extremo FROM: buscar board (hardware) Y label (puerto) por separado ──
        b_f = find_closest_board(line['bx'], line['by'], pg, boards)
        l_f = find_closest_label(line['bx'], line['by'], pg, labels, label_threshold)

        # ── Extremo TO: ídem ──────────────────────────────────────────────────
        b_t = find_closest_board(line['ex'], line['ey'], pg, boards)
        l_t = find_closest_label(line['ex'], line['ey'], pg, labels, label_threshold)

        # ── Resolver endpoint FROM ──
        from_ep = None
        # Prioridad 1: tiene board cercano → endpoint de placa
        if b_f:
            parsed = parse_board_shape(b_f['text'])
            if parsed:
                port_txt = ''
                if l_f and not is_annotation(l_f['text']) and not is_odf_label(l_f['text']):
                    port_txt = l_f['text']
                from_ep = build_endpoint("F:", parsed, port_txt, site, project_name, b_f['text'])
        # Prioridad 2: sin board → buscar label ODF con radio ampliado
        else:
            odf_l = find_closest_odf_label(line['bx'], line['by'], pg, labels, threshold=0.6)
            if odf_l:
                odf = extract_odf_name(odf_l['text'], site)
                from_ep = odf if odf.startswith('F:') else f"F:{odf}"

        # ── Resolver endpoint TO ──
        to_ep = None
        # Prioridad 1: tiene board (hardware) → combinar con label (puerto)
        if b_t:
            parsed = parse_board_shape(b_t['text'])
            if parsed:
                port_txt = ''
                if l_t and not is_annotation(l_t['text']) and not is_odf_label(l_t['text']):
                    port_txt = l_t['text']
                to_ep = build_endpoint("T:", parsed, port_txt, site, project_name, b_t['text'])
        # Prioridad 2: buscar label ODF con radio ampliado
        else:
            odf_l = find_closest_odf_label(line['ex'], line['ey'], pg, labels, threshold=0.6)
            if odf_l:
                odf = extract_odf_name(odf_l['text'], site)
                to_ep = odf if odf.startswith('T:') else f"T:{odf}"

        if from_ep and to_ep:
            # Calcular posición promedio en X para ordenar de izquierda a derecha
            avg_x = (line['bx'] + line['ex']) / 2
            
            # Intentar encontrar el "Equipo" (A, B, C) más cercano a la placa FROM
            equipment_label = "SIN_EQUIPO"
            if b_f:
                # Buscar label tipo "Equipo X" o "Rack X" arriba del board
                closest_eq = None
                min_eq_dist = 5.0  # límite de búsqueda
                for l in labels.values():
                    if l['page'] == pg and ('EQUIPO' in l['text'].upper() or 'RACK' in l['text'].upper()):
                        dist = ((l['x'] - b_f['x'])**2 + (l['y'] - b_f['y'])**2)**0.5
                        if dist < min_eq_dist:
                            min_eq_dist = dist
                            closest_eq = l['text'].strip()
                if closest_eq:
                    equipment_label = closest_eq

            raw_connections.append({
                'FROM:': from_ep, 
                'TO:': to_ep, 
                'page': pg, 
                'equipment': equipment_label,
                'x': avg_x
            })

    # Mostrar primeras 25 crudas
    print(f"\n[RAW] {len(raw_connections)} conexiones antes de filtros. Primeras 25:")
    for i, c in enumerate(raw_connections[:25], 1):
        print(f"  {i:02d}. {c['FROM:']:50s} | {c['TO:']}")

    # ── 4. Filtros ─────────────────────────────────────────────────────────────
    filtered = []
    for c in raw_connections:
        f_ep = c['FROM:']
        t_ep = c['TO:']

        # Reglas de negocio: Anti-Loopback y Whitelist
        if not validate_telecom_connection(f_ep, t_ep):
            continue

        filtered.append(c)

    # ── 5. Deduplicar ──────────────────────────────────────────────────────────
    seen = set()
    unique = []
    for c in filtered:
        # Usar tupla de FROM y TO para la unicidad
        key = (c['FROM:'], c['TO:'])
        if key not in seen:
            seen.add(key)
            unique.append(c)

    # ── 6. Ordenar ─────────────────────────────────────────────────────────────
    # Ordenar por Página, luego por Equipo, luego por X (izquierda a derecha)
    unique.sort(key=lambda c: (c['page'], c['equipment'], c['x']))

    # Limpiar campos auxiliares para el Excel final (pero mantener equipo para el agrupador)
    final_list = []
    for c in unique:
        final_list.append({'FROM:': c['FROM:'], 'TO:': c['TO:'], '_eq': c['equipment']})

    return final_list


# ─────────────────────────────────────────────────────────────────────────────
# Generación de Excel
# ─────────────────────────────────────────────────────────────────────────────
def generate_excel(connections: list, output_path: str) -> int:
    """Guarda las conexiones en Excel con separadores de equipo."""
    if not connections:
        print("[WARN] Sin conexiones para guardar.")
        return 0
    
    rows = []
    current_eq = None
    
    for c in connections:
        eq = c.get('_eq', 'SIN_EQUIPO')
        if eq != current_eq:
            # Insertar fila de cabecera de equipo (como en la captura 2)
            rows.append({'FROM:': eq, 'TO:': ''})
            current_eq = eq
        rows.append({'FROM:': c['FROM:'], 'TO:': c['TO:']})

    df = pd.DataFrame(rows, columns=['FROM:', 'TO:'])
    # NO deduplicar aquí para no borrar los headers
    
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    
    df.to_excel(output_path, index=False)
    print(f"[OK] Excel guardado: {output_path} ({len(df)} filas)")
    return len(df)
