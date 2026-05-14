import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os

# ─── Namespace Visio ────────────────────────────────────────────────────────
NS = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}

# ─── Normalización de Nombres de Placa ──────────────────────────────────────
# Reglas: quitar prefijos G2/TNG2, quitar sufijos S+digits / 01 / 00 al final
# Ej: G2WSMD901→WSMD9, G2DAP→DAP, G2DWSS2001→DWSS20, M808SAF3→M808SA, OPM8→OPM8
BOARD_STRIP_PREFIX = re.compile(r'^(?:TNG2|G2)', re.IGNORECASE)
# Sufijo a eliminar: solo "01" o "00" al final (revision digits)
# NO eliminamos S+digits si queda parte del nombre (ej: DWSS20 tiene S20 que es parte del nombre)
# Regla: eliminar solo si son exactamente 2 dígitos "00" o "01" al final
BOARD_STRIP_SUFFIX = re.compile(r'(?:S\d+|0[01])$', re.IGNORECASE)

# Nombres conocidos que NO deben perder su sufijo (excepción a la regla)
BOARD_KNOWN_EXCEPTIONS = {
    'DWSS2001': 'DWSS20',
    'OH9S01':   'OH9S',
    'WSMD901':  'WSMD9',
    'M808SAF3': 'M808SA',
}

def normalize_board(raw: str) -> str:
    """Normaliza el nombre de una placa DWDM Huawei."""
    s = raw.strip().upper()
    s = BOARD_STRIP_PREFIX.sub('', s)
    # Verificar excepciones conocidas primero
    if s in BOARD_KNOWN_EXCEPTIONS:
        return BOARD_KNOWN_EXCEPTIONS[s]
    # Regla general: quitar sufijo de revision solo si son 2 digits exactos (00/01)
    s = re.sub(r'0[01]$', '', s)
    return s

# ─── Normalización de Sitio ──────────────────────────────────────────────────
# "MAIPU VTR" → "MAIP_VTR" (4 chars primer token + _ + operadora)
def normalize_site(raw: str) -> str:
    """Devuelve XXXX_OPE a partir de un string de sitio libre."""
    raw = raw.strip().upper()
    # Eliminar prefijos RD_ si ya están
    raw = re.sub(r'^RD_', '', raw)
    parts = re.split(r'[\s_]+', raw)
    if len(parts) >= 2:
        site_part = parts[0][:4]
        ope_part  = parts[1]
        return f"{site_part}_{ope_part}"
    return parts[0][:8]  # fallback

# ─── Helpers de Geometría ───────────────────────────────────────────────────
def _get_float(shape, cell_name, default=0.0):
    cell = shape.find(f'.//v:Cell[@N="{cell_name}"]', NS)
    if cell is not None:
        try:
            return float(cell.attrib.get('V', default))
        except (ValueError, TypeError):
            pass
    return default

def _get_text(shape):
    elem = shape.find('.//v:Text', NS)
    if elem is None:
        return ''
    # Visio puede fragmentar texto caracter a caracter con nodos cp/pp.
    # itertext() los concatena todos; limpiamos espacios y saltos internos.
    raw = ''.join(elem.itertext())
    # Colapsar cualquier secuencia de whitespace a nada (el texto real no tiene espacios)
    # pero preservar el espacio antes del bloque de slot "(1_B04)"
    # Estrategia: quitar \n y espacios aislados entre letras/dígitos individuales
    cleaned = re.sub(r'(?<=\S)\s+(?=\S)', '', raw)  # quitar whitespace entre chars
    return cleaned.strip()

# ─── Clasificación de Shapes ─────────────────────────────────────────────────
# En diagramas DWDM Huawei:
#   BOARD  → shape vertical, height >> width, tiene el nombre de placa + subrack/slot
#   LABEL  → shape pequeño, tiene el nombre de puerto (VI_2, OUT, IN(P), etc.)
#   LINE   → tiene BeginX/EndX, es el conector físico
#   ODF    → shape con texto "ODF" o patrón "From/To Site_Board"

BOARD_ASPECT_RATIO = 2.0   # h/w mínimo para ser considerado board

def classify_shapes(root):
    """Separa shapes en boards, labels y lines."""
    boards = []
    labels = []
    lines  = []

    for shape in root.findall('.//v:Shape', NS):
        sid  = shape.get('ID', '')
        text = _get_text(shape)
        pinx = _get_float(shape, 'PinX')
        piny = _get_float(shape, 'PinY')
        w    = _get_float(shape, 'Width')
        h    = _get_float(shape, 'Height')
        bx   = _get_float(shape, 'BeginX', None)

        # ─ Línea / conector
        bx_elem = shape.find('./v:Cell[@N="BeginX"]', NS)
        if bx_elem is not None:
            bx = float(bx_elem.attrib.get('V', 0))
            by = _get_float(shape, 'BeginY')
            ex = _get_float(shape, 'EndX')
            ey = _get_float(shape, 'EndY')
            lines.append({'id': sid, 'bx': bx, 'by': by, 'ex': ex, 'ey': ey})
            continue

        # ─ Board vs Label por geometría
        if text:
            if h > BOARD_ASPECT_RATIO * max(w, 0.001):
                boards.append({'id': sid, 'x': pinx, 'y': piny, 'w': w, 'h': h, 'text': text})
            else:
                labels.append({'id': sid, 'x': pinx, 'y': piny, 'w': w, 'h': h, 'text': text})

    return boards, labels, lines

# Labels que deben ser ignorados como puertos (son anotaciones externas)
LABEL_BLACKLIST_PATTERNS = [
    re.compile(r'^(?:From|To)\s*[A-Z]', re.IGNORECASE),  # "From LACISTERNA..." / "To Maipu..."
    re.compile(r'^\d+\.\d+dBm$', re.IGNORECASE),          # "0.5dBm"
    re.compile(r'^\d+$'),                                   # solo dígitos
    re.compile(r'v\d', re.IGNORECASE),                      # versiones
]

def is_annotation_label(text: str) -> bool:
    """True si el texto parece una anotación técnica, no un nombre de puerto."""
    if not text or len(text) < 1:
        return True
    # Labels muy largos son referencias externas, no puertos
    if len(text) > 20:
        return True
    for pat in LABEL_BLACKLIST_PATTERNS:
        if pat.match(text):
            return True
    return False

def find_closest_board(x, y, boards, y_tolerance=0.6, x_threshold=3.0):
    """
    Encuentra el board cuya franja vertical abarca Y (con tolerancia)
    y está más cercano en X.
    """
    best     = None
    min_dist = x_threshold
    for b in boards:
        half_h = b['h'] / 2 + y_tolerance
        if abs(b['y'] - y) <= half_h:
            dist = abs(b['x'] - x)
            if dist < min_dist:
                min_dist = dist
                best = b
    return best

def find_closest_label(x, y, labels, threshold=0.6):
    """Encuentra el label más cercano a un punto dado."""
    best     = None
    min_dist = threshold
    for l in labels:
        dist = ((l['x'] - x) ** 2 + (l['y'] - y) ** 2) ** 0.5
        if dist < min_dist:
            min_dist = dist
            best = l
    return best

def find_odf_label(x, y, labels, threshold=1.5):
    """
    Versión más permisiva para encontrar labels ODF/externos
    que pueden estar más alejados.
    """
    return find_closest_label(x, y, labels, threshold=threshold)

# ─── Parser de Texto de Board ────────────────────────────────────────────────
# Formato Huawei en boards: "G2WSMD901\n1B_04" o "G2DAP\n1B_01"
# Subrack = primer token alfanumérico del segundo bloque (ej: "1B")
# Slot     = número final del segundo bloque (ej: "04")
# Formato real en el Visio: "G2WSMD901(1_B04)" o "G2DAP(1_B01)(1_TNG2OACU25S(T)2"
# El PRIMER grupo de paréntesis es el subrack/slot: (SUBRACK_SLOT)
# SUBRACK = número + letra (ej: 1B, 2A) o número solo (ej: 12)
# SLOT    = número de 2 dígitos
BOARD_TEXT_PATTERN = re.compile(
    r'^([A-Z0-9]+)'           # board_raw al inicio (G2WSMD901, G2DAP, OPM8...)
    r'\((\d+[A-Z]?)'          # (subrack: 1B, 2A, 1...
    r'[_\-]'
    r'([A-Z]?\d+)\)',         # slot: B04, 12, 01 con letra opcional
    re.IGNORECASE
)

def parse_board_text(text: str):
    """
    Parsea el texto compacto de un shape de board Huawei.
    Formato esperado tras clean: "G2WSMD901(1_B04)" o "OPM8(1_B06)"
    Retorna (board_norm, subrack_str, slot_zfill2) o None.
    """
    clean = text.strip().upper()
    m = BOARD_TEXT_PATTERN.match(clean)
    if m:
        board_raw, sub_prefix, slot_raw = m.groups()
        board_norm = normalize_board(board_raw)
        # subrack = sub_prefix + letra del slot si aplica
        # Ej: sub_prefix="1", slot_raw="B04" → subrack="1B", slot="04"
        slot_letter_match = re.match(r'^([A-Z])(\d+)$', slot_raw, re.IGNORECASE)
        if slot_letter_match:
            letter, num = slot_letter_match.groups()
            subrack = f"{sub_prefix}{letter}"
            slot_str = str(int(num)).zfill(2)
        else:
            # slot_raw es solo dígitos: ej "04"
            subrack  = sub_prefix
            slot_str = str(int(slot_raw)).zfill(2)
        return board_norm, subrack, slot_str
    return None

# ─── Detección de ODF ────────────────────────────────────────────────────────
ODF_PATTERN = re.compile(r'ODF|FROM\s+|TO\s+', re.IGNORECASE)

def is_odf_shape(text: str) -> bool:
    return bool(ODF_PATTERN.search(text))

def format_odf_endpoint(prefix: str, text: str, site_norm: str) -> str:
    """
    Formatea un endpoint ODF externo.
    Ej: "From MAIPU_VTR_B" → "F:ODF_MAIPU_VTR_B"
    """
    # Intentar extraer el sufijo del ODF
    m = re.search(r'(?:ODF|from|to)\s*[_\s]*([\w\s]+)', text, re.IGNORECASE)
    if m:
        suffix = m.group(1).strip().replace(' ', '_').upper()
        return f"{prefix}ODF_{suffix}"
    # Fallback: usar el sitio
    site_upper = site_norm.upper().replace('_', '_')
    return f"{prefix}ODF_{site_upper}"

# ─── Formateador de Endpoint ──────────────────────────────────────────────────
def _clean_port(port: str) -> str:
    """Elimina solo los paréntesis de cierre sin par al final del puerto.
    Preserva puertos como IN(P) y OUT(P) intactos."""
    if not port:
        return 'P'
    # Contar paréntesis: si hay más ')' que '(' al final, quitarlos
    opens  = port.count('(')
    closes = port.count(')')
    if closes > opens:
        port = port[:-(closes - opens)]
    return port.strip() or 'P'

def format_endpoint(prefix: str, board_info: tuple, port_text: str, site_norm: str) -> str:
    """
    Construye el endpoint final.
    board_info = (board_norm, subrack, slot)
    Ej: F:RD_MAIP_VTR_1B(04)-WSMD9-OUT
    """
    board_norm, subrack, slot = board_info
    port = _clean_port(port_text)
    return f"{prefix}RD_{site_norm}_{subrack}({slot})-{board_norm}-{port}"

# ─── Detección de Sitio ───────────────────────────────────────────────────────
# En los diagramas Huawei, el sitio suele aparecer en labels grandes.
# El texto puede estar colapsado (sin espacios) o con espacios:
# "MAIPU VTR", "MAIPUVTR", "From MAIPU VTR_..."
SITE_PATTERNS = [
    # Con espacio: "MAIPU VTR"
    re.compile(r'\b([A-Z]{3,8})\s+(VTR|MOVISTAR|WOM|ENTEL|CLARO|FON)\b', re.IGNORECASE),
    # Sin espacio: "MAIPUVTR", "LACISTERNAVTR" (colapsado)
    re.compile(r'\b([A-Z]{3,8})(VTR|MOVISTAR|WOM|ENTEL|CLARO|FON)\b', re.IGNORECASE),
    # En labels tipo "From MAIPU VTR_Board..."
    re.compile(r'(?:From|To)\s*([A-Z]{3,8})\s*(VTR|MOVISTAR|WOM|ENTEL|CLARO|FON)', re.IGNORECASE),
]

def detect_site(labels: list, boards: list) -> str:
    """Detecta el nombre del sitio a partir de los textos del diagrama.
    Prioriza VTR > FON > otros si hay múltiples coincidencias."""
    candidates = [l['text'] for l in labels] + [b['text'] for b in boards]
    results = []
    for text in candidates:
        text_up = text.upper()
        for pattern in SITE_PATTERNS:
            m = pattern.search(text_up)
            if m:
                city, ope = m.groups()
                if city.upper() in ('OPM', 'G2OH', 'G2DAP', 'G2WS'):
                    continue
                results.append(normalize_site(f"{city} {ope}"))
    if not results:
        return "SITE_UNK"
    # Priorizar resultados que contienen VTR sobre FON
    vtr_hits = [r for r in results if 'VTR' in r]
    if vtr_hits:
        return vtr_hits[0]
    return results[0]

# ─── Motor Principal ──────────────────────────────────────────────────────────
def parse_vsdx_connections(file_path: str, telco_name: str = "ClaroVTR") -> list:
    """
    Parsea un archivo VSDX de diagrama DWDM Huawei y retorna lista de conexiones.
    Cada conexión es un dict {'FROM:': str, 'TO:': str}.
    """
    extracted = []

    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            pages = sorted([
                f for f in z.namelist()
                if re.match(r'visio/pages/page\d+\.xml$', f)
            ])

            print(f"[INFO] Páginas encontradas: {len(pages)}")

            for page_path in pages:
                with z.open(page_path) as f:
                    root = ET.fromstring(f.read())

                boards, labels, lines = classify_shapes(root)

                print(f"  [{page_path}] Boards={len(boards)}, Labels={len(labels)}, Lines={len(lines)}")

                if not lines:
                    print(f"  [{page_path}] Sin líneas, saltando.")
                    continue

                # Detectar sitio en esta página
                site_norm = detect_site(labels, boards)
                print(f"  [{page_path}] Sitio detectado: {site_norm}")

                # Debug: primeros 5 boards y labels
                for b in boards[:5]:
                    print(f"    BOARD: '{b['text'][:40]}' @ ({b['x']:.2f}, {b['y']:.2f}) W={b['w']:.2f} H={b['h']:.2f}")
                for l in labels[:5]:
                    print(f"    LABEL: '{l['text'][:40]}' @ ({l['x']:.2f}, {l['y']:.2f})")

                for line in lines:
                    # ── Extremo FROM (Begin) ──
                    f_board = find_closest_board(line['bx'], line['by'], boards)
                    f_label = find_closest_label(line['bx'], line['by'], labels)

                    # ── Extremo TO (End) ──
                    t_board = find_closest_board(line['ex'], line['ey'], boards)
                    t_label = find_closest_label(line['ex'], line['ey'], labels)

                    f_info = parse_board_text(f_board['text']) if f_board else None
                    t_info = parse_board_text(t_board['text']) if t_board else None

                    # FILTRO loopback: ignorar solo si board+subrack+slot+puerto son iguales
                    f_port_txt = f_label['text'] if f_label and not is_annotation_label(f_label['text']) else ''
                    t_port_txt = t_label['text'] if t_label and not is_annotation_label(t_label['text']) else ''
                    if f_info and t_info and f_info == t_info and f_port_txt == t_port_txt:
                        continue

                    # ── Resolución de endpoint FROM ──
                    from_ep = None
                    if f_board:
                        parsed = parse_board_text(f_board['text'])
                        if parsed:
                            from_ep = format_endpoint("F:", parsed, f_port_txt, site_norm)
                    elif f_label and is_odf_shape(f_label['text']):
                        from_ep = format_odf_endpoint("F:", f_label['text'], site_norm)

                    # ── Resolución de endpoint TO ──
                    to_ep = None
                    if t_board:
                        parsed = parse_board_text(t_board['text'])
                        if parsed:
                            to_ep = format_endpoint("T:", parsed, t_port_txt, site_norm)
                    elif t_label and is_odf_shape(t_label['text']):
                        to_ep = format_odf_endpoint("T:", t_label['text'], site_norm)

                    if from_ep and to_ep and from_ep != to_ep:
                        extracted.append({'FROM:': from_ep, 'TO:': to_ep})


    except zipfile.BadZipFile:
        print(f"[ERROR] No es un ZIP válido: {file_path}")
    except Exception as e:
        import traceback
        print(f"[ERROR] {e}")
        traceback.print_exc()

    # Deduplicar
    seen = set()
    unique = []
    for c in extracted:
        key = (c['FROM:'], c['TO:'])
        if key not in seen:
            seen.add(key)
            unique.append(c)

    print(f"\n[RESULT] Total conexiones únicas: {len(unique)}")
    return unique


def generate_excel(connections: list, output_path: str) -> int:
    """Guarda las conexiones en Excel con columnas FROM: y TO:."""
    if not connections:
        print("[WARN] Sin conexiones para guardar.")
        return 0
    df = pd.DataFrame(connections, columns=['FROM:', 'TO:'])
    df.drop_duplicates(inplace=True)
    os.makedirs(os.path.dirname(output_path), exist_ok=True) if os.path.dirname(output_path) else None
    df.to_excel(output_path, index=False)
    print(f"[OK] Excel guardado en: {output_path} ({len(df)} filas)")
    return len(df)
