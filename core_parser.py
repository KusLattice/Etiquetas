import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re

LABEL_THRESHOLD = 1.0
NS = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}

def _cell(shape, name, default=0.0):
    c = shape.find(f'./v:Cell[@N="{name}"]', NS)
    if c is not None:
        try: return float(c.attrib.get('V', default))
        except: return default
    return default

def _cell_exists(shape, name):
    return shape.find(f'./v:Cell[@N="{name}"]', NS) is not None

def _text(shape):
    elem = shape.find('.//v:Text', NS)
    if elem is None: return ''
    return ''.join(elem.itertext()).strip()

def find_closest_board(x, y, page, boards):
    best, min_dist = None, 9999.0
    for b in boards:
        if b['page'] != page: continue
        if abs(b['y'] - y) <= b['h'] / 2 + 0.5:
            dist = abs(b['x'] - x)
            if dist < min_dist:
                min_dist = dist
                best = b
    return best

def find_closest_label(x, y, page, labels, threshold=LABEL_THRESHOLD):
    best, min_dist = None, 9999.0
    for l in labels:
        if l['page'] != page: continue
        dist = ((l['x'] - x) ** 2 + (l['y'] - y) ** 2) ** 0.5
        if dist < min_dist:
            min_dist = dist
            best = l
    return best if min_dist <= threshold else None

class BaseParser:
    def build_endpoint(self, prefix, board, label, site_name):
        raise NotImplementedError

class ClaroVTRParser(BaseParser):
    METADATA_PATTERNS = [
        r'^\d+(\.\d+)?KM', r'^V\d+\.\d+', r'^\d+\)?$', r'THZ',
        r'DBM', r'\d+\.\d+GHZ', r'^\+?-?\d+(\.\d+)?$',
        r'^APD$|^PIN$', r'150KM|80KM', r'VERSION|SOFTWARE|RESERVA',
        r'^\(?\d+X\d+\)?$',
    ]

    def is_metadata(self, text):
        if not text or len(text) < 2: return True
        t = text.upper().strip().replace(' ', '').replace('\n', '')
        for pat in self.METADATA_PATTERNS:
            if re.search(pat, t): return True
        return False

    def normalize_board_name(self, raw):
        b = raw.strip().upper()
        b = re.sub(r'^(TNG2|G2)', '', b)
        b = re.sub(r'(?<![A-Z0-9])01$|00$', '', b)
        b = re.sub(r'S\d+$', '', b)
        b = re.sub(r'F\d+$', '', b)
        return b

    def parse_board_text(self, text):
        clean = re.sub(r'\s+', '', text).upper()
        m = re.search(r'([A-Z0-9]+)\((\d+)[_\-]([A-Z])(\d+)\)', clean)
        if m:
            board_raw, rack, sub_letter, slot_digits = m.groups()
            board = self.normalize_board_name(board_raw)
            subrack = f"{rack}{sub_letter}"
            slot = f"({slot_digits.zfill(2)})"
            return board, subrack, slot
        return None, None, None

    def clean_port(self, text):
        if not text: return None
        t = text.strip().replace('\n', ' ')
        t = re.sub(r'\)+$', '', t)
        t = t.strip('-_ ')
        return t if t else None

    def build_odf_endpoint(self, prefix, label_text):
        t = re.sub(r'^(From|To)\s+', '', label_text.strip(), flags=re.IGNORECASE)
        m = re.match(r'^([^_]+(?:_[^_]+)*?)_(?:G2|P2|TNG2|LIN|LOUT)', t, re.IGNORECASE)
        if m:
            site_raw = m.group(1).strip()
            site_normalized = site_raw.replace(' ', '_').upper()
            return f"{prefix}ODF_{site_normalized}"
        return None

    def build_endpoint(self, prefix, board, label, site_name):
        if board is None and label is not None:
            ltext = label['text']
            if re.match(r'^(From|To)\s+', ltext, re.IGNORECASE):
                odf = self.build_odf_endpoint(prefix, ltext)
                if odf: return {'full_path': odf, 'fingerprint': odf}
            return None
        if board is None: return None
        board_name, subrack, slot = self.parse_board_text(board['text'])
        if not board_name or not subrack or not slot: return None
        port = self.clean_port(label['text']) if label else 'UNKNOWN'
        if port and self.is_metadata(port): port = 'UNKNOWN'
        site_parts = site_name.replace('_', ' ').split()
        site_tag = site_parts[0][:4] + '_VTR' if site_parts else 'SITIO_VTR'
        full_path = f"{prefix}RD_{site_tag}_{subrack}{slot}-{board_name}-{port}"
        fingerprint = f"{subrack}-{slot}-{board_name}"
        return {'full_path': full_path, 'fingerprint': fingerprint}

def parse_vsdx_connections(file_path, telco_name="ClaroVTR"):
    parser_map = {"ClaroVTR": ClaroVTRParser()}
    parser = parser_map.get(telco_name, ClaroVTRParser())
    extracted = []
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            pages = sorted([f for f in z.namelist()
                if 'visio/pages/page' in f and f.endswith('.xml') and 'rels' not in f])
            for page_path in pages:
                root = ET.fromstring(z.read(page_path))
                boards, labels, lines = [], [], []
                site_name = "SITIO"
                for shape in root.findall('.//v:Shape', NS):
                    txt = _text(shape)
                    px, py = _cell(shape, 'PinX'), _cell(shape, 'PinY')
                    w, h = _cell(shape, 'Width'), _cell(shape, 'Height')
                    if 'VTR' in txt.upper() and 5 < len(txt) < 40:
                        site_name = txt.strip()
                    if _cell_exists(shape, 'BeginX'):
                        lines.append({'bx': _cell(shape,'BeginX'), 'by': _cell(shape,'BeginY'),
                                      'ex': _cell(shape,'EndX'),   'ey': _cell(shape,'EndY'),
                                      'page': page_path})
                    elif h > 1.0 and w < 1.0 and txt:
                        boards.append({'x':px,'y':py,'w':w,'h':h,'text':txt,'page':page_path})
                    elif txt:
                        labels.append({'x':px,'y':py,'text':txt,'page':page_path})
                print(f"[{page_path}] boards={len(boards)} labels={len(labels)} lines={len(lines)} site='{site_name}'")
                for line in lines:
                    page = line['page']
                    board_f = find_closest_board(line['bx'], line['by'], page, boards)
                    label_f = find_closest_label(line['bx'], line['by'], page, labels)
                    board_t = find_closest_board(line['ex'], line['ey'], page, boards)
                    label_t = find_closest_label(line['ex'], line['ey'], page, labels)
                    has_from = board_f is not None or (label_f and re.match(r'^(From|To)\s+', label_f['text'], re.IGNORECASE))
                    has_to   = board_t is not None or (label_t and re.match(r'^(From|To)\s+', label_t['text'], re.IGNORECASE))
                    if not has_from or not has_to: continue
                    f_info = parser.build_endpoint("F:", board_f, label_f, site_name)
                    t_info = parser.build_endpoint("T:", board_t, label_t, site_name)
                    if not f_info or not t_info: continue
                    if f_info['fingerprint'] == t_info['fingerprint']: continue
                    extracted.append({'FROM:DE': f_info['full_path'], 'TO:HASTA': t_info['full_path']})
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback; traceback.print_exc()
    return extracted

def generate_excel(connections, output_path):
    if not connections: return 0
    df = pd.DataFrame(connections)
    df.drop_duplicates(inplace=True)
    df.to_excel(output_path, index=False)
    print(f"[Excel] {len(df)} filas → {output_path}")
    return len(df)
