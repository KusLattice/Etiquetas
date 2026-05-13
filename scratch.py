import zipfile, re
import xml.etree.ElementTree as ET

path = r'C:/Users/kus/OneDrive/Escritorio/HUB Claro Maipú/Visio/FiberConnectionDiagram_MAIPU VTR.vsdx'
ns = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}

def get_num(shape, cell_name, default=0.0):
    cell = shape.find(f'.//v:Cell[@N="{cell_name}"]', ns)
    if cell is not None:
        try:
            return float(cell.attrib.get('V', default))
        except:
            return default
    return default

def get_text(shape):
    elem = shape.find('.//v:Text', ns)
    if elem is None: return ''
    # Filter out text from sub-elements like cp, pp unless it's direct text
    return ''.join(elem.itertext()).strip()

boards = []
labels = []
lines = []

with zipfile.ZipFile(path, 'r') as z:
    for page_name in [f for f in z.namelist() if 'visio/pages/page' in f and f.endswith('.xml') and not 'rels' in f]:
        root = ET.fromstring(z.read(page_name))
        for shape in root.findall('.//v:Shape', ns):
            text = get_text(shape)
            pinx = get_num(shape, 'PinX')
            piny = get_num(shape, 'PinY')
            w = get_num(shape, 'Width')
            h = get_num(shape, 'Height')
            bx = get_num(shape, 'BeginX', None)
            by = get_num(shape, 'BeginY', None)
            ex = get_num(shape, 'EndX', None)
            ey = get_num(shape, 'EndY', None)
            
            if bx is not None and ex is not None:
                lines.append({'bx':bx, 'by':by, 'ex':ex, 'ey':ey, 'page': page_name})
            elif h > 1.0 and w < 1.0 and text:
                boards.append({'x':pinx, 'y':piny, 'w':w, 'h':h, 'text':text, 'page': page_name})
            elif text:
                labels.append({'x':pinx, 'y':piny, 'text':text, 'page': page_name})

print(f'Found {len(lines)} lines, {len(boards)} boards, {len(labels)} labels.')

def find_closest_board(x, y, page):
    best = None
    min_dist = 9999
    for b in boards:
        if b['page'] != page: continue
        if abs(b['y'] - y) <= b['h']/2 + 0.5:
            dist = abs(b['x'] - x)
            if dist < min_dist:
                min_dist = dist
                best = b
    return best

def find_closest_label(x, y, page):
    best = None
    min_dist = 9999
    for l in labels:
        if l['page'] != page: continue
        dist = ((l['x'] - x)**2 + (l['y'] - y)**2)**0.5
        if dist < min_dist:
            min_dist = dist
            best = l
    return best if min_dist < 0.5 else None

for line in lines[:5]:
    b_left = find_closest_board(line['bx'], line['by'], line['page'])
    l_left = find_closest_label(line['bx'], line['by'], line['page'])
    b_right = find_closest_board(line['ex'], line['ey'], line['page'])
    l_right = find_closest_label(line['ex'], line['ey'], line['page'])
    
    print('LINE:')
    print('  Left:', b_left['text'].replace('\n','')[:20] if b_left else 'None', '|', l_left['text'] if l_left else 'None')
    print('  Right:', b_right['text'].replace('\n','')[:20] if b_right else 'None', '|', l_right['text'] if l_right else 'None')
