import sys, zipfile, re
import xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding='utf-8')

path = r'C:\Users\kus\OneDrive\Escritorio\HUB Claro Maipú\Visio\FiberConnectionDiagram_MAIPU VTR.vsdx'
NS = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}

def get_text(shape):
    elem = shape.find('.//v:Text', NS)
    if elem is None: return ''
    raw = ''.join(elem.itertext())
    return re.sub(r'(?<=\S)\s+(?=\S)', '', raw).strip()

def get_num(shape, cell_name, default=0.0):
    cell = shape.find(f'.//v:Cell[@N="{cell_name}"]', NS)
    if cell is not None:
        try: return float(cell.attrib.get('V', default))
        except: pass
    return default

def get_none(shape, cell_name):
    cell = shape.find(f'.//v:Cell[@N="{cell_name}"]', NS)
    if cell is not None:
        try: return float(cell.attrib.get('V', 0))
        except: pass
    return None

for page_num in [2, 3]:
    target = f'visio/pages/page{page_num}.xml'
    boards, labels, lines = [], [], []

    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read(target))
        for shape in root.findall('.//v:Shape', NS):
            text = get_text(shape)
            pinx, piny = get_num(shape,'PinX'), get_num(shape,'PinY')
            w, h = get_num(shape,'Width'), get_num(shape,'Height')
            bx, ex = get_none(shape,'BeginX'), get_none(shape,'EndX')
            if bx is not None and ex is not None:
                lines.append({'bx':bx,'by':get_num(shape,'BeginY'),'ex':ex,'ey':get_num(shape,'EndY')})
            elif h > 0.4 and h > w * 1.5 and text:
                boards.append({'x':pinx,'y':piny,'w':w,'h':h,'text':text})
            elif text:
                labels.append({'x':pinx,'y':piny,'text':text})

    print(f'\n{"="*60}')
    print(f'PAGE{page_num}: {len(boards)} boards, {len(labels)} labels, {len(lines)} lines')
    print('BOARDS:')
    for b in boards:
        print(f'  {b["text"][:45]:45s} @ ({b["x"]:.3f},{b["y"]:.3f}) H={b["h"]:.3f}')

    print('ALL LABELS:')
    for l in sorted(labels, key=lambda x: x['y'], reverse=True):
        print(f'  {l["text"][:55]:55s} @ ({l["x"]:.3f},{l["y"]:.3f})')

    print('FIRST 15 LINES:')
    for ln in lines[:15]:
        # Para cada extremo, buscar label más cercano
        def nearest(x, y, threshold=9999):
            best, md = None, threshold
            for l in labels:
                d = ((l['x']-x)**2+(l['y']-y)**2)**0.5
                if d < md:
                    md = d
                    best = l
            return best, md
        lb, db = nearest(ln['bx'], ln['by'])
        le, de = nearest(ln['ex'], ln['ey'])
        print(f'  ({ln["bx"]:.2f},{ln["by"]:.2f})->({ln["ex"]:.2f},{ln["ey"]:.2f})')
        print(f'    FROM label: {lb["text"][:30] if lb else "None"} d={db:.3f}')
        print(f'    TO   label: {le["text"][:30] if le else "None"} d={de:.3f}')
