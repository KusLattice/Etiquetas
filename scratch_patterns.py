import zipfile
import xml.etree.ElementTree as ET

path = r'C:/Users/kus/OneDrive/Escritorio/HUB Claro Maipú/Visio/FiberConnectionDiagram_MAIPU VTR.vsdx'
ns = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}

def get_text(shape):
    elem = shape.find('.//v:Text', ns)
    if elem is None: return ''
    return ''.join(elem.itertext()).strip()

with zipfile.ZipFile(path, 'r') as z:
    for page_name in [f for f in z.namelist() if 'visio/pages/page' in f and f.endswith('.xml') and not 'rels' in f]:
        print(f'\n--- PAGE: {page_name} ---')
        root = ET.fromstring(z.read(page_name))
        for shape in root.findall('.//v:Shape', ns):
            text = get_text(shape).replace('\n', ' ')
            if text:
                # Print shape ID and text if it looks relevant
                if len(text) > 2:
                    print(f'[{shape.attrib.get("ID")}] {text}')
