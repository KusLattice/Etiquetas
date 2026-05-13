import zipfile
import xml.etree.ElementTree as ET
import os

file_path = 'c:/Users/kus/OneDrive/Escritorio/HUB Claro Maipú/Visio/FiberConnectionDiagram_MAIPU VTR.vsdx'

def test():
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    ns = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}
    
    with zipfile.ZipFile(file_path, 'r') as z:
        pages = [f for f in z.namelist() if f.startswith('visio/pages/page') and f.endswith('.xml')]
        print(f"Pages found: {pages}")
        for page_path in pages:
            print(f"--- Checking {page_path} ---")
            with z.open(page_path) as f:
                content = f.read()
                root = ET.fromstring(content)
                
                all_shapes = root.findall('.//v:Shape', ns)
                print(f"Total shapes in XML: {len(all_shapes)}")
                
                for i, shape in enumerate(all_shapes):
                    sid = shape.get('ID')
                    cells = shape.findall('./v:Cell', ns)
                    cell_names = [c.get('N') for c in cells]
                    
                    if 'BeginX' in cell_names:
                        name = shape.get('NameU', 'NO_NAME')
                        txt_elem = shape.find('./v:Text', ns)
                        txt = "".join(txt_elem.itertext()).strip() if txt_elem is not None else ""
                        print(f"  FOUND CONNECTOR ID {sid}: NameU='{name}', Text='{txt}'")

if __name__ == "__main__":
    test()
