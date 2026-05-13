import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os

class BaseParser:
    def format_endpoint(self, prefix, board_data):
        raise NotImplementedError("Debe implementarse en la subclase")

class ClaroVTRParser(BaseParser):
    def __init__(self):
        # FILTRO 2: Metadatos técnicos a ignorar (no son conexiones físicas)
        self.metadata_blacklist = [
            r"\d+KM",          # ej: 150km, 80km
            r"V\d+\.\d+",      # ej: V3.1, V4.2
            r"^\d+\)?$",       # ej: 1, 1), 2)
            r"VERSION", 
            r"SOFTWARE",
            r"\(?\d+X\d+\)?",  # ej: (10X10)
            r"RESERVA"
        ]

    def clean_string(self, text):
        """FILTRO 3: Saneamiento de strings y eliminación de paréntesis huérfanos"""
        if not text: return ""
        # Limpieza de saltos de línea y espacios
        text = text.strip().replace('\n', ' ')
        # ELIMINAR PARÉNTESIS HUÉRFANOS AL FINAL: ej: "OH9S-1)" -> "OH9S-1"
        text = re.sub(r"\)+$", "", text)
        # Limpiar caracteres especiales residuales al inicio/final
        text = text.strip("-_ ")
        return text

    def is_metadata(self, text):
        """FILTRO 2: Discriminador de etiquetas de anotación vs puertos reales"""
        if not text or len(text) < 2: return True
        clean_text = text.upper().strip()
        for pattern in self.metadata_blacklist:
            if re.search(pattern, clean_text):
                return True
        return False

    def parse_hardware_info(self, board_text):
        """Normalización estricta de Hardware (Board, Subrack, Slot)"""
        clean_text = re.sub(r"\s+", "", board_text).upper()
        
        # Pattern: Identifica Board y el bloque (Subrack_Slot)
        match = re.search(r"([^()]+)\((\d+)[_-]([A-Z0-9]+)\)", clean_text)
        if match:
            board_raw, sub_num, slot_raw = match.groups()
            
            # Limpieza de Board según norma: Quitar G2/TNG2 y sufijos Sxx/01/00
            board = re.sub(r"^(TNG2|G2)", "", board_raw)
            board = re.sub(r"S\d+$|01$|00$", "", board)
            
            # Normalización de Slot y Subrack
            sub_let_match = re.search(r"([A-Z])", slot_raw)
            sub_let = sub_let_match.group(1) if sub_let_match else ""
            slot_num_match = re.search(r"(\d+)$", slot_raw)
            slot_num = slot_num_match.group(1).zfill(2) if slot_num_match else "00"
            
            return board, f"{sub_num}{sub_let}", f"({slot_num})"
        
        return board_text, "??", "(??)"

    def format_endpoint(self, prefix, board_data):
        """Aplica filtros de depuración y genera el endpoint final"""
        site = board_data.get('site', 'SITIO').split()[0]
        raw_text = board_data.get('text', '')
        
        # FILTRO 2: Validar si es una anotación de texto metadata
        if self.is_metadata(raw_text):
            return None

        clean_text = self.clean_string(raw_text)
        
        # Segmentación Board vs Puerto
        parenthesis_end = clean_text.find(")")
        if parenthesis_end != -1:
            raw_board = clean_text[:parenthesis_end+1]
            port = self.clean_string(clean_text[parenthesis_end+1:])
        else:
            raw_board = clean_text
            port = "PORT"

        board, subrack, slot = self.parse_hardware_info(raw_board)
        
        # Si falló el parseo de hardware crítico, marcamos para revisión
        if subrack == "??" or slot == "(??)":
            return None

        return {
            'full_path': f"{prefix}RD_{site}_VTR_{subrack}{slot}-{board}-{port}",
            'fingerprint': f"{subrack}-{slot}-{board}" # Para FILTRO 1 (Loopbacks)
        }

def parse_vsdx_connections(file_path, telco_name="ClaroVTR"):
    """Motor de extracción VSDX con Pipeline de Depuración Integrado"""
    parser_map = {
        "ClaroVTR": ClaroVTRParser(),
    }
    parser = parser_map.get(telco_name, ClaroVTRParser())
    
    extracted_connections = []
    
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            pages = [f for f in z.namelist() if f.startswith('visio/pages/page') and f.endswith('.xml')]
            for page_path in pages:
                with z.open(page_path) as f:
                    root = ET.fromstring(f.read())
                    ns = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}
                    
                    shapes = {}
                    lines = []
                    site_name = "SITIO"
                    
                    # 1. Escaneo de Shapes y Detección de Sitio
                    for shape in root.findall('.//v:Shape', ns):
                        sid = shape.get('ID')
                        name = shape.get('NameU', '')
                        
                        # Extraer coordenadas
                        px = float(shape.find('./v:Cell[@N="PinX"]', ns).get('V')) if shape.find('./v:Cell[@N="PinX"]', ns) is not None else 0
                        py = float(shape.find('./v:Cell[@N="PinY"]', ns).get('V')) if shape.find('./v:Cell[@N="PinY"]', ns) is not None else 0
                        
                        # Extraer texto
                        txt_elem = shape.find('./v:Text', ns)
                        txt = "".join(txt_elem.itertext()).strip() if txt_elem is not None else ""
                        
                        # Detectar Nombre de Sitio (el texto más grande o con patrón VTR)
                        if "VTR" in txt.upper() and len(txt) > len(site_name):
                            site_name = txt.replace(" ", "_").upper()

                        # Detección de conectores: Buscamos si tiene celdas de inicio/fin (BeginX/EndX)
                        # Esto es mucho más robusto que buscar por nombre (NameU)
                        begin_x_elem = shape.find('./v:Cell[@N="BeginX"]', ns)
                        if begin_x_elem is not None:
                            bx = float(begin_x_elem.get('V'))
                            by = float(shape.find('./v:Cell[@N="BeginY"]', ns).get('V')) if shape.find('./v:Cell[@N="BeginY"]', ns) is not None else 0
                            ex = float(shape.find('./v:Cell[@N="EndX"]', ns).get('V')) if shape.find('./v:Cell[@N="EndX"]', ns) is not None else 0
                            ey = float(shape.find('./v:Cell[@N="EndY"]', ns).get('V')) if shape.find('./v:Cell[@N="EndY"]', ns) is not None else 0
                            lines.append({'bx': bx, 'by': by, 'ex': ex, 'ey': ey})
                        else:
                            shapes[sid] = {'id': sid, 'x': px, 'y': py, 'text': txt}

                    # 2. Análisis de Proximidad y Aplicación de Filtros
                    for line in lines:
                        from_shape = find_closest_shape(line['bx'], line['by'], shapes)
                        to_shape = find_closest_shape(line['ex'], line['ey'], shapes)
                        
                        if from_shape and to_shape and from_shape['id'] != to_shape['id']:
                            from_shape['site'] = site_name
                            to_shape['site'] = site_name
                            
                            f_info = parser.format_endpoint("F:", from_shape)
                            t_info = parser.format_endpoint("T:", to_shape)
                            
                            # FILTRO 2: Ignorar si es metadato o error de parseo
                            if f_info and t_info:
                                # FILTRO 1: Ignorar Loopbacks (Patchcord a la misma tarjeta/slot)
                                if f_info['fingerprint'] != t_info['fingerprint']:
                                    extracted_connections.append({
                                        'FROM:DE': f_info['full_path'],
                                        'TO:HASTA': t_info['full_path']
                                    })
                                    
    except Exception as e:
        print(f"Error en parser core: {e}")

    return extracted_connections

def find_closest_shape(x, y, shapes, threshold=2.0):
    """Busca el objeto de texto más cercano a un extremo de línea"""
    closest = None
    min_dist = threshold
    for sid, s in shapes.items():
        if not s['text'] or len(s['text']) < 2: continue
        dist = ((x - s['x'])**2 + (y - s['y'])**2)**0.5
    return closest

def generate_excel(connections, output_path):
    """Generación de Excel con Filtro 4 (Deduplicación)"""
    if not connections: return 0
    
    df = pd.DataFrame(connections)
    
    # FILTRO 4: Eliminar duplicados exactos (redundancia de líneas)
    initial_count = len(df)
    df.drop_duplicates(inplace=True)
    final_count = len(df)
    
    df.to_excel(output_path, index=False)
    return final_count
