from PIL import Image
import os

def make_transparent(src, dest):
    print(f"Processing {src} -> {dest}")
    img = Image.open(src).convert("RGBA")
    datas = img.getdata()

    new_data = []
    for item in datas:
        # If the pixel is very white, make it transparent
        # We use a threshold to handle slight variations in white
        if item[0] > 230 and item[1] > 230 and item[2] > 230:
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)

    img.putdata(new_data)
    
    # Crop to content
    bbox = img.getbbox()
    if bbox:
        img = img.crop(bbox)
        
    img.save(dest, "PNG")

logos = {
    "claro": r"C:\Users\kus\.gemini\antigravity\brain\977dd3f3-804f-465a-a96d-76eccfbb0d3f\claro_white_bg_1778784668401.png",
    "movistar": r"C:\Users\kus\.gemini\antigravity\brain\977dd3f3-804f-465a-a96d-76eccfbb0d3f\movistar_white_bg_1778784680331.png",
    "wom": r"C:\Users\kus\.gemini\antigravity\brain\977dd3f3-804f-465a-a96d-76eccfbb0d3f\wom_white_bg_1778784703245.png",
    "entel": r"C:\Users\kus\.gemini\antigravity\brain\977dd3f3-804f-465a-a96d-76eccfbb0d3f\entel_white_bg_1778784720261.png",
    "tigo": r"C:\Users\kus\.gemini\antigravity\brain\977dd3f3-804f-465a-a96d-76eccfbb0d3f\tigo_white_bg_1778784739410.png"
}

os.makedirs("resources", exist_ok=True)

for name, path in logos.items():
    if os.path.exists(path):
        make_transparent(path, f"resources/{name}.png")
    else:
        print(f"Skipping {name}, path not found: {path}")
