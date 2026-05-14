from PIL import Image
import os

for logo in ["claro", "movistar", "wom", "entel", "tigo"]:
    path = f"resources/{logo}.png"
    if os.path.exists(path):
        img = Image.open(path)
        # Check pixel (0,0) transparency
        pixel = img.getpixel((0,0))
        print(f"{logo}: {pixel}")
    else:
        print(f"{logo}: Not found")
