from PIL import Image
import os

def process_logos():
    res_dir = "resources"
    for filename in os.listdir(res_dir):
        if filename.endswith(".png"):
            path = os.path.join(res_dir, filename)
            img = Image.open(path).convert("RGBA")
            datas = img.getdata()

            new_data = []
            for item in datas:
                # Remove white/near-white background
                if item[0] > 220 and item[1] > 220 and item[2] > 220:
                    new_data.append((255, 255, 255, 0))
                else:
                    new_data.append(item)

            img.putdata(new_data)
            # Crop to content if possible
            bbox = img.getbbox()
            if bbox:
                img = img.crop(bbox)
            
            img.save(path)
            print(f"Procesado: {filename}")

if __name__ == "__main__":
    process_logos()
