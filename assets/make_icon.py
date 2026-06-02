"""
Generates assets/crm.ico — a simple MISA-branded icon.
Requires Pillow (pip install pillow).  Falls back to a minimal ICO stub if
Pillow is not installed so the shortcut creation still succeeds.
"""
import os
import struct

OUT = os.path.join(os.path.dirname(__file__), "crm.ico")

def _try_pillow():
    from PIL import Image, ImageDraw, ImageFont
    sizes = [256, 128, 64, 48, 32, 16]
    frames = []
    for sz in sizes:
        img = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        # Green rounded square background
        margin = max(1, sz // 16)
        d.rounded_rectangle(
            [margin, margin, sz - margin - 1, sz - margin - 1],
            radius=sz // 8,
            fill=(27, 92, 63, 255),          # #1B5C3F
        )
        # Gold "M" letter centred
        font_size = int(sz * 0.55)
        try:
            font = ImageFont.truetype("arialbd.ttf", font_size)
        except Exception:
            try:
                font = ImageFont.truetype("Arial_Bold.ttf", font_size)
            except Exception:
                font = ImageFont.load_default()
        text = "M"
        bbox = d.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        tx = (sz - tw) // 2 - bbox[0]
        ty = (sz - th) // 2 - bbox[1]
        d.text((tx, ty), text, font=font, fill=(201, 151, 74, 255))   # #C9974A gold
        frames.append(img)
    frames[0].save(OUT, format="ICO", sizes=[(s, s) for s in sizes],
                   append_images=frames[1:])
    print(f"Icon written: {OUT}")

def _minimal_ico():
    """Write a 16×16 solid green ICO without Pillow."""
    green = (63, 92, 27)   # BGR order for BMP
    px = bytes([green[0], green[1], green[2], 0] * 16 * 16)
    # BMP info header (40 bytes) for 16×16×32bpp
    bih = struct.pack("<IiiHHIIiiII", 40, 16, 32, 1, 32, 0,
                      len(px), 0, 0, 0, 0)
    bmp_data = bih + px
    # ICO header + 1 entry
    ico_header = struct.pack("<HHH", 0, 1, 1)
    dir_entry  = struct.pack("<BBBBHHII", 16, 16, 0, 0, 1, 32,
                             len(bmp_data), 6 + 16)
    with open(OUT, "wb") as f:
        f.write(ico_header + dir_entry + bmp_data)
    print(f"Minimal icon written: {OUT}")

if __name__ == "__main__":
    try:
        _try_pillow()
    except ImportError:
        _minimal_ico()
