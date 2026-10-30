"""
Icon generator script using Python Pillow (PIL).
Creates icon-16.png, icon-25.png, icon-32.png, icon-48.png, and icon-80.png in assets/
"""
import os
import sys

try:
    from PIL import Image, ImageDraw
except ImportError:
    print("Pillow is not installed. Run 'pip install pillow' or open 'generate-icons.html' in your browser.")
    sys.exit(0)

os.makedirs("assets", exist_ok=True)
sizes = [16, 25, 32, 48, 80]

for size in sizes:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Blue rounded rectangle background
    r = int(size * 0.18)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=(0, 120, 212, 255))

    # Lightning bolt polygon
    s = size
    points = [
        (s * 0.55, s * 0.18),
        (s * 0.32, s * 0.52),
        (s * 0.50, s * 0.52),
        (s * 0.44, s * 0.82),
        (s * 0.70, s * 0.44),
        (s * 0.52, s * 0.44),
    ]
    draw.polygon(points, fill=(255, 255, 255, 255))

    out_path = os.path.join("assets", f"icon-{size}.png")
    img.save(out_path, "PNG")
    print(f"Generated: {out_path}")

print("Done! All icons created.")
