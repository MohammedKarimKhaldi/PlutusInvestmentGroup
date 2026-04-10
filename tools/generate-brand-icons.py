#!/usr/bin/env python3

from pathlib import Path
from shutil import rmtree
import subprocess
import sys

try:
    from PIL import Image, ImageDraw
except ImportError as exc:
    raise SystemExit(
        "Pillow is required to generate brand icons. "
        "Create a virtualenv and install pillow before running this script."
    ) from exc


ROOT = Path(__file__).resolve().parents[1]
RESOURCES_DIR = ROOT / "src-tauri" / "icons"
ICNS_ICONSET_DIR = RESOURCES_DIR / "icon.iconset"

NAVY = (20, 24, 95, 255)
WHITE = (255, 255, 255, 255)
TRANSPARENT = (0, 0, 0, 0)


def scale_points(points, size):
    factor = size / 1024.0
    return [(round(x * factor), round(y * factor)) for x, y in points]


def draw_mark(draw, size, fill):
    polygons = [
        [(358, 120), (616, 120), (760, 218), (660, 290), (468, 182)],
        [(300, 132), (430, 208), (330, 386), (236, 330)],
        [(230, 374), (350, 446), (350, 852), (230, 780)],
        [(376, 230), (514, 310), (514, 860), (376, 940)],
        [(664, 266), (760, 212), (760, 334), (664, 388)],
        [(548, 352), (662, 286), (662, 414), (548, 480)],
        [(494, 414), (584, 466), (504, 514), (416, 462)],
        [(416, 462), (504, 514), (504, 646), (416, 594)],
        [(504, 514), (760, 366), (760, 494), (588, 592), (504, 646)],
    ]

    for polygon in polygons:
        draw.polygon(scale_points(polygon, size), fill=fill)


def create_canvas(size, transparent=False):
    if transparent:
      return Image.new("RGBA", (size, size), TRANSPARENT)
    image = Image.new("RGBA", (size, size), WHITE)
    draw = ImageDraw.Draw(image)
    radius = round(size * 0.22)
    draw.rounded_rectangle((0, 0, size - 1, size - 1), radius=radius, fill=WHITE)
    return image


def generate_base_icons():
    RESOURCES_DIR.mkdir(parents=True, exist_ok=True)

    full_icon = create_canvas(1024, transparent=False)
    draw_full = ImageDraw.Draw(full_icon)
    draw_mark(draw_full, 1024, NAVY)
    full_icon.save(RESOURCES_DIR / "icon.png")
    full_icon.save(RESOURCES_DIR / "icon-square-1024.png")
    full_icon.save(
        RESOURCES_DIR / "icon.ico",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )

    foreground = create_canvas(1024, transparent=True)
    draw_foreground = ImageDraw.Draw(foreground)
    draw_mark(draw_foreground, 1024, NAVY)
    foreground.save(RESOURCES_DIR / "icon-foreground-1024.png")

    return full_icon, foreground


def generate_icns(full_icon):
    if ICNS_ICONSET_DIR.exists():
        rmtree(ICNS_ICONSET_DIR)
    ICNS_ICONSET_DIR.mkdir(parents=True)

    iconset_sizes = {
        "icon_16x16.png": 16,
        "icon_16x16@2x.png": 32,
        "icon_32x32.png": 32,
        "icon_32x32@2x.png": 64,
        "icon_128x128.png": 128,
        "icon_128x128@2x.png": 256,
        "icon_256x256.png": 256,
        "icon_256x256@2x.png": 512,
        "icon_512x512.png": 512,
        "icon_512x512@2x.png": 1024,
    }

    for filename, size in iconset_sizes.items():
        full_icon.resize((size, size), Image.Resampling.LANCZOS).save(ICNS_ICONSET_DIR / filename)

    subprocess.run(
        ["iconutil", "-c", "icns", str(ICNS_ICONSET_DIR), "-o", str(RESOURCES_DIR / "icon.icns")],
        check=True,
    )
    rmtree(ICNS_ICONSET_DIR)


def main():
    full_icon, _foreground = generate_base_icons()
    generate_icns(full_icon)
    print("Generated Plutus brand icons for Tauri.")


if __name__ == "__main__":
    sys.exit(main())
