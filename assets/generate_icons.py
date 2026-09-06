#!/usr/bin/env python3
"""Generate the two macOS application iconsets used by the build scripts."""

from pathlib import Path
import shutil
import subprocess

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent


def rounded(draw: ImageDraw.ImageDraw, box, radius, fill):
    draw.rounded_rectangle(box, radius=radius, fill=fill)


def make_icon(kind: str, output: Path) -> None:
    scale = 4
    size = 1024
    image = Image.new("RGBA", (size * scale, size * scale), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    def box(values):
        return tuple(int(value * scale) for value in values)

    # macOS-style dark tile with a subtle inner border.
    rounded(draw, box((42, 42, 982, 982)), 210 * scale, (20, 27, 43, 255))
    rounded(draw, box((58, 58, 966, 966)), 194 * scale, (31, 42, 65, 255))
    rounded(draw, box((72, 72, 952, 952)), 180 * scale, (38, 52, 80, 255))

    if kind == "export":
        # Spreadsheet panel.
        rounded(draw, box((190, 180, 834, 720)), 56 * scale, (232, 244, 255, 255))
        rounded(draw, box((190, 180, 834, 300)), 56 * scale, (63, 188, 221, 255))
        draw.rectangle(box((190, 246, 834, 300)), fill=(63, 188, 221, 255))
        for x in (352, 512, 672):
            draw.line((x * scale, 300 * scale, x * scale, 720 * scale), fill=(158, 190, 211, 255), width=10 * scale)
        for y in (405, 510, 615):
            draw.line((190 * scale, y * scale, 834 * scale, y * scale), fill=(158, 190, 211, 255), width=10 * scale)
        # Downward export arrow.
        draw.line((512 * scale, 430 * scale, 512 * scale, 826 * scale), fill=(255, 196, 76, 255), width=54 * scale)
        draw.polygon([box((398, 735)), box((626, 735)), box((512, 870))], fill=(255, 196, 76, 255))
    else:
        # Cloud silhouette.
        rounded(draw, box((222, 430, 802, 730)), 110 * scale, (221, 233, 255, 255))
        draw.ellipse(box((270, 300, 550, 610)), fill=(221, 233, 255, 255))
        draw.ellipse(box((430, 250, 730, 650)), fill=(221, 233, 255, 255))
        draw.rectangle(box((340, 510, 680, 700)), fill=(221, 233, 255, 255))
        # Upward upload arrow.
        draw.line((512 * scale, 810 * scale, 512 * scale, 420 * scale), fill=(111, 229, 177, 255), width=54 * scale)
        draw.polygon([box((398, 500)), box((626, 500)), box((512, 360))], fill=(111, 229, 177, 255))

    image = image.resize((size, size), Image.Resampling.LANCZOS)
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output, format="PNG", optimize=True)


def make_icns(png_path: Path, icns_path: Path) -> None:
    iconset = icns_path.with_suffix(".iconset")
    if iconset.exists():
        shutil.rmtree(iconset)
    iconset.mkdir(parents=True)
    source = Image.open(png_path)
    for size in (16, 32, 128, 256, 512, 1024):
        source.resize((size, size), Image.Resampling.LANCZOS).save(iconset / f"icon_{size}x{size}.png")
        if size <= 512:
            source.resize((size * 2, size * 2), Image.Resampling.LANCZOS).save(iconset / f"icon_{size}x{size}@2x.png")
    subprocess.run(["/usr/bin/iconutil", "-c", "icns", str(iconset), "-o", str(icns_path)], check=True)
    shutil.rmtree(iconset)


def main() -> None:
    for kind, name in (("export", "exporter"), ("upload", "uploader")):
        png = ROOT / f"{name}.png"
        icns = ROOT / f"{name}.icns"
        make_icon(kind, png)
        make_icns(png, icns)
        print(icns)


if __name__ == "__main__":
    main()
