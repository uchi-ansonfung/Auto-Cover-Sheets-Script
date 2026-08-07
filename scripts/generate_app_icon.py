#!/usr/bin/env python3
"""
Generate Automatic Exhibit Cover Sheets app icons (PNG + multi-size ICO).

Requires Pillow:  pip install Pillow

Usage (from repo root):
  python scripts/generate_app_icon.py
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

REPO_ROOT = Path(__file__).resolve().parents[1]
ASSETS = REPO_ROOT / "assets"
MASTER_SIZE = 1024
ICO_SIZES = (16, 24, 32, 48, 64, 128, 256)

# Match CustomTkinter default blue theme used by the GUI.
BLUE_DEEP = (31, 106, 165)  # #1F6AA5
BLUE_MID = (59, 142, 208)  # #3B8ED0
BLUE_LIGHT = (120, 180, 230)
WHITE = (255, 255, 255)
PAGE_EDGE = (220, 228, 236)
SHADOW = (12, 40, 70)


def _rounded_rect(
    draw: ImageDraw.ImageDraw,
    box: tuple[float, float, float, float],
    radius: float,
    fill: tuple[int, ...],
    outline: tuple[int, ...] | None = None,
    width: int = 1,
) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def _page_box(
    cx: float,
    cy: float,
    page_w: float,
    page_h: float,
    dx: float = 0.0,
    dy: float = 0.0,
) -> tuple[float, float, float, float]:
    left = cx - page_w / 2 + dx
    top = cy - page_h / 2 + dy
    return (left, top, left + page_w, top + page_h)


def _draw_page_shadow(
    canvas: Image.Image,
    box: tuple[float, float, float, float],
    radius: float,
    blur: float,
    alpha: int,
    offset: tuple[float, float] = (0, 0),
) -> None:
    """Soft drop shadow for a rounded page rectangle."""
    layer = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    ox, oy = offset
    shifted = (box[0] + ox, box[1] + oy, box[2] + ox, box[3] + oy)
    _rounded_rect(draw, shifted, radius, fill=(*SHADOW, alpha))
    if blur > 0:
        layer = layer.filter(ImageFilter.GaussianBlur(radius=blur))
    canvas.alpha_composite(layer)


def draw_master(size: int = MASTER_SIZE) -> Image.Image:
    """Draw the stacked cover-sheet icon at the given square size."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    s = float(size)

    # App tile background (rounded square).
    margin = s * 0.06
    tile = (margin, margin, s - margin, s - margin)
    tile_r = s * 0.18
    _rounded_rect(draw, tile, tile_r, fill=(*BLUE_DEEP, 255))

    # Subtle top highlight on the tile.
    highlight = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    hdraw = ImageDraw.Draw(highlight)
    mid_y = s * 0.42
    hdraw.rounded_rectangle(
        (margin, margin, s - margin, mid_y + tile_r),
        radius=tile_r,
        fill=(*BLUE_MID, 70),
    )
    # Clip highlight to upper portion with a soft fade mask.
    fade = Image.new("L", (size, size), 0)
    fdraw = ImageDraw.Draw(fade)
    for y in range(int(margin), int(mid_y)):
        t = (y - margin) / max(mid_y - margin, 1)
        a = int(255 * (1.0 - t * t))
        fdraw.line([(0, y), (size, y)], fill=a)
    highlight.putalpha(
        Image.composite(
            fade,
            Image.new("L", (size, size), 0),
            highlight.split()[3],
        )
    )
    img.alpha_composite(highlight)

    # Letter-page proportions (~8.5 x 11).
    page_h = s * 0.58
    page_w = page_h * (8.5 / 11.0)
    cx, cy = s * 0.50, s * 0.52
    page_r = max(s * 0.035, 2.0)

    # Back → mid → front offsets (stacked, slightly fanned).
    layers = (
        {"dx": -s * 0.055, "dy": s * 0.045, "shadow_a": 55},
        {"dx": s * 0.010, "dy": s * 0.012, "shadow_a": 70},
        {"dx": s * 0.055, "dy": -s * 0.035, "shadow_a": 90},
    )

    for layer in layers:
        box = _page_box(cx, cy, page_w, page_h, layer["dx"], layer["dy"])
        _draw_page_shadow(
            img,
            box,
            page_r,
            blur=s * 0.018,
            alpha=layer["shadow_a"],
            offset=(s * 0.012, s * 0.018),
        )

    draw = ImageDraw.Draw(img)
    for i, layer in enumerate(layers):
        box = _page_box(cx, cy, page_w, page_h, layer["dx"], layer["dy"])
        is_front = i == len(layers) - 1
        fill = WHITE if is_front else (248, 250, 252)
        _rounded_rect(draw, box, page_r, fill=fill, outline=PAGE_EDGE, width=max(1, int(s * 0.006)))

        if is_front:
            # Abstract centered "title" block — no letters.
            left, top, right, bottom = box
            pw, ph = right - left, bottom - top
            line_w = pw * 0.55
            line_h = max(ph * 0.045, 2.0)
            lx0 = left + (pw - line_w) / 2
            ly0 = top + ph * 0.42
            _rounded_rect(
                draw,
                (lx0, ly0, lx0 + line_w, ly0 + line_h),
                line_h / 2,
                fill=(*BLUE_MID, 255),
            )
            # Second thinner rule beneath (cover-sheet title wrap feel).
            line_w2 = pw * 0.38
            line_h2 = max(ph * 0.028, 1.5)
            lx1 = left + (pw - line_w2) / 2
            ly1 = ly0 + line_h * 2.2
            _rounded_rect(
                draw,
                (lx1, ly1, lx1 + line_w2, ly1 + line_h2),
                line_h2 / 2,
                fill=(*BLUE_LIGHT, 220),
            )
            # Thin blue header strip at top of front cover.
            strip_h = max(ph * 0.07, 2.0)
            strip_inset = pw * 0.08
            _rounded_rect(
                draw,
                (
                    left + strip_inset,
                    top + ph * 0.08,
                    right - strip_inset,
                    top + ph * 0.08 + strip_h,
                ),
                strip_h / 2,
                fill=(*BLUE_DEEP, 200),
            )

    return img


def _resize_preview(master: Image.Image, size: int) -> Image.Image:
    """Downscale for QA previews; slight sharpen at tiny sizes."""
    out = master.resize((size, size), Image.Resampling.LANCZOS)
    if size <= 32:
        out = out.filter(ImageFilter.UnsharpMask(radius=0.6, percent=120, threshold=2))
    return out


def main() -> None:
    ASSETS.mkdir(parents=True, exist_ok=True)
    master = draw_master(MASTER_SIZE)

    png_path = ASSETS / "app-icon.png"
    master.save(png_path, format="PNG", optimize=True)
    print(f"Wrote {png_path.relative_to(REPO_ROOT)} ({MASTER_SIZE}x{MASTER_SIZE})")

    ico_path = ASSETS / "app-icon.ico"
    # Pass the full-res master + sizes list so Pillow embeds every resolution.
    # (Saving a pre-resized 16px image with sizes= only keeps that one size.)
    master.save(ico_path, format="ICO", sizes=[(n, n) for n in ICO_SIZES])
    print(f"Wrote {ico_path.relative_to(REPO_ROOT)} sizes={list(ICO_SIZES)}")

    # Small preview thumbs for visual QA (not packaged; gitignored).
    preview_dir = ASSETS / "_icon_preview"
    preview_dir.mkdir(exist_ok=True)
    for n in (16, 32, 64, 256):
        p = preview_dir / f"app-icon-{n}.png"
        _resize_preview(master, n).save(p, format="PNG")
    print(f"Wrote previews under {preview_dir.relative_to(REPO_ROOT)}/")


if __name__ == "__main__":
    main()
