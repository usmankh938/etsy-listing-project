#!/usr/bin/env python3
"""Shop banner and icon for the Sable invoice shop.

The big banner is 3360×840, the size Etsy’s seller handbook calls optimal
for a big shop banner (4:1, same shape as the 1600×400 file Etsy Help
lists as the recommended upload). The icon is the 500×500 shop icon.

Run from the repo root:

    python3 etsy/make_banner.py
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parents[1]
FONTS = ROOT / "fonts"
PREVIEW = ROOT / "preview"
OUT = ROOT / "etsy"

PAPER = (231, 225, 214)  # #E7E1D6, same ground as the listing hero
FOREST = (30, 58, 50)  # #1E3A32
CREAM = (247, 244, 238)  # #F7F4EE
BRASS = (122, 98, 56)  # #7A6238
INK = (28, 25, 21)  # #1C1915

W, H = 3360, 840


def font(name: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONTS / name), size)


def text_top(draw, xy, text, face, fill, tracking=0):
    """Draw text so `xy` is the top-left of the ink, not the em box."""
    x, y = xy
    bbox = face.getbbox(text)
    baseline = y - bbox[1]
    if tracking == 0:
        draw.text((x, baseline), text, font=face, fill=fill)
        width = bbox[2] - bbox[0]
    else:
        cursor = x
        for ch in text:
            draw.text((cursor, baseline), ch, font=face, fill=fill)
            cursor += face.getlength(ch) + tracking
        width = cursor - tracking - x
    return width, bbox[3] - bbox[1]


def draw_meta(draw, x, y, items, face, fill):
    """Small brass labels separated by dots, with even gaps."""
    start = x
    gap = 16
    dot = 3
    height = face.getbbox(items[0])[3] - face.getbbox(items[0])[1]
    for index, item in enumerate(items):
        width, _ = text_top(draw, (x, y), item, face, fill, tracking=1.6)
        x += width
        if index == len(items) - 1:
            break
        cx = x + gap + dot / 2
        cy = y + height / 2 + 1
        draw.ellipse((cx - dot / 2, cy - dot / 2, cx + dot / 2, cy + dot / 2), fill=fill)
        x = cx + dot / 2 + gap
    return x - start


def ink_width(face, text, tracking=0):
    if tracking == 0:
        bbox = face.getbbox(text)
        return bbox[2] - bbox[0]
    return sum(face.getlength(ch) + tracking for ch in text) - tracking


def shadow_for(size, blur=28, alpha=78):
    w, h = size
    plate = Image.new("RGBA", (w + blur * 4, h + blur * 4), (0, 0, 0, 0))
    ImageDraw.Draw(plate).rectangle(
        (blur * 2, blur * 2, blur * 2 + w, blur * 2 + h),
        fill=(28, 25, 21, alpha),
    )
    return plate.filter(ImageFilter.GaussianBlur(blur))


def load_page(path: Path, width: int) -> Image.Image:
    page = Image.open(path).convert("RGB")
    height = round(page.height * (width / page.width))
    return page.resize((width, height), Image.Resampling.LANCZOS)


def paste_page(canvas: Image.Image, page: Image.Image, origin, shadow=True):
    x, y = origin
    if shadow:
        blur = 36
        plate = shadow_for(page.size, blur=blur, alpha=70)
        canvas.alpha_composite(plate, (x - blur * 2 + 8, y - blur * 2 + 22))
    frame = Image.new("RGBA", page.size, (255, 255, 255, 255))
    frame.paste(page, (0, 0))
    # A hairline keeps the white sheet from melting into the paper ground.
    ImageDraw.Draw(frame).rectangle(
        (0, 0, page.width - 1, page.height - 1),
        outline=(201, 194, 184, 255),
    )
    canvas.alpha_composite(frame, (x, y))


def build_banner() -> Image.Image:
    canvas = Image.new("RGBA", (W, H), PAPER + (255,))

    # The sheet is large enough that "INVOICE" and the table header stay
    # readable, and tall enough to run off the bottom edge.
    page_w = 1460
    letter = load_page(PREVIEW / "invoice-us-letter.png", page_w)
    a4 = load_page(PREVIEW / "invoice-a4.png", page_w - 36)

    # Right margin stays clear of the edge. The bottom of the frame lands
    # in the white gap under the first line item, not through the type.
    page_x = 1740
    page_y = 52
    paste_page(canvas, a4, (page_x - 34, page_y - 20))
    paste_page(canvas, letter, (page_x, page_y))

    draw = ImageDraw.Draw(canvas)

    display = font("PlayfairDisplay-Bold.ttf", 148)
    italic = font("PlayfairDisplay-Italic.ttf", 40)
    regular = font("Montserrat-Regular.ttf", 22)
    bold = font("Montserrat-Bold.ttf", 18)

    left = 220
    eyebrow = "EDITABLE INVOICE"
    title = "Sable"
    line_a = "For freelancers and"
    line_b = "small businesses."
    meta_items = ("US LETTER", "A4", "CANVA")

    eyebrow_h = bold.getbbox(eyebrow)[3] - bold.getbbox(eyebrow)[1]
    title_h = display.getbbox(title)[3] - display.getbbox(title)[1]
    italic_h = italic.getbbox(line_a)[3] - italic.getbbox(line_a)[1]
    meta_h = regular.getbbox(meta_items[0])[3] - regular.getbbox(meta_items[0])[1]

    gap_after_eyebrow = 22
    gap_after_title = 28
    gap_between_italic = 6
    gap_before_rule = 32
    rule_h = 2
    gap_after_rule = 26

    block_h = (
        eyebrow_h
        + gap_after_eyebrow
        + title_h
        + gap_after_title
        + italic_h
        + gap_between_italic
        + italic_h
        + gap_before_rule
        + rule_h
        + gap_after_rule
        + meta_h
    )
    y = (H - block_h) // 2

    text_top(draw, (left, y), eyebrow, bold, BRASS, tracking=6)
    y += eyebrow_h + gap_after_eyebrow

    text_top(draw, (left, y), title, display, FOREST)
    y += title_h + gap_after_title

    text_top(draw, (left, y), line_a, italic, INK)
    y += italic_h + gap_between_italic
    text_top(draw, (left, y), line_b, italic, INK)
    y += italic_h + gap_before_rule

    rule_w = 168
    draw.rectangle((left, y, left + rule_w, y + rule_h), fill=FOREST)
    y += rule_h + gap_after_rule

    meta_w = draw_meta(draw, left, y, meta_items, regular, BRASS)

    # Keep the wordmark clear of the page. Fail the build if a future
    # copy change runs the italic into the sheet.
    widest = max(
        ink_width(bold, eyebrow, tracking=6),
        ink_width(display, title),
        ink_width(italic, line_a),
        ink_width(italic, line_b),
        meta_w,
    )
    limit = page_x - 34 - 48
    if left + widest > limit:
        raise SystemExit(f"Banner type is {left + widest}px wide; page starts at {limit}px")

    return canvas.convert("RGB")


def build_icon() -> Image.Image:
    size = 500
    canvas = Image.new("RGBA", (size, size), PAPER + (255,))
    draw = ImageDraw.Draw(canvas)
    cx = cy = size // 2
    outer = 188
    draw.ellipse((cx - outer, cy - outer, cx + outer, cy + outer), fill=FOREST)
    ring = 174
    draw.ellipse(
        (cx - ring, cy - ring, cx + ring, cy + ring),
        outline=BRASS + (255,),
        width=6,
    )
    face = font("PlayfairDisplay-Bold.ttf", 168)
    bbox = face.getbbox("S")
    draw.text(
        (cx - (bbox[2] + bbox[0]) / 2, cy - (bbox[3] + bbox[1]) / 2 - 6),
        "S",
        font=face,
        fill=CREAM,
    )
    return canvas.convert("RGB")


def main():
    banner = build_banner()
    banner.save(OUT / "shop-banner.png", "PNG", optimize=True)
    small = banner.resize((1600, 400), Image.Resampling.LANCZOS)
    small.save(OUT / "shop-banner-1600.png", "PNG", optimize=True)
    icon = build_icon()
    icon.save(OUT / "shop-icon.png", "PNG", optimize=True)
    print(f"shop-banner.png {banner.size}")
    print(f"shop-banner-1600.png {small.size}")
    print(f"shop-icon.png {icon.size}")


if __name__ == "__main__":
    main()
