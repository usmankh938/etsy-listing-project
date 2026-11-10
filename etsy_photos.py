#!/usr/bin/env python3
"""Etsy listing photos for the Sable invoice.

Upload order is the filename order. The first image is the square thumbnail.
Every file is at least 2000px on the short side.
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image

from build import SRC, TMP, chrome_shot, write_preview_html

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "preview" / "etsy"
BG = (231, 225, 214)  # #E7E1D6


def sheet(svg: Path, html: Path, canvas: tuple[int, int], page: tuple[int, int]) -> None:
    from build import _font_css

    raw = svg.read_text(encoding="utf-8").split("?>", 1)[-1].strip()
    cw, ch = canvas
    pw, ph = page

    html.write_text(
        f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
{_font_css()}
html, body {{ margin: 0; width: {cw}px; height: {ch}px; overflow: hidden; background: #E7E1D6; }}
.stage {{ width: {cw}px; height: {ch}px; display: flex; align-items: center; justify-content: center; }}
.sheet {{
  width: {pw}px; height: {ph}px; background: #fff;
  box-shadow: 0 18px 40px rgba(48,36,20,.16), 0 2px 6px rgba(48,36,20,.05);
}}
.sheet svg {{ width: {pw}px; height: {ph}px; display: block; }}
</style>
</head>
<body>
<div class="stage"><div class="sheet">{raw}</div></div>
</body>
</html>
""",
        encoding="utf-8",
    )


def how_to_edit(html: Path) -> None:
    from build import _font_css

    html.write_text(
        f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
{_font_css()}
html, body {{ margin: 0; width: 1000px; height: 1000px; background: #E7E1D6; }}
.frame {{
  box-sizing: border-box;
  width: 1000px; height: 1000px;
  padding: 78px 80px 70px;
  display: flex;
  flex-direction: column;
  justify-content: center;
}}
.rule {{ height: 2px; width: 64px; background: #1E3A32; }}
.kicker {{
  margin-top: 26px;
  font-family: Montserrat, sans-serif;
  font-weight: 700;
  font-size: 13px;
  letter-spacing: 0.26em;
  color: #7A6238;
}}
h1 {{
  font-family: "Playfair Display", serif;
  font-weight: 400;
  font-size: 76px;
  line-height: 0.96;
  color: #1E3A32;
  margin: 16px 0 0;
}}
.lead {{
  font-family: "Playfair Display", serif;
  font-style: italic;
  font-size: 22px;
  color: #5E5953;
  margin-top: 18px;
}}
.steps {{ margin-top: 52px; }}
.step {{
  display: grid;
  grid-template-columns: 78px 1fr;
  column-gap: 8px;
  align-items: start;
  padding: 22px 0 20px;
  border-top: 1px solid #D5CDC2;
}}
.step:last-child {{ border-bottom: 1px solid #D5CDC2; }}
.n {{
  font-family: "Playfair Display", serif;
  font-size: 28px;
  color: #1E3A32;
  line-height: 1;
  padding-top: 2px;
}}
.t {{
  font-family: Montserrat, sans-serif;
  font-weight: 700;
  font-size: 18px;
  letter-spacing: 0.03em;
  color: #1C1915;
}}
.d {{
  font-family: Montserrat, sans-serif;
  font-size: 16px;
  line-height: 1.45;
  color: #5E5953;
  margin-top: 5px;
}}
</style>
</head>
<body>
<div class="frame">
  <div class="rule"></div>
  <div class="kicker">HOW TO EDIT</div>
  <h1>Open. Type.<br>Send.</h1>
  <div class="lead">A free Canva account is all you need.</div>
  <div class="steps">
    <div class="step">
      <div class="n">01</div>
      <div>
        <div class="t">Open your copy</div>
        <div class="d">Use the link in your download. Canva gives you your own invoice.</div>
      </div>
    </div>
    <div class="step">
      <div class="n">02</div>
      <div>
        <div class="t">Replace the sample</div>
        <div class="d">Business name, client, line items, and the total.</div>
      </div>
    </div>
    <div class="step">
      <div class="n">03</div>
      <div>
        <div class="t">Download a PDF</div>
        <div class="d">Send it to your client, or print it at home.</div>
      </div>
    </div>
  </div>
</div>
</body>
</html>
""",
        encoding="utf-8",
    )


def shoot(html: Path, dest: Path, width: int, height: int, scale: int = 2) -> None:
    chrome_shot(html, dest, width, height, scale)
    image = Image.open(dest)
    print(f"  {dest.name}: {image.size[0]} x {image.size[1]}")


def band(src: Path, dest: Path, page_w: float, y0: float, y1: float, short: int = 2200) -> None:
    image = Image.open(src).convert("RGB")
    ppi = image.width / page_w
    top = max(0, int(round(y0 * ppi)))
    bottom = min(image.height, int(round(y1 * ppi)))
    crop = image.crop((0, top, image.width, bottom))
    scale = short / min(crop.size)
    resized = crop.resize(
        (int(round(crop.width * scale)), int(round(crop.height * scale))),
        Image.Resampling.LANCZOS,
    )
    resized.save(dest, "PNG", optimize=True)
    print(f"  {dest.name}: {resized.size[0]} x {resized.size[1]}")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    TMP.mkdir(parents=True, exist_ok=True)
    letter_svg = SRC / "invoice-us-letter.svg"
    a4_svg = SRC / "invoice-a4.svg"

    # 1. Square cover. This is the search thumbnail, so the page stays centered.
    cover_html = TMP / "etsy-cover.html"
    sheet(letter_svg, cover_html, (1000, 1000), (703, 910))
    shoot(cover_html, OUT / "01-cover.png", 1000, 1000)

    # 2–3. Full pages buyers can actually read. 4:5, short side 2000.
    letter_html = TMP / "etsy-letter.html"
    sheet(letter_svg, letter_html, (1000, 1250), (866, 1121))
    shoot(letter_html, OUT / "02-us-letter.png", 1000, 1250)

    a4_html = TMP / "etsy-a4.html"
    # A4 proportion 8.2677 / 11.6929, height matched to the letter card.
    a4_h = 1121
    a4_w = int(round(a4_h * 8.2677 / 11.6929))
    sheet(a4_svg, a4_html, (1000, 1250), (a4_w, a4_h))
    shoot(a4_html, OUT / "03-a4.png", 1000, 1250)

    # High-resolution letter, used only to crop the close-ups.
    hi_html = TMP / "etsy-letter-hi.html"
    write_preview_html({"w": 8.5, "h": 11.0}, letter_svg, hi_html)
    hi_png = TMP / "letter-4x.png"
    shoot(hi_html, hi_png, 816, 1056, scale=4)

    # 4. Header through the line-item table, edge to edge.
    # Ends in the white gap under the table (table ends at 6.66in).
    band(hi_png, OUT / "04-items.png", 8.5, 0.0, 6.80, short=2400)

    # 5. Notes, total, signature, and footer, edge to edge.
    band(hi_png, OUT / "05-totals.png", 8.5, 6.50, 11.0, short=2200)

    # 6. Editing steps, same colors as the invoice.
    steps_html = TMP / "etsy-steps.html"
    how_to_edit(steps_html)
    shoot(steps_html, OUT / "06-how-to-edit.png", 1000, 1000)


if __name__ == "__main__":
    main()
