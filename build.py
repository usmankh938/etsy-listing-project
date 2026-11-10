#!/usr/bin/env python3
"""Build the Sable invoice: Canva-editable PPTX, SVG, and listing previews.

The PowerPoint is the Canva source. Canva imports PPTX text as editable text.
SVG, PNG, and PDF are the print proof and the Etsy listing images.
"""

from __future__ import annotations

import html
import subprocess
import sys
import time
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "source"
PREVIEW = ROOT / "preview"
TMP = Path("/tmp/invoice-build")
CHROME = "google-chrome"

INK = "1C1915"
FOREST = "1E3A32"
CREAM = "F7F4EE"
BRASS = "7A6238"
MUTED = "5E5953"
RULE = "C9C2B8"
ZEBRA = "F6F3EE"
PILL = "F4EFE6"
WHITE = "FFFFFF"

SERIF = "Playfair Display"
SANS = "Montserrat"

ITEMS = [
    ("Brand workshop", "1", "900.00", "900.00"),
    ("Visual identity", "1", "2,200.00", "2,200.00"),
    ("Packaging design", "1", "1,150.00", "1,150.00"),
    ("Social templates", "2", "240.00", "480.00"),
    ("", "", "", ""),
    ("", "", "", ""),
]


def rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color)


def money(value: str) -> float:
    return float(value.replace(",", "").replace("$", "") or 0)


def check_totals() -> None:
    got = sum(money(row[3]) for row in ITEMS)
    if abs(got - money("4,730.00")) > 0.001:
        raise SystemExit(f"Line items sum to {got}, expected 4730")


def layout(page: dict) -> dict:
    """Return draw order for one page. All geometry is in inches."""
    width, height = page["w"], page["h"]
    left = 0.58
    right = width - 0.58
    extra = height - 11.0
    row_h = 0.42 + (extra * 0.55) / 6

    shapes: list[dict] = []
    texts: list[dict] = []

    def add_shape(kind: str, id_: str, x, y, w, h, fill=None, stroke=None, sw=1.0):
        shapes.append(
            {
                "kind": kind,
                "id": id_,
                "x": x,
                "y": y,
                "w": w,
                "h": h,
                "fill": fill,
                "stroke": stroke,
                "sw": sw,
            }
        )

    def add_text(
        id_,
        text,
        x,
        y,
        w,
        h,
        *,
        font=SANS,
        size=10.5,
        color=INK,
        align="left",
        bold=False,
        italic=False,
        tracking=0,
        valign="middle",
        line_pt=None,
        wrap=True,
        pad_l=0.02,
        pad_r=0.02,
    ):
        texts.append(
            {
                "id": id_,
                "text": text,
                "x": x,
                "y": y,
                "w": w,
                "h": h,
                "font": font,
                "size": size,
                "color": color,
                "align": align,
                "bold": bold,
                "italic": italic,
                "tracking": tracking,
                "valign": valign,
                "line_pt": line_pt if line_pt is not None else size * 1.35,
                "wrap": wrap,
                "pad_l": pad_l,
                "pad_r": pad_r,
            }
        )

    def hline(id_, x, y, w, color, pt):
        add_shape("rect", id_, x, y, w, pt / 72.0, fill=color)

    # --- Letterhead -------------------------------------------------------
    hline("color_forest_rule_top", left, 0.36, right - left, FOREST, 1.75)

    head_y, head_h = 0.50, 0.50
    mark = 0.62
    mark_y = head_y + (head_h - mark) / 2
    ring = 0.045
    add_shape(
        "ellipse",
        "color_brass_ring",
        left - ring,
        mark_y - ring,
        mark + ring * 2,
        mark + ring * 2,
        stroke=BRASS,
        sw=1.25,
    )
    add_shape("ellipse", "color_forest_monogram", left, mark_y, mark, mark, fill=FOREST)
    add_text(
        "monogram",
        "W",
        left,
        mark_y,
        mark,
        mark,
        font=SERIF,
        size=18,
        color=CREAM,
        align="center",
        pad_l=0,
        pad_r=0,
    )

    meta_w = 3.15
    meta_x = right - meta_w
    name_x = left + mark + 0.16
    name_w = meta_x - name_x - 0.18
    add_text(
        "business_name",
        "WILLOW STUDIO",
        name_x,
        head_y,
        name_w,
        head_h,
        font=SERIF,
        size=22,
        color=FOREST,
        tracking=0.7,
        wrap=False,
    )

    tag_y = head_y + head_h + 0.02
    tag_h = 0.28
    add_text(
        "tagline",
        "Studio for small brands",
        name_x,
        tag_y,
        name_w,
        tag_h,
        font=SERIF,
        size=12,
        color=MUTED,
        italic=True,
        valign="middle",
    )

    add_text(
        "doc_title",
        "INVOICE",
        meta_x,
        head_y,
        meta_w,
        head_h,
        font=SERIF,
        size=28,
        color=FOREST,
        align="right",
        tracking=1.1,
        wrap=False,
        pad_r=0,
    )

    label_w = 1.20
    value_x = meta_x + label_w + 0.08
    value_w = meta_w - label_w - 0.08
    # Tuck the status pill against the invoice number. The number is
    # right-aligned to the content edge; 0.78in covers "INV-1042" at 13pt.
    pill_w, pill_h = 0.64, 0.24
    pill_x = right - 0.78 - 0.12 - pill_w
    pill_y = tag_y + (tag_h - pill_h) / 2
    add_shape("roundrect", "color_brass_pill", pill_x, pill_y, pill_w, pill_h, fill=PILL, stroke=BRASS, sw=0.75)
    add_text(
        "status",
        "DUE",
        pill_x,
        pill_y,
        pill_w,
        pill_h,
        font=SANS,
        size=8,
        color=BRASS,
        align="center",
        bold=True,
        tracking=0.9,
        pad_l=0,
        pad_r=0,
        wrap=False,
    )
    add_text(
        "invoice_number",
        "INV-1042",
        value_x,
        tag_y,
        value_w,
        tag_h,
        font=SERIF,
        size=13,
        color=INK,
        align="right",
        pad_r=0,
        wrap=False,
    )

    issued_y = tag_y + tag_h + 0.06
    date_h = 0.24
    add_text(
        "label_issued",
        "Issued",
        meta_x,
        issued_y,
        label_w,
        date_h,
        size=10.5,
        color=MUTED,
        align="right",
        wrap=False,
    )
    add_text(
        "date_issued",
        "November 10, 2026",
        value_x,
        issued_y,
        value_w,
        date_h,
        size=10.5,
        align="right",
        pad_r=0,
        wrap=False,
    )
    due_y = issued_y + date_h + 0.02
    add_text(
        "label_due",
        "Due",
        meta_x,
        due_y,
        label_w,
        date_h,
        size=10.5,
        color=MUTED,
        align="right",
        wrap=False,
    )
    add_text(
        "date_due",
        "November 24, 2026",
        value_x,
        due_y,
        value_w,
        date_h,
        size=10.5,
        align="right",
        pad_r=0,
        wrap=False,
    )

    divider_y = due_y + date_h + 0.14
    hline("color_forest_divider", left, divider_y, right - left, FOREST, 0.9)

    # --- Parties ----------------------------------------------------------
    bill_label_y = divider_y + 0.18
    label_h = 0.22
    name_y = bill_label_y + label_h
    name_h = 0.32
    addr_y = name_y + name_h
    addr_h = 0.90
    gap = 0.32
    col_w = (right - left - gap) / 2
    from_x = left + col_w + gap

    add_text(
        "label_bill_to",
        "BILL TO",
        left,
        bill_label_y,
        col_w,
        label_h,
        size=9,
        color=BRASS,
        bold=True,
        tracking=1.5,
        valign="middle",
    )
    add_text(
        "bill_to_name",
        "Avery Lane",
        left,
        name_y,
        col_w,
        name_h,
        font=SERIF,
        size=15,
        valign="middle",
    )
    add_text(
        "bill_to_address",
        "1842 Market Street\nPortland, OR 97204\navery@client.co",
        left,
        addr_y,
        col_w,
        addr_h,
        size=10.5,
        color=INK,
        valign="top",
        line_pt=15,
    )

    add_text(
        "label_from",
        "FROM",
        from_x,
        bill_label_y,
        col_w,
        label_h,
        size=9,
        color=BRASS,
        bold=True,
        tracking=1.5,
    )
    add_text(
        "from_name",
        "Willow Studio",
        from_x,
        name_y,
        col_w,
        name_h,
        font=SERIF,
        size=15,
    )
    add_text(
        "from_address",
        "210 Oak Avenue, Suite 4\nAustin, TX 78701\nhello@willowstudio.co\n+1 (512) 555-0148",
        from_x,
        addr_y,
        col_w,
        addr_h,
        size=10.5,
        valign="top",
        line_pt=15,
    )

    # --- Line items -------------------------------------------------------
    table_y = addr_y + addr_h + 0.16
    header_h = 0.36
    add_shape("rect", "color_forest_table_header", left, table_y, right - left, header_h, fill=FOREST)

    # Every money figure shares this right edge, inset from the page edge
    # so the digits sit inside the header bar and the total bar.
    figure_right = right - 0.16
    amount_w = 1.38
    rate_w = 1.18
    qty_w = 0.58
    col_gap = 0.10
    amount_x = figure_right - amount_w
    rate_x = amount_x - col_gap - rate_w
    qty_x = rate_x - col_gap - qty_w
    desc_x = left + 0.12
    desc_w = qty_x - col_gap - desc_x

    columns = [
        ("col_description", "DESCRIPTION", desc_x, desc_w, "left"),
        ("col_qty", "QTY", qty_x, qty_w, "center"),
        ("col_rate", "RATE", rate_x, rate_w, "right"),
        ("col_amount", "AMOUNT", amount_x, amount_w, "right"),
    ]
    for id_, label, x, w, align in columns:
        add_text(
            id_,
            label,
            x,
            table_y,
            w,
            header_h,
            size=8.5,
            color=CREAM,
            align=align,
            bold=True,
            tracking=0.8,
            wrap=False,
            pad_l=0.0 if align == "right" else 0.02,
            pad_r=0.0 if align == "right" else 0.02,
        )

    row_top = table_y + header_h
    keys = ("desc", "qty", "rate", "amount")
    boxes = [
        (desc_x, desc_w, "left"),
        (qty_x, qty_w, "center"),
        (rate_x, rate_w, "right"),
        (amount_x, amount_w, "right"),
    ]
    for index, item in enumerate(ITEMS, start=1):
        y = row_top + (index - 1) * row_h
        filled = bool(item[0])
        if filled and index % 2 == 0:
            add_shape("rect", f"color_paper_row_{index}", left, y, right - left, row_h, fill=ZEBRA)
        if index == len(ITEMS):
            hline(f"rule_table_bottom", left, y + row_h - (1.15 / 72.0), right - left, FOREST, 1.15)
        else:
            hline(f"rule_row_{index}", left, y + row_h - (1.0 / 72.0), right - left, RULE, 1.0)
        for key, value, (x, w, align) in zip(keys, item, boxes):
            add_text(
                f"item_{index}_{key}",
                value,
                x,
                y,
                w,
                row_h,
                size=11,
                align=align,
                wrap=False,
                pad_l=0.0 if align != "left" else 0.02,
                pad_r=0.0 if align == "right" else 0.02,
            )

    table_end = row_top + len(ITEMS) * row_h

    # --- Notes, payment, totals ------------------------------------------
    block_y = table_end + 0.26
    totals_w = 2.95
    totals_x = right - totals_w
    notes_w = totals_x - left - 0.28

    add_text(
        "label_notes",
        "NOTES",
        left,
        block_y,
        notes_w,
        0.22,
        size=9,
        color=BRASS,
        bold=True,
        tracking=1.5,
    )
    add_text(
        "notes",
        "Payment is due within 14 days of the issue date.\n"
        "Include the invoice number with your transfer.\n"
        "Amounts in USD.",
        left,
        block_y + 0.24,
        notes_w,
        0.72,
        size=10.5,
        valign="top",
        line_pt=15,
    )
    add_text(
        "label_payment",
        "PAYMENT",
        left,
        block_y + 1.02,
        notes_w,
        0.22,
        size=9,
        color=BRASS,
        bold=True,
        tracking=1.5,
    )
    add_text(
        "payment",
        "Bank transfer   ·   PayPal   ·   Card",
        left,
        block_y + 1.26,
        notes_w,
        0.28,
        size=11,
        valign="middle",
    )

    tot_h = 0.30
    pairs = [
        ("label_subtotal", "Subtotal", "subtotal", "4,730.00", 0.00),
        ("label_discount", "Discount", "discount", "0.00", 0.32),
        ("label_tax", "Tax", "tax", "0.00", 0.64),
    ]
    for label_id, label, value_id, value, dy in pairs:
        add_text(
            label_id,
            label,
            totals_x,
            block_y + dy,
            1.25,
            tot_h,
            size=10.5,
            color=MUTED,
        )
        add_text(
            value_id,
            value,
            totals_x + 1.20,
            block_y + dy,
            figure_right - (totals_x + 1.20),
            tot_h,
            size=11,
            align="right",
            pad_r=0,
            wrap=False,
        )

    bar_y = block_y + 1.08
    bar_h = 0.52
    add_shape("rect", "color_forest_total", totals_x, bar_y, totals_w, bar_h, fill=FOREST)
    add_text(
        "label_total",
        "TOTAL",
        totals_x + 0.12,
        bar_y,
        1.15,
        bar_h,
        size=11,
        color=CREAM,
        bold=True,
        tracking=1.1,
    )
    add_text(
        "total",
        "$4,730.00",
        totals_x + 1.15,
        bar_y,
        figure_right - (totals_x + 1.15),
        bar_h,
        font=SERIF,
        size=18,
        color=CREAM,
        align="right",
        pad_r=0,
        wrap=False,
    )
    bar_end = bar_y + bar_h

    # --- Signature and footer, anchored to the bottom of the page --------
    footer_h = 0.28
    footer_text_y = height - 0.46 - footer_h
    footer_rule_y = footer_text_y - 0.14
    sig_y = footer_rule_y - 0.78
    if sig_y < bar_end + 0.36:
        raise SystemExit(
            f"{page['id']}: signature ({sig_y:.2f}) collides with total ({bar_end:.2f})"
        )

    sig_w = 2.55
    add_shape("rect", "rule_signature", left, sig_y, sig_w, 0.75 / 72.0, fill=FOREST)
    add_text(
        "label_signature",
        "Signature",
        left,
        sig_y + 0.06,
        sig_w,
        0.22,
        size=8.5,
        color=MUTED,
        tracking=0.6,
        valign="top",
    )
    date_line_x = right - sig_w
    add_shape("rect", "rule_date", date_line_x, sig_y, sig_w, 0.75 / 72.0, fill=FOREST)
    add_text(
        "label_sign_date",
        "Date",
        date_line_x,
        sig_y + 0.06,
        sig_w,
        0.22,
        size=8.5,
        color=MUTED,
        align="right",
        tracking=0.6,
        valign="top",
        pad_r=0,
    )

    hline("color_forest_footer_rule", left, footer_rule_y, right - left, FOREST, 0.9)
    add_text(
        "thanks",
        "Thank you for your business.",
        left,
        footer_text_y,
        4.15,
        footer_h,
        font=SERIF,
        size=12,
        color=FOREST,
        italic=True,
    )
    add_text(
        "footer_contact",
        "willowstudio.co    ·    hello@willowstudio.co",
        left + 3.4,
        footer_text_y,
        right - left - 3.4,
        footer_h,
        size=9,
        color=MUTED,
        align="right",
        pad_r=0,
        wrap=False,
    )

    return {
        "shapes": shapes,
        "texts": texts,
        "report": {
            "page": page["id"],
            "divider_y": round(divider_y, 3),
            "table_y": round(table_y, 3),
            "table_end": round(table_end, 3),
            "bar_end": round(bar_end, 3),
            "sig_y": round(sig_y, 3),
            "footer_text_y": round(footer_text_y, 3),
            "row_h": round(row_h, 3),
        },
    }


PAGES = [
    {"id": "us-letter", "name": "US Letter", "w": 8.5, "h": 11.0},
    {"id": "a4", "name": "A4", "w": 8.2677, "h": 11.6929},
]


def _set_run(run, spec: dict) -> None:
    font = spec["font"]
    size = spec["size"]
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bool(spec["bold"])
    run.font.italic = bool(spec["italic"])
    run.font.color.rgb = rgb(spec["color"])
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        node = r_pr.find(qn(tag))
        if node is None:
            node = etree.SubElement(r_pr, qn(tag))
        node.set("typeface", font)
    r_pr.set("spc", str(int(round(spec["tracking"] * 100))))
    r_pr.set("kern", "0")


def _end_para(paragraph, spec: dict) -> None:
    end = paragraph._p.find(qn("a:endParaRPr"))
    if end is None:
        end = etree.SubElement(paragraph._p, qn("a:endParaRPr"))
    end.set("sz", str(int(round(spec["size"] * 100))))
    end.set("spc", str(int(round(spec["tracking"] * 100))))
    end.set("kern", "0")
    if spec["bold"]:
        end.set("b", "1")
    if spec["italic"]:
        end.set("i", "1")
    for tag in ("a:latin", "a:ea", "a:cs"):
        node = end.find(qn(tag))
        if node is None:
            node = etree.SubElement(end, qn(tag))
        node.set("typeface", spec["font"])
    solid = end.find(qn("a:solidFill"))
    if solid is None:
        solid = etree.SubElement(end, qn("a:solidFill"))
    srgb = solid.find(qn("a:srgbClr"))
    if srgb is None:
        srgb = etree.SubElement(solid, qn("a:srgbClr"))
    srgb.set("val", spec["color"])


def write_pptx(page: dict, spec: dict, path: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(page["w"])
    prs.slide_height = Inches(page["h"])
    prs.core_properties.title = f"Sable Invoice — {page['name']}"
    prs.core_properties.subject = "Editable small-business invoice for Canva"
    prs.core_properties.category = "Invoice template"
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(WHITE)

    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    valign_map = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE}

    for shape in spec["shapes"]:
        kind = MSO_SHAPE.OVAL if shape["kind"] == "ellipse" else (
            MSO_SHAPE.ROUNDED_RECTANGLE if shape["kind"] == "roundrect" else MSO_SHAPE.RECTANGLE
        )
        box = slide.shapes.add_shape(
            kind, Inches(shape["x"]), Inches(shape["y"]), Inches(shape["w"]), Inches(shape["h"])
        )
        box.name = shape["id"]
        if shape["kind"] == "roundrect":
            box.adjustments[0] = 0.5
        if shape["fill"]:
            box.fill.solid()
            box.fill.fore_color.rgb = rgb(shape["fill"])
        else:
            box.fill.background()
        if shape["stroke"]:
            box.line.color.rgb = rgb(shape["stroke"])
            box.line.width = Pt(shape["sw"])
        else:
            box.line.fill.background()
        # Shapes are color blocks. Drop the default text body so Canva
        # doesn't show an empty "Click to add text" on every rectangle.
        tx_body = box._element.find(qn("p:txBody"))
        if tx_body is not None:
            box._element.remove(tx_body)

    for spec_text in spec["texts"]:
        box = slide.shapes.add_textbox(
            Inches(spec_text["x"]),
            Inches(spec_text["y"]),
            Inches(spec_text["w"]),
            Inches(spec_text["h"]),
        )
        box.name = spec_text["id"]
        frame = box.text_frame
        frame.word_wrap = bool(spec_text["wrap"])
        frame.auto_size = None
        frame.margin_left = Inches(spec_text["pad_l"])
        frame.margin_right = Inches(spec_text["pad_r"])
        frame.margin_top = Inches(0.01)
        frame.margin_bottom = Inches(0.01)
        frame.paragraphs[0].alignment = align_map[spec_text["align"]]
        body_pr = frame._txBody.find(qn("a:bodyPr"))
        body_pr.set("anchor", "ctr" if spec_text["valign"] == "middle" else "t")

        lines = spec_text["text"].split("\n") if spec_text["text"] else [""]
        for i, line in enumerate(lines):
            paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
            paragraph.alignment = align_map[spec_text["align"]]
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(0)
            paragraph.line_spacing = Pt(spec_text["line_pt"])
            run = paragraph.add_run()
            run.text = line
            _set_run(run, spec_text)
            _end_para(paragraph, spec_text)

    prs.save(path)


def _esc(text: str) -> str:
    return html.escape(text, quote=False)


def write_svg(page: dict, spec: dict, path: Path) -> None:
    width_px = page["w"] * 96
    height_px = page["h"] * 96
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        (
            f'<svg xmlns="http://www.w3.org/2000/svg" '
            f'width="{page["w"]:.4f}in" height="{page["h"]:.4f}in" '
            f'viewBox="0 0 {width_px:.2f} {height_px:.2f}">'
        ),
        f"<title>Sable Invoice — {page['name']}</title>",
        "<desc>Original editable small-business invoice. Import the matching PPTX into Canva to edit the text.</desc>",
        f'<rect id="page" x="0" y="0" width="{width_px:.2f}" height="{height_px:.2f}" fill="#FFFFFF"/>',
    ]

    def px(inches: float) -> float:
        return inches * 96

    for shape in spec["shapes"]:
        x, y, w, h = px(shape["x"]), px(shape["y"]), px(shape["w"]), px(shape["h"])
        fill = f'#{shape["fill"]}' if shape["fill"] else "none"
        stroke = f'#{shape["stroke"]}' if shape["stroke"] else "none"
        sw = shape["sw"] * 96 / 72 if shape["stroke"] else 0
        if shape["kind"] == "ellipse":
            parts.append(
                f'<ellipse id="{shape["id"]}" cx="{x + w / 2:.2f}" cy="{y + h / 2:.2f}" '
                f'rx="{w / 2:.2f}" ry="{h / 2:.2f}" fill="{fill}" stroke="{stroke}" '
                f'stroke-width="{sw:.2f}"/>'
            )
        else:
            rx = h / 2 if shape["kind"] == "roundrect" else 0
            parts.append(
                f'<rect id="{shape["id"]}" x="{x:.2f}" y="{y:.2f}" width="{w:.2f}" height="{h:.2f}" '
                f'rx="{rx:.2f}" fill="{fill}" stroke="{stroke}" stroke-width="{sw:.2f}"/>'
            )

    anchor = {"left": "start", "center": "middle", "right": "end"}
    for spec_text in spec["texts"]:
        size_px = spec_text["size"] * 96 / 72
        track_px = spec_text["tracking"] * 96 / 72
        if spec_text["align"] == "left":
            x = px(spec_text["x"] + spec_text["pad_l"])
        elif spec_text["align"] == "right":
            x = px(spec_text["x"] + spec_text["w"] - spec_text["pad_r"])
        else:
            x = px(spec_text["x"] + spec_text["w"] / 2)
        lines = spec_text["text"].split("\n") if spec_text["text"] else [""]
        common = (
            f'id="{spec_text["id"]}" font-family="{spec_text["font"]}" '
            f'font-size="{size_px:.2f}" font-weight="{"700" if spec_text["bold"] else "400"}" '
            f'font-style="{"italic" if spec_text["italic"] else "normal"}" '
            f'letter-spacing="{track_px:.2f}" fill="#{spec_text["color"]}" '
            f'text-anchor="{anchor[spec_text["align"]]}" xml:space="preserve"'
        )
        if spec_text["valign"] == "middle" and len(lines) == 1:
            y = px(spec_text["y"] + spec_text["h"] / 2)
            parts.append(
                f'<text {common} x="{x:.2f}" y="{y:.2f}" dominant-baseline="central">{_esc(lines[0])}</text>'
            )
        else:
            line_px = spec_text["line_pt"] * 96 / 72
            y0 = px(spec_text["y"]) + size_px * 0.92
            tspans = []
            for i, line in enumerate(lines):
                dy = 0 if i == 0 else line_px
                tspans.append(f'<tspan x="{x:.2f}" dy="{dy:.2f}">{_esc(line)}</tspan>')
            parts.append(f'<text {common} y="{y0:.2f}">{"".join(tspans)}</text>')

    parts.append("</svg>")
    path.write_text("\n".join(parts) + "\n", encoding="utf-8")


def _font_css() -> str:
    faces = [
        ("Playfair Display", 400, "normal", ROOT / "fonts/PlayfairDisplay-Regular.ttf"),
        ("Playfair Display", 400, "italic", ROOT / "fonts/PlayfairDisplay-Italic.ttf"),
        ("Playfair Display", 700, "normal", ROOT / "fonts/PlayfairDisplay-Bold.ttf"),
        ("Montserrat", 400, "normal", ROOT / "fonts/Montserrat-Regular.ttf"),
        ("Montserrat", 700, "normal", ROOT / "fonts/Montserrat-Bold.ttf"),
    ]
    rules = []
    for family, weight, style, file in faces:
        rules.append(
            "@font-face {"
            f'font-family:"{family}";'
            f'src:url("{file.resolve().as_uri()}") format("truetype");'
            f"font-weight:{weight};font-style:{style};"
            "}"
        )
    return "\n".join(rules)


def write_preview_html(page: dict, svg_path: Path, html_path: Path) -> None:
    svg = svg_path.read_text(encoding="utf-8")
    svg = svg.split("?>", 1)[-1].strip()
    html_path.write_text(
        f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
{_font_css()}
@page {{ size: {page["w"]:.4f}in {page["h"]:.4f}in; margin: 0; }}
* {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
html, body {{ margin: 0; padding: 0; width: {page["w"]:.4f}in; height: {page["h"]:.4f}in; overflow: hidden; background: #fff; }}
svg {{ display: block; width: {page["w"]:.4f}in; height: {page["h"]:.4f}in; }}
</style>
</head>
<body>
{svg}
</body>
</html>
""",
        encoding="utf-8",
    )


def write_hero_html(svg_path: Path, html_path: Path) -> None:
    svg = svg_path.read_text(encoding="utf-8").split("?>", 1)[-1].strip()
    html_path.write_text(
        f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
{_font_css()}
html, body {{ margin: 0; width: 1000px; height: 1000px; overflow: hidden; background: #E7E1D6; }}
.stage {{ width: 1000px; height: 1000px; display: flex; align-items: center; justify-content: center; }}
.sheet {{ width: 680px; height: 880px; background: #fff; box-shadow: 0 22px 48px rgba(48,36,20,.16), 0 2px 6px rgba(48,36,20,.05); }}
.sheet svg {{ width: 680px; height: 880px; display: block; }}
</style>
</head>
<body>
<div class="stage"><div class="sheet">{svg}</div></div>
</body>
</html>
""",
        encoding="utf-8",
    )


def _chrome_base() -> list[str]:
    profile = TMP / f"chrome-{time.time_ns()}"
    profile.mkdir(parents=True, exist_ok=True)
    return [
        CHROME,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--hide-scrollbars",
        "--no-first-run",
        f"--user-data-dir={profile}",
        "--virtual-time-budget=12000",
    ]


def _run_chrome(cmd: list[str], output: Path) -> None:
    """Chrome headless writes the file, then sometimes never exits."""
    if output.exists():
        output.unlink()
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    deadline = time.time() + 30
    stable = 0
    last_size = -1
    while time.time() < deadline:
        if output.exists():
            size = output.stat().st_size
            if size > 1000 and size == last_size:
                stable += 1
                if stable >= 2:
                    break
            else:
                stable = 0
            last_size = size
        time.sleep(0.25)
    proc.kill()
    proc.wait(timeout=5)
    if not output.exists() or output.stat().st_size < 1000:
        raise SystemExit(f"Chrome did not write {output}")


def chrome_shot(html_path: Path, png_path: Path, width: int, height: int, scale: int) -> None:
    png_path.parent.mkdir(parents=True, exist_ok=True)
    cmd = _chrome_base() + [
        "--force-device-scale-factor=" + str(scale),
        f"--window-size={width},{height}",
        f"--screenshot={png_path}",
        html_path.resolve().as_uri(),
    ]
    _run_chrome(cmd, png_path)


def chrome_pdf(html_path: Path, pdf_path: Path) -> None:
    cmd = _chrome_base() + [
        "--no-pdf-header-footer",
        f"--print-to-pdf={pdf_path}",
        html_path.resolve().as_uri(),
    ]
    _run_chrome(cmd, pdf_path)
    reader = PdfReader(str(pdf_path))
    if len(reader.pages) != 1:
        raise SystemExit(f"{pdf_path.name} is {len(reader.pages)} pages, expected 1")
    box = reader.pages[0].mediabox
    print(f"  pdf {pdf_path.name}: {float(box.width):.2f} x {float(box.height):.2f} pt")


def write_layer_map(specs: list[tuple[dict, dict]], path: Path) -> None:
    lines = [
        "Editable layers in the Canva source",
        "Select a text box by this name in PowerPoint's Selection Pane.",
        "After import, the same words are editable in Canva.",
        "",
    ]
    # Names are identical on both sizes; list them once, from the letter file.
    _, spec = specs[0]
    for text in spec["texts"]:
        sample = text["text"].replace("\n", " / ") or "(blank row — click and type)"
        lines.append(f"{text['id']}: {sample}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    check_totals()
    SRC.mkdir(parents=True, exist_ok=True)
    PREVIEW.mkdir(parents=True, exist_ok=True)
    TMP.mkdir(parents=True, exist_ok=True)

    built: list[tuple[dict, dict]] = []
    for page in PAGES:
        spec = layout(page)
        built.append((page, spec))
        print(page["id"], spec["report"])
        pptx_path = SRC / f"invoice-{page['id']}.pptx"
        svg_path = SRC / f"invoice-{page['id']}.svg"
        write_pptx(page, spec, pptx_path)
        write_svg(page, spec, svg_path)
        html_path = TMP / f"invoice-{page['id']}.html"
        write_preview_html(page, svg_path, html_path)
        png_path = PREVIEW / f"invoice-{page['id']}.png"
        # 96 px per inch, times the device scale, matches the SVG viewBox.
        chrome_shot(html_path, png_path, round(page["w"] * 96), round(page["h"] * 96), 3)
        shot = Image.open(png_path)
        print(f"  png {png_path.name}: {shot.size[0]} x {shot.size[1]}")
        chrome_pdf(html_path, PREVIEW / f"invoice-{page['id']}.pdf")

    write_layer_map(built, SRC / "layers.txt")

    letter_svg = SRC / "invoice-us-letter.svg"
    hero_html = TMP / "hero.html"
    write_hero_html(letter_svg, hero_html)
    hero_png = PREVIEW / "hero-us-letter.png"
    chrome_shot(hero_html, hero_png, 1000, 1000, 2)
    hero = Image.open(hero_png)
    print(f"  hero {hero.size[0]} x {hero.size[1]}")
    print("done")


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as exc:
        sys.stderr.write((exc.stderr or b"").decode()[:2000])
        raise
