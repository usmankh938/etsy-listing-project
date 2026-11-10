# Sable invoice template

Original small-business invoice for an Etsy listing. Buyers edit it in Canva. You keep the master files in this folder.

The sample studio (Willow Studio), client, and prices are placeholders. They are what show in the listing photos. Buyers replace them with their own.

## Files

| File | Use |
| --- | --- |
| `source/invoice-us-letter.pptx` | Canva source, US Letter (8.5 × 11 in) |
| `source/invoice-a4.pptx` | Canva source, A4 (210 × 297 mm) |
| `source/invoice-us-letter.svg` | Vector version of the same layout |
| `source/invoice-a4.svg` | Vector version, A4 |
| `source/layers.txt` | Every text box name and the sample words in it |
| `preview/etsy/01-cover.png` | First Etsy photo, square thumbnail |
| `preview/etsy/02-us-letter.png` | Letter page, 4:5 |
| `preview/etsy/03-a4.png` | A4 page, 4:5 |
| `preview/etsy/04-items.png` | Close-up of the header and line items |
| `preview/etsy/05-totals.png` | Close-up of the notes and total |
| `preview/etsy/06-how-to-edit.png` | How a buyer edits it in Canva |
| `preview/invoice-us-letter.png` | Flat Letter proof |
| `preview/invoice-a4.png` | Flat A4 proof |
| `preview/invoice-us-letter.pdf` | Print proof, US Letter |
| `preview/invoice-a4.pdf` | Print proof, A4 |
| `etsy/listing-copy.md` | Title, tags, and description to paste into Etsy |
| `etsy/buyer-instructions.md` | Words for the PDF buyers download |

Import the PowerPoint into Canva. The text comes in as editable text. The SVG is the vector backup. The PNG and PDF are the listing photos and the print proof.

## Turn it into a Canva template

1. Sign in at [canva.com](https://www.canva.com).
2. Drag `source/invoice-us-letter.pptx` onto the Canva home page. Canva turns each slide into an editable page.
3. Click a few text boxes. The fonts should be **Playfair Display** and **Montserrat**. Both are free in Canva. If a box shows a missing font, select the words and choose that font.
4. Check the page size in the resize control. It should be **8.5 × 11 in**.
5. If Canva opened it as a presentation (1920 × 1080), make a new design at **Custom size → 8.5 × 11 in**. On the imported page press Ctrl+A or Cmd+A, copy, and paste onto the Letter page. Drag the pasted group so it fills the page.
6. Repeat with `source/invoice-a4.pptx` and a new design at **A4**.
7. Share → **Template link** (Canva Pro). Each buyer who opens that link gets their own copy. Put the link in the buyer PDF. Keep these PPTX files as your masters. Do not upload the PowerPoint, the SVG, or the font files as the Etsy download.

A free Canva account can edit the file. The link that gives every buyer a separate copy is a Canva Pro feature, which is what this kind of listing uses.

## What a buyer can change

Every word is its own text box: business name, monogram letter, client, dates, line items, notes, and totals. The forest shapes (monogram, table header, total bar, rules) are separate objects, so the color can be changed without retyping. Two line-item rows are blank on purpose. Duplicate a row in Canva to add another.

`source/layers.txt` lists each box.

## Colors

| | Hex | Where |
| --- | --- | --- |
| Forest | `#1E3A32` | Monogram, rules, table header, total |
| Cream | `#F7F4EE` | Type on the forest bars |
| Brass | `#7A6238` | Section labels and the DUE pill |
| Ink | `#1C1915` | Body text |
| Muted | `#5E5953` | Dates, footer, secondary labels |
| Rule | `#C9C2B8` | Lines between line items |
| Zebra | `#F6F3EE` | Alternating item rows |

## Rebuild

The previews are already built. To regenerate them after a layout change:

```bash
python3 build.py
python3 etsy_photos.py
```

That needs Python 3 with `python-pptx`, `pillow`, and `pypdf`, plus Chrome. Fonts used for the proof are in `fonts/` (Playfair Display and Montserrat, SIL Open Font License). Canva already has both, so buyers do not need these files.

## What the buyer receives

A short PDF that contains the Canva template link and the steps in `etsy/buyer-instructions.md`. Paste your link in before you upload that PDF to Etsy.
