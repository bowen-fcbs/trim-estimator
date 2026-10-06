#!/usr/bin/env python3
"""Build the Four Corners Building Supply Window & Door Monthly RPA Review.

Produces a one-page, fillable PDF that matches the company Monthly RPA Review
form (same layout, colors, fonts, field names) and adds an "Orders In & Quotes"
section: written business and quoted dollars for the month vs. budget.

    pip install reportlab
    python3 forms/rpa/build_rpa_window_door.py
"""
from pathlib import Path

from reportlab.lib.colors import Color
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

HERE = Path(__file__).resolve().parent
OUT = HERE / "Four_Corners_Building_Supply_RPA_Window_Door.pdf"
LOGO = HERE / "four_corners_logo.png"

PAGE_W, PAGE_H = letter
L, R = 36, 576  # left / right margins

ACCENT = Color(0.788, 0.416, 0.173)
BAND = Color(0.961, 0.941, 0.925)
RULE = Color(0.8, 0.8, 0.8)
INK = Color(0.267, 0.267, 0.267)
MUTED = Color(0.533, 0.533, 0.533)
WHITE = Color(1, 1, 1)
BLACK = Color(0, 0, 0)

c = canvas.Canvas(str(OUT), pagesize=letter)
c.setTitle("Four Corners Building Supply — Window & Door Monthly RPA Review")
c.setAuthor("Four Corners Building Supply")
c.setSubject("Window & Door Department — Monthly RPA (Results, Pipeline, Activity) Review")


def Y(top):
    """Layout below is written top-down (like reading the page); PDF is bottom-up."""
    return PAGE_H - top


def field(name, tip, x0, top, x1, bottom, size=8, multiline=False, maxlen=100):
    c.acroForm.textfield(
        name=name, tooltip=tip, x=x0, y=Y(bottom), width=x1 - x0, height=bottom - top,
        fontName="Helvetica", fontSize=size, textColor=BLACK, fillColor=WHITE,
        borderColor=RULE, borderWidth=1, forceBorder=True, maxlen=maxlen,
        fieldFlags="multiline" if multiline else "",
    )


def cell(x0, top, x1, bottom, lw=0.35):
    c.setStrokeColor(RULE)
    c.setLineWidth(lw)
    c.rect(x0, Y(bottom), x1 - x0, bottom - top, stroke=1, fill=0)


def band(top, title, note=None):
    """Section title bar: orange tab, beige band, orange underline. Returns next top."""
    c.setFillColor(ACCENT)
    c.rect(L, Y(top + 18), 4, 18, stroke=0, fill=1)
    c.setFillColor(BAND)
    c.rect(L + 4, Y(top + 18), R - L - 4, 18, stroke=0, fill=1)
    c.setStrokeColor(ACCENT)
    c.setLineWidth(0.75)
    c.line(L, Y(top + 18), R, Y(top + 18))
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(46, Y(top + 13), title)
    if note:
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 8)
        c.drawString(46 + stringWidth(title, "Helvetica-Bold", 9) + 4, Y(top + 13), "  " + note)
    return top + 21


def header_row(top, cols, labels, size=6.5, h=14):
    """Beige column-header row. Returns the top of the first data row."""
    c.setFillColor(BAND)
    c.rect(L, Y(top + h), R - L, h, stroke=0, fill=1)
    for (x0, x1), label in zip(zip(cols, cols[1:]), labels):
        cell(x0, top, x1, top + h, lw=0.4)
        if label:
            c.setFillColor(INK)
            c.setFont("Helvetica-Bold", size)
            c.drawCentredString((x0 + x1) / 2, Y(top + 10), label)
    return top + h


def input_rows(top, cols, rows, row_h, inset, names, tips, size=8, maxlens=None):
    """Grid of bordered cells with a text field in each. Returns bottom of last row."""
    for r in range(rows):
        t, b = top + r * row_h, top + (r + 1) * row_h
        for i, (x0, x1) in enumerate(zip(cols, cols[1:])):
            cell(x0, t, x1, b)
            field(names[i].format(r=r + 1), tips[i].format(r=r + 1),
                  x0 + inset, t + inset, x1 - inset, b - inset, size=size,
                  maxlen=(maxlens or [100] * len(names))[i])
    return top + rows * row_h


GAP = 8  # space between sections

# ---------- masthead ----------
# Logo keeps the exact size/position it has on the company form (whitespace cropped).
sx, sy = 146.6221008 / 615, 90.3274231 / 439
c.drawImage(str(LOGO), 34.0352173 + 32 * sx, 791.3470764 - 283 * sy, width=537 * sx, height=143 * sy)

c.setFillColor(INK)
c.setFont("Helvetica-Bold", 13)
c.drawCentredString(PAGE_W / 2, Y(34.1), "MONTHLY RPA REVIEW")
c.setFillColor(ACCENT)
c.setFont("Helvetica-Bold", 8.5)
c.drawCentredString(PAGE_W / 2, Y(47), "WINDOW & DOOR DEPARTMENT", charSpace=0.8)

top = 70
c.setStrokeColor(ACCENT)
c.setLineWidth(1.5)
c.line(L, Y(top), R, Y(top))

# ---------- salesperson / month / date ----------
ft = top + 9
c.setFillColor(MUTED)
c.setFont("Helvetica-Bold", 7)
for label, lx, x0, x1, name in [
    ("SALESPERSON", 36, 92.9, 229, "salesperson"),
    ("MONTH", 236, 265.7, 389, "month"),
    ("DATE", 396, 419.1, 549, "date"),
]:
    c.drawString(lx, Y(ft + 6), label)
    field(name, label.title(), x0, ft, x1, ft + 13)
c.setStrokeColor(RULE)
c.setLineWidth(0.5)
c.line(L, Y(ft + 19), R, Y(ft + 19))
top = ft + 19 + GAP

# ---------- results ----------
cols = [36, 219.6, 397.8, 486.9, 576]
top = band(top, "RESULTS")
top = header_row(top, cols, ["LAST MONTH SALES: GOAL", "LAST MONTH SALES: ACTUAL", "% OVER / UNDER", ""])
top = input_rows(top, cols, 1, 22, 2, ["res_1_1", "res_1_2", "res_1_3", "res_1_4"],
                 ["Last month sales: goal", "Last month sales: actual", "% over / under", "Results notes"])
top += GAP

# ---------- NEW: orders in / written business & quoted dollars vs budget ----------
cols = [36, 196, 316, 436, 506, 576]
top = band(top, "ORDERS IN & QUOTES", "(for the month vs. budget)")
top = header_row(top, cols, ["FOR THE MONTH", "BUDGET", "ACTUAL", "$ OVER / UNDER", "% OF BUDGET"])
for key, label in [("wb", "ORDERS IN / WRITTEN BUSINESS"), ("quoted", "QUOTED DOLLARS")]:
    b = top + 22
    c.setFillColor(BAND)
    c.rect(cols[0], Y(b), cols[1] - cols[0], 22, stroke=0, fill=1)
    cell(cols[0], top, cols[1], b)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(cols[0] + 8, Y(top + 13.5), label)
    pretty = label.title()
    input_rows(top, cols[1:], 1, 22, 2,
               [f"{key}_budget", f"{key}_actual", f"{key}_over_under", f"{key}_pct_of_budget"],
               [f"{pretty}: budget for the month", f"{pretty}: actual for the month",
                f"{pretty}: $ over / under budget", f"{pretty}: % of budget"])
    top = b
top += GAP

# ---------- pipeline ----------
cols = [36, 216, 396, 576]
top = band(top, "PIPELINE")
top = header_row(top, cols, ["PIPELINE VALUE", "SALES BUDGET", "SALES PROJECTION"], size=7)
top = input_rows(top, cols, 1, 28, 4, ["pipeline_1", "pipeline_2", "pipeline_3"],
                 ["Pipeline value", "Sales budget", "Sales projection"], size=10)
top += GAP

# ---------- new starts ----------
cols = [36, 225, 349.2, 462.6, 576]
top = band(top, "NEW STARTS", "(include exterior, interior, Windows / Doors)")
top = header_row(top, cols, ["CLIENT NAME", "PRODUCT / LINE", "START DATE", "REVENUE"])
top = input_rows(top, cols, 5, 20, 2, ["ns_{r}_1", "ns_{r}_2", "ns_{r}_3", "ns_{r}_4"],
                 ["New start {r}: client name", "New start {r}: product / line",
                  "New start {r}: start date", "New start {r}: revenue"])
top += GAP

# ---------- prospect updates ----------
cols = [36, 187.2, 576]
top = band(top, "PROSPECT UPDATES")
top = header_row(top, cols, ["CLIENT", "NOTES / UPDATES"])
top = input_rows(top, cols, 5, 22, 2, ["pu_{r}_1", "pu_{r}_2"],
                 ["Prospect {r}: client", "Prospect {r}: notes / updates"])
top += GAP

# ---------- misc ----------
FOOTER_RULE = 750
top = band(top, "MISC TOPICS / ITEMS TO DISCUSS")
box_top, box_bottom = top, FOOTER_RULE - 10
assert box_bottom - box_top >= 80, f"misc box too short: {box_bottom - box_top}pt"
cell(L, box_top, R, box_bottom, lw=0.5)
field("misc_notes", "Misc topics / items to discuss", L + 4, box_top + 4, R - 4, box_bottom - 4,
      size=9, multiline=True, maxlen=None)

# ---------- footer ----------
c.setStrokeColor(ACCENT)
c.setLineWidth(1)
c.line(L, Y(FOOTER_RULE), R, Y(FOOTER_RULE))
c.setFillColor(MUTED)
c.setFont("Helvetica", 6.5)
c.drawCentredString(PAGE_W / 2, Y(FOOTER_RULE + 9),
                    "Four Corners Building Supply  ·  Window & Door Monthly RPA Review  ·  Confidential")

c.showPage()
c.save()
print(f"wrote {OUT.relative_to(Path.cwd()) if OUT.is_relative_to(Path.cwd()) else OUT}")
