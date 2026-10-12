"""Single Vercel entrypoint (Flask). Replaces the old api/menu.py + generate.py.

Routes:
  GET /api/menu      -> {"menu": {...} | null}   (from Supabase)
  PUT /api/menu      -> {"menu": {...}} saves it
  POST /api/generate -> {year, month, holidays: [...]} -> PDF download

Env vars (Vercel Project Settings): SUPABASE_URL, SUPABASE_SERVICE_KEY.
"""

import calendar
import datetime
import io
import json
import os
import urllib.request
from xml.sax.saxutils import escape

from flask import Flask, Response, jsonify, request
from PIL import Image
from pypdf import PdfReader
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as canvas_mod
from reportlab.platypus import Paragraph, Table, TableStyle

app = Flask(__name__)

# ---- layout constants (identical to desktop app) ----
PAGE_W, PAGE_H = 841.92, 595.32
COL_WIDTHS = [134.7, 134.6, 134.7, 134.6, 134.6, 134.8]
TABLE_X, TABLE_TOP, TABLE_BOTTOM = 17.0, 125.2, 514.1
ROW_HEIGHTS = [27.1, 32.0, 18.4, 17.5, 17.6, 17.6, 17.5, 18.3, 29.0, 18.4,
               17.6, 17.5, 17.5, 17.6, 17.6, 17.5, 17.5, 17.5, 17.5, 17.7]
C_TITLE, C_DAYS = HexColor("#7DB85B"), HexColor("#8FBF6B")
C_SECTION = HexColor("#FFE7A8")
C_WHITE, C_ALT = HexColor("#FFFFFF"), HexColor("#EAF4E2")
C_GRID, C_BLACK = HexColor("#C4D8B4"), HexColor("#000000")
FONT = "Helvetica-Bold"
TITLE_SIZE, DAY_SIZE, SECTION_SIZE, DATA_SIZE = 17.5, 12.25, 12.0, 10.5
BG_X0, BG_TOP, BG_W, BG_H = 0.0, 60.84, 841.92, 473.64
PX_PER_PT_X, PX_PER_PT_Y = 1024 / 841.92, 577 / 473.64
PILL_PX = (522, 20, 763, 59)
PILL_FILL_PX = (528, 25, 758, 54)
PILL_FILL_COLOR = (162, 200, 145)
PILL_TEXT_COLOR = HexColor("#8E1F28")
PILL_FONT_SIZE, PILL_MAX_TEXT_W = 19.5, 188.0
DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

_bg_cache = {}


def _sb_env():
    return (os.environ["SUPABASE_URL"].rstrip("/"),
            os.environ["SUPABASE_SERVICE_KEY"])


def _sb_get_menu():
    base, key = _sb_env()
    req = urllib.request.Request(
        base + "/rest/v1/app_data?key=eq.menu&select=value",
        headers={"apikey": key, "Authorization": "Bearer " + key})
    with urllib.request.urlopen(req, timeout=20) as resp:
        rows = json.loads(resp.read().decode())
    if not rows:
        raise ValueError("Cloud menu is empty - import menu.xlsx in the Edit tab first.")
    return rows[0]["value"]


def _sb_save_menu(menu):
    base, key = _sb_env()
    req = urllib.request.Request(
        base + "/rest/v1/app_data",
        data=json.dumps({"key": "menu", "value": menu}).encode(),
        method="POST",
        headers={"apikey": key, "Authorization": "Bearer " + key,
                 "Content-Type": "application/json",
                 "Prefer": "resolution=merge-duplicates"})
    with urllib.request.urlopen(req, timeout=20):
        pass


def _get_background():
    if "bg" in _bg_cache:
        return _bg_cache["bg"]
    base, _ = _sb_env()
    with urllib.request.urlopen(
            base + "/storage/v1/object/public/assets/format.pdf",
            timeout=30) as resp:
        pdf_bytes = resp.read()
    img = list(PdfReader(io.BytesIO(pdf_bytes)).pages[0].images)[0]
    bg = Image.open(io.BytesIO(img.data)).convert("RGB")
    x0, y0, x1, y1 = PILL_FILL_PX
    bg.paste(Image.new("RGB", (x1 - x0, y1 - y0), PILL_FILL_COLOR), (x0, y0))
    _bg_cache["bg"] = bg
    return bg


def monday_weeks(year, month):
    first = datetime.date(year, month, 1)
    last = datetime.date(year, month, calendar.monthrange(year, month)[1])
    monday = first - datetime.timedelta(days=first.weekday())
    if monday + datetime.timedelta(days=5) < first:
        monday += datetime.timedelta(days=7)
    weeks = []
    while monday <= last:
        weeks.append([monday + datetime.timedelta(days=i) for i in range(6)])
        monday += datetime.timedelta(days=7)
    return weeks


def sheet_for_week(i):
    return "week%d" % ((i % 5) + 1)


def week_title(year, month, week):
    start, end = week[0], week[-1]
    mon = calendar.month_name[month].upper()
    rng = ("%d-%d %s" % (start.day, end.day, start.strftime("%b"))
           if start.month == end.month else
           "%d %s-%d %s" % (start.day, start.strftime("%b"), end.day, end.strftime("%b")))
    return "%s MENU \u2022 %s" % (mon, rng)


def build_pdf(year, month, holiday_dates, menu):
    from reportlab.pdfbase.pdfmetrics import stringWidth
    weeks = monday_weeks(year, month)
    first_day = datetime.date(year, month, 1)
    bg_reader = ImageReader(_get_background())
    buf = io.BytesIO()
    c = canvas_mod.Canvas(buf, pagesize=(PAGE_W, PAGE_H))
    st_title = ParagraphStyle("t", fontName=FONT, fontSize=TITLE_SIZE, leading=TITLE_SIZE + 2,
                              alignment=TA_CENTER, textColor=C_BLACK)
    st_day = ParagraphStyle("d", fontName=FONT, fontSize=DAY_SIZE, leading=DAY_SIZE + 2.5,
                            alignment=TA_CENTER, textColor=C_BLACK)
    st_sec = ParagraphStyle("s", fontName=FONT, fontSize=SECTION_SIZE, leading=SECTION_SIZE + 2,
                            alignment=TA_CENTER, textColor=C_BLACK)
    st_data = ParagraphStyle("b", fontName=FONT, fontSize=DATA_SIZE, leading=DATA_SIZE + 1.5,
                             alignment=TA_CENTER, textColor=C_BLACK)
    for wi, week in enumerate(weeks):
        m = menu.get(sheet_for_week(wi))
        if m is None:
            raise ValueError("Menu has no '%s' data" % sheet_for_week(wi))
        c.drawImage(bg_reader, BG_X0, PAGE_H - BG_TOP - BG_H, width=BG_W, height=BG_H)
        pill_text = "%s MENU" % calendar.month_name[month].upper()
        size = PILL_FONT_SIZE
        if stringWidth(pill_text, FONT, size) > PILL_MAX_TEXT_W:
            size = size * PILL_MAX_TEXT_W / stringWidth(pill_text, FONT, size)
        c.setFont(FONT, size)
        c.setFillColor(PILL_TEXT_COLOR)
        pill_cx = (PILL_PX[0] + PILL_PX[2]) / 2 / PX_PER_PT_X
        pill_cy = BG_TOP + (PILL_PX[1] + PILL_PX[3]) / 2 / PX_PER_PT_Y
        c.drawCentredString(pill_cx, PAGE_H - (pill_cy + size * 0.36), pill_text)
        is_pad = [d < first_day for d in week]
        is_hol = [d in holiday_dates for d in week]

        def cell(t):
            return Paragraph(escape(t) if t else "", st_data)

        day_cells = [Paragraph("%s<br/>%d/%d/%d" % (DAY_NAMES[i], d.day, d.month, d.year), st_day)
                     for i, d in enumerate(week)]
        brow = [[cell("" if is_pad[i] else
                        (("HOLIDAY" if r == 0 else "") if is_hol[i]
                         else m["breakfast"][r][i]))
                 for i in range(6)] for r in range(4)]
        srow = [[cell("" if (is_pad[i] or is_hol[i]) else m["snack"][0][i]) for i in range(6)]]
        lrow = [[cell("" if (is_pad[i] or is_hol[i]) else m["lunch"][r][i]) for i in range(6)]
                for r in range(10)]
        data = ([[Paragraph(escape(week_title(year, month, week)), st_title)] + [""] * 5]
                + [day_cells]
                + [[Paragraph(escape(m["sections"][0]), st_sec)] + [""] * 5]
                + brow
                + [[Paragraph(escape(m["sections"][1]), st_sec)] + [""] * 5]
                + srow
                + [[Paragraph(escape(m["sections"][2]), st_sec)] + [""] * 5]
                + lrow)
        cmds = [("SPAN", (0, 0), (-1, 0)), ("SPAN", (0, 2), (-1, 2)),
                ("SPAN", (0, 7), (-1, 7)), ("SPAN", (0, 9), (-1, 9)),
                ("BACKGROUND", (0, 0), (-1, 0), C_TITLE),
                ("BACKGROUND", (0, 1), (-1, 1), C_DAYS),
                ("BACKGROUND", (0, 2), (-1, 2), C_SECTION),
                ("BACKGROUND", (0, 7), (-1, 7), C_SECTION),
                ("BACKGROUND", (0, 9), (-1, 9), C_SECTION),
                ("GRID", (0, 0), (-1, -1), 0.75, C_GRID),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]
        for r in (3, 4, 5, 6, 8) + tuple(range(10, 20)):
            base = 3 if r <= 6 else (8 if r == 8 else 10)
            cmds.append(("BACKGROUND", (0, r), (-1, r),
                         C_ALT if (r - base) % 2 == 1 else C_WHITE))
        t = Table(data, colWidths=COL_WIDTHS, rowHeights=ROW_HEIGHTS)
        t.setStyle(TableStyle(cmds))
        t.wrapOn(c, 0, 0)
        t.drawOn(c, TABLE_X, PAGE_H - TABLE_BOTTOM)
        c.showPage()
    c.save()
    return buf.getvalue()


@app.route("/api/menu", methods=["GET", "PUT"])
def menu():
    try:
        if request.method == "GET":
            try:
                return jsonify({"menu": _sb_get_menu()})
            except ValueError:
                return jsonify({"menu": None})
        body = request.get_json(force=True, silent=True) or {}
        if "menu" not in body:
            return jsonify({"error": "missing 'menu'"}), 400
        _sb_save_menu(body["menu"])
        return jsonify({"ok": True})
    except Exception as ex:  # noqa: BLE001
        return jsonify({"error": str(ex)}), 500


@app.route("/api/generate", methods=["POST"])
def generate():
    try:
        body = request.get_json(force=True, silent=True) or {}
        year, month = int(body["year"]), int(body["month"])
        hols = {datetime.date.fromisoformat(d) for d in body.get("holidays", [])}
        pdf = build_pdf(year, month, hols, _sb_get_menu())
        return Response(pdf, mimetype="application/pdf",
                        headers={"Content-Disposition": 'attachment; filename="Menu.pdf"'})
    except Exception as ex:  # noqa: BLE001
        msg = ("Missing cloud menu - import menu.xlsx first. (%s)" % ex
               if "empty" in str(ex) else str(ex))
        return Response(msg, status=500, mimetype="text/plain")
