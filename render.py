"""Render tabel rekap jadi gambar PNG mirip template STOK BODY HARIAN QC."""
import os
import glob as _glob
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import engine

# Cari font yang tersedia di sistem (beda-beda antara lokal vs Streamlit Cloud)
_REGULAR_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
    "/usr/share/fonts/TTF/DejaVuSans.ttf",
    "/Library/Fonts/Arial.ttf",
    "C:/Windows/Fonts/arial.ttf",
]
_BOLD_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
    "/usr/share/fonts/TTF/DejaVuSans-Bold.ttf",
    "/Library/Fonts/Arial Bold.ttf",
    "C:/Windows/Fonts/arialbd.ttf",
]


def _find_font(candidates, keyword):
    for p in candidates:
        if os.path.exists(p):
            return p
    # fallback: cari font apapun yang cocok di sistem
    for pattern in ("/usr/share/fonts/**/*.ttf", "/usr/local/share/fonts/**/*.ttf"):
        found = _glob.glob(pattern, recursive=True)
        if found:
            match = [f for f in found if keyword.lower() in os.path.basename(f).lower()]
            return match[0] if match else found[0]
    return None


_FONT_REGULAR = _find_font(_REGULAR_CANDIDATES, "bold" if False else "sans")
_FONT_BOLD = _find_font(_BOLD_CANDIDATES, "bold") or _FONT_REGULAR


def _font(size, bold=False):
    path = _FONT_BOLD if bold else _FONT_REGULAR
    if path:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    # fallback terakhir: font bawaan PIL (ukuran tetap, tapi app tidak crash)
    try:
        return ImageFont.load_default(size)
    except TypeError:
        return ImageFont.load_default()


DISPLAY_COLS = [
    ("OP100_PERIKSA",     "OP100", "PERIKSA",       (255, 255, 255)),
    ("OP105_PERBAIKAN",   "OP105", "PERBAIKAN",     (255, 255, 255)),
    ("OP107_TDKSET_IJP",  "OP107", "TDK SET\niJP",  (255, 255, 255)),
    ("OP107_TDKSET_CRJP", "OP107", "TDK SET\nCRJP", (255, 255, 255)),
    ("OP110_GERINDA",     "OP110", "GERINDA",       (146, 208, 80)),
    ("OP115_CEKULANG",    "OP115", "CEK\nULANG",    (189, 215, 238)),
    ("OP120_TAP",         "OP120", "TAP",           (255, 153, 204)),
    ("OP166_FITTINGST",   "OP166", "FITTING\n& ST", (255, 255, 255)),
    ("OP166_KIRIM",       "OP166", "KIRIM",         (255, 255, 255)),
]

COL_W_NO = 40
COL_W_TYPE = 190
COL_W_ISI = 55
COL_W_DATA = 78
COL_W_TOTAL = 90
ROW_H = 28
HEADER_H1 = 30
HEADER_H2 = 34
TITLE_H = 120
MARGIN = 10

BLACK = (0, 0, 0)
GREY_BANNER = (217, 217, 217)
YELLOW = (255, 255, 0)
GRID = (0, 0, 0)


def _text_center(draw, box, text, font, fill=BLACK):
    x0, y0, x1, y1 = box
    lines = text.split("\n")
    line_h = font.size + 3
    total_h = line_h * len(lines)
    y = y0 + ((y1 - y0) - total_h) / 2
    for line in lines:
        w = draw.textlength(line, font=font)
        x = x0 + ((x1 - x0) - w) / 2
        draw.text((x, y), line, font=font, fill=fill)
        y += line_h


LEGAL_BOX_W = 340
LEGAL_ROW_H = 22
LEGAL_SIGN_H = 42
COMPANY_LINE_H = 26
TANGGAL_LINE_H = 28


def render_report(report_df: pd.DataFrame, tanggal_str: str, sesi: str = "PAGI", kota: str = "Gresik") -> Image.Image:
    n_data_cols = len(DISPLAY_COLS)
    table_w = COL_W_NO + COL_W_TYPE + COL_W_ISI + n_data_cols * COL_W_DATA + COL_W_TOTAL
    img_w = table_w + 2 * MARGIN

    legal_box_h = LEGAL_ROW_H + LEGAL_ROW_H + LEGAL_SIGN_H + LEGAL_ROW_H
    title_block_h = legal_box_h
    top_h = COMPANY_LINE_H + title_block_h + TANGGAL_LINE_H

    groups = [g for g in engine.GROUP_ORDER if g in report_df["RowGroup"].unique()]
    n_rows = len(report_df)
    n_banners = len(groups)
    body_h = n_rows * ROW_H + n_banners * ROW_H
    footer_h = ROW_H * 3
    img_h = top_h + HEADER_H1 + HEADER_H2 + body_h + footer_h + 2 * MARGIN

    img = Image.new("RGB", (int(img_w), int(img_h)), "white")
    d = ImageDraw.Draw(img)

    f_title = _font(24, bold=True)
    f_sub = _font(13, bold=True)
    f_legal = _font(11, bold=True)
    f_legal_sm = _font(10)
    f_h = _font(11, bold=True)
    f_hh = _font(9, bold=True)
    f_cell = _font(12)
    f_cell_b = _font(12, bold=True)

    # ---------- Baris 1: nama perusahaan, center dari ujung kiri ke kanan ----------
    y = MARGIN
    _text_center(d, (MARGIN, y, MARGIN + table_w, y + COMPANY_LINE_H), "PT. SURYA PERTIWI NUSANTARA", f_sub)
    y += COMPANY_LINE_H

    # ---------- Kotak legalisasi (kanan) ----------
    legal_x0 = MARGIN + table_w - LEGAL_BOX_W
    legal_y0 = y
    col_w3 = LEGAL_BOX_W / 3

    d.rectangle([legal_x0, legal_y0, legal_x0 + LEGAL_BOX_W, legal_y0 + LEGAL_ROW_H], outline=GRID)
    _text_center(d, (legal_x0, legal_y0, legal_x0 + LEGAL_BOX_W, legal_y0 + LEGAL_ROW_H),
                 f"{kota}, {tanggal_str}", f_legal_sm)

    y2 = legal_y0 + LEGAL_ROW_H
    for i, label in enumerate(["DIBUAT", "DIPERIKSA", "DISETUJUI"]):
        x0 = legal_x0 + i * col_w3
        d.rectangle([x0, y2, x0 + col_w3, y2 + LEGAL_ROW_H], outline=GRID)
        _text_center(d, (x0, y2, x0 + col_w3, y2 + LEGAL_ROW_H), label, f_legal)

    y3 = y2 + LEGAL_ROW_H
    for i in range(3):
        x0 = legal_x0 + i * col_w3
        d.rectangle([x0, y3, x0 + col_w3, y3 + LEGAL_SIGN_H], outline=GRID)

    y4 = y3 + LEGAL_SIGN_H
    for i, label in enumerate(["INDIRECT", "FOREMAN", "SUPERVISOR"]):
        x0 = legal_x0 + i * col_w3
        d.rectangle([x0, y4, x0 + col_w3, y4 + LEGAL_ROW_H], outline=GRID)
        _text_center(d, (x0, y4, x0 + col_w3, y4 + LEGAL_ROW_H), label, f_legal)

    # ---------- Judul, center di area kiri (sebelah tabel legalisasi) ----------
    title_area_x1 = legal_x0
    _text_center(d, (MARGIN, legal_y0, title_area_x1, legal_y0 + legal_box_h), "STOK BODY HARIAN QC TK", f_title)

    # ---------- Baris TANGGAL ----------
    y = legal_y0 + legal_box_h
    d.text((MARGIN, y + 6), f"TANGGAL : {tanggal_str}  ( {sesi} )", font=f_sub, fill=BLACK)
    y += TANGGAL_LINE_H

    y = MARGIN + top_h
    header1_y0 = y
    header1_y1 = y + HEADER_H1
    cx = MARGIN
    d.rectangle([cx, header1_y0, cx + COL_W_NO, header1_y1 + HEADER_H2], outline=GRID)
    _text_center(d, (cx, header1_y0, cx + COL_W_NO, header1_y1 + HEADER_H2), "NO", f_h)
    cx += COL_W_NO
    d.rectangle([cx, header1_y0, cx + COL_W_TYPE, header1_y1 + HEADER_H2], outline=GRID)
    _text_center(d, (cx, header1_y0, cx + COL_W_TYPE, header1_y1 + HEADER_H2), "TYPE", f_h)
    cx += COL_W_TYPE
    d.rectangle([cx, header1_y0, cx + COL_W_ISI, header1_y1 + HEADER_H2], outline=GRID)
    _text_center(d, (cx, header1_y0, cx + COL_W_ISI, header1_y1 + HEADER_H2), "ISI\nPALET/\nKRT", f_hh)
    cx += COL_W_ISI

    i = 0
    while i < len(DISPLAY_COLS):
        key, op, label, color = DISPLAY_COLS[i]
        span = 1
        j = i + 1
        while j < len(DISPLAY_COLS) and DISPLAY_COLS[j][1] == op:
            span += 1
            j += 1
        w = COL_W_DATA * span
        d.rectangle([cx, header1_y0, cx + w, header1_y1], fill=color, outline=GRID)
        _text_center(d, (cx, header1_y0, cx + w, header1_y1), op, f_h)
        cx += w
        i = j
    d.rectangle([cx, header1_y0, cx + COL_W_TOTAL, header1_y1 + HEADER_H2], outline=GRID)
    _text_center(d, (cx, header1_y0, cx + COL_W_TOTAL, header1_y1 + HEADER_H2), "GRAND\nTOTAL", f_h)

    cx = MARGIN + COL_W_NO + COL_W_TYPE + COL_W_ISI
    y2 = header1_y1
    for key, op, label, color in DISPLAY_COLS:
        d.rectangle([cx, y2, cx + COL_W_DATA, y2 + HEADER_H2], fill=color, outline=GRID)
        _text_center(d, (cx, y2, cx + COL_W_DATA, y2 + HEADER_H2), label, f_hh)
        cx += COL_W_DATA

    y = header1_y1 + HEADER_H2

    col_totals = {k: 0 for k, *_ in DISPLAY_COLS}
    no_counter = 1

    for grp in groups:
        d.rectangle([MARGIN, y, MARGIN + table_w, y + ROW_H], fill=GREY_BANNER, outline=GRID)
        d.text((MARGIN + 5, y + 6), grp, font=f_cell_b, fill=BLACK)
        y += ROW_H

        sub = report_df[report_df["RowGroup"] == grp]
        for _, r in sub.iterrows():
            cx = MARGIN
            highlight = bool(r.get("IsAksesoris")) and False  # tidak dipakai untuk warna, biarkan putih
            row_fill = "white"

            d.rectangle([cx, y, cx + COL_W_NO, y + ROW_H], outline=GRID, fill=row_fill)
            _text_center(d, (cx, y, cx + COL_W_NO, y + ROW_H), str(no_counter), f_cell)
            cx += COL_W_NO
            d.rectangle([cx, y, cx + COL_W_TYPE, y + ROW_H], outline=GRID, fill=row_fill)
            d.text((cx + 5, y + 6), str(r["RowLabel"]), font=f_cell, fill=BLACK)
            cx += COL_W_TYPE
            d.rectangle([cx, y, cx + COL_W_ISI, y + ROW_H], outline=GRID, fill=row_fill)
            isi_val = str(r["ISI"]).strip()
            _text_center(d, (cx, y, cx + COL_W_ISI, y + ROW_H), isi_val if isi_val and isi_val != "nan" else "", f_cell)
            cx += COL_W_ISI
            for key, op, label, color in DISPLAY_COLS:
                val = int(r.get(key, 0) or 0)
                d.rectangle([cx, y, cx + COL_W_DATA, y + ROW_H], outline=GRID, fill=row_fill)
                _text_center(d, (cx, y, cx + COL_W_DATA, y + ROW_H), str(val) if val else "", f_cell)
                col_totals[key] += val
                cx += COL_W_DATA
            gt = int(r["GrandTotal"])
            d.rectangle([cx, y, cx + COL_W_TOTAL, y + ROW_H], outline=GRID, fill=row_fill)
            _text_center(d, (cx, y, cx + COL_W_TOTAL, y + ROW_H), str(gt), f_cell_b)

            no_counter += 1
            y += ROW_H

    totals = engine.compute_totals(report_df)
    footer_rows = [
        ("TOTAL PCS 1", totals["pcs1"], False),
        ("TOTAL AKSESORIS", totals["aksesoris"], False),
        ("TOTAL BODY", totals["body"], True),
    ]
    for label, value, show_cols in footer_rows:
        d.rectangle([MARGIN, y, MARGIN + COL_W_NO + COL_W_TYPE + COL_W_ISI, y + ROW_H], outline=GRID, fill=GREY_BANNER)
        d.text((MARGIN + 5, y + 6), label, font=f_cell_b, fill=BLACK)
        cx = MARGIN + COL_W_NO + COL_W_TYPE + COL_W_ISI
        if show_cols:
            for key, op, lab, color in DISPLAY_COLS:
                d.rectangle([cx, y, cx + COL_W_DATA, y + ROW_H], outline=GRID, fill=GREY_BANNER)
                _text_center(d, (cx, y, cx + COL_W_DATA, y + ROW_H), str(col_totals[key]), f_cell_b)
                cx += COL_W_DATA
        else:
            d.rectangle([cx, y, cx + n_data_cols * COL_W_DATA, y + ROW_H], outline=GRID, fill=GREY_BANNER)
            cx += n_data_cols * COL_W_DATA
        d.rectangle([cx, y, cx + COL_W_TOTAL, y + ROW_H], outline=GRID, fill=GREY_BANNER)
        _text_center(d, (cx, y, cx + COL_W_TOTAL, y + ROW_H), str(value), f_cell_b)
        y += ROW_H

    return img
