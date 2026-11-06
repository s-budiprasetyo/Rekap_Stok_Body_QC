"""
Render sebuah worksheet openpyxl menjadi gambar PNG (PIL) yang meniru tampilan Excel:
lebar kolom, tinggi baris, baris/kolom tersembunyi, merge, warna isi (termasuk warna tema),
border, font (Arial -> Liberation Sans yang ikut dibundel), perataan, wrap teks.
Dipakai supaya preview/PNG Stok Opname = tampilan file Excel format Oktober, tanpa LibreOffice.
"""
import os
import re
import colorsys
import xml.etree.ElementTree as ET

from PIL import Image, ImageDraw, ImageFont
from openpyxl.utils import get_column_letter, range_boundaries

FONT_DIR = os.path.join(os.path.dirname(__file__), "fonts")
PX_PER_PT = 96 / 72
LINE_THIN = 2

_font_cache = {}


def _font(bold: bool, size_px: int):
    key = (bold, size_px)
    if key not in _font_cache:
        path = os.path.join(FONT_DIR, "LiberationSans-Bold.ttf" if bold else "LiberationSans-Regular.ttf")
        try:
            _font_cache[key] = ImageFont.truetype(path, size_px)
        except Exception:
            _font_cache[key] = ImageFont.load_default()
    return _font_cache[key]


# ---------------------------------------------------------------- warna
def _theme_colors(wb):
    """Urutan indeks tema Excel: 0=lt1 1=dk1 2=lt2 3=dk2 4..9=accent1..6."""
    default = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]
    try:
        root = ET.fromstring(wb.loaded_theme)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = root.find(".//a:clrScheme", ns)
        vals = []
        for child in scheme:
            el = child[0]
            vals.append(el.get("val") if el.tag.endswith("srgbClr") else el.get("lastClr"))
        # urutan di XML: dk1, lt1, dk2, lt2, accent1..6
        dk1, lt1, dk2, lt2 = vals[0], vals[1], vals[2], vals[3]
        return [lt1, dk1, lt2, dk2] + vals[4:10]
    except Exception:
        return default


def _apply_tint(hex_rgb: str, tint: float):
    if not tint:
        return hex_rgb
    r, g, b = (int(hex_rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def _color(c, theme, default=None):
    """openpyxl Color -> (r,g,b) atau default."""
    if c is None:
        return default
    try:
        if c.type == "rgb" and isinstance(c.rgb, str):
            return tuple(int(c.rgb[-6:][i:i + 2], 16) for i in (0, 2, 4))
        if c.type == "theme":
            hx = _apply_tint(theme[c.theme] if c.theme < len(theme) else "000000", c.tint or 0)
            return tuple(int(hx[i:i + 2], 16) for i in (0, 2, 4))
        if c.type == "indexed":
            return default
    except Exception:
        pass
    return default


# ---------------------------------------------------------------- format angka
def _fmt(v, nf: str) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        if '"-"' in nf and v == 0:
            return "-"
        if "#,##0" in nf:
            return f"{v:,.0f}"
        if isinstance(v, float):
            if v == int(v):
                return str(int(v))
            return f"{v:.1f}" if nf.startswith("0.0") else f"{v:g}"
        return str(v)
    return str(v)


def _wrap(draw, text, font, max_w):
    lines = []
    for para in text.split("\n"):
        words = para.split(" ")
        cur = ""
        for w in words:
            trial = (cur + " " + w).strip() if cur else w
            if draw.textlength(trial, font=font) <= max_w:
                cur = trial
                continue
            if cur:
                lines.append(cur)
                cur = ""
            # kata tunggal terlalu panjang -> potong per karakter (seperti Excel)
            if draw.textlength(w, font=font) > max_w:
                piece = ""
                for ch in w:
                    if draw.textlength(piece + ch, font=font) <= max_w or not piece:
                        piece += ch
                    else:
                        lines.append(piece)
                        piece = ch
                cur = piece
            else:
                cur = w
        lines.append(cur)
    return lines


# ---------------------------------------------------------------- render utama
def render_sheet(ws, values: dict, area: str, wb=None, scale: float = 1.0) -> Image.Image:
    """area mis. 'A1:U82'. values = hasil format_so.evaluate_formulas(ws)."""
    c0, r0, c1, r1 = range_boundaries(area)
    theme = _theme_colors(wb) if wb is not None else _theme_colors(ws.parent)
    default_w = ws.sheet_format.defaultColWidth or 9.14
    default_h = ws.sheet_format.defaultRowHeight or 15

    # ---- ukuran kolom/baris
    col_w = {}
    for c in range(c0, c1 + 1):
        col_w[c] = default_w
    hidden_c = set()
    for dim in ws.column_dimensions.values():
        if dim.min is None:
            continue
        for c in range(dim.min, min(dim.max, c1) + 1):
            if c0 <= c <= c1:
                if dim.width:
                    col_w[c] = dim.width
                if dim.hidden:
                    hidden_c.add(c)
    row_h = {r: (ws.row_dimensions[r].height or default_h) if r in ws.row_dimensions else default_h
             for r in range(r0, r1 + 1)}
    hidden_r = {r for r in range(r0, r1 + 1) if r in ws.row_dimensions and ws.row_dimensions[r].hidden}

    xs, x = {}, 2
    for c in range(c0, c1 + 1):
        xs[c] = x
        x += 0 if c in hidden_c else round(col_w[c] * 7 * scale)
    W = x - 2
    ys, y = {}, 2
    for r in range(r0, r1 + 1):
        ys[r] = y
        y += 0 if r in hidden_r else round(row_h[r] * PX_PER_PT * scale)
    H = y - 2

    img = Image.new("RGB", (W + 6, H + 6), (255, 255, 255))
    d = ImageDraw.Draw(img)

    # ---- merge
    merge_of = {}   # (c,r) -> (mc0,mr0,mc1,mr1)
    for m in ws.merged_cells.ranges:
        a, b, c_, d_ = m.min_col, m.min_row, m.max_col, m.max_row
        for rr in range(b, d_ + 1):
            for cc in range(a, c_ + 1):
                merge_of[(cc, rr)] = (a, b, c_, d_)

    def rect_of(cc, rr):
        a, b, c_, d_ = merge_of.get((cc, rr), (cc, rr, cc, rr))
        a, c_ = max(a, c0), min(c_, c1)
        b, d_ = max(b, r0), min(d_, r1)
        x0 = xs[a]
        x1 = xs[c_] + (0 if c_ in hidden_c else round(col_w[c_] * 7 * scale))
        y0 = ys[b]
        y1 = ys[d_] + (0 if d_ in hidden_r else round(row_h[d_] * PX_PER_PT * scale))
        return x0, y0, x1, y1

    def is_empty(cc, rr):
        if (cc, rr) in merge_of:
            a, b, _, _ = merge_of[(cc, rr)]
            return ws.cell(b, a).value in (None, "")
        return ws.cell(rr, cc).value in (None, "")

    # ---- 1) isi warna
    for rr in range(r0, r1 + 1):
        if rr in hidden_r:
            continue
        for cc in range(c0, c1 + 1):
            if cc in hidden_c:
                continue
            if (cc, rr) in merge_of and merge_of[(cc, rr)][:2] != (cc, rr):
                continue
            cell = ws.cell(rr, cc)
            if cell.fill and cell.fill.fill_type == "solid":
                col = _color(cell.fill.fgColor, theme)
                if col:
                    x0, y0, x1, y1 = rect_of(cc, rr)
                    d.rectangle([x0, y0, x1, y1], fill=col)

    # ---- 2) teks
    for rr in range(r0, r1 + 1):
        if rr in hidden_r:
            continue
        for cc in range(c0, c1 + 1):
            if cc in hidden_c:
                continue
            if (cc, rr) in merge_of and merge_of[(cc, rr)][:2] != (cc, rr):
                continue
            cell = ws.cell(rr, cc)
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                v = values.get(cell.coordinate, 0)
            if v is None or v == "":
                continue
            text = _fmt(v, cell.number_format or "General")
            x0, y0, x1, y1 = rect_of(cc, rr)
            f = cell.font
            size_px = max(6, round((f.sz or 11) * PX_PER_PT * scale))
            fnt = _font(bool(f.b), size_px)
            color = _color(f.color, theme, (0, 0, 0))
            al = cell.alignment
            h = al.horizontal or ("right" if isinstance(v, (int, float)) else "left")
            va = al.vertical or "bottom"
            pad = 4

            if text == "●":
                dia = size_px * 0.62
                cx, cy = (x0 + x1) / 2, (y0 + y1) / 2 + (size_px * 0.1 if va != "center" else 0)
                d.ellipse([cx - dia / 2, cy - dia / 2, cx + dia / 2, cy + dia / 2], fill=color)
                continue

            if al.wrap_text:
                lines = _wrap(d, text, fnt, max(10, x1 - x0 - 4))
            else:
                lines = text.split("\n")
            lh = round(size_px * 1.17)
            block_h = lh * len(lines)
            if va == "center":
                ty = y0 + (y1 - y0 - block_h) / 2
            elif va == "top":
                ty = y0 + 2
            else:
                ty = y1 - block_h - 2

            # area teks boleh meluber ke sel kosong di sebelahnya (seperti Excel), kecuali wrap
            avail_x0, avail_x1 = x0, x1
            if not al.wrap_text:
                last_c = merge_of.get((cc, rr), (cc, rr, cc, rr))[2]
                first_c = merge_of.get((cc, rr), (cc, rr, cc, rr))[0]
                if h in ("left", "center"):
                    k = last_c + 1
                    while k <= c1 and is_empty(k, rr):
                        avail_x1 = rect_of(k, rr)[2] if k not in hidden_c else avail_x1
                        k += 1
                if h in ("right", "center"):
                    k = first_c - 1
                    while k >= c0 and is_empty(k, rr):
                        avail_x0 = rect_of(k, rr)[0] if k not in hidden_c else avail_x0
                        k -= 1

            for i, line in enumerate(lines):
                tw = d.textlength(line, font=fnt)
                if h == "center":
                    tx = (x0 + x1) / 2 - tw / 2
                elif h == "right":
                    tx = x1 - tw - pad
                else:
                    tx = x0 + pad
                ly = ty + i * lh
                if tx >= avail_x0 and tx + tw <= avail_x1:
                    d.text((tx, ly), line, font=fnt, fill=color)
                else:
                    tmp = Image.new("RGBA", (int(tw) + 8, lh + 8), (0, 0, 0, 0))
                    ImageDraw.Draw(tmp).text((2, 2), line, font=fnt, fill=color + (255,))
                    cx0 = max(0, int(avail_x0 - tx + 2))
                    cx1 = min(tmp.width, int(avail_x1 - tx + 2))
                    if cx1 > cx0:
                        img.paste(tmp.crop((cx0, 0, cx1, tmp.height)), (int(tx - 2 + cx0), int(ly - 2)),
                                  tmp.crop((cx0, 0, cx1, tmp.height)))

    # ---- 3) border
    def has(side):
        return side is not None and side.style not in (None, "none")

    for rr in range(r0, r1 + 1):
        if rr in hidden_r:
            continue
        for cc in range(c0, c1 + 1):
            if cc in hidden_c:
                continue
            cell = ws.cell(rr, cc)
            b = cell.border
            if not b:
                continue
            mrect = merge_of.get((cc, rr))
            x0, y0, x1, y1 = (xs[cc], ys[rr],
                              xs[cc] + round(col_w[cc] * 7 * scale),
                              ys[rr] + round(row_h[rr] * PX_PER_PT * scale))
            edges = []
            if has(b.left) and (not mrect or cc == mrect[0]):
                edges.append(((x0, y0), (x0, y1)))
            if has(b.right) and (not mrect or cc == mrect[2]):
                edges.append(((x1, y0), (x1, y1)))
            if has(b.top) and (not mrect or rr == mrect[1]):
                edges.append(((x0, y0), (x1, y0)))
            if has(b.bottom) and (not mrect or rr == mrect[3]):
                edges.append(((x0, y1), (x1, y1)))
            for p, q in edges:
                d.line([p, q], fill=(0, 0, 0), width=LINE_THIN)
    return img
