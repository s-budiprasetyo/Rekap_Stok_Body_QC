"""
Laporan Stok Opname memakai FORMAT FILE OKTOBER (so/template_stok_body.xlsx).
Template dibuka apa adanya (font, warna, border, merge, lebar kolom, baris tersembunyi, rumus
total, area cetak), lalu HANYA sel angka stok + tanggal yang diisi ulang. Tidak ada tata letak
yang digambar ulang -- jadi hasilnya identik dengan file Oktober.
"""
import io
import os
import re
import datetime

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

import engine

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template_stok_body.xlsx")

FIRST_ROW, LAST_ROW = 11, 75          # baris data di template
DATA_COLS = ["G", "J", "K", "L", "N", "P", "Q", "R", "S", "T"]

# kolom report_df (aplikasi) -> kolom di template Oktober
COLMAP = {
    "OP100_PERIKSA": "G",       # OP 100 PERIKSA
    "OP105_PERBAIKAN": "J",     # OP 105 PERBAIKAN
    "OP115_CEKULANG": "K",      # OP 115 CEK ULANG
    "OP107_TDKSET_IJP": "L",    # TDK SET CR (iJP)
    "OP107_TDKSET_CRJP": "P",  # CRJP (header OP 107) -> kolom sendiri, terpisah dari iJP
    "OP110_GERINDA": "N",       # OP 110 GERINDA
    "OP120_TAP": "Q",           # OP 120 TAP
    "OP166_FITTINGST": "S",     # OP 166 FITTING
    "OP166_KIRIM": "T",         # OP 166 KIRIM
}
BULAN_UPPER = [b.upper() for b in engine.BULAN_ID]


def _row_map(ws) -> dict:
    """Label di kolom B template -> nomor baris."""
    rows = {}
    for r in range(FIRST_ROW, LAST_ROW + 1):
        v = ws.cell(r, 2).value
        if v:
            rows[str(v).strip()] = r
    return rows


def build_workbook(report_df, tanggal_str: str, sesi: str = "SORE"):
    """Return openpyxl Workbook (rumus total tetap rumus, seperti file Oktober)."""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active
    rows = _row_map(ws)

    # penanda "sel ini dipakai di template": berisi angka atau '-'  -> nol ditulis '-'
    marker = {}
    for r in range(FIRST_ROW, LAST_ROW + 1):
        for c in DATA_COLS:
            cell = ws[f"{c}{r}"]
            marker[(c, r)] = cell.value is not None
            cell.value = None

    for _, rec in report_df.iterrows():
        r = rows.get(str(rec["RowLabel"]).strip())
        if r is None:
            continue
        per_col = {}
        for key, col in COLMAP.items():
            per_col[col] = per_col.get(col, 0) + int(rec.get(key, 0) or 0)
        for col, val in per_col.items():
            if val > 0:
                ws[f"{col}{r}"].value = val
                if ws.row_dimensions[r].hidden:
                    ws.row_dimensions[r].hidden = False
            elif marker.get((col, r)):
                ws[f"{col}{r}"].value = "-"

    d = engine.parse_id_date(tanggal_str)
    if d:
        bulan = BULAN_UPPER[d.month - 1]
        ws["R2"].value = f"Gresik, {d.day} {bulan} {d.year}"
        ws["A7"].value = f"TANGGAL   :    {d.day}"
        ws["C7"].value = f"{bulan} {d.year}  (      {sesi}      )"
        ws.title = bulan
    else:
        ws["A7"].value = f"TANGGAL   :    {tanggal_str}"
        ws["C7"].value = f"(      {sesi}      )"
    return wb


def build_excel_bytes(report_df, tanggal_str: str, sesi: str = "SORE") -> bytes:
    wb = build_workbook(report_df, tanggal_str, sesi)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------
# Hitung nilai rumus (SUM saja) supaya preview PNG menampilkan angka total
# ----------------------------------------------------------------------
def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def evaluate_formulas(ws) -> dict:
    """{ 'U11': 254, ... } untuk semua sel berisi =SUM(range,range,...)."""
    out = {}

    def cell_val(coord):
        v = ws[coord].value
        if isinstance(v, str) and v.startswith("="):
            if coord not in out:
                out[coord] = eval_formula(v)
            return out[coord]
        return v

    def eval_formula(f):
        m = re.fullmatch(r"=SUM\((.*)\)", f.strip())
        if not m:
            return 0
        total = 0
        for part in m.group(1).split(","):
            c1, r1, c2, r2 = range_boundaries(part.strip())
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    total += _num(cell_val(f"{get_column_letter(c)}{r}"))
        return total

    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                if c.coordinate not in out:
                    out[c.coordinate] = eval_formula(c.value)
    return out
