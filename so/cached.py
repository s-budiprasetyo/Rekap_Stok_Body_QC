"""Fungsi ber-cache untuk PNG/Excel Stok Opname (dipakai halaman Stok Opname & Ambil Data)."""
import io
import streamlit as st

from so import format_so, sheet_png

SESI_SO = "SORE"


@st.cache_data(show_spinner=False)
def so_png(report_df, tanggal_str: str) -> bytes:
    """PNG = tampilan file Excel format Oktober (dirender dari template yang sama)."""
    wb = format_so.build_workbook(report_df, tanggal_str, SESI_SO)
    ws = wb.active
    img = sheet_png.render_sheet(ws, format_so.evaluate_formulas(ws), "A1:U82", wb=wb)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def so_xlsx(report_df, tanggal_str: str) -> bytes:
    return format_so.build_excel_bytes(report_df, tanggal_str, SESI_SO)
