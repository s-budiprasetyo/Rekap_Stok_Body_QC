"""Loader & cache bersama yang dipakai di kedua halaman (Input & Ambil Data)."""
import io
import streamlit as st
import engine
import gsheets
import render
import excel_export

MASTER_PATH = "master_type.csv"
ALL_OP_LOCATIONS = engine.ALL_OP_LOCATIONS


@st.cache_data
def _load_master():
    return engine.load_master(MASTER_PATH)


def get_master_and_manual_rows():
    master_df = _load_master()
    manual_rows = master_df[master_df["IsManual"] == "TRUE"]["RowLabel"].tolist()
    return master_df, manual_rows


def gsheets_ready() -> bool:
    return gsheets.is_configured(st)


@st.cache_data(ttl=60, show_spinner=False)
def cached_snapshots():
    """Daftar snapshot histori — disimpan sementara 60 detik supaya tidak nge-hit
    Google Sheets berkali-kali tiap kali ada interaksi (ketik/geser input)."""
    return gsheets.list_history_snapshots(st)


@st.cache_data(ttl=60, show_spinner=False)
def cached_history_map(tanggal: str, sesi: str):
    return gsheets.fetch_history_snapshot(st, tanggal, sesi)


@st.cache_data(show_spinner=False)
def cached_combine_dedup(files):
    """Cache hasil baca+dedup file mentah SAP berdasarkan isi filenya sendiri — supaya
    tidak di-parse ulang tiap kali ada interaksi lain di halaman (isi manual, dsb),
    yang sebelumnya menyebabkan lag/kedipan di UI."""
    return engine.combine_raw_files_dedup(files)


@st.cache_data(show_spinner=False)
def cached_single_file_parse(file):
    return engine.combine_raw_files_dedup([file])


@st.cache_data(show_spinner=False)
def cached_render_png(report_df, tanggal_str: str, sesi: str) -> bytes:
    """Cache hasil render gambar PNG berdasarkan isi report_df+tanggal+sesi -- render PIL
    cukup berat (dipanggil ulang di setiap rerun kalau tidak di-cache), ini salah satu
    sumber lag/kedipan yang masih tersisa tiap kali ada interaksi lain di halaman setelah
    Generate (ketik manual, dsb) padahal laporannya sendiri tidak berubah."""
    img = render.render_report(report_df, tanggal_str, sesi)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def cached_build_excel(report_df, tanggal_str: str, sesi: str) -> bytes:
    """Cache hasil build file Excel -- sama alasannya dgn cached_render_png di atas."""
    return excel_export.build_excel(report_df, tanggal_str, sesi)
