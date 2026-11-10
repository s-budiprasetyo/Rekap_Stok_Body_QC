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


@st.cache_data(ttl=10, show_spinner=False)
def cached_snapshots():
    """Daftar snapshot histori — disimpan sementara 10 detik supaya tidak nge-hit
    Google Sheets berkali-kali tiap kali ada interaksi. Pendek (10 dtk) supaya perubahan di Sheet cepat terbaca."""
    return gsheets.list_history_snapshots(st)


@st.cache_data(ttl=10, show_spinner=False)
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


@st.cache_data(ttl=10, show_spinner=False)
def cached_snap_harian():
    return gsheets.list_all_history_snapshots(st)


@st.cache_data(ttl=10, show_spinner=False)
def cached_snap_so():
    from so import gs_so
    return gs_so.list_so_snapshots(st)


@st.cache_data(ttl=10, show_spinner=False)
def cached_so_report(tanggal_str: str):
    from so import gs_so
    return gs_so.fetch_so_report(st, tanggal_str)


@st.cache_data(ttl=15, show_spinner=False)
def _poll_sheet():
    from so import gs_so
    _, manual_rows = get_master_and_manual_rows()
    return gs_so.sync_so_from_harian(st, manual_rows)


def clear_sheet_caches():
    for f in (cached_snapshots, cached_history_map, cached_snap_harian, cached_snap_so, cached_so_report):
        f.clear()


@st.fragment(run_every=20)
def sheet_watcher():
    """Cek Google Sheet tiap ~20 detik (tanpa tombol). (1) menerapkan cell yang Anda ubah di tab
    bulan ke tab SO; (2) kalau isi Sheet berubah, halaman dimuat ulang otomatis & hasil
    mengikuti angka terbaru."""
    if not gsheets_ready():
        return
    fp, _changed = _poll_sheet()
    if fp is None:
        return
    last = st.session_state.get("_sheet_fp")
    if last is None:
        st.session_state["_sheet_fp"] = fp
    elif fp != last and not st.session_state.get("_dlg_busy"):  # sedang buka dialog -> tunda
        st.session_state["_sheet_fp"] = fp
        clear_sheet_caches()
        st.session_state["_regen"] = True
        st.rerun()


def sync_now():
    """Sinkron perubahan tab bulan -> tab SO saat itu juga (tanpa menunggu pemeriksaan berkala)."""
    from so import gs_so
    _, manual_rows = get_master_and_manual_rows()
    try:
        gs_so.sync_so_from_harian(st, manual_rows)
    except Exception:
        pass
