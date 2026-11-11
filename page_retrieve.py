"""Halaman Ambil Data: ambil ulang laporan dari histori -- Stok Body Harian QC atau Stok Opname."""
import streamlit as st

import common
import engine
import gsheets
import ui_calendar
from so import cached as so_cached, gs_so

JENIS_HARIAN = "Stok Body Harian QC"
JENIS_SO = "Stok Opname"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _build_result(q):
    """Bangun hasil dari permintaan (jenis, tanggal, sesi) -- dibaca ulang dari Google Sheet tiap
    render (cache 10 dtk), jadi perubahan di Sheet langsung ikut. Return (res, error)."""
    jenis, tanggal_str, sesi = q["jenis"], q["tanggal"], q["sesi"]
    if jenis == JENIS_SO:
        report_df = common.cached_so_report(tanggal_str)
        if report_df is None:
            return None, f"Tidak ditemukan Stok Opname tersimpan untuk **{tanggal_str}**."
        return {"jenis": jenis, "tanggal": tanggal_str, "sesi": "SORE", "missing": [],
                "png": so_cached.so_png(report_df, tanggal_str),
                "xlsx": so_cached.so_xlsx(report_df, tanggal_str),
                "base": f"Stok_Opname_{tanggal_str.replace(' ', '_')}_SORE"}, None
    history_map = common.cached_history_map(tanggal_str, sesi)
    if not history_map:
        return None, f"Tidak ditemukan laporan tersimpan untuk **{tanggal_str} ({sesi})**."
    master_df, _ = common.get_master_and_manual_rows()
    report_df, missing = engine.build_report_table_from_snapshot(master_df, history_map)
    return {"jenis": jenis, "tanggal": tanggal_str, "sesi": sesi, "missing": missing,
            "png": common.cached_render_png(report_df, tanggal_str, sesi),
            "xlsx": common.cached_build_excel(report_df, tanggal_str, sesi),
            "base": f"Stok_Body_{tanggal_str.replace(' ', '_')}_{sesi}"}, None


@st.fragment
def _show_result():
    q = st.session_state.get("ambil_query")
    if not q:
        return
    res, err = _build_result(q)
    if err:
        st.error(err)
        return
    if res["missing"]:
        with st.expander(f"⚠️ {len(res['missing'])} Type tidak ada di histori tersimpan (diisi 0)"):
            st.write(res["missing"])
    st.image(res["png"], caption=f"{res['jenis']} — {res['tanggal']} ({res['sesi']})")
    colA, colB = st.columns(2)
    with colA:
        st.download_button("⬇️ Download Gambar (PNG)", data=res["png"], file_name=f"{res['base']}.png",
                           mime="image/png", use_container_width=True)
    with colB:
        st.download_button("⬇️ Download Excel", data=res["xlsx"], file_name=f"{res['base']}.xlsx",
                           mime=XLSX_MIME, use_container_width=True)


def render_page():
    st.title("📂 Ambil Data")
    st.caption("Ambil ulang laporan dari histori yang tersimpan di Google Sheets, lalu unduh gambar / Excel-nya.")

    if not common.gsheets_ready():
        st.warning(
            "Google Sheets belum dikonfigurasi — fitur ini belum bisa dipakai. "
            "Butuh secret `gcp_service_account` dan `gsheet_name` (lihat README)."
        )
        return

    common.sheet_watcher()
    jenis = st.radio("Data yang diambil", [JENIS_HARIAN, JENIS_SO], horizontal=True, key="ambil_jenis")
    is_so = jenis == JENIS_SO
    snaps = [s for s in (common.cached_snap_so() if is_so else common.cached_snap_harian()) if s["date_obj"]]
    if not snaps:
        st.info(f"Belum ada histori {jenis} yang tersimpan di Google Sheets.")
        return

    dates = sorted({s["date_obj"] for s in snaps})
    picked = ui_calendar.date_picker_apply(
        f"ambil_cal_{'so' if is_so else 'harian'}", enabled_dates=dates, default=dates[-1], label="Tanggal"
    )

    if is_so:
        sesi = st.selectbox("Sesi", ["SORE"], disabled=True, key="ambil_sesi_so",
                            help="Stok Opname selalu masuk sesi SORE.")
    else:
        opts = [x for x in ["PAGI", "SORE"] if any(s["date_obj"] == picked and s["sesi"] == x for s in snaps)] or ["PAGI", "SORE"]
        sesi = st.selectbox("Sesi", opts, index=opts.index("SORE") if "SORE" in opts else 0,
                            key=f"ambil_sesi_{picked}")

    tanggal_str = ui_calendar.fmt_id(picked)
    common.watch(("so" if is_so else "harian", tanggal_str))
    if st.button("🔍 Ambil Gambar", type="primary"):
        if is_so:
            common.cached_so_report.clear()
        st.session_state["ambil_query"] = {"jenis": jenis, "tanggal": tanggal_str, "sesi": sesi}

    _show_result()
