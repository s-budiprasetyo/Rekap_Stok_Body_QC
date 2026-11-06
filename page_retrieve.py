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


@st.cache_data(ttl=60, show_spinner=False)
def _snap_harian():
    return gsheets.list_all_history_snapshots(st)


@st.cache_data(ttl=60, show_spinner=False)
def _snap_so():
    return gs_so.list_so_snapshots(st)


@st.fragment
def _show_result():
    res = st.session_state.get("ambil_result")
    if not res:
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

    jenis = st.radio("Data yang diambil", [JENIS_HARIAN, JENIS_SO], horizontal=True, key="ambil_jenis")
    is_so = jenis == JENIS_SO
    snaps = [s for s in (_snap_so() if is_so else _snap_harian()) if s["date_obj"]]
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
    if st.button("🔍 Ambil Gambar", type="primary"):
        with st.spinner("Mencari & memuat data..."):
            if is_so:
                report_df = gs_so.fetch_so_report(st, tanggal_str)
                if report_df is None:
                    st.session_state.pop("ambil_result", None)
                    st.error(f"Tidak ditemukan Stok Opname tersimpan untuk **{tanggal_str}**.")
                    return
                st.session_state["ambil_result"] = {
                    "jenis": jenis, "tanggal": tanggal_str, "sesi": "SORE", "missing": [],
                    "png": so_cached.so_png(report_df, tanggal_str),
                    "xlsx": so_cached.so_xlsx(report_df, tanggal_str),
                    "base": f"Stok_Opname_{tanggal_str.replace(' ', '_')}_SORE",
                }
            else:
                history_map = gsheets.fetch_history_snapshot(st, tanggal_str, sesi)
                if not history_map:
                    st.session_state.pop("ambil_result", None)
                    st.error(f"Tidak ditemukan laporan tersimpan untuk **{tanggal_str} ({sesi})**.")
                    return
                master_df, _ = common.get_master_and_manual_rows()
                report_df, missing = engine.build_report_table_from_snapshot(master_df, history_map)
                st.session_state["ambil_result"] = {
                    "jenis": jenis, "tanggal": tanggal_str, "sesi": sesi, "missing": missing,
                    "png": common.cached_render_png(report_df, tanggal_str, sesi),
                    "xlsx": common.cached_build_excel(report_df, tanggal_str, sesi),
                    "base": f"Stok_Body_{tanggal_str.replace(' ', '_')}_{sesi}",
                }

    _show_result()
