"""Halaman Stok Opname (st.navigation). Semua logika hitung ada di so/engine_so.py."""
import base64
import datetime
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd

import common
import engine
import ui_calendar
from so import cached, engine_so, gs_so

SESI_SO = "SORE"  # semua laporan stok opname masuk sesi sore
MANUAL_NOTE = "CRJP, U104ST, S90097, TX60V1"


@st.cache_data(show_spinner=False)
def _cached_read_so(file):
    return engine_so.read_so_file(file)


def _trigger_download(data: bytes, filename: str, mime: str):
    b64 = base64.b64encode(data).decode()
    components.html(
        f'<a id="dl" href="data:{mime};base64,{b64}" download="{filename}" style="display:none"></a>'
        f'<script>document.getElementById("dl").click();</script>', height=0, width=0)


@st.dialog("Simpan sebagai histori Stok Opname?")
def _save_dialog(data_bytes, filename, mime, report_df, tanggal_str, gs_ready):
    st.write(f"Simpan laporan Stok Opname **{tanggal_str} ({SESI_SO})** ke histori (tab terpisah dari Stok Body Harian)?")
    c1, c2 = st.columns(2)
    choice = None
    if c1.button("✅ Ya, simpan", use_container_width=True, type="primary"):
        choice = True
    if c2.button("❌ Tidak, download saja", use_container_width=True):
        choice = False
    if choice is not None:
        if choice:
            if gs_ready:
                ok, msg = gs_so.save_so(st, report_df, tanggal_str, SESI_SO)
                (st.success if ok else st.error)(msg)
            else:
                st.warning("Google Sheets belum dikonfigurasi — tidak bisa disimpan.")
        _trigger_download(data_bytes, filename, mime)
        st.caption("Unduhan dimulai otomatis. Kalau tidak jalan, klik tombol di bawah:")
        st.download_button("⬇️ Download sekarang", data=data_bytes, file_name=filename, mime=mime,
                           use_container_width=True)


def _checklist(info: dict):
    lines = []
    for op in engine_so.OPS:
        n = info["per_op"].get(op, 0)
        lines.append(f"✅ **{op}** — {n} pcs" if n else f"⬜ {op} — tidak ada data di file ini")
    st.markdown("  \n".join(lines))
    if info["unknown_lokasi"]:
        st.warning("Lokasi di luar OP yang dikenal (diabaikan): " +
                   ", ".join(f"{k} ({v} pcs)" for k, v in info["unknown_lokasi"].items()))


@st.dialog("🛠️ Perbaikan Data", width="large")
def _corrections_dialog(type_opts, warna_opts, so_date):
    st.caption(
        "Angka di tabel ini akan MENGURANGI stok hasil upload Stok Opname "
        "(mis. barang yang sudah diterima Packing). Isi manual di tabel, atau upload Excel."
    )
    st.session_state.setdefault("so_corr_ver", 0)
    st.session_state.setdefault("so_corr_table", st.session_state.get("so_corrections", engine_so.empty_corrections()))

    up = st.file_uploader("Upload Excel perbaikan (opsional)", type=["xlsx"], key="so_corr_file")
    default_op = st.selectbox(
        "Operation untuk baris dari Excel (dipakai kalau di Excel tidak ada kolom operation)",
        engine_so.OPS, index=engine_so.OPS.index("OP166"), key="so_corr_defop",
    )
    sel_date, sel_lok = None, None
    if up is not None:
        try:
            dates, loks = engine_so.corr_file_options(up)
        except Exception as e:
            dates, loks = [], []
            st.error(f"Gagal membaca Excel perbaikan: {e}")
        if dates:
            idx = dates.index(so_date) if so_date in dates else len(dates) - 1
            sel_date = st.selectbox(
                "Tanggal posting yang dipakai (hanya tanggal ini yang mengurangi stok)",
                dates, index=idx, format_func=ui_calendar.fmt_id, key=f"so_corr_date_{up.file_id if hasattr(up,'file_id') else up.name}")
        if loks:
            sel_lok = st.multiselect(
                "Lokasi awal yang dipakai (QCSK = Shuttle Kiln, biasanya tidak dipakai)",
                loks, default=["QC"] if "QC" in loks else loks, key="so_corr_lok")
        if st.button("➕ Masukkan isi Excel ke tabel", key="so_corr_import"):
            try:
                new = engine_so.read_correction_excel(up, default_op, sel_date, sel_lok)
                merged = engine_so.clean_corrections(pd.concat([st.session_state["so_corr_table"], new], ignore_index=True))
                st.session_state["so_corr_table"] = merged
                st.session_state["so_corr_ver"] += 1
                st.success(f"{len(new)} baris ({int(new['Jumlah'].sum())} pcs) dari Excel dimasukkan ke tabel.")
            except Exception as e:
                st.error(f"Gagal membaca Excel perbaikan: {e}")

    table = st.session_state["so_corr_table"]
    t_opts = sorted(set(type_opts) | set(table["Type"].dropna().astype(str)))
    w_opts = sorted(set(warna_opts) | set(table["Kode Warna"].dropna().astype(str)))
    edited = st.data_editor(
        table, key=f"so_corr_editor_{st.session_state['so_corr_ver']}", num_rows="dynamic",
        use_container_width=True, hide_index=True,
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=t_opts, required=True),
            "Kode Warna": st.column_config.SelectboxColumn("Kode Warna", options=w_opts, required=True),
            "Jenis Forming": st.column_config.SelectboxColumn("Jenis Forming", options=["BC", "HP", "MC"], required=True),
            "Kode Pabrik": st.column_config.SelectboxColumn("Kode Pabrik", options=["1", "2"], required=True),
            "Operation": st.column_config.SelectboxColumn("Operation", options=engine_so.OPS, required=True),
            "Jumlah": st.column_config.NumberColumn("Jumlah", min_value=1, step=1, format="%d", required=True),
        },
    )
    c1, c2 = st.columns(2)
    if c1.button("✅ Terapkan Perbaikan", type="primary", use_container_width=True):
        clean = engine_so.clean_corrections(edited)
        st.session_state["so_corrections"] = clean
        st.session_state["so_corr_table"] = clean
        st.session_state["so_corr_ver"] += 1
        st.rerun()
    if c2.button("🗑️ Hapus semua perbaikan", use_container_width=True):
        st.session_state["so_corrections"] = engine_so.empty_corrections()
        st.session_state["so_corr_table"] = engine_so.empty_corrections()
        st.session_state["so_corr_ver"] += 1
        st.rerun()


@st.fragment
def _result_section(master_df):
    """Preview + Perbaikan + Download. Fragment: klik di sini tidak me-reload seluruh halaman."""
    inp = st.session_state.get("so_inputs")
    if not inp:
        return
    corr = st.session_state.get("so_corrections", engine_so.empty_corrections())
    grouped_adj, log = engine_so.apply_corrections(inp["grouped"], corr)
    report_df, unmatched = engine_so.build_so_report(grouped_adj, master_df, inp["manual_values"])
    tanggal = inp["tanggal_str"]

    if not unmatched.empty:
        with st.expander(f"⚠️ {len(unmatched)} baris data SAP TIDAK cocok ke master"):
            st.dataframe(unmatched[["Lokasi", "Type", "Warna", "KodeNorm", "Jenis Forming", "Qty"]], use_container_width=True)

    png = cached.so_png(report_df, tanggal)
    xlsx = cached.so_xlsx(report_df, tanggal)
    st.image(png, caption=f"Preview — {tanggal} ({SESI_SO})")

    if st.button("🛠️ Perbaikan Data", use_container_width=True):
        type_opts = sorted(inp["grouped"]["Type"].unique())
        warna_opts = sorted(inp["grouped"]["Warna"].unique())
        _corrections_dialog(type_opts, warna_opts, engine.parse_id_date(tanggal))

    if len(log):
        total_kurang = int(log["Dikurangi"].sum())
        with st.expander(f"Perbaikan diterapkan: {len(log)} baris, total −{total_kurang} pcs"):
            st.dataframe(log, use_container_width=True, hide_index=True)

    base = f"Stok_Opname_{tanggal.replace(' ', '_')}_{SESI_SO}"
    gs_ready = common.gsheets_ready()
    cA, cB = st.columns(2)
    with cA:
        if st.button("⬇️ Download Gambar (PNG)", use_container_width=True):
            _save_dialog(png, f"{base}.png", "image/png", report_df, tanggal, gs_ready)
    with cB:
        if st.button("⬇️ Download Excel", use_container_width=True):
            _save_dialog(xlsx, f"{base}.xlsx",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         report_df, tanggal, gs_ready)


def render_page():
    st.title("📦 Stok Opname")
    st.caption("Upload 1 file 'Report Stock Opname', pilih tanggal histori, lalu Generate. Semua laporan masuk sesi SORE.")

    master_df, _ = common.get_master_and_manual_rows()
    gs_ready = common.gsheets_ready()

    # 1. Upload
    st.header("1️⃣ Upload File Stok Opname")
    so_file = st.file_uploader("Upload file Report Stock Opname (.xlsx) — cukup 1 file", type=["xlsx"], key="so_file")
    grouped, info = None, None
    if so_file is not None:
        try:
            grouped, info = _cached_read_so(so_file)
            st.success(f"File dibaca: {info['n_rows']} pcs.")
            _checklist(info)
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    # 2. Tanggal histori (khusus data CRJP/U104ST/S90097/TX60V1)
    st.header("2️⃣ Tanggal Histori")
    st.caption(f"Tanggal ini HANYA untuk mengambil data {MANUAL_NOTE} dari histori (sesi SORE). PCS 1 tidak diambil dari sini.")
    hist_date, manual_values = None, {}
    if gs_ready:
        dates = sorted({s["date_obj"] for s in common.cached_snapshots() if s["sesi"] == SESI_SO and s["date_obj"]})
        if dates:
            so_date = info["so_date"] if info else None
            default = so_date if so_date in dates else dates[-1]
            hist_date = ui_calendar.date_picker_apply("so_hist", enabled_dates=dates, default=default,
                                                       label="Ambil data dari histori tanggal:")
            hmap = common.cached_history_map(ui_calendar.fmt_id(hist_date), SESI_SO) if hist_date else {}
            manual_values = engine_so.manual_from_history(master_df, hmap)
            with st.expander(f"Nilai {MANUAL_NOTE} yang diambil"):
                st.dataframe(pd.DataFrame({"Type": list(manual_values), "Jumlah": list(manual_values.values())}),
                             hide_index=True, use_container_width=True)
        else:
            st.info(f"Belum ada histori sesi SORE di Google Sheets — nilai {MANUAL_NOTE} dianggap 0.")
    else:
        st.caption(f"(Google Sheets belum terhubung — nilai {MANUAL_NOTE} dianggap 0.)")

    # 3. Generate
    st.header("3️⃣ Generate")
    default_d = info["so_date"] if (info and info["so_date"]) else ui_calendar.today_wib()
    # key memuat default -> ganti file/tanggal = default baru benar-benar dipakai
    tgl_pick = ui_calendar.date_picker_apply(
        f"so_tgl_{default_d.isoformat()}", default=default_d, label="Tanggal laporan (otomatis dari file, bisa diubah)")
    tanggal_str = ui_calendar.fmt_id(tgl_pick)
    if st.button("🖼️ Generate", type="primary", disabled=grouped is None):
        st.session_state["so_inputs"] = {"grouped": grouped, "manual_values": manual_values, "tanggal_str": tanggal_str}
        st.session_state["so_corrections"] = engine_so.empty_corrections()
        st.session_state["so_corr_table"] = engine_so.empty_corrections()
        st.session_state["so_corr_ver"] = st.session_state.get("so_corr_ver", 0) + 1
    if grouped is None:
        st.info("⬆️ Upload file Stok Opname dulu di Langkah 1.")

    _result_section(master_df)
