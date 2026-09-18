import io
import streamlit as st
import pandas as pd

import engine
import render
import excel_export
import gsheets

st.set_page_config(page_title="Rekap Stok Body Harian QC", layout="wide")
st.title("📋 Rekap Stok Body Harian QC")

with st.expander("🔧 Tes Koneksi Google Sheets (klik untuk cek)"):
    if st.button("Tes Koneksi Sekarang"):
        ok, msg = gsheets.test_connection(st)
        (st.success if ok else st.error)(msg)

MASTER_PATH = "master_type.csv"


@st.cache_data
def _load_master():
    return engine.load_master(MASTER_PATH)


master_df = _load_master()
MANUAL_ROWS = master_df[master_df["IsManual"] == "TRUE"]["RowLabel"].tolist()
gsheets_ready = gsheets.is_configured(st)

# ------------------------------------------------------------------
# 1. Upload file mentah SAP (+ opsi mode histori)
# ------------------------------------------------------------------
st.header("1️⃣ Upload File Mentah SAP")

use_history = False
history_snapshot = None
history_map = {}

if gsheets_ready:
    snapshots = gsheets.list_history_snapshots(st)
    if snapshots:
        pakai = st.radio(
            "Mau pakai data histori kemarin? (OP105-OP166 & manual diambil dari histori, "
            "Anda cuma perlu upload file OP100 hari ini)",
            ["Tidak, upload semua file seperti biasa", "Ya, pakai histori"],
            index=0,
        )
        use_history = pakai.startswith("Ya")
    else:
        st.caption("Belum ada histori tersimpan di Google Sheets — upload semua file seperti biasa dulu.")
else:
    st.caption("(Histori Google Sheets belum terhubung — upload semua file seperti biasa.)")

grouped = pd.DataFrame()
op100_only_file = None
unmatched_op100 = pd.DataFrame()

if use_history:
    labels = [s["label"] for s in gsheets.list_history_snapshots(st)]
    snap_list = gsheets.list_history_snapshots(st)
    chosen_label = st.selectbox("Pilih snapshot histori (untuk OP105-OP166 & manual):", labels)
    history_snapshot = next(s for s in snap_list if s["label"] == chosen_label)
    history_map = gsheets.fetch_history_snapshot(st, history_snapshot["tanggal"], history_snapshot["sesi"])
    st.caption(f"Histori dipilih: **{chosen_label}** ({len(history_map)} Type tersimpan).")

    op100_only_file = st.file_uploader("Upload file OP100 hari ini (.xlsx)", type=["xlsx"])
    if op100_only_file:
        try:
            grouped = engine.combine_raw_files([op100_only_file])
            st.success("File OP100 dibaca.")
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")
else:
    st.caption("Upload file Posisi_Stock dari OP100, OP105, OP107, OP110, OP115, OP120, OP166 (boleh sebagian dulu).")
    raw_files = st.file_uploader("Upload file .xlsx", type=["xlsx"], accept_multiple_files=True)
    if raw_files:
        try:
            grouped = engine.combine_raw_files(raw_files)
            st.success(f"{len(raw_files)} file dibaca.")
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

# ------------------------------------------------------------------
# 2. Input manual
# ------------------------------------------------------------------
st.header("2️⃣ Input Manual")
if use_history:
    st.caption("Nilai default diambil dari histori yang dipilih — edit kalau ada perubahan hari ini.")
else:
    st.caption("Tidak wajib diisi — kosongkan / biarkan 0 kalau stok tidak ada hari ini.")

manual_values = {}
cols = st.columns(3)
for idx, label in enumerate(MANUAL_ROWS):
    default_val = 0
    if use_history and label in history_map:
        target_col = master_df.loc[master_df["RowLabel"] == label, "ManualTargetColumn"].iloc[0]
        if target_col:
            default_val = int(history_map[label].get(target_col, 0) or 0)
    with cols[idx % 3]:
        manual_values[label] = st.number_input(label, min_value=0, value=default_val, step=1, key=f"manual_{label}")

# ------------------------------------------------------------------
# 3. Tanggal & Sesi
# ------------------------------------------------------------------
st.header("3️⃣ Tanggal & Sesi")
col1, col2 = st.columns(2)
with col1:
    tanggal_str = st.text_input("Tanggal (otomatis hari ini, bisa diubah kalau perlu)", value=engine.today_id_date())
with col2:
    sesi = st.selectbox("Sesi", ["PAGI", "SORE"])

# ------------------------------------------------------------------
# 4. Generate
# ------------------------------------------------------------------
st.header("4️⃣ Generate Laporan")

can_generate = (not grouped.empty) if not use_history else (op100_only_file is not None)

if st.button("🖼️ Generate", type="primary", disabled=not can_generate):
    if use_history:
        op100_values, unmatched_op100 = engine.match_op100_only(grouped, master_df)
        report_df, missing_types = engine.build_report_table_from_history(
            master_df, history_map, op100_values, manual_values
        )
        st.session_state["missing_types"] = missing_types
        st.session_state["unmatched_df"] = unmatched_op100
    else:
        report_df, unmatched_df = engine.build_report_table(grouped, master_df, manual_values)
        st.session_state["missing_types"] = []
        st.session_state["unmatched_df"] = unmatched_df

    st.session_state["report_df"] = report_df
    st.session_state["tanggal_str"] = tanggal_str
    st.session_state["sesi"] = sesi

if not can_generate:
    st.info("⬆️ Upload file yang diperlukan dulu di Langkah 1.")

if "report_df" in st.session_state:
    report_df = st.session_state["report_df"]
    unmatched_df = st.session_state.get("unmatched_df", pd.DataFrame())
    missing_types = st.session_state.get("missing_types", [])

    if not unmatched_df.empty:
        with st.expander(f"⚠️ {len(unmatched_df)} baris data SAP TIDAK cocok ke master (cek Type/Warna/Kode Pabrik/Forming)"):
            st.dataframe(unmatched_df[["Lokasi", "Type", "Warna", "KodeNorm", "Jenis Forming", "Qty"]], use_container_width=True)

    if missing_types:
        with st.expander(f"⚠️ {len(missing_types)} Type tidak ditemukan di histori yang dipilih (diisi 0 untuk kolom selain OP100)"):
            st.write(missing_types)

    img = render.render_report(report_df, st.session_state["tanggal_str"], st.session_state["sesi"])
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    st.image(buf.getvalue(), caption="Preview Laporan")

    colA, colB, colC = st.columns(3)
    with colA:
        st.download_button(
            "⬇️ Download Gambar (PNG)", data=buf.getvalue(),
            file_name=f"Stok_Body_{st.session_state['tanggal_str'].replace(' ', '_')}_{st.session_state['sesi']}.png",
            mime="image/png",
        )
    with colB:
        xlsx_bytes = excel_export.build_excel(report_df, st.session_state["tanggal_str"], st.session_state["sesi"])
        st.download_button(
            "⬇️ Jadikan Excel", data=xlsx_bytes,
            file_name=f"Stok_Body_{st.session_state['tanggal_str'].replace(' ', '_')}_{st.session_state['sesi']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with colC:
        if gsheets_ready:
            if st.button("☁️ Simpan ke Histori (Google Sheets)"):
                ok, msg = gsheets.append_history(st, report_df, st.session_state["tanggal_str"], st.session_state["sesi"])
                (st.success if ok else st.error)(msg)
        else:
            st.caption("Histori Google Sheets belum dikonfigurasi (opsional, lihat README).")

    with st.expander("Lihat tabel data laporan"):
        st.dataframe(report_df, use_container_width=True)
