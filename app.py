import io
import streamlit as st
import pandas as pd

import engine
import render
import excel_export
import gsheets

st.set_page_config(page_title="Rekap Stok Body Harian QC", layout="wide")
st.title("📋 Rekap Stok Body Harian QC")

MASTER_PATH = "master_type.csv"  # simpan 1 file bareng app.py di repo


@st.cache_data
def _load_master():
    return engine.load_master(MASTER_PATH)


master_df = _load_master()

# Baris master yang butuh input manual
MANUAL_ROWS = master_df[master_df["IsManual"] == "TRUE"]["RowLabel"].tolist()

# ------------------------------------------------------------------
# 1. Upload file mentah SAP
# ------------------------------------------------------------------
st.header("1️⃣ Upload File Mentah SAP")
st.caption("Upload file Posisi_Stock dari OP100, OP105, OP107, OP110, OP115, OP120, OP166 (boleh sebagian dulu).")
raw_files = st.file_uploader("Upload file .xlsx", type=["xlsx"], accept_multiple_files=True)

grouped = pd.DataFrame()
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
st.caption("Tidak wajib diisi — kosongkan / biarkan 0 kalau stok tidak ada hari ini.")

manual_values = {}
cols = st.columns(3)
for idx, label in enumerate(MANUAL_ROWS):
    with cols[idx % 3]:
        manual_values[label] = st.number_input(label, min_value=0, value=0, step=1, key=f"manual_{label}")

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

report_df = None
if st.button("🖼️ Generate", type="primary", disabled=grouped.empty):
    report_df, unmatched_df = engine.build_report_table(grouped, master_df, manual_values)
    st.session_state["report_df"] = report_df
    st.session_state["unmatched_df"] = unmatched_df
    st.session_state["tanggal_str"] = tanggal_str
    st.session_state["sesi"] = sesi

if grouped.empty:
    st.info("⬆️ Upload file SAP dulu di Langkah 1.")

if "report_df" in st.session_state:
    report_df = st.session_state["report_df"]
    unmatched_df = st.session_state["unmatched_df"]

    if not unmatched_df.empty:
        with st.expander(f"⚠️ {len(unmatched_df)} baris data SAP TIDAK cocok ke master (cek Type/Warna/Kode Pabrik/Forming)"):
            st.dataframe(unmatched_df[["Lokasi", "Type", "Warna", "KodeNorm", "Jenis Forming", "Qty"]], use_container_width=True)

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
        if gsheets.is_configured(st):
            if st.button("☁️ Simpan ke Histori (Google Sheets)"):
                ok, msg = gsheets.append_history(st, report_df, st.session_state["tanggal_str"], st.session_state["sesi"])
                (st.success if ok else st.error)(msg)
        else:
            st.caption("Histori Google Sheets belum dikonfigurasi (opsional, lihat README).")

    with st.expander("Lihat tabel data laporan"):
        st.dataframe(report_df, use_container_width=True)
