import io
import streamlit as st

import engine
import render
import excel_export
import gdrive


def render_page():
    st.title("📂 Ambil Data")
    st.caption(
        "Pilih tanggal & sesi, aplikasi akan membaca ulang file Excel yang tersimpan "
        "otomatis di Google Drive (setiap kali Anda klik Generate di halaman Input), "
        "lalu me-render ulang gambar laporannya."
    )

    if not gdrive.is_configured(st):
        st.warning(
            "Google Drive belum dikonfigurasi — fitur ini belum bisa dipakai. "
            "Butuh secret `gcp_service_account` yang sama seperti untuk Google Sheets."
        )
        return

    col1, col2 = st.columns(2)
    with col1:
        tanggal_str = st.text_input("Tanggal", value=engine.today_id_date(), key="ambil_tanggal")
    with col2:
        sesi = st.selectbox("Sesi", ["PAGI", "SORE"], key="ambil_sesi")

    if st.button("🔍 Ambil Gambar", type="primary"):
        with st.spinner("Mencari & memuat data dari Google Drive..."):
            file_info = gdrive.find_latest_file(st, tanggal_str, sesi)
            if not file_info:
                st.session_state.pop("ambil_img_bytes", None)
                st.error(f"Tidak ditemukan laporan tersimpan untuk **{tanggal_str} ({sesi})**.")
            else:
                try:
                    xlsx_bytes = gdrive.download_file(st, file_info["id"])
                    report_df = excel_export.read_report_df(xlsx_bytes)
                    img = render.render_report(report_df, tanggal_str, sesi)
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    st.session_state["ambil_img_bytes"] = buf.getvalue()
                    st.session_state["ambil_xlsx_bytes"] = xlsx_bytes
                    st.session_state["ambil_file_name"] = file_info["name"]
                    st.session_state["ambil_tanggal_res"] = tanggal_str
                    st.session_state["ambil_sesi_res"] = sesi
                except Exception as e:
                    st.session_state.pop("ambil_img_bytes", None)
                    st.error(f"Gagal memuat/merender data: {e}")

    if "ambil_img_bytes" in st.session_state:
        st.success(f"Dimuat dari file: `{st.session_state['ambil_file_name']}`")
        st.image(
            st.session_state["ambil_img_bytes"],
            caption=f"{st.session_state['ambil_tanggal_res']} ({st.session_state['ambil_sesi_res']})",
        )
        colA, colB = st.columns(2)
        with colA:
            st.download_button(
                "⬇️ Download Gambar (PNG)", data=st.session_state["ambil_img_bytes"],
                file_name=(
                    f"Stok_Body_{st.session_state['ambil_tanggal_res'].replace(' ', '_')}_"
                    f"{st.session_state['ambil_sesi_res']}.png"
                ),
                mime="image/png",
            )
        with colB:
            st.download_button(
                "⬇️ Download Excel", data=st.session_state["ambil_xlsx_bytes"],
                file_name=st.session_state["ambil_file_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
