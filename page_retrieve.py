import io
import streamlit as st

import common
import engine
import render
import excel_export
import gsheets


def render_page():
    st.title("📂 Ambil Data")
    st.caption(
        "Pilih tanggal & sesi, aplikasi akan mengolah ulang data yang tersimpan otomatis "
        "di Google Sheets (setiap kali Generate di halaman Input Harian berhasil), lalu "
        "me-render ulang gambar laporannya."
    )

    gsheets_ready = common.gsheets_ready()
    if not gsheets_ready:
        st.warning(
            "Google Sheets belum dikonfigurasi — fitur ini belum bisa dipakai. "
            "Butuh secret `gcp_service_account` dan `gsheet_name` (lihat README)."
        )
        return

    master_df, _ = common.get_master_and_manual_rows()

    col1, col2 = st.columns(2)
    with col1:
        tanggal_str = st.text_input("Tanggal", value=engine.today_id_date(), key="ambil_tanggal")
    with col2:
        sesi = st.selectbox("Sesi", ["PAGI", "SORE"], key="ambil_sesi")

    if st.button("🔍 Ambil Gambar", type="primary"):
        with st.spinner("Mencari & memuat data..."):
            history_map = gsheets.fetch_history_snapshot(st, tanggal_str, sesi)
            if not history_map:
                st.session_state.pop("ambil_img_bytes", None)
                st.error(
                    f"Tidak ditemukan laporan tersimpan untuk **{tanggal_str} ({sesi})**. "
                    "Pastikan laporan untuk tanggal & sesi itu sudah pernah di-Generate di "
                    "halaman Input Harian (tersimpan otomatis sejak fitur ini aktif)."
                )
            else:
                report_df, missing_types = engine.build_report_table_from_snapshot(master_df, history_map)
                img = render.render_report(report_df, tanggal_str, sesi)
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                st.session_state["ambil_img_bytes"] = buf.getvalue()
                st.session_state["ambil_report_df"] = report_df
                st.session_state["ambil_missing_types"] = missing_types
                st.session_state["ambil_tanggal_res"] = tanggal_str
                st.session_state["ambil_sesi_res"] = sesi

    if "ambil_img_bytes" in st.session_state:
        missing_types = st.session_state.get("ambil_missing_types", [])
        if missing_types:
            with st.expander(f"⚠️ {len(missing_types)} Type tidak ada di histori tersimpan (diisi 0)"):
                st.write(missing_types)

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
            xlsx_bytes = excel_export.build_excel(
                st.session_state["ambil_report_df"],
                st.session_state["ambil_tanggal_res"],
                st.session_state["ambil_sesi_res"],
            )
            st.download_button(
                "⬇️ Download Excel", data=xlsx_bytes,
                file_name=(
                    f"Stok_Body_{st.session_state['ambil_tanggal_res'].replace(' ', '_')}_"
                    f"{st.session_state['ambil_sesi_res']}.xlsx"
                ),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
