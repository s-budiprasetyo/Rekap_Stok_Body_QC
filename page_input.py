import io
import streamlit as st
import pandas as pd

import common
import engine
import render
import excel_export
import gsheets

# Simpan otomatis ke Histori Google Sheets SEMENTARA DIMATIKAN (atas permintaan Anda,
# 28 Okt 2026) selama masa testing, karena ada data histori yang ternyata tidak sesuai
# harapan. Ubah jadi True kapan pun sudah yakin aman untuk dinyalakan lagi.
AUTO_SAVE_HISTORI = False


def _render_checklist(report: dict):
    """Tampilkan centang per lokasi OP yang terdeteksi dari file yang diupload,
    plus keterangan duplikat (kalau ada, otomatis diresolve: pakai file pertama)."""
    used = report.get("used", {})
    duplicates = report.get("duplicates", {})

    lines = []
    for op in common.ALL_OP_LOCATIONS:
        if op in used:
            lines.append(f"✅ **{op}** — `{used[op]}`")
        else:
            lines.append(f"⬜ {op} — belum diupload")
    st.markdown("  \n".join(lines))

    if duplicates:
        dup_msgs = []
        for op, fnames in duplicates.items():
            ignored = ", ".join(f"`{n}`" for n in fnames)
            dup_msgs.append(f"**{op}**: dipakai file pertama (`{used.get(op)}`), diabaikan: {ignored}")
        st.warning(
            "⚠️ Ada file **double** untuk lokasi OP yang sama. Aplikasi otomatis memakai "
            "file yang **pertama** diupload dan mengabaikan sisanya (tidak dijumlahkan):\n\n"
            + "\n\n".join(dup_msgs)
        )
    else:
        st.success("Tidak ada file operation yang double.")


def render_page():
    st.title("📋 Rekap Stok Body Harian QC")

    with st.expander("🔧 Tes Koneksi Google Sheets (klik untuk cek)"):
        colT1, colT2 = st.columns(2)
        with colT1:
            if st.button("Tes Koneksi Sekarang"):
                ok, msg = gsheets.test_connection(st)
                (st.success if ok else st.error)(msg)
        with colT2:
            if st.button("🩹 Perbaiki Header Histori"):
                ok, msg = gsheets.repair_header(st)
                (st.success if ok else st.error)(msg)

        st.caption(
            "Histori otomatis tersimpan per-tab per-bulan (mis. 'Oktober 2026') dalam 1 file Google Sheet "
            "yang sama — tidak perlu bikin file baru tiap bulan. Dipakai sekali saja kalau ada tab lama "
            "yang perlu diganti nama (mis. tab lama 'Histori' -> 'September 2026'):"
        )
        colR1, colR2, colR3 = st.columns([2, 2, 1])
        with colR1:
            old_tab = st.text_input("Nama tab lama", value="Histori", key="rename_old")
        with colR2:
            new_tab = st.text_input("Nama tab baru", value="September 2026", key="rename_new")
        with colR3:
            st.write("")
            st.write("")
            if st.button("Ganti Nama"):
                ok, msg = gsheets.rename_tab(st, old_tab, new_tab)
                (st.success if ok else st.error)(msg)

    master_df, MANUAL_ROWS = common.get_master_and_manual_rows()
    gsheets_ready = common.gsheets_ready()

    # ------------------------------------------------------------------
    # 1. Upload file mentah SAP (+ opsi mode histori)
    # ------------------------------------------------------------------
    st.header("1️⃣ Upload File Mentah SAP")

    use_history = False
    history_snapshot = None
    history_map = {}

    if gsheets_ready:
        snap_list_top = common.cached_snapshots()
        if snap_list_top:
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
    upload_report = None

    if use_history:
        snap_list = common.cached_snapshots()
        labels = [s["label"] for s in snap_list]
        # default: snapshot paling baru dengan Sesi SORE (kemarin sore), kalau tidak ada -> paling atas
        default_idx = next((i for i, s in enumerate(snap_list) if s["sesi"] == "SORE"), 0)
        chosen_label = st.selectbox(
            "Pilih snapshot histori (untuk OP105-OP166 & manual) — biasanya: kemarin, Sore:",
            labels, index=default_idx,
        )
        history_snapshot = next(s for s in snap_list if s["label"] == chosen_label)
        history_map = common.cached_history_map(history_snapshot["tanggal"], history_snapshot["sesi"])
        st.caption(f"Histori dipilih: **{chosen_label}** ({len(history_map)} Type tersimpan).")

        op100_only_file = st.file_uploader("Upload file OP100 hari ini (.xlsx)", type=["xlsx"])
        if op100_only_file:
            try:
                grouped, upload_report = common.cached_single_file_parse(op100_only_file)
                st.success("File OP100 dibaca.")
            except Exception as e:
                st.error(f"Gagal membaca file: {e}")
    else:
        st.caption("Upload file Posisi_Stock dari OP100, OP105, OP107, OP110, OP115, OP120, OP166 (boleh sebagian dulu).")
        raw_files = st.file_uploader("Upload file .xlsx", type=["xlsx"], accept_multiple_files=True)
        if raw_files:
            try:
                grouped, upload_report = common.cached_combine_dedup(raw_files)
                st.success(f"{len(raw_files)} file dibaca.")
            except Exception as e:
                st.error(f"Gagal membaca file: {e}")

    if upload_report is not None:
        _render_checklist(upload_report)

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
        st.session_state.pop("auto_save_msg", None)

        # Simpan otomatis ke histori Google Sheets begitu Generate selesai (point 5) —
        # SEMENTARA DIMATIKAN (AUTO_SAVE_HISTORI = False) atas permintaan Anda selama masa
        # testing, supaya data di tab Histori Google Sheets tidak ikut berubah/bertambah
        # tanpa sengaja saat sedang coba-coba. Pakai tombol "Simpan ke Histori" manual di
        # bawah kalau memang mau menyimpan. Nyalakan lagi dengan ubah AUTO_SAVE_HISTORI
        # jadi True kalau fitur 'Ambil Data' sudah dipastikan aman.
        if AUTO_SAVE_HISTORI:
            if gsheets_ready:
                ok_save, msg_save = gsheets.append_history(st, report_df, tanggal_str, sesi)
                st.session_state["auto_save_msg"] = (ok_save, f"☁️ Tersimpan otomatis — {msg_save}" if ok_save else msg_save)
            else:
                st.session_state["auto_save_msg"] = (
                    False,
                    "Google Sheets belum dikonfigurasi — laporan ini TIDAK tersimpan otomatis, "
                    "jadi belum bisa diambil lagi lewat halaman 'Ambil Data'.",
                )

    if not can_generate:
        st.info("⬆️ Upload file yang diperlukan dulu di Langkah 1.")

    if "report_df" in st.session_state:
        report_df = st.session_state["report_df"]
        unmatched_df = st.session_state.get("unmatched_df", pd.DataFrame())
        missing_types = st.session_state.get("missing_types", [])

        if "auto_save_msg" in st.session_state:
            ok_save, msg_save = st.session_state["auto_save_msg"]
            (st.caption if ok_save else st.warning)(msg_save)

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
