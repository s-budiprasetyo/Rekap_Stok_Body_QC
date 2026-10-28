import base64
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd

import common
import engine
import gsheets

# Simpan otomatis ke Histori Google Sheets setiap kali Generate berhasil -- DIMATIKAN lagi
# (29 Okt 2026) atas permintaan Anda: simpan-otomatis tanpa konfirmasi dirasa kurang aman
# kalau aplikasi dipakai orang lain yang perilakunya di luar prediksi. Sebagai gantinya,
# konfirmasi simpan sekarang muncul sebagai dialog tiap kali tombol Download diklik (lihat
# _confirm_save_and_download di bawah) -- jadi keputusan simpan selalu eksplisit per unduhan.
AUTO_SAVE_HISTORI = False


def _trigger_browser_download(data: bytes, filename: str, mime: str):
    """Memicu unduhan file browser secara otomatis (tanpa klik tombol native kedua) lewat
    link <a download> yang di-klik otomatis via JS, dibungkus di iframe komponen Streamlit.
    Ini trik yang umum dipakai krn Streamlit sendiri tidak punya cara memicu unduhan murni
    dari kode server. CATATAN: beberapa browser/pengaturan keamanan bisa saja memblokir
    auto-klik di dalam iframe -- makanya tombol unduh manual tetap ditampilkan sebagai
    cadangan tepat di bawahnya, supaya file tidak pernah "hilang begitu saja" kalau trik
    ini gagal jalan."""
    b64 = base64.b64encode(data).decode()
    html = f"""
    <html><body>
    <a id="auto-dl-link" href="data:{mime};base64,{b64}" download="{filename}" style="display:none"></a>
    <script>document.getElementById('auto-dl-link').click();</script>
    </body></html>
    """
    components.html(html, height=0, width=0)


@st.dialog("Simpan sebagai histori?")
def _confirm_save_and_download(data_bytes: bytes, filename: str, mime: str,
                                report_df: pd.DataFrame, tanggal_str: str, sesi: str,
                                gsheets_ready: bool):
    st.write(f"Simpan laporan **{tanggal_str} ({sesi})** ini ke histori Stok Body (Google Sheets) juga?")
    col1, col2 = st.columns(2)
    choice = None
    with col1:
        if st.button("✅ Ya, simpan", use_container_width=True, type="primary"):
            choice = True
    with col2:
        if st.button("❌ Tidak, download saja", use_container_width=True):
            choice = False

    if choice is not None:
        if choice:
            if gsheets_ready:
                ok, msg = gsheets.append_history(st, report_df, tanggal_str, sesi)
                (st.success if ok else st.error)(msg)
            else:
                st.warning("Google Sheets belum dikonfigurasi — tidak bisa disimpan.")
        _trigger_browser_download(data_bytes, filename, mime)
        st.caption("Unduhan dimulai otomatis. Kalau tidak jalan, klik tombol di bawah ini:")
        st.download_button("⬇️ Download sekarang", data=data_bytes, file_name=filename, mime=mime,
                            use_container_width=True)


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

    # Catatan penting: key widget number_input di bawah SENGAJA menyertakan mode +
    # snapshot histori yang dipilih (bukan cuma nama field). Di Streamlit, argumen
    # `value=` pada widget HANYA dipakai saat widget dgn key itu pertama kali dibuat di
    # sesi browser ini -- rerun berikutnya akan selalu pakai nilai lama di session_state,
    # walau default_val yang baru dihitung berbeda. Kalau key-nya cuma "manual_<label>"
    # (sama terus apapun mode/snapshot-nya), begitu widget itu sempat ter-render SEKALI
    # dengan nilai 0 (misal waktu histori masih rusak), field itu akan "macet" di 0
    # selamanya walau histori sudah diperbaiki -- persis masalah CRJP/U104ST yang dilaporkan.
    # Dengan menambahkan mode+snapshot ke key, ganti snapshot/mode = widget baru = default
    # yang baru dihitung benar-benar dipakai lagi.
    history_key_part = history_snapshot["label"] if (use_history and history_snapshot) else "no-history"

    manual_values = {}
    cols = st.columns(3)
    for idx, label in enumerate(MANUAL_ROWS):
        default_val = 0
        if use_history and label in history_map:
            target_col = master_df.loc[master_df["RowLabel"] == label, "ManualTargetColumn"].iloc[0]
            if target_col:
                default_val = int(history_map[label].get(target_col, 0) or 0)
        with cols[idx % 3]:
            manual_values[label] = st.number_input(
                label, min_value=0, value=default_val, step=1,
                key=f"manual_{label}__{use_history}__{history_key_part}",
            )

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

        # Catatan: tidak ada simpan-otomatis di sini lagi. Simpan ke histori sekarang
        # ditanyakan lewat dialog konfirmasi tiap kali tombol Download (PNG/Excel) di
        # bawah diklik -- lihat _confirm_save_and_download di atas.

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

        tanggal_cur = st.session_state["tanggal_str"]
        sesi_cur = st.session_state["sesi"]
        png_bytes = common.cached_render_png(report_df, tanggal_cur, sesi_cur)
        xlsx_bytes = common.cached_build_excel(report_df, tanggal_cur, sesi_cur)
        st.image(png_bytes, caption="Preview Laporan")

        png_filename = f"Stok_Body_{tanggal_cur.replace(' ', '_')}_{sesi_cur}.png"
        xlsx_filename = f"Stok_Body_{tanggal_cur.replace(' ', '_')}_{sesi_cur}.xlsx"

        colA, colB, colC = st.columns(3)
        with colA:
            if st.button("⬇️ Download Gambar (PNG)", use_container_width=True):
                _confirm_save_and_download(
                    png_bytes, png_filename, "image/png",
                    report_df, tanggal_cur, sesi_cur, gsheets_ready,
                )
        with colB:
            if st.button("⬇️ Jadikan Excel", use_container_width=True):
                _confirm_save_and_download(
                    xlsx_bytes, xlsx_filename,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    report_df, tanggal_cur, sesi_cur, gsheets_ready,
                )
        with colC:
            if gsheets_ready:
                if st.button("☁️ Simpan ke Histori saja (tanpa download)", use_container_width=True):
                    ok, msg = gsheets.append_history(st, report_df, tanggal_cur, sesi_cur)
                    (st.success if ok else st.error)(msg)
            else:
                st.caption("Histori Google Sheets belum dikonfigurasi (opsional, lihat README).")

        with st.expander("Lihat tabel data laporan"):
            st.dataframe(report_df, use_container_width=True)
