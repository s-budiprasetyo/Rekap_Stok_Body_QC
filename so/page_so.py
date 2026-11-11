"""Halaman Stok Opname (st.navigation). Semua logika hitung ada di so/engine_so.py."""
import hashlib
import streamlit as st
import pandas as pd

import common
import dl_helper
import engine
import ui_calendar
from so import cached, engine_so, gs_so

SESI_SO = "SORE"  # semua laporan stok opname masuk sesi sore
MANUAL_NOTE = "CRJP, U104ST, S90097, TX60V1"

CSS = """
<style>
.st-key-so_custom_btn button{
  background:linear-gradient(135deg,#4f46e5,#06b6d4)!important;color:#fff!important;border:0!important;
  min-height:3.4rem;font-size:1.15rem!important;font-weight:700!important;border-radius:.8rem!important;
  box-shadow:0 4px 12px rgba(79,70,229,.35);letter-spacing:.02em}
.st-key-so_custom_btn button:hover{filter:brightness(1.1);transform:translateY(-1px)}
.st-key-so_custom_btn button:disabled{background:#cbd5e1!important;box-shadow:none;color:#64748b!important}
.st-key-so_custom_btn button [data-testid="stMarkdownContainer"] p{font-size:1.25rem!important;font-weight:700!important}
.st-key-so_custom_btn button [data-testid="stMarkdownContainer"] span[role="img"]{font-size:2.3rem!important;margin-right:.2rem;vertical-align:middle!important;line-height:1}
.st-key-so_gen_btn button{min-height:3.4rem;border-radius:.8rem!important}
.st-key-so_gen_btn button [data-testid="stMarkdownContainer"] p{font-size:1.25rem!important;font-weight:700!important}
</style>
"""


@st.cache_data(show_spinner=False)
def _cached_read_so(file):
    return engine_so.read_so_file(file)


@st.dialog("Simpan sebagai histori Stok Opname?", on_dismiss="rerun")
def _save_dialog(data_bytes, filename, mime, report_df, tanggal_str, gs_ready):
    st.session_state["_dlg_busy"] = True
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
                with st.spinner("Menyimpan ke histori..."):
                    ok, msg = gs_so.save_so(st, report_df, tanggal_str, SESI_SO)
                dl_helper.flash("success" if ok else "error", msg)
            else:
                dl_helper.flash("warning", "Google Sheets belum dikonfigurasi — tidak bisa disimpan.")
        dl_helper.queue_download(data_bytes, filename, mime)
        st.rerun()  # dialog langsung hilang, unduhan jalan otomatis


def _checklist(info: dict):
    lines = []
    for op in engine_so.OPS:
        n = info["per_op"].get(op, 0)
        lines.append(f"✅ **{op}** — {n} pcs" if n else f"⬜ {op} — tidak ada data di file ini")
    st.markdown("  \n".join(lines))
    if info["unknown_lokasi"]:
        st.warning("Lokasi di luar OP yang dikenal (diabaikan): " +
                   ", ".join(f"{k} ({v} pcs)" for k, v in info["unknown_lokasi"].items()))


def _dialog_css(maximized: bool):
    if maximized:
        box = "width:96vw!important;max-width:96vw!important;height:92vh!important;max-height:92vh!important;"
    else:
        box = "width:min(900px,94vw)!important;max-width:94vw!important;"
    st.markdown(f"""<style>
[data-testid="stDialog"]:has(.st-key-so_dlg_ctl) div:has(> section[role="dialog"]){{{box}position:relative;}}
[data-testid="stDialog"]:has(.st-key-so_dlg_ctl) section[role="dialog"]{{{'height:100%;overflow:auto;' if maximized else ''}}}
.st-key-so_dlg_ctl{{position:absolute;top:.85rem;right:3.6rem;width:auto!important;z-index:10}}
.st-key-so_dlg_ctl button{{min-height:2rem;padding:.1rem .4rem}}
.st-key-so_dlg_ctl button span[role="img"]{{font-size:1.7rem}}
</style>""", unsafe_allow_html=True)


def _open_customize(grouped, so_date):
    applied = st.session_state.get("so_corrections", engine_so.empty_corrections())
    st.session_state["so_corr_table"] = applied
    st.session_state["so_corr_origin"] = "applied"
    st.session_state["so_corr_ver"] = st.session_state.get("so_corr_ver", 0) + 1
    _customize_dialog(sorted(grouped["Type"].unique()), sorted(grouped["Warna"].unique()), so_date)


@st.dialog("⚙️🔧 Customize — Pengurang Stok", width="large", on_dismiss="rerun")
def _customize_dialog(type_opts, warna_opts, so_date):
    st.session_state["_dlg_busy"] = True
    maxed = st.session_state.get("so_dlg_max", False)
    _dialog_css(maxed)
    with st.container(key="so_dlg_ctl"):
        if st.button(":material/fullscreen_exit:" if maxed else ":material/fullscreen:",
                     key="so_dlg_max_btn", type="tertiary", help="Perkecil" if maxed else "Perbesar (maximize)"):
            st.session_state["so_dlg_max"] = not maxed
            st.rerun(scope="fragment")

    applied = st.session_state.get("so_corrections", engine_so.empty_corrections())
    source = st.session_state.get("so_corr_source")
    st.caption("Angka di tabel MENGURANGI stok hasil upload Stok Opname (mis. barang yang sudah diterima Packing). "
               "Isi manual di tabel, atau upload Excel Surat Jalan.")
    n_a, p_a = engine_so.corr_summary(applied)
    if n_a:
        st.info(f"Pengurang yang sedang dipakai: **{(source or {}).get('label', 'input manual')}** — {n_a} baris, −{p_a} pcs.")
    else:
        st.caption("Belum ada pengurang yang dipakai.")

    up = st.file_uploader("Upload Excel pengurang (opsional)", type=["xlsx"], key="so_corr_file")
    default_op = st.selectbox(
        "Operation untuk baris dari Excel (dipakai kalau di Excel tidak ada kolom operation)",
        engine_so.OPS, index=engine_so.OPS.index("OP166"), key="so_corr_defop")

    seed_key = st.session_state.get("so_corr_origin", "applied")
    new_df, new_sig, new_label = None, None, None
    if up is not None:
        raw = up.getvalue()
        fid = hashlib.sha1(raw).hexdigest()[:8]
        try:
            dates, loks = engine_so.corr_file_options(up)
        except Exception as e:
            dates, loks = [], []
            st.error(f"Gagal membaca Excel pengurang: {e}")
        sel_date, sel_lok = None, None
        if dates:
            idx = dates.index(so_date) if so_date in dates else len(dates) - 1
            sel_date = st.selectbox("Tanggal posting yang dipakai (hanya tanggal ini yang mengurangi stok)",
                                    dates, index=idx, format_func=ui_calendar.fmt_id, key=f"so_corr_date_{fid}")
        if loks:
            sel_lok = st.multiselect("Lokasi awal yang dipakai (QCSK = Shuttle Kiln, biasanya tidak dipakai)",
                                     loks, default=["QC"] if "QC" in loks else loks, key=f"so_corr_lok_{fid}")
        try:
            new_df = engine_so.read_correction_excel(up, default_op, sel_date, sel_lok)
            new_sig = engine_so.file_signature(raw, sel_date, sel_lok, default_op)
            new_label = f"{up.name} ({ui_calendar.fmt_id(sel_date) if sel_date else 'semua tanggal'})"
        except Exception as e:
            st.error(f"Gagal membaca Excel pengurang: {e}")

    if new_df is not None:
        n_n, p_n = engine_so.corr_summary(new_df)
        same = bool(n_a) and (source or {}).get("sig") == new_sig and engine_so.tables_equal(new_df, applied)
        if same:
            st.success("File ini **sudah dipakai** sebagai pengurang — tidak akan dihitung dua kali.")
            want = "applied"
        elif n_a:
            st.warning(f"Sudah ada pengurang ({n_a} baris, −{p_a} pcs). File baru ini berisi {n_n} baris, −{p_n} pcs. "
                       "Pilih salah satu (keduanya TIDAK dijumlahkan):")
            pick = st.radio("Pengurang yang dipakai", ["Pakai data baru dari file ini (mengganti yang lama)",
                                                       "Tetap pakai pengurang yang sudah ada"],
                            key=f"so_corr_pick_{new_sig}", label_visibility="collapsed")
            want = new_sig if pick.startswith("Pakai data baru") else "applied"
        else:
            st.success(f"File dibaca: {n_n} baris, −{p_n} pcs. Cek tabel di bawah, lalu klik Terapkan.")
            want = new_sig
        if want != seed_key:  # ganti isi tabel (BUKAN menjumlah) sebelum editor dibuat
            st.session_state["so_corr_table"] = new_df if want == new_sig else applied
            st.session_state["so_corr_origin"] = want
            st.session_state["so_corr_ver"] += 1
        pending_label, pending_sig = new_label, new_sig
    else:
        pending_label, pending_sig = "input manual", None

    table = st.session_state["so_corr_table"]
    t_opts = sorted(set(type_opts) | set(table["Type"].dropna().astype(str)))
    w_opts = sorted(set(warna_opts) | set(table["Kode Warna"].dropna().astype(str)))
    edited = st.data_editor(
        table, key=f"so_corr_editor_{st.session_state['so_corr_ver']}", num_rows="dynamic",
        use_container_width=True, hide_index=True, height=520 if maxed else 300,
        column_config={
            "Type": st.column_config.SelectboxColumn("Type", options=t_opts, required=True),
            "Kode Warna": st.column_config.SelectboxColumn("Kode Warna", options=w_opts, required=True),
            "Jenis Forming": st.column_config.SelectboxColumn("Jenis Forming", options=["BC", "HP", "MC"], required=True),
            "Kode Pabrik": st.column_config.SelectboxColumn("Kode Pabrik", options=["1", "2"], required=True),
            "Operation": st.column_config.SelectboxColumn("Operation", options=engine_so.OPS, required=True),
            "Jumlah": st.column_config.NumberColumn("Jumlah", min_value=1, step=1, format="%d", required=True),
        })
    clean = engine_so.clean_corrections(edited)
    if not engine_so.tables_equal(clean, applied):
        n_c, p_c = engine_so.corr_summary(clean)
        st.warning(f"Tabel ini ({n_c} baris, −{p_c} pcs) **belum diterapkan** — klik Terapkan agar laporan berkurang.")
    c1, c2 = st.columns(2)
    if c1.button("✅ Terapkan", type="primary", use_container_width=True):
        used_new = st.session_state.get("so_corr_origin") not in ("applied",) and pending_sig is not None
        st.session_state["so_corrections"] = clean
        st.session_state["so_corr_source"] = (
            {"sig": pending_sig, "label": pending_label} if used_new else
            ({**(source or {}), "label": (source or {}).get("label", "input manual") + " (diedit)"} if n_a and not engine_so.tables_equal(clean, applied) else
             (source or {"sig": None, "label": "input manual"})))
        n_c, p_c = engine_so.corr_summary(clean)
        dl_helper.flash("success", f"Pengurang diterapkan: {n_c} baris, −{p_c} pcs.")
        st.rerun()
    if c2.button("🗑️ Hapus semua pengurang", use_container_width=True):
        st.session_state["so_corrections"] = engine_so.empty_corrections()
        st.session_state["so_corr_source"] = None
        dl_helper.flash("success", "Semua pengurang dihapus.")
        st.rerun()


@st.fragment
def _result_section(master_df):
    """Preview + Download. Fragment: klik di sini tidak me-reload seluruh halaman."""
    inp = st.session_state.get("so_inputs")
    if not inp:
        return
    # Nilai CRJP/U104ST/S90097/TX60V1 dibaca ULANG dari Google Sheet tiap render (bukan dikunci
    # saat Generate), jadi perubahan di Sheet langsung ikut.
    manual_values = inp["manual_values"]
    if inp.get("hist_tanggal"):
        manual_values = engine_so.manual_from_history(
            master_df, common.cached_history_map(inp["hist_tanggal"], SESI_SO))
    corr = st.session_state.get("so_corrections", engine_so.empty_corrections())
    grouped_adj, log = engine_so.apply_corrections(inp["grouped"], corr)  # selalu dari stok ASLI
    report_df, unmatched = engine_so.build_so_report(grouped_adj, master_df, manual_values)
    tanggal = inp["tanggal_str"]

    if not unmatched.empty:
        with st.expander(f"⚠️ {len(unmatched)} baris data SAP TIDAK cocok ke master"):
            st.dataframe(unmatched[["Lokasi", "Type", "Warna", "KodeNorm", "Jenis Forming", "Qty"]], use_container_width=True)

    png = cached.so_png(report_df, tanggal)
    xlsx = cached.so_xlsx(report_df, tanggal)
    st.image(png, caption=f"Preview — {tanggal} ({SESI_SO})")

    if len(log):
        total_kurang = int(log["Dikurangi"].sum())
        with st.expander(f"Pengurang diterapkan: {len(log)} baris, total −{total_kurang} pcs"):
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
    st.session_state["_dlg_busy"] = False
    dl_helper.run_pending()
    common.sheet_watcher()
    common.watch()
    st.markdown(CSS, unsafe_allow_html=True)
    st.title("📦 Stok Opname")
    st.caption("Upload 1 file 'Report Stock Opname', pilih tanggal histori, lalu Generate. Semua laporan masuk sesi SORE.")

    master_df, _ = common.get_master_and_manual_rows()
    gs_ready = common.gsheets_ready()

    # 1. Upload
    st.header("1️⃣ Upload File Stok Opname")
    so_file = st.file_uploader("Upload file Report Stock Opname (.xlsx) — cukup 1 file", type=["xlsx"], key="so_file")
    grouped, info = None, None
    sig = hashlib.sha1(so_file.getvalue()).hexdigest() if so_file is not None else None
    if st.session_state.get("so_file_sig") != sig:  # file SO berganti/dihapus -> bersihkan pengurang lama
        st.session_state["so_file_sig"] = sig
        for k in ("so_inputs", "so_corrections", "so_corr_source", "so_corr_table", "so_corr_origin"):
            st.session_state.pop(k, None)
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
            if hist_date:
                common.watch(("harian", ui_calendar.fmt_id(hist_date)))
            hmap = common.cached_history_map(ui_calendar.fmt_id(hist_date), SESI_SO) if hist_date else {}
            manual_values = engine_so.manual_from_history(master_df, hmap)
            with st.expander(f"Nilai {MANUAL_NOTE} yang diambil (otomatis ikut perubahan di Google Sheet)"):
                _tc = master_df.set_index("RowLabel")["ManualTargetColumn"].to_dict()
                st.dataframe(pd.DataFrame({"Type": list(manual_values), "Kolom Sheet": [_tc.get(k, "") or "(langsung total)" for k in manual_values],
                                           "Jumlah": list(manual_values.values())}),
                             hide_index=True, use_container_width=True)
        else:
            st.info(f"Belum ada histori sesi SORE di Google Sheets — nilai {MANUAL_NOTE} dianggap 0.")
    else:
        st.caption(f"(Google Sheets belum terhubung — nilai {MANUAL_NOTE} dianggap 0.)")

    # 3. Generate + Customize
    st.header("3️⃣ Generate")
    default_d = info["so_date"] if (info and info["so_date"]) else ui_calendar.today_wib()
    tgl_pick = ui_calendar.date_picker_apply(
        f"so_tgl_{default_d.isoformat()}", default=default_d, label="Tanggal laporan (otomatis dari file, bisa diubah)")
    tanggal_str = ui_calendar.fmt_id(tgl_pick)

    cg, cc = st.columns(2)
    with cg:
        with st.container(key="so_gen_btn"):
            gen = st.button("🖼️ Generate", type="primary", disabled=grouped is None, use_container_width=True)
    with cc:
        with st.container(key="so_custom_btn"):
            custom = st.button(":material/settings: :material/build: Customize", disabled=grouped is None,
                               use_container_width=True, help="Kurangi stok (mis. barang yang sudah diterima Packing)")
    if gen:
        st.session_state["so_inputs"] = {"grouped": grouped, "manual_values": manual_values, "tanggal_str": tanggal_str,
                                         "hist_tanggal": ui_calendar.fmt_id(hist_date) if hist_date else None}
    if custom:
        _open_customize(grouped, engine.parse_id_date(tanggal_str))

    corr = st.session_state.get("so_corrections", engine_so.empty_corrections())
    n_c, p_c = engine_so.corr_summary(corr)
    if n_c:
        label = (st.session_state.get("so_corr_source") or {}).get("label", "input manual")
        st.caption(f"🔧 Pengurang aktif: **{label}** — {n_c} baris, −{p_c} pcs")
    elif grouped is not None:
        st.caption("Belum ada pengurang (Customize) — laporan memakai stok apa adanya.")
    if grouped is None:
        st.info("⬆️ Upload file Stok Opname dulu di Langkah 1.")

    _result_section(master_df)
