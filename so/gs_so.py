"""
Simpan histori Stok Opname -- TERPISAH dari histori Stok Body Harian:
 * tab bernama 'SO <Bulan Tahun>' (mis. 'SO Oktober 2026'), jadi tidak bercampur dengan
   tab 'Oktober 2026' milik Rekap Harian (yang dibaca fitur Ambil Data / pilih histori);
 * kalau di Secrets ada `gsheet_name_so`, dipakai Google Sheet (file) lain khusus SO;
   kalau tidak ada, dipakai file yang sama dgn `gsheet_name` tapi tab-nya terpisah.
Perilaku UPSERT: Tanggal yang sama ditimpa, tidak dobel.
"""
import datetime
import pandas as pd

import gsheets

HEADER = gsheets.HEADER_ROW
DATA_COLS = gsheets.DATA_COLS


def tab_name(tanggal_str: str) -> str:
    return "SO " + gsheets._tab_name_for_tanggal_str(tanggal_str)


def _spreadsheet_name(st) -> str:
    try:
        if "gsheet_name_so" in st.secrets and str(st.secrets["gsheet_name_so"]).strip():
            return str(st.secrets["gsheet_name_so"]).strip()
    except Exception:
        pass
    return st.secrets["gsheet_name"]


def save_so(st, report_df: pd.DataFrame, tanggal_str: str, sesi: str = "SORE"):
    if not gsheets.is_configured(st):
        return False, "Google Sheets belum dikonfigurasi (lihat Secrets)."
    try:
        client = gsheets._get_client(st)
        sh = client.open(_spreadsheet_name(st))
        name = tab_name(tanggal_str)
        try:
            ws = sh.worksheet(name)
        except Exception:
            ws = sh.add_worksheet(title=name, rows=1000, cols=20)
        existing = ws.get_all_values()
        kept, removed = [], 0
        for row in existing[1:]:
            if len(row) > 2 and row[1] == tanggal_str and row[2] == sesi:
                removed += 1
            else:
                kept.append(row)
        now = datetime.datetime.now(gsheets.WIB).isoformat(timespec="seconds")
        new_rows = [
            [now, tanggal_str, sesi, r["RowLabel"], r["RowGroup"]]
            + [int(r.get(c, 0) or 0) for c in DATA_COLS] + [int(r["GrandTotal"])]
            for _, r in report_df.iterrows()
        ]
        try:  # terapkan lagi cell tab bulan yang pernah Anda ubah untuk tanggal ini
            ovr_ws = sh.worksheet(ovr_name(name))
            new_rows = apply_overrides_to_rows(new_rows, load_overrides(ovr_ws), tanggal_str)
        except Exception:
            pass
        ws.clear()
        ws.update(values=[HEADER] + kept + new_rows, value_input_option="USER_ENTERED")
        protect_sheet(st, sh, ws)
        refresh_basis(st, tanggal_str)  # basis = angka tab bulan saat ini
        msg = f"{len(new_rows)} baris tersimpan ke tab '{name}' ({sh.title})"
        if removed:
            msg += f", menimpa {removed} baris lama {tanggal_str}"
        return True, msg + "."
    except Exception as e:
        return False, f"Gagal menyimpan histori Stok Opname: {e}"


def _open(st):
    return gsheets._get_client(st).open(_spreadsheet_name(st))


def list_so_snapshots(st):
    """Daftar tanggal Stok Opname yang tersimpan: [{'tanggal':str,'date_obj':date,'label':str}], terbaru dulu.
    Membaca semua tab berawalan 'SO '."""
    import engine
    if not gsheets.is_configured(st):
        return []
    try:
        sh = _open(st)
        tabs = [w for w in sh.worksheets() if w.title.startswith("SO ")]
    except Exception:
        return []
    seen = set()
    for ws in tabs:
        try:
            for row in ws.get_all_values()[1:]:
                if len(row) > 2 and row[1].strip():
                    seen.add(row[1].strip())
        except Exception:
            continue
    items = []
    for t in seen:
        d = engine.parse_id_date(t)
        if d:
            items.append({"tanggal": t, "date_obj": d, "label": f"{t} ( SORE )"})
    items.sort(key=lambda x: x["date_obj"], reverse=True)
    return items


def fetch_so_report(st, tanggal_str: str):
    """Susun ulang report_df (RowLabel, RowGroup, kolom data, GrandTotal) dari data tersimpan. {} kalau kosong."""
    if not gsheets.is_configured(st):
        return None
    try:
        sh = _open(st)
        records = sh.worksheet(tab_name(tanggal_str)).get_all_records()
    except Exception:
        return None
    rows = {}
    for r in records:
        if str(r.get("Tanggal", "")).strip() != tanggal_str:
            continue
        label = str(r.get("RowLabel", "")).strip()
        if not label:
            continue
        rows[label] = {"RowLabel": label, "RowGroup": r.get("RowGroup", ""),
                       **{c: int(r.get(c, 0) or 0) for c in DATA_COLS},
                       "GrandTotal": int(r.get("GrandTotal", 0) or 0)}
    if not rows:
        return None
    return pd.DataFrame(list(rows.values()))


# ======================================================================
# Sinkron perubahan manual di tab bulan ('Oktober 2026') -> tab SO
# ======================================================================
# Aturan (permintaan pengguna):
#  * Tab SO punya angka sendiri; angka SO dan Harian BOLEH berbeda (selisih dibiarkan).
#  * Hanya CELL yang diubah di tab bulan (sesi SORE) yang ikut mengubah cell yang sama
#    (baris Type + kolom OP) di tab SO pada tanggal yang sama: nilainya DISALIN
#    (mis. Harian 50->48 membuat SO 51->48; 50->53 membuat SO 53). Cell lain tidak disentuh.
#  * Klik / mode edit / Esc / ketik ulang angka yang sama = tidak ada perubahan nilai = SO tidak berubah.
#  * Cara mendeteksi "cell yang diubah": tab 'Basis SO <Bulan>' menyimpan angka tab bulan
#    PADA SAAT terakhir diperiksa/disimpan. Cell yang angkanya sekarang beda dari basis = diubah.
#  * Tanggal SO lama yang belum punya basis: baris manual (CRJP/U104ST/S90097/TX60V1)
#    mengikuti tab bulan (sesuai aturan lama), lalu basis dibuat. Baris lain tidak diubah.
import hashlib

_NC = len(DATA_COLS)
_C0 = 5
_CG = _C0 + _NC


def basis_name(so_title: str) -> str:
    return "Basis " + so_title


def _int(x):
    try:
        return int(float(str(x).replace(",", "").strip() or 0))
    except Exception:
        return 0


def _pad(row):
    row = list(row)
    return (row + [""] * (_CG + 1 - len(row)))[:_CG + 1] if len(row) < _CG + 1 else row


def _norm(vals):
    return [[str(c).strip() for c in r] for r in (_strip_tail(r) for r in vals)]


def _strip_tail(r):
    r = list(r)
    while r and str(r[-1]).strip() == "":
        r.pop()
    return r


def _index(rows):
    return {(r[1].strip(), r[3].strip()): i for i, r in enumerate(rows) if len(r) > 3 and r[3].strip()}


def _write(ws, header, rows):
    ws.clear()
    ws.update(values=[header] + rows, value_input_option="USER_ENTERED")


def _harian_sore(har_vals):
    out = {}
    for r in har_vals[1:]:
        if len(r) > 3 and r[2].strip() == "SORE" and r[3].strip():
            out[(r[1].strip(), r[3].strip())] = _pad(r)
    return out


def sync_month(so_ws, har_ws, basis_ws, manual_labels, ovr=None):
    """Sinkronkan 1 bulan. Return (changed, har_vals, so_vals_final)."""
    so_vals = so_ws.get_all_values()
    har_vals = har_ws.get_all_values()
    base_vals = basis_ws.get_all_values()
    so_hdr = so_vals[0] if so_vals else HEADER
    so_rows = [_pad(r) for r in so_vals[1:]]
    base_hdr = base_vals[0] if base_vals else HEADER
    base_rows = [_pad(r) for r in base_vals[1:]]
    har = _harian_sore(har_vals)
    so_idx, base_idx = _index(so_rows), _index(base_rows)
    so_changed = base_changed = False

    for tgl in {t for (t, _) in so_idx}:
        labels = [l for (t, l) in har if t == tgl]
        if not labels:
            continue
        has_base = any((tgl, l) in base_idx for l in labels)
        if not has_base:
            for l in labels:
                h = har[(tgl, l)]
                if l in manual_labels and (tgl, l) in so_idx:
                    s = so_rows[so_idx[(tgl, l)]]
                    if [_int(x) for x in s[_C0:_CG + 1]] != [_int(x) for x in h[_C0:_CG + 1]]:
                        s[_C0:_CG + 1] = h[_C0:_CG + 1]
                        so_changed = True
                base_rows.append(list(h))
                base_idx[(tgl, l)] = len(base_rows) - 1
                base_changed = True
            continue
        for l in labels:
            h = har[(tgl, l)]
            if (tgl, l) not in base_idx:
                base_rows.append(list(h))
                base_idx[(tgl, l)] = len(base_rows) - 1
                base_changed = True
                continue
            b = base_rows[base_idx[(tgl, l)]]
            s = so_rows[so_idx[(tgl, l)]] if (tgl, l) in so_idx else None
            data_changed, delta = False, 0
            for i in range(_C0, _CG):
                hv, bv = _int(h[i]), _int(b[i])
                if hv != bv:
                    if s is not None:
                        delta += hv - _int(s[i])          # total SO menyesuaikan dgn cell yang diganti
                        s[i] = hv                          # SALIN nilai baru dari tab bulan (bukan selisih)
                        data_changed = True
                        if ovr is not None:
                            ovr[(tgl, l, DATA_COLS[i - _C0])] = hv
                    b[i] = hv
                    base_changed = True
            hg, bg = _int(h[_CG]), _int(b[_CG])
            if s is not None and data_changed:
                s[_CG] = _int(s[_CG]) + delta      # GrandTotal ikut menyesuaikan cell yang diganti
                so_changed = True
            elif hg != bg and s is not None:
                s[_CG] = hg                          # GrandTotal di tab bulan diubah langsung
                so_changed = True
            if hg != bg:
                b[_CG] = hg
                base_changed = True

    if so_changed:
        _write(so_ws, so_hdr, so_rows)
    if base_changed:
        _write(basis_ws, base_hdr, base_rows)
    final_so = so_vals if not so_changed else [so_hdr] + so_rows
    return so_changed, har_vals, final_so


def ovr_name(so_title: str) -> str:
    return "Koreksi " + so_title


OVR_HEADER = ["Tanggal", "RowLabel", "Kolom", "Nilai"]


def load_overrides(ws) -> dict:
    """{(Tanggal, RowLabel, Kolom): nilai} -- cell tab bulan yang pernah Anda ubah (dipakai lagi
    kalau SO tanggal itu disimpan ulang, supaya perubahan Anda tidak hilang)."""
    out = {}
    for r in ws.get_all_values()[1:]:
        if len(r) > 3 and r[0].strip() and r[1].strip() and r[2].strip():
            out[(r[0].strip(), r[1].strip(), r[2].strip())] = _int(r[3])
    return out


def save_overrides(ws, ovr: dict):
    rows = [[t, l, c, v] for (t, l, c), v in sorted(ovr.items())]
    _write(ws, OVR_HEADER, rows)


def apply_overrides_to_rows(rows, ovr: dict, tanggal_str: str):
    """rows = baris SO (list) untuk 1 tanggal; terapkan nilai Koreksi, total ikut menyesuaikan."""
    for r in rows:
        for ci, col in enumerate(DATA_COLS):
            key = (tanggal_str, str(r[3]).strip(), col)
            if key in ovr:
                i = _C0 + ci
                r[_CG] = _int(r[_CG]) + (ovr[key] - _int(r[i]))
                r[i] = ovr[key]
    return rows


_PROTECTED = set()


def protect_sheet(st, sh, ws):
    """Kunci tab: hanya akun aplikasi (service account) yang boleh menulis. Dicoba sekali per tab
    per proses; kalau gagal (mis. sudah terkunci) diabaikan -- ini pengaman, bukan syarat jalan."""
    key = (sh.id, ws.id)
    if key in _PROTECTED:
        return
    _PROTECTED.add(key)
    try:
        meta = sh.fetch_sheet_metadata()
        for sheet in meta.get("sheets", []):
            if sheet["properties"]["sheetId"] == ws.id and sheet.get("protectedRanges"):
                return
        email = str(st.secrets["gcp_service_account"]["client_email"])
        sh.batch_update({"requests": [{"addProtectedRange": {"protectedRange": {
            "range": {"sheetId": ws.id},
            "description": "Dikunci oleh aplikasi - ubah angka di tab bulan, bukan di sini.",
            "warningOnly": False, "editors": {"users": [email]}}}}]})
    except Exception:
        pass


def _get_or_create(sh, title):
    try:
        return sh.worksheet(title)
    except Exception:
        return sh.add_worksheet(title=title, rows=1000, cols=20)


def sync_so_from_harian(st, manual_labels):
    """Dipanggil berkala. Return (fingerprint, changed). (None, False) kalau tak bisa."""
    if not gsheets.is_configured(st):
        return None, False
    try:
        sh_h = gsheets._get_client(st).open(st.secrets["gsheet_name"])
        sh_s = _open(st)
        titles_h = {w.title for w in sh_h.worksheets()}
        titles_s = {w.title: w for w in sh_s.worksheets()}
        fp = hashlib.md5()
        changed = False
        for month in gsheets._current_and_previous_tab_names():
            so_title = "SO " + month
            if month not in titles_h:
                continue
            har_ws = sh_h.worksheet(month)
            if so_title not in titles_s:
                fp.update(repr(_norm(har_ws.get_all_values())).encode())
                continue
            basis_ws = titles_s.get(basis_name(so_title)) or _get_or_create(sh_s, basis_name(so_title))
            ovr_ws = titles_s.get(ovr_name(so_title)) or _get_or_create(sh_s, ovr_name(so_title))
            ovr = load_overrides(ovr_ws)
            ovr_before = dict(ovr)
            ch, har_vals, so_vals = sync_month(titles_s[so_title], har_ws, basis_ws, set(manual_labels), ovr)
            if ovr != ovr_before:
                save_overrides(ovr_ws, ovr)
            changed |= ch
            for w in (titles_s[so_title], basis_ws, ovr_ws):
                protect_sheet(st, sh_s, w)
            fp.update(repr(_norm(har_vals)).encode())
            fp.update(repr(_norm(so_vals)).encode())
        return fp.hexdigest(), changed
    except Exception:
        return None, False


def refresh_basis(st, tanggal_str: str, only_if_exists: bool = False):
    """Set basis tanggal ini = angka tab bulan (sesi SORE) SEKARANG. Dipanggil setelah SO atau
    Harian-SORE disimpan lewat aplikasi, supaya penyimpanan lewat aplikasi tidak dianggap 'edit manual'."""
    try:
        sh_h = gsheets._get_client(st).open(st.secrets["gsheet_name"])
        sh_s = _open(st)
        month = gsheets._tab_name_for_tanggal_str(tanggal_str)
        so_title = "SO " + month
        try:
            sh_s.worksheet(so_title)
        except Exception:
            return  # belum ada SO bulan ini
        har = _harian_sore(sh_h.worksheet(month).get_all_values())
        basis_ws = _get_or_create(sh_s, basis_name(so_title))
        vals = basis_ws.get_all_values()
        hdr = vals[0] if vals else HEADER
        rows = [_pad(r) for r in vals[1:]]
        has = any(r[1].strip() == tanggal_str for r in rows)
        if only_if_exists and not has:
            return
        rows = [r for r in rows if r[1].strip() != tanggal_str]
        rows += [list(r) for (t, _), r in har.items() if t == tanggal_str]
        _write(basis_ws, hdr, rows)
    except Exception:
        pass
