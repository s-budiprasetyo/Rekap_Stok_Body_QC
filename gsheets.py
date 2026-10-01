"""
Simpan histori harian ke Google Sheets (opsional).
Kalau credential belum diisi di st.secrets, semua fungsi di sini
akan gagal dengan aman (return False) tanpa bikin app crash.
"""
import datetime
import pandas as pd


def is_configured(st) -> bool:
    try:
        return "gcp_service_account" in st.secrets and "gsheet_name" in st.secrets
    except Exception:
        return False


def _get_client(st):
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def test_connection(st) -> tuple[bool, str]:
    """Cek koneksi ke Google Sheets tanpa menulis data apapun. Return (sukses, pesan detail)."""
    if not is_configured(st):
        return False, (
            "Secrets belum diisi. Butuh 2 hal di Streamlit Secrets: "
            "`gsheet_name` (nama Google Sheet) dan blok `[gcp_service_account]` "
            "(isi dari file JSON service account)."
        )
    try:
        client = _get_client(st)
    except Exception as e:
        return False, f"Gagal baca credential service account. Cek format Secrets Anda. Detail: {e}"

    try:
        sh = client.open(st.secrets["gsheet_name"])
    except Exception as e:
        return False, (
            f"Berhasil login sebagai service account, TAPI gagal membuka Google Sheet "
            f"bernama '{st.secrets.get('gsheet_name')}'. Kemungkinan: (1) nama sheet di Secrets "
            f"tidak persis sama dengan nama file Google Sheet Anda, atau (2) sheet belum di-Share "
            f"ke email service account sebagai Editor. Detail teknis: {e}"
        )

    try:
        worksheets = [w.title for w in sh.worksheets()]
    except Exception as e:
        return False, f"Sheet ditemukan tapi gagal baca daftar tab di dalamnya. Detail: {e}"

    return True, f"✅ Koneksi berhasil! Sheet '{sh.title}' ditemukan, tab yang ada: {worksheets}"

DATA_COLS = ["OP100_PERIKSA", "OP105_PERBAIKAN", "OP107_TDKSET_IJP", "OP107_TDKSET_CRJP",
             "OP110_GERINDA", "OP115_CEKULANG", "OP120_TAP", "OP166_FITTINGST", "OP166_KIRIM"]
HEADER_ROW = ["Timestamp", "Tanggal", "Sesi", "RowLabel", "RowGroup"] + DATA_COLS + ["GrandTotal"]


def _tab_name_for_date(date_obj) -> str:
    """datetime.date -> 'Oktober 2026' (nama tab per bulan)."""
    import engine
    return f"{engine.BULAN_ID[date_obj.month - 1]} {date_obj.year}"


def _tab_name_for_tanggal_str(tanggal_str: str) -> str:
    """'16 September 2026' -> 'September 2026'. Fallback ke bulan ini kalau gagal parse."""
    import engine
    d = engine.parse_id_date(tanggal_str)
    if d is None:
        d = datetime.date.today()
    return _tab_name_for_date(d)


def _current_and_previous_tab_names():
    """Nama tab bulan ini & bulan sebelumnya — dipakai supaya histori tetap nyambung pas pergantian bulan."""
    today = datetime.date.today()
    cur = _tab_name_for_date(today)
    prev_month_last_day = today.replace(day=1) - datetime.timedelta(days=1)
    prev = _tab_name_for_date(prev_month_last_day)
    return cur, prev


def _ensure_header(ws) -> bool:
    """Pastikan baris 1 berisi header yang benar. Sisipkan kalau belum ada/salah. Return True kalau ada perubahan."""
    values = ws.get_all_values()
    if not values:
        ws.append_row(HEADER_ROW)
        return True
    first_row = values[0]
    if first_row[:1] != ["Timestamp"]:
        ws.insert_row(HEADER_ROW, index=1)
        return True
    return False


def list_history_snapshots(st):
    """Return list of dict {'tanggal':str,'sesi':str,'label':str,'date_obj':date}, terbaru dulu.
    Membaca tab bulan ini DAN bulan sebelumnya, supaya histori tidak putus pas awal bulan."""
    import engine
    if not is_configured(st):
        return []
    try:
        client = _get_client(st)
        sh = client.open(st.secrets["gsheet_name"])
    except Exception:
        return []

    cur_tab, prev_tab = _current_and_previous_tab_names()
    records = []
    for tab_name in (cur_tab, prev_tab):
        try:
            ws = sh.worksheet(tab_name)
            records.extend(ws.get_all_records())
        except Exception:
            continue  # tab belum ada, lewati saja

    seen = {}
    for r in records:
        tanggal, sesi = str(r.get("Tanggal", "")).strip(), str(r.get("Sesi", "")).strip()
        if not tanggal or not sesi:
            continue
        seen[(tanggal, sesi)] = True

    items = []
    for (tanggal, sesi) in seen:
        d = engine.parse_id_date(tanggal)
        items.append({"tanggal": tanggal, "sesi": sesi, "date_obj": d,
                       "label": f"{tanggal} ( {sesi} )"})
    items.sort(key=lambda x: (x["date_obj"] or datetime.date.min, x["sesi"]), reverse=True)
    return items


def fetch_history_snapshot(st, tanggal: str, sesi: str):
    """Return dict {RowLabel: {col_key: nilai, ...}} untuk snapshot tanggal+sesi tertentu.
    Tab ditentukan otomatis dari bulan pada tanggal tsb."""
    if not is_configured(st):
        return {}
    try:
        client = _get_client(st)
        sh = client.open(st.secrets["gsheet_name"])
        ws = sh.worksheet(_tab_name_for_tanggal_str(tanggal))
        records = ws.get_all_records()
    except Exception:
        return {}

    result = {}
    for r in records:
        if str(r.get("Tanggal", "")).strip() != tanggal or str(r.get("Sesi", "")).strip() != sesi:
            continue
        label = str(r.get("RowLabel", "")).strip()
        if not label:
            continue
        result[label] = {c: int(r.get(c, 0) or 0) for c in DATA_COLS}
    return result


def repair_header(st) -> tuple[bool, str]:
    """Cek & tambahkan header ke tab bulan ini & bulan sebelumnya kalau belum ada."""
    if not is_configured(st):
        return False, "Google Sheets belum dikonfigurasi."
    try:
        client = _get_client(st)
        sh = client.open(st.secrets["gsheet_name"])
    except Exception as e:
        return False, f"Gagal konek: {e}"

    fixed = []
    for tab_name in _current_and_previous_tab_names():
        try:
            ws = sh.worksheet(tab_name)
        except Exception:
            continue
        if _ensure_header(ws):
            fixed.append(tab_name)

    if fixed:
        return True, f"✅ Header diperbaiki di tab: {', '.join(fixed)}."
    return True, "Header sudah benar di semua tab yang ada, tidak ada yang perlu diperbaiki."


def rename_tab(st, old_name: str, new_name: str) -> tuple[bool, str]:
    """Migrasi manual sekali-pakai: ganti nama 1 tab (mis. 'Histori' -> 'September 2026')."""
    if not is_configured(st):
        return False, "Google Sheets belum dikonfigurasi."
    if not old_name.strip() or not new_name.strip():
        return False, "Nama tab lama dan baru harus diisi."
    try:
        client = _get_client(st)
        sh = client.open(st.secrets["gsheet_name"])
    except Exception as e:
        return False, f"Gagal konek: {e}"

    try:
        ws = sh.worksheet(old_name)
    except Exception:
        return False, f"Tab '{old_name}' tidak ditemukan (mungkin sudah pernah dipindah sebelumnya)."

    try:
        sh.worksheet(new_name)
        return False, f"Tab '{new_name}' sudah ada. Tidak melakukan apa-apa supaya tidak menimpa data yang sudah ada."
    except Exception:
        pass  # bagus, berarti belum ada -> aman untuk rename

    try:
        ws.update_title(new_name)
        return True, f"✅ Tab '{old_name}' berhasil diganti nama jadi '{new_name}'."
    except Exception as e:
        return False, f"Gagal mengganti nama tab: {e}"


def append_history(st, report_df: pd.DataFrame, tanggal_str: str, sesi: str) -> tuple[bool, str]:
    """Tambahkan 1 baris per Type ke tab sesuai bulan tanggal_str (dibuat otomatis kalau belum ada)."""
    if not is_configured(st):
        return False, "Google Sheets belum dikonfigurasi (lihat Secrets)."
    try:
        client = _get_client(st)
        sh = client.open(st.secrets["gsheet_name"])
        tab_name = _tab_name_for_tanggal_str(tanggal_str)
        try:
            ws = sh.worksheet(tab_name)
        except Exception:
            ws = sh.add_worksheet(title=tab_name, rows=1000, cols=20)
        _ensure_header(ws)

        now = datetime.datetime.now().isoformat(timespec="seconds")
        rows_to_add = []
        for _, r in report_df.iterrows():
            row = [now, tanggal_str, sesi, r["RowLabel"], r["RowGroup"]] + \
                  [int(r.get(c, 0) or 0) for c in DATA_COLS] + [int(r["GrandTotal"])]
            rows_to_add.append(row)
        ws.append_rows(rows_to_add, value_input_option="USER_ENTERED")
        return True, f"{len(rows_to_add)} baris tersimpan ke tab '{tab_name}'."
    except Exception as e:
        return False, f"Gagal menyimpan ke Google Sheets: {e}"
