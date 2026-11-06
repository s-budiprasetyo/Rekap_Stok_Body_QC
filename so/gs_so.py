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
        ws.clear()
        ws.update(values=[HEADER] + kept + new_rows, value_input_option="USER_ENTERED")
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
