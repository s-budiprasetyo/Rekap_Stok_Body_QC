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


def append_history(st, report_df: pd.DataFrame, tanggal_str: str, sesi: str) -> tuple[bool, str]:
    """Tambahkan 1 baris per Type ke sheet 'Histori'. Return (sukses, pesan)."""
    if not is_configured(st):
        return False, "Google Sheets belum dikonfigurasi (lihat Secrets)."
    try:
        client = _get_client(st)
        sh = client.open(st.secrets["gsheet_name"])
        try:
            ws = sh.worksheet("Histori")
        except Exception:
            ws = sh.add_worksheet(title="Histori", rows=1000, cols=20)
            header = ["Timestamp", "Tanggal", "Sesi", "RowLabel", "RowGroup"] + \
                     [k for k, *_ in [
                         ("OP100_PERIKSA",), ("OP105_PERBAIKAN",), ("OP107_TDKSET_IJP",),
                         ("OP107_TDKSET_CRJP",), ("OP110_GERINDA",), ("OP115_CEKULANG",),
                         ("OP120_TAP",), ("OP166_FITTINGST",), ("OP166_KIRIM",),
                     ]] + ["GrandTotal"]
            ws.append_row(header)

        now = datetime.datetime.now().isoformat(timespec="seconds")
        data_cols = ["OP100_PERIKSA", "OP105_PERBAIKAN", "OP107_TDKSET_IJP", "OP107_TDKSET_CRJP",
                     "OP110_GERINDA", "OP115_CEKULANG", "OP120_TAP", "OP166_FITTINGST", "OP166_KIRIM"]
        rows_to_add = []
        for _, r in report_df.iterrows():
            row = [now, tanggal_str, sesi, r["RowLabel"], r["RowGroup"]] + \
                  [int(r.get(c, 0) or 0) for c in data_cols] + [int(r["GrandTotal"])]
            rows_to_add.append(row)
        ws.append_rows(rows_to_add, value_input_option="USER_ENTERED")
        return True, f"{len(rows_to_add)} baris tersimpan ke Google Sheets."
    except Exception as e:
        return False, f"Gagal menyimpan ke Google Sheets: {e}"
