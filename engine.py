"""
Engine v2 — merekap data mentah SAP jadi laporan STOK BODY HARIAN QC,
berbasis master_type.csv (baris & kunci pencocokan TETAP, ditentukan user).
"""
import re
import datetime
import pandas as pd

REPORT_COLUMNS = [
    ("OP100_PERIKSA",     "OP100", "PERIKSA"),
    ("OP105_PERBAIKAN",   "OP105", "PERBAIKAN"),
    ("OP107_TDKSET_IJP",  "OP107", "TDK SET\niJP"),
    ("OP107_TDKSET_CRJP", "OP107", "TDK SET\nCRJP"),
    ("OP110_GERINDA",     "OP110", "GERINDA"),
    ("OP115_CEKULANG",    "OP115", "CEK\nULANG"),
    ("OP120_TAP",         "OP120", "TAP"),
    ("OP166_FITTINGST",   "OP166", "FITTING\n& ST"),
    ("OP166_KIRIM",       "OP166", "KIRIM"),
]
REPORT_COLUMN_KEYS = [c[0] for c in REPORT_COLUMNS]
GROUP_ORDER = ["TYPE C", "TYPE L", "TYPE S", "TYPE U"]

LOKASI_TO_COLKEY = {
    "OP100": "OP100_PERIKSA",
    "OP105": "OP105_PERBAIKAN",
    "OP107": "OP107_TDKSET_IJP",   # dari file OP107 SAP selalu masuk kolom iJP; CRJP manual
    "OP110": "OP110_GERINDA",
    "OP115": "OP115_CEKULANG",
    "OP120": "OP120_TAP",
    "OP166": "OP166_FITTINGST",
}

RAW_COLUMNS = ["No.", "Nama Lokasi", "Type", "Warna", "Kode Pabrik", "Jenis Forming", "Jenis Proses", "Qty"]

BULAN_ID = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"]


def today_id_date() -> str:
    d = datetime.date.today()
    return f"{d.day} {BULAN_ID[d.month-1]} {d.year}"


def normalize_lokasi(nama_lokasi: str) -> str:
    s = str(nama_lokasi).upper().replace("QC", "").strip()
    m = re.search(r"OP\s*\d+", s)
    return m.group(0).replace(" ", "") if m else s


def normalize_kode_pabrik(v) -> str:
    """Samakan '01'/'1'/1/1.0 semua jadi '1'."""
    s = str(v).strip()
    s = re.sub(r"\D", "", s)
    return str(int(s)) if s else ""


def read_raw_file(file) -> pd.DataFrame:
    df = pd.read_excel(file, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    rename_map = {}
    for c in df.columns:
        cl = c.lower().replace(".", "").strip()
        if cl == "no":
            rename_map[c] = "No."
        elif "lokasi" in cl:
            rename_map[c] = "Nama Lokasi"
        elif cl == "type":
            rename_map[c] = "Type"
        elif "warna" in cl:
            rename_map[c] = "Warna"
        elif "pabrik" in cl:
            rename_map[c] = "Kode Pabrik"
        elif "forming" in cl:
            rename_map[c] = "Jenis Forming"
        elif "proses" in cl:
            rename_map[c] = "Jenis Proses"
        elif cl == "qty":
            rename_map[c] = "Qty"
    df = df.rename(columns=rename_map)
    missing = [c for c in RAW_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Kolom tidak ditemukan: {missing}. Kolom yang ada: {list(df.columns)}")
    df = df[RAW_COLUMNS].copy()
    df = df.dropna(subset=["Type", "Nama Lokasi"])
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(0).astype(int)
    for c in ["Type", "Warna", "Kode Pabrik", "Jenis Forming", "Jenis Proses", "Nama Lokasi"]:
        df[c] = df[c].astype(str).str.strip()
    df["Lokasi"] = df["Nama Lokasi"].apply(normalize_lokasi)
    df["KodeNorm"] = df["Kode Pabrik"].apply(normalize_kode_pabrik)
    return df


def combine_raw_files(files) -> pd.DataFrame:
    frames = [read_raw_file(f) for f in files]
    if not frames:
        return pd.DataFrame()
    combined = pd.concat(frames, ignore_index=True)
    grouped = combined.groupby(
        ["Lokasi", "Type", "Warna", "KodeNorm", "Jenis Forming"], as_index=False
    )["Qty"].sum()
    return grouped


def load_master(file) -> pd.DataFrame:
    df = pd.read_csv(file, dtype=str).fillna("")
    df["KodeNorm"] = df["KodePabrik"].apply(lambda v: normalize_kode_pabrik(v) if v else "")
    return df


def build_report_table(grouped_raw: pd.DataFrame, master_df: pd.DataFrame, manual_values: dict):
    """
    manual_values: dict {RowLabel: int}. Nilai otomatis dimasukkan ke ManualTargetColumn
    baris tsb (kalau ada), kalau tidak ada -> langsung nambah ke Grand Total saja.
    """
    auto_master = master_df[master_df["IsManual"] != "TRUE"].copy()
    data = {row["RowLabel"]: {k: 0 for k in REPORT_COLUMN_KEYS} for _, row in master_df.iterrows()}

    unmatched = []
    for _, r in grouped_raw.iterrows():
        lokasi = r["Lokasi"]
        type_col = f"Type_{lokasi}"
        if type_col not in auto_master.columns:
            unmatched.append(r)
            continue
        cand = auto_master[
            (auto_master[type_col] == r["Type"]) &
            (auto_master["Warna"] == r["Warna"]) &
            (auto_master["KodeNorm"] == r["KodeNorm"]) &
            (auto_master["JenisForming"].str.upper() == r["Jenis Forming"].upper())
        ]
        if cand.empty:
            unmatched.append(r)
            continue
        if len(cand) > 1:
            if lokasi == "OP107":
                pick = cand[cand["OnlyFromOP107"] == "TRUE"]
            else:
                pick = cand[cand["OnlyFromOP107"] != "TRUE"]
            cand = pick if not pick.empty else cand.iloc[[0]]
        row_label = cand.iloc[0]["RowLabel"]
        col_key = LOKASI_TO_COLKEY.get(lokasi)
        if col_key:
            data[row_label][col_key] += int(r["Qty"])

    for row_label, val in (manual_values or {}).items():
        if row_label not in data or not val:
            continue
        master_row = master_df[master_df["RowLabel"] == row_label]
        target_col = master_row.iloc[0]["ManualTargetColumn"] if not master_row.empty else ""
        v = int(val)
        if target_col:
            data[row_label][target_col] = data[row_label].get(target_col, 0) + v
        else:
            data[row_label]["_DIRECT_TOTAL_"] = data[row_label].get("_DIRECT_TOTAL_", 0) + v

    records = []
    for _, m in master_df.iterrows():
        label = m["RowLabel"]
        vals = data[label]
        grand = sum(vals.get(k, 0) for k in REPORT_COLUMN_KEYS) + vals.get("_DIRECT_TOTAL_", 0)
        rec = {"RowLabel": label, "RowGroup": m["RowGroup"], "ISI": m["ISI"],
               "IsAksesoris": m["IsAksesoris"] == "TRUE", "GrandTotal": grand}
        for k in REPORT_COLUMN_KEYS:
            rec[k] = vals.get(k, 0)
        records.append(rec)

    report_df = pd.DataFrame(records)
    order_map = {g: i for i, g in enumerate(GROUP_ORDER)}
    report_df["_order"] = report_df["RowGroup"].map(order_map).fillna(99)
    report_df = report_df.sort_values(["_order"], kind="stable").drop(columns="_order").reset_index(drop=True)

    unmatched_df = pd.DataFrame(unmatched) if unmatched else pd.DataFrame()
    return report_df, unmatched_df


def compute_totals(report_df: pd.DataFrame) -> dict:
    aksesoris = int(report_df[report_df["IsAksesoris"]]["GrandTotal"].sum())
    pcs1 = int(report_df[~report_df["IsAksesoris"]]["GrandTotal"].sum())
    return {"pcs1": pcs1, "aksesoris": aksesoris, "body": pcs1 + aksesoris}
