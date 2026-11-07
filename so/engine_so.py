"""
Logika murni (tanpa Streamlit) untuk halaman Stok Opname:
- baca file 'Report Stock Opname' (1 file, semua OP jadi satu, 1 baris = 1 pcs)
- baca file perbaikan (Excel) / tabel perbaikan ketikan
- kurangi stok sesuai perbaikan
- susun report_df memakai engine.build_report_table (format laporan SAMA dgn Rekap Harian)
Modul root (engine, render, excel_export) dipakai apa adanya & TIDAK diubah.
"""
import re
import datetime
import pandas as pd

import engine

OPS = engine.ALL_OP_LOCATIONS  # OP100 ... OP166
CORR_COLS = ["Type", "Kode Warna", "Jenis Forming", "Kode Pabrik", "Operation", "Jumlah"]
FORMING_VALUES = {"BC", "HP", "MC"}


def format_id_date(d: datetime.date) -> str:
    return f"{d.day} {engine.BULAN_ID[d.month - 1]} {d.year}"


def _find_header_row(raw: pd.DataFrame, must_have: set, max_scan=15):
    for i in range(min(max_scan, len(raw))):
        vals = {str(v).strip().lower() for v in raw.iloc[i].tolist() if pd.notna(v)}
        if must_have <= vals:
            return i
    return None


# ----------------------------------------------------------------------
# File Stok Opname
# ----------------------------------------------------------------------
def read_so_file(file):
    """Return (grouped_df, info). grouped_df kolom: Lokasi, Type, Warna, KodeNorm, Jenis Forming, Qty
    (format sama dgn engine.combine_raw_files), info = {so_date, n_rows, per_op, unknown_lokasi}."""
    raw = pd.read_excel(file, header=None, dtype=str)
    hdr = _find_header_row(raw, {"type", "lokasi"})
    if hdr is None:
        raise ValueError("Kolom 'Type' dan 'Lokasi' tidak ditemukan. Pastikan ini file 'Report Stock Opname'.")
    df = raw.iloc[hdr + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[hdr].tolist()]

    ren = {}
    for c in df.columns:
        cl = c.lower()
        if cl == "type":
            ren[c] = "Type"
        elif "warna" in cl:
            ren[c] = "Warna"
        elif "forming" in cl:
            ren[c] = "Jenis Forming"
        elif "pabrik" in cl:
            ren[c] = "Kode Pabrik"
        elif cl == "lokasi" or "lokasi" in cl:
            ren[c] = "Lokasi"
        elif "tanggal" in cl:
            ren[c] = "Tanggal"
        elif cl in ("qty", "jumlah"):
            ren[c] = "Qty"
    df = df.rename(columns=ren)
    missing = [c for c in ["Type", "Warna", "Jenis Forming", "Kode Pabrik", "Lokasi"] if c not in df.columns]
    if missing:
        raise ValueError(f"Kolom tidak ditemukan: {missing}. Kolom yang ada: {list(df.columns)}")

    df = df.dropna(subset=["Type", "Lokasi"]).copy()
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce").fillna(1).astype(int) if "Qty" in df.columns else 1
    for c in ["Type", "Warna", "Jenis Forming", "Kode Pabrik", "Lokasi"]:
        df[c] = df[c].astype(str).str.strip()
    df["Lokasi"] = df["Lokasi"].apply(engine.normalize_lokasi)
    df["KodeNorm"] = df["Kode Pabrik"].apply(engine.normalize_kode_pabrik)

    so_date = None
    if "Tanggal" in df.columns:
        dts = pd.to_datetime(df["Tanggal"], errors="coerce").dropna()
        if not dts.empty:
            so_date = dts.dt.date.mode().iloc[0]

    grouped = df.groupby(["Lokasi", "Type", "Warna", "KodeNorm", "Jenis Forming"], as_index=False)["Qty"].sum()
    per_op = df.groupby("Lokasi")["Qty"].sum().to_dict()
    info = {
        "so_date": so_date,
        "n_rows": int(df["Qty"].sum()),
        "per_op": {op: int(per_op.get(op, 0)) for op in OPS},
        "unknown_lokasi": {k: int(v) for k, v in per_op.items() if k not in OPS},
    }
    return grouped, info


# ----------------------------------------------------------------------
# Perbaikan
# ----------------------------------------------------------------------
def empty_corrections() -> pd.DataFrame:
    return pd.DataFrame({
        "Type": pd.Series(dtype="object"), "Kode Warna": pd.Series(dtype="object"),
        "Jenis Forming": pd.Series(dtype="object"), "Kode Pabrik": pd.Series(dtype="object"),
        "Operation": pd.Series(dtype="object"), "Jumlah": pd.Series(dtype="int64"),
    })


def _norm_op(v) -> str:
    s = str(v).strip().upper()
    if re.fullmatch(r"\d{3}", s):
        s = "OP" + s
    m = re.search(r"OP\s*\d+", s)
    return m.group(0).replace(" ", "") if m else s


def clean_corrections(df: pd.DataFrame) -> pd.DataFrame:
    """Bersihkan tabel perbaikan: buang baris kosong/Jumlah<=0, rapikan nilai, gabung baris kembar."""
    if df is None or len(df) == 0:
        return empty_corrections()
    d = df.copy()
    for c in CORR_COLS:
        if c not in d.columns:
            d[c] = None
    d["Jumlah"] = pd.to_numeric(d["Jumlah"], errors="coerce").fillna(0).astype(int)
    d = d.dropna(subset=["Type"])
    d["Type"] = d["Type"].astype(str).str.strip()
    d = d[(d["Type"] != "") & (d["Type"].str.lower() != "nan") & (d["Jumlah"] > 0)].copy()
    if d.empty:
        return empty_corrections()
    d["Kode Warna"] = d["Kode Warna"].fillna("").astype(str).str.strip()
    d["Jenis Forming"] = d["Jenis Forming"].fillna("").astype(str).str.strip().str.upper()
    d["Kode Pabrik"] = d["Kode Pabrik"].apply(lambda v: engine.normalize_kode_pabrik(v) if pd.notna(v) else "")
    d["Operation"] = d["Operation"].fillna("").apply(_norm_op)
    d = d.groupby(["Type", "Kode Warna", "Jenis Forming", "Kode Pabrik", "Operation"], as_index=False)["Jumlah"].sum()
    return d[CORR_COLS].reset_index(drop=True)


def _read_corr_raw(file):
    try:
        file.seek(0)
    except Exception:
        pass
    raw = pd.read_excel(file, header=None, dtype=str)
    hdr = _find_header_row(raw, {"type"})
    if hdr is None:
        raise ValueError("Kolom 'TYPE' tidak ditemukan di file perbaikan.")
    df = raw.iloc[hdr + 1:].copy()
    df.columns = [str(c).strip() for c in raw.iloc[hdr].tolist()]
    return df


def _find_col(df, pred):
    return next((c for c in df.columns if pred(c.lower())), None)


def corr_file_options(file):
    """Isi pilihan dari file perbaikan: (list tanggal posting, list lokasi awal)."""
    df = _read_corr_raw(file)
    c_post = _find_col(df, lambda cl: "posting" in cl)
    c_lok = _find_col(df, lambda cl: "lok" in cl and "awal" in cl)
    dates = []
    if c_post:
        dts = pd.to_datetime(df[c_post], errors="coerce", dayfirst=False).dropna()
        dates = sorted({d.date() for d in dts})
    loks = []
    if c_lok:
        loks = sorted({str(v).strip() for v in df[c_lok].dropna()})
    return dates, loks


def read_correction_excel(file, default_op: str = "OP166", posting_date=None, lokasi_awal=None) -> pd.DataFrame:
    """Baca Excel perbaikan (mis. laporan Packing 'Terima': TYPE, WARNA, PABRIK, JENIS FORMING, ..., TOTAL).
    Kolom Operation dipakai kalau ada di file; kalau tidak, semua baris diberi default_op."""
    df = _read_corr_raw(file)
    c_post = _find_col(df, lambda cl: "posting" in cl)
    if posting_date is not None and c_post:
        pd_ = pd.to_datetime(df[c_post], errors="coerce").dt.date
        df = df[pd_ == posting_date]
    c_lok = _find_col(df, lambda cl: "lok" in cl and "awal" in cl)
    if lokasi_awal is not None and c_lok:
        df = df[df[c_lok].astype(str).str.strip().isin(set(lokasi_awal))]
    cols = {c: c.lower() for c in df.columns}

    def pick(pred):
        return next((c for c, cl in cols.items() if pred(cl)), None)

    c_type = pick(lambda cl: cl == "type")
    c_warna = pick(lambda cl: "warna" in cl)
    c_pabrik = pick(lambda cl: "pabrik" in cl)
    c_total = pick(lambda cl: cl in ("total", "jumlah", "qty") or cl.startswith("total"))
    c_op = pick(lambda cl: cl in ("op", "operation", "operasi", "kode op") or cl.startswith("operation"))
    # kolom forming = salah satu kolom berawalan 'jenis' yang isinya BC/HP/MC
    c_form = None
    for c, cl in cols.items():
        if cl.startswith("jenis") or "forming" in cl:
            vals = df[c].dropna().astype(str).str.strip().str.upper()
            if len(vals) and (vals.isin(FORMING_VALUES).mean() > 0.8):
                c_form = c
                break
    need = {"TYPE": c_type, "WARNA": c_warna, "PABRIK": c_pabrik, "JENIS FORMING": c_form, "TOTAL": c_total}
    miss = [k for k, v in need.items() if v is None]
    if miss:
        raise ValueError(f"Kolom perbaikan tidak ditemukan: {miss}. Kolom yang ada: {list(df.columns)}")

    out = pd.DataFrame({
        "Type": df[c_type], "Kode Warna": df[c_warna], "Jenis Forming": df[c_form],
        "Kode Pabrik": df[c_pabrik],
        "Operation": df[c_op] if c_op else default_op,
        "Jumlah": pd.to_numeric(df[c_total], errors="coerce"),
    })
    out = out.dropna(subset=["Type", "Jumlah"])
    return clean_corrections(out)


def apply_corrections(grouped: pd.DataFrame, corr: pd.DataFrame):
    """Kurangi stok SO sesuai tabel perbaikan. Return (grouped_baru, log_df).
    Kalau perbaikan > stok yang ada -> dikurangi sampai 0 saja (dicatat di log)."""
    out = grouped.copy()
    corr = clean_corrections(corr) if corr is not None else empty_corrections()
    log = []
    for _, c in corr.iterrows():
        mask = (
            (out["Lokasi"] == c["Operation"]) &
            (out["Type"].str.upper() == c["Type"].upper()) &
            (out["Warna"].str.upper() == c["Kode Warna"].upper()) &
            (out["KodeNorm"] == c["Kode Pabrik"]) &
            (out["Jenis Forming"].str.upper() == c["Jenis Forming"].upper())
        )
        want = int(c["Jumlah"])
        avail = int(out.loc[mask, "Qty"].sum())
        if not mask.any():
            took, status = 0, "❌ Tidak ditemukan di data SO (OP/Type/Warna/Forming/Pabrik tidak cocok)"
        else:
            took = min(avail, want)
            idx = out.index[mask][0]
            out.loc[idx, "Qty"] = int(out.loc[idx, "Qty"]) - took
            status = "✅ OK" if took == want else f"⚠️ Stok hanya {avail}, dikurangi sampai 0"
        log.append({**{k: c[k] for k in CORR_COLS}, "Dikurangi": took, "Status": status})
    return out, pd.DataFrame(log)


# ----------------------------------------------------------------------
# Laporan
# ----------------------------------------------------------------------
def manual_from_history(master_df: pd.DataFrame, history_map: dict) -> dict:
    """Nilai baris manual (CRJP, U104ST, S90097, TX60V1) dari 1 snapshot histori. {RowLabel: int}."""
    vals = {}
    for _, m in master_df[master_df["IsManual"] == "TRUE"].iterrows():
        target = m["ManualTargetColumn"]
        v = int((history_map.get(m["RowLabel"], {}) or {}).get(target, 0) or 0) if target else 0
        vals[m["RowLabel"]] = v
    return vals


def build_so_report(grouped: pd.DataFrame, master_df: pd.DataFrame, manual_values: dict):
    """Return (report_df, unmatched_df) -- format sama persis dgn laporan Rekap Harian."""
    return engine.build_report_table(grouped, master_df, manual_values)


# ----------------------------------------------------------------------
# Penjaga pengurang ganda
# ----------------------------------------------------------------------
def file_signature(raw_bytes: bytes, posting_date, lokasi_awal, default_op: str) -> str:
    """Sidik jari 'file + pilihan' pengurang: file yang sama dgn tanggal/lokasi/OP yang sama
    selalu menghasilkan sidik jari yang sama."""
    import hashlib
    h = hashlib.sha1(raw_bytes).hexdigest()[:12]
    lok = ",".join(sorted(lokasi_awal)) if lokasi_awal else "-"
    return f"{h}|{posting_date}|{lok}|{default_op}"


def corr_summary(df: pd.DataFrame):
    """(jumlah baris, total pcs)."""
    d = clean_corrections(df)
    return len(d), int(d["Jumlah"].sum()) if len(d) else 0


def tables_equal(a: pd.DataFrame, b: pd.DataFrame) -> bool:
    a, b = clean_corrections(a), clean_corrections(b)
    key = CORR_COLS[:-1]
    a = a.sort_values(key).reset_index(drop=True)
    b = b.sort_values(key).reset_index(drop=True)
    return a.equals(b)
