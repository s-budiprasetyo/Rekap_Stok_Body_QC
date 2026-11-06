"""
Pemilih tanggal gaya kalender Windows (latar abu terang) + tombol Apply di kanan bawah.
Dipakai di SEMUA pilihan tanggal aplikasi (Input Harian, Stok Opname, Ambil Data).

    d = ui_calendar.date_picker_apply("kunci_unik", enabled_dates=[date, ...] atau None, default=date)

 * titik kecil di bawah angka  = tanggal itu punya data histori (kalau enabled_dates diberikan)
 * kotak biru                  = tanggal yang dipilih
 * kotak garis tipis           = hari ini (WIB)
 * angka abu-abu               = bukan bulan yang sedang tampil / tidak punya data
Tanggal baru HANYA berlaku setelah tombol Apply diklik.
"""
import calendar
import datetime
from zoneinfo import ZoneInfo

import streamlit as st

import engine

_HARI = ["Min", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"]
_WIB = ZoneInfo("Asia/Jakarta")


def today_wib() -> datetime.date:
    return datetime.datetime.now(_WIB).date()


def fmt_id(d: datetime.date) -> str:
    return f"{d.day} {engine.BULAN_ID[d.month - 1]} {d.year}"


def _css(key: str) -> str:
    w = f".st-key-calwrap_{key}"
    dlg = f'[data-testid="stDialog"]:has({w})'
    return f"""
<style>
{dlg} div:has(> section[role="dialog"]), {dlg} section[role="dialog"] {{ background:#f2f2f2 !important; }}
{dlg} h1, {dlg} h2, {dlg} h3, {dlg} p, {dlg} span, {dlg} label, {dlg} svg {{ color:#1a1a1a; }}
{w} [data-testid="stHorizontalBlock"] {{ flex-wrap:nowrap !important; gap:0.1rem !important; }}
{w} [data-testid="stColumn"] {{ min-width:0 !important; }}
{w} button {{ min-height:2.5rem; padding:0; border-radius:2px; font-size:1.05rem; position:relative; }}
{w} [data-testid="stBaseButton-tertiary"] {{ background:transparent; border:2px solid transparent; }}
{w} [data-testid="stBaseButton-tertiary"]:not(:disabled) * {{ color:#1a1a1a !important; }}
{w} [data-testid="stBaseButton-tertiary"]:hover:not(:disabled) {{ background:#e1e1e1; }}
{w} [data-testid="stBaseButton-tertiary"]:disabled {{ opacity:1; background:transparent; }}
{w} [data-testid="stBaseButton-tertiary"]:disabled * {{ color:#a3a3a3 !important; }}
/* titik = ada data */
{w} [class*="st-key-{key}__dD"] button::after {{
  content:""; position:absolute; bottom:3px; left:50%; width:4px; height:4px;
  margin-left:-2px; border-radius:50%; background:#1a1a1a; }}
{w} [class*="st-key-{key}__dD"] button:disabled::after {{ background:#a3a3a3; }}
/* hari ini = kotak garis tipis */
{w} [class*="st-key-{key}__dT"] [data-testid="stBaseButton-tertiary"],
{w} [class*="st-key-{key}__dDT"] [data-testid="stBaseButton-tertiary"] {{ border:2px solid #c9c9c9; }}
/* terpilih = kotak biru */
{w} [data-testid="stBaseButton-primary"] {{
  background:#0078d7; border:2px solid #0b3d6e; box-shadow:inset 0 0 0 2px #ffffff; }}
{w} [data-testid="stBaseButton-primary"] * {{ color:#ffffff !important; }}
{w} [class*="st-key-{key}__dD"] [data-testid="stBaseButton-primary"]::after {{ background:#ffffff; }}
{w} [data-testid="stIconMaterial"] {{ font-size:1.7rem !important; }}
{w} .cal-head {{ text-align:center; font-weight:600; color:#1a1a1a; font-size:0.85rem; padding:0.35rem 0; }}
{w} .cal-month {{ font-size:1.1rem; font-weight:400; color:#1a1a1a; padding-top:0.45rem; white-space:nowrap; }}
{w} [data-testid="stBaseButton-secondary"] {{ background:#0078d7; color:#fff; border:none; padding:0 1.2rem; }}
{w} [data-testid="stBaseButton-secondary"] * {{ color:#fff !important; }}
</style>
"""


def _select(key: str, d: datetime.date):
    st.session_state[f"{key}__pending"] = d


def _shift_month(key: str, delta: int):
    y, m = st.session_state[f"{key}__view"]
    m += delta
    if m < 1:
        y, m = y - 1, 12
    elif m > 12:
        y, m = y + 1, 1
    st.session_state[f"{key}__view"] = (y, m)


@st.dialog("Pilih tanggal", width="small")
def _calendar_dialog(key: str, enabled_dates):
    enabled = set(enabled_dates) if enabled_dates is not None else None
    st.markdown(_css(key), unsafe_allow_html=True)
    y, m = st.session_state[f"{key}__view"]
    pending = st.session_state.get(f"{key}__pending")
    today = today_wib()

    with st.container(key=f"calwrap_{key}"):
        c0, c1, c2 = st.columns([8, 1, 1])
        c0.markdown(f'<div class="cal-month">{engine.BULAN_ID[m - 1]} {y}</div>', unsafe_allow_html=True)
        c1.button("", key=f"{key}__prev", type="tertiary", icon=":material/keyboard_arrow_up:",
                  on_click=_shift_month, args=(key, -1))
        c2.button("", key=f"{key}__next", type="tertiary", icon=":material/keyboard_arrow_down:",
                  on_click=_shift_month, args=(key, 1))

        head = st.columns(7)
        for i, h in enumerate(_HARI):
            head[i].markdown(f'<div class="cal-head">{h}</div>', unsafe_allow_html=True)

        for week in calendar.Calendar(firstweekday=6).monthdatescalendar(y, m):
            cols = st.columns(7)
            for i, d in enumerate(week):
                has = enabled is not None and d in enabled
                ok = enabled is None or has
                flag = ("D" if has else "") + ("T" if d == today else "")
                cols[i].button(
                    str(d.day), key=f"{key}__d{flag}_{d.isoformat()}",
                    type="primary" if pending == d else "tertiary",
                    disabled=not ok, on_click=_select, args=(key, d), use_container_width=True,
                )

        _, right = st.columns([3, 2])
        with right:
            if st.button("Apply", key=f"{key}__apply", use_container_width=True, disabled=pending is None):
                st.session_state[f"{key}__applied"] = pending
                st.rerun()


def date_picker_apply(key: str, enabled_dates=None, default=None, label: str = "Tanggal"):
    """Tombol tanggal -> kalender + Apply. Return datetime.date (atau None kalau belum ada default)."""
    enabled = set(enabled_dates) if enabled_dates is not None else None
    applied = st.session_state.get(f"{key}__applied")
    if applied is None or (enabled is not None and applied not in enabled):
        applied = default
        st.session_state[f"{key}__applied"] = applied

    st.caption(label)
    text = f"📅  {fmt_id(applied)}" if applied else "📅  Pilih tanggal"
    if st.button(text, key=f"{key}__open"):
        base = applied or today_wib()
        st.session_state[f"{key}__view"] = (base.year, base.month)
        st.session_state[f"{key}__pending"] = applied
        _calendar_dialog(key, enabled_dates)
    return applied
