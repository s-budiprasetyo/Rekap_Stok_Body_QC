"""Unduhan otomatis + pesan singkat setelah dialog ditutup.

Alur: tombol di dialog -> queue_download() -> st.rerun() (dialog hilang) ->
run_pending() di awal halaman memasang iframe tak terlihat yang meng-klik link unduhan
SEKALI, lalu item dibuang dari antrean (tidak terunduh ulang pada interaksi berikutnya)."""
import base64
import streamlit as st
import streamlit.components.v1 as components

_KEY = "_pending_download"
_FLASH = "_flash_msgs"


def queue_download(data: bytes, filename: str, mime: str):
    st.session_state[_KEY] = (data, filename, mime)


def flash(level: str, msg: str):
    st.session_state.setdefault(_FLASH, []).append((level, msg))


def run_pending():
    for level, msg in st.session_state.pop(_FLASH, []):
        if level == "success":
            st.toast(msg, icon="✅")
        elif level == "error":
            st.error(msg)
        elif level == "warning":
            st.warning(msg)
        else:
            st.toast(msg)
    item = st.session_state.pop(_KEY, None)
    if item:
        data, filename, mime = item
        b64 = base64.b64encode(data).decode()
        components.html(
            f'<a id="dl" href="data:{mime};base64,{b64}" download="{filename}" style="display:none"></a>'
            f'<script>document.getElementById("dl").click();</script>', height=0, width=0)
