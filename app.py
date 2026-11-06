import streamlit as st

import page_input
import page_retrieve
from so import page_so

st.set_page_config(page_title="Rekap Stok Body Harian QC", layout="wide")

pages = [
    st.Page(page_input.render_page, title="Input Harian", icon="📋", url_path="input-harian", default=True),
    st.Page(page_retrieve.render_page, title="Ambil Data", icon="📂", url_path="ambil-data"),
    st.Page(page_so.render_page, title="Stok Opname", icon="📦", url_path="stok-opname"),
]

nav = st.navigation(pages)
nav.run()
