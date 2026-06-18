import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

# Pustaka untuk Export PPT & PDF
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# =========================================================================
# ⚙️ KONFIGURASI OPERASIONAL & URL RESOURCE
# =========================================================================
st.set_page_config(
    page_title="G348T Task Force Telkomsel Enom",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ID Spreadsheet "348 Tsel"
SPREADSHEET_ID = "1FGKOzWoUrbf3PXN_ahgG1t-83JZT4H4sioQepePbBxM"
csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv"

def format_drive_image_url(url_or_id):
    val = str(url_or_id).strip()
    if val == "" or val.lower() == "nan":
        return None
    match = re.search(r'([a-zA-Z0-9_-]{33})', val)
    if match:
        doc_id = match.group(1)
    else:
        doc_id = val
    return f"https://lh3.googleusercontent.com/d/{doc_id}"

st.title("⚡ Infrastructure Power & Asset Dashboard")
st.caption("Live Synchronization with Google Sheets Resource (348 Tsel) & Voltage RST Pattern Analytics")
st.markdown("---")

# =========================================================================
# 💾 LOAD DATA DENGAN ANTI-CRASH PROTECTION
# =========================================================================
@st.cache_data(ttl=5)
def load_financial_data(url):
    try:
        df = pd.read_csv(url)
        return df
    except Exception as e:
        return str(e)

with st.spinner("🔄 Sedang memuat resource data terbaru..."):
    df_raw = load_financial_data(csv_url)

if isinstance(df_raw, str):
    st.error("⚠️ Gagal memuat data dari link Google Sheets yang diberikan.")
    st.markdown(f"**Detail Error:** `{df_raw}`")
    st.stop()

elif df_raw is None or df_raw.empty:
    st.warning("⚠️ Data di dalam spreadsheet kosong.")
    st.stop()

else:
    df = df_raw.dropna(how='all', axis=1).copy()
    df = df.dropna(how='all', axis=0)

    # =========================================================================
    # 🔍 DETEKSI KOLOM SECARA DINAMIS
    # =========================================================================
    potential_names = ['Nama Gardu', 'Component', 'Site ID', 'Nama', 'Site Name', 'Tower ID', 'SITE_ID']
    KOLOM_UTAMA = next((c for c in df.columns if c in potential_names), df.columns[0])

    potential_power = ['Daya (VA)', 'Kapasitas', 'Daya', 'Power (VA)', 'Daya PLN']
    KOLOM_DAYA = next((c for c in df.columns if c in potential_power), None)

    potential_loc = ['Lokasi', 'Location', 'Alamat', 'Wilayah', 'Witel']
    KOLOM_LOKASI = next((c for c in df.columns if c in potential_loc), None)

    potential_img = ['ID_Foto_Drive', 'Link Foto', 'Gambar', 'Foto', 'Drive Image', 'Link Gambar', 'FOTO']
    KOLOM_FOTO = next((c for c in df.columns if c in potential_img), None)

    # 🔹 Deteksi Otomatis Kolom Tegangan RST (Case Insensitive)
    KOLOM_VR = next((c for c in df.columns if any(p in c.upper() for p in ['VOLTAGE R', 'TEGANGAN R', 'VOLT R', 'V_R'])), None)
    KOLOM_VS = next((c for c in df.columns if any(p in c.upper() for p in ['VOLTAGE S', 'TEGANGAN S', 'VOLT S', 'V_S'])), None)
    KOLOM_VT = next((c for c in df.columns if any(p in c.upper() for p in ['VOLTAGE T', 'TEGANGAN T', 'VOLT T', 'V_T'])), None)

    # --- SIDEBAR PANEL FILTER ---
    st.sidebar.header("⚙️ Panel Filter Resource")
    unique_items = ["Semua Komponen"] + sorted(list(df[KOLOM_UTAMA].astype(str).unique()))
    selected_item = st.sidebar.selectbox(f"Pilih {KOLOM_UTAMA}:", unique_items)

    if st.sidebar.button("🔄 Refresh & Sinkronkan Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # --- PROSES FILTER DATA ---
    df_filtered = df.copy()
    if selected_item != "Semua Komponen":
        df_filtered = df_filtered
