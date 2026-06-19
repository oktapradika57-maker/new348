import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd

# Pengaturan halaman dasbor
st.set_page_config(page_title="Dasbor Analisa Data", layout="wide")

# 1. KONEKSI DATA (Mengambil data dari Google Sheets)
# Catatan: URL di bawah diganti dengan URL Google Sheets Anda nantinya
url = "https://docs.google.com/spreadsheets/d/Halaman_Spreadsheet_Anda_Paling_Baru"
conn = st.connection("gsheets", type=GSheetsConnection)
data = conn.read(spreadsheet=url)

# ---- MENU ANALISA (Kriteria No. 4) ----
st.sidebar.title("Menu Navigasi")
menu = st.sidebar.radio("Pilih Menu:", ["Ringkasan & Analisa", "Detail Data & Pencarian"])

# ---- KONDISI MENU 1: RINGKASAN & ANALISA ----
if menu == "Ringkasan & Analisa":
    st.title("📊 Menu Analisa & Grafik")
    
    # ---- FILTER PENCARIAN / SELEKSI (Kriteria No. 3) ----
    st.subheader("Filter Data")
    # Misalkan di spreadsheet Anda ada kolom bernama 'Kategori' atau 'Status'
    if 'Kategori' in data.columns:
        list_kategori = data['Kategori'].unique()
        pilihan_kategori = st.multiselect("Pilih Kategori:", list_kategori, default=list_kategori)
        # Saring data berdasarkan filter
        data_filtered = data[data['Kategori'].isin(pilihan_kategori)]
    else:
        data_filtered = data

    # ---- HASIL ANALISA (Kriteria No. 2) ----
    st.write("---")
    col1, col2, col3 = st.columns(3)
    
    # Contoh menampilkan metrik ringkasan (sesuaikan dengan nama kolom Anda)
    with col1:
        st.metric(label="Total Data", value=len(data_filtered))
    with col2:
        if 'Biaya' in data_filtered.columns:
            st.metric(label="Total Pengiriman/Biaya", value=f"Rp {data_filtered['Biaya'].sum():,}")
            
    # Contoh Grafik Analisa
    st.subheader("Tren / Grafik Analisa")
    if 'Tanggal' in data_filtered.columns and 'Biaya' in data_filtered.columns:
        st.line_chart(data=data_filtered, x='Tanggal', y='Biaya')
    else:
        st.info("Tambahkan kolom 'Tanggal' dan 'Biaya' di Sheets untuk memunculkan grafik tren otomatis.")

# ---- KONDISI MENU 2: DETAIL DATA & PENCARIAN ----
elif menu == "Detail Data & Pencarian":
    st.title("🔍 Detail Data & Pencarian Konten")
    
    # ---- FILTER PENCARIAN TENTU (Kriteria No. 3) ----
    search_query = st.text_input("Ketik kata kunci untuk mencari data di semua kolom:")
    
    # Logika pencarian teks
    if search_query:
        # Mencari teks di seluruh baris data
        data_searched = data[data.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)]
    else:
        data_searched = data

    # ---- DETAIL DATA (Kriteria No. 1) ----
    st.subheader("Tabel Detail Data Master")
    st.dataframe(data_searched, use_container_width=True)