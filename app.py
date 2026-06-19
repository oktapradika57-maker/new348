import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd

# Pengaturan halaman dasbor
st.set_page_config(page_title="Dasbor Analisa Data", layout="wide")

# URL Spreadsheet resmi milik Anda (sudah disesuaikan ke format sharing publik)
URL_SHEETS = "https://docs.google.com/spreadsheets/d/1FGKOzWoUrbf3PXN_ahgG1t-83JZT4H4sioQepePbBxM/edit?usp=sharing"

st.title("📊 Dasbor Analisa & Monitoring")

try:
    # 1. KONEKSI DATA (Mengambil data dari Google Sheets Anda)
    conn = st.connection("gsheets", type=GSheetsConnection)
    # ttl=60 artinya Streamlit akan mengecek perubahan data di Sheets setiap 1 menit
    data = conn.read(spreadsheet=URL_SHEETS, ttl=60)
    
    # Membersihkan baris yang sepenuhnya kosong di spreadsheet
    data = data.dropna(how='all')

    # ---- KRITERIA 4: MENU ANALISA (Navigasi Sidebar) ----
    st.sidebar.title("📌 Navigasi Dasbor")
    menu = st.sidebar.radio("Pilih Menu:", ["📈 Hasil Analisa", "🔍 Detail Data & Pencarian"])

    # =========================================================
    # KONDISI MENU 1: HASIL ANALISA
    # =========================================================
    if menu == "📈 Hasil Analisa":
        st.header("Hasil Analisa & Ringkasan Metrik")
        
        # ---- KRITERIA 3: FILTER PENCARIAN & SELEKSI DINAMIS ----
        st.subheader("🛠️ Filter Data Berdasarkan Kolom")
        kolom_pilihan = data.columns.tolist()
        kolom_filter = st.selectbox("Pilih kolom yang ingin difilter:", kolom_pilihan)
        
        if kolom_filter:
            opsi_filter = data[kolom_filter].dropna().unique().tolist()
            pilihan_nilai = st.multiselect(f"Pilih nilai untuk [{kolom_filter}]:", opsi_filter, default=opsi_filter)
            # Menyaring data berdasarkan pilihan filter
            data_filtered = data[data[kolom_filter].isin(pilihan_nilai)]
        else:
            data_filtered = data

        st.write("---")
        
        # ---- KRITERIA 2: HASIL ANALISA (Metrik Angka) ----
        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="Total Volume / Baris Data", value=f"{len(data_filtered)} item")
            
        # Sistem mendeteksi otomatis jika ada kolom berisi angka (misal: Biaya, Jumlah, Nominal)
        kolom_numerik = data_filtered.select_dtypes(include=['number']).columns.tolist()
        if kolom_numerik:
            with col2:
                kolom_sum = st.selectbox("Pilih kolom angka untuk dihitung totalnya:", kolom_numerik)
                total_nilai = data_filtered[kolom_sum].sum()
                st.metric(label=f"Total Akumulasi {kolom_sum}", value=f"{total_nilai:,.0f}")
        else:
            with col2:
                st.info("Info: Tidak ada kolom bertipe angka terdeteksi untuk kalkulasi.")

        # ---- GRAFIK ANALISA DINAMIS ----
        st.subheader("📊 Visualisasi Grafik")
        if len(kolom_pilihan) >= 2:
            col_x = st.selectbox("Pilih Sumbu X (Kategori / Tanggal):", kolom_pilihan)
            if kolom_numerik:
                col_y = st.selectbox("Pilih Sumbu Y (Nilai Parameter Angka):", kolom_numerik)
                st.bar_chart(data=data_filtered, x=col_x, y=col_y)
            else:
                st.warning("Tambahkan kolom berisi data angka di Google Sheets untuk memunculkan grafik.")
        else:
            st.info("Data di spreadsheet terlalu sedikit untuk dianalisis dalam bentuk grafik.")

    # =========================================================
    # KONDISI MENU 2: DETAIL DATA & PENCARIAN
    # =========================================================
    elif menu == "🔍 Detail Data & Pencarian":
        st.header("Detail Data Master & Pencarian")
        
        # ---- KRITERIA 3: FILTER PENCARIAN KATA KUNCI GLOBAL ----
        search_query = st.text_input("🔍 Ketik kata kunci bebas untuk mencari data di semua kolom:")
        
        if search_query:
            # Mencari teks di seluruh baris tanpa peduli huruf besar/kecil
            data_searched = data[data.astype(str).apply(lambda x: x.str.contains(search_query, case=False)).any(axis=1)]
            st.success(f"Ditemukan {len(data_searched)} baris data yang cocok.")
        else:
            data_searched = data

        # ---- KRITERIA 1: DETAIL DATA (Tabel Lengkap) ----
        st.subheader("📋 Tabel Data Lengkap")
        st.dataframe(data_searched, use_container_width=True)

except Exception as e:
    st.error("🔒 Koneksi Gagal atau Akses Ditolak")
    st.warning("Silakan periksa kembali apakah setelan Google Sheets Anda sudah diubah menjadi 'Siapa saja yang memiliki link' (Anyone with the link).")
    st.info("Pesan Error Teknis:")
    st.code(str(e))
