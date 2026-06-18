import streamlit as st
import pandas as pd
import numpy as np
import re

# Konfigurasi Halaman
st.set_page_config(
    page_title="Genset & Site Analytics Dashboard",
    page_icon="📸",
    layout="wide"
)

# ID Spreadsheet Baru Anda
SPREADSHEET_ID = "1FGKOzWoUrbf3PXN_ahgG1t-83JZT4H4sioQepePbBxM"
# Menggunakan alternatif endpoint Google Viz API yang jauh lebih stabil untuk Pandas
csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv"

# Fungsi Konversi Link Google Drive ke Direct Image URL
def convert_drive_url(url):
    if pd.isna(url) or not isinstance(url, str):
        return None
    match = re.search(r'(?:\/file\/d\/|id=)([\w-]+)', url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=view&id={file_id}"
    return None

# Memuat Data
@st.cache_data(ttl=30)
def load_data(url):
    try:
        df = pd.read_csv(url)
        return df
    except Exception as e:
        st.error(f"Gagal memuat data dari Google Sheets. Pastikan link sudah 'Anyone with link can view'. Error: {e}")
        return None

df_raw = load_data(csv_url)

if df_raw is not None and not df_raw.empty:
    # Membersihkan kolom kosong akibat pembacaan export gviz
    df_raw = df_raw.dropna(how='all', axis=1)
    cols = df_raw.columns
    
    # Deteksi Otomatis Kolom Utama
    site_id_col = next((c for c in cols if 'site' in c.lower() and 'id' in c.lower()), None)
    start_col = next((c for c in cols if 'start' in c.lower() and ('backup' in c.lower() or 'waktu' in c.lower())), None)
    stop_col = next((c for c in cols if 'stop' in c.lower() and ('backup' in c.lower() or 'waktu' in c.lower())), None)
    rh_start_col = next((c for c in cols if 'rh' in c.lower() and 'start' in c.lower()), None)
    rh_stop_col = next((c for c in cols if 'rh' in c.lower() and 'stop' in c.lower()), None)
    
    evidence_cols = [c for c in cols if any(keyword in c.lower() for keyword in ['gambar', 'foto', 'link', 'dokumentasi', 'evidence'])]

    if not site_id_col:
        st.warning("⚠️ Kolom 'Site ID' tidak ditemukan secara otomatis. Menampilkan semua data.")
        site_id_col = cols[0]  # Gunakan kolom pertama sebagai fallback

    # Pre-processing dasar
    df = df_raw.copy()
    
    if start_col and stop_col:
        df[start_col] = pd.to_datetime(df[start_col], errors='coerce', format='mixed')
        df[stop_col] = pd.to_datetime(df[stop_col], errors='coerce', format='mixed')
        df['Durasi Nyata (Jam)'] = ((df[stop_col] - df[start_col]).dt.total_seconds() / 3600).round(2).fillna(0)
    else:
        df['Durasi Nyata (Jam)'] = 0

    if rh_start_col and rh_stop_col:
        df[rh_start_col] = pd.to_numeric(df[rh_start_col], errors='coerce').fillna(0)
        df[rh_stop_col] = pd.to_numeric(df[rh_stop_col], errors='coerce').fillna(0)
        df['Durasi RH (Jam)'] = np.where(df[rh_stop_col] >= df[rh_start_col], df[rh_stop_col] - df[rh_start_col], 0).round(2)
    else:
        df['Durasi RH (Jam)'] = 0

    # Pastikan data string aman
    df[site_id_col] = df[site_id_col].astype(str)

    # --- PANEL UTAMA & FILTER ---
    st.title("📸 Genset Analytics & Documentation Dashboard")
    
    unique_sites = sorted(df[site_id_col].dropna().unique().tolist())
    selected_site = st.selectbox("📍 Pilih Site ID untuk Dianalisa:", ["Semua Site"] + unique_sites)
    
    if selected_site == "Semua Site":
        df_selected = df
    else:
        df_selected = df[df[site_id_col] == selected_site]

    # --- KPI GRID ---
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Records", f"{len(df_selected)} Baris")
    with col2:
        st.metric("Total Durasi Nyata", f"{df_selected['Durasi Nyata (Jam)'].sum():.2f} Jam")
    with col3:
        st.metric("Total Durasi Running Hours", f"{df_selected['Durasi RH (Jam)'].sum():.2f} Jam")

    st.markdown("---")

    # --- GRAFIK ANALISA DURASI ---
    st.subheader("📊 Grafik Analisa Durasi Backup")
    if not df_selected.empty:
        chart_df = pd.DataFrame({
            'Durasi Nyata (Jam)': df_selected['Durasi Nyata (Jam)'].values,
            'Durasi RH (Jam)': df_selected['Durasi RH (Jam)'].values
        }, index=df_selected[site_id_col].values if selected_site == "Semua Site" else range(1, len(df_selected) + 1))
        st.bar_chart(chart_df, color=["#10b981", "#3b82f6"])

    st.markdown("---")

    # --- GALERI GAMBAR BERURUTAN ---
    st.subheader("🖼️ Dokumentasi Gambar Berurutan (Google Drive)")
    
    if len(evidence_cols) == 0:
        st.info("Tidak terdeteksi adanya kolom gambar/dokumentasi khusus di spreadsheet Anda.")
    else:
        for idx, row in df_selected.iterrows():
            st.markdown(f"#### 📅 Log Data ke-{idx+1} | Site: **{row[site_id_col]}**")
            img_slots = st.columns(len(evidence_cols))
            
            for i, col_name in enumerate(evidence_cols):
                raw_url = row[col_name]
                direct_img_url = convert_drive_url(raw_url)
                
                with img_slots[i]:
                    if direct_img_url:
                        st.image(direct_img_url, caption=f"{col_name}", use_container_width=True)
                    else:
                        st.caption(f"❌ {col_name}: Tidak ada gambar")
            st.markdown("---")

    # --- DATA TABLE VIEW ---
    st.subheader("📋 Data Mentah Terfilter")
    st.dataframe(df_selected, use_container_width=True)
else:
    st.info("Belum ada data atau pembacaan spreadsheet kosong.")
