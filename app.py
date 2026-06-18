import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Pustaka untuk Export PPT & PDF
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# =========================================================================
# ⚙️ KONFIGURASI OPERASIONAL & URL RESOURCE BARU
# =========================================================================
st.set_page_config(
    page_title="PLN & Infrastructure Asset Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ID Spreadsheet Baru yang Anda Berikan (348 Tsel)
SPREADSHEET_ID = "1FGKOzWoUrbf3PXN_ahgG1t-83JZT4H4sioQepePbBxM"
# Format ekspor resmi yang stabil untuk membaca tab utama secara langsung
csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv"

# Fungsi penarik gambar langsung (Direct Stream) Google Drive
def format_drive_image_url(url_or_id):
    if "id=" in str(url_or_id):
        doc_id = str(url_or_id).split("id=")[1].split("&")[0]
    elif "file/d/" in str(url_or_id):
        doc_id = str(url_or_id).split("file/d/")[1].split("/")[0]
    else:
        doc_id = str(url_or_id)
    return f"https://docs.google.com/uc?export=view&id={doc_id}"

st.title("⚡ Infrastructure Power & Asset Dashboard")
st.caption("Live Synchronization with Google Sheets Resource (348 Tsel) & Google Drive")
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

# Penanganan Jika Koneksi Bermasalah / Spreadsheet Kosong
if isinstance(df_raw, str):
    st.error("⚠️ Gagal memuat data dari link Google Sheets yang diberikan.")
    st.markdown(f"**Detail Error:** `{df_raw}`")
    st.info("💡 Pastikan pengaturan berbagi di Google Sheets tersebut sudah diset ke **'Anyone with the link can view'**.")
    st.stop()

elif df_raw is None or df_raw.empty:
    st.warning("⚠️ Koneksi berhasil, namun tidak ada baris data yang terbaca di dalam spreadsheet.")
    st.stop()

else:
    # Bersihkan kolom kosong gaib akibat format tabel di spreadsheet
    df = df_raw.dropna(how='all', axis=1).copy()
    df = df.dropna(how='all', axis=0)

    # =========================================================================
    # 🔍 DETEKSI KOLOM SECARA DINAMIS (Menyesuaikan Otomatis dengan Isi Sheet)
    # =========================================================================
    # Deteksi Kolom Komponen Utama / Nama Gardu / Site ID
    potential_names = ['Nama Gardu', 'Component', 'Site ID', 'Nama', 'Site Name']
    KOLOM_UTAMA = next((c for c in df.columns if c in potential_names), df.columns[0])

    # Deteksi Kolom Daya / Kapasitas
    potential_power = ['Daya (VA)', 'Kapasitas', 'Daya', 'Power (VA)', 'Daya PLN']
    KOLOM_DAYA = next((c for c in df.columns if c in potential_power), None)

    # Deteksi Kolom Lokasi
    potential_loc = ['Lokasi', 'Location', 'Alamat', 'Wilayah']
    KOLOM_LOKASI = next((c for c in df.columns if c in potential_loc), None)

    # Deteksi Kolom Link Gambar Google Drive
    potential_img = ['ID_Foto_Drive', 'Link Foto', 'Gambar', 'Foto', 'Drive Image']
    KOLOM_FOTO = next((c for c in df.columns if c in potential_img), None)

    # --- SIDEBAR PANEL FILTER ---
    st.sidebar.header("⚙️ Panel Filter Resource")
    
    # Filter Berdasarkan Komponen Utama yang Ditemukan
    unique_items = ["Semua Komponen"] + sorted(list(df[KOLOM_UTAMA].astype(str).unique()))
    selected_item = st.sidebar.selectbox(f"Pilih {KOLOM_UTAMA}:", unique_items)

    if st.sidebar.button("🔄 Refresh & Sinkronkan Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # --- PROSES FILTER DATA ---
    df_filtered = df.copy()
    if selected_item != "Semua Komponen":
        df_filtered = df_filtered[df_filtered[KOLOM_UTAMA].astype(str) == selected_item]

    # --- TAMPILAN UTAMA DASHBOARD ---
    st.subheader("📋 Detail Rincian Data")
    st.dataframe(df_filtered, use_container_width=True, hide_index=True)

    # --- BAGIAN VISUALISASI GAMBAR GOOGLE DRIVE ---
    if KOLOM_FOTO and KOLOM_FOTO in df_filtered.columns:
        st.markdown("---")
        st.subheader("🖼️ Dokumentasi Visual Aset (Google Drive)")
        
        # Tampilkan dalam bentuk grid 3 kolom
        grid_cols = st.columns(3)
        for idx, row in df_filtered.reset_index(drop=True).iterrows():
            with grid_cols[idx % 3]:
                detail_text = f"⚡ Daya: {row[KOLOM_DAYA]}" if KOLOM_DAYA else ""
                loc_text = f"📍 Lokasi: {row[KOLOM_LOKASI]}" if KOLOM_LOKASI else ""
                
                st.info(f"**{row[KOLOM_UTAMA]}**\n\n{detail_text}\n\n{loc_text}")
                
                if pd.notnull(row[KOLOM_FOTO]) and str(row[KOLOM_FOTO]).strip() != "":
                    img_url = format_drive_image_url(row[KOLOM_FOTO])
                    st.image(img_url, use_container_width=True, caption=f"Asset View - {row[KOLOM_UTAMA]}")
                else:
                    st.warning("Belum ada dokumentasi foto di Google Drive untuk item ini.")

    st.markdown("---")

    # =========================================================================
    # 📥 PANEL GENERATOR EXPORT DATA DYNAMIC (PPTX & PDF)
    # =========================================================================
    st.subheader("📥 Export Laporan Sesuai Pilihan")
    
    # 1. Generator Fungsi PowerPoint (PPTX)
    def generate_pptx(data_target):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"Laporan Analisa Infrastruktur - {selected_item}"
        
        # Ambil maksimal 5 kolom pertama untuk tabel biar muat di slide PPT
        cols_to_include = data_target.columns[:5]
        rows, cols = data_target.shape[0] + 1, len(cols_to_include)
        
        left, top, width, height = Inches(0.5), Inches(2), Inches(9), Inches(3.5)
        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table
        
        # Header Slide Tabel
        for c_idx, col_name in enumerate(cols_to_include):
            table.cell(0, c_idx).text = str(col_name)
            
        # Mengisi baris data
        for r_idx, (_, row) in enumerate(data_target.iterrows()):
            for c_idx, col_name in enumerate(cols_to_include):
                table.cell(r_idx + 1, c_idx).text = str(row[col_name])
                
        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output

    # 2. Generator Fungsi Dokumen PDF
    def generate_pdf(data_target):
        output = BytesIO()
        doc = SimpleDocTemplate(output, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        story.append(Paragraph(f"<b>LAPORAN DATA ASSET INFRASTRUKTUR</b>", styles['Title']))
        story.append(Paragraph(f"Kategori Filter: {selected_item}", styles['Normal']))
        story.append(Spacer(1, 15))
        
        # Konversi dataframe ke list untuk ReportLab Table
        cols_to_print = data_target.columns[:4] # Batasi kolom agar tidak tumpah dari kertas PDF
        table_content = [list(cols_to_print)]
        
        for _, row in data_target.iterrows():
            row_data = [str(row[c]) for c in cols_to_print]
            table_content.append(row_data)
            
        pdf_table = Table(table_content)
        pdf_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#10b981")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#f9fafb")),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#e5e7eb"))
        ]))
        
        story.append(pdf_table)
        doc.build(story)
        output.seek(0)
        return output

    # Tombol Unduh Laporan Dinamis
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        ppt_file = generate_pptx(df_filtered)
        st.download_button(
            label="📊 Download Laporan PowerPoint (.pptx)",
            data=ppt_file,
            file_name=f"Laporan_Asset_{selected_item.replace(' ', '_')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
        
    with col_btn2:
        pdf_file = generate_pdf(df_filtered)
        st.download_button(
            label="📄 Download Laporan PDF (.pdf)",
            data=pdf_file,
            file_name=f"Laporan_Asset_{selected_item.replace(' ', '_')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
