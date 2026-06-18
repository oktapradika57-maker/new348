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
    page_title="PLN & Infrastructure Asset Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ID Spreadsheet "348 Tsel"
SPREADSHEET_ID = "1FGKOzWoUrbf3PXN_ahgG1t-83JZT4H4sioQepePbBxM"
# Menggunakan format gviz/tq kembali untuk memastikan seluruh data struktural ditarik dengan aman
csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv"

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
@st.cache_data(ttl=2)
def load_financial_data(url):
    try:
        df = pd.read_csv(url)
        return df
    except Exception as e:
        return str(e)

with st.spinner("🔄 Sedang memuat resource data terbaru dari Google Sheets..."):
    df_raw = load_financial_data(csv_url)

if isinstance(df_raw, str):
    st.error("⚠️ Gagal memuat data dari link Google Sheets yang diberikan.")
    st.markdown(f"**Detail Error:** `{df_raw}`")
    st.stop()

elif df_raw is None or df_raw.empty:
    st.warning("⚠️ Data yang diunduh kosong. Silakan periksa apakah tab pertama di Google Sheets Anda berisi data atau kosong.")
    st.stop()

else:
    # Bersihkan baris/kolom yang benar-benar kosong
    df = df_raw.dropna(how='all', axis=1).copy()
    df = df.dropna(how='all', axis=0)

    # =========================================================================
    # 🔍 DETEKSI KOLOM SECARA DINAMIS (DENGAN FALLBACK AMAN)
    # =========================================================================
    columns_list = [str(c).strip() for c in df.columns]
    
    KOLOM_UTAMA = next((c for c in df.columns if any(p in str(c).upper() for p in ['SITE ID', 'SITE_ID', 'NAMA GARDU', 'COMPONENT', 'TOWER ID', 'NAMA'])), df.columns[0])
    KOLOM_DAYA = next((c for c in df.columns if any(p in str(c).upper() for p in ['DAYA', 'KAPASITAS', 'POWER'])), None)
    KOLOM_LOKASI = next((c for c in df.columns if any(p in str(c).upper() for p in ['LOKASI', 'LOCATION', 'ALAMAT', 'WITEL'])), None)
    KOLOM_FOTO = next((c for c in df.columns if any(p in str(c).upper() for p in ['FOTO', 'LINK FOTO', 'GAMBAR', 'DRIVE IMAGE'])), None)

    # Deteksi fasa Voltage RST
    KOLOM_VR = next((c for c in df.columns if any(p in str(c).upper() for p in ['VOLTAGE R', 'TEGANGAN R', 'VOLT R', 'V_R', 'VR'])), None)
    KOLOM_VS = next((c for c in df.columns if any(p in str(c).upper() for p in ['VOLTAGE S', 'TEGANGAN S', 'VOLT S', 'V_S', 'VS'])), None)
    KOLOM_VT = next((c for c in df.columns if any(p in str(c).upper() for p in ['VOLTAGE T', 'TEGANGAN T', 'VOLT T', 'V_T', 'VT'])), None)

    # --- SIDEBAR PANEL FILTER ---
    st.sidebar.header("⚙️ Panel Filter Resource")
    
    # Pastikan data kolom utama dikonversi ke string dan dibersihkan dari nan
    df[KOLOM_UTAMA] = df[KOLOM_UTAMA].fillna("Tanpa Nama").astype(str)
    unique_items = ["Semua Komponen"] + sorted(list(df[KOLOM_UTAMA].unique()))
    selected_item = st.sidebar.selectbox(f"Pilih {KOLOM_UTAMA}:", unique_items)

    if st.sidebar.button("🔄 Clear Cache & Refresh", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # --- PROSES FILTER DATA ---
    df_filtered = df.copy()
    if selected_item != "Semua Komponen":
        df_filtered = df_filtered[df_filtered[KOLOM_UTAMA] == selected_item]

    # =========================================================================
    # 📊 SEKSI GRAFIK ANALISA TEGANGAN VOLTAGE RST
    # =========================================================================
    st.subheader("📈 Pola Power & Grafik Analisa Tegangan RST")
    
    if (KOLOM_VR or KOLOM_VS or KOLOM_VT) and not df_filtered.empty:
        try:
            chart_data = df_filtered.copy()
            build_chart = pd.DataFrame(index=chart_data[KOLOM_UTAMA])
            
            # Bersihkan nilai string ke numerik secara aman
            for name, col_obj in [('Voltage R', KOLOM_VR), ('Voltage S', KOLOM_VS), ('Voltage T', KOLOM_VT)]:
                if col_obj:
                    clean_val = chart_data[col_obj].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                    build_chart[name] = pd.to_numeric(clean_val, errors='coerce').fillna(0).values
            
            st.line_chart(build_chart)
        except Exception as chart_err:
            st.info("💡 Grafik tren belum dapat ditampilkan karena format tipe data numerik fasa sedang disinkronkan.")
    else:
        st.info("💡 Pilih salah satu item atau pastikan fasa Voltage R, S, T terisi untuk melihat grafik tren daya.")

    st.markdown("---")

    # --- TAMPILAN TABEL UTAMA ---
    st.subheader("📋 Detail Rincian Data")
    st.dataframe(df_filtered, use_container_width=True, hide_index=True)

    # --- BAGIAN VISUALISASI GAMBAR GOOGLE DRIVE ---
    if KOLOM_FOTO and KOLOM_FOTO in df_filtered.columns:
        st.markdown("---")
        st.subheader("🖼️ Dokumentasi Visual Aset (Google Drive)")
        
        grid_cols = st.columns(3)
        for idx, row in df_filtered.reset_index(drop=True).iterrows():
            with grid_cols[idx % 3]:
                detail_text = f"⚡ Daya: {row[KOLOM_DAYA]}" if KOLOM_DAYA else ""
                loc_text = f"📍 Lokasi: {row[KOLOM_LOKASI]}" if KOLOM_LOKASI else ""
                
                vr_val = f" R: {row[KOLOM_VR]}V" if KOLOM_VR else ""
                vs_val = f" S: {row[KOLOM_VS]}V" if KOLOM_VS else ""
                vt_val = f" T: {row[KOLOM_VT]}V" if KOLOM_VT else ""
                volt_summary = f"📊 Voltase:{vr_val}{vs_val}{vt_val}" if (KOLOM_VR or KOLOM_VS or KOLOM_VT) else ""

                st.info(f"**{row[KOLOM_UTAMA]}**\n\n{detail_text}\n\n{loc_text}\n\n{volt_summary}")
                
                cell_foto_val = str(row[KOLOM_FOTO]).strip()
                if cell_foto_val != "" and cell_foto_val.lower() != "nan":
                    img_url = format_drive_image_url(cell_foto_val)
                    if img_url:
                        st.image(img_url, use_container_width=True, caption=f"Asset View - {row[KOLOM_UTAMA]}")
                else:
                    st.warning("⚠️ Kolom foto kosong / belum diisi.")

    st.markdown("---")

    # =========================================================================
    # 📥 PANEL GENERATOR EXPORT DATA DYNAMIC (PPTX & PDF)
    # =========================================================================
    st.subheader("📥 Export Laporan Sesuai Pilihan")
    
    def generate_pptx(data_target):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"Laporan Analisa Infrastruktur"
        cols_to_include = data_target.columns[:5]
        rows, cols = data_target.shape[0] + 1, len(cols_to_include)
        left, top, width, height = Inches(0.5), Inches(2), Inches(9), Inches(3.5)
        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table
        for c_idx, col_name in enumerate(cols_to_include):
            table.cell(0, c_idx).text = str(col_name)
        for r_idx, (_, row) in enumerate(data_target.iterrows()):
            for c_idx, col_name in enumerate(cols_to_include):
                table.cell(r_idx + 1, c_idx).text = str(row[col_name])
        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output

    def generate_pdf(data_target):
        output = BytesIO()
        doc = SimpleDocTemplate(output, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        story.append(Paragraph(f"<b>LAPORAN DATA ASSET INFRASTRUKTUR</b>", styles['Title']))
        story.append(Spacer(1, 15))
        cols_to_print = data_target.columns[:4]
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

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        ppt_file = generate_pptx(df_filtered)
        st.download_button(
            label="📊 Download Laporan PowerPoint (.pptx)",
            data=ppt_file,
            file_name="Laporan_Asset_PLN.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )
    with col_btn2:
        pdf_file = generate_pdf(df_filtered)
        st.download_button(
            label="📄 Download Laporan PDF (.pdf)",
            data=pdf_file,
            file_name="Laporan_Asset_PLN.pdf",
            mime="application/pdf",
            use_container_width=True
        )
