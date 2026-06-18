import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
from io import BytesIO

# Pustaka untuk Export PPT & PDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# =========================================================================
# ⚙️ KONFIGURASI HALAMAN
# =========================================================================
st.set_page_config(
    page_title="Infrastructure Asset Report",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

SPREADSHEET_ID = "1FGKOzWoUrbf3PXN_ahgG1t-83JZT4H4sioQepePbBxM"
csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv"

# Fungsi bypass Google Drive agar foto wajib muncul langsung di Streamlit & PPT
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

st.title("⚡ Infrastructure Asset & Site Report")
st.caption("Live Synchronization with Google Sheets (348 Tsel) & Full PPT Document Generation")
st.markdown("---")

# =========================================================================
# 💾 LOAD DATA
# =========================================================================
@st.cache_data(ttl=2)
def load_financial_data(url):
    try:
        df = pd.read_csv(url)
        return df
    except Exception as e:
        return str(e)

with st.spinner("🔄 Memuat database ..."):
    df_raw = load_financial_data(csv_url)

if isinstance(df_raw, str):
    st.error(f"⚠️ Gagal memuat data. Detail Error: `{df_raw}`")
    st.stop()
elif df_raw is None or df_raw.empty:
    st.warning("⚠️ Data di dalam spreadsheet kosong.")
    st.stop()
else:
    df = df_raw.dropna(how='all', axis=1).copy()
    df = df.dropna(how='all', axis=0)

    # 🔍 DETEKSI DUA PILIHAN PENCARIAN (SITE ID ATAU COMPONENT)
    KOLOM_UTAMA = next((c for c in df.columns if any(p in str(c).upper() for p in ['SITE ID', 'SITE_ID', 'TOWER ID', 'COMPONENT'])), df.columns[0])
    KOLOM_DAYA = next((c for c in df.columns if any(p in str(c).upper() for p in ['DAYA', 'KAPASITAS', 'POWER'])), None)
    KOLOM_LOKASI = next((c for c in df.columns if any(p in str(c).upper() for p in ['LOKASI', 'LOCATION', 'ALAMAT', 'WITEL'])), None)
    KOLOM_FOTO = next((c for c in df.columns if any(p in str(c).upper() for p in ['FOTO', 'LINK FOTO', 'GAMBAR', 'DRIVE IMAGE'])), None)

    # Deteksi fasa Voltage
    KOLOM_VR = next((c for c in df.columns if any(p in str(c).upper() for p in ['VOLTAGE R', 'TEGANGAN R', 'VR'])), None)
    KOLOM_VS = next((c for c in df.columns if any(p in str(c).upper() for p in ['VOLTAGE S', 'TEGANGAN S', 'VS'])), None)
    KOLOM_VT = next((c for c in df.columns if any(p in str(c).upper() for p in ['VOLTAGE T', 'TEGANGAN T', 'VT'])), None)

    # --- SIDEBAR: PENCARIAN SITE ID ---
    st.sidebar.header("⚙️ Pencarian Site ID / Asset")
    df[KOLOM_UTAMA] = df[KOLOM_UTAMA].fillna("Unknown").astype(str).str.strip()
    
    search_query = st.sidebar.text_input("🔍 Cari Berdasarkan Site ID:", "")
    
    if search_query:
        df_filtered = df[df[KOLOM_UTAMA].str.contains(search_query, case=False, na=False)]
    else:
        unique_items = sorted(list(df[KOLOM_UTAMA].unique()))
        selected_item = st.sidebar.selectbox(f"Atau Pilih dari Daftar {KOLOM_UTAMA}:", unique_items)
        df_filtered = df[df[KOLOM_UTAMA] == selected_item]

    if st.sidebar.button("🔄 Clear Cache & Refresh", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # =========================================================================
    # 🏗️ LAYOUT BARU: DIBUAT 3 BAGIAN UTAMA YANG RAPI per BARIS DATA
    # =========================================================================
    if df_filtered.empty:
        st.warning("⚠️ Site ID atau data tidak ditemukan.")
    else:
        for idx, row in df_filtered.reset_index(drop=True).iterrows():
            st.markdown(f"### 📍 Site Profil: {row[KOLOM_UTAMA]}")
            
            # Membuat Pembagian Layout Tiga Bagian Sejajar
            layout_col1, layout_col2, layout_col3 = st.columns([1.5, 1.2, 1.3])
            
            # --- BAGIAN 1: KETERANGAN HURUF DAN ANGKA YANG RAPI ---
            with layout_col1:
                st.markdown("📝 **Spesifikasi & Informasi Teknis**")
                
                # Menampilkan metrik utama daya jika ada
                if KOLOM_DAYA and pd.notnull(row[KOLOM_DAYA]):
                    st.metric(label="Kapasitas / Daya Terpasang", value=f"{row[KOLOM_DAYA]}")
                
                # Tampilkan seluruh kolom teks pendukung dalam bentuk list ringkas yang bersih
                info_html = ""
                for col in df_filtered.columns:
                    if col not in [KOLOM_UTAMA, KOLOM_FOTO, KOLOM_VR, KOLOM_VS, KOLOM_VT, KOLOM_DAYA]:
                        val_str = str(row[col]).strip()
                        if val_str != "" and val_str.lower() != "nan":
                            info_html += f"**{col}**: {val_str}  \n"
                st.markdown(info_html if info_html else "*Tidak ada keterangan tambahan.*")

            # --- BAGIAN 2: POLA POWER & VOLTAGE ANALYTICS ---
            with layout_col2:
                st.markdown("📈 **Status Pola Tegangan (R-S-T)**")
                vr = row[KOLOM_VR] if KOLOM_VR and pd.notnull(row[KOLOM_VR]) else 0
                vs = row[KOLOM_VS] if KOLOM_VS and pd.notnull(row[KOLOM_VS]) else 0
                vt = row[KOLOM_VT] if KOLOM_VT and pd.notnull(row[KOLOM_VT]) else 0
                
                # Mini Grid internal untuk status angka voltase fasa
                v_c1, v_c2, v_c3 = st.columns(3)
                v_c1.metric("Fasa R", f"{vr}V")
                v_c2.metric("Fasa S", f"{vs}V")
                v_c3.metric("Fasa T", f"{vt}V")
                
                # Tampilkan pola visual dalam bentuk chart batang mini khusus per fasa
                volt_df = pd.DataFrame({"Voltase (V)": [float(str(vr).replace('V','').strip() or 0), 
                                                        float(str(vs).replace('V','').strip() or 0), 
                                                        float(str(vt).replace('V','').strip() or 0)]}, 
                                       index=["Rasa R", "Fasa S", "Fasa T"])
                st.bar_chart(volt_df, height=140)

            # --- BAGIAN 3: FOTO HASIL DARI GOOGLE DRIVE (WAJIB KELIHATAN) ---
            with layout_col3:
                st.markdown("🖼️ **Dokumentasi Hasil Lapangan**")
                cell_foto = str(row[KOLOM_FOTO]).strip() if KOLOM_FOTO else ""
                
                if cell_foto != "" and cell_foto.lower() != "nan":
                    img_url = format_drive_image_url(cell_foto)
                    if img_url:
                        st.image(img_url, use_container_width=True, caption=f"Kondisi Fisik {row[KOLOM_UTAMA]}")
                else:
                    st.warning("⚠️ Foto belum diinput atau link kosong di Google Sheets.")
            
            st.markdown("---")

        # =========================================================================
        # 📥 PANEL EXPORT PPTX MERANGKUM LENGKAP DENGAN GAMBAR & LAYOUT
        # =========================================================================
        st.subheader("📥 Export Laporan Komprehensif per Site ID")
        
        def generate_full_report_pptx(data_target):
            prs = Presentation()
            
            for _, r_data in data_target.iterrows():
                # Gunakan Blank Layout agar peletakan text dan gambar bisa di-custom mirip layout dashboard
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                
                # 1. Judul Slide (Header)
                tx_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.8))
                tf = tx_box.text_frame
                p = tf.paragraphs[0]
                p.text = f"SITE REPORT: {r_data[KOLOM_UTAMA]}"
                p.font.size = Pt(24)
                p.font.bold = True
                
                # 2. Kotak Keterangan Huruf & Angka (Sebelah Kiri)
                desc_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(4.5), Inches(5.5))
                desc_tf = desc_box.text_frame
                desc_tf.word_wrap = True
                
                p_daya = desc_tf.paragraphs[0]
                p_daya.text = f"Daya/Kapasitas: {r_data[KOLOM_DAYA] if KOLOM_DAYA else '-'}\n"
                p_daya.font.bold = True
                
                # Tambahkan baris keterangan teknis lainnya
                for col_name in data_target.columns:
                    if col_name not in [KOLOM_UTAMA, KOLOM_FOTO, KOLOM_VR, KOLOM_VS, KOLOM_VT, KOLOM_DAYA]:
                        val_txt = str(r_data[col_name]).strip()
                        if val_txt != "" and val_txt.lower() != "nan":
                            p_info = desc_tf.add_paragraph()
                            p_info.text = f"• {col_name}: {val_txt}"
                            p_info.font.size = Pt(11)
                
                # Tambahkan keterangan Voltage RST di bagian bawah teks kiri
                p_volt = desc_tf.add_paragraph()
                p_volt.text = f"\nVoltase RST: R={r_data[KOLOM_VR]}V | S={r_data[KOLOM_VS]}V | T={r_data[KOLOM_VT]}V"
                p_volt.font.size = Pt(12)
                p_volt.font.bold = True
                p_volt.font.color.rgb = colors.HexColor("#10b981")

                # 3. Download & Tempel Gambar Dokumentasi Google Drive (Sebelah Kanan)
                f_cell = str(r_data[KOLOM_FOTO]).strip() if KOLOM_FOTO else ""
                if f_cell != "" and f_cell.lower() != "nan":
                    g_url = format_drive_image_url(f_cell)
                    try:
                        # Stream gambar langsung dari URL internet ke memori PPTX
                        response = requests.get(g_url, timeout=5)
                        if response.status_code == 200:
                            img_stream = BytesIO(response.content)
                            # Letakkan di sisi kanan slide (Left=5.2 inci, Top=1.5 inci, Width=4.3 inci)
                            slide.shapes.add_picture(img_stream, Inches(5.2), Inches(1.5), width=Inches(4.3))
                    except Exception:
                        # Jika timeout/gagal download gambar, beri penanda box teks error di PPT
                        err_box = slide.shapes.add_textbox(Inches(5.2), Inches(1.5), Inches(4.3), Inches(2))
                        err_box.text_frame.text = "[Gagal memuat visual dokumentasi dari Drive ke PPT]"
            
            output = BytesIO()
            prs.save(output)
            output.seek(0)
            return output

        # Tombol aksi download PPT yang merangkum seluruh layout dan data fasa
        if st.button("📊 Generate & Download PPTX Report Terpilih", type="primary", use_container_width=True):
            with st.spinner("⏳ Mengunduh gambar & menyusun struktur slide PPTX..."):
                full_ppt = generate_full_report_pptx(df_filtered)
                st.download_button(
                    label="📥 Klik di Sini Untuk Mengunduh File .PPTX",
                    data=full_ppt,
                    file_name=f"Full_Report_Site_{search_query or 'Filtered'}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
