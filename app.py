import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
from io import BytesIO

# Pustaka untuk Export PPT & PDF
from pptx import Presentation
from pptx.util import Inches, Pt

# =========================================================================
# ⚙️ KONFIGURASI HALAMAN UTAMA
# =========================================================================
st.set_page_config(
    page_title="Infrastructure Asset Report",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ID Resmi Spreadsheet 348 Tsel Anda
SPREADSHEET_ID = "1FGKOzWoUrbf3PXN_ahgG1t-83JZT4H4sioQepePbBxM"
csv_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:csv"

# 🔥 ENGINE RE-FORMAT LINK GOOGLE DRIVE AGAR WAJIB TEMBUS & TAMPIL DI STREAMLIT
def format_drive_image_url(url_or_id):
    val = str(url_or_id).strip()
    if val == "" or val.lower() == "nan" or "http" not in val.lower():
        # Cek jika isinya sudah berupa ID murni tanpa http
        if len(val) >= 25 and "/" not in val:
            return f"https://drive.google.com/uc?export=view&id={val}"
        return None
    
    # Ekstraksi ID 33 karakter unik dari link sharing panjang Google Drive
    match = re.search(r'([a-zA-Z0-9_-]{33})', val)
    if match:
        doc_id = match.group(1)
        return f"https://drive.google.com/uc?export=view&id={doc_id}"
    return val

st.title("⚡ Infrastructure Asset & Site Report")
st.caption("Sistem Monitoring Terintegrasi Google Sheets (348 Tsel) & Dokumentasi Visual Drive")
st.markdown("---")

# =========================================================================
# 💾 PROSES PENGAMBILAN DATA
# =========================================================================
@st.cache_data(ttl=2)
def load_financial_data(url):
    try:
        df = pd.read_csv(url)
        return df
    except Exception as e:
        return str(e)

with st.spinner("🔄 Sedang menyinkronkan data lapangan dari Google Sheets..."):
    df_raw = load_financial_data(csv_url)

if isinstance(df_raw, str):
    st.error(f"⚠️ Gagal memuat data dari Spreadsheet. Detail: `{df_raw}`")
    st.stop()
elif df_raw is None or df_raw.empty:
    st.warning("⚠️ Data di dalam lembar kerja spreadsheet kosong.")
    st.stop()
else:
    # Bersihkan spasi kosong pada nama-nama kolom
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    df = df_raw.dropna(how='all', axis=0).copy()

    # =========================================================================
    # 🔍 PEMETAAN HARDCODE KOLOM SPREADSHEET ANDA (ANTI-MELESET)
    # =========================================================================
    # Mencari kolom Site ID Anda secara cerdas
    KOLOM_UTAMA = next((c for c in df.columns if any(p in c.upper() for p in ['SITE ID', 'SITE_ID', 'SITEID', 'TOWER ID'])), df.columns[0])
    
    # Mencari kolom link gambar yang Anda ketik di Google Sheets
    KOLOM_FOTO = next((c for c in df.columns if any(p in c.upper() for p in ['LINK FOTO', 'DOKUMENTASI', 'DRIVE', 'FOTO', 'URL FOTO', 'GAMBAR', 'LINK'])), None)
    
    # Mencari kolom Voltase fasa RST
    KOLOM_VR = next((c for c in df.columns if any(p in c.upper() for p in ['VOLTAGE R', 'TEGANGAN R', 'VR', 'V_R'])), None)
    KOLOM_VS = next((c for c in df.columns if any(p in c.upper() for p in ['VOLTAGE S', 'TEGANGAN S', 'VS', 'V_S'])), None)
    KOLOM_VT = next((c for c in df.columns if any(p in c.upper() for p in ['VOLTAGE T', 'TEGANGAN T', 'VT', 'V_T'])), None)

    # --- SIDEBAR: PENCARIAN & FILTER SITE ID ---
    st.sidebar.header("⚙️ Kontrol Pencarian")
    df[KOLOM_UTAMA] = df[KOLOM_UTAMA].fillna("Kosong").astype(str).str.strip()
    
    search_query = st.sidebar.text_input("🔍 Cari Masukkan Site ID:", "").strip()
    
    if search_query:
        df_filtered = df[df[KOLOM_UTAMA].str.contains(search_query, case=False, na=False)]
    else:
        unique_items = sorted(list(df[KOLOM_UTAMA].unique()))
        selected_item = st.sidebar.selectbox("Atau Pilih Sesuai List ID:", unique_items)
        df_filtered = df[df[KOLOM_UTAMA] == selected_item]

    if st.sidebar.button("🔄 Paksa Muat Ulang Data (Clear Cache)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    # =========================================================================
    # 🏗️ RE-DESIGN LAYOUT 3 BAGIAN UTAMA (FOTO DI BAGIAN KHUSUS & RAPI)
    # =========================================================================
    if df_filtered.empty:
        st.warning("⚠️ Data laporan untuk pencarian tersebut tidak ditemukan.")
    else:
        for idx, row in df_filtered.reset_index(drop=True).iterrows():
            
            # Pembungkus Kontainer Utama per Item agar Tampilan Terkunci Rapi
            with st.container(border=True):
                st.markdown(f"### 📍 SITE REPORT ID: <span style='color:#10b981'>{row[KOLOM_UTAMA]}</span>", unsafe_allow_html=True)
                
                # Membagi Layar Menjadi 3 Bagian Layout Sesuai Permintaan Anda
                layout_col1, layout_col2, layout_col3 = st.columns([1.5, 1.2, 1.3])
                
                # --- BAGIAN 1: KETERANGAN HURUF DAN ANGKA YANG RAPI ---
                with layout_col1:
                    st.markdown("📝 **Keterangan & Spesifikasi Teknik**")
                    info_text = ""
                    for col in df_filtered.columns:
                        # Tampilkan seluruh metadata teks terkecuali data gambar dan voltase fasa
                        if col not in [KOLOM_UTAMA, KOLOM_FOTO, KOLOM_VR, KOLOM_VS, KOLOM_VT]:
                            val_str = str(row[col]).strip()
                            if val_str != "" and val_str.lower() != "nan":
                                info_text += f"**{col}** : {val_str}  \n"
                    
                    if info_text:
                        st.info(info_text)
                    else:
                        st.caption("*Tidak ada parameter huruf/angka tambahan.*")

                # --- BAGIAN 2: STATUS POLA DAYA & TEGANGAN (ANGKA METRIK) ---
                with layout_col2:
                    st.markdown("📊 **Pola Tegangan Arus (R-S-T)**")
                    vr_val = row[KOLOM_VR] if KOLOM_VR and pd.notnull(row[KOLOM_VR]) else "-"
                    vs_val = row[KOLOM_VS] if KOLOM_VS and pd.notnull(row[KOLOM_VS]) else "-"
                    vt_val = row[KOLOM_VT] if KOLOM_VT and pd.notnull(row[KOLOM_VT]) else "-"
                    
                    # Layout angka metrik internal fasa
                    sub_c1, sub_c2, sub_c3 = st.columns(3)
                    sub_c1.metric("Fasa R", f"{vr_val} V")
                    sub_c2.metric("Fasa S", f"{vs_val} V")
                    sub_c3.metric("Fasa T", f"{vt_val} V")
                    
                    # Grafik mini pembentuk pola kestabilan daya fasa
                    try:
                        v_r_num = float(str(vr_val).replace('V','').strip() or 0)
                        v_s_num = float(str(vs_val).replace('V','').strip() or 0)
                        v_t_num = float(str(vt_val).replace('V','').strip() or 0)
                        volt_df = pd.DataFrame({"Tegangan (Volt)": [v_r_num, v_s_num, v_t_num]}, index=["Fasa R", "Fasa S", "Fasa T"])
                        st.bar_chart(volt_df, height=130)
                    except:
                        pass

                # --- BAGIAN 3: TEMPAT KHUSUS FOTO DARI DRIVE (WAJIB KELIHATAN) ---
                with layout_col3:
                    st.markdown("🖼️ **Dokumentasi Hasil Lapangan**")
                    cell_foto_content = str(row[KOLOM_FOTO]).strip() if KOLOM_FOTO else ""
                    
                    if cell_foto_content != "" and cell_foto_content.lower() != "nan":
                        img_url_stream = format_drive_image_url(cell_foto_content)
                        if img_url_stream:
                            # Trik penayangan bypass via markdown HTML img tag untuk menjamin foto tidak pecah
                            st.markdown(
                                f'<img src="{img_url_stream}" style="width:100%; border-radius:8px; border:2px solid #e5e7eb;" alt="Memuat Dokumentasi Drive...">', 
                                unsafe_allow_html=True
                            )
                            st.caption(f"Visual Validasi - ID {row[KOLOM_UTAMA]}")
                        else:
                            st.warning("⚠️ Format tautan di dalam kolom foto tidak valid.")
                    else:
                        st.error("❌ Link foto kosong / belum diinput pada baris sheet ini.")
            
            st.markdown("<br>", unsafe_allow_html=True)

        # =========================================================================
        # 📥 TARIKAN DATA PPT YANG MERANGKUM TOTAL PER SITE ID LENGKAP DENGAN FOTO
        # =========================================================================
        st.markdown("---")
        st.subheader("📥 Penarikan Dokumen Laporan Akhir")
        
        def generate_total_report_pptx(data_source):
            prs = Presentation()
            for _, r_data in data_source.iterrows():
                # Membuat slide kosong baru per baris Site ID
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                
                # Pembuatan Banner Judul atas Slide PPT
                title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.8))
                tf = title_box.text_frame
                p_title = tf.paragraphs[0]
                p_title.text = f"LAPORAN DATA TEKNIS INFRASTRUKTUR: {r_data[KOLOM_UTAMA]}"
                p_title.font.size = Pt(20)
                p_title.font.bold = True
                
                # Penyusunan Teks Huruf & Angka Keterangan (Sisi Kiri Slide PPT)
                text_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(4.5), Inches(5.3))
                text_tf = text_box.text_frame
                text_tf.word_wrap = True
                
                # Masukkan ringkasan fasa voltase di baris awal PPT
                v_line = text_tf.paragraphs[0]
                vr = r_data[KOLOM_VR] if KOLOM_VR else "-"
                vs = r_data[KOLOM_VS] if KOLOM_VS else "-"
                vt = r_data[KOLOM_VT] if KOLOM_VT else "-"
                v_line.text = f"Status Pola Voltase RST: R={vr}V | S={vs}V | T={vt}V\n"
                v_line.font.bold = True
                v_line.font.size = Pt(12)
                
                for col_name in data_source.columns:
                    if col_name not in [KOLOM_UTAMA, KOLOM_FOTO, KOLOM_VR, KOLOM_VS, KOLOM_VT]:
                        val_txt = str(r_data[col_name]).strip()
                        if val_txt != "" and val_txt.lower() != "nan":
                            p_spec = text_tf.add_paragraph()
                            p_spec.text = f"• {col_name}: {val_txt}"
                            p_spec.font.size = Pt(11)
                
                # Proses Download Otomatis Gambar Dari Drive Dan Tempel ke Sisi Kanan Slide PPT
                f_link = str(r_data[KOLOM_FOTO]).strip() if KOLOM_FOTO else ""
                if f_link != "" and f_link.lower() != "nan":
                    g_direct_link = format_drive_image_url(f_link)
                    try:
                        # Request penarikan byte image internet langsung dari server Google Drive
                        resp = requests.get(g_direct_link, timeout=6)
                        if resp.status_code == 200:
                            img_bytes = BytesIO(resp.content)
                            # Letakkan foto di kanan slide (Sesuai tatanan layout dashboard)
                            slide.shapes.add_picture(img_bytes, Inches(5.3), Inches(1.3), width=Inches(4.2))
                    except Exception:
                        err_box = slide.shapes.add_textbox(Inches(5.3), Inches(1.3), Inches(4.2), Inches(2))
                        err_box.text_frame.text = "[Gambar gagal dimuat otomatis karena masalah koneksi atau hak akses file Drive Dibatasi]"
            
            output_buffer = BytesIO()
            prs.save(output_buffer)
            output_buffer.seek(0)
            return output_buffer

        if st.button("📊 Generate Dokumen PowerPoint (.PPTX) Berdasarkan Filter", type="primary", use_container_width=True):
            with st.spinner("⏳ Menarik data huruf, angka, dan aset foto dari Google Drive untuk disusun ke PPTX..."):
                final_ppt = generate_total_report_pptx(df_filtered)
                st.download_button(
                    label="📥 Klik di Sini Untuk Mengunduh File PPTX Hasil Report",
                    data=final_ppt,
                    file_name=f"Laporan_Lengkap_Site_{search_query or 'Selected'}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
