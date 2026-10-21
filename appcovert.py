import streamlit as st
import zipfile
import pandas as pd
from bs4 import BeautifulSoup
import io

# 1. BIKIN LAYOUT JADI LEBAR
st.set_page_config(page_title="KMZ TO EXCEL", page_icon="😡", layout="wide")

# 2. INJECT CSS (TAMBAH FONT BARU & FIX WARNA)
st.markdown("""
    <style>
    /* --- IMPORT FONT POPPINS DARI GOOGLE FONTS --- */
    @import url('https://share.google/zBR6gSn2lTbWmBzAW');
    
    /* Paksa semua teks pakai font Poppins dan warna terang */
    html, body, [class*="css"], .stApp, .stApp p, .stApp span, .stApp div, .stApp label, .stApp h3, h1 {
        font-family: 'Poppins', sans-serif !important;
        color: #E2E8F0 !important;
    }
    
    /* Warna background utama (Dark Navy) */
    .stApp {
        background-color: #0F172A !important;
    }
    
    /* Judul Gradasi Keren */
    h1 {
        background: -webkit-linear-gradient(#00c6ff, #0072ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        text-align: center;
        padding-bottom: 10px;
    }
    
    /* --- FIX KOTAK UPLOADER --- */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #0072ff !important;
        border-radius: 15px !important;
        background-color: #1E293B !important; 
        padding: 30px !important;
        transition: all 0.3s ease;
    }
    
    /* Paksa teks di dalem kotak uploader jadi kelihatan */
    [data-testid="stFileUploadDropzone"] div, 
    [data-testid="stFileUploadDropzone"] span, 
    [data-testid="stFileUploadDropzone"] small {
        color: #E2E8F0 !important; 
    }
    
    /* Tombol Upload (Browse files) */
    [data-testid="stFileUploadDropzone"] button {
        background-color: #0072ff !important;
        color: #ffffff !important;
        border: 1px solid #00c6ff !important;
        border-radius: 8px !important;
    }
    
    [data-testid="stFileUploadDropzone"]:hover {
        border-color: #00c6ff !important;
    }
    
    /* --- Tombol Download --- */
    div.stDownloadButton > button {
        background: linear-gradient(90deg, #00C9FF 0%, #92FE9D 100%);
        color: #000 !important;
        border-radius: 30px;
        border: none;
        padding: 10px 25px;
        font-weight: 700;
        box-shadow: 0 4px 15px rgba(0, 201, 255, 0.3);
        transition: all 0.3s ease;
        width: 100%;
        margin-top: 15px;
    }
    div.stDownloadButton > button:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(0, 201, 255, 0.5);
    }
    
    /* Sembunyiin menu default */
    #MainMenu, footer, header {visibility: hidden;}

    /* --- CSS ANIMASI IKON TERBANG --- */
    .bg-animation {
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        z-index: 0; 
        overflow: hidden;
        pointer-events: none; 
    }
    .bg-animation span {
        position: absolute;
        display: block;
        list-style: none;
        width: 20px;
        height: 20px;
        animation: animate 20s linear infinite;
        bottom: -150px;
        font-size: 2rem;
        opacity: 0.2;
    }
    
    .bg-animation span:nth-child(1) { left: 25%; animation-delay: 0s; font-size: 3rem; }
    .bg-animation span:nth-child(2) { left: 10%; animation-delay: 2s; animation-duration: 12s; }
    .bg-animation span:nth-child(3) { left: 70%; animation-delay: 4s; }
    .bg-animation span:nth-child(4) { left: 40%; animation-delay: 0s; animation-duration: 18s; }
    .bg-animation span:nth-child(5) { left: 65%; animation-delay: 0s; font-size: 2.5rem; }
    .bg-animation span:nth-child(6) { left: 75%; animation-delay: 3s; }
    .bg-animation span:nth-child(7) { left: 35%; animation-delay: 7s; }
    .bg-animation span:nth-child(8) { left: 50%; animation-delay: 15s; animation-duration: 45s; }
    .bg-animation span:nth-child(9) { left: 20%; animation-delay: 2s; animation-duration: 35s; font-size: 4rem; }
    .bg-animation span:nth-child(10){ left: 85%; animation-delay: 0s; animation-duration: 11s; }

    @keyframes animate {
        0% { transform: translateY(0) rotate(0deg); opacity: 0.6; }
        100% { transform: translateY(-1000px) rotate(720deg); opacity: 0; }
    }
    </style>
    
    <div class="bg-animation">
        <span>📍</span><span>🌍</span><span>📁</span><span>🗺️</span><span>📍</span>
        <span>🛰️</span><span>🌍</span><span>📌</span><span>🗺️</span><span>🛰️</span>
    </div>
""", unsafe_allow_html=True)

# Layout Form
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    # --- MENAMBAHKAN FOTO/LOGO DI TENGAH ---
    # Lu bisa ganti link gambar di atribut src="..." di bawah ini sama link foto/logo lu sendiri
    st.markdown("""
        <div style='text-align: center;'>
            <img src='data:image/jpeg;base64,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' width='120' style='margin-bottom: -10px;'>
        </div>
    """, unsafe_allow_html=True)
    
    st.title("Geodata Extractor Pro")
    st.markdown("<p style='text-align:center;'>Ubah kumpulan file KMZ/KML lu menjadi tabel Excel secara otomatis. Support Batch Processing.</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    uploaded_files = st.file_uploader("Seret dan Lepaskan file di sini (Support multiple files)", type=['kmz', 'kml'], accept_multiple_files=True)

def process_single_file(file):
    try:
        filename = file.name.lower()
        kml_data = None
        
        if filename.endswith('.kmz'):
            kmz = zipfile.ZipFile(file, 'r')
            kml_filename = [name for name in kmz.namelist() if name.endswith('.kml')][0]
            kml_data = kmz.open(kml_filename, 'r').read()
        elif filename.endswith('.kml'):
            kml_data = file.read()
            
        soup = BeautifulSoup(kml_data, 'xml')
        data_list = []
        
        placemarks = soup.find_all('Placemark')
        for placemark in placemarks:
            row_data = {}
            name_tag = placemark.find('name')
            if name_tag: row_data['Nama Placemark'] = name_tag.text.strip()
                
            point_tag = placemark.find('Point')
            if point_tag:
                coords_tag = point_tag.find('coordinates')
                if coords_tag:
                    coords = coords_tag.text.strip().split(',')
                    if len(coords) >= 2:
                        row_data['Longitude (Asli)'] = float(coords[0].strip())
                        row_data['Latitude (Asli)'] = float(coords[1].strip())
            
            extended_data = placemark.find('ExtendedData')
            if extended_data:
                for data in extended_data.find_all('Data'):
                    name = data.get('name')
                    value = data.find('value')
                    if name and value: row_data[name] = value.text.strip()
                for simple_data in extended_data.find_all('SimpleData'):
                    name = simple_data.get('name')
                    if name: row_data[name] = simple_data.text.strip()
            
            desc_tag = placemark.find('description')
            if desc_tag and desc_tag.text:
                html_soup = BeautifulSoup(desc_tag.text, 'html.parser')
                tables = html_soup.find_all('table')
                if tables:
                    for table in tables:
                        for tr in table.find_all('tr'):
                            cells = tr.find_all(['td', 'th'])
                            if len(cells) >= 2:
                                key = cells[0].text.strip().replace(':', '')
                                val = cells[1].text.strip()
                                if key.lower() in ['lat', 'latitude']: key = 'Latitude (Tabel)'
                                elif key.lower() in ['long', 'longitude']: key = 'Longitude (Tabel)'
                                row_data[key] = val
                                
            data_list.append(row_data)
        
        if not data_list: return None, f"⚠️ Tidak ada data Placemark di {file.name}"
            
        df = pd.DataFrame(data_list)
        cols = df.columns.tolist()
        front_cols = ['Nama Placemark', 'Latitude (Asli)', 'Longitude (Asli)']
        front_cols = [c for c in front_cols if c in cols]
        other_cols = [c for c in cols if c not in front_cols]
        return df[front_cols + other_cols], None
    except Exception as e:
        return None, f"❌ Error di {file.name}: {e}"

if uploaded_files:
    st.markdown("<h3 style='text-align:center;'>📊 Hasil Proses</h3>", unsafe_allow_html=True)
    with st.spinner(f"Sedang meracik {len(uploaded_files)} file..."):
        
        if len(uploaded_files) == 1:
            file = uploaded_files[0]
            df, error = process_single_file(file)
            
            if error: st.error(error)
            else:
                st.success(f"✅ Mantap! Ketemu {len(df)} titik dari {file.name}")
                st.dataframe(df, height=250, use_container_width=True)
                
                base_name = file.name.rsplit('.', 1)[0]
                export_name = f"{base_name}_extracted.xlsx"
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                
                col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
                with col_btn2:
                    st.download_button(
                        label=f"📥 DOWNLOAD {export_name.upper()}",
                        data=output.getvalue(),
                        file_name=export_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            zip_buffer = io.BytesIO()
            berhasil_count, gagal_count = 0, 0
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for file in uploaded_files:
                    df, error = process_single_file(file)
                    if error:
                        st.error(error)
                        gagal_count += 1
                    else:
                        base_name = file.name.rsplit('.', 1)[0]
                        export_name = f"{base_name}_extracted.xlsx"
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False)
                        zip_file.writestr(export_name, excel_buffer.getvalue())
                        berhasil_count += 1
                        
            if berhasil_count > 0:
                st.success(f"✅ Selesai Cuy! {berhasil_count} file berhasil dibungkus ke dalam 1 ZIP.")
                if gagal_count > 0: st.warning(f"⚠️ Ada {gagal_count} file gagal.")
                
                col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
                with col_btn2:
                    st.download_button(
                        label="📦 DOWNLOAD SEMUA (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="Geodata_Batch_Export.zip",
                        mime="application/zip"
                    )
            else:
                st.error("❌ Semua file gagal diproses.")
