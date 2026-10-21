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
    st.markdown("data:image/jpeg;base64,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' width='120' style='margin-bottom: -10px;'>
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
