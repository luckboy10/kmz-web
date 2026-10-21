import streamlit as st
import zipfile
import pandas as pd
from bs4 import BeautifulSoup
import io

# 1. BIKIN LAYOUT JADI LEBAR
st.set_page_config(page_title="Geodata Extractor", page_icon="📍", layout="wide")

# 2. INJECT CSS (FIX TEKS & ANIMASI TERBANG)
st.markdown("""
    <style>
    /* FIX TEKS: Paksa semua teks jadi warna terang biar kebaca di background gelap */
    .stApp, .stApp p, .stApp span, .stApp div, .stApp label, .stApp h3 {
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
        font-family: 'Segoe UI', sans-serif;
        font-weight: 800;
        text-align: center;
        padding-bottom: 10px;
    }
    
    /* Area Uploader File */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #0072ff !important;
        border-radius: 15px !important;
        background-color: rgba(0, 114, 255, 0.1) !important;
        padding: 30px !important;
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        background-color: rgba(0, 114, 255, 0.2) !important;
        border-color: #00c6ff !important;
    }
    
    /* Tombol Download Jadi Estetik */
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
        z-index: 0; /* Taruh di belakang */
        overflow: hidden;
        pointer-events: none; /* Biar kaga ngehalangin klik */
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
        opacity: 0.2; /* Dibuat agak transparan biar teks tetep jelas */
    }
    
    /* Variasi posisi, ukuran, dan kecepatan per ikon */
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
    
    <!-- INJEK HTML IKON TERBANG -->
    <div class="bg-animation">
        <span>📍</span>
        <span>🌍</span>
        <span>📁</span>
        <span>🗺️</span>
        <span>📍</span>
        <span>🛰️</span>
        <span>🌍</span>
        <span>📌</span>
        <span>🗺️</span>
        <span>🛰️</span>
    </div>
""", unsafe_allow_html=True)

# Layout: Bikin kolom biar form-nya di tengah
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    st.title("📍 Geodata Extractor Pro")
    st.markdown("<p style='text-align:center;'>Ubah kumpulan file KMZ/KML lu menjadi tabel Excel secara otomatis. Support Batch Processing.</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    uploaded_files = st.file_uploader("Seret dan Lepaskan file di sini (Support multiple files)", type=['kmz', 'kml'], accept_multiple_files=True)

# Fungsi pemrosesan file (Tidak diubah, murni backend)
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

# Eksekusi kalau file di-upload
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
