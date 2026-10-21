import streamlit as st
import zipfile
import pandas as pd
from bs4 import BeautifulSoup
import io

# 1. BIKIN LAYOUT JADI LEBAR (WIDE)
st.set_page_config(page_title="Geodata Extractor", page_icon="📍", layout="wide")

# 2. INJECT CUSTOM CSS BIAR TAMPILAN KEREN
st.markdown("""
    <style>
    /* Ubah background utama (Dark Theme yang elegan) */
    .stApp {
        background-color: #0e1117;
        color: #fafafa;
    }
    
    /* Ubah tampilan header judul dengan efek gradasi */
    h1 {
        background: -webkit-linear-gradient(#4facfe, #00f2fe);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-weight: 800;
        text-align: center;
        padding-bottom: 20px;
    }
    
    /* Bikin teks deskripsi di tengah */
    .stMarkdown p {
        text-align: center;
        font-size: 1.1rem;
        color: #b0bec5;
    }

    /* Percantik area kotak Uploader */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #4facfe !important;
        border-radius: 15px !important;
        background-color: rgba(79, 172, 254, 0.05) !important;
        padding: 30px !important;
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        background-color: rgba(79, 172, 254, 0.15) !important;
        border-color: #00f2fe !important;
    }
    
    /* Tampilan tabel dataframe yang lebih modern */
    [data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid #262730;
    }
    
    /* Ubah Tombol Download jadi Keren */
    div.stDownloadButton > button {
        background: linear-gradient(90deg, #00C9FF 0%, #92FE9D 100%);
        color: #000000 !important;
        border-radius: 30px;
        border: none;
        padding: 10px 25px;
        font-weight: 700;
        letter-spacing: 0.5px;
        box-shadow: 0 4px 15px rgba(0, 201, 255, 0.3);
        transition: all 0.3s ease;
        width: 100%;
        margin-top: 15px;
    }
    div.stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0, 201, 255, 0.5);
    }
    
    /* Sembunyiin menu Streamlit di pojok kanan & footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# Layout: Bikin 3 kolom biar uploader-nya nggak selebar layar (taruh di tengah)
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    st.title("📍 ngubah kmz jadi excel jing")
    st.write("Ubah kumpulan file KMZ/KML lu menjadi tabel Excel secara otomatis. Support Batch Processing (Banyak File).")
    st.markdown("---")
    
    uploaded_files = st.file_uploader("Seret dan Lepaskan file di sini (Support multiple files)", type=['kmz', 'kml'], accept_multiple_files=True)

# Fungsi inti untuk memproses satu file
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
            if name_tag:
                row_data['Nama Placemark'] = name_tag.text.strip()
                
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
                    if name and value:
                        row_data[name] = value.text.strip()
                for simple_data in extended_data.find_all('SimpleData'):
                    name = simple_data.get('name')
                    if name:
                        row_data[name] = simple_data.text.strip()
            
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
                                
                                if key.lower() in ['lat', 'latitude']:
                                    key = 'Latitude (Tabel)'
                                elif key.lower() in ['long', 'longitude']:
                                    key = 'Longitude (Tabel)'
                                    
                                row_data[key] = val
                                
            data_list.append(row_data)
        
        if not data_list:
            return None, f"⚠️ Tidak ada data Placemark di {file.name}"
            
        df = pd.DataFrame(data_list)
        cols = df.columns.tolist()
        front_cols = ['Nama Placemark', 'Latitude (Asli)', 'Longitude (Asli)']
        front_cols = [c for c in front_cols if c in cols]
        other_cols = [c for c in cols if c not in front_cols]
        df = df[front_cols + other_cols]
        
        return df, None
    except Exception as e:
        return None, f"❌ Error di {file.name}: {e}"

# Proses Data
if uploaded_files:
    st.markdown("### 📊 Hasil Proses")
    with st.spinner(f"Sedang meracik {len(uploaded_files)} file untuk lu..."):
        
        # JIKA CUMA 1 FILE
        if len(uploaded_files) == 1:
            file = uploaded_files[0]
            df, error = process_single_file(file)
            
            if error:
                st.error(error)
            else:
                st.success(f"✅ Mantap! Ketemu {len(df)} titik koordinat dari file **{file.name}**")
                
                # Biar tabelnya ga kepanjangan di layar, kita batasin tingginya
                st.dataframe(df, height=250, use_container_width=True)
                
                base_name = file.name.rsplit('.', 1)[0]
                export_name = f"{base_name}_extracted.xlsx"
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                excel_data = output.getvalue()
                
                # Taruh tombol di tengah
                col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
                with col_btn2:
                    st.download_button(
                        label=f"📥 DOWNLOAD {export_name.upper()}",
                        data=excel_data,
                        file_name=export_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
        # JIKA MULTIPLE FILE (Batch ZIP)
        else:
            zip_buffer = io.BytesIO()
            berhasil_count = 0
            gagal_count = 0
            
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
                if gagal_count > 0:
                    st.warning(f"⚠️ Ada {gagal_count} file yang gagal diekstrak.")
                
                col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
                with col_btn2:
                    st.download_button(
                        label="📦 DOWNLOAD SEMUA (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="Geodata_Batch_Export.zip",
                        mime="application/zip"
                    )
            else:
                st.error("❌ Waduh, semua filenya gagal diproses. cek lagi anjing bener ga file lu.")
