import streamlit as st
import zipfile
import pandas as pd
from bs4 import BeautifulSoup
import io

# Konfigurasi halaman web
st.set_page_config(page_title="KMZ to Excel Converter", page_icon="🌍", layout="centered")

st.title("🌍 KMZ to Excel Converter")
st.write("Upload file KMZ di bawah ini. Sistem akan otomatis membaca tabel deskripsi dan mengekstrak koordinat asli (Tikor) langsung dari geometry KMZ.")

# Widget untuk upload file
uploaded_file = st.file_uploader("Pilih file .kmz", type=['kmz'])

if uploaded_file is not None:
    with st.spinner("Sedang memproses data..."):
        try:
            # Baca KMZ langsung dari memori (tanpa save ke disk)
            kmz = zipfile.ZipFile(uploaded_file, 'r')
            
            # Cari file KML di dalam zip KMZ
            kml_filename = [name for name in kmz.namelist() if name.endswith('.kml')][0]
            kml_data = kmz.open(kml_filename, 'r').read()
            
            # Parsing KML
            soup = BeautifulSoup(kml_data, 'xml')
            data_list = []
            
            placemarks = soup.find_all('Placemark')
            for placemark in placemarks:
                row_data = {}
                
                name_tag = placemark.find('name')
                if name_tag:
                    row_data['Nama Placemark'] = name_tag.text.strip()
                    
                # Ambil Koordinat Asli dari Geometry KMZ
                point_tag = placemark.find('Point')
                if point_tag:
                    coords_tag = point_tag.find('coordinates')
                    if coords_tag:
                        coords = coords_tag.text.strip().split(',')
                        if len(coords) >= 2:
                            row_data['Longitude (Asli KMZ)'] = float(coords[0].strip())
                            row_data['Latitude (Asli KMZ)'] = float(coords[1].strip())
                
                # Parsing Tabel HTML di Deskripsi (Pop-up)
                desc_tag = placemark.find('description')
                if desc_tag and desc_tag.text:
                    html_soup = BeautifulSoup(desc_tag.text, 'html.parser')
                    tables = html_soup.find_all('table')
                    
                    if tables:
                        for tr in tables[0].find_all('tr'):
                            tds = tr.find_all('td')
                            if len(tds) >= 2:
                                key = tds[0].text.strip()
                                val = tds[1].text.strip()
                                # Hindari duplikat nama kolom dengan Tikor asli
                                if key.lower() in ['lat', 'latitude']:
                                    key = 'Latitude (Tabel)'
                                elif key.lower() in ['long', 'longitude']:
                                    key = 'Longitude (Tabel)'
                                row_data[key] = val
                                
                data_list.append(row_data)
            
            if not data_list:
                st.warning("Tidak ada data Placemark yang ditemukan di dalam KMZ ini.")
            else:
                # Convert ke DataFrame
                df = pd.DataFrame(data_list)
                
                # Rapikan urutan kolom
                cols = df.columns.tolist()
                front_cols = ['Nama Placemark', 'Latitude (Asli KMZ)', 'Longitude (Asli KMZ)']
                front_cols = [c for c in front_cols if c in cols]
                other_cols = [c for c in cols if c not in front_cols]
                df = df[front_cols + other_cols]
                
                st.success(f"Sukses! Berhasil mengekstrak {len(df)} titik ODP.")
                
                # Tampilkan preview data di web
                st.dataframe(df.head())
                
                # Proses pembuatan file Excel di memori
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                excel_data = output.getvalue()
                
                # Tombol Download
                st.download_button(
                    label="📥 Download Data ke Excel",
                    data=excel_data,
                    file_name="Hasil_Ekstrak_KMZ.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Terjadi error saat memproses file: {e}")