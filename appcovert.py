import streamlit as st
import zipfile
import pandas as pd
from bs4 import BeautifulSoup
import io

st.set_page_config(page_title="KML/KMZ to Excel", page_icon="🌍", layout="centered")

st.title("🌍 KML/KMZ to Excel Converter")
st.write("Upload file KMZ atau KML di bawah ini. Sistem akan membaca koordinat asli (Tikor) dan mengekstrak tabel pop-up HTML menjadi kolom Excel.")

uploaded_file = st.file_uploader("Pilih file", type=['kmz', 'kml'])

if uploaded_file is not None:
    with st.spinner("Sedang memproses data..."):
        try:
            filename = uploaded_file.name.lower()
            kml_data = None
            
            if filename.endswith('.kmz'):
                kmz = zipfile.ZipFile(uploaded_file, 'r')
                kml_filename = [name for name in kmz.namelist() if name.endswith('.kml')][0]
                kml_data = kmz.open(kml_filename, 'r').read()
            elif filename.endswith('.kml'):
                kml_data = uploaded_file.read()
            
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
                st.warning("Tidak ada data Placemark/Titik yang ditemukan di dalam file ini.")
            else:
                df = pd.DataFrame(data_list)
                
                cols = df.columns.tolist()
                front_cols = ['Nama Placemark', 'Latitude (Asli)', 'Longitude (Asli)']
                front_cols = [c for c in front_cols if c in cols]
                other_cols = [c for c in cols if c not in front_cols]
                df = df[front_cols + other_cols]
                
                st.success(f"Sukses! Berhasil mengekstrak {len(df)} titik.")
                st.dataframe(df.head())
                
                # --- LOGIK AUTO RENAME FILE ---
                # Ambil nama asli, pisah berdasarkan titik dari belakang, ambil nama depannya aja
                base_name = uploaded_file.name.rsplit('.', 1)[0]
                export_filename = f"{base_name}_export_excel.xlsx"
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                excel_data = output.getvalue()
                
                st.download_button(
                    label=f"📥 Download {export_filename}",
                    data=excel_data,
                    file_name=export_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Error woy, cek filenya: {e}")
