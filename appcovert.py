import streamlit as st
import zipfile
import pandas as pd
from bs4 import BeautifulSoup
import io

st.set_page_config(page_title="KML/KMZ to Excel", page_icon="🌍", layout="centered")

st.title("🌍 KML/KMZ to Excel Converter (Batch Support)")
st.write("Upload file KMZ/KML di bawah ini. Lu bisa upload **banyak file sekaligus** (tinggal block/drag banyak file). Jika lebih dari 1 file, hasil otomatis di-ZIP.")

# Fungsi inti untuk memproses satu file (biar kodenya nggak berulang)
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
            return None, f"Tidak ada data Placemark di {file.name}"
            
        df = pd.DataFrame(data_list)
        cols = df.columns.tolist()
        front_cols = ['Nama Placemark', 'Latitude (Asli)', 'Longitude (Asli)']
        front_cols = [c for c in front_cols if c in cols]
        other_cols = [c for c in cols if c not in front_cols]
        df = df[front_cols + other_cols]
        
        return df, None
    except Exception as e:
        return None, f"Error di {file.name}: {e}"

# --- Uploader File (Ditambah accept_multiple_files=True) ---
uploaded_files = st.file_uploader("Pilih file (Bisa pilih/blok lebih dari 1 file)", type=['kmz', 'kml'], accept_multiple_files=True)

if uploaded_files:
    with st.spinner(f"Sedang memproses {len(uploaded_files)} file..."):
        
        # ----------------------------------------------------
        # JIKA CUMA 1 FILE (Otomatis download Excel langsung)
        # ----------------------------------------------------
        if len(uploaded_files) == 1:
            file = uploaded_files[0]
            df, error = process_single_file(file)
            
            if error:
                st.error(error)
            else:
                st.success(f"Sukses! Berhasil mengekstrak {len(df)} titik dari file {file.name}")
                st.dataframe(df.head())
                
                base_name = file.name.rsplit('.', 1)[0]
                export_name = f"{base_name}_export_excel.xlsx"
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                excel_data = output.getvalue()
                
                st.download_button(
                    label=f"📥 Download {export_name}",
                    data=excel_data,
                    file_name=export_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        # ----------------------------------------------------
        # JIKA MULTIPLE FILE (Otomatis masukin ke ZIP)
        # ----------------------------------------------------
        else:
            zip_buffer = io.BytesIO()
            berhasil_count = 0
            gagal_count = 0
            
            # Buat file ZIP baru di memori
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for file in uploaded_files:
                    df, error = process_single_file(file)
                    
                    if error:
                        st.warning(error)
                        gagal_count += 1
                    else:
                        base_name = file.name.rsplit('.', 1)[0]
                        export_name = f"{base_name}_export_excel.xlsx"
                        
                        # Tulis dataframe ke Excel di memori, lalu masukin ke dalam ZIP
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False)
                        zip_file.writestr(export_name, excel_buffer.getvalue())
                        berhasil_count += 1
                        
            if berhasil_count > 0:
                st.success(f"Selesai diproses! {berhasil_count} file berhasil di-convert ke Excel (Tersimpan di dalam ZIP).")
                if gagal_count > 0:
                    st.warning(f"Ada {gagal_count} file yang gagal atau kosong.")
                
                # Tombol download ZIP
                st.download_button(
                    label="📦 Download Semua Hasil (Format ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Hasil_Export_KMZ_Batch.zip",
                    mime="application/zip"
                )
            else:
                st.error("Semua file gagal diproses. Pastikan file KMZ/KML valid.")
