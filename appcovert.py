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
    st.title("📍 Ubah KMZ Loh Jdi EXcel jing!!!!")
    st.markdown("<p style='text-align:center;'>Ubah kumpulan file KMZ/KML lu menjadi tabel Excel secara otomatis. Support Batch Processing.</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    uploaded_files = st.file_uploader("Seret dan Lepaskan file di sini (Support multiple files)", type=['kmz', 'kml'], accept_multiple_files
