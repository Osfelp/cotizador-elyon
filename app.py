import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF
from datetime import date, datetime
import tempfile
import requests
import os
from io import BytesIO
import math
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill

# --- CONFIGURACIÓN DE COLORES Y RUTAS ---
COLOR_VERDE = "#008f39"   
COLOR_GEMINI_DARK = "#1e1f20" 
COLOR_TEXTO = "#e3e3e3"

URL_LOGO_WEB = "https://raw.githubusercontent.com/Osfelp/IMAGENES-ELYON-/main/logo-elyon-GRIS.png"
URL_LOGO_PDF = "https://raw.githubusercontent.com/Osfelp/IMAGENES-ELYON-/main/logo-elyon.png"
URL_WATERMARK_PDF = "https://raw.githubusercontent.com/Osfelp/IMAGENES-ELYON-/main/logo-elyon-GRIS2%2B.png"

st.set_page_config(page_title="Formato de Cotización | Elyon", page_icon="🎬", layout="wide")

# --- DESCARGA RÁPIDA DE IMÁGENES ---
@st.cache_data(show_spinner=False)
def obtener_imagen_bytes(url):
    try:
        req = requests.get(url)
        if req.status_code == 200: return req.content
    except: pass
    return None

# --- DISEÑO (CSS PREMIUM UI/UX) ---
st.markdown(f"""
    <style>
    .stApp {{ background-color: {COLOR_GEMINI_DARK}; color: {COLOR_TEXTO}; }}
    h1, h2, h3, h4, h5, label, .stMetric {{ color: {COLOR_VERDE} !important; font-family: 'Lexend', sans-serif; }}
    
    .stTextInput>div>div>input, .stNumberInput>div>div>input, .stDateInput>div>div>input {{
        background-color: #161718 !important; color: white !important;
        border: 1px solid #3a3b3c !important; border-radius: 8px !important; padding: 8px 12px !important;
    }}
    [data-testid="stDataEditor"] {{ background-color: #161718; border: 1px solid #3a3b3c; border-radius: 10px; }}
    
    [data-testid="metric-container"] {{
        background: linear-gradient(145deg, #1e1f20 0%, #111213 100%);
        border: 1px solid #3a3b3c; padding: 20px 15px !important; border-radius: 12px !important;
        box-shadow: 0px 6px 15px rgba(0,0,0,0.3); text-align: center;
    }}
    [data-testid="metric-container"] label {{ justify-content: center; font-size: 0.95rem !important; opacity: 0.8; margin-bottom: 5px; }}
    [data-testid="metric-container"] div {{ justify-content: center; font-size: 2.1rem !important; color: white !important; font-weight: bold; }}

    div[data-testid="stHorizontalBlock"]:last-of-type [data-testid="column"]:nth-child(1) button {{
        background: linear-gradient(135deg, #008f39 0%, #00501f 100%) !important; 
        color: white !important; height: 3.8em !important; border-radius: 8px !important; border: none !important; font-weight: bold !important; box-shadow: 0 4px 10px rgba(0,0,0,0.4) !important;
    }}
    div[data-testid="stHorizontalBlock"]:last-of-type [data-testid="column"]:nth-child(2) button {{
        background: linear-gradient(135deg, #d35400 0%, #903a00 100%) !important; 
        color: white !important; height: 3.8em !important; border-radius: 8px !important; border: none !important; font-weight: bold !important; box-shadow: 0 4px 10px rgba(0,0,0,0.4) !important;
    }}
    div[data-testid="stHorizontalBlock"]:last-of-type [data-testid="column"]:nth-child(3) button {{
        background: linear-gradient(135deg, #2980b9 0%, #154360 100%) !important; 
        color: white !important; height: 3.8em !important; border-radius: 8px !important; border: none !important; font-weight: bold !important; box-shadow: 0 4px 10px rgba(0,0,0,0.4) !important;
    }}
    div[data-testid="stHorizontalBlock"]:last-of-type button:hover {{ transform: translateY(-2px); transition: 0.2s; }}

    .sidebar-card {{ background-color: #161718; border: 1px solid #3a3b3c; padding: 18px; border-radius: 12px; margin-bottom: 10px; box-shadow: 0px 4px 12px rgba(0,0,0,0.4); }}
    [data-testid="stSidebar"] [data-testid="stDownloadButton"] button {{ background: linear-gradient(135deg, #2980b9 0%, #154360 100%) !important; color: white !important; border: none !important; font-weight: bold !important; height: 3em !important; }}
    
    .btn-edit button {{ background-color: #2c3e50 !important; color: white !important; height: 36px !important; padding: 0px 15px !important; border-radius: 8px !important; border: none !important; margin-top: 13px !important; }}
    .btn-edit button:hover {{ background-color: #34495e !important; transform: scale(1.05); }}
    
    .btn-delete button {{ background-color: #7b241c !important; color: white !important; height: 36px !important; padding: 0px 15px !important; border-radius: 8px !important; border: none !important; margin-top: 13px !important; }}
    .btn-delete button:hover {{ background-color: #922b21 !important; transform: scale(1.05); }}
    
    .group-header {{ background: linear-gradient(90deg, #161718 0%, #1e1f20 100%); padding: 12px 20px; border-radius: 8px; border: 1px solid #3a3b3c; border-left: 5px solid {COLOR_VERDE}; margin-top: 20px; color: white; font-weight: 700; font-size: 1.1em; }}
    .preview-box {{ background-color: #161718; padding: 25px; border-radius: 12px; border: 1px solid #3a3b3c; margin-top: 15px; box-shadow: 0px 4px 10px rgba(0,0,0,0.2); }}
    </style>
""", unsafe_allow_html=True)

# --- FUNCIONES MATEMÁTICAS ---
def formato_moneda(valor):
    try: return f"$ {int(round(float(valor))):,}".replace(",", ".")
    except: return "$ 0"

def redondear_al_mil(valor):
    try:
        if pd.isna(valor): return 0
        v = float(valor); return 0 if v <= 0 else int(math.ceil(v / 1000.0) * 1000)
    except: return 0

# --- CLASE PDF
