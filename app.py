import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF
from datetime import date
import tempfile
import requests
import os
from io import BytesIO
import math
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill

# Librerías para Automatización (Google Drive y Correo)
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- CONFIGURACIÓN DE ASESORES ---
ASESORES = {
    "Oscar Buitrago": {"email": "oscar@ejemplo.com", "drive_folder": "TU_ID_AQUI"},
    "Administración": {"email": "admin@ejemplo.com", "drive_folder": "TU_ID_AQUI"}
}

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

# --- DISEÑO (CSS PREMIUM) ---
st.markdown(f"""
    <style>
    .stApp {{ background-color: {COLOR_GEMINI_DARK}; color: {COLOR_TEXTO}; }}
    h1, h2, h3, h4, h5, label, .stMetric {{ color: {COLOR_VERDE} !important; font-family: 'Lexend', sans-serif; }}
    .stTextInput>div>div>input, .stNumberInput>div>div>input, .stDateInput>div>div>input {{
        background-color: #282a2c !important; color: white !important;
        border: 1px solid #444746 !important; border-radius: 8px !important; padding: 8px 12px !important;
    }}
    [data-testid="stDataEditor"] {{ background-color: #282a2c; border: 1px solid #444746; border-radius: 10px; }}
    
    div[data-testid="column"]:nth-of-type(1) [data-testid="stDownloadButton"] > button {{ background: linear-gradient(135deg, #008f39 0%, #006829 100%) !important; color: white; height: 3.5em; font-weight: bold; border-radius: 8px; width: 100%; border: none; }}
    div[data-testid="column"]:nth-of-type(2) [data-testid="stDownloadButton"] > button {{ background: linear-gradient(135deg, #d35400 0%, #e67e22 100%) !important; color: white; height: 3.5em; font-weight: bold; border-radius: 8px; width: 100%; border: none; }}
    div[data-testid="column"]:nth-of-type(3) [data-testid="stDownloadButton"] > button {{ background: linear-gradient(135deg, #2980b9 0%, #1f618d 100%) !important; color: white; height: 3.5em; font-weight: bold; border-radius: 8px; width: 100%; border: none; }}
    
    .btn-edit>button {{ background-color: #3a3b3c !important; color: white !important; height: auto !important; margin-top: 15px; padding: 5px 15px; }}
    .btn-delete>button {{ background-color: #ff4b4b !important; color: white !important; height: auto !important; margin-top: 15px; padding: 5px 15px; }}
    .group-header {{ background: linear-gradient(90deg, #2a2a2c 0%, #1e1f20 100%); padding: 10px 15px; border-radius: 6px; border-left: 5px solid {COLOR_VERDE}; margin-top: 20px; color: white; font-weight: 700; font-size: 1.1em; }}
    .preview-box {{ background-color: #1e1f20; padding: 20px; border-radius: 10px; border: 1px solid #444746; margin-top: 10px; }}
    div[data-testid="metric-container"] {{ background-color: #282a2c; border: 1px solid #444746; padding: 15px 20px; border-radius: 10px; }}
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

# --- AUTOMATIZACIÓN PROTEGIDA ---
def enviar_correo_automatico(destinatario, nombre_cliente, pdf_bytes, xlsx_bytes, filename):
    if "gmail" not in st.secrets: return False # Sale sin error si no hay contraseñas
    try:
        msg = MIMEMultipart()
        msg['From'] = st.secrets["gmail"]["user"]
        msg['To'] = destinatario
        msg['Subject'] = f"Nueva Cotización: {nombre_cliente} - {filename}"
        
        msg.attach(MIMEText(f"Se ha generado la cotización {filename} para el cliente {nombre_cliente}.\nAdjunto PDF y Cuadro de Control.", 'plain'))
        
        part_pdf = MIMEBase('application', 'octet-stream')
        part_pdf.set_payload(pdf_bytes)
        encoders.encode_base64(part_pdf)
        part_pdf.add_header('Content-Disposition', f"attachment; filename={filename}.pdf")
        msg.attach(part_pdf)

        part_xl = MIMEBase('application', 'octet-stream')
        part_xl.set_payload(xlsx_bytes)
        encoders.encode_base64(part_xl)
        part_xl.add_header('Content-Disposition', f"attachment; filename={filename}_Control.xlsx")
        msg.attach(part_xl)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(st.secrets["gmail"]["user"], st.secrets["gmail"]["password"])
        server.send_message(msg)
        server.quit()
        return True
    except: return False

def subir_a_drive(archivo_bytes, nombre_archivo, folder_id, tipo_mimo):
    if "google_drive" not in st.secrets: return False
    try:
        creds = service_account.Credentials.from_service_account_info(st.secrets["google_drive"])
        service = build('drive', 'v3', credentials=creds)
        file_metadata = {'name': nombre_archivo, 'parents': [folder_id]}
        media = MediaIoBaseUpload(BytesIO(archivo_bytes), mimetype=tipo_mimo, resumable=True)
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return True
    except: return False

# --- CLASE PDF ---
class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.set_auto_page_break(auto=True, margin=15)
    def footer(self):
        self.set_y(-15); self.set_font('Arial', 'I', 8); self.set_text_color(128, 128, 128)

# --- ENCABEZADO ---
col_l, col_t = st.columns([1, 2.5])
with col_l:
    try: st.image(URL_LOGO_WEB, width=320)
    except: st.write("ELYON")
with col_t:
    st.markdown("<h1 style='margin-top: 10px; font-size: 3em;'>Formato de cotización</h1>", unsafe_allow_html=True)
    st.write("Plataforma de Gestión Comercial | **ELYON PRODUCCIONES**")
    asesor_selec = st.selectbox("👤 ASESOR COMERCIAL", list(ASESORES.keys()))

st.markdown("<hr>", unsafe_allow_html=True)

# --- I. DATOS GENERALES ---
st.markdown("### I. DATOS GENERALES")
c1, c2, c3, c4, c5, c6, c7 = st.columns([1.5, 1.5, 1.5, 1, 0.8, 0.8, 1])
cotizacion_num = c1.text_input("COTIZACIÓN NO.", value="ELY-", key="val_cot")
ciudad = c2.text_input("CIUDAD", key="val_ciu")
fecha_emision = c3.date_input("FECHA DE EMISIÓN", date.today(), key="val_fec")
validez_dias = c4.number_input("VALIDEZ (DÍAS)", min_value=1, value=15, key="val_val")
with c7:
    st.markdown("<div style='margin-top: 32px;'></div>", unsafe_allow_html=True)
    sin_fee = st.checkbox("Sin Fee", key="val_sinf")
fee_porcentaje_input = c5.number_input("FEE %", min_value=0, value=10, disabled=sin_fee, key="val_fee")
fee_pct_final = 0 if sin_fee else fee_porcentaje_input
iva_porcentaje = c6.number_input("IVA %", min_value=0, value=19, key="val_iva") 

c8, c9, c10, c11 = st.columns(4)
solicitante_empresa = c8.text_input("EMPRESA", key="val_emp")
solicitante_contacto = c9.text_input("CONTACTO", key="val_con")
solicitante_email = c10.text_input("EMAIL", placeholder="ejemplo@correo.com", key="val_ema")
solicitante_telefono = c11.text_input("TELÉFONO", placeholder="Ej: 3001234567", key="val_tel")

c12, c13, c14, c15 = st.columns(4)
evento_referencia = c12.text_input("REFERENCIA EVENTO", key="val_ref")
evento_fecha = c13.date_input("FECHA EVENTO", date.today(), key="val_fev") 
evento_lugar = c14.text_input("LUGAR EVENTO", key="val_lug")
evento_asistentes = c15.text_input("No. ASISTENTES", key="val_asi")

# --- II. GESTIÓN POR CATEGORÍAS (Motor Blindado) ---
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("### II. DETALLE DEL SERVICIO")

DEFAULT_VALS = {"Descripción": "", "Días": 1.0, "Cantidad": 1.0, "Costo Unitario": 0.0, "Incremento (%)": 40.0}

if 'cat_order' not in st.session_state: st.session_state.cat_order = ["General"]
if 'categorias' not in st.session_state: st.session_state.categorias = {"General": pd.DataFrame([DEFAULT_VALS])}
if 'edited_categorias' not in st.session_state: st.session_state.edited_categorias = {}

def agregar_nueva_seccion():
    n = st.session_state.nueva_cat_input.strip()
    if n and n not in st.session_state.cat_order:
        st.session_state.cat_order.append(n)
        st.session_state.categorias[n] = pd.DataFrame([DEFAULT_VALS])
    st.session_state.nueva_cat_input = ""

def guardar_edicion(old_name, idx):
    new_name = st.session_state[f"edit_input_{old_name}"].strip()
    if new_name and new_name != old_name and new_name not in st.session_state.categorias:
        baked_df = st.session_state.edited_categorias.get(old_name, st.session_state.categorias[old_name]).copy()
        st.session_state.categorias[new_name] = baked_df
        st.session_state.cat_order[idx] = new_name
        st.session_state.categorias.pop(old_name, None)
        st.session_state.edited_categorias.pop(old_name, None)
    st.session_state[f"edit_mode_{old_name}"] = False

col_nueva1, col_nueva2 = st.columns([10, 1])
col_nueva1.text_input("Añadir nueva sección...", label_visibility="collapsed", placeholder="Escribe un título y presiona Enter...", key="nueva_cat_input", on_change=agregar_nueva_seccion)
st.markdown("<br>", unsafe_allow_html=True)

df_global = pd.DataFrame()
config_columnas = {
    "Descripción": st.column_config.TextColumn("Descripción", width="large"),
    "Días": st.column_config.NumberColumn("Días", width="small", min_value=0.0, step=1.0, default=1.0),
    "Cantidad": st.column_config.NumberColumn("Cant.", width="small", min_value=0.0, step=1.0, default=1.0),
    "Costo Unitario": st.column_config.NumberColumn("Costo Unitario ($)", width="medium", step=1.0, default=0.0), 
    "Incremento (%)": st.column_config.NumberColumn("Incr. (%)", width="small", step=1.0, default=40.0)
}

for i, cat_name in enumerate(list(st.session_state.cat_order)):
    col_tit, col_edit, col_del = st.columns([8, 1, 1])
    edit_key = f"edit_mode_{cat_name}"
    
    if st.session_state.get(edit_key, False):
        with col_tit: st.text_input("Renombrar (Enter para guardar):", value=cat_name, key=f"edit_input_{cat_name}", on_change=guardar_edicion, args=(cat_name, i), label_visibility="collapsed")
    else:
        with col_tit: st.markdown(f"#### 📁 {cat_name.upper()}")
            
    with col_edit:
        if not st.session_state.get(edit_key, False):
            st.markdown('<div class="btn-edit">', unsafe_allow_html=True)
            if st.button("✏️", key=f"btn_edit_{cat_name}"):
                st.session_state[edit_key] = True; st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    with col_del:
        if cat_name != "General":
            st.markdown('<div class="btn-delete">', unsafe_allow_html=True)
            if st.button("🗑️", key=f"del_{cat_name}"):
                st.session_state.cat_order.pop(i)
                st.session_state.categorias.pop(cat_name, None)
                st.session_state.edited_categorias.pop(cat_name, None); st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
    
    df_base = st.session_state.categorias.get(cat_name, pd.DataFrame([DEFAULT_VALS]))
    edited_df = st.data_editor(df_base, num_rows="dynamic", use_container_width=True, hide_index=True, key=f"editor_{cat_name}", column_config=config_columnas)
    st.session_state.edited_categorias[cat_name] = edited_df.copy()
    
    df_calc = edited_df.copy()
    df_calc["Días"] = pd.to_numeric(df_calc["Días"], errors='coerce').fillna(1.0)
    df_calc["Cantidad"] = pd.to_numeric(df_calc["Cantidad"], errors='coerce').fillna(1.0)
    df_calc["Costo Unitario"] = pd.to_numeric(df_calc["Costo Unitario"], errors='coerce').fillna(0.0)
    df_calc["Incremento (%)"] = pd.to_numeric(df_calc["Incremento (%)"], errors='coerce').fillna(40.0)
    
    precio_base = df_calc["Costo Unitario"] / (1 - (df_calc["Incremento (%)"] / 100))
    df_calc["Precio Unitario"] = (np.ceil(precio_base / 1000.0) * 1000).fillna(0).astype(int)
    subtotal_base = df_calc["Días"] * df_calc["Cantidad"] * df_calc["Precio Unitario"]
    df_calc["Subtotal Item"] = (np.ceil(subtotal_base / 1000.0) * 1000).fillna(0).astype(int)
    df_calc["Categoría"] = cat_name 
    df_global = pd.concat([df_global, df_calc], ignore_index=True)

# --- CÁLCULOS GLOBALES ---
df_validos = df_global[df_global["Descripción"].str.strip() != ""] if not df_global.empty else pd.DataFrame()
subtotal_neto = redondear_al_mil(df_validos["Subtotal Item"].sum()) if not df_validos.empty else 0
fee_total = redondear_al_mil(subtotal_neto * (fee_pct_final / 100.0))
iva_total = redondear_al_mil((subtotal_neto + fee_total) * (iva_porcentaje / 100.0))
total_general = redondear_al_mil(subtotal_neto + fee_total + iva_total)

# --- III. PREVISUALIZACIÓN ---
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("### 📄 III. PREVISUALIZACIÓN DE COTIZACIÓN")
st.markdown('<div class="preview-box">', unsafe_allow_html=True)
st.markdown(f"**Empresa:** {solicitante_empresa} &nbsp;&nbsp;|&nbsp;&nbsp; **Referencia:** {evento_referencia} &nbsp;&nbsp;|&nbsp;&nbsp; **Fecha:** {evento_fecha.strftime('%d/%m/%Y')}")

if not df_validos.empty:
    for cat_name in st.session_state.cat_order:
        subset = df_validos[df_validos["Categoría"] == cat_name]
        if not subset.empty:
            st.markdown(f'<div class="group-header">{cat_name.upper()}</div>', unsafe_allow_html=True)
            subset_format = subset[["Descripción", "Días", "Cantidad", "Precio Unitario", "Subtotal Item"]].copy()
            subset_format["Precio Unitario"] = subset_format["Precio Unitario"].apply(formato_moneda)
            subset_format["Subtotal Item"] = subset_format["Subtotal Item"].apply(formato_moneda)
            st.dataframe(subset_format, use_container_width=True, hide_index=True)
            sub_cat_mostrar = redondear_al_mil(subset['Subtotal Item'].sum())