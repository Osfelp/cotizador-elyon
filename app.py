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

# Librerías para Automatización
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# --- CONFIGURACIÓN DE ASESORES ---
# Edita los correos y los IDs de carpeta de Drive aquí
ASESORES = {
    "Oscar Buitrago": {"email": "oscar@ejemplo.com", "drive_folder": "ID_CARPETA_OSCAR"},
    "María Pérez": {"email": "maria@ejemplo.com", "drive_folder": "ID_CARPETA_MARIA"},
    "Vendedor General": {"email": "administracion@ejemplo.com", "drive_folder": "ID_CARPETA_GENERAL"}
}

COLOR_VERDE = "#008f39"   
COLOR_GEMINI_DARK = "#1e1f20" 
COLOR_TEXTO = "#e3e3e3"

URL_LOGO_WEB = "https://raw.githubusercontent.com/Osfelp/IMAGENES-ELYON-/main/logo-elyon-GRIS.png"
URL_LOGO_PDF = "https://raw.githubusercontent.com/Osfelp/IMAGENES-ELYON-/main/logo-elyon.png"
URL_WATERMARK_PDF = "https://raw.githubusercontent.com/Osfelp/IMAGENES-ELYON-/main/logo-elyon-GRIS2%2B.png"

st.set_page_config(page_title="Gestión Elyon", page_icon="🎬", layout="wide")

# --- FUNCIONES DE AUTOMATIZACIÓN ---

def enviar_correo_automatico(destinatario, nombre_cliente, pdf_bytes, xlsx_bytes, filename):
    try:
        msg = MIMEMultipart()
        msg['From'] = st.secrets["gmail"]["user"]
        msg['To'] = destinatario
        msg['Subject'] = f"Nueva Cotización: {nombre_cliente} - {filename}"
        
        cuerpo = f"Hola, se ha generado una nueva cotización para {nombre_cliente}.\nAdjunto encontrarás el PDF comercial y el Cuadro de Control en Excel."
        msg.attach(MIMEText(cuerpo, 'plain'))

        # Adjuntar PDF
        part_pdf = MIMEBase('application', 'octet-stream')
        part_pdf.set_payload(pdf_bytes)
        encoders.encode_base64(part_pdf)
        part_pdf.add_header('Content-Disposition', f"attachment; filename= {filename}.pdf")
        msg.attach(part_pdf)

        # Adjuntar Excel
        part_xl = MIMEBase('application', 'octet-stream')
        part_xl.set_payload(xlsx_bytes)
        encoders.encode_base64(part_xl)
        part_xl.add_header('Content-Disposition', f"attachment; filename= {filename}_Control.xlsx")
        msg.attach(part_xl)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(st.secrets["gmail"]["user"], st.secrets["gmail"]["password"])
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Error enviando correo: {e}")
        return False

def subir_a_drive(archivo_bytes, nombre_archivo, folder_id, tipo_mimo):
    try:
        creds = service_account.Credentials.from_service_account_info(st.secrets["google_drive"])
        service = build('drive', 'v3', credentials=creds)
        file_metadata = {'name': nombre_archivo, 'parents': [folder_id]}
        media = MediaIoBaseUpload(BytesIO(archivo_bytes), mimetype=tipo_mimo, resumable=True)
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return True
    except Exception as e:
        st.error(f"Error Drive: {e}")
        return False

# --- UI Y LOGICA DE COTIZACIÓN ---

# CABEZOTE
col_l, col_t = st.columns([1, 2.5])
with col_l:
    try: st.image(URL_LOGO_WEB, width=320)
    except: st.write("ELYON")
with col_t:
    st.markdown("<h1 style='color:#008f39;'>Formato de cotización</h1>", unsafe_allow_html=True)
    asesor_selec = st.selectbox("👤 ASESOR COMERCIAL", list(ASESORES.keys()))

st.markdown("<hr>", unsafe_allow_html=True)

# --- MEMORIA Y TABLAS ---
if 'cat_order' not in st.session_state: st.session_state.cat_order = ["General"]
if 'categorias' not in st.session_state: st.session_state.categorias = {"General": pd.DataFrame([{"Descripción": "", "Días": 1.0, "Cantidad": 1.0, "Costo Unitario": 0.0, "Incremento (%)": 40.0}])}

# (Aquí va toda la lógica de Datos Generales y Detalle de Servicio que ya tenemos perfeccionada...)
# [Por brevedad mantengo la estructura del motor que ya te funcionó]

st.markdown("### I. DATOS GENERALES")
c1, c2, c3, c4, c5, c6, c7 = st.columns([1.5, 1.5, 1.5, 1, 0.8, 0.8, 1])
cotizacion_num = c1.text_input("COTIZACIÓN NO.", value="ELY-", key="val_cot")
solicitante_empresa = st.text_input("EMPRESA CLIENTE", key="val_emp")
# ... resto de inputs con sus keys ...

# --- III. PREVISUALIZACIÓN ---
# (Lógica de previsualización que ya tienes)

# --- IV. EXPORTAR Y AUTOMATIZAR ---
st.markdown("<hr>", unsafe_allow_html=True)
col_b1, col_b2, col_b3 = st.columns(3)

# Generación de datos
# pdf_data = generar_pdf() 
# xlsx_ctrl = generar_excel_estilizado(df_validos, True)

if col_b1.download_button("📥 COT PDF"):
    # Al descargar, se dispara la automatización
    info_asesor = ASESORES[asesor_selec]
    
    # 1. Enviar Correo
    with st.spinner("Enviando copia al asesor..."):
        enviar_correo_automatico(info_asesor["email"], solicitante_empresa, pdf_data, xlsx_ctrl, cotizacion_num)
    
    # 2. Subir a Drive
    with st.spinner("Guardando respaldo en Drive..."):
        subir_a_drive(pdf_data, f"{cotizacion_num}_{solicitante_empresa}.pdf", info_asesor["drive_folder"], "application/pdf")
    
    st.success(f"✅ Cotización procesada y enviada a {asesor_selec}")

# ... Resto del código de botones ...