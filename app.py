import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF
from datetime import date
import tempfile
import requests
import csv
import os
from io import StringIO
import math

# --- CONFIGURACIÓN DE COLORES Y RUTAS ---
COLOR_VERDE = "#008f39"   
# Gris clásico estilo Gemini
COLOR_GEMINI_DARK = "#1e1f20" 
COLOR_TEXTO = "#e3e3e3"

# Rutas RAW de GitHub para carga directa
URL_LOGO_WEB = "https://raw.githubusercontent.com/Osfelp/IMAGENES-ELYON-/main/logo-elyon-GRIS.png"
URL_LOGO_PDF = "https://raw.githubusercontent.com/Osfelp/IMAGENES-ELYON-/main/logo-elyon.png"
URL_WATERMARK_PDF = "https://raw.githubusercontent.com/Osfelp/IMAGENES-ELYON-/main/logo-elyon-GRIS2%2B.png"

st.set_page_config(page_title="Formato de Cotización | Elyon", page_icon="🎬", layout="wide")

# --- DESCARGA RÁPIDA DE IMÁGENES (Para hacer el botón instantáneo) ---
@st.cache_data(show_spinner=False)
def obtener_imagen_bytes(url):
    try:
        req = requests.get(url)
        if req.status_code == 200:
            return req.content
    except:
        pass
    return None

# --- DISEÑO MEJORADO (CSS PREMIUM) ---
st.markdown(f"""
    <style>
    /* Fondo General Gris Gemini */
    .stApp {{ background-color: {COLOR_GEMINI_DARK}; color: {COLOR_TEXTO}; }}
    h1, h2, h3, h4, h5, label, .stMetric {{ color: {COLOR_VERDE} !important; font-family: 'Lexend', sans-serif; }}
    
    /* Inputs y Tablas */
    .stTextInput>div>div>input, .stNumberInput>div>div>input, .stDateInput>div>div>input {{
        background-color: #282a2c !important; color: white !important;
        border: 1px solid #444746 !important; border-radius: 8px !important;
        padding: 8px 12px !important; transition: all 0.3s ease;
    }}
    .stTextInput>div>div>input:focus, .stNumberInput>div>div>input:focus {{
        border-color: {COLOR_VERDE} !important; box-shadow: 0 0 5px rgba(0, 143, 57, 0.5) !important;
    }}
    [data-testid="stDataEditor"] {{ background-color: #282a2c; border: 1px solid #444746; border-radius: 10px; overflow: hidden; }}
    
    /* Base de los botones */
    [data-testid="stDownloadButton"] > button, .stButton>button {{
        color: white !important; font-weight: 800 !important; font-size: 1.05em !important;
        border-radius: 8px !important; border: none !important;
        width: 100%; height: 3.5em; transition: all 0.3s ease !important;
    }}
    
    /* COLORES INDIVIDUALES PARA LOS BOTONES DE DESCARGA */
    /* 1. COT PDF (Verde Elyon) */
    div[data-testid="column"]:nth-of-type(1) [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(135deg, #008f39 0%, #00b347 100%) !important;
        box-shadow: 0 4px 6px rgba(0, 143, 57, 0.2) !important;
    }}
    div[data-testid="column"]:nth-of-type(1) [data-testid="stDownloadButton"] > button:hover {{
        transform: translateY(-3px) !important; box-shadow: 0 6px 12px rgba(0, 143, 57, 0.4) !important;
        background: linear-gradient(135deg, #00b347 0%, #008f39 100%) !important;
    }}

    /* 2. COT EXCEL (Naranja/Dorado) */
    div[data-testid="column"]:nth-of-type(2) [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(135deg, #e67e22 0%, #f39c12 100%) !important;
        box-shadow: 0 4px 6px rgba(230, 126, 34, 0.2) !important;
    }}
    div[data-testid="column"]:nth-of-type(2) [data-testid="stDownloadButton"] > button:hover {{
        transform: translateY(-3px) !important; box-shadow: 0 6px 12px rgba(230, 126, 34, 0.4) !important;
        background: linear-gradient(135deg, #f39c12 0%, #e67e22 100%) !important;
    }}

    /* 3. CUADRO CONTROL (Azul) */
    div[data-testid="column"]:nth-of-type(3) [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(135deg, #2980b9 0%, #3498db 100%) !important;
        box-shadow: 0 4px 6px rgba(41, 128, 185, 0.2) !important;
    }}
    div[data-testid="column"]:nth-of-type(3) [data-testid="stDownloadButton"] > button:hover {{
        transform: translateY(-3px) !important; box-shadow: 0 6px 12px rgba(41, 128, 185, 0.4) !important;
        background: linear-gradient(135deg, #3498db 0%, #2980b9 100%) !important;
    }}

    /* Botones deshabilitados (Gris) */
    [data-testid="stDownloadButton"] > button:disabled {{
        background: #2a2a2c !important; color: #666666 !important;
        box-shadow: none !important; transform: none !important; border: 1px solid #444746 !important;
    }}
    
    /* Botones generales de la app (Como el + de añadir sección) */
    .stButton>button {{ background: linear-gradient(135deg, #008f39 0%, #00b347 100%) !important; }}
    .stButton>button:hover {{ background: linear-gradient(135deg, #00b347 0%, #008f39 100%) !important; }}
    
    .btn-delete>button {{ background: #ff4b4b !important; box-shadow: none !important; height: auto; }}
    .btn-delete>button:hover {{ background: #ff0000 !important; transform: scale(1.05) !important; }}
    
    /* Tarjetas y Encabezados */
    hr {{ border: 0; height: 1px; background: #444746; margin: 25px 0; }}
    .group-header {{
        background: linear-gradient(90deg, #2a2a2c 0%, #1e1f20 100%);
        padding: 10px 15px; border-radius: 6px;
        border-left: 5px solid {COLOR_VERDE}; margin-top: 20px;
        color: white; font-weight: 700; font-size: 1.1em;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }}
    .compact-title {{ margin-top: 10px; margin-bottom: 10px; color: {COLOR_VERDE}; font-weight: 800; font-size: 1.2em; letter-spacing: 1px; }}
    
    /* Cajas de Resultados (Metrics) */
    div[data-testid="metric-container"] {{
        background-color: #282a2c; border: 1px solid #444746;
        padding: 15px 20px; border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center;
    }}
    </style>
""", unsafe_allow_html=True)

# --- FUNCIONES MATEMÁTICAS ---
def formato_moneda(valor):
    try:
        return f"$ {int(round(float(valor))):,}".replace(",", ".")
    except:
        return "$ 0"

def redondear_al_mil(valor):
    try:
        if pd.isna(valor): return 0
        v = float(valor)
        if v <= 0: return 0
        return int(math.ceil(v / 1000.0) * 1000)
    except:
        return 0

# --- CLASE PDF PERSONALIZADA (PIE DE PÁGINA) ---
class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, 'by Oscar Buitrago / Elyon Producciones.', align='C')

# --- ENCABEZADO ---
col_l, col_t = st.columns([1, 2.5])
with col_l:
    try:
        st.image(URL_LOGO_WEB, width=320)
    except:
        st.write("ELYON")
with col_t:
    st.markdown("<h1 style='margin-top: 10px; font-size: 3em;'>Formato de cotización</h1>", unsafe_allow_html=True)
    st.write("Plataforma de Gestión Comercial | **ELYON PRODUCCIONES**")

st.markdown("<hr>", unsafe_allow_html=True)

# --- I. INFORMACIÓN GENERAL ---
st.markdown("<div class='compact-title'>I. DATOS GENERALES</div>", unsafe_allow_html=True)

c1, c2, c3, c4, c5, c6, c7 = st.columns([1.5, 1.5, 1.5, 1, 0.8, 0.8, 1])
cotizacion_num = c1.text_input("COTIZACIÓN NO.", value="ELY-", key="gen_cot_num")
ciudad = c2.text_input("CIUDAD", key="gen_ciudad")
fecha_emision = c3.date_input("FECHA DE EMISIÓN", date.today(), key="gen_fecha_emision")
validez_dias = c4.number_input("VALIDEZ (DÍAS)", min_value=1, value=15, key="gen_validez")

with c7:
    st.markdown("<div style='margin-top: 32px;'></div>", unsafe_allow_html=True)
    sin_fee = st.checkbox("Sin Fee", key="sin_fee_chk")

fee_porcentaje_input = c5.number_input("FEE %", min_value=0, value=10, disabled=sin_fee, key="gen_fee")
fee_pct_final = 0 if sin_fee else fee_porcentaje_input

iva_porcentaje = c6.number_input("IVA %", min_value=0, value=19, key="gen_iva") 

st.markdown("<div class='compact-title'>DATOS DEL SOLICITANTE</div>", unsafe_allow_html=True)
c8, c9, c10, c11 = st.columns(4)
solicitante_empresa = c8.text_input("EMPRESA", key="sol_empresa")
solicitante_contacto = c9.text_input("CONTACTO", key="sol_contacto")
solicitante_email = c10.text_input("EMAIL", key="sol_email")
solicitante_telefono = c11.text_input("TELÉFONO", key="sol_telefono")

st.markdown("<div class='compact-title'>DETALLE EVENTO</div>", unsafe_allow_html=True)
c12, c13, c14, c15 = st.columns(4)
evento_referencia = c12.text_input("REFERENCIA", key="ev_referencia")
evento_fecha = c13.date_input("FECHA EVENTO", date.today(), key="ev_fecha") 
evento_lugar = c14.text_input("LUGAR EVENTO", key="ev_lugar")
evento_asistentes = c15.text_input("No. ASISTENTES", key="ev_asistentes")

st.markdown("<hr>", unsafe_allow_html=True)

# --- II. GESTIÓN POR CATEGORÍAS ---
st.markdown("<div class='compact-title'>II. DETALLE DEL SERVICIO (POR SECCIONES)</div>", unsafe_allow_html=True)

if 'categorias' not in st.session_state:
    st.session_state.categorias = {
        "General": pd.DataFrame([{"Descripción": "", "Días": 1.0, "Cantidad": 1.0, "Costo Unitario": 0.0, "Incremento (%)": 40.0}])
    }

col_nueva1, col_nueva2 = st.columns([10, 1])
with col_nueva1:
    nueva_cat = st.text_input("Añadir nueva sección...", label_visibility="collapsed", placeholder="Nombre de la nueva sección...", key="nueva_cat_input")
with col_nueva2:
    if st.button("➕"):
        if nueva_cat and nueva_cat not in st.session_state.categorias:
            st.session_state.categorias[nueva_cat] = pd.DataFrame([{"Descripción": "", "Días": 1.0, "Cantidad": 1.0, "Costo Unitario": 0.0, "Incremento (%)": 40.0}])
            st.rerun()

df_global = pd.DataFrame()

if not st.session_state.categorias:
    st.info("👆 Crea una sección para comenzar.")
else:
    for cat_name in list(st.session_state.categorias.keys()):
        col_tit, col_del = st.columns([8, 1])
        with col_tit:
            st.markdown(f"#### 📁 {cat_name.upper()}")
        with col_del:
            if cat_name != "General":
                st.markdown('<div class="btn-delete">', unsafe_allow_html=True)
                if st.button(f"🗑️", key=f"del_{cat_name}"):
                    del st.session_state.categorias[cat_name]
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
        
        if cat_name in st.session_state.categorias:
            df_cat = st.session_state.categorias[cat_name]
            config_cols = {
                "Descripción": st.column_config.TextColumn("Descripción", width="large"),
                "Días": st.column_config.NumberColumn("Días", width="small", min_value=0.0, step=1.0, default=1.0),
                "Cantidad": st.column_config.NumberColumn("Cant.", width="small", min_value=0.0, step=1.0, default=1.0),
                "Costo Unitario": st.column_config.NumberColumn("Costo Unitario ($)", width="medium", step=1.0, default=0.0), 
                "Incremento (%)": st.column_config.NumberColumn("Incr. (%)", width="small", step=1.0, default=40.0)
            }
            edited_df = st.data_editor(df_cat, num_rows="dynamic", use_container_width=True, hide_index=True, key=f"editor_{cat_name}", column_config=config_cols)
            
            df_calc = edited_df.copy()
            df_calc["Días"] = pd.to_numeric(df_calc["Días"], errors='coerce').fillna(1.0)
            df_calc["Cantidad"] = pd.to_numeric(df_calc["Cantidad"], errors='coerce').fillna(1.0)
            df_calc["Costo Unitario"] = pd.to_numeric(df_calc["Costo Unitario"], errors='coerce').fillna(0.0)
            df_calc["Incremento (%)"] = pd.to_numeric(df_calc["Incremento (%)"], errors='coerce').fillna(40.0)
            df_calc["Incremento (%)"] = df_calc["Incremento (%)"].apply(lambda x: 99.9 if x >= 100 else x)
            
            precio_base = df_calc["Costo Unitario"] / (1 - (df_calc["Incremento (%)"] / 100))
            df_calc["Precio Unitario"] = (np.ceil(precio_base / 1000.0) * 1000).fillna(0).astype(int)
            
            subtotal_base = df_calc["Días"] * df_calc["Cantidad"] * df_calc["Precio Unitario"]
            df_calc["Subtotal Item"] = (np.ceil(subtotal_base / 1000.0) * 1000).fillna(0).astype(int)
            
            df_calc["Categoría"] = cat_name 
            df_global = pd.concat([df_global, df_calc], ignore_index=True)

st.markdown("<hr>", unsafe_allow_html=True)

# --- III. VISTA PREVIA ---
st.markdown("<div class='compact-title'>III. VISTA PREVIA DE VALORES</div>", unsafe_allow_html=True)

subtotal_neto = 0
if not df_global.empty:
    df_validos = df_global[df_global["Descripción"].str.strip() != ""]
    for cat_name in st.session_state.categorias.keys():
        subset = df_validos[df_validos["Categoría"] == cat_name]
        if not subset.empty:
            st.markdown(f'<div class="group-header">{cat_name.upper()}</div>', unsafe_allow_html=True)
            subset_format = subset[["Descripción", "Días", "Cantidad", "Precio Unitario", "Subtotal Item"]].copy()
            subset_format["Precio Unitario"] = subset_format["Precio Unitario"].apply(formato_moneda)
            subset_format["Subtotal Item"] = subset_format["Subtotal Item"].apply(formato_moneda)
            st.dataframe(subset_format, use_container_width=True, hide_index=True)
            
            sub_cat_mostrar = redondear_al_mil(subset['Subtotal Item'].sum())
            st.write(f"**Subtotal {cat_name}:** {formato_moneda(sub_cat_mostrar)}")

    subtotal_neto = redondear_al_mil(df_validos["Subtotal Item"].sum() if not df_validos.empty else 0)

# FEE, IVA Y TOTAL AL MIL MÁS CERCANO
fee_total = redondear_al_mil(subtotal_neto * (fee_pct_final / 100.0))
iva_total = redondear_al_mil((subtotal_neto + fee_total) * (iva_porcentaje / 100.0))
total_general = redondear_al_mil(subtotal_neto + fee_total + iva_total)

st.markdown("<br>", unsafe_allow_html=True)
res1, res2, res3, res4 = st.columns(4)
res1.metric("SUBTOTAL (NETO)", formato_moneda(subtotal_neto))
res2.metric(f"FEE {fee_pct_final}%", formato_moneda(fee_total))
res3.metric(f"IVA {iva_porcentaje}%", formato_moneda(iva_total))
res4.metric("TOTAL GENERAL", formato_moneda(total_general))

# --- MOTOR DE PDF RÁPIDO ---
def generar_pdf():
    pdf = PDF(format='letter')
    pdf.add_page()
    
    tmp_wm_path, tmp_logo_path = None, None
    wm_bytes = obtener_imagen_bytes(URL_WATERMARK_PDF)
    logo_bytes = obtener_imagen_bytes(URL_LOGO_PDF)

    if wm_bytes:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_wm:
            tmp_wm.write(wm_bytes)
            tmp_wm_path = tmp_wm.name
        pdf.image(tmp_wm_path, x=35, y=85, w=145)

    if logo_bytes:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo:
            tmp_logo.write(logo_bytes)
            tmp_logo_path = tmp_logo.name
        pdf.image(tmp_logo_path, x=10, y=10, w=60)

    pdf.set_font("Arial", 'B', 16)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, "COTIZACIÓN COMERCIAL", ln=True, align="R")
    
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 5, f"Referencia: {cotizacion_num}", ln=True, align="R")
    pdf.cell(0, 5, f"Fecha: {fecha_emision.strftime('%d/%m/%Y')}", ln=True, align="R")
    pdf.ln(15)

    pdf.set_fill_color(0, 143, 57)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 8, " INFORMACIÓN GENERAL", ln=True, fill=True)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", '', 10)
    pdf.cell(100, 8, f" Cliente/Empresa: {solicitante_empresa}", border='LT')
    pdf.cell(90, 8, f" Contacto: {solicitante_contacto}", border='RT', ln=True)
    pdf.cell(100, 8, f" Ciudad: {ciudad}", border='L')
    pdf.cell(90, 8, f" Email: {solicitante_email}", border='R', ln=True)
    pdf.cell(190, 8, f" Teléfono: {solicitante_telefono}", border='LRB', ln=True)
    
    pdf.ln(5)
    pdf.set_fill_color(30, 30, 30)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 8, " DETALLE DEL EVENTO", ln=True, fill=True)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", '', 10)
    pdf.cell(100, 8, f" Referencia Evento: {evento_referencia}", border='LR')
    fecha_evento_str = evento_fecha.strftime('%d/%m/%Y') if hasattr(evento_fecha, 'strftime') else str(evento_fecha)
    pdf.cell(90, 8, f" Fecha Evento: {fecha_evento_str}", border='R', ln=True)
    pdf.cell(100, 8, f" Lugar Evento: {evento_lugar}", border='LRB')
    pdf.cell(90, 8, f" No. Asistentes: {evento_asistentes}", border='RB', ln=True)
    
    categorias_usadas = df_validos["Categoría"].unique()
    cat_ordenadas = ["General"] + [c for c in categorias_usadas if c != "General"]
    
    for cat in cat_ordenadas:
        if cat in categorias_usadas:
            pdf.ln(8)
            pdf.set_fill_color(230, 230, 230)
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, 8, f" SECCIÓN: {cat.upper()}", ln=True, fill=True, border=1)
            
            pdf.set_fill_color(30, 30, 30) 
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", 'B', 8)
            pdf.cell(75, 8, "DESCRIPCIÓN", 1, 0, 'L', True)
            pdf.cell(15, 8, "DIAS", 1, 0, 'C', True)
            pdf.cell(20, 8, "CANT.", 1, 0, 'C', True)
            pdf.cell(40, 8, "UNITARIO", 1, 0, 'C', True)
            pdf.cell(40, 8, "SUBTOTAL", 1, 1, 'C', True)

            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", '', 8)
            subset_cat = df_validos[df_validos["Categoría"] == cat]
            
            for _, row in subset_cat.iterrows():
                pdf.cell(75, 8, str(row["Descripción"]), 1)
                pdf.cell(15, 8, str(int(row["Días"])), 1, 0, 'C')
                pdf.cell(20, 8, str(int(row["Cantidad"])), 1, 0, 'C')
                pdf.cell(40, 8, formato_moneda(row['Precio Unitario']), 1, 0, 'R')
                pdf.cell(40, 8, formato_moneda(row['Subtotal Item']), 1, 1, 'R')
                
            pdf.set_font("Arial", 'B', 9)
            pdf.cell(150, 8, f"Subtotal {cat}:", 1, 0, 'R')
            pdf.cell(40, 8, formato_moneda(redondear_al_mil(subset_cat['Subtotal Item'].sum())), 1, 1, 'R')

    pdf.ln(10)
    pdf.set_font("Arial", 'B', 11)
    
    pdf.set_fill_color(230, 230, 230)
    pdf.cell(150, 8, "SUBTOTAL", 1, 0, 'R', fill=True)
    pdf.cell(40, 8, formato_moneda(subtotal_neto), 1, 1, 'R')
    
    pdf.cell(150, 8, f"FEE {fee_pct_final}%", 1, 0, 'R', fill=True)
    pdf.cell(40, 8, formato_moneda(fee_total), 1, 1, 'R')

    pdf.cell(150, 8, f"IVA {iva_porcentaje}%", 1, 0, 'R', fill=True)
    pdf.cell(40, 8, formato_moneda(iva_total), 1, 1, 'R')
    
    pdf.set_font("Arial", 'B', 12)
    pdf.set_text_color(255, 255, 255)
    pdf.set_fill_color(0, 143, 57) 
    pdf.cell(150, 10, "VALOR TOTAL NETO (ANTES DE IVA)", 1, 0, 'R', fill=True)
    pdf.cell(40, 10, formato_moneda(redondear_al_mil(subtotal_neto + fee_total)), 1, 1, 'R', fill=True)

    pdf.ln(3)
    pdf.cell(150, 10, "TOTAL GENERAL", 1, 0, 'R', fill=True)
    pdf.cell(40, 10, formato_moneda(total_general), 1, 1, 'R', fill=True)

    pdf.ln(10)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", 'B', 7)
    pdf.cell(0, 6, "TÉRMINOS Y CONDICIONES GENERALES DEL SERVICIO", ln=True)
    pdf.set_font("Arial", '', 6)
    
    texto_terminos = (
        "1.  Alcance del servicio: La cotización incluye únicamente lo descrito en la propuesta. Servicios o requerimientos adicionales serán cotizados y facturados por separado.\n"
        "2.  Cambios o ajustes: Cualquier modificación técnica, conceptual, logística o de alcance solicitada por el cliente generará costos adicionales.\n"
        "3.  Orden de Compra y ejecución: El proyecto se iniciará únicamente con la recepción de la Orden de Compra (OC) y el anticipo acordado. Cualquier retraso en estos procesos extenderá los tiempos de entrega y libera a ELYON PRODUCCIONES S.A.S. de compromisos de fecha.\n"
        "4.  Suspensión o cancelación: En caso de suspensión o cancelación del proyecto, el cliente deberá asumir la totalidad de los costos ejecutados, materiales adquiridos, gastos administrativos y penalidades de proveedores, si aplican.\n"
        "5.  Forma de pago y mora: El plazo de pago es de hasta 30 días calendario a partir de la fecha de emisión de la factura, salvo acuerdo escrito. El no pago oportuno generará intereses de mora a la tasa máxima legal vigente en Colombia, más los costos administrativos y de cobranza, hasta la cancelación total de la obligación.\n"
        "6.  Retrasos atribuibles al cliente: Retrasos en aprobaciones, entrega de información, artes, documentos o pagos afectarán el cronograma sin responsabilidad para ELYON PRODUCCIONES S.A.S.\n"
        "7.  Validez de la cotización: La cotización tiene una vigencia de 15 días calendario. Vencido este plazo, los valores podrán ser ajustados.\n"
        "8.  Aceptación y mérito ejecutivo: La aceptación de la cotización por cualquier medio implica la aceptación total de estos términos y condiciones y el reconocimiento expreso de la obligación de pago. La cotización aprobada y la factura correspondiente prestan mérito ejecutivo para efectos de cobro judicial conforme a la ley."
    )
    pdf.multi_cell(0, 3, texto_terminos)

    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    temp_pdf_path = temp_pdf.name
    temp_pdf.close()

    pdf.output(temp_pdf_path)
    
    if tmp_wm_path and os.path.exists(tmp_wm_path): os.remove(tmp_wm_path)
    if tmp_logo_path and os.path.exists(tmp_logo_path): os.remove(tmp_logo_path)
        
    with open(temp_pdf_path, "rb") as f:
        pdf_bytes = f.read()
    os.remove(temp_pdf_path)
    
    return pdf_bytes

# --- IV. EXPORTAR ---
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<div class='compact-title'>IV. EXPORTAR DOCUMENTOS</div>", unsafe_allow_html=True)

listo_para_descarga = not df_global.empty and not df_global["Descripción"].str.strip().eq("").all() and solicitante_empresa and solicitante_contacto

if not listo_para_descarga:
    st.info("⚠️ Completa los datos de **Empresa**, **Contacto** y añade al menos un **Ítem** para habilitar los botones de descarga.")

col_btn1, col_btn2, col_btn3 = st.columns(3)

pdf_data, csv_cot_data, csv_ctrl_data = b"", b"", b""

if listo_para_descarga:
    df_validos = df_global[df_global["Descripción"].str.strip() != ""]
    
    pdf_data = generar_pdf()

    def generar_csv_estructurado(df_data, es_control=False):
        output = StringIO()
        writer = csv.writer(output, delimiter=';')
        
        if es_control:
            writer.writerow(["CUADRO DE CONTROL INTERNO"])
            writer.writerow(["Evento:", evento_referencia])
            writer.writerow([])
            
            if not df_data.empty:
                categorias = df_data["Categoría"].unique()
                cat_ordenadas = ["General"] + [c for c in categorias if c != "General"]
                total_costo_general = 0

                for cat in cat_ordenadas:
                    if cat in categorias:
                        subset = df_data[df_data["Categoría"] == cat]
                        writer.writerow([f"SECCIÓN: {cat.upper()}"])
                        writer.writerow(["DESCRIPCIÓN", "VALOR UNITARIO", "VALOR TOTAL"])
                        
                        sub_cat = 0
                        for _, row in subset.iterrows():
                            desc = str(row["Descripción"]).strip()
                            dias = int(row["Días"]) if row["Días"].is_integer() else row["Días"]
                            cant = int(row["Cantidad"]) if row["Cantidad"].is_integer() else row["Cantidad"]
                            
                            unit = int(round(row["Costo Unitario"]))
                            sub_item = int(round(dias * cant * unit))
                            sub_cat += sub_item
                            writer.writerow([desc, unit, sub_item])
                        
                        total_costo_general += sub_cat
                        writer.writerow(["", f"Subtotal {cat}:", sub_cat])
                        writer.writerow([])
                writer.writerow(["", "COSTO TOTAL (NETO)", int(total_costo_general)])
        else:
            writer.writerow(["", "", "", "COTIZACIÓN COMERCIAL"])
            writer.writerow(["", "", "", "Referencia:", cotizacion_num])
            fecha_str = fecha_emision.strftime('%d/%m/%Y') if hasattr(fecha_emision, 'strftime') else str(fecha_emision)
            writer.writerow(["", "", "", "Fecha:", fecha_str])
            writer.writerow([])
            
            writer.writerow(["INFORMACIÓN GENERAL"])
            writer.writerow(["Cliente/Empresa:", solicitante_empresa, "Contacto:", solicitante_contacto])
            writer.writerow(["Ciudad:", ciudad, "Email:", solicitante_email])
            writer.writerow(["Teléfono:", solicitante_telefono])
            writer.writerow([])
            
            writer.writerow(["DETALLE DEL EVENTO"])
            evento_fecha_str = evento_fecha.strftime('%d/%m/%Y') if hasattr(evento_fecha, 'strftime') else str(evento_fecha)
            writer.writerow(["Referencia Evento:", evento_referencia, "Fecha Evento:", evento_fecha_str])
            writer.writerow(["Lugar Evento:", evento_lugar, "No. Asistentes:", evento_asistentes])
            writer.writerow([])

            if not df_data.empty:
                categorias = df_data["Categoría"].unique()
                cat_ordenadas = ["General"] + [c for c in categorias if c != "General"]

                for cat in cat_ordenadas:
                    if cat in categorias:
                        subset = df_data[df_data["Categoría"] == cat]
                        writer.writerow([f"SECCIÓN: {cat.upper()}"])
                        writer.writerow(["DESCRIPCIÓN", "DIAS", "CANT.", "PRECIO UNITARIO", "SUBTOTAL"])
                        
                        sub_cat = 0
                        for _, row in subset.iterrows():
                            desc = str(row["Descripción"]).strip()
                            dias = int(row["Días"]) if row["Días"].is_integer() else row["Días"]
                            cant = int(row["Cantidad"]) if row["Cantidad"].is_integer() else row["Cantidad"]
                            unit = int(row["Precio Unitario"])
                            sub_item = int(row["Subtotal Item"])
                            sub_cat += sub_item
                            writer.writerow([desc, dias, cant, unit, sub_item])
                        
                        writer.writerow(["", "", "", f"Subtotal {cat}:", sub_cat])
                        writer.writerow([])

                writer.writerow(["", "", "", "SUBTOTAL", int(subtotal_neto)])
                writer.writerow(["", "", "", f"FEE {int(fee_pct_final)}%", int(fee_total)])
                writer.writerow(["", "", "", f"IVA {int(iva_porcentaje)}%", int(iva_total)])
                writer.writerow(["", "", "", "VALOR TOTAL NETO", int(redondear_al_mil(subtotal_neto + fee_total))])
                writer.writerow(["", "", "", "TOTAL GENERAL", int(total_general)])

        return output.getvalue().encode('utf-8-sig')

    csv_cot_data = generar_csv_estructurado(df_validos, es_control=False)
    csv_ctrl_data = generar_csv_estructurado(df_validos, es_control=True)

# --- BOTONES DE DESCARGA DIRECTA (CON COLORES APLICADOS VÍA CSS) ---
with col_btn1:
    st.download_button(
        label="📥 COT PDF",
        data=pdf_data,
        file_name=f"{cotizacion_num}_{solicitante_empresa}.pdf",
        mime="application/pdf",
        use_container_width=True,
        disabled=not listo_para_descarga
    )

with col_btn2:
    st.download_button(
        label="📊 COT EXCEL",
        data=csv_cot_data,
        file_name=f"{cotizacion_num}_{solicitante_empresa}_Cotizacion.csv",
        mime="text/csv",
        use_container_width=True,
        disabled=not listo_para_descarga
    )

with col_btn3:
    st.download_button(
        label="📈 Exportar CUADRO CONTROL",
        data=csv_ctrl_data,
        file_name=f"Control_{cotizacion_num}_{solicitante_empresa}.csv",
        mime="text/csv",
        use_container_width=True,
        disabled=not listo_para_descarga
    )