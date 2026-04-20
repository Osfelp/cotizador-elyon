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
        if req.status_code == 200:
            return req.content
    except:
        pass
    return None

# --- DISEÑO (CSS PREMIUM) ---
st.markdown(f"""
    <style>
    .stApp {{ background-color: {COLOR_GEMINI_DARK}; color: {COLOR_TEXTO}; }}
    h1, h2, h3, h4, h5, label, .stMetric {{ color: {COLOR_VERDE} !important; font-family: 'Lexend', sans-serif; }}
    .stTextInput>div>div>input, .stNumberInput>div>div>input, .stDateInput>div>div>input {{
        background-color: #282a2c !important; color: white !important;
        border: 1px solid #444746 !important; border-radius: 8px !important;
        padding: 8px 12px !important;
    }}
    [data-testid="stDataEditor"] {{ background-color: #282a2c; border: 1px solid #444746; border-radius: 10px; }}
    
    [data-testid="stDownloadButton"] > button, .stButton>button {{
        color: white !important; font-weight: 800 !important; border-radius: 8px !important; border: none !important;
        width: 100%; height: 3.5em; transition: all 0.3s ease !important;
    }}
    
    /* 1. COT PDF */
    div[data-testid="column"]:nth-of-type(1) [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(135deg, #008f39 0%, #006829 100%) !important;
    }}
    /* 2. COT EXCEL */
    div[data-testid="column"]:nth-of-type(2) [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(135deg, #d35400 0%, #e67e22 100%) !important;
    }}
    /* 3. CUADRO CONTROL */
    div[data-testid="column"]:nth-of-type(3) [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(135deg, #2980b9 0%, #1f618d 100%) !important;
    }}
    
    .btn-delete>button {{ background: #ff4b4b !important; height: auto !important; }}
    .group-header {{
        background: linear-gradient(90deg, #2a2a2c 0%, #1e1f20 100%);
        padding: 10px 15px; border-radius: 6px; border-left: 5px solid {COLOR_VERDE};
        margin-top: 20px; color: white; font-weight: 700;
    }}
    div[data-testid="metric-container"] {{
        background-color: #282a2c; border: 1px solid #444746; padding: 15px 20px; border-radius: 10px;
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

# --- CLASE PDF ---
class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(128, 128, 128)

# --- ENCABEZADO ---
col_l, col_t = st.columns([1, 2.5])
with col_l:
    try: st.image(URL_LOGO_WEB, width=320)
    except: st.write("ELYON")
with col_t:
    st.markdown("<h1 style='margin-top: 10px; font-size: 3em;'>Formato de cotización</h1>", unsafe_allow_html=True)
    st.write("Plataforma de Gestión Comercial | **ELYON PRODUCCIONES**")

st.markdown("<hr>", unsafe_allow_html=True)

# --- I. DATOS GENERALES ---
st.markdown("<div class='compact-title'>I. DATOS GENERALES</div>", unsafe_allow_html=True)
c1, c2, c3, c4, c5, c6, c7 = st.columns([1.5, 1.5, 1.5, 1, 0.8, 0.8, 1])
cotizacion_num = c1.text_input("COTIZACIÓN NO.", value="ELY-")
ciudad = c2.text_input("CIUDAD")
fecha_emision = c3.date_input("FECHA DE EMISIÓN", date.today())
validez_dias = c4.number_input("VALIDEZ (DÍAS)", min_value=1, value=15)
with c7:
    st.markdown("<div style='margin-top: 32px;'></div>", unsafe_allow_html=True)
    sin_fee = st.checkbox("Sin Fee")
fee_porcentaje_input = c5.number_input("FEE %", min_value=0, value=10, disabled=sin_fee)
fee_pct_final = 0 if sin_fee else fee_porcentaje_input
iva_porcentaje = c6.number_input("IVA %", min_value=0, value=19) 

c8, c9, c10, c11 = st.columns(4)
solicitante_empresa = c8.text_input("EMPRESA")
solicitante_contacto = c9.text_input("CONTACTO")
solicitante_email = c10.text_input("EMAIL")
solicitante_telefono = c11.text_input("TELÉFONO")

c12, c13, c14, c15 = st.columns(4)
evento_referencia = c12.text_input("REFERENCIA")
evento_fecha = c13.date_input("FECHA EVENTO", date.today()) 
evento_lugar = c14.text_input("LUGAR EVENTO")
evento_asistentes = c15.text_input("No. ASISTENTES")

# --- II. GESTIÓN POR CATEGORÍAS ---
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<div class='compact-title'>II. DETALLE DEL SERVICIO</div>", unsafe_allow_html=True)

if 'categorias' not in st.session_state:
    st.session_state.categorias = {"General": pd.DataFrame([{"Descripción": "", "Días": 1.0, "Cantidad": 1.0, "Costo Unitario": 0.0, "Incremento (%)": 40.0}])}

col_nueva1, col_nueva2 = st.columns([10, 1])
nueva_cat = col_nueva1.text_input("Añadir nueva sección...", label_visibility="collapsed", placeholder="Nombre de la nueva sección...")
if col_nueva2.button("➕"):
    if nueva_cat and nueva_cat not in st.session_state.categorias:
        st.session_state.categorias[nueva_cat] = pd.DataFrame([{"Descripción": "", "Días": 1.0, "Cantidad": 1.0, "Costo Unitario": 0.0, "Incremento (%)": 40.0}])
        st.rerun()

df_global = pd.DataFrame()
for cat_name in list(st.session_state.categorias.keys()):
    col_tit, col_del = st.columns([8, 1])
    col_tit.markdown(f"#### 📁 {cat_name.upper()}")
    if cat_name != "General":
        if col_del.button(f"🗑️", key=f"del_{cat_name}"):
            del st.session_state.categorias[cat_name]
            st.rerun()
    
    df_cat = st.session_state.categorias[cat_name]
    edited_df = st.data_editor(df_cat, num_rows="dynamic", use_container_width=True, hide_index=True, key=f"ed_{cat_name}")
    
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

# --- III. VISTA PREVIA Y TOTALES ---
st.markdown("<hr>", unsafe_allow_html=True)
df_validos = df_global[df_global["Descripción"].str.strip() != ""] if not df_global.empty else pd.DataFrame()
subtotal_neto = redondear_al_mil(df_validos["Subtotal Item"].sum()) if not df_validos.empty else 0
fee_total = redondear_al_mil(subtotal_neto * (fee_pct_final / 100.0))
iva_total = redondear_al_mil((subtotal_neto + fee_total) * (iva_porcentaje / 100.0))
total_general = redondear_al_mil(subtotal_neto + fee_total + iva_total)

res1, res2, res3, res4 = st.columns(4)
res1.metric("SUBTOTAL (NETO)", formato_moneda(subtotal_neto))
res2.metric(f"FEE {fee_pct_final}%", formato_moneda(fee_total))
res3.metric(f"IVA {iva_porcentaje}%", formato_moneda(iva_total))
res4.metric("TOTAL GENERAL", formato_moneda(total_general))

# --- GENERADOR DE EXCEL PROFESIONAL (openpyxl) ---
def generar_excel_estilizado(df_data, es_control=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"
    
    # Estilos
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="008f39" if not es_control else "2980b9", end_color="008f39" if not es_control else "2980b9", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    currency_format = '$ #,##0'
    
    curr_row = 1
    
    # Encabezado General
    ws.cell(curr_row, 1, "CUADRO DE CONTROL INTERNO" if es_control else "COTIZACIÓN COMERCIAL").font = Font(bold=True, size=14)
    curr_row += 1
    ws.cell(curr_row, 1, f"Evento: {evento_referencia}")
    curr_row += 1
    ws.cell(curr_row, 1, f"Fecha Emisión: {fecha_emision.strftime('%d/%m/%Y')}")
    curr_row += 2
    
    if not df_data.empty:
        categorias = df_data["Categoría"].unique()
        for cat in categorias:
            # Título de Sección
            ws.cell(curr_row, 1, f"SECCIÓN: {cat.upper()}").font = Font(bold=True)
            curr_row += 1
            
            # Encabezados de tabla
            if es_control:
                headers = ["DESCRIPCIÓN", "DIAS", "CANT.", "COSTO UNIT.", "COSTO TOTAL", "PRECIO VENTA", "GANANCIA"]
            else:
                headers = ["DESCRIPCIÓN", "DIAS", "CANT.", "PRECIO UNITARIO", "SUBTOTAL"]
            
            for col_idx, text in enumerate(headers, 1):
                cell = ws.cell(curr_row, col_idx, text)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
                cell.alignment = Alignment(horizontal="center")
            
            curr_row += 1
            subset = df_data[df_data["Categoría"] == cat]
            
            for _, row in subset.iterrows():
                dias, cant = float(row["Días"]), float(row["Cantidad"])
                
                if es_control:
                    costo_u = float(row["Costo Unitario"])
                    costo_t = dias * cant * costo_u
                    venta_t = float(row["Subtotal Item"])
                    vals = [row["Descripción"], dias, cant, costo_u, costo_t, venta_t, venta_t - costo_t]
                else:
                    vals = [row["Descripción"], dias, cant, float(row["Precio Unitario"]), float(row["Subtotal Item"])]
                
                for col_idx, val in enumerate(headers, 1):
                    cell = ws.cell(curr_row, col_idx, vals[col_idx-1])
                    cell.border = border
                    # Aplicar formato moneda a columnas de dinero
                    if "UNIT" in headers[col_idx-1] or "TOTAL" in headers[col_idx-1] or "VENTA" in headers[col_idx-1] or "GANANCIA" in headers[col_idx-1] or "SUBTOTAL" in headers[col_idx-1] or "PRECIO" in headers[col_idx-1]:
                        cell.number_format = currency_format
                curr_row += 1
            curr_row += 1

    # Totales Finales (Solo para Cotización)
    if not es_control:
        resumen = [("SUBTOTAL", subtotal_neto), (f"FEE {fee_pct_final}%", fee_total), (f"IVA {iva_porcentaje}%", iva_total), ("TOTAL GENERAL", total_general)]
        for label, value in resumen:
            ws.cell(curr_row, 4, label).font = Font(bold=True)
            cell_val = ws.cell(curr_row, 5, value)
            cell_val.number_format = currency_format
            cell_val.font = Font(bold=True)
            curr_row += 1

    # Ajuste de ancho de columnas
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        ws.column_dimensions[column].width = max_length + 5

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- MOTOR PDF ---
def generar_pdf():
    pdf = PDF(format='letter')
    pdf.add_page()
    wm_b, logo_b = obtener_imagen_bytes(URL_WATERMARK_PDF), obtener_imagen_bytes(URL_LOGO_PDF)
    if wm_b:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as t:
            t.write(wm_b); t_p = t.name
        pdf.image(t_p, x=68, y=80, w=80); os.remove(t_p)
    if logo_b:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as t:
            t.write(logo_b); t_p = t.name
        pdf.image(t_p, x=10, y=10, w=60); os.remove(t_p)

    pdf.set_font("Arial", 'B', 16); pdf.cell(190, 10, "COTIZACIÓN COMERCIAL", ln=True, align="R")
    pdf.set_font("Arial", '', 10); pdf.cell(190, 5, f"Referencia: {cotizacion_num}", ln=True, align="R")
    pdf.cell(190, 5, f"Fecha: {fecha_emision.strftime('%d/%m/%Y')}", ln=True, align="R"); pdf.ln(15)

    pdf.set_fill_color(0, 143, 57); pdf.set_text_color(255, 255, 255); pdf.set_font("Arial", 'B', 11)
    pdf.cell(190, 8, " INFORMACIÓN GENERAL", ln=True, fill=True)
    pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 10)
    pdf.cell(100, 8, f" Cliente/Empresa: {solicitante_empresa}", border='LT')
    pdf.cell(90, 8, f" Contacto: {solicitante_contacto}", border='RT', ln=True)
    pdf.cell(100, 8, f" Ciudad: {ciudad}", border='L')
    pdf.cell(90, 8, f" Email: {solicitante_email}", border='R', ln=True)
    pdf.cell(190, 8, f" Teléfono: {solicitante_telefono}", border='LRB', ln=True); pdf.ln(5)

    pdf.set_fill_color(30, 30, 30); pdf.set_text_color(255, 255, 255)
    pdf.cell(190, 8, " DETALLE DEL EVENTO", ln=True, fill=True)
    pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 10)
    pdf.cell(100, 8, f" Referencia: {evento_referencia}", border='LR')
    pdf.cell(90, 8, f" Fecha: {evento_fecha.strftime('%d/%m/%Y')}", border='R', ln=True)
    pdf.cell(100, 8, f" Lugar: {evento_lugar}", border='LRB')
    pdf.cell(90, 8, f" Asistentes: {evento_asistentes}", border='RB', ln=True)

    for cat in df_validos["Categoría"].unique():
        pdf.ln(8); pdf.set_fill_color(230, 230, 230); pdf.set_font("Arial", 'B', 10)
        pdf.cell(190, 8, f" SECCIÓN: {cat.upper()}", ln=True, fill=True, border=1)
        pdf.set_fill_color(30, 30, 30); pdf.set_text_color(255, 255, 255); pdf.set_font("Arial", 'B', 8)
        pdf.cell(75, 8, "DESCRIPCIÓN", 1, 0, 'L', True); pdf.cell(15, 8, "DIAS", 1, 0, 'C', True)
        pdf.cell(20, 8, "CANT.", 1, 0, 'C', True); pdf.cell(40, 8, "UNITARIO", 1, 0, 'C', True)
        pdf.cell(40, 8, "SUBTOTAL", 1, 1, 'C', True); pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 8)
        sub_c = df_validos[df_validos["Categoría"] == cat]
        for _, r in sub_c.iterrows():
            pdf.cell(75, 8, str(r["Descripción"]), 1); pdf.cell(15, 8, str(int(r["Días"])), 1, 0, 'C')
            pdf.cell(20, 8, str(int(r["Cantidad"])), 1, 0, 'C'); pdf.cell(40, 8, formato_moneda(r['Precio Unitario']), 1, 0, 'R')
            pdf.cell(40, 8, formato_moneda(r['Subtotal Item']), 1, 1, 'R')

    pdf.ln(10); pdf.set_font("Arial", 'B', 11); pdf.set_fill_color(230, 230, 230)
    pdf.cell(150, 8, "SUBTOTAL", 1, 0, 'R', fill=True); pdf.cell(40, 8, formato_moneda(subtotal_neto), 1, 1, 'R')
    pdf.cell(150, 8, f"FEE {fee_pct_final}%", 1, 0, 'R', fill=True); pdf.cell(40, 8, formato_moneda(fee_total), 1, 1, 'R')
    pdf.cell(150, 8, f"IVA {iva_porcentaje}%", 1, 0, 'R', fill=True); pdf.cell(40, 8, formato_moneda(iva_total), 1, 1, 'R')
    pdf.set_text_color(255, 255, 255); pdf.set_fill_color(0, 143, 57); pdf.set_font("Arial", 'B', 12)
    pdf.cell(150, 10, "TOTAL GENERAL", 1, 0, 'R', fill=True); pdf.cell(40, 10, formato_moneda(total_general), 1, 1, 'R', fill=True)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(tmp.name); tmp.close()
    with open(tmp.name, "rb") as f: b = f.read()
    os.remove(tmp.name); return b

# --- BOTONES DE DESCARGA ---
st.markdown("<hr>", unsafe_allow_html=True)
listo = not df_validos.empty and solicitante_empresa
col_b1, col_b2, col_b3 = st.columns(3)

if listo:
    xlsx_cot = generar_excel_estilizado(df_validos, False)
    xlsx_ctrl = generar_excel_estilizado(df_validos, True)
    pdf_file = generar_pdf()

    col_b1.download_button("📥 COT PDF", pdf_file, f"Cot_{cotizacion_num}.pdf", "application/pdf")
    col_b2.download_button("📊 COT EXCEL", xlsx_cot, f"Cot_{cotizacion_num}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    col_b3.download_button("📈 CUADRO CONTROL", xlsx_ctrl, f"Control_{cotizacion_num}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("⚠️ Completa los datos para habilitar descargas.")

st.markdown("<div style='text-align: center; font-size: 10px; color: #1e1f20; padding-top: 60px;'>by Oscar Buitrago / Elyon Producciones.</div>", unsafe_allow_html=True)