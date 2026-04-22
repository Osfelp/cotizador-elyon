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
    
    /* Botones de Descarga (Grandes) */
    div[data-testid="column"]:nth-of-type(1) [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(135deg, #008f39 0%, #006829 100%) !important; color: white; height: 3.5em; font-weight: bold; border-radius: 8px; width: 100%; border: none;
    }}
    div[data-testid="column"]:nth-of-type(2) [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(135deg, #d35400 0%, #e67e22 100%) !important; color: white; height: 3.5em; font-weight: bold; border-radius: 8px; width: 100%; border: none;
    }}
    div[data-testid="column"]:nth-of-type(3) [data-testid="stDownloadButton"] > button {{
        background: linear-gradient(135deg, #2980b9 0%, #1f618d 100%) !important; color: white; height: 3.5em; font-weight: bold; border-radius: 8px; width: 100%; border: none;
    }}
    
    /* Botones de gestión (Lápiz y Borrar) */
    .btn-edit>button {{ background-color: #3a3b3c !important; color: white !important; height: auto !important; margin-top: 15px; padding: 5px 15px; }}
    .btn-edit>button:hover {{ background-color: #555555 !important; transform: scale(1.05) !important; }}
    .btn-delete>button {{ background-color: #ff4b4b !important; color: white !important; height: auto !important; margin-top: 15px; padding: 5px 15px; }}
    .btn-delete>button:hover {{ background-color: #ff0000 !important; transform: scale(1.05) !important; }}
    
    .group-header {{
        background: linear-gradient(90deg, #2a2a2c 0%, #1e1f20 100%);
        padding: 10px 15px; border-radius: 6px; border-left: 5px solid {COLOR_VERDE};
        margin-top: 20px; color: white; font-weight: 700; font-size: 1.1em;
    }}
    .preview-box {{
        background-color: #1e1f20; padding: 20px; border-radius: 10px; border: 1px solid #444746; margin-top: 10px;
    }}
    div[data-testid="metric-container"] {{
        background-color: #282a2c; border: 1px solid #444746; padding: 15px 20px; border-radius: 10px;
    }}
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

st.markdown("<hr>", unsafe_allow_html=True)

# --- I. DATOS GENERALES ---
st.markdown("### I. DATOS GENERALES")
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
solicitante_email = c10.text_input("EMAIL", placeholder="ejemplo@correo.com")
solicitante_telefono = c11.text_input("TELÉFONO", placeholder="Ej: 3001234567")

c12, c13, c14, c15 = st.columns(4)
evento_referencia = c12.text_input("REFERENCIA EVENTO")
evento_fecha = c13.date_input("FECHA EVENTO", date.today()) 
evento_lugar = c14.text_input("LUGAR EVENTO")
evento_asistentes = c15.text_input("No. ASISTENTES")

# --- II. GESTIÓN POR CATEGORÍAS ---
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("### II. DETALLE DEL SERVICIO")

DEFAULT_VALS = {"Descripción": "", "Días": 1.0, "Cantidad": 1.0, "Costo Unitario": 0.0, "Incremento (%)": 40.0}

if 'cat_order' not in st.session_state:
    st.session_state.cat_order = ["General"]
if 'categorias' not in st.session_state:
    st.session_state.categorias = {"General": pd.DataFrame([DEFAULT_VALS])}

def agregar_nueva_seccion():
    n = st.session_state.nueva_cat_input.strip()
    if n and n not in st.session_state.cat_order:
        st.session_state.cat_order.append(n)
        st.session_state.categorias[n] = pd.DataFrame([DEFAULT_VALS])
    st.session_state.nueva_cat_input = ""

def guardar_edicion(old_name, idx):
    new_name = st.session_state[f"edit_input_{old_name}"].strip()
    if new_name and new_name != old_name and new_name not in st.session_state.categorias:
        st.session_state.categorias[new_name] = st.session_state.categorias.pop(old_name)
        st.session_state.cat_order[idx] = new_name
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

# Mostrar secciones
for i, cat_name in enumerate(list(st.session_state.cat_order)):
    col_tit, col_edit, col_del = st.columns([8, 1, 1])
    
    edit_key = f"edit_mode_{cat_name}"
    
    if st.session_state.get(edit_key, False):
        with col_tit:
            st.text_input("Renombrar sección (Presiona Enter para guardar):", value=cat_name, key=f"edit_input_{cat_name}", on_change=guardar_edicion, args=(cat_name, i), label_visibility="collapsed")
    else:
        with col_tit:
            st.markdown(f"#### 📁 {cat_name.upper()}")
            
    with col_edit:
        if not st.session_state.get(edit_key, False):
            st.markdown('<div class="btn-edit">', unsafe_allow_html=True)
            if st.button("✏️", key=f"btn_edit_{cat_name}"):
                st.session_state[edit_key] = True
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    with col_del:
        if cat_name != "General":
            st.markdown('<div class="btn-delete">', unsafe_allow_html=True)
            if st.button("🗑️", key=f"del_{cat_name}"):
                st.session_state.cat_order.pop(i)
                st.session_state.categorias.pop(cat_name, None)
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
    
    # Tabla de datos
    df_cat = st.session_state.categorias.get(cat_name, pd.DataFrame([DEFAULT_VALS]))
    edited_df = st.data_editor(df_cat, num_rows="dynamic", use_container_width=True, hide_index=True, key=f"ed_{cat_name}", column_config=config_columnas)
    st.session_state.categorias[cat_name] = edited_df 
    
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

# --- III. PREVISUALIZACIÓN DE COTIZACIÓN ---
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
            st.markdown(f"<div style='text-align: right; color: {COLOR_VERDE}; font-weight: bold;'>Subtotal {cat_name}: {formato_moneda(sub_cat_mostrar)}</div>", unsafe_allow_html=True)
else:
    st.info("Agrega descripciones en las tablas de arriba para ver la previsualización.")

st.markdown("<br>", unsafe_allow_html=True)
res1, res2, res3, res4 = st.columns(4)
res1.metric("SUBTOTAL NETO", formato_moneda(subtotal_neto))
res2.metric(f"FEE {fee_pct_final}%", formato_moneda(fee_total))
res3.metric(f"IVA {iva_porcentaje}%", formato_moneda(iva_total))
res4.metric("TOTAL GENERAL", formato_moneda(total_general))
st.markdown('</div>', unsafe_allow_html=True)

# --- GENERADOR DE EXCEL PROFESIONAL ---
def generar_excel_estilizado(df_data, es_control=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="008f39" if not es_control else "2980b9", end_color="008f39" if not es_control else "2980b9", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    currency_format = '$ #,##0'
    
    curr_row = 1
    ws.cell(curr_row, 1, "CUADRO DE CONTROL INTERNO" if es_control else "COTIZACIÓN COMERCIAL").font = Font(bold=True, size=14)
    curr_row += 1
    ws.cell(curr_row, 1, f"Evento: {evento_referencia}")
    curr_row += 1
    ws.cell(curr_row, 1, f"Fecha Emisión: {fecha_emision.strftime('%d/%m/%Y')}")
    curr_row += 2
    
    ws.cell(curr_row, 1, "INFORMACIÓN GENERAL").font = Font(bold=True)
    curr_row += 1
    ws.cell(curr_row, 1, "Cliente/Empresa:").font = Font(bold=True); ws.cell(curr_row, 2, solicitante_empresa)
    ws.cell(curr_row, 3, "Contacto:").font = Font(bold=True); ws.cell(curr_row, 4, solicitante_contacto)
    curr_row += 1
    
    ws.cell(curr_row, 1, "Ciudad:").font = Font(bold=True); ws.cell(curr_row, 2, ciudad)
    ws.cell(curr_row, 3, "Email:").font = Font(bold=True)
    cell_email = ws.cell(curr_row, 4, solicitante_email)
    if solicitante_email:
        cell_email.hyperlink = f"mailto:{solicitante_email}"
        cell_email.font = Font(color="0000FF", underline="single")
    curr_row += 1
    
    ws.cell(curr_row, 1, "Teléfono:").font = Font(bold=True)
    try:
        cell_tel = ws.cell(curr_row, 2, int(solicitante_telefono.replace(" ", "").replace("+", "")))
        cell_tel.number_format = '0'
    except:
        ws.cell(curr_row, 2, solicitante_telefono)
    curr_row += 2
    
    total_costo_general, total_venta_general, total_ganancia_general = 0, 0, 0

    if not df_data.empty:
        for cat in st.session_state.cat_order:
            if cat not in df_data["Categoría"].values: continue
            
            ws.cell(curr_row, 1, f"SECCIÓN: {cat.upper()}").font = Font(bold=True)
            curr_row += 1
            
            headers = ["DESCRIPCIÓN", "DIAS", "CANT.", "COSTO UNIT.", "COSTO TOTAL", "PRECIO VENTA UNIT.", "PRECIO VENTA", "GANANCIA", "% GANANCIA"] if es_control else ["DESCRIPCIÓN", "DIAS", "CANT.", "PRECIO UNITARIO", "SUBTOTAL"]
            for col_idx, text in enumerate(headers, 1):
                cell = ws.cell(curr_row, col_idx, text)
                cell.font, cell.fill, cell.border, cell.alignment = header_font, header_fill, border, Alignment(horizontal="center")
            
            curr_row += 1
            subset = df_data[df_data["Categoría"] == cat]
            sub_costo, sub_venta, sub_gan = 0, 0, 0
            
            for _, row in subset.iterrows():
                d, c = float(row["Días"]), float(row["Cantidad"])
                if es_control:
                    cu = float(row["Costo Unitario"])
                    ct = d * c * cu
                    vu = float(row["Precio Unitario"]) 
                    vt = float(row["Subtotal Item"])   
                    gan = vt - ct
                    pct = (gan / vt) if vt > 0 else 0
                    vals = [row["Descripción"], d, c, cu, ct, vu, vt, gan, pct]
                    sub_costo += ct; sub_venta += vt; sub_gan += gan
                else:
                    vals = [row["Descripción"], d, c, float(row["Precio Unitario"]), float(row["Subtotal Item"])]
                
                for idx, v in enumerate(vals, 1):
                    cell = ws.cell(curr_row, idx, v); cell.border = border
                    h_name = headers[idx-1]
                    if es_control and h_name == "% GANANCIA":
                        cell.number_format = '0.0%'
                        if isinstance(v, (int, float)):
                            if v >= 0.35: cell.fill, cell.font = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), Font(color="006100")
                            else: cell.fill, cell.font = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"), Font(color="9C0006")
                    elif any(x in h_name for x in ["UNIT", "TOTAL", "VENTA", "GANANCIA", "SUBTOTAL", "PRECIO"]):
                        cell.number_format = currency_format
                curr_row += 1
            
            if es_control:
                sub_pct = (sub_gan / sub_venta) if sub_venta > 0 else 0
                vals_sub = ["", "", "", "Subtotales:", sub_costo, "", sub_venta, sub_gan, sub_pct]
                for idx, v in enumerate(vals_sub, 1):
                    cell = ws.cell(curr_row, idx, v)
                    if v != "":
                        cell.font, cell.border = Font(bold=True), border
                        if idx in [5, 7, 8]: cell.number_format = currency_format
                        elif idx == 9:
                            cell.number_format = '0.0%'
                            if v >= 0.35: cell.fill, cell.font = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), Font(bold=True, color="006100")
                            else: cell.fill, cell.font = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"), Font(bold=True, color="9C0006")
                total_costo_general += sub_costo; total_venta_general += sub_venta; total_ganancia_general += sub_gan
            curr_row += 1

    if es_control:
        tot_pct = (total_ganancia_general / total_venta_general) if total_venta_general > 0 else 0
        vals_tot = ["", "", "", "TOTALES GENERALES:", total_costo_general, "", total_venta_general, total_ganancia_general, tot_pct]
        for idx, v in enumerate(vals_tot, 1):
            cell = ws.cell(curr_row, idx, v)
            if v != "":
                cell.font, cell.border = Font(bold=True), border
                if idx in [5, 7, 8]: cell.number_format = currency_format
                elif idx == 9:
                    cell.number_format = '0.0%'
                    if v >= 0.35: cell.fill, cell.font = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), Font(bold=True, color="006100")
                    else: cell.fill, cell.font = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"), Font(bold=True, color="9C0006")
        curr_row += 1

    if not es_control:
        resumen = [("SUBTOTAL", subtotal_neto), (f"FEE {fee_pct_final}%", fee_total), (f"IVA {iva_porcentaje}%", iva_total), ("TOTAL GENERAL", total_general)]
        for label, val in resumen:
            ws.cell(curr_row, 4, label).font = Font(bold=True)
            cv = ws.cell(curr_row, 5, val); cv.number_format = currency_format; cv.font = Font(bold=True)
            curr_row += 1

    for col in ws.columns:
        max_len = 0; column = col[0].column_letter
        for cell in col:
            try: 
                if len(str(cell.value)) > max_len: max_len = len(str(cell.value))
            except: pass
        ws.column_dimensions[column].width = max_len + 5

    output = BytesIO(); wb.save(output); return output.getvalue()

# --- MOTOR PDF CON MANEJO DE TEXTOS LARGOS ---
def generar_pdf():
    pdf = PDF(format='letter'); pdf.add_page()
    wm_b, logo_b = obtener_imagen_bytes(URL_WATERMARK_PDF), obtener_imagen_bytes(URL_LOGO_PDF)
    if wm_b:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as t: t.write(wm_b); t_p = t.name
        pdf.image(t_p, x=68, y=100, w=80); os.remove(t_p)
    if logo_b:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as t: t.write(logo_b); t_p = t.name
        pdf.image(t_p, x=10, y=10, w=60); os.remove(t_p)

    pdf.set_font("Arial", 'B', 16); pdf.cell(190, 10, "COTIZACIÓN COMERCIAL", ln=True, align="R")
    pdf.set_font("Arial", '', 10); pdf.cell(190, 5, f"Referencia: {cotizacion_num}", ln=True, align="R")
    pdf.cell(190, 5, f"Fecha: {fecha_emision.strftime('%d/%m/%Y')}", ln=True, align="R"); pdf.ln(15)

    pdf.set_fill_color(0, 143, 57); pdf.set_text_color(255, 255, 255); pdf.set_font("Arial", 'B', 11)
    pdf.cell(190, 8, " INFORMACIÓN GENERAL", ln=True, fill=True)
    pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 10)
    pdf.cell(100, 8, f" Cliente: {solicitante_empresa}", border='LT'); pdf.cell(90, 8, f" Contacto: {solicitante_contacto}", border='RT', ln=True)
    pdf.cell(100, 8, f" Ciudad: {ciudad}", border='L'); pdf.cell(90, 8, f" Email: {solicitante_email}", border='R', ln=True)
    pdf.cell(190, 8, f" Teléfono: {solicitante_telefono}", border='LRB', ln=True); pdf.ln(5)

    pdf.set_fill_color(30, 30, 30); pdf.set_text_color(255, 255, 255)
    pdf.cell(190, 8, " DETALLE DEL EVENTO", ln=True, fill=True)
    pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 10)
    pdf.cell(100, 8, f" Referencia: {evento_referencia}", border='LR'); pdf.cell(90, 8, f" Fecha: {evento_fecha.strftime('%d/%m/%Y')}", border='R', ln=True)
    pdf.cell(100, 8, f" Lugar: {evento_lugar}", border='LRB'); pdf.cell(90, 8, f" Asistentes: {evento_asistentes}", border='RB', ln=True)

    for cat in st.session_state.cat_order:
        if cat not in df_validos["Categoría"].values: continue
        
        pdf.ln(8); pdf.set_fill_color(230, 230, 230); pdf.set_font("Arial", 'B', 10); pdf.cell(190, 8, f" SECCIÓN: {cat.upper()}", ln=True, fill=True, border=1)
        pdf.set_fill_color(30, 30, 30); pdf.set_text_color(255, 255, 255); pdf.set_font("Arial", 'B', 8)
        pdf.cell(75, 8, "DESCRIPCIÓN", 1, 0, 'L', True); pdf.cell(15, 8, "DIAS", 1, 0, 'C', True); pdf.cell(20, 8, "CANT.", 1, 0, 'C', True); pdf.cell(40, 8, "UNITARIO", 1, 0, 'C', True); pdf.cell(40, 8, "SUBTOTAL", 1, 1, 'C', True)
        pdf.set_text_color(0, 0, 0); pdf.set_font("Arial", '', 8)
        
        sc = df_validos[df_validos["Categoría"] == cat]
        for _, r in sc.iterrows():
            desc = str(r["Descripción"])
            
            ancho_columna = 73.0
            ancho_texto = pdf.get_string_width(desc)
            lineas_estimadas = math.ceil(ancho_texto / ancho_columna) if ancho_texto > 0 else 1
            lineas_estimadas = max(1, lineas_estimadas)
            altura_fila = max(8, lineas_estimadas * 5)
            
            if pdf.get_y() + altura_fila > 255:
                pdf.add_page()
            
            x_start = pdf.get_x()
            y_start = pdf.get_y()
            
            pdf.multi_cell(75, 5 if lineas_estimadas > 1 else 8, desc, 1, 'L')
            y_end = pdf.get_y()
            
            altura_real = max(8, y_end - y_start)
            
            pdf.set_xy(x_start + 75, y_start)
            pdf.cell(15, altura_real, str(int(r["Días"])), 1, 0, 'C')
            pdf.cell(20, altura_real, str(int(r["Cantidad"])), 1, 0, 'C')
            pdf.cell(40, altura_real, formato_moneda(r['Precio Unitario']), 1, 0, 'R')
            pdf.cell(40, altura_real, formato_moneda(r['Subtotal Item']), 1, 1, 'R')
            
            pdf.set_y(y_start + altura_real)

        pdf.set_font("Arial", 'B', 9)
        pdf.cell(150, 8, f"Subtotal {cat}:", 1, 0, 'R')
        pdf.cell(40, 8, formato_moneda(redondear_al_mil(sc['Subtotal Item'].sum())), 1, 1, 'R')

    pdf.ln(10); pdf.set_font("Arial", 'B', 11); pdf.set_fill_color(230, 230, 230)
    
    if pdf.get_y() + 40 > 255: pdf.add_page()
        
    pdf.cell(150, 8, "SUBTOTAL", 1, 0, 'R', fill=True); pdf.cell(40, 8, formato_moneda(subtotal_neto), 1, 1, 'R')
    pdf.cell(150, 8, f"FEE {fee_pct_final}%", 1, 0, 'R', fill=True); pdf.cell(40, 8, formato_moneda(fee_total), 1, 1, 'R')
    pdf.cell(150, 8, f"IVA {iva_porcentaje}%", 1, 0, 'R', fill=True); pdf.cell(40, 8, formato_moneda(iva_total), 1, 1, 'R')
    pdf.set_text_color(255, 255, 255); pdf.set_fill_color(0, 143, 57); pdf.set_font("Arial", 'B', 12)
    pdf.cell(150, 10, "VALOR TOTAL NETO (ANTES DE IVA)", 1, 0, 'R', fill=True); pdf.cell(40, 10, formato_moneda(redondear_al_mil(subtotal_neto + fee_total)), 1, 1, 'R', fill=True)

    pdf.ln(3)
    pdf.cell(150, 10, "TOTAL GENERAL", 1, 0, 'R', fill=True)
    pdf.cell(40, 10, formato_moneda(total_general), 1, 1, 'R', fill=True)

    pdf.ln(10)
    
    if pdf.get_y() + 30 > 260: pdf.add_page()
    
    pdf.set_text_color(100, 100, 100); pdf.set_font("Arial", 'B', 6)
    pdf.cell(0, 4, "TÉRMINOS Y CONDICIONES GENERALES DEL SERVICIO", ln=True); pdf.set_font("Arial", '', 6)
    texto_terminos = (
        "1. Alcance del servicio: La cotización incluye únicamente lo descrito en la propuesta. Servicios o requerimientos adicionales serán cotizados y facturados por separado.\n"
        "2. Cambios o ajustes: Cualquier modificación técnica, conceptual, logística o de alcance solicitada por el cliente generará costos adicionales.\n"
        "3. Orden de Compra y ejecución: El proyecto se iniciará únicamente con la recepción de la Orden de Compra (OC) y el anticipo acordado. Cualquier retraso en estos procesos extenderá los tiempos de entrega y libera a ELYON PRODUCCIONES S.A.S. de compromisos de fecha.\n"
        "4. Suspensión o cancelación: En caso de suspensión o cancelación del proyecto, el cliente deberá asumir la totalidad de los costos ejecutados, materiales adquiridos, gastos administrativos y penalidades de proveedores, si aplican.\n"
        "5. Forma de pago y mora: El plazo de pago es de hasta 30 días calendario a partir de la fecha de emisión de la factura, salvo acuerdo escrito. El no pago oportuno generará intereses de mora a la tasa máxima legal vigente en Colombia, más los costos administrativos y de cobranza, hasta la cancelación total de la obligación.\n"
        "6. Retrasos atribuibles al cliente: Retrasos en aprobaciones, entrega de información, artes, documentos o pagos afectarán el cronograma sin responsabilidad para ELYON PRODUCCIONES S.A.S.\n"
        "7. Validez de la cotización: La cotización tiene una vigencia de 15 días calendario. Vencido este plazo, los valores podrán ser ajustados.\n"
        "8. Aceptación y mérito ejecutivo: La aceptación de la cotización por cualquier medio implica la aceptación total de estos términos y condiciones y el reconocimiento expreso de la obligación de pago. La cotización aprobada y la factura correspondiente prestan mérito ejecutivo para efectos de cobro judicial conforme a la ley."
    )
    pdf.multi_cell(0, 3, texto_terminos)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf"); pdf.output(tmp.name); tmp.close()
    with open(tmp.name, "rb") as f: b = f.read()
    os.remove(tmp.name); return b

# --- IV. BOTONES DE DESCARGA ---
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("### IV. EXPORTAR DOCUMENTOS")

listo_para_descargar = not df_validos.empty
if not listo_para_descargar:
    st.info("⚠️ Agrega al menos un ítem a la tabla para habilitar las descargas.")

col_b1, col_b2, col_b3 = st.columns(3)
empresa_nombre = solicitante_empresa if solicitante_empresa else "Cliente"

pdf_data = generar_pdf() if listo_para_descargar else b""
xlsx_cot = generar_excel_estilizado(df_validos, False) if listo_para_descargar else b""
xlsx_ctrl = generar_excel_estilizado(df_validos, True) if listo_para_descargar else b""

col_b1.download_button("📥 COT PDF", pdf_data, f"Cot_{cotizacion_num}_{empresa_nombre}.pdf", "application/pdf", disabled=not listo_para_descargar, use_container_width=True)
col_b2.download_button("📊 COT EXCEL", xlsx_cot, f"Cot_{cotizacion_num}_{empresa_nombre}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disabled=not listo_para_descargar, use_container_width=True)
col_b3.download_button("📈 CUADRO CONTROL", xlsx_ctrl, f"Control_{cotizacion_num}_{empresa_nombre}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disabled=not listo_para_descargar, use_container_width=True)

st.markdown("<div style='text-align: center; font-size: 10px; color: #444746; padding-top: 60px;'>by Oscar Buitrago / Elyon Producciones.</div>", unsafe_allow_html=True)