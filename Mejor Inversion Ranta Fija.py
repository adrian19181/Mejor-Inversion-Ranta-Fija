import sys
import io
import datetime
import os
import re
import unicodedata
import requests
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import streamlit.components.v1 as components

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

if __name__ == "__main__" and not st.runtime.exists():
    from streamlit.web import cli as stcli
    sys.argv = ["streamlit", "run", __file__]
    sys.exit(stcli.main())

st.set_page_config(page_title="Mejor Inversión Renta Fija", page_icon="💰", layout="centered", initial_sidebar_state="collapsed")

components.html("""
<script>
const doc = window.parent.document;
doc.addEventListener('click', function(e) {
    if (e.target && e.target.tagName === 'INPUT' && (e.target.type === 'number' || e.target.type === 'text')) {
        e.target.select();
    }
});
</script>
""", height=0, width=0)

st.markdown("""
<style>
    h1, h2, h3, h4, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { text-align: center !important; }
    button[data-baseweb="tab"] { font-size: 16px !important; font-weight: 600 !important; }
    [data-testid="stInputInstruction"], small[data-testid="stInputInstruction"], div[data-testid="InputInstruction"], div[data-testid="InputInstructions"], div[data-testid="stFormInstruction"], div[data-baseweb="typo-caption"] { display: none !important; }
    input::-webkit-input-placeholder, input::-moz-placeholder, input:-ms-input-placeholder, input::placeholder { color: transparent !important; opacity: 0 !important; }
    div[data-baseweb="input"] input { text-align: center !important; font-size: 18px !important; font-weight: 600 !important; }
    div[data-testid="stNumberInputContainer"] button, button[data-testid="stNumberInputStepDown"], button[data-testid="stNumberInputStepUp"] { display: none !important; }
    input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none !important; margin: 0 !important; }
    input[type=number] { -moz-appearance: textfield !important; }
    .table-container { width: 100%; overflow-x: auto; margin-top: 15px; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); background-color: #0F172A; }
    .custom-table { width: 100%; border-collapse: separate; border-spacing: 0; table-layout: auto; font-family: -apple-system, sans-serif; }
    .custom-table th { background-color: #1E293B; color: #FFFFFF; text-align: center !important; vertical-align: middle !important; padding: 10px 8px; font-size: 13px; font-weight: 700; border-bottom: 2px solid #334155; border-top: 1px solid #334155; }
    .custom-table td { text-align: center !important; vertical-align: middle !important; padding: 10px 8px; font-size: 14px; white-space: nowrap; border-bottom: 1px solid #334155; }
    .custom-table th:nth-child(1) { position: sticky; left: 0; z-index: 5; background-color: #1E293B; min-width: 110px; max-width: 130px; white-space: normal !important; word-wrap: break-word !important; border-right: 2px solid #64748B !important; }
    .custom-table td:nth-child(1) { position: sticky; left: 0; z-index: 3; background-color: #0F172A; min-width: 110px; max-width: 130px; white-space: normal !important; word-wrap: break-word !important; border-right: 2px solid #64748B !important; }
    .custom-table tr:nth-child(even) td { background-color: #182234 !important; }
    .custom-table tr:hover td { background-color: #243248 !important; }
    .kpi-flex-container { display: flex; flex-direction: row; justify-content: space-between; gap: 8px; width: 100%; margin-top: 10px; }
    .kpi-card-compact { flex: 1; background-color: #1E293B; border: 1px solid #334155; border-radius: 10px; padding: 10px 4px; text-align: center; box-shadow: 0 2px 6px rgba(0,0,0,0.2); }
    .kpi-title-compact { color: #94A3B8; font-size: 11px; font-weight: 600; margin-bottom: 4px; line-height: 1.2; }
    .kpi-value-compact { font-size: 22px; font-weight: 800; line-height: 1.1; }

    /* BOTONES UNIFICADOS ESTILO EXCEL / PREMIUM */
    div[data-testid="stButton"] > button, .stDownloadButton > button {
        background-color: #1E222B !important;
        border: 1.5px solid #107C41 !important;
        border-radius: 8px !important;
        padding: 0.4rem 0.6rem !important;
        width: 100% !important; 
        font-size: 0.85rem !important;
        font-weight: 700 !important;
    }
    div[data-testid="stButton"] > button *, .stDownloadButton > button * {
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 700 !important;
    }
    div[data-testid="stButton"] > button:hover, .stDownloadButton > button:hover {
        background-color: #107C41 !important;
        border-color: #107C41 !important;
    }
</style>
""", unsafe_allow_html=True)

FILE_ID_INVERSIONES = "1DPASEKWo7lTbu8tug041iK5lMXLlyJvG"
DRIVE_URL_INVERSIONES = f"https://docs.google.com/spreadsheets/d/{FILE_ID_INVERSIONES}/export?format=xlsx"
LOCAL_BACKUP_INVERSIONES = "Inversiones_Ecuador.xlsx"

FILE_ID_RIESGO = "1ZtmezQ9fQfpUmpRdHa66Fp0j733j5qwv"
DRIVE_URL_RIESGO = f"https://docs.google.com/spreadsheets/d/{FILE_ID_RIESGO}/export?format=xlsx"
LOCAL_BACKUP_RIESGO = "Calificaciones_Riesgo.xlsx"

def limpiar_numeros(serie):
    s = pd.Series(serie).astype(str).str.strip()
    s = s.str.replace(r'[%\$]', '', regex=True)
    s = s.apply(lambda x: x.replace(',', '') if ',' in str(x) and '.' in str(x) else str(x).replace(',', '.'))
    s = s.replace(['nan', 'NAN', 'None', ''], np.nan)
    return pd.to_numeric(s, errors='coerce')

def formatear_nombre_banco(nombre):
    if pd.isna(nombre) or nombre is None:
        return ""
    nombre_str = str(nombre).strip()
    palabras = nombre_str.split()
    if len(palabras) > 1:
        mitad = (len(palabras) + 1) // 2
        return f"{' '.join(palabras[:mitad])}<br>{' '.join(palabras[mitad:])}"
    return nombre_str

def es_columna_vacia(col_series):
    s = col_series.fillna('').astype(str).str.strip().str.lower()
    return s.isin(['', 'nan', 'none']).all()

def normalizar_texto(texto):
    if pd.isna(texto): return ""
    texto = str(texto).strip().lower()
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

@st.cache_data(ttl=3600)
def cargar_calificaciones_riesgo():
    content = None
    try:
        response = requests.get(DRIVE_URL_RIESGO, timeout=15)
        if response.status_code == 200:
            content = response.content
    except Exception:
        pass

    if content is None and os.path.exists(LOCAL_BACKUP_RIESGO):
        try:
            with open(LOCAL_BACKUP_RIESGO, 'rb') as f:
                content = f.read()
        except Exception:
            pass

    if content is None:
        return {}

    mapping = {}

    try:
        xl = pd.ExcelFile(io.BytesIO(content))
        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet)
            if df.empty or len(df.columns) < 2: continue

            col_inst = df.columns[0]
            col_rating = df.columns[-1]

            for _, row in df.iterrows():
                inst_raw = str(row[col_inst]).strip()
                rating_raw = str(row[col_rating]).strip()
                if inst_raw and rating_raw and rating_raw.lower() not in ['nan', 'none', '']:
                    norm_k = normalizar_texto(inst_raw)
                    mapping[norm_k] = rating_raw
        if mapping:
            return mapping
    except Exception:
        pass

    try:
        txt = content.decode('utf-8', errors='ignore')
        pattern_split = r'(?=(?:BANCO|COOPERATIVA|CORPORACI[ÓO]N|ASOCIACI[ÓO]N|MUTUALISTA|CAJA|TITULARIZACI[ÓO]N|BANECUADOR|BIESS|FONDO)\b)'
        entries = re.split(pattern_split, txt)
        pattern_next_inst = r'\s+(?:BANCO|COOPERATIVA|CORPORACI[ÓO]N|ASOCIACI[ÓO]N|MUTUALISTA|CAJA|TITULARIZACI[ÓO]N|BANECUADOR|BIESS|FONDO)\b.*$'

        for entry in entries:
            entry = entry.strip()
            if not entry or "INSTITUCIÓN FINANCIERA" in entry: continue

            parts = [p.strip() for p in entry.split(',')]
            if len(parts) >= 4:
                nombre = parts[0]
                raw_last = parts[-1]
                rating_clean = re.sub(pattern_next_inst, '', raw_last, flags=re.IGNORECASE).strip()

                if rating_clean and not rating_clean.startswith("http") and len(rating_clean) <= 30:
                    norm_n = normalizar_texto(nombre)
                    mapping[norm_n] = rating_clean
    except Exception:
        pass

    return mapping

def buscar_calificacion_riesgo(banco_nombre, mapping):
    if not mapping or not banco_nombre:
        return "N/A"

    b_norm = normalizar_texto(banco_nombre)

    alias_map = {
        "jep": "juventud ecuatoriana progresista",
        "sanfra": "san francisco",
        "produbanco": "banco produbanco s.a.",
        "cooperco": "erco",
        "pacifico": "banco pacifico s.a."
    }

    target_search = b_norm
    for alias_key, alias_val in alias_map.items():
        if alias_key in b_norm:
            target_search = alias_val
            break

    if target_search in mapping:
        return mapping[target_search]

    for k, v in mapping.items():
        if target_search in k or k in target_search:
            return v

    b_tokens = [t for t in target_search.split() if t not in ['banco', 'cooperativa', 'de', 'ahorro', 'y', 'credito', 's.a.', 'ltda', 'limitada', 'c.a.', 'n.a.', 'asociacion', 'mutualista', 'coop']]
    if b_tokens:
        tok_str = ' '.join(b_tokens)
        for k, v in mapping.items():
            if tok_str in k:
                return v

    return "N/A"

# ==============================================================================
# FUNCIÓN GENERADORA DE PDF DINÁMICA DE ENTRADAS + TABLA DE RESULTADOS
# ==============================================================================
def generar_pdf_dinamico(resultados_df, tipo, monto, dias, map_riesgo):
    if resultados_df is None or resultados_df.empty:
        return None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )
    story = []
    styles = getSampleStyleSheet()

    style_title = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=15, leading=17, textColor=colors.HexColor('#0F172A'), alignment=1, spaceAfter=2)
    style_sub = ParagraphStyle('DocSub', parent=styles['Normal'], fontSize=8.5, leading=10, textColor=colors.HexColor('#475569'), alignment=1, spaceAfter=8)
    style_h2 = ParagraphStyle('DocH2', parent=styles['Heading2'], fontSize=10.5, leading=12, textColor=colors.HexColor('#1E293B'), spaceBefore=6, spaceAfter=4)
    style_cell = ParagraphStyle('DocCell', parent=styles['Normal'], fontSize=7.5, leading=9, textColor=colors.HexColor('#0F172A'), alignment=1)
    style_cell_left = ParagraphStyle('DocCellLeft', parent=style_cell, alignment=0)
    style_cell_bold = ParagraphStyle('DocCellBold', parent=style_cell, fontName='Helvetica-Bold')
    style_cell_header = ParagraphStyle('DocCellHeader', parent=style_cell, fontName='Helvetica-Bold', textColor=colors.white, alignment=1)

    titulos_tipo = {
        'pf': ("Reporte - Mejores Inversiones a Plazo Fijo", "Capital a Invertir ($ USD)", "Plazo deseado (Días)"),
        'ap': ("Reporte - Mejores Opciones de Ahorro Programado", "Aporte Mensual / Apertura ($ USD)", "Plazo a cumplir (Días)"),
        'ca': ("Reporte - Mejores Cuentas de Ahorros", "Saldo Promedio a Mantener ($ USD)", "Plazo a simular (Días)")
    }

    titulo_txt, label_monto, label_dias = titulos_tipo.get(tipo, ("Reporte de Inversiones", "Monto", "Plazo"))
    fecha_hoy = datetime.date.today().strftime('%d/%m/%Y')

    story.append(Paragraph(f"💰 {titulo_txt}", style_title))
    story.append(Paragraph(f"Fecha de generación: {fecha_hoy}", style_sub))

    # --- DATOS DE ENTRADA DEL FORMULARIO ---
    story.append(Paragraph("📋 Parámetros de Entrada del Formulario", style_h2))
    input_data = [
        [Paragraph(f"<b>{label_monto}:</b> ${monto:,.2f}", style_cell_left), Paragraph(f"<b>{label_dias}:</b> {dias} días", style_cell_left)]
    ]
    t_input = Table(input_data, colWidths=[270, 270])
    t_input.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(t_input)
    story.append(Spacer(1, 8))

    # --- TABLA COMPLETA GENERADA ---
    story.append(Paragraph("🏆 Tabla de Resultados Generada", style_h2))

    if tipo == 'pf':
        headers = ["Institución Financiera", "Calificación de Riesgo", "Rango Plazo Aplicado", "TNA (%)", "TNA Periodo (%)", "Ganancia Estimada ($)"]
        col_widths = [130, 75, 95, 60, 85, 95]
        table_rows = [[Paragraph(h, style_cell_header) for h in headers]]

        for idx, row in resultados_df.iterrows():
            banco = str(row['banco']).strip()
            calif = buscar_calificacion_riesgo(banco, map_riesgo)
            
            if (row['dias_minimo'] in [0, 1]) and row['dias_maximo'] == 9999:
                rango = "Libre / Sin plazo"
            elif row['dias_minimo'] == row['dias_maximo']:
                rango = f"{int(row['dias_minimo'])} días"
            else:
                rango = f"{int(row['dias_minimo'])} - {int(row['dias_maximo'])} días"
            if not row.get('es_exacto', True):
                rango += " (Más cercano)"

            tna_val = float(row['tna'])
            tna_periodo_val = tna_val * (dias / 360.0)
            ganancia_val = float(row.get('ganancia_estimada', 0))

            table_rows.append([
                Paragraph(banco, style_cell_left),
                Paragraph(calif, style_cell),
                Paragraph(rango, style_cell),
                Paragraph(f"{tna_val:.2f}%", style_cell_bold),
                Paragraph(f"{tna_periodo_val:.2f}%", style_cell),
                Paragraph(f"${ganancia_val:,.2f}", style_cell_bold)
            ])

    elif tipo == 'ap':
        headers = ["Institución / Producto", "Calificación de Riesgo", "Rango Plazo Aplicado", "TNA (%)", "TEA (%)", "TEA Periodo (%)"]
        col_widths = [140, 75, 95, 65, 80, 85]
        table_rows = [[Paragraph(h, style_cell_header) for h in headers]]

        for idx, row in resultados_df.iterrows():
            banco = str(row['banco']).strip()
            prod = str(row.get('producto', '')).title()
            inst = f"{banco}<br/><font size=6.5 color='#475569'>{prod}</font>" if prod else banco
            calif = buscar_calificacion_riesgo(banco, map_riesgo)

            if (row['dias_minimo'] in [0, 1]) and row['dias_maximo'] == 9999:
                rango = "Libre / Sin plazo"
            elif row['dias_minimo'] == row['dias_maximo']:
                rango = f"{int(row['dias_minimo'])} días"
            else:
                rango = f"{int(row['dias_minimo'])} - {int(row['dias_maximo'])} días"
            if not row.get('es_exacto', True):
                rango += " (Más cercano)"

            tna_val = float(row['tna'])
            tea_val = ((1 + (tna_val / 1200.0)) ** 12 - 1) * 100.0
            tea_periodo_val = ((1 + (tna_val / 1200.0)) ** (12.0 * (dias / 360.0)) - 1) * 100.0

            table_rows.append([
                Paragraph(inst, style_cell_left),
                Paragraph(calif, style_cell),
                Paragraph(rango, style_cell),
                Paragraph(f"{tna_val:.2f}%", style_cell_bold),
                Paragraph(f"{tea_val:.2f}%", style_cell),
                Paragraph(f"{tea_periodo_val:.2f}%", style_cell_bold)
            ])

    elif tipo == 'ca':
        headers = ["Institución / Producto", "Calificación de Riesgo", "Rango de Saldos Aplicado", "TNA (%)", "TNA Periodo (%)", "TEA (%)", "Ganancia Estimada ($)"]
        col_widths = [115, 65, 95, 55, 65, 60, 85]
        table_rows = [[Paragraph(h, style_cell_header) for h in headers]]

        for idx, row in resultados_df.iterrows():
            banco = str(row['banco']).strip()
            prod = str(row.get('producto', '')).title()
            inst = f"{banco}<br/><font size=6.5 color='#475569'>{prod}</font>" if prod else banco
            calif = buscar_calificacion_riesgo(banco, map_riesgo)

            if pd.isna(row['monto_maximo']):
                rango_saldos = f"Desde ${float(row['monto_minimo']):,.2f}"
            else:
                rango_saldos = f"${float(row['monto_minimo']):,.2f} - ${float(row['monto_maximo']):,.2f}"

            tna_val = float(row['tna'])
            tna_periodo_val = tna_val * (dias / 360.0)
            tea_val = ((1 + (tna_val / 1200.0)) ** 12 - 1) * 100.0
            ganancia_val = float(row.get('ganancia_estimada', 0))

            table_rows.append([
                Paragraph(inst, style_cell_left),
                Paragraph(calif, style_cell),
                Paragraph(rango_saldos, style_cell),
                Paragraph(f"{tna_val:.2f}%", style_cell_bold),
                Paragraph(f"{tna_periodo_val:.2f}%", style_cell),
                Paragraph(f"${tea_val:.2f}%", style_cell),
                Paragraph(f"${ganancia_val:,.2f}", style_cell_bold)
            ])

    t_results = Table(table_rows, colWidths=col_widths)
    t_results.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E293B')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')])
    ]))

    story.append(t_results)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def generar_pdf_bytes(df, tipo, monto, dias, map_riesgo):
    return generar_pdf_dinamico(df, tipo, monto, dias, map_riesgo)

def obtener_ranking_bancos(df, monto, dias, tipo='pf'):
    if df.empty or monto <= 0 or dias <= 0:
        return pd.DataFrame()
        
    resultados = []
    
    for banco in df['banco'].unique():
        df_banco = df[df['banco'] == banco].copy()
        banco_upper = str(banco).upper()
        
        if tipo == 'ca':
            mask_monto = (df_banco['monto_minimo'].isna() | (df_banco['monto_minimo'] <= monto)) & \
                         (df_banco['monto_maximo'].isna() | (df_banco['monto_maximo'] >= monto))
            df_monto_valido = df_banco[mask_monto].copy()
            
            if df_monto_valido.empty:
                continue
                
            for prod in df_monto_valido['producto'].unique():
                df_prod = df_monto_valido[df_monto_valido['producto'] == prod]
                mejor = df_prod.sort_values(by='tna', ascending=False).iloc[0].copy()
                mejor['es_exacto'] = True
                resultados.append(mejor)
            continue

        if tipo == 'ap' and "SOLIDARIO" in banco_upper:
            df_monto_match = df_banco[(df_banco['monto_minimo'].isna() | (df_banco['monto_minimo'] <= monto)) & \
                                      (df_banco['monto_maximo'].isna() | (df_banco['monto_maximo'] >= monto))]
            
            if df_monto_match.empty:
                continue
                
            row_monto = df_monto_match.iloc[0]
            tna_base = row_monto['tna_base']
            
            if tna_base <= 0:
                continue
                
            df_bonos = df_banco[(df_banco['bono'] > 0) & (df_banco['dias_minimo'] > 0)].copy()
            
            applied_bono = 0.0
            applied_dmin = 0
            applied_dmax = 9999
            es_exacto = True
            
            if not df_bonos.empty:
                exact_bono = df_bonos[(df_bonos['dias_minimo'] <= dias) & (df_bonos['dias_maximo'] >= dias)]
                if not exact_bono.empty:
                    best_bono_row = exact_bono.iloc[0]
                    applied_bono = float(best_bono_row['bono'])
                    applied_dmin = int(best_bono_row['dias_minimo'])
                    applied_dmax = int(best_bono_row['dias_maximo']) if float(best_bono_row['dias_maximo']) > 0 else 9999
                    es_exacto = True
                else:
                    df_bonos['distancia_dias'] = df_bonos.apply(
                        lambda row: min(abs(dias - row['dias_minimo']), abs(dias - row['dias_maximo'])), axis=1
                    )
                    min_dist = df_bonos['distancia_dias'].min()
                    if min_dist <= 10:
                        closest_bono_row = df_bonos[df_bonos['distancia_dias'] == min_dist].iloc[0]
                        applied_bono = float(closest_bono_row['bono'])
                        applied_dmin = int(closest_bono_row['dias_minimo'])
                        applied_dmax = int(closest_bono_row['dias_maximo']) if float(closest_bono_row['dias_maximo']) > 0 else 9999
                        es_exacto = False
                    else:
                        applied_dmin = int(row_monto['dias_minimo']) if 'dias_minimo' in row_monto and pd.notna(row_monto['dias_minimo']) else 0
                        applied_dmax = int(row_monto['dias_maximo']) if 'dias_maximo' in row_monto and pd.notna(row_monto['dias_maximo']) else 9999
                        es_exacto = True
            
            tna_total = tna_base + applied_bono
            
            mejor = pd.Series({
                'banco': banco,
                'producto': row_monto.get('producto', 'Futuro Seguro'),
                'monto_minimo': row_monto['monto_minimo'],
                'monto_maximo': row_monto['monto_maximo'],
                'dias_minimo': applied_dmin,
                'dias_maximo': applied_dmax,
                'tna': tna_total,
                'tna_base': tna_base,
                'bono': applied_bono,
                'es_exacto': es_exacto
            })
            resultados.append(mejor)
            continue

        mask_monto = (df_banco['monto_minimo'].isna() | (df_banco['monto_minimo'] <= monto)) & \
                     (df_banco['monto_maximo'].isna() | (df_banco['monto_maximo'] >= monto))
        df_monto_valido = df_banco[mask_monto].copy()
        
        if df_monto_valido.empty:
            continue
        
        exact_dias = df_monto_valido[(df_monto_valido['dias_minimo'] <= dias) & (df_monto_valido['dias_maximo'] >= dias)]
        
        if not exact_dias.empty:
            mejor = exact_dias.sort_values(by='tna', ascending=False).iloc[0].copy()
            mejor['es_exacto'] = True
            resultados.append(mejor)
        else:
            df_monto_valido['distancia_dias'] = df_monto_valido.apply(
                lambda row: min(abs(dias - row['dias_minimo']), abs(dias - row['dias_maximo'])), axis=1
            )
            min_dist = df_monto_valido['distancia_dias'].min()
            
            if min_dist <= 10:
                closest_rows = df_monto_valido[df_monto_valido['distancia_dias'] == min_dist].copy()
                mejor = closest_rows.sort_values(by='tna', ascending=False).iloc[0].copy()
                mejor['es_exacto'] = False
                resultados.append(mejor)
            
    if resultados:
        res_df = pd.DataFrame(resultados).sort_values(by='tna', ascending=False).reset_index(drop=True)
        if tipo in ['pf', 'ca'] and dias > 0:
            res_df['ganancia_estimada'] = monto * (res_df['tna'] / 100.0) * (dias / 360.0)
        return res_df
        
    return pd.DataFrame()

@st.cache_data(ttl=3600)
def cargar_datos_inversiones():
    content = None
    try:
        response = requests.get(DRIVE_URL_INVERSIONES, timeout=15)
        if response.status_code == 200: content = io.BytesIO(response.content)
    except Exception: pass

    xl = None
    if content is not None:
        try:
            xl = pd.ExcelFile(content)
        except Exception: pass

    if xl is None and os.path.exists(LOCAL_BACKUP_INVERSIONES):
        try:
            xl = pd.ExcelFile(LOCAL_BACKUP_INVERSIONES)
        except Exception: pass

    if xl is None:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    tablas_pf = []
    tablas_ap = []
    tablas_ca = []

    BANCOS_AP_VALIDOS = [
        "AUSTRO", "JARDIN AZUAYO", "MACHALA", 
        "PRODUBANCO", "29 DE OCTUBRE", "COOPERCO", "JEP", "SOLIDARIO", "ALIANZA"
    ]

    for sheet in xl.sheet_names:
        if sheet == "Tasas Referenciales": continue
        
        try:
            df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
            if len(df_raw) < 3: continue
            
            banco_norm = str(sheet).upper()

            non_empty_cols = [i for i in range(df_raw.shape[1]) if not es_columna_vacia(df_raw.iloc[:, i])]
            if not non_empty_cols: continue

            bloques_tablas = []
            curr_block = [non_empty_cols[0]]
            for col_idx in non_empty_cols[1:]:
                if col_idx == curr_block[-1] + 1:
                    curr_block.append(col_idx)
                else:
                    bloques_tablas.append(curr_block)
                    curr_block = [col_idx]
            bloques_tablas.append(curr_block)

            if len(bloques_tablas) >= 1:
                cols_pf = bloques_tablas[0]
                df_bloque_raw = df_raw.iloc[:, cols_pf].copy()
                row1_headers = [normalizar_texto(h) for h in df_bloque_raw.iloc[1].tolist()]
                df_data = df_bloque_raw.iloc[2:].copy().reset_index(drop=True)
                df_data.columns = [f"col_{i}_{h}" for i, h in enumerate(row1_headers)]

                mmin_col = next((c for c in df_data.columns if 'monto_min' in c or 'aporte_min' in c), None)
                mmax_col = next((c for c in df_data.columns if 'monto_max' in c or 'aporte_max' in c), None)
                dmin_col = next((c for c in df_data.columns if 'dias_min' in c or 'plazo_min' in c or 'dias' in c or 'plazo' in c), None)
                dmax_col = next((c for c in df_data.columns if 'dias_max' in c or 'plazo_max' in c), None)
                tna_col = next((c for c in df_data.columns if 'tasa_nominal' in c or 'tasa_pago' in c or 'tna' in c or 'tasa' in c), None)

                if tna_col is None and len(df_data.columns) >= 1:
                    tna_col = df_data.columns[-1]

                if tna_col:
                    mmin = limpiar_numeros(df_data[mmin_col]) if mmin_col else pd.Series(np.nan, index=df_data.index)
                    mmax = limpiar_numeros(df_data[mmax_col]) if mmax_col else pd.Series(np.nan, index=df_data.index)
                    
                    if dmin_col:
                        dmin = limpiar_numeros(df_data[dmin_col]).fillna(1).astype(int)
                        dmax = limpiar_numeros(df_data[dmax_col]).fillna(9999).astype(int) if dmax_col else pd.Series(9999, index=df_data.index)
                    else:
                        dmin = pd.Series(1, index=df_data.index)
                        dmax = pd.Series(9999, index=df_data.index)

                    tna_base = limpiar_numeros(df_data[tna_col]).fillna(0.0)

                    temp_df = pd.DataFrame({
                        'banco': sheet,
                        'monto_minimo': mmin,
                        'monto_maximo': mmax,
                        'dias_minimo': dmin,
                        'dias_maximo': dmax,
                        'tna': tna_base,
                        'tna_base': tna_base,
                        'bono': 0.0
                    })
                    temp_df = temp_df[temp_df['tna'] > 0]
                    if not temp_df.empty:
                        tablas_pf.append(temp_df)

            if len(bloques_tablas) >= 2 and any(b in banco_norm for b in BANCOS_AP_VALIDOS):
                cols_ap = bloques_tablas[1]
                df_bloque_raw = df_raw.iloc[:, cols_ap].copy()
                row1_headers = [normalizar_texto(h) for h in df_bloque_raw.iloc[1].tolist()]
                df_data = df_bloque_raw.iloc[2:].copy().reset_index(drop=True)
                df_data.columns = [f"col_{i}_{h}" for i, h in enumerate(row1_headers)]

                mmin_col = next((c for c in df_data.columns if 'monto_min' in c or 'aporte_min' in c), None)
                mmax_col = next((c for c in df_data.columns if 'monto_max' in c or 'aporte_max' in c), None)
                dmin_col = next((c for c in df_data.columns if 'dias_min' in c or 'plazo_min' in c or 'dias' in c or 'plazo' in c), None)
                dmax_col = next((c for c in df_data.columns if 'dias_max' in c or 'plazo_max' in c), None)
                tna_col = next((c for c in df_data.columns if 'tasa_nominal' in c or 'tasa_pago' in c or 'tna' in c or 'tasa' in c), None)
                bono_col = next((c for c in df_data.columns if 'bonificac' in c or 'premio' in c or 'bono' in c), None)

                if tna_col is None and len(df_data.columns) >= 1:
                    tna_col = df_data.columns[-1]

                if tna_col:
                    mmin = limpiar_numeros(df_data[mmin_col]) if mmin_col else pd.Series(np.nan, index=df_data.index)
                    mmax = limpiar_numeros(df_data[mmax_col]) if mmax_col else pd.Series(np.nan, index=df_data.index)
                    
                    if dmin_col:
                        dmin = limpiar_numeros(df_data[dmin_col]).fillna(0).astype(int)
                        dmax = limpiar_numeros(df_data[dmax_col]).fillna(9999).astype(int) if dmax_col else pd.Series(9999, index=df_data.index)
                    else:
                        dmin = pd.Series(0, index=df_data.index)
                        dmax = pd.Series(9999, index=df_data.index)

                    tna_base = limpiar_numeros(df_data[tna_col]).fillna(0.0)
                    bono_val = limpiar_numeros(df_data[bono_col]).fillna(0.0) if bono_col else pd.Series(0.0, index=df_data.index)
                    tna_total = tna_base + bono_val

                    prod_name = "Ahorro Programado"
                    if "JARDIN AZUAYO" in banco_norm: prod_name = "Mi Ahorro Chanchito"
                    elif "PRODUBANCO" in banco_norm: prod_name = "Flexi Ahorro"
                    elif "JEP" in banco_norm: prod_name = "Ahorro Fondo JEP"
                    elif "SOLIDARIO" in banco_norm: prod_name = "Futuro Seguro"
                    elif "AUSTRO" in banco_norm: prod_name = "Ahorro Programado"
                    elif "MACHALA" in banco_norm: prod_name = "Ahorro Programado"
                    elif "29 DE OCTUBRE" in banco_norm: prod_name = "Ahorro Programado"
                    elif "COOPERCO" in banco_norm: prod_name = "Ahorro Programado"
                    elif "ALIANZA" in banco_norm: prod_name = "Ahorro Programado"

                    temp_df = pd.DataFrame({
                        'banco': sheet,
                        'producto': prod_name,
                        'monto_minimo': mmin,
                        'monto_maximo': mmax,
                        'dias_minimo': dmin,
                        'dias_maximo': dmax,
                        'tna': tna_total,
                        'tna_base': tna_base,
                        'bono': bono_val
                    })
                    temp_df = temp_df[temp_df['tna'] > 0]
                    if not temp_df.empty:
                        tablas_ap.append(temp_df)

            indices_ca = []
            if "AUSTRO" in banco_norm: indices_ca = [2]
            elif "JARDIN AZUAYO" in banco_norm: indices_ca = [2]
            elif "MUTUALISTA AZUAY" in banco_norm: indices_ca = [1, 2]
            elif "COOPERCO" in banco_norm: indices_ca = [2]
            elif "JEP" in banco_norm: indices_ca = [2]
            elif "SOLIDARIO" in banco_norm: indices_ca = [2]

            for idx_ca in indices_ca:
                if len(bloques_tablas) > idx_ca:
                    cols_ca = bloques_tablas[idx_ca]
                    df_bloque_raw = df_raw.iloc[:, cols_ca].copy()
                    row1_headers = [normalizar_texto(h) for h in df_bloque_raw.iloc[1].tolist()]
                    df_data = df_bloque_raw.iloc[2:].copy().reset_index(drop=True)
                    df_data.columns = [f"col_{i}_{h}" for i, h in enumerate(row1_headers)]

                    mmin_col = next((c for c in df_data.columns if 'monto_min' in c or 'saldo_min' in c or 'monto' in c or 'saldo' in c), None)
                    mmax_col = next((c for c in df_data.columns if 'monto_max' in c or 'saldo_max' in c), None)
                    tna_col = next((c for c in df_data.columns if 'tasa_nominal' in c or 'tna' in c or 'tasa_anual' in c or 'tasa' in c), None)

                    if tna_col is None and len(df_data.columns) >= 1:
                        tna_col = df_data.columns[-1]

                    if tna_col:
                        mmin = limpiar_numeros(df_data[mmin_col]) if mmin_col else pd.Series(0.0, index=df_data.index)
                        mmax = limpiar_numeros(df_data[mmax_col]) if mmax_col else pd.Series(np.nan, index=df_data.index)
                        tna_base = limpiar_numeros(df_data[tna_col]).fillna(0.0)

                        main_header = str(df_bloque_raw.iloc[0].dropna().values[0]).strip() if not df_bloque_raw.iloc[0].dropna().empty else "Cuenta de Ahorros"

                        temp_df = pd.DataFrame({
                            'banco': sheet,
                            'producto': main_header.title(),
                            'monto_minimo': mmin,
                            'monto_maximo': mmax,
                            'dias_minimo': 0,
                            'dias_maximo': 9999,
                            'tna': tna_base,
                            'tna_base': tna_base,
                            'bono': 0.0
                        })
                        temp_df = temp_df[temp_df['tna'] > 0]
                        if not temp_df.empty:
                            tablas_ca.append(temp_df)

        except Exception:
            continue

    df_final_pf = pd.concat(tablas_pf, ignore_index=True) if tablas_pf else pd.DataFrame()
    df_final_ap = pd.concat(tablas_ap, ignore_index=True) if tablas_ap else pd.DataFrame()
    df_final_ca = pd.concat(tablas_ca, ignore_index=True) if tablas_ca else pd.DataFrame()
    return df_final_pf, df_final_ap, df_final_ca

@st.cache_data(ttl=3600)
def cargar_tasas_referenciales():
    content = None
    try:
        response = requests.get(DRIVE_URL_INVERSIONES, timeout=5)
        if response.status_code == 200: content = io.BytesIO(response.content)
    except Exception: pass
    
    xl_source = content if content is not None else LOCAL_BACKUP_INVERSIONES
    try:
        is_valid_source = isinstance(xl_source, io.BytesIO) or (isinstance(xl_source, str) and os.path.exists(xl_source))
        if is_valid_source:
            xl_doc = pd.ExcelFile(xl_source)
            if "Tasas Referenciales" in xl_doc.sheet_names: 
                return pd.read_excel(xl_doc, sheet_name="Tasas Referenciales")
    except Exception: pass
    return pd.DataFrame()

# ---------------------------------------------------------
# INTERFAZ PRINCIPAL
# ---------------------------------------------------------
col_title, col_btn_act = st.columns([3, 1], vertical_alignment="center")
with col_title:
    st.markdown("<h1 style='text-align: center; margin: 0;'>Mejor Inversión Renta Fija</h1>", unsafe_allow_html=True)
with col_btn_act:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

df_pf, df_ap, df_ca = cargar_datos_inversiones()
map_riesgo = cargar_calificaciones_riesgo()

if df_pf.empty and df_ap.empty and df_ca.empty:
    st.warning("No se encontraron datos de inversiones disponibles.")
    st.stop()

tab_pf, tab_ap, tab_ca = st.tabs(["💰 Plazo Fijo", "📅 Ahorro Programado", "💳 Cuentas de Ahorros"])

with tab_pf:
    st.markdown("<br><h3 style='text-align: center;'>📋 Parámetros de Inversión</h3>", unsafe_allow_html=True)
    with st.form(key="formulario_pf"):
        col1, col2 = st.columns(2)
        with col1: monto_input = st.number_input("Capital a Invertir ($ USD)", min_value=0.0, value=None, step=1.0)
        with col2: plazo_input = st.number_input("Plazo deseado (Días)", min_value=0, max_value=3600, value=None, step=1)
        submit_pf = st.form_submit_button(label="🎯 Calcular Mejores Opciones", use_container_width=True)

    if submit_pf:
        if monto_input is None or plazo_input is None:
            st.warning("Por favor ingresa tanto el monto como el plazo para realizar la consulta.")
        elif monto_input <= 0 or plazo_input <= 0:
            st.warning("El capital y el plazo deben ser valores mayores a cero (0).")
        else:
            monto = float(monto_input)
            plazo_dias = int(plazo_input)

            if not df_pf.empty:
                resultados_pf = obtener_ranking_bancos(df_pf, monto, plazo_dias, tipo='pf')

                if resultados_pf.empty:
                    st.error(f"No hay opciones disponibles para un monto de ${monto:,.2f} a {plazo_dias} días.")
                else:
                    st.markdown("---")
                    st.markdown("<h3 style='text-align: center;'>🏆 Ranking Completo Plazo Fijo</h3>", unsafe_allow_html=True)
                    ganador = resultados_pf.iloc[0]
                    st.success(f"🥇 **Opción Ganadora:** **{ganador['banco']}** con una TNA de **{ganador['tna']:.2f}%** y una ganancia estimada de **${ganador['ganancia_estimada']:,.2f}**")

                    pdf_data = generar_pdf_bytes(resultados_pf, 'pf', monto, plazo_dias, map_riesgo)
                    if pdf_data:
                        st.download_button(label="📄 Exportar a PDF", data=pdf_data, file_name="Mejor Inversion.pdf", mime="application/pdf", use_container_width=True)
                        st.markdown("<br>", unsafe_allow_html=True)

                    html_rows = ""
                    for idx, row in resultados_pf.iterrows():
                        banco_fmt = formatear_nombre_banco(row['banco'])
                        calif_riesgo = buscar_calificacion_riesgo(row['banco'], map_riesgo)
                        calif_fmt = f"<span style='color: #38bdf8; font-weight: bold;'>{calif_riesgo}</span>"
                        
                        if (row['dias_minimo'] == 0 or row['dias_minimo'] == 1) and row['dias_maximo'] == 9999:
                            rango = "Libre / Sin plazo"
                        elif row['dias_minimo'] == 0 and row['dias_maximo'] == 0:
                            rango = "Libre / Sin plazo"
                        elif row['dias_minimo'] == row['dias_maximo']:
                            rango = f"{int(row['dias_minimo'])} días"
                        else:
                            rango = f"{int(row['dias_minimo'])} - {int(row['dias_maximo'])} días"
                            
                        if not row['es_exacto']:
                            rango += f"<br><span style='font-size:10px; color:#ef4444;'>(Más cercano)</span>"
                            
                        tna = f"{row['tna']:.2f}%"
                        if 'bono' in row and row['bono'] > 0:
                            tna += f"<br><span style='font-size:10px; color:#94A3B8;'>({row['tna_base']:.2f}% Base + {row['bono']:.2f}% Bono)</span>"

                        tna_val = float(row['tna'])
                        tna_periodo_val = tna_val * (plazo_dias / 360.0)
                        tna_periodo_str = f"{tna_periodo_val:.2f}%"

                        ganancia = f"${row['ganancia_estimada']:,.2f}"
                        html_rows += f"<tr><td><strong>{banco_fmt}</strong></td><td>{calif_fmt}</td><td>{rango}</td><td><span style='color: #22c55e; font-weight: bold;'>{tna}</span></td><td><span style='color: #f59e0b; font-weight: bold;'>{tna_periodo_str}</span></td><td><span style='color: #3b82f6; font-weight: bold;'>{ganancia}</span></td></tr>"

                    st.markdown(f"""<div class="table-container"><table class="custom-table"><thead><tr><th>Institución<br>Financiera</th><th>Calificación<br>de Riesgo</th><th>Rango Plazo<br>Aplicado</th><th>TNA<br>(%)</th><th>TNA Periodo<br>(%)</th><th>Ganancia<br>Estimada ($)</th></tr></thead><tbody>{html_rows}</tbody></table></div>""", unsafe_allow_html=True)

with tab_ap:
    if df_ap.empty:
        st.info("No se encontraron productos de Ahorro Programado en la base de datos.")
    else:
        st.markdown("<br><h3 style='text-align: center;'>📋 Parámetros de Ahorro</h3>", unsafe_allow_html=True)
        with st.form(key="formulario_ap"):
            col1, col2 = st.columns(2)
            with col1: aporte_ap_input = st.number_input("Aporte Mensual / Apertura ($ USD)", min_value=0.0, value=None, step=1.0)
            with col2: dias_ap_input = st.number_input("Plazo a cumplir (Días)", min_value=0, max_value=3600, value=None, step=1)
            submit_ap = st.form_submit_button(label="🎯 Calcular Mejores Opciones", use_container_width=True)

        if submit_ap:
            if aporte_ap_input is None or dias_ap_input is None:
                st.warning("Por favor ingresa tanto el aporte mensual como el plazo para realizar la consulta.")
            elif aporte_ap_input <= 0 or dias_ap_input <= 0:
                st.warning("El aporte mensual y el plazo deben ser valores mayores a cero (0).")
            else:
                aporte_ap = float(aporte_ap_input)
                dias_ap = int(dias_ap_input)

                resultados_ap = obtener_ranking_bancos(df_ap, aporte_ap, dias_ap, tipo='ap')

                if resultados_ap.empty:
                    st.error(f"No hay opciones disponibles de Ahorro Programado para ${aporte_ap:,.2f} a {dias_ap} días.")
                else:
                    st.markdown("---")
                    st.markdown("<h3 style='text-align: center;'>🏆 Ranking Completo Ahorro Programado</h3>", unsafe_allow_html=True)
                    ganador_ap = resultados_ap.iloc[0]
                    
                    st.success(f"🥇 **Opción Ganadora:** **{ganador_ap['banco']} ({str(ganador_ap['producto']).title()})** con una TNA de **{ganador_ap['tna']:.2f}%**")

                    pdf_data_ap = generar_pdf_bytes(resultados_ap, 'ap', aporte_ap, dias_ap, map_riesgo)
                    if pdf_data_ap:
                        st.download_button(label="📄 Exportar a PDF", data=pdf_data_ap, file_name="Mejor Inversion.pdf", mime="application/pdf", use_container_width=True)
                        st.markdown("<br>", unsafe_allow_html=True)

                    html_rows_ap = ""
                    for idx, row in resultados_ap.iterrows():
                        banco_fmt = formatear_nombre_banco(row['banco'])
                        prod_fmt = f"<span style='font-size:11px; color:#94A3B8;'>{str(row['producto']).title()}</span>"
                        institucion = f"{banco_fmt}<br>{prod_fmt}"
                        calif_riesgo = buscar_calificacion_riesgo(row['banco'], map_riesgo)
                        calif_fmt = f"<span style='color: #38bdf8; font-weight: bold;'>{calif_riesgo}</span>"
                        
                        if (row['dias_minimo'] == 0 or row['dias_minimo'] == 1) and row['dias_maximo'] == 9999:
                            rango_dias = "Libre / Sin plazo"
                        elif row['dias_minimo'] == 0 and row['dias_maximo'] == 0:
                            rango_dias = "Libre / Sin plazo"
                        elif row['dias_minimo'] == row['dias_maximo']:
                            rango_dias = f"{int(row['dias_minimo'])} días"
                        else:
                            rango_dias = f"{int(row['dias_minimo'])} - {int(row['dias_maximo'])} días"
                            
                        if not row['es_exacto']:
                            rango_dias += f"<br><span style='font-size:10px; color:#ef4444;'>(Más cercano)</span>"
                            
                        tna = f"{row['tna']:.2f}%"
                        if 'bono' in row and row['bono'] > 0:
                            tna += f"<br><span style='font-size:10px; color:#94A3B8;'>({row['tna_base']:.2f}% Base + {row['bono']:.2f}% Bono)</span>"

                        tna_val = float(row['tna'])
                        tea_val = ((1 + (tna_val / 1200.0)) ** 12 - 1) * 100.0
                        tea_str = f"{tea_val:.2f}%"

                        tea_periodo_val = ((1 + (tna_val / 1200.0)) ** (12.0 * (dias_ap / 360.0)) - 1) * 100.0
                        tea_periodo_str = f"{tea_periodo_val:.2f}%"
                            
                        html_rows_ap += f"<tr><td><strong>{institucion}</strong></td><td>{calif_fmt}</td><td>{rango_dias}</td><td><span style='color: #22c55e; font-weight: bold;'>{tna}</span></td><td><span style='color: #38bdf8; font-weight: bold;'>{tea_str}</span></td><td><span style='color: #f59e0b; font-weight: bold;'>{tea_periodo_str}</span></td></tr>"

                    st.markdown(f"""<div class="table-container"><table class="custom-table"><thead><tr><th>Institución /<br>Producto</th><th>Calificación<br>de Riesgo</th><th>Rango de<br>Plazo Aplicado</th><th>TNA<br>(%)</th><th>TEA<br>(%)</th><th>TEA Periodo<br>(%)</th></tr></thead><tbody>{html_rows_ap}</tbody></table></div>""", unsafe_allow_html=True)

with tab_ca:
    if df_ca.empty:
        st.info("No se encontraron productos de Cuentas de Ahorros en la base de datos.")
    else:
        st.markdown("<br><h3 style='text-align: center;'>📋 Parámetros de Cuentas de Ahorros</h3>", unsafe_allow_html=True)
        with st.form(key="formulario_ca"):
            col1, col2 = st.columns(2)
            with col1: monto_ca_input = st.number_input("Saldo Promedio a Mantener ($ USD)", min_value=0.0, value=None, step=1.0)
            with col2: dias_ca_input = st.number_input("Plazo a simular (Días)", min_value=0, max_value=3600, value=None, step=1)
            submit_ca = st.form_submit_button(label="🎯 Calcular Mejores Opciones", use_container_width=True)

        if submit_ca:
            if monto_ca_input is None or dias_ca_input is None:
                st.warning("Por favor ingresa tanto el saldo promedio como el plazo a simular para realizar la consulta.")
            elif monto_ca_input <= 0 or dias_ca_input <= 0:
                st.warning("El saldo y el plazo deben ser valores mayores a cero (0).")
            else:
                monto_ca = float(monto_ca_input)
                dias_ca = int(dias_ca_input)

                resultados_ca = obtener_ranking_bancos(df_ca, monto_ca, dias_ca, tipo='ca')

                if resultados_ca.empty:
                    st.error(f"No hay opciones disponibles de Cuentas de Ahorros para un saldo de ${monto_ca:,.2f}.")
                else:
                    st.markdown("---")
                    st.markdown("<h3 style='text-align: center;'>🏆 Ranking Completo Cuentas de Ahorros</h3>", unsafe_allow_html=True)
                    ganador_ca = resultados_ca.iloc[0]
                    
                    st.success(f"🥇 **Opción Ganadora:** **{ganador_ca['banco']} ({str(ganador_ca['producto']).title()})** con una TNA de **{ganador_ca['tna']:.2f}%** y una ganancia estimada de **${ganador_ca['ganancia_estimada']:,.2f}**")

                    pdf_data_ca = generar_pdf_bytes(resultados_ca, 'ca', monto_ca, dias_ca, map_riesgo)
                    if pdf_data_ca:
                        st.download_button(label="📄 Exportar a PDF", data=pdf_data_ca, file_name="Mejor Inversion.pdf", mime="application/pdf", use_container_width=True)
                        st.markdown("<br>", unsafe_allow_html=True)

                    html_rows_ca = ""
                    for idx, row in resultados_ca.iterrows():
                        banco_fmt = formatear_nombre_banco(row['banco'])
                        prod_fmt = f"<span style='font-size:11px; color:#94A3B8;'>{str(row['producto']).title()}</span>"
                        institucion = f"{banco_fmt}<br>{prod_fmt}"
                        calif_riesgo = buscar_calificacion_riesgo(row['banco'], map_riesgo)
                        calif_fmt = f"<span style='color: #38bdf8; font-weight: bold;'>{calif_riesgo}</span>"
                        
                        rango_saldos = ""
                        if pd.isna(row['monto_maximo']):
                            rango_saldos = f"Desde ${float(row['monto_minimo']):,.2f}"
                        else:
                            rango_saldos = f"${float(row['monto_minimo']):,.2f} -${float(row['monto_maximo']):,.2f}"
                            
                        tna = f"{row['tna']:.2f}%"
                        tna_val = float(row['tna'])
                        
                        tna_periodo_val = tna_val * (dias_ca / 360.0)
                        tna_periodo_str = f"{tna_periodo_val:.2f}%"
                        
                        tea_val = ((1 + (tna_val / 1200.0)) ** 12 - 1) * 100.0
                        tea_str = f"{tea_val:.2f}%"

                        ganancia = f"${row['ganancia_estimada']:,.2f}"

                        html_rows_ca += f"<tr><td><strong>{institucion}</strong></td><td>{calif_fmt}</td><td>{rango_saldos}</td><td><span style='color: #22c55e; font-weight: bold;'>{tna}</span></td><td><span style='color: #f59e0b; font-weight: bold;'>{tna_periodo_str}</span></td><td><span style='color: #38bdf8; font-weight: bold;'>{tea_str}</span></td><td><span style='color: #3b82f6; font-weight: bold;'>{ganancia}</span></td></tr>"

                    st.markdown(f"""<div class="table-container"><table class="custom-table"><thead><tr><th>Institución /<br>Producto</th><th>Calificación<br>de Riesgo</th><th>Rango de<br>Saldos Aplicado</th><th>TNA<br>(%)</th><th>TNA Periodo<br>(%)</th><th>TEA<br>(%)</th><th>Ganancia<br>Estimada ($)</th></tr></thead><tbody>{html_rows_ca}</tbody></table></div>""", unsafe_allow_html=True)

st.markdown("<br><br><hr>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center;'>Tasas Referenciales Banco Central</h3>", unsafe_allow_html=True)

df_ref = cargar_tasas_referenciales()
if not df_ref.empty and len(df_ref.columns) >= 2:
    df_ref['fecha_dt'] = pd.to_datetime(df_ref.iloc[:, 0], dayfirst=True, errors='coerce')
    df_ref['tasa_val'] = limpiar_numeros(df_ref.iloc[:, 1])
    df_validos = df_ref.dropna(subset=['fecha_dt', 'tasa_val']).sort_values('fecha_dt')

    if not df_validos.empty:
        fecha_actual = datetime.datetime.now()
        df_mes_actual = df_validos[(df_validos['fecha_dt'].dt.year == fecha_actual.year) & (df_validos['fecha_dt'].dt.month == fecha_actual.month)]
        
        tasa_actual = df_mes_actual.iloc[-1]['tasa_val'] if not df_mes_actual.empty else df_validos.iloc[-1]['tasa_val']
        fecha_kpi = df_mes_actual.iloc[-1]['fecha_dt'] if not df_mes_actual.empty else df_validos.iloc[-1]['fecha_dt']

        meses_nombres = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        etiqueta_mes = f"{meses_nombres[fecha_kpi.month - 1]} {fecha_kpi.year}"
        tasa_promedio_anio = df_validos[df_validos['fecha_dt'].dt.year == fecha_kpi.year]['tasa_val'].mean()

        st.markdown(f"""
        <div class="kpi-flex-container">
            <div class="kpi-card-compact"><div class="kpi-title-compact">Referencial Actual<br>({etiqueta_mes})</div><div class="kpi-value-compact" style="color: #22c55e;">{tasa_actual:.2f}%</div></div>
            <div class="kpi-card-compact"><div class="kpi-title-compact">Promedio Anual<br>({fecha_kpi.year})</div><div class="kpi-value-compact" style="color: #38bdf8;">{tasa_promedio_anio:.2f}%</div></div>
        </div><br>
        """, unsafe_allow_html=True)

        df_validos['anio'] = df_validos['fecha_dt'].dt.year
        anios_unicos = sorted(df_validos['anio'].unique(), reverse=True)
        opciones_anio = ["Todos"] + [str(a) for a in anios_unicos]

        anio_seleccionado = st.selectbox("📅 Filtrar por año:", opciones_anio, index=0)

        if anio_seleccionado != "Todos":
            df_chart = df_validos[df_validos['anio'] == int(anio_seleccionado)].copy()
        else:
            df_chart = df_validos.copy()

        fig = px.line(
            df_chart,
            x='fecha_dt',
            y='tasa_val',
            labels={'fecha_dt': 'Fecha', 'tasa_val': 'Tasa (%)'}
        )
        fig.update_traces(line_color='#22c55e', line_width=2.5)
        fig.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#FFFFFF'),
            margin=dict(l=10, r=10, t=10, b=10),
            xaxis=dict(showgrid=True, gridcolor='#334155', title=None),
            yaxis=dict(showgrid=True, gridcolor='#334155', title='Tasa (%)'),
            dragmode=False
        )
        st.plotly_chart(
            fig,
            use_container_width=True,
            config={
                'scrollZoom': False,
                'displayModeBar': False,
                'doubleClick': False,
                'showAxisDragHandles': False,
                'staticPlot': False
            }
        )

        if not df_chart.empty:
            max_idx = df_chart['tasa_val'].idxmax()
            min_idx = df_chart['tasa_val'].idxmin()
            
            row_max = df_chart.loc[max_idx]
            row_min = df_chart.loc[min_idx]
            
            tasa_max = row_max['tasa_val']
            fecha_max = f"{meses_nombres[row_max['fecha_dt'].month - 1]} {row_max['fecha_dt'].year}"
            
            tasa_min = row_min['tasa_val']
            fecha_min = f"{meses_nombres[row_min['fecha_dt'].month - 1]} {row_min['fecha_dt'].year}"
            
            tasa_prom = df_chart['tasa_val'].mean()
            
            st.markdown(f"""
            <div class="kpi-flex-container">
                <div class="kpi-card-compact">
                    <div class="kpi-title-compact">Pico Máximo</div>
                    <div class="kpi-value-compact" style="color: #22c55e;">{tasa_max:.2f}%</div>
                    <div style="font-size: 11px; color: #94A3B8; margin-top: 2px;">{fecha_max}</div>
                </div>
                <div class="kpi-card-compact">
                    <div class="kpi-title-compact">Pico Mínimo</div>
                    <div class="kpi-value-compact" style="color: #ef4444;">{tasa_min:.2f}%</div>
                    <div style="font-size: 11px; color: #94A3B8; margin-top: 2px;">{fecha_min}</div>
                </div>
                <div class="kpi-card-compact">
                    <div class="kpi-title-compact">Promedio Periodo</div>
                    <div class="kpi-value-compact" style="color: #38bdf8;">{tasa_prom:.2f}%</div>
                    <div style="font-size: 11px; color: #94A3B8; margin-top: 2px;">Filtro: {anio_seleccionado}</div>
                </div>
            </div><br>
            """, unsafe_allow_html=True)

            df_mensual = df_chart.copy()
            df_mensual['year_month'] = df_mensual['fecha_dt'].dt.strftime('%Y-%m')
            df_serie = df_mensual.groupby('year_month')['tasa_val'].last().values
            
            if len(df_serie) >= 3:
                y = df_serie[-12:]
                n = len(y)
                x = np.arange(1, n + 1)
                
                slope, intercept = np.polyfit(x, y, 1)
                
                pred_1 = slope * (n + 1) + intercept
                pred_3 = slope * (n + 3) + intercept
                
                residuals = y - (slope * x + intercept)
                se = np.sqrt(np.sum(residuals**2) / (n - 2)) if n > 2 else 0
                ci_95 = 1.96 * se
                margen_1 = max(0.05, ci_95)
                
                tasa_1m = max(0, pred_1)
                rango_1m = f"[{max(0, tasa_1m - margen_1):.2f}% - {tasa_1m + margen_1:.2f}%]"
                
                tasa_3m = max(0, pred_3)
                tendencia_str = "↗️ Alza" if slope > 0.015 else ("↘️ Baja" if slope < -0.015 else "➡️ Estable")
                
                diff_1m = tasa_1m - y[-1]
                icono_1m = "↗️" if diff_1m > 0.015 else ("↘️" if diff_1m < -0.015 else "➡️")
                
                rango_3m = f"[{max(0, min(tasa_1m, tasa_3m) - margen_1):.2f}% - {max(tasa_1m, tasa_3m) + margen_1:.2f}%]"
                
                st.markdown(f"""
                <div class="kpi-flex-container">
                    <div class="kpi-card-compact">
                        <div class="kpi-title-compact">Proyección Próximo Mes (t+1)</div>
                        <div class="kpi-value-compact" style="color: #a855f7;">{icono_1m} {tasa_1m:.2f}%</div>
                        <div style="font-size: 11px; color: #94A3B8; margin-top: 2px;">IC 95%: {rango_1m}</div>
                    </div>
                    <div class="kpi-card-compact">
                        <div class="kpi-title-compact">Tendencia 3 Meses (t+3)</div>
                        <div class="kpi-value-compact" style="color: #a855f7;">{tendencia_str}</div>
                        <div style="font-size: 11px; color: #94A3B8; margin-top: 2px;">Rango Esperado: {rango_3m}</div>
                    </div>
                </div><br>
                """, unsafe_allow_html=True)