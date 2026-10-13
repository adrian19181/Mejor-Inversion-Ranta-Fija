import sys
import io
import datetime
import requests
import pandas as pd
import numpy as np
import streamlit as st
import streamlit.components.v1 as components

if __name__ == "__main__" and not st.runtime.exists():
    from streamlit.web import cli as stcli
    sys.argv = ["streamlit", "run", __file__]
    sys.exit(stcli.main())

st.set_page_config(page_title="Mejor Inversión Renta Fija", page_icon="💰", layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    h1, h2, h3, h4, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { text-align: center !important; }
    button[data-baseweb="tab"] { font-size: 16px !important; font-weight: 600 !important; }
    [data-testid="stInputInstruction"], small[data-testid="stInputInstruction"], div[data-testid="InputInstruction"], div[data-baseweb="typo-caption"] { display: none !important; }
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
    div.row-widget.stRadio > div { flex-direction: column; gap: 16px; padding: 10px 0px; }
    div.row-widget.stRadio > div > label { background-color: #1E293B; padding: 20px 15px; border-radius: 12px; cursor: pointer; transition: all 0.2s ease; border: 2px solid #334155; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    div.row-widget.stRadio > div > label:hover { background-color: #243248; border-color: #3b82f6; }
    div.row-widget.stRadio > div > label p { font-size: 17px !important; font-weight: 700 !important; color: #F8FAFC !important; }
</style>
""", unsafe_allow_html=True)

FILE_ID = "1DPASEKWo7lTbu8tug041iK5lMXLlyJvG"
DRIVE_URL = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"
LOCAL_BACKUP_PATH = "Inversiones_Ecuador.xlsx"

def limpiar_numeros(serie):
    s = pd.Series(serie).astype(str).str.replace(',', '.', regex=False)
    s = s.str.replace('%', '', regex=False).str.replace('$', '', regex=False).str.strip()
    s = s.replace(['nan', 'NAN', 'None', ''], np.nan)
    return pd.to_numeric(s, errors='coerce')

def safe_get(df, cols, default_val):
    for c in cols:
        if c in df.columns: return df[c]
    return pd.Series([default_val]*len(df))

@st.cache_data(ttl=3600)
def cargar_datos_inversiones():
    content = None
    try:
        response = requests.get(DRIVE_URL, timeout=10)
        if response.status_code == 200: content = io.BytesIO(response.content)
    except Exception: pass

    xl = pd.ExcelFile(content) if content is not None else (pd.ExcelFile(LOCAL_BACKUP_PATH) if pd.io.common.file_exists(LOCAL_BACKUP_PATH) else None)
    if xl is None: return pd.DataFrame(), pd.DataFrame()

    tablas_pf = []
    tablas_ap = []

    for sheet in xl.sheet_names:
        if sheet == "Tasas Referenciales": continue
        df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        if len(df_raw) < 3: continue
        
        banco_norm = sheet.upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
        row0 = df_raw.iloc[0].fillna('').astype(str).str.strip().str.upper()
        row1 = df_raw.iloc[1].fillna('').astype(str).str.strip()
        inicios = [(i, val) for i, val in enumerate(row0) if val != '' and val != 'NAN' and not val.startswith('UNNAMED')]
        
        for idx in range(len(inicios)):
            start_col, title = inicios[idx]
            end_col = inicios[idx+1][0] if idx + 1 < len(inicios) else len(row1)
            
            headers = row1[start_col:end_col].tolist()
            valid_cols_idx = [j for j, h in enumerate(headers) if h != '' and h.upper() != 'NAN']
            
            es_ap = False
            kw_ap = ["PROGRAMADO", "CHANCHITO", "FONDO AHORRO", "AHORRO FONDO", "FUTURO SEGURO", "FLEXI"]
            if any(kw in title for kw in kw_ap) or ("AHORRO" in title and "VISTA" not in title and "CUENTA" not in title and "PLUS" not in title):
                es_ap = True
                
            if es_ap:
                cols_idx = []
                grupo_nombre = "🟢 Grupo 1: Tradicionales (Jardín Azuayo, Machala, 29 Oct, Produbanco, JEP)"
                
                if "JARDIN AZUAYO" in banco_norm:
                    cols_idx = [5, 6, 7] # F, G, H
                elif "MACHALA" in banco_norm:
                    cols_idx = [7, 8, 9] # H, I, J
                elif "29 DE OCTUBRE" in banco_norm:
                    cols_idx = [8, 9, 10] # I, J, K
                elif "PRODUBANCO" in banco_norm:
                    cols_idx = [7, 8, 9] # H, I, J
                elif "JEP" in banco_norm:
                    cols_idx = [7, 8, 9] # H, I, J
                elif "AUSTRO" in banco_norm:
                    grupo_nombre = "🔵 Grupo 2: Banco del Austro"
                elif "COOPERCO" in banco_norm:
                    grupo_nombre = "🟡 Grupo 3: Cooperco"
                elif "SOLIDARIO" in banco_norm:
                    grupo_nombre = "🟣 Grupo 4: Banco Solidario"

                if cols_idx:
                    max_col_needed = max(cols_idx)
                    if df_raw.shape[1] <= max_col_needed:
                        df_raw = df_raw.reindex(columns=range(max_col_needed + 1))
                        
                    df_temp = df_raw.iloc[2:, cols_idx].copy()
                    df_temp.columns = ['dias_minimo', 'dias_maximo', 'tna']
                    df_temp = df_temp.dropna(how='all')
                    
                    df_ap = pd.DataFrame()
                    df_ap['banco'] = [sheet] * len(df_temp)
                    df_ap['producto'] = [title] * len(df_temp)
                    df_ap['tipo_ahorro'] = grupo_nombre
                    df_ap['monto_minimo'] = 0.0
                    df_ap['monto_maximo'] = 99999999.0
                    
                    p_max = limpiar_numeros(df_temp['dias_maximo']).fillna(9999).astype(int)
                    p_min_raw = limpiar_numeros(df_temp['dias_minimo'])
                    if p_min_raw.isna().all():
                        p_min_list = [1]
                        for i in range(len(p_max)-1):
                            p_min_list.append(p_max.iloc[i] + 1)
                        p_min = pd.Series(p_min_list, index=p_max.index).astype(int)
                    else:
                        p_min = p_min_raw.fillna(1).astype(int)
                    
                    if p_max.max() <= 120:
                        p_min = p_min * 30
                        p_max = p_max * 30
                        
                    df_ap['dias_minimo'] = p_min
                    df_ap['dias_maximo'] = p_max
                    df_ap['tna'] = limpiar_numeros(df_temp['tna']).fillna(0.0)
                    
                    df_ap = df_ap[df_ap['tna'] > 0]
                    if not df_ap.empty: tablas_ap.append(df_ap)
                else:
                    if not valid_cols_idx: continue
                    df_t = df_raw.iloc[2:, start_col:end_col].copy()
                    df_t.columns = headers
                    df_t = df_t.iloc[:, valid_cols_idx]
                    df_t = df_t.dropna(how='all')
                    for c in df_t.columns: df_t[c] = limpiar_numeros(df_t[c])
                        
                    df_ap = pd.DataFrame()
                    df_ap['banco'] = [sheet] * len(df_t)
                    df_ap['producto'] = [title] * len(df_t)
                    df_ap['tipo_ahorro'] = grupo_nombre
                    
                    df_ap['monto_minimo'] = limpiar_numeros(safe_get(df_t, ['monto_minimo', 'cuota_minima'], 0.0)).fillna(0.0)
                    df_ap['monto_maximo'] = limpiar_numeros(safe_get(df_t, ['monto_maximo', 'cuota_maxima'], 99999999.0)).fillna(99999999.0)
                    dmin = limpiar_numeros(safe_get(df_t, ['dias_minimo', 'plazo_minimo_dias'], 1)).fillna(1).astype(int)
                    dmax = limpiar_numeros(safe_get(df_t, ['dias_maximo', 'plazo_maximo_dias'], 9999)).fillna(9999).astype(int)
                    
                    if dmax.max() <= 120:
                        dmin = dmin * 30
                        dmax = dmax * 30
                        
                    df_ap['dias_minimo'] = dmin
                    df_ap['dias_maximo'] = dmax
                    df_ap['tna'] = limpiar_numeros(safe_get(df_t, ['tasa_nominal_porcentaje', 'tasa_anual_porcentaje'], 0.0)).fillna(0.0)
                    
                    df_ap = df_ap[df_ap['tna'] > 0]
                    if not df_ap.empty: tablas_ap.append(df_ap)
                        
            elif idx == 0:
                if not valid_cols_idx: continue
                df_t = df_raw.iloc[2:, start_col:end_col].copy()
                df_t.columns = headers
                df_t = df_t.iloc[:, valid_cols_idx]
                df_t = df_t.dropna(how='all')

                df_pf = pd.DataFrame()
                df_pf['banco'] = [sheet] * len(df_t)
                df_pf['monto_minimo'] = limpiar_numeros(safe_get(df_t, ['monto_minimo'], 0.0)).fillna(0.0)
                df_pf['monto_maximo'] = limpiar_numeros(safe_get(df_t, ['monto_maximo'], 99999999.0)).fillna(99999999.0)
                df_pf['dias_minimo'] = limpiar_numeros(safe_get(df_t, ['dias_minimo'], 1)).fillna(1).astype(int)
                df_pf['dias_maximo'] = limpiar_numeros(safe_get(df_t, ['dias_maximo'], 9999)).fillna(9999).astype(int)
                df_pf['tasa_nominal_porcentaje'] = limpiar_numeros(safe_get(df_t, ['tasa_nominal_porcentaje', 'tasa_pago_vencimiento_porcentaje'], 0.0)).fillna(0.0)
                
                df_pf = df_pf[df_pf['tasa_nominal_porcentaje'] > 0]
                if not df_pf.empty: tablas_pf.append(df_pf)

    df_final_pf = pd.concat(tablas_pf, ignore_index=True) if tablas_pf else pd.DataFrame()
    df_final_ap = pd.concat(tablas_ap, ignore_index=True) if tablas_ap else pd.DataFrame()
    return df_final_pf, df_final_ap

@st.cache_data(ttl=3600)
def cargar_tasas_referenciales():
    content = None
    try:
        response = requests.get(DRIVE_URL, timeout=5)
        if response.status_code == 200: content = io.BytesIO(response.content)
    except Exception: pass
    xl_source = content if content else LOCAL_BACKUP_PATH
    try:
        xl_doc = pd.ExcelFile(xl_source)
        if "Tasas Referenciales" in xl_doc.sheet_names: return pd.read_excel(xl_doc, sheet_name="Tasas Referenciales")
    except Exception: pass
    return pd.DataFrame()

def formatear_nombre_banco(nombre):
    palabras = nombre.strip().split()
    if len(palabras) > 1:
        mitad = (len(palabras) + 1) // 2
        return f"{' '.join(palabras[:mitad])}<br>{' '.join(palabras[mitad:])}"
    return nombre

st.markdown("<h1 style='text-align: center;'>Mejor Inversión Renta Fija</h1>", unsafe_allow_html=True)

col_btn1, col_btn2 = st.columns([3, 1])
with col_btn2:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

df_pf, df_ap = cargar_datos_inversiones()
if df_pf.empty:
    st.warning("No se encontraron datos de inversiones disponibles.")
    st.stop()

tab_pf, tab_ap = st.tabs(["💰 Plazo Fijo", "📅 Ahorro Programado"])

with tab_pf:
    st.markdown("<br><h3 style='text-align: center;'>📋 Parámetros de Inversión</h3>", unsafe_allow_html=True)
    with st.form(key="formulario_pf"):
        col1, col2 = st.columns(2)
        with col1: monto_input = st.number_input("Capital a Invertir ($ USD)", min_value=1.0, value=None, step=1.0)
        with col2: plazo_input = st.number_input("Plazo deseado (Días)", min_value=1, max_value=3600, value=None, step=1)
        submit_pf = st.form_submit_button(label="🎯 Calcular Mejores Opciones", use_container_width=True)

    monto = monto_input if monto_input is not None else 5000.0
    plazo_dias = plazo_input if plazo_input is not None else 180

    filtrado_pf = df_pf[(df_pf['monto_minimo'] <= monto) & (df_pf['monto_maximo'] >= monto) & 
                        (df_pf['dias_minimo'] <= plazo_dias) & (df_pf['dias_maximo'] >= plazo_dias)].copy()

    if filtrado_pf.empty:
        st.error(f"No hay opciones disponibles para un monto de ${monto:,.2f} a {plazo_dias} días.")
    else:
        filtrado_pf['ganancia_estimada'] = monto * (filtrado_pf['tasa_nominal_porcentaje'] / 100.0) * (plazo_dias / 360.0)
        top5 = filtrado_pf.sort_values(by='tasa_nominal_porcentaje', ascending=False).head(5).reset_index(drop=True)
        
        st.markdown("---")
        st.markdown("<h3 style='text-align: center;'>🏆 Mejor Opción Plazo Fijo</h3>", unsafe_allow_html=True)
        ganador = top5.iloc[0]
        st.success(f"🥇 **Opción Ganadora:** **{ganador['banco']}** con una TNA de **{ganador['tasa_nominal_porcentaje']:.2f}%** y una ganancia estimada de **${ganador['ganancia_estimada']:,.2f}**")

        html_rows = ""
        for idx, row in top5.iterrows():
            banco_fmt = formatear_nombre_banco(row['banco'])
            rango = f"{int(row['dias_minimo'])} - {int(row['dias_maximo'])} días"
            tna = f"{row['tasa_nominal_porcentaje']:.2f}%"
            ganancia = f"${row['ganancia_estimada']:,.2f}"
            html_rows += f"<tr><td><strong>{banco_fmt}</strong></td><td>{rango}</td><td><span style='color: #22c55e; font-weight: bold;'>{tna}</span></td><td><span style='color: #3b82f6; font-weight: bold;'>{ganancia}</span></td></tr>"

        st.markdown(f"""<div class="table-container"><table class="custom-table"><thead><tr><th>Institución<br>Financiera</th><th>Rango Plazo<br>Aplicado</th><th>TNA<br>(%)</th><th>Ganancia<br>Estimada ($)</th></tr></thead><tbody>{html_rows}</tbody></table></div>""", unsafe_allow_html=True)

with tab_ap:
    if df_ap.empty:
        st.info("No se encontraron productos de Ahorro Programado en la base de datos.")
    else:
        st.markdown("<br><h3 style='text-align: center;'>🔍 Seleccionar Grupo a Comparar</h3>", unsafe_allow_html=True)
        
        grupo_seleccionado = st.radio(
            "Filtrar por Instituciones con estructura similar:",
            [
                "🟢 Grupo 1: Tradicionales (Jardín Azuayo, Machala, 29 Oct, Produbanco, JEP)",
                "🔵 Grupo 2: Banco del Austro",
                "🟡 Grupo 3: Cooperco",
                "🟣 Grupo 4: Banco Solidario"
            ],
            label_visibility="collapsed"
        )
        
        st.markdown("<br><h3 style='text-align: center;'>📋 Parámetros de Ahorro</h3>", unsafe_allow_html=True)
        with st.form(key="formulario_ap"):
            col1, col2 = st.columns(2)
            with col1: cuota_input = st.number_input("Ahorro Mensual ($ USD)", min_value=1.0, value=None, step=1.0)
            with col2: dias_ap_input = st.number_input("Plazo a cumplir (Días)", min_value=1, max_value=3600, value=None, step=1)
            submit_ap = st.form_submit_button(label="🎯 Calcular Mejores Opciones", use_container_width=True)

        cuota = cuota_input if cuota_input is not None else 100.0
        dias_ap = dias_ap_input if dias_ap_input is not None else 360

        df_ap_grupo = df_ap[df_ap['tipo_ahorro'] == grupo_seleccionado].copy()

        filtrado_ap = df_ap_grupo[(df_ap_grupo['monto_minimo'] <= cuota) & (df_ap_grupo['monto_maximo'] >= cuota) & 
                                  (df_ap_grupo['dias_minimo'] <= dias_ap) & (df_ap_grupo['dias_maximo'] >= dias_ap)].copy()

        if filtrado_ap.empty:
            st.error(f"No hay opciones en este grupo para una cuota de ${cuota:,.2f} a {dias_ap} días.")
        else:
            top3_ap = filtrado_ap.sort_values(by='tna', ascending=False).head(3).reset_index(drop=True)
            
            st.markdown("---")
            st.markdown("<h3 style='text-align: center;'>🏆 Mejor Opción Ahorro Programado</h3>", unsafe_allow_html=True)
            ganador_ap = top3_ap.iloc[0]
            
            st.success(f"🥇 **Opción Ganadora:** **{ganador_ap['banco']} ({ganador_ap['producto'].title()})** con una TNA de **{ganador_ap['tna']:.2f}%**")

            html_rows_ap = ""
            for idx, row in top3_ap.iterrows():
                banco_fmt = formatear_nombre_banco(row['banco'])
                prod_fmt = f"<span style='font-size:11px; color:#94A3B8;'>{row['producto'].title()}</span>"
                institucion = f"{banco_fmt}<br>{prod_fmt}"
                rango_dias = f"{int(row['dias_minimo'])} - {int(row['dias_maximo'])} días"
                tna = f"{row['tna']:.2f}%"
                
                html_rows_ap += f"<tr><td><strong>{institucion}</strong></td><td>{rango_dias}</td><td><span style='color: #22c55e; font-weight: bold;'>{tna}</span></td></tr>"

            st.markdown(f"""<div class="table-container"><table class="custom-table"><thead><tr><th>Institución /<br>Producto</th><th>Rango de<br>Plazo</th><th>TNA<br>(%)</th></tr></thead><tbody>{html_rows_ap}</tbody></table></div>""", unsafe_allow_html=True)

st.markdown("<br><br><hr>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center;'>📊 Tasas Pasivas Referenciales (BCE)</h3>", unsafe_allow_html=True)

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