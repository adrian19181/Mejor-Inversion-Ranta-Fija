import sys
import io
import datetime
import requests
import pandas as pd
import numpy as np
import streamlit as st

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
    return (s == '').all() or (s == 'nan').all() or (s == 'none').all()

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
        
        banco_norm = str(sheet).upper()

        # Identificar bloques de tablas separados por columnas totalmente vacías
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

        # =========================================================
        # TABLA 1: PLAZO FIJO (PRIMER BLOQUE DE COLUMNAS)
        # =========================================================
        if len(bloques_tablas) >= 1:
            cols_pf = bloques_tablas[0]
            df_pf_raw = df_raw.iloc[:, cols_pf].copy()
            row1_pf = df_pf_raw.iloc[1].fillna('').astype(str).str.strip().str.lower().tolist()
            df_pf_data = df_pf_raw.iloc[2:].copy().reset_index(drop=True)
            df_pf_data.columns = [f"col_{i}_{h}" for i, h in enumerate(row1_pf)]

            mmin_col = next((c for c in df_pf_data.columns if 'monto_min' in c), None)
            mmax_col = next((c for c in df_pf_data.columns if 'monto_max' in c), None)
            dmin_col = next((c for c in df_pf_data.columns if 'dias_min' in c or 'plazo_min' in c), None)
            dmax_col = next((c for c in df_pf_data.columns if 'dias_max' in c or 'plazo_max' in c), None)
            tna_col = next((c for c in df_pf_data.columns if 'tasa_nominal' in c or 'tasa_pago' in c or 'tna' in c), None)

            if dmin_col is None and len(df_pf_data.columns) >= 2:
                dmin_col = df_pf_data.columns[0] if len(df_pf_data.columns) == 3 else (df_pf_data.columns[2] if len(df_pf_data.columns) >= 4 else df_pf_data.columns[0])
            if tna_col is None and len(df_pf_data.columns) >= 1:
                tna_col = df_pf_data.columns[-1]

            if dmin_col and tna_col:
                mmin = limpiar_numeros(df_pf_data[mmin_col]).fillna(0.0) if mmin_col else 0.0
                mmax = limpiar_numeros(df_pf_data[mmax_col]).fillna(99999999.0) if mmax_col else 99999999.0
                dmin = limpiar_numeros(df_pf_data[dmin_col]).fillna(1).astype(int)
                dmax = limpiar_numeros(df_pf_data[dmax_col]).fillna(9999).astype(int) if dmax_col else dmin
                tna = limpiar_numeros(df_pf_data[tna_col]).fillna(0.0)

                temp_pf = pd.DataFrame({
                    'banco': sheet,
                    'monto_minimo': mmin,
                    'monto_maximo': mmax,
                    'dias_minimo': dmin,
                    'dias_maximo': dmax,
                    'tasa_nominal_porcentaje': tna
                })
                temp_pf = temp_pf[temp_pf['tasa_nominal_porcentaje'] > 0]
                if not temp_pf.empty:
                    tablas_pf.append(temp_pf)

        # =========================================================
        # TABLA 2: AHORRO PROGRAMADO (SEGUNDO BLOQUE DE COLUMNAS)
        # =========================================================
        if len(bloques_tablas) >= 2:
            cols_ap = bloques_tablas[1]
            df_ap_raw = df_raw.iloc[:, cols_ap].copy()
            row1_ap = df_ap_raw.iloc[1].fillna('').astype(str).str.strip().str.lower().tolist()
            df_ap_data = df_ap_raw.iloc[2:].copy().reset_index(drop=True)
            df_ap_data.columns = [f"col_{i}_{h}" for i, h in enumerate(row1_ap)]

            dmin_col = next((c for c in df_ap_data.columns if 'dias_min' in c or 'plazo_min' in c), None)
            dmax_col = next((c for c in df_ap_data.columns if 'dias_max' in c or 'plazo_max' in c), None)
            tna_col = next((c for c in df_ap_data.columns if 'tna' in c or 'tasa' in c or 'tasa_nominal' in c), None)

            if dmin_col is None and len(df_ap_data.columns) >= 1: dmin_col = df_ap_data.columns[0]
            if dmax_col is None and len(df_ap_data.columns) >= 2: dmax_col = df_ap_data.columns[1]
            if tna_col is None and len(df_ap_data.columns) >= 1: tna_col = df_ap_data.columns[-1]

            if dmin_col and tna_col:
                dmin = limpiar_numeros(df_ap_data[dmin_col]).fillna(1).astype(int)
                dmax = limpiar_numeros(df_ap_data[dmax_col]).fillna(9999).astype(int) if dmax_col else dmin
                tna = limpiar_numeros(df_ap_data[tna_col]).fillna(0.0)

                prod_name = "Ahorro Programado"
                if "JARDIN AZUAYO" in banco_norm: prod_name = "Mi Ahorro Chanchito"
                elif "PRODUBANCO" in banco_norm: prod_name = "Flexi Ahorro"
                elif "JEP" in banco_norm: prod_name = "Ahorro Fondo JEP"
                elif "MACHALA" in banco_norm: prod_name = "Ahorro Programado"

                temp_ap = pd.DataFrame({
                    'banco': sheet,
                    'producto': prod_name,
                    'tipo_ahorro': "🟢 Grupo 1: Tradicionales (Jardín Azuayo, Machala, 29 Oct, Produbanco, JEP)",
                    'dias_minimo': dmin,
                    'dias_maximo': dmax,
                    'tna': tna
                })
                temp_ap = temp_ap[temp_ap['tna'] > 0]
                if not temp_ap.empty:
                    tablas_ap.append(temp_ap)

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

st.markdown("<h1 style='text-align: center;'>Mejor Inversión Renta Fija</h1>", unsafe_allow_html=True)

col_btn1, col_btn2 = st.columns([3, 1])
with col_btn2:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

df_pf, df_ap = cargar_datos_inversiones()
if df_pf.empty and df_ap.empty:
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

    if not df_pf.empty:
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
            ["🟢 Grupo 1: Tradicionales (Jardín Azuayo, Machala, 29 Oct, Produbanco, JEP)"],
            label_visibility="collapsed"
        )
        
        st.markdown("<br><h3 style='text-align: center;'>📋 Parámetros de Ahorro</h3>", unsafe_allow_html=True)
        with st.form(key="formulario_ap"):
            dias_ap_input = st.number_input("Plazo a cumplir (Días)", min_value=1, max_value=3600, value=None, step=1)
            submit_ap = st.form_submit_button(label="🎯 Calcular Mejores Opciones", use_container_width=True)

        dias_ap = dias_ap_input if dias_ap_input is not None else 360
        df_ap_grupo = df_ap[df_ap['tipo_ahorro'] == grupo_seleccionado].copy()
        
        if df_ap_grupo.empty:
            st.error("No hay datos cargados para este grupo.")
        else:
            resultados_finales = []
            
            for banco in df_ap_grupo['banco'].unique():
                df_banco = df_ap_grupo[df_ap_grupo['banco'] == banco].copy()
                
                exact_match = df_banco[(df_banco['dias_minimo'] <= dias_ap) & (df_banco['dias_maximo'] >= dias_ap)]
                
                if not exact_match.empty:
                    mejor_opcion = exact_match.sort_values(by='tna', ascending=False).iloc[0]
                    resultados_finales.append(mejor_opcion)
                else:
                    df_banco['distancia'] = df_banco.apply(
                        lambda row: min(abs(dias_ap - row['dias_minimo']), abs(dias_ap - row['dias_maximo'])), axis=1
                    )
                    mejor_opcion = df_banco.sort_values(by=['distancia', 'tna'], ascending=[True, False]).iloc[0]
                    resultados_finales.append(mejor_opcion)

            if resultados_finales:
                resultados_ap = pd.DataFrame(resultados_finales).sort_values(by='tna', ascending=False).reset_index(drop=True)
                
                st.markdown("---")
                st.markdown("<h3 style='text-align: center;'>🏆 Mejores Opciones Ahorro Programado</h3>", unsafe_allow_html=True)
                ganador_ap = resultados_ap.iloc[0]
                
                st.success(f"🥇 **Opción Ganadora:** **{ganador_ap['banco']} ({str(ganador_ap['producto']).title()})** con una TNA de **{ganador_ap['tna']:.2f}%**")

                html_rows_ap = ""
                for idx, row in resultados_ap.iterrows():
                    banco_fmt = formatear_nombre_banco(row['banco'])
                    prod_fmt = f"<span style='font-size:11px; color:#94A3B8;'>{str(row['producto']).title()}</span>"
                    institucion = f"{banco_fmt}<br>{prod_fmt}"
                    
                    es_exacto = (row['dias_minimo'] <= dias_ap <= row['dias_maximo'])
                    
                    if row['dias_minimo'] == row['dias_maximo']:
                        rango_dias = f"{int(row['dias_minimo'])} días"
                    else:
                        rango_dias = f"{int(row['dias_minimo'])} - {int(row['dias_maximo'])} días"
                        
                    if not es_exacto:
                        rango_dias += f"<br><span style='font-size:10px; color:#ef4444;'>(Más cercano)</span>"
                        
                    tna = f"{row['tna']:.2f}%"
                    
                    html_rows_ap += f"<tr><td><strong>{institucion}</strong></td><td>{rango_dias}</td><td><span style='color: #22c55e; font-weight: bold;'>{tna}</span></td></tr>"

                st.markdown(f"""<div class="table-container"><table class="custom-table"><thead><tr><th>Institución /<br>Producto</th><th>Rango de<br>Plazo Aplicado</th><th>TNA<br>(%)</th></tr></thead><tbody>{html_rows_ap}</tbody></table></div>""", unsafe_allow_html=True)

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