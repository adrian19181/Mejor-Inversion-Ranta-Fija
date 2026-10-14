import sys
import io
import datetime
import os
import unicodedata
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
</style>
""", unsafe_allow_html=True)

FILE_ID = "1DPASEKWo7lTbu8tug041iK5lMXLlyJvG"
DRIVE_URL = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"
LOCAL_BACKUP_PATH = "Inversiones_Ecuador.xlsx"

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

def obtener_ranking_bancos(df, monto, dias, tipo='pf'):
    """Evalúa las condiciones por banco de forma aislada. Excluye entradas <= 0 o que no cumplan el monto y plazo mínimo."""
    if df.empty or monto <= 0 or dias <= 0:
        return pd.DataFrame()
        
    resultados = []
    
    for banco in df['banco'].unique():
        df_banco = df[df['banco'] == banco].copy()
        banco_upper = str(banco).upper()
        
        # LÓGICA AISLADA BANCO SOLIDARIO: AHORRO PROGRAMADO
        if tipo == 'ap' and "SOLIDARIO" in banco_upper:
            df_monto_match = df_banco[(df_banco['monto_minimo'].isna() | (df_banco['monto_minimo'] <= monto)) & 
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
                    best_bono_row = exact_bono.sort_values(by='bono', ascending=False).iloc[0]
                    applied_bono = float(best_bono_row['bono'])
                    applied_dmin = int(best_bono_row['dias_minimo'])
                    applied_dmax = int(best_bono_row['dias_maximo'])
                    es_exacto = True
                else:
                    df_bonos['diff_dias'] = df_bonos['dias_minimo'] - dias
                    df_tol = df_bonos[(df_bonos['diff_dias'] > 0) & (df_bonos['diff_dias'] <= 10)]
                    if not df_tol.empty:
                        best_tol_row = df_tol.sort_values(by='diff_dias', ascending=True).iloc[0]
                        applied_bono = float(best_tol_row['bono'])
                        applied_dmin = int(best_tol_row['dias_minimo'])
                        applied_dmax = int(best_tol_row['dias_maximo'])
                        es_exacto = False
                    else:
                        applied_dmin = 0
                        applied_dmax = 9999
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

        # LÓGICA ESTÁNDAR PARA EL RESTO DE BANCOS (Y PLAZO FIJO)
        mask_monto = (df_banco['monto_minimo'].isna() | (df_banco['monto_minimo'] <= monto)) & \
                     (df_banco['monto_maximo'].isna() | (df_banco['monto_maximo'] >= monto))
        df_monto_valido = df_banco[mask_monto].copy()
        
        if df_monto_valido.empty:
            continue
        
        # OBLIGATORIEDAD DE PLAZO MÍNIMO: Si el plazo ingresado es menor al plazo mínimo exigido, se descarta
        min_dias_banco = df_monto_valido['dias_minimo'].min()
        if pd.notna(min_dias_banco) and dias < min_dias_banco:
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
        if tipo == 'pf':
            res_df['ganancia_estimada'] = monto * (res_df['tna'] / 100.0) * (dias / 360.0)
        return res_df
        
    return pd.DataFrame()

@st.cache_data(ttl=3600)
def cargar_datos_inversiones():
    content = None
    try:
        response = requests.get(DRIVE_URL, timeout=15)
        if response.status_code == 200: content = io.BytesIO(response.content)
    except Exception: pass

    xl = None
    if content is not None:
        try:
            xl = pd.ExcelFile(content)
        except Exception: pass

    if xl is None and os.path.exists(LOCAL_BACKUP_PATH):
        try:
            xl = pd.ExcelFile(LOCAL_BACKUP_PATH)
        except Exception: pass

    if xl is None:
        return pd.DataFrame(), pd.DataFrame()

    tablas_pf = []
    tablas_ap = []

    BANCOS_AP_VALIDOS = [
        "AUSTRO", "JARDIN AZUAYO", "MACHALA", 
        "PRODUBANCO", "29 DE OCTUBRE", "COOPERCO", "JEP", "SOLIDARIO"
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

            # BLOQUE 1: PLAZO FIJO
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

            # BLOQUE 2: AHORRO PROGRAMADO
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
        except Exception:
            continue

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
    
    xl_source = content if content is not None else LOCAL_BACKUP_PATH
    try:
        is_valid_source = isinstance(xl_source, io.BytesIO) or (isinstance(xl_source, str) and os.path.exists(xl_source))
        if is_valid_source:
            xl_doc = pd.ExcelFile(xl_source)
            if "Tasas Referenciales" in xl_doc.sheet_names: 
                return pd.read_excel(xl_doc, sheet_name="Tasas Referenciales")
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

                    html_rows = ""
                    for idx, row in resultados_pf.iterrows():
                        banco_fmt = formatear_nombre_banco(row['banco'])
                        
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

                        ganancia = f"${row['ganancia_estimada']:,.2f}"
                        html_rows += f"<tr><td><strong>{banco_fmt}</strong></td><td>{rango}</td><td><span style='color: #22c55e; font-weight: bold;'>{tna}</span></td><td><span style='color: #3b82f6; font-weight: bold;'>{ganancia}</span></td></tr>"

                    st.markdown(f"""<div class="table-container"><table class="custom-table"><thead><tr><th>Institución<br>Financiera</th><th>Rango Plazo<br>Aplicado</th><th>TNA<br>(%)</th><th>Ganancia<br>Estimada ($)</th></tr></thead><tbody>{html_rows}</tbody></table></div>""", unsafe_allow_html=True)

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

                    html_rows_ap = ""
                    for idx, row in resultados_ap.iterrows():
                        banco_fmt = formatear_nombre_banco(row['banco'])
                        prod_fmt = f"<span style='font-size:11px; color:#94A3B8;'>{str(row['producto']).title()}</span>"
                        institucion = f"{banco_fmt}<br>{prod_fmt}"
                        
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