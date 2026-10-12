import sys
import io
import datetime
import requests
import pandas as pd
import numpy as np
import streamlit as st
import streamlit.components.v1 as components

# =========================================================================
# CONFIGURACIÓN Y AUTO-LANZAMIENTO LOCAL AL HACER DOBLE CLICK
# =========================================================================
if __name__ == "__main__" and not st.runtime.exists():
    from streamlit.web import cli as stcli
    sys.argv = ["streamlit", "run", __file__]
    sys.exit(stcli.main())

# Configuración de página
st.set_page_config(
    page_title="Mejor Inversión Renta Fija",
    page_icon="💰",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Estilos CSS personalizados
st.markdown("""
<style>
    /* Centrar todos los títulos y encabezados */
    h1, h2, h3, h4, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
        text-align: center !important;
    }

    /* Ocultar texto auxiliar e instrucciones ("Press Enter to submit form", etc.) */
    [data-testid="stInputInstruction"], 
    small[data-testid="stInputInstruction"], 
    div[data-testid="InputInstruction"],
    div[data-baseweb="typo-caption"] { 
        display: none !important; 
    }
    input::-webkit-input-placeholder, input::-moz-placeholder, input:-ms-input-placeholder, input::placeholder { 
        color: transparent !important; 
        opacity: 0 !important; 
    }

    /* Centrado de texto en las casillas del formulario */
    div[data-baseweb="input"] input { 
        text-align: center !important; 
        font-size: 18px !important; 
        font-weight: 600 !important; 
    }

    /* Ocultar botones de incremento y decremento (+/-) y flechas */
    div[data-testid="stNumberInputContainer"] button, 
    button[data-testid="stNumberInputStepDown"], 
    button[data-testid="stNumberInputStepUp"] { 
        display: none !important; 
    }
    input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { 
        -webkit-appearance: none !important; 
        margin: 0 !important; 
    }
    input[type=number] { 
        -moz-appearance: textfield !important; 
    }

    /* Contenedor principal de tabla con desplazamiento horizontal */
    .table-container { 
        width: 100%; 
        overflow-x: auto; 
        margin-top: 15px; 
        border-radius: 10px; 
        box-shadow: 0 2px 8px rgba(0,0,0,0.15); 
        background-color: #0F172A; 
    }
    .custom-table { 
        width: 100%; 
        border-collapse: separate; 
        border-spacing: 0; 
        table-layout: auto; 
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; 
    }
    .custom-table th { 
        background-color: #1E293B; 
        color: #FFFFFF; 
        text-align: center !important; 
        vertical-align: middle !important; 
        padding: 10px 8px; 
        font-size: 13px; 
        font-weight: 700; 
        line-height: 1.25; 
        white-space: nowrap; 
        border-bottom: 2px solid #334155; 
        border-top: 1px solid #334155; 
    }
    .custom-table td { 
        text-align: center !important; 
        vertical-align: middle !important; 
        padding: 10px 8px; 
        font-size: 14px; 
        white-space: nowrap; 
        border-bottom: 1px solid #334155; 
    }
    .custom-table th:nth-child(1) { 
        position: sticky; 
        left: 0; 
        z-index: 5; 
        background-color: #1E293B; 
        min-width: 110px; 
        max-width: 130px; 
        white-space: normal !important; 
        word-wrap: break-word !important; 
        border-right: 2px solid #64748B !important; 
    }
    .custom-table td:nth-child(1) { 
        position: sticky; 
        left: 0; 
        z-index: 3; 
        background-color: #0F172A; 
        min-width: 110px; 
        max-width: 130px; 
        white-space: normal !important; 
        word-wrap: break-word !important; 
        line-height: 1.2; 
        border-right: 2px solid #64748B !important; 
    }
    .custom-table tr:nth-child(even) td { background-color: #182234 !important; }
    .custom-table tr:hover td { background-color: #243248 !important; }

    /* Diseños de recuadros / tarjetas para KPIs optimizados para móvil lado a lado */
    .kpi-flex-container {
        display: flex;
        flex-direction: row;
        justify-content: space-between;
        gap: 8px;
        width: 100%;
        margin-top: 10px;
    }
    .kpi-card-compact {
        flex: 1;
        background-color: #1E293B;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 10px 4px;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.2);
    }
    .kpi-title-compact {
        color: #94A3B8;
        font-size: 11px;
        font-weight: 600;
        margin-bottom: 4px;
        line-height: 1.2;
    }
    .kpi-value-compact {
        font-size: 22px;
        font-weight: 800;
        line-height: 1.1;
    }
</style>
""", unsafe_allow_html=True)

# Inyección JS para autoselección
components.html("""
<script>
    const parentDoc = window.parent.document;
    function habilitarAutoSeleccion() {
        const inputs = parentDoc.querySelectorAll('input[type="number"]');
        inputs.forEach(input => {
            if (!input.dataset.autoselectBound) {
                input.dataset.autoselectBound = "true";
                const handler = function(e) { setTimeout(() => { try { e.target.select(); e.target.setSelectionRange(0, 9999); } catch (err) {} }, 50); };
                input.addEventListener('focus', handler, true);
                input.addEventListener('click', handler, true);
                input.addEventListener('touchstart', handler, true);
            }
        });
    }
    habilitarAutoSeleccion();
    const observer = new MutationObserver(habilitarAutoSeleccion);
    observer.observe(parentDoc.body, { childList: true, subtree: true });
</script>
""", height=0, width=0)

FILE_ID = "1DPASEKWo7lTbu8tug041iK5lMXLlyJvG"
DRIVE_URL = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"
LOCAL_BACKUP_PATH = "Inversiones_Ecuador.xlsx"

# =========================================================================
# CARGA DE DATOS BANCARIOS (Caché)
# =========================================================================
@st.cache_data(ttl=3600)
def cargar_datos_inversiones():
    content = None
    try:
        response = requests.get(DRIVE_URL, timeout=10)
        if response.status_code == 200:
            content = io.BytesIO(response.content)
    except Exception:
        pass

    if content is None:
        try:
            xl = pd.ExcelFile(LOCAL_BACKUP_PATH)
        except Exception:
            st.error("No se pudo descargar el archivo desde Google Drive ni encontrar la copia local.")
            return pd.DataFrame()
    else:
        xl = pd.ExcelFile(content)

    registros = []
    for sheet in xl.sheet_names:
        if sheet == "Tasas Referenciales":
            continue
        df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        if len(df_raw) < 3:
            continue
        cols_raw = df_raw.iloc[1].values
        plazo_fijo_cols, col_indices = [], []
        for idx, col_name in enumerate(cols_raw):
            if pd.isna(col_name) or str(col_name).strip() == '':
                break
            plazo_fijo_cols.append(str(col_name).strip())
            col_indices.append(idx)
        df_pf = df_raw.iloc[2:, col_indices].copy()
        df_pf.columns = plazo_fijo_cols
        df_pf = df_pf.dropna(how='all')
        for col in df_pf.columns:
            df_pf[col] = pd.to_numeric(df_pf[col], errors='coerce')
        if 'monto_minimo' not in df_pf.columns: df_pf['monto_minimo'] = 0.0
        if 'monto_maximo' not in df_pf.columns: df_pf['monto_maximo'] = 99999999.0
        if 'dias_minimo' not in df_pf.columns: df_pf['dias_minimo'] = 1
        if 'dias_maximo' not in df_pf.columns: df_pf['dias_maximo'] = 9999
        if 'tasa_nominal_porcentaje' not in df_pf.columns and 'tasa_pago_vencimiento_porcentaje' in df_pf.columns:
            df_pf['tasa_nominal_porcentaje'] = df_pf['tasa_pago_vencimiento_porcentaje']
        df_pf['banco'] = sheet
        df_pf['monto_minimo'] = df_pf['monto_minimo'].fillna(0.0)
        df_pf['monto_maximo'] = df_pf['monto_maximo'].fillna(99999999.0)
        df_pf['dias_minimo'] = df_pf['dias_minimo'].fillna(1).astype(int)
        df_pf['dias_maximo'] = df_pf['dias_maximo'].fillna(9999).astype(int)
        df_pf['tasa_nominal_porcentaje'] = df_pf['tasa_nominal_porcentaje'].fillna(0.0)
        registros.append(df_pf[['banco', 'monto_minimo', 'monto_maximo', 'dias_minimo', 'dias_maximo', 'tasa_nominal_porcentaje']])
    return pd.concat(registros, ignore_index=True) if registros else pd.DataFrame()

# =========================================================================
# CARGA DE TASAS REFERENCIALES
# =========================================================================
@st.cache_data(ttl=3600)
def cargar_tasas_referenciales():
    content = None
    try:
        response = requests.get(DRIVE_URL, timeout=5)
        if response.status_code == 200:
            content = io.BytesIO(response.content)
    except Exception:
        pass

    if content is not None:
        try:
            xl_drive = pd.ExcelFile(content)
            if "Tasas Referenciales" in xl_drive.sheet_names:
                return pd.read_excel(xl_drive, sheet_name="Tasas Referenciales")
        except Exception:
            pass

    try:
        xl_local = pd.ExcelFile(LOCAL_BACKUP_PATH)
        if "Tasas Referenciales" in xl_local.sheet_names:
            return pd.read_excel(xl_local, sheet_name="Tasas Referenciales")
    except Exception:
        pass

    return pd.DataFrame()

def formatear_nombre_banco(nombre):
    palabras = nombre.strip().split()
    if len(palabras) > 1:
        mitad = (len(palabras) + 1) // 2
        linea1 = " ".join(palabras[:mitad])
        linea2 = " ".join(palabras[mitad:])
        return f"{linea1}<br>{linea2}"
    return nombre

# =========================================================================
# INTERFAZ PRINCIPAL
# =========================================================================
st.markdown("<h1 style='text-align: center;'>Mejor Inversión Renta Fija</h1>", unsafe_allow_html=True)

col_btn1, col_btn2 = st.columns([3, 1])
with col_btn2:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

df_datos = cargar_datos_inversiones()
if df_datos.empty:
    st.warning("No se encontraron datos de inversiones disponibles.")
    st.stop()

st.markdown("<h3 style='text-align: center;'>📋 Parámetros de Inversión</h3>", unsafe_allow_html=True)

with st.form(key="formulario_inversion"):
    col1, col2 = st.columns(2)
    with col1:
        monto_input = st.number_input("Capital a Invertir ($ USD)", min_value=1.0, value=None, step=1.0)
    with col2:
        plazo_input = st.number_input("Plazo deseado (Días)", min_value=1, max_value=3600, value=None, step=1)
    submit_button = st.form_submit_button(label="🎯 Calcular Mejores Opciones", use_container_width=True)

monto = monto_input if monto_input is not None else 5000.0
plazo_dias = plazo_input if plazo_input is not None else 180

filtrado = df_datos[(df_datos['monto_minimo'] <= monto) & (df_datos['monto_maximo'] >= monto) & 
                    (df_datos['dias_minimo'] <= plazo_dias) & (df_datos['dias_maximo'] >= plazo_dias)].copy()

if filtrado.empty:
    st.error(f"No hay opciones disponibles para un monto de ${monto:,.2f} a {plazo_dias} días.")
else:
    filtrado['ganancia_estimada'] = monto * (filtrado['tasa_nominal_porcentaje'] / 100.0) * (plazo_dias / 360.0)
    top5 = filtrado.sort_values(by='tasa_nominal_porcentaje', ascending=False).head(5).reset_index(drop=True)
    
    st.markdown("---")
    st.markdown("<h3 style='text-align: center;'>🏆 Mejor Opción Plazo Fijo</h3>", unsafe_allow_html=True)
    ganador = top5.iloc[0]
    st.success(f"🥇 **Opción Ganadora:** **{ganador['banco']}** con una TNA de **{ganador['tasa_nominal_porcentaje']:.2f}%** y una ganancia estimada de **${ganador['ganancia_estimada']:,.2f}**")

    html_rows = ""
    for idx, row in top5.iterrows():
        banco_fmt = formatear_nombre_banco(row['banco'])
        rango = f"{row['dias_minimo']} - {row['dias_maximo']} días"
        tna = f"{row['tasa_nominal_porcentaje']:.2f}%"
        ganancia = f"${row['ganancia_estimada']:,.2f}"
        html_rows += f"<tr><td><strong>{banco_fmt}</strong></td><td>{rango}</td><td><span style='color: #22c55e; font-weight: bold;'>{tna}</span></td><td><span style='color: #3b82f6; font-weight: bold;'>{ganancia}</span></td></tr>"

    table_html = f"""<div class="table-container"><table class="custom-table"><thead><tr><th>Institución<br>Financiera</th><th>Rango Plazo<br>Aplicado</th><th>TNA<br>(%)</th><th>Ganancia<br>Estimada ($)</th></tr></thead><tbody>{html_rows}</tbody></table></div>"""
    st.markdown(table_html, unsafe_allow_html=True)

# =========================================================================
# KPIS DE TASAS REFERENCIALES (BCE) LADO A LADO COMPACTOS
# =========================================================================
df_ref = cargar_tasas_referenciales()

if not df_ref.empty and len(df_ref.columns) >= 2:
    df_ref['fecha_dt'] = pd.to_datetime(df_ref.iloc[:, 0], dayfirst=True, errors='coerce')
    tasas_limpias = df_ref.iloc[:, 1].astype(str).str.replace(',', '.').str.replace('%', '').str.strip()
    df_ref['tasa_val'] = pd.to_numeric(tasas_limpias, errors='coerce')
    
    df_validos = df_ref.dropna(subset=['fecha_dt', 'tasa_val']).copy()

    if not df_validos.empty:
        df_validos = df_validos.sort_values('fecha_dt')
        fecha_actual = datetime.datetime.now()
        
        anio_actual = fecha_actual.year
        mes_actual = fecha_actual.month
        df_mes_actual = df_validos[(df_validos['fecha_dt'].dt.year == anio_actual) & (df_validos['fecha_dt'].dt.month == mes_actual)]

        if not df_mes_actual.empty:
            tasa_actual = df_mes_actual.iloc[-1]['tasa_val']
            fecha_kpi = df_mes_actual.iloc[-1]['fecha_dt']
        else:
            tasa_actual = df_validos.iloc[-1]['tasa_val']
            fecha_kpi = df_validos.iloc[-1]['fecha_dt']

        meses_nombres = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        nombre_mes = meses_nombres[fecha_kpi.month - 1]
        etiqueta_mes_anio = f"{nombre_mes} {fecha_kpi.year}"

        df_anio_actual = df_validos[df_validos['fecha_dt'].dt.year == fecha_kpi.year]
        tasa_promedio_anio = df_anio_actual['tasa_val'].mean()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("<h3 style='text-align: center;'>📊 Tasas Pasivas Referenciales (BCE)</h3>", unsafe_allow_html=True)

        kpis_side_by_side = f"""
        <div class="kpi-flex-container">
            <div class="kpi-card-compact">
                <div class="kpi-title-compact">Referencial Actual<br>({etiqueta_mes_anio})</div>
                <div class="kpi-value-compact" style="color: #22c55e;">{tasa_actual:.2f}%</div>
            </div>
            <div class="kpi-card-compact">
                <div class="kpi-title-compact">Promedio Anual<br>({fecha_kpi.year})</div>
                <div class="kpi-value-compact" style="color: #38bdf8;">{tasa_promedio_anio:.2f}%</div>
            </div>
        </div>
        """
        st.markdown(kpis_side_by_side, unsafe_allow_html=True)
    else:
        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.error("📌 Error de lectura: No se pudo procesar la hoja de 'Tasas Referenciales'.")
else:
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.error("📌 Error de lectura: No se pudo procesar la hoja de 'Tasas Referenciales'.")