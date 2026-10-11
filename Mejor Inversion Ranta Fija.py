import sys
import io
import requests
import pandas as pd
import numpy as np
import streamlit as st

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

# Estilos CSS personalizados para limpiar casillas, auto-seleccionar texto y congelar columna 1
st.markdown("""
<style>
    /* Ocultar texto auxiliar/instrucciones dentro o debajo de las casillas ("Press Enter to submit form", etc.) */
    [data-testid="stInputInstruction"], 
    small[data-testid="stInputInstruction"], 
    div[data-testid="InputInstruction"],
    div[data-baseweb="typo-caption"] {
        display: none !important;
    }

    /* Auto-selección visual de texto */
    div[data-baseweb="input"] input {
        text-align: center !important;
        font-size: 18px !important;
        font-weight: 600 !important;
        user-select: all !important;
        -webkit-user-select: all !important;
    }

    /* Ocultar botones + y - propios de Streamlit */
    div[data-testid="stNumberInputContainer"] button,
    button[data-testid="stNumberInputStepDown"],
    button[data-testid="stNumberInputStepUp"] {
        display: none !important;
    }

    /* Ocultar botones de flechas nativos del navegador */
    input[type=number]::-webkit-inner-spin-button, 
    input[type=number]::-webkit-outer-spin-button { 
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
        padding: 10px 12px;
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
        padding: 12px 10px;
        font-size: 14px;
        white-space: nowrap;
        border-bottom: 1px solid #334155;
    }

    /* --- CONGELAR ÚNICAMENTE LA COLUMNA 1 (Institución Financiera) --- */
    .custom-table th:nth-child(1) {
        position: sticky;
        left: 0;
        z-index: 5;
        background-color: #1E293B;
        min-width: 150px;
        border-right: 2px solid #64748B !important;
    }
    .custom-table td:nth-child(1) {
        position: sticky;
        left: 0;
        z-index: 3;
        background-color: #0F172A;
        min-width: 150px;
        border-right: 2px solid #64748B !important;
    }

    /* Mantenimiento del fondo opaco al alternar filas y hacer hover */
    .custom-table tr:nth-child(even) td {
        background-color: #182234 !important;
    }
    .custom-table tr:hover td {
        background-color: #243248 !important;
    }
</style>

<script>
    // JS para forzar la selección total del texto al tocar o enfocar la casilla en móvil o PC
    (function() {
        function autoSelectInput(e) {
            if (e.target && e.target.tagName === 'INPUT') {
                setTimeout(function() {
                    try {
                        e.target.select();
                        e.target.setSelectionRange(0, 9999);
                    } catch(err) {}
                }, 50);
            }
        }
        document.addEventListener('focusin', autoSelectInput, true);
        document.addEventListener('click', autoSelectInput, true);
        document.addEventListener('touchstart', autoSelectInput, true);
    })();
</script>
""", unsafe_allow_html=True)

# Enlace de exportación de Google Drive
FILE_ID = "1JlrtCIMfv9H0kRDsruZVwOTBKS5_6eQP"
DRIVE_URL = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"
LOCAL_BACKUP_PATH = "Inversiones_Ecuador.xlsx"

# =========================================================================
# FUNCIÓN DE CARGA Y NORMALIZACIÓN DE DATOS (CON CACHÉ)
# =========================================================================
@st.cache_data(ttl=3600)
def cargar_datos_inversiones():
    content = None
    try:
        response = requests.get(DRIVE_URL, timeout=10)
        if response.status_code == 200:
            content = io.BytesIO(response.content)
    except Exception:
        content = None

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
        df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        if len(df_raw) < 3:
            continue

        cols_raw = df_raw.iloc[1].values
        
        plazo_fijo_cols = []
        col_indices = []
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

        if 'monto_minimo' not in df_pf.columns:
            df_pf['monto_minimo'] = 0.0
        if 'monto_maximo' not in df_pf.columns:
            df_pf['monto_maximo'] = 99999999.0
        if 'dias_minimo' not in df_pf.columns:
            df_pf['dias_minimo'] = 1
        if 'dias_maximo' not in df_pf.columns:
            df_pf['dias_maximo'] = 9999

        if 'tasa_nominal_porcentaje' not in df_pf.columns and 'tasa_pago_vencimiento_porcentaje' in df_pf.columns:
            df_pf['tasa_nominal_porcentaje'] = df_pf['tasa_pago_vencimiento_porcentaje']

        df_pf['banco'] = sheet
        df_pf['monto_minimo'] = df_pf['monto_minimo'].fillna(0.0)
        df_pf['monto_maximo'] = df_pf['monto_maximo'].fillna(99999999.0)
        df_pf['dias_minimo'] = df_pf['dias_minimo'].fillna(1).astype(int)
        df_pf['dias_maximo'] = df_pf['dias_maximo'].fillna(9999).astype(int)
        df_pf['tasa_nominal_porcentaje'] = df_pf['tasa_nominal_porcentaje'].fillna(0.0)

        registros.append(df_pf[['banco', 'monto_minimo', 'monto_maximo', 'dias_minimo', 'dias_maximo', 'tasa_nominal_porcentaje']])

    if registros:
        return pd.concat(registros, ignore_index=True)
    return pd.DataFrame()

# =========================================================================
# INTERFAZ PRINCIPAL DE STREAMLIT
# =========================================================================
st.title("Mejor Inversión Renta Fija")

col_btn1, col_btn2 = st.columns([3, 1])
with col_btn2:
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

df_datos = cargar_datos_inversiones()

if df_datos.empty:
    st.warning("No se encontraron datos de inversiones disponibles.")
    st.stop()

st.subheader("📋 Parámetros de Inversión")

with st.form(key="formulario_inversion"):
    col1, col2 = st.columns(2)
    with col1:
        monto_input = st.number_input(
            "Capital a Invertir ($ USD)",
            min_value=1.0,
            value=None,
            placeholder="",
            step=1.0
        )
    with col2:
        plazo_input = st.number_input(
            "Plazo deseado (Días)",
            min_value=1,
            max_value=3600,
            value=None,
            placeholder="",
            step=1
        )
    submit_button = st.form_submit_button(label="🎯 Calcular Mejores Opciones", use_container_width=True)

# Asignación de valores predeterminados cuando el usuario no escribe en el campo
monto = monto_input if monto_input is not None else 5000.0
plazo_dias = plazo_input if plazo_input is not None else 180

# =========================================================================
# FILTRADO Y CÁLCULO DEL TOP 5
# =========================================================================
filtrado = df_datos[
    (df_datos['monto_minimo'] <= monto) &
    (df_datos['monto_maximo'] >= monto) &
    (df_datos['dias_minimo'] <= plazo_dias) &
    (df_datos['dias_maximo'] >= plazo_dias)
].copy()

if filtrado.empty:
    st.error(f"No hay opciones disponibles para un monto de ${monto:,.2f} a {plazo_dias} días.")
else:
    # Fórmula de interés simple: Ganancia = Monto * (TNA / 100) * (Días / 360)
    filtrado['ganancia_estimada'] = monto * (filtrado['tasa_nominal_porcentaje'] / 100.0) * (plazo_dias / 360.0)
    
    # Ordenar estrictamente por Tasa Nominal (TNA) descendente
    top5 = filtrado.sort_values(
        by='tasa_nominal_porcentaje',
        ascending=False
    ).head(5).reset_index(drop=True)

    st.markdown("---")
    st.subheader("🏆 Top 5 Mejores Opciones")

    # Opción ganadora destacada
    ganador = top5.iloc[0]
    st.success(f"🥇 **Opción Ganadora:** **{ganador['banco']}** con una TNA de **{ganador['tasa_nominal_porcentaje']:.2f}%** y una ganancia estimada de **${ganador['ganancia_estimada']:,.2f}**")

    # Construcción de la tabla (sin columna Rank, dejando 'Institución Financiera' como columna 1 fija)
    html_rows = ""
    for idx, row in top5.iterrows():
        banco = row['banco']
        rango = f"{row['dias_minimo']} - {row['dias_maximo']} días"
        tna = f"{row['tasa_nominal_porcentaje']:.2f}%"
        ganancia = f"${row['ganancia_estimada']:,.2f}"
        
        html_rows += f"<tr><td><strong>{banco}</strong></td><td>{rango}</td><td><span style='color: #22c55e; font-weight: bold;'>{tna}</span></td><td><span style='color: #3b82f6; font-weight: bold;'>{ganancia}</span></td></tr>"

    table_html = f"""<div class="table-container"><table class="custom-table"><thead><tr><th>Institución<br>Financiera</th><th>Rango Plazo<br>Aplicado</th><th>TNA<br>(%)</th><th>Ganancia<br>Estimada ($)</th></tr></thead><tbody>{html_rows}</tbody></table></div>"""

    st.markdown(table_html, unsafe_allow_html=True)