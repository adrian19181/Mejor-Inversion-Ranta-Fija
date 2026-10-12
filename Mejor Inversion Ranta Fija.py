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

# Estilos CSS
st.markdown("""
<style>
    [data-testid="stInputInstruction"], 
    small[data-testid="stInputInstruction"], 
    div[data-testid="InputInstruction"],
    div[data-baseweb="typo-caption"] { display: none !important; }
    input::-webkit-input-placeholder, input::-moz-placeholder, input:-ms-input-placeholder, input::placeholder { color: transparent !important; opacity: 0 !important; }
    div[data-baseweb="input"] input { text-align: center !important; font-size: 18px !important; font-weight: 600 !important; }
    div[data-testid="stNumberInputContainer"] button, button[data-testid="stNumberInputStepDown"], button[data-testid="stNumberInputStepUp"] { display: none !important; }
    input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none !important; margin: 0 !important; }
    input[type=number] { -moz-appearance: textfield !important; }
    .table-container { width: 100%; overflow-x: auto; margin-top: 15px; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); background-color: #0F172A; }
    .custom-table { width: 100%; border-collapse: separate; border-spacing: 0; table-layout: auto; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }
    .custom-table th { background-color: #1E293B; color: #FFFFFF; text-align: center !important; vertical-align: middle !important; padding: 10px 8px; font-size: 13px; font-weight: 700; line-height: 1.25; white-space: nowrap; border-bottom: 2px solid #334155; border-top: 1px solid #334155; }
    .custom-table td { text-align: center !important; vertical-align: middle !important; padding: 10px 8px; font-size: 14px; white-space: nowrap; border-bottom: 1px solid #334155; }
    .custom-table th:nth-child(1) { position: sticky; left: 0; z-index: 5; background-color: #1E293B; min-width: 110px; max-width: 130px; white-space: normal !important; word-wrap: break-word !important; border-right: 2px solid #64748B !important; }
    .custom-table td:nth-child(1) { position: sticky; left: 0; z-index: 3; background-color: #0F172A; min-width: 110px; max-width: 130px; white-space: normal !important; word-wrap: break-word !important; line-height: 1.2; border-right: 2px solid #64748B !important; }
    .custom-table tr:nth-child(even) td { background-color: #182234 !important; }
    .custom-table tr:hover td { background-color: #243248 !important; }
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
# CARGA DE TASAS REFERENCIALES (Con logs en Terminal)
# =========================================================================
@st.cache_data(ttl=3600)
def cargar_tasas_referenciales():
    print("\n\n" + "="*50)
    print("--- DEBUG: INICIANDO CARGA DE TASAS REFERENCIALES ---")
    content = None
    
    try:
        print("DEBUG: Intentando descargar desde Drive...")
        response = requests.get(DRIVE_URL, timeout=5)
        if response.status_code == 200:
            content = io.BytesIO(response.content)
            print("DEBUG: Descarga de Drive exitosa.")
        else:
            print(f"DEBUG: Falló Drive. Código de estado HTTP: {response.status_code}")
    except Exception as e:
        print(f"DEBUG: Error al conectar con Drive: {e}")

    if content is not None:
        try:
            xl_drive = pd.ExcelFile(content)
            print(f"DEBUG: Hojas detectadas en el Excel de Drive: {xl_drive.sheet_names}")
            if "Tasas Referenciales" in xl_drive.sheet_names:
                df = pd.read_excel(xl_drive, sheet_name="Tasas Referenciales")
                print("DEBUG: Hoja 'Tasas Referenciales' leída correctamente desde Drive.")
                print(f"DEBUG: Columnas crudas en Drive: {df.columns.tolist()}")
                print(f"DEBUG: Primeras 3 filas crudas:\n{df.head(3)}")
                return df
            else:
                print("DEBUG: La hoja 'Tasas Referenciales' NO EXISTE en el archivo de Drive.")
        except Exception as e:
            print(f"DEBUG: Error procesando Excel de Drive: {e}")

    print(f"DEBUG: Intentando cargar archivo local: {LOCAL_BACKUP_PATH}")
    try:
        xl_local = pd.ExcelFile(LOCAL_BACKUP_PATH)
        print(f"DEBUG: Hojas detectadas en el Excel Local: {xl_local.sheet_names}")
        if "Tasas Referenciales" in xl_local.sheet_names:
            df = pd.read_excel(xl_local, sheet_name="Tasas Referenciales")
            print("DEBUG: Hoja 'Tasas Referenciales' leída correctamente desde Local.")
            print(f"DEBUG: Columnas crudas en Local: {df.columns.tolist()}")
            print(f"DEBUG: Primeras 3 filas crudas:\n{df.head(3)}")
            return df
        else:
            print("DEBUG: La hoja 'Tasas Referenciales' NO EXISTE en el archivo Local.")
    except Exception as e:
        print(f"DEBUG: Error procesando Excel Local: {e}")

    print("DEBUG: No se pudo cargar la hoja de ninguna fuente.")
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
    st.subheader("🏆 Mejor Opción Plazo Fijo")
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
# KPIS DE TASAS REFERENCIALES (BCE)
# =========================================================================
df_ref = cargar_tasas_referenciales()

print("--- DEBUG: PROCESANDO DATOS PARA KPIS ---")
print(f"DEBUG: Tamaño del df_ref extraído: {df_ref.shape}")

try:
    if not df_ref.empty and len(df_ref.columns) >= 2:
        print("DEBUG: Iniciando parseo de fechas...")
        df_ref['fecha_dt'] = pd.to_datetime(df_ref.iloc[:, 0], dayfirst=True, errors='coerce')
        print(f"DEBUG: Total fechas Nulas post-parseo: {df_ref['fecha_dt'].isna().sum()}")
        
        print("DEBUG: Limpiando y parseando tasas a float...")
        tasas_limpias = df_ref.iloc[:, 1].astype(str).str.replace(',', '.').str.replace('%', '').str.strip()
        df_ref['tasa_val'] = pd.to_numeric(tasas_limpias, errors='coerce')
        print(f"DEBUG: Total tasas Nulas post-parseo: {df_ref['tasa_val'].isna().sum()}")
        
        df_validos = df_ref.dropna(subset=['fecha_dt', 'tasa_val']).copy()
        print(f"DEBUG: Total de filas válidas finales: {len(df_validos)}")

        if not df_validos.empty:
            df_validos = df_validos.sort_values('fecha_dt')
            fecha_actual = datetime.datetime.now()
            print(f"DEBUG: Fecha actual del sistema: {fecha_actual}")
            
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

            print(f"DEBUG: KPIs calculados -> Actual: {tasa_actual} | Promedio: {tasa_promedio_anio}")

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("---")
            st.subheader("📊 Tasas Pasivas Referenciales (BCE)")
            kpi_col1, kpi_col2 = st.columns(2)
            with kpi_col1:
                st.metric(label=f"Referencial Actual ({etiqueta_mes_anio})", value=f"{tasa_actual:.2f}%")
            with kpi_col2:
                st.metric(label=f"Promedio Anual ({fecha_kpi.year})", value=f"{tasa_promedio_anio:.2f}%")
        else:
            print("DEBUG: df_validos quedó vacío después de eliminar los nulos.")
            st.markdown("<br><hr>", unsafe_allow_html=True)
            st.error("📌 Error de lectura: No se pudo procesar la hoja de 'Tasas Referenciales'. Verifica que contenga fechas y tasas numéricas.")
    else:
        print("DEBUG: Condición IF no cumplida (el DataFrame está vacío o tiene menos de 2 columnas).")
        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.error("📌 Error de lectura: No se pudo procesar la hoja de 'Tasas Referenciales'. Verifica que contenga fechas y tasas numéricas.")
except Exception as e:
    print(f"DEBUG: Ocurrió un error inesperado al calcular los KPIs: {e}")
    import traceback
    traceback.print_exc()
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.error("📌 Ocurrió un error inesperado durante el procesamiento de las Tasas Referenciales.")