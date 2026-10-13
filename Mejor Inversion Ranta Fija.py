import pandas as pd
import streamlit as st

# Datos de ejemplo con la estructura habitual de tasas
data = [
    {"institucion": "29 de Octubre", "producto": "Ahorro Programado", "plazo_min": 30, "plazo_max": 690, "tna": 6.00},
    {"institucion": "Jardín Azuayo", "producto": "Ahorro Programado", "plazo_min": 360, "plazo_max": 720, "tna": 6.50}
]
df = pd.DataFrame(data)

# Captura del plazo
plazo_dias = st.number_input("Plazo a cumplir (Días)", min_value=1, value=361)

# Normalización de texto en la columna del producto
df['producto_clean'] = df['producto'].str.strip().str.lower()

# 1. Ajuste de tolerancia para el año comercial (360 vs 365 días)
# Si el usuario ingresa 361 días, se permite un margen de ajuste para rangos cerrados en 360
tolerancia_dias = 5

df_filtrado = df[
    (df['plazo_min'] <= (plazo_dias + tolerancia_dias)) & 
    (df['plazo_max'] >= (plazo_dias - tolerancia_dias))
].copy()

# 2. Selección de la mejor tasa (TNA)
if not df_filtrado.empty:
    df_resultado = df_filtrado.sort_values(by="tna", ascending=False)
    opcion_ganadora = df_resultado.iloc[0]
    
    st.write(f"Opción Ganadora: {opcion_ganadora['institucion']} con TNA de {opcion_ganadora['tna']:.2f}%")
    st.dataframe(df_resultado[['institucion', 'producto', 'plazo_min', 'plazo_max', 'tna']])
else:
    st.warning("No se encontraron instituciones para el plazo ingresado.")