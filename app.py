import os
from datetime import date, time
import pandas as pd
import streamlit as st

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_FILE = os.path.join(BASE_DIR, "PARAMETROS_AGUA.xlsx")
EXPORT_DIR = os.path.join(BASE_DIR, "exportaciones")
os.makedirs(EXPORT_DIR, exist_ok=True)

COLUMNAS_SALIDA = [
    "UBICACIÓN",
    "ESTACIÓN DE MUESTREO",
    "FECHA DE MUESTREO",
    "HORA DE MUESTREO",
    "UTILIDAD",
    "TIPO DE ANÁLISIS",
    "PARÁMETROS",
    "RESULTADO",
    "UNIDAD",
    "NORMA APLICABLE",
    "OBSERVACIÓN",
]

st.set_page_config(
    page_title="Monitoreo de Calidad de Agua",
    page_icon="💧",
    layout="wide",
)

@st.cache_data
def cargar_parametros():
    return pd.read_excel(MASTER_FILE, sheet_name="PARAMETROS", dtype=object)

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def mostrar_limite(minimo, maximo):
    minimo = limpiar_texto(minimo)
    maximo = limpiar_texto(maximo)
    if minimo and maximo:
        return f"Rango: {minimo} – {maximo}"
    if minimo:
        return f"Mínimo: {minimo}"
    if maximo:
        return f"Máximo: {maximo}"
    return "Sin límite establecido"

def construir_registros(datos, parametros, resultados, observaciones):
    filas = []
    for i, (_, p) in enumerate(parametros.iterrows()):
        filas.append({
            "UBICACIÓN": datos["ubicacion"],
            "ESTACIÓN DE MUESTREO": datos["estacion"],
            "FECHA DE MUESTREO": datos["fecha"].strftime("%d/%m/%Y"),
            "HORA DE MUESTREO": datos["hora"].strftime("%H:%M:%S"),
            "UTILIDAD": datos["utilidad"],
            "TIPO DE ANÁLISIS": limpiar_texto(p["TIPO DE ANÁLISIS"]),
            "PARÁMETROS": limpiar_texto(p["PARAMETRO"]),
            "RESULTADO": resultados[i],
            "UNIDAD": limpiar_texto(p["UNIDAD"]),
            "NORMA APLICABLE": limpiar_texto(p["NORMA APLICABLE"]),
            "OBSERVACIÓN": observaciones[i],
        })
    return pd.DataFrame(filas, columns=COLUMNAS_SALIDA)

st.title("💧 Monitoreo de Calidad de Agua")
st.caption("Versión 1 — captura de resultados y exportación para Power BI")

parametros = cargar_parametros()

st.subheader("1. Datos del monitoreo")

c1, c2, c3 = st.columns(3)
with c1:
    utilidad = st.selectbox(
        "UTILIDAD",
        ["CONSUMO", "RIEGO", "FSMA"],
        index=0,
    )
with c2:
    ubicacion = st.text_input("UBICACIÓN", placeholder="Ej. LA RINCONADA")
with c3:
    estacion = st.text_input(
        "ESTACIÓN DE MUESTREO",
        placeholder="Ej. OSMOSIS - LA RINCONADA",
    )

c4, c5 = st.columns(2)
with c4:
    fecha = st.date_input("FECHA DE MUESTREO", value=date.today(), format="DD/MM/YYYY")
with c5:
    hora = st.time_input("HORA DE MUESTREO", value=time(8, 0))

parametros_filtrados = parametros[
    parametros["UTILIDAD"].astype(str).str.strip().str.upper() == utilidad
].copy()

if parametros_filtrados.empty:
    st.error(f"No existen parámetros configurados para {utilidad}.")
    st.stop()

st.subheader(f"2. Resultados — {utilidad}")

resultados = []
observaciones = []

# Agrupación visual por tipo de análisis
grupos = list(parametros_filtrados.groupby("TIPO DE ANÁLISIS", sort=False))

for tipo, grupo in grupos:
    st.markdown(f"### {tipo}")
    for idx, (_, p) in enumerate(grupo.iterrows()):
        original_index = p.name
        # índice único dentro del dataframe filtrado
        posicion = parametros_filtrados.index.get_loc(original_index)

        parametro = limpiar_texto(p["PARAMETRO"])
        unidad = limpiar_texto(p["UNIDAD"])
        limite = mostrar_limite(p["VALOR MIN ESTABLECIDO"], p["VALOR MAX ESTABLECIDO"])

        a, b, c = st.columns([4, 2, 3])
        with a:
            st.write(f"**{parametro}**")
            st.caption(limite)
        with b:
            resultado = st.text_input(
                "Resultado",
                key=f"resultado_{utilidad}_{original_index}",
                label_visibility="collapsed",
                placeholder="Ingrese resultado",
            )
        with c:
            st.caption(f"Unidad: {unidad}")

        obs = st.text_input(
            "Observación",
            key=f"obs_{utilidad}_{original_index}",
            label_visibility="collapsed",
            placeholder="Observación (opcional)",
        )

        # Guardamos por índice del dataframe filtrado
        while len(resultados) <= posicion:
            resultados.append("")
            observaciones.append("")
        resultados[posicion] = resultado.strip()
        observaciones[posicion] = obs.strip()

st.divider()

if st.button("💾 GUARDAR Y GENERAR EXCEL", type="primary", use_container_width=True):
    if not ubicacion.strip():
        st.error("Ingrese la UBICACIÓN.")
        st.stop()
    if not estacion.strip():
        st.error("Ingrese la ESTACIÓN DE MUESTREO.")
        st.stop()

    # Validamos que exista al menos un resultado.
    if not any(r.strip() for r in resultados):
        st.error("Ingrese al menos un resultado antes de guardar.")
        st.stop()

    datos = {
        "ubicacion": ubicacion.strip(),
        "estacion": estacion.strip(),
        "fecha": fecha,
        "hora": hora,
        "utilidad": utilidad,
    }

    # Solo exportamos los parámetros que tienen resultado.
    # Esto reproduce la estructura de tu base real: una fila por resultado
    # efectivamente informado por el laboratorio.
    parametros_exportar = parametros_filtrados.reset_index(drop=True).copy()
    resultados_exportar = []
    observaciones_exportar = []
    filas_parametros = []

    for i, resultado in enumerate(resultados):
        if resultado.strip():
            filas_parametros.append(i)
            resultados_exportar.append(resultado)
            observaciones_exportar.append(observaciones[i])

    parametros_exportar = parametros_exportar.iloc[filas_parametros].reset_index(drop=True)

    salida = construir_registros(
        datos,
        parametros_exportar,
        resultados_exportar,
        observaciones_exportar,
    )

    nombre = f"MONITOREO_{utilidad}_{fecha.strftime('%Y%m%d')}_{hora.strftime('%H%M%S')}.xlsx"
    ruta = os.path.join(EXPORT_DIR, nombre)

    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        salida.to_excel(writer, index=False, sheet_name="BASE_DATOS")

    st.success(f"Muestra generada: {nombre}")
    st.dataframe(salida, use_container_width=True, hide_index=True)

    with open(ruta, "rb") as f:
        st.download_button(
            "⬇️ DESCARGAR EXCEL",
            data=f,
            file_name=nombre,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
