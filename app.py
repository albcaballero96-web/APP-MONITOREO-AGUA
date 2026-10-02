import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import date, time

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Monitoreo de Calidad de Agua",
    page_icon="💧",
    layout="wide"
)

ARCHIVO_PARAMETROS = "PARAMETROS_AGUA.xlsx"

# Columnas definitivas de salida
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
    "NORMA APLICABLE"
]


# ============================================================
# CARGAR MAESTRO DE PARÁMETROS
# ============================================================

@st.cache_data
def cargar_parametros():
    df = pd.read_excel(ARCHIVO_PARAMETROS)

    # Normalizar nombres de columnas
    df.columns = [str(col).strip() for col in df.columns]

    # Compatibilidad con las columnas generadas anteriormente
    if "PARAMETRO" in df.columns and "PARÁMETROS" not in df.columns:
        df = df.rename(columns={"PARAMETRO": "PARÁMETROS"})

    return df


try:
    df_parametros = cargar_parametros()

except Exception as e:
    st.error(
        f"No se pudo cargar el archivo `{ARCHIVO_PARAMETROS}`. "
        f"Verifica que esté en la misma carpeta que `app.py`."
    )
    st.stop()


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def es_olor_sabor(parametro):
    """
    Determina si el parámetro debe utilizar el selector
    Aceptable / No aceptable.
    """
    parametro_normalizado = str(parametro).strip().upper()

    return parametro_normalizado in ["OLOR", "SABOR"]


def es_numerico_parametro(parametro):
    """
    Determina si el parámetro debe ser ingresado como valor numérico.

    Olor y Sabor son tratados de forma independiente.
    Todos los demás parámetros se consideran numéricos,
    incluyendo parámetros como Temperatura y pH.
    """
    return not es_olor_sabor(parametro)


def validar_numero(valor):
    """
    Valida que el resultado ingresado sea numérico.

    Permite:
    0
    1
    6.61
    0.0013
    1500
    1,25

    No permite:
    abc
    6.6 mg
    <1
    Aceptable
    etc.
    """

    if valor is None:
        return False

    valor = str(valor).strip()

    if valor == "":
        return False

    # Permitir punto o coma decimal
    patron = r"^\d+([.,]\d+)?$"

    return bool(re.fullmatch(patron, valor))


def convertir_numero(valor):
    """
    Convierte coma decimal a punto para mantener
    consistencia en el archivo Excel.
    """
    return str(valor).strip().replace(",", ".")


def generar_excel(df):
    """
    Genera el Excel directamente en memoria.
    """

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="BD_AGUA"
        )

        # Ajustar anchos de columnas
        worksheet = writer.sheets["BD_AGUA"]

        anchos = {
            "A": 25,
            "B": 35,
            "C": 18,
            "D": 18,
            "E": 15,
            "F": 25,
            "G": 40,
            "H": 20,
            "I": 20,
            "J": 30
        }

        for columna, ancho in anchos.items():
            worksheet.column_dimensions[columna].width = ancho

    output.seek(0)

    return output


# ============================================================
# TÍTULO
# ============================================================

st.title("💧 Registro de Análisis de Calidad de Agua")

st.markdown(
    """
    Registra los resultados de laboratorio de acuerdo con el tipo de agua.
    
    **Todos los parámetros del tipo de agua seleccionado deben ser completados
    antes de poder descargar el archivo Excel.**
    """
)


# ============================================================
# SELECCIÓN DE UTILIDAD
# ============================================================

utilidades_disponibles = sorted(
    df_parametros["UTILIDAD"].dropna().astype(str).unique().tolist()
)

utilidad = st.selectbox(
    "UTILIDAD / TIPO DE AGUA",
    utilidades_disponibles,
    index=None,
    placeholder="Seleccione el tipo de agua..."
)


# ============================================================
# SI NO SE HA SELECCIONADO UTILIDAD
# ============================================================

if utilidad is None:
    st.info("Seleccione un tipo de agua para comenzar.")
    st.stop()


# ============================================================
# FILTRAR PARÁMETROS
# ============================================================

df_utilidad = df_parametros[
    df_parametros["UTILIDAD"].astype(str).str.strip() == utilidad
].copy()


if df_utilidad.empty:
    st.error(
        f"No se encontraron parámetros configurados para la utilidad: {utilidad}"
    )
    st.stop()


# ============================================================
# INFORMACIÓN DE LA MUESTRA
# ============================================================

st.subheader("📋 Información de la muestra")

col1, col2 = st.columns(2)

with col1:
    ubicacion = st.text_input(
        "UBICACIÓN *",
        placeholder="Ej. LA RINCONADA"
    )

with col2:
    estacion = st.text_input(
        "ESTACIÓN DE MUESTREO *",
        placeholder="Ej. RESERVORIO - LA RINCONADA"
    )

col3, col4 = st.columns(2)

with col3:
    fecha_muestreo = st.date_input(
        "FECHA DE MUESTREO *",
        value=date.today(),
        format="DD/MM/YYYY"
    )

with col4:
    hora_muestreo = st.time_input(
        "HORA DE MUESTREO *",
        value=time(8, 0)
    )


st.divider()


# ============================================================
# PARÁMETROS
# ============================================================

st.subheader(
    f"🧪 Resultados de laboratorio — {utilidad}"
)

st.caption(
    "Todos los parámetros son obligatorios. "
    "En parámetros numéricos puede marcar '< LC' cuando el resultado "
    "del laboratorio corresponda a un valor menor al límite cuantificable."
)


# Diccionario para guardar resultados
resultados = {}

errores_numericos = []

parametros_faltantes = []


# ============================================================
# AGRUPAR POR TIPO DE ANÁLISIS
# ============================================================

tipos_analisis = (
    df_utilidad["TIPO DE ANÁLISIS"]
    .fillna("Sin clasificar")
    .astype(str)
    .drop_duplicates()
    .tolist()
)


for tipo_analisis in tipos_analisis:

    df_grupo = df_utilidad[
        df_utilidad["TIPO DE ANÁLISIS"].fillna("Sin clasificar").astype(str)
        == tipo_analisis
    ]

    st.markdown(f"### {tipo_analisis}")

    for _, fila in df_grupo.iterrows():

        parametro = str(fila["PARÁMETROS"]).strip()

        unidad = (
            ""
            if pd.isna(fila.get("UNIDAD"))
            else str(fila["UNIDAD"]).strip()
        )

        norma = (
            ""
            if pd.isna(fila.get("NORMA APLICABLE"))
            else str(fila["NORMA APLICABLE"]).strip()
        )

        # ----------------------------------------------------
        # Olor y Sabor
        # ----------------------------------------------------

        if es_olor_sabor(parametro):

            col1, col2, col3 = st.columns([4, 2, 2])

            with col1:
                st.markdown(f"**{parametro}**")

            with col2:
                st.caption(f"Unidad: {unidad}")

            with col3:

                resultado = st.selectbox(
                    parametro,
                    options=[
                        "Seleccione...",
                        "Aceptable",
                        "No aceptable"
                    ],
                    key=f"resultado_{utilidad}_{parametro}",
                    label_visibility="collapsed"
                )

            if resultado == "Seleccione...":
                resultados[parametro] = ""
                parametros_faltantes.append(parametro)
            else:
                resultados[parametro] = resultado

        # ----------------------------------------------------
        # PARÁMETROS NUMÉRICOS
        # ----------------------------------------------------

        else:

            col1, col2, col3, col4 = st.columns([4, 1.5, 1.5, 2])

            with col1:
                st.markdown(f"**{parametro}**")

                # Mostrar límites establecidos
                minimo = fila.get("VALOR MIN ESTABLECIDO")
                maximo = fila.get("VALOR MAX ESTABLECIDO")

                texto_limite = ""

                if pd.notna(minimo) and str(minimo).strip() != "":
                    texto_limite += f"Min: {minimo}"

                if pd.notna(maximo) and str(maximo).strip() != "":
                    if texto_limite:
                        texto_limite += " | "
                    texto_limite += f"Max: {maximo}"

                if texto_limite:
                    st.caption(texto_limite)

            with col2:
                st.caption(f"Unidad: {unidad}")

            with col3:
                st.caption(f"Norma: {norma}")

            # ------------------------------------------------
            # CHECKBOX < LC
            # ------------------------------------------------

            with col4:

                usar_lc = st.checkbox(
                    "< LC",
                    key=f"lc_{utilidad}_{parametro}"
                )

            # ------------------------------------------------
            # CAMPO DE RESULTADO
            # ------------------------------------------------

            resultado_ingresado = st.text_input(
                f"Resultado - {parametro}",
                key=f"valor_{utilidad}_{parametro}",
                placeholder="Ingrese valor numérico",
                disabled=usar_lc,
                label_visibility="collapsed"
            )

            # ------------------------------------------------
            # SI ESTÁ MARCADO < LC
            # ------------------------------------------------

            if usar_lc:

                resultados[parametro] = "< LC"

            # ------------------------------------------------
            # SI NO ESTÁ MARCADO < LC
            # ------------------------------------------------

            else:

                resultado_ingresado = resultado_ingresado.strip()

                if resultado_ingresado == "":

                    resultados[parametro] = ""
                    parametros_faltantes.append(parametro)

                else:

                    if validar_numero(resultado_ingresado):

                        resultados[parametro] = convertir_numero(
                            resultado_ingresado
                        )

                    else:

                        resultados[parametro] = resultado_ingresado

                        errores_numericos.append(
                            parametro
                        )

                        st.error(
                            f"❌ **{parametro}** debe contener únicamente "
                            f"un valor numérico."
                        )


    st.divider()


# ============================================================
# RESUMEN DE VALIDACIÓN
# ============================================================

st.subheader("🔎 Validación del registro")


# Validación de datos generales

errores_generales = []

if not ubicacion.strip():
    errores_generales.append("UBICACIÓN")

if not estacion.strip():
    errores_generales.append("ESTACIÓN DE MUESTREO")


# ------------------------------------------------------------
# Mostrar parámetros faltantes
# ------------------------------------------------------------

if parametros_faltantes:

    st.warning(
        f"⚠️ Faltan completar **{len(parametros_faltantes)} parámetro(s)**."
    )

    with st.expander("Ver parámetros pendientes"):

        for parametro in parametros_faltantes:
            st.write(f"- {parametro}")


# ------------------------------------------------------------
# Mostrar errores numéricos
# ------------------------------------------------------------

if errores_numericos:

    st.error(
        f"❌ Hay **{len(errores_numericos)} parámetro(s)** "
        f"con valores no numéricos."
    )

    with st.expander("Ver parámetros con error"):

        for parametro in errores_numericos:
            st.write(f"- {parametro}")


# ------------------------------------------------------------
# Mostrar errores generales
# ------------------------------------------------------------

if errores_generales:

    st.warning(
        "⚠️ Complete los siguientes datos obligatorios: "
        + ", ".join(errores_generales)
    )


# ============================================================
# VALIDACIÓN FINAL
# ============================================================

registro_completo = (
    len(parametros_faltantes) == 0
    and len(errores_numericos) == 0
    and len(errores_generales) == 0
)


if registro_completo:

    st.success(
        f"✅ Registro completo. "
        f"Se han completado los **{len(df_utilidad)} parámetros** "
        f"correspondientes a {utilidad}."
    )

else:

    st.info(
        "🔒 La descarga permanecerá bloqueada hasta completar "
        "todos los parámetros y corregir los valores ingresados."
    )


# ============================================================
# GENERAR EXCEL
# ============================================================

st.divider()

if st.button(
    "📥 GENERAR Y DESCARGAR EXCEL",
    type="primary",
    use_container_width=True,
    disabled=not registro_completo
):

    # --------------------------------------------------------
    # CREAR REGISTROS
    # --------------------------------------------------------

    registros = []

    for _, fila in df_utilidad.iterrows():

        parametro = str(fila["PARÁMETROS"]).strip()

        resultado = resultados.get(parametro, "")

        tipo_analisis = (
            ""
            if pd.isna(fila.get("TIPO DE ANÁLISIS"))
            else str(fila["TIPO DE ANÁLISIS"]).strip()
        )

        unidad = (
            ""
            if pd.isna(fila.get("UNIDAD"))
            else str(fila["UNIDAD"]).strip()
        )

        norma = (
            ""
            if pd.isna(fila.get("NORMA APLICABLE"))
            else str(fila["NORMA APLICABLE"]).strip()
        )

        registros.append({
            "UBICACIÓN": ubicacion.strip(),
            "ESTACIÓN DE MUESTREO": estacion.strip(),
            "FECHA DE MUESTREO": fecha_muestreo.strftime("%d/%m/%Y"),
            "HORA DE MUESTREO": hora_muestreo.strftime("%H:%M:%S"),
            "UTILIDAD": utilidad,
            "TIPO DE ANÁLISIS": tipo_analisis,
            "PARÁMETROS": parametro,
            "RESULTADO": resultado,
            "UNIDAD": unidad,
            "NORMA APLICABLE": norma
        })


    # --------------------------------------------------------
    # CREAR DATAFRAME
    # --------------------------------------------------------

    df_exportacion = pd.DataFrame(
        registros,
        columns=COLUMNAS_SALIDA
    )


    # --------------------------------------------------------
    # SEGURIDAD EXTRA:
    # NO PERMITIR EXPORTAR FILAS VACÍAS
    # --------------------------------------------------------

    if df_exportacion["RESULTADO"].astype(str).str.strip().eq("").any():

        st.error(
            "❌ El archivo no puede ser generado porque existen "
            "parámetros sin resultado."
        )

        st.stop()


    # --------------------------------------------------------
    # GENERAR ARCHIVO
    # --------------------------------------------------------

    archivo_excel = generar_excel(df_exportacion)


    nombre_archivo = (
        f"Monitoreo_Agua_"
        f"{utilidad}_"
        f"{fecha_muestreo.strftime('%Y%m%d')}.xlsx"
    )


    st.success(
        "✅ Todos los parámetros fueron completados correctamente. "
        "El archivo está listo para descargar."
    )


    st.download_button(
        label="⬇️ DESCARGAR EXCEL",
        data=archivo_excel,
        file_name=nombre_archivo,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )
