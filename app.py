import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import date, time
import streamlit.components.v1 as components


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Monitoreo de Calidad de Agua",
    page_icon="💧",
    layout="wide"
)

ARCHIVO_PARAMETROS = "PARAMETROS_AGUA.xlsx"

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
# CARGAR MAESTRO
# ============================================================

@st.cache_data
def cargar_parametros():

    df = pd.read_excel(ARCHIVO_PARAMETROS)

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    # Compatibilidad con versiones anteriores
    if "PARAMETRO" in df.columns and "PARÁMETROS" not in df.columns:
        df = df.rename(
            columns={
                "PARAMETRO": "PARÁMETROS"
            }
        )

    return df


try:

    df_parametros = cargar_parametros()

except Exception as e:

    st.error(
        f"No se pudo cargar el archivo `{ARCHIVO_PARAMETROS}`."
    )

    st.stop()


# ============================================================
# FUNCIONES
# ============================================================

def es_olor_sabor(parametro):

    parametro_normalizado = (
        str(parametro)
        .strip()
        .upper()
    )

    return parametro_normalizado in [
        "OLOR",
        "SABOR"
    ]


def validar_numero(valor):

    if valor is None:
        return False

    valor = str(valor).strip()

    if valor == "":
        return False

    # Permite:
    # 10
    # 10.5
    # 0.0013
    # 10,5

    patron = r"^\d+([.,]\d+)?$"

    return bool(
        re.fullmatch(
            patron,
            valor
        )
    )


def convertir_a_numero(valor):

    """
    Convierte el texto ingresado por el usuario
    en un verdadero número Python.

    Esto permite que Excel lo reconozca como número
    y no como texto.
    """

    valor = str(valor).strip().replace(",", ".")

    numero = float(valor)

    # Si es un entero, guardar como entero
    # para que Excel muestre 5 en lugar de 5.0
    if numero.is_integer():

        return int(numero)

    return numero


def generar_excel(df):

    """
    Genera el archivo Excel en memoria.
    """

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
        date_format="DD/MM/YYYY",
        datetime_format="DD/MM/YYYY HH:MM:SS"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="BD_AGUA"
        )

        worksheet = writer.sheets["BD_AGUA"]

        # ----------------------------------------------------
        # ANCHOS
        # ----------------------------------------------------

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

            worksheet.column_dimensions[
                columna
            ].width = ancho

        # ----------------------------------------------------
        # FORMATO FECHA
        # ----------------------------------------------------

        # Columna C = FECHA DE MUESTREO
        for celda in worksheet["C"][1:]:

            celda.number_format = "DD/MM/YYYY"

        # ----------------------------------------------------
        # FORMATO HORA
        # ----------------------------------------------------

        # Columna D = HORA DE MUESTREO
        for celda in worksheet["D"][1:]:

            celda.number_format = "HH:MM:SS"

    output.seek(0)

    return output


# ============================================================
# JAVASCRIPT PARA NAVEGACIÓN CON ENTER
# ============================================================

def activar_enter_siguiente():

    """
    Intenta hacer que ENTER pase al siguiente campo
    de resultado.

    Se ejecuta en el navegador.
    """

    components.html(
        """
        <script>

        const iniciarNavegacion = () => {

            // Buscar todos los inputs visibles
            const inputs = window.parent.document.querySelectorAll(
                'input'
            );

            inputs.forEach((input, index) => {

                // Evitar agregar el evento varias veces
                if (input.dataset.enterNavigation === "true") {
                    return;
                }

                input.dataset.enterNavigation = "true";

                input.addEventListener(
                    "keydown",
                    function(event) {

                        if (event.key !== "Enter") {
                            return;
                        }

                        event.preventDefault();

                        const inputsActualizados =
                            Array.from(
                                window.parent.document.querySelectorAll(
                                    'input'
                                )
                            ).filter(
                                el =>
                                !el.disabled &&
                                el.offsetParent !== null
                            );

                        const posicion =
                            inputsActualizados.indexOf(this);

                        if (
                            posicion >= 0 &&
                            posicion + 1 <
                            inputsActualizados.length
                        ) {

                            inputsActualizados[
                                posicion + 1
                            ].focus();

                        }

                    }
                );

            });

        };


        iniciarNavegacion();


        // Streamlit actualiza el DOM constantemente.
        // Volvemos a revisar periódicamente.

        setInterval(
            iniciarNavegacion,
            1000
        );

        </script>
        """,
        height=0
    )


# ============================================================
# TÍTULO
# ============================================================

st.title(
    "💧 Registro de Análisis de Calidad de Agua"
)

st.markdown(
    """
    Registra los resultados de laboratorio de acuerdo con
    el tipo de agua seleccionado.

    **Todos los parámetros deben ser completados antes
    de poder descargar el archivo Excel.**
    """
)


# ============================================================
# SELECCIÓN DE UTILIDAD
# ============================================================

utilidades_disponibles = sorted(
    df_parametros[
        "UTILIDAD"
    ]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

utilidad = st.selectbox(
    "UTILIDAD / TIPO DE AGUA",
    utilidades_disponibles,
    index=None,
    placeholder="Seleccione el tipo de agua..."
)


if utilidad is None:

    st.info(
        "Seleccione un tipo de agua para comenzar."
    )

    st.stop()


# ============================================================
# FILTRAR PARÁMETROS
# ============================================================

df_utilidad = df_parametros[
    df_parametros[
        "UTILIDAD"
    ]
    .astype(str)
    .str.strip()
    == utilidad
].copy()


if df_utilidad.empty:

    st.error(
        f"No se encontraron parámetros para {utilidad}."
    )

    st.stop()


# ============================================================
# INFORMACIÓN DE LA MUESTRA
# ============================================================

st.subheader(
    "📋 Información de la muestra"
)

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
# RESULTADOS
# ============================================================

st.subheader(
    f"🧪 Resultados de laboratorio — {utilidad}"
)

st.caption(
    "Todos los parámetros son obligatorios. "
    "Para resultados menores al límite cuantificable, "
    "marque < LC. Estos resultados se exportarán como 0."
)


resultados = {}

parametros_faltantes = []

errores_numericos = []


# ============================================================
# TIPOS DE ANÁLISIS
# ============================================================

tipos_analisis = (
    df_utilidad[
        "TIPO DE ANÁLISIS"
    ]
    .fillna("Sin clasificar")
    .astype(str)
    .drop_duplicates()
    .tolist()
)


contador_parametro = 0


for tipo_analisis in tipos_analisis:

    df_grupo = df_utilidad[
        df_utilidad[
            "TIPO DE ANÁLISIS"
        ]
        .fillna("Sin clasificar")
        .astype(str)
        == tipo_analisis
    ]


    st.markdown(
        f"### {tipo_analisis}"
    )


    for _, fila in df_grupo.iterrows():

        parametro = str(
            fila["PARÁMETROS"]
        ).strip()

        unidad = (
            ""
            if pd.isna(fila.get("UNIDAD"))
            else str(
                fila["UNIDAD"]
            ).strip()
        )

        norma = (
            ""
            if pd.isna(
                fila.get("NORMA APLICABLE")
            )
            else str(
                fila["NORMA APLICABLE"]
            ).strip()
        )


        # ====================================================
        # OLOR / SABOR
        # ====================================================

        if es_olor_sabor(parametro):

            col1, col2, col3 = st.columns(
                [4, 2, 2]
            )

            with col1:

                st.markdown(
                    f"**{parametro}**"
                )

            with col2:

                st.caption(
                    f"Unidad: {unidad}"
                )

            with col3:

                resultado = st.selectbox(
                    parametro,
                    [
                        "Seleccione...",
                        "Aceptable",
                        "No aceptable"
                    ],
                    key=f"resultado_{utilidad}_{parametro}",
                    label_visibility="collapsed"
                )


            if resultado == "Seleccione...":

                resultados[parametro] = ""

                parametros_faltantes.append(
                    parametro
                )

            else:

                resultados[parametro] = resultado


        # ====================================================
        # PARÁMETROS NUMÉRICOS
        # ====================================================

        else:

            col1, col2, col3, col4 = st.columns(
                [4, 1.5, 1.5, 2]
            )


            with col1:

                st.markdown(
                    f"**{parametro}**"
                )

                minimo = fila.get(
                    "VALOR MIN ESTABLECIDO"
                )

                maximo = fila.get(
                    "VALOR MAX ESTABLECIDO"
                )

                texto_limite = ""


                if (
                    pd.notna(minimo)
                    and str(minimo).strip() != ""
                ):

                    texto_limite += (
                        f"Min: {minimo}"
                    )


                if (
                    pd.notna(maximo)
                    and str(maximo).strip() != ""
                ):

                    if texto_limite:

                        texto_limite += " | "

                    texto_limite += (
                        f"Max: {maximo}"
                    )


                if texto_limite:

                    st.caption(
                        texto_limite
                    )


            with col2:

                st.caption(
                    f"Unidad: {unidad}"
                )


            with col3:

                st.caption(
                    f"Norma: {norma}"
                )


            # ------------------------------------------------
            # CHECKBOX < LC
            # ------------------------------------------------

            with col4:

                usar_lc = st.checkbox(
                    "< LC",
                    key=f"lc_{utilidad}_{parametro}"
                )


            # ------------------------------------------------
            # CAMPO NUMÉRICO
            # ------------------------------------------------

            resultado_ingresado = st.text_input(
                f"Resultado - {parametro}",
                key=f"valor_{utilidad}_{parametro}",
                placeholder="Ingrese valor numérico",
                disabled=usar_lc,
                label_visibility="collapsed"
            )


            # ------------------------------------------------
            # < LC
            # ------------------------------------------------

            if usar_lc:

                # IMPORTANTE:
                # En Excel se guardará como 0 NUMÉRICO
                resultados[parametro] = 0


            # ------------------------------------------------
            # RESULTADO NORMAL
            # ------------------------------------------------

            else:

                resultado_ingresado = (
                    resultado_ingresado.strip()
                )


                if resultado_ingresado == "":

                    resultados[parametro] = ""

                    parametros_faltantes.append(
                        parametro
                    )


                else:

                    if validar_numero(
                        resultado_ingresado
                    ):

                        try:

                            resultados[parametro] = (
                                convertir_a_numero(
                                    resultado_ingresado
                                )
                            )

                        except Exception:

                            resultados[parametro] = (
                                resultado_ingresado
                            )

                            errores_numericos.append(
                                parametro
                            )


                    else:

                        resultados[parametro] = (
                            resultado_ingresado
                        )

                        errores_numericos.append(
                            parametro
                        )

                        st.error(
                            f"❌ **{parametro}** "
                            "debe contener únicamente "
                            "un valor numérico."
                        )


        contador_parametro += 1


    st.divider()


# ============================================================
# ACTIVAR NAVEGACIÓN ENTER
# ============================================================

activar_enter_siguiente()


# ============================================================
# VALIDACIÓN
# ============================================================

st.subheader(
    "🔎 Validación del registro"
)


errores_generales = []


if not ubicacion.strip():

    errores_generales.append(
        "UBICACIÓN"
    )


if not estacion.strip():

    errores_generales.append(
        "ESTACIÓN DE MUESTREO"
    )


# ============================================================
# PARÁMETROS FALTANTES
# ============================================================

if parametros_faltantes:

    st.warning(
        f"⚠️ Faltan completar "
        f"**{len(parametros_faltantes)} parámetro(s)**."
    )


    with st.expander(
        "Ver parámetros pendientes"
    ):

        for parametro in parametros_faltantes:

            st.write(
                f"- {parametro}"
            )


# ============================================================
# ERRORES NUMÉRICOS
# ============================================================

if errores_numericos:

    st.error(
        f"❌ Hay "
        f"**{len(errores_numericos)} parámetro(s)** "
        "con valores no numéricos."
    )


    with st.expander(
        "Ver parámetros con error"
    ):

        for parametro in errores_numericos:

            st.write(
                f"- {parametro}"
            )


# ============================================================
# ERRORES GENERALES
# ============================================================

if errores_generales:

    st.warning(
        "⚠️ Complete los siguientes datos obligatorios: "
        + ", ".join(
            errores_generales
        )
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
        f"Se han completado los "
        f"**{len(df_utilidad)} parámetros** "
        f"correspondientes a {utilidad}."
    )

else:

    st.info(
        "🔒 La descarga permanecerá bloqueada "
        "hasta completar todos los parámetros."
    )


# ============================================================
# EXPORTACIÓN
# ============================================================

st.divider()


if st.button(
    "📥 GENERAR Y DESCARGAR EXCEL",
    type="primary",
    use_container_width=True,
    disabled=not registro_completo
):

    registros = []


    for _, fila in df_utilidad.iterrows():

        parametro = str(
            fila["PARÁMETROS"]
        ).strip()

        resultado = resultados.get(
            parametro,
            ""
        )


        tipo_analisis = (
            ""
            if pd.isna(
                fila.get(
                    "TIPO DE ANÁLISIS"
                )
            )
            else str(
                fila[
                    "TIPO DE ANÁLISIS"
                ]
            ).strip()
        )


        unidad = (
            ""
            if pd.isna(
                fila.get("UNIDAD")
            )
            else str(
                fila["UNIDAD"]
            ).strip()
        )


        norma = (
            ""
            if pd.isna(
                fila.get(
                    "NORMA APLICABLE"
                )
            )
            else str(
                fila[
                    "NORMA APLICABLE"
                ]
            ).strip()
        )


        registros.append({

            "UBICACIÓN":
                ubicacion.strip(),

            "ESTACIÓN DE MUESTREO":
                estacion.strip(),

            # IMPORTANTE:
            # Aquí NO usamos strftime.
            # Se guarda directamente como objeto date.
            "FECHA DE MUESTREO":
                fecha_muestreo,

            # Hora como objeto time
            "HORA DE MUESTREO":
                hora_muestreo,

            "UTILIDAD":
                utilidad,

            "TIPO DE ANÁLISIS":
                tipo_analisis,

            "PARÁMETROS":
                parametro,

            "RESULTADO":
                resultado,

            "UNIDAD":
                unidad,

            "NORMA APLICABLE":
                norma
        })


    # ========================================================
    # DATAFRAME
    # ========================================================

    df_exportacion = pd.DataFrame(
        registros,
        columns=COLUMNAS_SALIDA
    )


    # ========================================================
    # SEGURIDAD
    # ========================================================

    if (
        df_exportacion[
            "RESULTADO"
        ]
        .astype(str)
        .str.strip()
        .eq("")
        .any()
    ):

        st.error(
            "❌ El archivo no puede ser generado "
            "porque existen parámetros sin resultado."
        )

        st.stop()


    # ========================================================
    # GENERAR EXCEL
    # ========================================================

    archivo_excel = generar_excel(
        df_exportacion
    )


    nombre_archivo = (
        f"Monitoreo_Agua_"
        f"{utilidad}_"
        f"{fecha_muestreo.strftime('%Y%m%d')}.xlsx"
    )


    st.success(
        "✅ Todos los parámetros fueron "
        "completados correctamente."
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
```

