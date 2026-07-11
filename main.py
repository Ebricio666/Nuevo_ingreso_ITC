import io
import re
import unicodedata
from difflib import SequenceMatcher
from datetime import date, datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONSTANTES CHASIDE
# ============================================================

CHASIDE_AREAS = ["C", "H", "A", "S", "I", "D", "E"]

CHASIDE_AREAS_LONG = {
    "C": "Administrativo",
    "H": "Humanidades y Sociales",
    "A": "Artístico",
    "S": "Ciencias de la Salud",
    "I": "Enseñanzas Técnicas",
    "D": "Defensa y Seguridad",
    "E": "Ciencias Experimentales"
}

CHASIDE_INTERESES_ITEMS = {
    "C": [1, 12, 20, 53, 64, 71, 78, 85, 91, 98],
    "H": [9, 25, 34, 41, 56, 67, 74, 80, 89, 95],
    "A": [3, 11, 21, 28, 36, 45, 50, 57, 81, 96],
    "S": [8, 16, 23, 33, 44, 52, 62, 70, 87, 92],
    "I": [6, 19, 27, 38, 47, 54, 60, 75, 83, 97],
    "D": [5, 14, 24, 31, 37, 48, 58, 65, 73, 84],
    "E": [17, 32, 35, 42, 49, 61, 68, 77, 88, 93]
}

CHASIDE_APTITUDES_ITEMS = {
    "C": [2, 15, 46, 51],
    "H": [30, 63, 72, 86],
    "A": [22, 39, 76, 82],
    "S": [4, 29, 40, 69],
    "I": [10, 26, 59, 90],
    "D": [13, 18, 43, 66],
    "E": [7, 55, 79, 94]
}

CHASIDE_PERFILES_CARRERA = {
    "Arquitectura": ["A", "I", "C"],
    "Contador Público": ["C", "D"],
    "Licenciatura en Administración": ["C", "D"],
    "Ingeniería Ambiental": ["I", "C", "E"],
    "Ingeniería Bioquímica": ["I", "C", "E"],
    "Ingeniería en Gestión Empresarial": ["C", "D", "H"],
    "Ingeniería Industrial": ["C", "D", "H"],
    "Ingeniería en Inteligencia Artificial": ["I", "E"],
    "Ingeniería Mecatrónica": ["I", "E"],
    "Ingeniería en Sistemas Computacionales": ["I", "E"]
}

CHASIDE_DESC_INTENSIDAD = {
    "Sin perfil": "La elección de carrera muestra baja correspondencia con el perfil vocacional identificado.",
    "Perfil en riesgo": "El perfil vocacional presenta una coincidencia mínima con la carrera elegida.",
    "Perfil en transición": "Existe congruencia funcional entre elección profesional y perfil vocacional, aunque aún en proceso de consolidación.",
    "Joven promesa": "Existe alta congruencia entre el perfil vocacional y la carrera elegida."
}

CHASIDE_COLUMNA_NOMBRE = "Ingrese su nombre completo"
CHASIDE_COLUMNA_CARRERA = "¿A qué carrera desea ingresar?"
CHASIDE_COLUMNA_EMAIL_1 = "Dirección de correo electrónico"
CHASIDE_COLUMNA_EMAIL_2 = "Escriba su correo electrónico"

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Perfil Integral de Aspirantes",
    page_icon="🎓",
    layout="wide"
)


# ============================================================
# NAVEGACIÓN LATERAL
# ============================================================

st.sidebar.title("🎓 Panel de navegación")

modulo_activo = st.sidebar.radio(
    "Selecciona un módulo",
    [
        "📂 Carga de archivos",
        "📘 EVALUATEC 2026",
        "🎓 Historial de Aspirantes",
        "🧭 Diagnóstico vocacional CHASIDE",
        "👤 Perfil individual"
    ]
)    
# ============================================================
# UTILIDADES GENERALES COMPARTIDAS
# ============================================================

def util_normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFD", texto)

    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )

    return " ".join(texto.split())


def util_limpiar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFD", texto)

    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )

    return re.sub(r"\s+", " ", texto)


def util_limpiar_texto_visible(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).replace("\n", " ")
    return re.sub(r"\s+", " ", texto).strip()
def util_encontrar_columna(df, posibles_nombres):
    """Busca una columna ignorando mayúsculas, acentos y espacios."""

    columnas_normalizadas = {
        util_limpiar_texto(columna): columna
        for columna in df.columns
    }

    for posible in posibles_nombres:
        posible_limpio = util_limpiar_texto(posible)

        if posible_limpio in columnas_normalizadas:
            return columnas_normalizadas[posible_limpio]

        for columna_limpia, columna_original in columnas_normalizadas.items():
            if posible_limpio in columna_limpia:
                return columna_original

    return None
def dataframe_a_excel_bytes(dic_hojas):
    """Convierte uno o varios DataFrames a archivo Excel en memoria."""

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nombre_hoja, df_hoja in dic_hojas.items():
            nombre_limpio = str(nombre_hoja)[:31] if nombre_hoja else "Hoja"

            df_hoja.to_excel(
                writer,
                index=False,
                sheet_name=nombre_limpio
            )

    output.seek(0)
    return output.getvalue()


def cargar_datos_globales():
    """
    Carga centralizada de Historial, EVALUATEC y CHASIDE.
    Guarda todo en st.session_state para no cargar archivos varias veces.
    """

    st.title("📂 Carga de archivos")
    st.caption(
        "Carga aquí las bases principales. Después podrás navegar entre módulos "
        "sin volver a subir los archivos."
    )

    # ------------------------------------------------------------
    # Historial
    # ------------------------------------------------------------

    st.markdown("## 1. Historial de Aspirantes")

    archivo_historial = st.file_uploader(
        "Carga el Excel de Historial de Aspirantes",
        type=["xlsx", "xls"],
        key="global_historial_archivo"
    )

    if archivo_historial is not None:
        try:
            with st.spinner("Procesando Historial de Aspirantes..."):
                df_historial, df_bitacora = hist_procesar_archivo_excel(
                    archivo_historial.getvalue()
                )

            if df_historial.empty:
                st.error("No se pudieron identificar registros de aspirantes.")
                st.dataframe(df_bitacora, use_container_width=True)

            else:
                columna_sexo = util_encontrar_columna(
                    df_historial,
                    ["Género", "Genero", "Sexo"]
                )

                if columna_sexo is not None:
                    df_historial["Sexo_normalizado"] = df_historial[
                        columna_sexo
                    ].apply(hist_normalizar_sexo)
                else:
                    df_historial["Sexo_normalizado"] = "Sin especificar"

                df_historial["Rango_promedio"] = df_historial[
                    "Promedio_normalizado_100"
                ].apply(hist_clasificar_rango_promedio)

                columna_escuela = util_encontrar_columna(
                    df_historial,
                    [
                        "Escuela de Procedencia",
                        "Escuela Procedencia",
                        "Procedencia",
                        "Escuela"
                    ]
                )

                if columna_escuela is not None:
                    df_historial["Bachillerato_procedencia_original"] = df_historial[
                        columna_escuela
                    ].fillna("Sin dato").astype(str)

                    df_historial["Bachillerato_procedencia"] = df_historial[
                        columna_escuela
                    ].apply(hist_normalizar_escuela_procedencia)

                    df_historial["Estado_procedencia"] = df_historial[
                        columna_escuela
                    ].apply(hist_clasificar_estado_procedencia)

                else:
                    df_historial["Bachillerato_procedencia_original"] = "Sin dato"
                    df_historial["Bachillerato_procedencia"] = "Sin dato"
                    df_historial["Estado_procedencia"] = "Sin dato"

                st.session_state["df_historial_global"] = df_historial
                st.session_state["df_bitacora_historial"] = df_bitacora

                st.success(
                    f"Historial cargado correctamente: {len(df_historial):,} aspirantes."
                )

        except Exception as error:
            st.error(f"No fue posible procesar el Historial: {error}")

    elif "df_historial_global" in st.session_state:
        st.success(
            f"Historial ya cargado: {len(st.session_state['df_historial_global']):,} aspirantes."
        )
    else:
        st.info("Aún no se ha cargado el Historial de Aspirantes.")

    st.markdown("---")

    # ------------------------------------------------------------
    # EVALUATEC
    # ------------------------------------------------------------

    st.markdown("## 2. EVALUATEC 2026")

    archivos_eval = st.file_uploader(
        "Carga los 3 archivos CSV de EVALUATEC",
        type=["csv"],
        accept_multiple_files=True,
        key="global_eval_archivos"
    )

    if archivos_eval:
        if len(archivos_eval) != 3:
            st.warning(
                f"Cargaste {len(archivos_eval)} archivo(s). Se requieren exactamente 3."
            )
        else:
            datos_por_bloque = {}
            errores = []

            for archivo in archivos_eval:
                try:
                    df_archivo, areas_detectadas = eval_procesar_archivo(
                        archivo
                    )

                    bloque = df_archivo["Bloque"].iloc[0]

                    datos_por_bloque[bloque] = {
                        "df": df_archivo,
                        "areas": areas_detectadas,
                        "archivo": archivo.name
                    }

                except Exception as error:
                    errores.append(f"{archivo.name}: {error}")

            if errores:
                for error in errores:
                    st.error(error)

            if datos_por_bloque:
                st.session_state["datos_eval_global"] = datos_por_bloque

                st.success(
                    f"EVALUATEC cargado correctamente: {len(datos_por_bloque)} bloques."
                )

    elif "datos_eval_global" in st.session_state:
        st.success(
            f"EVALUATEC ya cargado: {len(st.session_state['datos_eval_global'])} bloques."
        )
    else:
        st.info("Aún no se han cargado los archivos EVALUATEC.")

    st.markdown("---")

    # ------------------------------------------------------------
    # CHASIDE
    # ------------------------------------------------------------

    st.markdown("## 3. CHASIDE")

    url_chaside = st.text_input(
        "Pega el enlace de respuestas de Google Forms / Google Sheets",
        value=st.session_state.get(
            "url_chaside_global",
            "https://docs.google.com/spreadsheets/u/2/d/1YHMEb5hftOZfV-CMWoUsUgJh1xmsgTY3YYwAtq1dGQA/edit?resourcekey&gid=1491376423#gid=1491376423"
        ),
        key="global_url_chaside"
    )

    st.session_state["url_chaside_global"] = url_chaside

    peso_intereses = st.slider(
        "Peso de intereses",
        min_value=0.0,
        max_value=1.0,
        value=st.session_state.get("peso_intereses_chaside", 0.8),
        step=0.1,
        key="global_peso_intereses_chaside"
    )

    peso_aptitudes = round(1 - peso_intereses, 2)

    st.session_state["peso_intereses_chaside"] = peso_intereses
    st.session_state["peso_aptitudes_chaside"] = peso_aptitudes

    st.caption(
        f"Pesos activos → Intereses: {peso_intereses:.1f} | "
        f"Aptitudes: {peso_aptitudes:.1f}"
    )

    if st.button("Procesar respuestas CHASIDE", use_container_width=True):
        try:
            with st.spinner("Cargando respuestas CHASIDE desde Google Forms..."):
                df_chaside_raw = chaside_cargar_respuestas(url_chaside)

                resultado_chaside = chaside_procesar_respuestas(
                    df_chaside_raw,
                    peso_intereses=peso_intereses,
                    peso_aptitudes=peso_aptitudes
                )

            if isinstance(resultado_chaside, tuple):
                df_chaside = resultado_chaside[0]
            else:
                df_chaside = resultado_chaside

            if not isinstance(df_chaside, pd.DataFrame):
                raise ValueError(
                    "La función chaside_procesar_respuestas no regresó un DataFrame válido."
                )

            st.session_state["df_chaside_raw_global"] = df_chaside_raw
            st.session_state["df_chaside_global"] = df_chaside

            total_chaside = int(df_chaside.shape[0])

            st.success(
                f"CHASIDE cargado correctamente: {total_chaside} respuestas."
            )

        except Exception as error:
            st.error(f"No fue posible cargar/procesar CHASIDE: {error}")

    elif "df_chaside_global" in st.session_state:
        total_chaside = int(st.session_state["df_chaside_global"].shape[0])

        st.success(
            f"CHASIDE ya cargado: {total_chaside} respuestas."
        )

    else:
        st.info("Aún no se han procesado las respuestas CHASIDE.")
    # ------------------------------------------------------------
    # Estado general
    # ------------------------------------------------------------

    st.markdown("## Estado general de carga")

    estado_historial = (
        "✅ Cargado"
        if "df_historial_global" in st.session_state
        else "❌ Pendiente"
    )

    estado_eval = (
        "✅ Cargado"
        if "datos_eval_global" in st.session_state
        else "❌ Pendiente"
    )

    estado_chaside = (
        "✅ Cargado"
        if "df_chaside_global" in st.session_state
        else "❌ Pendiente"
    )

    col1, col2, col3 = st.columns(3)

    col1.metric("Historial", estado_historial)
    col2.metric("EVALUATEC", estado_eval)
    col3.metric("CHASIDE", estado_chaside)
    
# ============================================================
# ============================================================
# MÓDULO 1: EVALUATEC 2026
# ============================================================
# ============================================================

EVAL_ETIQUETAS_AREAS = {
    "ING": "Inglés",
    "MAT": "Matemáticas",
    "COM": "Comprensión lectora",
    "RLM": "Razonamiento lógico-matemático",
    "PM": "Pensamiento matemático",
    "ARQ": "Arquitectura",
    "FIS": "Física",
    "ADMN": "Administración"
}

EVAL_ORDEN_AREAS = [
    "ING",
    "MAT",
    "COM",
    "RLM",
    "PM",
    "FIS",
    "ARQ",
    "ADMN"
]

EVAL_BLOQUES = {
    "ADM": "Administración",
    "ARQ": "Arquitectura",
    "ING": "Ingeniería"
}

EVAL_ICONOS_BLOQUES = {
    "ADM": "📘",
    "ARQ": "🏛️",
    "ING": "⚙️"
}

EVAL_ORDEN_NIVELES = [
    "Bajo",
    "Básico",
    "Satisfactorio",
    "Alto"
]

EVAL_RANGOS_NIVELES = {
    "Bajo": "0–24%",
    "Básico": "25–49%",
    "Satisfactorio": "50–74%",
    "Alto": "75–100%"
}

EVAL_COLORES_NIVELES = {
    "Bajo": "#E74C3C",
    "Básico": "#F39C12",
    "Satisfactorio": "#F1C40F",
    "Alto": "#27AE60"
}


def eval_limpiar_nombre_carrera(valor):
    if pd.isna(valor):
        return "Sin carrera especificada"

    return " ".join(str(valor).strip().split())


def eval_leer_csv_archivo(archivo):
    contenido = archivo.getvalue()

    codificaciones = [
        "utf-8",
        "utf-8-sig",
        "latin-1",
        "cp1252"
    ]

    separadores = [
        ",",
        ";",
        "\t"
    ]

    for codificacion in codificaciones:
        for separador in separadores:
            try:
                df = pd.read_csv(
                    io.BytesIO(contenido),
                    encoding=codificacion,
                    sep=separador
                )

                if len(df.columns) > 1:
                    return df

            except Exception:
                continue

    return pd.read_csv(
        io.BytesIO(contenido),
        encoding="latin-1"
    )


def eval_identificar_bloque_archivo(nombre_archivo):
    nombre = util_normalizar_texto(nombre_archivo)

    if "administracion" in nombre:
        return "ADM"

    if "arquitectura" in nombre:
        return "ARQ"

    if "ingenieria" in nombre:
        return "ING"

    return None


def eval_clasificar_inicio(valor):
    if pd.isna(valor):
        return "No inició"

    texto = util_normalizar_texto(valor)

    valores_no_inicio = [
        "",
        "no",
        "n",
        "false",
        "falso",
        "0",
        "no inicio",
        "no iniciado",
        "pendiente",
        "null",
        "nan",
        "none"
    ]

    if texto in valores_no_inicio:
        return "No inició"

    if "no inicio" in texto:
        return "No inició"

    return "Inició"


def eval_convertir_porcentaje(valor):
    if pd.isna(valor):
        return np.nan

    texto = str(valor).strip()

    if texto == "":
        return np.nan

    texto = texto.replace("%", "")
    texto = texto.replace(",", ".")

    try:
        numero = float(texto)
    except ValueError:
        return np.nan

    if 0 <= numero <= 1:
        return numero * 100

    if 0 <= numero <= 100:
        return numero

    return np.nan


def eval_hex_a_rgba(color_hex, alpha=0.15):
    color_hex = color_hex.lstrip("#")

    if len(color_hex) != 6:
        return f"rgba(120, 120, 120, {alpha})"

    rojo = int(color_hex[0:2], 16)
    verde = int(color_hex[2:4], 16)
    azul = int(color_hex[4:6], 16)

    return f"rgba({rojo}, {verde}, {azul}, {alpha})"


def eval_detectar_columnas_areas(df):
    areas_detectadas = {}

    for columna in df.columns:
        texto = util_normalizar_texto(columna)

        texto_compacto = re.sub(
            r"[^a-z0-9]",
            "",
            texto
        )

        if "seccion" not in texto_compacto:
            continue

        if "porcentajecorrectas" not in texto_compacto:
            continue

        coincidencia = re.search(
            r"seccion([a-z0-9]+?)porcentajecorrectas",
            texto_compacto
        )

        if coincidencia:
            codigo = coincidencia.group(1).upper()
            areas_detectadas[codigo] = columna

    areas_ordenadas = {}

    for codigo in EVAL_ORDEN_AREAS:
        if codigo in areas_detectadas:
            areas_ordenadas[codigo] = areas_detectadas[codigo]

    for codigo, columna in areas_detectadas.items():
        if codigo not in areas_ordenadas:
            areas_ordenadas[codigo] = columna

    return areas_ordenadas


def eval_procesar_archivo(archivo):
    df = eval_leer_csv_archivo(archivo)

    bloque = eval_identificar_bloque_archivo(archivo.name)

    if bloque is None:
        raise ValueError(
            "No se identificó el bloque académico. "
            "El archivo debe contener Administración, Arquitectura o Ingeniería."
        )

    columna_carrera = util_encontrar_columna(
        df,
        ["Carrera"]
    )

    columna_inicio = util_encontrar_columna(
        df,
        [
            "InicioExamen",
            "Inicio Examen",
            "Inició examen",
            "Inicio"
        ]
    )

    if columna_carrera is None:
        raise ValueError(
            f"{archivo.name}: no se encontró la columna Carrera."
        )

    if columna_inicio is None:
        raise ValueError(
            f"{archivo.name}: no se encontró la columna InicioExamen."
        )

    areas_detectadas = eval_detectar_columnas_areas(df)

    if not areas_detectadas:
        raise ValueError(
            f"{archivo.name}: no se detectaron columnas de áreas evaluadas."
        )

    df["Archivo_origen"] = archivo.name
    df["Bloque"] = bloque

    df["Carrera_normalizada"] = df[
        columna_carrera
    ].apply(
        eval_limpiar_nombre_carrera
    )

    df["Estatus_inicio"] = df[
        columna_inicio
    ].apply(
        eval_clasificar_inicio
    )

    for codigo, columna in areas_detectadas.items():
        df[f"Area_{codigo}"] = df[
            columna
        ].apply(
            eval_convertir_porcentaje
        )

    columnas_areas = [
        f"Area_{codigo}"
        for codigo in areas_detectadas.keys()
    ]

    df["Promedio_global_individual"] = df[
        columnas_areas
    ].mean(axis=1)

    return df, areas_detectadas


def eval_clasificar_nivel_desempeno(valor):
    if pd.isna(valor):
        return None

    if 0 <= valor < 25:
        return "Bajo"

    if 25 <= valor < 50:
        return "Básico"

    if 50 <= valor < 75:
        return "Satisfactorio"

    if 75 <= valor <= 100:
        return "Alto"

    return None


def eval_crear_promedio_dimensiones(df, areas_detectadas):
    df_iniciaron = df[
        df["Estatus_inicio"] == "Inició"
    ].copy()

    resultados = []

    for codigo in areas_detectadas.keys():
        columna = f"Area_{codigo}"

        if columna not in df_iniciaron.columns:
            continue

        promedio = df_iniciaron[columna].mean()

        if pd.notna(promedio):
            resultados.append(
                {
                    "Código": codigo,
                    "Dimensión": EVAL_ETIQUETAS_AREAS.get(
                        codigo,
                        codigo
                    ),
                    "Promedio": round(float(promedio), 1)
                }
            )

    return pd.DataFrame(resultados)


def eval_crear_distribucion_por_dimension(df, areas_detectadas):
    df_iniciaron = df[
        df["Estatus_inicio"] == "Inició"
    ].copy()

    registros = []

    for codigo in areas_detectadas.keys():
        columna = f"Area_{codigo}"

        if columna not in df_iniciaron.columns:
            continue

        valores = df_iniciaron[columna].dropna()

        if valores.empty:
            continue

        total = len(valores)

        niveles = valores.apply(
            eval_clasificar_nivel_desempeno
        )

        conteos = niveles.value_counts()

        for nivel in EVAL_ORDEN_NIVELES:
            cantidad = int(conteos.get(nivel, 0))
            porcentaje = cantidad / total * 100

            registros.append(
                {
                    "Código": codigo,
                    "Dimensión": EVAL_ETIQUETAS_AREAS.get(
                        codigo,
                        codigo
                    ),
                    "Nivel": nivel,
                    "Aspirantes": cantidad,
                    "Total": total,
                    "Porcentaje": porcentaje
                }
            )

    tabla = pd.DataFrame(registros)

    if tabla.empty:
        return tabla

    promedios = eval_crear_promedio_dimensiones(
        df,
        areas_detectadas
    )

    tabla = tabla.merge(
        promedios[
            [
                "Código",
                "Promedio"
            ]
        ],
        on="Código",
        how="left"
    )

    orden_codigo = {
        codigo: indice
        for indice, codigo in enumerate(EVAL_ORDEN_AREAS)
    }

    tabla["Orden"] = tabla["Código"].map(
        orden_codigo
    ).fillna(999)

    tabla = tabla.sort_values(
        [
            "Orden",
            "Nivel"
        ]
    )

    tabla["Nivel"] = pd.Categorical(
        tabla["Nivel"],
        categories=EVAL_ORDEN_NIVELES,
        ordered=True
    )

    tabla["Etiqueta"] = tabla["Porcentaje"].apply(
        lambda valor: f"{valor:.0f}%"
        if valor >= 8
        else ""
    )

    return tabla


def eval_crear_diagnostico_carrera(
    df_carrera,
    df_bloque,
    areas_detectadas,
    carrera_seleccionada
):
    promedio_carrera = eval_crear_promedio_dimensiones(
        df_carrera,
        areas_detectadas
    )

    if promedio_carrera.empty:
        return "No hay información suficiente para generar un diagnóstico."

    ranking = promedio_carrera.sort_values(
        "Promedio",
        ascending=True
    ).reset_index(drop=True)

    prioridades = ranking.head(2)
    fortaleza = ranking.iloc[-1]

    promedio_global_carrera = df_carrera[
        (
            df_carrera["Estatus_inicio"] == "Inició"
        )
        &
        (
            df_carrera["Promedio_global_individual"].notna()
        )
    ][
        "Promedio_global_individual"
    ].mean()

    promedio_global_bloque = df_bloque[
        (
            df_bloque["Estatus_inicio"] == "Inició"
        )
        &
        (
            df_bloque["Promedio_global_individual"].notna()
        )
    ][
        "Promedio_global_individual"
    ].mean()

    diferencia = promedio_global_carrera - promedio_global_bloque

    if diferencia >= 0:
        comparacion = f"{diferencia:.1f} puntos por encima"
    else:
        comparacion = f"{abs(diferencia):.1f} puntos por debajo"

    texto_prioridades = ", ".join(
        [
            f"{fila['Dimensión']} ({fila['Promedio']:.1f}%)"
            for _, fila in prioridades.iterrows()
        ]
    )

    bloque = df_bloque["Bloque"].iloc[0]

    return (
        f"**{carrera_seleccionada}** presenta un promedio global de "
        f"**{promedio_global_carrera:.1f}%**, "
        f"{comparacion} del promedio general de "
        f"**{EVAL_BLOQUES[bloque]}**. "
        f"Las principales áreas de fortalecimiento son "
        f"**{texto_prioridades}**. "
        f"La dimensión con mejor resultado es "
        f"**{fortaleza['Dimensión']}** "
        f"({fortaleza['Promedio']:.1f}%)."
    )


def eval_mostrar_radar_carrera(
    df_carrera,
    df_bloque,
    areas_detectadas,
    carrera_seleccionada,
    nombre_bloque
):
    promedio_carrera = eval_crear_promedio_dimensiones(
        df_carrera,
        areas_detectadas
    )

    promedio_bloque = eval_crear_promedio_dimensiones(
        df_bloque,
        areas_detectadas
    )

    if promedio_carrera.empty:
        st.info("No hay datos suficientes para generar el radar.")
        return

    orden_codigo = {
        codigo: indice
        for indice, codigo in enumerate(EVAL_ORDEN_AREAS)
    }

    promedio_carrera = promedio_carrera.sort_values(
        "Código",
        key=lambda serie: serie.map(orden_codigo)
    )

    codigos = promedio_carrera["Código"].tolist()
    etiquetas = promedio_carrera["Dimensión"].tolist()
    valores_carrera = promedio_carrera["Promedio"].tolist()

    valores_bloque = []

    for codigo in codigos:
        fila_bloque = promedio_bloque[
            promedio_bloque["Código"] == codigo
        ]

        if fila_bloque.empty:
            valores_bloque.append(0)
        else:
            valores_bloque.append(
                float(
                    fila_bloque["Promedio"].iloc[0]
                )
            )

    ranking_bajo = promedio_carrera.sort_values(
        "Promedio",
        ascending=True
    ).head(2)

    etiquetas_bajas = ranking_bajo["Dimensión"].tolist()
    valores_bajos = ranking_bajo["Promedio"].tolist()

    etiquetas_cerradas = etiquetas + [etiquetas[0]]
    valores_carrera_cerrados = valores_carrera + [
        valores_carrera[0]
    ]
    valores_bloque_cerrados = valores_bloque + [
        valores_bloque[0]
    ]

    fig = go.Figure()

    fig.add_trace(
        go.Scatterpolar(
            r=valores_bloque_cerrados,
            theta=etiquetas_cerradas,
            mode="lines+markers",
            name=f"Promedio {nombre_bloque}",
            line=dict(
                color="#9E9E9E",
                width=2,
                dash="dash"
            ),
            marker=dict(
                color="#9E9E9E",
                size=6
            ),
            hovertemplate=(
                "<b>Promedio del archivo</b><br>"
                "%{theta}: %{r:.1f}%"
                "<extra></extra>"
            )
        )
    )

    fig.add_trace(
        go.Scatterpolar(
            r=valores_carrera_cerrados,
            theta=etiquetas_cerradas,
            mode="lines+markers",
            name=carrera_seleccionada,
            line=dict(
                color="#4C78A8",
                width=4
            ),
            marker=dict(
                color="#4C78A8",
                size=8
            ),
            fill="toself",
            fillcolor=eval_hex_a_rgba(
                "#4C78A8",
                alpha=0.14
            ),
            hovertemplate=(
                f"<b>{carrera_seleccionada}</b><br>"
                "%{theta}: %{r:.1f}%"
                "<extra></extra>"
            )
        )
    )

    fig.add_trace(
        go.Scatterpolar(
            r=valores_bajos,
            theta=etiquetas_bajas,
            mode="markers",
            name="Áreas prioritarias",
            marker=dict(
                color="#E74C3C",
                size=14,
                line=dict(
                    color="white",
                    width=2
                )
            ),
            hovertemplate=(
                "<b>Área prioritaria</b><br>"
                "%{theta}: %{r:.1f}%"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(
        title=f"Perfil de dimensiones · {carrera_seleccionada}",
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 100],
                ticksuffix="%"
            )
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.15,
            xanchor="center",
            x=0.5
        ),
        height=560,
        margin=dict(
            t=80,
            b=80,
            l=40,
            r=40
        )
    )

    col_grafica, col_prioridades = st.columns(
        [3, 1]
    )

    with col_grafica:
        st.plotly_chart(
            fig,
            use_container_width=True
        )

    with col_prioridades:
        st.markdown("### 🔴 Áreas prioritarias")
        st.caption(
            "Dimensiones con menor promedio en la carrera."
        )

        for _, fila in ranking_bajo.iterrows():
            st.metric(
                fila["Dimensión"],
                f"{fila['Promedio']:.1f}%"
            )

        st.markdown("---")

        st.caption(
            f"Línea gris: promedio de {nombre_bloque}. "
            "Puntos rojos: dimensiones prioritarias."
        )


def eval_mostrar_barras_distribucion_dimension(
    df_carrera,
    areas_detectadas,
    carrera_seleccionada
):
    tabla = eval_crear_distribucion_por_dimension(
        df_carrera,
        areas_detectadas
    )

    if tabla.empty:
        st.info(
            "No hay información suficiente para generar la gráfica."
        )
        return

    dimensiones = (
        tabla[
            [
                "Código",
                "Dimensión",
                "Orden"
            ]
        ]
        .drop_duplicates()
        .sort_values("Orden")
    )

    nombres_dimensiones = dimensiones[
        "Dimensión"
    ].tolist()

    promedios = (
        tabla[
            [
                "Dimensión",
                "Promedio"
            ]
        ]
        .drop_duplicates()
        .set_index("Dimensión")
        .reindex(nombres_dimensiones)
        .reset_index()
    )

    fig = go.Figure()

    for nivel in EVAL_ORDEN_NIVELES:
        datos_nivel = tabla[
            tabla["Nivel"] == nivel
        ].copy()

        datos_nivel = (
            datos_nivel
            .set_index("Dimensión")
            .reindex(nombres_dimensiones)
            .reset_index()
        )

        fig.add_trace(
            go.Bar(
                x=datos_nivel["Dimensión"],
                y=datos_nivel["Porcentaje"],
                name=f"{nivel} · {EVAL_RANGOS_NIVELES[nivel]}",
                marker_color=EVAL_COLORES_NIVELES[nivel],
                text=datos_nivel["Etiqueta"],
                textposition="inside",
                insidetextanchor="middle",
                customdata=np.column_stack(
                    [
                        datos_nivel["Aspirantes"],
                        datos_nivel["Total"],
                        datos_nivel["Porcentaje"]
                    ]
                ),
                hovertemplate=(
                    "<b>%{x}</b><br>"
                    f"<b>Rango de calificación:</b> "
                    f"{EVAL_RANGOS_NIVELES[nivel]}<br>"
                    "<b>Aspirantes:</b> %{customdata[0]} de %{customdata[1]}<br>"
                    "<b>Distribución:</b> %{customdata[2]:.1f}%"
                    "<extra></extra>"
                )
            )
        )

    fig.add_trace(
        go.Scatter(
            x=promedios["Dimensión"],
            y=promedios["Promedio"],
            yaxis="y2",
            mode="markers+text",
            name="Promedio real obtenido",
            text=[
                f"{valor:.1f}%"
                for valor in promedios["Promedio"]
            ],
            textposition="top center",
            marker=dict(
                color="#111111",
                size=11,
                symbol="diamond",
                line=dict(
                    color="white",
                    width=1.5
                )
            ),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "<b>Promedio real obtenido:</b> %{y:.1f}%"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(
        title=(
            "Distribución de resultados y promedio real por dimensión · "
            f"{carrera_seleccionada}"
        ),
        barmode="stack",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.12,
            xanchor="center",
            x=0.5
        ),
        xaxis=dict(
            title="Dimensiones",
            categoryorder="array",
            categoryarray=nombres_dimensiones
        ),
        yaxis=dict(
            title="Distribución de aspirantes",
            range=[0, 100],
            ticksuffix="%"
        ),
        yaxis2=dict(
            title="Promedio de calificación obtenida",
            overlaying="y",
            side="right",
            range=[0, 100],
            ticksuffix="%",
            showgrid=False
        ),
        height=620,
        margin=dict(
            t=115,
            b=80,
            l=65,
            r=75
        )
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    st.caption(
        "Las barras muestran cómo se distribuyen los aspirantes en los rangos "
        "de calificación. El rombo negro indica el promedio real obtenido "
        "en cada dimensión."
    )


def render_evaluatec():
    st.title("📘 Resultados EVALUATEC 2026")
    st.caption("Perfil académico por carrera.")

    st.sidebar.markdown("---")
    st.sidebar.subheader("Carga EVALUATEC")

    archivos_subidos = st.sidebar.file_uploader(
        "Carga los 3 archivos oficiales EVALUATEC",
        type=["csv"],
        accept_multiple_files=True,
        key="eval_archivos"
    )

    if not archivos_subidos:
        st.info(
            "Carga los tres archivos CSV: Administración, Arquitectura e Ingeniería."
        )
        st.stop()

    if len(archivos_subidos) != 3:
        st.warning(
            f"Actualmente cargaste {len(archivos_subidos)} archivo(s). "
            "Deben cargarse exactamente 3."
        )
        st.stop()

    datos_por_bloque = {}
    errores = []

    for archivo in archivos_subidos:
        try:
            df_archivo, areas_detectadas = eval_procesar_archivo(
                archivo
            )

            bloque = df_archivo["Bloque"].iloc[0]

            datos_por_bloque[bloque] = {
                "df": df_archivo,
                "areas": areas_detectadas,
                "archivo": archivo.name
            }

        except Exception as error:
            errores.append(
                f"{archivo.name}: {error}"
            )

    if errores:
        for error in errores:
            st.error(error)

    if not datos_por_bloque:
        st.stop()

    bloques_disponibles = [
        codigo
        for codigo in EVAL_BLOQUES
        if codigo in datos_por_bloque
    ]

    bloque_seleccionado = st.radio(
        "Selecciona el archivo o bloque académico",
        options=bloques_disponibles,
        horizontal=True,
        format_func=lambda codigo: (
            f"{EVAL_ICONOS_BLOQUES[codigo]} {EVAL_BLOQUES[codigo]}"
        ),
        label_visibility="collapsed",
        key="eval_bloque"
    )

    informacion_bloque = datos_por_bloque[
        bloque_seleccionado
    ]

    df_bloque = informacion_bloque["df"].copy()
    areas_detectadas = informacion_bloque["areas"]
    nombre_bloque = EVAL_BLOQUES[bloque_seleccionado]

    st.markdown(f"## {nombre_bloque}")

    st.caption(
        f"Archivo analizado: {informacion_bloque['archivo']}"
    )

    carreras_disponibles = sorted(
        df_bloque["Carrera_normalizada"]
        .dropna()
        .unique()
    )

    carrera_seleccionada = st.selectbox(
        "Selecciona la carrera",
        options=carreras_disponibles,
        key=f"eval_carrera_{bloque_seleccionado}"
    )

    df_carrera = df_bloque[
        df_bloque["Carrera_normalizada"]
        == carrera_seleccionada
    ].copy()

    total_registrados = len(df_carrera)

    total_iniciaron = int(
        (
            df_carrera["Estatus_inicio"] == "Inició"
        ).sum()
    )

    total_no_iniciaron = int(
        (
            df_carrera["Estatus_inicio"] == "No inició"
        ).sum()
    )

    porcentaje_inicio = (
        total_iniciaron / total_registrados * 100
        if total_registrados > 0
        else 0
    )

    promedio_global = df_carrera[
        (
            df_carrera["Estatus_inicio"] == "Inició"
        )
        &
        (
            df_carrera["Promedio_global_individual"].notna()
        )
    ][
        "Promedio_global_individual"
    ].mean()

    st.markdown(f"### Perfil de {carrera_seleccionada}")

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("Registrados", f"{total_registrados:,}")
    col2.metric("Iniciaron", f"{total_iniciaron:,}")
    col3.metric("No iniciaron", f"{total_no_iniciaron:,}")
    col4.metric("% de inicio", f"{porcentaje_inicio:.1f}%")
    col5.metric(
        "Promedio global",
        f"{promedio_global:.1f}%"
        if pd.notna(promedio_global)
        else "Sin dato"
    )

    st.markdown("## Perfil de dimensiones")

    eval_mostrar_radar_carrera(
        df_carrera=df_carrera,
        df_bloque=df_bloque,
        areas_detectadas=areas_detectadas,
        carrera_seleccionada=carrera_seleccionada,
        nombre_bloque=nombre_bloque
    )

    st.markdown("## Distribución de calificaciones por dimensión")

    eval_mostrar_barras_distribucion_dimension(
        df_carrera=df_carrera,
        areas_detectadas=areas_detectadas,
        carrera_seleccionada=carrera_seleccionada
    )

    st.markdown("## Diagnóstico ejecutivo")

    diagnostico = eval_crear_diagnostico_carrera(
        df_carrera=df_carrera,
        df_bloque=df_bloque,
        areas_detectadas=areas_detectadas,
        carrera_seleccionada=carrera_seleccionada
    )

    st.info(diagnostico)


# ============================================================
# ============================================================
# MÓDULO 2: HISTORIAL DE ASPIRANTES
# ============================================================
# ============================================================

def hist_nombres_unicos(encabezados):
    usados = {}
    resultado = []

    for posicion, encabezado in enumerate(encabezados, start=1):

        if pd.isna(encabezado) or str(encabezado).strip() == "":
            nombre = f"Columna_sin_nombre_{posicion}"
        else:
            nombre = str(encabezado).strip()

        if nombre in usados:
            usados[nombre] += 1
            nombre = f"{nombre}_{usados[nombre]}"
        else:
            usados[nombre] = 1

        resultado.append(nombre)

    return resultado


def hist_buscar_fila_encabezados(df_crudo):
    palabras_clave = [
        "matricula/id",
        "matricula",
        "id",
        "apellido paterno",
        "apellido materno",
        "nombre (s)",
        "nombre"
    ]

    limite = min(len(df_crudo), 40)

    for indice in range(limite):

        valores = [
            util_limpiar_texto(valor)
            for valor in df_crudo.iloc[indice].tolist()
        ]

        coincidencias = sum(
            any(palabra in valor for valor in valores)
            for palabra in palabras_clave
        )

        if coincidencias >= 2:
            return indice

    return None


def hist_obtener_nombre_carrera(nombre_hoja, df_crudo):
    limite = min(len(df_crudo), 15)

    for indice in range(limite):

        fila = df_crudo.iloc[indice].tolist()

        for posicion, valor in enumerate(fila):

            if util_limpiar_texto(valor) == "carrera":

                if posicion + 1 < len(fila):
                    posible_carrera = fila[posicion + 1]

                    if pd.notna(posible_carrera):
                        return str(posible_carrera).strip()

    return str(nombre_hoja).strip()


def hist_convertir_promedio(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return np.nan, "Sin dato"

    if isinstance(valor, (datetime, date, pd.Timestamp)):
        return np.nan, "Dato dudoso: formato fecha"

    texto = str(valor).strip()
    texto = texto.replace("\xa0", " ")
    texto = texto.replace(",", ".")
    texto = texto.lstrip("'").strip()

    try:
        numero = float(texto)

    except (TypeError, ValueError):
        return np.nan, "Dato dudoso: no numérico"

    if 0 <= numero <= 10:
        return round(numero * 10, 2), "Convertido de escala 0-10"

    if 10 < numero <= 100:
        return round(numero, 2), "Válido: escala 0-100"

    return np.nan, "Dato dudoso: fuera de rango"


def hist_clasificar_rango_promedio(valor):
    if pd.isna(valor):
        return "Sin dato"

    if 60 <= valor < 70:
        return "60-69"

    if 70 <= valor < 80:
        return "70-79"

    if 80 <= valor < 90:
        return "80-89"

    if 90 <= valor <= 100:
        return "90-100"

    return "Fuera de rango"


def hist_normalizar_sexo(valor):
    if pd.isna(valor):
        return "Sin especificar"

    texto = util_limpiar_texto(valor)

    if texto in ["hombre", "masculino", "m", "male"]:
        return "Hombre"

    if texto in ["mujer", "femenino", "f", "female"]:
        return "Mujer"

    return "Sin especificar"


def hist_clasificar_estado_procedencia(valor):
    if pd.isna(valor):
        return "Sin dato"

    texto = util_limpiar_texto(valor)

    if texto in ["", "nan", "none", "escuela de procedencia"]:
        return "Sin dato"

    palabras_jalisco = [
        "jalisco",
        "tuxpan",
        "cihuatlan",
        "autlan",
        "guadalajara",
        "zapopan",
        "tonala",
        "sayula",
        "zapotiltic",
        "zapotlan",
        "ciudad guzman",
        "tequila",
        "casimiro castillo",
        "el grullo",
        "union de tula",
        "tamazula",
        "teocuitatlan",
        "universidad de guadalajara",
        "udeg"
    ]

    if any(palabra in texto for palabra in palabras_jalisco):
        return "Jalisco"

    palabras_michoacan = [
        "michoacan",
        "coahuayana",
        "coalcoman",
        "morelia",
        "zamora",
        "lazaro cardenas",
        "uruapan",
        "apatzingan",
        "maravatio"
    ]

    if any(palabra in texto for palabra in palabras_michoacan):
        return "Michoacán"

    palabras_nayarit = [
        "nayarit",
        "tepic",
        "bahia de banderas",
        "santiago ixcuintla",
        "compostela"
    ]

    if any(palabra in texto for palabra in palabras_nayarit):
        return "Nayarit"

    palabras_guanajuato = [
        "guanajuato",
        "leon",
        "irapuato",
        "celaya",
        "salamanca"
    ]

    if any(palabra in texto for palabra in palabras_guanajuato):
        return "Guanajuato"

    if "nuevo leon" in texto or "monterrey" in texto:
        return "Nuevo León"

    if "sinaloa" in texto or "culiacan" in texto:
        return "Sinaloa"

    if "durango" in texto:
        return "Durango"

    if "sonora" in texto or "hermosillo" in texto:
        return "Sonora"

    if "baja california" in texto or "tijuana" in texto:
        return "Baja California"

    if "veracruz" in texto:
        return "Veracruz"

    if "ciudad de mexico" in texto or "cdmx" in texto:
        return "Ciudad de México"

    if any(
        palabra in texto
        for palabra in ["canada", "canadá", "usa", "united states"]
    ):
        return "Internacional"

    return "Colima"


def hist_obtener_numero_institucion(texto, expresiones):
    for expresion in expresiones:

        coincidencia = re.search(expresion, texto)

        if coincidencia:
            return coincidencia.group(1)

    return None


def hist_normalizar_escuela_procedencia(valor):
    if pd.isna(valor):
        return "Sin dato"

    texto_visible = util_limpiar_texto_visible(valor)
    texto = util_limpiar_texto(valor)
    texto_compacto = re.sub(r"[^a-z0-9]", "", texto)

    if texto in ["", "nan", "none", "escuela de procedencia"]:
        return "Sin dato"

    if (
        "universidad de colima" in texto
        or "u de c" in texto
        or "udec" in texto
        or "bachillerato udec" in texto
        or re.search(r"\bbachillerato\s*([1-9]|[12][0-9]|30)\b", texto)
    ):
        return "Universidad de Colima (U de C)"

    if (
        "telebachillerato" in texto
        or "tele bachillerato" in texto
        or "telebach" in texto
        or "telebach" in texto_compacto
    ):
        return "Telebachillerato"

    if (
        "colegio de bachilleres" in texto
        or "colegio bachilleres" in texto
        or "colegio de bach" in texto
        or "colegio bach" in texto
        or "cobach" in texto_compacto
        or "coba" in texto_compacto
    ):
        return "Colegio de Bachilleres"

    if "cbtis" in texto_compacto or "cbti" in texto_compacto:

        numero = hist_obtener_numero_institucion(
            texto,
            [
                r"cbtis\s*#?\s*(\d+)",
                r"cbti[s]?\s*#?\s*(\d+)"
            ]
        )

        if numero:
            return f"CBTis {numero}"

        return "CBTis"

    if "cetis" in texto_compacto:

        numero = hist_obtener_numero_institucion(
            texto,
            [r"cetis\s*#?\s*(\d+)"]
        )

        if numero:
            return f"CETis {numero}"

        return "CETis"

    if "cbta" in texto_compacto:

        numero = hist_obtener_numero_institucion(
            texto,
            [r"cbta\s*#?\s*(\d+)"]
        )

        if numero:
            return f"CBTA {numero}"

        return "CBTA"

    if "emsad" in texto_compacto:

        numero = hist_obtener_numero_institucion(
            texto,
            [r"emsad\s*#?\s*(\d+)"]
        )

        if numero:
            return f"EMSAD {numero}"

        return "EMSAD"

    if "isenco" in texto_compacto:
        return "ISENCO"

    if "conalep" in texto_compacto:
        return "CONALEP"

    if "cecyte" in texto_compacto:
        return "CECyTE"

    if "icep" in texto_compacto:
        return "ICEP"

    if (
        "universidad de guadalajara" in texto
        or "udeg" in texto_compacto
        or "prepa regional tuxpan" in texto
    ):
        return "Universidad de Guadalajara (UdeG)"

    if "anahuac" in texto:
        return "Preparatoria Anáhuac"

    if "campoverde" in texto_compacto or "campo verde" in texto:
        return "Colegio Campoverde"

    if "adonai" in texto:
        return "Instituto Adonai"

    if "prepa en linea" in texto:
        return "Prepa en Línea SEP"

    if "acredita" in texto and "bach" in texto:
        return "Acredita-Bach SEP"

    return texto_visible.title()


def hist_procesar_hoja(contenido_archivo, nombre_hoja):
    archivo = io.BytesIO(contenido_archivo)

    df_crudo = pd.read_excel(
        archivo,
        sheet_name=nombre_hoja,
        header=None,
        dtype=object
    )

    fila_encabezados = hist_buscar_fila_encabezados(df_crudo)

    if fila_encabezados is None:
        return None, {
            "Hoja": nombre_hoja,
            "Estatus": "No procesada",
            "Detalle": "No se identificó una fila de encabezados."
        }

    carrera = hist_obtener_nombre_carrera(nombre_hoja, df_crudo)

    encabezados = hist_nombres_unicos(
        df_crudo.iloc[fila_encabezados].tolist()
    )

    df = df_crudo.iloc[fila_encabezados + 1:].copy()
    df.columns = encabezados
    df = df.dropna(how="all").copy()

    columna_id = util_encontrar_columna(
        df,
        ["Matrícula/ID", "Matrícula", "ID"]
    )

    if columna_id is not None:
        df = df[df[columna_id].notna()].copy()

    df["Carrera"] = carrera
    df["Hoja_origen"] = nombre_hoja

    columna_promedio = util_encontrar_columna(
        df,
        [
            "Promedio Bachillerato",
            "Promedio de Bachillerato",
            "Promedio"
        ]
    )

    if columna_promedio is not None:

        df["Promedio_original"] = df[columna_promedio]

        resultado = df[columna_promedio].apply(hist_convertir_promedio)

        df["Promedio_normalizado_100"] = resultado.apply(
            lambda x: x[0]
        )

        df["Estatus_promedio"] = resultado.apply(
            lambda x: x[1]
        )

    else:
        df["Promedio_original"] = np.nan
        df["Promedio_normalizado_100"] = np.nan
        df["Estatus_promedio"] = "No se encontró columna de promedio"

    return df, {
        "Hoja": nombre_hoja,
        "Estatus": "Procesada",
        "Detalle": f"{len(df):,} aspirantes identificados."
    }


@st.cache_data(show_spinner=False)
def hist_procesar_archivo_excel(contenido_archivo):
    archivo = io.BytesIO(contenido_archivo)
    excel = pd.ExcelFile(archivo)

    bases = []
    bitacora = []

    for hoja in excel.sheet_names:

        df_hoja, resultado = hist_procesar_hoja(
            contenido_archivo,
            hoja
        )

        bitacora.append(resultado)

        if df_hoja is not None and not df_hoja.empty:
            bases.append(df_hoja)

    if not bases:
        return pd.DataFrame(), pd.DataFrame(bitacora)

    df_general = pd.concat(
        bases,
        ignore_index=True,
        sort=False
    )

    return df_general, pd.DataFrame(bitacora)


def hist_crear_tabla_calificaciones_por_sexo(df):
    orden_rangos = ["60-69", "70-79", "80-89", "90-100"]

    df_calificaciones = df[
        df["Rango_promedio"].isin(orden_rangos)
    ].copy()

    if df_calificaciones.empty:
        return pd.DataFrame()

    tabla = (
        df_calificaciones
        .groupby(["Sexo_normalizado", "Rango_promedio"])
        .size()
        .reset_index(name="Aspirantes")
    )

    tabla["Rango_promedio"] = pd.Categorical(
        tabla["Rango_promedio"],
        categories=orden_rangos,
        ordered=True
    )

    tabla["Porcentaje"] = (
        tabla
        .groupby("Sexo_normalizado")["Aspirantes"]
        .transform(lambda x: (x / x.sum()) * 100)
    )

    tabla["Etiqueta"] = tabla["Porcentaje"].apply(
        lambda valor: f"{valor:.1f}%" if valor >= 5 else ""
    )

    return tabla


def hist_crear_distribucion_calificaciones_bachillerato(df, top_n=10):
    orden_rangos = ["60-69", "70-79", "80-89", "90-100"]

    df_valido = df[
        (
            df["Bachillerato_procedencia"] != "Sin dato"
        )
        &
        (
            df["Rango_promedio"].isin(orden_rangos)
        )
    ].copy()

    if df_valido.empty:
        return pd.DataFrame()

    totales = (
        df_valido
        .groupby("Bachillerato_procedencia")
        .size()
        .reset_index(name="Total")
        .sort_values("Total", ascending=False)
    )

    top_bachilleratos = totales.head(top_n).copy()

    escuelas_top = top_bachilleratos[
        "Bachillerato_procedencia"
    ].tolist()

    df_valido = df_valido[
        df_valido["Bachillerato_procedencia"].isin(escuelas_top)
    ].copy()

    tabla = (
        df_valido
        .groupby(
            [
                "Bachillerato_procedencia",
                "Rango_promedio"
            ]
        )
        .size()
        .reset_index(name="Aspirantes")
    )

    tabla = tabla.merge(
        top_bachilleratos,
        on="Bachillerato_procedencia",
        how="left"
    )

    tabla["Porcentaje"] = (
        tabla["Aspirantes"]
        / tabla["Total"]
        * 100
    )

    tabla["Rango_promedio"] = pd.Categorical(
        tabla["Rango_promedio"],
        categories=orden_rangos,
        ordered=True
    )

    tabla["Etiqueta"] = tabla["Porcentaje"].apply(
        lambda valor: f"{valor:.0f}%" if valor >= 8 else ""
    )

    tabla["Escuela_etiqueta"] = tabla.apply(
        lambda fila: (
            f"{fila['Bachillerato_procedencia']} "
            f"(n={int(fila['Total'])})"
        ),
        axis=1
    )

    orden_escuelas = [
        f"{fila['Bachillerato_procedencia']} (n={int(fila['Total'])})"
        for _, fila in top_bachilleratos.iterrows()
    ]

    tabla["Escuela_etiqueta"] = pd.Categorical(
        tabla["Escuela_etiqueta"],
        categories=orden_escuelas[::-1],
        ordered=True
    )

    return tabla


def hist_crear_mapa_colores_carreras(df):
    paleta = (
        px.colors.qualitative.Alphabet
        + px.colors.qualitative.Dark24
        + px.colors.qualitative.Light24
        + px.colors.qualitative.Bold
    )

    carreras = sorted(
        df["Carrera"]
        .dropna()
        .astype(str)
        .unique()
    )

    return {
        carrera: paleta[indice % len(paleta)]
        for indice, carrera in enumerate(carreras)
    }


def hist_mostrar_grafica_calificaciones(df):
    tabla = hist_crear_tabla_calificaciones_por_sexo(df)

    if tabla.empty:
        st.info("No hay promedios válidos entre 60 y 100.")
        return

    fig = px.bar(
        tabla,
        x="Porcentaje",
        y="Sexo_normalizado",
        color="Rango_promedio",
        orientation="h",
        barmode="stack",
        text="Etiqueta",
        custom_data=["Aspirantes"],
        category_orders={
            "Sexo_normalizado": [
                "Mujer",
                "Hombre",
                "Sin especificar"
            ],
            "Rango_promedio": [
                "60-69",
                "70-79",
                "80-89",
                "90-100"
            ]
        },
        color_discrete_map={
            "60-69": "#E74C3C",
            "70-79": "#F39C12",
            "80-89": "#F1C40F",
            "90-100": "#27AE60"
        }
    )

    fig.update_traces(
        textposition="inside",
        insidetextanchor="middle",
        hovertemplate=(
            "<b>Sexo:</b> %{y}<br>"
            "<b>Rango:</b> %{fullData.name}<br>"
            "<b>Aspirantes:</b> %{customdata[0]}<br>"
            "<b>Porcentaje:</b> %{x:.1f}%"
            "<extra></extra>"
        )
    )

    fig.update_layout(
        title="Promedio de bachillerato por sexo",
        legend_title_text="Rango de promedio",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.08,
            xanchor="center",
            x=0.5
        ),
        xaxis=dict(
            title="Porcentaje de aspirantes",
            range=[0, 100],
            ticksuffix="%"
        ),
        yaxis_title="",
        height=420,
        margin=dict(t=100, b=40, l=30, r=30)
    )

    st.plotly_chart(fig, use_container_width=True)


def hist_mostrar_grafica_semaforo_bachillerato(df):
    tabla = hist_crear_distribucion_calificaciones_bachillerato(
        df,
        top_n=10
    )

    if tabla.empty:
        st.info(
            "No hay suficientes promedios válidos para relacionar "
            "con bachilleratos."
        )
        return

    fig = px.bar(
        tabla,
        x="Porcentaje",
        y="Escuela_etiqueta",
        color="Rango_promedio",
        orientation="h",
        barmode="stack",
        text="Etiqueta",
        custom_data=["Aspirantes", "Total"],
        category_orders={
            "Rango_promedio": [
                "60-69",
                "70-79",
                "80-89",
                "90-100"
            ]
        },
        color_discrete_map={
            "60-69": "#E74C3C",
            "70-79": "#F39C12",
            "80-89": "#F1C40F",
            "90-100": "#27AE60"
        }
    )

    fig.update_traces(
        textposition="inside",
        insidetextanchor="middle",
        hovertemplate=(
            "<b>Bachillerato:</b> %{y}<br>"
            "<b>Rango:</b> %{fullData.name}<br>"
            "<b>Aspirantes:</b> %{customdata[0]} de %{customdata[1]}<br>"
            "<b>Porcentaje:</b> %{x:.1f}%"
            "<extra></extra>"
        )
    )

    fig.update_layout(
        title="Distribución de calificaciones por bachillerato",
        legend_title_text="Semáforo",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.08,
            xanchor="center",
            x=0.5
        ),
        xaxis=dict(
            title="Porcentaje de aspirantes",
            range=[0, 100],
            ticksuffix="%"
        ),
        yaxis_title="",
        height=720,
        margin=dict(t=100, b=40, l=320, r=30)
    )

    st.plotly_chart(fig, use_container_width=True)


def hist_mostrar_concentrado_estados(df, max_estados=5):
    resumen = (
        df[
            df["Estado_procedencia"] != "Sin dato"
        ]
        .groupby("Estado_procedencia")
        .size()
        .reset_index(name="Aspirantes")
        .sort_values("Aspirantes", ascending=False)
    )

    if resumen.empty:
        st.info("No hay información de estado de procedencia.")
        return

    total = resumen["Aspirantes"].sum()

    if len(resumen) > max_estados:

        principales = resumen.head(max_estados - 1).copy()
        otros = resumen.iloc[max_estados - 1:]["Aspirantes"].sum()

        fila_otros = pd.DataFrame({
            "Estado_procedencia": ["Otros estados"],
            "Aspirantes": [otros]
        })

        resumen = pd.concat(
            [principales, fila_otros],
            ignore_index=True
        )

    resumen["Porcentaje"] = (
        resumen["Aspirantes"]
        / total
        * 100
    ).round(1)

    columnas = st.columns(len(resumen))

    for columna, (_, fila) in zip(columnas, resumen.iterrows()):

        columna.metric(
            fila["Estado_procedencia"],
            f"{int(fila['Aspirantes']):,}",
            f"{fila['Porcentaje']:.1f}%",
            delta_color="off"
        )


def hist_mostrar_sunburst_udec(df, mapa_colores_carreras):
    df_udec = df[
        df["Bachillerato_procedencia"]
        == "Universidad de Colima (U de C)"
    ].copy()

    if df_udec.empty:
        st.info("No se encontraron aspirantes provenientes de la U de C.")
        return

    resumen = (
        df_udec
        .groupby("Carrera")
        .size()
        .reset_index(name="Aspirantes")
        .sort_values("Aspirantes", ascending=False)
    )

    total = int(resumen["Aspirantes"].sum())

    labels = ["Universidad de Colima (U de C)"]
    parents = [""]
    values = [total]
    ids = ["root_udec"]
    colores = ["#595959"]
    textos = [f"U de C<br>n={total}"]

    for indice, fila in resumen.reset_index(drop=True).iterrows():

        carrera = fila["Carrera"]
        aspirantes = int(fila["Aspirantes"])
        porcentaje = aspirantes / total * 100

        labels.append(carrera)
        parents.append("root_udec")
        values.append(aspirantes)
        ids.append(f"udec_carrera_{indice}")

        colores.append(
            mapa_colores_carreras.get(carrera, "#9E9E9E")
        )

        textos.append(
            f"{carrera}<br>{porcentaje:.1f}%"
        )

    fig = go.Figure(
        go.Sunburst(
            ids=ids,
            labels=labels,
            parents=parents,
            values=values,
            text=textos,
            textinfo="text",
            branchvalues="total",
            marker=dict(
                colors=colores,
                line=dict(color="#111217", width=1)
            ),
            insidetextorientation="radial",
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Aspirantes: %{value}<br>"
                "Participación: %{percentParent:.1%}"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(
        title="Universidad de Colima → carreras elegidas",
        height=560,
        margin=dict(t=70, b=20, l=20, r=20)
    )

    st.plotly_chart(fig, use_container_width=True)


def hist_mostrar_sunburst_otros_bachilleratos(
    df,
    mapa_colores_carreras,
    top_n=10
):
    df_otros = df[
        (
            df["Bachillerato_procedencia"]
            != "Universidad de Colima (U de C)"
        )
        &
        (
            df["Bachillerato_procedencia"] != "Sin dato"
        )
    ].copy()

    if df_otros.empty:
        st.info("No hay registros de otros bachilleratos.")
        return

    totales = (
        df_otros
        .groupby("Bachillerato_procedencia")
        .size()
        .reset_index(name="Aspirantes")
        .sort_values("Aspirantes", ascending=False)
    )

    escuelas_top = totales.head(top_n)[
        "Bachillerato_procedencia"
    ].tolist()

    df_otros["Escuela_sunburst"] = np.where(
        df_otros["Bachillerato_procedencia"].isin(escuelas_top),
        df_otros["Bachillerato_procedencia"],
        "Otros bachilleratos"
    )

    resumen_escuelas = (
        df_otros
        .groupby("Escuela_sunburst")
        .size()
        .reset_index(name="Aspirantes")
        .sort_values("Aspirantes", ascending=False)
    )

    resumen_carreras = (
        df_otros
        .groupby(["Escuela_sunburst", "Carrera"])
        .size()
        .reset_index(name="Aspirantes")
    )

    total_externos = int(resumen_escuelas["Aspirantes"].sum())

    tonos_gris = [
        "#424242",
        "#505050",
        "#5E5E5E",
        "#6C6C6C",
        "#7A7A7A",
        "#888888",
        "#969696",
        "#A4A4A4",
        "#B2B2B2",
        "#C0C0C0",
        "#6A6A6A"
    ]

    labels = ["Otros bachilleratos"]
    parents = [""]
    values = [total_externos]
    ids = ["root_otros"]
    colores = ["#303030"]
    textos = [f"Otros<br>n={total_externos}"]

    ids_escuelas = {}
    totales_escuelas = {}

    for indice, fila in resumen_escuelas.reset_index(drop=True).iterrows():

        escuela = fila["Escuela_sunburst"]
        aspirantes = int(fila["Aspirantes"])
        porcentaje = aspirantes / total_externos * 100
        id_escuela = f"escuela_{indice}"

        ids_escuelas[escuela] = id_escuela
        totales_escuelas[escuela] = aspirantes

        labels.append(escuela)
        parents.append("root_otros")
        values.append(aspirantes)
        ids.append(id_escuela)

        colores.append(
            tonos_gris[indice % len(tonos_gris)]
        )

        textos.append(
            f"{escuela}<br>{porcentaje:.1f}%"
        )

    for indice, fila in resumen_carreras.reset_index(drop=True).iterrows():

        carrera = fila["Carrera"]
        escuela = fila["Escuela_sunburst"]
        aspirantes = int(fila["Aspirantes"])

        total_escuela = totales_escuelas[escuela]
        porcentaje = aspirantes / total_escuela * 100

        labels.append(carrera)
        parents.append(ids_escuelas[escuela])
        values.append(aspirantes)
        ids.append(f"carrera_{indice}")

        colores.append(
            mapa_colores_carreras.get(carrera, "#9E9E9E")
        )

        textos.append(
            f"{carrera}<br>{porcentaje:.1f}%"
        )

    fig = go.Figure(
        go.Sunburst(
            ids=ids,
            labels=labels,
            parents=parents,
            values=values,
            text=textos,
            textinfo="text",
            branchvalues="total",
            marker=dict(
                colors=colores,
                line=dict(color="#111217", width=1)
            ),
            insidetextorientation="radial",
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Aspirantes: %{value}<br>"
                "Participación: %{percentParent:.1%}"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(
        title="Otros bachilleratos → carreras elegidas",
        height=560,
        margin=dict(t=70, b=20, l=20, r=20)
    )

    st.plotly_chart(fig, use_container_width=True)


def render_historial():
    st.title("🎓 Historial de Aspirantes")
    st.caption(
        "Análisis general y análisis por carrera de aspirantes de nuevo ingreso."
    )

    st.sidebar.markdown("---")
    st.sidebar.subheader("Carga Historial")

    archivo_subido = st.sidebar.file_uploader(
        "Carga el archivo Excel de aspirantes",
        type=["xlsx", "xls"],
        key="hist_archivo"
    )

    if archivo_subido is None:
        st.info("Carga un archivo Excel para iniciar el análisis.")
        st.stop()

    contenido_archivo = archivo_subido.getvalue()

    with st.spinner("Leyendo e integrando hojas del archivo..."):
        df_general, df_bitacora = hist_procesar_archivo_excel(
            contenido_archivo
        )

    if df_general.empty:
        st.error("No se pudieron identificar registros de aspirantes.")
        st.dataframe(df_bitacora, use_container_width=True)
        st.stop()

    columna_sexo = util_encontrar_columna(
        df_general,
        ["Género", "Genero", "Sexo"]
    )

    if columna_sexo is not None:
        df_general["Sexo_normalizado"] = df_general[columna_sexo].apply(
            hist_normalizar_sexo
        )
    else:
        df_general["Sexo_normalizado"] = "Sin especificar"

    df_general["Rango_promedio"] = df_general[
        "Promedio_normalizado_100"
    ].apply(hist_clasificar_rango_promedio)
    columna_escuela = util_encontrar_columna(
        df_general,
        [
            "Escuela de Procedencia",
            "Escuela Procedencia",
            "Procedencia",
            "Escuela"
        ]
    )
    
    if columna_escuela is not None:

        df_general["Bachillerato_procedencia_original"] = df_general[
            columna_escuela
        ].fillna("Sin dato").astype(str)

        df_general["Bachillerato_procedencia"] = df_general[
            columna_escuela
        ].apply(hist_normalizar_escuela_procedencia)

        df_general["Estado_procedencia"] = df_general[
            columna_escuela
        ].apply(hist_clasificar_estado_procedencia)
    else:
        df_general["Bachillerato_procedencia_original"] = "Sin dato"
        df_general["Bachillerato_procedencia"] = "Sin dato"
        df_general["Estado_procedencia"] = "Sin dato"

    mapa_colores_carreras = hist_crear_mapa_colores_carreras(df_general)

    seccion_activa = st.radio(
        "Navegación",
        [
            "📊 Análisis general",
            "🎓 Análisis por carrera"
        ],
        horizontal=True,
        label_visibility="collapsed",
        key="hist_navegacion_principal"
    )

    if seccion_activa == "📊 Análisis general":

        st.subheader("Análisis general de aspirantes")

        total = len(df_general)
        mujeres = df_general["Sexo_normalizado"].eq("Mujer").sum()
        hombres = df_general["Sexo_normalizado"].eq("Hombre").sum()
        sin_especificar = df_general["Sexo_normalizado"].eq(
            "Sin especificar"
        ).sum()

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Aspirantes", f"{total:,}")
        col2.metric("Mujeres", f"{mujeres:,}")
        col3.metric("Hombres", f"{hombres:,}")
        col4.metric("Sin especificar", f"{sin_especificar:,}")

        st.markdown("### Distribución de calificaciones por sexo")
        hist_mostrar_grafica_calificaciones(df_general)

        st.markdown("### Estado de procedencia")
        hist_mostrar_concentrado_estados(df_general)

        st.markdown(
            "## Distribución de calificaciones por bachillerato"
        )
        hist_mostrar_grafica_semaforo_bachillerato(df_general)

        st.markdown("## Origen académico y carrera elegida")

        col_udec, col_otros = st.columns(2)

        with col_udec:
            hist_mostrar_sunburst_udec(
                df_general,
                mapa_colores_carreras
            )

        with col_otros:
            hist_mostrar_sunburst_otros_bachilleratos(
                df_general,
                mapa_colores_carreras,
                top_n=10
            )

    elif seccion_activa == "🎓 Análisis por carrera":

        st.subheader("Análisis por carrera")

        carreras = sorted(
            df_general["Carrera"]
            .dropna()
            .astype(str)
            .unique()
        )

        carrera_seleccionada = st.selectbox(
            "Selecciona una carrera",
            options=carreras,
            key="hist_selector_carrera"
        )

        df_carrera = df_general[
            df_general["Carrera"] == carrera_seleccionada
        ].copy()

        st.markdown(f"## {carrera_seleccionada}")

        total_carrera = len(df_carrera)
        mujeres_carrera = df_carrera["Sexo_normalizado"].eq(
            "Mujer"
        ).sum()
        hombres_carrera = df_carrera["Sexo_normalizado"].eq(
            "Hombre"
        ).sum()
        sin_especificar_carrera = df_carrera["Sexo_normalizado"].eq(
            "Sin especificar"
        ).sum()

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Aspirantes", f"{total_carrera:,}")
        col2.metric("Mujeres", f"{mujeres_carrera:,}")
        col3.metric("Hombres", f"{hombres_carrera:,}")
        col4.metric(
            "Sin especificar",
            f"{sin_especificar_carrera:,}"
        )

        st.markdown("### Distribución de calificaciones por sexo")
        hist_mostrar_grafica_calificaciones(df_carrera)

        st.markdown("### Estado de procedencia")
        hist_mostrar_concentrado_estados(df_carrera)

        st.markdown(
            "## Distribución de calificaciones por bachillerato"
        )
        hist_mostrar_grafica_semaforo_bachillerato(df_carrera)

def chaside_color_semaforo(semaforo):
    """Define color visual del diagnóstico CHASIDE."""

    if semaforo == "Verde":
        return "#E8F5E9", "#2E7D32", "Perfil acorde"

    if semaforo == "Amarillo":
        return "#FFF8E1", "#F9A825", "Perfil por revisar"

    if semaforo == "Rojo":
        return "#FFEBEE", "#C62828", "Perfil en riesgo"

    if semaforo == "Respondió siempre igual":
        return "#F3F4F6", "#6B7280", "Respuesta no confiable"

    return "#F3F4F6", "#6B7280", "Sin sugerencia clara"


def chaside_recomendacion_ejecutiva_solo_link(fila):
    """Genera recomendación ejecutiva para CHASIDE."""

    semaforo = fila["Semáforo vocacional"]
    nivel = fila["Nivel de intensidad"]
    carrera_elegida = fila[CHASIDE_COLUMNA_CARRERA]
    carrera_mejor = fila["Carrera_Mejor_Perfilada"]

    if semaforo == "Respondió siempre igual":
        return (
            "El patrón de respuesta muestra baja variabilidad. "
            "Se recomienda interpretar el resultado con cautela y considerar "
            "una entrevista breve o reaplicación del instrumento."
        )

    if pd.isna(nivel) or str(nivel).strip() == "" or str(nivel) == "nan":
        nivel = "Sin nivel definido"

    if nivel == "Sin perfil":
        return (
            "El perfil vocacional muestra baja correspondencia con la carrera elegida. "
            "Se recomienda orientación vocacional individual y seguimiento temprano."
        )

    if nivel == "Perfil en riesgo":
        return (
            "Existe coincidencia mínima entre el perfil vocacional y la carrera elegida. "
            "Se recomienda seguimiento tutorial durante el primer semestre."
        )

    if nivel == "Perfil en transición":
        return (
            "El perfil muestra congruencia funcional con la carrera elegida. "
            "Se recomienda acompañamiento preventivo para consolidar su permanencia."
        )

    if nivel == "Joven promesa":
        return (
            "El perfil muestra alta congruencia con la carrera elegida. "
            "Se recomienda promover actividades de alto desempeño y seguimiento positivo."
        )

    if carrera_mejor != carrera_elegida and carrera_mejor != "Sin sugerencia clara":
        return (
            f"El perfil sugiere mayor afinidad hacia {carrera_mejor}. "
            "Se recomienda revisión vocacional complementaria."
        )

    return (
        "El resultado puede considerarse regular. "
        "Se recomienda seguimiento académico preventivo durante el primer semestre."
    )


def chaside_render_reporte_ejecutivo_solo_link():
    """Reporte ejecutivo CHASIDE sin archivo de aspirantes."""

    st.markdown("## Reporte ejecutivo CHASIDE")

    url_chaside = st.text_input(
        "Pega el enlace de respuestas de Google Forms / Google Sheets",
        value="https://docs.google.com/spreadsheets/d/1YHMEb5hftOZfV-CMWoUsUgJh1xmsgTY3YYwAtq1dGQA/edit?resourcekey=&gid=1491376423#gid=1491376423",
        key="chaside_url_google_solo_link"
    )

    peso_intereses = st.slider(
        "Peso de intereses",
        min_value=0.0,
        max_value=1.0,
        value=0.8,
        step=0.1,
        key="chaside_peso_intereses_solo_link"
    )

    peso_aptitudes = round(1 - peso_intereses, 2)

    st.caption(
        f"Pesos activos → Intereses: {peso_intereses:.1f} | "
        f"Aptitudes: {peso_aptitudes:.1f}"
    )

    try:
        with st.spinner("Cargando y procesando respuestas CHASIDE..."):
            df_chaside_raw = chaside_cargar_respuestas(url_chaside)

            resultado_chaside = chaside_procesar_respuestas(
                df_chaside_raw,
                peso_intereses=peso_intereses,
                peso_aptitudes=peso_aptitudes
            )

        if isinstance(resultado_chaside, tuple):
            df_chaside = resultado_chaside[0]
        else:
            df_chaside = resultado_chaside

        if not isinstance(df_chaside, pd.DataFrame):
            raise ValueError(
                "La función chaside_procesar_respuestas no regresó un DataFrame válido."
            )

    except Exception as error:
        st.error(f"No fue posible procesar CHASIDE: {error}")
        return

    total_chaside = int(df_chaside.shape[0])

    st.success(
        f"CHASIDE procesado correctamente: {total_chaside} respuestas."
    )

    col1, col2, col3 = st.columns(3)

    col1.metric("Respuestas", total_chaside)

    col2.metric(
        "Perfil adecuado",
        int((df_chaside["Semáforo vocacional"] == "Verde").sum())
    )

    col3.metric(
        "Requiere revisión",
        int((df_chaside["Semáforo vocacional"] != "Verde").sum())
    )

    st.markdown("## Reporte individual para docente")

    df_selector = df_chaside.copy()

    df_selector["Etiqueta"] = df_selector.apply(
        lambda fila: (
            f"{fila[CHASIDE_COLUMNA_NOMBRE]} · "
            f"{fila[CHASIDE_COLUMNA_CARRERA]} · "
            f"{fila['Semáforo vocacional']}"
        ),
        axis=1
    )

    carreras_disponibles = sorted(
        df_selector[CHASIDE_COLUMNA_CARRERA]
        .dropna()
        .astype(str)
        .unique()
    )

    if not carreras_disponibles:
        st.warning("No hay carreras disponibles para mostrar.")
        return

    carrera_seleccionada = st.selectbox(
        "Selecciona carrera",
        options=carreras_disponibles,
        key="chaside_selector_carrera_solo_link"
    )

    df_selector_carrera = df_selector[
        df_selector[CHASIDE_COLUMNA_CARRERA].astype(str) == carrera_seleccionada
    ].copy()

    opciones = sorted(
        df_selector_carrera["Etiqueta"]
        .dropna()
        .astype(str)
        .unique()
    )

    if not opciones:
        st.warning("No hay estudiantes disponibles para esta carrera.")
        return

    seleccion = st.selectbox(
        "Selecciona estudiante",
        options=opciones,
        key="chaside_selector_estudiante_solo_link"
    )

    fila = df_selector_carrera[
        df_selector_carrera["Etiqueta"] == seleccion
    ].iloc[0]
    
    fondo, borde, titulo = chaside_color_semaforo(
        fila["Semáforo vocacional"]
    )

    recomendacion = chaside_recomendacion_ejecutiva_solo_link(fila)

    st.markdown(
        f"""
        <div style="
            background-color:{fondo};
            border-left:10px solid {borde};
            padding:20px 24px;
            border-radius:16px;
            color:#111111;
            margin-bottom:16px;
        ">
            <h3 style="margin-top:0; color:#111111;">{titulo}</h3>
            <p style="margin-bottom:0;">
                <b>Diagnóstico:</b> {fila['Diagnóstico vocacional']}
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown("### Identificación")
        st.write(f"**Nombre:** {fila[CHASIDE_COLUMNA_NOMBRE]}")
        st.write(f"**Carrera elegida:** {fila[CHASIDE_COLUMNA_CARRERA]}")

        if CHASIDE_COLUMNA_EMAIL_1 in fila.index:
            st.write(f"**Correo Google:** {fila[CHASIDE_COLUMNA_EMAIL_1]}")

        if CHASIDE_COLUMNA_EMAIL_2 in fila.index:
            st.write(f"**Correo escrito:** {fila[CHASIDE_COLUMNA_EMAIL_2]}")

    with col_b:
        st.markdown("### Perfil vocacional")
        st.write(
            f"**Área fuerte:** {fila['Area_Fuerte_Ponderada']} · "
            f"{CHASIDE_AREAS_LONG.get(fila['Area_Fuerte_Ponderada'], '')}"
        )
        st.write(f"**Semáforo vocacional:** {fila['Semáforo vocacional']}")
        st.write(f"**Nivel de intensidad:** {fila['Nivel de intensidad']}")
        st.write(f"**Carrera mejor perfilada:** {fila['Carrera_Mejor_Perfilada']}")

    st.markdown("### Recomendación docente")
    st.info(recomendacion)

    st.markdown("### Puntajes por área CHASIDE")

    tabla_areas = []

    for area in CHASIDE_AREAS:
        tabla_areas.append(
            {
                "Área": area,
                "Descripción": CHASIDE_AREAS_LONG[area],
                "Intereses": fila[f"INTERES_{area}"],
                "Aptitudes": fila[f"APTITUD_{area}"],
                "Puntaje combinado": fila[f"PUNTAJE_COMBINADO_{area}"]
            }
        )

    st.dataframe(
        pd.DataFrame(tabla_areas).sort_values(
            "Puntaje combinado",
            ascending=False
        ),
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        label="⬇️ Descargar respuestas CHASIDE procesadas",
        data=dataframe_a_excel_bytes(
            {
                "CHASIDE procesado": df_chaside
            }
        ),
        file_name="chaside_procesado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_chaside_procesado_solo_link"
    )


def render_modulo_chaside_con_carga():
    """Muestra CHASIDE únicamente desde el enlace de respuestas."""

    st.title("🧭 Diagnóstico vocacional CHASIDE")
    st.caption(
        "Este módulo se alimenta directamente del enlace de respuestas de Google Forms / Google Sheets."
    )

    chaside_render_reporte_ejecutivo_solo_link()
# ============================================================
# MÓDULO 3: PERFIL INDIVIDUAL
# ============================================================
    
def perfil_normalizar_nombre(valor):
    """Normaliza nombres para cruzar Historial contra EVALUATEC."""

    if pd.isna(valor):
        return ""

    texto = str(valor).upper().strip()
    texto = unicodedata.normalize("NFD", texto)

    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )

    texto = re.sub(r"[^A-ZÑ\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()

    return texto


def perfil_nombre_visible(valor):
    """Convierte el nombre a formato visible."""

    if pd.isna(valor):
        return "Sin nombre"

    texto = str(valor).strip()
    texto = re.sub(r"\s+", " ", texto)

    return texto.title()


def perfil_simplificar_carrera(valor):
    """Normaliza carreras para validación flexible."""

    if pd.isna(valor):
        return ""

    texto = util_limpiar_texto(valor)

    reemplazos = [
        "licenciatura en",
        "lic. en",
        "licenciatura",
        "lic ",
        "ingenieria en",
        "ingeniería en",
        "ing. en",
        "ing ",
        "carrera de",
        "programa de"
    ]

    for reemplazo in reemplazos:
        texto = texto.replace(reemplazo, "")

    texto = texto.replace(".", " ")
    texto = re.sub(r"[^a-z0-9\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()

    return texto


def perfil_encontrar_nombre_historial(df):
    """Detecta columnas separadas de apellido y nombre en Historial."""

    col_apellido_paterno = util_encontrar_columna(
        df,
        [
            "Apellido paterno",
            "Primer apellido",
            "Paterno"
        ]
    )

    col_apellido_materno = util_encontrar_columna(
        df,
        [
            "Apellido materno",
            "Segundo apellido",
            "Materno"
        ]
    )

    col_nombre = util_encontrar_columna(
        df,
        [
            "Nombre (s)",
            "Nombre(s)",
            "Nombres",
            "Nombre"
        ]
    )

    if col_apellido_paterno is None or col_nombre is None:
        return None, None, None

    return col_apellido_paterno, col_apellido_materno, col_nombre


def perfil_encontrar_nombre_evaluatec(df):
    """Detecta columna de nombre completo en EVALUATEC."""

    posibles_columnas = [
        "Nombre completo",
        "NombreCompleto",
        "Nombre del aspirante",
        "Aspirante",
        "Alumno",
        "Estudiante",
        "Participante",
        "Nombre",
        "Nombre(s)"
    ]

    return util_encontrar_columna(df, posibles_columnas)
def perfil_preparar_historial(contenido_archivo):
    """Procesa Historial y prepara nombre, carrera y datos generales."""

    df_historial, df_bitacora = hist_procesar_archivo_excel(
        contenido_archivo
    )

    if df_historial.empty:
        return pd.DataFrame(), df_bitacora

    col_apellido_paterno, col_apellido_materno, col_nombre = perfil_encontrar_nombre_historial(
        df_historial
    )

    if col_apellido_paterno is None or col_nombre is None:
        st.error(
            "En el Historial no pude identificar las columnas de apellido paterno y nombre."
        )
        return pd.DataFrame(), df_bitacora

    if col_apellido_materno is None:
        df_historial["Nombre_completo_visible"] = (
            df_historial[col_apellido_paterno].fillna("").astype(str)
            + " "
            + df_historial[col_nombre].fillna("").astype(str)
        )
    else:
        df_historial["Nombre_completo_visible"] = (
            df_historial[col_apellido_paterno].fillna("").astype(str)
            + " "
            + df_historial[col_apellido_materno].fillna("").astype(str)
            + " "
            + df_historial[col_nombre].fillna("").astype(str)
        )

    df_historial["Nombre_completo_visible"] = df_historial[
        "Nombre_completo_visible"
    ].apply(perfil_nombre_visible)

    df_historial["Nombre_match"] = df_historial[
        "Nombre_completo_visible"
    ].apply(perfil_normalizar_nombre)

    df_historial = df_historial[
        df_historial["Nombre_match"].notna()
        &
        (df_historial["Nombre_match"].str.strip() != "")
        &
        (~df_historial["Nombre_match"].str.contains("AULA", na=False))
    ].copy()

    df_historial["Letra_apellido"] = df_historial[
        col_apellido_paterno
    ].fillna("").astype(str).str.strip().str[:1].str.upper()

    df_historial["Carrera_match"] = df_historial[
        "Carrera"
    ].apply(perfil_simplificar_carrera)

    columna_id = util_encontrar_columna(
        df_historial,
        [
            "Matrícula/ID",
            "Matricula/ID",
            "Matrícula",
            "Matricula",
            "ID"
        ]
    )

    if columna_id is not None:
        df_historial["ID_aspirante"] = df_historial[columna_id]
    else:
        df_historial["ID_aspirante"] = "Sin dato"

    columna_sexo = util_encontrar_columna(
        df_historial,
        ["Género", "Genero", "Sexo"]
    )

    if columna_sexo is not None:
        df_historial["Sexo_normalizado"] = df_historial[
            columna_sexo
        ].apply(hist_normalizar_sexo)
    else:
        df_historial["Sexo_normalizado"] = "Sin especificar"

    df_historial["Rango_promedio"] = df_historial[
        "Promedio_normalizado_100"
    ].apply(hist_clasificar_rango_promedio)

    columna_escuela = util_encontrar_columna(
        df_historial,
        [
            "Escuela de Procedencia",
            "Escuela Procedencia",
            "Procedencia",
            "Escuela"
        ]
    )

    if columna_escuela is not None:

        df_historial["Bachillerato_procedencia_original"] = df_historial[
            columna_escuela
        ].fillna("Sin dato").astype(str)

        df_historial["Bachillerato_procedencia"] = df_historial[
            columna_escuela
        ].apply(hist_normalizar_escuela_procedencia)

        df_historial["Estado_procedencia"] = df_historial[
            columna_escuela
        ].apply(hist_clasificar_estado_procedencia)

    else:
        df_historial["Bachillerato_procedencia_original"] = "Sin dato"
        df_historial["Bachillerato_procedencia"] = "Sin dato"
        df_historial["Estado_procedencia"] = "Sin dato"

    return df_historial, df_bitacora


def perfil_preparar_evaluatec(archivos_subidos):
    """Procesa los 3 CSV de EVALUATEC y prepara nombre normalizado."""

    bases = []
    errores = []

    for archivo in archivos_subidos:
        try:
            df_archivo, areas_detectadas = eval_procesar_archivo(
                archivo
            )

            columna_nombre = perfil_encontrar_nombre_evaluatec(
                df_archivo
            )

            if columna_nombre is None:
                errores.append(
                    f"{archivo.name}: no encontré columna de nombre del aspirante."
                )
                continue

            df_archivo["Nombre_completo_visible"] = df_archivo[
                columna_nombre
            ].apply(perfil_nombre_visible)

            df_archivo["Nombre_match"] = df_archivo[
                columna_nombre
            ].apply(perfil_normalizar_nombre)

            df_archivo["Carrera_match"] = df_archivo[
                "Carrera_normalizada"
            ].apply(perfil_simplificar_carrera)

            df_archivo["Areas_detectadas_lista"] = ",".join(
                areas_detectadas.keys()
            )

            bases.append(df_archivo)

        except Exception as error:
            errores.append(
                f"{archivo.name}: {error}"
            )

    if not bases:
        return pd.DataFrame(), errores

    df_evaluatec = pd.concat(
        bases,
        ignore_index=True,
        sort=False
    )

    return df_evaluatec, errores


def perfil_crear_base_cruzada(df_historial, df_evaluatec):
    """Cruza Historial y EVALUATEC usando nombre como llave principal."""

    hist = df_historial.copy()
    eval_df = df_evaluatec.copy()

    hist = hist.add_prefix("hist_")
    eval_df = eval_df.add_prefix("eval_")

    df_cruzado = hist.merge(
        eval_df,
        left_on="hist_Nombre_match",
        right_on="eval_Nombre_match",
        how="outer",
        indicator=True
    )

    df_cruzado["Nombre_visible"] = df_cruzado[
        "hist_Nombre_completo_visible"
    ].combine_first(
        df_cruzado["eval_Nombre_completo_visible"]
    )

    df_cruzado["Carrera_historial"] = df_cruzado[
        "hist_Carrera"
    ]

    df_cruzado["Carrera_evaluatec"] = df_cruzado[
        "eval_Carrera_normalizada"
    ]

    df_cruzado["Carrera_referencia"] = df_cruzado[
        "Carrera_historial"
    ].combine_first(
        df_cruzado["Carrera_evaluatec"]
    )

    df_cruzado["Letra_apellido"] = df_cruzado[
        "hist_Letra_apellido"
    ]

    df_cruzado["Letra_apellido"] = df_cruzado[
        "Letra_apellido"
    ].fillna(
        df_cruzado["Nombre_visible"]
        .fillna("")
        .astype(str)
        .str[:1]
        .str.upper()
    )

    df_cruzado["Carrera_coincide"] = (
        df_cruzado["hist_Carrera_match"]
        == df_cruzado["eval_Carrera_match"]
    )

    df_cruzado["Estatus_cruce"] = np.select(
        [
            df_cruzado["_merge"] == "both",
            df_cruzado["_merge"] == "left_only",
            df_cruzado["_merge"] == "right_only"
        ],
        [
            "Coincide en ambas bases",
            "Solo en Historial",
            "Solo en EVALUATEC"
        ],
        default="Sin clasificar"
    )

    df_cruzado.loc[
        (
            df_cruzado["_merge"] == "both"
        )
        &
        (
            df_cruzado["Carrera_coincide"] == False
        ),
        "Estatus_cruce"
    ] = "Coincide por nombre, carrera distinta"

    def crear_etiqueta_selector(fila):
        nombre = fila["Nombre_visible"]
        carrera = fila["Carrera_referencia"]
        estatus = fila["Estatus_cruce"]

        if pd.isna(nombre):
            nombre = "Sin nombre"

        if pd.isna(carrera):
            carrera = "Sin carrera"

        if estatus == "Coincide en ambas bases":
            return f"{nombre} · {carrera}"

        if estatus == "Coincide por nombre, carrera distinta":
            return f"{nombre} · {carrera} · Validar carrera"

        if estatus == "Solo en Historial":
            return f"{nombre} · {carrera} · Solo Historial"

        if estatus == "Solo en EVALUATEC":
            return f"{nombre} · {carrera} · Solo EVALUATEC"

        return f"{nombre} · {carrera}"

    df_cruzado["Etiqueta_selector"] = df_cruzado.apply(
        crear_etiqueta_selector,
        axis=1
    )

    return df_cruzado


def perfil_valor(fila, columna, default="Sin dato"):
    """Obtiene un valor seguro para mostrar."""

    if columna not in fila.index:
        return default

    valor = fila[columna]

    if pd.isna(valor) or str(valor).strip() == "":
        return default

    return valor


def perfil_formato_porcentaje(valor):
    """Formatea valores como porcentaje."""

    if pd.isna(valor):
        return "Sin dato"

    try:
        return f"{float(valor):.1f}%"
    except Exception:
        return "Sin dato"


def perfil_obtener_areas_individuales(fila):
    """Extrae resultados EVALUATEC por dimensión para un aspirante."""

    registros = []

    for codigo in EVAL_ORDEN_AREAS:
        columna = f"eval_Area_{codigo}"

        if columna not in fila.index:
            continue

        valor = fila[columna]

        if pd.isna(valor):
            continue

        registros.append(
            {
                "Código": codigo,
                "Dimensión": EVAL_ETIQUETAS_AREAS.get(codigo, codigo),
                "Resultado": float(valor)
            }
        )

    return pd.DataFrame(registros)


def perfil_clasificar_nivelacion(promedio):
    """Clasifica necesidad de nivelación individual."""

    if pd.isna(promedio):
        return "Sin dato"

    if promedio < 50:
        return "Nivelación prioritaria"

    if promedio < 70:
        return "Seguimiento académico"

    return "Fortalecimiento regular"


def perfil_mostrar_resultados_evaluatec_individual(tabla_areas):
    """Muestra gráfica individual de áreas EVALUATEC."""

    if tabla_areas.empty:
        st.info("No hay resultados EVALUATEC para este aspirante.")
        return

    tabla = tabla_areas.sort_values(
        "Resultado",
        ascending=True
    )

    colores = []

    for valor in tabla["Resultado"]:
        if valor < 25:
            colores.append(EVAL_COLORES_NIVELES["Bajo"])
        elif valor < 50:
            colores.append(EVAL_COLORES_NIVELES["Básico"])
        elif valor < 75:
            colores.append(EVAL_COLORES_NIVELES["Satisfactorio"])
        else:
            colores.append(EVAL_COLORES_NIVELES["Alto"])

    fig = go.Figure(
        go.Bar(
            x=tabla["Resultado"],
            y=tabla["Dimensión"],
            orientation="h",
            marker_color=colores,
            text=[
                f"{valor:.1f}%"
                for valor in tabla["Resultado"]
            ],
            textposition="outside",
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Resultado: %{x:.1f}%"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(
        title="Resultados EVALUATEC por dimensión",
        xaxis=dict(
            title="Calificación obtenida",
            range=[0, 100],
            ticksuffix="%"
        ),
        yaxis_title="",
        height=max(420, len(tabla) * 70),
        margin=dict(t=70, b=40, l=230, r=60)
    )

    st.plotly_chart(fig, use_container_width=True)


def perfil_clasificar_por_cuartiles(valor, serie_referencia):
    """
    Clasifica al estudiante contra su grupo de referencia.
    Usa lógica de boxplot:
    - Atípico alto: joven promesa
    - Atípico bajo: alto riesgo
    - Dentro de cuartiles: semáforo de acompañamiento
    """

    serie = pd.Series(serie_referencia).dropna()

    if pd.isna(valor) or serie.empty or len(serie) < 5:
        return {
            "Clasificación": "Sin referencia suficiente",
            "Nivel": "Sin dato",
            "Q1": np.nan,
            "Q2": np.nan,
            "Q3": np.nan,
            "LI": np.nan,
            "LS": np.nan
        }

    q1 = serie.quantile(0.25)
    q2 = serie.quantile(0.50)
    q3 = serie.quantile(0.75)
    iqr = q3 - q1

    limite_inferior = q1 - 1.5 * iqr
    limite_superior = q3 + 1.5 * iqr

    if valor < limite_inferior:
        clasificacion = "Perfil en alto riesgo de no acreditar el semestre"
        nivel = "Atípico bajo"

    elif valor > limite_superior:
        clasificacion = "Joven promesa"
        nivel = "Atípico alto"

    elif valor < q1:
        clasificacion = "Riesgo preventivo"
        nivel = "Debajo de Q1"

    elif valor < q2:
        clasificacion = "Seguimiento académico"
        nivel = "Entre Q1 y Q2"

    elif valor < q3:
        clasificacion = "Desempeño esperado"
        nivel = "Entre Q2 y Q3"

    else:
        clasificacion = "Fortaleza académica"
        nivel = "Arriba de Q3"

    return {
        "Clasificación": clasificacion,
        "Nivel": nivel,
        "Q1": q1,
        "Q2": q2,
        "Q3": q3,
        "LI": limite_inferior,
        "LS": limite_superior
    }


def perfil_obtener_grupo_referencia(fila, df_evaluatec):
    """
    Usa como grupo de referencia la carrera EVALUATEC del estudiante.
    Si no hay carrera EVALUATEC, usa todos los registros EVALUATEC.
    """

    carrera_eval = perfil_valor(
        fila,
        "Carrera_evaluatec",
        ""
    )

    if carrera_eval != "":
        grupo = df_evaluatec[
            df_evaluatec["Carrera_normalizada"] == carrera_eval
        ].copy()

        if not grupo.empty:
            return grupo

    return df_evaluatec.copy()

def perfil_clasificar_area_boxplot(resultado, columna_grupo, grupo_referencia):
    """
    Clasifica una dimensión contra el grupo de referencia usando lógica de boxplot.
    """

    if columna_grupo not in grupo_referencia.columns:
        return "Estudiante regular", "#111111"

    serie = grupo_referencia[columna_grupo].dropna()

    if pd.isna(resultado) or serie.empty or len(serie) < 5:
        return "Estudiante regular", "#111111"

    q1 = serie.quantile(0.25)
    q3 = serie.quantile(0.75)
    iqr = q3 - q1

    limite_inferior = q1 - 1.5 * iqr
    limite_superior = q3 + 1.5 * iqr

    if resultado < limite_inferior:
        return "Estudiante en riesgo de no acreditar", "#C62828"

    if resultado > limite_superior:
        return "Joven promesa", "#2E7D32"

    return "Estudiante regular", "#111111"


def perfil_texto_chaside_coloreado(texto, tipo="positivo"):
    """Devuelve texto HTML coloreado para el dictamen CHASIDE."""

    if tipo == "positivo":
        color = "#2E7D32"

    elif tipo == "negativo":
        color = "#C62828"

    else:
        color = "#111111"

    return f"<b style='color:{color};'>{texto}</b>"

def perfil_generar_pdf_dictamen(nombre, carrera, configuracion, dictamen_tutoria):
    """Genera PDF del dictamen tutorial."""

    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib import colors

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=1.8 * cm,
        leftMargin=1.8 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm
    )

    styles = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "TituloDictamen",
        parent=styles["Title"],
        fontSize=18,
        leading=22,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#111111"),
        spaceAfter=14
    )

    estilo_info = ParagraphStyle(
        "InfoDictamen",
        parent=styles["Normal"],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#444444"),
        spaceAfter=16
    )

    estilo_punto = ParagraphStyle(
        "PuntoDictamen",
        parent=styles["Normal"],
        fontSize=10.5,
        leading=16,
        textColor=colors.HexColor("#111111"),
        spaceAfter=12
    )

    elementos = []

    elementos.append(
        Paragraph("Dictamen tutorial", estilo_titulo)
    )

    elementos.append(
        Paragraph(
            f"<b>Estudiante:</b> {nombre}<br/>"
            f"<b>Carrera:</b> {carrera}<br/>"
            f"<b>Clasificación:</b> {configuracion['titulo']}",
            estilo_info
        )
    )

    for numero, (titulo, contenido) in enumerate(dictamen_tutoria, start=1):
        contenido_pdf = str(contenido)

        contenido_pdf = contenido_pdf.replace("<span", "<font")
        contenido_pdf = contenido_pdf.replace("</span>", "</font>")

        elementos.append(
            Paragraph(
                f"<b>{numero}. {titulo}:</b> {contenido_pdf}",
                estilo_punto
            )
        )

        elementos.append(Spacer(1, 8))

    doc.build(elementos)

    buffer.seek(0)
    return buffer.getvalue()
    
def render_perfil_individual():
    st.title("👤 Perfil individual del aspirante")
    
    serie = grupo_referencia[columna_grupo].dropna()

    if pd.isna(resultado) or serie.empty or len(serie) < 5:
        return "Estudiante regular", "#111111"

    q1 = serie.quantile(0.25)
    q3 = serie.quantile(0.75)
    iqr = q3 - q1

    limite_inferior = q1 - 1.5 * iqr
    limite_superior = q3 + 1.5 * iqr

    if resultado < limite_inferior:
        return "Estudiante en riesgo de no acreditar", "#C62828"

    if resultado > limite_superior:
        return "Joven promesa", "#2E7D32"

    return "Estudiante regular", "#111111"
    """
    Calcula posición del estudiante contra sus compañeros.
    No grafica el boxplot; solo usa su lógica para clasificar.
    """

    grupo = perfil_obtener_grupo_referencia(
        fila,
        df_evaluatec
    )

    promedio_eval = perfil_valor(
        fila,
        "eval_Promedio_global_individual",
        np.nan
    )

    resultado_global = perfil_clasificar_por_cuartiles(
        promedio_eval,
        grupo["Promedio_global_individual"]
    )

    registros_dimensiones = []

    for _, area in tabla_areas.iterrows():
        codigo = area["Código"]
        columna = f"Area_{codigo}"

        if columna not in grupo.columns:
            continue

        resultado_dimension = perfil_clasificar_por_cuartiles(
            area["Resultado"],
            grupo[columna]
        )

        registros_dimensiones.append(
            {
                "Dimensión": area["Dimensión"],
                "Resultado": area["Resultado"],
                "Clasificación": resultado_dimension["Clasificación"],
                "Nivel": resultado_dimension["Nivel"],
                "Q1": resultado_dimension["Q1"],
                "Q2": resultado_dimension["Q2"],
                "Q3": resultado_dimension["Q3"]
            }
        )

    tabla_contexto = pd.DataFrame(registros_dimensiones)

    return resultado_global, tabla_contexto, grupo


def perfil_color_clasificacion(clasificacion):
    """Color visual para semáforo."""

    if clasificacion == "Perfil en alto riesgo de no acreditar el semestre":
        return "error"

    if clasificacion in ["Riesgo preventivo", "Seguimiento académico"]:
        return "warning"

    if clasificacion in ["Desempeño esperado", "Fortaleza académica"]:
        return "success"

    if clasificacion == "Joven promesa":
        return "success"

    return "info"


def perfil_mostrar_contexto_academico(resultado_global, tabla_contexto, grupo):
    """Muestra el semáforo contextual del estudiante."""

    st.markdown("## Posición académica respecto a sus compañeros")

    st.caption(
        "La clasificación se calcula internamente con lógica de boxplot y cuartiles. "
        "No compara contra una calificación ideal, sino contra el desempeño de sus compañeros de referencia."
    )

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Grupo de referencia",
        f"{len(grupo):,} aspirantes"
    )

    col2.metric(
        "Q1",
        perfil_formato_porcentaje(resultado_global["Q1"])
    )

    col3.metric(
        "Mediana",
        perfil_formato_porcentaje(resultado_global["Q2"])
    )

    col4.metric(
        "Q3",
        perfil_formato_porcentaje(resultado_global["Q3"])
    )

    clasificacion = resultado_global["Clasificación"]
    tipo = perfil_color_clasificacion(clasificacion)

    mensaje = (
        f"**Clasificación global:** {clasificacion}. "
        f"Ubicación: {resultado_global['Nivel']}."
    )

    if tipo == "error":
        st.error(mensaje)
    elif tipo == "warning":
        st.warning(mensaje)
    elif tipo == "success":
        st.success(mensaje)
    else:
        st.info(mensaje)

    if not tabla_contexto.empty:
        st.markdown("### Semáforo por dimensión")

        tabla_vista = tabla_contexto[
            [
                "Dimensión",
                "Resultado",
                "Clasificación",
                "Nivel"
            ]
        ].copy()

        tabla_vista["Resultado"] = tabla_vista["Resultado"].apply(
            lambda valor: f"{valor:.1f}%"
        )

        st.dataframe(
            tabla_vista,
            use_container_width=True,
            hide_index=True
        )


def perfil_generar_acciones_acompanamiento(resultado_global, tabla_contexto):
    """
    Genera acciones de prevención y corrección según el lugar del estudiante
    frente a sus compañeros.
    """

    clasificacion_global = resultado_global["Clasificación"]

    acciones_prevencion = []
    acciones_correccion = []

    if tabla_contexto.empty:
        return {
            "Prevención": [
                "Validar manualmente el expediente porque no se encontraron resultados suficientes para comparar."
            ],
            "Corrección": [
                "Revisar captura de nombre, carrera y resultados EVALUATEC."
            ]
        }

    dimensiones_riesgo = tabla_contexto[
        tabla_contexto["Clasificación"].isin(
            [
                "Perfil en alto riesgo de no acreditar el semestre",
                "Riesgo preventivo",
                "Seguimiento académico"
            ]
        )
    ].sort_values("Resultado")

    dimensiones_fuertes = tabla_contexto[
        tabla_contexto["Clasificación"].isin(
            [
                "Fortaleza académica",
                "Joven promesa"
            ]
        )
    ].sort_values("Resultado", ascending=False)

    if clasificacion_global == "Perfil en alto riesgo de no acreditar el semestre":
        acciones_prevencion.append(
            "Canalizar a seguimiento académico temprano durante las primeras semanas del semestre."
        )

        acciones_correccion.append(
            "Diseñar un plan de regularización con metas semanales y revisión de avances."
        )

    elif clasificacion_global in ["Riesgo preventivo", "Seguimiento académico"]:
        acciones_prevencion.append(
            "Mantener monitoreo preventivo para evitar rezago durante el arranque del semestre."
        )

        acciones_correccion.append(
            "Asignar actividades de reforzamiento en las dimensiones con menor resultado."
        )

    elif clasificacion_global == "Joven promesa":
        acciones_prevencion.append(
            "Ofrecer actividades de mayor reto académico para mantener motivación y desarrollo."
        )

        acciones_correccion.append(
            "No requiere corrección inmediata; se recomienda seguimiento de consolidación."
        )

    else:
        acciones_prevencion.append(
            "Mantener acompañamiento regular y observación del desempeño en las primeras evaluaciones."
        )

        acciones_correccion.append(
            "Reforzar únicamente las dimensiones ubicadas por debajo de la mediana del grupo."
        )

    if not dimensiones_riesgo.empty:
        principales = dimensiones_riesgo.head(2)

        texto_dimensiones = ", ".join(
            [
                f"{fila['Dimensión']} ({fila['Resultado']:.1f}%)"
                for _, fila in principales.iterrows()
            ]
        )

        acciones_correccion.append(
            f"Priorizar trabajo académico en: {texto_dimensiones}."
        )

    if not dimensiones_fuertes.empty:
        principales_fuertes = dimensiones_fuertes.head(2)

        texto_fuertes = ", ".join(
            [
                f"{fila['Dimensión']} ({fila['Resultado']:.1f}%)"
                for _, fila in principales_fuertes.iterrows()
            ]
        )

        acciones_prevencion.append(
            f"Aprovechar como fortalezas de apoyo: {texto_fuertes}."
        )

    return {
        "Prevención": acciones_prevencion,
        "Corrección": acciones_correccion
    }

def perfil_clasificar_alerta_tutoria(
    promedio_bach,
    promedio_eval,
    resultado_global,
    tabla_contexto
):
    """
    Clasifica al estudiante en un semáforo tutorial.
    """

    clasificacion_global = resultado_global["Clasificación"]

    dimensiones_riesgo = 0

    if not tabla_contexto.empty:
        dimensiones_riesgo = tabla_contexto[
            tabla_contexto["Clasificación"].isin(
                [
                    "Perfil en alto riesgo de no acreditar el semestre",
                    "Riesgo preventivo",
                    "Seguimiento académico"
                ]
            )
        ].shape[0]

    if clasificacion_global == "Perfil en alto riesgo de no acreditar el semestre":
        return "red_flag"

    if pd.notna(promedio_eval) and promedio_eval < 45:
        return "red_flag"

    if dimensiones_riesgo >= 3:
        return "irregular"

    if clasificacion_global in ["Riesgo preventivo", "Seguimiento académico"]:
        return "observacion"

    if pd.notna(promedio_eval) and promedio_eval < 60:
        return "observacion"

    return "regular"


def perfil_configuracion_alerta_tutoria(nivel_alerta):
    """
    Define color, título y estilo visual del dictamen.
    """

    configuraciones = {
        "regular": {
            "titulo": "Alumno regular",
            "color_fondo": "#E8F5E9",
            "color_borde": "#2E7D32",
            "color_texto": "#111111"
        },
        "observacion": {
            "titulo": "Alumno en observación",
            "color_fondo": "#FFFDE7",
            "color_borde": "#FBC02D",
            "color_texto": "#111111"
        },
        "irregular": {
            "titulo": "Alumno irregular",
            "color_fondo": "#FFF3E0",
            "color_borde": "#F57C00",
            "color_texto": "#111111"
        },
        "red_flag": {
            "titulo": "Red flag académico",
            "color_fondo": "#FFEBEE",
            "color_borde": "#C62828",
            "color_texto": "#111111"
        }
    }

    return configuraciones.get(
        nivel_alerta,
        configuraciones["observacion"]
    )


def perfil_promedio_grupo_historial(fila, df_historial):
    """
    Obtiene promedio de bachillerato del grupo de referencia en historial.
    Usa la carrera del historial cuando está disponible.
    """

    carrera_historial = perfil_valor(
        fila,
        "Carrera_historial",
        ""
    )

    if carrera_historial != "":
        grupo_historial = df_historial[
            df_historial["Carrera"] == carrera_historial
        ].copy()

        if not grupo_historial.empty:
            return grupo_historial[
                "Promedio_normalizado_100"
            ].mean()

    return df_historial[
        "Promedio_normalizado_100"
    ].mean()

def perfil_texto_comparativo(valor_estudiante, valor_grupo):
    """
    Redacta comparación simple entre estudiante y grupo.
    """

    if pd.isna(valor_estudiante) or pd.isna(valor_grupo):
        return "sin información suficiente para realizar una comparación con sus pares"

    diferencia = valor_estudiante - valor_grupo

    if diferencia > 0:
        return (
            f"superior en {abs(diferencia):.1f} puntos porcentuales "
            f"respecto a sus pares"
        )

    if diferencia < 0:
        return (
            f"inferior en {abs(diferencia):.1f} puntos porcentuales "
            f"respecto a sus pares"
        )

    return "similar al promedio de sus pares"


def perfil_texto_comparativo_coloreado(valor_estudiante, valor_grupo):
    """
    Redacta comparación contra pares con color:
    rojo si está por debajo, verde si está por arriba.
    """

    if pd.isna(valor_estudiante) or pd.isna(valor_grupo):
        return "sin información suficiente para realizar una comparación con sus pares"

    diferencia = valor_estudiante - valor_grupo

    if diferencia > 0:
        return (
            f"<b style='color:#2E7D32;'>superior en {abs(diferencia):.1f} "
            f"puntos porcentuales respecto a sus pares</b>"
        )

    if diferencia < 0:
        return (
            f"<b style='color:#C62828;'>inferior en {abs(diferencia):.1f} "
            f"puntos porcentuales respecto a sus pares</b>"
        )

    return (
        "<b style='color:#333333;'>similar al promedio de sus pares</b>"
    )
    
def perfil_describir_dimensiones(tabla_contexto):
    """
    Identifica dimensiones prioritarias y fortalezas.
    """

    if tabla_contexto.empty:
        return {
            "prioritarias": "no se identificaron dimensiones prioritarias por falta de información",
            "fortalezas": "no se identificaron fortalezas por falta de información",
            "texto_prioritarias": "",
            "texto_fortalezas": ""
        }

    tabla_ordenada = tabla_contexto.sort_values(
        "Resultado",
        ascending=True
    )

    prioritarias = tabla_ordenada.head(2)
    fortalezas = tabla_ordenada.tail(2).sort_values(
        "Resultado",
        ascending=False
    )

    texto_prioritarias = ", ".join(
        [
            f"{fila['Dimensión']} ({fila['Resultado']:.1f}%)"
            for _, fila in prioritarias.iterrows()
        ]
    )

    texto_fortalezas = ", ".join(
        [
            f"{fila['Dimensión']} ({fila['Resultado']:.1f}%)"
            for _, fila in fortalezas.iterrows()
        ]
    )

    dimension_mas_baja = prioritarias.iloc[0]
    dimension_mas_fuerte = fortalezas.iloc[0]

    descripcion_prioritaria = (
        f"la dimensión con mayor necesidad de atención es "
        f"**{dimension_mas_baja['Dimensión']}**, con un resultado de "
        f"**{dimension_mas_baja['Resultado']:.1f}%**, ubicada como "
        f"**{dimension_mas_baja['Clasificación']}**"
    )

    descripcion_fortaleza = (
        f"mientras que la dimensión con mejor desempeño es "
        f"**{dimension_mas_fuerte['Dimensión']}**, con "
        f"**{dimension_mas_fuerte['Resultado']:.1f}%**, clasificada como "
        f"**{dimension_mas_fuerte['Clasificación']}**"
    )

    return {
        "prioritarias": texto_prioritarias,
        "fortalezas": texto_fortalezas,
        "texto_prioritarias": descripcion_prioritaria,
        "texto_fortalezas": descripcion_fortaleza
    }

def perfil_generar_dictamen_tutoria(
    fila,
    df_historial,
    grupo_referencia,
    resultado_global,
    tabla_contexto
):
    """
    Genera dictamen descriptivo para tutoría en formato ordenado.
    """

    nombre = perfil_valor(
        fila,
        "Nombre_visible"
    )

    escuela = perfil_valor(
        fila,
        "hist_Bachillerato_procedencia_original",
        perfil_valor(fila, "hist_Bachillerato_procedencia")
    )

    estado = perfil_valor(
        fila,
        "hist_Estado_procedencia"
    )

    promedio_bach = perfil_valor(
        fila,
        "hist_Promedio_normalizado_100",
        np.nan
    )

    promedio_eval = perfil_valor(
        fila,
        "eval_Promedio_global_individual",
        np.nan
    )

    promedio_bach_grupo = perfil_promedio_grupo_historial(
        fila,
        df_historial
    )

    promedio_eval_grupo = grupo_referencia[
        "Promedio_global_individual"
    ].mean()

    comparativo_bach = perfil_texto_comparativo_coloreado(
        promedio_bach,
        promedio_bach_grupo
    )

    comparativo_eval = perfil_texto_comparativo_coloreado(
        promedio_eval,
        promedio_eval_grupo
    )

    
    nivel_alerta = perfil_clasificar_alerta_tutoria(
        promedio_bach=promedio_bach,
        promedio_eval=promedio_eval,
        resultado_global=resultado_global,
        tabla_contexto=tabla_contexto
    )

    acciones = perfil_generar_acciones_acompanamiento(
        resultado_global=resultado_global,
        tabla_contexto=tabla_contexto
    )

    acciones_prevencion = " ".join(
        acciones["Prevención"]
    )

    acciones_correccion = " ".join(
        acciones["Corrección"]
    )

    # ------------------------------------------------------------
    # Dimensiones EVALUATEC comparadas contra pares
    # ------------------------------------------------------------

    if tabla_contexto.empty:
        texto_oportunidad = "Validar manualmente los resultados del estudiante."
        texto_dimensiones = (
            "No se cuenta con información suficiente para comparar "
            "las dimensiones del EVALUATEC."
        )

    else:
        tabla_ordenada = tabla_contexto.sort_values(
            "Resultado",
            ascending=True
        )

        areas_oportunidad = tabla_ordenada.head(2)
        areas_fuertes = tabla_ordenada.tail(2).sort_values(
            "Resultado",
            ascending=False
        )

        textos_fuertes = []

        for _, fila_area in areas_fuertes.iterrows():
            codigo = fila_area.get("Código", None)
            dimension = fila_area["Dimensión"]
            resultado = fila_area["Resultado"]

            columna_grupo = None

            for codigo_eval, etiqueta_eval in EVAL_ETIQUETAS_AREAS.items():
                if etiqueta_eval == dimension:
                    columna_grupo = f"Area_{codigo_eval}"
                    break

            if columna_grupo is not None and columna_grupo in grupo_referencia.columns:
                promedio_grupo_area = grupo_referencia[columna_grupo].mean()
                comparativo_area = perfil_texto_comparativo_coloreado(
                    resultado,
                    promedio_grupo_area
                )

                textos_fuertes.append(
                    f"<b>{dimension}</b> ({resultado:.1f}%), {comparativo_area} "
                    f"(<b>{perfil_formato_porcentaje(promedio_grupo_area)}</b>)"
                )
            else:
                textos_fuertes.append(
                    f"<b>{dimension}</b> ({resultado:.1f}%)"
                )

        textos_oportunidad = []

        for _, fila_area in areas_oportunidad.iterrows():
            dimension = fila_area["Dimensión"]
            resultado = fila_area["Resultado"]

            columna_grupo = None

            for codigo_eval, etiqueta_eval in EVAL_ETIQUETAS_AREAS.items():
                if etiqueta_eval == dimension:
                    columna_grupo = f"Area_{codigo_eval}"
                    break

            if columna_grupo is not None and columna_grupo in grupo_referencia.columns:
                promedio_grupo_area = grupo_referencia[columna_grupo].mean()
                comparativo_area = perfil_texto_comparativo_coloreado(
                    resultado,
                    promedio_grupo_area
                )
                
                textos_oportunidad.append(
                    f"<b>{dimension}</b> ({resultado:.1f}%), {comparativo_area} "
                    f"(<b>{perfil_formato_porcentaje(promedio_grupo_area)}</b>)"
                )
            else:
                textos_oportunidad.append(
                    f"<b>{dimension}</b> ({resultado:.1f}%)"
                )

        texto_fortalezas = "; ".join(textos_fuertes)
        texto_oportunidad = "; ".join(textos_oportunidad)

        texto_dimensiones = (
            f"Las áreas con mejor desempeño fueron {texto_fortalezas}. "
            f"Las áreas de oportunidad principales fueron {texto_oportunidad}."
        )

    puntos = [
        (
            "Nombre y procedencia",
            (
                f"{nombre} proviene de <b>{escuela}</b>, con procedencia registrada en "
                f"<b>{estado}</b>. Su promedio de bachillerato es "
                f"<b>{perfil_formato_porcentaje(promedio_bach)}</b>, el cual es "
                f"{comparativo_bach} del grupo de referencia "
                f"(<b>{perfil_formato_porcentaje(promedio_bach_grupo)}</b>)."
            )
        ),
        (
            "Resultado global EVALUATEC",
            (
                f"Obtuvo una calificación global de "
                f"<b>{perfil_formato_porcentaje(promedio_eval)}</b>. Este resultado es "
                f"{comparativo_eval} respecto al promedio de sus compañeros "
                f"(<b>{perfil_formato_porcentaje(promedio_eval_grupo)}</b>)."
            )
        ),
        (
            "Dimensiones EVALUATEC",
            texto_dimensiones
        )
    ]

    return nivel_alerta, puntos
def perfil_crear_contexto_academico(fila, df_evaluatec, tabla_areas):
    """
    Calcula posición del estudiante contra sus compañeros.
    No grafica el boxplot; solo usa su lógica para clasificar.
    """

    grupo = perfil_obtener_grupo_referencia(
        fila,
        df_evaluatec
    )

    promedio_eval = perfil_valor(
        fila,
        "eval_Promedio_global_individual",
        np.nan
    )

    resultado_global = perfil_clasificar_por_cuartiles(
        promedio_eval,
        grupo["Promedio_global_individual"]
    )

    registros_dimensiones = []

    for _, area in tabla_areas.iterrows():
        codigo = area["Código"]
        columna = f"Area_{codigo}"

        if columna not in grupo.columns:
            continue

        resultado_dimension = perfil_clasificar_por_cuartiles(
            area["Resultado"],
            grupo[columna]
        )

        registros_dimensiones.append(
            {
                "Código": codigo,
                "Dimensión": area["Dimensión"],
                "Resultado": area["Resultado"],
                "Clasificación": resultado_dimension["Clasificación"],
                "Nivel": resultado_dimension["Nivel"],
                "Q1": resultado_dimension["Q1"],
                "Q2": resultado_dimension["Q2"],
                "Q3": resultado_dimension["Q3"]
            }
        )

    tabla_contexto = pd.DataFrame(registros_dimensiones)

    return resultado_global, tabla_contexto, grupo


def perfil_clasificar_area_boxplot(resultado, columna_grupo, grupo_referencia):
    """
    Clasifica una dimensión contra el grupo de referencia usando lógica de boxplot.
    """

    if columna_grupo not in grupo_referencia.columns:
        return "Estudiante regular", "#111111"

    serie = grupo_referencia[columna_grupo].dropna()

    if pd.isna(resultado) or serie.empty or len(serie) < 5:
        return "Estudiante regular", "#111111"

    q1 = serie.quantile(0.25)
    q3 = serie.quantile(0.75)
    iqr = q3 - q1

    limite_inferior = q1 - 1.5 * iqr
    limite_superior = q3 + 1.5 * iqr

    if resultado < limite_inferior:
        return "Estudiante en riesgo de no acreditar", "#C62828"

    if resultado > limite_superior:
        return "Joven promesa", "#2E7D32"

    return "Estudiante regular", "#111111"
def perfil_normalizar_correo(valor):
    """Normaliza correos para cruces entre bases."""

    if pd.isna(valor):
        return ""

    texto = str(valor).strip().lower()

    if texto in ["", "nan", "none", "sin dato"]:
        return ""

    return texto


def perfil_extraer_correos_fila(fila):
    """
    Extrae posibles correos desde Historial y EVALUATEC.
    Busca columnas que contengan correo/email/mail.
    """

    correos = []

    for columna in fila.index:
        columna_limpia = util_limpiar_texto(columna)

        if (
            "correo" in columna_limpia
            or "email" in columna_limpia
            or "mail" in columna_limpia
        ):
            correo = perfil_normalizar_correo(fila[columna])

            if correo != "":
                correos.append(correo)

    return list(set(correos))


def perfil_score_nombre_tokens(nombre_a, nombre_b):
    """
    Calcula similitud flexible por nombre.
    Funciona mejor cuando falta un apellido o cambia el orden.
    """

    nombre_a = perfil_normalizar_nombre(nombre_a)
    nombre_b = perfil_normalizar_nombre(nombre_b)

    if nombre_a == "" or nombre_b == "":
        return 0

    tokens_a = set(nombre_a.split())
    tokens_b = set(nombre_b.split())

    tokens_a = {token for token in tokens_a if len(token) >= 3}
    tokens_b = {token for token in tokens_b if len(token) >= 3}

    if not tokens_a or not tokens_b:
        return 0

    interseccion = len(tokens_a.intersection(tokens_b))
    union = len(tokens_a.union(tokens_b))

    score_tokens = interseccion / union

    score_texto = SequenceMatcher(
        None,
        " ".join(sorted(tokens_a)),
        " ".join(sorted(tokens_b))
    ).ratio()

    return max(score_tokens, score_texto)

def perfil_interpretar_intensidad_chaside(nivel):
    """Interpreta el nivel de intensidad CHASIDE con color."""

    if pd.isna(nivel):
        return "no se cuenta con un nivel de intensidad definido"

    nivel = str(nivel).strip()

    if nivel == "Joven promesa":
        return (
            "se interpreta como "
            + perfil_texto_chaside_coloreado(
                "un perfil con alta congruencia vocacional",
                "positivo"
            )
            + ", por lo que se espera mejor adaptación académica, mayor claridad en su elección "
            + "y potencial para aprovechar actividades de alto desempeño"
        )

    if nivel == "Perfil en transición":
        return (
            "se interpreta como "
            + perfil_texto_chaside_coloreado(
                "un perfil funcional, aunque todavía en proceso de consolidación",
                "positivo"
            )
            + "; podría presentar "
            + perfil_texto_chaside_coloreado(
                "algunas dificultades de adaptación o acreditación en ciertas asignaturas",
                "negativo"
            )
            + " durante su formación académica, por lo que se recomienda acompañamiento preventivo"
        )

    if nivel == "Perfil en riesgo":
        return (
            "se interpreta como "
            + perfil_texto_chaside_coloreado(
                "una coincidencia vocacional baja",
                "negativo"
            )
            + "; puede requerir seguimiento tutorial más cercano para prevenir desmotivación, "
            + "bajo desempeño o dudas sobre la carrera elegida"
        )

    if nivel == "Sin perfil":
        return (
            "se interpreta como "
            + perfil_texto_chaside_coloreado(
                "baja correspondencia entre la carrera elegida y el perfil vocacional detectado",
                "negativo"
            )
            + "; se recomienda orientación vocacional complementaria y seguimiento individual"
        )

    if nivel == "Sin nivel definido":
        return (
            "no permite establecer una intensidad vocacional clara, por lo que se sugiere interpretar "
            "el resultado con cautela"
        )

    return "requiere interpretación complementaria por parte del tutor o responsable académico"

def perfil_generar_pdf_dictamen(nombre, carrera, nivel_alerta, dictamen_tutoria):
    """Genera PDF simple del dictamen tutorial."""

    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib import colors

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=1.8 * cm,
        leftMargin=1.8 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm
    )

    styles = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "TituloDictamen",
        parent=styles["Title"],
        fontSize=18,
        leading=22,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#111111"),
        spaceAfter=12
    )

    estilo_subtitulo = ParagraphStyle(
        "SubtituloDictamen",
        parent=styles["Normal"],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#555555"),
        spaceAfter=16
    )

    estilo_punto = ParagraphStyle(
        "PuntoDictamen",
        parent=styles["Normal"],
        fontSize=10.5,
        leading=16,
        textColor=colors.HexColor("#111111"),
        spaceAfter=12
    )

    elementos = []

    elementos.append(
        Paragraph("Dictamen tutorial", estilo_titulo)
    )

    elementos.append(
        Paragraph(
            f"<b>Estudiante:</b> {nombre}<br/>"
            f"<b>Carrera:</b> {carrera}<br/>"
            f"<b>Clasificación tutorial:</b> {nivel_alerta}",
            estilo_subtitulo
        )
    )

    for numero, (titulo, contenido) in enumerate(dictamen_tutoria, start=1):
        contenido_pdf = str(contenido)
        contenido_pdf = contenido_pdf.replace("<span", "<font")
        contenido_pdf = contenido_pdf.replace("</span>", "</font>")

        elementos.append(
            Paragraph(
                f"<b>{numero}. {titulo}:</b> {contenido_pdf}",
                estilo_punto
            )
        )

        elementos.append(Spacer(1, 6))

    doc.build(elementos)

    buffer.seek(0)
    return buffer.getvalue()
    
def render_perfil_individual():
    st.title("👤 Perfil individual del aspirante")
    st.caption(
        "Cruce entre Historial de Aspirantes y EVALUATEC usando nombre como llave principal."
    )

    st.sidebar.markdown("---")
    st.sidebar.subheader("Carga Perfil Individual")

    archivos_eval = st.sidebar.file_uploader(
        "Carga los 3 CSV de EVALUATEC",
        type=["csv"],
        accept_multiple_files=True,
        key="perfil_eval_archivos"
    )

    archivo_historial = st.sidebar.file_uploader(
        "Carga el Excel de Historial de Aspirantes",
        type=["xlsx", "xls"],
        key="perfil_historial_archivo"
    )

    if not archivos_eval or archivo_historial is None:
        st.info(
            "Para construir el perfil individual carga los archivos de EVALUATEC "
            "y el Excel de Historial de Aspirantes."
        )
        st.stop()

    if len(archivos_eval) != 3:
        st.warning(
            f"Cargaste {len(archivos_eval)} archivo(s) de EVALUATEC. "
            "Se requieren exactamente 3."
        )
        st.stop()

    with st.spinner("Procesando y cruzando bases de datos..."):
        df_evaluatec, errores_eval = perfil_preparar_evaluatec(
            archivos_eval
        )

        df_historial, df_bitacora = perfil_preparar_historial(
            archivo_historial.getvalue()
        )

    if errores_eval:
        for error in errores_eval:
            st.warning(error)

    if df_evaluatec.empty:
        st.error("No se pudo procesar la base de EVALUATEC.")
        st.stop()

    if df_historial.empty:
        st.error("No se pudo procesar el Historial de Aspirantes.")
        st.dataframe(df_bitacora, use_container_width=True)
        st.stop()

    df_cruzado = perfil_crear_base_cruzada(
        df_historial=df_historial,
        df_evaluatec=df_evaluatec
    )

    st.markdown("## Búsqueda del aspirante")

    sistema_busqueda = st.radio(
        "Selecciona sistema de búsqueda",
        [
            "🔤 Búsqueda por inicial del apellido",
            "🎓 Búsqueda por carrera"
        ],
        horizontal=True
    )
    
    df_busqueda = df_cruzado.copy()

    if sistema_busqueda == "🔤 Búsqueda por inicial del apellido":

        letras_disponibles = sorted(
            [
                letra
                for letra in df_busqueda["Letra_apellido"]
                .dropna()
                .astype(str)
                .str.upper()
                .unique()
                if letra != ""
            ]
        )

        letra_seleccionada = st.selectbox(
            "Selecciona la letra del apellido",
            options=letras_disponibles
        )

        df_filtrado = df_busqueda[
            df_busqueda["Letra_apellido"] == letra_seleccionada
        ].copy()

    else:
        carreras_disponibles = sorted(
            df_busqueda["Carrera_referencia"]
            .dropna()
            .astype(str)
            .unique()
        )

        carrera_seleccionada = st.selectbox(
            "Selecciona la carrera",
            options=carreras_disponibles
        )

        df_filtrado = df_busqueda[
            df_busqueda["Carrera_referencia"] == carrera_seleccionada
        ].copy()

    df_filtrado = df_filtrado.sort_values(
        "Nombre_visible"
    ).reset_index(drop=False)

    if df_filtrado.empty:
        st.warning("No hay aspirantes con esos criterios.")
        st.stop()

    opciones = {
        fila["Etiqueta_selector"]: fila["index"]
        for _, fila in df_filtrado.iterrows()
    }

    opcion_seleccionada = st.selectbox(
        "Selecciona al aspirante",
        options=list(opciones.keys())
    )

    indice_original = opciones[opcion_seleccionada]
    fila = df_cruzado.loc[indice_original]

    st.markdown("---")
    st.markdown("## Tarjeta de información del aspirante")
    nombre = perfil_valor(fila, "Nombre_visible")
    carrera_historial = perfil_valor(fila, "Carrera_historial")
    carrera_eval = perfil_valor(fila, "Carrera_evaluatec")
    estatus_cruce = perfil_valor(fila, "Estatus_cruce")

    promedio_bach = perfil_valor(
        fila,
        "hist_Promedio_normalizado_100",
        np.nan
    )

    promedio_eval = perfil_valor(
        fila,
        "eval_Promedio_global_individual",
        np.nan
    )

    sexo = perfil_valor(fila, "hist_Sexo_normalizado")

    escuela = perfil_valor(
        fila,
        "hist_Bachillerato_procedencia_original",
        perfil_valor(fila, "hist_Bachillerato_procedencia")
    )

    estado = perfil_valor(fila, "hist_Estado_procedencia")
    matricula = perfil_valor(fila, "hist_ID_aspirante")
    
    dictamen_chaside = (
        "No se encontró registro CHASIDE para este aspirante. "
        "Se recomienda solicitarle responder la escala para complementar el dictamen tutorial."
    )
    # ------------------------------------------------------------
    # CHASIDE para dictamen tutorial
    # ------------------------------------------------------------

    dictamen_chaside = (
        "No se encontró registro CHASIDE para este aspirante. "
        "Se recomienda solicitarle responder la escala para complementar el dictamen tutorial."
    )

    url_chaside_perfil = st.session_state.get(
        "url_chaside_global",
        "https://docs.google.com/spreadsheets/d/1YHMEb5hftOZfV-CMWoUsUgJh1xmsgTY3YYwAtq1dGQA/edit?resourcekey=&gid=1491376423#gid=1491376423"
    )

    try:
        df_chaside_raw_perfil = chaside_cargar_respuestas(
            url_chaside_perfil
        )

        resultado_chaside_perfil = chaside_procesar_respuestas(
            df_chaside_raw_perfil,
            peso_intereses=0.8,
            peso_aptitudes=0.2
        )

        if isinstance(resultado_chaside_perfil, tuple):
            df_chaside_perfil = resultado_chaside_perfil[0]
        else:
            df_chaside_perfil = resultado_chaside_perfil

        nombre_actual_norm = perfil_normalizar_nombre(nombre)
        carrera_actual_norm = perfil_simplificar_carrera(carrera_historial)

        correos_aspirante = perfil_extraer_correos_fila(fila)

        df_chaside_match = df_chaside_perfil.copy()

        df_chaside_match["Nombre_match"] = df_chaside_match[
            CHASIDE_COLUMNA_NOMBRE
        ].apply(perfil_normalizar_nombre)

        df_chaside_match["Carrera_match"] = df_chaside_match[
            CHASIDE_COLUMNA_CARRERA
        ].apply(perfil_simplificar_carrera)

        df_chaside_match["Correo_CHASIDE_1"] = ""

        if CHASIDE_COLUMNA_EMAIL_1 in df_chaside_match.columns:
            df_chaside_match["Correo_CHASIDE_1"] = df_chaside_match[
                CHASIDE_COLUMNA_EMAIL_1
            ].apply(perfil_normalizar_correo)

        df_chaside_match["Correo_CHASIDE_2"] = ""

        if CHASIDE_COLUMNA_EMAIL_2 in df_chaside_match.columns:
            df_chaside_match["Correo_CHASIDE_2"] = df_chaside_match[
                CHASIDE_COLUMNA_EMAIL_2
            ].apply(perfil_normalizar_correo)

        df_chaside_match["Coincide_correo"] = df_chaside_match.apply(
            lambda fila_ch: (
                fila_ch["Correo_CHASIDE_1"] in correos_aspirante
                or fila_ch["Correo_CHASIDE_2"] in correos_aspirante
            ),
            axis=1
        )

        df_chaside_match["Coincide_carrera"] = (
            df_chaside_match["Carrera_match"] == carrera_actual_norm
        )

        df_chaside_match["Score_nombre"] = df_chaside_match[
            CHASIDE_COLUMNA_NOMBRE
        ].apply(
            lambda valor: perfil_score_nombre_tokens(
                valor,
                nombre
            )
        )

        df_chaside_match = df_chaside_match.sort_values(
            [
                "Coincide_correo",
                "Coincide_carrera",
                "Score_nombre"
            ],
            ascending=[
                False,
                False,
                False
            ]
        )

        if not df_chaside_match.empty:
            mejor_match = df_chaside_match.iloc[0]

            if (
                mejor_match["Coincide_correo"]
                or (
                    mejor_match["Coincide_carrera"]
                    and mejor_match["Score_nombre"] >= 0.45
                )
                or mejor_match["Score_nombre"] >= 0.70
            ):
                fila_chaside = mejor_match

                semaforo_chaside = fila_chaside["Semáforo vocacional"]
                diagnostico_chaside = fila_chaside["Diagnóstico vocacional"]
                nivel_chaside = fila_chaside["Nivel de intensidad"]
                carrera_mejor_chaside = fila_chaside["Carrera_Mejor_Perfilada"]
                carrera_respondida_chaside = fila_chaside[CHASIDE_COLUMNA_CARRERA]

                interpretacion_intensidad = perfil_interpretar_intensidad_chaside(
                    nivel_chaside
                )

                if semaforo_chaside == "Verde":
                    semaforo_html = perfil_texto_chaside_coloreado(
                        f"{semaforo_chaside} ({diagnostico_chaside})",
                        "positivo"
                    )

                elif semaforo_chaside in ["Rojo", "Respondió siempre igual"]:
                    semaforo_html = perfil_texto_chaside_coloreado(
                        f"{semaforo_chaside} ({diagnostico_chaside})",
                        "negativo"
                    )

                else:
                    semaforo_html = f"<b>{semaforo_chaside} ({diagnostico_chaside})</b>"

                if diagnostico_chaside == "Perfil adecuado":
                    texto_adecuacion = (
                        "lo que indica "
                        + perfil_texto_chaside_coloreado(
                            "un perfil vocacional adecuado para la carrera elegida",
                            "positivo"
                        )
                    )

                    texto_carrera_mejor = (
                        f"La carrera respondida en CHASIDE fue <b>{carrera_respondida_chaside}</b>, "
                        "y el resultado mantiene "
                        + perfil_texto_chaside_coloreado(
                            "correspondencia con dicha elección",
                            "positivo"
                        )
                        + "."
                    )

                else:
                    texto_adecuacion = (
                        "lo que indica "
                        + perfil_texto_chaside_coloreado(
                            "un perfil vocacional no completamente adecuado para la carrera elegida",
                            "negativo"
                        )
                    )

                    texto_carrera_mejor = (
                        f"La carrera respondida en CHASIDE fue <b>{carrera_respondida_chaside}</b>; "
                        f"sin embargo, el perfil sugiere mayor afinidad hacia "
                        + perfil_texto_chaside_coloreado(
                            carrera_mejor_chaside,
                            "positivo"
                        )
                        + "."
                    )

                dictamen_chaside = (
                    f"La estudiante sí contestó la escala CHASIDE, donde se detectó un perfil "
                    f"{semaforo_html}, {texto_adecuacion}. "
                    f"{texto_carrera_mejor} "
                    f"Por otro lado, el nivel de intensidad se detectó como "
                    f"<b>{nivel_chaside}</b>, lo cual {interpretacion_intensidad}."
                )

    except Exception:
        pass

    # ------------------------------------------------------------
    # Cálculos EVALUATEC para dictamen, sin mostrar gráficas
    # ------------------------------------------------------------

    tabla_areas = perfil_obtener_areas_individuales(fila)

    resultado_global, tabla_contexto, grupo_referencia = perfil_crear_contexto_academico(
        fila=fila,
        df_evaluatec=df_evaluatec,
        tabla_areas=tabla_areas
    )                

    nivel_alerta, dictamen_tutoria = perfil_generar_dictamen_tutoria(
        fila=fila,
        df_historial=df_historial,
        grupo_referencia=grupo_referencia,
        resultado_global=resultado_global,
        tabla_contexto=tabla_contexto
    )
    dictamen_tutoria.append(
        (
            "Resultado vocacional CHASIDE",
            dictamen_chaside
        )
    )
    
    configuracion = perfil_configuracion_alerta_tutoria(
        nivel_alerta
    )

    st.markdown("## Dictamen tutorial")

    st.markdown(
        f"""
        <div style="
            background-color: {configuracion['color_fondo']};
            border-left: 10px solid {configuracion['color_borde']};
            padding: 20px 24px;
            border-radius: 16px;
            margin-top: 12px;
            margin-bottom: 18px;
            color: #111111;
        ">
            <h3 style="
                margin-top: 0;
                margin-bottom: 4px;
                color: #111111;
            ">
                {configuracion['titulo']}
            </h3>
            <p style="margin-bottom: 0; color:#111111;">
                Clasificación general para orientar el seguimiento tutorial.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    for numero, (titulo, contenido) in enumerate(
        dictamen_tutoria,
        start=1
    ):

        html_punto = f'''
        <div style="
            background-color: #FFFFFF;
            border: 1px solid #F0D6D6;
            border-left: 6px solid {configuracion["color_borde"]};
            padding: 14px 18px;
            border-radius: 12px;
            margin-bottom: 10px;
            color: #111111;
            font-size: 17px;
            line-height: 1.55;
        ">
            <b style="color:#111111;">{numero}. {titulo}:</b> {contenido}
        </div>
        '''

        st.markdown(
            html_punto,
            unsafe_allow_html=True
        )


# ============================================================
# ============================================================
# FUNCIONES BASE CHASIDE
# ============================================================

def chaside_transformar_url_google_sheets(url):
    """
    Convierte una URL editable de Google Sheets a una URL CSV descargable.
    Usa por defecto el gid de la pestaña de respuestas CHASIDE.
    """

    url = str(url).strip()

    if "export?format=csv" in url:
        return url

    if "docs.google.com/spreadsheets" not in url:
        return url

    try:
        file_id = url.split("/d/")[1].split("/")[0]

        gid = "1491376423"

        if "gid=" in url:
            gid = url.split("gid=")[-1].split("&")[0].split("#")[0]

        resourcekey = ""

        if "resourcekey=" in url:
            resourcekey = url.split("resourcekey=")[-1].split("&")[0].split("#")[0]

        url_csv = (
            f"https://docs.google.com/spreadsheets/d/"
            f"{file_id}/export?format=csv&gid={gid}"
        )

        if resourcekey != "":
            url_csv = f"{url_csv}&resourcekey={resourcekey}"

        return url_csv

    except Exception:
        raise ValueError(
            "No se pudo transformar el enlace de Google Sheets. "
            "Pega el enlace completo de la hoja de respuestas."
        )
@st.cache_data(show_spinner=False)
def chaside_cargar_respuestas(url):
    """
    Carga respuestas CHASIDE desde Google Sheets.
    El archivo debe estar compartido como 'Cualquier persona con el enlace puede ver'.
    """

    url_csv = chaside_transformar_url_google_sheets(url)

    try:
        return pd.read_csv(url_csv)

    except Exception as error:
        raise ValueError(
            "No fue posible leer la hoja de respuestas. "
            "Verifica que el Google Sheet esté compartido como "
            "'Cualquier persona con el enlace puede ver' y que el enlace corresponda "
            "a la pestaña de respuestas del formulario. "
            f"Detalle técnico: {error}"
        )
def chaside_col_item(columnas_items, i):
    return columnas_items[i - 1]


def chaside_procesar_respuestas(
    df_raw,
    peso_intereses=0.8,
    peso_aptitudes=0.2
):
    """Procesa respuestas CHASIDE y genera diagnóstico vocacional."""

    df = df_raw.copy()
    df.columns = df.columns.str.strip()

    faltantes = [
        columna
        for columna in [
            CHASIDE_COLUMNA_NOMBRE,
            CHASIDE_COLUMNA_CARRERA
        ]
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"Faltan columnas requeridas en CHASIDE: {faltantes}. "
            f"Columnas detectadas: {list(df.columns)}"
        )

    columnas_items = df.columns[6:104]

    if len(columnas_items) != 98:
        raise ValueError(
            f"Se esperaban 98 reactivos CHASIDE, "
            f"pero se detectaron {len(columnas_items)}. "
            "Revisa que el enlace corresponda a la hoja de respuestas del formulario."
        )

    df_items = (
        df[columnas_items]
        .astype(str)
        .apply(lambda col: col.str.strip().str.lower())
        .replace({
            "sí": 1,
            "si": 1,
            "s": 1,
            "1": 1,
            "true": 1,
            "verdadero": 1,
            "x": 1,
            "no": 0,
            "n": 0,
            "0": 0,
            "false": 0,
            "falso": 0,
            "": 0,
            "nan": 0
        })
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
        .astype(int)
    )

    df[columnas_items] = df_items

    df["Desv_Intrapersona"] = df[columnas_items].std(axis=1)
    umbral = df["Desv_Intrapersona"].quantile(0.10)
    df["Respondio_Siempre_Igual"] = df["Desv_Intrapersona"] <= umbral

    for area in CHASIDE_AREAS:
        df[f"INTERES_{area}"] = df[
            [
                chaside_col_item(columnas_items, i)
                for i in CHASIDE_INTERESES_ITEMS[area]
            ]
        ].sum(axis=1)

        df[f"APTITUD_{area}"] = df[
            [
                chaside_col_item(columnas_items, i)
                for i in CHASIDE_APTITUDES_ITEMS[area]
            ]
        ].sum(axis=1)

    for area in CHASIDE_AREAS:
        df[f"PUNTAJE_COMBINADO_{area}"] = (
            df[f"INTERES_{area}"] * peso_intereses
            +
            df[f"APTITUD_{area}"] * peso_aptitudes
        )

        df[f"TOTAL_{area}"] = (
            df[f"INTERES_{area}"]
            +
            df[f"APTITUD_{area}"]
        )

    df["Area_Fuerte_Ponderada"] = df.apply(
        lambda fila: max(
            CHASIDE_AREAS,
            key=lambda area: fila[f"PUNTAJE_COMBINADO_{area}"]
        ),
        axis=1
    )

    columnas_score = [
        f"PUNTAJE_COMBINADO_{area}"
        for area in CHASIDE_AREAS
    ]

    df["Score_CHASIDE"] = df[columnas_score].max(axis=1)

    def evaluar_coincidencia(area_chaside, carrera):
        carrera = str(carrera).strip()
        perfil = CHASIDE_PERFILES_CARRERA.get(carrera)

        if not perfil:
            return "Sin perfil definido"

        if area_chaside in perfil:
            return "Coherente"

        return "No coincidente"

    df["Coincidencia_CHASIDE"] = df.apply(
        lambda fila: evaluar_coincidencia(
            fila["Area_Fuerte_Ponderada"],
            fila[CHASIDE_COLUMNA_CARRERA]
        ),
        axis=1
    )

    def carrera_mejor_perfilada(fila):
        if fila["Respondio_Siempre_Igual"]:
            return "Información no confiable"

        area = fila["Area_Fuerte_Ponderada"]
        carrera_actual = str(fila[CHASIDE_COLUMNA_CARRERA]).strip()

        sugeridas = [
            carrera
            for carrera, letras in CHASIDE_PERFILES_CARRERA.items()
            if area in letras
        ]

        if carrera_actual in sugeridas:
            return carrera_actual

        if sugeridas:
            return ", ".join(sugeridas)

        return "Sin sugerencia clara"

    df["Carrera_Mejor_Perfilada"] = df.apply(
        carrera_mejor_perfilada,
        axis=1
    )

    def diagnostico(fila):
        carrera_actual = str(fila[CHASIDE_COLUMNA_CARRERA]).strip()
        mejor = str(fila["Carrera_Mejor_Perfilada"]).strip()

        if fila["Respondio_Siempre_Igual"]:
            return "Información no confiable"

        if mejor == carrera_actual:
            return "Perfil adecuado"

        if mejor == "Sin sugerencia clara":
            return "Sin sugerencia clara"

        return f"Sugerencia: {mejor}"

    df["Diagnóstico vocacional"] = df.apply(
        diagnostico,
        axis=1
    )

    def semaforo(fila):
        if fila["Respondio_Siempre_Igual"]:
            return "Respondió siempre igual"

        if fila["Diagnóstico vocacional"] == "Sin sugerencia clara":
            return "Sin sugerencia"

        if (
            fila["Diagnóstico vocacional"] == "Perfil adecuado"
            and fila["Coincidencia_CHASIDE"] == "Coherente"
        ):
            return "Verde"

        if fila["Diagnóstico vocacional"] == "Perfil adecuado":
            return "Amarillo"

        if str(fila["Diagnóstico vocacional"]).startswith("Sugerencia:"):
            return "Amarillo"

        return "Rojo"

    df["Semáforo vocacional"] = df.apply(
        semaforo,
        axis=1
    )

    # ------------------------------------------------------------
    # Nivel de intensidad vocacional
    # ------------------------------------------------------------

    df["Nivel de intensidad"] = "Sin nivel definido"

    for carrera, grupo in df.groupby(CHASIDE_COLUMNA_CARRERA):

        indices_amarillos = grupo[
            grupo["Semáforo vocacional"] == "Amarillo"
        ].sort_values(
            "Score_CHASIDE",
            ascending=True
        ).index.tolist()

        indices_verdes = grupo[
            grupo["Semáforo vocacional"] == "Verde"
        ].sort_values(
            "Score_CHASIDE",
            ascending=True
        ).index.tolist()

        if len(indices_amarillos) > 0:
            total_amarillos = len(indices_amarillos)

            for posicion, indice in enumerate(indices_amarillos, start=1):
                rank_pct = posicion / total_amarillos

                if rank_pct <= 0.25:
                    df.loc[indice, "Nivel de intensidad"] = "Sin perfil"
                else:
                    df.loc[indice, "Nivel de intensidad"] = "Perfil en riesgo"

        if len(indices_verdes) > 0:
            total_verdes = len(indices_verdes)

            for posicion, indice in enumerate(indices_verdes, start=1):
                rank_pct = posicion / total_verdes

                if rank_pct > 0.75:
                    df.loc[indice, "Nivel de intensidad"] = "Joven promesa"
                else:
                    df.loc[indice, "Nivel de intensidad"] = "Perfil en transición"
    return df
# ============================================================
# EJECUCIÓN DE LA APP
# ============================================================

if modulo_activo == "📂 Carga de archivos":
    cargar_datos_globales()

elif modulo_activo == "📘 EVALUATEC 2026":
    render_evaluatec()

elif modulo_activo == "🎓 Historial de Aspirantes":
    render_historial()

elif modulo_activo == "🧭 Diagnóstico vocacional CHASIDE":
    render_modulo_chaside_con_carga()

elif modulo_activo == "👤 Perfil individual":
    render_perfil_individual()

    configuracion = perfil_configuracion_alerta_tutoria(
        nivel_alerta
    )

    st.markdown("## Dictamen tutorial")

    st.markdown(
        f"""
        <div style="
            background-color: {configuracion['color_fondo']};
            border-left: 10px solid {configuracion['color_borde']};
            padding: 20px 24px;
            border-radius: 16px;
            margin-top: 12px;
            margin-bottom: 18px;
            color: #111111;
        ">
            <h3 style="
                margin-top: 0;
                margin-bottom: 4px;
                color: #111111;
            ">
                {configuracion['titulo']}
            </h3>
            <p style="margin-bottom: 0; color:#111111;">
                Clasificación general para orientar el seguimiento tutorial.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    for numero, (titulo, contenido) in enumerate(
        dictamen_tutoria,
        start=1
    ):

        html_punto = f'''
        <div style="
            background-color: #FFFFFF;
            border: 1px solid #F0D6D6;
            border-left: 6px solid {configuracion["color_borde"]};
            padding: 14px 18px;
            border-radius: 12px;
            margin-bottom: 10px;
            color: #111111;
            font-size: 17px;
            line-height: 1.55;
        ">
            <b style="color:#111111;">{numero}. {titulo}:</b> {contenido}
        </div>
        '''

        st.markdown(
            html_punto,
            unsafe_allow_html=True
        )

    pdf_dictamen = perfil_generar_pdf_dictamen(
        nombre=nombre,
        carrera=carrera_historial,
        configuracion=configuracion,
        dictamen_tutoria=dictamen_tutoria
    )

    nombre_archivo_pdf = (
        "dictamen_tutorial_"
        + perfil_normalizar_nombre(nombre).lower().replace(" ", "_")
        + ".pdf"
    )

    st.download_button(
        label="⬇️ Descargar dictamen tutorial en PDF",
        data=pdf_dictamen,
        file_name=nombre_archivo_pdf,
        mime="application/pdf",
        use_container_width=True
    )
