# app_conciliador_264_260.py
# App para consolidar reportes 264 y reemplazar fechas desde reportes 260
# Autor: Generado para Sr. Stark
#
# Ejecución:
#   streamlit run app_conciliador_264_260.py
#
# Dependencias:
#   pip install streamlit pandas openpyxl xlrd xlsxwriter

import io
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Conciliador 264 vs 260",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# FUNCIONES DE LIMPIEZA
# ============================================================

def normalizar_texto(valor) -> str:
    """
    Normaliza texto para comparar nombres de columnas y valores.
    Quita tildes, símbolos raros y espacios dobles.
    """
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.upper()
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_doc(valor) -> str:
    """
    Normaliza el Doc Paciente.
    Ejemplos:
    - 305521 -> 305521
    - CLI306039 -> 306039
    - CLI 306039 -> 306039
    """
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()
    digitos = re.sub(r"\D+", "", texto)

    if digitos == "":
        return ""

    # Quitar ceros a la izquierda solo si existen, sin dejar vacío
    digitos = digitos.lstrip("0") or "0"
    return digitos


def limpiar_numero(valor) -> float:
    """
    Convierte valores monetarios a número.
    Soporta formatos:
    - 123456
    - 123.456
    - 123,456
    - $ 123.456,78
    - 123456.78
    """
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if texto == "":
        return 0.0

    texto = texto.replace("$", "").replace(" ", "")

    # Si tiene punto y coma, asumimos formato colombiano: 1.234.567,89
    if "." in texto and "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto and "." not in texto:
        # Puede ser decimal con coma o miles con coma
        partes = texto.split(",")
        if len(partes[-1]) == 2:
            texto = texto.replace(",", ".")
        else:
            texto = texto.replace(",", "")
    else:
        # Si solo tiene puntos, puede ser miles o decimal.
        partes = texto.split(".")
        if len(partes) > 2:
            texto = texto.replace(".", "")

    texto = re.sub(r"[^0-9.\-]+", "", texto)

    try:
        return float(texto)
    except ValueError:
        return 0.0


def convertir_fecha(valor):
    """
    Convierte una fecha a datetime.
    Si no puede, retorna NaT.
    """
    if pd.isna(valor):
        return pd.NaT

    try:
        return pd.to_datetime(valor, errors="coerce")
    except Exception:
        return pd.NaT


def primer_no_vacio(serie):
    """
    Retorna el primer valor no vacío de una serie.
    """
    for valor in serie:
        if pd.notna(valor) and str(valor).strip() != "":
            return valor
    return ""


def valores_unicos_limpios(serie):
    """
    Retorna valores únicos limpios en texto separados por coma.
    """
    valores = []
    for valor in serie:
        if pd.notna(valor) and str(valor).strip() != "":
            texto = str(valor).strip()
            if texto not in valores:
                valores.append(texto)
    return ", ".join(valores)


def hacer_columnas_unicas(columnas):
    """
    Evita columnas duplicadas.
    """
    nuevas = []
    contador = {}

    for col in columnas:
        nombre = str(col).strip() if pd.notna(col) else ""
        if nombre == "":
            nombre = "SIN_NOMBRE"

        if nombre not in contador:
            contador[nombre] = 0
            nuevas.append(nombre)
        else:
            contador[nombre] += 1
            nuevas.append(f"{nombre}_{contador[nombre]}")

    return nuevas


def encontrar_columna(columnas, debe_contener=None, no_debe_contener=None, alternativas=None):
    """
    Encuentra una columna por palabras clave normalizadas.
    """
    debe_contener = debe_contener or []
    no_debe_contener = no_debe_contener or []
    alternativas = alternativas or []

    columnas_norm = [(col, normalizar_texto(col)) for col in columnas]

    # Búsqueda por lista de alternativas exactas normalizadas
    for alt in alternativas:
        alt_norm = normalizar_texto(alt)
        for col, col_norm in columnas_norm:
            if col_norm == alt_norm:
                return col

    # Búsqueda por contenido obligatorio/prohibido
    for col, col_norm in columnas_norm:
        if all(palabra in col_norm for palabra in debe_contener) and not any(palabra in col_norm for palabra in no_debe_contener):
            return col

    return None


def detectar_fila_encabezado(raw: pd.DataFrame) -> int:
    """
    Detecta la fila de encabezado del reporte.
    Funciona para reportes tipo Crystal Reports con filas de título arriba.
    """
    claves = [
        "FECHA",
        "DOC",
        "PCTE",
        "PACIENTE",
        "CONT",
        "APB",
        "CONVENIO",
        "IDENTIFICACION",
        "SERVICIO",
        "CANT",
        "VLR",
        "VALOR",
    ]

    mejor_indice = 0
    mejor_puntaje = -1

    limite = min(len(raw), 30)

    for i in range(limite):
        fila_texto = " ".join(normalizar_texto(x) for x in raw.iloc[i].tolist())
        puntaje = sum(1 for clave in claves if clave in fila_texto)

        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_indice = i

    return mejor_indice


def leer_excel_crystal(archivo) -> pd.DataFrame:
    """
    Lee un archivo Excel .xls o .xlsx proveniente de Crystal Reports.
    Devuelve un DataFrame limpio con encabezados detectados.
    """
    nombre = getattr(archivo, "name", str(archivo))
    extension = Path(nombre).suffix.lower()

    if hasattr(archivo, "getvalue"):
        contenido = archivo.getvalue()
    else:
        with open(archivo, "rb") as f:
            contenido = f.read()

    buffer = io.BytesIO(contenido)

    try:
        if extension == ".xls":
            raw = pd.read_excel(buffer, header=None, dtype=object, engine="xlrd")
        else:
            raw = pd.read_excel(buffer, header=None, dtype=object, engine="openpyxl")
    except ImportError as e:
        if extension == ".xls":
            raise ImportError(
                "Falta instalar xlrd para leer archivos .xls. Ejecuta: pip install xlrd"
            ) from e
        raise
    except Exception as e:
        raise ValueError(f"No se pudo leer el archivo {nombre}. Detalle: {e}") from e

    raw = raw.dropna(how="all").reset_index(drop=True)

    if raw.empty:
        raise ValueError(f"El archivo {nombre} está vacío o no tiene datos legibles.")

    fila_header = detectar_fila_encabezado(raw)
    columnas = hacer_columnas_unicas(raw.iloc[fila_header].tolist())

    df = raw.iloc[fila_header + 1:].copy()
    df.columns = columnas

    # Quitar filas completamente vacías
    df = df.dropna(how="all").reset_index(drop=True)

    # Quitar filas que son encabezados repetidos o basura del reporte
    def es_fila_basura(fila):
        texto = " ".join(normalizar_texto(x) for x in fila.tolist())
        if texto.strip() == "":
            return True
        if "SUCURSAL" in texto and "CONVENIO" in texto:
            return True
        if "FECHA" in texto and "DOC" in texto and ("PACIENTE" in texto or "PCTE" in texto):
            return True
        if "TOTAL" in texto and "REGISTROS" in texto:
            return True
        if "CRYSTAL" in texto:
            return True
        return False

    mascara_basura = df.apply(es_fila_basura, axis=1)
    df = df.loc[~mascara_basura].reset_index(drop=True)

    # Quitar columnas totalmente vacías
    df = df.dropna(axis=1, how="all")

    return df


# ============================================================
# PROCESAMIENTO DE ARCHIVOS 260
# ============================================================

def procesar_archivo_260(archivo) -> pd.DataFrame:
    """
    Extrae del 260:
    - FECHA_260: primera columna del reporte
    - DOC_PACIENTE
    """
    nombre_archivo = getattr(archivo, "name", str(archivo))
    df = leer_excel_crystal(archivo)

    if df.empty:
        return pd.DataFrame()

    fecha_col = df.columns[0]

    doc_col = encontrar_columna(
        df.columns,
        debe_contener=["DOC"],
        no_debe_contener=["CONT"],
        alternativas=[
            "Doc Paciente",
            "DOC PACIENTE",
            "DOC. PCTE.",
            "Doc Pcte",
            "DOC PCTE",
        ],
    )

    if doc_col is None:
        raise ValueError(
            f"No encontré la columna Doc Paciente en el archivo 260: {nombre_archivo}"
        )

    # IMPORTANTE: crear el DataFrame con el mismo índice del archivo leído.
    # Si se asigna un texto fijo sobre un DataFrame vacío, pandas deja esa columna en NaN.
    # Luego groupby elimina los NaN y el reporte final sale en cero. Trampa elegante, como siempre.
    salida = pd.DataFrame(index=df.index)
    salida["ARCHIVO_260"] = nombre_archivo
    salida["DOC_PACIENTE_ORIGINAL_260"] = df[doc_col]
    salida["DOC_NORM"] = df[doc_col].apply(normalizar_doc)
    salida["FECHA_260"] = df[fecha_col].apply(convertir_fecha)

    salida = salida[
        (salida["DOC_NORM"] != "") &
        (salida["FECHA_260"].notna())
    ].copy()

    salida = salida.drop_duplicates(
        subset=["ARCHIVO_260", "DOC_NORM", "FECHA_260"]
    ).reset_index(drop=True)

    return salida


def consolidar_fechas_260(archivos_260) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Consolida fechas de todos los 260.
    Si un mismo Doc Paciente tiene varias fechas diferentes, lo marca como conflicto.
    """
    lista = []

    for archivo in archivos_260:
        tmp = procesar_archivo_260(archivo)
        if not tmp.empty:
            lista.append(tmp)

    if not lista:
        return pd.DataFrame(), pd.DataFrame()

    base = pd.concat(lista, ignore_index=True)

    resumen = (
        base.groupby("DOC_NORM", as_index=False)
        .agg(
            FECHA_260=("FECHA_260", "min"),
            FECHAS_260_DISTINTAS=("FECHA_260", lambda x: x.dropna().nunique()),
            ARCHIVOS_260=("ARCHIVO_260", valores_unicos_limpios),
            DOC_PACIENTE_ORIGINAL_260=("DOC_PACIENTE_ORIGINAL_260", primer_no_vacio),
        )
    )

    resumen["CONFLICTO_FECHA_260"] = resumen["FECHAS_260_DISTINTAS"].apply(
        lambda x: "SI" if x > 1 else "NO"
    )

    conflictos = resumen[resumen["CONFLICTO_FECHA_260"] == "SI"].copy()

    return resumen, conflictos


# ============================================================
# PROCESAMIENTO DE ARCHIVOS 264
# ============================================================

def procesar_archivo_264(archivo) -> pd.DataFrame:
    """
    Lee y normaliza un archivo 264.
    """
    nombre_archivo = getattr(archivo, "name", str(archivo))
    sede = Path(nombre_archivo).stem

    df = leer_excel_crystal(archivo)

    fecha_col = encontrar_columna(
        df.columns,
        debe_contener=["FECHA"],
        alternativas=["FECHA", "Fecha"]
    )

    doc_col = encontrar_columna(
        df.columns,
        debe_contener=["DOC"],
        no_debe_contener=["CONT"],
        alternativas=[
            "DOC. PCTE.",
            "DOC PCTE",
            "DOC PACIENTE",
            "Doc Paciente",
        ],
    )

    doc_cont_col = encontrar_columna(
        df.columns,
        debe_contener=["DOC", "CONT"],
        alternativas=["DOC. CONT.", "DOC CONTABLE", "Doc Contable"]
    )

    apb_col = encontrar_columna(df.columns, debe_contener=["APB"], alternativas=["APB"])
    convenio_col = encontrar_columna(df.columns, debe_contener=["CONVENIO"], alternativas=["CONVENIO"])
    identificacion_col = encontrar_columna(df.columns, debe_contener=["IDENTIFICACION"], alternativas=["IDENTIFICACION", "Identificacion"])
    paciente_col = encontrar_columna(df.columns, debe_contener=["PACIENTE"], alternativas=["PACIENTE", "Paciente"])
    servicio_col = encontrar_columna(df.columns, debe_contener=["SERVICIO"], alternativas=["SERVICIO", "Servicio"])

    cant_col = encontrar_columna(df.columns, debe_contener=["CANT"], alternativas=["CANT", "Cantidad"])

    vlr_serv_col = encontrar_columna(
        df.columns,
        debe_contener=["VLR", "SERV"],
        alternativas=["VLR  SERV", "VLR SERV", "VALOR SERVICIO", "Valor Servicio"]
    )

    vlr_docto_col = encontrar_columna(
        df.columns,
        debe_contener=["VLR", "DOCTO"],
        alternativas=["VLR DOCTO", "VALOR DOCUMENTO", "Valor Documento"]
    )

    vlr_pcte_col = encontrar_columna(
        df.columns,
        debe_contener=["VLR", "PCTE"],
        alternativas=["VLR PCTE", "VALOR PACIENTE", "Valor Paciente"]
    )

    columnas_minimas = {
        "FECHA": fecha_col,
        "DOC_PACIENTE": doc_col,
        "VLR_DOCTO": vlr_docto_col,
    }

    faltantes = [k for k, v in columnas_minimas.items() if v is None]
    if faltantes:
        raise ValueError(
            f"El archivo 264 {nombre_archivo} no tiene columnas mínimas requeridas: {', '.join(faltantes)}"
        )

    # IMPORTANTE: crear el DataFrame con el mismo índice del archivo leído.
    # Si ARCHIVO_264 y SEDE quedan en NaN, el groupby final excluye todas las filas.
    salida = pd.DataFrame(index=df.index)
    salida["ARCHIVO_264"] = nombre_archivo
    salida["SEDE"] = sede
    salida["FECHA_ORIGINAL_264"] = df[fecha_col].apply(convertir_fecha)
    salida["DOC_PACIENTE_ORIGINAL_264"] = df[doc_col]
    salida["DOC_NORM"] = df[doc_col].apply(normalizar_doc)

    salida["DOC_CONTABLE"] = df[doc_cont_col] if doc_cont_col else ""
    salida["APB"] = df[apb_col] if apb_col else ""
    salida["CONVENIO"] = df[convenio_col] if convenio_col else ""
    salida["IDENTIFICACION"] = df[identificacion_col] if identificacion_col else ""
    salida["PACIENTE"] = df[paciente_col] if paciente_col else ""
    salida["SERVICIO"] = df[servicio_col] if servicio_col else ""

    salida["CANT"] = df[cant_col].apply(limpiar_numero) if cant_col else 0
    salida["VLR_SERV"] = df[vlr_serv_col].apply(limpiar_numero) if vlr_serv_col else 0
    salida["VLR_DOCTO"] = df[vlr_docto_col].apply(limpiar_numero)
    salida["VLR_PCTE"] = df[vlr_pcte_col].apply(limpiar_numero) if vlr_pcte_col else 0

    salida = salida[salida["DOC_NORM"] != ""].copy()
    salida = salida.dropna(how="all").reset_index(drop=True)

    return salida


def consolidar_264(archivos_264) -> pd.DataFrame:
    """
    Consolida todos los 264 cargados.
    """
    lista = []

    for archivo in archivos_264:
        tmp = procesar_archivo_264(archivo)
        if not tmp.empty:
            lista.append(tmp)

    if not lista:
        return pd.DataFrame()

    return pd.concat(lista, ignore_index=True)


# ============================================================
# REPORTE FINAL
# ============================================================

def generar_reporte_final(base_264: pd.DataFrame, fechas_260: pd.DataFrame, mantener_fecha_original_si_falta=True):
    """
    Cruza 264 con fechas 260, reemplaza fechas y agrupa una fila por factura/doc paciente.
    """
    if base_264.empty:
        raise ValueError("La base 264 está vacía.")

    if fechas_260.empty:
        base = base_264.copy()
        base["FECHA_260"] = pd.NaT
        base["ARCHIVOS_260"] = ""
        base["CONFLICTO_FECHA_260"] = "NO"
    else:
        base = base_264.merge(
            fechas_260[["DOC_NORM", "FECHA_260", "ARCHIVOS_260", "CONFLICTO_FECHA_260"]],
            on="DOC_NORM",
            how="left"
        )

    if mantener_fecha_original_si_falta:
        base["FECHA_CORRECTA"] = base["FECHA_260"].combine_first(base["FECHA_ORIGINAL_264"])
    else:
        base["FECHA_CORRECTA"] = base["FECHA_260"]

    base["ESTADO_FECHA"] = base["FECHA_260"].apply(
        lambda x: "REEMPLAZADA_CON_260" if pd.notna(x) else "SIN_FECHA_260"
    )

    # Agrupar una fila por factura/documento dentro de cada archivo/sede.
    agrupado = (
        base.groupby(["ARCHIVO_264", "SEDE", "DOC_NORM"], as_index=False)
        .agg(
            DOC_PACIENTE=("DOC_PACIENTE_ORIGINAL_264", primer_no_vacio),
            FECHA_ORIGINAL_264=("FECHA_ORIGINAL_264", "min"),
            FECHA_CORRECTA=("FECHA_CORRECTA", "min"),
            DOC_CONTABLE=("DOC_CONTABLE", primer_no_vacio),
            APB=("APB", primer_no_vacio),
            CONVENIO=("CONVENIO", primer_no_vacio),
            IDENTIFICACION=("IDENTIFICACION", primer_no_vacio),
            PACIENTE=("PACIENTE", primer_no_vacio),
            VLR_DOCTO=("VLR_DOCTO", "max"),
            VLR_DOCTO_MIN=("VLR_DOCTO", "min"),
            VLR_DOCTO_VALORES_DISTINTOS=("VLR_DOCTO", lambda x: x.nunique(dropna=True)),
            VLR_SERVICIOS_SUM=("VLR_SERV", "sum"),
            VLR_PCTE=("VLR_PCTE", "max"),
            LINEAS_SERVICIO=("SERVICIO", "count"),
            SERVICIOS_DISTINTOS=("SERVICIO", lambda x: x.nunique(dropna=True)),
            ESTADO_FECHA=("ESTADO_FECHA", primer_no_vacio),
            ARCHIVOS_260=("ARCHIVOS_260", primer_no_vacio),
            CONFLICTO_FECHA_260=("CONFLICTO_FECHA_260", primer_no_vacio),
        )
    )

    agrupado["DIFERENCIA_VLR_DOCTO_INTERNA"] = agrupado["VLR_DOCTO"] - agrupado["VLR_DOCTO_MIN"]
    agrupado["ALERTA_VALOR_DOCTO"] = agrupado["VLR_DOCTO_VALORES_DISTINTOS"].apply(
        lambda x: "REVISAR" if x > 1 else "OK"
    )

    columnas_finales = [
        "SEDE",
        "ARCHIVO_264",
        "DOC_PACIENTE",
        "DOC_NORM",
        "FECHA_ORIGINAL_264",
        "FECHA_CORRECTA",
        "ESTADO_FECHA",
        "DOC_CONTABLE",
        "APB",
        "CONVENIO",
        "IDENTIFICACION",
        "PACIENTE",
        "VLR_DOCTO",
        "VLR_SERVICIOS_SUM",
        "VLR_PCTE",
        "LINEAS_SERVICIO",
        "SERVICIOS_DISTINTOS",
        "ARCHIVOS_260",
        "CONFLICTO_FECHA_260",
        "ALERTA_VALOR_DOCTO",
        "DIFERENCIA_VLR_DOCTO_INTERNA",
    ]

    agrupado = agrupado[columnas_finales].copy()

    # Ordenar por sede, fecha y doc
    agrupado = agrupado.sort_values(
        by=["SEDE", "FECHA_CORRECTA", "DOC_PACIENTE"],
        ascending=[True, True, True]
    ).reset_index(drop=True)

    sin_fecha_260 = agrupado[agrupado["ESTADO_FECHA"] == "SIN_FECHA_260"].copy()
    alertas_valor = agrupado[agrupado["ALERTA_VALOR_DOCTO"] == "REVISAR"].copy()

    resumen_general = pd.DataFrame({
        "INDICADOR": [
            "Facturas / Doc Paciente 264",
            "Valor total 264 simplificado",
            "Facturas con fecha reemplazada desde 260",
            "Facturas sin fecha en 260",
            "Facturas con conflicto de fecha 260",
            "Facturas con alerta de valor documento",
        ],
        "VALOR": [
            len(agrupado),
            agrupado["VLR_DOCTO"].sum(),
            (agrupado["ESTADO_FECHA"] == "REEMPLAZADA_CON_260").sum(),
            (agrupado["ESTADO_FECHA"] == "SIN_FECHA_260").sum(),
            (agrupado["CONFLICTO_FECHA_260"] == "SI").sum(),
            (agrupado["ALERTA_VALOR_DOCTO"] == "REVISAR").sum(),
        ]
    })

    resumen_sede = (
        agrupado.groupby("SEDE", as_index=False)
        .agg(
            FACTURAS=("DOC_PACIENTE", "count"),
            VALOR_TOTAL=("VLR_DOCTO", "sum"),
            CON_FECHA_260=("ESTADO_FECHA", lambda x: (x == "REEMPLAZADA_CON_260").sum()),
            SIN_FECHA_260=("ESTADO_FECHA", lambda x: (x == "SIN_FECHA_260").sum()),
        )
        .sort_values("VALOR_TOTAL", ascending=False)
    )

    resumen_convenio = (
        agrupado.groupby(["SEDE", "CONVENIO"], as_index=False)
        .agg(
            FACTURAS=("DOC_PACIENTE", "count"),
            VALOR_TOTAL=("VLR_DOCTO", "sum"),
            CON_FECHA_260=("ESTADO_FECHA", lambda x: (x == "REEMPLAZADA_CON_260").sum()),
            SIN_FECHA_260=("ESTADO_FECHA", lambda x: (x == "SIN_FECHA_260").sum()),
        )
        .sort_values(["SEDE", "VALOR_TOTAL"], ascending=[True, False])
    )

    return {
        "reporte_final": agrupado,
        "sin_fecha_260": sin_fecha_260,
        "alertas_valor": alertas_valor,
        "resumen_general": resumen_general,
        "resumen_sede": resumen_sede,
        "resumen_convenio": resumen_convenio,
    }


def crear_excel_descarga(resultados: dict, conflictos_260: pd.DataFrame, log_archivos: pd.DataFrame) -> bytes:
    """
    Genera archivo Excel con varias hojas.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd hh:mm:ss", date_format="yyyy-mm-dd") as writer:
        workbook = writer.book

        formato_titulo = workbook.add_format({
            "bold": True,
            "font_size": 13,
            "bg_color": "#1F4E78",
            "font_color": "white",
            "border": 1,
        })

        formato_header = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAF7",
            "border": 1,
            "text_wrap": True,
            "valign": "top",
        })

        formato_moneda = workbook.add_format({
            "num_format": '$ #,##0',
            "border": 1,
        })

        formato_fecha = workbook.add_format({
            "num_format": "yyyy-mm-dd hh:mm:ss",
            "border": 1,
        })

        formato_normal = workbook.add_format({
            "border": 1,
        })

        hojas = {
            "REPORTE_FINAL": resultados["reporte_final"],
            "RESUMEN_GENERAL": resultados["resumen_general"],
            "RESUMEN_SEDE": resultados["resumen_sede"],
            "RESUMEN_CONVENIO": resultados["resumen_convenio"],
            "SIN_FECHA_260": resultados["sin_fecha_260"],
            "ALERTAS_VALOR": resultados["alertas_valor"],
            "CONFLICTOS_260": conflictos_260,
            "LOG_ARCHIVOS": log_archivos,
        }

        for nombre_hoja, df in hojas.items():
            df = df.copy()

            # Evitar filas completamente vacías
            df = df.dropna(how="all")

            df.to_excel(writer, sheet_name=nombre_hoja, index=False, startrow=1)

            ws = writer.sheets[nombre_hoja]

            ws.write(0, 0, nombre_hoja, formato_titulo)

            # Encabezados
            for col_num, value in enumerate(df.columns):
                ws.write(1, col_num, value, formato_header)

            # Autofiltro
            if len(df.columns) > 0:
                ws.autofilter(1, 0, max(1, len(df) + 1), len(df.columns) - 1)
                ws.freeze_panes(2, 0)

            # Anchos y formatos
            for idx, col in enumerate(df.columns):
                serie = df[col].astype(str) if not df.empty else pd.Series(dtype=str)
                max_len = max([len(str(col))] + [len(x) for x in serie.head(500).tolist()])
                ancho = min(max(max_len + 2, 12), 45)
                ws.set_column(idx, idx, ancho)

                col_norm = normalizar_texto(col)

                if "VLR" in col_norm or "VALOR" in col_norm or "DIFERENCIA" in col_norm:
                    ws.set_column(idx, idx, max(ancho, 16), formato_moneda)
                elif "FECHA" in col_norm:
                    ws.set_column(idx, idx, max(ancho, 20), formato_fecha)
                else:
                    ws.set_column(idx, idx, ancho, formato_normal)

        # Hoja LEEME
        leeme = workbook.add_worksheet("LEEME")
        leeme.write(0, 0, "Uso del reporte", formato_titulo)
        textos = [
            "1. REPORTE_FINAL contiene una fila por factura / Doc Paciente tomada desde los 264.",
            "2. FECHA_CORRECTA se toma desde los 260 cuando existe cruce por Doc Paciente.",
            "3. Si no existe fecha en 260, el estado queda SIN_FECHA_260 y se conserva la fecha original del 264 si esa opción fue seleccionada.",
            "4. VLR_DOCTO se toma del 264 agrupado por documento. No se suma fila por fila para evitar inflación del valor.",
            "5. SIN_FECHA_260 muestra documentos del 264 que no fueron encontrados en ningún 260.",
            "6. CONFLICTOS_260 muestra Doc Paciente con más de una fecha distinta en los archivos 260.",
            "7. ALERTAS_VALOR muestra documentos donde el mismo Doc Paciente tiene valores de documento diferentes dentro del 264.",
        ]

        for i, texto in enumerate(textos, start=2):
            leeme.write(i, 0, texto)

        leeme.set_column(0, 0, 120)

    output.seek(0)
    return output.getvalue()


def crear_log_archivos(archivos_264, archivos_260) -> pd.DataFrame:
    """
    Crea log de archivos cargados.
    """
    registros = []

    for archivo in archivos_264:
        registros.append({
            "TIPO": "264",
            "ARCHIVO": getattr(archivo, "name", str(archivo)),
            "FECHA_PROCESO": datetime.now(),
        })

    for archivo in archivos_260:
        registros.append({
            "TIPO": "260",
            "ARCHIVO": getattr(archivo, "name", str(archivo)),
            "FECHA_PROCESO": datetime.now(),
        })

    return pd.DataFrame(registros)


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.title("📊 Conciliador de informes 264 vs 260")
st.caption(
    "Carga varios informes 264 y varios informes 260. "
    "La app reemplaza la fecha del 264 usando la fecha del 260 por Doc Paciente "
    "y genera un reporte simplificado por factura."
)

with st.expander("📌 ¿Qué hace esta app?", expanded=True):
    st.markdown(
        """
        **Lógica del proceso:**

        1. Cargas todos los reportes **264** de las sedes.
        2. Cargas todos los reportes **260** disponibles.
        3. La app toma del **260** únicamente:
           - La fecha de la primera columna.
           - El Doc Paciente.
        4. Cruza contra el **264** por Doc Paciente.
        5. Reemplaza la fecha del 264 por la fecha encontrada en el 260.
        6. Agrupa el 264 para dejar **una sola fila por factura / Doc Paciente**.
        7. Genera un Excel final sin filas vacías.
        """
    )

col1, col2 = st.columns(2)

with col1:
    archivos_264 = st.file_uploader(
        "Carga aquí los informes 264 de todas las sedes",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        key="uploader_264"
    )

with col2:
    archivos_260 = st.file_uploader(
        "Carga aquí los informes 260",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        key="uploader_260"
    )

st.sidebar.header("⚙️ Opciones")

mantener_fecha_original = st.sidebar.checkbox(
    "Si no hay fecha en 260, conservar fecha original del 264",
    value=True
)

nombre_salida = st.sidebar.text_input(
    "Nombre del archivo de salida",
    value="reporte_general_simplificado_fechas_correctas.xlsx"
)

procesar = st.button("🚀 Procesar reportes", type="primary")

if procesar:
    if not archivos_264:
        st.error("Debes cargar al menos un archivo 264. Sin base maestra no hay milagro, solo fe administrativa.")
        st.stop()

    if not archivos_260:
        st.warning(
            "No cargaste archivos 260. La app generará el reporte simplificado, "
            "pero no podrá reemplazar fechas."
        )

    try:
        with st.spinner("Procesando archivos..."):
            fechas_260, conflictos_260 = consolidar_fechas_260(archivos_260) if archivos_260 else (pd.DataFrame(), pd.DataFrame())
            base_264 = consolidar_264(archivos_264)

            resultados = generar_reporte_final(
                base_264=base_264,
                fechas_260=fechas_260,
                mantener_fecha_original_si_falta=mantener_fecha_original
            )

            log_archivos = crear_log_archivos(archivos_264, archivos_260)

            excel_bytes = crear_excel_descarga(
                resultados=resultados,
                conflictos_260=conflictos_260,
                log_archivos=log_archivos
            )

        reporte_final = resultados["reporte_final"]

        st.success("Proceso terminado correctamente.")

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)

        kpi1.metric("Facturas / Doc Paciente", f"{len(reporte_final):,}")
        kpi2.metric("Valor total", f"${reporte_final['VLR_DOCTO'].sum():,.0f}")
        kpi3.metric(
            "Con fecha 260",
            f"{(reporte_final['ESTADO_FECHA'] == 'REEMPLAZADA_CON_260').sum():,}"
        )
        kpi4.metric(
            "Sin fecha 260",
            f"{(reporte_final['ESTADO_FECHA'] == 'SIN_FECHA_260').sum():,}"
        )

        st.subheader("Vista previa del reporte final")
        st.dataframe(reporte_final.head(500), use_container_width=True)

        if not resultados["sin_fecha_260"].empty:
            st.warning(
                f"Hay {len(resultados['sin_fecha_260'])} facturas del 264 sin fecha encontrada en los 260."
            )
            with st.expander("Ver facturas sin fecha 260"):
                st.dataframe(resultados["sin_fecha_260"], use_container_width=True)

        if not conflictos_260.empty:
            st.warning(
                f"Hay {len(conflictos_260)} Doc Paciente con más de una fecha distinta en los 260."
            )
            with st.expander("Ver conflictos de fechas en 260"):
                st.dataframe(conflictos_260, use_container_width=True)

        if not resultados["alertas_valor"].empty:
            st.warning(
                f"Hay {len(resultados['alertas_valor'])} documentos con valores distintos dentro del 264. Revisar ALERTAS_VALOR."
            )

        st.download_button(
            label="⬇️ Descargar Excel final",
            data=excel_bytes,
            file_name=nombre_salida if nombre_salida.endswith(".xlsx") else f"{nombre_salida}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("Ocurrió un error procesando los archivos.")
        st.exception(e)

else:
    st.info("Carga los archivos y presiona **Procesar reportes**.")
