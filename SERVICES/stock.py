import os
import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd


CARPETA_STOCK = "ARCHIVOS/STOCK"


# ============================================================
# BUSCAR ÚLTIMO STOCK
# ============================================================

def buscar_ultimo_stock():

    archivos = []

    if not os.path.exists(CARPETA_STOCK):
        raise FileNotFoundError(
            f"No existe la carpeta {CARPETA_STOCK}."
        )

    for archivo in os.listdir(CARPETA_STOCK):

        ruta = os.path.join(
            CARPETA_STOCK,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        nombre = archivo.lower().strip()

        if nombre.endswith(".xlsx") and "stock" in nombre:
            archivos.append(ruta)

    if not archivos:
        raise FileNotFoundError(
            "No se encontró ningún archivo STOCK."
        )

    return max(
        archivos,
        key=os.path.getmtime
    )


# ============================================================
# NORMALIZAR DNI
# ============================================================

def normalizar_dni(valor):

    if valor is None:
        return ""

    texto = str(valor).strip()

    if texto.endswith(".0"):
        texto = texto[:-2]

    texto = "".join(
        caracter
        for caracter in texto
        if caracter.isdigit()
    )

    return texto


# ============================================================
# CONVERTIR FECHA
# ============================================================

def convertir_fecha(valor):

    if valor is None:
        return pd.NaT

    return pd.to_datetime(
        valor,
        dayfirst=True,
        errors="coerce"
    )


# ============================================================
# CONVERTIR CELDA EXCEL
# ============================================================

def convertir_numero_excel(valor):

    if valor is None:
        return ""

    texto = str(valor).strip()

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


# ============================================================
# LEER SHARED STRINGS
#
# IMPORTANTE:
# Los XLSX pueden guardar textos en sharedStrings.xml.
#
# Los cargamos solamente si existen.
# ============================================================

def cargar_shared_strings(zf):

    try:
        archivo = "xl/sharedStrings.xml"

        if archivo not in zf.namelist():
            return None

        shared_strings = []

        with zf.open(archivo) as archivo_xml:

            for evento, elemento in ET.iterparse(
                archivo_xml,
                events=("end",)
            ):

                if elemento.tag.endswith("}si"):

                    textos = []

                    for nodo in elemento.iter():

                        if nodo.tag.endswith("}t"):

                            if nodo.text:
                                textos.append(
                                    nodo.text
                                )

                    shared_strings.append(
                        "".join(textos)
                    )

                    elemento.clear()

        return shared_strings

    except Exception:

        return None


# ============================================================
# OBTENER VALOR DE CELDA
# ============================================================

def obtener_valor_celda(
    celda,
    shared_strings
):

    tipo = celda.attrib.get("t")

    valor = None

    for hijo in celda:

        if hijo.tag.endswith("}v"):

            valor = hijo.text
            break

        if hijo.tag.endswith("}is"):

            textos = []

            for nodo in hijo.iter():

                if nodo.tag.endswith("}t"):

                    if nodo.text:
                        textos.append(
                            nodo.text
                        )

            valor = "".join(textos)
            break

    if valor is None:
        return ""

    # --------------------------------------------------------
    # SHARED STRING
    # --------------------------------------------------------

    if tipo == "s":

        try:

            indice = int(valor)

            if (
                shared_strings is not None
                and 0 <= indice < len(shared_strings)
            ):
                return shared_strings[indice]

        except Exception:
            pass

        return ""

    # --------------------------------------------------------
    # BOOLEAN
    # --------------------------------------------------------

    if tipo == "b":

        return "1" if valor == "1" else "0"

    return valor


# ============================================================
# CONVERTIR FECHA SERIAL DE EXCEL
# ============================================================

def convertir_fecha_excel(valor):

    if valor is None or valor == "":
        return pd.NaT

    # --------------------------------------------------------
    # Si ya es una fecha textual
    # --------------------------------------------------------

    texto = str(valor).strip()

    # --------------------------------------------------------
    # Número serial de Excel
    # --------------------------------------------------------

    try:

        numero = float(texto)

        if numero > 0:

            return (
                pd.Timestamp("1899-12-30")
                + pd.to_timedelta(
                    numero,
                    unit="D"
                )
            )

    except Exception:
        pass

    # --------------------------------------------------------
    # Fecha normal
    # --------------------------------------------------------

    return pd.to_datetime(
        texto,
        dayfirst=True,
        errors="coerce"
    )


# ============================================================
# OBTENER LETRA DE COLUMNA
# ============================================================

def obtener_referencia_columna(referencia):

    match = re.match(
        r"([A-Z]+)",
        referencia.upper()
    )

    if not match:
        return ""

    return match.group(1)


# ============================================================
# LEER STOCK DIRECTAMENTE DESDE XLSX
#
# NO USAMOS:
#
# - pandas.read_excel
# - openpyxl
# - calamine
#
# Leemos directamente:
#
# XLSX -> ZIP -> XML
#
# y procesamos fila por fila.
# ============================================================

def cargar_stock():

    ruta = buscar_ultimo_stock()

    print()
    print("========================================")
    print("STOCK ENCONTRADO")
    print("========================================")

    print(ruta)

    print()
    print("Leyendo STOCK en modo XML streaming...")

    # ========================================================
    # ABRIR XLSX COMO ZIP
    # ========================================================

    with zipfile.ZipFile(
        ruta,
        "r"
    ) as zf:

        # ====================================================
        # SHARED STRINGS
        # ====================================================

        shared_strings = cargar_shared_strings(
            zf
        )

        # ====================================================
        # BUSCAR HOJA PRINCIPAL
        #
        # En la mayoría de los stocks es sheet1.xml.
        # ====================================================

        hojas = [
            nombre
            for nombre in zf.namelist()
            if nombre.startswith(
                "xl/worksheets/"
            )
            and nombre.endswith(".xml")
        ]

        if not hojas:

            raise FileNotFoundError(
                "No se encontró ninguna hoja XML dentro del STOCK."
            )

        hoja = hojas[0]

        print()
        print(
            "Hoja utilizada:",
            hoja
        )

        # ====================================================
        # VARIABLES
        # ====================================================

        indice_dni = None
        indice_fecha = None

        stock_dict = {}

        contador = 0

        encabezados_encontrados = False

        # ====================================================
        # ABRIR XML EN STREAMING
        # ====================================================

        with zf.open(hoja) as archivo_xml:

            for evento, fila_xml in ET.iterparse(
                archivo_xml,
                events=("end",)
            ):

                # ------------------------------------------------
                # SOLO PROCESAMOS FILAS
                # ------------------------------------------------

                if not fila_xml.tag.endswith("}row"):
                    continue

                # =================================================
                # ENCABEZADOS
                # =================================================

                if not encabezados_encontrados:

                    encabezados = {}

                    for celda in fila_xml:

                        if not celda.tag.endswith("}c"):
                            continue

                        referencia = celda.attrib.get(
                            "r",
                            ""
                        )

                        columna = obtener_referencia_columna(
                            referencia
                        )

                        valor = obtener_valor_celda(
                            celda,
                            shared_strings
                        )

                        nombre = str(
                            valor
                        ).strip()

                        encabezados[
                            columna
                        ] = nombre

                    # --------------------------------------------
                    # BUSCAR COLUMNAS
                    # --------------------------------------------

                    for columna, nombre in encabezados.items():

                        nombre_normalizado = (
                            nombre
                            .strip()
                            .upper()
                        )

                        if nombre_normalizado == "NUM_DOC":

                            indice_dni = columna

                        elif (
                            nombre_normalizado
                            == "FECHA_ASIG_ESTUDIO"
                        ):

                            indice_fecha = columna

                    if indice_dni is None:

                        raise KeyError(
                            "No se encontró la columna "
                            "'NUM_DOC' en STOCK."
                        )

                    if indice_fecha is None:

                        raise KeyError(
                            "No se encontró la columna "
                            "'FECHA_ASIG_ESTUDIO' en STOCK."
                        )

                    print()
                    print(
                        "Columnas encontradas:"
                    )

                    print(
                        "- NUM_DOC"
                    )

                    print(
                        "- FECHA_ASIG_ESTUDIO"
                    )

                    encabezados_encontrados = True

                    fila_xml.clear()

                    continue

                # =================================================
                # PROCESAR REGISTRO
                # =================================================

                valores = {}

                for celda in fila_xml:

                    if not celda.tag.endswith("}c"):
                        continue

                    referencia = celda.attrib.get(
                        "r",
                        ""
                    )

                    columna = obtener_referencia_columna(
                        referencia
                    )

                    if columna not in (
                        indice_dni,
                        indice_fecha
                    ):
                        continue

                    valores[
                        columna
                    ] = obtener_valor_celda(
                        celda,
                        shared_strings
                    )

                # -------------------------------------------------
                # DNI
                # -------------------------------------------------

                dni = normalizar_dni(
                    valores.get(
                        indice_dni,
                        ""
                    )
                )

                if not dni:

                    fila_xml.clear()

                    continue

                # -------------------------------------------------
                # FECHA
                # -------------------------------------------------

                fecha_original = valores.get(
                    indice_fecha,
                    ""
                )

                fecha = convertir_fecha_excel(
                    fecha_original
                )

                # =================================================
                # CONSERVAR SOLO FECHA MÁS RECIENTE
                # =================================================

                if dni not in stock_dict:

                    stock_dict[dni] = (
                        fecha_original,
                        fecha
                    )

                else:

                    fecha_actual = (
                        stock_dict[dni][1]
                    )

                    if (
                        pd.notna(fecha)
                        and (
                            pd.isna(
                                fecha_actual
                            )
                            or
                            fecha > fecha_actual
                        )
                    ):

                        stock_dict[dni] = (
                            fecha_original,
                            fecha
                        )

                contador += 1

                # -------------------------------------------------
                # PROGRESO
                # -------------------------------------------------

                if contador % 100000 == 0:

                    print(
                        "Registros procesados:",
                        contador,
                        "| DNI únicos:",
                        len(stock_dict)
                    )

                # -------------------------------------------------
                # LIBERAR MEMORIA
                # -------------------------------------------------

                fila_xml.clear()

        # ========================================================
        # RESULTADOS
        # ========================================================

        print()
        print(
            "Registros procesados:",
            contador
        )

        print(
            "DNI únicos:",
            len(stock_dict)
        )

    # ============================================================
    # CREAR DATAFRAME PEQUEÑO
    # ============================================================

    documentos = []
    fechas = []

    for dni, valores in stock_dict.items():

        documentos.append(
            dni
        )

        fechas.append(
            valores[0]
        )

    df = pd.DataFrame({
        "NUM_DOC": documentos,
        "FECHA_ASIG_ESTUDIO": fechas
    })

    # ============================================================
    # RESULTADO
    # ============================================================

    print()
    print("========================================")
    print("STOCK PROCESADO")
    print("========================================")

    print(
        "Registros finales:",
        len(df)
    )

    return df


# ============================================================
# PREPARAR STOCK
# ============================================================

def preparar_stock(df_stock):

    columnas_necesarias = [
        "NUM_DOC",
        "FECHA_ASIG_ESTUDIO"
    ]

    for columna in columnas_necesarias:

        if columna not in df_stock.columns:

            raise KeyError(
                f"No se encontró la columna "
                f"'{columna}' en STOCK."
            )

    stock = df_stock[
        columnas_necesarias
    ].copy()

    stock["NUM_DOC"] = (
        stock["NUM_DOC"]
        .astype("string")
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
    )

    return stock