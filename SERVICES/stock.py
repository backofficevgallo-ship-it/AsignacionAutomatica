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

    return "".join(
        caracter
        for caracter in texto
        if caracter.isdigit()
    )


# ============================================================
# CONVERTIR FECHA EXCEL
# ============================================================

def convertir_fecha_excel(valor):

    if valor is None:
        return pd.NaT

    texto = str(valor).strip()

    if not texto:
        return pd.NaT

    # --------------------------------------------------------
    # FECHA SERIAL DE EXCEL
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
    # FECHA NORMAL
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
# NAMESPACE EXCEL
# ============================================================

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


# ============================================================
# OBTENER VALOR SIMPLE DE CELDA
#
# IMPORTANTE:
# NO CARGAMOS sharedStrings.xml ENTERO.
#
# Si una celda es shared string, se devuelve el índice.
# ============================================================

def obtener_valor_celda(celda):

    tipo = celda.attrib.get("t")

    # --------------------------------------------------------
    # CELDA NORMAL
    # --------------------------------------------------------

    nodo_v = celda.find(
        f"{{{NS_MAIN}}}v"
    )

    if nodo_v is not None:

        valor = nodo_v.text or ""

        # ----------------------------------------------------
        # SHARED STRING
        #
        # Devolvemos el índice.
        # No cargamos todos los textos.
        # ----------------------------------------------------

        if tipo == "s":

            try:
                return ("SHARED", int(valor))
            except Exception:
                return ("SHARED", -1)

        # ----------------------------------------------------
        # BOOLEAN
        # ----------------------------------------------------

        if tipo == "b":

            return "1" if valor == "1" else "0"

        return valor

    # --------------------------------------------------------
    # INLINE STRING
    # --------------------------------------------------------

    nodo_is = celda.find(
        f"{{{NS_MAIN}}}is"
    )

    if nodo_is is not None:

        textos = []

        for nodo in nodo_is.iter():

            if nodo.tag.endswith("}t"):

                if nodo.text:
                    textos.append(
                        nodo.text
                    )

        return "".join(textos)

    return ""


# ============================================================
# RESOLVER SHARED STRING
#
# IMPORTANTE:
# SOLO SE BUSCA EL TEXTO NECESARIO.
#
# No se carga sharedStrings.xml completo.
# ============================================================

def buscar_shared_string(zf, indice_buscado):

    if indice_buscado < 0:
        return ""

    archivo = "xl/sharedStrings.xml"

    if archivo not in zf.namelist():
        return ""

    indice_actual = 0

    with zf.open(archivo) as archivo_xml:

        for evento, elemento in ET.iterparse(
            archivo_xml,
            events=("end",)
        ):

            if not elemento.tag.endswith("}si"):
                continue

            if indice_actual == indice_buscado:

                textos = []

                for nodo in elemento.iter():

                    if nodo.tag.endswith("}t"):

                        if nodo.text:
                            textos.append(
                                nodo.text
                            )

                elemento.clear()

                return "".join(textos)

            indice_actual += 1
            elemento.clear()

    return ""


# ============================================================
# LEER STOCK
# ============================================================

def cargar_stock():

    ruta = buscar_ultimo_stock()

    print()
    print("========================================")
    print("STOCK ENCONTRADO")
    print("========================================")

    print(ruta)

    print()
    print("Leyendo STOCK con XML de bajo consumo...")

    # ========================================================
    # ABRIR XLSX
    # ========================================================

    with zipfile.ZipFile(
        ruta,
        "r"
    ) as zf:

        # ====================================================
        # BUSCAR HOJA
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

        columna_dni = None
        columna_fecha = None

        encabezados_encontrados = False

        # ====================================================
        # PRIMER PASO
        #
        # LEER ENCABEZADOS Y DETERMINAR COLUMNAS
        # ====================================================

        with zf.open(hoja) as archivo_xml:

            for evento, fila_xml in ET.iterparse(
                archivo_xml,
                events=("end",)
            ):

                if not fila_xml.tag.endswith("}row"):
                    continue

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
                        celda
                    )

                    # ------------------------------------------------
                    # Para encabezados normalmente no hace falta
                    # resolver shared strings, pero por seguridad
                    # lo hacemos solamente acá.
                    # ------------------------------------------------

                    if (
                        isinstance(valor, tuple)
                        and valor[0] == "SHARED"
                    ):

                        valor = buscar_shared_string(
                            zf,
                            valor[1]
                        )

                    encabezados[
                        columna
                    ] = str(
                        valor
                    ).strip()

                for columna, nombre in encabezados.items():

                    nombre_normalizado = (
                        nombre
                        .strip()
                        .upper()
                    )

                    if nombre_normalizado == "NUM_DOC":

                        columna_dni = columna

                    elif (
                        nombre_normalizado
                        == "FECHA_ASIG_ESTUDIO"
                    ):

                        columna_fecha = columna

                fila_xml.clear()

                break

        # ====================================================
        # VALIDAR COLUMNAS
        # ====================================================

        if columna_dni is None:

            raise KeyError(
                "No se encontró la columna "
                "'NUM_DOC' en STOCK."
            )

        if columna_fecha is None:

            raise KeyError(
                "No se encontró la columna "
                "'FECHA_ASIG_ESTUDIO' en STOCK."
            )

        print()
        print("Columnas encontradas:")
        print("- NUM_DOC")
        print("- FECHA_ASIG_ESTUDIO")

        # ====================================================
        # CACHE DE SHARED STRINGS
        #
        # SOLO GUARDAMOS LOS STRINGS QUE REALMENTE APARECEN
        # EN LAS DOS COLUMNAS QUE UTILIZAMOS.
        # ====================================================

        shared_cache = {}

        # ====================================================
        # DICCIONARIO FINAL
        #
        # DNI -> fecha más reciente
        # ====================================================

        stock_dict = {}

        contador = 0

        # ====================================================
        # PROCESAR STOCK
        # ====================================================

        with zf.open(hoja) as archivo_xml:

            primera_fila = True

            for evento, fila_xml in ET.iterparse(
                archivo_xml,
                events=("end",)
            ):

                if not fila_xml.tag.endswith("}row"):
                    continue

                # ------------------------------------------------
                # SALTAR ENCABEZADOS
                # ------------------------------------------------

                if primera_fila:

                    primera_fila = False
                    fila_xml.clear()
                    continue

                valor_dni = ""
                valor_fecha = ""

                # =================================================
                # SOLO LEER LAS DOS COLUMNAS NECESARIAS
                # =================================================

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
                        columna_dni,
                        columna_fecha
                    ):
                        continue

                    valor = obtener_valor_celda(
                        celda
                    )

                    # ------------------------------------------------
                    # SHARED STRING
                    # ------------------------------------------------

                    if (
                        isinstance(valor, tuple)
                        and valor[0] == "SHARED"
                    ):

                        indice = valor[1]

                        if indice in shared_cache:

                            valor = shared_cache[
                                indice
                            ]

                        else:

                            valor = buscar_shared_string(
                                zf,
                                indice
                            )

                            shared_cache[
                                indice
                            ] = valor

                    # ------------------------------------------------
                    # GUARDAR VALOR
                    # ------------------------------------------------

                    if columna == columna_dni:

                        valor_dni = valor

                    elif columna == columna_fecha:

                        valor_fecha = valor

                # =================================================
                # NORMALIZAR DNI
                # =================================================

                dni = normalizar_dni(
                    valor_dni
                )

                if not dni:

                    fila_xml.clear()
                    continue

                # =================================================
                # FECHA
                # =================================================

                fecha = convertir_fecha_excel(
                    valor_fecha
                )

                # =================================================
                # CONSERVAR FECHA MÁS RECIENTE
                # =================================================

                if dni not in stock_dict:

                    stock_dict[dni] = (
                        valor_fecha,
                        fecha
                    )

                else:

                    fecha_actual = (
                        stock_dict[dni][1]
                    )

                    if (
                        pd.notna(fecha)
                        and (
                            pd.isna(fecha_actual)
                            or fecha > fecha_actual
                        )
                    ):

                        stock_dict[dni] = (
                            valor_fecha,
                            fecha
                        )

                contador += 1

                # =================================================
                # PROGRESO
                # =================================================

                if contador % 5000 == 0:

                    print(
                        "Registros procesados:",
                        contador,
                        "| DNI únicos:",
                        len(stock_dict)
                    )

                # =================================================
                # LIBERAR XML
                # =================================================

                fila_xml.clear()

        # ========================================================
        # RESULTADO
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
    # CREAR DATAFRAME FINAL
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