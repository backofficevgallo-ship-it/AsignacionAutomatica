import os
import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd


CARPETA_STOCK = "ARCHIVOS/STOCK"

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


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

    if valor is None or valor == "":
        return 0

    texto = str(valor).strip()

    if not texto:
        return 0

    # ========================================================
    # FECHA SERIAL DE EXCEL
    #
    # La mayoría de las fechas del STOCK vienen como número.
    # No usamos pandas para convertirlas.
    # ========================================================

    try:

        numero = float(texto)

        if numero > 0:
            return numero

    except (ValueError, TypeError):
        pass

    # ========================================================
    # FECHA TEXTUAL
    #
    # Solo usamos pandas si realmente no es un número.
    # ========================================================

    try:

        fecha = pd.to_datetime(
            texto,
            dayfirst=True,
            errors="coerce"
        )

        if pd.isna(fecha):
            return 0

        # Convertimos la fecha a número comparable
        return (
            fecha - pd.Timestamp("1899-12-30")
        ).total_seconds() / 86400

    except Exception:

        return 0


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
# LEER VALOR CRUDO DE CELDA
#
# IMPORTANTE:
# Si es shared string devolvemos:
#
# ("SHARED", indice)
#
# No cargamos sharedStrings completo.
# ============================================================

def obtener_valor_crudo(celda):

    tipo = celda.attrib.get("t")

    nodo_v = celda.find(
        f"{{{NS_MAIN}}}v"
    )

    if nodo_v is not None:

        valor = nodo_v.text or ""

        if tipo == "s":

            try:

                return (
                    "SHARED",
                    int(valor)
                )

            except Exception:

                return (
                    "SHARED",
                    -1
                )

        if tipo == "b":

            return (
                "1"
                if valor == "1"
                else "0"
            )

        return valor

    # ========================================================
    # INLINE STRING
    # ========================================================

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
# OBTENER VALOR DE SHARED STRING ESPECÍFICO
#
# Se usa solamente para encabezados.
# ============================================================

def buscar_shared_string(
    zf,
    indice_buscado
):

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

                return "".join(textos)

            indice_actual += 1

            elemento.clear()

    return ""


# ============================================================
# OBTENER HOJA PRINCIPAL
# ============================================================

def buscar_hoja(zf):

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

    return hojas[0]


# ============================================================
# LEER ENCABEZADOS
# ============================================================

def leer_encabezados(
    zf,
    hoja
):

    with zf.open(hoja) as archivo_xml:

        for evento, elemento in ET.iterparse(
            archivo_xml,
            events=("end",)
        ):

            if not elemento.tag.endswith("}row"):
                continue

            encabezados = {}

            for celda in elemento:

                if not celda.tag.endswith("}c"):
                    continue

                referencia = celda.attrib.get(
                    "r",
                    ""
                )

                columna = obtener_referencia_columna(
                    referencia
                )

                valor = obtener_valor_crudo(
                    celda
                )

                # --------------------------------------------
                # RESOLVER SHARED STRING DEL ENCABEZADO
                # --------------------------------------------

                if (
                    isinstance(valor, tuple)
                    and valor[0] == "SHARED"
                ):

                    valor = buscar_shared_string(
                        zf,
                        valor[1]
                    )

                encabezados[columna] = (
                    str(valor)
                    .strip()
                    .upper()
                )

            elemento.clear()

            return encabezados

    return {}


# ============================================================
# PRIMERA PASADA:
# DETECTAR COLUMNAS SHARED
#
# Solo guardamos los índices de shared strings que realmente
# aparecen en las dos columnas que necesitamos.
# ============================================================

def obtener_indices_shared_necesarios(
    zf,
    hoja,
    indice_dni,
    indice_fecha
):

    indices = set()

    with zf.open(hoja) as archivo_xml:

        for evento, elemento in ET.iterparse(
            archivo_xml,
            events=("end",)
        ):

            if not elemento.tag.endswith("}row"):
                continue

            for celda in elemento:

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

                valor = obtener_valor_crudo(
                    celda
                )

                if (
                    isinstance(valor, tuple)
                    and valor[0] == "SHARED"
                ):

                    if valor[1] >= 0:
                        indices.add(
                            valor[1]
                        )

            elemento.clear()

    return indices


# ============================================================
# CARGAR SOLO SHARED STRINGS NECESARIOS
#
# NO cargamos todo sharedStrings.xml.
# Solo guardamos los índices que realmente aparecen en
# NUM_DOC / FECHA_ASIG_ESTUDIO.
# ============================================================

def cargar_shared_strings_necesarios(
    zf,
    indices_necesarios
):

    resultado = {}

    if not indices_necesarios:
        return resultado

    archivo = "xl/sharedStrings.xml"

    if archivo not in zf.namelist():
        return resultado

    indice_actual = 0

    with zf.open(archivo) as archivo_xml:

        for evento, elemento in ET.iterparse(
            archivo_xml,
            events=("end",)
        ):

            if not elemento.tag.endswith("}si"):
                continue

            if indice_actual in indices_necesarios:

                textos = []

                for nodo in elemento.iter():

                    if nodo.tag.endswith("}t"):

                        if nodo.text:
                            textos.append(
                                nodo.text
                            )

                resultado[indice_actual] = (
                    "".join(textos)
                )

            indice_actual += 1

            elemento.clear()

            # -----------------------------------------------
            # Si ya encontramos todos, podemos terminar.
            # -----------------------------------------------

            if len(resultado) == len(
                indices_necesarios
            ):
                break

    return resultado


# ============================================================
# RESOLVER VALOR
# ============================================================

def resolver_valor(
    valor,
    shared_strings
):

    if (
        isinstance(valor, tuple)
        and valor[0] == "SHARED"
    ):

        return shared_strings.get(
            valor[1],
            ""
        )

    return valor


# ============================================================
# CARGAR STOCK
#
# ESTRATEGIA:
#
# 1. Abrir XLSX.
# 2. Leer encabezados.
# 3. Detectar shared strings necesarios.
# 4. Cargar solamente esos shared strings.
# 5. Recorrer nuevamente las filas.
#
# No se carga el STOCK completo en pandas.
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

    stock_dict = {}

    contador = 0

    with zipfile.ZipFile(
        ruta,
        "r"
    ) as zf:

        # ====================================================
        # HOJA
        # ====================================================

        hoja = buscar_hoja(zf)

        print()
        print(
            "Hoja utilizada:",
            hoja
        )

        # ====================================================
        # ENCABEZADOS
        # ====================================================

        encabezados = leer_encabezados(
            zf,
            hoja
        )

        indice_dni = None
        indice_fecha = None

        for columna, nombre in encabezados.items():

            if nombre == "NUM_DOC":

                indice_dni = columna

            elif nombre == "FECHA_ASIG_ESTUDIO":

                indice_fecha = columna

        if indice_dni is None:

            raise KeyError(
                "No se encontró la columna NUM_DOC en STOCK."
            )

        if indice_fecha is None:

            raise KeyError(
                "No se encontró la columna "
                "FECHA_ASIG_ESTUDIO en STOCK."
            )

        print()
        print(
            "Columnas encontradas:"
        )

        print("- NUM_DOC")
        print("- FECHA_ASIG_ESTUDIO")

        # ====================================================
        # DETECTAR SHARED STRINGS NECESARIOS
        # ====================================================

        print()
        print(
            "Analizando valores necesarios del STOCK..."
        )

        indices_shared = (
            obtener_indices_shared_necesarios(
                zf,
                hoja,
                indice_dni,
                indice_fecha
            )
        )

        print(
            "Shared strings necesarios:",
            len(indices_shared)
        )

        # ====================================================
        # CARGAR SOLO LOS SHARED NECESARIOS
        # ====================================================

        shared_strings = (
            cargar_shared_strings_necesarios(
                zf,
                indices_shared
            )
        )

        print(
            "Shared strings cargados:",
            len(shared_strings)
        )

        # ====================================================
        # SEGUNDA PASADA
        # PROCESAR STOCK
        # ====================================================

        with zf.open(hoja) as archivo_xml:

            for evento, elemento in ET.iterparse(
                archivo_xml,
                events=("end",)
            ):

                if not elemento.tag.endswith("}row"):
                    continue

                # ------------------------------------------------
                # SALTAR ENCABEZADO
                # ------------------------------------------------

                if contador == 0:

                    # No usamos contador para detectar realmente
                    # el encabezado porque puede haber filas vacías.
                    pass

                dni = ""
                fecha_original = ""

                for celda in elemento:

                    if not celda.tag.endswith("}c"):
                        continue

                    referencia = celda.attrib.get(
                        "r",
                        ""
                    )

                    columna = obtener_referencia_columna(
                        referencia
                    )

                    if columna == indice_dni:

                        valor = obtener_valor_crudo(
                            celda
                        )

                        valor = resolver_valor(
                            valor,
                            shared_strings
                        )

                        dni = normalizar_dni(
                            valor
                        )

                    elif columna == indice_fecha:

                        valor = obtener_valor_crudo(
                            celda
                        )

                        valor = resolver_valor(
                            valor,
                            shared_strings
                        )

                        fecha_original = valor

                # ------------------------------------------------
                # EVITAR PROCESAR EL ENCABEZADO
                # ------------------------------------------------

                if dni.upper() == "NUM_DOC":

                    elemento.clear()
                    continue

                # ------------------------------------------------
                # DNI VACÍO
                # ------------------------------------------------

                if not dni:

                    elemento.clear()
                    continue

                # ------------------------------------------------
                # FECHA
                # ------------------------------------------------

                fecha_nueva = convertir_fecha_excel(
                    fecha_original
                )

                # =================================================
                # # CONSERVAR FECHA MÁS RECIENTE
                # # =================================================
                
                anterior = stock_dict.get(
                    dni
                    )
                
                if anterior is None:
                    stock_dict[dni] = (
                        fecha_original,
                        fecha_nueva
                    )
                else:

                    fecha_actual = anterior[1]
                    if fecha_nueva > fecha_actual:

                        stock_dict[dni] = (
                            fecha_original,
                            fecha_nueva
                            )

                contador += 1

                if contador % 5000 == 0:

                    print(
                        "Registros procesados:",
                        contador,
                        "| DNI únicos:",
                        len(stock_dict)
                    )

                # ------------------------------------------------
                # LIBERAR XML
                # ------------------------------------------------

                elemento.clear()

    # ============================================================
    # RESULTADOS
    # ============================================================

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
    # DATAFRAME FINAL
    # ============================================================

    documentos = []
    fechas = []

    for dni, valores in stock_dict.items():

        documentos.append(dni)

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