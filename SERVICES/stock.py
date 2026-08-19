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

# ============================================================
# CONVERTIR FECHA EXCEL - VERSION RAPIDA
# ============================================================

def convertir_fecha_excel(valor):

    if valor is None or valor == "":
        return None

    texto = str(valor).strip()

    if not texto:
        return None

    # --------------------------------------------------------
    # FECHA SERIAL DE EXCEL
    # --------------------------------------------------------

    try:

        numero = float(texto)

        if numero > 0:

            return (
                pd.Timestamp("1899-12-30")
                + pd.Timedelta(days=numero)
            )

    except Exception:
        pass

    # --------------------------------------------------------
    # FECHA TEXTUAL
    # --------------------------------------------------------

    try:

        return pd.to_datetime(
            texto,
            dayfirst=True,
            errors="coerce"
        )

    except Exception:

        return None

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

def cargar_shared_strings(zf):

    archivo = "xl/sharedStrings.xml"

    if archivo not in zf.namelist():
        return None

    shared_strings = []

    with zf.open(archivo) as archivo_xml:

        for evento, elemento in ET.iterparse(
            archivo_xml,
            events=("end",)
        ):

            if not elemento.tag.endswith("}si"):
                continue

            textos = []

            for nodo in elemento.iter():

                if nodo.tag.endswith("}t"):

                    if nodo.text:
                        textos.append(nodo.text)

            shared_strings.append("".join(textos))

            elemento.clear()

    return shared_strings

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
# LEER STOCK
# ============================================================

# ============================================================
# CARGAR STOCK
#
# XML STREAMING + BAJO CONSUMO
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

    with zipfile.ZipFile(ruta, "r") as zf:

        # ========================================================
        # SHARED STRINGS
        # ========================================================

        shared_strings = cargar_shared_strings(zf)

        # ========================================================
        # BUSCAR HOJA
        # ========================================================

        hojas = [
            nombre
            for nombre in zf.namelist()
            if nombre.startswith("xl/worksheets/")
            and nombre.endswith(".xml")
        ]

        if not hojas:
            raise FileNotFoundError(
                "No se encontró ninguna hoja XML dentro del STOCK."
            )

        hoja = hojas[0]

        print()
        print("Hoja utilizada:", hoja)

        indice_dni = None
        indice_fecha = None

        encabezados_encontrados = False

        # ========================================================
        # XML STREAMING
        # ========================================================

        with zf.open(hoja) as archivo_xml:

            contexto = ET.iterparse(
                archivo_xml,
                events=("start", "end")
            )

            for evento, elemento in contexto:

                # ------------------------------------------------
                # SOLO NOS INTERESAN LOS ROW
                # ------------------------------------------------

                if evento != "end":
                    continue

                if not elemento.tag.endswith("}row"):
                    continue

                # =================================================
                # ENCABEZADOS
                # =================================================

                if not encabezados_encontrados:

                    encabezados = {}

                    for celda in elemento:

                        if not celda.tag.endswith("}c"):
                            continue

                        referencia = celda.attrib.get("r", "")

                        columna = obtener_referencia_columna(
                            referencia
                        )

                        valor = obtener_valor_celda(
                            celda,
                            shared_strings
                        )

                        nombre = str(valor).strip().upper()

                        encabezados[columna] = nombre

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
                            "No se encontró la columna FECHA_ASIG_ESTUDIO en STOCK."
                        )

                    print()
                    print("Columnas encontradas:")
                    print("- NUM_DOC")
                    print("- FECHA_ASIG_ESTUDIO")

                    encabezados_encontrados = True

                    elemento.clear()

                    continue

                # =================================================
                # LEER SOLO LAS DOS COLUMNAS NECESARIAS
                # =================================================

                dni = ""
                fecha_original = ""

                for celda in elemento:

                    if not celda.tag.endswith("}c"):
                        continue

                    referencia = celda.attrib.get("r", "")

                    columna = obtener_referencia_columna(
                        referencia
                    )

                    if columna == indice_dni:

                        dni = normalizar_dni(
                            obtener_valor_celda(
                                celda,
                                shared_strings
                            )
                        )

                    elif columna == indice_fecha:

                        fecha_original = obtener_valor_celda(
                            celda,
                            shared_strings
                        )

                # ------------------------------------------------
                # IGNORAR DNI VACÍO
                # ------------------------------------------------

                if not dni:

                    elemento.clear()
                    continue

                # =================================================
                # COMPARAR FECHAS
                #
                # Guardamos solamente la fecha textual y un
                # timestamp temporal para comparar.
                # =================================================

                fecha_nueva = convertir_fecha_excel(
                    fecha_original
                )

                if dni not in stock_dict:

                    stock_dict[dni] = (
                        fecha_original,
                        fecha_nueva
                    )

                else:

                    fecha_actual = stock_dict[dni][1]

                    if (
                        pd.notna(fecha_nueva)
                        and (
                            pd.isna(fecha_actual)
                            or fecha_nueva > fecha_actual
                        )
                    ):

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
                # MUY IMPORTANTE
                # ------------------------------------------------

                elemento.clear()

        # ========================================================
        # LIBERAR SHARED STRINGS
        # ========================================================

        shared_strings = None

    print()
    print("Registros procesados:", contador)
    print("DNI únicos:", len(stock_dict))

    # ============================================================
    # CREAR DATAFRAME
    # ============================================================

    documentos = []
    fechas = []

    for dni, valores in stock_dict.items():

        documentos.append(dni)
        fechas.append(valores[0])

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