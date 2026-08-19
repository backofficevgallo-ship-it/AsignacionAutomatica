import os
import re
import zipfile
import xml.etree.ElementTree as ET
import datetime

import pandas as pd


CARPETA_STOCK = "ARCHIVOS/STOCK"

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


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

        if (
            os.path.isfile(ruta)
            and archivo.lower().endswith(".xlsx")
            and "stock" in archivo.lower()
        ):
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

    texto = str(valor or "").strip()

    if texto.endswith(".0"):
        texto = texto[:-2]

    return re.sub(
        r"\D",
        "",
        texto
    )


# ============================================================
# CONVERTIR FECHA
# ============================================================

def convertir_fecha_excel(valor):

    texto = str(valor or "").strip()

    if not texto:
        return 0

    # Fecha serial de Excel
    try:

        numero = float(texto)

        if numero > 0:
            return numero

    except (ValueError, TypeError):
        pass

    # Fecha textual
    try:

        partes = texto.replace("-", "/").split("/")

        if len(partes) == 3:

            dia = int(partes[0])
            mes = int(partes[1])
            anio = int(partes[2])

            if anio < 100:
                anio += 2000

            fecha = datetime.datetime(
                anio,
                mes,
                dia
            )

            base = datetime.datetime(
                1899,
                12,
                30
            )

            return (
                fecha - base
            ).days

    except Exception:
        pass

    return 0


# ============================================================
# COLUMNA EXCEL
# ============================================================

def columna_excel(referencia):

    match = re.match(
        r"([A-Z]+)",
        referencia.upper()
    )

    return match.group(1) if match else ""


# ============================================================
# SHARED STRINGS
#
# Solo se cargan las necesarias.
# ============================================================

def cargar_shared_strings(zf):

    archivo = "xl/sharedStrings.xml"

    if archivo not in zf.namelist():
        return []

    resultado = []

    with zf.open(archivo) as xml:

        for _, elemento in ET.iterparse(
            xml,
            events=("end",)
        ):

            if not elemento.tag.endswith("}si"):
                continue

            texto = "".join(
                nodo.text or ""
                for nodo in elemento.iter()
                if nodo.tag.endswith("}t")
            )

            resultado.append(texto)

            elemento.clear()

    return resultado


# ============================================================
# VALOR DE CELDA
# ============================================================

def valor_celda(celda, shared):

    tipo = celda.attrib.get("t")

    nodo = celda.find(
        f"{{{NS}}}v"
    )

    if nodo is not None:

        valor = nodo.text or ""

        if tipo == "s":

            try:
                return shared[int(valor)]

            except Exception:
                return ""

        return valor

    # INLINE STRING

    nodo = celda.find(
        f"{{{NS}}}is"
    )

    if nodo is not None:

        return "".join(
            nodo.text or ""
            for nodo in nodo.iter()
            if nodo.tag.endswith("}t")
        )

    return ""


# ============================================================
# CARGAR STOCK
#
# SOLO LEE:
#
# NUM_DOC
# FECHA_ASIG_ESTUDIO
# ============================================================

def cargar_stock():

    ruta = buscar_ultimo_stock()

    print()
    print("========================================")
    print("STOCK ENCONTRADO")
    print("========================================")
    print(ruta)

    with zipfile.ZipFile(ruta, "r") as zf:

        hojas = [
            x
            for x in zf.namelist()
            if x.startswith("xl/worksheets/")
            and x.endswith(".xml")
        ]

        if not hojas:
            raise FileNotFoundError(
                "No se encontró la hoja del STOCK."
            )

        hoja = hojas[0]

        print()
        print("Leyendo solamente NUM_DOC y FECHA_ASIG_ESTUDIO...")
        print("Hoja utilizada:", hoja)

        shared = cargar_shared_strings(zf)

        stock_dict = {}

        indice_dni = None
        indice_fecha = None

        contador = 0

        with zf.open(hoja) as xml:

            for _, fila in ET.iterparse(
                xml,
                events=("end",)
            ):

                if not fila.tag.endswith("}row"):
                    continue

                # =================================================
                # PRIMERA FILA = ENCABEZADOS
                # =================================================

                if indice_dni is None:

                    for celda in fila:

                        if not celda.tag.endswith("}c"):
                            continue

                        columna = columna_excel(
                            celda.attrib.get("r", "")
                        )

                        nombre = valor_celda(
                            celda,
                            shared
                        ).strip().upper()

                        if nombre == "NUM_DOC":
                            indice_dni = columna

                        elif nombre == "FECHA_ASIG_ESTUDIO":
                            indice_fecha = columna

                    if indice_dni is None:
                        raise KeyError(
                            "No se encontró NUM_DOC en STOCK."
                        )

                    if indice_fecha is None:
                        raise KeyError(
                            "No se encontró FECHA_ASIG_ESTUDIO en STOCK."
                        )

                    print()
                    print("Columnas encontradas:")
                    print("- NUM_DOC")
                    print("- FECHA_ASIG_ESTUDIO")

                    fila.clear()
                    continue

                # =================================================
                # SOLO LEER LAS DOS COLUMNAS
                # =================================================

                dni = ""
                fecha_original = ""

                for celda in fila:

                    if not celda.tag.endswith("}c"):
                        continue

                    columna = columna_excel(
                        celda.attrib.get("r", "")
                    )

                    if columna == indice_dni:

                        dni = normalizar_dni(
                            valor_celda(
                                celda,
                                shared
                            )
                        )

                    elif columna == indice_fecha:

                        fecha_original = valor_celda(
                            celda,
                            shared
                        )

                if not dni:

                    fila.clear()
                    continue

                fecha = convertir_fecha_excel(
                    fecha_original
                )

                anterior = stock_dict.get(dni)

                if (
                    anterior is None
                    or fecha > anterior[1]
                ):

                    stock_dict[dni] = (
                        fecha_original,
                        fecha
                    )

                contador += 1

                if contador % 5000 == 0:

                    print(
                        "Registros procesados:",
                        contador,
                        "| DNI únicos:",
                        len(stock_dict)
                    )

                fila.clear()

    print()
    print("Registros procesados:", contador)
    print("DNI únicos:", len(stock_dict))

    # ============================================================
    # DATAFRAME FINAL
    # ============================================================

    df = pd.DataFrame(
        [
            {
                "NUM_DOC": dni,
                "FECHA_ASIG_ESTUDIO": valores[0]
            }
            for dni, valores in stock_dict.items()
        ]
    )

    print()
    print("========================================")
    print("STOCK PROCESADO")
    print("========================================")
    print("Registros finales:", len(df))

    return df


# ============================================================
# PREPARAR STOCK
# ============================================================

def preparar_stock(df_stock):

    columnas = [
        "NUM_DOC",
        "FECHA_ASIG_ESTUDIO"
    ]

    for columna in columnas:

        if columna not in df_stock.columns:

            raise KeyError(
                f"No se encontró la columna '{columna}' en STOCK."
            )

    stock = df_stock[columnas].copy()

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