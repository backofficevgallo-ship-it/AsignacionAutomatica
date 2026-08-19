import os
import pandas as pd
from openpyxl import load_workbook


CARPETA_STOCK = "ARCHIVOS/STOCK"


def buscar_ultimo_stock():

    archivos = []

    for archivo in os.listdir(CARPETA_STOCK):

        ruta = os.path.join(
            CARPETA_STOCK,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        nombre = archivo.lower().strip()

        if (
            "stock" in nombre
            and nombre.endswith(".xlsx")
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


def cargar_stock():

    ruta = buscar_ultimo_stock()

    print()
    print("========================================")
    print("STOCK ENCONTRADO")
    print("========================================")
    print(ruta)

    # ========================================================
    # ABRIR EXCEL EN MODO SOLO LECTURA
    # ========================================================

    print()
    print("Leyendo STOCK en modo memoria reducida...")

    wb = load_workbook(
        filename=ruta,
        read_only=True,
        data_only=True
    )

    ws = wb.active

    # ========================================================
    # BUSCAR ENCABEZADOS
    # ========================================================

    encabezados = next(
        ws.iter_rows(
            min_row=1,
            max_row=1,
            values_only=True
        )
    )

    encabezados = [
        str(x).strip() if x is not None else ""
        for x in encabezados
    ]

    if "NUM_DOC" not in encabezados:
        wb.close()

        raise KeyError(
            "No se encontró la columna 'NUM_DOC' en STOCK."
        )

    if "FECHA_ASIG_ESTUDIO" not in encabezados:
        wb.close()

        raise KeyError(
            "No se encontró la columna "
            "'FECHA_ASIG_ESTUDIO' en STOCK."
        )

    indice_dni = encabezados.index(
        "NUM_DOC"
    )

    indice_fecha = encabezados.index(
        "FECHA_ASIG_ESTUDIO"
    )

    # ========================================================
    # LEER SOLO LAS DOS COLUMNAS NECESARIAS
    # ========================================================

    documentos = []
    fechas = []

    for fila in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        dni = (
            fila[indice_dni]
            if indice_dni < len(fila)
            else None
        )

        fecha = (
            fila[indice_fecha]
            if indice_fecha < len(fila)
            else None
        )

        documentos.append(dni)
        fechas.append(fecha)

    wb.close()

    # ========================================================
    # CREAR DATAFRAME
    # ========================================================

    df = pd.DataFrame({
        "NUM_DOC": documentos,
        "FECHA_ASIG_ESTUDIO": fechas
    })

    # ========================================================
    # NORMALIZAR DNI
    # ========================================================

    df["NUM_DOC"] = (
        df["NUM_DOC"]
        .astype("string")
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
    )

    print()
    print(
        "Cantidad de registros:",
        len(df)
    )

    return df


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