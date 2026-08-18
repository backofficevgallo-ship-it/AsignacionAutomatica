import os
import pandas as pd


CARPETA_STOCK = "ARCHIVOS/STOCK"


def buscar_ultimo_stock():

    archivos = []

    for archivo in os.listdir(CARPETA_STOCK):

        ruta = os.path.join(CARPETA_STOCK, archivo)

        if not os.path.isfile(ruta):
            continue

        nombre = archivo.lower().strip()

        if (
            "stock" in nombre
            and nombre.endswith((".xlsx", ".xls"))
        ):
            archivos.append(ruta)

    if not archivos:
        raise FileNotFoundError(
            "No se encontró ningún archivo STOCK."
        )

    return max(archivos, key=os.path.getmtime)


def cargar_stock():

    ruta = buscar_ultimo_stock()

    print()
    print("========================================")
    print("STOCK ENCONTRADO")
    print("========================================")
    print(ruta)

    df = pd.read_excel(ruta)

    # Limpiar encabezados
    df.columns = df.columns.astype(str).str.strip()

    print()
    print("Cantidad de registros:", len(df))

    return df


def preparar_stock(df_stock):

    columnas_necesarias = [
        "NUM_DOC",
        "FECHA_ASIG_ESTUDIO"
    ]

    for columna in columnas_necesarias:
        if columna not in df_stock.columns:
            raise KeyError(
                f"No se encontró la columna '{columna}' en STOCK."
            )

    stock = df_stock[
        columnas_necesarias
    ].copy()

    # Normalizar DNI para evitar problemas con .0,
    # espacios o valores numéricos.
    stock["NUM_DOC"] = (
        stock["NUM_DOC"]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    return stock