import os
import pandas as pd


CARPETA_STOCK_LIGERO = "ARCHIVOS/STOCK_LIGERO"


# ============================================================
# BUSCAR STOCK LIGERO
# ============================================================

def buscar_stock_ligero():

    if not os.path.exists(CARPETA_STOCK_LIGERO):
        raise FileNotFoundError(
            f"No existe la carpeta {CARPETA_STOCK_LIGERO}."
        )

    archivos = []

    for archivo in os.listdir(CARPETA_STOCK_LIGERO):

        ruta = os.path.join(
            CARPETA_STOCK_LIGERO,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        if archivo.lower().endswith(".csv"):
            archivos.append(ruta)

    if not archivos:
        raise FileNotFoundError(
            "No se encontró stock_ligero.csv."
        )

    return max(
        archivos,
        key=os.path.getmtime
    )


# ============================================================
# CARGAR STOCK
# ============================================================

def cargar_stock():

    ruta = buscar_stock_ligero()

    print()
    print("========================================")
    print("STOCK LIGERO ENCONTRADO")
    print("========================================")

    print(ruta)

    # --------------------------------------------------------
    # LEER SOLAMENTE LAS DOS COLUMNAS
    # --------------------------------------------------------

    df = pd.read_csv(
        ruta,
        sep=";",
        encoding="utf-8-sig",
        usecols=[
            "NUM_DOC",
            "FECHA_ASIG_ESTUDIO"
        ],
        dtype={
            "NUM_DOC": "string"
        }
    )

    # --------------------------------------------------------
    # NORMALIZAR DNI
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # ELIMINAR DNI VACÍOS
    # --------------------------------------------------------

    df = df[
        df["NUM_DOC"].notna()
        & (df["NUM_DOC"] != "")
    ].copy()

    # --------------------------------------------------------
    # ELIMINAR DUPLICADOS
    # --------------------------------------------------------

    df = df.drop_duplicates(
        subset=["NUM_DOC"],
        keep="first"
    )

    print()
    print("Columnas encontradas:")
    print("- NUM_DOC")
    print("- FECHA_ASIG_ESTUDIO")

    print()
    print(
        "Registros finales:",
        len(df)
    )

    print()
    print("========================================")
    print("STOCK PROCESADO")
    print("========================================")

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