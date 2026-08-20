import os
import pandas as pd


CARPETA_STOCK = "ARCHIVOS/STOCK"


# ============================================================
# BUSCAR ÚLTIMO STOCK
# ============================================================

def buscar_ultimo_stock():

    if not os.path.exists(CARPETA_STOCK):
        raise FileNotFoundError(
            f"No existe la carpeta {CARPETA_STOCK}."
        )

    archivos = []

    for archivo in os.listdir(CARPETA_STOCK):

        ruta = os.path.join(
            CARPETA_STOCK,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        if archivo.lower().endswith((".csv")):
            archivos.append(ruta)

    if not archivos:
        raise FileNotFoundError(
            "No se encontró ningún archivo de STOCK."
        )

    return max(
        archivos,
        key=os.path.getmtime
    )


# ============================================================
# CARGAR STOCK
# ============================================================

def cargar_stock():

    ruta = buscar_ultimo_stock()

    print()
    print("========================================")
    print("STOCK ENCONTRADO")
    print("========================================")

    print(ruta)

    print()
    print("Leyendo solamente:")
    print("- NUM_DOC")
    print("- FECHA_ASIG_ESTUDIO")

    # ========================================================
    # LEER SOLAMENTE LAS COLUMNAS NECESARIAS
    # ========================================================

    df = pd.read_csv(
        ruta,
        usecols=[
            "NUM_DOC",
            "FECHA_ASIG_ESTUDIO"
        ]
    )

    # ========================================================
    # NORMALIZAR COLUMNAS
    # ========================================================

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # VERIFICAR COLUMNAS
    # ========================================================

    columnas_necesarias = [
        "NUM_DOC",
        "FECHA_ASIG_ESTUDIO"
    ]

    for columna in columnas_necesarias:

        if columna not in df.columns:

            raise KeyError(
                f"No se encontró la columna "
                f"'{columna}' en STOCK."
            )

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

    # ========================================================
    # ELIMINAR DNI VACÍOS
    # ========================================================

    df = df[
        df["NUM_DOC"].notna()
        & (df["NUM_DOC"] != "")
    ].copy()

    # ========================================================
    # ELIMINAR DUPLICADOS
    # ========================================================

    df = df.drop_duplicates(
        subset=["NUM_DOC"],
        keep="first"
    )

    # ========================================================
    # RESULTADO
    # ========================================================

    print()
    print("Columnas utilizadas:")
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