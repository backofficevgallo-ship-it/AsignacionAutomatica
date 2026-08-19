import os
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

    if pd.isna(valor):
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
# CARGAR STOCK
#
# Se utiliza Calamine para evitar el alto consumo de memoria
# de openpyxl al abrir el archivo STOCK en Render.
# ============================================================

def cargar_stock():

    ruta = buscar_ultimo_stock()

    print()
    print("========================================")
    print("STOCK ENCONTRADO")
    print("========================================")
    print(ruta)

    print()
    print("Leyendo STOCK con Calamine...")

    # ========================================================
    # LEER SOLAMENTE LAS DOS COLUMNAS NECESARIAS
    # ========================================================

    df_stock = pd.read_excel(
        ruta,
        engine="calamine",
        usecols=[
            "NUM_DOC",
            "FECHA_ASIG_ESTUDIO"
        ]
    )

    # ========================================================
    # VERIFICAR COLUMNAS
    # ========================================================

    if "NUM_DOC" not in df_stock.columns:
        raise KeyError(
            "No se encontró la columna 'NUM_DOC' en STOCK."
        )

    if "FECHA_ASIG_ESTUDIO" not in df_stock.columns:
        raise KeyError(
            "No se encontró la columna "
            "'FECHA_ASIG_ESTUDIO' en STOCK."
        )

    print()
    print("Columnas encontradas:")
    print("- NUM_DOC")
    print("- FECHA_ASIG_ESTUDIO")

    # ========================================================
    # NORMALIZAR DNI
    # ========================================================

    df_stock["NUM_DOC"] = (
        df_stock["NUM_DOC"]
        .apply(normalizar_dni)
    )

    # ========================================================
    # ELIMINAR DNI VACÍOS
    # ========================================================

    df_stock = df_stock[
        df_stock["NUM_DOC"] != ""
    ].copy()

    # ========================================================
    # NORMALIZAR FECHA
    # ========================================================

    df_stock["FECHA_ASIG_ESTUDIO"] = pd.to_datetime(
        df_stock["FECHA_ASIG_ESTUDIO"],
        dayfirst=True,
        errors="coerce"
    )

    # ========================================================
    # ORDENAR POR FECHA
    #
    # El más reciente queda primero.
    # ========================================================

    df_stock = df_stock.sort_values(
        "FECHA_ASIG_ESTUDIO",
        ascending=False,
        na_position="last"
    )

    # ========================================================
    # QUEDARSE CON UN SOLO REGISTRO POR DNI
    #
    # Como está ordenado de más reciente a más antiguo,
    # keep="first" conserva la fecha más reciente.
    # ========================================================

    df_stock = df_stock.drop_duplicates(
        subset=["NUM_DOC"],
        keep="first"
    )

    # ========================================================
    # RESULTADO FINAL
    # ========================================================

    df_stock.reset_index(
        drop=True,
        inplace=True
    )

    print()
    print("========================================")
    print("STOCK PROCESADO")
    print("========================================")

    print(
        "Registros finales:",
        len(df_stock)
    )

    return df_stock


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