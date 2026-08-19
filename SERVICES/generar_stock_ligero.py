import os
import pandas as pd


CARPETA_STOCK = "ARCHIVOS/STOCK"
CARPETA_SALIDA = "ARCHIVOS/STOCK_LIGERO"


def buscar_ultimo_stock():

    archivos = []

    if not os.path.exists(CARPETA_STOCK):
        raise FileNotFoundError(
            f"No existe la carpeta {CARPETA_STOCK}"
        )

    for archivo in os.listdir(CARPETA_STOCK):

        ruta = os.path.join(
            CARPETA_STOCK,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        nombre = archivo.lower()

        if nombre.endswith(".xlsx") and "stock" in nombre:
            archivos.append(ruta)

    if not archivos:
        raise FileNotFoundError(
            "No se encontró ningún STOCK."
        )

    return max(
        archivos,
        key=os.path.getmtime
    )


def generar_stock_ligero():

    ruta_stock = buscar_ultimo_stock()

    print()
    print("========================================")
    print("GENERANDO STOCK LIGERO")
    print("========================================")

    print("Stock:", ruta_stock)

    # --------------------------------------------------------
    # LEER SOLAMENTE LAS DOS COLUMNAS
    # --------------------------------------------------------

    df = pd.read_excel(
        ruta_stock,
        usecols=[
            "NUM_DOC",
            "FECHA_ASIG_ESTUDIO"
        ]
    )

    print(
        "Registros originales:",
        len(df)
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
    # CONVERTIR FECHA
    # --------------------------------------------------------

    df["FECHA_ASIG_ESTUDIO"] = pd.to_datetime(
        df["FECHA_ASIG_ESTUDIO"],
        dayfirst=True,
        errors="coerce"
    )

    # --------------------------------------------------------
    # ELIMINAR DNI VACÍOS
    # --------------------------------------------------------

    df = df[
        df["NUM_DOC"].notna()
        & (df["NUM_DOC"] != "")
    ].copy()

    # --------------------------------------------------------
    # ORDENAR POR FECHA
    #
    # La fecha más reciente queda primero.
    # --------------------------------------------------------

    df = df.sort_values(
        "FECHA_ASIG_ESTUDIO",
        ascending=False,
        na_position="last"
    )

    # --------------------------------------------------------
    # UN SOLO REGISTRO POR DNI
    #
    # Conservamos la fecha más reciente.
    # --------------------------------------------------------

    df = df.drop_duplicates(
        subset=["NUM_DOC"],
        keep="first"
    )

    # --------------------------------------------------------
    # CREAR CARPETA
    # --------------------------------------------------------

    os.makedirs(
        CARPETA_SALIDA,
        exist_ok=True
    )

    ruta_salida = os.path.join(
        CARPETA_SALIDA,
        "stock_ligero.csv"
    )

    # --------------------------------------------------------
    # GUARDAR CSV
    # --------------------------------------------------------

    df.to_csv(
    ruta_salida,
    index=False,
    sep=";",
    encoding="utf-8-sig"
)

    print()
    print("========================================")
    print("STOCK LIGERO GENERADO")
    print("========================================")

    print(
        "Registros finales:",
        len(df)
    )

    print(
        "Archivo:",
        ruta_salida
    )

    return ruta_salida


if __name__ == "__main__":

    generar_stock_ligero()