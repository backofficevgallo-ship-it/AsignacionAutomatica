import os
import pandas as pd


CARPETA_COLCHON = "ARCHIVOS/COLCHON"


# ============================================================
# BUSCAR ÚLTIMO ARCHIVO DE COLCHÓN
# ============================================================

def buscar_ultimo_colchon():

    if not os.path.exists(CARPETA_COLCHON):
        raise FileNotFoundError(
            f"No existe la carpeta {CARPETA_COLCHON}."
        )

    archivos = []

    for archivo in os.listdir(CARPETA_COLCHON):

        ruta = os.path.join(
            CARPETA_COLCHON,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        nombre = archivo.lower().strip()

        if (
            "colchon" in nombre
            and nombre.endswith((".xlsx", ".xls"))
        ):
            archivos.append(ruta)

    if not archivos:
        raise FileNotFoundError(
            "No se encontró ningún archivo COLCHON."
        )

    return max(
        archivos,
        key=os.path.getmtime
    )


# ============================================================
# CARGAR COLCHÓN
# ============================================================

def cargar_colchon():

    ruta = buscar_ultimo_colchon()

    print()
    print("========================================")
    print("COLCHON ENCONTRADO")
    print("========================================")

    print(ruta)

    print()
    print("Leyendo solamente:")
    print("- DNI")
    print("- MONTO")

    # ========================================================
    # LEER SOLAMENTE LAS COLUMNAS NECESARIAS
    # ========================================================

    df = pd.read_excel(
        ruta,
        usecols=[
            "DNI",
            "MONTO"
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

    if "DNI" not in df.columns:
        raise KeyError(
            "No se encontró la columna DNI en COLCHON."
        )

    if "MONTO" not in df.columns:
        raise KeyError(
            "No se encontró la columna MONTO en COLCHON."
        )

    # ========================================================
    # NORMALIZAR DNI
    # ========================================================

    df["DNI"] = (
        df["DNI"]
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
        df["DNI"].notna()
        & (df["DNI"] != "")
    ].copy()

    # ========================================================
    # NORMALIZAR MONTO
    # ========================================================

    df["MONTO"] = pd.to_numeric(
        df["MONTO"],
        errors="coerce"
    )

    print()
    print(
        "Cantidad de registros:",
        len(df)
    )

    print()
    print("Columnas utilizadas:")
    print("- DNI")
    print("- MONTO")

    return df


# ============================================================
# CRUZAR COLCHÓN
# ============================================================

def cruzar_colchon(
    df_asignacion,
    df_colchon
):

    # ========================================================
    # VERIFICAR COLUMNAS
    # ========================================================

    if "DNI" not in df_colchon.columns:
        raise KeyError(
            "No se encontró la columna DNI en COLCHON."
        )

    if "MONTO" not in df_colchon.columns:
        raise KeyError(
            "No se encontró la columna MONTO en COLCHON."
        )

    if "DNI" not in df_asignacion.columns:
        raise KeyError(
            "No se encontró la columna DNI en ASIGNACION."
        )

    # ========================================================
    # COPIA
    # ========================================================

    colchon = df_colchon[
        [
            "DNI",
            "MONTO"
        ]
    ].copy()

    # ========================================================
    # NORMALIZAR DNI
    # ========================================================

    colchon["DNI"] = (
        colchon["DNI"]
        .astype("string")
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
    )

    df_asignacion["DNI"] = (
        df_asignacion["DNI"]
        .astype("string")
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
    )

    # ========================================================
    # DNI REPETIDOS
    #
    # Igual que BUSCARV:
    # usamos la primera coincidencia.
    # ========================================================

    colchon = colchon.drop_duplicates(
        subset="DNI",
        keep="first"
    )

    # ========================================================
    # CREAR MAPA DNI → MONTO
    # ========================================================

    mapa = colchon.set_index(
        "DNI"
    )["MONTO"]

    # ========================================================
    # CRUCE
    # ========================================================

    df_asignacion["COLCHON"] = (
        df_asignacion["DNI"].map(mapa)
    )

    # ========================================================
    # RESULTADOS
    # ========================================================

    encontrados = (
        df_asignacion["COLCHON"]
        .notna()
        .sum()
    )

    no_encontrados = (
        df_asignacion["COLCHON"]
        .isna()
        .sum()
    )

    print()
    print("========================================")
    print("CRUCE COLCHON")
    print("========================================")

    print(
        "Clientes en ASIGNACION OK:",
        len(df_asignacion)
    )

    print(
        "Encontrados en COLCHON:",
        encontrados
    )

    print(
        "No encontrados:",
        no_encontrados
    )

    return df_asignacion