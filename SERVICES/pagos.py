import os
import pandas as pd


CARPETA_PAGOS = "ARCHIVOS/PAGOS"


# ============================================================
# BUSCAR ÚLTIMO ARCHIVO DE PAGOS
# ============================================================

def buscar_ultimo_pago():

    archivos = []

    for archivo in os.listdir(CARPETA_PAGOS):

        ruta = os.path.join(
            CARPETA_PAGOS,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        nombre = archivo.lower().strip()

        if (
            "pagos" in nombre
            and "agosto" in nombre
            and nombre.endswith((".xlsx", ".xls"))
        ):
            archivos.append(ruta)

    if not archivos:

        raise FileNotFoundError(
            "No se encontró ningún archivo de PAGOS AGOSTO."
        )

    return max(
        archivos,
        key=os.path.getmtime
    )


# ============================================================
# CARGAR PAGOS AGOSTO
# ============================================================

def cargar_pagos_agosto():

    carpeta = "ARCHIVOS/PAGOS_LIGERO"

    if not os.path.exists(carpeta):

        raise FileNotFoundError(
            f"No existe la carpeta {carpeta}."
        )

    archivos = []

    for archivo in os.listdir(carpeta):

        ruta = os.path.join(
            carpeta,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        if archivo.lower().endswith(".csv"):

            archivos.append(ruta)

    if not archivos:

        raise FileNotFoundError(
            "No se encontró ningún CSV "
            "de PAGOS LIGERO."
        )

    ruta = max(
        archivos,
        key=os.path.getmtime
    )

    print()
    print("========================================")
    print("PAGOS AGOSTO LIGERO ENCONTRADO")
    print("========================================")

    print(ruta)

    print()
    print("Leyendo DNI e Importe...")

    # --------------------------------------------------------
    # LEER CSV
    # --------------------------------------------------------

    df = pd.read_csv(
        ruta,
        sep=";",
        encoding="utf-8-sig",
        usecols=[
            "DNI",
            "Importe"
        ]
    )

    # --------------------------------------------------------
    # NORMALIZAR COLUMNAS
    # --------------------------------------------------------

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # NORMALIZAR DNI
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # NORMALIZAR IMPORTE
    # --------------------------------------------------------

    df["Importe"] = pd.to_numeric(
        df["Importe"],
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
    print("- Importe")

    return df

# ============================================================
# CRUZAR PAGOS AGOSTO
# ============================================================

def cruzar_pagos_agosto(
    df_asignacion,
    df_pagos
):

    # --------------------------------------------------------
    # Verificar columnas
    # --------------------------------------------------------

    if "DNI" not in df_pagos.columns:

        raise KeyError(
            "No se encontró DNI en PAGOS AGOSTO."
        )

    if "Importe" not in df_pagos.columns:

        raise KeyError(
            "No se encontró Importe en PAGOS AGOSTO."
        )

    if "DNI" not in df_asignacion.columns:

        raise KeyError(
            "No se encontró DNI en ASIGNACION."
        )

    # --------------------------------------------------------
    # Copia
    # --------------------------------------------------------

    pagos = df_pagos[
        [
            "DNI",
            "Importe"
        ]
    ].copy()

    # --------------------------------------------------------
    # Normalizar DNI
    # --------------------------------------------------------

    pagos["DNI"] = (
        pagos["DNI"]
        .astype(str)
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
    )

    df_asignacion["DNI"] = (
        df_asignacion["DNI"]
        .astype(str)
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
    )

    # --------------------------------------------------------
    # Si hay DNI repetidos en pagos,
    # por ahora usamos la primera coincidencia,
    # igual que BUSCARV.
    # --------------------------------------------------------

    pagos = pagos.drop_duplicates(
        subset="DNI",
        keep="first"
    )

    # --------------------------------------------------------
    # Crear mapa DNI → IMPORTE
    # --------------------------------------------------------

    mapa = pagos.set_index(
        "DNI"
    )["Importe"]

    # --------------------------------------------------------
    # Cruce
    #
    # Si no encuentra DNI:
    # queda vacío.
    # --------------------------------------------------------

    df_asignacion["PAGO AGOSTO"] = (
        df_asignacion["DNI"].map(mapa)
    )

    # --------------------------------------------------------
    # Resultado
    # --------------------------------------------------------

    encontrados = (
        df_asignacion["PAGO AGOSTO"].notna().sum()
    )

    no_encontrados = (
        df_asignacion["PAGO AGOSTO"].isna().sum()
    )

    print()
    print("========================================")
    print("CRUCE PAGOS AGOSTO")
    print("========================================")

    print(
        "Clientes en ASIGNACION OK:",
        len(df_asignacion)
    )

    print(
        "Encontrados en PAGOS:",
        encontrados
    )

    print(
        "No encontrados:",
        no_encontrados
    )

    return df_asignacion