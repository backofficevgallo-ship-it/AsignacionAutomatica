import os
import pandas as pd


CARPETA_PAGOS = "ARCHIVOS/PAGOS"


# ============================================================
# BUSCAR ÚLTIMO ARCHIVO DE PAGOS
# ============================================================

def buscar_ultimo_pago():

    if not os.path.exists(CARPETA_PAGOS):
        raise FileNotFoundError(
            f"No existe la carpeta {CARPETA_PAGOS}."
        )

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

    ruta = buscar_ultimo_pago()

    print()
    print("========================================")
    print("PAGOS AGOSTO")
    print("========================================")

    print(ruta)

    print()
    print("Leyendo solamente:")
    print("- DNI")
    print("- Importe")

    # ========================================================
    # LEER SOLAMENTE LAS COLUMNAS NECESARIAS
    # ========================================================

    df = pd.read_excel(
        ruta,
        usecols=[
            "DNI",
            "Importe"
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
            "No se encontró DNI en PAGOS AGOSTO."
        )

    if "Importe" not in df.columns:
        raise KeyError(
            "No se encontró Importe en PAGOS AGOSTO."
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
    # NORMALIZAR IMPORTE
    # ========================================================

    df["Importe"] = pd.to_numeric(
        df["Importe"],
        errors="coerce"
    )

    # ========================================================
    # ELIMINAR DNI VACÍOS
    # ========================================================

    df = df[
        df["DNI"].notna()
        & (df["DNI"] != "")
    ].copy()

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

    # ========================================================
    # VERIFICAR COLUMNAS
    # ========================================================

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

    # ========================================================
    # NORMALIZAR DNI DE PAGOS
    # ========================================================

    pagos = df_pagos[
        [
            "DNI",
            "Importe"
        ]
    ].copy()

    pagos["DNI"] = (
        pagos["DNI"]
        .astype("string")
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
    )

    # ========================================================
    # NORMALIZAR DNI ASIGNACION
    # ========================================================

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
    # ELIMINAR DNI VACIOS DE PAGOS
    # ========================================================

    pagos = pagos[
        pagos["DNI"].notna()
        & (pagos["DNI"] != "")
    ].copy()

    # ========================================================
    # DNI REPETIDOS
    #
    # Igual que BUSCARV:
    # usamos la primera coincidencia.
    # ========================================================

    pagos = pagos.drop_duplicates(
        subset="DNI",
        keep="first"
    )

    # ========================================================
    # MAPA DNI → IMPORTE
    # ========================================================

    mapa = dict(
        zip(
            pagos["DNI"],
            pagos["Importe"]
        )
    )

    del pagos


    df_asignacion["PAGO AGOSTO"] = (
        df_asignacion["DNI"].map(mapa)
    )

    del mapa

    # ========================================================
    # RESULTADOS
    # ========================================================

    encontrados = (
        df_asignacion["PAGO AGOSTO"]
        .notna()
        .sum()
    )

    no_encontrados = (
        df_asignacion["PAGO AGOSTO"]
        .isna()
        .sum()
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