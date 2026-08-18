import os
import pandas as pd


CARPETA_COLCHON = "ARCHIVOS/COLCHON"


# ============================================================
# BUSCAR ÚLTIMO ARCHIVO DE COLCHÓN
# ============================================================

def buscar_ultimo_colchon():

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

    df = pd.read_excel(ruta)

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    print()
    print(
        "Cantidad de registros:",
        len(df)
    )

    print()
    print("Columnas encontradas:")

    for columna in df.columns:
        print("-", columna)

    return df


# ============================================================
# CRUZAR COLCHÓN
# ============================================================

def cruzar_colchon(df_asignacion, df_colchon):

    # --------------------------------------------------------
    # Verificar columnas
    # --------------------------------------------------------

    if "DNI" not in df_colchon.columns:
        raise KeyError(
            "No se encontró la columna DNI en COLCHON."
        )

    if "MONTO" not in df_colchon.columns:
        raise KeyError(
            "No se encontró la columna CUOTAS en COLCHON."
        )

    if "DNI" not in df_asignacion.columns:
        raise KeyError(
            "No se encontró la columna DNI en ASIGNACION."
        )

    # --------------------------------------------------------
    # Copia
    # --------------------------------------------------------

    colchon = df_colchon[
        [
            "DNI",
            "MONTO"
        ]
    ].copy()

    # --------------------------------------------------------
    # Normalizar DNI
    # --------------------------------------------------------

    colchon["DNI"] = (
        colchon["DNI"]
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
    # Si hay DNI repetidos en Colchón,
    # usamos la primera coincidencia,
    # igual que un BUSCARV.
    # --------------------------------------------------------

    colchon = colchon.drop_duplicates(
        subset="DNI",
        keep="first"
    )

    # --------------------------------------------------------
    # Crear mapa DNI → CUOTAS
    # --------------------------------------------------------

    mapa = colchon.set_index(
        "DNI"
    )["MONTO"]

    # --------------------------------------------------------
    # Cruce
    #
    # Si no encuentra DNI:
    # queda vacío.
    # --------------------------------------------------------

    df_asignacion["COLCHON"] = (
        df_asignacion["DNI"].map(mapa)
    )

    # --------------------------------------------------------
    # Resultado
    # --------------------------------------------------------

    encontrados = (
        df_asignacion["COLCHON"].notna().sum()
    )

    no_encontrados = (
        df_asignacion["COLCHON"].isna().sum()
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