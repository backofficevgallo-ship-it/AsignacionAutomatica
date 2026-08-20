import os
import pandas as pd


CARPETA_MORIA = "ARCHIVOS/MORIA"


# ============================================================
# BUSCAR ÚLTIMO ARCHIVO DE MORIA
# ============================================================

def buscar_ultimo_moria():

    archivos = []

    for archivo in os.listdir(CARPETA_MORIA):

        ruta = os.path.join(
            CARPETA_MORIA,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        nombre = archivo.lower().strip()

        if (
            "moria" in nombre
            and nombre.endswith((".xlsx", ".xls"))
        ):
            archivos.append(ruta)

    if not archivos:

        raise FileNotFoundError(
            "No se encontró ningún archivo MORIA."
        )

    return max(
        archivos,
        key=os.path.getmtime
    )


# ============================================================
# CARGAR MORIA
# ============================================================

def cargar_moria():

    ruta = buscar_ultimo_moria()

    print()
    print("========================================")
    print("MORIA ENCONTRADO")
    print("========================================")
    print(ruta)

    print()
    print("Leyendo solamente:")
    print("- DNI")
    print("- FECHA CARGA")
    print("- MONTO BCO")
    print("- VTO")

    # ========================================================
    # LEER SOLAMENTE LAS COLUMNAS NECESARIAS
    # ========================================================

    df = pd.read_excel(
        ruta,
        usecols=[
            "DNI",
            "FECHA CARGA",
            "MONTO BCO",
            "VTO"
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
        "DNI",
        "FECHA CARGA",
        "MONTO BCO",
        "VTO"
    ]

    for columna in columnas_necesarias:

        if columna not in df.columns:

            raise KeyError(
                f"No se encontró la columna "
                f"'{columna}' en MORIA."
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
    # CONVERTIR FECHAS
    # ========================================================

    df["FECHA CARGA"] = pd.to_datetime(
        df["FECHA CARGA"],
        errors="coerce",
        dayfirst=True
    )

    df["VTO"] = pd.to_datetime(
        df["VTO"],
        errors="coerce",
        dayfirst=True
    )

    # ========================================================
    # CONVERTIR MONTO BCO
    # ========================================================

    df["MONTO BCO"] = pd.to_numeric(
        df["MONTO BCO"],
        errors="coerce"
    )

    # ========================================================
    # ELIMINAR DUPLICADOS
    #
    # Igual que BUSCARV:
    # primera coincidencia.
    # ========================================================

    df = df.drop_duplicates(
        subset=["DNI"],
        keep="first"
    )

    # ========================================================
    # RESULTADO
    # ========================================================

    print()
    print(
        "Cantidad de registros:",
        len(df)
    )

    print()
    print("Columnas utilizadas:")
    print("- DNI")
    print("- FECHA CARGA")
    print("- MONTO BCO")
    print("- VTO")

    return df


# ============================================================
# CRUZAR MORIA
# ============================================================

def cruzar_moria(
    df_asignacion,
    df_moria
):

    # --------------------------------------------------------
    # Verificar columnas
    # --------------------------------------------------------

    columnas_necesarias = [
        "DNI",
        "FECHA CARGA",
        "MONTO BCO",
        "VTO"
    ]

    for columna in columnas_necesarias:

        if columna not in df_moria.columns:

            raise KeyError(
                f"No se encontró la columna "
                f"{columna} en MORIA."
            )

    if "DNI" not in df_asignacion.columns:

        raise KeyError(
            "No se encontró DNI en ASIGNACION."
        )

    # --------------------------------------------------------
    # Copia
    # --------------------------------------------------------

    moria = df_moria[
        [
            "DNI",
            "FECHA CARGA",
            "MONTO BCO",
            "VTO"
        ]
    ].copy()

    # --------------------------------------------------------
    # Normalizar DNI
    # --------------------------------------------------------

    moria["DNI"] = (
        moria["DNI"]
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
    # Convertir fechas
    # --------------------------------------------------------

    moria["FECHA CARGA"] = pd.to_datetime(
    moria["FECHA CARGA"],
    errors="coerce",
    dayfirst=True
    )

    moria["VTO"] = pd.to_datetime(
    moria["VTO"],
    errors="coerce",
    dayfirst=True
)

    # --------------------------------------------------------
    # Convertir MONTO BCO a número
    # --------------------------------------------------------

    moria["MONTO BCO"] = pd.to_numeric(
        moria["MONTO BCO"],
        errors="coerce"
    )

    # --------------------------------------------------------
    # Si hay DNI repetidos en MORIA,
    # por ahora usamos la primera coincidencia,
    # igual que BUSCARV.
    # --------------------------------------------------------

    moria = moria.drop_duplicates(
        subset="DNI",
        keep="first"
    )

    # --------------------------------------------------------
    # Crear mapas
    # --------------------------------------------------------

    mapa_promesas = moria.set_index(
        "DNI"
    )["MONTO BCO"]

    mapa_alta = moria.set_index(
        "DNI"
    )["FECHA CARGA"]

    mapa_vencimiento = moria.set_index(
        "DNI"
    )["VTO"]

    # --------------------------------------------------------
    # CRUCE PROMESAS
    #
    # Si no encuentra:
    # PROMESAS = 0
    # --------------------------------------------------------

    df_asignacion["PROMESAS"] = (
        df_asignacion["DNI"]
        .map(mapa_promesas)
        .fillna(0)
    )

    # --------------------------------------------------------
    # CRUCE ALTA
    #
    # Si no encuentra:
    # queda vacío
    # --------------------------------------------------------

    df_asignacion["ALTA"] = (
        df_asignacion["DNI"]
        .map(mapa_alta)
    )

    # --------------------------------------------------------
    # CRUCE VENCIMIENTO
    #
    # Si no encuentra:
    # queda vacío
    # --------------------------------------------------------

    df_asignacion["VENCIMIENTO"] = (
        df_asignacion["DNI"]
        .map(mapa_vencimiento)
    )

    # --------------------------------------------------------
    # RESULTADO
    # --------------------------------------------------------

    encontrados = (
        df_asignacion["DNI"]
        .isin(moria["DNI"])
        .sum()
    )

    no_encontrados = (
        len(df_asignacion)
        - encontrados
    )

    print()
    print("========================================")
    print("CRUCE MORIA")
    print("========================================")

    print(
        "Clientes en ASIGNACION OK:",
        len(df_asignacion)
    )

    print(
        "Encontrados en MORIA:",
        encontrados
    )

    print(
        "No encontrados:",
        no_encontrados
    )

    return df_asignacion