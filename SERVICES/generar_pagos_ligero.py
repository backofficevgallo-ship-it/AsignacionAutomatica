import os
import pandas as pd


# ============================================================
# CONFIGURACION
# ============================================================

CARPETA_PAGOS = "ARCHIVOS/PAGOS"
CARPETA_SALIDA = "ARCHIVOS/PAGOS_LIGERO"

ARCHIVO_SALIDA = "pagos_ligero.csv"


# ============================================================
# BUSCAR PAGOS AGOSTO
# ============================================================

def buscar_pagos_agosto():

    archivos = []

    if not os.path.exists(CARPETA_PAGOS):
        raise FileNotFoundError(
            f"No existe la carpeta {CARPETA_PAGOS}."
        )

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
# GENERAR PAGOS LIGERO
# ============================================================

def generar_pagos_ligero():

    ruta_origen = buscar_pagos_agosto()

    print()
    print("========================================")
    print("GENERANDO PAGOS LIGERO")
    print("========================================")

    print("Archivo origen:")
    print(ruta_origen)

    # --------------------------------------------------------
    # Leer solamente DNI e Importe
    # --------------------------------------------------------

    df = pd.read_excel(
        ruta_origen,
        usecols=[
            "DNI",
            "Importe"
        ]
    )

    # --------------------------------------------------------
    # Normalizar columnas
    # --------------------------------------------------------

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Normalizar DNI
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

    df["Importe"] = (
        df["Importe"]
        .round(0)
        .astype("Int64")
        )

    # --------------------------------------------------------
    # Crear carpeta
    # --------------------------------------------------------

    os.makedirs(
        CARPETA_SALIDA,
        exist_ok=True
    )

    ruta_salida = os.path.join(
        CARPETA_SALIDA,
        ARCHIVO_SALIDA
    )

    # --------------------------------------------------------
    # Guardar CSV
    #
    # IMPORTANTE:
    # usamos ; porque Excel en Argentina
    # trabaja habitualmente con este separador.
    # --------------------------------------------------------

    df.to_csv(
        ruta_salida,
        index=False,
        sep=";",
        encoding="utf-8-sig"
    )

    print()
    print("PAGOS LIGERO GENERADO")
    print("========================================")

    print(
        "Registros:",
        len(df)
    )

    print(
        "Columnas:",
        list(df.columns)
    )

    print(
        "Guardado en:",
        ruta_salida
    )

    return ruta_salida


# ============================================================
# EJECUTAR
# ============================================================

if __name__ == "__main__":

    generar_pagos_ligero()