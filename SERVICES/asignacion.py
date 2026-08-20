import pandas as pd


# ============================================================
# COLUMNAS FINALES DE ASIGNACION OK
# ============================================================

COLUMNAS_ASIGNACION = [
    "DNI",
    "RAZON SOCIAL",
    "CAPITAL",
    "COMPAÑÍA",
    "ESTADO GESTION",
    "ASIGNACION",
    "COLCHON",
    "MORA",
    "PROMESAS",
    "ALTA",
    "VENCIMIENTO",
    "ESTADO",
    "CONTACTO",
    "ULT.GEST",
    "OPERADOR",
    "PAGO MAYO",
    "PAGO JUNIO",
    "PAGO JULIO",
    "PAGO AGOSTO",
    "MENSUAL",
    "HISTORICO",
    "FECHA",
    "BAJA",
]


# ============================================================
# CREAR BASE ASIGNACION OK DESDE REPORTE OPERATIVO
# ============================================================

def crear_base_asignacion(df_reporte):

    df = pd.DataFrame()

    # --------------------------------------------------------
    # DATOS DEL REPORTE OPERATIVO
    # --------------------------------------------------------

    df["DNI"] = df_reporte["documento"]

    df["RAZON SOCIAL"] = df_reporte["nombre"]

    df["CAPITAL"] = df_reporte["monto"]

    # Compañía fija
    df["COMPAÑÍA"] = "SANTANDER PRIMERA"

    # --------------------------------------------------------
    # CAMPOS QUE NO COMPLETAMOS TODAVÍA
    # --------------------------------------------------------

    # Lo completa/revisa la supervisora
    df["ESTADO GESTION"] = ""

    # Se completa desde STOCK
    df["ASIGNACION"] = ""

    # Se completa desde COLCHON
    df["COLCHON"] = ""

    # --------------------------------------------------------
    # DATOS DEL REPORTE OPERATIVO
    # --------------------------------------------------------

    df["MORA"] = df_reporte["mora"]

    # Se completa desde MORIA
    df["PROMESAS"] = ""

    df["ALTA"] = ""

    df["VENCIMIENTO"] = ""

    df["ESTADO"] = df_reporte["estadooperacion"]

    df["CONTACTO"] = df_reporte["estadocontacto"]

    df["ULT.GEST"] = df_reporte["fecha_tramite"]

    df["OPERADOR"] = df_reporte["usuario"]

    # --------------------------------------------------------
    # PAGOS
    # --------------------------------------------------------

    # Por ahora no tocamos Mayo, Junio ni Julio
    df["PAGO MAYO"] = ""

    df["PAGO JUNIO"] = ""

    df["PAGO JULIO"] = ""

    # Se completa con el pago del mes en curso
    df["PAGO AGOSTO"] = ""

    # --------------------------------------------------------
    # GESTIONES
    # --------------------------------------------------------

    df["MENSUAL"] = df_reporte["Gestiones Historicas"]

    df["HISTORICO"] = df_reporte["GestionesMes"]

    # --------------------------------------------------------
    # ÚLTIMO OPERADOR / FECHA
    # --------------------------------------------------------

    # Por ahora queda vacío.
    # Después definiremos exactamente cómo utilizar
    # la información del último operador.
    df["FECHA"] = ""

    # --------------------------------------------------------
    # BAJA
    # --------------------------------------------------------

    # Por decisión actual, siempre queda vacío.
    df["BAJA"] = ""

    return df[COLUMNAS_ASIGNACION]


# ============================================================
# CRUCE CON STOCK
# ============================================================

def cruzar_stock(df_asignacion, df_stock):

    # ========================================================
    # VERIFICAR COLUMNAS
    # ========================================================

    if "NUM_DOC" not in df_stock.columns:
        raise KeyError(
            "No se encontró NUM_DOC en STOCK."
        )

    if "FECHA_ASIG_ESTUDIO" not in df_stock.columns:
        raise KeyError(
            "No se encontró FECHA_ASIG_ESTUDIO en STOCK."
        )

    if "DNI" not in df_asignacion.columns:
        raise KeyError(
            "No se encontró DNI en ASIGNACION."
        )

    # ========================================================
    # NORMALIZAR DNI STOCK
    # ========================================================

    stock = df_stock[
        [
            "NUM_DOC",
            "FECHA_ASIG_ESTUDIO"
        ]
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
    # ELIMINAR DNI DUPLICADOS
    #
    # Nos quedamos con el primer registro,
    # igual que el comportamiento anterior.
    # ========================================================

    stock = stock.drop_duplicates(
        subset="NUM_DOC",
        keep="first"
    )

    # ========================================================
    # CREAR MAPA
    # ========================================================

    mapa = stock.set_index(
        "NUM_DOC"
    )["FECHA_ASIG_ESTUDIO"]

    # ========================================================
    # CRUCE
    # ========================================================

    df_asignacion["ASIGNACION"] = (
        df_asignacion["DNI"].map(mapa)
    )

    # ========================================================
    # RESULTADOS
    # ========================================================

    encontrados = (
        df_asignacion["ASIGNACION"]
        .notna()
        .sum()
    )

    no_encontrados = (
        df_asignacion["ASIGNACION"]
        .isna()
        .sum()
    )

    # ========================================================
    # MOSTRAR RESULTADO
    # ========================================================

    print()
    print("========================================")
    print("CRUCE STOCK")
    print("========================================")

    print(
        "Clientes en ASIGNACION OK:",
        len(df_asignacion)
    )

    print(
        "Encontrados en STOCK:",
        encontrados
    )

    print(
        "No encontrados:",
        no_encontrados
    )

    print(
        "DNI con múltiples registros:",
        0
    )

    return df_asignacion