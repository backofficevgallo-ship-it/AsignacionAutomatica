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

    # ========================================================
    # NORMALIZAR DNI DEL STOCK
    # ========================================================

    stock_dni = (
        df_stock["NUM_DOC"]
        .astype("string")
        .str.strip()
        .str.replace(
            r"\.0$",
            "",
            regex=True
        )
    )

    # ========================================================
    # NORMALIZAR FECHA
    # ========================================================

    stock_fecha = pd.to_datetime(
        df_stock["FECHA_ASIG_ESTUDIO"],
        dayfirst=True,
        errors="coerce"
    )

    # ========================================================
    # CONTAR DNI CON MÚLTIPLES REGISTROS
    # ========================================================

    dnis_multiples = (
        stock_dni.value_counts()
        .gt(1)
        .sum()
    )

    # ========================================================
    # CREAR TABLA MÍNIMA PARA EL CRUCE
    # ========================================================

    stock = pd.DataFrame({
        "DNI": stock_dni,
        "FECHA": stock_fecha
    })

    # ========================================================
    # TOMAR LA FECHA MÁS RECIENTE POR DNI
    # ========================================================

    stock = (
        stock
        .dropna(subset=["DNI"])
        .sort_values(
            "FECHA",
            ascending=False,
            na_position="last"
        )
        .drop_duplicates(
            subset="DNI",
            keep="first"
        )
    )

    # ========================================================
    # MAPA DNI → FECHA MÁS RECIENTE
    # ========================================================

    mapa = dict(
        zip(
            stock["DNI"],
            stock["FECHA"]
        )
    )

    # ========================================================
    # NORMALIZAR DNI DEL REPORTE
    # ========================================================

    dni_asignacion = (
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
    # CRUCE
    # ========================================================

    df_asignacion["ASIGNACION"] = (
        dni_asignacion.map(mapa)
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
        dnis_multiples
    )

    return df_asignacion