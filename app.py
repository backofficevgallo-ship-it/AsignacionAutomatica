import os
import pandas as pd

from flask import Flask, render_template, send_file, request
from drive import descargar_ultimo_de_carpeta

from SERVICES.asignacion import crear_base_asignacion
from SERVICES.asignacion import cruzar_stock
from SERVICES.stock import cargar_stock

from SERVICES.colchon import cargar_colchon
from SERVICES.colchon import cruzar_colchon

from SERVICES.pagos import cargar_pagos_agosto
from SERVICES.pagos import cruzar_pagos_agosto

from SERVICES.moria import cargar_moria
from SERVICES.moria import cruzar_moria


# ============================================================
# FLASK
# ============================================================

app = Flask(__name__)


# ============================================================
# CARPETAS
# ============================================================

CARPETA_REPORTING = "ARCHIVOS/REPORTING"
CARPETA_STOCK = "ARCHIVOS/STOCK"
CARPETA_COLCHON = "ARCHIVOS/COLCHON"
CARPETA_PAGOS = "ARCHIVOS/PAGOS"
CARPETA_MORIA = "ARCHIVOS/MORIA"

CARPETA_OUTPUT = "OUTPUT"


# ============================================================
# BUSCAR ÚLTIMO REPORTE OPERATIVO
# ============================================================

def buscar_ultimo_reporte_operativo():

    archivos = []

    if not os.path.exists(CARPETA_REPORTING):
        raise FileNotFoundError(
            "No existe la carpeta ARCHIVOS/REPORTING."
        )

    for archivo in os.listdir(CARPETA_REPORTING):

        ruta = os.path.join(
            CARPETA_REPORTING,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        nombre = archivo.lower().strip()

        if (
            (
                "reporte operativo" in nombre
                or "r.operativo" in nombre
            )
            and
            nombre.endswith((".xlsx", ".xls"))
        ):
            archivos.append(ruta)

    if not archivos:

        raise FileNotFoundError(
            "No se encontró ningún REPORTE OPERATIVO."
        )

    return max(
        archivos,
        key=os.path.getmtime
    )


# ============================================================
# CARGAR REPORTE OPERATIVO
# ============================================================

def cargar_reporte_operativo(ruta):

    print()
    print("========================================")
    print("REPORTE OPERATIVO")
    print("========================================")

    print(ruta)

    if not os.path.exists(ruta):

        raise FileNotFoundError(
            "No se encontró el Reporte Operativo seleccionado."
        )

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

    return df


# ============================================================
# NORMALIZAR DNI
# ============================================================

def normalizar_dni(valor):

    if pd.isna(valor):
        return ""

    valor = str(valor).strip()

    if valor.endswith(".0"):
        valor = valor[:-2]

    valor = "".join(
        caracter
        for caracter in valor
        if caracter.isdigit()
    )

    return valor


# ============================================================
# BUSCAR COLUMNA DNI
# ============================================================

def buscar_columna_dni_reporte(reporte):

    for columna in reporte.columns:

        nombre = str(columna).strip().upper()

        if nombre in (
            "DNI",
            "NUM_DOC",
            "DOCUMENTO",
            "DOC.CLI.",
            "DOC_CLI"
        ):
            return columna

    return None


# ============================================================
# SEPARAR FALLECIDOS Y NO GESTIONAR
# ============================================================

def separar_categorias(
    reporte,
    asignacion
):

    if "estadooperacion" not in reporte.columns:

        print()
        print(
            "No se encontró la columna "
            "'estadooperacion' en el REPORTE OPERATIVO."
        )

        return (
            asignacion,
            pd.DataFrame(),
            pd.DataFrame()
        )

    reporte_temp = reporte.copy()
    asignacion_temp = asignacion.copy()

    # --------------------------------------------------------
    # BUSCAR DNI
    # --------------------------------------------------------

    columna_dni_reporte = buscar_columna_dni_reporte(
        reporte_temp
    )

    if columna_dni_reporte is None:

        print()
        print(
            "No se encontró una columna DNI "
            "en el REPORTE OPERATIVO."
        )

        return (
            asignacion,
            pd.DataFrame(),
            pd.DataFrame()
        )

    # --------------------------------------------------------
    # NORMALIZAR DNI
    # --------------------------------------------------------

    reporte_temp["_DNI_CRUCE"] = (
        reporte_temp[columna_dni_reporte]
        .apply(normalizar_dni)
    )

    asignacion_temp["_DNI_CRUCE"] = (
        asignacion_temp["DNI"]
        .apply(normalizar_dni)
    )

    # --------------------------------------------------------
    # ESTADO OPERACION
    # --------------------------------------------------------

    estado_operacion = (
        reporte_temp["estadooperacion"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ========================================================
    # FALLECIDO
    # ========================================================

    reporte_temp["_ES_FALLECIDO"] = (
        estado_operacion
        .str.contains(
            "FALLECIDO",
            na=False
        )
    )

    # ========================================================
    # INCOBRABLE
    #
    # SOLO EXACTAMENTE "INCOBRABLE"
    # ========================================================

    reporte_temp["_ES_INCOBRABLE"] = (
        estado_operacion == "INCOBRABLE"
    )

    # --------------------------------------------------------
    # DNI FALLECIDOS
    # --------------------------------------------------------

    dni_fallecidos = set(
        reporte_temp.loc[
            reporte_temp["_ES_FALLECIDO"],
            "_DNI_CRUCE"
        ]
    )

    dni_fallecidos.discard("")

    # --------------------------------------------------------
    # DNI INCOBRABLE
    # --------------------------------------------------------

    dni_incobrables = set(
        reporte_temp.loc[
            reporte_temp["_ES_INCOBRABLE"],
            "_DNI_CRUCE"
        ]
    )

    dni_incobrables.discard("")

    # ========================================================
    # FALLECIDOS
    # ========================================================

    es_fallecido = (
        asignacion_temp["_DNI_CRUCE"]
        .isin(dni_fallecidos)
    )

    fallecidos = asignacion_temp.loc[
        es_fallecido
    ].copy()

    # ========================================================
    # NO GESTIONAR
    # ========================================================

    es_incobrable = (
        asignacion_temp["_DNI_CRUCE"]
        .isin(dni_incobrables)
    )

    # FALLECIDO TIENE PRIORIDAD

    es_incobrable = (
        es_incobrable
        & ~es_fallecido
    )

    no_gestionar = asignacion_temp.loc[
        es_incobrable
    ].copy()

    # ========================================================
    # ASIGNACION OK
    # ========================================================

    asignacion_ok = asignacion_temp.loc[
        ~es_fallecido
        & ~es_incobrable
    ].copy()

    # ========================================================
    # ELIMINAR AUXILIAR
    # ========================================================

    for dataframe in (
        fallecidos,
        no_gestionar,
        asignacion_ok
    ):

        dataframe.drop(
            columns=["_DNI_CRUCE"],
            inplace=True,
            errors="ignore"
        )

    # ========================================================
    # PREPARAR FALLECIDOS
    # ========================================================

    columnas_fallecidos = [
        "NUM_DOC",
        "RAZON_SOCIAL",
        "CAPITAL",
        "COMPAÑÍA",
        "ASIGNACION",
        "ULT.GEST",
        "CONTACTO",
        "ESTADO"
    ]

    if "NUM_DOC" not in fallecidos.columns:

        fallecidos["NUM_DOC"] = (
            fallecidos["DNI"]
        )

    if "RAZON_SOCIAL" not in fallecidos.columns:

        if "RAZON SOCIAL" in fallecidos.columns:

            fallecidos["RAZON_SOCIAL"] = (
                fallecidos["RAZON SOCIAL"]
            )

    for columna in columnas_fallecidos:

        if columna not in fallecidos.columns:

            fallecidos[columna] = ""

    fallecidos = fallecidos[
        columnas_fallecidos
    ].copy()

    # ========================================================
    # PREPARAR NO GESTIONAR
    # ========================================================

    columnas_no_gestionar = [
        "NUM_DOC",
        "RAZON_SOCIAL",
        "CAPITAL",
        "COMPAÑÍA",
        "ESTADO",
        "CONTACTO"
    ]

    if "NUM_DOC" not in no_gestionar.columns:

        no_gestionar["NUM_DOC"] = (
            no_gestionar["DNI"]
        )

    if "RAZON_SOCIAL" not in no_gestionar.columns:

        if "RAZON SOCIAL" in no_gestionar.columns:

            no_gestionar["RAZON_SOCIAL"] = (
                no_gestionar["RAZON SOCIAL"]
            )

    for columna in columnas_no_gestionar:

        if columna not in no_gestionar.columns:

            no_gestionar[columna] = ""

    no_gestionar = no_gestionar[
        columnas_no_gestionar
    ].copy()

    # ========================================================
    # RESULTADO
    # ========================================================

    print()
    print("========================================")
    print("SEPARACION DE CATEGORIAS")
    print("========================================")

    print(
        "Casos FALLECIDOS:",
        len(fallecidos)
    )

    print(
        "Casos NO GESTIONAR - INCOBRABLE:",
        len(no_gestionar)
    )

    print(
        "Casos en ASIGNACION OK:",
        len(asignacion_ok)
    )

    print(
        "Total:",
        len(fallecidos)
        + len(no_gestionar)
        + len(asignacion_ok)
    )

    return (
        asignacion_ok,
        fallecidos,
        no_gestionar
    )


# ============================================================
# GUARDAR ASIGNACION
# ============================================================

def guardar_asignacion(
    df_asignacion,
    df_fallecidos,
    df_no_gestionar
):

    os.makedirs(
        CARPETA_OUTPUT,
        exist_ok=True
    )

    ruta_salida = os.path.join(
        CARPETA_OUTPUT,
        "ASIGNACION_OK_PRUEBA.xlsx"
    )

    # ========================================================
    # COLUMNAS FALLECIDOS
    # ========================================================

    columnas_fallecidos = [
        "NUM_DOC",
        "RAZON_SOCIAL",
        "CAPITAL",
        "COMPAÑÍA",
        "ASIGNACION",
        "ULT.GEST",
        "CONTACTO",
        "ESTADO"
    ]

    for columna in columnas_fallecidos:

        if columna not in df_fallecidos.columns:

            df_fallecidos[columna] = ""

    df_fallecidos = df_fallecidos[
        columnas_fallecidos
    ].copy()

    # ========================================================
    # COLUMNAS NO GESTIONAR
    # ========================================================

    columnas_no_gestionar = [
        "NUM_DOC",
        "RAZON_SOCIAL",
        "CAPITAL",
        "COMPAÑÍA",
        "ESTADO",
        "CONTACTO"
    ]

    for columna in columnas_no_gestionar:

        if columna not in df_no_gestionar.columns:

            df_no_gestionar[columna] = ""

    df_no_gestionar = df_no_gestionar[
        columnas_no_gestionar
    ].copy()

    # ========================================================
    # FECHAS
    # ========================================================

    columnas_fecha_asignacion = [
        "MORA",
        "ULT.GEST",
        "ALTA",
        "VENCIMIENTO"
    ]

    for columna in columnas_fecha_asignacion:

        if columna in df_asignacion.columns:

            df_asignacion[columna] = pd.to_datetime(
                df_asignacion[columna],
                errors="coerce"
            ).dt.date

    if "ULT.GEST" in df_fallecidos.columns:

        df_fallecidos["ULT.GEST"] = pd.to_datetime(
            df_fallecidos["ULT.GEST"],
            errors="coerce"
        ).dt.date

    # ========================================================
    # GENERAR EXCEL
    # ========================================================

    with pd.ExcelWriter(
        ruta_salida,
        engine="openpyxl"
    ) as writer:

        # ----------------------------------------------------
        # ASIGNACION OK
        # ----------------------------------------------------

        df_asignacion.to_excel(
            writer,
            sheet_name="ASIGNACION OK",
            index=False
        )

        # ----------------------------------------------------
        # FALLECIDOS
        # ----------------------------------------------------

        df_fallecidos.to_excel(
            writer,
            sheet_name="FALLECIDOS",
            index=False
        )

        # ----------------------------------------------------
        # NO GESTIONAR
        # ----------------------------------------------------

        df_no_gestionar.to_excel(
            writer,
            sheet_name="NO GESTIONAR",
            index=False
        )

        # ====================================================
        # FORMATO FECHA CORTA
        # ====================================================

        hoja_asignacion = writer.sheets[
            "ASIGNACION OK"
        ]

        for columna in [
            "MORA",
            "ULT.GEST",
            "ALTA",
            "VENCIMIENTO"
        ]:

            if columna in df_asignacion.columns:

                numero_columna = (
                    df_asignacion.columns.get_loc(
                        columna
                    ) + 1
                )

                for fila in range(
                    2,
                    hoja_asignacion.max_row + 1
                ):

                    hoja_asignacion.cell(
                        row=fila,
                        column=numero_columna
                    ).number_format = "m/d/yyyy"

        # ====================================================
        # FALLECIDOS
        # ====================================================

        hoja_fallecidos = writer.sheets[
            "FALLECIDOS"
        ]

        if "ULT.GEST" in df_fallecidos.columns:

            numero_columna = (
                df_fallecidos.columns.get_loc(
                    "ULT.GEST"
                ) + 1
            )

            for fila in range(
                2,
                hoja_fallecidos.max_row + 1
            ):

                hoja_fallecidos.cell(
                    row=fila,
                    column=numero_columna
                ).number_format = "m/d/yyyy"

    print()
    print("========================================")
    print("ARCHIVO GENERADO")
    print("========================================")

    print(ruta_salida)

    return ruta_salida


# ============================================================
# ACTUALIZAR ARCHIVOS DESDE GOOGLE DRIVE
#
# IMPORTANTE:
# EL REPORTE OPERATIVO YA NO SE DESCARGA DESDE DRIVE.
#
# SOLAMENTE:
# STOCK
# COLCHON
# PAGOS
# MORIA
# ============================================================

def actualizar_archivos_desde_drive():

    print()
    print("========================================")
    print("ACTUALIZANDO ARCHIVOS DESDE GOOGLE DRIVE")
    print("========================================")

    # --------------------------------------------------------
    # STOCK
    # --------------------------------------------------------

    ruta_stock = descargar_ultimo_de_carpeta(
        "STOCK",
        CARPETA_STOCK
    )

    # --------------------------------------------------------
    # COLCHON
    # --------------------------------------------------------

    ruta_colchon = descargar_ultimo_de_carpeta(
        "COLCHON",
        CARPETA_COLCHON
    )

    # --------------------------------------------------------
    # PAGOS
    # --------------------------------------------------------

    ruta_pagos = descargar_ultimo_de_carpeta(
        "PAGOS",
        CARPETA_PAGOS
    )

    # --------------------------------------------------------
    # MORIA
    # --------------------------------------------------------

    ruta_moria = descargar_ultimo_de_carpeta(
        "MORIA",
        CARPETA_MORIA
    )

    print()
    print("========================================")
    print("ARCHIVOS ACTUALIZADOS")
    print("========================================")

    print(
        "Stock:",
        ruta_stock
    )

    print(
        "Colchon:",
        ruta_colchon
    )

    print(
        "Pagos:",
        ruta_pagos
    )

    print(
        "Moria:",
        ruta_moria
    )


# ============================================================
# PROCESAR AUTOMATIZACION
# ============================================================

def ejecutar_automatizacion(ruta_reporte):

    import gc

    # ========================================================
    # 1. ACTUALIZAR ARCHIVOS DESDE DRIVE
    # ========================================================

    actualizar_archivos_desde_drive()

    # ========================================================
    # 2. REPORTE OPERATIVO
    # ========================================================

    reporte = cargar_reporte_operativo(
        ruta_reporte
    )

    # ========================================================
    # 3. BASE ASIGNACION
    # ========================================================

    asignacion = crear_base_asignacion(
        reporte
    )

    # ========================================================
    # 4. STOCK
    # ========================================================

    stock = cargar_stock()

    asignacion = cruzar_stock(
        asignacion,
        stock
    )

    # STOCK YA NO SE NECESITA
    del stock
    gc.collect()

    # ========================================================
    # 5. COLCHON
    # ========================================================

    colchon = cargar_colchon()

    asignacion = cruzar_colchon(
        asignacion,
        colchon
    )

    # COLCHON YA NO SE NECESITA
    del colchon
    gc.collect()

    # ========================================================
    # 6. PAGOS AGOSTO
    # ========================================================

    pagos_agosto = cargar_pagos_agosto()

    asignacion = cruzar_pagos_agosto(
        asignacion,
        pagos_agosto
    )

    # PAGOS YA NO SE NECESITAN
    del pagos_agosto
    gc.collect()

    # ========================================================
    # 7. MORIA
    # ========================================================

    moria = cargar_moria()

    asignacion = cruzar_moria(
        asignacion,
        moria
    )

    # MORIA YA NO SE NECESITA
    del moria
    gc.collect()

    # ========================================================
    # 8. SEPARAR CATEGORIAS
    # ========================================================

    (
        asignacion,
        fallecidos,
        no_gestionar
    ) = separar_categorias(
        reporte,
        asignacion
    )

    # ========================================================
    # EL REPORTE YA NO SE NECESITA
    # ========================================================

    del reporte
    gc.collect()

    # ========================================================
    # 9. GUARDAR RESULTADO
    # ========================================================

    ruta = guardar_asignacion(
        asignacion,
        fallecidos,
        no_gestionar
    )

    # ========================================================
    # RESULTADO
    # ========================================================

    resultado = {
        "asignacion": len(asignacion),
        "fallecidos": len(fallecidos),
        "no_gestionar": len(no_gestionar),
        "total": (
            len(asignacion)
            + len(fallecidos)
            + len(no_gestionar)
        )
    }

    return ruta, resultado


# ============================================================
# PAGINA PRINCIPAL
# ============================================================

@app.route("/")
def inicio():

    return render_template(
        "index.html"
    )


# ============================================================
# EJECUTAR
# ============================================================

@app.route(
    "/ejecutar",
    methods=["POST"]
)
def ejecutar():

    try:

        # ====================================================
        # RECIBIR REPORTE OPERATIVO
        # ====================================================

        archivo = request.files.get(
            "reporte"
        )

        if not archivo or archivo.filename == "":

            raise ValueError(
                "No seleccionaste ningún Reporte Operativo."
            )

        # ====================================================
        # GUARDAR REPORTE LOCALMENTE
        # ====================================================

        os.makedirs(
            CARPETA_REPORTING,
            exist_ok=True
        )

        ruta_reporte = os.path.join(
            CARPETA_REPORTING,
            archivo.filename
        )

        archivo.save(
            ruta_reporte
        )

        print()
        print("========================================")
        print("REPORTE OPERATIVO RECIBIDO")
        print("========================================")

        print(
            "Archivo:",
            archivo.filename
        )

        # ====================================================
        # EJECUTAR
        # ====================================================

        ruta, resultado = (
            ejecutar_automatizacion(
                ruta_reporte
            )
        )

        return render_template(
            "index.html",
            resultado=resultado
        )

    except Exception as e:

        print()
        print("========================================")
        print("ERROR")
        print("========================================")

        print(e)

        return render_template(
            "index.html",
            error=str(e)
        ), 500


# ============================================================
# DESCARGAR RESULTADO
# ============================================================

@app.route("/descargar")
def descargar():

    ruta = os.path.join(
        CARPETA_OUTPUT,
        "ASIGNACION_OK_PRUEBA.xlsx"
    )

    if not os.path.exists(ruta):

        return (
            "Todavía no se generó ningún archivo.",
            404
        )

    return send_file(
        ruta,
        as_attachment=True,
        download_name="ASIGNACION_OK_PRUEBA.xlsx"
    )


# ============================================================
# INICIAR SERVIDOR
# ============================================================

if __name__ == "__main__":

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )