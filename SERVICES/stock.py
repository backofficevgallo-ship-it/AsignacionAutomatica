import os
from openpyxl import load_workbook


CARPETA_STOCK = "ARCHIVOS/STOCK"


def buscar_ultimo_stock():

    archivos = []

    for archivo in os.listdir(CARPETA_STOCK):

        ruta = os.path.join(
            CARPETA_STOCK,
            archivo
        )

        if not os.path.isfile(ruta):
            continue

        nombre = archivo.lower().strip()

        if (
            "stock" in nombre
            and nombre.endswith(".xlsx")
        ):
            archivos.append(ruta)

    if not archivos:

        raise FileNotFoundError(
            "No se encontró ningún archivo STOCK."
        )

    return max(
        archivos,
        key=os.path.getmtime
    )


def cargar_stock():

    ruta = buscar_ultimo_stock()

    print()
    print("========================================")
    print("STOCK ENCONTRADO")
    print("========================================")
    print(ruta)

    print()
    print("Leyendo STOCK en modo memoria reducida...")

    wb = load_workbook(
        filename=ruta,
        read_only=True,
        data_only=True
    )

    ws = wb.active

    # ========================================================
    # ENCABEZADOS
    # ========================================================

    encabezados = next(
        ws.iter_rows(
            min_row=1,
            max_row=1,
            values_only=True
        )
    )

    encabezados = [
        str(x).strip() if x is not None else ""
        for x in encabezados
    ]

    if "NUM_DOC" not in encabezados:

        wb.close()

        raise KeyError(
            "No se encontró la columna 'NUM_DOC' en STOCK."
        )

    if "FECHA_ASIG_ESTUDIO" not in encabezados:

        wb.close()

        raise KeyError(
            "No se encontró la columna "
            "'FECHA_ASIG_ESTUDIO' en STOCK."
        )

    indice_dni = encabezados.index(
        "NUM_DOC"
    )

    indice_fecha = encabezados.index(
        "FECHA_ASIG_ESTUDIO"
    )

    # ========================================================
    # CREAR DICCIONARIO DIRECTAMENTE
    # ========================================================

    stock = {}

    print()
    print("Procesando registros del STOCK...")

    for fila in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if indice_dni >= len(fila):
            continue

        dni = fila[indice_dni]

        if dni is None:
            continue

        dni = str(dni).strip()

        if dni.endswith(".0"):
            dni = dni[:-2]

        fecha = (
            fila[indice_fecha]
            if indice_fecha < len(fila)
            else None
        )

        # Guardamos solamente un registro por DNI.
        #
        # El cruce posterior necesita la fecha de asignación.
        # Si aparece varias veces, conservamos la última
        # información leída.

        stock[dni] = fecha

    wb.close()

    print()
    print(
        "DNI únicos encontrados en STOCK:",
        len(stock)
    )

    return stock


def preparar_stock(stock):

    return stock