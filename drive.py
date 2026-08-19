import os
import json

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


# ============================================================
# GOOGLE DRIVE
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly"
]


# ============================================================
# CARGAR CREDENCIALES
#
# LOCAL:
#   Usa credenciales_drive.json
#
# RENDER:
#   Usa variable de entorno GOOGLE_CREDENTIALS
# ============================================================

def cargar_credenciales():

    # --------------------------------------------------------
    # RENDER
    # --------------------------------------------------------

    credenciales_json = os.environ.get(
        "GOOGLE_CREDENTIALS"
    )

    if credenciales_json:

        try:

            datos = json.loads(
                credenciales_json
            )

            return (
                service_account.Credentials
                .from_service_account_info(
                    datos,
                    scopes=SCOPES
                )
            )

        except Exception as e:

            raise RuntimeError(
                "No se pudieron cargar las "
                "credenciales de Google Drive "
                "desde GOOGLE_CREDENTIALS."
            ) from e

    # --------------------------------------------------------
    # PC LOCAL
    # --------------------------------------------------------

    archivo_credenciales = (
        "credenciales_drive.json"
    )

    if not os.path.exists(
        archivo_credenciales
    ):

        raise FileNotFoundError(
            "No se encontró "
            "'credenciales_drive.json' "
            "y tampoco está definida "
            "la variable GOOGLE_CREDENTIALS."
        )

    return (
        service_account.Credentials
        .from_service_account_file(
            archivo_credenciales,
            scopes=SCOPES
        )
    )


# ============================================================
# CONEXIÓN GOOGLE DRIVE
# ============================================================

credenciales = cargar_credenciales()

drive = build(
    "drive",
    "v3",
    credentials=credenciales
)


# ============================================================
# BUSCAR CARPETA
# ============================================================

def buscar_carpeta_drive(nombre_carpeta):

    # --------------------------------------------------------
    # Primero buscamos por nombre exacto
    # --------------------------------------------------------

    resultado = drive.files().list(
        q=(
            "name = '"
            + nombre_carpeta
            + "' "
            "and mimeType = "
            "'application/vnd.google-apps.folder' "
            "and trashed = false"
        ),
        spaces="drive",
        fields="files(id,name,mimeType)"
    ).execute()

    carpetas = resultado.get(
        "files",
        []
    )

    # --------------------------------------------------------
    # Si no encuentra, hacemos búsqueda general
    # para tolerar diferencias de mayúsculas/minúsculas
    # --------------------------------------------------------

    if not carpetas:

        resultado = drive.files().list(
            q=(
                "mimeType = "
                "'application/vnd.google-apps.folder' "
                "and trashed = false"
            ),
            spaces="drive",
            pageSize=100,
            fields="files(id,name,mimeType)"
        ).execute()

        todas_las_carpetas = resultado.get(
            "files",
            []
        )

        nombre_buscado = (
            nombre_carpeta
            .strip()
            .casefold()
        )

        carpetas = [
            carpeta
            for carpeta in todas_las_carpetas
            if carpeta["name"]
            .strip()
            .casefold()
            == nombre_buscado
        ]

    # --------------------------------------------------------
    # Si sigue sin encontrar
    # --------------------------------------------------------

    if not carpetas:

        raise FileNotFoundError(
            f"No se encontró la carpeta "
            f"'{nombre_carpeta}' en Google Drive."
        )

    carpeta = carpetas[0]

    print()
    print(
        "CARPETA DRIVE ENCONTRADA:"
    )

    print(
        "Nombre:",
        carpeta["name"]
    )

    print(
        "ID:",
        carpeta["id"]
    )

    return carpeta["id"]


# ============================================================
# BUSCAR ÚLTIMO EXCEL DE UNA CARPETA
# ============================================================

def buscar_ultimo_excel_drive(
    carpeta_id
):

    resultado = drive.files().list(
        q=(
            "'"
            + carpeta_id
            + "' in parents "
            "and trashed = false"
        ),
        spaces="drive",
        orderBy="modifiedTime desc",
        pageSize=50,
        fields=(
            "files("
            "id,"
            "name,"
            "mimeType,"
            "modifiedTime"
            ")"
        )
    ).execute()

    archivos = resultado.get(
        "files",
        []
    )

    print()
    print("========================================")
    print("BUSCANDO ARCHIVOS DENTRO DE LA CARPETA")
    print("========================================")

    print(
        "ID CARPETA:",
        carpeta_id
    )

    print("ARCHIVOS ENCONTRADOS:")

    for archivo in archivos:

        print(
            "NOMBRE:",
            archivo.get("name")
        )

        print(
            "TIPO:",
            archivo.get("mimeType")
        )

        print(
            "MODIFICADO:",
            archivo.get("modifiedTime")
        )

        print("----------------------------------------")

    tipos_excel = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel"
    ]

    archivos_excel = [
        archivo
        for archivo in archivos
        if archivo.get("mimeType") in tipos_excel
    ]

    print(
        "CANTIDAD DE EXCEL ENCONTRADOS:",
        len(archivos_excel)
    )

    if not archivos_excel:

        return None

    return archivos_excel[0]


# ============================================================
# DESCARGAR ARCHIVO
# ============================================================

def descargar_archivo_drive(
    archivo,
    ruta_local
):

    print()
    print(
        "========================================"
    )

    print(
        "DESCARGANDO DESDE GOOGLE DRIVE"
    )

    print(
        "Archivo:",
        archivo["name"]
    )

    print(
        "========================================"
    )

    request = drive.files().get_media(
        fileId=archivo["id"]
    )

    carpeta_local = os.path.dirname(
        ruta_local
    )

    if carpeta_local:

        os.makedirs(
            carpeta_local,
            exist_ok=True
        )

    with open(
        ruta_local,
        "wb"
    ) as archivo_local:

        downloader = MediaIoBaseDownload(
            archivo_local,
            request
        )

        terminado = False

        while not terminado:

            estado, terminado = (
                downloader.next_chunk()
            )

            if estado:

                porcentaje = int(
                    estado.progress() * 100
                )

                print(
                    f"Descarga: {porcentaje}%"
                )

    print()

    print(
        "Guardado localmente en:",
        ruta_local
    )


# ============================================================
# DESCARGAR ÚLTIMO ARCHIVO DE UNA CARPETA
# ============================================================

def descargar_ultimo_de_carpeta(
    nombre_carpeta,
    carpeta_local
):

    print()
    print(
        "========================================"
    )

    print(
        "BUSCANDO CARPETA:",
        nombre_carpeta
    )

    print(
        "========================================"
    )

    # --------------------------------------------------------
    # BUSCAR CARPETA
    # --------------------------------------------------------

    carpeta_id = buscar_carpeta_drive(
        nombre_carpeta
    )

    # --------------------------------------------------------
    # BUSCAR ÚLTIMO EXCEL
    # --------------------------------------------------------

    archivo = buscar_ultimo_excel_drive(
        carpeta_id
    )

    if archivo is None:

        raise FileNotFoundError(
            f"No se encontró ningún Excel "
            f"en la carpeta '{nombre_carpeta}'."
        )

    # --------------------------------------------------------
    # CREAR CARPETA LOCAL
    # --------------------------------------------------------

    os.makedirs(
        carpeta_local,
        exist_ok=True
    )

    # --------------------------------------------------------
    # # RUTA LOCAL
    # # --------------------------------------------------------
    
    nombre_archivo = archivo["name"]

    # Google Drive puede devolver el archivo sin extensión
    if not nombre_archivo.lower().endswith((".xlsx", ".xls")):
        if archivo.get("mimeType") == (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            nombre_archivo += ".xlsx"

        elif archivo.get("mimeType") == (
            "application/vnd.ms-excel"
            ):
            nombre_archivo += ".xls"

        ruta_local = os.path.join(
            carpeta_local,
            nombre_archivo
            )

    # --------------------------------------------------------
    # DESCARGAR
    # --------------------------------------------------------

    descargar_archivo_drive(
        archivo,
        ruta_local
    )

    return ruta_local