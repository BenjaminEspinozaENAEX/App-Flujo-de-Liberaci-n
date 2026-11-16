# ============================================================
# 01_CARGAR_ARCHIVO_FLUJO
# APP_ESTRATEGIAS_LIBERACION
#
# NUEVO FORMATO ÚNICO:
#   7 archivos obligatorios y 1 opcional:
#   - Liberador 1, 2, 3, 4 y 5.
#   - Diccionario CECO-Plantas.
#   - Diccionario Usuarios-Cargos.
#   - Cambios (opcional).
#   Cada archivo puede ser CSV, Parquet o Excel.
#
# La aplicación:
#   1. valida cada nivel;
#   2. conserva todas las reglas originales, incluidos * y V;
#   3. reconstruye una tabla interna:
#      CECO | Planta | Desde | Hasta | TipoDoc |
#      Lib1 | Lib2 | Lib3 | Lib4 | Lib5
#   4. incorpora Planta y cargos desde los dos diccionarios;
#   5. deja los archivos normalizados en session_state;
#   6. incorpora el historial de cambios cuando se adjunta.
# ============================================================

from __future__ import annotations

import base64
import hashlib
import hmac
import importlib.util
import re
import threading
import zipfile
from datetime import datetime
from email.utils import parsedate_to_datetime
from zoneinfo import ZoneInfo
from io import BytesIO
from pathlib import Path
from urllib.parse import urlparse
from textwrap import dedent
from typing import Any, Callable

import pandas as pd
import requests
import streamlit as st


# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent

LOGO_CANDIDATES = [
    PROJECT_DIR / "assets" / "logo.svg",
    BASE_DIR / "assets" / "logo.svg",
    PROJECT_DIR / "assets" / "logo.png",
    BASE_DIR / "assets" / "logo.png",
    PROJECT_DIR / "assets" / "logo.jpg",
    BASE_DIR / "assets" / "logo.jpg",
    PROJECT_DIR / "assets" / "logo.jpeg",
    BASE_DIR / "assets" / "logo.jpeg",
]

LEVELS = (1, 2, 3, 4, 5)

BASE_RULE_COLUMNS = [
    "CompanyCode",
    "BillingAddress",
    "AccountCategory",
    "CostCenter",
    "cus_POClasedeDocumento",
    "PurchaseGroup",
    "TotalCost Bajo",
    "TotalCost Alto",
]

COMMON_COLUMNS = [
    *BASE_RULE_COLUMNS,
    "User",
    "Required",
    "Tooltip",
]

GROUP_COLUMNS = [
    *BASE_RULE_COLUMNS,
    "Group",
    "User",
    "Required",
    "Tooltip",
]

FLOW_COLUMNS = [
    "CECO",
    "Planta",
    "Desde",
    "Hasta",
    "TipoDoc",
    "Lib1",
    "Lib2",
    "Lib3",
    "Lib4",
    "Lib5",
]

FLOW_KEY_COLUMNS = ["CECO", "Desde", "Hasta", "TipoDoc"]
LIB_COLS = ["Lib1", "Lib2", "Lib3", "Lib4", "Lib5"]

DOC_TYPES = {"AZNB", "AZSR"}
LS_LABEL = "Liberador Servicios"
LS_GROUP_VALUES = {
    "liberacion servicios",
    "liberación servicios",
    "liberador servicios",
}


CECO_DICTIONARY_COLUMNS = ["CECO", "Planta", "Centro"]
USER_DICTIONARY_COLUMNS = ["Correo", "Cargo"]

CHANGE_COLUMNS = [
    "FechaHora",
    "Usuario",
    "CECO",
    "Desde",
    "Hasta",
    "TipoDoc",
    "Campo",
    "ValorAntes",
    "ValorDespues",
    "Nota",
]

FILE_ROLE_LIBERATORS = {level: f"liberador_{level}" for level in LEVELS}
FILE_ROLE_CECO = "dic_ceco"
FILE_ROLE_USERS = "dic_users"
FILE_ROLE_CHANGES = "cambios"
REQUIRED_FILE_ROLES = [
    *[FILE_ROLE_LIBERATORS[level] for level in LEVELS],
    FILE_ROLE_CECO,
    FILE_ROLE_USERS,
]
OPTIONAL_FILE_ROLES = [FILE_ROLE_CHANGES]
ALL_FILE_ROLES = [*REQUIRED_FILE_ROLES, *OPTIONAL_FILE_ROLES]

SESSION_DATA_KEY = "flujo_liberacion_data"
SESSION_FILE_KEY = "flujo_liberacion_file_name"
SESSION_CASE_KEY = "flujo_liberacion_last_case"
SESSION_FILE_BYTES_KEY = "flujo_liberacion_file_bytes"
SESSION_SIGNATURE_KEY = "flujo_liberacion_upload_signature_v05"
SESSION_VALIDATION_KEY = "flujo_liberacion_validation_v05"
SESSION_SOURCE_FILES_KEY = "flujo_liberacion_source_files_v06"
SESSION_CHANGES_KEY = "flujo_liberacion_changes_v01"
SESSION_REMOTE_FILES_KEY = "flujo_liberacion_remote_files_v01"
SESSION_REMOTE_SOURCE_KEY = "flujo_liberacion_remote_source_v01"
SESSION_VERSION_KEY = "flujo_liberacion_version_v01"
SESSION_EDITOR_COUNTER_KEY = "flujo_liberacion_editor_counter_v01"
SESSION_EDITOR_MESSAGE_KEY = "flujo_liberacion_editor_message_v01"
HISTORY_KEY_PAGE_03 = "mod_liberadores_history_v04"
SESSION_SAVED_COUNT_KEY = "flujo_liberacion_saved_changes_v01"

CHILE_TZ = ZoneInfo("America/Santiago")
MESES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
    "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]
NOMBRE_DESCARGA = "Flujo de Liberacion Estrategias ENAEX"
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# ============================================================
# ESTILOS Y LOGO
# ============================================================

def compact_html(value: str) -> str:
    return re.sub(r">\s+<", "><", dedent(value).strip())


def aplicar_estilos() -> None:
    st.markdown(
        """
        <style>
            .stMainBlockContainer,
            .block-container {
                padding-top: 6.5rem !important;
                padding-bottom: 2.5rem !important;
            }

            .fl-logo-wrap {
                width: 100%;
                min-height: 90px;
                display: flex;
                justify-content: center;
                align-items: center;
                margin: .6rem 0 12px;
            }

            .fl-logo-wrap img {
                width: 220px;
                max-width: min(60vw, 220px);
                max-height: 88px;
                object-fit: contain;
            }

            .fl-title {
                text-align: center;
                color: #17365D;
                font-size: 2rem;
                font-weight: 850;
                margin: .2rem 0;
            }

            .fl-subtitle {
                text-align: center;
                color: #64748B;
                font-size: 1rem;
                margin-bottom: 1.2rem;
            }

            .fl-section-title {
                color: #17365D;
                font-size: 1.1rem;
                font-weight: 850;
                margin: .4rem 0 .6rem;
            }

            .fl-help {
                color: #64748B;
                font-size: .9rem;
                margin-bottom: .8rem;
            }

            .format-card {
                border: 1px solid #BFDBFE;
                border-radius: 14px;
                background: #EFF6FF;
                padding: 14px 16px;
                margin: .6rem 0 1rem;
            }

            .format-title {
                color: #17365D;
                font-weight: 850;
                margin-bottom: 6px;
            }

            .format-columns {
                color: #334155;
                font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
                font-size: .82rem;
                overflow-wrap: anywhere;
            }

            .level-card {
                border: 1px solid #E2E8F0;
                border-radius: 12px;
                padding: 12px;
                background: #FFFFFF;
                min-height: 105px;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def buscar_logo() -> Path | None:
    return next(
        (path for path in LOGO_CANDIDATES if path.exists() and path.is_file()),
        None,
    )


def mostrar_logo() -> None:
    path = buscar_logo()
    if path is None:
        return

    try:
        raw = path.read_bytes()
        if path.suffix.lower() == ".svg":
            mime = "image/svg+xml"
        elif path.suffix.lower() == ".png":
            mime = "image/png"
        else:
            mime = "image/jpeg"

        encoded = base64.b64encode(raw).decode("utf-8")
        st.markdown(
            f'<div class="fl-logo-wrap"><img src="data:{mime};base64,{encoded}" '
            'alt="Logo ENAEX"></div>',
            unsafe_allow_html=True,
        )
    except (OSError, UnicodeError):
        pass


# ============================================================
# LIMPIEZA
# ============================================================

def clean_text(value: Any) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value).strip()
    return "" if text.lower() in {
        "", "nan", "none", "null", "<na>", "n/a", "—", "-"
    } else text


def normalized_key(value: Any) -> str:
    return clean_text(value).casefold()


def normalize_column_name(value: Any) -> str:
    text = clean_text(value)
    text = re.sub(r"\s+", " ", text)
    return text


def parse_bound(value: Any, *, low: bool) -> float:
    text = clean_text(value)

    if text in {"", "*"}:
        return 1.0 if low else 1e18

    normalized = text.replace(" ", "")

    if "." in normalized and "," in normalized:
        if normalized.rfind(",") > normalized.rfind("."):
            normalized = normalized.replace(".", "").replace(",", ".")
        else:
            normalized = normalized.replace(",", "")
    elif normalized.count(".") > 1:
        normalized = normalized.replace(".", "")
    elif normalized.count(",") > 1:
        normalized = normalized.replace(",", "")
    elif "," in normalized:
        left, right = normalized.rsplit(",", 1)
        normalized = left + right if len(right) == 3 else left + "." + right
    elif "." in normalized:
        left, right = normalized.rsplit(".", 1)
        if len(right) == 3 and left.replace("-", "").isdigit():
            normalized = left + right

    try:
        return float(normalized)
    except ValueError as error:
        raise ValueError(f"Límite no válido: {value!r}") from error


def normalize_required(value: Any) -> str:
    text = normalized_key(value)
    if text in {"true", "verdadero", "1", "sí", "si", "x"}:
        return "TRUE"
    if text in {"false", "falso", "0", "no"}:
        return "FALSE"
    return clean_text(value) or "TRUE"


def liberator_from_row(row: pd.Series) -> str:
    group = clean_text(row.get("Group", ""))
    user = clean_text(row.get("User", ""))

    if normalized_key(group) in LS_GROUP_VALUES:
        return LS_LABEL

    return user


# ============================================================
# CONEXIÓN REMOTA MEDIANTE REQUESTS + SECRETS
# ============================================================

class RemoteUploadedFile:
    """Adaptador mínimo compatible con read_uploaded_table."""
    def __init__(
        self,
        name: str,
        raw: bytes,
        last_modified: str = "",
        fecha_guardado: datetime | None = None,
    ):
        self.name = Path(name).name
        self._raw = bytes(raw)
        # Fecha de modificación informada por SharePoint (Last-Modified).
        self.last_modified = last_modified or ""
        # Fecha en que se presionó "Guardar cambios" (versión guardada).
        self.fecha_guardado = fecha_guardado

    def getvalue(self) -> bytes:
        return self._raw


# Niveles que pueden no existir en el flujo (Compra Directa usa 4 niveles).
# Si no se cargan, se crea un archivo vacío solo con encabezados.
OPTIONAL_LEVELS = (5,)


def completar_niveles_opcionales(
    uploaded: dict[str, Any],
) -> dict[str, Any]:
    if not uploaded:
        return uploaded

    for level in OPTIONAL_LEVELS:
        role = FILE_ROLE_LIBERATORS[level]
        if role not in uploaded:
            header = ",".join(COMMON_COLUMNS) + "\n"
            uploaded[role] = RemoteUploadedFile(
                f"Liberador_{level}_Compra_Directa_ENAEX_vacio.csv",
                header.encode("utf-8"),
            )

    return uploaded


def secret_section() -> Any:
    try:
        return st.secrets["sharepoint_flujo"]
    except Exception as error:
        raise ValueError(
            "No existe la sección [sharepoint_flujo] en Secrets."
        ) from error


def secret_value(
    key: str,
    default: Any = None,
    *,
    required: bool = False,
) -> Any:
    section = secret_section()
    try:
        value = section.get(key, default)
    except Exception:
        value = default

    if required and not clean_text(value):
        raise ValueError(
            f"Falta configurar sharepoint_flujo.{key} en Secrets."
        )

    return value


def calcular_sha256(texto: str) -> str:
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()


def validar_clave_conexion(clave_ingresada: str) -> bool:
    expected_hash = clean_text(
        secret_value("access_key_sha256", required=True)
    ).casefold()

    if not clave_ingresada or not expected_hash:
        return False

    return hmac.compare_digest(
        calcular_sha256(clave_ingresada).casefold(),
        expected_hash,
    )


def obtener_urls_remotas() -> list[str]:
    """
    Lee una colección de URLs sin asociarlas previamente a ningún rol.

    El rol se determina exclusivamente después de descargar cada archivo,
    usando el nombre real entregado por SharePoint.
    """
    section = secret_section()

    try:
        urls_section = section["urls"]
    except Exception as error:
        raise ValueError(
            "Falta configurar [sharepoint_flujo.urls] en Secrets."
        ) from error

    urls: list[str] = []

    # Admite claves arbitrarias: archivo_a, enlace_01, version_actual, etc.
    try:
        items = urls_section.items()
    except Exception as error:
        raise ValueError(
            "[sharepoint_flujo.urls] debe contener enlaces con claves únicas."
        ) from error

    for _, configured_url in items:
        url = clean_text(configured_url)
        if url:
            urls.append(url)

    min_urls = len(REQUIRED_FILE_ROLES) - len(OPTIONAL_LEVELS)
    if len(urls) < min_urls:
        raise ValueError(
            f"Debes configurar al menos {min_urls} enlaces remotos."
        )

    if len(urls) > 8:
        raise ValueError(
            "Solo se admiten siete archivos obligatorios "
            "y un archivo de Cambios opcional."
        )

    allowed_host = clean_text(
        secret_value(
            "allowed_host",
            "empresassk-my.sharepoint.com",
        )
    ).casefold()

    validated: list[str] = []
    seen: set[str] = set()

    for url in urls:
        parsed = urlparse(url)

        if parsed.scheme.casefold() != "https" or not parsed.netloc:
            raise ValueError(
                "Existe una URL remota que no es HTTPS válida."
            )

        hostname = (parsed.hostname or "").casefold()
        if allowed_host and (
            hostname != allowed_host
            and not hostname.endswith("." + allowed_host)
        ):
            raise ValueError(
                "Existe una URL que no pertenece al host autorizado."
            )

        normalized_url = url.casefold()
        if normalized_url in seen:
            raise ValueError(
                "La configuración contiene enlaces duplicados."
            )

        seen.add(normalized_url)
        validated.append(url)

    return validated


def preparar_url_descarga(url: str) -> str:
    """Convierte un enlace compartido de SharePoint en descarga directa."""
    text = clean_text(url)
    if not text:
        raise ValueError("La URL de descarga está vacía.")

    # En enlaces compartidos de SharePoint, download=1 fuerza el archivo.
    if re.search(r"(?:\?|&)download=1(?:&|$)", text, flags=re.I):
        return text

    separator = "&" if "?" in text else "?"
    return f"{text}{separator}download=1"


def request_headers_from_secrets() -> dict[str, str]:
    headers = {
        "User-Agent": "ENAEX-Estrategias-Liberacion/1.0",
        "Accept": "application/octet-stream, application/zip, */*",
    }

    bearer_token = clean_text(secret_value("bearer_token", ""))
    if bearer_token:
        headers["Authorization"] = f"Bearer {bearer_token}"

    return headers


def request_auth_from_secrets() -> tuple[str, str] | None:
    username = clean_text(secret_value("username", ""))
    password = clean_text(secret_value("password", ""))

    if username and password:
        return username, password

    return None


def limpiar_nombre_content_disposition(value: str) -> str:
    """Extrae y decodifica el filename enviado por SharePoint."""
    from urllib.parse import unquote

    text = clean_text(value)
    if not text:
        return ""

    filename_star = re.search(
        r"filename\*\s*=\s*UTF-8''([^;]+)",
        text,
        flags=re.I,
    )
    if filename_star:
        return Path(
            unquote(filename_star.group(1).strip().strip('"'))
        ).name

    filename = re.search(
        r'filename\s*=\s*"([^"]+)"',
        text,
        flags=re.I,
    )
    if filename:
        return Path(
            unquote(filename.group(1).strip())
        ).name

    filename_plain = re.search(
        r"filename\s*=\s*([^;]+)",
        text,
        flags=re.I,
    )
    if filename_plain:
        return Path(
            unquote(
                filename_plain.group(1).strip().strip('"')
            )
        ).name

    return ""


def detectar_nombre_archivo_remoto(
    response: requests.Response,
) -> str:
    """
    Obtiene el nombre real del archivo.

    Prioridad:
    1. Content-Disposition.
    2. Encabezado x-ms-file-name.
    3. Nombre en la URL final.
    """
    candidates = [
        limpiar_nombre_content_disposition(
            response.headers.get("Content-Disposition", "")
        ),
        clean_text(response.headers.get("x-ms-file-name", "")),
        Path(urlparse(response.url).path).name,
    ]

    supported = {
        ".csv",
        ".parquet",
        ".pq",
        ".xlsx",
        ".xls",
        ".xlsm",
    }

    for candidate in candidates:
        safe_name = Path(candidate).name
        if Path(safe_name).suffix.casefold() in supported:
            return safe_name

    raise ValueError(
        "SharePoint no entregó un nombre de archivo reconocible. "
        "Verifica que el enlace apunte directamente al archivo."
    )


def validar_nombre_y_rol_remoto(
    file_name: str,
) -> str:
    """
    Clasifica únicamente por el nombre descargado.

    No usa orden, posición del enlace ni nombre de la clave en Secrets.
    """
    role = detect_file_role(file_name)

    if role is None:
        raise ValueError(
            f"No se pudo identificar el archivo por su nombre: {file_name}."
        )

    return role


def validar_respuesta_remota(
    response: requests.Response,
    expected_host: str,
) -> None:
    response.raise_for_status()

    final = urlparse(response.url)
    if final.scheme.casefold() != "https":
        raise ValueError(
            "La descarga fue redirigida a una URL no segura."
        )

    allow_redirect_host = bool(
        secret_value("allow_redirect_host", True)
    )
    if (
        not allow_redirect_host
        and final.hostname
        and final.hostname.casefold() != expected_host.casefold()
    ):
        raise ValueError(
            "La descarga fue redirigida a otro host."
        )

    content_type = response.headers.get(
        "Content-Type",
        "",
    ).casefold()
    beginning = response.content[:500].lower()

    if (
        "text/html" in content_type
        or b"<html" in beginning
        or b"<!doctype html" in beginning
    ):
        raise ValueError(
            "SharePoint devolvió una página HTML en lugar del archivo."
        )

    if not response.content:
        raise ValueError(
            "SharePoint devolvió un archivo vacío."
        )

    max_file_mb = int(secret_value("max_file_mb", 100))
    if len(response.content) > max_file_mb * 1024 * 1024:
        raise ValueError(
            f"El archivo remoto supera el máximo de {max_file_mb} MB."
        )


@st.cache_data(show_spinner=False, ttl=300)
def descargar_archivo_remoto_cache(
    url: str,
    headers_items: tuple[tuple[str, str], ...],
    auth: tuple[str, str] | None,
    timeout_seconds: int,
    verify_ssl: bool,
) -> tuple[str, bytes, str]:
    response = requests.get(
        preparar_url_descarga(url),
        headers=dict(headers_items),
        auth=auth,
        allow_redirects=True,
        timeout=timeout_seconds,
        verify=verify_ssl,
    )

    expected_host = urlparse(url).hostname or ""
    validar_respuesta_remota(response, expected_host)
    file_name = detectar_nombre_archivo_remoto(response)
    last_modified = response.headers.get("Last-Modified", "")

    return file_name, response.content, last_modified


def conectar_y_obtener_archivos(
    clave_ingresada: str,
    progress_callback: Callable[[int, str], None] | None = None,
) -> dict[str, RemoteUploadedFile]:
    if not validar_clave_conexion(clave_ingresada):
        raise ValueError("Clave de conexión incorrecta.")

    urls = obtener_urls_remotas()
    total = len(urls)

    headers = request_headers_from_secrets()
    auth = request_auth_from_secrets()
    timeout_seconds = int(secret_value("timeout_seconds", 90))
    verify_ssl = bool(secret_value("verify_ssl", True))

    uploaded: dict[str, RemoteUploadedFile] = {}
    downloaded_names: list[str] = []

    for index, url in enumerate(urls, start=1):
        if progress_callback is not None:
            progress_callback(
                int((index - 1) / total * 100),
                f"Descargando archivo {index} de {total}...",
            )

        file_name, raw, last_modified = descargar_archivo_remoto_cache(
            url=url,
            headers_items=tuple(sorted(headers.items())),
            auth=auth,
            timeout_seconds=timeout_seconds,
            verify_ssl=verify_ssl,
        )

        role = validar_nombre_y_rol_remoto(file_name)

        if role in uploaded:
            raise ValueError(
                "Se descargaron dos archivos para el mismo rol: "
                f"{role_label(role)}. Revisa sus nombres."
            )

        uploaded[role] = RemoteUploadedFile(
            file_name,
            raw,
            last_modified,
        )
        downloaded_names.append(file_name)

        if progress_callback is not None:
            progress_callback(
                int(index / total * 100),
                (
                    f"Archivo {index} de {total} descargado · "
                    f"{file_name} → {role_label(role)}."
                ),
            )

    uploaded = completar_niveles_opcionales(uploaded)

    missing = [
        role
        for role in REQUIRED_FILE_ROLES
        if role not in uploaded
    ]
    if missing:
        raise ValueError(
            "Después de clasificar por nombre faltan: "
            + ", ".join(role_label(role) for role in missing)
            + ". Archivos detectados: "
            + ", ".join(downloaded_names)
        )

    if len(uploaded) not in {7, 8}:
        raise ValueError(
            "La conexión debe producir siete archivos obligatorios "
            "y, opcionalmente, Cambios."
        )

    return uploaded



# ============================================================
# LECTURA DE CSV / PARQUET / EXCEL
# ============================================================

def read_csv_flexible(raw: bytes) -> pd.DataFrame:
    attempts = [
        {"sep": None, "engine": "python", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "utf-8-sig"},
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": "\t", "encoding": "utf-8-sig"},
        {"sep": None, "engine": "python", "encoding": "latin-1"},
    ]

    last_error: Exception | None = None

    for options in attempts:
        try:
            frame = pd.read_csv(BytesIO(raw), dtype=object, **options)
            if len(frame.columns) >= 2:
                return frame
        except Exception as error:
            last_error = error

    raise ValueError("No fue posible interpretar el CSV.") from last_error


def read_uploaded_table(uploaded_file: Any) -> tuple[pd.DataFrame, bytes]:
    raw = uploaded_file.getvalue()
    if not raw:
        raise ValueError("El archivo está vacío.")

    suffix = Path(uploaded_file.name).suffix.lower()

    try:
        if suffix == ".csv":
            frame = read_csv_flexible(raw)
        elif suffix in {".parquet", ".pq"}:
            frame = pd.read_parquet(BytesIO(raw))
        elif suffix in {".xlsx", ".xls", ".xlsm"}:
            frame = pd.read_excel(BytesIO(raw), dtype=object)
        else:
            raise ValueError(
                "Formato no admitido. Usa CSV, Parquet, XLSX, XLS o XLSM."
            )
    except ValueError:
        raise
    except Exception as error:
        raise ValueError(
            f"No fue posible leer {uploaded_file.name}."
        ) from error

    frame.columns = [normalize_column_name(column) for column in frame.columns]
    frame = frame.dropna(axis=0, how="all")
    if not frame.empty:
        # Si el archivo solo trae encabezados, se conservan sus columnas.
        frame = frame.dropna(axis=1, how="all")

    return frame, raw


# ============================================================
# VALIDACIÓN POR NIVEL
# ============================================================

def expected_columns_for_level(level: int) -> list[str]:
    # Group es opcional en cualquier nivel.
    return COMMON_COLUMNS


def normalize_level_dataframe(
    dataframe: pd.DataFrame,
    level: int,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    frame = dataframe.copy()
    frame.columns = [normalize_column_name(column) for column in frame.columns]

    required = expected_columns_for_level(level)
    missing = [column for column in required if column not in frame.columns]

    if missing:
        raise ValueError(
            f"Liberador {level}: faltan columnas obligatorias: "
            + ", ".join(missing)
        )

    if "Group" not in frame.columns:
        frame.insert(
            frame.columns.get_loc("User"),
            "Group",
            "",
        )

    frame = frame.loc[:, GROUP_COLUMNS].copy()

    for column in GROUP_COLUMNS:
        frame[column] = frame[column].map(clean_text)

    frame["Required"] = frame["Required"].map(normalize_required)

    invalid_ranges: list[dict[str, Any]] = []
    empty_approver_rows: list[int] = []

    lows: list[float] = []
    highs: list[float] = []

    for index, row in frame.iterrows():
        excel_row = int(index) + 2

        try:
            low = parse_bound(row["TotalCost Bajo"], low=True)
            high = parse_bound(row["TotalCost Alto"], low=False)
        except ValueError as error:
            invalid_ranges.append({
                "Nivel": level,
                "Fila": excel_row,
                "Detalle": str(error),
            })
            low, high = 1.0, 1e18

        if low > high:
            invalid_ranges.append({
                "Nivel": level,
                "Fila": excel_row,
                "Detalle": "TotalCost Bajo es mayor que TotalCost Alto.",
            })

        lows.append(low)
        highs.append(high)

        if not liberator_from_row(row):
            empty_approver_rows.append(excel_row)

    if invalid_ranges:
        preview = "; ".join(
            f"fila {item['Fila']}: {item['Detalle']}"
            for item in invalid_ranges[:8]
        )
        raise ValueError(f"Liberador {level}: rangos inválidos. {preview}")

    frame["_DesdeNum"] = lows
    frame["_HastaNum"] = highs
    frame["_Nivel"] = level
    frame["_Liberador"] = [
        liberator_from_row(row) for _, row in frame.iterrows()
    ]

    explicit_mask = (
        frame["CostCenter"].ne("")
        & frame["CostCenter"].ne("*")
        & frame["cus_POClasedeDocumento"].isin(DOC_TYPES)
    )

    duplicate_columns = [
        *BASE_RULE_COLUMNS,
        "Group",
        "User",
    ]
    duplicate_rows = int(
        frame.duplicated(subset=duplicate_columns, keep=False).sum()
    )

    report = {
        "level": level,
        "rows": len(frame),
        "explicit_rows": int(explicit_mask.sum()),
        "special_rows": int((~explicit_mask).sum()),
        "empty_approver_rows": empty_approver_rows,
        "duplicate_rows": duplicate_rows,
        "service_group_rows": int(
            frame["Group"].map(normalized_key).isin(LS_GROUP_VALUES).sum()
        ),
    }

    return frame.reset_index(drop=True), report


# ============================================================
# DICCIONARIOS
# ============================================================

def normalize_ceco_dictionary(
    dataframe: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    frame = dataframe.copy()
    frame.columns = [normalize_column_name(column) for column in frame.columns]

    missing = [
        column
        for column in ["CECO", "Planta"]
        if column not in frame.columns
    ]
    if missing:
        raise ValueError(
            "Diccionario CECO-Plantas: faltan columnas obligatorias: "
            + ", ".join(missing)
        )

    if "Centro" not in frame.columns:
        frame["Centro"] = ""

    frame = frame.loc[:, CECO_DICTIONARY_COLUMNS].copy()

    for column in CECO_DICTIONARY_COLUMNS:
        frame[column] = frame[column].map(clean_text)

    frame = frame[frame["CECO"].ne("")].copy()

    duplicate_rows = int(
        frame.duplicated(subset=["CECO"], keep=False).sum()
    )

    conflicting_cecos: list[dict[str, Any]] = []
    for ceco, group in frame.groupby("CECO", sort=False):
        plants = [
            value
            for value in group["Planta"].drop_duplicates().tolist()
            if value
        ]
        if len(plants) > 1:
            conflicting_cecos.append({
                "CECO": ceco,
                "Plantas": " | ".join(plants),
            })

    result = (
        frame.drop_duplicates(subset=["CECO"], keep="last")
        .sort_values("CECO", kind="stable")
        .reset_index(drop=True)
    )

    report = {
        "rows": len(result),
        "duplicate_rows": duplicate_rows,
        "conflicts": pd.DataFrame(conflicting_cecos),
        "empty_plants": int(result["Planta"].eq("").sum()),
    }
    return result, report


def normalize_user_dictionary(
    dataframe: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    frame = dataframe.copy()
    frame.columns = [normalize_column_name(column) for column in frame.columns]

    aliases = {
        "Email": "Correo",
        "Mail": "Correo",
        "Usuario": "Correo",
        "Rol": "Cargo",
        "Role": "Cargo",
    }
    frame = frame.rename(
        columns={
            column: aliases.get(column, column)
            for column in frame.columns
        }
    )

    missing = [
        column
        for column in USER_DICTIONARY_COLUMNS
        if column not in frame.columns
    ]
    if missing:
        raise ValueError(
            "Diccionario Usuarios-Cargos: faltan columnas obligatorias: "
            + ", ".join(missing)
        )

    frame = frame.loc[:, USER_DICTIONARY_COLUMNS].copy()
    frame["Correo"] = frame["Correo"].map(clean_text)
    frame["Cargo"] = frame["Cargo"].map(clean_text)
    frame = frame[frame["Correo"].ne("")].copy()

    duplicate_rows = int(
        frame.duplicated(subset=["Correo"], keep=False).sum()
    )

    result = (
        frame.drop_duplicates(subset=["Correo"], keep="last")
        .sort_values("Correo", key=lambda values: values.str.casefold())
        .reset_index(drop=True)
    )

    report = {
        "rows": len(result),
        "duplicate_rows": duplicate_rows,
        "empty_roles": int(result["Cargo"].eq("").sum()),
    }
    return result, report


def apply_ceco_dictionary(
    flow: pd.DataFrame,
    ceco_dictionary: pd.DataFrame,
) -> tuple[pd.DataFrame, list[str], list[str]]:
    result = flow.copy()
    plant_by_ceco = dict(
        zip(
            ceco_dictionary["CECO"],
            ceco_dictionary["Planta"],
        )
    )
    known_cecos = set(ceco_dictionary["CECO"].astype(str))

    result["Planta"] = result["CECO"].map(plant_by_ceco).fillna("")

    missing_cecos = sorted(
        set(result["CECO"].astype(str)) - known_cecos
    )
    cecos_without_plant = sorted(
        result.loc[
            result["CECO"].isin(known_cecos)
            & result["Planta"].eq(""),
            "CECO",
        ]
        .drop_duplicates()
        .astype(str)
        .tolist()
    )
    return result, missing_cecos, cecos_without_plant


def normalize_changes_file(
    dataframe: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Normaliza el archivo opcional de historial de cambios."""
    frame = dataframe.copy()
    frame.columns = [
        normalize_column_name(column)
        for column in frame.columns
    ]

    aliases = {
        "Fecha Hora": "FechaHora",
        "Fecha_Hora": "FechaHora",
        "Responsable": "Usuario",
        "Tipo Documento": "TipoDoc",
        "TipoDocumento": "TipoDoc",
        "Valor Antes": "ValorAntes",
        "Valor Después": "ValorDespues",
        "Valor Despues": "ValorDespues",
        "Observación": "Nota",
        "Observacion": "Nota",
        "Motivo": "Nota",
    }
    frame = frame.rename(
        columns={
            column: aliases.get(column, column)
            for column in frame.columns
        }
    )

    # El archivo puede venir de versiones anteriores con menos columnas.
    for column in CHANGE_COLUMNS:
        if column not in frame.columns:
            frame[column] = ""

    result = frame.loc[:, CHANGE_COLUMNS].copy()

    for column in CHANGE_COLUMNS:
        result[column] = result[column].map(clean_text)

    if not result.empty:
        non_empty = [
            any(clean_text(value) for value in row)
            for row in result.itertuples(index=False)
        ]
        result = result[non_empty]
    result = result.reset_index(drop=True)

    report = {
        "rows": len(result),
        "empty_user": int(result["Usuario"].eq("").sum()),
        "empty_ceco": int(result["CECO"].eq("").sum()),
        "empty_note": int(result["Nota"].eq("").sum()),
    }
    return result, report


# ============================================================
# RECONSTRUCCIÓN DEL FLUJO INTERNO
# ============================================================

def company_from_ceco(ceco: str) -> str:
    match = re.match(r"^(EC\d{2})", clean_text(ceco).upper())
    return match.group(1) if match else ""


def canonical_rule_key(row: pd.Series) -> tuple[str, str, float, float]:
    return (
        clean_text(row["CostCenter"]),
        clean_text(row["cus_POClasedeDocumento"]).upper(),
        float(row["_DesdeNum"]),
        float(row["_HastaNum"]),
    )


def display_bound(value: float, *, high: bool) -> int | float:
    if high and value >= 1e17:
        return 999_999_999_999

    if float(value).is_integer():
        return int(value)

    return float(value)


def reconstruct_flow(
    level_frames: dict[int, pd.DataFrame],
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Reconstruye el flujo indexando cada nivel una sola vez.
    Mismo resultado que la versión original, pero sin filtrar el
    DataFrame completo por cada regla (de ~25 s a menos de 1 s).
    """
    rule_keys: set[tuple[str, str, float, float]] = set()

    # nivel -> (CECO, TipoDoc, Desde, Hasta) -> [(CompanyCode, Liberador)]
    index_by_level: dict[
        int,
        dict[tuple[Any, Any, Any, Any], list[tuple[Any, Any]]],
    ] = {}

    for level in LEVELS:
        frame = level_frames.get(level, pd.DataFrame())
        level_index: dict[
            tuple[Any, Any, Any, Any],
            list[tuple[Any, Any]],
        ] = {}
        index_by_level[level] = level_index

        if frame is None or frame.empty:
            continue

        explicit = frame[
            frame["CostCenter"].ne("")
            & frame["CostCenter"].ne("*")
            & frame["cus_POClasedeDocumento"].isin(DOC_TYPES)
        ]

        for ceco_raw, doc_raw, low, high, company, liberator in zip(
            explicit["CostCenter"],
            explicit["cus_POClasedeDocumento"],
            explicit["_DesdeNum"],
            explicit["_HastaNum"],
            explicit["CompanyCode"],
            explicit["_Liberador"],
        ):
            rule_keys.add(
                (
                    clean_text(ceco_raw),
                    clean_text(doc_raw).upper(),
                    float(low),
                    float(high),
                )
            )
            level_index.setdefault(
                (ceco_raw, doc_raw, low, high),
                [],
            ).append((company, liberator))

    records: list[dict[str, Any]] = []
    conflicts: list[dict[str, Any]] = []

    for ceco, doc, low, high in sorted(rule_keys):
        record: dict[str, Any] = {
            "CECO": ceco,
            "Planta": "",
            "Desde": display_bound(low, high=False),
            "Hasta": display_bound(high, high=True),
            "TipoDoc": doc,
            "Lib1": "",
            "Lib2": "",
            "Lib3": "",
            "Lib4": "",
            "Lib5": "",
        }

        expected_company = company_from_ceco(ceco)

        for level in LEVELS:
            matches = index_by_level[level].get(
                (ceco, doc, low, high),
                [],
            )

            if expected_company:
                preferred = [
                    item
                    for item in matches
                    if item[0] in {expected_company, "*", ""}
                ]
                if preferred:
                    matches = preferred

            values = [
                value
                for value in (clean_text(item[1]) for item in matches)
                if value
            ]

            unique_values = list(dict.fromkeys(values))

            if len(unique_values) == 1:
                record[f"Lib{level}"] = unique_values[0]
            elif len(unique_values) > 1:
                conflicts.append({
                    "CECO": ceco,
                    "TipoDoc": doc,
                    "Desde": display_bound(low, high=False),
                    "Hasta": display_bound(high, high=True),
                    "Nivel": level,
                    "Liberadores": " | ".join(unique_values),
                })
                record[f"Lib{level}"] = unique_values[0]

        records.append(record)

    flow = pd.DataFrame(records, columns=FLOW_COLUMNS)

    if not flow.empty:
        flow = (
            flow.sort_values(
                ["CECO", "TipoDoc", "Desde", "Hasta"],
                kind="stable",
            )
            .reset_index(drop=True)
        )

    duplicate_rule_rows = int(
        flow.duplicated(subset=FLOW_KEY_COLUMNS, keep=False).sum()
    )

    empty_flow_rows = (
        flow[LIB_COLS]
        .apply(lambda row: not any(clean_text(value) for value in row), axis=1)
        .sum()
        if not flow.empty
        else 0
    )

    report = {
        "conflicts": pd.DataFrame(conflicts),
        "duplicate_rule_rows": duplicate_rule_rows,
        "empty_flow_rows": int(empty_flow_rows),
        "rule_count": len(flow),
        "ceco_count": int(flow["CECO"].nunique()) if not flow.empty else 0,
    }

    return flow, report


# ============================================================
# FIRMA, ESTADO Y CARGA
# ============================================================

def files_signature(files: dict[str, Any]) -> str:
    parts: list[str] = []

    for role in ALL_FILE_ROLES:
        uploaded = files.get(role)
        if uploaded is None:
            parts.append(f"{role}|NO_CARGADO")
            continue

        raw = uploaded.getvalue()
        digest = hashlib.sha1(raw).hexdigest()
        parts.append(
            f"{role}|{uploaded.name}|{len(raw)}|{digest}"
        )

    return "||".join(parts)



def clear_active_files() -> None:
    for key in [
        SESSION_DATA_KEY,
        SESSION_FILE_KEY,
        SESSION_CASE_KEY,
        SESSION_FILE_BYTES_KEY,
        SESSION_SIGNATURE_KEY,
        SESSION_VALIDATION_KEY,
        SESSION_SOURCE_FILES_KEY,
        SESSION_CHANGES_KEY,
        SESSION_REMOTE_FILES_KEY,
        SESSION_REMOTE_SOURCE_KEY,
        SESSION_VERSION_KEY,
        SESSION_EDITOR_MESSAGE_KEY,
        SESSION_SAVED_COUNT_KEY,
    ]:
        st.session_state.pop(key, None)


def load_seven_files(
    uploaded_files: dict[str, Any],
    progress_callback: Callable[[int, str], None] | None = None,
) -> tuple[dict[str, Any], dict[str, bytes], dict[str, Any]]:
    """
    Lee y consolida siete archivos obligatorios y un archivo
    de cambios opcional.
    """
    def report(percent: int, message: str) -> None:
        if progress_callback is not None:
            progress_callback(
                max(0, min(100, int(percent))),
                message,
            )

    level_frames: dict[int, pd.DataFrame] = {}
    level_reports: dict[int, dict[str, Any]] = {}
    source_bytes: dict[str, bytes] = {}

    ordered_roles = [
        *[FILE_ROLE_LIBERATORS[level] for level in LEVELS],
        FILE_ROLE_CECO,
        FILE_ROLE_USERS,
    ]
    if FILE_ROLE_CHANGES in uploaded_files:
        ordered_roles.append(FILE_ROLE_CHANGES)

    total_files = len(ordered_roles)
    completed_files = 0

    def file_progress(
        file_number: int,
        stage_fraction: float,
    ) -> int:
        base = (file_number - 1) / total_files
        value = base + stage_fraction / total_files
        # Reserva el 15% final para reconstrucción y activación.
        return int(2 + value * 78)

    report(
        1,
        f"Preparando carga de {total_files} archivos...",
    )

    # 1–5: liberadores.
    for level in LEVELS:
        role = FILE_ROLE_LIBERATORS[level]
        file_number = ordered_roles.index(role) + 1

        report(
            file_progress(file_number, 0.10),
            (
                f"Archivo {file_number} de {total_files} · "
                f"Leyendo Liberador {level}..."
            ),
        )
        frame_raw, raw = read_uploaded_table(uploaded_files[role])

        report(
            file_progress(file_number, 0.55),
            (
                f"Archivo {file_number} de {total_files} · "
                f"Validando Liberador {level}..."
            ),
        )
        frame, validation_report = normalize_level_dataframe(
            frame_raw,
            level,
        )

        level_frames[level] = frame
        level_reports[level] = validation_report
        source_bytes[role] = raw
        completed_files += 1

        report(
            file_progress(file_number, 1.0),
            (
                f"Archivo {file_number} de {total_files} completado · "
                f"Liberador {level} · {len(frame):,} filas"
            ).replace(",", "."),
        )

    # Diccionario CECO.
    role = FILE_ROLE_CECO
    file_number = ordered_roles.index(role) + 1
    report(
        file_progress(file_number, 0.10),
        (
            f"Archivo {file_number} de {total_files} · "
            "Leyendo Diccionario CECO-Plantas..."
        ),
    )
    ceco_raw, ceco_bytes = read_uploaded_table(
        uploaded_files[role]
    )
    report(
        file_progress(file_number, 0.55),
        (
            f"Archivo {file_number} de {total_files} · "
            "Validando Diccionario CECO-Plantas..."
        ),
    )
    ceco_dictionary, ceco_report = normalize_ceco_dictionary(
        ceco_raw
    )
    source_bytes[role] = ceco_bytes
    completed_files += 1
    report(
        file_progress(file_number, 1.0),
        (
            f"Archivo {file_number} de {total_files} completado · "
            f"CECO-Plantas · {len(ceco_dictionary):,} registros"
        ).replace(",", "."),
    )

    # Diccionario usuarios.
    role = FILE_ROLE_USERS
    file_number = ordered_roles.index(role) + 1
    report(
        file_progress(file_number, 0.10),
        (
            f"Archivo {file_number} de {total_files} · "
            "Leyendo Diccionario Usuarios-Cargos..."
        ),
    )
    users_raw, users_bytes = read_uploaded_table(
        uploaded_files[role]
    )
    report(
        file_progress(file_number, 0.55),
        (
            f"Archivo {file_number} de {total_files} · "
            "Validando Diccionario Usuarios-Cargos..."
        ),
    )
    users_dictionary, users_report = normalize_user_dictionary(
        users_raw
    )
    source_bytes[role] = users_bytes
    completed_files += 1
    report(
        file_progress(file_number, 1.0),
        (
            f"Archivo {file_number} de {total_files} completado · "
            f"Usuarios-Cargos · {len(users_dictionary):,} registros"
        ).replace(",", "."),
    )

    # Archivo opcional de cambios.
    changes_dataframe = pd.DataFrame(columns=CHANGE_COLUMNS)
    changes_report = {
        "rows": 0,
        "empty_user": 0,
        "empty_ceco": 0,
        "empty_note": 0,
        "loaded": False,
    }

    if FILE_ROLE_CHANGES in uploaded_files:
        role = FILE_ROLE_CHANGES
        file_number = ordered_roles.index(role) + 1
        report(
            file_progress(file_number, 0.10),
            (
                f"Archivo {file_number} de {total_files} · "
                "Leyendo historial de cambios..."
            ),
        )
        changes_raw, changes_bytes = read_uploaded_table(
            uploaded_files[role]
        )
        report(
            file_progress(file_number, 0.55),
            (
                f"Archivo {file_number} de {total_files} · "
                "Validando historial de cambios..."
            ),
        )
        changes_dataframe, changes_report = normalize_changes_file(
            changes_raw
        )
        changes_report["loaded"] = True
        source_bytes[role] = changes_bytes
        completed_files += 1
        report(
            file_progress(file_number, 1.0),
            (
                f"Archivo {file_number} de {total_files} completado · "
                f"Cambios · {len(changes_dataframe):,} registros"
            ).replace(",", "."),
        )

    report(
        82,
        (
            f"{completed_files} de {total_files} archivos procesados · "
            "Reconstruyendo el flujo..."
        ),
    )
    flow, flow_report = reconstruct_flow(level_frames)

    if flow.empty:
        raise ValueError(
            "Los cinco archivos de liberadores no contienen reglas con "
            "CostCenter explícito y TipoDoc AZNB/AZSR."
        )

    report(
        91,
        (
            f"Flujo reconstruido · {len(flow):,} reglas · "
            "Aplicando diccionarios..."
        ).replace(",", "."),
    )
    flow, missing_cecos, cecos_without_plant = apply_ceco_dictionary(
        flow,
        ceco_dictionary,
    )
    flow_report["cecos_without_dictionary"] = missing_cecos
    flow_report["cecos_without_plant"] = cecos_without_plant

    report(97, "Preparando la versión activa...")
    data: dict[str, Any] = {
        "flujo": flow,
        "liberadores": level_frames,
        "liberador_1": level_frames[1],
        "liberador_2": level_frames[2],
        "liberador_3": level_frames[3],
        "liberador_4": level_frames[4],
        "liberador_5": level_frames[5],
        "dic_users": users_dictionary,
        "dic_ceco": ceco_dictionary,
        "cambios": changes_dataframe,
        "dic_rangos": pd.DataFrame(
            columns=["Orden", "Desde", "Hasta"]
        ),
    }

    validation = {
        "levels": level_reports,
        "flow": flow_report,
        "dic_ceco": ceco_report,
        "dic_users": users_report,
        "cambios": changes_report,
        "file_count": total_files,
    }

    report(
        100,
        (
            f"Carga completada · {total_files} de {total_files} "
            "archivos procesados."
        ),
    )
    return data, source_bytes, validation



# ============================================================
# EXPORTACIÓN DE LOS SIETE ARCHIVOS
# ============================================================

def dataframe_for_export(
    data: dict[str, Any],
    role: str,
) -> pd.DataFrame:
    """Obtiene una copia limpia del archivo lógico solicitado."""
    if role == FILE_ROLE_CECO:
        frame = data.get("dic_ceco", pd.DataFrame())
        columns = CECO_DICTIONARY_COLUMNS
    elif role == FILE_ROLE_USERS:
        frame = data.get("dic_users", pd.DataFrame())
        columns = USER_DICTIONARY_COLUMNS
    elif role == FILE_ROLE_CHANGES:
        frame = data.get("cambios", pd.DataFrame())
        columns = CHANGE_COLUMNS
    else:
        level = next(
            (
                current_level
                for current_level in LEVELS
                if FILE_ROLE_LIBERATORS[current_level] == role
            ),
            None,
        )
        if level is None:
            raise ValueError(f"Rol de archivo no reconocido: {role}")

        frame = data.get("liberadores", {}).get(
            level,
            pd.DataFrame(),
        )
        columns = GROUP_COLUMNS

    if not isinstance(frame, pd.DataFrame):
        raise ValueError(
            f"No hay datos disponibles para {role_label(role)}."
        )

    result = frame.drop(
        columns=[
            "_DesdeNum",
            "_HastaNum",
            "_Nivel",
            "_Liberador",
        ],
        errors="ignore",
    ).copy()

    for column in columns:
        if column not in result.columns:
            result[column] = ""

    return result.loc[:, columns].copy()


def source_extension(
    role: str,
    file_names: dict[str, str],
) -> str:
    """Determina la extensión del archivo originalmente cargado."""
    name = clean_text(file_names.get(role, ""))
    suffix = Path(name).suffix.lower()

    if suffix in {".parquet", ".pq"}:
        return ".parquet"
    if suffix == ".csv":
        return ".csv"
    if suffix in {".xlsx", ".xls", ".xlsm"}:
        return ".xlsx"

    return ".csv"


def dataframe_to_csv_bytes(dataframe: pd.DataFrame) -> bytes:
    return dataframe.to_csv(
        index=False,
        sep=";",
        lineterminator="\n",
    ).encode("utf-8-sig")


def available_parquet_engine() -> str | None:
    """Retorna el motor Parquet disponible en el entorno."""
    if importlib.util.find_spec("pyarrow") is not None:
        return "pyarrow"

    if importlib.util.find_spec("fastparquet") is not None:
        return "fastparquet"

    return None


def dataframe_to_parquet_bytes(dataframe: pd.DataFrame) -> bytes:
    engine = available_parquet_engine()

    if engine is None:
        raise ValueError(
            "La exportación Parquet requiere instalar `pyarrow` "
            "o `fastparquet`. Agrega `pyarrow` a requirements.txt."
        )

    output = BytesIO()

    try:
        dataframe.to_parquet(
            output,
            index=False,
            engine=engine,
        )
    except Exception as error:
        raise ValueError(
            f"No fue posible generar el archivo Parquet con {engine}."
        ) from error

    return output.getvalue()


def dataframe_to_excel_bytes(
    dataframe: pd.DataFrame,
) -> bytes:
    output = BytesIO()

    try:
        with pd.ExcelWriter(
            output,
            engine="openpyxl",
        ) as writer:
            dataframe.to_excel(
                writer,
                index=False,
                sheet_name="Sheet1",
            )
    except Exception as error:
        raise ValueError(
            "No fue posible generar el archivo Excel."
        ) from error

    return output.getvalue()


def export_file_name(
    role: str,
    extension: str,
) -> str:
    if role == FILE_ROLE_CECO:
        stem = "Diccionario_CECO_Plantas"
    elif role == FILE_ROLE_USERS:
        stem = "Diccionario_Usuarios_Cargos"
    elif role == FILE_ROLE_CHANGES:
        stem = "Cambios"
    else:
        level = next(
            level
            for level in LEVELS
            if FILE_ROLE_LIBERATORS[level] == role
        )
        stem = f"Liberador_{level}_Compra_Directa_ENAEX"

    return f"{stem}{extension}"


def nombre_exportacion(
    role: str,
    extension: str,
    file_names: dict[str, str],
) -> str:
    """
    Conserva el nombre con que se cargó cada archivo para poder
    reemplazarlo en SharePoint sin romper su vínculo.
    """
    original = clean_text(file_names.get(role, ""))
    stem = Path(original).stem if original else ""
    if not stem or stem.endswith("_vacio"):
        return export_file_name(role, extension)
    return f"{stem}{extension}"


def build_export_zip(
    data: dict[str, Any],
    file_names: dict[str, str],
    export_mode: str,
) -> bytes:
    """
    Genera un ZIP con los siete archivos.

    export_mode:
      - source: conserva CSV/Parquet/Excel según el archivo de entrada.
      - csv: convierte todos a CSV.
      - parquet: convierte todos a Parquet.
    """
    if export_mode not in {"source", "csv", "parquet"}:
        raise ValueError("Modo de exportación no válido.")

    output = BytesIO()

    with zipfile.ZipFile(
        output,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as archive:
        export_roles = list(REQUIRED_FILE_ROLES)
        changes = data.get("cambios", pd.DataFrame())
        if isinstance(changes, pd.DataFrame) and not changes.empty:
            export_roles.append(FILE_ROLE_CHANGES)

        for role in export_roles:
            dataframe = dataframe_for_export(data, role)

            if export_mode == "csv":
                extension = ".csv"
            elif export_mode == "parquet":
                extension = ".parquet"
            else:
                extension = source_extension(role, file_names)

            if extension == ".csv":
                content = dataframe_to_csv_bytes(dataframe)
            elif extension == ".parquet":
                content = dataframe_to_parquet_bytes(dataframe)
            elif extension == ".xlsx":
                content = dataframe_to_excel_bytes(dataframe)
            else:
                raise ValueError(
                    f"Extensión de exportación no válida: {extension}"
                )

            archive.writestr(
                nombre_exportacion(role, extension, file_names),
                content,
            )

    return output.getvalue()


def render_export_section(data: dict[str, Any]) -> None:
    st.markdown("---")
    st.markdown(
        '<div class="fl-section-title">Descargar versión completa</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Descarga los siete archivos conservando su formato de origen "
        "o conviértelos todos a CSV o Parquet."
    )

    file_names = st.session_state.get(SESSION_FILE_KEY, {})
    if not isinstance(file_names, dict):
        file_names = {}

    input_extensions = {
        source_extension(role, file_names)
        for role in REQUIRED_FILE_ROLES
    }

    if input_extensions == {".parquet"}:
        default_index = 2
    elif input_extensions == {".csv"}:
        default_index = 1
    else:
        default_index = 0

    parquet_engine = available_parquet_engine()
    options = ["source", "csv"]

    if parquet_engine is not None:
        options.append("parquet")
    elif ".parquet" in input_extensions:
        st.error(
            "Se cargaron archivos Parquet, pero el entorno no tiene un motor "
            "Parquet instalado. Agrega `pyarrow` a requirements.txt."
        )
        return
    else:
        st.info(
            "La opción Parquet se habilitará automáticamente al instalar "
            "`pyarrow` o `fastparquet`."
        )

    default_mode = (
        "parquet"
        if default_index == 2 and "parquet" in options
        else "csv"
        if default_index == 1
        else "source"
    )

    selected_mode = st.radio(
        "Formato de descarga",
        options=options,
        index=options.index(default_mode),
        format_func=lambda value: {
            "source": "Conservar formato original de cada archivo",
            "csv": "Convertir los 7 archivos a CSV",
            "parquet": (
                "Convertir los 7 archivos a Parquet "
                f"({parquet_engine})"
            ),
        }[value],
        horizontal=False,
        key="export_format_mode_v02",
    )

    try:
        zip_bytes = build_export_zip(
            data=data,
            file_names=file_names,
            export_mode=selected_mode,
        )
    except ValueError as error:
        st.error(str(error))
        return

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    mode_label = {
        "source": "FORMATO_ORIGINAL",
        "csv": "CSV",
        "parquet": "PARQUET",
    }[selected_mode]

    st.download_button(
        "⬇️ Descargar ZIP con los 7 archivos",
        data=zip_bytes,
        file_name=nombre_descarga_version(
            version_activa(file_names)
        ),
        mime="application/zip",
        type="primary",
        use_container_width=True,
        key=f"download_seven_files_{selected_mode}_v01",
    )

# ============================================================
# INTERFAZ
# ============================================================

def render_header() -> None:
    mostrar_logo()

    st.markdown(
        '<div class="fl-title">01 Cargar Versión</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="fl-subtitle">
            Carga la versión manualmente o mediante conexión SharePoint.
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_format_help() -> None:
    columns = " | ".join(GROUP_COLUMNS)

    st.markdown(
        compact_html(
            f"""
            <div class="format-card">
                <div class="format-title">Formato único desde ahora</div>
                <div class="format-columns">{columns}</div>
            </div>
            """
        ),
        unsafe_allow_html=True,
    )

    st.caption(
        "Cada nivel puede cargarse como CSV, Parquet o Excel. "
        "La columna Group puede venir vacía, pero debe existir cuando el archivo "
        "la utiliza para representar Liberacion Servicios."
    )


def detect_file_role(name: str) -> str | None:
    """Identifica los cinco niveles, dos diccionarios y cambios opcionales."""
    text = Path(name).stem.casefold()
    simplified = re.sub(r"[^a-z0-9]+", " ", text).strip()

    ceco_terms = (
        ("diccionario" in simplified or "dic" in simplified)
        and ("ceco" in simplified or "centro costo" in simplified)
        and ("planta" in simplified or "centro" in simplified)
    )
    if ceco_terms:
        return FILE_ROLE_CECO

    user_terms = (
        ("diccionario" in simplified or "dic" in simplified)
        and (
            "usuario" in simplified
            or "usuarios" in simplified
            or "correo" in simplified
        )
        and (
            "cargo" in simplified
            or "cargos" in simplified
            or "rol" in simplified
        )
    )
    if user_terms:
        return FILE_ROLE_USERS

    changes_terms = (
        "cambio" in simplified
        or "cambios" in simplified
        or "historial" in simplified
        or "auditoria" in simplified
        or "auditoría" in simplified
    )
    if changes_terms:
        return FILE_ROLE_CHANGES

    patterns = [
        r"liberador[_\-\s]*(\d)",
        r"nivel[_\-\s]*(\d)",
        r"lib[_\-\s]*(\d)",
    ]

    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            level = int(match.group(1))
            if level in LEVELS:
                return FILE_ROLE_LIBERATORS[level]

    return None


def role_label(role: str) -> str:
    if role == FILE_ROLE_CECO:
        return "Diccionario CECO-Plantas"
    if role == FILE_ROLE_USERS:
        return "Diccionario Usuarios-Cargos"
    if role == FILE_ROLE_CHANGES:
        return "Historial de Cambios (opcional)"

    for level in LEVELS:
        if role == FILE_ROLE_LIBERATORS[level]:
            return f"Liberador {level}"

    return role


def obtener_version_desde_nombre_archivo(
    nombre_archivo: str,
) -> dict[str, Any]:
    """
    Detecta versiones con fecha y hora en formatos:
    - DDMMYYYY_HHMMSS
    - YYYYMMDD_HHMMSS
    """
    stem = Path(clean_text(nombre_archivo)).stem

    patterns = [
        (r"(?<!\d)(\d{2})(\d{2})(\d{4})[_-](\d{2})(\d{2})(\d{2})(?!\d)", "DMY"),
        (r"(?<!\d)(\d{4})(\d{2})(\d{2})[_-](\d{2})(\d{2})(\d{2})(?!\d)", "YMD"),
    ]

    for pattern, order in patterns:
        matches = list(re.finditer(pattern, stem))
        if not matches:
            continue

        match = matches[-1]
        values = [int(value) for value in match.groups()]

        try:
            if order == "DMY":
                day, month, year, hour, minute, second = values
            else:
                year, month, day, hour, minute, second = values

            version_date = datetime(
                year,
                month,
                day,
                hour,
                minute,
                second,
            )
        except ValueError:
            continue

        return {
            "version_detectada": True,
            "fecha_version": version_date,
            "texto_version": version_date.strftime(
                "%d-%m-%Y %H:%M:%S"
            ),
            "fecha_corta": version_date.strftime("%d-%m-%Y"),
        }

    return {
        "version_detectada": False,
        "fecha_version": None,
        "texto_version": "No detectada",
        "fecha_corta": "No detectada",
    }


def obtener_version_archivos_activos(
    file_names: dict[str, str],
) -> dict[str, Any]:
    """Obtiene la versión común o más reciente desde los nombres cargados."""
    detected: list[datetime] = []

    for name in file_names.values():
        version = obtener_version_desde_nombre_archivo(name)
        value = version.get("fecha_version")
        if isinstance(value, datetime):
            detected.append(value)

    if not detected:
        return {
            "version_detectada": False,
            "fecha_version": None,
            "texto_version": "No detectada",
            "fecha_corta": "No detectada",
        }

    # En una versión exportada todos los archivos normalmente comparten sello.
    # Si hay diferencias, se muestra el más reciente.
    version_date = max(detected)
    return {
        "version_detectada": True,
        "fecha_version": version_date,
        "texto_version": version_date.strftime("%d-%m-%Y %H:%M:%S"),
        "fecha_corta": version_date.strftime("%d-%m-%Y"),
    }


# ============================================================
# VERSIÓN POR FECHA
# ============================================================

def ahora_chile() -> datetime:
    return datetime.now(CHILE_TZ).replace(tzinfo=None)


def fecha_larga(fecha: datetime) -> str:
    return f"{fecha.day} de {MESES_ES[fecha.month - 1]} de {fecha.year}"


def construir_info_version(
    fecha: datetime,
    origen: str,
) -> dict[str, Any]:
    return {
        "version_detectada": True,
        "fecha_version": fecha,
        "texto_version": f"{fecha_larga(fecha)} · {fecha:%H:%M}",
        "fecha_corta": fecha.strftime("%d-%m-%Y"),
        "origen": origen,
    }


def fecha_desde_last_modified(valor: str) -> datetime | None:
    try:
        fecha = parsedate_to_datetime(clean_text(valor))
    except (TypeError, ValueError, IndexError):
        return None
    if fecha is None:
        return None
    if fecha.tzinfo is not None:
        fecha = fecha.astimezone(CHILE_TZ).replace(tzinfo=None)
    return fecha


def determinar_version_carga(
    uploaded_files: dict[str, Any],
) -> dict[str, Any]:
    """
    Fecha de la versión, en este orden de prioridad:
    1. Sello de fecha en el nombre de los archivos.
    2. Última modificación informada por SharePoint.
    3. Fecha y hora de la carga.
    """
    guardadas = [
        fecha
        for uploaded in uploaded_files.values()
        if isinstance(
            fecha := getattr(uploaded, "fecha_guardado", None),
            datetime,
        )
    ]
    if guardadas:
        return construir_info_version(
            max(guardadas),
            "cambios guardados en la app",
        )

    names = {
        role: getattr(uploaded, "name", "")
        for role, uploaded in uploaded_files.items()
    }
    por_nombre = obtener_version_archivos_activos(names)
    if por_nombre.get("version_detectada"):
        return construir_info_version(
            por_nombre["fecha_version"],
            "fecha indicada en el nombre de los archivos",
        )

    fechas = [
        fecha
        for uploaded in uploaded_files.values()
        if (
            fecha := fecha_desde_last_modified(
                getattr(uploaded, "last_modified", "")
            )
        )
        is not None
    ]
    if fechas:
        return construir_info_version(
            max(fechas),
            "última modificación de los archivos en SharePoint",
        )

    return construir_info_version(ahora_chile(), "fecha de carga")


def version_activa(file_names: dict[str, str]) -> dict[str, Any]:
    guardada = st.session_state.get(SESSION_VERSION_KEY)
    if isinstance(guardada, dict) and guardada.get("fecha_version"):
        return guardada

    por_nombre = obtener_version_archivos_activos(file_names)
    if por_nombre.get("version_detectada"):
        return construir_info_version(
            por_nombre["fecha_version"],
            "fecha indicada en el nombre de los archivos",
        )

    info = construir_info_version(ahora_chile(), "fecha de carga")
    st.session_state[SESSION_VERSION_KEY] = info
    return info


def nombre_descarga_version(
    info: dict[str, Any],
    extension: str = ".zip",
) -> str:
    fecha = info.get("fecha_version")
    if not isinstance(fecha, datetime):
        fecha = ahora_chile()
    return f"Version {fecha_larga(fecha)} - {NOMBRE_DESCARGA}{extension}"


# ============================================================
# EDICIÓN DE USUARIOS (USER) EN TIEMPO REAL
# ============================================================

def cambios_vigentes(data: dict[str, Any]) -> pd.DataFrame:
    """Historial de cambios actual, incluido lo registrado en la página 03."""
    history_03 = st.session_state.get(HISTORY_KEY_PAGE_03)
    previous = data.get("cambios", pd.DataFrame())
    if isinstance(history_03, list) and len(history_03) > (
        len(previous) if isinstance(previous, pd.DataFrame) else 0
    ):
        previous = pd.DataFrame(history_03).reindex(columns=CHANGE_COLUMNS)
    elif isinstance(previous, pd.DataFrame) and not previous.empty:
        previous = dataframe_for_export(data, FILE_ROLE_CHANGES)
    else:
        previous = pd.DataFrame(columns=CHANGE_COLUMNS)
    return previous.fillna("").reset_index(drop=True)


@st.cache_resource(show_spinner=False)
def almacen_version_guardada() -> dict[str, Any]:
    """
    Versión guardada compartida por todas las sesiones de la app.
    Sobrevive a recargar la página (F5) y la ven todos los usuarios,
    hasta que la app se reinicia (por ejemplo, al actualizar el código).
    """
    return {"lock": threading.Lock(), "fecha": None, "zip": None}


def obtener_version_guardada() -> tuple[datetime | None, bytes | None]:
    almacen = almacen_version_guardada()
    with almacen["lock"]:
        return almacen["fecha"], almacen["zip"]


def guardar_version_en_app(fecha: datetime, zip_bytes: bytes) -> None:
    almacen = almacen_version_guardada()
    with almacen["lock"]:
        almacen["fecha"] = fecha
        almacen["zip"] = bytes(zip_bytes)


def descartar_version_guardada() -> None:
    almacen = almacen_version_guardada()
    with almacen["lock"]:
        almacen["fecha"] = None
        almacen["zip"] = None


def fecha_archivos_sharepoint(
    uploaded_files: dict[str, Any],
) -> datetime | None:
    fechas = [
        fecha
        for uploaded in uploaded_files.values()
        if (
            fecha := fecha_desde_last_modified(
                getattr(uploaded, "last_modified", "")
            )
        )
        is not None
    ]
    return max(fechas) if fechas else None


def archivos_desde_version_guardada(
    fecha: datetime,
    zip_bytes: bytes,
) -> dict[str, RemoteUploadedFile]:
    archivos: dict[str, RemoteUploadedFile] = {}
    with zipfile.ZipFile(BytesIO(zip_bytes)) as archive:
        for name in archive.namelist():
            role = detect_file_role(name)
            if role:
                archivos[role] = RemoteUploadedFile(
                    name,
                    archive.read(name),
                    fecha_guardado=fecha,
                )
    return archivos


def marcar_cambios_guardados(total_cambios: int, zip_bytes: bytes) -> None:
    hoy = ahora_chile()
    guardar_version_en_app(hoy, zip_bytes)
    st.session_state[SESSION_VERSION_KEY] = construir_info_version(
        hoy,
        "cambios guardados en la app",
    )
    st.session_state[SESSION_SAVED_COUNT_KEY] = total_cambios
    st.toast(
        f"Versión {hoy:%d-%m-%Y} guardada en la app. Para respaldarla "
        "también en OneDrive, reemplaza los archivos con los del ZIP.",
        icon="💾",
    )


def render_fecha_version(version: dict[str, Any]) -> None:
    """Muestra la fecha completa de la versión, sin recortarla."""
    fecha = version.get("fecha_version")
    if isinstance(fecha, datetime):
        fecha_texto = fecha.strftime("%d-%m-%Y")
        hora_texto = fecha.strftime("%H:%M")
    else:
        fecha_texto = clean_text(version.get("fecha_corta")) or "—"
        hora_texto = ""

    st.markdown(
        f"""
        <div style="line-height:1.15; padding-top:2px;">
            <div style="font-size:0.875rem; opacity:0.8;
                        margin-bottom:6px;">Versión</div>
            <div style="font-size:clamp(1.35rem, 2.1vw, 2.1rem);
                        font-weight:400; white-space:nowrap;
                        overflow:visible;">{fecha_texto}</div>
            <div style="font-size:0.8rem; opacity:0.65;
                        margin-top:4px;">{hora_texto}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_boton_guardar(data: dict[str, Any]) -> None:
    file_names = st.session_state.get(SESSION_FILE_KEY, {})
    if not isinstance(file_names, dict):
        file_names = {}

    cambios = cambios_vigentes(data)
    total = len(cambios)
    if SESSION_SAVED_COUNT_KEY not in st.session_state:
        st.session_state[SESSION_SAVED_COUNT_KEY] = total
    pendientes = max(0, total - st.session_state[SESSION_SAVED_COUNT_KEY])

    export_data = dict(data)
    export_data["cambios"] = cambios

    try:
        zip_bytes = build_export_zip(
            data=export_data,
            file_names=file_names,
            export_mode="source",
        )
    except ValueError as error:
        st.error(str(error))
        return

    hoy = ahora_chile()
    st.download_button(
        f"Guardar cambios · {hoy:%d-%m-%Y}",
        data=zip_bytes,
        file_name=nombre_descarga_version(
            construir_info_version(hoy, "cambios guardados")
        ),
        mime="application/zip",
        icon="💾",
        type="primary" if pendientes else "secondary",
        use_container_width=True,
        on_click=marcar_cambios_guardados,
        args=(total, zip_bytes),
        key="guardar_cambios_version_v01",
        help=(
            "Descarga la versión con la fecha de hoy y deja la app en esa "
            "versión. Para que sea permanente, reemplaza los archivos en "
            "OneDrive con los del ZIP (mismo nombre)."
        ),
    )
    if pendientes:
        st.caption(f"⚠️ {pendientes} cambio(s) sin guardar")
    else:
        st.caption("✅ Sin cambios pendientes")


def reconstruir_version(
    updated: dict[str, Any],
    registros: list[dict[str, Any]],
    origen: str,
) -> None:
    """
    Reconstruye la versión completa desde los datos en memoria
    (incluye lo modificado en la página 03), con las mismas
    validaciones de la carga, y agrega los registros a Cambios.
    """
    data = st.session_state.get(SESSION_DATA_KEY)
    if not isinstance(data, dict):
        raise ValueError("No hay una versión activa.")

    original_names = dict(
        st.session_state.get(SESSION_FILE_KEY, {}) or {}
    )
    names = dict(original_names)

    sources: dict[str, bytes] = {}
    for role in REQUIRED_FILE_ROLES:
        sources[role] = (
            dataframe_for_export(updated, role)
            .to_csv(index=False)
            .encode("utf-8")
        )
        names[role] = (
            Path(names.get(role) or export_file_name(role, ".csv")).stem
            + ".csv"
        )

    previous = cambios_vigentes(data)

    all_changes = pd.concat(
        [previous, pd.DataFrame(registros, columns=CHANGE_COLUMNS)],
        ignore_index=True,
    )
    sources[FILE_ROLE_CHANGES] = (
        all_changes.to_csv(index=False).encode("utf-8")
    )
    names[FILE_ROLE_CHANGES] = (
        Path(names.get(FILE_ROLE_CHANGES) or "Cambios").stem + ".csv"
    )

    uploaded = {
        current_role: RemoteUploadedFile(
            names.get(current_role) or export_file_name(current_role, ".csv"),
            raw,
        )
        for current_role, raw in sources.items()
    }

    new_data, new_sources, validation = load_seven_files(uploaded)

    st.session_state[SESSION_DATA_KEY] = new_data
    # Se conservan los nombres originales (y su extensión) para que la
    # descarga pueda reemplazar los archivos en SharePoint.
    st.session_state[SESSION_FILE_KEY] = {
        current_role: (
            original_names.get(current_role) or uploaded_file.name
        )
        for current_role, uploaded_file in uploaded.items()
    }
    st.session_state[SESSION_SOURCE_FILES_KEY] = new_sources
    st.session_state[SESSION_FILE_BYTES_KEY] = new_sources
    st.session_state[SESSION_VALIDATION_KEY] = validation
    st.session_state[SESSION_CHANGES_KEY] = new_data.get(
        "cambios",
        pd.DataFrame(columns=CHANGE_COLUMNS),
    )
    st.session_state[SESSION_VERSION_KEY] = construir_info_version(
        ahora_chile(),
        origen,
    )
    st.session_state.pop(SESSION_CASE_KEY, None)


def registro_cambio(
    campo: str,
    antes: str,
    despues: str,
    nota: str,
    *,
    ceco: str = "",
    desde: str = "",
    hasta: str = "",
    tipo_doc: str = "",
) -> dict[str, Any]:
    return {
        "FechaHora": ahora_chile().strftime("%d-%m-%Y %H:%M:%S"),
        "Usuario": "Edición directa en la app",
        "CECO": ceco,
        "Desde": desde,
        "Hasta": hasta,
        "TipoDoc": tipo_doc,
        "Campo": campo,
        "ValorAntes": antes,
        "ValorDespues": despues,
        "Nota": nota,
    }


def aplicar_edicion_usuarios(
    level: int,
    ediciones: list[tuple[Any, str, str]],
) -> None:
    """Aplica cambios de la columna User de un liberador."""
    data = st.session_state.get(SESSION_DATA_KEY)
    if not isinstance(data, dict):
        raise ValueError("No hay una versión activa.")

    liberators = dict(data.get("liberadores", {}))
    frame = liberators[level].copy()
    registros: list[dict[str, Any]] = []

    for index, antes, despues in ediciones:
        row = frame.loc[index]
        frame.at[index, "User"] = despues
        registros.append(
            registro_cambio(
                f"Liberador {level}",
                antes,
                despues,
                "Cambio de usuario liberador",
                ceco=clean_text(row.get("CostCenter", "")),
                desde=clean_text(row.get("TotalCost Bajo", "")),
                hasta=clean_text(row.get("TotalCost Alto", "")),
                tipo_doc=clean_text(row.get("cus_POClasedeDocumento", "")),
            )
        )

    liberators[level] = frame
    updated = dict(data)
    updated["liberadores"] = liberators
    reconstruir_version(updated, registros, "edición en la app")


def usuarios_en_liberadores(data: dict[str, Any]) -> set[str]:
    usados: set[str] = set()
    for frame in data.get("liberadores", {}).values():
        if isinstance(frame, pd.DataFrame) and "User" in frame.columns:
            usados.update(
                frame["User"].astype(str).str.strip().str.lower()
            )
    return usados


def aplicar_diccionario_usuarios(
    nuevo: pd.DataFrame,
    registros: list[dict[str, Any]],
) -> None:
    data = st.session_state.get(SESSION_DATA_KEY)
    if not isinstance(data, dict):
        raise ValueError("No hay una versión activa.")
    updated = dict(data)
    updated["dic_users"] = nuevo.reset_index(drop=True)
    reconstruir_version(updated, registros, "edición en la app")


def comparar_diccionarios_usuarios(
    actual: dict[str, str],
    nuevo: dict[str, str],
) -> list[dict[str, Any]]:
    """Genera los registros de Cambios entre dos diccionarios correo→cargo."""
    registros: list[dict[str, Any]] = []
    for correo in sorted(set(nuevo) - set(actual)):
        registros.append(
            registro_cambio(
                "Usuarios-Cargos",
                "",
                f"{correo} · {nuevo[correo]}",
                "Usuario agregado",
            )
        )
    for correo in sorted(set(actual) - set(nuevo)):
        registros.append(
            registro_cambio(
                "Usuarios-Cargos",
                f"{correo} · {actual[correo]}",
                "",
                "Usuario eliminado",
            )
        )
    for correo in sorted(set(actual) & set(nuevo)):
        if actual[correo] != nuevo[correo]:
            registros.append(
                registro_cambio(
                    "Usuarios-Cargos",
                    f"{correo} · {actual[correo]}",
                    f"{correo} · {nuevo[correo]}",
                    "Cargo actualizado",
                )
            )
    return registros


def mapa_usuarios(frame: pd.DataFrame) -> dict[str, str]:
    """Correo (minúsculas) → Cargo, conservando el último valor."""
    resultado: dict[str, str] = {}
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        return resultado
    for correo, cargo in zip(frame["Correo"], frame["Cargo"]):
        clave = clean_text(correo).lower()
        if clave:
            resultado[clave] = clean_text(cargo)
    return resultado


def frame_desde_mapa(mapa: dict[str, str]) -> pd.DataFrame:
    return pd.DataFrame(
        [{"Correo": correo, "Cargo": cargo} for correo, cargo in mapa.items()],
        columns=USER_DICTIONARY_COLUMNS,
    )


def guardar_mensaje(tipo: str, texto: str) -> None:
    st.session_state[SESSION_EDITOR_MESSAGE_KEY] = (tipo, texto)


def render_editor_usuarios(data: dict[str, Any]) -> None:
    dictionary = data.get("dic_users", pd.DataFrame())
    if not isinstance(dictionary, pd.DataFrame):
        dictionary = pd.DataFrame(columns=USER_DICTIONARY_COLUMNS)
    dictionary = dictionary.loc[:, USER_DICTIONARY_COLUMNS].copy()
    actual = mapa_usuarios(dictionary)
    usados = usuarios_en_liberadores(data)

    sin_registrar = sorted(
        correo
        for correo in usados - set(actual)
        if EMAIL_PATTERN.match(correo)
    )
    sin_cargo = sum(1 for cargo in actual.values() if not cargo)

    columnas = st.columns(3)
    columnas[0].metric("Usuarios", len(actual))
    columnas[1].metric("Sin cargo", sin_cargo)
    columnas[2].metric(
        "En liberadores sin registrar",
        len(sin_registrar),
        help="Correos que aparecen en los liberadores pero no en este diccionario.",
    )

    if sin_registrar:
        with st.expander(
            f"Agregar los {len(sin_registrar)} correos de los liberadores "
            "que no están registrados"
        ):
            st.write(", ".join(sin_registrar))
            if st.button(
                "➕ Agregarlos al diccionario (cargo por completar)",
                key="agregar_sin_registrar_v01",
            ):
                nuevo = dict(actual)
                for correo in sin_registrar:
                    nuevo[correo] = ""
                try:
                    aplicar_diccionario_usuarios(
                        frame_desde_mapa(nuevo),
                        comparar_diccionarios_usuarios(actual, nuevo),
                    )
                    guardar_mensaje(
                        "success",
                        f"Se agregaron {len(sin_registrar)} usuarios. "
                        "Completa su cargo en la tabla.",
                    )
                except (ValueError, KeyError) as error:
                    guardar_mensaje("error", f"No se pudo agregar: {error}")
                st.session_state[SESSION_EDITOR_COUNTER_KEY] = (
                    st.session_state.get(SESSION_EDITOR_COUNTER_KEY, 0) + 1
                )
                st.rerun()

    st.caption(
        "Edita el **Cargo** con doble clic. Para **agregar** un usuario usa "
        "la fila vacía al final de la tabla (➕) y completa Correo y Cargo. "
        "Para **eliminar**, selecciona la fila y presiona el ícono de papelera. "
        "Cada cambio se aplica al instante y queda en Cambios."
    )

    counter = st.session_state.get(SESSION_EDITOR_COUNTER_KEY, 0)
    edited = st.data_editor(
        dictionary,
        key=f"editor_usuarios_{counter}",
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Correo": st.column_config.TextColumn(
                "Correo ✏️",
                help="Correo del usuario, por ejemplo nombre.apellido@enaex.com",
            ),
            "Cargo": st.column_config.TextColumn(
                "Cargo ✏️",
                help="Cargo del usuario.",
            ),
        },
    )

    filas_originales = [
        (clean_text(c).lower(), clean_text(g))
        for c, g in zip(dictionary["Correo"], dictionary["Cargo"])
    ]
    filas_editadas = [
        (clean_text(c).lower(), clean_text(g))
        for c, g in zip(edited["Correo"].tolist(), edited["Cargo"].tolist())
    ]
    if filas_editadas == filas_originales:
        render_carga_masiva_usuarios(actual, counter)
        return

    conteo_original: dict[str, int] = {}
    for correo, _ in filas_originales:
        conteo_original[correo] = conteo_original.get(correo, 0) + 1
    conteo_editado: dict[str, int] = {}

    nuevo: dict[str, str] = {}
    errores: list[str] = []
    pendientes = 0

    for correo_raw, cargo_raw in zip(
        edited["Correo"].tolist(),
        edited["Cargo"].tolist(),
    ):
        correo = clean_text(correo_raw).lower()
        cargo = clean_text(cargo_raw)

        if not correo:
            if cargo:
                pendientes += 1
            continue
        conteo_editado[correo] = conteo_editado.get(correo, 0) + 1
        # Solo se valida lo nuevo: los datos que ya venían en el
        # diccionario se conservan aunque tengan errores.
        if correo not in actual and not EMAIL_PATTERN.match(correo):
            errores.append(f"«{correo}» no es un correo válido")
            continue
        if correo in nuevo:
            if conteo_editado[correo] > conteo_original.get(correo, 0):
                errores.append(f"«{correo}» está repetido")
            continue
        if correo not in actual and not cargo:
            # Usuario nuevo: se espera a que tenga cargo.
            pendientes += 1
            continue
        nuevo[correo] = cargo

    # Los usuarios nuevos que esperan cargo no cuentan como eliminados.
    eliminados_en_uso = sorted(
        (set(actual) - set(nuevo)) & usados
        - {
            clean_text(c).lower()
            for c in edited["Correo"].tolist()
        }
    )
    if eliminados_en_uso:
        errores.append(
            "no se puede eliminar porque está asignado en los liberadores: "
            + ", ".join(eliminados_en_uso)
        )

    if errores:
        guardar_mensaje(
            "error",
            "No se aplicó el cambio: " + "; ".join(errores) + ".",
        )
        st.session_state[SESSION_EDITOR_COUNTER_KEY] = counter + 1
        st.rerun()

    # Correos vigentes que siguen en la tabla pero quedaron pendientes
    # no deben borrarse.
    for correo in actual:
        if correo not in nuevo and correo in {
            clean_text(c).lower() for c in edited["Correo"].tolist()
        }:
            nuevo[correo] = actual[correo]

    registros = comparar_diccionarios_usuarios(actual, nuevo)

    if pendientes and not registros:
        st.info(
            "Completa **Correo y Cargo** de la fila nueva para agregar "
            "el usuario."
        )

    if registros:
        try:
            aplicar_diccionario_usuarios(frame_desde_mapa(nuevo), registros)
        except (ValueError, KeyError) as error:
            guardar_mensaje("error", f"No se pudo aplicar el cambio: {error}")
        else:
            agregados = sum(
                1 for r in registros if r["Nota"] == "Usuario agregado"
            )
            eliminados = sum(
                1 for r in registros if r["Nota"] == "Usuario eliminado"
            )
            cargos = sum(
                1 for r in registros if r["Nota"] == "Cargo actualizado"
            )
            partes = [
                texto
                for cantidad, texto in [
                    (agregados, f"{agregados} agregado(s)"),
                    (cargos, f"{cargos} cargo(s) actualizado(s)"),
                    (eliminados, f"{eliminados} eliminado(s)"),
                ]
                if cantidad
            ]
            guardar_mensaje(
                "success",
                "Usuarios-Cargos: " + ", ".join(partes) + ".",
            )
        st.session_state[SESSION_EDITOR_COUNTER_KEY] = counter + 1
        st.rerun()

    render_carga_masiva_usuarios(actual, counter)


def render_carga_masiva_usuarios(
    actual: dict[str, str],
    counter: int,
) -> None:
    with st.expander("📥 Cargar varios usuarios desde un archivo"):
        st.caption(
            "Sube un Excel o CSV con columnas **Correo** y **Cargo** "
            "(también se aceptan Email, Mail, Usuario, Rol). Los correos "
            "nuevos se agregan y, si un correo ya existe, se actualiza "
            "su cargo. No se elimina a nadie."
        )
        archivo = st.file_uploader(
            "Archivo de usuarios",
            type=["csv", "xlsx", "xls", "xlsm"],
            key=f"carga_usuarios_{counter}",
        )
        if archivo is not None:
            try:
                raw_frame, _ = read_uploaded_table(archivo)
                cargados, _ = normalize_user_dictionary(raw_frame)
            except ValueError as error:
                st.error(str(error))
                return

            nuevos_archivo = mapa_usuarios(cargados)
            invalidos = [
                c for c in nuevos_archivo if not EMAIL_PATTERN.match(c)
            ]
            combinado = dict(actual)
            for correo, cargo in nuevos_archivo.items():
                if correo in invalidos:
                    continue
                if correo not in combinado or (
                    cargo and cargo != combinado[correo]
                ):
                    combinado[correo] = cargo or combinado.get(correo, "")

            registros_archivo = comparar_diccionarios_usuarios(
                actual,
                combinado,
            )
            agregados = sum(
                1 for r in registros_archivo
                if r["Nota"] == "Usuario agregado"
            )
            actualizados = len(registros_archivo) - agregados

            if invalidos:
                st.warning(
                    "Se omitirán por no ser correos válidos: "
                    + ", ".join(invalidos)
                )

            if not registros_archivo:
                st.info("El archivo no trae usuarios nuevos ni cargos distintos.")
            elif st.button(
                f"Aplicar: {agregados} nuevo(s), "
                f"{actualizados} cargo(s) actualizado(s)",
                type="primary",
                key=f"aplicar_carga_usuarios_{counter}",
            ):
                try:
                    aplicar_diccionario_usuarios(
                        frame_desde_mapa(combinado),
                        registros_archivo,
                    )
                    guardar_mensaje(
                        "success",
                        f"Usuarios-Cargos: {agregados} agregado(s) y "
                        f"{actualizados} cargo(s) actualizado(s) "
                        "desde archivo.",
                    )
                except (ValueError, KeyError) as error:
                    guardar_mensaje("error", f"No se pudo aplicar: {error}")
                st.session_state[SESSION_EDITOR_COUNTER_KEY] = counter + 1
                st.rerun()


def render_editor_liberador(
    level: int,
    frame: pd.DataFrame,
    known_users: set[str],
) -> None:
    visible = frame.drop(
        columns=["_DesdeNum", "_HastaNum", "_Nivel", "_Liberador"],
        errors="ignore",
    )

    if visible.empty:
        st.info(f"Liberador {level} no tiene reglas en esta versión.")
        return

    st.caption(
        "Haz doble clic en una celda de **User** para cambiar el "
        "liberador. El cambio se aplica al instante en toda la app "
        "y queda registrado en Cambios."
    )

    counter = st.session_state.get(SESSION_EDITOR_COUNTER_KEY, 0)
    edited = st.data_editor(
        visible,
        key=f"editor_liberador_{level}_{counter}",
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        disabled=[column for column in visible.columns if column != "User"],
        column_config={
            "User": st.column_config.TextColumn(
                "User ✏️",
                help="Correo del liberador. Editable.",
            )
        },
    )

    ediciones: list[tuple[Any, str, str]] = []
    invalidos: list[str] = []
    for index in visible.index:
        antes = clean_text(visible.at[index, "User"])
        despues = clean_text(edited.at[index, "User"]).lower()
        if despues == antes.lower():
            continue
        if not EMAIL_PATTERN.match(despues):
            invalidos.append(despues or "(vacío)")
            continue
        ediciones.append((index, antes, despues))

    if invalidos:
        st.session_state[SESSION_EDITOR_COUNTER_KEY] = counter + 1
        st.session_state[SESSION_EDITOR_MESSAGE_KEY] = (
            "error",
            "No se aplicó el cambio: "
            + ", ".join(invalidos)
            + " no es un correo válido.",
        )
        st.rerun()

    if not ediciones:
        return

    try:
        aplicar_edicion_usuarios(level, ediciones)
    except (ValueError, KeyError) as error:
        st.session_state[SESSION_EDITOR_MESSAGE_KEY] = (
            "error",
            f"No se pudo aplicar el cambio: {error}",
        )
    else:
        nuevos = sorted(
            {despues for _, _, despues in ediciones} - known_users
        )
        mensaje = (
            f"Liberador {level}: {len(ediciones)} cambio(s) aplicado(s)."
        )
        if nuevos:
            mensaje += (
                " Atención: " + ", ".join(nuevos)
                + " no está en el diccionario Usuarios-Cargos. "
                "Agrégalo con su cargo en la pestaña Usuarios-Cargos."
            )
            st.session_state[SESSION_EDITOR_MESSAGE_KEY] = ("warning", mensaje)
        else:
            st.session_state[SESSION_EDITOR_MESSAGE_KEY] = ("success", mensaje)

    st.session_state[SESSION_EDITOR_COUNTER_KEY] = counter + 1
    st.rerun()


def render_remote_connection() -> dict[str, Any]:
    st.caption(
        "Los enlaces y las credenciales técnicas se leen desde "
        "`st.secrets[\"sharepoint_flujo\"]`."
    )

    aviso = st.session_state.pop("flujo_liberacion_remote_msg_v01", None)
    if aviso:
        st.info(aviso)

    with st.form("form_remote_flujo_v02"):
        access_key = st.text_input(
            "Clave de conexión",
            type="password",
            placeholder="Ingresa la clave autorizada",
        )
        connect = st.form_submit_button(
            "Conectar y cargar archivos",
            use_container_width=True,
            type="primary",
        )

    if connect:
        progress_bar = st.progress(
            0,
            text="0% · Preparando conexión...",
        )
        status = st.empty()

        def remote_progress(percent: int, message: str) -> None:
            progress_bar.progress(
                max(0, min(100, percent)),
                text=f"{percent}% · {message}",
            )
            status.caption(message)

        try:
            # Siempre se leen los archivos actuales de OneDrive.
            descargar_archivo_remoto_cache.clear()
            remote_files = conectar_y_obtener_archivos(
                access_key,
                progress_callback=remote_progress,
            )

            # Si en la app hay una versión guardada más reciente que la de
            # OneDrive, se carga esa.
            fecha_guardada, zip_guardado = obtener_version_guardada()
            fecha_onedrive = fecha_archivos_sharepoint(remote_files)
            if fecha_guardada is not None and zip_guardado:
                if fecha_onedrive is None or fecha_guardada > fecha_onedrive:
                    remote_files = archivos_desde_version_guardada(
                        fecha_guardada,
                        zip_guardado,
                    )
                else:
                    # OneDrive ya tiene algo más nuevo: manda OneDrive.
                    descartar_version_guardada()

            st.session_state[SESSION_REMOTE_FILES_KEY] = remote_files
            st.session_state[SESSION_REMOTE_SOURCE_KEY] = "sharepoint_urls"
            progress_bar.progress(
                100,
                text=(
                    f"100% · {len(remote_files)} archivos "
                    "descargados correctamente."
                ),
            )
            status.success(
                "Conexión realizada. Los archivos están listos para validar."
            )
        except requests.RequestException:
            st.session_state.pop(SESSION_REMOTE_FILES_KEY, None)
            progress_bar.progress(0, text="Conexión interrumpida.")
            status.empty()
            st.error(
                "No fue posible descargar uno de los archivos. "
                "Revisa los enlaces, permisos y conectividad."
            )
        except ValueError as error:
            st.session_state.pop(SESSION_REMOTE_FILES_KEY, None)
            progress_bar.progress(0, text="Conexión interrumpida.")
            status.empty()
            st.error(str(error))

    value = st.session_state.get(SESSION_REMOTE_FILES_KEY, {})
    if not isinstance(value, dict):
        value = {}

    usando_guardada = any(
        isinstance(getattr(f, "fecha_guardado", None), datetime)
        for f in value.values()
    )
    if usando_guardada:
        fecha = max(
            f.fecha_guardado
            for f in value.values()
            if isinstance(getattr(f, "fecha_guardado", None), datetime)
        )
        st.info(
            f"Se cargó la versión guardada en la app el "
            f"**{fecha_larga(fecha)} · {fecha:%H:%M}**, que es más reciente "
            "que los archivos de OneDrive."
        )
        if st.button(
            "Descartar la versión guardada y usar la de OneDrive",
            key="descartar_version_guardada_v01",
        ):
            descartar_version_guardada()
            clear_active_files()
            st.session_state["flujo_liberacion_remote_msg_v01"] = (
                "Versión guardada descartada. Conéctate de nuevo para "
                "cargar los archivos de OneDrive."
            )
            st.rerun()

    return value


def render_uploader() -> dict[str, Any]:
    """Carga siete archivos obligatorios y un historial opcional."""
    selected_files = st.file_uploader(
        "Seleccionar Liberador 1–4 (5 opcional), 2 diccionarios y Cambios opcional",
        type=["csv", "parquet", "pq", "xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        key="liberadores_diccionarios_uploader_v06",
        label_visibility="collapsed",
        help=(
            "Selecciona Liberador 1–5, Diccionario CECO-Plantas, "
            "Diccionario Usuarios-Cargos y, opcionalmente, Cambios."
        ),
    )

    if not selected_files:
        return {}

    if len(selected_files) > 8:
        st.error(
            "Selecciona como máximo ocho archivos: siete obligatorios "
            "y el historial de cambios opcional."
        )
        return {}

    uploaded: dict[str, Any] = {}
    unresolved: list[str] = []
    duplicated_roles: list[str] = []

    for uploaded_file in selected_files:
        role = detect_file_role(uploaded_file.name)

        if role is None:
            unresolved.append(uploaded_file.name)
            continue

        if role in uploaded:
            duplicated_roles.append(role)
            continue

        uploaded[role] = uploaded_file

    if unresolved:
        st.error(
            "No pude identificar: "
            + ", ".join(unresolved)
            + ". Revisa los nombres de los archivos."
        )

    if duplicated_roles:
        labels = ", ".join(
            role_label(role)
            for role in sorted(set(duplicated_roles))
        )
        st.error(f"Hay archivos duplicados para: {labels}.")

    if unresolved or duplicated_roles:
        return {}

    return uploaded



def render_validation_report(report: dict[str, Any]) -> None:
    level_reports = report.get("levels", {})
    flow_report = report.get("flow", {})

    with st.expander("Informe de validación", expanded=False):
        rows = []

        for level in LEVELS:
            item = level_reports.get(level, {})
            rows.append({
                "Nivel": level,
                "Filas": item.get("rows", 0),
                "Reglas CECO": item.get("explicit_rows", 0),
                "Reglas especiales": item.get("special_rows", 0),
                "Grupo Servicios": item.get("service_group_rows", 0),
                "Duplicadas": item.get("duplicate_rows", 0),
                "Sin liberador": len(item.get("empty_approver_rows", [])),
            })

        st.dataframe(
            pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True,
        )

        dictionary_rows = [
            {
                "Archivo": "Diccionario CECO-Plantas",
                "Filas": report.get("dic_ceco", {}).get("rows", 0),
                "Duplicadas": report.get("dic_ceco", {}).get(
                    "duplicate_rows", 0
                ),
                "Valores vacíos": report.get("dic_ceco", {}).get(
                    "empty_plants", 0
                ),
            },
            {
                "Archivo": "Diccionario Usuarios-Cargos",
                "Filas": report.get("dic_users", {}).get("rows", 0),
                "Duplicadas": report.get("dic_users", {}).get(
                    "duplicate_rows", 0
                ),
                "Valores vacíos": report.get("dic_users", {}).get(
                    "empty_roles", 0
                ),
            },
        ]

        changes_report = report.get("cambios", {})
        if changes_report.get("loaded"):
            dictionary_rows.append({
                "Archivo": "Historial de Cambios",
                "Filas": changes_report.get("rows", 0),
                "Duplicadas": 0,
                "Valores vacíos": (
                    changes_report.get("empty_user", 0)
                    + changes_report.get("empty_ceco", 0)
                    + changes_report.get("empty_note", 0)
                ),
            })

        st.markdown("#### Diccionarios e historial")
        st.dataframe(
            pd.DataFrame(dictionary_rows),
            use_container_width=True,
            hide_index=True,
        )

        ceco_conflicts = report.get("dic_ceco", {}).get(
            "conflicts", pd.DataFrame()
        )
        if isinstance(ceco_conflicts, pd.DataFrame) and not ceco_conflicts.empty:
            st.error(
                "Hay CECO asociados a más de una planta en el diccionario."
            )
            st.dataframe(
                ceco_conflicts,
                use_container_width=True,
                hide_index=True,
            )

        missing_cecos = flow_report.get("cecos_without_dictionary", [])
        if missing_cecos:
            st.error(
                f"Hay {len(missing_cecos)} CECO del flujo que no existen "
                "en el diccionario CECO-Plantas."
            )

        cecos_without_plant = flow_report.get("cecos_without_plant", [])
        if cecos_without_plant:
            st.warning(
                f"Hay {len(cecos_without_plant)} CECO presentes en el "
                "diccionario, pero con Planta vacía."
            )

        conflicts = flow_report.get("conflicts", pd.DataFrame())

        if isinstance(conflicts, pd.DataFrame) and not conflicts.empty:
            st.error(
                f"Se detectaron {len(conflicts)} conflictos: una misma regla "
                "tiene más de un liberador en el mismo nivel."
            )
            st.dataframe(
                conflicts,
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.success("No se detectaron conflictos de liberadores.")

        if flow_report.get("duplicate_rule_rows", 0):
            st.error(
                "La tabla interna contiene reglas CECO/TipoDoc/rango duplicadas."
            )

        if flow_report.get("empty_flow_rows", 0):
            st.warning(
                f"Hay {flow_report['empty_flow_rows']} reglas internas "
                "sin ningún liberador."
            )


def render_active_state() -> None:
    data = st.session_state.get(SESSION_DATA_KEY)

    if not isinstance(data, dict):
        st.info("Carga los siete archivos para activar la aplicación.")
        return

    flow = data.get("flujo", pd.DataFrame())
    liberators = data.get("liberadores", {})

    if not isinstance(flow, pd.DataFrame) or flow.empty:
        st.warning("No existe un flujo interno válido.")
        return

    changes = data.get("cambios", pd.DataFrame())
    changes_count = (
        len(changes)
        if isinstance(changes, pd.DataFrame)
        else 0
    )
    active_count = 8 if changes_count else 7

    file_names = st.session_state.get(SESSION_FILE_KEY, {})
    if not isinstance(file_names, dict):
        file_names = {}
    version = version_activa(file_names)

    st.success(
        (
            f"{active_count} archivos activos · "
            f"**{len(flow):,} reglas internas** · "
            f"**{flow['CECO'].nunique():,} CECO**"
            + (
                f" · **{changes_count:,} cambios históricos**"
                if changes_count
                else ""
            )
        ).replace(",", ".")
    )

    summary_metrics = st.columns([1.35, 1.35, 1, 1, 1])
    with summary_metrics[0]:
        render_fecha_version(version)
    with summary_metrics[1]:
        render_boton_guardar(data)
    summary_metrics[2].metric(
        "Archivos activos",
        active_count,
    )
    summary_metrics[3].metric(
        "Reglas internas",
        f"{len(flow):,}".replace(",", "."),
    )
    summary_metrics[4].metric(
        "CECO",
        f"{flow['CECO'].nunique():,}".replace(",", "."),
    )

    st.caption(
        f"Versión del **{version['texto_version']}** "
        f"· según {version.get('origen', 'fecha de carga')}"
    )

    metrics = st.columns(5)
    for level, column in zip(LEVELS, metrics):
        frame = liberators.get(level, pd.DataFrame())
        column.metric(
            f"Liberador {level}",
            f"{len(frame):,}".replace(",", "."),
        )

    report = st.session_state.get(SESSION_VALIDATION_KEY, {})
    if isinstance(report, dict):
        render_validation_report(report)

    with st.expander("Vista previa del flujo reconstruido", expanded=False):
        st.dataframe(
            flow.head(200),
            use_container_width=True,
            hide_index=True,
        )

    tab_labels = (
        [f"Liberador {level}" for level in LEVELS]
        + ["CECO-Plantas", "Usuarios-Cargos"]
    )
    if changes_count:
        tab_labels.append("Cambios")

    tabs = st.tabs(tab_labels)

    mensaje = st.session_state.pop(SESSION_EDITOR_MESSAGE_KEY, None)
    if mensaje:
        tipo, texto = mensaje
        getattr(st, tipo, st.info)(texto)

    users_dictionary = data.get("dic_users", pd.DataFrame())
    known_users = (
        set(
            users_dictionary["Correo"].astype(str).str.strip().str.lower()
        )
        if isinstance(users_dictionary, pd.DataFrame)
        and "Correo" in users_dictionary.columns
        else set()
    )

    for level, tab in zip(LEVELS, tabs[:5]):
        with tab:
            render_editor_liberador(
                level,
                liberators.get(level, pd.DataFrame()),
                known_users,
            )

    with tabs[5]:
        st.dataframe(
            data.get("dic_ceco", pd.DataFrame()).head(300),
            use_container_width=True,
            hide_index=True,
        )

    with tabs[6]:
        render_editor_usuarios(data)

    if changes_count:
        with tabs[7]:
            st.dataframe(
                changes.head(500),
                use_container_width=True,
                hide_index=True,
            )

    render_export_section(data)

    if st.button(
        "🗑️ Quitar la versión cargada",
        use_container_width=True,
    ):
        clear_active_files()
        st.rerun()


def main() -> None:
    aplicar_estilos()
    render_header()

    st.markdown(
        '<div class="fl-section-title">Cargar versión completa</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="fl-help">
            Carga manualmente Liberador 1–5, CECO–Plantas y
            Usuarios–Cargos, con Cambios opcional, o conecta SharePoint y clasifica cada archivo por su nombre real.
        </div>
        """,
        unsafe_allow_html=True,
    )

    load_method = st.radio(
        "Método de carga",
        options=["Carga manual", "Conexión SharePoint"],
        horizontal=True,
        key="flujo_load_method_v01",
    )

    if load_method == "Carga manual":
        st.session_state.pop(SESSION_REMOTE_FILES_KEY, None)
        uploaded_files = render_uploader()
    else:
        uploaded_files = render_remote_connection()

    # Antes de cargar los archivos obligatorios no se muestra contenido adicional.
    if not uploaded_files:
        if load_method == "Carga manual":
            st.info(
                "Los nombres deben identificar Liberador 1–4 (5 opcional), "
                "CECO-Plantas, Usuarios-Cargos y opcionalmente Cambios."
            )
        else:
            st.info(
                "Configura las URLs en Secrets, ingresa la clave autorizada "
                "y conecta los archivos remotos."
            )
        return

    uploaded_files = completar_niveles_opcionales(uploaded_files)

    missing = [
        role
        for role in REQUIRED_FILE_ROLES
        if role not in uploaded_files
    ]

    if missing:
        st.warning(
            "Faltan archivos: "
            + ", ".join(role_label(role) for role in missing)
            + "."
        )
        return

    signature = files_signature(uploaded_files)

    needs_load = (
        st.session_state.get(SESSION_SIGNATURE_KEY) != signature
        or SESSION_DATA_KEY not in st.session_state
    )

    if needs_load:
        st.markdown("#### 1. Lectura y validación de archivos")
        file_progress_bar = st.progress(
            0,
            text="0% · Preparando los archivos...",
        )
        file_status = st.empty()

        st.markdown("#### 2. Construcción del flujo")
        flow_progress_bar = st.progress(
            0,
            text="Pendiente · Esperando la validación de archivos.",
        )
        flow_status = st.empty()

        def update_progress(percent: int, message: str) -> None:
            if percent <= 80:
                file_percent = max(
                    0,
                    min(100, round(percent / 80 * 100)),
                )
                file_progress_bar.progress(
                    file_percent,
                    text=f"{file_percent}% · {message}",
                )
                file_status.caption(message)
                return

            file_progress_bar.progress(
                100,
                text="100% · Archivos leídos y validados.",
            )
            file_status.success(
                "Etapa 1 completada. Todos los archivos fueron procesados."
            )

            flow_percent = max(
                1,
                min(100, round((percent - 80) / 20 * 100)),
            )
            flow_progress_bar.progress(
                flow_percent,
                text=f"{flow_percent}% · {message}",
            )
            flow_status.caption(message)

        try:
            data, source_bytes, validation = load_seven_files(
                uploaded_files,
                progress_callback=update_progress,
            )

            names = {
                role: uploaded_file.name
                for role, uploaded_file in uploaded_files.items()
            }

            st.session_state[SESSION_DATA_KEY] = data
            st.session_state[SESSION_FILE_KEY] = names
            st.session_state[SESSION_SOURCE_FILES_KEY] = source_bytes
            st.session_state[SESSION_FILE_BYTES_KEY] = source_bytes
            st.session_state[SESSION_SIGNATURE_KEY] = signature
            st.session_state[SESSION_VALIDATION_KEY] = validation
            st.session_state[SESSION_VERSION_KEY] = (
                determinar_version_carga(uploaded_files)
            )
            st.session_state.pop(HISTORY_KEY_PAGE_03, None)
            st.session_state[SESSION_SAVED_COUNT_KEY] = len(
                data.get("cambios", pd.DataFrame())
            )
            st.session_state[SESSION_CHANGES_KEY] = data.get(
                "cambios",
                pd.DataFrame(columns=CHANGE_COLUMNS),
            )
            st.session_state.pop(SESSION_CASE_KEY, None)

            loaded_count = len(uploaded_files)
            file_progress_bar.progress(
                100,
                text=(
                    f"100% · {loaded_count} de {loaded_count} "
                    "archivos validados correctamente."
                ),
            )
            file_status.success(
                "Etapa 1 completada: archivos cargados y validados."
            )
            flow_progress_bar.progress(
                100,
                text="100% · Flujo construido y versión activada.",
            )
            flow_status.success(
                "Etapa 2 completada: flujo reconstruido correctamente."
            )
            st.toast(
                f"{loaded_count} archivos cargados correctamente.",
                icon="✅",
            )
            st.rerun()

        except (ValueError, requests.RequestException) as error:
            file_progress_bar.progress(
                0,
                text="Carga interrumpida.",
            )
            flow_progress_bar.progress(
                0,
                text="Construcción interrumpida.",
            )
            file_status.empty()
            flow_status.empty()
            clear_active_files()
            st.error(str(error))
            return

    # Solo después de una carga válida aparecen los demás elementos.
    render_active_state()


if __name__ == "__main__":
    main()
