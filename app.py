import hashlib
import io
import json
import os
import re
import sqlite3
from datetime import datetime, timezone

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Cisco Price Manager",
    page_icon="🌐",
    layout="wide",
)

APP_TITLE = "Cisco Price Manager"
DATABASE = os.getenv("CISCO_PRICE_MANAGER_DB", "cisco_price_manager.db")
CISCO_TOKEN_URL = "https://id.cisco.com/oauth2/default/v1/token"

# Part Number de los ocho equipos principales identificados en la
# cotización compartida. Se comparan como identificadores exactos.
# Para otras cotizaciones se conserva además una detección secundaria.
KNOWN_EQUIPMENT_PART_NUMBERS = {
    "C1121-4P",
    "C1121-8P",
    "C1121X-8P",
    "C8111-G2-MX",
    "C8231-E-G2",
    "C8235-E-G2",
    "C8375-E-G2",
    "C8475-G2",
}

COLUMN_ALIASES = {
    "part_number": [
        "Part Number", "PartNumber", "SKU", "Product Number",
    ],
    "description": [
        "Description", "Product Description",
    ],
    "list_price": [
        "Unit List Price", "UnitListPrice", "List Price", "ListPrice",
    ],
    "quantity": [
        "Qty", "Quantity",
    ],
    "discount": [
        "Disc (%)", "Disc(%)", "Discount", "Discount (%)",
        "Discount(%)",
    ],
    "net_price": [
        "Unit Net Price", "UnitNetPrice", "Net Unit Price",
    ],
    "extended_price": [
        "Extended Net Price", "ExtendedNetPrice", "Extended Price",
    ],
    "lead_time": [
        "Estimated Lead Time", "Estimated Lead Time (Days)",
        "EstimatedLeadTimeDays", "Lead Time Days",
    ],
}


# ============================================================
# CONFIGURACIÓN DE SECRETOS
# ============================================================

def get_setting(name, default=None):
    """Lee primero variables de entorno y después Streamlit Secrets."""
    value = os.getenv(name)
    if value not in (None, ""):
        return value

    try:
        value = st.secrets.get(name, default)
        if value not in (None, ""):
            return value
    except Exception:
        pass

    return default


# ============================================================
# BASE DE DATOS
# ============================================================

def get_db_connection():
    conn = sqlite3.connect(DATABASE, timeout=30)
    conn.execute("PRAGMA busy_timeout = 30000")
    return conn


def initialize_database():
    """
    Crea una tabla nueva y compatible para el historial.
    No elimina, renombra ni modifica la tabla price_history anterior.
    """
    conn = get_db_connection()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS price_history_v2 (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                part_number TEXT NOT NULL,
                description TEXT,
                list_price REAL,
                source TEXT,
                status TEXT,
                updated_at TEXT
            )
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_price_history_v2_part_number
            ON price_history_v2(part_number)
        """)
        conn.commit()
    finally:
        conn.close()


def save_price_history(
    part_number,
    description,
    list_price,
    source,
    status,
):
    conn = get_db_connection()
    try:
        conn.execute("""
            INSERT INTO price_history_v2 (
                part_number,
                description,
                list_price,
                source,
                status,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            str(part_number),
            str(description or ""),
            None if list_price is None else float(list_price),
            str(source or ""),
            str(status or ""),
            datetime.now(timezone.utc).isoformat(),
        ))
        conn.commit()
    finally:
        conn.close()


def load_price_history(limit=500):
    conn = get_db_connection()
    try:
        return pd.read_sql_query("""
            SELECT
                part_number AS 'Part Number',
                description AS 'Description',
                list_price AS 'List Price',
                source AS 'Source',
                status AS 'Status',
                updated_at AS 'Updated At'
            FROM price_history_v2
            ORDER BY id DESC
            LIMIT ?
        """, conn, params=(int(limit),))
    finally:
        conn.close()


# ============================================================
# UTILIDADES
# ============================================================

def normalize_text(value):
    if value is None:
        return ""
    value = str(value).strip().lower()
    value = re.sub(r"\s+", "", value)
    value = re.sub(r"[_\-()%]", "", value)
    return value


def safe_float(value, default=0.0):
    """Convierte valores numéricos y formatos comunes de Excel."""
    if value is None:
        return default

    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass

    if isinstance(value, (int, float)):
        try:
            return float(value)
        except (TypeError, ValueError, OverflowError):
            return default

    text = str(value).strip()
    if not text or text.upper() in {"N/A", "NA", "-", "---", "NONE"}:
        return default

    text = text.replace("$", "").replace("€", "").replace("£", "")
    text = text.replace("%", "").replace(" ", "")

    # Admite tanto 1,234.56 como 1.234,56.
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        pieces = text.split(",")
        if len(pieces) == 2 and len(pieces[1]) in (1, 2):
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")

    try:
        return float(text)
    except (TypeError, ValueError, OverflowError):
        return default


def find_column(headers, aliases):
    normalized_headers = {
        normalize_text(value): index
        for index, value in enumerate(headers)
        if value is not None
    }
    for alias in aliases:
        key = normalize_text(alias)
        if key in normalized_headers:
            return normalized_headers[key]
    return None


def safe_get(values, column_index, default=None):
    if column_index is None or column_index < 0:
        return default
    if column_index >= len(values):
        return default
    value = values[column_index]
    return default if value is None else value


def safe_set_cell(worksheet, row_number, column_number, value):
    """Escribe sin intentar modificar celdas secundarias combinadas."""
    if column_number is None or column_number < 1:
        return False
    cell = worksheet.cell(row=row_number, column=column_number)
    if isinstance(cell, MergedCell):
        return False
    cell.value = value
    return True


# ============================================================
# IDENTIFICACIÓN DE EQUIPOS PRINCIPALES
# ============================================================

def is_primary_equipment(part_number, description):
    """
    Primero identifica los ocho Part Number conocidos.
    Para otros archivos, aplica una detección secundaria conservadora.
    """
    sku = str(part_number or "").strip().upper()
    desc = str(description or "").strip().lower()

    if not sku:
        return False

    if sku in KNOWN_EQUIPMENT_PART_NUMBERS:
        return True

    if sku.startswith("CON-"):
        return False

    excluded_terms = (
        "cable", "power supply", "power cord", "mount kit",
        "rack", "license", "licence", "accessory", "accessories",
        "transceiver", "antenna", "bracket", "spare",
    )
    if any(term in desc for term in excluded_terms):
        return False

    # Detección secundaria para cotizaciones futuras.
    is_cisco_product_id = bool(re.match(r"^C\d{3,5}", sku))
    mentions_router = "router" in desc
    return is_cisco_product_id and mentions_router


# ============================================================
# LECTURA DE LA COTIZACIÓN
# ============================================================

def read_quote(file_bytes, sheet_name=None):
    workbook = load_workbook(
        io.BytesIO(file_bytes),
        data_only=False,
        read_only=False,
    )

    if sheet_name and sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
    else:
        worksheet = workbook.active

    header_row = None
    for row_number in range(1, worksheet.max_row + 1):
        values = [
            worksheet.cell(row_number, column).value
            for column in range(1, worksheet.max_column + 1)
        ]
        normalized = {normalize_text(value) for value in values}
        if "partnumber" in normalized and "description" in normalized:
            header_row = row_number
            break

    if header_row is None:
        raise ValueError(
            "No se encontró la fila de encabezados con las columnas "
            "'Part Number' y 'Description'. No se modificó el archivo."
        )

    headers = [
        worksheet.cell(header_row, column).value
        for column in range(1, worksheet.max_column + 1)
    ]
    columns = {
        key: find_column(headers, aliases)
        for key, aliases in COLUMN_ALIASES.items()
    }

    if columns["part_number"] is None:
        raise ValueError("No se encontró la columna 'Part Number'.")
    if columns["description"] is None:
        raise ValueError("No se encontró la columna 'Description'.")

    rows = []
    for excel_row in range(header_row + 1, worksheet.max_row + 1):
        raw_values = [
            worksheet.cell(excel_row, column).value
            for column in range(1, worksheet.max_column + 1)
        ]

        sku_value = safe_get(raw_values, columns["part_number"])
        if sku_value is None or not str(sku_value).strip():
            continue

        # Evita tratar filas de totales como productos.
        sku = str(sku_value).strip()
        description = str(
            safe_get(raw_values, columns["description"], "") or ""
        ).strip()

        quantity_value = safe_get(raw_values, columns["quantity"], 1)
        quantity = safe_float(quantity_value, 1.0)
        if quantity == 0 and quantity_value in (None, ""):
            quantity = 1.0

        row_data = {
            "Excel Row": excel_row,
            "Part Number": sku,
            "Description": description,
            "Unit List Price": safe_float(
                safe_get(raw_values, columns["list_price"])
            ),
            "Qty": quantity,
            "Disc (%)": safe_float(
                safe_get(raw_values, columns["discount"])
            ),
            "Unit Net Price": safe_float(
                safe_get(raw_values, columns["net_price"])
            ),
            "Extended Net Price": safe_float(
                safe_get(raw_values, columns["extended_price"])
            ),
            "Estimated Lead Time": safe_get(
                raw_values, columns["lead_time"], ""
            ),
            "Cisco Status": "Pendiente",
            "Price Source": "Cotización original",
            "Last Updated": "",
        }
        row_data["Is Equipment"] = is_primary_equipment(
            sku, description
        )
        rows.append(row_data)

    if not rows:
        raise ValueError("No se encontraron líneas de productos.")

    equipment_count = sum(bool(row["Is Equipment"]) for row in rows)
    if equipment_count == 0:
        raise ValueError(
            "No se identificaron equipos principales. Revisa los "
            "Part Number del archivo o actualiza la lista "
            "KNOWN_EQUIPMENT_PART_NUMBERS al inicio de app.py."
        )

    workbook.close()
    return {
        "rows": rows,
        "header_row": header_row,
        "sheet_name": worksheet.title,
        "columns": columns,
    }


# ============================================================
# CONEXIÓN CON CISCO
# ============================================================

def get_cisco_token():
    client_id = get_setting("CISCO_CLIENT_ID")
    client_secret = get_setting("CISCO_CLIENT_SECRET")

    if not client_id or not client_secret:
        raise ValueError(
            "Faltan CISCO_CLIENT_ID o CISCO_CLIENT_SECRET. "
            "Configúralos en Codespaces Secrets; no los escribas "
            "en el código ni los compartas en el chat."
        )

    response = requests.post(
        CISCO_TOKEN_URL,
        data={"grant_type": "client_credentials"},
        auth=(client_id, client_secret),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    response.raise_for_status()

    token = response.json().get("access_token")
    if not token:
        raise ValueError("Cisco no devolvió un access_token.")
    return token


def get_nested_value(data, path):
    if not path:
        return data

    current = data
    for part in str(path).split("."):
        if isinstance(current, list):
            try:
                current = current[int(part)]
            except (ValueError, IndexError, TypeError):
                return None
        elif isinstance(current, dict):
            current = current.get(part)
        else:
            return None

        if current is None:
            return None

    return current


def query_cisco_price(token, part_number):
    """
    Ejecuta la consulta GraphQL definida por el usuario con la
    documentación/esquema aprobados por Cisco.

    Secrets requeridos:
      CISCO_GRAPHQL_URL
      CISCO_PRICE_QUERY
      CISCO_PRICE_VARIABLES_TEMPLATE
      CISCO_PRICE_JSON_PATH
      CISCO_PRICE_FIELD
    El marcador {{SKU}} debe aparecer en la consulta o en la plantilla
    de variables para que cada solicitud consulte el Part Number correcto.
    """
    graphql_url = get_setting("CISCO_GRAPHQL_URL")
    graphql_query_template = get_setting("CISCO_PRICE_QUERY")
    variables_template = get_setting(
        "CISCO_PRICE_VARIABLES_TEMPLATE", "{}"
    )

    json_path = get_setting("CISCO_PRICE_JSON_PATH", "data")
    price_field = get_setting("CISCO_PRICE_FIELD", "listPrice")
    description_field = get_setting(
        "CISCO_DESCRIPTION_FIELD", "description"
    )
    lead_time_field = get_setting(
        "CISCO_LEAD_TIME_FIELD", "leadTime"
    )

    if not graphql_url or not graphql_query_template:
        raise ValueError(
            "La conexión GraphQL aún no está configurada. Faltan "
            "CISCO_GRAPHQL_URL y/o CISCO_PRICE_QUERY."
        )

    sku_json = json.dumps(str(part_number))
    query_has_marker = "{{SKU}}" in str(graphql_query_template)
    variables_has_marker = "{{SKU}}" in str(variables_template)

    if not query_has_marker and not variables_has_marker:
        raise ValueError(
            "La consulta no contiene el marcador {{SKU}}. Debe estar "
            "en CISCO_PRICE_QUERY o CISCO_PRICE_VARIABLES_TEMPLATE "
            "para consultar cada Part Number individualmente."
        )

    graphql_query = str(graphql_query_template).replace(
        "{{SKU}}", sku_json
    )
    variables_text = str(variables_template).replace(
        "{{SKU}}", sku_json
    )

    try:
        variables = json.loads(variables_text)
    except json.JSONDecodeError as error:
        raise ValueError(
            "CISCO_PRICE_VARIABLES_TEMPLATE no contiene JSON válido: "
            f"{error}"
        ) from error

    if not isinstance(variables, dict):
        raise ValueError(
            "CISCO_PRICE_VARIABLES_TEMPLATE debe ser un objeto JSON."
        )

    response = requests.post(
        graphql_url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        json={"query": graphql_query, "variables": variables},
        timeout=45,
    )
    response.raise_for_status()

    try:
        result = response.json()
    except ValueError as error:
        raise ValueError(
            "La respuesta de Cisco no es JSON válido."
        ) from error

    if not isinstance(result, dict):
        raise ValueError("La respuesta GraphQL de Cisco tiene formato inesperado.")

    if result.get("errors"):
        raise ValueError(
            "Cisco GraphQL devolvió errores: "
            + json.dumps(result["errors"], ensure_ascii=False)
        )

    product_data = get_nested_value(result, json_path)
    if isinstance(product_data, list):
        product_data = product_data[0] if product_data else None

    if not isinstance(product_data, dict):
        return None

    price_value = get_nested_value(product_data, price_field)
    if price_value is None:
        return None

    parsed_price = safe_float(price_value, default=None)
    if parsed_price is None or parsed_price < 0:
        raise ValueError(
            f"Cisco devolvió un precio inválido para {part_number}."
        )

    description = get_nested_value(product_data, description_field)
    lead_time = get_nested_value(product_data, lead_time_field)

    return {
        "part_number": str(part_number),
        "list_price": parsed_price,
        "description": description,
        "lead_time": lead_time,
    }


def refresh_equipment_prices(rows):
    equipment_rows = [row for row in rows if row.get("Is Equipment")]
    if not equipment_rows:
        raise ValueError("No hay equipos principales para consultar.")

    token = get_cisco_token()
    unique_skus = list(dict.fromkeys(
        row["Part Number"] for row in equipment_rows
    ))

    updated = 0
    errors = []

    for sku in unique_skus:
        try:
            product = query_cisco_price(token, sku)
            if not product:
                message = "Cisco no devolvió un producto/precio para este Part Number."
                errors.append(f"{sku}: {message}")
                for row in equipment_rows:
                    if row["Part Number"] == sku:
                        row["Cisco Status"] = "Sin resultado en Cisco"
                save_price_history(
                    sku, "", None, "Cisco API", message
                )
                continue

            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for row in equipment_rows:
                if row["Part Number"] != sku:
                    continue

                row["Unit List Price"] = product["list_price"]
                if product.get("description"):
                    row["Description"] = str(product["description"])
                if product.get("lead_time") is not None:
                    row["Estimated Lead Time"] = product["lead_time"]

                row["Cisco Status"] = "Actualizado desde Cisco"
                row["Price Source"] = "Cisco API"
                row["Last Updated"] = timestamp

            save_price_history(
                sku,
                product.get("description") or "",
                product["list_price"],
                "Cisco API",
                "Actualizado",
            )
            updated += 1

        except Exception as error:
            message = str(error)
            errors.append(f"{sku}: {message}")
            for row in equipment_rows:
                if row["Part Number"] == sku:
                    row["Cisco Status"] = "Error de consulta"
            try:
                save_price_history(
                    sku, "", None, "Cisco API", message
                )
            except Exception:
                pass

    return updated, errors


# ============================================================
# CÁLCULO DE PRECIOS
# ============================================================

def calculate_prices(rows):
    for row in rows:
        list_price = max(0.0, safe_float(row.get("Unit List Price")))
        quantity = max(0.0, safe_float(row.get("Qty"), 1.0))
        discount = min(100.0, max(0.0, safe_float(row.get("Disc (%)"))))

        net_price = round(list_price * (1.0 - discount / 100.0), 2)
        extended_price = round(net_price * quantity, 2)

        row["Qty"] = quantity
        row["Disc (%)"] = discount
        row["Unit Net Price"] = net_price
        row["Extended Net Price"] = extended_price

    return rows


# ============================================================
# EXPORTACIÓN DE EXCEL
# ============================================================

def export_quote(file_bytes, rows, header_row, sheet_name=None):
    """
    Conserva el libro original y actualiza únicamente las celdas
    correspondientes a las columnas encontradas por encabezado.
    Los importes se escriben como valores numéricos, sin fórmulas
    con letras de columna fijas.
    """
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=False)

    if sheet_name and sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
    else:
        worksheet = workbook.active

    headers = [
        worksheet.cell(header_row, column).value
        for column in range(1, worksheet.max_column + 1)
    ]
    columns = {
        key: find_column(headers, aliases)
        for key, aliases in COLUMN_ALIASES.items()
    }

    required = ("part_number", "description")
    missing = [name for name in required if columns[name] is None]
    if missing:
        workbook.close()
        raise ValueError(
            "No se puede exportar: faltan columnas requeridas "
            + ", ".join(missing)
            + ". El archivo original no fue modificado."
        )

    calculate_prices(rows)

    for row_data in rows:
        excel_row = int(row_data["Excel Row"])
        is_equipment = bool(row_data.get("Is Equipment"))

        # Solo cambia descripción y precio de lista de equipos
        # principales que hayan sido identificados.
        if is_equipment:
            safe_set_cell(
                worksheet, excel_row, columns["description"] + 1,
                row_data.get("Description", ""),
            )
            if columns["list_price"] is not None:
                safe_set_cell(
                    worksheet, excel_row, columns["list_price"] + 1,
                    safe_float(row_data.get("Unit List Price")),
                )
            if columns["quantity"] is not None:
                safe_set_cell(
                    worksheet, excel_row, columns["quantity"] + 1,
                    safe_float(row_data.get("Qty"), 1.0),
                )
            if columns["discount"] is not None:
                safe_set_cell(
                    worksheet, excel_row, columns["discount"] + 1,
                    safe_float(row_data.get("Disc (%)")),
                )
            if columns["lead_time"] is not None:
                lead_time = row_data.get("Estimated Lead Time")
                if lead_time not in (None, ""):
                    safe_set_cell(
                        worksheet, excel_row, columns["lead_time"] + 1,
                        lead_time,
                    )

        # Los precios netos se calculan con los índices reales de la
        # cotización. No se modifican Part Number ni otras columnas.
        if columns["net_price"] is not None:
            safe_set_cell(
                worksheet, excel_row, columns["net_price"] + 1,
                safe_float(row_data.get("Unit Net Price")),
            )

        if columns["extended_price"] is not None:
            safe_set_cell(
                worksheet, excel_row, columns["extended_price"] + 1,
                safe_float(row_data.get("Extended Net Price")),
            )

    # Mantener el recálculo automático activado al abrir el libro.
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except Exception:
        pass

    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    output.seek(0)
    return output.getvalue()


# ============================================================
# ESTADO DE LA APLICACIÓN
# ============================================================

initialize_database()

DEFAULT_STATE = {
    "quote_rows": None,
    "quote_bytes": None,
    "quote_header_row": None,
    "quote_sheet_name": None,
    "quote_name": None,
    "quote_hash": None,
    "uploader_version": 0,
    "last_refresh_message": None,
    "last_refresh_errors": [],
}

for state_key, default_value in DEFAULT_STATE.items():
    if state_key not in st.session_state:
        st.session_state[state_key] = default_value


# ============================================================
# INTERFAZ
# ============================================================

st.title("🌐 Cisco Price Manager")
st.caption(
    "Gestión de cotizaciones Cisco, actualización de precios, "
    "cálculo de importes e historial."
)
st.divider()

st.subheader("1. Cargar cotización Cisco")

uploaded_file = st.file_uploader(
    "Selecciona la cotización original de Cisco Workspace (.xlsx)",
    type=["xlsx"],
    key=f"quote_upload_{st.session_state.uploader_version}",
)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_hash = hashlib.sha256(file_bytes).hexdigest()

    if file_hash != st.session_state.quote_hash:
        try:
            parsed = read_quote(file_bytes)
            st.session_state.quote_rows = parsed["rows"]
            st.session_state.quote_bytes = file_bytes
            st.session_state.quote_header_row = parsed["header_row"]
            st.session_state.quote_sheet_name = parsed["sheet_name"]
            st.session_state.quote_name = uploaded_file.name
            st.session_state.quote_hash = file_hash
            st.session_state.last_refresh_message = None
            st.session_state.last_refresh_errors = []
            st.success(
                f"Cotización cargada correctamente: {uploaded_file.name}"
            )
        except Exception as error:
            st.error(f"No se pudo leer la cotización: {error}")

if st.session_state.quote_rows is not None:
    if st.button("Cerrar cotización actual", type="secondary"):
        for key in (
            "quote_rows", "quote_bytes", "quote_header_row",
            "quote_sheet_name", "quote_name", "quote_hash",
            "last_refresh_message",
        ):
            st.session_state[key] = None
        st.session_state.last_refresh_errors = []
        st.session_state.uploader_version += 1
        st.rerun()


# ============================================================
# PROCESAMIENTO
# ============================================================

if st.session_state.quote_rows is not None:
    rows = st.session_state.quote_rows
    calculate_prices(rows)

    equipment_rows = [row for row in rows if row.get("Is Equipment")]
    all_skus = {row["Part Number"] for row in rows}

    st.divider()
    st.subheader("2. Resumen de la cotización")

    metric1, metric2, metric3, metric4 = st.columns(4)
    metric1.metric("Equipos principales", len(equipment_rows))
    metric2.metric("Líneas de cotización", len(rows))
    metric3.metric("Part Number diferentes", len(all_skus))

    equipment_total = sum(
        safe_float(row.get("Extended Net Price"))
        for row in equipment_rows
    )
    metric4.metric("Total equipos", f"${equipment_total:,.2f}")

    st.caption(
        f"Archivo activo: {st.session_state.quote_name} | "
        f"Hoja: {st.session_state.quote_sheet_name}"
    )

    if len(equipment_rows) != len(KNOWN_EQUIPMENT_PART_NUMBERS):
        st.warning(
            f"Se identificaron {len(equipment_rows)} equipos principales. "
            "La lista de ocho Part Number conocidos se aplica de forma "
            "exacta; si esta cotización debe contener los ocho, revisa "
            "que todos estén presentes en la columna Part Number."
        )

    st.divider()
    st.subheader("3. Equipos principales identificados")
    st.info(
        "La consulta a Cisco se ejecuta únicamente para los equipos "
        "principales identificados. Los servicios CON-, licencias, "
        "cables y accesorios permanecen como líneas de la cotización."
    )

    display_columns = [
        "Excel Row",
        "Part Number",
        "Description",
        "Unit List Price",
        "Qty",
        "Disc (%)",
        "Unit Net Price",
        "Extended Net Price",
        "Estimated Lead Time",
        "Cisco Status",
    ]
    equipment_df = pd.DataFrame(equipment_rows, columns=display_columns)

    editor_key = f"equipment_editor_{st.session_state.quote_hash[:12]}"
    edited_df = st.data_editor(
        equipment_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key=editor_key,
        disabled=["Excel Row", "Part Number", "Cisco Status"],
        column_config={
            "Unit List Price": st.column_config.NumberColumn(
                "Precio de lista USD", min_value=0, format="$%.2f"
            ),
            "Qty": st.column_config.NumberColumn(
                "Cantidad", min_value=0, step=1
            ),
            "Disc (%)": st.column_config.NumberColumn(
                "Descuento (%)", min_value=0, max_value=100,
                format="%.2f"
            ),
            "Unit Net Price": st.column_config.NumberColumn(
                "Precio neto USD", format="$%.2f"
            ),
            "Extended Net Price": st.column_config.NumberColumn(
                "Precio extendido USD", format="$%.2f"
            ),
        },
    )

    edited_by_row = {
        int(item["Excel Row"]): item
        for item in edited_df.to_dict("records")
    }

    for row in rows:
        if not row.get("Is Equipment"):
            continue
        edited = edited_by_row.get(int(row["Excel Row"]))
        if edited is None:
            continue

        row["Description"] = str(edited.get("Description") or "")
        row["Unit List Price"] = max(
            0.0, safe_float(edited.get("Unit List Price"))
        )
        row["Qty"] = max(0.0, safe_float(edited.get("Qty"), 1.0))
        row["Disc (%)"] = min(
            100.0, max(0.0, safe_float(edited.get("Disc (%)")))
        )

    calculate_prices(rows)
    st.session_state.quote_rows = rows

    st.divider()
    st.subheader("4. Actualizar información desde Cisco")

    st.write(
        f"Se consultarán {len({r['Part Number'] for r in equipment_rows})} "
        "Part Number únicos de los equipos principales."
    )

    if st.button(
        f"🔄 Actualizar los {len(equipment_rows)} equipos desde Cisco",
        type="primary",
        use_container_width=True,
        disabled=(len(equipment_rows) == 0),
    ):
        try:
            with st.spinner(
                "Solicitando token y consultando los Part Number en Cisco..."
            ):
                updated, errors = refresh_equipment_prices(rows)

            calculate_prices(rows)
            st.session_state.quote_rows = rows
            st.session_state.last_refresh_message = (
                f"Consultas exitosas: {updated} de "
                f"{len({r['Part Number'] for r in equipment_rows})} "
                "Part Number únicos."
            )
            st.session_state.last_refresh_errors = errors
            st.rerun()
        except Exception as error:
            st.error(f"No se pudo iniciar la actualización de Cisco: {error}")

    if st.session_state.last_refresh_message:
        st.success(st.session_state.last_refresh_message)
    if st.session_state.last_refresh_errors:
        st.warning("Algunas consultas no pudieron completarse:")
        for error in st.session_state.last_refresh_errors:
            st.write(f"- {error}")

    st.divider()
    with st.expander("Ver todas las líneas de la cotización"):
        all_rows_df = pd.DataFrame(rows)
        columns_to_show = [
            "Excel Row", "Part Number", "Description",
            "Unit List Price", "Qty", "Disc (%)",
            "Unit Net Price", "Extended Net Price",
            "Is Equipment", "Cisco Status",
        ]
        st.dataframe(
            all_rows_df[columns_to_show],
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
    st.subheader("5. Exportar cotización actualizada")
    st.write(
        "La exportación parte del archivo original y actualiza únicamente "
        "las celdas de datos correspondientes. No renombra ni reorganiza "
        "las columnas de Cisco Workspace."
    )

    try:
        excel_output = export_quote(
            st.session_state.quote_bytes,
            rows,
            st.session_state.quote_header_row,
            st.session_state.quote_sheet_name,
        )
        original_name = st.session_state.quote_name or "cotizacion.xlsx"
        output_name = (
            original_name.rsplit(".", 1)[0] + "_actualizada.xlsx"
        )
        st.download_button(
            label="📥 Descargar cotización actualizada",
            data=excel_output,
            file_name=output_name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )
    except Exception as error:
        st.error(f"No se pudo generar el Excel: {error}")

    st.divider()
    with st.expander("Historial de consultas de precios"):
        try:
            history_df = load_price_history()
            if history_df.empty:
                st.info("Todavía no hay consultas guardadas.")
            else:
                st.dataframe(
                    history_df,
                    use_container_width=True,
                    hide_index=True,
                )
        except Exception as error:
            st.error(f"No se pudo cargar el historial: {error}")

else:
    st.info(
        "Carga una cotización Cisco en formato .xlsx para comenzar."
    )


# ============================================================
# PIE DE PÁGINA
# ============================================================

st.divider()
st.caption(
    "Cisco Price Manager | Los precios se actualizan desde Cisco "
    "únicamente cuando la API autorizada está configurada y devuelve "
    "información válida. La aplicación no inventa precios ni datos."
)
