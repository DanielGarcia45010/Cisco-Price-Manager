
import io
import os
import re
import json
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
    layout="wide"
)

APP_TITLE = "Cisco Price Manager"
DATABASE = "cisco_price_manager.db"

CISCO_TOKEN_URL = "https://id.cisco.com/oauth2/default/v1/token"

# Columnas esperadas en la cotización Cisco
COLUMN_ALIASES = {
    "part_number": [
        "partnumber",
        "part number",
        "sku",
        "productnumber"
    ],
    "description": [
        "description",
        "productdescription"
    ],
    "list_price": [
        "unitlistprice",
        "unit list price",
        "listprice"
    ],
    "quantity": [
        "qty",
        "quantity"
    ],
    "discount": [
        "disc(%)",
        "disc (%)",
        "discount",
        "discount(%)",
        "discount (%)"
    ],
    "net_price": [
        "unitnetprice",
        "unit net price",
        "netunitprice"
    ],
    "extended_price": [
        "extendednetprice",
        "extended net price",
        "extendedprice"
    ],
    "lead_time": [
        "estimatedleadtimedays",
        "estimated lead time (days)",
        "leadtimedays"
    ]
}


# ============================================================
# CONFIGURACIÓN DE SECRETOS
# ============================================================

def get_setting(name, default=None):
    """
    Lee primero variables de entorno y luego Streamlit Secrets.
    No se solicitan credenciales al usuario en la aplicación.
    """

    value = os.getenv(name)

    if value:
        return value

    try:
        value = st.secrets.get(name)
        if value:
            return value
    except Exception:
        pass

    return default


# ============================================================
# BASE DE DATOS
# ============================================================

def initialize_database():

    conn = sqlite3.connect(DATABASE)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS price_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_number TEXT NOT NULL,
            description TEXT,
            list_price REAL,
            source TEXT,
            status TEXT,
            updated_at TEXT
        )
    """)

    conn.commit()
    conn.close()


def save_price_history(
    part_number,
    description,
    list_price,
    source,
    status
):

    conn = sqlite3.connect(DATABASE)

    conn.execute("""
        INSERT INTO price_history (
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
        list_price,
        str(source),
        str(status),
        datetime.now(timezone.utc).isoformat()
    ))

    conn.commit()
    conn.close()


def load_price_history():

    conn = sqlite3.connect(DATABASE)

    df = pd.read_sql_query("""
        SELECT
            part_number AS 'Part Number',
            description AS 'Description',
            list_price AS 'List Price',
            source AS 'Source',
            status AS 'Status',
            updated_at AS 'Updated At'
        FROM price_history
        ORDER BY id DESC
        LIMIT 500
    """, conn)

    conn.close()

    return df


# ============================================================
# UTILIDADES
# ============================================================

def normalize_text(value):

    if value is None:
        return ""

    value = str(value).strip().lower()

    value = re.sub(r"\s+", "", value)

    value = value.replace("_", "")
    value = value.replace("-", "")
    value = value.replace("(", "")
    value = value.replace(")", "")
    value = value.replace("%", "")

    return value


def safe_float(value, default=0.0):

    if value is None:
        return default

    if isinstance(value, (int, float)):
        try:
            if pd.isna(value):
                return default
            return float(value)
        except Exception:
            return default

    value = str(value).strip()

    if not value or value.upper() in ("N/A", "-", "---"):
        return default

    value = value.replace("$", "")
    value = value.replace(",", "")
    value = value.replace("%", "")

    try:
        return float(value)
    except Exception:
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


def safe_get(row, column_index, default=None):

    if column_index is None:
        return default

    if column_index >= len(row):
        return default

    value = row[column_index]

    if value is None:
        return default

    return value


# ============================================================
# IDENTIFICACIÓN DE EQUIPOS PRINCIPALES
# ============================================================

def is_primary_equipment(part_number, description):
    """
    Identifica los equipos físicos principales.

    No considera como equipos independientes:
    - Servicios CON-*
    - Cables
    - Fuentes de poder
    - Licencias
    - Accesorios
    - Kits de montaje
    - Opciones ambientales

    En la cotización compartida identifica los 8 routers.
    """

    sku = str(part_number or "").strip().upper()
    desc = str(description or "").strip().lower()

    if not sku or not desc:
        return False

    # Los servicios de soporte no son equipos principales.
    if sku.startswith("CON-"):
        return False

    # El producto principal debe estar descrito como router.
    if "router" not in desc:
        return False

    # Excluir accesorios que mencionan routers en su descripción.
    excluded_terms = [
        "rack",
        "wallmount",
        "eco-friendly",
        "power supply",
        "power cord",
        "mount kit"
    ]

    if any(term in desc for term in excluded_terms):
        return False

    # Los modelos principales de esta cotización comienzan por C.
    if not re.match(r"^C\d{3,5}", sku):
        return False

    return True


# ============================================================
# LECTURA DE LA COTIZACIÓN
# ============================================================

def read_quote(file_bytes):

    workbook = load_workbook(
        io.BytesIO(file_bytes),
        data_only=False
    )

    worksheet = workbook.active

    header_row = None

    # Encontrar automáticamente la fila de encabezados.
    for row_number in range(1, worksheet.max_row + 1):

        values = [
            worksheet.cell(row_number, column).value
            for column in range(1, worksheet.max_column + 1)
        ]

        normalized = [
            normalize_text(value)
            for value in values
        ]

        if (
            "partnumber" in normalized
            and "description" in normalized
        ):
            header_row = row_number
            break

    if header_row is None:
        raise ValueError(
            "No se encontró la tabla de productos. "
            "Se requieren las columnas Part Number y Description."
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
        raise ValueError("No se encontró la columna Part Number.")

    if columns["description"] is None:
        raise ValueError("No se encontró la columna Description.")

    rows = []

    for excel_row in range(header_row + 1, worksheet.max_row + 1):

        raw_values = [
            worksheet.cell(excel_row, column).value
            for column in range(1, worksheet.max_column + 1)
        ]

        sku = safe_get(
            raw_values,
            columns["part_number"]
        )

        description = safe_get(
            raw_values,
            columns["description"],
            ""
        )

        if not sku:
            continue

        sku = str(sku).strip()

        if not sku:
            continue

        row_data = {
            "Excel Row": excel_row,
            "Part Number": sku,
            "Description": str(description or ""),
            "Unit List Price": safe_float(
                safe_get(raw_values, columns["list_price"])
            ),
            "Qty": safe_float(
                safe_get(raw_values, columns["quantity"]),
                1
            ),
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
                raw_values,
                columns["lead_time"],
                ""
            ),
            "Cisco Status": "Pendiente",
            "Price Source": "Cotización original",
            "Last Updated": ""
        }

        row_data["Is Equipment"] = is_primary_equipment(
            sku,
            description
        )

        rows.append(row_data)

    if not rows:
        raise ValueError("No se encontraron productos en la cotización.")

    equipment_count = sum(
        1 for row in rows if row["Is Equipment"]
    )

    if equipment_count == 0:
        raise ValueError(
            "No se identificaron equipos principales. "
            "Verifica los Part Number y las descripciones."
        )

    return {
        "rows": rows,
        "header_row": header_row,
        "sheet_name": worksheet.title,
        "columns": columns
    }


# ============================================================
# CONEXIÓN CON CISCO
# ============================================================

def get_cisco_token():

    client_id = get_setting("CISCO_CLIENT_ID")
    client_secret = get_setting("CISCO_CLIENT_SECRET")

    if not client_id or not client_secret:
        raise ValueError(
            "Faltan CISCO_CLIENT_ID o CISCO_CLIENT_SECRET "
            "en los Secrets de Codespaces."
        )

    response = requests.post(
        CISCO_TOKEN_URL,
        data={
            "grant_type": "client_credentials"
        },
        auth=(client_id, client_secret),
        headers={
            "Content-Type": "application/x-www-form-urlencoded"
        },
        timeout=30
    )

    response.raise_for_status()

    token_data = response.json()

    token = token_data.get("access_token")

    if not token:
        raise ValueError(
            "Cisco no devolvió un access_token."
        )

    return token


def get_nested_value(data, path):

    if not path:
        return data

    current = data

    for part in path.split("."):

        if isinstance(current, list):

            if not current:
                return None

            try:
                current = current[int(part)]
            except Exception:
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
    Consulta un Part Number usando la operación GraphQL
    configurada en los Secrets de Cisco.

    No se inventa una operación GraphQL: debe configurarse
    con la consulta y el esquema autorizados por Cisco.
    """

    graphql_url = get_setting("CISCO_GRAPHQL_URL")
    graphql_query = get_setting("CISCO_PRICE_QUERY")
    variables_template = get_setting(
        "CISCO_PRICE_VARIABLES_TEMPLATE",
        "{}"
    )

    json_path = get_setting(
        "CISCO_PRICE_JSON_PATH",
        "data"
    )

    price_field = get_setting(
        "CISCO_PRICE_FIELD",
        "listPrice"
    )

    description_field = get_setting(
        "CISCO_DESCRIPTION_FIELD",
        "description"
    )

    lead_time_field = get_setting(
        "CISCO_LEAD_TIME_FIELD",
        "leadTime"
    )

    if not graphql_url or not graphql_query:
        raise ValueError(
            "La API GraphQL de Cisco todavía no está configurada. "
            "Faltan CISCO_GRAPHQL_URL y CISCO_PRICE_QUERY."
        )

    # Reemplaza el marcador por el SKU consultado.
    variables_text = variables_template.replace(
        "{{SKU}}",
        json.dumps(str(part_number))
    )

    variables = json.loads(variables_text)

    response = requests.post(
        graphql_url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        },
        json={
            "query": graphql_query,
            "variables": variables
        },
        timeout=45
    )

    response.raise_for_status()

    result = response.json()

    if result.get("errors"):
        raise ValueError(
            "Cisco GraphQL devolvió errores: "
            + str(result["errors"])
        )

    product_data = get_nested_value(
        result,
        json_path
    )

    if isinstance(product_data, list):

        if not product_data:
            return None

        product_data = product_data[0]

    if not isinstance(product_data, dict):
        return None

    price = get_nested_value(
        product_data,
        price_field
    )

    description = get_nested_value(
        product_data,
        description_field
    )

    lead_time = get_nested_value(
        product_data,
        lead_time_field
    )

    if price is None:
        return None

    return {
        "part_number": part_number,
        "list_price": safe_float(price),
        "description": description,
        "lead_time": lead_time
    }


# ============================================================
# ACTUALIZACIÓN DE LOS 8 EQUIPOS
# ============================================================

def refresh_equipment_prices(rows):

    equipment_rows = [
        row for row in rows
        if row["Is Equipment"]
    ]

    if not equipment_rows:
        raise ValueError("No hay equipos para consultar.")

    token = get_cisco_token()

    updated = 0
    errors = []

    # Evitar consultas repetidas del mismo SKU.
    unique_skus = list(dict.fromkeys(
        row["Part Number"]
        for row in equipment_rows
    ))

    for sku in unique_skus:

        try:

            product = query_cisco_price(
                token,
                sku
            )

            if not product:
                errors.append(
                    f"{sku}: Cisco no devolvió información."
                )
                continue

            for row in equipment_rows:

                if row["Part Number"] != sku:
                    continue

                row["Unit List Price"] = product["list_price"]

                if product.get("description"):
                    row["Description"] = product["description"]

                if product.get("lead_time") is not None:
                    row["Estimated Lead Time"] = product["lead_time"]

                row["Cisco Status"] = "Actualizado desde Cisco"
                row["Price Source"] = "Cisco API"
                row["Last Updated"] = datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )

            save_price_history(
                sku,
                product.get("description"),
                product["list_price"],
                "Cisco API",
                "Actualizado"
            )

            updated += 1

        except Exception as error:

            errors.append(
                f"{sku}: {str(error)}"
            )

            for row in equipment_rows:

                if row["Part Number"] == sku:
                    row["Cisco Status"] = "Error de consulta"

            save_price_history(
                sku,
                "",
                None,
                "Cisco API",
                str(error)
            )

    return updated, errors


# ============================================================
# CÁLCULO DE PRECIOS
# ============================================================

def calculate_prices(rows):

    for row in rows:

        list_price = safe_float(
            row.get("Unit List Price")
        )

        quantity = safe_float(
            row.get("Qty"),
            1
        )

        discount = safe_float(
            row.get("Disc (%)")
        )

        if quantity < 0:
            quantity = 0

        if discount < 0:
            discount = 0

        if discount > 100:
            discount = 100

        net_price = round(
            list_price * (1 - discount / 100),
            2
        )

        extended_price = round(
            net_price * quantity,
            2
        )

        row["Qty"] = quantity
        row["Disc (%)"] = discount
        row["Unit Net Price"] = net_price
        row["Extended Net Price"] = extended_price

    return rows


# ============================================================
# EXPORTACIÓN SEGURA DE EXCEL
# ============================================================

def safe_set_cell(ws, row, column, value):
    """
    No escribe en celdas secundarias de rangos combinados.
    """

    cell = ws.cell(
        row=row,
        column=column
    )

    if isinstance(cell, MergedCell):
        return False

    cell.value = value

    return True


def export_quote(file_bytes, rows, header_row):

    workbook = load_workbook(
        io.BytesIO(file_bytes)
    )

    worksheet = workbook.active

    headers = [
        worksheet.cell(header_row, column).value
        for column in range(1, worksheet.max_column + 1)
    ]

    columns = {
        key: find_column(headers, aliases)
        for key, aliases in COLUMN_ALIASES.items()
    }

    row_by_excel_number = {
        row["Excel Row"]: row
        for row in rows
    }

    for excel_row, row_data in row_by_excel_number.items():

        # Solo se actualiza la información comercial editable
        # de los equipos principales.
        if row_data["Is Equipment"]:

            safe_set_cell(
                worksheet,
                excel_row,
                columns["description"] + 1,
                row_data["Description"]
            )

            safe_set_cell(
                worksheet,
                excel_row,
                columns["list_price"] + 1,
                safe_float(row_data["Unit List Price"])
            )

            safe_set_cell(
                worksheet,
                excel_row,
                columns["quantity"] + 1,
                safe_float(row_data["Qty"])
            )

            safe_set_cell(
                worksheet,
                excel_row,
                columns["discount"] + 1,
                safe_float(row_data["Disc (%)"])
            )

        # Se recalculan los importes de todas las líneas
        # sin modificar sus Part Number ni sus descripciones.
        if columns["net_price"] is not None:

            safe_set_cell(
                worksheet,
                excel_row,
                columns["net_price"] + 1,
                f"=ROUND(F{excel_row}-((F{excel_row}*J{excel_row})/100),2)"
            )

        if columns["extended_price"] is not None:

            safe_set_cell(
                worksheet,
                excel_row,
                columns["extended_price"] + 1,
                f"=ROUND(H{excel_row}*I{excel_row},2)"
            )

    # Solicitar a Excel que recalcule las fórmulas al abrir.
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except Exception:
        pass

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# ESTADO DE LA APLICACIÓN
# ============================================================

initialize_database()

if "quote_rows" not in st.session_state:
    st.session_state.quote_rows = None

if "quote_bytes" not in st.session_state:
    st.session_state.quote_bytes = None

if "quote_header_row" not in st.session_state:
    st.session_state.quote_header_row = None

if "quote_name" not in st.session_state:
    st.session_state.quote_name = None

if "uploader_version" not in st.session_state:
    st.session_state.uploader_version = 0


# ============================================================
# ENCABEZADO
# ============================================================

st.title("🌐 Cisco Price Manager")

st.caption(
    "Gestión de cotizaciones Cisco, actualización de precios "
    "y exportación a Excel."
)

st.divider()


# ============================================================
# CARGA DEL ARCHIVO
# ============================================================

st.subheader("1. Cargar cotización Cisco")

uploaded_file = st.file_uploader(
    "Selecciona la cotización en formato Excel (.xlsx)",
    type=["xlsx"],
    key=f"quote_upload_{st.session_state.uploader_version}"
)

if uploaded_file is not None:

    current_name = uploaded_file.name

    if current_name != st.session_state.quote_name:

        try:

            file_bytes = uploaded_file.getvalue()

            parsed = read_quote(file_bytes)

            st.session_state.quote_rows = parsed["rows"]
            st.session_state.quote_bytes = file_bytes
            st.session_state.quote_header_row = parsed["header_row"]
            st.session_state.quote_name = current_name

            st.success(
                f"Cotización cargada correctamente: {current_name}"
            )

        except Exception as error:

            st.error(
                f"No se pudo leer la cotización: {error}"
            )


# ============================================================
# CERRAR COTIZACIÓN ACTUAL
# ============================================================

if st.session_state.quote_rows is not None:

    if st.button(
        "Cerrar cotización actual",
        type="secondary"
    ):

        st.session_state.quote_rows = None
        st.session_state.quote_bytes = None
        st.session_state.quote_header_row = None
        st.session_state.quote_name = None

        # Cambiar la clave crea un nuevo cargador vacío.
        st.session_state.uploader_version += 1

        st.rerun()


# ============================================================
# PROCESAMIENTO DE LA COTIZACIÓN
# ============================================================

if st.session_state.quote_rows is not None:

    rows = st.session_state.quote_rows

    calculate_prices(rows)

    equipment_rows = [
        row for row in rows
        if row["Is Equipment"]
    ]

    all_skus = set(
        row["Part Number"]
        for row in rows
    )

    st.divider()

    st.subheader("2. Resumen de la cotización")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Equipos principales",
        len(equipment_rows)
    )

    col2.metric(
        "Líneas de cotización",
        len(rows)
    )

    col3.metric(
        "SKU diferentes",
        len(all_skus)
    )

    equipment_total = sum(
        safe_float(row["Extended Net Price"])
        for row in equipment_rows
    )

    col4.metric(
        "Total equipos",
        f"${equipment_total:,.2f}"
    )

    st.caption(
        f"Archivo activo: {st.session_state.quote_name}"
    )

    st.divider()

    # --------------------------------------------------------
    # EQUIPOS PRINCIPALES
    # --------------------------------------------------------

    st.subheader("3. Equipos principales identificados")

    st.info(
        "Cisco se consultará únicamente para estos equipos. "
        "Los servicios CON-, licencias, cables y accesorios "
        "se mantienen como líneas relacionadas de la cotización."
    )

    equipment_df = pd.DataFrame(equipment_rows)

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
        "Cisco Status"
    ]

    equipment_df = equipment_df[display_columns]

    edited_df = st.data_editor(
        equipment_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="equipment_editor",
        disabled=[
            "Excel Row",
            "Part Number",
            "Cisco Status"
        ],
        column_config={
            "Unit List Price": st.column_config.NumberColumn(
                "Precio de lista USD",
                min_value=0,
                format="$%.2f"
            ),
            "Qty": st.column_config.NumberColumn(
                "Cantidad",
                min_value=0,
                step=1
            ),
            "Disc (%)": st.column_config.NumberColumn(
                "Descuento (%)",
                min_value=0,
                max_value=100,
                format="%.2f"
            ),
            "Unit Net Price": st.column_config.NumberColumn(
                "Precio neto USD",
                format="$%.2f"
            ),
            "Extended Net Price": st.column_config.NumberColumn(
                "Precio extendido USD",
                format="$%.2f"
            )
        }
    )

    # Guardar los cambios realizados en la tabla.
    edited_by_row = {
        int(item["Excel Row"]): item
        for item in edited_df.to_dict("records")
    }

    for row in rows:

        if not row["Is Equipment"]:
            continue

        edited = edited_by_row.get(
            int(row["Excel Row"])
        )

        if edited:

            row["Description"] = edited["Description"]

            row["Unit List Price"] = safe_float(
                edited["Unit List Price"]
            )

            row["Qty"] = safe_float(
                edited["Qty"],
                1
            )

            row["Disc (%)"] = safe_float(
                edited["Disc (%)"]
            )

    calculate_prices(rows)

    st.divider()

    # --------------------------------------------------------
    # CONSULTA CISCO
    # --------------------------------------------------------

    st.subheader("4. Actualizar información desde Cisco")

    st.write(
        "La actualización consulta los Part Number de los "
        "equipos principales, no los 51 SKU de todas las líneas."
    )

    if st.button(
        "🔄 Actualizar los 8 equipos desde Cisco",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "Consultando información de los equipos en Cisco..."
            ):

                updated, errors = refresh_equipment_prices(rows)

                calculate_prices(rows)

            st.session_state.quote_rows = rows

            if updated:
                st.success(
                    f"Se actualizaron {updated} equipos desde Cisco."
                )

            if errors:

                st.warning(
                    "Algunas consultas no pudieron completarse:"
                )

                for error in errors:
                    st.write(f"- {error}")

            if not updated and not errors:
                st.warning(
                    "Cisco no devolvió información para los equipos."
                )

            st.rerun()

        except Exception as error:

            st.error(
                f"No se pudo realizar la consulta a Cisco: {error}"
            )

    # --------------------------------------------------------
    # TODAS LAS LÍNEAS DE LA COTIZACIÓN
    # --------------------------------------------------------

    st.divider()

    with st.expander(
        "Ver las líneas completas de la cotización"
    ):

        all_rows_df = pd.DataFrame(rows)

        columns_to_show = [
            "Excel Row",
            "Part Number",
            "Description",
            "Unit List Price",
            "Qty",
            "Disc (%)",
            "Unit Net Price",
            "Extended Net Price",
            "Is Equipment",
            "Cisco Status"
        ]

        st.dataframe(
            all_rows_df[columns_to_show],
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # EXPORTAR EXCEL
    # --------------------------------------------------------

    st.divider()

    st.subheader("5. Exportar cotización actualizada")

    st.write(
        "Se conservará la estructura del archivo original, "
        "incluyendo sus celdas combinadas y formato."
    )

    try:

        calculate_prices(rows)

        excel_output = export_quote(
            st.session_state.quote_bytes,
            rows,
            st.session_state.quote_header_row
        )

        original_name = st.session_state.quote_name

        output_name = (
            original_name.rsplit(".", 1)[0]
            + "_actualizada.xlsx"
        )

        st.download_button(
            label="📥 Descargar cotización actualizada",
            data=excel_output,
            file_name=output_name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    except Exception as error:

        st.error(
            f"No se pudo generar el Excel: {error}"
        )

    # --------------------------------------------------------
    # HISTORIAL
    # --------------------------------------------------------

    st.divider()

    with st.expander("Historial de consultas de precios"):

        history_df = load_price_history()

        if history_df.empty:
            st.info(
                "Todavía no hay consultas guardadas."
            )
        else:
            st.dataframe(
                history_df,
                use_container_width=True,
                hide_index=True
            )


else:

    st.info(
        "Carga una cotización Cisco en formato .xlsx "
        "para comenzar."
    )


# ============================================================
# PIE DE PÁGINA
# ============================================================

st.divider()

st.caption(
    "Cisco Price Manager | Los precios de Cisco solo se "
    "actualizan cuando la API autorizada está configurada "
    "y devuelve información válida."
)
