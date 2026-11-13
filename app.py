
import os
import io
import json
import sqlite3
from datetime import datetime, timezone

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook


# ============================================================
# 1. CONFIGURACION GENERAL
# ============================================================

st.set_page_config(
    page_title="Cisco Price Manager",
    page_icon="🌐",
    layout="wide"
)

APP_TITLE = "Cisco Price Manager"
DATABASE = "cisco_price_manager.db"

TOKEN_URL = "https://id.cisco.com/oauth2/default/v1/token"


# ============================================================
# 2. CONFIGURACION SEGURA
# Lee primero variables de entorno y luego Streamlit Secrets.
# Nunca escribir credenciales directamente en este archivo.
# ============================================================

def get_setting(name, default=""):
    value = os.getenv(name)

    if value:
        return value

    try:
        value = st.secrets.get(name, default)
        return value if value is not None else default
    except Exception:
        return default


CLIENT_ID = get_setting("CISCO_CLIENT_ID")
CLIENT_SECRET = get_setting("CISCO_CLIENT_SECRET")

GRAPHQL_URL = get_setting("CISCO_GRAPHQL_URL")
PRICE_QUERY = get_setting("CISCO_PRICE_QUERY")
PRICE_VARIABLES = get_setting(
    "CISCO_PRICE_VARIABLES_TEMPLATE"
)
PRICE_JSON_PATH = get_setting("CISCO_PRICE_JSON_PATH")


# ============================================================
# 3. BASE DE DATOS LOCAL
# Guarda descuentos, precios manuales e historial.
# ============================================================

def get_connection():
    return sqlite3.connect(DATABASE)


def initialize_database():

    with get_connection() as conn:

        conn.execute("""
            CREATE TABLE IF NOT EXISTS products (
                sku TEXT PRIMARY KEY,
                manual_discount REAL DEFAULT 0,
                manual_list_price REAL,
                updated_at TEXT
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS price_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sku TEXT NOT NULL,
                price REAL,
                source TEXT,
                recorded_at TEXT
            )
        """)


def get_saved_product(sku):

    with get_connection() as conn:

        row = conn.execute(
            """
            SELECT manual_discount, manual_list_price
            FROM products
            WHERE sku = ?
            """,
            (sku,)
        ).fetchone()

    if row:
        return {
            "discount": row[0],
            "manual_price": row[1]
        }

    return {
        "discount": 0.0,
        "manual_price": None
    }


def save_product(sku, discount, manual_price=None):

    now = datetime.now().isoformat(timespec="seconds")

    with get_connection() as conn:

        conn.execute(
            """
            INSERT INTO products (
                sku,
                manual_discount,
                manual_list_price,
                updated_at
            )
            VALUES (?, ?, ?, ?)
            ON CONFLICT(sku) DO UPDATE SET
                manual_discount = excluded.manual_discount,
                manual_list_price = excluded.manual_list_price,
                updated_at = excluded.updated_at
            """,
            (
                str(sku),
                float(discount),
                manual_price,
                now
            )
        )


def save_price_history(sku, price, source):

    with get_connection() as conn:

        conn.execute(
            """
            INSERT INTO price_history (
                sku,
                price,
                source,
                recorded_at
            )
            VALUES (?, ?, ?, ?)
            """,
            (
                str(sku),
                float(price),
                source,
                datetime.now().isoformat(timespec="seconds")
            )
        )


def load_price_history():

    with get_connection() as conn:

        return pd.read_sql_query(
            """
            SELECT
                sku AS SKU,
                price AS Price,
                source AS Source,
                recorded_at AS Date
            FROM price_history
            ORDER BY recorded_at DESC
            LIMIT 500
            """,
            conn
        )


initialize_database()


# ============================================================
# 4. UTILIDADES
# ============================================================

def normalize_header(value):

    if value is None:
        return ""

    return (
        str(value)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace(".", "")
    )


HEADER_ALIASES = {
    "sku": [
        "partnumber",
        "sku",
        "productnumber",
        "materialnumber"
    ],
    "description": [
        "description",
        "productdescription",
        "itemdescription"
    ],
    "list_price": [
        "unitlistprice",
        "listprice",
        "unitprice"
    ],
    "discount": [
        "disc(%)",
        "disc%",
        "discount(%)",
        "discount%",
        "discount"
    ],
    "quantity": [
        "qty",
        "quantity"
    ],
    "net_price": [
        "unitnetprice",
        "netunitprice"
    ],
    "extended": [
        "extendednetprice",
        "extendedprice",
        "totalnetprice"
    ]
}


def find_column_mapping(worksheet):

    best_row = None
    best_mapping = {}
    best_score = 0

    for row_number in range(1, min(worksheet.max_row, 60) + 1):

        values = {}

        for col in range(1, worksheet.max_column + 1):

            raw = worksheet.cell(
                row_number,
                col
            ).value

            normalized = normalize_header(raw)

            if normalized:
                values[normalized] = col

        mapping = {}

        for field, aliases in HEADER_ALIASES.items():

            for alias in aliases:

                normalized_alias = normalize_header(alias)

                if normalized_alias in values:
                    mapping[field] = values[normalized_alias]
                    break

        score = len(mapping)

        if "sku" in mapping and score > best_score:

            best_row = row_number
            best_mapping = mapping
            best_score = score

    if best_row is None:
        raise ValueError(
            "No se encontró una columna Part Number o SKU."
        )

    return best_row, best_mapping


def safe_float(value, default=0.0):

    if value is None or value == "":
        return default

    try:

        if isinstance(value, str):
            value = value.replace("$", "").replace(",", "").strip()

        return float(value)

    except (ValueError, TypeError):
        return default


def safe_optional_float(value):

    if value is None or value == "":
        return None

    try:

        if isinstance(value, str):
            value = value.replace("$", "").replace(",", "").strip()

        return float(value)

    except (ValueError, TypeError):
        return None


# ============================================================
# 5. LECTURA DE COTIZACIONES EXCEL
# ============================================================

def read_quote(file_bytes):

    workbook = load_workbook(
        io.BytesIO(file_bytes)
    )

    worksheet = workbook.active

    header_row, mapping = find_column_mapping(worksheet)

    records = []

    for row_number in range(header_row + 1, worksheet.max_row + 1):

        sku_value = worksheet.cell(
            row_number,
            mapping["sku"]
        ).value

        if sku_value is None:
            continue

        sku = str(sku_value).strip()

        if not sku:
            continue

        description = ""

        if "description" in mapping:

            description = worksheet.cell(
                row_number,
                mapping["description"]
            ).value or ""

        list_price = None

        if "list_price" in mapping:

            list_price = safe_optional_float(
                worksheet.cell(
                    row_number,
                    mapping["list_price"]
                ).value
            )

        quantity = 1

        if "quantity" in mapping:

            quantity = safe_float(
                worksheet.cell(
                    row_number,
                    mapping["quantity"]
                ).value,
                1
            )

        discount = 0

        if "discount" in mapping:

            discount = safe_float(
                worksheet.cell(
                    row_number,
                    mapping["discount"]
                ).value,
                0
            )

        net_price = None

        if "net_price" in mapping:

            net_price = safe_optional_float(
                worksheet.cell(
                    row_number,
                    mapping["net_price"]
                ).value
            )

        extended = None

        if "extended" in mapping:

            extended = safe_optional_float(
                worksheet.cell(
                    row_number,
                    mapping["extended"]
                ).value
            )

        saved = get_saved_product(sku)

        # El descuento guardado en la aplicación tiene prioridad.
        if saved["discount"] is not None:
            discount = saved["discount"]

        # Si no existe precio en el Excel, usar un precio manual
        # guardado previamente. Nunca inventar un precio.
        if list_price is None and saved["manual_price"] is not None:
            list_price = saved["manual_price"]

        records.append({
            "Excel Row": row_number,
            "Part Number": sku,
            "Description": str(description),
            "Unit List Price": list_price,
            "Qty": quantity,
            "Disc (%)": discount,
            "Unit Net Price": net_price,
            "Extended Net Price": extended,
            "Price Source": (
                "Cotización original"
                if list_price is not None
                else "Sin precio"
            ),
            "Cisco Status": "Pendiente",
            "Last Updated": ""
        })

    if not records:
        raise ValueError(
            "No se encontraron productos en la cotización."
        )

    return workbook, worksheet.title, header_row, mapping, pd.DataFrame(records)


# ============================================================
# 6. AUTENTICACION OAUTH CISCO
# ============================================================

def get_cisco_token():

    if not CLIENT_ID or not CLIENT_SECRET:

        raise RuntimeError(
            "Faltan CISCO_CLIENT_ID y CISCO_CLIENT_SECRET."
        )

    response = requests.post(
        TOKEN_URL,
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json"
        },
        data={
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET
        },
        timeout=30
    )

    if response.status_code != 200:

        raise RuntimeError(
            f"Autenticación Cisco fallida: HTTP "
            f"{response.status_code}. "
            "Verifica la aprobación y configuración de la aplicación."
        )

    payload = response.json()

    token = payload.get("access_token")

    if not token:
        raise RuntimeError(
            "Cisco no devolvió un token de acceso."
        )

    return token


# ============================================================
# 7. ADAPTADOR GRAPHQL CISCO
#
# Esta función queda preparada, pero requiere la consulta
# REAL publicada para la aplicación aprobada.
# No se inventan nombres de campos ni operaciones.
# ============================================================

def get_nested_value(data, path):

    current = data

    for part in path.split("."):

        if isinstance(current, list):
            current = current[int(part)]

        elif isinstance(current, dict):
            current = current[part]

        else:
            raise KeyError(
                f"No existe la ruta de respuesta: {path}"
            )

    return current


def query_cisco_price(sku, token):

    required = {
        "CISCO_GRAPHQL_URL": GRAPHQL_URL,
        "CISCO_PRICE_QUERY": PRICE_QUERY,
        "CISCO_PRICE_VARIABLES_TEMPLATE": PRICE_VARIABLES,
        "CISCO_PRICE_JSON_PATH": PRICE_JSON_PATH
    }

    missing = [
        name for name, value in required.items()
        if not value
    ]

    if missing:

        raise RuntimeError(
            "La conexión Cisco todavía no está configurada. "
            "Faltan: " + ", ".join(missing)
        )

    # El marcador {{SKU}} se sustituye dentro del JSON de variables.
    variables_text = PRICE_VARIABLES.replace(
        "{{SKU}}",
        json.dumps(str(sku))
    )

    variables = json.loads(variables_text)

    response = requests.post(
        GRAPHQL_URL,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        },
        json={
            "query": PRICE_QUERY,
            "variables": variables
        },
        timeout=45
    )

    if response.status_code != 200:

        raise RuntimeError(
            f"Consulta Cisco fallida: HTTP {response.status_code}"
        )

    payload = response.json()

    if payload.get("errors"):

        raise RuntimeError(
            "Error GraphQL: "
            + json.dumps(
                payload["errors"],
                ensure_ascii=False
            )[:1500]
        )

    raw_price = get_nested_value(
        payload,
        PRICE_JSON_PATH
    )

    price = float(raw_price)

    if price < 0:
        raise ValueError("Cisco devolvió un precio negativo.")

    return price


# ============================================================
# 8. ACTUALIZACION DE PRECIOS
# ============================================================

def refresh_cisco_prices(dataframe):

    if not all([
        CLIENT_ID,
        CLIENT_SECRET,
        GRAPHQL_URL,
        PRICE_QUERY,
        PRICE_VARIABLES,
        PRICE_JSON_PATH
    ]):

        raise RuntimeError(
            "La conexión real con Cisco aún no está configurada. "
            "Puedes utilizar el modo local mientras esperas "
            "la aprobación y la documentación de Cisco."
        )

    token = get_cisco_token()

    df = dataframe.copy()

    unique_skus = (
        df["Part Number"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    results = {}

    progress = st.progress(0)
    status = st.empty()

    for index, sku in enumerate(unique_skus):

        status.write(
            f"Consultando Cisco: {sku}"
        )

        try:

            price = query_cisco_price(
                sku,
                token
            )

            results[sku] = {
                "price": price,
                "status": "Actualizado por Cisco",
                "error": ""
            }

            save_price_history(
                sku,
                price,
                "Cisco API"
            )

        except Exception as error:

            results[sku] = {
                "price": None,
                "status": "No actualizado",
                "error": str(error)
            }

        progress.progress(
            (index + 1) / len(unique_skus)
        )

    for index, row in df.iterrows():

        sku = str(row["Part Number"])
        result = results[sku]

        if result["price"] is not None:

            df.at[index, "Unit List Price"] = result["price"]
            df.at[index, "Price Source"] = "Cisco API"
            df.at[index, "Cisco Status"] = result["status"]

        else:

            df.at[index, "Cisco Status"] = (
                result["status"] + ": " + result["error"]
            )

    df = calculate_prices(df)

    df["Last Updated"] = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    status.success("Consulta finalizada.")

    return df


# ============================================================
# 9. CALCULO DE PRECIOS
# ============================================================

def calculate_prices(dataframe):

    df = dataframe.copy()

    for index, row in df.iterrows():

        list_price = safe_optional_float(
            row["Unit List Price"]
        )

        discount = safe_float(
            row["Disc (%)"],
            0
        )

        quantity = safe_float(
            row["Qty"],
            1
        )

        if discount < 0 or discount > 100:
            raise ValueError(
                f"Descuento inválido para {row['Part Number']}. "
                "Debe estar entre 0 y 100."
            )

        if quantity < 0:
            raise ValueError(
                f"Cantidad inválida para {row['Part Number']}."
            )

        if list_price is None:

            df.at[index, "Unit Net Price"] = None
            df.at[index, "Extended Net Price"] = None

        else:

            net_price = list_price * (
                1 - discount / 100
            )

            df.at[index, "Unit Net Price"] = net_price

            df.at[index, "Extended Net Price"] = (
                net_price * quantity
            )

    return df


# ============================================================
# 10. GUARDAR CONFIGURACION MANUAL
# ============================================================

def persist_manual_values(df):

    for _, row in df.iterrows():

        sku = str(row["Part Number"])

        discount = safe_float(
            row["Disc (%)"],
            0
        )

        source = str(row.get("Price Source", ""))

        manual_price = None

        if source in [
            "Precio manual",
            "Precio editado manualmente"
        ]:

            manual_price = safe_optional_float(
                row["Unit List Price"]
            )

        save_product(
            sku,
            discount,
            manual_price
        )


# ============================================================
# 11. EXPORTAR EXCEL CONSERVANDO LA COTIZACION ORIGINAL
# ============================================================

def write_cell(worksheet, row, column, value):

    if column is None:
        return

    cell = worksheet.cell(row, column)

    if value is None or pd.isna(value):
        cell.value = None
    else:
        cell.value = float(value)


def export_quote(
    workbook,
    worksheet_name,
    header_row,
    mapping,
    dataframe
):

    worksheet = workbook[worksheet_name]

    # Se actualizan únicamente las columnas que existen.
    # Las demás hojas y el resto del formato se conservan.

    for _, record in dataframe.iterrows():

        excel_row = int(record["Excel Row"])

        write_cell(
            worksheet,
            excel_row,
            mapping.get("list_price"),
            record["Unit List Price"]
        )

        write_cell(
            worksheet,
            excel_row,
            mapping.get("discount"),
            record["Disc (%)"]
        )

        write_cell(
            worksheet,
            excel_row,
            mapping.get("net_price"),
            record["Unit Net Price"]
        )

        write_cell(
            worksheet,
            excel_row,
            mapping.get("extended"),
            record["Extended Net Price"]
        )

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# 12. INTERFAZ PRINCIPAL
# ============================================================

st.title("🌐 Cisco Price Manager")

st.caption(
    "Gestión de cotizaciones, precios de lista y descuentos Cisco"
)

st.markdown(
    """
    Carga una cotización Cisco, revisa los productos,
    administra descuentos y actualiza precios cuando
    la conexión autorizada esté disponible.
    """
)

# ---------------- SIDEBAR ----------------

with st.sidebar:

    st.header("Configuración")

    st.subheader("Estado de Cisco API")

    credentials_ready = bool(
        CLIENT_ID and CLIENT_SECRET
    )

    query_ready = bool(
        GRAPHQL_URL
        and PRICE_QUERY
        and PRICE_VARIABLES
        and PRICE_JSON_PATH
    )

    if credentials_ready:
        st.success("Credenciales OAuth configuradas")
    else:
        st.warning("Credenciales OAuth pendientes")

    if query_ready:
        st.success("Consulta de precios configurada")
    else:
        st.warning("Consulta GraphQL pendiente")

    if credentials_ready and query_ready:
        st.info("Conexión preparada para probar")
    else:
        st.info("La aplicación funciona en modo local")

    st.divider()

    st.caption(
        "Los precios manuales no se presentan como "
        "precios verificados por Cisco."
    )


# ---------------- CARGA DE COTIZACION ----------------

uploaded_file = st.file_uploader(
    "Cargar cotización Cisco (.xlsx)",
    type=["xlsx"],
    help="Carga una cotización descargada de Cisco CCW."
)

if uploaded_file:

    file_bytes = uploaded_file.getvalue()

    quote_key = (
        uploaded_file.name,
        len(file_bytes)
    )

    if st.session_state.get("quote_key") != quote_key:

        try:

            (
                workbook,
                worksheet_name,
                header_row,
                mapping,
                dataframe
            ) = read_quote(file_bytes)

            st.session_state.quote_key = quote_key
            st.session_state.workbook_bytes = file_bytes
            st.session_state.worksheet_name = worksheet_name
            st.session_state.header_row = header_row
            st.session_state.mapping = mapping
            st.session_state.quote_df = dataframe

        except Exception as error:

            st.error(
                f"No se pudo leer la cotización: {error}"
            )

            st.stop()

    df = st.session_state.quote_df.copy()

    # ---------------- INDICADORES ----------------

    total_lines = len(df)
    unique_skus = df["Part Number"].nunique()

    missing_prices = df["Unit List Price"].apply(
        lambda value: safe_optional_float(value) is None
    ).sum()

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Líneas", total_lines)
    col2.metric("SKU únicos", unique_skus)
    col3.metric("Sin precio", int(missing_prices))
    col4.metric(
        "Moneda",
        "USD"
    )

    st.divider()

    # ---------------- TABLA EDITABLE ----------------

    st.subheader("Productos de la cotización")

    st.caption(
        "Puedes modificar cantidades, descuentos y precios "
        "manualmente. Los precios de Cisco se identificarán "
        "por separado."
    )

    editable_columns = [
        "Excel Row",
        "Part Number",
        "Description",
        "Unit List Price",
        "Qty",
        "Disc (%)",
        "Unit Net Price",
        "Extended Net Price",
        "Price Source",
        "Cisco Status",
        "Last Updated"
    ]

    edited = st.data_editor(
        df[editable_columns],
        use_container_width=True,
        hide_index=True,
        disabled=[
            "Excel Row",
            "Part Number",
            "Description",
            "Unit Net Price",
            "Extended Net Price",
            "Price Source",
            "Cisco Status",
            "Last Updated"
        ],
        column_config={
            "Unit List Price": st.column_config.NumberColumn(
                "Unit List Price",
                min_value=0,
                format="%.2f"
            ),
            "Qty": st.column_config.NumberColumn(
                "Qty",
                min_value=0,
                step=1
            ),
            "Disc (%)": st.column_config.NumberColumn(
                "Disc (%)",
                min_value=0,
                max_value=100,
                step=0.1,
                format="%.2f%%"
            )
        },
        key="products_editor"
    )

    # Detectar cambios manuales en precio.
    for index in df.index:

        old_price = safe_optional_float(
            df.at[index, "Unit List Price"]
        )

        new_price = safe_optional_float(
            edited.at[index, "Unit List Price"]
        )

        if new_price != old_price:

            df.at[index, "Price Source"] = (
                "Precio editado manualmente"
            )

            df.at[index, "Cisco Status"] = (
                "Precio manual; pendiente de validación Cisco"
            )

    df["Qty"] = edited["Qty"].values
    df["Disc (%)"] = edited["Disc (%)"].values
    df["Unit List Price"] = edited["Unit List Price"].values

    try:

        df = calculate_prices(df)

        persist_manual_values(df)

        st.session_state.quote_df = df

    except Exception as error:

        st.error(str(error))

    st.divider()

    # ---------------- ACCIONES ----------------

    st.subheader("Acciones")

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button(
            "🧮 Recalcular",
            use_container_width=True
        ):

            try:

                df = calculate_prices(
                    st.session_state.quote_df
                )

                st.session_state.quote_df = df

                st.success("Precios recalculados.")

            except Exception as error:

                st.error(str(error))

    with col2:

        if st.button(
            "🔄 Actualizar desde Cisco",
            type="primary",
            use_container_width=True
        ):

            try:

                updated = refresh_cisco_prices(
                    st.session_state.quote_df
                )

                st.session_state.quote_df = updated

                persist_manual_values(updated)

                st.success(
                    "Proceso de actualización finalizado."
                )

            except Exception as error:

                st.error(str(error))

    with col3:

        if st.button(
            "🧹 Limpiar cotización",
            use_container_width=True
        ):

            for key in [
                "quote_key",
                "workbook_bytes",
                "worksheet_name",
                "header_row",
                "mapping",
                "quote_df"
            ]:

                st.session_state.pop(key, None)

            st.rerun()

    # ---------------- RESULTADOS ----------------

    st.divider()

    st.subheader("Resumen de precios")

    result = st.session_state.quote_df.copy()

    total = pd.to_numeric(
        result["Extended Net Price"],
        errors="coerce"
    ).sum()

    valid_prices = result["Unit List Price"].apply(
        lambda value: safe_optional_float(value) is not None
    ).sum()

    updated_prices = (
        result["Cisco Status"]
        .astype(str)
        .str.contains("Actualizado por Cisco", case=False)
        .sum()
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Total calculado",
        f"USD {total:,.2f}"
    )

    col2.metric(
        "Líneas con precio",
        f"{int(valid_prices)} / {len(result)}"
    )

    col3.metric(
        "Líneas actualizadas por Cisco",
        int(updated_prices)
    )

    st.dataframe(
        result[
            [
                "Part Number",
                "Description",
                "Unit List Price",
                "Qty",
                "Disc (%)",
                "Unit Net Price",
                "Extended Net Price",
                "Price Source",
                "Cisco Status",
                "Last Updated"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )

    # ---------------- EXPORTAR ----------------

    st.divider()

    st.subheader("Exportar cotización")

    try:

        original_workbook = load_workbook(
            io.BytesIO(
                st.session_state.workbook_bytes
            )
        )

        excel_bytes = export_quote(
            original_workbook,
            st.session_state.worksheet_name,
            st.session_state.header_row,
            st.session_state.mapping,
            result
        )

        st.download_button(
            "📥 Descargar cotización actualizada",
            data=excel_bytes,
            file_name="Cisco_Price_Manager_Updated.xlsx",
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


else:

    st.info(
        "Carga una cotización Cisco en formato XLSX "
        "para comenzar."
    )

    st.markdown(
        """
        **Funciones disponibles en modo local**

        - Importar cotizaciones.
        - Identificar SKU.
        - Editar precios y descuentos.
        - Calcular precios netos.
        - Guardar descuentos y precios manuales localmente.
        - Consultar historial.
        - Exportar la cotización.
        """
    )


# ============================================================
# 13. HISTORIAL DE PRECIOS
# ============================================================

st.divider()

with st.expander("📊 Historial de precios"):

    history = load_price_history()

    if history.empty:

        st.info(
            "Todavía no hay consultas de precios registradas."
        )

    else:

        st.dataframe(
            history,
            use_container_width=True,
            hide_index=True
        )

        history_csv = history.to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            "Descargar historial CSV",
            data=history_csv,
            file_name="Cisco_Price_History.csv",
            mime="text/csv"
        )


st.caption(
    "Cisco Price Manager | Los precios solo se consideran "
    "verificados por Cisco cuando la API devuelve una respuesta válida."
)

