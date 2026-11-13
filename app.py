
import os
import io
import json
import time
import requests
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from datetime import datetime

# ============================================================
# CONFIGURACION
# ============================================================

st.set_page_config(
    page_title="Cisco Price Manager",
    page_icon="🌐",
    layout="wide"
)

TOKEN_URL = "https://id.cisco.com/oauth2/default/v1/token"

# No se inventa el endpoint GraphQL.
# Se debe configurar con la URL exacta indicada por Cisco
# para la aplicación y entorno que hayan sido aprobados.

CLIENT_ID = os.getenv("CISCO_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("CISCO_CLIENT_SECRET", "")
GRAPHQL_URL = os.getenv("CISCO_GRAPHQL_URL", "")

PRICE_QUERY = os.getenv("CISCO_PRICE_QUERY", "")
VARIABLES_TEMPLATE = os.getenv(
    "CISCO_PRICE_VARIABLES_TEMPLATE", ""
)
PRICE_JSON_PATH = os.getenv("CISCO_PRICE_JSON_PATH", "")

# ============================================================
# ESTILOS
# ============================================================

st.title("🌐 Cisco Price Manager")
st.caption(
    "Administración y actualización de precios de productos Cisco"
)

# ============================================================
# AUTENTICACION CISCO
# ============================================================

def get_cisco_token():

    if not CLIENT_ID or not CLIENT_SECRET:
        raise ValueError(
            "Faltan CISCO_CLIENT_ID o CISCO_CLIENT_SECRET."
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
            f"Error de autenticación Cisco: HTTP "
            f"{response.status_code}. "
            "Revisa el registro y los permisos de la aplicación."
        )

    data = response.json()

    token = data.get("access_token")

    if not token:
        raise RuntimeError(
            "Cisco no devolvió un access_token."
        )

    return token


# ============================================================
# FUNCIONES GRAPHQL
# ============================================================

def get_json_path(data, path):

    """
    Lee un campo anidado de una respuesta JSON.

    Ejemplo:
    data.catalog.product.listPrice
    """

    current = data

    for part in path.split("."):

        if isinstance(current, list):
            current = current[int(part)]

        elif isinstance(current, dict):
            current = current[part]

        else:
            raise KeyError(
                f"No se encontró la ruta: {path}"
            )

    return current


def query_cisco_price(sku, token):

    """
    Consulta un SKU mediante la operación GraphQL
    configurada con la documentación oficial de Cisco.

    La consulta, variables y ruta de respuesta deben
    corresponder exactamente al esquema autorizado.
    """

    if not GRAPHQL_URL:
        raise ValueError(
            "No está configurado CISCO_GRAPHQL_URL."
        )

    if not PRICE_QUERY:
        raise ValueError(
            "No está configurado CISCO_PRICE_QUERY."
        )

    if not VARIABLES_TEMPLATE:
        raise ValueError(
            "No está configurado "
            "CISCO_PRICE_VARIABLES_TEMPLATE."
        )

    if not PRICE_JSON_PATH:
        raise ValueError(
            "No está configurado CISCO_PRICE_JSON_PATH."
        )

    variables_text = VARIABLES_TEMPLATE.replace(
        "{{SKU}}", sku
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
            f"Error Cisco HTTP {response.status_code}"
        )

    result = response.json()

    if result.get("errors"):
        raise RuntimeError(
            "Cisco devolvió errores GraphQL: "
            + json.dumps(result["errors"], ensure_ascii=False)
        )

    price = get_json_path(
        result,
        PRICE_JSON_PATH
    )

    if price is None:
        raise ValueError(
            "Cisco no devolvió un precio."
        )

    price = float(price)

    if price < 0:
        raise ValueError(
            "El precio devuelto no es válido."
        )

    return price


# ============================================================
# LECTURA DE COTIZACION
# ============================================================

def find_header_row(workbook):

    ws = workbook.active

    for row_number in range(1, 40):

        values = [
            str(cell.value).strip()
            if cell.value is not None else ""
            for cell in ws[row_number]
        ]

        if (
            "Part Number" in values
            and "Unit List Price" in values
        ):
            return row_number

    raise ValueError(
        "No se encontró la fila de encabezados "
        "de la cotización Cisco."
    )


def load_quote(file_bytes):

    workbook = load_workbook(
        io.BytesIO(file_bytes)
    )

    ws = workbook.active

    header_row = find_header_row(workbook)

    headers = {}

    for column in range(1, ws.max_column + 1):

        value = ws.cell(
            header_row,
            column
        ).value

        if value is not None:
            headers[str(value).strip()] = column

    required = [
        "Part Number",
        "Unit List Price",
        "Unit Net Price",
        "Disc(%)",
        "Qty",
        "Extended Net Price"
    ]

    missing = [
        field for field in required
        if field not in headers
    ]

    if missing:
        raise ValueError(
            "Faltan columnas requeridas: "
            + ", ".join(missing)
        )

    records = []

    for row in range(header_row + 1, ws.max_row + 1):

        sku = ws.cell(
            row,
            headers["Part Number"]
        ).value

        if not sku:
            continue

        record = {
            "Excel Row": row,
            "Part Number": str(sku).strip()
        }

        for field in headers:

            if field == "Part Number":
                continue

            record[field] = ws.cell(
                row,
                headers[field]
            ).value

        records.append(record)

    df = pd.DataFrame(records)

    if df.empty:
        raise ValueError(
            "La cotización no contiene SKU."
        )

    return workbook, df, headers, header_row


# ============================================================
# ACTUALIZACION DE PRECIOS
# ============================================================

def refresh_prices(df):

    df = df.copy()

    token = get_cisco_token()

    unique_skus = (
        df["Part Number"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    price_results = {}

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

            price_results[sku] = {
                "price": price,
                "status": "Actualizado",
                "error": ""
            }

        except Exception as error:

            price_results[sku] = {
                "price": None,
                "status": "Sin actualizar",
                "error": str(error)
            }

        progress.progress(
            (index + 1) / len(unique_skus)
        )

        time.sleep(0.1)

    status.success(
        "Finalizó la consulta de SKU."
    )

    df["Cisco List Price"] = df["Part Number"].map(
        lambda sku: price_results.get(
            str(sku),
            {}
        ).get("price")
    )

    df["Cisco Status"] = df["Part Number"].map(
        lambda sku: price_results.get(
            str(sku),
            {}
        ).get("status", "Sin consultar")
    )

    df["Cisco Error"] = df["Part Number"].map(
        lambda sku: price_results.get(
            str(sku),
            {}
        ).get("error", "")
    )

    # Solo reemplazar precios cuando Cisco devuelve
    # un valor válido. Los descuentos se conservan.

    for index, row in df.iterrows():

        new_price = row["Cisco List Price"]

        if pd.notna(new_price):

            discount = pd.to_numeric(
                row.get("Disc(%)", 0),
                errors="coerce"
            )

            quantity = pd.to_numeric(
                row.get("Qty", 1),
                errors="coerce"
            )

            discount = (
                0 if pd.isna(discount) else discount
            )

            quantity = (
                1 if pd.isna(quantity) else quantity
            )

            net_price = new_price * (
                1 - discount / 100
            )

            df.at[index, "Unit List Price"] = new_price
            df.at[index, "Unit Net Price"] = net_price
            df.at[index, "Extended Net Price"] = (
                net_price * quantity
            )

    df["Last Updated"] = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    return df


# ============================================================
# EXPORTAR EXCEL
# ============================================================

def export_quote(
    workbook,
    df,
    headers,
    header_row
):

    ws = workbook.active

    for _, record in df.iterrows():

        excel_row = int(record["Excel Row"])

        for field in [
            "Unit List Price",
            "Unit Net Price",
            "Extended Net Price"
        ]:

            if field in headers:

                value = record.get(field)

                if pd.notna(value):

                    ws.cell(
                        excel_row,
                        headers[field]
                    ).value = float(value)

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# INTERFAZ
# ============================================================

with st.sidebar:

    st.header("Conexión Cisco")

    if CLIENT_ID and CLIENT_SECRET:

        st.success(
            "Credenciales configuradas"
        )

    else:

        st.warning(
            "Faltan credenciales OAuth"
        )

    if GRAPHQL_URL and PRICE_QUERY:

        st.success(
            "Consulta GraphQL configurada"
        )

    else:

        st.info(
            "Consulta de precios pendiente "
            "de configurar con el esquema oficial."
        )

    st.caption(
        "Las credenciales se leen desde variables "
        "de entorno. Nunca se muestran en pantalla."
    )


uploaded_file = st.file_uploader(
    "Cargar cotización Cisco",
    type=["xlsx"]
)

if uploaded_file:

    try:

        workbook, original_df, headers, header_row = (
            load_quote(uploaded_file.getvalue())
        )

        if "quote_df" not in st.session_state:
            st.session_state.quote_df = original_df.copy()

        df = st.session_state.quote_df

        st.success(
            "Cotización cargada correctamente."
        )

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Líneas",
            len(df)
        )

        col2.metric(
            "SKU diferentes",
            df["Part Number"].nunique()
        )

        col3.metric(
            "Moneda",
            "USD"
        )

        st.subheader("Productos de la cotización")

        editable_columns = [
            column for column in [
                "Part Number",
                "Description",
                "Unit List Price",
                "Qty",
                "Disc(%)",
                "Unit Net Price",
                "Extended Net Price"
            ]
            if column in df.columns
        ]

        edited_df = st.data_editor(
            df[editable_columns],
            use_container_width=True,
            disabled=[
                column for column in editable_columns
                if column not in [
                    "Qty",
                    "Disc(%)"
                ]
            ],
            key="quote_editor"
        )

        # Incorporar cantidades y descuentos editados.
        for column in ["Qty", "Disc(%)"]:

            if column in edited_df.columns:
                df[column] = edited_df[column]

        st.session_state.quote_df = df

        st.divider()

        col1, col2 = st.columns(2)

        with col1:

            if st.button(
                "🔄 Actualizar precios desde Cisco",
                type="primary",
                use_container_width=True
            ):

                try:

                    updated_df = refresh_prices(
                        st.session_state.quote_df
                    )

                    st.session_state.quote_df = updated_df

                    st.success(
                        "Proceso de actualización terminado."
                    )

                except Exception as error:

                    st.error(str(error))

        with col2:

            if st.button(
                "🧮 Recalcular precios",
                use_container_width=True
            ):

                df = st.session_state.quote_df.copy()

                for index, row in df.iterrows():

                    price = pd.to_numeric(
                        row.get("Unit List Price"),
                        errors="coerce"
                    )

                    discount = pd.to_numeric(
                        row.get("Disc(%)", 0),
                        errors="coerce"
                    )

                    quantity = pd.to_numeric(
                        row.get("Qty", 1),
                        errors="coerce"
                    )

                    if pd.notna(price):

                        discount = (
                            0 if pd.isna(discount)
                            else discount
                        )

                        quantity = (
                            1 if pd.isna(quantity)
                            else quantity
                        )

                        net = price * (
                            1 - discount / 100
                        )

                        df.at[index, "Unit Net Price"] = net

                        df.at[index, "Extended Net Price"] = (
                            net * quantity
                        )

                st.session_state.quote_df = df

                st.success(
                    "Precios recalculados."
                )

        st.subheader("Resultado")

        result_df = st.session_state.quote_df

        display_columns = [
            column for column in [
                "Part Number",
                "Description",
                "Unit List Price",
                "Qty",
                "Disc(%)",
                "Unit Net Price",
                "Extended Net Price",
                "Cisco Status",
                "Last Updated"
            ]
            if column in result_df.columns
        ]

        st.dataframe(
            result_df[display_columns],
            use_container_width=True,
            hide_index=True
        )

        total = pd.to_numeric(
            result_df.get(
                "Extended Net Price",
                pd.Series(dtype=float)
            ),
            errors="coerce"
        ).sum()

        st.metric(
            "Total de productos",
            f"USD {total:,.2f}"
        )

        export_bytes = export_quote(
            workbook,
            result_df,
            headers,
            header_row
        )

        st.download_button(
            "📥 Descargar cotización actualizada",
            data=export_bytes,
            file_name="Cisco_Quote_Updated.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    except Exception as error:

        st.error(
            f"No fue posible procesar la cotización: {error}"
        )

else:

    st.info(
        "Carga una cotización Cisco en formato XLSX "
        "para comenzar."
    )
