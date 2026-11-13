import io
from datetime import datetime
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(page_title="Cisco Price Manager", page_icon="📊", layout="wide")

HEADERS = [
    "Part Number", "Smart Account Mandatory", "Description",
    "Service Duration (Months)", "Estimated Lead Time (Days)",
    "Unit List Price", "Pricing Term", "Qty", "Unit Net Price",
    "Disc(%)", "Extended Net Price"
]

def as_number(value, default=0.0):
    try:
        if value is None or value == "":
            return default
        return float(value)
    except (TypeError, ValueError):
        return default

def parse_cisco_quote(file):
    wb = load_workbook(file, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row = None
    for r in range(1, min(ws.max_row, 60) + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if "Part Number" in vals and "Unit List Price" in vals:
            header_row = r
            break
    if header_row is None:
        raise ValueError("No encontré la fila de encabezados de una cotización Cisco.")

    colmap = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v in HEADERS:
            colmap[v] = c

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        sku = ws.cell(r, colmap["Part Number"]).value if "Part Number" in colmap else None
        if not sku:
            continue
        row = {}
        for h in HEADERS:
            c = colmap.get(h)
            row[h] = ws.cell(r, c).value if c else None
        row["Part Number"] = str(row["Part Number"]).strip()
        row["Unit List Price"] = as_number(row["Unit List Price"])
        row["Qty"] = as_number(row["Qty"], 1)
        row["Disc(%)"] = as_number(row["Disc(%)"])
        row["Unit Net Price"] = round(row["Unit List Price"] * (1 - row["Disc(%)"] / 100), 2)
        row["Extended Net Price"] = round(row["Unit Net Price"] * row["Qty"], 2)
        rows.append(row)

    meta = {
        "Customer": ws["A3"].value,
        "Quote Date": ws["A10"].value,
        "Estimate ID": ws["K10"].value,
        "Deal ID": ws["K11"].value,
        "Price List": ws["K12"].value,
    }
    return pd.DataFrame(rows, columns=HEADERS), meta

def normalize_df(df):
    for col in HEADERS:
        if col not in df.columns:
            df[col] = None
    df = df[HEADERS].copy()
    df["Part Number"] = df["Part Number"].fillna("").astype(str).str.strip()
    for col in ["Unit List Price", "Qty", "Disc(%)"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Qty"] = df["Qty"].replace(0, 1)
    df["Unit Net Price"] = (df["Unit List Price"] * (1 - df["Disc(%)"] / 100)).round(2)
    df["Extended Net Price"] = (df["Unit Net Price"] * df["Qty"]).round(2)
    return df[df["Part Number"] != ""].reset_index(drop=True)

def classify(row):
    sku = str(row["Part Number"]).upper()
    desc = str(row["Description"] or "").lower()
    if sku.startswith("CON-") or "support" in desc or "maintenance" in desc:
        return "Service"
    if "subscription" in desc:
        return "Subscription"
    return "Product"

def export_quote(df, meta, currency):
    df = normalize_df(df)
    out = io.BytesIO()
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Price Estimate"

    ws["A1"] = "Price Estimate"
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="17365D")
    ws.merge_cells("A1:K1")

    ws["A3"] = meta.get("Customer") or "Customer"
    ws["A3"].font = Font(bold=True, size=12)
    ws["A9"] = "Price Estimate for planning and information purposes only; not a binding offer from Cisco."
    ws.merge_cells("A9:K9")
    ws["A10"] = "Date:"
    ws["B10"] = datetime.now().strftime("%d-%b-%Y")
    ws["I10"] = "Estimate ID:"
    ws["K10"] = meta.get("Estimate ID") or "Internal"
    ws["I11"] = "Deal ID:"
    ws["K11"] = meta.get("Deal ID") or "NA"
    ws["I12"] = "Price List:"
    ws["K12"] = meta.get("Price List") or "User supplied / pending Cisco validation"
    ws["A13"] = f"All prices are shown in {currency}"

    header_row = 14
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(header_row, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for i, (_, row) in enumerate(df.iterrows(), header_row + 1):
        values = [row[h] for h in HEADERS]
        for c, value in enumerate(values, 1):
            ws.cell(i, c, value)
        for c in [6, 9, 11]:
            ws.cell(i, c).number_format = f'"{currency}" #,##0.00'
        ws.cell(i, 10).number_format = '0.00"%"'
        ws.cell(i, 8).number_format = '0.##'

    end_row = header_row + len(df)
    summary_row = end_row + 2
    categories = {"Product": 0.0, "Service": 0.0, "Subscription": 0.0}
    for _, row in df.iterrows():
        categories[classify(row)] += float(row["Extended Net Price"])
    for offset, (label, value) in enumerate([
        ("Product Total", categories["Product"]),
        ("Service Total", categories["Service"]),
        ("Subscription Total", categories["Subscription"]),
        ("Total Price", sum(categories.values()))
    ]):
        rr = summary_row + offset
        ws.cell(rr, 9, label).font = Font(bold=True)
        ws.cell(rr, 11, round(value, 2)).font = Font(bold=True)
        ws.cell(rr, 11).number_format = f'"{currency}" #,##0.00'
        if label == "Total Price":
            ws.cell(rr, 9).fill = PatternFill("solid", fgColor="D9EAF7")
            ws.cell(rr, 11).fill = PatternFill("solid", fgColor="D9EAF7")

    widths = [25, 24, 65, 20, 23, 18, 18, 10, 18, 12, 22]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A15"
    ws.auto_filter.ref = f"A{header_row}:K{end_row}"
    ws.row_dimensions[header_row].height = 34
    thin = Side(style="thin", color="D9E2F3")
    for row in ws.iter_rows(min_row=header_row, max_row=end_row, min_col=1, max_col=11):
        for cell in row:
            cell.border = Border(bottom=thin)
            if cell.row > header_row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(out)
    out.seek(0)
    return out

st.title("Cisco Price Manager")
st.caption("Gestión de cotizaciones, descuentos manuales y comparación de precios Cisco")

with st.sidebar:
    st.header("Configuración")
    currency = st.selectbox("Moneda", ["USD", "COP", "EUR"], index=0)
    st.info("Los descuentos se ingresan manualmente por SKU. Los precios no se consultan todavía en Cisco.")
    st.markdown("**Estado de integración**")
    st.warning("Cisco API: pendiente de autorización y configuración")

tab_quote, tab_skus, tab_compare = st.tabs(["Cotización desde Excel", "Consulta de SKU", "Comparar precios"])

if "data" not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=HEADERS)
if "meta" not in st.session_state:
    st.session_state.meta = {}

with tab_quote:
    st.subheader("Importar una cotización Cisco")
    uploaded = st.file_uploader("Selecciona el archivo Excel de Cisco (.xlsx)", type=["xlsx"], key="quote_upload")
    if uploaded is not None:
        try:
            data, meta = parse_cisco_quote(uploaded)
            st.session_state.data = data
            st.session_state.meta = meta
            st.success(f"Se importaron {len(data)} líneas de producto.")
        except Exception as e:
            st.error(f"No pude leer la cotización: {e}")

    if not st.session_state.data.empty:
        meta = st.session_state.meta
        c1, c2, c3 = st.columns(3)
        c1.metric("Cliente", str(meta.get("Customer") or "No indicado"))
        c2.metric("Estimate ID", str(meta.get("Estimate ID") or "No indicado"))
        c3.metric("SKU / líneas", len(st.session_state.data))

        st.markdown("#### Edita cantidades, precios de lista y descuentos")
        st.caption("Los precios netos y extendidos se recalculan automáticamente al exportar.")
        edited = st.data_editor(
            st.session_state.data,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "Unit List Price": st.column_config.NumberColumn(format="%.2f", min_value=0.0),
                "Qty": st.column_config.NumberColumn(format="%.0f", min_value=1),
                "Disc(%)": st.column_config.NumberColumn(format="%.2f", min_value=0.0, max_value=100.0),
                "Unit Net Price": st.column_config.NumberColumn(disabled=True, format="%.2f"),
                "Extended Net Price": st.column_config.NumberColumn(disabled=True, format="%.2f"),
            },
            key="quote_editor"
        )
        edited = normalize_df(edited)
        st.session_state.data = edited
        total = edited["Extended Net Price"].sum()
        a, b, c = st.columns(3)
        a.metric("Product / líneas", len(edited))
        b.metric("Total estimado", f"{currency} {total:,.2f}")
        c.metric("Descuento promedio", f"{edited['Disc(%)'].mean() if len(edited) else 0:.2f}%")

        xlsx = export_quote(edited, meta, currency)
        st.download_button("Descargar cotización actualizada", data=xlsx,
                           file_name="Cisco_Price_Manager_Quote.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Carga una cotización Cisco para conservar sus líneas y editar descuentos.")

with tab_skus:
    st.subheader("Agregar SKU manualmente")
    st.write("Puedes pegar uno o varios SKU, uno por línea. Para cada referencia podrás completar descripción, precio de lista, cantidad y descuento.")
    sku_text = st.text_area("SKU Cisco", placeholder="C1121X-8P\nCON-L1NCD-C1121X8P\nPWR-66W-AC-V2", height=120)
    if st.button("Agregar SKU a la cotización", type="primary"):
        new_rows = []
        existing = set(st.session_state.data["Part Number"].astype(str)) if not st.session_state.data.empty else set()
        for sku in [x.strip() for x in sku_text.splitlines() if x.strip()]:
            if sku not in existing:
                new_rows.append({
                    "Part Number": sku, "Smart Account Mandatory": "",
                    "Description": "", "Service Duration (Months)": "",
                    "Estimated Lead Time (Days)": "", "Unit List Price": 0.0,
                    "Pricing Term": "", "Qty": 1, "Unit Net Price": 0.0,
                    "Disc(%)": 0.0, "Extended Net Price": 0.0
                })
        if new_rows:
            st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame(new_rows)], ignore_index=True)
            st.success(f"Se agregaron {len(new_rows)} SKU. Completa los datos en la pestaña Cotización desde Excel.")
        else:
            st.warning("No hay SKU nuevos para agregar.")

    if not st.session_state.data.empty:
        st.dataframe(st.session_state.data[["Part Number", "Description", "Unit List Price", "Disc(%)", "Unit Net Price"]], use_container_width=True)

with tab_compare:
    st.subheader("Comparar cotización anterior contra nueva")
    old_file = st.file_uploader("Cotización anterior (.xlsx)", type=["xlsx"], key="old_file")
    new_file = st.file_uploader("Cotización nueva (.xlsx)", type=["xlsx"], key="new_file")
    if old_file and new_file:
        try:
            old_df, _ = parse_cisco_quote(old_file)
            new_df, _ = parse_cisco_quote(new_file)
            old_map = old_df.groupby("Part Number")["Unit List Price"].last()
            new_map = new_df.groupby("Part Number")["Unit List Price"].last()
            keys = sorted(set(old_map.index) | set(new_map.index))
            comp = pd.DataFrame({"Part Number": keys})
            comp["Precio anterior"] = comp["Part Number"].map(old_map)
            comp["Precio nuevo"] = comp["Part Number"].map(new_map)
            comp["Variación absoluta"] = comp["Precio nuevo"] - comp["Precio anterior"]
            comp["Variación (%)"] = ((comp["Precio nuevo"] / comp["Precio anterior"]) - 1) * 100
            comp["Estado"] = comp.apply(
                lambda r: "Nuevo SKU" if pd.isna(r["Precio anterior"]) else
                          ("Ya no aparece" if pd.isna(r["Precio nuevo"]) else
                           ("Sin cambio" if round(r["Variación absoluta"], 2) == 0 else "Cambio de precio")), axis=1)
            st.dataframe(comp, use_container_width=True)
            st.download_button("Descargar comparación CSV", comp.to_csv(index=False).encode("utf-8-sig"),
                               "Comparacion_precios_Cisco.csv", "text/csv")
        except Exception as e:
            st.error(f"No pude comparar los archivos: {e}")

st.divider()
st.caption("Cisco Price Manager · Versión inicial · Los importes solo son válidos como datos cargados por el usuario hasta conectar una fuente oficial Cisco.")
