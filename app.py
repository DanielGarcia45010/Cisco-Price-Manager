from __future__ import annotations

import io
import re
from collections import defaultdict
from copy import copy
from datetime import datetime
from typing import Any

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# CISCO PRICE MANAGER - COTIZACION CISCO -> CALCULADORA
# ============================================================

st.set_page_config(
    page_title="Cisco Price Manager",
    page_icon="🌐",
    layout="wide",
)

APP_TITLE = "Cisco Price Manager"
SCENARIOS = (12, 24, 36, 48)
QUOTE_ALIASES = {
    "part_number": ["Part Number", "PartNumber", "SKU", "Product Number", "Product Number/Part Number"],
    "description": ["Description", "Product Description", "Item Description"],
    "duration": ["Service Duration (Months)", "Service Duration", "Duration (Months)", "ServiceDurationMonths"],
    "list_price": ["Unit List Price", "UnitListPrice", "List Price", "ListPrice"],
    "quantity": ["Qty", "Quantity"],
    "discount": ["Disc(%)", "Disc (%)", "Discount", "Discount (%)", "Discount(%)"],
    "unit_net": ["Unit Net Price", "UnitNetPrice", "Net Unit Price"],
    "extended": ["Extended Net Price", "ExtendedNetPrice", "Extended Price", "ExtendedNetAmount"],
    "pricing_term": ["Pricing Term", "PricingTerm"],
}


# -----------------------------
# Utilidades de lectura
# -----------------------------

def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text.replace("\n", " ")


def as_number(value: Any, default: float | None = None) -> float | None:
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text or text in {"---", "-", "n/a", "na", "none"}:
        return default
    text = text.replace("$", "").replace("USD", "").replace(" ", "")
    # Admite 1,234.56 y 1.234,56.
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts[-1]) in (1, 2):
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")
    try:
        return float(text)
    except (TypeError, ValueError):
        return default


def parse_duration(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value) if value > 0 else None
    match = re.search(r"(\d+(?:[.,]\d+)?)", str(value))
    if not match:
        return None
    return as_number(match.group(1).replace(",", "."))


def classify_part_number(part_number: str) -> str:
    """Regla acordada: intangible si el PN contiene CON o LIC."""
    pn = (part_number or "").upper()
    return "Intangible" if ("CON" in pn or "LIC" in pn) else "Tangible"


def find_quote_table(workbook_bytes: bytes) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Detecta automáticamente la hoja y fila de encabezados de la cotización Cisco."""
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=True, read_only=True)
    selected = None
    header_row = None
    column_map: dict[str, int] = {}

    aliases = {key: {normalize_text(x) for x in values} for key, values in QUOTE_ALIASES.items()}
    for ws in wb.worksheets:
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 40), values_only=True), start=1):
            normalized = {normalize_text(v): i for i, v in enumerate(row) if v is not None}
            pn_col = next((normalized[a] for a in aliases["part_number"] if a in normalized), None)
            desc_col = next((normalized[a] for a in aliases["description"] if a in normalized), None)
            if pn_col is not None and desc_col is not None:
                selected, header_row = ws, row_idx
                for field, names in aliases.items():
                    found = next((normalized[a] for a in names if a in normalized), None)
                    if found is not None:
                        column_map[field] = found
                break
        if selected is not None:
            break

    if selected is None or header_row is None:
        raise ValueError("No encontré una tabla con las columnas Part Number y Description en el archivo Cisco.")

    records: list[dict[str, Any]] = []
    for row in selected.iter_rows(min_row=header_row + 1, values_only=True):
        pn_value = row[column_map["part_number"]] if column_map["part_number"] < len(row) else None
        if pn_value is None or not str(pn_value).strip():
            # La tabla de Cisco termina en su primera fila vacía; debajo puede
            # haber totales/notas que no son productos.
            break
        pn = str(pn_value).strip()
        desc = row[column_map["description"]] if "description" in column_map and column_map["description"] < len(row) else ""
        qty = as_number(row[column_map["quantity"]], 1.0) if "quantity" in column_map and column_map["quantity"] < len(row) else 1.0
        qty = qty if qty is not None else 1.0
        list_price = as_number(row[column_map["list_price"]]) if "list_price" in column_map and column_map["list_price"] < len(row) else None
        discount = as_number(row[column_map["discount"]], 0.0) if "discount" in column_map and column_map["discount"] < len(row) else 0.0
        unit_net = as_number(row[column_map["unit_net"]]) if "unit_net" in column_map and column_map["unit_net"] < len(row) else None
        extended = as_number(row[column_map["extended"]]) if "extended" in column_map and column_map["extended"] < len(row) else None
        if unit_net is None and list_price is not None:
            unit_net = list_price * (1 - (discount or 0.0) / 100.0)
        if extended is None and unit_net is not None:
            extended = unit_net * qty
        if extended is None:
            extended = 0.0
        duration_value = row[column_map["duration"]] if "duration" in column_map and column_map["duration"] < len(row) else None
        duration = parse_duration(duration_value)
        pricing_term = row[column_map["pricing_term"]] if "pricing_term" in column_map and column_map["pricing_term"] < len(row) else None
        records.append({
            "part_number": pn,
            "description": str(desc or "").strip(),
            "duration_raw": duration_value,
            "duration": duration,
            "quantity": qty,
            "list_price": list_price,
            "discount": discount or 0.0,
            "unit_net": unit_net,
            "extended": float(extended),
            "pricing_term": str(pricing_term or "").strip(),
            "classification": classify_part_number(pn),
        })

    wb.close()
    if not records:
        raise ValueError("La hoja contiene encabezados, pero no encontré productos debajo de ellos.")
    metadata = {"sheet": selected.title, "header_row": header_row, "line_count": len(records)}
    return records, metadata


# -----------------------------
# Proyecciones de servicios
# -----------------------------

def service_projection(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    projections = []
    for item in records:
        if item["classification"] != "Intangible":
            continue
        base = float(item["extended"] or 0.0)
        duration = item["duration"]
        values = {}
        for months in SCENARIOS:
            if duration and duration > 0:
                values[months] = base * months / duration
                rule = f"Proporcional desde {duration:g} meses"
            else:
                values[months] = base
                rule = "Sin duración: se conserva el valor original"
        projections.append({**item, "projection": values, "projection_rule": rule})
    return projections


def aggregate_tangibles(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, dict[str, Any]] = {}
    for item in records:
        if item["classification"] != "Tangible":
            continue
        pn = item["part_number"]
        if pn not in grouped:
            grouped[pn] = {
                "part_number": pn,
                "description": item["description"],
                "quantity": 0.0,
                "extended": 0.0,
            }
        grouped[pn]["quantity"] += float(item["quantity"] or 0.0)
        grouped[pn]["extended"] += float(item["extended"] or 0.0)
    return list(grouped.values())


# -----------------------------
# Construcción del Excel
# -----------------------------

def copy_row_style(ws, source_row: int, target_row: int, max_col: int = 11) -> None:
    for col in range(1, max_col + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.protection:
            target.protection = copy(source.protection)
    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def get_country_options(wb) -> list[str]:
    ws = wb["CPEs & Aduanas"]
    values = []
    for row in range(15, ws.max_row + 1):
        value = ws.cell(row, 2).value
        if value:
            values.append(str(value))
    return values


def make_calculator_sheet(wb, sheet_name: str, products: list[dict[str, Any]], months: int,
                          shipping_total: float, margin: float, return_rate: float,
                          nrc_rate: float, markup: float, country: str,
                          service_total: float) -> None:
    ws = wb[sheet_name]
    original_first, original_last = 9, 18
    count = len(products)
    if count == 0:
        products = [{"part_number": "SIN PRODUCTOS TANGIBLES", "description": "", "quantity": 0, "extended": 0.0}]
        count = 1

    extra = max(0, count - (original_last - original_first + 1))
    if extra:
        ws.insert_rows(20, amount=extra)

    # Parámetros generales existentes en la plantilla.
    ws["B1"] = nrc_rate
    ws["B2"] = return_rate
    ws["B3"] = months
    ws["B4"] = margin
    ws["B6"] = markup
    ws["D5"] = country

    # El vínculo externo del libro original no se puede resolver desde esta app.
    # La fila externa se desplaza si se insertan productos; se limpia más abajo.
    external_row = 28 + extra
    ws.cell(external_row, 1).value = "External CPE Calculation (not supplied)"
    ws.cell(external_row, 2).value = 0

    first_row = 9
    last_row = first_row + count - 1
    for idx, product in enumerate(products):
        r = first_row + idx
        if r > 18:
            copy_row_style(ws, 18, r)
        ws.cell(r, 1).value = product["part_number"]
        ws.cell(r, 2).value = round(float(product["extended"] or 0.0), 2)
        # Shipping & Handling se introduce como costo total del quote, una sola vez.
        ws.cell(r, 3).value = round(float(shipping_total), 2) if idx == 0 else 0.0
        # Se aplica la tasa de aduanas al costo tangible de cada SKU.
        ws.cell(r, 4).value = f"=$B$5*B{r}"
        ws.cell(r, 5).value = f"=SUM(B{r}:D{r})"
        ws.cell(r, 6).value = f"=B{r}/(1-$B$4)"
        ws.cell(r, 7).value = f"=C{r}*(1+$B$6)"
        ws.cell(r, 8).value = f"=D{r}*(1+$B$6)"
        ws.cell(r, 9).value = f"=SUM(F{r}:H{r})"
        ws.cell(r, 10).value = f"=(G{r}+H{r})*1.1"
        ws.cell(r, 11).value = f"=-PMT($B$2/12,$B$3,(I{r}-J{r}),0)"
        for col in (2, 3, 4, 5, 6, 7, 8, 9, 10, 11):
            ws.cell(r, col).number_format = '$#,##0.00;[Red]-$#,##0.00'

    # Limpia las filas sobrantes de la tabla original cuando hay menos de 10 SKU.
    for r in range(first_row + count, 19):
        for col in range(1, 12):
            ws.cell(r, col).value = None

    summary_row = 21 + extra
    ws.cell(summary_row, 2).value = f"=SUM(E{first_row}:E{last_row})"
    # La celda SFP/CPE externa de la plantilla se reutiliza como costo de servicios
    # proyectado para el plazo seleccionado.
    ws.cell(summary_row - 1, 3).value = "Services"
    ws.cell(summary_row, 3).value = round(float(service_total), 2)
    ws.cell(summary_row, 4).value = f"=C{summary_row}+B{summary_row}"

    # Total Sales Price: equipo (incluye shipping/aduanas) + servicios con margen.
    ws.cell(summary_row + 1, 2).value = f"=SUM(I{first_row}:I{last_row})"
    ws.cell(summary_row + 1, 3).value = f"=C{summary_row}/(1-$B$4)"
    ws.cell(summary_row + 1, 4).value = f"=SUM(B{summary_row + 1}:C{summary_row + 1})"

    ws.cell(summary_row + 2, 2).value = f"=SUM(J{first_row}:J{last_row})"
    ws.cell(summary_row + 2, 3).value = f"=(C{summary_row}/(1-$B$4))*$B$1"
    ws.cell(summary_row + 2, 4).value = f"=C{summary_row + 2}+B{summary_row + 2}"
    ws.cell(summary_row + 3, 2).value = f"=SUM(K{first_row}:K{last_row})"
    ws.cell(summary_row + 3, 3).value = f"=-PMT($B$2/12,$B$3,C{summary_row + 2}/20%*80%,0)"
    ws.cell(summary_row + 3, 4).value = f"=C{summary_row + 3}+B{summary_row + 3}"

    # Etiquetas auxiliares de trazabilidad.
    ws["A7"] = "UNIT COST - Cisco tangible products"
    ws["A8"] = "Part Number"
    ws["B8"] = "Equipment Cost"
    ws["C8"] = "Shipping & Handling Cost"
    ws["D8"] = "Custom and Duties Cost"
    ws["E8"] = "Total Cost"
    ws["F8"] = "Equipment Price"
    ws["G8"] = "Shipping Price"
    ws["H8"] = "Custom & Duties Price"
    ws["I8"] = "Total Price"
    ws["J8"] = "NRC"
    ws["K8"] = "MRC"


def build_output_workbook(template_bytes: bytes, records: list[dict[str, Any]],
                          shipping_total: float, margin: float, return_rate: float,
                          nrc_rate: float, markup: float, country: str) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes), data_only=False)
    if "Calculadora" not in wb.sheetnames or "CPEs & Aduanas" not in wb.sheetnames:
        raise ValueError("La plantilla debe contener las hojas 'Calculadora' y 'CPEs & Aduanas'.")

    # El archivo original incluye una referencia a otro libro ('CPE Calculation (8475)').
    # Se elimina el vínculo externo porque no se ha cargado ese libro auxiliar.
    wb._external_links = []

    products = aggregate_tangibles(records)
    services = service_projection(records)
    service_totals = {m: sum(x["projection"][m] for x in services) for m in SCENARIOS}

    base_name = "Calculadora"
    wb[base_name].title = "Calculadora 12M"
    for months in SCENARIOS[1:]:
        new_ws = wb.copy_worksheet(wb["Calculadora 12M"])
        new_ws.title = f"Calculadora {months}M"

    for months in SCENARIOS:
        make_calculator_sheet(
            wb, f"Calculadora {months}M", products, months,
            shipping_total, margin, return_rate, nrc_rate, markup,
            country, service_totals[months]
        )

    # Hoja de proyección de servicios/intangibles.
    if "Proyeccion Servicios" in wb.sheetnames:
        del wb["Proyeccion Servicios"]
    ws = wb.create_sheet("Proyeccion Servicios")
    headers = ["Part Number", "Description", "Qty", "Service Duration (Months)",
               "Cisco Extended Net Price", "12 Months", "24 Months", "36 Months", "48 Months", "Projection Rule"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for item in services:
        ws.append([
            item["part_number"], item["description"], item["quantity"],
            item["duration_raw"], round(item["extended"], 2),
            round(item["projection"][12], 2), round(item["projection"][24], 2),
            round(item["projection"][36], 2), round(item["projection"][48], 2),
            item["projection_rule"],
        ])
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=9):
        for cell in row:
            cell.number_format = '$#,##0.00;[Red]-$#,##0.00'
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    widths = [25, 65, 10, 22, 22, 16, 16, 16, 16, 44]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Hoja de detalle de todos los renglones de la cotización.
    if "Detalle Cisco" in wb.sheetnames:
        del wb["Detalle Cisco"]
    detail = wb.create_sheet("Detalle Cisco")
    detail_headers = ["Part Number", "Description", "Classification", "Qty", "Service Duration (Months)",
                      "Unit List Price", "Discount (%)", "Unit Net Price", "Extended Net Price", "Pricing Term"]
    detail.append(detail_headers)
    for cell in detail[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for item in records:
        detail.append([
            item["part_number"], item["description"], item["classification"], item["quantity"],
            item["duration_raw"], item["list_price"], item["discount"], item["unit_net"],
            round(item["extended"], 2), item["pricing_term"],
        ])
    for row in detail.iter_rows(min_row=2, min_col=6, max_col=9):
        for cell in row:
            cell.number_format = '$#,##0.00;[Red]-$#,##0.00'
    detail.freeze_panes = "A2"
    detail.auto_filter.ref = detail.dimensions
    detail_widths = [27, 68, 17, 10, 24, 18, 15, 18, 20, 18]
    for i, width in enumerate(detail_widths, start=1):
        detail.column_dimensions[get_column_letter(i)].width = width

    # Fuerza el recálculo de fórmulas al abrir en Excel.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# -----------------------------
# Interfaz Streamlit
# -----------------------------

st.title("🌐 Cisco Price Manager")
st.subheader("Generador de calculadora a partir de una cotización Cisco")
st.markdown(
    "Carga el Excel descargado de Cisco Commerce y la plantilla de la calculadora. "
    "La aplicación clasifica los productos por Part Number, proyecta los intangibles "
    "a 12, 24, 36 y 48 meses y genera un libro con las cuatro calculadoras."
)

with st.expander("Reglas aplicadas", expanded=False):
    st.markdown(
        "- **Intangible:** el Part Number contiene `CON` o `LIC` (sin distinguir mayúsculas/minúsculas).\n"
        "- **Tangible:** todos los demás Part Number. La descripción no cambia esta clasificación.\n"
        "- **Equipo:** el costo tangible tomado de Cisco se mantiene igual en los cuatro escenarios.\n"
        "- **Servicios/intangibles:** si Cisco indica duración, el valor se escala proporcionalmente; "
        "si no indica duración, se conserva el valor original en todos los escenarios.\n"
        "- Las proyecciones de 24, 36 y 48 meses son estimaciones, no cotizaciones oficiales de Cisco."
    )

col1, col2 = st.columns(2)
with col1:
    quote_file = st.file_uploader("1. Cotización Cisco Commerce (.xlsx)", type=["xlsx"], key="quote")
with col2:
    template_file = st.file_uploader("2. Plantilla de calculadora (.xlsx)", type=["xlsx"], key="template")

st.markdown("### Parámetros de la calculadora")
p1, p2, p3, p4, p5 = st.columns(5)
with p1:
    nrc_rate = st.number_input("NRC", min_value=0.0, max_value=1.0, value=0.20, step=0.01, format="%.2f")
with p2:
    return_rate = st.number_input("Return Rate", min_value=0.0, max_value=1.0, value=0.10, step=0.01, format="%.2f")
with p3:
    margin = st.number_input("Margin", min_value=0.0, max_value=0.95, value=0.20, step=0.01, format="%.2f")
with p4:
    markup = st.number_input("Markup", min_value=0.0, max_value=5.0, value=0.10, step=0.01, format="%.2f")
with p5:
    shipping_total = st.number_input("Shipping & Handling total", min_value=0.0, value=500.0, step=50.0)

country_options = ["Belize"]
if template_file is not None:
    try:
        temp_wb = load_workbook(io.BytesIO(template_file.getvalue()), data_only=False, read_only=True)
        if "CPEs & Aduanas" in temp_wb.sheetnames:
            found = get_country_options(temp_wb)
            if found:
                country_options = found
        temp_wb.close()
    except Exception:
        pass
country_default_index = country_options.index("Belize") if "Belize" in country_options else 0
country = st.selectbox("Country (Customs & Duties)", country_options, index=country_default_index)

if quote_file is not None:
    try:
        preview_records, preview_meta = find_quote_table(quote_file.getvalue())
        tangible_count = sum(1 for x in preview_records if x["classification"] == "Tangible")
        intangible_count = len(preview_records) - tangible_count
        tangible_total = sum(x["extended"] for x in preview_records if x["classification"] == "Tangible")
        intangible_total = sum(x["extended"] for x in preview_records if x["classification"] == "Intangible")
        st.markdown("### Resumen de la cotización detectada")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Renglones", len(preview_records))
        m2.metric("Tangibles", tangible_count)
        m3.metric("Intangibles", intangible_count)
        m4.metric("Costo tangible Cisco", f"${tangible_total:,.2f}")
        st.caption(f"Hoja detectada: {preview_meta['sheet']} · Encabezados en fila {preview_meta['header_row']}")
        st.caption(f"Valor Cisco de intangibles antes de proyección: ${intangible_total:,.2f}")
        with st.expander("Vista previa de clasificación", expanded=False):
            st.dataframe(
                [{"Part Number": x["part_number"], "Description": x["description"],
                  "Classification": x["classification"], "Qty": x["quantity"],
                  "Duration": x["duration_raw"], "Extended Net Price": x["extended"]}
                 for x in preview_records],
                use_container_width=True,
                hide_index=True,
            )
    except Exception as exc:
        st.error(f"No pude leer la cotización: {exc}")

if st.button("Generar calculadora Excel", type="primary", use_container_width=True,
              disabled=(quote_file is None or template_file is None)):
    try:
        records, metadata = find_quote_table(quote_file.getvalue())
        output_bytes = build_output_workbook(
            template_file.getvalue(), records, shipping_total, margin,
            return_rate, nrc_rate, markup, country
        )
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Cisco_Price_Manager_{timestamp}.xlsx"
        st.success("Calculadora generada correctamente.")
        st.download_button(
            "Descargar Excel generado",
            data=output_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.info(
            "El archivo contiene Calculadora 12M, 24M, 36M y 48M, además de las hojas "
            "Proyeccion Servicios y Detalle Cisco. Excel recalculará las fórmulas al abrir el archivo."
        )
        st.warning(
            "La plantilla original tenía un vínculo externo a 'CPE Calculation (8475)'. "
            "Como ese archivo auxiliar no se carga aquí, el vínculo se eliminó y su entrada se dejó en cero "
            "para evitar referencias rotas. Revisa ese componente si tu oferta requiere costos SFP/CPE adicionales."
        )
    except Exception as exc:
        st.error(f"No fue posible generar el Excel: {exc}")
        st.exception(exc)

st.divider()
st.caption("Cisco Price Manager · Procesamiento local del archivo cargado · No requiere credenciales ni API de Cisco")
