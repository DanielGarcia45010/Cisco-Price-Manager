from __future__ import annotations

import io
import re
from copy import copy
from datetime import datetime
from typing import Any

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation


# ============================================================
# CISCO PRICE MANAGER
# Una sola hoja Calculadora con plazo seleccionable en Excel.
# ============================================================

st.set_page_config(page_title="Cisco Price Manager", page_icon="🌐", layout="wide")

COUNTRY_RATES = {
    "Anguilla": 0.15, "Antigua": 0.15, "Aruba": 0.18, "Bahamas": 0.60,
    "Barbados": 0.15, "Belize": 0.20, "BVI": 0.15, "Cayman": 0.15,
    "Colombia": 0.26, "Costa Rica": 0.14, "Curacao": 0.17, "Dominica": 0.15,
    "Dominican Rep": 0.18, "Ecuador": 0.34, "El Salvador": 0.18, "Grenada": 0.15,
    "Guatemala": 0.18, "Haiti": 0.18, "Honduras": 0.18, "Jamaica": 0.26,
    "Mexico": 0.18, "Nicaragua": 0.18, "Panama": 0.15, "Peru": 0.34,
    "Puerto Rico": 0.15, "St Kitts": 0.15, "St Lucia": 0.15, "St Marteen": 0.17,
    "St Vincent": 0.15, "Trinidad": 0.15, "Turks and Caicos": 0.15, "USVI": 0.15,
}

HARDWARE_SKUS = {
    "C1121-4P", "C1121X-8P", "C1121-8P",
    "C8200L-1N-4T", "C8300-1N1S-6T", "C8300-2N2S-6T",
}

QUOTE_ALIASES = {
    "part_number": ["Part Number", "PartNumber", "SKU", "Product Number"],
    "description": ["Description", "Product Description", "Item Description"],
    "duration": ["Service Duration (Months)", "Service Duration", "Duration (Months)"],
    "quantity": ["Qty", "Quantity"],
    "unit_net": ["Unit Net Price", "Net Unit Price"],
    "extended": ["Extended Net Price", "Extended Price"],
    "list_price": ["Unit List Price", "List Price"],
    "discount": ["Disc(%)", "Disc (%)", "Discount", "Discount (%)"],
}


def norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def to_number(value: Any, default: float | None = None) -> float | None:
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("$", "").replace("USD", "").replace(" ", "")
    if not s or s.lower() in {"-", "---", "n/a", "na", "none"}:
        return default
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        tail = s.split(",")[-1]
        s = s.replace(",", ".") if len(tail) in (1, 2) else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return default


def parse_duration(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value) if value > 0 else None
    match = re.search(r"(\d+(?:[.,]\d+)?)", str(value))
    if not match:
        return None
    result = to_number(match.group(1).replace(",", "."))
    return result if result and result > 0 else None


def read_cisco_quote(file_bytes: bytes) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    aliases = {k: {norm(v) for v in vals} for k, vals in QUOTE_ALIASES.items()}
    selected_ws = None
    header_row = None
    columns: dict[str, int] = {}

    for ws in wb.worksheets:
        for row_num, values in enumerate(
            ws.iter_rows(min_row=1, max_row=min(ws.max_row, 50), values_only=True), start=1
        ):
            found = {norm(v): i for i, v in enumerate(values) if v is not None}
            if (found.keys() & aliases["part_number"]) and (found.keys() & aliases["description"]):
                selected_ws, header_row = ws, row_num
                for field, names in aliases.items():
                    columns[field] = next((found[n] for n in names if n in found), -1)
                break
        if selected_ws:
            break

    if selected_ws is None or header_row is None:
        wb.close()
        raise ValueError("No encontré las columnas Part Number y Description en la cotización.")

    products = []
    for values in selected_ws.iter_rows(min_row=header_row + 1, values_only=True):
        pn_col = columns["part_number"]
        pn_value = values[pn_col] if pn_col >= 0 and pn_col < len(values) else None
        if pn_value is None or not str(pn_value).strip():
            # La tabla de Cisco termina en la primera fila sin Part Number.
            break

        def get(field: str):
            col = columns.get(field, -1)
            return values[col] if col >= 0 and col < len(values) else None

        qty = to_number(get("quantity"), 1.0) or 1.0
        list_price = to_number(get("list_price"))
        discount = to_number(get("discount"), 0.0) or 0.0
        unit_net = to_number(get("unit_net"))
        extended = to_number(get("extended"))
        if unit_net is None and list_price is not None:
            unit_net = list_price * (1 - discount / 100)
        if extended is None and unit_net is not None:
            extended = unit_net * qty

        products.append({
            "part_number": str(pn_value).strip(),
            "description": str(get("description") or "").strip(),
            "quantity": qty,
            "duration": parse_duration(get("duration")),
            "duration_raw": get("duration"),
            "extended": float(extended or 0.0),
        })

    wb.close()
    if not products:
        raise ValueError("No se encontraron productos en la cotización.")
    return products, {
        "sheet": selected_ws.title,
        "header_row": header_row,
        "line_count": len(products),
    }


def group_quote_by_equipment(products: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Cada fila principal de hardware inicia un grupo. Todos los renglones de
    Cisco Commerce hasta el siguiente hardware pertenecen a ese equipo.
    El coste del grupo incluye tangibles e intangibles. Solo los renglones
    con Service Duration > 0 se escalan con el plazo seleccionado.
    """
    groups: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None

    for product in products:
        pn = product["part_number"].upper()
        if pn in HARDWARE_SKUS:
            if current is not None:
                groups.append(current)
            current = {
                "part_number": product["part_number"],
                "description": product["description"],
                "hardware_cost": float(product["extended"]),
                "components": [product],
                "secured": False,
            }
        elif current is not None:
            current["components"].append(product)
            if "SEC-SV" in pn or "SECURITY" in product["description"].upper():
                current["secured"] = True

    if current is not None:
        groups.append(current)

    return groups


def equipment_cost_at_term(group: dict[str, Any], months: int) -> float:
    total = 0.0
    for component in group["components"]:
        value = float(component["extended"] or 0.0)
        duration = component["duration"]
        if duration and duration > 0:
            value *= months / duration
        total += value
    return round(total, 2)


def add_dropdown(ws, cell_ref: str, formula: str, prompt: str, error: str) -> None:
    validation = DataValidation(type="list", formula1=formula, allow_blank=False)
    validation.promptTitle = "Selecciona una opción"
    validation.prompt = prompt
    validation.errorTitle = "Valor no válido"
    validation.error = error
    validation.showErrorMessage = True
    validation.showInputMessage = True
    ws.add_data_validation(validation)
    validation.add(ws[cell_ref])


def copy_style(source, target) -> None:
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    if source.alignment:
        target.alignment = copy(source.alignment)


def create_output(
    template_bytes: bytes,
    quote_products: list[dict[str, Any]],
    country: str,
    months: int,
    nrc_rate: float,
    shipping_per_equipment: float,
    margin: float,
    return_rate: float,
    markup: float,
) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes), data_only=False)
    if "Calculadora" not in wb.sheetnames:
        raise ValueError("La plantilla debe contener la hoja 'Calculadora'.")

    ws = wb["Calculadora"]
    # Guardar los nombres de los 10 equipos tal como están en la plantilla.
    original_labels = [ws.cell(r, 1).value for r in range(9, 19)]

    # Solo se conserva la hoja Calculadora.
    for name in list(wb.sheetnames):
        if name != "Calculadora":
            del wb[name]
    wb._external_links = []

    groups = group_quote_by_equipment(quote_products)
    if not groups:
        raise ValueError("No encontré equipos principales en la cotización.")
    if len(groups) > 10:
        raise ValueError(
            f"La cotización contiene {len(groups)} equipos principales y la plantilla tiene "
            "10 filas disponibles. Amplía la plantilla antes de continuar."
        )

    # Inputs principales de la calculadora.
    ws["B1"] = nrc_rate
    ws["B2"] = return_rate
    ws["B3"] = months
    ws["B4"] = margin
    ws["B5"] = '=IFERROR(INDEX($N$2:$N$33,MATCH(D5,$M$2:$M$33,0)),0)'
    ws["D5"] = country
    ws["B6"] = markup

    # Lista de países y tasas en columnas ocultas de la misma hoja.
    ws["M1"], ws["N1"] = "Country", "Customs Rate"
    for row, (name, rate) in enumerate(COUNTRY_RATES.items(), start=2):
        ws.cell(row, 13).value = name
        ws.cell(row, 14).value = rate
        ws.cell(row, 14).number_format = "0%"

    # Tabla de soporte oculta con cada componente de cada equipo.
    # P = número de equipo, Q = SKU, R = descripción, S = valor neto,
    # T = duración en meses (0 significa que no se proyecta).
    headers = ["Equipment Index", "Part Number", "Description", "Extended Net Price", "Service Duration"]
    for col, header in enumerate(headers, start=16):
        ws.cell(1, col).value = header

    support_row = 2
    for equipment_idx, group in enumerate(groups, start=1):
        for component in group["components"]:
            ws.cell(support_row, 16).value = equipment_idx
            ws.cell(support_row, 17).value = component["part_number"]
            ws.cell(support_row, 18).value = component["description"]
            ws.cell(support_row, 19).value = float(component["extended"] or 0.0)
            ws.cell(support_row, 20).value = float(component["duration"] or 0.0)
            ws.cell(support_row, 19).number_format = '$#,##0.00'
            support_row += 1

    # Base física del router, separada de licencias/servicios para aduanas.
    ws["U1"], ws["V1"] = "Equipment Index", "Physical Equipment Cost"
    for equipment_idx, group in enumerate(groups, start=1):
        ws.cell(equipment_idx + 1, 21).value = equipment_idx
        ws.cell(equipment_idx + 1, 22).value = group["hardware_cost"]
        ws.cell(equipment_idx + 1, 22).number_format = '$#,##0.00'

    last_support = max(2, support_row - 1)
    # Fórmula B: componentes sin duración + componentes con duración escalados.
    # Fórmula D: tasa de aduanas × solo el precio físico del router.
    for row in range(9, 19):
        idx = row - 8
        if idx <= len(groups):
            group = groups[idx - 1]
            # Conservar el nombre/etiqueta original de la calculadora cuando
            # corresponde al mismo SKU; si no, usar la identificación de la cotización.
            template_label = original_labels[idx - 1] if idx - 1 < len(original_labels) else None
            if template_label and str(template_label).upper().startswith(group["part_number"].upper()):
                label = template_label
            else:
                label = group["part_number"] + (" (Secured)" if group["secured"] else "")
            ws.cell(row, 1).value = label
            ws.cell(row, 2).value = (
                f'=SUMIFS($S$2:$S${last_support},$P$2:$P${last_support},{idx},$T$2:$T${last_support},0)'
                f'+SUMPRODUCT(($P$2:$P${last_support}={idx})*($T$2:$T${last_support}>0)'
                f'*$S$2:$S${last_support}*IFERROR($B$3/$T$2:$T${last_support},0))'
            )
            ws.cell(row, 3).value = shipping_per_equipment
            ws.cell(row, 4).value = f'=$B$5*INDEX($V$2:$V${len(groups)+1},ROW()-8)'
            ws.cell(row, 2).number_format = '$#,##0.00;[Red]-$#,##0.00'
            ws.cell(row, 4).number_format = '$#,##0.00;[Red]-$#,##0.00'
        else:
            for col in range(1, 12):
                ws.cell(row, col).value = None

    # Mantener las fórmulas originales de la plantilla para E:K.
    # Si se encuentran vacías, reconstruirlas siguiendo las fórmulas originales.
    for row in range(9, 19):
        if ws.cell(row, 1).value is None:
            continue
        ws.cell(row, 5).value = f"=SUM(B{row}:D{row})"
        ws.cell(row, 6).value = f"=B{row}/(1-$B$4)"
        ws.cell(row, 7).value = f"=C{row}*(1+$B$6)"
        ws.cell(row, 8).value = f"=D{row}*(1+$B$6)"
        ws.cell(row, 9).value = f"=SUM(F{row}:H{row})"
        ws.cell(row, 10).value = f"=(G{row}+H{row})*1.1"
        ws.cell(row, 11).value = f"=-PMT($B$2/12,$B$3,(I{row}-J{row}),0)"

    # Resumen: los costes con duración se muestran en Services, pero ya están
    # incluidos en Equipment Cost por dispositivo. Se separan en el resumen
    # para evitar sumarlos dos veces.
    ws["B20"] = "Equipment & Other Costs"
    ws["C20"] = "Projected Services"
    ws["D20"] = "Total"
    ws["B21"] = "=SUM(E9:E18)-C21"
    ws["C21"] = (
        f'=SUMPRODUCT(($T$2:$T${last_support}>0)*$S$2:$S${last_support}'
        f'*IFERROR($B$3/$T$2:$T${last_support},0))'
    )
    ws["D21"] = "=SUM(B21:C21)"

    ws["A22"] = "Total Sales Price"
    ws["B22"] = "=SUM(I9:I18)-C22"
    ws["C22"] = "=IFERROR(C21/(1-$B$4),0)"
    ws["D22"] = "=SUM(B22:C22)"

    ws["A23"] = "NRC"
    ws["B23"] = "=SUM(J9:J18)"
    ws["C23"] = "=IFERROR((C21/(1-$B$4))*$B$1,0)"
    ws["D23"] = "=SUM(B23:C23)"

    ws["A24"] = "MRC"
    ws["B24"] = "=SUM(K9:K18)-C24"
    ws["C24"] = "=-PMT($B$2/12,$B$3,C23/20%*80%,0)"
    ws["D24"] = "=SUM(B24:C24)"

    for row in range(21, 25):
        for col in range(2, 5):
            ws.cell(row, col).number_format = '$#,##0.00;[Red]-$#,##0.00'

    # Validaciones desplegables: país y plazo de proyección.
    add_dropdown(
        ws, "D5", "=$M$2:$M$33",
        "Selecciona el país para calcular aduanas.",
        "Selecciona un país de la lista.",
    )
    add_dropdown(
        ws, "B3", '"12,24,36,48"',
        "Selecciona el plazo de la proyección en meses.",
        "Selecciona 12, 24, 36 o 48 meses.",
    )

    # Ocultar tablas auxiliares dentro de la misma hoja; no crear hojas adicionales.
    for col in range(13, 23):
        ws.column_dimensions[chr(64 + col)].hidden = True

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A9"
    ws["B1"].number_format = "0%"
    ws["B2"].number_format = "0%"
    ws["B4"].number_format = "0%"
    ws["B5"].number_format = "0%"
    ws["B6"].number_format = "0%"
    ws["D5"].alignment = Alignment(horizontal="left")
    ws["B3"].number_format = '0 "months"'

    # Excel recalcula todas las fórmulas al abrir el archivo.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# STREAMLIT
# ============================================================

st.title("🌐 Cisco Price Manager")
st.subheader("Calculadora dinámica en una sola hoja")
st.write(
    "Carga la cotización de Cisco Commerce y la plantilla original. "
    "El Excel generado conserva únicamente la hoja **Calculadora**. "
    "En esa hoja podrás cambiar el país y el plazo desde listas desplegables."
)

with st.expander("Reglas de cálculo", expanded=False):
    st.markdown(
        "- **Equipment Cost:** suma el router y todos los componentes asociados, tangibles e intangibles.\n"
        "- Solo los renglones de la cotización con un valor en **Service Duration (Months)** "
        "se proyectan proporcionalmente al plazo seleccionado.\n"
        "- Los componentes sin duración conservan su valor original.\n"
        "- **Custom and Duties Cost:** tasa del país × precio físico del router, sin sumar licencias ni servicios.\n"
        "- El NRC es editable en la aplicación y también en la celda B1 del Excel."
    )

left, right = st.columns(2)
with left:
    quote_file = st.file_uploader("1. Cotización Cisco Commerce (.xlsx)", type=["xlsx"])
with right:
    template_file = st.file_uploader("2. Plantilla original (.xlsx)", type=["xlsx"])

st.markdown("### Parámetros iniciales")
p1, p2, p3 = st.columns(3)
with p1:
    country = st.selectbox("País para aduanas", list(COUNTRY_RATES), index=list(COUNTRY_RATES).index("Belize"))
with p2:
    months = st.selectbox("Plazo inicial de proyección", [12, 24, 36, 48], index=0)
with p3:
    nrc_rate = st.number_input("NRC", min_value=0.0, max_value=1.0, value=0.20, step=0.01, format="%.2f")

p4, p5, p6 = st.columns(3)
with p4:
    shipping_per_equipment = st.number_input(
        "Shipping & Handling por equipo (USD)", min_value=0.0, value=500.0, step=50.0
    )
with p5:
    margin = st.number_input("Margin", min_value=0.0, max_value=0.95, value=0.20, step=0.01, format="%.2f")
with p6:
    return_rate = st.number_input("Return Rate", min_value=0.0, max_value=1.0, value=0.10, step=0.01, format="%.2f")

markup = st.number_input("Markup", min_value=0.0, max_value=5.0, value=0.10, step=0.01, format="%.2f")

if quote_file is not None:
    try:
        preview_products, quote_meta = read_cisco_quote(quote_file.getvalue())
        preview_groups = group_quote_by_equipment(preview_products)
        st.markdown("### Vista previa de equipos")
        st.metric("Equipos identificados", len(preview_groups))
        st.dataframe(
            [{
                "Equipo": g["part_number"] + (" (Secured)" if g["secured"] else ""),
                "Equipment Cost (12 meses)": equipment_cost_at_term(g, 12),
                "Base para aduanas": g["hardware_cost"],
                "Componentes asociados": len(g["components"]),
            } for g in preview_groups],
            use_container_width=True,
            hide_index=True,
        )
        st.caption(
            f"Cotización: hoja '{quote_meta['sheet']}' · "
            f"fila de encabezados {quote_meta['header_row']} · "
            f"{quote_meta['line_count']} renglones leídos."
        )
    except Exception as exc:
        st.error(f"No pude leer la cotización: {exc}")

if st.button(
    "Generar calculadora Excel",
    type="primary",
    use_container_width=True,
    disabled=(quote_file is None or template_file is None),
):
    try:
        quote_products, _ = read_cisco_quote(quote_file.getvalue())
        groups = group_quote_by_equipment(quote_products)
        output_bytes = create_output(
            template_file.getvalue(),
            quote_products,
            country,
            months,
            nrc_rate,
            shipping_per_equipment,
            margin,
            return_rate,
            markup,
        )
        filename = f"Cisco_Calculadora_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.success("Calculadora generada correctamente. El archivo tiene una sola hoja.")
        st.download_button(
            "Descargar Excel",
            data=output_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        first_group = groups[0]
        st.info(
            f"Verificación del primer equipo: {first_group['part_number']} — "
            f"Equipment Cost a 12 meses: USD {equipment_cost_at_term(first_group, 12):,.2f}; "
            f"base de aduanas: USD {first_group['hardware_cost']:,.2f}. "
            "En Excel podrás cambiar el plazo en B3 y el país en D5."
        )
    except Exception as exc:
        st.error(f"No fue posible generar el Excel: {exc}")
        st.exception(exc)

st.divider()
st.caption("Cisco Price Manager · Procesamiento local · No utiliza API ni credenciales de Cisco")
