from __future__ import annotations

import io
import re
from copy import copy
from datetime import datetime
from typing import Any

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


# ============================================================
# CISCO PRICE MANAGER
# Una sola hoja de calculadora con escenarios seleccionables.
# No requiere credenciales ni API de Cisco.
# ============================================================

st.set_page_config(
    page_title="Cisco Price Manager",
    page_icon="🌐",
    layout="wide",
)

COUNTRY_RATES = {
    "Anguilla": 0.15,
    "Antigua": 0.15,
    "Aruba": 0.18,
    "Bahamas": 0.60,
    "Barbados": 0.15,
    "Belize": 0.20,
    "BVI": 0.15,
    "Cayman": 0.15,
    "Colombia": 0.26,
    "Costa Rica": 0.14,
    "Curacao": 0.17,
    "Dominica": 0.15,
    "Dominican Rep": 0.18,
    "Ecuador": 0.34,
    "El Salvador": 0.18,
    "Grenada": 0.15,
    "Guatemala": 0.18,
    "Haiti": 0.18,
    "Honduras": 0.18,
    "Jamaica": 0.26,
    "Mexico": 0.18,
    "Nicaragua": 0.18,
    "Panama": 0.15,
    "Peru": 0.34,
    "Puerto Rico": 0.15,
    "St Kitts": 0.15,
    "St Lucia": 0.15,
    "St Marteen": 0.17,
    "St Vincent": 0.15,
    "Trinidad": 0.15,
    "Turks and Caicos": 0.15,
    "USVI": 0.15,
}

# Los modelos que aparecen en la tabla principal de la calculadora original.
# Se incluyen exclusivamente equipos; no accesorios, fuentes, memorias ni licencias.
HARDWARE_SKUS = {
    "C1121-4P",
    "C1121X-8P",
    "C1121-8P",
    "C8200L-1N-4T",
    "C8300-1N1S-6T",
    "C8300-2N2S-6T",
}

QUOTE_ALIASES = {
    "part_number": ["Part Number", "PartNumber", "SKU", "Product Number"],
    "description": ["Description", "Product Description", "Item Description"],
    "duration": ["Service Duration (Months)", "Service Duration", "Duration (Months)"],
    "list_price": ["Unit List Price", "List Price"],
    "quantity": ["Qty", "Quantity"],
    "discount": ["Disc(%)", "Disc (%)", "Discount", "Discount (%)"],
    "unit_net": ["Unit Net Price", "Net Unit Price"],
    "extended": ["Extended Net Price", "Extended Price"],
    "pricing_term": ["Pricing Term"],
}


def norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def number(value: Any, default: float | None = None) -> float | None:
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("$", "").replace("USD", "").replace(" ", "")
    if not s or s.lower() in {"-", "---", "n/a", "na", "none"}:
        return default
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        s = s.replace(",", ".") if len(parts[-1]) in (1, 2) else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return default


def parse_months(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value) if value > 0 else None
    m = re.search(r"(\d+(?:[.,]\d+)?)", str(value))
    return number(m.group(1).replace(",", ".")) if m else None


def read_cisco_quote(file_bytes: bytes) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    selected_ws = None
    header_row = None
    cmap: dict[str, int] = {}

    alias_norm = {k: {norm(x) for x in v} for k, v in QUOTE_ALIASES.items()}
    for ws in wb.worksheets:
        for r_idx, row in enumerate(
            ws.iter_rows(min_row=1, max_row=min(ws.max_row, 50), values_only=True),
            start=1,
        ):
            found = {norm(v): i for i, v in enumerate(row) if v is not None}
            if any(a in found for a in alias_norm["part_number"]) and any(
                a in found for a in alias_norm["description"]
            ):
                selected_ws, header_row = ws, r_idx
                for field, aliases in alias_norm.items():
                    col = next((found[a] for a in aliases if a in found), None)
                    if col is not None:
                        cmap[field] = col
                break
        if selected_ws:
            break

    if selected_ws is None or header_row is None:
        wb.close()
        raise ValueError("No encontré las columnas Part Number y Description en la cotización.")

    rows: list[dict[str, Any]] = []
    for row in selected_ws.iter_rows(min_row=header_row + 1, values_only=True):
        pn_idx = cmap["part_number"]
        pn_value = row[pn_idx] if pn_idx < len(row) else None
        if pn_value is None or not str(pn_value).strip():
            break

        def get(field: str, default=None):
            idx = cmap.get(field)
            return row[idx] if idx is not None and idx < len(row) else default

        qty = number(get("quantity"), 1.0) or 1.0
        list_price = number(get("list_price"))
        discount = number(get("discount"), 0.0) or 0.0
        unit_net = number(get("unit_net"))
        extended = number(get("extended"))
        if unit_net is None and list_price is not None:
            unit_net = list_price * (1 - discount / 100)
        if extended is None and unit_net is not None:
            extended = unit_net * qty

        pn = str(pn_value).strip()
        rows.append({
            "part_number": pn,
            "description": str(get("description", "") or "").strip(),
            "quantity": qty,
            "duration_raw": get("duration"),
            "duration": parse_months(get("duration")),
            "list_price": list_price,
            "discount": discount,
            "unit_net": unit_net,
            "extended": float(extended or 0.0),
            "pricing_term": str(get("pricing_term", "") or "").strip(),
        })

    wb.close()
    if not rows:
        raise ValueError("No se encontraron productos debajo de los encabezados.")
    return rows, {"sheet": selected_ws.title, "header_row": header_row, "line_count": len(rows)}


def identify_equipment(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    Mantiene cada aparición física del equipo como renglón separado.
    El sufijo '(Secured)' se asigna cuando en el mismo grupo de configuración
    aparece un SKU de seguridad SL-...-SEC-SV.
    """
    equipment: list[dict[str, Any]] = []
    current_group: list[dict[str, Any]] = []

    def flush_group(group: list[dict[str, Any]]) -> None:
        if not group:
            return
        has_security = any(
            "SEC-SV" in str(x["part_number"]).upper()
            or "SECURITY" in str(x["description"]).upper()
            for x in group
        )
        for x in group:
            pn = x["part_number"].upper()
            if pn in HARDWARE_SKUS:
                label = f"{x['part_number']} (Secured)" if has_security else x["part_number"]
                equipment.append({
                    "item": label,
                    "part_number": x["part_number"],
                    "description": x["description"],
                    "quantity": x["quantity"],
                    "equipment_cost": float(x["extended"] or 0.0),
                })

    for item in rows:
        pn = item["part_number"].upper()
        if pn in HARDWARE_SKUS:
            # Cada SKU de hardware inicia una configuración nueva.
            flush_group(current_group)
            current_group = [item]
        elif current_group:
            current_group.append(item)
    flush_group(current_group)
    return equipment


def service_projection_totals(rows: list[dict[str, Any]], months: int) -> tuple[float, list[dict[str, Any]]]:
    """
    Proyecta únicamente renglones clasificados por la regla acordada:
    Part Number contiene CON o LIC. La duración se toma de Cisco.
    """
    detail = []
    total = 0.0
    for x in rows:
        pn = x["part_number"].upper()
        if "CON" not in pn and "LIC" not in pn:
            continue
        base = float(x["extended"] or 0.0)
        duration = x["duration"]
        if duration and duration > 0:
            projected = base * months / duration
            rule = f"Proporcional desde {duration:g} meses"
        else:
            projected = base
            rule = "Sin duración: conserva el valor original"
        total += projected
        detail.append({**x, "projected": projected, "rule": rule})
    return total, detail


def clone_row_style(ws, source_row: int, target_row: int, max_col: int = 11) -> None:
    for col in range(1, max_col + 1):
        src, dst = ws.cell(source_row, col), ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
    if source_row in ws.row_dimensions and ws.row_dimensions[source_row].height:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def copy_cell_style(source, target) -> None:
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    if source.alignment:
        target.alignment = copy(source.alignment)


def parse_scenario_months(raw: str) -> list[int]:
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if not parts:
        raise ValueError("Indica al menos un plazo, por ejemplo: 12,24,36.")
    values = []
    for p in parts:
        if not re.fullmatch(r"\d{1,3}", p):
            raise ValueError(f"El plazo '{p}' no es válido. Usa meses enteros separados por comas.")
        n = int(p)
        if n < 1 or n > 600:
            raise ValueError("Los plazos deben estar entre 1 y 600 meses.")
        if n not in values:
            values.append(n)
    if len(values) > 8:
        raise ValueError("Puedes generar máximo 8 escenarios en una misma hoja.")
    return values


def create_output(template_bytes: bytes, quote_rows: list[dict[str, Any]],
                  months_list: list[int], country: str, shipping_per_equipment: float,
                  margin: float, return_rate: float, markup: float) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes), data_only=False)

    if "Calculadora" not in wb.sheetnames:
        raise ValueError("La plantilla debe contener la hoja 'Calculadora'.")

    # Conservar únicamente la hoja que el usuario solicitó.
    ws = wb["Calculadora"]
    for sheet_name in list(wb.sheetnames):
        if sheet_name != "Calculadora":
            del wb[sheet_name]
    ws.title = "Calculadora"
    wb._external_links = []

    equipment = identify_equipment(quote_rows)
    if not equipment:
        raise ValueError(
            "No encontré equipos principales de la lista de la calculadora "
            "(C1121-4P, C1121X-8P, C1121-8P, C8200L-1N-4T, C8300-1N1S-6T, C8300-2N2S-6T)."
        )

    # Países y tasas embebidos en columnas ocultas de la misma hoja.
    # Así se conserva la lista desplegable sin crear la hoja 'CPEs & Aduanas'.
    country_start = 2
    ws["M1"] = "Country list"
    ws["N1"] = "Customs rate"
    for i, (name, rate) in enumerate(COUNTRY_RATES.items(), start=country_start):
        ws.cell(i, 13).value = name
        ws.cell(i, 14).value = rate
        ws.cell(i, 14).number_format = "0%"
    ws.column_dimensions["M"].hidden = True
    ws.column_dimensions["N"].hidden = True

    # Usar el bloque original de la calculadora como plantilla visual.
    base_block_height = 32
    first_product_row = 9
    max_products = len(equipment)
    if max_products > 10:
        ws.insert_rows(19, amount=max_products - 10)

    # Guardar estilos del bloque original para replicarlo verticalmente.
    style_rows = {}
    for r in range(1, 31):
        style_rows[r] = [copy(ws.cell(r, c)._style) if ws.cell(r, c).has_style else None
                         for c in range(1, 15)]
    row_heights = {r: ws.row_dimensions[r].height for r in range(1, 31)}

    # Primera sección usa el bloque original. Los escenarios siguientes se copian
    # en la misma hoja, uno debajo del otro.
    block_height = max(31, 24 + max(0, max_products - 10) + 3)
    required_end = block_height * len(months_list)
    if required_end > ws.max_row:
        ws.insert_rows(ws.max_row + 1, amount=required_end - ws.max_row)

    # Replicar estilos de filas del bloque original para los escenarios posteriores.
    for idx in range(1, len(months_list)):
        offset = idx * block_height
        for r in range(1, 31):
            target_r = r + offset
            for c in range(1, 15):
                src_cell = ws.cell(r, c)
                dst_cell = ws.cell(target_r, c)
                copy_cell_style(src_cell, dst_cell)
                if row_heights.get(r) is not None:
                    ws.row_dimensions[target_r].height = row_heights[r]

    # Encabezado visual de cada escenario y fórmulas de calculadora.
    for idx, months in enumerate(months_list):
        off = idx * block_height
        r1 = 1 + off
        # Parámetros generales de la plantilla.
        ws.cell(r1, 1).value = "NRC"
        ws.cell(r1, 2).value = 0.20  # Valor original de la plantilla; no requiere cálculo manual.
        ws.cell(r1 + 1, 1).value = "Return Rate"
        ws.cell(r1 + 1, 2).value = return_rate
        ws.cell(r1 + 2, 1).value = "Term (Months)"
        ws.cell(r1 + 2, 2).value = months
        ws.cell(r1 + 3, 1).value = "Marging"
        ws.cell(r1 + 3, 2).value = margin
        ws.cell(r1 + 4, 1).value = "Customs & Duties. %"
        ws.cell(r1 + 4, 2).value = f'=IFERROR(INDEX($N$2:$N${len(COUNTRY_RATES)+1},MATCH(D{r1+4},$M$2:$M${len(COUNTRY_RATES)+1},0)),0)'
        ws.cell(r1 + 4, 3).value = "Country"
        ws.cell(r1 + 4, 4).value = country
        ws.cell(r1 + 5, 1).value = "Markup"
        ws.cell(r1 + 5, 2).value = markup

        # Título que identifica el escenario sin agregar otra hoja.
        ws.cell(r1 + 6, 1).value = f"CALCULATOR PROJECTION — {months} MONTHS"
        ws.cell(r1 + 6, 1).font = Font(bold=True, size=12, color="17365D")

        headers = [
            "Item", "Equipment Cost", "Shipping & Handling Cost",
            "Custom and Duties Cost", "Total Cost", "Equipment Price",
            "Shipping Price", "Custom & Duties Price", "Total Price", "NRC", "MRC"
        ]
        for c, value in enumerate(headers, start=1):
            cell = ws.cell(r1 + 7, c)
            cell.value = value
            cell.font = copy(ws["A8"].font) if ws["A8"].has_style else Font(bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="center")

        first = r1 + 8
        last = first + len(equipment) - 1

        for j, item in enumerate(equipment):
            r = first + j
            if j >= 10:
                clone_row_style(ws, r1 + 17, r)
            ws.cell(r, 1).value = item["item"]
            ws.cell(r, 2).value = round(item["equipment_cost"], 2)
            ws.cell(r, 3).value = round(shipping_per_equipment, 2)
            ws.cell(r, 4).value = f"=$B${r1+4}*B{r}"
            ws.cell(r, 5).value = f"=SUM(B{r}:D{r})"
            ws.cell(r, 6).value = f"=B{r}/(1-$B${r1+3})"
            ws.cell(r, 7).value = f"=C{r}*(1+$B${r1+5})"
            ws.cell(r, 8).value = f"=D{r}*(1+$B${r1+5})"
            ws.cell(r, 9).value = f"=SUM(F{r}:H{r})"
            # Fórmula NRC de la calculadora original. No se pide al usuario multiplicar nada.
            ws.cell(r, 10).value = f"=(G{r}+H{r})*1.1"
            ws.cell(r, 11).value = f"=-PMT($B${r1+1}/12,$B${r1+2},(I{r}-J{r}),0)"
            for c in range(2, 12):
                ws.cell(r, c).number_format = '$#,##0.00;[Red]-$#,##0.00'

        # Limpia las filas sobrantes del bloque original.
        for r in range(first + len(equipment), first + max(10, len(equipment))):
            for c in range(1, 12):
                ws.cell(r, c).value = None

        # Resumen de costos y precios. La celda de servicios sustituye el cálculo
        # externo de SFP/CPE que dependía de un libro no proporcionado.
        summary = r1 + 20 + max(0, len(equipment) - 10)
        service_total, _ = service_projection_totals(quote_rows, months)
        ws.cell(summary - 1, 2).value = "Equipment"
        ws.cell(summary - 1, 3).value = "Services"
        ws.cell(summary, 1).value = "Total Cost"
        ws.cell(summary, 2).value = f"=SUM(E{first}:E{last})"
        ws.cell(summary, 3).value = round(service_total, 2)
        ws.cell(summary, 4).value = f"=B{summary}+C{summary}"

        ws.cell(summary + 1, 1).value = "Total Sales Price"
        ws.cell(summary + 1, 2).value = f"=SUM(I{first}:I{last})"
        ws.cell(summary + 1, 3).value = f"=C{summary}/(1-$B${r1+3})"
        ws.cell(summary + 1, 4).value = f"=SUM(B{summary+1}:C{summary+1})"

        ws.cell(summary + 2, 1).value = "NRC"
        ws.cell(summary + 2, 2).value = f"=SUM(J{first}:J{last})"
        ws.cell(summary + 2, 3).value = f"=(C{summary}/(1-$B${r1+3}))*$B${r1}"
        ws.cell(summary + 2, 4).value = f"=B{summary+2}+C{summary+2}"

        ws.cell(summary + 3, 1).value = "MRC"
        ws.cell(summary + 3, 2).value = f"=SUM(K{first}:K{last})"
        ws.cell(summary + 3, 3).value = f"=-PMT($B${r1+1}/12,$B${r1+2},C{summary+2}/20%*80%,0)"
        ws.cell(summary + 3, 4).value = f"=B{summary+3}+C{summary+3}"

        for rr in range(summary, summary + 4):
            for cc in range(2, 5):
                ws.cell(rr, cc).number_format = '$#,##0.00;[Red]-$#,##0.00'

        # Eliminar cualquier referencia externa remanente de la plantilla.
        for rr in range(r1, summary + 5):
            for cc in range(1, 15):
                val = ws.cell(rr, cc).value
                if isinstance(val, str) and "CPE Calculation" in val:
                    ws.cell(rr, cc).value = None

        # Lista desplegable de países en cada bloque.
        dv = DataValidation(
            type="list",
            formula1=f"=$M$2:$M${len(COUNTRY_RATES)+1}",
            allow_blank=False,
        )
        dv.error = "Selecciona un país de la lista."
        dv.errorTitle = "País no válido"
        dv.prompt = "Selecciona el país para calcular aduanas."
        dv.promptTitle = "País"
        ws.add_data_validation(dv)
        dv.add(ws.cell(r1 + 4, 4))

    # Dejar una única hoja y mostrar solo el área de la calculadora.
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A8"
    ws.column_dimensions["A"].width = 27
    for col in "BCDEFGHIJK":
        ws.column_dimensions[col].width = 19
    ws.column_dimensions["C"].width = 23
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["M"].hidden = True
    ws.column_dimensions["N"].hidden = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_area = f"A1:K{block_height * len(months_list)}"

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ============================================================
# INTERFAZ
# ============================================================

st.title("🌐 Cisco Price Manager")
st.subheader("Calculadora Cisco en una sola hoja de Excel")
st.write(
    "Carga la cotización de Cisco Commerce y tu plantilla. "
    "La aplicación genera una única hoja llamada **Calculadora**, con los escenarios "
    "de meses que indiques y la lista desplegable de países para aduanas."
)

with st.expander("Reglas de cálculo", expanded=False):
    st.markdown(
        "- La tabla principal muestra únicamente los modelos de equipo de la calculadora original.\n"
        "- Cada aparición del mismo SKU se conserva como un renglón independiente.\n"
        "- Los accesorios, fuentes, memorias y licencias no se muestran como equipos.\n"
        "- **Custom and Duties Cost = tasa del país × Equipment Cost** en cada renglón.\n"
        "- NRC conserva el parámetro y las fórmulas de la plantilla; no tienes que multiplicarlo manualmente por 0.2.\n"
        "- Las proyecciones de servicios se calculan según la duración indicada por Cisco."
    )

c1, c2 = st.columns(2)
with c1:
    quote_file = st.file_uploader("1. Cotización de Cisco Commerce (.xlsx)", type=["xlsx"])
with c2:
    template_file = st.file_uploader("2. Plantilla original de calculadora (.xlsx)", type=["xlsx"])

st.markdown("### Parámetros")
p1, p2, p3, p4 = st.columns(4)
with p1:
    months_raw = st.text_input("Plazos a proyectar (meses)", value="12,24,36,48",
                               help="Escribe los plazos separados por comas. Ejemplo: 12,18,24,36.")
with p2:
    country = st.selectbox("País para aduanas", list(COUNTRY_RATES.keys()),
                           index=list(COUNTRY_RATES.keys()).index("Belize"))
with p3:
    shipping_per_equipment = st.number_input(
        "Shipping & Handling por equipo (USD)", min_value=0.0, value=500.0, step=50.0
    )
with p4:
    margin = st.number_input("Margin", min_value=0.0, max_value=0.95, value=0.20,
                             step=0.01, format="%.2f")

p5, p6 = st.columns(2)
with p5:
    return_rate = st.number_input("Return Rate", min_value=0.0, max_value=1.0,
                                  value=0.10, step=0.01, format="%.2f")
with p6:
    markup = st.number_input("Markup", min_value=0.0, max_value=5.0,
                             value=0.10, step=0.01, format="%.2f")

if quote_file is not None:
    try:
        preview_rows, meta = read_cisco_quote(quote_file.getvalue())
        preview_equipment = identify_equipment(preview_rows)
        st.markdown("### Equipos identificados")
        st.metric("Renglones de equipos", len(preview_equipment))
        st.dataframe(
            [{
                "Item": x["item"],
                "Part Number Cisco": x["part_number"],
                "Equipment Cost": x["equipment_cost"],
                "Qty": x["quantity"],
            } for x in preview_equipment],
            use_container_width=True,
            hide_index=True,
        )
        st.caption(f"Cotización detectada: hoja {meta['sheet']} · fila de encabezados {meta['header_row']}")
    except Exception as exc:
        st.error(f"No pude leer la cotización: {exc}")

if st.button("Generar Excel — una sola hoja", type="primary", use_container_width=True,
             disabled=(quote_file is None or template_file is None)):
    try:
        months_list = parse_scenario_months(months_raw)
        quote_rows, meta = read_cisco_quote(quote_file.getvalue())
        output_bytes = create_output(
            template_file.getvalue(),
            quote_rows,
            months_list,
            country,
            shipping_per_equipment,
            margin,
            return_rate,
            markup,
        )
        filename = f"Cisco_Calculadora_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        st.success("Excel generado. El archivo contiene únicamente la hoja Calculadora.")
        st.download_button(
            "Descargar calculadora Excel",
            data=output_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.info(
            f"Escenarios incluidos en una sola hoja: {', '.join(str(x) for x in months_list)} meses. "
            "En cada bloque puedes cambiar el país desde la lista desplegable; la tasa de aduanas "
            "se actualiza automáticamente."
        )
    except Exception as exc:
        st.error(f"No fue posible generar el Excel: {exc}")
        st.exception(exc)

st.divider()
st.caption("Cisco Price Manager · Procesamiento local · No utiliza API ni credenciales de Cisco")
