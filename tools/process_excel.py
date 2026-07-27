#!/usr/bin/env python3
"""Compila Boost_CN.xlsx como única fuente de datos del dashboard."""

from __future__ import annotations

import hashlib
import json
import os
import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
SOURCE = DATA / "Boost_CN.xlsx"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": MAIN, "r": REL}
VERSION = "10.0.0"
CHUNK_SIZE = 4000


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def norm(value):
    return "".join(
        char
        for char in unicodedata.normalize("NFD", clean(value).lower())
        if unicodedata.category(char) != "Mn"
    )


def number(value):
    text = clean(value).replace("$", "").replace(",", "").replace("%", "")
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def integer(value):
    value = number(value)
    return None if value is None else int(value)


def excel_date(value):
    if value in (None, ""):
        return ""
    try:
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).strftime("%Y-%m-%d")
    except (TypeError, ValueError, OverflowError):
        pass
    for date_format in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(clean(value), date_format).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def column_index(reference):
    index = 0
    for char in re.match(r"[A-Z]+", reference).group(0):
        index = index * 26 + ord(char) - 64
    return index - 1


def read_workbook(path):
    sheets = {}
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {item.attrib["Id"]: item.attrib["Target"] for item in relationships}
        shared_strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for string_item in root.findall("a:si", NS):
                shared_strings.append(
                    "".join(node.text or "" for node in string_item.iter(f"{{{MAIN}}}t"))
                )
        for sheet in workbook.find("a:sheets", NS):
            target = rel_map[sheet.attrib[f"{{{REL}}}id"]]
            target = target if target.startswith("xl/") else f"xl/{target.lstrip('/')}"
            rows = []
            xml_sheet = ET.fromstring(archive.read(target))
            for xml_row in xml_sheet.findall(".//a:sheetData/a:row", NS):
                cells = {}
                for cell in xml_row.findall("a:c", NS):
                    position = column_index(cell.attrib["r"])
                    cell_type = cell.attrib.get("t")
                    value_node = cell.find("a:v", NS)
                    value = "" if value_node is None else value_node.text
                    if cell_type == "s" and value != "":
                        value = shared_strings[int(value)]
                    elif cell_type == "inlineStr":
                        value = "".join(
                            node.text or "" for node in cell.iter(f"{{{MAIN}}}t")
                        )
                    cells[position] = value
                if cells:
                    rows.append([cells.get(i, "") for i in range(max(cells) + 1)])
            sheets[sheet.attrib["name"]] = rows
    return sheets


def records(sheets, sheet_name):
    rows = sheets[sheet_name]
    headers = [clean(value) for value in rows[0]]
    output = []
    for source_row in rows[1:]:
        padded = source_row + [""] * len(headers)
        output.append({headers[i]: padded[i] for i in range(len(headers))})
    return headers, output


def write_json(path, value, *, pretty=False):
    temporary = path.with_suffix(path.suffix + ".tmp")
    with temporary.open("w", encoding="utf-8", newline="\n") as handle:
        json.dump(
            value,
            handle,
            ensure_ascii=False,
            indent=2 if pretty else None,
            separators=None if pretty else (",", ":"),
        )
        handle.write("\n")
    os.replace(temporary, path)


def numeric_ratio(values, predicate):
    usable = [number(value) for value in values if clean(value)]
    if not usable:
        return 0.0
    return sum(1 for value in usable if value is not None and predicate(value)) / len(usable)


def text_ratio(values):
    usable = [clean(value) for value in values if clean(value)]
    if not usable:
        return 0.0
    return sum(1 for value in usable if number(value) is None) / len(usable)


def detect_base_at_layout(rows):
    sample = rows[: min(150, len(rows))]
    standard = (
        numeric_ratio([row.get("Semana") for row in sample], lambda value: 1 <= value <= 53)
        + numeric_ratio([row.get("Ceco") for row in sample], lambda value: 1000 <= value <= 999999)
        + text_ratio([row.get("Tienda") for row in sample])
    ) / 3
    shifted = (
        numeric_ratio([row.get("Semana") for row in sample], lambda value: 1000 <= value <= 999999)
        + text_ratio([row.get("Ceco") for row in sample])
        + numeric_ratio([row.get("Tienda") for row in sample], lambda value: 1 <= value <= 53)
    ) / 3
    repaired = shifted >= 0.90 and shifted > standard + 0.50
    return {
        "mode": "shifted_semana_ceco_tienda" if repaired else "standard",
        "repaired": repaired,
        "confidence": round(max(standard, shifted), 4),
        "standardScore": round(standard, 4),
        "shiftedScore": round(shifted, 4),
    }


def base_at_fields(row, layout):
    if layout["repaired"]:
        return row.get("Tienda"), row.get("Semana"), row.get("Ceco")
    return row.get("Semana"), row.get("Ceco"), row.get("Tienda")


if not SOURCE.is_file():
    raise SystemExit(f"No se encontró la fuente única: {SOURCE.relative_to(ROOT)}")

sheets = read_workbook(SOURCE)
required_sheets = ["Instruccion", "Base_Boost_CN", "Condicion", "Base_AT", "Item_ADT"]
missing_sheets = [name for name in required_sheets if name not in sheets]
if missing_sheets:
    raise SystemExit(f"Faltan pestañas requeridas: {missing_sheets}")

instruction_headers, instructions = records(sheets, "Instruccion")
base_headers, base = records(sheets, "Base_Boost_CN")
condition_headers, conditions = records(sheets, "Condicion")
at_headers, base_at = records(sheets, "Base_AT")
adt_headers, item_adt = records(sheets, "Item_ADT")

required_headers = {
    "Base_Boost_CN": [
        "Region", "DM", "Año", "Semana", "Dia", "DayPart", "CeCo",
        "Tienda", "Categoria", "#Producto", "Producto", "Unidad Vendida",
    ],
    "Condicion": ["Producto", "Item", "Venta"],
    "Base_AT": ["División", "DM", "Año", "Semana", "Ceco", "Tienda", "AT R", "AT AA", "AT ppto"],
    "Item_ADT": ["CeCo", "Tienda", "Semana", "DM", "TPLH", "item/ADT"],
}
headers_by_sheet = {
    "Base_Boost_CN": base_headers,
    "Condicion": condition_headers,
    "Base_AT": at_headers,
    "Item_ADT": adt_headers,
}
for sheet_name, expected in required_headers.items():
    missing = [header for header in expected if header not in headers_by_sheet[sheet_name]]
    if missing:
        raise SystemExit(
            f"{sheet_name}: faltan encabezados {missing}; encontrados {headers_by_sheet[sheet_name]}"
        )

exceptions = []
condition_map = {}
condition_duplicates = Counter()
invalid_sales = 0
for row_number, row in enumerate(conditions, 2):
    key = norm(row.get("Producto"))
    sale = number(row.get("Venta"))
    item = clean(row.get("Item"))
    condition_duplicates[key] += 1
    if not key:
        exceptions.append({"sheet": "Condicion", "row": row_number, "error": "Producto vacío"})
        continue
    if sale is None:
        invalid_sales += 1
        exceptions.append({
            "sheet": "Condicion",
            "row": row_number,
            "error": "Venta no numérica",
            "value": row.get("Venta"),
        })
    if key not in condition_map:
        condition_map[key] = {
            "Item": item,
            "PrecioVenta": sale,
            "Producto": clean(row.get("Producto")),
            "Fila": row_number,
        }
for key, count in condition_duplicates.items():
    if key and count > 1:
        exceptions.append({
            "sheet": "Condicion",
            "error": "Producto duplicado normalizado",
            "key": key,
            "count": count,
        })

sales_rows = []
unmatched_products = 0
max_date = ""
for row_number, row in enumerate(base, 2):
    product = clean(row.get("Producto"))
    match = condition_map.get(norm(product))
    date = excel_date(row.get("Dia"))
    units = number(row.get("Unidad Vendida"))
    output = {
        "Region": clean(row.get("Region")),
        "DM": clean(row.get("DM")),
        "Anio": integer(row.get("Año")),
        "Semana": integer(row.get("Semana")),
        "Dia": date,
        "DayPart": clean(row.get("DayPart")),
        "CeCo": clean(row.get("CeCo")).removesuffix(".0"),
        "Tienda": clean(row.get("Tienda")),
        "Categoria": clean(row.get("Categoria")),
        "ProductoId": clean(row.get("#Producto")).removesuffix(".0"),
        "Producto": product,
        "UnidadVendida": units,
        "Item": match["Item"] if match else "Sin relación",
        "PrecioVenta": match["PrecioVenta"] if match else None,
    }
    errors = []
    if not match:
        unmatched_products += 1
        errors.append("Producto sin relación en Condicion")
    if units is None:
        errors.append("Unidad Vendida inválida")
    if not date:
        errors.append("Fecha inválida")
    if output["Semana"] is None or not 1 <= output["Semana"] <= 53:
        errors.append("Semana inválida")
    if date and date > max_date:
        max_date = date
    output["Venta"] = round((units or 0) * (output["PrecioVenta"] or 0), 6)
    output["Valid"] = not errors
    if errors:
        exceptions.append({
            "sheet": "Base_Boost_CN",
            "row": row_number,
            "errors": errors,
            "Producto": product,
        })
    sales_rows.append(output)

at_layout = detect_base_at_layout(base_at)
if at_layout["repaired"]:
    exceptions.append({
        "sheet": "Base_AT",
        "type": "schema_repair",
        "message": "Encabezados desplazados detectados; Semana, CeCo y Tienda fueron realineados.",
        "details": at_layout,
    })

at_rows = []
at_keys = Counter()
for row_number, row in enumerate(base_at, 2):
    raw_week, raw_ceco, raw_store = base_at_fields(row, at_layout)
    output = {
        "Division": clean(row.get("División")),
        "DM": clean(row.get("DM")),
        "Anio": integer(row.get("Año")),
        "Semana": integer(raw_week),
        "CeCo": clean(raw_ceco).removesuffix(".0"),
        "Tienda": clean(raw_store),
        "ATR": number(row.get("AT R")),
        "ATAA": number(row.get("AT AA")),
        "ATPpto": number(row.get("AT ppto")),
    }
    output["Key"] = "|".join([
        output["CeCo"],
        str(output["Semana"] or ""),
        str(output["Anio"] or ""),
    ])
    identity_valid = bool(
        output["CeCo"]
        and output["Tienda"]
        and output["Semana"] is not None
        and 1 <= output["Semana"] <= 53
        and output["Anio"] is not None
    )
    available_metrics = [key for key in ("ATR", "ATAA", "ATPpto") if output[key] is not None]
    output["Valid"] = identity_valid and bool(available_metrics)
    if not output["Valid"]:
        exceptions.append({
            "sheet": "Base_AT",
            "row": row_number,
            "error": "Identidad AT inválida o sin métricas",
            "key": output["Key"],
        })
    elif len(available_metrics) < 3:
        exceptions.append({
            "sheet": "Base_AT",
            "row": row_number,
            "warning": "Comparativo AT incompleto; se conservaron las métricas disponibles",
            "key": output["Key"],
            "missingMetrics": [
                key for key in ("ATR", "ATAA", "ATPpto") if output[key] is None
            ],
        })
    at_keys[output["Key"]] += 1
    at_rows.append(output)

at_best = {}
for row in at_rows:
    score = int(row["Valid"]) * 10 + sum(row[key] is not None for key in ("ATR", "ATAA", "ATPpto"))
    previous = at_best.get(row["Key"])
    previous_score = (
        int(previous["Valid"]) * 10
        + sum(previous[key] is not None for key in ("ATR", "ATAA", "ATPpto"))
        if previous
        else -1
    )
    if score > previous_score:
        at_best[row["Key"]] = row
at_rows = list(at_best.values())
for key, count in at_keys.items():
    if count > 1:
        exceptions.append({
            "sheet": "Base_AT",
            "error": "Duplicado CeCo+Semana+Año consolidado",
            "key": key,
            "count": count,
        })

adt_rows = []
adt_keys = Counter()
for row_number, row in enumerate(item_adt, 2):
    output = {
        "CeCo": clean(row.get("CeCo")).removesuffix(".0"),
        "Tienda": clean(row.get("Tienda")),
        "Semana": integer(row.get("Semana")),
        "DM": clean(row.get("DM")),
        "TPLH": number(row.get("TPLH")),
        "ItemADT": number(row.get("item/ADT")),
    }
    output["Key"] = "|".join([output["CeCo"], str(output["Semana"] or "")])
    output["Valid"] = bool(
        output["CeCo"]
        and output["Tienda"]
        and output["Semana"] is not None
        and 1 <= output["Semana"] <= 53
        and output["TPLH"] is not None
        and output["ItemADT"] is not None
    )
    if not output["Valid"]:
        exceptions.append({
            "sheet": "Item_ADT",
            "row": row_number,
            "error": "Registro de productividad incompleto o inválido",
            "key": output["Key"],
        })
    adt_keys[output["Key"]] += 1
    adt_rows.append(output)

adt_best = {}
for row in adt_rows:
    score = int(row["Valid"]) * 10 + sum(row[key] is not None for key in ("TPLH", "ItemADT"))
    previous = adt_best.get(row["Key"])
    previous_score = (
        int(previous["Valid"]) * 10
        + sum(previous[key] is not None for key in ("TPLH", "ItemADT"))
        if previous
        else -1
    )
    if score > previous_score:
        adt_best[row["Key"]] = row
adt_rows = list(adt_best.values())
for key, count in adt_keys.items():
    if count > 1:
        exceptions.append({
            "sheet": "Item_ADT",
            "error": "Duplicado CeCo+Semana consolidado",
            "key": key,
            "count": count,
        })

for stale in DATA.glob("records-*.json"):
    stale.unlink()
chunks = []
for start in range(0, len(sales_rows), CHUNK_SIZE):
    filename = f"records-{start // CHUNK_SIZE + 1:03d}.json"
    chunks.append(filename)
    write_json(DATA / filename, sales_rows[start : start + CHUNK_SIZE])

source_hash = hashlib.sha256(SOURCE.read_bytes()).hexdigest()
sheet_profile = {
    name: {
        "rowsIncludingHeader": len(rows),
        "dataRows": max(len(rows) - 1, 0),
        "columns": len(rows[0]) if rows else 0,
        "headers": [clean(value) for value in rows[0]] if rows else [],
    }
    for name, rows in sheets.items()
}
base_keys = {
    f"{row['CeCo']}|{row['Semana']}|{row['Anio']}"
    for row in sales_rows
    if row["Valid"]
}
adt_base_keys = {
    f"{row['CeCo']}|{row['Semana']}"
    for row in sales_rows
    if row["Valid"]
}
at_matched = sum(row["Key"] in base_keys for row in at_rows if row["Valid"])
adt_matched = sum(row["Key"] in adt_base_keys for row in adt_rows if row["Valid"])

write_json(DATA / "base-at.json", at_rows)
write_json(DATA / "item-adt.json", adt_rows)
write_json(DATA / "condition.json", list(condition_map.values()))
write_json(DATA / "instructions.json", instructions)
write_json(DATA / "exceptions.json", exceptions, pretty=True)
write_json(DATA / "workbook-profile.json", sheet_profile, pretty=True)

manifest = {
    "version": VERSION,
    "engine": "excel-only",
    "sourceFile": "data/Boost_CN.xlsx",
    "sourceSha256": source_hash,
    "generatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
    "updatedDate": max_date,
    "sheetsRead": list(sheets),
    "sheetCount": len(sheets),
    "sheetProfile": sheet_profile,
    "chunks": chunks,
    "sourceRowsIncludingHeader": len(sheets["Base_Boost_CN"]),
    "dataRows": len(sales_rows),
    "loadedRows": len(sales_rows),
    "validRows": sum(row["Valid"] for row in sales_rows),
    "discardedRows": 0,
    "conditionRowsIncludingHeader": len(sheets["Condicion"]),
    "conditionDataRows": len(conditions),
    "conditionDuplicateKeys": sum(1 for count in condition_duplicates.values() if count > 1),
    "conditionInvalidSales": invalid_sales,
    "unmatchedProducts": unmatched_products,
    "baseATRowsIncludingHeader": len(sheets["Base_AT"]),
    "baseATSourceRows": len(base_at),
    "baseATDataRows": len(at_rows),
    "baseATValidRows": sum(row["Valid"] for row in at_rows),
    "baseATMatchedRows": at_matched,
    "baseATIncompleteMetricRows": sum(
        row["Valid"] and any(row[key] is None for key in ("ATR", "ATAA", "ATPpto"))
        for row in at_rows
    ),
    "atDuplicateKeys": sum(1 for count in at_keys.values() if count > 1),
    "baseATLayout": at_layout,
    "itemADTRowsIncludingHeader": len(sheets["Item_ADT"]),
    "itemADTSourceRows": len(item_adt),
    "itemADTDataRows": len(adt_rows),
    "itemADTValidRows": sum(row["Valid"] for row in adt_rows),
    "itemADTMatchedRows": adt_matched,
    "itemADTDuplicateKeys": sum(1 for count in adt_keys.values() if count > 1),
    "exceptionCount": len(exceptions),
    "totals": {
        "units": round(sum(row["UnidadVendida"] or 0 for row in sales_rows), 6),
        "sales": round(sum(row["Venta"] for row in sales_rows), 2),
    },
    "metrics": [
        "Venta", "UnidadVendida", "ATR", "ATAA", "ATPpto", "TPLH", "ItemADT",
    ],
}
write_json(DATA / "manifest-data.json", manifest, pretty=True)
write_json(DATA / "audit-summary.json", manifest, pretty=True)

audit_lines = [
    "BOOST CN v10.0 - AUDITORÍA DEL MOTOR EXCEL",
    "",
    "FUENTE ÚNICA",
    f"- Archivo: {manifest['sourceFile']}",
    f"- SHA-256: {source_hash}",
    f"- Pestañas leídas: {manifest['sheetCount']} ({', '.join(manifest['sheetsRead'])})",
    "",
    "MÉTRICAS CONSERVADAS",
    "- Venta, Unidades, AT Real, AT Año Anterior, AT Presupuesto, TPLH e Item/ADT.",
    "",
    "RESULTADO",
    f"- Registros de venta: {manifest['dataRows']} ({manifest['validRows']} válidos)",
    f"- Base AT: {manifest['baseATDataRows']} ({manifest['baseATMatchedRows']} relacionados)",
    f"- Productividad: {manifest['itemADTDataRows']} ({manifest['itemADTMatchedRows']} relacionados)",
    f"- Productos sin relación: {manifest['unmatchedProducts']}",
    f"- Excepciones registradas: {manifest['exceptionCount']}",
    f"- Unidades: {manifest['totals']['units']}",
    f"- Venta calculada: {manifest['totals']['sales']}",
    "",
    "NORMALIZACIÓN BASE_AT",
    f"- Modo: {at_layout['mode']}",
    f"- Reparación aplicada: {'Sí' if at_layout['repaired'] else 'No'}",
    f"- Confianza: {at_layout['confidence']:.2%}",
    "",
    "ACTUALIZACIÓN",
    "- Reemplazar únicamente data/Boost_CN.xlsx.",
    "- GitHub Actions regenera y valida todos los JSON automáticamente.",
    "- index.html no contiene datos ni requiere edición para actualizar métricas.",
]
(ROOT / "AUDITORIA_TECNICA.txt").write_text("\n".join(audit_lines) + "\n", encoding="utf-8")

print(json.dumps(manifest, ensure_ascii=False, indent=2))
