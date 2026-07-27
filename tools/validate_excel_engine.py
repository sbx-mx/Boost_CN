#!/usr/bin/env python3
"""Valida integridad, relaciones y métricas generadas desde Boost_CN.xlsx."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"


def load(name):
    with (DATA / name).open(encoding="utf-8") as handle:
        return json.load(handle)


def require(condition, message):
    if not condition:
        raise SystemExit(f"ERROR: {message}")


manifest = load("manifest-data.json")
source = ROOT / manifest["sourceFile"]
require(source.is_file(), "no existe el Excel fuente")
require(
    hashlib.sha256(source.read_bytes()).hexdigest() == manifest["sourceSha256"],
    "los JSON no corresponden al Excel actual",
)
require(manifest["engine"] == "excel-only", "el manifiesto no declara motor Excel")
require(
    manifest["sheetsRead"] == ["Instruccion", "Base_Boost_CN", "Condicion", "Base_AT", "Item_ADT"],
    "no se leyeron las cinco pestañas requeridas en su orden",
)

records = []
for chunk in manifest["chunks"]:
    require((DATA / chunk).is_file(), f"falta el fragmento {chunk}")
    records.extend(load(chunk))
require(len(records) == manifest["dataRows"], "el total de registros no coincide")
require(
    sum(bool(row.get("Valid")) for row in records) == manifest["validRows"],
    "el total de registros válidos no coincide",
)
require(
    round(sum(float(row.get("UnidadVendida") or 0) for row in records), 6)
    == manifest["totals"]["units"],
    "el total de unidades no coincide",
)
require(
    round(sum(float(row.get("Venta") or 0) for row in records), 2)
    == manifest["totals"]["sales"],
    "el total de venta no coincide",
)

required_sales = {
    "Region", "DM", "Anio", "Semana", "Dia", "DayPart", "CeCo", "Tienda",
    "Categoria", "ProductoId", "Producto", "UnidadVendida", "Item",
    "PrecioVenta", "Venta", "Valid",
}
require(all(required_sales <= set(row) for row in records), "faltan campos en ventas")

at_rows = load("base-at.json")
adt_rows = load("item-adt.json")
require(len(at_rows) == manifest["baseATDataRows"], "el total Base_AT no coincide")
require(len(adt_rows) == manifest["itemADTDataRows"], "el total Item_ADT no coincide")
require(
    all(
        1 <= row["Semana"] <= 53
        and row["CeCo"]
        and row["Tienda"]
        and any(row[key] is not None for key in ("ATR", "ATAA", "ATPpto"))
        for row in at_rows
        if row["Valid"]
    ),
    "Base_AT contiene semanas, tiendas o métricas inválidas",
)
require(
    all(
        1 <= row["Semana"] <= 53
        and row["CeCo"]
        and row["Tienda"]
        and all(row[key] is not None for key in ("TPLH", "ItemADT"))
        for row in adt_rows
        if row["Valid"]
    ),
    "Item_ADT contiene semanas, tiendas o métricas inválidas",
)
require(
    manifest["baseATMatchedRows"] == manifest["baseATValidRows"],
    "hay registros AT válidos que no se relacionan con ventas",
)
require(
    manifest["itemADTMatchedRows"] == manifest["itemADTValidRows"],
    "hay registros de productividad válidos que no se relacionan con ventas",
)
require(
    set(manifest["metrics"])
    == {"Venta", "UnidadVendida", "ATR", "ATAA", "ATPpto", "TPLH", "ItemADT"},
    "se perdió una o más métricas requeridas",
)
require(load("audit-summary.json") == manifest, "audit-summary y manifest difieren")
require((DATA / "workbook-profile.json").is_file(), "falta el perfil del Excel")

print(
    "VALIDACIÓN APROBADA · "
    f"{manifest['dataRows']} ventas · "
    f"{manifest['baseATDataRows']} AT · "
    f"{manifest['itemADTDataRows']} productividad · "
    "7 métricas"
)
