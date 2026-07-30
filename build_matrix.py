"""Build matrix XLSX from cleaned Almazar Document Intelligence JSON."""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from project_paths import OUTPUT_DIR, PROCESSING_DIR

CURRENCY = "EUR"

COST_NAME_ROW = 1
APPLY_IF_ROW = 2
RATE_BY_ROW = 3
BRACKET_ROW = 4
COLUMN_HEADER_ROW = 5
DATA_START_ROW = 6

HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
BOLD = Font(bold=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

SHIPMENT_HEADERS = [
    "Origin Postal Code Zone",
    "Destination Postal Code Zone",
]

ORIGIN_BARCELONA = "ES Barcelona"
ORIGIN_TENERIFE_WH = "ES Tenerife (Almacén)"
INSURANCE_FCL_40_LABEL = (
    "l[100% - Shipment Insurance (Carga EN CELRA/CORAL'40)%]"
)
EMERGENCY_BUNKER_NOTE = "TBD"

TENERIFE_DIST_BRACKETS = (
    "Hasta 150 kg",
    "Hasta 3000 kg",
    "Mas de 3000 kg",
    "Minimo",
)

FCL_SIZE_BRACKETS = ("20 pies", "40 pies")

ISLAND_DESTINATIONS = (
    "Las Palmas",
    "Lanzarote",
    "Fuerteventura",
    "La Palma",
    "La Gomera",
    "El Hierro",
)

GRUPAJE_DESTINATIONS = ("Las Palmas", "Lanzarote", "Tenerife")

FCL_ORIGINS = ("Las Palmas", "Tenerife")


@dataclass(frozen=True)
class CostColumnSpec:
    bracket_label: str
    rate_unit: str = "Flat"


@dataclass
class CostBlock:
    title: str
    apply_if: str
    rate_by: str
    columns: list[CostColumnSpec] = field(default_factory=list)
    uses_shared_currency: bool = False


@dataclass
class MatrixRow:
    shipment: dict[str, Any]
    costs: dict[tuple[str, str], float | str | None] = field(default_factory=dict)
    text_costs: dict[tuple[str, str], str] = field(default_factory=dict)


@dataclass
class SkippedItem:
    source: str
    label: str
    detail: str


def _norm_key(text: object) -> str:
    raw = unicodedata.normalize("NFKD", str(text or ""))
    raw = "".join(ch for ch in raw if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", "", raw.casefold())


def _row_value(row: dict[str, Any], key: str) -> Any:
    if key in row:
        return row[key]
    colon_key = f"{key}:"
    if colon_key in row:
        return row[colon_key]
    return None


def _parse_euro(value: object) -> float | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("€", "").replace(" ", "").strip()
    if text.endswith("%"):
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        whole, frac = text.split(",", 1)
        if len(frac) == 2:
            text = f"{whole}.{frac}"
        elif len(frac) == 3:
            if whole in ("0", "") or (whole.isdigit() and int(whole) < 10):
                text = f"{whole}.{frac}"
            else:
                text = f"{whole}.{frac[:2]}"
        else:
            text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def _zone_label(place: str) -> str:
    return f"ES {place.strip()}"


def _load_fields(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        raise ValueError("Expected JSON root object.")
    if "analyzeResult" in payload:
        from clean_di_json import clean_analyze_payload

        return clean_analyze_payload(payload)["documents"][0]["fields"]
    documents = payload.get("documents")
    if isinstance(documents, list) and documents:
        fields = documents[0].get("fields")
        if isinstance(fields, dict):
            return fields
    raise ValueError("JSON does not contain document fields.")


def _iter_sections(
    rows: list[dict[str, Any]], *, name_key: str = "RateName"
) -> Iterator[tuple[str, list[dict[str, Any]]]]:
    current_name: str | None = None
    buffer: list[dict[str, Any]] = []

    def flush() -> Iterator[tuple[str, list[dict[str, Any]]]]:
        nonlocal buffer, current_name
        if current_name is not None and buffer:
            yield current_name, buffer
        buffer = []

    for row in rows:
        rate_name = row.get(name_key)
        if rate_name:
            yield from flush()
            current_name = str(rate_name).strip()
            buffer = []
            continue
        if current_name is not None:
            buffer.append(row)
    yield from flush()


def _section_rows(
    rows: list[dict[str, Any]], *name_fragments: str
) -> list[dict[str, Any]]:
    targets = {_norm_key(fragment) for fragment in name_fragments}
    for section_name, section_rows in _iter_sections(rows):
        key = _norm_key(section_name)
        if any(target in key or key in target for target in targets):
            return section_rows
    return []


def _match_origin(origin: object, expected: str) -> bool:
    return _norm_key(origin) == _norm_key(expected)


def _match_weight(weight: object, expected: str) -> bool:
    return _norm_key(weight) == _norm_key(expected)


def _parse_fcl_table(section_rows: list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, float | None]]:
    """Map (origin, size) -> {cost1, cost2, cost3}."""
    table: dict[tuple[str, str], dict[str, float | None]] = {}
    for row in section_rows:
        origin = _row_value(row, "Origin")
        size = _row_value(row, "Value1")
        if not origin or not size:
            continue
        if _norm_key(origin) == _norm_key("Origen"):
            continue
        size_norm = _norm_key(size)
        if "20" in size_norm:
            size_key = "20 pies"
        elif "40" in size_norm:
            size_key = "40 pies"
        else:
            continue
        origin_key = str(origin).strip()
        table[(origin_key, size_key)] = {
            "cost1": _parse_euro(_row_value(row, "Cost1")),
            "cost2": _parse_euro(_row_value(row, "Cost2")),
            "cost3": _parse_euro(_row_value(row, "Cost3")),
        }
    return table


def _parse_grupajes(section_rows: list[dict[str, Any]]) -> dict[str, dict[str, float | None]]:
    rates: dict[str, dict[str, float | None]] = {}
    for row in section_rows:
        origin = _row_value(row, "Origin")
        if not origin:
            continue
        name = str(origin).strip()
        if _norm_key(name) in {
            _norm_key("Minimo"),
            _norm_key("Entrega /Reexp en destino"),
            _norm_key("Recargo Transporte 22/06/22"),
            _norm_key("Fuel (Fes)"),
            _norm_key("Despacho Origen"),
            _norm_key("Despacho Destino"),
        } or "seguro" in _norm_key(name):
            continue
        per_kg = _parse_euro(_row_value(row, "Cost2"))
        minimum = _parse_euro(_row_value(row, "Cost3"))
        if per_kg is None and minimum is None:
            continue
        rates[name] = {"per_kg": per_kg, "minimum": minimum}
    return rates


def _parse_tenerife_distribution(
    section_rows: list[dict[str, Any]],
) -> dict[str, float | None]:
    mapping: dict[str, float | None] = {}
    for row in section_rows:
        weight = _row_value(row, "Weight")
        price = _parse_euro(_row_value(row, "Price"))
        if not weight or price is None:
            continue
        w = str(weight).strip()
        if _match_weight(w, "Hasta 150 kg"):
            mapping["Hasta 150 kg"] = price
        elif _match_weight(w, "Hasta 3000 kg"):
            mapping["Hasta 3000 kg"] = price
        elif "3000" in _norm_key(w) and "mas" in _norm_key(w):
            mapping["Mas de 3000 kg"] = price
        elif "minimo" in _norm_key(w):
            mapping["Minimo"] = price
    return mapping


def _parse_wh_distribution(
    section_rows: list[dict[str, Any]],
) -> dict[str, float | None]:
    rates: dict[str, float | None] = {}
    for row in section_rows:
        dest = _row_value(row, "Weight")
        price = _parse_euro(_row_value(row, "Price"))
        if not dest or price is None:
            continue
        rates[str(dest).strip()] = price
    return rates


def _parse_otros_conceptos(section_rows: list[dict[str, Any]]) -> dict[str, str | float | None]:
    data: dict[str, str | float | None] = {}
    for row in section_rows:
        weight = _row_value(row, "Weight")
        price = _row_value(row, "Price")
        if not weight:
            continue
        key = _norm_key(weight)
        if "despacho" in key:
            data["customs_combined"] = _parse_euro(price) or str(price).strip()
        elif "seguro" in key:
            data["insurance"] = str(price).strip() if price else None
    return data


def _cost_blocks() -> list[CostBlock]:
    default_apply = "Apply if: Applies in all items"
    return [
        CostBlock(
            title="Transport cost (Distribución en la Isla de Tenerife)",
            apply_if=default_apply,
            rate_by="Rate by: Weight (Distribución Isla Tenerife)",
            columns=[
                CostColumnSpec(label, "Flat") for label in TENERIFE_DIST_BRACKETS
            ],
            uses_shared_currency=True,
        ),
        CostBlock(
            title="Transport cost (Envios)",
            apply_if=default_apply,
            rate_by="Rate by: Per kg (Distribución desde almacén Tenerife)",
            columns=[CostColumnSpec("Flat", "Flat")],
        ),
        CostBlock(
            title=(
                "Origin Customs Clearance (ALL) + Destination Customs Clearance "
                "(Grupajes (Despacho Destino))"
            ),
            apply_if=default_apply,
            rate_by="Rate by: Per shipment (Otros conceptos)",
            columns=[CostColumnSpec("Flat", "Flat")],
        ),
        CostBlock(
            title="Shipment Insurance (Cargas Completas (Seguro))",
            apply_if=default_apply,
            rate_by="Rate by: Percent of freight",
            columns=[CostColumnSpec("Flat", "Flat")],
        ),
        CostBlock(
            title="Transport cost + BAF (Cargas EN CORAL)",
            apply_if=default_apply,
            rate_by="Rate by: Container (Carga en Coral, Tarifa 2026)",
            columns=[CostColumnSpec(s, "Flat") for s in FCL_SIZE_BRACKETS],
            uses_shared_currency=True,
        ),
        CostBlock(
            title="Transport cost + BAF (Cargas EN CELRA)",
            apply_if=default_apply,
            rate_by="Rate by: Container (Carga en Celra, Tarifa 2026)",
            columns=[CostColumnSpec(s, "Flat") for s in FCL_SIZE_BRACKETS],
            uses_shared_currency=True,
        ),
        CostBlock(
            title="Transport cost + BAF (CARGA EN CELRA / CORAL)",
            apply_if=default_apply,
            rate_by="Rate by: Container (Celra + Coral, Tarifa 2026)",
            columns=[CostColumnSpec("40 pies", "Flat")],
            uses_shared_currency=True,
        ),
        CostBlock(
            title="Transport cost (Grupajes)",
            apply_if=default_apply,
            rate_by="Rate by: Per kg (Zona Destino / Grupajes, Tarifa 2025)",
            columns=[
                CostColumnSpec("Per kg", "Per kg"),
                CostColumnSpec("Minimum", "Flat"),
            ],
            uses_shared_currency=True,
        ),
        CostBlock(
            title="BAF (CARGA EN CELRA / CORAL, 40)",
            apply_if=default_apply,
            rate_by="Rate by: BAF 40ft (Celra + Coral)",
            columns=[CostColumnSpec("Flat", "Flat")],
        ),
        CostBlock(
            title="BAF (20FT/40FT)",
            apply_if=default_apply,
            rate_by="Rate by: BAF by origin (Celra + Coral)",
            columns=[
                CostColumnSpec("20 pies", "Flat"),
                CostColumnSpec("40 pies", "Flat"),
            ],
            uses_shared_currency=True,
        ),
        CostBlock(
            title="Emergency Bunker Surcharge",
            apply_if=default_apply,
            rate_by="Rate by: TBD",
            columns=[CostColumnSpec("Flat", "Flat")],
        ),
        CostBlock(
            title="BAF (40FT)",
            apply_if=default_apply,
            rate_by="Rate by: BAF 40ft (Celra + Coral)",
            columns=[CostColumnSpec("Flat", "Flat")],
        ),
        CostBlock(
            title=(
                "Transport cost + BAF (CARGA EN CELRA / CORAL) "
                "[by Origen]"
            ),
            apply_if=default_apply,
            rate_by="Rate by: Tarifa 2026 by origin (Celra + Coral)",
            columns=[CostColumnSpec("Flat", "Flat")],
        ),
        CostBlock(
            title="Shipment Insurance (Carga EN CELRA/CORAL'40)",
            apply_if=default_apply,
            rate_by="Rate by: Percent of freight",
            columns=[CostColumnSpec("Flat", "Flat")],
        ),
    ]


def cost_key(block: CostBlock, spec: CostColumnSpec) -> tuple[str, str]:
    return (block.title, spec.bracket_label)


def block_column_width(block: CostBlock) -> int:
    if block.uses_shared_currency:
        return 1 + len(block.columns)
    return 2 * len(block.columns)


def _has_cost(value: float | str | None) -> bool:
    return value is not None and value != ""


def _build_lane_rows() -> list[tuple[str, str]]:
    lanes: list[tuple[str, str]] = []
    for dest in GRUPAJE_DESTINATIONS:
        lanes.append((ORIGIN_BARCELONA, _zone_label(dest)))
    for dest in ISLAND_DESTINATIONS:
        lanes.append((ORIGIN_TENERIFE_WH, _zone_label(dest)))
    lanes.append((_zone_label("Tenerife"), _zone_label("Tenerife")))
    for origin in FCL_ORIGINS:
        zone = _zone_label(origin)
        lanes.append((zone, zone))
    return lanes


def build_matrix(
    fields: dict[str, Any],
) -> tuple[list[MatrixRow], list[CostBlock], list[SkippedItem]]:
    main = fields.get("MainCosts") or []
    main2 = fields.get("MainCosts2") or []
    skipped: list[SkippedItem] = []

    coral_rows = _section_rows(main, "Carga en Coral")
    celra_rows = _section_rows(main, "Carga en Celra")
    combo_rows = _section_rows(main, "Carga en Celra + Coral", "Celra + Coral")
    grupaje_rows = _section_rows(main, "Zona Destino/Grupajes", "Grupajes")

    tenerife_dist_rows = _section_rows(
        main2, "Disstribucion isla tenerife", "Distribucion isla tenerife"
    )
    wh_dist_rows = _section_rows(
        main2,
        "Distribution desde almacen de tenerife",
        "Distribución desde almacén de Tenerife",
    )
    otros_rows = _section_rows(main2, "Otros conceptos")

    coral = _parse_fcl_table(coral_rows)
    celra = _parse_fcl_table(celra_rows)
    combo = _parse_fcl_table(combo_rows)
    grupajes = _parse_grupajes(grupaje_rows)
    tenerife_dist = _parse_tenerife_distribution(tenerife_dist_rows)
    wh_dist = _parse_wh_distribution(wh_dist_rows)
    otros = _parse_otros_conceptos(otros_rows)

    customs = otros.get("customs_combined")
    insurance_cc = otros.get("insurance")

    cost_blocks = _cost_blocks()
    matrix_rows: list[MatrixRow] = []

    for origin_zone, dest_zone in _build_lane_rows():
        shipment = {
            "Origin Postal Code Zone": origin_zone,
            "Destination Postal Code Zone": dest_zone,
        }
        costs: dict[tuple[str, str], float | str | None] = {}
        text_costs: dict[tuple[str, str], str] = {}

        dest_name = dest_zone.replace("ES ", "", 1)
        origin_name = origin_zone.replace("ES ", "", 1).replace(" (Almacén)", "")

        # Tenerife island distribution
        if origin_zone == _zone_label("Tenerife") and dest_zone == _zone_label("Tenerife"):
            block = cost_blocks[0]
            for bracket in TENERIFE_DIST_BRACKETS:
                costs[cost_key(block, CostColumnSpec(bracket))] = tenerife_dist.get(
                    bracket
                )

        # Warehouse distribution (envíos)
        if origin_zone == ORIGIN_TENERIFE_WH:
            block = cost_blocks[1]
            for island in ISLAND_DESTINATIONS:
                if _norm_key(island) == _norm_key(dest_name):
                    costs[cost_key(block, block.columns[0])] = wh_dist.get(island)
                    break

        # Customs (flat on all rows when value present)
        block = cost_blocks[2]
        if isinstance(customs, (int, float)):
            costs[cost_key(block, block.columns[0])] = float(customs)
        elif customs:
            text_costs[cost_key(block, block.columns[0])] = str(customs)

        # Shipment insurance completas
        block = cost_blocks[3]
        if insurance_cc:
            text_costs[cost_key(block, block.columns[0])] = str(insurance_cc)

        # FCL Coral / Celra
        for block_index, table, title in (
            (4, coral, "Carga en Coral"),
            (5, celra, "Carga en Celra"),
        ):
            block = cost_blocks[block_index]
            if origin_zone == dest_zone and _norm_key(origin_name) in {
                _norm_key(o) for o in FCL_ORIGINS
            }:
                for size in FCL_SIZE_BRACKETS:
                    row = table.get((origin_name, size))
                    if row:
                        costs[cost_key(block, CostColumnSpec(size))] = row.get("cost3")

        # Celra + Coral combined (40ft tariff)
        block = cost_blocks[6]
        if origin_zone == dest_zone:
            row_40 = combo.get((origin_name, "40 pies")) or combo.get(
                ("Las Palmas", "40 pies")
            )
            if row_40 and _norm_key(origin_name) == _norm_key("Las Palmas"):
                costs[cost_key(block, block.columns[0])] = row_40.get("cost3")

        # Grupajes from Barcelona
        if origin_zone == ORIGIN_BARCELONA:
            block = cost_blocks[7]
            g = None
            for key, values in grupajes.items():
                if _norm_key(key) == _norm_key(dest_name):
                    g = values
                    break
            if g:
                costs[cost_key(block, block.columns[0])] = g.get("per_kg")
                costs[cost_key(block, block.columns[1])] = g.get("minimum")

        # BAF from Celra+Coral and Coral/Celra tables (BAFF = cost2)
        baf_by_origin: dict[str, dict[str, float | None]] = {}
        for table in (coral, celra, combo):
            for (orig, size), vals in table.items():
                baf_by_origin.setdefault(orig, {})[size] = vals.get("cost2")

        if origin_zone == dest_zone and origin_name in FCL_ORIGINS:
            combo_baf_40 = (combo.get((origin_name, "40 pies")) or {}).get("cost2")
            combo_baf_20 = (combo.get((origin_name, "20 pies")) or {}).get("cost2")
            table_baf = baf_by_origin.get(origin_name, {})

            block = cost_blocks[8]
            costs[cost_key(block, block.columns[0])] = combo_baf_40 or table_baf.get(
                "40 pies"
            )

            block = cost_blocks[9]
            costs[cost_key(block, block.columns[0])] = combo_baf_20 or table_baf.get(
                "20 pies"
            )
            costs[cost_key(block, block.columns[1])] = combo_baf_40 or table_baf.get(
                "40 pies"
            )

            block = cost_blocks[11]
            costs[cost_key(block, block.columns[0])] = combo_baf_40 or table_baf.get(
                "40 pies"
            )

            block = cost_blocks[12]
            combo_row = combo.get((origin_name, "40 pies"))
            if combo_row:
                costs[cost_key(block, block.columns[0])] = combo_row.get("cost3")

            block = cost_blocks[13]
            if combo.get((origin_name, "40 pies")) or (
                _norm_key(origin_name) == _norm_key("Las Palmas")
                and combo.get(("Las Palmas", "40 pies"))
            ):
                text_costs[cost_key(block, block.columns[0])] = INSURANCE_FCL_40_LABEL

        matrix_rows.append(MatrixRow(shipment=shipment, costs=costs, text_costs=text_costs))

    for section_name, _ in _iter_sections(main):
        if _norm_key(section_name) in {
            _norm_key("RGA EN CORAL"),
            _norm_key("CARGA EN CERLA"),
            _norm_key("Transporte desde barcelona"),
        }:
            skipped.append(
                SkippedItem(
                    source="MainCosts",
                    label=section_name,
                    detail="Section not mapped to matrix columns in this build.",
                )
            )

    return matrix_rows, cost_blocks, skipped


def write_rates_sheet(
    worksheet,
    matrix_rows: list[MatrixRow],
    cost_blocks: list[CostBlock],
) -> None:
    shipment_count = len(SHIPMENT_HEADERS)

    def write_merged_row(row_index: int, values: list[str]) -> None:
        column_index = shipment_count + 1
        for block_index, block in enumerate(cost_blocks):
            width = block_column_width(block)
            value = values[block_index] if block_index < len(values) else ""
            cell = worksheet.cell(row=row_index, column=column_index, value=value)
            cell.font = BOLD
            cell.fill = HEADER_FILL
            cell.alignment = LEFT
            if width > 1:
                worksheet.merge_cells(
                    start_row=row_index,
                    start_column=column_index,
                    end_row=row_index,
                    end_column=column_index + width - 1,
                )
            column_index += width

    write_merged_row(COST_NAME_ROW, [block.title for block in cost_blocks])
    write_merged_row(APPLY_IF_ROW, [block.apply_if for block in cost_blocks])
    write_merged_row(RATE_BY_ROW, [block.rate_by for block in cost_blocks])

    for col_idx, header in enumerate(SHIPMENT_HEADERS, start=1):
        cell = worksheet.cell(row=COLUMN_HEADER_ROW, column=col_idx, value=header)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.alignment = LEFT

    column_index = shipment_count + 1
    for block in cost_blocks:
        if block.uses_shared_currency:
            currency_cell = worksheet.cell(row=BRACKET_ROW, column=column_index)
            currency_cell.fill = HEADER_FILL
            worksheet.cell(row=COLUMN_HEADER_ROW, column=column_index, value=CURRENCY)
            for offset, spec in enumerate(block.columns):
                spec_col = column_index + 1 + offset
                bracket_cell = worksheet.cell(
                    row=BRACKET_ROW, column=spec_col, value=spec.bracket_label
                )
                bracket_cell.font = BOLD
                bracket_cell.fill = HEADER_FILL
                bracket_cell.alignment = CENTER
                unit_cell = worksheet.cell(
                    row=COLUMN_HEADER_ROW, column=spec_col, value=spec.rate_unit
                )
                unit_cell.font = BOLD
                unit_cell.fill = HEADER_FILL
                unit_cell.alignment = CENTER
            column_index += block_column_width(block)
            continue

        for spec in block.columns:
            bracket_cell = worksheet.cell(
                row=BRACKET_ROW, column=column_index, value=spec.bracket_label
            )
            bracket_cell.font = BOLD
            bracket_cell.fill = HEADER_FILL
            bracket_cell.alignment = CENTER
            worksheet.merge_cells(
                start_row=BRACKET_ROW,
                start_column=column_index,
                end_row=BRACKET_ROW,
                end_column=column_index + 1,
            )
            worksheet.cell(row=COLUMN_HEADER_ROW, column=column_index, value=CURRENCY)
            unit_cell = worksheet.cell(
                row=COLUMN_HEADER_ROW, column=column_index + 1, value=spec.rate_unit
            )
            unit_cell.font = BOLD
            unit_cell.fill = HEADER_FILL
            unit_cell.alignment = CENTER
            column_index += 2

    for row_offset, matrix_row in enumerate(matrix_rows):
        excel_row = DATA_START_ROW + row_offset
        for col_idx, header in enumerate(SHIPMENT_HEADERS, start=1):
            worksheet.cell(
                row=excel_row, column=col_idx, value=matrix_row.shipment.get(header)
            )

        column_index = shipment_count + 1
        for block in cost_blocks:
            if block.uses_shared_currency:
                worksheet.cell(row=excel_row, column=column_index, value=CURRENCY)
                for offset, spec in enumerate(block.columns):
                    spec_col = column_index + 1 + offset
                    key = cost_key(block, spec)
                    if key in matrix_row.text_costs:
                        worksheet.cell(
                            row=excel_row,
                            column=spec_col,
                            value=matrix_row.text_costs[key],
                        )
                        continue
                    value = matrix_row.costs.get(key)
                    if _has_cost(value) and isinstance(value, (int, float)):
                        cell = worksheet.cell(row=excel_row, column=spec_col, value=value)
                        cell.number_format = "0.00"
                column_index += block_column_width(block)
                continue

            for spec in block.columns:
                key = cost_key(block, spec)
                if key in matrix_row.text_costs:
                    worksheet.cell(
                        row=excel_row,
                        column=column_index,
                        value=matrix_row.text_costs[key],
                    )
                    column_index += 2
                    continue
                value = matrix_row.costs.get(key)
                if _has_cost(value) and isinstance(value, (int, float)):
                    worksheet.cell(row=excel_row, column=column_index, value=CURRENCY)
                    cell = worksheet.cell(
                        row=excel_row, column=column_index + 1, value=value
                    )
                    cell.number_format = "0.00"
                column_index += 2

    for col_idx in range(1, worksheet.max_column + 1):
        worksheet.column_dimensions[get_column_letter(col_idx)].width = 20


def write_skipped_sheet(worksheet, items: list[SkippedItem]) -> None:
    headers = ["Source", "Label", "Detail"]
    for col_idx, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=1, column=col_idx, value=header)
        cell.font = BOLD
        cell.fill = HEADER_FILL
    for row_idx, item in enumerate(items, start=2):
        worksheet.cell(row=row_idx, column=1, value=item.source)
        worksheet.cell(row=row_idx, column=2, value=item.label)
        worksheet.cell(row=row_idx, column=3, value=item.detail)
    for col_idx in range(1, 4):
        worksheet.column_dimensions[get_column_letter(col_idx)].width = 36


def build_matrix_workbook(cleaned_path: Path, output_path: Path | None = None) -> Path:
    fields = _load_fields(cleaned_path)
    matrix_rows, cost_blocks, skipped = build_matrix(fields)
    if not matrix_rows:
        raise ValueError("No matrix rows produced.")

    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = cleaned_path.stem.replace(".cleaned", "")
        output_path = OUTPUT_DIR / f"matrix_{stem}_{timestamp}.xlsx"

    workbook = Workbook()
    rates_sheet = workbook.active
    rates_sheet.title = "Rates"
    write_rates_sheet(rates_sheet, matrix_rows, cost_blocks)

    skipped_sheet = workbook.create_sheet("did not added")
    write_skipped_sheet(skipped_sheet, skipped)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Almazar rate matrix XLSX.")
    parser.add_argument(
        "cleaned",
        nargs="?",
        type=Path,
        help="Cleaned JSON path (default: newest *.cleaned.json in processing/)",
    )
    parser.add_argument("-o", "--output", type=Path, help="Output XLSX path")
    args = parser.parse_args()

    if args.cleaned:
        cleaned_path = args.cleaned
    else:
        candidates = sorted(
            PROCESSING_DIR.glob("*.cleaned.json"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not candidates:
            raise FileNotFoundError(f"No cleaned JSON in {PROCESSING_DIR}")
        cleaned_path = candidates[0]

    output_path = build_matrix_workbook(cleaned_path, output_path=args.output)
    print(f"Matrix written to {output_path}")


if __name__ == "__main__":
    main()
