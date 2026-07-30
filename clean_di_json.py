"""Strip Azure Document Intelligence analyze JSON down to document fields only."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from file_selection import resolve_input_json
from project_paths import PROCESSING_DIR

_VALUE_KEYS = (
    "valueString",
    "valueNumber",
    "valueInteger",
    "valueDate",
    "valueTime",
    "valuePhoneNumber",
    "valueCountryRegion",
    "valueSelectionMark",
    "valueSignature",
    "valueBoolean",
)


def _scalar_from_field(field: dict[str, Any]) -> Any:
    for key in _VALUE_KEYS:
        if key in field:
            return field[key]
    if "valueCurrency" in field:
        currency = field["valueCurrency"]
        if isinstance(currency, dict):
            amount = currency.get("amount")
            code = currency.get("currencyCode")
            if amount is not None and code:
                return f"{amount} {code}"
            return amount
        return currency
    if "valueAddress" in field:
        address = field["valueAddress"]
        if isinstance(address, dict):
            parts = [
                address.get("streetAddress"),
                address.get("city"),
                address.get("state"),
                address.get("postalCode"),
                address.get("countryRegion"),
            ]
            return ", ".join(part for part in parts if part)
        return address
    return None


def clean_di_field(field: object) -> Any:
    """Return the business value for one DI field node (no spans, polygons, etc.)."""
    if not isinstance(field, dict):
        return field

    field_type = field.get("type")
    if field_type == "array":
        items = field.get("valueArray")
        if not isinstance(items, list):
            return []
        return [clean_di_field(item) for item in items]

    if field_type == "object":
        obj = field.get("valueObject")
        if not isinstance(obj, dict):
            return {}
        return {name: clean_di_field(child) for name, child in obj.items()}

    if field_type == "string":
        return field.get("valueString")

    scalar = _scalar_from_field(field)
    if scalar is not None:
        return scalar
    if field_type is not None:
        return None
    return field


def clean_document_fields(document: dict[str, Any]) -> dict[str, Any]:
    fields = document.get("fields")
    if not isinstance(fields, dict):
        return {}
    return {name: clean_di_field(field) for name, field in fields.items()}


def clean_analyze_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """
    Keep job metadata and per-document field values only.

    Drops analyzeResult.content, pages, tables, paragraphs, and per-field
    boundingRegions / spans / confidence / content.
    """
    analyze = payload.get("analyzeResult")
    if not isinstance(analyze, dict):
        raise ValueError("JSON is missing 'analyzeResult'.")

    documents = analyze.get("documents")
    if not isinstance(documents, list):
        raise ValueError("JSON analyzeResult has no documents.")

    cleaned_documents: list[dict[str, Any]] = []
    for document in documents:
        if not isinstance(document, dict):
            continue
        entry: dict[str, Any] = {"fields": clean_document_fields(document)}
        doc_type = document.get("docType")
        if doc_type:
            entry["docType"] = doc_type
        confidence = document.get("confidence")
        if confidence is not None:
            entry["confidence"] = confidence
        cleaned_documents.append(entry)

    result: dict[str, Any] = {
        "status": payload.get("status"),
        "modelId": analyze.get("modelId"),
        "apiVersion": analyze.get("apiVersion"),
        "documents": cleaned_documents,
    }
    for key in ("createdDateTime", "lastUpdatedDateTime"):
        if key in payload:
            result[key] = payload[key]
    return result


def load_cleaned_fields(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as handle:
        payload = json.load(handle)
    if not isinstance(payload, dict):
        raise ValueError("Expected JSON root object.")
    return clean_analyze_payload(payload)


def write_cleaned_json(source: Path, destination: Path) -> dict[str, Any]:
    cleaned = load_cleaned_fields(source)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("w", encoding="utf-8") as handle:
        json.dump(cleaned, handle, indent=2, ensure_ascii=False)
        handle.write("\n")
    return cleaned


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Clean Azure DI JSON: keep document fields, drop layout OCR payload."
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        help="Path to raw DI JSON (default: sole file in input/)",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output path (default: processing/<stem>.cleaned.json)",
    )
    args, _unknown = parser.parse_known_args()

    source = resolve_input_json(args.input)
    dest = args.output or (PROCESSING_DIR / f"{source.stem}.cleaned.json")
    cleaned = write_cleaned_json(source, dest)

    print(f"Wrote {dest}")
    for index, doc in enumerate(cleaned.get("documents", [])):
        fields = doc.get("fields", {})
        prefix = f"  document[{index}]" if len(cleaned.get("documents", [])) > 1 else " "
        for name, value in fields.items():
            if isinstance(value, list):
                print(f"{prefix} {name}: {len(value)} row(s)")
            else:
                print(f"{prefix} {name}: {value!r}")


if __name__ == "__main__":
    main()
