"""Select Azure Document Intelligence JSON files from input/."""

from __future__ import annotations

from pathlib import Path

from project_paths import INPUT_DIR, ensure_workspace_dirs


def list_input_json_files() -> list[Path]:
    ensure_workspace_dirs()
    return sorted(
        path
        for path in INPUT_DIR.glob("*.json")
        if path.is_file() and not path.name.startswith("~$")
    )


def resolve_input_json(path: Path | None = None) -> Path:
    if path is not None:
        resolved = path.expanduser().resolve()
        if not resolved.is_file():
            raise FileNotFoundError(f"Input file not found: {resolved}")
        return resolved

    files = list_input_json_files()
    if not files:
        raise FileNotFoundError(
            f"No JSON files found in {INPUT_DIR}. "
            "Place Azure Document Intelligence exports there first."
        )
    if len(files) == 1:
        return files[0]
    return prompt_input_json()


def prompt_input_json() -> Path:
    files = list_input_json_files()
    if not files:
        raise FileNotFoundError(
            f"No JSON files found in {INPUT_DIR}. "
            "Place Azure Document Intelligence exports there first."
        )

    print(f"\nSelect input JSON file from {INPUT_DIR}:")
    for index, path in enumerate(files, start=1):
        default_marker = " (default)" if index == 1 else ""
        print(f"  {index}. {path.name}{default_marker}")
    print("\nEnter file number, or press Enter for 1:")

    while True:
        raw = input("> ").strip()
        if not raw:
            return files[0]
        if raw.isdigit():
            choice = int(raw) - 1
            if 0 <= choice < len(files):
                return files[choice]
        print("Invalid selection. Try again.")
