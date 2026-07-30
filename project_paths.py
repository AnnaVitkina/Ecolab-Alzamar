"""Project folder paths for local runs and Google Colab."""

from __future__ import annotations

import os
from pathlib import Path

COLAB_CODE_DIR = Path("/content/Ecolab-Alzamar")

COLAB_INPUT_DIR = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Ecolab/RMT_ALZAMAR/input"
)
COLAB_PROCESSING_DIR = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Ecolab/RMT_ALZAMAR/processing"
)
COLAB_OUTPUT_DIR = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_Ecolab/RMT_ALZAMAR/output"
)


def resolve_code_dir() -> Path:
    """Folder containing pipeline scripts (works when imported or run via exec())."""
    try:
        return Path(__file__).resolve().parent
    except NameError:
        env = os.environ.get("ALZAMAR_CODE_DIR")
        if env:
            return Path(env).expanduser().resolve()
        if COLAB_CODE_DIR.is_dir():
            return COLAB_CODE_DIR.resolve()
        return Path.cwd().resolve()


_LOCAL_ROOT = resolve_code_dir()


def is_colab_environment() -> bool:
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return COLAB_CODE_DIR.is_dir()


def _apply_colab_paths() -> None:
    global CODE_DIR, INPUT_DIR, PROCESSING_DIR, OUTPUT_DIR
    CODE_DIR = COLAB_CODE_DIR
    INPUT_DIR = COLAB_INPUT_DIR
    PROCESSING_DIR = COLAB_PROCESSING_DIR
    OUTPUT_DIR = COLAB_OUTPUT_DIR


def _apply_local_paths() -> None:
    global CODE_DIR, INPUT_DIR, PROCESSING_DIR, OUTPUT_DIR
    CODE_DIR = _LOCAL_ROOT
    INPUT_DIR = _LOCAL_ROOT / "input"
    PROCESSING_DIR = _LOCAL_ROOT / "processing"
    OUTPUT_DIR = _LOCAL_ROOT / "output"


if is_colab_environment():
    _apply_colab_paths()
else:
    _apply_local_paths()


def ensure_workspace_dirs() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
