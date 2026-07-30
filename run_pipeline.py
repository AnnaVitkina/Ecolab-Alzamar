#!/usr/bin/env python3
"""CLI / Colab entry point for the Almazar end-to-end pipeline."""

from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path
from types import ModuleType


def _code_dir_candidates() -> list[Path]:
    out: list[Path] = []
    try:
        out.append(Path(__file__).resolve().parent)
    except NameError:
        pass
    env = os.environ.get("ALZAMAR_CODE_DIR")
    if env:
        out.append(Path(env).expanduser())
    out.append(Path("/content/Ecolab-Alzamar"))
    out.append(Path.cwd())
    return out


def _resolve_code_dir() -> Path:
    for root in _code_dir_candidates():
        root = root.resolve()
        if (root / "pipeline.py").is_file():
            return root
    raise RuntimeError(
        "Could not find Almazar code directory (need pipeline.py). "
        "Set ALZAMAR_CODE_DIR or run from /content/Ecolab-Alzamar."
    )


def _ensure_code_on_path(code_dir: Path) -> None:
    code_dir_str = str(code_dir)
    if code_dir_str not in sys.path:
        sys.path.insert(0, code_dir_str)


def _load_pipeline_module(code_dir: Path) -> ModuleType:
    """Load pipeline.py from disk (avoids stale/wrong ``pipeline`` in sys.modules)."""
    pipeline_path = code_dir / "pipeline.py"
    module_name = "alzamar_pipeline"
    spec = importlib.util.spec_from_file_location(module_name, pipeline_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load pipeline module from {pipeline_path}")

    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def _in_notebook() -> bool:
    try:
        from IPython import get_ipython

        shell = get_ipython()
        return shell is not None and shell.__class__.__name__ in (
            "ZMQInteractiveShell",
            "Shell",
        )
    except ImportError:
        return False


def _should_raise_system_exit() -> bool:
    if _in_notebook():
        return False
    try:
        from project_paths import is_colab_environment

        return not is_colab_environment()
    except ImportError:
        return True


def _run() -> int:
    code_dir = _resolve_code_dir()
    _ensure_code_on_path(code_dir)

    from project_paths import is_colab_environment

    pipeline = _load_pipeline_module(code_dir)

    if is_colab_environment() or _in_notebook():
        result = pipeline.run_pipeline()
        pipeline.print_summary(result)
        return 0

    return int(pipeline.main())


if __name__ == "__main__":
    exit_code = _run()
    if _should_raise_system_exit():
        raise SystemExit(exit_code)
