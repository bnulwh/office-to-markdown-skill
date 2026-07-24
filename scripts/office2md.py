#!/usr/bin/env python3
"""Unified Office-to-Markdown converter.

Auto-detects file type by extension and dispatches to the appropriate
converter module (pptx2md, docx2md, xlsx2md).

Usage:
  python office2md.py input.pptx [-o output.md]
  python office2md.py input.docx [-o output.md]
  python office2md.py input.xlsx [-o output.md]

Supported formats: .pptx, .docx, .xlsx, .xlsm, .pptm, .docm
"""

import sys
import argparse
from pathlib import Path

# Mapping of file extensions to converter modules
_EXT_MAP = {
    ".pptx": "pptx2md",
    ".pptm": "pptx2md",
    ".docx": "docx2md",
    ".docm": "docx2md",
    ".xlsx": "xlsx2md",
    ".xlsm": "xlsx2md",
}


def _import_converter(module_name: str):
    """Dynamically import a converter module from the same directory."""
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        module_name,
        str(Path(__file__).parent / f"{module_name}.py"),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def convert(input_path: str, output_path: str | None = None) -> str:
    """Convert an Office file to Markdown. Auto-detects format by extension."""
    inp = Path(input_path)
    ext = inp.suffix.lower()

    if ext not in _EXT_MAP:
        supported = ", ".join(sorted(_EXT_MAP.keys()))
        raise ValueError(f"Unsupported file type '{ext}'. Supported: {supported}")

    module_name = _EXT_MAP[ext]
    mod = _import_converter(module_name)
    return mod.convert(str(inp), output_path)


def main():
    ap = argparse.ArgumentParser(
        description="Convert Office files (.pptx/.docx/.xlsx) to Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument("input", help="Input Office file path (.pptx, .docx, .xlsx)")
    ap.add_argument("-o", "--output", help="Output .md file (default: <input>.md)")
    args = ap.parse_args()

    inp = Path(args.input)
    if not inp.exists():
        print(f"Error: file not found: {inp}", file=sys.stderr)
        sys.exit(1)

    ext = inp.suffix.lower()
    if ext not in _EXT_MAP:
        supported = ", ".join(sorted(_EXT_MAP.keys()))
        print(f"Error: unsupported file type '{ext}'. Supported: {supported}",
              file=sys.stderr)
        sys.exit(1)

    out = args.output or str(inp.with_suffix(".md"))

    try:
        convert(str(inp), out)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
