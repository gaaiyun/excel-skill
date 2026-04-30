#!/usr/bin/env python3
"""
recalc.py -- Recalculate formulas in openpyxl-generated .xlsx files.

> TODO: Implement LibreOffice-based formula recalculation.

openpyxl writes formulas as strings; the .xlsx file has no cached values.
This script is intended to open the file in LibreOffice headless mode,
trigger a full recalculation, and save the result so that cached values
are populated.

Planned usage:
    python scripts/recalc.py input.xlsx
    python scripts/recalc.py input.xlsx --output recalculated.xlsx

Implementation plan:
    1. Locate LibreOffice binary (soffice / libreoffice)
    2. Use --headless --calc --convert-to xlsx to open, recalc, save
    3. Alternatively: use a LibreOffice Basic macro to force Ctrl+Shift+F9
    4. Fall back to openpyxl data_only=True warning if LO is not installed
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def find_libreoffice() -> str | None:
    """Find the LibreOffice executable on the system.

    TODO: Implement cross-platform detection:
      - Linux/macOS: 'soffice' or 'libreoffice' in PATH
      - Windows: check 'C:\\Program Files\\LibreOffice\\program\\soffice.exe'
    """
    # Try PATH first
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path

    # TODO: Add Windows-specific paths
    # TODO: Add macOS-specific paths (/Applications/LibreOffice.app/...)

    return None


def recalc_with_libreoffice(input_path: Path, output_path: Path | None) -> int:
    """Open the .xlsx in LibreOffice headless, recalc, and save.

    TODO: Implement the actual recalculation logic:
      - Use subprocess to call soffice --headless --calc
      - Option A: convert-to xlsx (simple but loses some formatting)
      - Option B: use a .py macro that calls document.calculateAll()
      - Option C: use a Basic macro with SendKeys Ctrl+Shift+F9
    """
    lo_path = find_libreoffice()
    if lo_path is None:
        print(
            "[error] LibreOffice not found. Install it to enable formula recalculation.\n"
            "        Download: https://www.libreoffice.org/download/\n"
            "        Alternative: open the file in Excel (auto-recalc on open).",
            file=sys.stderr,
        )
        return 1

    # TODO: Implement subprocess call
    # Example (needs testing):
    # subprocess.run([
    #     lo_path,
    #     "--headless",
    #     "--calc",
    #     "--convert-to", "xlsx",
    #     "--outdir", str(output_path.parent),
    #     str(input_path),
    # ])

    print(f"[stub] Would recalculate: {input_path}")
    print(f"[stub] LibreOffice found at: {lo_path}")
    print("[stub] Actual recalculation logic not yet implemented.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Recalculate formulas in openpyxl-generated .xlsx files"
    )
    parser.add_argument("input", type=Path, help="Input .xlsx file")
    parser.add_argument("--output", "-o", type=Path, default=None, help="Output path (default: overwrite input)")
    args = parser.parse_args()

    if not args.input.exists():
        print(f"[error] File not found: {args.input}", file=sys.stderr)
        return 1

    output = args.output or args.input
    return recalc_with_libreoffice(args.input, output)


if __name__ == "__main__":
    sys.exit(main())
