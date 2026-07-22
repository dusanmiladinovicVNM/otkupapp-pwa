"""Legacy compatibility entry point.

Use `02_enrich_financials.py` directly for new automation.
"""

from __future__ import annotations

import runpy
from pathlib import Path


if __name__ == "__main__":
    runpy.run_path(
        str(Path(__file__).resolve().parent / "02_enrich_financials.py"),
        run_name="__main__",
    )
