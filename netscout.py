"""NETSCOUT Automation v0.89 command entry point.

Run from this folder:

    python3 -m netscout --help
"""

from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
for path in (ROOT, ROOT / "src", ROOT / "vendor"):
    text = str(path)
    if text not in sys.path:
        sys.path.insert(0, text)

from V8.v8_cli import main


if __name__ == "__main__":
    raise SystemExit(main())
