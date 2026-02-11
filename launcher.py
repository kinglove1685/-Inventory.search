from __future__ import annotations

import os
import sys
from pathlib import Path

from streamlit.web import cli as stcli


def _base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent


def main() -> None:
    base = _base_dir()
    script_path = base / "app.py"
    if not script_path.exists():
        raise FileNotFoundError(f"app.py not found: {script_path}")

    # Keep behavior consistent with `streamlit run app.py`.
    os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "false")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    sys.argv = [
        "streamlit",
        "run",
        str(script_path),
        "--global.developmentMode",
        "false",
        "--server.port",
        "3000",
        "--server.headless",
        "false",
    ]
    stcli.main()


if __name__ == "__main__":
    main()
