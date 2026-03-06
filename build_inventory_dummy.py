from __future__ import annotations

from pathlib import Path

from extract_inventory_sheet import extract_sheet_values


OUTPUT_NAME = "inventory_dummy.xlsx"
PREFERRED_SOURCE = "재고관련 프로그램제작.xlsx"
TARGET_SHEET = "재고장"


def find_source(base: Path) -> Path:
    preferred = base / PREFERRED_SOURCE
    if preferred.exists():
        return preferred

    candidates = [
        p
        for p in base.glob("*.xlsx")
        if p.name != OUTPUT_NAME and not p.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError("No source .xlsx file found in current folder.")
    return sorted(candidates, key=lambda p: p.name)[0]


def main() -> None:
    base = Path(__file__).resolve().parent
    src = find_source(base)
    out = base / OUTPUT_NAME
    extract_sheet_values(src, out, TARGET_SHEET)
    print(f"Source: {src}")
    print(f"Output: {out}")
    print("Done")


if __name__ == "__main__":
    main()
