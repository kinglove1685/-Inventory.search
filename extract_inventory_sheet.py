from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook, load_workbook


def extract_sheet_values(src: Path, dst: Path, sheet_name: str = "재고장") -> None:
    wb = load_workbook(src, data_only=True, read_only=True)
    if sheet_name not in wb.sheetnames:
        available = ", ".join(wb.sheetnames)
        raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {available}")

    ws = wb[sheet_name]
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = sheet_name

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        for col_idx, value in enumerate(row, start=1):
            out_ws.cell(row=row_idx, column=col_idx, value="" if value is None else str(value).strip())

    out_wb.save(dst)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create a lightweight workbook by copying values from one sheet."
    )
    parser.add_argument("source", type=Path, help="Source Excel file path")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("inventory_dummy.xlsx"),
        help="Output Excel file path (default: inventory_dummy.xlsx)",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        default="재고장",
        help="Sheet name to extract (default: 재고장)",
    )
    args = parser.parse_args()

    extract_sheet_values(args.source, args.output, args.sheet)
    print(f"Done: {args.output}")


if __name__ == "__main__":
    main()
