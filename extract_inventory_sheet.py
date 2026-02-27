from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook, load_workbook


def extract_sheet_values_to_text(src: Path, dst: Path, sheet_name: str = "재고장") -> None:
    wb = load_workbook(src, data_only=True, read_only=True)
    if sheet_name not in wb.sheetnames:
        names = ", ".join(wb.sheetnames)
        raise ValueError(f"'{sheet_name}' 시트를 찾지 못했습니다. 사용 가능한 시트: {names}")

    ws = wb[sheet_name]
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = sheet_name

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        for col_idx, value in enumerate(row, start=1):
            if value is None:
                text = ""
            else:
                text = str(value).strip()
            out_ws.cell(row=row_idx, column=col_idx, value=text)

    out_wb.save(dst)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="원본 엑셀에서 재고장 시트 값만 추출해 경량 엑셀 파일을 생성합니다."
    )
    parser.add_argument("source", type=Path, help="원본 엑셀 경로")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("재고장_더미.xlsx"),
        help="출력 엑셀 경로 (기본값: 재고장_더미.xlsx)",
    )
    parser.add_argument(
        "-s",
        "--sheet",
        default="재고장",
        help="추출할 시트 이름 (기본값: 재고장)",
    )
    args = parser.parse_args()

    extract_sheet_values_to_text(args.source, args.output, args.sheet)
    print(f"완료: {args.output}")


if __name__ == "__main__":
    main()
