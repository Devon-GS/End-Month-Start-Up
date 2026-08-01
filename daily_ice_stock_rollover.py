from __future__ import annotations

import argparse
import re
from pathlib import Path

import openpyxl


MONTH_NAMES = [
    "JANUARY",
    "FEBRUARY",
    "MARCH",
    "APRIL",
    "MAY",
    "JUNE",
    "JULY",
    "AUGUST",
    "SEPTEMBER",
    "OCTOBER",
    "NOVEMBER",
    "DECEMBER",
]


def month_name(month: int) -> str:
    return MONTH_NAMES[month - 1]


def next_month(month: int, year: int) -> tuple[int, int]:
    if month == 12:
        return 1, year + 1
    return month + 1, year


def parse_tax_year_folder(folder: Path) -> tuple[int, int]:
    match = re.match(r"^ICE STOCK FROM MARCH (\d{4}) TO FEB(?:RUARY)? (\d{4})$", folder.name.upper())
    if not match:
        raise ValueError(f"Unrecognized tax year folder: {folder.name}")
    start_year = int(match.group(1))
    end_year = int(match.group(2))
    if end_year != start_year + 1:
        raise ValueError(f"Unexpected year range in folder: {folder.name}")
    return start_year, end_year


def find_latest_tax_year_folder(root: Path) -> Path:
    candidates: list[tuple[int, Path]] = []
    for folder in root.iterdir():
        if not folder.is_dir():
            continue
        try:
            start_year, _ = parse_tax_year_folder(folder)
        except ValueError:
            continue
        candidates.append((start_year, folder))
    if not candidates:
        raise FileNotFoundError(f"No ICE STOCK tax year folders found in {root}")
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def parse_sheet_title(title: str) -> tuple[int, int]:
    match = re.match(r"^([A-Z]+)\s+(\d{2})$", title.strip(), re.IGNORECASE)
    if not match:
        raise ValueError(f"Unrecognized sheet title: {title}")
    month_text = match.group(1).upper()
    year_suffix = int(match.group(2))
    if month_text not in MONTH_NAMES:
        raise ValueError(f"Unrecognized month in sheet title: {title}")
    return MONTH_NAMES.index(month_text) + 1, 2000 + year_suffix


def find_latest_sheet(wb: openpyxl.Workbook):
    for title in reversed(wb.sheetnames):
        if re.match(r"^[A-Z]+\s+\d{2}$", title.strip(), re.IGNORECASE):
            return wb[title]
    raise FileNotFoundError("No monthly ICE STOCK sheets found in workbook")


def build_new_sheet_title(month: int, year: int) -> str:
    return f"{month_name(month)} {str(year)[-2:]}"


def clear_range(ws, cell_range: str) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.value = None


def negative_value(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return -abs(value)
    if value == "0":
        return 0
    try:
        numeric = float(value)
        return -abs(numeric)
    except (TypeError, ValueError):
        return value


def process_ice_stock(root: Path, dry_run: bool = False) -> tuple[Path, str]:
    stock_root = root / "DAILY PROCEDURE DOCUMENTS" / "DAILY ICE STOCK"
    tax_year_folder = find_latest_tax_year_folder(stock_root)
    workbook_path = tax_year_folder / "1 ICE RECON.xlsx"
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = openpyxl.load_workbook(workbook_path)
    source_ws = find_latest_sheet(wb)
    source_month, source_year = parse_sheet_title(source_ws.title)
    next_month_number, next_year = next_month(source_month, source_year)
    new_sheet_title = build_new_sheet_title(next_month_number, next_year)

    if dry_run:
        return workbook_path, new_sheet_title

    data_wb = openpyxl.load_workbook(workbook_path, data_only=True)
    data_ws = data_wb[source_ws.title]

    new_ws = wb.copy_worksheet(source_ws)
    new_ws.title = new_sheet_title

    source_month_name = month_name(source_month)
    new_month_name = month_name(next_month_number)

    a3_value = source_ws["A3"].value
    if isinstance(a3_value, str):
        new_ws["A3"] = a3_value.replace(source_month_name, new_month_name)
    else:
        new_ws["A3"] = new_month_name

    b37_value = data_ws["B37"].value
    if b37_value is None or b37_value == 0 or b37_value == "0":
        b37_value = data_ws["B36"].value
    new_ws["F5"] = negative_value(b37_value)

    clear_range(new_ws, "C6:E36")
    clear_range(new_ws, "H6:H36")

    wb.save(workbook_path)
    return workbook_path, new_sheet_title


def main() -> int:
    parser = argparse.ArgumentParser(description="Append the next DAILY ICE STOCK sheet to 1 ICE RECON.xlsx.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Project root containing 'DAILY PROCEDURE DOCUMENTS'. Defaults to the current directory.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be created without writing files.")
    args = parser.parse_args()

    workbook_path, sheet_title = process_ice_stock(args.root, dry_run=args.dry_run)
    if args.dry_run:
        print(f"Would update: {workbook_path}")
        print(f"Would append sheet: {sheet_title}")
    else:
        print(f"Updated: {workbook_path}")
        print(f"Added sheet: {sheet_title}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
