from __future__ import annotations

import argparse
import calendar
import re
import shutil
from datetime import date
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


WEEKDAY_NAMES = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]


def month_name(month: int) -> str:
    return MONTH_NAMES[month - 1]


def next_month(month: int, year: int) -> tuple[int, int]:
    if month == 12:
        return 1, year + 1
    return month + 1, year


def parse_workbook_name(path: Path) -> tuple[int, int, int]:
    match = re.match(r"^(\d{2})\s+DAILY SALES\s+([A-Z]+)\s+(\d{4})\.XLSX$", path.name.upper())
    if not match:
        raise ValueError(f"Unrecognized workbook name: {path.name}")

    prefix = int(match.group(1))
    month_text = match.group(2)
    year = int(match.group(3))

    try:
        month = MONTH_NAMES.index(month_text) + 1
    except ValueError as exc:
        raise ValueError(f"Unrecognized month name in workbook: {path.name}") from exc

    return prefix, month, year


def find_latest_tax_year_folder(root: Path) -> Path:
    candidates = []
    for folder in root.iterdir():
        if not folder.is_dir():
            continue
        match = re.match(r"^DAILY SALES TAX YR MARCH (\d{4}) TO FEB (\d{4})$", folder.name.upper())
        if not match:
            continue
        start_year = int(match.group(1))
        end_year = int(match.group(2))
        if end_year != start_year + 1:
            continue
        candidates.append((start_year, folder))
    if not candidates:
        raise FileNotFoundError(f"No DAILY SALES tax year folders found in {root}")
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def build_weekday_label(day: date) -> str:
    return WEEKDAY_NAMES[day.weekday()]


def clear_range(ws, cell_range: str) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.value = None


def set_date_range(ws, start_row: int, month: int, year: int) -> None:
    days_in_month = calendar.monthrange(year, month)[1]
    for offset in range(days_in_month):
        row = start_row + offset
        current = date(year, month, offset + 1)
        ws[f"A{row}"] = current
        ws[f"B{row}"] = build_weekday_label(current)


def update_reference_dates(ws, month: int, year: int) -> None:
    ref_date = date(year, month, 1)
    ws["B40"] = ref_date
    ws["B41"] = ref_date


def process_daily_sales(root: Path, dry_run: bool = False) -> tuple[Path, Path]:
    sales_root = root / "DAILY PROCEDURE DOCUMENTS" / "DAILY SALES"
    tax_year_folder = find_latest_tax_year_folder(sales_root)
    source_workbook = []
    for path in tax_year_folder.iterdir():
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$"):
            try:
                prefix, month, year = parse_workbook_name(path)
            except ValueError:
                continue
            source_workbook.append((prefix, month, year, path))
    if not source_workbook:
        raise FileNotFoundError(f"No DAILY SALES workbook found in {tax_year_folder}")

    source_prefix, source_month, source_year, source_path = sorted(source_workbook, key=lambda item: item[0])[-1]
    next_month_number, next_year = next_month(source_month, source_year)
    next_prefix = source_prefix + 1
    target_workbook = tax_year_folder / f"{next_prefix:02d} DAILY SALES {month_name(next_month_number)} {next_year}.xlsx"

    if dry_run:
        return source_path, target_workbook

    shutil.copy2(source_path, target_workbook)

    wb = openpyxl.load_workbook(target_workbook)
    ws = wb["TOTAL SALES"]

    ws["G3"] = month_name(next_month_number)
    set_date_range(ws, 5, next_month_number, next_year)
    clear_range(ws, "D5:G35")
    clear_range(ws, "J5:T35")
    update_reference_dates(ws, next_month_number, next_year)
    clear_range(ws, "C40:G41")
    clear_range(ws, "J40:T41")
    ws["A37"] = 1

    wb.save(target_workbook)
    return source_path, target_workbook


def main() -> int:
    parser = argparse.ArgumentParser(description="Create the next DAILY SALES workbook from the latest one.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Project root containing 'DAILY PROCEDURE DOCUMENTS'. Defaults to the current directory.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be created without writing files.")
    args = parser.parse_args()

    source_workbook, target_workbook = process_daily_sales(args.root, dry_run=args.dry_run)
    if args.dry_run:
        print(f"Would copy: {source_workbook}")
        print(f"Would create: {target_workbook}")
    else:
        print(f"Created: {target_workbook}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
