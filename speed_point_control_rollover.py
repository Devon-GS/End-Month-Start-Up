from __future__ import annotations

import argparse
import calendar
import re
import shutil
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill


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


def parse_workbook_name(path: Path) -> tuple[int, int, int]:
    match = re.match(r"^(\d+)\s+SPEEDPOINT DAILY\s+([A-Z]+)\s+(\d{4})\.XLSX$", path.name.upper())
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
        match = re.match(r"^DAILY SPEEDPOINT FROM MARCH (\d{4}) TO FEB (\d{4})$", folder.name.upper())
        if not match:
            continue
        start_year = int(match.group(1))
        end_year = int(match.group(2))
        if end_year != start_year + 1:
            continue
        candidates.append((start_year, folder))
    if not candidates:
        raise FileNotFoundError(f"No SPEEDPOINT tax year folders found in {root}")
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def set_speedpoint_dates(ws, month: int, year: int) -> None:
    days_in_month = calendar.monthrange(year, month)[1]
    date_rows = list(range(4, 34, 2)) + list(range(42, 74, 2))
    if len(date_rows) < days_in_month:
        raise ValueError("Not enough date rows for the target month")

    for day, row in enumerate(date_rows[:days_in_month], start=1):
        ws[f"A{row}"] = date(year, month, day)


def numeric_or_zero(value):
    if value is None or value == 0 or value == "0":
        return 0
    return value


def remove_highlight(ws, cell_range: str) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.fill = PatternFill()


def process_speed_point_control(root: Path, dry_run: bool = False) -> tuple[Path, Path]:
    speed_root = root / "SPEED POINT CONTROL" / "1 - DAILY SPEEDPOINT CONTROL"
    tax_year_folder = find_latest_tax_year_folder(speed_root)
    source_workbooks = []
    for path in tax_year_folder.iterdir():
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$"):
            try:
                prefix, month, year = parse_workbook_name(path)
            except ValueError:
                continue
            source_workbooks.append((prefix, month, year, path))
    if not source_workbooks:
        raise FileNotFoundError(f"No SPEEDPOINT workbook found in {tax_year_folder}")

    source_prefix, source_month, source_year, source_path = sorted(source_workbooks, key=lambda item: item[0])[-1]
    next_month_number, next_year = next_month(source_month, source_year)
    next_prefix = source_prefix + 1
    target_workbook = tax_year_folder / f"{next_prefix} SPEEDPOINT DAILY {month_name(next_month_number)} {next_year}.xlsx"

    if dry_run:
        return source_path, target_workbook

    shutil.copy2(source_path, target_workbook)
    wb = openpyxl.load_workbook(target_workbook)
    data_wb = openpyxl.load_workbook(source_path, data_only=True)
    ws = wb["Speed Points"]
    set_speedpoint_dates(ws, next_month_number, next_year)

    source_ws = data_wb["Speed Points"]

    b73_value = source_ws["B73"].value
    if b73_value is None or b73_value == 0 or b73_value == "0":
        b4_value = numeric_or_zero(source_ws["B71"].value) + 1
    else:
        b4_value = numeric_or_zero(b73_value) + 1
    ws["B4"] = b4_value

    d73_value = source_ws["D73"].value
    if d73_value is None or d73_value == 0 or d73_value == "0":
        d4_value = numeric_or_zero(source_ws["D71"].value) + 1
    else:
        d4_value = numeric_or_zero(source_ws["D73"].value) + 1
    ws["D4"] = d4_value

    for cell_range in [
        "B5:B33",
        "B42:B73",
        "D5:D33",
        "D42:D73",
        "C4:C33",
        "C42:C73",
        "E4:E33",
        "E42:E73",
        "G4:P33",
        "G42:P73",
        "R4:S33",
        "R42:S73",
    ]:
        for row in ws[cell_range]:
            for cell in row:
                cell.value = None

    for cell_range in [
        "C4:C33",
        "C42:C73",
        "E4:E33",
        "E42:E73",
        "G4:P33",
        "G42:P73",
    ]:
        remove_highlight(ws, cell_range)

    wb.save(target_workbook)
    return source_path, target_workbook


def main() -> int:
    parser = argparse.ArgumentParser(description="Create the next SPEEDPOINT CONTROL workbook from the latest one.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Project root containing 'SPEED POINT CONTROL'. Defaults to the current directory.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be created without writing files.")
    args = parser.parse_args()

    source_workbook, target_workbook = process_speed_point_control(args.root, dry_run=args.dry_run)
    if args.dry_run:
        print(f"Would copy: {source_workbook}")
        print(f"Would create: {target_workbook}")
    else:
        print(f"Created: {target_workbook}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
