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


def month_name(month: int) -> str:
    return MONTH_NAMES[month - 1]


def next_month(month: int, year: int) -> tuple[int, int]:
    if month == 12:
        return 1, year + 1
    return month + 1, year


def parse_workbook_name(path: Path) -> tuple[int, int, int]:
    match = re.match(r"^(\d{2})\s+GAS BOTTLE\s+([A-Z]+)\s+(\d{4})\.XLSX$", path.name.upper())
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
        match = re.match(r"^14 GAS BOTTLES TAX YR MARCH (\d{4}) TO FEB (\d{4})$", folder.name.upper())
        if not match:
            continue
        start_year = int(match.group(1))
        end_year = int(match.group(2))
        if end_year != start_year + 1:
            continue
        candidates.append((start_year, folder))
    if not candidates:
        raise FileNotFoundError(f"No GAS BOTTLES tax year folders found in {root}")
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def find_latest_source_workbook(folder: Path) -> Path:
    candidates = []
    for path in folder.iterdir():
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$"):
            try:
                prefix, _, _ = parse_workbook_name(path)
            except ValueError:
                continue
            candidates.append((prefix, path))
    if not candidates:
        raise FileNotFoundError(f"No GAS BOTTLES workbook found in {folder}")
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def clear_range(ws, cell_range: str) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.value = None


def set_date_range(ws, start_row: int, month: int, year: int) -> None:
    days_in_month = calendar.monthrange(year, month)[1]
    for offset in range(days_in_month):
        row = start_row + offset
        ws[f"A{row}"] = date(year, month, offset + 1)


def positive_value(value):
    if value is None:
        return None
    if value == 0 or value == "0":
        return 0
    if isinstance(value, (int, float)):
        return abs(value)
    try:
        return abs(float(value))
    except (TypeError, ValueError):
        return value


def numeric_or_zero(value):
    if value is None or value == 0 or value == "0":
        return 0
    return value


def negative_value(value):
    if value is None:
        return None
    if value == 0 or value == "0":
        return 0
    if isinstance(value, (int, float)):
        return -abs(value)
    try:
        return -abs(float(value))
    except (TypeError, ValueError):
        return value


def process_gas_bottles(root: Path, dry_run: bool = False) -> tuple[Path, Path]:
    gas_root = root / "DAILY PROCEDURE DOCUMENTS" / "GAS BOTTLES"
    tax_year_folder = find_latest_tax_year_folder(gas_root)
    source_workbook = find_latest_source_workbook(tax_year_folder)
    source_prefix, source_month, source_year = parse_workbook_name(source_workbook)
    next_month_number, next_year = next_month(source_month, source_year)
    next_prefix = source_prefix + 1
    target_workbook = tax_year_folder / f"{next_prefix:02d} Gas bottle {month_name(next_month_number)} {next_year}.xlsx"

    if dry_run:
        return source_workbook, target_workbook

    shutil.copy2(source_workbook, target_workbook)

    wb = openpyxl.load_workbook(target_workbook)
    data_wb = openpyxl.load_workbook(source_workbook, data_only=True)
    ws = wb.active
    source_ws = data_wb.active

    ws.title = f"GAS RECON {month_name(next_month_number)} {next_year}"
    wb.active = wb.sheetnames.index(ws.title)

    set_date_range(ws, 5, next_month_number, next_year)

    b36_value = source_ws["B36"].value
    if b36_value is None or b36_value == 0 or b36_value == "0":
        b35_value = source_ws["B35"].value
        if b35_value is None or b35_value == 0 or b35_value == "0":
            b4_value = ((numeric_or_zero(source_ws["B34"].value) + numeric_or_zero(source_ws["D34"].value)) - numeric_or_zero(source_ws["F34"].value)) - numeric_or_zero(source_ws["G34"].value)
        else:
            b4_value = numeric_or_zero(b35_value) + numeric_or_zero(source_ws["D35"].value) - numeric_or_zero(source_ws["F35"].value)
    else:
        b4_value = b36_value
    ws["B4"] = b4_value

    c36_value = source_ws["C36"].value
    if c36_value is None or c36_value == 0 or c36_value == "0":
        c35_value = source_ws["C35"].value
        if c35_value is None or c35_value == 0 or c35_value == "0":
            c4_value = ((numeric_or_zero(source_ws["C34"].value) - numeric_or_zero(source_ws["E34"].value)) + numeric_or_zero(source_ws["F34"].value)) - numeric_or_zero(source_ws["H34"].value)
        else:
            c4_value = numeric_or_zero(c35_value) - numeric_or_zero(source_ws["E35"].value) + numeric_or_zero(source_ws["F35"].value) + numeric_or_zero(source_ws["H35"].value)
    else:
        c4_value = c36_value
    ws["C4"] = c4_value

    clear_range(ws, "B5:I37")
    clear_range(ws, "J4:J37")

    wb.save(target_workbook)
    return source_workbook, target_workbook


def main() -> int:
    parser = argparse.ArgumentParser(description="Create the next GAS BOTTLES workbook from the latest one.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Project root containing 'DAILY PROCEDURE DOCUMENTS'. Defaults to the current directory.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be created without writing files.")
    args = parser.parse_args()

    source_workbook, target_workbook = process_gas_bottles(args.root, dry_run=args.dry_run)
    if args.dry_run:
        print(f"Would copy: {source_workbook}")
        print(f"Would create: {target_workbook}")
    else:
        print(f"Created: {target_workbook}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
