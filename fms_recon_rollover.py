from __future__ import annotations

import argparse
import re
from copy import copy
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
        if re.match(r"^[A-Z]+(?:\s+[A-Z]+)?\s+\d{2}$", title.strip(), re.IGNORECASE):
            return wb[title]
    raise FileNotFoundError("No monthly FMS sheet found in workbook")


def build_new_sheet_title(month: int, year: int) -> str:
    return f"{month_name(month)} {str(year)[-2:]}"


def find_latest_tax_year_folder(root: Path) -> Path:
    candidates: list[tuple[int, Path]] = []
    for folder in root.iterdir():
        if not folder.is_dir():
            continue
        match = re.match(r"^FMS MARCH (\d{4}) TO FEBRUARY (\d{4})$", folder.name.upper())
        if not match:
            continue
        start_year = int(match.group(1))
        end_year = int(match.group(2))
        if end_year != start_year + 1:
            continue
        candidates.append((start_year, folder))
    if not candidates:
        raise FileNotFoundError(f"No FMS tax year folders found in {root}")
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def clear_range(ws, cell_range: str) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.value = None


def reset_fill(ws, cell_range: str) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.fill = PatternFill()


def set_fill(ws, cell_range: str, color: str) -> None:
    fill = PatternFill(fill_type="solid", fgColor=color)
    for row in ws[cell_range]:
        for cell in row:
            cell.fill = copy(fill)


def set_running_formula(ws, column: str, start_row: int, end_row: int) -> None:
    for row in range(start_row, end_row + 1):
        ws[f"{column}{row}"] = f"=SUM(E{row}*H{row})"


def set_month_dates(ws, month: int, year: int) -> None:
    days_in_month = (date(year + int(month == 12), 1 if month == 12 else month + 1, 1) - date(year, month, 1)).days
    for offset in range(days_in_month):
        row = 8 + offset
        ws[f"A{row}"] = date(year, month, offset + 1)


def process_fms_recon(root: Path, dry_run: bool = False) -> tuple[Path, str]:
    fms_root = root / "FMS Recon"
    tax_year_folder = find_latest_tax_year_folder(fms_root)
    workbook_candidates = [
        path
        for path in tax_year_folder.iterdir()
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$")
    ]
    if not workbook_candidates:
        raise FileNotFoundError(f"Workbook not found in: {tax_year_folder}")
    workbook_path = sorted(workbook_candidates)[-1]

    wb = openpyxl.load_workbook(workbook_path)
    source_ws = find_latest_sheet(wb)
    source_month, source_year = parse_sheet_title(source_ws.title)
    next_month_number, next_year = next_month(source_month, source_year)
    new_sheet_title = build_new_sheet_title(next_month_number, next_year)

    if dry_run:
        return workbook_path, new_sheet_title

    new_ws = wb.copy_worksheet(source_ws)
    new_ws.title = new_sheet_title

    clear_range(new_ws, "A8:A210")
    clear_range(new_ws, "B8:E210")
    clear_range(new_ws, "K8:L210")
    clear_range(new_ws, "M8:M210")
    reset_fill(new_ws, "A8:A210")
    set_fill(new_ws, "I8:I210", "FFC4BD97")
    set_running_formula(new_ws, "I", 9, 210)
    reset_fill(new_ws, "M8:M210")

    wb.save(workbook_path)
    return workbook_path, new_sheet_title


def main() -> int:
    parser = argparse.ArgumentParser(description="Create the next FMS Recon sheet in the existing workbook.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Project root containing 'FMS Recon'. Defaults to the current directory.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be created without writing files.")
    args = parser.parse_args()

    workbook_path, sheet_title = process_fms_recon(args.root, dry_run=args.dry_run)
    if args.dry_run:
        print(f"Would update: {workbook_path}")
        print(f"Would append sheet: {sheet_title}")
    else:
        print(f"Updated: {workbook_path}")
        print(f"Added sheet: {sheet_title}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
