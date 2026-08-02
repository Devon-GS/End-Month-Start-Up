from __future__ import annotations

import argparse
import re
import shutil
from copy import copy
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
	match = re.match(r"^(\d{2})\s+CASHER SHORT AND OVER DAILY\s+([A-Z]+)\s+(\d{4})\.XLSX$", path.name.upper())
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
	candidates = [
			p
			for p in root.iterdir()
			if p.is_dir() and "CASHERS SHORTS AND OVERS TAX YR" in p.name.upper()
		]
	if not candidates:
			raise FileNotFoundError(f"No cashiers tax year folders found in {root}")
		
	def get_sort_number(path: Path) -> int:
		try:
			# Splits by space and converts the first part ("10", "2", etc.) to an integer
			return int(path.name.split()[0])
		except ValueError:
			return 0  # Fallback just in case a folder doesn't start with a number

	return sorted(candidates, key=get_sort_number)[-1]

def find_latest_source_workbook(tax_year_folder: Path) -> Path:
    candidates = []
    for path in tax_year_folder.iterdir():
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$"):
            try:
                prefix, _, _ = parse_workbook_name(path)
            except ValueError:
                continue
            candidates.append((prefix, path))
    if not candidates:
        raise FileNotFoundError(f"No source workbook found in {tax_year_folder}")
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def set_month_dates(ws, month: int, year: int) -> None:
    days_in_month = (date(year + int(month == 12), 1 if month == 12 else month + 1, 1) - date(year, month, 1)).days
    for day in range(1, days_in_month + 1):
        cell = ws.cell(row=day + 1, column=1)
        cell.value = date(year, month, day)
        cell.number_format = "dd/mm/yyyy"


def clear_cells(ws, ranges: list[str]) -> None:
    for cell_range in ranges:
        for row in ws[cell_range]:
            for cell in row:
                cell.value = None


def reset_range_style(ws, cell_range: str) -> None:
    from openpyxl.styles import PatternFill

    for row in ws[cell_range]:
        for cell in row:
            new_font = copy(cell.font)
            new_font.color = "FF000000"
            cell.font = new_font

            cell.fill = PatternFill()


def process_cashiers_shorts_and_overs(root: Path, dry_run: bool = False) -> tuple[Path, Path]:
    cashiers_root = root / "DAILY PROCEDURE DOCUMENTS" / "CASHIERS SHORTS AND OVERS"
    tax_year_folder = find_latest_tax_year_folder(cashiers_root)
    source_workbook = find_latest_source_workbook(tax_year_folder)
    source_prefix, source_month, source_year = parse_workbook_name(source_workbook)
    next_month_number, next_year = next_month(source_month, source_year)
    next_prefix = source_prefix + 1

    target_workbook = tax_year_folder / f"{next_prefix:02d} CASHER SHORT AND OVER DAILY {month_name(next_month_number)} {next_year}.xlsx"

    if dry_run:
        return source_workbook, target_workbook

    shutil.copy2(source_workbook, target_workbook)

    wb = openpyxl.load_workbook(target_workbook)
    ws = wb["Cashier Short"]

    set_month_dates(ws, next_month_number, next_year)
    clear_cells(
        ws,
        [
            "B2:F32",
            "H2:H32",
        ],
    )
    reset_range_style(ws, "B2:F32")

    wb.save(target_workbook)
    return source_workbook, target_workbook


def main() -> int:
    parser = argparse.ArgumentParser(description="Create the next CASHIERS SHORTS AND OVERS workbook from the latest one.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Project root containing 'DAILY PROCEDURE DOCUMENTS'. Defaults to the current directory.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be created without writing files.")
    args = parser.parse_args()

    source_workbook, target_workbook = process_cashiers_shorts_and_overs(args.root, dry_run=args.dry_run)
    if args.dry_run:
        print(f"Would copy: {source_workbook}")
        print(f"Would create: {target_workbook}")
    else:
        print(f"Created: {target_workbook}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
