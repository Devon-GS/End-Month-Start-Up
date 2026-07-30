from __future__ import annotations

import argparse
import re
import shutil
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


def parse_sheet_title(title: str) -> tuple[int, int]:
    match = re.match(r"^BILL PAY B LABEL\s+([A-Z]+)\s+(\d{2})$", title.upper())
    if not match:
        raise ValueError(f"Unrecognized sheet title: {title}")

    month_text = match.group(1)
    year_suffix = int(match.group(2))
    try:
        month = MONTH_NAMES.index(month_text) + 1
    except ValueError as exc:
        raise ValueError(f"Unrecognized month name in sheet title: {title}") from exc

    year = 2000 + year_suffix
    return month, year


def find_latest_sheet(wb: openpyxl.Workbook):
    for title in reversed(wb.sheetnames):
        if title.upper().startswith("BILL PAY B LABEL "):
            return wb[title]
    raise FileNotFoundError("No BILL PAY B LABEL sheets found in workbook")


def build_new_sheet_title(month: int, year: int) -> str:
    return f"BILL PAY B LABEL {month_name(month).title()} {str(year)[-2:]}"


def find_latest_tax_year_folder(root: Path) -> Path:
    candidates = []
    for folder in root.iterdir():
        if not folder.is_dir():
            continue
        match = re.match(r"^BILL PAYMENT RECON MARCH (\d{4}) TO FEB (\d{4})$", folder.name.upper())
        if not match:
            continue
        start_year = int(match.group(1))
        end_year = int(match.group(2))
        if end_year != start_year + 1:
            continue
        candidates.append((start_year, folder))
    if not candidates:
        raise FileNotFoundError(f"No BILL PAYMENT RECON tax year folders found in {root}")
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def process_bill_payment_recon(root: Path, dry_run: bool = False) -> tuple[Path, str]:
    recon_root = root / "DAILY PROCEDURE DOCUMENTS" / "DAILY BILL PAYMENT RECON"
    tax_year_folder = find_latest_tax_year_folder(recon_root)
    recon_file = tax_year_folder / "BILL PAYMENTS.xlsx"
    if not recon_file.exists():
        raise FileNotFoundError(f"Workbook not found: {recon_file}")

    wb = openpyxl.load_workbook(recon_file)
    source_ws = find_latest_sheet(wb)
    source_month, source_year = parse_sheet_title(source_ws.title)
    next_month_number, next_year = next_month(source_month, source_year)
    new_sheet_title = build_new_sheet_title(next_month_number, next_year)

    if dry_run:
        return recon_file, new_sheet_title

    data_wb = openpyxl.load_workbook(recon_file, data_only=True)
    data_ws = data_wb[source_ws.title]
    new_ws = wb.copy_worksheet(source_ws)
    new_ws.title = new_sheet_title

    source_value = data_ws["E36"].value
    if source_value is None:
        raise ValueError("E36 has no cached value to copy into B6")

    new_ws["B6"] = abs(source_value)
    for row in range(7, 37):
        new_ws[f"B{row}"] = f"=SUM(E{row - 1})"

    for row in range(6, 37):
        new_ws[f"C{row}"] = None
        new_ws[f"D{row}"] = None
        new_ws[f"F{row}"] = None

    wb.save(recon_file)
    return recon_file, new_sheet_title


def main() -> int:
    parser = argparse.ArgumentParser(description="Append the next BILL PAYMENT RECON sheet to BILL PAYMENTS.xlsx.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Project root containing 'DAILY PROCEDURE DOCUMENTS'. Defaults to the current directory.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be created without writing files.")
    args = parser.parse_args()

    workbook_path, new_sheet_title = process_bill_payment_recon(args.root, dry_run=args.dry_run)
    if args.dry_run:
        print(f"Would update: {workbook_path}")
        print(f"Would append sheet: {new_sheet_title}")
    else:
        print(f"Updated: {workbook_path}")
        print(f"Added sheet: {new_sheet_title}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
