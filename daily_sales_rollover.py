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
            
        # Updated regex: (?:\s*\(\d+\))? optionally matches " (1)", " (2)", etc. at the end
        match = re.match(
            r"^DAILY SALES TAX YR MARCH (\d{4}) TO FEB (\d{4})(?:\s*\(\d+\))?$", 
            folder.name.upper()
        )
        
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


def update_monthly_day_range(ws, start_row: int, month: int, year: int) -> None:
    days_in_month = calendar.monthrange(year, month)[1]
    for offset in range(days_in_month):
        row = start_row + offset
        current = date(year, month, offset + 1)
        ws[f"A{row}"] = current


def process_product_sheet(ws, data_ws, month: int, year: int) -> None:
    update_monthly_day_range(ws, 6, month, year)

    k36_value = data_ws["K36"].value
    if k36_value is None or k36_value == 0 or k36_value == "0":
        k36_value = data_ws["K35"].value
    ws["K2"] = k36_value

    t36_value = data_ws["T36"].value
    if t36_value is None or t36_value == 0 or t36_value == "0":
        t36_value = data_ws["T35"].value
    ws["T2"] = t36_value

    clear_range(ws, "B6:J36")
    clear_range(ws, "M6:M36")
    clear_range(ws, "O6:P36")
    clear_range(ws, "R6:S36")


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
    data_wb = openpyxl.load_workbook(source_path, data_only=True)
    total_sales = wb["TOTAL SALES"]
    ulp_t1 = wb["ULP 95 T1"]
    source_ulp_t1 = data_wb["ULP 95 T1"]

    total_sales["G3"] = month_name(next_month_number)
    set_date_range(total_sales, 5, next_month_number, next_year)
    clear_range(total_sales, "D5:G35")
    clear_range(total_sales, "J5:T35")
    update_reference_dates(total_sales, next_month_number, next_year)
    clear_range(total_sales, "C40:G41")
    clear_range(total_sales, "J40:T41")
    total_sales["A37"] = 1

    ref_date = date(next_year, next_month_number, 1)
    ulp_t1["W1"] = ref_date
    update_monthly_day_range(ulp_t1, 6, next_month_number, next_year)

    k36_value = source_ulp_t1["K36"].value
    if k36_value is None or k36_value == 0 or k36_value == "0":
        k36_value = source_ulp_t1["K35"].value
    ulp_t1["K2"] = k36_value

    t36_value = source_ulp_t1["T36"].value
    if t36_value is None or t36_value == 0 or t36_value == "0":
        t36_value = source_ulp_t1["T35"].value
    ulp_t1["T2"] = t36_value

    clear_range(ulp_t1, "B6:J36")
    clear_range(ulp_t1, "M6:M36")
    clear_range(ulp_t1, "O6:P36")
    clear_range(ulp_t1, "R6:S36")

    ulp_t3 = wb["ULP 95 T3"]
    process_product_sheet(ulp_t3, data_wb["ULP 95 T3"], next_month_number, next_year)

    process_product_sheet(wb["Tdiesel ULS T2 (OLD LRP)"], data_wb["Tdiesel ULS T2 (OLD LRP)"], next_month_number, next_year)
    process_product_sheet(wb["Tdiesel ULS T4"], data_wb["Tdiesel ULS T4"], next_month_number, next_year)

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
