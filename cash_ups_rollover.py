from __future__ import annotations

import argparse
import re
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


def _load_openpyxl():
    site_packages = Path(r"C:\Users\Devon\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\Lib\site-packages")
    if site_packages.exists():
        sys.path.insert(0, str(site_packages))
    import openpyxl  # type: ignore
    from openpyxl.cell.cell import MergedCell  # type: ignore

    return openpyxl, MergedCell


openpyxl, MergedCell = _load_openpyxl()


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


@dataclass(frozen=True)
class MonthFolder:
    path: Path
    prefix: int
    month: int
    year: int


def month_name(month: int) -> str:
    return MONTH_NAMES[month - 1]


def next_month(month: int, year: int) -> tuple[int, int]:
    if month == 12:
        return 1, year + 1
    return month + 1, year


def parse_month_folder(folder: Path) -> MonthFolder:
    match = re.match(r"^(\d{2})\s+([A-Z]+)\s+(\d{4})$", folder.name.upper())
    if not match:
        raise ValueError(f"Unrecognized month folder name: {folder.name}")

    prefix = int(match.group(1))
    month_text = match.group(2)
    year = int(match.group(3))

    try:
        month = MONTH_NAMES.index(month_text) + 1
    except ValueError as exc:
        raise ValueError(f"Unrecognized month name in folder: {folder.name}") from exc

    return MonthFolder(path=folder, prefix=prefix, month=month, year=year)


def find_latest_tax_year_folder(cash_ups_root: Path) -> Path:
    candidates = [p for p in cash_ups_root.iterdir() if p.is_dir() and p.name.upper().startswith("CASH UPS TAX YR")]
    if not candidates:
        raise FileNotFoundError(f"No tax year folders found in {cash_ups_root}")
    return sorted(candidates)[-1]


def find_latest_month_folder(tax_year_folder: Path) -> MonthFolder:
    candidates = []
    for folder in tax_year_folder.iterdir():
        if not folder.is_dir():
            continue
        try:
            month_folder = parse_month_folder(folder)
            find_latest_source_workbook(month_folder.path)
        except (ValueError, FileNotFoundError):
            continue
        candidates.append(month_folder)
    if not candidates:
        raise FileNotFoundError(f"No month folders found in {tax_year_folder}")
    return sorted(candidates, key=lambda item: item.prefix)[-1]


def find_latest_source_workbook(month_folder: Path) -> Path:
    candidates = []
    for path in month_folder.iterdir():
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$"):
            match = re.search(r"CONTROL SHEET DAILY\s+(\d{1,2})\s+[A-Z]+\s+\d{4}\.XLSX$", path.name.upper())
            if match:
                candidates.append((int(match.group(1)), path))
    if not candidates:
        raise FileNotFoundError(f"No source workbook found in {month_folder}")
    return sorted(candidates, key=lambda item: item[0])[-1][1]


def increment_batch_label(value: object) -> object:
    if not isinstance(value, str):
        return value

    match = re.search(r"(\d+)(?!.*\d)", value)
    if not match:
        raise ValueError(f"Could not find batch number in {value!r}")

    number_text = match.group(1)
    next_number = str(int(number_text) + 1).zfill(len(number_text))
    return value[: match.start(1)] + next_number + value[match.end(1) :]


def iter_cells_in_range(ws, cell_range: str) -> Iterable:
    if ":" not in cell_range:
        yield ws[cell_range]
        return

    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(cell_range)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            yield cell


def clear_unprotected_cells(ws, ranges: Iterable[str]) -> None:
    for cell_range in ranges:
        for cell in iter_cells_in_range(ws, cell_range):
            if isinstance(cell, MergedCell):
                continue
            if cell.protection.locked:
                continue
            cell.value = None


def apply_cash_ups_rules(target_workbook: Path, source_workbook: Path) -> None:
    data_only_source = openpyxl.load_workbook(source_workbook, data_only=True)
    target = openpyxl.load_workbook(target_workbook)

    ws_target = target["CASH UP"]
    ws_source_data = data_only_source["CASH UP"]

    clear_unprotected_cells(
        ws_target,
        [
            "B3:E11",
            "B13:E15",
            "A18:E33",
            "I26",
            "B36:E36",
            "B38:E38",
            "D44:D47",
            "H6:H8",
        ],
    )

    ws_target["A2"] = increment_batch_label(ws_target["A2"].value)
    ws_target["D42"] = ws_source_data["D48"].value

    target.save(target_workbook)


def build_next_month_folder(current_month: MonthFolder) -> MonthFolder:
    next_month_number, next_year = next_month(current_month.month, current_month.year)
    return MonthFolder(
        path=current_month.path.parent / f"{current_month.prefix + 1:02d} {month_name(next_month_number)} {next_year}",
        prefix=current_month.prefix + 1,
        month=next_month_number,
        year=next_year,
    )


def build_next_tax_year_folder(tax_year_folder: Path, next_month_folder: MonthFolder) -> Path:
    # Preserve the same naming pattern as the existing tax year folder.
    match = re.search(r"^(.*?)(\d{4}).*?(\d{4})$", tax_year_folder.name)
    if not match:
        return tax_year_folder

    start_year = next_month_folder.year if next_month_folder.month == 3 else int(match.group(2))
    end_year = start_year + 1
    prefix = match.group(1).rstrip()
    if prefix.endswith("TO Feb"):
        return tax_year_folder.parent / f"{prefix} {start_year} TO Feb {end_year}"
    return tax_year_folder.parent / f"{tax_year_folder.name.split(str(match.group(2)))[0]}{start_year} TO Feb {end_year}"


def process_cash_ups(root: Path, dry_run: bool = False) -> tuple[Path, Path]:
    cash_ups_root = root / "DAILY PROCEDURE DOCUMENTS" / "CASH UPS"
    tax_year_folder = find_latest_tax_year_folder(cash_ups_root)
    month_folder = find_latest_month_folder(tax_year_folder)
    source_workbook = find_latest_source_workbook(month_folder.path)
    next_month_folder = build_next_month_folder(month_folder)

    # If the month advances from February to March, move into the new tax year.
    if next_month_folder.month == 3 and month_folder.month == 2:
        tax_year_folder = build_next_tax_year_folder(tax_year_folder, next_month_folder)
        next_month_folder = MonthFolder(
            path=tax_year_folder / next_month_folder.path.name,
            prefix=next_month_folder.prefix,
            month=next_month_folder.month,
            year=next_month_folder.year,
        )

    target_workbook = next_month_folder.path / f"Control Sheet Daily 01 {month_name(next_month_folder.month)} {next_month_folder.year}.xlsx"

    if dry_run:
        return source_workbook, target_workbook

    next_month_folder.path.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_workbook, target_workbook)
    apply_cash_ups_rules(target_workbook, source_workbook)
    return source_workbook, target_workbook


def main() -> int:
    parser = argparse.ArgumentParser(description="Create the next CASH UPS month workbook from the latest one.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Project root containing 'DAILY PROCEDURE DOCUMENTS'. Defaults to the current directory.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Show what would be created without writing files.")
    args = parser.parse_args()

    source_workbook, target_workbook = process_cash_ups(args.root, dry_run=args.dry_run)
    if args.dry_run:
        print(f"Would copy: {source_workbook}")
        print(f"Would create: {target_workbook}")
    else:
        print(f"Created: {target_workbook}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
