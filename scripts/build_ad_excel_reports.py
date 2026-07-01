#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
from datetime import date, datetime
import shutil
import sys
from pathlib import Path

from lib.department_mapping import DepartmentMatcher, load_departments, load_dept_codes
from lib.financial_year import get_financial_year, get_run_folder
from lib.report_model import build_report_bundle, unmatched_records
from lib.spreadsheet_safety import safe_cell_value
from lib.scan_ad_input import load_scan_records, validate_scan_records


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build Excel dashboards and department workbooks from Scan-ADComputers exports."
    )
    parser.add_argument("--input", help="Combined Scan-ADComputers CSV or JSON export.")
    parser.add_argument("--servers", help="Server Scan-ADComputers CSV or JSON export.")
    parser.add_argument("--workstations", help="Workstation Scan-ADComputers CSV or JSON export.")
    parser.add_argument(
        "--departments",
        default="config/dept_list.txt",
        help="Department list file.",
    )
    parser.add_argument(
        "--dept-codes",
        default="config/dept_codes.txt",
        help="Department code mapping file.",
    )
    parser.add_argument(
        "--output-root",
        default="reports",
        help="Root folder for generated output.",
    )
    parser.add_argument(
        "--logs-root",
        default="logs",
        help="Root folder for generated logs.",
    )
    parser.add_argument(
        "--as-of-date",
        default=None,
        help="Run date in YYYY-MM-DD format. Defaults to today.",
    )
    parser.add_argument(
        "--financial-year-start-month",
        type=int,
        default=4,
        help="Financial year start month. Default is 4 for April.",
    )
    parser.add_argument(
        "--report-scope",
        choices=("all", "main", "departments", "department"),
        default="all",
        help="Report output scope. Default is all.",
    )
    parser.add_argument(
        "--department",
        help="Department name to generate when --report-scope department is used.",
    )
    return parser.parse_args(argv)


def resolve_run_date(raw_value: str | None) -> date:
    if not raw_value:
        return date.today()
    try:
        return datetime.strptime(raw_value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError("--as-of-date must be in YYYY-MM-DD format") from exc


def prepare_output_folders(output_root: Path, logs_root: Path, run_date: date, financial_year: str) -> dict[str, Path]:
    run_root = output_root / financial_year / get_run_folder(run_date)
    folders = {
        "root": run_root,
        "source": run_root / "source",
        "consolidated": run_root / "consolidated",
        "departments": run_root / "departments",
        "logs": logs_root / "excel-reporting" / financial_year / get_run_folder(run_date),
    }
    for path in folders.values():
        path.mkdir(parents=True, exist_ok=True)
    return folders


def copy_source_files(paths: list[Path], destination: Path) -> None:
    seen: set[Path] = set()
    for path in paths:
        if path in seen:
            continue
        seen.add(path)
        shutil.copy2(path, destination / path.name)


def load_all_records(args: argparse.Namespace) -> tuple[list[dict[str, object]], list[Path]]:
    records: list[dict[str, object]] = []
    source_paths: list[Path] = []

    if args.input:
        input_path = Path(args.input)
        batch = load_scan_records(input_path)
        validate_scan_records(batch, str(input_path))
        records.extend(batch)
        source_paths.append(input_path)

    if args.servers:
        server_path = Path(args.servers)
        batch = load_scan_records(server_path, forced_computer_type="Server")
        validate_scan_records(batch, str(server_path))
        records.extend(batch)
        source_paths.append(server_path)

    if args.workstations:
        workstation_path = Path(args.workstations)
        batch = load_scan_records(workstation_path, forced_computer_type="Workstation")
        validate_scan_records(batch, str(workstation_path))
        records.extend(batch)
        source_paths.append(workstation_path)

    return records, source_paths


def write_unmatched_log(path: Path, records: list[dict[str, object]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "ComputerType",
                "Name",
                "CN",
                "DNSHostName",
                "OUPath",
                "DistinguishedName",
                "Description",
                "SourceFile",
            ]
        )
        for record in records:
            writer.writerow(
                [
                    safe_cell_value(record.get("ComputerType")),
                    safe_cell_value(record.get("Name")),
                    safe_cell_value(record.get("CN")),
                    safe_cell_value(record.get("DNSHostName")),
                    safe_cell_value(record.get("OUPath")),
                    safe_cell_value(record.get("DistinguishedName")),
                    safe_cell_value(record.get("Description")),
                    safe_cell_value(record.get("SourceFile")),
                ]
            )


def validate_report_scope(args: argparse.Namespace) -> bool:
    if args.report_scope == "department" and not args.department:
        print("--department is required when --report-scope department is used.", file=sys.stderr)
        return False
    if args.report_scope != "department" and args.department:
        print("--department can only be used with --report-scope department.", file=sys.stderr)
        return False
    return True


def find_department_name(departments: list[str], requested_department: str) -> str | None:
    requested = requested_department.casefold()
    for department in departments:
        if department.casefold() == requested:
            return department
    return None


def departments_for_scope(args: argparse.Namespace, available_departments: list[str]) -> list[str]:
    if args.report_scope == "main":
        return []
    if args.report_scope in ("all", "departments"):
        return available_departments

    matched_department = find_department_name(available_departments, args.department)
    if matched_department is None:
        available = ", ".join(available_departments)
        raise ValueError(
            f"Department '{args.department}' was not found in the generated report data. "
            f"Available departments: {available}"
        )
    return [matched_department]


def write_consolidated_report(
    *,
    create_consolidated_workbook,
    folders: dict[str, Path],
    bundle,
    financial_year: str,
    run_date_label: str,
) -> Path:
    consolidated_name = f"AD_Dashboard_{financial_year}_{run_date_label}.xlsx"
    consolidated_path = folders["consolidated"] / consolidated_name
    create_consolidated_workbook(
        output_path=consolidated_path,
        bundle=bundle,
        financial_year=financial_year,
        run_date_label=run_date_label,
    )
    return consolidated_path


def write_department_report(
    *,
    create_department_workbook,
    sanitize_file_component,
    folders: dict[str, Path],
    bundle,
    department: str,
    financial_year: str,
    run_date_label: str,
) -> Path:
    department_records = bundle.records_for_department(department)
    department_folder = folders["departments"] / sanitize_file_component(department)
    workbook_path = department_folder / f"{sanitize_file_component(department)}_{financial_year}.xlsx"
    create_department_workbook(
        output_path=workbook_path,
        department=department,
        server_records=[record for record in department_records if record["ComputerType"] == "Server"],
        workstation_records=[
            record for record in department_records if record["ComputerType"] == "Workstation"
        ],
        financial_year=financial_year,
        run_date_label=run_date_label,
        department_summary=bundle.department_summary(department),
    )
    return workbook_path


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if not validate_report_scope(args):
        return 2

    if not any((args.input, args.servers, args.workstations)):
        print("Specify at least one of --input, --servers, or --workstations.", file=sys.stderr)
        return 2

    try:
        from lib.excel_writer import (
            create_consolidated_workbook,
            create_department_workbook,
            sanitize_file_component,
        )
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    run_date = resolve_run_date(args.as_of_date)
    financial_year = get_financial_year(run_date, start_month=args.financial_year_start_month)

    departments_path = Path(args.departments)
    dept_codes_path = Path(args.dept_codes)
    departments = load_departments(departments_path)
    if "Domain Controllers" not in departments:
        departments.append("Domain Controllers")

    matcher = DepartmentMatcher(departments=departments, code_map=load_dept_codes(dept_codes_path))
    records, source_paths = load_all_records(args)
    bundle = build_report_bundle(records=records, matcher=matcher, departments=departments, as_of_date=run_date)

    folders = prepare_output_folders(Path(args.output_root), Path(args.logs_root), run_date, financial_year)
    copy_source_files(source_paths, folders["source"])

    run_date_label = run_date.isoformat()
    consolidated_path: Path | None = None
    if args.report_scope in ("all", "main"):
        consolidated_path = write_consolidated_report(
            create_consolidated_workbook=create_consolidated_workbook,
            folders=folders,
            bundle=bundle,
            financial_year=financial_year,
            run_date_label=run_date_label,
        )

    try:
        selected_departments = departments_for_scope(args, bundle.departments_with_unknown)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    for department in selected_departments:
        write_department_report(
            create_department_workbook=create_department_workbook,
            sanitize_file_component=sanitize_file_component,
            folders=folders,
            bundle=bundle,
            department=department,
            financial_year=financial_year,
            run_date_label=run_date_label,
        )

    unmatched = unmatched_records(bundle.records)
    write_unmatched_log(folders["logs"] / "unmatched_devices.csv", unmatched)

    print(f"Financial year: {financial_year}")
    print(f"Run root: {folders['root']}")
    print(f"Log root: {folders['logs']}")
    print(f"Consolidated workbook: {consolidated_path if consolidated_path else 'not generated'}")
    print(f"Departments processed: {len(selected_departments)}")
    print(f"Total devices: {len(bundle.records)}")
    print(f"Unmatched devices: {len(unmatched)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
