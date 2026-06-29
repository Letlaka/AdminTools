from __future__ import annotations

import csv
import io
import sys
from contextlib import redirect_stderr
from pathlib import Path
import tempfile
import types
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import build_ad_excel_reports


FIELDNAMES = [
    "ComputerType",
    "Name",
    "CN",
    "DNSHostName",
    "Description",
    "OUPath",
    "CanonicalName",
    "DistinguishedName",
    "Created",
    "Enabled",
    "IPv4Address",
    "LastLogonDate",
    "LastLogonTimestamp",
    "LastSeenDate",
    "DaysSinceLastSeen",
    "InactiveThresholdDays",
    "IsStale",
    "StaleStatus",
    "ObjectGUID",
    "OperatingSystem",
    "OperatingSystemVersion",
]


class BuildAdExcelReportsTests(unittest.TestCase):
    def _write_csv(self, path: Path) -> None:
        with path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
            writer.writeheader()
            writer.writerow(
                {
                    "ComputerType": "Workstation",
                    "CN": "GPTRANSPORT01",
                    "DNSHostName": "gptransport01.example.local",
                    "OUPath": "gauteng/GPtransport",
                    "DistinguishedName": "CN=GPTRANSPORT01,OU=GPtransport,DC=example,DC=local",
                    "OperatingSystem": "Windows 11 Enterprise",
                    "Enabled": "True",
                }
            )
            writer.writerow(
                {
                    "ComputerType": "Workstation",
                    "CN": "HEALTH01",
                    "DNSHostName": "health01.example.local",
                    "OUPath": "gauteng/Health",
                    "DistinguishedName": "CN=HEALTH01,OU=Health,DC=example,DC=local",
                    "OperatingSystem": "Windows 11 Enterprise",
                    "Enabled": "True",
                }
            )

    def _write_department_files(self, root: Path) -> tuple[Path, Path]:
        departments_path = root / "dept_list.txt"
        dept_codes_path = root / "dept_codes.txt"
        departments_path.write_text("GPtransport\nHealth\n", encoding="utf-8")
        dept_codes_path.write_text("", encoding="utf-8")
        return departments_path, dept_codes_path

    def _fake_excel_writer(self, calls: dict[str, list[object]]) -> types.ModuleType:
        module = types.ModuleType("lib.excel_writer")

        def create_consolidated_workbook(**kwargs: object) -> None:
            calls["consolidated"].append(kwargs["output_path"])

        def create_department_workbook(**kwargs: object) -> None:
            calls["departments"].append(kwargs["department"])

        def sanitize_file_component(value: str) -> str:
            return value.replace(" ", "_")

        module.create_consolidated_workbook = create_consolidated_workbook
        module.create_department_workbook = create_department_workbook
        module.sanitize_file_component = sanitize_file_component
        return module

    def _run_report(self, extra_args: list[str]) -> tuple[int, dict[str, list[object]]]:
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            input_path = root / "workstations.csv"
            output_root = root / "reports"
            departments_path, dept_codes_path = self._write_department_files(root)
            self._write_csv(input_path)

            calls: dict[str, list[object]] = {"consolidated": [], "departments": []}
            argv = [
                "--workstations",
                str(input_path),
                "--departments",
                str(departments_path),
                "--dept-codes",
                str(dept_codes_path),
                "--output-root",
                str(output_root),
                "--as-of-date",
                "2026-06-29",
                *extra_args,
            ]
            with patch.dict(sys.modules, {"lib.excel_writer": self._fake_excel_writer(calls)}):
                return build_ad_excel_reports.main(argv), calls

    def test_department_scope_generates_only_named_department(self) -> None:
        exit_code, calls = self._run_report(
            ["--report-scope", "department", "--department", "GPtransport"]
        )

        self.assertEqual(exit_code, 0)
        self.assertEqual(calls["consolidated"], [])
        self.assertEqual(calls["departments"], ["GPtransport"])

    def test_main_scope_generates_only_consolidated_workbook(self) -> None:
        exit_code, calls = self._run_report(["--report-scope", "main"])

        self.assertEqual(exit_code, 0)
        self.assertEqual(len(calls["consolidated"]), 1)
        self.assertEqual(calls["departments"], [])

    def test_departments_scope_generates_all_department_workbooks_without_main(self) -> None:
        exit_code, calls = self._run_report(["--report-scope", "departments"])

        self.assertEqual(exit_code, 0)
        self.assertEqual(calls["consolidated"], [])
        self.assertEqual(calls["departments"], ["GPtransport", "Health", "Domain Controllers", "Unknown"])

    def test_department_scope_rejects_unknown_department(self) -> None:
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            exit_code, calls = self._run_report(
                ["--report-scope", "department", "--department", "Unknown Department"]
            )

        self.assertEqual(exit_code, 2)
        self.assertEqual(calls["consolidated"], [])
        self.assertEqual(calls["departments"], [])
        self.assertIn("Department 'Unknown Department' was not found", stderr.getvalue())

    def test_department_scope_requires_department_name(self) -> None:
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            exit_code, calls = self._run_report(["--report-scope", "department"])

        self.assertEqual(exit_code, 2)
        self.assertEqual(calls["consolidated"], [])
        self.assertEqual(calls["departments"], [])
        self.assertIn("--department is required", stderr.getvalue())


if __name__ == "__main__":
    unittest.main()
