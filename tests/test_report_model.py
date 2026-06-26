from __future__ import annotations

from datetime import date
import sys
from pathlib import Path
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from lib.department_mapping import DepartmentMatcher
from lib.report_model import build_report_bundle, classify_operating_system


class ReportModelTests(unittest.TestCase):
    def test_classify_server_versions(self) -> None:
        cases = [
            ("Windows Server 2008 Standard", "", "Windows Server 2008", True),
            ("Windows Server 2008 R2 Standard", "", "Windows Server 2008 R2", True),
            ("Windows Server 2012 Standard", "", "Windows Server 2012", True),
            ("Windows Server 2012 R2 Standard", "", "Windows Server 2012 R2", True),
            ("Windows Server 2003", "", "Windows Server 2003 or Older", True),
            ("Windows Server 2016 Datacenter", "", "Windows Server 2016", False),
            ("Windows Server 2019 Datacenter", "", "Windows Server 2019", False),
            ("Windows Server 2022 Datacenter", "", "Windows Server 2022", False),
        ]
        for operating_system, operating_system_version, expected_bucket, expected_legacy in cases:
            with self.subTest(operating_system=operating_system):
                result = classify_operating_system(
                    {
                        "ComputerType": "Server",
                        "OperatingSystem": operating_system,
                        "OperatingSystemVersion": operating_system_version,
                    }
                )
                self.assertEqual(result["OSLifecycleBucket"], expected_bucket)
                self.assertEqual(result["OSIsLegacy"], expected_legacy)

    def test_classify_workstation_versions(self) -> None:
        cases = [
            ("Windows 10 Enterprise", "10.0.19045", "Windows 10", True),
            ("Windows 11 Enterprise", "10.0.22631", "Windows 11", False),
            ("Windows 7 Professional", "6.1", "Windows 7", True),
            ("", "", "Non-Windows / Unknown", False),
        ]
        for operating_system, operating_system_version, expected_bucket, expected_legacy in cases:
            with self.subTest(operating_system=operating_system):
                result = classify_operating_system(
                    {
                        "ComputerType": "Workstation",
                        "OperatingSystem": operating_system,
                        "OperatingSystemVersion": operating_system_version,
                    }
                )
                self.assertEqual(result["OSLifecycleBucket"], expected_bucket)
                self.assertEqual(result["OSIsLegacy"], expected_legacy)

    def test_device_records_include_normalized_inactivity_and_dashboard_excludes_stale_column(self) -> None:
        records = [
            {
                "ComputerType": "Server",
                "CN": "SRCSTALE",
                "DNSHostName": "srcstale.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=SRCSTALE,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows Server 2019",
                "StaleStatus": "Stale",
                "DaysSinceLastSeen": 120,
                "LastLogonDate": date(2026, 6, 20),
            },
            {
                "ComputerType": "Server",
                "CN": "SOURCE_DAYS_STALE",
                "DNSHostName": "source-days-stale.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=SOURCE_DAYS_STALE,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows Server 2019",
                "StaleStatus": "NotEvaluated",
                "DaysSinceLastSeen": 749,
            },
            {
                "ComputerType": "Server",
                "CN": "COMPUTEDSTALE",
                "DNSHostName": "computedstale.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=COMPUTEDSTALE,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows Server 2019",
                "LastLogonDate": date(2026, 3, 1),
            },
            {
                "ComputerType": "Workstation",
                "CN": "FRESH",
                "DNSHostName": "fresh.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=FRESH,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows 11 Enterprise",
                "LastLogonDate": date(2026, 6, 1),
            },
            {
                "ComputerType": "Workstation",
                "CN": "UNKNOWN",
                "DNSHostName": "unknown.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=UNKNOWN,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows 11 Enterprise",
            },
            {
                "ComputerType": "Workstation",
                "CN": "SOURCE_STALE_UNKNOWN",
                "DNSHostName": "source-stale-unknown.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=SOURCE_STALE_UNKNOWN,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows 11 Enterprise",
                "StaleStatus": "Stale",
            },
        ]
        matcher = DepartmentMatcher(departments=["Health"], code_map={})
        bundle = build_report_bundle(
            records=records,
            matcher=matcher,
            departments=["Health"],
            as_of_date=date(2026, 6, 25),
        )

        dashboard_row = bundle.dashboard_rows()[0]
        self.assertNotIn("Stale", dashboard_row)
        records_by_cn = {record["CN"]: record for record in bundle.records}
        self.assertNotIn("StaleDays", records_by_cn["SRCSTALE"])
        self.assertEqual(records_by_cn["SRCSTALE"]["InactivityDays"], 120)
        self.assertEqual(records_by_cn["SRCSTALE"]["InactivityStatus"], "Stale")
        self.assertEqual(records_by_cn["SOURCE_DAYS_STALE"]["InactivityDays"], 749)
        self.assertEqual(records_by_cn["SOURCE_DAYS_STALE"]["InactivityStatus"], "Stale")
        self.assertEqual(records_by_cn["COMPUTEDSTALE"]["InactivityDays"], 116)
        self.assertEqual(records_by_cn["COMPUTEDSTALE"]["InactivityStatus"], "Stale")
        self.assertEqual(records_by_cn["FRESH"]["InactivityDays"], 24)
        self.assertEqual(records_by_cn["FRESH"]["InactivityStatus"], "Fresh")
        self.assertIsNone(records_by_cn["UNKNOWN"]["InactivityDays"])
        self.assertEqual(records_by_cn["UNKNOWN"]["InactivityStatus"], "Unknown")
        self.assertIsNone(records_by_cn["SOURCE_STALE_UNKNOWN"]["InactivityDays"])
        self.assertEqual(records_by_cn["SOURCE_STALE_UNKNOWN"]["InactivityStatus"], "Stale")

    def test_last_seen_date_is_normalized_from_last_logon_date_for_servers_and_workstations(self) -> None:
        records = [
            {
                "ComputerType": "Server",
                "CN": "SERVER01",
                "DNSHostName": "server01.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=SERVER01,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows Server 2019",
                "LastLogonDate": "2026-06-20T14:30:00",
            },
            {
                "ComputerType": "Workstation",
                "CN": "PC01",
                "DNSHostName": "pc01.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=PC01,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows 11 Enterprise",
                "LastLogonDate": "2026/06/01",
            },
            {
                "ComputerType": "Workstation",
                "CN": "PRESERVE01",
                "DNSHostName": "preserve01.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=PRESERVE01,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows 11 Enterprise",
                "LastSeenDate": "2026-06-10",
                "LastLogonDate": "2026-06-01",
            },
        ]
        matcher = DepartmentMatcher(departments=["Health"], code_map={})
        bundle = build_report_bundle(
            records=records,
            matcher=matcher,
            departments=["Health"],
            as_of_date=date(2026, 6, 25),
        )

        records_by_cn = {record["CN"]: record for record in bundle.records}
        self.assertEqual(records_by_cn["SERVER01"]["LastSeenDate"], date(2026, 6, 20))
        self.assertEqual(records_by_cn["SERVER01"]["InactivityDays"], 5)
        self.assertEqual(records_by_cn["PC01"]["LastSeenDate"], date(2026, 6, 1))
        self.assertEqual(records_by_cn["PC01"]["InactivityDays"], 24)
        self.assertEqual(records_by_cn["PRESERVE01"]["LastSeenDate"], date(2026, 6, 10))
        self.assertEqual(records_by_cn["PRESERVE01"]["InactivityDays"], 15)

    def test_bundle_summaries_include_legacy_counts(self) -> None:
        records = [
            {
                "ComputerType": "Server",
                "CN": "SRV2008",
                "DNSHostName": "srv2008.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=SRV2008,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows Server 2008 R2 Standard",
            },
            {
                "ComputerType": "Server",
                "CN": "SRV2022",
                "DNSHostName": "srv2022.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=SRV2022,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows Server 2022 Datacenter",
            },
            {
                "ComputerType": "Workstation",
                "CN": "PCWIN7",
                "DNSHostName": "pcwin7.example.local",
                "OUPath": "example/Health",
                "DistinguishedName": "CN=PCWIN7,OU=Health,DC=example,DC=local",
                "OperatingSystem": "Windows 7 Professional",
            },
        ]
        matcher = DepartmentMatcher(departments=["Health"], code_map={})
        bundle = build_report_bundle(records=records, matcher=matcher, departments=["Health"])

        legacy_rows = bundle.legacy_department_rows()
        self.assertEqual(legacy_rows[0]["LegacyServers"], 1)
        self.assertEqual(legacy_rows[0]["LegacyWorkstations"], 1)

        spotlight_rows = bundle.legacy_spotlight_rows()
        spotlight_lookup = {row["OSLifecycleBucket"]: row["Count"] for row in spotlight_rows}
        self.assertEqual(spotlight_lookup["Windows Server 2008 R2"], 1)

        summary = bundle.department_summary("Health")
        self.assertEqual(summary["LegacyServers"], 1)
        self.assertEqual(summary["LegacyWorkstations"], 1)


if __name__ == "__main__":
    unittest.main()
