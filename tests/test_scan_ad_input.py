from __future__ import annotations

import csv
import json
import sys
from pathlib import Path
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from lib.scan_ad_input import load_scan_records, validate_scan_records


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


class ScanAdInputTests(unittest.TestCase):
    def test_load_csv_records(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "devices.csv"
            with path.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
                writer.writeheader()
                writer.writerow(
                    {
                        "ComputerType": "Workstation",
                        "CN": "PC01",
                        "DNSHostName": "pc01.example.local",
                        "OUPath": "example/Health",
                        "DistinguishedName": "CN=PC01,OU=Health,DC=example,DC=local",
                        "OperatingSystem": "Windows 11 Enterprise",
                        "Enabled": "True",
                    }
                )

            records = load_scan_records(path)
            validate_scan_records(records, str(path))
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["Enabled"], True)
            self.assertEqual(records[0]["ComputerType"], "Workstation")


    def test_load_csv_parses_scan_ad_slash_dates(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "devices.csv"
            with path.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
                writer.writeheader()
                writer.writerow(
                    {
                        "ComputerType": "Workstation",
                        "CN": "PC01",
                        "DNSHostName": "pc01.example.local",
                        "OUPath": "example/Health",
                        "DistinguishedName": "CN=PC01,OU=Health,DC=example,DC=local",
                        "OperatingSystem": "Windows 11 Enterprise",
                        "LastLogonDate": "2026/06/01 10:06:51",
                        "LastSeenDate": "2026/06/01 10:06:51",
                    }
                )

            records = load_scan_records(path)
            self.assertEqual(records[0]["LastLogonDate"].date().isoformat(), "2026-06-01")
            self.assertEqual(records[0]["LastSeenDate"].date().isoformat(), "2026-06-01")

    def test_load_csv_parses_windows_filetime_last_logon_timestamp(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "devices.csv"
            with path.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.DictWriter(handle, fieldnames=FIELDNAMES)
                writer.writeheader()
                writer.writerow(
                    {
                        "ComputerType": "Server",
                        "CN": "SRV01",
                        "DNSHostName": "srv01.example.local",
                        "OUPath": "example/Health",
                        "DistinguishedName": "CN=SRV01,OU=Health,DC=example,DC=local",
                        "OperatingSystem": "Windows Server 2019",
                        "LastLogonTimestamp": "134220964118949845",
                    }
                )

            records = load_scan_records(path)
            self.assertEqual(records[0]["LastLogonTimestamp"].date().isoformat(), "2026-05-01")
    def test_load_json_records(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "devices.json"
            payload = [
                {
                    "ComputerType": "Server",
                    "CN": "SRV01",
                    "DNSHostName": "srv01.example.local",
                    "OUPath": "example/Infrastructure Development",
                    "DistinguishedName": "CN=SRV01,OU=Servers,DC=example,DC=local",
                    "OperatingSystem": "Windows Server 2022",
                }
            ]
            path.write_text(json.dumps(payload), encoding="utf-8")

            records = load_scan_records(path)
            validate_scan_records(records, str(path))
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]["ComputerType"], "Server")


if __name__ == "__main__":
    unittest.main()
