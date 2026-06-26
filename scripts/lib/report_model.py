from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
import re

from lib.department_mapping import DepartmentMatcher


SERVER_OS_BUCKETS: list[str] = [
    "Windows Server 2003 or Older",
    "Windows Server 2008",
    "Windows Server 2008 R2",
    "Windows Server 2012",
    "Windows Server 2012 R2",
    "Windows Server 2016",
    "Windows Server 2019",
    "Windows Server 2022",
    "Windows Server Newer/Unknown",
]

WORKSTATION_OS_BUCKETS: list[str] = [
    "Windows XP or Older",
    "Windows 7",
    "Windows 8/8.1",
    "Windows 10",
    "Windows 11",
    "Windows Workstation Other",
]

NON_WINDOWS_BUCKET = "Non-Windows / Unknown"
LEGACY_SERVER_BUCKETS: set[str] = {
    "Windows Server 2003 or Older",
    "Windows Server 2008",
    "Windows Server 2008 R2",
    "Windows Server 2012",
    "Windows Server 2012 R2",
}
LEGACY_WORKSTATION_BUCKETS: set[str] = {
    "Windows XP or Older",
    "Windows 7",
    "Windows 8/8.1",
    "Windows 10",
}
STALE_DEVICE_THRESHOLD_DAYS = 90

LEGACY_SERVER_SPOTLIGHT_BUCKETS: list[str] = [
    "Windows Server 2003 or Older",
    "Windows Server 2008",
    "Windows Server 2008 R2",
    "Windows Server 2012",
    "Windows Server 2012 R2",
]

DETAIL_COLUMNS: list[str] = [
    "ComputerType",
    "Department",
    "DepartmentMatchSource",
    "OSFamily",
    "OSNormalized",
    "OSLifecycleBucket",
    "OSIsLegacy",
    "InactivityDays",
    "InactivityStatus",
    "Name",
    "CN",
    "DNSHostName",
    "Description",
    "OUPath",
    "CanonicalName",
    "DistinguishedName",
    "Enabled",
    "OperatingSystem",
    "OperatingSystemVersion",
    "IPv4Address",
    "LastSeenDate",
    "DaysSinceLastSeen",
    "StaleStatus",
    "IsStale",
    "ObjectGUID",
    "SourceFile",
]


def _dedupe_key(record: dict[str, object]) -> str:
    for key in ("ObjectGUID", "DNSHostName", "CN", "Name"):
        value = record.get(key)
        if value:
            return str(value).strip().upper()
    return str(id(record))


def classify_record(record: dict[str, object], matcher: DepartmentMatcher) -> dict[str, object]:
    values = [
        ("OUPath", str(record.get("OUPath") or "")),
        ("DistinguishedName", str(record.get("DistinguishedName") or "")),
        ("DNSHostName", str(record.get("DNSHostName") or "")),
        ("CN", str(record.get("CN") or "")),
        ("Description", str(record.get("Description") or "")),
    ]
    match = matcher.match_values(values)
    record_copy = dict(record)
    if match:
        record_copy["Department"] = match.department
        record_copy["DepartmentMatchSource"] = match.matched_by
    else:
        record_copy["Department"] = "Unknown"
        record_copy["DepartmentMatchSource"] = "unmatched"
    return record_copy


def normalize_computer_type(value: object) -> str:
    text = str(value or "").strip().lower()
    if text == "server":
        return "Server"
    if text == "workstation":
        return "Workstation"
    if "server" in text:
        return "Server"
    return "Workstation"


def _parse_version_numbers(value: object) -> tuple[int, ...]:
    text = str(value or "").strip()
    matches = re.findall(r"\d+", text)
    return tuple(int(match) for match in matches)


def classify_operating_system(record: dict[str, object]) -> dict[str, object]:
    operating_system = str(record.get("OperatingSystem") or "").strip()
    operating_system_version = str(record.get("OperatingSystemVersion") or "").strip()
    combined = f"{operating_system} {operating_system_version}".lower()
    computer_type = normalize_computer_type(record.get("ComputerType"))
    version_numbers = _parse_version_numbers(operating_system_version)
    build_number = version_numbers[2] if len(version_numbers) >= 3 else None

    if "windows" not in combined and "server" not in combined:
        return {
            "OSFamily": NON_WINDOWS_BUCKET,
            "OSNormalized": NON_WINDOWS_BUCKET,
            "OSLifecycleBucket": NON_WINDOWS_BUCKET,
            "OSIsLegacy": False,
        }

    if computer_type == "Server":
        if "2003" in combined or "2000" in combined or "nt 4" in combined:
            bucket = "Windows Server 2003 or Older"
        elif "2008 r2" in combined:
            bucket = "Windows Server 2008 R2"
        elif "2008" in combined:
            bucket = "Windows Server 2008"
        elif "2012 r2" in combined:
            bucket = "Windows Server 2012 R2"
        elif "2012" in combined:
            bucket = "Windows Server 2012"
        elif "2016" in combined:
            bucket = "Windows Server 2016"
        elif "2019" in combined:
            bucket = "Windows Server 2019"
        elif "2022" in combined:
            bucket = "Windows Server 2022"
        elif operating_system_version.startswith("6.0"):
            bucket = "Windows Server 2008"
        elif operating_system_version.startswith("6.1"):
            bucket = "Windows Server 2008 R2"
        elif operating_system_version.startswith("6.2"):
            bucket = "Windows Server 2012"
        elif operating_system_version.startswith("6.3"):
            bucket = "Windows Server 2012 R2"
        elif operating_system_version.startswith("5."):
            bucket = "Windows Server 2003 or Older"
        else:
            bucket = "Windows Server Newer/Unknown"

        return {
            "OSFamily": "Windows Server",
            "OSNormalized": bucket,
            "OSLifecycleBucket": bucket,
            "OSIsLegacy": bucket in LEGACY_SERVER_BUCKETS,
        }

    if "xp" in combined or operating_system_version.startswith("5."):
        bucket = "Windows XP or Older"
    elif "windows 7" in combined or operating_system_version.startswith("6.1"):
        bucket = "Windows 7"
    elif "windows 8.1" in combined or operating_system_version.startswith("6.3"):
        bucket = "Windows 8/8.1"
    elif "windows 8" in combined or operating_system_version.startswith("6.2"):
        bucket = "Windows 8/8.1"
    elif "windows 11" in combined:
        bucket = "Windows 11"
    elif "windows 10" in combined:
        bucket = "Windows 10"
    elif operating_system_version.startswith("10.0") and build_number is not None:
        bucket = "Windows 11" if build_number >= 22000 else "Windows 10"
    elif "windows" in combined:
        bucket = "Windows Workstation Other"
    else:
        bucket = NON_WINDOWS_BUCKET

    return {
        "OSFamily": "Windows Workstation" if bucket != NON_WINDOWS_BUCKET else NON_WINDOWS_BUCKET,
        "OSNormalized": bucket,
        "OSLifecycleBucket": bucket,
        "OSIsLegacy": bucket in LEGACY_WORKSTATION_BUCKETS,
    }



def _parse_int(value: object) -> int | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value if value >= 0 else None
    text = str(value).strip()
    if not text:
        return None
    try:
        parsed = int(float(text))
    except ValueError:
        return None
    return parsed if parsed >= 0 else None


def _parse_date(value: object) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None

    for candidate in (text, text.replace("Z", "+00:00")):
        try:
            return datetime.fromisoformat(candidate).date()
        except ValueError:
            pass

    for date_format in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text.split()[0], date_format).date()
        except ValueError:
            pass

    return None


def _normalized_last_seen_date(record: dict[str, object]) -> date | None:
    for field_name in ("LastSeenDate", "LastLogonDate", "LastLogonTimestamp"):
        record_date = _parse_date(record.get(field_name))
        if record_date is not None:
            return record_date
    return None


def _days_since_record_date(record: dict[str, object], as_of_date: date) -> int | None:
    record_date = _normalized_last_seen_date(record)
    if record_date is None:
        return None
    elapsed_days = (as_of_date - record_date).days
    return elapsed_days if elapsed_days >= 0 else None


def _inactivity_days(record: dict[str, object], as_of_date: date) -> int | None:
    source_days = _parse_int(record.get("DaysSinceLastSeen"))
    if source_days is not None:
        return source_days
    return _days_since_record_date(record, as_of_date)


def _inactivity_status(record: dict[str, object], inactivity_days: int | None) -> str:
    if str(record.get("StaleStatus") or "").strip().lower() == "stale":
        return "Stale"
    if inactivity_days is None:
        return "Unknown"
    if inactivity_days >= STALE_DEVICE_THRESHOLD_DAYS:
        return "Stale"
    return "Fresh"


def _record_sort_key(record: dict[str, object]) -> tuple[object, ...]:
    return (
        0 if record.get("OSIsLegacy") else 1,
        str(record.get("OSLifecycleBucket") or ""),
        str(record.get("DNSHostName") or record.get("CN") or record.get("Name") or ""),
    )


@dataclass
class ReportBundle:
    records: list[dict[str, object]]
    departments: list[str]
    as_of_date: date

    @property
    def departments_with_unknown(self) -> list[str]:
        baseline = list(self.departments)
        if "Unknown" not in baseline:
            baseline.append("Unknown")
        if any(record["Department"] == "Domain Controllers" for record in self.records):
            if "Domain Controllers" not in baseline:
                baseline.append("Domain Controllers")
        return baseline

    def records_for_type(self, computer_type: str) -> list[dict[str, object]]:
        matching = [record for record in self.records if record["ComputerType"] == computer_type]
        return sorted(matching, key=_record_sort_key)

    def records_for_department(self, department: str) -> list[dict[str, object]]:
        matching = [record for record in self.records if record["Department"] == department]
        return sorted(matching, key=_record_sort_key)

    def dashboard_rows(self) -> list[dict[str, object]]:
        rows: list[dict[str, object]] = []
        for department in self.departments_with_unknown:
            department_records = self.records_for_department(department)
            server_count = sum(1 for record in department_records if record["ComputerType"] == "Server")
            workstation_count = sum(
                1 for record in department_records if record["ComputerType"] == "Workstation"
            )
            enabled_count = sum(1 for record in department_records if record.get("Enabled") is True)
            disabled_count = sum(1 for record in department_records if record.get("Enabled") is False)
            rows.append(
                {
                    "Department": department,
                    "Servers": server_count,
                    "Workstations": workstation_count,
                    "Total": server_count + workstation_count,
                    "Enabled": enabled_count,
                    "Disabled": disabled_count,
                }
            )
        return rows

    def os_summary_rows(self, computer_type: str) -> list[dict[str, object]]:
        buckets = SERVER_OS_BUCKETS if computer_type == "Server" else WORKSTATION_OS_BUCKETS
        rows: list[dict[str, object]] = []
        typed_records = self.records_for_type(computer_type)
        for bucket in buckets:
            bucket_records = [record for record in typed_records if record.get("OSLifecycleBucket") == bucket]
            rows.append(
                {
                    "OSLifecycleBucket": bucket,
                    "Count": len(bucket_records),
                    "Legacy": "Yes" if any(record.get("OSIsLegacy") for record in bucket_records) else "No",
                }
            )

        non_windows_count = sum(
            1 for record in typed_records if record.get("OSLifecycleBucket") == NON_WINDOWS_BUCKET
        )
        if non_windows_count > 0:
            rows.append(
                {
                    "OSLifecycleBucket": NON_WINDOWS_BUCKET,
                    "Count": non_windows_count,
                    "Legacy": "No",
                }
            )
        return rows

    def legacy_department_rows(self) -> list[dict[str, object]]:
        rows: list[dict[str, object]] = []
        for department in self.departments_with_unknown:
            department_records = self.records_for_department(department)
            legacy_servers = sum(
                1
                for record in department_records
                if record["ComputerType"] == "Server" and record.get("OSIsLegacy")
            )
            legacy_workstations = sum(
                1
                for record in department_records
                if record["ComputerType"] == "Workstation" and record.get("OSIsLegacy")
            )
            rows.append(
                {
                    "Department": department,
                    "LegacyServers": legacy_servers,
                    "LegacyWorkstations": legacy_workstations,
                    "LegacyTotal": legacy_servers + legacy_workstations,
                }
            )
        return rows

    def legacy_spotlight_rows(self, department: str | None = None) -> list[dict[str, object]]:
        records = self.records if department is None else self.records_for_department(department)
        rows: list[dict[str, object]] = []
        for bucket in LEGACY_SERVER_SPOTLIGHT_BUCKETS:
            count = sum(
                1
                for record in records
                if record["ComputerType"] == "Server" and record.get("OSLifecycleBucket") == bucket
            )
            rows.append({"OSLifecycleBucket": bucket, "Count": count})
        return rows

    def department_summary(self, department: str) -> dict[str, object]:
        department_records = self.records_for_department(department)
        server_records = [record for record in department_records if record["ComputerType"] == "Server"]
        workstation_records = [
            record for record in department_records if record["ComputerType"] == "Workstation"
        ]

        summary = {
            "Servers": len(server_records),
            "Workstations": len(workstation_records),
            "Total": len(department_records),
            "LegacyServers": sum(1 for record in server_records if record.get("OSIsLegacy")),
            "LegacyWorkstations": sum(
                1 for record in workstation_records if record.get("OSIsLegacy")
            ),
            "ServerOSSummary": self._records_os_summary(server_records, SERVER_OS_BUCKETS),
            "WorkstationOSSummary": self._records_os_summary(
                workstation_records,
                WORKSTATION_OS_BUCKETS,
            ),
            "LegacySpotlight": self.legacy_spotlight_rows(department=department),
        }
        return summary

    @staticmethod
    def _records_os_summary(
        records: list[dict[str, object]],
        buckets: list[str],
    ) -> list[dict[str, object]]:
        rows: list[dict[str, object]] = []
        for bucket in buckets:
            bucket_records = [record for record in records if record.get("OSLifecycleBucket") == bucket]
            rows.append(
                {
                    "OSLifecycleBucket": bucket,
                    "Count": len(bucket_records),
                    "Legacy": "Yes" if any(record.get("OSIsLegacy") for record in bucket_records) else "No",
                }
            )
        return rows


def build_report_bundle(
    records: list[dict[str, object]],
    matcher: DepartmentMatcher,
    departments: list[str],
    as_of_date: date | None = None,
) -> ReportBundle:
    effective_as_of_date = as_of_date or date.today()
    deduped: dict[str, dict[str, object]] = {}
    for record in records:
        typed_record = dict(record)
        typed_record["ComputerType"] = normalize_computer_type(typed_record.get("ComputerType"))
        typed_record.update(classify_operating_system(typed_record))
        typed_record["LastSeenDate"] = _normalized_last_seen_date(typed_record)
        inactivity_days = _inactivity_days(typed_record, effective_as_of_date)
        typed_record["InactivityDays"] = inactivity_days
        typed_record["InactivityStatus"] = _inactivity_status(typed_record, inactivity_days)
        typed_record = classify_record(typed_record, matcher)
        deduped[_dedupe_key(typed_record)] = typed_record

    sorted_records = sorted(
        deduped.values(),
        key=lambda record: (
            str(record.get("Department") or ""),
            str(record.get("ComputerType") or ""),
            *_record_sort_key(record),
        ),
    )
    return ReportBundle(records=sorted_records, departments=departments, as_of_date=effective_as_of_date)


def serialize_for_sheet(records: list[dict[str, object]]) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for record in sorted(records, key=_record_sort_key):
        row: dict[str, object] = {}
        for column in DETAIL_COLUMNS:
            row[column] = record.get(column)
        rows.append(row)
    return rows


def unmatched_records(records: list[dict[str, object]]) -> list[dict[str, object]]:
    return [record for record in records if record.get("Department") == "Unknown"]
