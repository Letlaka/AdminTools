from __future__ import annotations

import csv
import json
from datetime import datetime, timedelta
from pathlib import Path


SCAN_AD_FIELD_MAP: dict[str, str] = {
    "ComputerType": "ComputerType",
    "Name": "Name",
    "CN": "CN",
    "DNSHostName": "DNSHostName",
    "Description": "Description",
    "OUPath": "OUPath",
    "CanonicalName": "CanonicalName",
    "DistinguishedName": "DistinguishedName",
    "Created": "Created",
    "Enabled": "Enabled",
    "IPv4Address": "IPv4Address",
    "LastLogonDate": "LastLogonDate",
    "LastLogonTimestamp": "LastLogonTimestamp",
    "LastSeenDate": "LastSeenDate",
    "DaysSinceLastSeen": "DaysSinceLastSeen",
    "InactiveThresholdDays": "InactiveThresholdDays",
    "IsStale": "IsStale",
    "StaleStatus": "StaleStatus",
    "ObjectGUID": "ObjectGUID",
    "OperatingSystem": "OperatingSystem",
    "OperatingSystemVersion": "OperatingSystemVersion",
}


def _parse_bool(value: object) -> bool | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"true", "1", "yes", "y"}:
        return True
    if text in {"false", "0", "no", "n"}:
        return False
    return None


def _parse_int(value: object) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(str(value).strip())
    except ValueError:
        return None


def _parse_windows_filetime(value: str) -> datetime | None:
    try:
        ticks = int(value)
    except ValueError:
        return None
    if ticks <= 0:
        return None
    try:
        return datetime(1601, 1, 1) + timedelta(microseconds=ticks / 10)
    except OverflowError:
        return None


def _parse_datetime(value: object) -> datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value

    text = str(value).strip()
    for parser in (
        datetime.fromisoformat,
        lambda raw: datetime.strptime(raw, "%m/%d/%Y %H:%M:%S"),
        lambda raw: datetime.strptime(raw, "%m/%d/%Y %I:%M:%S %p"),
        lambda raw: datetime.strptime(raw, "%Y-%m-%d %H:%M:%S"),
        lambda raw: datetime.strptime(raw, "%Y-%m-%d"),
        lambda raw: datetime.strptime(raw, "%Y/%m/%d %H:%M:%S"),
        lambda raw: datetime.strptime(raw, "%Y/%m/%d"),
        _parse_windows_filetime,
    ):
        try:
            parsed = parser(text)
        except ValueError:
            continue
        if parsed is not None:
            return parsed
    return None


def _normalize_record(
    row: dict[str, object],
    source_path: Path,
    forced_computer_type: str | None = None,
) -> dict[str, object]:
    normalized: dict[str, object] = {}
    for source_key, target_key in SCAN_AD_FIELD_MAP.items():
        normalized[target_key] = row.get(source_key)

    if forced_computer_type:
        normalized["ComputerType"] = forced_computer_type

    normalized["Enabled"] = _parse_bool(normalized.get("Enabled"))
    normalized["IsStale"] = _parse_bool(normalized.get("IsStale"))
    normalized["DaysSinceLastSeen"] = _parse_int(normalized.get("DaysSinceLastSeen"))
    normalized["InactiveThresholdDays"] = _parse_int(normalized.get("InactiveThresholdDays"))
    normalized["Created"] = _parse_datetime(normalized.get("Created"))
    normalized["LastLogonDate"] = _parse_datetime(normalized.get("LastLogonDate"))
    normalized["LastLogonTimestamp"] = _parse_datetime(normalized.get("LastLogonTimestamp"))
    normalized["LastSeenDate"] = _parse_datetime(normalized.get("LastSeenDate"))
    normalized["SourceFile"] = str(source_path)
    return normalized


def _load_csv(path: Path, forced_computer_type: str | None = None) -> list[dict[str, object]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = [row for row in reader]

    return [_normalize_record(row, path, forced_computer_type=forced_computer_type) for row in rows]


def _load_json(path: Path, forced_computer_type: str | None = None) -> list[dict[str, object]]:
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    if isinstance(payload, dict):
        rows = [payload]
    elif isinstance(payload, list):
        rows = payload
    else:
        raise ValueError(f"Unsupported JSON structure in {path}")

    return [_normalize_record(row, path, forced_computer_type=forced_computer_type) for row in rows]


def load_scan_records(path: Path, forced_computer_type: str | None = None) -> list[dict[str, object]]:
    if not path.exists():
        raise FileNotFoundError(path)

    suffix = path.suffix.lower()
    if suffix == ".csv":
        return _load_csv(path, forced_computer_type=forced_computer_type)
    if suffix == ".json":
        return _load_json(path, forced_computer_type=forced_computer_type)

    raise ValueError(f"Unsupported input format for {path}. Expected .csv or .json")


def validate_scan_records(records: list[dict[str, object]], source_label: str) -> None:
    required_keys = {"CN", "DNSHostName", "DistinguishedName", "OUPath", "OperatingSystem"}
    missing = [
        key
        for key in required_keys
        if not any(record.get(key) not in (None, "") for record in records)
    ]
    if missing and records:
        raise ValueError(
            f"Input '{source_label}' does not look like a Scan-ADComputers export. "
            f"Missing usable fields: {', '.join(sorted(missing))}"
        )
