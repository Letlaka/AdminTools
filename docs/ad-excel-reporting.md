# AD Excel Reporting

`build_ad_excel_reports.py` converts `Scan-ADComputers.ps1` CSV or JSON exports
into Excel workbooks organized by financial year and department.

## Current Reporting Features

- consolidated workbook with:
  - `Dashboard`
  - `All Servers`
  - `All Workstations`
  - one sheet per department
- one department workbook per department with:
  - `Summary`
  - `Servers`
  - `Workstations`
- department classification using:
  - `config/dept_list.txt`
  - `config/dept_codes.txt`

The runtime config files are intentionally environment-local and ignored by git
so site-specific department names and codes do not get committed. Public sample
files are committed under `config/*.sample.*`; copy them to the runtime names
before running the reporting script:

```powershell
Copy-Item config/dept_list.sample.txt config/dept_list.txt
Copy-Item config/dept_codes.sample.txt config/dept_codes.txt
```

## Windows Version Tracking

The reporting pipeline derives these normalized OS fields from
`OperatingSystem` and `OperatingSystemVersion`:

- `OSFamily`
- `OSNormalized`
- `OSLifecycleBucket`
- `OSIsLegacy`
- `LastSeenDate`
- `InactivityDays`
- `InactivityStatus`

Server lifecycle buckets:

- `Windows Server 2003 or Older`
- `Windows Server 2008`
- `Windows Server 2008 R2`
- `Windows Server 2012`
- `Windows Server 2012 R2`
- `Windows Server 2016`
- `Windows Server 2019`
- `Windows Server 2022`
- `Windows Server Newer/Unknown`

Workstation lifecycle buckets:

- `Windows XP or Older`
- `Windows 7`
- `Windows 8/8.1`
- `Windows 10`
- `Windows 11`
- `Windows Workstation Other`
- `Non-Windows / Unknown`

Legacy defaults:

- servers: `2012 R2` and older
- workstations: `Windows 8.1` and older

## Dashboard Additions

The consolidated workbook dashboard places each chart beside its source table, using table-width-based chart anchors and enough row spacing to keep charts from overlapping the next section. Sections appear in this order:

- department device summary
- legacy device summary by department
- legacy server spotlight for:
  - `2003 or Older`
  - `2008`
  - `2008 R2`
  - `2012`
  - `2012 R2`
- server OS lifecycle summary
- workstation OS lifecycle summary

## Detail Sheet Additions

All detail sheets include:

- `OSFamily`
- `OSNormalized`
- `OSLifecycleBucket`
- `OSIsLegacy`
- `LastSeenDate`
- `InactivityDays`
- `InactivityStatus`

`LastSeenDate`, `InactivityDays`, and `InactivityStatus` are normalized reporting fields. `LastSeenDate` preserves a valid source `LastSeenDate` when present; otherwise it is filled from source `LastLogonDate`, with `LastLogonTimestamp` as a final fallback. `InactivityDays` prefers source `DaysSinceLastSeen` when present, otherwise it is calculated from the normalized `LastSeenDate`. `InactivityStatus` is `Stale` at 90 days or more, `Fresh` below 90 days, and `Unknown` when no usable source exists. Source `StaleStatus = Stale` still takes priority for the status, but raw source fields such as `DaysSinceLastSeen`, `StaleStatus`, and `IsStale` remain later in the detail sheets for traceability.

Legacy devices are sorted to the top of detail sheets and highlighted in the
generated workbooks.

## Selective Report Generation

The reporting utility can generate all workbooks, only the consolidated workbook,
all department workbooks, or one named department workbook.

Generate only the consolidated dashboard workbook:

```powershell
uv run python scripts/build_ad_excel_reports.py `
  --workstations reports/ad-computers/Workstations_example_corp_local_20260629.csv `
  --report-scope main `
  --as-of-date 2026-06-29
```

Generate all department workbooks without the consolidated dashboard:

```powershell
uv run python scripts/build_ad_excel_reports.py `
  --workstations reports/ad-computers/Workstations_example_corp_local_20260629.csv `
  --report-scope departments `
  --as-of-date 2026-06-29
```

Generate one department workbook only:

```powershell
uv run python scripts/build_ad_excel_reports.py `
  --workstations reports/ad-computers/Workstations_example_corp_local_20260629.csv `
  --report-scope department `
  --department "GPtransport" `
  --as-of-date 2026-06-29
```

`--department` is valid only with `--report-scope department`. The department
name is matched case-insensitively against the generated department names.

## Dependency

Install the Excel dependency:

```powershell
uv sync
```

If you are not using `uv`, install from `requirements-reporting.txt` into your
chosen Python 3.13+ environment.

## Department Config Format

Start from `config/dept_list.sample.txt`, then save the local runtime file as
`config/dept_list.txt`. It contains one department name per line:

```text
Finance
Human Resources
Information Technology
```

Start from `config/dept_codes.sample.txt`, then save the local runtime file as
`config/dept_codes.txt`. It maps department codes to department names:

```text
FIN=Finance
HR=Human Resources
IT=Information Technology
```

## Example

```powershell
uv run python scripts/build_ad_excel_reports.py `
  --servers reports/ad-computers/Servers_example_corp_local_20260625.csv `
  --workstations reports/ad-computers/Workstations_example_corp_local_20260625.csv `
  --as-of-date 2026-06-25
```

By default, Excel output is written under:

```text
reports/<financial-year>/<run-date>/
```
