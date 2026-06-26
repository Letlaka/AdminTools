# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Repository-standard documentation files:
  - `CONTRIBUTING.md`
  - `SECURITY.md`
  - `CHANGELOG.md`
- New focused documentation under `docs/`:
  - `overview.md`
  - `usage.md`
  - `parameters.md`
  - `outputs.md`
  - `examples.md`
  - `troubleshooting.md`
- Root `README.md` converted to a documentation entry point.
- AD Excel reporting workflow documentation under `docs/ad-excel-reporting.md`.

### Fixed

- Full AD scans now log query progress while retrieving computer objects.
- The base AD inventory query no longer requests the extended `ipv4Address` property, which could make broad workstation scans appear to hang.
- Empty-result `Scan-ADComputers.ps1` runs now complete cleanly without null-binding export and summary errors.

### Added

- Python Excel reporting pipeline under `scripts/build_ad_excel_reports.py` for department dashboards and per-department workbooks from `Scan-ADComputers` exports.
- Windows lifecycle tracking in Excel reports with normalized server and workstation OS buckets, legacy highlighting, and legacy server spotlight summaries.
- Financial-year based report filing structure with consolidated, department, source, and log folders.
- Shared Python helpers for department matching, Scan-ADComputers input normalization, financial year calculation, OS normalization, and Excel workbook generation.
