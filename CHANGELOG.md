# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-06-29

### Added

- Added `Scan-ADComputers.ps1` for Active Directory computer inventory, validation, stale-device reporting, targeted scans, operational checks, and CSV, JSON, or HTML exports.
- Added `Get-ADAdminActivity.ps1` for Domain Controller Security log reporting of Active Directory administrative activity, with optional privileged-admin filtering.
- Added `Manage-ADUserAccounts.ps1` for user account reports, locked-out account detail, single-user audit lookup, and explicit unlock, enable, and password-reset actions.
- Added `AdminToolsCommon.psm1` with shared helpers for logging, path validation, credential resolution, CSV sanitization, output safety, and retry behavior.
- Added reusable credential support through explicit credentials, Microsoft SecretManagement secret names, and DPAPI-protected CLIXML credential files.
- Added safe output defaults, including repository-local report folders, central log folders, no silent overwrites, CSV formula sanitization, and guarded network input/output paths.
- Added `-NoProgress` support for unattended computer scans.
- Added JSON config support for `Scan-ADComputers.ps1` with committed sample server and workstation config files.
- Added sample configuration files for server lists, workstation lists, department names, and department code mappings.
- Added the Python Excel reporting pipeline under `scripts/build_ad_excel_reports.py` for consolidated dashboards and per-department workbooks from `Scan-ADComputers` exports.
- Added shared Python reporting helpers for Scan-ADComputers input normalization, department matching, financial-year path calculation, report modelling, OS lifecycle classification, and workbook generation.
- Added Windows lifecycle tracking in Excel reports, including normalized server and workstation OS buckets, legacy highlighting, and legacy server spotlight summaries.
- Added normalized reporting fields for Excel detail sheets, including `LastSeenDate`, `InactivityDays`, and `InactivityStatus`.
- Added financial-year based Excel report output structure with consolidated, department, source, and log folders.
- Added selective Excel report generation for consolidated-only, all-departments-only, and single-department output.
- Added focused documentation under `docs/` for overview, usage, parameters, outputs, examples, troubleshooting, computer scanning, admin activity reporting, user account management, and AD Excel reporting.
- Added repository-standard documentation files: `CONTRIBUTING.md`, `SECURITY.md`, and `CHANGELOG.md`.
- Added manual generation tooling in `tools/New-AdminToolsManual.ps1`.
- Added PowerShell Pester tests and Python unit tests covering shared helpers, report generation, input normalization, department matching, workbook output, and release-critical behavior.

### Changed

- Converted the root `README.md` into a documentation entry point with quick-start commands and links to the detailed manuals.
- Refactored script structure and function names for readability, consistency, and shared helper reuse.
- Standardized logging across the PowerShell scripts and documented default log paths.
- Enhanced Pester test configuration and updated tests for the current script behavior.
- Updated documentation to describe environment-specific setup, sample config usage, log paths, safety controls, and report outputs.

### Fixed

- Full AD scans now log query progress while retrieving computer objects.
- The base AD inventory query no longer requests the extended `ipv4Address` property, which could make broad workstation scans appear to hang.
- Empty-result `Scan-ADComputers.ps1` runs now complete cleanly without null-binding export and summary errors.
- Addressed PowerShell ScriptAnalyzer warnings around password handling, helper naming, and performance-stage functions.
- Fixed department matching edge cases for full department phrases, aliases, quoted values, and code-prefix routing in Excel reports.
- Fixed Excel report date normalization so `LastSeenDate` can be populated consistently from source `LastSeenDate`, `LastLogonDate`, or `LastLogonTimestamp`.
