# Changelog

All notable changes to **VBA Excel UI** are documented in this file.

The project follows [Semantic Versioning](https://semver.org/) for its public
VBA API. Dates use the `YYYY-MM-DD` format.

## [Unreleased]

No unreleased changes are currently documented.

## [1.0.1] - 2026-07-25

### Added

- Added `CONTRIBUTING.md` with repository-specific contribution, testing,
  compatibility, WinAPI, and release guidance.
- Added `CODE_OF_CONDUCT.md`.
- Added `SECURITY.md` with supported-version and private-reporting guidance.
- Added a repository-specific `.gitignore`.
- Added `.gitattributes` to preserve deterministic text handling for exported
  VBA modules and binary handling for Excel workbooks.
- Added GitHub issue templates for bug reports and feature requests.
- Added the GitHub issue-template chooser configuration.
- Added a pull-request template tailored to Excel UI, snapshot, WinAPI,
  diagnostics, recovery, and compatibility changes.

### Changed

- Redesigned `README.md` as the primary project, API, architecture, integration,
  testing, recovery, and release reference.
- Updated the core, demo, and regression-test module metadata to version
  `1.0.1`.
- Updated module documentation dates to `2026-07-25`.
- Reduced repetitive comments while preserving section banners, procedure
  headers, and declaration-level inline comments.
- Corrected stale or imprecise documentation concerning:
  - `UI_ShowExcelUI` versus explicit snapshot restoration;
  - fire-and-forget versus structured-result behavior;
  - Ribbon and title-bar state handling;
  - index-based per-window snapshot restoration;
  - title-bar style restoration limitations;
  - demo-module dependencies and button assignments.
- Synchronized `demo/EXCEL_UI_DEMO.xlsm` with the exported versioned VBA
  modules.

### Validation

The release candidate was validated manually in desktop Microsoft Excel:

- `Debug -> Compile VBAProject`
- `Test_EXCEL_UI_RunCore`
- `Test_EXCEL_UI_RunTitleBarOnly`
- `Test_EXCEL_UI_RunAll`

All three regression runners completed successfully.

### Compatibility

- No public procedure signature changed.
- No public enum member or value changed.
- No migration is required for existing callers.
- No executable VBA behavior was intentionally changed.
- GitHub Actions workflows are intentionally not included in this release.

[Unreleased]: https://github.com/danielep71/VBA-EXCEL_UI/compare/v1.0.1...HEAD
[1.0.1]: https://github.com/danielep71/VBA-EXCEL_UI/releases/tag/v1.0.1
