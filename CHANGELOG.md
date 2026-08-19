# Changelog

All notable changes to **VBA Excel UI** are documented in this file.

The project follows [Semantic Versioning](https://semver.org/) for its public
VBA API. Dates use the `YYYY-MM-DD` format.

## [Unreleased]

### Planned for 1.1.1

A corrective release addressing the findings of the independent `v1.1.0` review
recorded in `docs/INDEPENDENT_CODE_REVIEW_V1.1.0_2026-08-19.md`. The public API
is unchanged: no procedure, enum or parameter is added, removed or renamed, and
no existing call site requires modification.

### Added

- Added explicit-target title-bar entry points to `M_EXCEL_UI_TITLEBAR`:
  `UI_TryGetActiveTitleBarHwnd`, `UI_TryGetTitleBarVisibleForHwnd`,
  `UI_TrySetTitleBarVisibleForHwndIfNeeded` and
  `UI_InternalIsTitleBarFrameAlive`. Callers that must read a frame now and
  write it back later can resolve the target window once and keep it, rather
  than re-resolving `Application.Hwnd` at each end of the operation.
- Added `UI_SnapshotTryGetActiveWindow` and
  `UI_SnapshotTryResolveTitleBarFrame` to `M_EXCEL_UI_SNAPSHOT`, which capture
  the identity of the top-level frame a title-bar value was read from and prove
  that frame is still present before anything is written back to it.
- Added `UI_RuntimeTryAppendFailureEntry` and
  `UI_RuntimeMarkFailureListTruncated` to `M_EXCEL_UI_RUNTIME`, which separate
  the fallible allocation from the infallible status update, and record a
  truncation marker in a slot that already exists when the list cannot grow.
- Added the regression seams `UI_InternalResetTitleBarBaselineForHwnd`,
  `UI_InternalIsFrameRefreshPending`, `UI_InternalInjectFrameRefreshFailure`
  and `UI_InternalInjectFailureListGrowthFailure`. They exist because neither a
  `SetWindowPos` failure nor an exhausted allocation can be produced on demand,
  and a recovery path that cannot be executed is indistinguishable from one
  that was never written.
- Added `Test_EXCEL_UI_RunTitleBarSdiIdentity`, a regression runner covering
  title-bar restoration across two workbook windows. It verifies that a
  snapshot restores the frame it was captured from while a different window is
  active, and that a captured frame which has since closed is reported rather
  than redirected. The runner is destructive and is invoked explicitly; it is
  not part of `Test_EXCEL_UI_RunAll`.
- Added `TST_Case_TitleBarFrameRefreshDebtRetried` to the title-bar regression
  pack, which injects a frame-refresh failure and verifies that the outstanding
  repaint is recorded and retried on the next call rather than short-circuited
  as a no-op.
- Added `TST_Case_FailureAccumulatorDegradesSafely` to the core regression
  pack, which injects a failure-list growth failure and verifies that the
  status outputs survive, the truncation is reported, and nothing raises.
- Added `Test_EXCEL_UI_RunReleaseCertification`, a single runner that executes
  every mandatory regression unit, counts units, failures, skips and cleanup
  separately, verifies the host state afterwards rather than assuming it, and
  emits a JSON evidence document and a text report naming the exact Excel
  build, bitness and operating system the verdict was obtained on. It refuses
  to start when an explicit snapshot already exists, rather than degrading into
  a partial run that reads like a complete one.

### Changed

- Replaced the single process-wide title-bar baseline with a frame-state
  registry keyed by top-level window handle. Operating on one workbook window
  no longer discards the baseline captured for another, and entries whose
  window has closed are reclaimed before the registry grows.
- The title-bar baseline is now refreshed rather than captured once. While the
  component does not own a hidden state for a window, the live owned style bits
  are re-adopted on every call, so a legitimate frame change made by Excel or
  another add-in survives the next hide and show instead of being reverted to
  bits captured earlier in the session.
- The snapshot now retains the top-level window handle, the owning Excel
  `Window` object and a diagnostic label for the captured title bar, and
  restores through them. A captured frame that is no longer open is reported as
  a title-bar failure naming the window, instead of the captured value being
  applied to whichever workbook window is active at restore time.
- `FailureCount` is now documented as authoritative and `FailureList` as best
  effort. The list can hold fewer entries than the count when an allocation
  fails, but never silently: a `Diagnostics` truncation marker is written into
  an existing slot whenever growth failed.
- A skipped regression case is now counted rather than only logged. Under the
  certification runner a skipped mandatory case makes the run `INCOMPLETE` and
  therefore not a pass; the legacy runners keep their previous behavior
  exactly, because the accounting is inert outside a certification run.
- `UI_TryGetTitleBarVisible` and `UI_TrySetTitleBarVisibleIfNeeded` are now
  documented and implemented as active-window wrappers over the explicit-target
  entry points. Their signatures and behavior for existing callers are
  unchanged.

### Fixed

- Fixed title-bar snapshot restoration not being identity-safe under the Single
  Document Interface. `Application.Hwnd` reports the active workbook window's
  handle, and the snapshot re-resolved it on restore, so activating a different
  workbook between capture and restore applied one window's captured title-bar
  state to another. Every API call succeeded, so the misdirection was silent and
  the originally captured frame was left unrestored.
  (`ICR-UI-P1-01`, #14)
- Fixed the title-bar owned-bit baseline being a single process-wide value that
  a second workbook window silently displaced, and that was never refreshed
  after another component legitimately changed the owned frame bits.
  (`ICR-UI-P2-04`, #15)
- Fixed a title-bar style write and its non-client frame refresh not being
  treated as one unit of work. `SetWindowLong` could succeed while
  `SetWindowPos` failed, after which the desired style already matched the
  current style and the next call short-circuited, reporting success over a
  frame Windows had never re-measured. The outstanding refresh is now recorded
  against the window and retried before the no-op test.
  (`ICR-UI-P2-03`, #16)
- Fixed failure accumulation being able to raise from inside an error handler.
  `UI_RuntimeAddFailure` grew the failure list with no error boundary and
  assumed the buffer already held a `String` array whose bound agreed with the
  count. An allocation failure, or a buffer holding anything else, replaced the
  original failure with the failure to record it, could abort a pass designed
  to continue, and could bypass the `ScreenUpdating` restoration in
  `UI_RuntimeEndQuietUpdate`. The status outputs are now set before anything
  fallible is attempted, the entry text degrades rather than failing, and the
  allocation is isolated behind a Boolean contract.
  (`ICR-UI-P2-02`, #17)

- Fixed the regression harness being unable to distinguish a complete pass from
  a partial one. `Test_EXCEL_UI_RunAll` executed no multi-window case, could
  skip the snapshot cases silently when a snapshot already existed, suppressed
  cleanup failures, and reported its outcome only as Immediate Window prose with
  no counters. A green result therefore carried far less information than it
  appeared to. `Test_EXCEL_UI_RunAll` is unchanged and remains the interactive
  runner; release certification now has its own gate.
  (`ICR-UI-P2-07`, #18)

### Documentation

- Added `docs/INDEPENDENT_CODE_REVIEW_V1.1.0_2026-08-19.md`, the independent
  code and repository review of the `v1.1.0` tag at commit
  `96360379a4bca7703cf649a69a2162961dfa6c9e`. Every issue in the `1.1.1`
  milestone cites it as a stable in-repo reference.

### Compatibility

```text
Existing calls affected: none
Backward compatible:     Yes
Release type:            patch
```

- No public procedure was added, removed or renamed.
- No existing parameter changed name, position, type or default.
- No enum member or value changed.
- The `Stage | Detail` diagnostic format is unchanged. Two new entries can now
  appear: a `TitleBar` entry reporting that the captured frame is no longer
  available, where the previous build silently applied the captured value to the
  active window; and a `Diagnostics` entry reporting that the failure list could
  not be grown.
- `FailureList` may now hold fewer entries than `FailureCount` under memory
  pressure. Callers that assumed the two always agreed should read the count as
  authoritative and treat the list as descriptive.
- Snapshot storage remains in memory only and does not survive a VBA project
  reset or an Excel restart. It now also retains one `Window` reference for the
  captured title-bar frame, released by `UI_ClearExcelUIStateSnapshot`, by a
  replacing capture, or by a project reset.

### Known limitations

- Ribbon scope under the Single Document Interface remains uncharacterized and
  is documented as application-level without supporting evidence. See `#21`.
- The `README.md` title-bar scope statement still describes the pre-fix
  behavior and is corrected with the rest of the documentation work. See `#19`.
- `tools/reformat.py` does not round-trip the committed production modules
  byte-for-byte, so a formatter check cannot yet be made blocking. See `#20`.

## [1.1.0] - 2026-08-19

Backward-compatible feature release. Every public `UI_...` procedure and enum
member from `1.0.1` is preserved, with unchanged names, parameter order and
defaults. No migration is required.

### Added

- Added `UIWindowTargetScope`, a public enum selecting which Excel windows
  receive window-level UI changes:
  - `UI_TargetAllExcelWindows` (0) - every current Excel window;
  - `UI_TargetActiveWindow` (1) - `Application.ActiveWindow` only;
  - `UI_TargetActiveWorkbookWindows` (2) - every window of the active workbook.
- Added an optional trailing `TargetScope` argument to `UI_SetExcelUI` and
  `UI_SetExcelUI_WithResult`, defaulting to `UI_TargetAllExcelWindows`.
  Targeting affects only Headings, Workbook Tabs and Gridlines; the Ribbon,
  Status Bar, Scroll Bars, Formula Bar and Title Bar keep their existing
  application-level and main-window scope.
- Added `UI_CaptureExcelUIState_WithResult`, returning the established
  `Boolean` + `FailureCount` + ordered `FailureList` contract for snapshot
  capture.
- Added `UI_ResetExcelUIToSnapshot_WithResult`, returning the same contract for
  snapshot restoration.
- Added `INSTALLATION.md`, documenting the four-module production package,
  import order, dependency graph, fresh installation, upgrade from the
  single-module architecture, upgrade from intermediate `1.1.0` builds,
  validation and troubleshooting.
- Added a snapshot-lifetime section to `INSTALLATION.md`, documenting that a
  captured snapshot retains one live `Window` reference per captured window,
  that those references are released only by `UI_ClearExcelUIStateSnapshot`, a
  replacing capture or a project reset, and that restoring deliberately retains
  the snapshot rather than releasing it.
- Added `tools/reformat.py`, a deterministic house-style reformatter for
  exported `.bas` modules.
- Added regression coverage for identity-safe window restoration, structured
  snapshot capture and restore results, title-bar owned-bit preservation,
  active-window targeting, active-workbook-window targeting, invalid target
  scopes, title-bar show recovery without a captured baseline, and per-element
  application-level capture and restoration.
- Added the `Test_EXCEL_UI_RunSnapshotIdentity` regression runner.

### Changed

- Replaced index-based per-window snapshot restoration with identity-based
  matching. The snapshot now retains each captured `Window` object and restores
  through that reference, so reordered windows restore correctly, windows
  opened after capture are left unchanged, and a closed or recreated window is
  reported rather than having its captured state applied to whichever window
  now occupies the same collection index. Diagnostic captions are stored
  separately and never participate in matching.
- Replaced whole-value title-bar style restoration with an owned-bit merge.
  `TITLEBAR_OWNED_STYLE_MASK` (`&HCF0000`) defines the exact bits this
  component claims - `WS_CAPTION`, `WS_SYSMENU`, `WS_THICKFRAME`,
  `WS_MINIMIZEBOX` and `WS_MAXIMIZEBOX` - and every write merges only those
  into the live style, preserving unrelated changes made by Excel or another
  component after capture.
- Decomposed the production implementation into four cohesive modules while
  keeping `M_EXCEL_UI` as the public facade:
  - `M_EXCEL_UI` - public API, visibility validation, apply orchestration;
  - `M_EXCEL_UI_RUNTIME` - shared fail-soft host operations, result buffers,
    diagnostics, quiet-update scope;
  - `M_EXCEL_UI_SNAPSHOT` - snapshot state, retained window identities, capture
    and restoration;
  - `M_EXCEL_UI_TITLEBAR` - WinAPI declarations, owned style bits, frame
    refresh.

  The dependency graph is acyclic. `M_EXCEL_UI_RUNTIME` and
  `M_EXCEL_UI_TITLEBAR` have no project-module dependency. All internal modules
  use `Option Explicit` and `Option Private Module`.
- Reformatted every `.bas` module to the project house style. Verified
  behaviour-neutral: 3,648 logical statements across the seven modules,
  statement-for-statement identical to their predecessors.
- Stopped version-controlling `demo/EXCEL_UI_DEMO.xlsm`. Tested macro-enabled
  demo workbooks are now distributed as GitHub Release assets only; the demo
  source modules remain in the repository.
- Updated `README.md` for the modular architecture, targeting scopes, snapshot
  identity model, structured diagnostics, installation and release checklist.
- Updated `CONTRIBUTING.md` and the pull-request template for the four-module
  package.
- Moved the diagnostic window-label builder into `M_EXCEL_UI_RUNTIME` as
  `UI_RuntimeBuildWindowLabel`, shared by the apply and snapshot paths, and
  removed the private copy from `M_EXCEL_UI_SNAPSHOT`. The fallback label used
  when Excel exposes neither a caption nor a parent workbook name is now
  `Excel window`.
- Updated all source, demo and regression module metadata to version `1.1.0`.

### Fixed

- Fixed `UI_ShowExcelUI` silently failing to restore the title bar when no
  owned-bit baseline had been captured and the frame was already hidden - the
  state reached after a VBA project reset, which is precisely when the
  documented emergency recovery path is needed. The operation reported success
  through both diagnostic paths while the title bar stayed hidden, and no later
  call could recover it. A show with no captured baseline now restores the full
  owned frame. Introduced during this release cycle by the owned-bit merge; not
  present in `1.0.1`.
- Fixed one failed application-level property read discarding the entire
  snapshot. `Application.DisplayStatusBar`, `DisplayScrollBars` and
  `DisplayFormulaBar` were read directly under an active error handler, so an
  ordinary host refusal cleared the Ribbon state, the frame state and every
  captured window identity. Because `UI_CaptureExcelUIState` returns nothing,
  the loss was silent until restore time. All three reads now route through the
  fail-soft helper, record a `Known` flag, and continue the pass.
- Fixed restoration writing default `False` values over good host state after a
  partial capture. Status Bar, Scroll Bars and Formula Bar now carry `Known`
  flags, and restoration writes each only when its captured value is
  meaningful.
- Fixed every `Err`-derived diagnostic reporting `0: ` with an empty
  description. `UI_RuntimeBuildErrorText`, `UI_TitleBarBuildRuntimeErrorText`
  and `TST_BuildRuntimeErrorText` read the `Err` object after executing
  `On Error Resume Next`, and any form of `On Error` resets `Err`. The guard
  intended to stop the formatter raising inside an error handler blanked the
  values it existed to report. All three now capture `Err.Number`,
  `Err.Description`, `Err.Source` and `Erl` before protecting themselves. This
  affected the `Unexpected` stage on both the Immediate Window path and the
  `FailureList` returned by every `_WithResult` API.
- Fixed `TST_SetWindowPos` in the regression harness declaring no `Alias`, so
  VBA searched `user32.dll` for an export of that literal name and raised error
  453. The defect was latent because `TST_TryRefreshWindowFrame` had no caller.
- Fixed one unusable Excel window aborting the rest of a multi-window pass.
  `UI_ApplyWindowLevelState` had no error handler, and the caller's handler
  ends in `Resume Safe_Exit`, so an error raised while processing one window
  abandoned every window still to be visited. The trigger was in the failure
  path rather than the writes: composing a diagnostic read
  `TargetWindow.Caption`, which can itself raise on the window that is already
  failing. The procedure now handles errors locally, records one entry naming
  the window, and returns so the enumeration continues; the label is built once
  on entry, so no window property is read while composing a failure message.

### Compatibility

```text
Existing calls affected: none
Backward compatible:     Yes
Release type:            minor
```

- No public procedure was removed or renamed.
- No existing parameter changed name, position, type or default.
- `TargetScope` is declared after `FailureCount` and `FailureList` in
  `UI_SetExcelUI_WithResult`, so existing positional callers are unaffected.
- No enum member or value changed. `UIVisibility` is unchanged.
- The `Stage | Detail` diagnostic format is unchanged. One new stage value,
  `Window [label]`, can now appear in a failure list.
- `UI_ShowExcelUI` still means "show all managed elements", not "restore the
  captured baseline".
- Snapshot storage remains in memory only and does not survive a VBA project
  reset or an Excel restart.
- Installation now requires all four `src/` modules. Importing only
  `M_EXCEL_UI.bas` is not a valid installation; see `INSTALLATION.md`.

### Validation

Validated manually in desktop Microsoft Excel for Windows:

```text
Debug -> Compile VBAProject        PASS
Test_EXCEL_UI_RunCore              PASS
Test_EXCEL_UI_RunTitleBarOnly      PASS
Test_EXCEL_UI_RunSnapshotIdentity  PASS
Test_EXCEL_UI_RunAll               PASS
```

Manual checks completed: `UI_HideExcelUI` / `UI_ShowExcelUI` recovery, and
capture / hide / reset validation.

### Known limitations

- Ribbon and title-bar control remain best effort and depend on Excel version,
  window state, Windows desktop composition and other loaded add-ins.
- The title-bar show-recovery regression case reproduces the observable
  precondition rather than a real VBA project reset, because VBA offers no
  supported way to clear another module's private state from code.
- The per-element application-level capture case cannot force a host read
  failure; it guards the independence contract rather than the failing read
  itself.
- Hidden Excel UI is not a security boundary.

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

[Unreleased]: https://github.com/danielep71/VBA-EXCEL_UI/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/danielep71/VBA-EXCEL_UI/compare/v1.0.1...v1.1.0
[1.0.1]: https://github.com/danielep71/VBA-EXCEL_UI/releases/tag/v1.0.1
