<div align="center">

# 🪟 VBA Excel UI

### A structured Windows Excel UI controller for application-style workbooks

**Tri-state visibility control · Targeted window scopes · Best-effort execution · Structured diagnostics · Identity-safe snapshot restore · Owned-bit title-bar management · Modular architecture**

<br>

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-32%20%2F%2064--bit-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://github.com/danielep71/VBA-EXCEL_UI)
[![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)](#requirements)
[![API](https://img.shields.io/badge/API-Backward_Compatible-6f42c1?style=for-the-badge)](#public-api)
[![Modules](https://img.shields.io/badge/Production_Modules-4-c2185b?style=for-the-badge)](INSTALLATION.md)
[![Status](https://img.shields.io/badge/Status-Release_Candidate-d97706?style=for-the-badge)](#status)

<br>

[![Release](https://img.shields.io/github/v/release/danielep71/VBA-EXCEL_UI?style=flat-square&label=release&color=217346)](https://github.com/danielep71/VBA-EXCEL_UI/releases)
[![License](https://img.shields.io/github/license/danielep71/VBA-EXCEL_UI?style=flat-square&color=2ea44f)](LICENSE)
[![Issues](https://img.shields.io/github/issues/danielep71/VBA-EXCEL_UI?style=flat-square&color=d73a49)](https://github.com/danielep71/VBA-EXCEL_UI/issues)

<br>

**No installer · No COM add-in · No third-party DLL · No non-standard VBA reference**

[Installation](INSTALLATION.md)
&nbsp;·&nbsp;
[Quick start](#quick-start)
&nbsp;·&nbsp;
[Public API](#public-api)
&nbsp;·&nbsp;
[Target scopes](#target-scopes)
&nbsp;·&nbsp;
[Architecture](#architecture)
&nbsp;·&nbsp;
[Regression tests](#regression-testing)
&nbsp;·&nbsp;
[Download demo release assets](https://github.com/danielep71/VBA-EXCEL_UI/releases)

---

<p align="center">
  <img width="100%"
       alt="VBA Excel UI — structured Excel interface controller"
       src="https://github.com/user-attachments/assets/702a3603-3744-4012-8a4c-fcf44d39bba8">
</p>

---

</div>

---

> [!IMPORTANT]
> Version 1.1.0 uses a **four-module production package**. Importing only
> `M_EXCEL_UI.bas` is not a valid installation. See [INSTALLATION.md](INSTALLATION.md).

> [!NOTE]
> The macro-enabled demo workbook is **not versioned in the repository**.
> Tested `.xlsm` demo builds are distributed only as GitHub Release assets.

## ✨ What this project is

**VBA Excel UI** is a focused VBA component for controlling the visible Excel shell on Windows.

It provides one stable project-facing API for:

- showing, hiding, or leaving unchanged individual Excel UI elements;
- applying application-level and window-level settings consistently;
- targeting all Excel windows, only the active window, or all windows of the active workbook;
- controlling the Excel title bar through bitness-safe WinAPI calls;
- capturing and restoring an in-memory UI baseline;
- restoring each captured Excel window by retained object identity rather than collection index;
- returning ordered structured diagnostics;
- maintaining a deterministic show-all emergency recovery path.

The public API remains concentrated in `M_EXCEL_UI`. Internal responsibilities are separated into private project modules for runtime services, snapshot state, and title-bar handling.

---

<a id="managed-ui-surface"></a>

## 🎚️ Managed UI surface

| UI element | Scope | Mechanism | Targetable |
|---|---|---|:---:|
| Ribbon | Excel application | Ribbon command with best-effort state reads | No |
| Status Bar | Excel application | `Application.DisplayStatusBar` | No |
| Scroll Bars | Excel application | `Application.DisplayScrollBars` | No |
| Formula Bar | Excel application | `Application.DisplayFormulaBar` | No |
| Headings | Excel window | `Window.DisplayHeadings` | Yes |
| Workbook Tabs | Excel window | `Window.DisplayWorkbookTabs` | Yes |
| Gridlines | Excel window | `Window.DisplayGridlines` | Yes |
| Title Bar | Excel main window | Owned-bit WinAPI update on `Application.Hwnd` | No |

Application-level changes affect the current Excel process. `TargetScope` applies only to Headings, Workbook Tabs, and Gridlines.

---

<a id="quick-start"></a>

# ⚡ Quick start

## 1. Import the complete production package

Import all four files from `src/`:

```text
src/M_EXCEL_UI_RUNTIME.bas
src/M_EXCEL_UI_TITLEBAR.bas
src/M_EXCEL_UI_SNAPSHOT.bas
src/M_EXCEL_UI.bas
```

Recommended dependency-first order:

```text
M_EXCEL_UI_RUNTIME
M_EXCEL_UI_TITLEBAR
M_EXCEL_UI_SNAPSHOT
M_EXCEL_UI
```

Then compile:

```text
VBA Editor → Debug → Compile VBAProject
```

For upgrade and troubleshooting instructions, see [INSTALLATION.md](INSTALLATION.md).

## 2. Apply selective UI control

```vb
UI_SetExcelUI _
    Ribbon:=UI_Hide, _
    StatusBar:=UI_Show, _
    ScrollBars:=UI_Hide, _
    FormulaBar:=UI_LeaveUnchanged, _
    Headings:=UI_Hide, _
    WorkbookTabs:=UI_Hide, _
    Gridlines:=UI_Hide, _
    TitleBar:=UI_Hide
```

Only explicitly requested elements are changed.

The default target scope remains all current Excel windows, preserving pre-v1.1.0 behavior.

## 3. Target a specific window scope

Only window-level elements are affected by `TargetScope`.

Active window only:

```vb
UI_SetExcelUI _
    Headings:=UI_Hide, _
    WorkbookTabs:=UI_Hide, _
    Gridlines:=UI_Hide, _
    TargetScope:=UI_TargetActiveWindow
```

All windows belonging to the active workbook:

```vb
UI_SetExcelUI _
    Headings:=UI_Show, _
    Gridlines:=UI_Show, _
    TargetScope:=UI_TargetActiveWorkbookWindows
```

Application-level elements still operate at their established scope even when a restricted target is selected.

## 4. Request structured diagnostics

```vb
Dim OK As Boolean
Dim FailureCount As Long
Dim FailureList As Variant
Dim i As Long

OK = UI_SetExcelUI_WithResult( _
        StatusBar:=UI_Show, _
        Headings:=UI_Hide, _
        Gridlines:=UI_Hide, _
        FailureCount:=FailureCount, _
        FailureList:=FailureList, _
        TargetScope:=UI_TargetActiveWindow)

If Not OK Then
    For i = 1 To FailureCount
        Debug.Print FailureList(i)
    Next i
End If
```

Failure entries use:

```text
Stage | Detail
```

## 5. Hide or show the complete managed shell

```vb
UI_HideExcelUI
UI_ShowExcelUI
```

`UI_ShowExcelUI` means **show every managed element**. It does not restore a captured custom baseline.

## 6. Capture and restore a managed baseline

```vb
UI_CaptureExcelUIState
UI_HideExcelUI
UI_ResetExcelUIToSnapshot
```

Structured capture and reset:

```vb
Dim OK As Boolean
Dim FailureCount As Long
Dim FailureList As Variant

OK = UI_CaptureExcelUIState_WithResult( _
        FailureCount:=FailureCount, _
        FailureList:=FailureList)

OK = UI_ResetExcelUIToSnapshot_WithResult( _
        FailureCount:=FailureCount, _
        FailureList:=FailureList)
```

When the snapshot is no longer needed:

```vb
UI_ClearExcelUIStateSnapshot
```

Snapshot capture/restore retains its established all-managed-windows semantics. `TargetScope` applies to selective UI application, not to snapshot capture or restore.

---

<a id="public-api"></a>

# 🧩 Public API

## Public enums

```vb
Public Enum UIVisibility
    UI_LeaveUnchanged = -1
    UI_Hide = 0
    UI_Show = 1
End Enum
```

```vb
Public Enum UIWindowTargetScope
    UI_TargetAllExcelWindows = 0
    UI_TargetActiveWindow = 1
    UI_TargetActiveWorkbookWindows = 2
End Enum
```

## API reference

| Member | Type | Purpose | Diagnostic behavior |
|---|---|---|---|
| `UIVisibility` | Public enum | Show, hide, or leave unchanged | Not applicable |
| `UIWindowTargetScope` | Public enum | Select window-level target scope | Invalid values are controlled by the apply path |
| `UI_SetExcelUI` | Public `Sub` | Apply selective managed UI state | Logs failures |
| `UI_SetExcelUI_WithResult` | Public `Function` | Apply selective state with structured result | Boolean + count + optional list |
| `UI_HideExcelUI` | Public `Sub` | Hide all managed UI elements | Fail-soft; logs failures |
| `UI_ShowExcelUI` | Public `Sub` | Show all managed UI elements | Fail-soft; logs failures |
| `UI_CaptureExcelUIState` | Public `Sub` | Capture a managed baseline | Best effort; logs failures |
| `UI_CaptureExcelUIState_WithResult` | Public `Function` | Capture with ordered structured diagnostics | Boolean + count + optional list |
| `UI_ResetExcelUIToSnapshot` | Public `Sub` | Restore the current snapshot | Best effort; logs failures |
| `UI_ResetExcelUIToSnapshot_WithResult` | Public `Function` | Restore with ordered structured diagnostics | Boolean + count + optional list |
| `UI_HasExcelUIStateSnapshot` | Public `Function` | Report snapshot availability | Returns `Boolean` |
| `UI_ClearExcelUIStateSnapshot` | Public `Sub` | Discard snapshot state | No return value |

The v1.1.0 targeting extension is backward compatible:

- existing public names are preserved;
- existing parameter order is preserved;
- `TargetScope` is optional and trailing;
- the default is `UI_TargetAllExcelWindows`;
- existing calls therefore keep their previous behavior.

---

<a id="target-scopes"></a>

# 🎯 Target scopes

`UIWindowTargetScope` controls only:

```text
Headings
Workbook Tabs
Gridlines
```

Supported values:

| Value | Meaning |
|---|---|
| `UI_TargetAllExcelWindows` | Apply to every current Excel window; compatibility default |
| `UI_TargetActiveWindow` | Apply only to `Application.ActiveWindow` |
| `UI_TargetActiveWorkbookWindows` | Apply to all windows in `ActiveWorkbook.Windows` |

Ribbon, Status Bar, Scroll Bars, Formula Bar, and Title Bar keep their existing application/main-window scope.

An unsupported target value is handled through the established best-effort diagnostics contract. Valid application-level operations can still proceed, while window-level writes are suppressed when no safe scope can be resolved.

---

<a id="architecture"></a>

# 🏗️ Architecture

```mermaid
flowchart TD
    CALLER[Workbook, add-in, demo, or tests]
    FACADE[M_EXCEL_UI<br/>public facade and targeting]
    RUNTIME[M_EXCEL_UI_RUNTIME<br/>host operations and diagnostics]
    SNAPSHOT[M_EXCEL_UI_SNAPSHOT<br/>snapshot state and identity-safe restore]
    TITLEBAR[M_EXCEL_UI_TITLEBAR<br/>WinAPI and owned style bits]
    EXCEL[Excel object model / Windows API]

    CALLER --> FACADE
    FACADE --> RUNTIME
    FACADE --> SNAPSHOT
    FACADE --> TITLEBAR
    SNAPSHOT --> RUNTIME
    SNAPSHOT --> TITLEBAR
    RUNTIME --> EXCEL
    TITLEBAR --> EXCEL
    SNAPSHOT --> EXCEL
```

## Module responsibilities

| Module | Responsibility | Caller-facing API |
|---|---|---|
| `M_EXCEL_UI` | Public facade, enums, tri-state validation, target resolution, general apply orchestration | Yes |
| `M_EXCEL_UI_RUNTIME` | Shared diagnostics, result buffers, Ribbon/property helpers, quiet-update scope | No |
| `M_EXCEL_UI_SNAPSHOT` | Snapshot state, capture, restore, retained Window identity resolution | No |
| `M_EXCEL_UI_TITLEBAR` | WinAPI declarations, title-bar state, owned-bit merging and frame refresh | Internal test seam only |

All four modules use `Option Explicit` and `Option Private Module`.

## Dependency constraints

- `M_EXCEL_UI_RUNTIME` has no project-module dependency.
- `M_EXCEL_UI_TITLEBAR` has no project-module dependency.
- `M_EXCEL_UI_SNAPSHOT` depends on runtime and title-bar services.
- `M_EXCEL_UI` depends on all three internal modules.
- No circular dependency or duplicate mutable state is intended.

---

## 📸 Snapshot lifecycle and identity

The snapshot is stored in module memory and is lost after VBA reset, project unload, or Excel exit.

For each Excel window, the snapshot engine retains the exact captured `Window` object reference plus a diagnostic label. Restore never selects a target by `Application.Windows` collection index.

Expected behavior:

- window reordering does not redirect captured state;
- activation changes do not redirect captured state;
- a still-live captured window can be restored after its collection position changes;
- a newly opened window is left unchanged;
- a closed, recreated, or otherwise unusable captured window is reported as a best-effort failure;
- state is never intentionally applied to a different replacement window.

Every captured element carries a `Known` flag recording whether its value was actually readable — the Ribbon, the title bar, the three application-level properties, and each window's Headings, Workbook Tabs and Gridlines. Capture continues after an element-level failure, and restoration never writes a value that was not successfully captured. A partial capture is therefore still a usable snapshot.

> [!CAUTION]
> Snapshot restoration is not persistent or transactional. It cannot survive a VBA project reset or guarantee rollback after Excel process termination.

> [!IMPORTANT]
> A snapshot retains a live `Window` reference per captured window, and restoring does not release them: the snapshot is deliberately kept so it can be replayed. Call `UI_ClearExcelUIStateSnapshot` when the baseline is no longer needed — `Workbook_BeforeClose` is the natural place. See [INSTALLATION.md](INSTALLATION.md#snapshot-lifetime-and-window-references).

---

## 🪟 Title-bar ownership

The title-bar subsystem owns only these `GWL_STYLE` bits:

- `WS_CAPTION`;
- `WS_SYSMENU`;
- `WS_THICKFRAME`;
- `WS_MINIMIZEBOX`;
- `WS_MAXIMIZEBOX`.

Showing the title bar merges the captured owned bits into the current style. It does not restore an entire historical style value, so unrelated style changes are preserved.

When a show is requested and no baseline was ever captured for the current handle, the full owned frame is restored instead. That case is reached after a VBA project reset, because the window style belongs to the running Excel process and survives while module state does not. Without the fallback a show would re-apply the current hidden bits, report success, and leave the title bar hidden — so this is what keeps `UI_ShowExcelUI` a real recovery path.

The WinAPI path remains Windows-only and best effort. It supports VBA7 32-bit, VBA7 64-bit, and the legacy 32-bit declaration path through conditional compilation.

---

## 🛡️ Execution and diagnostics

The component deliberately uses best-effort processing:

1. a failure is recorded or logged;
2. later unrelated operations are still attempted;
3. `ScreenUpdating` is restored where the component changed it;
4. fire-and-forget APIs do not raise ordinary element-level failures.

The `WithResult` APIs return:

| Output | Meaning |
|---|---|
| `True` | No failure was recorded |
| `False` | One or more failures were recorded |
| `FailureCount` | Number of failures |
| `FailureList` | Optional 1-based ordered string array |

Output buffers are cleared deterministically on entry.

---

<a id="regression-testing"></a>

# ✅ Regression testing

Optional test module:

```text
test/M_EXCEL_UI_REGRESSION_TESTS.bas
```

Public runners:

```vb
Test_EXCEL_UI_RunCore
Test_EXCEL_UI_RunTitleBarOnly
Test_EXCEL_UI_RunSnapshotIdentity
Test_EXCEL_UI_RunAll
```

Recommended validation sequence:

```text
Debug → Compile VBAProject
Test_EXCEL_UI_RunCore
Test_EXCEL_UI_RunTitleBarOnly
Test_EXCEL_UI_RunSnapshotIdentity
Test_EXCEL_UI_RunAll
UI_HideExcelUI
UI_ShowExcelUI
```

The suite covers:

- tri-state behavior and wrappers;
- structured diagnostics;
- snapshot lifecycle and no-snapshot handling;
- identity-safe restoration and replacement-window non-interference;
- title-bar owned-bit preservation and real WinAPI round-trips;
- active-window targeting;
- active-workbook-window targeting;
- invalid-target-scope diagnostics and application-level continuation;
- title-bar show recovery when no owned-bit baseline was captured;
- per-element application-level capture and idempotent restoration.

> [!IMPORTANT]
> Tests manipulate the real Excel UI of the current process. Run them in a controlled Excel instance.

---

## 🖼️ Demo and release assets

Version-controlled demo source:

```text
demo/M_EXCEL_UI_DEMO.bas
demo/M_DEMO_BUILDER.bas
```

The macro-enabled demo workbook is intentionally **not committed to Git**.

For tagged releases, a tested workbook should be published as a GitHub Release asset, for example:

```text
EXCEL_UI_DEMO_v1.1.0.xlsm
```

Release-asset preparation should include:

1. import the exact production and demo source from the tagged release candidate;
2. compile the VBA project;
3. run the applicable regression sequence;
4. perform manual UI recovery checks;
5. save the `.xlsm`;
6. optionally calculate and publish a SHA-256 checksum;
7. attach the workbook to the GitHub Release.

Users who want a ready-to-run workbook should obtain it from the repository's **Releases** page rather than from the source tree.

---

## 🆘 Emergency recovery

Run:

```vb
UI_ShowExcelUI
```

This does not require an existing snapshot and is the preferred recovery command after interrupted development, missing snapshot state, or a VBA reset.

---

## 📦 Repository structure

```text
VBA-EXCEL_UI/
├─ .github/
├─ demo/
│  ├─ M_DEMO_BUILDER.bas
│  └─ M_EXCEL_UI_DEMO.bas
├─ src/
│  ├─ M_EXCEL_UI.bas
│  ├─ M_EXCEL_UI_RUNTIME.bas
│  ├─ M_EXCEL_UI_SNAPSHOT.bas
│  └─ M_EXCEL_UI_TITLEBAR.bas
├─ test/
│  └─ M_EXCEL_UI_REGRESSION_TESTS.bas
├─ CHANGELOG.md
├─ CODE_OF_CONDUCT.md
├─ CONTRIBUTING.md
├─ INSTALLATION.md
├─ LICENSE
├─ README.md
└─ SECURITY.md
```

The source repository intentionally contains no versioned demo `.xlsm`. Release binaries are attached to tagged GitHub Releases.

## Documentation map

| Document | Purpose |
|---|---|
| [INSTALLATION.md](INSTALLATION.md) | Fresh install, upgrade, required modules, targeting, validation, troubleshooting |
| [README.md](README.md) | Project overview, API, architecture, behavior, limitations |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Development workflow and module-boundary rules |
| [CHANGELOG.md](CHANGELOG.md) | Release history |
| [SECURITY.md](SECURITY.md) | Security reporting and safe-use boundaries |
| [Regression tests](test/M_EXCEL_UI_REGRESSION_TESTS.bas) | Behavioral verification |
| [GitHub Releases](https://github.com/danielep71/VBA-EXCEL_UI/releases) | Tested binary demo workbooks and release notes |

---

<a id="requirements"></a>

# 💻 Requirements

- Microsoft Excel desktop for Windows;
- a macro-enabled workbook or add-in host;
- VBA project access for importing `.bas` modules;
- 32-bit or 64-bit Office;
- host policy permitting the required WinAPI calls.

Unsupported:

- Excel for macOS;
- Excel for the web;
- non-Excel VBA hosts;
- environments that block required WinAPI calls.

No third-party DLL, COM component, package manager, or non-standard VBA reference is required.

---

## 🔍 Scope and limitations

- **Windows only.** Title-bar control depends on WinAPI.
- **Current Excel instance.** Application-level properties affect the running Excel process.
- **Targeted window operations.** Headings, Workbook Tabs, and Gridlines support all windows, active window, or active-workbook windows.
- **Identity-safe but in-memory snapshots.** Window identity is retained by object reference, but snapshots do not survive VBA reset or Excel restart.
- **Changed window set after capture.** New windows are unchanged; missing captured windows produce diagnostics.
- **Best-effort Ribbon and frame behavior.** Excel, Windows, add-ins, and window mode can affect visible results.
- **No durable transaction.** Process termination can prevent restoration.
- **Not a security boundary.** Hidden Excel UI does not prevent other code or informed users from changing state.
- **Binary demo not source-controlled.** The ready-to-run `.xlsm` is a release asset, not part of the repository tree.

---

## 🧭 v1.1.0 scope status

Completed:

- identity-safe per-window snapshot restoration;
- owned-bit title-bar restoration;
- structured snapshot capture and restore results;
- internal four-module decomposition;
- installation and dependency documentation;
- additional public window-targeting scopes;
- active-window and active-workbook-window regression coverage;
- invalid-target-scope structured diagnostics.

Remaining release-maintenance work:

- synchronize the demo source/workbook for the final release candidate;
- publish the validated `.xlsm` as a GitHub Release asset;
- update `CHANGELOG.md` and final release notes;
- review the complete release branch diff;
- open and review the release pull request;
- merge, tag `v1.1.0`, and publish the release.

---

## ✅ Release checklist

```text
[ ] Confirm current branch is release/v1.1.0
[ ] Import all four src/ production modules
[ ] Import the regression module
[ ] Debug → Compile VBAProject
[ ] Run Test_EXCEL_UI_RunCore
[ ] Run Test_EXCEL_UI_RunTitleBarOnly
[ ] Run Test_EXCEL_UI_RunSnapshotIdentity
[ ] Run Test_EXCEL_UI_RunAll
[ ] Perform UI_HideExcelUI / UI_ShowExcelUI recovery
[ ] Perform capture / hide / reset validation
[ ] Validate active-window targeting
[ ] Validate active-workbook-window targeting
[ ] Review exported .bas diffs and CRLF handling
[ ] Build and validate EXCEL_UI_DEMO_v1.1.0.xlsm outside Git tracking
[ ] Calculate demo SHA-256 if published in release notes
[ ] Update README, INSTALLATION, CHANGELOG, and release notes
[ ] Review the complete release branch diff
[ ] Open and review the release pull request
[ ] Merge and tag v1.1.0
[ ] Publish GitHub Release
[ ] Attach validated EXCEL_UI_DEMO_v1.1.0.xlsm as release asset
```

---

<a id="status"></a>

## 📌 Status

The `release/v1.1.0` line preserves the established public `UI_...` surface while adding identity-safe snapshots, safer title-bar ownership, structured snapshot results, a cohesive four-module internal architecture, and backward-compatible window targeting.

---

## 👤 Author

**Daniele Penza**

## 📄 License

Licensed under the [MIT License](LICENSE).
