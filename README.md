<img width="1774" height="338" alt="excel_ui_Banner 2" src="https://github.com/user-attachments/assets/702a3603-3744-4012-8a4c-fcf44d39bba8" />

# VBA-EXCEL UI

> Windows Excel UI control for VBA — show, hide, snapshot, and restore Ribbon, bars, worksheet aids, and title bar through a structured tri-state API

---

<p align="center">
  <img alt="Platform" src="https://img.shields.io/badge/Platform-Excel_VBA-217346">
  <img alt="Office" src="https://img.shields.io/badge/Office-32%2F64--bit-blue">
  <img alt="Layer" src="https://img.shields.io/badge/Layer-UI controller-6f42c1">
  <img alt="OS" src="https://img.shields.io/badge/OS-Windows-0078D6">
  <img alt="API" src="https://img.shields.io/badge/API-WinAPI-blue">
  <img alt="Status" src="https://img.shields.io/badge/Status-FINAL-brightgreen">
  <img alt="License" src="https://img.shields.io/badge/License-MIT-green">
</p>

---

## Part of a larger framework

This repository is part of the **Excel VBA Runtime Framework**:

👉 https://github.com/danielep71/excel-vba-runtime-framework

It is intended for Excel/VBA solutions that need to:

- control the Excel shell from one place
- present a cleaner, more guided user experience
- reduce visual noise during workbook-driven workflows
- apply consistent UI behavior across projects
- support explicit snapshot / restore of managed UI state

`VBA-EXCEL_UI` can be used on its own, but it also fits naturally into a broader runtime architecture alongside modules for execution control, event-driven interaction, and application-style workbook behavior.

---

<img width="1536" height="1024" alt="EXCEL UI - HOME" src="https://github.com/user-attachments/assets/574869d3-f17b-4daa-a17e-aa4c79e15bf7" />

---

## Overview

`VBA-EXCEL_UI` is a compact but structured **VBA toolkit for controlling the Excel UI shell on Windows**.

It centralizes control of the Excel interface elements that are most commonly relevant in workbook-driven solutions, including:

- Ribbon
- Status Bar
- Scroll Bars
- Formula Bar
- Headings
- Workbook Tabs
- Gridlines
- Title Bar

The project is designed for scenarios where Excel should behave less like a general-purpose spreadsheet and more like a controlled application surface.

Typical examples include:

- guided workbook workflows
- kiosk-like or presentation-oriented shells
- controlled data-entry experiences
- demo environments
- application-style Excel solutions
- repeatable UI setup and restoration during automation

The repository currently includes:

- a **core UI control module**
- a **structured-result path** for callers that need diagnostics as data
- an **explicit snapshot / reset workflow** for managed UI state
- a **worksheet-based demo module**
- a **regression test harness**

---

## What this repository includes

### Core UI controller

The core module exposes a tri-state API for selective UI control:

- `UI_SetExcelUI`
- `UI_SetExcelUI_WithResult`
- `UI_HideExcelUI`
- `UI_ShowExcelUI`

### Explicit snapshot / restore workflow

The repository also includes an explicit managed-state lifecycle:

- `UI_CaptureExcelUIState`
- `UI_ResetExcelUIToSnapshot`
- `UI_HasExcelUIStateSnapshot`
- `UI_ClearExcelUIStateSnapshot`

### Interactive demo

The demo module builds a worksheet-based showcase and lets users:

- select which UI elements to affect
- apply **SHOW** or **HIDE** only to selected elements
- synchronize check boxes back to current UI state
- use preset selection profiles
- capture the current managed UI baseline
- reset the managed UI back to the captured baseline

### Regression tests

The regression harness validates the public behavior of the toolkit, including:

- show-all baseline behavior
- selective hide
- selective show
- leave-unchanged / no-op semantics
- convenience wrappers
- structured-result success paths
- snapshot lifecycle behavior
- reset-without-snapshot behavior
- `ScreenUpdating` preservation
- title-bar round-trip behavior

---

## Why this exists

Excel exposes many interface elements that are useful in ordinary spreadsheet work, but not always desirable in structured VBA solutions.

In real projects, it is often useful to:

- reduce visual noise
- guide user interaction
- hide implementation-oriented Excel chrome
- standardize the workbook interface before or during automation
- restore a known visual state afterward

`VBA-EXCEL_UI` exists to provide a **single control layer** for those needs.

Instead of scattering raw UI toggles across workbooks and macros, this repository gives you one reusable API for controlled Excel presentation behavior.

---

## Why not just toggle Excel properties directly

You can hide many Excel UI elements directly through the object model, but doing that ad hoc tends to become fragile and inconsistent.

`VBA-EXCEL_UI` adds structure that simple one-off property writes do not provide:

- an explicit **tri-state API**
- clear **leave unchanged** semantics
- centralized **best-effort processing**
- optional **structured diagnostics**
- explicit **snapshot / restore** behavior
- wrapped **WinAPI title-bar control**
- reduced no-op churn where possible

This makes the code easier to reuse, easier to test, and easier to maintain.

---

## Repository contents

```text
/README.md
/src/M_EXCEL_UI.bas
/demo/M_EXCEL_UI_DEMO.bas
/demo/M_DEMO_BUILDER.bas
/demo/EXCEL_UI_DEMO.xlsm
/test/M_EXCEL_UI_REGRESSION_TESTS.bas
```

---

## Main components

### `M_EXCEL_UI.bas`

Core Excel UI control module.

Public surface:

- `UIVisibility`
- `UI_SetExcelUI`
- `UI_SetExcelUI_WithResult`
- `UI_HideExcelUI`
- `UI_ShowExcelUI`
- `UI_CaptureExcelUIState`
- `UI_ResetExcelUIToSnapshot`
- `UI_HasExcelUIStateSnapshot`
- `UI_ClearExcelUIStateSnapshot`

Managed UI elements:

**Application-level**

- Ribbon
- Status Bar
- Scroll Bars
- Formula Bar

**Window-level**

- Headings
- Workbook Tabs
- Gridlines

**Window frame**

- Title Bar through WinAPI on the Excel window represented by `Application.Hwnd`

### `M_EXCEL_UI_DEMO.bas`

Worksheet-based demo and demo-sheet builder.

Main public procedures:

- `Demo_CreateDemoSheet`
- `Demo_ShowSelectedUI`
- `Demo_HideSelectedUI`
- `Demo_SyncCheckBoxesToUI`
- `Demo_SelectAllUI`
- `Demo_ClearAllUI`
- `Demo_PresetKiosk`
- `Demo_PresetAnalyst`
- `Demo_PresetMinimal`
- `Demo_CaptureUIState`
- `Demo_ResetUIToCapturedState`

The demo provides:

- selective **SHOW** / **HIDE** actions driven by worksheet check boxes
- a builder macro that creates or rebuilds the demo sheet
- **Sync Checkboxes** support to read current UI state back into the demo controls
- convenience actions such as:
  - Select All
  - Clear All
- preset selection profiles
- explicit state workflow buttons:
  - Capture State
  - Reset State
- explanatory notes rendered directly on the demo sheet

### `M_EXCEL_UI_REGRESSION_TESTS.bas`

Regression-test harness for the public API.

Public runners:

- `Test_EXCEL_UI_RunAll`
- `Test_EXCEL_UI_RunCore`
- `Test_EXCEL_UI_RunTitleBarOnly`

The harness validates:

- show-all baseline behavior
- selective hide
- selective show
- leave-unchanged / no-op semantics
- convenience wrappers
- title-bar hide / show round-trip
- structured-result clean success path
- structured-result no-op success path
- structured-result success path without failure-list capture
- explicit snapshot lifecycle
- reset without snapshot
- `ScreenUpdating` preservation around quiet-update behavior

It snapshots current UI state before test execution and attempts to restore it afterward.

---

## Core design

### Tri-state API

The core module uses an explicit tri-state enum rather than Boolean optional arguments:

```vb
Public Enum UIVisibility
    UI_LeaveUnchanged = -1
    UI_Hide = 0
    UI_Show = 1
End Enum
```

This avoids the ambiguity of omitted Boolean arguments and makes caller intent explicit.

### Best-effort processing

`UI_SetExcelUI` and `UI_SetExcelUI_WithResult` use best-effort processing.

That means one failed UI element does **not** prevent the module from attempting the remaining requested changes.

This is especially important when working across different Excel surfaces, host configurations, or WinAPI-sensitive areas.

### Fire-and-forget vs structured diagnostics

The repository provides two complementary entry points:

#### `UI_SetExcelUI`

Best-effort, fail-soft application with diagnostics written to the Immediate Window.

#### `UI_SetExcelUI_WithResult`

Best-effort, fail-soft application that returns:

- a Boolean success flag
- `FailureCount`
- optional `FailureList` as a 1-based string array

This allows callers to choose between a simpler procedural API and a structured inspection path without relying on a dedicated class module.

### Explicit snapshot / reset

The snapshot lifecycle is deliberately separate from `UI_ShowExcelUI`.

Use:

- `UI_CaptureExcelUIState`
- `UI_ResetExcelUIToSnapshot`
- `UI_HasExcelUIStateSnapshot`
- `UI_ClearExcelUIStateSnapshot`

when you want to:

1. capture the current managed UI baseline
2. apply a constrained shell
3. restore the captured baseline later

`UI_ShowExcelUI` means **show all managed UI**, not **restore previous state**.

### Reduced redraw where possible

The core apply path tries to reduce unnecessary UI churn by:

- temporarily suppressing `Application.ScreenUpdating` where possible
- skipping no-op writes when the current state already matches the requested target

This improves smoothness for object-model UI elements, even though it cannot fully suppress Ribbon refresh or WinAPI non-client frame repaint behavior.

### Title-bar control

Excel does not expose title-bar visibility directly through the object model.

For that reason, the module uses WinAPI to update the style of the Excel window represented by `Application.Hwnd`.

This makes title-bar behavior available through the same public API surface as the rest of the managed UI.

---

## Public API

### `UI_SetExcelUI`

Selective UI control entry point.

Example:

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

### `UI_SetExcelUI_WithResult`

Selective UI control entry point returning structured diagnostics.

Example:

```vb
Dim OK As Boolean
Dim FailureCount As Long
Dim FailureList As Variant
Dim i As Long

OK = UI_SetExcelUI_WithResult( _
        Ribbon:=UI_Hide, _
        StatusBar:=UI_Show, _
        ScrollBars:=UI_Hide, _
        FormulaBar:=UI_LeaveUnchanged, _
        Headings:=UI_Hide, _
        WorkbookTabs:=UI_Hide, _
        Gridlines:=UI_Hide, _
        TitleBar:=UI_Hide, _
        FailureCount:=FailureCount, _
        FailureList:=FailureList)

If Not OK Then
    For i = 1 To FailureCount
        Debug.Print FailureList(i)
    Next i
End If
```

### `UI_HideExcelUI`

Hide all managed UI elements.

```vb
UI_HideExcelUI
```

### `UI_ShowExcelUI`

Show all managed UI elements.

```vb
UI_ShowExcelUI
```

### `UI_CaptureExcelUIState`

Capture the current managed Excel UI state explicitly.

```vb
UI_CaptureExcelUIState
```

### `UI_ResetExcelUIToSnapshot`

Best-effort restore to the most recently captured managed UI snapshot.

```vb
UI_ResetExcelUIToSnapshot
```

### `UI_HasExcelUIStateSnapshot`

Return whether a snapshot is currently available.

```vb
Debug.Print UI_HasExcelUIStateSnapshot
```

### `UI_ClearExcelUIStateSnapshot`

Clear any previously captured managed UI snapshot.

```vb
UI_ClearExcelUIStateSnapshot
```

---

## Demo quick start

1. Import `M_EXCEL_UI.bas`
2. Import `M_EXCEL_UI_DEMO.bas`
3. Run:

```vb
Demo_CreateDemoSheet
```

4. Use the generated `DEMO_UI` worksheet to:

- select target UI elements with check boxes
- apply **SHOW SELECTED UI**
- apply **HIDE SELECTED UI**
- synchronize check boxes from current UI state
- apply preset selections
- capture the current managed UI state
- reset the managed UI to the captured state

---

## Regression test quick start

1. Import `M_EXCEL_UI.bas`
2. Import `M_EXCEL_UI_REGRESSION_TESTS.bas`
3. Run one of:

```vb
Test_EXCEL_UI_RunCore
```

```vb
Test_EXCEL_UI_RunTitleBarOnly
```

```vb
Test_EXCEL_UI_RunAll
```

Suggested order for manual validation:

1. `Test_EXCEL_UI_RunCore`
2. `Test_EXCEL_UI_RunTitleBarOnly`
3. `Test_EXCEL_UI_RunAll`

---

## Importing into VBA

### Core only

Import:

- `M_EXCEL_UI.bas`

### Core + demo

Import:

- `M_EXCEL_UI.bas`
- `M_EXCEL_UI_DEMO.bas`

### Core + tests

Import:

- `M_EXCEL_UI.bas`
- `M_EXCEL_UI_REGRESSION_TESTS.bas`

### Full repository behavior

Import:

- `M_EXCEL_UI.bas`
- `M_EXCEL_UI_DEMO.bas`
- `M_EXCEL_UI_REGRESSION_TESTS.bas`

---

## Typical use cases

### Application-like Excel solutions

Hide unnecessary Excel chrome and present a cleaner, guided interface to the user.

### Controlled workflows

Ensure that a workbook opens or runs under a consistent visual configuration.

### Demo and presentation environments

Temporarily suppress distracting UI elements for workshops, demonstrations, or executive walkthroughs.

### Protected or guided user interaction

Limit visible interface elements so users focus on intended actions and input areas.

### Repeatable UI orchestration in larger frameworks

Use a centralized UI-control layer as part of a broader workbook runtime architecture.

---

## Requirements

- Microsoft Excel for Windows
- VBA-enabled workbook or add-in host
- 32-bit or 64-bit Office/VBA supported by conditional compilation
- WinAPI access available in the host environment

---

## Limitations

- Windows only
- title-bar control is WinAPI-based, not object-model-based
- title-bar behavior is best effort and may remain somewhat sensitive to Excel version, window state, and Windows desktop composition behavior
- `UI_ShowExcelUI` means **show all managed UI**, not **restore previous custom state**
- explicit snapshot/reset is **best effort**, especially for per-window restore when the set or order of open windows has changed
- the demo module’s window-level synchronization reads from `ActiveWindow`
- the demo-sheet builder performs a destructive rebuild of the `DEMO_UI` sheet
- reduced redraw does not fully eliminate Ribbon or non-client frame refresh flicker

---

## Design notes

- use centralized wrappers rather than scattering raw UI toggles throughout a project
- use selective tri-state control when only some elements should change
- use hide-all / show-all wrappers for consistent workbook presentation flows
- treat title-bar control as best effort because it depends on supported Windows behavior
- pair this module with execution-control and event-driven components for more complete Excel application design

---

## Position in the framework

Within the **Excel VBA Runtime Framework**, `VBA-EXCEL_UI` is the component responsible for **UI management and interface control**.

It is intended to work alongside complementary components for:

- execution control and performance management
- event-driven interaction
- broader Excel application architecture

Framework home:

👉 https://github.com/danielep71/excel-vba-runtime-framework

---

## Wiki

For additional examples, notes, and repository-level guidance, see the project wiki:

[EXCEL UI Wiki](https://github.com/danielep71/VBA-EXCEL_UI/wiki)

---

## Status

This repository is intended as a reusable VBA component for Excel-based solutions on Windows.

---

## Author

Daniele Penza

---

## License

This project is licensed under the terms in `LICENSE`.
