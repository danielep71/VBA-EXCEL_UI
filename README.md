# Excel UI

VBA utilities for controlling the Excel UI shell on Windows, including Ribbon, status bar, scroll bars, formula bar, worksheet aids, and WinAPI-based title-bar visibility.

## Overview

This repository provides a small VBA toolkit for Excel UI control and validation. It currently includes:

- a **core module** for applying UI visibility changes through a safe tri-state API
- a **demo module** for building and driving a worksheet-based showcase of the UI features
- a **regression test module** for validating the public behavior of the toolkit

The project is intended for workbook-driven solutions that need a constrained, kiosk-like, presentation-oriented, or application-style Excel shell.

## Repository Contents

```text
/README.md
/src/M_EXCEL_UI.bas
/src/M_EXCEL_UI_DEMO.bas
/test/M_EXCEL_UI_REGRESSION_TESTS.bas
```

## Modules

### `M_EXCEL_UI.bas`

Core Excel UI control module.

It exposes:

- `K_UIVisibility`
- `K_SetExcelUI`
- `K_HideExcelUI`
- `K_ShowExcelUI`

Managed UI elements:

Application-level:

- Ribbon
- Status Bar
- Scroll Bars
- Formula Bar

Window-level:

- Headings
- Workbook Tabs
- Gridlines

Window-frame:

- Title Bar (via WinAPI on the Excel window represented by `Application.Hwnd`)

### `M_EXCEL_UI_DEMO.bas`

Worksheet-based demo and demo-sheet builder.

It provides:

- selective **SHOW** / **HIDE** actions driven by worksheet check boxes
- a builder macro that creates or rebuilds the demo sheet
- **Sync Checkboxes** support to read current UI state back into the demo controls
- selection helpers such as:
  - Select All
  - Clear All
- preset selection profiles such as:
  - Kiosk
  - Analyst
  - Minimal
- explanatory notes rendered directly on the demo sheet

Main public procedures include:

- `Demo_ShowSelectedExcelUI`
- `Demo_HideSelectedExcelUI`
- `Demo_SyncCheckBoxesToCurrentUI`
- `Demo_SelectAllUI`
- `Demo_ClearAllUI`
- `Demo_PresetKiosk`
- `Demo_PresetAnalyst`
- `Demo_PresetMinimal`
- `Demo_CreateExcelUISheet`

### `M_EXCEL_UI_REGRESSION_TESTS.bas`

Regression-test harness for the public API.

It provides:

- `Test_EXCEL_UI_RunAll`
- `Test_EXCEL_UI_RunCore`
- `Test_EXCEL_UI_RunTitleBarOnly`

The test harness validates:

- show-all baseline behavior
- selective hide
- selective show
- leave-unchanged / no-op semantics
- convenience wrappers
- title-bar hide / show round-trip

It also snapshots current UI state before test execution and attempts to restore it afterward.

## Core Design

### Tri-state API

The core module uses a tri-state enum instead of Boolean optional arguments:

```vb
Public Enum K_UIVisibility
    K_UI_LeaveUnchanged = -1
    K_UI_Hide = 0
    K_UI_Show = 1
End Enum
```

This avoids the ambiguity of omitted Boolean arguments and makes caller intent explicit.

### Best-effort processing

`K_SetExcelUI` is designed so that one failed UI element does not prevent the rest of the requested changes from being attempted.

### Title-bar control

Excel does not expose title-bar visibility directly through the object model, so the module uses WinAPI to update the style of the Excel window represented by `Application.Hwnd`.

## Public API

### `K_SetExcelUI`

Selective UI control entry point.

Example:

```vb
K_SetExcelUI _
    Ribbon:=K_UI_Hide, _
    StatusBar:=K_UI_Show, _
    ScrollBars:=K_UI_Hide, _
    FormulaBar:=K_UI_LeaveUnchanged, _
    Headings:=K_UI_Hide, _
    WorkbookTabs:=K_UI_Hide, _
    Gridlines:=K_UI_Hide, _
    TitleBar:=K_UI_Hide
```

### `K_HideExcelUI`

Hide all managed UI elements.

```vb
K_HideExcelUI
```

### `K_ShowExcelUI`

Show all managed UI elements.

```vb
K_ShowExcelUI
```

## Demo Quick Start

1. Import `M_EXCEL_UI.bas`.
2. Import `M_EXCEL_UI_DEMO.bas`.
3. Run:

```vb
Demo_CreateExcelUISheet
```

4. Use the generated `Demo` worksheet to:
   - select target UI elements with check boxes
   - apply **SHOW SELECTED UI**
   - apply **HIDE SELECTED UI**
   - synchronize check boxes from current UI state
   - apply preset selections

## Regression Test Quick Start

1. Import `M_EXCEL_UI.bas`.
2. Import `M_EXCEL_UI_REGRESSION_TESTS.bas`.
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

## Requirements

- Microsoft Excel for Windows
- VBA-enabled workbook or add-in host
- 32-bit or 64-bit Office/VBA supported by conditional compilation
- WinAPI access available in the host environment

## Limitations

- Windows only
- no snapshot/restore of prior user-specific UI state in the core module
- `K_ShowExcelUI` means **show all managed UI**, not **restore previous custom state**
- title-bar control is WinAPI-based, not object-model-based
- title-bar behavior is best effort and may remain somewhat sensitive to Excel version, window state, and Windows desktop composition behavior
- the demo module’s window-level synchronization reads from `ActiveWindow`
- the demo-sheet builder performs a destructive rebuild of the `Demo` sheet

## Notes

- Ribbon control relies on `Application.ExecuteExcel4Macro`
- title-bar control targets the Excel window represented by `Application.Hwnd`
- the toolkit is designed for practical workbook UI control, demos, and regression validation rather than for persistent per-user shell personalization

## Suggested Use Cases

- kiosk-like Excel workbooks
- guided demos and training workbooks
- controlled presentation environments
- reducing visible Excel chrome in client-facing solutions
- regression validation after refactoring WinAPI/UI code

## Status

This repository is intended as a reusable VBA utility component for Excel-based solutions on Windows.
