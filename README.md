# VBA-EXCEL_UI

> Centralized control of Excel UI elements and application interface behavior in VBA

---

## Part of a larger framework

This module is part of the **Excel VBA Runtime Framework**:

👉 https://github.com/danielep71/excel-vba-runtime-framework

The framework provides a structured runtime layer for:

- execution control
- UI management
- event-driven interaction

Within that framework, `VBA-EXCEL_UI` acts as the **UI Controller**.

It provides the foundation for:

- centralized Excel interface control
- application-like workbook presentation
- consistent visual environments
- reusable UI-management patterns in VBA projects

---

<img width="1536" height="1024" alt="UI Home reduced" src="https://github.com/user-attachments/assets/27c7ba79-2c45-41d0-b076-f215902c7df9" />


## Overview

`VBA-EXCEL_UI` is a **compact but structured VBA toolkit for controlling the Excel UI shell on Windows**.

It is designed for workbook-driven solutions that need a constrained, kiosk-like, presentation-oriented, or application-style Excel shell, while still preserving a practical and reusable API for ordinary VBA projects.

The repository currently includes:

- a **core module** for applying Excel UI visibility changes through a safe tri-state API
- a **structured-result path** for callers that want diagnostics as data rather than only Immediate Window logging
- an **explicit snapshot / reset path** for capturing and restoring the current managed UI baseline
- a **demo module** for building and driving a worksheet-based showcase of the UI features
- a **regression test module** for validating the public behavior of the toolkit

This makes the project useful both for:

- creating cleaner and more controlled Excel user experiences
- building guided workbook workflows
- demonstrating and validating repeatable UI-management behavior in VBA

---

## Why this exists

Excel exposes many interface elements that are useful in general-purpose spreadsheet work, but not always desirable in structured VBA solutions.

In many real projects, it is useful to:

- reduce visual noise
- guide user interaction
- hide implementation-oriented Excel chrome
- standardize the workbook interface before or during automation
- restore a known visual state after execution

`VBA-EXCEL_UI` exists to provide a **single control layer** for those needs.

---

## Repository contents

```text
/README.md
/src/M_EXCEL_UI.bas
/demo/M_EXCEL_UI_DEMO.bas
/demo/EXCEL_UI_DEMO.xlsm
/test/M_EXCEL_UI_REGRESSION_TESTS.bas
```

---

## Components

### `M_EXCEL_UI.bas`

Core Excel UI control module.

It exposes:

- `K_UIVisibility`
- `K_SetExcelUI`
- `K_SetExcelUI_WithResult`
- `K_HideExcelUI`
- `K_ShowExcelUI`
- `K_CaptureExcelUIState`
- `K_ResetExcelUIToSnapshot`
- `K_HasExcelUIStateSnapshot`
- `K_ClearExcelUIStateSnapshot`

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

**Window-frame**

- Title Bar (via WinAPI on the Excel window represented by `Application.Hwnd`)

---

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
- explicit state workflow buttons:
  - Capture State
  - Reset State
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
- `Demo_CaptureCurrentExcelUIState`
- `Demo_ResetExcelUIToCapturedState`
- `Demo_CreateExcelUISheet`

---

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
- structured-result clean success path
- structured-result no-op / leave-unchanged success path
- structured-result success path without failure-list capture
- explicit snapshot lifecycle
- reset without snapshot
- `ScreenUpdating` preservation around quiet-update behavior

It also snapshots current UI state before test execution and attempts to restore it afterward.

---

## Core capabilities

- centralized control of Excel UI elements
- tri-state behavior for each element:
  - show
  - hide
  - leave unchanged
- workbook-window UI control where applicable
- application-level UI control where applicable
- title-bar visibility control through WinAPI
- reusable wrappers for common “hide all” / “show all” patterns
- structured-result path for safer integration
- explicit snapshot / reset workflow for managed UI state
- reduced redraw and no-op avoidance where possible

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

---

## Core design

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

---

### Best-effort processing

`K_SetExcelUI` and `K_SetExcelUI_WithResult` are designed so that one failed UI element does not prevent the rest of the requested changes from being attempted.

---

### Fire-and-forget vs structured diagnostics

The toolkit provides two complementary entry points:

- `K_SetExcelUI`  
  Best-effort, fail-soft application with diagnostics written to the Immediate Window

- `K_SetExcelUI_WithResult`  
  Best-effort, fail-soft application that returns:
  - a Boolean success flag
  - `FailureCount`
  - optional `FailureList` as a 1-based string array

This allows callers to choose between a simple procedural API and a structured inspection model without relying on a class module.

---

### Explicit snapshot / reset

The toolkit also provides an explicit state lifecycle that is separate from `K_ShowExcelUI`:

- `K_CaptureExcelUIState`
- `K_ResetExcelUIToSnapshot`
- `K_HasExcelUIStateSnapshot`
- `K_ClearExcelUIStateSnapshot`

This is intended for workflows where callers want to:

1. capture the current managed UI baseline
2. apply a constrained shell
3. restore the captured baseline later

---

### Reduced redraw where possible

The core apply path tries to reduce unnecessary UI churn by:

- temporarily suppressing `Application.ScreenUpdating` where possible
- skipping no-op writes when the current state already matches the requested target

This improves smoothness for object-model UI elements, though it cannot fully suppress Ribbon or WinAPI non-client refresh.

---

### Title-bar control

Excel does not expose title-bar visibility directly through the object model, so the module uses WinAPI to update the style of the Excel window represented by `Application.Hwnd`.

---

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

---

### `K_SetExcelUI_WithResult`

Selective UI control entry point returning structured diagnostics.

Example:

```vb
Dim OK As Boolean
Dim FailureCount As Long
Dim FailureList As Variant
Dim i As Long

OK = K_SetExcelUI_WithResult( _
        Ribbon:=K_UI_Hide, _
        StatusBar:=K_UI_Show, _
        ScrollBars:=K_UI_Hide, _
        FormulaBar:=K_UI_LeaveUnchanged, _
        Headings:=K_UI_Hide, _
        WorkbookTabs:=K_UI_Hide, _
        Gridlines:=K_UI_Hide, _
        TitleBar:=K_UI_Hide, _
        FailureCount:=FailureCount, _
        FailureList:=FailureList)

If Not OK Then
    For i = 1 To FailureCount
        Debug.Print FailureList(i)
    Next i
End If
```

---

### `K_HideExcelUI`

Hide all managed UI elements.

```vb
K_HideExcelUI
```

---

### `K_ShowExcelUI`

Show all managed UI elements.

```vb
K_ShowExcelUI
```

---

### `K_CaptureExcelUIState`

Capture the current managed Excel UI state explicitly.

```vb
K_CaptureExcelUIState
```

---

### `K_ResetExcelUIToSnapshot`

Best-effort restore to the most recently captured managed UI snapshot.

```vb
K_ResetExcelUIToSnapshot
```

---

### `K_HasExcelUIStateSnapshot`

Return whether a snapshot is currently available.

```vb
Debug.Print K_HasExcelUIStateSnapshot
```

---

### `K_ClearExcelUIStateSnapshot`

Clear any previously captured managed UI snapshot.

```vb
K_ClearExcelUIStateSnapshot
```

---

## Demo quick start

1. Import `M_EXCEL_UI.bas`
2. Import `M_EXCEL_UI_DEMO.bas`
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
- `K_ShowExcelUI` means **show all managed UI**, not **restore previous custom state**
- explicit snapshot/reset is **best effort**, especially for per-window restore when the set or order of open windows has changed
- the demo module’s window-level synchronization reads from `ActiveWindow`
- the demo-sheet builder performs a destructive rebuild of the `Demo` sheet
- reduced redraw does not fully eliminate Ribbon or non-client frame refresh flicker

---

## Design notes

- use centralized wrappers rather than scattering raw UI toggles throughout a project
- use selective tri-state control when only some elements should change
- use “hide all” / “show all” wrappers for consistent workbook presentation flows
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

## Status

This repository is intended as a reusable VBA component for Excel-based solutions on Windows.

---

## Author

Daniele Penza

---

## License

This project is licensed under the terms in `LICENSE`.
