# Excel UI

VBA utilities to control the Excel UI shell, including Ribbon, status bar, formula bar, gridlines, workbook tabs, and WinAPI-based title-bar visibility.

## Overview

This repository contains a VBA module for managing Excel UI visibility through a safe tri-state API. It combines:

- Excel object-model UI control
- WinAPI-based title-bar control for the Excel window represented by `Application.Hwnd`

The module is designed for workbook-driven solutions that need a constrained, kiosk-like, or application-style Excel shell.

## Features

- Show, hide, or leave unchanged each managed UI element
- Best-effort processing so one failed UI element does not stop the rest
- Windows-only title-bar control via WinAPI
- 32-bit and 64-bit Office/VBA compatibility
- Simple convenience wrappers for hide-all and show-all behavior

## Managed UI Elements

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

- Title Bar (via WinAPI)

## Public API

### Enum: `K_UIVisibility`

```vb
Public Enum K_UIVisibility
    K_UI_LeaveUnchanged = -1
    K_UI_Hide = 0
    K_UI_Show = 1
End Enum
```

### Procedures

- `K_SetExcelUI`
- `K_HideExcelUI`
- `K_ShowExcelUI`

## Example Usage

### Hide everything managed by the module

```vb
K_HideExcelUI
```

### Show everything managed by the module

```vb
K_ShowExcelUI
```

### Selective control

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

## Design Notes

- The API uses a tri-state enum to avoid the ambiguity of Boolean optional arguments.
- `K_ShowExcelUI` means **show all managed UI**. It does **not** restore a previously captured user-specific UI state.
- The title-bar logic targets the Excel window represented by `Application.Hwnd`.
- Ribbon control relies on `Application.ExecuteExcel4Macro`.

## Requirements

- Microsoft Excel for Windows
- VBA7-compatible Office supported, including 32-bit and 64-bit builds
- WinAPI access available in the host environment

## Limitations

- Windows only
- No snapshot/restore of prior UI state
- Title-bar control is implemented through WinAPI, not through the Excel object model

## Suggested Repository Contents

```text
/README.md
/src/EXCEL_UI.bas
```

## Importing into VBA

1. Open the VBA editor.
2. Export or download the module file from this repository.
3. Import the `.bas` module into your VBA project.
4. Call `K_SetExcelUI`, `K_HideExcelUI`, or `K_ShowExcelUI` as needed.

## Status

This repository is intended as a reusable VBA utility component for Excel-based solutions.
