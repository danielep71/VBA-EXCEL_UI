# Installation and Upgrade Guide

This guide applies to the modular v1.1.0 production package.

> [!IMPORTANT]
> `M_EXCEL_UI.bas` is the public facade, not a standalone implementation.
> Every production installation requires all four files in `src/`.

> [!NOTE]
> The ready-to-run macro-enabled demo workbook is not versioned in Git.
> Tested `.xlsm` demo builds are distributed only as GitHub Release assets.

## Required production modules

| Recommended import order | Repository path | VBA module name | Responsibility |
|---:|---|---|---|
| 1 | `src/M_EXCEL_UI_RUNTIME.bas` | `M_EXCEL_UI_RUNTIME` | Shared host operations, diagnostics, result buffers, quiet-update scope |
| 2 | `src/M_EXCEL_UI_TITLEBAR.bas` | `M_EXCEL_UI_TITLEBAR` | WinAPI declarations, owned title-bar bits, frame refresh |
| 3 | `src/M_EXCEL_UI_SNAPSHOT.bas` | `M_EXCEL_UI_SNAPSHOT` | Snapshot state, capture, restore, retained Window identities |
| 4 | `src/M_EXCEL_UI.bas` | `M_EXCEL_UI` | Public `UI_...` facade, targeting, and general apply orchestration |

The import order is recommended for clarity. VBA resolves project-level references after all modules are present and the project is compiled.

## Optional source modules

| Path | Required for production | Purpose |
|---|:---:|---|
| `test/M_EXCEL_UI_REGRESSION_TESTS.bas` | No | Regression and release validation |
| `demo/M_EXCEL_UI_DEMO.bas` | No | Demo actions |
| `demo/M_DEMO_BUILDER.bas` | No | Demo worksheet construction |

## Binary demo distribution

The repository does not version `demo/EXCEL_UI_DEMO.xlsm`.

For a tagged release, the validated workbook should be attached to the GitHub Release, preferably with a versioned name such as:

```text
EXCEL_UI_DEMO_v1.1.0.xlsm
```

A SHA-256 checksum may also be published in the release notes.

This keeps source control text-focused while still giving end users a ready-to-run workbook.

## Dependency graph

```text
M_EXCEL_UI
├── M_EXCEL_UI_RUNTIME
├── M_EXCEL_UI_TITLEBAR
└── M_EXCEL_UI_SNAPSHOT
    ├── M_EXCEL_UI_RUNTIME
    └── M_EXCEL_UI_TITLEBAR
```

`M_EXCEL_UI_RUNTIME` and `M_EXCEL_UI_TITLEBAR` do not depend on another project module. The graph is deliberately acyclic.

## Fresh installation

1. Open the destination macro-enabled workbook or add-in.
2. Open the VBA Editor with `Alt+F11`.
3. Save a backup before changing the VBA project.
4. Import the four production modules in dependency-first order.
5. Confirm the Project Explorer contains exactly:

   ```text
   M_EXCEL_UI
   M_EXCEL_UI_RUNTIME
   M_EXCEL_UI_SNAPSHOT
   M_EXCEL_UI_TITLEBAR
   ```

6. Run:

   ```text
   Debug → Compile VBAProject
   ```

7. Perform a recovery round-trip:

   ```vb
   UI_HideExcelUI
   UI_ShowExcelUI
   ```

8. Capture and restore a baseline:

   ```vb
   UI_CaptureExcelUIState
   UI_HideExcelUI
   UI_ResetExcelUIToSnapshot
   ```

## Upgrade from v1.0.1 or another single-module installation

The old `M_EXCEL_UI` module contains internal implementation that is now distributed across four modules.

1. Back up the workbook or add-in.
2. In the VBA Editor, remove the existing `M_EXCEL_UI`.
3. Export a backup if required.
4. Remove any experimental modules named:

   ```text
   M_EXCEL_UI_RUNTIME
   M_EXCEL_UI_SNAPSHOT
   M_EXCEL_UI_TITLEBAR
   ```

5. Import the four current files from `src/`.
6. Compile the project.
7. Run the validation sequence below.

Do not paste the new facade over the old module while leaving old private helpers in place. That can create duplicate procedure names or mixed state.

## Upgrade from an intermediate v1.1.0 development build

Replace the complete production set together:

```text
M_EXCEL_UI_RUNTIME
M_EXCEL_UI_TITLEBAR
M_EXCEL_UI_SNAPSHOT
M_EXCEL_UI
```

This avoids combining a newer facade with an older internal module.

## Public targeting API

Version 1.1.0 adds:

```vb
Public Enum UIWindowTargetScope
    UI_TargetAllExcelWindows = 0
    UI_TargetActiveWindow = 1
    UI_TargetActiveWorkbookWindows = 2
End Enum
```

`TargetScope` is an optional trailing argument on the selective apply APIs.

The default:

```vb
UI_TargetAllExcelWindows
```

preserves existing behavior.

Targeting affects only:

```text
Headings
Workbook Tabs
Gridlines
```

Application-level elements and the title bar retain their established scope.

Example:

```vb
UI_SetExcelUI _
    Headings:=UI_Hide, _
    Gridlines:=UI_Hide, _
    TargetScope:=UI_TargetActiveWindow
```

Active workbook windows:

```vb
UI_SetExcelUI _
    WorkbookTabs:=UI_Show, _
    Gridlines:=UI_Show, _
    TargetScope:=UI_TargetActiveWorkbookWindows
```

## Validation sequence

With the optional regression module imported:

```text
Debug → Compile VBAProject
Test_EXCEL_UI_RunCore
Test_EXCEL_UI_RunTitleBarOnly
Test_EXCEL_UI_RunSnapshotIdentity
Test_EXCEL_UI_RunAll
UI_HideExcelUI
UI_ShowExcelUI
```

The regression harness includes active-window, active-workbook-window, and invalid-target-scope cases.

The tests manipulate the real Excel UI. Run them in a controlled Excel instance.

## Public integration

Call only the documented facade members from workbook or add-in code:

```vb
UI_SetExcelUI
UI_SetExcelUI_WithResult
UI_HideExcelUI
UI_ShowExcelUI
UI_CaptureExcelUIState
UI_CaptureExcelUIState_WithResult
UI_ResetExcelUIToSnapshot
UI_ResetExcelUIToSnapshot_WithResult
UI_HasExcelUIStateSnapshot
UI_ClearExcelUIStateSnapshot
```

Use these public enums:

```vb
UIVisibility
UIWindowTargetScope
```

Do not call `UI_Runtime...`, `UI_Snapshot...`, or title-bar worker routines from normal application code. Those procedures are internal implementation seams.

## Module settings

The production modules use:

```vb
Option Explicit
Option Private Module
```

`Option Private Module` keeps project members available inside the containing VBA project while preventing normal cross-project automation exposure.

## Common compilation problems

### Sub or Function not defined

Cause: one or more required production modules are missing or an old module version is mixed with the current facade.

Resolution: replace all four production modules from the same release or commit.

### Ambiguous name detected

Cause: a module was imported twice, or old helpers remain in another module.

Resolution: remove duplicate or experimental modules, then re-import the complete package.

### Invalid target-scope diagnostic

If `UI_SetExcelUI_WithResult` reports a `TargetScope` failure, verify that the value is one of:

```text
UI_TargetAllExcelWindows
UI_TargetActiveWindow
UI_TargetActiveWorkbookWindows
```

Application-level operations can continue under the fail-soft contract, while unsafe window-level writes are suppressed.

### Expected module name differs from file name

The authoritative VBA module name is the `Attribute VB_Name` value. After import, Project Explorer must show the four exact production module names listed above.

### Title bar does not change

Confirm:

- Excel desktop for Windows is being used;
- organizational policy permits WinAPI calls;
- no add-in is immediately rewriting the same window-style bits;
- the project compiles for the installed Office bitness.

Title-bar behavior is best effort and can vary by Excel/Windows environment.

### Excel UI remains hidden after an interrupted test

Run:

```vb
UI_ShowExcelUI
```

This recovery path does not require a snapshot.

## Building the demo release asset

For a release candidate:

1. Start from the exact release branch/tag source.
2. Import the four production modules.
3. Import:

   ```text
   demo/M_EXCEL_UI_DEMO.bas
   demo/M_DEMO_BUILDER.bas
   ```

4. Compile the VBA project.
5. Run the regression and manual recovery checks.
6. Save the workbook outside the Git-tracked source tree as:

   ```text
   EXCEL_UI_DEMO_v1.1.0.xlsm
   ```

7. Optionally calculate its SHA-256.
8. Attach the workbook to the GitHub Release.

The binary workbook should not be committed back into the repository.

## Line endings and repository work

Exported `.bas` files are expected to use CRLF under `.gitattributes`. Markdown and repository configuration files should remain LF.

The demo `.xlsm` is ignored by repository policy and belongs only in release packaging.

## Removing the component

1. Call `UI_ShowExcelUI`.
2. Clear any snapshot if the project is still running:

   ```vb
   UI_ClearExcelUIStateSnapshot
   ```

3. Remove all four production modules.
4. Remove optional demo/test modules when they are no longer needed.
5. Compile the remaining VBA project.
