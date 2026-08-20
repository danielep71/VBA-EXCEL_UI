---
name: 🐞 Bug report
about: Report incorrect UI behavior, restoration failure, compatibility defects, or crashes
title: "[Bug]: "
labels: bug
---

## 🐞 Description

Describe the defect clearly and explain its practical impact.

> [!IMPORTANT]
> Do not use this public template for a suspected security vulnerability.
> Follow the private reporting process in `SECURITY.md`.

## 🔖 Version and source state

Identify the exact source tested.

```text
Release tag:     <e.g. v1.0.1, or N/A>
Commit SHA:      <full 40-character SHA if using main or another snapshot>
Branch:          <main / release branch / feature branch / N/A>
Source obtained: <official repository / tagged release archive / other>
```

Do not write only “latest.”

## 🎛️ Affected UI surface

Check all that apply:

- [ ] Ribbon
- [ ] Status Bar
- [ ] Scroll Bars
- [ ] Formula Bar
- [ ] Headings
- [ ] Workbook Tabs
- [ ] Gridlines
- [ ] Title Bar / Excel main-window frame
- [ ] Snapshot capture
- [ ] Snapshot reset / restoration
- [ ] Structured diagnostics
- [ ] `ScreenUpdating` preservation
- [ ] Demo worksheet or controls
- [ ] Regression harness
- [ ] Other

## 🔢 Exact call and observed result

Provide the smallest exact call that reproduces the issue.

```vba
Option Explicit

Public Sub ReproduceIssue()

    Dim FailureCount As Long
    Dim FailureList  As Variant
    Dim OK           As Boolean

    OK = UI_SetExcelUI_WithResult( _
            Ribbon:=UI_LeaveUnchanged, _
            StatusBar:=UI_LeaveUnchanged, _
            ScrollBars:=UI_LeaveUnchanged, _
            FormulaBar:=UI_LeaveUnchanged, _
            Headings:=UI_LeaveUnchanged, _
            WorkbookTabs:=UI_LeaveUnchanged, _
            Gridlines:=UI_LeaveUnchanged, _
            TitleBar:=UI_LeaveUnchanged, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

    Debug.Print "Succeeded: "; OK
    Debug.Print "FailureCount: "; FailureCount

End Sub
```

Adapt the example to the actual problem.

```text
Call:
Returned Boolean:
FailureCount:
FailureList:
Immediate Window output:
Runtime error, if any:
```

## 🖥️ UI state before and after

Describe the relevant state precisely.

```text
Before:
Expected after:
Observed after:
```

Where relevant, include:

- which workbook and window were active;
- number and order of open Excel windows;
- whether Excel was maximized, minimized, or in normal state;
- whether the Ribbon or title bar had already been changed by another macro or add-in;
- whether an explicit snapshot existed.

## 🔁 Steps to reproduce

1.
2.
3.

State whether the defect reproduces:

```text
Always / intermittently / once only
```

## 🧪 Environment

```text
Excel product:       <Microsoft 365 / Excel 2021 / other>
Excel version:
Excel build:
Office bitness:      32-bit / 64-bit
VBA version:         VBA7 / legacy VBA
Windows version:
Workbook type:       .xlsm / .xlsb / .xlam / other
Excel window state:  maximized / normal / minimized
Open Excel windows:
Other add-ins active:
```

## ✅ Regression-harness result

Run the relevant procedures when practical.

```text
Debug → Compile VBAProject             →
Test_EXCEL_UI_RunReleaseCertification  →
```

If certification cannot be run, the narrower runners still help:

```text
Test_EXCEL_UI_RunCore               →
Test_EXCEL_UI_RunTitleBarOnly       →
Test_EXCEL_UI_RunSnapshotIdentity   →
Test_EXCEL_UI_RunAll                →
```

Paste the relevant Immediate Window output. Certification also writes a JSON
document and a text report to `%TEMP%`, either of which can be attached.

## 🆘 Recovery result

State whether Excel could be restored using:

```vb
UI_ShowExcelUI
```

and, where relevant:

```vb
UI_ResetExcelUIToSnapshot
```

```text
UI_ShowExcelUI restored the interface:       Yes / No / Not tested
Snapshot reset restored the interface:       Yes / No / Not applicable
Excel restart required:                      Yes / No
Other recovery action:
```

A failure that leaves Excel difficult to recover is higher priority than a
cosmetic defect.

## 🔀 Scope and interaction

Check any that apply:

- [ ] One workbook only
- [ ] All workbooks in the current Excel process
- [ ] One Excel window
- [ ] Multiple Excel windows
- [ ] More than one Excel process
- [ ] Interaction with another add-in
- [ ] Interaction with custom Ribbon XML
- [ ] Interaction with Protected View
- [ ] Interaction with full-screen mode
- [ ] Interaction with a VBA project reset

## 📎 Additional context

Add sanitized screenshots, logs, or links that help reproduce the issue.

Do not attach workbooks containing confidential, personal, client, credential,
or production data. Create a minimal sanitized reproduction instead.
