---
name: "🐞 Bug report"
about: "Report incorrect UI behavior, restoration, recovery, compatibility, or diagnostics"
title: "[Bug] "
labels: "bug"
assignees: ""
---

<!--
Thank you for helping improve VBA-EXCEL_UI.

Before submitting:
- Search open and closed issues for an existing report.
- Describe one observable defect per issue.
- Reduce the problem to the smallest reproducible workbook or VBA sample.
- Record the exact release tag or full 40-character commit SHA tested.
- Remove client data, credentials, workbook internals, and private material.

Suspected security vulnerabilities must not be reported publicly. Follow
SECURITY.md and use a private reporting channel.

Keep the headings below. Replace prompts with evidence. Use "N/A — reason" only
when a field genuinely does not apply. Screenshots support copyable evidence;
they do not replace it.
-->

## 🔍 Summary

<!--
Describe what fails, when it fails, and why it matters. State observable facts
before proposing a cause.
-->



## 🎯 Expected and actual behavior

**Expected**

<!-- Link the documented or public API contract when possible. -->



**Actual**

<!-- Include exact visible state, return values, errors, and recovery impact. -->



## 🔁 Minimal reproduction

**Frequency**

- [ ] Every time
- [ ] Intermittent — approximately <!-- e.g. 3 of 10 attempts -->
- [ ] Happened once

**Starting state**

<!--
Describe the state before step 1: active workbook/window, number of workbooks
and windows, managed UI visibility, snapshot availability, ScreenUpdating,
window state, and relevant add-ins.
-->



**Steps**

1. <!-- Begin from a clean Excel session or precisely described state. -->
2. <!-- Next action. -->
3. <!-- Action that exposes the defect. -->

**Minimal VBA**

<!--
Use a disposable workbook containing all four production modules from one exact
source state. Adapt this example to the smallest failing call. Do not overwrite
or clear a snapshot owned by another workflow.
-->

~~~vb
Option Explicit

Public Sub ReproduceIssue()
    Dim FailureCount As Long
    Dim FailureList As Variant
    Dim Succeeded As Boolean

    Succeeded = UI_SetExcelUI_WithResult( _
        Headings:=UI_Hide, _
        Gridlines:=UI_Hide, _
        FailureCount:=FailureCount, _
        FailureList:=FailureList, _
        TargetScope:=UI_TargetActiveWindow)

    Debug.Print "Succeeded="; Succeeded
    Debug.Print "FailureCount="; FailureCount

    'Use the deterministic show-all path after a constrained-state test.
    UI_ShowExcelUI
End Sub
~~~

**Reproduction cleanup**

<!--
Did normal cleanup complete? Was UI_ShowExcelUI required? Was the snapshot
restored or cleared? Did ScreenUpdating return to its original value? State
whether a VBA reset, End statement, Excel restart, or process termination was
required; each can hide the original cleanup defect.
-->



---

## 📦 Exact source identity

<!--
"Latest" is not reproducible. Use a release tag or the full output of
git rev-parse HEAD. If source was edited, attach or link a minimal sanitized
diff.
-->

| Field | Value |
|---|---|
| Release tag | <!-- e.g. v1.1.2 / N/A --> |
| Full 40-character commit SHA | <!-- required for branch or commit source --> |
| Branch | <!-- main / release branch / feature branch / N/A --> |
| Source origin | <!-- official release / repository / fork / copied files / other --> |
| Local modifications | <!-- none, or describe/link exact diff --> |
| Imported production files | <!-- list all four BAS filenames --> |
| Debug → Compile VBAProject | <!-- PASS / FAIL with exact error --> |
| Official demo asset used | <!-- filename/version / No --> |

Do not combine production modules from different tags, commits, branches, or
release assets.

---

## 💻 Host environment

| Field | Value |
|---|---|
| Excel product | <!-- Microsoft 365 / Excel 2021 / other --> |
| Excel version and full build | <!-- File → Account → About Excel --> |
| Office bitness | <!-- 32-bit / 64-bit --> |
| VBA generation | <!-- VBA7 / other --> |
| Windows version and build | |
| Workbook host | <!-- XLSM / XLSB / XLAM / PERSONAL.XLSB / other --> |
| Excel window state | <!-- normal / maximized / minimized / full screen --> |
| Open workbooks | <!-- count and sanitized names if relevant --> |
| Open Excel windows | <!-- count and ownership by workbook --> |
| Other relevant add-ins | <!-- none, or names/versions --> |
| Ribbon customization | <!-- none / custom Ribbon XML / other --> |
| Protected View or policy | <!-- relevant policy state / N/A --> |
| Regional or display settings | <!-- if relevant: language, DPI, monitors --> |

The source contains conditional 32-bit and 64-bit Office declarations.
Execution on one bitness does not certify the other.

---

## 🎛️ Affected surface and scope

**Managed UI**

- [ ] Ribbon
- [ ] Status Bar
- [ ] Scroll Bars
- [ ] Formula Bar
- [ ] Headings
- [ ] Workbook Tabs
- [ ] Gridlines
- [ ] Title Bar / native Excel frame

**Runtime or repository**

- [ ] Target-scope resolution
- [ ] Snapshot capture
- [ ] Snapshot restoration
- [ ] Snapshot ownership or clearing
- [ ] Window/hWnd identity
- [ ] Structured diagnostics
- [ ] ScreenUpdating preservation
- [ ] Emergency recovery
- [ ] Demo source or workbook
- [ ] Regression or certification harness
- [ ] Repository tooling, CI, documentation, or release evidence
- [ ] Other

**Observed reach**

- [ ] One workbook
- [ ] Multiple workbooks in one Excel process
- [ ] One Excel window
- [ ] Multiple Excel windows
- [ ] More than one Excel process
- [ ] Interaction with another macro or add-in
- [ ] Interaction with custom Ribbon XML
- [ ] Interaction with Protected View
- [ ] Interaction with full-screen mode
- [ ] Interaction with a VBA project reset

---

## 🪟 State and identity evidence

<!-- Complete the relevant rows. Preserve raw values where practical. -->

| Field | Before | Expected after | Actual after |
|---|---|---|---|
| Active workbook/window | | | |
| Application.Hwnd | | | |
| Relevant Window.hWnd | | | |
| TargetScope | | | |
| Ribbon | | | |
| Status Bar | | | |
| Scroll Bars | | | |
| Formula Bar | | | |
| Headings | | | |
| Workbook Tabs | | | |
| Gridlines | | | |
| Title Bar | | | |
| ScreenUpdating | | | |
| Snapshot available | | | |

**Identity or lifecycle transitions**

- [ ] Active window changed between capture and restore
- [ ] Captured window was closed
- [ ] Workbook window was recreated
- [ ] New window was opened after capture
- [ ] Native hWnd changed or may have been reused
- [ ] Snapshot was replaced by another capture
- [ ] Caller already owned a snapshot
- [ ] VBA project reset occurred
- [ ] None observed

Add the exact transition sequence:



---

## ⚠️ Error and structured-result evidence

<!--
Capture diagnostics immediately after the failing call. A later call can replace
or clear them. FailureCount is authoritative; FailureList is best effort.
-->

| Field | Value |
|---|---|
| Public procedure called | |
| Exact named arguments | |
| Returned Boolean | |
| FailureCount | |
| FailureList entries | <!-- preserve order and sanitize names --> |
| Immediate Window output | |
| Err.Number | <!-- decimal and/or hexadecimal / N/A --> |
| Err.Source | |
| Err.Description | |
| Visible partial changes | |

If FailureCount is non-zero while FailureList is shorter or unavailable, report
both without inferring that no failure occurred.

---

## 🆘 Recovery result

| Recovery action | Result |
|---|---|
| UI_ResetExcelUIToSnapshot | <!-- restored / partial / failed / no snapshot / not attempted --> |
| UI_ShowExcelUI | <!-- restored / partial / failed / not attempted --> |
| UI_ClearExcelUIStateSnapshot | <!-- completed / failed / not appropriate --> |
| VBA project reset | <!-- required / occurred earlier / no --> |
| Excel restart | <!-- required / no --> |
| Process termination | <!-- required / no --> |
| Remaining incorrect state | <!-- describe / none --> |

A defect that leaves Excel difficult to recover is higher priority than a
bounded cosmetic defect.

Do not run a snapshot-mutating self-test while application code owns the shared
snapshot slot.

---

## 🧪 Validation evidence

<!--
Static CI checks repository text. It does not launch Excel. Do not copy a
published result unless you ran it against the exact source identified above.
-->

| Check | Result |
|---|---|
| Debug → Compile VBAProject | <!-- PASS / FAIL / NOT RUN --> |
| python3 tools/check_repo.py | <!-- PASS / FAIL / NOT RUN --> |
| Test_EXCEL_UI_RunReleaseCertification | <!-- PASS / FAIL / INCOMPLETE / NOT RUN --> |
| Mandatory units / failures / skips / cleanup | <!-- exact counters/outcome --> |
| Tested source SHA | <!-- full SHA / N/A --> |
| Evidence text or JSON | <!-- sanitized attachment / filename / N/A --> |
| Targeted runner or case | <!-- exact name/result / N/A --> |
| Manual recovery check | <!-- PASS / FAIL / NOT RUN --> |

Useful narrower runners include:

~~~text
Test_EXCEL_UI_RunCore
Test_EXCEL_UI_RunTitleBarOnly
Test_EXCEL_UI_RunSnapshotIdentity
Test_EXCEL_UI_RunTitleBarSdiIdentity
Test_EXCEL_UI_RunAll
Test_EXCEL_UI_RunRibbonSdiProbe
~~~

The Ribbon SDI probe characterizes behavior; it does not certify a release.

---

## 📎 Logs, screenshots, and additional context

<!--
Paste short logs as text. Attach a minimal workbook only when plain-text VBA
cannot reproduce the issue. Explain relevant activation, window lifecycle,
display, remote-session, add-in, policy, or project-reset conditions.
-->



Do not attach:

- client, personal, or production data;
- credentials or connection strings;
- proprietary VBA unrelated to the defect;
- private reviews or audit material;
- unsanitized certification output;
- weaponized or security-sensitive proof-of-concept files.

---

## ✅ Reporter checklist

- [ ] I searched open and closed issues for this problem
- [ ] I reported one focused defect
- [ ] I supplied an exact release tag or full commit SHA, not latest
- [ ] I listed the exact production modules and any local modifications
- [ ] I described the starting UI, window, and snapshot state
- [ ] I provided a minimal repeatable example where possible
- [ ] I captured return values, failures, and errors before another call changed them
- [ ] I documented recovery and remaining Excel state
- [ ] I separated static checks from Excel runtime evidence
- [ ] I removed credentials, client data, private material, and sensitive names
- [ ] This is not a suspected security vulnerability; private reports follow SECURITY.md
