# Installing VBA Excel UI

<div align="center">

**Source-first deployment for Windows Excel VBA**

[![Excel VBA](https://img.shields.io/badge/Excel_VBA-32%20%2F%2064--bit-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://github.com/danielep71/VBA-EXCEL_UI)
[![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)](#requirements)
[![Modules](https://img.shields.io/badge/Production_Modules-4-c2185b?style=for-the-badge)](#production-package)

[Repository](https://github.com/danielep71/VBA-EXCEL_UI)
·
[Releases](https://github.com/danielep71/VBA-EXCEL_UI/releases)
·
[Public API](#public-api)
·
[Validation](#validation)
·
[Recovery](#emergency-recovery)

</div>

---

> [!IMPORTANT]
> VBA Excel UI is a **four-module production package**. Import all four source
> modules from the same release. Importing only M_EXCEL_UI.bas is not a valid
> installation.

> [!NOTE]
> No installer, COM registration, administrator access, third-party DLL, or
> non-standard VBA reference is required. The component uses the VBA project
> itself plus standard Windows APIs supplied by the operating system.

## Contents

- [Choose a deployment model](#choose-a-deployment-model)
- [Requirements](#requirements)
- [Production package](#production-package)
- [Install in a workbook or add-in](#install-in-a-workbook-or-add-in)
- [First smoke test](#first-smoke-test)
- [Public API](#public-api)
- [Target scopes](#target-scopes)
- [Snapshot ownership](#snapshot-ownership)
- [Current release boundaries](#current-release-boundaries)
- [Validation](#validation)
- [Emergency recovery](#emergency-recovery)
- [Upgrade](#upgrade)
- [Troubleshooting](#troubleshooting)
- [Remove the component](#remove-the-component)

---

## Choose a deployment model

### Embed the source in your VBA project

This is the normal production deployment.

Use it when:

- a workbook or add-in owns its VBA source;
- the UI controller must travel with that file;
- source review and deterministic versioning matter;
- no shared machine-level installation is wanted.

Import the four production modules, call the public facade in M_EXCEL_UI, and
save the host as a macro-enabled Office file.

### Use the release demo

Use the demo to evaluate the component before integrating it.

The macro-enabled demo workbook is distributed as a **GitHub Release asset**.
It is not maintained as a binary file in the source tree. Download the demo
from the release matching the source version you intend to use.

Do not use the demo modules as production dependencies.

---

## Requirements

### Supported source platform

| Requirement | Value |
|---|---|
| Host | Microsoft Excel for Windows |
| VBA generation | VBA7 |
| Office bitness | 32-bit or 64-bit source paths are present |
| File type | XLSM, XLAM, XLSB, or another VBA-capable Excel host |
| References | Standard Excel/VBA references only |
| Windows APIs | user32 and kernel32 |
| Administrator rights | Not required |

The title-bar module contains conditional declarations for 32-bit and 64-bit
Office. That is a **source compatibility claim**, not proof that every bitness,
Office channel, Windows build, and Excel SDI configuration has been exercised
for a particular release.

Always consult the release notes and certification evidence for the exact
version you deploy.

### Unsupported environments

- Excel for macOS;
- Excel for the web;
- hosts without VBA7;
- protected or managed environments that prohibit macros or the required
  Windows API calls.

### Macro security

Use your organization’s normal trusted-location, signing, or macro-security
policy. Do not instruct users to weaken global Office security settings.

For distributed workbooks or add-ins, a signed VBA project and a documented
publisher trust process are recommended.

---

## Production package

Import these four files from the same tag or release:

| Order | Source file | VBA module name | Responsibility |
|---:|---|---|---|
| 1 | src/M_EXCEL_UI_RUNTIME.bas | M_EXCEL_UI_RUNTIME | Shared execution, diagnostics, result buffers, and host-state handling |
| 2 | src/M_EXCEL_UI_TITLEBAR.bas | M_EXCEL_UI_TITLEBAR | Bitness-safe title-bar WinAPI access and frame-state management |
| 3 | src/M_EXCEL_UI_SNAPSHOT.bas | M_EXCEL_UI_SNAPSHOT | In-memory capture, snapshot identity, lifecycle, and restoration |
| 4 | src/M_EXCEL_UI.bas | M_EXCEL_UI | Stable public facade and public enums |

The import order is recommended for clarity. VBA resolves project modules as a
project, so compile success—not import order—is the actual dependency check.

### Optional development files

| File | Purpose | Production dependency |
|---|---|:---:|
| test/M_EXCEL_UI_REGRESSION_TESTS.bas | Regression and release-certification runners | No |
| demo/M_EXCEL_UI_DEMO.bas | Demo behavior | No |
| demo/M_DEMO_BUILDER.bas | Builds the distributed demo workbook | No |

Keep test and demo modules out of a production workbook unless you explicitly
need them.

### Runtime architecture

~~~mermaid
flowchart TD
    A[Workbook or add-in] --> B[M_EXCEL_UI public facade]
    B --> C[M_EXCEL_UI_RUNTIME]
    B --> D[M_EXCEL_UI_SNAPSHOT]
    B --> E[M_EXCEL_UI_TITLEBAR]
    D --> C
    D --> E
    E --> C
~~~

Application code should call M_EXCEL_UI. The other modules are implementation
dependencies and may expose procedures only so standard VBA modules can
collaborate and the regression harness can inspect controlled seams.

---

## Install in a workbook or add-in

### 1. Download one exact version

Prefer a tagged release:

1. Open the project’s
   [Releases page](https://github.com/danielep71/VBA-EXCEL_UI/releases).
2. Select the required version.
3. Download that release’s source archive.
4. Keep all four production modules from that same archive.

Do not mix modules from different commits, tags, branches, or release assets.

If you work from a commit rather than a tag, record the full commit SHA in your
own build or deployment notes.

### 2. Back up the destination

Before changing an existing VBA project:

- save and close other Excel files that are not needed;
- create a recoverable copy of the destination workbook or add-in;
- export any existing modules that will be replaced.

### 3. Open the Visual Basic Editor

In Excel:

1. Open the destination file.
2. Press **Alt+F11**.
3. In Project Explorer, select the correct VBA project.
4. If Project Explorer is hidden, press **Ctrl+R**.

Confirm the project you selected belongs to the intended workbook or add-in.

### 4. Import all four production modules

For each production BAS file:

1. Choose **File > Import File**.
2. Select the BAS file.
3. Confirm the expected module appears under **Modules**.

The final project must contain:

- M_EXCEL_UI_RUNTIME
- M_EXCEL_UI_TITLEBAR
- M_EXCEL_UI_SNAPSHOT
- M_EXCEL_UI

If a module with the same name already exists, follow the
[upgrade procedure](#upgrade) instead of creating duplicate or auto-renamed
modules.

### 5. Compile the complete VBA project

Choose **Debug > Compile VBAProject**.

Compilation is mandatory. It catches:

- missing modules;
- duplicate public names;
- unsupported declarations;
- accidental module renaming;
- broken host-project references;
- syntax damage introduced during manual copying.

If **Compile VBAProject** is disabled, the project is already compiled or no
compile-relevant change is pending. Make a harmless edit and undo it if you
need to force the command to become available.

### 6. Save in a macro-capable format

Use a file type that preserves VBA, such as:

- Excel Macro-Enabled Workbook, XLSM;
- Excel Binary Workbook, XLSB;
- Excel Add-In, XLAM.

Saving as XLSX removes the VBA project.

### 7. Add a recovery entry point

Place a small macro in a module owned by your workbook or add-in:

~~~vb
Public Sub RestoreExcelShell()
    UI_ShowExcelUI
End Sub
~~~

This provides an obvious manual recovery command if application logic stops
while the Excel shell is constrained.

---

## First smoke test

Run smoke tests in a disposable workbook or a recoverable copy. Start with an
element that does not depend on native title-bar mutation.

### Structured selective test

~~~vb
Public Sub SmokeTestExcelUI()
    Dim failureCount As Long
    Dim failureList As Variant
    Dim succeeded As Boolean

    succeeded = UI_SetExcelUI_WithResult( _
        StatusBar:=UI_Hide, _
        FailureCount:=failureCount, _
        FailureList:=failureList)

    Debug.Print "Hide succeeded: "; succeeded
    Debug.Print "Failure count: "; failureCount

    succeeded = UI_SetExcelUI_WithResult( _
        StatusBar:=UI_Show, _
        FailureCount:=failureCount, _
        FailureList:=failureList)

    Debug.Print "Show succeeded: "; succeeded
    Debug.Print "Failure count: "; failureCount
End Sub
~~~

Expected result:

- the status bar hides and then shows;
- both calls return True;
- FailureCount is zero.

### Snapshot round-trip test

Only run this example when the caller does not already own a component
snapshot. The component has one project-level snapshot slot.

~~~vb
Public Sub SmokeTestSnapshotRoundTrip()
    Dim failureCount As Long
    Dim failureList As Variant

    If UI_HasExcelUIStateSnapshot Then
        Debug.Print "Snapshot already exists; test skipped."
        Exit Sub
    End If

    If Not UI_CaptureExcelUIState_WithResult( _
            FailureCount:=failureCount, _
            FailureList:=failureList) Then
        Debug.Print "Capture reported failures: "; failureCount
    End If

    Call UI_SetExcelUI( _
        Headings:=UI_Hide, _
        Gridlines:=UI_Hide, _
        TargetScope:=UI_TargetActiveWindow)

    If Not UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=failureCount, _
            FailureList:=failureList) Then
        Debug.Print "Restore reported failures: "; failureCount
    End If

    UI_ClearExcelUIStateSnapshot
End Sub
~~~

Do not use a test helper that blindly captures and clears the shared snapshot
inside a caller-owned workflow.

---

## Public API

Production callers should use only M_EXCEL_UI.

### Visibility values

| Value | Meaning |
|---|---|
| UI_LeaveUnchanged | Do not touch this element |
| UI_Hide | Request hidden state |
| UI_Show | Request visible state |

Omitted visibility arguments are equivalent to UI_LeaveUnchanged. This makes
selective calls explicit and prevents omitted arguments from becoming an
accidental hide request.

### Primary entry points

| Procedure | Use |
|---|---|
| UI_SetExcelUI | Selectively change managed elements; fire-and-forget diagnostics |
| UI_SetExcelUI_WithResult | Selectively change elements and receive structured success/failure data |
| UI_HideExcelUI | Apply the component’s deterministic hide-all preset |
| UI_ShowExcelUI | Apply the deterministic show-all recovery preset |
| UI_CaptureExcelUIState | Capture the current managed state |
| UI_CaptureExcelUIState_WithResult | Capture and receive structured diagnostics |
| UI_HasExcelUIStateSnapshot | Test whether a snapshot slot is currently populated |
| UI_ResetExcelUIToSnapshot | Restore the retained snapshot |
| UI_ResetExcelUIToSnapshot_WithResult | Restore and receive structured diagnostics |
| UI_ClearExcelUIStateSnapshot | Release the retained snapshot and object references |

### Structured-result contract

The WithResult procedures:

- return True only when no failure was recorded;
- return False when one or more failures were recorded;
- write the authoritative number of failures to FailureCount;
- optionally return a 1-based String array in FailureList;
- use ordered Stage | Detail text entries for individual failures;
- continue best-effort where the operation permits it.

Treat FailureCount as authoritative. FailureList is a convenience buffer and
its allocation or growth can itself fail under resource pressure. A False
return must not be interpreted as “nothing happened”; inspect the diagnostics
and the resulting Excel state.

### Selective example

~~~vb
Public Sub EnterApplicationView()
    Dim failureCount As Long
    Dim failureList As Variant

    If Not UI_SetExcelUI_WithResult( _
            Ribbon:=UI_Hide, _
            StatusBar:=UI_Hide, _
            FormulaBar:=UI_Hide, _
            Headings:=UI_Hide, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Hide, _
            TargetScope:=UI_TargetActiveWorkbookWindows, _
            FailureCount:=failureCount, _
            FailureList:=failureList) Then

        Debug.Print "UI update failures: "; failureCount
    End If
End Sub
~~~

Ribbon, Status Bar, Scroll Bars, and Formula Bar are not controlled by
TargetScope. Title-bar mutation is associated with the Excel application frame
used by the current implementation. Headings, Workbook Tabs, and Gridlines use
TargetScope.

---

## Target scopes

| Scope | Window-level targets |
|---|---|
| UI_TargetAllExcelWindows | Every current Excel window |
| UI_TargetActiveWindow | Application.ActiveWindow only |
| UI_TargetActiveWorkbookWindows | Windows belonging to ActiveWorkbook |

TargetScope controls:

- headings;
- workbook tabs;
- gridlines.

It does not make application-level properties window-local.

### Scope guidance

Use **UI_TargetActiveWindow** when a command deliberately affects only the
window the user is working in.

Use **UI_TargetActiveWorkbookWindows** when every current view of one workbook
must be consistent without changing unrelated workbooks.

Use **UI_TargetAllExcelWindows** for a deliberate Excel-instance-wide mode.
This remains the default for backward compatibility.

Excel uses an SDI-style window model in current Windows versions. Test
multi-workbook and multi-window behavior that matters to your host project;
single-window testing is not sufficient evidence for window identity or
targeting behavior.

---

## Snapshot ownership

The component maintains one in-memory snapshot slot per loaded VBA project.

### Lifecycle

1. Capture before applying a temporary UI mode.
2. Check the structured capture result.
3. Apply the required UI changes.
4. Restore while the captured windows still exist.
5. Inspect the structured restore result.
6. Clear the snapshot when it is no longer needed.

### Ownership rules

- A new capture replaces the previous snapshot.
- Restore retains the snapshot so a caller can inspect, retry, or decide when
  to clear it.
- Clear releases retained Excel Window references.
- A partial capture can still leave a snapshot available.
- Newly opened windows are not part of an earlier snapshot.
- Closed or recreated windows can make full restoration impossible.
- Callers sharing the same VBA project must coordinate ownership of the single
  snapshot slot.

Use UI_HasExcelUIStateSnapshot before an optional self-test or helper captures
state. A helper must not overwrite or clear a snapshot owned by its caller.

### Recommended host pattern

~~~vb
Public Sub BeginManagedView()
    Dim failureCount As Long
    Dim failureList As Variant

    If UI_HasExcelUIStateSnapshot Then
        Err.Raise vbObjectError + 2000, _
                  "BeginManagedView", _
                  "A UI snapshot is already owned by another workflow."
    End If

    If Not UI_CaptureExcelUIState_WithResult( _
            FailureCount:=failureCount, _
            FailureList:=failureList) Then
        Debug.Print "Capture was partial: "; failureCount
    End If

    Call UI_HideExcelUI
End Sub

Public Sub EndManagedView()
    Dim failureCount As Long
    Dim failureList As Variant

    If UI_HasExcelUIStateSnapshot Then
        Call UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=failureCount, _
            FailureList:=failureList)
        UI_ClearExcelUIStateSnapshot
    Else
        UI_ShowExcelUI
    End If
End Sub
~~~

Adapt the error policy to your host. The important parts are explicit ownership,
structured restoration, and a show-all fallback.

---

## Current release boundaries

The following boundaries apply to the current v1.1.2 baseline and are tracked
for the v1.1.3 correctness-and-hardening release. Read them before relying on
snapshot restoration or release certification in a safety-critical workflow.

### Runtime correctness

| Area | v1.1.2 boundary | v1.1.3 branch state |
|---|---|---|
| Ribbon restore | A changed active window can make restore target the wrong window instead of failing closed. | Open: [#23](https://github.com/danielep71/VBA-EXCEL_UI/issues/23) |
| Title-bar identity | Snapshot restore does not pair the retained Excel Window identity with its native hWnd strongly enough for every SDI recreation case. | Open: [#45](https://github.com/danielep71/VBA-EXCEL_UI/issues/45) |
| Recycled hWnd | A recycled native handle can collide with a stored same-style title-bar registry entry. | Reopened: [#32](https://github.com/danielep71/VBA-EXCEL_UI/issues/32) |
| Captionless baseline | Showing a title bar from a non-zero captionless baseline can report success without restoring the caption. | Open: [#6](https://github.com/danielep71/VBA-EXCEL_UI/issues/6) |
| Self-test ownership | The tagged self-test can clear a snapshot it has just refused because the caller owns it. | Corrected on the release branch: [#43](https://github.com/danielep71/VBA-EXCEL_UI/issues/43) |
| Quiet-update ownership | A suppressed or ignored write can be recorded as an achieved transition. | Corrected on the release branch: [#26](https://github.com/danielep71/VBA-EXCEL_UI/issues/26) |

Until the relevant corrections are released and certified:

- keep UI_ShowExcelUI available as the emergency recovery path;
- do not treat a successful structured result as stronger than the known
  limits of the deployed version;
- exercise active-window changes, closed/recreated windows, and title-bar
  recovery in a disposable Excel session;
- on v1.1.2, do not run a destructive self-test while application code owns the
  shared snapshot; the repaired release-branch runners refuse before mutation
  and have regression coverage for caller-owned snapshot preservation.

### Assurance and release evidence

Repository checks and the current regression runners are useful, but they are
not by themselves complete proof of every supported runtime configuration.
The v1.1.3 release work includes:

- full cleanup proof
  ([#35](https://github.com/danielep71/VBA-EXCEL_UI/issues/35));
- a mandatory certification case inventory
  ([#42](https://github.com/danielep71/VBA-EXCEL_UI/issues/42));
- exact-source certification evidence
  ([#46](https://github.com/danielep71/VBA-EXCEL_UI/issues/46));
- a full public-API contract gate, implemented on the branch and still awaiting
  the final release-diff evidence
  ([#47](https://github.com/danielep71/VBA-EXCEL_UI/issues/47));
- accurate release closure documentation
  ([#48](https://github.com/danielep71/VBA-EXCEL_UI/issues/48));
- an exact-head review and certification gate
  ([#49](https://github.com/danielep71/VBA-EXCEL_UI/issues/49)).

Certification claims should identify the exact source commit, the runner,
Excel version/build, Office bitness, Windows version, relevant window
configuration, outcome, and retained evidence. Never infer that a closed issue
or a passing static check proves the final release source.

Private assessments and audit material are not installation artifacts and must
not be published without explicit authorization. Public issues, release notes,
and approved certification evidence must remain self-contained.

---

## Validation

Validation has three layers. Use the layer appropriate to the deployment, and
do not confuse one layer with another.

### Consumer integration gate

For every workbook or add-in:

- import the four production modules from one exact version;
- compile the complete destination VBA project;
- run the low-risk structured smoke test;
- confirm the emergency recovery macro is available;
- test the host’s own entry and exit workflows;
- reopen the saved file and repeat the critical path.

### Developer regression gate

For source changes, import the optional regression module and use the relevant
runner:

| Runner | Purpose |
|---|---|
| Test_EXCEL_UI_RunCore | Core regression pack without title-bar mutation |
| Test_EXCEL_UI_RunTitleBarOnly | Title-bar-focused pack |
| Test_EXCEL_UI_RunSnapshotIdentity | Snapshot window-identity cases |
| Test_EXCEL_UI_RunTitleBarSdiIdentity | Multi-window SDI title-bar identity cases |
| Test_EXCEL_UI_RunAll | Broad developer regression pack |
| Test_EXCEL_UI_RunRibbonSdiProbe | Ribbon SDI characterization only; not release certification |
| Test_EXCEL_UI_RunReleaseCertification | Release-oriented certification runner |

Run destructive or shell-mutating tests only in a controlled Excel session.
Close unrelated workbooks, save work first, and preserve the full Immediate
Window or generated evidence required by the runner.

The Ribbon SDI probe characterizes behavior; it does not certify automatic
Ribbon activation or introduce that future feature.

### Repository static gate

From the repository root:

~~~text
python tools/check_repo.py
~~~

The repository checker validates source hygiene and repository invariants. It
cannot execute Excel, Windows APIs, SDI transitions, or VBA runtime cleanup.
A passing static gate is required evidence, not runtime certification.

### Release-candidate gate

Before a release is tagged:

1. Freeze the exact candidate commit.
2. Run the repository gate on that commit.
3. Export or build the VBA artifacts from that exact source.
4. Run the mandatory runtime case inventory.
5. Record environment and result evidence.
6. Confirm the evidence identifies the exact candidate SHA.
7. Re-run any gate invalidated by a later source or documentation change.
8. Tag only the exact reviewed and certified commit.

Do not certify one tree and release another.

### Runtime evidence

The v1.1.2 release notes record runtime evidence for:

- Excel 16.0 build 20131;
- Windows 64-bit NT 10.00;
- Office x64;
- VBA7.

That evidence does not automatically certify Office x86 or every other
Windows, Office, language, update-channel, DPI, and multi-window combination.

---

## Emergency recovery

### Preferred recovery

Run:

~~~vb
UI_ShowExcelUI
~~~

This is the deterministic show-all recovery preset. It does not require a
snapshot and is therefore the correct fallback when no trustworthy baseline
exists.

### Snapshot recovery

If the current workflow owns a valid snapshot:

~~~vb
Public Sub RestoreOwnedSnapshot()
    Dim failureCount As Long
    Dim failureList As Variant

    If Not UI_HasExcelUIStateSnapshot Then
        UI_ShowExcelUI
        Exit Sub
    End If

    If Not UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=failureCount, _
            FailureList:=failureList) Then
        Debug.Print "Restore failures: "; failureCount
        UI_ShowExcelUI
    End If

    UI_ClearExcelUIStateSnapshot
End Sub
~~~

For v1.1.2, apply the known Ribbon and title-bar boundaries described above.

### If the VBA project is stopped

1. Press **Alt+F8**.
2. Run your RestoreExcelShell macro.
3. If macros cannot run, close the affected Excel instance after saving any
   recoverable work.
4. Reopen Excel without automatically re-entering the constrained UI mode.

Design Workbook_Open, Auto_Open, and add-in startup code so recovery is still
possible when initialization fails halfway through.

---

## Upgrade

### Safe replacement procedure

1. Read the target release notes and migration information.
2. Back up the destination file.
3. Export any local modifications to the existing component modules.
4. Open the destination in the Visual Basic Editor.
5. Remove the four old production modules.
6. Import all four new production modules from the same exact release.
7. Replace the regression module too if you use it.
8. Choose **Debug > Compile VBAProject**.
9. Run the consumer integration gate.
10. Run the developer or release gates required by your use case.
11. Save, close, reopen, and verify the host workflow.

Replacing the four modules as a unit avoids mixed internal contracts.

### Preserve host code, not component forks

Keep workbook-specific orchestration in separate host modules. This makes a
future component upgrade a clean four-module replacement.

If you must modify component source:

- record the upstream release and commit;
- keep the patch in version control;
- re-run the full relevant validation;
- do not claim upstream certification for the modified tree.

### Upgrade from v1.1.2 to v1.1.3

v1.1.3 is a correctness-and-hardening release. Replace the complete production
package and regression module, then validate against the final exact v1.1.3
source and published evidence.

Automatic Ribbon activation and the rebuilt demo belong to v1.2.0. Do not
expect those behavioral features in v1.1.3.

---

## Troubleshooting

### Compile error: Sub or Function not defined

Most often, one of the four production modules is missing or came from a
different version.

Check module names and replace the complete package from one exact release.

### Compile error in a Windows declaration

Confirm:

- the host is Excel for Windows;
- VBA7 is available;
- the BAS file was imported rather than partially copied;
- the full M_EXCEL_UI_TITLEBAR module came from the selected release.

### Ambiguous name detected

The project probably contains:

- duplicate imports;
- an auto-renamed copy such as M_EXCEL_UI1;
- host code using the same public procedure or enum name.

Remove or rename the conflicting code, then compile again.

### UI call returns False

Inspect FailureCount and FailureList. A best-effort operation can change some
elements while another element fails. Use UI_ShowExcelUI when safe recovery is
more important than preserving a partial state.

### FailureList is empty but FailureCount is non-zero

Treat FailureCount as authoritative. The optional detail buffer is best-effort
and may be unavailable even though failures were counted.

### Snapshot capture returned False

A partial snapshot may still exist. Check UI_HasExcelUIStateSnapshot, inspect
the failure details, and decide whether the captured subset is sufficient.
Do not assume False means the snapshot slot is empty.

### Snapshot restore returned False

Common causes include:

- a captured window was closed;
- a window was recreated;
- an element could not be written;
- no snapshot existed;
- a known identity or native-frame boundary applies to the deployed version.

Inspect the result, run UI_ShowExcelUI when appropriate, and clear the snapshot
only when its owner has finished with it.

### Title bar does not change

Confirm:

- Excel is running on Windows;
- the VBA project compiled for the installed Office bitness;
- policy or endpoint protection is not blocking the API call;
- the returned diagnostics;
- the known title-bar boundaries for the deployed release.

Test title-bar behavior in a disposable session before integrating it into
automatic startup code.

### Ribbon result is unexpected after switching windows

Ribbon visibility is associated with the active window’s Ribbon command
context. v1.1.2 has a tracked wrong-target restore defect. Use a controlled
single-window recovery path or UI_ShowExcelUI, and upgrade after the correction
is released and certified.

### Excel remains constrained after an unhandled error

Add error cleanup to the host workflow and keep an independent
RestoreExcelShell macro. Do not rely exclusively on snapshot restoration.

### Changes disappear after saving

The file was probably saved in a non-macro format. Save as XLSM, XLSB, or XLAM.

---

## Demo release asset

The distributed demo workbook is a convenience artifact, not the authoritative
source.

Before using it:

- download it from the release matching your source version;
- verify the release and asset identity;
- treat its macros with the same security review as other downloaded VBA;
- use the source modules as the basis for code review and integration.

The demo builder and demo source under the demo directory are development
inputs. A locally generated demo must not be represented as an official release
asset unless it was built, checked, and published through the release process.

---

## Repository text and export hygiene

The repository’s .gitattributes and .editorconfig define text normalization and
editor behavior for source and documentation files.

For VBA modules:

- preserve the repository’s expected encoding and line-ending policy;
- import and export with a process that does not silently rewrite declarations,
  attributes, or non-ASCII text;
- review diffs after any Visual Basic Editor export;
- run the repository checker before committing.

Do not store secrets, credentials, private audit material, or sensitive
workbook data in source, test evidence, demo assets, or documentation.

---

## Remove the component

### From a workbook or add-in

1. Run UI_ShowExcelUI.
2. If your workflow owns a snapshot, restore or clear it deliberately.
3. Remove host calls to the public API.
4. Remove the four production modules.
5. Remove the optional regression and demo modules, if present.
6. Choose **Debug > Compile VBAProject**.
7. Save, close, and reopen the host.
8. Confirm the Excel shell starts in the expected state.

There is no machine-level uninstaller because the component creates no
machine-level installation.

### Before deleting an add-in

Disable or remove it through Excel’s add-in management first. Confirm its
startup or shutdown code cannot reapply a constrained UI state.

---

## Installation checklist

### Production integration

- [ ] Destination is backed up.
- [ ] All four modules came from one exact release or commit.
- [ ] Module names match the production package.
- [ ] Complete VBA project compiles.
- [ ] File is saved in a macro-capable format.
- [ ] Low-risk structured smoke test passes.
- [ ] Host entry, error, and exit paths are tested.
- [ ] UI_ShowExcelUI recovery macro is available.
- [ ] Snapshot ownership is explicit.
- [ ] Relevant v1.1.2 or later release boundaries were reviewed.
- [ ] Saved file was closed, reopened, and retested.

### Release or redistributed build

- [ ] Exact source SHA is recorded.
- [ ] Static repository gate passes on that SHA.
- [ ] Runtime cases match the mandatory inventory.
- [ ] Office, Windows, bitness, Excel build, and window setup are recorded.
- [ ] Evidence came from the exact candidate source.
- [ ] Full cleanup and recovery behavior were exercised.
- [ ] Public API contracts were checked.
- [ ] Documentation matches the final source.
- [ ] Private material is excluded.
- [ ] The tag points to the reviewed and certified commit.

---

## Installation principle

> Import one complete four-module version, compile the destination project,
> validate the behavior you rely on, preserve an independent show-all recovery
> path, and tie every assurance claim to the final exact source.
