<!--
Thank you for contributing to VBA-EXCEL_UI.

Keep the pull request focused: one logical UI behavior, recovery, API, testing,
documentation, or repository-maintenance change per PR.

Read CONTRIBUTING.md, CODE_OF_CONDUCT.md, and SECURITY.md before completing this
template.
-->

## 📋 Summary

Describe the change and why it is needed.

## 🔗 Related issue

```text
Closes #
```

Remove this section when no issue applies.

## 🧩 Type of change

- [ ] 🐞 Functional or compatibility bug fix
- [ ] 🆘 Recovery or host-state preservation fix
- [ ] ✨ Backward-compatible feature
- [ ] ♻️ Refactor with no intended public behavior change
- [ ] 🧪 Regression-harness or test change
- [ ] 🖼️ Demo workbook or demo-builder change
- [ ] 📚 Documentation-only change
- [ ] 🔒 Security-related change
- [ ] 🔧 Repository or release-maintenance change

## 🎛️ Affected UI surface

Check all that apply:

- [ ] Ribbon
- [ ] Status Bar
- [ ] Scroll Bars
- [ ] Formula Bar
- [ ] Headings
- [ ] Workbook Tabs
- [ ] Gridlines
- [ ] Title Bar / WinAPI frame
- [ ] Snapshot capture
- [ ] Snapshot restoration
- [ ] Structured diagnostics
- [ ] `ScreenUpdating`
- [ ] Demo
- [ ] Tests
- [ ] Repository configuration only

## 📐 Public API and Semantic Versioning

Describe changes to:

- public procedure or function names;
- signatures or parameter order;
- optional-parameter defaults;
- enum values;
- show / hide / leave-unchanged semantics;
- fire-and-forget behavior;
- structured-result behavior;
- application or window targeting;
- snapshot meaning;
- recovery behavior.

```text
Public behavior changed:
Backward compatible:
Suggested release: patch / minor / major
Migration required:
```

Write `No public behavior change` where applicable.

## 🧭 Scope and state ownership

Describe which Excel surface is affected:

```text
Application-level:
Window-level:
Excel main window:
All open windows / active window / specified target:
```

Explain:

- whether the change affects the current Excel process or one workbook;
- what state this module owns;
- how the change interacts with other macros or add-ins;
- how prior state is captured or preserved;
- what happens when the target window collection changes.

## 🪟 Ribbon or WinAPI method

For Ribbon or title-bar changes, explain:

```text
API or command used:
Style bits read or written:
32-bit path:
64-bit path:
GetLastError treatment:
Frame-refresh treatment:
Application.Hwnd treatment:
Interaction with other add-ins:
```

Write `Not applicable` when the change does not touch these areas.

## 📸 Snapshot and recovery behavior

For snapshot, reset, or constrained-shell changes, explain:

```text
Captured state:
Complete or partial capture:
Window identity strategy:
Behavior after VBA reset:
Behavior when windows change:
Emergency recovery path:
```

Confirm whether `UI_ShowExcelUI` remains a usable fallback.

## 📋 Diagnostics and failure policy

Describe:

- whether failures are logged, returned, or raised;
- whether best-effort continuation is preserved;
- whether failure ordering changes;
- whether new machine-readable categories are introduced;
- how unexpected runtime errors are represented;
- how modified host state is restored after failure.

```text
Failure contract:
Logging contract:
Structured-result contract:
```

## 🧪 Testing performed

```text
Debug → Compile VBAProject       →
Test_EXCEL_UI_RunCore            →
Test_EXCEL_UI_RunTitleBarOnly    →
Test_EXCEL_UI_RunAll             →
Manual hide/show recovery        →
Manual capture/hide/reset        →
```

Include relevant Immediate Window output for a functional fix or substantial
behavioral change.

## 💻 Validation environment

```text
Excel product:
Excel version:
Excel build:
Office bitness:
Windows version:
Workbook type:
Excel window state:
Open Excel windows:
Other add-ins active:
```

List every environment actually tested. Do not imply untested environments were
validated.

## ✅ Contract checklist

### Source and compilation

- [ ] The VBA project compiles cleanly.
- [ ] Changed modules were re-exported to the correct repository paths.
- [ ] Exported VBA files retain CRLF line endings under `.gitattributes`.
- [ ] The textual diff contains only intended changes.
- [ ] No Office lock, backup, editor-state, or generated file is included.
- [ ] No confidential, personal, credential, client, or production data is included.

### Public behavior and compatibility

- [ ] Existing public names and signatures remain compatible, or the breaking
      change and major-version rationale are explicit.
- [ ] Existing enum values retain their meaning.
- [ ] `UI_ShowExcelUI` still means show all unless a deliberate major-version
      contract change is proposed.
- [ ] Application-level and window-level scope is documented.
- [ ] New public members use the established `UI_...` namespace.
- [ ] `Option Private Module` exposure remains intentional.

### Error and host-state behavior

- [ ] Best-effort continuation remains deliberate and documented.
- [ ] Failures are not silently discarded.
- [ ] `Application.ScreenUpdating` is restored.
- [ ] No production procedure introduces an unsolicited `MsgBox`.
- [ ] Narrow `On Error Resume Next` blocks restore the intended handler promptly.
- [ ] Emergency recovery remains available.
- [ ] Invalid `UIVisibility` values remain controlled.

### WinAPI and Ribbon safety

- [ ] 32-bit and 64-bit declarations are correct.
- [ ] Handles and pointer-sized values use appropriate types.
- [ ] Valid zero WinAPI returns are distinguished from failures.
- [ ] Style ownership and restoration are documented.
- [ ] Required frame refresh is performed.
- [ ] No user-controlled dynamic macro command or arbitrary OS execution is added.

### Snapshot behavior

- [ ] Snapshot completeness and unavailable reads are handled explicitly.
- [ ] Window identity or index behavior is documented and tested.
- [ ] Changes in window count or order are covered where relevant.
- [ ] In-memory lifetime and VBA-reset behavior remain accurate.
- [ ] Reset-without-snapshot behavior remains controlled.

### Tests and documentation

- [ ] A corrected defect has a permanent regression test.
- [ ] Focused and complete applicable suites pass.
- [ ] Manual recovery was tested.
- [ ] README and Wiki pages were updated where relevant.
- [ ] `CONTRIBUTING.md`, `SECURITY.md`, or templates were updated when project
      process or risk boundaries changed.
- [ ] Version metadata and release notes are synchronized where applicable.

### Binary demo workbook

- [ ] No binary workbook change is included.
- [ ] Or: the binary change is intentional, described, and accompanied by source.
- [ ] The workbook opens without repair warnings.
- [ ] The embedded VBA project compiles.
- [ ] Screenshots are supplied for visible layout changes.
- [ ] No unintended external links, connections, names, or hidden data were added.

## 📚 Documentation updated

Check all that apply:

- [ ] README
- [ ] Wiki
- [ ] Module or procedure headers
- [ ] CONTRIBUTING
- [ ] SECURITY
- [ ] Issue or pull-request templates
- [ ] Demo guidance
- [ ] Release notes
- [ ] No documentation change required

## 📎 Reviewer notes

Describe trade-offs, known limitations, unresolved environment coverage, or
follow-up work.
