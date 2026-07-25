# 🤝 Contributing to VBA-EXCEL_UI

<p align="left">
  <img alt="Contributions" src="https://img.shields.io/badge/Contributions-Welcome-217346">
  <img alt="Language" src="https://img.shields.io/badge/Language-Excel_VBA-00599C">
  <img alt="Platform" src="https://img.shields.io/badge/Platform-Windows-0078D6">
  <img alt="Style" src="https://img.shields.io/badge/Style-House_Conventions-6f42c1">
  <img alt="Tests" src="https://img.shields.io/badge/Tests-Regression_Harness-d97706">
  <img alt="License" src="https://img.shields.io/badge/License-MIT-2ea44f">
</p>

Thank you for your interest in improving **VBA-EXCEL_UI**.

This is a focused Excel interface-control project. Contributions are evaluated
primarily on whether they preserve or improve:

- explicit UI-state semantics;
- compatibility across supported Windows Excel environments;
- safe recovery after constrained-shell operations;
- predictable best-effort and structured-diagnostic behavior;
- clear ownership of application, window, Ribbon, and WinAPI effects;
- regression coverage for changed behavior;
- readability in the VBA Editor;
- backward compatibility of the established `UI_...` public surface.

A contribution is not complete merely because it compiles. It should also explain
the host behavior it changes, the Excel scope it affects, the failure policy, and
how the result was validated.

---

## 💬 Before you start

<p align="left">
  <img alt="Step" src="https://img.shields.io/badge/Step-Open_an_Issue_First-217346">
  <img alt="Scope" src="https://img.shields.io/badge/Scope-Agree_before_building-blue">
</p>

Open an issue before beginning non-trivial work so the intended API, platform
scope, UI ownership, and compatibility impact can be agreed in advance.

Use:

- [Bug report](.github/ISSUE_TEMPLATE/bug_report.md) for reproducible defects;
- [Feature request](.github/ISSUE_TEMPLATE/feature_request.md) for enhancements;
- [SECURITY.md](SECURITY.md) for suspected vulnerabilities.

Blank public issues are disabled so that reports contain enough environment,
state, recovery, and reproduction evidence to be actionable.

Good issues include:

- a reproducible failure on a specific Excel, Office-bitness, or Windows version;
- incorrect show, hide, or leave-unchanged behavior;
- a Ribbon or title-bar compatibility problem;
- a failure to restore `Application.ScreenUpdating`;
- a snapshot or reset defect;
- an incorrect structured failure count or failure message;
- a per-window restoration problem;
- a focused diagnostic, recovery, test, demo, or documentation improvement;
- a backward-compatible API extension with a clear use case.

Tiny corrections such as typographical fixes, broken links, comment corrections,
or obvious one-line defects may go directly to a pull request.

Suspected security vulnerabilities must be reported privately under
[SECURITY.md](SECURITY.md), not through a public issue.

All project interaction is governed by
[CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md).

---

## 🧭 Contribution priorities

The project particularly welcomes:

1. **Reproducible compatibility findings**  
   Include the exact Excel version, Office bitness, Windows version, workbook
   state, and UI element involved.

2. **Recovery improvements**  
   Changes that reduce the risk of leaving Excel in a constrained or damaged
   interface state.

3. **Identity-safe state restoration**  
   Improvements that preserve the correct relationship between captured
   per-window state and the window later restored.

4. **Structured diagnostics**  
   Clearer, machine-readable, and backward-compatible failure reporting.

5. **Regression coverage**  
   Permanent tests for corrected defects, including failure paths and host-state
   preservation.

6. **Documentation and demo quality**  
   Accurate examples, limitations, troubleshooting guidance, and reproducible
   demonstrations.

Feature count is not the primary objective. A small, explicit, well-tested
improvement is preferred to a broad change with ambiguous host behavior.

---

## 🧰 Project layout

<p align="left">
  <img alt="Source" src="https://img.shields.io/badge/Source-Exported_VBA-217346">
  <img alt="Demo" src="https://img.shields.io/badge/Demo-Versioned_XLSM-6f42c1">
  <img alt="Workflow" src="https://img.shields.io/badge/Workflow-Import_Edit_Export-0969da">
</p>

```text
VBA-EXCEL_UI/
├─ .gitattributes
├─ .github/
│  ├─ ISSUE_TEMPLATE/
│  │  ├─ bug_report.md
│  │  ├─ config.yml
│  │  └─ feature_request.md
│  └─ PULL_REQUEST_TEMPLATE.md
├─ .gitignore
├─ demo/
│  ├─ EXCEL_UI_DEMO.xlsm
│  ├─ M_DEMO_BUILDER.bas
│  └─ M_EXCEL_UI_DEMO.bas
├─ images/
├─ src/
│  └─ M_EXCEL_UI.bas
├─ test/
│  └─ M_EXCEL_UI_REGRESSION_TESTS.bas
├─ CODE_OF_CONDUCT.md
├─ CONTRIBUTING.md
├─ LICENSE
├─ README.md
└─ SECURITY.md
```

### Production source

```text
src/M_EXCEL_UI.bas
```

Owns:

- the public `UI_...` API;
- tri-state validation;
- application-level property control;
- per-window property control;
- Ribbon reads and writes;
- WinAPI title-bar handling;
- snapshot state;
- best-effort orchestration;
- Immediate Window logging;
- structured result accumulation.

### Demo source

```text
demo/M_EXCEL_UI_DEMO.bas
demo/M_DEMO_BUILDER.bas
```

Owns the worksheet demonstration and its builder workflow.

### Demo workbook

```text
demo/EXCEL_UI_DEMO.xlsm
```

The binary demo workbook is intentionally versioned because it is a user-facing
artifact. Binary changes receive special review requirements described below.

### Regression source

```text
test/M_EXCEL_UI_REGRESSION_TESTS.bas
```

Owns public test runners, assertions, state capture, restoration, and UI-specific
regression cases.

---


## 🧹 Repository hygiene and line endings

The repository includes:

```text
.gitignore
.gitattributes
```

### `.gitignore`

The ignore policy excludes:

- Office lock, temporary, backup, and recovery files;
- operating-system metadata;
- local editor and IDE settings;
- logs and diagnostic dumps;
- generated build and test-output folders;
- local virtual environments and caches;
- private-key and certificate-container files.

It deliberately does **not** ignore:

- exported `.bas`, `.cls`, `.frm`, or `.frx` source assets;
- the official `demo/EXCEL_UI_DEMO.xlsm` workbook;
- Markdown documentation;
- repository images.

Do not add a broad rule such as:

```text
*.xlsm
*.bas
images/
```

because it would hide intentional repository assets.

### `.gitattributes`

The attributes policy enforces:

```text
*.bas  → CRLF
*.cls  → CRLF
*.frm  → CRLF
```

This matches the Windows VBA Editor workflow and avoids line-ending churn across
contributors.

The following are explicitly binary:

- `.frx` UserForm companion files;
- Excel and Office workbooks;
- archives;
- images;
- PDFs.

Markdown and repository configuration files use LF.

When `.gitattributes` is introduced or materially changed, review the GitHub
Desktop diff carefully. A one-time renormalization may make unchanged files
appear modified if their prior line endings did not match the new policy. Do not
commit a broad normalization change accidentally as part of an unrelated pull
request.


## 🌿 Branch and pull-request workflow

### Branch from the correct baseline

For maintenance work, branch from the current release baseline or the branch
identified in the issue.

Examples:

```text
fix/titlebar-frame-refresh
fix/snapshot-window-identity
test/ribbon-failure-path
docs/recovery-guidance
feature/structured-reset-result
```

Use a release branch only when preparing a release:

```text
release/v1.0.1
```

Do not make routine development changes directly on `main`.

### Keep pull requests focused

GitHub automatically loads
[`.github/PULL_REQUEST_TEMPLATE.md`](.github/PULL_REQUEST_TEMPLATE.md) into new
pull requests after the template is merged into the default branch. Complete all
applicable sections rather than deleting required evidence.

A pull request should normally address one coherent concern:

- one defect;
- one API extension;
- one test expansion;
- one documentation improvement;
- one internal refactor with unchanged public behavior.

Avoid combining:

- public API changes;
- broad formatting changes;
- demo workbook redesign;
- unrelated refactoring;
- release metadata updates

unless the pull request is explicitly a coordinated release preparation.

### Recommended commit style

Use concise conventional prefixes where practical:

```text
fix: restore window state by captured identity
feat: add structured snapshot reset result
test: cover invalid UIVisibility values
refactor: extract WinAPI frame adapter
docs: clarify emergency recovery
chore: prepare v1.0.1
```

Commits should be small enough to review and revert independently.

---

## 🔁 Edit, compile, test, and export workflow

1. Confirm the current Git branch.

2. Import the required modules into a macro-enabled workbook.

   Recommended order:

   ```text
   src/M_EXCEL_UI.bas
   demo/M_EXCEL_UI_DEMO.bas          when needed
   demo/M_DEMO_BUILDER.bas           when needed
   test/M_EXCEL_UI_REGRESSION_TESTS.bas
   ```

3. Make the source change in the VBA Editor.

4. Compile with:

   ```text
   Debug → Compile VBAProject
   ```

5. Run the narrowest relevant regression runner.

6. Run the complete applicable regression sequence.

7. Perform the required manual UI-recovery checks.

8. Re-export each changed module over the matching repository file.

9. Review the textual diff in GitHub Desktop.

10. Update documentation and version metadata where applicable.

11. Commit and push the branch.

12. Open a pull request against the agreed base branch.

> [!IMPORTANT]
> The VBA project embedded in `demo/EXCEL_UI_DEMO.xlsm` does not update merely
> because an exported `.bas` file changes. When the demo workbook is part of the
> change, synchronize the embedded modules deliberately and verify the workbook
> independently.

---

## 🧱 Coding standards

<p align="left">
  <img alt="Option Explicit" src="https://img.shields.io/badge/Option_Explicit-Required-217346">
  <img alt="Private module" src="https://img.shields.io/badge/Option_Private_Module-Intentional-6f42c1">
  <img alt="Procedure headers" src="https://img.shields.io/badge/Procedure_Headers-Required-0969da">
</p>

### Module settings

- Every module must use `Option Explicit`.
- Production and test modules should retain `Option Private Module` unless an
  approved API-design change explicitly requires otherwise.
- Public members must be intentional and documented.
- New project-scoped helpers should normally be `Private`.
- Do not expose implementation details merely to make testing easier; prefer a
  narrow internal test seam when needed.

### Procedure headers

Every material procedure should use the established structured header style.

Include the fields relevant to the routine:

```text
PURPOSE
WHY THIS EXISTS
INPUTS
RETURNS
BEHAVIOR
ERROR POLICY
DEPENDENCIES
NOTES
CALLED FROM
UPDATED
```

Do not add headings that contain no useful information.

### Body structure

Use the sections a routine needs, generally in a clear execution order:

```text
DECLARE
INITIALIZE
VALIDATE INPUTS
CAPTURE STATE
APPLY
RESTORE
RETURN SUCCESS
SAFE EXIT
FAIL
```

Place explanatory comments above the code they describe. Use inline comments
primarily for declarations and short non-obvious constants.

### Naming

Preserve the established namespaces:

| Pattern | Intended scope |
|---|---|
| `UI_...` | Public production API or tightly related internal UI helpers |
| `Demo_...` | Public demonstration entry points |
| `TST_...` | Private regression-harness helpers |
| `Test_EXCEL_UI_...` | Public regression runners |

New public names require explicit review because they become part of the
compatibility contract.

### Formatting

- Preserve the repository’s indentation and line-continuation style.
- Prefer descriptive identifiers over abbreviations.
- Use `vbNullString` for intentional empty strings.
- Keep constants typed explicitly.
- Use `LongPtr` where required for handles and pointer-sized WinAPI values.
- Do not make broad whitespace-only changes in a behavioral pull request.
- Keep exported module attributes intact.

---

## 🎛️ Public API and compatibility contract

The established public surface includes:

```text
UIVisibility
UI_SetExcelUI
UI_SetExcelUI_WithResult
UI_HideExcelUI
UI_ShowExcelUI
UI_CaptureExcelUIState
UI_ResetExcelUIToSnapshot
UI_HasExcelUIStateSnapshot
UI_ClearExcelUIStateSnapshot
```

### Backward-compatible changes

Normally suitable for a minor release:

- adding a new optional procedure or helper;
- adding a new structured-result entry point;
- improving internal restoration logic;
- expanding tests;
- improving diagnostics without invalidating documented parsing assumptions;
- splitting internal implementation into additional private modules;
- adding an optional targeting mode while preserving current defaults.

### Breaking changes

Normally require a major release:

- removing or renaming a public member;
- changing enum values;
- reordering parameters in a public procedure;
- changing an existing optional parameter’s default meaning;
- changing `UI_ShowExcelUI` to mean restore rather than show all;
- changing fail-soft procedures to raise errors by default;
- changing global/window scope in a way that alters existing callers;
- removing `Option Private Module` to create a different exposure model.

State the expected Semantic Versioning impact in every pull request that changes
public behavior.

---

## 🧯 Error-handling contract

### Fire-and-forget procedures

The existing fire-and-forget procedures are deliberately fail-soft:

- ordinary element-level failures are not raised to callers;
- failures are written to the Immediate Window;
- processing continues with later requested elements where safe;
- quiet-update state is closed on handled exits.

Do not silently change this behavior in a patch or minor release.

### Structured-result procedure

`UI_SetExcelUI_WithResult` returns:

- `True` when no failure is recorded;
- `False` when one or more failures are recorded;
- `FailureCount` as the number of recorded failures;
- optional `FailureList` as a 1-based array preserving failure order.

When extending diagnostics:

- preserve deterministic ordering;
- keep messages concise and actionable;
- distinguish stage from detail;
- do not lose the original runtime error number and description;
- add a regression test for any machine-readable contract.

### Error handlers

- Use explicit local error handlers.
- Restore modified host state before returning where possible.
- Do not introduce modal `MsgBox` calls into the production module.
- Do not use `On Error Resume Next` across a broad procedure.
- Limit `On Error Resume Next` to narrow host-probing operations and restore the
  intended error handler immediately afterward.
- Avoid swallowing failures without logging or recording them.

---

## 🪟 WinAPI discipline

Title-bar handling is the highest-risk implementation area.

### Declarations

- Preserve VBA7 and pre-VBA7 conditional compilation where currently supported.
- Preserve 32-bit and 64-bit branches.
- Use `LongPtr` for window handles and pointer-sized style values under VBA7.
- Do not assume that a zero API return always means failure.
- Clear and read `GetLastError` according to the WinAPI contract.

### Style ownership

A contribution that changes window styles must document:

- the exact style bits read or written;
- whether the full style or only an owned mask is restored;
- behavior when another add-in changes the frame;
- behavior when `Application.Hwnd` changes;
- behavior in maximized, normal, and restored window states;
- the recovery path after failure.

### Frame refresh

Any style update must evaluate whether a non-client frame refresh is required.
Changes to `SetWindowPos` flags require explicit review and title-bar regression
testing.

### No dynamic code execution

Do not introduce:

- dynamically constructed VBA code;
- arbitrary `Shell` execution;
- external executable downloads;
- user-controlled WinAPI function names;
- user-controlled Excel 4 macro command strings.

The current Ribbon command is fixed and narrowly scoped. Any expansion of legacy
macro usage requires a security review.

---

## 📸 Snapshot and restore discipline

A snapshot contribution must define:

- which UI elements are captured;
- how unavailable reads are represented;
- whether the snapshot is complete or partial;
- how window identity is matched;
- what happens when windows are opened, closed, or reordered;
- what happens after a VBA project reset;
- whether the snapshot is cleared automatically or explicitly;
- how failures are reported.

Module-level snapshot state is intentionally in-memory. Do not imply durability
across Excel sessions unless a separately reviewed persistence design is added.

A restore operation must remain best effort unless a future major version defines
a different transactional contract.

---

## 🧪 Testing requirements

Import:

```text
test/M_EXCEL_UI_REGRESSION_TESTS.bas
```

Run:

```vb
Test_EXCEL_UI_RunCore
Test_EXCEL_UI_RunTitleBarOnly
Test_EXCEL_UI_RunAll
```

### Minimum requirements

- The VBA project compiles.
- All existing applicable tests pass.
- Changed behavior is covered by a focused regression case.
- A corrected defect receives a permanent named regression test.
- `Application.ScreenUpdating` is restored after the test.
- The Excel UI remains usable after completion.
- A manual emergency recovery test succeeds.

### Manual recovery tests

Run:

```vb
UI_HideExcelUI
UI_ShowExcelUI
```

Then run:

```vb
UI_CaptureExcelUIState
UI_HideExcelUI
UI_ResetExcelUIToSnapshot
```

Verify the interface after both sequences.

### Environment reporting

For Ribbon, WinAPI, multi-window, or compatibility changes, state:

- Excel product and version;
- Office 32-bit or 64-bit;
- Windows version;
- Excel window state;
- number of open workbook windows;
- whether other Excel add-ins were active;
- exact runner used;
- observed result.

### Failure-path testing

Where practical, cover:

- invalid enum values;
- failed Ribbon reads or writes;
- failed Boolean property reads or writes;
- invalid or changed window handles;
- title-bar style read/write failure;
- frame-refresh failure;
- changed window collection between capture and restore;
- unexpected failure while ScreenUpdating is suppressed;
- structured failure ordering and counts.

Do not weaken a valid test merely to make a changed implementation pass.

---

## 🖼️ Demo-workbook changes

The binary demo workbook is intentionally committed, but binary diffs are not
human-readable.

A pull request that changes `demo/EXCEL_UI_DEMO.xlsm` must also provide:

- the exported text module changes that explain the behavior;
- a description of the visible workbook change;
- screenshots when layout or controls change;
- confirmation that the workbook opens without repair warnings;
- confirmation that macros compile;
- confirmation that the demo controls work;
- confirmation that no personal, confidential, or client data is present;
- confirmation that no unintended external links, connections, names, or hidden
  content were introduced.

Do not change the binary workbook merely because Excel rewrote metadata. Include
it only when the user-facing artifact actually changed.

---

## 📚 Documentation expectations

Update documentation when a contribution changes:

- the public API;
- default behavior;
- managed UI scope;
- supported environments;
- error or diagnostic contracts;
- snapshot behavior;
- title-bar behavior;
- demo usage;
- regression runners;
- limitations or recovery guidance.

Depending on the change, update:

```text
README.md
CONTRIBUTING.md
SECURITY.md
.github/ISSUE_TEMPLATE/*
.github/PULL_REQUEST_TEMPLATE.md
Wiki pages
module and procedure headers
release notes
```

Documentation must distinguish:

- implemented behavior;
- tested behavior;
- best-effort behavior;
- planned roadmap items.

Avoid unsupported words such as “safe,” “complete,” “universal,” or
“production-ready” without a clearly stated scope and validation boundary.

---

## 🔐 Security-sensitive contributions

Read [SECURITY.md](SECURITY.md) before changing:

- WinAPI declarations or window-style handling;
- legacy Excel 4 macro calls;
- binary demo workbooks;
- external links, connections, or file access;
- macro-security guidance;
- persistence or configuration storage;
- any code that invokes the operating system.

Do not include vulnerability details in a public pull request before a coordinated
fix is ready.

---

## 📋 Pull-request checklist

Use this checklist in the pull-request description:

```text
[ ] The issue and intended scope are linked
[ ] The change is focused and reviewable
[ ] Public API and Semantic Versioning impact are stated
[ ] VBA project compiles
[ ] Relevant focused tests pass
[ ] Test_EXCEL_UI_RunAll passes where applicable
[ ] Manual show-all recovery succeeds
[ ] Snapshot/reset behavior is validated where applicable
[ ] ScreenUpdating is restored
[ ] Excel / Office / Windows environment is reported
[ ] Changed .bas modules were re-exported from the VBE
[ ] VBA files retain CRLF line endings under .gitattributes
[ ] Binary workbook and image files remain classified as binary
[ ] Text diffs were reviewed
[ ] Binary workbook changes are justified and validated
[ ] README / Wiki / headers are updated
[ ] No confidential or personal data is included
[ ] Security-sensitive changes received appropriate review
```

---

## 🚫 Changes that will normally be declined

- unexplained broad rewrites;
- changes that break the public API without a major-version plan;
- copied code without compatible licensing and attribution;
- dynamic code execution or unnecessary operating-system calls;
- arbitrary external dependencies;
- modal UI in the production controller;
- disabling macro-security protections;
- binary workbook changes without corresponding source and validation evidence;
- test removals or weakened assertions without a documented contract change;
- claims not supported by the implementation or tests;
- unrelated formatting churn;
- generated or AI-assisted changes that the contributor cannot explain and
  validate.

Use of development tools, including AI assistants, does not reduce the
contributor’s responsibility for correctness, licensing, security, and review.

---

## ⚖️ Review and acceptance

The maintainer may:

- request a smaller scope;
- ask for additional environment evidence;
- adapt naming or implementation to preserve project coherence;
- defer a feature to a later release;
- decline a contribution that does not fit the project boundary.

Acceptance is based on technical fit, maintainability, compatibility, and
verification—not only on whether the code appears to work in one workbook.

---

## 📄 License

By contributing, you agree that your contribution may be distributed under the
project’s [MIT License](LICENSE).

---

## 👤 Maintainer

Maintained by **Daniele Penza**.
