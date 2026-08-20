# Contributing to VBA-EXCEL_UI

Thank you for improving VBA-EXCEL_UI.

The project prioritizes:

- backward compatibility of the documented `UI_...` public surface;
- safe recovery after constrained-shell operations;
- identity-safe Window restoration;
- explicit state ownership;
- deterministic diagnostics;
- 32-bit and 64-bit Windows Excel compatibility;
- permanent regression coverage;
- readable exported VBA source.

Read [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md) and [SECURITY.md](SECURITY.md) before contributing. Open an issue before non-trivial work.

## Project layout

```text
VBA-EXCEL_UI/
├─ src/
│  ├─ M_EXCEL_UI.bas
│  ├─ M_EXCEL_UI_RUNTIME.bas
│  ├─ M_EXCEL_UI_SNAPSHOT.bas
│  └─ M_EXCEL_UI_TITLEBAR.bas
├─ demo/
│  ├─ M_DEMO_BUILDER.bas
│  └─ M_EXCEL_UI_DEMO.bas
├─ test/
│  └─ M_EXCEL_UI_REGRESSION_TESTS.bas
├─ INSTALLATION.md
├─ README.md
└─ ...
```

## Production module boundaries

### `M_EXCEL_UI`

Owns:

- the supported public `UI_...` API;
- `UIVisibility`;
- general tri-state validation;
- general application/window apply orchestration;
- compatibility wrappers.

It must remain the facade. Internal implementation details should not be moved into caller code.

### `M_EXCEL_UI_RUNTIME`

Owns shared low-level services used by the facade and snapshot engine:

- ordered failure accumulation;
- Immediate Window diagnostic formatting;
- Ribbon reads and writes;
- generic Boolean property reads and writes;
- `ScreenUpdating` quiet-update scopes.

It must not depend on another project module.

### `M_EXCEL_UI_SNAPSHOT`

Owns:

- all mutable snapshot state;
- retained Excel `Window` references;
- capture and restoration orchestration;
- identity resolution and missing-window diagnostics.

It may depend on runtime and title-bar services. It must not duplicate snapshot state in the facade.

### `M_EXCEL_UI_TITLEBAR`

Owns:

- Win32/Win64 declarations;
- exact title-bar style-bit ownership;
- handle-specific captured style state;
- style merging and non-client frame refresh.

It must not depend on another project module.

## Dependency rule

The production dependency graph must remain acyclic:

```text
M_EXCEL_UI
├── M_EXCEL_UI_RUNTIME
├── M_EXCEL_UI_TITLEBAR
└── M_EXCEL_UI_SNAPSHOT
    ├── M_EXCEL_UI_RUNTIME
    └── M_EXCEL_UI_TITLEBAR
```

Do not introduce:

- a callback from runtime or title-bar modules into the facade;
- a second copy of mutable snapshot or title-bar state;
- generic utility modules with unrelated responsibilities;
- public implementation details solely to simplify testing.

## Branch workflow

Do not make routine development changes directly on `main`.

Use a branch appropriate to the agreed scope, for example:

```text
fix/snapshot-window-identity
test/titlebar-owned-bits
docs/install-modules
ci/static-release-gate
feature/window-target-scope
release/v1.1.1
```

Confirm the current branch in GitHub Desktop before every commit.

## Import, edit, compile, test, export

Recommended import order:

```text
src/M_EXCEL_UI_RUNTIME.bas
src/M_EXCEL_UI_TITLEBAR.bas
src/M_EXCEL_UI_SNAPSHOT.bas
src/M_EXCEL_UI.bas
test/M_EXCEL_UI_REGRESSION_TESTS.bas
demo/M_EXCEL_UI_DEMO.bas          when needed
demo/M_DEMO_BUILDER.bas           when needed
```

Workflow:

1. Confirm the current branch.
2. Import the required modules into a controlled workbook.
3. Compile with `Debug → Compile VBAProject`.
4. Run the narrowest focused runner.
5. Run the complete applicable sequence.
6. Perform manual recovery.
7. Re-export each changed module over the matching repository path.
8. Review the GitHub Desktop diff.
9. Update documentation and version metadata.
10. Commit and push.
11. Open a pull request against the agreed base.

The demo workbook is not version-controlled. It is built from the exported demo modules and published as a GitHub Release asset, so changing an exported `.bas` file does not update any committed binary. Rebuild and validate the workbook separately when it is in release scope.

## Required validation

For production code changes:

```text
Debug → Compile VBAProject
Test_EXCEL_UI_RunCore
Test_EXCEL_UI_RunTitleBarOnly
Test_EXCEL_UI_RunSnapshotIdentity
Test_EXCEL_UI_RunAll
UI_HideExcelUI / UI_ShowExcelUI
capture / hide / reset
```

Record only environments actually tested.

## Public API compatibility

Preserve unless an explicitly approved breaking release requires otherwise:

- public procedure and function names;
- parameter order;
- optional defaults;
- enum values;
- show/hide/leave-unchanged semantics;
- fire-and-forget behavior;
- ordered structured-result behavior;
- application and window scope;
- `UI_ShowExcelUI` as the emergency show-all operation.

New backward-compatible parameters must normally be optional and trailing.

## Snapshot changes

Document and test:

- captured state;
- partial capture behavior;
- Window identity strategy;
- new, missing, closed, or recreated windows;
- failure ordering;
- behavior after VBA reset;
- emergency recovery.

Never restore per-window state by collection index.

## Title-bar changes

Document and test:

- exact style bits owned;
- 32-bit and 64-bit declarations;
- valid zero WinAPI returns;
- `GetLastError` treatment;
- `Application.Hwnd` changes;
- frame refresh;
- preservation of unrelated current style bits.

Do not restore an entire stale `GWL_STYLE` value.

## Error policy

Production entry points are fail-soft unless the public contract explicitly says otherwise.

- Continue after unrelated element-level failures.
- Do not silently discard failures.
- Restore `ScreenUpdating`.
- Do not introduce unsolicited `MsgBox` calls.
- Keep `On Error Resume Next` scopes narrow.
- Preserve insertion order in structured diagnostics.

## Source style

Every module must use `Option Explicit`.

Production and test modules should retain `Option Private Module` unless an approved API change requires otherwise.

Use the established structured comment style with relevant sections such as:

```text
PURPOSE
WHY THIS EXISTS
INPUTS
RETURNS
BEHAVIOR
ERROR POLICY
DEPENDENCIES
NOTES
UPDATED
```

Prefer explicit procedure boundaries and cohesive modules over clever abstraction.

## Repository hygiene

- Exported `.bas`, `.cls`, and `.frm` files use CRLF.
- Markdown and repository configuration use LF.
- Workbooks, images, `.frx`, archives, and PDFs are binary.
- Do not commit Office lock files, backups, logs, credentials, client data, or unrelated formatting churn.
- Review every binary workbook change separately.

## Pull requests

Keep a pull request focused on one coherent concern.

Include:

- related issue;
- public behavior and Semantic Versioning assessment;
- affected UI scope;
- state ownership;
- diagnostics and failure policy;
- validation environment;
- focused and full test evidence;
- recovery evidence;
- documentation changes;
- binary demo impact.

A refactor is complete only when the public API still compiles and behaves identically under the regression harness.
