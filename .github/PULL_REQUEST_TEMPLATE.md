## Summary

Describe the change and why it is needed.

## Related issue

```text
Closes #
```

## Type of change

- [ ] Functional or compatibility fix
- [ ] Recovery or host-state fix
- [ ] Backward-compatible feature
- [ ] Internal refactor with no intended public behavior change
- [ ] Regression-test change
- [ ] Demo change
- [ ] Documentation-only change
- [ ] Repository or release maintenance
- [ ] Security-related change

## Affected surface

- [ ] Ribbon
- [ ] Status Bar
- [ ] Scroll Bars
- [ ] Formula Bar
- [ ] Headings
- [ ] Workbook Tabs
- [ ] Gridlines
- [ ] Title Bar / WinAPI
- [ ] Snapshot capture
- [ ] Snapshot restoration
- [ ] Structured diagnostics
- [ ] `ScreenUpdating`
- [ ] Module dependencies
- [ ] Demo
- [ ] Tests
- [ ] Documentation only

## Public API and Semantic Versioning

```text
Public behavior changed:
Backward compatible:
Suggested release: patch / minor / major
Migration required:
```

Confirm changes to names, signatures, parameter order/defaults, enum values, targeting, snapshot meaning, diagnostics, and recovery. Write `No public behavior change` where applicable.

## Module ownership and dependencies

```text
M_EXCEL_UI:
M_EXCEL_UI_RUNTIME:
M_EXCEL_UI_SNAPSHOT:
M_EXCEL_UI_TITLEBAR:
Dependency graph changed:
Circular dependency introduced: no / explain
Mutable state duplicated: no / explain
```

## Snapshot and recovery

```text
Captured state:
Window identity strategy:
Behavior for new windows:
Behavior for missing/closed/recreated windows:
Behavior after VBA reset:
Failure ordering:
Emergency recovery path:
```

## Ribbon or WinAPI method

```text
API or command used:
Owned style bits:
32-bit path:
64-bit path:
GetLastError treatment:
Frame refresh:
Application.Hwnd treatment:
Unrelated style bits preserved:
```

Write `Not applicable` when appropriate.

## Diagnostics and failure policy

```text
Failure contract:
Logging contract:
Structured-result contract:
ScreenUpdating restoration:
```

## Testing performed

```text
Debug → Compile VBAProject             →
python3 tools/check_repo.py            →
Test_EXCEL_UI_RunReleaseCertification  →
Manual UI_HideExcelUI / UI_ShowExcelUI →
Manual capture / hide / reset          →
```

Paste the certification verdict line:

```text
RESULT:
```

`INCOMPLETE`, a non-zero `skipped` count, or `cleanup=FAILED` is not a pass.

Narrower runners used while iterating, if any:

```text
Test_EXCEL_UI_RunAll                  →
Test_EXCEL_UI_RunCore                 →
Test_EXCEL_UI_RunTitleBarOnly         →
Test_EXCEL_UI_RunSnapshotIdentity     →
Test_EXCEL_UI_RunTitleBarSdiIdentity  →
```

## Validation environment

```text
Excel product/version/build:
Office bitness:
Windows version:
Workbook type:
Excel window state:
Open Excel windows:
Other add-ins:
```

List only environments actually tested.

## Source checklist

- [ ] Current branch was confirmed before committing.
- [ ] All four required production modules were present during compilation.
- [ ] Changed modules were exported to the correct repository paths.
- [ ] CRLF was preserved for exported VBA source.
- [ ] No conflict markers or duplicate procedures remain.
- [ ] The textual diff contains only intended changes.
- [ ] No lock, backup, generated, confidential, credential, client, or production-data file is included.

## Compatibility checklist

- [ ] Existing public names, signatures, defaults, and enum values remain compatible, or the breaking rationale is explicit.
- [ ] `UI_ShowExcelUI` remains an emergency show-all path.
- [ ] Best-effort continuation remains deliberate.
- [ ] Failures are not silently discarded.
- [ ] `ScreenUpdating` is restored.
- [ ] No unsolicited production `MsgBox` was introduced.
- [ ] Invalid enum values remain controlled.

## Module-boundary checklist

- [ ] `M_EXCEL_UI` remains the public facade.
- [ ] Runtime and title-bar modules have no project-module dependencies.
- [ ] Snapshot state exists only in `M_EXCEL_UI_SNAPSHOT`.
- [ ] Title-bar mutable state exists only in `M_EXCEL_UI_TITLEBAR`.
- [ ] No circular dependency was introduced.
- [ ] Internal modules retain `Option Private Module`.

## Snapshot checklist

- [ ] Per-window restore does not use collection index.
- [ ] Retained Window identity behavior is documented and tested.
- [ ] New windows remain unchanged.
- [ ] Missing captured windows produce controlled diagnostics.
- [ ] Reset-without-snapshot remains controlled.
- [ ] In-memory lifetime remains documented.

## Title-bar checklist

- [ ] Exact owned style mask is preserved or deliberately reviewed.
- [ ] Unrelated current style bits are preserved.
- [ ] 32-bit and 64-bit declarations are correct.
- [ ] Valid zero returns are distinguished from failures.
- [ ] Required frame refresh is performed.

## Documentation and demo

- [ ] README
- [ ] INSTALLATION
- [ ] CONTRIBUTING
- [ ] CHANGELOG / release notes
- [ ] Wiki
- [ ] Module headers
- [ ] Demo guidance
- [ ] No documentation change required

- [ ] No binary demo workbook change is included.
- [ ] Or: the binary change is intentional, described, synchronized with exported source, and tested.

## Reviewer notes

Describe trade-offs, known limitations, untested environments, and follow-up work.
