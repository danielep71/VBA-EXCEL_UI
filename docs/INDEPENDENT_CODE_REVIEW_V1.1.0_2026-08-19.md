# Independent Code and Repository Review — VBA Excel UI v1.1.0

> \*\*Repository:\*\* \[`danielep71/VBA-EXCEL\_UI`](https://github.com/danielep71/VBA-EXCEL\_UI)  
> \*\*Tag reviewed:\*\* \[`v1.1.0`](https://github.com/danielep71/VBA-EXCEL\_UI/tree/v1.1.0)  
> \*\*Commit reviewed:\*\* \[`96360379a4bca7703cf649a69a2162961dfa6c9e`](https://github.com/danielep71/VBA-EXCEL\_UI/commit/96360379a4bca7703cf649a69a2162961dfa6c9e)  
> \*\*Branch relationship at review time:\*\* `main` and `v1.1.0` resolve to the same commit  
> \*\*Review date:\*\* 2026-08-19  
> \*\*Suggested repository path:\*\* `docs/INDEPENDENT\_CODE\_REVIEW\_V1.1.0\_2026-08-19.md`

\---

## 1\. Executive assessment

### Overall repository score: **8.0 / 10**

### Production-code quality score: **8.3 / 10**

### Repository-quality score: **7.6 / 10**

### Architecture and modularity score: **9.5 / 10**

### Release-engineering and reproducibility score: **6.1 / 10**

Version `v1.1.0` is a substantial and generally successful evolution of VBA Excel UI. It is no longer a single-module convenience macro: it is a deliberately structured Windows/Excel UI-control component with:

* a stable public façade;
* explicit tri-state visibility semantics;
* backward-compatible window-target scopes;
* best-effort continuation after element-level failures;
* structured failure results;
* snapshot capture and restoration;
* retained Excel `Window` identities for per-window state;
* isolated Win32/Win64 title-bar handling;
* owned-style-bit merging rather than whole-style replacement;
* a dedicated regression harness;
* unusually comprehensive documentation and repository policy for a pure-VBA project.

The strongest parts of the release are:

* the four-module dependency architecture;
* the separation of façade, runtime, snapshot, and WinAPI responsibilities;
* the per-window snapshot identity model for Headings, Workbook Tabs, and Gridlines;
* the use of per-element `Known` flags after partial capture;
* backward-compatible public API evolution;
* the owned-bit title-bar merge;
* careful treatment of ambiguous zero returns from `GetWindowLong` and `SetWindowLong`;
* deterministic ordered diagnostics;
* source-level recovery design;
* an issue-driven release review that found and fixed several genuine P1/P2 defects before tagging;
* strong installation, contribution, security, and line-ending policies.

A material correctness defect nevertheless remains in the title-bar snapshot contract on modern Excel's Single Document Interface (SDI):

> \*\*The snapshot records one title-bar Boolean but does not retain the top-level window handle or Excel `Window` identity from which that Boolean was captured. Restoration resolves `Application.Hwnd` again, and Microsoft documents that this returns the currently active workbook window's handle. Activating another workbook between capture and restore can therefore apply the captured title-bar state to the wrong top-level Excel window.\*\*

The defect is structurally distinct from the otherwise strong per-window snapshot implementation. Headings, Workbook Tabs, and Gridlines are restored through retained `Window` objects; the title bar is not. This creates a split identity model inside one advertised snapshot:

```text
Window properties  -> retained captured Window object -> identity-safe
Title bar           -> current Application.Hwnd        -> active-window-dependent
```

A second SDI concern is that Ribbon state is represented as one application-level Boolean even though modern Excel gives each workbook window its own Ribbon UI. The exact behavior of the fixed XLM `Show.TOOLBAR` command across multiple SDI windows needs empirical characterization and an explicit public contract.

Other important hardening items are:

* make the failure accumulator itself fail-safe, because it is called from error handlers;
* make title-bar style mutation transactional, or remember that a frame refresh is still pending after `SetWindowPos` fails;
* strengthen the regression harness so `RunAll` really runs every release-critical case and cleanup failures cannot be silently suppressed;
* add automated source/static gates and a Windows/Excel execution gate;
* publish machine-readable test evidence tied to the exact release commit and Excel environment;
* update the tagged README, which still describes `v1.1.0` as a release candidate and retains an unchecked pre-release checklist;
* update the demo so it exercises the principal `v1.1.0` features;
* harden and test the repository reformatter.

### Independent verdict

> \*\*VBA Excel UI v1.1.0 is a strong professional VBA component with excellent modularization, API discipline, documentation, and per-window snapshot engineering. It is suitable for controlled Windows Excel use when the exact operating environment is validated. It should not yet claim fully identity-safe title-bar snapshot restoration across multiple modern Excel workbook windows. That SDI identity defect should be the first corrective item in v1.1.1.\*\*

\---

# 2\. Review scope and methodology

## 2.1 Exact source basis

The review was performed against the exact `v1.1.0` revision identified above. The tag and `main` were identical at review time.

### Production source

* [`src/M\_EXCEL\_UI.bas`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/src/M_EXCEL_UI.bas)
* [`src/M\_EXCEL\_UI\_RUNTIME.bas`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/src/M_EXCEL_UI_RUNTIME.bas)
* [`src/M\_EXCEL\_UI\_SNAPSHOT.bas`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/src/M_EXCEL_UI_SNAPSHOT.bas)
* [`src/M\_EXCEL\_UI\_TITLEBAR.bas`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/src/M_EXCEL_UI_TITLEBAR.bas)

### Tests and demo

* [`test/M\_EXCEL\_UI\_REGRESSION\_TESTS.bas`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/test/M_EXCEL_UI_REGRESSION_TESTS.bas)
* [`demo/M\_EXCEL\_UI\_DEMO.bas`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/demo/M_EXCEL_UI_DEMO.bas)
* [`demo/M\_DEMO\_BUILDER.bas`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/demo/M_DEMO_BUILDER.bas)

### Repository tooling and policy

* [`README.md`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/README.md)
* [`INSTALLATION.md`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/INSTALLATION.md)
* [`CHANGELOG.md`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/CHANGELOG.md)
* [`CONTRIBUTING.md`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/CONTRIBUTING.md)
* [`SECURITY.md`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/SECURITY.md)
* [`.gitattributes`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/.gitattributes)
* [`.gitignore`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/.gitignore)
* [`tools/reformat.py`](https://github.com/danielep71/VBA-EXCEL_UI/blob/v1.1.0/tools/reformat.py)
* GitHub issue templates and pull-request template;
* the release pull request, closed release-cycle issues, commit status, and workflow state.

### Platform references

The review also checked the implementation against primary Microsoft documentation for:

* [Excel Single Document Interface behavior](https://learn.microsoft.com/en-us/office/vba/excel/concepts/programming-for-the-single-document-interface-in-excel);
* `Application.Hwnd` and workbook-window behavior under SDI;
* [SetWindowLong / SetWindowLongPtr behavior](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlonga);
* the requirement to call `SetWindowPos(..., SWP\_FRAMECHANGED)` after frame-style changes;
* [VBA `Err.LastDllError`](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/lastdllerror-property).

## 2.2 Review dimensions

The assessment covered:

1. functional correctness;
2. state ownership and lifecycle;
3. multi-window identity semantics;
4. Win32/Win64 API correctness;
5. public API compatibility;
6. error handling and diagnostics;
7. snapshot and emergency recovery;
8. test design and release evidence;
9. documentation accuracy;
10. security and supply-chain boundaries;
11. maintainability and tooling;
12. repository structure, governance, automation, and release quality.

## 2.3 Execution boundary

Desktop Microsoft Excel for Windows was not available in the review environment. The reviewer therefore did **not**:

* import the `.bas` modules into the Visual Basic Editor;
* execute `Debug -> Compile VBAProject`;
* run the four public regression runners;
* reproduce title-bar behavior on live 32-bit and 64-bit Office installations;
* inspect or execute the macro-enabled demo release asset;
* independently verify its checksum or exact source binding.

The review distinguishes among:

* **confirmed source behavior**, established from deterministic control flow;
* **platform-contract findings**, established by combining source behavior with Microsoft documentation;
* **committed manual evidence**, reported in `CHANGELOG.md` and the release pull request;
* **operational state**, such as status checks and workflow runs visible for the exact release commit;
* **runtime hypotheses**, which require an Excel validation case before final closure.

## 2.4 Confidence model

The SDI title-bar identity finding is high confidence because:

1. the snapshot stores only title-bar visibility and no title-bar handle or `Window` identity;
2. title-bar read/write functions resolve `Application.Hwnd` at call time;
3. Microsoft documents that `Application.Hwnd` returns the **active window's** handle under SDI;
4. activating another workbook between capture and restore is ordinary user behavior.

The Ribbon SDI concern is classified as a contract and test-coverage gap rather than a confirmed wrong-result defect because the exact cross-window behavior of the fixed XLM command should be measured in Excel.

\---

# 3\. Hard repository metrics

## 3.1 Production surface

|Metric|v1.1.0|
|-|-:|
|Production modules|**4**|
|Public enums|**2**|
|Documented public callable members|**10**|
|Managed UI elements|**8**|
|Window target scopes|**3**|
|Production-source size|**approximately 200.7 KB**|
|Test-module size|**approximately 212.3 KB**|
|Demo-source size|**approximately 220.3 KB**|
|Logical VBA statements reported by release review|**3,648** across seven VBA modules|
|Regression cases reported by release review|**22**|
|Public test runners|**4**|

## 3.2 Public API inventory

### Enums

```vb
Public Enum UIVisibility
    UI\_LeaveUnchanged = -1
    UI\_Hide = 0
    UI\_Show = 1
End Enum
```

```vb
Public Enum UIWindowTargetScope
    UI\_TargetAllExcelWindows = 0
    UI\_TargetActiveWindow = 1
    UI\_TargetActiveWorkbookWindows = 2
End Enum
```

### Public procedures/functions

```text
UI\_SetExcelUI
UI\_SetExcelUI\_WithResult
UI\_HideExcelUI
UI\_ShowExcelUI
UI\_CaptureExcelUIState
UI\_CaptureExcelUIState\_WithResult
UI\_ResetExcelUIToSnapshot
UI\_ResetExcelUIToSnapshot\_WithResult
UI\_HasExcelUIStateSnapshot
UI\_ClearExcelUIStateSnapshot
```

## 3.3 Managed UI surface

|UI element|Implemented scope|Publicly targetable|
|-|-|:-:|
|Ribbon|documented as application-level|No|
|Status Bar|application-level|No|
|Scroll Bars|application-level|No|
|Formula Bar|application-level|No|
|Headings|Excel `Window`|Yes|
|Workbook Tabs|Excel `Window`|Yes|
|Gridlines|Excel `Window`|Yes|
|Title Bar|top-level window identified through `Application.Hwnd`|No|

## 3.4 Test and release evidence

|Artifact|Current state|
|-|-:|
|Public regression runners|**4**|
|Main-pack cases|**22** reported|
|Manual compile result in changelog|**PASS**|
|Manual runner results in changelog|**4 PASS**|
|Manual recovery checks in changelog|**PASS** reported|
|GitHub Actions workflows|**0**|
|Status checks on release SHA|**0 visible**|
|Workflow runs on release SHA|**0 visible**|
|Formal GitHub review submissions on release PR|**0 visible**|
|Open repository issues at review time|**0**|

The manual results are useful release evidence, but they are not machine-readable, environment-complete, or independently reproducible from the repository.

## 3.5 Repository footprint

The repository is approximately **20 MB**, materially larger than the source itself because it includes several high-resolution media assets. The VBA production code is compact relative to that footprint.

\---

# 4\. Scoring methodology

A score of 10 requires:

* correct behavior throughout the documented public domain;
* identity-safe operation under modern Excel's multi-window model;
* no known path that applies captured state to the wrong host object;
* explicit and enforceable error contracts;
* deterministic regression coverage of critical paths;
* exact-environment test evidence tied to the release commit;
* automated source, documentation, and release gates;
* authoritative documentation with no release-state drift;
* reproducible binary release provenance;
* maintainable module and tooling boundaries.

## Weighted scorecard

|Area|Weight|Score|Weighted contribution|
|-|-:|-:|-:|
|Functional correctness|16%|**7.9**|1.264|
|State, identity, and recovery correctness|14%|**7.6**|1.064|
|Architecture and modularity|10%|**9.5**|0.950|
|Public API design and compatibility|8%|**9.1**|0.728|
|Error handling and diagnostics|9%|**8.0**|0.720|
|WinAPI and platform engineering|8%|**7.8**|0.624|
|Regression testing|10%|**8.2**|0.820|
|CI and release reproducibility|8%|**5.8**|0.464|
|Documentation and governance|7%|**8.0**|0.560|
|Maintainability and tooling|6%|**8.0**|0.480|
|Security and supply-chain hygiene|4%|**8.3**|0.332|
|**Total**|**100%**||**8.006 / 10**|

Rounded overall score:

```text
8.0 / 10
```

## Score interpretation

|Score|Interpretation|
|-:|-|
|9.5–10.0|Exceptional, independently certified, and operationally enforced|
|9.0–9.4|Advanced professional component with limited remaining gaps|
|8.0–8.9|Strong implementation requiring material targeted hardening|
|7.0–7.9|Good foundation with significant correctness or assurance gaps|
|Below 7.0|Major design, correctness, or repository-control deficiencies|

\---

# 5\. Component scores

|Component|Score|Assessment|
|-|-:|-|
|`M\_EXCEL\_UI`|**9.1**|Strong façade, backward-compatible targeting, validation, orchestration, and per-window continuation|
|`M\_EXCEL\_UI\_RUNTIME`|**8.2**|Cohesive shared services; failure accumulation itself needs a fail-safe boundary|
|`M\_EXCEL\_UI\_SNAPSHOT`|**8.4**|Excellent retained-object model for window properties; title-bar/Ribbon state remains singleton and not identity-bound|
|`M\_EXCEL\_UI\_TITLEBAR`|**7.7**|Good owned-bit and bitness design; SDI handle semantics, singleton baseline, and refresh-failure transaction gap are material|
|Regression harness|**8.1**|Broad and thoughtful; `RunAll` is incomplete, important cases can be skipped, cleanup failures can be hidden|
|Demo layer|**7.2**|Polished source and builder, but does not demonstrate the principal v1.1.0 targeting and structured-result features|
|`tools/reformat.py`|**6.8**|Useful deterministic intent; regex-based rewriting is not generally token-safe and lacks tests/check mode|
|Documentation|**8.0**|Extensive and professional; tagged README is stale and SDI scope claims need correction|
|Security policy|**8.6**|Clear trust boundaries, private reporting, safe-use guidance, and explicit non-security role of UI hiding|
|Repository quality|**7.6**|Excellent structure and editorial governance; weak automation, evidence, release-state synchronization, and independent review controls|

\---

# 6\. Architectural review

## 6.1 Dependency architecture

The production dependency graph is clear and acyclic:

```text
M\_EXCEL\_UI
├── M\_EXCEL\_UI\_RUNTIME
├── M\_EXCEL\_UI\_TITLEBAR
└── M\_EXCEL\_UI\_SNAPSHOT
    ├── M\_EXCEL\_UI\_RUNTIME
    └── M\_EXCEL\_UI\_TITLEBAR
```

This is a major improvement over the earlier single-module architecture.

### `M\_EXCEL\_UI`

Owns:

* the documented public API;
* the two public enums;
* visibility validation;
* target-scope resolution;
* selective application orchestration;
* compatibility wrappers.

### `M\_EXCEL\_UI\_RUNTIME`

Owns:

* ordered failure collection;
* structured output buffers;
* Immediate Window logging;
* Ribbon read/write helpers;
* generic Boolean property read/write helpers;
* quiet `ScreenUpdating` scopes;
* shared diagnostic labels and error text.

### `M\_EXCEL\_UI\_SNAPSHOT`

Owns:

* all mutable snapshot state;
* retained Excel `Window` references;
* per-element `Known` flags;
* capture and restore orchestration;
* dead-window detection;
* snapshot lifecycle and release.

### `M\_EXCEL\_UI\_TITLEBAR`

Owns:

* every WinAPI declaration;
* the exact title-bar style mask;
* handle-scoped title-bar baseline state;
* owned-bit merging;
* frame refresh;
* the pure arithmetic merge test seam.

## 6.2 Visibility boundaries

All four production modules use:

```vb
Option Explicit
Option Private Module
```

This is appropriate for a component intended to expose a project-facing API while hiding implementation seams from cross-project automation and the Macro dialog.

The two title-bar test seams are `Public` only because VBA cannot otherwise reach them from another standard module in the same project. `Option Private Module` prevents normal external project exposure. That is a reasonable testability trade-off, although a static public-surface checker should explicitly distinguish supported façade members from internal same-project seams.

## 6.3 Cohesion

The decomposition is based on state ownership, not arbitrary file size. That is the correct criterion:

* snapshot state is centralized;
* title-bar state is centralized;
* the façade does not duplicate either;
* the runtime module is largely stateless;
* the title-bar module deliberately has no project dependency;
* no circular dependency was identified.

## 6.4 Architectural limitation

The principal architecture gap is not module decomposition but **scope modeling**:

```text
Application-wide state       -> one value
Window object-model state    -> one value per retained Window object
Top-level SDI frame state     -> one value resolved through current Application.Hwnd
Ribbon state under SDI        -> one value, host behavior not fully characterized
```

Modern Excel does not have one permanent top-level frame per process. Each workbook window has its own top-level SDI window and Ribbon. Therefore, the architecture needs a first-class model for:

* process-global state;
* workbook-window state;
* active-window-only state;
* all-SDI-window state;
* state keyed by a specific top-level `HWND`.

## 6.5 Architectural verdict

**9.5 / 10** for decomposition and dependency discipline.

No wholesale rewrite is needed. The correct next architectural step is to extend the existing ownership model to explicit SDI frame and Ribbon scopes rather than undo the four-module design.

\---

# 7\. Production-code review

## 7.1 `M\_EXCEL\_UI` — **9.1 / 10**

### Strengths

#### Stable façade

The public `UI\_...` surface remains concentrated in one module. Callers do not need to understand snapshot internals, WinAPI details, or the structured-result implementation.

#### Backward-compatible extension

`TargetScope` is optional, trailing, and defaults to:

```vb
UI\_TargetAllExcelWindows
```

Existing call sites retain their prior parameter positions and behavior for Headings, Workbook Tabs, and Gridlines.

#### Complete input validation

The general apply path validates all eight `UIVisibility` inputs and the target-scope enum. Invalid values are accumulated as structured failures rather than allowing the first one to abort validation of later values.

This is better than a conventional fail-fast VBA implementation because it gives a caller a complete diagnosis in one pass.

#### Correct separation of scopes

The façade correctly limits `TargetScope` to:

```text
Headings
Workbook Tabs
Gridlines
```

Ribbon, Status Bar, Scroll Bars, Formula Bar, and Title Bar retain their own established mechanisms.

#### Per-window continuation

`UI\_ApplyWindowLevelState` now has a local error boundary. An unusable window no longer escapes to the outer apply handler and silently prevents later windows from being attempted.

The diagnostic label is built before the write sequence so a failing `Window.Caption` read cannot turn a property failure into a pass-level abort.

#### No-op semantics

`UI\_LeaveUnchanged` is a real no-op, not an alias for the current Boolean value. The design also uses IfNeeded helpers to avoid unnecessary writes when the current state is readable.

### Minor concerns

#### Quiet-scope success is not verified

The runtime helper attempts:

```vb
Application.ScreenUpdating = False
```

under a suppressed-error scope and records that quiet mode changed. It does not verify that the assignment succeeded or read back the resulting value.

This is low severity because `ScreenUpdating` is a performance/visual-stability aid rather than a correctness precondition, but the state flag should mean “the write succeeded,” not merely “the write was attempted.”

#### Application state can change while targets are enumerated

The target resolver and later write loop operate against live Excel collections. Activation, workbook closure, or add-in activity can change the target set during the operation. The per-window local boundary limits damage, but a deterministic snapshot of target references before mutation would make failure ordering and target scope more stable.

### Recommended improvements

1. Resolve window targets into retained object references before any write.
2. Record whether `ScreenUpdating=False` actually succeeded.
3. Add a machine-readable static API inventory for the façade.
4. Add a public-result case where multiple visibility values and the target scope are simultaneously invalid, asserting complete deterministic ordering.

\---

## 7.2 `M\_EXCEL\_UI\_RUNTIME` — **8.2 / 10**

### Strengths

#### Cohesive stateless services

The module is appropriately limited to operations shared by the façade and snapshot engine.

#### Structured-result contract

The result API is clear:

|Output|Meaning|
|-|-|
|Function result `True`|no failure recorded|
|Function result `False`|one or more failures recorded|
|`FailureCount`|exact count|
|`FailureList`|optional one-based ordered array|

Output buffers are cleared on entry, avoiding stale data from a prior call.

#### Stable diagnostic format

Failure entries use:

```text
Stage | Detail
```

Fire-and-forget paths use a consistent Immediate Window line. The shared window-label builder avoids risky reads while composing an error message.

#### Error-text correction

The runtime error-text builder captures:

```text
Err.Number
Err.Description
Err.Source
Erl
```

before using `On Error Resume Next`. This correctly avoids the prior defect where the protective error statement reset the very error state being formatted.

#### Fixed and bounded XLM input

The Ribbon setter constructs only one of two fixed commands:

```text
Show.TOOLBAR("Ribbon",True)
Show.TOOLBAR("Ribbon",False)
```

No user-controlled macro text is interpolated, which materially limits injection risk.

### Finding: failure recording can itself raise

The central failure path calls `UI\_RuntimeAddFailure`, which:

1. marks the pass unsuccessful;
2. increments `FailureCount`;
3. performs `ReDim Preserve` on a string array;
4. writes the new failure entry.

The routine does not have its own protective error boundary.

This matters because it is invoked from error handlers. An allocation error, a Variant type conflict, or another unexpected failure in the diagnostic path can:

* mask the original Excel/WinAPI error;
* escape a function documented as fail-soft;
* prevent quiet-scope cleanup;
* leave `FailureCount` inconsistent with `FailureList`;
* convert a recoverable element failure into an unhandled VBA exception.

The release pull request itself identifies this as a `v1.1.1` follow-up, which is an accurate assessment.

### Finding: success often means “call returned,” not “state was achieved”

The generic setter and Ribbon setter report success after the host call completes. They do not read back the state.

This is acceptable as a best-effort default, but it should be explicit:

```text
operation accepted by host != visual state independently verified
```

This distinction is especially important for:

* Ribbon behavior under policy restrictions;
* title-bar changes normalized by Windows/Excel;
* transient window states;
* add-ins that immediately rewrite the same property.

### XLM policy dependency

The component relies on Excel 4 macro execution for Ribbon control. The command is fixed and does not create an arbitrary-code input surface, but enterprise Office policy can restrict legacy macro capabilities.

The installation and troubleshooting documentation should explicitly identify XLM availability as a host-policy dependency alongside WinAPI permission.

### Recommended improvements

1. Make `UI\_RuntimeAddFailure` non-raising under every ordinary condition.
2. Keep a small emergency fallback string when array growth fails.
3. Preserve the original error detail before entering the diagnostic path.
4. Add optional readback verification for release validation.
5. Add a fault-injection seam for result-buffer growth and logging failure.
6. Document XLM policy requirements and expected failure stages.

\---

## 7.3 `M\_EXCEL\_UI\_SNAPSHOT` — **8.4 / 10**

### Strengths

#### Correct retained-object identity model

For Headings, Workbook Tabs, and Gridlines, the snapshot retains the exact captured Excel `Window` object and later probes that retained object before restoring it.

This avoids the three major hazards of index-based restoration:

```text
reordering
window replacement
new windows occupying old collection positions
```

The consequences are correct:

* reordered surviving windows restore correctly;
* a newly opened window is unchanged;
* a closed captured window is reported and skipped;
* captured state is not redirected to a replacement window.

#### Per-element `Known` flags

Every managed state carries a `Known` flag after the v1.1.0 fixes. A read that fails does not silently become a captured `False` value.

Restoration writes only values that were successfully captured.

This is a strong partial-snapshot contract:

```text
partial capture remains usable
unknown values are not invented
known values remain recoverable
```

#### Failure isolation

Capture continues after element-level read failures. Restoration continues after property-level and window-level failures. Only a genuinely unexpected pass-level capture error clears the partially built snapshot.

#### Explicit lifecycle

Snapshot replacement, retention, replay, and clearing are documented. The implementation deliberately retains the snapshot after restoration so it can be replayed.

The installation guide correctly explains that retained `Window` references must eventually be released through:

```vb
UI\_ClearExcelUIStateSnapshot
```

### Material limitation: not every part of the snapshot has an identity

The module-level state for the title bar is only:

```vb
Private m\_SnapshotTitleBarKnown   As Boolean
Private m\_SnapshotTitleBarVisible As Boolean
```

There is no corresponding:

```text
captured Window object
captured HWND
captured workbook identity
captured top-level frame label
```

The title-bar worker therefore cannot restore the same frame that was read unless that frame also happens to be active at restoration time.

Ribbon state is likewise represented by one Boolean and one `Known` flag, despite modern Excel's one-Ribbon-per-workbook-window SDI model.

### Collection-mutation edge

Capture sizes parallel arrays from:

```vb
Application.Windows.Count
```

and then enumerates the live collection. If another macro, event, or add-in changes the collection during the pass, the count and enumeration can diverge. The unexpected error handler safely clears the partial snapshot, but a two-phase capture would be more deterministic:

1. retain the target `Window` objects;
2. size arrays from the retained count;
3. capture properties from that stable object list.

### Resource lifetime

Retained COM object references are necessary for identity safety, but they can prolong object/workbook lifetime. The documentation now covers this well. The remaining improvement is an optional diagnostic helper that reports snapshot age and retained-window count, so long-running host solutions can detect a forgotten snapshot.

### Recommended improvements

1. Add explicit title-bar identity state.
2. Characterize and model Ribbon state under SDI.
3. Resolve the complete window object list before allocating property arrays.
4. Add snapshot metadata: capture timestamp, captured active window label, window count, completeness count.
5. Add an optional `UI\_GetExcelUIStateSnapshotInfo` result for diagnostics without exposing mutable internals.

\---

## 7.4 `M\_EXCEL\_UI\_TITLEBAR` — **7.7 / 10**

### Strengths

#### WinAPI isolation

All operating-system-sensitive code is confined to one module. The façade and snapshot engine do not repeat platform declarations.

#### 32-bit and 64-bit compatibility

The module handles:

```text
VBA7 + Win64
VBA7 + Win32
pre-VBA7 32-bit
```

with the expected `GetWindowLongPtr`/`SetWindowLongPtr` route on 64-bit Office and the correct handle type on VBA7 32-bit Office.

#### Exact style ownership

The module claims only:

```text
WS\_CAPTION
WS\_SYSMENU
WS\_THICKFRAME
WS\_MINIMIZEBOX
WS\_MAXIMIZEBOX
```

through:

```text
TITLEBAR\_OWNED\_STYLE\_MASK = \&HCF0000
```

This is materially safer than restoring a complete historical `GWL\_STYLE` value.

#### Pure merge function

The merge rule is isolated as pure arithmetic:

```vb
(CurrentStyle And Not TITLEBAR\_OWNED\_STYLE\_MASK) Or \_
(OwnedStyleBits And TITLEBAR\_OWNED\_STYLE\_MASK)
```

That makes the central ownership policy deterministic and testable without touching a live Excel window.

#### Ambiguous-zero handling

Before `GetWindowLong` and `SetWindowLong`, the module clears last error and then checks both the API return and the last-error value. This follows the Win32 requirement for distinguishing a valid zero return from failure.

#### Emergency show fallback

If the first operation after a VBA project reset is “show” while the frame is already hidden, the module restores the full owned frame rather than capturing zero and reporting a false no-op success.

This fixes a real recovery blocker found during the v1.1.0 release review.

### Material limitations

#### SDI identity

The module stores one current-handle baseline:

```text
m\_OriginalMainWindowOwnedStyleBits
m\_OriginalMainWindowHwnd
m\_HasOriginalMainWindowOwnedStyleBits
```

Whenever `Application.Hwnd` changes, the singleton baseline is replaced.

On Excel 2013 and later, `Application.Hwnd` is the active workbook window's handle, not a permanent process-wide frame. Therefore:

* the component does not own one stable “main window”;
* switching active workbooks changes the target;
* title-bar snapshot restoration is active-window-dependent;
* one handle's baseline can be overwritten by another handle's baseline;
* a show operation may restore the owned bits most recently captured for the current handle only;
* the public description “Excel main window” is too imprecise for SDI.

#### Non-transactional style change

The write sequence is:

```text
SetWindowLong / SetWindowLongPtr
SetWindowPos(... SWP\_FRAMECHANGED)
```

If the style write succeeds but the frame refresh fails:

1. the function reports failure;
2. the style has already changed;
3. there is no rollback;
4. no pending-refresh state is recorded;
5. a later retry can see `NewStyle = CurrentStyle` and take the no-op path;
6. that no-op path does not call `SetWindowPos` again.

Microsoft explicitly states that frame-style changes do not take effect correctly until `SetWindowPos` is called with `SWP\_FRAMECHANGED`. The retry path must therefore not treat “style bits already match” as complete success when the prior frame refresh failed.

#### Baseline staleness

Owned bits are captured once per observed handle. If Excel or another add-in later changes one of the owned frame bits while the title bar is visible, a future show can restore the stale first-captured bit set rather than the most recent legitimate visible frame state.

The release pull request already identifies this as a follow-up.

#### `GetLastError` access in VBA

The current declarations call `GetLastError` directly. VBA also exposes `Err.LastDllError`, which Microsoft documents as the immediate post-DLL-call error channel for VBA. The direct API route is not automatically wrong, but `Err.LastDllError` is the language-native mechanism and reduces the risk that an intervening call changes the thread error state.

### Recommended improvements

1. Accept an explicit target `HWND` or `Window` object in title-bar internal APIs.
2. Store baselines in a keyed collection by `HWND` rather than one singleton.
3. Capture and restore the exact title-bar target in the snapshot.
4. Track `FrameRefreshPending` after a successful style write followed by failed refresh.
5. Retry `SetWindowPos` even when style bits already match if refresh remains pending.
6. Consider rollback to the prior style when frame refresh fails.
7. Use `Err.LastDllError` immediately after DLL calls, or document and test the direct `GetLastError` design.
8. Add readback and visual-state verification in the release harness.
9. Add `SWP\_NOACTIVATE` if future code mutates non-active SDI windows.

\---

# 8\. Dedicated SDI scope and identity review

This section is central to the assessment because Excel's window architecture changed materially in Excel 2013.

## 8.1 Relevant Excel platform model

Microsoft documents that modern Excel uses a Single Document Interface:

* each workbook window is a separate top-level operating-system window;
* each workbook window has its own Ribbon UI;
* multiple workbook windows can still belong to one Excel process;
* `Application.Windows` remains available;
* `Application.Hwnd` returns the **active window's** top-level handle;
* developers may need to cache or propagate UI state while users switch workbook windows.

Therefore, the following terms are not interchangeable:

```text
Excel process
active workbook window
all workbook windows in the process
one specific top-level HWND
application object-model property
Ribbon instance
```

## 8.2 Current title-bar capture/restore sequence

The current snapshot sequence is effectively:

```text
CAPTURE
1. Read Application.Hwnd for whichever workbook window is active.
2. Read WS\_CAPTION from that handle.
3. Store only Visible/Not Visible.

RESTORE
1. Read Application.Hwnd again.
2. Apply the stored Boolean to whichever workbook window is active now.
```

No identity check connects the two handles.

## 8.3 Deterministic counterexample

Assume two workbook windows in one Excel instance:

```text
Window A -> HWND\_A -> title bar visible
Window B -> HWND\_B -> title bar hidden
```

Sequence:

```vb
'Window A active
UI\_CaptureExcelUIState

'User activates Window B
UI\_ResetExcelUIToSnapshot
```

Current title-bar result:

```text
Captured value from: HWND\_A
Applied value to:    HWND\_B
HWND\_A restored:     No
HWND\_B modified:     Yes
Failure reported:    No, provided both API operations succeed
```

This is not an ambiguous collection-reordering case. It follows directly from resolving the active handle at two different times.

## 8.4 Why retained `Window` identity does not solve it

The snapshot engine correctly retains every Excel `Window` object for Headings, Workbook Tabs, and Gridlines. However, the title-bar state is captured before the per-window arrays and is stored outside them as a singleton.

The title-bar worker is therefore unaware of:

* which retained `Window` was active at capture;
* which `HWND` carried the captured title bar;
* whether that window is still open;
* whether the active window changed;
* whether the user intended active-window or all-window title-bar semantics.

## 8.5 Public-contract implications

### `UI\_CaptureExcelUIState` / `UI\_ResetExcelUIToSnapshot`

The phrase “identity-safe snapshot restore” is accurate for the three object-model window properties, but not for the title bar. The documentation should qualify the guarantee until fixed.

### `UI\_HideExcelUI` / `UI\_ShowExcelUI`

The wrappers apply Headings, Workbook Tabs, and Gridlines to all Excel windows by default, but the title-bar function uses one active `Application.Hwnd`.

Thus “hide/show the complete managed shell” means:

```text
all current Excel Window objects for three properties
one active top-level SDI frame for the title bar
application/object-model behavior for the remaining properties
```

That mixed scope is not obvious from the public method name.

### `TargetScope`

The documentation says `TargetScope` does not apply to Title Bar. That is internally consistent, but the alternative scope is not clearly defined. It is neither truly process-wide nor identity-bound; it is active-frame-at-execution-time.

## 8.6 Ribbon implications

Microsoft documents one Ribbon UI per workbook window under SDI. The current library stores:

```text
m\_SnapshotRibbonKnown
m\_SnapshotRibbonVisible
```

and uses `Application.CommandBars("Ribbon")` / fixed XLM commands without activating and validating each workbook window.

Possible actual host behaviors include:

* the command affects only the active Ribbon;
* Excel propagates the state to all Ribbon instances;
* behavior varies by Excel build or window state;
* reads and writes are asymmetric;
* a newly opened workbook receives a default state rather than the current cached state.

The source alone cannot select among these possibilities. This is why the Ribbon item is classified as a P2 contract/assurance gap rather than the same confirmed P1 identity defect as the title bar.

## 8.7 Correct target models

The project should choose and document one of these title-bar models.

### Model A — Active window only

```text
UI\_SetExcelUI TitleBar:=... operates on the current active workbook window.
Snapshot captures the active Window/HWND and restores only that exact frame.
```

This is the smallest change and most closely matches current implementation intent.

### Model B — Follow `TargetScope`

```text
Title bar becomes a targetable window-level element.
All, active, or active-workbook windows can be selected.
```

This is more expressive but changes the conceptual public model. It can remain backward-compatible if the default preserves the established active-window behavior.

### Model C — All SDI windows for show/hide wrappers

```text
Selective title-bar calls remain active-window-only.
UI\_HideExcelUI/UI\_ShowExcelUI enumerate all current top-level workbook windows.
Snapshot stores one frame state per captured Window/HWND.
```

This best matches the “complete managed shell” wording but requires more implementation work.

## 8.8 Required regression matrix

At minimum, add live Excel cases for:

|Case|Required assertion|
|-|-|
|Capture A, activate B, restore|A restored; B unchanged|
|Capture A and B with different frame states|each frame receives its own captured state|
|Close captured A, create replacement|replacement unchanged; A reported missing|
|Switch active window during title-bar show/hide|documented target receives the change|
|Style write succeeds, frame refresh fails|retry refreshes rather than no-op|
|Two workbook windows with different Ribbon visibility|read/write/snapshot behavior matches documented scope|
|Open a new workbook after Ribbon hide|behavior is characterized and documented|

## 8.9 SDI verdict

> \*\*The per-window object-model snapshot design is genuinely strong. The title-bar and Ribbon subsystems have not yet been brought to the same SDI-aware identity standard.\*\*

\---

# 9\. Public API and compatibility review

## 9.1 Naming and discoverability

The public naming convention is consistent:

```text
UI\_<Verb><Object>
UI\_<Verb><Object>\_WithResult
```

The two enums provide readable named arguments and avoid Boolean ambiguity.

## 9.2 Tri-state design

`UIVisibility` is an effective public contract:

|Value|Meaning|
|-|-|
|`UI\_LeaveUnchanged`|do not read or write for the purpose of applying a new state|
|`UI\_Hide`|request hidden state|
|`UI\_Show`|request visible state|

This is materially better than many UI helper APIs that require callers to construct separate include/exclude lists or pass nullable Variants.

## 9.3 Structured and fire-and-forget pairs

The package offers both:

```text
compatibility Sub -> logs failures
structured Function -> returns Boolean/count/list
```

This is a good migration pattern for VBA. Existing callers remain simple, while governed callers can inspect results.

## 9.4 Semantic-versioning assessment

The source API change from v1.0.1 is backward-compatible:

* no existing public procedure was removed;
* no existing enum member changed;
* the new scope parameter is optional and trailing;
* existing positional calls remain valid;
* show/hide semantics remain intact.

However, the release statement “no migration required” needs precision.

### Caller-code migration

```text
None required
```

### Deployment/install migration

```text
Required for existing one-module installations
```

A v1.0.1 workbook must replace the prior single module with all four v1.1.0 production modules. `INSTALLATION.md` documents this clearly, but the headline compatibility statement should distinguish source compatibility from package-layout migration.

## 9.5 Scope transparency

The public API is clear for the three targetable window properties. It is less clear for the title bar and Ribbon under SDI. The API reference should explicitly state:

```text
TitleBar target: active top-level workbook window at execution time (current v1.1.0 behavior)
Ribbon target: host-dependent; exact SDI behavior characterized in supported environments
```

until a stronger model is implemented.

## 9.6 Public API verdict

**9.1 / 10**

The interface is compact, readable, and compatibility-conscious. The remaining issue is not naming but the precision of cross-window scope guarantees.

\---

# 10\. Error handling, diagnostics, and transactional behavior

## 10.1 Positive design

The component generally follows a strong fail-soft model:

1. perform one operation;
2. record failure rather than raise;
3. continue with unrelated operations;
4. restore `ScreenUpdating` if changed;
5. expose ordered structured results when requested.

No unsolicited `MsgBox` is used in production code.

`On Error Resume Next` scopes are usually narrow and followed by an explicit restore of the intended handler.

## 10.2 Diagnostic ordering

The ordering is deterministic within a pass. This is useful for:

* regression tests;
* issue reproduction;
* human interpretation;
* machine comparison of expected failures.

## 10.3 Diagnostic atomicity gap

The failure list is mutable state built while handling another failure. It should satisfy a stronger invariant:

```text
The diagnostic path may degrade detail, but it must never raise or erase the original failure.
```

Current `ReDim Preserve` growth does not guarantee that invariant.

A robust pattern is:

```vb
Private Sub UI\_RuntimeTryRecordFailure(...)
    On Error GoTo MinimalFallback
    'normal array accumulation
    Exit Sub

MinimalFallback:
    'preserve Succeeded=False and FailureCount at a coherent value
    'optionally retain one static fallback message
End Sub
```

## 10.4 Partial mutation

Most Boolean property writes are naturally atomic from the component's perspective: either the property assignment raises or it does not.

The title-bar path is a multi-step transaction:

```text
read current style
compute merge
write style
refresh frame
```

It needs an explicit transaction state because the style can be committed before the visible refresh completes.

## 10.5 Logging behavior

Immediate Window logging is appropriate for development but is not durable operational telemetry. For integration into governed workbooks, consider an optional callback or returned result only; do not introduce workbook writes or event-based logging by default.

## 10.6 Recommended diagnostic contract

Document three levels:

|Result class|Meaning|
|-|-|
|Verified success|host state read back and matched|
|Best-effort success|host call completed; visible state not independently confirmed|
|Failure|call failed or readback did not match|

The current Boolean contract can remain unchanged, while detailed status text or a future result type can distinguish the first two.

\---

# 11\. Regression-test review

## 11.1 Test inventory

The harness exposes:

```vb
Test\_EXCEL\_UI\_RunCore
Test\_EXCEL\_UI\_RunTitleBarOnly
Test\_EXCEL\_UI\_RunSnapshotIdentity
Test\_EXCEL\_UI\_RunAll
```

The release review reports 22 cases, covering:

* show-all baseline;
* selective hide;
* selective show;
* leave-unchanged/no-op;
* convenience wrappers;
* structured-result success;
* structured-result output clearing;
* invalid visibility values;
* invalid target scope;
* active-window targeting;
* active-workbook-window targeting;
* `ScreenUpdating` preservation;
* snapshot capture and restore;
* no-snapshot restoration;
* partial application-level capture independence;
* closed-window identity failure;
* replacement-window non-interference;
* title-bar round-trip;
* title-bar owned-bit preservation;
* post-reset show recovery.

## 11.2 Strong qualities

### Historical defect regression

The harness contains cases specifically designed around real defects found during the release cycle. This is much stronger than testing only happy paths.

### Live WinAPI validation

The title-bar pack includes both:

* a pure arithmetic merge test;
* a live `Application.Hwnd` round-trip.

The combination is appropriate because arithmetic correctness alone does not prove frame rendering.

### Identity scenario

The dedicated snapshot runner creates a captured temporary window, closes it, creates a replacement, and verifies that the replacement does not inherit captured state. This directly guards the principal per-window identity promise.

### Host-state preservation intent

The harness snapshots the existing host state and attempts to restore it after pass or failure. That is essential for tests that deliberately hide Excel UI.

## 11.3 Material limitations

### `RunAll` does not run all release-critical tests

`Test\_EXCEL\_UI\_RunAll` invokes the main regression pack, but the dedicated `Test\_EXCEL\_UI\_RunSnapshotIdentity` runner is separate. The documented release sequence runs both, but the name `RunAll` overstates its coverage.

Recommended correction:

```text
Test\_EXCEL\_UI\_RunAll -> calls core, title-bar, and identity packs
```

or rename the existing procedure to:

```text
Test\_EXCEL\_UI\_RunGeneral
```

while preserving the old public name as a compatibility wrapper that now runs everything.

### Important cases can be skipped

When a pre-existing EXCEL\_UI snapshot exists, snapshot-destructive cases are skipped because the harness cannot reconstruct the caller's prior snapshot object.

A green `RunAll` can therefore mean:

```text
all requested cases passed, but several snapshot lifecycle cases did not execute
```

This is acceptable for an interactive developer run but not sufficient for release certification. A release runner should fail or clearly return `INCOMPLETE` when mandatory cases are skipped.

### Test cleanup uses index-based state

The harness's own pre-test window-state capture/restoration uses `Application.Windows` index order. The production code correctly rejects that identity strategy.

If windows are opened, closed, or reordered during a failing test, cleanup can restore state to the wrong window or fail to restore the original environment exactly.

The harness should retain exact `Window` objects just as the production snapshot does.

### Cleanup failure is suppressed

The main runner calls `TST\_RestoreState` under `On Error Resume Next`. A cleanup error does not necessarily fail the test run, and the harness logs PASS before cleanup.

For UI-manipulating tests, cleanup is part of test correctness. A run that passes assertions but leaves Excel in an altered or constrained state is not a clean pass.

Recommended result classes:

```text
PASS
FAIL\_ASSERTION
FAIL\_CLEANUP
INCOMPLETE\_SKIPPED
```

### No machine-readable counters

There is no durable result artifact containing:

* case count;
* assertion count;
* passed/failed/skipped counts;
* exact Excel version/build/bitness;
* Windows version;
* release commit SHA;
* failure details.

### Fault-injection limitations

The source candidly notes that some host-refusal states cannot be forced deterministically. The right response is to add test seams at the runtime boundary, not to rely only on naturally occurring Excel failures.

## 11.4 Missing critical cases

Add permanent cases for:

```text
SDI\_TitleBar\_CaptureA\_ActivateB\_Restore
SDI\_TitleBar\_TwoWindowDistinctStates
SDI\_TitleBar\_ClosedCapturedFrame
SDI\_Ribbon\_TwoWorkbookWindows
TitleBar\_FrameRefreshFailureRetry
FailureAccumulator\_FaultInjection
QuietScope\_WriteRefused
RunAll\_ReportsSkippedAsIncomplete
Harness\_CleanupFailureIsFailure
Snapshot\_WindowCollectionMutation
```

## 11.5 Test verdict

**8.1 / 10**

The harness is thoughtful and materially better than typical VBA tests. It needs stronger release-runner semantics, exact object identity for its own cleanup, deterministic fault injection, and automated result capture.

\---

# 12\. Demo and repository tooling review

## 12.1 Demo layer

The demo source is visually and operationally ambitious. It provides:

* checkbox-driven show/hide selection;
* presets;
* current-state synchronization;
* capture/reset actions;
* a repeatable worksheet builder;
* explanatory note areas;
* both Forms and ActiveX checkbox compatibility.

However, the release pull request correctly states that the demo does not yet exercise the headline v1.1.0 features.

### Missing v1.1.0 demonstrations

The demo should expose:

* a target-scope selector;
* active-window targeting;
* active-workbook-window targeting;
* `UI\_SetExcelUI\_WithResult` output;
* structured snapshot-capture results;
* structured snapshot-restore results;
* failure-list rendering;
* snapshot availability and retained-window count;
* an explicit warning about active SDI title-bar scope.

### Demo-builder size

`M\_DEMO\_BUILDER.bas` is larger than any production module. It is generic and potentially reusable, but it materially increases review surface for a small UI component.

Options:

1. keep it, but document it as a separately reusable demo infrastructure module;
2. move it to a shared demo-framework repository and pin a version;
3. reduce the local copy to the helpers actually required by this demo.

## 12.2 Binary demo distribution

The repository intentionally excludes `.xlsm` demo files and instructs users to obtain tested binaries from GitHub Releases. This is sound source-control hygiene.

For stronger supply-chain assurance, each release should publish:

```text
EXCEL\_UI\_DEMO\_v1.1.0.xlsm
EXCEL\_UI\_DEMO\_v1.1.0.sha256
RELEASE\_MANIFEST\_v1.1.0.json
```

The manifest should identify:

* release tag and full commit SHA;
* source module hashes;
* demo module hashes;
* workbook hash;
* Excel version/build/bitness used to compile and test;
* Windows version;
* test result artifact hash;
* build timestamp.

The existence and exact source binding of the binary release asset could not be independently confirmed through the review interface.

## 12.3 `tools/reformat.py`

### Positive design

The formatter aims to make only mechanical changes:

* hoist options;
* normalize section banners;
* rename standard labels;
* align declarations;
* normalize rules and CRLF;
* preserve executable behavior.

The release pull request reports a statement-for-statement equivalence check across 3,648 logical statements after the v1.1.0 reformat. That is useful evidence for the specific committed transformation.

### General safety limitation

The formatter uses regular expressions over whole source lines. For example, label-reference replacement searches for patterns such as:

```text
GoTo SafeExit
Resume CleanExit
```

without a VBA tokenizer that separates:

```text
executable code
string literals
comments
conditional-compilation text
```

A line containing the same phrase inside a string or comment can be rewritten even though the tool claims never to touch executable tokens.

### Encoding limitation

The tool reads and writes Latin-1. A future module containing characters outside that encoding can fail or be corrupted.

### Operational gaps

The tool lacks:

* argument validation;
* a `--check` mode;
* idempotence verification;
* golden-file tests;
* parsing tests for strings/comments;
* a dry run;
* a diff summary;
* CI enforcement.

### Recommended tooling contract

```text
reformat.py --check <paths>        exits nonzero on drift
reformat.py --write <paths>        rewrites files
reformat.py --diff <paths>         prints proposed changes
pytest                             validates golden cases and idempotence
```

The tool should either tokenize VBA lines or explicitly skip transformations once a string/comment boundary is reached.

\---

# 13\. Dedicated repository-quality assessment

## Repository-quality score: **7.6 / 10**

This score is deliberately separate from production-code quality.

The repository has unusually high **editorial and structural maturity** for a VBA project, but substantially lower **operational and enforceable maturity**. In practical terms:

> \*\*The repository explains good engineering discipline better than it automatically proves and enforces it.\*\*

## 13.1 Repository-quality scorecard

|Repository dimension|Weight|Score|Weighted contribution|
|-|-:|-:|-:|
|Structure and discoverability|15%|**9.2**|1.380|
|Documentation quality|15%|**8.1**|1.215|
|Contribution and governance policy|12%|**8.7**|1.044|
|Source and binary hygiene|10%|**9.0**|0.900|
|Testing assets|12%|**8.1**|0.972|
|Automation and status enforcement|15%|**5.3**|0.795|
|Release reproducibility and provenance|10%|**6.5**|0.650|
|Review/community controls|6%|**6.8**|0.408|
|Maintenance backlog discipline|5%|**5.4**|0.270|
|**Total**|**100%**||**7.634 / 10**|

Rounded:

```text
7.6 / 10
```

## 13.2 Strong repository qualities

### Clear source layout

The repository separates:

```text
src/
demo/
test/
tools/
.github/
```

The four required production modules are visible and documented. The binary demo is intentionally excluded from source control.

### Complete governance documents

The repository contains:

* MIT license;
* Code of Conduct;
* contribution guidelines;
* security policy;
* installation and upgrade guide;
* changelog;
* bug-report template;
* feature-request template;
* pull-request template.

These documents are project-specific rather than generic boilerplate. They discuss:

* public API compatibility;
* module ownership;
* SDI/WinAPI concerns;
* recovery;
* snapshot lifetime;
* exact validation steps;
* source export and line endings;
* binary demo treatment;
* security reporting.

### Strong text/binary policy

`.gitattributes` is particularly good. It establishes:

* CRLF working-tree behavior for exported VBA source;
* LF for Markdown and cross-platform scripts;
* binary merge treatment for Office files and `.frx` files;
* explicit GitHub Linguist classification;
* archive exclusions;
* optional diff-driver guidance.

`.gitignore` is similarly tailored and prevents:

* Office lock files;
* local demo binaries;
* release staging output;
* logs and dumps;
* editor metadata;
* secret material;
* Python caches.

### Issue-driven release hardening

The v1.1.0 release cycle used detailed issues to track and fix defects such as:

* title-bar recovery after project reset;
* partial snapshot loss;
* default-false replay after failed capture;
* broken `Err` diagnostics;
* one unusable window aborting a multi-window pass;
* retained snapshot-reference documentation.

The issue descriptions included root-cause analysis, expected behavior, reproduction logic, environment notes, and proposed fixes. This is high-quality engineering practice.

### Detailed release pull request

The release pull request contains:

* a semantic-versioning statement;
* feature inventory;
* defect inventory;
* validation results;
* static validation results;
* diff summary;
* explicit out-of-scope items;
* known follow-ups;
* release checklist.

The merge commit is signed/verified, and the tag and `main` align.

### Security framing

The repository clearly states that hidden UI is not a security boundary. It identifies the actual attack and integrity surfaces:

* VBA macros;
* XLM Ribbon commands;
* WinAPI calls;
* process-wide UI state;
* binary release assets.

This is a credible and appropriately scoped security policy.

## 13.3 Repository-quality weaknesses

### No automated workflow

There is no `.github/workflows` execution gate for:

* source static checks;
* API compatibility;
* test registration;
* documentation freshness;
* reformatter idempotence;
* release manifest validation;
* desktop Excel regression.

The exact release commit had no visible status checks or workflow runs.

For a pure VBA project, full hosted execution is difficult, but zero automation is not the only option. A substantial static gate can run on a normal hosted runner, while the live Excel pack can run on a protected self-hosted Windows runner.

### Manual test evidence is not an artifact

The changelog reports PASS results, but the repository does not publish a machine-readable test file tied to:

* full commit SHA;
* Excel product/version/build;
* Office bitness;
* Windows version;
* exact case inventory;
* pass/fail/skip counts;
* test start/end timestamps;
* cleanup result.

A reviewer cannot distinguish:

```text
complete pass
pass with skipped snapshot cases
pass on one environment only
pass before final source export
```

### Tagged README is stale

The exact `v1.1.0` README still contains:

* a `Release Candidate` badge;
* a “remaining release-maintenance work” section;
* an entirely unchecked release checklist;
* references to the `release/v1.1.0` line;
* tasks to merge, tag, and publish the release, even though the reviewed state is already tagged and merged.

This is a public release-state inconsistency, not a cosmetic detail. The root README is the first artifact most users see.

### No formal review record

The release PR was merged without visible formal GitHub review submissions. A solo-maintained repository can legitimately self-review, but a release with WinAPI and host-state recovery logic would benefit from at least one independent approval or a recorded external review artifact.

### Known follow-ups are not open backlog items

The release PR explicitly lists follow-ups including:

* failure accumulation raising inside an error handler;
* direct `GetLastError` use;
* title-bar owned bits captured once and never refreshed;
* demo not exercising v1.1.0 features.

At review time, the repository had no open issues. Known technical debt should be represented in the issue tracker so it survives release-note context and can be prioritized.

### Release asset provenance

The repository policy says tested `.xlsm` workbooks are release assets. The release process makes a checksum optional rather than required. There is no committed or generated manifest binding the binary to the exact source and validation environment.

### Demo lag

The demo is a primary adoption surface, but it does not expose the headline v1.1.0 functionality. This weakens the practical value of the release asset even if the underlying code is correct.

### Media footprint

Several large images/GIFs make the repository approximately 20 MB, while the production source is only about 0.2 MB. This is not a correctness problem, but it increases clone/archive size and can be improved by:

* optimized WebP/PNG assets;
* GitHub release/attachment hosting;
* a small static screenshot in the repository;
* keeping large animated media outside the source tree.

## 13.4 Recommended repository automation

### Hosted static workflow

Run on every pull request:

```text
1. Verify required files and module names.
2. Verify Option Explicit / Option Private Module policy.
3. Parse Public Sub/Function/Enum inventory.
4. Compare public API against a versioned compatibility manifest.
5. Detect duplicate procedure names.
6. Verify GoTo/Resume labels.
7. Verify #If/#Else/#End If balance.
8. Verify Declare statements and bitness branches.
9. Verify source line endings and encoding.
10. Run reformatter --check and idempotence tests.
11. Verify README/changelog version and release-state markers.
12. Verify test runner/case registration.
13. Verify no tracked Office lock file or ignored release binary.
14. Run Markdown link checks.
```

### Protected Excel workflow

On trusted branches/tags only:

```text
1. Start isolated Excel on a dedicated Windows runner.
2. Create a clean macro-enabled workbook.
3. Import all four production modules and test module.
4. Compile VBAProject.
5. Run a machine-readable CI bridge.
6. Execute every mandatory runner, including SDI multi-window tests.
7. Reject skipped mandatory cases.
8. Verify cleanup and emergency recovery.
9. Export JSON/CSV/TXT result artifacts.
10. Record Excel/Office/Windows environment.
```

### Release workflow

For a tag:

```text
1. Require green static and Excel checks on the exact SHA.
2. Build demo from exact tagged source.
3. Run the same regression suite against the demo workbook.
4. Compute SHA-256.
5. Generate release manifest.
6. Publish workbook, checksum, manifest, and test evidence.
7. Verify README status is stable-release, not release-candidate.
```

## 13.5 Repository-quality verdict

> \*\*The repository is professionally written and thoughtfully governed, but its controls are primarily documentary and manual. The next maturity step is not more prose; it is executable enforcement, evidence provenance, and release-state synchronization.\*\*

\---

# 14\. Documentation and release-management review

## 14.1 README

### Strengths

The README provides:

* a clear statement of purpose;
* managed-surface table;
* quick-start examples;
* full public API table;
* target-scope explanation;
* architecture diagram;
* snapshot lifecycle;
* title-bar ownership explanation;
* diagnostics contract;
* regression instructions;
* recovery command;
* requirements and limitations.

It is substantially above the norm for VBA repositories.

### Defects

#### Release-state drift

The tagged README still describes a pre-release state. This should be corrected immediately in `main` and in the next patch tag.

#### SDI imprecision

The table describes:

```text
Ribbon    -> Excel application
Title Bar -> Excel main window
```

For modern Excel, each workbook has a top-level window and Ribbon. `Application.Hwnd` is the active window's handle. The wording should distinguish:

```text
process-wide object-model property
active SDI workbook frame
per-workbook Ribbon instance
all open workbook windows
```

#### “Show every managed element” overstatement

`UI\_ShowExcelUI` shows all targetable object-model window elements under the default scope, but title-bar execution is active-frame-dependent. The current phrase should be qualified until the SDI model is fixed.

#### XLM requirement

Requirements should identify the fixed Excel 4 macro command used for Ribbon control and note that organizational policy can block it.

## 14.2 Installation guide

`INSTALLATION.md` is excellent. It correctly states:

* all four production modules are required;
* import order is recommended, not semantically required after all modules exist;
* v1.0.1 single-module installations must be replaced as a complete set;
* mixed internal versions are unsafe;
* snapshot object references have a release lifecycle;
* `Workbook\_BeforeClose` is a natural clear point;
* the binary demo belongs in releases, not Git.

The main improvement is to add an SDI behavior section with a two-workbook example and explicit title-bar/Ribbon scope.

## 14.3 Changelog

The changelog entry for v1.1.0 is detailed and candid. It records:

* added features;
* architectural changes;
* fixed defects;
* compatibility;
* manual validation;
* known limitations.

The principal evidence gap is environment precision. It should record:

```text
Excel product/channel
Excel version
Excel build
Office bitness
Windows version/build
window count and SDI scenario
add-ins enabled/disabled
workbook type
```

## 14.4 Contribution guide

The contribution guide is strong in:

* module ownership;
* acyclic dependency rules;
* source workflow;
* compatibility requirements;
* snapshot and title-bar review questions;
* validation sequence;
* line-ending discipline.

It should add:

* SDI frame identity requirements;
* mandatory issue creation for release-deferred defects;
* no green release result when mandatory cases were skipped;
* exact-environment evidence requirements;
* required independent approval or attached external review for tagged releases.

## 14.5 Pull-request template

The template asks the right questions about:

* public behavior;
* module dependencies;
* snapshot and recovery;
* WinAPI methods;
* diagnostics;
* validation environment.

The release PR was detailed, but there were no visible formal review submissions. Template quality cannot substitute for approval evidence.

\---

# 15\. Security and platform assessment

## 15.1 Security posture

No high-severity security vulnerability was identified in the reviewed production source.

Positive properties include:

* no network access;
* no shell execution;
* no dynamic DLL selection;
* no arbitrary WinAPI target supplied by a worksheet/user input;
* no arbitrary XLM command construction;
* no installer or background process;
* no third-party DLL;
* no credential storage;
* no automatic update mechanism;
* no modal production UI;
* explicit warning that UI hiding is not access control.

## 15.2 Integrity risks

The main risk category is host-state integrity, not confidentiality:

* wrong window receives title-bar state under SDI;
* partial frame mutation after failed refresh;
* diagnostic failure can mask original host failure;
* tests can leave UI altered when cleanup failure is suppressed;
* retained COM references can prolong object lifetime;
* an external add-in can race or rewrite the same state.

## 15.3 XLM Ribbon command

The Ribbon command is fixed and therefore not an input-injection vector. Nevertheless:

* XLM is a legacy macro mechanism;
* enterprise policy may restrict it;
* CommandBars/Ribbon behavior under SDI should be validated on supported Office channels;
* the component should report clear failure text when policy blocks execution.

## 15.4 WinAPI correctness

The implementation correctly:

* uses `LongPtr` where required;
* separates 32-bit and 64-bit declarations;
* validates zero handles;
* disambiguates ambiguous zero returns;
* merges only owned bits;
* uses `SWP\_FRAMECHANGED`;
* avoids changing position, size, or Z-order.

The next corrections are transactionality and explicit SDI handle identity, not basic declaration repair.

## 15.5 Supply-chain posture

Plain-text source is easy to inspect. The `.xlsm` release asset is executable content and should be:

* generated from exact tagged source;
* scanned according to organizational policy;
* accompanied by SHA-256;
* bound to a release manifest;
* tested after final save;
* never treated as trustworthy merely because it appears in a release.

## 15.6 Security verdict

**8.6 / 10** for policy and source posture.

The identified P1 is a correctness/integrity defect rather than an exploit in the conventional sense, unless a concrete adversarial availability or integrity impact is demonstrated.

\---

# 16\. Findings summary

|ID|Severity|Area|Finding|
|-|-|-|-|
|ICR-UI-P1-01|**P1**|SDI / snapshot correctness|Title-bar snapshot state is not bound to the captured SDI window; activation changes can restore the state to the wrong top-level frame|
|ICR-UI-P2-01|**P2**|Ribbon / SDI contract|Ribbon state is modeled as one application Boolean although modern Excel has one Ribbon per workbook window; cross-window behavior is unspecified and untested|
|ICR-UI-P2-02|**P2**|Error handling|`UI\_RuntimeAddFailure` can raise while handling another failure and mask the original error|
|ICR-UI-P2-03|**P2**|WinAPI transaction|A successful style write followed by failed frame refresh leaves partial mutation; a retry can no-op without refreshing|
|ICR-UI-P2-04|**P2**|Title-bar ownership|Owned title-bar bits are cached once per currently observed handle and can become stale or be displaced when active SDI windows change|
|ICR-UI-P2-05|**P2**|CI / release evidence|The exact release commit has no automated workflow/status evidence; manual PASS claims are not machine-readable or environment-complete|
|ICR-UI-P2-06|**P2**|Documentation|The tagged README still identifies v1.1.0 as a release candidate and retains an unchecked pre-release checklist|
|ICR-UI-P2-07|**P2**|Regression harness|`RunAll` excludes the dedicated identity runner, mandatory cases may be skipped, and cleanup failures can be suppressed|
|ICR-UI-P2-08|**P2**|Demo / adoption|The demo does not exercise window targeting or structured v1.1.0 result APIs|
|ICR-UI-P3-01|**P3**|DLL diagnostics|The WinAPI layer calls `GetLastError` directly rather than using the VBA-native immediate `Err.LastDllError` channel|
|ICR-UI-P3-02|**P3**|Verification|Ribbon, property, and title-bar setters usually do not read back the achieved state|
|ICR-UI-P3-03|**P3**|Quiet scope|`ScreenUpdating` change state is recorded without confirming the write succeeded|
|ICR-UI-P3-04|**P3**|Tooling|The regex reformatter can rewrite matching text inside strings/comments and lacks tests, idempotence, and check mode|
|ICR-UI-P3-05|**P3**|Release provenance|The binary demo and manual test evidence are not bound by a required checksum/manifest to the exact source and environment|
|ICR-UI-P3-06|**P3**|Backlog governance|Known v1.1.1 follow-ups are recorded in the merged PR but not represented by open issues|
|ICR-UI-P3-07|**P3**|Compatibility wording|“No migration required” is true for caller code but not for deployment from the prior single-module package|
|ICR-UI-P3-08|**P3**|Repository footprint|Large media assets dominate a small source repository and can be optimized or moved outside Git history|

\---

# 17\. Detailed findings

## ICR-UI-P1-01 — Title-bar snapshot restoration is not SDI identity-safe

### Severity

**P1 — material correctness and release-claim defect**

### Affected public behavior

```text
UI\_CaptureExcelUIState
UI\_CaptureExcelUIState\_WithResult
UI\_ResetExcelUIToSnapshot
UI\_ResetExcelUIToSnapshot\_WithResult
```

The same underlying scope issue also affects selective and show/hide title-bar operations.

### Root cause

Snapshot state contains only:

```vb
m\_SnapshotTitleBarKnown
m\_SnapshotTitleBarVisible
```

The title-bar worker resolves:

```vb
xlHnd = Application.Hwnd
```

on every read/write.

Microsoft documents that under Excel SDI:

```text
Application.Hwnd returns the active workbook window's handle.
```

The capture and restore calls can therefore resolve different handles.

### Reproduction logic

```text
1. Open workbook windows A and B in the same Excel process.
2. Give their title bars distinct states.
3. Activate A.
4. Capture the UI snapshot.
5. Activate B.
6. Restore the snapshot.
7. Observe that the captured A title-bar Boolean is applied through B's active HWND.
```

### Impact

* captured state can be written to the wrong top-level Excel window;
* the original captured frame remains unrestored;
* no mismatch is reported;
* the “identity-safe snapshot” claim is only partially true;
* show/hide wrapper scope is inconsistent across managed elements;
* multi-monitor/multi-workbook users are most exposed.

### Why severity is P1

The defect is:

* inside ordinary supported modern Excel behavior;
* deterministic from source and platform contract;
* silent when API calls succeed;
* capable of modifying the wrong host object;
* directly related to a headline v1.1.0 feature: identity-safe restoration.

### Required remediation

1. Define exact title-bar scope.
2. Capture the target `Window` reference and/or top-level `HWND`.
3. Pass the explicit target into title-bar read/write helpers.
4. Store owned-bit baselines per `HWND`.
5. Probe liveness/identity before restore.
6. Report a missing captured frame rather than redirecting state.
7. Add two-window SDI regression cases.
8. Qualify README claims until the fix is released.

### Suggested design

```vb
Private Type tTitleBarSnapshot
#If VBA7 Then
    Hwnd            As LongPtr
#Else
    Hwnd            As Long
#End If
    WindowRef       As Object
    WindowLabel     As String
    Known           As Boolean
    Visible         As Boolean
    OwnedStyleBits  As LongPtr   'conditional type in actual code
End Type
```

The internal title-bar surface should accept an explicit handle:

```vb
UI\_TryGetTitleBarVisibleForHwnd
UI\_TrySetTitleBarVisibleForHwndIfNeeded
```

The existing no-argument helpers can remain as active-window compatibility wrappers.

\---

## ICR-UI-P2-01 — Ribbon scope under SDI is unspecified and unverified

### Severity

**P2 — public contract and assurance gap**

### Root cause

Modern Excel provides one Ribbon UI per workbook window. The component stores and restores one Ribbon Boolean and uses the current application/CommandBars context.

### Risk

A user can reasonably interpret “Ribbon — Excel application” to mean one state across every workbook window. That is not automatically guaranteed by the modern Excel UI model.

### Required work

1. Build a two-workbook SDI test workbook.
2. Characterize Ribbon read/write behavior on supported Excel builds.
3. Test active-window switching before and after capture.
4. Test a workbook opened after Ribbon hide.
5. Decide whether state is active-window, all-window, or cached-and-propagated.
6. Document exact behavior.
7. If all-window synchronization is promised, implement activation/enumeration logic.

### Closure criterion

A P2 contract can be closed either by:

* implementing deterministic all-window behavior; or
* narrowing the public contract to the empirically confirmed active-window behavior.

\---

## ICR-UI-P2-02 — Failure accumulation can mask the original failure

### Severity

**P2 — fail-soft contract defect**

### Root cause

`UI\_RuntimeAddFailure` grows a dynamic array without an internal error boundary and is called from failure paths.

### Failure sequence

```text
Original Excel/WinAPI failure
    -> UI\_RuntimeHandleFailure
        -> UI\_RuntimeAddFailure
            -> ReDim Preserve or assignment raises
                -> original failure detail lost or replaced
```

### Impact

* structured APIs can raise unexpectedly;
* fire-and-forget paths can stop processing;
* `ScreenUpdating` cleanup may be bypassed;
* failure count/list invariants may break;
* the most important diagnostic—the original failure—can disappear.

### Remediation

* make result accumulation non-raising;
* preserve original detail in local scalars before recording;
* use a minimal fallback if list growth fails;
* add a fault-injection test;
* document whether failure-count overflow/list-allocation failure is itself returned.

\---

## ICR-UI-P2-03 — Title-bar write and frame refresh are not transactional

### Severity

**P2 — host-state integrity and recovery**

### Root cause

`SetWindowLong` can succeed before `SetWindowPos(... SWP\_FRAMECHANGED)` fails.

The module then returns failure but retains no state that the frame refresh is incomplete.

### Retry defect

On retry, the desired style can already equal the current style. The code short-circuits before `SetWindowPos`, so the frame cache can remain stale.

### Remediation options

#### Option 1 — Pending refresh flag

```text
style write success + refresh failure -> mark HWND pending
next call -> refresh before no-op evaluation
```

#### Option 2 — Rollback

```text
refresh failure -> restore prior style -> attempt refresh -> return original failure
```

#### Option 3 — Always refresh on matching bits after prior failure

This is the minimum corrective behavior.

### Regression requirement

The API boundary needs a test seam so `SetWindowPos` can be forced to fail after a successful style write. The next call must prove that refresh is retried.

\---

## ICR-UI-P2-04 — Title-bar baseline can become stale or be displaced

### Severity

**P2 — interoperability and state-ownership defect**

### Root cause

The module captures owned bits once for the currently observed `Application.Hwnd`. It does not refresh the baseline after another component legitimately changes those owned bits.

Under SDI, activating another workbook replaces the singleton handle/baseline pair.

### Impact

* later show can restore stale frame bits;
* baseline for window A is lost when B becomes active;
* interoperability with other add-ins is weaker than the owned-bit narrative suggests;
* full owned mask fallback may restore controls that were deliberately absent.

### Remediation

* maintain state per `HWND`;
* distinguish “current visible baseline” from “component last-applied state”;
* recapture when the component does not currently own a hidden state;
* document conflict policy with other add-ins;
* add cross-component simulation tests using synthetic style-bit changes.

\---

## ICR-UI-P2-05 — No automated release evidence on the exact SHA

### Severity

**P2 — release engineering and reproducibility**

### Current state

The release documents manual PASS results, but the exact commit has no visible automated status checks or workflow runs.

### Why this matters

A pure VBA project is especially vulnerable to drift between:

```text
repository export
locally imported workbook
compiled workbook
tested workbook
tagged commit
published demo asset
```

Manual evidence does not prove these are identical.

### Remediation

* add hosted static validation;
* add protected self-hosted Excel regression;
* export a machine-readable result artifact;
* bind it to the full SHA and environment;
* require green checks before tag publication;
* include the test artifact in the release.

\---

## ICR-UI-P2-06 — Tagged README remains in release-candidate state

### Severity

**P2 — public documentation and release governance**

### Evidence

The tagged README contains:

```text
Status: Release Candidate
Remaining release-maintenance work
unchecked release checklist
merge/tag/publish tasks
release/v1.1.0 line wording
```

### Impact

* users cannot tell whether the tag is stable or pre-release;
* documentation contradicts the changelog and repository state;
* the release process did not include a final documentation freeze;
* trust in other status claims is reduced.

### Remediation

* change status to stable/current release;
* remove or archive the pre-release checklist;
* move future release tasks to an issue/project;
* add CI that checks README version/status markers against the tag/changelog.

\---

## ICR-UI-P2-07 — Regression runner semantics permit incomplete green results

### Severity

**P2 — assurance completeness**

### Problems

1. `RunAll` excludes the dedicated snapshot-identity runner.
2. Snapshot cases are skipped when a prior snapshot exists.
3. Skips do not convert release validation to incomplete.
4. Harness cleanup uses collection indexes.
5. Cleanup errors are suppressed.
6. PASS is logged before cleanup completes.

### Remediation

* create one release-certification runner;
* retain test-window object identities;
* reject pre-existing snapshots in certification mode;
* count pass/fail/skip/cleanup;
* fail on mandatory skip;
* fail separately on cleanup failure;
* export machine-readable evidence.

\---

## ICR-UI-P2-08 — Demo does not exercise v1.1.0's main features

### Severity

**P2 — release completeness and adoption quality**

### Missing user journeys

```text
select target scope
apply to active window
apply to active workbook windows
display structured failure count/list
show snapshot-capture partial failures
show restore identity failure
explain SDI title-bar target
```

### Remediation

Update both demo source and release workbook, then test the final binary against the exact tag. Include screenshots that reflect the current controls.

\---

## ICR-UI-P3-01 — Prefer immediate `Err.LastDllError` handling

### Severity

**P3 — platform hardening**

Microsoft's VBA guidance is to inspect `Err.LastDllError` immediately after a DLL call that returns a failure indicator. The direct `GetLastError` declaration can work, but the project should standardize one approach and test it.

Recommended:

```vb
SetLastError 0
Result = SetWindowLong...
LastErr = Err.LastDllError
```

Capture it before any other VBA or API operation.

\---

## ICR-UI-P3-02 — State writes are not read back

### Severity

**P3 — assurance precision**

The component often interprets “no runtime error” as success. Add optional strict verification in the regression/release path:

```text
write
settle
read
compare
```

Production calls can remain best effort to avoid excessive host interaction.

\---

## ICR-UI-P3-03 — Quiet-scope state flag is optimistic

### Severity

**P3 — cleanup precision**

Only mark the quiet scope as changed when the `ScreenUpdating=False` write succeeds. Otherwise cleanup should not imply ownership of a state transition that did not occur.

\---

## ICR-UI-P3-04 — Reformatter is not token-safe

### Severity

**P3 — maintenance tooling**

The release-specific transformation was independently checked at statement level, but the general tool can alter matching text inside comments or strings. Add tokenizer-aware processing and golden tests before treating it as a general source formatter.

\---

## ICR-UI-P3-05 — Release asset and evidence provenance are incomplete

### Severity

**P3 — supply chain and reproducibility**

Make checksums and release manifests mandatory, not optional. Publish the exact final workbook tested after final save.

\---

## ICR-UI-P3-06 — Known debt should be in the issue tracker

### Severity

**P3 — maintenance governance**

Convert every explicit merged-PR follow-up into an issue with severity, acceptance criteria, and target milestone. A zero-open-issue dashboard should not conceal known deferred defects.

\---

## ICR-UI-P3-07 — Compatibility wording conflates code and deployment

### Severity

**P3 — release communication**

Use:

```text
Caller-code migration: none
Installation/package migration from v1.0.1: replace one module with four
```

This preserves the valid SemVer claim without understating deployment work.

\---

# 18\. Prioritized remediation plan

## Release Gate 1 — Correct SDI frame identity

1. Decide active-window versus all-window title-bar semantics.
2. Add explicit `HWND`/`Window` parameters to internal title-bar helpers.
3. Store snapshot title-bar identity.
4. Store title-bar baseline state per `HWND`.
5. Detect missing/closed captured frames.
6. Add two-window SDI regression cases.
7. Update public scope documentation.

**Exit criterion:** activating another workbook between capture and restore never redirects title-bar state.

## Release Gate 2 — Make failure paths and frame updates recoverable

1. Make failure accumulation non-raising.
2. Preserve original error detail under diagnostic degradation.
3. Track title-bar frame-refresh pending state.
4. Retry or roll back after `SetWindowPos` failure.
5. Add deterministic fault injection.

**Exit criterion:** every expected failure returns through the documented structured contract, and a failed frame refresh cannot become a false no-op success.

## Release Gate 3 — Define Ribbon behavior under SDI

1. Test at least two workbook windows.
2. Test active-window switching.
3. Test new-window behavior after hide/show.
4. Test relevant Office channels and policies.
5. Document exact scope.
6. Implement propagation/caching if application-wide behavior is promised.

**Exit criterion:** the README's Ribbon scope statement is backed by repeatable evidence.

## Release Gate 4 — Create a real release-certification runner

1. Make one runner execute every mandatory case.
2. Reject pre-existing snapshot state or isolate the test workbook.
3. Track pass/fail/skip/cleanup counts.
4. Retain test-window object identities.
5. Treat cleanup failure as failure.
6. Emit JSON/CSV/TXT evidence.

**Exit criterion:** one command produces an unambiguous complete/incomplete/pass/fail result.

## Release Gate 5 — Automate repository controls

1. Add hosted static checks.
2. Add protected Windows/Excel CI where operationally feasible.
3. Verify public API compatibility against a manifest.
4. Verify test registration and source structure.
5. Verify README/changelog/release status.
6. Verify formatter idempotence.
7. Require checks on release PRs/tags.

**Exit criterion:** a tagged release cannot be published from a SHA that has no exact validation evidence.

## Release Gate 6 — Synchronize public release artifacts

1. Correct README stable-release status.
2. Remove the obsolete unchecked release checklist.
3. Update demo controls for v1.1.0 features.
4. Build the demo from exact source.
5. Publish checksum and release manifest.
6. Record exact validation environment.
7. open issues for all deferred work.

**Exit criterion:** source, docs, demo, evidence, and tag all describe the same release state.

## Release Gate 7 — Harden maintenance tooling

1. Add `--check`, `--write`, and `--diff` modes.
2. Add argument validation.
3. Add tokenizer-aware comment/string handling.
4. Add encoding policy.
5. Add golden tests and idempotence tests.
6. Run tooling checks in CI.

\---

# 19\. Recommended v1.1.1 scope

A focused v1.1.1 should prioritize correctness and release assurance rather than new UI features.

## Must fix

```text
P1 SDI title-bar identity
failure accumulator fail-safety
title-bar frame-refresh retry/rollback
README stable-release state
complete release runner semantics
```

## Should fix

```text
Ribbon SDI characterization and documentation
per-HWND title-bar baseline cache
exact test-environment artifact
static CI gate
open issue backlog for deferred findings
```

## Could fix

```text
demo target-scope controls
formatter check mode and tests
media optimization
strict state-readback mode
snapshot metadata/status helper
```

## Avoid in v1.1.1

To keep the patch reviewable, do not combine the corrective release with:

* a new object-oriented public API;
* macOS support;
* RibbonX packaging;
* persistent snapshot storage;
* COM add-in architecture;
* unrelated demo-framework redesign.

\---

# 20\. Release-readiness assessment

## Suitable now

The component is suitable for:

* controlled Windows Excel workbooks;
* application-style workbook shells;
* teaching and demonstrations;
* internal productivity tools;
* environments where the exact Excel/Office build is tested;
* single-active-window workflows;
* per-window Headings, Workbook Tabs, and Gridlines control;
* snapshot restore where title-bar cross-window identity is not relied upon;
* governed callers that use `\_WithResult` APIs and maintain an accessible recovery macro.

## Required operating controls

For production use of v1.1.0:

1. pin the exact tag/commit;
2. import all four production modules from the same release;
3. compile in the target workbook/add-in;
4. run all four documented runners in a clean Excel session;
5. verify no mandatory snapshot cases were skipped;
6. test with the actual workbook-window count and Office bitness;
7. keep `UI\_ShowExcelUI` independently accessible;
8. clear snapshots at the intended lifecycle point;
9. avoid relying on title-bar snapshot identity across active-window changes;
10. characterize Ribbon behavior in the target Office environment;
11. do not treat hidden UI as security or access control.

## Not suitable for an unconditional claim

Do not currently claim:

```text
fully identity-safe restoration of every managed UI element across multiple SDI windows
one application-wide title-bar state
independently certified 32-bit and 64-bit behavior on the exact tag
release evidence enforced by CI
cryptographically source-bound demo binary
```

## Release-readiness score

```text
7.5 / 10
```

The release is usable and thoughtfully engineered, but one cross-window correctness issue and several evidence/governance gaps prevent a stronger certification.

\---

# 21\. Final verdict

VBA Excel UI v1.1.0 demonstrates a level of engineering discipline rarely seen in small pure-VBA repositories:

* a compact, stable public API;
* explicit visibility and target enums;
* cohesive modular architecture;
* fail-soft orchestration;
* structured diagnostics;
* partial-snapshot semantics;
* retained object identity for window properties;
* careful WinAPI bit ownership;
* emergency recovery logic;
* a meaningful regression harness;
* excellent installation and security documentation;
* detailed release issues and changelog history.

The release's central remaining weakness is an architectural mismatch between modern Excel's SDI window model and the singleton title-bar/Ribbon state model. The title-bar case is a confirmed correctness defect: capture and restore can target different top-level windows because both resolve the active `Application.Hwnd` independently.

The repository itself is polished but not yet self-enforcing. Its documentation, templates, and release narrative are stronger than its automated validation, release provenance, and status synchronization.

> \*\*Final overall score: 8.0 / 10\*\*  
> \*\*Production-code quality: 8.3 / 10\*\*  
> \*\*Repository quality: 7.6 / 10\*\*  
> \*\*Classification: strong professional VBA component requiring one material SDI correctness fix and a targeted upgrade from manual process controls to executable release assurance.\*\*

\---

# Appendix A — Representative assurance inventory

|Area|v1.1.0 evidence|
|-|-|
|Exact release commit|`96360379a4bca7703cf649a69a2162961dfa6c9e`|
|Production modules|4|
|Public callable members|10|
|Public enums|2|
|Managed UI elements|8|
|Window target scopes|3|
|Regression runners|4|
|Regression cases|22 reported|
|Manual compile|PASS reported|
|Manual runners|4 PASS reported|
|Manual recovery|PASS reported|
|Hosted workflows|none|
|Exact-SHA status checks|none visible|
|Machine-readable test artifact|none identified|
|Binary checksum manifest|none identified|
|Formal PR review submissions|none visible|
|Open issues at review time|0|

\---

# Appendix B — Suggested GitHub issues

1. `Bind title-bar snapshot state to the captured SDI Window/HWND`
2. `Store title-bar owned-bit baselines per HWND`
3. `Characterize and document Ribbon visibility across SDI workbook windows`
4. `Make UI\_RuntimeAddFailure non-raising inside error paths`
5. `Retry or roll back title-bar frame refresh after SetWindowPos failure`
6. `Make Test\_EXCEL\_UI\_RunAll execute the snapshot-identity pack`
7. `Fail release validation when mandatory test cases are skipped`
8. `Treat test-harness cleanup failure as a failed run`
9. `Retain exact Window identities in test cleanup`
10. `Add hosted static validation for exported VBA source`
11. `Add protected desktop-Excel regression workflow and result artifact`
12. `Generate a release manifest binding source, demo workbook, checksum, and environment`
13. `Update the tagged README from release-candidate to stable-release state`
14. `Extend the demo with TargetScope and structured-result controls`
15. `Harden tools/reformat.py with token-safe parsing and golden tests`
16. `Document Excel 4 macro policy requirements for Ribbon control`
17. `Separate caller-code compatibility from installation migration wording`
18. `Optimize or externalize large repository media assets`

\---

# Appendix C — Evidence confidence

|Conclusion|Confidence|
|-|-|
|Four-module architecture and dependency ownership|High|
|Public API compatibility assessment|High|
|Per-window retained-object snapshot design|High|
|SDI title-bar identity defect|High|
|Title-bar frame-refresh partial-transaction risk|High from source and Win32 contract|
|Failure-accumulator masking risk|High from source|
|Ribbon cross-window behavior is under-specified|High|
|Exact Ribbon behavior on each Office build|Not independently executed|
|Manual regression results recorded in repository|High as committed claims|
|Exact desktop-Excel regression result independently reproduced|Not performed|
|Exact release asset source binding/checksum|Not independently verified|
|Branch-protection configuration|Not independently verified|
|Performance impact|Low concern; no formal benchmark performed|

\---

# Appendix D — Review limitations

This report is a static, source-and-platform-contract review. It does not replace:

* compilation in the target VBA project;
* live testing on supported Excel channels;
* 32-bit and 64-bit Office validation;
* Windows accessibility and multi-monitor checks;
* interaction testing with other add-ins;
* enterprise XLM policy testing;
* binary workbook malware scanning;
* an independent human code review before high-governance deployment.

The findings should be converted into permanent regression cases and closed against exact commits, not only acknowledged in release prose.

