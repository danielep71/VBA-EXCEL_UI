# Independent Code and Repository Review — VBA Excel UI v1.1.1

> **Repository:** [`danielep71/VBA-EXCEL_UI`](https://github.com/danielep71/VBA-EXCEL_UI)  
> **Release reviewed:** [`v1.1.1`](https://github.com/danielep71/VBA-EXCEL_UI/releases/tag/v1.1.1)  
> **Tag / merge commit reviewed:** [`b4df9be28a19ccc4ffc76df1a2cadfa2423b31e1`](https://github.com/danielep71/VBA-EXCEL_UI/commit/b4df9be28a19ccc4ffc76df1a2cadfa2423b31e1)  
> **Exact tree reviewed:** `90e3d5fdff7007b44bce69704ef3757b3e5b3484`  
> **Comparison baseline:** `v1.1.0`  
> **Review date:** 2026-08-20   
> **Suggested repository path:** `docs/INDEPENDENT_CODE_REVIEW_V1.1.1_2026-08-20.md`

---

## 1. Executive assessment

### Overall repository score: **8.5 / 10**

### Production-code quality score: **8.7 / 10**

### Repository-quality score: **8.0 / 10**

### Architecture and modularity score: **9.5 / 10**

### Regression and release-certification score: **8.1 / 10**

### CI and release-engineering score: **8.2 / 10**

`v1.1.1` is a substantial and technically serious corrective release. It is not a cosmetic patch over `v1.1.0`. The implementation now contains:

* explicit-handle title-bar read and write operations;
* title-bar snapshot state bound to the frame captured under Excel's Single Document Interface;
* a per-window frame-state registry;
* self-healing title-bar baselines while the component does not own a hidden state;
* retryable non-client-frame refresh debt;
* a failure accumulator designed not to raise while another error is being handled;
* a dedicated SDI title-bar regression runner;
* a release-certification runner with completeness, failure, skip and cleanup counters;
* machine-readable and human-readable certification output;
* a hosted static repository gate;
* a versioned project-public-surface manifest;
* measured Ribbon behavior under SDI;
* much stronger README, installation, contribution and release documentation.

The central `v1.1.0` title-bar defect has been corrected in the intended direction. Snapshot restore no longer reads `Application.Hwnd` at restore time and silently applies the captured title-bar state to whichever workbook happens to be active. The snapshot now retains the captured handle and the associated Excel `Window`, validates them before writing, and uses the explicit-handle title-bar path.

The release also closes the most important transactional weakness in the title-bar subsystem. A successful `GWL_STYLE` write followed by a failed `SetWindowPos(... SWP_FRAMECHANGED)` is now recorded as outstanding refresh debt and retried before a later no-op decision can report success.

The release is nevertheless **not free of material correctness and certification defects**.

The most important remaining production defect is already measured and recorded by the repository itself:

> **Ribbon snapshot restoration is still not window-identity-safe. A Ribbon value captured from workbook window A is restored to whichever window is active, potentially window B, while A remains unrestored and the operation reports success.**

This is not merely an undocumented uncertainty. The repository's own SDI probe reproduced it, and open issue `#23` correctly labels it **P1**. The README now discloses it, which is a significant governance improvement, but documentation does not make the behavior correct. A caller using the snapshot API in a multi-window session can still receive a silent wrong-target result.

A second material defect remains in the title-bar state registry. Registry lookup is keyed only by the numeric `hWnd`. It searches for a matching value before compacting dead entries. If Windows destroys a workbook window and later reuses the same handle value, the new window can inherit stale `OwnedStyleBits`, `ComponentHidden` and `RefreshPending` state from the destroyed one. This was identified in the release pull request and remained unresolved when the release was merged.

The new release-certification runner is directionally excellent but not yet reliable enough to serve as the sole behavioral release gate:

* it treats a correctly restored `ScreenUpdating = False` baseline as cleanup failure;
* its own error handler logs before preserving `Err`, clearing the original error before attempting to re-raise it;
* it checks workbook count, but not Excel window count, so an extra `Workbook.NewWindow` can leak without failing cleanup;
* it does not compare the complete managed UI state after cleanup, while the underlying test cleanup suppresses restore errors;
* its JSON identifies the host but not the tag, commit SHA or source hashes that were executed.

Repository quality has improved sharply but is held back by three visible contradictions:

1. the public GitHub Wiki remains materially obsolete and gives installation instructions that are invalid for the four-module package;
2. the latest published demo is still the `v1.1.0` asset, does not expose the principal current features, and has broken presets;
3. the release pull request was merged seconds after two P2 review comments were posted, leaving both threads unresolved in the tagged release.

### Independent verdict

> **`v1.1.1` is a strong corrective release with a mature core architecture and much better assurance infrastructure. It is suitable for controlled single-window use and for multi-window use that does not rely on Ribbon snapshot restoration. It should not yet be described as fully window-identity-safe across the complete managed UI surface, and its release-certification runner needs a corrective pass before it can be treated as authoritative automation.**

---

## 2. What materially improved from v1.1.0

The score improvement from **8.0** to **8.5** is justified by real engineering changes rather than documentation volume.

|`v1.1.0` review area|`v1.1.1` status|Assessment|
|-|-|-|
|Title-bar snapshot restored through active `Application.Hwnd`|**Corrected**|Captured handle plus retained `Window`; explicit-target restore|
|One process-wide title-bar baseline|**Corrected in normal operation**|Per-`hWnd` registry and baseline refresh|
|Style write followed by failed frame refresh could become a false no-op|**Corrected**|Refresh debt is recorded and retried|
|Failure-list growth could raise inside an error handler|**Corrected**|Count remains authoritative; append is guarded and truncation is surfaced where possible|
|Ribbon scope under SDI unmeasured|**Measured**|Active-window-only model established on one host|
|Ribbon snapshot identity behavior unknown|**Confirmed defect; not fixed**|Open P1 issue `#23`|
|`Test_EXCEL_UI_RunAll` could be partial without an unambiguous verdict|**Partially corrected**|New release-certification runner, but runner contains defects discussed below|
|No hosted static gate|**Corrected**|`tools/check_repo.py` and GitHub Actions workflow added|
|No versioned public-surface inventory|**Partially corrected**|Name-level manifest added; signatures and enum values remain unguarded|
|README retained pre-release state|**Corrected**|Stable release state, current architecture and known limitations documented|
|Demo did not exercise `v1.1.0` features|**Open**|Explicitly deferred to `v1.2.0` as issue `#22`|
|Regex formatter had no check mode|**Partially corrected**|`--check` and `--write` added; token-safety issue remains|
|No durable review backlog|**Improved**|P1 Ribbon and P2 demo issues are explicitly tracked|

The important point is that the design has converged. The remaining work does not require abandoning the four-module architecture. It requires completing the SDI contract, hardening identity and evidence handling, and removing documentation duplication that is already drifting.

---

# 3. Review scope and methodology

## 3.1 Exact source basis

The review was performed against the exact tree referenced by the `v1.1.1` release tag and merge commit.

The release pull-request head and the tagged merge commit resolve to the same tree:

```text
90e3d5fdff7007b44bce69704ef3757b3e5b3484
```

This matters because the hosted static workflow passed on the pull-request head. Since the tag carries the same tree, the static result applies to the reviewed source tree even though no separate tag-triggered workflow is present.

The reviewed scope included:

### Production VBA

* `src/M_EXCEL_UI.bas`
* `src/M_EXCEL_UI_RUNTIME.bas`
* `src/M_EXCEL_UI_SNAPSHOT.bas`
* `src/M_EXCEL_UI_TITLEBAR.bas`

### Regression and release certification

* `test/M_EXCEL_UI_REGRESSION_TESTS.bas`

### Demo source

* `demo/M_EXCEL_UI_DEMO.bas`
* `demo/M_DEMO_BUILDER.bas`

### Repository automation

* `.github/workflows/static-checks.yml`
* `tools/check_repo.py`
* `tools/reformat.py`
* `tools/public_api_manifest.txt`

### Documentation and governance

* `README.md`
* `INSTALLATION.md`
* `CHANGELOG.md`
* `CONTRIBUTING.md`
* `SECURITY.md`
* `CODE_OF_CONDUCT.md`
* `.github` issue and pull-request templates
* `docs/RIBBON_SDI_BEHAVIOR.md`
* the prior independent `v1.1.0` review
* the public GitHub Wiki home page
* release pull request `#24`, its review threads and workflow state
* current open issues and repository metadata

## 3.2 Execution boundary

Desktop Microsoft Excel was not available in the review environment. The reviewer therefore did **not**:

* import the modules into the Visual Basic Editor;
* execute `Debug -> Compile VBAProject`;
* run `Test_EXCEL_UI_RunReleaseCertification` independently;
* reproduce the Ribbon or title-bar SDI tests interactively;
* execute 32-bit Office paths;
* inspect the binary demo workbook directly.

The review distinguishes among:

* **source-confirmed behavior**, established from the committed control flow;
* **repository-reported Excel evidence**, particularly the certification block in `CHANGELOG.md`;
* **host measurements committed by the project**, particularly `docs/RIBBON_SDI_BEHAVIOR.md`;
* **review-process evidence**, including unresolved pull-request threads;
* **operational state not independently verified**, including branch protection and a behavioral test run on the exact tag outside the repository's recorded evidence.

## 3.3 External platform contracts used

The review uses two stable platform facts:

* modern Excel's Ribbon and top-level workbook frames are window-sensitive under SDI, as measured by the repository's own probe;
* VBA's `Err` object is cleared by an `On Error` statement, which is also the reason the production code correctly captures `Err.Number`, `Err.Description`, `Err.Source` and `Erl` before entering protected diagnostic formatting.

No conclusion depends on unpublished or proprietary behavior.

---

# 4. Hard repository metrics

## 4.1 Source scale

|Area|Files|Exact repository bytes|
|-|-:|-:|
|Production VBA|4|**266,745**|
|Regression / certification VBA|1|**314,489**|
|Demo VBA|2|**220,245**|
|Python repository tooling|3|**27,335**|
|Versioned review / SDI documentation in `docs/`|2|**107,015**|

The project is no longer a small utility module. Its production code is roughly a quarter megabyte of exported VBA, while the test harness is larger than the entire production package. The scale is justified by the platform-sensitive behavior being tested, but it makes static validation and accurate module headers increasingly important.

## 4.2 Public and project-visible surfaces

|Surface|Count|
|-|-:|
|Supported facade enums|2|
|Supported facade callable members|10|
|**Supported caller-facing facade total**|**12**|
|Project-public members listed in `tools/public_api_manifest.txt`|**38**|
|Public regression / certification / characterization entry points|at least 7|
|Mandatory release-certification units|3|

The 38-member manifest deliberately includes internal project-visible helpers and test seams. That is useful for detecting namespace drift, but it is not the same as the supported 12-member caller-facing API.

## 4.3 Release and repository state

|Metric|Current state|
|-|-:|
|Release comparison|`v1.1.0...v1.1.1`|
|Release PR|`#24`|
|Files changed in release PR|22|
|Release PR additions / deletions|11,911 / 2,849|
|Release PR commits|30|
|Hosted workflows|1|
|Open issues|2|
|Stars|3|
|Forks|0|
|Visibility|Public|
|License|MIT|
|Repository size reported by GitHub|approximately 20.3 MB|
|Default branch|`main`|
|`main` versus `v1.1.1`|identical at review time|

## 4.4 Recorded behavioral evidence

The release records the following manual Excel certification:

```text
Excel 16.0 build 20131
Windows (64-bit) NT 10.00
Office / VBA x64, VBA7
2026-08-20 20:13:12

RESULT: PASS | COMPLETE | units=3 failed=0 skipped=0 cleanup=OK
  PASS  RegressionPack
  PASS  SnapshotIdentity
  PASS  TitleBarSdiIdentity
```

This is meaningful evidence, but it is one host, one bitness and one Office build. The generated JSON does not carry a release tag, commit SHA or source hash.

---

# 5. Scoring methodology

A score of 10 requires:

* correct behavior throughout the documented supported domain;
* no silent wrong-target UI operation;
* exact and testable state ownership;
* deterministic recovery from partial WinAPI operations;
* regression tests that distinguish complete from partial execution;
* cleanup verification against the actual entry state;
* behavioral evidence bound to the exact reviewed source;
* automated static and Excel-hosted release gates;
* internally consistent README, source headers, Wiki and demo assets;
* release governance that does not merge unresolved material review findings.

## Weighted scorecard

|Area|Weight|Score|Weighted contribution|
|-|-:|-:|-:|
|Functional correctness|18%|**8.1**|1.458|
|Architecture and modularity|12%|**9.5**|1.140|
|WinAPI and SDI state management|12%|**8.6**|1.032|
|Error handling and diagnostics|10%|**8.7**|0.870|
|Public API and compatibility|8%|**8.8**|0.704|
|Regression and certification|12%|**8.1**|0.972|
|Repository documentation and governance|12%|**8.0**|0.960|
|CI and release engineering|10%|**8.2**|0.820|
|Maintainability and tooling|6%|**8.4**|0.504|
|**Total**|**100%**||**8.460 / 10**|

Rounded overall score:

```text
8.5 / 10
```

## Score interpretation

|Score|Interpretation|
|-:|-|
|9.5-10.0|Exceptional; full contract independently evidenced and release-gated|
|9.0-9.4|Advanced professional component with limited non-material gaps|
|8.0-8.9|Strong implementation with material but targeted hardening remaining|
|7.0-7.9|Good foundation with significant correctness or assurance gaps|
|Below 7.0|Major architectural, correctness or governance deficiencies|

---

# 6. Component scores

|Component|Score|Assessment|
|-|-:|-|
|`M_EXCEL_UI`|**8.8**|Stable facade and backward-compatible targeting; source header still states obsolete scope and version metadata|
|`M_EXCEL_UI_RUNTIME`|**9.0**|Strong non-raising diagnostic redesign; truncation and quiet-scope edge contracts remain imperfect|
|`M_EXCEL_UI_SNAPSHOT`|**9.0**|Title-bar capture/restore is materially safer; Ribbon remains a confirmed wrong-target path|
|`M_EXCEL_UI_TITLEBAR`|**8.6**|Excellent explicit-target and refresh-debt design; numeric-handle registry is vulnerable to handle reuse|
|`M_EXCEL_UI_REGRESSION_TESTS`|**8.1**|Ambitious certification and SDI coverage; certification has several correctness and cleanup-verification defects|
|Demo source and release asset|**6.5**|Attractive foundation, but current feature journeys are absent, presets are recorded as broken, and error diagnostics retain an old defect|
|`tools/check_repo.py`|**8.0**|Valuable hosted gate; parser and API-contract coverage are narrower than its release-safety claims|
|`tools/reformat.py`|**7.8**|Idempotent check/write modes are useful; label rewriting is not token-aware and can alter string data|
|Root documentation|**8.8**|Candid, technically rich and substantially updated; a few internal contradictions and stale source headers remain|
|Public GitHub Wiki|**4.5**|Materially obsolete and unsafe as installation guidance for the current release|
|Repository governance|**7.8**|Strong templates and issue backlog; release merged with two unresolved P2 review threads|

---

# 7. Architectural review

## 7.1 Current dependency model

```text
M_EXCEL_UI
├── M_EXCEL_UI_RUNTIME
├── M_EXCEL_UI_TITLEBAR
└── M_EXCEL_UI_SNAPSHOT
    ├── M_EXCEL_UI_RUNTIME
    └── M_EXCEL_UI_TITLEBAR
```

The graph remains acyclic.

### `M_EXCEL_UI`

Owns:

* supported enums;
* caller-facing procedures and functions;
* tri-state validation;
* target-scope resolution;
* application and window apply orchestration;
* compatibility wrappers.

### `M_EXCEL_UI_RUNTIME`

Owns:

* failure count and list handling;
* Immediate Window logging;
* generic Boolean reads and writes;
* Ribbon reads and writes;
* redraw suppression.

### `M_EXCEL_UI_SNAPSHOT`

Owns:

* all snapshot mutable state;
* retained `Window` references;
* per-element `Known` flags;
* title-bar captured handle, window and label;
* capture and restoration sequencing.

### `M_EXCEL_UI_TITLEBAR`

Owns:

* Win32/Win64 declarations;
* exact owned style bits;
* explicit target-handle operations;
* per-handle frame registry;
* baseline ownership;
* frame refresh debt.

## 7.2 Architectural strengths

### One public facade

Caller code remains concentrated on the established `UI_...` surface. The corrective release does not expose a second competing object model.

### Exact mutable-state ownership

Snapshot state and title-bar state are no longer mixed into the public facade. This is critical because both have lifecycles that survive a single call.

### Explicit handle targeting

The title-bar subsystem now has read and write paths that accept a supplied `hWnd`, instead of forcing every operation through the active `Application.Hwnd`.

### Owned-bit merge

Only the five declared frame bits are changed:

```text
WS_CAPTION
WS_SYSMENU
WS_THICKFRAME
WS_MINIMIZEBOX
WS_MAXIMIZEBOX
```

Unrelated current style bits are preserved.

### Refresh debt

The subsystem now distinguishes:

```text
style already equals requested state
```

from:

```text
style equals requested state but Windows still owes a non-client refresh
```

That distinction is one of the best changes in `v1.1.1`.

## 7.3 Architectural weakness: identity remains split

A top-level frame is represented by two different identity signals:

* numeric `hWnd` for WinAPI access;
* retained Excel `Window` object for workbook-window identity.

The snapshot module retains both, but the title-bar registry itself stores only the numeric handle. Windows can reuse handle values after destruction. A per-handle registry therefore needs either:

* an associated Excel `Window` identity;
* a generation token attached to the native window;
* or another lifecycle signal that invalidates a recycled numeric handle.

Compaction based on `IsWindow` is insufficient once a value has been reused, because the recycled handle is alive.

## 7.4 Architectural verdict

**9.5 / 10**

No architectural rewrite is recommended. The existing decomposition should be retained. The remaining identity problem is local to the frame registry and can be corrected without changing the caller-facing API.

---

# 8. Production-code review

## 8.1 `M_EXCEL_UI` — **8.8 / 10**

### Strengths

* tri-state intent prevents omitted optional Booleans from silently becoming hide operations;
* every visibility enum and target-scope enum is validated defensively;
* application-level operations can continue when an invalid window target is rejected;
* window-level writes are isolated per target window;
* no-op writes are suppressed where reads are available;
* structured and fire-and-forget entry points share one worker;
* legacy parameter order is preserved;
* `UI_ShowExcelUI` remains separate from snapshot restoration.

### No executable regression in v1.1.1

The release diff changes no substantive facade behavior. Public procedure names, parameter order, defaults and enum values remain stable.

### Source-header drift

The module header still says:

```text
Application-level:
  Ribbon, Status Bar, Scroll Bars, Formula Bar

Main-window frame:
  Title Bar

VERSION
  1.1.0
```

That is no longer the project's own measured contract. Under SDI:

* Ribbon acts on the active workbook window;
* title-bar operations act on the active frame unless an explicit target is used internally;
* only Status Bar, Scroll Bars and Formula Bar are genuinely process-wide.

This does not change execution, but it places obsolete architecture text at the top of the public facade, where future maintainers are most likely to trust it.

### Public-wrapper scope asymmetry

`UI_HideExcelUI` and `UI_ShowExcelUI`:

* apply Headings, Workbook Tabs and Gridlines to all current Excel windows by default;
* apply Ribbon and Title Bar only to the active window;
* apply Status Bar, Scroll Bars and Formula Bar process-wide.

The behavior is now documented, but the wrapper names continue to sound more globally uniform than the operations actually are. This should be addressed in the `v1.2.0` scope design rather than through a breaking patch.

---

## 8.2 `M_EXCEL_UI_RUNTIME` — **9.0 / 10**

### Failure accumulation redesign

The new failure path is materially safer:

1. overall success is set to `False` first;
2. `FailureCount` is incremented regardless of list state;
3. entry text is built defensively;
4. list growth is attempted behind a non-raising function;
5. a diagnostic truncation marker is written into an existing slot where possible;
6. logging remains independent of list capture.

This removes the contradiction in which an error-reporting path could itself raise and prevent later best-effort operations.

### Important contract nuance

The documentation says truncation is made visible. That is true when an existing list slot can be overwritten. It cannot be guaranteed when the **first allocation itself** fails, because no array slot exists in which to place the marker. In that case:

* `FailureCount` remains correct;
* the list can remain absent or shorter;
* count/list disagreement is the only durable signal.

The changelog correctly tells callers to treat `FailureCount` as authoritative. The stronger source comments should be narrowed to the behavior actually guaranteed.

### Quiet-update scope remains optimistic

`UI_RuntimeBeginQuietUpdate` runs under `On Error Resume Next` and sets `QuietModeChanged = True` immediately after requesting `Application.ScreenUpdating = False`. It does not verify that the assignment succeeded.

The practical risk is low, but the declared contract — “True only when this scope actually changed the setting” — is stronger than the implementation.

A stronger implementation would:

```vb
Application.ScreenUpdating = False
QuietModeChanged = (Application.ScreenUpdating = False)
```

under the same narrow protected scope.

### No post-write readback

The generic property and Ribbon setters treat a non-raising write as success. Excel's object model usually raises when a write is refused, but a readback would distinguish “call returned” from “requested state achieved.” This remains a P3 hardening item.

---

## 8.3 `M_EXCEL_UI_SNAPSHOT` — **9.0 / 10**

### Title-bar identity correction

The snapshot now retains:

```text
m_SnapshotTitleBarHwnd
m_SnapshotTitleBarWindow
m_SnapshotTitleBarLabel
m_SnapshotTitleBarVisible
m_SnapshotTitleBarKnown
```

Capture resolves the active frame once and reads through that retained handle. Restore refuses to fall back to whichever window is active.

The restore path validates:

* the captured handle is nonzero;
* the handle still names a live native window;
* the retained Excel `Window` still responds, when one was available;
* the explicit-target title-bar write succeeds.

This is a strong correction to the `v1.1.0` P1 defect.

### Remaining identity limitation

The handle and retained `Window` are validated independently; Excel exposes no direct mapping from one to the other. Normally this is sufficient:

* destroyed native frame -> `IsWindow` fails;
* closed Excel window -> retained object probe fails.

The remaining edge is a surviving Excel `Window` whose native frame was recreated while the old handle value was recycled for another frame. Both independent liveness checks can then pass without proving they still refer to the same object. The per-handle registry finding discussed below makes this more than a theoretical design concern.

### Ribbon remains a confirmed wrong-target path

The snapshot still stores only:

```text
RibbonKnown
RibbonVisible
```

It does not retain the `Window` from which the Ribbon value was read. Restore uses the active-window Ribbon mechanisms. The project probe reproduced:

```text
captured from A
restored while B active
B receives the write/no-op
A remains unrestored
result reports success
```

This is the most important remaining production defect.

### Resource lifetime

The additional retained title-bar `Window` reference increases the importance of `UI_ClearExcelUIStateSnapshot`. The repository documents this clearly. Restore intentionally retains the snapshot for replay; it is not a release operation.

---

## 8.4 `M_EXCEL_UI_TITLEBAR` — **8.6 / 10**

### Explicit-target API

The subsystem now supports:

```text
read active handle
read visibility for supplied handle
write visibility for supplied handle
validate supplied handle
```

That is the right abstraction for SDI.

### Per-window registry

Each registry slot records:

```text
Hwnd
OwnedStyleBits
HasBaseline
ComponentHidden
RefreshPending
```

This is substantially more accurate than one process-wide cache.

### Self-healing baseline

When the component does not own a hidden state, it refreshes the baseline from the live style. This allows legitimate changes made by Excel or another add-in to survive the next hide/show cycle.

### Transactional refresh debt

After a successful style write:

* component ownership is updated;
* a failed frame refresh sets `RefreshPending = True`;
* the next operation retries the refresh before deciding that the style is already correct.

This is excellent state-machine design.

### P2: recycled `hWnd` can retrieve stale registry state

Registry lookup scans for numeric equality and returns a slot before compaction:

```vb
For Slot = 1 To m_FrameStateCount
    If m_FrameStates(Slot).Hwnd = TargetHwnd Then
        UI_FrameStateIndexForHwnd = Slot
        Exit Function
    End If
Next Slot
```

Compaction removes entries only when `IsWindow` says a handle is dead. If Windows reuses the numeric value, the recycled handle is alive, so the stale slot survives and is returned.

The new frame can inherit:

* the previous window's owned bits;
* the previous component-hidden flag;
* an outstanding refresh debt.

The most dangerous state is a stale `ComponentHidden = True`. The worker then refuses to refresh its baseline from the new window and can apply the destroyed window's baseline to the replacement.

### Recommended correction

Store an identity corroborator in each slot. The least disruptive design is:

```text
Hwnd + retained Excel Window object
```

The active wrapper can pass `Application.ActiveWindow`; snapshot restore already holds the captured `Window`. A numeric-handle match with a different object must reset or replace the slot.

Where no Excel `Window` object is available, use a weaker explicitly documented path or attach a native generation marker to the window.

### `GetLastError` handling

The code clears and reads the thread last-error through declared WinAPI functions. This is carefully ordered, but VBA's supported bridge is `Err.LastDllError`, captured immediately after the API call. Moving to `Err.LastDllError` would simplify declarations and align with Visual Basic's documented marshaling contract.

### No achieved-state readback

After a style write and frame refresh, the code does not re-read the style. A final readback of the owned mask would provide stronger evidence that the requested state was achieved.

---

## 8.5 Demo source — **6.5 / 10**

The demo remains the weakest code area.

### Missing current journeys

The source and published workbook do not expose:

* `UIWindowTargetScope` selection;
* structured `*_WithResult` diagnostics;
* rendered failure lists;
* snapshot capture -> mutate -> restore -> clear lifecycle;
* multi-window SDI behavior;
* explicit title-bar recovery behavior;
* Ribbon's active-window limitation.

### Broken presets

Open issue `#22` records that the preset controls do not function. The README candidly discloses this and defers the rebuild to `v1.2.0`.

### Demo error diagnostics retain an old production defect

`Demo_GetRuntimeErrorText` executes:

```vb
On Error Resume Next
```

before reading `Err.Number`, `Err.Description`, `Err.Source` and `Erl`.

Any `On Error` statement clears the `Err` object. The function therefore risks producing the same empty `0: ` diagnostics that were previously fixed in the production and test error builders.

The correction is the established project pattern:

```vb
Dim ErrNumber As Long
Dim ErrDescription As String
Dim ErrSource As String
Dim ErrLine As Long

ErrNumber = Err.Number
ErrDescription = Err.Description
ErrSource = Err.Source
ErrLine = Erl

On Error Resume Next
'format captured values
```

### Distribution asset

The latest published demo remains `EXCEL_UI_DEMO_v1.1.0.xlsm`. It is not a `v1.1.1` artifact, does not demonstrate the corrected title-bar behavior, and has no published checksum in the current release documentation.

---

# 9. Public API and compatibility review

## 9.1 Supported facade

The supported caller-facing API remains:

### Enums

```text
UIVisibility
UIWindowTargetScope
```

### Callables

```text
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

No supported name, parameter position, type, default or enum value was changed in `v1.1.1`.

## 9.2 Source compatibility versus package compatibility

The README now explains this well:

* caller code does not require modification;
* all four production modules must nevertheless be replaced together;
* mixing internal module versions is not a supported upgrade.

## 9.3 API manifest is name-level only

`tools/public_api_manifest.txt` records:

```text
module
kind
member name
```

It does **not** record:

* parameter names;
* parameter order;
* `ByVal` versus `ByRef`;
* parameter types;
* optional defaults;
* function return types;
* enum members or values;
* conditional-compilation signature differences.

A breaking change such as moving `FailureCount`, changing a default target scope, or changing an enum numeric value can pass the current API gate as long as the member name remains.

### Recommended manifest split

Use two files:

```text
tools/supported_api_manifest.txt
  full normalized declarations for M_EXCEL_UI
  enum members and values

tools/project_public_surface_manifest.txt
  name-level inventory of internal project-visible seams
```

The first is the SemVer gate. The second is a namespace and dead-surface gate.

## 9.4 Scope model needs a `v1.2.0` decision

Current targeting semantics are asymmetric:

|Element|Effective scope|
|-|-|
|Status Bar / Scroll Bars / Formula Bar|process|
|Headings / Tabs / Gridlines|supplied `TargetScope`|
|Ribbon|active window only|
|Title Bar|active window through public API; explicit target internally|

A future release should decide whether Ribbon and Title Bar:

* remain explicitly active-window only;
* follow `TargetScope`;
* receive their own target policy;
* or are restored through a separate activation policy.

The API should not silently activate workbooks unless the caller opts into that observable behavior.

---

# 10. Regression and release-certification review

## 10.1 Strengths

The test module now contains distinct entry points for:

* core iteration;
* title-bar-only iteration;
* full legacy pack;
* snapshot identity;
* SDI title-bar identity;
* Ribbon SDI characterization;
* release certification.

The dedicated certification runner improves the assurance model by:

* rejecting a pre-existing snapshot rather than silently skipping cases;
* running three mandatory units separately;
* continuing after one unit fails;
* counting failed units;
* counting skipped mandatory work;
* checking selected cleanup conditions;
* emitting JSON and text evidence;
* raising when the verdict is not PASS/COMPLETE.

This is a major improvement over `Test_EXCEL_UI_RunAll` as a release signal.

## 10.2 P2: `ScreenUpdating` cleanup is compared with `True`, not the baseline

The runner captures:

```vb
OldScreenUpdating = Application.ScreenUpdating
```

but cleanup uses:

```vb
If Not Application.ScreenUpdating Then
    CleanupOK = False
    CleanupDetail = "ScreenUpdating was left suppressed"
End If
```

A caller that deliberately enters certification with `ScreenUpdating = False` can have that state restored perfectly and still receive a failed cleanup verdict.

The correct assertion is:

```vb
If Application.ScreenUpdating <> OldScreenUpdating Then
```

The release's recorded certification began with the ordinary `True` baseline, so the defect did not affect the published PASS.

## 10.3 P2: certification error handler destroys the error it intends to re-raise

The handler is structurally:

```vb
Err_Handler:
    m_CertActive = False
    TST_Log PROC, "FAIL", Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
```

`TST_Log` begins with `On Error Resume Next`. VBA clears the global `Err` properties when an `On Error` statement executes. The original error is therefore no longer available after the logging call.

Consequences:

* the original number, source and description are not rethrown;
* a precondition failure or failed verdict can be replaced by a different runtime error or lose its diagnostic identity;
* callers cannot reliably distinguish certification refusal from certification failure.

The correct pattern is already used elsewhere in the repository:

```vb
FailNumber = Err.Number
FailSource = Err.Source
FailDescription = Err.Description

m_CertActive = False
TST_Log PROC, "FAIL", FailDescription
Err.Raise FailNumber, FailSource, FailDescription
```

Add a regression case that intentionally triggers:

* the pre-existing-snapshot precondition; and
* a synthetic failed certification verdict;

then asserts the exact number, source and description received by the caller.

## 10.4 P2: cleanup can miss a leaked Excel window

The runner records:

```text
Workbooks.Count
```

but not:

```text
Application.Windows.Count
```

`ThisWorkbook.NewWindow` adds an Excel window without adding a workbook. A failed cleanup in the snapshot-identity unit can therefore leave a stray window while:

* workbook count remains unchanged;
* the anchor window remains usable;
* no snapshot remains;
* ScreenUpdating is restored.

The certification can report `cleanup=OK` despite leaked window state.

Record and compare both counts, and preferably retain entry-window identities rather than relying only on counts.

## 10.5 P2: complete managed UI state is not verified after cleanup

The underlying regression pack performs best-effort state restoration under suppressed error handling. Certification later checks only:

* snapshot absence;
* workbook count;
* `ScreenUpdating`;
* anchor-window liveness.

It does not verify the entry values of:

* Ribbon;
* Status Bar;
* Scroll Bars;
* Formula Bar;
* Title Bar;
* Headings;
* Workbook Tabs;
* Gridlines.

A failed cleanup write can therefore leave Excel altered while certification still reports cleanup OK.

A real release certificate should capture a full managed-state baseline before the first unit and compare the achieved state after all cleanup. Where Ribbon identity cannot be restored safely, the certificate should report the limitation explicitly rather than silently omitting the comparison.

## 10.6 Certification record arrays are themselves fail-soft without integrity checks

`TST_CertRecordUnit` and `TST_CertRecordSkip` increment counters and resize parallel arrays under `On Error Resume Next`. A memory or array failure can leave:

* a count incremented;
* only some arrays resized;
* default Boolean values interpreted as failed units;
* evidence generation operating on inconsistent arrays.

This is unlikely in ordinary runs, but the production failure accumulator was redesigned specifically to avoid this class of failure. The certification evidence path should apply the same discipline.

## 10.7 Evidence is host-bound, not source-bound

The JSON records:

```text
component
schema
timestampLocal
Excel version and build
operating system
bitness
VBA generation
unit counts and results
skip detail
cleanup result
```

It does not record:

```text
release tag
commit SHA
tree SHA
hash of imported .bas files
test-module hash
```

The files are written to `%TEMP%` on a best-effort basis and are not versioned in the repository.

The changelog's manual statement that certification occurred after the last production-source change is useful, but it is not machine-verifiable provenance.

## 10.8 Environment coverage

The current certificate covers one environment:

```text
Excel 16.0 build 20131
Windows x64
VBA7
```

The package claims 32-bit and 64-bit Office support. The 32-bit path is source-reviewed and conditionally declared, but not evidenced by the current release certificate.

---

# 11. Dedicated repository-quality assessment

## Repository-quality score: **8.0 / 10**

The repository is unusually strong for a VBA component in structure, policy and transparency. The score is reduced by the stale public Wiki, outdated demo, review-gate failure and source-evidence gap.

## 11.1 Repository structure — **9.4 / 10**

Strengths:

* clear `src/`, `test/`, `demo/`, `docs/`, `tools/` separation;
* no production workbook binary committed;
* explicit CRLF policy for exported VBA;
* source-focused release archive rules;
* cohesive four-module production package;
* versioned independent review and SDI behavior study;
* issue and pull-request templates tailored to state ownership and recovery.

Weaknesses:

* large test and title-bar modules are approaching a reviewability threshold;
* source header metadata is not checked by the static gate;
* the Wiki exists outside the tagged source tree and has drifted badly.

## 11.2 Root README and versioned docs — **8.8 / 10**

The root README is now technically candid and professionally structured. It correctly states:

* the four-module requirement;
* source compatibility versus package compatibility;
* actual active-window Ribbon and title-bar behavior;
* the known Ribbon identity limitation;
* the current demo limitation;
* the `v1.1.1` corrective scope.

Remaining inconsistencies:

* the target-scope section later says Ribbon and Title Bar retain “application/main-window scope,” contradicting the active-window explanation above it;
* “snapshot capture/restore retains all-managed-windows semantics” is too broad when Ribbon is not restored by captured identity;
* the public facade and test-module headers still report version `1.1.0` and obsolete scope descriptions.

## 11.3 Public GitHub Wiki — **4.5 / 10**

The public Wiki home page was last edited on 2026-04-21 and describes an architecture that predates both `v1.1.0` and `v1.1.1`.

It currently tells users, among other things, that:

```text
/src/M_EXCEL_UI.bas
/demo/EXCEL_UI_DEMO.xlsm
```

represent the typical repository structure.

It instructs “Core only” users to import only:

```text
M_EXCEL_UI.bas
```

That installation is invalid for the current four-module package and will not compile.

It also:

* calls Ribbon application-level;
* omits the modular runtime, snapshot and title-bar modules;
* omits `UIWindowTargetScope`;
* omits structured snapshot result APIs;
* lists only the old test runners;
* describes old window-order snapshot limitations;
* treats the demo workbook as version-controlled.

This is not harmless historical prose. GitHub exposes the Wiki as a first-class repository navigation item, and its quick start conflicts directly with `INSTALLATION.md`.

### Recommended decision

Choose one:

1. **Disable the Wiki** and keep all authoritative documentation versioned in the repository; or
2. generate/synchronize the Wiki from tagged Markdown during release preparation.

Maintaining two independent documentation systems is already producing unsafe drift.

## 11.4 Issue and backlog quality — **9.0 / 10**

The repository now openly tracks:

* P1 Ribbon snapshot identity (`#23`);
* P2 demo modernization and broken presets (`#22`).

Both issues have clear acceptance criteria and are assigned to a future milestone. This is good governance and much better than silently burying known limitations in prose.

The weakness is that the release is marked Stable while a P1 silent wrong-target path remains open. The README caveat is clear, but release-status wording should be interpreted as stable **within the documented limitation**, not as complete multi-window snapshot correctness.

## 11.5 Pull-request and review governance — **7.1 / 10**

Release PR `#24` was comprehensive and included a semantic-versioning statement, certification result and detailed change description.

However, two P2 review threads were posted shortly before merge and remained unresolved:

* recycled `hWnd` registry state;
* `ScreenUpdating` cleanup baseline.

The PR was merged approximately seconds after those comments appeared. The review was recorded as a comment rather than a blocking change request, so it did not gate merge.

### Recommended controls

* require all review conversations resolved before merge;
* require at least one approving review after the latest pushed commit;
* prevent merge while a review bot is still posting findings;
* require the static workflow;
* require a manually uploaded or automated Excel certification artifact for release branches;
* enable automatic release-branch deletion after merge.

## 11.6 CI and automation — **8.2 / 10**

### Strengths

* hosted read-only workflow;
* exact-tag tree was covered by a successful PR-head static run because both share the same tree;
* module names, options, CRLF, ASCII, tabs, trailing whitespace, banner widths, directive balance, selected labels, PtrSafe declarations, duplicate names, public-surface names, release-state markers, binary exclusions, Markdown links and formatter idempotence are checked;
* local and CI gates use the same script.

### Weaknesses

* no automated Excel runtime workflow;
* no tag trigger;
* actions are referenced by major tags rather than immutable commit SHAs;
* exact supported API signatures and enum values are not checked;
* static parsing has false-negative paths described below;
* source version metadata is not checked;
* release evidence is not uploaded by workflow.

## 11.7 Static checker limitations — **8.0 / 10**

`tools/check_repo.py` is a valuable addition, but several checks are weaker than their labels imply.

### Jump targets are module-scoped, not procedure-scoped

Labels are collected for the entire module. A procedure missing its own `Safe_Exit:` can pass because another procedure defines the same label elsewhere in the module.

### Conditional-compilation tracking is not nested

`check_ptrsafe` uses one Boolean. An inner `#Else` for `#If Win64` turns off the outer VBA7 state, so several VBA7 Win32 declarations and declarations after the inner branch are not actually checked.

### Procedure end kinds are not matched

The procedure stack pops on any `End Sub`, `End Function` or `End Property`; it does not confirm the closing kind matches the opener.

### WinAPI aliases are not checked

The checker would not detect the historical “prefixed Declare without Alias” defect that existed in the test harness.

### Binary coverage is narrower than repository policy

The checker directly forbids only a limited set of workbook extensions. `.gitignore` and `.gitattributes` cover more formats, but `git add -f` can bypass ignore rules.

### No runner / case inventory

The checker does not validate that:

* every public certification runner is documented;
* every intended case is registered;
* the certification unit registry contains the required units;
* public test-module headers match the executable surface.

## 11.8 Release assets and supply chain — **7.8 / 10**

Strengths:

* source tag and `main` are identical;
* merge commit is GitHub-verified;
* no workbook binary is stored in the source tree;
* macro-enabled artifacts are explicitly treated as executable content;
* security policy and release-source guidance are strong.

Weaknesses:

* no `v1.1.1` demo asset;
* latest demo is known incomplete and partially broken;
* checksum publication remains optional;
* certification evidence is not attached or source-bound;
* GitHub Actions are not pinned to immutable SHAs.

## 11.9 Discoverability and adoption — **7.5 / 10**

The README, images, topics and focused problem statement are good. External adoption remains early:

```text
3 stars
0 forks
```

That is not a code-quality defect, but it means there is little independent field evidence across Excel builds, add-in combinations and enterprise policies. A reliable current demo and Discussions/Q&A channel would improve adoption and issue discovery.

---

# 12. Security and platform assessment

No high-severity security vulnerability was identified in the production source.

Positive controls include:

* no external dependency or installer;
* no arbitrary shell execution;
* no network access;
* fixed Excel 4 macro strings rather than caller-supplied macro text;
* exact style-bit ownership;
* handle validation before writes;
* read-only hosted workflow permissions;
* strong security documentation;
* explicit statement that hidden UI is not an access-control boundary.

Platform-sensitive risks are primarily integrity and availability risks:

* wrong-window Ribbon changes;
* stale frame state after handle reuse;
* persistent UI alteration after interrupted execution;
* other add-ins writing the same Ribbon or frame state;
* downloaded macro-enabled release assets.

Recommended supply-chain improvements:

* publish SHA-256 for every `.xlsm` asset;
* attach certification JSON and text to the release;
* include source-tree fingerprint in certification;
* pin GitHub Actions by full commit SHA;
* consider signing the VBA project or publishing a signing/deployment guide for enterprise consumers.

---

# 13. Findings summary

|ID|Severity|Area|Finding|
|-|-|-|-|
|ICR-UI-111-P1-01|**P1**|Snapshot correctness|Ribbon snapshot restoration applies the captured value to the active window, not the captured window, and reports success|
|ICR-UI-111-P2-01|**P2**|Title-bar state|Per-`hWnd` registry can reuse stale state when Windows recycles a numeric handle|
|ICR-UI-111-P2-02|**P2**|Certification|Cleanup compares `ScreenUpdating` with `True`, not with the captured baseline|
|ICR-UI-111-P2-03|**P2**|Certification error handling|Certification handler calls a logging procedure containing `On Error` before preserving `Err`, then attempts to re-raise cleared error properties|
|ICR-UI-111-P2-04|**P2**|Certification cleanup|Cleanup checks workbook count but not Excel window count and does not verify full managed UI-state restoration|
|ICR-UI-111-P2-05|**P2**|API governance|Public API manifest detects names only, not signatures, defaults or enum values|
|ICR-UI-111-P2-06|**P2**|Documentation|Public Wiki gives invalid single-module installation instructions and obsolete scope/API information|
|ICR-UI-111-P2-07|**P2**|Demo|Demo remains incomplete and broken; its runtime-error builder still clears `Err` before reading it|
|ICR-UI-111-P2-08|**P2**|Release evidence|Behavioral certificate is manual, one-host and not bound to tag, commit or source hashes|
|ICR-UI-111-P2-09|**P2**|Review governance|Release was merged with two unresolved P2 review threads that describe defects present in the tag|
|ICR-UI-111-P3-01|**P3**|Static checker|Jump labels are checked module-wide, conditional-compilation state is not nested, aliases and end-kind pairing are not checked|
|ICR-UI-111-P3-02|**P3**|Formatter|Regex label rewriting is not string/comment aware and can alter literal text|
|ICR-UI-111-P3-03|**P3**|Source documentation|Facade and test-module headers retain `1.1.0` metadata and obsolete scope/public-surface descriptions|
|ICR-UI-111-P3-04|**P3**|WinAPI diagnostics|Direct declared `GetLastError` remains instead of `Err.LastDllError`|
|ICR-UI-111-P3-05|**P3**|Achieved-state verification|Ribbon, object-model and frame writes have no final readback contract|
|ICR-UI-111-P3-06|**P3**|Quiet-update scope|`QuietModeChanged` can be set without verifying `ScreenUpdating` actually changed|
|ICR-UI-111-P3-07|**P3**|Diagnostics|First failure-list allocation failure cannot always carry an explicit truncation marker|
|ICR-UI-111-P3-08|**P3**|CI supply chain|Workflow does not trigger on tags and actions are not pinned by immutable SHA|
|ICR-UI-111-P3-09|**P3**|Platform assurance|Only one x64 Office build is behaviorally certified|
|ICR-UI-111-P3-10|**P3**|Maintainability|Test and title-bar modules are becoming difficult to review as single files|

---

# 14. Detailed material findings

## ICR-UI-111-P1-01 — Ribbon snapshot restoration is not window-identity-safe

### Severity

**P1 — silent wrong-target state restoration**

### Evidence

The repository's own probe records:

```text
captured from  A  (Ribbon visible)
restored to    B  (active at restore time)
A              still hidden, never restored
result         success, no failure reported
```

The snapshot stores one Ribbon Boolean and no captured `Window` identity.

### Impact

A multi-window caller can believe the baseline was restored when:

* the captured window remains altered;
* another window was changed or treated as a no-op;
* no structured failure identifies the miss.

This violates the core safety promise more seriously than an ordinary best-effort failure because the operation reports success against the wrong target.

### Recommended immediate patch

A minimal `v1.1.2` safety correction does **not** need to activate another window:

1. capture and retain the `Window` from which Ribbon state was read;
2. at restore, compare it with `Application.ActiveWindow` by retained object identity;
3. if it is not active, return an ordered `Ribbon` failure and perform no Ribbon write;
4. if the captured window is closed, return a specific failure;
5. keep the snapshot for replay.

That removes silent misapplication without introducing activation side effects.

### Recommended `v1.2.0` capability

Add an explicit optional policy:

```text
DoNotActivateCapturedWindow  default, safe refusal
ActivateCapturedWindow       opt-in, write, restore prior focus
```

Document event firing and reentrancy implications.

---

## ICR-UI-111-P2-01 — Recycled native handles can retrieve stale frame state

### Severity

**P2 — rare but real native identity collision**

### Evidence

The registry returns the first numerically equal `Hwnd` before compaction. Compaction relies on `IsWindow` and therefore cannot distinguish a recycled value from the original frame.

### Impact

A new workbook frame can inherit:

```text
OwnedStyleBits
HasBaseline
ComponentHidden
RefreshPending
```

from a destroyed window.

### Recommended correction

Extend each registry slot with an identity corroborator and validate it on lookup. Where possible, store the Excel `Window` object alongside the handle. On mismatch:

* clear the stale slot;
* initialize a fresh baseline from the live frame;
* never carry `ComponentHidden` or refresh debt across generations.

Add a deterministic injection seam that simulates a recycled handle/identity mismatch, because real handle reuse is not reliably producible on demand.

---

## ICR-UI-111-P2-02 — Certification rejects a valid `ScreenUpdating = False` baseline

### Severity

**P2 — false release-gate failure**

### Current condition

```vb
If Not Application.ScreenUpdating Then
```

### Required condition

```vb
If Application.ScreenUpdating <> OldScreenUpdating Then
```

### Test

Run certification from both entry states:

```text
ScreenUpdating = True
ScreenUpdating = False
```

Both must restore the exact entry value and produce the same PASS/COMPLETE verdict when all other conditions pass.

---

## ICR-UI-111-P2-03 — Certification clears the original error before rethrowing it

### Severity

**P2 — failed certification can lose its failure identity**

### Root cause

`TST_Log` uses `On Error Resume Next`, which clears `Err`. The handler reads `Err` again after logging.

### Correction

Capture all fields before any call:

```vb
FailNumber = Err.Number
FailSource = Err.Source
FailDescription = Err.Description
FailLine = Erl
```

Then log the captured values and raise from the captured values.

### Regression requirement

Assert exact propagation for:

* precondition rejection;
* failed mandatory unit;
* incomplete verdict;
* cleanup failure.

---

## ICR-UI-111-P2-04 — Certification cleanup is incomplete

### Severity

**P2 — PASS can coexist with leaked window or UI state**

### Missing checks

* entry versus exit `Application.Windows.Count`;
* identity of every pre-existing window;
* full application property baseline;
* full per-window property baseline;
* title-bar achieved state;
* Ribbon state where safely addressable;
* cleanup failures emitted by `TST_RestoreState`.

### Correction

Create a certification-specific full-state snapshot independent of the component's own snapshot, keyed by retained test `Window` objects. Cleanup must return a structured result rather than suppressing every error.

---

## ICR-UI-111-P2-05 — Public API manifest does not protect the actual compatibility contract

### Severity

**P2 — future breaking changes can pass the compatibility gate**

### Current guarantee

```text
member name exists
```

### Required guarantee

```text
full normalized declaration unchanged
optional defaults unchanged
enum members and values unchanged
```

### Correction

Parse multi-line declarations into a normalized signature and include enum members/values. Add unit tests for the checker itself.

---

## ICR-UI-111-P2-06 — Public Wiki is unsafe for current users

### Severity

**P2 — installation and behavior documentation defect**

### Current unsafe instruction

```text
Import M_EXCEL_UI.bas only
```

### Actual requirement

```text
M_EXCEL_UI_RUNTIME
M_EXCEL_UI_TITLEBAR
M_EXCEL_UI_SNAPSHOT
M_EXCEL_UI
```

### Correction

Disable the Wiki immediately or replace its home page with a short pointer to versioned `README.md` and `INSTALLATION.md`. Re-enable only if synchronization is automated.

---

## ICR-UI-111-P2-07 — Demo is not a trustworthy current adoption surface

### Severity

**P2 — user-facing integration and diagnostics defect**

### Required work

* fix preset controls;
* add target-scope selector;
* render structured failures;
* add snapshot-clear action;
* demonstrate multi-window behavior;
* disclose and demonstrate Ribbon active-window semantics;
* fix `Demo_GetRuntimeErrorText`;
* audit every `OnAction` string;
* build from exact tag;
* publish checksum.

---

## ICR-UI-111-P2-08 — Certification evidence is not bound to source

### Severity

**P2 — release provenance gap**

### Correction options

Preferred:

* automate import, compile and certification from the checked-out tag on a controlled Windows/Excel runner;
* hash every imported `.bas` file;
* include commit, tree and hashes in JSON;
* upload evidence as workflow and release assets.

Minimum manual alternative:

* generate a source-fingerprint file before import;
* include it in the workbook or injected certification bridge;
* attach JSON, text and fingerprints to the release.

---

## ICR-UI-111-P2-09 — Review findings were not allowed to gate release

### Severity

**P2 — release-governance defect**

### Impact

Both unresolved comments describe defects that remain in the tag. Static checks cannot find either because they are behavioral and semantic.

### Correction

Make review-thread resolution a release requirement. A release PR should not be merged until the review provider has completed and the latest commit has been reviewed.

---

# 15. Static tooling review

## 15.1 `tools/check_repo.py`

The checker is worth retaining and extending. Recommended next checks:

```text
per-procedure label resolution
matching procedure begin/end kind
nested conditional-compilation stack
Declare Alias checks
complete Office-binary extension set
module VERSION consistency
supported full API signatures
enum values
required workflow presence
test runner and certification-unit inventory
README / Wiki authoritative-source policy
```

Add a small Python test suite with deliberately malformed fixture modules. At present the checker itself has no committed regression harness.

## 15.2 `tools/reformat.py`

The new modes are good:

```text
--check
--write
legacy explicit source/destination mode
```

The remaining safety defect is in `rename_labels`. It applies regex substitutions to whole lines, including:

* comments;
* string literals;
* generated command text.

For example, a VBA literal containing:

```vb
"GoTo Fail"
```

can be rewritten even though it is data, not a control-flow token.

The tool's claim that it “never touches an executable token” is therefore too strong.

### Correction

Implement a minimal VBA line lexer that distinguishes:

* code;
* quoted strings with doubled quotes;
* apostrophe comments;
* `Rem` comments.

Apply label-reference replacement only in the code portion.

Add idempotence and preservation fixtures for strings and comments.

---

# 16. Prioritized remediation plan

## Release Gate A — Remove silent Ribbon misapplication

1. Capture Ribbon owning `Window`.
2. In a patch release, refuse restore when it is not active.
3. Return ordered `Ribbon` failure.
4. Add closed-window and wrong-active-window regressions.
5. In `v1.2.0`, add opt-in activation policy.

## Release Gate B — Harden title-bar frame identity

1. Add registry identity corroborator.
2. Reset slot on handle-generation mismatch.
3. Add recycled-handle simulation seam.
4. Verify no stale baseline, hidden flag or refresh debt crosses generations.

## Release Gate C — Correct certification

1. Compare `ScreenUpdating` with captured baseline.
2. Preserve `Err` before logging.
3. Record and verify window count.
4. Verify complete managed UI-state cleanup.
5. make unit/skip record accumulation internally consistent under failure.
6. add certification self-tests for pass, failure, skip and cleanup rejection.

## Release Gate D — Bind evidence to source

1. Add commit/tree/source hashes to JSON.
2. Automate Excel run on controlled runner.
3. upload evidence artifact.
4. trigger release gate on tags.
5. certify x86 and at least one additional Office build/channel.

## Repository Gate E — Restore one authoritative documentation system

1. Disable or regenerate Wiki.
2. update facade and test module headers.
3. remove README scope contradiction.
4. ensure documentation map identifies one authoritative source per topic.

## Adoption Gate F — Rebuild demo

1. fix presets;
2. add all current journeys;
3. audit control links;
4. build from tag;
5. publish SHA-256;
6. add a smoke-test checklist and evidence.

## Static Gate G — Protect compatibility and tooling

1. full API-signature manifest;
2. enum-value manifest;
3. checker tests;
4. token-aware formatter;
5. immutable action pins;
6. required review-thread resolution.

---

# 17. Release-readiness assessment

## Suitable now

`v1.1.1` is suitable for:

* single-workbook application-style workbooks;
* controlled demonstrations using source examples rather than the old binary demo;
* multi-window use where Ribbon snapshot restoration is not relied upon;
* projects that keep `UI_ShowExcelUI` readily accessible;
* projects that clear snapshots during shutdown;
* Windows x64 Excel environments similar to the certified host;
* source-pinned internal deployment with local compile and certification.

## Conditional use

Use with explicit constraints for:

* multiple open workbook windows;
* other add-ins that modify title-bar bits;
* environments that begin with `ScreenUpdating = False`;
* long-running sessions that create and destroy many workbook windows;
* governed enterprise deployment requiring reproducible release evidence.

## Not yet suitable for an unconditional claim

Do not claim:

```text
complete identity-safe restoration of every managed UI element across SDI windows
```

until the Ribbon defect is closed.

Do not treat the current release-certification JSON as proof of exact-source execution until source identity is included.

Do not direct users to the current Wiki for installation.

---

# 18. Final verdict

`v1.1.1` is a materially better release than `v1.1.0`.

Its strongest achievements are:

* correction of the title-bar wrong-active-window defect;
* explicit-target WinAPI design;
* per-window frame state;
* self-healing baselines;
* retryable frame-refresh debt;
* non-raising production failure accumulation;
* dedicated SDI testing;
* candid Ribbon characterization;
* a real static repository gate;
* stronger release and contribution documentation.

The component's architecture is now mature. The remaining weaknesses are concentrated and actionable rather than systemic.

The decisive limitation is that the snapshot is only partly identity-safe: title-bar and ordinary window properties are now protected, but Ribbon remains a measured silent wrong-target operation. The new release certificate improves assurance dramatically, but several defects prevent it from being accepted as an authoritative gate without correction.

> **Final score: 8.5 / 10**  
> **Classification: strong professional VBA UI component with a mature architecture, one known P1 multi-window snapshot defect, several targeted certification defects, and repository documentation/release-governance cleanup still required.**

---

# Appendix A — Recommended GitHub issues

1. **P1 — Refuse Ribbon snapshot restore when the captured window is not active**
2. **P2 — Bind title-bar registry slots to window generation / Excel Window identity**
3. **P2 — Compare certification ScreenUpdating cleanup with the entry baseline**
4. **P2 — Preserve Err before certification logging and rethrow**
5. **P2 — Verify Excel window count and complete UI state during certification cleanup**
6. **P2 — Version certification evidence with tag, commit, tree and source hashes**
7. **P2 — Replace name-only supported API manifest with full signatures and enum values**
8. **P2 — Disable or regenerate the stale GitHub Wiki**
9. **P2 — Rebuild the demo and fix runtime diagnostics**
10. **P3 — Add regression fixtures for tools/check_repo.py**
11. **P3 — Make reformat.py token-aware**
12. **P3 — Move WinAPI last-error reads to Err.LastDllError**
13. **P3 — Add achieved-state readback to property, Ribbon and title-bar writes**
14. **P3 — Pin GitHub Actions by commit SHA and add tag triggers**
15. **P3 — Add x86 and second-build certification evidence**

---

# Appendix B — Recommended score-to-10 roadmap

## To reach 9.0

* close Ribbon P1 or at minimum refuse wrong-target restore;
* fix all certification P2 findings;
* close `hWnd` reuse risk;
* disable or update Wiki;
* rebuild current demo.

## To reach 9.5

* automate Excel certification from exact tag;
* source-bind evidence;
* certify x86 and multiple Office environments;
* protect full API signatures and enum values;
* require resolved review threads and passing runtime evidence.

## To reach 10.0

* complete opt-in Ribbon window activation with focus/event policy;
* prove all supported scope semantics across the environment matrix;
* publish signed/checksummed release artifacts;
* provide deterministic current demo and behavioral smoke evidence;
* maintain zero known silent wrong-target paths.

---

# Appendix C — Evidence confidence

|Conclusion|Confidence|
|-|-|
|Exact tag, commit and tree reviewed|High|
|`main` and `v1.1.1` identical at review time|High|
|Four-module architecture and public facade compatibility|High|
|Title-bar explicit-target and refresh-debt design|High|
|Ribbon wrong-target defect|High; measured and tracked as P1 by repository|
|`hWnd` reuse registry defect|High by control-flow inspection|
|ScreenUpdating certification defect|High by direct source inspection|
|Certification Err-clobber defect|High by source and VBA Err contract|
|Certification cleanup incompleteness|High by direct source inspection|
|Static workflow passed on exact reviewed tree|High; PR head and tag share tree|
|Manual Excel certification result|Medium-high as committed evidence; not independently executed|
|32-bit behavior|Medium; source-reviewed, not release-certified|
|Demo preset failure|High as repository-recorded open issue; binary not independently executed|
|Wiki staleness|High for public Wiki home page reviewed on 2026-08-20|
|Branch-protection enforcement|Not independently verified|
|Release-asset contents|Not independently inspected|

---

# Appendix D — Source files that should be changed first

```text
src/M_EXCEL_UI_SNAPSHOT.bas
    capture Ribbon Window identity
    refuse unsafe wrong-window restore

src/M_EXCEL_UI_TITLEBAR.bas
    corroborate registry handle identity

test/M_EXCEL_UI_REGRESSION_TESTS.bas
    ScreenUpdating baseline
    preserve Err before logging
    verify Application.Windows.Count
    verify complete state cleanup
    source-bound evidence fields

tools/check_repo.py
    full signature / enum gate
    procedure-scoped labels
    nested directives
    Alias checks
    metadata and runner inventory

tools/reformat.py
    token-aware label rewriting

demo/M_EXCEL_UI_DEMO.bas
    fix Err capture
    expose current features
    fix presets

README.md / Wiki
    eliminate scope contradictions
    disable or synchronize Wiki
```

