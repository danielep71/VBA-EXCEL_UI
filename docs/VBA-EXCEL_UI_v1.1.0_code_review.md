# Code Review — VBA-EXCEL_UI `release/v1.1.0`

**Repository:** [danielep71/VBA-EXCEL_UI](https://github.com/danielep71/VBA-EXCEL_UI)
**Branch reviewed:** `release/v1.1.0` (23 commits ahead of `main`)
**Base:** `main` @ `v1.0.1`
**Review date:** 2026-08-18

**Scope of review:** all four `src/` production modules, the 5,249-line regression harness, both demo modules, and the repository documentation set (README, INSTALLATION, CONTRIBUTING, SECURITY, CHANGELOG, `.gitignore`, `.gitattributes`, issue/PR templates).

---

## Contents

- [Overall assessment](#overall-assessment)
- [Scoring](#scoring)
- [What the branch does well](#what-the-branch-does-well)
- [P1 — Blockers](#p1--blockers)
- [P2 — Fix before tagging](#p2--fix-before-tagging)
- [P3 — Follow-ups](#p3--follow-ups)
- [Suggested go/no-go](#suggested-gono-go)
- [Appendix A — Files changed](#appendix-a--files-changed)
- [Appendix B — Automated checks run](#appendix-b--automated-checks-run)

---

## Overall assessment

**7.1 / 10 — strong engineering, two real blockers.**

The modular decomposition is genuinely well done. The dependency direction is clean (`RUNTIME` ← `SNAPSHOT` → `TITLEBAR`, facade on top, no cycles), there are no cross-module name collisions in VBA's flat global namespace, no undefined references, no dead public surface, and no `MsgBox` anywhere in `src/`. Extracting `UI_InternalMergeTitleBarStyleBits` as a pure function so the merge policy can be tested without Windows is a good instinct.

However, the title-bar refactor introduced a silent regression in exactly the path the README sells hardest, and the snapshot capture has a fail-soft hole that can destroy a snapshot the caller believes was taken.

---

## Scoring

| Dimension | Weight | Score | Notes |
|---|---:|---:|---|
| Architecture & modularity | 10% | 9.0 | Clean four-module split; no cycles; no namespace collisions |
| Public API design & backward compat | 10% | 9.0 | `TargetScope` appended last — positional callers unaffected |
| Error handling & fail-soft correctness | 20% | 6.0 | Several documented "does not raise" contracts are not enforced |
| WinAPI / platform correctness | 15% | 5.5 | Recovery regression + unreliable `GetLastError` strategy |
| Snapshot state & lifecycle | 15% | 6.5 | Identity-safe restore is right; capture path and COM lifetime are not |
| Test coverage & harness quality | 12% | 7.5 | Broad and disciplined, but misses the stateful title-bar path |
| Documentation | 8% | 8.5 | Excellent depth; CHANGELOG missing, two files contradict the branch |
| Code style & consistency | 5% | 9.0 | House style applied consistently; headers complete |
| Release readiness & repo hygiene | 5% | 6.0 | Version metadata drift, stale docs, no CI |
| **Weighted total** | **100%** | **7.13** | |

---

## What the branch does well

Worth stating explicitly, because the findings below are all deltas against an already-solid baseline.

- **The decomposition is real, not cosmetic.** The 3,439-line monolith became four modules with a defensible ownership boundary each. `M_EXCEL_UI_TITLEBAR` deliberately carries no project-module dependency (it duplicates `BuildErrorText` rather than depending on `RUNTIME`), and the README documents that choice accurately.
- **Identity-safe window restoration is the correct design.** `UI_SnapshotTryResolveWindow` retains the exact captured `Window` object and validates it with a non-mutating probe, rather than re-enumerating and comparing with `Is` — the header comment correctly explains that COM may hand back a different wrapper for the same live window. This is a genuine improvement over v1.0's index-based restore.
- **The merge policy is isolated and testable.** Pulling `UI_InternalMergeTitleBarStyleBits` out as a pure function, `Public` only for same-project test access under `Option Private Module`, is exactly the right seam.
- **Backward compatibility was handled carefully.** `TargetScope` is appended after `FailureCount`/`FailureList` in `UI_SetExcelUI_WithResult`, so existing positional callers are unaffected. The public `UI_...` surface is preserved.
- **Verified clean by static cross-reference:** no `UI_*` identifier is called without being defined; no public `src/` procedure is unreferenced outside its own module; the only duplicate procedure names are `#If`/`#Else` bitness variants (expected).

---

## P1 — Blockers

### P1-1. `UI_ShowExcelUI` silently fails to restore the title bar after a VBA project reset

**File:** `src/M_EXCEL_UI_TITLEBAR.bas`
**Lines:** 454–485 (capture block and short-circuit)
**Introduced by:** commit `83bc9d4` — *feat(titlebar): preserve unrelated window style bits*
**Severity:** High — silent failure of the documented emergency recovery path

```vba
If (Not m_HasOriginalMainWindowOwnedStyleBits) Or _
    (m_OriginalMainWindowHwnd <> xlHnd) Then

    m_OriginalMainWindowOwnedStyleBits = _
        CurrentStyle And TITLEBAR_OWNED_STYLE_MASK

    m_OriginalMainWindowHwnd = xlHnd
    m_HasOriginalMainWindowOwnedStyleBits = True
End If
```

**Trace of the documented recovery scenario:**

1. A workbook hides the title bar. `WS_CAPTION` is cleared on the Win32 window — that is process state, and it survives a VBA project reset.
2. The VBA project resets (code edit, `End`, unhandled error). `m_HasOriginalMainWindowOwnedStyleBits` returns to `False`.
3. The user runs `UI_ShowExcelUI` — the documented emergency path.
4. This is the first title-bar call of the module's lifetime, so the capture branch fires and stores `CurrentStyle And MASK`, which is **`0`** because the bar is currently hidden.
5. `NewStyle = merge(CurrentStyle, 0) = CurrentStyle`. The short-circuit at line 482 returns **`True`**.
6. The title bar remains hidden, and success is reported.

The README states that `UI_ShowExcelUI` is "the preferred emergency recovery command when… VBA project state was reset." It is broken in precisely that case, and because it reports success, neither the Immediate-Window path nor `UI_SetExcelUI_WithResult` surfaces the failure. v1.0's full-style restore did not have this failure mode.

**Suggested fix** — when showing with no previously captured bits, do not capture from the currently-hidden style; fall back to the full owned mask:

```vba
'------------------------------------------------------------------------------
' COMPUTE NEW STYLE
'------------------------------------------------------------------------------
    If IsVisible Then
        If m_HasOriginalMainWindowOwnedStyleBits Then
            NewStyle = UI_InternalMergeTitleBarStyleBits( _
                CurrentStyle:=CurrentStyle, _
                OwnedStyleBits:=m_OriginalMainWindowOwnedStyleBits)
        Else
            'No captured baseline (e.g. after a VBA project reset). Restore the
            'full owned frame rather than re-applying the current hidden state.
            NewStyle = UI_InternalMergeTitleBarStyleBits( _
                CurrentStyle:=CurrentStyle, _
                OwnedStyleBits:=TITLEBAR_OWNED_STYLE_MASK)
        End If
    Else
        NewStyle = UI_InternalMergeTitleBarStyleBits( _
            CurrentStyle:=CurrentStyle, _
            OwnedStyleBits:=0)
    End If
```

Additionally, restrict the capture branch so it only runs on the hide path, or guard it so an all-zero owned-bit read is never stored as the "original." Pair the fix with a regression case that drives the stateful path (see P2-8).

---

### P1-2. One failed application-property read destroys the entire snapshot

**File:** `src/M_EXCEL_UI_SNAPSHOT.bas`
**Lines:** 160–162 (raw reads), 269–277 (`Fail:` handler)
**Severity:** High — silent loss of a snapshot the caller believes exists

```vba
m_SnapshotStatusBarVisible = Application.DisplayStatusBar
m_SnapshotScrollBarsVisible = Application.DisplayScrollBars
m_SnapshotFormulaBarVisible = Application.DisplayFormulaBar
```

These are raw property reads executing under `On Error GoTo Fail`. Every other read in this module goes through `UI_RuntimeTryGetBooleanProperty`. If any of the three raises, control lands at `Fail:`, which calls `UI_SnapshotClear` and discards **everything** — Ribbon state, title-bar state, and all per-window state.

The severity comes from the caller contract: `UI_CaptureExcelUIState` is a `Sub` with no return value. A typical sequence is capture → `UI_HideExcelUI` → workflow → reset. If step 1 silently failed, the user is left in a constrained shell with no snapshot to return to — the exact scenario the component exists to prevent.

This is also asymmetric with the restore path, which routes the same three properties through `UI_RuntimeTrySetBooleanPropertyIfNeeded`.

**Suggested fix:** route all three through `UI_RuntimeTryGetBooleanProperty`, add `m_SnapshotStatusBarKnown` / `ScrollBarsKnown` / `FormulaBarKnown` flags matching the Ribbon and title-bar pattern, record per-element failures via `UI_RuntimeHandleFailure`, and continue the capture pass. Gate the corresponding restore writes on those `Known` flags (see P2-5).

---

### P1-3. `CHANGELOG.md` has no `1.1.0` entry

**File:** `CHANGELOG.md`
**Lines:** 8–9, 66–67
**Severity:** Release blocker (documentation)

The file still reads:

```markdown
## [Unreleased]

No unreleased changes are currently documented.
```

and the compare link is still `v1.0.1...HEAD`. The project's own release checklist requires a CHANGELOG update, the README's "v1.1.0 scope status" section already lists it as outstanding, and the project documents SemVer discipline explicitly. The entry needs to cover: identity-safe snapshot restoration, owned-bit title-bar handling, structured snapshot capture/restore results, the four-module decomposition, window target scopes, the new `INSTALLATION.md`, and the demo-workbook distribution policy change.

---

## P2 — Fix before tagging

### P2-1. `UI_ApplyWindowLevelState` has no error handler, breaking per-window best-effort

**File:** `src/M_EXCEL_UI.bas` — lines 1234–1312; `.Caption` reads at 1280, 1294, 1308

The procedure header states "Unexpected errors return to the caller's fail-soft handler." That handler is `UI_ApplyExcelUIState`'s `Fail:` → `Resume SafeExit`, which **aborts the remaining windows in the enumeration**. This contradicts the documented best-effort contract ("One failed element does not prevent later requested elements from being attempted").

Compounding it, `TargetWindow.Caption` is read live while constructing each failure message. On a window that is closing mid-enumeration, that read itself raises and kills the pass.

**Fix:** give the procedure a local `On Error GoTo Fail` that records the failure and returns to the loop; precompute the caption label once per window (as the snapshot path already does with `m_SnapshotWindowLabels`) rather than reading `.Caption` on each failure branch.

### P2-2. `UI_RuntimeAddFailure` can raise, from inside active error handlers

**File:** `src/M_EXCEL_UI_RUNTIME.bas` — lines 119–168

No `On Error` statement, yet the caller `UI_RuntimeHandleFailure` documents "ERROR POLICY: Does not raise." The array growth logic

```vba
If IsEmpty(FailureList) Then
    ReDim Arr(1 To 1)
Else
    Arr = FailureList
    ReDim Preserve Arr(1 To FailureCount)
End If

Arr(FailureCount) = Stage & " | " & Detail
```

assumes `FailureCount` and the array bound never desync. If they ever do, `Arr(FailureCount)` raises error 9 — from inside an active `Fail:` block, which is unrecoverable in VBA and will reset the project.

**Fix:** add `On Error Resume Next`, and derive the target bound from `UBound(Arr)` rather than trusting `FailureCount`.

### P2-3. `GetLastError` declared directly instead of using `Err.LastDllError`

**File:** `src/M_EXCEL_UI_TITLEBAR.bas` — lines 278, 579, 674, 762

Microsoft documents that the VBA runtime may issue its own API calls between a `Declare`d function returning and a separately-declared `GetLastError` executing, clobbering the thread's last-error value. VBA captures the value for you in `Err.LastDllError` immediately after the `Declare` call.

The README advertises "handles valid zero API returns using `GetLastError`" as a feature. As written, a genuine `SetWindowLong` failure returning `0` can be reported as success if the last-error value was reset in between.

**Fix:** replace the `GetLastError` / `SetLastError` declares with `Err.LastDllError` reads taken on the statement immediately following each API call.

### P2-4. Snapshot retains live `Window` COM references indefinitely

**File:** `src/M_EXCEL_UI_SNAPSHOT.bas` — lines 68, 210

Identity-safe matching is the right call, but nothing releases the retained `Window` objects when a workbook closes. Retained wrappers can prevent the Excel instance from terminating cleanly on quit, and leave zombie COM objects that only fail on next probe.

**Fix:** at minimum, document calling `UI_ClearExcelUIStateSnapshot` from `Workbook_BeforeClose`; preferably add an explicit lifecycle section to `INSTALLATION.md` alongside the import instructions.

### P2-5. Asymmetric `Known` flags across snapshot state

Ribbon, title bar, and all per-window properties have `*Known` companion flags. `StatusBar`, `ScrollBars`, and `FormulaBar` do not, so `UI_SnapshotRestoreCore` (lines 416–447) writes them unconditionally. Once P1-2 is fixed, these need `Known` flags too — otherwise a partial capture will restore default `False` values over good state.

### P2-6. `demo/M_EXCEL_UI_DEMO.bas` version metadata is stale

Header reads `VERSION 1.0.1` while all four `src/` modules, `demo/M_DEMO_BUILDER.bas`, and the test module read `1.1.0`. The release checklist includes "Update module version metadata."

### P2-7. Two documentation files contradict this branch

| File | Line | Problem |
|---|---:|---|
| `CONTRIBUTING.md` | 28 | Repository tree still shows `demo/EXCEL_UI_DEMO.xlsm` |
| `CONTRIBUTING.md` | 150 | Instructs contributors to synchronize the binary demo workbook |
| `SECURITY.md` | 201 | References "the official `demo/EXCEL_UI_DEMO.xlsm` artifact" |

This branch deletes that file (commit `14f9e98`) and `.gitignore` now excludes `demo/*.xlsm`. README and INSTALLATION were updated correctly; these two were not. Roughly a five-minute fix that otherwise ships a self-contradicting repository.

### P2-8. Test gap on the exact P1-1 code path

`TST_Case_TitleBarOwnedBitPreservation` (`test/M_EXCEL_UI_REGRESSION_TESTS.bas:3085`) exercises only `UI_InternalMergeTitleBarStyleBits` with synthetic style values. The stateful capture-and-short-circuit logic in `UI_TrySetTitleBarVisible` — where the bug lives — has no coverage. `TST_Case_TitleBarRoundTrip` (`:3026`) always starts from a visible title bar, so it never reaches the cold-start-while-hidden case.

**Suggested case:** hide the title bar via direct WinAPI (the harness already has `TST_TrySetWindowStyle`), leaving module state untouched to simulate a project reset, then assert that `UI_ShowExcelUI` actually restores `WS_CAPTION`.

### P2-9. The demo exercises none of the v1.1.0 features

No `TargetScope` usage and no `_WithResult` snapshot API usage anywhere in `demo/`. It is the artifact most users open first, and it currently demonstrates only the v1.0 surface.

### P2-10. Owned style bits are captured once and never refreshed

After a show, a subsequent hide does not re-capture (the guard at line 454 is already satisfied). Any change Excel or another add-in makes to `WS_MAXIMIZEBOX` / `WS_THICKFRAME` between operations is silently reverted on the next show.

---

## P3 — Follow-ups

| # | Finding | Location |
|---:|---|---|
| P3-1 | `Err.Number` tested without a preceding `Err.Clear`. A stale error entering the procedure makes a successful `CommandBars` read look like a failure and forces the Excel4 fallback unnecessarily. | `M_EXCEL_UI_RUNTIME.bas:405–415` |
| P3-2 | `SetWindowPos` omits `SWP_NOACTIVATE`. Standard for a frame-change refresh; without it the call can activate the window. | `M_EXCEL_UI_TITLEBAR.bas:759–760` |
| P3-3 | `UI_RuntimeBuildErrorText` and `UI_TitleBarBuildRuntimeErrorText` are byte-identical. Deliberate (keeps `TITLEBAR` dependency-free, as the README documents) — worth an inline comment saying so, or accept the dependency. | `RUNTIME.bas:641`, `TITLEBAR.bas:785` |
| P3-4 | Capture reads `.Count`, `ReDim`s, then `For Each`. A window opened by an event mid-enumeration overflows `i`. | `M_EXCEL_UI_SNAPSHOT.bas:190–207` |
| P3-5 | `UI_RuntimeClearResultBuffer` is called in `UI_SetExcelUI_WithResult` and again inside `UI_ApplyExcelUIState`. Harmless but redundant. | `M_EXCEL_UI.bas:415, 978` |
| P3-6 | `UI_HideExcelUI` / `UI_ShowExcelUI` do not accept `TargetScope`. Reasonable for the emergency path — state it explicitly in the API reference table. | `M_EXCEL_UI.bas:214, 280` |
| P3-7 | `UI_TrySetTitleBarVisibleIfNeeded` is a pure pass-through; the "IfNeeded" logic lives in the private callee. Misleading relative to the `RUNTIME` helpers that share the naming convention. | `M_EXCEL_UI_TITLEBAR.bas:152` |
| P3-8 | LongPtr sign-extension invariant is undocumented. `(CurrentStyle And Not TITLEBAR_OWNED_STYLE_MASK)` is correct today only because `GWL_STYLE` on Excel's main window never sets bit 31. Subtle enough that someone will "fix" it later. | `M_EXCEL_UI_TITLEBAR.bas:362–364` |
| P3-9 | No CI. `.github/` has templates but no workflows. | `.github/` |

### On P3-9 — a text-only CI job is viable

Excel cannot run in GitHub Actions, but a lint workflow over the exported `.bas` files would automatically catch several findings above:

- version metadata identical across all modules (would have caught P2-6);
- CRLF preserved on `.bas` per `.gitattributes` (currently correct — verified);
- `Option Explicit` and `Option Private Module` present in every module;
- procedure-header required fields present;
- CHANGELOG contains an entry matching the tag being pushed (would have caught P1-3).

---

## Suggested go/no-go

**Ship-blocking:** P1-1, P1-2, P1-3, plus P2-1 and P2-7.

**P1-1 is the priority.** It is a silent regression against the component's core recovery promise, it reports success so no diagnostic path surfaces it, and it sits in the subsystem the documentation already flags as the most environment-sensitive.

Everything else here is polish on top of a genuinely well-structured codebase. The four-module split is a real improvement over the 3,439-line monolith, and the identity-safe window restoration is the right design.

**Recommended order of work:**

1. Fix P1-1 and add the stateful title-bar regression case (P2-8).
2. Fix P1-2 and add `Known` flags for the three application-level properties (P2-5).
3. Fix P2-1 (window-loop error handling and precomputed captions).
4. Sweep documentation: P1-3 CHANGELOG, P2-7 CONTRIBUTING/SECURITY, P2-6 demo version.
5. Re-run the full harness — `RunCore`, `RunTitleBarOnly`, `RunSnapshotIdentity`, `RunAll` — plus the manual recovery and capture/hide/reset checks.
6. Treat P2-2, P2-3, P2-4, P2-9, P2-10 as either a `v1.1.1` follow-up or fold them in if the release is not time-boxed.

---

## Appendix A — Files changed

`main...release/v1.1.0` — 11 files, 6,102 insertions, 4,614 deletions.

| File | Change |
|---|---|
| `src/M_EXCEL_UI.bas` | Reduced to facade + apply worker (1,394 lines) |
| `src/M_EXCEL_UI_RUNTIME.bas` | New — 698 lines |
| `src/M_EXCEL_UI_SNAPSHOT.bas` | New — 774 lines |
| `src/M_EXCEL_UI_TITLEBAR.bas` | New — 816 lines |
| `test/M_EXCEL_UI_REGRESSION_TESTS.bas` | Expanded to 5,249 lines |
| `INSTALLATION.md` | New — 328 lines |
| `README.md` | Substantially rewritten |
| `CONTRIBUTING.md` | Condensed |
| `.github/PULL_REQUEST_TEMPLATE.md` | Updated |
| `.gitignore` | Demo `.xlsm` now excluded |
| `demo/EXCEL_UI_DEMO.xlsm` | Deleted (now a release asset) |

---

## Appendix B — Automated checks run

| Check | Result |
|---|---|
| Duplicate procedure names across modules (VBA flat namespace) | Pass — only `#If`/`#Else` bitness variants |
| `UI_*` identifiers called but never defined | Pass — none |
| Public `src/` procedures unreferenced outside their own module | Pass — none |
| Line endings on exported `.bas` | Pass — CRLF, matches `.gitattributes` |
| Trailing whitespace in `src/` | Pass — none |
| `MsgBox` / `InputBox` in `src/` | Pass — none (matches `SECURITY.md`) |
| Version metadata consistency across modules | **Fail** — `demo/M_EXCEL_UI_DEMO.bas` at `1.0.1` (P2-6) |
| References to the deleted demo workbook | **Fail** — `CONTRIBUTING.md`, `SECURITY.md` (P2-7) |
| CHANGELOG entry for the release tag | **Fail** — no `1.1.0` section (P1-3) |
| Procedures in `src/` with no error handling | 4 genuine: `UI_ApplyWindowLevelState`, `UI_RuntimeHandleFailure`, `UI_RuntimeAddFailure`, `UI_SnapshotClear` |

*Note: line numbers refer to `release/v1.1.0` at the time of review (HEAD `376fdfd`).*
