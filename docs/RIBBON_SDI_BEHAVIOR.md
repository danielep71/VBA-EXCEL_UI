# Ribbon behaviour under the Single Document Interface

> **Status:** measured on one host. Model determined: **active-window only.**
> **Characterization:** [#21](https://github.com/danielep71/VBA-EXCEL_UI/issues/21)
> (`ICR-UI-P2-01`) is complete.
> **Corrective work:** fail-closed restore is tracked by
> [#23](https://github.com/danielep71/VBA-EXCEL_UI/issues/23) for v1.1.3;
> automatic activation remains
> [#44](https://github.com/danielep71/VBA-EXCEL_UI/issues/44) for v1.2.0.
> **Produced by:** `Test_EXCEL_UI_RunRibbonSdiProbe`
> in `test/M_EXCEL_UI_REGRESSION_TESTS.bas`.

---

## 1. Why this document exists

Before this measurement, `README.md` stated the Ribbon's scope as **Excel
application**, which a reader could reasonably take to mean *one state shared
by every workbook window*. The README now records the measured active-window
scope.

Modern Excel uses the Single Document Interface: each workbook window is its own
top-level window with its own Ribbon UI. Nothing in the component verifies that
the documented scope is the scope Excel actually implements, and the component
has no way to address a specific window's Ribbon in any case — both mechanisms
it uses are application-scoped:

```text
Application.CommandBars("Ribbon").Visible      no window argument
Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")   no window argument
Application.ExecuteExcel4Macro("Show.TOOLBAR(""Ribbon"",…)")  no window argument
```

The old documented scope was therefore an **assumption**. This document replaced
it with a measurement and remains the evidence behind the current contract.

The same problem shape was already found for the title bar
(`ICR-UI-P1-01`, #14): one Boolean describing a per-window resource, restored
through whichever window happened to be active. That correction moved title-bar
restore to a retained `Window`; v1.1.3 pairs that Window with the native hWnd
read from the same object. Whether the Ribbon shared the active-window defect, was genuinely
application-wide, or was something in between is exactly what this probe
answered.

---

## 2. How to reproduce

Save unsaved work first — the probe creates and closes temporary workbooks and
toggles the Ribbon.

1. Import all four `src/` modules and `test/M_EXCEL_UI_REGRESSION_TESTS.bas`.
2. **Debug → Compile VBAProject.**
3. In the Immediate Window (**Ctrl+G**):

```vba
Test_EXCEL_UI_RunRibbonSdiProbe
```

The probe writes `EXCEL_UI_certification_*.ribbon.txt` and `.ribbon.json` to
`%TEMP%` and prints both. Paste the text table into a new block in section 4
below, one block per host tested.

---

## 3. What is measured, and why three mechanisms

Each observation records the Ribbon through every mechanism available, because
they can legitimately disagree and the disagreement is itself a finding.

| Reading | Why it is recorded |
|---|---|
| `CommandBars("Ribbon").Visible` | What the component's primary read returns. |
| `CommandBars("Ribbon").Height` | The more sensitive signal. A Ribbon that is *collapsed* rather than hidden can report `Visible = True` while its height falls to the tab strip. A component trusting `Visible` alone would call that state shown. |
| `Get.ToolBar(7,"Ribbon")` | The component's XLM fallback, used when the object model refuses. |

Because none of the three takes a window argument, a per-window Ribbon can only
surface as a **difference between readings taken while different windows are
active**, or as a disagreement between the mechanisms.

### Scenarios

| # | Scenario | Question it answers |
|---|---|---|
| 1 | Baseline, window A | What does a visible Ribbon read as on this host? |
| 2 | Hide with A active, observe A then B | Does hiding affect a window that already existed? |
| 3 | Show with B active, observe B then A | Is the behaviour symmetric? It need not be. |
| 4 | Create window C *after* a hide, observe C then A | Does a new window inherit the hidden state? This is the case a component storing one Boolean cannot reason about at all. |
| 5 | Capture on A, hide, activate B, restore, observe B then A | Does the snapshot contract hold for the Ribbon? |

---

## 4. Measurements

> One block per host, recorded verbatim from the probe. Interpretation belongs
> in section 5, not here.

### Host 1 — Excel 16.0 build 20131, Windows (64-bit) NT 10.00, x64, VBA7

Recorded 2026-08-19 18:51:14.

```text
scenario              window  CommandBars.Visible  Height  XLM
1-Baseline            A       True                 178     True
2-HiddenOnA           A       False                178     False
2-HiddenOnA           B       True                 178     True
3-ShownOnB            B       True                 178     True
3-ShownOnB            A       False                178     False
4-NewWindowAfterHide  C       False                178     False
4-NewWindowAfterHide  A       False                178     False
5-RestoredFromB       B       True                 178     True
5-RestoredFromB       A       False                178     False
```

---

## 5. Interpretation

### The Ribbon is per-window

Scenario 2 settles it. Hiding the Ribbon while window A was active left A
reporting hidden and **B reporting visible**. Scenario 3 is the symmetric
confirmation: showing it while B was active left B visible and **A still
hidden** from the previous scenario.

Each write therefore affects the active window and nothing else. The documented
scope, *Excel application*, is wrong for this host.

### A window created after a hide inherits the hidden state

Scenario 4: with the Ribbon hidden and A active, a newly created window C
reported hidden. Note the limit of what this shows — at that moment the "last
written state" and "the active window's state" were both hidden, so this
measurement cannot separate *inherits from the active window* from *inherits
from the last write*. Distinguishing them needs a scenario that sets them
differently, and no current decision depends on the answer.

### Scenario 5 is a live defect, not a curiosity

The sequence was: show with A active, capture, hide A, activate B, restore.

The captured value was read from **A**. The restore applied it to **B**, which
was merely the window that happened to be active. A was left hidden. Nothing was
reported.

```text
captured from  A  (Ribbon visible)
restored to    B  (already visible - no-op)
A              still hidden, never restored
result         success
```

That is the same defect as `ICR-UI-P1-01`, in the same snapshot, on a different
element. It was invisible before this probe because no test ever activated a
second window across a Ribbon capture and restore.

### The `Height` column is inert on this host

`CommandBars("Ribbon").Height` reported **178 in all nine observations**,
including every hidden one. It does not track visibility on this build and is
worthless as a signal here. The *collapsed versus hidden* ambiguity the column
was added to detect did not arise.

### The two read mechanisms agree exactly

`CommandBars("Ribbon").Visible` and `Get.ToolBar(7,"Ribbon")` returned the same
value in **all nine observations**. The component's XLM fallback is faithful to
its primary read, at least on this host, so a divergence between them is not a
risk that needs designing around.

### Model

| Model | Verdict |
|---|---|
| Application-wide | **Ruled out** by scenarios 2 and 3. |
| **Active-window only** | **Selected.** Every write affects the active window alone. |
| Cached and propagated | **Partly true** — new windows inherit at creation (scenario 4) — but this is a property of window creation, not of the write scope. |
| Host-dependent | Not ruled out. One host is one data point. |

## 6. Decision and consequences

The measured model has the same shape as the title-bar defect fixed in `1.1.1`,
but **it cannot take the same fix**, and that difference drives the split below.

The title bar could be corrected inside a patch release because
`SetWindowLong` accepts an explicit `HWND`: the component could simply write to
the captured window without touching anything else. The Ribbon has no such API.
Every mechanism available is application-scoped and acts on the active window:

```text
Application.CommandBars("Ribbon").Visible                     no window argument
Application.ExecuteExcel4Macro("Show.TOOLBAR(""Ribbon"",…)")  no window argument
```

Restoring the Ribbon to a specific window therefore requires **activating that
window**, writing, and restoring focus. That is a visible side effect which can
fire `Workbook_WindowActivate` handlers in caller code, and it does not belong
in a corrective patch release.

### Characterization completed for `1.1.1`

- [x] Model determined and recorded above.
- [x] `README.md` states the measured active-window scope instead of claiming
      application scope.

### Required for `1.1.3`

- [ ] The snapshot **reports** the Ribbon as unrestorable when the captured
      window is not active, rather than applying the captured value to whichever
      window is. This fail-closed correction is #23. It activates no window and
      therefore introduces no new focus or event side effect.

### Deferred to `1.2.0`

- [ ] Under #44, restore the Ribbon to its captured window by activating it,
      writing, and restoring focus — with an opt-out, because the activation is
      observable.
- [ ] Decide and document whether `UI_SetExcelUI(Ribbon:=…)` should apply to
      every window or only the active one, which is a public-contract question
      the `UIWindowTargetScope` enum already raises for other elements.
- [ ] Establish whether a new window inherits from the active window or from the
      last write, if either behaviour is to be promised.

## 7. Limitations

- The probe measures what the host reports, not what is drawn on screen. On the
  host measured, `Height` proved inert and could not have detected a collapsed
  Ribbon; if a future host reports differently, that column becomes meaningful
  again.
- One host is one data point. Ribbon behaviour can vary by Office channel,
  update ring and administrative policy; a single green block does not establish
  behaviour for the supported range.
- Any add-in that manipulates the Ribbon can influence these readings. Prefer a
  clean Excel session with other add-ins disabled.
