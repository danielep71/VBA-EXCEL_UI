# Ribbon behaviour under the Single Document Interface

> **Status:** measurements pending.
> **Issue:** [#21](https://github.com/danielep71/VBA-EXCEL_UI/issues/21) —
> `ICR-UI-P2-01`, Ribbon scope under SDI is unspecified and unverified.
> **Produced by:** `Test_EXCEL_UI_RunRibbonSdiProbe`
> in `test/M_EXCEL_UI_REGRESSION_TESTS.bas`.

---

## 1. Why this document exists

`README.md` states the Ribbon's scope as **Excel application**, which a reader
will reasonably take to mean *one state shared by every workbook window*.

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

So the documented scope is currently an **assumption**. This document replaces it
with a measurement.

The same problem shape was already found and fixed for the title bar
(`ICR-UI-P1-01`, #14): one Boolean describing a per-window resource, restored
through whichever window happened to be active. Whether the Ribbon shares that
defect, is genuinely application-wide, or is something in between is exactly what
the probe answers.

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

> Paste probe output here, one block per host. Do not summarise — record the
> table verbatim, then interpret it in section 5.

### Host 1 — *(pending)*

```text
Excel <version> build <build> | <operating system> | <bitness>

scenario              window  CommandBars.Visible  Height  XLM
```

---

## 5. Interpretation

*(to be completed once at least one host has been measured)*

The measurements will place the Ribbon into exactly one of these models. Each
carries a different obligation for the component:

| Model | Meaning | What the component must then do |
|---|---|---|
| **Application-wide** | One state, every window agrees | Nothing. `README.md` is already correct and the snapshot contract holds. |
| **Active-window only** | Each window has its own state; the API addresses the active one | Same class of defect as `ICR-UI-P1-01`. The snapshot must capture the Ribbon's window identity, and `README.md` must stop claiming application scope. |
| **Cached and propagated** | New windows inherit the state at creation | `README.md` must say so explicitly, and scenario 4 governs what a caller can expect after opening a workbook. |
| **Host-dependent** | Differs by build, channel or policy | The component must document the scope as best effort and name the builds tested, exactly as it already does for title-bar control. |

---

## 6. Decision and consequences

*(to be completed)*

- [ ] Model selected, with the measurements that justify it.
- [ ] `README.md` Ribbon scope statement corrected to match (tracked on #19).
- [ ] If the model is anything other than application-wide, a follow-up issue
      opened for the propagation or identity work, scoped to `v1.2.0` — it
      changes observable behaviour and does not belong in a patch release.
- [ ] Assertions added to the regression suite **only after** the model is
      decided. Writing them first would encode the guess this document exists to
      remove.

---

## 7. Limitations

- The probe measures what the host reports, not what is drawn on screen. A
  Ribbon that is visually collapsed but reported as visible is detectable only
  through the `Height` column.
- One host is one data point. Ribbon behaviour can vary by Office channel,
  update ring and administrative policy; a single green block does not establish
  behaviour for the supported range.
- Any add-in that manipulates the Ribbon can influence these readings. Prefer a
  clean Excel session with other add-ins disabled.
