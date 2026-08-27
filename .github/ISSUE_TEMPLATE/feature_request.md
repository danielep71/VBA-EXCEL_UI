---
name: "✨ Feature request"
about: "Propose a backward-compatible UI, targeting, recovery, diagnostics, demo, or assurance capability"
title: "[Feature] "
labels: "enhancement"
assignees: ""
---

<!--
Thank you for proposing an improvement to VBA-EXCEL_UI.

Start with the user problem and observable outcome. Maintainers may choose a
different implementation.

Before submitting:
- Search open and closed issues, README, INSTALLATION, and the Wiki.
- Keep one request focused on one user outcome.
- Use the bug template when existing documented behavior is incorrect.
- Remove client data, credentials, workbook internals, and private material.

New behavior normally belongs in a minor release or later planned milestone.
Do not assume a correctness-and-hardening patch should absorb an adjacent
behavioral feature.

Keep the headings below and replace prompts with the proposal.
-->

## 🎯 Problem and user story

<!--
Describe the limitation in current behavior. Who experiences it, under what
conditions, and what practical cost does it create?

Suggested form: As a <user>, I need <capability>, so that <outcome>.
-->



## 🔧 Current workaround and cost

<!--
What do users do today? Explain extra code, activation side effects, reliability
risk, recovery burden, compatibility cost, or missing evidence. Write
"No known workaround" if true.
-->



## 🧭 Desired outcome

<!--
Describe success from the caller’s perspective without requiring one particular
implementation. Include observable state, scope, diagnostics, and recovery.
-->



## ✅ Acceptance criteria

<!--
List verifiable outcomes. Cover normal behavior, important boundaries,
failure/refusal, compatibility, cleanup, and documentation. Avoid task-only
criteria such as "add a module."
-->

- [ ] <!-- Observable normal-path outcome -->
- [ ] <!-- Scope, identity, or boundary outcome -->
- [ ] <!-- Failure or fail-closed outcome -->
- [ ] <!-- Recovery and cleanup outcome -->
- [ ] <!-- Compatibility and documentation outcome -->

---

## 🎚️ Affected area

**Managed UI**

- [ ] Ribbon
- [ ] Status Bar
- [ ] Scroll Bars
- [ ] Formula Bar
- [ ] Headings
- [ ] Workbook Tabs
- [ ] Gridlines
- [ ] Title Bar / native Excel frame

**Runtime or repository**

- [ ] Public UI API
- [ ] UIVisibility or UIWindowTargetScope semantics
- [ ] Application-level properties
- [ ] Active-window behavior
- [ ] Window-level targeting
- [ ] Snapshot capture, ownership, or lifecycle
- [ ] Snapshot restoration or identity
- [ ] Structured diagnostics
- [ ] ScreenUpdating ownership
- [ ] Emergency recovery
- [ ] WinAPI or Office bitness
- [ ] Demo source or release workbook
- [ ] Regression or certification harness
- [ ] Documentation, tooling, CI, or release evidence
- [ ] Other

**Why these areas are affected**



---

## 🧭 Targeting and ownership

### Intended target

- [ ] Current Excel Application instance
- [ ] Every current Excel window
- [ ] Active window only
- [ ] Windows belonging to ActiveWorkbook
- [ ] A caller-supplied Excel Window
- [ ] A caller-supplied Workbook
- [ ] Native frame associated with a verified Window/hWnd pair
- [ ] Other

### Ownership model

<!--
Explain who owns any captured baseline or temporary state, how nested callers
interact, what happens when another workbook/add-in changes the same UI, and
when retained references are released.
-->



### Activation and SDI behavior

<!--
Does the feature activate another Excel window? Can active-window changes alter
the target? What happens when a window closes, is recreated, or receives a new
hWnd? State any visible activation or focus side effect.
-->



---

## 🔄 Failure, recovery, and diagnostics

### Failure behavior

- [ ] Fail closed before mutation when the intended target cannot be proved
- [ ] Best-effort continuation across independent UI elements
- [ ] All-or-nothing behavior
- [ ] Caller-selectable policy
- [ ] Other

Explain which partial outcomes are permitted:



### Diagnostic contract

- [ ] Existing fire-and-forget procedure is sufficient
- [ ] Boolean structured result
- [ ] FailureCount and ordered FailureList
- [ ] New machine-readable category or type
- [ ] Immediate Window logging
- [ ] Existing diagnostic contract remains unchanged

Describe proposed stages, return values, defaults, and error behavior:



### Recovery and lifecycle

Explain:

- whether prior state must be captured;
- whether the operation is reversible;
- what happens when a snapshot already exists;
- what happens after a VBA project reset;
- what UI_ShowExcelUI should do;
- when retained Window references are released;
- whether a full Excel restart can be required.



---

## 🧩 Candidate API or design

<!--
Optional. Show how the capability might feel to a VBA caller. Public names,
parameters, defaults, enums, and result contracts become compatibility
commitments, so prefer the smallest surface that expresses the outcome.
-->

~~~vb
'Optional illustrative API — maintainers may choose another design.
~~~

| Contract question | Proposal |
|---|---|
| Public procedure or enum | |
| Parameter names, order, and types | |
| Defaults | |
| Return value | |
| FailureCount / FailureList impact | |
| Target scope | |
| Snapshot impact | |
| Recovery behavior | |

---

## ⚖️ Compatibility and release impact

### Public behavior

- [ ] Internal or documentation-only; no public behavior change
- [ ] Backward-compatible additive functionality
- [ ] Backward-compatible correction to documented behavior
- [ ] Changes an existing default, result, error, targeting, or recovery semantic
- [ ] Removes or renames public API, or otherwise requires caller changes

### Deployment

- [ ] The existing four-module production package remains complete
- [ ] Adds or changes required runtime files or import guidance
- [ ] No new external dependency
- [ ] Adds an external dependency or host prerequisite
- [ ] Must support both 32-bit and 64-bit Office source paths
- [ ] Intentionally targets a narrower Excel/Windows environment

~~~text
Existing calls affected:
Backward compatible:     Yes / No / Unsure
Migration required:
Suggested release:       patch / minor / major / unsure
Suggested milestone:     v1.2.0 / future / unsure
~~~

Explain caller impact, migration, deprecation, and the oldest intended host:



<details>
<summary>Semantic-versioning guidance</summary>

<br>

- **Patch:** backward-compatible defect corrections.
- **Minor:** backward-compatible new behavior.
- **Major:** incompatible public API or behavioral changes.

A correctness fix can change defective behavior without automatically requiring
a major release. Maintainers assign the final milestone and release.

</details>

---

## 🏗️ Design constraints and invariants

<!-- Explain how the proposal preserves or intentionally changes each applicable rule. -->

- M_EXCEL_UI remains the supported public facade.
- The production package remains source-first and reviewable.
- UI_LeaveUnchanged never becomes an accidental mutation.
- Application-level and window-level scopes remain explicit.
- Ribbon command text remains fixed and non-injectable.
- Active-window side effects are observable and documented.
- Native title-bar writes preserve every unowned style bit.
- Window object identity and native hWnd identity are paired defensibly.
- An unprovable target fails closed before state is written.
- Snapshot ownership, replacement, retention, and clearing remain explicit.
- FailureCount remains authoritative and FailureList remains best effort.
- ScreenUpdating is restored to the caller-visible baseline.
- UI_ShowExcelUI remains an independent emergency recovery path.
- No behavior is described as certified without exact-source runtime evidence.

**Applicable constraints and proposed treatment**



---

## 🧪 Verification plan

<!--
Static CI does not run Excel. Runtime features normally require focused Excel
checks plus regression and cleanup evidence tied to the exact source.
-->

| Verification area | Proposed case and expected evidence |
|---|---|
| Normal path | |
| No-op / already-correct state | |
| Invalid argument or unavailable target | |
| Active-window transition | |
| Multiple workbooks/windows | |
| Closed, recreated, or new window | |
| Native-handle reuse, if relevant | |
| Failure or fault injection | |
| Partial-write and recovery | |
| Snapshot ownership | |
| ScreenUpdating cleanup | |
| Win32 / 32-bit Office | |
| Win64 / 64-bit Office | |
| Static repository gate | |
| Release certification | |
| Demo or release asset | |

### Proposed regression coverage

- [ ] Core regression case
- [ ] Title-bar regression case
- [ ] Snapshot identity case
- [ ] Ribbon SDI characterization
- [ ] Multi-window test
- [ ] Failure-path or fault-injection test
- [ ] Captionless or same-style native-frame test
- [ ] Caller-owned snapshot preservation
- [ ] Cleanup proof
- [ ] Public API contract gate
- [ ] Manual recovery test
- [ ] Demo update
- [ ] 32-bit Office runtime validation
- [ ] 64-bit Office runtime validation

List proposed test names or assertions:



---

## 🔀 Alternatives and non-goals

### Alternatives considered

<!--
Consider composing existing UI_SetExcelUI calls, caller-side code, current
snapshot APIs, documentation-only guidance, a workbook-specific macro, custom
Ribbon XML/add-in code, or leaving the capability outside this repository.
-->



### Explicit non-goals

<!--
Prevent the request from absorbing adjacent behavior. Link separate proposals.
For example, a v1.1.3 wrong-target correction does not automatically include
v1.2.0 automatic Ribbon activation or demo modernization.
-->



---

## 📚 Documentation and additional context

Identify expected changes:

- [ ] README
- [ ] INSTALLATION
- [ ] SECURITY
- [ ] CHANGELOG
- [ ] CONTRIBUTING
- [ ] Wiki
- [ ] Public API manifest
- [ ] Module and procedure headers
- [ ] Demo guidance or asset
- [ ] Release evidence or certification inventory
- [ ] No documentation impact

Include related issues, prior art, sanitized screenshots, or pseudocode:



Do not include credentials, client data, proprietary workbook content, private
reviews, or sensitive environment information.

---

## ✅ Requester checklist

- [ ] I searched open and closed issues, README, INSTALLATION, and the Wiki
- [ ] I described a user problem and observable outcome, not only an implementation
- [ ] I used the bug template if existing behavior is incorrect
- [ ] I supplied testable acceptance criteria
- [ ] I identified targeting, ownership, activation, and recovery behavior
- [ ] I considered compatibility, deployment, Office bitness, and release impact
- [ ] I separated corrective patch scope from new behavioral feature scope
- [ ] I proposed runtime verification in addition to static checks where required
- [ ] I identified explicit alternatives and non-goals
- [ ] I removed credentials, client data, private material, and sensitive names
- [ ] I understand maintainers may choose another design and assign the milestone
