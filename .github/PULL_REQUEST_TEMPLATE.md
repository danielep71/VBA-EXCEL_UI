<!--
  Sections that do not apply can be deleted outright.
  A template filled with "Not applicable" fifteen times hides the parts that
  matter, so deleting is preferred over padding.

  The collapsed sections near the bottom are relevant only when the change
  touches that subsystem. Expand the ones that apply; delete the rest.
-->

## 📌 Summary

<!-- What changed, and why it was needed. One paragraph is usually enough. -->

## 🔗 Related issue

```text
Closes #
```

---

## 🏷️ Type of change

- [ ] 🐛 Functional or compatibility fix
- [ ] 🆘 Recovery or host-state fix
- [ ] ✨ Backward-compatible feature
- [ ] ♻️ Internal refactor with no intended public behavior change
- [ ] 🧪 Regression-test change
- [ ] 🖼️ Demo change
- [ ] 📖 Documentation-only change
- [ ] 🧹 Repository or release maintenance
- [ ] 🔐 Security-related change

## 🎚️ Affected surface

**Application-level**

- [ ] 📊 Status Bar
- [ ] ↕️ Scroll Bars
- [ ] ƒ Formula Bar

**Active-window** — per workbook window; no API accepts a window argument

- [ ] 🎀 Ribbon
- [ ] 🪟 Title Bar / WinAPI

**Targetable window-level** — honours `UIWindowTargetScope`

- [ ] 🔢 Headings
- [ ] 📑 Workbook Tabs
- [ ] ▦ Gridlines

**Subsystem**

- [ ] 📸 Snapshot capture
- [ ] 🔄 Snapshot restoration
- [ ] 🧾 Structured diagnostics
- [ ] ⚙️ `ScreenUpdating`
- [ ] 🧱 Module dependencies
- [ ] 🖼️ Demo
- [ ] 🧪 Tests
- [ ] 📖 Documentation only

---

## 🔒 Public API and Semantic Versioning

```text
Public behavior changed:
Backward compatible:
Suggested release:        patch / minor / major
Migration required:
```

Confirm changes to names, signatures, parameter order and defaults, enum values,
targeting, snapshot meaning, diagnostics and recovery. Write
`No public behavior change` where applicable.

> [!IMPORTANT]
> Adding or removing a `Public` member in `src/` requires an intentional edit to
> `tools/public_api_manifest.txt`, and CI fails otherwise. That friction is
> deliberate: a change to the public surface is exactly what breaks callers, and
> it is invisible in a diff of several thousand lines.

---

## ✅ Testing performed

```text
Debug → Compile VBAProject             →
python3 tools/check_repo.py            →
Test_EXCEL_UI_RunReleaseCertification  →
Manual UI_HideExcelUI / UI_ShowExcelUI →
Manual capture / hide / reset          →
```

**Certification verdict**

```text
RESULT:
```

> [!CAUTION]
> All four counters form the verdict. `failed=0` alone is not a pass:
> `INCOMPLETE`, a non-zero `skipped` count, or `cleanup=FAILED` each mean the
> run cannot certify this change.

<details>
<summary>🔬 Narrower runners used while iterating</summary>

<br>

```text
Test_EXCEL_UI_RunAll                  →
Test_EXCEL_UI_RunCore                 →
Test_EXCEL_UI_RunTitleBarOnly         →
Test_EXCEL_UI_RunSnapshotIdentity     →
Test_EXCEL_UI_RunTitleBarSdiIdentity  →
Test_EXCEL_UI_RunRibbonSdiProbe       →
```

</details>

## 🖥️ Validation environment

The certification runner writes a JSON document and a text report to `%TEMP%`,
both naming the exact host. **Attaching one replaces the fields below**, and
removes the transcription errors that come with retyping build numbers.

<details>
<summary>Or record the environment by hand</summary>

<br>

```text
Excel product/version/build:
Office bitness:
Windows version:
Workbook type:
Excel window state:
Open Excel windows:
Other add-ins:
```

</details>

List only environments actually tested.

---

## 📋 Always required

### 🧹 Source

- [ ] Current branch was confirmed before committing.
- [ ] All four required production modules were present during compilation.
- [ ] Changed modules were exported to the correct repository paths.
- [ ] CRLF was preserved for exported VBA source.
- [ ] No conflict markers or duplicate procedures remain.
- [ ] The textual diff contains only intended changes.
- [ ] No lock, backup, generated, confidential, credential, client, or
      production-data file is included.

### 🔒 Compatibility

- [ ] Existing public names, signatures, defaults and enum values remain
      compatible, or the breaking rationale is explicit.
- [ ] `UI_ShowExcelUI` remains an emergency show-all path.
- [ ] Best-effort continuation remains deliberate.
- [ ] Failures are not silently discarded.
- [ ] `ScreenUpdating` is restored.
- [ ] No unsolicited production `MsgBox` was introduced.
- [ ] Invalid enum values remain controlled.

### 📖 Documentation

- [ ] README · INSTALLATION · CONTRIBUTING · Wiki · module headers · demo
      guidance, as affected
- [ ] `CHANGELOG.md` entry added
- [ ] No documentation change required
- [ ] No binary demo workbook change is included, **or** the binary change is
      intentional, described, synchronized with exported source, and tested

> [!TIP]
> Documentation belongs in **this** pull request, not a follow-up. A follow-up
> documentation commit is a commit that does not get written.

---

## 🧱 Module ownership and dependencies

<details>
<summary>Expand when the change touches module boundaries or shared state</summary>

<br>

```text
M_EXCEL_UI:
M_EXCEL_UI_RUNTIME:
M_EXCEL_UI_SNAPSHOT:
M_EXCEL_UI_TITLEBAR:
Dependency graph changed:
Circular dependency introduced:   no / explain
Mutable state duplicated:         no / explain
```

- [ ] `M_EXCEL_UI` remains the public facade.
- [ ] Runtime and title-bar modules have no project-module dependencies.
- [ ] Snapshot state exists only in `M_EXCEL_UI_SNAPSHOT`.
- [ ] Title-bar mutable state exists only in `M_EXCEL_UI_TITLEBAR`.
- [ ] No circular dependency was introduced.
- [ ] Internal modules retain `Option Private Module`.
- [ ] Any new test seam is documented as unsupported **and has a caller**.

</details>

## 📸 Snapshot and recovery

<details>
<summary>Expand when the change touches capture, restoration or identity</summary>

<br>

```text
Captured state:
Window identity strategy:
Behavior for new windows:
Behavior for missing/closed/recreated windows:
Behavior after VBA reset:
Failure ordering:
Emergency recovery path:
```

- [ ] Per-window restore does not use collection index.
- [ ] Retained Window identity behavior is documented and tested.
- [ ] New windows remain unchanged.
- [ ] Missing captured windows produce controlled diagnostics.
- [ ] Reset-without-snapshot remains controlled.
- [ ] In-memory lifetime remains documented.
- [ ] State is never applied to an object that cannot be proven to be the one it
      was captured from.

</details>

## 🪟 Ribbon or WinAPI method

<details>
<summary>Expand when the change touches the Ribbon or the window frame</summary>

<br>

```text
API or command used:
Owned style bits:
32-bit path:
64-bit path:
GetLastError treatment:
Frame refresh:
Target window treatment:
Unrelated style bits preserved:
```

- [ ] Exact owned style mask is preserved or deliberately reviewed.
- [ ] Unrelated current style bits are preserved.
- [ ] 32-bit and 64-bit declarations are correct.
- [ ] Valid zero returns are distinguished from failures.
- [ ] Required frame refresh is performed, and a failed refresh is retried
      rather than short-circuited.
- [ ] Multi-window behavior was tested with more than one workbook open.

</details>

## 🧾 Diagnostics and failure policy

<details>
<summary>Expand when the change touches failure handling or result contracts</summary>

<br>

```text
Failure contract:
Logging contract:
Structured-result contract:
ScreenUpdating restoration:
```

- [ ] Insertion order is preserved in structured diagnostics.
- [ ] Anything reachable from an error handler cannot itself raise.
- [ ] Outputs that cannot fail are set before anything that can.

</details>

---

## 💬 Reviewer notes

<!--
  Trade-offs, known limitations, environments not tested, and follow-up work.
  A limitation stated here is a decision. The same limitation discovered later
  is a defect.
-->
