<!--
Thank you for contributing to VBA-EXCEL_UI.

Complete every applicable section. Use "N/A — reason" where evidence or a
section genuinely does not apply. Delete instructional comments before the PR
is ready for review.

Do not manufacture evidence. Do not describe a local run, another branch, an
earlier commit, or a similar source tree as certification of the PR head.

The hosted Static checks workflow performs text and repository analysis only.
It does not import or compile VBA, start Excel, call Windows APIs, exercise
Office-bitness branches, validate SDI window identity, run the regression
harness, or certify a release workbook.
-->

## 🎯 Summary

<!--
In one or two paragraphs explain:
- the problem or requirement;
- the observable behavior that changes;
- why this implementation was chosen.

Describe more than the files or procedures edited.
-->



### Linked issues

<!-- Use a closing keyword only when this PR fully satisfies the issue. -->

Closes #

Related to #

### Acceptance-criteria status

<!--
Map each acceptance criterion to implementation and evidence. If an issue will
remain partly open, say so explicitly and do not use a closing keyword.
-->

| Acceptance criterion | Implementation | Verification | Status |
|---|---|---|---|
| <!-- criterion --> | <!-- file/procedure --> | <!-- test/evidence --> | <!-- met / not met --> |

---

## 🧭 Change classification

<!-- Select every category that applies. -->

- [ ] Bug or compatibility fix
- [ ] Recovery or host-state fix
- [ ] Backward-compatible enhancement
- [ ] Refactor with no intended behavior change
- [ ] Public API or contract change
- [ ] Regression-test or fault-injection change
- [ ] Demo source or release-asset change
- [ ] Documentation only
- [ ] Repository tooling, workflow, or governance
- [ ] Release preparation or evidence
- [ ] Security-related change

### Affected technical boundaries

**Managed UI**

- [ ] Ribbon
- [ ] Status Bar
- [ ] Scroll Bars
- [ ] Formula Bar
- [ ] Headings
- [ ] Workbook Tabs
- [ ] Gridlines
- [ ] Title Bar / native frame

**Runtime and contracts**

- [ ] Target-scope resolution
- [ ] Snapshot capture, ownership, or lifecycle
- [ ] Snapshot restoration or window identity
- [ ] Structured diagnostics or failure semantics
- [ ] Application.ScreenUpdating ownership
- [ ] WinAPI declaration or Office bitness
- [ ] Public enum, procedure, parameter, default, or result contract
- [ ] Required module set or dependency graph
- [ ] Regression runner, mandatory inventory, cleanup, or evidence
- [ ] Demo build or macro-enabled release artifact
- [ ] Repository automation or release provenance
- [ ] None of the above

---

## 📐 Scope and contract impact

### In scope

<!-- What this PR deliberately changes. -->

- <!-- item -->

### Out of scope

<!--
State work a reviewer might reasonably expect but this PR deliberately defers.
Name the follow-up issue and milestone. Do not silently move behavior between
v1.1.3 correctness/hardening and v1.2.0 feature work.
-->

- <!-- item / issue / milestone -->

### Public API and compatibility

~~~text
Public behavior changed:
Backward compatible:
Suggested release:        patch / minor / major
Migration required:
~~~

Select the most accurate statement:

- [ ] No supported public API name, signature, parameter, default, enum value,
      target-scope rule, result contract, or recovery behavior changes
- [ ] Additive and source-compatible public surface change
- [ ] Existing behavior changes to correct a defect
- [ ] Breaking API or deployment change
- [ ] Documentation or tooling only; runtime compatibility is unchanged

**Compatibility or migration notes:**

<!--
Include any caller action, import change, host requirement, support-policy
change, or reason an existing workbook can observe different behavior.
-->



> [!IMPORTANT]
> Any change to a Public declaration under src requires an intentional update to
> tools/public_api_manifest.txt, which records the full normalised signature:
> parameter order and names, ByVal/ByRef, types, Optional status, defaults,
> return types and enum values. Regenerate it with tools/vba_api.py --write.
>
> Passing the manifest gate proves the declared contract matches that file. It
> does not prove behavior. A procedure whose signature is untouched can still
> change what it does, and that belongs in the statement above.
>
> A change under [supported] is a Semantic Versioning event for external
> callers and must be declared as one here and in CHANGELOG.md. The gate
> enforces that: the manifest carries the facade as it stood at the last
> release in its [baseline] section, and when the two differ the CHANGELOG
> Supported API contract row has to name the release type or the build fails.
> Regenerating the manifest alone does not clear it.
>
> A change under [project-public] only has to be deliberate. Neither is the
> same as the deployment rule that all four src modules are replaced together.

### Production source package

- [ ] The complete four-module production package remains:
      M_EXCEL_UI_RUNTIME, M_EXCEL_UI_TITLEBAR, M_EXCEL_UI_SNAPSHOT,
      and M_EXCEL_UI
- [ ] Required source files changed; installation, inventories, and release
      guidance were updated
- [ ] Not applicable — no production source or package change

---

## 🔧 Implementation notes

<!--
Summarize the design, invariants, important control flow, and non-obvious
trade-offs. Address application-level versus window-level state, active-window
behavior, identity proof, fail-closed decisions, cleanup, and recovery where
relevant.
-->



### Files and responsibilities

| File | Responsibility in this PR |
|---|---|
| <!-- path --> | <!-- change and reason --> |

---

## ✅ Verification

### Static repository gate

<!--
Run the commands against the exact PR head. Link the GitHub Actions run when it
is available. The current workflow does not upload a separate result artifact.
-->

| Evidence | Result |
|---|---|
| Exact PR head SHA | <!-- full 40-character SHA --> |
| Working tree used locally | <!-- clean / dirty; a dirty tree cannot certify exact head --> |
| python3 tools/check_repo.py | <!-- PASS / FAIL / NOT RUN --> |
| python3 tools/reformat.py --check src/*.bas test/*.bas demo/*.bas | <!-- PASS / FAIL / NOT RUN --> |
| git diff --check | <!-- PASS / FAIL / NOT RUN --> |
| Static checks workflow | <!-- PASS / FAIL / PENDING / NOT RUN --> |
| Workflow run URL | <!-- direct GitHub Actions URL / N/A --> |

### Excel and VBA execution

<!--
Choose one applicability statement. Production VBA, regression, demo, package,
and runtime-contract changes require execution against source demonstrably
bound to the PR head. A static pass is not a substitute.
-->

- [ ] Required and completed against the exact PR head source
- [ ] Required but incomplete — reason and release consequence documented below
- [ ] Not required — documentation/repository-only change with no executable
      source, runtime-contract, test, builder, or packaging impact

| Evidence | Result |
|---|---|
| Tested commit SHA | <!-- full 40-character SHA / N/A --> |
| Source identity method | <!-- clean checkout, hashes, export manifest, other / N/A --> |
| Source or package tested | <!-- exact PR source, demo workbook, add-in, other / N/A --> |
| Debug → Compile VBAProject | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Regression entry point | <!-- runner name / N/A --> |
| Certification verdict | <!-- PASS / FAIL / INCOMPLETE / NOT RUN / N/A --> |
| Mandatory units | <!-- expected / executed / skipped / N/A --> |
| Failures | <!-- number / N/A --> |
| Cleanup outcome | <!-- OK / FAILED / not observed / N/A --> |
| Evidence files | <!-- generated TXT/JSON names or approved links / N/A --> |
| Excel version and build | <!-- exact value / N/A --> |
| Office bitness | <!-- 32-bit / 64-bit / N/A --> |
| Windows version | <!-- version/build / N/A --> |
| Workbook and window setup | <!-- workbooks, windows, active-window transitions / N/A --> |
| Other relevant add-ins or policy | <!-- detail / none observed / N/A --> |

> [!CAUTION]
> A release-certification verdict requires every mandatory unit, zero
> unexplained skips, zero failures, and successful cleanup. A single green
> counter is not enough.
>
> The current runner can write text and JSON host evidence, but exact-source
> identity must still be demonstrated. Evidence from an earlier commit does not
> certify a later PR head merely because the VBA diff appears equivalent.

### Targeted and manual checks

<!--
List focused checks beyond the standard runner: wrong-target refusal,
activation changes, closed/recreated windows, recycled-handle seams,
captionless baselines, caller-owned snapshots, diagnostic allocation failure,
recovery, demo controls, or release-asset smoke tests.
-->

| Scenario | Expected result | Actual result / evidence |
|---|---|---|
| <!-- scenario --> | <!-- expected --> | <!-- PASS / FAIL / N/A + detail --> |

### Narrower iteration runners

<!-- These aid diagnosis and iteration; they do not replace the release gate. -->

| Runner | Result |
|---|---|
| Test_EXCEL_UI_RunCore | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Test_EXCEL_UI_RunTitleBarOnly | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Test_EXCEL_UI_RunSnapshotIdentity | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Test_EXCEL_UI_RunTitleBarSdiIdentity | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Test_EXCEL_UI_RunAll | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Test_EXCEL_UI_RunRibbonSdiProbe | <!-- characterization only / result / N/A --> |

> [!NOTE]
> Test_EXCEL_UI_RunRibbonSdiProbe is characterization, not release
> certification. It must not be used to imply automatic Ribbon activation or a
> v1.2.0 behavior change.

---

## 🧪 Regression coverage

- [ ] Existing cases cover the changed behavior
- [ ] New or amended cases cover the success path
- [ ] Failure, partial-write, refusal, and cleanup paths are covered where
      applicable
- [ ] Multi-workbook or multi-window behavior is covered where applicable
- [ ] Win32 and Win64 declaration impact is assessed
- [ ] Fault-injection seams remain deterministic and reset after use
- [ ] New mandatory cases are registered in the certification inventory
- [ ] Case accounting fails visibly on omissions or skips
- [ ] Tests preserve caller-owned snapshots and host state
- [ ] No regression change is required — rationale below

**New or amended case names:**

~~~text
N/A — explain, or list exact procedure names
~~~

**Coverage rationale:**



---

## ⚠️ Risk, rollback, and recovery

### Risk level

- [ ] Low — documentation, metadata, or mechanically verified change
- [ ] Medium — localized runtime or tooling behavior with bounded impact
- [ ] High — native frame, wrong-target, shared snapshot, cleanup, public API,
      release runner, or supply-chain impact

### Principal risks

<!--
What can fail despite normal-path tests? Consider active-window changes,
Window/hWnd mismatches, handle reuse, partial native writes, captionless
baselines, snapshot replacement, ScreenUpdating leakage, conditional
compilation, source/evidence drift, and binary identity.
-->

- <!-- principal risk -->

### Rollback

<!-- State the exact code, artifact, or release rollback. -->

- <!-- rollback step -->

### Excel recovery

<!--
State the recovery required if Excel is left constrained or identity becomes
uncertain. Do not rely only on snapshot restoration or normal cleanup after a
hard VBA reset.
-->

- [ ] UI_ShowExcelUI remains available as the deterministic show-all path
- [ ] Caller-owned snapshot state is not blindly cleared
- [ ] Full Excel-process restart was considered for uncertain native state
- [ ] N/A — no runtime or host-state impact

---

## 📋 Source and repository hygiene

- [ ] The branch and exact PR head were confirmed before evidence was recorded
- [ ] All four production modules were present during compilation
- [ ] Changed VBA modules were exported to the correct repository paths
- [ ] Repository line-ending and encoding policy was preserved
- [ ] No conflict markers, duplicate procedures, temporary comments, or
      placeholder implementation remain
- [ ] The complete textual diff contains only intended changes
- [ ] No Office lock file, generated workbook, backup, log, local evidence,
      credential, client data, private review, or confidential file is included
- [ ] Generated XLSM, XLAM, and certification outputs remain outside the Git
      source tree unless an explicit reviewed policy exception applies
- [ ] Comments and procedure headers describe current behavior rather than
      intended future behavior

---

## 🧱 Module ownership and dependencies

<details>
<summary>Expand when module boundaries, shared state, or visibility changes</summary>

<br>

~~~text
M_EXCEL_UI:
M_EXCEL_UI_RUNTIME:
M_EXCEL_UI_SNAPSHOT:
M_EXCEL_UI_TITLEBAR:
Dependency graph changed:
Circular dependency introduced:   no / explain
Mutable state duplicated:         no / explain
~~~

- [ ] M_EXCEL_UI remains the supported public facade
- [ ] Internal modules retain Option Private Module
- [ ] Snapshot state remains owned by M_EXCEL_UI_SNAPSHOT
- [ ] Title-bar registry state remains owned by M_EXCEL_UI_TITLEBAR
- [ ] Runtime services do not acquire inappropriate mutable subsystem state
- [ ] No circular dependency was introduced
- [ ] Any new internal or test seam has an actual caller and documented scope
- [ ] Production source remains a coherent four-module import unit

</details>

## 📸 Snapshot, identity, and recovery

<details>
<summary>Expand when capture, restoration, lifecycle, or window identity changes</summary>

<br>

~~~text
Captured state:
Snapshot owner:
Window identity strategy:
Window/hWnd pairing:
Behavior for new windows:
Behavior for missing, closed, or recreated windows:
Behavior for recycled hWnd:
Behavior after VBA reset:
Failure ordering:
Emergency recovery:
~~~

- [ ] Per-window restoration does not depend on collection index
- [ ] Excel Window object and Window.hWnd evidence are paired deliberately
- [ ] Application.Hwnd is not treated as stable identity across activation
- [ ] IsWindow is not treated as proof of original-window identity
- [ ] A mismatched or unprovable target fails closed before mutation
- [ ] Recycled native handles cannot inherit unrelated registry ownership
- [ ] New windows remain unchanged during restore
- [ ] Missing or recreated windows produce controlled diagnostics
- [ ] Capture replacement and clear release retained Window references
- [ ] Self-tests preserve a snapshot owned by their caller
- [ ] Reset-without-snapshot remains controlled
- [ ] UI_ShowExcelUI remains independent of snapshot availability

</details>

## 🪟 Ribbon and native title-bar behavior

<details>
<summary>Expand when Ribbon commands, activation, WinAPI, or frame state changes</summary>

<br>

~~~text
Excel 4 command or WinAPI used:
Caller-controlled command text:
Owned style bits:
32-bit path:
64-bit path:
Valid-zero / GetLastError treatment:
Frame refresh:
Target identity:
Activation side effect:
Unrelated style bits preserved:
~~~

- [ ] Ribbon command text remains fixed and non-injectable
- [ ] Active-window Ribbon behavior is explicit and tested
- [ ] Wrong-target restoration refuses mutation or follows the approved
      versioned policy
- [ ] Automatic activation is not introduced into a corrective patch unless
      explicitly approved for that release
- [ ] Exact owned style mask is preserved or deliberately reviewed
- [ ] Unrelated current style bits are preserved
- [ ] Win32 and Win64 declarations and types are correct
- [ ] Valid zero returns are distinguished from API failures
- [ ] Native handle liveness and identity are treated separately
- [ ] A successful style write followed by refresh failure remains recoverable
- [ ] Required frame refresh is performed and refresh debt is not lost
- [ ] Captionless non-zero baselines are read back or otherwise verified
- [ ] Multi-window behavior was tested with more than one workbook window

</details>

## 🧾 Diagnostics, failure policy, and cleanup

<details>
<summary>Expand when result contracts, error paths, quiet mode, or certification changes</summary>

<br>

~~~text
Failure contract:
FailureCount authority:
FailureList behavior:
Logging contract:
ScreenUpdating baseline:
Cleanup proof:
Certification completeness:
~~~

- [ ] FailureCount remains authoritative
- [ ] FailureList remains best effort and ordered
- [ ] Anything reachable from an error handler cannot destroy the original
      failure evidence
- [ ] Outputs that cannot fail are set before allocations or formatting
- [ ] Best-effort continuation is deliberate and documented
- [ ] No unsolicited production MsgBox was introduced
- [ ] ScreenUpdating is restored to the observed caller baseline
- [ ] Cleanup verifies snapshot, workbook/window, and quiet-update state as
      applicable
- [ ] Mandatory case omissions cannot produce a passing release verdict
- [ ] Evidence states PASS, FAIL, or INCOMPLETE unambiguously

</details>

---

## 📚 Documentation and release hygiene

- [ ] CHANGELOG.md updated under Unreleased for material behavior or governance
- [ ] Previously released changelog sections were not rewritten
- [ ] README, INSTALLATION, SECURITY, CONTRIBUTING, Wiki, and demo guidance were
      updated where affected
- [ ] Procedure headers, examples, and public issue acceptance criteria agree
      with the implementation
- [ ] Version stamps remain unchanged unless this is the deliberate
      release-stamp commit
- [ ] Public API, required modules, tests, static checks, and runtime evidence
      are not manually overstated
- [ ] Current released limitations are not documented as already fixed
- [ ] v1.2.0 Ribbon activation or demo modernization was not pulled into a
      v1.1.3 correctness fix
- [ ] Private review material is not linked, quoted, committed, or exposed
- [ ] No documentation change is required — rationale below

**Documentation rationale or Wiki follow-up:**



### Release or artifact packaging

<!-- Complete for a release PR, demo asset, checksum, manifest, or evidence change. -->

- [ ] The candidate tag and full SHA are explicit
- [ ] The exact reviewed PR head is the intended tagged commit
- [ ] Static and runtime evidence refer to that exact source
- [ ] Evidence invalidated by later changes was rerun
- [ ] The demo workbook was built from the claimed source and compiled
- [ ] The exact release asset was smoke-tested
- [ ] Published hashes identify the actual attached files
- [ ] Source identity, runtime certification, binary identity, and
      source-to-binary provenance are stated as separate claims
- [ ] Release notes close only issues whose implementation, tests,
      documentation, CI, and evidence were verified
- [ ] N/A — this PR does not prepare or publish a release artifact

---

## 👀 Reviewer focus

<!--
Point reviewers to the decisions most likely to be wrong despite green checks.
Name files, procedures, invariants, and important diff sections.
-->

1. <!-- first review focus -->
2. <!-- second review focus -->
3. <!-- third review focus -->

### Unresolved questions or accepted trade-offs

<!-- Write None when there are none. Link every deferred release-relevant item. -->



---

## ✅ Final author check

- [ ] The PR title describes the observable change
- [ ] The linked issue title, body, comments, labels, milestone, and acceptance
      criteria remain accurate
- [ ] Every closing keyword refers to a fully satisfied issue
- [ ] P1 and P2 release blockers remain open unless final evidence is complete
- [ ] The evidence above belongs to the exact SHA claimed
- [ ] Debug → Compile VBAProject succeeded when executable VBA changed
- [ ] Runtime certification was performed when required
- [ ] Static checks and git diff --check pass
- [ ] No merge markers, placeholder evidence, unexplained N/A, or unrelated
      edits remain
- [ ] I reviewed the full diff, including comments, documentation, tests, CI,
      and release claims
- [ ] The PR is ready to merge as presented; no unrecorded follow-up is required
