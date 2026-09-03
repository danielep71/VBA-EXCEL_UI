<!--
  Keep this pull request focused on one coherent outcome.
  Complete every common section. Delete optional profile blocks that do not apply.
  Use NOT RUN or NOT APPLICABLE with a reason; never manufacture PASS evidence.
  Record only checks and environments exercised against the exact candidate.
  Report vulnerabilities privately through SECURITY.md; do not disclose secrets,
  exploitable details, confidential workbooks, or restricted data in a pull request.
-->

<div align="center">

# 🔀 VBA Excel UI Pull Request

### Window identity · Host-state ownership · Recovery · Exact evidence

[![Contributing](https://img.shields.io/badge/guide-CONTRIBUTING-217346?style=flat-square)](../CONTRIBUTING.md)
[![Security](https://img.shields.io/badge/security-private%20reporting-d73a49?style=flat-square)](../SECURITY.md)
[![Release](https://img.shields.io/badge/release-RELEASING-6f42c1?style=flat-square)](../RELEASING.md)
[![Changelog](https://img.shields.io/badge/changes-Unreleased-d97706?style=flat-square)](../CHANGELOG.md)

</div>

---

> [!IMPORTANT]
> Static checks cannot prove window identity, title-bar behavior, or restoration. Executable changes require interactive Windows Excel evidence from the exact candidate.

## 📌 Summary

<!-- State the observable outcome and why it is needed. Prefer one precise purpose. -->

## 🔗 Related issues

```text
Closes #
Related to #
```

Use a closing keyword only when this pull request satisfies the issue's complete acceptance criteria.

## 🧭 Change classification

- [ ] Defect correction
- [ ] Backward-compatible capability
- [ ] Breaking API, behavior, deployment, or migration change
- [ ] Internal refactor with no intended supported-behavior change
- [ ] Test, fixture, reference-data, or validation change
- [ ] Performance change
- [ ] Security or trust-boundary hardening
- [ ] Documentation-only change
- [ ] Repository tooling, workflow, or governance change
- [ ] Packaging or release preparation
- [ ] Excel-state ownership or recovery change
- [ ] Window identity, title-bar, Ribbon, or WinAPI change
- [ ] Snapshot, diagnostic, or structured-result change

## 🎚️ Affected surface

- [ ] Public `UI_*` facade
- [ ] Runtime and application-level UI state
- [ ] Snapshot capture and restoration
- [ ] Title-bar and native-window integration
- [ ] Multi-window targeting and identity
- [ ] Diagnostics and recovery
- [ ] No runtime or supported surface — documentation/repository-only

---

## 📐 Scope and contract impact

### In scope

- <!-- Deliberate outcome -->

### Out of scope

- <!-- Reasonable adjacent work deliberately deferred -->

### Supported behavior and compatibility

```text
Supported behavior changed:       Yes / No
Backward compatible:              Yes / No / Uncertain
Suggested release impact:         none / patch / minor / major / uncertain
New supported members:
Removed or renamed members:
Changed signatures or defaults:
Changed results, errors, state, or side effects:
Migration required:
Known limitation introduced or retained:
```

Assess compatibility against documented behavior, not merely the VBA `Public` keyword. Infrastructure callbacks, Ribbon entry points, test seams, and `Application.Run` targets are not automatically supported API.

### Production source and package

All four production modules: `M_EXCEL_UI`, `M_EXCEL_UI_RUNTIME`, `M_EXCEL_UI_SNAPSHOT`, and `M_EXCEL_UI_TITLEBAR`.

- [ ] Required source files and import order are unchanged.
- [ ] Required source files or order changed and `INSTALLATION.md` was updated.
- [ ] No production source/package impact.

## 🔧 Implementation notes

```text
Approach and key invariant:
Alternatives considered:
New dependency, reference, or generated input:
State ownership and cleanup:
Failure behavior:
```

Explain decisions a future reviewer cannot safely infer from the diff.

---

## ✅ Verification

### Candidate identity

| Evidence | Result |
| --- | --- |
| Exact PR HEAD SHA | <!-- Full 40-character SHA --> |
| Base branch and base SHA | <!-- Branch + full SHA --> |
| Working tree used locally | <!-- clean / dirty; explain --> |
| Source or package tested | <!-- Exact candidate source / artifact / N/A --> |

Evidence from another commit does not certify this candidate.

### Static and repository checks

- `python3 tools/check_repo.py`
- `python3 tools/reformat.py --check src/*.bas test/*.bas demo/*.bas`
- `git diff --check`

| Check | Result / evidence |
| --- | --- |
| Hosted required checks | <!-- PASS / FAIL / NOT RUN + workflow URL --> |
| Local static command | <!-- Command + PASS / FAIL / NOT RUN --> |
| Formatting / `git diff --check` | <!-- PASS / FAIL --> |
| Machine-readable artifact | <!-- Name / URL / not produced --> |

### Excel and VBA execution

- [ ] Required and completed against the exact PR HEAD.
- [ ] Required but incomplete — reason and merge/release consequence stated.
- [ ] Not required — documentation/repository-only change with no executable or packaging impact.

Relevant entry points:

- `Test_EXCEL_UI_RunReleaseCertification`
- Relevant focused `Test_EXCEL_UI_*` runners
- Manual `UI_HideExcelUI` / `UI_ShowExcelUI` and capture/hide/reset scenarios

| Evidence | Result |
| --- | --- |
| Tested commit SHA | <!-- Full SHA or N/A --> |
| `Debug → Compile VBAProject` | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Regression/certification entry point | <!-- Exact procedure --> |
| Completion state | <!-- PASS / FAIL / INCOMPLETE / NOT RUN --> |
| Cases / assertions / failures | <!-- Counts or N/A --> |
| Skipped / cleanup outcome | <!-- Counts and state or N/A --> |
| Focused and manual checks | <!-- Scenarios + result --> |
| Evidence file or workflow | <!-- Name / URL / N/A --> |

### Validation environment

```text
Excel product, version, and build:
Office bitness:                    32-bit / 64-bit
Windows version/build:
Workbook or add-in host:
Deployment model:
Workbook type and Excel window state
Number of open Excel windows
Other add-ins
Title-bar and Ribbon configuration
```

Record only tested environments. Source inspection does not constitute host execution, and one Office bitness does not execute the other conditional branch.

### Regression coverage

- [ ] Existing tests cover the changed success path.
- [ ] New or amended tests cover each corrected defect.
- [ ] Boundary, invalid-input, failure, fallback, and cleanup paths are covered as applicable.
- [ ] Test entry points and inventory/count metadata remain synchronized.
- [ ] Expected results come from the contract or an independent reference.
- [ ] No regression change is needed — rationale recorded below.

```text
Coverage rationale and new test names:
Unexecuted or deferred coverage:
```

---

## ⚠️ Risk, rollback, and recovery

- [ ] Low — documentation, metadata, or mechanically verified change.
- [ ] Medium — bounded runtime, tooling, or compatibility impact.
- [ ] High — numerical integrity, shared Excel state, native API, security, release, or breaking impact.

```text
Principal failure modes:
Residual risk after validation:
Rollback or revert procedure:
Excel-process, workbook, data, or artifact recovery:
Conditions that make rollback unsafe:
```

## 🔐 Security, data, and provenance

- [ ] No credential, secret, signing material, internal URL, or personal path is included.
- [ ] No client, employer, counterparty, student, personal, or restricted production data is included.
- [ ] Test data is synthetic, anonymized, or explicitly redistributable.
- [ ] External algorithms, code, datasets, and market/vendor data have attributable provenance and compatible licensing.
- [ ] Formula, command, path, callback, deserialization, and external-content injection surfaces were assessed.
- [ ] No security-sensitive detail belongs in private disclosure instead of this pull request.
- [ ] Generated evidence identifies its inputs, tool/runtime version, candidate SHA, and limitations.

```text
Security or privacy impact:
Source/data provenance:
New trust boundary:
```

## 📚 Documentation and release hygiene

- [ ] `README.md` reflects supported behavior and examples.
- [ ] `INSTALLATION.md` reflects paths, dependencies, import order, validation, upgrades, and removal.
- [ ] `CONTRIBUTING.md` reflects development and evidence requirements.
- [ ] `CHANGELOG.md` records material change under `[Unreleased]`.
- [ ] `SECURITY.md` reflects supported versions or trust boundaries.
- [ ] `RELEASING.md` reflects certification, package, provenance, or recovery changes.
- [ ] Source headers, API references, demos, Wiki pages, and counts remain synchronized.
- [ ] Version markers remain unchanged unless this is the deliberate release-stamp change.
- [ ] No documentation change is required — reason recorded below.

```text
Documentation impact:
Release, artifact, or migration impact:
```

---

## 🧩 Project-specific review

<details>
<summary><strong>🧱 Module ownership and API manifest</strong></summary>

Keep when module boundaries or supported surface can change.

- [ ] `M_EXCEL_UI` remains the public facade.
- [ ] Runtime and title-bar modules retain their allowed dependency boundary.
- [ ] Snapshot state exists only in `M_EXCEL_UI_SNAPSHOT`.
- [ ] Title-bar mutable state exists only in `M_EXCEL_UI_TITLEBAR`.
- [ ] No circular dependency or duplicated mutable state is introduced.
- [ ] Internal modules retain `Option Private Module`.
- [ ] `tools/public_api_manifest.txt` intentionally reflects every public-surface change.

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

</details>
<details>
<summary><strong>📸 Snapshot, identity, and recovery</strong></summary>

Keep when capture, restoration, targeting, or reset behavior changes.

- [ ] Per-window restoration never relies on collection index.
- [ ] Object identity and recycled-window-handle risks are addressed.
- [ ] New, missing, closed, and recreated windows have explicit behavior.
- [ ] Reset without a valid snapshot is controlled.
- [ ] State is never applied to a window that cannot be proven to be the captured target.
- [ ] `UI_ShowExcelUI` remains an emergency show-all recovery route.

</details>
<details>
<summary><strong>🪟 Ribbon and WinAPI</strong></summary>

Keep for Ribbon, title-bar, or native-frame changes.

- [ ] Only owned style bits are changed and unrelated bits are preserved.
- [ ] 32-bit and 64-bit declarations are correct.
- [ ] Valid zero returns are distinguished from native failures.
- [ ] Required frame refresh is performed and failures remain observable.
- [ ] Multi-window behavior is tested with more than one workbook open.

</details>
<details>
<summary><strong>🧾 Diagnostics and cleanup</strong></summary>

Keep when failure ordering, results, or Excel state can change.

- [ ] Structured diagnostics preserve insertion order and real post-operation state.
- [ ] Error handlers cannot replace the original failure.
- [ ] `ScreenUpdating` and other caller-owned state are restored.
- [ ] Best-effort continuation is deliberate and failures are not discarded.
- [ ] Certification reports complete, failed, skipped, and cleanup counters.

</details>

---

## 👀 Reviewer focus

```text
Highest-risk decision:
Files and procedures to inspect first:
Evidence to challenge:
Known boundary not proved by this pull request:
Unresolved question or accepted trade-off:
```

## ☑️ Final author check

- [ ] The title describes the observable outcome.
- [ ] The pull request has one coherent purpose and no unrelated churn.
- [ ] Linked issue acceptance criteria are met or remaining work is explicit.
- [ ] Compatibility and release impact are assessed.
- [ ] Evidence belongs to the exact candidate claimed.
- [ ] Required checks are terminal and passing; incomplete work is not presented as PASS.
- [ ] Executable VBA was compiled and tested when required.
- [ ] Failure, cleanup, and recovery behavior were reviewed.
- [ ] The complete diff, including comments, metadata, binary companions, and documentation, was reviewed.
- [ ] No merge marker, stale placeholder, unexplained N/A, accidental binary, or private material remains.

---

**Review principle:** approve the smallest coherent change whose contract, evidence, risk, and recovery can all be explained from this pull request.
