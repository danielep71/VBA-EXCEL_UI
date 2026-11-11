<!--
  Keep this pull request focused on one coherent outcome.
  Complete every section; the reference lists are in CONTRIBUTING.md.
  Use NOT RUN or NOT APPLICABLE with a reason; never manufacture PASS evidence.
  Record only checks and environments exercised against the exact candidate.
  Report vulnerabilities privately through SECURITY.md; do not disclose secrets,
  exploitable details, confidential workbooks, or restricted data in a pull request.
-->

# 🔀 VBA Excel UI Pull Request

Guides: [CONTRIBUTING](../CONTRIBUTING.md) · [SECURITY](../SECURITY.md) · [RELEASING](../RELEASING.md) · [CHANGELOG](../CHANGELOG.md)

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

## 🧭 Classification and affected surface

```text
Classification:
Affected surface:
```

Pick from the [classification and surface lists](../CONTRIBUTING.md#pull-request-reference) in CONTRIBUTING.md; name every one that applies.

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
Production modules or import order:  unchanged / changed (INSTALLATION.md updated) / N/A
```

Assess compatibility against documented behavior, not merely the VBA `Public` keyword. Infrastructure callbacks, Ribbon entry points, test seams, and `Application.Run` targets are not automatically supported API.

## 🔧 Implementation notes

```text
Approach and key invariant:
Alternatives considered:
New dependency, reference, or generated input:
State ownership and cleanup:
Failure behavior:
```

Explain decisions a future reviewer cannot safely infer from the diff.

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

| Check | Result / evidence |
| --- | --- |
| Hosted required checks | <!-- PASS / FAIL / NOT RUN + workflow URL --> |
| Local static command | <!-- python3 tools/verify.py --ci — its RESULT line, or NOT RUN --> |
| Formatting / `git diff --check` | <!-- PASS / FAIL --> |
| Machine-readable artifact | <!-- Name / URL / not produced --> |

### Excel and VBA execution

- [ ] Required and completed against the exact PR HEAD.
- [ ] Required but incomplete — reason and merge/release consequence stated.
- [ ] Not required — documentation/repository-only change with no executable or packaging impact.

| Evidence | Result |
| --- | --- |
| Tested commit SHA | <!-- Full SHA or N/A --> |
| `Debug → Compile VBAProject` | <!-- PASS / FAIL / NOT RUN / N/A --> |
| Regression/certification entry point | <!-- Test_EXCEL_UI_RunReleaseCertification, or the focused Test_EXCEL_UI_* runner --> |
| Completion state | <!-- PASS / FAIL / INCOMPLETE / NOT RUN --> |
| Cases / assertions / failures | <!-- Counts or N/A --> |
| Skipped / cleanup outcome | <!-- Counts and state or N/A --> |
| Focused and manual checks | <!-- UI_HideExcelUI / UI_ShowExcelUI, capture/hide/reset scenarios + result --> |
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

See the [regression coverage list](../CONTRIBUTING.md#regression-coverage) in CONTRIBUTING.md.

```text
Coverage rationale and new test names:
Unexecuted or deferred coverage:
```

## ⚠️ Risk, rollback, and recovery

```text
Risk level:                    Low / Medium / High (CONTRIBUTING.md → Risk levels)
Principal failure modes:
Residual risk after validation:
Rollback or revert procedure:
Excel-process, workbook, data, or artifact recovery:
Conditions that make rollback unsafe:
```

## 🔐 Security, data, and provenance

- [ ] Every item of the [security and provenance list](../CONTRIBUTING.md#security-and-provenance-review) in CONTRIBUTING.md holds, or the exception is stated below.

```text
Security or privacy impact:
Source/data provenance:
New trust boundary:
```

## 📚 Documentation and project-specific review

```text
Documents updated:
Release, artifact, or migration impact:
Review profiles applied:   module ownership / snapshots / Ribbon and WinAPI / diagnostics / none
Profile items not met:
```

Check the [documentation list](../CONTRIBUTING.md#documentation-and-release-hygiene) and apply each [review profile](../CONTRIBUTING.md#project-specific-review-profiles) your change touches; state "None" with a reason when no document changes.

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
