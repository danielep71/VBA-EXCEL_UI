<div align="center">

# 🤝 Contributing to VBA Excel UI

### Reviewable changes to a structured Windows Excel UI controller

[![Contributions](https://img.shields.io/badge/Contributions-Welcome-2ea44f?style=flat-square)](#ways-to-contribute)
[![Conduct](https://img.shields.io/badge/Conduct-Required-6f42c1?style=flat-square)](CODE_OF_CONDUCT.md)
[![Security](https://img.shields.io/badge/Security-Private_reporting-d73a49?style=flat-square)](SECURITY.md)
[![Workflow](https://img.shields.io/badge/Workflow-Source--first-0969da?style=flat-square)](#source-first-vba)
[![License](https://img.shields.io/badge/License-MIT-217346?style=flat-square)](LICENSE)

<br>

**Focused scope · Reviewable source · Reproducible evidence · Honest limitations**

<br>

[Start here](#start-here)
&nbsp;·&nbsp;
[Workflow](#development-workflow)
&nbsp;·&nbsp;
[VBA rules](#source-first-vba)
&nbsp;·&nbsp;
[Validation](#validation-and-evidence)
&nbsp;·&nbsp;
[Pull requests](#pull-requests)

</div>

---

Thank you for helping improve **VBA Excel UI**.

Contributions are welcome when they strengthen correctness, clarity,
maintainability, compatibility, documentation, tests, or reproducibility. The
standard is not simply that a change works once: another person must be able to
review it, reproduce the evidence, and understand its operational boundaries.

Participation is governed by the [Code of Conduct](CODE_OF_CONDUCT.md).
Report suspected vulnerabilities privately under [SECURITY.md](SECURITY.md);
never disclose sensitive details in a public issue or pull request.

---

<a id="start-here"></a>

## 🧭 Start here

Before opening work:

1. Read the README, this guide, the [Code of Conduct](CODE_OF_CONDUCT.md),
   [Security Policy](SECURITY.md), and [CHANGELOG.md](CHANGELOG.md).
2. Search open and closed issues and pull requests for related work.
3. Open an issue before a non-trivial feature, public-API change, dependency,
   architectural change, compatibility break, or broad refactor.
4. Agree the observable contract and validation approach before implementation.
5. Keep credentials, personal/client data, proprietary workbooks, and restricted
   reference material out of the repository.

Small documentation corrections and narrowly obvious fixes may go directly to a
focused pull request.

> [!IMPORTANT]
> UI restoration is an ownership problem. A stale snapshot or window handle must
> never be treated as authority to overwrite unrelated current Excel state.

---

<a id="ways-to-contribute"></a>

## 🌱 Ways to contribute

| Contribution | Good first action |
|---|---|
| 🐛 Reproducible defect | Open an issue with minimal inputs, expected behavior, observed behavior, and environment. |
| ✨ Feature or API change | Open an issue describing users, contract, alternatives, compatibility, and validation. |
| 🧪 Tests or reference evidence | Explain provenance, independence, precision, coverage, and expected failure detection. |
| 📖 Documentation | Identify the affected behavior and keep examples executable and current. |
| ⚙️ Repository/tooling | Explain developer impact, failure behavior, portability, and maintenance cost. |
| 🔐 Security concern | Follow [SECURITY.md](SECURITY.md); do not open a public report. |
| 💬 Usage question | Use the repository's supported discussion or issue channel without sensitive data. |

A proposal may be adapted, deferred, or declined when it is out of scope,
duplicates an existing capability, weakens a contract, or creates maintenance
cost disproportionate to its benefit.

---

## 📁 Repository model

This is a **source-first Excel UI library**. The Git diff, not an opaque workbook, is
the review artifact.

| Location | Purpose |
|---|---|
| `src/` | Authoritative production modules for UI, runtime, snapshots, and title-bar control |
| `test/` | Regression and release-certification source |
| `demo/` | Demo and reproducible host setup |
| `tools/` | Repository/static validation |
| `docs/` | Architecture, API, recovery, and release guidance |

The README and current tree are authoritative if a listed optional directory is
not present.

---

<a id="development-workflow"></a>

## 🌿 Development workflow

1. Fork or clone the repository and start from the current `main`.
2. Create a short, focused branch such as `fix/clear-description`,
   `feat/clear-description`, `docs/clear-description`, or
   `test/clear-description`.
3. Reproduce the existing behavior before changing it.
4. Define the intended contract, affected callers, compatibility impact, and
   evidence plan.
5. Make the smallest coherent source change; do not mix unrelated formatting,
   refactoring, generated files, or cleanup.
6. Compile and run the relevant static, regression, host, and manual checks.
7. Re-export changed VBA components and review the complete text/binary diff.
8. Update documentation and release notes required by the change.
9. Push the branch and open a pull request with evidence and limitations.

Repository maintainers may use the repository's configured direct-push workflow
where permitted. External contributions and reviewable portfolio changes should
use branches and pull requests.

### Commit discipline

Write imperative, specific subjects, normally in this form:

```text
fix: preserve formulas during write-back
feat: add explicit tail calculation
test: cover cleanup after initialization failure
docs: clarify supported Office environments
chore: harden repository validation
```

Keep commits reviewable. Reference the issue when one exists. Do not include
secrets, private links, generated attribution boilerplate, or unverifiable test
claims in commit messages.

---

<a id="source-first-vba"></a>

## 📦 Source-first VBA

Exported source is authoritative.

- Use `Option Explicit`.
- Preserve the repository's VBE export metadata, module names, encoding, and
  line-ending policy.
- Match `.bas`, `.cls`, and `.frm` filenames to their component identity.
- Keep every required `.frm` / `.frx` pair together; treat `.frx` as binary.
- Do not edit a binary form resource as text.
- Do not use a workbook or add-in as the only record of a code change.
- Do not commit Office lock files, recovery copies, local exports, test output,
  or generated binaries unless the repository explicitly designates them as
  source.
- Qualify workbook, worksheet, range, and application references.
- Avoid implicit active-workbook, active-sheet, selection, and default-member
  dependencies.
- Keep `On Error Resume Next` scopes narrow and intentional.
- Preserve useful diagnostic context and clean up on success and failure.
- Avoid new references, APIs, dependencies, or platform assumptions until their
  support and deployment impact is agreed.

### Public contracts and compatibility

Treat documented procedures, functions, classes, enums, parameters, defaults,
return values, errors, side effects, workbook formats, and supported platforms
as contracts.

A contract-changing contribution must:

1. identify affected callers and migration needs;
2. explain what changes and what remains unchanged;
3. add or update regression coverage;
4. update user-facing documentation and examples; and
5. state whether the release impact is patch, minor, or major.

Do not make an internal helper public merely to simplify a test. Use an explicit
test seam where the project supports one.

### Excel state ownership

Assume these surfaces belong to the caller or host unless the project explicitly
owns them:

```text
Application.Calculation
Application.EnableEvents
Application.ScreenUpdating
Application.DisplayAlerts
Application.StatusBar
active workbook / worksheet / selection
window styles, shortcuts, timers, names, links, connections, and shapes
```

Capture state before changing it. Restore only state the component successfully
changed and still owns. Cleanup must not conceal the original failure.

---

## 🧩 Project engineering contract

| Area | Required behavior |
|---|---|
| **Module boundaries** | Preserve the UI, runtime, snapshot, and title-bar dependency direction. |
| **Window identity** | Never restore per-window state by collection index; handle closed, recreated, and newly opened windows explicitly. |
| **Snapshot lifecycle** | Document partial capture, failure ordering, reset behavior, and emergency recovery. |
| **Title-bar ownership** | Restore only the style bits owned; handle `GetLastError`, valid zero returns, handle changes, and frame refresh. |
| **Failure behavior** | Production entry points are fail-soft unless documented otherwise; never silently discard failures or leave screen updating suppressed. |
| **Platform support** | Validate affected 32-bit and 64-bit declarations on a real Windows Excel host. |

---

<a id="validation-and-evidence"></a>

## 🧪 Validation and evidence

Validation must be proportional to risk and reproducible from the exact source
under review.

- Compile the complete VBA project.
- Run `Test_EXCEL_UI_RunReleaseCertification` for production changes.
- Exercise `UI_HideExcelUI` / `UI_ShowExcelUI` and capture / hide / reset flows.
- Treat any `INCOMPLETE`, skipped unit, failed unit, or cleanup failure as not passing.
- Run hosted static checks and record their result.
- Record exact commit, Excel/Windows/Office bitness, multi-window scope, verdict, skips, and cleanup.

Hosted CI can validate source and repository invariants but cannot certify VBA behavior inside Excel. Quote the real-host certification verdict in the pull request and state any environment not exercised.

### Evidence principles

- Test the behavior, not only the implementation path.
- Add a permanent regression for every corrected defect.
- Include ordinary, boundary, invalid-input, error, and cleanup paths.
- Use an independent source for expected numerical results.
- State skips and unavailable environments explicitly; a skipped check is not a
  pass.
- Do not claim compatibility, accuracy, performance, or certification beyond
  what was actually observed.
- Treat cleanup failures and incomplete runs as failures.
- Never generate expected values with the implementation under test.

### Suggested evidence block

```text
Source
------
Commit / tag:
Files or components changed:

Environment
-----------
Excel:
Office bitness:
Operating system:
Locale / date system:
Deployment or host:

Checks
------
Compile:
Static checks:
Focused tests:
Full regression:
Manual / UI / platform checks:
Cleanup:

Evidence
--------
Independent reference and version:
Inputs / workload:
Tolerance or acceptance rule:
Expected:
Observed:
Worst discrepancy / dispersion:

Limitations
-----------
Skipped or unverified:
Follow-up:
```

Remove non-applicable fields, but do not omit a material limitation.

---

## 📖 Documentation and release notes

Update the README, API and architecture material, recovery instructions, demo guidance, and the `[Unreleased]` section of `CHANGELOG.md` when behavior changes.

Documentation must say:

- what users can rely on;
- inputs, outputs, defaults, side effects, and failure behavior;
- supported and untested environments;
- installation or migration steps;
- numerical or platform assumptions; and
- any known limitation introduced or exposed by the change.

Do not edit a released version or tag merely to describe unreleased work. Release
numbers, artifacts, hashes, and dates belong to the repository's release
workflow.

---

## 🔐 Security, privacy, and provenance

- Follow [SECURITY.md](SECURITY.md) for vulnerability reports.
- Use synthetic, anonymized, or explicitly redistributable examples and data.
- Remove names, email addresses, account identifiers, workbook properties,
  document metadata, credentials, tokens, private URLs, and machine-specific
  paths.
- Verify the license and redistribution rights of copied code, formulas,
  reference tables, images, and generated material.
- Cite material algorithms and external reference data precisely enough for a
  reviewer to verify them.
- You remain responsible for the correctness, licensing, security, and
  reviewability of tool-assisted contributions.

---

<a id="pull-requests"></a>

## 🚀 Pull requests

A pull request should answer:

```text
What problem does this solve?
What observable contract changes?
What remains compatible?
How was it validated from this exact source?
What evidence is independent?
What remains unverified?
```

### Checklist

```text
[ ] Scope is focused and the related issue is linked
[ ] Public API, compatibility, and release impact are assessed
[ ] Exported VBA source and required binary companions are synchronized
[ ] Relevant compile, static, regression, and manual checks are recorded
[ ] Numerical/performance evidence is independent and reproducible where relevant
[ ] Error, boundary, recovery, and cleanup paths are covered
[ ] Caller-owned Excel state and platform/bitness concerns are addressed
[ ] README, contracts, examples, and release notes are updated
[ ] No confidential, restricted, generated, or accidental binary content is added
[ ] Unverified environments and skipped checks are stated plainly
[ ] Final diff contains no unrelated formatting or local artifacts
```

Reviews may request changes to scope, tests, contracts, compatibility,
documentation, or evidence. Discussion must remain technical and respectful
under the [Code of Conduct](CODE_OF_CONDUCT.md).

---

## 🤝 Review and maintainer decisions

Reviewers evaluate correctness, safety, maintainability, compatibility,
evidence, documentation, and fit with the project's direction. Approval of an
idea does not guarantee acceptance of every implementation detail.

The maintainer may edit, squash, defer, or decline a contribution to protect the
coherence and supportability of the project. Contributors will be credited
through Git history and release notes where appropriate.

---

## 📄 Licensing

By contributing, you agree that your contribution is licensed under the
repository's [MIT License](LICENSE). You must have the right to submit every
part of the contribution, including code, tests, data, images, and generated
material.

---

## 👤 Maintainer

Maintained by **Daniele Penza**.

For ordinary contributions, use GitHub issues and pull requests. For sensitive
security matters, use the private channel in [SECURITY.md](SECURITY.md).

---

### Contribution principle

> Make the contract explicit, keep the diff focused, and leave evidence another
> person can reproduce.
