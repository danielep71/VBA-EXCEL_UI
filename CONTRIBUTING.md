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

Portable text (Markdown, Python, YAML and JSON) uses LF. Exported VBA keeps its
existing CRLF policy. Windows-native `.bat`, `.cmd`, `.ps1`, `.psm1`, `.psd1`,
`.vbs`, `.reg` and `.ini` files use CRLF in editors and Git checkouts. This is a
repository consistency policy, not a claim that every host rejects LF.

### Disposable mutation controls

Create throwaway mutants and local control output only in the repository-root
`/.mutation-scratch/` directory (created locally on demand, never committed).
Do not place mutants beside `src/` or authoritative test fixtures. Do not add
broad `*.bas`, `*mutant*`, `*control*` or evidence-directory ignore rules.

The recorded v1.1.3 control inventory is:

| Issue | Disposable artifact type | Purpose |
| --- | --- | --- |
| #43 | Four one-runner-at-a-time refusal mutants | Prove caller-owned snapshots survive each refusal |
| #43 | Cleanup-skipping mutant | Prove the success-path snapshot-release assertion detects leakage |
| #6 / #66 | Title-bar module with the v1.1.2 fallback condition restored | Negative control for captionless recovery; currently hangs, not passing evidence |
| #45 | Active-frame-pair comparison control | Prove disagreement fails at `.Disagreed.Refused` |

This inventories recorded types and purposes, not recoverable local filenames:
the historical throwaway files were not committed. For #32, a true unchanged-
hWnd generation control remains outstanding; do not count the distinct-handle
seam as that evidence. Preserve reviewed results separately from disposable code.

### Exported-source rules

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

External compatibility is a promise about the `[supported]` facade in
`M_EXCEL_UI`. It is not the same promise as the deployment rule that all four
`src/` modules are replaced together: internal boundaries may move in a patch
release without any caller-visible change, which is why a mixed-version project
fails to compile even when the external API is untouched.

Preserve the `[supported]` facade unless an explicitly approved breaking release
requires otherwise:

- public procedure and function names;
- parameter order;
- optional defaults;
- enum values;
- show/hide/leave-unchanged semantics;
- fire-and-forget behavior;
- ordered structured-result behavior;
- application and window scope;
- `UI_ShowExcelUI` as the emergency show-all operation.

New backward-compatible parameters must normally be optional and trailing.

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

### Required validation

For production code changes:

```text
Debug → Compile VBAProject
Test_EXCEL_UI_RunReleaseCertification
UI_HideExcelUI / UI_ShowExcelUI
capture / hide / reset
```

Certification is the gate. Quote its verdict line in the pull request:

```text
RESULT: PASS | COMPLETE | units=3 failed=0 skipped=0 cleanup=OK
```

A run reporting `INCOMPLETE`, any `skipped` count above zero, or
`cleanup=FAILED` is not a pass, whatever the assertions that did execute
reported. The narrower runners remain useful while iterating, but do not
substitute for certification: `Test_EXCEL_UI_RunAll` executes no multi-window
case and produces no machine-readable evidence.

CI runs the static gate on every pull request. It cannot execute VBA — a hosted
runner has no Excel — so certification remains a manual step on a real host, and
the two are complementary rather than alternatives.

Record only environments actually tested.

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

### Three SHAs, three different claims

A release involves three commits, and evidence attached to one of them says
nothing about the others. Naming which is meant is not pedantry: the v1.1.1 and
v1.1.2 tags were published from commits carrying no automated evidence at all,
and the only way to notice that is to keep the three apart.

| SHA | What it is | What evidence on it proves |
|---|---|---|
| **PR head** | The last commit on `release/v1.1.3` before merge | The reviewed tree passed the static gate and the bot reviewed *this* tree. It does not prove anything about `main`. |
| **Merge SHA** | The merge commit created on `main` | `main` contains the reviewed tree. The merge itself is a new commit whose tree can differ from the PR head if anything was resolved during merge. |
| **Tag SHA** | What `v1.1.3` points at | The artefact people install from. It should equal the merge SHA, and until the tag trigger existed nothing checked that it had ever been built. |

Each has a distinct owner. The static gate runs on the PR head and, since the
`v*` trigger, on the tag SHA as well. Binding a certification run to a specific
tree and its hashes belongs to
[#46](https://github.com/danielep71/VBA-EXCEL_UI/issues/46). Proving the tag was
created from the reviewed merge commit belongs to
[#49](https://github.com/danielep71/VBA-EXCEL_UI/issues/49). Do not present
evidence gathered on one as evidence for another; a green check on the PR head
is not a certified tag.


## 📖 Documentation and release notes

Installation or packaging changes must keep [INSTALLATION.md](INSTALLATION.md) current. Release preparation must follow [RELEASING.md](RELEASING.md).

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

### The wiki and the VERSION file

The root `VERSION` file is the single release number this repository states.
Module headers and the wiki track badges both derive from it, and it moves once
per release rather than in several places independently.

The wiki is a separate Git-backed repository with no Actions surface of its own,
so it cannot check itself. `.github/workflows/wiki-badges.yml` clones it
read-only and runs `tools/wiki_badges.py`, which asserts that every content page
carries a `wiki_tracks-vX.Y.Z` badge, that every badge agrees, and that they
agree with `VERSION`. Reserved navigation pages — `_Sidebar.md` and friends —
are exempt by an explicit list rather than by a leading-underscore pattern,
which would silently exempt any future page named that way.

Two things about that arrangement are deliberate. The expected version comes
from this repository and not from the wiki, because deriving it from the first
page read would make the check circular: fourteen pages agreeing on a stale
release would pass. And the workflow runs on a schedule as well as on push,
because a wiki edit does not touch this repository and therefore starts nothing
— without the timer, drift introduced in the wiki editor stays invisible until
the next commit lands here.

`tools/check_repo.py` never touches the network. It runs the badge rules against
their own fixtures so the policy cannot rot, and leaves the clone to the
workflow.

Every run records what it inspected: the expected track, the wiki commit it read,
and an ordered inventory naming each page and the badge it carried. A passing run
that records nothing cannot be tied to anything later, and a wiki of fourteen
agreeing pages would be indistinguishable from a wiki of one.

#### What the badge means, and when the two legitimately disagree

`wiki_tracks-v1.1.3` means *this page was written against the v1.1.3 contract*.
It does not mean the tag exists. The badge is a release-candidate claim, and it
has to be, because the wiki is reviewed before the tag is cut rather than after
— a badge that could only be written post-tag would make the gate impossible to
satisfy at the moment it matters.

That produces one deliberate window per release, in wave 7:

1. `VERSION` moves to the new release, together with the module headers under
   [#36](https://github.com/danielep71/VBA-EXCEL_UI/issues/36).
2. The wiki gate goes red. Every page still claims the previous release, and
   the repository now states the new one. **This is the signal working**, not a
   fault.
3. The wiki pages are reviewed against the new contract and re-badged.
4. The gate goes green again, before the tree is frozen.

Order matters. Wiki edits do not touch this repository, so re-badging cannot
invalidate a frozen tree — but bumping `VERSION` can, because it is a tracked
file. Bump `VERSION` and re-badge the wiki *before* the release freeze, never
during it.

The wiki gates no merge. It makes a silent failure loud, which is the failure
that let the wiki fall a full release behind before v1.1.2.

### Markdown is edited as text

Never edit a Markdown file in a WYSIWYG editor. One escapes every character it
believes is markup, and the result is a document that renders as literal
backslashes: `## 1\. Executive assessment`, `M\_EXCEL\_UI`, `\---` where a
horizontal rule was meant. Both independent review archives arrived that way and
carried 669 escapes between them for two releases before anyone noticed.

`tools/check_repo.py` now fails on escaped Markdown punctuation in any tracked
document. The check is fence-aware and deliberately ignores fenced blocks and
inline code spans, because a backslash there is data — a Windows path, a regular
expression — and flagging it would make the gate wrong about correct documents.
An editor that escapes a fence escapes the prose around it too, so the rule
detects the editor rather than every symptom.


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

### Independent reviews are internal

From v1.1.3 onward this repository does not carry independent review documents.
They are internal working material: they drive the issue set, and the issue set
is the public record. `check_repo.py` denies any path matching
`*INDEPENDENT_CODE_REVIEW*`, in the Git index and in the working tree both — the
index catches the file after a commit, the working tree before one, which is the
only moment the mistake is cheap to undo.

The pattern covers the whole family rather than the documents withdrawn in
v1.1.3, because naming individual files would readmit the next review by the
simple fact of it having a new name.

This is a rule about what the project carries forward, not a claim that the
withdrawn documents were never public. The v1.1.0 and v1.1.1 archives shipped in
the `v1.1.1` and `v1.1.2` tags and remain reachable there and in this branch's
history. Rewriting that would invalidate published tag SHAs, including the exact
tree hashes the reviews themselves cite as what they reviewed.

Cite a finding by its identifier — `ICR-UI-*` for the v1.1.0 review,
`ICR-UI-111-*` for v1.1.1, `ICR-UI-112-*` for v1.1.2. Never quote or attach the
document. Issue bodies must be self-contained, because from v1.1.3 the
identifier is the only public handle a reader has, and
[#48](https://github.com/danielep71/VBA-EXCEL_UI/issues/48) publishes the
disposition table those identifiers resolve against.

### Updating a pinned Action

Every `uses:` reference in `.github/workflows/` is pinned to a full 40-character
commit SHA with a trailing comment naming the release it is, and
`tools/check_repo.py` fails on anything else. A version tag is a mutable
pointer: `actions/checkout@v4` moved to a different commit as recently as the
pin currently in the tree, so a workflow referencing it runs whatever the
upstream account publishes next, including whatever an attacker publishes after
compromising it.

The comment is not decoration. A bare 40-hex string cannot be reviewed — nobody
can tell a real release from an arbitrary commit by reading it. The comment
states the claim, and resolving the tag is what checks it.

To move a pin:

1. Find the intended release in the official Action repository and note its tag,
   for example `v4.4.0`. Read its release notes; a pin bump is a dependency
   change, not housekeeping.
2. Resolve that tag to a commit without trusting a web page:

   ```text
   git ls-remote https://github.com/actions/checkout.git refs/tags/v4.4.0
   ```

   If the tag is annotated, resolve `refs/tags/v4.4.0^{}` as well and pin the
   commit it dereferences to, not the tag object.
3. Replace the SHA and the version comment **together**, in one edit. A SHA
   updated without its comment, or a comment updated without its SHA, is worse
   than no pin: it is a false statement a reviewer will believe.
4. Run `python3 tools/check_repo.py`.

A repository-local action — `uses: ./…` — is versioned by this repository's own
history and is exempt. Nothing else is.


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
