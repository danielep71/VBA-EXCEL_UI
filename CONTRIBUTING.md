<div align="center">

# 🤝 Contributing

**Thank you for improving VBA Excel UI**

[![Conduct](https://img.shields.io/badge/read_first-code_of_conduct-6f42c1?style=flat-square)](CODE_OF_CONDUCT.md)
[![Security](https://img.shields.io/badge/read_first-security_policy-d73a49?style=flat-square)](SECURITY.md)
[![Checks](https://img.shields.io/badge/CI-static_gate-0969da?style=flat-square)](.github/workflows/static-checks.yml)
[![Gate](https://img.shields.io/badge/release-certification_runner-217346?style=flat-square)](#-required-validation)

</div>

---

## 🧭 Before you start

Read [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md) and [SECURITY.md](SECURITY.md).
**Open an issue before non-trivial work** — scope agreed in advance is far
cheaper than scope discovered in review.

This project holds these above convenience, and a change that trades one away
needs to say so explicitly:

| Priority | Why it is non-negotiable |
|---|---|
| 🔒 **Public API compatibility** | Callers are workbooks in production. A rename is a breaking change no matter how much better the name is. |
| 🆘 **Recoverable UI state** | A constrained shell that cannot be undone traps a user in an unusable Excel. |
| 🎯 **Identity-safe restoration** | State applied to the wrong window silently is worse than state not applied at all. |
| 📍 **Explicit state ownership** | Two modules holding the same mutable state means neither can be reasoned about. |
| 🧾 **Deterministic diagnostics** | A failure that is not reported did not happen, as far as the caller is concerned. |
| ⚙️ **32-bit and 64-bit parity** | A defect that only appears on the other bitness is invisible to the person who wrote it. |
| 🧪 **Permanent regression coverage** | A fix without a test is a fix with a scheduled regression. |
| 📖 **Readable exported source** | The `.bas` file is the review artifact; the VBE is not. |

---

## ⚡ Quick reference

```text
python3 tools/check_repo.py                        the gate CI runs
python3 tools/reformat.py --write src/*.bas …      fix house-style drift
python3 tools/reformat.py --selftest               check the formatter itself
Test_EXCEL_UI_RunReleaseCertification              certify behaviour in Excel
```

Static checks run on every pull request. They cannot execute VBA — a hosted
runner has no Excel — so certification stays a manual step on a real host. The
two are complementary, not alternatives.

## 📁 Project layout

```text
VBA-EXCEL_UI/
├─ src/
│  ├─ M_EXCEL_UI.bas
│  ├─ M_EXCEL_UI_RUNTIME.bas
│  ├─ M_EXCEL_UI_SNAPSHOT.bas
│  └─ M_EXCEL_UI_TITLEBAR.bas
├─ demo/
│  ├─ M_DEMO_BUILDER.bas
│  └─ M_EXCEL_UI_DEMO.bas
├─ test/
│  └─ M_EXCEL_UI_REGRESSION_TESTS.bas
├─ tools/
│  ├─ reformat.py                house-style formatter
│  ├─ vba_api.py                 public-declaration contract model
│  ├─ check_repo.py              the static gate CI runs
│  └─ public_api_manifest.txt    declaration contract and release baseline
├─ docs/
│  └─ …                          measurements and reviews
├─ .github/workflows/
│  └─ static-checks.yml
├─ INSTALLATION.md
├─ README.md
└─ …
```

> [!IMPORTANT]
> `tools/public_api_manifest.txt` records the complete normalised declaration of
> every `Public` member in `src/` — parameter order and names, `ByVal`/`ByRef`,
> types, `Optional` status, defaults, return types and enum values. Changing any
> of them requires an intentional edit there, and CI fails otherwise. That
> friction is the point: those are exactly the changes that break callers, and
> they are invisible in a diff of several thousand lines.
>
> The manifest has two sections, and they carry different promises. `[supported]`
> is the caller-facing facade in `M_EXCEL_UI`, covered by Semantic Versioning: a
> change there is a compatibility event for external callers. `[project-public]`
> is the helpers and regression seams that are `Public` only so an
> `Option Private Module` project can see them across its own modules; they are
> tracked so a compile-breaking change is never silent, but no external
> compatibility is claimed for them. Neither section is the deployment unit —
> importing all four `src/` modules together is, and that is a separate promise
> from external API compatibility.

## 🧱 Production module boundaries

### `M_EXCEL_UI`

Owns:

- the supported public `UI_...` API;
- `UIVisibility`;
- general tri-state validation;
- general application/window apply orchestration;
- compatibility wrappers.

It must remain the facade. Internal implementation details should not be moved into caller code.

### `M_EXCEL_UI_RUNTIME`

Owns shared low-level services used by the facade and snapshot engine:

- ordered failure accumulation;
- Immediate Window diagnostic formatting;
- Ribbon reads and writes;
- generic Boolean property reads and writes;
- `ScreenUpdating` quiet-update scopes.

It must not depend on another project module.

### `M_EXCEL_UI_SNAPSHOT`

Owns:

- all mutable snapshot state;
- retained Excel `Window` references;
- capture and restoration orchestration;
- identity resolution and missing-window diagnostics.

It may depend on runtime and title-bar services. It must not duplicate snapshot state in the facade.

### `M_EXCEL_UI_TITLEBAR`

Owns:

- Win32/Win64 declarations;
- exact title-bar style-bit ownership;
- handle-specific captured style state;
- style merging and non-client frame refresh.

It must not depend on another project module.

## 🔗 Dependency rule

The production dependency graph must remain acyclic:

```text
M_EXCEL_UI
├── M_EXCEL_UI_RUNTIME
├── M_EXCEL_UI_TITLEBAR
└── M_EXCEL_UI_SNAPSHOT
    ├── M_EXCEL_UI_RUNTIME
    └── M_EXCEL_UI_TITLEBAR
```

Do not introduce:

| ❌ Never | Because |
|---|---|
| A callback from runtime or title-bar into the facade | It creates a cycle, and the two lower modules are deliberately usable without the facade |
| A second copy of mutable snapshot or title-bar state | Two owners of one truth means neither can be reasoned about, and they will drift |
| Generic utility modules with unrelated responsibilities | "Helpers" is not a responsibility; the module boundary stops meaning anything |
| Public implementation details solely to simplify testing | Anything `Public` is surface that must be kept working |

> [!NOTE]
> Test seams are the deliberate exception, and they earn it: a recovery path that
> cannot be executed is indistinguishable from one that was never written. Each
> seam is `Public` only for same-project access, kept out of the cross-project
> namespace by `Option Private Module`, documented as unsupported, and has a
> caller. A seam without a caller is dead surface — remove it.

## 🌿 Branch workflow

Do not make routine development changes directly on `main`.

Use a branch appropriate to the agreed scope, for example:

```text
fix/snapshot-window-identity
test/titlebar-owned-bits
docs/install-modules
ci/static-release-gate
feature/window-target-scope
release/v<major>.<minor>.<patch>
```

| Prefix | For |
|---|---|
| `fix/` | A defect with an issue |
| `feature/` | Backward-compatible new capability |
| `test/` | Regression coverage only |
| `docs/` | Prose only, no code effect |
| `ci/` | Workflow, gate or tooling |
| `chore/` | Repository configuration and hygiene |
| `release/` | Integration branch for a version |

Confirm the current branch in GitHub Desktop before every commit.

> [!WARNING]
> Editing files through the GitHub web interface while holding unpushed local
> commits diverges the branch. If it happens, `git pull --rebase` replays your
> work on top; setting `pull.rebase true` once avoids the prompt entirely.

## 🔁 Import, edit, compile, test, export

Recommended import order:

```text
src/M_EXCEL_UI_RUNTIME.bas
src/M_EXCEL_UI_TITLEBAR.bas
src/M_EXCEL_UI_SNAPSHOT.bas
src/M_EXCEL_UI.bas
test/M_EXCEL_UI_REGRESSION_TESTS.bas
demo/M_EXCEL_UI_DEMO.bas          when needed
demo/M_DEMO_BUILDER.bas           when needed
```

Workflow:

1. Confirm the current branch.
2. Import the required modules into a controlled workbook.
3. Compile with `Debug → Compile VBAProject`.
4. Run the narrowest focused runner.
5. Run the complete applicable sequence.
6. Perform manual recovery.
7. Re-export each changed module over the matching repository path.
8. Run the static gate locally: `python3 tools/check_repo.py`.
9. Review the GitHub Desktop diff.
10. Update documentation and version metadata.
11. Commit and push.
12. Open a pull request against the agreed base.

Step 8 is the same gate CI runs, so running it locally costs seconds and saves a
round trip. If it reports house-style drift, the fix is mechanical:

```text
python3 tools/reformat.py --write src/*.bas test/*.bas demo/*.bas
```

Re-import any module it rewrites before committing, so the repository and the
VBE do not diverge.

The formatter must never alter a string literal. A literal is data the module
evaluates at run time, so rewriting one changes behaviour rather than layout,
and the change is invisible in a diff that is expected to be whitespace. Any
new transformation therefore uses `split_code_comment` to find where a comment
begins, and `sub_outside_literals` to substitute only outside quoted text —
matching an apostrophe or a keyword directly is what let padding be written
into a literal and a quoted label name be renamed. `--selftest` holds the
fixtures for both rules, and `check_repo.py` runs them.

Any change to a public declaration also requires an intentional edit to
`tools/public_api_manifest.txt`. `tools/vba_api.py --write` regenerates it, and
`--selftest` holds fixtures for every breaking class the gate claims to catch —
parameter reorder and rename, `ByVal`/`ByRef`, type, `Optional`, default, return
type and enum value — alongside formatting-only changes that must normalise
identically. `check_repo.py` runs those fixtures too, so the model is verified
on every push rather than only when someone suspects it.

A member declared in several compilation arms is recorded once. Arms nest, so a
`Win64` split inside a `VBA7` branch declares one member three times. Every arm
must be either the widest declaration or that same declaration with the pointer
types narrowed; arms that are simply identical satisfy that too, which is what a
`Win64` split that changes no signature looks like. Anything else is a real
disagreement between two compilations and is reported rather than normalised
away.

Two declarations are alternatives when they sit in different arms of the same
directive, or in two blocks whose predicates are a condition and its negation —
`#If VBA7` beside `#If Not VBA7` is one member. `#If VBA7` beside `#If Win64` is
not: both hold on a 64-bit VBA7 host, so the member is declared twice and the
compiler rejects it. The negation test is syntactic on purpose. Proving
arbitrary predicates disjoint is not something this tool should attempt, so
anything it cannot show complementary is reported for you to restructure or
confirm.

The arms themselves are recorded, in a fourth manifest field written as
`<effective predicate>#<arm index>` and joined by `>` for nesting. Deleting the
`#Else` arm of a VBA7 pair removes the member from every 32-bit build, and
replacing that `#Else` with `#ElseIf Mac Then` narrows it to one platform —
both while leaving the declaration byte-identical.

The predicate recorded is the effective one, not the arm's own condition. An
arm is reached only when every earlier arm of its block was not taken, so the
`#ElseIf VBA7` arm of an `#If Win64` block records `Not (Win64) And VBA7`, and
its trailing `#Else` records `Not (Win64) And Not (VBA7)`. Changing that
leading `Win64` to `Mac` therefore changes which hosts reach the later arms,
and the manifest says so. Recording `VBA7` and `Else` would not have: those
name a position in the block rather than a condition on the host.

The manifest also carries a `[baseline vX.Y.Z]` section holding the
`[supported]` facade as it stood at the last release. When the two differ,
`CHANGELOG.md` must say so in its `Supported API contract` row and name the
release type — patch, minor or major — or the gate fails. Regenerating the
manifest is therefore not enough on its own to make a facade change pass, which
is the point: the manifest records the change and the changelog declares what it
means. The baseline is rebased only at a release, with
`tools/vba_api.py --rebase-baseline vX.Y.Z`.

The demo workbook is not version-controlled. It is built from the exported demo modules and published as a GitHub Release asset, so changing an exported `.bas` file does not update any committed binary. Rebuild and validate the workbook separately when it is in release scope.

## 🔗 Updating a pinned Action

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

## ✅ Required validation

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

## 🔒 Public API compatibility

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

## 📸 Snapshot changes

Document and test:

- captured state;
- partial capture behavior;
- Window identity strategy;
- new, missing, closed, or recreated windows;
- failure ordering;
- behavior after VBA reset;
- emergency recovery.

Never restore per-window state by collection index.

## 🪟 Title-bar changes

Document and test:

- exact style bits owned;
- 32-bit and 64-bit declarations;
- valid zero WinAPI returns;
- `GetLastError` treatment;
- `Application.Hwnd` changes;
- frame refresh;
- preservation of unrelated current style bits.

Do not restore an entire stale `GWL_STYLE` value.

## 🛡️ Error policy

Production entry points are fail-soft unless the public contract explicitly says otherwise.

| Rule | Rationale |
|---|---|
| Continue after unrelated element-level failures | One unreachable window must not cost the caller the other seven elements |
| Never silently discard a failure | A caller acting on a false success is worse off than one told nothing happened |
| Always restore `ScreenUpdating` | Leaving it suppressed makes Excel look frozen long after the call returned |
| No unsolicited `MsgBox` | A library that blocks on a modal dialog cannot be automated |
| Keep `On Error Resume Next` scopes narrow | A wide scope swallows the error you did not anticipate, which is the one that matters |
| Preserve insertion order in diagnostics | Order is the only clue to which failure caused the others |

> [!CAUTION]
> Two rules protect a diagnostic from destroying the failure it describes. Both
> have been violated in this repository, twice each, in code written by someone
> who had just fixed the other instance.
>
> **1. Anything reachable from an error handler must not be able to raise.**
> A diagnostic that replaces the failure it was invoked to record is worse than
> no diagnostic at all. Set the outputs that cannot fail before attempting
> anything that can.
>
> **2. Never read `Err` after calling anything, or after any `On Error`.**
> Every form of `On Error` resets `Err`, and any procedure you call may contain
> one. Capture what you need into locals first:
>
> ```vb
> ErrNumber = Err.Number
> ErrDescription = Err.Description
> ErrSource = Err.Source
> ErrLine = Erl
> ```
>
> Then log, format and re-raise from the locals. Reading `Err` afterwards yields
> zero and an empty string — and `Err.Raise 0` does not raise at all, so a
> failure reported this way is not reported.
>
> Passing `Err.Number` as a **call argument** is safe, because arguments are
> evaluated before the call. That distinction is subtle enough to be worth
> stating rather than leaving to be rediscovered.

## ✒️ Source style

Every module must use `Option Explicit`.

Production and test modules should retain `Option Private Module` unless an approved API change requires otherwise.

Use the established structured comment style with relevant sections such as:

```text
PURPOSE
WHY THIS EXISTS
INPUTS
RETURNS
BEHAVIOR
ERROR POLICY
DEPENDENCIES
NOTES
UPDATED
```

Prefer explicit procedure boundaries and cohesive modules over clever abstraction.

## 🧹 Repository hygiene

| Kind | Policy |
|---|---|
| Exported `.bas`, `.cls`, `.frm` | CRLF, pure ASCII |
| Markdown and repository configuration | LF |
| Workbooks, images, `.frx`, archives, PDFs | Binary, never line-merged |
| Office lock files, backups, logs, credentials, client data | Never committed |
| Formatting churn unrelated to the change | Never mixed into a functional commit |

Line endings and binary handling are enforced by `.gitattributes`; exclusions by
`.gitignore`; and both by `tools/check_repo.py`, which fails the build if a
workbook binary or lock file is ever tracked. All three are meant to agree — if
you find them disagreeing, that is a defect worth reporting.

Binary workbooks are **release assets, not repository content**. Review any
binary change separately from source.

## 🚀 Pull requests

Keep a pull request focused on **one coherent concern**. A branch that fixes a
defect and tidies formatting produces a diff in which neither can be reviewed.

```text
[ ] Related issue linked
[ ] Public behaviour and Semantic Versioning assessment stated
[ ] Affected UI scope named
[ ] State ownership unchanged, or the change justified
[ ] Diagnostics and failure policy considered
[ ] Validation environment recorded
[ ] Certification verdict pasted
[ ] Recovery evidence (UI_HideExcelUI / UI_ShowExcelUI)
[ ] Documentation updated in the same pull request
[ ] CHANGELOG.md entry added
[ ] Binary demo impact assessed
```

> [!TIP]
> Documentation belongs in the **same** pull request as the change it describes.
> A follow-up documentation commit is a commit that does not get written.

A refactor is complete only when the public API still compiles and behaves
identically under the regression harness. "It should be equivalent" is a
hypothesis; the certification verdict is the test of it.
