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
   architectural change, compatibility break, or broad refactor. Use the issue
   forms; [Issue labels & automation](docs/ISSUE_LABELS.md) explains
   priorities, labels, milestones, and dependency status.
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
| `tests/` | Unit tests for the static validation in `tools/` |
| `docs/` | Issue and label rules, measured Ribbon SDI behaviour, and the v1.1.3 implementation sequence |
| `images/` | Documentation images, the only approved home for binary files |

The README and current tree are authoritative if a listed optional directory is
not present.

Binary files belong under `images/`. Anywhere else, `tools/check_repo.py`
rejects a tracked file that Git treats as binary (its `text` attribute is unset
in `.gitattributes`) or whose content contains a NUL byte, unless
`tools/binary_allowlist.txt` lists it with the SHA-256 of its exact content. The
digest makes the approval specific to that content: replacing the file fails
the gate until the entry is updated in the same change.

---

<a id="development-workflow"></a>

## 🌿 Development workflow

1. Fork or clone the repository and start from the base branch described
   below.
2. Create a short, focused branch such as `fix/clear-description`,
   `feat/clear-description`, `docs/clear-description`,
   `test/clear-description`, or `chore/clear-description`.
3. Reproduce the existing behavior before changing it.
4. Define the intended contract, affected callers, compatibility impact, and
   evidence plan.
5. Make the smallest coherent source change; do not mix unrelated formatting,
   refactoring, generated files, or cleanup.
6. Compile and run the relevant static, regression, host, and manual checks.
7. Re-export changed VBA components and review the complete text/binary diff.
8. Update documentation and release notes required by the change.
9. Push the branch and open a pull request with evidence and limitations.

<a id="base-branch"></a>

### Base branch

Start everyday work and governance changes from the current `main`. While a
release branch `release/vX.Y.Z` is open, VBA corrections for that release
target the release branch and reach `main` with the release; non-VBA repository
changes are made on the release branch and mirrored to `main`.

Every change to `main` and `release/**` goes through a pull request, including
mine. Repository rulesets enforce it: on `main` and `release/**` a pull request
must pass **Repository and module checks** on an up-to-date branch with every
review thread resolved, and is merged with a merge commit; a push to the pull
request dismisses earlier approvals. Force-pushes and deletion are blocked, and
`v*` tags cannot be moved or deleted. The rulesets are versioned in `.github/rulesets/`,
and the **Ruleset drift** workflow fails daily if a live ruleset differs from
its copy. To change a ruleset, change it in the repository settings, run
`python3 tools/check_rulesets.py --repository danielep71/VBA-EXCEL_UI --export`,
and commit the result in a pull request that explains why.

### Commit discipline

Write imperative, specific subjects, normally in this form:

```text
fix: preserve caller-owned UI snapshots
feat: add an explicit window-targeting option
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

Legacy root-only `MUTANT_*.bas`, `CONTROL_*.bas`, `commit-msg.txt`,
`msg-p2fix.txt` and `wiki/` are also ignored and excluded from source archives.
These narrow rules do not hide similarly named nested source or fixtures. Put
new mutation work in `/.mutation-scratch/` and new message files / Wiki clones
outside the repository. Inspect `git status --short` and stage explicit paths.
Ignore rules do not untrack a file that has already been committed.

The recorded v1.1.3 inventory accounts for ten variants (including the hanging
attempt and the insufficiently isolated generation control):

| Issue | Disposable artifact type | Purpose |
| --- | --- | --- |
| #43 | Four one-runner-at-a-time refusal mutants | Prove caller-owned snapshots survive each refusal |
| #43 | Cleanup-skipping mutant | Prove the success-path snapshot-release assertion detects leakage |
| #26 | Two quiet-scope variants | Detect the old entry-read behavior and missing ownership readback; no dedicated End-side mutation |
| #32 | Retained-Window generation-proof variant | Fails with the current distinct-handle seam; same-numeric-hWnd generation proof is still missing |
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

### Formatter verification (#52)

Exported VBA source is **ASCII only**, with CRLF checkout/output line endings.
The formatter rejects non-ASCII input, including comments, before writing; it
never silently transliterates or changes the VBE code page. The existing Git
attributes and Windows-native file overrides remain authoritative.

`tools/vba_lex.py` is the shared code/string/comment scanner used by the
formatter and VBA analyzer. It recognizes doubled
quotes, apostrophe comments, statement-position `Rem`, and bracketed identifiers.
It is not a complete VBA parser. The analyzer consumes the same lexical
boundaries for procedure-local resolution and targeted assignment analysis.

Conditional blocks (including their directives and whitespace) and continued
logical statements are preserved. `--check` reports that their formatting was
skipped; clean preserved input may pass. Malformed literals, interrupted or
unterminated continuations/conditionals, non-ASCII text and ambiguous label
renames fail before any file is written. Line endings normalize to CRLF.

Before publication, the formatter compares ordered executable tokens and
statement boundaries, preserving literal bytes and Option statements; only the
explicit house-label renaming and Option relocation are normalized. This
bounded equivalence check is not compilation or Excel runtime certification.
Every fixture is exercised by the existing formatter self-test in static CI.

`--write` preflights all inputs and replaces each changed file through a flushed
same-directory temporary file and `os.replace`. Existing permissions are kept,
symlinks are refused, and failed writes/replacements remove their temporary
file. This is atomic per file, not a multi-file transaction or a guarantee
against power loss. The legacy source/destination form uses the same writer.

Always inspect the resulting diff and re-import any changed VBA modules into
the VBE before committing them. A formatter change that leaves every module
byte-identical does not create a new Excel runtime result.

### Bounded VBA analyzer (#31)

`python tools/vba_analyze.py` checks all `.bas` files in `src/`, `test/` and
`demo/`. The static gate runs it and its independent malformed/benign fixtures
on configured pushes and pull requests. Diagnostics identify the source path,
physical start line, rule and compilation configuration. A continued statement
uses its first physical line. Dynamic dispatch appears as `INVENTORY` in the
gate log; these entries are unresolved host/binding work, not validated calls.

The analyzer collects procedure, DLL, constant, variable, type and enum names;
resolves project-prefixed references against visible declarations; and checks
simple unqualified assignment targets under `Option Explicit`. Parameters,
locals and visible module/project names are recognized after continuation and
colon normalization. This is declaration coverage, not type checking: member,
array-element, `With` and host-object resolution, argument compatibility and
complete expression/loop-variable analysis still require the VBA compiler.

Labels resolve within their procedure, opener/end kinds must agree, and
duplicates are checked independently for VBA7 x64, VBA7 x86 and legacy x86.
Conditional directives use a nested stack. Boolean `VBA7`/`Win64` expressions
and local Boolean `#Const` definitions are supported; unknown symbols or
unsupported expressions fail explicitly, including in inactive nested blocks.
Each effective VBA7 `Declare` requires `PtrSafe`; prefixed DLL declarations
require an explicit `Alias`. The eight known WinAPI contracts have versioned
ABI-v1 fixtures covering parameter types, ByVal passing and return types. This
is not a general DLL export/ABI verifier.

The error-preservation rule covers straight-line handler segments and direct
reads of `Err.Number`, `Err.Source` and `Err.Description` after a call or
`On Error` statement. Arguments are read before the call changes the state;
saved-local re-raises pass. A subsequent error-producing operation under
`On Error Resume Next` is a fresh source, while a scalar flag assignment is
not. Scalar capture intrinsics (`IIf`, `Len`, `CStr`, `CLng`, `CBool`, `CInt`,
`CDbl`, `CSng`) are treated as capture expressions. This does not prove their
arguments cannot raise. Branches, loops and label joins delimit the analysis;
there is no interprocedural or whole-control-flow proof, cleanup proof, or
protection against re-raise loops.

Run `Debug -> Compile VBAProject` and the relevant Excel tests after VBA edits.
Passing static analysis does not produce runtime evidence. The mandatory
certification-manifest registration check remains dependent on #42's manifest;
the current analyzer does not certify case completeness or execution.

### WinAPI error capture (#27)

In `src/` and `test/`, read a native error code only through
`Err.LastDllError`, and do not declare `GetLastError`. VBA sets
`Err.LastDllError` after every `Declare` call, the next `Declare` call replaces
it, and `Err.Clear` or any `On Error` statement resets it. Write every native
call whose error you report as three adjacent statements:

```vba
SetLastError 0                                  ' immediately before the call
ApiResult = GetWindowLongPtr(hWnd, GWL_STYLE)   ' result copied to a local
DllError = Err.LastDllError                     ' the very next statement
```

Put nothing between them: no other call, no formatting or logging helper and
no `On Error` statement. Take a regression seam before `SetLastError 0`, and
replace captured values only after the capture. `GetWindowLong[Ptr]` and
`SetWindowLong[Ptr]` can return zero legitimately, so treat zero with a
captured error of zero as a valid value, zero with a non-zero captured error as
a failure, and any non-zero result as success. Report a failure with the
exported API name, the operation, the captured error and the target hWnd.

The gate's `WinAPI error capture` check
(`tools/winapi_capture_fixtures.py`) enforces this sequence separately for VBA7
x64, VBA7 x86 and legacy x86, and proves itself against passing and failing
fixtures first. It reads statement order only: it does not follow `GoTo` and
cannot observe the error a call leaves, so
`TST_Case_TitleBarApiErrorCapture` must still pass on a real Excel host. The
demo modules in `demo/` are not yet held to this rule (#144).

The VBA that follows this rule, the gate check and the regression case exist
only on `release/v1.1.3` until v1.1.3 is released. On `main`, the title-bar
module still declares `GetLastError`, the check is not registered in
`tools/check_repo.py` and the case is absent, so a passing gate on `main` does
not enforce this sequence.

### Public contracts and compatibility

Demo cleanup (#71) is single-pass, not a host watchdog. With
`ReProtectAtEnd=True`, reset captures the fifteen Boolean `Worksheet.Protect`
options and `EnableSelection` before unprotecting, then attempts to restore
them only if that invocation's unprotect returned successfully. A missing or
incorrect password must not arm reprotection. With `ReProtectAtEnd=False`
(the existing default), leaving the sheet unprotected is intentional.

An operation error retains its number/source and description prefix; cleanup
failures append a separately identified diagnostic. A cleanup-only failure is
raised to the caller. This does not restore deleted cells, shapes, tables or
range-level permissions, nor guarantee recovery after Reset/End or a blocked
Excel call. `UserInterfaceOnly` is restored for the current session, not promised
across workbook reopening. Source fixtures do not establish Excel outcomes.
Before closure, use disposable worksheets and a bounded host harness to test
ordinary failure, missing/wrong password, reprotection and selection failure,
and read back every captured option at the exact tested commit.

`demo/M_DEMO_CLEANUP_TESTS.bas` is that harness; it exists only on
`release/v1.1.3` until v1.1.3 is released. Build a new demo workbook from the
commit under test:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-EXCEL_UI-VBA.ps1 -WorkbookPath "C:\work\EXCEL_UI_Demo.xlsm" -Mode Demo -NewWorkbook
```

Open it, run **Debug → Compile VBAProject**, and run
`Demo_Test_RunCleanupChecks` from the Immediate window or the Macros dialog. It
creates a temporary workbook, runs fourteen scenarios on sheets of its own,
closes the workbook without saving and prints one `PASS` or `FAIL` line per
scenario, then a summary line that starts with `DEMO CLEANUP CHECKS: PASS` or
`DEMO CLEANUP CHECKS: FAIL`. A failed run also raises an error after cleanup. An
operation error after a successful unprotect, a failed Protect and a failed
`EnableSelection` restoration have no worksheet cause that can be produced on
demand, so the harness injects them through the one-shot
`DEMO_InternalCleanupProbe*` seams of `M_DEMO_BUILDER`. Those seams stay inert
unless the harness arms them. Record the commit, Excel build and bitness with
the Immediate window output.

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
| **Title-bar ownership** | Restore only the style bits owned; capture `Err.LastDllError` immediately after each WinAPI call; handle valid zero returns, handle changes, and frame refresh. |
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
substitute for certification: `Test_EXCEL_UI_RunAll` does not run the
snapshot-identity and SDI identity units, and produces no machine-readable
evidence. It still opens and closes temporary windows and workbooks.

CI runs the static gate on every pull request. It cannot execute VBA — a hosted
runner has no Excel — so certification remains a manual step on a real host, and
the two are complementary rather than alternatives.

On Windows, the workbook tools shorten that step:

- `tools\Sync-EXCEL_UI-VBA.cmd` imports the modules of a configuration from
  `tools/vba-projects.json` into a workbook ([docs/WORKBOOK_SYNC.md](docs/WORKBOOK_SYNC.md));
- `tools\Export-EXCEL_UI-VBA.ps1` and `tools/compare_vba_export.py` check that a
  workbook holds exactly the VBA of a commit;
- `tools\Certify-EXCEL_UI.cmd` builds, compiles and certifies a development
  workbook from the checked-out commit and writes an evidence folder that
  `tools/check_excel_evidence.py` validates
  ([docs/EXCEL_CERTIFICATION.md](docs/EXCEL_CERTIFICATION.md)).

The required job also runs `tools/check_vba_calls.py`, which resolves calls
across the modules of each configuration
([docs/VBA_CALL_CHECKER.md](docs/VBA_CALL_CHECKER.md)). It is advisory: its
result appears in the job summary and never blocks a merge, because compiling
in Excel remains the authoritative check.

The Python tools under `tools/` need **Python 3.11 or later**, and the static
gate needs only the standard library. The issue-form check
(`python3 tools/check_issues.py --check-config`) and its tests in
`tests/test_check_issues.py` also need Node.js, because they run
`.github/scripts/labels-sync.mjs`; CI provides it. The required check runs the
gate on an exact Python 3.12 release; a second, non-required job runs the same
gate on Python 3.11 so that a change needing a newer interpreter is reported
before it reaches your machine.

Other workflows run alongside the gate and are not required checks:

- **CodeQL** analyses the Python tooling, the label scripts and the workflows
  with the `security-extended` queries. On a pull request it runs read-only and uploads
  nothing.
- **Workflow lint** runs a checksum-verified actionlint release on every
  workflow, after proving that it rejects `.github/actionlint/invalid-workflow.yml`.
- **Ruleset drift** compares the live rulesets with `.github/rulesets/`.
- **External links** checks the public URLs in the tracked Markdown weekly and
  on demand.
- **Wiki badges** checks the Wiki against `VERSION`.
- **Sync issue labels**, **Detect issue-label drift**, **Issue triage** and
  **Keep milestones issue-only** manage issues and labels, as
  [docs/ISSUE_LABELS.md](docs/ISSUE_LABELS.md) describes.

Record only environments actually tested.

### Testing the gate itself

The static gate is tested like any other code. `tests/` holds unittest modules
that copy this repository into a temporary Git repository, plant one defect,
and run one check against it:

```text
python3 -m unittest discover -s tests
```

Every entry in `CHECKS` in `tools/check_repo.py` has a passing fixture (the
current tree) and at least one failing fixture (a named defect).
`tests/test_coverage.py` fails when a check has no test in either direction, so
a new check needs a `GateTestCase` subclass in `tests/test_gate_checks.py`
before it can merge. Like the gate, the tests use only the standard library,
except that `tests/test_check_issues.py` also needs Node.js. The Python tools
run their self-tests with `--selftest`; the Node.js label scripts under
`.github/scripts/` use `--self-test`. `test/` is a different directory: it
holds the VBA regression suite.

The Python code under `tools/` and `tests/` must also be clean under Ruff and
`mypy --strict`, both configured in `pyproject.toml` against Python 3.11, the
oldest interpreter I support for the gate. CI runs both on `tools/` and
`tests/` and installs them from `tools/requirements-quality-ci.txt`, where
every wheel is hash-locked. To run
them locally, install them into a virtual environment the same way:

```text
python3 -m pip install --require-hashes --only-binary=:all: -r tools/requirements-quality-ci.txt
ruff check tools tests
mypy --strict tools
mypy --strict tests
```

A finding is fixed in the code, not silenced, unless the fix would change what
a check accepts or rejects. The few complexity exemptions carry their reason
beside them.

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

### External links

The static gate checks links inside the repository on every pull request. The
**External links** workflow checks every external link in the tracked Markdown
files each Monday, and on demand, with `tools/check_external_links.py`. It is not
a required check: another host can be down for reasons unrelated to your change.
A `404` or `410`, or a link that still fails after retries, fails the run; `401`,
`403` and `429` are reported as unverified. To run it locally:

```text
python3 tools/check_external_links.py
```

Add a prefix to `tools/external_link_exclusions.txt` only for a host that
refuses automated requests, with a comment saying why.

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
release would pass. The workflow declares scheduled and push triggers because a Wiki edit does
not touch this repository. The workflow is on both `main` and
`release/v1.1.3`, so its scheduled and manual runs operate from the default
branch. Do not treat a declared trigger
or the historical badge run as a fresh Wiki check. #53 and #49 require a runnable
check bound to the final repository and Wiki SHAs before release readiness.

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

1. Review the governed Wiki pages against the candidate contract, update every
   badge to `wiki_tracks-v1.1.3`, and record the Wiki SHA first, as required by
   [#53](https://github.com/danielep71/VBA-EXCEL_UI/issues/53).
2. Update root `VERSION` and the reviewed module headers together under
   [#36](https://github.com/danielep71/VBA-EXCEL_UI/issues/36).
3. Treat the interval between these separate repository updates as transitional;
   no release evidence is valid during that interval.
4. Run the badge check against the settled repository and recorded Wiki SHAs.
   Require a passing result before the exact-head freeze.

Order matters. Wiki edits do not touch this repository, so re-badging cannot
invalidate a frozen tree — but bumping `VERSION` can, because it is a tracked
file. Bump `VERSION` and re-badge the wiki *before* the release freeze, never
during it.

The currently configured required status check does not include the Wiki job.
#49 must make missing, failed, stale or unreadable final Wiki evidence block
release readiness; a historical passing run does not satisfy that requirement.

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
identifier is the only public handle a reader has. Those identifiers resolve
against the review finding crosswalk in `CHANGELOG.md` (under
`[Unreleased]`, tracked in
[#48](https://github.com/danielep71/VBA-EXCEL_UI/issues/48)). When you cite an
identifier that is not in the table yet, add its row in the same pull request:
`check_repo.py` fails on a cited identifier without exactly one row.

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

Dependabot (`.github/dependabot.yml`) opens a weekly grouped pull request when a
pinned Action has a newer release, and a monthly one for the hash-locked Python
quality tools. It changes the SHA and the version comment together; review it
with the steps above before merging. While a release branch is open, Dependabot
opens the same updates against it too, so neither branch falls behind; remove
those entries when the release branch is merged.

The actionlint release that **Workflow lint** downloads is neither an Action nor
a Python package, so Dependabot does not update it. Update
`ACTIONLINT_VERSION` and `ACTIONLINT_SHA256` in
`.github/workflows/actionlint.yml` by hand, together, in one edit.

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

<a id="pull-request-reference"></a>

### Pull request reference

The pull request template asks for the classification and the affected
surface in one line each. Choose from these lists.

#### Change classification

- Defect correction
- Backward-compatible capability
- Breaking API, behavior, deployment, or migration change
- Internal refactor with no intended supported-behavior change
- Test, fixture, reference-data, or validation change
- Performance change
- Security or trust-boundary hardening
- Documentation-only change
- Repository tooling, workflow, or governance change
- Packaging or release preparation
- Excel-state ownership or recovery change
- Window identity, title-bar, Ribbon, or WinAPI change
- Snapshot, diagnostic, or structured-result change

#### Affected surface

- Public `UI_*` facade
- Runtime and application-level UI state
- Snapshot capture and restoration
- Title-bar and native-window integration
- Multi-window targeting and identity
- Diagnostics and recovery
- No runtime or supported surface — documentation/repository-only

<a id="documentation-and-release-hygiene"></a>

### Documentation and release hygiene

Before you request review, check each document your change can affect:

- `README.md` reflects supported behavior and examples.
- `INSTALLATION.md` reflects paths, dependencies, import order, validation, upgrades, and removal.
- `CONTRIBUTING.md` reflects development and evidence requirements.
- `CHANGELOG.md` records material change under `[Unreleased]`.
- `SECURITY.md` reflects supported versions or trust boundaries.
- `RELEASING.md` reflects certification, package, provenance, or recovery changes.
- Source headers, API references, demos, Wiki pages, and counts remain synchronized.
- Version markers remain unchanged unless this is the deliberate release-stamp change.
- No documentation change is required — state the reason in the pull request.

<a id="regression-coverage"></a>

### Regression coverage

State in the pull request which of these hold, and why any other does not:

- Existing tests cover the changed success path.
- New or amended tests cover each corrected defect.
- Boundary, invalid-input, failure, fallback, and cleanup paths are covered as applicable.
- Test entry points and inventory/count metadata remain synchronized.
- Expected results come from the contract or an independent reference.
- No regression change is needed — state the rationale in the pull request.

<a id="security-and-provenance-review"></a>

### Security and provenance review

Every pull request must satisfy these, or state the exception:

- No credential, secret, signing material, internal URL, or personal path is included.
- No client, employer, counterparty, student, personal, or restricted production data is included.
- Test data is synthetic, anonymized, or explicitly redistributable.
- External algorithms, code, datasets, and market/vendor data have attributable provenance and compatible licensing.
- Formula, command, path, callback, deserialization, and external-content injection surfaces were assessed.
- No security-sensitive detail belongs in private disclosure instead of this pull request.
- Generated evidence identifies its inputs, tool/runtime version, candidate SHA, and limitations.

<a id="risk-levels"></a>

### Risk levels

- **Low** — documentation, metadata, or mechanically verified change.
- **Medium** — bounded runtime, tooling, or compatibility impact.
- **High** — numerical integrity, shared Excel state, native API, security,
  release, or breaking impact.

<a id="project-specific-review-profiles"></a>

### Project-specific review profiles

Apply every profile your change touches, and list in the pull request any item
it does not meet.

#### 🧱 Module ownership and API manifest

Apply when module boundaries or supported surface can change.

- `M_EXCEL_UI` remains the public facade.
- Runtime and title-bar modules retain their allowed dependency boundary.
- Snapshot state exists only in `M_EXCEL_UI_SNAPSHOT`.
- Title-bar mutable state exists only in `M_EXCEL_UI_TITLEBAR`.
- No circular dependency or duplicated mutable state is introduced.
- Internal modules retain `Option Private Module`.
- `tools/public_api_manifest.txt` intentionally reflects every public-surface change.

> [!IMPORTANT]
> Any change to a Public declaration under src requires an intentional update to
> tools/public_api_manifest.txt, which records the full normalised signature:
> parameter order and names, ByVal/ByRef, types, Optional status, defaults,
> return types and enum values. Regenerate it with tools/vba_api.py --write.
>
> Passing the manifest gate proves the declared contract matches that file. It
> does not prove behavior. A procedure whose signature is untouched can still
> change what it does, and that belongs under Verification.
>
> A change under [supported] is a Semantic Versioning event for external
> callers and must be declared as one in the pull request and in CHANGELOG.md. The gate
> enforces that: the manifest carries the facade as it stood at the last
> release in its [baseline] section, and when the two differ the CHANGELOG
> Supported API contract row has to name the release type or the build fails.
> Regenerating the manifest alone does not clear it.
>
> A change under [project-public] only has to be deliberate. Neither is the
> same as the deployment rule that all four src modules are replaced together.

#### 📸 Snapshot, identity, and recovery

Apply when capture, restoration, targeting, or reset behavior changes.

- Per-window restoration never relies on collection index.
- Object identity and recycled-window-handle risks are addressed.
- New, missing, closed, and recreated windows have explicit behavior.
- Reset without a valid snapshot is controlled.
- State is never applied to a window that cannot be proven to be the captured target.
- `UI_ShowExcelUI` remains an emergency show-all recovery route.

#### 🪟 Ribbon and WinAPI

Apply for Ribbon, title-bar, or native-frame changes.

- Only owned style bits are changed and unrelated bits are preserved.
- 32-bit and 64-bit declarations are correct.
- Valid zero returns are distinguished from native failures.
- `SetLastError 0` directly precedes each ambiguous-zero call, and
  `Err.LastDllError` is copied to a local in the statement directly after it.
- Required frame refresh is performed and failures remain observable.
- Multi-window behavior is tested with more than one workbook open.

#### 🧾 Diagnostics and cleanup

Apply when failure ordering, results, or Excel state can change.

- Structured diagnostics preserve insertion order and real post-operation state.
- Error handlers cannot replace the original failure.
- `ScreenUpdating` and other caller-owned state are restored.
- Best-effort continuation is deliberate and failures are not discarded.
- Certification reports complete, failed, skipped, and cleanup counters.

---

## 🤝 Review and decisions

Reviewers evaluate correctness, safety, maintainability, compatibility,
evidence, documentation, and fit with the project's direction. Approval of an
idea does not guarantee acceptance of every implementation detail.

I may edit, squash, defer, or decline a contribution to protect the
coherence and supportability of the project. Contributors will be credited
through Git history and release notes where appropriate.

---

## 📄 Licensing

By contributing, you agree that your contribution is licensed under the
repository's [MIT License](LICENSE). You must have the right to submit every
part of the contribution, including code, tests, data, images, and generated
material.

---

## 👤 Contact

I, **Daniele Penza**, maintain this repository.

For ordinary contributions, use GitHub issues and pull requests. For sensitive
security matters, use the private channel in [SECURITY.md](SECURITY.md).

---

### Contribution principle

> Make the contract explicit, keep the diff focused, and leave evidence another
> person can reproduce.
