<div align="center">

# 📄 Changelog

**All notable changes to VBA Excel UI**

[![Semantic Versioning](https://img.shields.io/badge/versioning-semver-6f42c1?style=flat-square)](https://semver.org/)
[![Format](https://img.shields.io/badge/format-keep_a_changelog-0969da?style=flat-square)](https://keepachangelog.com/)
[![Dates](https://img.shields.io/badge/dates-YYYY--MM--DD-217346?style=flat-square)](#)

</div>

---

Versioning applies to the **public VBA API** — every `UI_…` procedure, enum and
parameter in `M_EXCEL_UI`. Internal module boundaries are not covered by it, so
a release that changes nothing public may still require all four `src/` modules
to be replaced together.

<details>
<summary><strong>Section legend</strong></summary>

<br>

| Section | Contains |
|---|---|
| ➕ **Added** | New members, runners, tools or files |
| 🔧 **Changed** | Behaviour or contract changes to something that already existed |
| 🐛 **Fixed** | Defects, each citing its review finding and issue |
| 📖 **Documentation** | Corrections and additions to prose, with no code effect |
| ✅ **Validation** | The evidence a release was actually certified on |
| 🔗 **Compatibility** | What upgrading requires, and what becomes newly observable |
| ⚠️ **Known limitations** | What is deliberately not fixed, and where it is tracked |

Release types follow semver: 🩹 **patch** corrects defects, ✨ **minor** adds
backward-compatible capability, 💥 **major** may break callers.

</details>

## [Unreleased]

> 🩹 **Patch** · correctness release · public API unchanged

### 🧭 Release intent

A correctness release. Four of the five defects come from an independent review
of `v1.1.1` recorded in `docs/INDEPENDENT_CODE_REVIEW_V1.1.1_2026-08-20.md`; the
fifth was found while writing the regression case for one of them.

What they have in common is the direction of the failure. A registry that
answered for a window it could not identify, a certification runner that
returned silently after failing, a cleanup check that reported a leak where none
existed, a diagnostic that lost the error it was describing, a formatter that
rewrote the text it was meant to align — each reported success, or the wrong
failure, and none of them announced anything. The public API is unchanged.

### ➕ Added

- Added a re-entrancy guard to `Test_EXCEL_UI_RunReleaseCertification`. A nested
  invocation previously reset the outer run's unit records and cleared its
  active flag on exit, leaving the outer verdict describing work it never did.
  The refusal is raised before the error handler is armed, so it reaches the
  caller without disturbing the run in progress.
- Added `TST_CertEvaluateCleanup`, which extracts the certification cleanup
  decision from the runner so it can be tested with crafted inputs instead of a
  full destructive run, and gives the further state comparisons planned for this
  suite one place to live.
- Added `TST_Case_CertificationCleanupUsesBaseline` to the core regression pack,
  which asserts that a suppressed state matching its baseline is clean and that
  a genuine difference is still reported with both values named.
- Added `Test_EXCEL_UI_RunCertificationSelfTest`, which asserts that a failure
  inside certification reaches the caller with its error number and description
  intact. It is a standalone runner rather than a pack case, because the only
  errors travelling the handler path are raised after the counters have been
  reset, and the re-entrancy guard deliberately prevents reaching that path from
  inside a run.
- Added `tools/reformat.py --selftest`, a fixture set asserting the formatter's
  own transformation rules, wired into `check_repo.py` so it runs on every push.
  The formatter is the one tool whose defects the VBA suite cannot observe, and
  `--check` passing proves only that today's modules contain no construct that
  trips it.
- Added `TST_Case_TitleBarStaleFrameEntryNotReused` to the title-bar pack, which
  contradicts a registry entry that claims the frame and asserts the live window
  survives the next show untouched. A reissued handle cannot be produced on
  demand, so the case presents the registry with the same evidence one would:
  an entry claiming a hidden frame against a window carrying owned bits the
  component never wrote.

### 📖 Documentation

- `CONTRIBUTING.md` now states both rules that protect a diagnostic from
  destroying the failure it describes: nothing reachable from an error handler
  may raise, and `Err` must never be read after a call or after any `On Error`.
  Both had been violated twice, each time by someone who had just fixed the
  other instance. The safe exception — passing `Err.Number` as a call argument,
  where evaluation precedes the call — is stated explicitly, because a sweep
  that does not know it will "fix" correct code.

### 🐛 Fixed

- Fixed the release-certification runner not reaching two title-bar
  regression cases. `TST_RunRegressionPack` is the body of the `RegressionPack`
  certification unit, but `TST_Case_TitleBarFrameRefreshDebtRetried` and
  `TST_Case_TitleBarStaleFrameEntryNotReused` were registered only in the
  title-bar-only pack. Both passed there and neither appeared in the evidence a
  release is tagged on — including the case written for this release's own
  frame-registry fix. Every frame-state case is now registered in both. (#41)
- Fixed `tools/reformat.py` rewriting text inside string literals. Two
  transformations shared the assumption that an apostrophe or a keyword means
  the same thing everywhere on a line. Label renaming rewrote `GoTo Fail` inside
  quoted text, changing what a module printed at run time rather than where it
  jumped; and declaration alignment treated the first apostrophe as the start of
  a comment, so a `Const` whose literal contained one had the alignment padding
  written into the literal itself. Both now split the line at the comment marker
  that sits outside quoted text, and substitute only outside literals. Neither
  defect was reachable from any module in the repository, which is why the gate
  stayed green over them, and why the correction ships with its own fixtures.
  (#25)
- Fixed the title-bar frame registry treating a handle match as proof of
  identity. Windows reissues a window handle once the window holding it has
  closed, and `IsWindow` answers for whichever window holds the handle now, so
  liveness could not separate the two. An entry left by a closed window could
  therefore be applied to an unrelated window that had inherited its handle,
  and a show would write a frame that window never had. Every write now records
  the owned bits it leaves behind, and an entry that claims the frame — because
  this component hid it, or still owes it a repaint — must still match those
  bits before it is reused. An entry that cannot be proved is discarded and the
  window treated as one the component has never touched.
  (`ICR-UI-111-P2-01`, #32)
- Fixed certification cleanup reporting false leakage. It required
  `Application.ScreenUpdating` to be `True` rather than comparing it with the
  value captured on entry, so a run started from within a quiet-update scope
  failed certification even though the regression pack had correctly restored
  the suppressed value it found. `OldScreenUpdating` was captured and never
  read. Every cleanup check now compares the exit state with the entry state,
  and a genuine change names both values so it can be diagnosed from the
  evidence file alone. (`ICR-UI-111-P2-02`, #33)
- Fixed `Demo_GetRuntimeErrorText` reading the `Err` object after
  `On Error Resume Next`. Every form of `On Error` resets `Err`, so every
  unexpected-error diagnostic the demo produced reported `0:` with an empty
  description — the failure text was lost at exactly the moment it was needed.
  The fields are now captured before errors are suppressed. (#39)
- Fixed release certification destroying the error it re-raises. The handler
  called `TST_Log`, which contains `On Error Resume Next` and therefore clears
  `Err`, then read `Err` to re-raise. `Err.Number` was zero by that point, and
  `Err.Raise 0` does not raise — so **a failed certification returned silently**
  and a programmatic caller saw a normal return. All fields are now captured
  into locals before anything is called. (`ICR-UI-111-P2-03`, #34)

### 🔗 Compatibility

| Question | Answer |
|---|---|
| Existing calls affected | ✅ none |
| Backward compatible | ✅ yes |
| Release type | 🩹 patch |
| Modules to replace | ⚠️ all four, together |

- No public procedure was added, removed or renamed.
- No existing parameter changed name, position, type or default.
- No enum member or value changed.
- One title-bar behaviour is newly observable. While this component holds a
  frame hidden, a frame change made by Excel or another add-in now discards the
  stored baseline instead of overriding it: the next show adopts the live frame
  rather than restoring the one captured before the hide. From inside the
  component that case cannot be told apart from a handle Windows has issued to a
  different window, and adopting the live bits is the same self-healing rule
  that already applied while the component owned nothing. A caller that hid a
  frame, watched something else change it, and expected a show to undo both
  changes will now see only its own undone.
- `tools/reformat.py` no longer alters text inside a string literal. A module
  quoting a label name in a diagnostic, or placing an apostrophe inside a quoted
  string, was previously rewritten by `--write`. No module in this repository
  does either, so no tracked file changes as a result.

### ⚠️ Known limitations

- Ribbon snapshot restoration is still not window-identity-safe. Ribbon
  visibility is **per workbook window**, and every Ribbon mechanism Excel
  exposes acts on the active window and accepts no window argument, so reaching
  a captured window requires activating it — an observable side effect that does
  not belong in a patch release. Unchanged from `1.1.1`; deferred to `1.2.0`.
  See `#23` and `docs/RIBBON_SDI_BEHAVIOR.md`.
- A frame-state registry entry that claims nothing is reused on a handle match
  alone. That is safe, because its baseline is recaptured from the live window
  before anything is restored from it, but it is a weaker guarantee than the
  retained `Window` object the snapshot layer holds. The strong check exists
  only where a capture and a restore are paired.
- Ribbon behavior has been measured on one host only. It can vary by Office
  channel, update ring and administrative policy.
- The static gate cannot execute VBA, because a hosted runner has no Excel. It
  covers what is decidable from repository text and does not replace
  `Test_EXCEL_UI_RunReleaseCertification`.

---

## [1.1.1] - 2026-08-20

> 🩹 **Patch** · corrective release · public API unchanged

### 🧭 Release intent

A corrective release addressing the findings of the independent `v1.1.0` review
recorded in `docs/INDEPENDENT_CODE_REVIEW_V1.1.0_2026-08-19.md`. The public API
is unchanged: no procedure, enum or parameter is added, removed or renamed, and
no existing call site requires modification.

#### At a glance

| Finding | Issue | What it was |
|---|:--:|---|
| 🔴 `ICR-UI-P1-01` | [#14](https://github.com/danielep71/VBA-EXCEL_UI/issues/14) | Title-bar restoration wrote to whichever window was active |
| 🟠 `ICR-UI-P2-04` | [#15](https://github.com/danielep71/VBA-EXCEL_UI/issues/15) | One frame baseline per process, silently displaced by a second window |
| 🟠 `ICR-UI-P2-03` | [#16](https://github.com/danielep71/VBA-EXCEL_UI/issues/16) | A failed frame repaint reported as a successful no-op |
| 🟠 `ICR-UI-P2-02` | [#17](https://github.com/danielep71/VBA-EXCEL_UI/issues/17) | Diagnostics could raise and destroy the failure being recorded |
| 🟠 `ICR-UI-P2-07` | [#18](https://github.com/danielep71/VBA-EXCEL_UI/issues/18) | A green run could not be distinguished from a partial one |
| 🟠 `ICR-UI-P2-06` | [#19](https://github.com/danielep71/VBA-EXCEL_UI/issues/19) | Tagged documentation still described a pre-release state |
| 🟠 `ICR-UI-P2-05` | [#20](https://github.com/danielep71/VBA-EXCEL_UI/issues/20) | No automated check had ever run on any commit |
| 🟠 `ICR-UI-P2-01` | [#21](https://github.com/danielep71/VBA-EXCEL_UI/issues/21) | Ribbon scope documented without evidence |

🔴 P1 · 🟠 P2 — priorities as assigned by the independent review.

### ➕ Added

- Added explicit-target title-bar entry points to `M_EXCEL_UI_TITLEBAR`:
  `UI_TryGetActiveTitleBarHwnd`, `UI_TryGetTitleBarVisibleForHwnd`,
  `UI_TrySetTitleBarVisibleForHwndIfNeeded` and
  `UI_InternalIsTitleBarFrameAlive`. Callers that must read a frame now and
  write it back later can resolve the target window once and keep it, rather
  than re-resolving `Application.Hwnd` at each end of the operation.
- Added `UI_SnapshotTryGetActiveWindow` and
  `UI_SnapshotTryResolveTitleBarFrame` to `M_EXCEL_UI_SNAPSHOT`, which capture
  the identity of the top-level frame a title-bar value was read from and prove
  that frame is still present before anything is written back to it.
- Added `UI_RuntimeTryAppendFailureEntry` and
  `UI_RuntimeMarkFailureListTruncated` to `M_EXCEL_UI_RUNTIME`, which separate
  the fallible allocation from the infallible status update, and record a
  truncation marker in a slot that already exists when the list cannot grow.
- Added the regression seams `UI_InternalIsFrameRefreshPending`,
  `UI_InternalInjectFrameRefreshFailure` and
  `UI_InternalInjectFailureListGrowthFailure`. They exist because neither a
  `SetWindowPos` failure nor an exhausted allocation can be produced on demand,
  and a recovery path that cannot be executed is indistinguishable from one
  that was never written. Each seam has a caller: a seam without one is the
  same defect wearing different clothes.
- Added `Test_EXCEL_UI_RunTitleBarSdiIdentity`, a regression runner covering
  title-bar restoration across two workbook windows. It verifies that a
  snapshot restores the frame it was captured from while a different window is
  active, and that a captured frame which has since closed is reported rather
  than redirected. The runner is destructive and is invoked explicitly; it is
  not part of `Test_EXCEL_UI_RunAll`.
- Added `TST_Case_TitleBarFrameRefreshDebtRetried` to the title-bar regression
  pack, which injects a frame-refresh failure and verifies that the outstanding
  repaint is recorded and retried on the next call rather than short-circuited
  as a no-op.
- Added `.github/workflows/static-checks.yml`, the first automated check this
  repository has ever carried. It runs on every pull request and push to `main`
  and `release/**`, using only the Python standard library so that it cannot be
  broken by an upstream release and runs identically on a maintainer's machine.
- Added `tools/check_repo.py`, a static gate covering required files, module
  names against filenames, `Option` policy, encoding and line endings, banner
  widths, procedure and directive balance, label vocabulary, jump-target
  resolution, `PtrSafe` on VBA7 declarations, duplicate procedure names, the
  public API manifest, documentation release state, tracked binaries, markdown
  links and house-style conformance.
- Added `tools/public_api_manifest.txt`, recording all 38 public members
  declared in `src/`. A public member can no longer be added or removed without
  an intentional edit to the manifest, which is otherwise invisible in a diff of
  several thousand lines.
- Added `--check` and `--write` modes to `tools/reformat.py`, and module-name
  detection from the file's own `Attribute VB_Name`. Taking the name from the
  file removes a class of caller error: a mismatched name silently changes what
  the hoisting and title passes do, and the result still looks plausible.
- Added `TST_Case_FailureAccumulatorDegradesSafely` to the core regression
  pack, which injects a failure-list growth failure and verifies that the
  status outputs survive, the truncation is reported, and nothing raises.
- Added `Test_EXCEL_UI_RunRibbonSdiProbe`, a characterization probe that records
  Ribbon visibility through `CommandBars("Ribbon").Visible`, the same object's
  `Height`, and the legacy `Get.ToolBar` query, across five scenarios spanning
  two workbook windows and a window created after a hide. It is deliberately a
  probe rather than a test: it asserts nothing, because writing assertions
  before the host behavior is known would encode the assumption the exercise
  exists to remove.
- Added `Test_EXCEL_UI_RunReleaseCertification`, a single runner that executes
  every mandatory regression unit, counts units, failures, skips and cleanup
  separately, verifies the host state afterwards rather than assuming it, and
  emits a JSON evidence document and a text report naming the exact Excel
  build, bitness and operating system the verdict was obtained on. It refuses
  to start when an explicit snapshot already exists, rather than degrading into
  a partial run that reads like a complete one.

### 🔧 Changed

- Replaced the single process-wide title-bar baseline with a frame-state
  registry keyed by top-level window handle. Operating on one workbook window
  no longer discards the baseline captured for another, and entries whose
  window has closed are reclaimed before the registry grows.
- The title-bar baseline is now refreshed rather than captured once. While the
  component does not own a hidden state for a window, the live owned style bits
  are re-adopted on every call, so a legitimate frame change made by Excel or
  another add-in survives the next hide and show instead of being reverted to
  bits captured earlier in the session.
- The snapshot now retains the top-level window handle, the owning Excel
  `Window` object and a diagnostic label for the captured title bar, and
  restores through them. A captured frame that is no longer open is reported as
  a title-bar failure naming the window, instead of the captured value being
  applied to whichever workbook window is active at restore time.
- `FailureCount` is now documented as authoritative and `FailureList` as best
  effort. The list can hold fewer entries than the count when an allocation
  fails, but never silently: a `Diagnostics` truncation marker is written into
  an existing slot whenever growth failed.
- A skipped regression case is now counted rather than only logged. Under the
  certification runner a skipped mandatory case makes the run `INCOMPLETE` and
  therefore not a pass; the legacy runners keep their previous behavior
  exactly, because the accounting is inert outside a certification run.
- `UI_TryGetTitleBarVisible` and `UI_TrySetTitleBarVisibleIfNeeded` are now
  documented and implemented as active-window wrappers over the explicit-target
  entry points. Their signatures and behavior for existing callers are
  unchanged.
- Normalised all seven `.bas` modules to the formatter's normal form. Six had
  drifted from it, by 160 bytes in total: `Const PROC` declarations aligned on
  the `Dim` grid, which the formatter reserves for `Dim`, and between two and
  three trailing blank lines at end of file. No executable token changed.

### 🐛 Fixed

- Fixed title-bar snapshot restoration not being identity-safe under the Single
  Document Interface. `Application.Hwnd` reports the active workbook window's
  handle, and the snapshot re-resolved it on restore, so activating a different
  workbook between capture and restore applied one window's captured title-bar
  state to another. Every API call succeeded, so the misdirection was silent and
  the originally captured frame was left unrestored.
  (`ICR-UI-P1-01`, #14)
- Fixed the title-bar owned-bit baseline being a single process-wide value that
  a second workbook window silently displaced, and that was never refreshed
  after another component legitimately changed the owned frame bits.
  (`ICR-UI-P2-04`, #15)
- Fixed a title-bar style write and its non-client frame refresh not being
  treated as one unit of work. `SetWindowLong` could succeed while
  `SetWindowPos` failed, after which the desired style already matched the
  current style and the next call short-circuited, reporting success over a
  frame Windows had never re-measured. The outstanding refresh is now recorded
  against the window and retried before the no-op test.
  (`ICR-UI-P2-03`, #16)
- Fixed failure accumulation being able to raise from inside an error handler.
  `UI_RuntimeAddFailure` grew the failure list with no error boundary and
  assumed the buffer already held a `String` array whose bound agreed with the
  count. An allocation failure, or a buffer holding anything else, replaced the
  original failure with the failure to record it, could abort a pass designed
  to continue, and could bypass the `ScreenUpdating` restoration in
  `UI_RuntimeEndQuietUpdate`. The status outputs are now set before anything
  fallible is attempted, the entry text degrades rather than failing, and the
  allocation is isolated behind a Boolean contract.
  (`ICR-UI-P2-02`, #17)

- Fixed the regression harness being unable to distinguish a complete pass from
  a partial one. `Test_EXCEL_UI_RunAll` executed no multi-window case, could
  skip the snapshot cases silently when a snapshot already existed, suppressed
  cleanup failures, and reported its outcome only as Immediate Window prose with
  no counters. A green result therefore carried far less information than it
  appeared to. `Test_EXCEL_UI_RunAll` is unchanged and remains the interactive
  runner; release certification now has its own gate.
  (`ICR-UI-P2-07`, #18)

### 📖 Documentation

- Added `docs/INDEPENDENT_CODE_REVIEW_V1.1.0_2026-08-19.md`, the independent
  code and repository review of the `v1.1.0` tag at commit
  `96360379a4bca7703cf649a69a2162961dfa6c9e`. Every issue in the `1.1.1`
  milestone cites it as a stable in-repo reference.
- Added `docs/RIBBON_SDI_BEHAVIOR.md`, recording the measured behavior of the
  Ribbon across multiple workbook windows, the reasoning that selects one model
  from four candidates, and the split between what is corrected in `1.1.1` and
  what is deferred. Measured on Excel 16.0 build 20131, Windows x64, VBA7.
- Corrected `README.md`, which described a pre-release state after `v1.1.0` had
  been merged and tagged: a Release Candidate badge, an entirely unchecked
  release checklist, a remaining-maintenance list of work already completed, and
  references to a branch line that no longer exists. All are removed.
- Corrected the Ribbon's documented scope in `README.md` from *Excel
  application* to *active window*, matching the measurement, and recorded that
  Ribbon restoration is the one managed element that is not window-identity-safe.
- Corrected the title bar's documented scope to name the captured window rather
  than `Application.Hwnd`, and documented the per-window frame registry, the
  self-healing baseline and the retried frame refresh.
- Documented that `FailureCount` is authoritative while `FailureList` is best
  effort, and that a truncated list is always marked rather than silently short.
- Separated source compatibility from package migration in `README.md`: the
  public API is unchanged from `1.0.1`, but all four `src/` modules must be
  replaced together, because the boundaries between them changed at `1.1.0`.
- Documented `Test_EXCEL_UI_RunReleaseCertification` as the release gate, and
  `Test_EXCEL_UI_RunAll` explicitly as not being one.
- Corrected `INSTALLATION.md`, which stated that the title bar retained its
  established scope. It did not: the title bar follows the window a value was
  captured from, and `TargetScope` never applied to it. The section now
  separates the three genuinely application-level elements from the Ribbon and
  the title bar, and states how far the component can keep the per-window
  promise for each.
- Corrected the `INSTALLATION.md` title-bar troubleshooting section, which
  predated per-window frame state and offered no help to a reader who has hit a
  reported `TitleBar` failure or changed the frame of the wrong window. Both
  cases are now covered, along with the add-in interaction the self-healing
  baseline was written for.
- Corrected the branch-naming examples in `CONTRIBUTING.md`, which offered
  `release/v1.1.0` — a branch line that no longer exists.
- Restyled `.github/PULL_REQUEST_TEMPLATE.md`: glyphed type and surface
  sections, the affected-surface list regrouped by actual scope so that the
  Ribbon and title bar are no longer presented as application-level, the
  certification verdict promoted to a required field, and the five
  subsystem-specific sections collapsed behind `<details>` so a documentation
  change is not asked nine questions about WinAPI. Every checklist item is
  retained; the documentation items are consolidated onto one line. The
  environment section now says that attaching the certification evidence file
  replaces the hand-typed fields.
- Restyled `INSTALLATION.md` to match the other repository documents: a
  requirements table, glyphed headings, a routing table that sends a reader to
  the one upgrade path that applies to them, an indexed symptom table above the
  troubleshooting entries, and an emergency-recovery section placed before the
  problem list rather than buried inside it. A same-line upgrade section is
  added, naming the two outcomes that become newly observable without any call
  site changing. The guide is no longer pinned to a version, and the demo asset
  filename is a pattern rather than a fixed name.
- Restyled `CONTRIBUTING.md` to match `README.md` and `CHANGELOG.md`: glyphed
  headings, a priorities table stating why each non-negotiable is
  non-negotiable, a quick-reference block for the three commands a contributor
  needs, branch-prefix and error-policy tables, and a copyable pull-request
  checklist. The project layout now shows `tools/`, `docs/` and the workflow,
  and the public API manifest is explained where a contributor will meet it.
- Replaced the recommended validation sequence in `INSTALLATION.md`,
  `CONTRIBUTING.md`, the pull-request template and the bug-report template with
  `Test_EXCEL_UI_RunReleaseCertification`. All four still listed the four
  pre-`1.1.1` runners, so every document that told a reader how to validate a
  change directed them away from the gate. The narrower runners are retained as
  iteration aids, with `Test_EXCEL_UI_RunAll` marked explicitly as not a
  substitute.
- Documented in `CONTRIBUTING.md` that `tools/check_repo.py` is the same gate CI
  runs and can be run locally, that `tools/reformat.py --write` fixes style
  drift mechanically, and that a change to the public surface requires an
  intentional edit to `tools/public_api_manifest.txt`.
- Documented in `INSTALLATION.md` that the snapshot retains a top-level window
  handle and a `Window` object for the captured title-bar frame, and why neither
  identifies the frame on its own.
- Cross-referenced `docs/RIBBON_SDI_BEHAVIOR.md` from `INSTALLATION.md`, so a
  reader troubleshooting Ribbon behavior reaches the measurements.
- Removed the remaining version pins from `README.md`. The four-module
  requirement and the window-targeting compatibility note were both worded as
  properties of `1.1.0` rather than of the component, so each would have needed
  editing at every release.
- Recorded in `README.md` that the most recently published demo workbook is the
  `v1.1.0` asset, that it does not exercise the current feature set and that its
  preset controls do not work. The release-asset name is now given as a pattern
  rather than a fixed filename.
- Corrected the `diff.vba.xfuncname` recipe documented in `.gitattributes`. Git
  renders capture group 1 when a pattern has one, so the previous regex — which
  grouped only the visibility keyword — produced hunk headers reading `Public`
  instead of the procedure signature. The whole declaration is now the outer
  group.
- Broadened the binary-workbook exclusion in `.gitignore` from `demo/*.xlsm` to
  every macro-enabled and binary Office format, anywhere in the tree. The
  narrower rule left a workbook saved at the repository root, in `test/`, or
  with any other extension silently trackable, while the file's stated policy
  and `tools/check_repo.py` both said such files must not be tracked. The two
  now agree, instead of the gate catching what the ignore rules let through.
- Extended `.gitignore` coverage: merge and patch artifacts (`*.orig`, `*.rej`),
  release-certification evidence copied into the working tree, further Windows,
  macOS and Linux desktop metadata, and local state written by editors and
  coding assistants.
- Extended `.gitattributes` coverage: the remaining Office formats
  (`.dotm`, `.potm`, `.ppsx`, `.ppsm`, `.accde`, `.accdr`, `.mde`, `.ade`,
  `.xlw`, `.xll`, `.thmx`), Office lock files (`~$*`, `.laccdb`, `.ldb`),
  further archive, library, media, font and image formats. Lock files are
  already excluded by `.gitignore`; stating them here as well means one that
  reaches the index by accident still cannot be normalized or line-merged.

### ✅ Validation

Certified in desktop Microsoft Excel for Windows via
`Test_EXCEL_UI_RunReleaseCertification`.

| Host | Value |
|---|---|
| 🖥️ Excel | `16.0` build `20131` |
| 🪟 Operating system | Windows (64-bit) NT 10.00 |
| ⚙️ Bitness | x64 |
| 🧾 VBA generation | VBA7 |
| 🕒 Certified | 2026-08-20 20:13:12 |

```text
RESULT: PASS | COMPLETE | units=3 failed=0 skipped=0 cleanup=OK
  PASS  RegressionPack
  PASS  SnapshotIdentity
  PASS  TitleBarSdiIdentity
```

All four counters are part of the verdict. `failed=0` alone is not a pass:
`skipped=0` confirms nothing was silently omitted, and `cleanup=OK` confirms no
snapshot, stray workbook or suppressed screen update was left behind.

The runner also emits a JSON document and a text report naming the exact host,
written to the temporary folder, so a result can be attached to a release rather
than retyped from the Immediate Window.

Static checks additionally run on every pull request via
`.github/workflows/static-checks.yml`. They cannot execute VBA — a hosted runner
has no Excel — so the two gates are complementary rather than alternatives.

### 🔗 Compatibility

| Question | Answer |
|---|---|
| Existing calls affected | ✅ none |
| Backward compatible | ✅ yes |
| Release type | 🩹 patch |
| Modules to replace | ⚠️ all four, together |

- No public procedure was added, removed or renamed.
- No existing parameter changed name, position, type or default.
- No enum member or value changed.
- The `Stage | Detail` diagnostic format is unchanged. Two new entries can now
  appear: a `TitleBar` entry reporting that the captured frame is no longer
  available, where the previous build silently applied the captured value to the
  active window; and a `Diagnostics` entry reporting that the failure list could
  not be grown.
- `FailureList` may now hold fewer entries than `FailureCount` under memory
  pressure. Callers that assumed the two always agreed should read the count as
  authoritative and treat the list as descriptive.
- Snapshot storage remains in memory only and does not survive a VBA project
  reset or an Excel restart. It now also retains one `Window` reference for the
  captured title-bar frame, released by `UI_ClearExcelUIStateSnapshot`, by a
  replacing capture, or by a project reset.

### ⚠️ Known limitations

- Ribbon visibility is **per workbook window**, not application-wide. Restoring
  a snapshot therefore applies the captured Ribbon value to whichever window is
  active, which need not be the window it was captured from. Every Ribbon
  mechanism Excel exposes acts on the active window and none accepts a window
  argument, so reaching the captured window requires activating it — an
  observable side effect that does not belong in a patch release. Deferred to
  `1.2.0`; see `#23` and `docs/RIBBON_SDI_BEHAVIOR.md`.
- Ribbon behavior has been measured on one host only. It can vary by Office
  channel, update ring and administrative policy.
- The static gate cannot execute VBA, because a hosted runner has no Excel. It
  covers what is decidable from repository text and does not replace
  `Test_EXCEL_UI_RunReleaseCertification`, which remains the behavioral gate and
  is run by a maintainer on a real host.

## [1.1.0] - 2026-08-19

> ✨ **Minor** · backward-compatible feature release

Backward-compatible feature release. Every public `UI_...` procedure and enum
member from `1.0.1` is preserved, with unchanged names, parameter order and
defaults. No migration is required.

### ➕ Added

- Added `UIWindowTargetScope`, a public enum selecting which Excel windows
  receive window-level UI changes:
  - `UI_TargetAllExcelWindows` (0) - every current Excel window;
  - `UI_TargetActiveWindow` (1) - `Application.ActiveWindow` only;
  - `UI_TargetActiveWorkbookWindows` (2) - every window of the active workbook.
- Added an optional trailing `TargetScope` argument to `UI_SetExcelUI` and
  `UI_SetExcelUI_WithResult`, defaulting to `UI_TargetAllExcelWindows`.
  Targeting affects only Headings, Workbook Tabs and Gridlines; the Ribbon,
  Status Bar, Scroll Bars, Formula Bar and Title Bar keep their existing
  application-level and main-window scope.
- Added `UI_CaptureExcelUIState_WithResult`, returning the established
  `Boolean` + `FailureCount` + ordered `FailureList` contract for snapshot
  capture.
- Added `UI_ResetExcelUIToSnapshot_WithResult`, returning the same contract for
  snapshot restoration.
- Added `INSTALLATION.md`, documenting the four-module production package,
  import order, dependency graph, fresh installation, upgrade from the
  single-module architecture, upgrade from intermediate `1.1.0` builds,
  validation and troubleshooting.
- Added a snapshot-lifetime section to `INSTALLATION.md`, documenting that a
  captured snapshot retains one live `Window` reference per captured window,
  that those references are released only by `UI_ClearExcelUIStateSnapshot`, a
  replacing capture or a project reset, and that restoring deliberately retains
  the snapshot rather than releasing it.
- Added `tools/reformat.py`, a deterministic house-style reformatter for
  exported `.bas` modules.
- Added regression coverage for identity-safe window restoration, structured
  snapshot capture and restore results, title-bar owned-bit preservation,
  active-window targeting, active-workbook-window targeting, invalid target
  scopes, title-bar show recovery without a captured baseline, and per-element
  application-level capture and restoration.
- Added the `Test_EXCEL_UI_RunSnapshotIdentity` regression runner.

### 🔧 Changed

- Replaced index-based per-window snapshot restoration with identity-based
  matching. The snapshot now retains each captured `Window` object and restores
  through that reference, so reordered windows restore correctly, windows
  opened after capture are left unchanged, and a closed or recreated window is
  reported rather than having its captured state applied to whichever window
  now occupies the same collection index. Diagnostic captions are stored
  separately and never participate in matching.
- Replaced whole-value title-bar style restoration with an owned-bit merge.
  `TITLEBAR_OWNED_STYLE_MASK` (`&HCF0000`) defines the exact bits this
  component claims - `WS_CAPTION`, `WS_SYSMENU`, `WS_THICKFRAME`,
  `WS_MINIMIZEBOX` and `WS_MAXIMIZEBOX` - and every write merges only those
  into the live style, preserving unrelated changes made by Excel or another
  component after capture.
- Decomposed the production implementation into four cohesive modules while
  keeping `M_EXCEL_UI` as the public facade:
  - `M_EXCEL_UI` - public API, visibility validation, apply orchestration;
  - `M_EXCEL_UI_RUNTIME` - shared fail-soft host operations, result buffers,
    diagnostics, quiet-update scope;
  - `M_EXCEL_UI_SNAPSHOT` - snapshot state, retained window identities, capture
    and restoration;
  - `M_EXCEL_UI_TITLEBAR` - WinAPI declarations, owned style bits, frame
    refresh.

  The dependency graph is acyclic. `M_EXCEL_UI_RUNTIME` and
  `M_EXCEL_UI_TITLEBAR` have no project-module dependency. All internal modules
  use `Option Explicit` and `Option Private Module`.
- Reformatted every `.bas` module to the project house style. Verified
  behaviour-neutral: 3,648 logical statements across the seven modules,
  statement-for-statement identical to their predecessors.
- Stopped version-controlling `demo/EXCEL_UI_DEMO.xlsm`. Tested macro-enabled
  demo workbooks are now distributed as GitHub Release assets only; the demo
  source modules remain in the repository.
- Updated `README.md` for the modular architecture, targeting scopes, snapshot
  identity model, structured diagnostics, installation and release checklist.
- Updated `CONTRIBUTING.md` and the pull-request template for the four-module
  package.
- Moved the diagnostic window-label builder into `M_EXCEL_UI_RUNTIME` as
  `UI_RuntimeBuildWindowLabel`, shared by the apply and snapshot paths, and
  removed the private copy from `M_EXCEL_UI_SNAPSHOT`. The fallback label used
  when Excel exposes neither a caption nor a parent workbook name is now
  `Excel window`.
- Updated all source, demo and regression module metadata to version `1.1.0`.

### 🐛 Fixed

- Fixed `UI_ShowExcelUI` silently failing to restore the title bar when no
  owned-bit baseline had been captured and the frame was already hidden - the
  state reached after a VBA project reset, which is precisely when the
  documented emergency recovery path is needed. The operation reported success
  through both diagnostic paths while the title bar stayed hidden, and no later
  call could recover it. A show with no captured baseline now restores the full
  owned frame. Introduced during this release cycle by the owned-bit merge; not
  present in `1.0.1`.
- Fixed one failed application-level property read discarding the entire
  snapshot. `Application.DisplayStatusBar`, `DisplayScrollBars` and
  `DisplayFormulaBar` were read directly under an active error handler, so an
  ordinary host refusal cleared the Ribbon state, the frame state and every
  captured window identity. Because `UI_CaptureExcelUIState` returns nothing,
  the loss was silent until restore time. All three reads now route through the
  fail-soft helper, record a `Known` flag, and continue the pass.
- Fixed restoration writing default `False` values over good host state after a
  partial capture. Status Bar, Scroll Bars and Formula Bar now carry `Known`
  flags, and restoration writes each only when its captured value is
  meaningful.
- Fixed every `Err`-derived diagnostic reporting `0: ` with an empty
  description. `UI_RuntimeBuildErrorText`, `UI_TitleBarBuildRuntimeErrorText`
  and `TST_BuildRuntimeErrorText` read the `Err` object after executing
  `On Error Resume Next`, and any form of `On Error` resets `Err`. The guard
  intended to stop the formatter raising inside an error handler blanked the
  values it existed to report. All three now capture `Err.Number`,
  `Err.Description`, `Err.Source` and `Erl` before protecting themselves. This
  affected the `Unexpected` stage on both the Immediate Window path and the
  `FailureList` returned by every `_WithResult` API.
- Fixed `TST_SetWindowPos` in the regression harness declaring no `Alias`, so
  VBA searched `user32.dll` for an export of that literal name and raised error
  453. The defect was latent because `TST_TryRefreshWindowFrame` had no caller.
- Fixed one unusable Excel window aborting the rest of a multi-window pass.
  `UI_ApplyWindowLevelState` had no error handler, and the caller's handler
  ends in `Resume Safe_Exit`, so an error raised while processing one window
  abandoned every window still to be visited. The trigger was in the failure
  path rather than the writes: composing a diagnostic read
  `TargetWindow.Caption`, which can itself raise on the window that is already
  failing. The procedure now handles errors locally, records one entry naming
  the window, and returns so the enumeration continues; the label is built once
  on entry, so no window property is read while composing a failure message.

### ✅ Validation

Validated manually in desktop Microsoft Excel for Windows:

```text
Debug -> Compile VBAProject        PASS
Test_EXCEL_UI_RunCore              PASS
Test_EXCEL_UI_RunTitleBarOnly      PASS
Test_EXCEL_UI_RunSnapshotIdentity  PASS
Test_EXCEL_UI_RunAll               PASS
```

Manual checks completed: `UI_HideExcelUI` / `UI_ShowExcelUI` recovery, and
capture / hide / reset validation.

### 🔗 Compatibility

```text
Existing calls affected: none
Backward compatible:     Yes
Release type:            minor
```

- No public procedure was removed or renamed.
- No existing parameter changed name, position, type or default.
- `TargetScope` is declared after `FailureCount` and `FailureList` in
  `UI_SetExcelUI_WithResult`, so existing positional callers are unaffected.
- No enum member or value changed. `UIVisibility` is unchanged.
- The `Stage | Detail` diagnostic format is unchanged. One new stage value,
  `Window [label]`, can now appear in a failure list.
- `UI_ShowExcelUI` still means "show all managed elements", not "restore the
  captured baseline".
- Snapshot storage remains in memory only and does not survive a VBA project
  reset or an Excel restart.
- Installation now requires all four `src/` modules. Importing only
  `M_EXCEL_UI.bas` is not a valid installation; see `INSTALLATION.md`.

### ⚠️ Known limitations

- Ribbon and title-bar control remain best effort and depend on Excel version,
  window state, Windows desktop composition and other loaded add-ins.
- The title-bar show-recovery regression case reproduces the observable
  precondition rather than a real VBA project reset, because VBA offers no
  supported way to clear another module's private state from code.
- The per-element application-level capture case cannot force a host read
  failure; it guards the independence contract rather than the failing read
  itself.
- Hidden Excel UI is not a security boundary.

## [1.0.1] - 2026-07-25

> 🩹 **Patch** · documentation and repository governance

### ➕ Added

- Added `CONTRIBUTING.md` with repository-specific contribution, testing,
  compatibility, WinAPI, and release guidance.
- Added `CODE_OF_CONDUCT.md`.
- Added `SECURITY.md` with supported-version and private-reporting guidance.
- Added a repository-specific `.gitignore`.
- Added `.gitattributes` to preserve deterministic text handling for exported
  VBA modules and binary handling for Excel workbooks.
- Added GitHub issue templates for bug reports and feature requests.
- Added the GitHub issue-template chooser configuration.
- Added a pull-request template tailored to Excel UI, snapshot, WinAPI,
  diagnostics, recovery, and compatibility changes.

### 🔧 Changed

- Redesigned `README.md` as the primary project, API, architecture, integration,
  testing, recovery, and release reference.
- Updated the core, demo, and regression-test module metadata to version
  `1.0.1`.
- Updated module documentation dates to `2026-07-25`.
- Reduced repetitive comments while preserving section banners, procedure
  headers, and declaration-level inline comments.
- Corrected stale or imprecise documentation concerning:
  - `UI_ShowExcelUI` versus explicit snapshot restoration;
  - fire-and-forget versus structured-result behavior;
  - Ribbon and title-bar state handling;
  - index-based per-window snapshot restoration;
  - title-bar style restoration limitations;
  - demo-module dependencies and button assignments.
- Synchronized `demo/EXCEL_UI_DEMO.xlsm` with the exported versioned VBA
  modules.

### ✅ Validation

The release candidate was validated manually in desktop Microsoft Excel:

- `Debug -> Compile VBAProject`
- `Test_EXCEL_UI_RunCore`
- `Test_EXCEL_UI_RunTitleBarOnly`
- `Test_EXCEL_UI_RunAll`

All three regression runners completed successfully.

### 🔗 Compatibility

- No public procedure signature changed.
- No public enum member or value changed.
- No migration is required for existing callers.
- No executable VBA behavior was intentionally changed.
- GitHub Actions workflows are intentionally not included in this release.

[Unreleased]: https://github.com/danielep71/VBA-EXCEL_UI/compare/v1.1.1...HEAD
[1.1.1]: https://github.com/danielep71/VBA-EXCEL_UI/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/danielep71/VBA-EXCEL_UI/compare/v1.0.1...v1.1.0
[1.0.1]: https://github.com/danielep71/VBA-EXCEL_UI/releases/tag/v1.0.1
