# 🧪 Excel certification in one command

`tools/Certify-EXCEL_UI.ps1` certifies the commit you have checked out on a
Windows host with Excel. It builds a development workbook from that commit,
compiles it, runs `Test_EXCEL_UI_RunReleaseCertification`, and collects the
result into an evidence bundle that names the commit, the host and the source
it ran. `tools/check_excel_evidence.py` then validates the bundle against the
commit, on any machine with a clone. Both read what a bundle must contain from
one file, [the evidence policy](#-the-evidence-policy).

The script automates the compile and certification steps of
[RELEASING.md](../RELEASING.md). It does not replace the manual checks listed
there (hide/show, snapshot/restore, multi-window identity, title-bar behaviour,
cleanup on error paths), and it runs nothing but certification:
`Test_EXCEL_UI_RunRibbonSdiProbe` is a characterization that asserts nothing
(see [RIBBON_SDI_BEHAVIOR.md](RIBBON_SDI_BEHAVIOR.md)) and is not part of the
release gate, so its output is not part of a bundle.
`tools/Check-EXCEL_UI.ps1` runs this script unchanged and then the other
runners and the demo checks, with a log for each
([EXCEL_CHECKS.md](EXCEL_CHECKS.md)).

## ✅ Requirements

- Desktop Microsoft Excel for Windows on 64-bit Office, as in the
  [support baseline](../INSTALLATION.md). 32-bit Office is not certified.
- **Trust access to the VBA project object model** turned on in
  **File → Options → Trust Center → Trust Center Settings → Macro Settings**.
  The workbook is built and compiled through the VBA object model, so the
  script stops when this is off. It never turns it on for you.
- A macro setting that lets you enable the workbook's macros, such as
  **Disable VBA macros with notification**. The script opens the workbook with
  the Trust Center setting and never changes it. With **Disable VBA macros
  without notification**, Excel offers no **Enable Content** button and the run
  cannot proceed.
- Git on the `PATH`, and a clean checkout of the commit to certify.
- `tools/Sync-EXCEL_UI-VBA.ps1`, which builds the workbook.
- Python 3.11 or later, to validate the bundle at the end of the run. Without
  it the run still produces the bundle; validate it elsewhere.

Save your work in any open Excel window first. The script starts an Excel
instance of its own and leaves your other instances alone, but certification
creates, activates and closes windows, and a busy desktop makes the run harder
to follow.

## ⚡ Run it

```powershell
git switch --detach <full-sha>
git status
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-EXCEL_UI.ps1
```

or double-click `tools\Certify-EXCEL_UI.cmd`, which runs the same script and
keeps the window open at the end. To run without any prompt, add `-Unattended`
to either command; the launcher then skips its final pause. See
[Running unattended](#-running-unattended).

| Parameter | Effect |
| --- | --- |
| `-BundleDir <folder>` | Writes the bundle there instead of the default folder. It must be outside the checkout and new or empty. |
| `-AllowDirty` | Runs even though the working tree has changes. The bundle records it, and the validator rejects such a bundle. Use it only to try the tooling. |
| `-Python <exe>` | The Python 3.11+ executable for the validation. By default the script tries `py -3`, `python` and `python3`. |
| `-RepoRoot <folder>` | The repository. By default, the parent of the `tools` folder. |
| `-Unattended` | Runs without any prompt, as [Running unattended](#-running-unattended) describes. |

## 🤖 Running unattended

With `-Unattended` a run needs no key press and no click, from the start to the
final `RESULT:` line. I keep it opt-in: without the switch the script asks, as
the steps below describe, and `tools\Certify-EXCEL_UI.cmd` never adds it.

- The script sets `Application.AutomationSecurity` to
  `msoAutomationSecurityLow` on the Excel instance it starts, before that
  instance opens the workbook, so the macros of the workbook built from the
  checked-out commit run without the **Enable Content** message. Without the
  switch the instance follows the Trust Center (`msoAutomationSecurityByUI`).
- `host.json` records `automation_security` as `msoAutomationSecurityLow
  (-Unattended)`, or as `msoAutomationSecurityByUI`; the validator accepts
  those two values and no other. The evidence therefore says how the macros
  were enabled.
- It changes nothing but the setting of that instance, which ends when the
  instance quits. It writes nothing to the Trust Center, the registry or a
  Trusted Location, and it touches no other Excel instance.
- It never prompts. The enable-macros step polls every 2 seconds for at most
  60 seconds until the workbook macros answer, then fails with a message that
  gives the last error. The compile step waits at most 30 seconds for
  **Debug → Compile VBAProject** to grey out, then fails with a message that
  says the compile could not be confirmed. Excel's own retry of a call it
  rejects as busy (up to 60 seconds per call) still applies, and a modal
  message that Excel shows can hold a call until someone closes it.

To certify attended, run the script without `-Unattended`.

## 🔁 What it does

The script prints each step as it runs and records it in `session.log` and
`session.json`.

1. **Preflight.** Reads the full SHA of `HEAD` and refuses a working tree with
   changes to tracked files. Creates the bundle folder
   `EXCEL_UI-certification-<sha12>-<yyyyMMdd-HHmmss>` next to the checkout.
   Reads the Windows version and the Trust Center settings without changing
   them, and stops if VBA project access is off. Writes `source-digests.json`
   from the commit.
2. **Build.** Runs `tools\Sync-EXCEL_UI-VBA.ps1 -Mode Development -NewWorkbook`
   to create `EXCEL_UI-development.xlsm` in the bundle with the modules of the
   `development` configuration of `tools/vba-projects.json`.
3. **Open.** Starts a visible Excel instance, opens the workbook with the
   Trust Center macro setting (with `-Unattended`, with the macros of the
   workbook enabled for that instance), and writes `host.json`.
4. **Enable macros.** Calls a read-only public function of the workbook. If the
   macros are disabled, click **Enable Content** in the Excel message bar and
   press Enter in the script window. Type `Q` to stop, or `F` if Excel shows a
   compile error instead. With `-Unattended` it polls for at most 60 seconds
   and fails the step if the macros still do not run.
5. **Compile.** Shows the VBE and runs **Debug → Compile VBAProject**. The
   command turns grey once the project compiles. If it stays enabled, the
   script asks you to compile in the VBE yourself and answer `Y` (no error) or
   `N` (a compile error, which fails the run); with `-Unattended` it fails the
   step after 30 seconds instead of asking. It then records the workbook's
   VBA components and hides the VBE.
6. **Certify.** Runs `Test_EXCEL_UI_RunReleaseCertification` with
   `Application.Run`. Do not use Excel until it returns. A run that does not
   pass raises an error after it has written its evidence; if Excel shows a
   run-time error message, choose **End**. The script records the error and
   carries on collecting.
7. **Collect.** Copies the JSON document and the text report the run wrote to
   `%TEMP%` (`EXCEL_UI_certification_<yyyymmdd_hhnnss>.json` and `.txt`) into
   the bundle. Only files whose names and timestamps are not older than the
   start of the run are considered, and exactly one of each must exist. The
   originals stay in `%TEMP%`.
8. **Close.** Closes the workbook without saving, closes without saving any
   workbook a failed run left open, and quits the Excel instance it started.
   COM references that PowerShell has not yet released can keep that instance
   in memory after it quits, so a process of that instance still running 15
   seconds later is ended. A process is ended only when its id, its name and
   the start time recorded when the instance was started all match, so no
   other Excel process is touched, even one that reuses the id.

Finally it writes `session.json` and runs the validator.

| Exit code | Meaning |
| --- | --- |
| `0` | Every step passed, certification reported `PASS` and `COMPLETE`, and the bundle validated (or Python was not found). |
| `1` | A step failed, certification did not pass, or the bundle did not validate. |
| `2` | You cancelled the run. |

Keep a failed bundle as it is: it records what happened.

## 🗂️ The bundle

| File | Written by | Content |
| --- | --- | --- |
| `session.json` | the script | Commit SHA, whether the working tree had changes, start and end times in UTC and local time, each step with its status and detail, the workbook's VBA components, the SHA-256 of each collected evidence file, the result and the exit code of the run. |
| `host.json` | the script | Windows caption, version, build and architecture; Excel version, build, `Application.OperatingSystem` and bitness (read from the Excel process); the Click-to-Run version; the macro setting and VBA project access as found; and how the macros were enabled, `msoAutomationSecurityByUI` or `msoAutomationSecurityLow (-Unattended)`. |
| `source-digests.json` | the script | For each module of the `development` configuration at the commit: its path, module name, SHA-256 of the commit's content with CRLF line endings as a checkout writes it, and SHA-256 of the file the build imported. |
| `certification.json` | the workbook | The certification document: Excel version and build, bitness, unit counts and results, skips, cleanup with every finding, every unverified surface and the number of comparisons made, the case inventory (the manifest version, the expected cases, the ordered case results and every inventory finding), whether every unit and skip was recorded, and the verdict. |
| `certification.txt` | the workbook | The text report with its `RESULT:` line, a `CLEANUP` line per cleanup finding, an `UNVERIFIED` line per surface cleanup could not compare, a `CASES` line with the inventory counts, an `EXPECTED` line per expected case, a `CASE` line per case that started and an `INVENTORY` line per finding, and a `RECORDING` line when a unit, skip or finding could not be recorded. |
| `session.log` | the script | Everything the script printed, including the synchronizer's output. |
| `EXCEL_UI-development.xlsm` | the synchronizer | The workbook that ran. It is not part of the validation. |

`session.log` contains local paths, such as your `%TEMP%` folder, which can
include your user name. Review it before you publish it.

## 📐 The evidence policy

`.github/excel-evidence-policy.json` states the evidence contract in one place.
The script takes from it the names of the files it writes, the steps it
records and their order, the certification runner, the configuration and the
schema, and stops before it starts anything when the file is missing or
incomplete. The validator reads the whole file and checks it before it reads a
bundle: each entry must have the right type, no list may be empty, and no step
or file may appear twice. An unusable policy stops the validator with exit
code `2`.

| Key | Content |
| --- | --- |
| `description` | What the file is for. |
| `schema_version` | The format of the policy file itself. The validator accepts `1`. |
| `evidence_schema` | The `schema` value that `session.json`, `host.json`, `source-digests.json` and `certification.json` carry. |
| `component` | The `component` that `certification.json` names. |
| `manifest` | The project manifest, `tools/vba-projects.json`. |
| `configuration` | The configuration the workbook is built from, `development`. |
| `certification_module` | The module that holds the certification runner. |
| `certification_runner` | The macro the certify step runs, `Test_EXCEL_UI_RunReleaseCertification`. Without a `case_manifest`, its `TST_CertRunUnit` calls are the units a bundle must report. |
| `case_manifest` | Optional. The manifest of mandatory certification units and cases, `tools/certification-manifest.json`. When it is present, the units and cases a bundle must report come from the manifest at the commit, and a bundle without the case inventory is rejected. A policy without it describes the earlier contract, which had no case inventory. |
| `required_files` | The bundle file for each role: the session record, the host record, the source digests, the certification document and the text report. |
| `evidence_files` | The files the workbook writes and the collect step copies. `session.json` records the SHA-256 of each. |
| `steps` | The steps the script records, in the order it runs them. |
| `host_fields` | The `windows` and `excel` fields of `host.json` that must be non-empty text. |
| `excel_matches_certification` | Each Excel field of `host.json` and the `certification.json` key it must equal. |
| `bitness` | The Excel bitness values the validator accepts. |

To change the contract, change the policy and, in the same pull request, the
code and the text that implement it: a new step needs its own block in the
script, an entry in its help and one in the list under **What it does**; a
renamed file needs its row in the bundle table. `tests/test_excel_evidence_policy.py`
fails when the script, the validator or this guide disagrees with the policy,
so the mismatch shows up in CI instead of during an Excel run.

## 🔍 Validate a bundle

```bash
python3 tools/check_excel_evidence.py <bundle> --commit <full-sha>
```

`--commit` defaults to `HEAD`. The validator reads the policy of the checkout
it runs from; `--policy <file>` names another one. It prints one line per check
and a final `RESULT: PASS` or `RESULT: FAIL`, and exits with `0` when every
check passes, `1` when one fails, and `2` when it cannot run (no Git, an
unknown commit, no bundle folder or an unusable policy). It checks that:

- every file it needs is present;
- `session.json` names the commit and a clean working tree, every step passed,
  the run ended `PASS` with exit code `0`, and the collected evidence files
  still have the digests recorded when they were collected;
- `source-digests.json` equals the digests of the commit's `development`
  modules, and each imported file had the commit's content;
- the workbook held exactly those modules as standard modules, and no class
  module or form;
- `certification.json` reports `PASS` and `COMPLETE`, no failed or skipped
  unit, cleanup `OK`, recording `OK`, and the units of the commit's case
  manifest (or, under a policy without one, the units that the commit's
  `Test_EXCEL_UI_RunReleaseCertification` runs). A document without a
  `recording` field comes from a runner that did not report it, and is read
  as recording `OK`. A document that itemizes cleanup (`cleanupFindings`,
  `cleanupUnverified` and `cleanupChecked`, which come together) must hold no
  finding and have made at least one comparison when cleanup is `OK`, and must
  number its findings and unverified surfaces from 1, each with a code and a
  detail. A document without those fields comes from a runner that did not
  itemize cleanup, and is read as it was;
- under a policy that names a case manifest, `certification.json` carries the
  case inventory: `manifestVersion` equals the commit's manifest, `expectedCases`
  are its cases in order, `caseResults` lists each of them once, in order, as a
  case that started and passed, `inventory` is `OK`, and `inventoryFindings` is
  empty. A document that passes the unit counters and lacks the inventory is
  rejected, because the counters alone do not show that the required cases ran;
- the `RESULT:` line, the unit lines, any `RECORDING` line, the `CASES`,
  `EXPECTED`, `CASE` and `INVENTORY` lines and, for a document that itemizes
  cleanup, the `CLEANUP` and `UNVERIFIED` lines of `certification.txt` agree
  with `certification.json`;
- `host.json` names Windows and the Excel version, build and bitness, agrees
  with the Excel version, build, operating system and bitness that
  `certification.json` reports, and records `automation_security` as
  `msoAutomationSecurityByUI` or `msoAutomationSecurityLow (-Unattended)`, and
  no other value;
- `certification.json` was written during the recorded session.

To see the source digests a bundle needs for a commit:

```bash
python3 tools/check_excel_evidence.py --inventory --commit <full-sha>
```

## 🧾 What to record

In the pull request, or in the release [evidence record](../RELEASING.md#evidence-record),
record:

- the commit SHA and the validator's `RESULT: PASS` for that SHA;
- the `RESULT:` line of `certification.txt`;
- Excel version, build and bitness, and the Windows version, from `host.json`;
- the date of the run, from `session.json`;
- where the bundle is kept.

I accept a bundle as release evidence only when the validator passes it for
the exact commit being released. A bundle for any other commit, including the
pull-request head of a release that merges into a new commit, does not certify
the release.

<a id="the-case-inventory"></a>

## 🧭 The case inventory

A verdict such as `units=5 failed=0 skipped=0 cleanup=OK` proves only the
outcomes of the units that were dispatched. It does not prove that the cases a
release needs were present, unique and executed: a case can exist in source, run
in a focused pack and never appear in the release verdict. Certification
therefore keeps two separate parts of the verdict, the unit counters and the
case inventory, and a run is complete only when both hold.

[`tools/certification-manifest.json`](../tools/certification-manifest.json) is
the release contract. It lists the units in the order a run executes them, and
the mandatory cases of each. A case identifier is the name of the procedure that
runs it. The manifest also lists the cases a release must never lose
(`required_cases`: the five of [#42](https://github.com/danielep71/VBA-EXCEL_UI/issues/42),
the self-test and the ownership checks), the cases that exist but are deliberately
not mandatory (`focused_cases`, none today) and the interactive runners
(`interactive_runners`). #42 names `TST_Case_TitleBarSameStyleHandleReuse`; #32
replaced that case, so the manifest carries
`TST_Case_TitleBarSameHwndGenerationReuse` under that requirement.

| Unit | Execution | Cases |
| --- | --- | --- |
| `CertificationSelfTest` | isolated | `Test_EXCEL_UI_RunCertificationSelfTest` |
| `OwnershipCleanupChecks` | isolated | `Test_EXCEL_UI_RunOwnershipCleanupChecks` |
| `RegressionPack` | live | the 45 cases of `TST_RunRegressionPack` |
| `SnapshotIdentity` | live | `Test_EXCEL_UI_RunSnapshotIdentity` |
| `TitleBarSdiIdentity` | live | the six cases of `Test_EXCEL_UI_RunTitleBarSdiIdentity` |

**One list, proved equal.** The VBA cannot read the manifest, so the
certification module registers the same list in `TST_CertManifestCase`, runs each
case from `TST_CertDispatchCase` and reaches it through one `TST_CertRunCase`
call. The gate's `certification manifest` check (`tools/certification_manifest.py`)
fails when the registry, the dispatch, the call sites or the manifest disagree:
a case deleted from one place, dispatched twice, run twice, run outside the
wrapper, or added without being listed. Every `TST_Case_` procedure and every
public `Test_EXCEL_UI_` runner must be classified in the manifest, so a new one
needs a reviewed decision. Its fixtures break each of these in turn and expect
the gate to fail. It reads source text; it does not show that a case ran.

**Recording.** `TST_CertRunCase` records a start and exactly one terminal status
under the identifier of the case, and raises the failure of a failed case
unchanged after recording it. Outside a certification run it records nothing, so
the interactive runners behave as they did and are not release certification. At
the end of the run certification stores the expected cases in the record and
compares them exactly with the cases that ran. An expected case that never
started is `missing`, one that recorded a status without a start is `noStart`,
one that started more than once or recorded more statuses than starts is
`duplicate`, one that started and recorded fewer statuses is `noResult`, and a
case the manifest does not list is `unexpected`. Any of them makes the verdict
`FAIL | INCOMPLETE`, as does a comparison that could not run, whatever the unit
counters say.

**Isolated units.** `Test_EXCEL_UI_RunCertificationSelfTest` calls certification
to test a rejection path, and `Test_EXCEL_UI_RunOwnershipCleanupChecks` runs
nested certifications under run faults. A nested call clears the record and the
active flag, so inside a live run it would corrupt the accounting of the run that
executes it. Certification therefore runs both once, first, while no
certification is active, holds their outcome in local variables, and records
them once the live record is open. Each appears exactly once in the expected and
the observed cases. Both remain public runners that can be run alone, and
`tools\Check-EXCEL_UI.cmd` still runs them that way.

The entry baseline of the cleanup comparison is captured before either of them
runs, so that what they change is judged against the state the caller had. For
example, `Test_EXCEL_UI_RunSnapshotIdentity` leaves its own workbook active; with
another workbook active on entry, a baseline taken afterwards would call that
the entry state. What the capture records is held in a record of its own and
carried into the live record after it is cleared. What the isolated units leave
behind is compared with the baseline as soon as they return, before the live
units can change it again, and is reported as a cleanup finding whose code
begins with `isolated.`. The entry window is then made active again, so the live
units start where the caller was, and the comparison at the end of the run
reports anything the run itself leaves.

**Failures reach the caller.** `Test_EXCEL_UI_RunOwnershipCleanupChecks` forces
four runs of the real certification call, each under a one-shot run fault, and
asserts what the call raises and what it records: a cleanup failure (case E), a
failed mandatory unit with an exact inventory (case F), an inventory with one
missing case while every unit passes (case G) and a change left by an isolated
unit, reported as `isolated.app.DisplayStatusBar` although the run puts the host
back (case H). The error carries the number, the
source, the verdict, the failed units and the inventory findings, and no
evidence file is written for these runs.

**Evidence and versions.** `manifestVersion` in the document must equal
`manifest_version` of the manifest at the commit, and `TEST_CERT_MANIFEST_VERSION`
in the harness must equal both; change all three together when the list of cases
changes. The repository policy names the manifest, so a bundle without a case
inventory is rejected from now on. A bundle collected before the manifest existed
validates only under the policy of its own commit, which has no `case_manifest`,
by `--policy`.

## 🪪 Title-bar refusal coverage

A title-bar refusal is correct only if nothing reaches a frame. Each case
below records the complete `GWL_STYLE` of the frame that is active when the
refused call runs, through the harness WinAPI helpers, and fails unless the
value afterwards is identical. The caption bit, and with it title-bar
visibility, is part of that value. Where the captured state and the active
frame would otherwise agree, the case first changes the active frame through
the harness, so that a redirected write would show.

The table maps the identity scenarios of
[#45](https://github.com/danielep71/VBA-EXCEL_UI/issues/45) to the cases that
exercise them and to the certification unit that dispatches those cases. It
describes source coverage at the commit you check out: a row says that a case
exists and that `Test_EXCEL_UI_RunReleaseCertification` runs it, not that it
passed.

| Scenario | Refusal in the source | Cases and assertions | Certification unit |
| --- | --- | --- | --- |
| Active window changed after capture | None for the title bar: restore writes the captured frame through its retained handle, and the active frame must not change. The Ribbon is refused (`captured Ribbon window is not the active window`). | `TST_Case_TitleBarSdiRestoreTargetsCapturedFrame` (`.AnchorRestored`, `.SecondFrameUntouched`, active frame hidden first); `TST_Case_RibbonRestoreRefusesOtherActiveWindow` (`.SecondFrameUntouched`) | `TitleBarSdiIdentity` |
| Closed captured window | The captured hWnd no longer names a window (`captured title-bar window is no longer open`) | `TST_Case_TitleBarCapturedFrameClosedIsReported` (`.Reported`, `.SurvivorFrameUntouched`); `TST_Case_RibbonRestoreRefusesClosedWindow` (`.SurvivorFrameUntouched`) | `TitleBarSdiIdentity` |
| Object/handle mismatch at capture or show/hide | `UI_TryGetActiveFramePair`: `Application.hWnd` disagrees with the active Window's `Window.hWnd`, reached through `UI_InternalInjectFramePairFault` | `TST_Case_ActiveFramePairRefusesMismatch` (helper and wrapper only; no frame is compared) | `RegressionPack` |
| | | `TST_Case_TitleBarPairRefusalLeavesFrameUntouched` (`.Set`, `.FireAndForget`, `.Capture`, `.RestoreEntryFrame`, `.RestoreInvertedFrame`) | `TitleBarSdiIdentity` |
| Object/handle mismatch at restore | The retained Window no longer responds, or its `Window.hWnd` differs from the captured hWnd | No seam reaches these branches. `TST_Case_TitleBarClosedFrameReopenedIsRefused` reaches one of them only when Windows gives the closed frame's handle to the new window, and logs `handle reused=True` when it does. Otherwise this is source coverage only. | `TitleBarSdiIdentity`, when the handle is reused |
| Close and reopen | The liveness check, or the retained-Window checks when the handle is reused | `TST_Case_TitleBarClosedFrameReopenedIsRefused` (`.Reported`, `.ReplacementFrameUntouched` with the new active frame hidden first, `.SurvivorFrameUntouched`) | `TitleBarSdiIdentity` |
| Restore without a snapshot | `NoSnapshot` | `TST_Case_SnapshotResetResultNoSnapshot`, `TST_Case_ResetWithoutSnapshot_NoOp` (`.FrameUntouched`) | `RegressionPack` |

Two restore refusals cannot be reached from any public entry point:
`no captured title-bar frame identity` and
`captured title-bar Window identity is unavailable`. Capture records the title-bar state as known only together with
a live handle and its Window, so these branches guard a state that capture
never produces. They have source coverage only.

Execution evidence is recorded per certification run, not here. A unit passes
only when every case it dispatches ran and passed, because each case raises on
its first failed assertion and the unit stops there. A validated bundle whose
`certification.json` reports a unit as passed is therefore execution evidence
for every row that names that unit, for that commit and that host only. The
Immediate window holds each case's `PASS` or `FAIL` line and the `OBSERVED`
line of the close-and-reopen case. The bundle does not collect them, so record
them with the bundle when you need them; `tools/Check-EXCEL_UI.ps1` writes the
lines of each runner it runs, such as `Test_EXCEL_UI_RunTitleBarSdiIdentity`,
to that runner's log ([EXCEL_CHECKS.md](EXCEL_CHECKS.md)). The Excel builds and bitnesses that a
release must certify are governed by
[#51](https://github.com/danielep71/VBA-EXCEL_UI/issues/51).

## 🧱 Limits

- The validator binds the files to each other and to the commit. It cannot
  prove that Excel wrote them. Read `certification.txt` before you rely on a
  bundle, as you would read the Immediate window.
- One run certifies one host and one Office bitness. I certify releases on
  64-bit Office only: the validator also accepts an `x86` bundle, but a run on
  32-bit Office is not release evidence
  ([#51](https://github.com/danielep71/VBA-EXCEL_UI/issues/51)).
- The inventory proves that each mandatory case started and recorded one status
  in a run. It does not prove that a case asserts enough, and it is as good as
  the manifest: a case that is in no list is in no verdict.
- The script runs only `Test_EXCEL_UI_RunReleaseCertification`. The manual
  checks of [RELEASING.md](../RELEASING.md) and
  [the regression section of the README](../README.md#regression-testing)
  still apply.
- Evidence is collected from the temporary folders the script can see:
  `%TEMP%`, `%TMP%` and the user's own `TEMP` and `TMP`. An Excel instance
  started with another temporary folder writes evidence the script cannot
  find; the collect step then fails and names the folders it searched.
- Compilation is judged from the state of **Debug → Compile VBAProject**. When
  that state is unclear, your answer decides, and `session.json` records that
  you made it.
