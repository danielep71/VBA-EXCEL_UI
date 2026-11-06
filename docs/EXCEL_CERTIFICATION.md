# 🧪 Excel certification in one command

`tools/Certify-EXCEL_UI.ps1` certifies the commit you have checked out on a
Windows host with Excel. It builds a development workbook from that commit,
compiles it, runs `Test_EXCEL_UI_RunReleaseCertification`, and collects the
result into an evidence bundle that names the commit, the host and the source
it ran. `tools/check_excel_evidence.py` then validates the bundle against the
commit, on any machine with a clone.

The script automates the compile and certification steps of
[RELEASING.md](../RELEASING.md). It does not replace the manual checks listed
there (hide/show, snapshot/restore, multi-window identity, title-bar behaviour,
cleanup on error paths), and it runs nothing but certification:
`Test_EXCEL_UI_RunRibbonSdiProbe` is a characterization that asserts nothing
(see [RIBBON_SDI_BEHAVIOR.md](RIBBON_SDI_BEHAVIOR.md)) and is not part of the
release gate, so its output is not part of a bundle.

## ✅ Requirements

- Desktop Microsoft Excel for Windows, as in the
  [support baseline](../INSTALLATION.md).
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
keeps the window open at the end.

| Parameter | Effect |
| --- | --- |
| `-BundleDir <folder>` | Writes the bundle there instead of the default folder. It must be outside the checkout and new or empty. |
| `-AllowDirty` | Runs even though the working tree has changes. The bundle records it, and the validator rejects such a bundle. Use it only to try the tooling. |
| `-Python <exe>` | The Python 3.11+ executable for the validation. By default the script tries `py -3`, `python` and `python3`. |
| `-RepoRoot <folder>` | The repository. By default, the parent of the `tools` folder. |

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
   Trust Center macro setting, and writes `host.json`.
4. **Enable macros.** Calls a read-only public function of the workbook. If the
   macros are disabled, click **Enable Content** in the Excel message bar and
   press Enter in the script window. Type `Q` to stop, or `F` if Excel shows a
   compile error instead.
5. **Compile.** Shows the VBE and runs **Debug → Compile VBAProject**. The
   command turns grey once the project compiles. If it stays enabled, the
   script asks you to compile in the VBE yourself and answer `Y` (no error) or
   `N` (a compile error, which fails the run). It then records the workbook's
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
| `host.json` | the script | Windows caption, version, build and architecture; Excel version, build, `Application.OperatingSystem` and bitness (read from the Excel process); the Click-to-Run version; the macro setting and VBA project access as found. |
| `source-digests.json` | the script | For each module of the `development` configuration at the commit: its path, module name, SHA-256 of the commit's content with CRLF line endings as a checkout writes it, and SHA-256 of the file the build imported. |
| `certification.json` | the workbook | The certification document: Excel version and build, bitness, unit counts and results, skips, cleanup and the verdict. |
| `certification.txt` | the workbook | The text report with its `RESULT:` line. |
| `session.log` | the script | Everything the script printed, including the synchronizer's output. |
| `EXCEL_UI-development.xlsm` | the synchronizer | The workbook that ran. It is not part of the validation. |

`session.log` contains local paths, such as your `%TEMP%` folder, which can
include your user name. Review it before you publish it.

## 🔍 Validate a bundle

```bash
python3 tools/check_excel_evidence.py <bundle> --commit <full-sha>
```

`--commit` defaults to `HEAD`. The validator prints one line per check and a
final `RESULT: PASS` or `RESULT: FAIL`, and exits with `0` when every check
passes, `1` when one fails, and `2` when it cannot run (no Git, an unknown
commit or no bundle folder). It checks that:

- every file it needs is present;
- `session.json` names the commit and a clean working tree, every step passed,
  the run ended `PASS` with exit code `0`, and the collected evidence files
  still have the digests recorded when they were collected;
- `source-digests.json` equals the digests of the commit's `development`
  modules, and each imported file had the commit's content;
- the workbook held exactly those modules as standard modules, and no class
  module or form;
- `certification.json` reports `PASS` and `COMPLETE`, no failed or skipped
  unit, cleanup `OK`, and the units that the commit's
  `Test_EXCEL_UI_RunReleaseCertification` runs;
- the `RESULT:` line and the unit lines of `certification.txt` agree with
  `certification.json`;
- `host.json` names Windows and the Excel version, build and bitness, and
  agrees with the Excel version, build, operating system and bitness that
  `certification.json` reports;
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
them with the bundle when you need them. The Excel builds and bitnesses that a
release must certify are governed by
[#51](https://github.com/danielep71/VBA-EXCEL_UI/issues/51).

## 🧱 Limits

- The validator binds the files to each other and to the commit. It cannot
  prove that Excel wrote them. Read `certification.txt` before you rely on a
  bundle, as you would read the Immediate window.
- One run certifies one host and one Office bitness. Certify each environment
  you claim support for with a run of its own.
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
