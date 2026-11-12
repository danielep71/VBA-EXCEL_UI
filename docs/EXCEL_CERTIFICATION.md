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
| `certification_runner` | The macro the certify step runs, `Test_EXCEL_UI_RunReleaseCertification`. Its `TST_CertRunUnit` calls are the units a bundle must report. |
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

## 🧱 Limits

- The validator binds the files to each other and to the commit. It cannot
  prove that Excel wrote them. Read `certification.txt` before you rely on a
  bundle, as you would read the Immediate window.
- One run certifies one host and one Office bitness. I certify releases on
  64-bit Office only: the validator also accepts an `x86` bundle, but a run on
  32-bit Office is not release evidence
  ([#51](https://github.com/danielep71/VBA-EXCEL_UI/issues/51)).
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
