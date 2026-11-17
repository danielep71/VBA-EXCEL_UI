# 🧰 Excel checks in one command

`tools/Check-EXCEL_UI.ps1` runs the Excel checks of the commit you have checked
out on a Windows host with Excel, in one run:

- release certification, through `tools/Certify-EXCEL_UI.ps1`, unchanged;
- each `Test_EXCEL_UI_Run*` regression runner, with one log per runner;
- the demo cleanup harness `Demo_Test_RunCleanupChecks`;
- two timed builds of the demo sheet, the Windows default printer, and the
  demo's title-bar check box hiding and showing the title bar.

It writes everything into one bundle folder, with a `summary.txt` that has one
line per step and a final `RESULT: PASS` or `RESULT: FAIL` line, and exits with
a non-zero code when a step fails. Paste the summary into the pull request.

The certification bundle inside it is the release evidence, as
[EXCEL_CERTIFICATION.md](EXCEL_CERTIFICATION.md) describes. The runner and demo
logs are a record of what the other checks printed, bound to the commit by
`session.json`; no validator checks them.

## ✅ Requirements

The [requirements of certification](EXCEL_CERTIFICATION.md#-requirements)
apply: 64-bit desktop Excel for Windows, **Trust access to the VBA project
object model** turned on, a macro setting that lets you enable a workbook's
macros, Git, a clean checkout, and Python 3.11 or later to validate the
certification bundle. The script reads the Trust Center settings and never
changes them.

Save your work in any open Excel window first. The script starts Excel
instances of its own, one at a time, and closes only those.

## ⚡ Run it

```powershell
git switch --detach <full-sha>
git status
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Check-EXCEL_UI.ps1
```

or double-click `tools\Check-EXCEL_UI.cmd`, which runs the same script and
keeps the window open at the end.

| Parameter | Effect |
| --- | --- |
| `-BundleDir <folder>` | Writes the bundle there instead of the default folder `EXCEL_UI-checks-<sha12>-<yyyyMMdd-HHmmss>` next to the checkout. It must be outside the checkout and new or empty. |
| `-Runners <names>` | Runs these runners, in this order, instead of the default list below. Separate the names with commas. |
| `-IncludeRibbonProbe` | Also runs `Test_EXCEL_UI_RunRibbonSdiProbe`, after the other runners. |
| `-SlowBuildSeconds <seconds>` | A demo build that takes longer than this is a `WARN`. The default is `10`. |
| `-Python <exe>` | The Python 3.11+ executable that the certification script validates its bundle with. |
| `-RepoRoot <folder>` | The repository. By default, the parent of the `tools` folder. |

The default runners are `Test_EXCEL_UI_RunAll`, `Test_EXCEL_UI_RunCore`,
`Test_EXCEL_UI_RunTitleBarOnly`, `Test_EXCEL_UI_RunSnapshotIdentity`,
`Test_EXCEL_UI_RunTitleBarSdiIdentity`,
`Test_EXCEL_UI_RunCertificationSelfTest` and
`Test_EXCEL_UI_RunOwnershipCleanupChecks`.
`Test_EXCEL_UI_RunRibbonSdiProbe` asserts nothing
([RIBBON_SDI_BEHAVIOR.md](RIBBON_SDI_BEHAVIOR.md)), so it runs only when you
ask for it and its step is `INFO` whatever it reports. For example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Check-EXCEL_UI.ps1 -Runners Test_EXCEL_UI_RunCore,Test_EXCEL_UI_RunTitleBarOnly -IncludeRibbonProbe
```

## 🖱️ What you do during the run

- **Certification** asks you to click **Enable Content** for the development
  workbook it builds, and to press Enter in the script window. It may also ask
  you to confirm the compile, as
  [EXCEL_CERTIFICATION.md](EXCEL_CERTIFICATION.md#-what-it-does) describes.
- **The runners** reopen that same workbook. Excel remembers a workbook whose
  content you enabled as a trusted document, so it normally opens with its
  macros enabled. When the Trust Center does not keep trusted documents, for
  example by policy, the script asks for **Enable Content** again.
- **The demo** workbook is new, so the script asks for **Enable Content** once
  more. It may ask you to confirm its compile in the same way.

That is one **Enable Content** click per workbook. At each prompt, type `Q` to
stop the run, or `F` if Excel shows a compile error instead. Do not use Excel
while the runners and the demo checks run: they open, activate and close
windows.

### Avoiding the click

Excel enables the macros of a workbook in a trusted location without asking. To run without any **Enable Content** click, create a folder
for bundles only, such as `D:\EXCEL_UI-checks`, add it in **File → Options →
Trust Center → Trust Center Settings → Trusted Locations** with **Subfolders of
this location are also trusted**, and pass a new subfolder of it as
`-BundleDir`. The script never configures a trusted location: decide it
yourself, because every workbook in that folder then runs its macros without
asking. Do not trust the folder that holds your checkout or your downloads.

## 🔁 What it does

The script prints each step as it runs and records it in `session.log`. A step
that fails does not stop the steps that do not depend on it; a step that
cannot run because one it depends on failed is a `FAIL` that says which.

1. **preflight.** Reads the full SHA of `HEAD` and refuses a working tree with
   changes to tracked files, unknown runner names, and a bundle folder inside
   the checkout or not empty. Creates the bundle folder, reads the Windows
   version and the Trust Center settings without changing them, and stops if
   VBA project access is off. Nothing else runs when the preflight fails.
2. **certification.** Runs `tools\Certify-EXCEL_UI.ps1` as it is, in this
   console, with its bundle in the `certification` folder of this bundle. The
   step passes when that script exits with `0`. Its detail gives the exit code
   and the `RESULT:` line of `certification.txt`.
3. **open-development.** Opens the development workbook of the certification
   bundle, as its `session.json` names it, in an Excel instance of its own
   that follows the Trust Center macro setting, and waits until its macros
   run.
4. **One step per runner.** Runs `TST_RunRunnerWithLog` with
   `Application.Run`. That function runs the runner with the log sink set, so
   that everything the runner prints also goes to `runner-<name>.log`, and
   returns its verdict instead of raising. The step passes when the verdict is
   `RESULT: PASS` and the log ends with it.
5. **close-development.** Closes the workbook without saving, closes any
   workbook a runner left open, and quits the instance. A process of that
   instance still running 15 seconds later is ended only when its id, name and
   start time all match the instance the script started, as in certification.
6. **demo-build.** Creates `EXCEL_UI-demo.xlsm` in the bundle with
   `tools\Sync-EXCEL_UI-VBA.ps1 -Mode Demo -NewWorkbook` and writes its output
   to `demo-build.log`.
7. **demo-open** and **demo-compile.** Opens the demo workbook in another
   Excel instance of its own, waits for **Enable Content**, and runs
   **Debug → Compile VBAProject**.
8. **demo-cleanup-checks.** Runs `Demo_Test_RunCleanupChecks` through
   `Demo_Test_RunLogged`, which writes its scenario lines and its
   `DEMO CLEANUP CHECKS:` summary to `demo-cleanup-checks.log`.
9. **demo-create-sheet-1** and **demo-create-sheet-2.** Run
   `Demo_CreateDemoSheet` twice and record how long each run took. A run that
   raises is a `FAIL`; a run slower than `-SlowBuildSeconds` is a `WARN`.
10. **default-printer.** Records the name, the **Work offline** setting, the
    status and whether it is a network printer, of the Windows default
    printer, from `Win32_Printer`. Excel lays sheets out through that
    printer's driver, so a slow or offline printer explains a slow build. The
    step is `INFO`.
11. **demo-titlebar.** Runs the **CLEAR ALL** action, selects only the
    **Title bar** check box (`chkTitleBar`) of the demo sheet, runs the
    **HIDE SELECTED UI** action and then the **SHOW SELECTED UI** action, and
    reads the title bar back after each through `Demo_Test_TitleBarState`,
    which calls `UI_TryGetTitleBarVisible`. The step passes when the title
    bar reads `HIDDEN` after HIDE and `VISIBLE` after SHOW. Each read is a
    `CHECK` line in `demo-titlebar.log`.
12. **close-demo.** Closes the demo workbook without saving and quits that
    instance, as in step 5.

If the run stops on an error or you cancel it, the script still closes the
Excel instances it started, without saving, and writes the summary.

## 🗂️ The bundle

| File | Content |
| --- | --- |
| `summary.txt` | The commit, the Excel and Windows versions, one line per step and the `RESULT:` line. |
| `session.json` | The commit; start and end times in UTC and local time, and the duration; the runners and the `-SlowBuildSeconds` threshold; Windows, the Excel version, build and bitness, and the Trust Center settings as found; the certification exit code and `RESULT:` line; the default printer; each step with its status, detail, log, duration and time; the result and the exit code. |
| `session.log` | Everything the script printed, apart from the certification script's own output, which is in `certification/session.log`. |
| `certification/` | The certification bundle, which `tools/check_excel_evidence.py` validates. |
| `runner-<name>.log` | What the runner printed, then its verdict line. |
| `demo-build.log` | The output of the synchronizer. |
| `demo-cleanup-checks.log` | The scenario lines and the summary of `Demo_Test_RunCleanupChecks`, then its verdict line. |
| `demo-create-sheet-1.log`, `demo-create-sheet-2.log` | The start and verdict lines of each build. |
| `demo-titlebar.log` | The verdict of each demo action, the `CHECK` lines of the title-bar reads, and the step's `RESULT:` line. |
| `EXCEL_UI-demo.xlsm` | The demo workbook that ran. |

The logs and `summary.txt` are UTF-8. `session.log`, `session.json` and the
certification bundle contain local paths, which can include your user name,
and the printer name can include a server name. Review them before you publish
them.

## 📄 The summary

`summary.txt` has one line per step: the status, the step name, the detail and,
when the step wrote one, its log in brackets. A run on a host where everything
passes looks like this, with your commit, versions, timings and printer:

```text
Commit: <full-sha>
Host: Excel 16.0 build 20430, x64; Microsoft Windows 11 Pro 10.0.26100 build 26100.4349
PASS preflight: Commit <full-sha>; Microsoft Windows 11 Pro 10.0.26100 build 26100.4349
PASS certification: exit 0; RESULT: PASS | COMPLETE | units=3 failed=0 skipped=0 cleanup=OK [certification/session.log]
PASS open-development: Excel 16.0 build 20430, x64; process 12345
PASS Test_EXCEL_UI_RunAll: RESULT: PASS [runner-Test_EXCEL_UI_RunAll.log]
PASS Test_EXCEL_UI_RunCore: RESULT: PASS [runner-Test_EXCEL_UI_RunCore.log]
PASS Test_EXCEL_UI_RunTitleBarOnly: RESULT: PASS [runner-Test_EXCEL_UI_RunTitleBarOnly.log]
PASS Test_EXCEL_UI_RunSnapshotIdentity: RESULT: PASS [runner-Test_EXCEL_UI_RunSnapshotIdentity.log]
PASS Test_EXCEL_UI_RunTitleBarSdiIdentity: RESULT: PASS [runner-Test_EXCEL_UI_RunTitleBarSdiIdentity.log]
PASS Test_EXCEL_UI_RunCertificationSelfTest: RESULT: PASS [runner-Test_EXCEL_UI_RunCertificationSelfTest.log]
PASS Test_EXCEL_UI_RunOwnershipCleanupChecks: RESULT: PASS [runner-Test_EXCEL_UI_RunOwnershipCleanupChecks.log]
PASS close-development: Workbook closed without saving; Excel quit
PASS demo-build: EXCEL_UI-demo.xlsm built from the checkout [demo-build.log]
PASS demo-open: Excel 16.0 build 20430, x64; process 23456
PASS demo-compile: Debug > Compile VBAProject completed and is greyed out
PASS demo-cleanup-checks: RESULT: PASS [demo-cleanup-checks.log]
PASS demo-create-sheet-1: Demo_CreateDemoSheet took 2.4 s (threshold 10.0 s) [demo-create-sheet-1.log]
PASS demo-create-sheet-2: Demo_CreateDemoSheet took 2.1 s (threshold 10.0 s) [demo-create-sheet-2.log]
INFO default-printer: Microsoft Print to PDF; work offline False; status Idle; network False
PASS demo-titlebar: before VISIBLE; after HIDE HIDDEN; after SHOW VISIBLE [demo-titlebar.log]
PASS close-demo: Workbook closed without saving; Excel quit
RESULT: PASS
```

| Status | Meaning |
| --- | --- |
| `PASS` | The step passed. |
| `WARN` | A demo build passed but took longer than `-SlowBuildSeconds`. It does not fail the run. |
| `INFO` | A record, not a check: the default printer, and the Ribbon probe. It does not fail the run. |
| `FAIL` | The step failed, or could not run because a step it depends on failed. |
| `CANCELLED`, `NOT_RUN` | You cancelled the run in this step, or before this step. |

| Exit code | Meaning |
| --- | --- |
| `0` | Every step passed, warned or was recorded as `INFO`, and the summary ends with `RESULT: PASS`. |
| `1` | A step failed, or the run stopped on an error. The summary ends with `RESULT: FAIL`. |
| `2` | You cancelled the run, here or in the certification script. The summary ends with `RESULT: FAIL`. |

## 🧾 The log sink

The regression runners and the demo cleanup harness print only to the
Immediate window, which a script cannot read. Each has an opt-in log sink:

- In the development workbook, `TST_Log` of `M_EXCEL_UI_TEST_HARNESS` prints
  every line of the runners. While `TST_LogSinkStart` has named a file, it also
  appends the same line to that file.
- In the demo workbook, every line of `M_DEMO_CLEANUP_TESTS` goes through
  `DEMO_CT_Print`, which does the same while `Demo_Test_RunLogged` runs.

With no sink set, the Immediate window shows exactly what it showed before.
Each line is written in UTF-8 with a CR LF line break, and the file is opened
and closed for every line.

`TST_RunRunnerWithLog` and `Demo_Test_RunLogged` set the sink, run one runner
or demo macro, catch any error it raises, and print and log one verdict line
last: `RESULT: PASS`, or `RESULT: FAIL | Error <n> | Source: <source> |
<description>`. They return the same line. Catching the error keeps Excel from
stopping on a run-time error message during an unattended run.

Writing the log never fails or interrupts a runner. The first write that fails
stops writing for that run, and the verdict the function returns ends with
`| LOG WRITE FAILED: <reason>`. The script then records the step as a `FAIL`
whose detail keeps the runner's own verdict.

## 🧱 Limits

- Only the certification bundle is release evidence, and only once
  `tools/check_excel_evidence.py` passes it for the exact commit. The other
  steps do not replace the manual checks of [RELEASING.md](../RELEASING.md).
- Only `M_EXCEL_UI_TEST_HARNESS` and `M_DEMO_CLEANUP_TESTS` print through the
  sink. The demo actions report a failure with `Debug.Print` in
  `M_EXCEL_UI_DEMO`, and the production modules report one with
  `UI_RuntimeLogFailure`; those lines stay in the Immediate window. The
  title-bar step reads the result back instead.
- The title-bar step reads the frame Excel reports as active, as the demo's
  HIDE and SHOW actions target it. It does not check the other demo selections,
  which it clears first.
- The demo build timings include the time Excel spends in the printer driver
  and on a busy desktop. Compare them on the same host.
- VBA opens files by a name in the Windows code page. When the bundle path
  holds a character outside it, the workbooks cannot write their logs, and each
  runner step fails with `LOG WRITE FAILED`. Pass a `-BundleDir` whose path
  has none.
- Not a watchdog: a runner or demo macro that blocks inside Excel blocks the
  script too.
