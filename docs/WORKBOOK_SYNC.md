# Workbook synchronization and source round trip

The scripts in `tools/` build an Excel workbook from the modules of this
checkout and check afterwards that a workbook holds exactly the source of one
commit:

| File | Purpose |
| --- | --- |
| [`Sync-EXCEL_UI-VBA.ps1`](../tools/Sync-EXCEL_UI-VBA.ps1) | Imports the modules of one configuration, or of a folder, into an `.xlsm` workbook |
| [`Sync-EXCEL_UI-VBA.cmd`](../tools/Sync-EXCEL_UI-VBA.cmd) | Drag-and-drop launcher; without a workbook it creates a development workbook |
| [`Sync-EXCEL_UI-VBA-Downloads.cmd`](../tools/Sync-EXCEL_UI-VBA-Downloads.cmd) | Drag-and-drop launcher that imports modules from a folder, `Downloads` by default |
| [`Export-EXCEL_UI-VBA.ps1`](../tools/Export-EXCEL_UI-VBA.ps1) | Exports every VBA component of a workbook, with a manifest of SHA-256 digests |
| [`compare_vba_export.py`](../tools/compare_vba_export.py) | Compares an export with the source of one commit |
| [`vba-projects.json`](../tools/vba-projects.json) | The module set of each configuration |

The scripts change VBA code only. Compiling the workbook and running
`Test_EXCEL_UI_RunReleaseCertification` in Excel remain the authoritative
checks of behaviour.

## Requirements

- Windows with desktop Excel. The PowerShell scripts drive Excel through
  COM; the launchers run them with Windows PowerShell (`powershell.exe`).
- **Trust access to the VBA project object model** enabled: Excel → File →
  Options → Trust Center → Trust Center Settings → Macro Settings. Without
  it the scripts stop with a message that names this setting, and the
  workbook is not changed.
- A VBA project that is not password-locked.
- The workbook closed in every Excel window. A workbook that opens read-only
  stops the synchronization.
- Python 3.11 or later and Git for `compare_vba_export.py`. It does not need
  Excel, so you can run it on any machine that has the export folder and a
  clone of the repository.

## Quick start

Close the workbook and drag its `.xlsm` file onto `tools\Sync-EXCEL_UI-VBA.cmd`,
or run the launcher with the workbook's path:

```bat
tools\Sync-EXCEL_UI-VBA.cmd "C:\work\EXCEL_UI.xlsm"
```

The launcher asks for the configuration, synchronizes the workbook, shows the
result and waits for a key.

Started without a workbook, for example by double-clicking it, the launcher
creates a blank development workbook, `test-results\EXCEL_UI_Development.xlsm`,
and imports the modules of `src/` and `test/` into it
(`-Mode Development -NewWorkbook`). The file is replaced on every such run,
so its modules always come from the checked-out commit. `test-results/` is
ignored by Git. Open the workbook, run **Debug → Compile VBAProject**, and
run the regression suite or the release certification from there.

To run the synchronizer yourself, start it through `powershell.exe` with
`-ExecutionPolicy Bypass`, as the launchers do. Where the execution policy
disables scripts, as on many managed machines, `.\tools\Sync-EXCEL_UI-VBA.ps1`
alone is refused; the bypass applies to that one process and changes no
setting.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-EXCEL_UI-VBA.ps1 -WorkbookPath "C:\work\EXCEL_UI.xlsm" -Mode Production
```

| Parameter | Meaning |
| --- | --- |
| `-WorkbookPath` | The `.xlsm` workbook. Required. |
| `-Mode` | `Production`, `Development` or `Demo`. Asked when omitted. |
| `-NewWorkbook` | Create a blank `.xlsm` at `-WorkbookPath` first, replacing a file of that name. |
| `-SourceDir` | Import the modules found in this folder instead of a configuration. |
| `-Force` | With `-SourceDir`, import without asking for confirmation. |
| `-NoBackup` | Do not make the safety backup. |
| `-RepoRoot` | The checkout to read; by default the one that holds the script. |

## Configurations

A full synchronization imports the modules that one configuration of
`tools/vba-projects.json` lists:

| Mode | Modules | Use |
| --- | --- | --- |
| Production | The four modules of `src/` | The package of [INSTALLATION.md](../INSTALLATION.md) |
| Development | `src/` and `test/M_EXCEL_UI_REGRESSION_TESTS.bas` | Running the regression suite and the release certification |
| Demo | `src/` and the two modules of `demo/` | Building the demo workbook |

When `-Mode` is omitted, the script asks; answer `1`, `P` or `production`,
`2`, `D` or `development`, `3` or `demo`. Any other answer cancels with exit
code 2 and the workbook is not changed.

Every module the configuration lists is replaced from the source. A
synchronization never removes a module: modules that the configuration does
not list stay in the workbook. When some of them are repository modules, such
as `M_EXCEL_UI_REGRESSION_TESTS` in a workbook synchronized as production,
the script names them in a warning. For a workbook with exactly one
configuration, start from a new workbook with `-NewWorkbook`.

The manifest is the single list of the three sets. Each `.bas` file in
`src/`, `test/` and `demo/` must be in at least one configuration: an
unlisted module stops the script before Excel starts. The static gate
(`python3 tools/check_repo.py`, check *VBA project manifest*) fails when the
manifest and the tracked modules disagree, so add a new module to the
manifest in the same change.

## Importing modules from a folder

To try modules received outside the repository, for example modules
downloaded while you review a pull request, without copying them into the
working tree, drag the workbook onto `tools\Sync-EXCEL_UI-VBA-Downloads.cmd`
or pass its path as the argument. The launcher asks for the folder (press
Enter for `%USERPROFILE%\Downloads`) and runs the synchronizer with
`-SourceDir`:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-EXCEL_UI-VBA.ps1 -WorkbookPath "C:\work\EXCEL_UI.xlsm" -SourceDir "$env:USERPROFILE\Downloads"
```

Only the modules found in the folder are imported; every other module in the
workbook is left as it is. `-SourceDir` cannot be combined with `-Mode` or
`-NewWorkbook`.

- A file is identified only by the `Attribute VB_Name` it declares, never by
  its file name. A file with no such line, a malformed one, several (even
  identical) ones, or a name that is not a valid VBA name is ignored.
- A file is taken when its `VB_Name` is a module of the repository, with the
  same extension.
- A standard module (`.bas`) that the checkout does not have yet, such as a
  module added by the pull request under review, is also taken when its name
  matches one of the manifest's `project_name_patterns` that the repository's
  own module names match. Today that is `^M_`; the other patterns name
  procedures, not modules. Such a module is listed apart as a new module.
- When several files carry the same `VB_Name`, such as `M_EXCEL_UI.bas` and
  `M_EXCEL_UI (1).bas`, the newest is taken.
- Other files in the folder are ignored.

The script lists the chosen files with their dates and the ignored ones with
the reason, then asks for confirmation; `-Force` skips the question. Answering
anything but `Y` exits with code 2. When no file is usable, the script shows
why and exits with code 2. In both cases the workbook is not changed. The
repository files and the folder are never changed. A later full
synchronization brings the workbook back to the repository source.

## Safety behaviour

- **Checks before any change.** The manifest, the module set, the
  configuration answer and the folder selection are settled before the
  script makes a backup, creates a workbook or starts Excel.
- **VBA only.** Standard modules are removed and imported; no cell, formula,
  name, format, worksheet or other module is written. The script records the
  sheet names and code names before the import and refuses to save when they
  changed. It also refuses, before the first change, a module name that the
  workbook already uses for a sheet, a class or a form.
- **Its own Excel instance.** The script opens the workbook in a hidden
  Excel instance of its own, with events, macro execution
  (`AutomationSecurity`), alerts and external-link updates disabled, and
  calculation set to manual. It puts back the workbook's calculation mode
  before it saves, restores the instance's settings and quits it.
- **CRLF staging.** The modules are copied to a temporary folder with CRLF
  line endings, as the VBE expects, and imported from there. The folder is
  deleted at the end.
- **Safety backup.** Before it opens an existing workbook, the script copies
  it to `<name>.pre-vba-sync.<yyyyMMdd-HHmmss>.xlsm` in the same folder,
  unless you pass `-NoBackup`. The copy is deleted after a successful save
  and kept on failure. It is not a rollback archive. `-NewWorkbook` makes no
  backup, since it replaces the file.
- **No save on failure.** On an error the workbook is closed without saving.

## Exit codes

| Code | Meaning |
| --- | --- |
| 0 | The workbook was synchronized and saved. |
| 1 | Error. The message names the cause; the workbook was not saved. |
| 2 | Cancelled, or nothing to import. The workbook was not changed. |

The launchers print the outcome, wait for a key and return the same code.

## Source round trip

The round trip shows that a workbook holds exactly the source of one commit:
the modules its configuration lists, each with the same code.

1. **Check out the commit** with a clean `git status`.
2. **Synchronize** a new workbook from it, for example with
   `tools\Sync-EXCEL_UI-VBA.cmd` and no argument, or:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-EXCEL_UI-VBA.ps1 -WorkbookPath ".\test-results\EXCEL_UI_Development.xlsm" -Mode Development -NewWorkbook
   ```

3. **Compile and run** in Excel as you need, then save and close the
   workbook.
4. **Export** to a new folder. The workbook is opened read-only and not
   changed:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Export-EXCEL_UI-VBA.ps1 -WorkbookPath ".\test-results\EXCEL_UI_Development.xlsm" -OutDir ".\test-results\export"
   ```

   The folder must be new or empty. Beside the components the script writes
   `export-manifest.json`: the workbook's file name and SHA-256, taken
   before it opens the workbook and checked again after it closes it, and the
   name, type and SHA-256 of every exported file.
5. **Compare** with the source of the commit:

   ```sh
   python3 tools/compare_vba_export.py --export test-results/export --mode development --commit FULL_SHA --workbook test-results/EXCEL_UI_Development.xlsm
   ```

   `--commit` defaults to `HEAD`. The tool reads the source with
   `git show <commit>:<path>`, not from the working tree, so local changes do
   not affect the result. It prints the commit, the number of modules, the
   document modules it did not compare, the number of re-cased lines, the
   binding, one `FAIL` line per difference, and a last line
   `RESULT=PASS` or `RESULT=FAIL` with the counts and the binding. Exit code
   0 means pass, 1 a difference or a failed binding, 2 an error in the
   arguments or the input.

The comparison checks two things.

**Source parity.** The export must hold exactly the modules of the
configuration, with the same attribute lines and the same code. It ignores:

- line endings and trailing blank lines;
- the letter case of code outside string literals and comments. VBA names
  are case-insensitive and the VBE keeps one spelling per name across a
  project, so it re-cases identifiers on import. The lines that differ only
  that way are counted, not failed;
- document modules (`ThisWorkbook` and one per sheet), which Excel creates
  and whose attributes Excel writes. This repository has none, so they are
  not compared with the source; each must hold no code other than
  `Option Explicit`, or the comparison fails.

**Workbook binding.** With `--workbook`, the tool checks the workbook and the
exported files against `export-manifest.json`:

| Binding | Meaning |
| --- | --- |
| `VERIFIED` | The manifest records this workbook, by its bytes, and every file in the folder, unchanged. A renamed copy, as a download often is, gets a `NOTE`. |
| `FAIL` | The workbook, the files and the manifest do not agree, or `--workbook` is not a macro-enabled workbook. The result fails. |
| `UNVERIFIED` | There is no manifest, or no `--workbook`. Only source parity was checked. A manifest that does not match the files still fails. |

A manifest is a record, not a signature: it ties an export to one workbook
file, but it does not attest who made either.

## Limits

- The synchronizer imports standard modules (`.bas`) only. The manifest and
  the static gate keep `document_modules` empty, and a manifest that lists
  document modules is refused.
- It does not remove modules, so a workbook used with several configurations
  keeps the union of their modules. Build a new workbook to get exactly one.
- It needs Windows desktop Excel. The offline tests run the PowerShell
  functions in PowerShell 7 without Excel; the COM steps (opening, importing,
  saving and exporting) are checked only by running the scripts on Windows.
- A successful synchronization says nothing about whether the code compiles
  or behaves; compile it and run the regression suite or the release
  certification.

## Offline tests

The unit tests run on any platform, without Excel:

- `tests/test_sync_script.py` runs the synchronizer's folder selection,
  configuration answer, module sets, manifest completeness and the refusals
  that stop it before Excel in PowerShell 7 (`pwsh`), parses every script in
  `tools/`, and checks the launchers. The PowerShell tests are skipped where
  `pwsh` is not installed.
- `tests/test_compare_vba_export.py` builds exports from the source at
  `HEAD`, damages them the way a faulty round trip would, and checks each
  finding and binding result.
