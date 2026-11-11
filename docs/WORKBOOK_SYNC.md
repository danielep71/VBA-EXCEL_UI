# Workbook synchronization and source round trip

The scripts in `tools/` build an Excel workbook from the modules of this
checkout and check afterwards that a workbook holds exactly the source of one
commit. The synchronizer imports the modules through Excel; the builder writes
a workbook directly from the source, without Excel, and is
[not yet validated in Excel](#building-a-workbook-without-excel).

| File | Purpose |
| --- | --- |
| [`Sync-EXCEL_UI-VBA.ps1`](../tools/Sync-EXCEL_UI-VBA.ps1) | Imports the modules of one configuration, or of a folder, into an `.xlsm` workbook |
| [`Sync-EXCEL_UI-VBA.cmd`](../tools/Sync-EXCEL_UI-VBA.cmd) | Drag-and-drop launcher; without a workbook it creates a development workbook |
| [`Sync-EXCEL_UI-VBA-Downloads.cmd`](../tools/Sync-EXCEL_UI-VBA-Downloads.cmd) | Drag-and-drop launcher that imports modules from a folder, `Downloads` by default |
| [`Export-EXCEL_UI-VBA.ps1`](../tools/Export-EXCEL_UI-VBA.ps1) | Exports every VBA component of a workbook, with a manifest of SHA-256 digests |
| [`compare_vba_export.py`](../tools/compare_vba_export.py) | Compares an export with the source of one commit |
| [`build_workbook.py`](../tools/build_workbook.py) | Writes the `.xlsm` of one configuration from the source, without Excel |
| [`vba_project.py`](../tools/vba_project.py) | Writes the workbook's VBA project (`vbaProject.bin`) for `build_workbook.py` |
| [`lo_smoke.py`](../tools/lo_smoke.py) | Opens a built workbook in headless LibreOffice and checks its modules |
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
- Python 3.11 or later for `build_workbook.py`, on any platform. It needs
  neither Excel nor VBA project trust. `lo_smoke.py` also needs LibreOffice
  and its Python UNO bridge.

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
| Demo | `src/` and the modules of `demo/` | Building the demo workbook and running the demo cleanup harness |

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

Each configuration lists its modules in import order, which is the order in
which the script imports them, and `module_descriptions` gives each module a
one-line description. The module tables of [INSTALLATION.md](../INSTALLATION.md)
are generated from both; after a change to the manifest, run
`python3 tools/render_docs.py --write`.

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

## Building a workbook without Excel

`tools/build_workbook.py` writes a new macro-enabled workbook for one
configuration directly from the source, with the Python standard library
only:

```sh
python3 tools/build_workbook.py --mode Development test-results/EXCEL_UI_Development.xlsm
```

> [!IMPORTANT]
> The built workbook is not yet validated in Excel. Until a run on a Windows
> host has opened it, compiled it and passed
> `Test_EXCEL_UI_RunReleaseCertification`, I use the synchronizer for
> certification and for every other piece of evidence, and so should you.

| Argument | Meaning |
| --- | --- |
| `--mode` | `Production`, `Development` or `Demo`, in any letter case. Required. |
| `--override DIR` | Replace modules with the files of `DIR`, as described below. |
| `OUTPUT` | The `.xlsm` file to write. Its folder is created, and a file of that name is replaced. |

The workbook holds what a new Excel workbook holds after a full
synchronization with `-NewWorkbook`:

- the modules that the configuration lists, in the manifest's order, each
  with its source, its `Attribute` lines and CRLF line endings;
- one worksheet, `Sheet1`, with the code name `Sheet1`, and the workbook code
  name `ThisWorkbook`. Both document modules hold only `Option Explicit`.

The source rules are the synchronizer's: every listed module must be a `.bas`
file with one valid `Attribute VB_Name`, every `.bas` file in `src/`, `test/`
and `demo/` must be in a configuration, and no two modules may share a name.

The package and its VBA project (`xl/vbaProject.bin`) are written from the
Office file format specifications. The VBA project holds the module source and
no compiled code, so Excel compiles it from source when it opens the workbook.
Nothing drives Excel, so the build needs no **Trust access to the VBA project
object model**, and it runs on any platform.

### Overrides for negative controls

`--override DIR` builds the configuration with some of its modules replaced,
for example by a mutant for a negative control or by modules under review. A
file in `DIR` is identified as with `-SourceDir`: only by the
`Attribute VB_Name` it declares, never by its file name.

- A file whose `VB_Name` is a module of the configuration replaces that module
  in its place.
- A standard module (`.bas`) that the checkout does not have, whose name
  matches one of the manifest's `project_name_patterns` that the repository's
  own module names match (today `^M_`), is added after the configuration's
  modules, in name order.
- When several files carry the same `VB_Name`, the newest is taken; the tool
  names the older copies it ignored.

Where the synchronizer would ignore a file, the builder stops with exit code
1: a file without exactly one valid `VB_Name`, a `.cls` or `.frm` file, or a
`VB_Name` that names no module of the repository and matches no pattern. A
repository module that the configuration does not list, such as a demo module
in a development build, stops it too, because modules of different
configurations are never combined. A folder without any `.bas`, `.cls` or
`.frm` file stops it as well.

```sh
python3 tools/build_workbook.py --mode Development --override test-results/mutant test-results/EXCEL_UI_mutant.xlsm
```

### Reproducible output

The same mode and the same module source give the same bytes, on any platform
and from any checkout: line endings are normalized to CRLF, every part of the
package is stored uncompressed with a fixed time and fixed permissions, and the
project ID and the random values of the VBA project are derived from the
source. The tool prints the SHA-256 of the workbook, so two builds of one
commit can be compared by their digests.

| Code | Meaning |
| --- | --- |
| 0 | The workbook was written. |
| 1 | The source or the override folder cannot be built; the message names the cause and nothing is written. |
| 2 | An error in the arguments, such as an unknown mode or an output that is not an `.xlsm` file. |

### LibreOffice smoke check

`tools/lo_smoke.py` builds the workbook of a mode, opens it in headless
LibreOffice with macro execution off, and checks that the VBA project holds
exactly `ThisWorkbook`, `Sheet1` and the configuration's modules, each with the
code of its source file, and that the code names are those above:

```sh
python3 tools/lo_smoke.py --mode Development
```

Run it with a Python that has LibreOffice's UNO bridge; on Debian and Ubuntu,
that is `/usr/bin/python3` with the `python3-uno` package. Given a workbook
path, it checks that workbook instead of building one. It exits with 0 when the
check passes, 1 when it fails, and 2 when it cannot run, for example when
`soffice` or the UNO bridge is missing. The static-checks workflow runs it for
every mode in a job of its own, which is not a required check.

LibreOffice is not a supported host. The check shows that the package and its
VBA project can be read; it does not show that the code compiles or runs in
Excel.

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
- `tests/test_build_workbook.py` reads built workbooks back with readers of
  its own: the package parts and relationships, the compound file, the
  decompressed module and `dir` streams, and the `PROJECT` stream. It checks
  the code names, `--override`, the refusals and that a rebuild is
  byte-identical.
- `tests/test_lo_smoke.py` checks the comparison of the smoke check, and runs
  it end to end where LibreOffice and its UNO bridge are installed.
