"""Offline tests for tools/Sync-EXCEL_UI-VBA.ps1 and the other workbook scripts.

The synchronizer's functions are loaded from the script's syntax tree and run
in PowerShell 7 (pwsh), and the whole script is run for the refusals that
stop it before Excel starts. Nothing here starts Excel or touches a real
workbook. The tests are skipped where pwsh is not installed.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import time
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
SCRIPT = TOOLS / "Sync-EXCEL_UI-VBA.ps1"
PWSH = shutil.which("pwsh")

# Defines every function of the script and the variables they read, without
# running the script body: the top-level function definitions and
# assignments, not the try statement that does the work.
LOADER = r"""
$ErrorActionPreference = 'Stop'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($env:SYNC_SCRIPT, [ref]$null, [ref]$null)
foreach ($s in $ast.EndBlock.Statements) {
    if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst] -or
        $s -is [System.Management.Automation.Language.AssignmentStatementAst]) {
        . ([scriptblock]::Create($s.Extent.Text))
    }
}
"""

WRITE_RESULT = "| ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $env:RESULT -Encoding utf8"


def module(name: str) -> bytes:
    return f'Attribute VB_Name = "{name}"\r\nOption Explicit\r\n'.encode("cp1252")


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class SyncScriptTests(unittest.TestCase):
    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)

    def pwsh(self, body: str) -> Any:
        """Run body after the loader; it writes a JSON result to $env:RESULT."""
        result = self.root / "result.json"
        completed = subprocess.run(
            [str(PWSH), "-NoProfile", "-NonInteractive", "-Command", LOADER + body],
            env={**os.environ, "SYNC_SCRIPT": str(SCRIPT), "RESULT": str(result)},
            capture_output=True, text=True, timeout=120, check=False)
        self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
        return json.loads(result.read_text(encoding="utf-8-sig"))

    def failure(self, call: str) -> str:
        """The message of the exception ``call`` throws, or "" when it does not."""
        result = self.pwsh(f"$failed = ''\ntry {{ $null = {call} }} catch {{ $failed = $_.Exception.Message }}\n"
                           "@{ failed = $failed } " + WRITE_RESULT)
        return str(result["failed"])

    def fake_repository(self, configurations: dict[str, list[str]], extra: tuple[str, ...] = (),
                        documents: tuple[str, ...] = ()) -> Path:
        """A repository with the modules the configurations list, and ``extra`` unlisted ones."""
        repo = self.root / "repo"
        for directory in ("src", "test", "demo", "tools"):
            (repo / directory).mkdir(parents=True, exist_ok=True)
        for path in {p for paths in configurations.values() for p in paths} | set(extra):
            (repo / path).parent.mkdir(parents=True, exist_ok=True)
            (repo / path).write_bytes(module(Path(path).stem))
        manifest = {"schema_version": 1, "project_name_patterns": ["^M_", "^UI_"],
                    "document_modules": list(documents),
                    "configurations": {name: {"modules": paths} for name, paths in configurations.items()}}
        (repo / "tools" / "vba-projects.json").write_text(json.dumps(manifest), encoding="utf-8")
        return repo

    def standard_repository(self, extra: tuple[str, ...] = ()) -> Path:
        production = ["src/M_A.bas", "src/M_B.bas"]
        return self.fake_repository({"production": production,
                                     "development": [*production, "test/M_T.bas"],
                                     "demo": [*production, "demo/M_D.bas"]}, extra)

    # Sheet snapshot -----------------------------------------------------------

    def test_a_changed_sheet_is_named_in_the_refusal(self) -> None:
        unchanged = self.failure("Assert-SheetSnapshotUnchanged -Before @('Sheet1|Sheet1') -After @('Sheet1|Sheet1')")
        self.assertEqual(unchanged, "")
        changed = self.failure("Assert-SheetSnapshotUnchanged -Before @('Sheet1|') -After @('Sheet1|Sheet1')")
        self.assertIn("sheet 1 was 'Sheet1|' (name|CodeName) and is now 'Sheet1|Sheet1'", changed)
        count = self.failure("Assert-SheetSnapshotUnchanged -Before @('A|A') -After @('A|A', 'B|B')")
        self.assertIn("sheet count changed from 1 to 2", count)

    # Mode question and component sets ---------------------------------------

    def mode(self, answer: str) -> Any:
        return self.pwsh(f"$m = ConvertTo-SyncMode -Answer '{answer}'\n@{{ mode = $m }} " + WRITE_RESULT)["mode"]

    def test_mode_answer(self) -> None:
        # A number, a letter or a word in any case; anything else cancels.
        for answer, expected in (("1", "Production"), ("p", "Production"), (" PROD ", "Production"),
                                 ("production", "Production"), ("2", "Development"), ("D", "Development"),
                                 ("dev", "Development"), ("Development", "Development"), ("3", "Demo"),
                                 ("DEMO", "Demo"), ("", None), ("y", None), ("4", None), ("pd", None),
                                 ("demo x", None), ("m", None)):
            with self.subTest(answer=answer):
                self.assertEqual(self.mode(answer), expected)

    def manifest_sources(self, repo: Path, mode: str) -> list[str]:
        names = self.pwsh(f"$f = Get-ManifestSources -RepoPath '{repo}' -Mode {mode}\n"
                          "ConvertTo-Json -InputObject @($f | ForEach-Object { $_.FullName }) "
                          "| Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        return [Path(name).relative_to(repo).as_posix() for name in names]

    def test_each_mode_imports_its_configuration(self) -> None:
        def tracked(directory: str) -> set[str]:
            return {path.relative_to(ROOT).as_posix() for path in (ROOT / directory).glob("*.bas")}

        production = self.manifest_sources(ROOT, "Production")
        development = self.manifest_sources(ROOT, "Development")
        demo = self.manifest_sources(ROOT, "Demo")
        self.assertEqual(set(production), tracked("src"))
        self.assertEqual(set(development), tracked("src") | tracked("test"))
        self.assertEqual(set(demo), tracked("src") | tracked("demo"))
        # In the manifest's order, each module once.
        manifest = json.loads((TOOLS / "vba-projects.json").read_text(encoding="utf-8"))
        for mode, files in (("production", production), ("development", development), ("demo", demo)):
            self.assertEqual(files, manifest["configurations"][mode]["modules"])
            self.assertEqual(len(files), len(set(files)))

    def test_a_listed_module_must_exist_and_be_a_standard_module(self) -> None:
        repo = self.standard_repository()
        (repo / "src" / "M_B.bas").unlink()
        self.assertIn("not found: src/M_B.bas",
                      self.failure(f"Get-ManifestSources -RepoPath '{repo}' -Mode Production"))
        shutil.rmtree(repo)
        repo = self.fake_repository({"production": ["src/C_A.cls"]})
        self.assertIn("is not a standard module (.bas): src/C_A.cls",
                      self.failure(f"Get-ManifestSources -RepoPath '{repo}' -Mode Production"))

    def test_a_manifest_with_document_modules_is_refused(self) -> None:
        repo = self.fake_repository({"production": ["src/M_A.bas"]}, documents=("src/ThisWorkbook.cls",))
        self.assertIn("this synchronizer imports standard modules only",
                      self.failure(f"Get-ManifestSources -RepoPath '{repo}' -Mode Production"))

    def test_an_unlisted_module_is_refused(self) -> None:
        repo = self.standard_repository(extra=("demo/M_NEW.bas", "test/nested/M_DEEP.bas"))
        message = self.failure(f"Assert-ManifestComplete -RepoPath '{repo}' "
                               f"-SourceFiles @(Get-RepositorySources -RepoPath '{repo}')")
        self.assertIn("Module not listed in any configuration of tools/vba-projects.json", message)
        self.assertIn("M_NEW.bas", message)
        self.assertIn("M_DEEP.bas", message)
        self.assertNotIn("M_A.bas", message)

    def test_this_repository_lists_every_module(self) -> None:
        result = self.pwsh(f"$u = Get-UnlistedSources -RepoPath '{ROOT}' "
                           f"-SourceFiles @(Get-RepositorySources -RepoPath '{ROOT}')\n"
                           "ConvertTo-Json -InputObject @($u | ForEach-Object { $_.Name }) "
                           "| Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        self.assertEqual(result, [])

    def test_duplicate_repository_names_are_refused(self) -> None:
        repo = self.standard_repository()
        (repo / "demo" / "M_D.bas").write_bytes(module("M_A"))
        self.assertIn("Duplicate VBA module name 'M_A'",
                      self.failure(f"Get-RepositoryNames -SourceFiles @(Get-RepositorySources -RepoPath '{repo}')"))

    def test_new_module_patterns_are_those_naming_modules(self) -> None:
        # Of this repository's patterns only ^M_ matches a module name; the
        # others name procedures.
        manifest = json.loads((TOOLS / "vba-projects.json").read_text(encoding="utf-8"))
        patterns = ", ".join(f"'{p}'" for p in manifest["project_name_patterns"])
        result = self.pwsh(f"$n = @((Get-RepositoryNames -SourceFiles @(Get-RepositorySources -RepoPath '{ROOT}'))"
                           ".Values | ForEach-Object { $_.BaseName })\n"
                           f"$p = Get-NewModulePatterns -Patterns @({patterns}) -ModuleNames $n\n"
                           "ConvertTo-Json -InputObject @($p) | Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        self.assertEqual(result, ["^M_"])

    def test_modules_outside_the_configuration_are_left_in_place(self) -> None:
        result = self.pwsh(
            f"$r = Get-RepositoryNames -SourceFiles @(Get-RepositorySources -RepoPath '{ROOT}')\n"
            "$l = Get-LeftInPlaceModules -WorkbookModules @('Module1', 'M_EXCEL_UI', "
            "'M_EXCEL_UI_REGRESSION_TESTS', 'm_demo_builder') -RepositoryNames $r "
            "-ModeNames @('M_EXCEL_UI', 'M_EXCEL_UI_RUNTIME', 'M_EXCEL_UI_SNAPSHOT', 'M_EXCEL_UI_TITLEBAR')\n"
            "ConvertTo-Json -InputObject @($l) | Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        self.assertEqual(result, ["M_EXCEL_UI_REGRESSION_TESTS", "m_demo_builder"])

    # Folder selection ---------------------------------------------------------

    def select(self, folder: Path) -> dict[str, Any]:
        repo = self.standard_repository()
        stage = self.root / "stage"
        return dict(self.pwsh(
            f"$repo = Get-RepositoryNames -SourceFiles @(Get-RepositorySources -RepoPath '{repo}')\n"
            "$picked = @(); $shown = $null\n"
            f"$out = Select-FolderSources -FolderPath '{folder}' -RepositoryNames $repo "
            f"-NewModulePatterns @('^M_') -StagePath '{stage}' -InformationVariable shown 6>$null\n"
            "foreach ($o in $out) { if ($o -is [System.IO.FileInfo]) { $picked += $o.Name } }\n"
            "$lines = @(foreach ($r in $shown) { [string]$r })\n"
            "@{ picked = $picked; lines = $lines } " + WRITE_RESULT))

    def write(self, folder: Path, file_name: str, header: str, age: int = 0) -> Path:
        path = folder / file_name
        path.write_bytes((header + "Option Explicit\n").encode("cp1252"))
        if age:
            stamp = time.time() - age
            os.utime(path, (stamp, stamp))
        return path

    def test_folder_candidates_need_a_declared_vb_name(self) -> None:
        # The file name never stands in for a missing or invalid VB_Name.
        folder = self.root / "downloads"
        folder.mkdir()
        self.write(folder, "M_A.bas", "")
        self.write(folder, "M_A (1).bas", 'Attribute VB_Name = "M_A"\n')
        self.write(folder, "M_B.bas", "Attribute VB_Name = M_B\n")
        self.write(folder, "M_T.bas", 'Attribute VB_Name = "M_T"\nAttribute VB_Name = "m_t"\n')
        self.write(folder, "M_D.bas", 'Attribute VB_Name = "M D"\n')
        self.write(folder, "Other.bas", 'Attribute VB_Name = "Other"\n')
        self.write(folder, "UI_Helper.bas", 'Attribute VB_Name = "UI_Helper"\n')
        self.write(folder, "M_A.cls", 'Attribute VB_Name = "M_A"\n')
        self.write(folder, "M_NEW.cls", 'Attribute VB_Name = "M_NEW"\n')
        self.write(folder, "M_NEW.bas", 'Attribute VB_Name = "M_NEW"\n')
        self.write(folder, "notes.txt", 'Attribute VB_Name = "M_B"\n')
        result = self.select(folder)
        self.assertEqual(result["picked"], ["M_A.bas", "M_NEW.bas"])
        lines = "\n".join(result["lines"])
        for expected in ("M_A.bas (no Attribute VB_Name line; the file name is not used",
                         "M_B.bas (malformed Attribute VB_Name line)",
                         "M_T.bas (several Attribute VB_Name lines)",
                         "M_D.bas (VB_Name 'M D' is not a valid VBA name)",
                         "Other.bas (VB_Name 'Other' is not a repository module)",
                         "UI_Helper.bas (VB_Name 'UI_Helper' is not a repository module)",
                         "M_A.cls ('M_A' is a .bas module in the repository)",
                         "M_NEW.cls (VB_Name 'M_NEW' is not a repository module)",
                         "New modules, not in the repository checkout",
                         "M_A (1).bas"):
            self.assertIn(expected, lines)
        self.assertNotIn("notes.txt", lines)
        # The staged copy has CRLF line endings; the original is unchanged.
        self.assertEqual((self.root / "stage" / "M_A.bas").read_bytes(),
                         b'Attribute VB_Name = "M_A"\r\nOption Explicit\r\n')
        self.assertEqual((folder / "M_A (1).bas").read_bytes(), b'Attribute VB_Name = "M_A"\nOption Explicit\n')

    def test_the_newest_copy_of_a_module_wins(self) -> None:
        folder = self.root / "downloads"
        folder.mkdir()
        self.write(folder, "M_B.bas", 'Attribute VB_Name = "M_B"\n', age=3600)
        newest = self.write(folder, "M_B (2).bas", 'Attribute VB_Name = "M_B"\n')
        newest.write_bytes(newest.read_bytes() + b"' newest\n")
        self.write(folder, "M_B (1).bas", 'Attribute VB_Name = "M_B"\n', age=7200)
        result = self.select(folder)
        self.assertEqual(result["picked"], ["M_B.bas"])
        lines = "\n".join(result["lines"])
        self.assertIn("M_B.bas (older copy of 'M_B')", lines)
        self.assertIn("M_B (1).bas (older copy of 'M_B')", lines)
        self.assertTrue((self.root / "stage" / "M_B.bas").read_bytes().endswith(b"' newest\r\n"))

    def test_a_folder_without_a_usable_module_selects_nothing(self) -> None:
        folder = self.root / "downloads"
        folder.mkdir()
        self.write(folder, "M_A.bas", "")
        result = self.select(folder)
        self.assertEqual(result["picked"], [])
        # The reason is shown, so the user can see why nothing is imported.
        self.assertIn("M_A.bas (no Attribute VB_Name line", "\n".join(result["lines"]))

    # The whole script, stopped before Excel ---------------------------------

    def run_script(self, *arguments: str, answer: str | None = None) -> subprocess.CompletedProcess[str]:
        command = [str(PWSH), "-NoProfile"] + (["-NonInteractive"] if answer is None else [])
        return subprocess.run([*command, "-File", str(SCRIPT), *arguments], input=answer,
                              capture_output=True, text=True, timeout=120, check=False)

    def assertStops(self, completed: subprocess.CompletedProcess[str], code: int, message: str) -> None:
        self.assertEqual(completed.returncode, code, completed.stdout + completed.stderr)
        self.assertIn(message, completed.stdout + completed.stderr)
        self.assertNotIn("ComObject", completed.stdout + completed.stderr)

    def test_argument_refusals(self) -> None:
        built = self.root / "built" / "EXCEL_UI_Development.xlsm"
        existing = self.root / "EXCEL_UI.xlsm"
        existing.write_bytes(b"not opened")
        for arguments, message in (
                (["-WorkbookPath", str(built), "-NewWorkbook", "-SourceDir", str(self.root)],
                 "Use -NewWorkbook only with a full synchronization, not with -SourceDir."),
                (["-WorkbookPath", str(existing), "-Mode", "Demo", "-SourceDir", str(self.root)],
                 "Use -Mode only with a full synchronization, not with -SourceDir."),
                (["-WorkbookPath", str(self.root / "EXCEL_UI.xlsx"), "-NewWorkbook", "-Mode", "Demo"],
                 "WorkbookPath must point to an .xlsm file."),
                (["-WorkbookPath", str(self.root / "absent.xlsm"), "-Mode", "Demo"], "Workbook not found"),
                (["-WorkbookPath", str(existing), "-Mode", "Testing"], "Testing")):
            with self.subTest(message=message):
                self.assertStops(self.run_script(*arguments), 1, message)
                self.assertFalse(built.parent.exists())
        self.assertEqual(sorted(p.name for p in self.root.iterdir()), ["EXCEL_UI.xlsm"])

    def test_an_unlisted_module_stops_before_excel(self) -> None:
        repo = self.standard_repository(extra=("src/M_NEW.bas",))
        built = self.root / "built" / "EXCEL_UI_Development.xlsm"
        completed = self.run_script("-WorkbookPath", str(built), "-RepoRoot", str(repo),
                                    "-NewWorkbook", "-Mode", "Development")
        self.assertStops(completed, 1, "Module not listed in any configuration")
        self.assertIn("M_NEW.bas", completed.stdout + completed.stderr)
        self.assertFalse(built.parent.exists())

    def test_an_unknown_mode_answer_cancels(self) -> None:
        repo = self.standard_repository()
        book = self.root / "EXCEL_UI.xlsm"
        book.write_bytes(b"unchanged")
        completed = self.run_script("-WorkbookPath", str(book), "-RepoRoot", str(repo), answer="x\n")
        self.assertStops(completed, 2, "Cancelled: the workbook was not changed.")
        self.assertEqual(book.read_bytes(), b"unchanged")
        self.assertEqual(sorted(p.name for p in self.root.iterdir()), ["EXCEL_UI.xlsm", "repo"])

    def test_a_folder_import_with_nothing_to_import_or_declined_changes_nothing(self) -> None:
        repo = self.standard_repository()
        book = self.root / "EXCEL_UI.xlsm"
        book.write_bytes(b"unchanged")
        folder = self.root / "downloads"
        folder.mkdir()
        self.write(folder, "M_A.bas", "")
        completed = self.run_script("-WorkbookPath", str(book), "-RepoRoot", str(repo), "-SourceDir", str(folder))
        self.assertStops(completed, 2, "Nothing to import")
        self.write(folder, "M_B.bas", 'Attribute VB_Name = "M_B"\n')
        completed = self.run_script("-WorkbookPath", str(book), "-RepoRoot", str(repo), "-SourceDir", str(folder),
                                    answer="n\n")
        self.assertStops(completed, 2, "Cancelled: the workbook was not changed.")
        self.assertEqual(book.read_bytes(), b"unchanged")
        self.assertEqual(sorted(p.name for p in self.root.iterdir()), ["EXCEL_UI.xlsm", "downloads", "repo"])
        self.assertEqual(sorted(p.name for p in folder.iterdir()), ["M_A.bas", "M_B.bas"])


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class ScriptSyntaxTests(unittest.TestCase):
    def test_every_powershell_script_parses(self) -> None:
        scripts = sorted(TOOLS.glob("*.ps1"))
        self.assertTrue(scripts)
        for script in scripts:
            with self.subTest(script=script.name):
                completed = subprocess.run(
                    [str(PWSH), "-NoProfile", "-NonInteractive", "-Command",
                     "$e = $null; $null = [System.Management.Automation.Language.Parser]::ParseFile("
                     "$env:PS_SCRIPT, [ref]$null, [ref]$e); $e | ForEach-Object { $_.ToString() }"],
                    env={**os.environ, "PS_SCRIPT": str(script)},
                    capture_output=True, text=True, timeout=120, check=False)
                self.assertEqual(completed.returncode, 0, completed.stderr)
                self.assertEqual(completed.stdout.strip(), "", completed.stdout)


class SheetSnapshotOrderTests(unittest.TestCase):
    """Excel assigns a new workbook's sheet CodeNames only once its VBA project is read."""

    def test_the_first_snapshot_follows_reading_the_vba_components(self) -> None:
        text = SCRIPT.read_bytes().decode("ascii")
        read_project = text.index("$null = $vbProject.VBComponents.Count")
        read_names = text.index("foreach ($component in $vbProject.VBComponents) { $null = [string]$component.Name }")
        before = text.index("$before = Get-SheetSnapshot -Workbook $workbook")
        first_import = text.index("Import-StandardModule -VBProject $vbProject")
        self.assertLess(read_project, read_names)
        self.assertLess(read_names, before)
        self.assertLess(before, first_import)
        self.assertEqual(text.count("$before = Get-SheetSnapshot"), 1)


class LauncherTests(unittest.TestCase):
    """The launchers run the synchronizer as documented; no pwsh needed."""

    def read(self, name: str) -> str:
        return (TOOLS / name).read_bytes().decode("ascii")

    def test_launchers_run_the_synchronizer_and_pause(self) -> None:
        for name, arguments in (("Sync-EXCEL_UI-VBA.cmd", '-RepoRoot "%REPO_ROOT%" %BUILD%'),
                                ("Sync-EXCEL_UI-VBA-Downloads.cmd", '-SourceDir "%SOURCE%"')):
            with self.subTest(name=name):
                text = self.read(name)
                self.assertIn('set "SCRIPT=%~dp0Sync-EXCEL_UI-VBA.ps1"', text)
                self.assertIn('powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" '
                              '-WorkbookPath "%WORKBOOK%"', text)
                self.assertIn(arguments, text)
                self.assertIn("pause\r\nexit /b %RC%", text)

    def test_the_launcher_builds_a_development_workbook_without_an_argument(self) -> None:
        text = self.read("Sync-EXCEL_UI-VBA.cmd")
        self.assertIn('set "WORKBOOK=%CD%\\test-results\\EXCEL_UI_Development.xlsm"', text)
        self.assertIn('set "BUILD=-Mode Development -NewWorkbook"', text)

    def test_the_downloads_launcher_needs_a_workbook_and_defaults_to_downloads(self) -> None:
        text = self.read("Sync-EXCEL_UI-VBA-Downloads.cmd")
        self.assertIn('set "DEFAULT_SOURCE=%USERPROFILE%\\Downloads"', text)
        self.assertIn("No workbook given.", text)


if __name__ == "__main__":
    unittest.main()
