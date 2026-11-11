"""Offline tests for helpers of tools/Certify-EXCEL_UI.ps1.

The script's functions are loaded from its syntax tree and run in PowerShell 7
(pwsh) against a temporary Git repository. Nothing here starts Excel. The
tests are skipped where pwsh is not installed.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "tools" / "Certify-EXCEL_UI.ps1"
POLICY = ROOT / ".github" / "excel-evidence-policy.json"
PWSH = shutil.which("pwsh")
sys.path.insert(0, str(ROOT / "tools"))

import check_excel_evidence as evidence  # noqa: E402

# Defines the script's functions without running its body.
LOADER = r"""
$ErrorActionPreference = 'Stop'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($env:CERTIFY_SCRIPT, [ref]$null, [ref]$null)
foreach ($s in $ast.EndBlock.Statements) {
    if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
        . ([scriptblock]::Create($s.Extent.Text))
    }
}
$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$Latin1 = [System.Text.Encoding]::GetEncoding(28591)
$script:RepoRoot = $env:REPO
"""


def git(repo: Path, *args: str) -> None:
    subprocess.run(["git", "-C", str(repo), *args], check=True, capture_output=True)


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class InvokeGitTests(unittest.TestCase):
    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.repo = Path(temp.name) / "repo"
        self.repo.mkdir()
        git(self.repo, "init", "-q")
        git(self.repo, "config", "user.name", "Test")
        git(self.repo, "config", "user.email", "test@example.invalid")
        (self.repo / "a.txt").write_text("a\n", encoding="utf-8")
        git(self.repo, "add", "a.txt")
        git(self.repo, "commit", "-q", "-m", "one")
        self.result = Path(temp.name) / "result.json"

    def pwsh(self, body: str) -> Any:
        completed = subprocess.run(
            [str(PWSH), "-NoProfile", "-NonInteractive", "-Command", LOADER + body],
            env={**os.environ, "CERTIFY_SCRIPT": str(SCRIPT), "REPO": str(self.repo),
                 "RESULT": str(self.result)},
            capture_output=True, text=True, timeout=120, check=False)
        self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
        return json.loads(self.result.read_text(encoding="utf-8-sig"))

    def test_clean_tree_status_is_an_empty_byte_array(self) -> None:
        # The preflight decodes the status of a clean checkout, which git
        # prints as no bytes at all; that must stay a byte array, not $null.
        result = self.pwsh(
            "$bytes = Invoke-Git @('status', '--porcelain', '--untracked-files=no')\n"
            "$text = $Utf8NoBom.GetString($bytes)\n"
            "@{ isNull = ($null -eq $bytes); type = $bytes.GetType().FullName; text = $text } "
            "| ConvertTo-Json | Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        self.assertEqual(result, {"isNull": False, "type": "System.Byte[]", "text": ""})

    def test_output_is_returned_as_a_byte_array(self) -> None:
        result = self.pwsh(
            "$bytes = Invoke-Git @('rev-parse', '--verify', 'HEAD^{commit}')\n"
            "@{ type = $bytes.GetType().FullName; sha = $Utf8NoBom.GetString($bytes).Trim() } "
            "| ConvertTo-Json | Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        self.assertEqual(result["type"], "System.Byte[]")
        self.assertRegex(result["sha"], r"^[0-9a-f]{40}$")

    def test_a_failing_command_throws(self) -> None:
        result = self.pwsh(
            "$failed = ''\ntry { $null = Invoke-Git @('show', 'no-such-ref:x') } "
            "catch { $failed = $_.Exception.Message }\n"
            "@{ failed = $failed } | ConvertTo-Json | Set-Content -LiteralPath $env:RESULT -Encoding utf8")
        self.assertIn("git show no-such-ref:x failed", result["failed"])



@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class StopQuitProcessTests(unittest.TestCase):
    """The close step ends only the instance it started, and only after the wait."""

    def run_body(self, body: str) -> Any:
        with tempfile.TemporaryDirectory() as temp:
            result = Path(temp) / "result.json"
            completed = subprocess.run(
                [str(PWSH), "-NoProfile", "-NonInteractive", "-Command", LOADER + body],
                env={**os.environ, "CERTIFY_SCRIPT": str(SCRIPT), "REPO": temp, "RESULT": str(result)},
                capture_output=True, text=True, timeout=120, check=False)
            self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
            return json.loads(result.read_text(encoding="utf-8-sig"))

    # The helper processes are child PowerShell processes, so the tests run on
    # Windows, where the certification runs, as well as on Linux and macOS.
    HELPER = ("$name = (Get-Process -Id $PID).ProcessName\n"
              "$p = Start-Process -FilePath (Get-Process -Id $PID).Path -PassThru "
              "-ArgumentList '-NoProfile', '-NonInteractive', '-Command', '{command}'\n")
    SLEEPER = HELPER.format(command="Start-Sleep -Seconds 60")
    WRITE = " | ConvertTo-Json | Set-Content -LiteralPath $env:RESULT -Encoding utf8"

    def test_a_process_still_running_after_the_wait_is_ended(self) -> None:
        result = self.run_body(self.SLEEPER + "$r = Stop-QuitProcess -ProcessId $p.Id "
                               "-ExpectedStartTime $p.StartTime -WaitMilliseconds 300 "
                               "-ExpectedName $name\n$p.Refresh()\n"
                               "@{ result = $r; exited = $p.HasExited }" + self.WRITE)
        self.assertEqual(result, {"result": "stopped", "exited": True})

    def test_a_process_with_another_name_is_never_ended(self) -> None:
        result = self.run_body(self.SLEEPER + "$r = Stop-QuitProcess -ProcessId $p.Id "
                               "-ExpectedStartTime $p.StartTime -WaitMilliseconds 300\n"
                               "$p.Refresh()\n$alive = -not $p.HasExited\n$p.Kill()\n"
                               "@{ result = $r; alive = $alive }" + self.WRITE)
        self.assertEqual(result, {"result": "exited", "alive": True})

    def test_a_process_whose_id_was_reused_is_never_ended(self) -> None:
        # Same id and name, but not the instance that was started: its start
        # time differs from the one recorded.
        result = self.run_body(self.SLEEPER + "$r = Stop-QuitProcess -ProcessId $p.Id "
                               "-ExpectedStartTime $p.StartTime.AddSeconds(-5) -WaitMilliseconds 300 "
                               "-ExpectedName $name\n$p.Refresh()\n$alive = -not $p.HasExited\n$p.Kill()\n"
                               "@{ result = $r; alive = $alive }" + self.WRITE)
        self.assertEqual(result, {"result": "exited", "alive": True})

    def test_a_process_that_exits_or_is_gone_is_left_alone(self) -> None:
        result = self.run_body(
            self.HELPER.format(command="exit")
            + "$r = Stop-QuitProcess -ProcessId $p.Id -ExpectedStartTime (Get-Date) -WaitMilliseconds 5000 "
            "-ExpectedName $name\n"
            "$gone = Stop-QuitProcess -ProcessId 999999 -ExpectedStartTime (Get-Date) -WaitMilliseconds 100\n"
            "@{ result = $r; gone = $gone }" + self.WRITE)
        self.assertEqual(result, {"result": "exited", "gone": "exited"})


def policy_document() -> dict[str, Any]:
    value = json.loads(POLICY.read_text(encoding="utf-8"))
    assert isinstance(value, dict)
    return value


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class EvidencePolicyTests(unittest.TestCase):
    """The script takes its file names, steps, runner and configuration from the policy."""

    WRITE = " | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $env:RESULT -Encoding utf8"
    READ_ERROR = ("$m = ''\ntry { $null = Read-EvidencePolicy $env:POLICY } catch { $m = $_.Exception.Message }\n"
                  "@{ error = $m }" + WRITE)

    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.temp = Path(temp.name)
        self.policy = self.temp / "policy.json"
        self.result = self.temp / "result.json"

    def write_policy(self, change: Callable[[dict[str, Any]], None] | None = None) -> None:
        document = policy_document()
        if change is not None:
            change(document)
        self.policy.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")

    def run_body(self, body: str, repo: Path | None = None) -> Any:
        completed = subprocess.run(
            [str(PWSH), "-NoProfile", "-NonInteractive", "-Command", LOADER + body],
            env={**os.environ, "CERTIFY_SCRIPT": str(SCRIPT), "REPO": str(repo or self.temp),
                 "RESULT": str(self.result), "POLICY": str(self.policy)},
            capture_output=True, text=True, timeout=120, check=False)
        self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
        return json.loads(self.result.read_text(encoding="utf-8-sig"))

    def test_the_repository_policy_gives_the_names_steps_and_runner(self) -> None:
        self.policy = POLICY
        result = self.run_body(
            "$p = Read-EvidencePolicy $env:POLICY\n"
            "@{ steps = @($p.steps); runner = $p.certification_runner; files = $p.required_files; "
            "evidence = @($p.evidence_files); schema = $p.evidence_schema; configuration = $p.configuration }"
            + self.WRITE)
        document = policy_document()
        self.assertEqual(result, {
            "steps": document["steps"], "runner": document["certification_runner"],
            "files": document["required_files"], "evidence": document["evidence_files"],
            "schema": document["evidence_schema"], "configuration": document["configuration"]})

    def test_an_unusable_policy_stops_the_script(self) -> None:
        def drop(key: str) -> Callable[[dict[str, Any]], None]:
            return lambda document: document.pop(key)

        def set_value(key: str, value: Any) -> Callable[[dict[str, Any]], None]:
            return lambda document: document.update({key: value})

        cases: dict[str, tuple[Callable[[dict[str, Any]], None], str]] = {
            "no steps": (drop("steps"), "has no steps"),
            "no runner": (drop("certification_runner"), "has no certification_runner"),
            "empty runner": (set_value("certification_runner", " "), "certification_runner is not text"),
            "text schema": (set_value("evidence_schema", "1"), "evidence_schema is not a number"),
            "empty steps": (set_value("steps", []), "lists no steps"),
            "repeated step": (set_value("steps", ["build", "build"]), "distinct names"),
            "missing role": (lambda document: document["required_files"].pop("host"), "names no host file"),
            "other evidence": (set_value("evidence_files", ["certification.json", "host.json"]),
                               "not the certification.json and certification.txt this script collects"),
        }
        for name, (change, fragment) in cases.items():
            with self.subTest(name):
                self.write_policy(change)
                self.assertIn(fragment, self.run_body(self.READ_ERROR)["error"])
        with self.subTest("not JSON"):
            self.policy.write_text("{", encoding="utf-8")
            self.assertIn("is not valid JSON", self.run_body(self.READ_ERROR)["error"])
        with self.subTest("missing"):
            self.policy.unlink()
            self.assertIn("The evidence policy is missing", self.run_body(self.READ_ERROR)["error"])

    def test_set_step_records_a_policy_step_and_refuses_any_other(self) -> None:
        self.write_policy(lambda document: document.update(steps=["first", "second"]))
        result = self.run_body(
            "$p = Read-EvidencePolicy $env:POLICY\n"
            "$script:Session = [ordered]@{ steps = @(@($p.steps) | ForEach-Object {\n"
            "    [ordered]@{ name = $_; status = 'NOT_RUN'; detail = ''; at_utc = $null } }) }\n"
            "Set-Step 'second' 'PASS' 'done' | Out-Null\n"
            "$m = ''\ntry { Set-Step 'preflight' 'PASS' '' | Out-Null } catch { $m = $_.Exception.Message }\n"
            "@{ second = (Get-StepStatus 'second'); first = (Get-StepStatus 'first'); "
            "current = $script:CurrentStep; error = $m }" + self.WRITE)
        self.assertEqual(result, {"second": "PASS", "first": "NOT_RUN", "current": "second",
                                  "error": "The step 'preflight' is not in the evidence policy."})

    def test_source_digests_follow_the_policy_and_match_the_validator(self) -> None:
        repo = self.temp / "repo"
        repo.mkdir()
        manifest = {"schema_version": 1, "configurations": {
            "certified": {"modules": ["lib/M_ONE.bas"]},
            "development": {"modules": ["lib/M_ONE.bas", "lib/M_TWO.bas"]}}}
        files = {"config/projects.json": json.dumps(manifest) + "\n",
                 "lib/M_ONE.bas": 'Attribute VB_Name = "M_ONE"\nOption Explicit\n',
                 "lib/M_TWO.bas": 'Attribute VB_Name = "M_TWO"\r\nOption Explicit\r\n'}
        git(repo, "init", "-q")
        git(repo, "config", "user.name", "Test")
        git(repo, "config", "user.email", "test@example.invalid")
        for path, text in files.items():
            (repo / path).parent.mkdir(parents=True, exist_ok=True)
            (repo / path).write_bytes(text.encode("ascii"))
        git(repo, "add", "-A")
        git(repo, "commit", "-q", "-m", "fixture")
        self.write_policy(lambda document: document.update(
            manifest="config/projects.json", configuration="certified", evidence_schema=3))
        result = self.run_body(
            "$script:Policy = Read-EvidencePolicy $env:POLICY\n"
            "$sha = $Utf8NoBom.GetString((Invoke-Git @('rev-parse', 'HEAD'))).Trim()\n"
            "Get-SourceDigests $sha" + self.WRITE, repo=repo)
        for module in result["modules"]:
            module.pop("checked_out_sha256")
        sha = evidence.resolve_commit(repo, "HEAD")
        expected = evidence.source_inventory(repo, sha, evidence.load_policy(self.policy))
        self.assertEqual(expected["schema"], 3)
        self.assertEqual(expected["configuration"], "certified")
        self.assertEqual(result, expected)


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
@unittest.skipIf(os.name == "nt", "on Windows the preflight could reach a real Excel installation")
class ScriptRunTests(unittest.TestCase):
    """The whole script, run where its preflight fails before anything starts Excel."""

    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.repo = Path(temp.name) / "repo"
        self.bundle = Path(temp.name) / "bundle"
        self.repo.mkdir()
        git(self.repo, "init", "-q")
        git(self.repo, "config", "user.name", "Test")
        git(self.repo, "config", "user.email", "test@example.invalid")
        (self.repo / "a.txt").write_text("a\n", encoding="utf-8")

    def commit(self) -> None:
        git(self.repo, "add", "-A")
        git(self.repo, "commit", "-q", "-m", "fixture")

    def run_script(self) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [str(PWSH), "-NoProfile", "-NonInteractive", "-File", str(SCRIPT), "-RepoRoot", str(self.repo),
             "-BundleDir", str(self.bundle), "-Python", "no-such-python-for-this-test"],
            capture_output=True, text=True, timeout=120, check=False)

    def test_the_session_record_uses_the_policy_names_and_steps(self) -> None:
        document = policy_document()
        document["required_files"]["session"] = "run-record.json"
        document["evidence_schema"] = 7
        (self.repo / ".github").mkdir()
        (self.repo / ".github" / "excel-evidence-policy.json").write_text(json.dumps(document), encoding="utf-8")
        self.commit()
        completed = self.run_script()
        self.assertEqual(completed.returncode, 1, completed.stdout + completed.stderr)
        self.assertFalse((self.bundle / "session.json").exists())
        session = json.loads((self.bundle / "run-record.json").read_text(encoding="utf-8-sig"))
        self.assertEqual(session["schema"], 7)
        self.assertEqual([step["name"] for step in session["steps"]], document["steps"])
        self.assertEqual(session["steps"][0]["status"], "FAIL")
        self.assertEqual({step["status"] for step in session["steps"][1:]}, {"NOT_RUN"})

    def test_a_missing_policy_stops_the_script_before_the_bundle(self) -> None:
        self.commit()
        completed = self.run_script()
        self.assertEqual(completed.returncode, 1)
        self.assertIn("ERROR: The evidence policy is missing", completed.stdout)
        self.assertFalse(self.bundle.exists())


if __name__ == "__main__":
    unittest.main()
