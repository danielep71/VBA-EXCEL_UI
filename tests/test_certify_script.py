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
import tempfile
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "tools" / "Certify-EXCEL_UI.ps1"
PWSH = shutil.which("pwsh")

# Defines the script's functions without running its body.
LOADER = r"""
$ErrorActionPreference = 'Stop'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($env:CERTIFY_SCRIPT, [ref]$null, [ref]$null)
foreach ($s in $ast.EndBlock.Statements) {
    if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
        . ([scriptblock]::Create($s.Extent.Text))
    }
}
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
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


if __name__ == "__main__":
    unittest.main()
