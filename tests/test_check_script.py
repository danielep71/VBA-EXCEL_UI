"""Offline tests for tools/Check-EXCEL_UI.ps1 and the VBA entry points it calls.

The contract tests read the script and the VBA modules as text. The helper
tests load the script's functions from its syntax tree and run them in
PowerShell 7 (pwsh), and the run tests start the whole script where its
preflight fails before anything starts Excel. Nothing here starts Excel; the
pwsh tests are skipped where pwsh is not installed.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from typing import Any

from gate_fixture import house_style

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "tools" / "Check-EXCEL_UI.ps1"
LAUNCHER = ROOT / "tools" / "Check-EXCEL_UI.cmd"
CERTIFY = ROOT / "tools" / "Certify-EXCEL_UI.ps1"
CERTIFY_LAUNCHER = ROOT / "tools" / "Certify-EXCEL_UI.cmd"
POLICY = ROOT / ".github" / "excel-evidence-policy.json"
TEST_CERTIFICATION = ROOT / "test" / "M_EXCEL_UI_TEST_CERTIFICATION.bas"
TEST_HARNESS = ROOT / "test" / "M_EXCEL_UI_TEST_HARNESS.bas"
DEMO_TESTS = ROOT / "demo" / "M_DEMO_CLEANUP_TESTS.bas"
DEMO = ROOT / "demo" / "M_EXCEL_UI_DEMO.bas"
PWSH = shutil.which("pwsh")

# Defines the script's functions, and the ones it takes from the
# certification script, without running either body.
LOADER = r"""
$ErrorActionPreference = 'Stop'
$ast = [System.Management.Automation.Language.Parser]::ParseFile($env:CHECK_SCRIPT, [ref]$null, [ref]$null)
foreach ($s in $ast.EndBlock.Statements) {
    if ($s -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
        . ([scriptblock]::Create($s.Extent.Text))
    }
}
$Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$shared = Import-CertifyDefinitions -Path $env:CERTIFY_SCRIPT `
    -Functions @('Get-UtcStamp', 'Write-Log', 'Get-InnerMessage')
. $shared.Definitions
$script:SessionLog = $null
"""
WRITE = " | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $env:RESULT -Encoding utf8"


def read(path: Path) -> str:
    return path.read_text(encoding="ascii").replace("\r\n", "\n")


def script_value(name: str) -> str:
    """The single-quoted text a script variable is set to."""
    match = re.search(rf"^\${name} = '([^']*)'$", read(SCRIPT), re.M)
    assert match is not None, name
    return match.group(1)


def script_list(name: str) -> list[str]:
    """The single-quoted texts of a script variable set to @(...)."""
    match = re.search(rf"^\${name} = @\((.*?)\)$", read(SCRIPT), re.M | re.S)
    assert match is not None, name
    return re.findall(r"'([^']*)'", match.group(1))


def function_body(text: str, name: str) -> str:
    """The text of a PowerShell function, from its header to its closing brace."""
    match = re.search(rf"^function {re.escape(name)}\b.*?^}}$", text, re.M | re.S)
    assert match is not None, name
    return match.group(0)


def vba_procedure(path: Path, name: str) -> str:
    match = re.search(rf"^Public (?:Sub|Function) {name}\b.*?^End (?:Sub|Function)$", read(path), re.M | re.S)
    assert match is not None, name
    return match.group(0)


def case_labels(body: str) -> list[str]:
    return re.findall(r'^\s*Case "([^"]+)"$', body, re.M)


def vba_publics(*paths: Path) -> set[str]:
    names: set[str] = set()
    for path in paths:
        names.update(re.findall(r"^Public (?:Sub|Function) (\w+)", read(path), re.M))
    return names


def git(repo: Path, *args: str) -> None:
    subprocess.run(["git", "-C", str(repo), *args], check=True, capture_output=True)


class ContractTests(unittest.TestCase):
    """The script, the VBA entry points and the certification script agree."""

    def test_the_runner_list_is_what_the_vba_entry_point_runs(self) -> None:
        default = script_list("DefaultRunners")
        probe = script_value("ProbeRunner")
        body = vba_procedure(TEST_CERTIFICATION, script_value("RunnerMacro"))
        self.assertEqual(case_labels(body), [*default, probe])
        self.assertEqual(default, [
            "Test_EXCEL_UI_RunAll", "Test_EXCEL_UI_RunCore", "Test_EXCEL_UI_RunTitleBarOnly",
            "Test_EXCEL_UI_RunSnapshotIdentity", "Test_EXCEL_UI_RunTitleBarSdiIdentity",
            "Test_EXCEL_UI_RunCertificationSelfTest", "Test_EXCEL_UI_RunOwnershipCleanupChecks"])
        publics = vba_publics(*sorted((ROOT / "test").glob("*.bas")))
        for name in [*default, probe]:
            with self.subTest(name):
                self.assertIn(name, publics)
                self.assertRegex(body, rf"(?m)^\s+{name}$")
        policy = json.loads(POLICY.read_text(encoding="utf-8"))
        self.assertNotIn(policy["certification_runner"], case_labels(body))

    def test_each_runner_starts_after_the_active_frame_settles(self) -> None:
        text = read(SCRIPT)
        loop = text[text.index("foreach ($step in $runnerSteps)"):]
        loop = loop[:loop.index("Complete-CloseStep 'close-development'")]
        self.assertLess(loop.index("Wait-ActiveFrame $devState"), loop.index("Invoke-LoggedMacro"))
        self.assertIn("the active frame did not settle before the run", loop)
        settle = text[text.index("function Wait-ActiveFrame"):]
        settle = settle[:settle.index("\n}\n")]
        for read_back in ("$State.Excel.ActiveWindow", "$frameWindow.Hwnd", "$State.Excel.Hwnd"):
            with self.subTest(read_back):
                self.assertIn(read_back, settle)
        self.assertRegex(text, r"(?m)^\$ActiveFrameSettleSeconds = \d+$")

    def test_the_demo_macros_are_what_the_demo_entry_point_runs(self) -> None:
        body = vba_procedure(DEMO_TESTS, script_value("DemoLoggedMacro"))
        macros = [script_value(name) for name in
                  ("DemoCleanupMacro", "DemoBuildMacro", "DemoClearMacro", "DemoHideMacro", "DemoShowMacro")]
        self.assertEqual(case_labels(body), macros)
        publics = vba_publics(*sorted((ROOT / "demo").glob("*.bas")))
        for name in [*macros, script_value("DemoLoggedMacro"), script_value("DemoTitleBarStateMacro")]:
            with self.subTest(name):
                self.assertIn(name, publics)
        self.assertIn("UI_TryGetTitleBarVisible(IsVisible, FailMsg)",
                      vba_procedure(DEMO_TESTS, script_value("DemoTitleBarStateMacro")))

    def test_the_demo_sheet_and_check_box_names_are_the_demo_constants(self) -> None:
        demo = read(DEMO)
        for constant, variable in (("DEMO_SHEET_NAME", "DemoSheetName"), ("CB_TITLEBAR", "TitleBarCheckBox")):
            with self.subTest(constant):
                self.assertRegex(demo, rf'Private Const {constant} +As String = "{script_value(variable)}"')
        # The CLEAR ALL, HIDE SELECTED UI and SHOW SELECTED UI buttons run the
        # macros the script runs, and the HIDE and SHOW actions read chkTitleBar.
        flat = re.sub(r" _\n\s*", " ", demo)
        for caption, variable in (("CLEAR ALL", "DemoClearMacro"), ("HIDE SELECTED UI", "DemoHideMacro"),
                                  ("SHOW SELECTED UI", "DemoShowMacro")):
            with self.subTest(caption):
                self.assertIn(f'"{caption}", "{script_value(variable)}")', flat)
        self.assertIn("TitleBarVis = Demo_CheckBoxToUIVisibility(WS, CB_TITLEBAR,", demo)

    def test_every_shared_name_is_defined_by_the_certification_script(self) -> None:
        certify = read(CERTIFY)
        for name in script_list("SharedFunctions"):
            with self.subTest(name):
                self.assertRegex(certify, rf"(?m)^function {re.escape(name)}\b")
        for name in script_list("SharedValues"):
            with self.subTest(name):
                self.assertRegex(certify, rf"(?m)^\${name} = ")
        self.assertEqual(script_value("CertifyScriptName"), CERTIFY.name)

    def test_macro_security_is_lowered_only_behind_unattended_and_nothing_is_written(self) -> None:
        text = read(SCRIPT)
        certify = read(CERTIFY)
        # This script never sets AutomationSecurity itself: it calls the
        # certification script's function, so both start an instance alike.
        self.assertEqual(re.findall(r"AutomationSecurity\s*=", text), [])
        self.assertEqual(len(re.findall(r"\bSet-ExcelAutomationSecurity \$State\.Excel", text)), 1)
        self.assertEqual(re.findall(r"\$MsoAutomationSecurityLow\b", text),
                         ["$MsoAutomationSecurityLow"])
        # There, the only assignments are in one function, and the lowering is
        # the branch that runs only when -Unattended is given.
        self.assertEqual(re.findall(r"\.AutomationSecurity = (\S+)", certify),
                         ["$MsoAutomationSecurityLow", "$MsoAutomationSecurityByUI"])
        body = function_body(certify, "Set-ExcelAutomationSecurity")
        self.assertRegex(body, r"if \(\$Unattended\) \{\n\s+Invoke-Excel \{ \$Excel\.AutomationSecurity = "
                               r"\$MsoAutomationSecurityLow \}\n\s+\}\n\s+else \{\n\s+Invoke-Excel \{ "
                               r"\$Excel\.AutomationSecurity = \$MsoAutomationSecurityByUI \}")
        self.assertEqual(re.findall(r"\$MsoAutomationSecurityLow\b", certify.replace(body, "")),
                         ["$MsoAutomationSecurityLow"])
        self.assertRegex(certify, r"(?m)^\$MsoAutomationSecurityLow = 1$")
        # Nothing writes the Trust Center, the registry or a Trusted Location.
        for name, source in (("check", text), ("certify", certify), ("check launcher", read(LAUNCHER)),
                             ("certify launcher", read(CERTIFY_LAUNCHER))):
            for pattern in (r"Set-ItemProperty", r"New-ItemProperty", r"Remove-ItemProperty",
                            r"\breg(?:\.exe)? (?:add|delete)", r"VBAWarnings\s*=", r"AccessVBOM\s*=",
                            r"TrustedLocations?\s*=", r"\.SetValue\("):
                with self.subTest(name=name, pattern=pattern):
                    self.assertIsNone(re.search(pattern, source, re.I))

    def test_no_two_script_variables_differ_only_in_case(self) -> None:
        text = re.sub(r"^function [^\n]*\{.*?^\}[ \t]*$", "", read(SCRIPT), flags=re.MULTILINE | re.DOTALL)
        spellings: dict[str, set[str]] = {}
        for name in re.findall(r"\$(?:script:)?([A-Za-z_][A-Za-z0-9_]*)", text):
            spellings.setdefault(name.lower(), set()).add(name)
        clashes = {key: sorted(names) for key, names in spellings.items() if len(names) > 1}
        self.assertEqual(clashes, {})

    def test_the_launcher_passes_its_arguments_to_the_script(self) -> None:
        self.assertIn('set "SCRIPT=%~dp0Check-EXCEL_UI.ps1"', read(LAUNCHER))
        self.assertIn('-File "%SCRIPT%" %UNATTENDED_ARG% %*', read(LAUNCHER))
        self.assertEqual(SCRIPT.read_bytes().count(b"\r\n"), SCRIPT.read_bytes().count(b"\n"))
        self.assertEqual(LAUNCHER.read_bytes().count(b"\r\n"), LAUNCHER.read_bytes().count(b"\n"))

    def test_the_check_launcher_is_unattended_and_never_pauses(self) -> None:
        text = read(LAUNCHER)
        # It passes -Unattended itself, once, even when the caller also gives it.
        self.assertIn('set "UNATTENDED_ARG=-Unattended"', text)
        self.assertIn('if /i "%~1"=="-Unattended" set "UNATTENDED_ARG="', text)
        self.assertIn("shift /1", text)
        self.assertNotRegex(text, r"(?m)^\s*(?:if [^\n]* )?pause\b")
        self.assertEqual(text.count("powershell.exe "), 1)

    def test_the_certify_launcher_pauses_unless_unattended_is_an_argument(self) -> None:
        text = read(CERTIFY_LAUNCHER)
        self.assertIn('-File "%SCRIPT%" %*', text)
        # It never adds the switch: an attended run stays attended.
        self.assertNotIn("%UNATTENDED_ARG%", text)
        self.assertNotRegex(text, r'-File "%SCRIPT%"[^\n]*-Unattended')
        self.assertIn('set "PAUSE_AT_END=1"', text)
        self.assertIn('if /i "%~1"=="-Unattended" set "PAUSE_AT_END="', text)
        self.assertIn("shift /1", text)
        self.assertEqual(len(re.findall(r"(?m)^\s*pause\b", text)), 0)
        self.assertEqual(len(re.findall(r"(?m)^\s*if defined PAUSE_AT_END pause$", text)), 2)
        self.assertEqual(CERTIFY_LAUNCHER.read_bytes().count(b"\r\n"), CERTIFY_LAUNCHER.read_bytes().count(b"\n"))

    def test_both_scripts_take_an_unattended_switch_and_check_forwards_it(self) -> None:
        for path in (SCRIPT, CERTIFY):
            with self.subTest(path.name):
                param = re.search(r"(?s)\[CmdletBinding\(\)\]\nparam\((.*?)\n\)", read(path))
                assert param is not None
                self.assertIn("[switch]$Unattended", param.group(1))
        text = read(SCRIPT)
        self.assertRegex(text, r"if \(\$Unattended\) \{\n\s+\$certifyArguments \+= '-Unattended'\n\s+\}")
        self.assertIn("(Get-AutomationSecurityLabel)", text)
        self.assertNotIn("'msoAutomationSecurityByUI'", text)
        self.assertIn("msoAutomationSecurityLow (-Unattended)", read(CERTIFY))

    def test_no_prompt_is_reachable_when_unattended(self) -> None:
        check = read(SCRIPT)
        certify = read(CERTIFY)
        self.assertNotIn("Read-Host", check)
        # One place reads the console, and it refuses to when unattended.
        self.assertEqual(len(re.findall(r"\bRead-Host\b", certify)), 1)
        reader = function_body(certify, "Read-OperatorAnswer")
        self.assertRegex(reader, r"(?s)if \(\$Unattended\) \{\n\s+throw .*?\n\s+\}\n\s+return \(Read-Host \$Prompt\)")
        # Every caller leaves an unattended run before it reaches the prompt:
        # by throwing or, while it polls, by continuing the loop.
        callers = [name for name in re.findall(r"(?m)^function ([\w-]+)", certify)
                   if name != "Read-OperatorAnswer" and "Read-OperatorAnswer " in function_body(certify, name)]
        self.assertEqual(callers, ["Wait-ExcelMacros", "Resolve-CompileFallback"])
        for name in callers:
            with self.subTest(name):
                body = function_body(certify, name)
                guard = body.index("if ($Unattended) {")
                prompt = body.index("Read-OperatorAnswer ")
                self.assertLess(guard, prompt)
                self.assertRegex(body[guard:prompt], r"\n\s+(?:throw|continue)\b")
        # The script bodies outside the functions have no prompt of their own.
        for name, source in (("check", check), ("certify", certify)):
            with self.subTest(name):
                outside = re.sub(r"(?m)^function [\w-]+.*?^}$", "", source, flags=re.S)
                for pattern in (r"Read-Host", r"Read-OperatorAnswer ", r"\bPause\b", r"\[Console\]::Read", r"ReadKey",
                                r"Out-GridView", r"MessageBox"):
                    self.assertIsNone(re.search(pattern, outside), pattern)
        # Each wait an unattended run makes has a bound, which the help states.
        self.assertRegex(certify, r"(?m)^\$MacrosWaitSeconds = 60$")
        self.assertRegex(certify, r"(?m)^\$CompileWaitSeconds = 30$")
        self.assertIn("at most 60 seconds", certify)
        self.assertIn("at most 30 seconds", certify)


class LogSinkSourceTests(unittest.TestCase):
    """With no sink set, the harnesses print what they printed before."""

    def test_tst_log_prints_the_line_it_mirrors(self) -> None:
        body = vba_procedure(TEST_HARNESS, "TST_Log")
        self.assertIn('LineText = ProcName & " @ " & Stage & " | " & Detail\n'
                      "        Debug.Print LineText\n", body)
        self.assertIn("        TST_LogSinkWriteLine LineText\n", body)
        self.assertEqual(len(re.findall(r"(?m)^\s*Debug\.Print ", body)), 1)

    def test_the_sink_writes_nothing_without_a_path_and_never_raises(self) -> None:
        for path, name, state in ((TEST_HARNESS, "TST_LogSinkWriteLine", "m_LogSinkPath"),
                                  (DEMO_TESTS, "DEMO_CT_LogSinkWrite", "mLogSinkPath")):
            with self.subTest(name):
                text = read(path)
                match = re.search(rf"^(?:Public|Private) Sub {name}\b.*?^End Sub$", text, re.M | re.S)
                assert match is not None
                body = match.group(0)
                self.assertRegex(body, rf"If Len\({state}\) = 0 Or Len\({state[:-4]}Failure\) > 0 Then\n"
                                       r"\s+Exit Sub")
                self.assertNotIn("Err.Raise", body)
                self.assertIn(f"Resume {house_style.labels().safe_exit}", body)

    def test_every_demo_harness_line_goes_through_the_sink(self) -> None:
        text = read(DEMO_TESTS)
        self.assertEqual(re.findall(r"(?m)^[ \t]*Debug\.Print (.*)$", text), ["LineText"])
        self.assertEqual(len(re.findall(r"(?m)^[ \t]*DEMO_CT_Print ", text)), 9)

    def test_the_entry_points_report_instead_of_raising(self) -> None:
        for path, name in ((TEST_CERTIFICATION, "TST_RunRunnerWithLog"), (DEMO_TESTS, "Demo_Test_RunLogged")):
            with self.subTest(name):
                body = vba_procedure(path, name)
                handler = body.split(f"{house_style.labels().handler}:", 1)[1]
                self.assertNotIn("Err.Raise", handler)
                self.assertIn(f"Resume {house_style.labels().safe_exit}", handler)
                self.assertIn('Verdict = "RESULT: PASS"', body)
                self.assertIn('Verdict = "RESULT: FAIL | Error " & CStr(RaisedNumber)', body)
                self.assertIn('" | LOG WRITE FAILED: "', body)

    def test_a_run_that_only_skipped_is_an_incomplete_verdict_that_names_the_skips(self) -> None:
        body = vba_procedure(DEMO_TESTS, "Demo_Test_RunLogged")
        handler = body.split(f"{house_style.labels().handler}:", 1)[1]
        self.assertIn("If RaisedNumber = ERR_RUN_INCOMPLETE Then", handler)
        self.assertIn('Verdict = "RESULT: INCOMPLETE | " & RaisedDescription', handler)
        runner = vba_procedure(DEMO_TESTS, "Demo_Test_RunCleanupChecks")
        self.assertIn("If mSkipped <> 0 Then\n            Err.Raise ERR_RUN_INCOMPLETE, PROC_NAME", runner)
        self.assertIn('", skipped " & CStr(mSkipped)', runner)


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
class HelperTests(unittest.TestCase):
    """The script's helpers, run in pwsh."""

    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.temp = Path(temp.name)
        self.result = self.temp / "result.json"

    def run_body(self, body: str) -> Any:
        completed = subprocess.run(
            [str(PWSH), "-NoProfile", "-NonInteractive", "-Command", LOADER + body],
            env={**os.environ, "CHECK_SCRIPT": str(SCRIPT), "CERTIFY_SCRIPT": str(CERTIFY),
                 "RESULT": str(self.result), "TEMPDIR": str(self.temp)},
            capture_output=True, text=True, timeout=120, check=False)
        self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
        return json.loads(self.result.read_text(encoding="utf-8-sig"))

    def test_shared_definitions_come_from_the_certification_script(self) -> None:
        result = self.run_body(
            "$d = Import-CertifyDefinitions -Path $env:CERTIFY_SCRIPT -Functions @('Stop-QuitProcess') "
            "-Variables @('MsoAutomationSecurityByUI', 'CompileControlId', 'ReadinessMacro')\n"
            "$m = ''\ntry { $null = Import-CertifyDefinitions -Path $env:CERTIFY_SCRIPT "
            "-Functions @('Get-UtcStamp', 'No-SuchFunction') -Variables @('NoSuchValue') } "
            "catch { $m = $_.Exception.Message }\n"
            "$gone = ''\ntry { $null = Import-CertifyDefinitions -Path (Join-Path $env:TEMPDIR 'none.ps1') "
            "-Functions @('Get-UtcStamp') } catch { $gone = $_.Exception.Message }\n"
            "@{ values = $d.Values; text = $d.Definitions.ToString().StartsWith('function Stop-QuitProcess'); "
            "error = $m; gone = $gone }" + WRITE)
        self.assertEqual(result["values"], {"MsoAutomationSecurityByUI": 2, "CompileControlId": 578,
                                            "ReadinessMacro": "M_EXCEL_UI.UI_HasExcelUIStateSnapshot"})
        self.assertTrue(result["text"])
        self.assertIn("no longer defines No-SuchFunction, NoSuchValue.", result["error"])
        self.assertIn("The certification script is missing", result["gone"])

    # Loads the shared Excel helpers and their constants from the certification
    # script, as the script does, with stand-ins for the Excel application
    # (its Run call fails for the first calls) and for the VBE.
    HELPERS_LOAD = (
        "$names = @('Get-ComHResult', 'Invoke-Excel', 'Release-Com', 'Get-AutomationSecurityLabel', "
        "'Set-ExcelAutomationSecurity', 'Read-OperatorAnswer', 'Wait-ExcelMacros', 'Wait-CompileCommand', "
        "'Resolve-CompileFallback')\n"
        "$d = Import-CertifyDefinitions -Path $env:CERTIFY_SCRIPT -Functions $names `\n"
        "    -Variables @('MsoAutomationSecurityByUI', 'MsoAutomationSecurityLow', 'CompileControlId', "
        "'ReadinessMacro', 'MacrosWaitSeconds', 'MacrosPollSeconds', 'CompileWaitSeconds')\n"
        ". $d.Definitions\n"
        "foreach ($k in $d.Values.Keys) { Set-Variable -Name $k -Value $d.Values[$k] }\n"
        "function New-FakeExcel([int]$FailFirst) {\n"
        "    $e = [pscustomobject]@{ Calls = 0; AutomationSecurity = 0; FailFirst = $FailFirst }\n"
        "    $e | Add-Member -MemberType ScriptMethod -Name Run -Value {\n"
        "        param($Macro)\n"
        "        $this.Calls++\n"
        "        if ($this.Calls -le $this.FailFirst) { throw 'Cannot run the macro' }\n"
        "        return $null\n"
        "    }\n"
        "    return $e\n"
        "}\n"
        "function New-FakeVbe([bool]$Enabled) {\n"
        "    $control = [pscustomobject]@{ Enabled = $Enabled }\n"
        "    $bars = [pscustomobject]@{ Control = $control; Looks = 0 }\n"
        "    $bars | Add-Member -MemberType ScriptMethod -Name FindControl -Value {\n"
        "        param($Type, $Id)\n"
        "        $this.Looks++\n"
        "        return $this.Control\n"
        "    }\n"
        "    return [pscustomobject]@{ CommandBars = $bars }\n"
        "}\n"
    )

    def test_unattended_starts_the_instance_with_low_automation_security_and_says_so(self) -> None:
        result = self.run_body(
            self.HELPERS_LOAD +
            "$first = New-FakeExcel 0\n$second = New-FakeExcel 0\n"
            "$Unattended = $false\n$a = Set-ExcelAutomationSecurity $first\n"
            "$Unattended = $true\n$u = Set-ExcelAutomationSecurity $second\n"
            "@{ attended = $first.AutomationSecurity; attendedLabel = $a; "
            "unattended = $second.AutomationSecurity; unattendedLabel = $u; "
            "values = @($MsoAutomationSecurityByUI, $MsoAutomationSecurityLow) }" + WRITE)
        self.assertEqual(result["values"], [2, 1])
        self.assertEqual((result["attended"], result["attendedLabel"]), (2, "msoAutomationSecurityByUI"))
        self.assertEqual((result["unattended"], result["unattendedLabel"]),
                         (1, "msoAutomationSecurityLow (-Unattended)"))

    def test_an_unattended_macros_wait_polls_then_fails_with_the_reason_and_never_prompts(self) -> None:
        # pwsh runs with -NonInteractive here, so a Read-Host would raise
        # instead; the messages below show which path ran.
        result = self.run_body(
            self.HELPERS_LOAD +
            "$Unattended = $true\n"
            "$late = New-FakeExcel 3\n"
            "Wait-ExcelMacros -Excel $late -Qualified \"'W.xlsm'!\" -WorkbookName 'W.xlsm' -TimeoutSeconds 30 "
            "-PollSeconds 0.05\n"
            "$never = New-FakeExcel 1000000\n$clock = [Diagnostics.Stopwatch]::StartNew()\n$m = ''\n"
            "try { Wait-ExcelMacros -Excel $never -Qualified \"'W.xlsm'!\" -WorkbookName 'W.xlsm' "
            "-TimeoutSeconds 1 -PollSeconds 0.1 } catch { $m = $_.Exception.Message }\n"
            "@{ lateCalls = $late.Calls; neverCalls = $never.Calls; seconds = $clock.Elapsed.TotalSeconds; "
            "error = $m }" + WRITE)
        self.assertEqual(result["lateCalls"], 4)
        self.assertGreater(result["neverCalls"], 2)
        self.assertLess(result["seconds"], 30)
        self.assertIn("The macros of W.xlsm did not run within 1 seconds although -Unattended set "
                      "AutomationSecurity to msoAutomationSecurityLow; the last error was: Cannot run the macro.",
                      result["error"])
        self.assertIn("run without -Unattended to look", result["error"])

    def test_an_unattended_compile_that_cannot_be_confirmed_fails_instead_of_asking(self) -> None:
        result = self.run_body(
            self.HELPERS_LOAD +
            "$Unattended = $true\n$CompileWaitSeconds = 1\n"
            "$stuck = New-FakeVbe $true\n$done = New-FakeVbe $false\n"
            "$clock = [Diagnostics.Stopwatch]::StartNew()\n"
            "$stuckResult = Wait-CompileCommand $stuck\n$seconds = $clock.Elapsed.TotalSeconds\n"
            "$doneResult = Wait-CompileCommand $done\n"
            "$m = ''\ntry { $null = Resolve-CompileFallback } catch { $m = $_.Exception.Message }\n"
            "$p = ''\ntry { $null = Read-OperatorAnswer 'Question' } catch { $p = $_.Exception.Message }\n"
            "@{ stuck = $stuckResult; done = $doneResult; looks = $stuck.CommandBars.Looks; seconds = $seconds; "
            "error = $m; prompt = $p }" + WRITE)
        self.assertFalse(result["stuck"])
        self.assertTrue(result["done"])
        self.assertEqual(result["looks"], 2)
        self.assertLess(result["seconds"], 30)
        self.assertEqual(result["error"],
                         "Debug > Compile VBAProject was still enabled after 1 seconds, so the compile could not "
                         "be confirmed and -Unattended never asks. Run without -Unattended and compile in the VBE "
                         "to see why.")
        self.assertIn("An unattended run reached a prompt (Question)", result["prompt"])

    def test_an_attended_run_still_looks_for_five_seconds_and_asks(self) -> None:
        result = self.run_body(
            self.HELPERS_LOAD +
            "$Unattended = $false\n$vbe = New-FakeVbe $false\n$ok = Wait-CompileCommand $vbe\n"
            "$m = ''\ntry { $null = Read-OperatorAnswer 'Question' } catch { $m = $_.Exception.Message }\n"
            "@{ ok = $ok; looks = $vbe.CommandBars.Looks; error = $m }" + WRITE)
        self.assertTrue(result["ok"])
        self.assertEqual(result["looks"], 1)
        # Attended, the prompt is a real Read-Host, which a -NonInteractive host refuses.
        self.assertIn("NonInteractive", result["error"])

    def test_the_runner_list_is_resolved_and_checked(self) -> None:
        default = script_list("DefaultRunners")
        probe = script_value("ProbeRunner")
        args = (f"-Default @({', '.join(repr(name) for name in default)}) -Probe '{probe}'").replace('"', "'")
        result = self.run_body(
            f"$a = Resolve-RunnerList {args}\n"
            f"$b = Resolve-RunnerList -Requested @('test_excel_ui_runcore, Test_EXCEL_UI_RunAll', "
            f"'Test_EXCEL_UI_RunCore') -IncludeProbe {args}\n"
            f"$c = Resolve-RunnerList -Requested @('Test_EXCEL_UI_RunRibbonSdiProbe') -IncludeProbe {args}\n"
            f"$m = ''\ntry {{ $null = Resolve-RunnerList -Requested @('Test_EXCEL_UI_RunReleaseCertification') "
            f"{args} }} catch {{ $m = $_.Exception.Message }}\n"
            "@{ a = @($a); b = @($b); c = @($c); error = $m }" + WRITE)
        self.assertEqual(result["a"], default)
        self.assertEqual(result["b"], ["Test_EXCEL_UI_RunCore", "Test_EXCEL_UI_RunAll", probe])
        self.assertEqual(result["c"], [probe])
        self.assertIn("Unknown runner Test_EXCEL_UI_RunReleaseCertification.", result["error"])

    def test_a_macro_outcome_needs_a_verdict_and_a_log_that_ends_with_it(self) -> None:
        log = self.temp / "runner.log"
        result = self.run_body(
            "$log = Join-Path $env:TEMPDIR 'runner.log'\n"
            "function Outcome([string]$r, [string]$e = '') { $o = Get-MacroOutcome -Returned $r -Raised $e "
            "-LogPath $log; return @{ passed = $o.Passed; detail = $o.Detail } }\n"
            "$missing = Outcome 'RESULT: PASS'\n"
            "[IO.File]::WriteAllText($log, \"START`r`nline`r`nRESULT: PASS`r`n\", $Utf8NoBom)\n"
            "$pass = Outcome 'RESULT: PASS'\n"
            "$mismatch = Outcome 'RESULT: FAIL | Error 5 | Source: S | text'\n"
            "$logFailed = Outcome 'RESULT: PASS | LOG WRITE FAILED: error 75 writing x: Path/File access error'\n"
            "$none = Outcome ''\n$raised = Outcome '' 'Cannot run the macro'\n"
            "[IO.File]::AppendAllText($log, \"RESULT: FAIL | Error 5 | Source: S | text`r`n\", $Utf8NoBom)\n"
            "$fail = Outcome 'RESULT: FAIL | Error 5 | Source: S | text'\n"
            "@{ missing = $missing; pass = $pass; mismatch = $mismatch; logFailed = $logFailed; none = $none; "
            "raised = $raised; fail = $fail }" + WRITE)
        self.assertTrue(log.exists())
        self.assertEqual(result["pass"], {"passed": True, "detail": "RESULT: PASS"})
        self.assertEqual(result["fail"], {"passed": False, "detail": "RESULT: FAIL | Error 5 | Source: S | text"})
        self.assertFalse(result["missing"]["passed"])
        self.assertIn("the log does not end with that verdict", result["missing"]["detail"])
        self.assertIn("the log does not end with that verdict", result["mismatch"]["detail"])
        self.assertEqual(result["logFailed"], {
            "passed": False,
            "detail": "RESULT: PASS; the log could not be written: error 75 writing x: Path/File access error"})
        self.assertEqual(result["none"], {"passed": False, "detail": "no verdict was returned: ''"})
        self.assertEqual(result["raised"], {"passed": False, "detail": "Application.Run raised: Cannot run the macro"})

    def test_an_incomplete_verdict_is_a_skip_and_never_a_pass(self) -> None:
        result = self.run_body(
            "$log = Join-Path $env:TEMPDIR 'demo.log'\n"
            "$incomplete = 'RESULT: INCOMPLETE | skipped 2 of 49 scenarios: S45_ResetManualBreaks, S47_BreaksShown'\n"
            "function Outcome([string]$returned, [string]$logLast, [string]$raised = '') {\n"
            "  [IO.File]::WriteAllText($log, \"START`r`n  SKIP S45_X: x`r`n$logLast`r`n\", $Utf8NoBom)\n"
            "  $o = Get-MacroOutcome -Returned $returned -Raised $raised -LogPath $log\n"
            "  return @{ status = (Get-OutcomeStatus $o); passed = $o.Passed; detail = $o.Detail } }\n"
            "$skip = Outcome $incomplete $incomplete\n"
            "$pass = Outcome 'RESULT: PASS' 'RESULT: PASS'\n"
            "$fail = Outcome 'RESULT: FAIL | Error 5 | Source: S | text' 'RESULT: FAIL | Error 5 | Source: S | text'\n"
            "$mismatch = Outcome $incomplete 'RESULT: PASS'\n"
            "$logFailed = Outcome ($incomplete + ' | LOG WRITE FAILED: error 75') $incomplete\n"
            "$raised = Outcome '' '' 'Cannot run the macro'\n"
            "$verdicts = @('RESULT: INCOMPLETE', 'RESULT: INCOMPLETE | x', 'RESULT: INCOMPLETEX', "
            "'RESULT: PARTIAL | x') | "
            "ForEach-Object { (ConvertFrom-Verdict $_).HasVerdict }\n"
            "@{ skip = $skip; pass = $pass; fail = $fail; mismatch = $mismatch; logFailed = $logFailed; "
            "raised = $raised; verdicts = @($verdicts) }" + WRITE)
        self.assertEqual(result["skip"], {"status": "SKIP", "passed": False, "detail":
                         "RESULT: INCOMPLETE | skipped 2 of 49 scenarios: S45_ResetManualBreaks, S47_BreaksShown"})
        self.assertEqual((result["pass"]["status"], result["pass"]["passed"]), ("PASS", True))
        self.assertEqual((result["fail"]["status"], result["fail"]["passed"]), ("FAIL", False))
        for name in ("mismatch", "logFailed", "raised"):
            with self.subTest(name):
                self.assertEqual((result[name]["status"], result[name]["passed"]), ("FAIL", False))
        self.assertEqual(result["verdicts"], [True, True, False, False])

    def test_a_skipped_step_makes_the_run_incomplete_not_passed(self) -> None:
        result = self.run_body(
            "function S([string]$n, [string]$s, [string]$d = '', [string]$l = '') { "
            "[ordered]@{ name = $n; status = $s; detail = $d; log = $l } }\n"
            "$ok = @((S 'a' 'PASS'), (S 'b' 'WARN'), (S 'c' 'INFO'))\n"
            "$skipped = $ok + @(S 'demo-cleanup-checks' 'SKIP' 'RESULT: INCOMPLETE | skipped 1 of 49 scenarios: "
            "S47_BreaksShown' 'demo-cleanup-checks.log')\n"
            "$summary = Format-Summary -Commit ('0' * 40) -HostLine '' -Steps $skipped "
            "-Result (Get-RunResult -Steps $skipped -Cancelled $false).Result\n"
            "@{ skipped = (Get-RunResult -Steps $skipped -Cancelled $false); "
            "failed = (Get-RunResult -Steps ($skipped + @(S 'd' 'FAIL')) -Cancelled $false); "
            "notRun = (Get-RunResult -Steps ($skipped + @(S 'e' 'NOT_RUN')) -Cancelled $false); "
            "cancelled = (Get-RunResult -Steps $skipped -Cancelled $true); summary = @($summary) }" + WRITE)
        self.assertEqual(result["skipped"], {"Result": "INCOMPLETE", "ExitCode": 3})
        self.assertEqual(result["failed"], {"Result": "FAIL", "ExitCode": 1})
        self.assertEqual(result["notRun"], {"Result": "FAIL", "ExitCode": 1})
        self.assertEqual(result["cancelled"], {"Result": "FAIL", "ExitCode": 2})
        self.assertEqual(result["summary"][-2:], [
            "SKIP demo-cleanup-checks: RESULT: INCOMPLETE | skipped 1 of 49 scenarios: S47_BreaksShown "
            "[demo-cleanup-checks.log]",
            "RESULT: INCOMPLETE"])

    def test_results_timings_and_the_summary(self) -> None:
        result = self.run_body(
            "function S([string]$n, [string]$s, [string]$d = '', [string]$l = '') { "
            "[ordered]@{ name = $n; status = $s; detail = $d; log = $l } }\n"
            "$ok = @((S 'a' 'PASS'), (S 'b' 'WARN'), (S 'c' 'INFO'))\n"
            "$bad = $ok + @(S 'd' 'FAIL')\n$notRun = $ok + @(S 'e' 'NOT_RUN')\n"
            "$summary = Format-Summary -Commit ('0' * 40) -HostLine 'Excel 16.0 build 20430, x64' "
            "-Steps @((S 'preflight' 'PASS' 'ok'), (S 'Test_EXCEL_UI_RunAll' 'FAIL' 'RESULT: FAIL | x' "
            "'runner-Test_EXCEL_UI_RunAll.log')) -Result 'FAIL'\n"
            "@{ ok = (Get-RunResult -Steps $ok -Cancelled $false); "
            "bad = (Get-RunResult -Steps $bad -Cancelled $false); "
            "notRun = (Get-RunResult -Steps $notRun -Cancelled $false); "
            "cancelled = (Get-RunResult -Steps $ok -Cancelled $true); "
            "timing = @((Get-TimingStatus 9.9 10), (Get-TimingStatus 10 10), (Get-TimingStatus 10.1 10)); "
            "seconds = (Format-Seconds 12.345); summary = @($summary); "
            "args = (ConvertTo-ArgumentLine @('-File', 'C:\\a b\\x.ps1', '', 'say \"hi\"')) }" + WRITE)
        self.assertEqual(result["ok"], {"Result": "PASS", "ExitCode": 0})
        self.assertEqual(result["bad"], {"Result": "FAIL", "ExitCode": 1})
        self.assertEqual(result["notRun"], {"Result": "FAIL", "ExitCode": 1})
        self.assertEqual(result["cancelled"], {"Result": "FAIL", "ExitCode": 2})
        self.assertEqual(result["timing"], ["PASS", "PASS", "WARN"])
        self.assertEqual(result["seconds"], "12.3")
        self.assertEqual(result["summary"], [
            "Commit: " + "0" * 40,
            "Host: Excel 16.0 build 20430, x64",
            "PASS preflight: ok",
            "FAIL Test_EXCEL_UI_RunAll: RESULT: FAIL | x [runner-Test_EXCEL_UI_RunAll.log]",
            "RESULT: FAIL"])
        self.assertEqual(result["args"], '-File "C:\\a b\\x.ps1" "" "say \\"hi\\""')

    def test_a_step_that_cannot_run_is_a_failure_that_says_why(self) -> None:
        result = self.run_body(
            "$script:Steps = @((New-CheckStep 'a' 'a.log'), (New-CheckStep 'b' ''), (New-CheckStep 'c' ''))\n"
            "Set-CheckStep 'a' 'PASS' 'done' -Seconds 1.234 | Out-Null\n"
            "Skip-CheckStep @('a', 'b') 'a step failed' | Out-Null\n"
            "$m = ''\ntry { Set-CheckStep 'z' 'PASS' } catch { $m = $_.Exception.Message }\n"
            "@{ steps = $script:Steps; error = $m }" + WRITE)
        steps = {step["name"]: step for step in result["steps"]}
        self.assertEqual((steps["a"]["status"], steps["a"]["detail"], steps["a"]["seconds"]), ("PASS", "done", 1.23))
        self.assertEqual((steps["b"]["status"], steps["b"]["detail"]), ("FAIL", "not run: a step failed"))
        self.assertEqual(steps["c"]["status"], "NOT_RUN")
        self.assertEqual(result["error"], "The step 'z' is not in this run.")


@unittest.skipUnless(PWSH, "PowerShell 7 (pwsh) is not installed")
@unittest.skipIf(os.name == "nt", "on Windows the preflight could reach a real Excel installation")
class ScriptRunTests(unittest.TestCase):
    """The whole script, run where its preflight fails before anything starts Excel."""

    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.repo = Path(temp.name) / "repo"
        self.bundle = Path(temp.name) / "bundle"
        (self.repo / ".github").mkdir(parents=True)
        shutil.copyfile(POLICY, self.repo / ".github" / "excel-evidence-policy.json")
        (self.repo / "a.txt").write_text("a\n", encoding="utf-8")
        git(self.repo, "init", "-q")
        git(self.repo, "config", "user.name", "Test")
        git(self.repo, "config", "user.email", "test@example.invalid")
        git(self.repo, "add", "-A")
        git(self.repo, "commit", "-q", "-m", "fixture")

    def run_script(self, *extra: str, bundle: Path | None = None) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [str(PWSH), "-NoProfile", "-NonInteractive", "-File", str(SCRIPT), "-RepoRoot", str(self.repo),
             "-BundleDir", str(bundle or self.bundle), *extra],
            capture_output=True, text=True, timeout=120, check=False)

    def test_a_failed_preflight_still_writes_the_summary_and_the_session(self) -> None:
        completed = self.run_script("-Runners", "Test_EXCEL_UI_RunCore,Test_EXCEL_UI_RunAll", "-IncludeRibbonProbe")
        self.assertEqual(completed.returncode, 1, completed.stdout + completed.stderr)
        summary = (self.bundle / "summary.txt").read_text(encoding="utf-8").splitlines()
        self.assertTrue(summary[0].startswith("Commit: "))
        self.assertEqual(summary[-1], "RESULT: FAIL")
        steps = [line.split(":", 1)[0] for line in summary[1:-1]]
        self.assertEqual(steps, [
            "FAIL preflight", "FAIL certification", "FAIL open-development", "FAIL Test_EXCEL_UI_RunCore",
            "FAIL Test_EXCEL_UI_RunAll", "FAIL Test_EXCEL_UI_RunRibbonSdiProbe", "FAIL close-development",
            "FAIL demo-build", "FAIL demo-open", "FAIL demo-compile", "FAIL demo-cleanup-checks",
            "FAIL demo-create-sheet-1", "FAIL demo-create-sheet-2", "FAIL default-printer", "FAIL demo-titlebar",
            "FAIL close-demo"])
        self.assertTrue(all(line.endswith("not run: the preflight failed") for line in summary[2:-1]))
        self.assertNotIn("[", "".join(summary[1:-1]))
        self.assertIn("RESULT: FAIL", completed.stdout)
        session = json.loads((self.bundle / "session.json").read_text(encoding="utf-8-sig"))
        self.assertRegex(session["commit"], r"^[0-9a-f]{40}$")
        self.assertEqual(session["tool"], "tools/Check-EXCEL_UI.ps1")
        self.assertEqual((session["result"], session["exit_code"]), ("FAIL", 1))
        self.assertEqual(session["runners"],
                         ["Test_EXCEL_UI_RunCore", "Test_EXCEL_UI_RunAll", "Test_EXCEL_UI_RunRibbonSdiProbe"])
        self.assertEqual(session["slow_build_seconds"], 10)
        self.assertIsNotNone(session["finished_utc"])
        self.assertEqual(len(session["steps"]), len(steps))
        self.assertTrue((self.bundle / "session.log").exists())
        self.assertFalse((self.bundle / "certification").exists())

    def test_the_session_records_how_the_macros_are_enabled(self) -> None:
        for extra, expected in (((), "msoAutomationSecurityByUI"),
                                (("-Unattended",), "msoAutomationSecurityLow (-Unattended)")):
            with self.subTest(extra=extra):
                bundle = self.bundle.with_name("bundle-" + str(len(extra)))
                completed = self.run_script(*extra, bundle=bundle)
                self.assertEqual(completed.returncode, 1, completed.stdout + completed.stderr)
                session = json.loads((bundle / "session.json").read_text(encoding="utf-8-sig"))
                self.assertEqual(session["host"]["trust_center"]["automation_security"], expected)

    def test_the_script_stops_before_the_bundle_on_a_bad_start(self) -> None:
        with self.subTest("unknown runner"):
            completed = self.run_script("-Runners", "Test_EXCEL_UI_RunNothing")
            self.assertEqual(completed.returncode, 1)
            self.assertIn("ERROR: Unknown runner Test_EXCEL_UI_RunNothing.", completed.stdout)
        with self.subTest("bundle inside the checkout"):
            completed = self.run_script(bundle=self.repo / "bundle")
            self.assertEqual(completed.returncode, 1)
            self.assertIn("ERROR: the bundle folder must be outside the checkout", completed.stdout)
        with self.subTest("bundle not empty"):
            self.bundle.mkdir()
            (self.bundle / "x.txt").write_text("x", encoding="utf-8")
            completed = self.run_script()
            self.assertEqual(completed.returncode, 1)
            self.assertIn("ERROR: the bundle folder is not empty", completed.stdout)
            shutil.rmtree(self.bundle)
        with self.subTest("changed working tree"):
            (self.repo / "a.txt").write_text("b\n", encoding="utf-8")
            completed = self.run_script()
            self.assertEqual(completed.returncode, 1)
            self.assertIn("ERROR: the working tree has changes.", completed.stdout)
        self.assertFalse(self.bundle.exists())
        self.assertFalse((self.repo / "bundle").exists())


if __name__ == "__main__":
    unittest.main()
