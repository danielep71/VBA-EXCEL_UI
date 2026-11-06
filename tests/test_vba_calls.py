"""Tests for the advisory VBA call checker (tools/check_vba_calls.py).

Every rule has a fixture it must report next to a valid twin it must accept,
built from small synthetic modules. The last tests run the command line on a
temporary repository and the checker on this repository's own modules.
"""

from __future__ import annotations

import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_vba_calls as calls  # noqa: E402
from gate_fixture import REPO, git  # noqa: E402

PATTERNS = ["^UI_", "^M_"]


def mod(name: str, body: str, private_module: bool = False) -> str:
    header = f'Attribute VB_Name = "{name}"\nOption Explicit\n'
    if private_module:
        header += "Option Private Module\n"
    return header + body


def cls(name: str, body: str) -> str:
    return ('VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\n'
            f'Attribute VB_Name = "{name}"\nAttribute VB_PredeclaredId = False\nOption Explicit\n' + body)


def frm(name: str, body: str) -> str:
    return ('VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ' + name + '\n   Caption = "x"\nEnd\n'
            f'Attribute VB_Name = "{name}"\nAttribute VB_PredeclaredId = True\nOption Explicit\n' + body)


def manifest_for(files: dict[str, str], gating: bool = True, listed: list[str] | None = None) -> dict[str, Any]:
    config: dict[str, Any] = {"modules": listed if listed is not None else sorted(files), "gating": gating}
    return {"project_name_patterns": PATTERNS, "macro_name_functions": ["M_Qualify"],
            "document_modules": [], "configurations": {"fixture": config}}


def analyze(files: dict[str, str], gating: bool = True, listed: list[str] | None = None) -> dict[str, Any]:
    return calls.analyze_project(files, manifest_for(files, gating, listed))


def codes(report: dict[str, Any], severity: str = "error", configuration: str = "fixture") -> list[str]:
    findings = report["configurations"][configuration]["findings"]
    return sorted(f["code"] for f in findings if f["severity"] == severity)


def entries(report: dict[str, Any], configuration: str = "fixture") -> list[tuple[str, str]]:
    return [(e["channel"], e["status"]) for e in report["configurations"][configuration]["entry_points"]]


class ResolutionTests(unittest.TestCase):
    def test_valid_local_and_cross_module_calls(self) -> None:
        report = analyze({
            "a.bas": mod("M_A", "Public Sub UI_Run()\n    UI_Helper 1\n    M_B.UI_Remote\n    Call UI_Remote\nEnd Sub\n"
                                "Private Sub UI_Helper(ByVal x As Long)\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub UI_Remote()\nEnd Sub\n", private_module=True),
        })
        self.assertEqual(report["status"], "pass")
        self.assertEqual(codes(report), [])
        self.assertGreaterEqual(report["configurations"]["fixture"]["stats"]["resolved-calls"], 3)

    def test_missing_unqualified_target_only_for_project_names(self) -> None:
        missing = analyze({"a.bas": mod("M_A", "Public Sub UI_Run()\n    UI_Gone 1\nEnd Sub\n")})
        self.assertEqual(codes(missing), ["VBA-CALL-001"])
        self.assertEqual(missing["status"], "fail")
        host = analyze({"a.bas": mod("M_A", "Public Sub UI_Run()\n    Beep\n    HostThing 1\n"
                                            "    x = Len(\"a\")\nEnd Sub\n")})
        self.assertEqual(codes(host), [])
        self.assertEqual(host["status"], "pass")

    def test_project_names_match_case_insensitively(self) -> None:
        for name in ("UI_GONE", "ui_gone", "Ui_Gone"):
            with self.subTest(name=name):
                report = analyze({"a.bas": mod("M_A", f"Public Sub UI_Run()\n    {name} 1\nEnd Sub\n")})
                self.assertEqual(codes(report), ["VBA-CALL-001"])

    def test_missing_qualified_member_of_module_class_and_form(self) -> None:
        files = {
            "a.bas": mod("M_A", "Public Sub UI_Run()\n    Dim c As cThing\n    Set c = New cThing\n    c.Missing\n"
                                "    c.Present 1\n    M_B.UI_Nope\n    UF_X.lblUnknown.Caption = \"x\"\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub UI_Yes()\nEnd Sub\n"),
            "c.cls": cls("cThing", "Public Sub Present(ByVal x As Long)\nEnd Sub\n"),
            "f.frm": frm("UF_X", "Private Sub UserForm_Initialize()\nEnd Sub\n"),
        }
        report = analyze(files)
        self.assertEqual(codes(report), ["VBA-CALL-002", "VBA-CALL-002"])
        unknown = report["configurations"]["fixture"]["unknown"]
        self.assertIn("member not declared in form code (control or UserForm member)", unknown)

    def test_private_member_from_another_module(self) -> None:
        report = analyze({
            "a.bas": mod("M_A", "Public Sub UI_Run()\n    UI_Secret\n    M_B.UI_Secret\nEnd Sub\n"),
            "b.bas": mod("M_B", "Private Sub UI_Secret()\nEnd Sub\nPublic Sub UI_Own()\n    UI_Secret\nEnd Sub\n"),
        })
        self.assertEqual(codes(report), ["VBA-CALL-003", "VBA-CALL-003"])
        same_module = analyze({"b.bas": mod("M_B", "Private Sub UI_Secret()\nEnd Sub\n"
                                                   "Public Sub UI_Own()\n    UI_Secret\n    M_B.UI_Secret\nEnd Sub\n")})
        self.assertEqual(codes(same_module), [])

    def test_option_private_module_stays_project_visible(self) -> None:
        report = analyze({
            "a.bas": mod("M_A", "Public Sub UI_Run()\n    UI_Inner\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub UI_Inner()\nEnd Sub\n", private_module=True),
        })
        self.assertEqual(codes(report), [])

    def test_ambiguous_public_name_versus_qualified_and_own_module(self) -> None:
        files = {
            "a.bas": mod("M_A", "Public Sub UI_Same()\nEnd Sub\nPublic Sub UI_Local()\n    UI_Same\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub UI_Same()\nEnd Sub\n"),
            "c.bas": mod("M_C", "Public Sub UI_Caller()\n    M_A.UI_Same\n    M_B.UI_Same\n    UI_Same\nEnd Sub\n"),
        }
        report = analyze(files)
        findings = report["configurations"]["fixture"]["findings"]
        ambiguous = [f for f in findings if f["code"] == "VBA-CALL-004"]
        self.assertEqual(len(ambiguous), 1)
        self.assertEqual(ambiguous[0]["module"], "M_C")
        self.assertEqual({c["module"] for c in ambiguous[0]["candidates"]}, {"M_A", "M_B"})
        self.assertEqual(codes(report, "warning"), ["VBA-CALL-040"])
        unique = analyze({"a.bas": files["a.bas"], "c.bas": files["c.bas"].replace("    M_B.UI_Same\n", "")})
        self.assertEqual(codes(unique), [])
        self.assertEqual(codes(unique, "warning"), [])

    def test_local_shadowing_is_not_a_call(self) -> None:
        report = analyze({
            "a.bas": mod("M_A", "Public Sub UI_Run(ByVal UI_Param As Long)\n    Dim UI_Target(3) As Long\n"
                                "    UI_Target(1) = UI_Param\n    x = UI_Target(2) + UI_Param\nEnd Sub\n"),
            "b.bas": mod("M_B", "Public Sub UI_Target(ByVal a As Long, ByVal b As Long)\nEnd Sub\n"),
        })
        self.assertEqual(codes(report), [])

    def test_comparison_inside_argument_is_not_an_assignment(self) -> None:
        report = analyze({"a.bas": mod("M_A", "Public Sub UI_Assert(ByVal t As String, ByVal c As Boolean)\nEnd Sub\n"
                                              "Public Sub UI_Run()\n    UI_Assert \"x\", 1 = 1\n"
                                              "    UI_Assert \"y\", _\n        2 = 2\nEnd Sub\n")})
        self.assertEqual(codes(report), [])

    def test_class_without_default_instance_is_not_a_receiver(self) -> None:
        predeclared = cls("cPre", "Public Sub Present(ByVal x As Long)\nEnd Sub\n").replace(
            "VB_PredeclaredId = False", "VB_PredeclaredId = True")
        files = {
            "a.bas": mod("M_A", "Public Sub UI_Run()\n    cThing.Present 1\nEnd Sub\n"),
            "c.cls": cls("cThing", "Public Sub Present(ByVal x As Long)\nEnd Sub\n"
                                   "Public Sub Again()\n    Me.Present 1\nEnd Sub\n"),
            "p.cls": predeclared,
        }
        self.assertEqual(codes(analyze(files)), ["VBA-CALL-005"])
        files["a.bas"] = mod("M_A", "Public Sub UI_Run()\n    Dim c As cThing\n    Set c = New cThing\n"
                                    "    c.Present 1\n    cPre.Present 1\nEnd Sub\n")
        self.assertEqual(codes(analyze(files)), [])

    def test_recursive_calls_are_argument_checked(self) -> None:
        def check(body: str) -> list[str]:
            return codes(analyze({"a.bas": mod("M_A", body)}))

        self.assertEqual(check("Public Sub UI_Recur(ByVal n As Long)\n    UI_Recur\nEnd Sub\n"), ["VBA-CALL-010"])
        self.assertEqual(check("Public Sub UI_Recur(ByVal n As Long)\n    UI_Recur n - 1\n"
                               "    Call UI_Recur(n - 1)\nEnd Sub\n"), [])
        self.assertEqual(check("Public Function UI_F(ByVal n As Long) As Long\n    UI_F = n\n"
                               "    If n > 0 Then UI_F = UI_F + UI_F(n - 1)\nEnd Function\n"), [])

    def test_return_value_of_array_function_is_not_a_call(self) -> None:
        body = "Public Function UI_Arr() As Long()\n    Dim x As Long\n    x = UI_Arr(0)\nEnd Function\n"
        self.assertEqual(codes(analyze({"a.bas": mod("M_A", body)})), [])


class ArgumentTests(unittest.TestCase):
    TARGET = ("Public Sub UI_T(ByVal A As Long, Optional ByVal B As Long = 1, Optional C As Variant)\nEnd Sub\n"
              "Public Sub UI_P(ByVal A As Long, ParamArray Rest() As Variant)\nEnd Sub\n"
              "Public Function UI_F(ByVal A As Long) As Long\nEnd Function\n")

    def check(self, body: str) -> list[str]:
        return codes(analyze({"a.bas": mod("M_A", self.TARGET + "Public Sub UI_Run()\n" + body + "End Sub\n")}))

    def test_valid_shapes(self) -> None:
        self.assertEqual(self.check("    UI_T 1\n    UI_T 1, , 3\n    UI_T C:=3, A:=1\n    UI_T 1, C:=3\n"
                                    "    Call UI_T(1, 2)\n    UI_P 1, 2, 3, 4\n    UI_P 1\n"
                                    "    x = UI_F(1) + UI_F(A:=2)\n    UI_T (1)\n    UI_T (1), 2\n"
                                    "    UI_T _\n        1, _\n        2\n"), [])

    def test_too_few(self) -> None:
        self.assertEqual(self.check("    UI_T\n    x = UI_F()\n    x = UI_F\n    UI_P\n"), ["VBA-CALL-010"] * 4)

    def test_too_many_without_paramarray(self) -> None:
        self.assertEqual(self.check("    UI_T 1, 2, 3, 4\n    x = UI_F(1, 2)\n"), ["VBA-CALL-011"] * 2)

    def test_omitted_required_versus_omitted_optional(self) -> None:
        self.assertEqual(self.check("    UI_T , 2\n"), ["VBA-CALL-010", "VBA-CALL-012"])
        self.assertEqual(self.check("    UI_T 1, , 3\n"), [])

    def test_named_arguments(self) -> None:
        self.assertEqual(self.check("    UI_T 1, Z:=2\n"), ["VBA-CALL-013"])
        self.assertEqual(self.check("    UI_T 1, A:=2\n"), ["VBA-CALL-014"])
        self.assertEqual(self.check("    UI_T A:=1, A:=2\n"), ["VBA-CALL-014"])
        self.assertEqual(self.check("    UI_P 1, Rest:=2\n"), ["VBA-CALL-015"])
        self.assertEqual(self.check("    UI_T A:=1, 2\n"), ["VBA-CALL-016"])
        self.assertEqual(self.check("    UI_T B:=2, A:=1\n    UI_P A:=1\n"), [])

    def test_qualified_calls_are_argument_checked(self) -> None:
        files = {"t.bas": mod("M_T", self.TARGET)}
        bad = analyze({**files, "a.bas": mod("M_A", "Public Sub UI_Run()\n    M_T.UI_T\n"
                                                    "    x = M_T.UI_F(1, 2)\nEnd Sub\n")})
        self.assertEqual(codes(bad), ["VBA-CALL-010", "VBA-CALL-011"])
        good = analyze({**files, "a.bas": mod("M_A", "Public Sub UI_Run()\n    M_T.UI_T 1\n"
                                                     "    x = M_T.UI_F(1)\nEnd Sub\n")})
        self.assertEqual(codes(good), [])

    def test_with_member_as_first_argument(self) -> None:
        # "UI_T .A = 1, 2" passes a With member: a call with two arguments, not an assignment.
        box = cls("cBox", "Public A As Long\n")

        def run(body: str) -> list[str]:
            return codes(analyze({"c.cls": box, "a.bas": mod("M_A", self.TARGET + "Public Sub UI_Run()\n"
                                                                    "    Dim b As cBox\n    With b\n" + body
                                                                    + "    End With\nEnd Sub\n")}))

        self.assertEqual(run("        UI_T .A = 1, 2\n        UI_P .A = 1\n"), [])
        self.assertEqual(run("        UI_F.A = 1\n"), ["VBA-CALL-010"])
        self.assertEqual(run("        UI_T 1, 2, .A = 1, 4\n"), ["VBA-CALL-011"])


class PropertyTests(unittest.TestCase):
    CLASS = ("Private m As Long\nPublic Property Get Value(ByVal i As Long) As Long\nEnd Property\n"
             "Public Property Let Value(ByVal i As Long, ByVal v As Long)\nEnd Property\n"
             "Public Property Get ReadOnly() As Long\nEnd Property\n"
             "Public Property Set Target(ByVal o As Object)\nEnd Property\n")

    def check(self, body: str, class_body: str | None = None) -> list[str]:
        return codes(analyze({"c.cls": cls("cBox", class_body or self.CLASS),
                              "a.bas": mod("M_A", "Public Sub UI_Run()\n    Dim b As cBox\n" + body + "End Sub\n")}))

    def test_valid_family(self) -> None:
        self.assertEqual(self.check("    b.Value(1) = 2\n    x = b.Value(1)\n    x = b.ReadOnly\n"
                                    "    Set b.Target = Nothing\n"), [])

    def test_accessor_required_by_the_use(self) -> None:
        self.assertEqual(self.check("    b.ReadOnly = 1\n"), ["VBA-CALL-017"])
        self.assertEqual(self.check("    b.Target = 1\n"), ["VBA-CALL-017"])
        self.assertEqual(self.check("    x = b.Value(1, 2)\n"), ["VBA-CALL-011"])
        self.assertEqual(self.check("    b.Value = 2\n"), ["VBA-CALL-010"])

    def test_conflicting_declarations(self) -> None:
        duplicate = self.CLASS + "Public Property Get ReadOnly() As Long\nEnd Property\n"
        self.assertEqual(self.check("", duplicate), ["VBA-CALL-020"])
        clash = self.CLASS + "Public Sub ReadOnly()\nEnd Sub\n"
        self.assertEqual(self.check("", clash), ["VBA-CALL-020"])
        module = analyze({"a.bas": mod("M_A", "Public Sub UI_Twice()\nEnd Sub\nPublic Sub UI_Twice()\nEnd Sub\n")})
        self.assertEqual(codes(module), ["VBA-CALL-020"])


class DynamicTargetTests(unittest.TestCase):
    PROCS = ("Public Sub UI_Tick()\nEnd Sub\nPrivate Sub UI_Hidden()\nEnd Sub\n"
             "Public Function M_Qualify(ByVal n As String) As String\nEnd Function\n")

    def report(self, body: str, extra: str = "") -> dict[str, Any]:
        return analyze({"a.bas": mod("M_A", self.PROCS + "Public Sub UI_Run()\n    Dim v As String\n"
                                                         "    Dim o As Object\n    Dim c As cHook\n"
                                                         "    Dim s As Object\n" + body + "End Sub\n"),
                        "c.cls": cls("cHook", "Public Sub Fire()\nEnd Sub\n" + extra)})

    def test_literal_targets_that_resolve(self) -> None:
        good = self.report('    Application.Run "UI_Tick"\n    Application.OnTime Now, "M_A.UI_Tick"\n'
                           '    Application.OnKey "^d", M_Qualify("UI_Tick")\n'
                           '    x = Application.Run("\'Book.xlsm\'!UI_Tick")\n'
                           '    CallByName c, "Fire", VbMethod\n    s.OnAction = "UI_Tick"\n'
                           '    With s\n        .OnAction = "M_A.UI_Tick"\n    End With\n'
                           '    s.OnAction = "\'" & ThisWorkbook.Name & "\'!UI_Tick"\n'
                           '    s.OnAction = ThisWorkbook.Name & "!UI_Tick"\n')
        self.assertEqual(codes(good), [])
        self.assertEqual(codes(good, "warning"), [])
        statuses = entries(good)
        for channel in ("Application.Run", "Application.OnTime", "Application.OnKey", "CallByName", "OnAction"):
            self.assertIn((channel, "resolved"), statuses)
        self.assertNotIn("missing", [status for _, status in statuses])

    def test_constants_are_evaluated_whole(self) -> None:
        # A constant joined from literals names the whole result; anything else
        # in an initializer leaves the constant unknown, never its first literal.
        joined = self.report('    Const Target As String = "UI_" & "Tick"\n    Application.Run Target\n')
        joined_module = analyze({"a.bas": mod("M_A", 'Private Const M_TARGET As String = "UI_" & "Tick"\n'
                                              + self.PROCS + "Public Sub UI_Run()\n    Application.Run M_TARGET\n"
                                              "End Sub\n")})
        self.assertEqual(codes(joined), [])
        self.assertIn(("Application.Run", "resolved"), entries(joined))
        self.assertEqual(codes(joined_module), [])
        self.assertIn(("Application.Run", "resolved"), entries(joined_module))
        missing = self.report('    Const Target As String = "UI_" & "Gone"\n    Application.Run Target\n')
        self.assertEqual(codes(missing), ["VBA-CALL-030"])
        hidden = self.report('    Const Target As String = "UI_Tick" & Suffix\n    Application.Run Target\n')
        self.assertEqual(codes(hidden), [])
        self.assertIn(("Application.Run", "unknown-dynamic"), entries(hidden))

    def test_workbook_qualified_targets_stay_outside_the_project(self) -> None:
        # A target in another workbook says nothing about this project's
        # procedures, even when one of them has the same name.
        report = self.report('    Application.Run "Other.xlsm!UI_Hidden"\n'
                             '    Application.Run "\'Other Book.xlsm\'!UI_Tick"\n'
                             '    s.OnAction = "\'Other.xlsm\'!UI_Gone"\n')
        self.assertEqual(codes(report), [])
        self.assertEqual(codes(report, "warning"), [])
        self.assertEqual([status for _, status in entries(report)], ["unknown-external"] * 3)

    def test_literal_targets_that_do_not_resolve(self) -> None:
        bad = self.report('    Application.Run "UI_Gone"\n'
                          '    Excel.Application.OnTime EarliestTime:=Now, Procedure:="UI_Gone2"\n'
                          '    Application.OnKey "^d", "M_A.Helper"\n'
                          '    CallByName c, "Missing", VbMethod\n    v = M_Qualify("UI_Gone3")\n'
                          '    s.OnAction = "UI_Gone4"\n'
                          '    s.OnAction = "\'" & ThisWorkbook.Name & "\'!UI_Gone5"\n')
        self.assertEqual(codes(bad), ["VBA-CALL-030"] * 7)
        private = self.report('    Application.Run "UI_Hidden"\n')
        self.assertEqual(codes(private), [])
        self.assertEqual(codes(private, "warning"), ["VBA-CALL-031"])

    def test_dynamic_or_foreign_targets_are_unknown_not_defects(self) -> None:
        report = self.report('    Application.Run v\n    CallByName o, "Anything", VbMethod\n    s.OnAction = v\n'
                             '    s.OnAction = ActiveWorkbook.Name & "!UI_Gone"\n'
                             '    Application.Run "\'Other.xlsm\'!UI_Gone"\n    Application.Run "Elsewhere"\n')
        self.assertEqual(codes(report), [])
        self.assertEqual(report["status"], "pass")
        statuses = [status for _, status in entries(report)]
        self.assertEqual(statuses.count("unknown-dynamic"), 4)
        self.assertEqual(statuses.count("unknown-external"), 2)

    def test_addressof_targets(self) -> None:
        files = {"b.bas": mod("M_B", "Private Function UI_Proc(ByVal h As Long) As Long\nEnd Function\n"
                                     "Public UI_Value As Long\n")}
        good = analyze({**files, "a.bas": mod("M_A", "Private Function UI_Own(ByVal h As Long) As Long\n"
                                                      "End Function\nPublic Sub UI_Run()\n    x = AddressOf UI_Own\n"
                                                      "End Sub\n")})
        self.assertEqual(codes(good), [])
        self.assertIn(("AddressOf", "resolved"), entries(good))
        private = analyze({**files, "a.bas": mod("M_A", "Public Sub UI_Run()\n    x = AddressOf UI_Proc\nEnd Sub\n")})
        self.assertEqual(codes(private), ["VBA-CALL-003"])
        value = analyze({**files, "a.bas": mod("M_A", "Public Sub UI_Run()\n    x = AddressOf UI_Value\nEnd Sub\n")})
        self.assertEqual(codes(value), ["VBA-CALL-030"])
        missing = analyze({**files, "a.bas": mod("M_A", "Public Sub UI_Run()\n    x = AddressOf UI_Gone\nEnd Sub\n")})
        self.assertEqual(codes(missing), ["VBA-CALL-030"])

    def test_literal_target_in_a_known_module_is_checked_whatever_its_name(self) -> None:
        body = "Public Sub UI_Run()\n    Application.Run \"M_A.Missing\"\nEnd Sub\nPublic Sub Helper()\nEnd Sub\n"
        self.assertEqual(codes(analyze({"a.bas": mod("M_A", body)})), ["VBA-CALL-030"])
        valid = body.replace("M_A.Missing", "M_A.Helper")
        self.assertEqual(codes(analyze({"a.bas": mod("M_A", valid)})), [])

    def test_event_handlers_are_entry_points(self) -> None:
        report = self.report("", "Private WithEvents mApp As Excel.Application\n"
                                 "Private Sub mApp_SheetChange(ByVal S As Object, ByVal T As Object)\nEnd Sub\n"
                                 "Private Sub Class_Initialize()\nEnd Sub\n")
        self.assertIn(("event-handler (WithEvents)", "external-entry"), entries(report))
        self.assertIn(("event-handler (class)", "external-entry"), entries(report))
        self.assertEqual(codes(report), [])


class MacroParameterTests(unittest.TestCase):
    """A literal passed to a parameter that a procedure hands to OnAction is checked at the call."""

    BUTTON = ("Public Sub UI_Button(ByVal Shp As Object, Optional ByVal ActionMacro As String = \"{default}\")\n"
              "    If Len(ActionMacro) = 0 Then\n"
              "        Shp.OnAction = \"'\" & ThisWorkbook.Name & \"'!UI_Tick\"\n"
              "    ElseIf InStr(1, ActionMacro, \"!\") > 0 Then\n"
              "        Shp.OnAction = ActionMacro\n"
              "    Else\n"
              "        Shp.OnAction = \"'\" & ThisWorkbook.Name & \"'!\" & ActionMacro\n"
              "    End If\n"
              "End Sub\n"
              "Public Sub UI_Tick()\nEnd Sub\n")

    def report(self, caller: str, default: str = "UI_Tick", button: str | None = None) -> dict[str, Any]:
        source = (button or self.BUTTON).replace("{default}", default)
        return analyze({"b.bas": mod("M_B", source),
                        "a.bas": mod("M_A", "Public Sub UI_Run(ByVal s As Object)\n" + caller + "End Sub\n")})

    def test_literal_arguments_resolve(self) -> None:
        report = self.report("    UI_Button s, \"UI_Tick\"\n    M_B.UI_Button s, ActionMacro:=\"M_B.UI_Tick\"\n"
                             "    UI_Button s\n    UI_Button _\n        s, _\n        \"UI_Tick\"\n")
        self.assertEqual(codes(report), [])
        through = [e for e in entries(report) if "through UI_Button" in e[0]]
        self.assertEqual([status for _, status in through], ["resolved"] * 3)
        self.assertIn(("OnAction", "parameter"), entries(report))

    def test_missing_literal_argument_or_default_is_a_defect(self) -> None:
        report = self.report("    UI_Button s, \"UI_Gone\"\n    UI_Button s, ActionMacro:=\"UI_Gone2\"\n")
        self.assertEqual(codes(report), ["VBA-CALL-030"] * 2)
        messages = [f["message"] for f in report["configurations"]["fixture"]["findings"]]
        self.assertTrue(all("through UI_Button argument ActionMacro" in m for m in messages))
        omitted = self.report("    UI_Button s\n", default="UI_NoSuchDefault")
        self.assertEqual(codes(omitted), ["VBA-CALL-030"])
        supplied = self.report("    UI_Button s, \"UI_Tick\"\n", default="UI_NoSuchDefault")
        self.assertEqual(codes(supplied), [])

    def test_parameters_are_followed_through_other_procedures(self) -> None:
        relay = ("Public Sub UI_Relay(ByVal s As Object, ByVal Macro As String)\n"
                 "    UI_Button s, Macro\nEnd Sub\n")
        report = self.report("    UI_Relay s, \"UI_Gone\"\n    UI_Relay s, \"UI_Tick\"\n", button=self.BUTTON + relay)
        self.assertEqual(codes(report), ["VBA-CALL-030"])

    def test_parameters_that_may_change_are_unknown(self) -> None:
        reassigned = self.BUTTON.replace(
            "    If Len(", "    If Len(ActionMacro) = 1 Then ActionMacro = \"x\"\n    If Len(")
        report = self.report("    UI_Button s, \"UI_Gone\"\n", button=reassigned)
        self.assertEqual(codes(report), [])
        by_reference = self.BUTTON.replace("Optional ByVal ActionMacro", "Optional ActionMacro")
        report = self.report("    UI_Button s, \"UI_Gone\"\n", button=by_reference)
        self.assertEqual(codes(report), [])
        local = ("Public Sub UI_Grid(ByVal s As Object, ByVal Spec As Variant)\n    Dim Macro As String\n"
                 "    Macro = \"UI_Gone\"\n    Macro = CStr(Spec)\n    UI_Button s, Macro\nEnd Sub\n")
        report = self.report("    UI_Grid s, \"UI_Gone\"\n", button=self.BUTTON + local)
        self.assertEqual(codes(report), [])
        self.assertIn(("OnAction (through UI_Button argument ActionMacro)", "unknown-dynamic"), entries(report))

    def test_parameters_passed_by_reference_may_change(self) -> None:
        # A procedure that receives the parameter by reference may replace it
        # before the sink reads it, so the literal at the call is not checked.
        fixer = ("Public Sub UI_Fix(Macro As String)\n    Macro = \"UI_Tick\"\nEnd Sub\n")
        exposed = self.BUTTON.replace("    If Len(", "    UI_Fix ActionMacro\n    If Len(") + fixer
        report = self.report("    UI_Button s, \"UI_NotAMacro\"\n", button=exposed)
        self.assertEqual(codes(report), [])
        self.assertNotIn(("OnAction", "parameter"), entries(report))
        called = self.BUTTON.replace("    If Len(", "    Call UI_Fix(ActionMacro)\n    If Len(") + fixer
        self.assertEqual(codes(self.report("    UI_Button s, \"UI_NotAMacro\"\n", button=called)), [])
        external = self.BUTTON.replace("    If Len(", "    SomeLibraryProc ActionMacro\n    If Len(")
        self.assertEqual(codes(self.report("    UI_Button s, \"UI_NotAMacro\"\n", button=external)), [])

    def test_parameters_passed_by_value_or_to_intrinsics_stay_checked(self) -> None:
        # ByVal parameters, parenthesized arguments and VBA functions cannot
        # change the caller's variable, so the literal is still checked.
        reader = ("Public Sub UI_Read(ByVal Macro As String)\nEnd Sub\n"
                  "Public Sub UI_Fix(Macro As String)\n    Macro = \"UI_Tick\"\nEnd Sub\n")
        safe = self.BUTTON.replace(
            "    If Len(", "    UI_Read ActionMacro\n    UI_Fix (ActionMacro)\n"
            "    Debug.Print Trim$(ActionMacro), UCase$(ActionMacro)\n    If Len(") + reader
        report = self.report("    UI_Button s, \"UI_NotAMacro\"\n", button=safe)
        self.assertEqual(codes(report), ["VBA-CALL-030"])

    def test_target_in_a_module_of_another_configuration_is_a_defect(self) -> None:
        files = {"b.bas": mod("M_B", self.BUTTON.replace("{default}", "UI_Tick")),
                 "a.bas": mod("M_A", "Public Sub UI_Run(ByVal s As Object)\n    UI_Button s, \"UI_DemoOnly\"\n"
                                     "    Application.Run \"UI_DemoOnly\"\nEnd Sub\n"),
                 "d.bas": mod("M_D", "Public Sub UI_DemoOnly()\nEnd Sub\n")}
        manifest = {"project_name_patterns": PATTERNS, "macro_name_functions": [], "document_modules": [],
                    "configurations": {"production": {"modules": ["a.bas", "b.bas"], "gating": True},
                                       "demo": {"modules": ["a.bas", "b.bas", "d.bas"], "gating": True}}}
        report = calls.analyze_project(files, manifest)
        self.assertEqual(codes(report, configuration="production"), ["VBA-CALL-030"] * 2)
        self.assertEqual(codes(report, configuration="demo"), [])
        self.assertEqual(report["status"], "fail")


class ConfigurationTests(unittest.TestCase):
    def test_alternative_declarations_are_not_duplicates(self) -> None:
        report = analyze({"a.bas": mod("M_A", "#If VBA7 Then\nPublic Sub UI_X(ByVal p As LongPtr)\n#Else\n"
                                              "Public Sub UI_X(ByVal p As Long)\n#End If\nEnd Sub\n"
                                              "Public Sub UI_Run()\n    UI_X 1\nEnd Sub\n")})
        self.assertEqual(report["status"], "pass")
        self.assertEqual(codes(report), [])

    def test_every_environment_is_analyzed(self) -> None:
        report = analyze({"a.bas": mod("M_A", "Public Sub UI_Run()\nEnd Sub\n")})
        self.assertEqual(report["configurations"]["fixture"]["environments"], list(calls.ENVIRONMENTS))
        self.assertEqual(sorted(calls.ENVIRONMENTS), ["legacy-x86", "vba7-x64", "vba7-x86"])

    def test_environment_specific_defect_is_configuration_dependent(self) -> None:
        report = analyze({"a.bas": mod("M_A", "#If Win64 Then\nPublic Sub UI_Only64()\nEnd Sub\n#End If\n"
                                              "Public Sub UI_Run()\n    UI_Only64\nEnd Sub\n")})
        findings = report["configurations"]["fixture"]["findings"]
        self.assertEqual(codes(report), [])
        dependent = [f for f in findings if f["code"] == "VBA-CALL-001"]
        self.assertEqual(len(dependent), 1)
        self.assertEqual(dependent[0]["severity"], "unknown")
        self.assertTrue(dependent[0]["configuration_dependent"])
        self.assertEqual(sorted(dependent[0]["environments"]), ["legacy-x86", "vba7-x86"])

    def test_indeterminate_conditional_is_an_analysis_failure(self) -> None:
        report = analyze({"a.bas": mod("M_A", "#If FEATURE Then\nPublic Sub UI_Run()\n    UI_Gone\nEnd Sub\n"
                                              "#End If\n")})
        self.assertEqual(report["status"], "incomplete")
        self.assertIn("VBA-CALL-090", codes(report, "failure"))
        self.assertEqual(codes(report), [])
        defined = analyze({"a.bas": mod("M_A", "#Const FEATURE = True\n#If FEATURE Then\nPublic Sub UI_Run()\n"
                                               "End Sub\n#End If\n")})
        self.assertEqual(defined["status"], "pass")

    def test_parser_failure_never_passes(self) -> None:
        report = analyze({"a.bas": mod("M_A", "Public Sub UI_Run()\n    UI_Gone\n")})
        self.assertEqual(report["status"], "incomplete")
        self.assertIn("VBA-CALL-090", codes(report, "failure"))

    def test_incomplete_coverage_never_passes(self) -> None:
        missing_module = analyze({"a.bas": mod("M_A", "Public Sub UI_Run()\n    UI_Elsewhere\nEnd Sub\n")},
                                 listed=["a.bas", "b.bas"])
        self.assertEqual(missing_module["status"], "incomplete")
        self.assertEqual(codes(missing_module), [])
        self.assertIn("VBA-CALL-001", codes(missing_module, "unknown"))
        self.assertIn("VBA-CALL-091", codes(missing_module, "failure"))
        unlisted = analyze({"a.bas": mod("M_A", "Public Sub UI_Run()\nEnd Sub\n"),
                            "b.bas": mod("M_B", "Public Sub UI_Other()\nEnd Sub\n")}, listed=["a.bas"])
        self.assertEqual(unlisted["status"], "incomplete")
        self.assertEqual([f["code"] for f in unlisted["coverage_findings"]], ["VBA-CALL-091"])

    def test_document_modules_are_excluded_not_unassigned(self) -> None:
        files = {"a.bas": mod("M_A", "Public Sub UI_Run()\nEnd Sub\n"), "s.cls": cls("shSheet", "")}
        manifest = manifest_for(files, listed=["a.bas"])
        self.assertEqual(calls.analyze_project(files, manifest)["status"], "incomplete")
        manifest["document_modules"] = ["s.cls"]
        self.assertEqual(calls.analyze_project(files, manifest)["status"], "pass")
        self.assertEqual(calls.analyze_project(files, manifest, {"a.bas"})["status"], "incomplete")

    def test_non_gating_configuration_reports_warnings(self) -> None:
        report = analyze({"a.bas": mod("M_A", "Public Sub UI_Run()\n    UI_Gone\nEnd Sub\n")}, gating=False)
        self.assertEqual(report["status"], "pass")
        self.assertEqual(codes(report, "warning"), ["VBA-CALL-001"])


class CommandLineTests(unittest.TestCase):
    def repository(self, body: str) -> Path:
        directory = tempfile.TemporaryDirectory(prefix="vba-calls-")
        self.addCleanup(directory.cleanup)
        root = Path(directory.name)
        files = {"src/a.bas": mod("M_A", body)}
        for rel, text in files.items():
            (root / rel).parent.mkdir(parents=True, exist_ok=True)
            (root / rel).write_bytes(text.replace("\n", "\r\n").encode("cp1252"))
        manifest = manifest_for(files)
        (root / "tools").mkdir()
        (root / calls.MANIFEST).write_text(json.dumps(manifest), encoding="utf-8")
        git(root, "init", "-q")
        git(root, "add", "-A")
        return root

    def run_main(self, *argv: str) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            status = calls.main(list(argv))
        return status, out.getvalue(), err.getvalue()

    def test_passes_without_defect_and_writes_reports(self) -> None:
        root = self.repository("Public Sub UI_Run()\n    UI_Other 1\nEnd Sub\nPublic Sub UI_Other(ByVal x As Long)\n"
                               "End Sub\n")
        summary = root / "summary.md"
        summary.write_text("earlier step\n", encoding="utf-8")
        output = root / "out" / "report.json"
        status, out, _ = self.run_main("--root", str(root), "--output", str(output), "--summary", str(summary))
        self.assertEqual(status, 0)
        self.assertIn("PASS", out)
        report = json.loads(output.read_text(encoding="utf-8"))
        self.assertEqual(report["status"], "pass")
        self.assertTrue(report["advisory"])
        text = summary.read_text(encoding="utf-8")
        self.assertTrue(text.startswith("earlier step\n"), "the summary is appended, not overwritten")
        self.assertIn("## VBA call checker (advisory)", text)

    def test_fails_on_a_defect(self) -> None:
        root = self.repository("Public Sub UI_Run()\n    UI_Gone\nEnd Sub\n")
        status, out, _ = self.run_main("--root", str(root))
        self.assertEqual(status, 1)
        self.assertIn("VBA-CALL-001 src/a.bas:4 UI_Run", out)

    def test_fails_when_the_analysis_is_incomplete(self) -> None:
        root = self.repository("Public Sub UI_Run()\nEnd Sub\n")
        (root / "src" / "b.bas").write_bytes(mod("M_B", "").encode("cp1252"))
        git(root, "add", "-A")
        status, out, _ = self.run_main("--root", str(root))
        self.assertEqual(status, 1)
        self.assertIn("INCOMPLETE", out)

    def test_usage_and_input_errors_exit_2(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as raised:
            calls.main(["--no-such-option"])
        self.assertEqual(raised.exception.code, 2)
        root = self.repository("Public Sub UI_Run()\nEnd Sub\n")
        (root / calls.MANIFEST).write_text("{not json", encoding="utf-8")
        status, _, err = self.run_main("--root", str(root))
        self.assertEqual(status, 2)
        self.assertIn("ERROR", err)


class RepositoryTests(unittest.TestCase):
    """The checker on this repository's own modules and manifest."""

    # Defects the checker finds in the current source, by (configuration, code, path, line, target).
    # The demo builder's fallback OnAction target names a macro that no module declares. Remove the
    # entry once the source is fixed, so that the checker's verdict on this repository is "no defect".
    # The OnAction placeholder of the demo builder is tracked in #129; remove
    # the entry when that issue is fixed.
    KNOWN_DEFECTS = {
        ("demo", "VBA-CALL-030", "demo/M_DEMO_BUILDER.bas", 2291, "DEMO_DEMOButton_NotAssigned"),
    }

    def test_repository_has_no_defect_beyond_the_known_ones(self) -> None:
        report = calls.run_check(REPO)
        defects = {(name, f["code"], f["path"], f["line"], f["target"])
                   for name, result in report["configurations"].items()
                   for f in result["findings"] if f["severity"] == "error"}
        self.assertEqual(defects, self.KNOWN_DEFECTS)
        self.assertEqual(report["coverage_findings"], [])
        for name, result in report["configurations"].items():
            with self.subTest(configuration=name):
                self.assertTrue(result["complete"])
                self.assertEqual([f for f in result["findings"] if f["severity"] in {"failure", "warning"}], [])
                self.assertGreater(result["stats"].get("resolved-calls", 0), 0)
        self.assertEqual(set(report["configurations"]), {"production", "development", "demo"})

    def test_demo_buttons_are_checked_in_the_demo_configuration(self) -> None:
        report = calls.run_check(REPO)
        resolved = {e["target"] for e in report["configurations"]["demo"]["entry_points"]
                    if e["channel"].startswith("OnAction") and e["status"] == "resolved"}
        self.assertIn("Demo_CreateDemoSheet", resolved)
        self.assertIn("Demo_SyncCheckBoxesToUI", resolved)


if __name__ == "__main__":
    unittest.main()
