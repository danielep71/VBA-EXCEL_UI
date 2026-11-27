"""The #203 VBA syntax rules, one function at a time.

The gate check itself is proved against the real tree in test_gate_checks.
This file proves the pieces: the statement readers, the conservative
expression reader, the merge of findings across compilations, and the fixture
table that the gate runs before it reads a module.
"""

import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import vba_analyze  # noqa: E402
import vba_syntax  # noqa: E402
import vba_syntax_fixtures  # noqa: E402
from vba_lex import LexicalError  # noqa: E402

ALL = tuple(vba_analyze.CONFIGS)


def module(body: str) -> str:
    return f"Option Explicit\n{body}"


class Statements(unittest.TestCase):

    def test_short_keeps_the_head_and_the_tail_of_a_long_statement(self) -> None:
        text = "Call Foo(" + "a, " * 80 + "last)"
        shown = vba_syntax.short(text)
        self.assertTrue(shown.startswith("Call Foo(a, a"), shown)
        self.assertTrue(shown.endswith("last)"), shown)
        self.assertIn(" ... ", shown)
        self.assertEqual(vba_syntax.short("a  =\n  1"), "a = 1")

    def test_paren_message_reads_both_directions(self) -> None:
        self.assertIsNone(vba_syntax.paren_message("Foo(Bar(1), (2))"))
        self.assertEqual(vba_syntax.paren_message("Foo(1))"), "a closing parenthesis has no opening parenthesis")
        self.assertEqual(vba_syntax.paren_message(")("), "a closing parenthesis has no opening parenthesis")
        self.assertEqual(vba_syntax.paren_message("Foo((1)"), "1 opening parenthesis never closed")
        self.assertEqual(vba_syntax.paren_message("Foo(("), "2 opening parentheses never closed")

    def test_continuation_counts_are_keyed_by_the_first_physical_line(self) -> None:
        text = "A = 1\nB = 2 + _\n    3 + _\n    4\n' note _\nC = 5\n"
        counts, errors = vba_syntax.continuation_counts(text)
        self.assertEqual(counts, {1: 0, 2: 2, 5: 0, 6: 0})
        self.assertEqual(errors, [])

    def test_a_string_with_an_underscore_is_not_a_continuation(self) -> None:
        counts, _ = vba_syntax.continuation_counts('Text = "a _"\nText = "b" & _\n"c"\n')
        self.assertEqual(counts, {1: 0, 2: 1})

    def test_an_unterminated_string_is_returned_as_an_error(self) -> None:
        _, errors = vba_syntax.continuation_counts('A = 1\nText = "open\n')
        self.assertEqual(errors, [(2, 'Text = "open')])

    def test_leaves_take_the_body_of_an_inline_if(self) -> None:
        self.assertEqual(vba_syntax.leaves("If a Then b = 1"), ["b = 1"])
        self.assertEqual(vba_syntax.leaves("If a Then b = 1 Else b = 2"), ["b = 1", "b = 2"])
        self.assertEqual(vba_syntax.leaves("If a Then"), [])
        self.assertEqual(vba_syntax.leaves("ElseIf a Then"), [])
        self.assertEqual(vba_syntax.leaves("x = 1"), ["x = 1"])

    def test_classify_names_the_block_event(self) -> None:
        cases = {
            "If a Then": ("open", "if"), "If a Then b = 1": ("inline", "if"), "ElseIf a Then": ("mid", "if"),
            "Else": ("mid", "if"), "Else:": ("mid", "if"), "End If": ("close", "if"),
            "Select Case x": ("open", "select"), "Case 1": ("mid", "select"), "End Select": ("close", "select"),
            "With a": ("open", "with"), "End With": ("close", "with"), "For i = 1 To 2": ("open", "for"),
            "Next i": ("close", "for"), "Do While a": ("open", "do"), "Loop Until a": ("close", "do"),
            "While a": ("open", "while"), "Wend": ("close", "while"),
        }
        for code, event in cases.items():
            with self.subTest(code=code):
                self.assertEqual(vba_syntax.classify(code), event)
        for code in ("x = 1", "Done = 1", "Exit For", "Withdraw a"):
            with self.subTest(code=code):
                self.assertIsNone(vba_syntax.classify(code))

    def test_classify_refuses_what_the_text_does_not_settle(self) -> None:
        for code in ("If a", "Else x = 1", "ElseIf a Then b = 1"):
            with self.subTest(code=code), self.assertRaises(LexicalError):
                vba_syntax.classify(code)

    def test_next_count_counts_the_loops_one_next_closes(self) -> None:
        self.assertEqual(vba_syntax.next_count("Next"), 1)
        self.assertEqual(vba_syntax.next_count("Next i"), 1)
        self.assertEqual(vba_syntax.next_count("Next j, i"), 2)

    def test_a_statement_the_text_does_not_settle_silences_the_procedure(self) -> None:
        source = module("Sub A()\nIf a Then\nElse b = 2\nEnd Sub\n")
        found, _ = vba_syntax.findings({"a.bas": source})
        self.assertEqual([f for f in found if f.rule == "blocks"], [])


class Declarations(unittest.TestCase):

    def test_declarators_read_type_array_constant_and_visibility(self) -> None:
        found = vba_syntax.declarations("Public Const A As Long = 1, B = 2")
        self.assertEqual([(n, v.type, v.const, p) for n, v, p in found],
                         [("a", "long", True, True), ("b", None, True, True)])
        found = vba_syntax.declarations("Dim x, y As Long, z(1 To 2) As Byte, s As String * 5, n&")
        self.assertEqual([(n, v.type, v.array) for n, v, _ in found],
                         [("x", "variant", False), ("y", "long", False), ("z", "byte", True),
                          ("s", None, False), ("n", None, False)])

    def test_a_procedure_header_is_not_a_declaration(self) -> None:
        for code in ("Private Sub Foo()", "Public Declare Function F Lib \"x\" () As Long",
                     "Private Type T", "Public Enum E", "x = 1"):
            with self.subTest(code=code):
                self.assertEqual(vba_syntax.declarations(code), [])

    def test_parameters_default_to_byref_and_variant(self) -> None:
        params = vba_syntax.parse_parameters("ByVal a As Long, b, Optional ByRef c As LongPtr = 0, "
                                             "d() As String, ParamArray e() As Variant")
        assert params is not None
        self.assertEqual([(p.name, p.byref, p.type, p.array, p.paramarray) for p in params],
                         [("a", False, "long", False, False), ("b", True, "variant", False, False),
                          ("c", True, "longptr", False, False), ("d", True, "string", True, False),
                          ("e", True, "variant", True, True)])
        self.assertEqual(vba_syntax.parse_parameters(" "), ())
        self.assertIsNone(vba_syntax.parse_parameters("1"))


class Expressions(unittest.TestCase):

    def scope(self) -> vba_syntax.Scope:
        proc = vba_syntax.Procedure("P", "sub", 1, "Sub P()", None)
        proc.variables = {"ptr": vba_syntax.Variable("longptr", "LongPtr"),
                          "num": vba_syntax.Variable("long", "Long"),
                          "dbl": vba_syntax.Variable("double", "Double"),
                          "arr": vba_syntax.Variable("longptr", "LongPtr", array=True)}
        return vba_syntax.Scope(proc, vba_syntax.Module("m.bas"), vba_syntax.Project([]))

    def test_integer_arithmetic_over_a_longptr_names_it(self) -> None:
        scope = self.scope()
        for expression in ("Ptr And num", "(Ptr Or &H10) + 1", "Not Ptr", "num Mod 4 + Ptr", "-Ptr"):
            with self.subTest(expression=expression):
                self.assertEqual(vba_syntax.longptr_operand(expression, scope), "Ptr")

    def test_integer_arithmetic_without_a_longptr_names_nothing(self) -> None:
        self.assertIsNone(vba_syntax.longptr_operand("num + 1", self.scope()))

    def test_anything_else_leaves_the_expression_alone(self) -> None:
        scope = self.scope()
        for expression in ("CLng(Ptr)", "Ptr / 2", "Ptr & num", "Ptr = 0", "Ptr + dbl", "Ptr + 1.5", "Ptr + Foo",
                           "arr(1)", "Ptr + (num", "Ptr)", "Ptr + 1E3", "Ptr.x", "Ptr + #1/1/2000#"):
            with self.subTest(expression=expression):
                self.assertIsNone(vba_syntax.longptr_operand(expression, scope))

    def test_longptr_is_a_long_only_on_a_32_bit_host(self) -> None:
        self.assertEqual(vba_syntax.normalize("longptr", vba_analyze.CONFIGS["vba7-x64"]), "longlong")
        self.assertEqual(vba_syntax.normalize("longptr", vba_analyze.CONFIGS["vba7-x86"]), "long")
        self.assertEqual(vba_syntax.normalize("longptr", vba_analyze.CONFIGS["legacy-x86"]), "long")
        self.assertEqual(vba_syntax.normalize("integer", vba_analyze.CONFIGS["vba7-x64"]), "integer")


class Findings(unittest.TestCase):

    def test_a_finding_names_file_line_procedure_rule_and_statement(self) -> None:
        source = module("Sub A()\nDim x As Long\nx = (1\nEnd Sub\n")
        found, _ = vba_syntax.findings({"a.bas": source})
        self.assertEqual([str(f) for f in found],
                         ["a.bas:4: all compilations: A: parentheses: 1 opening parenthesis never closed: x = (1"])

    def test_findings_that_differ_by_compilation_name_only_those(self) -> None:
        source = module("#If VBA7 Then\nSub A()\nDim x As Long\nx = (1\nEnd Sub\n#End If\n")
        found, _ = vba_syntax.findings({"a.bas": source})
        self.assertEqual([f.configs for f in found], [("vba7-x64", "vba7-x86")])
        self.assertTrue(str(found[0]).startswith("a.bas:5: vba7-x64, vba7-x86: A: "), str(found[0]))

    def test_a_module_no_project_lists_is_read_as_a_project_of_its_own(self) -> None:
        callee = module("Public Sub Take(ByRef t As Long)\nEnd Sub\n")
        caller = module("Sub A()\nDim s As Integer\nTake s\nEnd Sub\n")
        sources = {"callee.bas": callee, "caller.bas": caller}
        listed, _ = vba_syntax.findings(sources, {"p": ["callee.bas", "caller.bas"]})
        self.assertEqual([f.rule for f in listed], ["byref-type"])
        unlisted, _ = vba_syntax.findings(sources, {"p": ["callee.bas"]})
        self.assertEqual(unlisted, [])

    def test_coverage_counts_what_each_compilation_read(self) -> None:
        source = module("Sub A()\nDim x As Long\nx = 1\nEnd Sub\n")
        _, coverage = vba_syntax.findings({"a.bas": source})
        for cov in coverage.values():
            self.assertEqual((len(cov.modules), len(cov.procedures)), (1, 1))
        self.assertEqual(len(coverage["vba7-x64"].assignments), 1)
        self.assertEqual(len(vba_syntax.summary(coverage)), len(ALL))

    def test_load_projects_reads_the_configurations_of_the_manifest(self) -> None:
        manifest = {"configurations": {"p": {"modules": ["a.bas", "b.bas"]}, "q": {"modules": "x"}, "r": 1}}
        self.assertEqual(vba_syntax.load_projects(manifest), {"p": ["a.bas", "b.bas"]})
        self.assertEqual(vba_syntax.load_projects({}), {})

    def test_the_tree_is_read_through_the_manifest(self) -> None:
        root = ROOT
        projects = vba_syntax.load_projects(json.loads((root / "tools" / "vba-projects.json").read_text("utf-8")))
        modules = sorted({rel for files in projects.values() for rel in files})
        found, coverage = vba_syntax.run(root, modules, projects)
        self.assertEqual([str(f) for f in found], [])
        self.assertEqual({len(cov.modules) for cov in coverage.values()}, {len(modules)})


class Fixtures(unittest.TestCase):

    def test_every_fixture_reports_what_it_names(self) -> None:
        errors, count = vba_syntax_fixtures.selftest()
        self.assertEqual(errors, [])
        self.assertEqual(count, len(vba_syntax_fixtures.CASES))

    def test_every_rule_has_a_failing_and_a_passing_fixture(self) -> None:
        failing = {case[2] for case in vba_syntax_fixtures.CASES if case[2]}
        self.assertEqual(failing, set(vba_syntax_fixtures.RULES))
        self.assertGreaterEqual(sum(1 for case in vba_syntax_fixtures.CASES if not case[2]), len(failing))

    def test_a_fixture_that_names_the_wrong_compilations_is_reported(self) -> None:
        sources = {"a.bas": vba_syntax_fixtures.BASE.replace(
            "    Owned = Total + 1\n", vba_syntax_fixtures.STYLE_DEFECT, 1)}
        check = vba_syntax_fixtures.violations
        self.assertEqual(check("c", sources, "longptr-assignment", "Owned", ("vba7-x64",)), [])
        self.assertTrue(check("c", sources, "longptr-assignment", "Owned", ALL))
        self.assertTrue(check("c", sources, "", "", ()))
        self.assertTrue(check("c", sources, "blocks", "Owned", ("vba7-x64",)))


if __name__ == "__main__":
    unittest.main()
