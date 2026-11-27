"""The house style of #210: its vocabulary, its formatter, its rules and the switch.

The formatter and the rules prove themselves with fixtures that the gate runs
on every push (house_format_fixtures, house_rules_fixtures). These tests prove
that those fixtures are run, that the formatter carries every real module into
the style without changing what it executes, and that switching the style on
leaves only the rules that need a person.
"""

import unittest
from unittest.mock import patch

from gate_fixture import (
    REPO,
    LexicalError,
    check_repo,
    executable_signature,
    fixture,
    house_format_fixtures,
    house_rules_fixtures,
    house_style,
    layout_edit,
    reformat,
    vba_house,
    vba_house_rules,
)


def modules() -> list[tuple[str, str, str]]:
    """(path, module name, text) of every standard module of the tree."""
    out = []
    for directory in ("src", "test", "demo"):
        for path in sorted((REPO / directory).glob("*.bas")):
            out.append((f"{directory}/{path.name}", path.stem, path.read_bytes().decode("ascii")))
    return out


class Vocabulary(unittest.TestCase):

    def test_the_two_label_sets_do_not_overlap(self) -> None:
        self.assertEqual(house_style.PREVIOUS_LABELS.names() & house_style.HOUSE_LABELS.names(), frozenset())

    def test_every_previous_label_has_a_house_name(self) -> None:
        for old in house_style.PREVIOUS_LABELS:
            self.assertIn(old, house_style.LABEL_RENAMES)
            self.assertIn(house_style.LABEL_RENAMES[old], house_style.HOUSE_LABELS)

    def test_the_switch_chooses_the_vocabulary(self) -> None:
        self.assertEqual(house_style.labels(True), house_style.HOUSE_LABELS)
        self.assertEqual(house_style.labels(False), house_style.PREVIOUS_LABELS)
        with patch.object(house_style, "HOUSE_STYLE_ACTIVE", True):
            self.assertEqual(house_style.labels(), house_style.HOUSE_LABELS)
            self.assertEqual(reformat.default_style(), "house")
        with patch.object(house_style, "HOUSE_STYLE_ACTIVE", False):
            self.assertEqual(house_style.labels(), house_style.PREVIOUS_LABELS)
            self.assertEqual(reformat.default_style(), "legacy")

    def test_a_title_is_centred_in_the_78_columns_after_the_apostrophe(self) -> None:
        line = house_style.centred_title("EXIT PROCEDURE")
        self.assertEqual(line, "'" + " " * 32 + "EXIT PROCEDURE")
        self.assertEqual(len(house_style.RULE_DASH), 79)

    def test_a_title_from_an_identifier_splits_words_and_keeps_acronyms(self) -> None:
        self.assertEqual(house_style.title_from_identifier("UI_SetExcelUI"), "UI SET EXCEL UI")
        self.assertEqual(house_style.title_from_identifier("TST_CertCaptureAppProps"), "TST CERT CAPTURE APP PROPS")
        self.assertEqual(house_style.title_from_identifier("DEMO_CT_ResetSuccess"), "DEMO CT RESET SUCCESS")


class Fixtures(unittest.TestCase):

    def test_the_formatter_fixtures_hold(self) -> None:
        self.assertEqual(house_format_fixtures.house_selftest(), [])

    def test_the_rule_fixtures_hold(self) -> None:
        failures, count = house_rules_fixtures.rules_selftest()
        self.assertEqual(failures, [])
        self.assertGreaterEqual(count, 30)

    def test_the_formatter_selftest_runs_the_house_fixtures(self) -> None:
        with patch.object(house_format_fixtures, "house_selftest", lambda: ["house fixture failed"]):
            self.assertIn("house fixture failed", reformat.selftest())

    def test_a_rule_that_stops_detecting_is_noticed(self) -> None:
        with patch.object(vba_house_rules, "COMPARE_CALL_RE", __import__("re").compile(r"(?!)")):
            failures, _ = house_rules_fixtures.rules_selftest()
        self.assertTrue(any("InStr" in f for f in failures))


class Comparison(unittest.TestCase):

    BEFORE = ('Private Sub A()\nConst PROC As String = "PROC"\nOn Error GoTo Err_Handler\nx = PROC\n'
              'Err_Handler:\nEnd Sub\n')
    AFTER = ('Private Sub A()\nConst PROC_NAME As String = "PROC"\nOn Error GoTo ErrorHandler\nx = PROC_NAME\n'
             'ErrorHandler:\nEnd Sub\n')

    def test_a_declared_rename_is_the_only_difference_accepted(self) -> None:
        self.assertEqual(executable_signature(vba_house.renamed_source(self.BEFORE, "M")),
                         executable_signature(self.AFTER))
        for changed in (self.AFTER.replace('"PROC"', '"PROC_NAME"'), self.AFTER.replace("x = PROC_NAME", "x = PROC"),
                        self.AFTER.replace("GoTo ErrorHandler", "GoTo Elsewhere")):
            self.assertNotEqual(executable_signature(vba_house.renamed_source(self.BEFORE, "M")),
                                executable_signature(changed))

    def test_a_member_named_like_the_constant_is_not_renamed(self) -> None:
        before = "Private Sub A()\nConst PROC As String = \"x\"\ny = Obj.PROC\nEnd Sub\n"
        self.assertIn("y = Obj.PROC\n", vba_house.renamed_source(before, "M") + "\n")

    def test_the_rename_stops_at_the_procedure_that_declares_the_constant(self) -> None:
        text = (self.BEFORE + "Public Sub B()\nEnd Sub\n")
        self.assertIn("Const PROC_NAME", vba_house.renamed_source(text, "M"))
        with self.assertRaises(LexicalError):
            vba_house.renamed_source(text.replace("Sub B()", "Sub PROC()"), "M")

    def test_the_formatter_refuses_what_the_comparison_could_not_tell_apart(self) -> None:
        text = ('Attribute VB_Name = "M_PROBE"\r\nOption Explicit\r\n\r\n'
                'Private Sub A()\r\n    Const PROC As String = "x"\r\nEnd Sub\r\n\r\n'
                "Public Sub PROC()\r\nEnd Sub\r\n")
        with self.assertRaises(LexicalError):
            reformat.reformat_text(text, "M_PROBE", "house")


class LayoutEdit(unittest.TestCase):
    """The helper that lets a fixture name a line by its words, not by its indentation."""

    TEXT = "    Dim A         As Long\r\n        If X Then\r\n            Y = 1\r\n        End If\r\n"

    def test_a_line_is_found_in_any_indentation_and_spacing(self) -> None:
        self.assertEqual(len(layout_edit.locate(self.TEXT, "Dim A As Long\r\n")), 1)
        self.assertEqual(len(layout_edit.locate(self.TEXT, "If X Then\r\n    Y = 1\r\n")), 1)

    def test_a_replacement_takes_the_indentation_of_the_line_it_replaces(self) -> None:
        out = layout_edit.replace(self.TEXT, "    If X Then\r\n", "    If True Then\r\n")
        self.assertEqual(out, self.TEXT.replace("If X Then", "If True Then"))
        out = layout_edit.replace(self.TEXT, "Y  =  1\r\n", "Y = 1\r\nZ = 2\r\n")
        assert out is not None
        self.assertIn("            Y = 1\r\n            Z = 2\r\n", out)

    def test_text_that_is_absent_or_ambiguous_is_not_replaced(self) -> None:
        self.assertIsNone(layout_edit.replace(self.TEXT, "Nothing here\r\n", ""))
        twice = self.TEXT + self.TEXT
        self.assertIsNone(layout_edit.replace(twice, "If X Then\r\n", "", unique=True))
        self.assertIsNotNone(layout_edit.replace(twice, "If X Then\r\n", ""))


class Tree(unittest.TestCase):
    """The formatter on every module of this repository, whatever the state of the switch."""

    def test_every_module_reaches_the_style_unchanged_in_what_it_executes_and_stable(self) -> None:
        for rel, name, text in modules():
            with self.subTest(module=rel):
                result = reformat.reformat_text(text, name, "house")
                self.assertEqual(reformat.reformat_text(result, name, "house"), result)
                self.assertEqual(executable_signature(vba_house.renamed_source(text, name)),
                                 executable_signature(result))

    def test_what_remains_for_a_person_is_only_what_rewriting_cannot_settle(self) -> None:
        settled = {"header", "sections", "proc-name", "select-case"}
        for rel, name, text in modules():
            with self.subTest(module=rel):
                result = reformat.reformat_text(text, name, "house")
                kinds = {f.rule for f in vba_house_rules.findings(result, name)}
                self.assertLessEqual(kinds, settled)

    def test_the_formatter_names_what_it_leaves(self) -> None:
        _, name, text = next(m for m in modules() if m[1] == "M_EXCEL_UI_RUNTIME")
        _, todos = vba_house.format_module(text, name)
        self.assertTrue(any(message.startswith("header lacks") for _, message in todos))

    def test_a_rename_that_is_ambiguous_is_refused_before_anything_is_written(self) -> None:
        _, name, text = next(m for m in modules() if m[1] == "M_EXCEL_UI_RUNTIME")
        with self.assertRaises(LexicalError):
            reformat.reformat_text(text + "\r\nFail:\r\n", name, "house")


class SwitchedOn(unittest.TestCase):
    """Stage B in miniature: reformat the tree, switch the style on, run the gate."""

    def test_the_formatted_tree_passes_every_check_except_the_rules_a_person_settles(self) -> None:
        with fixture() as repo:
            for rel, name, text in modules():
                repo.write(rel, reformat.reformat_text(text, name, "house").encode("ascii"))
            with patch.object(house_style, "HOUSE_STYLE_ACTIVE", True):
                for label, fn in check_repo.CHECKS:
                    found = repo.run(fn)
                    if label == "house-style rules":
                        self.assertTrue(found, "the rules should still have work for a person")
                        self.assertTrue(all("house-style-rules" in f for f in found), found[:3])
                    else:
                        self.assertEqual(found, [], label)


if __name__ == "__main__":
    unittest.main()
