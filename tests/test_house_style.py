"""The house style of #210: its vocabulary, its formatter, its rules and the switch.

The formatter and the rules prove themselves with fixtures that the gate runs
on every push (house_format_fixtures, house_rules_fixtures). These tests prove
that those fixtures are run, that the formatter carries every real module into
the style without changing what it executes, and that switching the style on
leaves only the rules that need a person.
"""

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from gate_fixture import (
    REPO,
    LexicalError,
    executable_signature,
    house_equiv,
    house_equiv_fixtures,
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

    def test_a_field_or_member_named_like_the_constant_is_not_renamed(self) -> None:
        before = 'Private Sub A()\nConst PROC As String = "x"\ny = Row!PROC\nz = Obj.PROC\nF PROC:=1\nEnd Sub\n'
        after = vba_house.renamed_source(before, "M")
        for kept in ("y = Row!PROC", "z = Obj.PROC", "F PROC:=1"):
            self.assertIn(kept, after)

    def test_a_member_split_over_two_lines_is_not_renamed(self) -> None:
        before = ('Private Sub A()\nConst PROC As String = "x"\nDebug.Print Obj. _\n    PROC\n'
                  "Err.Raise 5, _\n    PROC, _\n    Msg\nEnd Sub\n")
        after = vba_house.renamed_source(before, "M")
        self.assertIn("Obj. _\n    PROC\n", after)
        self.assertIn("    PROC_NAME, _\n", after)

    def test_a_name_that_differs_only_in_letter_case_collides(self) -> None:
        text = ('Private proc_name As String\nPrivate Sub A()\nConst PROC As String = "x"\n'
                "Debug.Print proc_name\nEnd Sub\n")
        with self.assertRaisesRegex(LexicalError, "M line 1"):
            vba_house.renamed_source(text, "M")

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
    """The tree is held to the style, and what a person still has to settle is listed."""

    def found(self) -> list[tuple[str, vba_house_rules.Finding]]:
        return [(name, item) for _, name, text in modules() for item in vba_house_rules.findings(text, name)]

    def test_the_switch_is_on_and_every_module_is_in_the_normal_form(self) -> None:
        self.assertTrue(house_style.HOUSE_STYLE_ACTIVE)
        for rel, name, text in modules():
            with self.subTest(module=rel):
                self.assertEqual(reformat.reformat_text(text, name), text)

    def test_the_pending_list_holds_the_findings_of_the_tree_and_nothing_else(self) -> None:
        listed = (REPO / vba_house_rules.PENDING_FILE).read_text(encoding="ascii")
        result = vba_house_rules.compare_pending(self.found(), listed)
        self.assertEqual(result, ([], [], []))

    def test_the_list_is_what_the_tool_writes(self) -> None:
        listed = (REPO / vba_house_rules.PENDING_FILE).read_text(encoding="ascii")
        self.assertEqual(vba_house_rules.pending_text(self.found()), listed)

    def test_every_remaining_finding_is_one_a_person_settles(self) -> None:
        self.assertLessEqual({item.rule for _, item in self.found()}, {"header", "sections", "proc-name"})

    def test_every_procedure_with_a_declare_section_names_itself_in_proc_name(self) -> None:
        # What remains is a procedure that has no DECLARE section at all.
        for name, item in self.found():
            if item.rule == "proc-name":
                with self.subTest(module=name, procedure=item.where):
                    self.assertTrue(item.message.startswith("no DECLARE section"), item)

    def test_a_bare_literal_stands_only_under_the_deliberate_comment(self) -> None:
        bare = 0
        for _, name, text in modules():
            lines = text.split("\r\n")
            for n, line in enumerate(lines):
                if line.lstrip().startswith("'The value is deliberate"):
                    bare += 1
                    self.assertEqual(lines[n].strip(), house_style.DELIBERATE_VALUE_COMMENT, (name, n + 1))
                    self.assertTrue(lines[n + 1].startswith("    Const PROC_NAME As String = "), (name, n + 2))
        self.assertEqual(bare, 61)

    def test_every_select_case_has_a_case_else(self) -> None:
        self.assertEqual([item for _, item in self.found() if item.rule == "select-case"], [])


class WritePending(unittest.TestCase):
    """reformat.py --write-pending writes the list, and only ever shrinks it."""

    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        root = Path(self.directory.name)
        _, self.name, text = next(m for m in modules() if m[1] == "M_EXCEL_UI_SNAPSHOT")
        self.module = root / "M_EXCEL_UI_SNAPSHOT.bas"
        self.module.write_bytes(text.encode("ascii"))
        self.target = root / "pending.txt"

    def run_it(self) -> tuple[int, str]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            status = reformat.write_pending([self.module], self.target)
        return status, out.getvalue()

    def test_it_creates_the_list_when_there_is_none(self) -> None:
        status, _ = self.run_it()
        self.assertEqual(status, 0)
        text = self.target.read_text(encoding="ascii")
        self.assertIn("M_EXCEL_UI_SNAPSHOT: header\n", text)
        self.assertEqual(self.run_it()[0], 0)
        self.assertEqual(self.target.read_text(encoding="ascii"), text)

    def test_it_removes_a_line_whose_finding_is_settled(self) -> None:
        self.run_it()
        text = self.target.read_text(encoding="ascii")
        self.target.write_text(text + "M_EXCEL_UI_SNAPSHOT.UI_SnapshotClear: sections\n", encoding="ascii")
        status, _ = self.run_it()
        self.assertEqual(status, 0)
        self.assertEqual(self.target.read_text(encoding="ascii"), text)

    def test_it_refuses_to_add_a_line_to_an_existing_list(self) -> None:
        self.run_it()
        lines = self.target.read_text(encoding="ascii").splitlines(keepends=True)
        gone = next(n for n, line in enumerate(lines)
                    if line.startswith("M_EXCEL_UI_SNAPSHOT.") and line.endswith(": header\n"))
        before = "".join(lines[:gone] + lines[gone + 1:])
        self.target.write_text(before, encoding="ascii")
        status, out = self.run_it()
        self.assertEqual(status, 1)
        self.assertIn("refusing to add 1 line", out)
        self.assertEqual(self.target.read_text(encoding="ascii"), before)

    def test_it_removes_the_file_when_nothing_is_left(self) -> None:
        self.target.write_text("M_PROBE: header\n", encoding="ascii")
        self.module.write_bytes(house_format_fixtures.HOUSE_MODULE.encode("ascii"))
        status, _ = self.run_it()
        self.assertEqual(status, 0)
        self.assertFalse(self.target.exists())

    def test_it_refuses_a_module_without_a_name(self) -> None:
        self.module.write_bytes(b"Option Explicit\r\n")
        status, out = self.run_it()
        self.assertEqual(status, 1)
        self.assertIn("no Attribute VB_Name", out)


class Equivalence(unittest.TestCase):
    """reformat.py --equivalent proves two trees differ only as the style allows."""

    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.base = self.root / "base"
        self.new = self.root / "new"
        for tree in (self.base, self.new):
            (tree / "src").mkdir(parents=True)
        (self.base / "src" / "M_PROBE.bas").write_bytes(house_equiv_fixtures.BASE.encode("ascii"))
        (self.new / "src" / "M_PROBE.bas").write_bytes(house_equiv_fixtures.STYLED.encode("ascii"))

    def run_it(self) -> tuple[int, str]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            status = reformat.main(["--equivalent", str(self.base), str(self.new)])
        return status, out.getvalue()

    def test_the_fixtures_hold(self) -> None:
        self.assertEqual(house_equiv_fixtures.equiv_selftest(), [])

    def test_the_formatted_fixture_is_accepted(self) -> None:
        status, out = self.run_it()
        self.assertEqual(status, 0, out)
        self.assertIn("ok   src/M_PROBE.bas", out)

    def test_each_executable_change_is_refused_with_its_difference(self) -> None:
        for name, before, after, part in (
            ("operator", "Value > PROBE_LIMIT", "Value >= PROBE_LIMIT", "base has `If Value > PROBE_LIMIT Then`"),
            ("deleted statement", "        Probe_Describe = Text\r\n", "", "a statement was deleted"),
            ("literal", '"zero"', '"nil"', '"nil"'),
            ("inserted statement", "        Probe_Describe = Text\r\n",
             "        Probe_Describe = Text\r\n        mState = 1\r\n", "a statement was inserted"),
        ):
            with self.subTest(change=name):
                text = house_equiv_fixtures.STYLED.replace(before, after, 1)
                self.assertNotEqual(text, house_equiv_fixtures.STYLED)
                (self.new / "src" / "M_PROBE.bas").write_bytes(text.encode("ascii"))
                status, out = self.run_it()
                self.assertEqual(status, 1)
                self.assertIn("FAIL M_PROBE: Function Probe_Describe:", out)
                self.assertIn(part, out)

    def test_a_module_missing_from_either_tree_is_refused(self) -> None:
        (self.new / "src" / "M_EXTRA.bas").write_bytes(house_equiv_fixtures.STYLED.encode("ascii"))
        status, out = self.run_it()
        self.assertEqual(status, 1)
        self.assertIn("src/M_EXTRA.bas: the module is not in the base tree", out)
        (self.new / "src" / "M_EXTRA.bas").unlink()
        (self.new / "src" / "M_PROBE.bas").unlink()
        status, out = self.run_it()
        self.assertEqual(status, 1)
        self.assertIn("the module is missing from the new tree", out)

    def test_an_empty_tree_is_refused_rather_than_accepted(self) -> None:
        (self.new / "src" / "M_PROBE.bas").unlink()
        (self.base / "src" / "M_PROBE.bas").unlink()
        status, out = self.run_it()
        self.assertEqual(status, 1)
        self.assertIn("no .bas module", out)

    def test_the_tree_is_equivalent_to_itself(self) -> None:
        for rel, name, text in modules():
            with self.subTest(module=rel):
                self.assertEqual(house_equiv.compare_modules(text, text, name), [])

    def test_a_statement_that_is_wrapped_differently_is_not_a_difference(self) -> None:
        wrapped = house_equiv.rows_of(["X = 1 + _", "    2"], 0, 1)
        self.assertEqual(wrapped[0].tokens, house_equiv.rows_of(["X = 1 + 2"], 0, 0)[0].tokens)


if __name__ == "__main__":
    unittest.main()
