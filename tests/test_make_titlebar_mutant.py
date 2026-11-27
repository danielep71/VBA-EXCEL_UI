"""The #6 title-bar negative controls that tools/make_titlebar_mutant.py writes (#66).

Each control must differ from src/M_EXCEL_UI_TITLEBAR.bas in the conditions
its variant names and nowhere else, keep the module identity and line
endings the synchronizer relies on, land only in the ignored scratch folder,
and refuse loudly when it cannot be produced exactly. The any-bit controls
put (style And WS_CAPTION) <> 0 back for the complete-mask test, and a small
model of the show decision records which fixture each control gets wrong.
"""

import difflib
import io
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO / "tools"))

import layout_edit  # noqa: E402
import make_titlebar_mutant as mutant  # noqa: E402
from gate_fixture import GIT_ENV  # noqa: E402

SOURCE = REPO / "src" / "M_EXCEL_UI_TITLEBAR.bas"


def lines(data: bytes) -> list[str]:
    return data.decode("ascii").split("\r\n")


def words(found: list[str]) -> list[str]:
    """Lines without their indentation: the controls hold in any layout of the module."""
    return [layout_edit.words(line) for line in found]


def changed(data: bytes, old: bytes, new: bytes) -> bytes:
    """``data`` with the lines of ``old`` replaced, wherever and however they are indented."""
    out = layout_edit.replace(data.decode("ascii"), old.decode("ascii"), new.decode("ascii"))
    assert out is not None
    return out.encode("ascii")


class MutantContent(unittest.TestCase):
    def setUp(self) -> None:
        self.source = SOURCE.read_bytes()
        self.control = mutant.mutate(self.source)

    def test_only_the_fallback_condition_differs(self) -> None:
        matcher = difflib.SequenceMatcher(a=lines(self.source), b=lines(self.control), autojunk=False)
        changes = [op for op in matcher.get_opcodes() if op[0] != "equal"]
        self.assertEqual(len(changes), 1, changes)
        tag, a0, a1, b0, b1 = changes[0]
        self.assertEqual(tag, "replace")
        self.assertEqual(words(lines(self.source)[a0:a1]), [
            "If Not m_FrameStates(Slot).HasBaseline _",
            "Or (RestoreBits And WS_CAPTION) <> WS_CAPTION Then",
        ])
        self.assertEqual(words(lines(self.control)[b0:b1]), ["If RestoreBits = 0 Then"])

    def test_the_condition_is_the_one_in_the_show_worker(self) -> None:
        text = self.source.decode("ascii")
        worker = text.index("Private Function UI_TrySetTitleBarVisibleForHwndWorker(")
        end = text.index("End Function", worker)
        self.assertEqual(len(layout_edit.locate(text[worker:end], mutant.CURRENT.decode("ascii"))), 1)

    def test_module_identity_and_line_endings_are_kept(self) -> None:
        self.assertTrue(self.control.startswith(b'Attribute VB_Name = "M_EXCEL_UI_TITLEBAR"\r\n'))
        self.assertEqual(self.control.count(b"\n"), self.control.count(b"\r\n"))
        self.control.decode("ascii")

    def test_a_missing_condition_is_refused(self) -> None:
        with self.assertRaisesRegex(mutant.MutantError, "found 0"):
            mutant.mutate(changed(self.source, mutant.CURRENT, b""))

    def test_a_duplicated_condition_is_refused(self) -> None:
        with self.assertRaisesRegex(mutant.MutantError, "found 2"):
            mutant.mutate(changed(self.source, mutant.CURRENT, mutant.CURRENT + mutant.CURRENT))

    def test_lf_line_endings_are_refused(self) -> None:
        with self.assertRaisesRegex(mutant.MutantError, "bare LF"):
            mutant.mutate(self.source.replace(b"\r\n", b"\n"))

    def test_another_module_is_refused(self) -> None:
        other = self.source.replace(b'"M_EXCEL_UI_TITLEBAR"', b'"M_EXCEL_UI"', 1)
        with self.assertRaisesRegex(mutant.MutantError, "VB_Name"):
            mutant.mutate(other)


class MutantFile(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.repo = Path(self.tmp.name)
        (self.repo / "src").mkdir()
        (self.repo / "src" / "M_EXCEL_UI_TITLEBAR.bas").write_bytes(SOURCE.read_bytes())

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def run_tool(self) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = mutant.main(["--repo", str(self.repo)])
        return code, out.getvalue(), err.getvalue()

    def test_writes_the_control_into_the_scratch_folder(self) -> None:
        code, out, _ = self.run_tool()
        self.assertEqual(code, 0)
        target = self.repo / ".mutation-scratch" / "titlebar-6" / "M_EXCEL_UI_TITLEBAR.bas"
        self.assertEqual(target.read_bytes(), mutant.mutate(SOURCE.read_bytes()))
        self.assertIn(f"line {mutant.mutated_line(SOURCE.read_bytes())}\n", out)
        self.assertEqual(sorted(p.name for p in target.parent.iterdir()), ["M_EXCEL_UI_TITLEBAR.bas"])

    def test_rerunning_replaces_the_control(self) -> None:
        self.assertEqual(self.run_tool()[0], 0)
        self.assertEqual(self.run_tool()[0], 0)

    def test_another_file_in_the_folder_is_refused(self) -> None:
        folder = self.repo / ".mutation-scratch" / "titlebar-6"
        folder.mkdir(parents=True)
        (folder / "M_EXCEL_UI.bas").write_bytes(b"x")
        code, _, err = self.run_tool()
        self.assertEqual(code, 1)
        self.assertIn("M_EXCEL_UI.bas", err)
        self.assertFalse((folder / "M_EXCEL_UI_TITLEBAR.bas").exists())

    def test_a_source_without_the_condition_writes_nothing(self) -> None:
        path = self.repo / "src" / "M_EXCEL_UI_TITLEBAR.bas"
        path.write_bytes(changed(path.read_bytes(), mutant.CURRENT, b""))
        code, _, err = self.run_tool()
        self.assertEqual(code, 1)
        self.assertIn("exactly once", err)
        self.assertFalse((self.repo / ".mutation-scratch").exists())


class VariantContent(unittest.TestCase):
    """Each variant replaces exactly the conditions it names, and no others."""

    EXPECTED = {
        "v112": [
            ["If Not m_FrameStates(Slot).HasBaseline _",
             "Or (RestoreBits And WS_CAPTION) <> WS_CAPTION Then"],
        ],
        "fallback": [["Or (RestoreBits And WS_CAPTION) <> WS_CAPTION Then"]],
        "reader": [["IsVisible = ((StyleValue And WS_CAPTION) = WS_CAPTION)"]],
        "readbacks": [
            ["If (VerifiedStyle And WS_CAPTION) <> WS_CAPTION Then"],
            ["If (VerifiedStyle And WS_CAPTION) <> WS_CAPTION Then"],
        ],
    }

    def setUp(self) -> None:
        self.source = SOURCE.read_bytes()

    def changes(self, variant: str) -> list[tuple[list[str], list[str]]]:
        control = mutant.mutate(self.source, variant)
        matcher = difflib.SequenceMatcher(a=lines(self.source), b=lines(control), autojunk=False)
        return [(lines(self.source)[a0:a1], lines(control)[b0:b1])
                for tag, a0, a1, b0, b1 in matcher.get_opcodes() if tag != "equal"]

    def test_each_variant_changes_only_its_conditions(self) -> None:
        for variant, expected in self.EXPECTED.items():
            with self.subTest(variant=variant):
                self.assertEqual([words(old) for old, _ in self.changes(variant)], expected)

    def test_the_any_bit_controls_put_the_any_bit_test_back(self) -> None:
        for variant in ("fallback", "reader", "readbacks"):
            for _old, new in self.changes(variant):
                with self.subTest(variant=variant):
                    self.assertEqual(len(new), 1)
                    self.assertRegex(new[0], r"\(\w+ And WS_CAPTION\) (= 0|<> 0)")

    def test_anybit_is_the_union_of_the_three_any_bit_variants(self) -> None:
        union = [pair for variant in ("fallback", "reader", "readbacks") for pair in self.changes(variant)]
        self.assertCountEqual(self.changes("anybit"), union)

    def test_every_variant_keeps_identity_and_line_endings(self) -> None:
        for variant in mutant.VARIANTS:
            with self.subTest(variant=variant):
                control = mutant.mutate(self.source, variant)
                self.assertTrue(control.startswith(b'Attribute VB_Name = "M_EXCEL_UI_TITLEBAR"\r\n'))
                self.assertEqual(control.count(b"\n"), control.count(b"\r\n"))
                control.decode("ascii")
                self.assertNotEqual(control, self.source)

    def test_every_replaced_condition_is_in_its_own_procedure(self) -> None:
        text = self.source.decode("ascii")
        reader = text[text.index("Public Function UI_TryGetTitleBarVisibleForHwnd("):]
        reader = reader[:reader.index("End Function")]
        worker = text[text.index("Private Function UI_TrySetTitleBarVisibleForHwndWorker("):]
        worker = worker[:worker.index("End Function")]
        self.assertEqual(len(layout_edit.locate(reader, mutant.READER_CURRENT.decode("ascii"))), 1)
        for current in (mutant.FALLBACK_CURRENT, mutant.NOOP_CURRENT, mutant.WRITE_CURRENT):
            self.assertEqual(len(layout_edit.locate(worker, current.decode("ascii"))), 1)

    def test_an_unknown_variant_is_refused(self) -> None:
        with self.assertRaisesRegex(mutant.MutantError, "unknown variant"):
            mutant.mutate(self.source, "nope")

    def test_each_variant_reports_a_line_for_every_condition(self) -> None:
        for variant, replacements in mutant.VARIANTS.items():
            with self.subTest(variant=variant):
                found = mutant.mutated_lines(self.source, variant)
                self.assertEqual(len(found), len(replacements))
                self.assertTrue(all(line > 0 for _label, line in found))

    def test_a_partly_missing_variant_is_refused(self) -> None:
        with self.assertRaisesRegex(mutant.MutantError, "post-write readback exactly once"):
            mutant.mutate(changed(self.source, mutant.WRITE_CURRENT, b""), "readbacks")


class ShowDecisionModel(unittest.TestCase):
    """A model of the show decision, to record what each control gets wrong.

    This is a Python model of the predicates in UI_TrySetTitleBarVisibleForHwndWorker
    and UI_TryGetTitleBarVisibleForHwnd, not Excel. The fixtures are the live
    owned bits the Excel cases write; the expectations are from the contract:
    a show succeeds only with a complete caption on the frame afterwards.
    """

    BORDER, DLGFRAME, CAPTION, OWNED = 0x800000, 0x400000, 0xC00000, 0xCF0000
    UNRELATED = 0x16000000
    FIXTURES = {
        "ZeroBaseline": 0,
        "CaptionlessBaseline": 0x40000,
        "BorderOnly": BORDER,
        "DlgFrameOnly": DLGFRAME,
        "VisibleNoOp": OWNED,
        "ForeignCaption": 0xC80000,
    }
    ANY_BIT = {
        "fallback": {"fallback"},
        "reader": {"reader"},
        "readbacks": {"noop", "write"},
        "anybit": {"fallback", "reader", "noop", "write"},
    }

    def has_caption(self, style: int, any_bit: bool) -> bool:
        masked = style & self.CAPTION
        return masked != 0 if any_bit else masked == self.CAPTION

    def show(self, owned: int, any_bit: set[str]) -> tuple[bool, int, str]:
        """Return success, the live style afterwards and the diagnostic."""
        current = self.UNRELATED | owned
        restore = current & self.OWNED
        if not self.has_caption(restore, "fallback" in any_bit):
            restore = self.OWNED
        new = (current & ~self.OWNED) | restore
        if new == current:
            if not self.has_caption(new, "noop" in any_bit):
                return False, new, "title-bar show completed without a visible caption"
            return True, new, ""
        if not self.has_caption(new, "write" in any_bit):
            return False, new, "title-bar show did not produce a visible caption"
        return True, new, ""

    def wrong(self, variant: str | None) -> dict[str, str]:
        any_bit = self.ANY_BIT.get(variant or "", set())
        out: dict[str, str] = {}
        for name, owned in self.FIXTURES.items():
            success, live, message = self.show(owned, any_bit)
            if not success:
                out[name] = f"reported failure: {message}"
            elif not self.has_caption(live, False):
                out[name] = "false success over a partial caption"
            elif (live & ~self.OWNED) != self.UNRELATED:
                out[name] = "changed an unrelated bit"
        return out

    def test_the_complete_mask_is_right_for_every_fixture(self) -> None:
        self.assertEqual(self.wrong(None), {})

    def test_foreign_and_visible_frames_are_left_as_found(self) -> None:
        for name in ("VisibleNoOp", "ForeignCaption"):
            owned = self.FIXTURES[name]
            self.assertEqual(self.show(owned, set())[1] & self.OWNED, owned)

    def test_a_partial_mask_fails_the_any_bit_fallback_loudly(self) -> None:
        wrong = self.wrong("fallback")
        self.assertEqual(sorted(wrong), ["BorderOnly", "DlgFrameOnly"])
        for message in wrong.values():
            self.assertEqual(message, "reported failure: title-bar show completed without a visible caption")

    def test_the_full_any_bit_state_reports_success_over_a_partial_caption(self) -> None:
        wrong = self.wrong("anybit")
        self.assertEqual(wrong, {"BorderOnly": "false success over a partial caption",
                                 "DlgFrameOnly": "false success over a partial caption"})

    def test_the_readbacks_alone_are_defence_in_depth(self) -> None:
        # With the fallback condition right, a show merges the full frame, so
        # neither readback can see a partial caption: no fixture separates them.
        self.assertEqual(self.wrong("readbacks"), {})

    def test_the_any_bit_reader_reports_a_partial_mask_as_visible(self) -> None:
        for owned in (self.BORDER, self.DLGFRAME):
            self.assertTrue(self.has_caption(self.UNRELATED | owned, True))
            self.assertFalse(self.has_caption(self.UNRELATED | owned, False))

    def test_the_two_bit_harness_predicate_agrees_with_the_mask_test(self) -> None:
        # The harness decides visibility from WS_BORDER and WS_DLGFRAME
        # separately; it must agree with (style And WS_CAPTION) = WS_CAPTION.
        for style in range(0, 0x1000000, 0x10000):
            two_bit = bool(style & self.BORDER) and bool(style & self.DLGFRAME)
            self.assertEqual(two_bit, self.has_caption(style, False), hex(style))


class ScratchFolder(unittest.TestCase):
    def test_the_control_path_is_ignored_and_not_archived(self) -> None:
        path = (mutant.SCRATCH / mutant.MODULE_FILE).as_posix()
        ignored = subprocess.run(["git", "check-ignore", "-q", path], cwd=REPO, env=GIT_ENV, check=False)
        self.assertEqual(ignored.returncode, 0, f"{path} is not ignored by Git")
        attr = subprocess.run(["git", "check-attr", "export-ignore", mutant.SCRATCH.parent.as_posix()],
                              cwd=REPO, env=GIT_ENV, check=True, capture_output=True, text=True)
        self.assertTrue(attr.stdout.strip().endswith("export-ignore: set"), attr.stdout)


if __name__ == "__main__":
    unittest.main()
