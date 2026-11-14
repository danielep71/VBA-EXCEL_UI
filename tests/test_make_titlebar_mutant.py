"""The #6 title-bar negative control that tools/make_titlebar_mutant.py writes (#66).

The control must differ from src/M_EXCEL_UI_TITLEBAR.bas in the captionless
fallback condition and nowhere else, keep the module identity and line
endings the synchronizer relies on, land only in the ignored scratch folder,
and refuse loudly when it cannot be produced exactly.
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

import make_titlebar_mutant as mutant  # noqa: E402
from gate_fixture import GIT_ENV  # noqa: E402

SOURCE = REPO / "src" / "M_EXCEL_UI_TITLEBAR.bas"


def lines(data: bytes) -> list[str]:
    return data.decode("ascii").split("\r\n")


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
        self.assertEqual(lines(self.source)[a0:a1], [
            "                If Not m_FrameStates(Slot).HasBaseline _",
            "                   Or (RestoreBits And WS_CAPTION) = 0 Then",
        ])
        self.assertEqual(lines(self.control)[b0:b1], ["                If RestoreBits = 0 Then"])

    def test_the_condition_is_the_one_in_the_show_worker(self) -> None:
        text = self.source.decode("ascii")
        worker = text.index("Private Function UI_TrySetTitleBarVisibleForHwndWorker(")
        end = text.index("End Function", worker)
        self.assertIn(mutant.CURRENT.decode("ascii"), text[worker:end])

    def test_module_identity_and_line_endings_are_kept(self) -> None:
        self.assertTrue(self.control.startswith(b'Attribute VB_Name = "M_EXCEL_UI_TITLEBAR"\r\n'))
        self.assertEqual(self.control.count(b"\n"), self.control.count(b"\r\n"))
        self.control.decode("ascii")

    def test_a_missing_condition_is_refused(self) -> None:
        changed = self.source.replace(mutant.CURRENT, b"")
        with self.assertRaisesRegex(mutant.MutantError, "found 0"):
            mutant.mutate(changed)

    def test_a_duplicated_condition_is_refused(self) -> None:
        changed = self.source.replace(mutant.CURRENT, mutant.CURRENT + mutant.CURRENT)
        with self.assertRaisesRegex(mutant.MutantError, "found 2"):
            mutant.mutate(changed)

    def test_lf_line_endings_are_refused(self) -> None:
        with self.assertRaisesRegex(mutant.MutantError, "bare LF"):
            mutant.mutate(self.source.replace(b"\r\n", b"\n"))

    def test_another_module_is_refused(self) -> None:
        changed = self.source.replace(b'"M_EXCEL_UI_TITLEBAR"', b'"M_EXCEL_UI"', 1)
        with self.assertRaisesRegex(mutant.MutantError, "VB_Name"):
            mutant.mutate(changed)


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
        path.write_bytes(path.read_bytes().replace(mutant.CURRENT, b""))
        code, _, err = self.run_tool()
        self.assertEqual(code, 1)
        self.assertIn("exactly once", err)
        self.assertFalse((self.repo / ".mutation-scratch").exists())


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
