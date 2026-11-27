"""The reformat.py command line, as CI and contributors invoke it.

The rule fixtures in reformat.selftest() prove what the formatter produces.
These tests prove what the command does with that: the exit status --check
reports, that --write is idempotent, and that a module it cannot identify is
refused rather than rewritten.
"""

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from gate_fixture import REPO

SCRIPT = str(REPO / "tools" / "reformat.py")

NORMAL = (
    'Attribute VB_Name = "M_PROBE"\r\n'
    "Option Explicit\r\n"
    "Option Private Module\r\n"
    "\r\n"
    "Private Sub Probe()\r\n"
    "    Dim Count               As Long\r\n"
    "    On Error GoTo Err_Handler\r\n"
    "    Count = 1\r\n"
    "Err_Handler:\r\n"
    "End Sub\r\n"
)

# The same module before formatting: a legacy label, a misaligned declaration,
# trailing whitespace and LF line endings.
DRIFTED = (
    'Attribute VB_Name = "M_PROBE"\n'
    "Option Explicit\n"
    "Option Private Module\n"
    "\n"
    "Private Sub Probe()\n"
    "    Dim Count As Long   \n"
    "    On Error GoTo Fail\n"
    "    Count = 1\n"
    "Fail:\n"
    "End Sub\n"
)

NAMELESS = "Option Explicit\r\n\r\nPrivate Sub Probe()\r\nEnd Sub\r\n"


def reformat(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run([sys.executable, SCRIPT, *args], capture_output=True, text=True)


def legacy(*args: str) -> subprocess.CompletedProcess[str]:
    """The previous normal form, which these fixtures are written in."""
    return reformat("--style", "legacy", *args)


class ReformatCli(unittest.TestCase):

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="excel-ui-reformat-")
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)

    def module(self, name: str, text: str) -> Path:
        path = self.root / name
        path.write_bytes(text.encode("ascii"))
        return path

    def test_check_passes_a_module_in_normal_form(self) -> None:
        path = self.module("M_PROBE.bas", NORMAL)
        result = legacy("--check", str(path))
        self.assertEqual(result.returncode, 0, result.stdout)
        self.assertIn(f"ok   {path}", result.stdout)

    def test_check_fails_on_drift_without_writing(self) -> None:
        path = self.module("M_PROBE.bas", DRIFTED)
        result = legacy("--check", str(path))
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("not in house-style normal form", result.stdout)
        self.assertEqual(path.read_bytes(), DRIFTED.encode("ascii"))

    def test_check_fails_when_any_one_of_several_modules_drifts(self) -> None:
        good = self.module("M_PROBE.bas", NORMAL)
        bad = self.module("M_OTHER.bas", DRIFTED.replace("M_PROBE", "M_OTHER"))
        result = legacy("--check", str(good), str(bad))
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn(f"ok   {good}", result.stdout)
        self.assertIn(f"FAIL {bad}", result.stdout)

    def test_write_produces_normal_form_and_is_idempotent(self) -> None:
        path = self.module("M_PROBE.bas", DRIFTED)
        first = legacy("--write", str(path))
        self.assertEqual(first.returncode, 0, first.stdout)
        once = path.read_bytes()
        self.assertEqual(once, NORMAL.encode("ascii"))

        second = legacy("--write", str(path))
        self.assertEqual(second.returncode, 0, second.stdout)
        self.assertEqual(path.read_bytes(), once)
        self.assertEqual(legacy("--check", str(path)).returncode, 0)

    def test_check_refuses_a_module_without_a_name(self) -> None:
        path = self.module("M_PROBE.bas", NAMELESS)
        result = legacy("--check", str(path))
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("no Attribute VB_Name", result.stdout)

    def test_write_refuses_a_module_without_a_name_and_leaves_every_file_alone(self) -> None:
        drifted = self.module("M_PROBE.bas", DRIFTED)
        nameless = self.module("M_NAMELESS.bas", NAMELESS)
        result = legacy("--write", str(drifted), str(nameless))
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("no Attribute VB_Name", result.stdout)
        self.assertEqual(drifted.read_bytes(), DRIFTED.encode("ascii"))
        self.assertEqual(nameless.read_bytes(), NAMELESS.encode("ascii"))

    def test_check_refuses_non_ascii_input(self) -> None:
        path = self.root / "M_PROBE.bas"
        path.write_bytes(NORMAL.replace("Count = 1", 'Count = Len("caf\xe9")').encode("latin-1"))
        result = legacy("--check", str(path))
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("non-ASCII", result.stdout)

    def test_selftest_passes(self) -> None:
        result = reformat("--selftest")
        self.assertEqual(result.returncode, 0, result.stdout)
        self.assertTrue(result.stdout.startswith("ok   self-test"), result.stdout)

    def test_usage_error_without_arguments(self) -> None:
        result = reformat()
        self.assertEqual(result.returncode, 2)
        self.assertIn("usage:", result.stderr)


HOUSE_PREVIOUS = (
    'Attribute VB_Name = "M_PROBE"\r\n'
    "Option Explicit\r\n"
    "\r\n"
    "Private Sub Probe()\r\n"
    "    On Error GoTo Err_Handler\r\n"
    "Safe_Exit:\r\n"
    "    Exit Sub\r\n"
    "Err_Handler:\r\n"
    "    Resume Safe_Exit\r\n"
    "End Sub\r\n"
)


class HouseStyleCli(unittest.TestCase):
    """The same command with --style house, which is what stage B of #210 will run."""

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory(prefix="excel-ui-reformat-")
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "M_PROBE.bas"
        self.path.write_bytes(HOUSE_PREVIOUS.encode("ascii"))

    def test_check_reports_the_previous_layout_as_drift(self) -> None:
        result = reformat("--style", "house", "--check", str(self.path))
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("not in house-style normal form", result.stdout)

    def test_write_renames_the_labels_and_is_idempotent(self) -> None:
        first = reformat("--style", "house", "--write", str(self.path))
        self.assertEqual(first.returncode, 0, first.stdout)
        text = self.path.read_bytes().decode("ascii")
        for old, new in (("Err_Handler", "ErrorHandler"), ("Safe_Exit", "SafeExit")):
            self.assertNotIn(old, text)
            self.assertIn(new + ":", text)
        self.assertIn("for a person to settle:", first.stdout)
        second = reformat("--style", "house", "--write", str(self.path))
        self.assertEqual(second.returncode, 0, second.stdout)
        self.assertEqual(self.path.read_bytes().decode("ascii"), text)
        self.assertEqual(reformat("--style", "house", "--check", str(self.path)).returncode, 0)

    def test_an_unknown_style_is_a_usage_error(self) -> None:
        result = reformat("--style", "other", "--check", str(self.path))
        self.assertEqual(result.returncode, 2)
        self.assertIn("usage:", result.stderr)


if __name__ == "__main__":
    unittest.main()
