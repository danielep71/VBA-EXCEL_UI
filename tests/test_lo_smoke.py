"""Tests for tools/lo_smoke.py.

The comparison of what LibreOffice loaded with the build is tested on its own,
without LibreOffice. The end-to-end check runs only where soffice and the
Python UNO bridge are installed; the static-checks workflow runs it in a job
of its own.
"""

from __future__ import annotations

import contextlib
import importlib.util
import io
import shutil
import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import lo_smoke  # noqa: E402

LOADED_PREFIX = "Rem Attribute VBA_ModuleType=VBAModule\nOption VBASupport 1\n"


def loaded(expected: dict[str, str]) -> dict[str, str]:
    """What LibreOffice shows for modules holding the expected code."""
    return {name: LOADED_PREFIX + code + "\n\n" for name, code in expected.items()}


class Comparison(unittest.TestCase):
    expected = lo_smoke.expected_modules(ROOT, "Development")
    sheets = [("Sheet1", "Sheet1")]

    def test_the_expected_modules_are_the_document_modules_and_the_configuration(self) -> None:
        self.assertEqual(list(self.expected), ["ThisWorkbook", "Sheet1", "M_EXCEL_UI_RUNTIME", "M_EXCEL_UI_TITLEBAR",
                                               "M_EXCEL_UI_SNAPSHOT", "M_EXCEL_UI", "M_EXCEL_UI_REGRESSION_TESTS"])
        self.assertEqual(self.expected["Sheet1"], "Option Explicit")
        self.assertNotIn("Attribute VB_Name", self.expected["M_EXCEL_UI"])
        self.assertNotIn("\r", self.expected["M_EXCEL_UI"])

    def test_a_faithful_load_passes(self) -> None:
        self.assertEqual(lo_smoke.findings(self.expected, loaded(self.expected), self.sheets, "ThisWorkbook"), [])

    def test_an_unmatched_procedure_libreoffice_comments_out_still_matches(self) -> None:
        modules = loaded(self.expected)
        modules["M_EXCEL_UI"] = modules["M_EXCEL_UI"].replace(
            "\nPublic Function", "\n" + lo_smoke.UNMATCHED_PREFIX + "Public Function", 1)
        self.assertNotEqual(modules["M_EXCEL_UI"], loaded(self.expected)["M_EXCEL_UI"])
        self.assertEqual(lo_smoke.findings(self.expected, modules, self.sheets, "ThisWorkbook"), [])

    def test_each_difference_is_reported(self) -> None:
        modules = loaded(self.expected)
        del modules["M_EXCEL_UI"]
        modules["Module1"] = LOADED_PREFIX
        modules["M_EXCEL_UI_SNAPSHOT"] += "' extra line\n"
        found = lo_smoke.findings(self.expected, modules, [("Sheet1", "")], "")
        self.assertEqual(found[:3], ["module M_EXCEL_UI is missing", "module Module1 is not in the build",
                                     "module M_EXCEL_UI_SNAPSHOT: its code differs from the source"])
        self.assertIn("sheets (name, code name)", found[3])
        self.assertIn("the workbook's code name is ''", found[4])
        self.assertEqual(len(found), 5)


class CommandLine(unittest.TestCase):
    def run_main(self, *argv: str) -> tuple[int, str]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
            try:
                code = lo_smoke.main(list(argv))
            except SystemExit as exit_:
                code = int(exit_.code or 0)
        return code, out.getvalue()

    def test_it_cannot_run_without_libreoffice(self) -> None:
        with mock.patch("shutil.which", return_value=None):
            code, out = self.run_main()
        self.assertEqual(code, 2)
        self.assertIn("SKIPPED: needs LibreOffice", out)

    def test_a_missing_workbook_or_an_unknown_mode_is_refused(self) -> None:
        self.assertEqual(self.run_main(str(ROOT / "missing.xlsm"))[0], 2)
        self.assertEqual(self.run_main("--mode", "Staging")[0], 2)

    @unittest.skipUnless(shutil.which("soffice") and importlib.util.find_spec("uno"),
                         "LibreOffice and its Python UNO bridge are not installed")
    def test_libreoffice_loads_every_module_of_the_development_build(self) -> None:
        code, out = self.run_main("--mode", "Development")
        self.assertEqual(code, 0, out)
        self.assertIn("lo-smoke: PASS: 7 modules expected, 7 loaded", out)


if __name__ == "__main__":
    unittest.main()
