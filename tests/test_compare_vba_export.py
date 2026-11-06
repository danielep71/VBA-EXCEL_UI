"""Tests for tools/compare_vba_export.py, without Excel.

Each test writes an "export" from the repository source at HEAD, with the
CRLF line endings the VBE uses, changes it the way a faulty round trip
would, and runs the comparison.
"""

from __future__ import annotations

import hashlib
import io
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from collections.abc import Callable
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import compare_vba_export as tool  # noqa: E402


def source(path: str) -> bytes:
    return subprocess.run(["git", "-C", str(ROOT), "show", f"HEAD:{path}"],
                          capture_output=True, check=True).stdout


MANIFEST: dict[str, Any] = json.loads(source("tools/vba-projects.json").decode("utf-8"))
PRODUCTION = len(MANIFEST["configurations"]["production"]["modules"])
SNAPSHOT = "src/M_EXCEL_UI_SNAPSHOT.bas"

CLASS_HEADER = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"


def document_module(name: str, base: str = "00020820", code: str = "") -> bytes:
    """A document module as Excel exports it: a sheet by default."""
    return (CLASS_HEADER
            + f'Attribute VB_Name = "{name}"\r\n'
            + f'Attribute VB_Base = "0{{{base}-0000-0000-C000-000000000046}}"\r\n'
            + "Attribute VB_GlobalNameSpace = False\r\n"
            + "Attribute VB_Creatable = False\r\n"
            + "Attribute VB_PredeclaredId = True\r\n"
            + "Attribute VB_Exposed = True\r\n"
            + "Attribute VB_TemplateDerived = False\r\n"
            + "Attribute VB_Customizable = True\r\n"
            + code).encode("cp1252")


class CompareExportTests(unittest.TestCase):
    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.folder = Path(temp.name)
        # The stand-in workbooks live in a folder of this test's own.
        books = tempfile.TemporaryDirectory()
        self.addCleanup(books.cleanup)
        self.books = Path(books.name)

    def clear(self) -> None:
        for path in self.folder.iterdir():
            path.unlink()

    def export(self, mode: str, documents: bool = True) -> None:
        """The modules of ``mode`` at HEAD and, by default, the document
        modules of a blank workbook, as the VBE exports them."""
        for path in MANIFEST["configurations"][mode]["modules"]:
            (self.folder / Path(path).name).write_bytes(source(path).replace(b"\n", b"\r\n"))
        if documents:
            (self.folder / "ThisWorkbook.cls").write_bytes(document_module("ThisWorkbook", "00020819"))
            (self.folder / "Sheet1.cls").write_bytes(document_module("Sheet1"))

    def run_tool(self, mode: str, *extra: str) -> tuple[int, str]:
        out, err = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(err):
            code = tool.main(["--root", str(ROOT), "--export", str(self.folder), "--mode", mode, *extra])
        return code, out.getvalue() + err.getvalue()

    def workbook(self, name: str, project: bytes = b"vba project") -> Path:
        """A stand-in for an .xlsm: a zip holding a VBA project part."""
        path = self.books / name
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types/>")
            archive.writestr("xl/vbaProject.bin", project)
        return path

    def write_manifest(self, workbook: Path) -> None:
        """export-manifest.json as Export-EXCEL_UI-VBA.ps1 writes it."""
        components = []
        for path in sorted(self.folder.iterdir()):
            if path.suffix in (".bas", ".cls", ".frm", ".frx"):
                name = path.stem if path.suffix == ".frx" else tool.parse(path.read_bytes(), path.name).name
                kind = {".bas": "standard", ".frm": "form", ".frx": "form-resource"}.get(path.suffix, "class")
                if path.suffix == ".cls" and b"VB_Base" in path.read_bytes():
                    kind = "document"
                components.append({"name": name, "type": kind, "file": path.name,
                                   "sha256": hashlib.sha256(path.read_bytes()).hexdigest().upper()})
        record = {"format": "excel-ui-vba-export-1",
                  "workbook": {"file": workbook.name, "sha256": hashlib.sha256(workbook.read_bytes()).hexdigest()},
                  "components": components}
        (self.folder / "export-manifest.json").write_text(json.dumps(record, indent=4), encoding="utf-8")

    def edit_manifest(self, change: Callable[[dict[str, Any]], object]) -> None:
        path = self.folder / "export-manifest.json"
        record = json.loads(path.read_text(encoding="utf-8"))
        change(record)
        path.write_text(json.dumps(record), encoding="utf-8")

    def edit(self, name: str, old: bytes, new: bytes) -> None:
        path = self.folder / name
        data = path.read_bytes()
        self.assertIn(old, data)
        path.write_bytes(data.replace(old, new, 1))

    def test_exact_export_passes_in_every_mode(self) -> None:
        for mode in tool.MODES:
            with self.subTest(mode=mode):
                self.clear()
                self.export(mode)
                code, out = self.run_tool(mode)
                self.assertEqual(code, 0, out)
                self.assertIn("RESULT=PASS", out)
                self.assertIn("Documents: 2 document module(s), not compared with the source: Sheet1, ThisWorkbook",
                              out)

    def test_an_explicit_commit_is_read(self) -> None:
        self.export("production")
        sha = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "HEAD"],
                             capture_output=True, text=True, check=True).stdout.strip()
        code, out = self.run_tool("production", "--commit", sha)
        self.assertEqual(code, 0, out)
        self.assertIn(f"Commit:    {sha}", out)

    def test_unknown_commit_is_an_error(self) -> None:
        self.export("production")
        code, out = self.run_tool("production", "--commit", "no-such-commit")
        self.assertEqual(code, 2, out)
        self.assertIn("unknown commit: no-such-commit", out)

    def test_trailing_blank_lines_and_line_endings_are_ignored(self) -> None:
        self.export("production")
        path = self.folder / "M_EXCEL_UI.bas"
        path.write_bytes(path.read_bytes().replace(b"\r\n", b"\n") + b"\r\n\r\n")
        self.assertEqual(self.run_tool("production")[0], 0)

    def test_changed_code_fails_with_its_line(self) -> None:
        self.export("production")
        self.edit("M_EXCEL_UI_SNAPSHOT.bas", b'"DisplayStatusBar"', b'"DisplayFormulaBar"')
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn(f"{SNAPSHOT}: code differs at line", out)
        self.assertIn("DisplayFormulaBar", out)

    def test_identifier_case_is_ignored_and_counted(self) -> None:
        # The VBE gives each name one spelling across the project.
        self.export("production")
        self.edit("M_EXCEL_UI_SNAPSHOT.bas", b"Option Private Module", b"option private module")
        self.edit("M_EXCEL_UI.bas", b"Option Explicit", b"OPTION EXPLICIT")
        code, out = self.run_tool("production")
        self.assertEqual(code, 0, out)
        self.assertIn("Re-cased:  2 line(s)", out)

    def test_case_change_in_a_string_or_comment_fails(self) -> None:
        for old, new in ((b'"DisplayStatusBar"', b'"DisplayStatusbar"'), (b"' PURPOSE", b"' Purpose")):
            with self.subTest(old=old):
                self.clear()
                self.export("production")
                self.edit("M_EXCEL_UI_SNAPSHOT.bas", old, new)
                code, out = self.run_tool("production")
                self.assertEqual(code, 1, out)
                self.assertIn(f"{SNAPSHOT}: code differs at line", out)

    def test_fold_case_keeps_strings_and_comments(self) -> None:
        self.assertEqual(tool.fold_case("    MsgBox \"It's OK\", , MSG_TITLE 'Show It"),
                         "    msgbox \"It's OK\", , msg_title 'Show It")
        self.assertEqual(tool.fold_case("    Rem Keep THIS"), "    rem Keep THIS")
        self.assertEqual(tool.fold_case("    X = 1: Rem Keep THIS"), "    x = 1: rem Keep THIS")
        self.assertEqual(tool.fold_case("Done: REM it's KEPT"), "done: rem it's KEPT")
        self.assertEqual(tool.fold_case("10 Rem Keep THIS"), "10 rem Keep THIS")
        self.assertEqual(tool.fold_case("    Remaining = Rem2"), "    remaining = rem2")
        self.assertEqual(tool.fold_case('    S = """A"" B"'), '    s = """A"" B"')

    def test_changed_attribute_of_a_module_fails(self) -> None:
        self.export("production")
        self.edit("M_EXCEL_UI.bas", b'Attribute VB_Name = "M_EXCEL_UI"\r\n',
                  b'Attribute VB_Name = "M_EXCEL_UI"\r\nAttribute VB_Description = "x"\r\n')
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn("src/M_EXCEL_UI.bas: attributes differ", out)

    def test_missing_and_extra_modules_fail(self) -> None:
        self.export("development")
        (self.folder / "M_EXCEL_UI_TITLEBAR.bas").unlink()
        code, out = self.run_tool("production")
        self.assertEqual(code, 1)
        self.assertIn("missing from the export: M_EXCEL_UI_TITLEBAR (src/M_EXCEL_UI_TITLEBAR.bas)", out)
        self.assertIn("not in the source for this mode: M_EXCEL_UI_REGRESSION_TESTS", out)

    def test_demo_modules_are_not_development_modules(self) -> None:
        self.export("demo")
        code, out = self.run_tool("development")
        self.assertEqual(code, 1)
        self.assertIn("missing from the export: M_EXCEL_UI_REGRESSION_TESTS", out)
        self.assertIn("not in the source for this mode: M_DEMO_BUILDER", out)
        self.assertIn("not in the source for this mode: M_EXCEL_UI_DEMO", out)

    def test_document_modules_may_hold_option_explicit_only(self) -> None:
        self.export("production", documents=False)
        (self.folder / "ThisWorkbook.cls").write_bytes(
            document_module("ThisWorkbook", "00020819", "Option Explicit\r\n\r\n"))
        (self.folder / "Chart1.cls").write_bytes(document_module("Chart1", "00020821", "\r\n"))
        code, out = self.run_tool("production")
        self.assertEqual(code, 0, out)
        self.assertIn("Documents: 2 document module(s)", out)

    def test_document_module_with_code_fails(self) -> None:
        self.export("production", documents=False)
        (self.folder / "Sheet1.cls").write_bytes(document_module(
            "Sheet1", code="Option Explicit\r\nPrivate Sub Worksheet_Activate()\r\nEnd Sub\r\n"))
        code, out = self.run_tool("production")
        self.assertEqual(code, 1, out)
        self.assertIn("document module Sheet1 (Sheet1.cls) holds code that is not in the source, first at line 2: "
                      "'Private Sub Worksheet_Activate()'", out)

    def test_class_module_is_not_a_document_module(self) -> None:
        self.export("production")
        (self.folder / "C_Extra.cls").write_bytes(
            (CLASS_HEADER + 'Attribute VB_Name = "C_Extra"\r\nOption Explicit\r\n').encode("cp1252"))
        code, out = self.run_tool("production")
        self.assertEqual(code, 1, out)
        self.assertIn("not in the source for this mode: C_Extra (C_Extra.cls)", out)

    def test_a_document_module_recorded_by_the_manifest_needs_no_vb_base(self) -> None:
        # Some exports drop VB_Base from a sheet module; the exporter's
        # record of the component type still identifies it.
        self.export("production", documents=False)
        (self.folder / "Sheet1.cls").write_bytes(
            (CLASS_HEADER + 'Attribute VB_Name = "Sheet1"\r\nOption Explicit\r\n').encode("cp1252"))
        self.assertEqual(self.run_tool("production")[0], 1)
        workbook = self.workbook("A.xlsm")
        self.write_manifest(workbook)
        self.edit_manifest(lambda r: [item.update(type="document") for item in r["components"]
                                      if item["file"] == "Sheet1.cls"])
        # The edit leaves every digest as it was, so the binding still holds.
        code, out = self.run_tool("production", "--workbook", str(workbook))
        self.assertEqual(code, 0, out)
        self.assertIn("binding=VERIFIED", out)

    def test_component_without_vb_name_is_an_error(self) -> None:
        self.export("production")
        (self.folder / "Stray.bas").write_bytes(b"Option Explicit\r\n")
        code, out = self.run_tool("production")
        self.assertEqual(code, 2)
        self.assertIn("Stray.bas: expected one Attribute VB_Name line, found 0", out)

    def test_bound_export_is_verified(self) -> None:
        workbook = self.workbook("A.xlsm")
        self.export("production")
        self.write_manifest(workbook)
        code, out = self.run_tool("production", "--workbook", str(workbook))
        self.assertEqual(code, 0, out)
        self.assertIn("Binding:   VERIFIED", out)
        self.assertIn(f"RESULT=PASS; modules={PRODUCTION}; differences=0; binding=VERIFIED", out)

    def test_renamed_workbook_is_verified_with_a_note(self) -> None:
        # A download is often renamed: the bytes bind it, the name only gets a note.
        self.export("production")
        self.write_manifest(self.workbook("EXCEL_UI.xlsm"))
        renamed = self.workbook("EXCEL_UI (1).xlsm")
        code, out = self.run_tool("production", "--workbook", str(renamed))
        self.assertEqual(code, 0, out)
        self.assertIn("NOTE export-manifest.json names the workbook EXCEL_UI.xlsm", out)
        self.assertIn("binding=VERIFIED", out)

    def test_export_compared_with_another_workbook_fails(self) -> None:
        self.export("production")
        self.write_manifest(self.workbook("A.xlsm", b"project A"))
        code, out = self.run_tool("production", "--workbook", str(self.workbook("B.xlsm", b"project B")))
        self.assertEqual(code, 1, out)
        self.assertIn("B.xlsm is not the workbook the export was made from", out)
        self.assertNotIn("NOTE", out)  # another workbook, not a renamed one
        self.assertIn("differences=0; binding=FAIL", out)

    def test_non_workbook_fails_even_when_the_manifest_names_it(self) -> None:
        readme = ROOT / "README.md"
        self.export("production")
        self.write_manifest(readme)
        code, out = self.run_tool("production", "--workbook", str(readme))
        self.assertEqual(code, 1, out)
        self.assertIn("README.md is not a macro-enabled workbook", out)

    def test_component_changed_after_the_export_fails_the_binding(self) -> None:
        # A file replaced by one that still matches the source: parity holds,
        # the binding does not.
        workbook = self.workbook("A.xlsm")
        self.export("production")
        self.write_manifest(workbook)
        path = self.folder / "M_EXCEL_UI.bas"
        path.write_bytes(path.read_bytes().replace(b"\r\n", b"\n"))
        code, out = self.run_tool("production", "--workbook", str(workbook))
        self.assertEqual(code, 1, out)
        self.assertIn("M_EXCEL_UI.bas changed after the export", out)
        self.assertIn("differences=0; binding=FAIL", out)

    def test_form_resources_are_bound(self) -> None:
        # A changed or unlisted .frx fails the binding like any component.
        workbook = self.workbook("A.xlsm")
        self.export("production")
        (self.folder / "UF_Picker.frx").write_bytes(b"form resources")
        self.write_manifest(workbook)
        self.assertEqual(self.run_tool("production", "--workbook", str(workbook))[0], 0)
        (self.folder / "UF_Picker.frx").write_bytes(b"other resources")
        code, out = self.run_tool("production", "--workbook", str(workbook))
        self.assertEqual(code, 1, out)
        self.assertIn("UF_Picker.frx changed after the export", out)
        self.write_manifest(workbook)
        (self.folder / "UF_Other.frx").write_bytes(b"form resources")
        code, out = self.run_tool("production", "--workbook", str(workbook))
        self.assertEqual(code, 1, out)
        self.assertIn("UF_Other.frx is in the export but not in export-manifest.json", out)

    def test_modified_manifest_fails(self) -> None:
        workbook = self.workbook("A.xlsm")
        cases: dict[str, tuple[Callable[[dict[str, Any]], object], str]] = {
            "dropped entry": (lambda r: r["components"].pop(0), "is in the export but not in export-manifest.json"),
            "extra entry": (lambda r: r["components"].append(
                {"name": "M_EXCEL_UI_REGRESSION_TESTS", "type": "standard",
                 "file": "M_EXCEL_UI_REGRESSION_TESTS.bas", "sha256": "0" * 64}),
                "lists M_EXCEL_UI_REGRESSION_TESTS.bas, which is not in the export"),
            "renamed component": (lambda r: r["components"][0].update(name="Other"), "names Other"),
            "workbook digest": (lambda r: r["workbook"].update(sha256="0" * 64),
                                "is not the workbook the export was made from"),
            "format": (lambda r: r.update(format="other"), "is not a valid export manifest"),
            "missing key": (lambda r: r.pop("workbook"), "is not a valid export manifest"),
            "missing workbook name": (lambda r: r["workbook"].pop("file"), "is not a valid export manifest"),
            "empty workbook name": (lambda r: r["workbook"].update(file=""), "workbook.file is not a file name"),
        }
        for label, (change, message) in cases.items():
            with self.subTest(label):
                self.clear()
                self.export("production")
                self.write_manifest(workbook)
                self.edit_manifest(change)
                code, out = self.run_tool("production", "--workbook", str(workbook))
                self.assertEqual(code, 1, out)
                self.assertIn(message, out)
                self.assertIn("binding=FAIL", out)

    def test_export_without_manifest_is_unverified(self) -> None:
        self.export("production")
        code, out = self.run_tool("production", "--workbook", str(self.workbook("A.xlsm")))
        self.assertEqual(code, 0, out)
        self.assertIn("no export-manifest.json: the export is not bound to a workbook", out)
        self.assertIn(f"RESULT=PASS; modules={PRODUCTION}; differences=0; binding=UNVERIFIED", out)

    def test_manifest_without_workbook_is_unverified_but_still_checked(self) -> None:
        self.export("production")
        self.write_manifest(self.workbook("A.xlsm"))
        code, out = self.run_tool("production")
        self.assertEqual(code, 0, out)
        self.assertIn("binding=UNVERIFIED", out)
        (self.folder / "M_EXCEL_UI.bas").unlink()
        code, out = self.run_tool("production")
        self.assertEqual(code, 1, out)
        self.assertIn("lists M_EXCEL_UI.bas, which is not in the export", out)

    def test_missing_workbook_or_export_is_an_error(self) -> None:
        self.export("production")
        code, out = self.run_tool("production", "--workbook", str(self.folder / "absent.xlsm"))
        self.assertEqual(code, 2, out)
        self.assertIn("workbook not found", out)
        out_buffer = io.StringIO()
        with redirect_stdout(out_buffer), redirect_stderr(out_buffer):
            code = tool.main(["--root", str(ROOT), "--export", str(self.folder / "absent"), "--mode", "demo"])
        self.assertEqual(code, 2)
        self.assertIn("export folder not found", out_buffer.getvalue())

    def test_userform_export_header_is_parsed(self) -> None:
        raw = ("VERSION 5.00\r\n"
               "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Picker \r\n"
               "   Caption         =   \"Pick\"\r\n"
               "   Begin {0BA686C6-F7D3-101A-993E-0000C0EF6F5E} Ctl \r\n"
               "   End\r\n"
               "End\r\n"
               'Attribute VB_Name = "UF_Picker"\r\n'
               "Attribute VB_GlobalNameSpace = False\r\n"
               "Option Explicit\r\n").encode("cp1252")
        component = tool.parse(raw, "UF_Picker.frm")
        self.assertEqual(component.name, "UF_Picker")
        self.assertEqual(component.attributes,
                         ('Attribute VB_Name = "UF_Picker"', "Attribute VB_GlobalNameSpace = False"))
        self.assertEqual(component.body, ("Option Explicit",))


if __name__ == "__main__":
    unittest.main()
