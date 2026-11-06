"""The module lists come from the Git index, so a new module cannot escape.

The inventory failures themselves (a module outside src/, test/ and demo/,
and a production set that disagrees with the documented one) are proved in
test_gate_checks.ModuleInventory. This file proves the other half: that a
module nobody listed by name is still read by every module check.
"""

import unittest
from collections.abc import Callable

from gate_fixture import Fixture, check_repo, fixture

EXTRA = "src/M_EXCEL_UI_EXTRA.bas"

# A minimal standard module in house-style normal form.
CLEAN = (
    'Attribute VB_Name = "M_EXCEL_UI_EXTRA"\r\n'
    "Option Explicit\r\n"
    "Option Private Module\r\n"
    "\r\n"
    "Private Sub Extra_Probe()\r\n"
    "    Dim Count               As Long\r\n"
    "    Count = 1\r\n"
    "End Sub\r\n"
)

# For each module check: the defect to plant in the extra module, and the
# finding that proves the check read it.
DEFECTS: list[tuple[str, Callable[[str], str], str]] = [
    ("module names",
     lambda text: text.replace('"M_EXCEL_UI_EXTRA"', '"M_OTHER"'),
     f"{EXTRA}: declares 'M_OTHER'"),
    ("option policy",
     lambda text: text.replace("Option Explicit\r\n", ""),
     f"{EXTRA}: missing Option Explicit"),
    ("encoding and line endings",
     lambda text: text + "\t\n",
     f"{EXTRA}: 1 bare LF"),
    ("banner rule widths",
     lambda text: text.replace("\r\nPrivate Sub", "\r\n'=====\r\nPrivate Sub"),
     f"{EXTRA}:5: rule is 6 cols"),
    ("house labels",
     lambda text: text.replace("End Sub", "Odd_Label:\r\nEnd Sub"),
     f"{EXTRA}:8: non-house label 'Odd_Label'"),
    ("VBA analyzer",
     lambda text: text.replace("End Sub", "    UI_Nowhere\r\nEnd Sub"),
     f"{EXTRA}:8: unresolved: "),
    ("house-style formatter",
     lambda text: text + "\r\n",
     f"{EXTRA}: not in house-style normal form"),
]


def run(repo: Fixture, label: str) -> list[str]:
    fn = dict(check_repo.CHECKS)[label]
    return repo.run(fn)


class DiscoverModules(unittest.TestCase):

    def test_sorts_tracked_modules_by_directory(self) -> None:
        found, misplaced = check_repo.discover_modules([
            "src/B.bas", "src/A.cls", "test/T.frm", "demo/D.BAS", "README.md", "src/notes.txt",
        ])
        self.assertEqual(found, {"src": ["src/A.cls", "src/B.bas"], "test": ["test/T.frm"],
                                 "demo": ["demo/D.BAS"]})
        self.assertEqual(misplaced, [])

    def test_reports_modules_elsewhere_or_nested(self) -> None:
        _, misplaced = check_repo.discover_modules(
            ["M_ROOT.bas", "docs/M.cls", "src/nested/M.bas", "tools/x.py"])
        self.assertEqual(misplaced, ["M_ROOT.bas", "docs/M.cls", "src/nested/M.bas"])

    def test_the_current_tree_matches_the_documented_production_set(self) -> None:
        with fixture() as repo:
            repo.run(lambda: None)
            self.assertEqual(sorted(check_repo.PRODUCTION_MODULES),
                             sorted(check_repo.EXPECTED_PRODUCTION_MODULES))
            self.assertEqual(check_repo.TEST_MODULES, ["test/M_EXCEL_UI_REGRESSION_TESTS.bas"])
            self.assertEqual(check_repo.DEMO_MODULES,
                             ["demo/M_DEMO_BUILDER.bas", "demo/M_DEMO_CLEANUP_TESTS.bas",
                              "demo/M_EXCEL_UI_DEMO.bas"])


class ExtraModuleReachesEveryModuleCheck(unittest.TestCase):

    def test_the_defect_table_covers_every_module_check(self) -> None:
        # A check that reads modules must appear in DEFECTS, or a new module
        # check could be added without proof that it sees discovered modules.
        module_checks = {"module names", "option policy", "encoding and line endings",
                         "banner rule widths", "house labels", "VBA analyzer",
                         "house-style formatter"}
        self.assertEqual({label for label, _, _ in DEFECTS}, module_checks)

    def test_a_clean_extra_module_passes_every_module_check(self) -> None:
        with fixture() as repo:
            repo.write(EXTRA, CLEAN)
            for label, _, _ in DEFECTS:
                with self.subTest(check=label):
                    self.assertEqual(run(repo, label), [])
            self.assertIn(EXTRA, check_repo.ALL_MODULES)

    def test_a_defect_in_an_extra_module_is_reported_by_every_module_check(self) -> None:
        with fixture() as repo:
            for label, plant, expected in DEFECTS:
                with self.subTest(check=label):
                    repo.write(EXTRA, plant(CLEAN))
                    found = run(repo, label)
                    self.assertTrue(any(expected in f for f in found),
                                    f"{label} did not report {expected!r}: {found}")


class ClassAndFormModules(unittest.TestCase):

    CLASS = (
        "VERSION 1.0 CLASS\r\n"
        "BEGIN\r\n"
        "  MultiUse = -1  'True\r\n"
        "END\r\n"
        'Attribute VB_Name = "C_EXCEL_UI_EXTRA"\r\n'
        "Option Explicit\r\n"
    )

    def test_a_class_module_is_discovered_and_named_after_its_header(self) -> None:
        with fixture() as repo:
            repo.write("src/C_EXCEL_UI_EXTRA.cls", self.CLASS)
            self.assertEqual(run(repo, "module names"), [])
            self.assertEqual(run(repo, "option policy"), [])
            self.assertIn("src/C_EXCEL_UI_EXTRA.cls", check_repo.PRODUCTION_MODULES)

    def test_a_class_module_with_the_wrong_name_is_reported(self) -> None:
        with fixture() as repo:
            repo.write("src/C_EXCEL_UI_EXTRA.cls", self.CLASS.replace('"C_EXCEL_UI_EXTRA"', '"C_X"'))
            found = run(repo, "module names")
            self.assertIn("module-name: src/C_EXCEL_UI_EXTRA.cls: declares 'C_X', "
                          "expected 'C_EXCEL_UI_EXTRA'", found)

    def test_the_formatter_rejects_a_layout_it_does_not_understand(self) -> None:
        with fixture() as repo:
            repo.write("demo/F_EXTRA.frm", self.CLASS)
            found = run(repo, "house-style formatter")
            self.assertIn("formatter: demo/F_EXTRA.frm: the house-style formatter supports "
                          "standard (.bas) modules only", found)


if __name__ == "__main__":
    unittest.main()
