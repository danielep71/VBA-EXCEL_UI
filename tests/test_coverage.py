"""Every entry in CHECKS must have a passing and a failing unittest fixture.

Without this, a check added to the gate tomorrow would have no evidence that
it ever rejects anything, and nothing would say so.
"""

import subprocess
import sys
import unittest
from collections.abc import Iterable
from pathlib import Path
from unittest.mock import patch

from gate_fixture import GateTestCase, check_repo, fixture

TESTS_DIR = Path(__file__).resolve().parent


def gate_test_cases() -> list[type[GateTestCase]]:
    """Import every test module and return each GateTestCase subclass in it."""
    unittest.defaultTestLoader.discover(str(TESTS_DIR), pattern="test_*.py")
    found: list[type[GateTestCase]] = []
    pending = list(GateTestCase.__subclasses__())
    while pending:
        case = pending.pop()
        found.append(case)
        pending.extend(case.__subclasses__())
    return found


def methods_by_check(cases: Iterable[type[GateTestCase]]) -> dict[str, set[str]]:
    covered: dict[str, set[str]] = {}
    for case in cases:
        names = {name for name in dir(case) if name.startswith("test_")}
        covered.setdefault(case.check, set()).update(names)
    return covered


def coverage_gaps(labels: Iterable[str], covered: dict[str, set[str]]) -> list[str]:
    """Return a description of every check lacking a passing or a failing test."""
    gaps = []
    for label in labels:
        names = covered.get(label, set())
        if not any(name.startswith("test_passes") for name in names):
            gaps.append(f"{label}: no test_passes_* fixture")
        if not any(name.startswith("test_fails") for name in names):
            gaps.append(f"{label}: no test_fails_* fixture")
    return gaps


def labels() -> list[str]:
    return [label for label, _ in check_repo.CHECKS]


class ChecksCoverage(unittest.TestCase):

    def test_every_check_has_a_passing_and_a_failing_fixture(self) -> None:
        gaps = coverage_gaps(labels(), methods_by_check(gate_test_cases()))
        self.assertEqual(gaps, [], "CHECKS entries without unittest coverage")

    def test_every_gate_test_names_an_existing_check(self) -> None:
        unknown = sorted(case.__name__ for case in gate_test_cases() if case.check not in labels())
        self.assertEqual(unknown, [], "GateTestCase classes naming no CHECKS entry")

    def test_an_untested_new_check_is_reported(self) -> None:
        extended = [*check_repo.CHECKS, ("brand new check", lambda: None)]
        with patch.object(check_repo, "CHECKS", extended):
            gaps = coverage_gaps(labels(), methods_by_check(gate_test_cases()))
        self.assertEqual(gaps, ["brand new check: no test_passes_* fixture",
                                "brand new check: no test_fails_* fixture"])

    def test_a_check_tested_in_one_direction_only_is_reported(self) -> None:
        covered = {"only passes": {"test_passes_clean"}, "only fails": {"test_fails_dirty"}}
        self.assertEqual(coverage_gaps(["only passes", "only fails"], covered),
                         ["only passes: no test_fails_* fixture",
                          "only fails: no test_passes_* fixture"])


class GateEndToEnd(unittest.TestCase):
    """The whole gate, run as CI runs it, against a copy of this repository."""

    def test_passes_on_the_current_tree_and_fails_on_a_defect(self) -> None:
        with fixture() as repo:
            script = str(repo.path("tools/check_repo.py"))
            clean = subprocess.run([sys.executable, script], capture_output=True, text=True)
            self.assertEqual(clean.returncode, 0, clean.stdout + clean.stderr)
            self.assertTrue(clean.stdout.rstrip().endswith("RESULT: PASS"), clean.stdout)

            repo.remove("LICENSE")
            broken = subprocess.run([sys.executable, script], capture_output=True, text=True)
            self.assertEqual(broken.returncode, 1, broken.stdout)
            self.assertIn("required-files: missing LICENSE", broken.stdout)
            self.assertTrue(broken.stdout.rstrip().endswith("RESULT: FAIL"), broken.stdout)


if __name__ == "__main__":
    unittest.main()
