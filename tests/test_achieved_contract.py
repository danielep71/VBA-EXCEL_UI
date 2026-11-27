"""The live verdict of the achieved-state unit and the evidence validator accept the same observations.

The VBA unit decides, observation by observation, which outcomes pass
(TST_AchIsAllowed, from two constants of test/M_EXCEL_UI_TEST_ACHIEVED.bas), and
tools/check_excel_evidence.py decides again from the manifest when it reads the
bundle. A run that passes live and is then rejected by the validator, or the
reverse, is a defect of the contract (#28). This test generates every outcome
combination of every scenario the unit can make, for every element and scope of
the manifest, and asserts that the two decisions are the same one.
"""

from __future__ import annotations

import dataclasses
import itertools
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import certification_manifest as manifests  # noqa: E402
import check_excel_evidence as evidence  # noqa: E402

MODULE = ROOT / "test" / "M_EXCEL_UI_TEST_ACHIEVED.bas"
MANIFEST = ROOT / "tools" / "certification-manifest.json"

WRITE, NOOP, UNVERIFIED, FAILED = "verified-write", "verified-no-op", "not-verifiable", "state-mismatch"
OUTCOMES = (WRITE, NOOP, UNVERIFIED, FAILED)
STEPS = ("hide", "hide", "show", "show")


def vba_constant(name: str) -> str:
    source = MODULE.read_bytes().decode("ascii")
    match = re.search(rf'Private Const {name}\s+As String = "([^"]*)"', source)
    assert match, f"{name} is not declared in {MODULE.name}"
    return match[1]


def row(element: str, scope: str, operation: str, outcome: str) -> evidence.ObservationRow:
    reason = "" if outcome in evidence.VERIFIED_OUTCOMES else "the reason"
    return evidence.ObservationRow(element, scope, operation, "requested", "observed", outcome, reason)


class Contract(unittest.TestCase):
    def setUp(self) -> None:
        self.manifest = manifests.parse_manifest(MANIFEST.read_text(encoding="utf-8"))
        self.verified_operations = vba_constant("TEST_ACH_VERIFIED_OPERATIONS")
        self.unverifiable_allowed = vba_constant("TEST_ACH_UNVERIFIABLE_ALLOWED")

    def live_allows(self, row_: evidence.ObservationRow) -> bool:
        """What TST_AchIsAllowed answers, with the constants read from the VBA."""
        if row_.outcome in evidence.VERIFIED_OUTCOMES:
            return f"|{row_.operation}|" in self.verified_operations
        if row_.outcome == UNVERIFIED:
            return f"|{row_.element}.{row_.operation}|" in self.unverifiable_allowed
        return True

    def live_passes(self, rows: list[evidence.ObservationRow]) -> bool:
        """The live verdict: no failed outcome, and no outcome the contract does not allow where it stands."""
        return all(self.live_allows(item) and item.outcome not in evidence.FAILED_OUTCOMES for item in rows)

    def accepted(self, observation: manifests.Observation, rows: list[evidence.ObservationRow]) -> bool:
        section = manifests.AchievedState(self.manifest.achieved.unit if self.manifest.achieved else "", (observation,))
        manifest = dataclasses.replace(self.manifest, achieved=section)
        document = evidence.AchievedDocument("OK", "", len(rows), tuple(rows), (), ())
        try:
            evidence.achieved_requirements(manifest, document, "document")
        except evidence.EvidenceError:
            return False
        return True

    @staticmethod
    def feasible(steps: tuple[str, ...]) -> bool:
        """The four steps of a scenario that all read back: the entry state decides which one changed it."""
        return steps in ((WRITE, NOOP, WRITE, NOOP), (NOOP, NOOP, WRITE, NOOP))

    def scenarios(self, observation: manifests.Observation) -> list[list[evidence.ObservationRow]]:
        """Every row list the unit can produce for an observation, with every outcome combination."""
        element, scope = observation.element, observation.scope
        shapes: list[list[evidence.ObservationRow]] = []
        # Begin failed: one entry observation. Only the Ribbon can end it as not-verifiable; any other
        # surface that cannot begin is readback-failed, which is a failed outcome.
        for outcome in ((UNVERIFIED, FAILED) if element == "Ribbon" else (FAILED,)):
            shapes.append([row(element, scope, "entry", outcome)])
        # Began: four steps and the restoration, with or without the entry row that names unreadable windows.
        # The Ribbon is a single surface, so it has no such row.
        prefixes: list[list[evidence.ObservationRow]] = [[]]
        if element != "Ribbon":
            prefixes += [[row(element, scope, "entry", UNVERIFIED)]]
        for prefix in prefixes:
            for combination in itertools.product(OUTCOMES, repeat=5):
                steps, restore = combination[:4], combination[4]
                if all(item in evidence.VERIFIED_OUTCOMES for item in steps) and not self.feasible(steps):
                    continue
                made = [row(element, scope, op, outcome) for op, outcome in zip(STEPS, steps, strict=True)]
                shapes.append([*prefix, *made, row(element, scope, "restore", restore)])
        return shapes

    def test_the_vba_constants_are_the_manifests_exemptions(self) -> None:
        section = self.manifest.achieved
        assert section is not None
        exempt = sorted(item.element for item in section.observations
                        if item.may_be_unverifiable or item.may_skip_unreadable_windows)
        listed = sorted(set(re.findall(r"\|([A-Za-z]+)\.entry", self.unverifiable_allowed)))
        self.assertEqual(listed, sorted(set(exempt)))
        self.assertEqual(sorted(set(self.verified_operations.strip("|").split("|"))), ["hide", "restore", "show"])
        self.assertEqual(len(re.findall(r"\|[A-Za-z]+\.entry", self.unverifiable_allowed)), len(listed))

    def test_the_vba_function_uses_those_constants(self) -> None:
        source = MODULE.read_bytes().decode("ascii")
        start = source.index("Private Function TST_AchIsAllowed(")
        body = source[start:source.index("End Function", start)]
        self.assertIn("TEST_ACH_VERIFIED_OPERATIONS", body)
        self.assertIn("TEST_ACH_UNVERIFIABLE_ALLOWED", body)
        record = source[source.index("Private Sub TST_AchRecordObservation("):]
        self.assertIn("TST_AchIsAllowed(m_AchElement, Operation, Outcome)", record)

    def test_live_pass_and_evidence_accept_are_the_same_decision(self) -> None:
        section = self.manifest.achieved
        assert section is not None
        checked = 0
        for observation in section.observations:
            for rows in self.scenarios(observation):
                live, offline = self.live_passes(rows), self.accepted(observation, rows)
                checked += 1
                self.assertEqual(
                    live, offline,
                    f"{observation.element}.{observation.scope}: the live verdict "
                    f"{'passes' if live else 'fails'} and the validator {'accepts' if offline else 'rejects'} "
                    + ", ".join(f"{item.operation}={item.outcome}" for item in rows))
        self.assertGreater(checked, 10000)

    def test_a_scenario_with_every_outcome_verified_passes_both(self) -> None:
        section = self.manifest.achieved
        assert section is not None
        for observation in section.observations:
            rows = [row(observation.element, observation.scope, op, outcome)
                    for op, outcome in zip((*STEPS, "restore"), (WRITE, NOOP, WRITE, NOOP, NOOP), strict=True)]
            self.assertTrue(self.live_passes(rows) and self.accepted(observation, rows), observation)

    def test_only_the_ribbon_passes_as_unverifiable_for_all_of_its_observations(self) -> None:
        section = self.manifest.achieved
        assert section is not None
        for observation in section.observations:
            rows = [row(observation.element, observation.scope, "entry", UNVERIFIED)]
            expected = observation.element == "Ribbon"
            self.assertEqual(self.live_passes(rows) and self.accepted(observation, rows), expected, observation)


if __name__ == "__main__":
    unittest.main()
