"""Tests for the mandatory-case manifest of release certification (tools/certification_manifest.py).

The gate-level fixtures, which break the real VBA modules one defect at a time,
are in test_gate_checks.py. These tests hold the manifest file, the parsers and
the in-memory fixtures of the module itself.
"""

from __future__ import annotations

import copy
import json
import sys
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import certification_manifest as manifests  # noqa: E402

MANIFEST = ROOT / "tools" / "certification-manifest.json"
NAMED_BY_THE_ISSUE = (
    "TST_Case_TitleBarFrameRefreshDebtRetried",
    "TST_Case_TitleBarStaleFrameEntryNotReused",
    "TST_Case_ActiveFramePairRefusesMismatch",
    "TST_Case_TitleBarSameHwndGenerationReuse",
    "TST_Case_TitleBarShowRejectsCaptionlessBaseline",
    "Test_EXCEL_UI_RunCertificationSelfTest",
    "Test_EXCEL_UI_RunOwnershipCleanupChecks",
)


def document() -> dict[str, Any]:
    loaded: dict[str, Any] = json.loads(MANIFEST.read_text(encoding="utf-8"))
    return loaded


class ManifestFileTests(unittest.TestCase):
    def setUp(self) -> None:
        self.manifest = manifests.parse_manifest(MANIFEST.read_text(encoding="utf-8"))

    def test_this_repository_manifest_is_valid(self) -> None:
        self.assertEqual(self.manifest.version, 1)
        self.assertEqual([unit.name for unit in self.manifest.units],
                         ["CertificationSelfTest", "OwnershipCleanupChecks", "RegressionPack", "SnapshotIdentity",
                          "TitleBarSdiIdentity"])
        self.assertEqual(len(self.manifest.cases), len(set(self.manifest.cases)))

    def test_isolated_units_come_first(self) -> None:
        executions = [unit.execution for unit in self.manifest.units]
        self.assertEqual(executions, ["isolated", "isolated", "live", "live", "live"])

    def test_the_cases_the_issue_names_are_required_and_mandatory(self) -> None:
        for case in NAMED_BY_THE_ISSUE:
            self.assertIn(case, self.manifest.required)
            self.assertIn(case, self.manifest.cases)

    def test_the_replaced_case_is_named_by_its_current_name(self) -> None:
        self.assertNotIn("TST_Case_TitleBarSameStyleHandleReuse", self.manifest.cases)
        notes = {item["case"]: item["note"] for item in document()["required_cases"]}
        self.assertIn("TST_Case_TitleBarSameStyleHandleReuse", notes["TST_Case_TitleBarSameHwndGenerationReuse"])

    def test_the_self_test_is_listed_exactly_once(self) -> None:
        self.assertEqual(self.manifest.cases.count("Test_EXCEL_UI_RunCertificationSelfTest"), 1)

    def test_pairs_follow_the_units_in_order(self) -> None:
        self.assertEqual(self.manifest.pairs[0], ("CertificationSelfTest", "Test_EXCEL_UI_RunCertificationSelfTest"))
        self.assertEqual(self.manifest.pairs[-1][0], "TitleBarSdiIdentity")


class ManifestValidationTests(unittest.TestCase):
    def test_an_unusable_manifest_is_rejected(self) -> None:
        def put(key: str, value: Any) -> Callable[[dict[str, Any]], None]:
            return lambda doc: doc.update({key: value})

        cases: dict[str, tuple[Callable[[dict[str, Any]], None], str]] = {
            "missing key": (lambda doc: doc.pop("units"), "missing units"),
            "unknown key": (put("extra", 1), "unknown extra"),
            "other schema": (put("schema_version", 2), "schema_version 2 is not supported"),
            "boolean schema": (put("schema_version", True), "schema_version True is not supported"),
            "text version": (put("manifest_version", "1"), "manifest_version must be a positive integer"),
            "zero version": (put("manifest_version", 0), "manifest_version must be a positive integer"),
            "no description": (put("description", " "), "description must be non-empty text"),
            "no units": (put("units", []), "units must be a non-empty list"),
            "unit not an object": (put("units", ["x"]), "units[0] must be an object"),
            "unit without cases": (lambda doc: doc["units"][0].pop("cases"), "units[0]: missing cases"),
            "unit with an unknown key": (lambda doc: doc["units"][0].update(extra=1), "units[0]: unknown extra"),
            "unit with another execution": (lambda doc: doc["units"][0].update(execution="later"),
                                            "units[0].execution must be one of"),
            "unit without runner and two cases": (lambda doc: doc["units"][0].update(cases=["Test_EXCEL_UI_A",
                                                                                           "Test_EXCEL_UI_B"]),
                                                  "a unit without a runner has exactly one case"),
            "case that is not an identifier": (lambda doc: doc["units"][2]["cases"].append("not a case"),
                                               "is not a valid name"),
            "repeated unit": (lambda doc: doc["units"].append(copy.deepcopy(doc["units"][0])),
                              "units lists CertificationSelfTest more than once"),
            "repeated case": (lambda doc: doc["units"][2]["cases"].append(doc["units"][2]["cases"][0]),
                              "cases are listed more than once"),
            "isolated after live": (lambda doc: doc["units"].append(
                {"name": "Late", "execution": "isolated", "cases": ["Test_EXCEL_UI_Late"]}),
                "isolated units must come before live units"),
            "required case not mandatory": (lambda doc: doc["required_cases"].append(
                {"case": "TST_Case_Gone", "note": "x"}), "is not a mandatory case of any unit"),
            "required case twice": (lambda doc: doc["required_cases"].append(doc["required_cases"][0]),
                                    "required_cases lists TST_Case_TitleBarFrameRefreshDebtRetried more than once"),
            "required entry without a note": (lambda doc: doc["required_cases"][0].pop("note"),
                                              "exactly case and note"),
            "focused case that is mandatory": (lambda doc: doc["focused_cases"].update(
                {"TST_Case_ShowAllBaseline": "no"}), "focused_cases lists mandatory cases"),
            "interactive runner that is a case": (lambda doc: doc["interactive_runners"].update(
                {"Test_EXCEL_UI_RunSnapshotIdentity": "no"}), "interactive_runners lists mandatory"),
            "interactive runner without a reason": (lambda doc: doc["interactive_runners"].update(
                {"Test_EXCEL_UI_RunMore": ""}), "must be non-empty text"),
        }
        for name, (change, fragment) in cases.items():
            with self.subTest(name):
                doc = document()
                change(doc)
                with self.assertRaises(manifests.ManifestError) as raised:
                    manifests.parse_manifest(json.dumps(doc))
                self.assertIn(fragment, str(raised.exception))

    def test_text_that_is_not_json_or_repeats_a_key_is_rejected(self) -> None:
        for text, fragment in (("{", "not valid JSON"), ('{"units": [], "units": []}', "duplicate key"),
                               ("[]", "not a JSON object")):
            with self.subTest(text):
                with self.assertRaises(manifests.ManifestError) as raised:
                    manifests.parse_manifest(text)
                self.assertIn(fragment, str(raised.exception))


class SourceParsingTests(unittest.TestCase):
    def test_statements_drop_comments_and_join_continuations(self) -> None:
        text = ("Public Sub A()\n"
                "    ' a comment with Case 9\n"
                "    TST_CertRunCase \"X\", _\n"
                "        True 'trailing\n"
                "    Note = \"it's\"\n"
                "End Sub\n")
        statements = manifests.logical_statements(text)
        self.assertEqual(statements, [(1, "Public Sub A()"), (3, 'TST_CertRunCase "X", True'),
                                      (5, 'Note = "it\'s"'), (6, "End Sub")])

    def test_procedures_are_found_with_their_visibility_and_kind(self) -> None:
        text = ("Private Declare PtrSafe Function Api Lib \"x\" () As Long\n"
                "Public Sub One()\nEnd Sub\n"
                "Private Function Two() As Long\n    Two = 1\nEnd Function\n"
                "Sub Three()\nEnd Sub\n")
        found = manifests.parse_procedures("m.bas", text)
        self.assertEqual([(item.name, item.kind, item.public) for item in found],
                         [("One", "Sub", True), ("Two", "Function", False), ("Three", "Sub", True)])

    def test_a_registry_entry_without_its_case_is_reported(self) -> None:
        text = ("Private Function TST_CertManifestCase()\n"
                "    Select Case Index\n"
                "        Case 1\n            UnitName = \"U\"\n"
                "        Case 2\n            UnitName = \"U\"\n            CaseId = \"C\"\n"
                "        Case Else\n            Exit Function\n"
                "    End Select\nEnd Function\n")
        procedure = manifests.parse_procedures("m.bas", text)[0]
        entries, problems = manifests.registry_entries(procedure)
        self.assertEqual(entries, [(2, "U", "C")])
        self.assertTrue(any("entry 1 must assign UnitName and CaseId" in problem for problem in problems), problems)
        self.assertTrue(any("are not 1..1 in order" in problem for problem in problems), problems)

    def test_a_dispatch_label_with_two_values_is_reported(self) -> None:
        text = ("Private Sub TST_CertDispatchCase()\n"
                "    Select Case CaseId\n"
                "        Case \"A\", \"B\"\n            Test_A\n"
                "        Case \"C\"\n            Call TST_Case_C\n"
                "        Case Else\n            Err.Raise 5\n"
                "    End Select\nEnd Sub\n")
        procedure = manifests.parse_procedures("m.bas", text)[0]
        branches, problems = manifests.dispatch_labels(procedure)
        self.assertEqual([(label, called) for _, label, called in branches], [("C", "TST_Case_C")])
        self.assertTrue(any("exactly one case" in problem for problem in problems), problems)

    def test_synthetic_units_are_not_counted_as_release_units(self) -> None:
        text = ("Public Sub Run()\n"
                "    TST_CertRunIsolatedUnit \"A\", X\n"
                "    TST_CertRunUnit \"B\"\n"
                "    TST_CertRunUnit \"SelfTest.ForcedFailure\"\n"
                "End Sub\n")
        procedure = manifests.parse_procedures("m.bas", text)[0]
        self.assertEqual(manifests.unit_calls(procedure), [("isolated", "A"), ("live", "B")])


class FixtureTests(unittest.TestCase):
    def test_the_selftest_proves_every_rule(self) -> None:
        errors, count = manifests.selftest()
        self.assertEqual(errors, [])
        self.assertGreaterEqual(count, 20)

    def test_the_consistent_fixture_reports_nothing(self) -> None:
        self.assertEqual(manifests.fixture_findings(), [])

    def test_every_defect_fixture_fails_only_with_its_rule_among_the_findings(self) -> None:
        for name, source, rule in manifests.source_cases():
            with self.subTest(name):
                found = manifests.fixture_findings(source=source)
                self.assertTrue(any(f"[{rule}]" in finding for finding in found), found)
        for name, manifest, rule in manifests.manifest_cases():
            with self.subTest(name):
                found = manifests.fixture_findings(manifest=manifest)
                self.assertTrue(any(f"[{rule}]" in finding for finding in found), found)

    def test_the_real_modules_agree_with_the_real_manifest(self) -> None:
        self.assertEqual(manifests.run(ROOT), [])

    def test_a_repository_without_a_policy_is_reported(self) -> None:
        self.assertTrue(any("does not name case_manifest" in finding for finding in manifests.run(ROOT / "docs")))


if __name__ == "__main__":
    unittest.main()
