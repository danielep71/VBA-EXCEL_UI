"""Tests for the Excel certification bundle validator (tools/check_excel_evidence.py).

Every bundle here is synthetic: a small Git repository stands in for the
commit, and the files Certify-EXCEL_UI.ps1 and the workbook would write are
generated from it. No test needs Excel or Windows.
"""

from __future__ import annotations

import contextlib
import copy
import dataclasses
import hashlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_excel_evidence as evidence  # noqa: E402

GIT_ENV = dict(os.environ, GIT_CONFIG_NOSYSTEM="1", GIT_CONFIG_GLOBAL=os.devnull,
               GIT_AUTHOR_NAME="Test", GIT_AUTHOR_EMAIL="test@example.invalid",
               GIT_COMMITTER_NAME="Test", GIT_COMMITTER_EMAIL="test@example.invalid")
for _name in ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_COMMON_DIR"):
    GIT_ENV.pop(_name, None)

UNITS = ["RegressionPack", "SnapshotIdentity", "TitleBarSdiIdentity"]
CASE_MANIFEST: dict[str, Any] = {
    "schema_version": 1,
    "manifest_version": 2,
    "description": "A fixture manifest.",
    "units": [
        {"name": "RegressionPack", "execution": "live", "runner": "TST_RunRegressionPack",
         "cases": ["TST_Case_A", "TST_Case_B"]},
        {"name": "SnapshotIdentity", "execution": "live", "cases": ["Test_EXCEL_UI_RunSnapshotIdentity"]},
        {"name": "TitleBarSdiIdentity", "execution": "live", "runner": "Test_EXCEL_UI_RunTitleBarSdiIdentity",
         "cases": ["TST_Case_Sdi"]},
    ],
    "required_cases": [{"case": "TST_Case_B", "note": "named by the issue"}],
    "focused_cases": {},
    "interactive_runners": {},
    "achieved_state": {
        "unit": "TitleBarSdiIdentity",
        "description": "The observations of the fixture.",
        "observations": [
            {"element": "StatusBar", "scope": "Application"},
            {"element": "Gridlines", "scope": "AllWindows"},
            {"element": "Ribbon", "scope": "ActiveWindow", "may_be_unverifiable": True},
        ],
    },
}
OBSERVED = [(item["element"], item["scope"]) for item in CASE_MANIFEST["achieved_state"]["observations"]]
ACHIEVED_KEYS = ("achieved", "achievedDetail", "achievedChecked", "achievedState", "achievedMismatches",
                 "achievedUnverifiable")
PAIRS = [(unit["name"], case) for unit in CASE_MANIFEST["units"] for case in unit["cases"]]
INVENTORY_KEYS = ("manifestVersion", "inventory", "inventoryDetail", "expectedCases", "caseResults",
                  "inventoryFindings")
RUNNER = (
    "Attribute VB_Name = \"M_EXCEL_UI_TEST_CERTIFICATION\"\n"
    "Option Explicit\n\n"
    "Public Sub Test_EXCEL_UI_RunReleaseCertification()\n"
    + "".join(f"        TST_CertRunUnit \"{unit}\"\n" for unit in UNITS)
    + "End Sub\n\n"
    "Private Sub TST_CertRunUnit( _\n    ByVal UnitName As String)\nEnd Sub\n"
)
RECORDING_DETAIL = ("certification recording FAILED | failures=1"
                    " | first: growing the unit-outcome buffer | error 7")
UNVERIFIED = [
    {"seq": 1, "code": "ribbon.scope",
     "detail": "the Ribbon is per window and every read acts on the active window"},
    {"seq": 2, "code": "frameRegistry.scope",
     "detail": "the read-only seam cannot enumerate the registry entries of windows opened during the run"},
]
FINDING = {"seq": 1, "code": "app.DisplayStatusBar", "detail": "DisplayStatusBar changed from True to False"}
CLEANUP_KEYS = ("cleanupFindings", "cleanupUnverified", "cleanupChecked")
DEVELOPMENT = ["src/M_EXCEL_UI.bas", "test/M_EXCEL_UI_TEST_CERTIFICATION.bas"]
MANIFEST = {
    "schema_version": 1,
    "configurations": {
        "production": {"modules": ["src/M_EXCEL_UI.bas"]},
        "development": {"modules": DEVELOPMENT},
    },
}
FILES = {
    ".gitattributes": "*.bas text eol=crlf\n",
    "tools/vba-projects.json": json.dumps(MANIFEST, indent=2) + "\n",
    "tools/certification-manifest.json": json.dumps(CASE_MANIFEST, indent=2) + "\n",
    "src/M_EXCEL_UI.bas": "Attribute VB_Name = \"M_EXCEL_UI\"\nOption Explicit\n\n'café is spelled in cp1252\n",
    "test/M_EXCEL_UI_TEST_CERTIFICATION.bas": RUNNER,
}


def git(root: Path, *args: str) -> str:
    result = subprocess.run(["git", *args], cwd=root, env=GIT_ENV, capture_output=True, check=True)
    return result.stdout.decode("utf-8").strip()


def make_repository(root: Path) -> str:
    """A repository whose one commit holds FILES; returns the commit SHA."""
    git(root, "init", "-q")
    for path, text in FILES.items():
        target = root / path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(text.encode("cp1252"))
    git(root, "add", "-A")
    git(root, "commit", "-q", "-m", "fixture")
    # Let Git write the working tree, as a fresh clone would.
    for path in FILES:
        (root / path).unlink()
    git(root, "checkout", "--", ".")
    return git(root, "rev-parse", "HEAD")


def observation_rows(element: str, scope: str) -> list[dict[str, Any]]:
    """The observations of one scenario: hide, hide again, show, show again and the restoration."""
    steps = [("hide", "hidden", "verified-write"), ("hide", "hidden", "verified-no-op"),
             ("show", "visible", "verified-write"), ("show", "visible", "verified-no-op"),
             ("restore", "visible", "verified-no-op")]
    return [{"element": element, "scope": scope, "operation": operation, "requested": state,
             "observed": state, "outcome": outcome, "reason": ""} for operation, state, outcome in steps]


def refresh_achieved(certification: dict[str, Any]) -> None:
    """Derive the summary fields of the achieved state from its observations, as the workbook does."""
    rows = certification["achievedState"]
    failed = [row for row in rows if row["outcome"] in evidence.FAILED_OUTCOMES]
    unverified = [row for row in rows if row["outcome"] == evidence.UNVERIFIABLE_OUTCOME]

    def item(position: int, row: dict[str, Any]) -> dict[str, Any]:
        text = f"{row['outcome']}: requested={row['requested']} observed={row['observed']}"
        return {"seq": position, "code": f"{row['element']}.{row['scope']}.{row['operation']}",
                "detail": text + (f" | {row['reason']}" if row["reason"] else "")}

    certification["achievedState"] = [dict(row, seq=position) for position, row in enumerate(rows, 1)]
    certification["achievedChecked"] = len(rows)
    certification["achievedMismatches"] = [item(position, row) for position, row in enumerate(failed, 1)]
    certification["achievedUnverifiable"] = [item(position, row) for position, row in enumerate(unverified, 1)]
    certification["achieved"] = "FAILED" if failed else "OK"
    certification["achievedDetail"] = "; ".join(entry["code"] + ": " + entry["detail"]
                                                for entry in certification["achievedMismatches"])


def result_line(result: dict[str, Any]) -> str:
    return (f"RESULT: {'PASS' if result['passed'] else 'FAIL'} | "
            f"{'COMPLETE' if result['complete'] else 'INCOMPLETE'} | units={result['units']} "
            f"failed={result['unitsFailed']} skipped={result['skipped']} cleanup={result['cleanup']}")


def report_text(result: dict[str, Any]) -> str:
    """The text report, as TST_CertBuildTextReport composes it."""
    lines = ["VBA Excel UI - release certification",
             f"Excel {result['excelVersion']} build {result['excelBuild']}",
             result["operatingSystem"], result["timestampLocal"], "-" * 60, result_line(result)]
    if result.get("recording") == "FAILED":
        lines.append(f"  RECORDING  {result['recordingDetail']}")
    lines += [f"  {'PASS' if unit['passed'] else 'FAIL'}  {unit['name']}"
              + (f" | {unit['detail']}" if unit["detail"] else "") for unit in result["unitResults"]]
    lines += [f"  SKIP  {item}" for item in result["skipDetail"]]
    lines += [f"  CLEANUP  {item['code']} | {item['detail']}" for item in result.get("cleanupFindings", [])]
    if result["cleanupDetail"] and not result.get("cleanupFindings"):
        lines.append(f"  CLEANUP  {result['cleanupDetail']}")
    lines += [f"  UNVERIFIED  {item['code']} | {item['detail']}" for item in result.get("cleanupUnverified", [])]
    if "caseResults" in result:
        results = result["caseResults"]
        started = sum(1 for row in results if row["started"])
        resulted = sum(1 for row in results if row["status"] != "NO-RESULT")
        lines.append(f"  CASES  expected={len(result['expectedCases'])} started={started} results={resulted}"
                     f" manifest={result['manifestVersion']} inventory={result['inventory']}")
        lines += [f"  EXPECTED  {item['unit']}  {item['case']}" for item in result["expectedCases"]]
        lines += [f"  CASE  {row['status']}  {row['unit'] or '-'}  {row['case']}"
                  + (f" | {row['detail']}" if row["detail"] else "") for row in results]
        lines += [f"  INVENTORY  {item['code']} | {item['detail']}" for item in result["inventoryFindings"]]
    if all(key in result for key in ACHIEVED_KEYS):
        checked = len(result["achievedState"])
        failed, unverifiable = len(result["achievedMismatches"]), len(result["achievedUnverifiable"])
        lines.append(f"  OBSERVATIONS  checked={checked} verified={checked - failed - unverifiable} failed={failed}"
                     f" unverifiable={unverifiable} achieved={result['achieved']}")
        lines += [f"  MISMATCH  {item['code']} | {item['detail']}" for item in result["achievedMismatches"]]
        lines += [f"  UNVERIFIABLE  {item['code']} | {item['detail']}" for item in result["achievedUnverifiable"]]
    return "\r\n".join(lines) + "\r\n"


def documents(root: Path, sha: str) -> dict[str, Any]:
    """The documents of a passing bundle for ``sha``."""
    inventory = evidence.source_inventory(root, sha)
    for module in inventory["modules"]:
        module["checked_out_sha256"] = hashlib.sha256((root / module["path"]).read_bytes()).hexdigest()
    certification = {
        "component": "VBA Excel UI", "schema": 1, "timestampLocal": "2026-11-06 10:03:00",
        "excelVersion": "16.0", "excelBuild": "17928", "operatingSystem": "Windows (64-bit) NT 10.00",
        "bitness": "x64", "vbaGeneration": "VBA7", "units": len(UNITS), "unitsFailed": 0, "skipped": 0,
        "cleanup": "OK", "cleanupDetail": "", "cleanupFindings": [], "cleanupUnverified": copy.deepcopy(UNVERIFIED),
        "cleanupChecked": 14, "manifestVersion": CASE_MANIFEST["manifest_version"], "inventory": "OK",
        "inventoryDetail": "",
        "expectedCases": [{"seq": index, "unit": unit, "case": case} for index, (unit, case) in enumerate(PAIRS, 1)],
        "caseResults": [{"seq": index, "unit": unit, "case": case, "started": True, "status": "PASS", "detail": ""}
                        for index, (unit, case) in enumerate(PAIRS, 1)],
        "inventoryFindings": [], "recording": "OK", "recordingDetail": "",
        "complete": True, "passed": True,
        "unitResults": [{"name": unit, "passed": True, "detail": ""} for unit in UNITS],
        "skipDetail": [],
        "achievedState": [row for element, scope in OBSERVED for row in observation_rows(element, scope)],
    }
    refresh_achieved(certification)
    host = {
        "schema": 1,
        "windows": {"caption": "Microsoft Windows 11 Pro", "version": "10.0.22631", "build": "22631.4460",
                    "display_version": "23H2", "architecture": "64-bit"},
        "excel": {"version": "16.0", "build": "17928", "operating_system": "Windows (64-bit) NT 10.00",
                  "bitness": "x64", "bitness_source": "process", "click_to_run_version": "16.0.17928.20156",
                  "click_to_run_platform": "x64"},
        "trust_center": {"macro_setting": "Disable VBA macros with notification", "vba_project_access": "On",
                         "automation_security": "msoAutomationSecurityByUI"},
    }
    session = {
        "schema": 1, "tool": "tools/Certify-EXCEL_UI.ps1", "commit": sha, "dirty": False,
        "allow_dirty": False, "dirty_paths": [],
        "started_utc": "2026-11-06T09:00:00Z", "started_local": "2026-11-06T10:00:00+01:00",
        "finished_utc": "2026-11-06T09:05:00Z", "finished_local": "2026-11-06T10:05:00+01:00",
        "steps": [{"name": name, "status": "PASS", "detail": "", "at_utc": "2026-11-06T09:01:00Z"}
                  for name in evidence.default_policy().steps],
        "workbook": {"file": "EXCEL_UI-development.xlsm", "sha256": "0" * 64},
        "components": [{"name": "ThisWorkbook", "type": 100}, {"name": "Sheet1", "type": 100}]
        + [{"name": module["name"], "type": 1} for module in inventory["modules"]],
        "evidence": [], "result": "PASS", "exit_code": 0, "message": "",
    }
    return {"session.json": session, "host.json": host, "source-digests.json": inventory,
            "certification.json": certification}


def write_bundle(folder: Path, docs: dict[str, Any], report: str | None = None) -> None:
    """Write a bundle; the session records the digests of the evidence it holds."""
    folder.mkdir(parents=True, exist_ok=True)
    certification = docs["certification.json"]
    raw = {
        # Print # writes the workbook's text in the ANSI code page, ending in CRLF.
        "certification.json": (json.dumps(certification, separators=(",", ":")) + "\r\n").encode("cp1252"),
        "certification.txt": ((report if report is not None else report_text(certification)) + "\r\n")
        .encode("cp1252"),
    }
    docs["session.json"]["evidence"] = [
        {"file": name, "source_name": f"EXCEL_UI_certification_20261106_100300.{name.split('.')[1]}",
         "sha256": hashlib.sha256(data).hexdigest()} for name, data in raw.items()]
    for name, data in raw.items():
        (folder / name).write_bytes(data)
    for name in ("session.json", "host.json", "source-digests.json"):
        (folder / name).write_text(json.dumps(docs[name], indent=2) + "\n", encoding="utf-8")


class Fixture(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name) / "repo"
        self.root.mkdir()
        self.sha = make_repository(self.root)
        self.base = documents(self.root, self.sha)
        self.bundle = Path(self.temp.name) / "bundle"

    def bundle_with(self, change: Callable[[dict[str, Any]], None] | None = None,
                    report: str | None = None) -> Path:
        docs = copy.deepcopy(self.base)
        if change is not None:
            change(docs)
        write_bundle(self.bundle, docs, report)
        return self.bundle

    def outcomes(self, folder: Path | None = None, sha: str | None = None) -> dict[str, tuple[bool, str]]:
        found = evidence.evaluate(folder or self.bundle, sha or self.sha, self.root)
        return {name: (passed, detail) for name, passed, detail in found}

    def failed(self, folder: Path | None = None, sha: str | None = None) -> dict[str, str]:
        return {name: detail for name, (passed, detail) in self.outcomes(folder, sha).items() if not passed}

    def run_main(self, *arguments: str) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = evidence.main([*arguments, "--root", str(self.root)])
        return code, out.getvalue(), err.getvalue()


class SourceTests(Fixture):
    def test_checked_out_adds_a_carriage_return_only_where_one_is_missing(self) -> None:
        self.assertEqual(evidence.checked_out(b"a\nb\r\nc\n"), b"a\r\nb\r\nc\r\n")
        self.assertEqual(evidence.checked_out(b"no newline"), b"no newline")

    def test_inventory_digests_equal_the_files_git_checks_out(self) -> None:
        inventory = evidence.source_inventory(self.root, self.sha)
        self.assertEqual([module["path"] for module in inventory["modules"]], DEVELOPMENT)
        self.assertEqual([module["name"] for module in inventory["modules"]],
                         ["M_EXCEL_UI", "M_EXCEL_UI_TEST_CERTIFICATION"])
        for module in inventory["modules"]:
            on_disk = (self.root / module["path"]).read_bytes()
            self.assertIn(b"\r\n", on_disk)
            self.assertEqual(module["sha256"], hashlib.sha256(on_disk).hexdigest(), module["path"])

    def test_units_are_read_from_the_certification_runner(self) -> None:
        self.assertEqual(evidence.certification_units(self.root, self.sha), UNITS)

    def test_inventory_option_prints_the_digests_for_the_commit(self) -> None:
        code, out, _ = self.run_main("--inventory", "--commit", self.sha)
        self.assertEqual(code, 0)
        self.assertEqual(json.loads(out), evidence.source_inventory(self.root, self.sha))


class PassingBundleTests(Fixture):
    def test_a_complete_bundle_passes_every_check(self) -> None:
        self.bundle_with()
        outcomes = self.outcomes()
        self.assertEqual(list(outcomes), [name for name, _ in evidence.CHECKS])
        self.assertEqual(self.failed(), {})
        self.assertEqual(outcomes["verdict"][1], result_line(self.base["certification.json"]))

    def test_main_prints_a_pass_line_per_check_and_exits_zero(self) -> None:
        self.bundle_with()
        code, out, _ = self.run_main(str(self.bundle), "--commit", self.sha)
        self.assertEqual(code, 0)
        self.assertEqual(out.count("  PASS  "), len(evidence.CHECKS))
        self.assertTrue(out.rstrip().endswith("RESULT: PASS"))

    def test_the_commit_defaults_to_head(self) -> None:
        self.bundle_with()
        code, _, _ = self.run_main(str(self.bundle))
        self.assertEqual(code, 0)

    def test_a_document_without_a_recording_field_passes(self) -> None:
        def earlier_runner(docs: dict[str, Any]) -> None:
            del docs["certification.json"]["recording"]
            del docs["certification.json"]["recordingDetail"]
        self.bundle_with(earlier_runner)
        self.assertEqual(self.failed(), {})

    def test_the_achieved_state_is_held_to_the_manifest(self) -> None:
        self.assertEqual(self.outcomes(self.bundle_with())["achieved state"],
                         (True, "3 elements and scopes, 15 observations read back"))

    def test_a_ribbon_that_could_not_be_verified_is_listed_and_passes(self) -> None:
        def unverifiable(docs: dict[str, Any]) -> None:
            certification = docs["certification.json"]
            rows = [row for row in certification["achievedState"] if row["element"] != "Ribbon"]
            rows.append({"element": "Ribbon", "scope": "ActiveWindow", "operation": "entry",
                         "requested": "not attempted", "observed": "unreadable", "outcome": "not-verifiable",
                         "reason": "no active window can be identified as the owner of the Ribbon"})
            certification["achievedState"] = rows
            refresh_achieved(certification)
        self.bundle_with(unverifiable)
        self.assertEqual(self.failed(), {})
        self.assertEqual(self.outcomes()["achieved state"],
                         (True, "3 elements and scopes, 11 observations read back, 1 not verifiable"))
        report = (self.bundle / "certification.txt").read_text(encoding="cp1252")
        self.assertIn("  UNVERIFIABLE  Ribbon.ActiveWindow.entry | not-verifiable: requested=not attempted", report)

    def test_a_manifest_without_the_section_does_not_require_the_achieved_state(self) -> None:
        manifest = copy.deepcopy(CASE_MANIFEST)
        del manifest["achieved_state"]
        self.assertEqual(manifest["manifest_version"], 2)
        policy = dataclasses.replace(evidence.default_policy(), case_manifest="tools/earlier-manifest.json")
        (self.root / "tools" / "earlier-manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
        git(self.root, "add", "-A")
        git(self.root, "commit", "-q", "-m", "an earlier manifest")
        later = git(self.root, "rev-parse", "HEAD")
        self.bundle_with()
        found = {name: detail for name, passed, detail in evidence.evaluate(self.bundle, later, self.root, policy)
                 if name == "achieved state"}
        self.assertIn("has no achieved_state section", found["achieved state"])

    def test_the_manifest_is_read_from_the_commit(self) -> None:
        manifest = evidence.case_manifest(self.root, self.sha, evidence.default_policy())
        self.assertEqual(manifest.version, 2)
        self.assertEqual(list(manifest.pairs), PAIRS)
        self.assertEqual(self.outcomes(self.bundle_with())["case inventory"],
                         (True, "4 cases of 3 units, manifest version 2, each started once and passed"))

    def test_a_document_without_itemized_cleanup_passes(self) -> None:
        def earlier_runner(docs: dict[str, Any]) -> None:
            for key in CLEANUP_KEYS:
                del docs["certification.json"][key]
        self.bundle_with(earlier_runner)
        self.assertNotIn("UNVERIFIED", (self.bundle / "certification.txt").read_text(encoding="cp1252"))
        self.assertEqual(self.failed(), {})

    def test_unverified_surfaces_do_not_fail_a_clean_cleanup(self) -> None:
        self.bundle_with()
        report = (self.bundle / "certification.txt").read_text(encoding="cp1252")
        self.assertIn("  UNVERIFIED  ribbon.scope | ", report)
        self.assertEqual(self.failed(), {})

    def test_an_ansi_report_with_a_non_ascii_detail_is_read(self) -> None:
        def detail(docs: dict[str, Any]) -> None:
            docs["certification.json"]["unitResults"][0]["detail"] = "résumé"
        self.bundle_with(detail)
        self.assertEqual(self.failed(), {})


class PolicyDrivenTests(Fixture):
    """The checker takes every rule from the policy it is given."""

    RENAMES = {"session.json": "run.json", "host.json": "machine.json", "source-digests.json": "sources.json",
               "certification.json": "verdict.json", "certification.txt": "verdict.txt"}
    STEPS = ["prepare", "run", "tidy"]

    def policy_file(self, change: Callable[[dict[str, Any]], None] | None = None) -> Path:
        document = json.loads(evidence.POLICY.read_text(encoding="utf-8"))
        if change is not None:
            change(document)
        path = Path(self.temp.name) / "policy.json"
        path.write_text(json.dumps(document, indent=2), encoding="utf-8")
        return path

    def test_an_explicit_copy_of_the_policy_passes_every_check(self) -> None:
        self.bundle_with()
        code, out, _ = self.run_main(str(self.bundle), "--commit", self.sha, "--policy", str(self.policy_file()))
        self.assertEqual(code, 0, out)
        self.assertEqual(out.count("  PASS  "), len(evidence.CHECKS))
        self.assertEqual(len(evidence.CHECKS), 11)

    def test_a_bundle_follows_renamed_files_steps_schema_and_component(self) -> None:
        def contract(document: dict[str, Any]) -> None:
            document["required_files"] = {role: self.RENAMES[name] for role, name in document["required_files"].items()}
            document["evidence_files"] = [self.RENAMES[name] for name in document["evidence_files"]]
            document.update(steps=self.STEPS, evidence_schema=4, component="Another component")

        def documents_for(docs: dict[str, Any]) -> None:
            for name in ("session.json", "host.json", "source-digests.json", "certification.json"):
                docs[name]["schema"] = 4
            docs["certification.json"]["component"] = "Another component"
            docs["session.json"]["steps"] = [{"name": step, "status": "PASS", "detail": "", "at_utc": None}
                                             for step in self.STEPS]

        policy = self.policy_file(contract)
        self.bundle_with(documents_for)
        session = json.loads((self.bundle / "session.json").read_text(encoding="utf-8"))
        for item in session["evidence"]:
            item["file"] = self.RENAMES[item["file"]]
        (self.bundle / "session.json").write_text(json.dumps(session), encoding="utf-8")
        for old, new in self.RENAMES.items():
            (self.bundle / old).rename(self.bundle / new)

        code, out, _ = self.run_main(str(self.bundle), "--commit", self.sha, "--policy", str(policy))
        self.assertEqual(code, 0, out)
        self.assertEqual(out.count("  PASS  "), len(evidence.CHECKS))
        self.assertIn("run.json, machine.json, sources.json, verdict.json, verdict.txt", out)
        # The repository policy does not accept that bundle.
        self.assertIn("files", self.failed())

    def renamed_documents(self) -> Path:
        """A policy and a bundle whose session and host records are named without .json."""
        names = {"session": "session.data", "host": "HOST.JSON"}
        policy = self.policy_file(lambda document: document["required_files"].update(names))
        self.bundle_with()
        (self.bundle / "session.json").rename(self.bundle / "session.data")
        (self.bundle / "host.json").rename(self.bundle / "HOST.JSON")
        return policy

    def test_json_documents_are_read_by_role_whatever_their_names(self) -> None:
        policy = self.renamed_documents()
        code, out, _ = self.run_main(str(self.bundle), "--commit", self.sha, "--policy", str(policy))
        self.assertEqual(code, 0, out)
        self.assertEqual(out.count("  PASS  "), len(evidence.CHECKS))
        self.assertIn("session.data, HOST.JSON, source-digests.json", out)

    def test_a_role_document_that_is_not_json_fails_clearly(self) -> None:
        policy = self.renamed_documents()
        (self.bundle / "session.data").write_text("{not json", encoding="utf-8")
        found = evidence.evaluate(self.bundle, self.sha, self.root, evidence.load_policy(policy))
        failed = {name: detail for name, passed, detail in found if not passed}
        self.assertNotIn("files", failed)
        for check in ("session", "evidence digests", "workbook modules", "timing"):
            self.assertIn("cannot check: session.data is not valid JSON", failed[check], check)
        self.assertNotIn("host", failed)

    def test_a_step_the_bundle_did_not_record_fails(self) -> None:
        self.bundle_with()
        policy = self.policy_file(lambda document: document["steps"].append("archive"))
        code, out, _ = self.run_main(str(self.bundle), "--commit", self.sha, "--policy", str(policy))
        self.assertEqual(code, 1)
        self.assertIn("FAIL  session:", out)
        self.assertIn("'archive'", out)

    def test_the_configuration_and_runner_come_from_the_policy(self) -> None:
        policy = evidence.default_policy()
        production = dataclasses.replace(policy, configuration="production")
        inventory = evidence.source_inventory(self.root, self.sha, production)
        self.assertEqual(inventory["configuration"], "production")
        self.assertEqual([module["path"] for module in inventory["modules"]], ["src/M_EXCEL_UI.bas"])
        other = dataclasses.replace(policy, certification_runner="Test_Other")
        with self.assertRaisesRegex(evidence.EvidenceError, "has no Test_Other"):
            evidence.certification_units(self.root, self.sha, other)


class NegativeTests(Fixture):
    def assert_fails(self, check: str, fragment: str, folder: Path | None = None, sha: str | None = None) -> None:
        failed = self.failed(folder, sha)
        self.assertIn(check, failed)
        self.assertIn(fragment, failed[check])

    def test_a_session_for_another_commit_fails(self) -> None:
        self.bundle_with(lambda docs: docs["session.json"].update(commit="f" * 40))
        self.assert_fails("session", "not " + self.sha)

    def test_a_bundle_checked_against_another_commit_fails(self) -> None:
        self.bundle_with()
        (self.root / "src" / "M_EXCEL_UI.bas").write_bytes(b'Attribute VB_Name = "M_EXCEL_UI"\r\n\'changed\r\n')
        git(self.root, "commit", "-q", "-am", "change")
        later = git(self.root, "rev-parse", "HEAD")
        failed = self.failed(sha=later)
        self.assertIn("session", failed)
        self.assertIn("source digests", failed)
        code, out, _ = self.run_main(str(self.bundle))
        self.assertEqual(code, 1)
        self.assertTrue(out.rstrip().endswith("RESULT: FAIL"))

    def test_a_source_digest_mismatch_fails(self) -> None:
        def wrong(docs: dict[str, Any]) -> None:
            docs["source-digests.json"]["modules"][0]["sha256"] = "0" * 64
        self.bundle_with(wrong)
        self.assert_fails("source digests", "src/M_EXCEL_UI.bas digest differs")

    def test_an_imported_file_that_differs_from_the_commit_fails(self) -> None:
        def wrong(docs: dict[str, Any]) -> None:
            docs["source-digests.json"]["modules"][1]["checked_out_sha256"] = "1" * 64
        self.bundle_with(wrong)
        self.assert_fails("source digests", "checked-out test/M_EXCEL_UI_TEST_CERTIFICATION.bas differs")

    def test_a_missing_module_digest_fails(self) -> None:
        self.bundle_with(lambda docs: docs["source-digests.json"]["modules"].pop())
        self.assert_fails("source digests", "are not the commit's development modules")

    def test_a_failed_verdict_fails_even_when_the_report_agrees(self) -> None:
        def failed(docs: dict[str, Any]) -> None:
            result = docs["certification.json"]
            result.update(passed=False, unitsFailed=1)
            result["unitResults"][1].update(passed=False, detail="1004: assertion failed")
        self.bundle_with(failed)
        failures = self.failed()
        self.assertIn("the verdict is FAIL | COMPLETE | units=3 failed=1", failures["verdict"])
        self.assertNotIn("report", failures)

    def test_a_skipped_unit_fails(self) -> None:
        def skipped(docs: dict[str, Any]) -> None:
            docs["certification.json"].update(complete=False, passed=False, skipped=1,
                                               skipDetail=["SnapshotIdentity: snapshot existed"])
        self.bundle_with(skipped)
        self.assert_fails("verdict", "INCOMPLETE")

    def test_a_cleanup_failure_fails(self) -> None:
        def leaked(docs: dict[str, Any]) -> None:
            docs["certification.json"].update(cleanup="FAILED", cleanupDetail="workbook count 2, expected 1")
        self.bundle_with(leaked)
        self.assert_fails("verdict", "cleanup FAILED: workbook count 2")

    @staticmethod
    def leaked_status_bar(docs: dict[str, Any]) -> None:
        docs["certification.json"].update(
            passed=False, cleanup="FAILED", cleanupDetail=FINDING["detail"], cleanupFindings=[dict(FINDING)])

    def test_an_itemized_cleanup_failure_fails_and_names_the_finding(self) -> None:
        self.bundle_with(self.leaked_status_bar)
        failures = self.failed()
        self.assertIn("cleanup FAILED: DisplayStatusBar changed from True to False", failures["verdict"])
        self.assertNotIn("report", failures)

    def test_a_clean_cleanup_that_holds_a_finding_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(cleanupFindings=[dict(FINDING)]))
        self.assert_fails("verdict", "cleanup is OK but cleanupFindings is not empty")

    def test_a_clean_cleanup_that_compared_nothing_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(cleanupChecked=0))
        self.assert_fails("verdict", "compared nothing")

    def test_a_failed_cleanup_without_a_finding_or_detail_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(passed=False, cleanup="FAILED"))
        self.assert_fails("verdict", "cleanup FAILED without a finding or a cleanupDetail")

    def test_some_cleanup_fields_without_the_others_fail(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].pop("cleanupChecked"))
        self.assert_fails("verdict", "cleanup is itemized but has no cleanupChecked")

    def test_a_finding_that_is_not_numbered_in_order_fails(self) -> None:
        def second(docs: dict[str, Any]) -> None:
            self.leaked_status_bar(docs)
            docs["certification.json"]["cleanupFindings"][0]["seq"] = 2
        self.bundle_with(second)
        self.assert_fails("verdict", "cleanupFindings[0] must have seq 1")

    def test_a_finding_without_a_code_or_a_detail_fails(self) -> None:
        def codeless(docs: dict[str, Any]) -> None:
            self.leaked_status_bar(docs)
            docs["certification.json"]["cleanupFindings"][0]["code"] = " "
        self.bundle_with(codeless)
        self.assert_fails("verdict", "cleanupFindings[0] has no code")

        def detailless(docs: dict[str, Any]) -> None:
            self.leaked_status_bar(docs)
            docs["certification.json"]["cleanupFindings"][0]["detail"] = ""
        self.bundle_with(detailless)
        self.assert_fails("verdict", "cleanupFindings[0] has no detail")

    def test_a_report_that_omits_a_finding_fails(self) -> None:
        result = copy.deepcopy(self.base)
        self.leaked_status_bar(result)
        report = report_text(result["certification.json"]).replace(
            f"  CLEANUP  {FINDING['code']} | {FINDING['detail']}\r\n", "")
        self.bundle_with(self.leaked_status_bar, report=report)
        self.assert_fails("report", "do not begin with the findings")

    def test_a_clean_report_with_a_cleanup_line_fails(self) -> None:
        lines = report_text(self.base["certification.json"]).split("\r\n")
        lines.insert(6, "  CLEANUP  app.DisplayStatusBar | DisplayStatusBar changed from True to False")
        self.bundle_with(report="\r\n".join(lines))
        self.assert_fails("report", "CLEANUP lines")

    def test_a_report_that_omits_an_unverified_surface_fails(self) -> None:
        report = report_text(self.base["certification.json"]).replace(
            f"  UNVERIFIED  {UNVERIFIED[0]['code']} | {UNVERIFIED[0]['detail']}\r\n", "")
        self.bundle_with(report=report)
        self.assert_fails("report", "UNVERIFIED lines")

    def test_a_report_with_an_extra_unverified_surface_fails(self) -> None:
        lines = report_text(self.base["certification.json"]).split("\r\n")
        lines.insert(6, "  UNVERIFIED  ribbon.exit | not compared")
        self.bundle_with(report="\r\n".join(lines))
        self.assert_fails("report", "UNVERIFIED lines")

    @staticmethod
    def recording_failed(docs: dict[str, Any]) -> None:
        docs["certification.json"].update(complete=False, passed=False, recording="FAILED",
                                           recordingDetail=RECORDING_DETAIL)

    def test_a_recording_failure_fails_and_is_named(self) -> None:
        self.bundle_with(self.recording_failed)
        failures = self.failed()
        self.assertIn(RECORDING_DETAIL, failures["verdict"])
        self.assertNotIn("report", failures)

    def test_a_report_that_omits_the_recording_failure_fails(self) -> None:
        result = copy.deepcopy(self.base)
        self.recording_failed(result)
        report = report_text(result["certification.json"]).replace(f"  RECORDING  {RECORDING_DETAIL}\r\n", "")
        self.bundle_with(self.recording_failed, report=report)
        self.assert_fails("report", "RECORDING lines [] differ")

    def test_a_recording_line_in_a_clean_report_fails(self) -> None:
        lines = report_text(self.base["certification.json"]).split("\r\n")
        lines.insert(6, f"  RECORDING  {RECORDING_DETAIL}")
        self.bundle_with(report="\r\n".join(lines))
        self.assert_fails("report", "RECORDING lines")

    def test_a_recording_value_other_than_ok_or_failed_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(recording="PARTIAL"))
        self.assert_fails("verdict", "recording must be OK or FAILED")

    def test_a_failed_recording_without_a_detail_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(recording="FAILED"))
        self.assert_fails("verdict", "recording FAILED without a recordingDetail")

    def test_a_recording_detail_on_a_clean_recording_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(recordingDetail=RECORDING_DETAIL))
        self.assert_fails("verdict", "recording is OK but recordingDetail is not empty")

    def test_units_other_than_the_runners_fail(self) -> None:
        def renamed(docs: dict[str, Any]) -> None:
            docs["certification.json"]["unitResults"][2]["name"] = "Other"
        self.bundle_with(renamed)
        self.assert_fails("verdict", "are not the commit's")

    def test_a_missing_file_fails(self) -> None:
        self.bundle_with()
        (self.bundle / "certification.txt").unlink()
        failed = self.failed()
        self.assertIn("certification.txt", failed["files"])
        self.assertIn("certification.txt is missing", failed["report"])
        self.assertIn("certification.txt is missing", failed["evidence digests"])
        self.assertNotIn("verdict", failed)

    def test_a_missing_host_record_fails(self) -> None:
        self.bundle_with()
        (self.bundle / "host.json").unlink()
        self.assert_fails("host", "host.json is missing")

    def test_a_report_whose_result_line_disagrees_fails(self) -> None:
        report = report_text(self.base["certification.json"]).replace("skipped=0", "skipped=1")
        self.bundle_with(report=report)
        self.assert_fails("report", "certification.json says")

    def test_a_report_with_two_result_lines_fails(self) -> None:
        line = result_line(self.base["certification.json"])
        self.bundle_with(report=report_text(self.base["certification.json"]) + line + "\r\n")
        self.assert_fails("report", "expected one RESULT line, found 2")

    def test_a_report_whose_unit_lines_disagree_fails(self) -> None:
        report = report_text(self.base["certification.json"]).replace("  PASS  SnapshotIdentity",
                                                                     "  FAIL  SnapshotIdentity")
        self.bundle_with(report=report)
        self.assert_fails("report", "unit lines")

    def test_evidence_edited_after_collection_fails(self) -> None:
        self.bundle_with()
        path = self.bundle / "certification.json"
        path.write_bytes(path.read_bytes().replace(b'"excelBuild":"17928"', b'"excelBuild":"17929"'))
        self.assert_fails("evidence digests", "certification.json differs from the copy")

    def test_a_dirty_working_tree_fails(self) -> None:
        self.bundle_with(lambda docs: docs["session.json"].update(dirty=True, allow_dirty=True))
        self.assert_fails("session", "-AllowDirty")

    def test_a_step_that_did_not_pass_fails(self) -> None:
        def compile_failed(docs: dict[str, Any]) -> None:
            docs["session.json"]["steps"][4].update(status="FAIL", detail="compile error")
        self.bundle_with(compile_failed)
        self.assert_fails("session", "compile=FAIL")

    def test_a_failed_run_result_fails(self) -> None:
        self.bundle_with(lambda docs: docs["session.json"].update(result="FAIL", exit_code=1))
        self.assert_fails("session", "ended 'FAIL'")

    def test_a_host_that_disagrees_with_the_workbook_fails(self) -> None:
        self.bundle_with(lambda docs: docs["host.json"]["excel"].update(bitness="x86"))
        self.assert_fails("host", "bitness 'x86' differs")

    @staticmethod
    def set_automation_security(value: object) -> Callable[[dict[str, Any]], None]:
        def change(docs: dict[str, Any]) -> None:
            docs["host.json"]["trust_center"]["automation_security"] = value
        return change

    def test_both_known_ways_of_enabling_the_macros_pass(self) -> None:
        self.assertEqual(evidence.AUTOMATION_SECURITY,
                         ("msoAutomationSecurityByUI", "msoAutomationSecurityLow (-Unattended)"))
        for value in evidence.AUTOMATION_SECURITY:
            with self.subTest(value):
                self.bundle_with(self.set_automation_security(value))
                self.assertEqual(self.failed(), {})

    def test_any_other_way_of_enabling_the_macros_fails(self) -> None:
        for value in ("msoAutomationSecurityLow", "msoAutomationSecurityForceDisable", "MSOAUTOMATIONSECURITYBYUI",
                      "msoAutomationSecurityLow (-Unattended) ", "", None, 1, ["msoAutomationSecurityByUI"]):
            with self.subTest(value=value):
                self.bundle_with(self.set_automation_security(value))
                self.assert_fails("host", "automation_security")

    def test_a_host_record_that_does_not_say_how_the_macros_were_enabled_fails(self) -> None:
        self.bundle_with(lambda docs: docs["host.json"]["trust_center"].pop("automation_security"))
        self.assert_fails("host", "automation_security None must be one of")
        self.bundle_with(lambda docs: docs["host.json"].pop("trust_center"))
        self.assert_fails("host", "has no 'trust_center'")
        self.bundle_with(lambda docs: docs["host.json"].update(trust_center="msoAutomationSecurityByUI"))
        self.assert_fails("host", "trust_center must be an object")

    def test_a_host_record_without_a_windows_build_fails(self) -> None:
        self.bundle_with(lambda docs: docs["host.json"]["windows"].update(build=""))
        self.assert_fails("host", "'build' must be non-empty")

    def test_a_workbook_with_an_extra_component_fails(self) -> None:
        def extra(docs: dict[str, Any]) -> None:
            docs["session.json"]["components"].append({"name": "Class1", "type": 2})
        self.bundle_with(extra)
        self.assert_fails("workbook modules", "Class1 (type 2)")

    def test_a_workbook_without_a_module_fails(self) -> None:
        def missing(docs: dict[str, Any]) -> None:
            docs["session.json"]["components"] = [item for item in docs["session.json"]["components"]
                                                  if item["name"] != "M_EXCEL_UI"]
        self.bundle_with(missing)
        self.assert_fails("workbook modules", "standard modules")

    def test_evidence_written_outside_the_session_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(timestampLocal="2026-11-05 10:03:00"))
        self.assert_fails("timing", "outside the session")

    def test_a_duplicate_json_key_fails(self) -> None:
        self.bundle_with()
        path = self.bundle / "host.json"
        path.write_text('{"schema": 1, "schema": 1}', encoding="utf-8")
        self.assert_fails("host", "duplicate key 'schema'")

    def test_an_unknown_commit_is_a_setup_error(self) -> None:
        self.bundle_with()
        code, _, err = self.run_main(str(self.bundle), "--commit", "0" * 40)
        self.assertEqual(code, 2)
        self.assertIn("is not a commit", err)

    def test_a_missing_bundle_folder_is_a_setup_error(self) -> None:
        code, _, err = self.run_main(str(self.bundle / "absent"))
        self.assertEqual(code, 2)
        self.assertIn("is not a folder", err)

    def test_a_bundle_and_inventory_together_are_a_usage_error(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as raised:
            evidence.main(["--inventory", str(self.bundle)])
        self.assertEqual(raised.exception.code, 2)


class InventoryTests(Fixture):
    """The case inventory is a part of the verdict that the unit counters cannot supply."""

    def assert_fails(self, check: str, fragment: str) -> None:
        failed = self.failed()
        self.assertIn(check, failed)
        self.assertIn(fragment, failed[check])

    @staticmethod
    def rows(docs: dict[str, Any]) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = docs["certification.json"]["caseResults"]
        return rows

    def test_a_pass_without_case_results_is_rejected(self) -> None:
        def earlier_runner(docs: dict[str, Any]) -> None:
            for key in INVENTORY_KEYS:
                del docs["certification.json"][key]
        self.bundle_with(earlier_runner)
        self.assert_fails("case inventory", "no case inventory")
        self.assertNotIn("verdict", self.failed())

    def test_some_inventory_fields_without_the_others_fail(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].pop("caseResults"))
        self.assert_fails("verdict", "the case inventory is itemized but has no caseResults")

    def test_a_missing_case_fails_even_when_every_unit_passed(self) -> None:
        def missing(docs: dict[str, Any]) -> None:
            self.rows(docs).pop(1)
            for index, row in enumerate(self.rows(docs), 1):
                row["seq"] = index
        self.bundle_with(missing)
        self.assert_fails("case inventory", "missing [('RegressionPack', 'TST_Case_B')]")
        self.assertEqual(self.base["certification.json"]["unitsFailed"], 0)

    def test_a_repeated_case_fails(self) -> None:
        def repeated(docs: dict[str, Any]) -> None:
            self.rows(docs).append(dict(self.rows(docs)[0], seq=len(self.rows(docs)) + 1))
        self.bundle_with(repeated)
        self.assert_fails("case inventory", "repeated [('RegressionPack', 'TST_Case_A')]")

    def test_a_case_the_manifest_does_not_list_fails(self) -> None:
        def unexpected(docs: dict[str, Any]) -> None:
            self.rows(docs).append({"seq": 5, "unit": "", "case": "TST_Case_New", "started": True,
                                    "status": "PASS", "detail": ""})
        self.bundle_with(unexpected)
        self.assert_fails("case inventory", "unexpected [('', 'TST_Case_New')]")

    def test_a_case_that_failed_or_never_reported_fails(self) -> None:
        self.bundle_with(lambda docs: self.rows(docs)[0].update(status="FAIL", detail="1004: failed"))
        self.assert_fails("case inventory", "cases that did not start and pass: TST_Case_A")
        self.bundle_with(lambda docs: self.rows(docs)[2].update(status="NO-RESULT"))
        self.assert_fails("case inventory", "cases that did not start and pass: Test_EXCEL_UI_RunSnapshotIdentity")

    def test_cases_in_another_order_fail(self) -> None:
        def swapped(docs: dict[str, Any]) -> None:
            rows = self.rows(docs)
            rows[0], rows[1] = rows[1], rows[0]
            for index, row in enumerate(rows, 1):
                row["seq"] = index
        self.bundle_with(swapped)
        self.assert_fails("case inventory", "another order")

    def test_expected_cases_that_are_not_the_manifest_fail(self) -> None:
        def shortened(docs: dict[str, Any]) -> None:
            docs["certification.json"]["expectedCases"].pop()
        self.bundle_with(shortened)
        self.assert_fails("case inventory", "expectedCases are not the commit's manifest")

    def test_a_manifest_version_that_differs_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(manifestVersion=1))
        self.assert_fails("case inventory", "manifestVersion is 1, the commit's manifest is 2")

    @staticmethod
    def inexact(docs: dict[str, Any]) -> None:
        docs["certification.json"].update(
            complete=False, passed=False, inventory="FAILED",
            inventoryDetail="missing:TST_Case_B: the expected case was never started",
            inventoryFindings=[{"seq": 1, "code": "missing:TST_Case_B",
                                "detail": "the expected case was never started"}])

    def test_an_inexact_inventory_fails_the_verdict_and_is_named(self) -> None:
        self.bundle_with(self.inexact)
        failures = self.failed()
        self.assertIn("the verdict is FAIL | INCOMPLETE | units=3 failed=0 skipped=0 cleanup=OK", failures["verdict"])
        self.assertIn("case inventory FAILED: missing:TST_Case_B", failures["verdict"])
        self.assertIn("missing:TST_Case_B", failures["case inventory"])
        self.assertNotIn("report", failures)

    def test_an_inventory_that_claims_ok_with_a_finding_or_detail_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(
            inventoryFindings=[{"seq": 1, "code": "missing:TST_Case_B", "detail": "never started"}]))
        self.assert_fails("verdict", "inventory is OK but inventoryFindings is not empty")
        self.bundle_with(lambda docs: docs["certification.json"].update(inventoryDetail="something"))
        self.assert_fails("verdict", "inventory is OK but inventoryDetail is not empty")

    def test_an_inventory_that_failed_without_saying_why_fails(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(inventory="FAILED"))
        self.assert_fails("verdict", "inventory FAILED without a finding or an inventoryDetail")

    def test_malformed_inventory_rows_fail(self) -> None:
        self.bundle_with(lambda docs: self.rows(docs)[0].update(status="DONE"))
        self.assert_fails("verdict", "status must be one of")
        self.bundle_with(lambda docs: self.rows(docs)[0].update(started="yes"))
        self.assert_fails("verdict", "started must be true or false")
        self.bundle_with(lambda docs: self.rows(docs)[1].update(seq=7))
        self.assert_fails("verdict", "caseResults[1] must have seq 2")
        self.bundle_with(lambda docs: docs["certification.json"].update(manifestVersion="1"))
        self.assert_fails("verdict", "manifestVersion must be a positive integer")

    def test_a_report_without_the_case_lines_fails(self) -> None:
        report = report_text(self.base["certification.json"]).replace(
            "  CASE  PASS  RegressionPack  TST_Case_A\r\n", "")
        self.bundle_with(report=report)
        self.assert_fails("report", "CASE lines")

    def test_a_report_whose_counts_differ_fails(self) -> None:
        report = report_text(self.base["certification.json"]).replace("started=4", "started=3")
        self.bundle_with(report=report)
        self.assert_fails("report", "CASES lines")

    def test_a_report_without_the_expected_lines_fails(self) -> None:
        report = report_text(self.base["certification.json"]).replace("  EXPECTED  RegressionPack  TST_Case_B\r\n", "")
        self.bundle_with(report=report)
        self.assert_fails("report", "EXPECTED lines")

    def test_a_report_that_omits_an_inventory_finding_fails(self) -> None:
        result = copy.deepcopy(self.base)
        self.inexact(result)
        report = report_text(result["certification.json"]).replace(
            "  INVENTORY  missing:TST_Case_B | the expected case was never started\r\n", "")
        self.bundle_with(self.inexact, report=report)
        self.assert_fails("report", "INVENTORY lines")

    def test_a_missing_manifest_at_the_commit_is_reported(self) -> None:
        self.bundle_with()
        git(self.root, "rm", "-q", "tools/certification-manifest.json")
        git(self.root, "commit", "-q", "-m", "drop the manifest")
        later = git(self.root, "rev-parse", "HEAD")
        failed = self.failed(sha=later)
        self.assertIn("has no tools/certification-manifest.json", failed["case inventory"])
        self.assertIn("has no tools/certification-manifest.json", failed["verdict"])

    def test_an_unusable_manifest_at_the_commit_is_reported(self) -> None:
        self.bundle_with()
        (self.root / "tools" / "certification-manifest.json").write_text("{}", encoding="utf-8")
        git(self.root, "commit", "-q", "-am", "break the manifest")
        later = git(self.root, "rev-parse", "HEAD")
        self.assertIn("missing schema_version", self.failed(sha=later)["case inventory"])


class AchievedStateTests(Fixture):
    """A write that was not read back at the requested state fails the bundle, whatever the setter returned."""

    def assert_fails(self, check: str, fragment: str) -> None:
        failed = self.failed()
        self.assertIn(check, failed)
        self.assertIn(fragment, failed[check])

    @staticmethod
    def rows(docs: dict[str, Any]) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = docs["certification.json"]["achievedState"]
        return rows

    def test_a_state_mismatch_fails_the_bundle_and_is_named(self) -> None:
        def mismatch(docs: dict[str, Any]) -> None:
            row = self.rows(docs)[0]
            row.update(observed="visible", outcome="state-mismatch",
                       reason="the setter returned True, but the state read back differs from the request")
            docs["certification.json"]["passed"] = False
            refresh_achieved(docs["certification.json"])
        self.bundle_with(mismatch)
        failed = self.failed()
        self.assertIn("achieved state", failed)
        self.assertIn("state-mismatch", failed["achieved state"])
        self.assertIn("StatusBar.Application.hide", failed["achieved state"])
        self.assertIn("achieved state FAILED", failed["verdict"])

    def test_a_mismatch_fails_even_when_the_document_claims_a_pass(self) -> None:
        def mismatch(docs: dict[str, Any]) -> None:
            self.rows(docs)[1].update(observed="visible", outcome="state-mismatch", reason="not read back")
            refresh_achieved(docs["certification.json"])
        self.bundle_with(mismatch)
        self.assertIn("verdict", self.failed())
        self.assertIn("achieved state", self.failed())

    def test_each_failed_outcome_fails_the_bundle(self) -> None:
        for outcome in ("write-failed", "readback-failed", "state-mismatch", "restore-failed"):
            with self.subTest(outcome):
                def change(docs: dict[str, Any], outcome: str = outcome) -> None:
                    self.rows(docs)[2].update(outcome=outcome, reason=f"{outcome} for the test")
                    refresh_achieved(docs["certification.json"])
                self.bundle_with(change)
                self.assert_fails("achieved state", outcome)

    def test_a_document_without_the_achieved_state_fails_when_the_manifest_requires_it(self) -> None:
        def earlier_runner(docs: dict[str, Any]) -> None:
            for key in ACHIEVED_KEYS:
                del docs["certification.json"][key]
        self.bundle_with(earlier_runner)
        self.assert_fails("achieved state", "no achieved state")

    def test_some_achieved_fields_without_the_others_fail(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].pop("achievedMismatches"))
        self.assert_fails("achieved state", "has no achievedMismatches")

    def test_an_observation_the_manifest_requires_but_the_run_never_made_fails(self) -> None:
        def missing(docs: dict[str, Any]) -> None:
            docs["certification.json"]["achievedState"] = [
                row for row in self.rows(docs) if row["element"] != "Gridlines"]
            refresh_achieved(docs["certification.json"])
        self.bundle_with(missing)
        self.assert_fails("achieved state", "no observation of Gridlines.AllWindows")

    def test_an_element_that_never_changed_or_never_found_the_state_correct_fails(self) -> None:
        for outcome, fragment in (("verified-write", "no hide or show observation that found it correct"),
                                  ("verified-no-op", "no hide or show observation that changed the state")):
            with self.subTest(outcome):
                def one_path(docs: dict[str, Any], outcome: str = outcome) -> None:
                    for row in self.rows(docs):
                        if row["element"] == "StatusBar" and row["operation"] != "restore":
                            row["outcome"] = outcome
                    refresh_achieved(docs["certification.json"])
                self.bundle_with(one_path)
                self.assert_fails("achieved state", fragment)

    def test_a_restoration_that_was_never_read_back_fails(self) -> None:
        def no_restoration(docs: dict[str, Any]) -> None:
            docs["certification.json"]["achievedState"] = [
                row for row in self.rows(docs) if not (row["element"] == "StatusBar" and row["operation"] == "restore")]
            refresh_achieved(docs["certification.json"])
        self.bundle_with(no_restoration)
        self.assert_fails("achieved state", "no restoration that read back")

    def test_only_the_ribbon_may_be_not_verifiable(self) -> None:
        def unverifiable(docs: dict[str, Any]) -> None:
            for row in self.rows(docs):
                if row["element"] == "StatusBar":
                    row.update(outcome="not-verifiable", reason="could not be read")
            refresh_achieved(docs["certification.json"])
        self.bundle_with(unverifiable)
        self.assert_fails("achieved state", "StatusBar.Application")

    def test_a_title_bar_that_was_not_verified_fails_the_bundle(self) -> None:
        """The live verdict fails a captionless entry frame, so the bundle must not accept one either."""
        manifest = copy.deepcopy(CASE_MANIFEST)
        manifest["achieved_state"]["observations"].append({"element": "TitleBar", "scope": "ActiveFrame"})
        text = json.dumps(manifest, indent=2) + "\n"
        (self.root / "tools" / "certification-manifest.json").write_text(text, encoding="utf-8")
        git(self.root, "add", "-A")
        git(self.root, "commit", "-q", "-m", "a manifest that lists the title bar")
        later = git(self.root, "rev-parse", "HEAD")

        def title_bar(docs: dict[str, Any]) -> None:
            self.rows(docs).append({"element": "TitleBar", "scope": "ActiveFrame", "operation": "entry",
                                    "requested": "not attempted", "observed": "unreadable",
                                    "outcome": "not-verifiable", "reason": "frame has no complete caption"})
            refresh_achieved(docs["certification.json"])
        self.bundle_with(title_bar)
        found = {name: detail for name, passed, detail in evidence.evaluate(self.bundle, later, self.root)
                 if name == "achieved state" and not passed}
        self.assertIn("TitleBar.ActiveFrame has a not-verifiable observation", found["achieved state"])

    def test_a_ribbon_that_was_verified_and_then_was_not_fails(self) -> None:
        """Only a Ribbon that is unverifiable for every observation is exempt; a mixed result is not."""
        for operation in ("hide", "restore"):
            with self.subTest(operation):
                def mixed(docs: dict[str, Any], operation: str = operation) -> None:
                    for row in self.rows(docs):
                        if row["element"] == "Ribbon" and row["operation"] == operation:
                            row.update(outcome="not-verifiable", reason="the Ribbon could not be read")
                    refresh_achieved(docs["certification.json"])
                self.bundle_with(mixed)
                self.assert_fails("achieved state", "Ribbon.ActiveWindow was verified and then was not")

    def test_a_ribbon_unverifiable_at_a_step_and_not_at_the_entry_fails(self) -> None:
        def steps(docs: dict[str, Any]) -> None:
            for row in self.rows(docs):
                if row["element"] == "Ribbon":
                    row.update(outcome="not-verifiable", reason="the Ribbon could not be read")
            refresh_achieved(docs["certification.json"])
        self.bundle_with(steps)
        self.assert_fails("achieved state", "at the entry operation")

    def test_an_observation_at_an_operation_the_unit_never_makes_fails(self) -> None:
        self.bundle_with(lambda docs: self.rows(docs)[0].update(operation="toggle"))
        self.assert_fails("achieved state", "'toggle'")

    def test_an_entry_observation_that_claims_to_be_verified_fails(self) -> None:
        def entry(docs: dict[str, Any]) -> None:
            self.rows(docs)[0].update(operation="entry")
            refresh_achieved(docs["certification.json"])
        self.bundle_with(entry)
        self.assert_fails("achieved state", "entry observation that is not an unverifiable one")

    def test_an_unverifiable_observation_without_a_reason_is_malformed(self) -> None:
        def no_reason(docs: dict[str, Any]) -> None:
            self.rows(docs)[0].update(outcome="not-verifiable")
            refresh_achieved(docs["certification.json"])
        self.bundle_with(no_reason)
        self.assert_fails("verdict", "has no reason")

    def test_an_outcome_the_evidence_does_not_define_is_malformed(self) -> None:
        def unknown(docs: dict[str, Any]) -> None:
            self.rows(docs)[0].update(outcome="probably-fine")
            refresh_achieved(docs["certification.json"])
        self.bundle_with(unknown)
        self.assert_fails("verdict", "is not one the evidence defines")

    def test_an_ok_status_over_a_failed_observation_is_malformed(self) -> None:
        def lie(docs: dict[str, Any]) -> None:
            self.rows(docs)[0].update(outcome="state-mismatch", reason="not read back")
            certification = docs["certification.json"]
            refresh_achieved(certification)
            certification["achieved"] = "OK"
            certification["achievedDetail"] = ""
        self.bundle_with(lie)
        self.assert_fails("verdict", "achieved is OK but an observation failed")

    def test_mismatches_that_are_not_the_failed_observations_are_malformed(self) -> None:
        def hide(docs: dict[str, Any]) -> None:
            self.rows(docs)[0].update(outcome="state-mismatch", reason="not read back")
            certification = docs["certification.json"]
            refresh_achieved(certification)
            certification["achievedMismatches"] = []
        self.bundle_with(hide)
        self.assert_fails("verdict", "achievedMismatches are not the failed observations")

    def test_a_count_that_is_not_the_observations_is_malformed(self) -> None:
        self.bundle_with(lambda docs: docs["certification.json"].update(achievedChecked=3))
        self.assert_fails("verdict", "achievedChecked is 3")

    def test_an_observation_without_its_states_is_malformed(self) -> None:
        self.bundle_with(lambda docs: self.rows(docs)[0].update(observed=""))
        self.assert_fails("verdict", "has an empty observed")

    def test_a_report_that_omits_a_mismatch_fails(self) -> None:
        def mismatch(docs: dict[str, Any]) -> None:
            self.rows(docs)[0].update(outcome="state-mismatch", reason="not read back")
            refresh_achieved(docs["certification.json"])
        docs = copy.deepcopy(self.base)
        mismatch(docs)
        honest = report_text(docs["certification.json"])
        omitted = "\r\n".join(line for line in honest.split("\r\n") if not line.startswith("  MISMATCH  "))
        self.bundle_with(mismatch, report=omitted)
        self.assert_fails("report", "MISMATCH lines")

    def test_a_report_that_omits_an_unverifiable_surface_fails(self) -> None:
        def unverifiable(docs: dict[str, Any]) -> None:
            for row in self.rows(docs):
                if row["element"] == "Ribbon":
                    row.update(outcome="not-verifiable", reason="the Ribbon could not be read")
            refresh_achieved(docs["certification.json"])
        docs = copy.deepcopy(self.base)
        unverifiable(docs)
        honest = report_text(docs["certification.json"])
        omitted = "\r\n".join(line for line in honest.split("\r\n") if not line.startswith("  UNVERIFIABLE  "))
        self.bundle_with(unverifiable, report=omitted)
        self.assert_fails("report", "UNVERIFIABLE lines")

    def test_a_report_whose_observation_counts_differ_fails(self) -> None:
        honest = report_text(self.base["certification.json"])
        self.bundle_with(report=honest.replace("checked=15 verified=15", "checked=15 verified=14"))
        self.assert_fails("report", "OBSERVATIONS lines")

    def test_an_earlier_report_with_observation_lines_is_not_read_as_clean(self) -> None:
        def earlier_runner(docs: dict[str, Any]) -> None:
            for key in ACHIEVED_KEYS:
                del docs["certification.json"][key]
        policy = dataclasses.replace(evidence.default_policy(), case_manifest=None)
        self.bundle_with(earlier_runner, report=report_text(self.base["certification.json"]))
        failed = {name for name, passed, _ in evidence.evaluate(self.bundle, self.sha, self.root, policy) if not passed}
        self.assertEqual(failed, {"report"})


class EarlierContractTests(Fixture):
    """A policy without a case manifest describes the contract that had no case inventory."""

    def earlier_policy(self) -> evidence.Policy:
        return dataclasses.replace(evidence.default_policy(), case_manifest=None)

    def outcomes_under(self, policy: evidence.Policy) -> dict[str, tuple[bool, str]]:
        found = evidence.evaluate(self.bundle, self.sha, self.root, policy)
        return {name: (passed, detail) for name, passed, detail in found}

    def test_a_bundle_without_an_inventory_passes_under_the_earlier_policy(self) -> None:
        def earlier_runner(docs: dict[str, Any]) -> None:
            for key in INVENTORY_KEYS:
                del docs["certification.json"][key]
        self.bundle_with(earlier_runner)
        outcomes = self.outcomes_under(self.earlier_policy())
        self.assertEqual({name for name, (passed, _) in outcomes.items() if not passed}, set())
        self.assertIn("predates", outcomes["case inventory"][1])
        # The units come from the commit's certification runner, as they did.
        self.assertEqual(evidence.certification_units(self.root, self.sha), UNITS)

    def test_the_current_policy_rejects_the_same_bundle(self) -> None:
        def earlier_runner(docs: dict[str, Any]) -> None:
            for key in INVENTORY_KEYS:
                del docs["certification.json"][key]
        self.bundle_with(earlier_runner)
        self.assertIn("case inventory", self.failed())

    def test_an_earlier_report_with_case_lines_is_not_read_as_clean(self) -> None:
        def earlier_runner(docs: dict[str, Any]) -> None:
            for key in INVENTORY_KEYS:
                del docs["certification.json"][key]
        self.bundle_with(earlier_runner, report=report_text(self.base["certification.json"]))
        failed = {name for name, (passed, _) in self.outcomes_under(self.earlier_policy()).items() if not passed}
        self.assertEqual(failed, {"report"})

    def test_a_bundle_with_an_inventory_also_passes_under_the_earlier_policy(self) -> None:
        self.bundle_with()
        outcomes = self.outcomes_under(self.earlier_policy())
        self.assertEqual({name for name, (passed, _) in outcomes.items() if not passed}, set())


@unittest.skipUnless((ROOT / ".git").exists(), "needs a Git checkout of this repository")
class RepositoryTests(unittest.TestCase):
    def test_this_repository_inventory_and_units_are_readable(self) -> None:
        sha = evidence.resolve_commit(ROOT, "HEAD")
        inventory = evidence.source_inventory(ROOT, sha)
        self.assertIn("test/M_EXCEL_UI_TEST_CERTIFICATION.bas", [module["path"] for module in inventory["modules"]])
        for module in inventory["modules"]:
            self.assertEqual(module["name"], Path(module["path"]).stem)
        units = evidence.certification_units(ROOT, sha)
        self.assertTrue(units)
        self.assertEqual(len(units), len(set(units)))


if __name__ == "__main__":
    unittest.main()
