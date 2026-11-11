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
RUNNER = (
    "Attribute VB_Name = \"M_EXCEL_UI_REGRESSION_TESTS\"\n"
    "Option Explicit\n\n"
    "Public Sub Test_EXCEL_UI_RunReleaseCertification()\n"
    + "".join(f"        TST_CertRunUnit \"{unit}\"\n" for unit in UNITS)
    + "End Sub\n\n"
    "Private Sub TST_CertRunUnit( _\n    ByVal UnitName As String)\nEnd Sub\n"
)
DEVELOPMENT = ["src/M_EXCEL_UI.bas", "test/M_EXCEL_UI_REGRESSION_TESTS.bas"]
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
    "src/M_EXCEL_UI.bas": "Attribute VB_Name = \"M_EXCEL_UI\"\nOption Explicit\n\n'café is spelled in cp1252\n",
    "test/M_EXCEL_UI_REGRESSION_TESTS.bas": RUNNER,
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


def result_line(result: dict[str, Any]) -> str:
    return (f"RESULT: {'PASS' if result['passed'] else 'FAIL'} | "
            f"{'COMPLETE' if result['complete'] else 'INCOMPLETE'} | units={result['units']} "
            f"failed={result['unitsFailed']} skipped={result['skipped']} cleanup={result['cleanup']}")


def report_text(result: dict[str, Any]) -> str:
    """The text report, as TST_CertBuildTextReport composes it."""
    lines = ["VBA Excel UI - release certification",
             f"Excel {result['excelVersion']} build {result['excelBuild']}",
             result["operatingSystem"], result["timestampLocal"], "-" * 60, result_line(result)]
    lines += [f"  {'PASS' if unit['passed'] else 'FAIL'}  {unit['name']}"
              + (f" | {unit['detail']}" if unit["detail"] else "") for unit in result["unitResults"]]
    lines += [f"  SKIP  {item}" for item in result["skipDetail"]]
    if result["cleanupDetail"]:
        lines.append(f"  CLEANUP  {result['cleanupDetail']}")
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
        "cleanup": "OK", "cleanupDetail": "", "complete": True, "passed": True,
        "unitResults": [{"name": unit, "passed": True, "detail": ""} for unit in UNITS],
        "skipDetail": [],
    }
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
                         ["M_EXCEL_UI", "M_EXCEL_UI_REGRESSION_TESTS"])
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
        self.assertEqual(len(evidence.CHECKS), 9)

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
        self.assert_fails("source digests", "checked-out test/M_EXCEL_UI_REGRESSION_TESTS.bas differs")

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


@unittest.skipUnless((ROOT / ".git").exists(), "needs a Git checkout of this repository")
class RepositoryTests(unittest.TestCase):
    def test_this_repository_inventory_and_units_are_readable(self) -> None:
        sha = evidence.resolve_commit(ROOT, "HEAD")
        inventory = evidence.source_inventory(ROOT, sha)
        self.assertIn("test/M_EXCEL_UI_REGRESSION_TESTS.bas", [module["path"] for module in inventory["modules"]])
        for module in inventory["modules"]:
            self.assertEqual(module["name"], Path(module["path"]).stem)
        units = evidence.certification_units(ROOT, sha)
        self.assertTrue(units)
        self.assertEqual(len(units), len(set(units)))


if __name__ == "__main__":
    unittest.main()
