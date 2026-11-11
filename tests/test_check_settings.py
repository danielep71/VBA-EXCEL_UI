"""Tests for the repository settings drift check (tools/check_settings.py)."""

from __future__ import annotations

import contextlib
import copy
import io
import json
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_settings  # noqa: E402

# Responses recorded with an administrator token, trimmed to the fields the
# check reads plus a few it must ignore.
ADMIN: dict[str, dict[str, Any]] = {
    "repository": {"status": 200, "body": {
        "id": 1, "full_name": "example-org/sample-project", "visibility": "public",
        "permissions": {"admin": True, "push": True, "pull": True},
        "has_issues": True, "has_projects": False, "has_wiki": True, "has_discussions": False,
        "allow_merge_commit": True, "allow_squash_merge": True, "allow_rebase_merge": True,
        "allow_auto_merge": False, "delete_branch_on_merge": True, "allow_update_branch": True,
        "allow_forking": True, "web_commit_signoff_required": False,
        "security_and_analysis": {
            "secret_scanning": {"status": "enabled"},
            "secret_scanning_push_protection": {"status": "enabled"},
            "dependabot_security_updates": {"status": "enabled"},
            "secret_scanning_validity_checks": {"status": "disabled"},
        },
    }},
    "private-vulnerability-reporting": {"status": 200, "body": {"enabled": True}},
    "vulnerability-alerts": {"status": 204},
    "actions/permissions": {"status": 200, "body": {
        "enabled": True, "allowed_actions": "selected",
        "selected_actions_url": "https://example.invalid"}},
    "actions/permissions/selected-actions": {"status": 200, "body": {
        "github_owned_allowed": True, "verified_allowed": False, "patterns_allowed": []}},
    "actions/permissions/workflow": {"status": 200, "body": {
        "default_workflow_permissions": "read", "can_approve_pull_request_reviews": False}},
}

# What a read-only token sees: no permissions or security fields, and the
# administrator-only endpoints refused.
READ_ONLY: dict[str, dict[str, Any]] = {
    **ADMIN,
    "repository": {"status": 200, "body": {
        key: value for key, value in ADMIN["repository"]["body"].items()
        if key not in ("permissions", "security_and_analysis")}},
    "vulnerability-alerts": {"status": 404, "body": {"message": "Not Found"}},
    "actions/permissions": {"status": 403, "body": {"message": "Resource not accessible"}},
    "actions/permissions/selected-actions": {"status": 403},
    "actions/permissions/workflow": {"status": 403},
}


class SettingsDriftTests(unittest.TestCase):
    def setUp(self) -> None:
        self.root = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.root)
        target = self.root / check_settings.SETTINGS_FILE
        target.parent.mkdir(parents=True)
        shutil.copyfile(ROOT / check_settings.SETTINGS_FILE, target)

    def run_tool(self, recorded: dict[str, Any], *extra: str) -> tuple[int, str]:
        source = self.write(recorded)
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            code = check_settings.main(["--root", str(self.root), "--input", str(source), *extra])
        return code, out.getvalue()

    def outcomes(self, recorded: dict[str, Any]) -> dict[str, check_settings.Result]:
        fetch = check_settings.cached(check_settings.recorded_fetcher(self.write(recorded)))
        results = check_settings.compare(check_settings.load_expected(self.root), fetch)
        return {result.name: result for result in results}

    def write(self, recorded: dict[str, Any]) -> Path:
        path = self.root / "recorded.json"
        path.write_text(json.dumps(recorded), encoding="utf-8")
        return path

    def test_every_setting_matches_the_administrator_recording(self) -> None:
        code, text = self.run_tool(ADMIN)
        self.assertEqual(code, 0)
        self.assertIn("Status: **PASS**", text)
        self.assertNotIn("## Differences", text)
        self.assertNotIn("## Not verifiable", text)
        self.assertEqual({r.outcome for r in self.outcomes(ADMIN).values()}, {"match"})

    def test_changed_setting_is_a_difference(self) -> None:
        recorded = copy.deepcopy(ADMIN)
        recorded["repository"]["body"]["security_and_analysis"][
            "secret_scanning_push_protection"]["status"] = "disabled"
        recorded["actions/permissions/workflow"]["body"]["default_workflow_permissions"] = "write"
        code, text = self.run_tool(recorded)
        self.assertEqual(code, 1)
        self.assertIn("Status: **FAIL**", text)
        self.assertIn('- `security_and_analysis.secret_scanning_push_protection`: '
                      'expected `"enabled"`, live `"disabled"`', text)
        self.assertIn('- `actions.default_workflow_permissions`: expected `"read"`, '
                      'live `"write"`', text)

    def test_unreadable_settings_are_reported_but_do_not_fail(self) -> None:
        code, text = self.run_tool(READ_ONLY)
        self.assertEqual(code, 0)
        self.assertIn("Status: **INCOMPLETE**", text)
        self.assertIn("- `actions.enabled`: HTTP 403 from actions/permissions", text)
        self.assertIn("- `security_and_analysis.secret_scanning`: repository has no "
                      "security_and_analysis.secret_scanning.status field", text)
        self.assertIn("- `security_and_analysis.dependabot_alerts`: HTTP 404 from "
                      "vulnerability-alerts without administrator access", text)
        self.assertEqual(self.run_tool(READ_ONLY, "--strict")[0], 1)

    def test_match_difference_and_unreadable_are_distinct_in_one_run(self) -> None:
        recorded = copy.deepcopy(READ_ONLY)
        recorded["repository"]["body"]["allow_squash_merge"] = False
        results = self.outcomes(recorded)
        self.assertEqual(results["general.allow_squash_merge"].outcome, "differ")
        self.assertIs(results["general.allow_squash_merge"].detail, False)
        self.assertEqual(results["general.has_issues"].outcome, "match")
        self.assertEqual(results["actions.allowed_actions"].outcome, "unverifiable")
        code, text = self.run_tool(recorded)
        self.assertEqual(code, 1)
        for heading in ("## Differences", "## Not verifiable", "## Matches"):
            self.assertIn(heading, text)

    def test_dependabot_alerts_off_is_a_difference_for_an_administrator(self) -> None:
        recorded = copy.deepcopy(ADMIN)
        recorded["vulnerability-alerts"] = {"status": 404}
        result = self.outcomes(recorded)["security_and_analysis.dependabot_alerts"]
        self.assertEqual((result.outcome, result.detail), ("differ", "disabled"))

    def test_server_error_and_missing_recording_are_errors(self) -> None:
        recorded = copy.deepcopy(ADMIN)
        recorded["actions/permissions/workflow"] = {"status": 500}
        self.assertEqual(self.run_tool(recorded)[0], 2)
        del recorded["actions/permissions/workflow"]
        self.assertEqual(self.run_tool(recorded)[0], 2)

    def test_unknown_or_unnormalized_settings_file_is_an_error(self) -> None:
        path = self.root / check_settings.SETTINGS_FILE
        data = json.loads(path.read_text(encoding="utf-8"))
        path.write_text(json.dumps(data), encoding="utf-8")
        self.assertEqual(self.run_tool(ADMIN)[0], 2)
        data["general"]["allow_squash"] = False
        path.write_text(check_settings.render(data), encoding="utf-8")
        self.assertEqual(self.run_tool(ADMIN)[0], 2)

    def test_settings_file_in_this_repository_is_valid(self) -> None:
        expected = check_settings.load_expected(ROOT)
        self.assertEqual(set(expected), set(check_settings.READERS))
        self.assertIs(expected["general"]["allow_merge_commit"], True)

    def test_summary_is_appended(self) -> None:
        summary = self.root / "summary.md"
        summary.write_text("before\n", encoding="utf-8")
        self.run_tool(ADMIN, "--summary", str(summary))
        self.assertTrue(summary.read_text(encoding="utf-8").startswith(
            "before\n# Repository settings drift\n"))


if __name__ == "__main__":
    unittest.main()
