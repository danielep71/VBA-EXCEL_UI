"""Tests for the ruleset drift check (tools/check_rulesets.py)."""

from __future__ import annotations

import contextlib
import copy
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_rulesets  # noqa: E402

LIVE = {
    "id": 1,
    "node_id": "RRS_x",
    "name": "Protect main",
    "target": "branch",
    "source_type": "Repository",
    "source": "example-org/sample-project",
    "enforcement": "active",
    "conditions": {"ref_name": {"exclude": [], "include": ["~DEFAULT_BRANCH"]}},
    "rules": [
        {"type": "pull_request", "parameters": {"required_approving_review_count": 0}},
        {"type": "deletion"},
        {"type": "non_fast_forward"},
    ],
    "created_at": "2026-01-01T00:00:00Z",
    "updated_at": "2026-01-02T00:00:00Z",
    "_links": {"self": {"href": "https://example.invalid"}},
}


class RulesetDriftTests(unittest.TestCase):
    def run_tool(self, root: Path, live: list[dict[str, object]], *extra: str) -> int:
        source = root / "live.json"
        source.write_text(json.dumps(live), encoding="utf-8")
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            return check_rulesets.main(["--root", str(root), "--input", str(source), *extra])

    def exported(self, live: list[dict[str, object]]) -> Path:
        root = Path(tempfile.mkdtemp())
        self.assertEqual(self.run_tool(root, live, "--export"), 0)
        return root

    def test_normalization_drops_generated_fields_and_sorts_rules(self) -> None:
        normalized = check_rulesets.normalize(copy.deepcopy(LIVE))
        self.assertEqual(sorted(normalized), sorted(check_rulesets.KEPT))
        self.assertEqual([rule["type"] for rule in normalized["rules"]],
                         ["deletion", "non_fast_forward", "pull_request"])

    def test_matching_rulesets_pass(self) -> None:
        root = self.exported([LIVE])
        self.assertTrue((root / ".github/rulesets/protect-main.json").is_file())
        self.assertEqual(self.run_tool(root, [LIVE]), 0)

    def test_changed_parameter_is_drift(self) -> None:
        root = self.exported([LIVE])
        weakened = copy.deepcopy(LIVE)
        weakened["rules"] = [rule for rule in LIVE["rules"]  # type: ignore[attr-defined]
                             if rule["type"] != "non_fast_forward"]
        self.assertEqual(self.run_tool(root, [weakened]), 1)

    def test_disabled_ruleset_is_drift(self) -> None:
        root = self.exported([LIVE])
        self.assertEqual(self.run_tool(root, [dict(LIVE, enforcement="disabled")]), 1)

    def test_unversioned_and_deleted_rulesets_are_drift(self) -> None:
        root = self.exported([LIVE])
        self.assertEqual(self.run_tool(root, [LIVE, dict(LIVE, name="New rule")]), 1)
        self.assertEqual(self.run_tool(root, []), 1)

    def test_hand_edited_file_must_be_normalized(self) -> None:
        root = self.exported([LIVE])
        path = root / ".github/rulesets/protect-main.json"
        path.write_text(json.dumps(dict(LIVE)), encoding="utf-8")
        self.assertEqual(self.run_tool(root, [LIVE]), 2)

    def test_versioned_copies_in_this_repository_are_normalized(self) -> None:
        versioned = check_rulesets.load_versioned(ROOT)
        self.assertIn("Protect main", versioned)
        for name, ruleset in versioned.items():
            path = ROOT / ".github/rulesets" / check_rulesets.file_name(name)
            self.assertEqual(path.read_text(encoding="utf-8"), check_rulesets.render(ruleset))


if __name__ == "__main__":
    unittest.main()
