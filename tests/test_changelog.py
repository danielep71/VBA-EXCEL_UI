"""Structure of CHANGELOG.md and agreement of VERSION with it (#83).

The rules run inside the "release state" check; test_gate_checks.ReleaseState
proves they are wired into it. These tests exercise each rule directly.
"""

import unittest

from gate_fixture import REPO, check_repo

URL = "https://github.com/danielep71/VBA-EXCEL_UI"

GOOD = f"""\
# Changelog

## Policy

| Release heading | `## [X.Y.Z] - YYYY-MM-DD` |

## [Unreleased]

### Added

- Something.

```text
## [9.9.9] - 2099-99-99
```

## [1.1.0] - 2026-08-19

- Feature.

## [1.0.1] - 2026-07-25

- Fix.

## [1.0.0] - 2026-07-01

- First.

[Unreleased]: {URL}/compare/v1.1.0...HEAD
[1.1.0]: {URL}/compare/v1.0.1...v1.1.0
[1.0.1]: {URL}/compare/v1.0.0...v1.0.1
[1.0.0]: {URL}/releases/tag/v1.0.0
"""


def findings(text: str = GOOD, version: str | None = "1.1.0") -> list[str]:
    return check_repo.changelog_findings(text, version)


def changed(old: str, new: str) -> str:
    if old not in GOOD:
        raise AssertionError(f"fixture does not contain {old!r}")
    return GOOD.replace(old, new, 1)


class ChangelogStructure(unittest.TestCase):

    def assertFinding(self, found: list[str], fragment: str) -> None:
        self.assertTrue(any(fragment in f for f in found), f"no finding mentions {fragment!r}: {found}")

    def test_the_current_changelog_and_version_pass(self) -> None:
        text = (REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        version = (REPO / "VERSION").read_text(encoding="utf-8").strip()
        self.assertEqual(check_repo.changelog_findings(text, version), [])

    def test_a_well_formed_changelog_passes(self) -> None:
        self.assertEqual(findings(), [])

    def test_a_heading_inside_a_fence_is_not_a_section(self) -> None:
        # GOOD carries "## [9.9.9] - 2099-99-99" in a fence; it must be ignored.
        self.assertEqual(findings(), [])

    # Unreleased first ----------------------------------------------------
    def test_unreleased_after_a_release_fails(self) -> None:
        text = changed("## [Unreleased]\n", "").replace(
            "## [1.0.1] - 2026-07-25\n", "## [Unreleased]\n\n## [1.0.1] - 2026-07-25\n")
        found = findings(text)
        self.assertFinding(found, "the first version section must be '## [Unreleased]'")
        self.assertFinding(found, "'## [Unreleased]' must be the first version section")

    def test_a_missing_unreleased_section_fails(self) -> None:
        found = findings(changed("## [Unreleased]\n", ""))
        self.assertFinding(found, "the first version section must be '## [Unreleased]'; "
                                  "found '## [1.1.0] - 2026-08-19'")

    # Headings and dates --------------------------------------------------
    def test_a_heading_without_the_dash_fails(self) -> None:
        self.assertFinding(findings(changed("## [1.0.1] - 2026-07-25", "## [1.0.1] 2026-07-25")),
                           "malformed release heading '## [1.0.1] 2026-07-25'")

    def test_a_heading_without_brackets_fails(self) -> None:
        self.assertFinding(findings(changed("## [1.0.1] - 2026-07-25", "## 1.0.1 - 2026-07-25")),
                           "malformed release heading '## 1.0.1 - 2026-07-25'")

    def test_a_two_part_version_fails(self) -> None:
        self.assertFinding(findings(changed("## [1.0.1] - 2026-07-25", "## [1.0] - 2026-07-25")),
                           "malformed release heading")

    def test_an_impossible_date_fails(self) -> None:
        self.assertFinding(findings(changed("2026-07-25", "2026-02-30")),
                           "[1.0.1] has an invalid date '2026-02-30'")

    def test_a_reversed_date_fails(self) -> None:
        self.assertFinding(findings(changed("2026-07-25", "2026-25-07")), "has an invalid date")

    # Order ---------------------------------------------------------------
    def test_releases_out_of_version_order_fail(self) -> None:
        text = changed("## [1.1.0] - 2026-08-19", "## [1.0.0] - 2026-08-19")
        text = text.replace("## [1.0.0] - 2026-07-01", "## [1.1.0] - 2026-07-01")
        self.assertFinding(findings(text, None), "follows [1.0.0]; releases must be newest first")

    def test_a_repeated_release_fails(self) -> None:
        self.assertFinding(findings(changed("## [1.0.1] - 2026-07-25", "## [1.1.0] - 2026-07-25")),
                           "[1.1.0] follows [1.1.0]")

    def test_releases_out_of_date_order_fail(self) -> None:
        self.assertFinding(findings(changed("2026-07-25", "2026-09-01")),
                           "[1.0.1] is dated 2026-09-01, after the newer [1.1.0] (2026-08-19)")

    def test_a_pre_release_sorts_before_its_release(self) -> None:
        self.assertLess(check_repo.version_key("1.2.0-rc.1"), check_repo.version_key("1.2.0"))
        self.assertLess(check_repo.version_key("1.1.9"), check_repo.version_key("1.2.0-rc.1"))
        self.assertLess(check_repo.version_key("1.9.0"), check_repo.version_key("1.10.0"))

    def test_pre_release_identifiers_follow_semver_precedence(self) -> None:
        key = check_repo.version_key
        self.assertLess(key("1.2.0-rc.2"), key("1.2.0-rc.10"))
        self.assertLess(key("1.2.0-rc"), key("1.2.0-rc.1"))
        self.assertLess(key("1.2.0-1"), key("1.2.0-alpha"))
        self.assertLess(key("1.2.0-alpha.beta"), key("1.2.0-beta"))

    def test_pre_releases_newest_first_pass(self) -> None:
        text = changed("## [1.1.0] - 2026-08-19\n",
                       "## [1.2.0-rc.10] - 2026-09-02\n\n- RC.\n\n"
                       "## [1.2.0-rc.2] - 2026-09-01\n\n- RC.\n\n## [1.1.0] - 2026-08-19\n")
        text = text.replace(f"[Unreleased]: {URL}/compare/v1.1.0...HEAD",
                            f"[Unreleased]: {URL}/compare/v1.2.0-rc.10...HEAD\n"
                            f"[1.2.0-rc.10]: {URL}/compare/v1.2.0-rc.2...v1.2.0-rc.10\n"
                            f"[1.2.0-rc.2]: {URL}/compare/v1.1.0...v1.2.0-rc.2")
        self.assertEqual(findings(text, "1.2.0-rc.10"), [])
        swapped = text.replace("## [1.2.0-rc.10] - 2026-09-02", "## [1.2.0-rc.X] - 2026-09-02")
        swapped = swapped.replace("## [1.2.0-rc.2] - 2026-09-01", "## [1.2.0-rc.10] - 2026-09-01")
        swapped = swapped.replace("## [1.2.0-rc.X] - 2026-09-02", "## [1.2.0-rc.2] - 2026-09-02")
        self.assertFinding(findings(swapped, "1.2.0-rc.2"),
                           "[1.2.0-rc.10] follows [1.2.0-rc.2]; releases must be newest first")

    # Link references -----------------------------------------------------
    def test_a_missing_unreleased_link_reference_fails(self) -> None:
        self.assertFinding(findings(changed(f"[Unreleased]: {URL}/compare/v1.1.0...HEAD\n", "")),
                           f"[Unreleased] has no link reference; expected "
                           f"[Unreleased]: {URL}/compare/v1.1.0...HEAD")

    def test_a_release_without_a_link_reference_fails(self) -> None:
        self.assertFinding(findings(changed(f"[1.0.1]: {URL}/compare/v1.0.0...v1.0.1\n", "")),
                           "release [1.0.1] has no link reference")

    def test_a_link_to_another_repository_fails(self) -> None:
        other = "https://github.com/someone/else/compare/v1.0.0...v1.0.1"
        found = findings(changed(f"{URL}/compare/v1.0.0...v1.0.1", other))
        self.assertFinding(found, f"link reference [1.0.1] points outside this repository: {other}")

    def test_any_foreign_link_reference_fails(self) -> None:
        found = findings(GOOD + "[guide]: https://example.com/guide\n")
        self.assertFinding(found, "link reference [guide] points outside this repository")

    def test_a_compare_link_from_the_wrong_base_fails(self) -> None:
        self.assertFinding(findings(changed(f"{URL}/compare/v1.0.1...v1.1.0", f"{URL}/compare/v1.0.0...v1.1.0")),
                           f"expected {URL}/compare/v1.0.1...v1.1.0")

    def test_a_link_to_another_release_fails(self) -> None:
        self.assertFinding(findings(changed(f"{URL}/releases/tag/v1.0.0", f"{URL}/releases/tag/v1.0.1")),
                           f"[1.0.0] links to {URL}/releases/tag/v1.0.1")

    def test_a_tag_link_is_accepted_for_any_release(self) -> None:
        self.assertEqual(findings(changed(f"{URL}/compare/v1.0.0...v1.0.1", f"{URL}/releases/tag/v1.0.1")), [])

    def test_an_unreleased_link_not_from_the_newest_release_fails(self) -> None:
        self.assertFinding(findings(changed(f"{URL}/compare/v1.1.0...HEAD", f"{URL}/compare/v1.0.1...HEAD")),
                           f"[Unreleased] links to {URL}/compare/v1.0.1...HEAD")

    def test_a_link_reference_defined_twice_fails(self) -> None:
        self.assertFinding(findings(GOOD + f"[1.0.0]: {URL}/releases/tag/v1.0.0\n"),
                           "link reference [1.0.0] is defined twice")

    # VERSION ---------------------------------------------------------------
    def test_version_ahead_of_the_newest_release_fails(self) -> None:
        self.assertFinding(findings(version="1.1.1"),
                           "VERSION is 1.1.1, but the newest release in CHANGELOG.md is [1.1.0]")

    def test_version_behind_the_newest_release_fails(self) -> None:
        self.assertFinding(findings(version="1.0.1"), "VERSION is 1.0.1")

    def test_version_without_any_release_fails(self) -> None:
        text = "## [Unreleased]\n\n- Work.\n"
        self.assertFinding(findings(text, "1.0.0"), "VERSION is 1.0.0, but CHANGELOG.md has no release heading")


if __name__ == "__main__":
    unittest.main()
