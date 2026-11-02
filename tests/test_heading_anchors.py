"""Heading anchors computed the way GitHub computes them (#84).

The markdown-links check resolves "#anchor" links against these, so an anchor
computed differently reports working links as broken and lets broken ones
pass. test_gate_checks.MarkdownLinks proves the check uses them.
"""

import unittest

from gate_fixture import GateTestCase, check_repo


def anchor(heading: str) -> str:
    return check_repo.github_anchor(heading)


def anchors(*lines: str) -> list[str]:
    return check_repo.heading_anchors("\n".join(lines) + "\n")


class GithubAnchor(unittest.TestCase):

    def test_the_cases_reported_in_the_issue(self) -> None:
        self.assertEqual(anchor("`UI_SetExcelUI` reference"), "ui_setexcelui-reference")
        self.assertEqual(anchor("Café — notes"), "café--notes")
        self.assertEqual(anchors("## Notes", "## Notes"), ["notes", "notes-1"])
        self.assertEqual(anchor("🧭 Release profile"), "-release-profile")

    def test_emoji_are_dropped_and_their_space_kept(self) -> None:
        self.assertEqual(anchor("🧪 Validation and evidence"), "-validation-and-evidence")
        self.assertEqual(anchor("Done ✅"), "done-")

    def test_code_spans_keep_their_content(self) -> None:
        self.assertEqual(anchor("The `UI_Hide` member"), "the-ui_hide-member")
        self.assertEqual(anchor("``a `tick` b``"), "a-tick-b")
        self.assertEqual(anchor("Run `tools/check_repo.py`"), "run-toolscheck_repopy")

    def test_underscores_are_kept_but_emphasis_is_not(self) -> None:
        self.assertEqual(anchor("M_EXCEL_UI module"), "m_excel_ui-module")
        self.assertEqual(anchor("snake_case and __dunder__"), "snake_case-and-dunder")
        self.assertEqual(anchor("_emphasis_ here"), "emphasis-here")
        self.assertEqual(anchor("**bold** and *italic*"), "bold-and-italic")

    def test_non_ascii_letters_and_digits_are_kept(self) -> None:
        self.assertEqual(anchor("Ünïcödé Straße"), "ünïcödé-straße")
        self.assertEqual(anchor("日本語 ２"), "日本語-２")
        self.assertEqual(anchor("Größe 10"), "größe-10")

    def test_punctuation_is_dropped_without_collapsing_hyphens(self) -> None:
        self.assertEqual(anchor("1. Executive assessment"), "1-executive-assessment")
        self.assertEqual(anchor("What's new? (v1.1.3)"), "whats-new-v113")
        self.assertEqual(anchor("A - B -- C"), "a---b----c")
        self.assertEqual(anchor("Q&A"), "qa")

    def test_links_images_html_and_entities_render_as_text(self) -> None:
        self.assertEqual(anchor("See [the guide](CONTRIBUTING.md)"), "see-the-guide")
        self.assertEqual(anchor("Logo ![banner](images/x.png) here"), "logo--here")
        self.assertEqual(anchor('<a id="x"></a>Install'), "install")
        self.assertEqual(anchor("Fish &amp; chips"), "fish--chips")

    def test_closing_hashes_are_not_heading_text(self) -> None:
        self.assertEqual(anchors("## Notes ##"), ["notes"])


class HeadingAnchors(unittest.TestCase):

    def test_duplicates_are_numbered_in_order(self) -> None:
        self.assertEqual(anchors("# Notes", "## Notes", "### Notes"), ["notes", "notes-1", "notes-2"])

    def test_duplicates_differing_only_in_markup_are_numbered(self) -> None:
        self.assertEqual(anchors("## `Notes`", "## *Notes*"), ["notes", "notes-1"])

    def test_a_suffixed_anchor_never_collides_with_a_real_heading(self) -> None:
        self.assertEqual(anchors("## Notes 1", "## Notes", "## Notes"), ["notes-1", "notes", "notes-2"])

    def test_headings_in_fenced_blocks_are_ignored(self) -> None:
        self.assertEqual(anchors("## Notes", "```bash", "# Notes", "```", "## Notes"), ["notes", "notes-1"])

    def test_lines_that_are_not_atx_headings_are_ignored(self) -> None:
        self.assertEqual(anchors("#hashtag", "####### seven", "    # indented code"), [])


class MarkdownLinkAnchors(GateTestCase):
    """The markdown-links check against a document using these anchors."""

    check = "markdown links"
    DOC = "docs/ANCHORS.md"

    def write(self, *lines: str) -> None:
        self.repo.write(self.DOC, "\n".join(lines) + "\n")

    def test_passes_with_links_that_work_on_github(self) -> None:
        self.write("# Anchors", "", "## `UI_SetExcelUI` reference", "## Café — notes",
                   "## Notes", "## Notes", "## 🧭 Release profile", "",
                   "[a](#ui_setexcelui-reference) [b](#café--notes) [c](#caf%C3%A9--notes)",
                   "[d](#notes) [e](#notes-1) [f](#-release-profile)")
        self.assertPasses()

    def test_fails_on_an_anchor_computed_the_old_way(self) -> None:
        self.write("## `UI_SetExcelUI` reference", "", "[a](#uisetexcelui-reference)")
        self.assertFails(f"{self.DOC}: broken internal anchor #uisetexcelui-reference")

    def test_fails_on_a_duplicate_suffix_that_does_not_exist(self) -> None:
        self.write("## Notes", "## Notes", "", "[a](#notes-2)")
        self.assertFails(f"{self.DOC}: broken internal anchor #notes-2")

    def test_fails_on_a_heading_that_only_exists_in_a_code_block(self) -> None:
        self.write("```bash", "# Install", "```", "", "[a](#install)")
        self.assertFails(f"{self.DOC}: broken internal anchor #install")


if __name__ == "__main__":
    unittest.main()
