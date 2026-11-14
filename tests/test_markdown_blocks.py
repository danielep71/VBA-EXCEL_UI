"""Fenced code blocks inside list items (#150).

The Markdown scanner behind the markdown-links, markdown-escapes and
changelog checks reads fences inside their list items: a fence may follow a
list marker, or be indented up to three columns past the item's content
column. Read at the top level instead, an opener after a marker is missed, its
closing fence opens a block that hides every heading after it, and a fence
indented four spaces under an ordered item is read as prose.
"""

import unittest

from gate_fixture import GateTestCase, check_repo


def doc(*lines: str) -> str:
    return "\n".join(lines) + "\n"


def kinds(*lines: str) -> list[str]:
    return [kind for _, _, kind in check_repo.markdown_line_kinds("\n".join(lines))]


def anchors(*lines: str) -> list[str]:
    return check_repo.heading_anchors(doc(*lines))


def escapes(*lines: str) -> list[str]:
    return check_repo.escape_findings("fixture.md", doc(*lines))


# The example in #150: a fence indented to its bullet item's content.
LIST_FENCE = ("- Example:", "  ```text", "  ## Hidden", "  ```", "", "## Install")

# #96: the opener follows the list marker on the same line.
MARKER_FENCE = ("- ```text", "  ## Hidden", "  ```", "", "## Install")

# #103: a fence indented four spaces under an ordered item, after a blank line.
ORDERED_FENCE = ("1. Create the scratch folder.", "", "    ```text", "    C:\\_scratch\\file", "    ```")


class ListContainedFences(unittest.TestCase):

    def test_a_fence_indented_to_the_item_content_hides_its_heading(self) -> None:
        self.assertEqual(anchors(*LIST_FENCE), ["install"])
        self.assertEqual(kinds(*LIST_FENCE), ["text", "fence", "code", "fence", "text", "text"])

    def test_a_fence_after_the_list_marker_opens_a_block(self) -> None:
        self.assertEqual(anchors(*MARKER_FENCE), ["install"])
        self.assertEqual(kinds(*MARKER_FENCE), ["fence", "code", "fence", "text", "text"])

    def test_a_fence_in_an_ordered_item_continuation_is_code(self) -> None:
        self.assertEqual(kinds(*ORDERED_FENCE), ["text", "text", "fence", "code", "fence"])
        self.assertEqual(escapes(*ORDERED_FENCE), [])

    def test_escapes_in_prose_after_the_list_are_still_reported(self) -> None:
        self.assertEqual(len(escapes(*ORDERED_FENCE, "", "The M\\_EXCEL\\_UI module.")), 1)
        self.assertEqual(len(escapes(*LIST_FENCE, "", "The M\\_EXCEL\\_UI module.")), 1)

    def test_a_fence_in_a_nested_item_closes_at_its_own_column(self) -> None:
        nested = ("- Outer", "  - Inner", "    ```text", "    ## Hidden", "    C:\\_x", "    ```",
                  "  - Next", "", "## After")
        self.assertEqual(anchors(*nested), ["after"])
        self.assertEqual(escapes(*nested), [])
        self.assertEqual(kinds(*nested),
                         ["text", "text", "fence", "code", "code", "fence", "text", "text", "text"])

    def test_a_fence_indented_three_past_the_content_column_still_opens(self) -> None:
        self.assertEqual(anchors("- Item", "     ```", "     ## Hidden", "     ```", "## After"), ["after"])
        self.assertEqual(anchors("1. Item", "      ```", "      ## Hidden", "      ```", "## After"),
                         ["after"])

    def test_four_columns_past_the_content_column_is_not_a_fence(self) -> None:
        self.assertEqual(kinds("- Item", "", "      ```", "", "## After"),
                         ["text", "text", "text", "text", "text"])

    def test_an_unindented_line_ends_the_item_and_its_fence(self) -> None:
        self.assertEqual(kinds("- ```", "  code", "Outside", "## After"), ["fence", "code", "text", "text"])
        self.assertEqual(anchors("- ```", "  code", "Outside", "## After"), ["after"])

    def test_a_lazy_paragraph_line_keeps_its_item(self) -> None:
        self.assertEqual(kinds("1. Step", "lazy line", "    ```", "    C:\\_x", "    ```"),
                         ["text", "text", "fence", "code", "fence"])

    def test_a_number_other_than_one_does_not_interrupt_a_paragraph(self) -> None:
        self.assertEqual(kinds("Since", "2026. A year", "    ```", "", "## After"),
                         ["text", "text", "text", "text", "text"])
        self.assertEqual(kinds("1. One", "2. Two", "   ```", "   x", "   ```"),
                         ["text", "text", "fence", "code", "fence"])

    def test_a_thematic_break_is_not_a_list_item(self) -> None:
        self.assertEqual(kinds("* * *", "```", "## Hidden", "```", "## After"),
                         ["text", "fence", "code", "fence", "text"])

    def test_a_comment_block_inside_an_item_hides_nothing_after_it(self) -> None:
        self.assertEqual(kinds("- <!-- note -->", "## After"), ["comment", "text"])


class TopLevelBlocks(unittest.TestCase):

    def test_a_top_level_fence_still_hides_its_content(self) -> None:
        self.assertEqual(anchors("```text", "## Hidden", "C:\\_x", "```", "## After"), ["after"])
        self.assertEqual(escapes("```text", "C:\\_scratch\\file", "```"), [])

    def test_indented_code_is_not_a_fence(self) -> None:
        self.assertEqual(kinds("Text", "", "    ```", "    ## Not a heading", "", "## Install"),
                         ["text", "text", "text", "text", "text", "text"])
        self.assertEqual(anchors("Text", "", "    ```", "", "## Install"), ["install"])

    def test_code_segments_read_fences_inside_list_items(self) -> None:
        self.assertEqual(list(check_repo.code_segments(doc("- ```bash", "  python3 tools/verify.py", "  ```",
                                                             "", "Run `x`."))),
                         [(2, "  python3 tools/verify.py"), (5, "x")])


class MarkdownGateOnListFences(GateTestCase):
    """The markdown-links check against the list example in #150."""

    check = "markdown links"
    DOC = "docs/LIST_FENCES.md"

    def test_a_heading_after_a_list_fence_can_be_linked(self) -> None:
        self.repo.write(self.DOC, doc("# List fences", *LIST_FENCE, "", *MARKER_FENCE, "",
                                      "[a](#install) [b](#install-1)"))
        self.assertPasses()

    def test_a_heading_inside_a_list_fence_cannot_be_linked(self) -> None:
        self.repo.write(self.DOC, doc("# List fences", *LIST_FENCE, "", "[a](#hidden)"))
        self.assertFails(f"{self.DOC}: broken internal anchor #hidden")


class MarkdownEscapeGateOnListFences(GateTestCase):
    """The markdown-escapes check against the ordered-list example in #103."""

    check = "markdown escapes"
    DOC = "docs/RIBBON_SDI_BEHAVIOR.md"

    def test_a_path_in_an_ordered_list_fence_passes(self) -> None:
        self.repo.append(self.DOC, doc("", *ORDERED_FENCE))
        self.assertPasses()

    def test_an_escape_after_the_list_fails(self) -> None:
        self.repo.append(self.DOC, doc("", *ORDERED_FENCE, "", "The M\\_EXCEL\\_UI module."))
        self.assertFails(f"{self.DOC}:", "escaped Markdown outside code: \\_, \\_")


if __name__ == "__main__":
    unittest.main()
