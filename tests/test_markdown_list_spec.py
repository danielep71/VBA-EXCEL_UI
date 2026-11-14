"""List item boundaries checked against CommonMark 0.31.2 (#150).

The Markdown scanner reads fences and comment blocks inside list items, so it
must close each item where CommonMark closes it. Each case lists, for every
non-blank line, its kind and the open list items, named by the line their
marker is on: "code in 1,3" is a line inside a fenced block, within the item
opened on line 1 and the item opened on line 3 inside it.

SPEC_LIST_EXAMPLES holds every example of sections 5.2 and 5.3, from 253 to
325, except 259, 260, 292 and 293, whose lists sit inside block quotes, which
the scanner does not track. SPEC_LINK_DEFINITION_EXAMPLES holds every example
of section 4.7, from 192 to 218. The expected structure is the one
commonmark.js 0.31.2, the specification's reference parser, reports with
source positions; cmark-gfm, which renders Markdown on GitHub, gives the same
HTML for every one of them.
"""

import unittest

from gate_fixture import check_repo


def structure(markdown: str) -> list[str]:
    found = []
    for _, line, kind, items in check_repo.markdown_line_blocks(markdown):
        if not line.strip():
            found.append("")
        else:
            found.append(kind + (" in " + ",".join(map(str, items)) if items else ""))
    return found


def anchors(*lines: str) -> list[str]:
    return check_repo.heading_anchors("\n".join(lines) + "\n")


SPEC_LINK_DEFINITION_EXAMPLES: dict[int, tuple[str, list[str]]] = {
    192: ('[foo]: /url "title"\n\n[foo]',
          ['text', '', 'text']),
    193: ("   [foo]: \n      /url  \n           'the title'  \n\n[foo]",
          ['text', 'text', 'text', '', 'text']),
    194: ("[Foo*bar\\]]:my_(url) 'title (with parens)'\n\n[Foo*bar\\]]",
          ['text', '', 'text']),
    195: ("[Foo bar]:\n<my url>\n'title'\n\n[Foo bar]",
          ['text', 'text', 'text', '', 'text']),
    196: ("[foo]: /url '\ntitle\nline1\nline2\n'\n\n[foo]",
          ['text', 'text', 'text', 'text', 'text', '', 'text']),
    197: ("[foo]: /url 'title\n\nwith blank line'\n\n[foo]",
          ['text', '', 'text', '', 'text']),
    198: ('[foo]:\n/url\n\n[foo]',
          ['text', 'text', '', 'text']),
    199: ('[foo]:\n\n[foo]',
          ['text', '', 'text']),
    200: ('[foo]: <>\n\n[foo]',
          ['text', '', 'text']),
    201: ('[foo]: <bar>(baz)\n\n[foo]',
          ['text', '', 'text']),
    202: ('[foo]: /url\\bar\\*baz "foo\\"bar\\baz"\n\n[foo]',
          ['text', '', 'text']),
    203: ('[foo]\n\n[foo]: url',
          ['text', '', 'text']),
    204: ('[foo]\n\n[foo]: first\n[foo]: second',
          ['text', '', 'text', 'text']),
    205: ('[FOO]: /url\n\n[Foo]',
          ['text', '', 'text']),
    206: ('[ΑΓΩ]: /φου\n\n[αγω]',
          ['text', '', 'text']),
    207: ('[foo]: /url',
          ['text']),
    208: ('[\nfoo\n]: /url\nbar',
          ['text', 'text', 'text', 'text']),
    209: ('[foo]: /url "title" ok',
          ['text']),
    210: ('[foo]: /url\n"title" ok',
          ['text', 'text']),
    211: ('    [foo]: /url "title"\n\n[foo]',
          ['text', '', 'text']),
    212: ('```\n[foo]: /url\n```\n\n[foo]',
          ['fence', 'code', 'fence', '', 'text']),
    213: ('Foo\n[bar]: /baz\n\n[bar]',
          ['text', 'text', '', 'text']),
    214: ('# [Foo]\n[foo]: /url\n> bar',
          ['text', 'text', 'text']),
    215: ('[foo]: /url\nbar\n===\n[foo]',
          ['text', 'text', 'text', 'text']),
    216: ('[foo]: /url\n===\n[foo]',
          ['text', 'text', 'text']),
    217: ('[foo]: /foo-url "foo"\n[bar]: /bar-url\n  "bar"\n[baz]: /baz-url\n\n[foo],\n[bar],\n[baz]',
          ['text', 'text', 'text', 'text', '', 'text', 'text', 'text']),
    218: ('[foo]\n\n> [foo]: /url',
          ['text', '', 'text']),
}


SPEC_LIST_EXAMPLES: dict[int, tuple[str, list[str]]] = {
    253: ('A paragraph\nwith two lines.\n\n    indented code\n\n> A block quote.',
          ['text', 'text', '', 'text', '', 'text']),
    254: ('1.  A paragraph\n    with two lines.\n\n        indented code\n\n    > A block quote.',
          ['text in 1', 'text in 1', '', 'text in 1', '', 'text in 1']),
    255: ('- one\n\n two',
          ['text in 1', '', 'text']),
    256: ('- one\n\n  two',
          ['text in 1', '', 'text in 1']),
    257: (' -    one\n\n     two',
          ['text in 1', '', 'text']),
    258: (' -    one\n\n      two',
          ['text in 1', '', 'text in 1']),
    261: ('-one\n\n2.two',
          ['text', '', 'text']),
    262: ('- foo\n\n\n  bar',
          ['text in 1', '', '', 'text in 1']),
    263: ('1.  foo\n\n    ```\n    bar\n    ```\n\n    baz\n\n    > bam',
          ['text in 1', '', 'fence in 1', 'code in 1', 'fence in 1', '', 'text in 1', '', 'text in 1']),
    264: ('- Foo\n\n      bar\n\n\n      baz',
          ['text in 1', '', 'text in 1', '', '', 'text in 1']),
    265: ('123456789. ok',
          ['text in 1']),
    266: ('1234567890. not ok',
          ['text']),
    267: ('0. ok',
          ['text in 1']),
    268: ('003. ok',
          ['text in 1']),
    269: ('-1. not ok',
          ['text']),
    270: ('- foo\n\n      bar',
          ['text in 1', '', 'text in 1']),
    271: ('  10.  foo\n\n           bar',
          ['text in 1', '', 'text in 1']),
    272: ('    indented code\n\nparagraph\n\n    more code',
          ['text', '', 'text', '', 'text']),
    273: ('1.     indented code\n\n   paragraph\n\n       more code',
          ['text in 1', '', 'text in 1', '', 'text in 1']),
    274: ('1.      indented code\n\n   paragraph\n\n       more code',
          ['text in 1', '', 'text in 1', '', 'text in 1']),
    275: ('   foo\n\nbar',
          ['text', '', 'text']),
    276: ('-    foo\n\n  bar',
          ['text in 1', '', 'text']),
    277: ('-  foo\n\n   bar',
          ['text in 1', '', 'text in 1']),
    278: ('-\n  foo\n-\n  ```\n  bar\n  ```\n-\n      baz',
          ['text in 1', 'text in 1', 'text in 3', 'fence in 3', 'code in 3', 'fence in 3', 'text in 7', 'text in 7']),
    279: ('-   \n  foo',
          ['text in 1', 'text in 1']),
    280: ('-\n\n  foo',
          ['text in 1', '', 'text']),
    281: ('- foo\n-\n- bar',
          ['text in 1', 'text in 2', 'text in 3']),
    282: ('- foo\n-   \n- bar',
          ['text in 1', 'text in 2', 'text in 3']),
    283: ('1. foo\n2.\n3. bar',
          ['text in 1', 'text in 2', 'text in 3']),
    284: ('*',
          ['text in 1']),
    285: ('foo\n*\n\nfoo\n1.',
          ['text', 'text', '', 'text', 'text']),
    286: (' 1.  A paragraph\n     with two lines.\n\n         indented code\n\n     > A block quote.',
          ['text in 1', 'text in 1', '', 'text in 1', '', 'text in 1']),
    287: ('  1.  A paragraph\n      with two lines.\n\n          indented code\n\n      > A block quote.',
          ['text in 1', 'text in 1', '', 'text in 1', '', 'text in 1']),
    288: ('   1.  A paragraph\n       with two lines.\n\n           indented code\n\n       > A block quote.',
          ['text in 1', 'text in 1', '', 'text in 1', '', 'text in 1']),
    289: ('    1.  A paragraph\n        with two lines.\n\n            indented code\n\n        > A block quote.',
          ['text', 'text', '', 'text', '', 'text']),
    290: ('  1.  A paragraph\nwith two lines.\n\n          indented code\n\n      > A block quote.',
          ['text in 1', 'text in 1', '', 'text in 1', '', 'text in 1']),
    291: ('  1.  A paragraph\n    with two lines.',
          ['text in 1', 'text in 1']),
    294: ('- foo\n  - bar\n    - baz\n      - boo',
          ['text in 1', 'text in 1,2', 'text in 1,2,3', 'text in 1,2,3,4']),
    295: ('- foo\n - bar\n  - baz\n   - boo',
          ['text in 1', 'text in 2', 'text in 3', 'text in 4']),
    296: ('10) foo\n    - bar',
          ['text in 1', 'text in 1,2']),
    297: ('10) foo\n   - bar',
          ['text in 1', 'text in 2']),
    298: ('- - foo',
          ['text in 1,1']),
    299: ('1. - 2. foo',
          ['text in 1,1,1']),
    300: ('- # Foo\n- Bar\n  ---\n  baz',
          ['text in 1', 'text in 2', 'text in 2', 'text in 2']),
    301: ('- foo\n- bar\n+ baz',
          ['text in 1', 'text in 2', 'text in 3']),
    302: ('1. foo\n2. bar\n3) baz',
          ['text in 1', 'text in 2', 'text in 3']),
    303: ('Foo\n- bar\n- baz',
          ['text', 'text in 2', 'text in 3']),
    304: ('The number of windows in my house is\n14.  The number of doors is 6.',
          ['text', 'text']),
    305: ('The number of windows in my house is\n1.  The number of doors is 6.',
          ['text', 'text in 2']),
    306: ('- foo\n\n- bar\n\n\n- baz',
          ['text in 1', '', 'text in 3', '', '', 'text in 6']),
    307: ('- foo\n  - bar\n    - baz\n\n\n      bim',
          ['text in 1', 'text in 1,2', 'text in 1,2,3', '', '', 'text in 1,2,3']),
    308: ('- foo\n- bar\n\n<!-- -->\n\n- baz\n- bim',
          ['text in 1', 'text in 2', '', 'comment', '', 'text in 6', 'text in 7']),
    309: ('-   foo\n\n    notcode\n\n-   foo\n\n<!-- -->\n\n    code',
          ['text in 1', '', 'text in 1', '', 'text in 5', '', 'comment', '', 'text']),
    310: ('- a\n - b\n  - c\n   - d\n  - e\n - f\n- g',
          ['text in 1', 'text in 2', 'text in 3', 'text in 4', 'text in 5', 'text in 6', 'text in 7']),
    311: ('1. a\n\n  2. b\n\n   3. c',
          ['text in 1', '', 'text in 3', '', 'text in 5']),
    312: ('- a\n - b\n  - c\n   - d\n    - e',
          ['text in 1', 'text in 2', 'text in 3', 'text in 4', 'text in 4']),
    313: ('1. a\n\n  2. b\n\n    3. c',
          ['text in 1', '', 'text in 3', '', 'text']),
    314: ('- a\n- b\n\n- c',
          ['text in 1', 'text in 2', '', 'text in 4']),
    315: ('* a\n*\n\n* c',
          ['text in 1', 'text in 2', '', 'text in 4']),
    316: ('- a\n- b\n\n  c\n- d',
          ['text in 1', 'text in 2', '', 'text in 2', 'text in 5']),
    317: ('- a\n- b\n\n  [ref]: /url\n- d',
          ['text in 1', 'text in 2', '', 'text in 2', 'text in 5']),
    318: ('- a\n- ```\n  b\n\n\n  ```\n- c',
          ['text in 1', 'fence in 2', 'code in 2', '', '', 'fence in 2', 'text in 7']),
    319: ('- a\n  - b\n\n    c\n- d',
          ['text in 1', 'text in 1,2', '', 'text in 1,2', 'text in 5']),
    320: ('* a\n  > b\n  >\n* c',
          ['text in 1', 'text in 1', 'text in 1', 'text in 4']),
    321: ('- a\n  > b\n  ```\n  c\n  ```\n- d',
          ['text in 1', 'text in 1', 'fence in 1', 'code in 1', 'fence in 1', 'text in 6']),
    322: ('- a',
          ['text in 1']),
    323: ('- a\n  - b',
          ['text in 1', 'text in 1,2']),
    324: ('1. ```\n   foo\n   ```\n\n   bar',
          ['fence in 1', 'code in 1', 'fence in 1', '', 'text in 1']),
    325: ('* foo\n  * bar\n\n  baz',
          ['text in 1', 'text in 1,2', '', 'text in 1']),
}


class SpecListExamples(unittest.TestCase):

    def test_each_example_closes_its_items_where_the_spec_does(self) -> None:
        for number, (markdown, expected) in SPEC_LIST_EXAMPLES.items():
            with self.subTest(example=number):
                self.assertEqual(structure(markdown), expected)


class SpecLinkDefinitionExamples(unittest.TestCase):

    def test_each_example_reads_as_the_spec_reads_it(self) -> None:
        for number, (markdown, expected) in SPEC_LINK_DEFINITION_EXAMPLES.items():
            with self.subTest(example=number):
                self.assertEqual(structure(markdown), expected)


# In each case below, the line after the first item ends it, so the two-space
# fence that follows is at the top level and hides its heading. Were the
# first item still open, the fence would take its column, the unindented
# heading would end the fence early, and its closing fence would open a block
# that hid "After".
TOP_LEVEL_FENCE = ("", "  ```text", "## Hidden", "  ```", "## After")


class ListBoundaries(unittest.TestCase):

    def test_a_different_list_type_starts_a_new_list_whatever_its_number(self) -> None:
        self.assertEqual(anchors("- Item", "2. Two", *TOP_LEVEL_FENCE), ["after"])
        self.assertEqual(structure("- Item\n2. Two"), ["text in 1", "text in 2"])

    def test_a_different_delimiter_starts_a_new_list_whatever_its_number(self) -> None:
        self.assertEqual(anchors("1. Item", "3) Three", *TOP_LEVEL_FENCE), ["after"])
        self.assertEqual(structure("1. Item\n3) Three"), ["text in 1", "text in 2"])

    def test_a_different_bullet_starts_a_new_list_even_when_empty(self) -> None:
        self.assertEqual(anchors("- Item", "+", *TOP_LEVEL_FENCE), ["after"])
        self.assertEqual(structure("- Item\n+"), ["text in 1", "text in 2"])

    def test_a_new_list_inside_an_outer_item_follows_the_same_rule(self) -> None:
        nested = ("- Outer", "  - Inner", "  2. Two", "", "    ```text", "  ## Hidden", "    ```", "## After")
        self.assertEqual(anchors(*nested), ["after"])
        self.assertEqual(structure("\n".join(nested)),
                         ["text in 1", "text in 1,2", "text in 1,3", "", "fence in 1", "code in 1",
                          "fence in 1", "text"])

    def test_the_start_number_rule_still_applies_inside_one_paragraph(self) -> None:
        self.assertEqual(structure("Para\n2. Two"), ["text", "text"])
        self.assertEqual(structure("- Para\n  2. Two"), ["text in 1", "text in 1"])
        self.assertEqual(structure("- Para\n  1. One"), ["text in 1", "text in 1,2"])

    def test_a_lazy_line_keeps_its_items(self) -> None:
        self.assertEqual(structure("- a\n  - b\nlazy"), ["text in 1", "text in 1,2", "text in 1,2"])

    def test_a_blank_line_then_less_indented_text_ends_the_item(self) -> None:
        self.assertEqual(structure("- a\n\n b\n\n  c"), ["text in 1", "", "text", "", "text"])

    def test_a_comment_block_ends_with_its_item(self) -> None:
        self.assertEqual(anchors("- <!--", "## Visible", "-->"), ["visible"])
        self.assertEqual(structure("- <!--\n  ## Hidden\n## Visible\n-->"),
                         ["comment in 1", "comment in 1", "text", "text"])


class LinkDefinitions(unittest.TestCase):
    """A paragraph of only link reference definitions, read as cmark-gfm reads it."""

    def test_an_underline_under_definitions_is_paragraph_text(self) -> None:
        # No text is left to make a heading of, so the paragraph stays open
        # and "lazy" continues it inside the item.
        for underline in ("  ===", "  ---"):
            with self.subTest(underline=underline):
                self.assertEqual(structure(f"- [a]: /u\n{underline}\nlazy"),
                                 ["text in 1", "text in 1", "text in 1"])

    def test_definitions_over_several_lines_or_with_tabs_are_recognised(self) -> None:
        self.assertEqual(structure('- [a]:\n  /u\n  "t"\n  ===\nlazy'), ["text in 1"] * 5)
        self.assertEqual(structure("- [a]:\t/u(x\n  ===\nlazy"), ["text in 1"] * 3)

    def test_text_after_a_definition_still_makes_a_heading(self) -> None:
        self.assertEqual(structure("- [a]: /u 't'x\n  ===\nlazy"), ["text in 1", "text in 1", "text"])

    def test_an_item_of_only_definitions_ends_at_its_second_blank_line(self) -> None:
        self.assertEqual(structure("- [a]: /u\n\n  x"), ["text in 1", "", "text in 1"])
        self.assertEqual(structure("- [a]: /u\n\n\n  x"), ["text in 1", "", "", "text"])

    def test_a_lazy_definition_keeps_its_indentation(self) -> None:
        # The indented lazy line is text, not a definition, so the item has
        # a block and continues across the blank lines.
        self.assertEqual(structure("1. [a]: /u\n [b]: /v\n\n\n   y"),
                         ["text in 1", "text in 1", "", "", "text in 1"])

    def test_a_definition_line_still_counts_as_a_paragraph_for_a_marker(self) -> None:
        # Both parsers read "2. Step" as paragraph text, so the fence after it
        # stays inside the bullet item and hides its heading.
        self.assertEqual(anchors("- [ref]: /url", "  2. Step", "", "     ```text", "  ## Visible",
                                 "     ```", "## After"), ["after"])


class BlankLines(unittest.TestCase):

    def test_spaces_up_to_the_content_column_continue_an_empty_item(self) -> None:
        self.assertEqual(structure("2.\n   \n   x"), ["text in 1", "", "text in 1"])
        self.assertEqual(structure("2.\n\n   x"), ["text in 1", "", "text"])


class OutOfScope(unittest.TestCase):
    """The documented limits of the scanner, pinned so a change is deliberate."""

    def test_a_block_quote_is_text_and_hides_no_fence(self) -> None:
        self.assertEqual(structure("> ```\n> ## Hidden\n> ```\n## After"), ["text"] * 4)
        self.assertEqual(structure("> - a\n> - b"), ["text", "text"])

    def test_a_block_quote_line_ends_a_paragraph(self) -> None:
        self.assertEqual(structure("- a\n> b"), ["text in 1", "text"])

    def test_html_blocks_other_than_comments_are_not_tracked(self) -> None:
        self.assertEqual(structure("<div>\n```\n## Hidden\n```\n## After"),
                         ["text", "fence", "code", "fence", "text"])

    def test_a_heading_inside_a_list_item_has_no_anchor(self) -> None:
        self.assertEqual(anchors("- ## In item"), [])


if __name__ == "__main__":
    unittest.main()
