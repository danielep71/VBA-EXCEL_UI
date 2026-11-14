"""List item boundaries checked against CommonMark 0.31.2, sections 5.2 and 5.3 (#150).

The Markdown scanner reads fences and comment blocks inside list items, so it
must close each item where CommonMark closes it. Each case lists, for every
non-blank line, its kind and the open list items, named by the line their
marker is on: "code in 1,3" is a line inside a fenced block, within the item
opened on line 1 and the item opened on line 3 inside it.

SPEC_LIST_EXAMPLES holds every example from 253 to 325 except 259, 260, 292
and 293, whose lists sit inside block quotes, which the scanner does not
track. The expected structure is that of the spec's HTML, as the CommonMark
reference parser reports it with source positions.
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


if __name__ == "__main__":
    unittest.main()
