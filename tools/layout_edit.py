"""Find and replace whole lines of VBA source without depending on its layout.

The fixtures that plant a defect in a real module used to name the line to
change with its exact indentation and column alignment. A layout change then
broke them, and not because the rule they prove had changed. These helpers
match lines by their words (indentation and runs of spaces do not count) and
give each replacement line the indentation of the line it replaces, so the same
fixture holds for the module in either layout.
"""
import re


def words(line: str) -> str:
    """A line with its indentation dropped and every run of spaces made one."""
    return " ".join(line.split())


def locate(text: str, old: str) -> list[tuple[int, int]]:
    """Character spans of each place where the lines of ``old`` stand, in order, in ``text``."""
    wanted = [words(line) for line in old.splitlines()]
    lines = text.splitlines(keepends=True)
    starts = [0]
    for line in lines:
        starts.append(starts[-1] + len(line))
    found: list[tuple[int, int]] = []
    n = 0
    while n + len(wanted) <= len(lines):
        if [words(line) for line in lines[n:n + len(wanted)]] == wanted:
            found.append((starts[n], starts[n + len(wanted)]))
            n += len(wanted)
        else:
            n += 1
    return found


def replace(text: str, old: str, new: str, *, unique: bool = False) -> str | None:
    """Replace the lines of ``old`` by the lines of ``new``; None when ``old`` is not there.

    ``old`` is looked for as written first. Failing that, its lines are matched
    by their words, and each replacement line is indented like the line of
    ``old`` it stands for (or, past its end, like the last one), moved by the
    difference between the indentation ``new`` and ``old`` give those lines.
    """
    exact = text.count(old)
    if exact:
        return None if unique and exact != 1 else text.replace(old, new)
    spans = locate(text, old)
    if not spans or (unique and len(spans) != 1):
        return None
    old_lines = old.splitlines()
    new_lines = new.splitlines()
    for start, stop in reversed(spans):
        block = text[start:stop].splitlines(keepends=True)
        eol = re.search(r"\r?\n$", block[-1])
        ending = eol.group(0) if eol else "\r\n"
        out = []
        for k, line in enumerate(new_lines):
            at = min(k, len(old_lines) - 1)
            lead = len(block[at]) - len(block[at].lstrip(" "))
            written = len(old_lines[at]) - len(old_lines[at].lstrip(" "))
            wanted = len(line) - len(line.lstrip(" "))
            out.append(" " * max(0, lead + wanted - written) + line.lstrip(" ") + ending)
        text = text[:start] + "".join(out) + text[stop:]
    return text
