"""Rewrite one VBA standard module into the house-style normal form.

The normal form is described in docs/VBA_HOUSE_STYLE.md. This module only
changes layout, comments, header text and the names that the style renames
(error-handling labels, the PROC constant). Everything else is reported, not
rewritten: ``format_module`` returns the findings a person has to settle.

The scanner of vba_lex.py protects string literals, comments and bracketed
identifiers. Conditional-compilation blocks are copied as they are. A
procedure whose blocks the indenter cannot pair up keeps its indentation and
is reported. The caller (reformat.py) compares the ordered executable tokens
before and after, so a defect here cannot change what the code does.
"""
import re
from dataclasses import dataclass
from datetime import date
from typing import NamedTuple

import house_style as hs
from vba_lex import LexicalError, code_only, protected_lines, split_code_comment, transform_code

Todo = tuple[str, str]

SIG_RE = re.compile(
    r"^(?P<vis>Public |Private |Friend )?(?P<static>Static )?"
    r"(?P<kind>Sub|Function|Property (?:Get|Let|Set))\s+(?P<name>[A-Za-z_]\w*)")
END_RE = re.compile(r"^End (?:Sub|Function|Property)\b")
RULE_RE = re.compile(r"^'[-=]{5,}$")
FIELD_LABEL_RE = re.compile(r"^' [A-Z][A-Z0-9 /&-]*$")
DATE_LINE_RE = re.compile(r"^'\s+(\d{4}-\d{2}-\d{2})\b")
LABEL_LINE_RE = re.compile(r"^([A-Za-z_]\w*):$")
ON_ERROR_RE = re.compile(r"^On Error GoTo (?!0\b)([A-Za-z_]\w*)$", re.I)
EXIT_RE = re.compile(r"^Exit (?:Sub|Function|Property)$", re.I)
DECL_RE = re.compile(
    r"^(?P<ind>\s*)(?P<head>(?:(?:Public|Private|Global|Friend)\s+(?:Const\s+)?"
    r"|(?:Dim|Static)\s+|Const\s+)(?:WithEvents\s+)?[A-Za-z_]\w*(?:\([^)]*\))?)"
    r"\s+As\s+(?P<rest>\S.*?)\s*$")
BLOCK_TYPE_RE = re.compile(r"^(?:Public |Private |Friend )?(?:Enum|Type)\s+\w+", re.I)
END_BLOCK_TYPE_RE = re.compile(r"^End (?:Enum|Type)\b", re.I)
SEPARATOR_RE = re.compile(r":(?!=)")

HANDLER_LABELS = frozenset({hs.HOUSE_LABELS.handler, hs.HOUSE_LABELS.clean_fail})
EXIT_LABELS = frozenset({hs.HOUSE_LABELS.safe_exit, hs.HOUSE_LABELS.clean_exit})


@dataclass
class Ln:
    """One physical line. ``fixed`` lines sit in a conditional block and are copied."""
    text: str
    fixed: bool = False


class Bail(Exception):
    """The indenter cannot classify the procedure; it is left as written."""


# --------------------------------------------------------------------------
# Small helpers


def split_lines(text: str) -> list[str]:
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def is_rule(text: str) -> bool:
    return RULE_RE.match(text) is not None


def is_blank(line: Ln) -> bool:
    return not line.text.strip()


def is_lone(line: Ln) -> bool:
    return line.text == "'"


def indent_of(text: str) -> int:
    return len(text) - len(text.lstrip(" "))


def code_of(text: str) -> str:
    return split_code_comment(text)[0].strip()


def continues(text: str) -> bool:
    code = split_code_comment(text)[0].rstrip()
    return code == "_" or code.endswith(" _")


def round_up(value: int, step: int = 4) -> int:
    return -(-value // step) * step


def conditional_lines(lines: list[str]) -> set[int]:
    """Lines of #If blocks and every conditional directive line."""
    out: set[int] = set()
    depth = 0
    for n, line in enumerate(lines):
        code = split_code_comment(line)[0].strip()
        directive = re.match(r"#(?:If|ElseIf|Else|End|Const)\b", line.lstrip(), re.I) is not None
        if re.match(r"#If\b", code, re.I):
            depth += 1
        if depth or directive:
            out.add(n)
        if re.match(r"#End\s+If\b", code, re.I):
            depth -= 1
    return out


def spans(lines: list[Ln], start: int, stop: int) -> list[tuple[int, int]]:
    """Logical lines of lines[start:stop]: a statement and its continuation lines."""
    out: list[tuple[int, int]] = []
    n = start
    while n < stop:
        m = n
        while m + 1 < stop and not lines[m].fixed and continues(lines[m].text):
            m += 1
        out.append((n, m))
        n = m + 1
    return out


def logical_code(lines: list[Ln], first: int, last: int) -> str:
    """The code of a logical line with literals and comments masked, one space joined."""
    parts = [code_only(lines[k].text).strip() for k in range(first, last + 1)]
    parts = [p[:-1].rstrip() if k < len(parts) - 1 and p.endswith("_") else p
             for k, p in enumerate(parts)]
    return " ".join(p for p in parts if p).strip()


def ensure_blanks(out: list[Ln], count: int) -> None:
    """Leave exactly ``count`` blank lines at the end of ``out``."""
    while out and is_blank(out[-1]) and not out[-1].fixed:
        out.pop()
    if out:
        out.extend(Ln("") for _ in range(count))


# --------------------------------------------------------------------------
# Renames


def rename_label_text(text: str, label_map: dict[str, str]) -> str:
    """Rewrite GoTo/Resume targets in one run of code or comment prose."""
    for old, rep in label_map.items():
        def renamed(m: re.Match[str], r: str = rep) -> str:
            return m.group(1) + " " + r

        text = re.sub(r"\b(GoTo|Resume)\s+" + old + r"\b", renamed, text, flags=re.I)
    return text


def rename_labels(lines: list[str], label_map: dict[str, str]) -> list[str]:
    """Rename label definitions and jumps, in code and in the prose describing them.

    Never inside a string literal: a diagnostic that quotes a label name is
    data the module prints.
    """
    out: list[str] = []
    for ln in lines:
        m = re.match(r"^([A-Za-z_]\w*):(?!=)", ln.strip())
        if m and m.group(1) in label_map:
            ln = ln.replace(m.group(1) + ":", label_map[m.group(1)] + ":", 1)
        code, comment = split_code_comment(ln)
        renamed = transform_code(code, lambda part: rename_label_text(part, label_map))
        if not comment.lower().startswith("rem"):
            comment = rename_label_text(comment, label_map)
        out.append(renamed + comment)
    return out


CONST_PROC_RE = re.compile(r"\s*Const\s+" + hs.PROC_CONSTANT_OLD + r"\s+As\b", re.I)


def standalone(code: str, word: str) -> list[re.Match[str]]:
    """The places where ``word`` stands as a name of its own in masked code, in any letter case.

    VBA reads names without regard to case, so ``proc_name`` is ``PROC_NAME``.
    A name after a dot or a bang (``obj.PROC``, ``row!PROC``, a ``With`` line's
    ``.PROC``) is a member and a name before ``:=`` is an argument label;
    neither refers to a variable, so neither counts. ``code`` should be a whole
    logical statement, so that the dot at the end of one physical line is seen
    by the name that starts the next.
    """
    found = []
    for m in re.finditer(r"(?<!\w)" + re.escape(word) + r"(?!\w)", code, re.I):
        before = code[:m.start()].rstrip()
        if before.endswith((".", "!")) or code[m.end():].lstrip().startswith(":="):
            continue
        found.append(m)
    return found


class Logical(NamedTuple):
    """One logical statement, masked: continuation lines joined, same offsets as the lines."""
    text: str
    starts: list[tuple[int, int]]       # (physical line, offset of that line in text)

    def where(self, offset: int) -> tuple[int, int]:
        """The physical line and column of an offset of ``text``."""
        line, start = [s for s in self.starts if s[1] <= offset][-1]
        return line, offset - start


def logical_masked(raw: list[str]) -> list[Logical]:
    """Every logical statement of the module with literals and comments masked."""
    out: list[Logical] = []
    n = 0
    while n < len(raw):
        text, starts = "", []
        while True:
            piece = code_only(raw[n])
            tail = piece.rstrip()
            wraps = (tail.endswith("_") and (len(tail) == 1 or tail[-2].isspace())) and n + 1 < len(raw)
            if wraps:
                piece = piece[:len(tail) - 1] + " " + piece[len(tail):]
            starts.append((n, len(text)))
            text += piece + " "
            n += 1
            if not wraps:
                break
        out.append(Logical(text, starts))
    return out


def standalone_at(logicals: list[Logical], word: str) -> list[tuple[int, int]]:
    """(line, column) of every place ``word`` stands as a name of its own."""
    return [lg.where(m.start()) for lg in logicals for m in standalone(lg.text, word)]


def proc_constant_scopes(raw: list[str], conditional: set[int]) -> list[tuple[int, int]]:
    """(signature, End) of each procedure that declares its own ``Const PROC As String``."""
    lines = [Ln(t.rstrip(), n in conditional) for n, t in enumerate(raw)]
    return [(sig, end) for sig, end in procedure_spans(lines)
            if any(CONST_PROC_RE.match(code_only(raw[n])) for n in range(sig, end + 1))]


def migrated_lines(raw: list[str], conditional: set[int], declaring: set[int]) -> set[int]:
    """Lines of procedures that already declare ``Const PROC_NAME`` and not ``Const PROC``."""
    lines = [Ln(t.rstrip(), n in conditional) for n, t in enumerate(raw)]
    done: set[int] = set()
    for sig, end in procedure_spans(lines):
        span = range(sig, end + 1)
        if not declaring.intersection(span) and any(
                re.match(r"\s*Const\s+" + hs.PROC_CONSTANT + r"\s+As\b", code_only(raw[n]), re.I) for n in span):
            done.update(span)
    return done


def rename_proc_constant(raw: list[str], conditional: set[int], module_name: str) -> list[str]:
    """Rename PROC to PROC_NAME inside the procedures that declare ``Const PROC``, and only there.

    A procedure's constant is local to it, so another procedure, variable,
    parameter, Declare or type member that is also called PROC is a different
    thing; where one exists outside those procedures the module is refused,
    with its line, rather than renamed along with the constant. The same goes
    for PROC_NAME, in any letter case: the new name would be shadowed by, or
    would shadow, whatever already carries it. What a name is, member, argument
    label or variable, is decided on the whole logical statement.
    """
    scopes = proc_constant_scopes(raw, conditional)
    if not scopes:
        return raw
    inside = {n for sig, end in scopes for n in range(sig, end + 1)}
    migrated = migrated_lines(raw, conditional, inside)
    logicals = logical_masked(raw)
    for line, _col in standalone_at(logicals, hs.PROC_CONSTANT_OLD):
        if line not in inside:
            raise LexicalError(
                f"{module_name} line {line + 1}: {hs.PROC_CONSTANT_OLD} names something other than the "
                f"constant of its own procedure; rename it before moving the constant to "
                f"{hs.PROC_CONSTANT}")
    for line, _col in standalone_at(logicals, hs.PROC_CONSTANT):
        if line not in migrated:
            raise LexicalError(
                f"{module_name} line {line + 1}: {hs.PROC_CONSTANT} (in any letter case) already names "
                f"something that the constant of a procedure would shadow or be shadowed by")
    out = list(raw)
    for line, col in sorted(standalone_at(logicals, hs.PROC_CONSTANT_OLD), reverse=True):
        width = len(hs.PROC_CONSTANT_OLD)
        out[line] = out[line][:col] + hs.PROC_CONSTANT + out[line][col + width:]
    return out


def renamed_source(text: str, module_name: str) -> str:
    """The module as the declared renames alone would leave it: what the formatter's output must equal in tokens."""
    raw = split_lines(text)
    conditional = conditional_lines(raw)
    raw = rename_labels(raw, dict(hs.LABEL_RENAMES))
    return "\n".join(rename_proc_constant(raw, conditional, module_name))


def check_renames(lines: list[str], label_map: dict[str, str], protected: set[int], module_name: str) -> None:
    """Refuse a rename that is ambiguous, collides with a name or crosses text the formatter does not edit.

    Collisions are read without regard to letter case, as VBA reads them: a
    variable called ``errorhandler`` collides with the label ErrorHandler, and
    a label spelled ``SAFEEXIT`` beside ``Safe_Exit`` is two spellings of one.
    """
    masked = [code_only(ln) for ln in lines]
    logicals = logical_masked(lines)
    targets: dict[str, set[str]] = {}
    for old, new in label_map.items():
        targets.setdefault(new, {new}).add(old)
    for new, spellings in targets.items():
        found: dict[str, int] = {}
        for lg in logicals:
            for spelling in spellings:
                for m in standalone(lg.text, spelling):
                    found.setdefault(m.group(0), lg.where(m.start())[0])
        odd = sorted(s for s in found if s not in spellings)
        if odd:
            raise LexicalError(f"{module_name} line {found[odd[0]] + 1}: {odd[0]} is {new} in another letter case; "
                               f"spell it as {new} or as its previous name before formatting")
        present = sorted(found)
        if len(present) > 1 and any(s != new for s in present):
            raise LexicalError(f"{module_name} line {found[present[-1]] + 1}: ambiguous label mapping "
                               f"{'/'.join(present)}")
        for lg in logicals:
            for m in standalone(lg.text, new):
                if not is_label_use(lg.text, m):
                    raise LexicalError(f"{module_name} line {lg.where(m.start())[0] + 1}: {new} already names "
                                       f"something that is not a label; the label would collide with it")
    for old in label_map:
        for i in protected:
            if standalone(masked[i], old):
                raise LexicalError(f"{module_name} line {i + 1}: unsupported label {old} in preserved region")


def is_label_use(code: str, m: re.Match[str]) -> bool:
    """Whether the name at ``m`` is a label definition or the target of GoTo or Resume."""
    before = code[:m.start()].rstrip()
    if re.search(r"(?:\bGoTo|\bResume)$", before, re.I):
        return True
    return before == "" and code[m.end():].lstrip().startswith(":") and not code[m.end():].lstrip().startswith(":=")


# --------------------------------------------------------------------------
# Header fields


class Fields(NamedTuple):
    ordered: list[tuple[str, list[str]]]
    todos: list[str]


def trim_lone(body: list[str]) -> list[str]:
    start, stop = 0, len(body)
    while start < stop and body[start] == "'":
        start += 1
    while stop > start and body[stop - 1] == "'":
        stop -= 1
    return body[start:stop]


def parse_fields(lines: list[str]) -> list[tuple[str, list[str]]] | None:
    """Split the lines between the title rule and the closing rule into fields."""
    out: list[tuple[str, list[str]]] = []
    for line in lines:
        if FIELD_LABEL_RE.match(line) and not is_rule(line):
            out.append((line[2:].strip(), []))
        elif out:
            out[-1][1].append(line)
        elif line != "'":
            return None
    return [(label, trim_lone(body)) for label, body in out]


def text_of(body: list[str]) -> str:
    return " ".join(line[1:].strip() for line in body if line != "'")


def valid_date(value: str) -> bool:
    try:
        date.fromisoformat(value)
    except ValueError:
        return False
    return True


def convert_updated(body: list[str]) -> tuple[list[str], list[str], str | None]:
    """Split an UPDATED body into the latest date and the history it came from."""
    texts = [line for line in body if line != "'"]
    if len(texts) == 1 and re.fullmatch(r"'   \d{4}-\d{2}-\d{2}", texts[0]) and valid_date(texts[0][4:]):
        return texts, [], None
    dates = [m.group(1) for line in texts if (m := DATE_LINE_RE.match(line)) and valid_date(m.group(1))]
    if not dates:
        return body, [], "UPDATED holds no date"
    return ["'   " + max(dates)], body, None


def arrange_fields(fields: list[tuple[str, list[str]]], *, is_sub: bool, required: tuple[str, ...]) -> Fields:
    """Order the fields as the style defines and report what is missing."""
    todos: list[str] = []
    standard: dict[str, list[str]] = {}
    extras: list[tuple[str, list[str]]] = []
    for label, body in fields:
        label = hs.FIELD_ALIASES.get(label, label)
        if label in hs.FIELDS and label not in standard:
            standard[label] = body
        else:
            if label in hs.FIELDS:
                todos.append(f"{label} appears twice")
            extras.append((label, body))
    history: list[str] = []
    for label, body in extras:
        if label == hs.HISTORY:
            history = body
    extras = [(label, body) for label, body in extras if label != hs.HISTORY]
    if "UPDATED" in standard:
        updated, moved, problem = convert_updated(standard["UPDATED"])
        standard["UPDATED"] = updated
        if moved:
            history = moved + (["'"] + history if history else [])
        if problem:
            todos.append(problem)
    if is_sub and "RETURNS" in standard:
        said = text_of(standard["RETURNS"]).rstrip(".").lower()
        if said in ("none", "nothing", "n/a", "no value", "no return value"):
            standard["RETURNS"] = ["'   Nothing"]
        elif said != "nothing":
            todos.append("RETURNS of a Sub is not Nothing")
    missing = [label for label in required if label not in standard]
    if missing:
        todos.append("header lacks " + ", ".join(missing))
    called = [e for e in extras if e[0] == hs.CALLED_FROM]
    other = [e for e in extras if e[0] != hs.CALLED_FROM]
    ordered = [(label, standard[label]) for label in hs.STANDARD_BEFORE_EXTRAS if label in standard]
    ordered += called + other
    if history:
        ordered.append((hs.HISTORY, history))
    if "UPDATED" in standard:
        ordered.append(("UPDATED", standard["UPDATED"]))
    return Fields(ordered, todos)


def render_header(rule: str, title: str, fields: list[tuple[str, list[str]]]) -> list[Ln]:
    out = ["'", rule, title, rule]
    for n, (label, body) in enumerate(fields):
        if n:
            out.append("'")
        out.append("' " + label)
        out.extend(body)
    out.append(rule)
    return [Ln(t) for t in out]


class HeaderAt(NamedTuple):
    start: int      # first line of the header block, a lone apostrophe included
    rule1: int
    rule2: int
    rule3: int


def locate_header(lines: list[Ln], first: int, stop: int) -> HeaderAt | None:
    """The comment header that opens lines[first:stop], if it has its three rules."""
    n = first
    while n < stop and is_blank(lines[n]):
        n += 1
    start = n
    if n < stop and is_lone(lines[n]):
        n += 1
    if n >= stop or lines[n].fixed or not is_rule(lines[n].text):
        return None
    rules = [n]
    m = n + 1
    while m < stop and lines[m].text.startswith("'") and not lines[m].fixed:
        if is_rule(lines[m].text):
            rules.append(m)
            if len(rules) == 3:
                break
        m += 1
    if len(rules) < 3 or rules[1] != rules[0] + 2:
        return None
    return HeaderAt(start, rules[0], rules[1], rules[2])


# --------------------------------------------------------------------------
# Banners and blank lines


def banner_title(lines: list[Ln], n: int) -> str | None:
    """The title when lines n..n+2 are a section banner (rule, ' TITLE, rule)."""
    if n + 2 >= len(lines):
        return None
    a, b, c = lines[n], lines[n + 1], lines[n + 2]
    if a.fixed or b.fixed or c.fixed or not is_rule(a.text) or not is_rule(c.text):
        return None
    if re.match(r"^' \S", b.text) is None:
        return None
    return b.text[2:].strip()


def normalise_banners(lines: list[Ln]) -> None:
    """Rule characters to dashes and section titles to upper case."""
    n = 0
    while n < len(lines):
        title = banner_title(lines, n)
        if title is not None:
            lines[n].text = hs.RULE_DASH
            lines[n + 1].text = "' " + title.upper()
            lines[n + 2].text = hs.RULE_DASH
            n += 3
        else:
            n += 1


def banner_positions(lines: list[Ln]) -> list[tuple[int, str]]:
    out: list[tuple[int, str]] = []
    n = 0
    while n < len(lines):
        title = banner_title(lines, n)
        if title is not None:
            out.append((n, title))
            n += 3
        else:
            n += 1
    return out


def space_around_banners(lines: list[Ln]) -> list[Ln]:
    """One blank line above each section banner and none below it."""
    marks = {n for n, _ in banner_positions(lines)}
    out: list[Ln] = []
    n = 0
    while n < len(lines):
        if n in marks:
            ensure_blanks(out, 1)
            out.extend(lines[n:n + 3])
            n += 3
            while n < len(lines) and is_blank(lines[n]) and not lines[n].fixed:
                n += 1
            continue
        out.append(lines[n])
        n += 1
    return out


def collapse_blanks(lines: list[Ln]) -> list[Ln]:
    out: list[Ln] = []
    for line in lines:
        if is_blank(line) and not line.fixed and out and is_blank(out[-1]) and not out[-1].fixed:
            continue
        out.append(line)
    return out


# --------------------------------------------------------------------------
# Declarations


def align_declarations(lines: list[Ln]) -> None:
    """Align As and the trailing comment of each run of one-line declarations."""
    n = 0
    while n < len(lines):
        run: list[tuple[int, re.Match[str], str]] = []
        while n < len(lines):
            line = lines[n]
            code, comment = split_code_comment(line.text)
            m = None if continues(line.text) else DECL_RE.match(code.rstrip())
            if m is None:
                break
            run.append((n, m, comment.strip()))
            n += 1
        if run:
            as_col = round_up(max(len(m.group("ind")) + len(m.group("head")) for _, m, _ in run) + 1)
            tails = [as_col + 3 + len(m.group("rest")) for _, m, c in run if c]
            comment_col = round_up(max(tails) + 1) if tails else 0
            for at, m, comment in run:
                text = (m.group("ind") + m.group("head")).ljust(as_col) + "As " + m.group("rest")
                if comment:
                    text = text.ljust(comment_col) + comment
                lines[at].text = text
        else:
            n += 1


def indent_declaration_units(lines: list[Ln], exact: bool) -> None:
    """Put declarations and their comments under a banner at the declaration indent.

    A unit is a declaration with its continuation lines, or an Enum/Type block.
    ``exact`` also lowers a unit that sits deeper; otherwise a unit moves only
    when it sits to the left of the declaration indent.
    """
    n = 0
    while n < len(lines):
        line = lines[n]
        if line.fixed or is_blank(line):
            n += 1
            continue
        if banner_title(lines, n) is not None:
            n += 3
            continue
        stop = n
        if BLOCK_TYPE_RE.match(code_of(line.text)):
            while stop + 1 < len(lines) and not END_BLOCK_TYPE_RE.match(code_of(lines[stop].text)):
                stop += 1
        else:
            while stop + 1 < len(lines) and not lines[stop].fixed and continues(lines[stop].text):
                stop += 1
        lead = indent_of(line.text)
        delta = hs.DECLARATION_INDENT - lead
        if delta > 0 or (exact and delta < 0 and not line.text.lstrip().startswith("'")):
            for k in range(n, stop + 1):
                if not lines[k].fixed and lines[k].text.strip():
                    lines[k].text = " " * max(0, indent_of(lines[k].text) + delta) + lines[k].text.lstrip(" ")
        n = stop + 1


def format_declaration_area(lines: list[Ln]) -> list[Ln]:
    normalise_banners(lines)
    lines = collapse_blanks(space_around_banners(lines))
    indent_declaration_units(lines, exact=False)
    align_declarations(lines)
    return lines


# --------------------------------------------------------------------------
# Group banners


def group_banner_runs(lines: list[Ln]) -> list[tuple[int, int, str]]:
    """Runs of column-0 comments that form a group banner: (first, last, title)."""
    out: list[tuple[int, int, str]] = []
    n = 0
    while n < len(lines):
        if lines[n].fixed or not lines[n].text.startswith("'"):
            n += 1
            continue
        m = n
        while m + 1 < len(lines) and not lines[m + 1].fixed and lines[m + 1].text.startswith("'"):
            m += 1
        run = [lines[k].text for k in range(n, m + 1)]
        rules = [k for k, t in enumerate(run) if is_rule(t)]
        texts = [k for k, t in enumerate(run) if t != "'" and not is_rule(t)]
        if len(rules) == 2 and len(texts) == 1 and rules[0] < texts[0] < rules[1] and \
                all(t == "'" for k, t in enumerate(run) if k not in (rules[0], rules[1], texts[0])):
            title = run[texts[0]][1:].strip()
            plain = run[texts[0]].startswith("' ") and not run[texts[0]].startswith("'  ")
            surrounded = rules[0] > 0 or texts[0] != rules[0] + 1 or rules[1] != texts[0] + 1 or len(run) > 3
            if title and (surrounded or not plain):
                out.append((n, m, title))
        n = m + 1
    return out


def group_banner_lines(title: str) -> list[Ln]:
    title = title.upper()
    return [Ln(t) for t in ("'", hs.RULE_DASH, "'", hs.centred_title(title), "'", hs.RULE_DASH, "'")]


def replace_group_banners(lines: list[Ln]) -> tuple[list[Ln], list[tuple[int, int]]]:
    """Rewrite every group banner to its canonical block; return the block extents."""
    out: list[Ln] = []
    blocks: list[tuple[int, int]] = []
    runs = {first: (last, title) for first, last, title in group_banner_runs(lines)}
    n = 0
    while n < len(lines):
        if n in runs:
            last, title = runs[n]
            block = group_banner_lines(title)
            blocks.append((len(out), len(out) + len(block) - 1))
            out.extend(block)
            n = last + 1
        else:
            out.append(lines[n])
            n += 1
    return out, blocks


# --------------------------------------------------------------------------
# Procedure bodies


def section_of(marks: list[tuple[int, str]], index: int) -> tuple[int, str] | None:
    found: tuple[int, str] | None = None
    for at, title in marks:
        if at > index:
            break
        found = (at, title)
    return found


def first_code_after(lines: list[Ln], start: int, stop: int) -> int | None:
    for n in range(start, stop):
        text = lines[n].text.strip()
        if text and not text.startswith("'"):
            return n
    return None


def next_section_start(marks: list[tuple[int, str]], at: int, stop: int) -> int:
    for other, _ in marks:
        if other > at:
            return other
    return stop


def set_title(lines: list[Ln], at: int, title: str) -> None:
    lines[at + 1].text = "' " + title


def insert_banner(lines: list[Ln], at: int, title: str) -> None:
    """Open a section at line ``at``, with a blank line above the banner."""
    lines[at:at] = [Ln(""), Ln(hs.RULE_DASH), Ln("' " + title), Ln(hs.RULE_DASH)]


def start_of_comment_run(lines: list[Ln], at: int, floor: int) -> int:
    """First line of the comment lines that sit directly above line ``at``."""
    n = at
    while n - 1 >= floor and lines[n - 1].text.lstrip().startswith("'") and not is_rule(lines[n - 1].text):
        n -= 1
    return n


def find_label(lines: list[Ln], labels: frozenset[str]) -> int | None:
    for n, ln in enumerate(lines):
        if not ln.fixed and (m := LABEL_LINE_RE.match(code_of(ln.text))) and m.group(1) in labels:
            return n
    return None


def hoist_labels(lines: list[Ln]) -> None:
    """A label written above a section banner belongs to the section below it."""
    n = 0
    while n < len(lines):
        m = None if lines[n].fixed else LABEL_LINE_RE.match(code_of(lines[n].text))
        if m:
            k = n + 1
            while k < len(lines) and is_blank(lines[k]) and not lines[k].fixed:
                k += 1
            if banner_title(lines, k) is not None and not lines[n].text.lstrip().startswith("'"):
                label = lines.pop(n)
                lines.insert(k + 2, label)
                continue
        n += 1


STANDARD_SECTIONS = (hs.DECLARE, hs.INITIALIZE, hs.EXIT_PROCEDURE, hs.ERROR_HANDLER)


def place_label_sections(lines: list[Ln]) -> None:
    """Give the clean-exit label and the handler label their standard sections."""
    for labels, title in ((EXIT_LABELS, hs.EXIT_PROCEDURE), (HANDLER_LABELS, hs.ERROR_HANDLER)):
        found = find_label(lines, labels)
        if found is None:
            continue
        here = section_of(banner_positions(lines), found)
        if here is None:
            insert_banner(lines, start_of_comment_run(lines, found, 0), title)
            continue
        at, current = here
        if current == title:
            continue
        if first_code_after(lines, at + 3, found + 1) == found and current not in STANDARD_SECTIONS:
            set_title(lines, at, title)
        else:
            insert_banner(lines, start_of_comment_run(lines, found, at + 3), title)
    # With no exit label, the Exit statement ends the section above the handler.
    marks = banner_positions(lines)
    titles = [t for _, t in marks]
    if hs.ERROR_HANDLER not in titles or hs.EXIT_PROCEDURE in titles:
        return
    handler = marks[titles.index(hs.ERROR_HANDLER)][0]
    codes = [n for n in range(handler) if code_of(lines[n].text) and not lines[n].text.lstrip().startswith("'")]
    if not codes or not EXIT_RE.match(code_of(lines[codes[-1]].text)):
        return
    last = codes[-1]
    here = section_of(marks, last)
    if here is not None and here[1] in hs.EXIT_TITLES and first_code_after(lines, here[0] + 3, last + 1) == last:
        set_title(lines, here[0], hs.EXIT_PROCEDURE)
    elif here is None or here[1] not in STANDARD_SECTIONS or here[1] == hs.DECLARE:
        insert_banner(lines, start_of_comment_run(lines, last, here[0] + 3 if here else 0), hs.EXIT_PROCEDURE)


def is_declaration(code: str) -> bool:
    return code.startswith(("Dim ", "Const ", "Static ", "ReDim ", "Public Const ", "Private Const "))


def split_initialize(lines: list[Ln], todos: list[str]) -> None:
    """Open INITIALIZE at the statement that arms the error handler."""
    marks = banner_positions(lines)
    armed = next((n for n, ln in enumerate(lines) if not ln.fixed and ON_ERROR_RE.match(code_of(ln.text))), None)
    here = None if armed is None else section_of(marks, armed)
    if armed is None or here is None:
        return
    at, current = here
    if current == hs.INITIALIZE:
        return
    blocked = any(code_of(ln.text) and not (current == hs.DECLARE and is_declaration(code_of(ln.text)))
                  and not ln.text.lstrip().startswith("'") for ln in lines[at + 3:armed])
    if blocked or any(title == hs.INITIALIZE for _, title in marks):
        todos.append("On Error GoTo is not in INITIALIZE")
        return
    if current == hs.DECLARE:
        insert_banner(lines, start_of_comment_run(lines, armed, at + 3), hs.INITIALIZE)
        return
    end = next_section_start(marks, at, len(lines))
    set_title(lines, at, hs.INITIALIZE)
    if any(ln.text.strip() for ln in lines[armed + 1:end]):
        insert_banner(lines, armed + 1, current)


BLOCK_ENDS = (("end if", "If"), ("end with", "With"), ("end select", "Select"), ("wend", "While"), ("loop", "Do"))
BLOCK_MIDDLES = (("elseif", "If"), ("else", "If"), ("case", "Select"))
BLOCK_STARTS = (("For", r"for\b"), ("Do", r"do\b"), ("While", r"while\b"), ("With", r"with\b"),
                ("Select", r"select case\b"))


def block_depth(stack: list[str]) -> int:
    """Levels of indentation the open blocks add; Select Case adds two."""
    return sum(2 if kind == "Select" else 1 for kind in stack)


def close_block(stack: list[str], kind: str, count: int = 1) -> None:
    for _ in range(count):
        if not stack or stack.pop() != kind:
            raise Bail


def middle_level(low: str, stack: list[str]) -> int | None:
    """Level of Else, ElseIf and Case, which sit one level out of their block."""
    for prefix, kind in BLOCK_MIDDLES:
        if re.match(prefix + r"\b", low):
            if not stack or stack[-1] != kind or (prefix == "else" and low != "else") \
                    or (prefix == "elseif" and not low.endswith(" then")):
                raise Bail
            return block_depth(stack) - 1
    return None


def statement_level(code: str, stack: list[str]) -> int:
    """Indent level of a statement; updates the stack of open blocks."""
    low = re.sub(r"\s+", " ", code.strip()).lower()
    for prefix, kind in BLOCK_ENDS:
        if re.match(prefix + r"\b", low):
            close_block(stack, kind)
            return block_depth(stack)
    if re.match(r"next\b", low):
        close_block(stack, "For", len([p for p in low[4:].split(",") if p.strip()]) or 1)
        return block_depth(stack)
    middle = middle_level(low, stack)
    if middle is not None:
        return middle
    level = block_depth(stack)
    if re.match(r"if\b", low):
        if low.endswith(" then"):
            stack.append("If")
        elif not re.search(r"\bthen\b", low):
            raise Bail
        return level
    for kind, pattern in BLOCK_STARTS:
        if re.match(pattern, low):
            stack.append(kind)
            break
    return level


@dataclass
class IndentPlan:
    """Where each statement and comment of a body goes, found before any line moves."""
    stack: list[str]
    edits: list[tuple[int, int, int]]       # (first line, last line, indent of the first)
    conditional: int = 0


def plan_comment(stripped: str, n: int, section: str, plan: IndentPlan) -> None:
    if section != hs.DECLARE and len(stripped) > 1 and stripped[1] != " " and not is_rule(stripped):
        plan.edits.append((n, n, hs.INTENT_INDENT + hs.INDENT_STEP * block_depth(plan.stack)))


def plan_statement(lines: list[Ln], first: int, last: int, plan: IndentPlan) -> None:
    code = logical_code(lines, first, last)
    if LABEL_LINE_RE.match(code):
        plan.edits.append((first, last, 0))
        return
    if SEPARATOR_RE.search(code) or code[:1].isdigit():
        raise Bail
    before = list(plan.stack)
    level = statement_level(code, plan.stack)
    if plan.conditional and plan.stack != before:
        raise Bail
    plan.edits.append((first, last, hs.CODE_INDENT + hs.INDENT_STEP * level))


def plan_line(lines: list[Ln], n: int, section: str | None, plan: IndentPlan) -> int:
    """Plan the logical line at ``n``; returns the index of the next one."""
    line = lines[n]
    stripped = line.text.strip()
    if not stripped:
        return n + 1
    if line.fixed and stripped.startswith("#"):
        code = code_of(line.text)
        plan.conditional += bool(re.match(r"#If\b", code, re.I)) - bool(re.match(r"#End\s+If\b", code, re.I))
        return n + 1
    if section is None:
        return n + 1
    if stripped.startswith("'"):
        plan_comment(stripped, n, section, plan)
        return n + 1
    last = n
    while last + 1 < len(lines) and continues(lines[last].text):
        last += 1
    if section != hs.DECLARE:
        plan_statement(lines, n, last, plan)
    return last + 1


def indent_body(lines: list[Ln], todos: list[str]) -> None:
    """Statements at 8 spaces plus 4 per open block; their one-line comments 4 less."""
    marks = dict(banner_positions(lines))
    if not marks:
        return
    plan = IndentPlan([], [])
    section: str | None = None
    n = 0
    try:
        while n < len(lines):
            if n in marks:
                section = marks[n]
                n += 3
            else:
                n = plan_line(lines, n, section, plan)
        if plan.stack:
            raise Bail
    except Bail:
        todos.append("indentation not changed (block structure not recognised)")
        return
    for first, last, want in plan.edits:
        delta = want - indent_of(lines[first].text)
        for k in range(first, last + 1):
            lines[k].text = " " * max(0, indent_of(lines[k].text) + delta) + lines[k].text.lstrip(" ")


PROC_NAME_RE = re.compile(r"^\s*Const\s+PROC_NAME\s+As\s+String\b", re.I)


def separate_proc_name(part: list[Ln]) -> None:
    """One blank line between Const PROC_NAME and the variables that follow it."""
    first = next((n for n, ln in enumerate(part) if ln.text.strip()), None)
    if first is None or part[first].fixed or not PROC_NAME_RE.match(part[first].text):
        return
    last = first
    while last + 1 < len(part) and continues(part[last].text):
        last += 1
    if last + 1 < len(part) and part[last + 1].text.strip():
        part.insert(last + 1, Ln(""))


def align_declare_sections(lines: list[Ln]) -> None:
    marks = banner_positions(lines)
    for pos, (at, title) in enumerate(marks):
        if title == hs.DECLARE:
            stop = marks[pos + 1][0] if pos + 1 < len(marks) else len(lines)
            part = lines[at + 3:stop]
            separate_proc_name(part)
            indent_declaration_units(part, exact=True)
            align_declarations(part)
            lines[at + 3:stop] = part


def format_body(lines: list[Ln], todos: list[str], *, trim: bool = True) -> list[Ln]:
    normalise_banners(lines)
    hoist_labels(lines)
    place_label_sections(lines)
    split_initialize(lines, todos)
    lines = collapse_blanks(space_around_banners(lines))
    if trim:
        while lines and is_blank(lines[0]):
            lines.pop(0)
        while lines and is_blank(lines[-1]) and not lines[-1].fixed:
            lines.pop()
    indent_body(lines, todos)
    align_declare_sections(lines)
    return lines


# --------------------------------------------------------------------------
# Procedures and the module


def code_line_count(body: list[Ln]) -> int:
    return sum(1 for ln in body if ln.text.strip() and not ln.text.lstrip().startswith("'"))


def format_procedure(lines: list[Ln], todos: list[Todo]) -> list[Ln]:
    """One procedure, from its signature to its End line."""
    last_sig = signature_end(lines, 0)
    signature, inner, end = lines[:last_sig + 1], lines[last_sig + 1:-1], lines[-1]
    m = signature_match(lines, 0)
    assert m is not None
    name = m.group("name")
    mine: list[str] = []
    at = locate_header(inner, 0, len(inner))
    parsed = None if at is None else parse_fields([ln.text for ln in inner[at.rule2 + 1:at.rule3]])
    if at is None or parsed is None:
        out = signature + format_body(list(inner), mine, trim=False) + [end]
        mine.append("no documentation header" if at is None else "header text before its first field")
    else:
        rest = list(inner[at.rule3 + 1:])
        while rest and (is_blank(rest[0]) or is_lone(rest[0])):
            rest.pop(0)
        body = format_body(rest, mine)
        if code_line_count(body) > 1 and not banner_positions(body):
            mine.append("body has no section banners")
        compact = (m.group("vis") or "").strip() == "Private" and code_line_count(body) <= hs.COMPACT_MAX_CODE_LINES
        arranged = arrange_fields(parsed, is_sub=m.group("kind") == "Sub",
                                  required=hs.COMPACT_FIELDS if compact else hs.FIELDS)
        mine.extend(arranged.todos)
        said = inner[at.rule1 + 1].text[1:].strip()
        title = hs.title_from_identifier(name) if said.lower() == name.lower() else said.upper()
        out = list(signature)
        if last_sig == 0:
            out.append(Ln(""))
        out.extend(render_header(hs.RULE_DASH, hs.centred_title(title), arranged.ordered))
        out.append(Ln(""))
        out.extend(body)
        if body:
            out.append(Ln(""))
        out.append(end)
    todos.extend((name, t) for t in mine)
    return out


def signature_end(lines: list[Ln], first: int) -> int:
    """Last line of the signature that starts at ``first``.

    A signature may be wrapped in #If ... #End If to hold a VBA7 and a legacy
    form; the whole conditional block is then the signature.
    """
    if lines[first].fixed and re.match(r"#If\b", code_of(lines[first].text), re.I):
        depth = 0
        for n in range(first, len(lines)):
            code = code_of(lines[n].text)
            depth += bool(re.match(r"#If\b", code, re.I)) - bool(re.match(r"#End\s+If\b", code, re.I))
            if depth == 0:
                return n
        raise LexicalError("unterminated conditional signature")
    last = first
    while continues(lines[last].text):
        last += 1
    return last


def signature_match(lines: list[Ln], first: int) -> re.Match[str] | None:
    """The signature pattern at ``first``, looking inside a conditional wrapper."""
    n = first
    while n < len(lines) and lines[n].fixed and lines[n].text.lstrip().startswith("#"):
        n += 1
    return SIG_RE.match(lines[n].text.strip()) if n < len(lines) else None


def procedure_spans(lines: list[Ln]) -> list[tuple[int, int]]:
    """(signature line, End line) of each procedure; a conditional signature counts."""
    out: list[tuple[int, int]] = []
    n = 0
    while n < len(lines):
        wrapped = lines[n].fixed and re.match(r"#If\b", code_of(lines[n].text), re.I) is not None \
            and n + 1 < len(lines) and SIG_RE.match(lines[n + 1].text.strip()) is not None
        if wrapped or (not lines[n].fixed and SIG_RE.match(lines[n].text.strip())):
            m = signature_end(lines, n) + 1
            while m < len(lines) and not (END_RE.match(lines[m].text.strip()) and not lines[m].fixed):
                m += 1
            if m == len(lines):
                raise LexicalError(f"procedure at line {n + 1} has no End")
            out.append((n, m))
            n = m + 1
        else:
            n += 1
    return out


def format_gap(lines: list[Ln]) -> tuple[list[Ln], int]:
    """Text between pieces of code: group banners, loose comments. Returns it and
    the number of blank lines that must follow it."""
    lines = list(lines)
    while lines and is_blank(lines[0]) and not lines[0].fixed:
        lines.pop(0)
    trailing = 0
    while lines and is_blank(lines[-1]) and not lines[-1].fixed:
        lines.pop()
        trailing += 1
    lines, blocks = replace_group_banners(lines)
    lines = collapse_blanks(lines)
    if blocks and blocks[-1][1] == len(lines) - 1:
        trailing = 1
    return lines, min(trailing, 1)


def split_top(top: list[Ln]) -> tuple[list[Ln], list[Ln]]:
    """Declarations, and the group banners or attached comments that lead to the first procedure."""
    cut = len(top)
    runs = group_banner_runs(top)
    if runs:
        cut = runs[0][0]
    stop = len(top)
    while stop and is_blank(top[stop - 1]):
        stop -= 1
    if stop == len(top):
        while stop and top[stop - 1].text.startswith("'") and not top[stop - 1].fixed:
            stop -= 1
        cut = min(cut, stop)
    return top[:cut], top[cut:]


def split_preamble(raw: list[str], conditional: set[int]) -> tuple[list[Ln], list[Ln], list[Ln]]:
    """Attribute lines, Option statements (hoisted from anywhere) and the rest of the module."""
    attributes: list[Ln] = []
    options: list[Ln] = []
    rest: list[Ln] = []
    for n, ln in enumerate(raw):
        line = Ln(ln.rstrip(), n in conditional)
        if line.text.startswith("Option "):
            if line.fixed:
                raise LexicalError("Option statement inside a conditional block")
            options.append(line)
        elif not rest and not line.text.strip():
            continue
        elif not rest and line.text.startswith("Attribute "):
            attributes.append(line)
        else:
            rest.append(line)
    if rest and is_lone(rest[0]):
        rest.pop(0)
    return attributes, options, rest


def component_header(rest: list[Ln], module_name: str, todos: list[Todo]) -> tuple[list[Ln], int]:
    """The rebuilt component header, and the index in ``rest`` where it ended."""
    at = locate_header(rest, 0, len(rest))
    parsed = None if at is None else parse_fields([ln.text for ln in rest[at.rule2 + 1:at.rule3]])
    if at is None or parsed is None:
        reason = "no component header" if at is None else "component header text before its first field"
        todos.append((module_name, reason))
        return [], 0
    arranged = arrange_fields(parsed, is_sub=False, required=hs.FIELDS)
    todos.extend((module_name, t) for t in arranged.todos)
    rule = hs.RULE_EQUALS if rest[at.rule1].text.startswith("'=") else hs.RULE_DASH
    end = at.rule3 + 1
    while end < len(rest) and is_lone(rest[end]):
        end += 1
    return render_header(rule, "' MODULE: " + module_name, arranged.ordered), end


def layout_units(body: list[Ln], todos: list[Todo]) -> list[tuple[list[Ln], int]]:
    """The declarations, the gaps and the procedures, each with the blank lines above it."""
    found = procedure_spans(body)
    first = found[0][0] if found else len(body)
    declarations, lead = split_top(body[:first])
    declarations = format_declaration_area(declarations)
    units: list[tuple[list[Ln], int]] = [(declarations, 1)]
    blanks = 2 if any(ln.text.strip() for ln in declarations) else 1
    gap_lines, gap_blanks = format_gap(lead)
    if gap_lines:
        units.append((gap_lines, blanks))
        blanks = gap_blanks
    previous_end = first
    for index, (sig, end) in enumerate(found):
        if index:
            gap_lines, gap_blanks = format_gap(body[previous_end:sig])
            if gap_lines:
                units.append((gap_lines, 1))
                blanks = gap_blanks
        units.append((format_procedure(body[sig:end + 1], todos), blanks))
        blanks = 1
        previous_end = end + 1
    if found and any(ln.text.strip() for ln in body[previous_end:]):
        units.append((format_gap(body[previous_end:])[0], 1))
    return units


def format_module(text: str, module_name: str) -> tuple[str, list[Todo]]:
    """The house-style form of one module, and what has to be settled by hand."""
    if not text.isascii():
        raise LexicalError("non-ASCII VBA input; repository exports must be ASCII")
    raw = split_lines(text)
    protected, _ = protected_lines(raw)
    conditional = conditional_lines(raw)
    label_map = dict(hs.LABEL_RENAMES)
    check_renames(raw, label_map, protected, module_name)
    raw = rename_labels(raw, label_map)
    raw = rename_proc_constant(raw, conditional, module_name)
    todos: list[Todo] = []
    attributes, options, rest = split_preamble(raw, conditional)
    out: list[Ln] = list(attributes) + [Ln("")] + options
    if options:
        out.append(Ln(""))
    header, header_end = component_header(rest, module_name, todos)
    out.extend(header)
    for piece, above in layout_units(rest[header_end:], todos):
        ensure_blanks(out, above)
        out.extend(piece)
    while out and is_blank(out[-1]) and not out[-1].fixed:
        out.pop()
    return "\r\n".join(ln.text for ln in out) + "\r\n", todos
