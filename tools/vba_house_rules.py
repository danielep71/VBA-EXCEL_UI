"""The house-style rules that rewriting cannot settle, read from the source text.

The formatter (vba_house.py) puts every module in the layout of the style. It
cannot write the text of a header field, add the PROC_NAME constant, give a
Select Case its Case Else or choose a compare mode, because each of those
changes what a person wrote or what the code does. This module reports them.

``findings`` reads one module and returns every deviation, whatever the state
of ``house_style.HOUSE_STYLE_ACTIVE``. The gate decides whether to scan the
tree with it (``check_repo.check_house_style_rules``); its own fixtures call
it directly, so each rule stays proved while the style is off.

While the modules are brought to the style, the findings that are still to be
settled by hand are listed in tools/house_style_pending.txt, one line per
finding. ``compare_pending`` holds the tree to that list in both directions: a
finding that is not listed fails, and so does a listed line that no finding
matches any more.
"""
import re
from collections import Counter
from collections.abc import Iterable
from typing import NamedTuple

import house_style as hs
from vba_house import (
    LABEL_LINE_RE,
    ON_ERROR_RE,
    Ln,
    banner_positions,
    code_of,
    conditional_lines,
    continues,
    locate_header,
    parse_fields,
    procedure_spans,
    signature_end,
    signature_match,
    split_lines,
)
from vba_lex import code_only, split_code_comment

RULES = ("header", "sections", "proc-name", "select-case", "compare-mode", "statements", "banners")


class Finding(NamedTuple):
    line: int
    rule: str
    where: str
    message: str

    def __str__(self) -> str:
        return f"{self.line}: {self.where}: {self.message}"


def joined(lines: list[Ln], first: int, last: int) -> tuple[str, str]:
    """Code of a logical line, comments dropped, and the same text with literals masked."""
    raw: list[str] = []
    masked: list[str] = []
    for n in range(first, last + 1):
        code = split_code_comment(lines[n].text)[0]
        text = code_only(code)
        if n < last and text.rstrip().endswith("_"):
            code, text = code.rstrip()[:-1], text.rstrip()[:-1]
        raw.append(code)
        masked.append(text)
    return " ".join(raw), " ".join(masked)


def split_arguments(masked: str, open_at: int) -> tuple[list[tuple[int, int]], int] | None:
    """Spans of the arguments of the call whose parenthesis opens at ``open_at``."""
    depth = 0
    start = open_at + 1
    spans: list[tuple[int, int]] = []
    for n in range(open_at, len(masked)):
        ch = masked[n]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                spans.append((start, n))
                return spans, n + 1
        elif ch == "," and depth == 1:
            spans.append((start, n))
            start = n + 1
    return None


COMPARE_CALL_RE = re.compile(r"(?<![\w.])(?:VBA\.)?(InStr|StrComp)\s*\(", re.I)
POSITIONAL_COMPARE = {"instr": 4, "strcomp": 3}      # the argument position of the compare mode


def missing_compare(name: str, args: list[str]) -> bool:
    """True when the call names no compare mode.

    The mode is the fourth positional argument of InStr (start, string1,
    string2, compare), the third of StrComp (string1, string2, compare), or
    ``Compare:=`` by name. What an argument is spelled does not matter: with
    three positional arguments InStr has a start and two strings, whatever the
    third one looks like.
    """
    named = [re.match(r"\s*(\w+)\s*:=", a) for a in args]
    if any(m and m.group(1).lower() == "compare" for m in named):
        return False
    positional = sum(1 for m in named if m is None)
    return positional < POSITIONAL_COMPARE[name.lower()]


class ProcedureView(NamedTuple):
    name: str
    kind: str
    visibility: str
    sig: int
    sig_last: int
    end: int
    header: tuple[int, int, int, int] | None      # start, rule1, rule2, rule3
    body_first: int
    code_lines: int


def read_procedures(lines: list[Ln]) -> list[ProcedureView]:
    out: list[ProcedureView] = []
    for sig, end in procedure_spans(lines):
        m = signature_match(lines, sig)
        if m is None:
            continue
        last = signature_end(lines, sig)
        at = locate_header(lines, last + 1, end)
        first = at.rule3 + 1 if at else last + 1
        code = sum(1 for k in range(first, end) if lines[k].text.strip() and not lines[k].text.lstrip().startswith("'"))
        out.append(ProcedureView(m.group("name"), m.group("kind"), (m.group("vis") or "").strip(), sig, last, end,
                                 (at.start, at.rule1, at.rule2, at.rule3) if at else None, first, code))
    return out


def header_findings(lines: list[Ln], rule: tuple[int, int, int, int] | None, where: str,
                    required: tuple[str, ...], is_sub: bool, line: int) -> list[Finding]:
    if rule is None:
        return [Finding(line + 1, "header", where, "no documentation header")]
    parsed = parse_fields([ln.text for ln in lines[rule[2] + 1:rule[3]]])
    if parsed is None:
        return [Finding(rule[1] + 1, "header", where, "header text before its first field")]
    labels = [label for label, _ in parsed]
    out: list[Finding] = []
    seen = [label for label in labels if label in hs.FIELDS]
    missing = [label for label in required if label not in seen]
    if missing:
        out.append(Finding(rule[1] + 1, "header", where, "header lacks " + ", ".join(missing)))
    if seen != sorted(seen, key=hs.FIELDS.index):
        out.append(Finding(rule[1] + 1, "header", where, "header fields are out of order"))
    for n, label in enumerate(labels):
        if label in hs.FIELDS:
            continue
        later = [x for x in labels[n + 1:] if x in hs.FIELDS and x != "UPDATED"]
        if later:
            out.append(Finding(rule[1] + 1, "header", where, f"section {label} must follow {later[-1]}"))
            break
    if labels and labels[-1] != "UPDATED" and "UPDATED" in labels:
        out.append(Finding(rule[1] + 1, "header", where, "UPDATED must be the last section"))
    for label, body in parsed:
        text = [t for t in body if t != "'"]
        if label == "UPDATED" and not (len(text) == 1 and re.fullmatch(r"'   \d{4}-\d{2}-\d{2}", text[0])):
            out.append(Finding(rule[1] + 1, "header", where, "UPDATED is not a single YYYY-MM-DD date"))
        if label == "RETURNS" and is_sub and not (text and re.match(r"'   Nothing\b", text[0])):
            out.append(Finding(rule[1] + 1, "header", where, "RETURNS of a Sub is not Nothing"))
    return out


def section_findings(lines: list[Ln], view: ProcedureView) -> list[Finding]:
    out: list[Finding] = []
    marks = [(n, t) for n, t in banner_positions(lines) if view.body_first <= n < view.end]
    prev, rank = "", -1
    for n, title in marks:
        current = hs.SECTION_RANK.get(title, hs.TASK_RANK)
        if current < rank:
            out.append(Finding(n + 1, "sections", view.name, f"section {title} is out of order (after {prev})"))
            break
        prev, rank = title, current
    armed = next((n for n in range(view.body_first, view.end) if ON_ERROR_RE.match(code_of(lines[n].text))), None)
    if armed is not None:
        owner = [t for n, t in marks if n < armed]
        if not owner or owner[-1] != hs.INITIALIZE:
            out.append(Finding(armed + 1, "sections", view.name, "On Error GoTo is not in INITIALIZE"))
        target = ON_ERROR_RE.match(code_of(lines[armed].text))
        assert target is not None
        label_at = next((n for n in range(armed, view.end)
                         if (m := LABEL_LINE_RE.match(code_of(lines[n].text))) and m.group(1) == target.group(1)), None)
        if label_at is None:
            out.append(Finding(armed + 1, "sections", view.name, f"label {target.group(1)}: not found"))
        else:
            if lines[label_at].text[:1].isspace():
                out.append(Finding(label_at + 1, "sections", view.name, f"label {target.group(1)}: is not at column 0"))
            exits = [n for n in range(view.body_first, label_at)
                     if re.match(r"Exit (Sub|Function|Property)$", code_of(lines[n].text), re.I)
                     and ([t for m, t in marks if m < n] or [""])[-1] == hs.EXIT_PROCEDURE]
            if not exits:
                out.append(Finding(label_at + 1, "sections", view.name,
                                   f"no Exit statement in EXIT PROCEDURE before {target.group(1)}:"))
    return out


def proc_name_findings(lines: list[Ln], view: ProcedureView, module: str) -> list[Finding]:
    marks = banner_positions(lines)
    declare = next((n for n, t in marks if t == hs.DECLARE and view.body_first <= n < view.end), None)
    body = range(view.body_first, view.end)
    has_local = any(re.match(r"(?:Dim|Static)\s+\w", code_of(lines[n].text), re.I) for n in body)
    has_handler = any(ON_ERROR_RE.match(code_of(lines[n].text)) for n in body)
    if declare is None:
        if has_local or has_handler:
            return [Finding(view.sig + 1, "proc-name", view.name,
                            "no DECLARE section for the local variables or the error handler")]
        return []
    first = next((n for n in range(declare + 3, view.end) if is_code(lines[n])), None)
    if first is None:
        return [Finding(declare + 1, "proc-name", view.name, "DECLARE is empty")]
    last = first
    while last + 1 < view.end and continues(lines[last].text):
        last += 1
    raw, _ = joined(lines, first, last)
    m = re.match(r'\s*Const\s+PROC_NAME\s+As\s+String\s*=\s*"((?:[^"]|"")*)"\s*$', raw, re.I)
    if m is None:
        return [Finding(first + 1, "proc-name", view.name, "DECLARE does not start with Const PROC_NAME As String")]
    value = m.group(1).lower()
    qualified = f"{module}.{view.name}"
    deliberate = first > 0 and lines[first - 1].text.strip() == hs.DELIBERATE_VALUE_COMMENT
    if value == view.name.lower():
        if deliberate:
            return []
        said = (f'PROC_NAME is the bare name "{m.group(1)}": expected "{qualified}", or the comment '
                f"{hs.DELIBERATE_VALUE_COMMENT} directly above it")
        return [Finding(first + 1, "proc-name", view.name, said)]
    if deliberate:
        said = (f'the deliberate-value comment stands above PROC_NAME "{m.group(1)}", '
                f'which is not the bare name "{view.name}"')
        return [Finding(first + 1, "proc-name", view.name, said)]
    if value != qualified.lower():
        said = f'PROC_NAME is "{m.group(1)}", expected "{qualified}"'
        return [Finding(first + 1, "proc-name", view.name, said)]
    return []


def statement_findings(lines: list[Ln], view: ProcedureView) -> list[Finding]:
    out: list[Finding] = []
    selects: list[list[int | bool]] = []
    n = view.sig
    while n < view.end:
        if not lines[n].text.strip() or lines[n].text.lstrip().startswith("'"):
            n += 1
            continue
        last = n
        while last + 1 < view.end and continues(lines[last].text):
            last += 1
        raw, masked = joined(lines, n, last)
        code = masked.strip()
        low = code.lower()
        if not code.startswith("#"):
            plain = code.replace(":=", "  ")
            if ":" in plain and not LABEL_LINE_RE.match(code):
                out.append(Finding(n + 1, "statements", view.name, "statement separator ':' (one statement per line)"))
            if re.fullmatch(r"next", low):
                out.append(Finding(n + 1, "statements", view.name, "Next without the counter name"))
            if re.match(r"select\s+case\b", low):
                selects.append([n + 1, False])
            elif re.match(r"case\s+else\b", low) and selects:
                selects[-1][1] = True
            elif re.match(r"end\s+select\b", low) and selects:
                at, has_else = selects.pop()
                if not has_else:
                    out.append(Finding(int(at), "select-case", view.name, "Select Case without Case Else"))
            for call in COMPARE_CALL_RE.finditer(masked):
                parsed = split_arguments(masked, call.end() - 1)
                if parsed is None:
                    continue
                args = [raw[a:b] for a, b in parsed[0]]
                if args == [""]:
                    continue
                if missing_compare(call.group(1), args):
                    said = f"{call.group(1)} without an explicit compare mode"
                    out.append(Finding(n + 1, "compare-mode", view.name, said))
        n = last + 1
    return out


def first_module_line(lines: list[Ln]) -> int:
    """The first line that is neither blank, an Attribute line nor an Option statement."""
    for n, ln in enumerate(lines):
        if ln.text.strip() and not ln.text.startswith(("Attribute ", "Option ")):
            return n
    return len(lines)


def is_code(line: Ln) -> bool:
    return bool(code_of(line.text)) and not line.text.lstrip().startswith("'")


def component_findings(lines: list[Ln], module: str) -> list[Finding]:
    first = first_module_line(lines)
    at = locate_header(lines, first, len(lines))
    if at is None or at.start != first:
        return [Finding(first + 1, "header", module, "no component header")]
    title = lines[at.rule1 + 1].text
    out: list[Finding] = []
    if title != f"' MODULE: {module}":
        out.append(Finding(at.rule1 + 2, "header", module, f"component title is not ' MODULE: {module}"))
    out.extend(header_findings(lines, (at.start, at.rule1, at.rule2, at.rule3), module, hs.FIELDS, False, first))
    return out


def findings(text: str, module: str) -> list[Finding]:
    """Every rule of the style that rewriting cannot settle, for one module."""
    raw = split_lines(text)
    fixed = conditional_lines(raw)
    lines = [Ln(t.rstrip(), n in fixed) for n, t in enumerate(raw)]
    out = component_findings(lines, module)
    header_at = locate_header(lines, first_module_line(lines), len(lines))
    skip = range(header_at.start, header_at.rule3 + 1) if header_at else range(0)
    for n, ln in enumerate(lines):
        if n not in skip and re.fullmatch(r"'={5,}", ln.text.rstrip()):
            out.append(Finding(n + 1, "banners", module, "banner rule uses '=' outside the component header"))
    for view in read_procedures(lines):
        compact = view.visibility == "Private" and view.code_lines <= hs.COMPACT_MAX_CODE_LINES
        required = hs.COMPACT_FIELDS if compact else hs.FIELDS
        out.extend(header_findings(lines, view.header, view.name, required, view.kind == "Sub", view.sig))
        out.extend(statement_findings(lines, view))
        out.extend(section_findings(lines, view))
        out.extend(proc_name_findings(lines, view, module))
    return sorted(out, key=lambda f: (f.line, f.rule, f.message))


def summarise(items: list[Finding]) -> dict[str, int]:
    counts = {rule: 0 for rule in RULES}
    for item in items:
        counts[item.rule] += 1
    return counts


# --------------------------------------------------------------------------
# The pending list

PENDING_FILE = "tools/house_style_pending.txt"
PENDING_HEADER = f"""\
# House-style findings I have not yet settled by hand (#210).
#
# The house-style rules check reads this file while the modules are brought to
# the style. A finding listed here is allowed. A finding that is not listed
# fails the check, and so does a line here that no finding matches any more
# (a stale line). Each line is "<Module>.<Procedure>: <rule>", or
# "<Module>: <rule>" for a finding of the module's own header, and stands for
# one finding. The rules are:
#   {", ".join(RULES)}.
#
# Settle a finding, then delete its line. Do not add a line for new code: new
# code follows the style. I delete this file when it is empty.
#
# Regenerate it with:
#   python3 tools/reformat.py --write-pending src/*.bas test/*.bas demo/*.bas
# Once the file exists, that only removes lines; it refuses to add one.
"""
PENDING_LINE = re.compile(r"^(?P<where>[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?): (?P<rule>[a-z][a-z-]*)$")


def pending_key(module: str, item: Finding) -> str:
    """The line of the pending list that stands for one finding."""
    where = module if item.where == module else f"{module}.{item.where}"
    return f"{where}: {item.rule}"


def pending_text(found: Iterable[tuple[str, Finding]]) -> str:
    """The pending list for these findings (module, finding), in the order given."""
    return PENDING_HEADER + "".join(pending_key(module, item) + "\n" for module, item in found)


class PendingEntry(NamedTuple):
    number: int
    key: str


def read_pending(text: str) -> tuple[list[PendingEntry], list[tuple[int, str]]]:
    """The entries of a pending list, and the lines that are neither a comment nor an entry."""
    entries: list[PendingEntry] = []
    problems: list[tuple[int, str]] = []
    for number, line in enumerate(text.replace("\r\n", "\n").split("\n"), 1):
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        m = PENDING_LINE.match(stripped)
        if m is None:
            problems.append((number, f'not a "<Module>.<Procedure>: <rule>" line: {stripped!r}'))
        elif m.group("rule") not in RULES:
            problems.append((number, f"unknown rule {m.group('rule')!r}"))
        else:
            entries.append(PendingEntry(number, stripped))
    return entries, problems


class PendingResult(NamedTuple):
    unlisted: list[tuple[str, Finding]]     # findings with no line left in the list
    stale: list[PendingEntry]               # lines that no finding matches
    problems: list[tuple[int, str]]         # lines that are not entries


def compare_pending(found: list[tuple[str, Finding]], text: str) -> PendingResult:
    """Hold the findings (module, finding) to the pending list, one line per finding.

    A line covers one finding of its module, procedure and rule, so a second
    finding of that kind in the procedure is not covered until it is listed.
    """
    entries, problems = read_pending(text)
    free = Counter(entry.key for entry in entries)
    unlisted: list[tuple[str, Finding]] = []
    for module, item in found:
        key = pending_key(module, item)
        if free[key] > 0:
            free[key] -= 1
        else:
            unlisted.append((module, item))
    stale: list[PendingEntry] = []
    for entry in reversed(entries):
        if free[entry.key] > 0:
            free[entry.key] -= 1
            stale.append(entry)
    return PendingResult(unlisted, sorted(stale), problems)


def additions(found: list[tuple[str, Finding]], text: str) -> list[str]:
    """The lines that writing the list for these findings would add to the list ``text``."""
    held = Counter(entry.key for entry in read_pending(text)[0])
    wanted = Counter(pending_key(module, item) for module, item in found)
    return sorted((wanted - held).elements())
