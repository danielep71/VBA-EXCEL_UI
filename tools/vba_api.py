#!/usr/bin/env python3
"""
Canonical public-declaration model for VBA-EXCEL_UI.

The public API manifest used to record ``module | kind | name``. That detects a
member appearing or disappearing and nothing else, so a patch could reorder
parameters, flip ByVal to ByRef, change a default, widen a return type or
renumber an enum while the gate stayed green. This module parses every public
declaration into one canonical line instead, and tools/check_repo.py diffs those
lines against tools/public_api_manifest.txt.

Two contracts are kept apart, because they carry different promises:

  supported       the caller-facing facade in M_EXCEL_UI, covered by Semantic
                  Versioning. Changing one of these lines is a compatibility
                  event and has to be declared as such.

  project-public  helpers and regression seams that are Public only because an
                  Option Private Module project needs them visible across its
                  own modules. They are tracked so a compile-breaking change is
                  never silent, but no external compatibility is claimed for
                  them.

Conditional declarations are folded rather than duplicated. A procedure declared
once under ``#If VBA7 Then`` and again under ``#Else`` is one logical member, so
the VBA7 arm is recorded and the legacy arm is proven to be the same declaration
with LongPtr narrowed to Long. An arm that differs in any other way is reported
rather than normalised away, because that is a real divergence between the two
compilations.

Comments never reach the manifest. Procedure headers in this codebase quote
example calls, and a parser that matched them would record an illustration as
public surface.

--selftest exercises every breaking-change class the gate claims to catch, plus
the formatting-only changes it must ignore. check_repo.py runs it, so the model
is verified on every push rather than only when someone suspects it.
"""

import os
import re
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

MANIFEST = "tools/public_api_manifest.txt"

SUPPORTED_MODULES = ["src/M_EXCEL_UI.bas"]

PROJECT_PUBLIC_MODULES = [
    "src/M_EXCEL_UI_RUNTIME.bas",
    "src/M_EXCEL_UI_SNAPSHOT.bas",
    "src/M_EXCEL_UI_TITLEBAR.bas",
]

SECTIONS = [
    ("supported", SUPPORTED_MODULES),
    ("project-public", PROJECT_PUBLIC_MODULES),
]

SECTION_RE = re.compile(r"^\[([a-z-]+)\]$")

DECL_RE = re.compile(
    r"^Public\s+(Sub|Function|Property\s+(?:Get|Let|Set)|Enum|Const|Type)\s+(\w+)"
)

IF_RE = re.compile(r"^#If\s+(.+?)\s+Then$")

POINTER_NARROWING = [
    (re.compile(r"\bLongPtr\b"), "Long"),
    (re.compile(r"\bLongLong\b"), "Long"),
]


class ApiError(Exception):
    """A declaration could not be modelled, which is a finding, not a crash."""


# --------------------------------------------------------------------------
# LEXICAL HELPERS
# --------------------------------------------------------------------------
def split_code_comment(line):
    """Split a line at the apostrophe that begins a comment, if there is one.

    An apostrophe inside a string literal is data. This is the same rule
    tools/reformat.py applies, kept here so the two tools cannot drift apart
    on what counts as code.

    Returns (code, comment); comment is "" when the line has none.
    """
    in_literal = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_literal = not in_literal
        elif ch == "'" and not in_literal:
            return line[:i], line[i:]
    return line, ""


def logical_lines(text):
    """Yield (line_number, code) with comments removed and continuations joined.

    line_number is the line the logical statement started on, so a finding
    points at the declaration rather than at its last continuation.
    """
    raw = text.replace("\r\n", "\n").split("\n")

    n = 0
    while n < len(raw):
        start = n + 1
        stripped = raw[n].strip()

        # A whole-line comment carries no code, and joining it to the next
        # line would splice a header example into a declaration.
        if stripped.startswith("'"):
            n += 1
            continue

        code, _ = split_code_comment(raw[n])
        code = code.strip()

        while code.endswith("_") and code[:-1].endswith(" "):
            n += 1
            if n >= len(raw):
                raise ApiError(f"line {start}: continuation runs past end of file")
            nxt, _ = split_code_comment(raw[n])
            code = code[:-1].rstrip() + " " + nxt.strip()

        if code:
            yield start, re.sub(r"\s+", " ", code).strip()
        n += 1


def balanced_split(text, sep=","):
    """Split on sep at paren depth zero and outside string literals."""
    parts = []
    depth = 0
    in_literal = False
    current = ""

    for ch in text:
        if ch == '"':
            in_literal = not in_literal
        if not in_literal:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == sep and depth == 0:
                parts.append(current.strip())
                current = ""
                continue
        current += ch

    if current.strip():
        parts.append(current.strip())
    return parts


def narrow_pointers(signature):
    """Rewrite a VBA7 signature into the declaration its #Else arm must carry."""
    for pattern, replacement in POINTER_NARROWING:
        signature = pattern.sub(replacement, signature)
    return signature


# --------------------------------------------------------------------------
# PARAMETER AND SIGNATURE MODEL
# --------------------------------------------------------------------------
def canonical_parameter(text):
    """Normalise one parameter to 'Optional ByVal Name As Type = Default'.

    Passing mode is written out even when the source omits it, because VBA
    defaults to ByRef and an omitted keyword is therefore a contract, not a
    blank. Recording it explicitly means adding the word ByRef later cannot
    read as a change when it is not one, and dropping ByVal cannot read as
    formatting when it is.
    """
    rest = text.strip()

    optional = False
    if re.match(r"^Optional\b", rest, re.I):
        optional = True
        rest = rest[len("Optional"):].strip()

        # ParamArray and Optional are mutually exclusive in VBA; a source that
        # combines them would not compile, so it is a finding here too.
        if re.match(r"^ParamArray\b", rest, re.I):
            raise ApiError(f"parameter {text!r} is both Optional and ParamArray")

    param_array = False
    if re.match(r"^ParamArray\b", rest, re.I):
        param_array = True
        rest = rest[len("ParamArray"):].strip()

    mode = "ByRef"
    m = re.match(r"^(ByVal|ByRef)\b", rest, re.I)
    if m:
        mode = "ByVal" if m.group(1).lower() == "byval" else "ByRef"
        rest = rest[len(m.group(1)):].strip()

    default = None
    parts = balanced_split(rest, "=")
    if len(parts) > 2:
        raise ApiError(f"parameter {text!r} has more than one default")
    if len(parts) == 2:
        rest, default = parts[0].strip(), parts[1].strip()

    m = re.match(r"^(\w+)\s*(\(\s*\))?\s*(?:As\s+(.+))?$", rest, re.I)
    if not m:
        raise ApiError(f"unparsable parameter {text!r}")

    name = m.group(1)
    array = "()" if m.group(2) else ""

    # An untyped parameter is Variant. Writing it out keeps a later explicit
    # 'As Variant' from showing up in the diff as a change.
    vtype = (m.group(3) or "Variant").strip()

    out = ""
    if optional:
        out += "Optional "
    if param_array:
        out += "ParamArray "
    out += f"{mode} {name}{array} As {vtype}"
    if default is not None:
        out += f" = {default}"
    return out


def canonical_procedure(kind, name, code):
    """Build the canonical text of a Sub/Function/Property declaration."""
    open_at = code.find("(")
    if open_at == -1:
        params, tail = "", code[code.find(name) + len(name):]
    else:
        depth = 0
        close_at = None
        in_literal = False
        for i in range(open_at, len(code)):
            ch = code[i]
            if ch == '"':
                in_literal = not in_literal
            elif not in_literal and ch == "(":
                depth += 1
            elif not in_literal and ch == ")":
                depth -= 1
                if depth == 0:
                    close_at = i
                    break
        if close_at is None:
            raise ApiError(f"{name}: unbalanced parentheses in the declaration")
        params = code[open_at + 1:close_at]
        tail = code[close_at + 1:]

    rendered = ", ".join(
        canonical_parameter(p) for p in balanced_split(params) if p
    )

    returns = ""
    m = re.match(r"^\s*As\s+(.+?)\s*$", tail, re.I)
    if m:
        returns = f" As {m.group(1).strip()}"
    elif tail.strip():
        raise ApiError(f"{name}: unexpected text after the parameter list: {tail!r}")

    if kind.startswith("Property") and not returns and kind.endswith("Get"):
        # A Property Get with no As clause returns Variant; say so.
        returns = " As Variant"
    if kind == "Function" and not returns:
        returns = " As Variant"

    return f"{name}({rendered}){returns}"


def canonical_enum(name, members):
    """Build 'Name(A = 0, B = 1)' with implicit values resolved.

    VBA assigns an omitted value as the previous member plus one. Recording the
    resolved number is the point: renumbering an early member silently shifts
    every later one, and that is exactly the break the gate has to see.
    """
    rendered = []
    nxt = 0
    for member in members:
        m = re.match(r"^\[?(\w+)\]?\s*(?:=\s*(.+))?$", member.strip())
        if not m:
            raise ApiError(f"{name}: unparsable enum member {member!r}")
        label = m.group(1)
        if m.group(2) is None:
            value = str(nxt)
        else:
            value = m.group(2).strip()
        try:
            nxt = int(value, 0) + 1
        except ValueError:
            # A member defined from another constant keeps its expression and
            # stops implicit numbering, which VBA would reject anyway.
            nxt = None
        rendered.append(f"{label} = {value}")
    return f"{name}({', '.join(rendered)})"


def canonical_type(name, members):
    rendered = []
    for member in members:
        m = re.match(r"^(\w+)\s*(\([^)]*\))?\s*As\s+(.+)$", member.strip(), re.I)
        if not m:
            raise ApiError(f"{name}: unparsable Type member {member!r}")
        array = m.group(2) or ""
        rendered.append(f"{m.group(1)}{array} As {m.group(3).strip()}")
    return f"{name}({', '.join(rendered)})"


# --------------------------------------------------------------------------
# DECLARATION EXTRACTION
# --------------------------------------------------------------------------
class Declaration:
    __slots__ = ("kind", "name", "text", "line", "cond", "arm")

    def __init__(self, kind, name, text, line, cond, arm):
        self.kind = kind
        self.name = name
        self.text = text
        self.line = line
        self.cond = cond
        self.arm = arm


def declarations_of_text(text):
    """Return every Public declaration, tagged with its compilation arm."""
    found = []
    stack = []

    pending = None          # (kind, name, line, cond, arm, [members])
    terminator = None

    for line, code in logical_lines(text):
        if IF_RE.match(code):
            stack.append([IF_RE.match(code).group(1), line, "then"])
            continue
        if re.match(r"^#ElseIf\s+.+\bThen$", code):
            if not stack:
                raise ApiError(f"line {line}: #ElseIf without #If")
            stack[-1][2] = "elseif"
            continue
        if code == "#Else":
            if not stack:
                raise ApiError(f"line {line}: #Else without #If")
            stack[-1][2] = "else"
            continue
        if code == "#End If":
            if not stack:
                raise ApiError(f"line {line}: #End If without #If")
            stack.pop()
            continue

        if pending is not None:
            if code == terminator:
                kind, name, at, cond, arm, members = pending
                body = (canonical_enum(name, members) if kind == "Enum"
                        else canonical_type(name, members))
                found.append(Declaration(kind, name, body, at, cond, arm))
                pending, terminator = None, None
            else:
                pending[5].append(code)
            continue

        m = DECL_RE.match(code)
        if not m:
            continue

        kind = re.sub(r"\s+", " ", m.group(1))
        name = m.group(2)
        cond = (stack[-1][0], stack[-1][1]) if stack else None
        arm = stack[-1][2] if stack else None

        if kind in ("Enum", "Type"):
            pending = [kind, name, line, cond, arm, []]
            terminator = f"End {kind}"
            continue

        if kind == "Const":
            body = re.sub(r"^Public\s+Const\s+", "", code).strip()
            found.append(Declaration(kind, name, body, line, cond, arm))
            continue

        found.append(
            Declaration(kind, name, canonical_procedure(kind, name, code),
                        line, cond, arm)
        )

    if pending is not None:
        raise ApiError(f"line {pending[2]}: unterminated {pending[0]} {pending[1]}")
    if stack:
        raise ApiError(f"line {stack[-1][1]}: unterminated #If")

    return found


def fold_conditionals(declarations):
    """Collapse a #If/#Else pair of one member into its logical contract.

    Returns (records, findings). A member declared in both arms is emitted once
    from the VBA7 arm; the other arm has to be that same declaration with the
    pointer types narrowed, or the two compilations do not agree and the
    difference is reported instead of hidden.
    """
    records = []
    findings = []

    grouped = {}
    order = []
    for decl in declarations:
        key = (decl.kind, decl.name)
        if key not in grouped:
            grouped[key] = []
            order.append(key)
        grouped[key].append(decl)

    for key in order:
        kind, name = key
        arms = grouped[key]

        if len(arms) == 1:
            records.append((kind, arms[0].text))
            continue

        if len(arms) != 2:
            findings.append(
                f"{name}: declared {len(arms)} times; only one #If/#Else pair "
                f"can be folded into a single contract"
            )
            continue

        first, second = arms
        if first.cond is None or first.cond != second.cond:
            findings.append(
                f"{name}: declared twice outside one conditional block "
                f"(lines {first.line} and {second.line})"
            )
            continue

        wide = next((d for d in arms if d.arm == "then"), None)
        narrow = next((d for d in arms if d.arm == "else"), None)
        if wide is None or narrow is None:
            findings.append(
                f"{name}: conditional arms are not a Then/Else pair "
                f"(lines {first.line} and {second.line})"
            )
            continue

        expected = narrow_pointers(wide.text)
        if narrow.text != expected:
            findings.append(
                f"{name}: the #Else arm is not the #If arm with pointer types "
                f"narrowed\n      #If arm  : {wide.text}\n"
                f"      #Else arm: {narrow.text}\n"
                f"      expected : {expected}"
            )
            continue

        records.append((kind, wide.text))

    return records, findings


def records_of_text(text, module):
    """Return (manifest lines, findings) for one module's source text."""
    records, findings = fold_conditionals(declarations_of_text(text))
    lines = sorted(f"{module}\t{kind}\t{body}" for kind, body in records)
    return lines, [f"{module}: {f}" for f in findings]


# --------------------------------------------------------------------------
# REPOSITORY SURFACE
# --------------------------------------------------------------------------
def read_module(repo, rel):
    with open(os.path.join(repo, rel), "rb") as fh:
        return fh.read().decode("latin-1")


def surface(repo=REPO):
    """Return ({section: [lines]}, findings) for the whole production tree."""
    sections = {}
    findings = []

    for section, modules in SECTIONS:
        lines = []
        for rel in modules:
            module = os.path.basename(rel)[:-4]
            try:
                got, bad = records_of_text(read_module(repo, rel), module)
            except ApiError as exc:
                findings.append(f"{rel}: {exc}")
                continue
            lines += got
            findings += bad
        sections[section] = sorted(lines)

    return sections, findings


def parse_manifest(path):
    """Read the manifest into {section: [lines]}."""
    sections = {name: [] for name, _ in SECTIONS}
    current = None

    with open(path, encoding="utf-8") as fh:
        for raw in fh:
            line = raw.rstrip("\n")
            if not line.strip() or line.lstrip().startswith("#"):
                continue
            m = SECTION_RE.match(line.strip())
            if m:
                current = m.group(1)
                if current not in sections:
                    raise ApiError(f"unknown manifest section [{current}]")
                continue
            if current is None:
                raise ApiError(f"manifest entry before any section header: {line!r}")
            sections[current].append(line)

    return {name: sorted(lines) for name, lines in sections.items()}


MANIFEST_HEADER = """\
# Public API manifest for VBA-EXCEL_UI
#
# One canonical declaration per line, generated by tools/vba_api.py and diffed
# by tools/check_repo.py. Editing a line here is how an intentional API change
# is declared; the gate fails on any difference that is not recorded.
#
# Format: <module>\\t<kind>\\t<canonical declaration>
#
# Passing mode, parameter names and order, Optional status, defaults, types,
# return types and enum values are all part of the recorded contract, so a
# reorder, a ByVal/ByRef flip, a widened return or a renumbered enum member is
# a manifest change and cannot land silently.
#
# A procedure declared under #If VBA7 Then and again under #Else is one member.
# The VBA7 arm is recorded; the gate separately proves the #Else arm is that
# same declaration with LongPtr narrowed to Long.
#
# [supported]
#   The caller-facing facade in M_EXCEL_UI. Covered by Semantic Versioning: a
#   change to any line below is a compatibility event and must be declared as
#   such in CHANGELOG.md and in the pull request.
#
# [project-public]
#   Helpers and regression seams that are Public only so an Option Private
#   Module project can see them across its own modules. Tracked so a
#   compile-breaking change is never silent. No external compatibility is
#   claimed for them, and importing the four modules together remains the
#   supported deployment unit.
"""


def render_manifest(sections):
    out = [MANIFEST_HEADER]
    for name, _ in SECTIONS:
        out.append(f"\n[{name}]\n")
        for line in sections[name]:
            out.append(line + "\n")
    return "".join(out)


# --------------------------------------------------------------------------
# SELF-TEST
# --------------------------------------------------------------------------
BASE = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit
Option Private Module

Public Enum UIVisibility
    UI_LeaveUnchanged = -1
    UI_Hide = 0
    UI_Show = 1
End Enum

Public Function UI_Probe( _
    ByVal Scope As UIVisibility, _
    Optional ByRef FailMsg As String = "") _
    As Boolean
End Function
"""


def _records(text):
    lines, findings = records_of_text(text, "M_SELFTEST")
    if findings:
        raise ApiError("; ".join(findings))
    return lines


def _mutate(old, new):
    if old not in BASE:
        raise ApiError(f"self-test fixture does not contain {old!r}")
    return BASE.replace(old, new)


BREAKING_CASES = [
    ("parameter reorder",
     "ByVal Scope As UIVisibility, _\n    Optional ByRef FailMsg As String = \"\"",
     "Optional ByRef FailMsg As String = \"\", _\n    ByVal Scope As UIVisibility"),
    ("parameter rename", "ByVal Scope As", "ByVal Target As"),
    ("ByVal to ByRef", "ByVal Scope As UIVisibility", "ByRef Scope As UIVisibility"),
    ("parameter type change", "Scope As UIVisibility", "Scope As Long"),
    ("Optional dropped", "Optional ByRef FailMsg", "ByRef FailMsg"),
    ("default value change", 'FailMsg As String = ""', 'FailMsg As String = "x"'),
    ("return type change", "As Boolean\nEnd Function", "As Long\nEnd Function"),
    ("enum value change", "UI_Hide = 0", "UI_Hide = 7"),
    ("enum member removed", "    UI_Show = 1\n", ""),
    ("enum member added", "    UI_Show = 1\n", "    UI_Show = 1\n    UI_Toggle = 2\n"),
    ("member removed", "Public Function UI_Probe(", "Private Function UI_Probe("),
    ("member renamed", "Public Function UI_Probe(", "Public Function UI_Probe2("),
]

NEUTRAL_CASES = [
    ("continuation reflow",
     "Public Function UI_Probe( _\n    ByVal Scope As UIVisibility, _\n"
     "    Optional ByRef FailMsg As String = \"\") _\n    As Boolean",
     "Public Function UI_Probe(ByVal Scope As UIVisibility, "
     "Optional ByRef FailMsg As String = \"\") As Boolean"),
    ("indentation change", "    UI_Hide = 0", "        UI_Hide = 0"),
    ("trailing comment added", "    UI_Hide = 0", "    UI_Hide = 0    'hide it"),
    ("header example added",
     "Public Function UI_Probe(",
     "'   Public Sub UI_NotReal(ByVal A As Long)\n'\nPublic Function UI_Probe("),
]

CONDITIONAL_OK = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

#If VBA7 Then
Public Function UI_Frame( _
    ByRef HwndOut As LongPtr) _
    As Boolean
#Else
Public Function UI_Frame( _
    ByRef HwndOut As Long) _
    As Boolean
#End If
"""

CONDITIONAL_BAD = CONDITIONAL_OK.replace(
    "    ByRef HwndOut As Long) _\n    As Boolean\n#End If",
    "    ByVal HwndOut As Long) _\n    As Boolean\n#End If",
)


def selftest():
    """Return a list of failure descriptions; empty means the model holds."""
    failures = []

    try:
        baseline = _records(BASE)
    except ApiError as exc:
        return [f"baseline fixture does not parse: {exc}"]

    if len(baseline) != 2:
        failures.append(
            f"baseline should record 2 members, recorded {len(baseline)}: {baseline}"
        )

    for label, old, new in BREAKING_CASES:
        try:
            changed = _records(_mutate(old, new))
        except ApiError as exc:
            failures.append(f"{label}\n      fixture did not parse: {exc}")
            continue
        if changed == baseline:
            failures.append(
                f"{label}\n      produced an identical manifest, so the gate "
                f"would not see it\n      records: {changed}"
            )

    for label, old, new in NEUTRAL_CASES:
        try:
            changed = _records(_mutate(old, new))
        except ApiError as exc:
            failures.append(f"{label}\n      fixture did not parse: {exc}")
            continue
        if changed != baseline:
            failures.append(
                f"{label}\n      should normalise identically\n"
                f"      expected: {baseline}\n      produced: {changed}"
            )

    lines, findings = records_of_text(CONDITIONAL_OK, "M_SELFTEST")
    if findings:
        failures.append(
            f"matched VBA7 pair\n      reported: {findings}"
        )
    elif len(lines) != 1 or "LongPtr" not in lines[0]:
        failures.append(
            f"matched VBA7 pair\n      should fold to one LongPtr record\n"
            f"      produced: {lines}"
        )

    _, findings = records_of_text(CONDITIONAL_BAD, "M_SELFTEST")
    if not findings:
        failures.append(
            "divergent VBA7 pair\n      arms differ by more than pointer width "
            "and the model accepted them"
        )

    return failures


def run_selftest():
    failures = selftest()
    total = len(BREAKING_CASES) + len(NEUTRAL_CASES) + 2
    if not failures:
        print(f"ok   self-test: {total} API-contract rules hold")
        return 0
    print(f"FAIL self-test: {len(failures)} rule(s) broken\n")
    for f in failures:
        print(f"  - {f}")
    return 1


# --------------------------------------------------------------------------
USAGE = """usage:
  vba_api.py --selftest   verify the contract model against its own fixtures
  vba_api.py --emit       print the manifest the current source would produce
  vba_api.py --write      regenerate tools/public_api_manifest.txt
"""


def main():
    args = sys.argv[1:]

    if args and args[0] == "--selftest":
        return run_selftest()

    if args and args[0] in ("--emit", "--write"):
        sections, findings = surface()
        for f in findings:
            print(f"FAIL {f}", file=sys.stderr)
        if findings:
            return 1
        text = render_manifest(sections)
        if args[0] == "--emit":
            sys.stdout.write(text)
        else:
            with open(os.path.join(REPO, MANIFEST), "w",
                      encoding="utf-8", newline="\n") as fh:
                fh.write(text)
            print(f"wrote {MANIFEST}")
        return 0

    print(USAGE, file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main())
