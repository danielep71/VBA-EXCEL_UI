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

SECTION_RE = re.compile(r"^\[([a-z-]+)(?:\s+(v[0-9][0-9A-Za-z.\-+]*))?\]$")

BASELINE_SECTION = "baseline"

DECL_RE = re.compile(
    r"^Public\s+(Sub|Function|Property\s+(?:Get|Let|Set)|Enum|Const|Type)\s+(\w+)"
)

IF_RE = re.compile(r"^#If\s+(.+?)\s+Then$")
ELSEIF_RE = re.compile(r"^#ElseIf\s+(.+?)\s+Then$")

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
    __slots__ = ("kind", "name", "text", "line", "path")

    def __init__(self, kind, name, text, line, path):
        self.kind = kind
        self.name = name
        self.text = text
        self.line = line

        # The full stack of enclosing conditionals as
        # ((condition, line, arm_index), ...). Nesting is ordinary here:
        # a Win64 split inside a VBA7 branch produces three declarations of
        # one member, and a model that assumed a flat Then/Else pair would
        # call that a duplicate.
        self.path = path

    def where(self):
        if not self.path:
            return f"line {self.line}, unconditional"
        trail = " > ".join(f"{cond} arm {arm} ({pred})"
                           for cond, _, arm, pred in self.path)
        return f"line {self.line}, {trail}"


def declarations_of_text(text):
    """Return every Public declaration, tagged with its compilation arm."""
    found = []
    stack = []

    pending = None          # (kind, name, line, path, [members])
    terminator = None

    for line, code in logical_lines(text):
        m_if = IF_RE.match(code)
        if m_if:
            # [condition of this block, line, arm index, effective predicate,
            #  conditions of the arms already opened]
            stack.append([m_if.group(1), line, 0, m_if.group(1),
                          [m_if.group(1)]])
            continue

        m_elseif = ELSEIF_RE.match(code)
        if m_elseif:
            if not stack:
                raise ApiError(f"line {line}: #ElseIf without #If")
            frame = stack[-1]
            frame[2] += 1
            frame[3] = effective_predicate(frame[4], m_elseif.group(1))
            frame[4] = frame[4] + [m_elseif.group(1)]
            continue

        if code == "#Else":
            if not stack:
                raise ApiError(f"line {line}: #Else without #If")
            frame = stack[-1]
            frame[2] += 1
            frame[3] = effective_predicate(frame[4], None)
            continue
        if code == "#End If":
            if not stack:
                raise ApiError(f"line {line}: #End If without #If")
            stack.pop()
            continue

        if pending is not None:
            if code == terminator:
                kind, name, at, path, members = pending
                body = (canonical_enum(name, members) if kind == "Enum"
                        else canonical_type(name, members))
                found.append(Declaration(kind, name, body, at, path))
                pending, terminator = None, None
            else:
                pending[4].append(code)
            continue

        m = DECL_RE.match(code)
        if not m:
            continue

        kind = re.sub(r"\s+", " ", m.group(1))
        name = m.group(2)
        path = tuple((cond, ln, arm, pred) for cond, ln, arm, pred, _ in stack)

        if kind in ("Enum", "Type"):
            pending = [kind, name, line, path, []]
            terminator = f"End {kind}"
            continue

        if kind == "Const":
            body = re.sub(r"^Public\s+Const\s+", "", code).strip()
            found.append(Declaration(kind, name, body, line, path))
            continue

        found.append(
            Declaration(kind, name, canonical_procedure(kind, name, code),
                        line, path)
        )

    if pending is not None:
        raise ApiError(f"line {pending[2]}: unterminated {pending[0]} {pending[1]}")
    if stack:
        raise ApiError(f"line {stack[-1][1]}: unterminated #If")

    return found


PTR_RE = re.compile(r"\b(?:LongPtr|LongLong)\b")


def arm_key(path):
    """Render a condition path as stable text, with no line numbers in it.

    The arms a member is declared in are part of its contract. Deleting the
    #Else arm of a VBA7 pair removes the member from every 32-bit build while
    leaving the recorded declaration identical, so the declaration alone is not
    enough to record. Line numbers are excluded deliberately: they churn on
    every edit above, and the contract is which arms exist, not where.
    """
    if not path:
        return ""
    return ">".join(f"{pred}#{arm}" for _, _, arm, pred in path)


def arms_field(paths):
    return " | ".join(sorted(arm_key(p) for p in paths if p))


def effective_predicate(preceding, own):
    """Return what actually has to hold for an arm to compile.

    An arm is reached only when every earlier arm of its block was not taken.
    Recording just the arm's own condition loses that: the #ElseIf VBA7 arm of
    an #If Win64 block is not VBA7, it is VBA7 on a host that is not Win64, and
    changing the leading condition to Mac changes which hosts reach the arm
    while leaving its own condition untouched. The trailing #Else is the same
    problem stated as a word — "Else" names a position, not a condition.
    """
    negated = [f"Not ({cond})" for cond in preceding]
    if own is None:
        return " And ".join(negated)
    return " And ".join(negated + [own])


def normalise_predicate(text):
    """Collapse a conditional predicate to a comparable form."""
    text = re.sub(r"\s+", " ", text).strip().casefold()
    while text.startswith("(") and text.endswith(")"):
        text = text[1:-1].strip()
    return text


def are_complementary(first, second):
    """Return True when two predicates are syntactic negations of each other.

    Declaring a member under #If VBA7 and again under #If Not VBA7 is two
    blocks and one member, and rejecting it as a duplicate was wrong. The test
    is deliberately syntactic: proving arbitrary predicates disjoint is not
    something this tool should attempt, so anything it cannot show complementary
    is reported for a person to restructure or confirm.
    """
    a = normalise_predicate(first)
    b = normalise_predicate(second)

    for outer, inner in ((a, b), (b, a)):
        if outer.startswith("not "):
            if normalise_predicate(outer[4:]) == inner:
                return True
    return False


def exclusivity_finding(name, arms):
    """Return a finding when two arms are not provably mutually exclusive.

    Two declarations are alternatives only when they sit in different arms of
    the *same* directive. Separate #If VBA7 and #If Win64 blocks are not
    alternatives: both conditions hold on a 64-bit VBA7 host, so the second
    declaration is a duplicate the compiler rejects. Folding them into one
    contract hid a build break behind a tidy manifest.
    """
    for i, first in enumerate(arms):
        for second in arms[i + 1:]:
            depth = 0
            limit = min(len(first.path), len(second.path))
            while depth < limit and first.path[depth] == second.path[depth]:
                depth += 1

            if depth == limit:
                return (
                    f"{name}: one declaration is reachable wherever the other "
                    f"is, so they are duplicates rather than alternatives "
                    f"({first.where()} and {second.where()})"
                )

            a_cond, a_line, _, a_pred = first.path[depth]
            b_cond, b_line, _, b_pred = second.path[depth]

            if (a_cond, a_line) == (b_cond, b_line):
                continue                     # different arms of one block

            if are_complementary(a_pred, b_pred):
                continue                     # #If X and #If Not X

            return (
                f"{name}: declared in two separate conditional blocks whose "
                f"predicates are not complementary ({a_pred} at line {a_line} "
                f"and {b_pred} at line {b_line}). A host satisfying both "
                f"compiles two declarations of the same member; write them as "
                f"arms of one block, or as a predicate and its negation\n"
                f"      {first.where()}\n      {second.where()}"
            )
    return None


def fold_conditionals(declarations):
    """Collapse the conditional arms of one member into its logical contract.

    Returns (records, findings).

    A member can be declared in more than one compilation arm, and the arms
    nest: a Win64 split inside a VBA7 branch declares one member three times.
    The rule is therefore stated over the whole set rather than over a Then and
    an Else. Every arm must be either the widest declaration or that same
    declaration with the pointer types narrowed. Arms that are simply identical
    satisfy it too, which is what a Win64 split that does not change any
    signature looks like, and reporting those as a divergence was a false
    positive.

    Anything else is a genuine disagreement between two compilations and is
    reported rather than normalised away.
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
            records.append((kind, arms[0].text, arms_field([arms[0].path])))
            continue

        # Two declarations reachable under the same conditions are a duplicate,
        # not a variant, whether that is twice at top level or twice inside one
        # arm. VBA would reject it; so does this.
        seen = {}
        duplicated = False
        for decl in arms:
            if decl.path in seen:
                duplicated = True
                findings.append(
                    f"{name}: declared twice under the same conditions "
                    f"({seen[decl.path].where()} and {decl.where()})"
                )
            else:
                seen[decl.path] = decl
        if duplicated:
            continue

        overlap = exclusivity_finding(name, arms)
        if overlap:
            findings.append(overlap)
            continue

        pointered = [decl for decl in arms if PTR_RE.search(decl.text)]

        if pointered:
            widest = {decl.text for decl in pointered}
            if len(widest) > 1:
                findings.append(
                    f"{name}: pointer-typed arms declare different contracts\n"
                    + "\n".join(f"      {decl.where()}: {decl.text}"
                                for decl in pointered)
                )
                continue
            wide = pointered[0].text
        else:
            uniform = {decl.text for decl in arms}
            if len(uniform) > 1:
                findings.append(
                    f"{name}: arms differ but none declares a pointer type, so "
                    f"the difference is not a pointer-width variant\n"
                    + "\n".join(f"      {decl.where()}: {decl.text}"
                                for decl in arms)
                )
                continue
            wide = arms[0].text

        allowed = {wide, narrow_pointers(wide)}
        divergent = [decl for decl in arms if decl.text not in allowed]
        if divergent:
            findings.append(
                f"{name}: an arm is neither the widest declaration nor that "
                f"declaration with pointer types narrowed\n"
                f"      widest   : {wide}\n"
                f"      narrowed : {narrow_pointers(wide)}\n"
                + "\n".join(f"      {decl.where()}: {decl.text}"
                            for decl in divergent)
            )
            continue

        records.append((kind, wide, arms_field(decl.path for decl in arms)))

    return records, findings


def records_of_text(text, module):
    """Return (manifest lines, findings) for one module's source text."""
    records, findings = fold_conditionals(declarations_of_text(text))
    lines = sorted(
        f"{module}\t{kind}\t{body}" + (f"\t{arms}" if arms else "")
        for kind, body, arms in records
    )
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
    """Read the manifest into ({section: [lines]}, baseline_version).

    The baseline section carries the release it was captured from in its own
    header, because a frozen contract that does not say what it is frozen at is
    not evidence of anything.
    """
    sections = {name: [] for name, _ in SECTIONS}
    sections[BASELINE_SECTION] = []
    baseline_version = None
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
                if current == BASELINE_SECTION:
                    if not m.group(2):
                        raise ApiError(
                            "the [baseline] section header must name the "
                            "release it was captured from, as [baseline vX.Y.Z]"
                        )
                    baseline_version = m.group(2)
                continue
            if current is None:
                raise ApiError(f"manifest entry before any section header: {line!r}")
            sections[current].append(line)

    if baseline_version is None:
        raise ApiError("the manifest has no [baseline vX.Y.Z] section")

    return ({name: sorted(lines) for name, lines in sections.items()},
            baseline_version)


MANIFEST_HEADER = """\
# Public API manifest for VBA-EXCEL_UI
#
# One canonical declaration per line, generated by tools/vba_api.py and diffed
# by tools/check_repo.py. Editing a line here is how an intentional API change
# is declared; the gate fails on any difference that is not recorded.
#
# Format: <module>\\t<kind>\\t<canonical declaration>[\\t<arms>]
#
# The fourth field lists the compilation arms a conditional member is declared
# in, as <arm predicate>#<arm index> joined by > for nesting. It is absent for
# an unconditional member. Which arms exist is part of the contract: deleting
# the #Else arm of a VBA7 pair removes the member from every 32-bit build while
# leaving the declaration itself identical, and replacing that #Else with
# #ElseIf Mac Then narrows it to one platform. Each arm therefore records its
# own predicate, not the position it happens to occupy.
#
# Passing mode, parameter names and order, Optional status, defaults, types,
# return types and enum values are all part of the recorded contract, so a
# reorder, a ByVal/ByRef flip, a widened return or a renumbered enum member is
# a manifest change and cannot land silently.
#
# A member declared in several compilation arms is recorded once. Arms nest, so
# a Win64 split inside a VBA7 branch is three declarations of one member. Every
# arm must be either the widest declaration or that declaration with pointer
# types narrowed; arms that are simply identical satisfy that too. Anything else
# is a real disagreement between two compilations and is reported.
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
#
# [baseline vX.Y.Z]
#   The [supported] facade as it stood at the named release. It is frozen
#   between releases and rebased only at one, with
#   tools/vba_api.py --rebase-baseline vX.Y.Z. The gate compares [supported]
#   against it, so an in-flight change to the external contract is detected in
#   the branch rather than in the release diff, and a shallow CI checkout can
#   detect it without any Git history.
"""


def render_manifest(sections, baseline_version, baseline_lines):
    out = [MANIFEST_HEADER]
    for name, _ in SECTIONS:
        out.append(f"\n[{name}]\n")
        for line in sections[name]:
            out.append(line + "\n")
    out.append(f"\n[{BASELINE_SECTION} {baseline_version}]\n")
    for line in baseline_lines:
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


def _mutate(old, new, source=None):
    source = BASE if source is None else source
    if old not in source:
        raise ApiError(f"self-test fixture does not contain {old!r}")
    return source.replace(old, new)


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


PROPERTY_BASE = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

Public Property Get UI_Mode() As Long
End Property
"""

NESTED_OK = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

#If VBA7 Then
#If Win64 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#Else
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#End If
#Else
Public Function UI_Frame(ByRef HwndOut As Long) As Boolean
#End If
"""

NESTED_BAD = NESTED_OK.replace(
    "#Else\nPublic Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean\n#End If",
    "#Else\nPublic Function UI_Frame(ByVal HwndOut As LongPtr) As Boolean\n#End If",
)

WIN64_IDENTICAL = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

#If Win64 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#Else
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#End If
"""

ELSEIF_CHAIN = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

#If Win64 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#ElseIf VBA7 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#Else
Public Function UI_Frame(ByRef HwndOut As Long) As Boolean
#End If
"""

ARM_REMOVED = CONDITIONAL_OK.replace(
    "#Else\nPublic Function UI_Frame( _\n    ByRef HwndOut As Long) _\n"
    "    As Boolean\n#End If",
    "#End If",
)

ELSEIF_INSTEAD_OF_ELSE = CONDITIONAL_OK.replace("#Else", "#ElseIf Mac Then")

def _three_arm(lead):
    return (
        'Attribute VB_Name = "M_SELFTEST"\n'
        "Option Explicit\n"
        "\n"
        f"#If {lead} Then\n"
        "Public Sub UI_Elsewhere()\n"
        "End Sub\n"
        "#ElseIf VBA7 Then\n"
        "Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean\n"
        "#Else\n"
        "Public Function UI_Frame(ByRef HwndOut As Long) As Boolean\n"
        "#End If\n"
    )


THREE_ARM_WIN64 = _three_arm("Win64")
THREE_ARM_MAC = _three_arm("Mac")


def _elseif_only(lead):
    """One member, in the #ElseIf arm, and nothing public before it.

    The three-arm fixture cannot prove that an #ElseIf arm records the
    conditions preceding it, because a member in its leading arm changes the
    manifest on its own and satisfies the comparison. Here the leading arm
    declares nothing, so only the member under test can carry the difference.
    """
    return (
        'Attribute VB_Name = "M_SELFTEST"\n'
        "Option Explicit\n"
        "\n"
        f"#If {lead} Then\n"
        "' no public declaration in the leading arm\n"
        "#ElseIf VBA7 Then\n"
        "Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean\n"
        "#End If\n"
    )


def _final_else_only(lead):
    """One member, in the final #Else, and nothing public before it."""
    return (
        'Attribute VB_Name = "M_SELFTEST"\n'
        "Option Explicit\n"
        "\n"
        f"#If {lead} Then\n"
        "' no public declaration in the leading arm\n"
        "#ElseIf VBA7 Then\n"
        "' no public declaration in the intermediate arm\n"
        "#Else\n"
        "Public Function UI_Frame(ByRef HwndOut As Long) As Boolean\n"
        "#End If\n"
    )


ELSEIF_ONLY_WIN64 = _elseif_only("Win64")
ELSEIF_ONLY_MAC = _elseif_only("Mac")
FINAL_ELSE_ONLY_WIN64 = _final_else_only("Win64")
FINAL_ELSE_ONLY_MAC = _final_else_only("Mac")

COMPLEMENTARY_BLOCKS = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

#If VBA7 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#End If

#If Not VBA7 Then
Public Function UI_Frame(ByRef HwndOut As Long) As Boolean
#End If
"""

ELSE_AND_NEGATED_BLOCK = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

#If VBA7 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#Else
Public Function UI_Frame(ByRef HwndOut As Long) As Boolean
#End If

#If Not VBA7 Then
Public Function UI_Frame(ByRef HwndOut As Long) As Boolean
#End If
"""

SEPARATE_BLOCKS = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

#If VBA7 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#End If

#If Win64 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#End If
"""

NESTED_INNER_DUPLICATE = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

#If VBA7 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#If Win64 Then
Public Function UI_Frame(ByRef HwndOut As LongPtr) As Boolean
#End If
#Else
Public Function UI_Frame(ByRef HwndOut As Long) As Boolean
#End If
"""

DUPLICATE_UNCONDITIONAL = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

Public Sub UI_Twice()
End Sub

Public Sub UI_Twice()
End Sub
"""

UNCONDITIONAL_PLUS_ARM = """\
Attribute VB_Name = "M_SELFTEST"
Option Explicit

Public Sub UI_Twice()
End Sub

#If VBA7 Then
Public Sub UI_Twice()
End Sub
#End If
"""

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
    ("standalone member added", "End Function\n",
     "End Function\n\nPublic Sub UI_Extra()\nEnd Sub\n"),
]

PROPERTY_CASES = [
    ("property accessor kind Get to Let",
     "Public Property Get UI_Mode() As Long",
     "Public Property Let UI_Mode(ByVal NewValue As Long)"),
    ("property accessor kind Get to Set",
     "Public Property Get UI_Mode() As Long",
     "Public Property Set UI_Mode(ByVal NewValue As Object)"),
    ("property return type change",
     "Public Property Get UI_Mode() As Long",
     "Public Property Get UI_Mode() As Variant"),
    ("property companion accessor added",
     "End Property\n",
     "End Property\n\nPublic Property Let UI_Mode(ByVal NewValue As Long)\n"
     "End Property\n"),
]

CONDITIONAL_CASES = [
    ("matched VBA7 pointer pair folds to one member", CONDITIONAL_OK, 1, False),
    ("VBA7 pair diverging beyond pointer width is reported",
     CONDITIONAL_BAD, None, True),
    ("nested VBA7 and Win64 arms fold to one member", NESTED_OK, 1, False),
    ("identical Win64 arms are not a divergence", WIN64_IDENTICAL, 1, False),
    ("ElseIf chain folds to one member", ELSEIF_CHAIN, 1, False),
    ("nested arm diverging beyond pointer width is reported",
     NESTED_BAD, None, True),
    ("member declared twice unconditionally is reported",
     DUPLICATE_UNCONDITIONAL, None, True),
    ("member declared unconditionally and in an arm is reported",
     UNCONDITIONAL_PLUS_ARM, None, True),
    ("separate VBA7 and Win64 blocks are duplicates, not alternatives",
     SEPARATE_BLOCKS, None, True),
    ("a declaration inside a nested arm of its own branch is a duplicate",
     NESTED_INNER_DUPLICATE, None, True),
    ("complementary #If X and #If Not X blocks are one member",
     COMPLEMENTARY_BLOCKS, 1, False),
    ("an #Else arm and a separately negated block overlap",
     ELSE_AND_NEGATED_BLOCK, None, True),
    ("a three-arm block folds to one member", THREE_ARM_WIN64, 2, False),
]

# Arms are contract. Each pair is (label, source, other source) whose recorded
# manifests must differ, because the declaration text alone is identical.
ARM_CASES = [
    ("deleting the #Else arm changes the recorded contract",
     CONDITIONAL_OK, ARM_REMOVED),
    ("a nested Win64 split is not the same contract as a flat VBA7 pair",
     CONDITIONAL_OK, NESTED_OK),
    ("replacing #Else with #ElseIf Mac changes the recorded contract",
     CONDITIONAL_OK, ELSEIF_INSTEAD_OF_ELSE),
    (
        "changing a preceding predicate changes an isolated #ElseIf arm",
        ELSEIF_ONLY_WIN64,
        ELSEIF_ONLY_MAC,
    ),
    (
        "changing preceding predicates changes an isolated final #Else arm",
        FINAL_ELSE_ONLY_WIN64,
        FINAL_ELSE_ONLY_MAC,
    ),
    (
        "changing a leading predicate changes a three-arm manifest",
        THREE_ARM_WIN64,
        THREE_ARM_MAC,
    ),
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

    try:
        property_baseline = _records(PROPERTY_BASE)
    except ApiError as exc:
        failures.append(f"property fixture does not parse: {exc}")
        property_baseline = None

    if property_baseline is not None:
        for label, old, new in PROPERTY_CASES:
            try:
                changed = _records(_mutate(old, new, PROPERTY_BASE))
            except ApiError as exc:
                failures.append(f"{label}\n      fixture did not parse: {exc}")
                continue
            if changed == property_baseline:
                failures.append(
                    f"{label}\n      produced an identical manifest, so the "
                    f"gate would not see it\n      records: {changed}"
                )

    for label, left, right in ARM_CASES:
        try:
            a, a_findings = records_of_text(left, "M_SELFTEST")
            b, b_findings = records_of_text(right, "M_SELFTEST")
        except ApiError as exc:
            failures.append(f"{label}\n      fixture did not parse: {exc}")
            continue
        if a_findings or b_findings:
            failures.append(
                f"{label}\n      a fixture was reported: "
                f"{a_findings + b_findings}"
            )
        elif a == b:
            failures.append(
                f"{label}\n      both recorded the same manifest, so the gate "
                f"would not see it\n      records: {a}"
            )

    for label, source, expected_count, expect_finding in CONDITIONAL_CASES:
        try:
            lines, findings = records_of_text(source, "M_SELFTEST")
        except ApiError as exc:
            failures.append(f"{label}\n      fixture did not parse: {exc}")
            continue

        if expect_finding:
            if not findings:
                failures.append(
                    f"{label}\n      the model accepted it silently\n"
                    f"      records: {lines}"
                )
            continue

        if findings:
            failures.append(f"{label}\n      reported: {findings}")
        elif len(lines) != expected_count:
            failures.append(
                f"{label}\n      expected {expected_count} record(s), got "
                f"{len(lines)}\n      records: {lines}"
            )
        elif "LongPtr" not in lines[0]:
            failures.append(
                f"{label}\n      the widest declaration was not the one "
                f"recorded\n      records: {lines}"
            )

    return failures


def run_selftest():
    failures = selftest()
    total = (len(BREAKING_CASES) + len(NEUTRAL_CASES)
             + len(PROPERTY_CASES) + len(CONDITIONAL_CASES) + len(ARM_CASES))
    if not failures:
        print(f"ok   self-test: {total} API-contract rules hold")
        return 0
    print(f"FAIL self-test: {len(failures)} rule(s) broken\n")
    for f in failures:
        print(f"  - {f}")
    return 1


# --------------------------------------------------------------------------
USAGE = """usage:
  vba_api.py --selftest              verify the model against its own fixtures
  vba_api.py --emit                  print the manifest the source would produce
  vba_api.py --write                 regenerate tools/public_api_manifest.txt
  vba_api.py --rebase-baseline vX.Y.Z
                                     freeze the current [supported] facade as
                                     the baseline for a release just made
"""


def main():
    args = sys.argv[1:]

    if args and args[0] == "--selftest":
        return run_selftest()

    if args and args[0] in ("--emit", "--write", "--rebase-baseline"):
        manifest_path = os.path.join(REPO, MANIFEST)

        try:
            recorded, baseline_version = parse_manifest(manifest_path)
        except (OSError, ApiError) as exc:
            print(f"FAIL {MANIFEST}: {exc}", file=sys.stderr)
            return 1

        baseline_lines = recorded[BASELINE_SECTION]

        if args[0] == "--rebase-baseline":
            if len(args) != 2 or not re.match(r"^v[0-9]", args[1]):
                print(USAGE, file=sys.stderr)
                return 2
            baseline_version = args[1]

        sections, findings = surface()
        for f in findings:
            print(f"FAIL {f}", file=sys.stderr)
        if findings:
            return 1

        if args[0] == "--rebase-baseline":
            baseline_lines = sections["supported"]

        text = render_manifest(sections, baseline_version, baseline_lines)

        if args[0] == "--emit":
            sys.stdout.write(text)
            return 0

        with open(manifest_path, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(text)
        print(f"wrote {MANIFEST}"
              + (f" with the baseline frozen at {baseline_version}"
                 if args[0] == "--rebase-baseline" else ""))
        return 0

    print(USAGE, file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main())
