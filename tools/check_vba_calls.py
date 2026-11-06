#!/usr/bin/env python3
"""Advisory VBA call checker: resolves calls across the modules of each VBA project.

Each configuration of ``tools/vba-projects.json`` is one VBA project. For each
one, and for each conditional-compilation environment of
``vba_analyze.CONFIGS``, this builds an inventory of the project's procedures
and checks only what it can resolve with certainty:

- unqualified and module-qualified calls to project procedures;
- Private members used from another module;
- unqualified names that are public in more than one standard module;
- argument counts and named arguments, including Optional and ParamArray;
- literal targets of ``Application.Run``, ``OnTime``, ``OnKey``, ``OnAction``,
  ``CallByName`` and ``AddressOf``, including a literal passed to a parameter
  that the called procedure hands to one of those sinks.

It is not a VBA compiler. What it cannot resolve is reported as unknown, never
as a defect, and an empty defect list is a clean verdict only when every module
parsed and every tracked module belongs to a configuration. Compiling and
running the project in Excel stay the authoritative checks.

Exit status: 0 when the analysis is complete and found no defect, 1 when it
found a defect or could not complete, 2 on a usage or input error.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from collections import Counter
from collections.abc import Iterable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, cast

from vba_analyze import CONFIGS, active_lines
from vba_lex import LexicalError

REPO = Path(__file__).resolve().parents[1]
TOOL_NAME = "VBA call checker (advisory)"
MANIFEST = "tools/vba-projects.json"
ENVIRONMENTS = tuple(CONFIGS)
VBA_SUFFIXES = {".bas": "standard", ".cls": "class", ".frm": "form"}

# Stable finding codes. "error" is a definite defect in a gating configuration;
# "warning" is advisory; "failure" means the analysis itself is incomplete.
CODES = {
    "VBA-CALL-001": ("error", "Unqualified call to a project-named procedure that no module declares"),
    "VBA-CALL-002": ("error", "Qualified member that the target module or class does not declare"),
    "VBA-CALL-003": ("error", "Private member used from another module"),
    "VBA-CALL-004": ("error", "Ambiguous unqualified name: public in more than one standard module"),
    "VBA-CALL-005": ("error", "Class name used as an object, but the class has no default instance"),
    "VBA-CALL-010": ("error", "Too few arguments: a required parameter is not supplied"),
    "VBA-CALL-011": ("error", "Too many positional arguments and no ParamArray"),
    "VBA-CALL-012": ("error", "Omitted positional argument for a required parameter"),
    "VBA-CALL-013": ("error", "Named argument does not match any parameter"),
    "VBA-CALL-014": ("error", "Argument supplied twice (repeated name or name already given by position)"),
    "VBA-CALL-015": ("error", "Named argument used for a ParamArray parameter"),
    "VBA-CALL-016": ("error", "Positional argument after a named argument"),
    "VBA-CALL-017": ("error", "Property accessor required by this use is not declared"),
    "VBA-CALL-020": ("error", "Conflicting declarations of one name in a module"),
    "VBA-CALL-030": ("error", "Literal macro or member target does not exist in the project"),
    "VBA-CALL-031": ("warning", "Literal macro target resolves only to a Private procedure"),
    "VBA-CALL-040": ("warning", "Same public name exported by more than one standard module"),
    "VBA-CALL-090": ("failure", "Module could not be analyzed"),
    "VBA-CALL-091": ("failure", "Project coverage is incomplete"),
}

KEYWORDS = frozenset("""
and as byref byval call case const declare dim do each else elseif empty end enum erase
error event exit explicit false for friend function get global gosub goto if implements in
is let like lib loop mod new next not nothing null on optional or paramarray preserve
private property public raiseevent redim rem resume return select set static step stop sub
then to true type typeof until wend while with withevents xor addressof imp eqv option
attribute base compare text binary module ptrsafe alias
""".split())

STATEMENT_KEYWORDS = frozenset("""
print open close input line write get put seek lock unlock kill name mkdir rmdir chdir
chdrive filecopy load unload beep randomize date time sendkeys appactivate savesetting
deletesetting reset width mid
""".split())

TOKEN_RE = re.compile(
    r"""
    (?P<ws>\s+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<date>\#\s*\d[\d/\-:. ]*(?:[AaPp][Mm])?\s*\#)
  | (?P<num>&[HhOo][0-9A-Fa-f]+&?|\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?[%&!#@^]?)
  | (?P<op>:=|<>|<=|>=|[-+*/\\^&=<>(),.:;!])
  | (?P<id>[A-Za-z_][A-Za-z0-9_]*[$%&@]?)
  | (?P<other>.)
    """,
    re.VERBOSE,
)


@dataclass(frozen=True)
class Tok:
    kind: str
    text: str
    # A "." after whitespace that follows the first identifier of a statement: "Check .Name = a, m"
    # passes a With member to Check, while "Check.Name = a" assigns a member of Check.
    spaced: bool = field(default=False, compare=False)

    @property
    def low(self) -> str:
        return self.text.casefold().rstrip("$%&@")

    def is_id(self) -> bool:
        return self.kind == "id"

    def is_op(self, *values: str) -> bool:
        return self.kind == "op" and self.text in values

    def is_kw(self, *words: str) -> bool:
        return self.kind == "id" and self.low in words

    def continues_chain(self) -> bool:
        """True for the ``.`` or ``!`` of a member access, false for a spaced ``.`` that starts a With member."""
        return self.is_op("!") or (self.is_op(".") and not self.spaced)


def tokenize(code: str) -> list[Tok]:
    """Tokens without whitespace; a "." keeps whether whitespace came before it."""
    tokens = []
    spaced = False
    for match in TOKEN_RE.finditer(code):
        kind = match.lastgroup or "other"
        if kind == "ws":
            spaced = True
            continue
        tokens.append(Tok(kind, match.group(), spaced and match.group() == "."))
        spaced = False
    return tokens


def split_top(tokens: list[Tok], separator: str) -> list[list[Tok]]:
    """Split on ``separator`` outside parentheses."""
    parts: list[list[Tok]] = [[]]
    depth = 0
    for token in tokens:
        if token.is_op("("):
            depth += 1
        elif token.is_op(")"):
            depth -= 1
        if depth == 0 and token.is_op(separator):
            parts.append([])
            continue
        parts[-1].append(token)
    return parts


def matching_paren(tokens: list[Tok], start: int) -> int:
    depth = 0
    for index in range(start, len(tokens)):
        if tokens[index].is_op("("):
            depth += 1
        elif tokens[index].is_op(")"):
            depth -= 1
            if depth == 0:
                return index
    return -1


def is_lvalue(tokens: list[Tok]) -> bool:
    """True when ``tokens`` form a member chain such as ``a``, ``.b(1)`` or ``a.b(i).c``."""
    index = 1 if tokens and tokens[0].is_op(".") else 0
    if index >= len(tokens):
        return False
    while index < len(tokens):
        if not tokens[index].is_id():
            return False
        index += 1
        if index < len(tokens) and tokens[index].is_op("("):
            close = matching_paren(tokens, index)
            if close < 0:
                return False
            index = close + 1
        if index == len(tokens):
            return True
        if not tokens[index].continues_chain():
            return False
        index += 1
    return False


def literal_value(token: Tok) -> str | None:
    if token.kind == "str":
        return token.text[1:-1].replace('""', '"')
    return None


# ---------------------------------------------------------------------------
# Inventory
# ---------------------------------------------------------------------------


@dataclass
class Param:
    name: str
    optional: bool = False
    paramarray: bool = False
    default: str | None = None
    type: str | None = None
    byval: bool = False


@dataclass
class Proc:
    module: str
    name: str
    accessor: str  # sub | function | get | let | set | declare
    visibility: str
    params: list[Param]
    line: int
    return_type: str | None = None
    statements: list[tuple[int, list[Tok]]] = field(default_factory=list)
    locals: dict[str, tuple[str | None, str]] = field(default_factory=dict)  # name -> (type, kind)
    assigned: set[str] | None = None  # names this procedure may assign; see assigned_names()

    def signature(self) -> str:
        parts = []
        for param in self.params:
            text = ("ParamArray " if param.paramarray else "Optional " if param.optional else "") + param.name
            if param.default is not None:
                text += " = " + param.default
            parts.append(text)
        return f"{self.name}({', '.join(parts)})"


@dataclass
class Member:
    name: str
    kind: str  # proc | var | const | enum | enum_member | type | event
    visibility: str
    line: int
    type: str | None = None
    value: str | None = None
    accessors: dict[str, Proc] = field(default_factory=dict)


@dataclass
class Module:
    path: str
    name: str
    kind: str
    option_private: bool = False
    predeclared: bool = False
    implements: list[str] = field(default_factory=list)
    members: dict[str, Member] = field(default_factory=dict)
    procs: list[Proc] = field(default_factory=list)
    withevents: set[str] = field(default_factory=set)
    failures: list[dict[str, Any]] = field(default_factory=list)
    findings: list[dict[str, Any]] = field(default_factory=list)


def parse_params(tokens: list[Tok]) -> list[Param]:
    params: list[Param] = []
    if not tokens:
        return params
    for part in split_top(tokens, ","):
        words = part
        param = Param(name="")
        index = 0
        while index < len(words) and words[index].is_kw("optional", "byval", "byref", "paramarray"):
            if words[index].is_kw("optional"):
                param.optional = True
            if words[index].is_kw("paramarray"):
                param.paramarray = True
            if words[index].is_kw("byval"):
                param.byval = True
            index += 1
        if index >= len(words) or not words[index].is_id():
            raise ValueError("unparsed parameter: " + " ".join(t.text for t in part))
        param.name = words[index].text.rstrip("$%&@")
        index += 1
        if index < len(words) and words[index].is_op("("):
            close = matching_paren(words, index)
            index = close + 1 if close >= 0 else len(words)
        if index < len(words) and words[index].is_kw("as"):
            type_end = index + 1
            while type_end < len(words) and not words[type_end].is_op("="):
                type_end += 1
            param.type = "".join(t.text for t in words[index + 1:type_end])
            index = type_end
        if index < len(words) and words[index].is_op("="):
            param.default = " ".join(t.text for t in words[index + 1:])
            param.optional = True
        params.append(param)
    return params


def constant_value(part: list[Tok]) -> str | None:
    """The string a ``Const`` declaration assigns, when its initializer joins only string literals.

    ``"UI_" & "Tick"`` is ``UI_Tick``. Any other initializer, even one that starts with a literal,
    gives None, so the constant stays unknown rather than resolving to a prefix of its value.
    """
    equals = next((i for i, t in enumerate(part) if t.is_op("=")), None)
    if equals is None or equals + 1 >= len(part):
        return None
    pieces = []
    for piece in split_top(part[equals + 1:], "&"):
        value = literal_value(piece[0]) if len(piece) == 1 else None
        if value is None:
            return None
        pieces.append(value)
    return "".join(pieces)


def type_after_as(tokens: list[Tok], start: int) -> str | None:
    if start < len(tokens) and tokens[start].is_kw("as"):
        index = start + 1
        if index < len(tokens) and tokens[index].is_kw("new"):
            index += 1
        text = []
        while index < len(tokens) and (tokens[index].is_id() or tokens[index].is_op(".")):
            text.append(tokens[index].text)
            index += 1
        return "".join(text) or None
    return None


def parse_header(tokens: list[Tok]) -> tuple[str, str, str, list[Tok], str | None] | None:
    """Return (visibility, accessor, name, param tokens, return type) for a procedure header."""
    index = 0
    visibility = "public"
    if index < len(tokens) and tokens[index].is_kw("public", "private", "friend", "global"):
        visibility = tokens[index].low.replace("global", "public")
        index += 1
    if index < len(tokens) and tokens[index].is_kw("static"):
        index += 1
    if index >= len(tokens):
        return None
    if tokens[index].is_kw("sub", "function"):
        accessor = tokens[index].low
        index += 1
    elif tokens[index].is_kw("property") and index + 1 < len(tokens) and tokens[index + 1].is_kw("get", "let", "set"):
        accessor = tokens[index + 1].low
        index += 2
    else:
        return None
    if index >= len(tokens) or not tokens[index].is_id():
        return None
    name = tokens[index].text.rstrip("$%&@")
    index += 1
    params: list[Tok] = []
    if index < len(tokens) and tokens[index].is_op("("):
        close = matching_paren(tokens, index)
        if close < 0:
            raise ValueError("unbalanced parameter list")
        params = tokens[index + 1:close]
        index = close + 1
    return visibility, accessor, name, params, type_after_as(tokens, index)


def parse_declare(tokens: list[Tok]) -> Proc | None:
    index = 0
    visibility = "public"
    if tokens and tokens[0].is_kw("public", "private", "global"):
        visibility = tokens[0].low.replace("global", "public")
        index = 1
    if index >= len(tokens) or not tokens[index].is_kw("declare"):
        return None
    index += 1
    if index < len(tokens) and tokens[index].is_kw("ptrsafe"):
        index += 1
    if index + 1 >= len(tokens) or not tokens[index].is_kw("sub", "function"):
        return None
    name = tokens[index + 1].text.rstrip("$%&@")
    index += 2
    while index < len(tokens) and not tokens[index].is_op("("):
        index += 1
    params: list[Param] = []
    if index < len(tokens):
        close = matching_paren(tokens, index)
        params = parse_params(tokens[index + 1:close])
    return Proc("", name, "declare", visibility, params, 0)


def statements_of(code: str) -> list[list[Tok]]:
    """Split one logical line into statements; drop a leading line label."""
    tokens = tokenize(code)
    if len(tokens) >= 2 and tokens[0].is_id() and tokens[1].is_op(":") and not tokens[0].is_kw(*KEYWORDS):
        tokens = tokens[2:]
    elif tokens and tokens[0].kind == "num":
        tokens = tokens[1:]
    return [with_argument(part) for part in split_top(tokens, ":") if part]


def with_argument(tokens: list[Tok]) -> list[Tok]:
    """Keep the spaced mark only on a "." right after the statement's first identifier.

    Elsewhere whitespace before "." carries no meaning here: a line continuation, joined as a space,
    may split a member access, and "x = a .b" is not valid VBA.
    """
    return [Tok(t.kind, t.text, t.spaced and i == 1 and tokens[0].is_id()) if t.spaced else t
            for i, t in enumerate(tokens)]


def add_member(module: Module, member: Member, accessor: str | None = None, proc: Proc | None = None) -> None:
    key = member.name.casefold()
    existing = module.members.get(key)
    if existing is None:
        module.members[key] = member
        if proc is not None:
            member.accessors[accessor or proc.accessor] = proc
        return
    property_pair = (
        existing.kind == "proc" and member.kind == "proc" and proc is not None
        and proc.accessor in {"get", "let", "set"}
        and all(a in {"get", "let", "set"} for a in existing.accessors)
        and proc.accessor not in existing.accessors
    )
    if property_pair and proc is not None:
        existing.accessors[proc.accessor] = proc
        return
    module.findings.append(finding(
        "VBA-CALL-020", module, None, member.line, member.name,
        f"{member.name} is declared at line {existing.line} and again at line {member.line}.",
    ))


def parse_module(  # noqa: C901 - one pass over every module-level statement form
        path: str, text: str, units: Iterable[tuple[int, str]]) -> Module:
    """Build a module's inventory from ``units``, its active logical lines in one environment."""
    kind = VBA_SUFFIXES[Path(path).suffix.casefold()]
    name_match = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.MULTILINE)
    module = Module(path, name_match.group(1) if name_match else Path(path).stem, kind)
    module.predeclared = kind == "form" or bool(
        re.search(r"^Attribute VB_PredeclaredId = True", text, re.MULTILINE | re.IGNORECASE))
    start = 0
    if name_match:
        start = text[:name_match.start()].count("\n")
    current: Proc | None = None
    block: str | None = None
    enum_visibility = "public"
    for first, code in units:
        if first <= start:
            continue
        for tokens in statements_of(code):
            head = tokens[0]
            if head.is_kw("attribute"):
                continue
            if block:
                if head.is_kw("end") and len(tokens) > 1 and tokens[1].is_kw(block):
                    block = None
                elif block == "enum" and head.is_id():
                    add_member(module, Member(head.text, "enum_member", enum_visibility, first))
                continue
            if current is not None:
                if head.is_kw("end") and len(tokens) > 1 and tokens[1].is_kw("sub", "function", "property"):
                    current = None
                    continue
                try:
                    header = parse_header(tokens)
                except ValueError:
                    header = None
                if header is not None:
                    module.failures.append(finding(
                        "VBA-CALL-090", module, current.name, first, header[2],
                        f"{current.name} (line {current.line}) is not closed before {header[2]}."))
                    current = None
                else:
                    current.statements.append((first, tokens))
                    continue
            # Module level
            if head.is_kw("option"):
                if len(tokens) > 2 and tokens[1].is_kw("private") and tokens[2].is_kw("module"):
                    module.option_private = True
                continue
            if head.is_kw("implements") and len(tokens) > 1:
                module.implements.append(tokens[1].text)
                continue
            declare = parse_declare(tokens)
            if declare is not None:
                declare.module, declare.line = module.name, first
                add_member(module, Member(declare.name, "proc", declare.visibility, first), "declare", declare)
                continue
            try:
                header = parse_header(tokens)
            except ValueError as error:
                module.failures.append(finding("VBA-CALL-090", module, None, first, None, str(error)))
                continue
            if header is not None:
                visibility, accessor, name, param_tokens, return_type = header
                try:
                    params = parse_params(param_tokens)
                except ValueError as error:
                    module.failures.append(finding("VBA-CALL-090", module, name, first, name, str(error)))
                    params = []
                current = Proc(module.name, name, accessor, visibility, params, first, return_type)
                for param in params:
                    current.locals[param.name.casefold()] = (param.type, "param")
                module.procs.append(current)
                add_member(module, Member(name, "proc", visibility, first, return_type), accessor, current)
                continue
            parse_module_declaration(module, tokens, first)
            lowered = [t.low for t in tokens[:3]]
            if "type" in lowered or "enum" in lowered:
                for index, token in enumerate(tokens[:3]):
                    if token.is_kw("type", "enum") and index + 1 < len(tokens):
                        block = token.low
                        enum_visibility = "private" if tokens[0].is_kw("private") else "public"
                        break
    if current is not None:
        module.failures.append(finding("VBA-CALL-090", module, current.name, current.line, current.name,
                                       f"{current.name} is not closed."))
    if block is not None:
        module.failures.append(finding("VBA-CALL-090", module, None, None, None, f"{block} block is not closed."))
    for proc in module.procs:
        collect_locals(proc)
    return module


def parse_module_declaration(module: Module, tokens: list[Tok], line: int) -> None:
    head = tokens[0]
    if not head.is_kw("public", "private", "dim", "global", "const", "friend", "type", "enum", "event"):
        return
    visibility = "private" if head.is_kw("private", "dim") else "public"
    index = 1 if head.is_kw("public", "private", "dim", "global", "friend") else 0
    if index < len(tokens) and tokens[index].is_kw("type", "enum"):
        if index + 1 < len(tokens):
            add_member(module, Member(tokens[index + 1].text, tokens[index].low, visibility, line))
        return
    if index < len(tokens) and tokens[index].is_kw("event"):
        if index + 1 < len(tokens):
            add_member(module, Member(tokens[index + 1].text, "event", "public", line))
        return
    is_const = index < len(tokens) and tokens[index].is_kw("const")
    if is_const:
        index += 1
        if head.is_kw("const"):
            visibility = "private"
    for part in split_top(tokens[index:], ","):
        withevents = bool(part) and part[0].is_kw("withevents")
        if withevents:
            part = part[1:]
        if not part or not part[0].is_id():
            continue
        name = part[0].text.rstrip("$%&@")
        position = 1
        if position < len(part) and part[position].is_op("("):
            position = matching_paren(part, position) + 1
        declared_type = type_after_as(part, position)
        value = None
        if is_const:
            value = constant_value(part)
        if withevents:
            module.withevents.add(name.casefold())
        add_member(module, Member(name, "const" if is_const else "var", visibility, line, declared_type, value))


def collect_locals(proc: Proc) -> None:
    for _line, tokens in proc.statements:
        if not tokens[0].is_kw("dim", "static", "const", "redim"):
            continue
        rest = tokens[1:]
        if rest and rest[0].is_kw("preserve"):
            rest = rest[1:]
        if tokens[0].is_kw("static") and rest and rest[0].is_kw("sub", "function", "property"):
            continue
        for part in split_top(rest, ","):
            if not part or not part[0].is_id():
                continue
            name = part[0].text.rstrip("$%&@").casefold()
            position = 1
            if position < len(part) and part[position].is_op("("):
                position = matching_paren(part, position) + 1
            value = None
            if tokens[0].is_kw("const"):
                value = constant_value(part)
            existing = proc.locals.get(name)
            if existing is None or tokens[0].is_kw("dim", "static", "const"):
                proc.locals[name] = (type_after_as(part, position), "const:" + (value or "") if value else "local")


# A statement that starts with one of these may write any variable it names.
WRITING_STATEMENTS = frozenset({"erase", "get", "input", "line", "lset", "mid", "redim", "rset"})

# VBA library functions that only read their arguments. A variable passed to any
# other procedure that the checker cannot show takes it ByVal may be changed.
READ_ONLY_FUNCTIONS = frozenset("""
abs asc ascb ascw cbool cbyte ccur cdate cdbl cdec chr chrb chrw cint clng clnglng clngptr csng
cstr cvar cverr format formatcurrency formatdatetime formatnumber formatpercent hex instr instrb
instrrev isarray isdate isempty iserror ismissing isnull isnumeric isobject join lbound lcase left
leftb len lenb ltrim mid midb oct replace right rightb rtrim sgn space split str strcomp strconv
string strreverse trim typename ubound ucase val vartype
""".split())


def assigned_names(proc: Proc) -> set[str]:
    """Names ``proc`` may assign: an over-approximation, so that a name absent from it is never written.

    A name counts as assigned when it is the target of an assignment (also after Set, Let, Then or
    Else), a For or For Each variable, or any name in a statement that can write its operands, such
    as Input, Get, Mid or ReDim.
    """
    if proc.assigned is None:
        names: set[str] = set()
        for _line, tokens in proc.statements:
            if tokens[0].is_kw(*WRITING_STATEMENTS):
                names.update(t.low for t in tokens if t.is_id())
            for index, token in enumerate(tokens):
                if not token.is_id() or index + 1 >= len(tokens):
                    continue
                previous = tokens[index - 1] if index else None
                following = tokens[index + 1]
                starts = previous is None or previous.is_kw("set", "let", "then", "else", "for", "each")
                if starts and (following.is_op("=", "(") or following.is_kw("in")):
                    names.add(token.low)
        proc.assigned = names
    return proc.assigned


def bind_arguments(params: list[Param], args: list[tuple[str | None, list[Tok]]]) -> dict[int, list[Tok]]:
    """Map each parameter index to the argument a call supplies for it; omitted arguments are absent."""
    paramarray = next((i for i, p in enumerate(params) if p.paramarray), len(params))
    bound: dict[int, list[Tok]] = {}
    position = 0
    for name, tokens in args:
        if name is None:
            if position < paramarray and tokens:
                bound[position] = tokens
            position += 1
            continue
        index = next((i for i, p in enumerate(params) if p.name.casefold() == name.casefold()), None)
        if index is not None and index < paramarray:
            bound[index] = tokens
    return bound


def parameter_default(param: Param) -> str | None:
    """The string literal a parameter defaults to, or None."""
    if param.default is None:
        return None
    tokens = tokenize(param.default)
    return literal_value(tokens[0]) if len(tokens) == 1 else None


# ---------------------------------------------------------------------------
# Resolution and checks
# ---------------------------------------------------------------------------


def finding(code: str, module: Module | None, procedure: str | None, line: int | None,
            target: str | None, message: str, **extra: Any) -> dict[str, Any]:
    severity, title = CODES[code]
    item = {
        "code": code,
        "severity": severity,
        "title": title,
        "path": module.path if module else None,
        "module": module.name if module else None,
        "procedure": procedure,
        "line": line,
        "target": target,
        "message": message,
    }
    item.update(extra)
    return item


@dataclass
class Project:
    name: str
    modules: dict[str, Module]  # casefold name -> module
    patterns: list[re.Pattern[str]]
    macro_functions: set[str]
    complete: bool

    def standard_modules(self) -> list[Module]:
        return [m for m in self.modules.values() if m.kind == "standard"]

    def is_project_name(self, name: str) -> bool:
        return any(pattern.search(name) for pattern in self.patterns)

    def global_candidates(
            self, name: str, exclude: Module | None) -> tuple[list[tuple[Module, Member]], list[tuple[Module, Member]]]:
        """Public and private declarations of ``name`` outside ``exclude`` that VBA treats as global."""
        public: list[tuple[Module, Member]] = []
        private: list[tuple[Module, Member]] = []
        key = name.casefold()
        for module in self.modules.values():
            if module is exclude:
                continue
            member = module.members.get(key)
            if member is None:
                continue
            global_kind = module.kind == "standard" or member.kind in {"enum", "enum_member"}
            if not global_kind:
                continue
            (private if member.visibility == "private" else public).append((module, member))
        return public, private


# (module, procedure), both casefolded -> {parameter (casefolded): the sink its value reaches}
MacroParameters = dict[tuple[str, str], dict[str, str]]

# Stands for ThisWorkbook.Name while a concatenated macro name is evaluated.
THIS_WORKBOOK = "\0"


class Analyzer:
    def __init__(self, project: Project, macro_parameters: MacroParameters | None = None):
        self.project = project
        # Parameters whose value a procedure hands to a macro sink (OnAction, Application.Run, ...).
        # A literal passed to one at a resolved call is checked as a macro target.
        self.macro_parameters: MacroParameters = {} if macro_parameters is None else macro_parameters
        self.findings: list[dict[str, Any]] = []
        self.stats: Counter[str] = Counter()
        self.unknown: Counter[str] = Counter()
        self.entry_points: list[dict[str, Any]] = []
        # Parameter names each procedure passes by reference; see passed_by_reference().
        self.by_reference: dict[int, set[str]] = {}

    # -- helpers ----------------------------------------------------------
    def report(self, code: str, module: Module, proc: Proc | None, line: int, target: str, message: str,
               **extra: Any) -> None:
        item = finding(code, module, proc.name if proc else None, line, target, message, **extra)
        if code in {"VBA-CALL-001", "VBA-CALL-002", "VBA-CALL-030"} and not self.project.complete:
            item.update(severity="unknown", resolution="coverage-incomplete")
        self.findings.append(item)

    def type_module(self, type_name: str | None) -> Module | None:
        if not type_name:
            return None
        return self.project.modules.get(type_name.split(".")[-1].casefold()) if (
            "." not in type_name or type_name.split(".")[0].casefold() == "vbaproject") else None

    def resolve_name(self, name: str, module: Module, proc: Proc | None) -> tuple[str, Any]:
        key = name.casefold()
        if key == "me" and module.kind in {"class", "form"}:
            return "module", module
        if proc is not None:
            if key in proc.locals:
                return "local", proc.locals[key]
            if key == proc.name.casefold():
                return "self", proc
        member = module.members.get(key)
        if member is not None:
            return "member", (module, member)
        public, private = self.project.global_candidates(name, module)
        if len({m.name for m, _ in public}) > 1:
            return "ambiguous", public
        if public:
            return "member", public[0]
        if key in self.project.modules:
            return "module", self.project.modules[key]
        if private:
            return "private", private
        return "external", None

    # -- argument checking ------------------------------------------------
    def split_args(self, tokens: list[Tok]) -> list[tuple[str | None, list[Tok]]]:
        if not tokens:
            return []
        args: list[tuple[str | None, list[Tok]]] = []
        for part in split_top(tokens, ","):
            if len(part) >= 2 and part[0].is_id() and part[1].is_op(":="):
                args.append((part[0].text, part[2:]))
            else:
                args.append((None, part))
        return args

    def check_args(self, target: Proc, params: list[Param], args: list[tuple[str | None, list[Tok]]],
                   module: Module, proc: Proc | None, line: int, label: str) -> None:
        self.stats["argument-checks"] += 1
        paramarray = next((i for i, p in enumerate(params) if p.paramarray), None)
        fixed = params[:paramarray] if paramarray is not None else params
        covered: set[int] = set()
        seen_named = False
        positional = 0
        for name, tokens in args:
            if name is None:
                if seen_named:
                    self.report("VBA-CALL-016", module, proc, line, label,
                                f"{label}: positional argument after a named argument.",
                                signature=target.signature())
                    return
                if positional >= len(fixed) and paramarray is None:
                    self.report("VBA-CALL-011", module, proc, line, label,
                                f"{label} accepts at most {len(fixed)} argument(s); "
                                f"{len([a for a in args if a[0] is None])} supplied.",
                                signature=target.signature())
                    return
                if positional < len(fixed):
                    if tokens:
                        covered.add(positional)
                    elif not fixed[positional].optional:
                        self.report("VBA-CALL-012", module, proc, line, label,
                                    f"{label}: required argument {fixed[positional].name!r} is omitted.",
                                    signature=target.signature())
                positional += 1
                continue
            seen_named = True
            index = next((i for i, p in enumerate(params) if p.name.casefold() == name.casefold()), None)
            if index is None:
                self.report("VBA-CALL-013", module, proc, line, label, f"{label} has no parameter named {name!r}.",
                            signature=target.signature())
                continue
            if params[index].paramarray:
                self.report("VBA-CALL-015", module, proc, line, label,
                            f"{label}: ParamArray {name!r} cannot be passed by name.",
                            signature=target.signature())
                continue
            if index in covered:
                self.report("VBA-CALL-014", module, proc, line, label, f"{label}: argument {name!r} is supplied twice.",
                            signature=target.signature())
                continue
            covered.add(index)
        missing = [p.name for i, p in enumerate(fixed) if not p.optional and i not in covered]
        if missing:
            self.report("VBA-CALL-010", module, proc, line, label,
                        f"{label}: required argument(s) {', '.join(missing)} not supplied.",
                        signature=target.signature())

    def check_member_call(self, owner: Module, member: Member, args: list[tuple[str | None, list[Tok]]] | None,
                          context: str, module: Module, proc: Proc | None, line: int, label: str) -> Module | None:
        """Validate a resolved member use; return the project module its value refers to, if known."""
        if member.kind != "proc":
            self.stats["resolved-data"] += 1
            return self.type_module(member.type)
        accessors = member.accessors
        self.stats["resolved-calls"] += 1
        if args is None:
            args = []  # A procedure named without parentheses is still called, with no arguments
        target: Proc | None
        if "sub" in accessors or "function" in accessors or "declare" in accessors:
            target = accessors.get("function") or accessors.get("sub") or accessors["declare"]
            if args is not None:
                self.check_args(target, target.params, args, module, proc, line, label)
                self.check_macro_arguments(target, args, module, proc, line)
            return self.type_module(target.return_type)
        accessor: Proc | None
        if context == "assign-set":
            accessor = accessors.get("set")
        elif context == "assign":
            accessor = accessors.get("let")
        else:
            accessor = accessors.get("get")
        if accessor is None:
            wanted = {"assign-set": "Property Set", "assign": "Property Let"}.get(context, "Property Get")
            self.report("VBA-CALL-017", module, proc, line, label,
                        f"{label} is used in a way that needs {wanted}, which {owner.name} does not declare.",
                        declared=sorted(accessors))
            return None
        params = accessor.params[:-1] if context.startswith("assign") else accessor.params
        if args is not None:
            self.check_args(accessor, params, args, module, proc, line, label)
        return self.type_module(accessor.return_type)

    def check_recursive_call(self, target: Proc, args: list[tuple[str | None, list[Tok]]] | None, context: str,
                             statement: bool, module: Module, proc: Proc | None, line: int, label: str) -> None:
        """Validate a procedure's use of its own name when that use is a recursive call.

        A Sub naming itself always calls itself. Inside a Function or Property Get, the bare name is the
        return value (assigned or read), so only a statement call, a Call statement or a parenthesised use
        is a recursive call. A parenthesised use of a parameterless function may index its array result,
        so it is not checked. Property Let and Set procedures are left alone.
        """
        if target.accessor == "sub":
            recursive = True
        elif target.accessor in {"function", "get"} and not context.startswith("assign"):
            recursive = statement or context == "call" or (args is not None and bool(target.params))
        else:
            recursive = False
        if not recursive:
            return
        self.stats["resolved-calls"] += 1
        self.check_args(target, target.params, args or [], module, proc, line, label)

    @staticmethod
    def lacks_default_instance(entity: Module, name: str) -> bool:
        """True for a class named directly (not through Me) whose VB_PredeclaredId is False."""
        return entity.kind == "class" and not entity.predeclared and name.casefold() != "me"

    # -- macro and member targets -------------------------------------------
    def literal_target(self, tokens: list[Tok], module: Module, proc: Proc | None) -> str | None:
        if not tokens:
            return None
        if len(tokens) == 1:
            value = literal_value(tokens[0])
            if value is not None:
                return value
            if tokens[0].is_id():
                kind, entity = self.resolve_name(tokens[0].text, module, proc)
                if kind == "local" and entity[1].startswith("const:"):
                    return cast(str, entity[1][6:])
                if kind == "member" and entity[1].kind == "const" and entity[1].value is not None:
                    return cast(str, entity[1].value)
        if tokens[0].is_id() and tokens[0].text.casefold() in self.project.macro_functions:
            if len(tokens) >= 3 and tokens[1].is_op("(") and matching_paren(tokens, 1) == len(tokens) - 1:
                return self.literal_target(tokens[2:-1], module, proc)
        return None

    @staticmethod
    def is_this_workbook_name(tokens: list[Tok]) -> bool:
        return (len(tokens) == 3 and tokens[0].is_kw("thisworkbook") and tokens[1].is_op(".")
                and tokens[2].is_kw("name"))

    @staticmethod
    def strip_this_workbook(text: str) -> str | None:
        """Remove a leading ``'<ThisWorkbook.Name>'!`` from a macro name; None if the workbook stays unresolved."""
        for prefix in (f"'{THIS_WORKBOOK}'!", f"{THIS_WORKBOOK}!"):
            if text.startswith(prefix):
                text = text[len(prefix):]
                break
        return None if THIS_WORKBOOK in text else text

    def stable_parameter(self, tokens: list[Tok], module: Module, proc: Proc | None) -> str | None:
        """The ByVal parameter of ``proc`` that ``tokens`` name, if nothing in the procedure can change it.

        It must never be assigned, and never passed by reference to a procedure that could assign it.
        """
        if proc is None or len(tokens) != 1 or not tokens[0].is_id():
            return None
        key = tokens[0].low
        param = next((p for p in proc.params if p.name.casefold() == key), None)
        if param is None or not param.byval or param.paramarray or proc.locals.get(key, (None, ""))[1] != "param":
            return None
        if key in assigned_names(proc) or key in self.passed_by_reference(module, proc):
            return None
        return key

    def call_sites(self, tokens: list[Tok]) -> list[tuple[int, list[Tok]]]:
        """The calls in one statement: the index of each callee name and the tokens of its arguments.

        A statement that starts with a name or a dotted chain, without ``Call`` and without an ``=``
        after it, is a call without parentheses: in ``Fix (x)`` the parentheses belong to the
        argument, which VBA then passes by value. Elsewhere a name followed by ``(`` takes the
        parenthesized list as its arguments.
        """
        sites: list[tuple[int, list[Tok]]] = []
        head = tokens[0]
        head_end = -1
        if head.is_id() and not head.is_kw(*KEYWORDS) and head.low not in STATEMENT_KEYWORDS:
            end = 0
            while end + 2 < len(tokens) and tokens[end + 1].continues_chain() and tokens[end + 2].is_id():
                end += 2
            rest = tokens[end + 1:]
            after = matching_paren(tokens, end + 1) + 1 if rest and rest[0].is_op("(") else end + 1
            if rest and not rest[0].is_op("=", ".", "!", ":=") and not (
                    after < len(tokens) and tokens[after].is_op("=")) and after > end:
                sites.append((end, rest))
                head_end = end
        for index, token in enumerate(tokens[:-1]):
            if index == head_end or not token.is_id() or token.is_kw(*KEYWORDS):
                continue
            if tokens[index + 1].is_op("("):
                close = matching_paren(tokens, index + 1)
                if close > index + 1:
                    sites.append((index, tokens[index + 2:close]))
        return sites

    def may_write_argument(self, tokens: list[Tok], index: int, args: list[tuple[str | None, list[Tok]]],
                           position: int, module: Module, proc: Proc) -> bool:
        """Whether the call at ``tokens[index]`` may change the variable passed as argument ``position``."""
        callee = tokens[index]
        qualified = index >= 2 and tokens[index - 1].continues_chain()
        if not qualified and callee.low in READ_ONLY_FUNCTIONS:
            return False
        owner: Module | None = module
        if qualified:
            owner = self.project.modules.get(tokens[index - 2].low) if tokens[index - 2].is_id() and (
                index < 3 or not tokens[index - 3].continues_chain()) else None
            member = owner.members.get(callee.low) if owner is not None else None
        else:
            kind, entity = self.resolve_name(callee.text, module, proc)
            if kind == "local":
                return False  # an array element, not a call
            member = entity[1] if kind == "member" else None
        target = next(iter(member.accessors.values()), None) if member is not None and member.kind == "proc" else None
        if target is None:
            return True
        bound = bind_arguments(target.params, args)
        argument = args[position][1]
        slot = next((i for i, value in bound.items() if value is argument), None)
        return slot is None or not target.params[slot].byval

    def passed_by_reference(self, module: Module, proc: Proc) -> set[str]:
        """Parameter names ``proc`` passes, as a bare argument, to a call that may change them."""
        cache = self.by_reference
        if id(proc) not in cache:
            params = {p.name.casefold() for p in proc.params}
            names: set[str] = set()
            for _line, tokens in proc.statements:
                for index, argument_tokens in self.call_sites(tokens):
                    args = self.split_args(argument_tokens)
                    for position, (_name, value) in enumerate(args):
                        if (len(value) == 1 and value[0].is_id() and value[0].low in params
                                and self.may_write_argument(tokens, index, args, position, module, proc)):
                            names.add(value[0].low)
            cache[id(proc)] = names
        return cache[id(proc)]

    def macro_value(self, tokens: list[Tok], module: Module, proc: Proc | None) -> tuple[str, str | None]:
        """Evaluate a macro-name expression: ("literal", name), ("parameter", name) or ("unknown", None).

        Literals, constants and macro-name functions are joined across ``&``. A leading
        ``"'" & ThisWorkbook.Name & "'!"`` names this project, so it is removed. An unassigned ByVal
        parameter of the enclosing procedure, alone or after that prefix, makes the parameter a macro
        parameter. Anything else is unknown.
        """
        parts = split_top(tokens, "&")
        text = ""
        for position, part in enumerate(parts):
            if self.is_this_workbook_name(part):
                text += THIS_WORKBOOK
                continue
            value = self.literal_target(part, module, proc)
            if value is not None:
                text += value
                continue
            parameter = self.stable_parameter(part, module, proc)
            if parameter is not None and position == len(parts) - 1 and self.strip_this_workbook(text) == "":
                return "parameter", parameter
            return "unknown", None
        local = self.strip_this_workbook(text)
        return ("literal", local) if local else ("unknown", None)

    def check_macro(self, channel: str, tokens: list[Tok], module: Module, proc: Proc | None, line: int,
                    via: str = "") -> None:
        """Check the macro name ``tokens`` evaluate to, as the target of ``channel``."""
        kind, value = self.macro_value(tokens, module, proc)
        if kind == "parameter" and proc is not None and value is not None:
            key = (module.name.casefold(), proc.name.casefold())
            self.macro_parameters.setdefault(key, {}).setdefault(value, channel)
            self.entry_points.append({"channel": channel + via, "module": module.name, "line": line,
                                      "target": f"{proc.name}({value})", "status": "parameter"})
            return
        if value is None:
            self.stats["dynamic-unknown"] += 1
            self.unknown[f"dynamic {channel} target"] += 1
            self.entry_points.append({"channel": channel + via, "module": module.name, "line": line,
                                      "target": None, "status": "unknown-dynamic"})
            return
        self.check_macro_target(channel, via, value, module, proc, line)

    def check_macro_arguments(self, target: Proc, args: list[tuple[str | None, list[Tok]]], module: Module,
                              proc: Proc | None, line: int) -> None:
        """Check what a call passes to the macro parameters of ``target``, or the defaults it leaves them."""
        carried = self.macro_parameters.get((target.module.casefold(), target.name.casefold()))
        if not carried:
            return
        bound = bind_arguments(target.params, args)
        for index, param in enumerate(target.params):
            channel = carried.get(param.name.casefold())
            if channel is None:
                continue
            via = f" (through {target.name} argument {param.name})"
            if index in bound:
                self.check_macro(channel, bound[index], module, proc, line, via)
                continue
            default = parameter_default(param)
            if default is not None:
                self.check_macro_target(channel, f" (the default of {target.name} argument {param.name})",
                                        default, module, proc, line)

    def check_macro_target(self, channel: str, via: str, value: str, module: Module, proc: Proc | None,
                           line: int) -> None:
        """Resolve a literal macro name against the standard modules of the project."""
        macro = value.split("!")[-1].strip("'")
        parts = macro.split(".")
        name = parts[-1]
        candidates = []
        for candidate in self.project.standard_modules():
            if len(parts) > 1 and candidate.name.casefold() != parts[-2].casefold():
                continue
            member = candidate.members.get(name.casefold())
            if member is not None and member.kind == "proc" and ({"sub", "function"} & set(member.accessors)):
                candidates.append((candidate, member))
        status = "resolved"
        # A target qualified by a standard module of this project, with no workbook prefix, can only
        # mean that module, so its member is checked whatever its name.
        qualifier = self.project.modules.get(parts[-2].casefold()) if len(parts) > 1 else None
        local_module = qualifier is not None and qualifier.kind == "standard"
        if not candidates:
            if (self.project.is_project_name(name) or local_module) and "!" not in value:
                status = "missing"
                self.report("VBA-CALL-030", module, proc, line, value,
                            f"{channel} target {value!r}{via} is not a procedure of any standard module "
                            f"of configuration {self.project.name}.", channel=channel)
            else:
                status = "unknown-external"
                self.unknown[f"{channel} target outside project"] += 1
        elif all(member.visibility == "private" for _, member in candidates):
            status = "private"
            self.report("VBA-CALL-031", module, proc, line, value,
                        f"{channel} target {value!r}{via} is Private in {candidates[0][0].name}.", channel=channel)
        self.stats["dynamic-literal"] += 1
        self.entry_points.append({"channel": channel + via, "module": module.name, "line": line, "target": value,
                                  "status": status})

    def check_callbyname(self, args: list[tuple[str | None, list[Tok]]], module: Module, proc: Proc | None,
                         line: int) -> None:
        if len(args) < 2:
            return
        receiver = self.chain_type(args[0][1], module, proc)
        member_name = self.literal_target(args[1][1], module, proc)
        entry = {"channel": "CallByName", "module": module.name, "line": line, "target": member_name}
        if receiver is None or member_name is None:
            self.stats["dynamic-unknown"] += 1
            self.unknown["dynamic CallByName target"] += 1
            self.entry_points.append({**entry, "status": "unknown-dynamic"})
            return
        member = receiver.members.get(member_name.casefold())
        if member is None:
            if receiver.kind == "class" and not receiver.implements:
                self.report("VBA-CALL-030", module, proc, line, f"{receiver.name}.{member_name}",
                            f"CallByName member {member_name!r} is not declared by {receiver.name}.",
                            channel="CallByName")
                status = "missing"
            else:
                status = "unknown-form-member"
                self.unknown["CallByName member not declared in form code"] += 1
        elif member.visibility == "private":
            self.report("VBA-CALL-003", module, proc, line, f"{receiver.name}.{member_name}",
                        f"CallByName targets Private member {member_name!r} of {receiver.name}.")
            status = "private"
        else:
            status = "resolved"
        self.stats["dynamic-literal"] += 1
        self.entry_points.append({**entry, "target": f"{receiver.name}.{member_name}", "status": status})

    # -- chains --------------------------------------------------------------
    def chain_type(self, tokens: list[Tok], module: Module, proc: Proc | None) -> Module | None:
        """Project module referred to by a simple receiver expression, or None."""
        if not tokens:
            return None
        if tokens[0].is_kw("new") and len(tokens) == 2:
            return self.type_module(tokens[1].text)
        if len(tokens) == 1 and tokens[0].is_id():
            kind, entity = self.resolve_name(tokens[0].text, module, proc)
            if kind == "local":
                return self.type_module(entity[0])
            if kind == "module":
                return None if self.lacks_default_instance(entity, tokens[0].text) else cast(Module, entity)
            if kind == "member" and entity[1].kind == "var":
                return self.type_module(entity[1].type)
        return None

    def process_chain(self, tokens: list[Tok], start: int, module: Module, proc: Proc | None, line: int,
                      context: str = "read", with_receiver: Module | None = None,
                      statement_args: list[Tok] | None = None) -> int:
        """Resolve one member chain beginning at ``start``; return the index after it."""
        index = start
        elements: list[tuple[str, list[Tok] | None]] = []
        leading_dot = tokens[index].is_op(".")
        if leading_dot:
            index += 1
        while index < len(tokens) and tokens[index].is_id():
            name = tokens[index].text.rstrip("$%&@")
            index += 1
            args: list[Tok] | None = None
            if index < len(tokens) and tokens[index].is_op("("):
                close = matching_paren(tokens, index)
                if close < 0:
                    close = len(tokens) - 1
                is_last = close + 1 >= len(tokens) or not tokens[close + 1].continues_chain()
                # In a statement call such as "Foo (a), b" the parentheses belong to the
                # first argument, not to the call: leave them for the trailing arguments.
                parenthesized_first_argument = statement_args is not None and is_last and close != len(tokens) - 1
                if not parenthesized_first_argument:
                    args = tokens[index + 1:close]
                    index = close + 1
            elements.append((name, args))
            if (index < len(tokens) and tokens[index].continues_chain() and index + 1 < len(tokens)
                    and tokens[index + 1].is_id()):
                index += 1
                continue
            break
        if not elements:
            return index + (0 if leading_dot else 1)
        for _name, args in elements:
            if args:
                self.scan_expression(args, module, proc, line, with_receiver)
        if statement_args is not None and index < len(tokens):
            trailing = tokens[index:]
            self.scan_expression(trailing, module, proc, line, with_receiver)
            name, args = elements[-1]
            if args is None:
                elements[-1] = (name, trailing)
            index = len(tokens)
        elif statement_args is not None and elements[-1][1] is None:
            elements[-1] = (elements[-1][0], [])
        self.resolve_elements(elements, leading_dot, module, proc, line, context, with_receiver,
                              statement=statement_args is not None)
        return index

    def resolve_elements(  # noqa: C901 - one branch per kind of name a member chain can start with
            self, elements: list[tuple[str, list[Tok] | None]], leading_dot: bool, module: Module,
            proc: Proc | None, line: int, context: str, with_receiver: Module | None, statement: bool) -> None:
        receiver: Module | None = None
        start = 0
        dynamic_receiver = ""
        if leading_dot:
            receiver = with_receiver
            if receiver is None:
                self.unknown["member of an unresolved With receiver"] += 1
                return
        else:
            name, args = elements[0]
            kind, entity = self.resolve_name(name, module, proc)
            last = len(elements) == 1
            element_context = context if last else "read"
            parsed = self.split_args(args) if args is not None else None
            if kind == "local":
                receiver = self.type_module(entity[0])
            elif kind == "self":
                if last:
                    self.check_recursive_call(entity, parsed, element_context, statement, module, proc, line, name)
                return
            elif kind == "member":
                owner, member = entity
                if member.kind in {"type", "enum"}:
                    return
                if name.casefold() in self.project.macro_functions and parsed:
                    self.check_macro("macro name", parsed[0][1], module, proc, line)
                receiver = self.check_member_call(owner, member, parsed, element_context, module, proc, line, name)
            elif kind == "module":
                if self.lacks_default_instance(entity, name):
                    self.report("VBA-CALL-005", module, proc, line, name,
                                f"{name} is a class without a default instance (VB_PredeclaredId = False); "
                                "call it through a variable that holds an instance.")
                    return
                receiver = entity
            elif kind == "ambiguous":
                self.report("VBA-CALL-004", module, proc, line, name,
                            f"{name} is public in {', '.join(m.name for m, _ in entity)}; "
                            "qualify the call with the module name.",
                            candidates=[{"module": m.name, "path": m.path, "line": mem.line} for m, mem in entity])
                return
            elif kind == "private":
                owners = ", ".join(m.name for m, _ in entity)
                self.report("VBA-CALL-003", module, proc, line, name,
                            f"{name} is Private in {owners} and not visible from {module.name}.",
                            candidates=[{"module": m.name, "path": m.path, "line": mem.line} for m, mem in entity])
                return
            else:
                lowered = name.casefold()
                if lowered == "callbyname" and parsed is not None:
                    self.check_callbyname(parsed, module, proc, line)
                    return
                if lowered == "addressof":
                    return
                if self.project.is_project_name(name) and (args is not None or statement or last):
                    self.report("VBA-CALL-001", module, proc, line, name,
                                f"{name} is not declared by any module of configuration {self.project.name}.")
                    return
                if args is not None or statement:
                    self.unknown["call to a VBA, host or library name"] += 1
                dynamic_receiver = lowered
            start = 1
        for position in range(start, len(elements)):
            name, args = elements[position]
            lowered = name.casefold()
            parsed = self.split_args(args) if args is not None else None
            last = position == len(elements) - 1
            if receiver is None:
                if dynamic_receiver == "excel" and lowered == "application":
                    dynamic_receiver = "application"
                    continue
                if dynamic_receiver == "application" and parsed is not None:
                    if lowered == "run" and parsed:
                        self.check_macro("Application.Run", parsed[0][1], module, proc, line)
                    elif lowered in {"ontime", "onkey"}:
                        named = [a for n, a in parsed if n and n.casefold() == "procedure"]
                        target = named[0] if named else (parsed[1][1] if len(parsed) > 1 else None)
                        if target is not None:
                            self.check_macro("Application." + name, target, module, proc, line)
                elif not leading_dot or position > start:
                    self.unknown["member of a receiver whose type is not a project class"] += 1
                break
            member = receiver.members.get(lowered)
            label = f"{receiver.name}.{name}"
            if member is None:
                if receiver.kind == "standard" or (receiver.kind == "class" and not receiver.implements):
                    self.report("VBA-CALL-002", module, proc, line, label,
                                f"{receiver.name} does not declare {name}.")
                else:
                    self.unknown["member not declared in form code (control or UserForm member)"] += 1
                receiver = None
                continue
            if member.visibility == "private" and receiver is not module:
                self.report("VBA-CALL-003", module, proc, line, label, f"{name} is Private in {receiver.name}.")
                receiver = None
                continue
            element_context = context if last else "read"
            receiver = self.check_member_call(receiver, member, parsed, element_context, module, proc, line, label)

    # -- expressions and statements -----------------------------------------
    def scan_expression(self, tokens: list[Tok], module: Module, proc: Proc | None, line: int,
                        with_receiver: Module | None) -> None:
        index = 0
        while index < len(tokens):
            token = tokens[index]
            previous = tokens[index - 1] if index else None
            if token.is_kw("new"):
                index += 2
                continue
            if token.is_kw("addressof") and index + 1 < len(tokens):
                self.check_addressof(tokens[index + 1], module, proc, line)
                index += 2
                continue
            if token.is_kw("is") and previous is not None:
                index += 2 if index + 1 < len(tokens) and tokens[index + 1].is_id() else 1
                continue
            if token.is_id() and index + 1 < len(tokens) and tokens[index + 1].is_op(":="):
                index += 2
                continue
            chain_start = (token.is_id() and not token.is_kw(*KEYWORDS)) or (
                token.is_op(".") and (previous is None
                                      or previous.is_op("(", ",", "=", "<>", "<", ">", "<=", ">=",
                                                        "+", "-", "*", "/", "\\", "&", "^", ";")
                                      or previous.is_kw(*KEYWORDS)))
            if chain_start and (previous is None or not previous.is_op(".", "!")):
                index = self.process_chain(tokens, index, module, proc, line, "read", with_receiver)
                continue
            index += 1

    def check_addressof(self, token: Tok, module: Module, proc: Proc | None, line: int) -> None:
        kind, entity = self.resolve_name(token.text, module, proc)
        target = token.text
        status = "unknown"
        if kind == "self" or (kind == "member" and entity[1].kind == "proc"):
            status = "resolved"
        elif kind == "member":
            status = "not-a-procedure"
            self.report("VBA-CALL-030", module, proc, line, target,
                        f"AddressOf target {target!r} is a {entity[1].kind} of {entity[0].name}, not a procedure.",
                        channel="AddressOf")
        elif kind == "private":
            status = "private"
            self.report("VBA-CALL-003", module, proc, line, target,
                        f"AddressOf target {target!r} is Private in {', '.join(m.name for m, _ in entity)}.",
                        channel="AddressOf",
                        candidates=[{"module": m.name, "path": m.path, "line": mem.line} for m, mem in entity])
        elif kind == "ambiguous":
            status = "ambiguous"
            self.report("VBA-CALL-004", module, proc, line, target,
                        f"AddressOf target {target!r} is public in {', '.join(m.name for m, _ in entity)}.",
                        channel="AddressOf",
                        candidates=[{"module": m.name, "path": m.path, "line": mem.line} for m, mem in entity])
        elif kind == "external" and self.project.is_project_name(target):
            status = "missing"
            self.report("VBA-CALL-030", module, proc, line, target, f"AddressOf target {target!r} is not declared.",
                        channel="AddressOf")
        self.entry_points.append({"channel": "AddressOf", "module": module.name, "line": line, "target": target,
                                  "status": status})

    def statement(  # noqa: C901 - one branch per VBA statement form, kept side by side
            self, tokens: list[Tok], module: Module, proc: Proc, line: int, with_stack: list[Module | None]) -> None:
        receiver = with_stack[-1] if with_stack else None
        head = tokens[0]
        low = head.low if head.is_id() else ""
        if low in {"dim", "static", "const", "redim", "erase", "attribute", "exit", "goto", "gosub", "return",
                   "resume", "stop", "option", "next", "wend", "implements", "rem"} or (
                       head.is_id() and low == "end" and len(tokens) == 1):
            if low == "redim":
                for part in split_top(tokens[1:], ","):
                    if part and part[0].is_kw("preserve"):
                        part = part[1:]
                    if len(part) > 2 and part[1].is_op("("):
                        close = matching_paren(part, 1)
                        self.scan_expression(part[2:close], module, proc, line, receiver)
            return
        if low == "end":
            if len(tokens) > 1 and tokens[1].is_kw("with") and with_stack:
                with_stack.pop()
            return
        if low == "on":
            return
        if low in {"if", "elseif"}:
            then = next((i for i, t in enumerate(tokens) if t.is_kw("then")), len(tokens))
            self.scan_expression(tokens[1:then], module, proc, line, receiver)
            tail = tokens[then + 1:]
            if tail:
                depth, split_at = 0, None
                for i, t in enumerate(tail):
                    depth += t.is_op("(") - t.is_op(")")
                    if depth == 0 and t.is_kw("else"):
                        split_at = i
                        break
                parts = [tail] if split_at is None else [tail[:split_at], tail[split_at + 1:]]
                for part in parts:
                    if part:
                        self.statement(part, module, proc, line, with_stack)
            return
        if low == "else":
            if len(tokens) > 1:
                self.statement(tokens[1:], module, proc, line, with_stack)
            return
        if low in {"select", "case", "while", "do", "loop", "debug"}:
            rest = tokens[1:]
            if low == "debug":
                rest = tokens[3:] if len(tokens) > 2 else []
            filtered = [t for t in rest if not t.is_kw("case", "else", "to", "while", "until")]
            self.scan_expression(filtered, module, proc, line, receiver)
            return
        if low == "for":
            body = tokens[1:]
            if body and body[0].is_kw("each"):
                inside = next((i for i, t in enumerate(body) if t.is_kw("in")), None)
                if inside is not None:
                    self.scan_expression(body[inside + 1:], module, proc, line, receiver)
            else:
                equals = next((i for i, t in enumerate(body) if t.is_op("=")), None)
                if equals is not None:
                    self.scan_expression([t for t in body[equals + 1:] if not t.is_kw("to", "step")],
                                         module, proc, line, receiver)
            return
        if low == "with":
            self.scan_expression(tokens[1:], module, proc, line, receiver)
            expression = tokens[1:]
            if expression and expression[0].is_op(".") and receiver is not None and len(expression) == 2:
                member = receiver.members.get(expression[1].low)
                with_stack.append(self.type_module(member.type) if member else None)
            else:
                with_stack.append(self.chain_type(expression, module, proc))
            return
        if low == "raiseevent":
            if len(tokens) > 2 and tokens[2].is_op("("):
                self.scan_expression(tokens[3:matching_paren(tokens, 2)], module, proc, line, receiver)
            return
        if low in STATEMENT_KEYWORDS and not (len(tokens) > 1 and tokens[1].is_op("=", ".")):
            self.scan_expression(tokens[1:], module, proc, line, receiver)
            return
        if low == "call":
            if len(tokens) > 1:
                self.stats["call-statements"] += 1
                self.process_chain(tokens, 1, module, proc, line, "call", receiver)
            return
        context = "assign"
        body = tokens
        if low in {"set", "let"}:
            context = "assign-set" if low == "set" else "assign"
            body = tokens[1:]
        equals = None
        depth = 0
        for i, t in enumerate(body):
            depth += t.is_op("(") - t.is_op(")")
            if depth == 0 and t.is_op("="):
                equals = i
                break
        if equals is not None and equals > 0 and is_lvalue(body[:equals]):
            lhs, rhs = body[:equals], body[equals + 1:]
            self.scan_expression(rhs, module, proc, line, receiver)
            if lhs[-1].is_id() and lhs[-1].low == "onaction":
                self.check_macro("OnAction", rhs, module, proc, line)
            self.process_chain(lhs, 0, module, proc, line, context, receiver)
            return
        if body and (body[0].is_id() or body[0].is_op(".")):
            self.stats["call-statements"] += 1
            self.process_chain(body, 0, module, proc, line, "read", receiver, statement_args=[])
            return
        self.scan_expression(body, module, proc, line, receiver)

    def run(self) -> None:
        for module in self.project.modules.values():
            for proc in module.procs:
                with_stack: list[Module | None] = []
                for line, tokens in proc.statements:
                    self.statement(tokens, module, proc, line, with_stack)
            self.classify_entry_points(module)

    def classify_entry_points(self, module: Module) -> None:
        for proc in module.procs:
            name = proc.name.casefold()
            prefix = name.split("_", 1)[0] if "_" in name else ""
            kind = None
            if prefix in module.withevents:
                kind = "event-handler (WithEvents)"
            elif module.kind == "form" and prefix == "userform":
                kind = "event-handler (UserForm)"
            elif module.kind in {"class", "form"} and name in {"class_initialize", "class_terminate"}:
                kind = "event-handler (class)"
            elif module.kind == "form" and prefix and proc.visibility == "private" and prefix not in {"uf"}:
                kind = "possible control event-handler"
            if kind:
                self.entry_points.append({"channel": kind, "module": module.name, "line": proc.line,
                                          "target": proc.name, "status": "external-entry"})


# ---------------------------------------------------------------------------
# Projects, coverage and the report
# ---------------------------------------------------------------------------


def duplicate_public_names(project: Project) -> list[dict[str, Any]]:
    owners: dict[str, list[tuple[Module, Member]]] = {}
    for module in project.standard_modules():
        for key, member in module.members.items():
            if member.visibility != "private":
                owners.setdefault(key, []).append((module, member))
    items = []
    for _key, entries in sorted(owners.items()):
        if len(entries) > 1:
            module, member = entries[0]
            items.append(finding("VBA-CALL-040", module, None, member.line, member.name,
                                 f"{member.name} is public in {', '.join(m.name for m, _ in entries)}; "
                                 "unqualified calls from other modules are ambiguous.",
                                 candidates=[{"module": m.name, "path": m.path, "line": mem.line}
                                             for m, mem in entries]))
    return items


def macro_parameters(project: Project) -> MacroParameters:
    """Find every parameter whose value reaches a macro sink, directly or through other procedures.

    A pass discovers the parameters handed to a sink, or passed on to a parameter already known to
    reach one; passes repeat until nothing new is found. Their findings are discarded: only the final
    pass, which knows every macro parameter, reports.
    """
    carried: MacroParameters = {}
    while True:
        before = sum(len(parameters) for parameters in carried.values())
        Analyzer(project, carried).run()
        if sum(len(parameters) for parameters in carried.values()) == before:
            return carried


def analyze_configuration(name: str, sources: dict[str, str], patterns: list[str], macro_functions: list[str],
                          environment: str, coverage_complete: bool) -> dict[str, Any]:
    modules: dict[str, Module] = {}
    failures: list[dict[str, Any]] = []
    declaration_findings: list[dict[str, Any]] = []
    for path, text in sorted(sources.items()):
        try:
            units = list(active_lines(text, CONFIGS[environment]))
        except LexicalError as error:
            failures.append(finding("VBA-CALL-090", None, None, None, None,
                                    f"{path}: source or conditional compilation cannot be read: {error}",
                                    path=path))
            continue
        module = parse_module(path, text, units)
        failures.extend(module.failures)
        declaration_findings.extend(module.findings)
        key = module.name.casefold()
        if key in modules:
            failures.append(finding("VBA-CALL-091", module, None, None, module.name,
                                    f"Module name {module.name} is used by {modules[key].path} and {path}."))
        modules[key] = module
    # VBA identifiers are case-insensitive, so project-name patterns are too.
    project = Project(name, modules, [re.compile(p, re.IGNORECASE) for p in patterns],
                      {m.casefold() for m in macro_functions},
                      coverage_complete and not failures)
    analyzer = Analyzer(project)
    if not failures:
        analyzer = Analyzer(project, macro_parameters(project))
        analyzer.run()
    inventory = {
        module.name: {
            "path": module.path,
            "kind": module.kind,
            "option_private_module": module.option_private,
            "predeclared": module.predeclared,
            "implements": module.implements,
            "procedures": [
                {"name": proc.name, "accessor": proc.accessor, "visibility": proc.visibility,
                 "line": proc.line, "signature": proc.signature()}
                for proc in module.procs
            ],
        }
        for module in modules.values()
    }
    return {
        "environment": environment,
        "complete": project.complete,
        "failures": failures,
        "findings": declaration_findings + analyzer.findings + duplicate_public_names(project),
        "stats": dict(analyzer.stats),
        "unknown": dict(analyzer.unknown),
        "entry_points": analyzer.entry_points,
        "inventory": inventory,
        "project": project,
    }


def merge_environments(runs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: dict[tuple[Any, ...], dict[str, Any]] = {}
    for run in runs:
        for item in run["findings"] + run["failures"]:
            key = (item["code"], item.get("path"), item.get("line"), item.get("target"), item["message"])
            entry = merged.setdefault(key, {**item, "environments": []})
            entry["environments"].append(run["environment"])
    all_envs = [run["environment"] for run in runs]
    for item in merged.values():
        item["configuration_dependent"] = item["environments"] != all_envs
        if item["configuration_dependent"] and item["severity"] == "error":
            item["severity"] = "unknown"
            item["resolution"] = "configuration-dependent"
    return sorted(merged.values(), key=lambda f: (f["severity"], f.get("path") or "", f.get("line") or 0, f["code"]))


def analyze_project(files: dict[str, str], manifest: dict[str, Any],
                    tracked: set[str] | None = None) -> dict[str, Any]:
    """Analyze every configuration of ``manifest`` over ``files`` (repository path -> source text)."""
    patterns = manifest.get("project_name_patterns", [])
    macro_functions = manifest.get("macro_name_functions", [])
    configurations = manifest.get("configurations", {})
    coverage: list[dict[str, Any]] = []
    vba_files = {p for p in files if Path(p).suffix.casefold() in VBA_SUFFIXES}
    listed = {m for c in configurations.values() for m in c.get("modules", [])}
    documents = set(manifest.get("document_modules", []))
    for path in sorted(vba_files - listed - documents):
        coverage.append(finding("VBA-CALL-091", None, None, None, path,
                                f"{path} is not assigned to any configuration in {MANIFEST}.", path=path))
    if tracked is not None:
        for path in sorted((documents | listed) - tracked):
            coverage.append(finding("VBA-CALL-091", None, None, None, path,
                                    f"{path} is listed in {MANIFEST} but not tracked.", path=path))
    results: dict[str, Any] = {}
    for name, config in configurations.items():
        sources = {p: files[p] for p in config.get("modules", []) if p in files}
        missing = [p for p in config.get("modules", []) if p not in files]
        complete = not missing
        runs = [analyze_configuration(name, sources, patterns, macro_functions, env, complete) for env in ENVIRONMENTS]
        findings = merge_environments(runs)
        for path in missing:
            findings.append(finding("VBA-CALL-091", None, None, None, path,
                                    f"{path} is missing from configuration {name}.", path=path))
        if not config.get("gating", True):
            for item in findings:
                if item["severity"] == "error":
                    item["severity"] = "warning"
                    item["resolution"] = "non-gating configuration"
        failures = [f for f in findings if f["severity"] == "failure"]
        last = runs[-1]
        results[name] = {
            "description": config.get("description", ""),
            "gating": config.get("gating", True),
            "modules_expected": config.get("modules", []),
            "modules_analyzed": sorted(sources),
            "environments": list(ENVIRONMENTS),
            "complete": complete and not failures,
            "stats": last["stats"],
            "unknown": last["unknown"],
            "entry_points": last["entry_points"],
            "inventory": last["inventory"],
            "findings": findings,
        }
    errors = [f for r in results.values() for f in r["findings"] if f["severity"] == "error"] + \
        [f for f in coverage if f["severity"] == "error"]
    incomplete = bool(coverage) or any(not r["complete"] for r in results.values() if r["gating"])
    status = "fail" if errors else "incomplete" if incomplete else "pass"
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": status,
        "advisory": True,
        "manifest": MANIFEST,
        "environments": {env: dict(CONFIGS[env]) for env in ENVIRONMENTS},
        "coverage_findings": coverage,
        "configurations": results,
        "unsupported": UNSUPPORTED,
    }


UNSUPPORTED = [
    "type compatibility of arguments",
    "receivers typed Object, Variant or a non-project library type",
    "member chains after a member whose type is not a project class",
    "controls and built-in members of UserForms (reported as unknown)",
    "macro names computed at run time, read from arrays or cells, or held in a variable assigned more than once",
    "procedures never called (no unused-code verdict; external entry points are classified instead)",
    "callback and event-handler signatures",
    "document modules (ThisWorkbook, worksheets), not analyzed: their Excel members are unknown to the checker",
    "conditional-compilation symbols other than VBA7, Win64 and module #Const (analysis fails closed)",
]


def tracked_files(root: Path) -> set[str]:
    """Repository paths Git tracks under ``root``."""
    completed = subprocess.run(["git", "-C", str(root), "ls-files", "-z"], check=False, capture_output=True)
    if completed.returncode:
        raise RuntimeError("git ls-files failed: " + completed.stderr.decode("utf-8", errors="replace").strip())
    return {item.decode("utf-8", errors="surrogateescape") for item in completed.stdout.split(b"\0") if item}


def run_check(root: Path) -> dict[str, Any]:
    """Analyze the tracked VBA modules of the repository at ``root`` against its manifest."""
    manifest = json.loads((root / MANIFEST).read_text(encoding="utf-8"))
    if not isinstance(manifest, dict) or not isinstance(manifest.get("configurations"), dict):
        raise ValueError(f"{MANIFEST} has no configurations object")
    tracked = tracked_files(root)
    files = {p: (root / p).read_bytes().decode("cp1252") for p in sorted(tracked)
             if Path(p).suffix.casefold() in VBA_SUFFIXES}
    return analyze_project(files, manifest, tracked)


def where(item: dict[str, Any]) -> str:
    location = f"{item.get('path') or ''}:{item.get('line') or ''}".rstrip(":")
    return f"{location} {item.get('procedure') or ''}".strip()


def configuration_counts(result: dict[str, Any]) -> tuple[Counter[str], dict[str, int]]:
    """Finding counts by severity, and the resolved and unknown counts of the last environment."""
    severities = Counter(f["severity"] for f in result["findings"])
    stats = result["stats"]
    counts = {
        "resolved procedure uses": stats.get("resolved-calls", 0),
        "argument checks": stats.get("argument-checks", 0),
        "literal macro targets": stats.get("dynamic-literal", 0),
        "unknown": sum(result["unknown"].values()),
    }
    return severities, counts


def text_report(report: dict[str, Any]) -> str:
    """Plain-text summary for a terminal."""
    lines = [f"{TOOL_NAME}: {report['status'].upper()}"]
    for item in report["coverage_findings"]:
        lines.append(f"  coverage: {item['message']}")
    for name, result in report["configurations"].items():
        severities, counts = configuration_counts(result)
        lines.append(f"{name}: {len(result['modules_analyzed'])} module(s), "
                     + ", ".join(f"{label} {value}" for label, value in counts.items())
                     + f"; {severities.get('error', 0)} defect(s), {severities.get('warning', 0)} warning(s), "
                     f"{severities.get('unknown', 0)} unknown, {severities.get('failure', 0)} analysis failure(s)")
        for item in result["findings"]:
            if item["severity"] in {"error", "warning", "failure"}:
                lines.append(f"  {item['severity']} {item['code']} {where(item)}: {item['message']}")
    return "\n".join(lines) + "\n"


def markdown_report(report: dict[str, Any]) -> str:
    """Markdown summary, for a job summary."""
    lines = [f"## {TOOL_NAME}", "",
             f"- **Status:** {report['status'].upper()} (advisory; not a release gate)",
             f"- **Manifest:** `{report['manifest']}`",
             f"- **Environments:** {', '.join(ENVIRONMENTS)}"]
    for item in report["coverage_findings"]:
        lines.append(f"- **Coverage:** {item['message']}")
    for name, result in report["configurations"].items():
        severities, counts = configuration_counts(result)
        lines += ["", f"### {name}{'' if result['gating'] else ' (non-gating)'}", "",
                  f"- Modules: {len(result['modules_analyzed'])} analyzed of "
                  f"{len(result['modules_expected'])} expected; complete: {result['complete']}",
                  "- " + "; ".join(f"{label.capitalize()}: {value}" for label, value in counts.items()),
                  f"- Findings: {severities.get('error', 0)} error, {severities.get('warning', 0)} warning, "
                  f"{severities.get('unknown', 0)} unknown, {severities.get('failure', 0)} analysis failure"]
        for reason, count in sorted(result["unknown"].items(), key=lambda kv: (-kv[1], kv[0])):
            lines.append(f"- Unknown: {reason}: {count}")
        if result["findings"]:
            lines += ["", "| Severity | Code | Location | Target | Finding |", "| --- | --- | --- | --- | --- |"]
            for item in result["findings"]:
                envs = "" if not item.get("configuration_dependent") else f" [only {', '.join(item['environments'])}]"
                lines.append("| {s} | {c} | {w} | {t} | {m}{e} |".format(
                    s=item["severity"], c=item["code"], w=where(item).replace("|", "\\|"),
                    t=str(item.get("target") or "").replace("|", "\\|"),
                    m=item["message"].replace("|", "\\|"), e=envs))
    lines += ["", "Unsupported (reported as unknown, never as defects): " + "; ".join(report["unsupported"]) + "."]
    return "\n".join(lines) + "\n"


def append_text(path: Path, text: str) -> None:
    """Append ``text`` to ``path``, as a job summary file expects; create the file if needed."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", newline="\n") as handle:
        handle.write(text)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Advisory VBA call checker: resolves calls across the modules of each configuration "
                    f"in {MANIFEST} and reports only what it can resolve with certainty.",
        epilog="Exit status: 0 no defect, 1 defect or incomplete analysis, 2 usage or input error.")
    parser.add_argument("--root", type=Path, default=REPO, help="repository root (default: this checkout)")
    parser.add_argument("--output", type=Path, help="write the full report as JSON to this file")
    parser.add_argument("--summary", type=Path,
                        help="append a Markdown summary to this file (for example $GITHUB_STEP_SUMMARY)")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        report = run_check(options.root)
        if options.output is not None:
            options.output.parent.mkdir(parents=True, exist_ok=True)
            options.output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8",
                                      newline="\n")
        if options.summary is not None:
            append_text(options.summary, markdown_report(report))
    except (OSError, UnicodeError, RuntimeError, ValueError, KeyError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(text_report(report), end="")
    return 0 if report["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
