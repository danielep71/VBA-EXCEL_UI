"""#203 VBA syntax rules: mistakes the compiler rejects and a parser can see.

Excel compiles a project only on the Windows host, so a statement that cannot
compile used to reach me there. This module reads the statements of every
module, per compilation, and reports six classes of compile error:

  * parentheses: a logical statement, after its continuation lines are joined
    and strings and comments are masked, opens or closes more parentheses than
    it matches (the ``FindingSummary)`` that nothing opened); a string or
    bracketed identifier that never closes is reported as ``lexical``;
  * blocks: ``If``/``End If``, ``Select Case``/``End Select``, ``With``/
    ``End With``, ``For``/``Next``, ``Do``/``Loop`` or ``While``/``Wend`` that
    do not pair up inside one procedure;
  * continuations: a logical statement with more than 24 line continuations;
  * procedure size: a procedure whose active code reaches 64 KB. Only the code
    of the compilation being read counts: comment-only lines, trailing comments,
    blank lines and the text of inactive conditional branches are dropped, and
    each logical statement counts its text plus a two-byte line ending. VBA
    limits the compiled code, which this module cannot see, so the measure is a
    conservative proxy that stays below the source size and is not the compiled
    size itself;
  * LongPtr assignment: under VBA7 x64, an expression of integer arithmetic
    that reads a ``LongPtr`` variable, assigned to a variable declared ``Long``;
  * ByRef type: a bare variable passed ``ByRef`` to a parameter of another
    declared type, such as a ``Long`` to a ``LongPtr``.

Each compilation of vba_analyze.CONFIGS is read on its own, with the same
conditional-compilation reader as the VBA analyzer, so a branch that opens a
block in one arm and not in the other is judged per arm. Findings that are
identical in several compilations are merged and name those compilations.

Every rule is conservative. Where the text does not settle the question (an
unknown type, a call whose return type is unknown, an argument that is not a
bare variable, an inline ``If`` that ends the line, a name that two modules
declare) the rule stays silent: a missed defect, never a false alarm. The type
rules read declarations only, so they ignore ``Variant`` and undeclared names.

This is bounded source evidence, not compilation: it does not replace
Debug -> Compile VBAProject on a real Excel host.
"""
import json
import re
from collections.abc import Iterable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from vba_analyze import BLOCK, CONFIGS, DECLARE, PROC, active_lines, comma_parts, statements
from vba_lex import LexicalError, code_only, split_code_comment

CONTINUATION_LIMIT = 24
PROCEDURE_LIMIT = 64 * 1024
LONGPTR_CONFIG = 'vba7-x64'  # the one compilation in which LongPtr is wider than Long
INTEGER_TYPES = frozenset({'boolean', 'byte', 'integer', 'long', 'longlong', 'longptr'})
SCALAR_TYPES = INTEGER_TYPES | {'currency', 'date', 'double', 'single', 'string'}
STATEMENT_WIDTH = 140

CONTINUED = re.compile(r'\s_\s*$')
END_PROCEDURE = re.compile(r'End\s+(?:Sub|Function|Property)', re.I)
END_BLOCK = re.compile(r'End\s+(?:Type|Enum)\b', re.I)
DECL_HEAD = re.compile(r'(?:(?:Public|Private|Global|Dim|Static|Const)\s+)+(?=\S)', re.I)
NOT_A_VARIABLE = re.compile(r'(?:Type|Enum|Event|Declare|Sub|Function|Property|Implements)\b', re.I)
DECLARATOR = re.compile(
    r'(?:WithEvents\s+)?(?P<name>[A-Za-z_]\w*)(?P<suffix>[%&!#@$^])?\s*(?P<array>\(.*?\))?\s*'
    r'(?:As\s+(?:New\s+)?(?P<type>[A-Za-z_][\w.]*)(?P<fixed>\s*\*\s*\w+)?)?\s*(?:=.*)?', re.I | re.S)
PARAMETER = re.compile(
    r'(?:Optional\s+)?(?:(?P<mode>ByVal|ByRef)\s+)?(?P<paramarray>ParamArray\s+)?(?P<name>[A-Za-z_]\w*)'
    r'(?P<suffix>[%&!#@$^])?\s*(?P<array>\(\s*\))?\s*(?:As\s+(?P<type>[A-Za-z_][\w.]*))?', re.I)
RESULT_TYPE = re.compile(r'\s*As\s+([A-Za-z_][\w.]*)', re.I)
THEN = re.compile(r'\bThen\b', re.I)
ELSE = re.compile(r'\bElse\b', re.I)
ASSIGNMENT = re.compile(r'(?:Let\s+)?(\w+)\s*=(?!=)\s*(.+)', re.I | re.S)
KEYWORD_LABEL = frozenset({'else', 'do', 'loop', 'wend', 'next', 'with'})
ARITHMETIC_WORDS = frozenset({'and', 'or', 'xor', 'not', 'mod'})
EXPRESSION_TOKEN = re.compile(
    r'(?P<num>&H[0-9A-Fa-f]+[&^]?|&O[0-7]+[&^]?|\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?[&%!#@^]?)'
    r'|(?P<id>[A-Za-z_]\w*[%&!#@$^]?)|(?P<op>[-+*\\])|(?P<open>\()|(?P<close>\))|(?P<space>\s+)')
CALL_OPEN = re.compile(r'(?<![\w.])([A-Za-z_]\w*)\(')
STATEMENT_CALL = re.compile(r'([A-Za-z_]\w*)\s+(?![=(])(\S.*)', re.S)
IDENTIFIER = re.compile(r'[A-Za-z_]\w*')

# How the findings name a block: what closes it and what opens it.
CLOSERS = {'if': 'End If', 'select': 'End Select', 'with': 'End With', 'for': 'Next', 'do': 'Loop',
           'while': 'Wend'}
OPENERS = {'if': 'If', 'select': 'Select Case', 'with': 'With', 'for': 'For', 'do': 'Do', 'while': 'While'}
SIMPLE_BLOCKS = {'case': ('mid', 'select'), 'with': ('open', 'with'), 'for': ('open', 'for'),
                 'next': ('close', 'for'), 'do': ('open', 'do'), 'loop': ('close', 'do'),
                 'while': ('open', 'while'), 'wend': ('close', 'while')}

# (file, line, procedure, rule, message, statement) before the compilations are merged.
Raw = tuple[str, int, str, str, str, str]
Event = tuple[str, str]  # (open | close | mid | inline, block kind)
Config = Mapping[str, bool]


@dataclass(frozen=True, order=True)
class Finding:
    file: str
    line: int
    procedure: str
    rule: str
    message: str
    statement: str
    configs: tuple[str, ...]

    def __str__(self) -> str:
        where = 'all compilations' if len(self.configs) == len(CONFIGS) else ', '.join(self.configs)
        return f'{self.file}:{self.line}: {where}: {self.procedure}: {self.rule}: {self.message}: {self.statement}'


@dataclass
class Coverage:
    """What one compilation read: distinct modules, procedures, statements and compared items."""
    modules: set[str] = field(default_factory=set)
    procedures: set[tuple[str, int]] = field(default_factory=set)
    statements: set[tuple[str, int]] = field(default_factory=set)
    assignments: set[tuple[str, int]] = field(default_factory=set)
    arguments: set[tuple[str, int, int]] = field(default_factory=set)
    sizes: dict[tuple[str, int], int] = field(default_factory=dict)


@dataclass(frozen=True)
class Variable:
    type: str | None  # lower case; 'variant' when undeclared; None when the text does not settle it
    text: str
    array: bool = False
    const: bool = False


@dataclass(frozen=True)
class Parameter:
    name: str
    byref: bool
    type: str | None
    text: str
    array: bool
    paramarray: bool


@dataclass(frozen=True)
class Routine:
    name: str
    public: bool
    params: tuple[Parameter, ...] | None  # None when a parameter could not be read


@dataclass(frozen=True)
class Stmt:
    line: int
    text: str
    code: str


@dataclass
class Procedure:
    name: str
    kind: str
    line: int
    header: str
    result: str | None
    variables: dict[str, Variable] = field(default_factory=dict)
    statements: list[Stmt] = field(default_factory=list)
    size: int = 0  # bytes of active code: statement text plus two bytes of line ending


@dataclass
class Module:
    file: str
    variables: dict[str, Variable] = field(default_factory=dict)
    public: set[str] = field(default_factory=set)
    routines: dict[str, list[Routine]] = field(default_factory=dict)
    procedures: list[Procedure] = field(default_factory=list)


@dataclass
class Project:
    modules: list[Module]
    symbols: dict[str, Variable] = field(default_factory=dict)


def short(statement: str) -> str:
    """One line of a statement; a long one keeps its head and its tail."""
    text = ' '.join(statement.split())
    if len(text) <= STATEMENT_WIDTH:
        return text
    return f'{text[:70]} ... {text[-65:]}'


def continuation_counts(text: str) -> tuple[dict[int, int], list[tuple[int, str]]]:
    """Line continuations per logical statement, keyed by its first physical line.

    The segmentation is the one of vba_analyze.logical_lines, so the keys are
    the line numbers that active_lines yields. A physical line whose string or
    bracketed identifier never closes is returned in the second list.
    """
    counts: dict[int, int] = {}
    errors: list[tuple[int, str]] = []
    first, count, pending = 0, 0, False
    for n, raw in enumerate(text.splitlines(), 1):
        try:
            code = split_code_comment(raw)[0]
        except LexicalError:
            errors.append((n, raw.strip()))
            code = ''
        if not pending:
            first, count = n, 0
        pending = bool(CONTINUED.search(code))
        if pending:
            count += 1
        else:
            counts[first] = count
    return counts, errors


def paren_message(masked: str) -> str | None:
    """Describe the first parenthesis problem of a masked statement, or None."""
    depth = 0
    for ch in masked:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
            if depth < 0:
                return 'a closing parenthesis has no opening parenthesis'
    if depth:
        return f'{depth} opening {"parenthesis" if depth == 1 else "parentheses"} never closed'
    return None


def matching_paren(code: str, start: int) -> int | None:
    """Index of the parenthesis that closes the one at ``start``."""
    depth = 0
    for i in range(start, len(code)):
        if code[i] == '(':
            depth += 1
        elif code[i] == ')':
            depth -= 1
            if depth == 0:
                return i
    return None


def parse_parameters(inner: str) -> tuple[Parameter, ...] | None:
    """The parameters of a header, or None when one of them is not understood."""
    params: list[Parameter] = []
    for part in comma_parts(inner) if inner.strip() else ():
        m = PARAMETER.match(part.strip())
        if not m:
            return None
        text = m['type']
        kind = None if m['suffix'] else (text.lower() if text else 'variant')
        params.append(Parameter(m['name'], (m['mode'] or 'ByRef').lower() == 'byref', kind,
                                text or 'Variant', bool(m['array']), bool(m['paramarray'])))
    return tuple(params)


def parse_header(code: str, start: int) -> tuple[tuple[Parameter, ...] | None, str | None]:
    """Parameters and result type of a Sub, Function, Property or Declare statement."""
    open_at = code.find('(', start)
    if open_at < 0:
        return (), None
    close_at = matching_paren(code, open_at)
    if close_at is None:
        return None, None
    result = RESULT_TYPE.match(code[close_at + 1:])
    return parse_parameters(code[open_at + 1:close_at]), result[1].lower() if result else None


def declarations(code: str) -> list[tuple[str, Variable, bool]]:
    """(name, variable, public) of each variable a Dim, Static, Private, Public, Global or Const declares."""
    head = DECL_HEAD.match(code)
    if not head or NOT_A_VARIABLE.match(code[head.end():]):
        return []
    keywords = head.group().lower().split()
    const = 'const' in keywords
    found: list[tuple[str, Variable, bool]] = []
    for part in comma_parts(code[head.end():]):
        m = DECLARATOR.fullmatch(part.strip())
        if not m:
            continue
        text = m['type']
        kind = None if m['suffix'] or m['fixed'] or (const and text is None) else (text.lower() if text else 'variant')
        found.append((m['name'].lower(), Variable(kind, text or 'Variant', bool(m['array']), const),
                      'public' in keywords or 'global' in keywords))
    return found


class Reader:
    """One module read in one compilation: its model and the findings that need no other module."""

    def __init__(self, file: str, counts: Mapping[int, int], cov: Coverage) -> None:
        self.module = Module(file)
        self.raws: list[Raw] = []
        self.counts, self.cov = counts, cov
        self.current: Procedure | None = None
        self.block: tuple[str, bool] | None = None
        self.closed: Procedure | None = None

    def read(self, text: str, config: Config) -> None:
        for n, raw in active_lines(text, dict(config)):
            self.cov.statements.add((self.module.file, n))
            self.check_statement(n, raw)
            before = self.current
            for stmt in statements(raw):
                code = code_only(stmt).strip()
                if code:
                    self.read_statement(n, stmt, code)
            owner = self.current or self.closed or before  # the End statement belongs to the procedure it closes
            if owner is not None:
                owner.size += len(raw.strip()) + 2
            if self.closed is not None:
                self.measure(self.closed)
                self.closed = None

    def check_statement(self, n: int, raw: str) -> None:
        """The rules that read a whole logical statement: parentheses and continuations."""
        header = PROC.match(code_only(raw).strip())
        where = header[3] if header else (self.current.name if self.current else '(module)')
        problem = paren_message(code_only(raw))
        if problem:
            self.raws.append((self.module.file, n, where, 'parentheses', problem, short(raw.strip())))
        if self.counts.get(n, 0) > CONTINUATION_LIMIT:
            self.raws.append((self.module.file, n, where, 'continuations',
                              f'{self.counts[n]} line continuations; VBA allows at most {CONTINUATION_LIMIT}',
                              short(raw.strip())))

    def read_statement(self, n: int, stmt: str, code: str) -> None:
        decl, proc, blk = DECLARE.match(code), PROC.match(code), BLOCK.match(code)
        if self.block and self.current is None:
            self.block = self.read_block(code)
        elif decl:
            params, _ = parse_header(code, decl.end())
            self.module.routines.setdefault(decl[4].lower(), []).append(
                Routine(decl[4], decl[1] is None or decl[1].lower() == 'public', params))
        elif proc:
            self.open_procedure(proc, code, stmt, n)
        elif END_PROCEDURE.fullmatch(code):
            self.close_procedure()
        elif blk and self.current is None:
            self.block = (blk[2].lower(), blk[1] is None or blk[1].lower() == 'public')
        elif self.current is not None:
            self.current.variables.update({name: var for name, var, _ in declarations(code)})
            self.current.statements.append(Stmt(n, stmt.strip(), code))
        else:
            for name, var, public in declarations(code):
                self.module.variables[name] = var
                if public:
                    self.module.public.add(name)

    def read_block(self, code: str) -> tuple[str, bool] | None:
        """Consume one line of a Type or Enum block; return the block that continues."""
        assert self.block is not None
        if END_BLOCK.match(code):
            return None
        member = re.match(r'(\w+)', code)
        if self.block[0] == 'enum' and member:
            self.module.variables[member[1].lower()] = Variable('long', 'Long', const=True)
            if self.block[1]:
                self.module.public.add(member[1].lower())
        return self.block

    def open_procedure(self, proc: re.Match[str], code: str, stmt: str, line: int) -> None:
        params, result = parse_header(code, proc.end())
        kind = proc[2].split()[0].lower()
        current = Procedure(proc[3], kind, line, short(stmt.strip()), result)
        for param in params or ():
            current.variables[param.name.lower()] = Variable(param.type, param.text, param.array)
        self.cov.procedures.add((self.module.file, line))
        if kind != 'property':
            self.module.routines.setdefault(proc[3].lower(), []).append(
                Routine(proc[3], proc[1] is None or proc[1].lower() != 'private', params))
        self.module.procedures.append(current)
        self.current = current

    def close_procedure(self) -> None:
        """End a procedure; its size is judged once the End line itself is counted."""
        self.closed, self.current = self.current, None

    def measure(self, current: Procedure) -> None:
        self.cov.sizes[(self.module.file, current.line)] = current.size
        if current.size >= PROCEDURE_LIMIT:
            self.raws.append((self.module.file, current.line, current.name, 'procedure-size',
                              f'{current.size} bytes of active code; VBA limits the compiled procedure to less '
                              f'than {PROCEDURE_LIMIT} bytes, and the active code is a conservative proxy of it',
                              current.header))


def leaves(code: str) -> list[str]:
    """The statements of an inline If, or the statement itself; none for a block If."""
    if re.match(r'if\b', code, re.I):
        m = THEN.search(code)
        tail = code[m.end():].strip() if m else ''
        e = ELSE.search(tail)
        parts = [tail[:e.start()], tail[e.end():]] if e else [tail]
        return [leaf for part in parts if part.strip() for leaf in leaves(part.strip())]
    if re.match(r'elseif\b', code, re.I):
        return []
    m = re.match(r'else\b\s*(.*)', code, re.I | re.S)
    if m:
        return leaves(m[1].strip()) if m[1].strip() else []
    return [code]


def classify(code: str) -> Event | None:
    """Return (event, block kind) for a statement that opens, closes or divides a block.

    Raises LexicalError when the statement cannot be classified with certainty.
    """
    label = re.fullmatch(r'(\w+):', code)
    if label and label[1].lower() in KEYWORD_LABEL:
        code = label[1]
    low = code.lower()
    word = re.match(r'[a-z_]\w*', low)
    first = word[0] if word else ''
    then = THEN.search(code)
    if first == 'if':
        if not then:
            raise LexicalError('If without Then')
        return ('open', 'if') if not code[then.end():].strip() else ('inline', 'if')
    if first == 'elseif':
        if not then or code[then.end():].strip():
            raise LexicalError('ElseIf that the text does not settle')
        return 'mid', 'if'
    if first == 'else':
        if low != 'else':
            raise LexicalError('Else followed by a statement')
        return 'mid', 'if'
    ended = re.match(r'end\s+(if|select|with)\b', low)
    if ended:
        return 'close', ended[1]
    if re.match(r'select\s+case\b', low):
        return 'open', 'select'
    return SIMPLE_BLOCKS.get(first)


def next_count(code: str) -> int:
    """How many loops one Next statement closes: ``Next i, j`` closes two."""
    names = re.sub(r'^next\b', '', code, flags=re.I).strip()
    return len(list(comma_parts(names))) if names else 1


def pair_problem(stack: list[tuple[str, int, str]], event: Event, code: str) -> str | None:
    """Why a closing or dividing statement does not fit the innermost open block."""
    shown = CLOSERS[event[1]] if event[0] == 'close' else code.split(None, 1)[0].rstrip(':')
    if not stack:
        return f'{shown} has no open {OPENERS[event[1]]}'
    kind, line, _ = stack[-1]
    if kind != event[1]:
        return f'{shown} does not match the innermost open block, {OPENERS[kind]} opened at line {line}'
    return None


def block_findings(proc: Procedure) -> list[tuple[int, str, str]]:
    """(line, message, statement) for each block of a procedure that does not pair up.

    The first mismatch is the only one reported, since the rest follows from
    it. A statement the text does not settle silences the whole procedure.
    """
    stack: list[tuple[str, int, str]] = []
    inline_line = -1
    for stmt in proc.statements:
        if stmt.line == inline_line:
            continue
        try:
            event = classify(stmt.code)
        except LexicalError:
            return []
        if event is None:
            continue
        if event[0] == 'inline':
            inline_line = stmt.line
            continue
        if event[0] == 'open':
            stack.append((event[1], stmt.line, stmt.code))
            continue
        for _ in range(next_count(stmt.code) if event == ('close', 'for') else 1):
            problem = pair_problem(stack, event, stmt.code)
            if problem:
                return [(stmt.line, problem, stmt.code)]
            if event[0] == 'close':
                stack.pop()
    return [(line, f'{OPENERS[kind]} opened at line {line} is never closed', text) for kind, line, text in stack]


def assignments(stmt: Stmt) -> Iterator[tuple[str, str]]:
    """(target, expression) of each simple assignment of a statement, inline If bodies included."""
    for leaf in leaves(stmt.code):
        m = ASSIGNMENT.fullmatch(leaf)
        if m:
            yield m[1], m[2]


def tokens(expression: str) -> list[tuple[str, str]] | None:
    """Tokens of an integer-arithmetic expression, or None for any other character."""
    out: list[tuple[str, str]] = []
    pos = 0
    while pos < len(expression):
        m = EXPRESSION_TOKEN.match(expression, pos)
        if not m:
            return None
        kind, value = m.lastgroup or '', m.group()
        if kind == 'num' and not value.startswith('&') and re.search(r'[.Ee!#@]', value):
            return None
        if kind != 'space':
            out.append((kind, value))
        pos = m.end()
    return out


class Scope:
    """Name resolution for one procedure of one module in one project."""

    def __init__(self, proc: Procedure, module: Module, project: Project) -> None:
        self.proc, self.module, self.project = proc, module, project

    def variable(self, name: str) -> Variable | None:
        key = name.lower()
        return self.proc.variables.get(key) or self.module.variables.get(key) or self.project.symbols.get(key)


def longptr_operand(expression: str, scope: Scope) -> str | None:
    """The LongPtr variable that the integer arithmetic of ``expression`` reads, or None.

    None whenever the expression holds anything but integer variables and
    constants, whole numbers, grouping parentheses and And, Or, Xor, Not, Mod,
    +, -, * and \\: a call, an array element, a member, a comparison, a
    concatenation, a division or a string. A conversion such as CLng(...) is a
    call, so it makes the expression silent.
    """
    stream = tokens(expression)
    if stream is None:
        return None
    pointer: str | None = None
    previous, depth = 'op', 0
    for kind, value in stream:
        if kind == 'id' and value.lower() in ARITHMETIC_WORDS:
            kind = 'op'
        elif kind == 'id':
            var = scope.variable(value)
            if var is None or var.array or var.type not in INTEGER_TYPES:
                return None
            pointer = pointer or (value if var.type == 'longptr' else None)
        elif kind == 'open':
            if previous in ('id', 'num', 'close'):
                return None
            depth += 1
        elif kind == 'close':
            depth -= 1
            if depth < 0:
                return None
        previous = kind
    return pointer if depth == 0 else None


def longptr_findings(module: Module, project: Project, cov: Coverage) -> Iterator[Raw]:
    for proc in module.procedures:
        scope = Scope(proc, module, project)
        for stmt in proc.statements:
            for target, expression in assignments(stmt):
                var = scope.variable(target)
                if proc.kind == 'function' and target.lower() == proc.name.lower():
                    var = Variable(proc.result, proc.result or 'Variant')
                if var is None or var.array or var.const or var.type != 'long':
                    continue
                cov.assignments.add((module.file, stmt.line))
                pointer = longptr_operand(expression, scope)
                if pointer:
                    yield (module.file, stmt.line, proc.name, 'longptr-assignment',
                           f'{target} is declared Long but the expression reads {pointer}, a LongPtr; '
                           f'convert it with CLng or declare {target} As LongPtr', short(stmt.text))


def normalize(kind: str, config: Config) -> str:
    """The type a declared type is under one compilation: LongPtr is Long or LongLong."""
    if kind == 'longptr':
        return 'longlong' if config['vba7'] and config['win64'] else 'long'
    return kind


def resolve_routine(name: str, module: Module, project: Project) -> Routine | None:
    """The one procedure or Declare that a bare name calls, or None when it is unknown or ambiguous."""
    local = module.routines.get(name.lower())
    if local:
        return local[0] if len(local) == 1 else None
    others = [r for other in project.modules if other is not module
              for r in other.routines.get(name.lower(), []) if r.public]
    return others[0] if len(others) == 1 else None


def call_sites(code: str, scope: Scope) -> Iterator[tuple[Routine, str]]:
    """(routine, argument text) of each call whose arguments pass ByRef.

    ``Foo(x)`` as a whole statement passes x by value, so it is skipped; so are
    calls through a qualifier and calls of a name that is also a variable.
    ``Call Foo(x)``, ``Foo x, y`` and a function call inside an expression
    pass by reference.
    """
    for m in CALL_OPEN.finditer(code):
        before = code[:m.start()].rstrip()
        if not before or re.search(r'\b(?:Then|Else)$', before, re.I) or scope.variable(m[1]):
            continue
        close = matching_paren(code, m.end() - 1)
        routine = resolve_routine(m[1], scope.module, scope.project)
        if close is not None and routine and routine.params is not None:
            yield routine, code[m.end():close]
    for leaf in leaves(code):
        m2 = STATEMENT_CALL.fullmatch(leaf)
        routine = resolve_routine(m2[1], scope.module, scope.project) if m2 and not scope.variable(m2[1]) else None
        if m2 and routine and routine.params is not None:
            yield routine, m2[2]


def compare_arguments(routine: Routine, args: str, stmt: Stmt, scope: Scope, config: Config,
                      cov: Coverage) -> Iterator[Raw]:
    params = routine.params or ()
    by_name = {p.name.lower(): p for p in params}
    for index, argument in enumerate(comma_parts(args) if args.strip() else ()):
        named = re.fullmatch(r'\s*(\w+)\s*:=\s*(.*)', argument, re.S)
        value = (named[2] if named else argument).strip()
        param = by_name.get(named[1].lower()) if named else (params[index] if index < len(params) else None)
        if param is None or param.paramarray:
            return
        var = scope.variable(value) if IDENTIFIER.fullmatch(value) else None
        if (var is None or var.const or var.array or param.array or not param.byref
                or var.type not in SCALAR_TYPES or param.type not in SCALAR_TYPES):
            continue
        cov.arguments.add((scope.module.file, stmt.line, index))
        if normalize(var.type or '', config) != normalize(param.type or '', config):
            yield (scope.module.file, stmt.line, scope.proc.name, 'byref-type',
                   f'{value} is declared {var.text} but {routine.name} takes {param.name} ByRef As '
                   f'{param.text}', short(stmt.text))


def byref_findings(module: Module, project: Project, config: Config, cov: Coverage) -> Iterator[Raw]:
    for proc in module.procedures:
        scope = Scope(proc, module, project)
        for stmt in proc.statements:
            for routine, args in call_sites(stmt.code, scope):
                yield from compare_arguments(routine, args, stmt, scope, config, cov)


def project_symbols(modules: Sequence[Module]) -> dict[str, Variable]:
    """Public module-level variables, constants and enum members; a name two modules declare is dropped."""
    seen: dict[str, Variable] = {}
    ambiguous: set[str] = set()
    for module in modules:
        for name in module.public:
            if name in seen:
                ambiguous.add(name)
            seen[name] = module.variables[name]
    return {name: var for name, var in seen.items() if name not in ambiguous}


def load_projects(manifest: Mapping[str, object]) -> dict[str, list[str]]:
    """Module paths of each configuration of tools/vba-projects.json."""
    configurations = manifest.get('configurations')
    projects: dict[str, list[str]] = {}
    if isinstance(configurations, dict):
        for name, entry in configurations.items():
            modules = entry.get('modules') if isinstance(entry, dict) else None
            if isinstance(modules, list):
                projects[str(name)] = [str(rel) for rel in modules]
    return projects


def read_modules(sources: Mapping[str, str], config_name: str, config: Config, cov: Coverage,
                 merged: dict[Raw, set[str]]) -> dict[str, Module]:
    """Read every module in one compilation; record the findings that need no other module."""
    parsed: dict[str, Module] = {}
    for rel, text in sources.items():
        counts, errors = continuation_counts(text)
        if errors:
            continue
        reader = Reader(rel, counts, cov)
        try:
            reader.read(text, config)
        except LexicalError as exc:
            found = re.search(r'line (\d+)', str(exc))
            merged.setdefault((rel, int(found[1]) if found else 1, '(module)', 'unreadable',
                               f'the conditional compilation could not be read: {exc}', ''),
                              set()).add(config_name)
            continue
        parsed[rel] = reader.module
        cov.modules.add(rel)
        for raw in reader.raws:
            merged.setdefault(raw, set()).add(config_name)
        for proc in reader.module.procedures:
            for line, message, statement in block_findings(proc):
                merged.setdefault((rel, line, proc.name, 'blocks', message, short(statement)), set()).add(config_name)
    return parsed


def read_projects(layout: Mapping[str, Sequence[str]], parsed: Mapping[str, Module], config_name: str,
                  config: Config, cov: Coverage, merged: dict[Raw, set[str]]) -> None:
    """Run the rules that need the declarations of the other modules of a project."""
    for files in layout.values():
        project = Project([parsed[rel] for rel in files if rel in parsed])
        project.symbols = project_symbols(project.modules)
        for module in project.modules:
            for raw in byref_findings(module, project, config, cov):
                merged.setdefault(raw, set()).add(config_name)
            if config_name == LONGPTR_CONFIG:
                for raw in longptr_findings(module, project, cov):
                    merged.setdefault(raw, set()).add(config_name)


def findings(sources: Mapping[str, str], projects: Mapping[str, Sequence[str]] | None = None
             ) -> tuple[list[Finding], dict[str, Coverage]]:
    """Findings for every module in every compilation, and what each compilation read.

    ``projects`` maps a configuration name to the modules of one VBA project.
    A module that no project lists is read as a project of its own, and with
    no ``projects`` every source belongs to one project.
    """
    layout = {name: [rel for rel in files if rel in sources]
              for name, files in (projects or {'project': list(sources)}).items()}
    listed = {rel for files in layout.values() for rel in files}
    layout.update({rel: [rel] for rel in sources if rel not in listed})
    merged: dict[Raw, set[str]] = {}
    for rel, text in sources.items():
        for n, raw in continuation_counts(text)[1]:
            merged.setdefault((rel, n, '(module)', 'lexical', 'a string or bracketed identifier never closes',
                               short(raw)), set()).update(CONFIGS)
    coverage = {name: Coverage() for name in CONFIGS}
    for name, config in CONFIGS.items():
        parsed = read_modules(sources, name, config, coverage[name], merged)
        read_projects(layout, parsed, name, config, coverage[name], merged)
    ordered = [Finding(file, line, proc, rule, message, statement, tuple(c for c in CONFIGS if c in names))
               for (file, line, proc, rule, message, statement), names in merged.items()]
    return sorted(ordered), coverage


def run(root: str | Path, modules: Iterable[str], projects: Mapping[str, Sequence[str]]
        ) -> tuple[list[Finding], dict[str, Coverage]]:
    """Read the given repository-relative modules, grouped into the projects of the manifest."""
    sources = {rel: (Path(root) / rel).read_bytes().decode('ascii', 'replace') for rel in modules}
    return findings(sources, projects)


def summary(coverage: Mapping[str, Coverage]) -> list[str]:
    """One line per compilation: what the rules read."""
    lines = []
    for name, cov in coverage.items():
        largest = max(cov.sizes.values(), default=0)
        lines.append(f'{name}: {len(cov.modules)} modules, {len(cov.procedures)} procedures, '
                     f'{len(cov.statements)} logical statements; {len(cov.assignments)} assignments to '
                     f'a Long variable, {len(cov.arguments)} ByRef arguments compared; '
                     f'largest procedure {largest} of {PROCEDURE_LIMIT} bytes of active code')
    return lines


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    projects = load_projects(json.loads((root / 'tools' / 'vba-projects.json').read_text(encoding='utf-8')))
    found, coverage = run(root, sorted({rel for files in projects.values() for rel in files}), projects)
    for item in found:
        print(item)
    for line in summary(coverage):
        print(line)
    return 1 if found else 0


if __name__ == '__main__':
    raise SystemExit(main())
