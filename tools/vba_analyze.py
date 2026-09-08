"""Bounded VBA source analyzer for #31. No Excel/type/dynamic-call resolution.

Analyze each supported Windows compilation independently, using the shared
physical-line lexer. Findings carry physical start lines, never generated-line
numbers. Unknown conditional symbols fail explicitly rather than hiding code.
"""
from dataclasses import dataclass, field
import re
from pathlib import Path
from vba_lex import code_only, split_code_comment, LexicalError

CONFIGS = {'vba7-x64': {'vba7': True, 'win64': True},
           'vba7-x86': {'vba7': True, 'win64': False},
           'legacy-x86': {'vba7': False, 'win64': False}}
PREFIX = re.compile(r'\b(?:UI_|TST_|Test_EXCEL_UI_|Demo_|DEMO_)\w+', re.I)
PROC = re.compile(r'^(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\b', re.I)
DECLARE = re.compile(r'^(?:(Public|Private)\s+)?Declare\s+(PtrSafe\s+)?(Sub|Function)\s+(\w+)\b', re.I)
VAR = re.compile(r'^(?:(Public|Private|Dim|Static|Global)\s+)(?:Const\s+)?(.+)$|^Const\s+(.+)$', re.I)
BLOCK = re.compile(r'^(?:(Public|Private)\s+)?(Type|Enum)\s+(\w+)', re.I)
ERR_READ = re.compile(r'\bErr\s*\.\s*(?:Number|Source|Description)\b', re.I)


@dataclass(frozen=True, order=True)
class Finding:
    file: str
    line: int
    rule: str
    message: str

    def __str__(self):
        return f'{self.file}:{self.line}: {self.rule}: {self.message}'


@dataclass
class Procedure:
    name: str
    kind: str
    line: int
    public: bool
    params: set = field(default_factory=set)
    locals: set = field(default_factory=set)
    statements: list = field(default_factory=list)


def expression(text, symbols):
    """Restricted boolean conditional grammar; no eval or unknown-symbol guess."""
    tokens = re.findall(r'\w+|<>|[()=]|\S', text.lower())
    pos = 0

    def atom():
        nonlocal pos
        if pos >= len(tokens):
            raise LexicalError('incomplete conditional expression')
        t = tokens[pos]; pos += 1
        if t == 'not':
            return not atom()
        if t == '(':
            value = disjunction()
            if pos >= len(tokens) or tokens[pos] != ')':
                raise LexicalError('unclosed conditional parenthesis')
            pos += 1
            return value
        if t in ('true', 'false', '0', '1'):
            return t in ('true', '1')
        if t in symbols:
            return symbols[t]
        raise LexicalError(f'unsupported conditional symbol {t}')

    def comparison():
        nonlocal pos
        value = atom()
        if pos < len(tokens) and tokens[pos] in ('=', '<>'):
            op = tokens[pos]; pos += 1
            other = atom()
            value = (value == other) if op == '=' else (value != other)
        return value

    def conjunction():
        nonlocal pos
        value = comparison()
        while pos < len(tokens) and tokens[pos] == 'and':
            pos += 1
            other = comparison(); value = value and other
        return value

    def disjunction():
        nonlocal pos
        value = conjunction()
        while pos < len(tokens) and tokens[pos] in ('or', 'xor'):
            op = tokens[pos]; pos += 1
            other = conjunction()
            value = value != other if op == 'xor' else value or other
        return value

    value = disjunction()
    if pos != len(tokens):
        raise LexicalError('unsupported conditional expression: ' + text)
    return value


def logical_lines(text):
    pending, first = '', 0
    for n, raw in enumerate(text.splitlines(), 1):
        raw = split_code_comment(raw)[0]
        if pending and not raw.strip():
            raise LexicalError(f'line {n}: interrupted continuation')
        if not pending:
            first = n
        continued = bool(re.search(r'\s_\s*$', raw))
        pending += (re.sub(r'\s_\s*$', ' ', raw) if continued else raw)
        if not continued:
            if pending.strip():
                yield first, pending.strip()
            pending = ''
    if pending:
        raise LexicalError('unterminated continuation')


def active_lines(text, config):
    stack, enabled, symbols = [], True, dict(config)
    for n, raw in logical_lines(text):
        code = code_only(raw).strip()
        m = re.match(r'#(If|ElseIf)\s+(.+?)\s+Then$', code, re.I)
        if m:
            try:
                truth = expression(m[2], symbols)
            except LexicalError as exc:
                raise LexicalError(f'line {n}: {exc}') from exc
            if m[1].lower() == 'if':
                stack.append([enabled, truth, False])
                enabled = enabled and truth
            else:
                if not stack or stack[-1][2]:
                    raise LexicalError(f'line {n}: unmatched #ElseIf')
                parent, taken, _ = stack[-1]
                enabled = parent and not taken and truth
                stack[-1][1] = taken or truth
            continue
        if re.fullmatch(r'#Else', code, re.I):
            if not stack or stack[-1][2]:
                raise LexicalError(f'line {n}: unmatched #Else')
            parent, taken, _ = stack[-1]
            enabled = parent and not taken
            stack[-1][1:] = [True, True]
            continue
        if re.fullmatch(r'#End\s+If', code, re.I):
            if not stack:
                raise LexicalError(f'line {n}: unmatched #End If')
            enabled = stack.pop()[0]
            continue
        m = re.match(r'#Const\s+(\w+)\s*=\s*(.*)', code, re.I)
        if m:
            if enabled:
                symbols[m[1].lower()] = expression(m[2], symbols)
            continue
        if code.startswith('#'):
            raise LexicalError(f'line {n}: unsupported conditional directive')
        if enabled:
            yield n, raw
    if stack:
        raise LexicalError('unterminated #If')


def comma_parts(code):
    depth, start = 0, 0
    for i, ch in enumerate(code):
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        elif ch == ',' and depth == 0:
            yield code[start:i]
            start = i + 1
    yield code[start:]


def names(code):
    result = set()
    for part in comma_parts(code):
        part = re.sub(r'\b(?:Optional|ByVal|ByRef|ParamArray)\s+', '', part.strip(), flags=re.I)
        m = re.match(r'(\w+)', part)
        if m:
            result.add(m[1].lower())
    return result


def statements(raw):
    """Split colon statements, retaining labels; ignore := and opaque regions."""
    code = code_only(raw)
    start = 0
    for m in re.finditer(r':(?!=)', code):
        part = raw[start:m.start()].strip()
        if start == 0 and re.fullmatch(r'\w+', part):
            yield part + ':'
        elif part:
            yield part
        start = m.end()
    if raw[start:].strip():
        yield raw[start:].strip()


# Versioned native ABI contracts: ordered parameter types and result type.
# P = pointer-sized in VBA7, Long in legacy. A/W exports share the same types.
ABI = {
    'getwindowlongptra': (['P', 'Long'], 'P'),
    'setwindowlongptra': (['P', 'Long', 'P'], 'P'),
    'getwindowlonga': (['P', 'Long'], 'Long'),
    'setwindowlonga': (['P', 'Long', 'Long'], 'Long'),
    'setwindowpos': (['P', 'P', 'Long', 'Long', 'Long', 'Long', 'Long'], 'Long'),
    'iswindow': (['P'], 'Long'),
    'getlasterror': ([], 'Long'),
    'setlasterror': (['Long'], None),
}


# Pure scalar intrinsics in the capture expressions used here. A project
# function or host/member invocation is never assumed to preserve Err.
PURE_CAPTURE = {'iif', 'len', 'cstr', 'clng', 'cbool', 'cint', 'cdbl', 'csng'}


def error_operation(code):
    """Possible call/host operation, excluding scalar local copies/captures."""
    assignment = re.match(r'(?:(?:Set|Let)\s+)?[A-Za-z_]\w*\s*=(?!=)', code, re.I)
    expression_code = code[assignment.end():] if assignment else code
    calls = re.findall(r'([A-Za-z_]\w*)\s*\(', expression_code)
    if any(name.lower() not in PURE_CAPTURE for name in calls):
        return True
    # Err field reads are values, not calls; other object members can invoke
    # host getters/setters even without parentheses.
    members = re.sub(r'\bErr\s*\.\s*(Number|Source|Description)\b', '', expression_code, flags=re.I)
    if re.search(r'\.\s*[A-Za-z_]', members):
        return True
    if assignment:
        # Arithmetic/concatenation may overflow, divide by zero or allocate.
        # A scalar flag/copy alone does not refresh a cleared Err object.
        return bool(re.search(r'[&+*/^\\-]|\bMod\b', expression_code, re.I))
    return bool(re.match(r'(?:Call\s+)?[A-Za-z_]\w*(?:\s|$)', code))


def analyze(sources):
    findings, dynamic = set(), set()
    for config_name, config in CONFIGS.items():
        modules = {}
        project = set()
        exported_procs = {}
        for file, text in sources.items():
            def add(n, rule, message):
                findings.add(Finding(file, n, rule, f'{config_name}: {message}'))
            try:
                lines = [(n, part) for n, raw in active_lines(text, config)
                         for part in statements(raw)]
            except LexicalError as exc:
                match = re.search(r'line (\d+)', str(exc))
                add(int(match[1]) if match else 1, 'conditional/lexical', str(exc))
                continue
            module_names, procedures, current, block = set(), [], None, None
            explicit = any(re.fullmatch(r'Option\s+Explicit', code_only(raw).strip(), re.I) for _, raw in lines)
            seen = set()
            for n, raw in lines:
                code = code_only(raw).strip()
                decl, proc, blk = DECLARE.match(code), PROC.match(code), BLOCK.match(code)
                if decl:
                    name = decl[4].lower(); module_names.add(name)
                    if decl[1] is None or decl[1].lower() == 'public':
                        project.add(name)
                    if config['vba7'] and not decl[2]:
                        add(n, 'ptrsafe', f'{decl[4]} requires PtrSafe')
                    alias = re.search(r'\bAlias\s+"([^"]+)"', raw, re.I)
                    if PREFIX.match(decl[4]) and not alias:
                        add(n, 'alias', f'{decl[4]} requires explicit DLL Alias')
                    export = alias[1].lower() if alias else name
                    if export in ABI:
                        args = code[code.find('(')+1:code.rfind(')')]
                        actual = re.findall(r'\bAs\s+(\w+)', args, re.I)
                        tail = re.search(r'\bAs\s+(\w+)', code[code.rfind(')')+1:], re.I)
                        expected, result = ABI[export]
                        def resolved(t):
                            return ('LongPtr' if config['vba7'] else 'Long') if t == 'P' else t
                        if (any(not re.match(r'\s*ByVal\b', arg, re.I) for arg in comma_parts(args) if arg.strip())
                                or (decl[3].lower() == 'sub') != (result is None)
                                or [t.lower() for t in actual] != [resolved(t).lower() for t in expected] or (tail[1].lower() if tail else None) != (resolved(result).lower() if result else None)):
                            add(n, 'winapi-signature', f'{decl[4]} disagrees with {export} ABI')
                    continue
                if proc:
                    kind, name = proc[2].split()[0].lower(), proc[3].lower()
                    accessor = proc[2].lower()
                    key = (name, accessor if kind == 'property' else 'procedure')
                    if key in seen:
                        add(n, 'duplicate', f'duplicate {proc[3]}')
                    seen.add(key)
                    if current:
                        add(n, 'structure', f'{current.name} is not closed before {name}')
                    params = code[code.find('(')+1:code.rfind(')')] if '(' in code else ''
                    current = Procedure(name, kind, n, proc[1] is None or proc[1].lower() != 'private', names(params))
                    procedures.append(current); module_names.add(name)
                    if current.public:
                        project.add(name)
                        global_key = key
                        if global_key in exported_procs and exported_procs[global_key] != file:
                            add(n, 'duplicate', f'{name} also public in {exported_procs[global_key]}')
                        exported_procs[global_key] = file
                    continue
                end = re.fullmatch(r'End\s+(Sub|Function|Property)', code, re.I)
                if end:
                    if not current or current.kind != end[1].lower():
                        add(n, 'structure', f'unmatched {code}')
                    current = None
                    continue
                if blk and current is None:
                    block = (blk[2].lower(), blk[1] is None or blk[1].lower() == 'public')
                    module_names.add(blk[3].lower())
                    if block[1]: project.add(blk[3].lower())
                    continue
                if block:
                    if re.match(r'End\s+(Type|Enum)', code, re.I):
                        block = None
                    elif block[0] == 'enum':
                        m = re.match(r'(\w+)', code)
                        if m:
                            module_names.add(m[1].lower())
                            if block[1]: project.add(m[1].lower())
                    continue
                for statement in statements(raw):
                    masked = code_only(statement).strip()
                    var = VAR.match(masked)
                    if var:
                        defined = names(var[2] or var[3])
                        if current:
                            current.locals.update(defined)
                        else:
                            module_names.update(defined)
                            if re.match(r'(Public|Global)\b', masked, re.I): project.update(defined)
                    if current:
                        current.statements.append((n, statement))
            if current:
                add(current.line, 'structure', f'unclosed {current.name}')
            modules[file] = (explicit, module_names, procedures)
        for file, (explicit, module_names, procedures) in modules.items():
            def add(n, rule, message):
                findings.add(Finding(file, n, rule, f'{config_name}: {message}'))
            for proc in procedures:
                known = project | module_names | proc.params | proc.locals | {proc.name}
                labels = {code_only(raw).strip()[:-1].lower() for _, raw in proc.statements if re.fullmatch(r'\w+:', code_only(raw).strip())}
                known |= labels
                seen_labels = set()
                for n, raw in proc.statements:
                    label = code_only(raw).strip()
                    if re.fullmatch(r'\w+:', label):
                        if label.lower() in seen_labels:
                            add(n, 'label', f'{proc.name}: duplicate label {label}')
                        seen_labels.add(label.lower())
                handlers = set()
                for _, raw in proc.statements:
                    m = re.search(r'\bOn\s+Error\s+GoTo\s+(\w+)', code_only(raw), re.I)
                    if m and m[1] != '0': handlers.add(m[1].lower())
                stale, in_handler, resume_next = False, False, False
                for n, raw in proc.statements:
                    code = code_only(raw).strip()
                    if re.fullmatch(r'\w+:', code):
                        in_handler = code[:-1].lower() in handlers
                        stale, resume_next = False, False
                        continue
                    for m in re.finditer(r'\b(?:GoTo|Resume)\s+(\w+)', code, re.I):
                        if m[1].lower() not in labels | {'next', '0'}:
                            add(n, 'label', f'{proc.name}: undefined local label {m[1]}')
                    if re.search(r'\bApplication\s*\.\s*Run\b|\.\s*OnAction\b|\bCallByName\b', code, re.I):
                        dynamic.add(f'{file}:{n}: {proc.name}: dynamic dispatch requires host/binding review')
                    if not VAR.match(code):
                        for m in PREFIX.finditer(code):
                            if m[0].lower() in known or code[:m.start()].rstrip().endswith('.') or re.match(r'\s*:=', code[m.end():]):
                                continue
                            add(n, 'unresolved', f'{proc.name}: unresolved project reference {m[0]}')
                    # Only direct, simple assignment LHS; inline If/Else bodies
                    # are checked without treating their conditions as writes.
                    for part in re.split(r'\bThen\b|\bElse\b', code, flags=re.I):
                        assignment = re.match(r'^\s*(?:(?:Let|Set)\s+)?([A-Za-z_]\w*)\s*=(?!=)', part, re.I)
                        if explicit and assignment and assignment[1].lower() not in known:
                            add(n, 'undeclared-assignment', f'{proc.name}: undeclared {assignment[1]}')
                    if stale and ERR_READ.search(code):
                        add(n, 'stale-err', f'{proc.name}: Err read after handler state was overwritten')
                    if re.match(r'On\s+Error\b', code, re.I):
                        stale = True
                        resume_next = bool(re.search(r'Resume\s+Next', code, re.I))
                    elif re.match(r'(If|Else|End If|For|Next|Do|Loop|Select|Case|Resume|Exit)\b', code, re.I):
                        # Beyond a straight-line segment: no interprocedural or
                        # branch-merge claim; start tracking at the next handler.
                        stale = False
                    elif code and not VAR.match(code):
                        if error_operation(code):
                            if resume_next or not in_handler:
                                stale = False  # new operation is the error source
                            else:
                                stale = True
    return sorted(findings), sorted(dynamic)


def run(root):
    sources = {str(p.relative_to(root)): p.read_text(encoding='ascii')
               for directory in ('src', 'test', 'demo') for p in (Path(root)/directory).glob('*.bas')}
    return analyze(sources)


if __name__ == '__main__':
    findings, inventory = run(Path(__file__).resolve().parents[1])
    for item in findings + inventory:
        print(item)
    raise SystemExit(bool(findings))
