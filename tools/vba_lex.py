"""Shared, conservative physical-line lexer; not a VBA parser or compiler.

Regions preserve their original spelling. Strings and bracketed identifiers
are opaque. Rem starts a comment only at a statement boundary.
"""
import re


class LexicalError(ValueError):
    """Input cannot safely be transformed."""


def regions(line):
    out, start, i = [], 0, 0
    boundary = True
    while i < len(line):
        ch = line[i]
        if ch == "'" or (boundary and re.match(r'Rem(?:\s|$)', line[i:], re.I)):
            if i > start:
                out.append(('code', line[start:i]))
            out.append(('comment', line[i:]))
            return out
        if ch in '\"[':
            kind = 'literal' if ch != '[' else 'identifier'
            close = {'"': '"', '[': ']'}[ch]
            if i > start:
                out.append(('code', line[start:i]))
            j = i + 1
            while j < len(line):
                if line[j] == close:
                    if ch == '"' and j + 1 < len(line) and line[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            if j == len(line):
                raise LexicalError('unterminated string or bracketed identifier')
            out.append((kind, line[i:j + 1]))
            i = start = j + 1
            boundary = False
            continue
        if ch == ':' and not line[i:i + 2] == ':=':
            boundary = True
        elif not ch.isspace():
            # A word at a statement boundary consumes that boundary. Then/Else
            # open an inline statement; member names never do.
            m = re.match(r'[A-Za-z_]\w*', line[i:])
            if m:
                word = m.group()
                boundary = word.lower() in ('then', 'else') and (i == 0 or line[i-1] != '.')
                i += len(word)
                continue
            boundary = False
        i += 1
    if start < len(line):
        out.append(('code', line[start:]))
    return out


def split_code_comment(line):
    parts = regions(line)
    return (''.join(s for k, s in parts if k != 'comment'),
            ''.join(s for k, s in parts if k == 'comment'))


def code_only(line):
    """Mask literals/comments without changing column offsets for analysis."""
    return ''.join(s if k == 'code' else ' ' * len(s) for k, s in regions(line))


def transform_code(line, fn):
    return ''.join(fn(s) if k == 'code' else s for k, s in regions(line))


def protected_lines(lines):
    """Preserve conditional blocks and complete continued logical statements."""
    protected, notes = set(), []
    depth, continuing = 0, False
    for n, line in enumerate(lines):
        code = split_code_comment(line)[0].strip()
        directive = line.lstrip().startswith('#') and bool(re.match(
            r'#(?:If|ElseIf|Else|End|Const)\b', line.lstrip(), re.I))
        if re.match(r'#If\b', code, re.I):
            depth += 1
        if depth or directive or continuing or code.endswith(' _') or code == '_':
            protected.add(n)
        if directive:
            notes.append(f'line {n+1}: conditional compilation preserved')
        if continuing and (not code or directive):
            raise LexicalError(f'line {n+1}: unsupported interrupted continuation')
        continuing = bool(re.search(r'\s_$', code)) or code == '_'
        if continuing:
            notes.append(f'line {n+1}: continued statement preserved (not reformatted)')
        if re.match(r'#End\s+If\b', code, re.I):
            depth -= 1
            if depth < 0:
                raise LexicalError('unmatched #End If')
    if depth or continuing:
        raise LexicalError('unterminated conditional block or continuation')
    return protected, notes


TOKEN = re.compile(r'[A-Za-z_]\w*|\d+(?:\.\d+)?|[^\s]')


def executable_signature(text, label_map=None):
    """Token sequence including statement boundaries, opaque literal bytes,
    directives and options; only declared house-label substitutions permitted.
    Whitespace changes never merge two identifier tokens unnoticed.
    """
    label_map = label_map or {}
    rows, options = [], []
    depth, in_proc = 0, False
    for line in text.replace('\r\n', '\n').split('\n'):
        tokens = []
        for kind, value in regions(line):
            if kind == 'comment':
                continue
            if kind != 'code':
                tokens.append((kind, value))
            else:
                tokens.extend(('code', t) for t in TOKEN.findall(value))
        if not tokens:
            continue
        for i, (kind, value) in enumerate(tokens):
            if kind != 'code' or value not in label_map:
                continue
            is_label = i == 0 and len(tokens) > 1 and tokens[1] == ('code', ':')
            is_jump = i > 0 and tokens[i-1][0] == 'code' and tokens[i-1][1].lower() in ('goto', 'resume')
            if is_label or is_jump:
                tokens[i] = ('code', label_map[value])
        statement = ' '.join(value for kind, value in tokens if kind == 'code')
        if re.match(r'# If\b', statement, re.I):
            depth += 1
        if re.match(r'(?:Public |Private |Friend )?(?:Sub|Function|Property)\b', statement, re.I):
            in_proc = True
        if tokens[0] == ('code', 'Option') and not depth and not in_proc:
            options.append(tokens)
        else:
            rows.append(tokens)
        if re.match(r'# End If\b', statement, re.I):
            depth -= 1
        if re.match(r'End (Sub|Function|Property)\b', statement, re.I):
            in_proc = False
    return options, rows
