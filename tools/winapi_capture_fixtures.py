"""#27 WinAPI error-capture contract: source rules and their fixtures.

Err.LastDllError belongs to the most recent Declare call, and Err.Clear or any
On Error statement resets it. A native error is therefore only trustworthy
when it is copied the moment the call returns. Each supported compilation is
checked separately, because the call usually sits in a VBA7 / Win64 split and
one arm can drift while the others stay correct:

  * Err.LastDllError is read only as ``<local> = Err.LastDllError``, and only
    in the statement directly after a Declare call;
  * every GetWindowLong[Ptr] / SetWindowLong[Ptr] call, whose zero result is
    ambiguous, copies its result to a local, is directly preceded by
    ``SetLastError 0`` and directly followed by that capture;
  * GetLastError is not declared, so there is one route to the error.

The contract covers the component in src/ and the regression helpers in test/.
The demo modules in demo/ are not yet held to it.

This is bounded source evidence, not proof of runtime behaviour: it does not
follow GoTo, and it cannot see an error replaced inside the call itself.
"""
import re
from collections.abc import Iterable, Iterator
from dataclasses import dataclass, field
from pathlib import Path

from vba_analyze import CONFIGS, DECLARE, PROC, VAR, active_lines, names, statements
from vba_lex import LexicalError, code_only

SCOPE = ('src/', 'test/')
AMBIGUOUS_ZERO = frozenset({
    'getwindowlonga', 'getwindowlongw', 'getwindowlongptra', 'getwindowlongptrw',
    'setwindowlonga', 'setwindowlongw', 'setwindowlongptra', 'setwindowlongptrw',
})
LAST_DLL_ERROR = re.compile(r'\bErr\s*\.\s*LastDllError\b', re.I)
CAPTURE = re.compile(r'(?:Let\s+)?(\w+)\s*=\s*Err\s*\.\s*LastDllError', re.I)
ON_ERROR = re.compile(r'On\s+Error\b', re.I)
END_PROC = re.compile(r'End\s+(?:Sub|Function|Property)', re.I)


@dataclass
class Declared:
    """Declare names of one module in one compilation, keyed by VBA name."""
    exports: dict[str, str] = field(default_factory=dict)

    def export_of(self, name: str) -> str | None:
        return self.exports.get(name.lower())

    def clears(self, code: str) -> bool:
        """True for ``SetLastError 0`` or ``Call SetLastError(0)`` by any alias."""
        m = re.fullmatch(r'(?:Call\s+)?(\w+)\s*(?:\(\s*0\s*\)|\s0)', code, re.I)
        return m is not None and self.export_of(m[1]) == 'setlasterror'

    def called(self, code: str) -> list[str]:
        """Declared names this statement invokes, in order."""
        return [m[0] for m in re.finditer(r'\b\w+\b', code) if m[0].lower() in self.exports]


@dataclass
class Procedure:
    name: str
    locals: set[str] = field(default_factory=set)
    body: list[tuple[int, str]] = field(default_factory=list)


def scan(text: str, config: dict[str, bool]) -> tuple[Declared, list[tuple[int, str]], list[Procedure]]:
    """Return declarations, GetLastError declarations and procedure bodies."""
    declared = Declared()
    get_last_error: list[tuple[int, str]] = []
    procedures: list[Procedure] = []
    current: Procedure | None = None
    for n, raw in active_lines(text, config):
        for statement in statements(raw):
            code = code_only(statement).strip()
            if not code:
                continue
            decl = DECLARE.match(code)
            if decl:
                alias = re.search(r'\bAlias\s+"([^"]+)"', statement, re.I)
                export = (alias[1] if alias else decl[4]).lower()
                declared.exports[decl[4].lower()] = export
                if export == 'getlasterror':
                    get_last_error.append((n, decl[4]))
                continue
            proc = PROC.match(code)
            if proc:
                current = Procedure(proc[3])
                procedures.append(current)
                continue
            if END_PROC.fullmatch(code):
                current = None
                continue
            if current is None:
                continue
            var = VAR.match(code)
            if var and re.match(r'(?:Dim|Static)\b', code, re.I):
                current.locals |= names(var[2] or var[3] or '')
                continue
            if re.fullmatch(r'Const\s+.*', code, re.I):
                continue
            current.body.append((n, code))
    return declared, get_last_error, procedures


def procedure_findings(procedure: Procedure, declared: Declared) -> Iterator[tuple[int, str]]:
    """Yield (line, message) for one procedure in one compilation."""
    body = procedure.body

    def is_api_call(code: str) -> bool:
        called = declared.called(code)
        return bool(called) and declared.export_of(called[0]) != 'setlasterror'

    for i, (n, code) in enumerate(body):
        before = body[i - 1][1] if i else ''
        after = body[i + 1][1] if i + 1 < len(body) else ''
        if LAST_DLL_ERROR.search(code):
            capture = CAPTURE.fullmatch(code)
            if not capture or capture[1].lower() not in procedure.locals:
                yield n, 'Err.LastDllError must be copied by itself to a local'
            if not before:
                yield n, 'Err.LastDllError read before any WinAPI call'
            elif ON_ERROR.match(before):
                yield n, f'Err.LastDllError read after an On Error statement ({before})'
            elif not is_api_call(before):
                yield n, f'Err.LastDllError read after an intervening statement ({before})'
        for name in declared.called(code):
            if declared.export_of(name) not in AMBIGUOUS_ZERO:
                continue
            result = re.fullmatch(r'(?:Let\s+)?(\w+)\s*=\s*' + re.escape(name) + r'\s*\(.*\)', code, re.I)
            if not result or result[1].lower() not in procedure.locals:
                yield n, f'{name} result must be copied to a local in the call statement'
            if not declared.clears(before):
                yield n, f'{name} is not immediately preceded by SetLastError 0'
            if not CAPTURE.fullmatch(after):
                yield n, f'{name} is not immediately followed by its Err.LastDllError capture'


def findings(sources: dict[str, str]) -> tuple[list[str], dict[str, int]]:
    """Findings for every module in every compilation, plus calls checked per compilation."""
    errors: set[str] = set()
    checked = dict.fromkeys(CONFIGS, 0)
    for config_name, config in CONFIGS.items():
        for file, text in sources.items():
            try:
                declared, get_last_error, procedures = scan(text, config)
            except LexicalError as exc:
                errors.add(f'{file}: {config_name}: {exc}')
                continue
            for n, name in get_last_error:
                errors.add(f'{file}:{n}: {config_name}: declares {name}; capture Err.LastDllError instead')
            for procedure in procedures:
                for n, message in procedure_findings(procedure, declared):
                    errors.add(f'{file}:{n}: {config_name}: {procedure.name}: {message}')
                checked[config_name] += sum(
                    1 for _, code in procedure.body for name in declared.called(code)
                    if declared.export_of(name) in AMBIGUOUS_ZERO)
    return sorted(errors), checked


def run(root: str | Path, modules: Iterable[str]) -> list[str]:
    """Check the tracked component and regression modules, and refuse a tree with nothing to check."""
    sources = {rel: (Path(root) / rel).read_text(encoding='ascii')
               for rel in modules if rel.startswith(SCOPE)}
    errors, checked = findings(sources)
    for config_name, count in checked.items():
        if count == 0:
            errors.append(f'{config_name}: no GetWindowLong/SetWindowLong call found; '
                          'the contract would pass without checking anything')
    return errors


# A compliant module: a Win64 split for the ambiguous read, a BOOL call whose
# error is still reported, and a seam taken before the capture sequence.
BASE = '''Attribute VB_Name = "M_FIXTURE"
Option Explicit
#If VBA7 Then
    #If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    #Else
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    #End If
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Sub FX_SetLastError Lib "kernel32" Alias "SetLastError" (ByVal dwErrCode As Long)
#Else
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare Sub FX_SetLastError Lib "kernel32" Alias "SetLastError" (ByVal dwErrCode As Long)
#End If

Private Function ReadStyle(ByVal hWnd As Long, ByRef StyleOut As Long, ByRef FailMsg As String) As Boolean
    Dim ApiResult As Long
    Dim DllError As Long
    Dim Fault As Long
    On Error GoTo Err_Handler
    Fault = TakeFault()
    If Fault = 3 Then
        FX_SetLastError 1234
    End If
    FX_SetLastError 0
#If VBA7 Then
    #If Win64 Then
    ApiResult = GetWindowLongPtr(hWnd, -16)
    #Else
    ApiResult = GetWindowLong(hWnd, -16)
    #End If
#Else
    ApiResult = GetWindowLong(hWnd, -16)
#End If
    DllError = Err.LastDllError
    If ApiResult = 0 And DllError <> 0 Then
        FailMsg = Describe("read", DllError)
        GoTo Safe_Exit
    End If
    StyleOut = ApiResult
    ReadStyle = True
Safe_Exit:
    Exit Function
Err_Handler:
    FailMsg = "unexpected"
End Function

Private Function Refresh(ByVal hWnd As Long, ByRef FailMsg As String) As Boolean
    Dim ApiOK As Long
    Dim DllError As Long
    On Error GoTo Err_Handler
    Call FX_SetLastError(0)
    ApiOK = SetWindowPos(hWnd, 0, 0, 0, 0, 0, &H27)
    DllError = Err.LastDllError
    If ApiOK = 0 Then
        FailMsg = Describe("refresh", DllError)
    End If
    Refresh = (ApiOK <> 0)
    Exit Function
Err_Handler:
    FailMsg = "unexpected"
End Function
'''

CAPTURE_LINE = '    DllError = Err.LastDllError\n    If ApiResult'
X64_CALL = '    ApiResult = GetWindowLongPtr(hWnd, -16)\n'
LEGACY_ARM = '#Else\n    ApiResult = GetWindowLong(hWnd, -16)\n#End If'
GET_LAST_ERROR = 'Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long\n'

# (name, source, substring every finding must contain, configurations that
# must report it). An empty tuple of configurations means "no finding".
CASES: list[tuple[str, str, str, tuple[str, ...]]] = [
    ('the compliant module passes', BASE, '', ()),
    ('a helper call between the call and the capture is rejected',
     BASE.replace(CAPTURE_LINE, '    FailMsg = Describe("read", 0)\n' + CAPTURE_LINE),
     'read after an intervening statement', tuple(CONFIGS)),
    ('an On Error statement between the call and the capture is rejected',
     BASE.replace(CAPTURE_LINE, '    On Error Resume Next\n' + CAPTURE_LINE),
     'read after an On Error statement', tuple(CONFIGS)),
    ('a capture moved into the failure branch is rejected',
     BASE.replace(CAPTURE_LINE, '    If ApiResult')
         .replace('        FailMsg = Describe("read", DllError)',
                  '        DllError = Err.LastDllError\n        FailMsg = Describe("read", DllError)'),
     'Err.LastDllError', tuple(CONFIGS)),
    ('a BOOL call reported after formatting is rejected',
     BASE.replace('    DllError = Err.LastDllError\n    If ApiOK = 0 Then\n        FailMsg = Describe("refresh", DllError)',
                  '    If ApiOK = 0 Then\n        FailMsg = Describe("refresh", 0)\n        DllError = Err.LastDllError'),
     'read after an intervening statement', tuple(CONFIGS)),
    ('a missing clear is rejected',
     BASE.replace('    FX_SetLastError 0\n', ''),
     'not immediately preceded by SetLastError 0', tuple(CONFIGS)),
    ('a clear that is not adjacent to the call is rejected',
     BASE.replace('    FX_SetLastError 0\n', '    FX_SetLastError 0\n    DllError = 0\n'),
     'not immediately preceded by SetLastError 0', tuple(CONFIGS)),
    ('a non-zero clear is rejected',
     BASE.replace('    FX_SetLastError 0\n', '    FX_SetLastError 5\n'),
     'not immediately preceded by SetLastError 0', tuple(CONFIGS)),
    ('a legacy arm that drifts is rejected in that compilation only',
     BASE.replace(LEGACY_ARM, '#Else\n    DllError = 0\n    ApiResult = GetWindowLong(hWnd, -16)\n#End If'),
     'not immediately preceded by SetLastError 0', ('legacy-x86',)),
    ('a 64-bit arm that drifts is rejected in that compilation only',
     BASE.replace(X64_CALL, X64_CALL + '    FailMsg = vbNullString\n'),
     'ReadStyle', ('vba7-x64',)),
    ('a result written straight to a ByRef parameter is rejected',
     BASE.replace(X64_CALL, '    StyleOut = GetWindowLongPtr(hWnd, -16)\n'),
     'result must be copied to a local', ('vba7-x64',)),
    ('a capture inside an expression is rejected',
     BASE.replace(CAPTURE_LINE, '    DllError = Err.LastDllError + 0\n    If ApiResult'),
     'Err.LastDllError', tuple(CONFIGS)),
    ('a GetLastError declaration is rejected',
     BASE.replace('#Else\nPrivate Declare Function GetWindowLong',
                  GET_LAST_ERROR + '#Else\nPrivate Declare Function GetWindowLong'),
     'declares GetLastError', ('vba7-x64', 'vba7-x86')),
]


def selftest() -> tuple[list[str], int]:
    """Every fixture must report exactly the configurations it names."""
    errors: list[str] = []
    for name, source, expected, configs in CASES:
        found, _ = findings({'fixture.bas': source})
        if not configs:
            if found:
                errors.append(f'#27 fixture {name}: unexpected {found}')
            continue
        named = {config for config in CONFIGS if any(f': {config}: ' in f and expected in f for f in found)}
        anywhere = {config for config in CONFIGS if any(f': {config}: ' in f for f in found)}
        if named != set(configs) or anywhere != set(configs):
            errors.append(f'#27 fixture {name}: expected {expected!r} in {configs}, got {found}')
    return errors, len(CASES)
