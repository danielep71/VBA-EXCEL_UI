"""#203 VBA syntax rules: one passing and one failing fixture for each rule.

Each case names the rule that must report, a fragment of its message and the
compilations that must report it; a case that names no rule must report
nothing. A rule that stays silent where it should not is as much a defect as a
rule that reports a statement that compiles, so the passing fixtures hold the
constructs that look like a defect and are not: strings and comments that
hold parentheses, quotes and underscores, ``Rem`` comments, ``Declare``
statements, ``Property`` procedures, line continuations, inline ``If``
statements and conversions that wrap a ``LongPtr``.

The two real defects are fixtures: ``FindingSummary)`` closing an argument
list that nothing opened (the harness at 51ce550), and
``FixtureOwned = EntryStyle And TST_TITLEBAR_OWNED_MASK`` assigning a
``LongPtr`` expression to a ``Long`` (the title-bar suite at 028c9c5).
"""
from collections.abc import Mapping, Sequence

from vba_analyze import CONFIGS
from vba_syntax import CONTINUATION_LIMIT, PROCEDURE_LIMIT, findings

ALL = tuple(CONFIGS)
X64 = ('vba7-x64',)
LEGACY = ('legacy-x86',)

# A module that compiles: every construct a rule could mistake for a defect.
BASE = '''Attribute VB_Name = "M_FIXTURE"
Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function FX_GetStyle Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function FX_Read Lib "fixture" (ByVal hWnd As LongPtr, ByRef Style As LongPtr) As Long
#Else
Private Declare Function FX_GetStyle Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function FX_Read Lib "fixture" (ByVal hWnd As Long, ByRef Style As Long) As Long
#End If

Private Const FX_MASK As Long = &HCF0000
Private Const FX_NOTE As String = "it's (a) 'note' _"
Private m_Total As Long

Private Type FX_Pair
    Left As Long
    Right As Long
End Type

Public Property Get FX_Total() As Long
    FX_Total = m_Total
End Property

Public Property Let FX_Total(ByVal Value As Long)
    m_Total = Value
End Property

Public Sub FX_Fill(ByRef Style As LongPtr, ByVal Flags As Long)
    Style = Flags
End Sub

Public Sub FX_Count(ByRef Total As Long)
    Total = Total + 1
End Sub

Public Function FX_Join(ByVal A As String, ByVal B As String) As String
    FX_Join = A & B
End Function

Public Function FX_Run(ByVal hWnd As Long, ByRef Msg As String) As Boolean
#If VBA7 Then
    Dim Style               As LongPtr
#Else
    Dim Style               As Long
#End If
    Dim Owned               As Long
    Dim Total               As Long
    Dim Index               As Long
    Dim Flag                As Boolean
    Dim Anything            As Variant
    Dim Text                As String
    Dim Pair                As FX_Pair
    Rem a Rem comment (
    Text = "one (" & "two" ' trailing comment )
    Text = "say ""hi"" (still text"
    Text = "ends with underscore _"
    Text = FX_NOTE
    If hWnd = 0 Then Exit Function
    If hWnd < 0 Then Msg = "bad": Exit Function
    If hWnd = 1 Then Total = 1 Else Total = 2
    FX_Fill Style, 1
    FX_Count Total
    Call FX_Count(Total)
    FX_Count (Total)
    FX_Count Anything
    Owned = CLng(Style And FX_MASK)
    Owned = CLng(Style) And FX_MASK
    Owned = FX_GetStyle(hWnd, 0) And FX_MASK
    Owned = Style / 2
    Anything = Style And FX_MASK
    Flag = (Style And FX_MASK) <> 0
    Owned = Total + 1
    For Index = 1 To 3
        Select Case Index
            Case 1
                Total = Total + 1
            Case Else
                Total = Total + 2
        End Select
    Next Index
    With Pair
        .Left = Total
    End With
    Do While Total > 0
        Total = Total - 1
    Loop
    If Total = 0 Then
        Msg = "zero"
    ElseIf Total < 0 Then
        Msg = "negative"
    Else
        Msg = "many"
    End If
    Msg = FX_Join("a", _
                  "b")
    Open Msg For Output As #1
    Close #1
    FX_Run = (Owned = 0)
End Function
'''

STYLE_DEFECT = '    Owned = Style And FX_MASK\n'
SECOND_MODULE = '''Attribute VB_Name = "M_OTHER"
Option Explicit

Public Const OT_MASK As Long = &HF0000
Public Sub OT_Narrow(ByRef Small As Integer)
    Small = 1
End Sub
Private Sub OT_Hidden(ByRef Small As Integer)
    Small = 2
End Sub
'''


def swap(old: str, new: str, source: str = BASE) -> str:
    """Return ``source`` with one passage replaced; the fixture must hold it."""
    assert old in source, old
    return source.replace(old, new, 1)


def after(anchor: str, addition: str, source: str = BASE) -> str:
    return swap(anchor, anchor + addition, source)


def continued(count: int) -> str:
    """A module with one statement of ``count`` line continuations."""
    pieces = ' _\n    & '.join(['"a"'] * (count + 1))
    return f'Private Sub FX_Long()\n    Dim Msg As String\n    Msg = {pieces}\nEnd Sub\n'


def sized(total: int) -> str:
    """A module of one procedure, ``total`` bytes of source from header to End."""
    head, tail = 'Private Sub FX_Big()\r\n', 'End Sub\r\n'
    pad = total - len(head) - len(tail)
    return f"{head}'{'x' * (pad - 3)}\r\n{tail}"


RULES = frozenset({'parentheses', 'lexical', 'blocks', 'continuations', 'procedure-size', 'longptr-assignment',
                   'byref-type', 'unreadable'})
Sources = Mapping[str, str]
# (name, sources, rule, message fragment, compilations that report it). A case
# with no rule must report nothing.
CASES: list[tuple[str, Sources, str, str, tuple[str, ...]]] = [
    ('the compliant module passes', {'a.bas': BASE}, '', '', ()),
    # parentheses
    ('a closing parenthesis nothing opened is rejected (FindingSummary)',
     {'a.bas': swap('    FX_Fill Style, 1\n', '    FX_Fill Style, _\n        FindingSummary)\n')},
     'parentheses', 'closing parenthesis has no opening', ALL),
    ('an opening parenthesis never closed is rejected',
     {'a.bas': swap('    Owned = Total + 1\n', '    Owned = (Total + 1\n')},
     'parentheses', '1 opening parenthesis never closed', ALL),
    ('a parenthesis dropped from a continued call is rejected',
     {'a.bas': swap('"b")', '"b"')},
     'parentheses', '1 opening parenthesis never closed', ALL),
    ('parentheses in a string, a comment and a Rem pass',
     {'a.bas': after('    Text = FX_NOTE\n', '    Text = ")))(" & Chr$(40) \' ((( \n    Rem )))\n')},
     '', '', ()),
    ('a string that never closes is rejected',
     {'a.bas': swap('    Text = FX_NOTE\n', '    Text = "never closes\n')},
     'lexical', 'never closes', ALL),
    ('a parenthesis in a branch the compilation does not read is judged in its own',
     {'a.bas': after('    Text = FX_NOTE\n', '#If VBA7 Then\n    Owned = (1\n#Else\n    Owned = (1)\n#End If\n')},
     'parentheses', 'never closed', ('vba7-x64', 'vba7-x86')),
    # blocks
    ('an If without End If is rejected',
     {'a.bas': swap('        Msg = "many"\n    End If\n', '        Msg = "many"\n')},
     'blocks', 'If opened at line', ALL),
    ('an End If without If is rejected',
     {'a.bas': swap('    FX_Run = (Owned = 0)\n', '    End If\n    FX_Run = (Owned = 0)\n')},
     'blocks', 'End If has no open If', ALL),
    ('a Loop that closes a For is rejected',
     {'a.bas': swap('    Next Index\n', '    Loop\n')},
     'blocks', 'Loop does not match the innermost open block, For', ALL),
    ('a Select Case never closed is rejected',
     {'a.bas': swap('        End Select\n', '')},
     'blocks', 'Next does not match the innermost open block, Select Case', ALL),
    ('a With never closed is rejected',
     {'a.bas': swap('    End With\n', '')},
     'blocks', 'With opened at line', ALL),
    ('a Do never closed is rejected',
     {'a.bas': swap('    Loop\n', '')},
     'blocks', 'Do opened at line', ALL),
    ('a Next without For is rejected',
     {'a.bas': swap('    FX_Run = (Owned = 0)\n', '    Next Index\n    FX_Run = (Owned = 0)\n')},
     'blocks', 'Next has no open For', ALL),
    ('a While without Wend is rejected',
     {'a.bas': swap('    FX_Run = (Owned = 0)\n', '    While Total > 0\n        Total = Total - 1\n    FX_Run = (Owned = 0)\n')},
     'blocks', 'While opened at line', ALL),
    ('a Case outside Select Case is rejected',
     {'a.bas': swap('    FX_Run = (Owned = 0)\n', '    Case 5\n    FX_Run = (Owned = 0)\n')},
     'blocks', 'Case has no open Select Case', ALL),
    ('an Else outside If is rejected',
     {'a.bas': swap('    FX_Run = (Owned = 0)\n', '    Else\n    FX_Run = (Owned = 0)\n')},
     'blocks', 'Else has no open If', ALL),
    ('an arm of a conditional block that drifts is rejected in that compilation only',
     {'a.bas': swap('    FX_Run = (Owned = 0)\n',
                    '#If VBA7 Then\n    If hWnd = 1 Then\n        Msg = "a"\n    End If\n#Else\n'
                    '    If hWnd = 1 Then\n        Msg = "a"\n#End If\n    FX_Run = (Owned = 0)\n')},
     'blocks', 'If opened at line', LEGACY),
    ('blocks that differ per arm and close in each pass',
     {'a.bas': swap('    FX_Run = (Owned = 0)\n',
                    '#If VBA7 Then\n    If hWnd = 1 Then\n        Msg = "a"\n    End If\n#Else\n'
                    '    For Index = 1 To 2\n        Msg = "a"\n    Next Index\n#End If\n'
                    '    FX_Run = (Owned = 0)\n')},
     '', '', ()),
    ('a Next that closes two loops passes',
     {'a.bas': swap('    FX_Run = (Owned = 0)\n',
                    '    For Index = 1 To 2\n        For Total = 1 To 2\n            Msg = "a"\n'
                    '    Next Total, Index\n    FX_Run = (Owned = 0)\n')},
     '', '', ()),
    ('an Else with a colon, an inline If and a loop on one line pass',
     {'a.bas': swap('    FX_Run = (Owned = 0)\n',
                    '    If hWnd = 1 Then\n        Msg = "a"\n    Else: Msg = "b"\n    End If\n'
                    '    For Index = 1 To 2: Msg = "a": Next Index\n    If hWnd = 3 Then Msg = "c": Msg = "d"\n'
                    '    FX_Run = (Owned = 0)\n')},
     '', '', ()),
    # continuations
    (f'a statement of {CONTINUATION_LIMIT} continuations passes',
     {'a.bas': continued(CONTINUATION_LIMIT)}, '', '', ()),
    (f'a statement of {CONTINUATION_LIMIT + 1} continuations is rejected',
     {'a.bas': continued(CONTINUATION_LIMIT + 1)},
     'continuations', f'{CONTINUATION_LIMIT + 1} line continuations', ALL),
    # procedure size
    ('a procedure one byte under the limit passes',
     {'a.bas': sized(PROCEDURE_LIMIT - 1)}, '', '', ()),
    ('a procedure of the limit is rejected',
     {'a.bas': sized(PROCEDURE_LIMIT)}, 'procedure-size', f'{PROCEDURE_LIMIT} bytes of source', ALL),
    # LongPtr assigned to Long
    ('a LongPtr expression assigned to a Long is rejected (EntryStyle And mask)',
     {'a.bas': swap('    Owned = Total + 1\n', STYLE_DEFECT)},
     'longptr-assignment', 'Owned is declared Long but the expression reads Style', X64),
    ('the same assignment through an inline If is rejected',
     {'a.bas': swap('    Owned = Total + 1\n', '    If hWnd = 1 Then Owned = Style Or 1\n')},
     'longptr-assignment', 'Owned is declared Long', X64),
    ('a grouped LongPtr expression assigned to a Long is rejected',
     {'a.bas': swap('    Owned = Total + 1\n', '    Owned = (Style And FX_MASK) + 1\n')},
     'longptr-assignment', 'Owned is declared Long', X64),
    ('a LongPtr assigned to a Long function result is rejected',
     {'a.bas': BASE + 'Public Function FX_Style(ByVal hWnd As LongPtr) As Long\n    FX_Style = hWnd\nEnd Function\n'},
     'longptr-assignment', 'FX_Style is declared Long', X64),
    ('a mask constant declared in another module is resolved',
     {'a.bas': swap('    Owned = Total + 1\n', '    Owned = Style And OT_MASK\n'), 'b.bas': SECOND_MODULE},
     'longptr-assignment', 'Owned is declared Long', X64),
    ('a LongPtr assigned to a LongPtr, a Variant and a Boolean passes',
     {'a.bas': swap('    Owned = Total + 1\n',
                    '    Style = Style And FX_MASK\n    Anything = Style Or 1\n    Flag = Style <> 0\n')},
     '', '', ()),
    ('an unknown name, a call and a string leave the assignment alone',
     {'a.bas': swap('    Owned = Total + 1\n',
                    '    Owned = Style And Unknown\n    Owned = Style + FX_Join("a", "b")\n'
                    '    Owned = Style & "a"\n')},
     '', '', ()),
    # ByRef type
    ('a Long passed to a ByRef LongPtr is rejected',
     {'a.bas': swap('    FX_Count Total\n', '    FX_Fill Total, 1\n')},
     'byref-type', 'Total is declared Long but FX_Fill takes Style ByRef As LongPtr', X64),
    ('a Long passed to a ByRef LongPtr through Call is rejected',
     {'a.bas': swap('    FX_Count Total\n', '    Call FX_Fill(Total, 1)\n')},
     'byref-type', 'Total is declared Long', X64),
    ('a Long passed to a ByRef LongPtr by name is rejected',
     {'a.bas': swap('    FX_Count Total\n', '    FX_Fill Flags:=1, Style:=Total\n')},
     'byref-type', 'Total is declared Long', X64),
    ('a Long passed to a ByRef LongPtr of a Declare is rejected',
     {'a.bas': swap('    FX_Count Total\n', '    Index = FX_Read(hWnd, Total)\n')},
     'byref-type', 'FX_Read takes Style ByRef As LongPtr', X64),
    ('an Integer passed to a ByRef Long is rejected in every compilation',
     {'a.bas': swap('    Dim Flag                As Boolean\n', '    Dim Flag                As Boolean\n    Dim Small As Integer\n')
      .replace('    FX_Count Total\n', '    FX_Count Small\n', 1)},
     'byref-type', 'Small is declared Integer but FX_Count takes Total ByRef As Long', ALL),
    ('a Boolean passed to a ByRef Long is rejected',
     {'a.bas': swap('    FX_Count Total\n', '    FX_Count Flag\n')},
     'byref-type', 'Flag is declared Boolean', ALL),
    ('a public procedure of another module of the project is resolved',
     {'a.bas': swap('    FX_Count Total\n', '    OT_Narrow Total\n'), 'b.bas': SECOND_MODULE},
     'byref-type', 'OT_Narrow takes Small ByRef As Integer', ALL),
    ('a private procedure of another module is not visible',
     {'a.bas': swap('    FX_Count Total\n', '    OT_Hidden Total\n'), 'b.bas': SECOND_MODULE},
     '', '', ()),
    ('a procedure that two modules declare is ambiguous and passes',
     {'a.bas': swap('    FX_Count Total\n', '    OT_Narrow Total\n'), 'b.bas': SECOND_MODULE,
      'c.bas': SECOND_MODULE.replace('M_OTHER', 'M_THIRD')},
     '', '', ()),
    ('an array element, a member, a constant and a parenthesised argument pass',
     {'a.bas': swap('    FX_Count Total\n', '    FX_Fill FX_MASK, 1\n    FX_Fill (Total), 1\n    FX_Fill Pair.Left, 1\n')},
     '', '', ()),
    ('a statement that calls with parentheses passes by value and is skipped',
     {'a.bas': swap('    FX_Count Total\n', '    FX_Fill(Total, 1)\n    If Flag Then FX_Fill(Total, 1)\n')},
     '', '', ()),
    # conditional compilation that cannot be read
    ('an unknown conditional symbol is reported, not skipped',
     {'a.bas': swap('    Text = FX_NOTE\n', '#If UnknownHost Then\n    Text = ""\n#End If\n    Text = FX_NOTE\n')},
     'unreadable', 'unsupported conditional symbol', ALL),
]


def violations(name: str, sources: Sources, rule: str, fragment: str, configs: Sequence[str]) -> list[str]:
    """Why the findings of one case are not the ones the case names."""
    found, _ = findings(sources)
    if not rule:
        return [f'{name}: unexpected {[str(f) for f in found]}'] if found else []
    named = [f for f in found if f.rule == rule and fragment in f.message]
    reported = {config for f in named for config in f.configs}
    others = [str(f) for f in found if f not in named]
    if reported != set(configs) or others:
        return [f'{name}: expected {rule!r} ({fragment!r}) in {tuple(configs)}; '
                f'got {[str(f) for f in found]}']
    return []


def selftest() -> tuple[list[str], int]:
    """Every fixture must report exactly the rule and compilations it names, and every rule has one."""
    errors: list[str] = [f'#203 fixture: no failing fixture for the {rule} rule'
                         for rule in sorted(RULES - {case[2] for case in CASES})]
    for name, sources, rule, fragment, configs in CASES:
        errors.extend(f'#203 fixture {message}' for message in violations(name, sources, rule, fragment, configs))
    return errors, len(CASES)
