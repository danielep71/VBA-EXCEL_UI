"""Versioned, independent defect/benign fixtures for the bounded #31 analyzer.

Expectations specify exact rule, source path and physical line (configuration
multiplicity is deliberately deduplicated). No fixture is allowed extra rules.
"""
from vba_analyze import analyze


def selftest():
    failures = []
    count = 0

    def case(name, source, expected=(), dynamic=0):
        nonlocal count
        count += 1
        sources = source if isinstance(source, dict) else {'fixture.bas': source}
        found, inventory = analyze(sources)
        actual = {(f.file, f.line, f.rule) for f in found}
        wanted = {(p, n, r) for p, n, r in expected}
        if actual != wanted or len(inventory) != dynamic:
            failures.append(f'{name}: expected {sorted(wanted)}, dynamic={dynamic}; '
                            f'got {sorted(actual)}, dynamic={len(inventory)}')

    def bad(name, source, line, rule):
        case(name, source, [('fixture.bas', line, rule)])

    bad('unresolved direct call', 'Sub Test()\nUI_Missing\nEnd Sub', 2, 'unresolved')
    bad('unresolved expression', 'Function Test()\nTest = UI_Missing()\nEnd Function', 2, 'unresolved')
    bad('continued call physical line', 'Sub Test()\nCall _\n UI_Missing()\nEnd Sub', 2, 'unresolved')
    bad('colon call', 'Sub Test(): UI_Missing: End Sub', 1, 'unresolved')
    bad('43 missing CreatedSnapshot Dim', 'Option Explicit\nSub Test()\nCreatedSnapshot = True\nEnd Sub', 3, 'undeclared-assignment')
    case('43 declared counterpart', 'Option Explicit\nSub Test()\nDim CreatedSnapshot As Boolean\nCreatedSnapshot = True\nEnd Sub')
    bad('Let assignment', 'Option Explicit\nSub Test()\nLet missing = 1\nEnd Sub', 3, 'undeclared-assignment')
    bad('Set assignment', 'Option Explicit\nSub Test()\nSet missing = Nothing\nEnd Sub', 3, 'undeclared-assignment')
    bad('continued assignment', 'Option Explicit\nSub Test()\nLet missing _\n = True\nEnd Sub', 3, 'undeclared-assignment')
    bad('inline Then assignment', 'Option Explicit\nSub Test()\nIf True Then missing = 1\nEnd Sub', 3, 'undeclared-assignment')
    bad('inline Else assignment', 'Option Explicit\nSub Test()\nIf True Then Debug.Print 1 Else missing = 1\nEnd Sub', 3, 'undeclared-assignment')
    bad('cross procedure label', 'Sub A()\nGoTo Safe_Exit\nEnd Sub\nSub B()\nSafe_Exit:\nEnd Sub', 2, 'label')
    bad('mismatched End', 'Function Test()\nEnd Sub', 2, 'structure')
    bad('mismatched property End', 'Property Get Value()\nEnd Function', 2, 'structure')
    bad('duplicate procedure', 'Sub Test()\nEnd Sub\nSub Test()\nEnd Sub', 3, 'duplicate')
    bad('Sub Function name collision', 'Sub Test()\nEnd Sub\nFunction Test()\nEnd Function', 3, 'duplicate')
    bad('nested PtrSafe after inner End', '#If VBA7 Then\n#If Win64 Then\n#Else\n#End If\nPrivate Declare Function Native Lib "custom" () As Long\n#End If', 5, 'ptrsafe')
    bad('unconditional VBA7 PtrSafe', 'Private Declare Function Native Lib "custom" () As Long', 1, 'ptrsafe')
    bad('missing prefixed Alias', 'Private Declare PtrSafe Function UI_Native Lib "custom" () As Long', 1, 'alias')
    case('explicit Alias', 'Private Declare PtrSafe Function UI_Native Lib "custom" Alias "Native" () As Long')
    bad('unmatched Else', '#Else\n', 1, 'conditional/lexical')
    bad('double Else', '#If VBA7 Then\n#Else\n#Else\n#End If', 3, 'conditional/lexical')
    bad('ElseIf after Else', '#If VBA7 Then\n#Else\n#ElseIf Win64 Then\n#End If', 3, 'conditional/lexical')
    bad('unknown condition fails explicitly', '#If UnknownHost Then\n#End If', 1, 'conditional/lexical')
    case('alternatives and nested ElseIf', '#If VBA7 Then\n#If Win64 Then\nSub Test()\n#ElseIf VBA7 Then\nSub Test()\n#Else\nSub Test()\n#End If\n#Else\nSub Test()\n#End If\nEnd Sub')
    case('boolean expressions and local Const', '#Const LocalFlag = True\n#If LocalFlag And (VBA7 Or Not VBA7) Then\nSub Test()\nEnd Sub\n#End If')
    case('benign literals comments members named arguments', '''Option Explicit
Sub Test(ByVal target As Object)
Dim text As String
text = "UI_Missing: GoTo Missing" ' UI_Absent
Rem UI_Absent: GoTo Missing
target.UI_Member = 1
Call target.TST_Member(UI_Named:=True)
If True Then Rem UI_Absent
End Sub''')
    case('doubled quotes and colon Rem', '''Sub Test()
Debug.Print "say ""UI_Absent"": still string": Rem TST_Absent
End Sub''')
    case('bracketed identifier member', 'Sub Test()\ntarget.[UI_Missing] = True\nEnd Sub')
    case('visible definitions across module categories', {
        'src/defs.bas': '''Public Const UI_Constant As Long = 1
Public UI_Global As Long
Public Type UI_Record
Value As Long
End Type
Public Enum UI_Kind
UI_First = 0
End Enum
Public Declare PtrSafe Function UI_Native Lib "custom" Alias "Native" () As Long
Public Function UI_Function() As Long
UI_Function = UI_First
End Function
Public Property Get UI_Value() As Long
UI_Value = UI_Constant
End Property
Public Property Let UI_Value(ByVal rhs As Long)
UI_Global = rhs
End Property''',
        'test/test.bas': '''Option Explicit
Sub TST_Test(ByVal parameter As Long)
Dim item As UI_Record, local As Long
local = UI_First + UI_Constant + UI_Function() + UI_Native()
UI_Global = parameter
End Sub''',
        'demo/demo.bas': 'Sub Demo_Test()\nTST_Test UI_Value\nEnd Sub'})
    case('procedure local declarations and return', '''Option Explicit
Private mValue As Long
Function Test(ByVal rhs As Long) As Long
Static held As Long
Const cap As Long = 10
Dim first As Long, second As Long
first = rhs: Let second = cap: held = first
mValue = second
Test = held
End Function''')
    case('private procedure names may repeat', {'a.bas': 'Private Sub TST_Local()\nEnd Sub', 'b.bas': 'Private Sub TST_Local()\nEnd Sub'})
    case('private names are not exported', {'a.bas': 'Private Sub TST_Local()\nEnd Sub', 'b.bas': 'Sub Test()\nTST_Local\nEnd Sub'}, [('b.bas', 2, 'unresolved')])
    case('local labels and same-line jump', 'Sub Test(): GoTo Safe_Exit\nSafe_Exit: Resume Next\nEnd Sub')
    case('dynamic strings inventoried', '''Sub Test()
Application.Run "UI_Missing"
shape.OnAction = "TST_Missing"
CallByName target, "Demo_Missing", 1
End Sub''', dynamic=3)

    # Historical #34: argument Err fields evaluate before TST_Log returns.
    log = '\nSub TST_Log(ByVal p As String, ByVal status As String, ByVal detail As String)\nEnd Sub'
    bad('34 read after logging', '''Sub Test()
On Error GoTo Err_Handler
Exit Sub
Err_Handler:
m_CertActive = False
TST_Log "PROC", "FAIL", Err.Description
Err.Raise Err.Number, Err.Source, Err.Description
End Sub''' + log, 7, 'stale-err')
    case('34 saved fields before logging', '''Option Explicit
Sub Test()
Dim savedNumber As Long, savedSource As String, savedDescription As String
On Error GoTo Err_Handler
Exit Sub
Err_Handler:
savedNumber = Err.Number
savedSource = Err.Source
savedDescription = Err.Description
TST_Log "PROC", "FAIL", savedDescription
Err.Raise savedNumber, savedSource, savedDescription
End Sub''' + log)
    bad('39 On Error before capture', '''Function Demo_GetRuntimeErrorText() As String
On Error Resume Next
Demo_GetRuntimeErrorText = CStr(Err.Number) & ": " & Err.Description
End Function''', 3, 'stale-err')
    case('39 captured before On Error', '''Function Demo_GetRuntimeErrorText() As String
Dim n As Long, description As String
n = Err.Number
description = Err.Description
On Error Resume Next
Demo_GetRuntimeErrorText = CStr(n) & ": " & description
End Function''')
    case('Err as call arguments only', '''Sub Test()
On Error GoTo Err_Handler
Exit Sub
Err_Handler:
TST_Log "PROC", "FAIL", Err.Description
End Sub''' + log)
    case('fresh failing host operation', '''Sub Test()
Dim n As Long
On Error Resume Next
Application.ExecuteExcel4Macro "GET.TOOLBAR(\"\"Ribbon\"\")"
n = Err.Number
End Sub''')
    case('new concatenation allocation', '''Sub Test()
Dim entry As String, detail As String, n As Long
On Error Resume Next
entry = "stage | " & detail
n = Err.Number
End Sub''')
    case('pure capture expression', '''Sub Test()
Dim n As Long, s As String, d As String
On Error GoTo Err_Handler
Exit Sub
Err_Handler:
n = Err.Number
s = IIf(Len(Err.Source) > 0, Err.Source, "PROC")
d = Err.Description
End Sub''')
    bad('direct On Error handler clobber', '''Sub Test()
On Error GoTo Err_Handler
Exit Sub
Err_Handler:
On Error GoTo 0
Debug.Print Err.Description
End Sub''', 6, 'stale-err')
    bad('call result clobbers later capture', '''Sub Test()
On Error GoTo Err_Handler
Exit Sub
Err_Handler:
value = target.Read()
value = Err.Number
End Sub''', 6, 'stale-err')

    # ABI v1: explicit independent native declarations. Mutations below alter
    # pointer/return types; expectations never import the analyzer's ABI table.
    native_v1 = [
        'Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal h As LongPtr, ByVal i As Long) As LongPtr',
        'Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal h As LongPtr, ByVal i As Long, ByVal v As LongPtr) As LongPtr',
        'Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal h As LongPtr, ByVal i As Long) As Long',
        'Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal h As LongPtr, ByVal i As Long, ByVal v As Long) As Long',
        'Function SetWindowPos Lib "user32" (ByVal h As LongPtr, ByVal after As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long',
        'Function IsWindow Lib "user32" (ByVal h As LongPtr) As Long',
        'Function GetLastError Lib "kernel32" () As Long',
        'Sub SetLastError Lib "kernel32" (ByVal code As Long)',
    ]
    for i, declaration in enumerate(native_v1):
        source = '#If VBA7 Then\nPrivate Declare PtrSafe ' + declaration + '\n#End If'
        case(f'ABI v1 {i} good', source)
        # First parameter (or return for zero-argument functions).
        mutated = source.replace('As LongPtr', 'As Integer', 1) if 'LongPtr' in source else source.replace('As Long', 'As Integer', 1)
        bad(f'ABI v1 {i} bad', mutated, 2, 'winapi-signature')
        legacy = '#If Not VBA7 Then\nPrivate Declare ' + declaration.replace('LongPtr', 'Long') + '\n#End If'
        case(f'ABI v1 {i} legacy', legacy)

    bad('duplicate local label', 'Sub Test()\nSafe_Exit:\nSafe_Exit:\nEnd Sub', 3, 'label')
    case('project prefix label is not a call', 'Sub Test()\nGoTo UI_Done\nUI_Done:\nEnd Sub')
    bad('ABI ByRef mismatch', '#If VBA7 Then\nPrivate Declare PtrSafe Function IsWindow Lib "user32" (h As LongPtr) As Long\n#End If', 2, 'winapi-signature')
    bad('ABI result mismatch', '#If VBA7 Then\nPrivate Declare PtrSafe Function IsWindow Lib "user32" (ByVal h As LongPtr) As LongPtr\n#End If', 2, 'winapi-signature')
    bad('scalar flag does not refresh cleared Err', 'Sub Test()\nOn Error Resume Next\nflag = False\nn = Err.Number\nEnd Sub', 4, 'stale-err')
    case('numeric literal is not a member call', 'Sub Test()\nOn Error GoTo Err_Handler\nExit Sub\nErr_Handler:\nflag = 1.5\nn = Err.Number\nEnd Sub')
    return failures, count


if __name__ == '__main__':
    findings, count = selftest()
    print(f'{count} VBA analyzer fixtures; {len(findings)} failures')
    for finding in findings:
        print(finding)
    raise SystemExit(bool(findings))
