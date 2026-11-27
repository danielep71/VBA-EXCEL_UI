"""Fixtures for the equivalence proof of house_equiv.py (#210).

Every case changes BASE in one way and says whether the proof must accept the
result. The gate runs equiv_selftest() on every push, whatever the state of
house_style.HOUSE_STYLE_ACTIVE.
"""
import house_equiv

BASE = (
    'Attribute VB_Name = "M_PROBE"\r\n'
    "Option Explicit\r\n"
    "\r\n"
    "Private Const PROBE_LIMIT As Long = 10\r\n"
    "\r\n"
    "Public Function Probe_Describe( _\r\n"
    "    ByVal Value As Long) _\r\n"
    "    As String\r\n"
    "\r\n"
    "    Dim Text As String\r\n"
    '    Const PROC As String = "M_PROBE.Probe_Describe"\r\n'
    "    On Error GoTo Err_Handler\r\n"
    "    If Value > PROBE_LIMIT Then\r\n"
    '        Err.Raise vbObjectError + 1, PROC, "Value above the limit"\r\n'
    "    End If\r\n"
    "    Select Case Value\r\n"
    "        Case 0\r\n"
    '            Text = "zero"\r\n'
    "        Case 1\r\n"
    '            Text = "one"\r\n'
    "    End Select\r\n"
    "    Probe_Describe = Text\r\n"
    "Safe_Exit:\r\n"
    "    Exit Function\r\n"
    "Err_Handler:\r\n"
    "    Err.Raise Err.Number, PROC, Err.Description\r\n"
    "End Function\r\n"
    "\r\n"
    "Private Sub Probe_Reset()\r\n"
    "    mState = 0\r\n"
    "End Sub\r\n"
)

# The same module as the style leaves it: the renames, Const PROC_NAME first, an empty Case Else, other layout.
STYLED = (
    'Attribute VB_Name = "M_PROBE"\r\n'
    "\r\n"
    "Option Explicit\r\n"
    "\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' PRIVATE CONSTANTS\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "    Private Const PROBE_LIMIT As Long = 10 'Largest accepted value\r\n"
    "\r\n"
    "Public Function Probe_Describe( _\r\n"
    "    ByVal Value As Long) _\r\n"
    "    As String\r\n"
    "'\r\n"
    "' PURPOSE\r\n"
    "'   Describes one value.\r\n"
    "\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' DECLARE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    '    Const PROC_NAME As String = "M_PROBE.Probe_Describe"\r\n'
    "\r\n"
    "    Dim Text As String 'Result text\r\n"
    "\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' INITIALIZE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "    'Route unexpected errors to the handler\r\n"
    "        On Error GoTo ErrorHandler\r\n"
    "\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' PICK TEXT\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "    'Reject a value above the limit\r\n"
    "        If Value > PROBE_LIMIT Then\r\n"
    "            Err.Raise vbObjectError + 1, _\r\n"
    '                PROC_NAME, "Value above the limit"\r\n'
    "        End If\r\n"
    "    'Pick the text\r\n"
    "        Select Case Value\r\n"
    "            Case 0\r\n"
    '                Text = "zero"\r\n'
    "            Case 1\r\n"
    '                Text = "one"\r\n'
    "            Case Else\r\n"
    "            'Leave the text empty for any other value\r\n"
    "        End Select\r\n"
    "        Probe_Describe = Text\r\n"
    "\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' EXIT PROCEDURE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "SafeExit:\r\n"
    "        Exit Function\r\n"
    "\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' ERROR HANDLER\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "ErrorHandler:\r\n"
    "        Err.Raise Err.Number, PROC_NAME, Err.Description\r\n"
    "\r\n"
    "End Function\r\n"
    "\r\n"
    "Private Sub Probe_Reset()\r\n"
    "        mState = 0\r\n"
    "End Sub\r\n"
)

# (name, text before, text after, part of the message that must appear; "" when the proof must accept)
# The edits start from STYLED, which the proof accepts for BASE.
CASES: list[tuple[str, str, str, str]] = [
    ("comments, blank lines and indentation", "        Probe_Describe = Text\r\n",
     "\r\n   'A comment\r\n  Probe_Describe = Text   'Another\r\n\r\n", ""),
    ("a statement wrapped over lines", "            Err.Raise vbObjectError + 1, _\r\n                PROC_NAME, ",
     "            Err.Raise vbObjectError + 1, PROC_NAME, _\r\n                ", ""),
    ("Const PROC_NAME moved after the variables", '    Const PROC_NAME As String = "M_PROBE.Probe_Describe"\r\n\r\n    Dim Text As String \'Result text\r\n',
     '    Dim Text As String \'Result text\r\n    Const PROC_NAME As String = "M_PROBE.Probe_Describe"\r\n', ""),
    ("an empty Case Else removed again", "            Case Else\r\n            'Leave the text empty for any other value\r\n", "", ""),
    ("a new procedure order is a difference", "Private Sub Probe_Reset()\r\n        mState = 0\r\nEnd Sub\r\n", "",
     "procedure 2 is Probe_Reset in base and nothing in new"),
    ("a changed operator", "If Value > PROBE_LIMIT Then", "If Value < PROBE_LIMIT Then", "base has `If Value > PROBE_LIMIT Then`"),
    ("a deleted statement", "        Probe_Describe = Text\r\n", "", "a statement was deleted"),
    ("an inserted statement", "        Probe_Describe = Text\r\n", "        Probe_Describe = Text\r\n        mState = 1\r\n",
     "a statement was inserted"),
    ("a changed literal", '"Value above the limit"', '"Value over the limit"', "Value over the limit"),
    ("a changed number", "vbObjectError + 1", "vbObjectError + 2", "vbObjectError"),
    ("a label renamed outside the declared names", "SafeExit:", "LeaveNow:", "LeaveNow"),
    ("a statement in the new Case Else", "            'Leave the text empty for any other value\r\n",
     '            Text = "other"\r\n', "Case Else"),
    ("a changed PROC_NAME literal", '"M_PROBE.Probe_Describe"', '"M_PROBE.Other"', "the literal changed"),
    ("a deleted PROC_NAME", '    Const PROC_NAME As String = "M_PROBE.Probe_Describe"\r\n\r\n', "", "new does not"),
    ("PROC_NAME declared twice", '    Const PROC_NAME As String = "M_PROBE.Probe_Describe"\r\n',
     '    Const PROC_NAME As String = "M_PROBE.Probe_Describe"\r\n    Const PROC_NAME As String = "M_PROBE.Probe_Describe"\r\n',
     "declared 2 times"),
    ("a statement removed from a short procedure", "        mState = 0\r\n", "", "a statement was deleted"),
    ("a changed Option statement", "Option Explicit\r\n", "Option Explicit\r\nOption Private Module\r\n", "Option statements differ"),
    ("a module-level declaration changed", "Private Const PROBE_LIMIT As Long = 10", "Private Const PROBE_LIMIT As Long = 11",
     "outside the procedures"),
]

# A base without any PROC constant, to which one is added.
BARE = BASE.replace('    Const PROC As String = "M_PROBE.Probe_Describe"\r\n', "").replace("PROC,", "Probe_Describe,")
BARE = BARE.replace("Err_Handler", "ErrorHandler").replace("Safe_Exit", "SafeExit")
ADDED = '    Const PROC_NAME As String = "{}"\r\n    Dim Text As String\r\n'

# (name, text of the new tree, part of the message that must appear; "" when the proof must accept)
ADDITIONS: list[tuple[str, str, str]] = [
    ("a PROC_NAME added with the procedure's own name",
     BARE.replace("    Dim Text As String\r\n", ADDED.format("M_PROBE.Probe_Describe")), ""),
    ("a PROC_NAME added with another procedure's name",
     BARE.replace("    Dim Text As String\r\n", ADDED.format("M_PROBE.Other")), "an added PROC_NAME must be"),
    ("a PROC_NAME added without the module", BARE.replace("    Dim Text As String\r\n", ADDED.format("Probe_Describe")),
     "an added PROC_NAME must be"),
]


def compare(base: str, new: str) -> list[str]:
    return house_equiv.compare_modules(base, new, "M_PROBE")


def equiv_selftest() -> list[str]:
    """Failure descriptions; empty means the proof accepts what it should and refuses the rest."""
    failures: list[str] = []
    if compare(BASE, STYLED):
        failures.append(f"the styled fixture must be equivalent to its base, got {compare(BASE, STYLED)!r}")
    if compare(STYLED, STYLED.replace("\r\n", "\n")):
        failures.append("a line-ending change must not be a difference")
    if not compare(STYLED, BASE):
        failures.append("the renames are declared for the base only: the new tree may not keep the previous names")
    for name, before, after, part in CASES:
        if before not in STYLED:
            failures.append(f"{name}: fixture text not found")
            continue
        found = compare(BASE, STYLED.replace(before, after, 1))
        if part == "" and found:
            failures.append(f"{name}: refused, but only an allowed difference was made: {found!r}")
        elif part and not any(part in item for item in found):
            failures.append(f"{name}: expected a difference naming {part!r}, got {found!r}")
    for name, new, part in ADDITIONS:
        found = compare(BARE, new)
        if part == "" and found:
            failures.append(f"{name}: refused, but only an allowed difference was made: {found!r}")
        elif part and not any(part in item for item in found):
            failures.append(f"{name}: expected a difference naming {part!r}, got {found!r}")
    return failures
