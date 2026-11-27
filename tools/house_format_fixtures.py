"""Independent expectations for the house-style normal form (#210).

Each fixture states the text the formatter must produce, so a failure names
the rule and not a byte count. The gate runs house_selftest() on every push,
whatever the state of house_style.HOUSE_STYLE_ACTIVE.

LEGACY_MODULE is a small module in the previous layout; HOUSE_MODULE is the
text the formatter must give for it. HOUSE_MODULE has no deviation from any
rule of the style, which house_rules_fixtures relies on.
"""
import contextlib
import io
import os
import tempfile
from pathlib import Path
from unittest.mock import patch

import house_style
import vba_house
import vba_house_rules
from vba_lex import LexicalError

LEGACY_MODULE = (
    'Attribute VB_Name = "M_PROBE"\r\n'
    'Option Explicit\r\n'
    'Option Private Module\r\n'
    '\r\n'
    "'==============================================================================\r\n"
    "' M_PROBE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' PURPOSE\r\n"
    "'   Holds the probe procedures of the style fixtures.\r\n"
    "'\r\n"
    "' WHY THIS EXISTS\r\n"
    "'   The formatter needs one module that shows every rule.\r\n"
    "'\r\n"
    "' INPUTS\r\n"
    "'   None at module level\r\n"
    "'\r\n"
    "' RETURNS\r\n"
    "'   Nothing at module level\r\n"
    "'\r\n"
    "' BEHAVIOR\r\n"
    "'   Offers a function, a handler and a compact helper.\r\n"
    "'\r\n"
    "' ERROR POLICY\r\n"
    "'   Raises descriptive errors.\r\n"
    "'\r\n"
    "' DEPENDENCIES\r\n"
    "'   VBA library only\r\n"
    "'\r\n"
    "' NOTES\r\n"
    "'   The fixtures read this header.\r\n"
    "'\r\n"
    "' UPDATED\r\n"
    "'   2026-11-02 - Added the helper (#210).\r\n"
    "'   2026-11-01 - First version.\r\n"
    "'\r\n"
    "' AUTHOR\r\n"
    "'   Probe\r\n"
    "'==============================================================================\r\n"
    '\r\n'
    "'==============================================================================\r\n"
    "' PRIVATE CONSTANTS\r\n"
    "'==============================================================================\r\n"
    '\r\n'
    "Private Const PROBE_LIMIT As Long = 10 'Largest accepted value\r\n"
    'Private Const PROBE_NAME As String = "x" \'Probe label\r\n'
    "Private mState As Long 'Probe state\r\n"
    '\r\n'
    'Public Function Probe_Describe( _\r\n'
    '    ByVal Value As Long) _\r\n'
    '    As String\r\n'
    "'\r\n"
    "'==============================================================================\r\n"
    "' Probe_Describe\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' PURPOSE\r\n"
    "'   Describes one value.\r\n"
    "'\r\n"
    "' WHY THIS EXISTS\r\n"
    "'   A caller needs one text per value.\r\n"
    "'\r\n"
    "' INPUTS\r\n"
    "'   Value\r\n"
    "'     The value to describe.\r\n"
    "'\r\n"
    "' RETURNS\r\n"
    "'   The description.\r\n"
    "'\r\n"
    "' BEHAVIOR\r\n"
    "'   - Picks the text.\r\n"
    "'\r\n"
    "' ERROR POLICY\r\n"
    "'   Raises on a value above PROBE_LIMIT.\r\n"
    "'\r\n"
    "' DEPENDENCIES\r\n"
    "'   None.\r\n"
    "'\r\n"
    "' CALLED FROM\r\n"
    "'   - M_PROBE\r\n"
    "'\r\n"
    "' NOTES\r\n"
    '\'   A text such as "GoTo Fail" is data.\r\n'
    "'\r\n"
    "' UPDATED\r\n"
    "'   2026-11-02 - Added (#210).\r\n"
    "'   2026-11-01 - Created.\r\n"
    "'==============================================================================\r\n"
    "'\r\n"
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' DECLARE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    '    Const PROC As String = "M_PROBE.Probe_Describe"\r\n'
    "    Dim Text As String 'Result text\r\n"
    "    Dim Index As Long 'Loop counter\r\n"
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' INITIALIZE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "    'Enable the handler\r\n"
    '    On Error GoTo Err_Handler\r\n'
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' VALIDATE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    '        If Value > PROBE_LIMIT Then\r\n'
    '        Err.Raise vbObjectError + 1, PROC, "Value above the limit; do not GoTo Fail"\r\n'
    '        End If\r\n'
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' PICK TEXT\r\n"
    "'------------------------------------------------------------------------------\r\n"
    '        For Index = 1 To 2\r\n'
    '        Select Case Value\r\n'
    '        Case 0\r\n'
    '        Text = "zero"\r\n'
    '        Case Else\r\n'
    '        Text = "other"\r\n'
    '        End Select\r\n'
    '        Next Index\r\n'
    '\r\n'
    '        Probe_Describe = Text\r\n'
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' RETURN SUCCESS\r\n"
    "'------------------------------------------------------------------------------\r\n"
    'Safe_Exit:\r\n'
    '    Exit Function\r\n'
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' ERROR HANDLER\r\n"
    "'------------------------------------------------------------------------------\r\n"
    'Err_Handler:\r\n'
    '        Probe_Describe = vbNullString\r\n'
    '        Resume Safe_Exit\r\n'
    '\r\n'
    'End Function\r\n'
    '\r\n'
    '\r\n'
    'Private Sub Probe_Reset()\r\n'
    "'\r\n"
    "'==============================================================================\r\n"
    "' Probe_Reset\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' PURPOSE\r\n"
    "'   Clears the probe state.\r\n"
    "'\r\n"
    "' RETURNS\r\n"
    "'   None.\r\n"
    "'\r\n"
    "' UPDATED\r\n"
    "'   2026-11-01\r\n"
    "'==============================================================================\r\n"
    "'\r\n"
    '\r\n'
    '        mState = 0\r\n'
    '\r\n'
    'End Sub\r\n'
)

HOUSE_MODULE = (
    'Attribute VB_Name = "M_PROBE"\r\n'
    '\r\n'
    'Option Explicit\r\n'
    'Option Private Module\r\n'
    '\r\n'
    "'\r\n"
    "'==============================================================================\r\n"
    "' MODULE: M_PROBE\r\n"
    "'==============================================================================\r\n"
    "' PURPOSE\r\n"
    "'   Holds the probe procedures of the style fixtures.\r\n"
    "'\r\n"
    "' WHY THIS EXISTS\r\n"
    "'   The formatter needs one module that shows every rule.\r\n"
    "'\r\n"
    "' INPUTS\r\n"
    "'   None at module level\r\n"
    "'\r\n"
    "' RETURNS\r\n"
    "'   Nothing at module level\r\n"
    "'\r\n"
    "' BEHAVIOR\r\n"
    "'   Offers a function, a handler and a compact helper.\r\n"
    "'\r\n"
    "' ERROR POLICY\r\n"
    "'   Raises descriptive errors.\r\n"
    "'\r\n"
    "' DEPENDENCIES\r\n"
    "'   VBA library only\r\n"
    "'\r\n"
    "' NOTES\r\n"
    "'   The fixtures read this header.\r\n"
    "'\r\n"
    "' AUTHOR\r\n"
    "'   Probe\r\n"
    "'\r\n"
    "' HISTORY\r\n"
    "'   2026-11-02 - Added the helper (#210).\r\n"
    "'   2026-11-01 - First version.\r\n"
    "'\r\n"
    "' UPDATED\r\n"
    "'   2026-11-02\r\n"
    "'==============================================================================\r\n"
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' PRIVATE CONSTANTS\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "    Private Const PROBE_LIMIT   As Long = 10    'Largest accepted value\r\n"
    '    Private Const PROBE_NAME    As String = "x" \'Probe label\r\n'
    "    Private mState              As Long         'Probe state\r\n"
    '\r\n'
    '\r\n'
    'Public Function Probe_Describe( _\r\n'
    '    ByVal Value As Long) _\r\n'
    '    As String\r\n'
    "'\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "'                                PROBE DESCRIBE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' PURPOSE\r\n"
    "'   Describes one value.\r\n"
    "'\r\n"
    "' WHY THIS EXISTS\r\n"
    "'   A caller needs one text per value.\r\n"
    "'\r\n"
    "' INPUTS\r\n"
    "'   Value\r\n"
    "'     The value to describe.\r\n"
    "'\r\n"
    "' RETURNS\r\n"
    "'   The description.\r\n"
    "'\r\n"
    "' BEHAVIOR\r\n"
    "'   - Picks the text.\r\n"
    "'\r\n"
    "' ERROR POLICY\r\n"
    "'   Raises on a value above PROBE_LIMIT.\r\n"
    "'\r\n"
    "' DEPENDENCIES\r\n"
    "'   None.\r\n"
    "'\r\n"
    "' NOTES\r\n"
    '\'   A text such as "GoTo ErrorHandler" is data.\r\n'
    "'\r\n"
    "' CALLED FROM\r\n"
    "'   - M_PROBE\r\n"
    "'\r\n"
    "' HISTORY\r\n"
    "'   2026-11-02 - Added (#210).\r\n"
    "'   2026-11-01 - Created.\r\n"
    "'\r\n"
    "' UPDATED\r\n"
    "'   2026-11-02\r\n"
    "'------------------------------------------------------------------------------\r\n"
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' DECLARE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    '    Const PROC_NAME As String = "M_PROBE.Probe_Describe"\r\n'
    '\r\n'
    "    Dim Text    As String   'Result text\r\n"
    "    Dim Index   As Long     'Loop counter\r\n"
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' INITIALIZE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "    'Enable the handler\r\n"
    '        On Error GoTo ErrorHandler\r\n'
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' VALIDATE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    '        If Value > PROBE_LIMIT Then\r\n'
    '            Err.Raise vbObjectError + 1, PROC_NAME, "Value above the limit; do not GoTo Fail"\r\n'
    '        End If\r\n'
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' PICK TEXT\r\n"
    "'------------------------------------------------------------------------------\r\n"
    '        For Index = 1 To 2\r\n'
    '            Select Case Value\r\n'
    '                Case 0\r\n'
    '                    Text = "zero"\r\n'
    '                Case Else\r\n'
    '                    Text = "other"\r\n'
    '            End Select\r\n'
    '        Next Index\r\n'
    '\r\n'
    '        Probe_Describe = Text\r\n'
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' EXIT PROCEDURE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    'SafeExit:\r\n'
    '        Exit Function\r\n'
    '\r\n'
    "'------------------------------------------------------------------------------\r\n"
    "' ERROR HANDLER\r\n"
    "'------------------------------------------------------------------------------\r\n"
    'ErrorHandler:\r\n'
    '        Probe_Describe = vbNullString\r\n'
    '        Resume SafeExit\r\n'
    '\r\n'
    'End Function\r\n'
    '\r\n'
    'Private Sub Probe_Reset()\r\n'
    '\r\n'
    "'\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "'                                 PROBE RESET\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' PURPOSE\r\n"
    "'   Clears the probe state.\r\n"
    "'\r\n"
    "' RETURNS\r\n"
    "'   Nothing\r\n"
    "'\r\n"
    "' UPDATED\r\n"
    "'   2026-11-01\r\n"
    "'------------------------------------------------------------------------------\r\n"
    '\r\n'
    '        mState = 0\r\n'
    '\r\n'
    'End Sub\r\n'
)


MODULE_HEAD = (
    'Attribute VB_Name = "M_PROBE"\r\n'
    "Option Explicit\r\n"
    "\r\n"
    "'==============================================================================\r\n"
    "' M_PROBE\r\n"
    "'------------------------------------------------------------------------------\r\n"
    "' PURPOSE\r\n"
    "'   Probe.\r\n"
    "'\r\n"
    "' UPDATED\r\n"
    "'   2026-11-01\r\n"
    "'==============================================================================\r\n"
    "\r\n"
)

PROCEDURE_HEAD = [
    "Private Sub Probe()",
    "'",
    "'==============================================================================",
    "' Probe",
    "'------------------------------------------------------------------------------",
    "' PURPOSE",
    "'   Probe.",
    "'",
    "' UPDATED",
    "'   2026-11-01",
    "'==============================================================================",
    "'",
    "",
]

BANNER = "'" + "-" * 78


def banner(title: str) -> list[str]:
    return [BANNER, "' " + title, BANNER]


def module_text(body: list[str], head: list[str] | None = None) -> str:
    lines = (PROCEDURE_HEAD if head is None else head) + body + ["End Sub"]
    return MODULE_HEAD + "\r\n".join(lines) + "\r\n"


def format_body(body: list[str]) -> tuple[list[str], list[tuple[str, str]]]:
    """The formatted lines of a procedure body, and the findings left for a person."""
    text, todos = vba_house.format_module(module_text(body), "M_PROBE")
    lines = text.split("\r\n")
    return lines[lines.index("Private Sub Probe()"):], todos


LEVELS_BEFORE = [
    "If A Then",
    "For I = 1 To 2",
    "Do While B",
    "x = 1",
    "Loop",
    "Next I",
    "ElseIf C Then",
    "With Obj",
    ".X = 1",
    "End With",
    "Else",
    "Select Case Q",
    "Case 1",
    "y = 2",
    "Case Else",
    "y = 3",
    "End Select",
    "End If",
    "If A Then z = 1 Else z = 2",
    "Foo _",
    "  a, _",
    "      b",
]

LEVELS_AFTER = [
    "        If A Then",
    "            For I = 1 To 2",
    "                Do While B",
    "                    x = 1",
    "                Loop",
    "            Next I",
    "        ElseIf C Then",
    "            With Obj",
    "                .X = 1",
    "            End With",
    "        Else",
    "            Select Case Q",
    "                Case 1",
    "                    y = 2",
    "                Case Else",
    "                    y = 3",
    "            End Select",
    "        End If",
    "        If A Then z = 1 Else z = 2",
    "        Foo _",
    "          a, _",
    "              b",
]

# (name, body before, lines the output must hold, lines it must not hold)
BODY_CASES: list[tuple[str, list[str], list[str], list[str]]] = [
    ("a label quoted in a literal is data, not a jump",
     banner("INITIALIZE") + ['    Msg = "use GoTo Fail nowhere"', "    On Error GoTo Fail"],
     ['        Msg = "use GoTo Fail nowhere"', "        On Error GoTo ErrorHandler"], []),
    ("one line, one jump renamed and one literal left alone",
     banner("WORK") + ['    If X Then Msg = "GoTo Fail" Else GoTo Fail'],
     ['        If X Then Msg = "GoTo Fail" Else GoTo ErrorHandler'], []),
    ("comment prose that describes a jump is renamed",
     banner("INITIALIZE") + ["    On Error GoTo Err_Handler 'then Resume Err_Handler"],
     ["        On Error GoTo ErrorHandler 'then Resume ErrorHandler"], []),
    ("the label definitions of the two exit forms are renamed",
     banner("EXIT PROCEDURE") + ["Safe_Exit:", "Exit Sub"] + banner("ERROR HANDLER") + ["Err_Handler:", "Resume Safe_Exit"],
     ["SafeExit:", "ErrorHandler:", "        Resume SafeExit"], ["Err_Handler:", "Safe_Exit:"]),
    ("the demo's exit labels are renamed",
     banner("EXIT PROCEDURE") + ["Clean_Exit:", "Exit Sub"] + banner("ERROR HANDLER") + ["Clean_Fail:", "Resume Clean_Exit"],
     ["CleanExit:", "CleanFail:", "        Resume CleanExit"], ["Clean_Exit:", "Clean_Fail:"]),
    ("the older label Fail becomes ErrorHandler",
     banner("WORK") + ["On Error GoTo Fail", "Fail:"], ["        On Error GoTo ErrorHandler", "ErrorHandler:"], ["Fail:"]),
    ("a bracketed identifier is opaque",
     banner("WORK") + ["    [Fail] = 2"], ["        [Fail] = 2"], []),
    ("PROC becomes PROC_NAME in code, also on a continuation line",
     banner("DECLARE") + ['    Const PROC As String = "M_PROBE.Probe"'] + banner("WORK")
     + ["    Err.Raise 5, PROC, _", '        "PROC stays"', "    x = Obj.PROC"],
     ['    Const PROC_NAME As String = "M_PROBE.Probe"', "        Err.Raise 5, PROC_NAME, _",
      '            "PROC stays"', "        x = Obj.PROC"], []),
    ("an apostrophe inside a literal does not start a comment",
     banner("DECLARE") + ['Const S As String = "a \'b\' c" \'note'],
     ['    Const S As String = "a \'b\' c"   \'note'], []),
    ("declarations align As and the comment on multiples of four",
     banner("DECLARE") + ["Dim A As Long 'x", "Dim LongerName As String 'y"],
     ["    Dim A           As Long     'x", "    Dim LongerName  As String   'y"], []),
    ("statements sit at 8 spaces plus 4 per block, with Select Case two levels deep",
     banner("WORK") + LEVELS_BEFORE, LEVELS_AFTER, []),
    ("one-line intent comments sit 4 left of the code they introduce",
     banner("WORK") + ["'Outside", "If A Then", "        'Inside", "x = 1", "End If"],
     ["    'Outside", "        'Inside", "            x = 1"], []),
    ("a comment with a space after the apostrophe is left where it is",
     banner("WORK") + ["      '   kept as written", "x = 1"], ["      '   kept as written"], []),
    ("a label goes to column 0",
     banner("WORK") + ["    Safe_Exit:"], ["SafeExit:"], []),
    ("a '=' rule becomes a dash rule and the title upper case",
     ["'" + "=" * 78, "' Work step", "'" + "=" * 78, "x = 1"],
     [BANNER, "' WORK STEP", "        x = 1"], ["'" + "=" * 78]),
    ("a section banner has one blank line above and none below",
     ["x = 1", BANNER, "' WORK", BANNER, "", "", "y = 2"],
     ["' WORK", "        y = 2"], []),
    ("statements in a conditional block are indented, the directives are not",
     banner("WORK") + ["#If VBA7 Then", "x = 1", "#Else", "x = 2", "#End If"],
     ["#If VBA7 Then", "        x = 1", "#Else", "        x = 2", "#End If"], []),
    ("RETURN SUCCESS with the clean-exit label becomes EXIT PROCEDURE",
     banner("RETURN SUCCESS") + ["Safe_Exit:", "Exit Sub"] + banner("FAIL") + ["Err_Handler:", "Resume Safe_Exit"],
     ["' EXIT PROCEDURE", "SafeExit:", "        Exit Sub", "' ERROR HANDLER", "ErrorHandler:", "        Resume SafeExit"],
     ["' RETURN SUCCESS", "' FAIL"]),
    ("a label written above its banner moves below it",
     banner("WORK") + ["x = 1", "Clean_Exit:"] + banner("CLEAN UP") + ["Exit Sub"],
     ["' EXIT PROCEDURE", "CleanExit:"], ["' CLEAN UP"]),
    ("a label that follows other statements gets its own banner",
     banner("WORK") + ["x = 1", "Safe_Exit:", "Exit Sub"],
     ["' WORK", "' EXIT PROCEDURE", "SafeExit:"], []),
    ("an Exit statement above the handler gets the EXIT PROCEDURE banner",
     banner("WORK") + ["x = 1", "Exit Sub"] + banner("ERROR HANDLER") + ["Err_Handler:", "Resume Next"],
     ["' EXIT PROCEDURE", "        Exit Sub"], []),
    ("the statement that arms the handler opens INITIALIZE after DECLARE",
     banner("DECLARE") + ["Dim A As Long", "On Error GoTo Err_Handler", "A = 1"],
     ["' INITIALIZE", "        On Error GoTo ErrorHandler"], []),
    ("a task section that starts with the handler statement is split",
     banner("ARRANGE") + ["On Error GoTo Err_Handler", "x = 1"],
     ["' INITIALIZE", "        On Error GoTo ErrorHandler", "' ARRANGE", "        x = 1"], []),
    ("a handler statement after other statements is reported, not moved",
     banner("ARRANGE") + ["x = 1", "On Error GoTo Err_Handler"],
     ["' ARRANGE"], ["' INITIALIZE"]),
]


def case_failures() -> list[str]:
    failures: list[str] = []
    for label, body, present, absent in BODY_CASES:
        try:
            produced, todos = format_body(body)
        except LexicalError as exc:
            failures.append(f"{label}\n      rejected: {exc}")
            continue
        for line in present:
            if line not in produced:
                failures.append(f"{label}\n      expected: {line!r}\n      produced: {produced[12:]!r}")
        for line in absent:
            if line in produced:
                failures.append(f"{label}\n      unexpected: {line!r}")
        again = vba_house.format_module(module_text(body), "M_PROBE")[0]
        if vba_house.format_module(again, "M_PROBE")[0] != again:
            failures.append(f"{label}\n      not idempotent on a second pass")
    produced, todos = format_body(banner("WORK") + ["x = 1", "End If"])
    if "        x = 1" in produced or not any("indentation not changed" in t for _, t in todos):
        failures.append("unbalanced blocks keep their indentation and are reported")
    produced, todos = format_body(banner("ARRANGE") + ["x = 1", "On Error GoTo Err_Handler"])
    if not any("not in INITIALIZE" in t for _, t in todos):
        failures.append("a late handler statement is reported")
    return failures


HEADER_BEFORE = [
    "Public Function Probe_Header(ByVal Value As Long) As Long",
    "'",
    "'==============================================================================",
    "' Probe_Header",
    "'------------------------------------------------------------------------------",
    "' PURPOSE",
    "'   Probe.",
    "'",
    "' WHY",
    "'   Why.",
    "'",
    "' RETURNS",
    "'   A value.",
    "'",
    "' INPUTS",
    "'   Value",
    "'     The value.",
    "'",
    "' CALLED FROM",
    "'   - Elsewhere",
    "'",
    "' ERROR POLICY",
    "'   None.",
    "'",
    "' NOTES",
    "'   Note.",
    "'",
    "' UPDATED",
    "'   2026-11-03 - Changed (#210).",
    "'   2026-11-01 - Created.",
    "'==============================================================================",
    "'",
    "",
    "    Probe_Header = Value",
    "End Function",
]

HEADER_AFTER = [
    "Public Function Probe_Header(ByVal Value As Long) As Long",
    "",
    "'",
    "'------------------------------------------------------------------------------",
    "'                                 PROBE HEADER",
    "'------------------------------------------------------------------------------",
    "' PURPOSE",
    "'   Probe.",
    "'",
    "' WHY THIS EXISTS",
    "'   Why.",
    "'",
    "' INPUTS",
    "'   Value",
    "'     The value.",
    "'",
    "' RETURNS",
    "'   A value.",
    "'",
    "' ERROR POLICY",
    "'   None.",
    "'",
    "' NOTES",
    "'   Note.",
    "'",
    "' CALLED FROM",
    "'   - Elsewhere",
    "'",
    "' HISTORY",
    "'   2026-11-03 - Changed (#210).",
    "'   2026-11-01 - Created.",
    "'",
    "' UPDATED",
    "'   2026-11-03",
    "'------------------------------------------------------------------------------",
    "",
    "    Probe_Header = Value",
    "",
    "End Function",
]


def header_failures() -> list[str]:
    failures: list[str] = []
    text, todos = vba_house.format_module(MODULE_HEAD + "\r\n".join(HEADER_BEFORE) + "\r\n", "M_PROBE")
    produced = text.split("\r\n")
    start = produced.index(HEADER_AFTER[0])
    got = produced[start:start + len(HEADER_AFTER)]
    if got != HEADER_AFTER:
        failures.append(f"header: a legacy header takes the standard form\n      produced: {got!r}")
    if [t for w, t in todos if w == "Probe_Header"] != ["header lacks BEHAVIOR, DEPENDENCIES"]:
        failures.append(f"header: missing fields are reported, not invented ({todos!r})")
    sub = HEADER_BEFORE[:]
    sub[0] = "Public Sub Probe_Header(ByVal Value As Long)"
    sub[12] = "'   None."
    sub[-2:] = ["    Value = 1", "End Sub"]
    text = vba_house.format_module(MODULE_HEAD + "\r\n".join(sub) + "\r\n", "M_PROBE")[0]
    if "' RETURNS\r\n'   Nothing\r\n" not in text:
        failures.append("header: RETURNS of a Sub that says None reads Nothing")
    sub[12] = "'   The count."
    text, todos = vba_house.format_module(MODULE_HEAD + "\r\n".join(sub) + "\r\n", "M_PROBE")
    if "'   The count.\r\n" not in text or not any("RETURNS of a Sub" in t for _, t in todos):
        failures.append("header: free text in RETURNS of a Sub is kept and reported")
    nodate = HEADER_BEFORE[:]
    nodate[28:30] = ["'   Changed once."]
    text, todos = vba_house.format_module(MODULE_HEAD + "\r\n".join(nodate) + "\r\n", "M_PROBE")
    if not any("UPDATED holds no date" in t for _, t in todos):
        failures.append("header: an UPDATED without a date is reported")
    single = HEADER_BEFORE[:]
    single[28:30] = ["'   2026-11-03"]
    text, _ = vba_house.format_module(MODULE_HEAD + "\r\n".join(single) + "\r\n", "M_PROBE")
    if "HISTORY" in text.split("' Probe_Header")[0] or "' HISTORY\r\n'   2026-11-03\r\n" in text:
        failures.append("header: a bare UPDATED date makes no HISTORY section")
    return failures


def refusal_failures() -> list[str]:
    failures: list[str] = []
    for label, body in [
        ("Fail and Err_Handler together", ["Fail:", "Err_Handler:"]),
        ("ErrorHandler beside Err_Handler", ["ErrorHandler:", "Err_Handler:"]),
        ("PROC beside PROC_NAME", ["Const PROC As String = \"x\"", "Const PROC_NAME As String = \"y\""]),
        ("a label name on a continued jump", ["On Error GoTo _", " Fail"]),
        ("an unterminated literal", ['x = "GoTo Fail']),
        ("a non-ASCII literal", ['x = "café"']),
        ("an unterminated conditional", ["#If VBA7 Then"]),
    ]:
        try:
            vba_house.format_module(module_text(body), "M_PROBE")
            failures.append(f"{label} must be rejected")
        except LexicalError:
            pass
    return failures


PROC_TWO = (
    "Private Sub First()\r\n"
    "    Const PROC As String = \"M_PROBE.First\"\r\n"
    "    Err.Raise 5, PROC, \"a\"\r\n"
    "End Sub\r\n"
    "\r\n"
    "Private Sub Second()\r\n"
    "    Const PROC As String = \"M_PROBE.Second\"\r\n"
    "    Err.Raise 6, PROC, \"b\"\r\n"
    "End Sub\r\n"
)


def proc_scope_failures() -> list[str]:
    """PROC is renamed inside the procedures that declare it, and nowhere else."""
    failures: list[str] = []
    head = 'Attribute VB_Name = "M_PROBE"\r\nOption Explicit\r\n\r\n'
    text, _ = vba_house.format_module(head + PROC_TWO, "M_PROBE")
    if text.count("Const PROC_NAME As String") != 2 or text.count("Err.Raise 5, PROC_NAME,") != 1 \
            or text.count("Err.Raise 6, PROC_NAME,") != 1 or " PROC " in text or "PROC," in text:
        failures.append("proc scope: two procedures that each declare Const PROC both move to PROC_NAME")
    mixed = PROC_TWO.replace('Const PROC As String = "M_PROBE.Second"', 'Const PROC_NAME As String = "M_PROBE.Second"')
    mixed = mixed.replace("Err.Raise 6, PROC,", "Err.Raise 6, PROC_NAME,")
    text, _ = vba_house.format_module(head + mixed, "M_PROBE")
    if text.count("Const PROC_NAME As String") != 2 or "Const PROC As" in text:
        failures.append("proc scope: a module half moved to PROC_NAME is finished, not refused")
    for label, extra, where in [
        ("an unrelated Public Sub PROC", "Public Sub PROC()\r\nEnd Sub\r\n", "line 14"),
        ("a module-level variable PROC", "Private PROC As Long\r\n", "line 4"),
        ("a Declare named PROC", 'Private Declare PtrSafe Function PROC Lib "user32" () As Long\r\n', "line 4"),
        ("a parameter named PROC in another procedure", "Private Sub Third(ByVal PROC As Long)\r\nEnd Sub\r\n", "line 14"),
    ]:
        source = head + extra + PROC_TWO if extra.startswith(("Private PROC", "Private Declare")) \
            else head + PROC_TWO + "\r\n" + extra
        try:
            vba_house.format_module(source, "M_PROBE")
            failures.append(f"proc scope: {label} must be refused")
        except LexicalError as exc:
            if "M_PROBE" not in str(exc) or where not in str(exc):
                failures.append(f"proc scope: {label} is refused without module and line: {exc}")
    return failures


def module_failures() -> list[str]:
    failures: list[str] = []
    text, todos = vba_house.format_module(LEGACY_MODULE, "M_PROBE")
    if text != HOUSE_MODULE:
        got, want = text.split("\r\n"), HOUSE_MODULE.split("\r\n")
        at = next((n for n, (a, b) in enumerate(zip(got, want, strict=False)) if a != b), min(len(got), len(want)))
        failures.append(f"module: the legacy fixture does not become the style fixture; line {at + 1}: "
                        f"{got[at:at + 1]!r} instead of {want[at:at + 1]!r}")
    if todos:
        failures.append(f"module: the style fixture reports {todos!r}")
    if vba_house.format_module(HOUSE_MODULE, "M_PROBE")[0] != HOUSE_MODULE:
        failures.append("module: the style fixture is not stable")
    if vba_house_rules.findings(HOUSE_MODULE, "M_PROBE"):
        failures.append("module: the style fixture breaks a rule")
    # Corrupting transformation: a changed statement must be refused before any write.
    import reformat

    def corrupt(text: str, name: str) -> tuple[str, list[tuple[str, str]]]:
        return text.replace("mState = 0", "mState = 1"), []

    with patch.object(vba_house, "format_module", corrupt):
        try:
            reformat.reformat_text(LEGACY_MODULE, "M_PROBE", "house")
            failures.append("module: a changed statement must be refused")
        except LexicalError:
            pass
    # Only the declared renames may differ.
    renamed = LEGACY_MODULE.replace("Safe_Exit", "SafeExit")
    if renamed == LEGACY_MODULE or reformat.reformat_text(renamed, "M_PROBE", "house") != HOUSE_MODULE:
        failures.append("module: a label that already has its house name is left alone")
    return failures


def cli_failures() -> list[str]:
    import reformat
    failures: list[str] = []
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "M_PROBE.bas"
        path.write_bytes(LEGACY_MODULE.encode("ascii"))
        with contextlib.redirect_stdout(io.StringIO()) as out:
            if reformat.check([path], "house") != 1:
                failures.append("cli: --check --style house must fail on the previous layout")
            if reformat.write([path], "house") != 0:
                failures.append("cli: --write --style house failed")
            if path.read_bytes().decode("ascii") != HOUSE_MODULE:
                failures.append("cli: --write --style house did not produce the style fixture")
            if reformat.check([path], "house") != 0:
                failures.append("cli: --check --style house must pass on its own output")
        if "for a person to settle" in out.getvalue():
            failures.append("cli: a conforming module needs nothing from a person")
        os.chmod(path, 0o644)
    return failures


def default_failures() -> list[str]:
    import reformat
    failures: list[str] = []
    for state, want in ((False, "legacy"), (True, "house")):
        with patch.object(house_style, "HOUSE_STYLE_ACTIVE", state):
            if reformat.default_style() != want:
                failures.append(f"the switch {state} must select the {want} normal form")
    return failures


def house_selftest() -> list[str]:
    """Failure descriptions of the house-style formatter; empty means every rule holds."""
    return (case_failures() + header_failures() + refusal_failures() + module_failures()
            + cli_failures() + default_failures() + proc_scope_failures())
