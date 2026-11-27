"""Failing and passing fixtures for the rules of vba_house_rules.py (#210).

Every rule is proved against HOUSE_MODULE, a module that breaks none, with one
change at a time. The gate runs rules_selftest() on every push, whatever the
state of house_style.HOUSE_STYLE_ACTIVE, so a rule stays proved whatever the
switch says.
"""
import house_style
import vba_house_rules
from house_format_fixtures import HOUSE_MODULE

DASHES = "'" + "-" * 78
DELIBERATE = "    " + house_style.DELIBERATE_VALUE_COMMENT + "\r\n"
QUALIFIED_CONST = '    Const PROC_NAME As String = "M_PROBE.Probe_Describe"\r\n'
BARE_CONST = '    Const PROC_NAME As String = "Probe_Describe"\r\n'
EQUALS = "'" + "=" * 78

BODY_END = "        Probe_Describe = Text\r\n"
COMPACT_DECLARED = (
    "\r\n" + DASHES + "\r\n' DECLARE\r\n" + DASHES + "\r\n"
    '    Const PROC_NAME As String = "M_PROBE.Probe_Reset"\r\n\r\n'
    "    Dim Count As Long 'Counter\r\n\r\n"
    + DASHES + "\r\n' RESET\r\n" + DASHES + "\r\n        mState = 0\r\n"
)

# (name, text before, text after, rule that must report, part of its message)
FAILING: list[tuple[str, str, str, str, str]] = [
    ("a procedure header lacks a field", "' NOTES\r\n'   A text such as \"GoTo ErrorHandler\" is data.\r\n'\r\n", "",
     "header", "header lacks NOTES"),
    ("header fields are out of order", "' RETURNS\r\n'   The description.\r\n'\r\n' BEHAVIOR\r\n'   - Picks the text.\r\n",
     "' BEHAVIOR\r\n'   - Picks the text.\r\n'\r\n' RETURNS\r\n'   The description.\r\n",
     "header", "out of order"),
    ("UPDATED is not a single date", "' UPDATED\r\n'   2026-11-02\r\n'---", "' UPDATED\r\n'   2026-11-02 - Added.\r\n'---",
     "header", "UPDATED is not a single"),
    ("the RETURNS of a Sub is not Nothing", "' RETURNS\r\n'   Nothing\r\n", "' RETURNS\r\n'   None.\r\n",
     "header", "RETURNS of a Sub"),
    ("the component title does not name the module", "' MODULE: M_PROBE", "' M_PROBE", "header", "component title"),
    ("a Public procedure may not use the compact header", "Private Sub Probe_Reset()", "Public Sub Probe_Reset()",
     "header", "header lacks WHY THIS EXISTS"),
    ("a procedure without a header", "End Sub\r\n", "End Sub\r\n\r\nPrivate Sub Bare()\r\n    mState = 1\r\nEnd Sub\r\n",
     "header", "no documentation header"),
    ("INITIALIZE after a task section", "' PICK TEXT", "' INITIALIZE", "sections", "out of order"),
    ("the handler is armed outside INITIALIZE", "' INITIALIZE\r\n", "' SET UP\r\n", "sections", "not in INITIALIZE"),
    ("EXIT PROCEDURE holds no Exit statement", "SafeExit:\r\n        Exit Function\r\n", "SafeExit:\r\n        x = 1\r\n",
     "sections", "no Exit statement"),
    ("the handler label is indented", "\r\nErrorHandler:\r\n", "\r\n    ErrorHandler:\r\n", "sections", "not at column 0"),
    ("the PROC_NAME text names another procedure", '"M_PROBE.Probe_Describe"', '"M_PROBE.Other"', "proc-name", "expected"),
    ("DECLARE does not start with PROC_NAME", "    Const PROC_NAME As String", "    Dim Zed As Long\r\n    Const PROC_NAME As String",
     "proc-name", "does not start with Const PROC_NAME"),
    ("a bare PROC_NAME without the deliberate-value comment", QUALIFIED_CONST, BARE_CONST,
     "proc-name", "is the bare name"),
    ("a bare PROC_NAME whose comment is not the exact text", QUALIFIED_CONST,
     DELIBERATE.replace("deliberate", "intended") + BARE_CONST, "proc-name", "is the bare name"),
    ("a bare PROC_NAME whose comment is not directly above it", QUALIFIED_CONST, DELIBERATE + "\r\n" + BARE_CONST,
     "proc-name", "is the bare name"),
    ("a PROC_NAME that is only the module", QUALIFIED_CONST, QUALIFIED_CONST.replace("M_PROBE.Probe_Describe", "M_PROBE"),
     "proc-name", "expected"),
    ("the deliberate-value comment above a qualified PROC_NAME", QUALIFIED_CONST, DELIBERATE + QUALIFIED_CONST,
     "proc-name", "not the bare name"),
    ("the deliberate-value comment above another procedure's name", QUALIFIED_CONST,
     DELIBERATE + QUALIFIED_CONST.replace("M_PROBE.Probe_Describe", "Other"), "proc-name", "not the bare name"),
    ("the constant keeps its previous name", "PROC_NAME", "PROC", "proc-name", "does not start with Const PROC_NAME"),
    ("a procedure with variables has no DECLARE", "' DECLARE\r\n", "' SET UP\r\n", "proc-name", "no DECLARE section"),
    ("a compact procedure whose PROC_NAME names another procedure", "        mState = 0\r\n", COMPACT_DECLARED.replace(
        "M_PROBE.Probe_Reset", "M_PROBE.WrongProcedure"), "proc-name", "expected"),
    ("a compact procedure with variables and no DECLARE", "        mState = 0\r\n",
     "    Dim Count As Long\r\n        mState = 0\r\n", "proc-name", "no DECLARE section"),
    ("a compact procedure that arms its handler outside INITIALIZE", "        mState = 0\r\n",
     DASHES + "\r\n' RESET\r\n" + DASHES + "\r\n        On Error GoTo ErrorHandler\r\n        mState = 0\r\n",
     "sections", "not in INITIALIZE"),
    ("a Select Case without Case Else", '                Case Else\r\n                    Text = "other"\r\n', "",
     "select-case", "without Case Else"),
    ("InStr with two arguments", BODY_END, BODY_END + '        Index = InStr(Text, "a")\r\n', "compare-mode", "InStr"),
    ("InStr with a numeric start and no mode", BODY_END, BODY_END + '        Index = InStr(1, Text, "a")\r\n',
     "compare-mode", "InStr"),
    ("InStr with a mode in the third place, where the search text goes", BODY_END,
     BODY_END + '        Index = InStr(Text, "a", vbTextCompare)\r\n', "compare-mode", "InStr"),
    ("InStr with a start and two strings, whatever they are called", BODY_END,
     BODY_END + "        Index = InStr(Index, Text, vbBinaryCompare)\r\n", "compare-mode", "InStr"),
    ("InStr named without Compare", BODY_END,
     BODY_END + '        Index = InStr(Start:=1, String1:=Text, String2:="a")\r\n', "compare-mode", "InStr"),
    ("StrComp with a mode in the second place", BODY_END,
     BODY_END + "        Index = StrComp(vbTextCompare, Text)\r\n", "compare-mode", "StrComp"),
    ("StrComp with two arguments", BODY_END, BODY_END + "        Index = StrComp(Text, Text)\r\n", "compare-mode", "StrComp"),
    ("a statement separator", BODY_END, BODY_END + "        Index = 1: mState = 2\r\n", "statements", "separator"),
    ("Next without its counter", "        Next Index\r\n", "        Next\r\n", "statements", "Next without"),
    ("an equals rule outside the component header", DASHES + "\r\n' PICK TEXT", EQUALS + "\r\n' PICK TEXT",
     "banners", "outside the component header"),
]

# Text that looks like a deviation and is not one.
PASSING: list[tuple[str, str, str]] = [
    ("a bare PROC_NAME with the deliberate-value comment directly above it", QUALIFIED_CONST, DELIBERATE + BARE_CONST),
    ("a bare PROC_NAME in another letter case with the comment", QUALIFIED_CONST,
     DELIBERATE + BARE_CONST.replace("Probe_Describe", "PROBE_DESCRIBE")),
    ("a compact procedure with a DECLARE that names itself", "        mState = 0\r\n",
     COMPACT_DECLARED.lstrip("\r\n")),
    ("the style fixture as it stands", "", ""),
    ("InStr with a start and a compare mode",
     BODY_END, BODY_END + '        Index = InStr(1, Text, "a", vbTextCompare)\r\n'),
    ("InStr with every argument named",
     BODY_END, BODY_END + '        Index = InStr(Compare:=vbTextCompare, Start:=1, String1:=Text, String2:="a")\r\n'),
    ("InStr with two strings and a named compare", BODY_END,
     BODY_END + '        Index = InStr(Text, "a", Compare:=vbTextCompare)\r\n'),
    ("StrComp with a compare mode", BODY_END, BODY_END + "        Index = StrComp(Text, Text, vbTextCompare)\r\n"),
    ("a call quoted in a literal or a comment", BODY_END,
     BODY_END + '        Text = "InStr(a, b): Next"   \'InStr(a, b)\r\n'),
    ("a named argument is not a separator", BODY_END, BODY_END + "        Index = Foo(Name:=1)\r\n"),
    ("a dash rule in the component header", EQUALS + "\r\n' MODULE: M_PROBE\r\n" + EQUALS,
     DASHES + "\r\n' MODULE: M_PROBE\r\n" + DASHES),
]


def mutate(before: str, after: str) -> str:
    if before and HOUSE_MODULE.count(before) < 1:
        raise ValueError(f"fixture text not found: {before!r}")
    return HOUSE_MODULE.replace(before, after, 1) if before else HOUSE_MODULE


def rules_selftest() -> tuple[list[str], int]:
    """Failure descriptions and the number of fixtures run."""
    failures: list[str] = []
    count = 0
    if vba_house_rules.findings(HOUSE_MODULE, "M_PROBE"):
        failures.append("the style fixture breaks a rule")
    for name, before, after, rule, part in FAILING:
        count += 1
        try:
            found = vba_house_rules.findings(mutate(before, after), "M_PROBE")
        except ValueError as exc:
            failures.append(f"{name}: {exc}")
            continue
        if not any(f.rule == rule and part in f.message for f in found):
            failures.append(f"{name}: expected a {rule} finding with {part!r}, got "
                            f"{[(f.rule, f.message) for f in found]!r}")
    for name, before, after in PASSING:
        count += 1
        found = vba_house_rules.findings(mutate(before, after), "M_PROBE")
        if found:
            failures.append(f"{name}: unexpected {[(f.rule, f.message) for f in found]!r}")
    return failures, count
