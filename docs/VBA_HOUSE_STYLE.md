# 📐 VBA house style

These are the rules for the source in this repository: how exported VBA is
written, formatted, analysed and kept compatible, and which Excel state it may
change. A review checks a change against them.
[CONTRIBUTING.md](../CONTRIBUTING.md) describes how to propose, validate and
submit a change, and [REPOSITORY_STRUCTURE.md](REPOSITORY_STRUCTURE.md) where
each kind of file belongs.

## Source files

Exported source is authoritative.

Portable text (Markdown, Python, YAML and JSON) uses LF. Exported VBA keeps its
existing CRLF policy. Windows-native `.bat`, `.cmd`, `.ps1`, `.psm1`, `.psd1`,
`.vbs`, `.reg` and `.ini` files use CRLF in editors and Git checkouts. This is a
repository consistency policy, not a claim that every host rejects LF.

## Exported-source rules

- Use `Option Explicit`.
- Preserve the repository's VBE export metadata, module names, encoding, and
  line-ending policy.
- Match `.bas`, `.cls`, and `.frm` filenames to their component identity.
- Keep every required `.frm` / `.frx` pair together; treat `.frx` as binary.
- Do not edit a binary form resource as text.
- Do not use a workbook or add-in as the only record of a code change.
- Do not commit Office lock files, recovery copies, local exports, test output,
  or generated binaries unless the repository explicitly designates them as
  source.
- Qualify workbook, worksheet, range, and application references.
- Avoid implicit active-workbook, active-sheet, selection, and default-member
  dependencies.
- Keep `On Error Resume Next` scopes narrow and intentional.
- Preserve useful diagnostic context and clean up on success and failure.
- Avoid new references, APIs, dependencies, or platform assumptions until their
  support and deployment impact is agreed.

## House style (#210)

I write every module of this repository in one form, and the gate holds new
code to it. This section defines the form: layout, banners, headers,
declarations, body sections, statements and names. [Formatter
verification (#52)](#formatter-verification-52) says how `tools/reformat.py`
produces it and proves that doing so changes nothing the code executes.

**The switch.** `HOUSE_STYLE_ACTIVE` in `tools/house_style.py` says whether the
tree is held to this form. It is `False` while the twelve modules are still in
the layout they had before: `tools/reformat.py --check` accepts that layout, the
*house labels* check accepts its spellings (`Err_Handler`, `Safe_Exit`,
`Clean_Exit`, `Clean_Fail`), and the *house-style rules* check runs its fixtures
but does not read the modules. While it is `False`, write to the layout the
modules have and run `python3 tools/reformat.py --write` as before; the form
below replaces it for the whole tree in one change. To see what that change does
to a module, run `python3 tools/reformat.py --style house --write` on a copy, or
`--style house --check` on the module itself. I reformat the modules and set the
constant to `True` in the same change, so that no commit holds a tree the gate
would reject.

### Files

- Line endings are CRLF, the file ends with CRLF, there are no tabs, no trailing
  whitespace and no non-ASCII character. The indent unit is 4 spaces.
- The top of a module is the `Attribute VB_Name` line, one blank line,
  `Option Explicit` (and `Option Private Module` where the module has it), one
  blank line, then the component header.
- Keep lines readable: aim for 100 characters, never pass 120. The contents of a
  string literal, and the trailing comment of an aligned declaration, do not
  count toward the limit; the code around them does.

### Banners

Every banner line is exactly 79 characters: an apostrophe and 78 dashes.

```text
'------------------------------------------------------------------------------
```

There are three kinds.

- **Section banner.** It opens a section of the declarations or of a procedure
  body. The title sits after one space and is UPPERCASE. One blank line goes
  above the banner and none below it: the first declaration, statement, comment
  or label follows the closing rule directly.

  ```text
  '------------------------------------------------------------------------------
  ' PRIVATE CONSTANTS
  '------------------------------------------------------------------------------
  ```

- **Procedure title banner.** Its title is UPPERCASE and centred in the 78
  columns after the apostrophe: `(78 - len(title)) // 2` leading spaces. It is a
  short verb phrase that says what the procedure does, not its identifier.
- **Group banner.** It separates families of procedures in a long module: an
  empty comment line above and below a centred title, all inside two rules and
  surrounded by one empty `'` line.

  ```text
  '
  '------------------------------------------------------------------------------
  '
  '                             FAULT INJECTION
  '
  '------------------------------------------------------------------------------
  '
  ```

A component header may use `=` for its rules, and keeps the character it has;
every other banner uses `-`.

### Headers

A module header and a procedure header use the same fields. The module header
sits after the options and one blank line: a lone `'` line, the rule, the title
`' MODULE: M_EXCEL_UI_RUNTIME` (always `MODULE:`, also for a document module),
the rule, the fields and a closing rule. A procedure header follows the
signature in the same shape, with the procedure title banner in place of the
module title. Each field is an UPPERCASE label indented one space, its text
indented three, with a lone `'` line between fields.

| # | Field | What it says |
|---|---|---|
| 1 | `PURPOSE` | One or two sentences on what it does. |
| 2 | `WHY THIS EXISTS` | The problem it solves, or why it is shaped this way. |
| 3 | `INPUTS` | `None`, or each parameter on its own line (3 spaces) with its description on the lines below (5 spaces), parameters separated by `'`. A module header says `None at module level`. |
| 4 | `RETURNS` | `Nothing` for a `Sub`; otherwise the value and its edge cases. A module header says `Nothing at module level`. |
| 5 | `BEHAVIOR` | What happens, in order. Bullets are `'     - item`. |
| 6 | `ERROR POLICY` | What raises, what is suppressed, what is best effort. `None` is allowed for a trivial getter. |
| 7 | `DEPENDENCIES` | Object models, other components and the members it calls. |
| 8 | `NOTES` | Design decisions, caveats and call examples (code examples indented 7 spaces). `None` when there is nothing to say. |
| 9 | `UPDATED` | The date of the last substantive change, as one `YYYY-MM-DD` date. |

Write `None` rather than omit a field. A `Private` procedure whose body has at
most 15 code lines (comments and blank lines do not count) may use a compact
header: `PURPOSE` and `UPDATED`, plus any other field that says something useful,
in the order above. A `Public`, `Friend` or event procedure always uses all nine.

A header may add sections of its own after the standard fields, in UPPERCASE.
Two have a fixed place in this repository. `CALLED FROM` follows `NOTES`. The
dated lines that used to fill `UPDATED` (`2026-11-15 - Added for #187.`) move,
unchanged, to a `HISTORY` section placed after every other section and before
`UPDATED`, which is always last. The citations in them (`ICR-UI-*`, issue
numbers) stay where the review-crosswalk check counts them.

When I change behaviour I set `UPDATED` to the day of the change; reformatting
alone keeps the date. I do not write a field from what the code seems to do:
where I cannot tell why something exists, the header says what the code shows.

### Declarations

The declarations at the top of a module use section banners, in this order, and
omit the empty ones: `PRIVATE TYPES`, `WINDOWS API DECLARATIONS`,
`PRIVATE CONSTANTS`, `PUBLIC CONSTANTS`, `PUBLIC ENUMS`, `PUBLIC TYPES`,
`PUBLIC STATE`, `PRIVATE STATE`. A module may add sections with a clear name of
its own, such as `TEST FAULT INJECTION STATE`.

- Everything under a section banner is indented 4 spaces.
- Within a run of one-line declarations, `As` sits in one column and the
  trailing comment in another, both on a multiple of 4. A blank line, a comment
  line or a statement ends a run. Every declaration gets a trailing comment with
  no space after the apostrophe (`'Live clock interval`).
- A blank line separates logical groups inside a section. Error-number
  constants form one block per procedure family, with a comment line above it.
- An API declaration keeps the `#If VBA7 Then` / `PtrSafe` / legacy arms nested
  4 spaces per level, one parameter per line after `( _`, each with `ByVal` or
  `ByRef`.

### Procedures

```text
<signature>
                                      one blank line (none after a signature
                                      that continues over several lines)
'
'------------------------------------------------------------------------------
'                              CENTRED TITLE
'------------------------------------------------------------------------------
' PURPOSE ...                         the nine fields, or the compact set
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------
                                      one blank line
<body sections>
                                      one blank line
End Sub
```

- The signature names its visibility (`Public`, `Private` or `Friend`), passes
  every parameter `ByVal` or `ByRef` and gives every parameter and every
  function an `As` type. A signature that does not fit on one line puts each
  parameter on its own line, indented 4 and ending in `, _`; the closing `)`
  follows the last parameter. A signature written once for VBA7 and once for
  the legacy compilation inside `#If` ... `#End If` counts as one signature.
- One blank line stands between `End Sub`, `End Function` or `End Property` and
  the next signature or group banner; two stand between the declarations and
  the first group banner; one stands between a group banner and the signature
  below it.

### Body sections

A body is a sequence of section banners. The standard ones come in this order;
a task section is a short UPPERCASE verb phrase for one logical step.

| Section | Holds |
|---|---|
| `DECLARE` | `Const PROC_NAME As String = "<Module>.<Procedure>"`, a blank line, then every `Dim`, with `As` and the comment aligned |
| `INITIALIZE` | `On Error GoTo ErrorHandler` and the initial state |
| task sections | `VALIDATE INPUTS`, `NORMALIZE INPUTS`, `LOAD SETTINGS`, `RETURN VALUE` and the like |
| `EXIT PROCEDURE` | the clean exit: the `SafeExit:` label, the cleanup, then `Exit Sub` or `Exit Function` |
| `ERROR HANDLER` | the `ErrorHandler:` label, then capture, cleanup and re-raise |

A trivial procedure, such as a one-line getter, keeps only the sections it
needs. `DECLARE`, `INITIALIZE`, `EXIT PROCEDURE` and `ERROR HANDLER` appear only
when there is something to put in them. The demo's cleanup flow, whose clean
exit and failure path are `CleanExit:` and `CleanFail:`, uses the same two
section titles.

### Statements

The examples use illustrative names. Inside a section, a statement or a small
group of closely related statements is preceded by a one-line comment at 4 spaces, with no space after the apostrophe.
The code sits at 8 spaces, and each nested block adds 4. The comment says why or
what, in sentence case and without a final period; a comment of several lines
continues at 4 spaces.

```vba
'------------------------------------------------------------------------------
' VALIDATE REFERENCES
'------------------------------------------------------------------------------
    'Reject a missing owner form
        If OwnerForm Is Nothing Then
            Err.Raise vbObjectError + UI_ERR_NO_OWNER, PROC_NAME, _
                "OwnerForm cannot be Nothing"
        End If
    'Store the owner form reference
        Set mOwnerForm = OwnerForm
```

- Labels (`ErrorHandler:`, `SafeExit:`, `CleanExit:`, `CleanFail:`) stand at
  column 0.
- `Else`, `ElseIf` and `End If` sit at the column of their `If`; `Next`, `Loop`,
  `Wend`, `End With` and `End Select` at the column of their opening statement.
  In a `Select Case`, each `Case` is indented 4 from `Select Case` and its
  statements 4 more, and `Case Else` is always written:

```vba
    'Pick the label for the style
        Select Case Style
            Case UI_StyleFormal
                Label = "Formal"
            Case Else
                Label = "Plain"
        End Select
```

The error handler captures the error before it does anything else:

```vba
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Suppress cleanup errors
        On Error Resume Next
    '...cleanup...
    'Restore normal error handling
        On Error GoTo 0
    'Raise a descriptive error
        Err.Raise ErrorNumber, PROC_NAME, "<What failed>: " & ErrorDescription
```

Raise a custom error as `vbObjectError + <constant>` with `PROC_NAME` as its
source. The rules for saved errors, `Err.LastDllError` and a handler that is
disabled before it re-raises are in the sections on the analyzer and on WinAPI
error capture.

Statement conventions:

- The `VBA.` qualifier on a library call is optional. Write it when the project
  could lose its reference (a plain `Len` then fails to compile while `VBA.Len`
  still does) or when the project defines a name that collides with a library
  function; otherwise keep the style of the module. Prefer the `$` string
  variants, `LenB(x) = 0` or `Len(x) = 0` for an empty test, and `vbNullString`
  rather than `""`.
- `InStr` and `StrComp` name their compare mode (`vbBinaryCompare` or
  `vbTextCompare`).
- `Next` names its counter (`Next CharIndex`).
- One statement per line; no `:` between statements. A single-line
  `If ... Then <one statement>` is allowed for guards and `Exit`.
- Break a long statement after an operator or a comma and indent the
  continuation 4 more than the statement. For `Err.Raise`, keep the number and
  the source on the first line and the message on the next.
- Name local variables in PascalCase and descriptively, with no Hungarian
  prefix (`NormalizedHoverMode`, not `strMode`). Release an object variable with
  `Set x = Nothing` on the exit path.

### Names

- A public or private procedure in a standard module is
  `<Prefix>_<Area>_<Verb><Object>` or `<Prefix>_<Verb><Object>`, with the prefix
  of its module family: `UI_` for the component, `TST_` and `Test_` for the
  regression modules, `DEMO_` and `Demo_` for the demo. An event handler keeps
  the name VBA requires. A class member is PascalCase.
- A constant is `UPPER_SNAKE_CASE` with the project prefix; a Win32 constant
  keeps its SDK name. A module-private variable is `m` followed by PascalCase,
  and a public variable `g` followed by the prefix and a name. An enumeration or
  type is `<Prefix>_PascalCase`, except a Win32 structure, which keeps its SDK
  name.
- The formatter and the gate never rename a procedure, a variable or a
  parameter. Public names are contracts (see below), and an internal name moves
  only in a change of its own.

### Names the style renames

Two names change, because the style fixes them. The formatter rewrites them in
code and in the comments that describe a jump, never inside a string literal,
and refuses a module where the rewrite would be ambiguous.

| Before | After |
|---|---|
| `Const PROC As String` and every use of `PROC` | `Const PROC_NAME As String` and `PROC_NAME` |
| `Err_Handler:`, `Fail:` | `ErrorHandler:` |
| `Safe_Exit:` | `SafeExit:` |
| `Clean_Exit:` | `CleanExit:` |
| `Clean_Fail:` | `CleanFail:` |

### What enforces what

| Rule | Where it is settled |
|---|---|
| Blank lines, banners and their characters, centred titles, header field order and spacing, `HISTORY`, `UPDATED` as one date, `RETURNS` of a `Sub` that says none, declaration alignment, section titles for the exit and the handler, `INITIALIZE` opened at the handler statement, label column, statement and comment indentation | `tools/reformat.py` rewrites them; `--check` and the *house-style formatter* check report drift |
| The renames above | `tools/reformat.py`, which compares the executable tokens before it writes |
| The nine fields present and in order, `UPDATED` a date, `RETURNS` of a `Sub` is `Nothing`, the section order, `On Error GoTo` in `INITIALIZE`, an `Exit` in `EXIT PROCEDURE`, `Const PROC_NAME` first in `DECLARE` with the procedure's own name, `Select Case` without `Case Else`, `InStr` and `StrComp` without a compare mode, one statement per line, `Next` without its counter, an `=` rule outside a component header | the *house-style rules* check (`tools/vba_house_rules.py`), which proves each rule on failing and passing fixtures (`tools/house_rules_fixtures.py`) |
| Header text, the content of a field, a title that says what the procedure does rather than spelling its identifier, local variable names, `vbNullString`, line length | I write and review them; no check rewrites them |

A title that repeats the identifier of its procedure becomes the words of that
identifier in upper case (`UI_SetExcelUI` becomes `UI SET EXCEL UI`); replacing
it by a verb phrase is part of the review.

The formatter never invents the text of a missing field and never adds the
`PROC_NAME` constant, because either would change what a person wrote or what
the module contains. It prints a `note` for each module that names the rules a
person still has to settle. The rule fixtures, the formatter fixtures
(`tools/house_format_fixtures.py`) and the tests under `tests/test_house_style.py`
run whatever the state of the switch, so each rule stays proved while the tree
is not yet held to it.

## Formatter verification (#52)

Exported VBA source is **ASCII only**, with CRLF checkout/output line endings.
The formatter rejects non-ASCII input, including comments, before writing; it
never silently transliterates or changes the VBE code page. The existing Git
attributes and Windows-native file overrides remain authoritative.

`tools/vba_lex.py` is the shared code/string/comment scanner used by the
formatter and VBA analyzer. It recognizes doubled
quotes, apostrophe comments, statement-position `Rem`, and bracketed identifiers.
It is not a complete VBA parser. The analyzer consumes the same lexical
boundaries for procedure-local resolution and targeted assignment analysis.

Two normal forms exist. `--style house` writes the form of the [house
style](#house-style-210); `--style legacy` writes the layout the modules had
before it. Without `--style`, `HOUSE_STYLE_ACTIVE` in `tools/house_style.py`
chooses, so `--check` keeps accepting the tree while it is `False`. The legacy
form protects conditional blocks and continued statements and reports that
their formatting was skipped. The house form (`tools/vba_house.py`) moves a
continued statement as one unit, shifting its continuation lines with its first
line. It copies conditional directives as written; the statements inside a
conditional block of a body are indented like any other, and one-line
declarations inside one are aligned. A procedure whose block statements do not
pair up, or that opens or closes one inside a conditional block, keeps its
indentation and is reported. Malformed literals, interrupted or unterminated
continuations or conditionals, non-ASCII text, an Option statement inside a
conditional block and ambiguous renames fail before any file is written. Line
endings normalize to CRLF.

Before publication, the formatter compares ordered executable tokens and
statement boundaries, preserving literal bytes and Option statements. Only
the renames the house style declares are accepted as differences: the label
names of the [table above](#names-the-style-renames) in a label or jump
position, and the identifier `PROC` where it stands as a name of its own. Option
relocation is normalized. The house form never moves, adds or removes a
statement. This bounded equivalence check is not compilation or Excel runtime
certification. Every fixture is exercised by the existing formatter self-test in
static CI, whatever the state of the switch.

`--write` preflights all inputs and replaces each changed file through a flushed
same-directory temporary file and `os.replace`. Existing permissions are kept,
symlinks are refused, and failed writes/replacements remove their temporary
file. This is atomic per file, not a multi-file transaction or a guarantee
against power loss. The legacy source/destination form uses the same writer.

With `--style house`, `--check` and `--write` also print a `note` per module
naming the rules that rewriting cannot settle (see [What enforces
what](#what-enforces-what)); the note does not change the exit status.

Always inspect the resulting diff and re-import any changed VBA modules into
the VBE before committing them. A formatter change that leaves every module
byte-identical does not create a new Excel runtime result.

## Bounded VBA analyzer (#31)

`python3 tools/vba_analyze.py` checks all `.bas` files in `src/`, `test/` and
`demo/`. The static gate runs it and its independent malformed/benign fixtures
on configured pushes and pull requests. Diagnostics identify the source path,
physical start line, rule and compilation configuration. A continued statement
uses its first physical line. Dynamic dispatch appears as `INVENTORY` in the
gate log; these entries are unresolved host/binding work, not validated calls.

The analyzer collects procedure, DLL, constant, variable, type and enum names;
resolves project-prefixed references against visible declarations; and checks
simple unqualified assignment targets under `Option Explicit`. Parameters,
locals and visible module/project names are recognized after continuation and
colon normalization. This is declaration coverage, not type checking: member,
array-element, `With` and host-object resolution, argument compatibility and
complete expression/loop-variable analysis still require the VBA compiler.

Labels resolve within their procedure, opener/end kinds must agree, and
duplicates are checked independently for VBA7 x64, VBA7 x86 and legacy x86.
Conditional directives use a nested stack. Boolean `VBA7`/`Win64` expressions
and local Boolean `#Const` definitions are supported; unknown symbols or
unsupported expressions fail explicitly, including in inactive nested blocks.
Each effective VBA7 `Declare` requires `PtrSafe`; prefixed DLL declarations
require an explicit `Alias`. The eight known WinAPI contracts have versioned
ABI-v1 fixtures covering parameter types, ByVal passing and return types. This
is not a general DLL export/ABI verifier.

The error-preservation rule covers straight-line handler segments and direct
reads of `Err.Number`, `Err.Source` and `Err.Description` after a call or
`On Error` statement. Arguments are read before the call changes the state;
saved-local re-raises pass. A subsequent error-producing operation under
`On Error Resume Next` is a fresh source, while a scalar flag assignment is
not. Scalar capture intrinsics (`IIf`, `Len`, `CStr`, `CLng`, `CBool`, `CInt`,
`CDbl`, `CSng`) are treated as capture expressions. This does not prove their
arguments cannot raise. Branches, loops and label joins delimit the analysis;
there is no interprocedural or whole-control-flow proof, cleanup proof, or
protection against re-raise loops.

Run `Debug -> Compile VBAProject` and the relevant Excel tests after VBA edits.
Passing static analysis does not produce runtime evidence. The analyzer does not
certify case completeness or execution. The gate's `certification manifest`
check (`tools/certification_manifest.py`, #42) holds the mandatory case
identifiers of `tools/certification-manifest.json` to the registry, dispatch and
call sites in `test/M_EXCEL_UI_TEST_CERTIFICATION.bas`; it reads source text, and
only the certification evidence shows that a case ran.

## WinAPI error capture (#27)

In `src/`, `test/` and `demo/`, read a native error code only through
`Err.LastDllError`, and do not declare `GetLastError`. VBA sets
`Err.LastDllError` after every `Declare` call, the next `Declare` call replaces
it, and `Err.Clear` or any `On Error` statement resets it. Write every native
call whose error you report as three adjacent statements:

```vba
SetLastError 0                                  ' immediately before the call
ApiResult = GetWindowLongPtr(hWnd, GWL_STYLE)   ' result copied to a local
DllError = Err.LastDllError                     ' the very next statement
```

Put nothing between them: no other call, no formatting or logging helper and
no `On Error` statement. Take a regression seam before `SetLastError 0`, and
replace captured values only after the capture. `GetWindowLong[Ptr]` and
`SetWindowLong[Ptr]` can return zero legitimately, so treat zero with a
captured error of zero as a valid value, zero with a non-zero captured error as
a failure, and any non-zero result as success. Report a failure with the
exported API name, the operation, the captured error and the target hWnd.

The gate's `WinAPI error capture` check
(`tools/winapi_capture_fixtures.py`) enforces this sequence separately for VBA7
x64, VBA7 x86 and legacy x86, and proves itself against passing and failing
fixtures first. It reads statement order only: it does not follow `GoTo` and
cannot observe the error a call leaves, so
`TST_Case_TitleBarApiErrorCapture` must still pass on a real Excel host. The
demo's title-bar read follows the same sequence (#144); its gate fixtures also
prove that a defect in a `src/`, `test/` or `demo/` module is reported.

## VBA syntax rules (#203)

Excel compiles a project only on the Windows host, so a statement that cannot
compile used to reach me there. The gate's `VBA syntax` check
(`tools/vba_syntax.py`, with its fixtures in `tools/vba_syntax_fixtures.py`)
reads every module of every configuration of `tools/vba-projects.json` and
fails on these six mistakes:

- **Parentheses.** After the continuation lines of a logical statement are
  joined and its strings and comments are masked, the statement opens or closes
  more parentheses than it matches. A string or bracketed identifier that never
  closes is reported as `lexical`.
- **Blocks.** Inside one procedure, an `If`/`End If`, `Select Case`/`End Select`,
  `With`/`End With`, `For`/`Next` (including `Next i, j`), `Do`/`Loop` or
  `While`/`Wend` that does not pair up, or an `Else`, `ElseIf` or `Case` outside
  its block. Only the first mismatch of a procedure is reported.
- **Continuations.** A logical statement with more than 24 line continuations.
- **Procedure size.** A procedure whose active code reaches 64 KB (65,536
  bytes), from its header to its `End`. Only the code of the compilation being
  read counts: comment-only lines, trailing comments, blank lines and the text
  of inactive conditional branches are dropped, and each statement counts its
  text plus two bytes of line ending. VBA limits the compiled code, which this
  check cannot see, so the measure is a conservative proxy that stays below the
  source size, not the compiled size itself.
- **LongPtr assignment.** Under VBA7 x64, an assignment to a variable declared
  `Long` of an expression that reads a `LongPtr` variable, such as
  `FixtureOwned = EntryStyle And TST_TITLEBAR_OWNED_MASK`. The expression must
  be integer arithmetic over declared variables and constants; wrap it in
  `CLng(...)`, or declare the target `LongPtr`, so that the rule is silent.
- **ByRef type.** A bare variable passed to a `ByRef` parameter of another
  declared scalar type, such as a `Long` passed to a `LongPtr`, in `Call Foo(x)`,
  `Foo x`, `Foo x:=y` and function calls inside expressions. `Foo(x)` as a whole
  statement and `Foo (x)` pass by value and are not read.

Each compilation (VBA7 x64, VBA7 x86 and legacy x86) is read separately with the
same conditional-compilation reader as the analyzer, so an `#If` arm that opens
a block the other arm does not is judged per arm. A finding names the file, the
first physical line, the compilations it holds in, the procedure, the rule and
the statement. `LongPtr` is wider than `Long` only on x64, so the `LongPtr`
rule reads that compilation alone, and a `Long` passed to a `LongPtr` is
reported there only.

Every rule stays silent when the text does not settle the question: a call or
an array element in the expression, a `Variant` or undeclared name, an
argument that is not a bare variable, a procedure name that two modules
declare, an inline `If` that ends the line. A rule that stays silent can miss
a defect; it never reports a statement that compiles. The check passes on the
current modules without suppressions. A conditional directive it cannot read is
reported as `unreadable`, never skipped.

`python3 tools/vba_syntax.py` prints the findings and what each compilation
read. The check is bounded source evidence and does not replace
`Debug -> Compile VBAProject` on a real Excel host.

## Public contracts and compatibility

Treat documented procedures, functions, classes, enums, parameters, defaults,
return values, errors, side effects, workbook formats, and supported platforms
as contracts.

A contract-changing contribution must:

1. identify affected callers and migration needs;
2. explain what changes and what remains unchanged;
3. add or update regression coverage;
4. update user-facing documentation and examples; and
5. state whether the release impact is patch, minor, or major.

Do not make an internal helper public merely to simplify a test. Use an explicit
test seam where the project supports one.

External compatibility is a promise about the `[supported]` facade in
`M_EXCEL_UI`. It is not the same promise as the deployment rule that all four
`src/` modules are replaced together: internal boundaries may move in a patch
release without any caller-visible change, which is why a mixed-version project
fails to compile even when the external API is untouched.

Preserve the `[supported]` facade unless an explicitly approved breaking release
requires otherwise:

- public procedure and function names;
- parameter order;
- optional defaults;
- enum values;
- show/hide/leave-unchanged semantics;
- fire-and-forget behavior;
- ordered structured-result behavior;
- application and window scope;
- `UI_ShowExcelUI` as the emergency show-all operation.

New backward-compatible parameters must normally be optional and trailing.

### Demo cleanup contract (#71)

Demo cleanup (#71) is single-pass, not a host watchdog. With
`ReProtectAtEnd=True`, reset captures the fifteen Boolean `Worksheet.Protect`
options and `EnableSelection` before unprotecting, then attempts to restore
them only if that invocation's unprotect returned successfully. A missing or
incorrect password must not arm reprotection. With `ReProtectAtEnd=False`
(the existing default), leaving the sheet unprotected is intentional.

An operation error retains its number/source and description prefix; cleanup
failures append a separately identified diagnostic. A cleanup-only failure is
raised to the caller. This does not restore deleted cells, shapes, tables or
range-level permissions, nor guarantee recovery after Reset/End or a blocked
Excel call. `UserInterfaceOnly` is restored for the current session, not promised
across workbook reopening. Source fixtures do not establish Excel outcomes.
[CONTRIBUTING.md](../CONTRIBUTING.md#demo-cleanup-checks-71) describes the
host harness that tests this contract.

### Demo fast-mode contract (#70)

`DEMO_FastMode_Begin` owns five Application settings: `ScreenUpdating`,
`EnableEvents`, `DisplayAlerts`, `Calculation` and `Cursor`. It reads all five
before its first write and marks each one as owned immediately before writing
it. When a write fails, Begin writes back every setting it marked, each in its
own attempt, and raises the original error with its number and source; a
failed write-back is appended to the description. `DEMO_FastMode_End` attempts
every owned restoration even after one fails, then gives up ownership and
raises one error that names each setting it could not restore.
`DEMO_FastMode_EndPreservingError` adds that failure to a caller's saved error,
or makes it the error when there is none. The cursor returns to its entry
value.

The worksheet window settings and `DisplayFormulaBar` that
`DEMO_Sheet_BuildTemplate` applies are the requested demo view, not temporary
settings: nothing restores them, and `DisplayFormulaBar` changes every window
of the Excel instance. Like the cleanup contract, this is single-pass: it does
not recover from a blocked Excel call or from Reset/End.

### Demo template validation (#74)

`DEMO_Sheet_BuildTemplate` checks every layout input before it creates or
resets a worksheet: the sheet name against Excel's naming rules, the content
and left-margin columns, the body, freeze and hidden rows, the widths, heights
and zoom. The checks that need the worksheet grid read it from the target or
another worksheet of the workbook, and a name held by a chart sheet is refused.
An invalid request therefore leaves an existing worksheet and the sheet count
as they were. `DEMO_Sheet_GetOrCreate` deletes a sheet it added when the rename
fails, and the template deletes a worksheet it created when it fails later.
Rebuilding an existing worksheet stays destructive: once the inputs are valid,
`DEMO_Sheet_Reset` clears it, and a later failure is reported without undoing
that reset.

Removing manual page breaks (`ResetAllPageBreaks`) and writing
`DisplayPageBreaks` make Excel lay the sheet out through the default printer
driver, which waits on a slow, offline or network printer. The template adds no
manual page breaks, so it resets with `RemovePageBreaks:=False` and keeps any
that exist. `DEMO_Sheet_HidePageBreaks` reads `DisplayPageBreaks` and writes it
only when it is `True`. An explicit `DEMO_Sheet_Reset` with the default
`RemovePageBreaks:=True` still removes manual breaks.

## Excel state ownership

Assume these surfaces belong to the caller or host unless the project explicitly
owns them:

```text
Application.Calculation
Application.EnableEvents
Application.ScreenUpdating
Application.DisplayAlerts
Application.StatusBar
active workbook / worksheet / selection
window styles, shortcuts, timers, names, links, connections, and shapes
```

Capture state before changing it. Restore only state the component successfully
changed and still owns. Cleanup must not conceal the original failure.

## Project engineering contract

| Area | Required behavior |
|---|---|
| **Module boundaries** | Preserve the UI, runtime, snapshot, and title-bar dependency direction. |
| **Window identity** | Never restore per-window state by collection index; handle closed, recreated, and newly opened windows explicitly. |
| **Snapshot lifecycle** | Document partial capture, failure ordering, reset behavior, and emergency recovery. |
| **Title-bar ownership** | Restore only the style bits owned; capture `Err.LastDllError` immediately after each WinAPI call; handle valid zero returns, handle changes, and frame refresh. |
| **Failure behavior** | Production entry points are fail-soft unless documented otherwise; never silently discard failures or leave screen updating suppressed. |
| **Platform support** | Keep the 32-bit and 64-bit declarations correct in every compilation branch; the static gate checks all three. Run affected changes on a real 64-bit Windows Excel host, the only certified bitness. |
