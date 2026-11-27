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

Conditional blocks (including their directives and whitespace) and continued
logical statements are preserved. `--check` reports that their formatting was
skipped; clean preserved input may pass. Malformed literals, interrupted or
unterminated continuations/conditionals, non-ASCII text and ambiguous label
renames fail before any file is written. Line endings normalize to CRLF.

Before publication, the formatter compares ordered executable tokens and
statement boundaries, preserving literal bytes and Option statements; only the
explicit house-label renaming and Option relocation are normalized. This
bounded equivalence check is not compilation or Excel runtime certification.
Every fixture is exercised by the existing formatter self-test in static CI.

`--write` preflights all inputs and replaces each changed file through a flushed
same-directory temporary file and `os.replace`. Existing permissions are kept,
symlinks are refused, and failed writes/replacements remove their temporary
file. This is atomic per file, not a multi-file transaction or a guarantee
against power loss. The legacy source/destination form uses the same writer.

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
Passing static analysis does not produce runtime evidence. The mandatory
certification-manifest registration check remains dependent on #42's manifest;
the current analyzer does not certify case completeness or execution.

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
