# VBA call checker

`tools/check_vba_calls.py` resolves the calls between the modules of each VBA
project this repository can build, and reports the ones that cannot work: a call
to a procedure no module declares, a Private procedure called from another
module, a wrong argument list, or a macro name that names no procedure. It reads
the source only, so it runs anywhere Python runs, without Excel.

It is advisory. It is not a VBA compiler, and it does not replace compiling the
project in the VBE (**Debug > Compile VBAProject**) or running
`Test_EXCEL_UI_RunReleaseCertification` on a real Excel host. Those stay the
authoritative checks.

## Contents

- [What it checks](#what-it-checks)
- [Projects and environments](#projects-and-environments)
- [Macro targets](#macro-targets)
- [What it leaves unknown](#what-it-leaves-unknown)
- [Run it](#run-it)
- [Why it is advisory](#why-it-is-advisory)

## What it checks

The checker reports only what it can resolve with certainty. A call it cannot
resolve is counted as unknown, never reported as a defect.

| Code | Severity | Finding |
| --- | --- | --- |
| `VBA-CALL-001` | error | An unqualified call to a project procedure that no module of the project declares |
| `VBA-CALL-002` | error | A module-qualified call (`M_EXCEL_UI.UI_Name`) to a member the module does not declare |
| `VBA-CALL-003` | error | A Private procedure, variable or `AddressOf` target used from another module |
| `VBA-CALL-004` | error | An unqualified name that is Public in more than one standard module |
| `VBA-CALL-005` | error | A class without a default instance used as an object |
| `VBA-CALL-010` | error | A required argument is not supplied |
| `VBA-CALL-011` | error | More positional arguments than parameters, and no `ParamArray` |
| `VBA-CALL-012` | error | A required positional argument is left empty (`UI_Name , 2`) |
| `VBA-CALL-013` | error | A named argument that matches no parameter |
| `VBA-CALL-014` | error | An argument supplied twice, by name or by name and position |
| `VBA-CALL-015` | error | A `ParamArray` parameter passed by name |
| `VBA-CALL-016` | error | A positional argument after a named argument |
| `VBA-CALL-017` | error | A property used in a way that needs a `Get`, `Let` or `Set` the property does not declare |
| `VBA-CALL-020` | error | One name declared twice in a module |
| `VBA-CALL-030` | error | A literal macro or member target that names no procedure of the project |
| `VBA-CALL-031` | warning | A literal macro target that resolves only to a Private procedure |
| `VBA-CALL-040` | warning | A Public name exported by more than one standard module |
| `VBA-CALL-090` | failure | A module the checker could not read or parse |
| `VBA-CALL-091` | failure | Coverage is incomplete: a tracked module belongs to no project, or a listed module is missing |

Argument checks count `Optional` parameters, defaults, `ParamArray` and named
arguments (`Name:=value`). Property procedures are matched to the accessor
their use needs: a read needs `Property Get`, an assignment `Property Let`, and
a `Set` assignment `Property Set`.

A name counts as a project name when it matches one of the
`project_name_patterns` of `tools/vba-projects.json` (`^M_`, `^UI_`, `^TST_`,
`^Test_EXCEL_UI_` and `^Demo_`, without regard to case). Every module name and
every `Sub`, `Function` and `Property` in this repository follows those
prefixes, so an unresolved call to such a name is a defect. An unresolved name
that matches no pattern may be a VBA, Excel or library name, so it stays
unknown. Calls to the Windows API `Declare` statements resolve like any other
module member.

## Projects and environments

A workbook holds one VBA project, so the checker analyses each configuration of
[`tools/vba-projects.json`](../tools/vba-projects.json) as a separate project
and never combines modules from different configurations:

| Configuration | Modules |
| --- | --- |
| `production` | the four modules of `src/` |
| `development` | `src/` and `test/M_EXCEL_UI_REGRESSION_TESTS.bas` |
| `demo` | `src/` and the two modules of `demo/` |

A call or macro target that resolves only to a module outside a configuration
is a defect in that configuration, because that workbook cannot reach it.

Each configuration is analysed once per conditional-compilation environment of
`tools/vba_analyze.py`: `vba7-x64` (VBA7 and Win64), `vba7-x86` (VBA7 without
Win64) and `legacy-x86` (neither). Branches of one `#If` block are never
combined. An error that occurs in some environments but not in all of them is
reported as unknown, with the environments it occurs in, rather than as a
defect. A conditional-compilation symbol other than `VBA7`, `Win64` or a module
`#Const` makes the module unreadable (`VBA-CALL-090`): the checker fails closed
instead of guessing.

## Macro targets

The checker resolves a macro name when it is fixed in the source:

- the first argument of `Application.Run`;
- the procedure argument of `Application.OnTime` and `Application.OnKey`,
  by position or as `Procedure:=`;
- the value assigned to an `OnAction` property, with or without a `With` block;
- the member name of `CallByName` when the object is a variable of a project
  class;
- the procedure after `AddressOf`.

A macro name may be a string literal, a constant, or a concatenation of them. A
leading `"'" & ThisWorkbook.Name & "'!"` or `ThisWorkbook.Name & "!"` names the
workbook that holds the code, so the checker removes it and resolves the rest in
the project. A name qualified by another workbook (`'Other.xlsm'!Macro`) may be
outside the project, so it is never reported missing.

A procedure can also pass a macro name on. When a procedure assigns one of its
own `ByVal` parameters to `OnAction` (or hands it to another sink) and never
assigns that parameter, the parameter carries a macro name. Every resolved call
to that procedure then has its argument for the parameter checked as a macro
target, and a call that omits the argument has the parameter's literal default
checked instead. This is how the checker follows the demo builder:
`DEMO_Btn_Add` in `demo/M_DEMO_BUILDER.bas` assigns its `ActionMacro` parameter
to the button's `OnAction`, so each literal macro name passed to `DEMO_Btn_Add`
must name a procedure of the `demo` configuration.

## What it leaves unknown

These are counted, by reason, in the report and the job summary, and never
reported as defects:

- calls to names that match no project pattern (VBA, Excel and library names);
- members of objects whose type is not a project module or class, including
  `Object`, `Variant` and Excel types, and members of an unresolved `With`;
- macro names computed at run time, read from arrays or cells, or held in a
  variable that is assigned more than once (the button grids of
  `DEMO_Btn_AddGrid` are of this kind);
- `ByRef` parameters, whose value the called code may change;
- the types of arguments, and the signatures of callbacks and event handlers;
- whether a procedure is ever called.

The checker does not compile VBA, so a module it accepts can still fail to
compile in Excel.

## Run it

From the repository root, with Python 3.11 or later and nothing else installed:

```text
python3 tools/check_vba_calls.py
python3 tools/check_vba_calls.py --output test-results/check_vba_calls.json --summary summary.md
```

| Option | Effect |
| --- | --- |
| `--root PATH` | Analyse the repository at `PATH` instead of this checkout |
| `--output FILE` | Write the full report as JSON: every finding, the per-configuration counts, the macro entry points and the procedure inventory |
| `--summary FILE` | Append a Markdown summary to `FILE`, as a CI job summary expects |

The checker reads the modules Git tracks, so add a new module to the index and
to a configuration of `tools/vba-projects.json` before you run it. It prints a
summary for each configuration and every error, warning and analysis failure,
with its file and line.

| Exit status | Meaning |
| --- | --- |
| `0` | The analysis is complete and found no defect |
| `1` | It found a defect, or the analysis is incomplete |
| `2` | A usage error, or the manifest or a module could not be read |

Its tests are in `tests/test_vba_calls.py`: each rule has a fixture it must
report and a valid twin it must accept, and one test runs the checker on this
repository's own modules.

## Why it is advisory

I run the checker in the `Static checks` workflow as a non-blocking step after
the required checks, and I keep it out of `tools/check_repo.py`. A static model
of VBA name resolution can be wrong where the compiler is not, and a false
defect must not block a merge that Excel would accept. Its verdict goes to the
job summary and to the `check-vba-calls` artifact, where I review it; its unit
tests are part of the required job, so a checker that stopped detecting a
defect still fails the build.

When it reports a defect, confirm it by compiling the affected configuration in
Excel. If Excel agrees, fix the VBA. If Excel compiles and runs the code, the
finding is a checker defect: fix the checker and add the case to its tests.
