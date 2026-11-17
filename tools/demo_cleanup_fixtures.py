"""#70/#71/#73/#74/#187 source contracts and negative fixtures, not Excel execution evidence."""
import re
from pathlib import Path

from titlebar_control_fixtures import findings as propagation_findings
from vba_analyze import logical_lines
from vba_lex import code_only

PROCEDURES = ('DEMO_Sheet_BuildTemplate', 'DEMO_Sheet_Reset', 'DEMO_FastMode_Begin')
# The restored value passes through the harness seam, which is inert unless armed.
SELECTION = 'WS.EnableSelection = DEMO_CleanupSelectionValue(EntryProtection.EnableSelection)'
ALLOW_OPTIONS = ('AllowFormattingCells', 'AllowFormattingColumns',
                 'AllowFormattingRows', 'AllowInsertingColumns',
                 'AllowInsertingRows', 'AllowInsertingHyperlinks',
                 'AllowDeletingColumns', 'AllowDeletingRows', 'AllowSorting',
                 'AllowFiltering', 'AllowUsingPivotTables')
# The five fast-mode settings in entry order, with the value Begin writes.
FAST_MODE = (('ScreenUpdating', 'False'), ('EnableEvents', 'False'),
             ('DisplayAlerts', 'False'), ('Calculation', 'xlCalculationManual'),
             ('Cursor', 'xlWait'))
# Layout errors of DEMO_Sheet_BuildTemplate that must be raised before its reset.
LAYOUT_ERRORS = (2013, 2014, 2015, 2016, 2017, 2019, 2020, 2021, 2022, 2023, 2024, 2027)
# The one reset of the template: it keeps manual page breaks (#187).
TEMPLATE_RESET = 'demo_sheet_reset ws, removepagebreaks:=false'
HIDE_BREAKS = 'demo_sheet_hidepagebreaks ws'
CALLER_END = ('DEMO_FastMode_EndPreservingError FastModeState, SavedErrNumber, '
              'SavedErrSource, SavedErrDescription')


def _code(block: str) -> list[str]:
    lines = (re.sub(r'\s+', ' ', code_only(raw).strip().lower()) for _, raw in logical_lines(block))
    return [line for line in lines if line]


def _procedure(source: str, kind: str, name: str) -> list[str] | None:
    match = re.search(rf'^(?:Public|Private) {kind} {name}\b.*?^End {kind}', source, re.M | re.S)
    return _code(match[0]) if match else None


def reset_findings(source: str) -> list[str]:
    match = re.search(r'^Public Sub DEMO_Sheet_Reset\b.*?^End Sub', source, re.M | re.S)
    if not match:
        return ['reset procedure missing']
    lines = [code_only(raw).strip().lower() for _, raw in logical_lines(match[0])]
    errors = []
    unprotect = 'ws.unprotect password:=protectpassword'
    ownership = 'didunprotect = true'
    if (lines.count(unprotect) != 1 or lines.count(ownership) != 1
            or lines.index(ownership) != lines.index(unprotect) + 1):
        errors.append('ownership must be armed only after Unprotect returns')
    if 'if didunprotect and reprotectatend then' not in lines:
        errors.append('reprotect requires invocation ownership and opt-in')
    cleanup = lines[lines.index('clean_exit:'):lines.index('clean_fail:')]
    if cleanup.count('on error resume next') != 1 or cleanup.count('on error goto 0') != 1:
        errors.append('cleanup requires one protected attempt then disabled handling')
    protects = [line for line in cleanup if line.startswith('ws.protect ')]
    if len(protects) != 1:
        return errors + ['expected one Protect attempt']
    for operation, prefix in [(protects[0], 'cleanuperr'),
                              (SELECTION.lower(), 'selectionerr')]:
        captures = [f'{prefix}{field} = err.{field}' for field in ('number', 'source', 'description')]
        if operation not in cleanup:
            errors.append(f'{prefix}: cleanup operation missing')
        else:
            pos = cleanup.index(operation)
            if cleanup[pos + 1:pos + 4] != captures:
                errors.append(f'{prefix}: immediate error capture missing')
    options = [('DrawingObjects', 'WS.ProtectDrawingObjects'),
               ('Contents', 'WS.ProtectContents'), ('Scenarios', 'WS.ProtectScenarios'),
               ('UserInterfaceOnly', 'WS.ProtectionMode')]
    options += [(name, '.' + name) for name in ALLOW_OPTIONS]
    for name, getter in options:
        capture = f'entryprotection.{name} = {getter}'.lower()
        binding = f'{name}:=EntryProtection.{name}'.lower()
        if (capture not in lines or binding not in protects[0]
                or (unprotect in lines and lines.index(capture) > lines.index(unprotect))):
            errors.append(f'protection option {name} is not captured before unprotect and rebound')
    if SELECTION.lower() not in cleanup:
        errors.append('selection policy restoration missing')
    return errors


def _begin_findings(begin: list[str]) -> list[str]:
    errors = []
    writes = [i for i, line in enumerate(begin) if line.startswith('application.')]
    first_write = writes[0] if writes else len(begin)
    for armed in ('stateout.captured = true', 'on error goto clean_fail'):
        if armed not in begin or begin.index(armed) > first_write:
            errors.append(f'fast mode: {armed!r} must precede the first write')
    for name, value in FAST_MODE:
        capture = f'stateout.{name} = application.{name}'.lower()
        if capture not in begin or begin.index(capture) > first_write:
            errors.append(f'fast mode: {name} is not captured before the first write')
        write = f'application.{name} = demo_fastmodevalue(demo_fm_{name}, false, {value})'.lower()
        mark = f'stateout.applied = stateout.applied or demo_fm_{name}'.lower()
        if write not in begin or begin[begin.index(write) - 1] != mark:
            errors.append(f'fast mode: {name} must be marked owned immediately before its write')
    cleanup = begin[begin.index('clean_exit:'):] if 'clean_exit:' in begin else []
    if 'rollbackfailures = demo_fastmode_restore(stateout)' not in cleanup:
        errors.append('fast mode: a failed entry must roll back what it wrote')
    return errors


def _restore_findings(restore: list[str]) -> list[str]:
    errors = []
    captures = [f'fail{field} = err.{field}' for field in ('number', 'source', 'description')]
    for name, _ in FAST_MODE:
        rewrite = f'application.{name} = demo_fastmodevalue(demo_fm_{name}, true, state.{name})'.lower()
        if rewrite not in restore:
            errors.append(f'fast mode: {name} restoration missing')
            continue
        pos = restore.index(rewrite)
        if restore[pos - 1] != 'err.clear' or restore[pos + 1:pos + 4] != captures:
            errors.append(f'fast mode: {name} restoration needs a cleared error and its immediate capture')
    handlers = [line for line in restore if line.startswith('on error ')]
    if handlers != ['on error resume next', 'on error goto 0']:
        errors.append('fast mode: every restoration must run under one Resume Next scope')
    elif 'state.applied = 0' not in restore[restore.index('on error goto 0'):]:
        errors.append('fast mode: ownership is released only after every restoration attempt')
    return errors


def _caller_findings(caller: list[str]) -> list[str]:
    errors = []
    if 'clean_exit:' not in caller or 'clean_fail:' not in caller:
        return ['demo build: Clean_Exit/Clean_Fail labels missing']
    tail = caller[caller.index('clean_exit:'):caller.index('clean_fail:')]
    if tail[:2] != ['clean_exit:', 'on error goto 0'] or CALLER_END.lower() not in tail:
        errors.append('demo build: cleanup must disable its handler, then end fast mode keeping the error')
    if any(line.startswith('application.cursor') for line in caller):
        errors.append('demo build: the cursor belongs to the fast-mode scope')
    return errors


def fast_mode_findings(source: str, demo: str) -> list[str]:
    """#70: capture before the first write, mark before each write, restore every mark."""
    begin = _procedure(source, 'Sub', 'DEMO_FastMode_Begin')
    restore = _procedure(source, 'Function', 'DEMO_FastMode_Restore')
    end = _procedure(source, 'Sub', 'DEMO_FastMode_End')
    caller = _procedure(demo, 'Sub', 'Demo_CreateDemoSheet')
    if begin is None or restore is None or end is None or caller is None:
        return ['fast-mode procedure or Demo_CreateDemoSheet missing']
    errors = _begin_findings(begin) + _restore_findings(restore) + _caller_findings(caller)
    if 'failures = demo_fastmode_restore(statein)' not in end:
        errors.append('fast mode: End must restore through the shared restoration')
    return errors


def checkbox_findings(source: str, demo: str) -> list[str]:
    """#73: the shape type dispatches Forms and ActiveX check boxes before any refusal."""
    errors = []
    resolve = _procedure(source, 'Function', 'DEMO_CB_ResolveControl')
    if resolve is None:
        return ['check boxes: DEMO_CB_ResolveControl missing']
    cases = [line for line in resolve if line.startswith('case ')]
    if cases != ['case msoformcontrol', 'case msoolecontrolobject', 'case else']:
        errors.append('check boxes: the shape type must route Forms and ActiveX controls before refusing a shape')
    for name in ('DEMO_CB_TryGetState', 'DEMO_CB_TrySetState'):
        proc = _procedure(source, 'Function', name)
        if proc is None or not any('demo_cb_resolvecontrol(' in line for line in proc):
            errors.append(f'check boxes: {name} must resolve the control through DEMO_CB_ResolveControl')
    if re.search(r'\bDemo_Try(?:Get|Set)CheckBoxState\b', demo):
        errors.append('check boxes: M_EXCEL_UI_DEMO must use DEMO_CB_TryGetState and DEMO_CB_TrySetState')
    return errors


def _raise_lines(lines: list[str], numbers: tuple[int, ...]) -> list[int]:
    found = []
    for index, line in enumerate(lines):
        match = re.match(r'err\.raise vbobjecterror \+ (\d+)\b', line)
        if match and int(match[1]) in numbers:
            found.append(index)
    return found


def template_findings(source: str) -> list[str]:
    """#74: validate before any reset or creation, and remove only an owned new sheet."""
    errors = []
    build = _procedure(source, 'Sub', 'DEMO_Sheet_BuildTemplate')
    create = _procedure(source, 'Function', 'DEMO_Sheet_GetOrCreate')
    if build is None or create is None:
        return ['template: DEMO_Sheet_BuildTemplate or DEMO_Sheet_GetOrCreate missing']
    reset = TEMPLATE_RESET
    if build.count(reset) != 1 or sum(line.startswith('demo_sheet_reset ') for line in build) != 1:
        return ['template: expected exactly one reset of the target worksheet']
    checks = _raise_lines(build, LAYOUT_ERRORS)
    find = next((i for i, line in enumerate(build) if 'demo_sheet_findworksheet(wb, ws_name,' in line), None)
    if not checks or max(checks) > build.index(reset) or find is None or find > build.index(reset):
        errors.append('template: every layout check must raise before the worksheet is reset')
    armed = build.index('on error goto clean_fail') if 'on error goto clean_fail' in build else -1
    for pure in ('nameproblem = demo_sheetnameproblem(ws_name)', 'contenttext = trim$(contentcolumns)'):
        if pure not in build or build.index(pure) > armed:
            errors.append(f'template: {pure!r} must run before the workbook is touched')
    owners = [i for i, line in enumerate(build) if line == 'createdsheet = true']
    if not owners or any(build[i - 1] != 'set ws = demo_sheet_getorcreate(wb, ws_name)' for i in owners):
        errors.append('template: only a worksheet this call created may be marked as owned')
    cleanup = build[build.index('clean_exit:'):] if 'clean_exit:' in build else []
    guard = 'if savederrnum <> 0 and createdsheet then'
    if guard not in cleanup or cleanup[cleanup.index(guard) + 1] != 'cleanupfailures = demo_sheet_deleteowned(ws)':
        errors.append('template: a failed call must remove the worksheet it created, and only that one')
    add = next((i for i, line in enumerate(create) if line.startswith('set ws = wb.worksheets.add(')), 0)
    if 'nameproblem = demo_sheetnameproblem(sheetname)' not in create[:add]:
        errors.append('get-or-create: the name must be checked before a sheet is added')
    rename = 'ws.name = demo_sheetnamevalue(sheetname)'
    if rename not in create:
        return errors + ['get-or-create: rename missing']
    pos = create.index(rename)
    captures = [f'nameerr{field} = err.{field}' for field in ('number', 'source', 'description')]
    tail = create[pos:]
    if create[pos + 1:pos + 4] != captures or 'cleanupfailures = demo_sheet_deleteowned(ws)' not in tail:
        errors.append('get-or-create: a failed rename must delete the new sheet before re-raising')
    return errors


def page_break_findings(source: str, demo: str) -> list[str]:
    """#187: no needless printer query from the template or a hidden-indicator reset."""
    errors = []
    build = _procedure(source, 'Sub', 'DEMO_Sheet_BuildTemplate')
    reset = _procedure(source, 'Sub', 'DEMO_Sheet_Reset')
    hide = _procedure(source, 'Sub', 'DEMO_Sheet_HidePageBreaks')
    if build is None or reset is None or hide is None:
        return ['page breaks: DEMO_Sheet_BuildTemplate, DEMO_Sheet_Reset or DEMO_Sheet_HidePageBreaks missing']
    if (TEMPLATE_RESET not in build or HIDE_BREAKS not in build
            or build.index(HIDE_BREAKS) < build.index(TEMPLATE_RESET)):
        errors.append('page breaks: the template must keep manual breaks and hide shown indicators after its reset')
    calls = [i for i, line in enumerate(reset) if line == 'ws.resetallpagebreaks']
    if len(calls) != 1 or reset[calls[0] - 1] != 'demo_cleanupprobehit demo_probe_page_breaks':
        errors.append('page breaks: the reset must remove manual breaks in one counted ResetAllPageBreaks call')
    if HIDE_BREAKS not in reset:
        errors.append('page breaks: the reset must hide indicators through DEMO_Sheet_HidePageBreaks')
    read, guard, write = 'shown = ws.displaypagebreaks', 'if shown then', 'ws.displaypagebreaks = false'
    if (hide.count(write) != 1 or read not in hide or guard not in hide
            or hide[hide.index(guard) - 1] != read or hide.index(write) < hide.index(guard)
            or hide[hide.index(write) - 1] != 'demo_cleanupprobehit demo_probe_display_breaks'):
        errors.append('page breaks: DisplayPageBreaks must be read first and written, counted, only when True')
    writes = sum(1 for line in _code(source) if re.match(r'(?:\w+\.)?displaypagebreaks\s*=', line))
    resets = sum(1 for line in _code(source) if line.endswith('.resetallpagebreaks'))
    if writes != 1 or resets != 1:
        errors.append('page breaks: DisplayPageBreaks and ResetAllPageBreaks belong to the reset and its helper only')
    if re.search(r'\b(?:DisplayPageBreaks|ResetAllPageBreaks)\b', ''.join(_code(demo)), re.I):
        errors.append('page breaks: M_EXCEL_UI_DEMO must leave page breaks to M_DEMO_BUILDER')
    return errors


def _page_break_selftest(source: str, demo: str) -> tuple[list[str], int]:
    """#187 negative fixtures: each mutation must be reported."""
    errors: list[str] = []
    count = 0
    page_break_mutations = [
        ('        DEMO_Sheet_Reset WS, RemovePageBreaks:=False\n', '        DEMO_Sheet_Reset WS\n'),
        ("    'Hide page-break indicators, writing the property only when they are shown\n"
         '        DEMO_Sheet_HidePageBreaks WS\n', ''),
        ('                DEMO_Sheet_HidePageBreaks WS\n', '                WS.DisplayPageBreaks = False\n'),
        ('            If Shown Then\n', '            If True Then\n'),
        ('            Shown = WS.DisplayPageBreaks\n', ''),
        ('                DEMO_CleanupProbeHit DEMO_PROBE_PAGE_BREAKS\n', ''),
        ('                DEMO_CleanupProbeHit DEMO_PROBE_DISPLAY_BREAKS\n', ''),
    ]
    for old, new in page_break_mutations:
        count += 1
        if source.count(old) != 1 or not page_break_findings(source.replace(old, new), demo):
            errors.append(f'negative page-break fixture not detected: {old.strip()}')
    count += 1
    if not page_break_findings(source, demo + '    ActiveSheet.ResetAllPageBreaks\n'):
        errors.append('negative page-break fixture not detected: a page-break call in M_EXCEL_UI_DEMO')
    return errors, count


def selftest(root: str | Path) -> tuple[list[str], int]:
    source = (Path(root) / 'demo/M_DEMO_BUILDER.bas').read_text(encoding='ascii')
    demo = (Path(root) / 'demo/M_EXCEL_UI_DEMO.bas').read_text(encoding='ascii')
    errors = (reset_findings(source) + fast_mode_findings(source, demo)
              + checkbox_findings(source, demo) + template_findings(source)
              + page_break_findings(source, demo))
    count = 0
    for name in PROCEDURES:
        errors += propagation_findings(source, name, 'clean_exit', 'clean_fail')
        fixture = f'''Public Sub {name}()
On Error GoTo Clean_Fail
Clean_Exit:
On Error GoTo 0
Err.Raise savedNumber, savedSource, savedDescription
Exit Sub
Clean_Fail:
Resume Clean_Exit
End Sub'''
        for replacement, expected in [('On Error GoTo 0', False), ('', True),
                                      ('On Error Resume Next', True),
                                      ('On Error GoTo Clean_Fail', True)]:
            count += 1
            mutated = fixture.replace('On Error GoTo 0', replacement)
            if bool(propagation_findings(mutated, name, 'clean_exit', 'clean_fail')) != expected:
                errors.append(f'{name}: propagation fixture {replacement!r} failed')
    mutations = [
        ('DidUnprotect And ReProtectAtEnd', 'WasProtected And ReProtectAtEnd'),
        ('DidUnprotect = True', 'DidUnprotect = False'),
        (SELECTION, 'Debug.Print 0'),
        (SELECTION, 'WS.EnableSelection = DEMO_CleanupSelectionValue(0)'),
    ] + [(f'{name}:=EntryProtection.{name}', f'{name}:=False') for name in ALLOW_OPTIONS]
    mutations += [(f'{prefix}{field} = Err.{field}', f'{prefix}{field} = 0')
                  for prefix in ('CleanupErr', 'SelectionErr')
                  for field in ('Number', 'Source', 'Description')]
    for old, new in mutations:
        count += 1
        if old not in source or not reset_findings(source.replace(old, new)):
            errors.append(f'negative source fixture not detected: {old}')
    fast_mutations = [
        ('        StateOut.Captured = True\n', ''),
        ('        StateOut.Cursor = Application.Cursor\n', ''),
        ('        On Error GoTo Clean_Fail\n', ''),
        ('        StateOut.Applied = StateOut.Applied Or DEMO_FM_CALCULATION\n', ''),
        ('            RollbackFailures = DEMO_FastMode_Restore(StateOut)\n', ''),
        ('            FailNumber = Err.Number\n', '            FailNumber = 0\n'),
        ('        Err.Clear\n        On Error GoTo 0\n',
         '        Err.Clear\n        On Error GoTo 0\n        On Error Resume Next\n'),
        ('            Application.Cursor = DEMO_FastModeValue(DEMO_FM_CURSOR, True, State.Cursor)\n', ''),
    ]
    for old, new in fast_mutations:
        count += 1
        if old not in source or not fast_mode_findings(source.replace(old, new), demo):
            errors.append(f'negative fast-mode fixture not detected: {old.strip()}')
    demo_mutations = [
        ('        DEMO_FastMode_EndPreservingError FastModeState, SavedErrNumber, _\n'
         '            SavedErrSource, SavedErrDescription\n', '        DEMO_FastMode_End FastModeState\n'),
        ('        DEMO_FastMode_Begin FastModeState\n',
         '        DEMO_FastMode_Begin FastModeState\n        Application.Cursor = xlWait\n'),
    ]
    for old, new in demo_mutations:
        count += 1
        if old not in demo or not fast_mode_findings(source, demo.replace(old, new)):
            errors.append(f'negative demo-build fixture not detected: {old.strip()}')
    checkbox_mutations = [
        ('                Case msoOLEControlObject\n                    Set Shp = Nothing\n', ''),
        ('        Select Case DEMO_CB_ResolveControl(WS, ControlName, Shp, CheckBoxOle, FailMsg)\n',
         '        Select Case DEMO_CB_KIND_FORMS\n'),
    ]
    for old, new in checkbox_mutations:
        count += 1
        if old not in source or not checkbox_findings(source.replace(old, new), demo):
            errors.append(f'negative check-box fixture not detected: {old.strip()}')
    count += 1
    if not checkbox_findings(source, demo + 'Private Function Demo_TryGetCheckBoxState()\n'):
        errors.append('negative check-box fixture not detected: a private copy in M_EXCEL_UI_DEMO')
    template_mutations = [
        ('        MaxCol = LimitSheet.Columns.Count\n',
         '        MaxCol = LimitSheet.Columns.Count\n        DEMO_Sheet_Reset WS, RemovePageBreaks:=False\n'),
        ('        DEMO_Sheet_Reset WS, RemovePageBreaks:=False\n', ''),
        ('        If SavedErrNum <> 0 And CreatedSheet Then\n', '        If SavedErrNum <> 0 Then\n'),
        ('            CleanupFailures = DEMO_Sheet_DeleteOwned(WS)\n', ''),
        ('        NameProblem = DEMO_SheetNameProblem(SheetName)\n', ''),
        ('            NameErrNumber = Err.Number\n', '            NameErrNumber = 0\n'),
    ]
    for old, new in template_mutations:
        count += 1
        if old not in source or not template_findings(source.replace(old, new)):
            errors.append(f'negative template fixture not detected: {old.strip()}')
    found, extra = _page_break_selftest(source, demo)
    return errors + found, count + extra
