"""#70/#71 source contracts and negative fixtures, not Excel execution evidence."""
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


def selftest(root: str | Path) -> tuple[list[str], int]:
    source = (Path(root) / 'demo/M_DEMO_BUILDER.bas').read_text(encoding='ascii')
    demo = (Path(root) / 'demo/M_EXCEL_UI_DEMO.bas').read_text(encoding='ascii')
    errors = reset_findings(source) + fast_mode_findings(source, demo)
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
    return errors, count
