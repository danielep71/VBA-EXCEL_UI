"""#71 source contracts and negative fixtures, not Excel execution evidence."""
from pathlib import Path
import re

from titlebar_control_fixtures import findings as propagation_findings
from vba_analyze import logical_lines
from vba_lex import code_only

PROCEDURES = ('DEMO_Sheet_BuildTemplate', 'DEMO_Sheet_Reset')
ALLOW_OPTIONS = ('AllowFormattingCells', 'AllowFormattingColumns',
                 'AllowFormattingRows', 'AllowInsertingColumns',
                 'AllowInsertingRows', 'AllowInsertingHyperlinks',
                 'AllowDeletingColumns', 'AllowDeletingRows', 'AllowSorting',
                 'AllowFiltering', 'AllowUsingPivotTables')


def reset_findings(source):
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
                              ('ws.enableselection = entryprotection.enableselection', 'selectionerr')]:
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
    if 'ws.enableselection = entryprotection.enableselection' not in cleanup:
        errors.append('selection policy restoration missing')
    return errors


def selftest(root):
    source = (Path(root) / 'demo/M_DEMO_BUILDER.bas').read_text(encoding='ascii')
    errors = reset_findings(source)
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
        ('WS.EnableSelection = EntryProtection.EnableSelection', 'Debug.Print 0'),
    ] + [(f'{name}:=EntryProtection.{name}', f'{name}:=False') for name in ALLOW_OPTIONS]
    mutations += [(f'{prefix}{field} = Err.{field}', f'{prefix}{field} = 0')
                  for prefix in ('CleanupErr', 'SelectionErr')
                  for field in ('Number', 'Source', 'Description')]
    for old, new in mutations:
        count += 1
        if old not in source or not reset_findings(source.replace(old, new)):
            errors.append(f'negative source fixture not detected: {old}')
    return errors, count
