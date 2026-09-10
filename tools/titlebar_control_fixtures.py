"""Narrow #66 source regression; not an Excel or native-call timeout test."""
from pathlib import Path
import re

from vba_analyze import CONFIGS, active_lines, statements
from vba_lex import code_only

CASE = 'TST_Case_TitleBarShowRejectsCaptionlessBaseline'


def findings(source):
    errors = []
    for config_name, config in CONFIGS.items():
        inside = False
        cleanup = False
        mode = 'handler'
        raises = 0
        for line, raw in active_lines(source, config):
            for part in statements(raw):
                code = code_only(part).strip().lower()
                if re.match(r'private sub ' + CASE.lower() + r'\b', code):
                    inside = True
                if not inside:
                    continue
                if code == 'safe_exit:':
                    cleanup = True
                if code == 'err_handler:':
                    cleanup = False
                if cleanup:
                    if code == 'on error goto 0':
                        mode = 'disabled'
                    elif code.startswith('on error '):
                        mode = 'enabled-or-suppressed'
                    if code.startswith('err.raise '):
                        raises += 1
                        if mode != 'disabled':
                            errors.append(f'{config_name}:{line}: cleanup re-raise must disable its handler')
                if code == 'end sub':
                    inside = False
        if raises != 1:
            errors.append(f'{config_name}: expected one saved-error re-raise in {CASE}')
    return errors


def selftest(root):
    source = f'''Private Sub {CASE}()
On Error GoTo Err_Handler
Safe_Exit:
On Error GoTo 0
If FailNumber <> 0 Then
Err.Raise FailNumber, FailSource, FailDescription
End If
Exit Sub
Err_Handler:
Resume Safe_Exit
End Sub
'''
    cases = [
        ('disarmed', source, False),
        ('original loop', source.replace('On Error GoTo 0\n', ''), True),
        ('swallowed failure', source.replace('On Error GoTo 0', 'On Error Resume Next'), True),
        ('re-enabled handler', source.replace('On Error GoTo 0', 'On Error GoTo 0\nOn Error GoTo Err_Handler'), True),
        ('missing propagation', source.replace('Err.Raise FailNumber, FailSource, FailDescription', 'Debug.Print FailNumber'), True),
    ]
    errors = []
    for name, text, expected_failure in cases:
        if bool(findings(text)) != expected_failure:
            errors.append(f'#66 fixture {name}: unexpected result')
    path = Path(root) / 'test/M_EXCEL_UI_REGRESSION_TESTS.bas'
    errors.extend(findings(path.read_text(encoding='ascii')))
    return errors
