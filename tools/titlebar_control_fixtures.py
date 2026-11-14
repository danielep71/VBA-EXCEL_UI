"""Narrow #66 source regressions; not an Excel or native-call timeout test.

findings() holds the captionless-baseline case to one saved-error re-raise
with its handler disabled. discarded_restores() requires every call of the
entry-style restoration helper to use the result it returns, so a failed
restoration cannot be dropped on a cleanup path.
"""
import re
from pathlib import Path

from vba_analyze import CONFIGS, active_lines, statements
from vba_lex import code_only

CASE = 'TST_Case_TitleBarShowRejectsCaptionlessBaseline'
RESTORE = 'TST_RestoreTitleBarStyle'


def findings(source: str, case: str = CASE, cleanup_label: str = 'safe_exit',
             handler_label: str = 'err_handler') -> list[str]:
    errors: list[str] = []
    for config_name, config in CONFIGS.items():
        inside = False
        cleanup = False
        mode = 'handler'
        raises = 0
        for line, raw in active_lines(source, config):
            for part in statements(raw):
                code = code_only(part).strip().lower()
                if re.match(r'(?:private|public) sub ' + re.escape(case.lower()) + r'\b', code):
                    inside = True
                if not inside:
                    continue
                if code == cleanup_label + ':':
                    cleanup = True
                if code == handler_label + ':':
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
            errors.append(f'{config_name}: expected one saved-error re-raise in {case}')
    return errors


def discarded_restores(source: str, helper: str = RESTORE) -> list[str]:
    """Report each statement that calls the restoration helper and drops its result.

    Assigning the result, testing it in an If and assigning the function's own
    return value inside its body all start with something else, or with the
    name followed by '=', so only a bare call or a Call statement matches.
    """
    errors: list[str] = []
    bare = re.compile(r'(?:call\s+)?' + re.escape(helper.lower()) + r'\b(?!\s*=)')
    for config_name, config in CONFIGS.items():
        for line, raw in active_lines(source, config):
            for part in statements(raw):
                if bare.match(code_only(part).strip().lower()):
                    errors.append(f'{config_name}:{line}: {helper} result discarded; '
                                  'check it and report a failed restoration')
    return errors


def selftest(root: str | Path) -> list[str]:
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
    restore_cases = [
        ('assigned result', f'RestoreOK = {RESTORE}(h, s, c, m)\n', False),
        ('tested result', f'If Not {RESTORE}( _\n    h, s, c, m) Then\nEnd If\n', False),
        ('own return value', f'{RESTORE} = True\n', False),
        ('bare call', f'{RESTORE} h, s, c, m\n', True),
        ('continued bare call', f'{RESTORE} _\n    h, s, c, m\n', True),
        ('Call statement', f'Call {RESTORE}(h, s, c, m)\n', True),
        ('call after a colon', f'x = 1: {RESTORE} h, s, c, m\n', True),
    ]
    for name, text, expected_failure in restore_cases:
        if bool(discarded_restores(text)) != expected_failure:
            errors.append(f'#66 restore fixture {name}: unexpected result')
    path = Path(root) / 'test/M_EXCEL_UI_REGRESSION_TESTS.bas'
    source = path.read_text(encoding='ascii')
    errors.extend(findings(source))
    errors.extend(discarded_restores(source))
    return errors
