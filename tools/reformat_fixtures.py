"""Independent expectations for #52, exercised by the existing static gate."""
import contextlib
import io
import os
from pathlib import Path
import tempfile
from unittest.mock import patch

from vba_lex import LexicalError, code_only, executable_signature, regions


CASES = [
    ('Rem at start', 'Rem GoTo Fail "unterminated', 'Rem GoTo Fail "unterminated'),
    ('Rem after colon', 'x = 1: Rem GoTo Fail', 'x = 1: Rem GoTo Fail'),
    ('Rem after Then', 'If x Then Rem GoTo Fail', 'If x Then Rem GoTo Fail'),
    ('Rem member', 'x = obj.Rem: GoTo Fail', 'x = obj.Rem: GoTo Err_Handler'),
    ('Rem prefix', 'Remember = "GoTo Fail"', 'Remember = "GoTo Fail"'),
    ('inline label', 'Fail: x = "Resume Fail": Resume Fail', 'Err_Handler: x = "Resume Fail": Resume Err_Handler'),
    ('trailing label comment', "Fail: 'diagnostic", "Err_Handler: 'diagnostic"),
    ('doubled quotes', 'x = "a ""GoTo Fail"" b": GoTo Fail', 'x = "a ""GoTo Fail"" b": GoTo Err_Handler'),
    ('apostrophe data', 'Const X As String = "a\'b"', 'Const X As String = "a\'b"'),
    ('wide declaration', 'Dim ANameThatExceedsTwentyCharacters As String', 'Dim ANameThatExceedsTwentyCharacters As String'),
    ('bracketed name', '[Fail] = 2', '[Fail] = 2'),
]


def extended_selftest():
    import reformat as f
    failures = []

    def require(condition, name):
        if not condition:
            failures.append('extended: ' + name)

    def module(fragment):
        return 'Attribute VB_Name = "M_PROBE"\r\nOption Explicit\r\nPrivate Sub Probe()\r\n' + fragment + '\r\nEnd Sub\r\n'

    for name, before, expected in CASES:
        try:
            source = module(before)
            out = f.reformat_text(source, 'M_PROBE')
            require(expected in out.split('\r\n'), name)
            require(out == f.reformat_text(out, 'M_PROBE'), name + ' idempotence')
            require(executable_signature(source, f.LABEL_MAP) == executable_signature(out, f.LABEL_MAP), name + ' tokens')
        except Exception as exc:
            failures.append(f'{name}: {exc}')

    for name, fragment in [
        ('continued declaration', 'Dim x As _\r\n    Long'),
        ('continued expression', 'x = "a" & _\r\n    "b"'),
        ('nested conditional', '#If VBA7 Then\r\n  #If Win64 Then\r\n Dim x As LongPtr   \r\n  #Else\r\n Dim x As Long\r\n  #End If\r\n#End If'),
    ]:
        source = module(fragment)
        out = f.reformat_text(source, 'M_PROBE')
        require(fragment in out, name + ' byte preservation')
        require(out == f.reformat_text(out, 'M_PROBE'), name + ' idempotence')

    for name, fragment in [
        ('unclosed quote', 'x = "GoTo Fail'),
        ('unclosed bracket', '[Fail = 2'),
        ('non-ASCII literal', 'x = "caf\u00e9"'),
        ('non-ASCII comment', "'caf\u00e9"),
        ('label collision', 'Fail:\r\nErr_Handler:'),
        ('unterminated conditional', '#If VBA7 Then'),
        ('continued label rename', 'On Error GoTo _\r\n Fail'),
    ]:
        try:
            f.reformat_text(module(fragment), 'M_PROBE')
            require(False, name + ' must reject')
        except LexicalError:
            pass

    require(''.join(code_only('x = "GoTo Missing": Rem Resume Missing').split()) == 'x=:', 'analyzer masks data/comments')
    require(regions('Rem "')[0][0] == 'comment', 'Rem owns unterminated quote')
    # Comparator negative controls: real executable/literal/options changes
    # must not be normalized away as if they were formatting.
    for before, after in [('x=1', 'x=2'), ('x="a"', 'x="b"'),
                          ('Dim x As Long', 'Dim xAs Long'),
                          ('x=1:y=2', 'x=1\ny=2')]:
        require(executable_signature(before) != executable_signature(after), 'negative control ' + before)
    require(executable_signature('Option Explicit\nx=1') != executable_signature('x=1'), 'option insertion rejected')
    require(executable_signature('#If X Then\nOption Explicit\n#End If') != executable_signature('Option Explicit\n#If X Then\n#End If'), 'conditional option movement rejected')
    with patch.object(f, 'align_declarations', lambda lines: [ln.replace('x = 1', 'x = 2') for ln in lines]):
        try:
            f.reformat_text(module('x = 1'), 'M_PROBE')
            require(False, 'corrupt transformation must be rejected')
        except LexicalError:
            pass

    # Old defect shapes are deliberately applied to their fixtures. These are
    # negative controls, not claims that every earlier release was executed.
    literal = 'x = "GoTo Fail"'
    require(f.rename_label_text(literal) != literal, 'historical literal rewrite control detects corruption')
    declaration = 'Const S As String = "a\'b"'
    require(declaration.split("'", 1)[0] != f.split_code_comment(declaration)[0], 'historical first-apostrophe split control')

    with tempfile.TemporaryDirectory() as directory:
        p = Path(directory) / 'probe.bas'
        original = module('Dim x As Long').encode('ascii')
        p.write_bytes(original)
        os.chmod(p, 0o640)
        with patch.object(f.os, 'replace', side_effect=OSError('injected replace failure')):
            try:
                f.atomic_write(p, b'new')
            except OSError:
                pass
        require(p.read_bytes() == original, 'replace failure preserves original')
        require(list(Path(directory).glob('.reformat-*')) == [], 'replace failure removes temporary')
        with patch.object(f.os, 'fsync', side_effect=OSError('injected flush failure')):
            try:
                f.atomic_write(p, b'new')
            except OSError:
                pass
        require(p.read_bytes() == original, 'flush failure preserves original')
        require(list(Path(directory).glob('.reformat-*')) == [], 'flush failure removes temporary')
        bad = Path(directory) / 'bad.bas'
        bad.write_bytes(module('x = "unfinished').encode('ascii'))
        with contextlib.redirect_stdout(io.StringIO()):
            require(f.write([p, bad]) == 1, 'batch validation fails')
        require(p.read_bytes() == original, 'batch preflight does not modify first file')
        f.atomic_write(p, b'complete')
        require(p.read_bytes() == b'complete', 'successful replace')
        require(p.stat().st_mode & 0o777 == 0o640, 'permission preservation')
    return failures
