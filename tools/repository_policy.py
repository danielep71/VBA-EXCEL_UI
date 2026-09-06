"""Offline policy probes. Git itself evaluates attributes and ignore precedence.

EditorConfig support is deliberately limited to this repository's flat sections,
fnmatch patterns and comma braces; unsupported brace syntax fails closed.
"""
import fnmatch
import os
from pathlib import Path
import re
import subprocess
import tempfile

WINDOWS = ('bat', 'cmd', 'ps1', 'psm1', 'psd1', 'vbs', 'reg', 'ini')
VBA = ('bas', 'cls', 'frm')
PORTABLE = ('md', 'py', 'yml', 'yaml', 'json', 'txt', 'sh')
FILES = ('.editorconfig', '.gitattributes', '.gitignore')


def git(root, *args, data=None):
    env = dict(os.environ, GIT_CONFIG_NOSYSTEM='1', GIT_CONFIG_GLOBAL=os.devnull)
    # Isolate fixture Git operations from caller repository overrides.
    for name in ('GIT_DIR', 'GIT_WORK_TREE', 'GIT_INDEX_FILE', 'GIT_COMMON_DIR'):
        env.pop(name, None)
    return subprocess.run(['git', '-c', 'core.excludesFile=' + os.devnull,
                           '-c', 'core.attributesFile=' + os.devnull, *args],
                          cwd=root, input=data, text=True, capture_output=True,
                          check=True, env=env).stdout


def editor_eol(text, path):
    pattern = None
    result = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith(('#', ';')):
            continue
        if line.startswith('[') and line.endswith(']'):
            pattern = line[1:-1]
        elif '=' in line and pattern:
            key, value = (s.strip().lower() for s in line.split('=', 1))
            if key != 'end_of_line':
                continue
            patterns = [pattern]
            if '{' in pattern:
                match = re.fullmatch(r'([^{}]*)\{([^{}]+)\}([^{}]*)', pattern)
                if not match or ',' not in match[2]:
                    raise ValueError('unsupported EditorConfig brace pattern: ' + pattern)
                patterns = [match[1] + item + match[3] for item in match[2].split(',')]
            if any(fnmatch.fnmatchcase(path if '/' in p else Path(path).name, p)
                   for p in patterns):
                result = value
    return result


def findings(root):
    root = Path(root)
    errors = []
    text = (root / '.editorconfig').read_text(encoding='utf-8')
    paths = [(f'{prefix}probe.{ext}', 'crlf' if ext in WINDOWS + VBA else 'lf')
             for prefix in ('', 'tools/nested/') for ext in WINDOWS + VBA + PORTABLE]
    for path, expected in paths:
        if editor_eol(text, path) != expected:
            errors.append('EditorConfig EOL: ' + path)
    values = git(root, 'check-attr', '-z', '--stdin', 'text', 'eol',
                 data='\0'.join(p for p, _ in paths) + '\0').split('\0')
    attrs = {(values[i], values[i+1]): values[i+2]
             for i in range(0, len(values)-1, 3)}
    for path, expected in paths:
        if attrs.get((path, 'eol')) != expected or attrs.get((path, 'text')) not in ('set', 'auto'):
            errors.append('Git EOL: ' + path)
    scratch = ['.mutation-scratch/probe.bas', '.mutation-scratch/nested/control.json']
    visible = ['src/mutant.bas', 'test/control.bas', 'test/fixtures/mutant.json',
               'docs/evidence/control.md', 'evidence/control.json',
               'test/.mutation-scratch/control.bas']
    output = git(root, 'check-ignore', '--no-index', '-z', '--stdin',
                 data='\0'.join(scratch + visible) + '\0')
    ignored = set(output.rstrip('\0').split('\0'))
    for path in scratch:
        if path not in ignored:
            errors.append('Scratch not ignored: ' + path)
    for path in visible:
        if path in ignored:
            errors.append('Authoritative path ignored: ' + path)
    archive = git(root, 'check-attr', 'export-ignore', '--', '.mutation-scratch')
    if not archive.strip().endswith(': set'):
        errors.append('Scratch archive exclusion missing')
    return errors


def check(root):
    """Evaluate tracked policy in isolation from personal Git settings."""
    with tempfile.TemporaryDirectory(prefix='excel-ui-policy-') as name:
        target = Path(name)
        git(target, 'init', '-q')
        for file in FILES:
            (target / file).write_bytes((Path(root) / file).read_bytes())
        return findings(target)


def selftest(root):
    """Positive baseline plus removed/shadowed rules and broad-ignore mutants."""
    originals = {file: (Path(root) / file).read_text(encoding='utf-8') for file in FILES}
    cases = [('baseline', None, None, False)]
    for ext in WINDOWS:
        cases.append(('removed Git ' + ext, '.gitattributes',
                      originals['.gitattributes'].replace(f'*.{ext} text eol=crlf\n', ''), True))
        cases.append(('shadow Git ' + ext, '.gitattributes',
                      originals['.gitattributes'] + f'\n*.{ext} text eol=lf\n', True))
        cases.append(('shadow editor ' + ext, '.editorconfig',
                      originals['.editorconfig'] + f'\n[*.{ext}]\nend_of_line = lf\n', True))
    cases.append(('removed editor override', '.editorconfig',
                  originals['.editorconfig'].replace('[*.{bat,cmd,ps1,psm1,psd1,vbs,reg,ini}]\nend_of_line = crlf', ''), True))
    cases.append(('removed scratch', '.gitignore',
                  originals['.gitignore'].replace('/.mutation-scratch/\n', ''), True))
    for rule in ('*.bas', '*mutant*', '*control*', '/evidence/', '/docs/evidence/'):
        cases.append(('broad ignore ' + rule, '.gitignore', originals['.gitignore'] + '\n' + rule + '\n', True))
    cases.append(('missing archive rule', '.gitattributes',
                  originals['.gitattributes'].replace('/.mutation-scratch export-ignore', ''), True))
    failures = []
    for label, file, replacement, expect_failure in cases:
        with tempfile.TemporaryDirectory(prefix='excel-ui-policy-test-') as name:
            target = Path(name)
            git(target, 'init', '-q')
            for key, value in originals.items():
                (target / key).write_text(replacement if key == file else value, encoding='utf-8')
            try:
                actual = bool(findings(target))
            except subprocess.CalledProcessError as exc:
                # check-ignore uses exit 1 when none match: missing scratch is failure.
                if exc.returncode != 1 or 'check-ignore' not in exc.cmd:
                    raise
                actual = True
            if actual != expect_failure:
                failures.append(label)
    return failures, len(cases)
