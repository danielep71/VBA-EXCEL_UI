#!/usr/bin/env python3
"""
House-style reformatter for VBA .bas modules.

Two normal forms exist. ``house`` is the style of docs/VBA_HOUSE_STYLE.md and is
built by vba_house.py. ``legacy`` is the previous normal form, described below,
which the tree keeps while house_style.HOUSE_STYLE_ACTIVE is False. ``--style``
picks one; without it the switch decides, so ``--check`` accepts the tree that
the switch names. ``--equivalent`` proves that two trees differ only as the
style allows (house_equiv.py), and ``--write-pending`` writes the list of
findings still to be settled by hand (tools/house_style_pending.txt).

The legacy form applies a bounded set of mechanical formatting transformations.
They are intended to be behaviour-neutral, and the named lexical rules below are
exercised by ``--selftest``:

  1. Option statements hoisted above the module header block, flush left,
     and the now-empty MODULE SETTINGS banner removed.
  2. Module and procedure header title lines de-centred to flush left,
     "MODULE: X" reduced to the module name.
  3. Error-handling labels renamed to the house convention.
  4. In-procedure section banners renamed to the house vocabulary.
  5. Procedure-level Dim/Const declarations aligned on the 20/19 grid.
  6. Trailing whitespace stripped, excess blank lines collapsed, rules
     normalised to 79 columns, and line endings normalised to CRLF.

The pipeline includes module-wide operations such as Option hoisting; it is not
line-local. No transformation may alter text inside a string literal: a literal
is data the module evaluates at run time, so rewriting one changes behaviour
rather than layout. Two rules follow, and ``--selftest`` enforces both rather
than leaving them to the reader:

  * text inside a literal is never substituted, so a label name quoted in a
    diagnostic survives a label rename;
  * a comment begins at the first apostrophe OUTSIDE a literal, so an
    apostrophe within quoted text is data and not a comment marker.

The shared vba_lex scanner protects opaque regions. Unsupported continued
statements and conditional blocks are preserved and reported. ASCII-only input
matches the repository gate. A token/statement comparison rejects executable
changes before any write. This is bounded lexical evidence, not VBA compilation.
"""

import os
import re
import stat
import sys
import tempfile
from collections.abc import Callable, Sequence
from pathlib import Path

import house_equiv
import house_style
import vba_house
import vba_house_rules
from vba_lex import LexicalError, code_only, executable_signature, protected_lines, split_code_comment, transform_code

STYLES = ("legacy", "house")


def default_style() -> str:
    """The normal form the tree is held to: the style when it is on, else the previous one."""
    return "house" if house_style.HOUSE_STYLE_ACTIVE else "legacy"

RULE_EQ = "'" + "=" * 78
RULE_DASH = "'" + "-" * 78

LABEL_MAP = {
    "SafeExit": "Safe_Exit",
    "Fail": "Err_Handler",
    "CleanExit": "Clean_Exit",
    "CleanFail": "Clean_Fail",
}

BANNER_MAP = {
    "SAFE EXIT": "RETURN SUCCESS",
    "FAIL": "ERROR HANDLER",
    "CLEAN EXIT": "RETURN SUCCESS",
    "CLEAN FAIL": "ERROR HANDLER",
    "MODULE SETTINGS": None,          # dropped; Options move to the top
    "DECLARE PRIVATE CONSTANTS": "PRIVATE CONSTANTS",
    "DECLARE: PRIVATE CONSTANTS": "PRIVATE CONSTANTS",
    "DECLARE: PRIVATE MODULE STATE": "PRIVATE MODULE STATE",
    "DECLARE: PUBLIC ENUMS": "PUBLIC ENUMS",
    "DECLARE: PRIVATE TYPES": "PRIVATE TYPES",
    "DECLARE: WIN32 / WIN64 API": "WIN32 / WIN64 API DECLARATIONS",
}

# split_code_comment removes the comment before this is applied, so there is
# deliberately no comment group here. There was one, and it treated the first
# apostrophe on the line as the start of a comment wherever it sat, including
# inside a quoted string, where the alignment padding was then written into
# the literal itself.
DECL_RE = re.compile(
    r"^(?P<ind>\s*)(?P<kw>Dim|Static|Const|Private Const|Public Const)\s+"
    r"(?P<name>[A-Za-z_]\w*(?:\(\))?)\s+"
    r"As\s+(?P<type>[A-Za-z_][\w.]*)"
    r"(?P<rest>\s*=\s*.+)?$"
)

PROC_RE = re.compile(
    r"^(Public |Private |Friend )?(Sub|Function|Property (?:Get|Let|Set))\s+[A-Za-z_]\w*"
)


def split_lines(text: str) -> list[str]:
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def sub_outside_literals(code: str, fn: Callable[[str], str]) -> str:
    """Compatibility wrapper around the shared region scanner."""
    return transform_code(code, fn)


def is_rule(line: str) -> bool:
    return re.match(r"^'[-=]{5,}$", line.strip()) is not None


def normalise_rules(lines: list[str]) -> list[str]:
    out: list[str] = []
    for ln in lines:
        s = ln.strip()
        if re.match(r"^'={5,}$", s):
            out.append(RULE_EQ)
        elif re.match(r"^'-{5,}$", s):
            out.append(RULE_DASH)
        else:
            out.append(ln)
    return out


def hoist_options(lines: list[str], module_name: str) -> list[str]:
    """Move Option statements to the top and drop the MODULE SETTINGS banner."""
    options: list[str] = []
    keep: list[str] = []
    i = 0
    while i < len(lines):
        ln = lines[i]
        s = ln.strip()

        if s.startswith("Option "):
            opt = split_code_comment(s)[0].strip()
            options.append(opt)
            i += 1
            continue

        # drop a MODULE SETTINGS banner triple
        if (
            is_rule(ln)
            and i + 2 < len(lines)
            and lines[i + 1].strip().lstrip("'").strip().upper() == "MODULE SETTINGS"
            and is_rule(lines[i + 2])
        ):
            i += 3
            while i < len(lines) and not lines[i].strip():
                i += 1
            continue

        keep.append(ln)
        i += 1


    # Attribute line stays first
    head: list[str] = []
    body = keep
    if body and body[0].startswith("Attribute VB_Name"):
        head, body = [body[0]], body[1:]

    while body and not body[0].strip():
        body = body[1:]
    # a lone "'" left over above the module banner
    if body and body[0].strip() == "'":
        body = body[1:]

    return head + options + [""] + body


def decentre_titles(lines: list[str], module_name: str) -> list[str]:
    """Flush-left the title line that sits between a '=== rule and a '--- rule."""
    out = list(lines)
    for i in range(1, len(out) - 1):
        if not (is_rule(out[i - 1]) and is_rule(out[i + 1])):
            continue
        if not out[i - 1].strip().startswith("'="):
            continue
        if not out[i + 1].strip().startswith("'-"):
            continue
        body = out[i].lstrip()
        if not body.startswith("'"):
            continue
        title = body[1:].strip()
        if not title:
            continue
        title = re.sub(r"^MODULE:\s*", "", title)
        if title.upper() in ("EXCEL_UI_DEMO", "DEMO_BUILDER", "EXCEL_UI_REGRESSION_TESTS"):
            title = module_name
        out[i] = "' " + title
    return out


def rename_banners(lines: list[str]) -> list[str]:
    out: list[str] = []
    i = 0
    while i < len(lines):
        ln = lines[i]
        if (
            is_rule(ln)
            and i + 2 < len(lines)
            and is_rule(lines[i + 2])
            and lines[i + 1].strip().startswith("'")
        ):
            name = lines[i + 1].strip().lstrip("'").strip()
            key = name.upper()
            if key in BANNER_MAP:
                new = BANNER_MAP[key]
                if new is None:
                    i += 3
                    continue
                out.extend([ln, "' " + new, lines[i + 2]])
                i += 3
                continue
        out.append(ln)
        i += 1
    return out


def rename_label_text(text: str) -> str:
    """Rewrite GoTo/Resume targets in one run of code or comment prose."""
    for old, rep in LABEL_MAP.items():
        def renamed(mm: re.Match[str], r: str = rep) -> str:
            return mm.group(1) + " " + r

        text = re.sub(r"\b(GoTo|Resume)\s+" + old + r"\b", renamed, text)
    return text


def rename_labels(lines: list[str]) -> list[str]:
    """Rename error-handling labels, in code and in the prose describing it.

    Not in string literals. A module quoting a label in a diagnostic - "use
    GoTo Fail nowhere" - had the quoted text rewritten along with the jump,
    which changed what the module printed at run time rather than where it
    jumped.
    """
    out: list[str] = []
    for ln in lines:
        s = ln.strip()
        m = re.match(r"^([A-Za-z_]\w*):", s)
        if m and m.group(1) in LABEL_MAP:
            ln = ln.replace(m.group(1) + ":", LABEL_MAP[m.group(1)] + ":", 1)

        code, comment = split_code_comment(ln)
        out.append(
            sub_outside_literals(code, rename_label_text)
            + (comment if comment.lower().startswith("rem") else rename_label_text(comment))
        )
    return out


def align_declarations(lines: list[str]) -> list[str]:
    """Align Dim/Const on the 20/19 grid inside procedures only."""
    out: list[str] = []
    in_proc = False
    for ln in lines:
        s = ln.strip()
        if PROC_RE.match(s):
            in_proc = True
        elif re.match(r"^End (Sub|Function|Property)\b", s):
            in_proc = False

        code, comment = split_code_comment(ln)

        if not in_proc or code.rstrip().endswith("_"):
            out.append(ln)
            continue

        m = DECL_RE.match(code.rstrip())
        if not m:
            out.append(ln)
            continue

        ind = m.group("ind")
        kw = m.group("kw")
        name = m.group("name")
        typ = m.group("type")
        rest = (m.group("rest") or "").rstrip()
        cmt = comment.strip()

        left = f"{kw} {name}"
        if kw == "Dim":
            width = len("Dim ") + 20
            # A name wider than the field must still keep a separating space,
            # otherwise "Dim LongName" + "As Long" fuses into one token.
            left = left.ljust(width) if len(left) < width else left + " "
        else:
            left = left + " "

        mid = f"As {typ}"
        if rest:
            mid = mid + " " + rest.strip()

        if cmt:
            mid = mid.ljust(19) if len(mid) < 19 else mid + " "
            out.append((ind + left + mid + cmt).rstrip())
        else:
            out.append((ind + left + mid).rstrip())
    return out


def strip_trailing(lines: list[str]) -> list[str]:
    return [ln.rstrip() for ln in lines]


def collapse_blanks(lines: list[str]) -> list[str]:
    out: list[str] = []
    blanks = 0
    for ln in lines:
        if not ln.strip():
            blanks += 1
            if blanks > 2:
                continue
        else:
            blanks = 0
        out.append(ln)
    return out


def reformat(path: str | Path, module_name: str, style: str | None = None) -> str:
    with open(path, encoding="latin-1") as fh:
        return reformat_text(fh.read(), module_name, style)


def reformat_text(text: str, module_name: str, style: str | None = None) -> str:
    """Validate, transform supported regions and verify executable equivalence."""
    return reformat_with_todos(text, module_name, style)[0]


def reformat_with_todos(text: str, module_name: str, style: str | None = None) -> tuple[str, list[tuple[str, str]]]:
    """The reformatted text and, for the house style, what a person still has to settle."""
    chosen = style or default_style()
    if chosen not in STYLES:
        raise ValueError(f"unknown style {chosen!r}")
    if chosen == "house":
        return reformat_house(text, module_name)
    return reformat_legacy(text, module_name), []


def reformat_house(text: str, module_name: str) -> tuple[str, list[tuple[str, str]]]:
    """Rewrite into the style, then prove the executable tokens are the same.

    Only the renames declared in house_style may differ: error-handling labels
    in a label or jump position, and PROC inside the procedures that declare it.
    """
    result, todos = vba_house.format_module(text, module_name)
    if executable_signature(vba_house.renamed_source(text, module_name)) != executable_signature(result):
        raise LexicalError("formatting would change executable tokens or statement boundaries")
    return result, todos


def reformat_legacy(text: str, module_name: str) -> str:
    """Validate, transform supported regions and verify executable equivalence."""
    if not text.isascii():
        raise LexicalError("non-ASCII VBA input; repository exports must be ASCII")
    lines = split_lines(text)
    protected, _ = protected_lines(lines)
    # Label renaming is module-wide. Refuse ambiguous old/new collisions or
    # jumps across opaque blocks rather than partly renaming a control flow.
    code = "\n".join(code_only(ln) for ln in lines)
    for old, new in LABEL_MAP.items():
        if re.search(r"\b" + old + r"\b", code) and re.search(r"\b" + new + r"\b", code):
            raise LexicalError(f"ambiguous label mapping {old}/{new}")
        if any(re.search(r"\b" + old + r"\b", code_only(lines[i])) for i in protected):
            raise LexicalError(f"unsupported label {old} in preserved region")
    # Opaque lines remain in place through the structural passes. Unique comment
    # placeholders also prevent declaration alignment on continuation tails.
    saved: dict[str, str] = {}
    for i in protected:
        key = f"' @formatter-preserve-{i}@"
        if key in text:
            raise LexicalError("reserved formatter marker in source")
        saved[key] = lines[i]
        lines[i] = key
    lines = strip_trailing(lines)
    lines = normalise_rules(lines)
    lines = hoist_options(lines, module_name)
    lines = rename_banners(lines)
    lines = decentre_titles(lines, module_name)
    lines = rename_labels(lines)
    lines = align_declarations(lines)
    lines = collapse_blanks(strip_trailing(lines))
    lines = [saved.get(line, line) for line in lines]
    while lines and not lines[-1].strip():
        lines.pop()
    result = "\r\n".join(lines) + "\r\n"
    if executable_signature(text, LABEL_MAP) != executable_signature(result, LABEL_MAP):
        raise LexicalError("formatting would change executable tokens or statement boundaries")
    return result


def atomic_write(path: str | Path, data: bytes) -> None:
    """Replace one regular file only after a complete, flushed temporary write.

    Same-directory replace is atomic; multi-file writes are not a transaction.
    Symlinks are refused. Existing permission bits are preserved.
    """
    path = Path(path)
    if path.is_symlink():
        raise OSError("refusing to replace a symbolic link")
    mode = stat.S_IMODE(path.stat().st_mode) if path.exists() else 0o644
    fd, name = tempfile.mkstemp(prefix=".reformat-", dir=path.parent)
    try:
        with os.fdopen(fd, "wb") as fh:
            fh.write(data)
            fh.flush()
            os.fsync(fh.fileno())
        os.chmod(name, mode)
        os.replace(name, path)
    finally:
        if os.path.exists(name):
            os.unlink(name)



VB_NAME_RE = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"')


def module_name_of(path: str | Path) -> str | None:
    """Read the module name from the file's own VB_Name attribute.

    Taking the name from the file rather than the command line removes a class
    of caller error: a mismatched name silently changes what hoist_options and
    decentre_titles do, and the result still looks plausible.
    """
    with open(path, encoding="latin-1") as fh:
        for line in fh:
            m = VB_NAME_RE.match(line.strip())
            if m:
                return str(m.group(1))
    return None


def hand_edits(path: str | Path, name: str, style: str) -> str:
    """One line naming the rules of the style that rewriting cannot settle."""
    if style != "house":
        return ""
    found = vba_house_rules.findings(Path(path).read_text(encoding="ascii"), name)
    counts = vba_house_rules.summarise(found)
    return ", ".join(f"{rule} {count}" for rule, count in counts.items() if count)


def write_pending(paths: Sequence[str | Path], target: str | Path | None = None) -> int:
    """Write the list of findings that are still to be settled by hand (tools/house_style_pending.txt).

    The list is what the house-style rules check allows while the modules are
    brought to the style. Once it exists this only removes lines: a finding
    that is not already listed is refused, so the list cannot grow to hide new
    code. With no findings left, the file is removed.
    """
    path = Path(target) if target else Path(__file__).resolve().parent.parent / vba_house_rules.PENDING_FILE
    found: list[tuple[str, vba_house_rules.Finding]] = []
    try:
        for source in paths:
            name = module_name_of(source)
            if name is None:
                raise LexicalError(f"{source}: no Attribute VB_Name")
            text = Path(source).read_text(encoding="ascii")
            found.extend((name, item) for item in vba_house_rules.findings(text, name))
        existing = path.read_text(encoding="ascii") if path.exists() else None
        if existing is not None:
            added = vba_house_rules.additions(found, existing)
            if added:
                print(f"FAIL {path}: refusing to add {len(added)} line(s); settle the finding or list it by hand: "
                      + "; ".join(added[:5]) + (" ..." if len(added) > 5 else ""))
                return 1
        if not found:
            if existing is not None:
                path.unlink()
            print(f"ok   no finding left; {path} removed" if existing is not None else "ok   no finding")
            return 0
        atomic_write(path, vba_house_rules.pending_text(found).encode("ascii"))
    except (ValueError, OSError) as exc:
        print(f"FAIL {exc}")
        return 1
    print(f"ok   {path}: {len(found)} finding(s) listed")
    return 0


def check(paths: Sequence[str | Path], style: str | None = None) -> int:
    """Report which files are not already in the formatter's normal form.

    Exit status is the point: this is what makes a formatter gate possible.
    Normal form is required to be idempotent, and the self-test verifies that
    property for every named fixture. A file that differs from its formatted
    output has drifted from the normal form this tool currently defines.
    """
    failed: list[str | Path] = []
    for path in paths:
        name = module_name_of(path)
        if name is None:
            print(f"FAIL {path}: no Attribute VB_Name")
            failed.append(path)
            continue

        try:
            expected = reformat(path, name, style).encode("ascii")
            _, notes = protected_lines(split_lines(Path(path).read_text(encoding="ascii")))
            if notes and (style or default_style()) == "legacy":
                print(f"note {path}: {len(notes)} conditional/continued lines preserved; "
                      f"unsupported formatting skipped")
        except (ValueError, OSError) as exc:
            print(f"FAIL {path}: {exc}")
            failed.append(path)
            continue
        with open(path, "rb") as fh:
            actual = fh.read()

        if actual == expected:
            print(f"ok   {path}")
            pending = hand_edits(path, name, style or default_style())
            if pending:
                print(f"note {path}: for a person to settle: {pending}")
        else:
            print(f"FAIL {path}: not in house-style normal form "
                  f"({len(actual) - len(expected):+d} bytes)")
            failed.append(path)

    return 1 if failed else 0


def write(paths: Sequence[str | Path], style: str | None = None) -> int:
    """Preflight every input, then atomically replace each changed file."""
    prepared: list[tuple[str | Path, bytes]] = []
    try:
        for path in paths:
            name = module_name_of(path)
            if name is None:
                raise LexicalError(f"{path}: no Attribute VB_Name")
            data = reformat(path, name, style).encode("ascii")
            prepared.append((path, data))
        for path, data in prepared:
            if Path(path).read_bytes() != data:
                atomic_write(path, data)
            print(f"ok   {path}")
            pending = hand_edits(path, module_name_of(path) or "", style or default_style())
            if pending:
                print(f"note {path}: for a person to settle: {pending}")
    except (ValueError, OSError) as exc:
        print(f"FAIL {exc}")
        return 1
    return 0


# --------------------------------------------------------------------------
# Self-test
#
# A defect in this file cannot be caught by the VBA regression suite, and
# --check passing proves only that today's modules happen not to contain a
# construct that trips it. Both defects corrected here were latent for exactly
# that reason: no module in the repository quoted a label name or put an
# apostrophe inside a literal, so the gate was green while the transformation
# was wrong.
#
# Each fixture is a module fragment with a single line under test. The
# expectation is stated as the line the formatter must produce, so a failure
# names the rule rather than a byte count.

SELFTEST_CASES = [
    (
        "a label name quoted in a literal is data, not a jump",
        '        Msg = "use GoTo Fail nowhere"',
        '        Msg = "use GoTo Fail nowhere"',
    ),
    (
        "a real jump is still renamed",
        "        On Error GoTo Fail",
        "        On Error GoTo Err_Handler",
    ),
    (
        "one line, one jump renamed and one literal left alone",
        '        If X Then Msg = "GoTo Fail" Else GoTo Fail',
        '        If X Then Msg = "GoTo Fail" Else GoTo Err_Handler',
    ),
    (
        "comment prose is renamed, because it describes the jump",
        "        On Error GoTo Fail  \'then Resume Fail",
        "        On Error GoTo Err_Handler  \'then Resume Err_Handler",
    ),
    (
        "an apostrophe inside a literal does not start a comment",
        '    Const S As String = "a \'b\' c"',
        '    Const S As String = "a \'b\' c"',
    ),
    (
        "a trailing comment after such a literal is still aligned",
        '    Const S As String = "a \'b\'" \'note',
        '    Const S As String = "a \'b\'" \'note',
    ),
    (
        "a doubled quote does not unbalance the literal scan",
        '    Const T As String = "he said ""go"" \'x\' now"',
        '    Const T As String = "he said ""go"" \'x\' now"',
    ),
    (
        "an ordinary declaration is still aligned on the 20/19 grid",
        "    Dim Msg As String \'Diagnostic buffer",
        "    Dim Msg                 As String          \'Diagnostic buffer",
    ),
    (
        "a label definition is still renamed",
        "Fail:",
        "Err_Handler:",
    ),
]

SELFTEST_HEAD = [
    'Attribute VB_Name = "M_SELFTEST"',
    "Option Explicit",
    "",
    "Private Sub SelfTest_Probe()",
]

SELFTEST_TAIL = ["End Sub"]


def selftest() -> list[str]:
    """Return a list of failure descriptions; empty means every rule holds."""
    failures: list[str] = []

    for label, before, after in SELFTEST_CASES:
        source = "\r\n".join(SELFTEST_HEAD + [before] + SELFTEST_TAIL) + "\r\n"
        produced = reformat_text(source, "M_SELFTEST", "legacy").split("\r\n")

        if after not in produced:
            got = [ln for ln in produced
                   if ln not in SELFTEST_HEAD + SELFTEST_TAIL and ln.strip()]
            failures.append(
                f"{label}\n      expected: {after!r}\n      produced: {got!r}"
            )
            continue

        # Idempotence is part of the contract check() relies on: a file that
        # differs from its own formatted output is said to have drifted, which
        # is only true while formatting twice equals formatting once.
        again = reformat_text("\r\n".join(produced), "M_SELFTEST", "legacy")
        if again != "\r\n".join(produced):
            failures.append(f"{label}\n      not idempotent on a second pass")

    from house_equiv_fixtures import equiv_selftest
    from house_format_fixtures import house_selftest
    from reformat_fixtures import extended_selftest
    failures.extend(extended_selftest())
    failures.extend(house_selftest())
    failures.extend(equiv_selftest())
    return failures


def run_selftest() -> int:
    failures = selftest()
    if not failures:
        print(f"ok   self-test: {len(SELFTEST_CASES)} historical rules and extended lexical/write fixtures hold")
        return 0
    print(f"FAIL self-test: {len(failures)} rule(s) broken\n")
    for f in failures:
        print(f"  - {f}")
    return 1



USAGE = """usage:
  reformat.py --selftest                         verify the rules themselves
  reformat.py [--style S] --check <file.bas> ...  report drift, exit 1 if any
  reformat.py [--style S] --write <file.bas> ...  normalise in place
  reformat.py --write-pending <file.bas> ...      write tools/house_style_pending.txt (never adds a line)
  reformat.py --equivalent BASE_DIR NEW_DIR       prove two trees differ only as the style allows
  reformat.py [--style S] <src> <dst> <module_name>   explicit form
S is "house" (the style in docs/VBA_HOUSE_STYLE.md) or "legacy" (the previous
normal form). Without --style the switch HOUSE_STYLE_ACTIVE in
tools/house_style.py decides.
"""


def main(argv: list[str]) -> int:
    args = list(argv)
    style: str | None = None
    if "--style" in args:
        at = args.index("--style")
        if at + 1 >= len(args) or args[at + 1] not in STYLES:
            sys.stderr.write(USAGE)
            return 2
        style = args[at + 1]
        del args[at:at + 2]
    if args and args[0] == "--selftest":
        return run_selftest()
    if args and args[0] == "--check":
        return check(args[1:], style)
    if args and args[0] == "--write":
        return write(args[1:], style)
    if args and args[0] == "--write-pending":
        return write_pending(args[1:])
    if args and args[0] == "--equivalent" and len(args) == 3:
        return house_equiv.run(args[1], args[2])
    if len(args) == 3:
        src, dst, name = args
        atomic_write(dst, reformat(src, name, style).encode("ascii"))
        print(f"wrote {dst}")
        return 0
    sys.stderr.write(USAGE)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
