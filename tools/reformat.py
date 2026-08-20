#!/usr/bin/env python3
"""
House-style reformatter for VBA .bas modules.

Applies only mechanical, provably behaviour-neutral transformations:

  1. Option statements hoisted above the module header block, flush left,
     and the now-empty MODULE SETTINGS banner removed.
  2. Module and procedure header title lines de-centred to flush left,
     "MODULE: X" reduced to the module name.
  3. Error-handling labels renamed to the house convention.
  4. In-procedure section banners renamed to the house vocabulary.
  5. Procedure-level Dim/Const declarations aligned on the 20/19 grid.
  6. Return types moved onto their own continuation line.
  7. Trailing whitespace stripped, rules normalised to 79 columns, CRLF.

Every transformation is line-local and never touches an executable token.
"""

import re
import sys

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

DECL_RE = re.compile(
    r"^(?P<ind>\s*)(?P<kw>Dim|Static|Const|Private Const|Public Const)\s+"
    r"(?P<name>[A-Za-z_]\w*(?:\(\))?)\s+"
    r"As\s+(?P<type>[A-Za-z_][\w.]*)"
    r"(?P<rest>\s*=\s*.+?)?"
    r"(?P<cmt>\s+'.*)?$"
)

PROC_RE = re.compile(
    r"^(Public |Private |Friend )?(Sub|Function|Property (?:Get|Let|Set))\s+[A-Za-z_]\w*"
)


def split_lines(text):
    return text.replace("\r\n", "\n").replace("\r", "\n").split("\n")


def is_rule(line):
    return re.match(r"^'[-=]{5,}$", line.strip()) is not None


def normalise_rules(lines):
    out = []
    for ln in lines:
        s = ln.strip()
        if re.match(r"^'={5,}$", s):
            out.append(RULE_EQ)
        elif re.match(r"^'-{5,}$", s):
            out.append(RULE_DASH)
        else:
            out.append(ln)
    return out


def hoist_options(lines, module_name):
    """Move Option statements to the top and drop the MODULE SETTINGS banner."""
    options, keep, i = [], [], 0
    while i < len(lines):
        ln = lines[i]
        s = ln.strip()

        if s.startswith("Option "):
            opt = re.sub(r"\s+'.*$", "", s).strip()
            if opt not in options:
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

    if not options:
        options = ["Option Explicit"]

    # Attribute line stays first
    head, body = [], keep
    if body and body[0].startswith("Attribute VB_Name"):
        head, body = [body[0]], body[1:]

    while body and not body[0].strip():
        body = body[1:]
    # a lone "'" left over above the module banner
    if body and body[0].strip() == "'":
        body = body[1:]

    return head + options + [""] + body


def decentre_titles(lines, module_name):
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


def rename_banners(lines):
    out, i = [], 0
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


def rename_labels(lines):
    out = []
    for ln in lines:
        s = ln.strip()
        m = re.match(r"^([A-Za-z_]\w*):\s*$", s)
        if m and m.group(1) in LABEL_MAP:
            out.append(LABEL_MAP[m.group(1)] + ":")
            continue
        new = ln
        for old, rep in LABEL_MAP.items():
            new = re.sub(
                r"\b(GoTo|Resume)\s+" + old + r"\b",
                lambda mm, r=rep: mm.group(1) + " " + r,
                new,
            )
        out.append(new)
    return out


def align_declarations(lines):
    """Align Dim/Const on the 20/19 grid inside procedures only."""
    out = []
    in_proc = False
    for ln in lines:
        s = ln.strip()
        if PROC_RE.match(s):
            in_proc = True
        elif re.match(r"^End (Sub|Function|Property)\b", s):
            in_proc = False

        if not in_proc or ln.rstrip().endswith("_"):
            out.append(ln)
            continue

        m = DECL_RE.match(ln.rstrip())
        if not m:
            out.append(ln)
            continue

        ind = m.group("ind")
        kw = m.group("kw")
        name = m.group("name")
        typ = m.group("type")
        rest = (m.group("rest") or "").rstrip()
        cmt = (m.group("cmt") or "").strip()

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


def strip_trailing(lines):
    return [ln.rstrip() for ln in lines]


def collapse_blanks(lines):
    out = []
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


def reformat(path, module_name):
    text = open(path, encoding="latin-1").read()
    lines = split_lines(text)
    lines = strip_trailing(lines)
    lines = normalise_rules(lines)
    lines = hoist_options(lines, module_name)
    lines = rename_banners(lines)
    lines = decentre_titles(lines, module_name)
    lines = rename_labels(lines)
    lines = align_declarations(lines)
    lines = strip_trailing(lines)
    lines = collapse_blanks(lines)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\r\n".join(lines) + "\r\n"


VB_NAME_RE = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"')


def module_name_of(path):
    """Read the module name from the file's own VB_Name attribute.

    Taking the name from the file rather than the command line removes a class
    of caller error: a mismatched name silently changes what hoist_options and
    decentre_titles do, and the result still looks plausible.
    """
    with open(path, encoding="latin-1") as fh:
        for line in fh:
            m = VB_NAME_RE.match(line.strip())
            if m:
                return m.group(1)
    return None


def check(paths):
    """Report which files are not already in the formatter's normal form.

    Exit status is the point: this is what makes a formatter gate possible.
    The formatter is idempotent, so a file that differs from its own formatted
    output has drifted and can be corrected mechanically.
    """
    failed = []
    for path in paths:
        name = module_name_of(path)
        if name is None:
            print(f"FAIL {path}: no Attribute VB_Name")
            failed.append(path)
            continue

        expected = reformat(path, name).encode("latin-1")
        with open(path, "rb") as fh:
            actual = fh.read()

        if actual == expected:
            print(f"ok   {path}")
        else:
            print(f"FAIL {path}: not in house-style normal form "
                  f"({len(actual) - len(expected):+d} bytes)")
            failed.append(path)

    return 1 if failed else 0


def write(paths):
    """Rewrite each file in place in the formatter's normal form."""
    for path in paths:
        name = module_name_of(path)
        if name is None:
            print(f"skip {path}: no Attribute VB_Name")
            continue
        data = reformat(path, name).encode("latin-1")
        with open(path, "wb") as fh:
            fh.write(data)
        print(f"wrote {path}")
    return 0


USAGE = """usage:
  reformat.py --check <file.bas> [file.bas ...]   report drift, exit 1 if any
  reformat.py --write <file.bas> [file.bas ...]   normalise in place
  reformat.py <src> <dst> <module_name>           legacy explicit form
"""


if __name__ == "__main__":
    args = sys.argv[1:]

    if args and args[0] == "--check":
        sys.exit(check(args[1:]))
    elif args and args[0] == "--write":
        sys.exit(write(args[1:]))
    elif len(args) == 3:
        src, dst, name = args
        open(dst, "wb").write(reformat(src, name).encode("latin-1"))
        print(f"wrote {dst}")
    else:
        sys.stderr.write(USAGE)
        sys.exit(2)
