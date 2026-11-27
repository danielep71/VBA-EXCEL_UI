"""Prove that two trees of VBA modules differ only where the house style allows (#210).

``python3 tools/reformat.py --equivalent BASE_DIR NEW_DIR`` reads every ``.bas``
module under both directories and compares the ordered executable tokens of
each procedure, and of the module outside its procedures. Nothing else may
differ except:

* comments, whitespace, blank lines and the way a statement is wrapped over
  lines with `` _``;
* the renames the style declares (error-handling labels, and ``PROC`` as
  ``PROC_NAME`` inside the procedures that declare it), which are applied to
  the base before it is compared;
* ``Const PROC_NAME As String = "<Module>.<Procedure>"``: added where the base
  has none, or moved within its procedure with the same literal;
* an empty ``Case Else`` directly before ``End Select``, which does nothing.

A changed operator, a deleted or inserted statement, a changed literal or a
different order of procedures is a difference. The comparison reads tokens; it
is not VBA compilation and does not replace the Excel certification run.
"""
import re
from collections.abc import Sequence
from pathlib import Path
from typing import NamedTuple

import vba_house
from vba_lex import LexicalError, regions

Token = tuple[str, str]
TOKEN = re.compile(r"[A-Za-z_]\w*|\d+(?:\.\d+)?|[^\s]")
VB_NAME = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"')
SHOWN_TOKENS = 14


class Row(NamedTuple):
    """One logical statement: its tokens and the physical line it starts on."""
    line: int
    tokens: tuple[Token, ...]


class Procedure(NamedTuple):
    kind: str
    name: str
    rows: list[Row]


def line_tokens(text: str) -> list[Token]:
    tokens: list[Token] = []
    for kind, value in regions(text):
        if kind == "comment":
            continue
        if kind == "code":
            tokens.extend(("code", t) for t in TOKEN.findall(value))
        else:
            tokens.append((kind, value))
    return tokens


def rows_of(lines: Sequence[str], first: int, last: int) -> list[Row]:
    """Logical statements of lines[first:last + 1]; a continued statement is one row."""
    rows: list[Row] = []
    pending: list[Token] = []
    start = first
    for n in range(first, last + 1):
        tokens = line_tokens(lines[n])
        if tokens and tokens[-1] == ("code", "_"):
            if not pending:
                start = n
            pending.extend(tokens[:-1])
            continue
        if pending:
            rows.append(Row(start + 1, tuple(pending + tokens)))
            pending = []
        elif tokens:
            rows.append(Row(n + 1, tuple(tokens)))
    if pending:
        raise LexicalError(f"line {start + 1}: a continued statement ends the procedure")
    return rows


def words(tokens: Sequence[Token]) -> list[str]:
    return [value.lower() for kind, value in tokens if kind == "code"]


def is_proc_name(row: Row) -> bool:
    kinds = [kind for kind, _ in row.tokens]
    return (len(row.tokens) == 6 and kinds[5] == "literal" and kinds[:5] == ["code"] * 5
            and words(row.tokens[:5]) == ["const", "proc_name", "as", "string", "="])


def without_empty_case_else(rows: list[Row]) -> list[Row]:
    """The rows without a ``Case Else`` that is followed at once by ``End Select``."""
    keep: list[Row] = []
    for n, row in enumerate(rows):
        following = words(rows[n + 1].tokens) if n + 1 < len(rows) else []
        if words(row.tokens) == ["case", "else"] and following == ["end", "select"]:
            continue
        keep.append(row)
    return keep


def shown(row: Row | None) -> str:
    if row is None:
        return "nothing"
    values = [value for _, value in row.tokens]
    text = " ".join(values[:SHOWN_TOKENS]) + (" ..." if len(values) > SHOWN_TOKENS else "")
    return f"`{text}` (line {row.line})"


def first_difference(base: list[Row], new: list[Row]) -> str | None:
    for n in range(min(len(base), len(new))):
        if base[n].tokens != new[n].tokens:
            tail_base = [r.tokens for r in base[n + 1:]]
            tail_new = [r.tokens for r in new[n + 1:]]
            if tail_base == [r.tokens for r in new[n:]]:
                return f"base has {shown(base[n])}, new has nothing there: a statement was deleted"
            if tail_new == [r.tokens for r in base[n:]]:
                return f"new has {shown(new[n])}, base has nothing there: a statement was inserted"
            return f"base has {shown(base[n])}, new has {shown(new[n])}"
    if len(base) > len(new):
        return f"base has {shown(base[len(new)])}, new has nothing there: a statement was deleted"
    if len(new) > len(base):
        return f"new has {shown(new[len(base)])}, base has nothing there: a statement was inserted"
    return None


def module_of(text: str, fallback: str) -> str:
    for line in text.split("\n"):
        m = VB_NAME.match(line.strip())
        if m:
            return m.group(1)
    return fallback


def split_module(text: str) -> tuple[list[Row], list[Row], list[Procedure]]:
    """(module-level Option rows, other module-level rows, procedures) of one module."""
    raw = vba_house.split_lines(text)
    fixed = vba_house.conditional_lines(raw)
    lines = [vba_house.Ln(t.rstrip(), n in fixed) for n, t in enumerate(raw)]
    spans = vba_house.procedure_spans(lines)
    plain = [t.rstrip() for t in raw]
    procedures: list[Procedure] = []
    outside: list[Row] = []
    at = 0
    for sig, end in spans:
        if at < sig:
            outside.extend(rows_of(plain, at, sig - 1))
        m = vba_house.signature_match(lines, sig)
        if m is None:
            raise LexicalError(f"line {sig + 1}: not a procedure signature")
        procedures.append(Procedure(re.sub(r"\s+", " ", m.group("kind")), m.group("name"), rows_of(plain, sig, end)))
        at = end + 1
    if at < len(plain):
        outside.extend(rows_of(plain, at, len(plain) - 1))
    options = [row for row in outside if words(row.tokens)[:1] == ["option"]]
    return options, [row for row in outside if row not in options], procedures


def compare_procedure(base: Procedure, new: Procedure, module: str) -> str | None:
    """The first difference of one procedure, or None when only allowed differences remain."""
    wanted = f'"{module}.{new.name}"'
    old_const = [row for row in base.rows if is_proc_name(row)]
    new_const = [row for row in new.rows if is_proc_name(row)]
    if len(new_const) > 1:
        return f"PROC_NAME is declared {len(new_const)} times"
    if old_const and [r.tokens for r in old_const] != [r.tokens for r in new_const]:
        if not new_const:
            return f"base declares {shown(old_const[0])}, new does not"
        return f"base declares {shown(old_const[0])}, new declares {shown(new_const[0])}: the literal changed"
    if not old_const and new_const and new_const[0].tokens[5] != ("literal", wanted):
        return f"new declares {shown(new_const[0])}, base has none: an added PROC_NAME must be {wanted}"
    rows_base = without_empty_case_else([r for r in base.rows if not is_proc_name(r)])
    rows_new = without_empty_case_else([r for r in new.rows if not is_proc_name(r)])
    return first_difference(rows_base, rows_new)


def compare_modules(base_text: str, new_text: str, module: str) -> list[str]:
    """What differs between two versions of one module, beyond what the style allows; first difference only."""
    try:
        base_text = vba_house.renamed_source(base_text, module)
        base_options, base_outside, base_procs = split_module(base_text)
        new_options, new_outside, new_procs = split_module(new_text)
    except LexicalError as exc:
        return [f"{module}: cannot be read: {exc}"]
    if [r.tokens for r in base_options] != [r.tokens for r in new_options]:
        return [f"{module}: Option statements differ: {first_difference(base_options, new_options)}"]
    base_keys = [(p.kind, p.name) for p in base_procs]
    new_keys = [(p.kind, p.name) for p in new_procs]
    if base_keys != new_keys:
        for n in range(max(len(base_keys), len(new_keys))):
            a = base_keys[n] if n < len(base_keys) else None
            b = new_keys[n] if n < len(new_keys) else None
            if a != b:
                return [f"{module}: procedure {n + 1} is {a[1] if a else 'nothing'} in base "
                        f"and {b[1] if b else 'nothing'} in new"]
    found = first_difference(base_outside, new_outside)
    if found:
        return [f"{module}: outside the procedures: {found}"]
    for old, new in zip(base_procs, new_procs, strict=True):
        found = compare_procedure(old, new, module)
        if found:
            return [f"{module}: {new.kind} {new.name}: {found}"]
    return []


def read_tree(directory: Path) -> dict[str, str]:
    """The .bas modules under a directory, by path relative to it."""
    out: dict[str, str] = {}
    for path in sorted(directory.rglob("*.bas")):
        out[path.relative_to(directory).as_posix()] = path.read_bytes().decode("ascii")
    return out


def compare_trees(base_dir: Path, new_dir: Path) -> list[tuple[str, list[str]]]:
    """(path, differences) for every module of either tree, in path order."""
    base, new = read_tree(base_dir), read_tree(new_dir)
    results: list[tuple[str, list[str]]] = []
    for rel in sorted(set(base) | set(new)):
        if rel not in new:
            results.append((rel, [f"{rel}: the module is missing from the new tree"]))
        elif rel not in base:
            results.append((rel, [f"{rel}: the module is not in the base tree"]))
        else:
            module = module_of(new[rel], Path(rel).stem)
            results.append((rel, compare_modules(base[rel], new[rel], module)))
    return results


def run(base_dir: str, new_dir: str) -> int:
    """Print one line per module and return 1 when any module differs beyond what is allowed."""
    try:
        results = compare_trees(Path(base_dir), Path(new_dir))
    except (OSError, UnicodeDecodeError) as exc:
        print(f"FAIL {exc}")
        return 1
    if not results:
        print(f"FAIL no .bas module under {base_dir} or {new_dir}")
        return 1
    failed = 0
    for rel, differences in results:
        if differences:
            failed += 1
            print(f"FAIL {differences[0]}")
        else:
            print(f"ok   {rel}")
    if failed:
        print(f"FAIL {failed} of {len(results)} modules differ in more than comments, layout, the declared "
              f"renames, Const PROC_NAME and an empty Case Else")
        return 1
    print(f"ok   {len(results)} modules differ only in comments, layout, the declared renames, "
          f"Const PROC_NAME and an empty Case Else")
    return 0
