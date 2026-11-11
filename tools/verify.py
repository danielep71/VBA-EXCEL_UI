#!/usr/bin/env python3
"""Run every static check of this repository, in one declared order.

This is the single list of the static checks. Contributors run it before they
push, and the static-checks workflow runs it in each of its jobs, so a check
added here runs everywhere at once:

1. check-repo: the static gate, tools/check_repo.py;
2. reformat: house-style normal form of src/, test/ and demo/ *.bas;
3. labels: the self-tests of the two Node.js label scripts;
4. check-issues: the issue forms against the label catalogue (Node.js);
5. unittest: the unit tests under tests/;
6. ruff: Ruff on tools/ and tests/;
7. mypy: mypy --strict on tools/, then on tests/.

The Python checks run with the interpreter that runs this script. Node.js, Ruff
and mypy are found on PATH. When one of them is missing, the checks that need
it are reported as SKIPPED; with --ci they fail instead, because a check that
did not run proves nothing.

Every selected check runs, even after an earlier one fails, so a single run
reports every finding. The script prints one result line per check and a final
RESULT line:

- PASS: every selected check ran and passed;
- INCOMPLETE: none failed, but at least one was skipped for a missing tool;
- FAIL: at least one failed.

Exit status: 0 for PASS or INCOMPLETE, 1 for FAIL, 2 for a usage error.
"""

from __future__ import annotations

import argparse
import os
import shlex
import shutil
import subprocess
import sys
import time
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# The VBA modules that reformat.py checks, expanded here rather than by a
# shell, so the command behaves the same on every platform.
BAS_GLOBS = ("src/*.bas", "test/*.bas", "demo/*.bas")

Runner = Callable[[Sequence[str], Path], int]
Which = Callable[[str], str | None]
Clock = Callable[[], float]


@dataclass(frozen=True)
class Check:
    name: str
    title: str
    commands: tuple[tuple[str, ...], ...]
    tool: str | None = None  # an executable found on PATH; None for the running Python
    hint: tuple[str, ...] = ()  # printed when the check fails


@dataclass(frozen=True)
class Result:
    name: str
    title: str
    status: str  # "PASS", "FAIL" or "SKIPPED"
    detail: str
    seconds: float


def bas_files(root: Path) -> list[str]:
    """Return the VBA modules that the formatter checks, relative to root."""
    files: list[str] = []
    for pattern in BAS_GLOBS:
        files.extend(path.relative_to(root).as_posix() for path in sorted(root.glob(pattern)))
    return files


def declared_checks(root: Path, python: str = sys.executable) -> list[Check]:
    """Return every static check, in the order in which it runs."""
    return [
        Check("check-repo", "Static gate", ((python, "tools/check_repo.py"),)),
        Check(
            "reformat", "House-style normal form",
            ((python, "tools/reformat.py", "--check", *bas_files(root)),),
            hint=(
                "Formatting drift is mechanical. Run:",
                "  python3 tools/reformat.py --write src/*.bas test/*.bas demo/*.bas",
                "then re-import the affected modules into the VBE before committing.",
            ),
        ),
        Check(
            "labels", "Label script self-tests",
            (("node", ".github/scripts/labels-sync.mjs", "--self-test"),
             ("node", ".github/scripts/labels-drift.mjs", "--self-test")),
            tool="node",
        ),
        Check("check-issues", "Issue forms against the label catalogue",
              ((python, "tools/check_issues.py", "--check-config"),), tool="node"),
        Check("unittest", "Unit tests", ((python, "-m", "unittest", "discover", "-s", "tests"),)),
        Check("ruff", "Ruff", (("ruff", "check", "tools", "tests"),), tool="ruff"),
        Check("mypy", "mypy --strict",
              (("mypy", "--strict", "tools"), ("mypy", "--strict", "tests")), tool="mypy"),
    ]


def short(command: Sequence[str]) -> str:
    """Name a command briefly: the program's file name and its first arguments."""
    words = [Path(command[0]).name, *command[1:3]]
    return shlex.join(words) + (" ..." if len(command) > 3 else "")


def run_command(command: Sequence[str], cwd: Path) -> int:
    """Run one command with inherited output and return its exit status."""
    return subprocess.run(list(command), cwd=cwd, check=False).returncode


def run_check(check: Check, root: Path, ci: bool, runner: Runner, which: Which, clock: Clock) -> Result:
    """Run every command of one check; the check passes only if all of them do."""
    if check.tool is not None and which(check.tool) is None:
        detail = f"{check.tool} not found on PATH"
        return Result(check.name, check.title, "FAIL" if ci else "SKIPPED", detail, 0.0)

    started = clock()
    failed: list[str] = []
    for command in check.commands:
        print(f"$ {shlex.join(command)}", flush=True)
        try:
            status = runner(command, root)
        except OSError as error:
            print(f"ERROR: {error}", file=sys.stderr, flush=True)
            status = 127
        if status != 0:
            failed.append(f"{short(command)} exited {status}")
    seconds = clock() - started

    if failed:
        return Result(check.name, check.title, "FAIL", "; ".join(failed), seconds)
    return Result(check.name, check.title, "PASS", "", seconds)


def verdict(results: Sequence[Result]) -> str:
    if any(r.status == "FAIL" for r in results):
        return "FAIL"
    if any(r.status == "SKIPPED" for r in results):
        return "INCOMPLETE"
    return "PASS"


def result_line(result: Result) -> str:
    line = f"{result.status:<8} {result.name:<13} {result.seconds:6.1f}s"
    return f"{line}  {result.detail}" if result.detail else line


def totals(results: Sequence[Result]) -> str:
    counts = {status: sum(r.status == status for r in results) for status in ("PASS", "FAIL", "SKIPPED")}
    return (f"checks={len(results)} passed={counts['PASS']} "
            f"failed={counts['FAIL']} skipped={counts['SKIPPED']}")


def result_report(results: Sequence[Result], unselected: Sequence[str]) -> str:
    """Plain-text recap: one line per check, then the RESULT line."""
    lines = [result_line(r) for r in results]
    if unselected:
        lines.append(f"not selected: {', '.join(unselected)}")
    lines.append(f"RESULT: {verdict(results)} | {totals(results)}")
    return "\n".join(lines) + "\n"


def markdown_report(results: Sequence[Result], unselected: Sequence[str]) -> str:
    """Markdown table of the checks, for a job summary."""
    lines = ["## Static checks", "", "| Check | Result | Duration |", "| --- | --- | --- |"]
    for r in results:
        status = f"{r.status} ({r.detail})" if r.detail else r.status
        lines.append(f"| `{r.name}` | {status.replace('|', '/')} | {r.seconds:.1f} s |")
    lines.append("")
    if unselected:
        lines.append(f"Not selected: {', '.join(f'`{n}`' for n in unselected)}.")
        lines.append("")
    lines.append(f"**RESULT: {verdict(results)}** | {totals(results)}")
    lines.append("")
    return "\n".join(lines) + "\n"


def append_text(path: Path, text: str) -> None:
    """Append text to path, as a job summary file expects; create the file if needed."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", newline="\n") as handle:
        handle.write(text)


def select(checks: Sequence[Check], only: Sequence[str], skip: Sequence[str]) -> tuple[list[Check], list[str]]:
    """Return the checks to run, and the names of those left out, in declared order."""
    known = [c.name for c in checks]
    unknown = sorted(set(only).union(skip).difference(known))
    if unknown:
        raise ValueError(f"unknown check: {', '.join(unknown)}; choose from {', '.join(known)}")
    chosen = [c for c in checks if (not only or c.name in only) and c.name not in skip]
    return chosen, [c.name for c in checks if c not in chosen]


def run_all(checks: Sequence[Check], root: Path, ci: bool, annotate: bool,
            runner: Runner, which: Which, clock: Clock) -> list[Result]:
    """Run the checks in order, printing a header before each and a result after it."""
    results: list[Result] = []
    for number, check in enumerate(checks, 1):
        print(f"==> [{number}/{len(checks)}] {check.name}: {check.title}", flush=True)
        result = run_check(check, root, ci, runner, which, clock)
        results.append(result)
        print(result_line(result), flush=True)
        if result.status == "FAIL":
            for line in check.hint:
                print(f"::notice::{line}" if annotate else f"hint: {line}", flush=True)
        print(flush=True)
    return results


def names(value: str) -> list[str]:
    return [name.strip() for name in value.split(",") if name.strip()]


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run every static check of this repository, in one declared order.",
        epilog="Exit status: 0 PASS or INCOMPLETE, 1 FAIL, 2 usage error.")
    parser.add_argument("--ci", action="store_true",
                        help="fail, rather than skip, a check whose tool (node, ruff, mypy) is missing")
    parser.add_argument("--summary", type=Path,
                        help="append a Markdown table of the results to this file (for example $GITHUB_STEP_SUMMARY)")
    parser.add_argument("--only", type=names, action="extend", default=[], metavar="NAME[,NAME]",
                        help="run only these checks")
    parser.add_argument("--skip", type=names, action="extend", default=[], metavar="NAME[,NAME]",
                        help="leave these checks out; they are listed as not selected")
    parser.add_argument("--list", action="store_true", help="print the checks and their commands, then exit")
    parser.add_argument("--root", type=Path, default=ROOT, help="repository root (default: this checkout)")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None, runner: Runner = run_command,
         which: Which = shutil.which, clock: Clock = time.monotonic) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    root = options.root.resolve()
    checks = declared_checks(root)

    if options.list:
        for check in checks:
            tool = f" (needs {check.tool})" if check.tool else ""
            print(f"{check.name}: {check.title}{tool}")
            for command in check.commands:
                print(f"    {shlex.join(command)}")
        return 0

    try:
        chosen, unselected = select(checks, options.only, options.skip)
    except ValueError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    if not chosen:
        print("ERROR: no check selected", file=sys.stderr)
        return 2

    annotate = os.environ.get("GITHUB_ACTIONS") == "true"
    results = run_all(chosen, root, options.ci, annotate, runner, which, clock)
    print(result_report(results, unselected), end="", flush=True)

    if options.summary is not None:
        try:
            append_text(options.summary, markdown_report(results, unselected))
        except OSError as error:
            print(f"ERROR: cannot write the summary: {error}", file=sys.stderr)
            return 1
    return 1 if verdict(results) == "FAIL" else 0


if __name__ == "__main__":
    raise SystemExit(main())
