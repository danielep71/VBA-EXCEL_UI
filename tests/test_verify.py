"""Tests for the single verification command (tools/verify.py).

No check really runs here: the runner, the PATH lookup and the clock are
injected, so the tests see which commands would run, in which order, and what
the script reports for each outcome.
"""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from collections.abc import Sequence
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import verify  # noqa: E402

ORDER = ["check-repo", "reformat", "render-docs", "labels", "check-issues", "unittest", "ruff", "mypy"]


class FakeRunner:
    """Record each command and return the exit status chosen for it."""

    def __init__(self, failing: Sequence[str] = ()) -> None:
        self.failing = failing
        self.commands: list[list[str]] = []

    def __call__(self, command: Sequence[str], cwd: Path) -> int:
        self.commands.append(list(command))
        joined = " ".join(command)
        return 1 if any(word in joined for word in self.failing) else 0


def which_without(*missing: str) -> verify.Which:
    def which(tool: str) -> str | None:
        return None if tool in missing else f"/usr/bin/{tool}"
    return which


class Clock:
    """A clock that advances one second per reading."""

    def __init__(self) -> None:
        self.now = 0.0

    def __call__(self) -> float:
        self.now += 1.0
        return self.now


def run(argv: Sequence[str], runner: FakeRunner, missing: Sequence[str] = ()) -> tuple[int, str]:
    out = io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
        code = verify.main(list(argv), runner=runner, which=which_without(*missing), clock=Clock())
    return code, out.getvalue()


def result_lines(output: str) -> dict[str, str]:
    """Return {check name: status} from the recap at the end of the output."""
    recap = output.split("\n\n")[-1]
    found: dict[str, str] = {}
    for line in recap.splitlines():
        words = line.split()
        if len(words) >= 2 and words[0] in ("PASS", "FAIL", "SKIPPED"):
            found[words[1]] = words[0]
    return found


class DeclaredChecks(unittest.TestCase):
    def test_declares_the_eight_checks_in_order(self) -> None:
        self.assertEqual([c.name for c in verify.declared_checks(ROOT)], ORDER)

    def test_python_checks_use_the_running_interpreter(self) -> None:
        commands = [c for check in verify.declared_checks(ROOT, python="/opt/python") for c in check.commands]
        python = [c for c in commands if c[0] not in ("node", "ruff", "mypy")]
        self.assertEqual(len(python), 5)
        self.assertEqual({c[0] for c in python}, {"/opt/python"})

    def test_expands_the_module_globs_itself(self) -> None:
        reformat = verify.declared_checks(ROOT)[1].commands[0]
        modules = reformat[3:]
        self.assertTrue(modules)
        self.assertFalse([m for m in modules if "*" in m])
        for prefix in ("src/", "test/", "demo/"):
            self.assertTrue([m for m in modules if m.startswith(prefix) and m.endswith(".bas")], prefix)


class RunOrder(unittest.TestCase):
    def test_runs_every_command_in_the_declared_order(self) -> None:
        runner = FakeRunner()
        code, _ = run([], runner)
        self.assertEqual(code, 0)
        programs = [" ".join(c[1:3]) if c[0] == sys.executable else " ".join(c[:3]) for c in runner.commands]
        self.assertEqual(programs, [
            "tools/check_repo.py",
            "tools/reformat.py --check",
            "tools/render_docs.py --check",
            "node .github/scripts/labels-sync.mjs --self-test",
            "node .github/scripts/labels-drift.mjs --self-test",
            "tools/check_issues.py --check-config",
            "-m unittest",
            "ruff check tools",
            "mypy --strict tools",
            "mypy --strict tests",
        ])

    def test_keeps_running_after_a_failure(self) -> None:
        runner = FakeRunner(failing=["check_repo.py"])
        code, output = run([], runner)
        self.assertEqual(code, 1)
        self.assertEqual(len(runner.commands), 10)
        self.assertEqual(result_lines(output)["check-repo"], "FAIL")
        self.assertEqual(result_lines(output)["mypy"], "PASS")

    def test_runs_both_commands_of_a_check_and_fails_it_on_either(self) -> None:
        runner = FakeRunner(failing=["--strict tools"])
        code, output = run(["--only", "mypy"], runner)
        self.assertEqual(code, 1)
        self.assertEqual([c[-1] for c in runner.commands], ["tools", "tests"])
        self.assertIn("mypy --strict tools exited 1", output)

    def test_only_and_skip_select_checks(self) -> None:
        runner = FakeRunner()
        code, output = run(["--skip", "ruff,mypy"], runner)
        self.assertEqual(code, 0)
        self.assertEqual(list(result_lines(output)), ORDER[:6])
        self.assertIn("not selected: ruff, mypy", output)

        runner = FakeRunner()
        run(["--only", "labels", "--only", "ruff"], runner)
        self.assertEqual([c[0] for c in runner.commands], ["node", "node", "ruff"])

    def test_rejects_an_unknown_check(self) -> None:
        runner = FakeRunner()
        code, _ = run(["--only", "lint"], runner)
        self.assertEqual(code, 2)
        self.assertEqual(runner.commands, [])

    def test_list_runs_nothing(self) -> None:
        runner = FakeRunner()
        code, output = run(["--list"], runner)
        self.assertEqual(code, 0)
        self.assertEqual(runner.commands, [])
        self.assertIn("mypy --strict tests", output)


class MissingTools(unittest.TestCase):
    def test_skips_a_missing_tool_locally(self) -> None:
        runner = FakeRunner()
        code, output = run([], runner, missing=["node", "ruff"])
        self.assertEqual(code, 0)
        statuses = result_lines(output)
        self.assertEqual(statuses["labels"], "SKIPPED")
        self.assertEqual(statuses["check-issues"], "SKIPPED")
        self.assertEqual(statuses["ruff"], "SKIPPED")
        self.assertEqual(statuses["mypy"], "PASS")
        self.assertIn("node not found on PATH", output)
        self.assertIn("RESULT: INCOMPLETE | checks=8 passed=5 failed=0 skipped=3", output)
        self.assertFalse([c for c in runner.commands if c[0] in ("node", "ruff")])

    def test_fails_a_missing_tool_with_ci(self) -> None:
        runner = FakeRunner()
        code, output = run(["--ci"], runner, missing=["mypy"])
        self.assertEqual(code, 1)
        self.assertEqual(result_lines(output)["mypy"], "FAIL")
        self.assertIn("RESULT: FAIL | checks=8 passed=7 failed=1 skipped=0", output)

    def test_a_check_left_out_by_skip_is_not_a_failure_with_ci(self) -> None:
        code, output = run(["--ci", "--skip", "ruff,mypy"], FakeRunner(), missing=["ruff", "mypy"])
        self.assertEqual(code, 0)
        self.assertIn("RESULT: PASS | checks=6 passed=6 failed=0 skipped=0", output)


class Reporting(unittest.TestCase):
    def test_prints_one_result_line_per_check_and_a_result_line(self) -> None:
        code, output = run([], FakeRunner())
        self.assertEqual(code, 0)
        self.assertEqual(list(result_lines(output)), ORDER)
        self.assertTrue(output.rstrip("\n").endswith("RESULT: PASS | checks=8 passed=8 failed=0 skipped=0"))

    def test_prints_the_formatting_hint_on_drift(self) -> None:
        _, output = run([], FakeRunner(failing=["reformat.py"]))
        self.assertIn("reformat.py --write src/*.bas test/*.bas demo/*.bas", output)

    def test_prints_the_render_hint_on_drift(self) -> None:
        code, output = run(["--only", "render-docs"], FakeRunner(failing=["render_docs.py"]))
        self.assertEqual(code, 1)
        self.assertIn("hint:   python3 tools/render_docs.py --write", output)

    def test_appends_a_markdown_summary(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            summary = Path(tmp) / "summary.md"
            summary.write_text("earlier step\n", encoding="utf-8")
            code, _ = run(["--summary", str(summary), "--skip", "mypy"], FakeRunner(failing=["ruff"]),
                          missing=["node"])
            text = summary.read_text(encoding="utf-8")
        self.assertEqual(code, 1)
        self.assertTrue(text.startswith("earlier step\n"))
        self.assertIn("| Check | Result | Duration |", text)
        self.assertIn("| `check-repo` | PASS | 1.0 s |", text)
        self.assertIn("| `labels` | SKIPPED (node not found on PATH) | 0.0 s |", text)
        self.assertIn("| `ruff` | FAIL (ruff check tools ... exited 1) | 1.0 s |", text)
        self.assertIn("Not selected: `mypy`.", text)
        self.assertIn("**RESULT: FAIL** | checks=7 passed=4 failed=1 skipped=2", text)


if __name__ == "__main__":
    unittest.main()
