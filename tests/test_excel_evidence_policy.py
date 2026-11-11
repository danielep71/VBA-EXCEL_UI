"""Tests for .github/excel-evidence-policy.json, the Excel evidence contract.

The policy is the one place that names the bundle files, the steps, the
certification runner and the configuration. These tests validate the policy
itself and fail when tools/Certify-EXCEL_UI.ps1, tools/check_excel_evidence.py
or docs/EXCEL_CERTIFICATION.md disagrees with it: a consumer that restates a
policy value instead of reading it, or a guide that describes another
contract.
"""

from __future__ import annotations

import ast
import contextlib
import copy
import dataclasses
import io
import json
import re
import sys
import tempfile
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_excel_evidence as evidence  # noqa: E402

POLICY = ROOT / ".github" / "excel-evidence-policy.json"
SCRIPT = ROOT / "tools" / "Certify-EXCEL_UI.ps1"
CHECKER = ROOT / "tools" / "check_excel_evidence.py"
GUIDE = ROOT / "docs" / "EXCEL_CERTIFICATION.md"
POLICY_PATH_TEXT = ".github/excel-evidence-policy.json"


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8").replace("\r\n", "\n")


# ----------------------------------------------------------------------------
# Readers
# ----------------------------------------------------------------------------
def powershell_parts(text: str) -> tuple[str, list[str], list[str]]:
    """Split a PowerShell script into code, string literals and comments.

    Comments and string literals are replaced by spaces in the code, so a
    pattern that matches the code matches neither.
    """
    code: list[str] = []
    strings: list[str] = []
    comments: list[str] = []
    index, size = 0, len(text)
    while index < size:
        char = text[index]
        if text.startswith("<#", index):
            end = text.find("#>", index + 2)
            end = size if end < 0 else end + 2
            comments.append(text[index:end])
            code.append(" ")
            index = end
        elif char == "#":
            end = text.find("\n", index)
            end = size if end < 0 else end
            comments.append(text[index:end])
            index = end
        elif text.startswith("@'\n", index) or text.startswith('@"\n', index):
            closing = "\n" + text[index + 1] + "@"
            end = text.find(closing, index + 3)
            end = size if end < 0 else end + len(closing)
            strings.append(text[index + 3:end - len(closing)])
            code.append(" ")
            index = end
        elif char in "'\"":
            body: list[str] = []
            index += 1
            while index < size:
                if char == '"' and text[index] == "`" and index + 1 < size:
                    body.append(text[index + 1])
                    index += 2
                elif text[index] == char and text.startswith(char * 2, index):
                    body.append(char)
                    index += 2
                elif text[index] == char:
                    index += 1
                    break
                else:
                    body.append(text[index])
                    index += 1
            strings.append("".join(body))
            code.append(" ")
        else:
            code.append(char)
            index += 1
    return "".join(code), strings, comments


def python_strings(source: str) -> list[str]:
    """The string constants of a Python module, docstrings excluded."""
    tree = ast.parse(source)
    docstrings: set[int] = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            body = node.body
            if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant):
                docstrings.add(id(body[0].value))
    return [node.value for node in ast.walk(tree)
            if isinstance(node, ast.Constant) and isinstance(node.value, str) and id(node) not in docstrings]


def section(markdown: str, heading: str) -> str:
    """The body of the level-two section whose heading contains ``heading``."""
    match = re.search(rf"^## [^\n]*{re.escape(heading)}[^\n]*\n(.*?)(?=^## |\Z)", markdown, re.MULTILINE | re.DOTALL)
    return match.group(1) if match else ""


def guide_steps(markdown: str) -> list[str]:
    """The steps the guide's numbered list describes, as step names."""
    titles = re.findall(r"^\d+\. \*\*([^*]+?)\.\*\*", section(markdown, "What it does"), re.MULTILINE)
    return [title.strip().lower().replace(" ", "-") for title in titles]


# ----------------------------------------------------------------------------
# Agreement: each returns the problems it finds
# ----------------------------------------------------------------------------
def checker_problems(policy: evidence.Policy, source: str) -> list[str]:
    """Values of the policy that the checker's code restates."""
    problems = []
    contained = [*policy.required_files, policy.certification_runner, policy.manifest, policy.certification_module]
    # A step may share its name with a host field, such as build.
    host_fields = {name for fields in policy.host_fields.values() for name in fields}
    exact = [*(step for step in policy.steps if step not in host_fields), policy.configuration, policy.component]
    for value in python_strings(source):
        problems += [f"the checker restates {name!r} in {value!r}" for name in contained if name in value]
        if value in exact:
            problems.append(f"the checker restates {value!r}")
    return problems


def script_problems(policy: evidence.Policy, text: str) -> list[str]:
    """Ways in which the PowerShell script disagrees with the policy."""
    problems = []
    code, strings, comments = powershell_parts(text)
    if not any(value.endswith("excel-evidence-policy.json") for value in strings):
        problems.append("the script does not read excel-evidence-policy.json")
    for value in strings:
        for name in (*policy.required_files, policy.certification_runner, policy.manifest):
            if name in value:
                problems.append(f"the script restates {name!r} in {value!r}")
        if value in (policy.configuration,):
            problems.append(f"the script restates {value!r}")
    used = re.findall(r"\bSet-Step\s+'([^']*)'", text)
    first_use = list(dict.fromkeys(used))
    if first_use != list(policy.steps):
        problems.append(f"the script records the steps {first_use}, not the policy's {list(policy.steps)}")
    help_block = comments[0] if comments and comments[0].startswith("<#") else ""
    described = re.findall(r"^\d+\. ([a-z][a-z0-9-]*):", help_block, re.MULTILINE)
    if described != list(policy.steps):
        problems.append(f"the script's help describes the steps {described}, not {list(policy.steps)}")
    roles = set(re.findall(r"\$PolicyFiles\.(\w+)", code))
    if roles != set(evidence.ROLES):
        problems.append(f"the script writes the files of roles {sorted(roles)}, not {sorted(evidence.ROLES)}")
    for section_name, fields in policy.host_fields.items():
        for field_name in fields:
            if not re.search(rf"^\s*{re.escape(field_name)}\s+=", code, re.MULTILINE):
                problems.append(f"the script does not record the {section_name} field {field_name!r}")
    return problems


def guide_problems(policy: evidence.Policy, markdown: str) -> list[str]:
    """Ways in which docs/EXCEL_CERTIFICATION.md disagrees with the policy."""
    problems = []
    if POLICY_PATH_TEXT not in markdown:
        problems.append(f"the guide does not name {POLICY_PATH_TEXT}")
    rows = re.findall(r"^\| `([^`]+)` \|", section(markdown, "The bundle"), re.MULTILINE)
    for name in policy.required_files:
        if name not in rows:
            problems.append(f"the bundle table has no row for {name}")
    steps = guide_steps(markdown)
    if steps != list(policy.steps):
        problems.append(f"the guide describes the steps {steps}, not {list(policy.steps)}")
    for value in (policy.certification_runner, policy.configuration, policy.manifest, *policy.evidence_files):
        if f"`{value}`" not in markdown:
            problems.append(f"the guide does not name `{value}`")
    described = section(markdown, "evidence policy")
    for key in evidence.POLICY_KEYS:
        if f"`{key}`" not in described:
            problems.append(f"the evidence policy section does not describe `{key}`")
    return problems


# ----------------------------------------------------------------------------
# Tests
# ----------------------------------------------------------------------------
class AgreementTests(unittest.TestCase):
    """The script, the checker and the guide agree with the repository policy."""

    def setUp(self) -> None:
        self.policy = evidence.load_policy(POLICY)

    def test_the_checker_reads_the_repository_policy(self) -> None:
        self.assertEqual(evidence.POLICY, POLICY)
        self.assertEqual(evidence.default_policy(), self.policy)

    def test_the_checker_restates_no_policy_value(self) -> None:
        self.assertEqual(checker_problems(self.policy, read(CHECKER)), [])

    def test_the_script_agrees_with_the_policy(self) -> None:
        self.assertEqual(script_problems(self.policy, read(SCRIPT)), [])

    def test_the_guide_agrees_with_the_policy(self) -> None:
        self.assertEqual(guide_problems(self.policy, read(GUIDE)), [])

    def test_the_policy_names_files_of_this_repository(self) -> None:
        manifest = json.loads(read(ROOT / self.policy.manifest))
        self.assertIn(self.policy.configuration, manifest["configurations"])
        modules = manifest["configurations"][self.policy.configuration]["modules"]
        self.assertIn(self.policy.certification_module, modules)
        module = (ROOT / self.policy.certification_module).read_text(encoding="cp1252")
        self.assertRegex(module, rf"(?m)^Public Sub {self.policy.certification_runner}\(\)")


class DisagreementTests(unittest.TestCase):
    """Each agreement check reports a consumer that drifts from the policy."""

    def setUp(self) -> None:
        self.policy = evidence.load_policy(POLICY)
        self.script = read(SCRIPT)
        self.guide = read(GUIDE)

    def renamed_step(self) -> evidence.Policy:
        steps = tuple("build-vba" if step == "build" else step for step in self.policy.steps)
        return dataclasses.replace(self.policy, steps=steps)

    def test_a_renamed_step_is_reported_in_the_script_and_the_guide(self) -> None:
        policy = self.renamed_step()
        self.assertTrue(any("records the steps" in problem for problem in script_problems(policy, self.script)))
        self.assertTrue(any("help describes" in problem for problem in script_problems(policy, self.script)))
        self.assertTrue(any("describes the steps" in problem for problem in guide_problems(policy, self.guide)))

    def test_reordered_steps_are_reported(self) -> None:
        policy = dataclasses.replace(self.policy, steps=tuple(reversed(self.policy.steps)))
        self.assertNotEqual(script_problems(policy, self.script), [])
        self.assertNotEqual(guide_problems(policy, self.guide), [])

    def test_a_renamed_file_is_reported_in_the_guide(self) -> None:
        files = dict(self.policy.files, host="machine.json")
        problems = guide_problems(dataclasses.replace(self.policy, files=files), self.guide)
        self.assertIn("the bundle table has no row for machine.json", problems)

    def test_a_renamed_runner_is_reported_in_the_guide(self) -> None:
        policy = dataclasses.replace(self.policy, certification_runner="Test_Other")
        self.assertIn("the guide does not name `Test_Other`", guide_problems(policy, self.guide))

    def test_a_script_that_restates_a_file_name_or_the_runner_is_reported(self) -> None:
        restated = self.script.replace("Write-JsonFile $PolicyFiles.host $hostRecord",
                                       "Write-JsonFile 'host.json' $hostRecord")
        self.assertNotEqual(restated, self.script)
        problems = script_problems(self.policy, restated)
        self.assertTrue(any("restates 'host.json'" in problem for problem in problems), problems)
        restated = self.script + "\n$CertificationMacro = 'Test_EXCEL_UI_RunReleaseCertification'\n"
        self.assertTrue(any("restates 'Test_EXCEL_UI_RunReleaseCertification'" in problem
                            for problem in script_problems(self.policy, restated)))

    def test_a_script_step_outside_the_policy_is_reported(self) -> None:
        extra = self.script.replace("Set-Step 'close' 'PASS' $closeDetail",
                                    "Set-Step 'close' 'PASS' $closeDetail\n        Set-Step 'archive' 'PASS' ''")
        self.assertNotEqual(extra, self.script)
        self.assertTrue(any("records the steps" in problem for problem in script_problems(self.policy, extra)))

    def test_a_host_field_the_script_does_not_record_is_reported(self) -> None:
        fields = dict(self.policy.host_fields, excel=(*self.policy.host_fields["excel"], "language"))
        problems = script_problems(dataclasses.replace(self.policy, host_fields=fields), self.script)
        self.assertIn("the script does not record the excel field 'language'", problems)

    def test_a_checker_that_restates_a_policy_value_is_reported(self) -> None:
        source = read(CHECKER) + '\nSTEPS = ("preflight", "build")\nSESSION = "session.json"\n'
        problems = checker_problems(self.policy, source)
        self.assertIn("the checker restates 'preflight'", problems)
        self.assertTrue(any("'session.json'" in problem for problem in problems), problems)

    def test_comments_and_docstrings_do_not_count_as_restating(self) -> None:
        self.assertEqual(checker_problems(self.policy, '"""Reads session.json."""\nX = 1\n'), [])
        code, strings, comments = powershell_parts("# session.json\n<# certify #>\n$a = 'it''s' + \"x`\"y\"\n")
        self.assertEqual(strings, ["it's", 'x"y'])
        self.assertEqual(len(comments), 2)
        self.assertNotIn("session.json", code)


class PolicyValidationTests(unittest.TestCase):
    """The checker validates the policy before it reads a bundle."""

    def setUp(self) -> None:
        self.document: dict[str, Any] = json.loads(read(POLICY))
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.path = Path(temp.name) / "policy.json"

    def load(self, change: Callable[[dict[str, Any]], None]) -> evidence.Policy:
        document = copy.deepcopy(self.document)
        change(document)
        self.path.write_text(json.dumps(document), encoding="utf-8")
        return evidence.load_policy(self.path)

    def test_the_repository_policy_is_valid(self) -> None:
        policy = evidence.load_policy(POLICY)
        self.assertEqual(policy.required_files, tuple(self.document["required_files"][role] for role in evidence.ROLES))
        self.assertEqual(list(policy.steps), self.document["steps"])
        self.assertEqual(sorted(self.document), sorted(evidence.POLICY_KEYS))

    def test_an_unusable_policy_is_rejected(self) -> None:
        def put(key: str, value: Any) -> Callable[[dict[str, Any]], None]:
            return lambda document: document.update({key: value})

        def put_in(table: str, key: str, value: Any) -> Callable[[dict[str, Any]], None]:
            return lambda document: document[table].update({key: value})

        steps = self.document["steps"]
        cases: dict[str, tuple[Callable[[dict[str, Any]], None], str]] = {
            "missing key": (lambda document: document.pop("steps"), "missing steps"),
            "unknown key": (put("stepz", []), "unknown stepz"),
            "no description": (put("description", ""), "description must be non-empty text"),
            "other policy schema": (put("schema_version", 2), "schema_version 2 is not supported"),
            "boolean evidence schema": (put("evidence_schema", True), "evidence_schema must be a positive integer"),
            "steps not a list": (put("steps", "preflight"), "steps must be a non-empty list"),
            "no steps": (put("steps", []), "steps must be a non-empty list"),
            "repeated step": (put("steps", [*steps, steps[0]]), f"steps lists {steps[0]} more than once"),
            "step with a space": (put("steps", ["pre flight"]), "steps[0] 'pre flight' is not a valid name"),
            "step not text": (put("steps", [1]), "steps[0] must be non-empty text"),
            "missing role": (lambda document: document["required_files"].pop("host"),
                             "required_files must name exactly the roles"),
            "file in a folder": (put_in("required_files", "host", "x/host.json"),
                                 "required_files.host 'x/host.json' is not a valid name"),
            "one file for two roles": (put_in("required_files", "host", "session.json"),
                                       "names the same file for two roles"),
            "two roles differing in case": (put_in("required_files", "host", "SESSION.json"),
                                            "names the same file for two roles"),
            "evidence outside the bundle": (put("evidence_files", ["other.json"]),
                                            "evidence_files other.json are not required files"),
            "session as evidence": (put("evidence_files", ["session.json"]),
                                    "must not include the session record"),
            "path outside the repository": (put("manifest", "../vba-projects.json"),
                                            "must not contain . or .. segments"),
            "runner not an identifier": (put("certification_runner", "Run Me"), "is not a valid name"),
            "missing host section": (lambda document: document["host_fields"].pop("excel"),
                                     "host_fields must have exactly the sections"),
            "unreported host field": (put_in("host_fields", "excel", ["version", "build"]),
                                      "host_fields.excel must include bitness"),
            "match outside the host fields": (put_in("excel_matches_certification", "language", "language"),
                                              "'language' is not in host_fields.excel"),
            "no bitness": (put("bitness", []), "bitness must be a non-empty list"),
        }
        for name, (change, fragment) in cases.items():
            with self.subTest(name):
                with self.assertRaises(evidence.PolicyError) as raised:
                    self.load(change)
                self.assertIn(fragment, str(raised.exception))

    def test_a_policy_that_is_not_json_or_repeats_a_key_is_rejected(self) -> None:
        for text, fragment in (("{", "is not valid UTF-8 JSON"), ('{"steps": [], "steps": []}', "duplicate key"),
                               ("[]", "is not a JSON object")):
            with self.subTest(text):
                self.path.write_text(text, encoding="utf-8")
                with self.assertRaises(evidence.PolicyError) as raised:
                    evidence.load_policy(self.path)
                self.assertIn(fragment, str(raised.exception))
        with self.assertRaises(evidence.PolicyError) as raised:
            evidence.load_policy(self.path.with_name("absent.json"))
        self.assertIn("cannot read the evidence policy", str(raised.exception))

    def test_main_reports_an_unusable_policy_as_a_setup_error(self) -> None:
        self.path.write_text("{}", encoding="utf-8")
        err = io.StringIO()
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
            code = evidence.main(["--inventory", "--policy", str(self.path)])
        self.assertEqual(code, 2)
        self.assertIn("ERROR:", err.getvalue())
        self.assertIn("missing", err.getvalue())


if __name__ == "__main__":
    unittest.main()
