#!/usr/bin/env python3
"""The mandatory-case manifest of release certification (#42), and its static proof.

tools/certification-manifest.json is the release contract of
Test_EXCEL_UI_RunReleaseCertification: the units and the mandatory cases a run
must execute exactly once, in the order it executes them. A case identifier is
the name of the procedure that runs it.

The VBA cannot read that file, so the certification module registers the same
list in TST_CertManifestCase, dispatches every case from TST_CertDispatchCase
and reaches each one through a single TST_CertRunCase call. Four hand-kept
lists that must agree are one list that can drift, so this module fails the
static gate when they do not, and its fixtures prove that it does:

- manifest: the file is well formed, its cases are unique, the cases the issue
  names are still listed, and a unit without a runner has exactly one case;
- registry: TST_CertManifestCase lists the manifest's cases, in order;
- dispatch: TST_CertDispatchCase has exactly one label per case, no other
  label, and each label calls the procedure it names;
- call-site: each case is reached through exactly one TST_CertRunCase call, a
  unit that has a runner reaches its cases in order from that runner, and no
  case the manifest does not list is wired in;
- unit: the certification runner runs the units in manifest order, isolated
  ones first, and TST_CertExecuteUnit dispatches each of them;
- resolve: every identifier names a real Sub;
- classification: every TST_Case_ procedure and every public Test_EXCEL_UI_
  runner is a mandatory case, a unit runner, the release runner, a listed
  focused case or a listed interactive runner, so a new one needs a decision;
- achieved: when the manifest has an achieved_state section (#28), the module that
  holds the runner of its unit names every element and scope it lists, as string
  literals of its code, so an observation cannot be deleted or renamed in the VBA
  while the manifest, and the evidence checker that reads it, still require it;
- version: TEST_CERT_MANIFEST_VERSION equals the manifest's version.

Identifiers that begin with SelfTest. are synthetic cases of the run faults
that prove a failed unit and an inexact inventory fail the outer call; they are
not release cases and are exempt.

This reads source text. It is not a VBA compiler and it does not show that a
case ran: the certification evidence does, and tools/check_excel_evidence.py
compares that evidence with this manifest.
"""

from __future__ import annotations

import json
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from vba_lex import LexicalError, split_code_comment

SCHEMA = 1
MANIFEST_KEYS = ("schema_version", "manifest_version", "description", "units", "required_cases",
                 "focused_cases", "interactive_runners")
# Keys a manifest may leave out. A manifest without achieved_state predates the achieved-state unit (#28).
OPTIONAL_MANIFEST_KEYS = ("achieved_state",)
ACHIEVED_KEYS = ("unit", "description", "observations")
OBSERVATION_KEYS = ("element", "scope")
OBSERVATION_FLAGS = ("may_be_unverifiable", "may_skip_unreadable_windows")
EXECUTIONS = ("isolated", "live")
CASE_ID = re.compile(r"(?:TST_Case_|Test_EXCEL_UI_)[A-Za-z0-9_]+")
IDENTIFIER = re.compile(r"[A-Za-z][A-Za-z0-9_]*")
SYNTHETIC = "SelfTest."

REGISTRY = "TST_CertManifestCase"
DISPATCHER = "TST_CertDispatchCase"
UNIT_DISPATCHER = "TST_CertExecuteUnit"
WRAPPER = "TST_CertRunCase"
VERSION_CONSTANT = "TEST_CERT_MANIFEST_VERSION"

PROCEDURE = re.compile(
    r"^(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.IGNORECASE)
END_PROCEDURE = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.IGNORECASE)
CASE_INDEX = re.compile(r"Case\s+(\d+)", re.IGNORECASE)
CASE_LABEL = re.compile(r'Case\s+"([^"]*)"', re.IGNORECASE)
UNIT_ASSIGNMENT = re.compile(r'UnitName\s*=\s*"([^"]*)"', re.IGNORECASE)
CASE_ASSIGNMENT = re.compile(r'CaseId\s*=\s*"([^"]*)"', re.IGNORECASE)
RUN_CASE_CALL = re.compile(rf'{WRAPPER}\s+"([^"]+)"', re.IGNORECASE)
RUN_UNIT_CALL = re.compile(r'TST_CertRunUnit\s+"([^"]+)"', re.IGNORECASE)
RUN_ISOLATED_CALL = re.compile(r'TST_CertRunIsolatedUnit\s+"([^"]+)"', re.IGNORECASE)
VERSION_DECLARATION = re.compile(rf"Public\s+Const\s+{VERSION_CONSTANT}\s+As\s+Long\s*=\s*(\d+)", re.IGNORECASE)


class ManifestError(ValueError):
    """A manifest file that cannot be used."""


@dataclass(frozen=True)
class Unit:
    name: str
    execution: str
    runner: str | None
    cases: tuple[str, ...]


@dataclass(frozen=True)
class Observation:
    """One element under one scope that the achieved-state unit must record (#28)."""

    element: str
    scope: str
    may_be_unverifiable: bool
    may_skip_unreadable_windows: bool = False


@dataclass(frozen=True)
class AchievedState:
    """The observations of the unit that reads back the state a write achieved."""

    unit: str
    observations: tuple[Observation, ...]


@dataclass(frozen=True)
class Manifest:
    version: int
    units: tuple[Unit, ...]
    required: tuple[str, ...]
    focused: dict[str, str]
    interactive: dict[str, str]
    achieved: AchievedState | None = None

    @property
    def pairs(self) -> tuple[tuple[str, str], ...]:
        """Every (unit, case) in the order a run executes them."""
        return tuple((unit.name, case) for unit in self.units for case in unit.cases)

    @property
    def cases(self) -> tuple[str, ...]:
        return tuple(case for _, case in self.pairs)


# ----------------------------------------------------------------------------
# The manifest file
# ----------------------------------------------------------------------------
def fail(message: str) -> ManifestError:
    return ManifestError(message)


def text_of(value: object, key: str, pattern: re.Pattern[str] | None = None) -> str:
    if not isinstance(value, str) or not value.strip():
        raise fail(f"{key} must be non-empty text")
    if pattern is not None and not pattern.fullmatch(value):
        raise fail(f"{key} {value!r} is not a valid name")
    return value


def table_of(value: object, key: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise fail(f"{key} must be an object")
    return dict(value)


def parse_unit(value: object, index: int) -> Unit:
    where = f"units[{index}]"
    if not isinstance(value, dict):
        raise fail(f"{where} must be an object")
    allowed = {"name", "execution", "runner", "cases"}
    unknown = sorted(set(value) - allowed)
    missing = sorted({"name", "execution", "cases"} - set(value))
    if unknown or missing:
        raise fail(f"{where}: " + "; ".join(([f"missing {', '.join(missing)}"] if missing else [])
                                             + ([f"unknown {', '.join(unknown)}"] if unknown else [])))
    name = text_of(value["name"], f"{where}.name", IDENTIFIER)
    execution = value["execution"]
    if execution not in EXECUTIONS:
        raise fail(f"{where}.execution must be one of {', '.join(EXECUTIONS)}")
    runner = text_of(value["runner"], f"{where}.runner", IDENTIFIER) if "runner" in value else None
    cases = value["cases"]
    if not isinstance(cases, list) or not cases:
        raise fail(f"{where}.cases must be a non-empty list")
    listed = tuple(text_of(case, f"{where}.cases[{position}]", CASE_ID) for position, case in enumerate(cases))
    if runner is None and len(listed) != 1:
        raise fail(f"{where}: a unit without a runner has exactly one case, the procedure that runs it")
    return Unit(name, str(execution), runner, listed)


def parse_manifest(text: str) -> Manifest:
    """Validate the manifest document; raise ManifestError on the first fault."""

    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise fail(f"duplicate key {key!r}")
            result[key] = value
        return result

    try:
        document = json.loads(text.lstrip("﻿"), object_pairs_hook=unique)
    except json.JSONDecodeError as error:
        raise fail(f"not valid JSON: {error}") from error
    if not isinstance(document, dict):
        raise fail("the manifest is not a JSON object")
    missing = [key for key in MANIFEST_KEYS if key not in document]
    unknown = sorted(key for key in document if key not in MANIFEST_KEYS + OPTIONAL_MANIFEST_KEYS)
    if missing or unknown:
        raise fail("; ".join(([f"missing {', '.join(missing)}"] if missing else [])
                             + ([f"unknown {', '.join(unknown)}"] if unknown else [])))
    text_of(document["description"], "description")
    if document["schema_version"] != SCHEMA or type(document["schema_version"]) is not int:
        raise fail(f"schema_version {document['schema_version']!r} is not supported (expected {SCHEMA})")
    version = document["manifest_version"]
    if type(version) is not int or version < 1:
        raise fail("manifest_version must be a positive integer")
    units_value = document["units"]
    if not isinstance(units_value, list) or not units_value:
        raise fail("units must be a non-empty list")
    units = tuple(parse_unit(unit, index) for index, unit in enumerate(units_value))
    names = [unit.name for unit in units]
    repeated = sorted({name for name in names if names.count(name) > 1})
    if repeated:
        raise fail(f"units lists {', '.join(repeated)} more than once")
    cases = [case for unit in units for case in unit.cases]
    repeated = sorted({case for case in cases if cases.count(case) > 1})
    if repeated:
        raise fail(f"cases are listed more than once: {', '.join(repeated)}")
    executions = [unit.execution for unit in units]
    if executions != sorted(executions, key=EXECUTIONS.index):
        raise fail("isolated units must come before live units: the run executes them first")
    required = parse_required(document["required_cases"], set(cases))
    focused = parse_reasons(document["focused_cases"], "focused_cases", CASE_ID)
    clash = sorted(set(focused) & set(cases))
    if clash:
        raise fail(f"focused_cases lists mandatory cases: {', '.join(clash)}")
    interactive = parse_reasons(document["interactive_runners"], "interactive_runners", IDENTIFIER)
    runners = {unit.runner for unit in units if unit.runner} | set(cases)
    clash = sorted(set(interactive) & runners)
    if clash:
        raise fail(f"interactive_runners lists mandatory or unit runners: {', '.join(clash)}")
    achieved = parse_achieved(document["achieved_state"], names) if "achieved_state" in document else None
    return Manifest(version, units, required, focused, interactive, achieved)


def parse_achieved(value: object, unit_names: list[str]) -> AchievedState:
    """Validate the achieved_state section: a unit of the manifest and its unique observations."""
    where = "achieved_state"
    table = table_of(value, where)
    missing = [key for key in ACHIEVED_KEYS if key not in table]
    unknown = sorted(key for key in table if key not in ACHIEVED_KEYS)
    if missing or unknown:
        raise fail(f"{where}: " + "; ".join(([f"missing {', '.join(missing)}"] if missing else [])
                                             + ([f"unknown {', '.join(unknown)}"] if unknown else [])))
    unit = text_of(table["unit"], f"{where}.unit", IDENTIFIER)
    if unit not in unit_names:
        raise fail(f"{where}.unit {unit!r} is not a unit of the manifest")
    text_of(table["description"], f"{where}.description")
    listed = table["observations"]
    if not isinstance(listed, list) or not listed:
        raise fail(f"{where}.observations must be a non-empty list")
    observations: list[Observation] = []
    for index, item in enumerate(listed):
        here = f"{where}.observations[{index}]"
        if not isinstance(item, dict):
            raise fail(f"{here} must be an object")
        unexpected = sorted(set(item) - set(OBSERVATION_KEYS) - set(OBSERVATION_FLAGS))
        absent = [key for key in OBSERVATION_KEYS if key not in item]
        if unexpected or absent:
            raise fail(f"{here}: " + "; ".join(([f"missing {', '.join(absent)}"] if absent else [])
                                                 + ([f"unknown {', '.join(unexpected)}"] if unexpected else [])))
        flags = {name: item.get(name, False) for name in OBSERVATION_FLAGS}
        for name, flag in flags.items():
            if type(flag) is not bool:
                raise fail(f"{here}.{name} must be true or false")
        if flags["may_be_unverifiable"] and flags["may_skip_unreadable_windows"]:
            raise fail(f"{here}: an observation is exempt as a whole or for its unreadable windows, not both")
        observations.append(Observation(text_of(item["element"], f"{here}.element", IDENTIFIER),
                                        text_of(item["scope"], f"{here}.scope", IDENTIFIER),
                                        flags["may_be_unverifiable"], flags["may_skip_unreadable_windows"]))
    pairs = [(item.element, item.scope) for item in observations]
    repeated = sorted({pair for pair in pairs if pairs.count(pair) > 1})
    if repeated:
        raise fail(f"{where}.observations lists {', '.join('.'.join(pair) for pair in repeated)} more than once")
    return AchievedState(unit, tuple(observations))


def parse_required(value: object, mandatory: set[str]) -> tuple[str, ...]:
    if not isinstance(value, list) or not value:
        raise fail("required_cases must be a non-empty list")
    required: list[str] = []
    for index, item in enumerate(value):
        where = f"required_cases[{index}]"
        if not isinstance(item, dict) or sorted(item) != ["case", "note"]:
            raise fail(f"{where} must be an object with exactly case and note")
        case = text_of(item["case"], f"{where}.case", CASE_ID)
        text_of(item["note"], f"{where}.note")
        if case not in mandatory:
            raise fail(f"{where}: {case} is not a mandatory case of any unit")
        required.append(case)
    repeated = sorted({case for case in required if required.count(case) > 1})
    if repeated:
        raise fail(f"required_cases lists {', '.join(repeated)} more than once")
    return tuple(required)


def parse_reasons(value: object, key: str, pattern: re.Pattern[str]) -> dict[str, str]:
    table = table_of(value, key)
    return {text_of(name, f"{key} name", pattern): text_of(reason, f"{key}.{name}") for name, reason in table.items()}


# ----------------------------------------------------------------------------
# The VBA source
# ----------------------------------------------------------------------------
@dataclass
class Procedure:
    name: str
    kind: str
    public: bool
    module: str
    line: int
    statements: list[tuple[int, str]]


def logical_statements(text: str) -> list[tuple[int, str]]:
    """The statements of a module with comments removed and continuations joined."""
    statements: list[tuple[int, str]] = []
    pending: list[str] = []
    start = 0
    for number, raw in enumerate(text.split("\n"), start=1):
        try:
            code = split_code_comment(raw)[0].rstrip()
        except LexicalError:
            code = raw.rstrip()
        continued = code.endswith(" _") or code == "_"
        piece = code[:-1].rstrip() if continued else code
        if not pending:
            start = number
        pending.append(piece.strip())
        if not continued:
            joined = " ".join(part for part in pending if part)
            if joined:
                statements.append((start, joined))
            pending = []
    return statements


def parse_procedures(module: str, text: str) -> list[Procedure]:
    procedures: list[Procedure] = []
    current: Procedure | None = None
    for line, statement in logical_statements(text):
        if current is None:
            match = PROCEDURE.match(statement)
            if match:
                current = Procedure(match[3], match[2].split()[0].title(), (match[1] or "Public").lower() == "public",
                                    module, line, [])
            continue
        if END_PROCEDURE.match(statement):
            procedures.append(current)
            current = None
        else:
            current.statements.append((line, statement))
    return procedures


def find_procedures(procedures: list[Procedure], name: str) -> list[Procedure]:
    return [procedure for procedure in procedures if procedure.name.lower() == name.lower()]


def where(procedure: Procedure, line: int | None = None) -> str:
    return f"{procedure.module}:{line if line is not None else procedure.line}"


def registry_entries(procedure: Procedure) -> tuple[list[tuple[int, str, str]], list[str]]:
    """The (index, unit, case) entries of TST_CertManifestCase, and what is wrong with their shape."""
    entries: list[tuple[int, str, str]] = []
    problems: list[str] = []
    index: int | None = None
    unit: str | None = None
    case: str | None = None
    line = procedure.line

    def flush() -> None:
        if index is None:
            return
        if unit is None or case is None:
            problems.append(f"{where(procedure, line)}: registry entry {index} must assign UnitName and CaseId")
        else:
            entries.append((index, unit, case))

    for number, statement in procedure.statements:
        match = CASE_INDEX.fullmatch(statement)
        if match:
            flush()
            index, unit, case, line = int(match[1]), None, None, number
        elif re.fullmatch(r"Case\s+Else", statement, re.IGNORECASE):
            flush()
            index = None
            break
        elif index is not None:
            assigned_unit = UNIT_ASSIGNMENT.fullmatch(statement)
            assigned_case = CASE_ASSIGNMENT.fullmatch(statement)
            if assigned_unit:
                unit = assigned_unit[1]
            elif assigned_case:
                case = assigned_case[1]
    else:
        flush()
    positions = [entry[0] for entry in entries]
    if positions != list(range(1, len(positions) + 1)):
        problems.append(f"{where(procedure)}: registry indexes {positions} are not 1..{len(positions)} in order")
    return entries, problems


def dispatch_labels(procedure: Procedure) -> tuple[list[tuple[int, str, str]], list[str]]:
    """The (line, label, called procedure) branches of TST_CertDispatchCase."""
    branches: list[tuple[int, str, str]] = []
    problems: list[str] = []
    label: tuple[int, str] | None = None
    called: str | None = None

    def flush() -> None:
        if label is not None:
            branches.append((label[0], label[1], called or ""))

    for number, statement in procedure.statements:
        if re.match(r"Case\b", statement, re.IGNORECASE):
            flush()
            label, called = None, None
            if re.fullmatch(r"Case\s+Else", statement, re.IGNORECASE):
                break
            match = CASE_LABEL.fullmatch(statement)
            if match:
                label = (number, match[1])
            else:
                problems.append(f"{where(procedure, number)}: a dispatch label must name exactly one case: {statement}")
        elif label is not None and called is None and not statement.lower().startswith(("select ", "end ")):
            match_call = re.match(r"(?:Call\s+)?(\w+)", statement, re.IGNORECASE)
            called = match_call[1] if match_call else ""
    else:
        flush()
    return branches, problems


def call_sites(procedures: list[Procedure]) -> list[tuple[str, Procedure, int]]:
    """Every TST_CertRunCase call with a literal identifier: (identifier, procedure, line)."""
    found: list[tuple[str, Procedure, int]] = []
    for procedure in procedures:
        if procedure.name.lower() == WRAPPER.lower():
            continue
        for line, statement in procedure.statements:
            match = RUN_CASE_CALL.match(statement)
            if match:
                found.append((match[1], procedure, line))
    return found


def unit_calls(procedure: Procedure) -> list[tuple[str, str]]:
    """The ("isolated" | "live", unit) calls of the certification runner, in order, without synthetic units."""
    calls: list[tuple[str, str]] = []
    for _, statement in procedure.statements:
        isolated = RUN_ISOLATED_CALL.match(statement)
        live = RUN_UNIT_CALL.match(statement)
        if isolated and not isolated[1].startswith(SYNTHETIC):
            calls.append(("isolated", isolated[1]))
        elif live and not live[1].startswith(SYNTHETIC):
            calls.append(("live", live[1]))
    return calls


def unit_branches(procedure: Procedure) -> dict[str, list[list[str]]]:
    """The statements of each Case "unit" branch of TST_CertExecuteUnit, one list per branch."""
    branches: dict[str, list[list[str]]] = {}
    current: list[str] | None = None
    for _, statement in procedure.statements:
        if re.match(r"Case\b", statement, re.IGNORECASE):
            match = CASE_LABEL.fullmatch(statement)
            current = []
            if match:
                branches.setdefault(match[1], []).append(current)
        elif current is not None:
            current.append(statement)
    return branches


# ----------------------------------------------------------------------------
# The rules
# ----------------------------------------------------------------------------
def sources_procedures(sources: dict[str, str]) -> list[Procedure]:
    procedures: list[Procedure] = []
    for module in sorted(sources):
        procedures.extend(parse_procedures(module, sources[module]))
    return procedures


def single(procedures: list[Procedure], name: str, findings: list[str]) -> Procedure | None:
    found = find_procedures(procedures, name)
    if len(found) != 1:
        findings.append(f"[call-site] {name} is defined {len(found)} times; expected exactly once")
        return None
    return found[0]


def registry_findings(manifest: Manifest, procedures: list[Procedure]) -> list[str]:
    findings: list[str] = []
    procedure = single(procedures, REGISTRY, findings)
    if procedure is None:
        return findings
    entries, problems = registry_entries(procedure)
    findings.extend(f"[registry] {problem}" for problem in problems)
    registered = [(unit, case) for _, unit, case in entries]
    expected = list(manifest.pairs)
    for unit, case in expected:
        if (unit, case) not in registered:
            findings.append(f"[registry] {where(procedure)}: {REGISTRY} does not register {case} (unit {unit})")
    for unit, case in registered:
        if (unit, case) not in expected:
            findings.append(f"[registry] {where(procedure)}: {REGISTRY} registers {case} (unit {unit}), "
                            "which the manifest does not list")
    repeated = sorted(case for case, count in Counter(case for _, case in registered).items() if count > 1)
    findings.extend(f"[registry] {where(procedure)}: {REGISTRY} registers {case} more than once" for case in repeated)
    if not findings and registered != expected:
        findings.append(f"[registry] {where(procedure)}: {REGISTRY} registers the cases in another order "
                        "than the manifest")
    return findings


def dispatch_findings(manifest: Manifest, procedures: list[Procedure]) -> list[str]:
    findings: list[str] = []
    procedure = single(procedures, DISPATCHER, findings)
    if procedure is None:
        return findings
    branches, problems = dispatch_labels(procedure)
    findings.extend(f"[dispatch] {problem}" for problem in problems)
    mandatory = set(manifest.cases)
    seen = Counter(label for _, label, _ in branches)
    for line, label, called in branches:
        if seen[label] > 1 and label in mandatory and [item for item in branches if item[1] == label][0][0] != line:
            findings.append(f"[dispatch] {where(procedure, line)}: {label} is dispatched more than once")
        if label not in mandatory:
            findings.append(f"[dispatch] {where(procedure, line)}: {label} is dispatched but the manifest "
                            "does not list it")
        elif called.lower() != label.lower():
            findings.append(f"[dispatch] {where(procedure, line)}: the branch for {label} calls {called or 'nothing'}")
    for case in manifest.cases:
        if seen[case] == 0:
            findings.append(f"[dispatch] {where(procedure)}: {case} has no dispatch branch")
    return findings


def call_site_findings(manifest: Manifest, procedures: list[Procedure]) -> list[str]:
    findings: list[str] = []
    sites = [site for site in call_sites(procedures) if not site[0].startswith(SYNTHETIC)]
    mandatory = set(manifest.cases)
    counts = Counter(identifier for identifier, _, _ in sites)
    for identifier, procedure, line in sites:
        if identifier not in mandatory:
            findings.append(f"[call-site] {where(procedure, line)}: {WRAPPER} runs {identifier}, "
                            "which the manifest does not list")
    for case in manifest.cases:
        if counts[case] == 0:
            findings.append(f"[call-site] {case} is never run through {WRAPPER}")
        elif counts[case] > 1:
            lines = ", ".join(where(procedure, line) for identifier, procedure, line in sites if identifier == case)
            findings.append(f"[call-site] {case} is run through {WRAPPER} {counts[case]} times ({lines})")
    for unit in manifest.units:
        if unit.runner is None:
            continue
        runner = find_procedures(procedures, unit.runner)
        if len(runner) != 1:
            findings.append(f"[call-site] the runner {unit.runner} of unit {unit.name} is defined {len(runner)} times")
            continue
        reached = [identifier for identifier, procedure, _ in sites if procedure is runner[0]]
        if reached != list(unit.cases):
            findings.append(f"[call-site] {where(runner[0])}: {unit.runner} runs {reached or 'no case'} through "
                            f"{WRAPPER}, not unit {unit.name}'s cases in order {list(unit.cases)}")
    return findings


def unit_findings(manifest: Manifest, procedures: list[Procedure], release_runner: str) -> list[str]:
    findings: list[str] = []
    runner = find_procedures(procedures, release_runner)
    if len(runner) != 1:
        findings.append(f"[unit] the release runner {release_runner} is defined {len(runner)} times")
    else:
        expected = [(unit.execution, unit.name) for unit in manifest.units]
        calls = unit_calls(runner[0])
        if calls != expected:
            findings.append(f"[unit] {where(runner[0])}: {release_runner} runs the units {calls}, "
                            f"not the manifest's {expected}")
    dispatcher = single(procedures, UNIT_DISPATCHER, findings)
    if dispatcher is None:
        return findings
    branches = unit_branches(dispatcher)
    for unit in manifest.units:
        found = branches.get(unit.name, [])
        if len(found) != 1:
            findings.append(f"[unit] {where(dispatcher)}: {UNIT_DISPATCHER} has {len(found)} branches for unit "
                            f"{unit.name}; expected one")
            continue
        statements = found[0]
        if unit.runner is not None:
            if not any(re.match(rf"{re.escape(unit.runner)}\b", statement, re.IGNORECASE) for statement in statements):
                findings.append(f"[unit] {where(dispatcher)}: the branch for unit {unit.name} does not run "
                                f"{unit.runner}")
        elif not any(re.match(rf'{WRAPPER}\s+"{re.escape(unit.cases[0])}"', statement, re.IGNORECASE)
                     for statement in statements):
            findings.append(f"[unit] {where(dispatcher)}: the branch for unit {unit.name} does not run "
                            f"{unit.cases[0]} through {WRAPPER}")
    return findings


def resolve_findings(manifest: Manifest, procedures: list[Procedure]) -> list[str]:
    findings: list[str] = []
    wanted = [(case, "mandatory case") for case in manifest.cases]
    wanted += [(unit.runner, f"runner of unit {unit.name}") for unit in manifest.units if unit.runner]
    wanted += [(case, "focused case") for case in manifest.focused]
    wanted += [(name, "interactive runner") for name in manifest.interactive]
    for name, role in wanted:
        found = find_procedures(procedures, name)
        if not found:
            findings.append(f"[resolve] {name}, a {role}, is not defined in any test module")
        elif role != "focused case" and not any(procedure.kind == "Sub" for procedure in found):
            findings.append(f"[resolve] {name}, a {role}, is not a Sub")
    return findings


def classification_findings(manifest: Manifest, procedures: list[Procedure], release_runner: str) -> list[str]:
    findings: list[str] = []
    known = set(manifest.cases) | set(manifest.focused)
    for procedure in procedures:
        if procedure.name.startswith("TST_Case_") and procedure.name not in known:
            findings.append(f"[classification] {where(procedure)}: {procedure.name} is neither a mandatory case nor a "
                            "listed focused case; list it in the manifest")
    allowed = set(manifest.cases) | set(manifest.interactive) | {release_runner}
    allowed |= {unit.runner for unit in manifest.units if unit.runner}
    for procedure in procedures:
        if procedure.public and procedure.kind == "Sub" and procedure.name.startswith("Test_EXCEL_UI_") \
                and procedure.name not in allowed:
            findings.append(f"[classification] {where(procedure)}: the public runner {procedure.name} is neither a "
                            "mandatory case, a unit runner nor a listed interactive runner; list it in the manifest")
    for name in manifest.interactive:
        for procedure in find_procedures(procedures, name):
            if not procedure.public:
                findings.append(f"[classification] {where(procedure)}: the interactive runner {name} is not public")
    return findings


STRING_LITERAL = re.compile(r'"((?:[^"]|"")*)"')


def code_literals(text: str) -> set[str]:
    """The string literals of the code of a module, comments excluded."""
    found: set[str] = set()
    for _, statement in logical_statements(text):
        found.update(match.replace('""', '"') for match in STRING_LITERAL.findall(statement))
    return found


def achieved_findings(manifest: Manifest, sources: dict[str, str], procedures: list[Procedure]) -> list[str]:
    """The module of the achieved-state unit's runner names every element and scope the manifest lists."""
    achieved = manifest.achieved
    if achieved is None:
        return []
    unit = next(unit for unit in manifest.units if unit.name == achieved.unit)
    if unit.runner is None:
        return [f"[achieved] unit {unit.name} has no runner, so no module holds its observations"]
    runner = find_procedures(procedures, unit.runner)
    if len(runner) != 1:
        return [f"[achieved] the runner {unit.runner} of unit {unit.name} is defined {len(runner)} times"]
    module = runner[0].module
    literals = code_literals(sources[module])
    findings: list[str] = []
    for observation in achieved.observations:
        for kind, name in (("element", observation.element), ("scope", observation.scope)):
            if name not in literals:
                findings.append(f"[achieved] {module}: the {kind} {name!r} of the observation "
                                f"{observation.element}.{observation.scope} is not a string literal of its code")
    return findings


def version_findings(manifest: Manifest, sources: dict[str, str]) -> list[str]:
    declared = [(module, int(match[1])) for module, text in sorted(sources.items())
                for match in VERSION_DECLARATION.finditer(text)]
    if len(declared) != 1:
        return [f"[version] {VERSION_CONSTANT} is declared {len(declared)} times; expected exactly once"]
    module, number = declared[0]
    if number != manifest.version:
        return [f"[version] {module}: {VERSION_CONSTANT} is {number}, the manifest's manifest_version is "
                f"{manifest.version}"]
    return []


def findings_for(manifest_text: str, sources: dict[str, str], release_runner: str) -> list[str]:
    """Every way in which the VBA in ``sources`` disagrees with the manifest."""
    try:
        manifest = parse_manifest(manifest_text)
    except ManifestError as error:
        return [f"[manifest] {error}"]
    procedures = sources_procedures(sources)
    findings: list[str] = []
    for rule in (registry_findings, dispatch_findings, call_site_findings, resolve_findings):
        findings.extend(rule(manifest, procedures))
    findings.extend(achieved_findings(manifest, sources, procedures))
    findings.extend(unit_findings(manifest, procedures, release_runner))
    findings.extend(classification_findings(manifest, procedures, release_runner))
    findings.extend(version_findings(manifest, sources))
    return findings


# ----------------------------------------------------------------------------
# The repository
# ----------------------------------------------------------------------------
def read_sources(root: Path) -> dict[str, str]:
    sources: dict[str, str] = {}
    for path in sorted((root / "test").glob("*.bas")):
        sources[path.relative_to(root).as_posix()] = path.read_bytes().decode("latin-1").replace("\r\n", "\n")
    return sources


def run(root: str | Path) -> list[str]:
    """Check the repository at ``root``; the evidence policy names the manifest and the release runner."""
    base = Path(root)
    try:
        policy = json.loads((base / ".github" / "excel-evidence-policy.json").read_text(encoding="utf-8"))
        manifest_path = policy["case_manifest"]
        release_runner = policy["certification_runner"]
    except (OSError, ValueError, KeyError, TypeError) as error:
        return [f"[manifest] the evidence policy does not name case_manifest and certification_runner: {error}"]
    try:
        manifest_text = (base / manifest_path).read_text(encoding="utf-8")
    except OSError as error:
        return [f"[manifest] cannot read {manifest_path}: {error.strerror or error}"]
    return findings_for(manifest_text, read_sources(base), str(release_runner))


# ----------------------------------------------------------------------------
# Fixtures
# ----------------------------------------------------------------------------
FIXTURE_MANIFEST: dict[str, Any] = {
    "schema_version": 1,
    "manifest_version": 3,
    "description": "A fixture manifest.",
    "units": [
        {"name": "IsoUnit", "execution": "isolated", "cases": ["Test_EXCEL_UI_RunIso"]},
        {"name": "PackUnit", "execution": "live", "runner": "TST_RunPack",
         "cases": ["TST_Case_A", "TST_Case_B"]},
        {"name": "SingleUnit", "execution": "live", "cases": ["Test_EXCEL_UI_RunSingle"]},
    ],
    "required_cases": [{"case": "TST_Case_B", "note": "named by the issue"}],
    "focused_cases": {"TST_Case_Focused": "run by one focused pack only"},
    "interactive_runners": {"Test_EXCEL_UI_RunAll": "an interactive pack"},
    "achieved_state": {
        "unit": "PackUnit",
        "description": "The observations of the fixture.",
        "observations": [
            {"element": "Alpha", "scope": "Here"},
            {"element": "Beta", "scope": "There", "may_be_unverifiable": True},
        ],
    },
}

FIXTURE_SOURCE = """Attribute VB_Name = "M_FIXTURE"
Option Explicit
Public Const TEST_CERT_MANIFEST_VERSION As Long = 3

Public Sub Test_EXCEL_UI_RunReleaseCertification()
    TST_CertRunIsolatedUnit "IsoUnit", Iso
    TST_CertRunUnit "PackUnit"
    TST_CertRunUnit "SingleUnit"
    TST_CertRunUnit "SelfTest.ForcedFailure"
End Sub

Private Sub TST_CertExecuteUnit(ByVal UnitName As String)
    Select Case UnitName
        Case "IsoUnit"
            TST_CertRunCase "Test_EXCEL_UI_RunIso", True
        Case "PackUnit"
            TST_RunPack IncludeTitleBarTests:=True, _
                CallerProc:="x"
        Case "SingleUnit"
            TST_CertRunCase "Test_EXCEL_UI_RunSingle", True
        Case "SelfTest.ForcedFailure"
            TST_CertRunCase "SelfTest.ForcedCase", True
        Case Else
    End Select
End Sub

Private Sub TST_RunPack(ByVal IncludeTitleBarTests As Boolean, ByVal CallerProc As String)
    TST_CertRunCase "TST_Case_A", IncludeTitleBarTests
    TST_CertRunCase "TST_Case_B", IncludeTitleBarTests
    TST_Observe "Alpha", "Here"
    TST_Observe "Beta", "There" ' "Gamma"
End Sub

Private Sub TST_Observe(ByVal Element As String, ByVal Scope As String)
End Sub

Private Function TST_CertManifestCase(ByVal Index As Long, ByRef UnitName As String, ByRef CaseId As String) As Boolean
    Select Case Index
        Case 1
            UnitName = "IsoUnit"
            CaseId = "Test_EXCEL_UI_RunIso"
        Case 2
            UnitName = "PackUnit"
            CaseId = "TST_Case_A"
        Case 3
            UnitName = "PackUnit"
            CaseId = "TST_Case_B"
        Case 4
            UnitName = "SingleUnit"
            CaseId = "Test_EXCEL_UI_RunSingle"
        Case Else
            Exit Function
    End Select
    TST_CertManifestCase = True
End Function

Private Sub TST_CertDispatchCase(ByVal CaseId As String, ByVal IncludeTitleBarTests As Boolean)
    Select Case CaseId
        Case "Test_EXCEL_UI_RunIso"
            Test_EXCEL_UI_RunIso
        Case "TST_Case_A"
            TST_Case_A IncludeTitleBarTests
        Case "TST_Case_B"
            TST_Case_B
        Case "Test_EXCEL_UI_RunSingle"
            Test_EXCEL_UI_RunSingle
        Case Else
            Err.Raise 5
    End Select
End Sub

Public Sub TST_CertRunCase(ByVal CaseId As String, ByVal IncludeTitleBarTests As Boolean)
    TST_CertDispatchCase CaseId, IncludeTitleBarTests
End Sub

Public Sub Test_EXCEL_UI_RunIso()
End Sub

Public Sub Test_EXCEL_UI_RunSingle()
End Sub

Public Sub Test_EXCEL_UI_RunAll()
End Sub

Private Sub TST_Case_A(ByVal IncludeTitleBarTests As Boolean)
End Sub

Public Sub TST_Case_B()
End Sub

Private Sub TST_Case_Focused()
End Sub
"""


def fixture_findings(manifest: dict[str, Any] | None = None, source: str | None = None) -> list[str]:
    text = json.dumps(FIXTURE_MANIFEST if manifest is None else manifest)
    return findings_for(text, {"test/M_FIXTURE.bas": (FIXTURE_SOURCE if source is None else source)},
                        "Test_EXCEL_UI_RunReleaseCertification")


def mutated(manifest: dict[str, Any]) -> dict[str, Any]:
    copy: dict[str, Any] = json.loads(json.dumps(manifest))
    return copy


def source_cases() -> list[tuple[str, str, str]]:
    """(name, mutated source, rule the gate must report) for each defect the issue names and more."""
    base = FIXTURE_SOURCE
    registry_b = "        Case 3\n            UnitName = \"PackUnit\"\n            CaseId = \"TST_Case_B\"\n"
    dispatch_b = "        Case \"TST_Case_B\"\n            TST_Case_B\n"
    new_case = "Public Sub TST_Case_C()\nEnd Sub\n"
    return [
        ("a deleted registry entry", base.replace(registry_b, "").replace("Case 4\n", "Case 3\n"), "registry"),
        ("a deleted dispatch label", base.replace(dispatch_b, ""), "dispatch"),
        ("a duplicated dispatch label", base.replace(dispatch_b, dispatch_b + dispatch_b), "dispatch"),
        ("a dispatch branch that calls another procedure",
         base.replace("            TST_Case_B\n", "            TST_Case_Focused\n"), "dispatch"),
        ("a case that is never run through the wrapper",
         base.replace('    TST_CertRunCase "TST_Case_B", IncludeTitleBarTests\n', ""), "call-site"),
        ("a case run twice", base.replace('    TST_CertRunCase "TST_Case_B", IncludeTitleBarTests\n',
                                          '    TST_CertRunCase "TST_Case_B", IncludeTitleBarTests\n'
                                          '    TST_CertRunCase "TST_Case_B", IncludeTitleBarTests\n'), "call-site"),
        ("a pack that runs its cases in another order",
         base.replace('    TST_CertRunCase "TST_Case_A", IncludeTitleBarTests\n    TST_CertRunCase "TST_Case_B", '
                      'IncludeTitleBarTests\n', '    TST_CertRunCase "TST_Case_B", IncludeTitleBarTests\n'
                      '    TST_CertRunCase "TST_Case_A", IncludeTitleBarTests\n'), "call-site"),
        ("an unmanifested mandatory case",
         base.replace('    TST_CertRunCase "TST_Case_B", IncludeTitleBarTests\n',
                      '    TST_CertRunCase "TST_Case_B", IncludeTitleBarTests\n'
                      '    TST_CertRunCase "TST_Case_C", IncludeTitleBarTests\n')
         .replace("        Case Else\n            Err.Raise 5\n",
                  '        Case "TST_Case_C"\n            TST_Case_C\n        Case Else\n            Err.Raise 5\n')
         + new_case, "call-site"),
        ("an unmanifested mandatory case in the dispatcher only",
         base.replace("        Case Else\n            Err.Raise 5\n",
                      '        Case "TST_Case_C"\n            TST_Case_C\n        Case Else\n            Err.Raise 5\n')
         + new_case, "dispatch"),
        ("a case procedure nobody classified", base + new_case, "classification"),
        ("a public runner nobody classified", base + "Public Sub Test_EXCEL_UI_RunNew()\nEnd Sub\n", "classification"),
        ("units run in another order",
         base.replace('    TST_CertRunUnit "PackUnit"\n    TST_CertRunUnit "SingleUnit"\n',
                      '    TST_CertRunUnit "SingleUnit"\n    TST_CertRunUnit "PackUnit"\n'), "unit"),
        ("a unit the dispatcher does not run",
         base.replace('        Case "SingleUnit"\n', '        Case "OtherUnit"\n'), "unit"),
        ("a manifest case that is not defined",
         base.replace("Public Sub Test_EXCEL_UI_RunSingle()\nEnd Sub\n", ""), "resolve"),
        ("a registered version that differs",
         base.replace("TEST_CERT_MANIFEST_VERSION As Long = 3", "TEST_CERT_MANIFEST_VERSION As Long = 2"), "version"),
        ("an achieved element the VBA no longer names", base.replace('TST_Observe "Alpha", "Here"', ""), "achieved"),
        ("an achieved scope the VBA no longer names",
         base.replace('TST_Observe "Beta", "There"', 'TST_Observe "Beta", "Elsewhere"'), "achieved"),
        ("an achieved scope that only a comment names",
         base.replace('TST_Observe "Beta", "There" \' "Gamma"', 'TST_Observe "Beta" \' "There"'), "achieved"),
    ]


def manifest_cases() -> list[tuple[str, dict[str, Any], str]]:
    """(name, mutated manifest, rule the gate must report)."""
    def change(edit: Any) -> dict[str, Any]:
        document = mutated(FIXTURE_MANIFEST)
        edit(document)
        return document

    return [
        ("a required case removed from every unit",
         change(lambda doc: doc["units"][1].update(cases=["TST_Case_A"])), "manifest"),
        ("a case listed twice",
         change(lambda doc: doc["units"][1].update(cases=["TST_Case_A", "TST_Case_A", "TST_Case_B"])), "manifest"),
        ("a unit without a runner that lists two cases",
         change(lambda doc: doc["units"][2].update(cases=["Test_EXCEL_UI_RunSingle", "TST_Case_Focused"])),
         "manifest"),
        ("an isolated unit after a live unit",
         change(lambda doc: doc["units"].append({"name": "Late", "execution": "isolated",
                                                 "cases": ["Test_EXCEL_UI_RunLate"]})), "manifest"),
        ("a focused case that is also mandatory",
         change(lambda doc: doc["focused_cases"].update({"TST_Case_A": "no"})), "manifest"),
        ("a manifest version that is not a number",
         change(lambda doc: doc.update(manifest_version="3")), "manifest"),
        ("an unknown key", change(lambda doc: doc.update(extra=1)), "manifest"),
        ("a registry that lacks a unit the manifest adds",
         change(lambda doc: doc["units"].append({"name": "More", "execution": "live",
                                                 "cases": ["Test_EXCEL_UI_RunMore"]})), "registry"),
        ("an achieved_state section for a unit the manifest does not list",
         change(lambda doc: doc["achieved_state"].update(unit="NoSuchUnit")), "manifest"),
        ("an achieved observation listed twice",
         change(lambda doc: doc["achieved_state"]["observations"].append({"element": "Alpha", "scope": "Here"})),
         "manifest"),
        ("an achieved observation with an unknown key",
         change(lambda doc: doc["achieved_state"]["observations"][0].update(extra=1)), "manifest"),
        ("an achieved section without observations",
         change(lambda doc: doc["achieved_state"].update(observations=[])), "manifest"),
    ]


def selftest() -> tuple[list[str], int]:
    """Prove the rules against a passing fixture and each defect; return (errors, fixtures run)."""
    errors: list[str] = []
    count = 1
    clean = fixture_findings()
    if clean:
        errors.append(f"certification manifest fixture: the consistent fixture reports {clean}")
    for name, source, rule in source_cases():
        count += 1
        found = fixture_findings(source=source)
        if not any(f"[{rule}]" in finding for finding in found):
            errors.append(f"certification manifest fixture {name}: expected a [{rule}] finding, got {found}")
    for name, manifest, rule in manifest_cases():
        count += 1
        found = fixture_findings(manifest=manifest)
        if not any(f"[{rule}]" in finding for finding in found):
            errors.append(f"certification manifest fixture {name}: expected a [{rule}] finding, got {found}")
    return errors, count


def main(arguments: list[str] | None = None) -> int:
    import argparse
    import sys

    parser = argparse.ArgumentParser(description="Check the certification manifest against the VBA test modules.")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1],
                        help="the repository (default: this checkout)")
    options = parser.parse_args(arguments)
    errors, count = selftest()
    findings = [*errors, *run(options.root)]
    for finding in findings:
        print(finding, file=sys.stderr)
    print(f"certification manifest: {count} fixtures, {len(findings)} finding(s)")
    return 1 if findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
