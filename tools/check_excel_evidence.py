#!/usr/bin/env python3
"""Validate an Excel certification bundle against one commit.

tools/Certify-EXCEL_UI.ps1 builds a development workbook from a commit, runs
the certification runner in Excel and collects the result in a bundle folder
(docs/EXCEL_CERTIFICATION.md). This check reads that folder and the commit,
never Excel, and reports one PASS or FAIL line per check:

- every file the bundle needs is present;
- the session record names the commit, a clean working tree and a run in which
  every step passed, and the digests it recorded for the collected evidence
  match;
- the source digests equal the SHA-256 of each module of the commit's
  certified configuration, with CRLF line endings as a checkout writes them,
  and the workbook held exactly those modules;
- the certification document reports PASS and COMPLETE with no failed or
  skipped unit, cleanup OK, and the units the commit's certification runner
  runs;
- the RESULT line of the text report agrees with the certification document;
- the host record names Windows and the Excel version, build and bitness, and
  agrees with the host the certification document reports;
- the certification document was written during the recorded session.

The evidence contract (the file names, the steps, the configuration, the
runner, the schema and the host fields) is read from
.github/excel-evidence-policy.json, or from the file --policy names; the
policy is validated before any bundle is read.

With --inventory it prints the source digests a bundle needs for the commit.

The check binds the files to each other and to the commit. It cannot prove
that Excel produced them; reviewing the report remains the operator's job.

Exit codes: 0 every check passed, 1 a check failed, 2 the check could not run
(no Git, an unknown commit, no bundle folder, or an unusable policy).
"""

from __future__ import annotations

import argparse
import functools
import hashlib
import json
import re
import subprocess
import sys
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, NoReturn

ROOT = Path(__file__).resolve().parents[1]
POLICY = ROOT / ".github" / "excel-evidence-policy.json"
POLICY_SCHEMA = 1

# VBComponent.Type: standard module, class module, UserForm, document module.
STANDARD_MODULE = 1
DOCUMENT_MODULE = 100

SHA = re.compile(r"[0-9a-f]{40}")
DIGEST = re.compile(r"[0-9a-f]{64}")
VB_NAME = re.compile(rb'^Attribute VB_Name = "([A-Za-z]\w*)"\r?$', re.MULTILINE)
UNIT_CALL = re.compile(r'^[ \t]*TST_CertRunUnit[ \t]+"([^"]+)"[ \t]*(?:\'.*)?$', re.MULTILINE)
RESULT_LINE = re.compile(r"^RESULT: .*$", re.MULTILINE)
UNIT_LINE = re.compile(r"^  (PASS|FAIL)  (\S+)", re.MULTILINE)
LONE_LF = re.compile(rb"(?<!\r)\n")


class EvidenceError(ValueError):
    """A bundle or commit that does not satisfy a check."""


class SetupError(RuntimeError):
    """A condition that stops the check from running at all."""


class PolicyError(SetupError):
    """An evidence policy the check cannot use."""


def require(condition: object, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


# ----------------------------------------------------------------------------
# The policy
# ----------------------------------------------------------------------------
# The roles of the bundle files; the policy names the file for each.
ROLES = ("session", "host", "source_digests", "certification_json", "certification_report")
HOST_SECTIONS = ("windows", "excel")
# The host fields the PASS line of the host check prints.
REPORTED_HOST_FIELDS = {"windows": ("caption", "version", "build"), "excel": ("version", "build", "bitness")}
POLICY_KEYS = ("description", "schema_version", "evidence_schema", "component", "manifest", "configuration",
               "certification_module", "certification_runner", "required_files", "evidence_files", "steps",
               "host_fields", "excel_matches_certification", "bitness")

STEP_NAME = re.compile(r"[a-z][a-z0-9]*(?:-[a-z0-9]+)*")
FILE_NAME = re.compile(r"[A-Za-z0-9][A-Za-z0-9._-]*")
REPOSITORY_PATH = re.compile(r"[A-Za-z0-9._-]+(?:/[A-Za-z0-9._-]+)*")
IDENTIFIER = re.compile(r"[A-Za-z][A-Za-z0-9_]*")


@dataclass(frozen=True)
class Policy:
    """The evidence contract that .github/excel-evidence-policy.json records."""

    path: Path
    schema: int
    component: str
    manifest: str
    configuration: str
    certification_module: str
    certification_runner: str
    files: dict[str, str]
    evidence_files: tuple[str, ...]
    steps: tuple[str, ...]
    host_fields: dict[str, tuple[str, ...]]
    excel_matches_certification: dict[str, str]
    bitness: tuple[str, ...]

    @property
    def required_files(self) -> tuple[str, ...]:
        return tuple(self.files[role] for role in ROLES)

    @property
    def session(self) -> str:
        return self.files["session"]

    @property
    def host(self) -> str:
        return self.files["host"]

    @property
    def digests(self) -> str:
        return self.files["source_digests"]

    @property
    def certification_json(self) -> str:
        return self.files["certification_json"]

    @property
    def certification_report(self) -> str:
        return self.files["certification_report"]


class PolicyReader:
    """Typed reads of a policy document that raise PolicyError on a fault."""

    def __init__(self, path: Path) -> None:
        self.path = path

    def fail(self, message: str) -> NoReturn:
        raise PolicyError(f"{self.path}: {message}")

    def text(self, value: object, key: str, pattern: re.Pattern[str] | None = None) -> str:
        if not isinstance(value, str) or not value.strip():
            self.fail(f"{key} must be non-empty text")
        if pattern is not None and not pattern.fullmatch(value):
            self.fail(f"{key} {value!r} is not a valid name")
        return value

    def number(self, value: object, key: str) -> int:
        if type(value) is not int or value < 1:
            self.fail(f"{key} must be a positive integer")
        return value

    def names(self, value: object, key: str, pattern: re.Pattern[str]) -> tuple[str, ...]:
        if not isinstance(value, list) or not value:
            self.fail(f"{key} must be a non-empty list")
        items = tuple(self.text(item, f"{key}[{index}]", pattern) for index, item in enumerate(value))
        duplicates = sorted({item for item in items if items.count(item) > 1})
        if duplicates:
            self.fail(f"{key} lists {', '.join(duplicates)} more than once")
        return items

    def table(self, value: object, key: str) -> dict[str, Any]:
        if not isinstance(value, dict) or not value:
            self.fail(f"{key} must be a non-empty object")
        return dict(value)

    def repository_path(self, value: object, key: str) -> str:
        result = self.text(value, key, REPOSITORY_PATH)
        if any(part in (".", "..") for part in result.split("/")):
            self.fail(f"{key} {result!r} must not contain . or .. segments")
        return result


def policy_document(read: PolicyReader, document: object) -> dict[str, Any]:
    """The policy's top-level object, with exactly the known keys."""
    if not isinstance(document, dict):
        read.fail("the policy is not a JSON object")
    missing = [key for key in POLICY_KEYS if key not in document]
    unknown = sorted(key for key in document if key not in POLICY_KEYS)
    if missing or unknown:
        read.fail("; ".join(([f"missing {', '.join(missing)}"] if missing else [])
                            + ([f"unknown {', '.join(unknown)}"] if unknown else [])))
    read.text(document["description"], "description")
    if read.number(document["schema_version"], "schema_version") != POLICY_SCHEMA:
        read.fail(f"schema_version {document['schema_version']} is not supported (expected {POLICY_SCHEMA})")
    return dict(document)


def policy_files(read: PolicyReader, document: dict[str, Any]) -> tuple[dict[str, str], tuple[str, ...]]:
    """The file of each role, and the evidence files among them."""
    files = read.table(document["required_files"], "required_files")
    if sorted(files) != sorted(ROLES):
        read.fail(f"required_files must name exactly the roles {', '.join(ROLES)}")
    named = {role: read.text(files[role], f"required_files.{role}", FILE_NAME) for role in ROLES}
    if len(set(named.values())) != len(named):
        read.fail("required_files names the same file for two roles")
    evidence_files = read.names(document["evidence_files"], "evidence_files", FILE_NAME)
    outside = [name for name in evidence_files if name not in named.values()]
    if outside:
        read.fail(f"evidence_files {', '.join(outside)} are not required files")
    if named["session"] in evidence_files:
        read.fail("evidence_files must not include the session record, which holds their digests")
    return named, evidence_files


def policy_host(read: PolicyReader, document: dict[str, Any]) -> tuple[dict[str, tuple[str, ...]], dict[str, str]]:
    """The host fields of each section, and the Excel fields paired with the certification document."""
    host = read.table(document["host_fields"], "host_fields")
    if sorted(host) != sorted(HOST_SECTIONS):
        read.fail(f"host_fields must have exactly the sections {', '.join(HOST_SECTIONS)}")
    fields = {section: read.names(host[section], f"host_fields.{section}", IDENTIFIER) for section in HOST_SECTIONS}
    for section, reported in REPORTED_HOST_FIELDS.items():
        absent = [key for key in reported if key not in fields[section]]
        if absent:
            read.fail(f"host_fields.{section} must include {', '.join(absent)}, which the host check reports")
    matches = read.table(document["excel_matches_certification"], "excel_matches_certification")
    for key in matches:
        if key not in fields["excel"]:
            read.fail(f"excel_matches_certification: {key!r} is not in host_fields.excel")
    pairs = {key: read.text(theirs, f"excel_matches_certification.{key}", IDENTIFIER)
             for key, theirs in matches.items()}
    return fields, pairs


def parse_policy(document: object, path: Path) -> Policy:
    """Validate a decoded policy document; raise PolicyError on the first fault."""
    read = PolicyReader(path)
    checked = policy_document(read, document)
    files, evidence_files = policy_files(read, checked)
    host_fields, matches = policy_host(read, checked)
    return Policy(
        path=path,
        schema=read.number(checked["evidence_schema"], "evidence_schema"),
        component=read.text(checked["component"], "component"),
        manifest=read.repository_path(checked["manifest"], "manifest"),
        configuration=read.text(checked["configuration"], "configuration", IDENTIFIER),
        certification_module=read.repository_path(checked["certification_module"], "certification_module"),
        certification_runner=read.text(checked["certification_runner"], "certification_runner", IDENTIFIER),
        files=files,
        evidence_files=evidence_files,
        steps=read.names(checked["steps"], "steps", STEP_NAME),
        host_fields=host_fields,
        excel_matches_certification=matches,
        bitness=read.names(checked["bitness"], "bitness", IDENTIFIER),
    )


def load_policy(path: Path) -> Policy:
    """Read and validate the policy at ``path``; raise PolicyError if it is unusable."""

    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise PolicyError(f"{path}: duplicate key {key!r}")
            result[key] = value
        return result

    try:
        document = json.loads(path.read_bytes().decode("utf-8-sig"), object_pairs_hook=unique)
    except OSError as error:
        raise PolicyError(f"cannot read the evidence policy {path}: {error.strerror or error}") from error
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PolicyError(f"{path} is not valid UTF-8 JSON: {error}") from error
    return parse_policy(document, path)


@functools.cache
def default_policy() -> Policy:
    """The policy of the repository this tool belongs to."""
    return load_policy(POLICY)


# ----------------------------------------------------------------------------
# The commit
# ----------------------------------------------------------------------------
def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    try:
        return subprocess.run(["git", "-C", str(root), *args], capture_output=True, check=False)
    except OSError as error:
        raise SetupError(f"cannot run git: {error}") from error


def resolve_commit(root: Path, revision: str) -> str:
    """Return the full SHA of ``revision``, or raise SetupError."""
    result = git(root, "rev-parse", "--verify", "--quiet", f"{revision}^{{commit}}")
    sha = result.stdout.decode("ascii", errors="replace").strip()
    if result.returncode or not SHA.fullmatch(sha):
        raise SetupError(f"{revision!r} is not a commit of the repository at {root}")
    return sha


def committed(root: Path, sha: str, path: str) -> bytes:
    """The bytes Git stores for ``path`` at ``sha``."""
    result = git(root, "show", f"{sha}:{path}")
    require(result.returncode == 0, f"commit {sha[:12]} has no {path}")
    return result.stdout


def checked_out(blob: bytes) -> bytes:
    """A module as a checkout writes it: .gitattributes gives *.bas eol=crlf.

    Git adds a carriage return before every line feed that lacks one, so a
    stored CRLF is left alone; this does the same.
    """
    return LONE_LF.sub(b"\r\n", blob)


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def configuration_modules(root: Path, sha: str, policy: Policy) -> list[str]:
    """The module paths of the policy's configuration at ``sha``."""
    manifest_path, configuration = policy.manifest, policy.configuration
    try:
        manifest = json.loads(committed(root, sha, manifest_path).decode("utf-8-sig"))
        modules = manifest["configurations"][configuration]["modules"]
    except (UnicodeDecodeError, json.JSONDecodeError, KeyError, TypeError) as error:
        raise EvidenceError(f"{manifest_path} at {sha[:12]} has no {configuration} module list: {error}") from error
    require(isinstance(modules, list) and modules and all(isinstance(path, str) and path for path in modules),
            f"{manifest_path} at {sha[:12]}: the {configuration} module list is empty or invalid")
    return list(modules)


def module_name(path: str, source: bytes) -> str:
    names = VB_NAME.findall(source)
    require(len(names) == 1, f"{path}: expected one Attribute VB_Name line, found {len(names)}")
    return str(names[0].decode("ascii"))


def source_inventory(root: Path, sha: str, policy: Policy | None = None) -> dict[str, Any]:
    """The source digests a bundle for ``sha`` must record."""
    policy = policy or default_policy()
    modules = []
    for path in configuration_modules(root, sha, policy):
        blob = committed(root, sha, path)
        modules.append({"path": path, "name": module_name(path, blob), "sha256": sha256(checked_out(blob))})
    return {"schema": policy.schema, "commit": sha, "configuration": policy.configuration, "line_endings": "CRLF",
            "modules": modules}


def certification_units(root: Path, sha: str, policy: Policy | None = None) -> list[str]:
    """The units the certification runner at ``sha`` runs, in order."""
    policy = policy or default_policy()
    module, runner = policy.certification_module, policy.certification_runner
    text = committed(root, sha, module).decode("cp1252", errors="replace")
    body = re.search(rf"^Public Sub {re.escape(runner)}\(\)(.*?)^End Sub", text, re.MULTILINE | re.DOTALL)
    if body is None:
        raise EvidenceError(f"{module} at {sha[:12]} has no {runner}")
    units = UNIT_CALL.findall(body.group(1))
    require(units, f"{runner} at {sha[:12]} runs no TST_CertRunUnit unit")
    return list(units)


# ----------------------------------------------------------------------------
# The bundle
# ----------------------------------------------------------------------------
def decode_text(raw: bytes) -> str:
    """Text the tools wrote as UTF-8, or the workbook wrote in the ANSI code page."""
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp1252", errors="replace")


def parse_json(raw: bytes, label: str) -> dict[str, Any]:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"{label}: duplicate key {key!r}")
            result[key] = value
        return result

    try:
        value = json.loads(decode_text(raw), object_pairs_hook=unique)
    except json.JSONDecodeError as error:
        raise EvidenceError(f"{label} is not valid JSON: {error}") from error
    require(isinstance(value, dict), f"{label} is not a JSON object")
    return dict(value)


def field(document: dict[str, Any], key: str, label: str) -> Any:
    require(key in document, f"{label} has no {key!r}")
    return document[key]


def text_field(document: dict[str, Any], key: str, label: str) -> str:
    value = field(document, key, label)
    require(isinstance(value, str) and value.strip(), f"{label}: {key!r} must be non-empty text")
    return str(value)


def count_field(document: dict[str, Any], key: str, label: str) -> int:
    value = field(document, key, label)
    require(type(value) is int and value >= 0, f"{label}: {key!r} must be a count")
    return int(value)


def timestamp(value: str, label: str) -> datetime:
    try:
        stamp = datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} is not an ISO 8601 time: {value!r}") from error
    require(stamp.utcoffset() is not None, f"{label} has no UTC offset: {value!r}")
    return stamp


@dataclass
class Bundle:
    folder: Path
    sha: str
    root: Path
    policy: Policy
    raw: dict[str, bytes]
    documents: dict[str, dict[str, Any] | None]
    errors: dict[str, str]

    def document(self, name: str) -> dict[str, Any]:
        value = self.documents.get(name)
        if value is None:
            raise EvidenceError(f"cannot check: {self.errors.get(name, name + ' is missing')}")
        return value

    def text(self, name: str) -> str:
        require(name in self.raw, f"cannot check: {name} is missing")
        return decode_text(self.raw[name]).replace("\r\n", "\n")


def load_bundle(folder: Path, sha: str, root: Path, policy: Policy) -> Bundle:
    raw: dict[str, bytes] = {}
    documents: dict[str, dict[str, Any] | None] = {}
    errors: dict[str, str] = {}
    for name in policy.required_files:
        path = folder / name
        if path.is_symlink() or not path.is_file():
            errors[name] = f"{name} is missing"
            continue
        raw[name] = path.read_bytes()
        if name.endswith(".json"):
            try:
                documents[name] = parse_json(raw[name], name)
            except EvidenceError as error:
                errors[name] = str(error)
    return Bundle(folder, sha, root, policy, raw, documents, errors)


# ----------------------------------------------------------------------------
# Checks: each returns a detail for its PASS line or raises EvidenceError
# ----------------------------------------------------------------------------
def check_files(bundle: Bundle) -> str:
    required = bundle.policy.required_files
    missing = [name for name in required if name not in bundle.raw]
    require(not missing, "missing: " + ", ".join(missing))
    return ", ".join(required)


def check_session(bundle: Bundle) -> str:
    policy = bundle.policy
    label = policy.session
    session = bundle.document(label)
    require(field(session, "schema", label) == policy.schema, f"{label}: unsupported schema")
    commit = field(session, "commit", label)
    require(commit == bundle.sha, f"{label} records commit {commit!r}, not {bundle.sha}")
    require(field(session, "dirty", label) is False,
            f"{label}: the run was made from a working tree with changes (-AllowDirty)")
    steps = field(session, "steps", label)
    require(isinstance(steps, list) and all(isinstance(step, dict) for step in steps),
            f"{label}: steps must be a list of objects")
    names = [step.get("name") for step in steps]
    require(names == list(policy.steps), f"{label}: steps are {names}, not {list(policy.steps)}")
    failed = [f"{step['name']}={step.get('status')} ({step.get('detail', '')})"
              for step in steps if step.get("status") != "PASS"]
    require(not failed, f"{label}: steps not passed: " + "; ".join(failed))
    result = field(session, "result", label)
    exit_code = field(session, "exit_code", label)
    require(result == "PASS" and exit_code == 0,
            f"{label}: the run ended {result!r} with exit code {exit_code!r}")
    started = timestamp(text_field(session, "started_utc", label), "started_utc")
    finished = timestamp(text_field(session, "finished_utc", label), "finished_utc")
    require(started <= finished, f"{label}: finished_utc precedes started_utc")
    return f"commit {bundle.sha[:12]}, clean tree, {len(policy.steps)} steps PASS"


def check_evidence_digests(bundle: Bundle) -> str:
    policy = bundle.policy
    label = policy.session
    session = bundle.document(label)
    entries = field(session, "evidence", label)
    require(isinstance(entries, list) and all(isinstance(item, dict) for item in entries),
            f"{label}: evidence must be a list of objects")
    recorded = {item.get("file"): item.get("sha256") for item in entries}
    wanted = policy.evidence_files
    require(sorted(recorded) == sorted(wanted) and len(entries) == len(wanted),
            f"{label}: evidence lists {sorted(map(str, recorded))}, not {sorted(wanted)}")
    for name in wanted:
        require(name in bundle.raw, f"cannot check: {name} is missing")
        require(recorded[name] == sha256(bundle.raw[name]),
                f"{name} differs from the copy the run collected (SHA-256 mismatch)")
    return " and ".join(wanted) + " as collected"


def check_sources(bundle: Bundle) -> str:
    policy = bundle.policy
    label = policy.digests
    digests = bundle.document(label)
    expected = source_inventory(bundle.root, bundle.sha, policy)
    for key in ("schema", "commit", "configuration", "line_endings"):
        require(digests.get(key) == expected[key],
                f"{label}: {key} is {digests.get(key)!r}, not {expected[key]!r}")
    modules = field(digests, "modules", label)
    require(isinstance(modules, list) and all(isinstance(item, dict) for item in modules),
            f"{label}: modules must be a list of objects")
    recorded = [(item.get("path"), item.get("name"), item.get("sha256")) for item in modules]
    wanted = [(item["path"], item["name"], item["sha256"]) for item in expected["modules"]]
    require([item[0] for item in recorded] == [item[0] for item in wanted],
            f"{label}: modules {[item[0] for item in recorded]} are not the commit's "
            f"{policy.configuration} modules {[item[0] for item in wanted]}")
    for (path, name, digest), (_, want_name, want_digest) in zip(recorded, wanted, strict=True):
        require(name == want_name, f"{label}: {path} is named {name!r}, not {want_name!r}")
        require(digest == want_digest, f"{label}: {path} digest differs from the commit's source")
    for item in modules:
        require(item.get("checked_out_sha256") == item.get("sha256"),
                f"{label}: the checked-out {item.get('path')} differs from the commit's source")
    return f"{len(wanted)} {policy.configuration} modules match commit {bundle.sha[:12]}"


def check_components(bundle: Bundle) -> str:
    policy = bundle.policy
    session = bundle.document(policy.session)
    components = field(session, "components", policy.session)
    require(isinstance(components, list) and all(isinstance(item, dict) for item in components),
            f"{policy.session}: components must be a list of objects")
    expected = sorted(item["name"] for item in source_inventory(bundle.root, bundle.sha, policy)["modules"])
    standard = sorted(str(item.get("name")) for item in components if item.get("type") == STANDARD_MODULE)
    others = sorted(f"{item.get('name')} (type {item.get('type')})" for item in components
                    if item.get("type") not in (STANDARD_MODULE, DOCUMENT_MODULE))
    require(standard == expected, f"the workbook held standard modules {standard}, not {expected}")
    require(not others, "the workbook held other components: " + ", ".join(others))
    return f"the workbook held exactly the {len(expected)} {policy.configuration} modules"


def certification_line(result: dict[str, Any], label: str) -> str:
    """The RESULT line the certification runner prints for ``result``."""
    passed = field(result, "passed", label)
    complete = field(result, "complete", label)
    require(isinstance(passed, bool) and isinstance(complete, bool),
            f"{label}: passed and complete must be true or false")
    cleanup = field(result, "cleanup", label)
    require(cleanup in ("OK", "FAILED"), f"{label}: cleanup must be OK or FAILED")
    return (f"RESULT: {'PASS' if passed else 'FAIL'} | {'COMPLETE' if complete else 'INCOMPLETE'}"
            f" | units={count_field(result, 'units', label)} failed={count_field(result, 'unitsFailed', label)}"
            f" skipped={count_field(result, 'skipped', label)} cleanup={cleanup}")


def unit_results(result: dict[str, Any], label: str) -> list[tuple[str, bool]]:
    units = field(result, "unitResults", label)
    require(isinstance(units, list) and all(isinstance(item, dict) for item in units),
            f"{label}: unitResults must be a list of objects")
    return [(str(item.get("name")), item.get("passed") is True) for item in units]


def check_verdict(bundle: Bundle) -> str:
    policy = bundle.policy
    label = policy.certification_json
    result = bundle.document(label)
    require(result.get("component") == policy.component and result.get("schema") == policy.schema,
            f"{label}: not a schema {policy.schema} {policy.component} certification document")
    line = certification_line(result, label)
    require(result["passed"] is True and result["complete"] is True,
            f"{label}: the verdict is {line.removeprefix('RESULT: ')}")
    require(result["unitsFailed"] == 0, f"{label}: {result['unitsFailed']} unit(s) failed")
    require(result["skipped"] == 0, f"{label}: {result['skipped']} unit(s) skipped")
    require(result["cleanup"] == "OK", f"{label}: cleanup FAILED: {result.get('cleanupDetail', '')}")
    require(field(result, "skipDetail", label) == [], f"{label}: skipDetail is not empty")
    units = unit_results(result, label)
    expected = certification_units(bundle.root, bundle.sha, policy)
    require([name for name, _ in units] == expected,
            f"{label}: units {[name for name, _ in units]} are not the commit's {expected}")
    require(result["units"] == len(expected), f"{label}: units={result['units']}, not {len(expected)}")
    failed = [name for name, passed in units if not passed]
    require(not failed, f"{label}: failed units: " + ", ".join(failed))
    return line


def check_report(bundle: Bundle) -> str:
    policy = bundle.policy
    label, report_label = policy.certification_json, policy.certification_report
    result = bundle.document(label)
    report = bundle.text(report_label)
    lines = [str(line).rstrip() for line in RESULT_LINE.findall(report)]
    require(len(lines) == 1, f"{report_label}: expected one RESULT line, found {len(lines)}")
    expected = certification_line(result, label)
    require(lines[0] == expected, f"{report_label} says {lines[0]!r}; {label} says {expected!r}")
    listed = [(name, verdict == "PASS") for verdict, name in UNIT_LINE.findall(report)]
    require(listed == unit_results(result, label),
            f"{report_label}: unit lines {listed} differ from {label}")
    return lines[0]


def check_host(bundle: Bundle) -> str:
    policy = bundle.policy
    label = policy.host
    host = bundle.document(label)
    require(host.get("schema") == policy.schema, f"{label}: unsupported schema")
    windows = field(host, "windows", label)
    excel = field(host, "excel", label)
    require(isinstance(windows, dict) and isinstance(excel, dict), f"{label}: windows and excel must be objects")
    for key in policy.host_fields["windows"]:
        text_field(windows, key, f"{label} windows")
    for key in policy.host_fields["excel"]:
        text_field(excel, key, f"{label} excel")
    require(excel["bitness"] in policy.bitness, f"{label}: Excel bitness must be one of {policy.bitness}")
    result = bundle.document(policy.certification_json)
    for key, theirs in policy.excel_matches_certification.items():
        require(excel[key] == result.get(theirs),
                f"{label} excel {key} {excel[key]!r} differs from {policy.certification_json} "
                f"{theirs} {result.get(theirs)!r}")
    return (f"Excel {excel['version']} build {excel['build']} {excel['bitness']}, "
            f"{windows['caption']} {windows['version']} build {windows['build']}")


def check_timing(bundle: Bundle) -> str:
    policy = bundle.policy
    label = policy.certification_json
    session = bundle.document(policy.session)
    result = bundle.document(label)
    written = text_field(result, "timestampLocal", label)
    try:
        stamp = datetime.strptime(written, "%Y-%m-%d %H:%M:%S")
    except ValueError as error:
        raise EvidenceError(f"{label}: timestampLocal {written!r} is not a local time") from error
    # The workbook writes local wall-clock time without an offset.
    started = timestamp(text_field(session, "started_local", policy.session), "started_local").replace(tzinfo=None)
    finished = timestamp(text_field(session, "finished_local", policy.session), "finished_local").replace(tzinfo=None)
    margin = timedelta(seconds=1)
    require(started.replace(microsecond=0) - margin <= stamp <= finished + margin,
            f"{label} was written at {written}, outside the session "
            f"{started:%Y-%m-%d %H:%M:%S} to {finished:%Y-%m-%d %H:%M:%S}")
    return f"written at {written}, during the session"


CHECKS: tuple[tuple[str, Callable[[Bundle], str]], ...] = (
    ("files", check_files),
    ("session", check_session),
    ("evidence digests", check_evidence_digests),
    ("source digests", check_sources),
    ("workbook modules", check_components),
    ("verdict", check_verdict),
    ("report", check_report),
    ("host", check_host),
    ("timing", check_timing),
)


def evaluate(folder: Path, sha: str, root: Path, policy: Policy | None = None) -> list[tuple[str, bool, str]]:
    """Run every check; return (name, passed, detail) for each."""
    bundle = load_bundle(folder, sha, root, policy or default_policy())
    outcomes = []
    for name, check in CHECKS:
        try:
            outcomes.append((name, True, check(bundle)))
        except EvidenceError as error:
            outcomes.append((name, False, str(error)))
    return outcomes


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate an Excel certification bundle against one commit.")
    parser.add_argument("bundle", nargs="?", type=Path, help="the bundle folder Certify-EXCEL_UI.ps1 wrote")
    parser.add_argument("--commit", default="HEAD", help="the commit the bundle must certify (default: HEAD)")
    parser.add_argument("--root", type=Path, default=ROOT, help="the repository (default: this checkout)")
    parser.add_argument("--policy", type=Path, default=POLICY,
                        help="the evidence policy (default: .github/excel-evidence-policy.json of this checkout)")
    parser.add_argument("--inventory", action="store_true",
                        help="print the source digests a bundle needs for the commit, and exit")
    options = parser.parse_args(arguments)
    if options.inventory == (options.bundle is not None):
        parser.error("give a bundle folder, or --inventory")
    try:
        policy = load_policy(options.policy)
        sha = resolve_commit(options.root, options.commit)
        if options.inventory:
            print(json.dumps(source_inventory(options.root, sha, policy), indent=2))
            return 0
        if not options.bundle.is_dir():
            raise SetupError(f"{options.bundle} is not a folder")
        outcomes = evaluate(options.bundle, sha, options.root, policy)
    except (SetupError, EvidenceError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    print(f"Excel certification bundle: {options.bundle}")
    print(f"Commit: {sha}")
    for name, passed, detail in outcomes:
        print(f"  {'PASS' if passed else 'FAIL'}  {name}: {detail}")
    passed = all(ok for _, ok, _ in outcomes)
    print(f"RESULT: {'PASS' if passed else 'FAIL'}")
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
