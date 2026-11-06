#!/usr/bin/env python3
"""Validate an Excel certification bundle against one commit.

tools/Certify-EXCEL_UI.ps1 builds a development workbook from a commit, runs
Test_EXCEL_UI_RunReleaseCertification in Excel and collects the result in a
bundle folder (docs/EXCEL_CERTIFICATION.md). This check reads that folder and
the commit, never Excel, and reports one PASS or FAIL line per check:

- every file the bundle needs is present;
- session.json names the commit, a clean working tree and a run in which every
  step passed, and the digests it recorded for the collected evidence match;
- source-digests.json equals the SHA-256 of each module of the commit's
  development configuration (tools/vba-projects.json), with CRLF line endings
  as a checkout writes them, and the workbook held exactly those modules;
- certification.json reports PASS and COMPLETE with no failed or skipped unit,
  cleanup OK, and the units the commit's certification runner runs;
- the RESULT line of certification.txt agrees with certification.json;
- host.json names Windows and the Excel version, build and bitness, and agrees
  with the host certification.json reports;
- certification.json was written during the recorded session.

With --inventory it prints the source digests a bundle needs for the commit.

The check binds the files to each other and to the commit. It cannot prove
that Excel produced them; reviewing the report remains the operator's job.

Exit codes: 0 every check passed, 1 a check failed, 2 the check could not run
(no Git, an unknown commit, or no bundle folder).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = "tools/vba-projects.json"
CONFIGURATION = "development"
CERTIFICATION_MODULE = "test/M_EXCEL_UI_REGRESSION_TESTS.bas"
CERTIFICATION_RUNNER = "Test_EXCEL_UI_RunReleaseCertification"
COMPONENT = "VBA Excel UI"
SCHEMA = 1

SESSION = "session.json"
HOST = "host.json"
DIGESTS = "source-digests.json"
CERTIFICATION_JSON = "certification.json"
CERTIFICATION_REPORT = "certification.txt"
REQUIRED_FILES = (SESSION, HOST, DIGESTS, CERTIFICATION_JSON, CERTIFICATION_REPORT)
EVIDENCE_FILES = (CERTIFICATION_JSON, CERTIFICATION_REPORT)

# The steps tools/Certify-EXCEL_UI.ps1 records, in the order it runs them.
STEPS = ("preflight", "build", "open", "enable-macros", "compile", "certify", "collect", "close")
BITNESS = ("x64", "x86")
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


def require(condition: object, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


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


def development_modules(root: Path, sha: str) -> list[str]:
    """The module paths of the development configuration at ``sha``."""
    try:
        manifest = json.loads(committed(root, sha, MANIFEST).decode("utf-8-sig"))
        modules = manifest["configurations"][CONFIGURATION]["modules"]
    except (UnicodeDecodeError, json.JSONDecodeError, KeyError, TypeError) as error:
        raise EvidenceError(f"{MANIFEST} at {sha[:12]} has no {CONFIGURATION} module list: {error}") from error
    require(isinstance(modules, list) and modules and all(isinstance(path, str) and path for path in modules),
            f"{MANIFEST} at {sha[:12]}: the {CONFIGURATION} module list is empty or invalid")
    return list(modules)


def module_name(path: str, source: bytes) -> str:
    names = VB_NAME.findall(source)
    require(len(names) == 1, f"{path}: expected one Attribute VB_Name line, found {len(names)}")
    return str(names[0].decode("ascii"))


def source_inventory(root: Path, sha: str) -> dict[str, Any]:
    """The source digests a bundle for ``sha`` must record."""
    modules = []
    for path in development_modules(root, sha):
        blob = committed(root, sha, path)
        modules.append({"path": path, "name": module_name(path, blob), "sha256": sha256(checked_out(blob))})
    return {"schema": SCHEMA, "commit": sha, "configuration": CONFIGURATION, "line_endings": "CRLF",
            "modules": modules}


def certification_units(root: Path, sha: str) -> list[str]:
    """The units the certification runner at ``sha`` runs, in order."""
    text = committed(root, sha, CERTIFICATION_MODULE).decode("cp1252", errors="replace")
    body = re.search(rf"^Public Sub {CERTIFICATION_RUNNER}\(\)(.*?)^End Sub",
                     text, re.MULTILINE | re.DOTALL)
    if body is None:
        raise EvidenceError(f"{CERTIFICATION_MODULE} at {sha[:12]} has no {CERTIFICATION_RUNNER}")
    units = UNIT_CALL.findall(body.group(1))
    require(units, f"{CERTIFICATION_RUNNER} at {sha[:12]} runs no TST_CertRunUnit unit")
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


def load_bundle(folder: Path, sha: str, root: Path) -> Bundle:
    raw: dict[str, bytes] = {}
    documents: dict[str, dict[str, Any] | None] = {}
    errors: dict[str, str] = {}
    for name in REQUIRED_FILES:
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
    return Bundle(folder, sha, root, raw, documents, errors)


# ----------------------------------------------------------------------------
# Checks: each returns a detail for its PASS line or raises EvidenceError
# ----------------------------------------------------------------------------
def check_files(bundle: Bundle) -> str:
    missing = [name for name in REQUIRED_FILES if name not in bundle.raw]
    require(not missing, "missing: " + ", ".join(missing))
    return ", ".join(REQUIRED_FILES)


def check_session(bundle: Bundle) -> str:
    session = bundle.document(SESSION)
    require(field(session, "schema", SESSION) == SCHEMA, f"{SESSION}: unsupported schema")
    commit = field(session, "commit", SESSION)
    require(commit == bundle.sha, f"{SESSION} records commit {commit!r}, not {bundle.sha}")
    require(field(session, "dirty", SESSION) is False,
            f"{SESSION}: the run was made from a working tree with changes (-AllowDirty)")
    steps = field(session, "steps", SESSION)
    require(isinstance(steps, list) and all(isinstance(step, dict) for step in steps),
            f"{SESSION}: steps must be a list of objects")
    names = [step.get("name") for step in steps]
    require(names == list(STEPS), f"{SESSION}: steps are {names}, not {list(STEPS)}")
    failed = [f"{step['name']}={step.get('status')} ({step.get('detail', '')})"
              for step in steps if step.get("status") != "PASS"]
    require(not failed, f"{SESSION}: steps not passed: " + "; ".join(failed))
    result = field(session, "result", SESSION)
    exit_code = field(session, "exit_code", SESSION)
    require(result == "PASS" and exit_code == 0,
            f"{SESSION}: the run ended {result!r} with exit code {exit_code!r}")
    started = timestamp(text_field(session, "started_utc", SESSION), "started_utc")
    finished = timestamp(text_field(session, "finished_utc", SESSION), "finished_utc")
    require(started <= finished, f"{SESSION}: finished_utc precedes started_utc")
    return f"commit {bundle.sha[:12]}, clean tree, {len(STEPS)} steps PASS"


def check_evidence_digests(bundle: Bundle) -> str:
    session = bundle.document(SESSION)
    entries = field(session, "evidence", SESSION)
    require(isinstance(entries, list) and all(isinstance(item, dict) for item in entries),
            f"{SESSION}: evidence must be a list of objects")
    recorded = {item.get("file"): item.get("sha256") for item in entries}
    require(sorted(recorded) == sorted(EVIDENCE_FILES) and len(entries) == len(EVIDENCE_FILES),
            f"{SESSION}: evidence lists {sorted(map(str, recorded))}, not {sorted(EVIDENCE_FILES)}")
    for name in EVIDENCE_FILES:
        require(name in bundle.raw, f"cannot check: {name} is missing")
        require(recorded[name] == sha256(bundle.raw[name]),
                f"{name} differs from the copy the run collected (SHA-256 mismatch)")
    return "certification.json and certification.txt as collected"


def check_sources(bundle: Bundle) -> str:
    digests = bundle.document(DIGESTS)
    expected = source_inventory(bundle.root, bundle.sha)
    for key in ("schema", "commit", "configuration", "line_endings"):
        require(digests.get(key) == expected[key],
                f"{DIGESTS}: {key} is {digests.get(key)!r}, not {expected[key]!r}")
    modules = field(digests, "modules", DIGESTS)
    require(isinstance(modules, list) and all(isinstance(item, dict) for item in modules),
            f"{DIGESTS}: modules must be a list of objects")
    recorded = [(item.get("path"), item.get("name"), item.get("sha256")) for item in modules]
    wanted = [(item["path"], item["name"], item["sha256"]) for item in expected["modules"]]
    require([item[0] for item in recorded] == [item[0] for item in wanted],
            f"{DIGESTS}: modules {[item[0] for item in recorded]} are not the commit's "
            f"{CONFIGURATION} modules {[item[0] for item in wanted]}")
    for (path, name, digest), (_, want_name, want_digest) in zip(recorded, wanted, strict=True):
        require(name == want_name, f"{DIGESTS}: {path} is named {name!r}, not {want_name!r}")
        require(digest == want_digest, f"{DIGESTS}: {path} digest differs from the commit's source")
    for item in modules:
        require(item.get("checked_out_sha256") == item.get("sha256"),
                f"{DIGESTS}: the checked-out {item.get('path')} differs from the commit's source")
    return f"{len(wanted)} {CONFIGURATION} modules match commit {bundle.sha[:12]}"


def check_components(bundle: Bundle) -> str:
    session = bundle.document(SESSION)
    components = field(session, "components", SESSION)
    require(isinstance(components, list) and all(isinstance(item, dict) for item in components),
            f"{SESSION}: components must be a list of objects")
    expected = sorted(item["name"] for item in source_inventory(bundle.root, bundle.sha)["modules"])
    standard = sorted(str(item.get("name")) for item in components if item.get("type") == STANDARD_MODULE)
    others = sorted(f"{item.get('name')} (type {item.get('type')})" for item in components
                    if item.get("type") not in (STANDARD_MODULE, DOCUMENT_MODULE))
    require(standard == expected, f"the workbook held standard modules {standard}, not {expected}")
    require(not others, "the workbook held other components: " + ", ".join(others))
    return f"the workbook held exactly the {len(expected)} {CONFIGURATION} modules"


def certification_line(result: dict[str, Any]) -> str:
    """The RESULT line Test_EXCEL_UI_RunReleaseCertification prints for ``result``."""
    label = CERTIFICATION_JSON
    passed = field(result, "passed", label)
    complete = field(result, "complete", label)
    require(isinstance(passed, bool) and isinstance(complete, bool),
            f"{label}: passed and complete must be true or false")
    cleanup = field(result, "cleanup", label)
    require(cleanup in ("OK", "FAILED"), f"{label}: cleanup must be OK or FAILED")
    return (f"RESULT: {'PASS' if passed else 'FAIL'} | {'COMPLETE' if complete else 'INCOMPLETE'}"
            f" | units={count_field(result, 'units', label)} failed={count_field(result, 'unitsFailed', label)}"
            f" skipped={count_field(result, 'skipped', label)} cleanup={cleanup}")


def unit_results(result: dict[str, Any]) -> list[tuple[str, bool]]:
    units = field(result, "unitResults", CERTIFICATION_JSON)
    require(isinstance(units, list) and all(isinstance(item, dict) for item in units),
            f"{CERTIFICATION_JSON}: unitResults must be a list of objects")
    return [(str(item.get("name")), item.get("passed") is True) for item in units]


def check_verdict(bundle: Bundle) -> str:
    result = bundle.document(CERTIFICATION_JSON)
    label = CERTIFICATION_JSON
    require(result.get("component") == COMPONENT and result.get("schema") == SCHEMA,
            f"{label}: not a schema {SCHEMA} {COMPONENT} certification document")
    line = certification_line(result)
    require(result["passed"] is True and result["complete"] is True,
            f"{label}: the verdict is {line.removeprefix('RESULT: ')}")
    require(result["unitsFailed"] == 0, f"{label}: {result['unitsFailed']} unit(s) failed")
    require(result["skipped"] == 0, f"{label}: {result['skipped']} unit(s) skipped")
    require(result["cleanup"] == "OK", f"{label}: cleanup FAILED: {result.get('cleanupDetail', '')}")
    require(field(result, "skipDetail", label) == [], f"{label}: skipDetail is not empty")
    units = unit_results(result)
    expected = certification_units(bundle.root, bundle.sha)
    require([name for name, _ in units] == expected,
            f"{label}: units {[name for name, _ in units]} are not the commit's {expected}")
    require(result["units"] == len(expected), f"{label}: units={result['units']}, not {len(expected)}")
    failed = [name for name, passed in units if not passed]
    require(not failed, f"{label}: failed units: " + ", ".join(failed))
    return line


def check_report(bundle: Bundle) -> str:
    result = bundle.document(CERTIFICATION_JSON)
    report = bundle.text(CERTIFICATION_REPORT)
    lines = [str(line).rstrip() for line in RESULT_LINE.findall(report)]
    require(len(lines) == 1, f"{CERTIFICATION_REPORT}: expected one RESULT line, found {len(lines)}")
    expected = certification_line(result)
    require(lines[0] == expected, f"{CERTIFICATION_REPORT} says {lines[0]!r}; {CERTIFICATION_JSON} says {expected!r}")
    listed = [(name, verdict == "PASS") for verdict, name in UNIT_LINE.findall(report)]
    require(listed == unit_results(result),
            f"{CERTIFICATION_REPORT}: unit lines {listed} differ from {CERTIFICATION_JSON}")
    return lines[0]


def check_host(bundle: Bundle) -> str:
    host = bundle.document(HOST)
    require(host.get("schema") == SCHEMA, f"{HOST}: unsupported schema")
    windows = field(host, "windows", HOST)
    excel = field(host, "excel", HOST)
    require(isinstance(windows, dict) and isinstance(excel, dict), f"{HOST}: windows and excel must be objects")
    for key in ("caption", "version", "build", "architecture"):
        text_field(windows, key, f"{HOST} windows")
    for key in ("version", "build", "operating_system", "bitness"):
        text_field(excel, key, f"{HOST} excel")
    require(excel["bitness"] in BITNESS, f"{HOST}: Excel bitness must be one of {BITNESS}")
    result = bundle.document(CERTIFICATION_JSON)
    for key, theirs in (("version", "excelVersion"), ("build", "excelBuild"),
                        ("operating_system", "operatingSystem"), ("bitness", "bitness")):
        require(excel[key] == result.get(theirs),
                f"{HOST} excel {key} {excel[key]!r} differs from {CERTIFICATION_JSON} {theirs} {result.get(theirs)!r}")
    return (f"Excel {excel['version']} build {excel['build']} {excel['bitness']}, "
            f"{windows['caption']} {windows['version']} build {windows['build']}")


def check_timing(bundle: Bundle) -> str:
    session = bundle.document(SESSION)
    result = bundle.document(CERTIFICATION_JSON)
    written = text_field(result, "timestampLocal", CERTIFICATION_JSON)
    try:
        stamp = datetime.strptime(written, "%Y-%m-%d %H:%M:%S")
    except ValueError as error:
        raise EvidenceError(f"{CERTIFICATION_JSON}: timestampLocal {written!r} is not a local time") from error
    # The workbook writes local wall-clock time without an offset.
    started = timestamp(text_field(session, "started_local", SESSION), "started_local").replace(tzinfo=None)
    finished = timestamp(text_field(session, "finished_local", SESSION), "finished_local").replace(tzinfo=None)
    margin = timedelta(seconds=1)
    require(started.replace(microsecond=0) - margin <= stamp <= finished + margin,
            f"{CERTIFICATION_JSON} was written at {written}, outside the session "
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


def evaluate(folder: Path, sha: str, root: Path) -> list[tuple[str, bool, str]]:
    """Run every check; return (name, passed, detail) for each."""
    bundle = load_bundle(folder, sha, root)
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
    parser.add_argument("--inventory", action="store_true",
                        help="print the source digests a bundle needs for the commit, and exit")
    options = parser.parse_args(arguments)
    if options.inventory == (options.bundle is not None):
        parser.error("give a bundle folder, or --inventory")
    try:
        sha = resolve_commit(options.root, options.commit)
        if options.inventory:
            print(json.dumps(source_inventory(options.root, sha), indent=2))
            return 0
        if not options.bundle.is_dir():
            raise SetupError(f"{options.bundle} is not a folder")
        outcomes = evaluate(options.bundle, sha, options.root)
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
