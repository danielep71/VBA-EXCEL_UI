#!/usr/bin/env python3
"""
Static release gate for VBA-EXCEL_UI.

Everything here can be decided from the repository checkout and Git index,
without Excel or network access. That is the whole design constraint: a hosted
runner has no Office, so the gate covers what is checkable there and
deliberately does not pretend to replace
Test_EXCEL_UI_RunReleaseCertification, which is the behavioural gate.

Each check is independent and reports every finding it has rather than stopping
at the first, because a contributor fixing one problem per push is a bad use of
their time.

Exit status is 0 when every check passed, 1 otherwise. --help prints the usage
and exits 0 without running a check.
"""

import ast
import fnmatch
import hashlib
import html
import json
import os
import re
import subprocess
import sys
import unicodedata
from collections.abc import Callable, Iterable, Iterator
from datetime import date
from types import ModuleType
from typing import TypeAlias
from urllib.parse import unquote

from vba_lex import code_only

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Every module check iterates the modules Git tracks, not a list written here.
# A hard-coded list let a new module, or a test module split out of the suite,
# skip every module check while the gate still reported PASS. The lists below
# are filled by load_modules() from `git ls-files`, before any check runs.
MODULE_DIRECTORIES = ("src", "test", "demo")

VBA_EXTENSIONS = (".bas", ".cls", ".frm")

PRODUCTION_MODULES: list[str] = []
TEST_MODULES: list[str] = []
DEMO_MODULES: list[str] = []
ALL_MODULES: list[str] = []

# Modules that other checks read by name: the title-bar, WinAPI and
# demo-cleanup regressions and the public API manifest inspect these files
# directly, so they must exist. Which modules make up each configuration is not
# written here: tools/vba-projects.json lists them, and the module-inventory
# and VBA-project-manifest checks hold the tracked modules to it.
REQUIRED_MODULES = [
    "src/M_EXCEL_UI_RUNTIME.bas",
    "src/M_EXCEL_UI_TITLEBAR.bas",
    "src/M_EXCEL_UI_SNAPSHOT.bas",
    "src/M_EXCEL_UI.bas",
    "test/M_EXCEL_UI_REGRESSION_TESTS.bas",
    "demo/M_DEMO_BUILDER.bas",
    "demo/M_EXCEL_UI_DEMO.bas",
]

REQUIRED_FILES = REQUIRED_MODULES + [
    ".gitattributes",
    ".gitignore",
    "README.md",
    "CHANGELOG.md",
    "INSTALLATION.md",
    "CONTRIBUTING.md",
    "SECURITY.md",
    "CODE_OF_CONDUCT.md",
    "LICENSE",
    "VERSION",
    "tools/reformat.py",
    "tools/vba_lex.py",
    "tools/vba_analyze.py",
    "tools/vba_analyze_fixtures.py",
    "tools/titlebar_control_fixtures.py",
    "tools/winapi_capture_fixtures.py",
    "tools/demo_cleanup_fixtures.py",
    "tools/reformat_fixtures.py",
    "tools/vba_api.py",
    "tools/wiki_badges.py",
    "tools/repository_policy.py",
    "tools/public_api_manifest.txt",
    "tools/binary_allowlist.txt",
    "tools/check_issues.py",
    "tools/verify.py",
    "tools/render_docs.py",
    "tools/vba-projects.json",
    ".github/labels.json",
    ".github/excel-evidence-policy.json",
    "docs/GOVERNANCE.md",
    "docs/ISSUE_LABELS.md",
    "docs/REPOSITORY_STRUCTURE.md",
    "docs/VBA_HOUSE_STYLE.md",
]

VERSION_RE = re.compile(r"^[0-9]+\.[0-9]+\.[0-9]+(?:[-+][0-9A-Za-z.\-+]*)?$")


def version_findings(raw: str) -> list[str]:
    """Return every problem with the contents of the root VERSION file."""
    findings: list[str] = []

    if raw != raw.rstrip("\n") + "\n":
        findings.append(
            "VERSION must hold one line ending in a single newline")

    if not VERSION_RE.match(raw.strip()):
        findings.append(
            f"VERSION does not hold a release number: {raw.strip()!r}")

    return findings


VERSION_SELFTEST_CASES = [
    ("a plain release number passes", "1.1.2\n", 0),
    ("a pre-release suffix passes", "1.2.0-rc.1\n", 0),
    ("a missing trailing newline is rejected", "1.1.2", 1),
    ("a doubled trailing newline is rejected", "1.1.2\n\n", 1),
    ("a leading v is rejected", "v1.1.2\n", 1),
    ("a two-part version is rejected", "1.1\n", 1),
    ("an empty file is rejected", "", 2),
    ("prose is rejected", "the next release\n", 1),
]

HOUSE_LABELS = {"Safe_Exit", "Fail_Num", "Err_Handler", "Clean_Exit", "Clean_Fail"}


RULE_RE = re.compile(r"^'(=+|-+)$")

FORBIDDEN_TRACKED_ARTIFACTS = [
    ("~$*", "Microsoft Office owner/lock file"),
    (".~lock.*#", "LibreOffice/OpenOffice lock file"),
    ("*.laccdb", "Microsoft Access lock file"),
    ("*.ldb", "Microsoft Access lock file"),
    ("*.xls", "Excel workbook binary"),
    ("*.xlsx", "Excel workbook package"),
    ("*.xlsm", "macro-enabled Excel workbook"),
    ("*.xlsb", "binary Excel workbook"),
    ("*.xlt", "Excel template binary"),
    ("*.xltx", "Excel template package"),
    ("*.xltm", "macro-enabled Excel template"),
    ("*.xla", "legacy Excel add-in"),
    ("*.xlam", "macro-enabled Excel add-in"),
    ("*.xll", "compiled Excel add-in"),
    ("*.xlw", "Excel workspace file"),
    ("*.docm", "macro-enabled Word document"),
    ("*.dotm", "macro-enabled Word template"),
    ("*.pptm", "macro-enabled PowerPoint presentation"),
    ("*.potm", "macro-enabled PowerPoint template"),
    ("*.ppsm", "macro-enabled PowerPoint slideshow"),
    ("*.accdb", "Microsoft Access database"),
    ("*.accde", "compiled Microsoft Access database"),
    ("*.accdr", "Microsoft Access runtime database"),
    ("*.mdb", "legacy Microsoft Access database"),
    ("*.mde", "compiled legacy Microsoft Access database"),
    ("*.ade", "Microsoft Access project binary"),
]

failures: list[str] = []
notes: list[str] = []


def fail(check: str, detail: str) -> None:
    failures.append(f"{check}: {detail}")


def note(check: str, detail: str) -> None:
    """Record something the reader should see that is not a failure.

    A correctly declared API change still has to be legible in the log. Raising
    it through fail() made the gate reject the very change it had just been
    told about, which trained the reader to regenerate the manifest again
    rather than to read the finding.
    """
    notes.append(f"{check}: {detail}")


def read(path: str) -> bytes:
    with open(os.path.join(REPO, path), "rb") as fh:
        return fh.read()


def read_lines(path: str) -> list[str]:
    return read(path).decode("latin-1").split("\r\n")


# --------------------------------------------------------------------------
def is_vba_module(rel: str) -> bool:
    return rel.lower().endswith(VBA_EXTENSIONS)


def is_standard_module(rel: str) -> bool:
    return rel.lower().endswith(".bas")


def discover_modules(tracked: Iterable[str]) -> tuple[dict[str, list[str]], list[str]]:
    """Sort tracked VBA modules by directory.

    Returns ({directory: [modules]}, misplaced). A module counts only when it
    sits directly in src/, test/ or demo/; anywhere else, including a nested
    directory, it would be imported by nobody and checked by nothing, so it is
    returned as misplaced.
    """
    found: dict[str, list[str]] = {directory: [] for directory in MODULE_DIRECTORIES}
    misplaced: list[str] = []
    for rel in sorted(tracked):
        if not is_vba_module(rel):
            continue
        directory, _, name = rel.partition("/")
        if directory in found and name and "/" not in name:
            found[directory].append(rel)
        else:
            misplaced.append(rel)
    return found, misplaced


def load_modules() -> list[str]:
    """Fill the module lists from the Git index and return misplaced modules.

    When the index cannot be read, git_tracked_files() has already recorded a
    finding and the lists stay empty, so the gate fails rather than passing on
    an empty module set.
    """
    tracked = git_tracked_files()
    found, misplaced = discover_modules(tracked or [])
    PRODUCTION_MODULES[:] = found["src"]
    TEST_MODULES[:] = found["test"]
    DEMO_MODULES[:] = found["demo"]
    ALL_MODULES[:] = PRODUCTION_MODULES + TEST_MODULES + DEMO_MODULES
    return misplaced


INSTALLATION_GUIDE = "INSTALLATION.md"
PRODUCTION_BLOCK_RE = re.compile(
    r"^<!-- generated:modules:production -->$(?P<body>.*?)^<!-- /generated:modules:production -->$",
    re.M | re.S)
TABLE_SOURCE_RE = re.compile(r"^\|[^|\n]*\|\s*`([^`]+)`", re.M)


def documented_import_order() -> list[str] | None:
    """The module paths of the production table of INSTALLATION.md, in order.

    Returns None after recording why the table cannot be read.
    """
    if not os.path.isfile(os.path.join(REPO, INSTALLATION_GUIDE)):
        fail("module-inventory", f"{INSTALLATION_GUIDE} is missing")
        return None
    m = PRODUCTION_BLOCK_RE.search(read(INSTALLATION_GUIDE).decode("utf-8"))
    if m is None:
        fail("module-inventory",
             f"{INSTALLATION_GUIDE}: no generated:modules:production block; the import "
             f"order must be generated from {VBA_PROJECT_MANIFEST} by tools/render_docs.py")
        return None
    return TABLE_SOURCE_RE.findall(m.group("body"))


def check_module_inventory() -> None:
    """Discover the modules every other module check iterates, and audit them.

    Four things fail here: a VBA module tracked outside src/, test/ or demo/;
    a src/ set that disagrees with the production configuration of
    tools/vba-projects.json; an INSTALLATION.md import order that is not the
    order of that configuration; and a production module that tools/vba_api.py
    assigns to no public API section, whose public declarations would
    otherwise never reach the manifest.
    """
    check = "module-inventory"
    for rel in load_modules():
        fail(check,
             f"{rel}: VBA module tracked outside src/, test/ or demo/; it "
             f"would escape every module check")

    discovered = set(PRODUCTION_MODULES)
    manifest = load_vba_project_manifest(check)
    production = configuration_modules(manifest, "production", check) if manifest is not None else None
    if production is not None:
        for rel in sorted(discovered - set(production)):
            fail(check,
                 f"{rel}: production module is not in the production configuration of "
                 f"{VBA_PROJECT_MANIFEST}; add it there in import order and run "
                 f"python3 tools/render_docs.py --write")
        for rel in sorted(set(production) - discovered):
            fail(check,
                 f"{rel}: listed in the production configuration of {VBA_PROJECT_MANIFEST} "
                 f"but not tracked in src/")
        documented = documented_import_order()
        if documented is not None and documented != production:
            fail(check,
                 f"{INSTALLATION_GUIDE}: the production import order is "
                 f"{', '.join(documented) or 'empty'}, but {VBA_PROJECT_MANIFEST} lists "
                 f"{', '.join(production)}; run python3 tools/render_docs.py --write")

    vba_api = import_vba_api()
    if vba_api is None:
        return
    sectioned = {rel for _, modules in vba_api.SECTIONS for rel in modules}
    for rel in sorted(discovered - sectioned):
        fail(check,
             f"{rel}: production module in no public API section of "
             f"tools/vba_api.py; its declarations would never be recorded")


# --------------------------------------------------------------------------
VBA_PROJECT_MANIFEST = "tools/vba-projects.json"


def expected_configurations() -> dict[str, set[str]]:
    """The standard modules each configuration of the manifest must list.

    production is src/; development adds test/; demo adds demo/. A workbook
    tool imports exactly what a configuration lists, so a module the manifest
    forgets is missing from every workbook built from it.
    """
    production = {rel for rel in PRODUCTION_MODULES if is_standard_module(rel)}
    return {
        "production": production,
        "development": production | {rel for rel in TEST_MODULES if is_standard_module(rel)},
        "demo": production | {rel for rel in DEMO_MODULES if is_standard_module(rel)},
    }


def load_vba_project_manifest(check: str = "vba-project-manifest") -> dict[str, object] | None:
    """Parse the manifest, or record why it cannot be read and return None."""
    path = os.path.join(REPO, VBA_PROJECT_MANIFEST)
    if not os.path.isfile(path):
        fail(check, f"{VBA_PROJECT_MANIFEST} is missing")
        return None
    try:
        manifest = json.loads(read(VBA_PROJECT_MANIFEST).decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        fail(check, f"{VBA_PROJECT_MANIFEST} is not valid JSON: {exc}")
        return None
    if not isinstance(manifest, dict):
        fail(check, f"{VBA_PROJECT_MANIFEST} is not a JSON object")
        return None
    return manifest


def configuration_modules(manifest: dict[str, object], name: str,
                          check: str = "vba-project-manifest") -> list[str] | None:
    """The module list of one configuration, or None after recording why not."""
    configurations = manifest.get("configurations")
    entry = configurations.get(name) if isinstance(configurations, dict) else None
    modules = entry.get("modules") if isinstance(entry, dict) else None
    if not isinstance(modules, list) or not all(isinstance(rel, str) for rel in modules):
        fail(check, f"{VBA_PROJECT_MANIFEST}: configuration {name!r} has no list of module paths")
        return None
    return modules


def check_vba_project_manifest() -> None:
    """Keep tools/vba-projects.json equal to the tracked modules.

    The workbook synchronizer, the export comparison and the call checker all
    read their module sets from this manifest. It must parse; each tracked
    standard module of src/, test/ and demo/ must be in a configuration; each
    listed path must be tracked; the three configurations must hold exactly
    the sets of expected_configurations(), each path once; it must list no
    document modules, since this repository has none; and module_descriptions
    must give every listed module a one-line description, which
    tools/render_docs.py writes into the module tables, and describe nothing
    else.
    """
    check = "vba-project-manifest"
    manifest = load_vba_project_manifest()
    if manifest is None:
        return

    if manifest.get("document_modules") != []:
        fail(check, f"{VBA_PROJECT_MANIFEST}: document_modules must be an empty list; "
                    f"this repository has no document modules")
    patterns = manifest.get("project_name_patterns", [])
    for pattern in patterns if isinstance(patterns, list) else [patterns]:
        try:
            re.compile(str(pattern))
        except re.error as exc:
            fail(check, f"{VBA_PROJECT_MANIFEST}: project name pattern {pattern!r} is invalid: {exc}")

    tracked = set(git_tracked_files() or [])
    expected = expected_configurations()
    listed: set[str] = set()
    for name, wanted in expected.items():
        modules = configuration_modules(manifest, name)
        if modules is None:
            continue
        listed.update(modules)
        for rel in sorted({rel for rel in modules if modules.count(rel) > 1}):
            fail(check, f"{VBA_PROJECT_MANIFEST}: {name} lists {rel} more than once")
        for rel in sorted(set(modules) - wanted):
            if not os.path.isfile(os.path.join(REPO, rel)):
                fail(check, f"{VBA_PROJECT_MANIFEST}: {name} lists {rel}, which does not exist")
            elif rel not in tracked:
                fail(check, f"{VBA_PROJECT_MANIFEST}: {name} lists {rel}, which Git does not track")
            else:
                fail(check, f"{VBA_PROJECT_MANIFEST}: {name} lists {rel}, which does not belong to it")
        for rel in sorted(wanted - set(modules)):
            fail(check, f"{VBA_PROJECT_MANIFEST}: {name} does not list {rel}")

    for rel in sorted(set().union(*expected.values()) - listed):
        fail(check, f"{rel}: tracked module is in no configuration of {VBA_PROJECT_MANIFEST}")

    for finding in module_description_findings(manifest.get("module_descriptions"), listed):
        fail(check, f"{VBA_PROJECT_MANIFEST}: {finding}")


def module_description_findings(descriptions: object, listed: set[str]) -> list[str]:
    """Every listed module needs a one-line description, and nothing else has one.

    tools/render_docs.py writes each description into a table cell, so a
    description must hold neither a line break nor a '|'.
    """
    if not isinstance(descriptions, dict):
        return ["module_descriptions must map each module path to its description"]
    findings: list[str] = []
    for rel in sorted(listed):
        text = descriptions.get(rel)
        if not isinstance(text, str) or not text.strip():
            findings.append(f"module_descriptions has no description for {rel}")
        elif "|" in text or "\n" in text:
            findings.append(f"the description of {rel} must be one line without '|', "
                            f"since it becomes a table cell")
    findings.extend(f"module_descriptions describes {rel}, which no configuration lists"
                    for rel in sorted(set(descriptions) - listed))
    return findings


# --------------------------------------------------------------------------
def check_required_files() -> None:
    for rel in REQUIRED_FILES:
        path = os.path.join(REPO, rel)
        if os.path.isdir(path):
            fail("required-files", f"{rel} is a directory, not a file")
        elif not os.path.isfile(path):
            fail("required-files", f"missing {rel}")


# --------------------------------------------------------------------------
def check_module_names() -> None:
    """Attribute VB_Name must match the filename.

    A mismatch imports cleanly and then silently shadows or duplicates a module
    in the VBE, which is far harder to diagnose than a failed check.

    A standard module carries the attribute on line 1. A class or form export
    starts with its VERSION/BEGIN header, so there the first attribute line is
    the one that counts.
    """
    for rel in ALL_MODULES:
        expected = os.path.splitext(os.path.basename(rel))[0]
        text = read(rel).decode("latin-1")
        pattern = r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"'
        if is_standard_module(rel):
            m = re.match(pattern, text)
            where = "on line 1"
        else:
            m = re.search(pattern, text, re.M)
            where = "after the module header"
        if not m:
            fail("module-name", f"{rel}: no Attribute VB_Name {where}")
        elif m.group(1) != expected:
            fail("module-name", f"{rel}: declares {m.group(1)!r}, expected {expected!r}")


# --------------------------------------------------------------------------
def check_option_policy() -> None:
    """Require Option Explicit everywhere and Option Private Module where it applies.

    Option Private Module is a standard-module statement, so it is required of
    the production and test .bas modules only.
    """
    for rel in ALL_MODULES:
        text = read(rel).decode("latin-1")
        if not re.search(r"^Option Explicit\s*$", text, re.M):
            fail("option-policy", f"{rel}: missing Option Explicit")
        if rel in PRODUCTION_MODULES + TEST_MODULES and is_standard_module(rel):
            if not re.search(r"^Option Private Module\s*$", text, re.M):
                fail("option-policy", f"{rel}: missing Option Private Module")


# --------------------------------------------------------------------------
def check_encoding_and_endings() -> None:
    for rel in ALL_MODULES:
        data = read(rel)
        bare_lf = data.count(b"\n") - data.count(b"\r\n")
        if bare_lf:
            fail("line-endings", f"{rel}: {bare_lf} bare LF")
        if b"\t" in data:
            fail("whitespace", f"{rel}: contains tab characters")
        non_ascii = sum(1 for b in data if b > 127)
        if non_ascii:
            fail("encoding", f"{rel}: {non_ascii} non-ASCII bytes")
        for n, line in enumerate(data.decode("latin-1").split("\r\n"), 1):
            if line != line.rstrip():
                fail("whitespace", f"{rel}:{n}: trailing whitespace")
                break


# --------------------------------------------------------------------------
def check_banner_rules() -> None:
    for rel in ALL_MODULES:
        for n, line in enumerate(read_lines(rel), 1):
            if RULE_RE.match(line) and len(line) != 79:
                fail("banner-width", f"{rel}:{n}: rule is {len(line)} cols, expected 79")


# --------------------------------------------------------------------------
def check_house_labels() -> None:
    """Retain the existing style policy separately from procedure resolution."""
    from vba_analyze import logical_lines, statements
    for rel in ALL_MODULES:
        for n, raw in logical_lines(read(rel).decode('ascii')):
            for statement in statements(raw):
                code = code_only(statement).strip()
                if re.fullmatch(r'\w+:', code) and code[:-1] not in HOUSE_LABELS:
                    fail('house-labels', f'{rel}:{n}: non-house label {code[:-1]!r}')


def check_vba_analyzer() -> None:
    from pathlib import Path

    from vba_analyze import run
    findings, inventory = run(Path(REPO), ALL_MODULES)
    for finding in findings:
        fail('VBA analyzer', str(finding))
    for entry in inventory:
        print('  INVENTORY ' + entry)


def check_vba_analyzer_selftest() -> None:
    from vba_analyze_fixtures import selftest
    findings, count = selftest()
    for finding in findings:
        fail('VBA analyzer self-test', finding)
    print(f'  {count} VBA analyzer fixtures')


def check_titlebar_control() -> None:
    from titlebar_control_fixtures import selftest
    for finding in selftest(REPO):
        fail('title-bar control regression', finding)


def check_winapi_error_capture() -> None:
    """Hold every WinAPI call in src/ and test/ to the #27 capture sequence."""
    from winapi_capture_fixtures import run, selftest
    errors, count = selftest()
    for finding in errors:
        fail('WinAPI error capture', finding)
    for finding in run(REPO, ALL_MODULES):
        fail('WinAPI error capture', finding)
    print(f'  {count} WinAPI error-capture fixtures')


def check_demo_cleanup() -> None:
    from demo_cleanup_fixtures import selftest
    findings, count = selftest(REPO)
    for finding in findings:
        fail('demo cleanup regression', finding)
    print(f'  {count} demo cleanup source fixtures')


# --------------------------------------------------------------------------
def import_vba_api() -> ModuleType | None:
    sys.path.insert(0, os.path.join(REPO, "tools"))
    try:
        import vba_api
    except Exception as exc:                                  # pragma: no cover
        fail("public-api", f"tools/vba_api.py could not be imported: {exc}")
        return None
    return vba_api


def check_public_api() -> None:
    """Diff the full declaration contract against the versioned manifest.

    Recording only module, kind and name detected a member appearing or
    disappearing and nothing else, so a reordered parameter, a ByVal turned
    ByRef, a changed default or a renumbered enum could break every caller
    while this check stayed green. The manifest now carries the whole
    normalised declaration, and any contract difference is a finding.

    The two sections are diffed separately so the report says which promise
    was broken: a [supported] change is a Semantic Versioning event for
    external callers, while a [project-public] change only has to be
    deliberate.
    """
    vba_api = import_vba_api()
    if vba_api is None:
        return

    manifest_path = os.path.join(REPO, "tools/public_api_manifest.txt")
    if not os.path.exists(manifest_path):
        fail("public-api", "tools/public_api_manifest.txt is missing")
        return

    try:
        recorded, baseline_version = vba_api.parse_manifest(manifest_path)
    except vba_api.ApiError as exc:
        fail("public-api", f"tools/public_api_manifest.txt: {exc}")
        return

    actual, findings = vba_api.surface(REPO)
    for finding in findings:
        fail("public-api", finding)

    drifted = False
    for section, _ in vba_api.SECTIONS:
        was = set(recorded.get(section, []))
        now = set(actual.get(section, []))

        for gone in sorted(was - now):
            drifted = True
            fail("public-api", f"[{section}] no longer declared: {gone}")
        for added in sorted(now - was):
            drifted = True
            fail("public-api",
                 f"[{section}] declared without a manifest entry: {added}")

    if drifted:
        fail("public-api",
             "run tools/vba_api.py --write to record an intended API change, "
             "and declare it in CHANGELOG.md; a [supported] change is a "
             "Semantic Versioning event")


SEMVER_ROW_RE = re.compile(
    r"^\|\s*Supported API contract\s*\|\s*(.+?)\s*\|\s*$", re.M)

SEMVER_WORDS = ("patch", "minor", "major")


def check_supported_api_declaration() -> None:
    """Require a Semantic Versioning statement whenever the facade moves.

    Documented policy is not a gate. The manifest carries the facade as it
    stood at the last release in its [baseline] section, so a change to the
    external contract is detectable in the branch that makes it rather than
    only in the release diff — and detectable from the checkout alone, which
    matters because CI clones one commit and has no history to compare against.

    When the two differ, CHANGELOG.md has to say so in words a reader sees,
    not only in a manifest a reader has to diff.
    """
    vba_api = import_vba_api()
    if vba_api is None:
        return

    manifest_path = os.path.join(REPO, "tools/public_api_manifest.txt")
    if not os.path.exists(manifest_path):
        return                       # already reported by check_public_api

    try:
        recorded, baseline_version = vba_api.parse_manifest(manifest_path)
    except vba_api.ApiError:
        return                       # already reported by check_public_api

    baseline = set(recorded.get(vba_api.BASELINE_SECTION, []))
    supported = set(recorded.get("supported", []))
    changed = baseline != supported

    changelog = read("CHANGELOG.md").decode("utf-8")
    match = SEMVER_ROW_RE.search(changelog)

    if match is None:
        fail("supported-api-declaration",
             "CHANGELOG.md has no '| Supported API contract | ... |' row; the "
             "facade cannot be silently unchanged, it has to be stated")
        return

    stated = match.group(1).strip()
    unchanged_claim = stated.lower().startswith("unchanged")

    if changed and unchanged_claim:
        fail("supported-api-declaration",
             f"[supported] differs from [baseline {baseline_version}], but "
             f"CHANGELOG.md still claims {stated!r}")
    elif changed and not any(w in stated.lower() for w in SEMVER_WORDS):
        fail("supported-api-declaration",
             f"[supported] differs from [baseline {baseline_version}]; "
             f"CHANGELOG.md must name the release type (patch, minor or "
             f"major), and says {stated!r}")
    elif not changed and not unchanged_claim:
        fail("supported-api-declaration",
             f"[supported] is identical to [baseline {baseline_version}], but "
             f"CHANGELOG.md claims {stated!r}")

    if changed:
        for gone in sorted(baseline - supported):
            note("supported-api-declaration",
                 f"no longer in the facade since {baseline_version}: {gone}")
        for added in sorted(supported - baseline):
            note("supported-api-declaration",
                 f"new in the facade since {baseline_version}: {added}")


def check_release_state_selftest() -> None:
    """Run the VERSION rules against their own fixtures.

    Two other checks derive from this file, so a rule that quietly accepted
    anything would weaken them without ever reporting a thing.
    """
    for label, raw, expected in VERSION_SELFTEST_CASES:
        found = version_findings(raw)
        if len(found) != expected:
            fail("release-state-selftest",
                 f"{label}: expected {expected} finding(s), got {len(found)}"
                 + (f" | {found}" if found else ""))


def check_wiki_badge_selftest() -> None:
    """Run the wiki badge rules against their own fixtures.

    The wiki itself is not reachable from here — check_repo.py reads the tree
    and never the network, so a machine without wiki access still gets a
    meaningful gate. The wiki-badges workflow does the clone. What can be
    checked offline is that the rules themselves still hold, which stops the
    badge policy rotting between the independent workflow's runs.
    """
    sys.path.insert(0, os.path.join(REPO, "tools"))
    try:
        import wiki_badges
    except Exception as exc:                                  # pragma: no cover
        fail("wiki-badge-selftest",
             f"tools/wiki_badges.py could not be imported: {exc}")
        return

    for finding in wiki_badges.selftest():
        fail("wiki-badge-selftest", finding.replace("\n", " | "))


def check_public_api_selftest() -> None:
    """Run the contract model's own fixtures.

    The manifest is only as good as the parser that produces it. A model that
    silently normalised a ByRef flip away would keep the gate green through
    exactly the change it exists to catch, and no VBA test can see that. The
    fixtures cover every breaking class the gate claims, plus the formatting
    changes it must ignore.
    """
    vba_api = import_vba_api()
    if vba_api is None:
        return

    if not hasattr(vba_api, "selftest"):
        fail("public-api-selftest", "tools/vba_api.py exposes no selftest()")
        return

    for finding in vba_api.selftest():
        fail("public-api-selftest", finding.replace("\n", " | "))


# --------------------------------------------------------------------------
# YAML permits whitespace before the colon, so `uses : x` is the same mapping
# as `uses: x`. Matching only the tight form let an unpinned reference through.
USES_RE = re.compile(
    r"^\s*(?:-\s*)?uses\s*:\s*(?P<ref>[^\s#]+)\s*(?P<rest>.*)$")

PINNED_RE = re.compile(r"^[\w.\-]+/[\w.\-]+(?:/[^@]+)?@(?P<sha>[0-9a-f]+)$")

# An exact release, not a series. `# v4` names the same mutable pointer the
# pin exists to escape, so it documents nothing a reviewer can check.
AUDIT_COMMENT_RE = re.compile(
    r"^#\s*v[0-9]+\.[0-9]+\.[0-9]+(?:[-+][0-9A-Za-z.\-+]*)?\s*$")

ON_KEY_RE = re.compile(r"""^(?:on|'on'|"on"|True)\s*:\s*$""")

GATE_WORKFLOW = ".github/workflows/static-checks.yml"


def yaml_block_lines(text: str) -> list[tuple[int, str]]:
    """Return [(indent, content)] for lines that carry YAML structure."""
    out: list[tuple[int, str]] = []
    for raw in text.split("\n"):
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        out.append((len(raw) - len(raw.lstrip(" ")), raw.strip()))
    return out


def child_block(lines: list[tuple[int, str]], start: int) -> tuple[list[tuple[int, str]], int]:
    """Return the lines nested under lines[start], and the index after them."""
    parent_indent = lines[start][0]
    end = start + 1
    while end < len(lines) and lines[end][0] > parent_indent:
        end += 1
    return lines[start + 1:end], end


def find_key(lines: list[tuple[int, str]], name: str) -> int | None:
    """Return the index of a mapping key at the block's own indentation."""
    if not lines:
        return None
    indent = lines[0][0]
    pattern = re.compile(rf"""^(?:{name}|'{name}'|"{name}")\s*:""")
    for i, (ind, content) in enumerate(lines):
        if ind == indent and pattern.match(content):
            return i
    return None


def has_push_tag_trigger(text: str) -> bool:
    """Return True only when on.push.tags contains the v* pattern.

    Searching the file for a tags: block matched one under pull_request, or
    under any other key, and reported a workflow as tag-triggered when no tag
    push would ever start it. The association has to be structural.
    """
    lines = yaml_block_lines(text)

    on_at = next((i for i, (ind, content) in enumerate(lines)
                  if ind == 0 and ON_KEY_RE.match(content)), None)
    if on_at is None:
        return False

    on_block, _ = child_block(lines, on_at)

    push_at = find_key(on_block, "push")
    if push_at is None:
        return False

    push_block, _ = child_block(on_block, push_at)

    tags_at = find_key(push_block, "tags")
    if tags_at is None:
        return False

    tags_block, _ = child_block(push_block, tags_at)

    for _, content in tags_block:
        if not content.startswith("-"):
            continue
        if content[1:].strip().strip("\"'") == "v*":
            return True
    return False


def workflow_findings(rel: str, text: str, require_tag_trigger: bool = False) -> list[str]:
    """Return every pin-policy violation in one workflow file.

    A version tag is a mutable pointer: actions/checkout@v4 has moved between
    commits, so a workflow referencing it runs whatever the upstream account
    publishes next — including whatever an attacker publishes after
    compromising it. Only a full commit SHA is immutable.

    The trailing version comment is required because a bare 40-hex string is
    unauditable in review: nobody can tell v4.4.0 from an arbitrary commit by
    reading it. The comment states the claim so a reviewer can resolve and
    verify it independently.
    """
    findings: list[str] = []

    for n, line in enumerate(text.split("\n"), 1):
        m = USES_RE.match(line)
        if not m:
            continue

        ref = m.group("ref").strip("\"'")
        rest = m.group("rest").strip()

        # A repository-local action is versioned by this repository's own
        # history, so there is nothing to pin it to.
        if ref.startswith("./"):
            continue

        pinned = PINNED_RE.match(ref)
        if not pinned:
            findings.append(
                f"{rel}:{n}: {ref} is not pinned to a commit SHA; a version "
                f"tag or branch can be moved by the upstream account"
            )
            continue

        sha = pinned.group("sha")
        if len(sha) != 40:
            findings.append(
                f"{rel}:{n}: {ref} is pinned to a {len(sha)}-character SHA; "
                f"a truncated SHA is ambiguous and must be the full 40"
            )
            continue

        if not AUDIT_COMMENT_RE.match(rest):
            findings.append(
                f"{rel}:{n}: {ref} carries no exact audited-version comment; "
                f"write the release it resolves to, as '# v1.2.3', so the pin "
                f"can be reviewed without resolving it by hand. A series such "
                f"as '# v4' names the mutable pointer the pin exists to escape"
            )

    if require_tag_trigger and not has_push_tag_trigger(text):
        findings.append(
            f"{rel}: no push trigger on tags 'v*'; a tag is the artefact "
            f"people install from, and nothing would run against it"
        )

    return findings


def check_workflow_policy() -> None:
    """Enforce immutable Action pins across every tracked workflow."""
    tracked = git_tracked_files()
    if tracked is None:
        return

    workflows = [
        rel for rel in tracked
        if rel.startswith(".github/workflows/")
        and rel.endswith((".yml", ".yaml"))
    ]

    if not workflows:
        fail("workflow-policy", "no workflow is tracked under .github/workflows")
        return

    if GATE_WORKFLOW not in workflows:
        fail("workflow-policy", f"{GATE_WORKFLOW} is missing")

    for rel in sorted(workflows):
        text = read(rel).decode("utf-8")
        for finding in workflow_findings(
                rel, text, require_tag_trigger=(rel == GATE_WORKFLOW)):
            fail("workflow-policy", finding)


WORKFLOW_HEAD = """\
name: Example

on:
  push:
    branches:
      - main
    tags:
      - 'v*'

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
"""

WORKFLOW_SELFTEST_CASES = [
    ("a pinned action with a tag trigger passes",
     WORKFLOW_HEAD
     + "      - uses: actions/checkout@" + "1" * 40 + " # v4.4.0\n",
     0),
    ("a repository-local action needs no pin",
     WORKFLOW_HEAD + "      - uses: ./.github/actions/house-style\n",
     0),
    ("a moving version tag is rejected",
     WORKFLOW_HEAD + "      - uses: actions/checkout@v4\n",
     1),
    ("a branch reference is rejected",
     WORKFLOW_HEAD + "      - uses: actions/checkout@main\n",
     1),
    ("a truncated SHA is rejected",
     WORKFLOW_HEAD + "      - uses: actions/checkout@" + "1" * 8 + " # v4.4.0\n",
     1),
    ("a full SHA without an audited-version comment is rejected",
     WORKFLOW_HEAD + "      - uses: actions/checkout@" + "1" * 40 + "\n",
     1),
    ("a full SHA with a non-version comment is rejected",
     WORKFLOW_HEAD
     + "      - uses: actions/checkout@" + "1" * 40 + " # latest\n",
     1),
    ("a missing tag trigger is rejected",
     WORKFLOW_HEAD.replace("    tags:\n      - 'v*'\n", "")
     + "      - uses: actions/checkout@" + "1" * 40 + " # v4.4.0\n",
     1),
    ("a tag trigger under pull_request does not satisfy on.push",
     WORKFLOW_HEAD.replace("    tags:\n      - 'v*'\n", "")
     .replace("on:\n", "on:\n  pull_request:\n    tags:\n      - 'v*'\n")
     + "      - uses: actions/checkout@" + "1" * 40 + " # v4.4.0\n",
     1),
    ("a tags key outside the on block does not satisfy on.push",
     WORKFLOW_HEAD.replace("    tags:\n      - 'v*'\n", "")
     + "      - uses: actions/checkout@" + "1" * 40 + " # v4.4.0\n"
     + "        with:\n          tags:\n            - 'v*'\n",
     1),
    ("whitespace before the colon does not hide a moving tag",
     WORKFLOW_HEAD + "      - uses : actions/checkout@v4\n",
     1),
    ("a quoted reference is read, not skipped",
     WORKFLOW_HEAD + '      - uses: "actions/checkout@v4"\n',
     1),
    ("a version series is not an audited version",
     WORKFLOW_HEAD
     + "      - uses: actions/checkout@" + "1" * 40 + " # v4\n",
     1),
    ("a major.minor comment is not an audited version",
     WORKFLOW_HEAD
     + "      - uses: actions/checkout@" + "1" * 40 + " # v4.4\n",
     1),
]


def check_workflow_policy_selftest() -> None:
    """Run the pin policy against its own fixtures.

    The policy is one regular expression away from accepting everything, and
    the only symptom would be a green gate. Every tracked workflow inherits
    this rule, so immutable pins have to be enforced rather than remembered.
    """
    for label, text, expected in WORKFLOW_SELFTEST_CASES:
        found = workflow_findings("fixture.yml", text, require_tag_trigger=True)
        if len(found) != expected:
            fail("workflow-policy-selftest",
                 f"{label}: expected {expected} finding(s), got {len(found)}"
                 + (f" | {found}" if found else ""))


# --------------------------------------------------------------------------
# Branch coverage. Which branches Dependabot updates and CodeQL analyses is
# configured by hand, and a branch that is renamed, added or forgotten loses
# its updates or its analysis without anything turning red.

DEPENDABOT_CONFIG = ".github/dependabot.yml"
CODEQL_WORKFLOW = ".github/workflows/codeql.yml"

# The settings a release-branch entry must share with its ecosystem's
# default-branch entry, so both branches receive the same updates.
MIRRORED_KEYS = ("directory", "schedule", "groups", "labels")

# The branch filters CodeQL must include, per event.
CODEQL_BRANCHES = {
    "push": ("main",),
    "pull_request": ("main", "release/**"),
}

# The only filter the check reads. Another one (paths, branches-ignore, types)
# can narrow what runs in ways the check cannot judge, so it is reported.
CODEQL_FILTER_KEYS = ("branches",)

# GitHub reads a branch pattern starting with this as an exclusion.
NEGATIVE_PATTERN_PREFIX = "!"

VERSION_CORE_RE = re.compile(
    r"^(?P<major>[0-9]+)\.(?P<minor>[0-9]+)\.(?P<patch>[0-9]+)(?P<suffix>[-+].*)?$")

RELEASE_BRANCH_RE = re.compile(r"^release/v(?P<core>[0-9]+\.[0-9]+\.[0-9]+)$")

YamlValue: TypeAlias = "str | list[YamlValue] | dict[str, YamlValue] | None"

# One structural line: (line number, indentation, content without comment).
YamlLine = tuple[int, int, str]

YAML_KEY_RE = re.compile(
    r"""^(?P<key>"[^"\\]*"|'[^']*'|[^\s"'#\-?:,\[\]{}&*!|>%@`][^:]*?)\s*:"""
    r"""(?:\s+(?P<value>.*))?$""")

# A plain scalar starting with one of these is a flow collection, an anchor,
# an alias, a tag, a block scalar or a reserved indicator.
YAML_UNSUPPORTED_START = "[]{}&*!|>%@`"


class YamlShapeError(ValueError):
    """The text uses YAML that the narrow reader below does not understand."""


def strip_yaml_comment(line: str) -> str:
    """Remove a trailing comment, leaving a # inside a quoted scalar alone."""
    quote = ""
    i = 0
    while i < len(line):
        ch = line[i]
        if quote:
            if ch == quote:
                if quote == "'" and line[i + 1:i + 2] == "'":
                    i += 2
                    continue
                quote = ""
        elif ch in "'\"" and (i == 0 or line[i - 1] == " "):
            quote = ch
        elif ch == "#" and (i == 0 or line[i - 1] in " \t"):
            return line[:i].rstrip()
        i += 1
    return line


def yaml_lines(text: str) -> list[YamlLine]:
    """Return the lines that carry structure, with comments removed."""
    out: list[YamlLine] = []
    for n, raw in enumerate(text.split("\n"), 1):
        body = strip_yaml_comment(raw.rstrip())
        if not body.strip():
            continue
        indent = len(body) - len(body.lstrip(" "))
        out.append((n, indent, body[indent:]))
    return out


class YamlReader:
    """A deliberately narrow reader for block-style YAML.

    It reads block mappings, block sequences and single-line plain or quoted
    scalars, and keeps every scalar as a string. Anything else — flow
    collections, anchors, aliases, tags, block scalars, multi-document
    streams, tab indentation — raises YamlShapeError rather than being guessed
    at: a reader that silently misread a shape would let the check pass a
    configuration it never really read. The gate uses only the standard
    library, so there is no full YAML parser to fall back on.
    """

    def __init__(self, rel: str, lines: list[YamlLine]) -> None:
        self.rel = rel
        self.lines = list(lines)

    def error(self, i: int, message: str) -> YamlShapeError:
        n = self.lines[min(i, len(self.lines) - 1)][0] if self.lines else 0
        return YamlShapeError(
            f"{self.rel}:{n}: {message}; the check reads block mappings, block "
            f"sequences and single-line scalars only")

    def parse(self) -> YamlValue:
        if not self.lines:
            return None
        value, end = self.block(0, self.lines[0][1])
        if end != len(self.lines):
            raise self.error(end, "inconsistent indentation")
        return value

    @staticmethod
    def is_item(content: str) -> bool:
        return content == "-" or content.startswith("- ")

    def nested(self, i: int, indent: int) -> tuple[YamlValue, int]:
        """Return the block nested under line i, or None when there is none."""
        if i + 1 < len(self.lines) and self.lines[i + 1][1] > indent:
            return self.block(i + 1, self.lines[i + 1][1])
        return None, i + 1

    def block(self, i: int, indent: int) -> tuple[YamlValue, int]:
        if self.is_item(self.lines[i][2]):
            return self.sequence(i, indent)
        return self.mapping(i, indent)

    def mapping(self, i: int, indent: int) -> tuple[YamlValue, int]:
        out: dict[str, YamlValue] = {}
        while i < len(self.lines):
            _, ind, content = self.lines[i]
            if ind < indent:
                break
            if ind > indent:
                raise self.error(i, "unexpected indentation")
            if self.is_item(content):
                raise self.error(i, "a sequence item where a mapping key was expected")
            m = YAML_KEY_RE.match(content)
            if not m:
                raise self.error(i, f"cannot read {content!r} as a mapping key")
            key = self.scalar(i, m.group("key"))
            if key in out:
                raise self.error(i, f"duplicate key {key!r}")
            if m.group("value"):
                out[key] = self.scalar(i, m.group("value"))
                i += 1
            else:
                out[key], i = self.nested(i, ind)
        return out, i

    def sequence(self, i: int, indent: int) -> tuple[YamlValue, int]:
        out: list[YamlValue] = []
        while i < len(self.lines):
            n, ind, content = self.lines[i]
            if ind < indent:
                break
            if ind > indent:
                raise self.error(i, "unexpected indentation")
            if not self.is_item(content):
                raise self.error(i, "a mapping key where a sequence item was expected")
            rest = content[1:].lstrip(" ")
            if not rest:
                item, i = self.nested(i, ind)
            elif YAML_KEY_RE.match(rest):
                # "- key: value" opens a mapping whose keys align with "key".
                column = ind + len(content) - len(rest)
                self.lines[i] = (n, column, rest)
                item, i = self.mapping(i, column)
            else:
                item = self.scalar(i, rest)
                i += 1
            out.append(item)
        return out, i

    def scalar(self, i: int, text: str) -> str:
        if text.startswith('"'):
            inner = text[1:-1]
            if len(text) < 2 or not text.endswith('"') or '"' in inner or "\\" in inner:
                raise self.error(i, f"cannot read the double-quoted scalar {text}")
            return inner
        if text.startswith("'"):
            inner = text[1:-1]
            if len(text) < 2 or not text.endswith("'") or "'" in inner.replace("''", ""):
                raise self.error(i, f"cannot read the single-quoted scalar {text}")
            return inner.replace("''", "'")
        if text[0] in YAML_UNSUPPORTED_START or text[0] == "\t":
            raise self.error(i, f"{text!r} is a flow collection, anchor, alias, tag "
                                f"or block scalar")
        if ": " in text or text in ("---", "..."):
            raise self.error(i, f"cannot read {text!r} as a plain scalar")
        return text


def parse_yaml(rel: str, text: str) -> YamlValue:
    """Parse a whole block-style YAML document, or raise YamlShapeError."""
    lines = yaml_lines(text)
    if lines and lines[0][1] != 0:
        raise YamlReader(rel, lines).error(0, "the document must start at column 0")
    return YamlReader(rel, lines).parse()


def open_release_branches(version: str) -> list[str] | None:
    """Return the release branches that may be open while VERSION is current.

    VERSION names the newest shipped release: the release-state check holds it
    to the newest dated CHANGELOG.md section, and it moves only in the change
    that dates the next one. While a release is being prepared, VERSION still
    names the previous release, so the open branch is the next patch, minor or
    major release after it. A pre-release VERSION may also be followed by its
    own final release.
    """
    m = VERSION_CORE_RE.match(version)
    if not m:
        return None
    major, minor, patch = (int(m.group(k)) for k in ("major", "minor", "patch"))
    branches: list[str] = []
    if (m.group("suffix") or "").startswith("-"):
        branches.append(f"release/v{major}.{minor}.{patch}")
    branches += [
        f"release/v{major}.{minor}.{patch + 1}",
        f"release/v{major}.{minor + 1}.0",
        f"release/v{major + 1}.0.0",
    ]
    return branches


def is_shipped_release(branch: str, version: str) -> bool:
    """Return True when the branch names VERSION or an earlier release."""
    branch_m = RELEASE_BRANCH_RE.match(branch)
    version_m = VERSION_CORE_RE.match(version)
    if not branch_m or not version_m:
        return False
    core = tuple(int(part) for part in branch_m.group("core").split("."))
    current = tuple(int(version_m.group(k)) for k in ("major", "minor", "patch"))
    return core <= current


DependabotEntry = dict[str, YamlValue]


def classify_dependabot_entries(
        rel: str, updates: list[YamlValue], version: str, candidates: list[str],
) -> tuple[dict[str, list[DependabotEntry]],
           dict[str, dict[str, list[DependabotEntry]]], list[str]]:
    """Sort entries into default-branch and release-branch entries per ecosystem.

    Returns the default entries, the release entries keyed by target branch,
    and a finding for every entry that names no branch the check accepts.
    """
    findings: list[str] = []
    defaults: dict[str, list[DependabotEntry]] = {}
    releases: dict[str, dict[str, list[DependabotEntry]]] = {}

    for k, entry in enumerate(updates):
        assert isinstance(entry, dict)
        ecosystem = entry.get("package-ecosystem")
        target = entry.get("target-branch")
        if not isinstance(ecosystem, str):
            findings.append(f"{rel}: updates[{k}] has no package-ecosystem")
            continue
        where = f"{rel}: updates[{k}] ({ecosystem})"
        if "directories" in entry:
            findings.append(f"{where} uses 'directories'; the check compares a "
                            f"single 'directory'")
        elif target is None:
            defaults.setdefault(ecosystem, []).append(entry)
        elif not isinstance(target, str):
            findings.append(f"{where} has a target-branch that is not a branch name")
        elif target in candidates:
            releases.setdefault(ecosystem, {}).setdefault(target, []).append(entry)
        elif is_shipped_release(target, version):
            findings.append(
                f"{where} targets {target}, but VERSION is {version}: that release "
                f"has shipped, so the entry is stale. Point it at the release "
                f"branch that is open now, or remove it when none is")
        else:
            findings.append(
                f"{where} targets {target}, which is not a release that can "
                f"follow VERSION {version} ({', '.join(candidates)}); the "
                f"default-branch entry names no target-branch")
    return defaults, releases, findings


def release_entry_findings(rel: str, ecosystem: str, target: str,
                           entries: list[DependabotEntry],
                           defaults: list[DependabotEntry]) -> list[str]:
    """Return the problems with one ecosystem's entries for the open release."""
    if not entries:
        return [f"{rel}: {ecosystem} has no entry for {target}, which other "
                f"ecosystems target; every ecosystem must cover the open "
                f"release branch"]
    findings: list[str] = []
    if len(entries) > 1:
        findings.append(f"{rel}: {ecosystem} has {len(entries)} entries for "
                        f"{target}; it needs exactly one")
    if len(defaults) != 1:
        return findings              # reported as a default-entry count
    for key in MIRRORED_KEYS:
        if any(entry.get(key) != defaults[0].get(key) for entry in entries):
            findings.append(
                f"{rel}: {ecosystem} entry for {target}: {key} differs from the "
                f"default-branch entry; both branches must receive the same "
                f"updates")
    return findings


def dependabot_findings(rel: str, text: str, version: str) -> list[str]:
    """Return every branch-coverage problem in a Dependabot configuration.

    Each ecosystem needs exactly one entry without target-branch, which
    Dependabot applies to the default branch. While a release is open, each
    ecosystem also needs one entry for that release branch, mirroring the
    default entry's settings, and no entry may name any other branch. After a
    release ships, the default entries alone are the correct configuration.
    """
    candidates = open_release_branches(version)
    if candidates is None:
        return [f"VERSION does not hold a release number ({version!r}), so the "
                f"open release branch cannot be derived"]

    try:
        config = parse_yaml(rel, text)
    except YamlShapeError as exc:
        return [str(exc)]

    updates = config.get("updates") if isinstance(config, dict) else None
    if not isinstance(updates, list) or not all(isinstance(e, dict) for e in updates):
        return [f"{rel}: 'updates' must be a block sequence of mappings"]

    defaults, releases, findings = classify_dependabot_entries(
        rel, updates, version, candidates)

    ecosystems = sorted(set(defaults) | set(releases))
    for ecosystem in ecosystems:
        count = len(defaults.get(ecosystem, []))
        if count != 1:
            findings.append(
                f"{rel}: {ecosystem} has {count} entries without target-branch; it "
                f"needs exactly one, for the default branch")

    targets = sorted({t for by_target in releases.values() for t in by_target})
    if len(targets) > 1:
        findings.append(
            f"{rel}: entries target {', '.join(targets)}; only one release "
            f"branch is open at a time")
        return findings

    for ecosystem in ecosystems:
        for target in targets:
            findings += release_entry_findings(
                rel, ecosystem, target, releases.get(ecosystem, {}).get(target, []),
                defaults.get(ecosystem, []))
    return findings


def codeql_findings(rel: str, text: str) -> list[str]:
    """Return every branch-coverage problem in the CodeQL workflow triggers."""
    lines = yaml_lines(text)
    on_at = next((i for i, (_, ind, content) in enumerate(lines)
                  if ind == 0 and ON_KEY_RE.match(content)), None)
    if on_at is None:
        return [f"{rel}: no 'on:' block mapping; the triggers cannot be read"]

    block: list[YamlLine] = []
    for line in lines[on_at + 1:]:
        if line[1] == 0:
            break
        block.append(line)

    try:
        triggers = YamlReader(rel, block).parse()
    except YamlShapeError as exc:
        return [str(exc)]

    if not isinstance(triggers, dict):
        return [f"{rel}: 'on' must be a block mapping of events"]

    findings: list[str] = []
    for event, required in CODEQL_BRANCHES.items():
        if event not in triggers:
            findings.append(f"{rel}: on.{event} is missing; CodeQL must analyse "
                            f"{event} events for {', '.join(required)}")
        else:
            findings += codeql_trigger_findings(rel, event, triggers[event], required)
    return findings


def codeql_trigger_findings(rel: str, event: str, trigger: YamlValue,
                            required: tuple[str, ...]) -> list[str]:
    """Return the problems with one CodeQL event's branch filter.

    GitHub evaluates branch patterns in order, and a later negative pattern
    ('!main') excludes a branch an earlier one included. Membership alone
    would then report coverage that does not exist, so a negative pattern is
    rejected rather than evaluated.
    """
    if trigger is None:
        return []                    # no filter: every branch is analysed
    if not isinstance(trigger, dict):
        return [f"{rel}: on.{event} must be a mapping with a branches list"]

    findings = [
        f"{rel}: on.{event}.{key} is a filter this check does not read; teach "
        f"tools/check_repo.py about it before relying on it"
        for key in trigger if key not in CODEQL_FILTER_KEYS
    ]
    branches = trigger.get("branches")
    if not isinstance(branches, list) or not all(isinstance(b, str) for b in branches):
        findings.append(f"{rel}: on.{event}.branches must be a list of branch names")
        return findings
    for branch in branches:
        assert isinstance(branch, str)
        if branch.startswith(NEGATIVE_PATTERN_PREFIX):
            findings.append(
                f"{rel}: on.{event}.branches has the negative pattern {branch!r}; "
                f"it can exclude a branch listed before it, and the check does "
                f"not evaluate pattern order")
    for branch in required:
        if branch not in branches:
            findings.append(f"{rel}: on.{event}.branches does not include {branch!r}")
    return findings


def check_branch_coverage() -> None:
    """Hold Dependabot and CodeQL to the default and the open release branch."""
    if not os.path.isfile(os.path.join(REPO, "VERSION")):
        return                       # already reported by required files

    version = read("VERSION").decode("utf-8").strip()

    for rel in (DEPENDABOT_CONFIG, CODEQL_WORKFLOW):
        if not os.path.isfile(os.path.join(REPO, rel)):
            fail("branch-coverage", f"{rel} is missing")
            continue
        text = read(rel).decode("utf-8")
        found = (dependabot_findings(rel, text, version) if rel == DEPENDABOT_CONFIG
                 else codeql_findings(rel, text))
        for finding in found:
            fail("branch-coverage", finding)


def dependabot_entry(template: str, target: str | None = None) -> str:
    """Return one fixture entry, for the default branch or for a target."""
    line = f"    target-branch: {target}\n" if target else ""
    return template.replace("    schedule:\n", line + "    schedule:\n", 1)


DEPENDABOT_ACTIONS = """\
  - package-ecosystem: github-actions
    directory: /
    schedule:
      interval: weekly
      day: monday
      time: "05:23"
      timezone: Europe/Rome
    groups:
      github-actions:
        patterns:
          - "*"
    labels:
      - ci
      - P3
    commit-message:
      prefix: ci
"""

DEPENDABOT_PIP = """\
  - package-ecosystem: pip
    directory: /tools
    schedule:
      interval: monthly
    groups:
      python-quality-tools:
        patterns:
          - "*"
    open-pull-requests-limit: 5
    labels:
      - ci
      - P3
"""


def dependabot_doc(*entries: str) -> str:
    return "# Fixture.\nversion: 2\nupdates:\n" + "\n".join(entries)


def dependabot_pair(release: str | None, actions: str = DEPENDABOT_ACTIONS,
                    pip: str = DEPENDABOT_PIP) -> str:
    """Return the shape of the tracked configuration with one release open."""
    entries = [dependabot_entry(DEPENDABOT_ACTIONS)]
    if release:
        entries.append(dependabot_entry(actions, release))
    entries.append(dependabot_entry(DEPENDABOT_PIP))
    if release:
        entries.append(dependabot_entry(pip, release))
    return dependabot_doc(*entries)


DEPENDABOT_SELFTEST_CASES = [
    ("the tracked shape passes", "1.1.2", dependabot_pair("release/v1.1.3"), 0),
    ("no release entry passes while no release is open", "1.1.2",
     dependabot_pair(None), 0),
    ("a minor release branch passes", "1.1.2", dependabot_pair("release/v1.2.0"), 0),
    ("a major release branch passes", "1.1.2", dependabot_pair("release/v2.0.0"), 0),
    ("a pre-release VERSION may be followed by its final release", "1.2.0-rc.1",
     dependabot_pair("release/v1.2.0"), 0),
    ("a missing release entry is rejected", "1.1.2",
     dependabot_doc(dependabot_entry(DEPENDABOT_ACTIONS),
                    dependabot_entry(DEPENDABOT_ACTIONS, "release/v1.1.3"),
                    dependabot_entry(DEPENDABOT_PIP)),
     1),
    ("a stale release branch is rejected", "1.1.3",
     dependabot_pair("release/v1.1.3"), 2),
    ("an older release branch is rejected", "1.1.3",
     dependabot_pair("release/v1.1.2"), 2),
    ("a release two patches ahead is rejected", "1.1.2",
     dependabot_pair("release/v1.1.4"), 2),
    ("an extra branch is rejected", "1.1.2",
     dependabot_pair("release/v1.1.3")
     + "\n" + dependabot_entry(DEPENDABOT_ACTIONS, "develop"),
     1),
    ("two open release branches are rejected", "1.1.2",
     dependabot_doc(dependabot_entry(DEPENDABOT_ACTIONS),
                    dependabot_entry(DEPENDABOT_ACTIONS, "release/v1.1.3"),
                    dependabot_entry(DEPENDABOT_PIP),
                    dependabot_entry(DEPENDABOT_PIP, "release/v1.2.0")),
     1),
    ("a second default entry is rejected", "1.1.2",
     dependabot_pair("release/v1.1.3") + "\n" + dependabot_entry(DEPENDABOT_PIP),
     1),
    ("a second release entry is rejected", "1.1.2",
     dependabot_pair("release/v1.1.3")
     + "\n" + dependabot_entry(DEPENDABOT_PIP, "release/v1.1.3"),
     1),
    ("a mismatched directory is rejected", "1.1.2",
     dependabot_pair("release/v1.1.3",
                     pip=DEPENDABOT_PIP.replace("directory: /tools", "directory: /")),
     1),
    ("a mismatched schedule is rejected", "1.1.2",
     dependabot_pair("release/v1.1.3",
                     actions=DEPENDABOT_ACTIONS.replace("day: monday", "day: friday")),
     1),
    ("mismatched groups are rejected", "1.1.2",
     dependabot_pair("release/v1.1.3",
                     pip=DEPENDABOT_PIP.replace("python-quality-tools:", "tools:")),
     1),
    ("mismatched labels are rejected", "1.1.2",
     dependabot_pair("release/v1.1.3",
                     actions=DEPENDABOT_ACTIONS.replace("      - P3\n", "      - P2\n")),
     1),
    ("a malformed VERSION is reported", "next", dependabot_pair("release/v1.1.3"), 1),
    ("a flow sequence is reported, not guessed at", "1.1.2",
     dependabot_pair("release/v1.1.3").replace("    labels:\n      - ci\n      - P3\n",
                                               "    labels: [ci, P3]\n", 1),
     1),
    ("an alias is reported, not guessed at", "1.1.2",
     dependabot_pair("release/v1.1.3").replace('time: "05:23"', "time: *at", 1),
     1),
    ("tab indentation is reported, not guessed at", "1.1.2",
     dependabot_pair("release/v1.1.3").replace("      day: monday", "\tday: monday", 1),
     1),
    ("a directories list is reported", "1.1.2",
     dependabot_pair(None).replace("    directory: /tools\n",
                                   "    directories:\n      - /tools\n"),
     1),
]

CODEQL_TRIGGERS = """\
name: CodeQL

on:
  pull_request:
    branches:
      - main
      - 'release/**'
  push:
    branches:
      - main
      - 'release/**'
  schedule:
    - cron: '41 4 * * 2'
  workflow_dispatch:

jobs:
  analyse:
    runs-on: ubuntu-latest
    steps:
      - run: |
          echo "a block scalar outside 'on' is not read"
"""

CODEQL_SELFTEST_CASES = [
    ("the tracked triggers pass", CODEQL_TRIGGERS, 0),
    ("an unfiltered push trigger passes",
     CODEQL_TRIGGERS.replace("  push:\n    branches:\n      - main\n      - 'release/**'\n",
                             "  push:\n"),
     0),
    ("a pull-request trigger without release/** is rejected",
     CODEQL_TRIGGERS.replace("      - main\n      - 'release/**'\n  push:",
                             "      - main\n  push:"),
     1),
    ("a push trigger without main is rejected",
     CODEQL_TRIGGERS.replace("  push:\n    branches:\n      - main\n",
                             "  push:\n    branches:\n"),
     1),
    ("a missing pull-request trigger is rejected",
     CODEQL_TRIGGERS.replace("  pull_request:\n    branches:\n      - main\n"
                             "      - 'release/**'\n", ""),
     1),
    ("an unread filter is reported",
     CODEQL_TRIGGERS.replace("  push:\n", "  push:\n    paths:\n      - tools/**\n"),
     1),
    ("a flow sequence of branches is reported, not guessed at",
     CODEQL_TRIGGERS.replace("    branches:\n      - main\n      - 'release/**'\n  push:",
                             "    branches: [main, 'release/**']\n  push:"),
     1),
    ("a negative pattern after main is rejected",
     CODEQL_TRIGGERS.replace("      - 'release/**'\n  push:",
                             "      - 'release/**'\n      - '!main'\n  push:"),
     1),
    ("a negative pattern after release/** is rejected",
     CODEQL_TRIGGERS.replace("      - 'release/**'\n  schedule:",
                             "      - 'release/**'\n      - '!release/**'\n  schedule:"),
     1),
    ("a missing on block is reported",
     CODEQL_TRIGGERS.replace("on:\n", "triggers:\n"),
     1),
]


def check_branch_coverage_selftest() -> None:
    """Run the branch-coverage rules against their own fixtures.

    A rule that read nothing would report nothing, and a green gate would then
    say every branch is covered when none is checked.
    """
    for label, version, text, expected in DEPENDABOT_SELFTEST_CASES:
        found = dependabot_findings("fixture.yml", text, version)
        if len(found) != expected:
            fail("branch-coverage-selftest",
                 f"{label}: expected {expected} finding(s), got {len(found)}"
                 + (f" | {found}" if found else ""))

    for label, text, expected in CODEQL_SELFTEST_CASES:
        found = codeql_findings("fixture.yml", text)
        if len(found) != expected:
            fail("branch-coverage-selftest",
                 f"{label}: expected {expected} finding(s), got {len(found)}"
                 + (f" | {found}" if found else ""))

# --------------------------------------------------------------------------
# Characters a Markdown renderer treats as escapable. A backslash before any of
# them is markup, not text.
ESCAPABLE = r"[\\`*_{}\[\]()#+\-.!|<>&\"'~=:;,/$%^?@]"

ESCAPE_RE = re.compile(r"\\(" + ESCAPABLE + ")")

# A fence may be indented by at most three spaces. Four or more make the line
# indented code, whose backticks are content and open no block.
FENCE_RE = re.compile(r"^ {0,3}(`{3,}|~{3,})")

INLINE_CODE_RE = re.compile(r"`+[^`]*`+")


def strip_inline_code(line: str) -> str:
    """Blank out inline code spans, preserving length so columns stay true."""
    return INLINE_CODE_RE.sub(lambda m: " " * len(m.group(0)), line)


def fence_opener(line: str) -> tuple[str, int] | None:
    """Return (character, length) when the line opens a fenced code block.

    A backtick fence's info string cannot contain a backtick, so a line such
    as ```Use `code` here``` is inline code, not an opener.
    """
    m = FENCE_RE.match(line)
    if not m:
        return None
    token = m.group(1)
    if token[0] == "`" and "`" in line[m.end():]:
        return None
    return token[0], len(token)


def closes_fence(line: str, fence: tuple[str, int]) -> bool:
    """Return whether the line closes the block that the given fence opened.

    A closing fence must use the same character, be at least as long as the
    one that opened it, and carry nothing but spaces after it. Collapsing every
    opener to three let a ``` line close a ```` block; accepting an info string
    let a ```python line inside a ```markdown example close it early. Either
    way the rest of the block was read as prose and its real closing fence as a
    new opener, hiding the headings after it.
    """
    m = FENCE_RE.match(line)
    return bool(m and m.group(1)[0] == fence[0] and len(m.group(1)) >= fence[1]
                and not line[m.end():].strip())


def escape_findings(rel: str, text: str) -> list[str]:
    r"""Report Markdown punctuation escaped outside code.

    A WYSIWYG editor escapes every character it thinks is markup, and both
    independent review archives arrived with several hundred of them: headings
    read `## 1\. Executive assessment`, module names read `M\_EXCEL\_UI`, and a
    horizontal rule read `\---`. The documents rendered as literal backslashes
    and nobody noticed for two releases.

    The check reads only rendered lines, and deliberately does not look inside
    fenced blocks, HTML comment blocks or inline code. A backslash there is data — a Windows path, a regular
    expression — and flagging it would make the gate wrong about correct
    documents. This is a deliberate boundary: the rule detects characteristic
    escaping in prose, not every possible backslash a WYSIWYG editor can emit.
    """
    findings: list[str] = []

    for n, line in rendered_markdown_lines(text):
        found = ESCAPE_RE.findall(strip_inline_code(line))
        if found:
            findings.append(
                f"{rel}:{n}: escaped Markdown outside code: "
                + ", ".join(f"\\{c}" for c in found)
                + "; this is what a WYSIWYG editor leaves behind, and it "
                  "renders as a literal backslash"
            )

    return findings


def check_markdown_escapes() -> None:
    """Fail on escaped Markdown punctuation in any tracked document."""
    tracked = git_tracked_files()
    if tracked is None:
        return

    for rel in sorted(r for r in tracked if r.endswith(".md")):
        for finding in escape_findings(rel, read(rel).decode("utf-8")):
            fail("markdown-escapes", finding)


ESCAPE_SELFTEST_CASES = [
    # Every escape class observed in the two withdrawn archives.
    ("an escaped underscore is reported (\\_)", "The M\\_EXCEL\\_UI module.\n", 1),
    ("an escaped asterisk is reported (\\*)", "\\*\\*Repository:\\*\\*\n", 1),
    ("an escaped hyphen is reported (\\-)", "\\---\n", 1),
    ("an escaped period is reported (\\.)", "## 1\\. Executive assessment\n", 1),
    ("an escaped bracket is reported (\\[)", "\\[a link](x)\n", 1),
    ("an escaped ampersand is reported (\\&)", "Discussions/Q\\&A channel\n", 1),

    # Fence boundaries: the character and its length both matter.
    ("a four-backtick fence is not closed by three backticks",
     "````text\nA\\_B\n```\nM\\_EXCEL\\_UI still inside\n````\n", 0),
    ("a four-backtick fence is closed by four backticks",
     "````text\nA\\_B\n````\nM\\_EXCEL\\_UI in prose\n", 1),
    ("a tilde fence is not closed by backticks",
     "~~~text\nA\\_B\n```\nM\\_EXCEL\\_UI still inside\n~~~\n", 0),
    ("a longer closing fence still closes the block",
     "```text\nA\\_B\n`````\nM\\_EXCEL\\_UI in prose\n", 1),

    # Content that legitimately contains backslashes or looks like it does.
    ("a URL is not an escape", "See https://example.com/a_b/c-d.html for more.\n", 0),
    ("a bare Windows path in a fence is data",
     "```text\n%TEMP%\\commit-msg.txt\n```\n", 0),
    ("a VBA listing in a fence is data",
     "```vb\nUI_Hide = 0\nUI\\_Show = 1\n```\n", 0),
    ("a regular expression in a fence is data",
     "```python\nre.compile(r\"^\\d+\\.\\d+$\")\n```\n", 0),
    ("an underscore without a backslash is not an escape",
     "The M_EXCEL_UI module and its UI_Hide member.\n", 0),
    ("a table row is not an escape", "| Area | State |\n|---|---|\n", 0),
    ("clean prose passes", "# Heading\n\nSome text with M_EXCEL_UI in it.\n", 0),
    ("an escaped heading is reported", "## 1\\. Executive assessment\n", 1),
    ("an escaped underscore is reported", "The M\\_EXCEL\\_UI module.\n", 1),
    ("an escaped rule is reported", "\\---\n", 1),
    ("a backslash inside a fenced block is data",
     "```text\nC:\\Users\\dpenza\\file.txt\n```\n", 0),
    ("an escapable character inside a fenced block is data",
     "```python\nre.compile(r\"\\.txt$\")\n```\n", 0),
    ("a tilde fence is honoured too",
     "~~~text\nA\\_B\n~~~\n", 0),
    ("a closed fence does not swallow the prose that follows",
     "```text\ncode\n```\n\nThen M\\_EXCEL\\_UI in prose.\n", 1),
    ("a backslash inside an inline code span is data",
     "Commit messages go in `%TEMP%\\commit-msg.txt` for encoding reasons.\n", 0),
    ("an escapable character inside an inline code span is data",
     "The pattern `\\.md$` matches documents.\n", 0),
    ("an escape beside an inline code span is still reported",
     "The `code` and the M\\_EXCEL\\_UI module.\n", 1),
    ("several escapes on one line are one finding",
     "\\*\\*Repository:\\*\\* M\\_EXCEL\\_UI\n", 1),
]


def check_markdown_escape_selftest() -> None:
    """Run the escape rules against their own fixtures.

    Fence handling is the part that can silently invert: a rule that treated
    the whole file as fenced would report nothing and look healthy.
    """
    for label, text, expected in ESCAPE_SELFTEST_CASES:
        found = escape_findings("fixture.md", text)
        if len(found) != expected:
            fail("markdown-escape-selftest",
                 f"{label}: expected {expected} finding(s), got {len(found)}"
                 + (f" | {found}" if found else ""))


# --------------------------------------------------------------------------
# Documented commands. A renamed tool or option leaves instructions that fail
# on the reader's machine, often a Windows host where finding out costs an
# Excel session. Every invocation of a repository tool that the documentation
# quotes as code is held to the tool itself. The changelog is history: it may
# name a tool or an option as it was when the entry was written.
DOCUMENTED_COMMANDS_EXCLUDED = ("CHANGELOG.md",)

TOOL_PATH_RE = re.compile(
    r"(?<![\w./\\-])(?:\.[\\/])?tools[\\/][\w.-]+\.(?:py|ps1|cmd)(?![\w.\\/-])")
COMMAND_TOKEN_RE = re.compile(r"\"[^\"]*\"|'[^']*'|\S+")
COMMAND_STOP_TOKENS = frozenset({"|", "||", "&", "&&", ";", ">", ">>", "2>", "2>&1", "<"})
PYTHON_OPTION_RE = re.compile(r"--[A-Za-z][\w-]*")
POWERSHELL_PARAMETER_RE = re.compile(r"-[A-Za-z]\w*")
POWERSHELL_SCRIPT_PARAM_RE = re.compile(
    r"^\s*(?:\[CmdletBinding\((?P<binding>[^)]*)\)\]\s*)?param\s*\(", re.I | re.M)
POWERSHELL_VARIABLE_RE = re.compile(r"\$\{?([A-Za-z_]\w*)")
# The parameters PowerShell adds to every advanced script.
POWERSHELL_COMMON_PARAMETERS = frozenset(name.casefold() for name in (
    "Verbose", "Debug", "ErrorAction", "WarningAction", "InformationAction",
    "ProgressAction", "ErrorVariable", "WarningVariable", "InformationVariable",
    "OutVariable", "OutBuffer", "PipelineVariable"))
CMD_SCRIPT_RE = re.compile(r"%~dp0(?P<script>[\w.-]+\.ps1)", re.I)

# Repository path -> source text, or None when no such tool is tracked.
SourceReader = Callable[[str], "str | None"]


def code_segments(text: str) -> Iterator[tuple[int, str]]:
    """Yield (line number, code) for each fenced code line and inline code span.

    A fenced line that ends in a continuation (a space and then a backslash or
    PowerShell's backtick) is joined with the next one and yielded at the line
    where it starts. HTML comment blocks are skipped, like everything GitHub
    does not render; prose is not read at all.
    """
    fence: tuple[str, int] | None = None
    in_comment = False
    pending: tuple[int, str] | None = None
    for n, line in enumerate(text.split("\n"), 1):
        if in_comment:
            in_comment = "-->" not in line
            continue
        if fence is not None:
            if closes_fence(line, fence):
                fence = None
                if pending is not None:
                    yield pending
                    pending = None
                continue
            start, joined = pending if pending is not None else (n, "")
            joined += line.rstrip("\r")
            if re.search(r"\s[\\`]$", joined):
                pending = (start, joined[:-1])
                continue
            pending = None
            yield start, joined
            continue
        fence = fence_opener(line)
        if fence is not None:
            continue
        if HTML_COMMENT_BLOCK_START_RE.match(line):
            in_comment = "-->" not in line[line.index("<!--") + 4:]
            continue
        for m in INLINE_CODE_RE.finditer(line):
            yield n, m.group(0).strip("`")
    if pending is not None:
        yield pending


def command_arguments(tail: str) -> list[str]:
    """The arguments after a tool path, up to the end of its command."""
    arguments: list[str] = []
    for token in COMMAND_TOKEN_RE.findall(tail):
        if token in COMMAND_STOP_TOKENS or token.startswith("#"):
            break
        arguments.append(token)
    return arguments


def python_tool_options(source: str) -> set[str]:
    """The options a Python tool accepts, read from its source without running it.

    Importing a tool to build its parser would run its module code inside the
    gate, so the source is parsed instead. A tool that uses argparse accepts the
    option strings of its add_argument calls, plus -h and --help unless its
    ArgumentParser sets add_help=False. A tool that reads sys.argv by hand
    accepts the option-shaped string literals of its comparisons (== and in),
    which is how such a parser recognises an argument.
    An option assembled at run time is not seen, and is reported as unknown
    rather than passed unchecked. Raises SyntaxError for unparsable source.
    """
    tree = ast.parse(source)
    declared: set[str] = set()
    uses_argparse = False
    adds_help = True
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        name = func.attr if isinstance(func, ast.Attribute) else func.id if isinstance(func, ast.Name) else ""
        if name == "add_argument":
            uses_argparse = True
            declared.update(arg.value for arg in node.args
                            if isinstance(arg, ast.Constant) and isinstance(arg.value, str)
                            and arg.value.startswith("-"))
        elif name == "ArgumentParser":
            adds_help = adds_help and not any(
                keyword.arg == "add_help" and isinstance(keyword.value, ast.Constant)
                and keyword.value.value is False for keyword in node.keywords)
    if uses_argparse:
        return declared | ({"-h", "--help"} if adds_help else set())
    return {constant.value for node in ast.walk(tree) if isinstance(node, ast.Compare)
            for operand in (node.left, *node.comparators) for constant in ast.walk(operand)
            if isinstance(constant, ast.Constant) and isinstance(constant.value, str)
            and PYTHON_OPTION_RE.fullmatch(constant.value)}


def powershell_code(source: str) -> str:
    """Return PowerShell source without comments and with every string emptied.

    What is left is code only, so a comma, a bracket or a $name inside a
    comment or a string cannot be taken for part of a parameter list.
    """
    out: list[str] = []
    i, n = 0, len(source)
    while i < n:
        c = source[i]
        if source.startswith("<#", i):
            end = source.find("#>", i + 2)
            i = n if end < 0 else end + 2
            out.append(" ")
        elif c == "#":
            end = source.find("\n", i)
            i = n if end < 0 else end
        elif c in "\"'":
            j = i + 1
            while j < n and source[j] != c:
                j += 2 if source[j] == "`" and c == '"' else 1
            out.append(c + c)
            i = j + 1
        else:
            out.append(c)
            i += 1
    return "".join(out)


def split_top_level(text: str) -> list[str]:
    """Split text at the commas outside every (), [] and {}."""
    pieces: list[str] = []
    depth = 0
    current: list[str] = []
    for c in text:
        if c in "([{":
            depth += 1
        elif c in ")]}":
            depth -= 1
        if c == "," and depth == 0:
            pieces.append("".join(current))
            current = []
        else:
            current.append(c)
    pieces.append("".join(current))
    return pieces


def powershell_parameters(source: str) -> set[str]:
    """The parameters a PowerShell script declares, lower-cased.

    Read from the script's own param() block, the first one in the code once
    comments are removed: a function's param() block comes later and declares
    the function's parameters, not the script's. An advanced script (one with
    [CmdletBinding()] or a [Parameter()] attribute) also accepts the common
    parameters, and -WhatIf and -Confirm when it supports ShouldProcess. A
    script without a param() block takes no parameters.
    """
    code = powershell_code(source)
    m = POWERSHELL_SCRIPT_PARAM_RE.search(code)
    if m is None:
        return set()
    depth, end = 1, m.end()
    while end < len(code) and depth:
        depth += {"(": 1, ")": -1}.get(code[end], 0)
        end += 1
    block = code[m.end():end - 1]
    names: set[str] = set()
    for piece in split_top_level(block):
        while True:
            stripped = re.sub(r"\[[^\[\]]*\]", " ", piece)
            if stripped == piece:
                break
            piece = stripped
        variable = POWERSHELL_VARIABLE_RE.search(piece)
        if variable:
            names.add(variable.group(1).casefold())
    binding = m.group("binding")
    if binding is not None or re.search(r"\[Parameter\b", block, re.I):
        names |= POWERSHELL_COMMON_PARAMETERS
    if binding is not None and re.search(r"SupportsShouldProcess", binding, re.I):
        names |= {"whatif", "confirm"}
    return names


def cmd_forward_target(rel: str, source: str) -> str | None:
    """The PowerShell script a .cmd launcher passes its own arguments (%*) to."""
    targets = set(CMD_SCRIPT_RE.findall(source))
    if "%*" not in source or len(targets) != 1:
        return None
    directory = os.path.dirname(rel)
    return f"{directory}/{targets.pop()}" if directory else targets.pop()


def powershell_findings(script: str, rel: str, source: str, arguments: list[str],
                        read_source: SourceReader) -> list[str]:
    """Report each -Parameter that the script, or the script a .cmd runs, does not declare."""
    used = [token.split(":")[0] for token in arguments
            if POWERSHELL_PARAMETER_RE.fullmatch(token.split(":")[0])]
    if not used:
        return []
    if rel.endswith(".cmd"):
        target = cmd_forward_target(rel, source)
        if target is None:
            return [f"{script} passes no arguments on, so {name} does nothing" for name in used]
        forwarded = read_source(target)
        if forwarded is None:
            return [f"{script} runs {target}, which is not tracked"]
        script, source = target, forwarded
    declared = powershell_parameters(source)
    return [f"{script} declares no parameter {name}" for name in used
            if name[1:].casefold() not in declared]


def invocation_findings(script: str, tail: str, read_source: SourceReader) -> list[str]:
    """Check one documented invocation of a tool: the tool, then its options."""
    rel = script.replace("\\", "/").removeprefix("./")
    source = read_source(rel)
    if source is None:
        return [f"{script}: no such tool is tracked"]
    arguments = command_arguments(tail)
    if not rel.endswith(".py"):
        return powershell_findings(script, rel, source, arguments, read_source)
    try:
        accepted = python_tool_options(source)
    except SyntaxError as exc:
        return [f"{script} cannot be parsed to read its options: {exc.msg}"]
    return [f"{script} accepts no option {option}" for option in
            (token.split("=")[0] for token in arguments) if PYTHON_OPTION_RE.fullmatch(option)
            and option not in accepted]


def documented_command_findings(rel: str, text: str, read_source: SourceReader) -> list[str]:
    """Report every quoted tool invocation of a document that would not run as written."""
    findings: list[str] = []
    for n, segment in code_segments(text):
        matches = list(TOOL_PATH_RE.finditer(segment))
        for index, m in enumerate(matches):
            end = matches[index + 1].start() if index + 1 < len(matches) else len(segment)
            findings.extend(f"{rel}:{n}: {finding}"
                            for finding in invocation_findings(m.group(0), segment[m.end():end], read_source))
    return findings


def check_documented_commands() -> None:
    """Hold every tool invocation the documentation quotes as code to the tool.

    Each tracked Markdown file except the changelog is read for tools/*.py,
    tools/*.ps1 and tools/*.cmd paths in fenced code and inline code. The tool
    must be tracked; each --option given to a Python tool must be one it
    accepts (python_tool_options); each -Parameter given to a PowerShell
    script, directly or through a .cmd launcher that passes its arguments on,
    must be declared in the script's param() block, in any letter case.
    Values, placeholders such as <sha> and the options of the program that
    runs the tool (powershell.exe -File, python3) are not checked.
    """
    tracked = git_tracked_files()
    if tracked is None:
        return
    known = set(tracked)

    def read_source(rel: str) -> str | None:
        if rel not in known or not os.path.isfile(os.path.join(REPO, rel)):
            return None
        return read(rel).decode("utf-8-sig", errors="replace")

    for rel in sorted(r for r in tracked if r.endswith(".md") and r not in DOCUMENTED_COMMANDS_EXCLUDED):
        for finding in documented_command_findings(rel, read(rel).decode("utf-8"), read_source):
            fail("documented-commands", finding)


DOCUMENTED_COMMANDS_SELFTEST_TOOLS = {
    "tools/run_check.py": (
        "import argparse\n"
        "parser = argparse.ArgumentParser()\n"
        "parser.add_argument('--check', action='store_true')\n"
        "parser.add_argument('-r', '--root')\n"),
    "tools/by_hand.py": (
        "import subprocess, sys\n"
        "if sys.argv[1:] == ['--selftest']:\n"
        "    subprocess.run(['git', 'ls-files', '--stage'])\n"),
    "tools/Run-Thing.ps1": (
        "<#\n.SYNOPSIS\nparam($NotAParameter)\n#>\n"
        "[CmdletBinding()]\n"
        "param(\n"
        "    [Parameter(Mandatory = $true)]\n"
        "    [ValidateSet(\"A, $B\", \"C)\")]\n"
        "    [string]$WorkbookPath,  # $Commented, too\n"
        "    [switch]$NewWorkbook = $false\n"
        ")\n"
        "function Inner { param($Inner) }\n"),
    "tools/Run-Thing.cmd": "@echo off\npowershell.exe -File \"%~dp0Run-Thing.ps1\" %*\n",
    "tools/Fixed.cmd": "@echo off\npowershell.exe -File \"%~dp0Run-Thing.ps1\" -NewWorkbook\n",
}

DOCUMENTED_COMMANDS_SELFTEST_CASES = [
    ("a valid invocation passes",
     "```text\npython3 tools/run_check.py --check --root <dir>\n```\n", 0),
    ("an inline-code invocation is read",
     "Run `python3 tools/run_check.py --chek` first.\n", 1),
    ("a valid inline-code invocation passes",
     "Run `py -3 tools\\run_check.py --check` first.\n", 0),
    ("a missing script is reported",
     "```text\npython3 tools/gone.py --check\n```\n", 1),
    ("an unknown option is reported",
     "```text\npython tools/run_check.py --check --chek\n```\n", 1),
    ("an option given with = is checked by its name",
     "`python3 tools/run_check.py --root=C:\\x --rot=y`\n", 1),
    ("argparse adds --help", "`python3 tools/run_check.py --help`\n", 0),
    ("an option a hand-written parser compares with passes",
     "`python3 tools/by_hand.py --selftest`\n", 0),
    ("an option a hand-written parser never compares with is reported",
     "`python3 tools/by_hand.py --write --stage`\n", 2),
    ("values, placeholders and short options are not options",
     "```text\npython3 tools/run_check.py <bundle> -r \"--quoted\" --check ...\n```\n", 0),
    ("a valid PowerShell invocation passes",
     "```powershell\npowershell.exe -NoProfile -File .\\tools\\Run-Thing.ps1 "
     "-WorkbookPath \"C:\\w.xlsm\" -NewWorkbook\n```\n", 0),
    ("PowerShell parameters are case-insensitive and include the common ones",
     "`tools/Run-Thing.ps1 -workbookpath x -NewWorkbook:$true -Verbose`\n", 0),
    ("an unknown PowerShell parameter is reported",
     "```powershell\n.\\tools\\Run-Thing.ps1 -WorkbookPath x -Mode Demo\n```\n", 1),
    ("comments, strings and inner functions declare no script parameter",
     "`tools\\Run-Thing.ps1 -Inner 1 -NotAParameter 2 -Commented 3 -B 4`\n", 4),
    ("a .cmd passes its arguments on to its script",
     "`tools\\Run-Thing.cmd -NewWorkbook`\n", 0),
    ("an unknown parameter through a .cmd is reported",
     "`tools\\Run-Thing.cmd -Bogus`\n", 1),
    ("a .cmd that passes nothing on takes no parameters",
     "`tools\\Fixed.cmd -NewWorkbook`\n", 1),
    ("a missing .cmd is reported", "Double-click `tools\\Gone.cmd`.\n", 1),
    ("two commands on one line are read apart",
     "```text\npython3 tools/run_check.py --check && python3 tools/by_hand.py --selftest\n```\n", 0),
    ("a continued fenced line is one command",
     "```bash\npython3 tools/run_check.py \\\n  --chek\n```\n", 1),
    ("prose is not code", "Run python3 tools/gone.py --check by hand.\n", 0),
    ("a path that merely ends in tools/ is not a tool", "`docs/tools/gone.py`\n", 0),
    ("an HTML comment is not code", "<!-- `python3 tools/gone.py` -->\n", 0),
]


def check_documented_commands_selftest() -> None:
    """Run the documented-command rules against their own fixtures.

    A rule that stopped finding invocations would report nothing and look
    healthy on the real documentation; the failing fixtures catch that.
    """
    for label, text, expected in DOCUMENTED_COMMANDS_SELFTEST_CASES:
        found = documented_command_findings("fixture.md", text, DOCUMENTED_COMMANDS_SELFTEST_TOOLS.get)
        if len(found) != expected:
            fail("documented-commands-selftest",
                 f"{label}: expected {expected} finding(s), got {len(found)}"
                 + (f" | {found}" if found else ""))


# Independent reviews are internal working material and are not carried by this
# repository from v1.1.3 onward. The pattern is deliberately the whole family
# rather than the two archives that were withdrawn: naming individual files
# would readmit the next review by the simple fact of it having a new name.
#
# This does not unpublish anything. Both withdrawn archives remain reachable at
# the v1.1.1 and v1.1.2 tags and in this branch's own history, and rewriting
# that would invalidate published tag SHAs — including the tree hashes the
# reviews themselves cite as what they reviewed. The rule is about what the
# project carries forward, not a claim about what was never public.
PRIVATE_PATTERNS = [
    ("*INDEPENDENT_CODE_REVIEW*", "independent reviews are internal and are not "
                                  "carried by this repository from v1.1.3 onward"),
    ("*_PRIVATE.*", "documents marked private are not published here"),
    ("*_CONFIDENTIAL.*", "documents marked confidential are not published here"),
]


def denies(rel: str, pattern: str) -> bool:
    """Match a path against a denylist pattern, ignoring case.

    fnmatch is case-sensitive on Linux, which is where the gate runs. A file
    named independent_code_review_v1.1.2.md passed every pattern while being
    exactly the document they exist to exclude, and it would have been caught
    only on a maintainer's Windows machine — the one place the check is
    optional.
    """
    rel = rel.casefold()
    pattern = pattern.casefold()
    return (fnmatch.fnmatchcase(rel, pattern)
            or fnmatch.fnmatchcase(os.path.basename(rel), pattern))


def denylist_findings(tracked: Iterable[str], present: Iterable[str]) -> list[tuple[str, str, str]]:
    """Return (path, reason, where) for every denied document."""
    findings: list[tuple[str, str, str]] = []

    for pattern, why in PRIVATE_PATTERNS:
        for rel in sorted(tracked):
            if denies(rel, pattern):
                findings.append((rel, why, "tracked"))

        for rel in sorted(set(present) - set(tracked)):
            if denies(rel, pattern):
                findings.append((rel, why, "working tree"))

    return findings


def check_private_review_denylist() -> None:
    """Fail when a private document is tracked or present in the tree.

    Both are checked. The index catches it after a commit; the working tree
    catches it before one, which is the only moment the mistake is cheap.
    """
    tracked = git_tracked_files()
    if tracked is None:
        return

    present: set[str] = set()
    for root, dirs, names in os.walk(REPO):
        dirs[:] = [x for x in dirs if x != ".git"]
        for name in names:
            present.add(os.path.relpath(os.path.join(root, name), REPO)
                        .replace(os.sep, "/"))

    for rel, why, where in denylist_findings(tracked, present):
        if where == "tracked":
            fail("private-review", f"{rel} is tracked, but {why}")
        else:
            fail("private-review",
                 f"{rel} is present in the working tree, but {why}; "
                 f"move it outside the repository before committing")


DENYLIST_SELFTEST_CASES = [
    ("the withdrawn v1.1.0 archive is denied",
     "docs/INDEPENDENT_CODE_REVIEW_V1.1.0_2026-08-19.md", True),
    ("the withdrawn v1.1.1 archive is denied",
     "docs/INDEPENDENT_CODE_REVIEW_V1.1.1_2026-08-20.md", True),
    ("the private v1.1.2 review is denied",
     "docs/INDEPENDENT_CODE_REVIEW_V1.1.2_2026-08-26.md", True),
    ("a future review is denied by the family pattern",
     "docs/INDEPENDENT_CODE_REVIEW_V2.0.0_2027-01-01.md", True),
    ("a lowercase variant is denied",
     "docs/independent_code_review_v1.1.2.md", True),
    ("a mixed-case variant is denied",
     "docs/Independent_Code_Review_v1.1.2.md", True),
    ("a review at the repository root is denied",
     "INDEPENDENT_CODE_REVIEW_V1.1.2_draft.md", True),
    ("a review in an unexpected directory is denied",
     "notes/scratch/INDEPENDENT_CODE_REVIEW_V1.1.2.md", True),
    ("a document marked private is denied", "docs/NOTES_PRIVATE.md", True),
    ("a document marked confidential is denied",
     "docs/AUDIT_CONFIDENTIAL.pdf", True),
    ("a lowercase private marker is denied", "docs/notes_private.md", True),
    ("the Ribbon behaviour note is allowed",
     "docs/RIBBON_SDI_BEHAVIOR.md", False),
    ("the changelog is allowed", "CHANGELOG.md", False),
    ("a module is allowed", "src/M_EXCEL_UI.bas", False),
    ("a file merely mentioning review is allowed",
     "docs/CODE_REVIEW_CHECKLIST.md", False),
    ("a private-sounding word inside a name is allowed",
     "docs/PRIVATE_MODULE_NOTES.md", False),
]


def check_private_review_denylist_selftest() -> None:
    """Run the denylist against its own fixtures.

    A denylist with no fixtures is a list of strings nobody has ever seen
    reject anything. Both directions matter: a pattern that denies everything
    would look identical to one that works, until it blocked a legitimate
    document and someone widened it back.
    """
    for label, rel, denied in DENYLIST_SELFTEST_CASES:
        hit = bool(denylist_findings([rel], [rel]))
        if hit != denied:
            fail("private-review-selftest",
                 f"{label}: {rel} should have been "
                 f"{'denied' if denied else 'allowed'}")


REPOSITORY_URL = "https://github.com/danielep71/VBA-EXCEL_UI"

UNRELEASED = "Unreleased"

# A level-two heading that starts with a bracket or a version number is a
# version section, so "## 1.2.0 - 2026-01-01" is reported as malformed rather
# than skipped as prose.
VERSION_SECTION_RE = re.compile(r"^##\s+(?:\[(?P<label>[^\]]*)\]|v?[0-9])")

RELEASE_HEADING_RE = re.compile(
    r"^## \[(?P<version>[0-9]+\.[0-9]+\.[0-9]+(?:-[0-9A-Za-z.\-]+)?(?:\+[0-9A-Za-z.\-]+)?)\]"
    r" - (?P<date>[0-9]{4}-[0-9]{2}-[0-9]{2})$")

LINK_REFERENCE_RE = re.compile(r"^ {0,3}\[(?P<label>[^\]]+)\]:\s*(?P<url>\S+)")


PreReleaseKey = tuple[tuple[int, int, str], ...]


def version_key(version: str) -> tuple[int, int, int, int, PreReleaseKey]:
    """Order release numbers by Semantic Versioning precedence.

    A pre-release sorts before its release. Pre-release identifiers compare
    one dot-separated field at a time: numeric fields numerically and below
    alphanumeric ones, alphanumeric fields as text, and a shorter list first
    when every shared field is equal, so rc.2 < rc.10 and rc < rc.1. Build
    metadata after "+" does not affect precedence.
    """
    core, _, pre = version.partition("+")[0].partition("-")
    major, minor, patch = (int(part) for part in core.split("."))
    fields: PreReleaseKey = tuple(
        (0, int(field), "") if field.isdigit() else (1, 0, field)
        for field in pre.split(".")
    ) if pre else ()
    return major, minor, patch, 0 if pre else 1, fields


def release_link_finding(version: str, older: str | None, url: str) -> str | None:
    """Return a problem with one release's link reference, or None.

    Two forms are accepted, the two this changelog already uses: a compare
    view from the next older release to this one, and the tag page itself,
    which is the only form the oldest release can have.
    """
    tag = f"{REPOSITORY_URL}/releases/tag/v{version}"
    if url == tag:
        return None
    m = re.fullmatch(re.escape(REPOSITORY_URL) + r"/compare/v(?P<base>[^/]+)\.\.\.v(?P<head>[^/]+)", url)
    if m and m.group("head") == version and (older is None or m.group("base") == older):
        return None
    expected = f"{REPOSITORY_URL}/compare/v{older}...v{version}" if older else tag
    return f"[{version}] links to {url}; expected {expected} or {tag}"


def changelog_findings(text: str, version: str | None) -> list[str]:  # noqa: C901
    """Return every structural problem with CHANGELOG.md, and with VERSION against it.

    The changelog and VERSION are the release authority that the wiki badges
    and release notes derive from, so a misordered heading, an impossible
    date, a missing or foreign link, or a VERSION that names no release would
    otherwise pass and propagate. The rules are listed one after another in a
    single pass over the document, which is clearer than splitting them.
    """
    findings: list[str] = []
    sections: list[tuple[int, str, str]] = []
    references: dict[str, tuple[int, str]] = {}

    # Fenced examples and HTML comment blocks are not rendered, so a release
    # heading or link reference inside one is not part of the changelog.
    for n, line in rendered_markdown_lines(text):
        section = VERSION_SECTION_RE.match(line)
        if section:
            sections.append((n, line, section.group("label") or ""))
        reference = LINK_REFERENCE_RE.match(line)
        if reference:
            label = reference.group("label")
            if label in references:
                findings.append(f"CHANGELOG.md:{n}: link reference [{label}] is defined twice")
            references[label] = (n, reference.group("url"))

    if not sections or sections[0][1] != f"## [{UNRELEASED}]":
        found = f"{sections[0][1]!r} at line {sections[0][0]}" if sections else "no version section"
        findings.append(f"CHANGELOG.md: the first version section must be '## [{UNRELEASED}]'; found {found}")

    releases: list[tuple[int, str, date | None]] = []
    for n, line, label in sections:
        if label == UNRELEASED and line == f"## [{UNRELEASED}]":
            if n != sections[0][0]:
                findings.append(f"CHANGELOG.md:{n}: '## [{UNRELEASED}]' must be the first version section")
            continue
        heading = RELEASE_HEADING_RE.match(line)
        if not heading:
            findings.append(f"CHANGELOG.md:{n}: malformed release heading {line!r}; "
                            f"expected '## [X.Y.Z] - YYYY-MM-DD'")
            continue
        try:
            released: date | None = date.fromisoformat(heading.group("date"))
        except ValueError:
            findings.append(f"CHANGELOG.md:{n}: [{heading.group('version')}] has an invalid date "
                            f"{heading.group('date')!r}")
            released = None
        releases.append((n, heading.group("version"), released))

    for (_, newer, newer_date), (n, older, older_date) in zip(releases, releases[1:], strict=False):
        if version_key(older) >= version_key(newer):
            findings.append(f"CHANGELOG.md:{n}: [{older}] follows [{newer}]; releases must be newest first")
        elif newer_date and older_date and older_date > newer_date:
            findings.append(f"CHANGELOG.md:{n}: [{older}] is dated {older_date}, after the newer "
                            f"[{newer}] ({newer_date})")

    for label, (n, url) in sorted(references.items(), key=lambda item: item[1][0]):
        if not url.startswith(REPOSITORY_URL + "/"):
            findings.append(f"CHANGELOG.md:{n}: link reference [{label}] points outside this "
                            f"repository: {url}")

    for i, (n, release, _) in enumerate(releases):
        link = references.get(release)
        if link is None:
            findings.append(f"CHANGELOG.md:{n}: release [{release}] has no link reference")
            continue
        previous = releases[i + 1][1] if i + 1 < len(releases) else None
        problem = release_link_finding(release, previous, link[1])
        if problem and link[1].startswith(REPOSITORY_URL + "/"):
            findings.append(f"CHANGELOG.md:{link[0]}: {problem}")

    unreleased = references.get(UNRELEASED)
    if releases and unreleased is None:
        findings.append(f"CHANGELOG.md: [{UNRELEASED}] has no link reference; expected "
                        f"[{UNRELEASED}]: {REPOSITORY_URL}/compare/v{releases[0][1]}...HEAD")
    elif unreleased and releases:
        expected = f"{REPOSITORY_URL}/compare/v{releases[0][1]}...HEAD"
        if unreleased[1] != expected and unreleased[1].startswith(REPOSITORY_URL + "/"):
            findings.append(f"CHANGELOG.md:{unreleased[0]}: [{UNRELEASED}] links to {unreleased[1]}; "
                            f"expected {expected}")

    if version is not None:
        if not releases:
            findings.append(f"VERSION is {version}, but CHANGELOG.md has no release heading")
        elif version != releases[0][1]:
            findings.append(f"VERSION is {version}, but the newest release in CHANGELOG.md is "
                            f"[{releases[0][1]}]; bump both together")

    return findings


# --------------------------------------------------------------------------
REVIEW_ID_RE = re.compile(r"\bICR-UI-(?:11[12]-)?P[0-9]-[0-9]{2}\b")

CROSSWALK_HEADING = "### 🔎 Review finding crosswalk"

CROSSWALK_ROW_RE = re.compile(
    r"^\| `(?P<id>[^`]+)` \| (?P<review>v1\.1\.[0-2]) \| \[#(?P<issue>[0-9]+)\]"
    r"\([^)]*/issues/(?P=issue)\) \| (?:`(?P<current>[^`]+)`|—) \| .+ \|$")


def crosswalk_findings(changelog: str, cited: dict[str, list[str]]) -> list[str]:
    """Check the review-finding crosswalk in CHANGELOG.md against the citations.

    ``cited`` maps each review identifier found in a tracked file to the files
    that cite it. Every cited identifier must appear in exactly one table row,
    each row must name an issue, and a renumbered v1.1.1 identifier must point
    at a row of the table.
    """
    start = changelog.find(CROSSWALK_HEADING + "\n")
    if start < 0:
        return [f"CHANGELOG.md has no '{CROSSWALK_HEADING}' section"]
    end = len(changelog)
    for marker in ("\n### ", "\n## "):
        at = changelog.find(marker, start + len(CROSSWALK_HEADING))
        if 0 <= at < end:
            end = at
    section = changelog[start:end]

    findings: list[str] = []
    rows: dict[str, str | None] = {}
    for line in section.splitlines():
        if not REVIEW_ID_RE.search(line) or not line.startswith("| `"):
            continue
        m = CROSSWALK_ROW_RE.match(line)
        if not m:
            findings.append(f"CHANGELOG.md crosswalk: malformed row: {line[:80]}")
            continue
        ident = m.group("id")
        if not REVIEW_ID_RE.fullmatch(ident):
            findings.append(f"CHANGELOG.md crosswalk: {ident!r} is not a review identifier")
        elif ident in rows:
            findings.append(f"CHANGELOG.md crosswalk: {ident} has more than one row")
        else:
            rows[ident] = m.group("current")

    for ident, current in sorted(rows.items()):
        if current is not None and current not in rows:
            findings.append(f"CHANGELOG.md crosswalk: {ident} points at {current}, "
                            "which has no row of its own")
    for ident in sorted(set(cited) - set(rows)):
        where = ", ".join(sorted(set(cited[ident]))[:3])
        findings.append(f"{ident} is cited in {where} but has no row in the "
                        "CHANGELOG.md review finding crosswalk")
    return findings


def check_review_crosswalk() -> None:
    """Every review identifier cited in the repository resolves through one row."""
    tracked = git_tracked_files()
    if tracked is None:
        return
    cited: dict[str, list[str]] = {}
    for rel in tracked:
        if not rel.lower().endswith((".md", ".bas", ".cls", ".frm")):
            continue
        text = read(rel).decode("utf-8", errors="replace")
        for ident in REVIEW_ID_RE.findall(text):
            cited.setdefault(ident, []).append(rel)
    for finding in crosswalk_findings(read("CHANGELOG.md").decode("utf-8"), cited):
        fail("review-crosswalk", finding)


def check_release_state() -> None:
    """Reject known stale release markers and a malformed root VERSION.

    This is the check that would have caught the tagged README still carrying a
    Release Candidate badge and an unchecked release checklist. It also checks
    the structure of CHANGELOG.md and that VERSION names its newest release.
    """
    readme = read("README.md").decode("utf-8")

    for marker in ("Release_Candidate", "Release Candidate"):
        if marker in readme:
            fail("release-state", f"README.md still contains {marker!r}")

    if re.search(r"^\[ \] ", readme, re.M):
        fail("release-state", "README.md contains an unchecked release checklist")

    changelog = read("CHANGELOG.md").decode("utf-8")
    if "_Nothing yet._" in changelog:
        fail("release-state", "CHANGELOG.md contains an unfilled placeholder section")

    # The root VERSION file is the repository's release authority and the wiki
    # badge check derives its expected value from it. Module-header alignment is
    # separate release work; this check must not imply those stamps already
    # derive from VERSION. A malformed value must fail here explicitly.
    #
    # Read it defensively. required-files already reports a missing VERSION one
    # check earlier, and reading it unconditionally here turned that finding
    # into an unhandled exception that killed the run before any finding was
    # printed.
    path = os.path.join(REPO, "VERSION")
    version: str | None = None
    if os.path.isfile(path):
        raw = read("VERSION").decode("utf-8")
        problems = version_findings(raw)
        for finding in problems:
            fail("release-state", finding)
        # A malformed VERSION is reported above; comparing it with the
        # changelog as well would only restate the same finding.
        if not problems:
            version = raw.strip()

    for finding in changelog_findings(changelog, version):
        fail("release-state", finding)


# --------------------------------------------------------------------------
def git_tracked_files() -> list[str] | None:
    """Return the exact index-backed file set, or fail closed.

    Walking the working tree is incorrect here: an ignored local demo workbook
    or release asset is permitted to exist, but a force-added copy must fail the
    gate. ``git ls-files`` makes that distinction mechanically.
    """
    try:
        result = subprocess.run(
            ["git", "-C", REPO, "ls-files", "-z"],
            capture_output=True,
            check=False,
        )
    except OSError as exc:
        fail("repository-hygiene", f"cannot enumerate tracked files: {exc}")
        return None

    if result.returncode:
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        fail("repository-hygiene",
             f"git ls-files failed ({result.returncode}): {detail}")
        return None

    return sorted(
        os.fsdecode(path)
        for path in result.stdout.split(b"\0")
        if path
    )


def git_ignored_files(paths: list[str]) -> list[str]:
    """Return tracked paths that the current .gitignore would hide."""
    if not paths:
        return []

    payload = b"\0".join(os.fsencode(path) for path in paths) + b"\0"
    try:
        result = subprocess.run(
            ["git", "-C", REPO, "check-ignore", "--no-index", "-z", "--stdin"],
            input=payload,
            capture_output=True,
            check=False,
        )
    except OSError as exc:
        fail("repository-hygiene", f"cannot evaluate .gitignore: {exc}")
        return []

    if result.returncode not in (0, 1):
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        fail("repository-hygiene",
             f"git check-ignore failed ({result.returncode}): {detail}")
        return []

    return sorted(
        os.fsdecode(path)
        for path in result.stdout.split(b"\0")
        if path
    )


def forbidden_tracked_reason(path: str) -> str | None:
    """Return why a tracked path is forbidden, or None when it is allowed."""
    name = path.rsplit("/", 1)[-1].lower()
    for pattern, reason in FORBIDDEN_TRACKED_ARTIFACTS:
        if fnmatch.fnmatchcase(name, pattern.lower()):
            return reason
    return None


def check_repository_hygiene() -> None:
    """Keep .gitignore, the Git index, and the release-binary policy aligned."""
    ignore_rules = {
        line.strip()
        for line in read(".gitignore").decode("utf-8").splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    }

    for pattern, _ in FORBIDDEN_TRACKED_ARTIFACTS:
        if pattern not in ignore_rules:
            fail("repository-hygiene",
                 f".gitignore is missing required blocked pattern {pattern!r}")

    tracked = git_tracked_files()
    if tracked is None:
        return

    forbidden_paths: set[str] = set()
    for rel in tracked:
        reason = forbidden_tracked_reason(rel)
        if reason:
            forbidden_paths.add(rel)
            fail("repository-hygiene", f"{rel}: {reason} must not be tracked")

    for rel in git_ignored_files(tracked):
        if rel not in forbidden_paths:
            fail("repository-hygiene",
                 f"{rel}: tracked file is hidden by the current .gitignore")

    # The name denylist above gives the more specific message, so a path it
    # already rejected is not reported a second time as an unapproved binary.
    for finding in tracked_binary_findings([r for r in tracked if r not in forbidden_paths]):
        fail("repository-hygiene", finding)


# Binaries are allowed in exactly two ways. Anything under an approved
# directory is documentation media by rule. Anything else must be listed in
# BINARY_ALLOWLIST with the SHA-256 of its exact content, so replacing it later
# fails the gate until the allowlist is updated in the same, reviewed change.
APPROVED_BINARY_DIRECTORIES = ("images/",)

BINARY_ALLOWLIST = "tools/binary_allowlist.txt"

ALLOWLIST_LINE_RE = re.compile(r"^(?P<sha>[0-9a-f]{64})\s+(?P<path>\S(?:.*\S)?)\s*$")


def git_text_attributes(paths: list[str]) -> dict[str, str] | None:
    """Return the value of Git's text attribute for each path, or fail closed."""
    if not paths:
        return {}

    payload = b"\0".join(os.fsencode(path) for path in paths) + b"\0"
    try:
        result = subprocess.run(
            ["git", "-C", REPO, "check-attr", "-z", "--stdin", "text"],
            input=payload,
            capture_output=True,
            check=False,
        )
    except OSError as exc:
        fail("repository-hygiene", f"cannot evaluate .gitattributes: {exc}")
        return None

    if result.returncode:
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        fail("repository-hygiene",
             f"git check-attr failed ({result.returncode}): {detail}")
        return None

    fields = [os.fsdecode(f) for f in result.stdout.split(b"\0")]
    return {fields[i]: fields[i + 2] for i in range(0, len(fields) - 2, 3)}


def binary_reason(attribute: str, data: bytes) -> str | None:
    """Return why a file is binary, or None when it is text.

    Either signal is enough. The attribute catches a file Git is told to treat
    as binary whatever it contains; the NUL byte catches binary content under
    a text extension, such as a workbook renamed notes.txt or a damaged module.
    """
    if attribute == "unset":
        return "Git treats it as binary (text attribute unset)"
    if b"\0" in data:
        return "its content contains a NUL byte"
    return None


def read_binary_allowlist() -> tuple[dict[str, str], list[str]]:
    """Return ({path: sha256}, findings) from BINARY_ALLOWLIST.

    A missing allowlist is an empty one; required-files reports its absence.
    """
    entries: dict[str, str] = {}
    findings: list[str] = []
    if not os.path.isfile(os.path.join(REPO, BINARY_ALLOWLIST)):
        return entries, findings

    for n, raw in enumerate(read(BINARY_ALLOWLIST).decode("utf-8").splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        m = ALLOWLIST_LINE_RE.match(line)
        if not m:
            findings.append(f"{BINARY_ALLOWLIST}:{n}: expected '<sha256>  <path>', got {line!r}")
        elif m.group("path") in entries:
            findings.append(f"{BINARY_ALLOWLIST}:{n}: {m.group('path')} is listed twice")
        else:
            entries[m.group("path")] = m.group("sha")
    return entries, findings


def tracked_binary_findings(tracked: list[str]) -> list[str]:
    """Report every tracked binary that is neither in an approved directory nor allowlisted."""
    attributes = git_text_attributes(tracked)
    if attributes is None:
        return []

    allowlist, findings = read_binary_allowlist()

    for rel in tracked:
        path = os.path.join(REPO, rel)
        if not os.path.isfile(path):
            continue                 # deleted in the working tree; nothing to classify
        data = read(rel)
        reason = binary_reason(attributes.get(rel, "unspecified"), data)
        if reason is None or rel.startswith(APPROVED_BINARY_DIRECTORIES):
            continue
        expected = allowlist.get(rel)
        if expected is None:
            findings.append(
                f"{rel}: binary file outside {', '.join(APPROVED_BINARY_DIRECTORIES)} "
                f"({reason}); move it there or list it with its SHA-256 in "
                f"{BINARY_ALLOWLIST}")
            continue
        actual = hashlib.sha256(data).hexdigest()
        if actual != expected:
            findings.append(
                f"{rel}: allowlisted binary has changed (SHA-256 {actual}, "
                f"{BINARY_ALLOWLIST} records {expected}); review the new content "
                f"and update the allowlist in the same change")

    tracked_set = set(tracked)
    for rel in sorted(allowlist):
        if rel not in tracked_set:
            findings.append(f"{BINARY_ALLOWLIST}: {rel} is listed but not tracked")

    return findings


# --------------------------------------------------------------------------
# ATX headings only: up to three spaces of indentation, one to six hashes, and
# then a space or the end of the line. An optional closing run of hashes is not
# part of the heading text.
ATX_HEADING_RE = re.compile(r"^ {0,3}#{1,6}(?:[ \t]+(?P<text>.*?))?(?:[ \t]+#+)?[ \t]*$")

CODE_SPAN_RE = re.compile(r"(?P<ticks>`+)(?P<code>.+?)(?P=ticks)")

CODE_PLACEHOLDER_RE = re.compile("\ue000([0-9]+)\ue001")

IMAGE_RE = re.compile(r"!\[[^\]]*\]\([^)]*\)")

LINK_RE = re.compile(r"\[(?P<text>[^\]]*)\](?:\([^)]*\)|\[[^\]]*\])")

AUTOLINK_RE = re.compile(r"<(?P<url>[A-Za-z][A-Za-z0-9+.\-]*:[^\s<>]*)>")

HTML_TAG_RE = re.compile(r"</?[A-Za-z][^>]*>")

# An underscore run opens emphasis only where it is not inside a word, so
# UI_SetExcelUI keeps its underscore while _note_ loses both.
UNDERSCORE_EMPHASIS_RE = re.compile(r"(?<![\w\\])(?P<run>_{1,3})(?=\S)(?P<body>.+?)(?<=\S)(?P=run)(?!\w)")


def heading_plain_text(markdown: str) -> str:
    """Return a heading's text as GitHub renders it, without inline Markdown.

    Code spans keep their content and lose their backticks. Images render no
    text. Links keep their text. HTML tags are dropped, entities decoded, and
    underscore emphasis delimiters removed; every other markup character is
    punctuation and is dropped when the anchor is made.
    """
    # Each code span becomes a placeholder, so the markup around it is read
    # as one piece: in _`Notes`_ the underscores still pair up as emphasis.
    # The placeholders use private-use characters that no other rule touches,
    # and the code content is put back only after the inline rules ran.
    codes: list[str] = []

    def hold(span: re.Match[str]) -> str:
        code = span.group("code")
        if len(code) > 2 and code.startswith(" ") and code.endswith(" ") and code.strip():
            code = code[1:-1]
        codes.append(code)
        return f"\ue000{len(codes) - 1}\ue001"

    text = _inline_text(CODE_SPAN_RE.sub(hold, markdown))
    text = CODE_PLACEHOLDER_RE.sub(lambda m: codes[int(m.group(1))], text)
    return text.strip()


def _inline_text(text: str) -> str:
    text = IMAGE_RE.sub("", text)
    text = LINK_RE.sub(lambda m: m.group("text"), text)
    text = AUTOLINK_RE.sub(lambda m: m.group("url"), text)
    text = INLINE_HTML_COMMENT_RE.sub("", text)
    text = HTML_TAG_RE.sub("", text)
    previous = None
    while previous != text:
        previous = text
        text = UNDERSCORE_EMPHASIS_RE.sub(lambda m: m.group("body"), text)
    return html.unescape(text)


def github_anchor(heading: str) -> str:
    """Return the anchor GitHub generates for one heading, before de-duplication.

    Lowercase the rendered text, keep Unicode letters, marks and decimal
    digits, the underscore and other connector punctuation, hyphens and
    spaces, drop everything else (emoji included), and turn each space into a
    hyphen. Nothing is collapsed, so "Café — notes" becomes "café--notes".
    """
    kept = []
    for ch in heading_plain_text(heading).lower():
        category = unicodedata.category(ch)
        if ch in " -" or category[0] in "LM" or category in ("Nd", "Pc"):
            kept.append(ch)
    return "".join(kept).replace(" ", "-")


HTML_COMMENT_BLOCK_START_RE = re.compile(r"^ {0,3}<!--")

INLINE_HTML_COMMENT_RE = re.compile(r"<!--.*?-->")


def rendered_markdown_lines(text: str) -> Iterator[tuple[int, str]]:
    """Yield (line number, line) for each line GitHub renders as Markdown.

    Read in one pass over fences and comments, so it serves both the heading
    anchors and the changelog's release headings and link references.

    Two kinds of block hide their lines, and each one's delimiters are plain
    content inside the other:

    - a fenced code block, from its opening fence to a closing fence;
    - an HTML comment block, which starts on a line that begins with <!--
      (after at most three spaces) and ends on the line that contains -->.

    A comment that opens anywhere else on a line is inline HTML. It belongs to
    that line only and hides nothing below it, even if it is never closed,
    because the next line's block structure is read first. So <!-- inside a
    code span, or after other text, never changes the state.
    """
    fence: tuple[str, int] | None = None
    in_comment = False
    for n, line in enumerate(text.split("\n"), 1):
        if in_comment:
            in_comment = "-->" not in line
            continue
        if fence is not None:
            if closes_fence(line, fence):
                fence = None
            continue
        fence = fence_opener(line)
        if fence is not None:
            continue
        if HTML_COMMENT_BLOCK_START_RE.match(line):
            in_comment = "-->" not in line[line.index("<!--") + 4:]
            continue
        yield n, line


def heading_anchors(text: str) -> list[str]:
    """Return every heading anchor in a document, in order, as GitHub assigns them.

    A repeated anchor gets -1, -2 and so on, and a suffixed anchor that
    collides with one already taken is suffixed again, so every anchor in the
    document is unique and each link resolves to one heading. Headings inside
    fenced code blocks or HTML comment blocks are not headings: GitHub renders
    neither, so they must neither create an anchor nor shift the numbering of
    a visible duplicate.
    """
    anchors: list[str] = []
    counts: dict[str, int] = {}
    taken: set[str] = set()
    for _, line in rendered_markdown_lines(text):
        m = ATX_HEADING_RE.match(line)
        if not m:
            continue
        base = github_anchor(m.group("text") or "")
        anchor = base
        while anchor in taken:
            counts[base] = counts.get(base, 0) + 1
            anchor = f"{base}-{counts[base]}"
        taken.add(anchor)
        anchors.append(anchor)
    return anchors


def check_markdown_links() -> None:
    md_files: list[str] = []
    for root, dirs, files in os.walk(REPO):
        dirs[:] = [d for d in dirs if d != ".git"]
        md_files += [
            os.path.relpath(os.path.join(root, f), REPO)
            for f in files
            if f.endswith(".md")
        ]

    for rel in sorted(md_files):
        text = read(rel).decode("utf-8", errors="replace")
        base = os.path.dirname(os.path.join(REPO, rel))

        anchors = set(re.findall(r'<a id="([^"]+)"></a>', text))
        anchors.update(heading_anchors(text))

        for target in re.findall(r"\]\((?!https?:)([^)]+)\)", text):
            path_part, _, anchor = target.partition("#")
            if not path_part:
                if anchor and unquote(anchor) not in anchors:
                    fail("markdown-links", f"{rel}: broken internal anchor #{anchor}")
                continue
            if not os.path.exists(os.path.normpath(os.path.join(base, path_part))):
                fail("markdown-links", f"{rel}: missing link target {path_part}")


# --------------------------------------------------------------------------
def check_formatter() -> None:
    sys.path.insert(0, os.path.join(REPO, "tools"))
    try:
        import reformat
    except Exception as exc:                                  # pragma: no cover
        fail("formatter", f"tools/reformat.py could not be imported: {exc}")
        return

    for rel in ALL_MODULES:
        if not is_standard_module(rel):
            # The formatter's rules are written for standard modules. A class
            # or form export has a header it would rewrite, so it fails here
            # explicitly until the formatter learns that layout.
            fail("formatter", f"{rel}: the house-style formatter supports "
                              f"standard (.bas) modules only")
            continue
        path = os.path.join(REPO, rel)
        name = reformat.module_name_of(path)
        if name is None:
            fail("formatter", f"{rel}: no Attribute VB_Name")
            continue
        try:
            expected = reformat.reformat(path, name).encode("ascii")
        except (ValueError, OSError) as exc:
            fail("formatter", f"{rel}: {exc}")
            continue
        if read(rel) != expected:
            fail("formatter", f"{rel}: not in house-style normal form "
                              f"(run tools/reformat.py --write)")


def check_formatter_selftest() -> None:
    """Run the formatter's own rule fixtures.

    The formatter is the one tool here whose defects the VBA suite cannot see,
    and --check passing proves only that today's modules contain no construct
    that trips it. Running the fixtures on every push is what turns that into
    a gate rather than a hope.
    """
    sys.path.insert(0, os.path.join(REPO, "tools"))
    try:
        import reformat
    except Exception as exc:                                  # pragma: no cover
        fail("formatter-selftest",
             f"tools/reformat.py could not be imported: {exc}")
        return

    if not hasattr(reformat, "selftest"):
        fail("formatter-selftest", "tools/reformat.py exposes no selftest()")
        return

    for finding in reformat.selftest():
        fail("formatter-selftest", finding.replace("\n", " | "))


# --------------------------------------------------------------------------
def check_repository_policy() -> None:
    import repository_policy
    for finding in repository_policy.check(REPO):
        fail('repository policy', finding)


def check_repository_policy_selftest() -> None:
    import repository_policy
    findings, count = repository_policy.selftest(REPO)
    for finding in findings:
        fail('repository policy self-test', finding)


def check_label_catalogue() -> None:
    """docs/ISSUE_LABELS.md must state the catalogue in .github/labels.json."""
    try:
        labels = json.loads(read(".github/labels.json").decode("utf-8"))["labels"]
        guide = read("docs/ISSUE_LABELS.md").decode("utf-8")
    except (OSError, ValueError, KeyError) as exc:
        fail("label-catalogue", f"cannot read the catalogue: {exc}")
        return

    count = len(labels)
    if f"badge/labels-{count}-" not in guide:
        fail("label-catalogue", f"docs/ISSUE_LABELS.md badge does not show {count} labels")
    if f"defines {count} issue labels" not in guide:
        fail("label-catalogue", f"docs/ISSUE_LABELS.md does not state {count} labels")

    rows = re.findall(r"^\| `([^`]+)` \| `#([0-9A-F]{6})` \| (.+?) \|$", guide, re.M)
    expected = [(label["name"], label["color"], label["description"]) for label in labels]
    if rows != expected:
        missing = sorted(set(expected) - set(rows))
        extra = sorted(set(rows) - set(expected))
        detail = "; ".join(
            [f"missing or different: {m[0]}" for m in missing]
            + [f"not in labels.json: {e[0]}" for e in extra]
        ) or "rows are out of order"
        fail("label-catalogue", f"docs/ISSUE_LABELS.md catalogue table: {detail}")


CHECKS: list[tuple[str, Callable[[], None]]] = [
    ("repository policy", check_repository_policy),
    ("repository policy self-test", check_repository_policy_selftest),
    ("module inventory", check_module_inventory),
    ("VBA project manifest", check_vba_project_manifest),
    ("required files", check_required_files),
    ("module names", check_module_names),
    ("option policy", check_option_policy),
    ("encoding and line endings", check_encoding_and_endings),
    ("banner rule widths", check_banner_rules),
    ("house labels", check_house_labels),
    ("VBA analyzer", check_vba_analyzer),
    ("VBA analyzer self-test", check_vba_analyzer_selftest),
    ("title-bar control regression", check_titlebar_control),
    ("WinAPI error capture", check_winapi_error_capture),
    ("demo cleanup regression", check_demo_cleanup),
    ("public API manifest", check_public_api),
    ("public API self-test", check_public_api_selftest),
    ("supported API declaration", check_supported_api_declaration),
    ("workflow pin policy", check_workflow_policy),
    ("workflow policy self-test", check_workflow_policy_selftest),
    ("branch coverage", check_branch_coverage),
    ("branch coverage self-test", check_branch_coverage_selftest),
    ("wiki badge self-test", check_wiki_badge_selftest),
    ("release state", check_release_state),
    ("review finding crosswalk", check_review_crosswalk),
    ("release state self-test", check_release_state_selftest),
    ("repository hygiene", check_repository_hygiene),
    ("markdown links", check_markdown_links),
    ("markdown escapes", check_markdown_escapes),
    ("markdown escape self-test", check_markdown_escape_selftest),
    ("documented commands", check_documented_commands),
    ("documented commands self-test", check_documented_commands_selftest),
    ("label catalogue", check_label_catalogue),
    ("private review denylist", check_private_review_denylist),
    ("private review denylist self-test", check_private_review_denylist_selftest),
    ("house-style formatter", check_formatter),
    ("formatter self-test", check_formatter_selftest),
]


USAGE = """usage:
  check_repo.py           run every static check against this checkout
  check_repo.py --help    print this message and exit

Exit status is 0 when every check passed, 1 when any check failed, and 2 for
an unknown argument.
"""


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    if args in (["-h"], ["--help"]):
        print(USAGE, end="")
        return 0
    if args:
        print(USAGE, end="", file=sys.stderr)
        return 2

    print(f"VBA-EXCEL_UI static gate\nrepository: {REPO}\n")

    for label, fn in CHECKS:
        before = len(failures)

        # A check that raises used to escape main() entirely: the traceback
        # went to stderr, sys.exit(main()) never ran, and the process exited 0.
        # A continuous-integration step reading only the status code would call
        # that a pass. A crash is a finding.
        try:
            fn()
        except Exception as exc:                              # noqa: BLE001
            fail(label, f"the check itself raised {type(exc).__name__}: {exc}")

        added = len(failures) - before
        print(f"  {'FAIL' if added else 'ok  '}  {label}"
              + (f"  ({added})" if added else ""))

    print()
    if notes:
        print(f"{len(notes)} note(s):\n")
        for n in notes:
            print(f"  {n}")
        print()

    if failures:
        print(f"{len(failures)} finding(s):\n")
        for f in failures:
            print(f"  {f}")
        print("\nRESULT: FAIL")
        return 1

    print("RESULT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
