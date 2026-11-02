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

Exit status is 0 when every check passed, 1 otherwise.
"""

import fnmatch
import hashlib
import os
import re
import subprocess
import sys
from collections.abc import Callable, Iterable
from types import ModuleType

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

# The production set as I document it, in the INSTALLATION.md import order.
# It is data, not the list the checks iterate: the module-inventory check fails
# when the discovered src/ set disagrees with it, so adding or removing a
# production module has to be a deliberate change here and in INSTALLATION.md.
EXPECTED_PRODUCTION_MODULES = [
    "src/M_EXCEL_UI_RUNTIME.bas",
    "src/M_EXCEL_UI_TITLEBAR.bas",
    "src/M_EXCEL_UI_SNAPSHOT.bas",
    "src/M_EXCEL_UI.bas",
]

# Modules that other checks read by name: the title-bar and demo-cleanup
# regressions inspect these files directly, so they must exist.
REQUIRED_MODULES = EXPECTED_PRODUCTION_MODULES + [
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
    "tools/demo_cleanup_fixtures.py",
    "tools/reformat_fixtures.py",
    "tools/vba_api.py",
    "tools/wiki_badges.py",
    "tools/repository_policy.py",
    "tools/public_api_manifest.txt",
    "tools/binary_allowlist.txt",
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


def check_module_inventory() -> None:
    """Discover the modules every other module check iterates, and audit them.

    Three things fail here: a VBA module tracked outside src/, test/ or demo/;
    a production set that disagrees with EXPECTED_PRODUCTION_MODULES; and a
    production module that tools/vba_api.py assigns to no public API section,
    whose public declarations would otherwise never reach the manifest.
    """
    for rel in load_modules():
        fail("module-inventory",
             f"{rel}: VBA module tracked outside src/, test/ or demo/; it "
             f"would escape every module check")

    discovered = set(PRODUCTION_MODULES)
    expected = set(EXPECTED_PRODUCTION_MODULES)
    for rel in sorted(discovered - expected):
        fail("module-inventory",
             f"{rel}: production module is not in EXPECTED_PRODUCTION_MODULES; "
             f"add it there and to the INSTALLATION.md import order")
    for rel in sorted(expected - discovered):
        fail("module-inventory",
             f"{rel}: listed in EXPECTED_PRODUCTION_MODULES but not tracked")

    vba_api = import_vba_api()
    if vba_api is None:
        return
    sectioned = {rel for _, modules in vba_api.SECTIONS for rel in modules}
    for rel in sorted(discovered - sectioned):
        fail("module-inventory",
             f"{rel}: production module in no public API section of "
             f"tools/vba_api.py; its declarations would never be recorded")


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
                    fail('labels', f'{rel}:{n}: non-house label {code[:-1]!r}')


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
# Characters a Markdown renderer treats as escapable. A backslash before any of
# them is markup, not text.
ESCAPABLE = r"[\\`*_{}\[\]()#+\-.!|<>&\"'~=:;,/$%^?@]"

ESCAPE_RE = re.compile(r"\\(" + ESCAPABLE + ")")

FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})")

INLINE_CODE_RE = re.compile(r"`+[^`]*`+")


def strip_inline_code(line: str) -> str:
    """Blank out inline code spans, preserving length so columns stay true."""
    return INLINE_CODE_RE.sub(lambda m: " " * len(m.group(0)), line)


def escape_findings(rel: str, text: str) -> list[str]:
    r"""Report Markdown punctuation escaped outside code.

    A WYSIWYG editor escapes every character it thinks is markup, and both
    independent review archives arrived with several hundred of them: headings
    read `## 1\. Executive assessment`, module names read `M\_EXCEL\_UI`, and a
    horizontal rule read `\---`. The documents rendered as literal backslashes
    and nobody noticed for two releases.

    The check is fence-aware, and deliberately does not look inside fenced
    blocks or inline code. A backslash there is data — a Windows path, a regular
    expression — and flagging it would make the gate wrong about correct
    documents. This is a deliberate boundary: the rule detects characteristic
    escaping in prose, not every possible backslash a WYSIWYG editor can emit.
    """
    findings: list[str] = []
    fence: tuple[str, int] | None = None   # (character, opening length)

    for n, line in enumerate(text.split("\n"), 1):
        m = FENCE_RE.match(line)
        if m:
            token = m.group(1)
            if fence is None:
                fence = (token[0], len(token))
            elif token[0] == fence[0] and len(token) >= fence[1]:
                # A closing fence must use the same character and be at least
                # as long as the one that opened it. Collapsing every opener to
                # three let a ``` line close a ```` block, and everything after
                # it was then read as prose that happened to look fenced.
                fence = None
            continue

        if fence is not None:
            continue

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


def check_release_state() -> None:
    """Reject known stale release markers and a malformed root VERSION.

    This is the check that would have caught the tagged README still carrying a
    Release Candidate badge and an unchecked release checklist.
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
    if not os.path.isfile(path):
        return

    for finding in version_findings(read("VERSION").decode("utf-8")):
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
        for heading in re.findall(r"^#{1,6}\s+(.+)$", text, re.M):
            slug = re.sub(r"[^a-z0-9 -]", "", heading.lower()).replace(" ", "-")
            anchors.add(slug)

        for target in re.findall(r"\]\((?!https?:)([^)]+)\)", text):
            path_part, _, anchor = target.partition("#")
            if not path_part:
                if anchor and anchor not in anchors:
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


CHECKS: list[tuple[str, Callable[[], None]]] = [
    ("repository policy", check_repository_policy),
    ("repository policy self-test", check_repository_policy_selftest),
    ("module inventory", check_module_inventory),
    ("required files", check_required_files),
    ("module names", check_module_names),
    ("option policy", check_option_policy),
    ("encoding and line endings", check_encoding_and_endings),
    ("banner rule widths", check_banner_rules),
    ("house labels", check_house_labels),
    ("VBA analyzer", check_vba_analyzer),
    ("VBA analyzer self-test", check_vba_analyzer_selftest),
    ("title-bar control regression", check_titlebar_control),
    ("demo cleanup regression", check_demo_cleanup),
    ("public API manifest", check_public_api),
    ("public API self-test", check_public_api_selftest),
    ("supported API declaration", check_supported_api_declaration),
    ("workflow pin policy", check_workflow_policy),
    ("workflow policy self-test", check_workflow_policy_selftest),
    ("wiki badge self-test", check_wiki_badge_selftest),
    ("release state", check_release_state),
    ("release state self-test", check_release_state_selftest),
    ("repository hygiene", check_repository_hygiene),
    ("markdown links", check_markdown_links),
    ("markdown escapes", check_markdown_escapes),
    ("markdown escape self-test", check_markdown_escape_selftest),
    ("private review denylist", check_private_review_denylist),
    ("private review denylist self-test", check_private_review_denylist_selftest),
    ("house-style formatter", check_formatter),
    ("formatter self-test", check_formatter_selftest),
]


def main() -> int:
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
