#!/usr/bin/env python3
"""
Static release gate for VBA-EXCEL_UI.

Everything here can be decided from the repository text alone, without Excel.
That is the whole design constraint: a hosted runner has no Office, so the gate
covers what is checkable there and deliberately does not pretend to replace
Test_EXCEL_UI_RunReleaseCertification, which is the behavioural gate.

Each check is independent and reports every finding it has rather than stopping
at the first, because a contributor fixing one problem per push is a bad use of
their time.

Exit status is 0 when every check passed, 1 otherwise.
"""

import fnmatch
import os
import re
import subprocess
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

PRODUCTION_MODULES = [
    "src/M_EXCEL_UI.bas",
    "src/M_EXCEL_UI_RUNTIME.bas",
    "src/M_EXCEL_UI_SNAPSHOT.bas",
    "src/M_EXCEL_UI_TITLEBAR.bas",
]

TEST_MODULES = ["test/M_EXCEL_UI_REGRESSION_TESTS.bas"]

DEMO_MODULES = [
    "demo/M_DEMO_BUILDER.bas",
    "demo/M_EXCEL_UI_DEMO.bas",
]

ALL_MODULES = PRODUCTION_MODULES + TEST_MODULES + DEMO_MODULES

REQUIRED_FILES = ALL_MODULES + [
    ".gitattributes",
    ".gitignore",
    "README.md",
    "CHANGELOG.md",
    "INSTALLATION.md",
    "CONTRIBUTING.md",
    "SECURITY.md",
    "CODE_OF_CONDUCT.md",
    "LICENSE",
    "tools/reformat.py",
    "tools/vba_api.py",
    "tools/public_api_manifest.txt",
]

HOUSE_LABELS = {"Safe_Exit", "Fail_Num", "Err_Handler", "Clean_Exit", "Clean_Fail"}

PROC_RE = re.compile(r"^(Public|Private|Friend)\s+(Sub|Function|Property\s+\w+)\s+(\w+)")
END_RE = re.compile(r"^End (Sub|Function|Property)\b")
LABEL_RE = re.compile(r"^([A-Za-z_]\w*):\s*$")
JUMP_RE = re.compile(r"\b(?:GoTo|Resume)\s+([A-Za-z_]\w*)")
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

failures = []
notes = []


def fail(check, detail):
    failures.append(f"{check}: {detail}")


def read(path):
    with open(os.path.join(REPO, path), "rb") as fh:
        return fh.read()


def read_lines(path):
    return read(path).decode("latin-1").split("\r\n")


# --------------------------------------------------------------------------
def check_required_files():
    for rel in REQUIRED_FILES:
        if not os.path.exists(os.path.join(REPO, rel)):
            fail("required-files", f"missing {rel}")


# --------------------------------------------------------------------------
def check_module_names():
    """Attribute VB_Name must match the filename.

    A mismatch imports cleanly and then silently shadows or duplicates a module
    in the VBE, which is far harder to diagnose than a failed check.
    """
    for rel in ALL_MODULES:
        expected = os.path.basename(rel)[:-4]
        text = read(rel).decode("latin-1")
        m = re.match(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"', text)
        if not m:
            fail("module-name", f"{rel}: no Attribute VB_Name on line 1")
        elif m.group(1) != expected:
            fail("module-name", f"{rel}: declares {m.group(1)!r}, expected {expected!r}")


# --------------------------------------------------------------------------
def check_option_policy():
    for rel in ALL_MODULES:
        text = read(rel).decode("latin-1")
        if not re.search(r"^Option Explicit\s*$", text, re.M):
            fail("option-policy", f"{rel}: missing Option Explicit")
        if rel in PRODUCTION_MODULES + TEST_MODULES:
            if not re.search(r"^Option Private Module\s*$", text, re.M):
                fail("option-policy", f"{rel}: missing Option Private Module")


# --------------------------------------------------------------------------
def check_encoding_and_endings():
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
def check_banner_rules():
    for rel in ALL_MODULES:
        for n, line in enumerate(read_lines(rel), 1):
            if RULE_RE.match(line) and len(line) != 79:
                fail("banner-width", f"{rel}:{n}: rule is {len(line)} cols, expected 79")


# --------------------------------------------------------------------------
def check_structure():
    """Procedure pairing, directive balance, label vocabulary, jump targets."""
    for rel in ALL_MODULES:
        lines = read_lines(rel)

        depth = 0
        for n, line in enumerate(lines, 1):
            s = line.strip()
            if s.startswith("#If "):
                depth += 1
            elif s.startswith("#End If"):
                depth -= 1
                if depth < 0:
                    fail("directives", f"{rel}:{n}: #End If without #If")
                    depth = 0
        if depth:
            fail("directives", f"{rel}: unbalanced conditional compilation ({depth:+d})")

        open_procs = []
        for n, line in enumerate(lines, 1):
            m = PROC_RE.match(line)
            if m:
                # a #If/#Else pair declares the same procedure twice
                if open_procs and open_procs[-1][1] == m.group(3):
                    continue
                open_procs.append((m.group(2).split()[0], m.group(3), n))
            elif END_RE.match(line):
                if not open_procs:
                    fail("procedures", f"{rel}:{n}: {line.strip()} without opener")
                else:
                    open_procs.pop()
        for kind, name, n in open_procs:
            fail("procedures", f"{rel}: unclosed {kind} {name} opened at line {n}")

        labels = set()
        for line in lines:
            m = LABEL_RE.match(line)
            if m:
                labels.add(m.group(1))
        for bad in sorted(labels - HOUSE_LABELS):
            fail("labels", f"{rel}: non-house label {bad!r}")

        for n, line in enumerate(lines, 1):
            if line.strip().startswith("'"):
                continue
            for m in JUMP_RE.finditer(line):
                target = m.group(1)
                if target in ("Next", "0"):
                    continue
                if target not in labels:
                    fail("jumps", f"{rel}:{n}: jump to undefined label {target!r}")


# --------------------------------------------------------------------------
def check_ptrsafe():
    """Every Declare inside a VBA7 branch must carry PtrSafe.

    Omitting it compiles on 32-bit and fails on 64-bit, so the defect is
    invisible on the machine that introduced it.
    """
    for rel in ALL_MODULES:
        lines = read_lines(rel)
        vba7 = False
        for n, line in enumerate(lines, 1):
            s = line.strip()
            if s.startswith("#If VBA7"):
                vba7 = True
            elif s.startswith("#Else"):
                vba7 = False
            elif s.startswith("#End If"):
                vba7 = False
            elif vba7 and re.match(r"^(Public |Private )?Declare\s+", s):
                if "PtrSafe" not in s:
                    fail("ptrsafe", f"{rel}:{n}: Declare without PtrSafe in a VBA7 branch")


# --------------------------------------------------------------------------
def check_duplicate_procedures():
    seen = {}
    for rel in ALL_MODULES:
        local = set()
        for line in read_lines(rel):
            m = PROC_RE.match(line)
            if m:
                local.add(m.group(3))
        for name in local:
            if name in seen:
                fail("duplicates", f"{name!r} defined in both {seen[name]} and {rel}")
            else:
                seen[name] = rel


# --------------------------------------------------------------------------
def import_vba_api():
    sys.path.insert(0, os.path.join(REPO, "tools"))
    try:
        import vba_api
    except Exception as exc:                                  # pragma: no cover
        fail("public-api", f"tools/vba_api.py could not be imported: {exc}")
        return None
    return vba_api


def check_public_api():
    """Diff the full declaration contract against the versioned manifest.

    Recording only module, kind and name detected a member appearing or
    disappearing and nothing else, so a reordered parameter, a ByVal turned
    ByRef, a changed default or a renumbered enum could break every caller
    while this check stayed green. The manifest now carries the whole
    normalised declaration, and any difference at all is a finding.

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
        recorded = vba_api.parse_manifest(manifest_path)
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


def check_public_api_selftest():
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
def check_release_state():
    """Documentation must not describe a release that has not happened.

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


# --------------------------------------------------------------------------
def git_tracked_files():
    """Return the exact index-backed file set, or fail closed.

    Walking the working tree is incorrect here: an ignored local demo workbook
    or release asset is permitted to exist, but a force-added copy must fail the
    gate. ``git ls-files`` makes that distinction mechanically.
    """
    try:
        result = subprocess.run(
            ["git", "-C", REPO, "ls-files", "-z"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
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


def git_ignored_files(paths):
    """Return tracked paths that the current .gitignore would hide."""
    if not paths:
        return []

    payload = b"\0".join(os.fsencode(path) for path in paths) + b"\0"
    try:
        result = subprocess.run(
            ["git", "-C", REPO, "check-ignore", "--no-index", "-z", "--stdin"],
            input=payload,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
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


def forbidden_tracked_reason(path):
    """Return why a tracked path is forbidden, or None when it is allowed."""
    name = path.rsplit("/", 1)[-1].lower()
    for pattern, reason in FORBIDDEN_TRACKED_ARTIFACTS:
        if fnmatch.fnmatchcase(name, pattern.lower()):
            return reason
    return None


def check_repository_hygiene():
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

    forbidden_paths = set()
    for rel in tracked:
        reason = forbidden_tracked_reason(rel)
        if reason:
            forbidden_paths.add(rel)
            fail("repository-hygiene", f"{rel}: {reason} must not be tracked")

    for rel in git_ignored_files(tracked):
        if rel not in forbidden_paths:
            fail("repository-hygiene",
                 f"{rel}: tracked file is hidden by the current .gitignore")


# --------------------------------------------------------------------------
def check_markdown_links():
    md_files = []
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
def check_formatter():
    sys.path.insert(0, os.path.join(REPO, "tools"))
    try:
        import reformat
    except Exception as exc:                                  # pragma: no cover
        fail("formatter", f"tools/reformat.py could not be imported: {exc}")
        return

    for rel in ALL_MODULES:
        path = os.path.join(REPO, rel)
        name = reformat.module_name_of(path)
        if name is None:
            fail("formatter", f"{rel}: no Attribute VB_Name")
            continue
        expected = reformat.reformat(path, name).encode("latin-1")
        if read(rel) != expected:
            fail("formatter", f"{rel}: not in house-style normal form "
                              f"(run tools/reformat.py --write)")


def check_formatter_selftest():
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
CHECKS = [
    ("required files", check_required_files),
    ("module names", check_module_names),
    ("option policy", check_option_policy),
    ("encoding and line endings", check_encoding_and_endings),
    ("banner rule widths", check_banner_rules),
    ("procedure structure", check_structure),
    ("PtrSafe declarations", check_ptrsafe),
    ("duplicate procedures", check_duplicate_procedures),
    ("public API manifest", check_public_api),
    ("public API self-test", check_public_api_selftest),
    ("release state", check_release_state),
    ("repository hygiene", check_repository_hygiene),
    ("markdown links", check_markdown_links),
    ("house-style formatter", check_formatter),
    ("formatter self-test", check_formatter_selftest),
]


def main():
    print(f"VBA-EXCEL_UI static gate\nrepository: {REPO}\n")

    for label, fn in CHECKS:
        before = len(failures)
        fn()
        added = len(failures) - before
        print(f"  {'FAIL' if added else 'ok  '}  {label}"
              + (f"  ({added})" if added else ""))

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
