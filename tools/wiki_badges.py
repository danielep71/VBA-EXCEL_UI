#!/usr/bin/env python3
"""
Wiki track-badge consistency check for VBA-EXCEL_UI.

Before v1.1.2 the wiki had drifted a full release behind and nothing signalled
it. Every page now carries a `wiki_tracks-vX.Y.Z` badge, but a convention held
up by memory is not a gate: a page added without a badge, or a sync that updates
thirteen pages of fourteen, reproduces the original failure while looking
correct.

This tool reads a wiki working copy from disk and asserts three things:

  * every content page carries a badge;
  * every badge names the same version;
  * that version is the one the repository states in its root VERSION file.

The wiki is a separate Git-backed repository with no Actions surface of its own,
so it cannot check itself. The clone belongs to a workflow in the main
repository; this tool receives a path and never touches the network.
tools/check_repo.py stays entirely offline for the same reason — every other
check reads the tree, and a gate that fails without network access is a gate
people route around.

GitHub's reserved navigation pages carry no content and are listed explicitly
below rather than matched by a leading-underscore pattern. A path pattern would
silently exempt any future page someone happened to name that way.

--selftest exercises the rules against fixtures on a temporary directory.
"""

import os
import re
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

VERSION_FILE = "VERSION"

# Reserved GitHub wiki navigation pages. Not required to carry a badge; if one
# does, it must agree with the rest.
NAVIGATION_PAGES = (
    "_Sidebar.md",
    "_Footer.md",
    "_Header.md",
)

BADGE_RE = re.compile(r"wiki_tracks-(?P<version>v[0-9][0-9A-Za-z.]*?)-[0-9A-Fa-f]{6}\b")

VERSION_RE = re.compile(r"^[0-9]+\.[0-9]+\.[0-9]+(?:[-+][0-9A-Za-z.\-+]*)?$")


class WikiError(Exception):
    """The check could not be run, which is a failure rather than a crash."""


def expected_version(repo=REPO):
    """Return the version the repository states, as a tag-shaped string.

    Deriving this from the first wiki page it happens to read would make the
    check circular: thirteen pages agreeing on a stale version would pass. It
    comes from the repository, which is the thing the wiki is supposed to track.
    """
    path = os.path.join(repo, VERSION_FILE)
    try:
        with open(path, encoding="utf-8") as fh:
            raw = fh.read()
    except OSError as exc:
        raise WikiError(f"cannot read {VERSION_FILE}: {exc}")

    value = raw.strip()
    if not VERSION_RE.match(value):
        raise WikiError(f"{VERSION_FILE} does not hold a release number: {value!r}")

    return "v" + value


def page_badges(text):
    """Return every track version named on a page, in order of appearance."""
    return [m.group("version") for m in BADGE_RE.finditer(text)]


def check_wiki(wiki_path, expected):
    """Return a list of findings for a wiki working copy.

    An empty list means every content page carries the expected badge and no
    page disagrees.
    """
    if not os.path.isdir(wiki_path):
        return [f"{wiki_path} is not a directory"]

    pages = sorted(
        name for name in os.listdir(wiki_path)
        if name.endswith(".md") and os.path.isfile(os.path.join(wiki_path, name))
    )

    if not pages:
        return [f"{wiki_path} contains no wiki pages; the clone is probably empty"]

    findings = []
    seen = {}

    for name in pages:
        with open(os.path.join(wiki_path, name), encoding="utf-8") as fh:
            badges = page_badges(fh.read())

        if not badges:
            if name not in NAVIGATION_PAGES:
                findings.append(
                    f"{name} carries no wiki_tracks badge; every content page "
                    f"must state the release it was written against"
                )
            continue

        distinct = sorted(set(badges))
        if len(distinct) > 1:
            findings.append(
                f"{name} carries more than one track badge: "
                f"{', '.join(distinct)}"
            )
            continue

        seen[name] = distinct[0]

    versions = sorted(set(seen.values()))

    if len(versions) > 1:
        findings.append(
            "the wiki does not agree with itself; pages name "
            + ", ".join(versions)
        )
        for name in sorted(seen):
            findings.append(f"    {name}: {seen[name]}")
        return findings

    if versions and versions[0] != expected:
        findings.append(
            f"the wiki tracks {versions[0]} and the repository states "
            f"{expected}; every page is stale by at least one release"
        )

    return findings


# --------------------------------------------------------------------------
# SELF-TEST
# --------------------------------------------------------------------------
def _badge(version):
    return (
        f"[![Version](https://img.shields.io/badge/wiki_tracks-{version}"
        f"-217346?style=flat-square)]"
        "(https://github.com/danielep71/VBA-EXCEL_UI/blob/main/CHANGELOG.md)\n"
    )


def _write_wiki(root, pages):
    os.makedirs(root, exist_ok=True)
    for name, text in pages.items():
        with open(os.path.join(root, name), "w", encoding="utf-8") as fh:
            fh.write(text)
    return root


GOOD = {
    "Home.md": _badge("v1.1.2") + "\n# Home\n",
    "Architecture.md": _badge("v1.1.2") + "\n# Architecture\n",
    "_Sidebar.md": "- [Home](Home)\n",
}


def _variant(**changes):
    pages = dict(GOOD)
    pages.update(changes)
    return pages


SELFTEST_CASES = [
    ("a fully badged wiki agreeing with VERSION passes", GOOD, "v1.1.2", 0),
    ("a navigation page needs no badge", GOOD, "v1.1.2", 0),
    ("a content page with no badge is reported",
     _variant(**{"Examples.md": "# Examples\n"}), "v1.1.2", 1),
    ("one page left behind is reported",
     _variant(**{"Architecture.md": _badge("v1.1.1") + "\n# Architecture\n"}),
     "v1.1.2", 3),
    ("a wiki that agrees with itself but not the repository is reported",
     {"Home.md": _badge("v1.1.1"), "Architecture.md": _badge("v1.1.1")},
     "v1.1.2", 1),
    ("a page carrying two different badges is reported",
     _variant(**{"Home.md": _badge("v1.1.2") + _badge("v1.1.1")}),
     "v1.1.2", 1),
    ("a navigation page disagreeing with the rest is reported",
     _variant(**{"_Sidebar.md": _badge("v1.1.1")}), "v1.1.2", 4),
    ("an empty clone is reported rather than passing vacuously",
     {}, "v1.1.2", 1),
]


def selftest():
    """Return a list of failure descriptions; empty means the rules hold."""
    failures = []

    with tempfile.TemporaryDirectory() as tmp:
        for i, (label, pages, expected, count) in enumerate(SELFTEST_CASES):
            root = _write_wiki(os.path.join(tmp, f"case{i}"), pages)
            found = check_wiki(root, expected)
            if len(found) != count:
                failures.append(
                    f"{label}: expected {count} finding(s), got {len(found)}"
                    + (f" | {found}" if found else "")
                )

    # The expected version must come from the repository, not from the wiki.
    try:
        version = expected_version()
    except WikiError as exc:
        failures.append(f"expected_version() could not read the repository: {exc}")
    else:
        if not version.startswith("v"):
            failures.append(
                f"expected_version() must return a tag-shaped string, got {version!r}"
            )

    return failures


def run_selftest():
    failures = selftest()
    if not failures:
        print(f"ok   self-test: {len(SELFTEST_CASES) + 1} wiki-badge rules hold")
        return 0
    print(f"FAIL self-test: {len(failures)} rule(s) broken\n")
    for f in failures:
        print(f"  - {f}")
    return 1


# --------------------------------------------------------------------------
USAGE = """usage:
  wiki_badges.py <wiki-path>   check a wiki working copy against VERSION
  wiki_badges.py --selftest    verify the rules against their own fixtures
"""


def main():
    args = sys.argv[1:]

    if args and args[0] == "--selftest":
        return run_selftest()

    if len(args) != 1:
        print(USAGE, file=sys.stderr)
        return 2

    try:
        expected = expected_version()
    except WikiError as exc:
        print(f"FAIL {exc}", file=sys.stderr)
        return 1

    findings = check_wiki(args[0], expected)

    if not findings:
        print(f"ok   wiki badges agree with {expected}")
        return 0

    print(f"FAIL wiki badges: {len(findings)} finding(s)\n", file=sys.stderr)
    for f in findings:
        print(f"  {f}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
