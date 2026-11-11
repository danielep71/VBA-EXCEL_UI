#!/usr/bin/env python3
"""Verify, describe, package and check a release of this repository.

Every subcommand reads the release from its annotated tag, so a step cannot run
against a different commit from the one the tag names:

- verify TAG: the tag exists, is annotated and equals "v" + VERSION at the
  tagged commit; CHANGELOG.md at that commit has a dated section for the
  version; the commit is reachable from origin/main; and, with --repository,
  the required check "Repository and module checks" concluded success on that
  exact commit.
- notes TAG: prints the CHANGELOG.md section of the version, the text the
  GitHub release carries.
- bundle TAG --out DIR: writes git archive output of the tag as .zip and
  .tar.gz, named VBA-EXCEL_UI-<version>, and a SHA256SUMS file. git archive
  honours the export-ignore attributes in .gitattributes, and the bundle is
  refused if an excluded path such as AGENTS.md is still in it.
- verify-published TAG --repository OWNER/NAME --dir DIR: the GitHub release
  for the tag exists and is not a draft, its body equals the notes, and its
  assets are exactly the files in DIR/SHA256SUMS plus that file, with matching
  SHA-256 values.

The GitHub REST API is read with the token in GITHUB_TOKEN, when set. Nothing
here writes to GitHub: the release workflow creates the release with gh.

Exit status: 0 when every check passed (verify without --repository prints
INCOMPLETE and also exits 0), 1 when a check failed, 2 for a usage, Git or
network error.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tarfile
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Protocol

from check_repo import LINK_REFERENCE_RE, RELEASE_HEADING_RE, VERSION_SECTION_RE, rendered_markdown_lines

ROOT = Path(__file__).resolve().parents[1]
API = "https://api.github.com"
PROJECT = "VBA-EXCEL_UI"
REQUIRED_CHECK = "Repository and module checks"
CHECK_APP = "github-actions"
SUMS = "SHA256SUMS"

# Paths that .gitattributes marks export-ignore and that a source archive must
# therefore never contain. A directory is written with its trailing slash.
EXCLUDED_FROM_ARCHIVE = ("AGENTS.md", ".github/")

# The lower-case tag form that RELEASING.md requires: "v" and a SemVer number.
TAG_RE = re.compile(
    r"^v(?P<version>[0-9]+\.[0-9]+\.[0-9]+(?:-[0-9A-Za-z.\-]+)?(?:\+[0-9A-Za-z.\-]+)?)$")

# Check-run conclusions that do not count against a release: a cancelled or
# skipped run neither passed nor failed. Every other completed conclusion
# other than success is a failure.
NOT_FAILED = {"success", "cancelled", "skipped"}

SUM_LINE_RE = re.compile(r"^(?P<hash>[0-9a-f]{64})  (?P<name>[^/\s][^/]*)$")


class ReleaseError(Exception):
    """A problem that stops a subcommand before it can decide anything."""


class NotFound(Exception):
    """The GitHub API answered 404."""


class Api(Protocol):
    """The two GitHub requests this tool makes; tests inject a fake."""

    def get_json(self, path: str) -> Any:
        """Return the decoded JSON at an API path; raise NotFound on 404."""

    def download(self, url: str) -> bytes:
        """Return the bytes of a release asset from its API URL."""


@dataclass
class GitHubApi:
    """The GitHub REST API, read with an optional token."""

    token: str
    base: str = API

    def headers(self, accept: str) -> dict[str, str]:
        return {"Accept": accept, "X-GitHub-Api-Version": "2022-11-28",
                "User-Agent": f"{PROJECT}-release-tool"}

    def open(self, url: str, accept: str) -> bytes:
        request = urllib.request.Request(url, headers=self.headers(accept))
        if self.token:
            # An asset download redirects to a storage host that must not
            # receive the token, and an unredirected header is not forwarded.
            request.add_unredirected_header("Authorization", f"Bearer {self.token}")
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                data: bytes = response.read()
                return data
        except urllib.error.HTTPError as error:
            if error.code == 404:
                raise NotFound(url) from error
            raise

    def get_json(self, path: str) -> Any:
        return json.loads(self.open(self.base + path, "application/vnd.github+json"))

    def download(self, url: str) -> bytes:
        return self.open(url, "application/octet-stream")


# --------------------------------------------------------------------------
# Git


def git(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(["git", *args], cwd=root, capture_output=True, check=False)


def git_text(root: Path, *args: str) -> str:
    """Run Git and return its output, or raise ReleaseError when it fails."""
    result = git(root, *args)
    if result.returncode != 0:
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        raise ReleaseError(f"git {' '.join(args)} failed: {detail}")
    return result.stdout.decode("utf-8")


def tag_version(tag: str) -> str:
    """Return the version a tag names, or raise ReleaseError for another form."""
    match = TAG_RE.match(tag)
    if not match:
        raise ReleaseError(f"{tag!r} is not a release tag; expected a lower-case 'vX.Y.Z'")
    return match.group("version")


def tag_commit(root: Path, tag: str) -> tuple[str, str] | None:
    """Return (object type, commit SHA) of a tag, or None when it does not exist."""
    ref = f"refs/tags/{tag}"
    if git(root, "rev-parse", "--verify", "--quiet", ref).returncode != 0:
        return None
    kind = git_text(root, "cat-file", "-t", ref).strip()
    sha = git_text(root, "rev-parse", "--verify", f"{ref}^{{commit}}").strip()
    return kind, sha


def file_at(root: Path, sha: str, path: str) -> str | None:
    """Return a file's text at a commit, or None when the commit lacks it."""
    result = git(root, "show", f"{sha}:{path}")
    if result.returncode != 0:
        return None
    return result.stdout.decode("utf-8")


# --------------------------------------------------------------------------
# Changelog


def release_heading(changelog: str, version: str) -> tuple[int, str] | None:
    """Return (line number, date) of the version's release heading, if any."""
    for n, line in rendered_markdown_lines(changelog):
        heading = RELEASE_HEADING_RE.match(line)
        if heading and heading.group("version") == version:
            return n, heading.group("date")
    return None


def release_notes(changelog: str, version: str) -> str | None:
    """Return the body of the version's changelog section, or None.

    The section runs from its heading to the next version heading. The link
    references that end the file and the rule that separates sections belong
    to no release, so they are removed from its end.
    """
    found = release_heading(changelog, version)
    if found is None:
        return None
    start = found[0]
    lines = changelog.split("\n")
    end = len(lines)
    for n, line in rendered_markdown_lines(changelog):
        if n > start and VERSION_SECTION_RE.match(line):
            end = n - 1
            break
    body = lines[start:end]
    while body and (not body[-1].strip() or body[-1].strip() == "---"
                    or LINK_REFERENCE_RE.match(body[-1])):
        body.pop()
    while body and not body[0].strip():
        body.pop(0)
    return "\n".join(body) + "\n" if body else None


def tag_notes(root: Path, tag: str) -> str:
    """Return the changelog section of a tag's version, read at the tagged commit."""
    version = tag_version(tag)
    resolved = tag_commit(root, tag)
    if resolved is None:
        raise ReleaseError(f"tag {tag} does not exist")
    changelog = file_at(root, resolved[1], "CHANGELOG.md")
    notes = release_notes(changelog, version) if changelog is not None else None
    if notes is None:
        raise ReleaseError(f"CHANGELOG.md at {tag} has no section for {version}")
    return notes


def normalized_notes(text: str) -> str:
    """Compare release bodies without line-ending or end-of-text differences."""
    return text.replace("\r\n", "\n").strip()


# --------------------------------------------------------------------------
# verify


def local_findings(root: Path, tag: str, main_ref: str) -> tuple[list[str], str | None]:
    """Return the findings Git can decide, and the tagged commit when it exists."""
    version = tag_version(tag)
    resolved = tag_commit(root, tag)
    if resolved is None:
        return [f"tag {tag} does not exist"], None
    kind, sha = resolved
    findings: list[str] = []
    if kind != "tag":
        findings.append(f"tag {tag} is a lightweight tag; a release tag must be annotated "
                        "(git tag -a)")

    raw = file_at(root, sha, "VERSION")
    if raw is None:
        findings.append(f"VERSION does not exist at {sha}")
    elif raw.strip() != version:
        findings.append(f"tag {tag} does not match VERSION {raw.strip()!r} at {sha}; "
                        f"expected tag v{raw.strip()}")

    changelog = file_at(root, sha, "CHANGELOG.md")
    heading = release_heading(changelog, version) if changelog is not None else None
    if changelog is None:
        findings.append(f"CHANGELOG.md does not exist at {sha}")
    elif heading is None:
        findings.append(f"CHANGELOG.md at {sha} has no '## [{version}] - YYYY-MM-DD' section")
    else:
        try:
            date.fromisoformat(heading[1])
        except ValueError:
            findings.append(f"CHANGELOG.md at {sha}: [{version}] has an invalid date {heading[1]!r}")
        if release_notes(changelog, version) is None:
            findings.append(f"CHANGELOG.md at {sha}: the [{version}] section is empty")

    if git(root, "rev-parse", "--verify", "--quiet", f"{main_ref}^{{commit}}").returncode != 0:
        findings.append(f"{main_ref} does not exist; fetch it before verifying")
    elif git(root, "merge-base", "--is-ancestor", sha, main_ref).returncode != 0:
        findings.append(f"{sha} is not reachable from {main_ref}; tag only a commit on main")
    return findings, sha


def check_run_findings(api: Api, repository: str, sha: str) -> list[str]:
    """Require a successful run of the required check on exactly this commit.

    Only the latest run of each check suite is read, so a job that failed and
    then passed on a re-run counts as passed. Runs from apps other than GitHub
    Actions are ignored, because any app with checks access can name a run.
    """
    query = urllib.parse.urlencode({"check_name": REQUIRED_CHECK, "filter": "latest",
                                    "per_page": "100"})
    data = api.get_json(f"/repos/{repository}/commits/{sha}/check-runs?{query}")
    if not isinstance(data, dict) or not isinstance(data.get("check_runs"), list):
        raise ReleaseError("the check-runs endpoint did not return a check_runs list")
    runs: list[dict[str, Any]] = data["check_runs"]
    if int(data.get("total_count", len(runs))) > len(runs):
        raise ReleaseError(f"{data.get('total_count')} check runs; refusing to judge a partial list")

    relevant = [run for run in runs
                if run.get("name") == REQUIRED_CHECK and run.get("head_sha") == sha
                and (run.get("app") or {}).get("slug") == CHECK_APP]
    if not relevant:
        return [f"no '{REQUIRED_CHECK}' check run exists on {sha}"]

    completed = [run for run in relevant if run.get("status") == "completed"]
    failed = sorted({str(run.get("conclusion")) for run in completed
                     if run.get("conclusion") not in NOT_FAILED})
    findings: list[str] = []
    if failed:
        findings.append(f"'{REQUIRED_CHECK}' concluded {', '.join(failed)} on {sha}")
    if not any(run.get("conclusion") == "success" for run in completed):
        states = sorted({str(run.get("conclusion") or run.get("status")) for run in relevant})
        findings.append(f"'{REQUIRED_CHECK}' has not concluded success on {sha} "
                        f"(found: {', '.join(states)})")
    return findings


def verify(root: Path, tag: str, main_ref: str, api: Api | None,
           repository: str | None) -> tuple[list[str], str | None]:
    """Return every finding, and the tagged commit SHA when the tag exists."""
    findings, sha = local_findings(root, tag, main_ref)
    if sha is not None and api is not None and repository is not None:
        findings += check_run_findings(api, repository, sha)
    return findings, sha


# --------------------------------------------------------------------------
# bundle


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def render_sums(hashes: dict[str, str]) -> str:
    """Return SHA256SUMS text in the format sha256sum --check reads."""
    return "".join(f"{digest}  {name}\n" for name, digest in sorted(hashes.items()))


def parse_sums(text: str) -> dict[str, str]:
    """Return {file name: SHA-256} from SHA256SUMS text, or raise ReleaseError."""
    hashes: dict[str, str] = {}
    for n, line in enumerate(text.splitlines(), 1):
        match = SUM_LINE_RE.match(line)
        if not match:
            raise ReleaseError(f"{SUMS}:{n}: malformed line {line!r}")
        if match.group("name") in hashes or match.group("name") == SUMS:
            raise ReleaseError(f"{SUMS}:{n}: {match.group('name')} is listed more than once or "
                               "lists the sums file itself")
        hashes[match.group("name")] = match.group("hash")
    if not hashes:
        raise ReleaseError(f"{SUMS} lists no file")
    return hashes


def archive_files(path: Path) -> set[str]:
    """Return the file members of a .zip or .tar.gz archive, without directories."""
    if path.name.endswith(".zip"):
        with zipfile.ZipFile(path) as archive:
            return {name for name in archive.namelist() if not name.endswith("/")}
    with tarfile.open(path, "r:gz") as archive:
        return {member.name for member in archive.getmembers() if not member.isdir()}


def archive_findings(name: str, members: Iterable[str], prefix: str) -> list[str]:
    """Return problems with an archive's member list."""
    findings: list[str] = []
    files = sorted(members)
    outside = [member for member in files if not member.startswith(prefix)]
    if outside:
        findings.append(f"{name}: {len(outside)} member(s) outside {prefix}, such as {outside[0]}")
    inside = {member[len(prefix):] for member in files if member.startswith(prefix)}
    for excluded in EXCLUDED_FROM_ARCHIVE:
        hits = sorted(member for member in inside
                      if member == excluded or (excluded.endswith("/") and member.startswith(excluded)))
        if hits:
            findings.append(f"{name}: contains {hits[0]}, which .gitattributes marks export-ignore")
    if "VERSION" not in inside:
        findings.append(f"{name}: does not contain VERSION")
    return findings


def bundle(root: Path, tag: str, out: Path) -> list[Path]:
    """Write the source archives and SHA256SUMS for a tag and return their paths."""
    version = tag_version(tag)
    if tag_commit(root, tag) is None:
        raise ReleaseError(f"tag {tag} does not exist")
    if out.exists() and any(out.iterdir()):
        raise ReleaseError(f"{out} is not empty; bundle into a new directory")
    out.mkdir(parents=True, exist_ok=True)

    stem = f"{PROJECT}-{version}"
    prefix = f"{stem}/"
    written: list[Path] = []
    for fmt, suffix in (("zip", ".zip"), ("tar.gz", ".tar.gz")):
        target = out / f"{stem}{suffix}"
        git_text(root, "archive", f"--format={fmt}", f"--prefix={prefix}",
                 f"--output={target}", f"refs/tags/{tag}")
        written.append(target)

    contents = [archive_files(path) for path in written]
    findings = [finding for path, files in zip(written, contents, strict=True)
                for finding in archive_findings(path.name, files, prefix)]
    if contents[0] != contents[1]:
        findings.append(f"{written[0].name} and {written[1].name} hold different files")
    if findings:
        raise ReleaseError("; ".join(findings))

    sums = out / SUMS
    sums.write_text(render_sums({path.name: sha256_bytes(path.read_bytes()) for path in written}),
                    encoding="utf-8", newline="\n")
    return [*written, sums]


# --------------------------------------------------------------------------
# verify-published


def published_findings(api: Api, repository: str, tag: str, notes: str,
                       sums_text: str) -> list[str]:
    """Compare the published GitHub release with the notes and SHA256SUMS."""
    expected = parse_sums(sums_text)
    try:
        release = api.get_json(f"/repos/{repository}/releases/tags/{urllib.parse.quote(tag, safe='')}")
    except NotFound:
        return [f"no published GitHub release exists for {tag}"]
    if not isinstance(release, dict):
        raise ReleaseError("the release endpoint did not return an object")

    findings: list[str] = []
    if release.get("draft"):
        findings.append(f"the release for {tag} is still a draft")
    if release.get("tag_name") != tag:
        findings.append(f"the release names tag {release.get('tag_name')!r}, not {tag}")
    if normalized_notes(str(release.get("body") or "")) != normalized_notes(notes):
        findings.append(f"the release body differs from the CHANGELOG.md section for {tag}")

    assets: dict[str, dict[str, Any]] = {str(asset.get("name")): asset
                                         for asset in release.get("assets") or []}
    wanted = set(expected) | {SUMS}
    for name in sorted(wanted - set(assets)):
        findings.append(f"asset {name} is missing from the release")
    for name in sorted(set(assets) - wanted):
        findings.append(f"asset {name} is not listed in {SUMS}")

    for name in sorted(wanted & set(assets)):
        asset = assets[name]
        if asset.get("state", "uploaded") != "uploaded":
            findings.append(f"asset {name} is in state {asset.get('state')!r}, not uploaded")
            continue
        actual = sha256_bytes(api.download(str(asset["url"])))
        want = sha256_bytes(sums_text.encode("utf-8")) if name == SUMS else expected[name]
        if actual != want:
            findings.append(f"asset {name} has SHA-256 {actual}; {SUMS} expects {want}")
        digest = asset.get("digest")
        if isinstance(digest, str) and digest.startswith("sha256:") and digest[7:] != actual:
            findings.append(f"asset {name}: GitHub reports {digest}, but the download hashes to {actual}")
    return findings


# --------------------------------------------------------------------------
# Command line


def repository_name(value: str) -> str:
    if not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", value):
        raise argparse.ArgumentTypeError("repository must have the form owner/name")
    return value


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    root.add_argument("--root", type=Path, default=ROOT, help="the Git repository (default: this one)")
    commands = root.add_subparsers(dest="command", required=True)

    check = commands.add_parser("verify", help="check a tag before it is published")
    check.add_argument("tag")
    check.add_argument("--repository", type=repository_name,
                       help="owner/name; also check the required check run (reads GITHUB_TOKEN)")
    check.add_argument("--main-ref", default="origin/main",
                       help="the ref the tagged commit must be reachable from (default: origin/main)")

    notes = commands.add_parser("notes", help="print the changelog section of a tag")
    notes.add_argument("tag")

    build = commands.add_parser("bundle", help="write the source archives and SHA256SUMS")
    build.add_argument("tag")
    build.add_argument("--out", type=Path, required=True, help="a new or empty directory")

    published = commands.add_parser("verify-published", help="check a published GitHub release")
    published.add_argument("tag")
    published.add_argument("--repository", type=repository_name, required=True,
                           help="owner/name (reads GITHUB_TOKEN)")
    published.add_argument("--dir", type=Path, required=True,
                           help="the directory that holds the bundle's SHA256SUMS")
    return root


def run(args: argparse.Namespace, api: Api | None) -> int:
    """Run one subcommand; api is None only when no --repository was given."""
    root: Path = args.root.resolve()
    tag: str = args.tag

    if args.command == "verify":
        findings, sha = verify(root, tag, args.main_ref, api, args.repository)
        print(f"Release verification for {tag}" + (f" at {sha}" if sha else ""))
        for finding in findings:
            print(f"  - {finding}")
        if findings:
            print("RESULT: FAIL")
            return 1
        if api is None or args.repository is None:
            print(f"  '{REQUIRED_CHECK}' was not checked; pass --repository to check it")
            print("RESULT: INCOMPLETE")
            return 0
        print("RESULT: PASS")
        return 0

    if args.command == "notes":
        sys.stdout.write(tag_notes(root, tag))
        return 0

    if args.command == "bundle":
        for path in bundle(root, tag, args.out):
            print(path)
        return 0

    if api is None:
        raise ReleaseError("verify-published needs --repository")
    sums_text = (args.dir / SUMS).read_text(encoding="utf-8")
    findings = published_findings(api, args.repository, tag, tag_notes(root, tag), sums_text)
    print(f"Published release check for {tag}")
    for finding in findings:
        print(f"  - {finding}")
    print("RESULT: " + ("FAIL" if findings else "PASS"))
    return 1 if findings else 0


def main(argv: list[str] | None = None, api: Api | None = None) -> int:
    args = parser().parse_args(argv)
    if api is None and getattr(args, "repository", None):
        api = GitHubApi(os.environ.get("GITHUB_TOKEN", ""))
    try:
        return run(args, api)
    except (ReleaseError, OSError, urllib.error.URLError, ValueError, KeyError, TypeError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
