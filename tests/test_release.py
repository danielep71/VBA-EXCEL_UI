"""Tests for the release tool (tools/release.py).

Every repository here is a temporary Git repository with a fabricated
origin/main, and every GitHub response is injected, so no test needs the
network or this repository's own tags.
"""

from __future__ import annotations

import contextlib
import hashlib
import io
import os
import subprocess
import sys
import tarfile
import tempfile
import unittest
import urllib.request
import zipfile
from pathlib import Path
from typing import Any
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import release  # noqa: E402

GIT_ENV = dict(os.environ, GIT_CONFIG_NOSYSTEM="1", GIT_CONFIG_GLOBAL=os.devnull,
               GIT_AUTHOR_NAME="Test", GIT_AUTHOR_EMAIL="test@example.invalid",
               GIT_COMMITTER_NAME="Test", GIT_COMMITTER_EMAIL="test@example.invalid")
for _name in ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_COMMON_DIR"):
    GIT_ENV.pop(_name, None)

URL = "https://github.com/example-org/sample-project"
REPOSITORY = "example-org/sample-project"

CHANGELOG = f"""\
# Changelog

```text
## [1.2.0] - 2099-01-01
```

## [Unreleased]

- Nothing released yet.

---

## [1.2.0] - 2026-12-01

> Minor release.

### Added

- A feature.

---

## [1.1.0] - 2026-08-19

- The first release.

[Unreleased]: {URL}/compare/v1.2.0...HEAD
[1.2.0]: {URL}/compare/v1.1.0...v1.2.0
[1.1.0]: {URL}/releases/tag/v1.1.0
"""

NOTES_120 = "> Minor release.\n\n### Added\n\n- A feature.\n"

FILES = {
    "VERSION": "1.2.0\n",
    "CHANGELOG.md": CHANGELOG,
    ".gitattributes": "* text=auto eol=lf\n.gitattributes export-ignore\n.github export-ignore\n"
                      ".github/** export-ignore\nAGENTS.md export-ignore\n",
    "AGENTS.md": "Assistant instructions.\n",
    ".github/workflows/static-checks.yml": "name: Static checks\n",
    "src/M_EXAMPLE.bas": "Attribute VB_Name = \"M_EXAMPLE\"\n",
}


def git(root: Path, *args: str) -> str:
    result = subprocess.run(["git", *args], cwd=root, env=GIT_ENV, capture_output=True, check=True)
    return result.stdout.decode("utf-8").strip()


def commit(root: Path, files: dict[str, str], message: str = "change") -> str:
    for path, text in files.items():
        target = root / path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(text, encoding="utf-8", newline="\n")
    git(root, "add", "-A")
    git(root, "commit", "-q", "-m", message)
    return git(root, "rev-parse", "HEAD")


def check_runs(sha: str, *runs: tuple[str, str | None]) -> dict[str, Any]:
    """A check-runs response holding one run per (status, conclusion)."""
    return {
        "total_count": len(runs),
        "check_runs": [
            {"name": release.REQUIRED_CHECK, "head_sha": sha, "status": status,
             "conclusion": conclusion, "app": {"slug": release.CHECK_APP}}
            for status, conclusion in runs
        ],
    }


class FakeApi:
    """Answers each API path by its longest matching prefix; anything else is a 404."""

    def __init__(self, responses: dict[str, Any], files: dict[str, bytes] | None = None) -> None:
        self.responses = responses
        self.files = files or {}
        self.paths: list[str] = []

    def get_json(self, path: str) -> Any:
        self.paths.append(path)
        for prefix in sorted(self.responses, key=len, reverse=True):
            if path.startswith(prefix):
                return self.responses[prefix]
        raise release.NotFound(path)

    def download(self, url: str) -> bytes:
        return self.files[url]


class ReleaseRepository(unittest.TestCase):
    """A repository whose main holds FILES, with origin/main pointing at it."""

    def setUp(self) -> None:
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name) / "repository"
        self.root.mkdir()
        git(self.root, "init", "-q", "-b", "main")
        self.sha = commit(self.root, FILES, "release 1.2.0")
        git(self.root, "update-ref", "refs/remotes/origin/main", self.sha)

    def tag(self, name: str, target: str | None = None, annotated: bool = True) -> None:
        args = ["tag", "-a", name, "-m", f"Release {name}"] if annotated else ["tag", name]
        git(self.root, *args, *([target] if target else []))

    def passing_api(self, sha: str | None = None) -> FakeApi:
        sha = sha or self.sha
        return FakeApi({f"/repos/{REPOSITORY}/commits/{sha}/check-runs":
                        check_runs(sha, ("completed", "success"))})

    def findings(self, tag: str, api: release.Api | None = None) -> list[str]:
        found, _ = release.verify(self.root, tag, "origin/main", api or self.passing_api(), REPOSITORY)
        return found

    def cli(self, *args: str, api: release.Api | None = None) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            status = release.main(["--root", str(self.root), *args], api=api)
        return status, out.getvalue(), err.getvalue()


class Verify(ReleaseRepository):
    def test_annotated_tag_on_main_with_successful_check_passes(self) -> None:
        self.tag("v1.2.0")
        self.assertEqual(self.findings("v1.2.0"), [])

    def test_check_runs_are_read_for_the_exact_commit(self) -> None:
        self.tag("v1.2.0")
        api = self.passing_api()
        self.findings("v1.2.0", api)
        self.assertEqual(len(api.paths), 1)
        self.assertIn(f"/commits/{self.sha}/check-runs?", api.paths[0])
        self.assertIn("check_name=Repository+and+module+checks", api.paths[0])

    def test_wrong_version_is_rejected(self) -> None:
        self.tag("v1.3.0")
        found = self.findings("v1.3.0")
        self.assertEqual(len(found), 2, found)
        self.assertIn("does not match VERSION '1.2.0'", found[0])
        self.assertIn("no '## [1.3.0] - YYYY-MM-DD' section", found[1])

    def test_lightweight_tag_is_rejected(self) -> None:
        self.tag("v1.2.0", annotated=False)
        found = self.findings("v1.2.0")
        self.assertEqual(len(found), 1, found)
        self.assertIn("lightweight", found[0])

    def test_missing_changelog_section_is_rejected(self) -> None:
        undated = CHANGELOG.replace("## [1.2.0] - 2026-12-01", "## [1.2.0]")
        sha = commit(self.root, {"CHANGELOG.md": undated})
        git(self.root, "update-ref", "refs/remotes/origin/main", sha)
        self.tag("v1.2.0")
        found = self.findings("v1.2.0", self.passing_api(sha))
        self.assertEqual(len(found), 1, found)
        self.assertIn("has no '## [1.2.0] - YYYY-MM-DD' section", found[0])

    def test_fenced_heading_is_not_a_changelog_section(self) -> None:
        fenced_only = CHANGELOG.replace("## [1.2.0] - 2026-12-01\n", "## [1.2.0-unrelated]\n")
        self.assertIsNone(release.release_heading(fenced_only, "1.2.0"))

    def test_missing_version_file_is_rejected(self) -> None:
        git(self.root, "rm", "-q", "VERSION")
        git(self.root, "commit", "-q", "-m", "drop VERSION")
        sha = git(self.root, "rev-parse", "HEAD")
        git(self.root, "update-ref", "refs/remotes/origin/main", sha)
        self.tag("v1.2.0")
        found = self.findings("v1.2.0", self.passing_api(sha))
        self.assertEqual(found, [f"VERSION does not exist at {sha}"])

    def test_commit_not_on_main_is_rejected(self) -> None:
        git(self.root, "checkout", "-q", "-b", "side")
        sha = commit(self.root, {"src/M_EXAMPLE.bas": "Attribute VB_Name = \"M_EXAMPLE\"\n' side\n"})
        self.tag("v1.2.0")
        found = self.findings("v1.2.0", self.passing_api(sha))
        self.assertEqual(len(found), 1, found)
        self.assertIn("is not reachable from origin/main", found[0])

    def test_missing_main_ref_is_rejected(self) -> None:
        git(self.root, "update-ref", "-d", "refs/remotes/origin/main")
        self.tag("v1.2.0")
        found = self.findings("v1.2.0")
        self.assertEqual(found, ["origin/main does not exist; fetch it before verifying"])

    def test_missing_tag_is_rejected(self) -> None:
        self.assertEqual(self.findings("v1.2.0"), ["tag v1.2.0 does not exist"])

    def test_malformed_tag_is_a_usage_error(self) -> None:
        self.tag("V1.2.0")
        status, _, err = self.cli("verify", "V1.2.0")
        self.assertEqual(status, 2)
        self.assertIn("lower-case 'vX.Y.Z'", err)


class RequiredCheck(ReleaseRepository):
    def setUp(self) -> None:
        super().setUp()
        self.tag("v1.2.0")

    def with_runs(self, response: dict[str, Any]) -> list[str]:
        api = FakeApi({f"/repos/{REPOSITORY}/commits/{self.sha}/check-runs": response})
        return self.findings("v1.2.0", api)

    def test_failed_check_is_rejected(self) -> None:
        found = self.with_runs(check_runs(self.sha, ("completed", "failure")))
        self.assertEqual(len(found), 2, found)
        self.assertIn("concluded failure", found[0])
        self.assertIn("has not concluded success", found[1])

    def test_failure_beside_a_success_is_rejected(self) -> None:
        found = self.with_runs(check_runs(self.sha, ("completed", "success"), ("completed", "timed_out")))
        self.assertEqual(len(found), 1, found)
        self.assertIn("concluded timed_out", found[0])

    def test_missing_check_is_rejected(self) -> None:
        found = self.with_runs(check_runs(self.sha))
        self.assertEqual(found, [f"no 'Repository and module checks' check run exists on {self.sha}"])

    def test_check_still_running_is_rejected(self) -> None:
        found = self.with_runs(check_runs(self.sha, ("in_progress", None)))
        self.assertEqual(len(found), 1, found)
        self.assertIn("found: in_progress", found[0])

    def test_running_check_beside_a_success_passes(self) -> None:
        # A tag push starts the static checks again while this runs.
        self.assertEqual(self.with_runs(check_runs(self.sha, ("completed", "success"), ("queued", None))), [])

    def test_cancelled_run_beside_a_success_passes(self) -> None:
        self.assertEqual(self.with_runs(check_runs(self.sha, ("completed", "cancelled"),
                                                   ("completed", "success"))), [])

    def test_run_on_another_commit_or_from_another_app_does_not_count(self) -> None:
        response = check_runs(self.sha, ("completed", "success"), ("completed", "success"))
        response["check_runs"][0]["head_sha"] = "0" * 40
        response["check_runs"][1]["app"] = {"slug": "someone-elses-app"}
        found = self.with_runs(response)
        self.assertEqual(len(found), 1, found)
        self.assertIn("no 'Repository and module checks' check run", found[0])

    def test_partial_list_is_an_error(self) -> None:
        response = check_runs(self.sha, ("completed", "success"))
        response["total_count"] = 101
        api = FakeApi({f"/repos/{REPOSITORY}/commits/{self.sha}/check-runs": response})
        status, _, err = self.cli("verify", "v1.2.0", "--repository", REPOSITORY, api=api)
        self.assertEqual(status, 2)
        self.assertIn("partial list", err)

    def test_cli_passes_and_reports_incomplete_without_repository(self) -> None:
        status, out, _ = self.cli("verify", "v1.2.0", "--repository", REPOSITORY, api=self.passing_api())
        self.assertEqual((status, out.splitlines()[-1]), (0, "RESULT: PASS"))
        status, out, _ = self.cli("verify", "v1.2.0")
        self.assertEqual((status, out.splitlines()[-1]), (0, "RESULT: INCOMPLETE"))

    def test_cli_fails_on_a_finding(self) -> None:
        api = FakeApi({f"/repos/{REPOSITORY}/commits/{self.sha}/check-runs":
                       check_runs(self.sha, ("completed", "failure"))})
        status, out, _ = self.cli("verify", "v1.2.0", "--repository", REPOSITORY, api=api)
        self.assertEqual((status, out.splitlines()[-1]), (1, "RESULT: FAIL"))


class Notes(ReleaseRepository):
    def test_notes_print_the_section_body(self) -> None:
        self.tag("v1.2.0")
        status, out, _ = self.cli("notes", "v1.2.0")
        self.assertEqual((status, out), (0, NOTES_120))

    def test_oldest_section_drops_the_link_references(self) -> None:
        self.assertEqual(release.release_notes(CHANGELOG, "1.1.0"), "- The first release.\n")

    def test_missing_section_is_an_error(self) -> None:
        self.tag("v1.3.0")
        status, _, err = self.cli("notes", "v1.3.0")
        self.assertEqual(status, 2)
        self.assertIn("has no section for 1.3.0", err)

    def test_this_repository_changelog_yields_notes(self) -> None:
        text = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
        notes = release.release_notes(text, "1.1.2")
        self.assertIsNotNone(notes)
        assert notes is not None
        self.assertNotIn("## [", notes)
        self.assertFalse(notes.rstrip().endswith("---"))


class Bundle(ReleaseRepository):
    def setUp(self) -> None:
        super().setUp()
        self.tag("v1.2.0")
        self.out = Path(self.directory.name) / "dist"

    def test_bundle_writes_archives_and_sums_without_excluded_files(self) -> None:
        paths = release.bundle(self.root, "v1.2.0", self.out)
        self.assertEqual([path.name for path in paths],
                         ["VBA-EXCEL_UI-1.2.0.zip", "VBA-EXCEL_UI-1.2.0.tar.gz", "SHA256SUMS"])
        with zipfile.ZipFile(paths[0]) as archive:
            zipped = sorted(name for name in archive.namelist() if not name.endswith("/"))
        with tarfile.open(paths[1], "r:gz") as tarball:
            tarred = sorted(member.name for member in tarball.getmembers() if not member.isdir())
        expected = ["VBA-EXCEL_UI-1.2.0/CHANGELOG.md", "VBA-EXCEL_UI-1.2.0/VERSION",
                    "VBA-EXCEL_UI-1.2.0/src/M_EXAMPLE.bas"]
        self.assertEqual(zipped, expected)
        self.assertEqual(tarred, expected)
        sums = release.parse_sums(paths[2].read_text(encoding="utf-8"))
        for path in paths[:2]:
            self.assertEqual(sums[path.name], hashlib.sha256(path.read_bytes()).hexdigest())

    def test_archive_still_holding_an_excluded_file_is_refused(self) -> None:
        sha = commit(self.root, {".gitattributes": "* text=auto eol=lf\n"})
        self.tag("v1.2.1", sha)
        with self.assertRaisesRegex(release.ReleaseError, "AGENTS.md"):
            release.bundle(self.root, "v1.2.1", self.out)

    def test_relative_output_is_resolved_against_the_caller(self) -> None:
        caller = Path(self.directory.name) / "caller"
        caller.mkdir()
        with contextlib.chdir(caller):
            status, out, err = self.cli("bundle", "v1.2.0", "--out", "dist")
        self.assertEqual(status, 0, err)
        written = sorted(path.name for path in (caller / "dist").iterdir())
        self.assertEqual(written, ["SHA256SUMS", "VBA-EXCEL_UI-1.2.0.tar.gz", "VBA-EXCEL_UI-1.2.0.zip"])
        self.assertFalse((self.root / "dist").exists())
        self.assertIn(str((caller / "dist").resolve()), out)

    def test_non_empty_output_directory_is_refused(self) -> None:
        self.out.mkdir()
        (self.out / "stale.zip").write_bytes(b"old")
        status, _, err = self.cli("bundle", "v1.2.0", "--out", str(self.out))
        self.assertEqual(status, 2)
        self.assertIn("is not empty", err)

    def test_malformed_sums_are_rejected(self) -> None:
        for text in ("", "abc  file.zip\n", f"{'0' * 64} file.zip\n", f"{'0' * 64}  SHA256SUMS\n"):
            with self.subTest(text=text), self.assertRaises(release.ReleaseError):
                release.parse_sums(text)


class VerifyPublished(ReleaseRepository):
    def setUp(self) -> None:
        super().setUp()
        self.tag("v1.2.0")
        self.out = Path(self.directory.name) / "dist"
        self.paths = release.bundle(self.root, "v1.2.0", self.out)
        self.files = {f"https://api.invalid/assets/{path.name}": path.read_bytes() for path in self.paths}
        self.release: dict[str, Any] = {
            "tag_name": "v1.2.0",
            "draft": False,
            "body": NOTES_120.replace("\n", "\r\n"),
            "assets": [{"name": path.name, "state": "uploaded",
                        "url": f"https://api.invalid/assets/{path.name}",
                        "digest": "sha256:" + hashlib.sha256(path.read_bytes()).hexdigest()}
                       for path in self.paths],
        }

    def check(self, *extra: str) -> tuple[int, str]:
        api = FakeApi({f"/repos/{REPOSITORY}/releases/tags/v1.2.0": self.release}, self.files)
        with contextlib.redirect_stderr(io.StringIO()):
            status, out, err = self.cli("verify-published", "v1.2.0", "--repository", REPOSITORY,
                                        "--dir", str(self.out), *extra, api=api)
        return status, out + err

    def test_stable_release_marked_prerelease_is_rejected(self) -> None:
        self.release["prerelease"] = True
        status, out = self.check()
        self.assertEqual(status, 1)
        self.assertIn("is marked a pre-release, but its version is stable", out)

    def test_matching_release_passes(self) -> None:
        self.assertEqual(self.check(), (0, "Published release check for v1.2.0\nRESULT: PASS\n"))

    def test_hash_mismatch_is_rejected(self) -> None:
        self.files["https://api.invalid/assets/VBA-EXCEL_UI-1.2.0.zip"] = b"tampered"
        status, out = self.check()
        self.assertEqual(status, 1)
        self.assertIn("asset VBA-EXCEL_UI-1.2.0.zip has SHA-256", out)
        self.assertIn("GitHub reports sha256:", out)

    def test_sums_asset_must_equal_the_local_file(self) -> None:
        self.files["https://api.invalid/assets/SHA256SUMS"] = b"other\n"
        status, out = self.check()
        self.assertEqual(status, 1)
        self.assertIn("asset SHA256SUMS has SHA-256", out)

    def test_draft_release_is_rejected(self) -> None:
        self.release["draft"] = True
        status, out = self.check()
        self.assertEqual(status, 1)
        self.assertIn("still a draft", out)

    def test_different_body_is_rejected(self) -> None:
        self.release["body"] = "Hand-written notes."
        status, out = self.check()
        self.assertEqual(status, 1)
        self.assertIn("body differs", out)

    def test_missing_and_extra_assets_are_rejected(self) -> None:
        self.release["assets"] = [*self.release["assets"][1:],
                                  {"name": "extra.xlsm", "state": "uploaded", "url": "x"}]
        status, out = self.check()
        self.assertEqual(status, 1)
        self.assertIn("asset VBA-EXCEL_UI-1.2.0.zip is missing", out)
        self.assertIn("asset extra.xlsm is neither listed in SHA256SUMS", out)

    def add_workbook(self, data: bytes = b"workbook") -> str:
        """Attach a workbook asset and return the SHA-256 of what was attached."""
        url = "https://api.invalid/assets/VBA-EXCEL_UI-Demo-v1.2.0.xlsm"
        self.files[url] = data
        self.release["assets"].append({"name": "VBA-EXCEL_UI-Demo-v1.2.0.xlsm",
                                       "state": "uploaded", "url": url})
        return hashlib.sha256(data).hexdigest()

    def test_declared_matching_workbook_passes(self) -> None:
        digest = self.add_workbook()
        status, out = self.check("--asset", f"VBA-EXCEL_UI-Demo-v1.2.0.xlsm={digest.upper()}")
        self.assertEqual((status, out.splitlines()[-1]), (0, "RESULT: PASS"))

    def test_declared_workbook_with_wrong_hash_is_rejected(self) -> None:
        self.add_workbook()
        status, out = self.check("--asset", f"VBA-EXCEL_UI-Demo-v1.2.0.xlsm={'0' * 64}")
        self.assertEqual(status, 1)
        self.assertIn("--asset expects " + "0" * 64, out)

    def test_declared_workbook_missing_is_rejected(self) -> None:
        status, out = self.check("--asset", f"VBA-EXCEL_UI-Demo-v1.2.0.xlsm={'0' * 64}")
        self.assertEqual(status, 1)
        self.assertIn("asset VBA-EXCEL_UI-Demo-v1.2.0.xlsm is missing", out)

    def test_undeclared_workbook_is_rejected(self) -> None:
        self.add_workbook()
        status, out = self.check()
        self.assertEqual(status, 1)
        self.assertIn("VBA-EXCEL_UI-Demo-v1.2.0.xlsm is neither listed in SHA256SUMS nor declared", out)

    def test_malformed_or_clashing_declarations_are_usage_errors(self) -> None:
        digest = "0" * 64
        for args in (["--asset", "workbook.xlsm"], ["--asset", f"dir/workbook.xlsm={digest}"],
                     ["--asset", f"SHA256SUMS={digest}"],
                     ["--asset", f"VBA-EXCEL_UI-1.2.0.zip={digest}"],
                     ["--asset", f"a.xlsm={digest}", "--asset", f"a.xlsm={digest}"]):
            with self.subTest(args=args):
                try:
                    status, _ = self.check(*args)
                except SystemExit as stop:  # argparse rejects a malformed value
                    status = int(str(stop.code))
                self.assertEqual(status, 2)

    def test_missing_release_is_rejected(self) -> None:
        api = FakeApi({}, self.files)
        status, out, _ = self.cli("verify-published", "v1.2.0", "--repository", REPOSITORY,
                                  "--dir", str(self.out), api=api)
        self.assertEqual(status, 1)
        self.assertIn("no published GitHub release exists for v1.2.0", out)


class PreRelease(ReleaseRepository):
    def test_suffix_decides_the_classification(self) -> None:
        cases = {"1.2.0": False, "1.2.0-rc.1": True, "1.2.0+build.5": False,
                 "1.2.0-rc.1+build.5": True, "1.2.0+build-5": False}
        for version, expected in cases.items():
            with self.subTest(version=version):
                self.assertIs(release.is_prerelease(version), expected)

    def test_cli_prints_true_or_false(self) -> None:
        self.assertEqual(self.cli("is-prerelease", "v1.2.0-rc.1"), (0, "true\n", ""))
        self.assertEqual(self.cli("is-prerelease", "v1.2.0"), (0, "false\n", ""))
        self.assertEqual(self.cli("is-prerelease", "V1.2.0")[0], 2)

    def published(self, flag: bool) -> list[str]:
        sha = commit(self.root, {"VERSION": "1.3.0-rc.1\n", "CHANGELOG.md": CHANGELOG.replace(
            "## [1.2.0] - 2026-12-01", "## [1.3.0-rc.1] - 2026-12-02\n\n- Candidate.\n\n---\n\n"
            "## [1.2.0] - 2026-12-01")})
        self.tag("v1.3.0-rc.1", sha)
        out = Path(self.directory.name) / "dist"
        paths = release.bundle(self.root, "v1.3.0-rc.1", out)
        files = {f"https://api.invalid/{path.name}": path.read_bytes() for path in paths}
        published = {"tag_name": "v1.3.0-rc.1", "draft": False, "prerelease": flag,
                     "body": "- Candidate.\n",
                     "assets": [{"name": path.name, "url": f"https://api.invalid/{path.name}"}
                                for path in paths]}
        api = FakeApi({f"/repos/{REPOSITORY}/releases/tags/v1.3.0-rc.1": published}, files)
        return release.published_findings(api, REPOSITORY, "v1.3.0-rc.1",
                                          release.tag_notes(self.root, "v1.3.0-rc.1"),
                                          (out / release.SUMS).read_text(encoding="utf-8"))

    def test_prerelease_published_as_prerelease_passes(self) -> None:
        self.assertEqual(self.published(True), [])

    def test_prerelease_published_as_stable_is_rejected(self) -> None:
        self.assertEqual(self.published(False), [
            "the release for v1.3.0-rc.1 is marked a stable release, but its version is a pre-release"])


class GitHubClient(unittest.TestCase):
    def test_token_is_not_forwarded_on_a_redirect(self) -> None:
        seen: list[urllib.request.Request] = []

        class Response(io.BytesIO):
            def __enter__(self) -> Response:
                return self

        def fake_urlopen(request: urllib.request.Request, timeout: float) -> Response:
            seen.append(request)
            return Response(b"asset")

        with mock.patch.object(urllib.request, "urlopen", fake_urlopen):
            data = release.GitHubApi("secret").download("https://api.invalid/assets/1")
        self.assertEqual(data, b"asset")
        self.assertEqual(seen[0].unredirected_hdrs.get("Authorization"), "Bearer secret")
        self.assertNotIn("Authorization", seen[0].headers)
        self.assertEqual(seen[0].headers.get("Accept"), "application/octet-stream")


if __name__ == "__main__":
    unittest.main()
