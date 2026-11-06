"""Temporary repositories for testing the static gate.

Each test gets its own copy of this repository's tracked files in a fresh Git
repository, mutates it, points tools/check_repo.py at it, and runs one check.
A copy of the real tree is the passing fixture; one deliberate defect is the
failing fixture. Building fixtures from the real tree means a test cannot pass
merely because its fixture is unrealistically small.

Only the standard library is used, like the gate itself. The issue-form tests
in test_check_issues.py also need Node.js.
"""

from __future__ import annotations

import contextlib
import io
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections.abc import Callable, Iterator
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
TOOLS = REPO / "tools"

if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))

# Imported here, after TOOLS is on sys.path, and re-exported for the tests.
import check_repo as check_repo  # noqa: E402
import reformat as reformat  # noqa: E402
import vba_api as vba_api  # noqa: E402
import wiki_badges as wiki_badges  # noqa: E402

# Fixture repositories must not depend on whoever runs the tests: a personal
# excludesFile or attributes file would change what git check-ignore and
# git check-attr report, and therefore what the gate decides.
GIT_ENV = dict(os.environ, GIT_CONFIG_NOSYSTEM="1", GIT_CONFIG_GLOBAL=os.devnull)
for _name in ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_COMMON_DIR"):
    GIT_ENV.pop(_name, None)

# The gate shells out to git with the ambient environment. The same isolation
# is applied to this test process so that the gate sees what the fixture says.
os.environ.update(GIT_CONFIG_NOSYSTEM="1", GIT_CONFIG_GLOBAL=os.devnull)
for _name in ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_COMMON_DIR"):
    os.environ.pop(_name, None)

# Large binaries are hard-linked rather than copied. Nothing writes to them in
# place: Fixture.write() and Fixture.remove() unlink a path before replacing
# it, so a test can never modify the template or the real checkout.
HARD_LINKED_PREFIXES = ("images/", ".git/objects/")


def git(root: Path, *args: str, data: bytes | None = None) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(["git", *args], cwd=root, input=data, capture_output=True,
                          env=GIT_ENV, check=True)


def tracked_files(root: Path) -> list[str]:
    out = git(root, "ls-files", "-z").stdout
    return sorted(os.fsdecode(p) for p in out.split(b"\0") if p)


_TEMPLATE: Path | None = None
_TEMPLATE_DIR: tempfile.TemporaryDirectory[str] | None = None


def template() -> Path:
    """Return a pristine fixture repository, built once per test run."""
    global _TEMPLATE, _TEMPLATE_DIR
    if _TEMPLATE is None:
        _TEMPLATE_DIR = tempfile.TemporaryDirectory(prefix="excel-ui-gate-template-")
        root = Path(_TEMPLATE_DIR.name) / "repo"
        root.mkdir()
        for rel in tracked_files(REPO):
            source = REPO / rel
            if not source.is_file():
                continue          # deleted in the working tree but still staged
            target = root / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, target)
        git(root, "init", "-q")
        git(root, "add", "-A")
        _TEMPLATE = root
    return _TEMPLATE


def _copy(source: str, target: str) -> None:
    rel = Path(target).as_posix()
    if any(f"/{prefix}" in rel for prefix in HARD_LINKED_PREFIXES):
        try:
            os.link(source, target)
            return
        except OSError:
            pass
    shutil.copy2(source, target)


class Fixture:
    """One disposable repository and the helpers to mutate it."""

    def __init__(self, root: Path) -> None:
        self.root = root

    def path(self, rel: str) -> Path:
        return self.root / rel

    def read_bytes(self, rel: str) -> bytes:
        return self.path(rel).read_bytes()

    def read_text(self, rel: str) -> str:
        return self.read_bytes(rel).decode("utf-8")

    def write(self, rel: str, data: bytes | str, *, track: bool = True) -> None:
        target = self.path(rel)
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists() or target.is_symlink():
            target.unlink()           # never write through a hard link
        target.write_bytes(data.encode("utf-8") if isinstance(data, str) else data)
        if track:
            git(self.root, "add", "-f", "--", rel)

    def replace(self, rel: str, old: str | bytes, new: str | bytes, count: int = 1) -> None:
        """Replace text in a file, failing loudly when the text is absent."""
        data = self.read_bytes(rel)
        old_b = old.encode("utf-8") if isinstance(old, str) else old
        new_b = new.encode("utf-8") if isinstance(new, str) else new
        if old_b not in data:
            raise AssertionError(f"{rel} does not contain {old_b!r}")
        self.write(rel, data.replace(old_b, new_b, count))

    def append(self, rel: str, data: str | bytes) -> None:
        extra = data.encode("utf-8") if isinstance(data, str) else data
        self.write(rel, self.read_bytes(rel) + extra)

    def remove(self, rel: str) -> None:
        git(self.root, "rm", "-q", "--cached", "--", rel)
        self.path(rel).unlink()

    def run(self, check: Callable[[], None]) -> list[str]:
        """Run one gate check against this repository and return its findings."""
        saved_path = list(sys.path)
        check_repo.REPO = str(self.root)
        del check_repo.failures[:]
        del check_repo.notes[:]
        try:
            with contextlib.redirect_stdout(io.StringIO()):
                check_repo.load_modules()
                check()
            return list(check_repo.failures)
        finally:
            check_repo.REPO = str(REPO)
            sys.path[:] = saved_path


@contextlib.contextmanager
def fixture() -> Iterator[Fixture]:
    with tempfile.TemporaryDirectory(prefix="excel-ui-gate-") as name:
        root = Path(name) / "repo"
        shutil.copytree(template(), root, symlinks=True, copy_function=_copy)
        yield Fixture(root)


class GateTestCase(unittest.TestCase):
    """A test of one CHECKS entry.

    Set ``check`` to the CHECKS label. Each subclass must provide at least one
    ``test_passes_*`` method and one ``test_fails_*`` method; the coverage
    test enforces both for every entry in CHECKS.
    """

    check: str = ""

    def setUp(self) -> None:
        self._context = fixture()
        self.repo = self._context.__enter__()
        self.addCleanup(self._context.__exit__, None, None, None)

    def check_function(self) -> Callable[[], None]:
        for label, fn in check_repo.CHECKS:
            if label == self.check:
                return fn
        raise AssertionError(f"no CHECKS entry is labelled {self.check!r}")

    def findings(self) -> list[str]:
        return self.repo.run(self.check_function())

    def assertPasses(self) -> None:
        found = self.findings()
        self.assertEqual(found, [], f"{self.check} reported findings on a clean fixture")

    def assertFails(self, *fragments: str) -> list[str]:
        """Assert the check reports a finding containing every fragment."""
        found = self.findings()
        self.assertTrue(found, f"{self.check} reported nothing for a defective fixture")
        for fragment in fragments:
            self.assertTrue(any(fragment in f for f in found),
                            f"no finding mentions {fragment!r}: {found}")
        return found
