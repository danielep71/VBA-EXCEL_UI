"""Tests for the generated documentation tables (tools/render_docs.py).

Each test builds a small repository in a temporary folder: a manifest and one
or two Markdown files. Only the test of the default file list needs Git.
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import render_docs  # noqa: E402

PRODUCTION = ["src/M_B.bas", "src/M_A.bas"]
MANIFEST: dict[str, Any] = {
    "module_descriptions": {
        "src/M_A.bas": "The facade",
        "src/M_B.bas": "Host operations",
        "test/M_T.bas": "Regression tests",
    },
    "configurations": {
        "production": {"modules": PRODUCTION},
        "development": {"modules": [*PRODUCTION, "test/M_T.bas"]},
    },
}
PRODUCTION_TABLE = (
    "<!-- generated:modules:production -->\n"
    "\n"
    "| Order | Repository source | VBE component | Responsibility |\n"
    "|---:|---|---|---|\n"
    "| 1 | `src/M_B.bas` | `M_B` | Host operations |\n"
    "| 2 | `src/M_A.bas` | `M_A` | The facade |\n"
    "\n"
    "<!-- /generated:modules:production -->\n"
)
DEVELOPMENT_TABLE = (
    "<!-- generated:modules:development -->\n"
    "\n"
    "| Order | Repository source | VBE component | Responsibility |\n"
    "|---:|---|---|---|\n"
    "| 3 | `test/M_T.bas` | `M_T` | Regression tests |\n"
    "\n"
    "<!-- /generated:modules:development -->\n"
)
EMPTY = ("<!-- generated:modules:production -->\n<!-- /generated:modules:production -->\n"
         "\nText between.\n\n"
         "<!-- generated:modules:development -->\nstale\n<!-- /generated:modules:development -->\n")
RENDERED = PRODUCTION_TABLE + "\nText between.\n\n" + DEVELOPMENT_TABLE


class Repo:
    """A temporary repository with a manifest and Markdown files."""

    def __init__(self, root: Path, manifest: dict[str, Any] = MANIFEST) -> None:
        self.root = root
        self.write("tools/vba-projects.json", json.dumps(manifest, indent=2) + "\n")

    def write(self, rel: str, text: str) -> None:
        path = self.root / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(text.encode("utf-8"))

    def read(self, rel: str) -> str:
        return (self.root / rel).read_bytes().decode("utf-8")

    def main(self, *arguments: str) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = render_docs.main([*arguments, "--root", str(self.root)])
        return code, out.getvalue(), err.getvalue()


class RenderTestCase(unittest.TestCase):
    def setUp(self) -> None:
        folder = tempfile.TemporaryDirectory(prefix="render-docs-")
        self.addCleanup(folder.cleanup)
        self.repo = Repo(Path(folder.name))

    def render(self, text: str, manifest: dict[str, Any] = MANIFEST) -> str:
        return render_docs.render_text("DOC.md", text, manifest)


class Render(RenderTestCase):
    def test_writes_each_table_between_its_markers(self) -> None:
        self.assertEqual(self.render(EMPTY), RENDERED)

    def test_rendering_is_idempotent(self) -> None:
        self.assertEqual(self.render(RENDERED), RENDERED)

    def test_production_lists_every_module_in_manifest_order(self) -> None:
        lines = render_docs.module_table(MANIFEST, "production")
        self.assertEqual([line.split("|")[2].strip() for line in lines[2:]], ["`src/M_B.bas`", "`src/M_A.bas`"])

    def test_a_configuration_lists_only_what_it_adds_to_one_it_contains(self) -> None:
        lines = render_docs.module_table(MANIFEST, "development")
        self.assertEqual(lines[2:], ["| 3 | `test/M_T.bas` | `M_T` | Regression tests |"])

    def test_keeps_crlf_line_endings(self) -> None:
        crlf = EMPTY.replace("\n", "\r\n")
        self.assertEqual(self.render(crlf), RENDERED.replace("\n", "\r\n"))

    def test_leaves_a_document_without_markers_alone(self) -> None:
        text = "# Title\n\nA `<!-- generated:modules:production -->` marker named in prose.\n"
        self.assertEqual(self.render(text), text)

    def test_the_current_documentation_is_up_to_date(self) -> None:
        code = subprocess.run([sys.executable, str(ROOT / "tools" / "render_docs.py"), "--check"],
                              cwd=ROOT, capture_output=True, text=True, check=False)
        self.assertEqual(code.returncode, 0, code.stdout + code.stderr)

    def test_the_installation_guide_has_a_table_for_each_configuration(self) -> None:
        text = (ROOT / "INSTALLATION.md").read_text(encoding="utf-8")
        for name in ("production", "development", "demo"):
            with self.subTest(configuration=name):
                self.assertIn(f"<!-- generated:modules:{name} -->", text)


class RenderErrors(RenderTestCase):
    def assertRenderError(self, text: str, fragment: str, manifest: dict[str, Any] = MANIFEST) -> None:
        with self.assertRaises(render_docs.RenderError) as caught:
            self.render(text, manifest)
        self.assertIn(fragment, str(caught.exception))

    def test_rejects_a_missing_closing_marker(self) -> None:
        self.assertRenderError("<!-- generated:modules:production -->\ntext\n",
                               "DOC.md:1: generated:modules:production has no closing")

    def test_rejects_a_closing_marker_without_an_opening_one(self) -> None:
        self.assertRenderError("<!-- /generated:modules:production -->\n", "DOC.md:1: /generated:modules:production")

    def test_rejects_a_closing_marker_for_another_block(self) -> None:
        self.assertRenderError(
            "<!-- generated:modules:production -->\n<!-- /generated:modules:development -->\n",
            "closes no open block")

    def test_rejects_a_block_opened_inside_another(self) -> None:
        self.assertRenderError(
            "<!-- generated:modules:production -->\n<!-- generated:modules:development -->\n",
            "DOC.md:2: generated:modules:development opens inside generated:modules:production")

    def test_rejects_a_malformed_marker(self) -> None:
        self.assertRenderError("<!-- generated: modules:production -->\n", "DOC.md:1: malformed generated marker")

    def test_rejects_an_unknown_configuration(self) -> None:
        self.assertRenderError("<!-- generated:modules:staging -->\n<!-- /generated:modules:staging -->\n",
                               "unknown configuration 'staging'")

    def test_rejects_an_unknown_kind(self) -> None:
        self.assertRenderError("<!-- generated:tables:production -->\n<!-- /generated:tables:production -->\n",
                               "unknown generated block kind 'tables'")

    def test_rejects_a_module_without_a_description(self) -> None:
        manifest = json.loads(json.dumps(MANIFEST))
        del manifest["module_descriptions"]["src/M_A.bas"]
        self.assertRenderError(EMPTY, "no description for src/M_A.bas", manifest)

    def test_rejects_a_description_that_would_break_the_table(self) -> None:
        manifest = json.loads(json.dumps(MANIFEST))
        manifest["module_descriptions"]["src/M_A.bas"] = "The | facade"
        self.assertRenderError(EMPTY, "must be one line without '|'", manifest)


class CommandLine(RenderTestCase):
    def test_check_passes_on_current_tables(self) -> None:
        self.repo.write("DOC.md", RENDERED)
        code, out, _ = self.repo.main("--check", "DOC.md")
        self.assertEqual(code, 0)
        self.assertIn("render-docs: PASS", out)

    def test_check_fails_with_a_diff_and_changes_nothing(self) -> None:
        self.repo.write("DOC.md", EMPTY)
        code, out, _ = self.repo.main("--check", "DOC.md")
        self.assertEqual(code, 1)
        self.assertIn("--- a/DOC.md", out)
        self.assertIn("+++ b/DOC.md", out)
        self.assertIn("-stale", out)
        self.assertIn("+| 3 | `test/M_T.bas` | `M_T` | Regression tests |", out)
        self.assertIn("run python3 tools/render_docs.py --write", out)
        self.assertEqual(self.repo.read("DOC.md"), EMPTY)

    def test_write_rewrites_the_tables(self) -> None:
        self.repo.write("DOC.md", EMPTY)
        code, out, _ = self.repo.main("--write", "DOC.md")
        self.assertEqual(code, 0)
        self.assertIn("rewrote DOC.md", out)
        self.assertEqual(self.repo.read("DOC.md"), RENDERED)
        self.assertEqual(self.repo.main("--check", "DOC.md")[0], 0)

    def test_an_error_exits_2_and_writes_nothing(self) -> None:
        broken = EMPTY + "<!-- generated:modules:staging -->\n<!-- /generated:modules:staging -->\n"
        self.repo.write("DOC.md", broken)
        code, _, err = self.repo.main("--write", "DOC.md")
        self.assertEqual(code, 2)
        self.assertIn("ERROR: unknown configuration 'staging'", err)
        self.assertEqual(self.repo.read("DOC.md"), broken)

    def test_an_unreadable_manifest_exits_2(self) -> None:
        self.repo.write("tools/vba-projects.json", "{ not json\n")
        self.repo.write("DOC.md", RENDERED)
        code, _, err = self.repo.main("--check", "DOC.md")
        self.assertEqual(code, 2)
        self.assertIn("tools/vba-projects.json cannot be read", err)

    def test_check_and_write_are_required_and_exclusive(self) -> None:
        for arguments in ((), ("--check", "--write")):
            with self.subTest(arguments=arguments), contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit) as caught:
                    render_docs.main(list(arguments))
                self.assertEqual(caught.exception.code, 2)

    def test_reads_the_tracked_markdown_files_by_default(self) -> None:
        self.repo.write("DOC.md", EMPTY)
        self.repo.write("docs/UNTRACKED.md", EMPTY)
        self.git("init", "-q")
        self.git("add", "DOC.md", "tools/vba-projects.json")
        code, out, _ = self.repo.main("--write")
        self.assertEqual(code, 0)
        self.assertEqual(self.repo.read("DOC.md"), RENDERED)
        self.assertEqual(self.repo.read("docs/UNTRACKED.md"), EMPTY)
        self.assertNotIn("UNTRACKED", out)

    def git(self, *arguments: str) -> None:
        env = dict(os.environ, GIT_CONFIG_NOSYSTEM="1", GIT_CONFIG_GLOBAL=os.devnull)
        for name in ("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_COMMON_DIR"):
            env.pop(name, None)
        subprocess.run(["git", *arguments], cwd=self.repo.root, env=env,
                       capture_output=True, check=True)


if __name__ == "__main__":
    unittest.main()
