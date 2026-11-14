#!/usr/bin/env python3
"""Write the generated tables of the documentation from tools/vba-projects.json.

A generated table sits between two marker lines:

    <!-- generated:modules:<configuration> -->
    <!-- /generated:modules:<configuration> -->

and is rewritten from the manifest, never edited by hand. The table lists the
modules of that configuration with their import order, the VBE component each
one becomes, and the description that module_descriptions in the manifest gives
it. A configuration that holds every module of another configuration lists only
the modules it adds to it: the development table shows the regression module,
not the production package again. The Order column is the module's position in
the configuration's list, which is the order in which the workbook tools import
it.

--check compares every block with what the manifest gives and prints a diff for
each one that differs; --write rewrites them. Without file arguments, both read
the Markdown files Git tracks.

Exit status: 0 when every block is current (or was rewritten), 1 when --check
finds a block that differs, 2 for an error in the arguments, a marker or the
manifest.
"""

from __future__ import annotations

import argparse
import difflib
import json
import re
import subprocess
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = "tools/vba-projects.json"
KINDS = ("modules",)
MARKER = re.compile(r"^<!-- (?P<end>/)?generated:(?P<kind>[\w-]+):(?P<name>[\w-]+) -->$")
ANY_MARKER = re.compile(r"^\s*<!--\s*/?generated:")
HEADER = ("| Order | Repository source | VBE component | Responsibility |",
          "|---:|---|---|---|")


class RenderError(Exception):
    """A marker, a manifest or an argument that cannot be rendered."""


@dataclass(frozen=True)
class Block:
    kind: str
    name: str
    start: int  # index of the opening marker line
    end: int  # index of the closing marker line


def load_manifest(root: Path) -> dict[str, Any]:
    path = root / MANIFEST
    try:
        manifest = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RenderError(f"{MANIFEST} cannot be read: {error}") from error
    if not isinstance(manifest, dict):
        raise RenderError(f"{MANIFEST} is not a JSON object")
    return manifest


def configurations(manifest: dict[str, Any]) -> dict[str, list[str]]:
    """Return {configuration: [module paths]} from the manifest."""
    found = manifest.get("configurations")
    if not isinstance(found, dict):
        raise RenderError(f"{MANIFEST} has no configurations")
    result: dict[str, list[str]] = {}
    for name, entry in found.items():
        modules = entry.get("modules") if isinstance(entry, dict) else None
        if not isinstance(modules, list) or not all(isinstance(rel, str) for rel in modules):
            raise RenderError(f"{MANIFEST}: configuration {name!r} has no list of module paths")
        result[name] = modules
    return result


def description(manifest: dict[str, Any], rel: str) -> str:
    descriptions = manifest.get("module_descriptions")
    text = descriptions.get(rel) if isinstance(descriptions, dict) else None
    if not isinstance(text, str) or not text.strip():
        raise RenderError(f"{MANIFEST}: module_descriptions has no description for {rel}")
    if "|" in text or "\n" in text:
        raise RenderError(f"{MANIFEST}: the description of {rel} must be one line without '|'")
    return text.strip()


def module_table(manifest: dict[str, Any], name: str) -> list[str]:
    """Return the table lines of one configuration."""
    projects = configurations(manifest)
    if name not in projects:
        raise RenderError(f"unknown configuration {name!r}; {MANIFEST} has {', '.join(projects)}")
    modules = projects[name]
    contained: set[str] = set()
    for other, others in projects.items():
        if other != name and set(others) < set(modules):
            contained.update(others)
    lines = list(HEADER)
    for order, rel in enumerate(modules, 1):
        if rel in contained:
            continue
        component = Path(rel).stem
        lines.append(f"| {order} | `{rel}` | `{component}` | {description(manifest, rel)} |")
    return lines


def render_block(manifest: dict[str, Any], kind: str, name: str) -> list[str]:
    """Return the lines between the two markers: a blank line, the table, a blank line."""
    if kind not in KINDS:
        raise RenderError(f"unknown generated block kind {kind!r}; known: {', '.join(KINDS)}")
    return ["", *module_table(manifest, name), ""]


def find_blocks(rel: str, lines: Sequence[str]) -> list[Block]:
    """Return every marked block of a document, or raise on a malformed marker."""
    blocks: list[Block] = []
    opened: tuple[str, str, int] | None = None
    for index, line in enumerate(lines):
        m = MARKER.match(line.rstrip("\r"))
        if m is None:
            if ANY_MARKER.search(line):
                raise RenderError(f"{rel}:{index + 1}: malformed generated marker: {line.strip()}")
            continue
        kind, name = m.group("kind"), m.group("name")
        if not m.group("end"):
            if opened is not None:
                raise RenderError(f"{rel}:{index + 1}: generated:{kind}:{name} opens inside "
                                  f"generated:{opened[0]}:{opened[1]}, which line {opened[2] + 1} opened")
            opened = (kind, name, index)
        elif opened is None or (opened[0], opened[1]) != (kind, name):
            raise RenderError(f"{rel}:{index + 1}: /generated:{kind}:{name} closes no open block")
        else:
            blocks.append(Block(kind, name, opened[2], index))
            opened = None
    if opened is not None:
        raise RenderError(f"{rel}:{opened[2] + 1}: generated:{opened[0]}:{opened[1]} has no "
                          f"closing <!-- /generated:{opened[0]}:{opened[1]} --> marker")
    return blocks


def render_text(rel: str, text: str, manifest: dict[str, Any]) -> str:
    """Return the document with every generated block rewritten."""
    newline = "\r\n" if "\r\n" in text else "\n"
    lines = text.split(newline)
    for block in reversed(find_blocks(rel, lines)):
        lines[block.start + 1:block.end] = render_block(manifest, block.kind, block.name)
    return newline.join(lines)


def tracked_markdown(root: Path) -> list[str]:
    try:
        out = subprocess.run(["git", "ls-files", "-z", "--", "*.md"], cwd=root,
                             capture_output=True, check=True).stdout
    except (OSError, subprocess.CalledProcessError) as error:
        raise RenderError(f"cannot list the Markdown files Git tracks: {error}") from error
    return sorted(p.decode("utf-8") for p in out.split(b"\0") if p)


def run(root: Path, files: Sequence[str], write: bool) -> int:
    manifest = load_manifest(root)
    drift = 0
    for rel in files or tracked_markdown(root):
        path = root / rel
        try:
            text = path.read_bytes().decode("utf-8")
        except (OSError, UnicodeDecodeError) as error:
            raise RenderError(f"{rel} cannot be read: {error}") from error
        rendered = render_text(rel, text, manifest)
        if rendered == text:
            continue
        if write:
            path.write_bytes(rendered.encode("utf-8"))
            print(f"rewrote {rel}")
            continue
        drift += 1
        sys.stdout.writelines(difflib.unified_diff(
            text.splitlines(keepends=True), rendered.splitlines(keepends=True),
            fromfile=f"a/{rel}", tofile=f"b/{rel}"))
    if drift:
        print(f"render-docs: FAIL: {drift} file(s) differ from {MANIFEST}; "
              f"run python3 tools/render_docs.py --write")
        return 1
    print("render-docs: PASS")
    return 0


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true", help="fail, with a diff, when a generated block differs")
    mode.add_argument("--write", action="store_true", help="rewrite every generated block")
    parser.add_argument("--root", type=Path, default=ROOT, help="repository root (default: this checkout)")
    parser.add_argument("files", nargs="*", help="Markdown files relative to the root (default: every tracked one)")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    options = parse_args(argv)
    try:
        return run(options.root.resolve(), options.files, options.write)
    except RenderError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
