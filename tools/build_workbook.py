#!/usr/bin/env python3
"""Build a macro-enabled workbook (.xlsm) from the VBA source of this repository.

Usage: python3 tools/build_workbook.py --mode Production|Development|Demo [--override DIR] OUTPUT

The workbook holds the modules that tools/vba-projects.json lists for the
configuration that --mode names, in the manifest's order, as one VBA project.
Its sheets are those of a new Excel workbook: one worksheet, Sheet1, with the
code name Sheet1, and the workbook's code name ThisWorkbook. Both document
modules hold only Option Explicit. The package and its vbaProject.bin are
written directly, with tools/vba_project.py, so the build needs neither Excel
nor "Trust access to the VBA project object model". Excel compiles the project
from source when it opens the workbook.

The rules on the source are those of tools/Sync-EXCEL_UI-VBA.ps1: every
module the configuration lists must be a .bas file that declares one valid
Attribute VB_Name; every .bas file in src/, test/ and demo/ must be listed in
at least one configuration; no two repository modules may share a VB_Name.

--override DIR replaces modules with the files of DIR, for a negative control
or a review. A file is identified only by the Attribute VB_Name it declares,
never by its file name, as with the synchronizer's -SourceDir. A file whose
VB_Name is a module of the configuration replaces that module in its place; a
standard module (.bas) whose name matches a project_name_patterns entry that
the repository's own module names match, but which is not a repository module,
is added after the configuration's modules, in name order. When several files
declare the same VB_Name, the newest is taken. Where the synchronizer would
ignore a file, the build stops instead: a file without exactly one valid
VB_Name, of another kind than the repository module of that name, naming a
repository module outside the configuration, or naming no module at all.

The output is the same, byte for byte, for the same mode and module source:
line endings are normalized to CRLF, every part of the package has a fixed
time and permissions and is stored without ZIP compression, and the project's
ID and the random values of its PROJECT stream are derived from the source.

Exit status: 0 when the workbook is written, 1 when the source or the
override folder cannot be built, 2 for an error in the arguments.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import re
import sys
import uuid
import zipfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.sax.saxutils import quoteattr

sys.path.insert(0, str(Path(__file__).resolve().parent))
from vba_project import Module, build_vba_project, document_module  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = "tools/vba-projects.json"
MODES = ("Production", "Development", "Demo")
SOURCE_DIRECTORIES = ("src", "test", "demo")
OVERRIDE_SUFFIXES = (".bas", ".cls", ".frm")
WORKBOOK_CODE_NAME = "ThisWorkbook"
# The worksheets of a new Excel workbook: (tab name, code name).
SHEETS = (("Sheet1", "Sheet1"),)

# The declaration a module's identity comes from, matched without regard to
# letter case as PowerShell's -match does in the synchronizer.
VB_NAME_LINE = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]*)"\s*$', re.IGNORECASE)
VB_NAME_START = re.compile(r"^\s*Attribute\s+VB_Name\b", re.IGNORECASE)
VBA_NAME = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,30}$")
LINE_BREAK = re.compile(rb"\r\n|\r|\n")

ZIP_TIME = (1980, 1, 1, 0, 0, 0)
PROJECT_NAMESPACE = uuid.UUID("5d0f8a64-3b8c-5c4e-9a51-6c3f0e2b7d19")

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
VBA_RELATIONSHIP = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
MACRO_WORKBOOK = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
VBA_CONTENT = "application/vnd.ms-office.vbaProject"
WORKSHEET_CONTENT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
STYLES_CONTENT = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
RELATIONSHIPS_CONTENT = "application/vnd.openxmlformats-package.relationships+xml"
XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
VBA_PART = "xl/vbaProject.bin"


class BuildError(Exception):
    """The configuration, the source or the override folder cannot be built."""


@dataclass(frozen=True)
class Component:
    """One standard module of the build: its VB_Name, the file it comes from,
    and whether that file replaced or added it through --override."""

    name: str
    path: Path
    origin: str = "repository"  # "repository", "override" or "new"


@dataclass(frozen=True)
class Overrides:
    """The files --override takes, by lower-case VB_Name, and what was set aside."""

    picked: dict[str, Path]
    new: frozenset[str]
    older: tuple[str, ...]


@dataclass(frozen=True)
class Build:
    mode: str
    components: tuple[Component, ...]
    project_id: str
    sha256: str
    ignored: tuple[str, ...] = ()


# ---------------------------------------------------------------------------
# Source


def declared_name(data: bytes) -> tuple[str, str]:
    """The VB_Name a file declares and, when it has no usable one, why not.

    Exactly one Attribute VB_Name line holding a valid VBA name is required; the
    file name is never used instead."""
    names: list[str] = []
    for line in LINE_BREAK.split(data):
        text = line.decode("latin-1")
        match = VB_NAME_LINE.match(text)
        if match:
            names.append(match.group(1))
        elif VB_NAME_START.match(text):
            return "", "malformed Attribute VB_Name line"
    if not names:
        return "", "no Attribute VB_Name line; the file name is not used as its identity"
    if len(names) > 1:
        return "", "several Attribute VB_Name lines"
    if not VBA_NAME.match(names[0]):
        return "", f"VB_Name '{names[0]}' is not a valid VBA name"
    return names[0], ""


def module_source(data: bytes) -> bytes:
    """A module's bytes as the VBE stores them: CRLF line endings, ending in one.
    The bytes are otherwise kept, in the project's code page."""
    text = b"\r\n".join(LINE_BREAK.split(data))
    return text if text.endswith(b"\r\n") else text + b"\r\n"


def mode_name(value: str) -> str:
    """A --mode value in the synchronizer's spelling; any letter case is accepted."""
    for mode in MODES:
        if mode.casefold() == value.casefold():
            return mode
    raise argparse.ArgumentTypeError(f"unknown mode {value!r}; choose from {', '.join(MODES)}")


def load_manifest(root: Path) -> dict[str, Any]:
    """tools/vba-projects.json. Like the synchronizer, the build writes standard
    modules only, so a manifest that lists document modules is refused."""
    try:
        manifest = json.loads((root / MANIFEST).read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise BuildError(f"{MANIFEST}: cannot be read ({error})") from error
    if not isinstance(manifest, dict) or not isinstance(manifest.get("configurations"), dict):
        raise BuildError(f"{MANIFEST}: no configurations")
    if manifest.get("document_modules"):
        raise BuildError(f"{MANIFEST} lists document modules; this build writes standard modules only")
    return manifest


def configuration_files(root: Path, manifest: dict[str, Any], mode: str) -> list[Path]:
    """The modules of one configuration, in the manifest's order. Each must
    exist and be a standard module."""
    configuration = manifest["configurations"].get(mode.lower())
    if not isinstance(configuration, dict) or not isinstance(configuration.get("modules"), list):
        raise BuildError(f"configuration '{mode.lower()}' not found in {MANIFEST}")
    files: list[Path] = []
    for entry in configuration["modules"]:
        relative = str(entry)
        if Path(relative).suffix.lower() != ".bas":
            raise BuildError(f"module listed in {MANIFEST} is not a standard module (.bas): {relative}")
        if not (root / relative).is_file():
            raise BuildError(f"module listed in {MANIFEST} not found: {relative}")
        files.append(root / relative)
    return files


def repository_sources(root: Path) -> list[Path]:
    """Every .bas file under src/, test/ and demo/, sorted by path."""
    files: list[Path] = []
    for directory in SOURCE_DIRECTORIES:
        if not (root / directory).is_dir():
            raise BuildError(f"required source directory not found: {directory}")
        files += [path for path in (root / directory).rglob("*") if path.is_file() and path.suffix.lower() == ".bas"]
    return sorted(files)


def repository_names(files: Sequence[Path]) -> dict[str, Path]:
    """The repository modules by lower-case VB_Name; a VB_Name used twice is refused."""
    names: dict[str, Path] = {}
    for path in files:
        name, problem = declared_name(path.read_bytes())
        if problem:
            raise BuildError(f"repository module {path}: {problem}")
        if name.lower() in names:
            raise BuildError(f"duplicate VBA module name '{name}' in the repository source")
        names[name.lower()] = path
    return names


def unlisted_sources(root: Path, manifest: dict[str, Any], files: Sequence[Path]) -> list[Path]:
    """The repository modules that no configuration of the manifest lists."""
    listed = {path.resolve() for mode in MODES for path in configuration_files(root, manifest, mode)}
    return [path for path in files if path.resolve() not in listed]


def new_module_patterns(patterns: Sequence[str], names: Sequence[str]) -> list[str]:
    """The project_name_patterns that name modules: those that at least one
    repository module name matches. The others name procedures."""
    return [pattern for pattern in patterns if any(re.search(pattern, name) for name in names)]


# ---------------------------------------------------------------------------
# Override folder


def select_overrides(folder: Path, names: dict[str, Path], patterns: Sequence[str]) -> Overrides:
    """The files of ``folder`` to build with, by lower-case VB_Name. A file the
    synchronizer's -SourceDir would ignore stops the build, except an older copy
    of a module, which is set aside as the synchronizer does."""
    picked: dict[str, Path] = {}
    new: set[str] = set()
    older: list[str] = []
    candidates = sorted(path for path in folder.iterdir()
                        if path.is_file() and path.suffix.lower() in OVERRIDE_SUFFIXES)
    for path in candidates:
        name, problem = declared_name(path.read_bytes())
        if problem:
            raise BuildError(f"override {path.name}: {problem}")
        key = name.lower()
        if key in names:
            if path.suffix.lower() != names[key].suffix.lower():
                raise BuildError(f"override {path.name}: '{name}' is a {names[key].suffix.lower()} module "
                                 "in the repository")
        elif path.suffix.lower() == ".bas" and any(re.search(pattern, name) for pattern in patterns):
            new.add(key)
        else:
            raise BuildError(f"override {path.name}: VB_Name '{name}' is not a repository module, and is not "
                             f"a new standard module named like one ({', '.join(patterns) or 'no pattern'})")
        if key not in picked:
            picked[key] = path
        elif path.stat().st_mtime > picked[key].stat().st_mtime:
            older.append(f"{picked[key].name} (older copy of '{name}')")
            picked[key] = path
        else:
            older.append(f"{path.name} (older copy of '{name}')")
    if not picked:
        raise BuildError(f"override folder {folder} holds no module (.bas, .cls or .frm)")
    return Overrides(picked, frozenset(new), tuple(older))


def components(root: Path, mode: str, override: Path | None = None) -> tuple[list[Component], Overrides | None]:
    """The standard modules of the build, in project order."""
    manifest = load_manifest(root)
    names = repository_names(repository_sources(root))
    unlisted = unlisted_sources(root, manifest, list(names.values()))
    if unlisted:
        raise BuildError(f"module not listed in any configuration of {MANIFEST}: "
                         + ", ".join(path.relative_to(root).as_posix() for path in unlisted))
    listed: list[Component] = []
    for path in configuration_files(root, manifest, mode):
        name, problem = declared_name(path.read_bytes())
        if problem:
            raise BuildError(f"{path.relative_to(root).as_posix()}: {problem}")
        listed.append(Component(name, path))
    if override is None:
        return listed, None
    if not override.is_dir():
        raise BuildError(f"override is not a folder: {override}")
    patterns = new_module_patterns([str(item) for item in manifest.get("project_name_patterns", [])],
                                   [declared_name(path.read_bytes())[0] for path in names.values()])
    overrides = select_overrides(override, names, patterns)
    in_mode = {component.name.lower() for component in listed}
    outside = sorted(overrides.picked[key].name for key in overrides.picked
                     if key in names and key not in in_mode)
    if outside:
        raise BuildError(f"override {', '.join(outside)}: not a module of the {mode} configuration; modules "
                         "of different configurations are never combined")
    built = [Component(component.name, overrides.picked[component.name.lower()], "override")
             if component.name.lower() in overrides.picked else component for component in listed]
    built += [Component(declared_name(overrides.picked[key].read_bytes())[0], overrides.picked[key], "new")
              for key in sorted(overrides.new)]
    return built, overrides


# ---------------------------------------------------------------------------
# Package


def vba_modules(parts: Sequence[Component]) -> list[Module]:
    """The document modules of the workbook and its sheets, then the standard modules."""
    modules = [document_module(WORKBOOK_CODE_NAME, workbook=True)]
    modules += [document_module(code_name, workbook=False) for _, code_name in SHEETS]
    modules += [Module(part.name, module_source(part.path.read_bytes())) for part in parts]
    return modules


def project_id(mode: str, modules: Sequence[Module]) -> str:
    """A project ID derived from the mode and every module's name and source."""
    digest = hashlib.sha256(mode.encode("ascii"))
    for module in modules:
        digest.update(b"\0" + module.name.encode("ascii") + b"\0" + module.source)
    return "{" + str(uuid.uuid5(PROJECT_NAMESPACE, digest.hexdigest())).upper() + "}"


def workbook_xml() -> str:
    sheets = "".join(f'<sheet name={quoteattr(tab)} sheetId="{index}" r:id="rId{index}"/>'
                     for index, (tab, _) in enumerate(SHEETS, 1))
    return (f'{XML_HEADER}<workbook xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
            f"<workbookPr codeName={quoteattr(WORKBOOK_CODE_NAME)}/>"
            f"<bookViews><workbookView/></bookViews><sheets>{sheets}</sheets></workbook>")


def worksheet_xml(code_name: str) -> str:
    return (f'{XML_HEADER}<worksheet xmlns="{MAIN}" xmlns:r="{RELATIONSHIPS}">'
            f"<sheetPr codeName={quoteattr(code_name)}/><sheetData/></worksheet>")


STYLES = (f'{XML_HEADER}<styleSheet xmlns="{MAIN}">'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill>'
          '<fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')


def package_parts(vba_bin: bytes) -> dict[str, bytes]:
    """Every part of the .xlsm package, in the order they are written."""
    count = len(SHEETS)
    worksheets = "".join(f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="{WORKSHEET_CONTENT}"/>'
                         for index in range(1, count + 1))
    sheet_relationships = "".join(
        f'<Relationship Id="rId{index}" Type="{RELATIONSHIPS}/worksheet" Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, count + 1))
    parts = {
        "[Content_Types].xml": (
            f'{XML_HEADER}<Types xmlns="{CONTENT_TYPES}">'
            f'<Default Extension="rels" ContentType="{RELATIONSHIPS_CONTENT}"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Default Extension="bin" ContentType="{VBA_CONTENT}"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{MACRO_WORKBOOK}"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{STYLES_CONTENT}"/>'
            f"{worksheets}</Types>"),
        "_rels/.rels": (
            f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">'
            f'<Relationship Id="rId1" Type="{RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"),
        "xl/workbook.xml": workbook_xml(),
        "xl/_rels/workbook.xml.rels": (
            f'{XML_HEADER}<Relationships xmlns="{PACKAGE_RELATIONSHIPS}">{sheet_relationships}'
            f'<Relationship Id="rIdStyles" Type="{RELATIONSHIPS}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rIdVBA" Type="{VBA_RELATIONSHIP}" Target="vbaProject.bin"/>'
            "</Relationships>"),
        "xl/styles.xml": STYLES,
    }
    parts.update((f"xl/worksheets/sheet{index}.xml", worksheet_xml(code_name))
                 for index, (_, code_name) in enumerate(SHEETS, 1))
    blobs = {name: text.encode("utf-8") for name, text in parts.items()}
    blobs[VBA_PART] = vba_bin
    return blobs


def zip_package(parts: dict[str, bytes]) -> bytes:
    """The ZIP of the parts, in the given order. Each part has a fixed time and
    permissions and is stored uncompressed, so the bytes do not depend on the
    platform or on the zlib build."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=ZIP_TIME)
            info.compress_type = zipfile.ZIP_STORED
            # ZipInfo takes the creating system from the platform; fix it to
            # Unix, the system the permission bits belong to.
            info.create_system = 3
            info.external_attr = 0o644 << 16
            archive.writestr(info, data)
    return buffer.getvalue()


def workbook_bytes(mode: str, parts: Sequence[Component]) -> tuple[bytes, str]:
    """The .xlsm of the build and its project ID."""
    modules = vba_modules(parts)
    identifier = project_id(mode, modules)
    try:
        vba_bin = build_vba_project(modules, identifier)
    except ValueError as error:
        raise BuildError(str(error)) from error
    return zip_package(package_parts(vba_bin)), identifier


def build(root: Path, mode: str, output: Path, override: Path | None = None) -> Build:
    """Write the workbook of ``mode`` to ``output`` and describe it."""
    if mode not in MODES:
        raise BuildError(f"unknown mode {mode!r}; choose from {', '.join(MODES)}")
    parts, overrides = components(root, mode, override)
    data, identifier = workbook_bytes(mode, parts)
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(output.name + ".partial")
    temporary.write_bytes(data)
    os.replace(temporary, output)
    return Build(mode, tuple(parts), identifier, hashlib.sha256(data).hexdigest(),
                 overrides.older if overrides is not None else ())


def parse_arguments(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a macro-enabled workbook from the VBA source of this repository, without Excel.",
        epilog="Exit status: 0 written, 1 the source or the override folder cannot be built, 2 usage error.")
    parser.add_argument("--mode", required=True, type=mode_name, metavar="|".join(MODES),
                        help=f"the configuration of {MANIFEST} to build")
    parser.add_argument("--override", type=Path, metavar="DIR",
                        help="replace modules with the files of DIR, identified by their Attribute VB_Name")
    parser.add_argument("output", type=Path, metavar="OUTPUT", help="the .xlsm file to write; it is replaced")
    arguments = parser.parse_args(argv)
    if arguments.output.suffix.lower() != ".xlsm":
        parser.error("OUTPUT must be an .xlsm file")
    return arguments


def main(argv: Sequence[str] | None = None) -> int:
    arguments = parse_arguments(argv)
    try:
        result = build(ROOT, arguments.mode, arguments.output, arguments.override)
    except (BuildError, OSError) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 1
    print(f"built {arguments.output}")
    for line in result.ignored:
        print(f"ignored: {line}")
    print(f"mode: {result.mode} ({len(result.components)} standard modules)")
    for component in result.components:
        origin = "" if component.origin == "repository" else f"  [{component.origin}: {component.path}]"
        print(f"  {component.name}{origin}")
    print(f"project: {result.project_id}")
    print(f"sha256: {result.sha256}")
    print("Excel compiles the project from source when it opens the workbook.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
