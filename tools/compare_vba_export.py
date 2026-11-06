#!/usr/bin/env python3
"""Compare the VBA exported from a workbook with the source of one commit.

The second half of the source round trip (docs/WORKBOOK_SYNC.md): a workbook
synchronized from one commit with tools/Sync-EXCEL_UI-VBA.ps1, then exported
with tools/Export-EXCEL_UI-VBA.ps1, must hold exactly the modules that
tools/vba-projects.json lists for its configuration, each with the same code
and attributes as the source of that commit. The source is read from Git at
that commit, not from the working tree.

Ignored: line endings, trailing blank lines, and the letter case of code
outside string literals and comments. VBA names are case-insensitive and the
VBE keeps one spelling per name across a project, so it re-cases identifiers
on import; the lines that differ only that way are counted. A workbook always
holds document modules (ThisWorkbook and one per sheet), which Excel creates
and writes, with attributes of its own; this repository has none, so they are
not compared with the source, but each must hold no code other than
``Option Explicit``. Any other component is compared.

Source parity says nothing about which workbook the files came from. The
exporter writes export-manifest.json: the SHA-256 of the workbook it opened
and of every file it exported. With --workbook the tool checks the workbook
and the files against it and reports the binding as VERIFIED, as FAIL, or as
UNVERIFIED for an export without a manifest or a run without --workbook. A
manifest is a record, not a signature: it binds an export to a workbook, but
it does not attest who made either.

Exit status: 0 when the export matches and the binding did not fail, 1 when
it does not, 2 for an error in the arguments or the input.
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import re
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

MANIFEST = "tools/vba-projects.json"
MODES = ("production", "development", "demo")
EXPORT_MANIFEST = "export-manifest.json"
EXPORT_FORMAT = "excel-ui-vba-export-1"
VBA_PROJECT_PART = "xl/vbaProject.bin"
VBA_ENCODING = "cp1252"
SUFFIXES = (".bas", ".cls", ".frm")
# A form's binary properties, which the VBE exports beside its .frm: not
# compared with the source, but bound to the workbook like the components.
RESOURCE_SUFFIX = ".frx"
VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]+)"$')
# The VB_Base of a document module names Excel's Workbook (00020819),
# Worksheet (00020820) or Chart (00020821) class.
DOCUMENT_BASE = re.compile(r'^Attribute VB_Base = "0\{(00020819|00020820|00020821)-')
# A Rem comment where a statement starts (see fold_case).
REM = re.compile(r"rem(?=\s|$)", re.IGNORECASE)
# Lines the VBE writes before the code of a class or form: the VERSION line,
# the BEGIN ... END property block (title case and nested in a form) and the
# Attribute lines.
CLASS_HEADER = re.compile(r"^(VERSION \d+\.\d+( CLASS)?|(?i:begin)(\s.*)?|(?i:end)|\s+\S.*)$")
# The only line a document module may hold: the VBE adds it to every new
# module when "Require Variable Declaration" is on.
EMPTY_DOCUMENT_LINE = re.compile(r"^\s*(option\s+explicit\s*)?$", re.IGNORECASE)


@dataclass(frozen=True)
class Component:
    name: str
    attributes: tuple[str, ...]
    body: tuple[str, ...]


@dataclass(frozen=True)
class Exported:
    file: str
    component: Component
    document: bool


def parse(raw: bytes, label: str) -> Component:
    """Split a VBA file into its VB_Name, its Attribute lines and its code."""
    lines = raw.decode(VBA_ENCODING).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    index = 0
    attributes: list[str] = []
    while index < len(lines):
        line = lines[index]
        if line.startswith("Attribute "):
            attributes.append(line)
        elif not CLASS_HEADER.match(line) or attributes:
            break
        index += 1
    body = lines[index:]
    while body and body[-1] == "":
        body.pop()
    names = [m.group(1) for m in map(VB_NAME.match, attributes) if m]
    if len(names) != 1:
        raise ValueError(f"{label}: expected one Attribute VB_Name line, found {len(names)}")
    return Component(names[0], tuple(attributes), tuple(body))


def fold_case(line: str) -> str:
    """``line`` with its code lower-cased, string literals and the comment kept.

    A comment starts at an apostrophe outside a string, or at ``Rem`` where a
    statement starts: the start of the line, after a line number or after a
    colon."""
    folded: list[str] = []
    in_string = False
    statement_start = True
    index = 0
    while index < len(line):
        char = line[index]
        if in_string:
            in_string = char != '"'
        elif char == '"':
            in_string, statement_start = True, False
        elif char == "'":
            return "".join(folded) + line[index:]
        elif char == ":":
            statement_start = True
        elif statement_start and REM.match(line, index):
            return "".join(folded) + "rem" + line[index + 3:]
        elif statement_start and char.isdigit() and not "".join(folded).strip(" \t0123456789"):
            pass
        elif not char.isspace():
            statement_start = False
        folded.append(char if in_string or char == '"' else char.lower())
        index += 1
    return "".join(folded)


def git_bytes(root: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(["git", "-C", str(root), *args], capture_output=True, check=False)


def committed(root: Path, sha: str, path: str) -> bytes:
    completed = git_bytes(root, "show", f"{sha}:{path}")
    if completed.returncode != 0:
        raise ValueError(f"{path}: not found at {sha}")
    return completed.stdout


def resolve_commit(root: Path, commit: str) -> str:
    resolved = git_bytes(root, "rev-parse", "--verify", "--quiet", f"{commit}^{{commit}}")
    if resolved.returncode != 0:
        raise ValueError(f"unknown commit: {commit}")
    return resolved.stdout.decode("ascii").strip()


def expected_components(root: Path, sha: str, mode: str) -> dict[str, tuple[str, Component]]:
    """The modules of ``mode`` at ``sha``, by lower-case VB_Name: (path, component)."""
    try:
        manifest: Any = json.loads(committed(root, sha, MANIFEST).decode("utf-8"))
        modules = manifest["configurations"][mode]["modules"]
        documents = manifest.get("document_modules", [])
    except (UnicodeDecodeError, json.JSONDecodeError, KeyError, TypeError, AttributeError) as error:
        raise ValueError(f"{MANIFEST} at {sha}: no configuration '{mode}' ({error})") from error
    if documents:
        raise ValueError(f"{MANIFEST} at {sha} lists document modules, which this tool does not compare")
    expected: dict[str, tuple[str, Component]] = {}
    for path in modules:
        component = parse(committed(root, sha, str(path)), str(path))
        key = component.name.lower()
        if key in expected:
            raise ValueError(f"{MANIFEST} at {sha}: VB_Name {component.name} is listed twice in '{mode}'")
        expected[key] = (str(path), component)
    return expected


def recorded_documents(folder: Path) -> set[str]:
    """The lower-case files export-manifest.json records as document modules."""
    try:
        record: Any = json.loads((folder / EXPORT_MANIFEST).read_bytes().decode("utf-8-sig"))
        return {str(item["file"]).lower() for item in record["components"] if item.get("type") == "document"}
    except (OSError, ValueError, KeyError, TypeError, AttributeError):
        return set()  # an absent or invalid manifest is reported by check_binding


def exported_components(folder: Path) -> dict[str, Exported]:
    """The exported components by lower-case VB_Name.

    A .cls is a document module when export-manifest.json records it as one
    or when its VB_Base names an Excel workbook, worksheet or chart class."""
    documents = recorded_documents(folder)
    exported: dict[str, Exported] = {}
    for path in sorted(folder.iterdir()):
        if path.suffix.lower() not in SUFFIXES:
            continue
        component = parse(path.read_bytes(), path.name)
        key = component.name.lower()
        if key in exported:
            raise ValueError(f"{path.name}: VB_Name {component.name} exported twice")
        document = path.suffix.lower() == ".cls" and (
            path.name.lower() in documents or any(DOCUMENT_BASE.match(line) for line in component.attributes))
        exported[key] = Exported(path.name, component, document)
    return exported


def first_difference(source: tuple[str, ...], export: tuple[str, ...]) -> str:
    for number, (left, right) in enumerate(zip(source, export, strict=False), start=1):
        if fold_case(left) != fold_case(right):
            return f"line {number}: source {left!r}, export {right!r}"
    shorter = min(len(source), len(export))
    return f"line {shorter + 1}: source has {len(source)} lines, export {len(export)}"


def document_findings(exported: dict[str, Exported], expected: dict[str, tuple[str, Component]]) -> list[str]:
    """Document modules that hold code: the source has none to compare it with."""
    findings = []
    for key in sorted(exported):
        item = exported[key]
        if not item.document or key in expected:
            continue
        code = [(number, line) for number, line in enumerate(item.component.body, start=1)
                if not EMPTY_DOCUMENT_LINE.match(line)]
        if code:
            number, line = code[0]
            findings.append(f"document module {item.component.name} ({item.file}) holds code that is not "
                            f"in the source, first at line {number}: {line!r}")
    return findings


def compare(expected: dict[str, tuple[str, Component]], exported: dict[str, Exported]) -> tuple[list[str], int]:
    """The differences between the source and the export, empty when they
    match, and the number of lines that differ only in letter case."""
    findings = [f"missing from the export: {expected[key][1].name} ({expected[key][0]})"
                for key in sorted(set(expected) - set(exported))]
    findings += [f"not in the source for this mode: {exported[key].component.name} ({exported[key].file})"
                 for key in sorted(set(exported) - set(expected)) if not exported[key].document]
    findings += document_findings(exported, expected)
    recased = 0
    for key in sorted(set(expected) & set(exported)):
        path, source = expected[key]
        export = exported[key].component
        if source.name != export.name:
            findings.append(f"{path}: VB_Name {source.name} exported as {export.name}")
        if source.attributes != export.attributes:
            changed = [line for line in difflib.unified_diff(
                source.attributes, export.attributes, lineterm="", n=0)
                if line[:1] in "+-" and line[:3] not in ("+++", "---")]
            findings.append(f"{path}: attributes differ: {'; '.join(changed)}")
        if [fold_case(line) for line in source.body] != [fold_case(line) for line in export.body]:
            findings.append(f"{path}: code differs at {first_difference(source.body, export.body)}")
        else:
            recased += sum(left != right for left, right in zip(source.body, export.body, strict=True))
    return findings, recased


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@dataclass(frozen=True)
class ExportRecord:
    workbook_name: str
    workbook_sha256: str
    files: dict[str, tuple[str, str, str]]  # lower-case file -> (file, VB_Name, sha256)


def load_export_manifest(path: Path) -> ExportRecord:
    """export-manifest.json; ValueError when it is not a valid manifest."""
    try:
        record: Any = json.loads(path.read_bytes().decode("utf-8-sig"))
        if not isinstance(record, dict) or record.get("format") != EXPORT_FORMAT:
            raise ValueError(f"format is not {EXPORT_FORMAT}")
        name = record["workbook"]["file"]
        if not isinstance(name, str) or not name:
            raise ValueError("workbook.file is not a file name")
        files = {item["file"].lower(): (item["file"], item["name"], item["sha256"].lower())
                 for item in record["components"]}
        if len(files) != len(record["components"]):
            raise ValueError("a file is listed twice")
        return ExportRecord(name, record["workbook"]["sha256"].lower(), files)
    except (ValueError, KeyError, TypeError, AttributeError) as error:
        raise ValueError(f"{EXPORT_MANIFEST} is not a valid export manifest: {error}") from error


def file_findings(folder: Path, record: ExportRecord) -> list[str]:
    """The exported files that are missing, unlisted, changed or renamed
    against the manifest."""
    listed = record.files
    present = {item.name.lower(): item for item in folder.iterdir()
               if item.is_file() and item.suffix.lower() in (*SUFFIXES, RESOURCE_SUFFIX)}
    findings = [f"{EXPORT_MANIFEST} lists {listed[key][0]}, which is not in the export"
                for key in sorted(set(listed) - set(present))]
    findings += [f"{present[key].name} is in the export but not in {EXPORT_MANIFEST}"
                 for key in sorted(set(present) - set(listed))]
    for key in sorted(set(listed) & set(present)):
        name, component, digest = listed[key]
        data = present[key].read_bytes()
        if sha256(data) != digest:
            findings.append(f"{name} changed after the export: its SHA-256 is not the one in {EXPORT_MANIFEST}")
        elif present[key].suffix.lower() != RESOURCE_SUFFIX:
            try:
                exported_name = parse(data, name).name
            except ValueError:
                continue  # reported by the source comparison
            if exported_name != component:
                findings.append(f"{name} declares VB_Name {exported_name}, {EXPORT_MANIFEST} names {component}")
    return findings


def is_macro_workbook(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            return VBA_PROJECT_PART in archive.namelist()
    except (zipfile.BadZipFile, OSError):
        return False


def check_binding(folder: Path, workbook: Path | None) -> tuple[str, list[str], list[str]]:
    """The binding of the export to the workbook, VERIFIED, UNVERIFIED or
    FAIL, the findings that make it FAIL and notes that do not.

    A manifest that does not match the exported files fails even without
    --workbook; without either, or without a manifest, the binding is
    UNVERIFIED. The workbook is bound by its bytes: a workbook renamed after
    the export, as a download often is, gets a note, not a failure."""
    path = folder / EXPORT_MANIFEST
    if not path.is_file():
        return "UNVERIFIED", [], []
    try:
        record = load_export_manifest(path)
    except ValueError as error:
        return "FAIL", [str(error)], []
    findings = file_findings(folder, record)
    notes: list[str] = []
    if workbook is not None:
        digest = sha256(workbook.read_bytes())
        same_bytes = digest == record.workbook_sha256
        if not same_bytes:
            findings.append(f"{workbook.name} is not the workbook the export was made from: "
                            f"sha256 {digest}, {EXPORT_MANIFEST} records {record.workbook_sha256}")
        if not is_macro_workbook(workbook):
            findings.append(f"{workbook.name} is not a macro-enabled workbook ({VBA_PROJECT_PART} not found)")
        if same_bytes and workbook.name.lower() != record.workbook_name.lower():
            notes.append(f"{EXPORT_MANIFEST} names the workbook {record.workbook_name}; "
                         "it was renamed since the export")
    if findings:
        return "FAIL", findings, notes
    return ("VERIFIED" if workbook is not None else "UNVERIFIED"), [], notes


def binding_note(binding: str, folder: Path, workbook: Path | None) -> str:
    if binding == "VERIFIED":
        return f"{EXPORT_MANIFEST} records this workbook and every exported file"
    if binding == "FAIL":
        return f"the export, {EXPORT_MANIFEST} and the workbook do not agree"
    if not (folder / EXPORT_MANIFEST).is_file():
        return f"no {EXPORT_MANIFEST}: the export is not bound to a workbook"
    return "no --workbook given: the export is not checked against a workbook"


def parse_arguments(arguments: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent.parent,
                        help="repository checkout to read the commit from (default: this one)")
    parser.add_argument("--export", type=Path, required=True, help="folder written by Export-EXCEL_UI-VBA.ps1")
    parser.add_argument("--mode", choices=MODES, required=True, help="configuration the workbook holds")
    parser.add_argument("--commit", default="HEAD", help="commit the workbook was synchronized from (default: HEAD)")
    parser.add_argument("--workbook", type=Path,
                        help="the .xlsm the export was made from, checked against export-manifest.json")
    return parser.parse_args(arguments)


def main(arguments: list[str] | None = None) -> int:
    options = parse_arguments(arguments)
    try:
        sha = resolve_commit(options.root, options.commit)
        if not options.export.is_dir():
            raise ValueError(f"export folder not found: {options.export}")
        if options.workbook is not None and not options.workbook.is_file():
            raise ValueError(f"workbook not found: {options.workbook}")
        expected = expected_components(options.root, sha, options.mode)
        exported = exported_components(options.export)
        findings, recased = compare(expected, exported)
        binding, binding_findings, binding_notes = check_binding(options.export, options.workbook)
    except ValueError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    documents = sorted(item.component.name for key, item in exported.items() if item.document and key not in expected)
    print(f"Commit:    {sha}")
    print(f"Mode:      {options.mode} ({len(expected)} modules expected, "
          f"{len(exported) - len(documents)} exported)")
    print(f"Documents: {len(documents)} document module(s), not compared with the source"
          + (f": {', '.join(documents)}" if documents else ""))
    if options.workbook is not None:
        print(f"Workbook:  {options.workbook.name} sha256 {sha256(options.workbook.read_bytes())}")
    print(f"Re-cased:  {recased} line(s) differ only in the letter case of code (ignored)")
    print(f"Binding:   {binding} - {binding_note(binding, options.export, options.workbook)}")
    for note in binding_notes:
        print(f"  NOTE {note}")
    for finding in findings + binding_findings:
        print(f"  FAIL {finding}")
    passed = not findings and binding != "FAIL"
    print(f"RESULT={'PASS' if passed else 'FAIL'}; modules={len(expected)}; "
          f"differences={len(findings)}; binding={binding}")
    return 0 if passed else 1


if __name__ == "__main__":
    sys.exit(main())
