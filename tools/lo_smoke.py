#!/usr/bin/env python3
"""Open a workbook built from source in headless LibreOffice and check its VBA project.

Usage: python3 tools/lo_smoke.py [--mode Production|Development|Demo] [WORKBOOK]

Without WORKBOOK, the workbook of --mode (default: Development) is built first
with tools/build_workbook.py, in a temporary folder. WORKBOOK must be a build
of that mode from the current source.

LibreOffice opens the workbook hidden, with its VBA project loaded and macro
execution off: nothing in the workbook runs. The check passes when the
project holds exactly the document modules ThisWorkbook and Sheet1 and the
modules that tools/vba-projects.json lists for the mode, each with the code of
its source file (without its Attribute lines, which LibreOffice does not keep),
and when the sheet and the workbook carry the code names of a new Excel
workbook.

This proves that the package and its vbaProject.bin can be read, not that the
VBA compiles or runs: LibreOffice is not a supported host, and this is never
Excel evidence. Run it with a Python that has LibreOffice's UNO bridge (on
Debian and Ubuntu, /usr/bin/python3 with the python3-uno package).

Exit status: 0 when the check passes, 1 when it fails or LibreOffice cannot
open the workbook, 2 when it cannot run: an error in the arguments, a missing
WORKBOOK, or LibreOffice (soffice) or its UNO bridge not found.
"""

from __future__ import annotations

import argparse
import importlib
import importlib.util
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_workbook  # noqa: E402
import vba_project  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
LIBRARY = "VBAProject"
# The lines LibreOffice puts before the code of a module it imports from VBA.
LOADED_HEADER = ("Rem Attribute VBA_ModuleType=", "Option VBASupport 1")
# LibreOffice comments out a procedure declaration whose End it cannot pair
# with it, which happens to a Sub declared in each branch of an #If block. The
# line is otherwise kept, so the prefix is removed before the comparison.
UNMATCHED_PREFIX = "Rem removed unmatched Sub/End: "
DOCUMENT_CODE = "Option Explicit"
CONNECT_ATTEMPTS = 120


def code_text(source: bytes) -> str:
    """A module's code as LibreOffice shows it: LF line endings, no Attribute
    lines, no trailing line break."""
    lines = build_workbook.module_source(source).decode(vba_project.ENCODING).split("\r\n")
    return "\n".join(line for line in lines if not line.startswith("Attribute ")).rstrip("\n")


def loaded_code(text: str) -> str:
    """The code of a module LibreOffice loaded, without the header it adds."""
    lines = text.replace("\r\n", "\n").split("\n")
    while lines and lines[0].startswith(LOADED_HEADER):
        lines.pop(0)
    return "\n".join(line.removeprefix(UNMATCHED_PREFIX) for line in lines).rstrip("\n")


def expected_modules(root: Path, mode: str) -> dict[str, str]:
    """Every module the built workbook holds, by name, with its code."""
    parts, _ = build_workbook.components(root, mode)
    expected = {build_workbook.WORKBOOK_CODE_NAME: DOCUMENT_CODE}
    expected.update((code_name, DOCUMENT_CODE) for _, code_name in build_workbook.SHEETS)
    expected.update((part.name, code_text(part.path.read_bytes())) for part in parts)
    return expected


def findings(expected: Mapping[str, str], loaded: Mapping[str, str], sheets: Sequence[tuple[str, str]],
             workbook_code_name: str) -> list[str]:
    """What LibreOffice loaded that differs from the build: modules, code and code names."""
    found: list[str] = []
    found += [f"module {name} is missing" for name in expected if name not in loaded]
    found += [f"module {name} is not in the build" for name in loaded if name not in expected]
    found += [f"module {name}: its code differs from the source" for name in expected
              if name in loaded and loaded_code(loaded[name]) != expected[name]]
    wanted = list(build_workbook.SHEETS)
    if list(sheets) != wanted:
        found.append(f"sheets (name, code name) are {list(sheets)}, not {wanted}")
    if workbook_code_name != build_workbook.WORKBOOK_CODE_NAME:
        found.append(f"the workbook's code name is {workbook_code_name!r}, not {build_workbook.WORKBOOK_CODE_NAME}")
    return found


# ---------------------------------------------------------------------------
# LibreOffice


def uno() -> Any:
    return importlib.import_module("uno")


def free_port() -> int:
    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        return int(probe.getsockname()[1])


def property_value(name: str, value: object) -> Any:
    item = uno().createUnoStruct("com.sun.star.beans.PropertyValue")
    item.Name = name
    item.Value = value
    return item


def connect(port: int) -> Any:
    local = uno().getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(CONNECT_ATTEMPTS):
        try:
            return resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
        except Exception:  # noqa: BLE001 - the office is still starting
            time.sleep(0.5)
    raise RuntimeError("LibreOffice did not accept a connection")


def load_vba(context: Any) -> None:
    """Let Calc import the VBA project of the workbooks it opens."""
    provider = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", context)
    access = provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationUpdateAccess",
        (property_value("nodepath", "/org.openoffice.Office.Calc/Filter/Import/VBA"),))
    access.setPropertyValue("Load", True)
    access.commitChanges()


# What LibreOffice loaded: the modules with their code, the sheets as (name,
# code name), and the workbook's code name.
Loaded = tuple[dict[str, str], list[tuple[str, str]], str]


def read_workbook(context: Any, workbook: Path) -> Loaded:
    """The modules of the workbook's VBA project with their code, its sheets
    as (name, code name), and its code name. No macro runs."""
    desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
    try:
        # MacroExecutionMode 0: NEVER_EXECUTE.
        document = desktop.loadComponentFromURL(
            workbook.resolve().as_uri(), "_blank", 0,
            (property_value("Hidden", True), property_value("MacroExecutionMode", 0)))
        if document is None:
            raise RuntimeError(f"LibreOffice could not open {workbook}")
        try:
            libraries = document.BasicLibraries
            modules: dict[str, str] = {}
            if libraries.hasByName(LIBRARY):
                library = libraries.getByName(LIBRARY)
                modules = {str(name): str(library.getByName(name)) for name in library.getElementNames()}
            sheets = [(str(sheet.Name), str(sheet.CodeName))
                      for sheet in (document.Sheets.getByIndex(index) for index in range(document.Sheets.Count))]
            return modules, sheets, str(document.CodeName)
        finally:
            document.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:  # noqa: BLE001 - the office may already be gone
            pass


def open_in_libreoffice(soffice: str, workbook: Path, profile: Path) -> Loaded:
    """Start a headless LibreOffice with a profile of its own and read the workbook."""
    port = free_port()
    office = subprocess.Popen(
        [soffice, "--headless", "--invisible", "--norestore", "--nologo", f"-env:UserInstallation={profile.as_uri()}",
         f"--accept=socket,host=127.0.0.1,port={port};urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        context = connect(port)
        load_vba(context)
        return read_workbook(context, workbook)
    finally:
        try:
            office.wait(timeout=60)
        except subprocess.TimeoutExpired:
            office.kill()
            office.wait()


def parse_arguments(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Open a workbook built from source in headless LibreOffice and check its VBA project.",
        epilog="Exit status: 0 pass, 1 fail, 2 cannot run (arguments, missing workbook, LibreOffice or UNO).")
    parser.add_argument("--mode", type=build_workbook.mode_name, default="Development",
                        metavar="|".join(build_workbook.MODES),
                        help="the configuration the workbook holds (default: Development)")
    parser.add_argument("workbook", nargs="?", type=Path, metavar="WORKBOOK",
                        help="a workbook built for --mode; without it, one is built first")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    arguments = parse_arguments(argv)
    if arguments.workbook is not None and not arguments.workbook.is_file():
        print(f"lo-smoke: ERROR: no workbook at {arguments.workbook}", file=sys.stderr)
        return 2
    soffice = shutil.which("soffice")
    if soffice is None or importlib.util.find_spec("uno") is None:
        print("lo-smoke: SKIPPED: needs LibreOffice (soffice) and its Python UNO bridge (python3-uno)")
        return 2
    try:
        expected = expected_modules(ROOT, arguments.mode)
    except (build_workbook.BuildError, OSError) as error:
        print(f"lo-smoke: FAIL: {error}")
        return 1
    with tempfile.TemporaryDirectory(prefix="lo-smoke-") as temporary:
        base = Path(temporary)
        workbook = arguments.workbook
        try:
            if workbook is None:
                workbook = base / f"EXCEL_UI_{arguments.mode}.xlsm"
                build_workbook.build(ROOT, arguments.mode, workbook)
                print(f"lo-smoke: built {arguments.mode} workbook")
            modules, sheets, code_name = open_in_libreoffice(soffice, workbook, base / "profile")
        except Exception as error:  # noqa: BLE001 - a build or UNO error, reported as a failure
            print(f"lo-smoke: FAIL: {type(error).__name__}: {error}")
            return 1
    problems = findings(expected, modules, sheets, code_name)
    for name in expected:
        print(f"  {name:<32} {'present' if name in modules else 'MISSING'}")
    for problem in problems:
        print(f"lo-smoke: {problem}")
    verdict = "FAIL" if problems else "PASS"
    print(f"lo-smoke: {verdict}: {len(expected)} modules expected, {len(modules)} loaded "
          "(LibreOffice read the package; this is not Excel evidence)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
