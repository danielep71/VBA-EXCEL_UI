"""Write the #6 title-bar negative control into /.mutation-scratch/ (#66).

The control is src/M_EXCEL_UI_TITLEBAR.bas with the captionless fallback
condition of UI_TrySetTitleBarVisibleForHwndWorker put back to its v1.1.2
form, which consults the value of the stored baseline instead of whether a
baseline was captured and whether it contains WS_CAPTION:

    If Not m_FrameStates(Slot).HasBaseline _
       Or (RestoreBits And WS_CAPTION) = 0 Then      current
    If RestoreBits = 0 Then                          v1.1.2

Nothing else changes: the module keeps its Attribute VB_Name line, its CRLF
line endings and every other byte. The file is written to
.mutation-scratch/titlebar-6/, which Git ignores and source archives exclude,
so the mutant can never be committed beside src/. That folder must hold
nothing else, because the workbook synchronizer imports every module it finds
in a -SourceDir folder.

The tool fails, and writes nothing, when the current condition is not found
exactly once, when the source is not CRLF ASCII or does not declare the
expected module, or when the folder holds another file.

Usage:
    python3 tools/make_titlebar_mutant.py [--repo PATH]

CONTRIBUTING.md describes how to run the control in Excel.
"""

from __future__ import annotations

import argparse
import hashlib
import sys
from pathlib import Path

SOURCE = Path("src") / "M_EXCEL_UI_TITLEBAR.bas"
SCRATCH = Path(".mutation-scratch") / "titlebar-6"
MODULE_FILE = "M_EXCEL_UI_TITLEBAR.bas"
VB_NAME_LINE = b'Attribute VB_Name = "M_EXCEL_UI_TITLEBAR"\r\n'

CURRENT = (
    b"                If Not m_FrameStates(Slot).HasBaseline _\r\n"
    b"                   Or (RestoreBits And WS_CAPTION) = 0 Then\r\n"
)
MUTANT = b"                If RestoreBits = 0 Then\r\n"


class MutantError(Exception):
    """The control cannot be produced from this source."""


def mutate(data: bytes) -> bytes:
    """Return the module with only the fallback condition replaced."""
    if data.count(b"\n") != data.count(b"\r\n"):
        raise MutantError(f"{SOURCE.as_posix()} has bare LF line endings; check it out with CRLF")
    if not data.startswith(VB_NAME_LINE):
        raise MutantError(f"{SOURCE.as_posix()} does not start with {VB_NAME_LINE.decode('ascii').strip()}")
    try:
        data.decode("ascii")
    except UnicodeDecodeError as exc:
        raise MutantError(f"{SOURCE.as_posix()} is not ASCII") from exc
    found = data.count(CURRENT)
    if found != 1:
        raise MutantError(
            f"expected the current captionless fallback condition exactly once in "
            f"{SOURCE.as_posix()}, found {found}; the worker has changed, so update "
            f"CURRENT in this tool before producing a control")
    if MUTANT in data:
        raise MutantError(f"{SOURCE.as_posix()} already contains the v1.1.2 condition")
    return data.replace(CURRENT, MUTANT)


def mutated_line(data: bytes) -> int:
    """Return the 1-based line number of the condition in the source."""
    return data[:data.index(CURRENT)].count(b"\r\n") + 1


def write(repo: Path) -> Path:
    """Write the control into the scratch folder and return its path."""
    data = (repo / SOURCE).read_bytes()
    control = mutate(data)
    folder = repo / SCRATCH
    if folder.exists():
        others = sorted(p.name for p in folder.iterdir() if p.name != MODULE_FILE)
        if others:
            raise MutantError(
                f"{SCRATCH.as_posix()} also holds {', '.join(others)}; the synchronizer would "
                f"import them too, so move them out first")
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / MODULE_FILE
    target.write_bytes(control)
    return target


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Write the #6 title-bar negative control into .mutation-scratch/titlebar-6/.")
    parser.add_argument("--repo", type=Path, default=Path(__file__).resolve().parent.parent,
                        help="repository root (default: the checkout that holds this tool)")
    args = parser.parse_args(argv)
    repo: Path = args.repo
    try:
        source = (repo / SOURCE).read_bytes()
        target = write(repo)
    except (MutantError, OSError) as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1
    control = target.read_bytes()
    print(f"control: {target}")
    print(f"condition replaced at {SOURCE.as_posix()} line {mutated_line(source)}")
    print(f"source  SHA-256 {hashlib.sha256(source).hexdigest()}")
    print(f"control SHA-256 {hashlib.sha256(control).hexdigest()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
