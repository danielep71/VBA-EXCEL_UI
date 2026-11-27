"""Write a title-bar negative control into /.mutation-scratch/ (#66, #6).

Each control is src/M_EXCEL_UI_TITLEBAR.bas with one or more decisions put
back to an older, weaker form. Nothing else changes: the module keeps its
Attribute VB_Name line, its CRLF line endings and every other byte. The file
is written to .mutation-scratch/titlebar-6/, which Git ignores and source
archives exclude, so a control can never be committed beside src/. That folder
must hold nothing else, because the workbook synchronizer imports every module
it finds in a -SourceDir folder.

Variants (--variant):

    v112       the show fallback condition as v1.1.2 had it, which consults
               the value of the stored baseline instead of whether one was
               captured and whether it carries the caption (#66, the default):

                   If Not m_FrameStates(Slot).HasBaseline _
                      Or (RestoreBits And WS_CAPTION) <> WS_CAPTION Then   current
                   If RestoreBits = 0 Then                                 v1.1.2

    fallback   the show fallback condition tests any caption bit:
                   Or (RestoreBits And WS_CAPTION) = 0 Then
    reader     UI_TryGetTitleBarVisibleForHwnd tests any caption bit:
                   IsVisible = ((StyleValue And WS_CAPTION) <> 0)
    readbacks  both show readbacks (the no-op branch and the post-write
               branch) test any caption bit:
                   If (VerifiedStyle And WS_CAPTION) = 0 Then
    anybit     all four any-bit tests together, as they stood when #6 was
               reopened for the complete-mask correction

WS_CAPTION is WS_BORDER Or WS_DLGFRAME. The current source tests
(style And WS_CAPTION) = WS_CAPTION everywhere, and the controls above put
the any-bit form, (style And WS_CAPTION) <> 0, back.

The tool fails, and writes nothing, when a condition to be replaced is not
found exactly once, when the source is not CRLF ASCII or does not declare the
expected module, or when the folder holds another file.

Usage:
    python3 tools/make_titlebar_mutant.py [--variant NAME] [--repo PATH]

CONTRIBUTING.md describes how to run the controls in Excel.
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

# The v1.1.2 control (#66). CURRENT and MUTANT keep their original names.
CURRENT = (
    b"                If Not m_FrameStates(Slot).HasBaseline _\r\n"
    b"                   Or (RestoreBits And WS_CAPTION) <> WS_CAPTION Then\r\n"
)
MUTANT = b"                If RestoreBits = 0 Then\r\n"

# The complete-mask decisions of #6, each paired with its any-bit form.
FALLBACK_CURRENT = b"                   Or (RestoreBits And WS_CAPTION) <> WS_CAPTION Then\r\n"
FALLBACK_MUTANT = b"                   Or (RestoreBits And WS_CAPTION) = 0 Then\r\n"
READER_CURRENT = b"        IsVisible = ((StyleValue And WS_CAPTION) = WS_CAPTION)\r\n"
READER_MUTANT = b"        IsVisible = ((StyleValue And WS_CAPTION) <> 0)\r\n"
# The two readbacks differ only in indentation and in the line after them, and
# the shorter one is a substring of the longer, so each pattern spans both lines.
NOOP_CURRENT = (
    b"If (VerifiedStyle And WS_CAPTION) <> WS_CAPTION Then\r\n"
    b"                        FailMsg = _\r\n"
)
NOOP_MUTANT = (
    b"If (VerifiedStyle And WS_CAPTION) = 0 Then\r\n"
    b"                        FailMsg = _\r\n"
)
WRITE_CURRENT = (
    b"If (VerifiedStyle And WS_CAPTION) <> WS_CAPTION Then\r\n"
    b'                FailMsg = "title-bar show did not produce a visible caption"\r\n'
)
WRITE_MUTANT = (
    b"If (VerifiedStyle And WS_CAPTION) = 0 Then\r\n"
    b'                FailMsg = "title-bar show did not produce a visible caption"\r\n'
)

DEFAULT_VARIANT = "v112"

# variant -> ordered (label, current, mutant) replacements
VARIANTS: dict[str, tuple[tuple[str, bytes, bytes], ...]] = {
    "v112": (("condition", CURRENT, MUTANT),),
    "fallback": (("fallback condition", FALLBACK_CURRENT, FALLBACK_MUTANT),),
    "reader": (("visibility reader", READER_CURRENT, READER_MUTANT),),
    "readbacks": (
        ("no-op readback", NOOP_CURRENT, NOOP_MUTANT),
        ("post-write readback", WRITE_CURRENT, WRITE_MUTANT),
    ),
    "anybit": (
        ("fallback condition", FALLBACK_CURRENT, FALLBACK_MUTANT),
        ("visibility reader", READER_CURRENT, READER_MUTANT),
        ("no-op readback", NOOP_CURRENT, NOOP_MUTANT),
        ("post-write readback", WRITE_CURRENT, WRITE_MUTANT),
    ),
}


class MutantError(Exception):
    """The control cannot be produced from this source."""


def _replacements(variant: str) -> tuple[tuple[str, bytes, bytes], ...]:
    try:
        return VARIANTS[variant]
    except KeyError:
        raise MutantError(f"unknown variant {variant!r}; choose one of {', '.join(VARIANTS)}") from None


def mutate(data: bytes, variant: str = DEFAULT_VARIANT) -> bytes:
    """Return the module with only the variant's conditions replaced."""
    replacements = _replacements(variant)
    if data.count(b"\n") != data.count(b"\r\n"):
        raise MutantError(f"{SOURCE.as_posix()} has bare LF line endings; check it out with CRLF")
    if not data.startswith(VB_NAME_LINE):
        raise MutantError(f"{SOURCE.as_posix()} does not start with {VB_NAME_LINE.decode('ascii').strip()}")
    try:
        data.decode("ascii")
    except UnicodeDecodeError as exc:
        raise MutantError(f"{SOURCE.as_posix()} is not ASCII") from exc
    for label, current, mutant in replacements:
        found = data.count(current)
        if found != 1:
            raise MutantError(
                f"expected the current {label} exactly once in "
                f"{SOURCE.as_posix()}, found {found}; the worker has changed, so update "
                f"the {variant} variant in this tool before producing a control")
        if mutant in data:
            raise MutantError(f"{SOURCE.as_posix()} already contains the {variant} {label}")
    for _label, current, mutant in replacements:
        data = data.replace(current, mutant)
    return data


def mutated_lines(data: bytes, variant: str = DEFAULT_VARIANT) -> list[tuple[str, int]]:
    """Return each replaced condition's label and 1-based line in the source."""
    return [(label, data[:data.index(current)].count(b"\r\n") + 1)
            for label, current, _mutant in _replacements(variant)]


def mutated_line(data: bytes, variant: str = DEFAULT_VARIANT) -> int:
    """Return the 1-based line number of the variant's first condition."""
    return mutated_lines(data, variant)[0][1]


def write(repo: Path, variant: str = DEFAULT_VARIANT) -> Path:
    """Write the control into the scratch folder and return its path."""
    data = (repo / SOURCE).read_bytes()
    control = mutate(data, variant)
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
        description="Write a title-bar negative control into .mutation-scratch/titlebar-6/.")
    parser.add_argument("--repo", type=Path, default=Path(__file__).resolve().parent.parent,
                        help="repository root (default: the checkout that holds this tool)")
    parser.add_argument("--variant", default=DEFAULT_VARIANT, choices=sorted(VARIANTS),
                        help=f"control to write (default: {DEFAULT_VARIANT})")
    args = parser.parse_args(argv)
    repo: Path = args.repo
    variant: str = args.variant
    try:
        source = (repo / SOURCE).read_bytes()
        target = write(repo, variant)
    except (MutantError, OSError) as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 1
    control = target.read_bytes()
    print(f"variant: {variant}")
    print(f"control: {target}")
    for label, line in mutated_lines(source, variant):
        print(f"{label} replaced at {SOURCE.as_posix()} line {line}")
    print(f"source  SHA-256 {hashlib.sha256(source).hexdigest()}")
    print(f"control SHA-256 {hashlib.sha256(control).hexdigest()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
