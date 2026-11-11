"""Tests for tools/build_workbook.py and tools/vba_project.py, without Excel.

The built workbook is read back with readers written here from the
specifications, independently of the writer: a ZIP and XML reading of the
package, an [MS-CFB] compound-file reader, an [MS-OVBA] decompressor, a dir
stream parser and the decryption of the PROJECT protection fields. A defect of
the writer therefore fails here, although these tests cannot show that Excel
opens, compiles or runs the workbook.
"""

from __future__ import annotations

import contextlib
import functools
import io
import json
import os
import random
import re
import shutil
import struct
import sys
import tempfile
import unittest
import zipfile
from collections.abc import Iterator
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import build_workbook  # noqa: E402
import vba_project  # noqa: E402

MANIFEST: dict[str, Any] = json.loads((ROOT / "tools" / "vba-projects.json").read_text(encoding="utf-8"))
ENDOFCHAIN = 0xFFFFFFFE
NOSTREAM = 0xFFFFFFFF
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
PACKAGE_RELATIONSHIPS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CONTENT_TYPES = "{http://schemas.openxmlformats.org/package/2006/content-types}"
VBA_RELATIONSHIP = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


# ---------------------------------------------------------------------------
# Independent readers


def decompress(data: bytes) -> bytes:
    """[MS-OVBA] 2.4.1.3.1: decompress a CompressedContainer."""
    if not data or data[0] != 0x01:
        raise ValueError("missing signature byte")
    out = bytearray()
    position = 1
    while position < len(data):
        start = position
        header = struct.unpack_from("<H", data, position)[0]
        position += 2
        if (header >> 12) & 0x7 != 0b011:
            raise ValueError("bad chunk signature")
        end = min(start + (header & 0x0FFF) + 3, len(data))
        if not header & 0x8000:
            out += data[position:position + 4096]
            position += 4096
            continue
        chunk_start = len(out)
        while position < end:
            flags = data[position]
            position += 1
            for bit in range(8):
                if position >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[position])
                    position += 1
                    continue
                token = struct.unpack_from("<H", data, position)[0]
                position += 2
                bit_count = max(4, (len(out) - chunk_start - 1).bit_length())
                length = (token & (0xFFFF >> bit_count)) + 3
                offset = (token >> (16 - bit_count)) + 1
                for _ in range(length):
                    out.append(out[-offset])
    return bytes(out)


def decrypt(hex_text: str) -> bytes:
    """[MS-OVBA] 2.4.3.3: decrypt a CMG, DPB or GC value."""
    raw = bytes.fromhex(hex_text)
    seed, version_encrypted, project_key_encrypted = raw[0], raw[1], raw[2]
    if seed ^ version_encrypted != 2:
        raise ValueError("bad version")
    unencrypted, encrypted1, encrypted2 = seed ^ project_key_encrypted, project_key_encrypted, version_encrypted
    plain = bytearray()
    for byte in raw[3:]:
        value = byte ^ ((encrypted2 + unencrypted) & 0xFF)
        encrypted2, encrypted1, unencrypted = encrypted1, byte, value
        plain.append(value)
    ignored = (seed & 6) // 2
    length = struct.unpack_from("<I", plain, ignored)[0]
    data = bytes(plain[ignored + 4:ignored + 4 + length])
    if len(data) != length:
        raise ValueError("truncated data")
    return data


class CompoundFile:
    """[MS-CFB] version 3 reader: FAT, directory tree and mini stream."""

    def __init__(self, blob: bytes) -> None:
        if blob[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ValueError("not a compound file")
        self.blob = blob
        self.sector = 1 << struct.unpack_from("<H", blob, 30)[0]
        self.mini = 1 << struct.unpack_from("<H", blob, 32)[0]
        fat_count, directory_start = struct.unpack_from("<II", blob, 44)
        self.cutoff, mini_fat_start = struct.unpack_from("<II", blob, 56)
        fat_sectors = struct.unpack_from("<109I", blob, 76)[:fat_count]
        fat = b"".join(self._sector(index) for index in fat_sectors)
        self.fat = list(struct.unpack(f"<{len(fat) // 4}I", fat))
        directory = self._chain_bytes(directory_start)
        self.entries = [self._entry(directory[index:index + 128]) for index in range(0, len(directory), 128)]
        root = self.entries[0]
        self.mini_stream = self._chain_bytes(root["start"])[:root["size"]] if root["size"] else b""
        mini_fat = self._chain_bytes(mini_fat_start) if mini_fat_start != ENDOFCHAIN else b""
        self.mini_fat = list(struct.unpack(f"<{len(mini_fat) // 4}I", mini_fat))

    def _sector(self, index: int) -> bytes:
        offset = self.sector * (index + 1)
        return self.blob[offset:offset + self.sector]

    @staticmethod
    def _chain(start: int, table: list[int]) -> list[int]:
        chain: list[int] = []
        while start != ENDOFCHAIN:
            if start in chain:
                raise ValueError("cyclic chain")
            chain.append(start)
            start = table[start]
        return chain

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self._sector(index) for index in self._chain(start, self.fat))

    @staticmethod
    def _entry(raw: bytes) -> dict[str, Any]:
        name_length = struct.unpack_from("<H", raw, 64)[0]
        return {"name": raw[:max(0, name_length - 2)].decode("utf-16-le"), "kind": raw[66],
                "left": struct.unpack_from("<I", raw, 68)[0], "right": struct.unpack_from("<I", raw, 72)[0],
                "child": struct.unpack_from("<I", raw, 76)[0], "start": struct.unpack_from("<I", raw, 116)[0],
                "size": struct.unpack_from("<Q", raw, 120)[0]}

    def children(self, sid: int) -> dict[str, int]:
        found: dict[str, int] = {}
        stack = [self.entries[sid]["child"]]
        while stack:
            node = stack.pop()
            if node == NOSTREAM:
                continue
            entry = self.entries[node]
            found[entry["name"]] = node
            stack += [entry["left"], entry["right"]]
        return found

    def read(self, path: str) -> bytes:
        sid = 0
        for part in path.split("/"):
            sid = self.children(sid)[part]
        entry = self.entries[sid]
        if entry["size"] >= self.cutoff:
            return self._chain_bytes(entry["start"])[:entry["size"]]
        if entry["size"] == 0:
            return b""
        chunks = [self.mini_stream[index * self.mini:(index + 1) * self.mini]
                  for index in self._chain(entry["start"], self.mini_fat)]
        return b"".join(chunks)[:entry["size"]]


def dir_records(data: bytes) -> list[tuple[int, bytes]]:
    """The records of a decompressed dir stream ([MS-OVBA] 2.3.4.2), in order."""
    records: list[tuple[int, bytes]] = []
    position = 0
    while position < len(data):
        identifier, size = struct.unpack_from("<HI", data, position)
        position += 6
        # PROJECTVERSION: its size field is a reserved 4, followed by 6 bytes.
        length = 6 if identifier == 0x0009 else size
        records.append((identifier, data[position:position + length]))
        position += length
    return records


def dir_modules(data: bytes) -> list[tuple[str, int]]:
    """Each module of a dir stream: its name and its MODULETYPE record id."""
    modules: list[tuple[str, int]] = []
    name = ""
    for identifier, payload in dir_records(data):
        if identifier == 0x0019:
            name = payload.decode("cp1252")
        elif identifier in (0x0021, 0x0022):
            modules.append((name, identifier))
    return modules


# ---------------------------------------------------------------------------
# Helpers


@contextlib.contextmanager
def temporary_folder() -> Iterator[Path]:
    with tempfile.TemporaryDirectory(prefix="build-workbook-") as name:
        yield Path(name)


def build_bytes(mode: str, root: Path = ROOT, override: Path | None = None) -> bytes:
    if root == ROOT and override is None:
        return repository_build(mode)
    return fresh_build(mode, root, override)


@functools.cache
def repository_build(mode: str) -> bytes:
    """The build of this checkout, made once per mode: a rebuild is compared with it."""
    return fresh_build(mode, ROOT, None)


def fresh_build(mode: str, root: Path, override: Path | None) -> bytes:
    with temporary_folder() as folder:
        target = folder / "book.xlsm"
        build_workbook.build(root, mode, target, override)
        return target.read_bytes()


def parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()}


def project(data: bytes) -> CompoundFile:
    return CompoundFile(parts(data)[build_workbook.VBA_PART])


def project_text(cfb: CompoundFile) -> list[str]:
    return cfb.read("PROJECT").decode("cp1252").split("\r\n")


def configuration(mode: str) -> list[str]:
    return [str(path) for path in MANIFEST["configurations"][mode.lower()]["modules"]]


def module_name(path: str) -> str:
    match = re.search(rb'^Attribute VB_Name = "([^"]+)"', (ROOT / path).read_bytes(), re.MULTILINE)
    assert match is not None
    return match.group(1).decode("ascii")


def crlf(data: bytes) -> bytes:
    text = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
    return text if text.endswith(b"\r\n") else text + b"\r\n"


@contextlib.contextmanager
def source_copy() -> Iterator[Path]:
    """A copy of the manifest and of src/, test/ and demo/, to change freely."""
    with temporary_folder() as root:
        (root / "tools").mkdir()
        shutil.copy(ROOT / build_workbook.MANIFEST, root / build_workbook.MANIFEST)
        for directory in build_workbook.SOURCE_DIRECTORIES:
            shutil.copytree(ROOT / directory, root / directory)
        yield root


def mutant(name: str, marker: str = "' mutant") -> bytes:
    return (f'Attribute VB_Name = "{name}"\r\nOption Explicit\r\n{marker}\r\n'
            "Public Sub Probe()\r\nEnd Sub\r\n").encode("ascii")


# ---------------------------------------------------------------------------
# vba_project.py


class Compression(unittest.TestCase):
    def test_the_specification_example_decompresses(self) -> None:
        # [MS-OVBA] 3.2.3, the "Normal Compression" example: checks the reader.
        compressed = bytes.fromhex("012FB000236161616263646582660070616768696A013808616B6C00306D6E6F70"
                                   "06710270041072737475761077" "78797A00" "3C")
        self.assertEqual(decompress(compressed), b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa")

    def test_round_trip(self) -> None:
        rng = random.Random(8)
        samples = [b"a", b"abc" * 3000, bytes(rng.randrange(256) for _ in range(9000)),
                   b"Option Explicit\r\n" * 700, bytes(range(256)) * 40, b"x" * 4096, b"x" * 8193]
        for data in samples:
            with self.subTest(length=len(data)):
                self.assertEqual(decompress(vba_project.compress(data)), data)

    def test_an_incompressible_full_chunk_is_stored_raw(self) -> None:
        rng = random.Random(3)
        data = bytes(rng.randrange(256) for _ in range(4096))
        compressed = vba_project.compress(data)
        self.assertEqual(struct.unpack_from("<H", compressed, 1)[0] & 0x8000, 0)
        self.assertEqual(decompress(compressed), data)

    def test_an_incompressible_short_last_chunk_names_the_module(self) -> None:
        rng = random.Random(97)
        noise = bytes(rng.randrange(33, 127) for _ in range(4000))
        module = vba_project.Module("M_NOISE", b'Attribute VB_Name = "M_NOISE"\r\n\' ' + noise + b"\r\n")
        with self.assertRaisesRegex(ValueError, "module M_NOISE: the last"):
            vba_project.build_vba_project([module], "{00000000-0000-0000-0000-000000000000}")


class ProjectStorage(unittest.TestCase):
    project_id = "{0B9A3C1D-1111-4222-8333-444455556666}"

    def modules(self) -> list[vba_project.Module]:
        body = b"\r\n".join(b"' line %d: some repeated text, repeated text" % index for index in range(400))
        return [vba_project.document_module("ThisWorkbook", True),
                vba_project.document_module("Sheet1", False),
                vba_project.Module("M_SMALL", b'Attribute VB_Name = "M_SMALL"\r\nOption Explicit\r\n'),
                vba_project.Module("M_LARGE", b'Attribute VB_Name = "M_LARGE"\r\n' + body + b"\r\n")]

    def test_every_stream_reads_back(self) -> None:
        cfb = CompoundFile(vba_project.build_vba_project(self.modules(), self.project_id))
        for module in self.modules():
            with self.subTest(module=module.name):
                self.assertEqual(decompress(cfb.read(f"VBA/{module.name}")), module.source)
        self.assertEqual(dir_modules(decompress(cfb.read("VBA/dir"))),
                         [("ThisWorkbook", 0x0022), ("Sheet1", 0x0022), ("M_SMALL", 0x0021), ("M_LARGE", 0x0021)])
        # Version 0xFFFF: the host compiles the project from source.
        self.assertEqual(cfb.read("VBA/_VBA_PROJECT")[:4], bytes.fromhex("CC61FFFF"))
        self.assertEqual(cfb.read("PROJECTwm")[-2:], b"\x00\x00")

    def test_the_dir_stream_declares_the_code_page_and_the_project_name(self) -> None:
        cfb = CompoundFile(vba_project.build_vba_project(self.modules(), self.project_id))
        records = dict(dir_records(decompress(cfb.read("VBA/dir"))))
        self.assertEqual(struct.unpack("<H", records[0x0003])[0], 1252)
        self.assertEqual(records[0x0004], b"VBAProject")
        self.assertEqual(struct.unpack("<H", records[0x000F])[0], 4)

    def test_the_protection_fields_decrypt_to_an_unprotected_project(self) -> None:
        cfb = CompoundFile(vba_project.build_vba_project(self.modules(), self.project_id))
        fields = dict(line.split("=", 1) for line in project_text(cfb) if line[:4] in ("CMG=", "DPB=", "GC=\""))
        self.assertEqual(decrypt(fields["CMG"].strip('"')), struct.pack("<I", 0))
        self.assertEqual(decrypt(fields["DPB"].strip('"')), b"\x00")
        self.assertEqual(decrypt(fields["GC"].strip('"')), b"\xff")

    def test_the_output_depends_only_on_the_modules_and_the_project_id(self) -> None:
        first = vba_project.build_vba_project(self.modules(), self.project_id)
        self.assertEqual(first, vba_project.build_vba_project(self.modules(), self.project_id))
        self.assertNotEqual(first, vba_project.build_vba_project(self.modules(), "{" + "1" * 8 + "-0000-0000-0000-"
                                                                 + "0" * 12 + "}"))

    def test_refusals(self) -> None:
        with self.assertRaisesRegex(ValueError, "unknown kind"):
            vba_project.Module("X", b"", kind="class")
        twice = [vba_project.Module("M_A", b"a\r\n"), vba_project.Module("m_a", b"b\r\n")]
        with self.assertRaisesRegex(ValueError, "same name"):
            vba_project.build_vba_project(twice, self.project_id)


# ---------------------------------------------------------------------------
# build_workbook.py


class Package(unittest.TestCase):
    data = build_bytes("Development")

    def test_the_package_declares_a_macro_enabled_workbook_and_its_vba_project(self) -> None:
        found = parts(self.data)
        types = ElementTree.fromstring(found["[Content_Types].xml"])
        defaults = {item.get("Extension"): item.get("ContentType") for item in types.iter(f"{CONTENT_TYPES}Default")}
        overrides = {item.get("PartName"): item.get("ContentType") for item in types.iter(f"{CONTENT_TYPES}Override")}
        self.assertEqual(defaults["bin"], "application/vnd.ms-office.vbaProject")
        self.assertEqual(overrides["/xl/workbook.xml"], "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        relationship = f"{PACKAGE_RELATIONSHIPS}Relationship"
        package_relationships = ElementTree.fromstring(found["_rels/.rels"])
        self.assertEqual([item.get("Target") for item in package_relationships.iter(relationship)],
                         ["xl/workbook.xml"])
        relationships = ElementTree.fromstring(found["xl/_rels/workbook.xml.rels"])
        targets = {item.get("Type"): item.get("Target") for item in relationships.iter(relationship)}
        self.assertEqual(targets[VBA_RELATIONSHIP], "vbaProject.bin")
        # Every part a relationship or an override names is in the package, and
        # every part but the relationships has a content type.
        for target in relationships.iter(relationship):
            self.assertIn("xl/" + str(target.get("Target")), found)
        for name in found:
            extension = name.rsplit(".", 1)[-1]
            self.assertTrue(f"/{name}" in overrides or extension in defaults, name)
        for part in overrides:
            self.assertIn(str(part).lstrip("/"), found)

    def test_every_part_is_stored_with_a_fixed_time(self) -> None:
        with zipfile.ZipFile(io.BytesIO(self.data)) as archive:
            for info in archive.infolist():
                with self.subTest(part=info.filename):
                    self.assertEqual(info.date_time, build_workbook.ZIP_TIME)
                    self.assertEqual(info.compress_type, zipfile.ZIP_STORED)
            self.assertIsNone(archive.testzip())

    def test_the_sheets_and_the_workbook_have_the_code_names_of_a_new_workbook(self) -> None:
        found = parts(self.data)
        book = ElementTree.fromstring(found["xl/workbook.xml"])
        self.assertEqual(book.find(f"{MAIN}workbookPr").get("codeName"), "ThisWorkbook")  # type: ignore[union-attr]
        self.assertEqual([sheet.get("name") for sheet in book.iter(f"{MAIN}sheet")], ["Sheet1"])
        sheet = ElementTree.fromstring(found["xl/worksheets/sheet1.xml"])
        self.assertEqual(sheet.find(f"{MAIN}sheetPr").get("codeName"), "Sheet1")  # type: ignore[union-attr]
        cfb = project(self.data)
        text = project_text(cfb)
        self.assertIn("Document=ThisWorkbook/&H00000000", text)
        self.assertIn("Document=Sheet1/&H00000000", text)
        for name, base in (("ThisWorkbook", "00020819"), ("Sheet1", "00020820")):
            with self.subTest(module=name):
                source = decompress(cfb.read(f"VBA/{name}")).decode("cp1252")
                self.assertIn(f'Attribute VB_Name = "{name}"', source)
                self.assertIn(f'Attribute VB_Base = "0{{{base}-', source)
                self.assertTrue(source.endswith("Option Explicit\r\n"))


class Modules(unittest.TestCase):
    def test_each_mode_holds_its_configuration_in_manifest_order(self) -> None:
        for mode in build_workbook.MODES:
            with self.subTest(mode=mode):
                cfb = project(build_bytes(mode))
                names = [module_name(path) for path in configuration(mode)]
                self.assertEqual(dir_modules(decompress(cfb.read("VBA/dir"))),
                                 [("ThisWorkbook", 0x0022), ("Sheet1", 0x0022)] + [(name, 0x0021) for name in names])
                self.assertEqual([line for line in project_text(cfb) if line.startswith("Module=")],
                                 [f"Module={name}" for name in names])

    def test_every_module_stream_round_trips_to_its_source(self) -> None:
        for mode in build_workbook.MODES:
            cfb = project(build_bytes(mode))
            for path in configuration(mode):
                with self.subTest(mode=mode, module=path):
                    self.assertEqual(decompress(cfb.read(f"VBA/{module_name(path)}")), crlf((ROOT / path).read_bytes()))

    def test_line_endings_do_not_change_the_workbook(self) -> None:
        with source_copy() as root:
            for path in configuration("Production"):
                (root / path).write_bytes((root / path).read_bytes().replace(b"\r\n", b"\n"))
            self.assertEqual(build_bytes("Production", root), build_bytes("Production"))

    def test_a_rebuild_is_byte_identical_and_a_source_change_is_not(self) -> None:
        first = build_bytes("Development")
        self.assertEqual(first, fresh_build("Development", ROOT, None))
        with source_copy() as root:
            path = root / "src" / "M_EXCEL_UI.bas"
            path.write_bytes(path.read_bytes() + b"' changed\r\n")
            changed = build_bytes("Development", root)
        self.assertNotEqual(first, changed)
        ids = [next(line for line in project_text(project(data)) if line.startswith("ID="))
               for data in (first, changed)]
        self.assertNotEqual(ids[0], ids[1])


class SourceRules(unittest.TestCase):
    def test_a_module_in_no_configuration_is_refused(self) -> None:
        with source_copy() as root:
            (root / "src" / "M_EXTRA.bas").write_bytes(mutant("M_EXTRA"))
            with self.assertRaisesRegex(build_workbook.BuildError, "not listed in any configuration.*M_EXTRA"):
                build_workbook.components(root, "Production")

    def test_a_listed_module_must_exist_and_be_a_standard_module(self) -> None:
        for entry, message in (("src/M_GONE.bas", "not found"), ("src/M_CLASS.cls", "not a standard module")):
            with self.subTest(entry=entry), source_copy() as root:
                manifest = json.loads((root / build_workbook.MANIFEST).read_text(encoding="utf-8"))
                manifest["configurations"]["production"]["modules"].append(entry)
                (root / build_workbook.MANIFEST).write_text(json.dumps(manifest), encoding="utf-8")
                with self.assertRaisesRegex(build_workbook.BuildError, message):
                    build_workbook.components(root, "Production")

    def test_a_vb_name_used_twice_is_refused(self) -> None:
        with source_copy() as root:
            (root / "demo" / "copy.bas").write_bytes(mutant("M_EXCEL_UI"))
            with self.assertRaisesRegex(build_workbook.BuildError, "duplicate VBA module name 'M_EXCEL_UI'"):
                build_workbook.components(root, "Production")

    def test_declared_name(self) -> None:
        cases = {
            b'Attribute VB_Name = "M_A"\r\n': ("M_A", ""),
            b'attribute vb_name = "M_A"\n': ("M_A", ""),
            b"Option Explicit\r\n": ("", "no Attribute VB_Name line"),
            b"Attribute VB_Name = M_A\r\n": ("", "malformed"),
            b'Attribute VB_Name = "M_A"\rAttribute VB_Name = "M_A"\r': ("", "several"),
            b'Attribute VB_Name = "1_A"\r\n': ("", "not a valid VBA name"),
        }
        for data, (name, problem) in cases.items():
            with self.subTest(data=data):
                found = build_workbook.declared_name(data)
                self.assertEqual(found[0], name)
                self.assertIn(problem, found[1])


class Override(unittest.TestCase):
    def setUp(self) -> None:
        folder = tempfile.TemporaryDirectory(prefix="override-")
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)

    def build(self, mode: str = "Development") -> bytes:
        return build_bytes(mode, override=self.folder)

    def test_a_module_is_replaced_by_its_vb_name_in_its_place(self) -> None:
        (self.folder / "anything.bas").write_bytes(mutant("M_EXCEL_UI"))
        cfb = project(self.build())
        self.assertEqual(decompress(cfb.read("VBA/M_EXCEL_UI")), mutant("M_EXCEL_UI"))
        names = [name for name, _ in dir_modules(decompress(cfb.read("VBA/dir")))]
        self.assertEqual(names[2:], [module_name(path) for path in configuration("Development")])
        snapshot = "src/M_EXCEL_UI_SNAPSHOT.bas"
        self.assertEqual(decompress(cfb.read("VBA/M_EXCEL_UI_SNAPSHOT")), crlf((ROOT / snapshot).read_bytes()))

    def test_a_new_module_named_like_the_repository_modules_is_added_last(self) -> None:
        (self.folder / "b.bas").write_bytes(mutant("M_NEW_B"))
        (self.folder / "a.bas").write_bytes(mutant("M_NEW_A"))
        cfb = project(self.build("Production"))
        names = [name for name, _ in dir_modules(decompress(cfb.read("VBA/dir")))]
        self.assertEqual(names[-2:], ["M_NEW_A", "M_NEW_B"])

    def test_the_newest_copy_of_a_module_wins(self) -> None:
        older, newer = self.folder / "M_EXCEL_UI.bas", self.folder / "M_EXCEL_UI (1).bas"
        older.write_bytes(mutant("M_EXCEL_UI", "' older"))
        newer.write_bytes(mutant("M_EXCEL_UI", "' newer"))
        os.utime(older, (1_000_000_000, 1_000_000_000))
        os.utime(newer, (1_000_000_100, 1_000_000_100))
        with temporary_folder() as output:
            result = build_workbook.build(ROOT, "Production", output / "book.xlsm", self.folder)
            source = decompress(project((output / "book.xlsm").read_bytes()).read("VBA/M_EXCEL_UI"))
        self.assertIn(b"' newer", source)
        self.assertEqual(result.ignored, ("M_EXCEL_UI.bas (older copy of 'M_EXCEL_UI')",))
        self.assertEqual([part.origin for part in result.components if part.name == "M_EXCEL_UI"], ["override"])

    def test_an_override_the_synchronizer_would_ignore_is_refused(self) -> None:
        cases = {
            "unknown module": ("x.bas", mutant("Foo"), "'Foo' is not a repository module"),
            "no declaration": ("M_EXCEL_UI.bas", b"Option Explicit\r\n", "no Attribute VB_Name line"),
            "two declarations": ("x.bas", mutant("M_EXCEL_UI") + mutant("M_EXCEL_UI"), "several"),
            "another kind": ("x.cls", mutant("M_EXCEL_UI"), "is a .bas module in the repository"),
            "a new class": ("x.cls", mutant("M_NEW"), "not a repository module"),
            "another configuration": ("x.bas", mutant("M_EXCEL_UI_DEMO"), "not a module of the Development"),
        }
        for label, (file_name, data, message) in cases.items():
            with self.subTest(case=label):
                for path in self.folder.iterdir():
                    path.unlink()
                (self.folder / file_name).write_bytes(data)
                with self.assertRaisesRegex(build_workbook.BuildError, message):
                    self.build()

    def test_an_empty_or_missing_folder_is_refused(self) -> None:
        (self.folder / "notes.txt").write_text("not a module", encoding="utf-8")
        with self.assertRaisesRegex(build_workbook.BuildError, "holds no module"):
            self.build()
        with self.assertRaisesRegex(build_workbook.BuildError, "not a folder"):
            build_bytes("Development", override=self.folder / "missing")


class CommandLine(unittest.TestCase):
    def run_main(self, *argv: str) -> tuple[int, str, str]:
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            try:
                code = build_workbook.main(list(argv))
            except SystemExit as exit_:
                code = int(exit_.code or 0)
        return code, out.getvalue(), err.getvalue()

    def test_a_build_writes_the_workbook_and_names_its_modules(self) -> None:
        with temporary_folder() as folder:
            target = folder / "nested" / "book.xlsm"
            code, out, _ = self.run_main("--mode", "production", str(target))
            self.assertEqual(code, 0)
            self.assertEqual(target.read_bytes(), build_bytes("Production"))
            self.assertEqual(list(target.parent.iterdir()), [target])
        self.assertIn("mode: Production (4 standard modules)", out)
        self.assertRegex(out, r"sha256: [0-9a-f]{64}")

    def test_an_unknown_mode_or_a_bad_output_is_a_usage_error(self) -> None:
        for argv in (["--mode", "Staging", "book.xlsm"], ["book.xlsm"], ["--mode", "Demo", "book.xlsx"]):
            with self.subTest(argv=argv):
                code, _, err = self.run_main(*argv)
                self.assertEqual(code, 2)
                self.assertIn("usage:", err)
        with self.assertRaisesRegex(build_workbook.BuildError, "unknown mode"):
            build_workbook.build(ROOT, "Staging", Path("book.xlsm"))

    def test_an_unusable_override_fails_and_writes_nothing(self) -> None:
        with temporary_folder() as folder:
            (folder / "x.bas").write_bytes(mutant("Foo"))
            target = folder / "book.xlsm"
            code, _, err = self.run_main("--mode", "Development", "--override", str(folder), str(target))
            self.assertEqual(code, 1)
            self.assertIn("FAIL: override x.bas", err)
            self.assertFalse(target.exists())


if __name__ == "__main__":
    unittest.main()
