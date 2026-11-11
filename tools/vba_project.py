"""Write an Office VBA project storage (vbaProject.bin) from module source.

tools/build_workbook.py embeds the result in the workbook it builds. The
writer follows two Microsoft specifications:

- [MS-OVBA]: the compression of the module and dir streams (2.4.1), the dir
  stream (2.3.4.2), the PROJECT and PROJECTwm streams (2.3.1 and 2.3.3), and
  the data encryption of the CMG, DPB and GC fields of PROJECT (2.4.3);
- [MS-CFB]: a version 3 compound file, with 512-byte sectors and a mini
  stream for the streams under 4096 bytes.

Each module stream holds the module's source only, with no compiled code, and
_VBA_PROJECT declares version 0xFFFF, so the host compiles the project from
source when it opens the workbook. Standard modules and document modules
(ThisWorkbook and the sheets) are written; class modules and forms are not.

The output depends only on the modules and the project ID: the protection
fields of PROJECT are encrypted with a random generator seeded with the
project ID. tests/test_build_workbook.py reads the output back with a reader
of its own. Standard library only.
"""

from __future__ import annotations

import random
import struct
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from typing import TypeAlias

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF = 4096
HEADER_DIFAT_ENTRIES = 109
CHUNK_SIZE = 4096

PROJECT_NAME = "VBAProject"
CODE_PAGE = 1252
ENCODING = "cp1252"
LCID_ENGLISH_US = 0x0409

# The VB_Base of a document module: Excel's Workbook and Worksheet classes.
WORKBOOK_BASE = "0{00020819-0000-0000-C000-000000000046}"
WORKSHEET_BASE = "0{00020820-0000-0000-C000-000000000046}"

# The references of a VBA project that Excel creates: OLE Automation and the
# Office object library. The host resolves each by its GUID; the path is a hint.
REFERENCES = (
    ("stdole", r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"),
    ("Office", r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#"
               r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"),
)

# _VBA_PROJECT: reserved 0x61CC, version 0xFFFF (compile from source), then
# three reserved bytes and no performance cache.
VBA_PROJECT_STREAM = bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])

STANDARD = "standard"
DOCUMENT = "document"


@dataclass(frozen=True)
class Module:
    """One module of the project: its name and its source as the VBE stores it,
    in the project code page with CRLF line endings and its Attribute lines."""

    name: str
    source: bytes
    kind: str = STANDARD

    def __post_init__(self) -> None:
        if self.kind not in (STANDARD, DOCUMENT):
            raise ValueError(f"module {self.name}: unknown kind {self.kind!r}")


def document_module(name: str, workbook: bool) -> Module:
    """The empty document module of the workbook or of one worksheet, with the
    attributes Excel gives it and Option Explicit."""
    base = WORKBOOK_BASE if workbook else WORKSHEET_BASE
    lines = [
        f'Attribute VB_Name = "{name}"',
        f'Attribute VB_Base = "{base}"',
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = True",
        "Attribute VB_TemplateDerived = False",
        "Attribute VB_Customizable = True",
        "Option Explicit",
        "",
    ]
    return Module(name, "\r\n".join(lines).encode(ENCODING), DOCUMENT)


# ---------------------------------------------------------------------------
# Compression ([MS-OVBA] 2.4.1)


def copy_token_layout(decompressed: int) -> tuple[int, int]:
    """The bit count of a copy token's offset, and the longest copy it can
    express, at ``decompressed`` bytes into the current chunk (2.4.1.3.19.1)."""
    bit_count = 4
    while (1 << bit_count) < decompressed:
        bit_count += 1
    return bit_count, (0xFFFF >> bit_count) + 3


def compress(data: bytes) -> bytes:
    """A CompressedContainer holding ``data``: a signature byte, then one
    chunk per 4096 bytes."""
    out = bytearray([0x01])
    for start in range(0, len(data), CHUNK_SIZE):
        out += compress_chunk(data[start:start + CHUNK_SIZE])
    return bytes(out)


def longest_match(chunk: bytes, position: int, candidates: Sequence[int]) -> tuple[int, int]:
    """The longest earlier copy of the bytes at ``position`` that a copy token
    can express: its length and its offset, or (0, 0)."""
    bit_count, max_length = copy_token_layout(position)
    limit = min(max_length, len(chunk) - position)
    best_length, best_offset = 0, 0
    # The most recent 256 occurrences are enough for source text, and keep a
    # long run of one repeated line from making the search quadratic.
    for candidate in reversed(candidates[-256:]):
        offset = position - candidate
        if offset > (1 << bit_count):
            break
        length = 0
        while length < limit and chunk[candidate + length] == chunk[position + length]:
            length += 1
        if length > best_length:
            best_length, best_offset = length, offset
            if length == limit:
                break
    return best_length, best_offset


def compress_chunk(chunk: bytes) -> bytes:
    """One CompressedChunk: a header, then token sequences of a flag byte and
    up to eight literal bytes or copy tokens (2.4.1.3.7)."""
    body = bytearray()
    position = 0
    seen: dict[bytes, list[int]] = {}

    def remember(index: int) -> None:
        if index + 3 <= len(chunk):
            seen.setdefault(chunk[index:index + 3], []).append(index)

    while position < len(chunk):
        flag_index = len(body)
        body.append(0)
        for bit in range(8):
            if position >= len(chunk):
                break
            length, offset = longest_match(chunk, position, seen.get(chunk[position:position + 3], [])) \
                if position else (0, 0)
            if length >= 3:
                bit_count, _ = copy_token_layout(position)
                body += struct.pack("<H", ((offset - 1) << (16 - bit_count)) | (length - 3))
                body[flag_index] |= 1 << bit
            else:
                length = 1
                body.append(chunk[position])
            for index in range(position, position + length):
                remember(index)
            position += length
    if len(body) + 2 > CHUNK_SIZE + 2:
        if len(chunk) != CHUNK_SIZE:
            # A chunk is stored raw only with exactly 4096 bytes, and only the
            # last chunk may be shorter, so a short last chunk that does not
            # compress cannot be written.
            raise ValueError(f"the last {len(chunk)} bytes do not compress into one chunk")
        return struct.pack("<H", 0x0FFF | (0b011 << 12)) + chunk
    header = ((len(body) + 2 - 3) & 0x0FFF) | (0b011 << 12) | (1 << 15)
    return struct.pack("<H", header) + bytes(body)


# ---------------------------------------------------------------------------
# Data encryption of the PROJECT protection fields ([MS-OVBA] 2.4.3)


def encrypt(data: bytes, project_id: str, rng: random.Random) -> str:
    """``data`` encrypted as 2.4.3.2 describes, as upper-case hexadecimal."""
    seed = rng.randrange(256)
    version = 2
    project_key = sum(project_id.encode(ENCODING)) & 0xFF
    out = bytearray([seed, seed ^ version, seed ^ project_key])
    unencrypted, encrypted1, encrypted2 = project_key, seed ^ project_key, seed ^ version

    def emit(byte: int) -> None:
        nonlocal unencrypted, encrypted1, encrypted2
        value = byte ^ ((encrypted2 + unencrypted) & 0xFF)
        out.append(value)
        encrypted2, encrypted1, unencrypted = encrypted1, value, byte

    for _ in range((seed & 6) // 2):
        emit(rng.randrange(256))
    for byte in struct.pack("<I", len(data)) + data:
        emit(byte)
    return out.hex().upper()


# ---------------------------------------------------------------------------
# Project streams ([MS-OVBA] 2.3)


def record(identifier: int, payload: bytes) -> bytes:
    """A dir stream record: its id, the size of its payload, the payload."""
    return struct.pack("<HI", identifier, len(payload)) + payload


def unicode(text: str) -> bytes:
    return text.encode("utf-16-le")


def mbcs(text: str) -> bytes:
    return text.encode(ENCODING)


def information_records() -> bytes:
    """PROJECTINFORMATION (2.3.4.2.1)."""
    return b"".join((
        record(0x0001, struct.pack("<I", 1)),                      # SYSKIND: 32-bit Windows
        record(0x0002, struct.pack("<I", LCID_ENGLISH_US)),        # LCID
        record(0x0014, struct.pack("<I", LCID_ENGLISH_US)),        # LCIDINVOKE
        record(0x0003, struct.pack("<H", CODE_PAGE)),              # CODEPAGE
        record(0x0004, mbcs(PROJECT_NAME)),                        # NAME
        record(0x0005, b"") + record(0x0040, b""),                 # DOCSTRING, and in Unicode
        record(0x0006, b"") + record(0x003D, b""),                 # HELPFILEPATH 1 and 2
        record(0x0007, struct.pack("<I", 0)),                      # HELPCONTEXT
        record(0x0008, struct.pack("<I", 0)),                      # LIBFLAGS
        struct.pack("<HIIH", 0x0009, 4, 1, 0),                     # VERSION: reserved 4, major, minor
        record(0x000C, b"") + record(0x003C, b""),                 # CONSTANTS, and in Unicode
    ))


def reference_records() -> bytes:
    """PROJECTREFERENCES (2.3.4.2.2): one registered reference per library."""
    out = bytearray()
    for name, libid in REFERENCES:
        out += record(0x0016, mbcs(name)) + record(0x003E, unicode(name))
        encoded = mbcs(libid)
        out += record(0x000D, struct.pack("<I", len(encoded)) + encoded + struct.pack("<IH", 0, 0))
    return bytes(out)


def module_records(module: Module) -> bytes:
    """MODULE (2.3.4.2.3.2): the module's name, stream, and type."""
    return b"".join((
        record(0x0019, mbcs(module.name)) + record(0x0047, unicode(module.name)),   # NAME
        record(0x001A, mbcs(module.name)) + record(0x0032, unicode(module.name)),   # STREAMNAME
        record(0x001C, b"") + record(0x0048, b""),                                  # DOCSTRING
        record(0x0031, struct.pack("<I", 0)),                                       # OFFSET: source at 0
        record(0x001E, struct.pack("<I", 0)),                                       # HELPCONTEXT
        record(0x002C, struct.pack("<H", 0xFFFF)),                                  # COOKIE
        record(0x0021 if module.kind == STANDARD else 0x0022, b""),                 # TYPE
        record(0x002B, b""),                                                        # terminator
    ))


def dir_stream(modules: Sequence[Module]) -> bytes:
    """The uncompressed dir stream (2.3.4.2)."""
    return b"".join((
        information_records(),
        reference_records(),
        record(0x000F, struct.pack("<H", len(modules))),   # PROJECTMODULES: count
        record(0x0013, struct.pack("<H", 0xFFFF)),         # PROJECTCOOKIE
        *(module_records(module) for module in modules),
        record(0x0010, b""),                               # terminator
    ))


def project_stream(project_id: str, modules: Sequence[Module]) -> bytes:
    """The PROJECT stream (2.3.1): the project's properties as text."""
    rng = random.Random(project_id)
    lines = [f'ID="{project_id}"']
    lines += [f"Document={module.name}/&H00000000" if module.kind == DOCUMENT else f"Module={module.name}"
              for module in modules]
    lines += [
        f'Name="{PROJECT_NAME}"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'CMG="{encrypt(struct.pack("<I", 0), project_id, rng)}"',
        f'DPB="{encrypt(bytes([0]), project_id, rng)}"',
        f'GC="{encrypt(bytes([0xFF]), project_id, rng)}"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    lines += [f"{module.name}=0, 0, 0, 0, C" for module in modules]
    return ("\r\n".join(lines) + "\r\n").encode(ENCODING)


def project_wm_stream(modules: Sequence[Module]) -> bytes:
    """The PROJECTwm stream (2.3.3): each module name in the code page and in Unicode."""
    out = bytearray()
    for module in modules:
        out += mbcs(module.name) + b"\x00" + unicode(module.name) + b"\x00\x00"
    return bytes(out + b"\x00\x00")


# ---------------------------------------------------------------------------
# Compound file ([MS-CFB])

# A storage: each child by name, a stream as its bytes, a storage as a Tree.
Tree: TypeAlias = Mapping[str, "bytes | Tree"]

STORAGE, STREAM, ROOT = 1, 2, 5


@dataclass
class Entry:
    name: str
    kind: int
    data: bytes = b""
    children: list[Entry] = field(default_factory=list)
    sid: int = 0
    left: int = NOSTREAM
    right: int = NOSTREAM
    child: int = NOSTREAM
    start: int = ENDOFCHAIN
    size: int = 0


def directory_entries(tree: Tree) -> list[Entry]:
    """Every entry of the compound file, root first, with its stream id and
    its children linked as a binary search tree in compound-file order."""
    root = Entry("Root Entry", ROOT)
    entries = [root]

    def add(parent: Entry, node: Tree) -> None:
        for name, value in node.items():
            entry = Entry(name, STREAM, value, size=len(value)) if isinstance(value, bytes) \
                else Entry(name, STORAGE)
            entry.sid = len(entries)
            parent.children.append(entry)
            entries.append(entry)
            if not isinstance(value, bytes):
                add(entry, value)

    def link(siblings: list[Entry]) -> int:
        # Names compare by length first, then by their upper-case form.
        ordered = sorted(siblings, key=lambda entry: (len(entry.name), entry.name.upper()))
        if not ordered:
            return NOSTREAM
        middle = len(ordered) // 2
        node = ordered[middle]
        node.left = link(ordered[:middle])
        node.right = link(ordered[middle + 1:])
        return node.sid

    add(root, tree)
    for entry in entries:
        entry.child = link(entry.children)
    return entries


class SectorWriter:
    """The sectors of a compound file after its header, and their FAT."""

    def __init__(self) -> None:
        self.sectors: list[bytes] = []
        self.fat: list[int] = []

    def chain(self, data: bytes) -> int:
        """Store ``data`` in a chain of new sectors and return the first."""
        count = max(1, -(-len(data) // SECTOR_SIZE))
        first = len(self.sectors)
        for index in range(count):
            self.sectors.append(data[index * SECTOR_SIZE:(index + 1) * SECTOR_SIZE].ljust(SECTOR_SIZE, b"\x00"))
            self.fat.append(first + index + 1 if index < count - 1 else ENDOFCHAIN)
        return first

    def close(self) -> list[int]:
        """Append the FAT sectors, which describe themselves too, and return their numbers."""
        per_sector = SECTOR_SIZE // 4
        count = 1
        while count * per_sector < len(self.sectors) + count:
            count += 1
        if count > HEADER_DIFAT_ENTRIES:
            raise ValueError("the VBA project is too large for a compound file without DIFAT sectors")
        first = len(self.sectors)
        table = self.fat + [FATSECT] * count
        table += [FREESECT] * (count * per_sector - len(table))
        data = struct.pack(f"<{len(table)}I", *table)
        self.sectors += [data[index * SECTOR_SIZE:(index + 1) * SECTOR_SIZE] for index in range(count)]
        return list(range(first, first + count))


def place_streams(entries: list[Entry], writer: SectorWriter) -> tuple[int, int]:
    """Store every stream: those under the cutoff in the mini stream, the others
    in sectors. Return the first sector and the sector count of the mini FAT."""
    mini = bytearray()
    mini_fat: list[int] = []
    for entry in entries:
        if entry.kind != STREAM or entry.size >= MINI_STREAM_CUTOFF:
            continue
        if entry.size == 0:
            continue
        count = -(-entry.size // MINI_SECTOR_SIZE)
        entry.start = len(mini_fat)
        mini_fat += [entry.start + index + 1 if index < count - 1 else ENDOFCHAIN for index in range(count)]
        mini += entry.data.ljust(count * MINI_SECTOR_SIZE, b"\x00")
    for entry in entries:
        if entry.kind == STREAM and entry.size >= MINI_STREAM_CUTOFF:
            entry.start = writer.chain(entry.data)
    if not mini:
        return ENDOFCHAIN, 0
    entries[0].start = writer.chain(bytes(mini))
    entries[0].size = len(mini)
    table = struct.pack(f"<{len(mini_fat)}I", *mini_fat)
    count = -(-len(table) // SECTOR_SIZE)
    return writer.chain(table.ljust(count * SECTOR_SIZE, b"\xff")), count


def directory_entry(entry: Entry) -> bytes:
    """One 128-byte directory entry; every node is black."""
    name = entry.name.encode("utf-16-le") + b"\x00\x00"
    is_storage = entry.kind == STORAGE
    return (name.ljust(64, b"\x00")
            + struct.pack("<HBB", len(name), entry.kind, 1)
            + struct.pack("<III", entry.left, entry.right, entry.child)
            + bytes(16) + struct.pack("<I", 0) + bytes(16)        # CLSID, state bits, times
            + struct.pack("<IQ", 0 if is_storage else entry.start, 0 if is_storage else entry.size))


UNUSED_ENTRY = bytes(64) + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) \
    + bytes(36) + struct.pack("<IQ", 0, 0)


def write_compound_file(tree: Tree) -> bytes:
    """A version 3 compound file holding ``tree``."""
    entries = directory_entries(tree)
    writer = SectorWriter()
    mini_fat_start, mini_fat_count = place_streams(entries, writer)
    directory = b"".join(directory_entry(entry) for entry in entries)
    while len(directory) % SECTOR_SIZE:
        directory += UNUSED_ENTRY
    directory_start = writer.chain(directory)
    fat_sectors = writer.close()
    difat = fat_sectors + [FREESECT] * (HEADER_DIFAT_ENTRIES - len(fat_sectors))
    header = (bytes.fromhex("D0CF11E0A1B11AE1") + bytes(16)
              + struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + bytes(6)
              + struct.pack("<IIIII", 0, len(fat_sectors), directory_start, 0, MINI_STREAM_CUTOFF)
              + struct.pack("<IIII", mini_fat_start, mini_fat_count, ENDOFCHAIN, 0)
              + struct.pack(f"<{HEADER_DIFAT_ENTRIES}I", *difat))
    return header + b"".join(writer.sectors)


# ---------------------------------------------------------------------------


def build_vba_project(modules: Sequence[Module], project_id: str) -> bytes:
    """The vbaProject.bin of ``modules``, in that order, under ``project_id``
    (a GUID in braces)."""
    names = [module.name.casefold() for module in modules]
    if len(set(names)) != len(names):
        raise ValueError("two modules have the same name")
    streams: dict[str, bytes | Tree] = {"_VBA_PROJECT": VBA_PROJECT_STREAM, "dir": compress(dir_stream(modules))}
    for module in modules:
        try:
            streams[module.name] = compress(module.source)
        except ValueError as error:
            raise ValueError(f"module {module.name}: {error}") from error
    return write_compound_file({
        "PROJECT": project_stream(project_id, modules),
        "PROJECTwm": project_wm_stream(modules),
        "VBA": streams,
    })
