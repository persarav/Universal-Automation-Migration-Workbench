"""Pure-Python VBA project extractor for XLSM/XLAM files.
Uses the Compound File Binary Format and MS-OVBA decompression. Never executes macros.
"""
from __future__ import annotations
import struct, zipfile, io

END = 0xFFFFFFFE
FREE = 0xFFFFFFFF


def _decompress_vba(data: bytes) -> bytes:
    if not data or data[0] != 1:
        raise ValueError("Invalid MS-OVBA compressed container signature")
    pos = 1
    out = bytearray()
    while pos + 2 <= len(data):
        header = int.from_bytes(data[pos:pos+2], "little")
        pos += 2
        signature = (header >> 12) & 0x7
        compressed = (header >> 15) & 1
        chunk_size = (header & 0x0FFF) + 3
        if signature != 0x3:
            raise ValueError(f"Invalid MS-OVBA chunk signature: {signature}")
        chunk_end = min(pos + chunk_size - 2, len(data))
        chunk = data[pos:chunk_end]
        pos = chunk_end
        start = len(out)
        if not compressed:
            out.extend(chunk[:4096])
            continue
        cp = 0
        while cp < len(chunk) and len(out) - start < 4096:
            flags = chunk[cp]
            cp += 1
            for bit in range(8):
                if cp >= len(chunk) or len(out) - start >= 4096:
                    break
                if flags & (1 << bit):
                    if cp + 2 > len(chunk):
                        break
                    token = int.from_bytes(chunk[cp:cp+2], "little")
                    cp += 2
                    current = len(out) - start
                    bit_count = max(4, (current - 1).bit_length())
                    length_mask = 0xFFFF >> bit_count
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    if offset > current:
                        raise ValueError("Invalid MS-OVBA copy-token offset")
                    for _ in range(length):
                        out.append(out[-offset])
                        if len(out) - start >= 4096:
                            break
                else:
                    out.append(chunk[cp])
                    cp += 1
        if pos >= len(data):
            break
    return bytes(out)


def _cfb_streams(blob: bytes) -> dict[str, bytes]:
    if blob[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
        raise ValueError("Not an OLE Compound File")
    sector_size = 1 << struct.unpack_from("<H", blob, 0x1E)[0]
    mini_sector_size = 1 << struct.unpack_from("<H", blob, 0x20)[0]
    num_fat = struct.unpack_from("<I", blob, 0x2C)[0]
    dir_start = struct.unpack_from("<I", blob, 0x30)[0]
    mini_cutoff = struct.unpack_from("<I", blob, 0x38)[0]
    mini_fat_start = struct.unpack_from("<I", blob, 0x3C)[0]
    num_mini_fat = struct.unpack_from("<I", blob, 0x40)[0]
    difat_start = struct.unpack_from("<I", blob, 0x44)[0]
    num_difat = struct.unpack_from("<I", blob, 0x48)[0]

    def sector(sid: int) -> bytes:
        start = (sid + 1) * sector_size
        return blob[start:start + sector_size]

    difat = [x for x in struct.unpack_from("<109I", blob, 0x4C) if x not in (FREE, END)]
    sid = difat_start
    for _ in range(num_difat):
        if sid in (FREE, END):
            break
        values = struct.unpack("<%dI" % (sector_size // 4), sector(sid))
        difat.extend(x for x in values[:-1] if x not in (FREE, END))
        sid = values[-1]
    difat = difat[:num_fat]
    fat: list[int] = []
    for fat_sid in difat:
        fat.extend(struct.unpack("<%dI" % (sector_size // 4), sector(fat_sid)))

    def chain(start: int, table: list[int], limit: int = 100000) -> list[int]:
        result, seen, cur = [], set(), start
        while cur not in (END, FREE) and cur < len(table) and cur not in seen and len(result) < limit:
            seen.add(cur)
            result.append(cur)
            cur = table[cur]
        return result

    def read_fat(start: int, size: int | None = None, table: list[int] | None = None) -> bytes:
        tab = table if table is not None else fat
        raw = b"".join(sector(i) for i in chain(start, tab))
        return raw if size is None else raw[:size]

    directory_data = read_fat(dir_start)
    entries = []
    for off in range(0, len(directory_data), 128):
        entry = directory_data[off:off+128]
        if len(entry) < 128:
            break
        name_len = struct.unpack_from("<H", entry, 64)[0]
        if name_len < 2:
            continue
        name = entry[:name_len-2].decode("utf-16le", "replace")
        entry_type = entry[66]
        start = struct.unpack_from("<I", entry, 116)[0]
        size = struct.unpack_from("<Q", entry, 120)[0]
        entries.append((name, entry_type, start, size))

    root = next((e for e in entries if e[1] == 5), None)
    root_data = read_fat(root[2], root[3]) if root else b""
    mini_fat_raw = read_fat(mini_fat_start, num_mini_fat * sector_size) if num_mini_fat else b""
    mini_fat = list(struct.unpack("<%dI" % (len(mini_fat_raw)//4), mini_fat_raw)) if mini_fat_raw else []

    def read_mini(start: int, size: int) -> bytes:
        chunks, seen, cur = [], set(), start
        while cur not in (END, FREE) and cur < len(mini_fat) and cur not in seen:
            seen.add(cur)
            off = cur * mini_sector_size
            chunks.append(root_data[off:off+mini_sector_size])
            cur = mini_fat[cur]
        return b"".join(chunks)[:size]

    streams = {}
    for name, entry_type, start, size in entries:
        if entry_type == 2:
            streams[name] = read_mini(start, size) if size < mini_cutoff else read_fat(start, size)
    return streams


def _module_offsets(dir_data: bytes) -> dict[str, tuple[str, int]]:
    """Read module name, stream name and MODULEOFFSET records from decompressed dir stream."""
    result = {}
    tag_name = b"\x19\x00"       # MODULENAME
    tag_unicode_name = b"\x47\x00" # MODULENAMEUNICODE
    tag_stream = b"\x1a\x00"      # MODULESTREAMNAME
    tag_unicode_stream = b"\x32\x00"
    tag_offset = b"\x31\x00\x04\x00\x00\x00" # MODULEOFFSET + cb=4
    pos = 0
    while True:
        pos = dir_data.find(tag_name, pos)
        if pos < 0:
            break
        try:
            name_len = int.from_bytes(dir_data[pos+2:pos+6], "little")
            module_name = dir_data[pos+6:pos+6+name_len].decode("latin1", "replace")
            cur = pos + 6 + name_len
            # Optional unicode module name record.
            if dir_data[cur:cur+2] == tag_unicode_name:
                n = int.from_bytes(dir_data[cur+2:cur+6], "little")
                cur += 6 + n
            # Stream name record (skip unicode byte sequences by following declared record lengths).
            if dir_data[cur:cur+2] != tag_stream:
                pos += 2
                continue
            stream_len = int.from_bytes(dir_data[cur+2:cur+6], "little")
            stream_name = dir_data[cur+6:cur+6+stream_len].decode("latin1", "replace")
            cur += 6 + stream_len
            if dir_data[cur:cur+2] == tag_unicode_stream:
                n = int.from_bytes(dir_data[cur+2:cur+6], "little")
                cur += 6 + n
            off_pos = dir_data.find(tag_offset, cur, min(cur+128, len(dir_data)))
            if off_pos >= 0:
                offset = int.from_bytes(dir_data[off_pos+6:off_pos+10], "little")
                result[module_name] = (stream_name, offset)
        except Exception:
            pass
        pos += 2
    return result


def extract_vba_modules(raw_xlsm: bytes) -> tuple[list[dict], list[str]]:
    """Extract standard VBA modules from a macro-enabled workbook without executing them."""
    notes = []
    try:
        with zipfile.ZipFile(io.BytesIO(raw_xlsm)) as zf:
            names = set(zf.namelist())
            vba_path = next((n for n in names if n.lower().endswith("vbaproject.bin")), None)
            if not vba_path:
                return [], ["No xl/vbaProject.bin was found in this workbook."]
            streams = _cfb_streams(zf.read(vba_path))
    except Exception as exc:
        return [], [f"Could not read VBA project container: {type(exc).__name__}: {exc}"]

    if "dir" not in streams:
        return [], ["VBA project has no dir stream; modules could not be enumerated."]
    try:
        directory = _decompress_vba(streams["dir"])
        module_map = _module_offsets(directory)
    except Exception as exc:
        return [], [f"Could not parse VBA module directory: {type(exc).__name__}: {exc}"]

    modules = []
    for module_name, (stream_name, offset) in module_map.items():
        stream = streams.get(stream_name)
        if stream is None:
            notes.append(f"Module stream missing: {stream_name}")
            continue
        if offset >= len(stream):
            notes.append(f"Invalid source offset for module {module_name}")
            continue
        try:
            code = _decompress_vba(stream[offset:])
            text = code.decode("cp1252", "replace")
            # Strip non-source bytes only if a stream produced empty content.
            modules.append({"filename": f"{module_name}.bas", "text": text, "language": "VBA module"})
        except Exception as exc:
            notes.append(f"Could not decompress module {module_name}: {type(exc).__name__}: {exc}")
    if modules:
        notes.append(f"Extracted {len(modules)} VBA module(s) directly from the OLE project; macros were not executed.")
    return modules, notes
