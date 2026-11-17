# LinBASIC Studio - the resource file (.res) of a VB6 project
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""A 32-bit .res file is a list of resources, each with a header:
   u32 DataSize, u32 HeaderSize, TYPE, NAME (0xFFFF + u16 number, or a UTF-16 name ending in 0),
   (align 4) u32 DataVersion, u16 MemoryFlags, u16 LanguageId, u32 Version, u32 Characteristics,
   then the data (align 4).  The first entry of the file is empty."""
import struct


def _id(b, p):
    if b[p:p + 2] == b"\xff\xff":
        return struct.unpack_from("<H", b, p + 2)[0], p + 4
    end = p
    while b[end:end + 2] != b"\0\0":
        end += 2
    return b[p:end].decode("utf-16-le").upper(), end + 2


def read_res(path):
    """[(type, name, data)]: type and name are an int, or an upper-case str"""
    with open(path, "rb") as f:
        b = f.read()
    out, p = [], 0
    while p + 8 <= len(b):
        size, hsize = struct.unpack_from("<II", b, p)
        if hsize < 16:
            break
        t, q = _id(b, p + 8)
        n, q = _id(b, q)
        data = b[p + hsize:p + hsize + size]
        if size or t != 0:
            out.append((t, n, data))
        p += hsize + size
        p = (p + 3) & ~3
    return out
