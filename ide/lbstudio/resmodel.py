# LinBASIC Studio - the resource file (.res) of a project, as the VB 6 Resource Editor reads and writes it
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""The model of the Resource Editor (phase 8.6), without Qt. What VB6 writes (VB-TESTS/DEV-8/rs1):

  - a 32-bit .res: an empty first entry (32 bytes), then the resources, each with a header (u32 DataSize,
    u32 HeaderSize, TYPE, NAME, align 4, u32 DataVersion 0, u16 MemoryFlags, u16 LanguageId, u32 Version 0,
    u32 Characteristics 0) and its data, aligned to 4;
  - the order: the string tables (by language column, then by block), the cursors (RT_CURSOR images, then
    their RT_GROUP_CURSOR), the icons (RT_ICON images, then their RT_GROUP_ICON), the bitmaps, the rest;
  - MemoryFlags 0 for the string tables, 0x1030 for the others; the images of the icons and the cursors are
    numbered 1, 2, ... in that order;
  - a string table: blocks of 16 strings (block (id >> 4) + 1, u16 length + UTF-16); an empty cell is not
    written, a block without strings neither;
  - a bitmap: the .bmp without its 14-byte file header; an icon: RT_ICON = the DIB, group entries (u8 width,
    u8 height of the DIB (twice the icon), u8 colours, u8 0, u16 planes, u16 bits, u32 size, u16 id); a
    cursor: RT_CURSOR = u16 hot spot x, u16 y + the DIB, group entries (u16 width, u16 height, u16 0, u16 0,
    u32 size, u16 id);
  - names of types and resources in capitals ("lower" -> "LOWER"); a new resource takes 101, or the
    largest id of its type + 1.
"""
import os
import struct

RT_CURSOR, RT_BITMAP, RT_ICON, RT_STRING, RT_GROUP_CURSOR, RT_GROUP_ICON = 1, 2, 3, 6, 12, 14
FLAGS = 0x1030
KIND_NAMES = {"bitmap": "Bitmap", "cursor": "Cursor", "icon": "Icon", "string": "String Table"}


def _id(b, p):
    if b[p:p + 2] == b"\xff\xff":
        return struct.unpack_from("<H", b, p + 2)[0], p + 4
    end = p
    while end + 1 < len(b) and b[end:end + 2] != b"\0\0":
        end += 2
    return b[p:end].decode("utf-16-le", "replace"), end + 2


def read_raw(data):
    """[(type, name, lang, flags, data)] of a .res file (the empty first entry left out)"""
    out, p = [], 0
    while p + 8 <= len(data):
        size, hsize = struct.unpack_from("<II", data, p)
        if hsize < 16 or p + hsize > len(data):
            break
        t, q = _id(data, p + 8)
        n, q = _id(data, q)
        q = (q + 3) & ~3
        flags, lang = (struct.unpack_from("<HH", data, q + 4) if q + 8 <= len(data) else (0, 0))
        body = data[p + hsize:p + hsize + size]
        if size or t != 0:
            out.append((t, n, lang, flags, body))
        p = (p + hsize + size + 3) & ~3
    return out


def _name_bytes(v):
    if isinstance(v, int):
        return struct.pack("<HH", 0xFFFF, v)
    return str(v).upper().encode("utf-16-le") + b"\0\0"


def entry_bytes(t, n, lang, flags, data):
    head = struct.pack("<II", len(data), 0) + _name_bytes(t) + _name_bytes(n)
    head += b"\0" * (-len(head) % 4)
    head += struct.pack("<IHHII", 0, flags, lang, 0, 0)
    head = head[:4] + struct.pack("<I", len(head)) + head[8:]
    return head + data + b"\0" * (-len(data) % 4)


EMPTY = entry_bytes(0, 0, 0, 0, b"")


def norm_name(v):
    """an id as the Resource Editor keeps it: a number, or a name in capitals (without the quotes)"""
    if isinstance(v, int):
        return v
    s = str(v).strip()
    if len(s) >= 2 and s[0] == s[-1] == '"':
        return s[1:-1].upper()
    try:
        return int(s)
    except ValueError:
        return s.upper()


def show_name(v):
    """an id as the tree and Edit Properties show it: 101, "INTRO\""""
    return str(v) if isinstance(v, int) else f'"{v}"'


class Item:
    """one resource of the tree: kind "bitmap" | "icon" | "cursor" | "custom"; type is the type of a custom
    resource (a number or a name); images of an icon / a cursor: [(entry fields, image data, image lang)]"""

    def __init__(self, kind, name, lang, data=b"", rtype=None, images=None):
        self.kind, self.name, self.lang, self.data = kind, name, lang, data
        self.rtype = rtype
        self.images = images or []

    def copy(self):
        return Item(self.kind, self.name, self.lang, self.data, self.rtype,
                    [(dict(e), d, lg) for e, d, lg in self.images])

    def type_key(self):
        return self.rtype if self.kind == "custom" else self.kind


class ResFile:
    """the resource file: string tables {lang: {id: text}} (in column order) and the other resources"""

    def __init__(self, lang=0x0409):
        self.lang = lang                        # the language of new resources (the system's)
        self.strings = {}                       # lang -> {id: text}
        self.items = []
        self.path = None

    # ------------------------------------------------------------ reading
    @classmethod
    def load(cls, path, lang=0x0409):
        with open(path, "rb") as f:
            data = f.read()
        r = cls.from_bytes(data, lang)
        r.path = path
        return r

    @classmethod
    def from_bytes(cls, data, lang=0x0409):
        r = cls(lang)
        raw = read_raw(data)
        images = {(t, n): (d, lg) for t, n, lg, fl, d in raw if t in (RT_ICON, RT_CURSOR)}
        for t, n, lg, fl, d in raw:
            if t == RT_STRING and isinstance(n, int):
                tab = r.strings.setdefault(lg, {})
                p, base = 0, (n - 1) * 16
                for i in range(16):
                    if p + 2 > len(d):
                        break
                    ln = struct.unpack_from("<H", d, p)[0]
                    p += 2
                    if ln:
                        tab[base + i] = d[p:p + 2 * ln].decode("utf-16-le", "replace")
                    p += 2 * ln
            elif t in (RT_GROUP_ICON, RT_GROUP_CURSOR):
                cur = t == RT_GROUP_CURSOR
                cnt = struct.unpack_from("<H", d, 4)[0] if len(d) >= 6 else 0
                imgs = []
                for k in range(cnt):
                    e = d[6 + 14 * k:20 + 14 * k]
                    if len(e) < 14:
                        break
                    if cur:
                        w, h, pl, bc, sz, iid = struct.unpack("<HHHHIH", e)
                        ent = {"w": w, "h": h, "planes": pl, "bits": bc}
                    else:
                        w, h, cc, rs, pl, bc, sz, iid = struct.unpack("<BBBBHHIH", e)
                        ent = {"w": w, "h": h, "colors": cc, "res": rs, "planes": pl, "bits": bc}
                    img, ilang = images.get((RT_CURSOR if cur else RT_ICON, iid), (b"", lg))
                    imgs.append((ent, img, ilang))
                r.items.append(Item("cursor" if cur else "icon", n, lg, images=imgs))
            elif t in (RT_ICON, RT_CURSOR):
                continue
            elif t == RT_BITMAP:
                r.items.append(Item("bitmap", n, lg, d))
            else:
                r.items.append(Item("custom", n, lg, d, rtype=t))
        return r

    # ------------------------------------------------------------ writing
    def to_bytes(self):
        out = [EMPTY]
        for lg, tab in self.strings.items():
            blocks = sorted({(i >> 4) + 1 for i, s in tab.items() if s})
            for blk in blocks:
                d = b""
                for i in range((blk - 1) * 16, blk * 16):
                    s = tab.get(i, "")
                    u = s.encode("utf-16-le")
                    d += struct.pack("<H", len(u) // 2) + u
                out.append(entry_bytes(RT_STRING, blk, lg, 0, d))
        for kind, rt_img, rt_grp in (("cursor", RT_CURSOR, RT_GROUP_CURSOR), ("icon", RT_ICON, RT_GROUP_ICON)):
            n = 0
            for it in self.items:
                if it.kind != kind:
                    continue
                grp = struct.pack("<HHH", 0, 2 if kind == "cursor" else 1, len(it.images))
                for ent, img, ilang in it.images:
                    n += 1
                    out.append(entry_bytes(rt_img, n, ilang, FLAGS, img))
                    if kind == "cursor":
                        grp += struct.pack("<HHHHIH", ent["w"], ent["h"], ent.get("planes", 0), ent.get("bits", 0),
                                           len(img), n)
                    else:
                        grp += struct.pack("<BBBBHHIH", ent["w"] & 0xFF, ent["h"] & 0xFF, ent.get("colors", 0) & 0xFF,
                                           ent.get("res", 0) & 0xFF, ent.get("planes", 1), ent.get("bits", 0), len(img), n)
                out.append(entry_bytes(rt_grp, it.name, it.lang, FLAGS, grp))
        for it in self.items:
            if it.kind == "bitmap":
                out.append(entry_bytes(RT_BITMAP, it.name, it.lang, FLAGS, it.data))
        for it in self.items:
            if it.kind == "custom":
                out.append(entry_bytes(it.rtype, it.name, it.lang, FLAGS, it.data))
        return b"".join(out)

    def save(self, path=None):
        path = path or self.path
        with open(path, "wb") as f:
            f.write(self.to_bytes())
        self.path = path

    # ------------------------------------------------------------ editing
    def next_id(self, type_key):
        ids = [it.name for it in self.items if it.type_key() == type_key and isinstance(it.name, int)]
        return max(ids) + 1 if ids else 101

    def add_bitmap(self, file_data):
        if file_data[:2] != b"BM" or len(file_data) < 54:
            raise ValueError("not a bitmap")
        it = Item("bitmap", self.next_id("bitmap"), self.lang, file_data[14:])
        self.items.append(it)
        return it

    def add_icon(self, file_data, cursor=False):
        res, typ, cnt = struct.unpack_from("<HHH", file_data, 0)
        if res != 0 or typ != (2 if cursor else 1) or cnt == 0:
            raise ValueError("not an icon" if not cursor else "not a cursor")
        imgs = []
        for k in range(cnt):
            bw, bh, cc, rs, f1, f2, size, off = struct.unpack_from("<BBBBHHII", file_data, 6 + 16 * k)
            dib = file_data[off:off + size]
            if len(dib) < 16:
                raise ValueError("a broken image")
            dw, dh = struct.unpack_from("<ii", dib, 4)
            planes, bits = struct.unpack_from("<HH", dib, 12)
            if cursor:
                imgs.append(({"w": bw or 256, "h": bh or 256, "planes": 0, "bits": 0},
                             struct.pack("<HH", f1, f2) + dib, self.lang))
            else:
                imgs.append(({"w": dw, "h": dh, "colors": cc, "res": 0, "planes": planes, "bits": bits},
                             dib, self.lang))
        kind = "cursor" if cursor else "icon"
        it = Item(kind, self.next_id(kind), self.lang, images=imgs)
        self.items.append(it)
        return it

    def add_custom(self, data, rtype="CUSTOM"):
        rtype = norm_name(rtype)
        it = Item("custom", None, self.lang, data, rtype=rtype)
        it.name = self.next_id(rtype)
        self.items.append(it)
        return it

    def duplicate(self, it):
        new = it.copy()
        new.name = self.next_id(it.type_key())
        self.items.insert(self.items.index(it) + 1, new)
        return new

    def remove(self, it):
        i = self.items.index(it)
        del self.items[i]
        return i

    # ------------------------------------------------------------ string tables
    def set_string(self, lang, sid, text):
        tab = self.strings.setdefault(lang, {})
        if text:
            tab[sid] = text
        else:
            tab.pop(sid, None)

    def string_ids(self):
        return sorted({i for tab in self.strings.values() for i in tab})


# ---------------------------------------------------------------- the tree of the Resource Editor
def tree(res):
    """[(folder caption, [(item caption, object)])] in the order the Resource Editor shows them: the folders
    alphabetically ("NAME" first), the items by id; a language other than the system's after the id"""
    folders = {}
    from . import lcid as L
    for lg, tab in res.strings.items():
        if not tab:                             # a column without strings is not in the file nor in the tree
            continue
        cap = "String Table" + ("" if lg == res.lang else f" [{L.name(lg)}]")
        folders.setdefault("String Table", []).append((cap, ("string", lg)))
    for it in res.items:
        fold = KIND_NAMES[it.kind] if it.kind != "custom" else show_name(it.rtype)
        cap = show_name(it.name) + ("" if it.lang == res.lang else f" [{L.name(it.lang)}]")
        folders.setdefault(fold, []).append((cap, it))
    return sorted(folders.items(), key=lambda kv: kv[0].upper())


def file_kind(path):
    ext = os.path.splitext(path)[1].lower()
    return {".bmp": "bitmap", ".dib": "bitmap", ".ico": "icon", ".cur": "cursor"}.get(ext, "custom")
