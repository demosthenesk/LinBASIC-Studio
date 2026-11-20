# LinBASIC Studio - .frx/.ctx/.dox/.pgx binary property stores
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""The .frm text refers to binary data as  Name = "Form1.frx":0012  (or with a
'$' prefix for long or multi-line strings).  Each referenced record starts at
that offset; its extent is found from the record's own length prefix, and is
cross-checked against the next referenced offset.

Record layouts (checked in Phase 0 on the VB98 samples of the MSDN library,
tools/rea/check_samples.py, 75 files: pictures, ImageList images, $ strings (u32 length)
and short Text (u8 length) confirmed; List/ItemData layout taken from a form that
VB6 itself saved, a test form saved in the VB6 IDE):

  picture   [u32 total][ 'lt\\0\\0' ][u32 size][image bytes: BMP/ICO/CUR/GIF/JPG/WMF]
            or ['lt\\0\\0'][u32 size][image]
  listimage [u32 total][CLSID_StdPicture (16 bytes)]['lt\\0\\0'][u32 size][image]   (ImageList)
  list      [u16 count][u16 max length] then count x ([u16 len][bytes])   List
  itemdata  same layout; the numbers are stored as text ("10", "-1")       ItemData
            (an empty list is only [u16 count = 0]; written by VB6 itself in
            that test form.  The second word is >= the longest item;
            VB6 wrote 3 for "10","20","-1", so readers must not rely on it.)
  string    [u32 len][bytes]  |  [u16 len][bytes]  |  [u8 len][bytes]   ($ references)
            VB6 (a test form saved by VB6): a string with a line break and at
            most 255 bytes is  Text = "Form1.frx":0161  with a u8 length; a longer one is
            Text = $"Form1.frx":0174  with a u32 length.  The records follow the order of the
            properties in the .frm.
            VB6 again (the designer test d_frx, saved by VB6): the Text of a TextBox is never a $
            reference: [u8 len] up to 254 bytes, else [0xFF][u16 len]; Caption, Tag and ToolTipText
            in the .frx are always $ references with [u32 len].  The second word of a List or
            ItemData record is the longest item but at least 3; a ListBox or ComboBox with a List
            always has an ItemData (0 for the items without one).
"""
import struct

PIC_MAGIC = b"lt\x00\x00"
CLSID_STDPICTURE = bytes.fromhex("0452e30b918fce119de300aa004bb851")   # {0BE35204-8F91-11CE-9DE3-00AA004BB851}


def sniff_image(data):
    if data[:2] == b"BM":
        return "bmp"
    if data[:4] in (b"\x00\x00\x01\x00",):
        return "ico"
    if data[:4] == b"\x00\x00\x02\x00":
        return "cur"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if data[:3] == b"\xff\xd8\xff":
        return "jpg"
    if data[:4] == b"\xd7\xcd\xc6\x9a" or data[:4] == b"\x01\x00\x09\x00":
        return "wmf"
    if data[:4] == b"\x01\x00\x00\x00" and data[40:44] == b" EMF":
        return "emf"
    return None


class FrxData:
    """a value that a .frx keeps, while the designer edits it (before it has an offset):
    kind is picture (image file bytes), list, itemdata, string or raw"""
    __slots__ = ("kind", "data", "width")

    def __init__(self, kind, data, width=None):
        self.kind, self.data, self.width = kind, data, width     # width: the 2nd word of a list

    def __eq__(self, o):
        return isinstance(o, FrxData) and (self.kind, self.data, self.width) == (o.kind, o.data, o.width)

    def __ne__(self, o):
        return not self == o

    def __hash__(self):
        return hash((self.kind, str(self.data)))

    def __repr__(self):
        return f"FrxData({self.kind}, {len(self.data) if self.data is not None else 0})"


def needs_frx(s):
    """VB6 keeps a string in the .frx when it has a line break or more than 255 characters"""
    return isinstance(s, str) and ("\r" in s or "\n" in s or len(s) > 255)


class FrxStore:
    def __init__(self, data, offsets=(), encoding="cp1252"):
        self.data = data
        self.encoding = encoding
        self.offsets = sorted(set(offsets))

    @classmethod
    def load(cls, path, offsets=(), encoding="cp1252"):
        with open(path, "rb") as f:
            return cls(f.read(), offsets, encoding)

    def _limit(self, off):
        """Upper bound of the record at off: next referenced offset or EOF."""
        for o in self.offsets:
            if o > off:
                return o
        return len(self.data)

    def raw(self, off):
        return self.data[off:self._limit(off)]

    # ---- decoders ---------------------------------------------------------
    def picture(self, off):
        """-> (format, image_bytes) or (None, b'') for an empty picture."""
        d = self.data
        if d[off + 4:off + 20] == CLSID_STDPICTURE and d[off + 20:off + 24] == PIC_MAGIC:
            size = struct.unpack_from("<I", d, off + 24)[0]
            img = d[off + 28:off + 28 + size]
        elif d[off + 4:off + 8] == PIC_MAGIC:
            size = struct.unpack_from("<I", d, off + 8)[0]
            img = d[off + 12:off + 12 + size]
        elif d[off:off + 4] == PIC_MAGIC:
            size = struct.unpack_from("<I", d, off + 4)[0]
            img = d[off + 8:off + 8 + size]
        else:
            img = self.raw(off)
        return (sniff_image(img), img) if img else (None, b"")

    def text(self, off):
        """the Text of a TextBox: [u8 len] or [0xFF][u16 len] (or [0xFF][0xFFFF][u32 len])"""
        d = self.data
        n, p = d[off], off + 1
        if n == 0xFF:
            n, p = struct.unpack_from("<H", d, p)[0], p + 2
            if n == 0xFFFF:
                n, p = struct.unpack_from("<I", d, p)[0], p + 4
        return d[p:p + n].decode(self.encoding, errors="replace")

    def string(self, off):
        lim = self._limit(off)
        d = self.data
        for fmt, hdr in (("<I", 4), ("<H", 2), ("<B", 1)):
            if off + hdr > len(d):
                continue
            n = struct.unpack_from(fmt, d, off)[0]
            if off + hdr + n <= lim and (off + hdr + n == lim or hdr == 4):
                return d[off + hdr:off + hdr + n].decode(self.encoding, errors="replace")
        return d[off:lim].decode(self.encoding, errors="replace")

    def list(self, off):
        d = self.data
        count = struct.unpack_from("<H", d, off)[0]
        if count == 0:
            return []
        p, out = off + 4, []                    # skip count and max length
        for _ in range(count):
            n = struct.unpack_from("<H", d, p)[0]
            out.append(d[p + 2:p + 2 + n].decode(self.encoding, errors="replace"))
            p += 2 + n
        return out

    def itemdata(self, off):
        return [int(s.strip() or 0) for s in self.list(off)]

    def list_width(self, off):
        """the second word of a list record (the longest item when VB6 wrote it)"""
        count = struct.unpack_from("<H", self.data, off)[0]
        return struct.unpack_from("<H", self.data, off + 2)[0] if count else None

    def data_of(self, prop_name, ref):
        """a record as FrxData, to be written again (the designer saves a new .frx)"""
        n = prop_name.lower()
        if n == "text" and not ref.is_string:
            return FrxData("text", self.text(ref.offset))
        if n == "text":
            return FrxData("text", self.string(ref.offset))
        if ref.is_string or n in ("caption", "tag", "tooltiptext"):
            return FrxData("string", self.string(ref.offset))
        if n == "list":
            return FrxData("list", self.list(ref.offset), self.list_width(ref.offset))
        if n == "itemdata":
            return FrxData("itemdata", self.itemdata(ref.offset), self.list_width(ref.offset))
        if n in ("picture", "icon", "mouseicon", "dragicon", "image", "downpicture",
                 "disabledpicture", "maskpicture", "palette"):
            fmt, img = self.picture(ref.offset)
            return FrxData("picture", img)
        return FrxData("raw", self.raw(ref.offset))

    def decode(self, prop_name, ref):
        """Decode by property name; unknown kinds come back as raw bytes."""
        n = prop_name.lower()
        if ref.is_string:
            return self.string(ref.offset)
        if n in ("picture", "icon", "mouseicon", "dragicon", "image", "downpicture",
                 "disabledpicture", "maskpicture", "palette") or \
                (n.startswith("bandpicture") and n[11:].isdigit()):     # the Picture of a band of a CoolBar
            return self.picture(ref.offset)
        if n == "list":
            return self.list(ref.offset)
        if n == "itemdata":
            return self.itemdata(ref.offset)
        if n == "text":
            return self.text(ref.offset)
        if n in ("caption", "tag", "tooltiptext"):
            return self.string(ref.offset)
        return self.raw(ref.offset)


class FrxWriter:
    """Builds a new .frx while the designer saves a form."""

    def __init__(self, encoding="cp1252"):
        self.buf = bytearray()
        self.encoding = encoding

    def _start(self):
        return len(self.buf)

    def picture(self, img, ole=False):
        """a picture: [u32 size] "lt\\0\\0" [u32 len] data, as VB6 writes the Picture / Icon of a form and of
        its intrinsic controls; ole: the picture of an ActiveX control (an ImageList's ListImage, a band of a
        CoolBar ...) - a persisted StdPicture, its CLSID first (VB6: DEV-8 f_cool6 Form1.frx)"""
        off = self._start()
        if ole and img:
            self.buf += struct.pack("<I", len(img) + 24) + CLSID_STDPICTURE + PIC_MAGIC + struct.pack("<I", len(img)) + img
        else:
            self.buf += struct.pack("<I", len(img) + 8) + PIC_MAGIC + struct.pack("<I", len(img)) + img
        return off

    def string(self, s):
        """Caption, Tag, ToolTipText: [u32 len], a $ reference -> (offset, True)"""
        b = s.encode(self.encoding, errors="replace")
        off = self._start()
        self.buf += struct.pack("<I", len(b)) + b
        return off, True

    def text(self, s):
        """the Text of a TextBox: [u8 len] or [0xFF][u16 len] -> (offset, False)"""
        b = s.encode(self.encoding, errors="replace")
        off = self._start()
        if len(b) < 0xFF:
            self.buf += struct.pack("<B", len(b))
        elif len(b) < 0xFFFF:
            self.buf += b"\xff" + struct.pack("<H", len(b))
        else:
            self.buf += b"\xff\xff\xff" + struct.pack("<I", len(b))
        self.buf += b
        return off, False

    def list(self, items, width=None):
        off = self._start()
        enc = [s.encode(self.encoding, errors="replace") for s in items]
        self.buf += struct.pack("<H", len(enc))
        if enc:
            self.buf += struct.pack("<H", width if width is not None else max([3] + [len(b) for b in enc]))
            for b in enc:
                self.buf += struct.pack("<H", len(b)) + b
        return off

    def itemdata(self, values, width=None):
        return self.list([str(int(v)) for v in values], width)

    def raw(self, data):
        off = self._start()
        if bytes(data[:4]) == b'LB\r\0':          # the OleObjectBlob of a DBGrid: its offsets count from the
            from .oleblob import relocate          # start of the .frx (oleblob.py)
            data = relocate(data, off)
        self.buf += data
        return off

    def write(self, d, ole=False):
        """FrxData -> (offset, is_dollar); ole: a property of an ActiveX control (its pictures are StdPicture
        objects with their CLSID)"""
        if d.kind == "picture":
            return self.picture(d.data, ole), False
        if d.kind == "string":
            return self.string(d.data)
        if d.kind == "text":
            return self.text(d.data)
        if d.kind == "list":
            return self.list(d.data, d.width), False
        if d.kind == "itemdata":
            return self.itemdata(d.data, d.width), False
        return self.raw(d.data), False

    def bytes(self):
        return bytes(self.buf)
