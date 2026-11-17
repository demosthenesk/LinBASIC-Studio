# LinBASIC Studio - reader for COM type libraries (MSFT format: .tlb .olb, TYPELIB resources of .dll/.ocx)
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""Pure-Python reader for the 'MSFT' type library format, written from the published
description of the format (the layout Wine's oleaut32 documents).  It is used

  * in Phase 0 (REA), to dump the object model of VB6.OLB, MSVBVM60.DLL and the OCX
    controls into the database the compiler, runtime and IDE are built from, and
  * by the IDE, to show the members of any type library the user references
    (Object Browser, IntelliSense, F1 through helpcontext).

    lib = TypeLib.load("VB6.OLB")          # or a PE file with a TYPELIB resource
    for t in lib.types: print(t.kind, t.name, [m.name for m in t.members])
"""
import struct

TKIND = {0: "enum", 1: "record", 2: "module", 3: "interface", 4: "dispatch",
         5: "coclass", 6: "alias", 7: "union"}
INVKIND = {1: "func", 2: "get", 4: "let", 8: "set"}
VT = {0: "Empty", 1: "Null", 2: "Integer", 3: "Long", 4: "Single", 5: "Double", 6: "Currency",
      7: "Date", 8: "String", 9: "Object", 10: "SCODE", 11: "Boolean", 12: "Variant",
      13: "IUnknown", 14: "Decimal", 16: "SByte", 17: "Byte", 18: "UShort", 19: "ULong",
      20: "LongLong", 21: "ULongLong", 22: "Long", 23: "ULong", 24: "Void", 25: "HRESULT",
      26: "Ptr", 27: "SafeArray", 28: "CArray", 29: "UserDefined", 30: "LPSTR", 31: "LPWSTR",
      36: "Record", 37: "LongPtr", 38: "ULongPtr", 64: "FileTime", 71: "ClipFormat"}
VT_BYREF, VT_ARRAY = 0x4000, 0x2000
# flags
TYPEFLAG = {0x1: "appobject", 0x2: "cancreate", 0x4: "licensed", 0x8: "predeclid", 0x10: "hidden",
            0x20: "control", 0x40: "dual", 0x80: "nonextensible", 0x100: "oleautomation",
            0x200: "restricted", 0x400: "aggregatable", 0x800: "replaceable", 0x1000: "dispatchable"}
FUNCFLAG = {0x1: "restricted", 0x2: "source", 0x4: "bindable", 0x8: "requestedit",
            0x10: "displaybind", 0x20: "defaultbind", 0x40: "hidden", 0x80: "usesgetlasterror",
            0x100: "defaultcollelem", 0x200: "uidefault", 0x400: "nonbrowsable",
            0x800: "replaceable", 0x1000: "immediatebind"}
PARAMFLAG_IN, PARAMFLAG_OUT, PARAMFLAG_LCID, PARAMFLAG_RETVAL = 1, 2, 4, 8
PARAMFLAG_OPT, PARAMFLAG_HASDEFAULT = 0x10, 0x20
IMPLTYPE_DEFAULT, IMPLTYPE_SOURCE = 1, 2


def _flags(v, table):
    return [n for b, n in table.items() if v & b]


def guid_str(b):
    d1, d2, d3 = struct.unpack_from("<IHH", b)
    return "{%08X-%04X-%04X-%s-%s}" % (d1, d2, d3, b[8:10].hex().upper(), b[10:16].hex().upper())


class TypeDesc:
    """vt + (for Ptr/SafeArray) inner TypeDesc, (UserDefined) a type name, (CArray) bounds."""
    __slots__ = ("vt", "inner", "ref", "dims")

    def __init__(self, vt, inner=None, ref=None, dims=None):
        self.vt, self.inner, self.ref, self.dims = vt, inner, ref, dims

    def vb(self):
        """The type as VB6 writes it."""
        if self.vt == 26:                                       # pointer -> ByRef / object ref
            return self.inner.vb() if self.inner else "Long"
        if self.vt == 27:
            return (self.inner.vb() if self.inner else "Variant") + "()"
        if self.vt == 28:
            return (self.inner.vb() if self.inner else "?") + "(" + ",".join(
                f"{lo} To {lo + n - 1}" for n, lo in self.dims) + ")"
        if self.vt == 29:
            return self.ref or "?"
        return VT.get(self.vt, f"VT{self.vt}")

    def is_ptr(self):
        return self.vt == 26

    def __repr__(self):
        return self.vb()


class Param:
    __slots__ = ("name", "type", "flags", "default")

    def __init__(self, name, type_, flags, default=None):
        self.name, self.type, self.flags, self.default = name, type_, flags, default

    @property
    def optional(self):
        return bool(self.flags & (PARAMFLAG_OPT | PARAMFLAG_HASDEFAULT))

    @property
    def retval(self):
        return bool(self.flags & PARAMFLAG_RETVAL)


class Member:
    __slots__ = ("name", "memid", "kind", "invkind", "type", "params", "flags", "helpcontext",
                 "helpstring", "value", "vtoffset", "entry", "funckind", "callconv", "nopt")

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k))

    @property
    def hidden(self):
        return bool((self.flags or 0) & 0x40)

    @property
    def restricted(self):
        return bool((self.flags or 0) & 0x1)

    def vb_params(self):
        """Parameters as VB sees them (no retval, no lcid)."""
        return [p for p in self.params or [] if not p.retval and not p.flags & PARAMFLAG_LCID]

    def vb_type(self):
        """Return type as VB sees it (HRESULT functions return their retval param)."""
        for p in self.params or []:
            if p.retval:
                t = p.type
                return t.inner if t.is_ptr() and t.inner else t
        return self.type

    def signature(self):
        if self.kind == "const":
            return f"Const {self.name} = {self.value!r}"
        if self.kind == "var":
            return f"{self.name} As {self.type.vb()}"
        ps = []
        for p in self.vb_params():
            t = p.type
            by = "ByRef " if t.is_ptr() and t.inner and t.inner.vt not in (9, 13, 29) else ""
            if t.is_ptr() and t.inner:
                t = t.inner
            s = f"{by}{p.name} As {t.vb()}"
            if p.default is not None:
                s += f" = {p.default!r}"
            ps.append(("[" + s + "]") if p.optional else s)
        rt = self.vb_type()
        kw = {"get": "Property Get", "let": "Property Let", "set": "Property Set"}.get(self.invkind)
        if kw:
            return f"{kw} {self.name}({', '.join(ps)})" + (f" As {rt.vb()}" if self.invkind == "get" else "")
        if rt is None or rt.vt in (24, 25):
            return f"Sub {self.name}({', '.join(ps)})"
        return f"Function {self.name}({', '.join(ps)}) As {rt.vb()}"


class TypeInfo:
    def __init__(self, lib, index):
        self.lib, self.index = lib, index
        self.kind = self.name = self.guid = None
        self.flags = 0
        self.members = []
        self.impl = []          # [(TypeInfo name, impl flags)]
        self.alias = None
        self.helpcontext = 0
        self.helpstring = ""
        self.version = (0, 0)

    @property
    def flag_names(self):
        return _flags(self.flags, TYPEFLAG)

    @property
    def default_interface(self):
        for n, f in self.impl:
            if f & IMPLTYPE_DEFAULT and not f & IMPLTYPE_SOURCE:
                return n
        return self.impl[0][0] if self.impl and not self.impl[0][1] & IMPLTYPE_SOURCE else None

    @property
    def event_interface(self):
        for n, f in self.impl:
            if f & IMPLTYPE_SOURCE and f & IMPLTYPE_DEFAULT:
                return n
        for n, f in self.impl:
            if f & IMPLTYPE_SOURCE:
                return n
        return None

    def __repr__(self):
        return f"TypeInfo({self.kind} {self.name}, {len(self.members)} members)"


class TypeLibError(Exception):
    pass


class TypeLib:
    def __init__(self, data, source=None):
        self.d = data
        self.source = source
        self.types = []
        self._parse()

    # ---- loading ---------------------------------------------------------------
    @classmethod
    def load(cls, path, index=1):
        with open(path, "rb") as f:
            data = f.read()
        if data[:4] == b"MSFT":
            return cls(data, path)
        if data[:2] == b"MZ":
            return cls(extract_typelib_resource(data, index), path)
        raise TypeLibError(f"{path}: not a type library (SLTG libraries are not supported)")

    # ---- low level -------------------------------------------------------------------
    def i32(self, o):
        return struct.unpack_from("<i", self.d, o)[0]

    def i16(self, o):
        return struct.unpack_from("<h", self.d, o)[0]

    def seg(self, n):
        return self.segs[n]

    def name_at(self, off):
        if off < 0:
            return None
        base = self.seg(7)[0] + off
        ln = self.d[base + 8]
        return self.d[base + 12:base + 12 + ln].decode("cp1252", "replace")

    def string_at(self, off):
        if off < 0:
            return ""
        base = self.seg(8)[0] + off
        ln = struct.unpack_from("<H", self.d, base)[0]
        return self.d[base + 2:base + 2 + ln].decode("cp1252", "replace")

    def guid_at(self, off):
        if off < 0:
            return None
        return guid_str(self.d[self.seg(5)[0] + off:self.seg(5)[0] + off + 16])

    def value_at(self, off):
        """Constant / default value (MSFT packs small integers into the offset itself)."""
        if off < 0:
            vt = (off & 0x7C000000) >> 26
            v = off & 0x3FFFFFF
            return self._coerce(vt, v)
        base = self.seg(11)[0] + off
        vt = struct.unpack_from("<H", self.d, base)[0]
        p = base + 2
        if vt in (0, 1):
            return None
        if vt == 8:                       # BSTR: length then chars
            n = struct.unpack_from("<i", self.d, p)[0]
            return self.d[p + 4:p + 4 + n].decode("cp1252", "replace") if n > 0 else ""
        if vt == 2:
            return struct.unpack_from("<h", self.d, p)[0]
        if vt in (3, 22, 10):
            return struct.unpack_from("<i", self.d, p)[0]
        if vt in (17, 16):
            return self.d[p]
        if vt in (18,):
            return struct.unpack_from("<H", self.d, p)[0]
        if vt in (19, 23):
            return struct.unpack_from("<I", self.d, p)[0]
        if vt == 4:
            return struct.unpack_from("<f", self.d, p)[0]
        if vt in (5, 7):
            return struct.unpack_from("<d", self.d, p)[0]
        if vt == 6:
            return struct.unpack_from("<q", self.d, p)[0] / 10000
        if vt == 11:
            return struct.unpack_from("<h", self.d, p)[0] != 0
        return struct.unpack_from("<i", self.d, p)[0]

    @staticmethod
    def _coerce(vt, v):
        if vt == 11:
            return v != 0
        if vt == 2 and v & 0x8000:
            return v - 0x10000
        return v

    # ---- type descriptions -------------------------------------------------------------
    def tdesc(self, t):
        if t < 0:
            vt = t & 0x7FFF
            base = vt & 0xFFF
            td = TypeDesc(base)
            if vt & VT_BYREF:
                td = TypeDesc(26, td)
            if vt & VT_ARRAY:
                td = TypeDesc(27, td)
            return td
        o = self.seg(9)[0] + t
        vt, _, w2, w3 = struct.unpack_from("<hhhh", self.d, o)
        vt &= 0x7FFF
        if vt in (26, 27):
            inner = TypeDesc(w2 & 0xFFF) if w3 < 0 else self.tdesc(w2 & 0xFFFF)
            return TypeDesc(vt, inner)
        if vt == 28:
            ao = self.seg(10)[0] + ((w3 << 16) | (w2 & 0xFFFF))
            inner = self.tdesc(self.i32(ao))
            ndims = struct.unpack_from("<H", self.d, ao + 4)[0]
            dims = [struct.unpack_from("<ii", self.d, ao + 8 + 8 * i) for i in range(ndims)]
            return TypeDesc(28, inner, dims=dims)
        if vt == 29:
            href = ((w3 & 0xFFFF) << 16) | (w2 & 0xFFFF)
            return TypeDesc(29, ref=self.href_name(href))
        return TypeDesc(vt)

    def href_name(self, href):
        if href == -1 or href is None:
            return None
        if href & 1:                       # imported: offset into the ImpInfo table
            o = self.seg(1)[0] + (href & ~3)
            flags, ofile, oguid = struct.unpack_from("<iii", self.d, o)
            libname = self._impfile_name(ofile)
            if flags & 0x10000:                   # oguid is an offset in the GUID table
                return f"{libname}:{self.guid_at(oguid)}"
            return f"{libname}:#{oguid}"          # oguid is a type index in that library
        idx = href // 0x64
        return self._names[idx] if 0 <= idx < len(self._names) else f"#{href}"

    def _impfile_name(self, ofile):
        o = self.seg(2)[0] + ofile
        n = struct.unpack_from("<H", self.d, o + 12)[0] >> 2
        return self.d[o + 14:o + 14 + n].decode("cp1252", "replace")

    # ---- parsing ---------------------------------------------------------------------------
    def _parse(self):
        d = self.d
        if d[:4] != b"MSFT":
            raise TypeLibError("not an MSFT type library")
        (magic1, magic2, posguid, lcid, lcid2, varflags, version, flags, ntypes, helpstring,
         hsctx, helpctx, ntnames, ntchars, nameoff, helpfile, custoff, r44, r48, disppos,
         nimp) = struct.unpack_from("<21i", d, 0)
        p = 0x54
        if varflags & 0x100:
            p += 4
        tioffs = struct.unpack_from(f"<{ntypes}i", d, p)
        p += 4 * ntypes
        self.segs = [struct.unpack_from("<iiii", d, p + 16 * i)[:2] for i in range(15)]
        self.name = self.name_at(nameoff)
        self.guid = self.guid_at(posguid)
        self.version = (version & 0xFFFF, version >> 16)
        self.lcid = lcid
        self.helpfile = self.string_at(helpfile)
        self.helpstring = self.string_at(helpstring)
        self.helpcontext = helpctx
        tbase = self.seg(0)[0]
        self._names = [self.name_at(self.i32(tbase + i * 0x64 + 0x34)) for i in range(ntypes)]
        for i in range(ntypes):
            self.types.append(self._typeinfo(i, tbase + i * 0x64))
        self.by_name = {t.name.lower(): t for t in self.types}

    def _typeinfo(self, i, o):
        d = self.d
        f = struct.unpack_from("<25i", d, o)
        ti = TypeInfo(self, i)
        ti.kind = TKIND.get(f[0] & 0xF, f"kind{f[0] & 0xF}")
        memoffset = f[1]
        cfuncs, cvars = f[6] & 0xFFFF, (f[6] >> 16) & 0xFFFF
        ti.guid = self.guid_at(f[11])
        ti.flags = f[12]
        ti.name = self._names[i]
        ti.version = (f[14] & 0xFFFF, f[14] >> 16)
        ti.helpstring = self.string_at(f[15])
        ti.helpcontext = f[17]
        cimpl = struct.unpack_from("<h", d, o + 0x4C)[0]
        datatype1 = f[21]
        if ti.kind == "coclass" and datatype1 >= 0:
            ro = datatype1
            seen = 0
            while ro >= 0 and seen < 64:
                rt, rflags, _, onext = struct.unpack_from("<iiii", d, self.seg(3)[0] + ro)
                ti.impl.append((self.href_name(rt), rflags))
                ro = onext
                seen += 1
        elif ti.kind in ("interface", "dispatch") and cimpl > 0 and datatype1 != -1:
            ti.impl.append((self.href_name(datatype1), 0))
        elif ti.kind == "alias":
            ti.alias = self.tdesc(datatype1)
        if memoffset >= 0 and (cfuncs or cvars):
            self._members(ti, memoffset, cfuncs, cvars)
        return ti

    def _members(self, ti, off, cfuncs, cvars):
        d = self.d
        infolen = self.i32(off)
        tbl = off + 4 + infolen         # memids[], names[], offsets[]
        n = cfuncs + cvars
        rec = off + 4
        for i in range(cfuncs):
            info = self.i32(rec)
            rl = info & 0xFFFF
            dtype, fflags = struct.unpack_from("<ii", d, rec + 4)
            vtoff, _ = struct.unpack_from("<hh", d, rec + 12)
            fkccic = self.i32(rec + 16)
            nargs, noargs = struct.unpack_from("<hh", d, rec + 20)
            has_def = bool(fkccic & 0x1000)
            nopt = (rl - 24 - nargs * 12 - (nargs * 4 if has_def else 0)) // 4
            opt = [self.i32(rec + 24 + 4 * k) for k in range(max(0, nopt))]
            pbase = rec + rl - nargs * 12
            dbase = pbase - nargs * 4
            params = []
            for a in range(nargs):
                pt, pname, pfl = struct.unpack_from("<iii", d, pbase + 12 * a)
                dv = None
                if has_def and pfl & PARAMFLAG_HASDEFAULT:
                    dv = self.value_at(self.i32(dbase + 4 * a))
                params.append(Param(self.name_at(pname) or f"arg{a + 1}", self.tdesc(pt), pfl, dv))
            memid = self.i32(tbl + 4 * i)
            name = self.name_at(self.i32(tbl + 4 * (n + i)))
            ti.members.append(Member(
                name=name, memid=memid, kind="func", invkind=INVKIND.get((fkccic >> 3) & 0xF, "func"),
                funckind=fkccic & 7, callconv=(fkccic >> 8) & 0xF, type=self.tdesc(dtype),
                params=params, flags=fflags & 0xFFFF, vtoffset=vtoff,
                nopt=noargs,                   # cParamsOpt: the last ones are optional Variants; -1: ParamArray
                helpcontext=opt[0] if len(opt) > 0 else 0,
                helpstring=self.string_at(opt[1]) if len(opt) > 1 else "",
                entry=opt[2] if len(opt) > 2 else None))
            rec += rl
        if cvars:
            rec = off + 4 + self.i32(tbl + 4 * (2 * n + cfuncs))
            for i in range(cvars):
                info = self.i32(rec)
                rl = info & 0xFFFF
                dtype, vflags = struct.unpack_from("<ii", d, rec + 4)
                varkind = struct.unpack_from("<h", d, rec + 12)[0]
                offsvalue = self.i32(rec + 16)
                nopt = (rl - 20) // 4
                opt = [self.i32(rec + 20 + 4 * k) for k in range(max(0, nopt))]
                memid = self.i32(tbl + 4 * (cfuncs + i))
                name = self.name_at(self.i32(tbl + 4 * (n + cfuncs + i)))
                const = varkind == 2
                ti.members.append(Member(
                    name=name, memid=memid, kind="const" if const else "var", type=self.tdesc(dtype),
                    flags=vflags & 0xFFFF, value=self.value_at(offsvalue) if const else None,
                    helpcontext=opt[0] if opt else 0,
                    helpstring=self.string_at(opt[1]) if len(opt) > 1 else ""))
                rec += rl

    def __repr__(self):
        return f"TypeLib({self.name} {self.guid} v{self.version[0]}.{self.version[1]}, {len(self.types)} types)"


def extract_typelib_resource(pe_bytes, index=1):
    """Return the raw TYPELIB resource of a PE image (.dll/.ocx/.exe)."""
    try:
        import pefile
    except ImportError:            # pragma: no cover
        raise TypeLibError("reading a TYPELIB resource needs the 'pefile' module")
    pe = pefile.PE(data=pe_bytes, fast_load=True)
    pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_RESOURCE"]])
    for t in getattr(pe, "DIRECTORY_ENTRY_RESOURCE", []).entries if hasattr(pe, "DIRECTORY_ENTRY_RESOURCE") else []:
        if t.name is not None and str(t.name).upper() == "TYPELIB":
            for e in t.directory.entries:
                if e.id == index or (index is None):
                    leaf = e.directory.entries[0].data.struct
                    return pe.get_data(leaf.OffsetToData, leaf.Size)
            leaf = t.directory.entries[0].directory.entries[0].data.struct
            return pe.get_data(leaf.OffsetToData, leaf.Size)
    raise TypeLibError("no TYPELIB resource")
