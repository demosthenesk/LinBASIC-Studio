#!/usr/bin/env python3
# LinBASIC Studio - Phase 0 (REA): VB6 object model database from the type libraries
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""Usage:  python3 tools/rea/dump_typelibs.py <VB6-REA folder> <output folder>

Reads the type libraries of a local VB6 installation (VB6-REA, made by
tools/rea/extract-vb6-msdn.sh) and writes one JSON file per library:
types, members, signatures, DISPIDs, enum values, helpcontext ids and, for the
VBA library, the MSVBVM60 export behind every runtime function.

Help strings are NOT copied (they are Microsoft's text); the IDE shows the
user's own MSDN help through helpfile + helpcontext instead.
"""
import json, os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "compiler"))
from lbvc.typelib import TypeLib, FUNCFLAG, _flags   # noqa: E402

LIBS = [  # (file in VB6-REA, TYPELIB resource index, output name)
    ("System32/stdole2.tlb", 1, "stdole"),
    ("System32/msvbvm60.dll", 1, "VBA"),
    ("System32/msvbvm60.dll", 3, "VBRUN"),
    ("VB98/VB6.OLB", 1, "VB"),
    ("VB98/VB6.EXE", 3, "VBInternal"),
    ("VB98/VB6EXT.OLB", 1, "VBIDE"),
    ("System32/mscomctl.OCX", 1, "MSComctlLib"),
    ("System32/mscomct2.ocx", 1, "MSComCtl2"),
    ("System32/ComCt332.ocx", 1, "ComCtl3"),
    ("System32/ComDlg32.OCX", 1, "MSComDlg"),
    ("System32/MSFlxGrd.ocx", 1, "MSFlexGridLib"),
    ("System32/richtx32.Ocx", 1, "RichTextLib"),
    ("System32/TabCtl32.Ocx", 1, "TabDlg"),
    ("System32/MSChrt20.ocx", 1, "MSChart20Lib"),
    ("System32/mci32.Ocx", 1, "MCI"),
    ("System32/MSComm32.Ocx", 1, "MSCommLib"),
    ("System32/PicClp32.Ocx", 1, "PicClip"),
    ("System32/scrrun.dll", 1, "Scripting"),
    ("CommonFiles/ado/msado28.tlb", 1, "ADODB"),
    ("CommonFiles/DAO/dao360.dll", 1, "DAO"),
    ("System32/dbgrid32.ocx", 1, "MSDBGrid"),
    ("System32/dblist32.Ocx", 1, "MSDBCtls"),
    ("System32/MSAdoDc.ocx", 1, "MSAdodcLib"),
    ("System32/MSDatGrd.ocx", 1, "MSDataGridLib"),
    ("System32/MSDatLst.Ocx", 1, "MSDataListLib"),
    ("System32/MSBind.dll", 1, "MSBind"),
    ("System32/msstdfmt.dll", 1, "StdFormat"),
    ("System32/MShflxgd.ocx", 1, "MSHierarchicalFlexGridLib"),
    ("System32/MSDbRptr.dll", 1, "MSDataReportLib"),
    ("System32/SysInfo.ocx", 1, "SysInfoLib"),
    # Phase 7
    ("System32/MSWINSCK.ocx", 1, "MSWinsockLib"),
    ("System32/MSINET.Ocx", 1, "InetCtlsObjects"),
    ("System32/msmask32.ocx", 1, "MSMask"),
    ("System32/MSDatRep.Ocx", 1, "MSDataRepeaterLib"),
    ("System32/msrdc20.Ocx", 1, "MSRDC"),
    ("System32/msmapi32.Ocx", 1, "MSMAPI"),
    ("System32/msrdo20.dll", 1, "RDO"),          # phase 7.8: Microsoft Remote Data Object 2.0 (VB6 Enterprise)
    # Phase 9.2: the WebBrowser control (Microsoft Internet Controls) of the Application Wizard
    ("System32/ieframe.dll", 1, "SHDocVw"),
]
PFLAGS = {1: "in", 2: "out", 4: "lcid", 8: "retval", 0x10: "optional", 0x20: "hasdefault"}


def exports_by_ordinal(path):
    try:
        import pefile
        pe = pefile.PE(path, fast_load=True)
        pe.parse_data_directories(directories=[pefile.DIRECTORY_ENTRY["IMAGE_DIRECTORY_ENTRY_EXPORT"]])
        return {e.ordinal: e.name.decode() for e in pe.DIRECTORY_ENTRY_EXPORT.symbols if e.name}
    except Exception:
        return {}


def main(src, out):
    os.makedirs(out, exist_ok=True)
    try:                                # what an earlier run wrote: kept for a library missing now
        old = {e["short"]: e for e in json.load(open(os.path.join(out, "index.json"), encoding="utf-8"))}
    except (OSError, ValueError):
        old = {}
    libs = []
    for rel, idx, short in LIBS:
        p = os.path.join(src, rel)
        if not os.path.exists(p):
            print("missing", rel, "(its JSON of an earlier run is kept)" if short in old else "")
            libs.append((None, rel, short, p))
            continue
        libs.append((TypeLib.load(p, idx), rel, short, p))
    # GUID -> qualified name, to resolve references between libraries
    by_guid = {}
    for lib, _, short, _ in libs:
        if lib is None:                 # a library missing now: the types of its JSON of an earlier run
            try:                        # (VB.json names DAO.Database even without dao360.dll)
                doc = json.load(open(os.path.join(out, short + ".json"), encoding="utf-8"))
            except (OSError, ValueError):
                continue
            for t in doc.get("types", []):
                if t.get("guid"):
                    by_guid[t["guid"]] = f"{doc['library']}.{t['name']}"
            continue
        for t in lib.types:
            if t.guid:
                by_guid[t.guid] = f"{lib.name}.{t.name}"

    ref_re = re.compile(r"[\w.]+:(\{[0-9A-F-]+\})")

    def fix(name):
        """'stdole2.tlb:{GUID}' -> 'stdole.IDispatch' when the library is loaded."""
        return ref_re.sub(lambda m: by_guid.get(m.group(1), m.group(0)), name) if name else name

    def tdesc(td):
        return None if td is None else fix(td.vb())

    index = []
    for lib, rel, short, path in libs:
        if lib is None:
            if short in old:
                index.append(old[short])
            continue
        exports = exports_by_ordinal(path) if short == "VBA" else {}
        types = []
        for t in lib.types:
            ty = {"name": t.name, "kind": t.kind, "guid": t.guid, "flags": t.flag_names,
                  "helpcontext": t.helpcontext}
            if t.impl:
                ty["implements"] = [{"name": fix(n), "default": bool(f & 1), "source": bool(f & 2)}
                                    for n, f in t.impl]
            if t.alias is not None:
                ty["alias"] = tdesc(t.alias)
            mems = []
            for m in t.members:
                e = {"name": m.name, "memid": m.memid, "kind": m.kind}
                if m.kind == "func":
                    e["invoke"] = m.invkind
                    e["returns"] = tdesc(m.vb_type())
                    e["params"] = [{"name": p.name, "type": tdesc(p.type.inner if p.type.is_ptr() and p.type.inner else p.type),
                                    "byref": p.type.is_ptr(), "flags": _flags(p.flags, PFLAGS),
                                    **({"default": p.default} if p.default is not None else {})}
                                   for p in m.params]
                    if m.nopt:              # cParamsOpt: the last n are optional (Variant); -1: ParamArray
                        e["optcount"] = m.nopt
                    if m.entry not in (None, -1, 0) and exports:
                        e["export"] = exports.get(m.entry, m.entry)
                else:
                    e["type"] = tdesc(m.type)
                    if m.kind == "const":
                        e["value"] = m.value
                fl = _flags(m.flags or 0, FUNCFLAG)
                if fl:
                    e["flags"] = fl
                if m.helpcontext:
                    e["helpcontext"] = m.helpcontext
                mems.append(e)
            if mems:
                ty["members"] = mems
            types.append(ty)
        doc = {"library": lib.name, "short": short, "guid": lib.guid,
               "version": "%d.%d" % lib.version, "lcid": lib.lcid, "helpfile": lib.helpfile,
               "file": os.path.basename(rel), "types": types}
        fn = os.path.join(out, short + ".json")
        with open(fn, "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=1, default=str)
        n = sum(len(t.get("members", [])) for t in types)
        index.append({"short": short, "library": lib.name, "guid": lib.guid, "version": doc["version"],
                      "file": doc["file"], "helpfile": lib.helpfile, "types": len(types), "members": n})
        print(f"{short:14} {lib.name:14} {len(types):4} types {n:5} members  -> {fn}")
    with open(os.path.join(out, "index.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
