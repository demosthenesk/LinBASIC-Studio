#!/usr/bin/env python3
# LinBASIC Studio - Phase 0 (REA): IDE commands, dialogs, shortcuts and runtime errors of VB6
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""Usage:  python3 tools/rea/dump_resources.py <VB6-REA folder> <output folder>

Writes
  ide_commands.json    menu bars and command captions of the VB6 IDE (VB6IDE.DLL strings 19000-19499),
                       dialog titles, and the keyboard accelerator tables
  runtime_errors.json  run-time error numbers and messages of MSVBVM60 (VBA and VB errors),
                       which the runtime returns from Err.Description; '|1' and '~' mark
                       the places where VB inserts a name
"""
import json, os, struct, sys
import pefile

VK = {0x70 + i: f"F{i + 1}" for i in range(12)}
VK.update({0x2E: "Del", 0x2D: "Ins", 0x24: "Home", 0x23: "End", 0x21: "PgUp", 0x22: "PgDn",
           0x26: "Up", 0x28: "Down", 0x25: "Left", 0x27: "Right", 0x09: "Tab", 0x08: "Backspace",
           0x0D: "Enter", 0x13: "Pause", 0x03: "Break", 0x1B: "Esc", 0x20: "Space"})


def resources(pe, type_id):
    for t in getattr(pe, "DIRECTORY_ENTRY_RESOURCE").entries:
        if t.id == type_id:
            for e in t.directory.entries:
                leaf = e.directory.entries[0].data.struct
                yield (str(e.name) if e.name else e.id), pe.get_data(leaf.OffsetToData, leaf.Size)


def strings(pe):
    out = {}
    for block, data in resources(pe, 6):
        p = 0
        for i in range(16):
            n = struct.unpack_from("<H", data, p)[0]
            p += 2
            if n:
                out[(block - 1) * 16 + i] = data[p:p + 2 * n].decode("utf-16le")
            p += 2 * n
    return out


def message_table(pe):
    out = {}
    for _, d in resources(pe, 11):
        nblocks = struct.unpack_from("<I", d, 0)[0]
        for b in range(nblocks):
            lo, hi, off = struct.unpack_from("<III", d, 4 + 12 * b)
            for mid in range(lo, hi + 1):
                ln, flags = struct.unpack_from("<HH", d, off)
                raw = d[off + 4:off + ln]
                txt = raw.decode("utf-16le" if flags & 1 else "cp1252").rstrip("\0\r\n ")
                if txt:
                    out[mid] = txt
                off += ln
    return out


def dialog_titles(pe):
    out = {}
    for name, d in resources(pe, 5):
        p = 26 if d[:4] == b"\x01\x00\xff\xff" else 18
        vals = []
        for _ in range(3):                    # menu, class, title
            if struct.unpack_from("<H", d, p)[0] == 0xFFFF:
                vals.append(None)
                p += 4
                continue
            s = ""
            while True:
                w = struct.unpack_from("<H", d, p)[0]
                p += 2
                if not w:
                    break
                s += chr(w)
            vals.append(s)
        out[str(name)] = vals[2]
    return out


def accelerators(pe):
    tables = {}
    for name, d in resources(pe, 9):
        rows = []
        for k in range(0, len(d), 8):
            fv, key, cmd, _ = struct.unpack_from("<HHHH", d, k)
            mods = ("Ctrl+" if fv & 8 else "") + ("Shift+" if fv & 4 else "") + ("Alt+" if fv & 16 else "")
            kn = VK.get(key, chr(key) if fv & 1 and 32 < key < 127 else hex(key))
            rows.append({"keys": mods + kn, "command": cmd})
        tables[str(name)] = rows
    return tables


def main(src, out):
    os.makedirs(out, exist_ok=True)
    ide = pefile.PE(os.path.join(src, "VB98", "VB6IDE.DLL"))
    s = strings(ide)
    doc = {
        "source": "VB6IDE.DLL",
        "menus": {k: v for k, v in s.items() if 19000 <= k < 20000},
        "dialogs": dialog_titles(ide),
        "accelerators": accelerators(ide),
    }
    with open(os.path.join(out, "ide_commands.json"), "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=1)
    # Run-time errors (resolved in Phase 0):
    #  * VBA errors: MSVBVM60 keeps their messages as string ids 10000+i, the i-th error in
    #    ascending number order; the numbers are the ids under which VB6IDE.DLL keeps the
    #    same messages (1-999). Both lists are checked to be identical.
    #  * VB (forms/controls) errors: MSVBVM60 string id = 1000 + error number (260-747).
    #    Where a number exists in both lists (429), Err.Description gives the VBA message.
    rt = strings(pefile.PE(os.path.join(src, "System32", "msvbvm60.dll")))
    nums = sorted(k for k in s if 1 <= k < 1000)
    vba = {n: rt[10000 + i] for i, n in enumerate(nums)}
    mismatch = [n for n in nums if vba[n] != s[n]]
    if mismatch:
        raise SystemExit(f"VBA error tables differ at {mismatch}")
    vb = {k - 1000: v for k, v in rt.items() if 1000 <= k < 2000 and k - 1000 not in vba}
    errs = {"vba": vba, "vb": vb}
    with open(os.path.join(out, "runtime_errors.json"), "w", encoding="utf-8") as f:
        json.dump({"source": "MSVBVM60.DLL (numbers of the VBA errors from VB6IDE.DLL)", **errs}, f, ensure_ascii=False, indent=1)
    print(f"{len(doc['menus'])} IDE command strings, {len(doc['dialogs'])} dialogs, "
          f"{sum(len(t) for t in doc['accelerators'].values())} shortcuts, "
          f"{len(errs['vba'])} VBA + {len(errs['vb'])} VB run-time errors")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
