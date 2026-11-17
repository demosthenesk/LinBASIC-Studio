#!/usr/bin/env python3
# LinBASIC Studio - Phase 0 (REA): writes docs/rea/vb6-feature-map.md from data/vb6
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
import json, os, sys
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..")
D = os.path.join(ROOT, "data", "vb6", "typelibs")
out = []
w = out.append

idx = json.load(open(f"{D}/index.json"))
L = {i["short"]: json.load(open(f"{D}/{i['short']}.json")) for i in idx}
w("# VB6 feature map (Phase 0, REA)\n")
w("Generated from a local Visual Basic 6.0 installation (VB6.EXE 6.0 SP6, MSVBVM60.DLL, OCX files, "
  "MSDN Library October 1998) by `tools/rea/dump_typelibs.py`, `tools/rea/dump_resources.py` and "
  "`tools/rea/make_feature_map.py`. Data: `data/vb6/`.\n")
w("## Type libraries\n")
w("| Library | File | Version | Types | Members | Help file |\n|---|---|---|---|---|---|")
for i in idx:
    w(f"| {i['library']} | {i['file']} | {i['version']} | {i['types']} | {i['members']} | {i['helpfile'] or '—'} |")


def ifaces(lib):
    T = {t["name"]: t for t in lib["types"]}
    for t in lib["types"]:
        if t["kind"] == "coclass" and t.get("implements") and "hidden" not in t["flags"] \
                and "restricted" not in t["flags"]:
            d = next((x["name"] for x in t["implements"] if x["default"] and not x["source"]), None)
            s = next((x["name"] for x in t["implements"] if x["source"]), None)
            yield t, T.get((d or "").split(".")[-1]), T.get((s or "").split(".")[-1])


def counts(di, ev):
    props, meth = set(), set()
    for m in (di or {}).get("members", []):
        if set(m.get("flags", [])) & {"hidden", "restricted"} or m["name"].startswith("_"):
            continue
        (props if m.get("invoke") in ("get", "let", "set") else meth).add(m["name"])
    return props, meth, [m["name"] for m in (ev or {}).get("members", [])]


w("\n## Intrinsic controls and objects (library VB)\n")
w("| Object | Properties | Methods | Events |\n|---|---|---|---|")
for t, di, ev in ifaces(L["VB"]):
    p, m, e = counts(di, ev)
    w(f"| {t['name']} | {len(p)} | {len(m)} | {len(e)}" + (f": {', '.join(e[:12])}{' …' if len(e) > 12 else ''}" if e else "") + " |")
for short, title in [("MSComctlLib", "Common Controls (MSCOMCTL.OCX)"), ("MSComCtl2", "Common Controls 2 (MSCOMCT2.OCX)"),
                     ("ComCtl3", "CoolBar (COMCT332.OCX)"), ("MSComDlg", "CommonDialog (COMDLG32.OCX)"),
                     ("MSFlexGridLib", "MSFlexGrid (MSFLXGRD.OCX)"), ("RichTextLib", "RichTextBox (RICHTX32.OCX)"),
                     ("TabDlg", "SSTab (TABCTL32.OCX)")]:
    w(f"\n## {title}\n")
    w("| Control | Properties | Methods | Events |\n|---|---|---|---|")
    for t, di, ev in ifaces(L[short]):
        if "control" in t["flags"]:
            p, m, e = counts(di, ev)
            w(f"| {t['name']} | {len(p)} | {len(m)} | {len(e)} |")
V = L["VBA"]
w("\n## VBA runtime library (MSVBVM60.DLL)\n")
w("Every function of the VBA modules maps to an MSVBVM60 export (`rtc*`), kept in `VBA.json` as `export`. "
  "Names `_B_str_X` / `_B_var_X` are the `X$` and `X` forms of the same function.\n")
w("| Module | Functions | Examples |\n|---|---|---|")
for t in V["types"]:
    if t["kind"] == "module" and t.get("members"):
        ms = [m for m in t["members"] if m["kind"] == "func"]
        w(f"| {t['name']} | {len(ms)} | " + ", ".join(f"{m['name']} → {m.get('export')}" for m in ms[:3]) + " |")
w("\nEnumerations: " + ", ".join(t["name"] for t in V["types"] if t["kind"] == "enum") + ". Objects: ErrObject, Collection.\n")
R = L["VBRUN"]
en = [t for t in R["types"] if t["kind"] in ("enum", "module")]
w(f"## VBRUN constants\n\n{len(en)} enumerations ({sum(len(t.get('members', [])) for t in en)} constants), e.g. "
  + ", ".join(t["name"] for t in en[:10]) + " …\n")
ide = json.load(open(os.path.join(ROOT, "data", "vb6", "ide_commands.json")))
M = ide["menus"]
rng = lambda a, b: [M[str(k)].replace("&", "") for k in range(a, b) if str(k) in M]
w("## IDE (VB6IDE.DLL)\n")
w("| Menu | Commands |\n|---|---|")
for n, a, b in [("File", 19144, 19176), ("Edit", 19176, 19208), ("View", 19208, 19240), ("Project", 19240, 19272),
                ("Format", 19272, 19300), ("Debug", 19300, 19336), ("Run", 19336, 19360), ("Tools", 19360, 19432),
                ("Window", 19432, 19464), ("Help", 19464, 19480)]:
    w(f"| {n} | {', '.join(c for c in rng(a, b) if 'Do not localize' not in c)} |")
w("\nToolbars: Standard, Edit, Debug, Form Editor. Context menus: " + ", ".join(rng(19024, 19061)) + ".\n")
w("Dialogs: " + ", ".join(sorted(set(v for v in ide["dialogs"].values()
                                     if v and v not in ("Microsoft Visual Basic", "Visual Basic", "None")))) + ".\n")
w("Keyboard shortcuts (main table): " + ", ".join(a["keys"] for a in ide["accelerators"]["1302"]) + ".\n")
E = json.load(open(os.path.join(ROOT, "data", "vb6", "runtime_errors.json")))
A, B = E["vba"], E["vb"]
w(f"## Run-time errors\n\n{len(A)} VBA errors (e.g. 13 {A['13']}, 91 {A['91']}) and {len(B)} VB errors of forms "
  f"and controls (e.g. 340 {B['340']}, 360 {B['360']}, 364 {B['364']}), with the messages MSVBVM60 gives "
  "`Err.Description` (`runtime_errors.json`). MSVBVM60 keeps the VBA messages as string ids 10000+i in "
  "ascending error order and the VB messages as string id 1000 + error number.\n")
w("## File formats\n")
w("`tools/rea/check_samples.py` on the VB98 samples of the MSDN library: every .frm/.ctl/.dob read and written back "
  "byte for byte; `.frx` pictures (BMP, ICO, WMF), ImageList images (`StdPicture` CLSID prefix) and `$` strings "
  "(u32 length) and short `Text` (u8 length) decoded on 75 files. `List`/`ItemData` taken from a form saved by VB6 "
  "itself (a test form saved in the VB6 IDE): `[u16 count][u16 max length]` then `[u16 length][text]` per item; "
  "ItemData numbers are stored as text. Bindings, OleObjectBlob and PropertyPages records are kept as raw bytes.\n")
w("## Open items\n")
w("* Types from libraries not loaded (msdatsrc, MSSTDFMT, DAO) appear as `file:{GUID}` in the JSON.")
open(os.path.join(ROOT, "docs", "rea", "vb6-feature-map.md"), "w", encoding="utf-8").write("\n".join(out) + "\n")
print("docs/rea/vb6-feature-map.md:", len(out), "lines")
