#!/usr/bin/env python3
# LinBASIC Studio - Phase 0: check the VB6 file readers against real VB6 projects
# Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later.
"""Usage: python3 tools/rea/check_samples.py <folder with .frm/.frx/.ctl/... files>

For every designer file: parse it, write it back and compare byte for byte, then
decode every binary-store reference and report the record layouts found."""
import collections, os, struct, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "compiler"))
from lbvc.project import load_source, FrxStore            # noqa: E402
from lbvc.project.vbp import load_project                 # noqa: E402

STORE_EXT = {".frm": ".frx", ".ctl": ".ctx", ".dob": ".dox", ".pag": ".pgx"}


def main(root):
    ok = bad = 0
    kinds = collections.Counter()
    problems = []
    for dp, _, fs in os.walk(root):
        for fn in sorted(fs):
            ext = os.path.splitext(fn)[1].lower()
            path = os.path.join(dp, fn)
            if ext == ".vbp":
                p = load_project(path)
                if p.dumps().encode(p.encoding, "surrogateescape") != open(path, "rb").read():
                    problems.append(f"{path}: vbp round trip differs")
                continue
            if ext not in (".frm", ".ctl", ".dob", ".pag", ".cls", ".bas"):
                continue
            try:
                sf = load_source(path)
            except Exception as e:
                bad += 1
                problems.append(f"{path}: {e}")
                continue
            if sf.dumps().encode(sf.encoding, "surrogateescape") != open(path, "rb").read():
                bad += 1
                problems.append(f"{path}: round trip differs")
                continue
            ok += 1
            refs = sf.frx_refs()
            if not refs:
                continue
            store = None
            for f2 in fs:
                if f2.lower() == os.path.splitext(fn)[0].lower() + STORE_EXT.get(ext, ".frx"):
                    store = os.path.join(dp, f2)
            if not store:
                problems.append(f"{path}: binary store missing")
                continue
            st = FrxStore.load(store, [r.value.offset for _, r in refs])
            for ctl, r in refs:
                try:
                    v = st.decode(r.name, r.value)
                except Exception as e:
                    problems.append(f"{fn} {ctl.name}.{r.name}@{r.value.offset:04X}: {e}")
                    continue
                if isinstance(v, tuple):
                    kinds[f"picture:{v[0]}"] += 1
                    if v[1] and v[0] is None:
                        problems.append(f"{fn} {ctl.name}.{r.name}: unknown image {v[1][:8].hex()}")
                elif isinstance(v, list):
                    kinds[f"{r.name.lower()}"] += 1
                elif isinstance(v, str):
                    kinds["string$" if r.value.is_string else f"string:{r.name}"] += 1
                else:
                    kinds[f"raw:{ctl.type_name}.{r.name}"] += 1
    print(f"designer/module files: {ok} round trip OK, {bad} failed")
    for k, n in sorted(kinds.items()):
        print(f"  {k:40} {n}")
    for p in problems:
        print("  !", p)


if __name__ == "__main__":
    main(sys.argv[1])
