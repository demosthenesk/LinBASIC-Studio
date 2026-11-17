"""The VB6 object model from data/vb6/typelibs (Phase 0, REA): the events of forms and
controls, for the Procedure box of the code window."""
import json
import os

from . import paths

_vb = None


def _types():
    global _vb
    if _vb is None:
        try:
            with open(os.path.join(paths.VB6_DATA, "typelibs", "VB.json"), encoding="utf-8") as f:
                _vb = {t["name"].lower(): t for t in json.load(f)["types"]}
        except (OSError, ValueError, KeyError):
            _vb = {}
    return _vb


def _short(t):
    t = t.split("\\")[-1]
    return t.split(".")[-1]


_libs = {}


def _lib_type(lib, name):
    if lib not in _libs:
        try:
            with open(os.path.join(paths.VB6_DATA, "typelibs", lib + ".json"), encoding="utf-8") as f:
                _libs[lib] = {t["name"].lower(): t for t in json.load(f)["types"]}
        except (OSError, ValueError, KeyError):
            _libs[lib] = {}
    return _libs[lib].get(name.lower())


_ocx_index = None


def ocx_events_iface(control):
    """(library, its events interface) of a control of the type libraries of the components"""
    global _ocx_index
    if _ocx_index is None:
        _ocx_index = {}
        folder = os.path.join(paths.VB6_DATA, "typelibs")
        try:
            names = sorted(os.listdir(folder))
        except OSError:
            names = []
        for fn in names:
            lib = fn[:-5]
            if not fn.endswith(".json") or lib in ("VB", "VBA", "VBRUN", "VBIDE", "VBInternal", "stdole", "index"):
                continue
            try:
                with open(os.path.join(folder, fn), encoding="utf-8") as f:
                    types = json.load(f).get("types", [])
            except (OSError, ValueError):
                continue
            for t in types:
                if t.get("kind") == "coclass" and "control" in t.get("flags", []) and "hidden" not in t.get("flags", []):
                    src = next((i["name"] for i in t.get("implements", []) if i.get("source")), None)
                    if src:
                        _ocx_index.setdefault(t["name"].lower(), (lib, src))
    hit = _ocx_index.get(control.lower())
    if not hit:
        return None
    t = _lib_type(hit[0], hit[1])
    return (hit[0], t) if t else None


# the objects of the designers (.dsr) in the code window: their events beyond the type library
DESIGNER_EVENTS = {
    "dataenvironment": [("Initialize", ""), ("Terminate", "")],
    "datareport": [("Activate", ""), ("Deactivate", ""), ("Initialize", ""), ("Terminate", ""), ("Resize", ""),
                   ("QueryClose", "Cancel As Integer, CloseMode As Integer")],
}


def events(type_name):
    """[(event, "params as VB6 writes them in a new procedure")] of a form or control type; a type
    "Lib.Interface" (ADODB.RecordsetEvents) is the event interface of that type library"""
    out = list(DESIGNER_EVENTS.get(type_name.lower(), []))
    if type_name.lower() == "datareport":
        type_name = "MSDataReportLib.DMSDataReportEvents"
    lib = None
    if "." in type_name:
        lib, _, iface = type_name.partition(".")
        t = _lib_type(lib, iface)
    else:
        t = _types().get(type_name.lower() + "events")
        if not t:                               # a control of a component (.ocx): its source interface
            hit = ocx_events_iface(type_name)
            if hit:
                lib, t = hit
    if not t:
        return out
    for m in t.get("members", []):
        if m.get("kind") != "func":
            continue
        ps = []
        for p in m.get("params", []):
            ty, byval = _short(p["type"]), not p.get("byref", True)
            if lib is not None:
                # as VB6 writes them: the types of the library qualified, its objects ByVal
                owner = next((x for x in (lib, "ADODB") if _lib_type(x, ty) is not None), None)
                lt = _lib_type(owner, ty) if owner else None
                if lt is not None:
                    if lt.get("kind") in ("dispatch", "interface", "coclass"):
                        byval = True
                        ty = ty.lstrip("_")
                    ty = f"{owner}.{ty}"
            ps.append(("ByVal " if byval else "") + f"{p['name']} As {ty}")
        out.append((m["name"], ", ".join(ps)))
    return out
