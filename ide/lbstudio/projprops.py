"""Project > Properties without Qt (phase 8.5): what the five pages of VB6's dialog show for a project and the
lines of the .vbp they change, in the order and with the values the VB6 IDE writes (VB6 oracle DEV-8/pp1).

A .vbp keeps its lines; a key the dialog adds goes where VB6 puts it (ORDER), a key it changes stays where it is.
What VB6 writes on OK:
  General   always (it is the page the dialog opens on): HelpFile="" even empty, Name, Startup, HelpContextID,
            Description (only when there is one), NoControlUpgrade=1 / RequireLicenseKey=1 (only when set),
            Unattended, Retained, ThreadPerObject, MaxNumberOfThreads; ThreadingModel only when the .vbp has it or
            it was changed (a project made an ActiveX DLL from a Standard EXE has none: Single Threaded)
  Make      once the page was shown: IconForm (when there is a form) and Title (the name of the .vbp file),
            the version, the six Version* texts (only those that are not empty), Command32, CondComp (only when
            there is one), RemoveUnusedControlInfo=0 (only when off)
  Compile   CompilationType (0 native, -1 P-Code), OptimizationType, the -1 / 0 switches, DllBaseAddress (only
            when it is not &H11000000)
  Component StartMode, ServerSupportFiles (1 / 0), CompatibleMode ("0" none, "1" project, "2" binary),
            CompatibleEXE32 (kept when No Compatibility is chosen), VersionCompatible32="1" (binary only)
  Debugging once the page was shown: DebugStartupOption (0 wait, 1 component, 2 program, 3 browser),
            DebugStartupComponent (the component, the program or the URL; no quotes), UseExistingBrowser=0 (only
            when off)
A new Project Type: Startup "Sub Main" (and IconForm = the form that was the startup object) when an EXE becomes
a component, StartMode 0 / 1 and CompatibleMode "0" / "1" for it, the Debugging lines dropped; on OK the message
"Project 'Start Mode' property has been changed."
"""
import os

# the keys of the top of a .vbp in the order VB6 writes them; the member lines (Form=, Class= ...), Reference= and
# Object= come before IconForm
ORDER = ["Type", "ResFile32", "IconForm", "Startup", "HelpFile", "RequireLicenseKey", "NoControlUpgrade", "Title", "ExeName32",
         "Command32", "Name", "HelpContextID", "Description", "CompatibleMode", "CompatibleEXE32",
         "VersionCompatible32", "MajorVer", "MinorVer", "RevisionVer", "AutoIncrementVer", "ServerSupportFiles",
         "DllBaseAddress", "VersionComments", "VersionCompanyName", "VersionFileDescription",
         "VersionLegalCopyright", "VersionLegalTrademarks", "VersionProductName", "CondComp", "CompilationType",
         "OptimizationType", "FavorPentiumPro(tm)", "CodeViewDebugInfo", "NoAliasing", "BoundsCheck", "OverflowCheck",
         "FlPointCheck", "FDIVCheck", "UnroundedFP", "StartMode", "Unattended", "Retained", "ThreadPerObject",
         "MaxNumberOfThreads", "ThreadingModel", "RemoveUnusedControlInfo", "DebugStartupOption",
         "DebugStartupComponent", "UseExistingBrowser"]
_RANK = {k.lower(): i for i, k in enumerate(ORDER)}

VERSION_INFO = [("Comments", "VersionComments"), ("Company Name", "VersionCompanyName"),
                ("File Description", "VersionFileDescription"), ("Legal Copyright", "VersionLegalCopyright"),
                ("Legal Trademarks", "VersionLegalTrademarks"), ("Product Name", "VersionProductName")]
# Advanced Optimizations: the check box and its key (-1 when checked)
ADVANCED = [("Assume No Aliasing", "NoAliasing"), ("Remove Array Bounds Checks", "BoundsCheck"),
            ("Remove Integer Overflow Checks", "OverflowCheck"), ("Remove Floating Point Error Checks", "FlPointCheck"),
            ("Allow Unrounded Floating Point Operations", "UnroundedFP"),
            ("Remove Safe Pentium(tm) FDIV Checks", "FDIVCheck")]
TYPES = [("Exe", "Standard EXE"), ("OleExe", "ActiveX EXE"), ("OleDll", "ActiveX DLL"), ("Control", "ActiveX Control")]
DLL_BASE = "&H11000000"
UI_KINDS = ("Form", "UserControl", "PropertyPage", "UserDocument")   # Unattended Execution is off with them
START_MODE_MSG = "Project 'Start Mode' property has been changed."


def compat_error(path):
    return f"Unable to set the version compatible component: {path}"


def _rank(key):
    return _RANK.get(key.strip().lower())


def put(vbp, key, value, quote=False):
    """key=value: in place when the .vbp has the key, otherwise where VB6 writes it"""
    if quote:
        value = '"' + value.replace('"', '""') + '"'
    for e in vbp.entries:
        if e[1] is not None and e[0].strip().lower() == key.lower():
            e[1] = value
            return
    r = _rank(key)
    pos, cur = len(vbp.entries), -1
    if r is not None:
        pos = 0
        for i, (k, v) in enumerate(vbp.entries):
            kr = _rank(k) if v is not None else None
            cur = kr if kr is not None else max(cur, 0)      # the member lines go with Type
            if cur < r:
                pos = i + 1
    vbp.entries.insert(pos, [key, value])


def drop(vbp, key):
    vbp.entries = [e for e in vbp.entries if not (e[1] is not None and e[0].strip().lower() == key.lower())]


def has(vbp, key):
    return any(v is not None and k.strip().lower() == key.lower() for k, v in vbp.entries)


def _num(v, default=0):
    try:
        return int(str(v).strip().strip('"'))
    except (TypeError, ValueError):
        return default


def _on(vbp, key):
    return _num(vbp.get(key), 0) != 0


def norm_type(t):
    return {"exe": "Exe", "oleexe": "OleExe", "oledll": "OleDll", "control": "Control"}.get((t or "Exe").lower(), "Exe")


def startup_items(ptype, forms):
    """the Startup Object list: Sub Main and the forms of a Standard EXE, (None) and Sub Main of a component
    (DEV-8/pp1 stages 15, 23)"""
    return ["Sub Main"] + list(forms) if ptype == "Exe" else ["(None)", "Sub Main"]


def read(vbp, modules=(), path=None):
    """the values the dialog shows for a project: modules are (kind, name) of its members"""
    g = vbp.get
    ptype = norm_type(g("Type"))
    forms = [n for k, n in modules if k == "Form"]
    ucs = [n for k, n in modules if k in ("UserControl", "UserDocument")]
    stem = os.path.splitext(os.path.basename(path))[0] if path else (g("Name") or "Project1")
    s = {
        "type": ptype, "name": g("Name") or "Project1", "startup": g("Startup") or ("Sub Main" if ptype == "Exe" else "(None)"),
        "helpfile": g("HelpFile") or "", "helpctx": g("HelpContextID") or "0", "desc": g("Description") or "",
        "unattended": _on(vbp, "Unattended"), "upgrade": not _on(vbp, "NoControlUpgrade"),
        "license": _on(vbp, "RequireLicenseKey"), "retained": _on(vbp, "Retained"),
        "threading": _num(g("ThreadingModel"), 0), "has_threading": has(vbp, "ThreadingModel"),
        "per_object": _on(vbp, "ThreadPerObject"), "threads": max(1, _num(g("MaxNumberOfThreads"), 1)),
        "major": _num(g("MajorVer")), "minor": _num(g("MinorVer")), "revision": _num(g("RevisionVer")),
        "autoinc": _on(vbp, "AutoIncrementVer"), "title": g("Title") if has(vbp, "Title") else stem,
        "iconform": g("IconForm") or (g("Startup") if g("Startup") in forms else (forms[0] if forms else "")),
        "info": {key: g(key) or "" for _, key in VERSION_INFO}, "command": g("Command32") or "",
        "condcomp": g("CondComp") or "", "remove_unused": not has(vbp, "RemoveUnusedControlInfo") or
        _on(vbp, "RemoveUnusedControlInfo"),
        "pcode": _num(g("CompilationType")) != 0, "opt": _num(g("OptimizationType")),
        "pentium": _on(vbp, "FavorPentiumPro(tm)"), "symdebug": _on(vbp, "CodeViewDebugInfo"),
        "advanced": {key: _on(vbp, key) for _, key in ADVANCED}, "dllbase": g("DllBaseAddress") or DLL_BASE,
        "startmode": _num(g("StartMode")), "remote": _on(vbp, "ServerSupportFiles"),
        "compat": _num(g("CompatibleMode")), "compat_path": g("CompatibleEXE32") or "",
        "use_existing": not has(vbp, "UseExistingBrowser") or _on(vbp, "UseExistingBrowser"),
        "forms": forms, "components": ucs,
        "ui": any(k in UI_KINDS for k, _ in modules),
    }
    opt = g("DebugStartupOption")
    s["debug"] = _num(opt) if opt is not None else (1 if ptype == "Control" and ucs else 0)
    comp = g("DebugStartupComponent") or ""
    s["debug_component"] = comp if s["debug"] == 1 else (ucs[0] if ucs else "")
    s["debug_program"] = comp if s["debug"] == 2 else ""
    s["debug_url"] = comp if s["debug"] == 3 else ""
    return s


def change_type(s, new):
    """a new Project Type chosen in the dialog: what VB6 changes with it (DEV-8/pp1 stages 15-20)"""
    old = s["type"]
    if new == old:
        return
    if new != "Exe" and s["startup"] not in ("Sub Main", "(None)"):
        if not s.get("iconform_set"):
            s["iconform"] = s["startup"]
            s["iconform_from_type"] = True
        s["startup"] = "Sub Main"
    s["type"] = new
    s["startmode"] = 0 if new == "Exe" else 1
    s["compat"] = 0 if new == "Exe" else 1


def enabled(s):
    """what is on in the dialog for these values (DEV-8/pp1)"""
    t = s["type"]
    comp = t != "Exe"
    return {
        "unattended": t in ("OleExe", "OleDll") and not s["ui"],
        "license": t == "Control",
        "retained": t == "OleDll" and s["unattended"] and not s["ui"],
        "threading": t in ("OleDll", "Control"),
        "thread_mode": t == "OleExe",
        "threads": t == "OleExe" and not s["per_object"],
        "native": not s["pcode"],
        "dllbase": t in ("OleDll", "Control"),
        "startmode": t == "OleExe",
        "remote": comp,
        "compat": comp,
        "compat_path": comp and s["compat"] != 0,
        "debugging": comp,
        "start_component": comp and bool(s["components"]),
    }


def compat_problem(s, project_dir):
    """what Project / Binary Compatibility cannot use, as VB6 names it in its message ("Unable to set the version
    compatible component: <path>"), or None: Binary needs a file that is there, even an empty box is refused;
    Project only a file that is there when one is named (DEV-8/pp1 stages 11, 22, 23)"""
    p = s["compat_path"].strip()
    if s["type"] == "Exe" or s["compat"] == 0:
        return None
    if not p:
        return "" if s["compat"] == 2 else None
    f = os.path.normpath(os.path.join(project_dir or ".", p.replace("\\", "/")))
    return None if os.path.exists(f) else p


def write(vbp, old, s, visited):
    """the lines VB6 writes on OK: old are the values the dialog opened with, s the values on OK, visited the
    pages that were shown ("general", "make", "compile", "component", "debugging"). Returns True when the
    Start Mode changed with the Project Type (VB6 says so)"""
    typed = s["type"] != old["type"]
    put(vbp, "Type", s["type"])
    # General
    put(vbp, "Name", s["name"].strip() or "Project1", quote=True)
    put(vbp, "Startup", s["startup"], quote=True)
    put(vbp, "HelpFile", s["helpfile"], quote=True)
    if s["upgrade"]:
        drop(vbp, "NoControlUpgrade")
    else:
        put(vbp, "NoControlUpgrade", "1")
    if s["license"] and s["type"] == "Control":
        put(vbp, "RequireLicenseKey", "1")
    else:
        drop(vbp, "RequireLicenseKey")
    put(vbp, "HelpContextID", str(_num(s["helpctx"])), quote=True)
    if s["desc"]:
        put(vbp, "Description", s["desc"], quote=True)
    else:
        drop(vbp, "Description")
    if "make" in visited or (typed and s.get("iconform_from_type") and not has(vbp, "IconForm")):
        if s["iconform"] and s["forms"]:
            put(vbp, "IconForm", s["iconform"], quote=True)
    if "make" in visited:
        put(vbp, "Title", s["title"], quote=True)
        put(vbp, "Command32", s["command"], quote=True)
        for _, key in VERSION_INFO:
            if s["info"][key]:
                put(vbp, key, s["info"][key], quote=True)
            else:
                drop(vbp, key)
        if s["condcomp"]:
            put(vbp, "CondComp", s["condcomp"], quote=True)
        else:
            drop(vbp, "CondComp")
        if s["remove_unused"]:
            drop(vbp, "RemoveUnusedControlInfo")
        else:
            put(vbp, "RemoveUnusedControlInfo", "0")
    else:
        put(vbp, "Command32", s["command"], quote=True)
    put(vbp, "CompatibleMode", str(s["compat"]), quote=True)
    if s["compat_path"]:
        put(vbp, "CompatibleEXE32", s["compat_path"], quote=True)
    else:
        drop(vbp, "CompatibleEXE32")
    if s["compat"] == 2:
        put(vbp, "VersionCompatible32", "1", quote=True)
    else:
        drop(vbp, "VersionCompatible32")
    put(vbp, "MajorVer", str(s["major"]))
    put(vbp, "MinorVer", str(s["minor"]))
    put(vbp, "RevisionVer", str(s["revision"]))
    put(vbp, "AutoIncrementVer", "1" if s["autoinc"] else "0")
    put(vbp, "ServerSupportFiles", "1" if s["remote"] else "0")
    if s["dllbase"].strip().upper() != DLL_BASE.upper():
        put(vbp, "DllBaseAddress", s["dllbase"].strip())
    else:
        drop(vbp, "DllBaseAddress")
    # Compile
    put(vbp, "CompilationType", "-1" if s["pcode"] else "0")
    put(vbp, "OptimizationType", str(s["opt"]))
    put(vbp, "FavorPentiumPro(tm)", "-1" if s["pentium"] else "0")
    put(vbp, "CodeViewDebugInfo", "-1" if s["symdebug"] else "0")
    for _, key in ADVANCED:
        put(vbp, key, "-1" if s["advanced"][key] else "0")
    # Component and the threads
    put(vbp, "StartMode", str(s["startmode"]))
    put(vbp, "Unattended", "-1" if s["unattended"] else "0")
    put(vbp, "Retained", "1" if s["retained"] else "0")
    put(vbp, "ThreadPerObject", "-1" if s["per_object"] else "0")
    put(vbp, "MaxNumberOfThreads", str(s["threads"]))
    if s["has_threading"] or s["threading"] != old["threading"]:
        put(vbp, "ThreadingModel", str(s["threading"]))
    # Debugging
    if "debugging" in visited:
        put(vbp, "DebugStartupOption", str(s["debug"]))
        comp = {1: s["debug_component"], 2: s["debug_program"], 3: s["debug_url"]}.get(s["debug"], "")
        if comp:
            put(vbp, "DebugStartupComponent", comp)
        else:
            drop(vbp, "DebugStartupComponent")
        if s["use_existing"]:
            drop(vbp, "UseExistingBrowser")
        else:
            put(vbp, "UseExistingBrowser", "0")
    elif typed:
        for key in ("DebugStartupOption", "DebugStartupComponent", "UseExistingBrowser"):
            drop(vbp, key)
    return typed and s["startmode"] != old["startmode"]
