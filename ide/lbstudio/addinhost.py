"""The add-ins of the user (phase 12.5): an add-in project of LinBASIC Studio (an ActiveX DLL with the Connect designer)
is made as a program; Make registers it (S["user_addins"]) and the Add-In Manager loads it: the IDE starts the program
and talks with it (runtime/addin.c: TCP on 127.0.0.1, a key, one message a line) - it calls the events of Connect
(OnConnection, OnDisconnection, OnStartupComplete ...), the add-in uses the objects of the IDE (vbide.py) and gets the
events of its command bar controls. VB6 loads the add-in as a COM DLL into its own process; here it is a process of
its own, without COM (decision of the user, 2026-11-27).

The protocol (as lbdbinfo -server writes values: e, m, n, s:, h:, i:, f:, d:, c:, t:, b:, y:, o:<id>[:<TypeName>]):
    the add-in:  K <key> | C <id> <M|G|L|S> <member> <arg>... (-> OK <value> | ER <n> <text>) | R <id>
    the IDE:     X connect <app> <mode> <inst> | X disconnect <mode> | X startup | X update | X shutdown |
                 X event <id> <event> <arg>... | X quit                (-> D <arg>... | ER <n> <text>)
"""
import os
import secrets
from contextlib import contextmanager

from PyQt6.QtCore import QObject, QProcess, QProcessEnvironment, QElapsedTimer, QCoreApplication, QEventLoop
from PyQt6.QtNetwork import QTcpServer, QHostAddress

from . import vbide
from .vbide import AddInError, NOTHING, VBObject, Out

# ext_ConnectMode, ext_DisconnectMode (AddInDesignerObjects)
CM_AFTERSTARTUP, CM_STARTUP, CM_EXTERNAL, CM_COMMANDLINE = 0, 1, 2, 3
DM_HOSTSHUTDOWN, DM_USERCLOSED = 0, 1
EMPTY = object()


def _esc(s):
    return s.replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n").replace("\r", "\\r")


def _unesc(s):
    out, i = [], 0
    while i < len(s):
        if s[i] == "\\" and i + 1 < len(s):
            out.append({"t": "\t", "n": "\n", "r": "\r"}.get(s[i + 1], s[i + 1]))
            i += 2
        else:
            out.append(s[i])
            i += 1
    return "".join(out)


# ---------------------------------------------------------------- the registry of the add-ins
def user_addins(main):
    return list(main.S["user_addins"] or [])


def register_addin(main, info):
    """Make of an add-in project: it goes into the Add-In Manager (again: its file and texts are updated, its Load
    Behavior kept)"""
    lst = user_addins(main)
    for a in lst:
        if a["progid"].lower() == info["progid"].lower():
            for k in ("exe", "vbp", "name", "description", "cmdsafe"):
                a[k] = info[k]
            break
    else:
        lst.append(info)
    main.S["user_addins"] = lst
    main.S.save()


def loaded_host(main, progid):
    h = getattr(main, "addin_hosts", {}).get(progid.lower())
    return h if h is not None and h.running() else None


def load_addin(main, progid, mode=CM_AFTERSTARTUP):
    """Loaded: the program of the add-in starts, then OnConnection"""
    if loaded_host(main, progid) is not None:
        return loaded_host(main, progid)
    info = next((a for a in user_addins(main) if a["progid"].lower() == progid.lower()), None)
    if info is None:
        return None
    h = AddInHost(main, info)
    if not h.start():
        main.addin_error(info, h.error)
        return None
    if not hasattr(main, "addin_hosts"):
        main.addin_hosts = {}
    main.addin_hosts[progid.lower()] = h
    install_file_hook(main)
    if mode in (CM_STARTUP, CM_COMMANDLINE):
        # VB6 (oracle evlog): OnConnection before the projects are there (VBProjects.Count 0), OnAddInsUpdate, then
        # the events of the projects loading (ItemRenamed "", ItemAdded of each component, ItemAdded), then
        # OnStartupComplete (load_startup_addins)
        h.hide_projects = True
        h.snap = Snap()
    else:
        h.snap = Snap.of(main)
    try:
        h.connect(mode)
    except AddInError as e:
        main.addin_error(info, f"{e.description} ({e.number})")
    if mode in (CM_STARTUP, CM_COMMANDLINE):
        try:
            if h.running():
                h.call("update")
        except AddInError as e:
            main.addin_error(info, f"{e.description} ({e.number})")
        h.hide_projects = False
        sync(h)
    return h


def unload_addin(main, progid, mode=DM_USERCLOSED):
    """Unloaded: OnDisconnection, then the program ends"""
    h = getattr(main, "addin_hosts", {}).pop(progid.lower(), None)
    if h is None:
        return
    try:
        if h.running():
            h.call("disconnect", mode)
    except AddInError as e:
        main.addin_error(h.info, f"{e.description} ({e.number})")
    h.stop()


def load_behavior_text(loaded, startup, cmdline):
    """the column Load Behavior of the Add-In Manager (VB6 oracles 12.5/addinmgr-*: "Startup / Command Line",
    "Startup / Loaded", all three "ALL")"""
    if loaded and startup and cmdline:
        return "ALL"
    return " / ".join(t for t, on in (("Startup", startup), ("Loaded", loaded), ("Command Line", cmdline)) if on)


# ---------------------------------------------------------------- the events of the IDE for the add-ins
# VBProjectsEvents, VBComponentsEvents, ReferencesEvents, VBControlsEvents, SelectedVBControlsEvents: each add-in has a
# picture of the IDE as it last saw it (Snap); after a change (a member of an add-in, Project Explorer, a designer ...)
# the differences are its events, in the order VB6 gives them (VB6 oracle: oracles/12.5/evlog-vb6*.txt).
# FileControlEvents come from the writing of the files (vbmodel.FILE_HOOK) and from Add / Remove File.
def ide_projects(main):
    g = getattr(main, "group", None)
    if g is not None:
        return list(g.projects)
    p = getattr(main, "project", None)
    return [p] if p is not None else []


def selected_module(main):
    """the component selected in the Project Explorer (None: a project, a folder, nothing)"""
    from PyQt6.QtCore import Qt
    tree = getattr(main, "tree", None)
    try:
        it = tree.currentItem() if tree is not None else None
        m = it.data(0, Qt.ItemDataRole.UserRole) if it is not None else None
    except RuntimeError:                        # (the tree being rebuilt)
        return None
    return m if m is not None and hasattr(m, "kind") else None


def form_roots(main):
    """the open form designers: root control -> (module, designer)"""
    out = {}
    for m, sw in list(getattr(main, "design_windows", {}).items()):
        d = sw.widget() if sw is not None else None
        if d is not None and hasattr(d, "root") and hasattr(d, "place_new"):
            out[id(d.root)] = (m, d)
    return out


class Snap:
    """what an add-in saw of the IDE"""

    def __init__(self):
        self.projects = []                      # [project]
        self.pnames = {}                        # id(project) -> name
        self.comps = {}                         # id(project) -> [(module, name)]
        self.refs = {}                          # id(project) -> [Reference line]
        self.selected = None
        self.active = None
        self.ctls = {}                          # id(form root) -> [(control, name, index)]
        self.sel_ctls = {}                      # id(form root) -> [control]

    @classmethod
    def of(cls, main):
        s = cls()
        for p in ide_projects(main):
            s.projects.append(p)
            s.pnames[id(p)] = p.name
            s.comps[id(p)] = [(m, m.name) for m in p.modules]
            s.refs[id(p)] = project_refs(p)
        s.selected = selected_module(main)
        s.active = getattr(main, "project", None)
        for k, (m, d) in form_roots(main).items():
            s.ctls[k] = [(c, c.name, c.index) for c in d.root.walk() if c is not d.root]
            s.sel_ctls[k] = [it.ctl for it in getattr(d, "selection", [])]
        return s


def project_refs(p):
    return [v for k, v in (p.vbp.entries if p.vbp else []) if v is not None and k.strip() == "Reference"]


def _arg(h, a):
    """an argument of an event: ("c", module) a VBComponent, ("p", project), ("r", line), ("x", control, module)"""
    if isinstance(a, tuple) and a:
        k = a[0]
        if k == "c":
            return h.vbe.wrap(vbide.VBComponent, a[1])
        if k == "p":
            return h.vbe.wrap(vbide.VBProject, a[1])
        if k == "r":
            return h.vbe.wrap(vbide.Reference, a[1], "Reference")
        if k == "x":
            return h.vbe.wrap(vbide.VBControl, a[1], a[2])
    return a


def _sources(h, typename, project=None, root=None):
    """the event objects of this type the add-in took (Events.VBComponentsEvents(Nothing): all the projects)"""
    out = []
    for o in list(h.vbe._wrap.values()):
        if not isinstance(o, vbide.SourceEvents) or o.TYPENAME != typename:
            continue
        k = getattr(o, "_key", None)
        if root is not None:
            if k is not root:
                continue
        elif isinstance(k, vbide.VBProject) and project is not None and k.p is not project:
            continue
        out.append(o)
    return out


def fire(h, typename, name, project, *args, root=None):
    """an event for one add-in; the values the add-in gave back (ByRef Cancel ...)"""
    backs = []
    for src in _sources(h, typename, project, root):
        if not h.running():
            break
        try:
            backs.append(h.fire_event(src, name, *[_arg(h, a) for a in args]))
        except AddInError as e:
            h.main.addin_error(h.info, f"{e.description} ({e.number})")
    return backs


def hosts(main):
    return [h for h in list(getattr(main, "addin_hosts", {}).values()) if h.running()]


def fire_all(main, typename, name, project, *args, root=None):
    out = []
    for h in hosts(main):
        out += fire(h, typename, name, project, *args, root=root)
    return out


def sync(h):
    """the events of what changed since the add-in last saw the IDE"""
    if not h.running() or getattr(h, "syncing", False) or getattr(h, "hide_projects", False):
        return
    h.syncing = True
    try:
        for _ in range(4):                      # (its event procedures may change the IDE again)
            if not _sync_once(h):
                break
    finally:
        h.syncing = False


def _sync_once(h):
    main = h.main
    old = getattr(h, "snap", None) or Snap()
    new = Snap.of(main)
    h.snap = new
    changed = False
    projs = new.projects
    sel_done = False
    act_done = False
    # projects that went: ItemRemoved, the References, the new active one, the files closed (oracle evlog)
    for p in old.projects:
        if p in projs:
            continue
        changed = True
        sel_done = True                         # (VB6: no ItemSelected when a project goes)
        fire(h, "VBProjectsEvents", "ItemRemoved", None, ("p", p))
        for line in old.refs.get(id(p), []):
            fire(h, "ReferencesEvents", "ItemRemoved", p, ("r", line))
        if new.active is not None and new.active in old.projects and new.active is not old.active and not act_done:
            fire(h, "VBProjectsEvents", "ItemActivated", None, ("p", new.active))
            act_done = True
        for m, _ in reversed(old.comps.get(id(p), [])):
            if m.path and os.path.exists(m.path):
                fire(h, "FileControlEvents", "AfterCloseFile", p, ("p", p), vbide.file_type(m), m.path, bool(m.dirty))
        if p.path:
            fire(h, "FileControlEvents", "AfterCloseFile", p, ("p", p), vbide.FT_PROJECT, p.path, bool(p.any_dirty()))
    # projects that came: (opened from a file: ItemRenamed "") the components, ItemAdded
    for p in projs:
        if p in old.projects:
            continue
        changed = True
        if p.path and os.path.exists(p.path):
            fire(h, "VBProjectsEvents", "ItemRenamed", None, ("p", p), "")
        for m in p.modules:
            fire(h, "VBComponentsEvents", "ItemAdded", p, ("c", m))
        fire(h, "VBProjectsEvents", "ItemAdded", None, ("p", p))
        if p is new.active:
            act_done = True                     # (a new project is the active one without ItemActivated)
        if new.selected is not None and new.selected in p.modules:
            sel_done = True                     # (nor ItemSelected for its component)
    for p in projs:
        if p not in old.projects:
            continue
        if old.pnames.get(id(p)) != p.name:
            changed = True
            fire(h, "VBProjectsEvents", "ItemRenamed", None, ("p", p), old.pnames.get(id(p), ""))
        before = old.comps.get(id(p), [])
        had = {id(m): n for m, n in before}
        for m, n in before:
            if m not in p.modules:
                changed = True
                fire(h, "VBComponentsEvents", "ItemRemoved", p, ("c", m))
        for m in p.modules:
            if id(m) not in had:
                changed = True
                # VB6: a code module is selected first, a form after its ItemAdded (oracle evlog)
                sel_first = m is new.selected and new.selected is not old.selected and not m.has_designer
                if sel_first:
                    fire(h, "VBComponentsEvents", "ItemSelected", p, ("c", m))
                    sel_done = True
                fire(h, "VBComponentsEvents", "ItemAdded", p, ("c", m))
            elif had[id(m)] != m.name:
                changed = True
                w = h.vbe._wrap.get((vbide.VBComponent, id(m)))
                if w is None or not w.pending:
                    fire(h, "VBComponentsEvents", "ItemRenamed", p, ("c", m), had[id(m)])
        old_refs = old.refs.get(id(p), [])
        cur = new.refs.get(id(p), [])
        for line in cur:
            if line not in old_refs:
                changed = True
                fire(h, "ReferencesEvents", "ItemAdded", p, ("r", line))
        for line in old_refs:
            if line not in cur:
                changed = True
                fire(h, "ReferencesEvents", "ItemRemoved", p, ("r", line))
    if new.selected is not None and new.selected is not old.selected and not sel_done:
        changed = True
        fire(h, "VBComponentsEvents", "ItemSelected", new.selected.project, ("c", new.selected))
    if new.active is not None and new.active is not old.active and old.active in projs and not act_done \
            and not getattr(h, "quiet_active", False):
        changed = True
        fire(h, "VBProjectsEvents", "ItemActivated", None, ("p", new.active))
    h.quiet_active = False
    # the controls of the forms the add-in watches
    roots = form_roots(main)
    for k, (m, d) in roots.items():
        root = d.root
        if not _sources(h, "VBControlsEvents", root=root) and not _sources(h, "SelectedVBControlsEvents", root=root):
            continue
        before = old.ctls.get(k)
        if before is None:
            continue
        cur = new.ctls.get(k, [])
        bids = {id(c): (n, i) for c, n, i in before}
        cids = {id(c) for c, _, _ in cur}
        for c, n, i in before:
            if id(c) not in cids:
                changed = True
                fire(h, "VBControlsEvents", "ItemRemoved", None, ("x", c, m), root=root)
        for c, n, i in cur:
            if id(c) not in bids:
                changed = True
                fire(h, "VBControlsEvents", "ItemAdded", None, ("x", c, m), root=root)
            elif bids[id(c)] != (n, i):
                changed = True
                on, oi = bids[id(c)]
                fire(h, "VBControlsEvents", "ItemRenamed", None, ("x", c, m), on, -1 if oi is None else oi, root=root)
        if not getattr(main, "_addin_api", 0):  # (VB6: no selection events for the members of an add-in)
            osel = old.sel_ctls.get(k, [])
            nsel = new.sel_ctls.get(k, [])
            for c in osel:
                if c not in nsel and id(c) in cids:
                    changed = True
                    fire(h, "SelectedVBControlsEvents", "ItemRemoved", None, ("x", c, m), root=root)
            for c in nsel:
                if c not in osel:
                    changed = True
                    fire(h, "SelectedVBControlsEvents", "ItemAdded", None, ("x", c, m), root=root)
    if changed:
        h.snap = Snap.of(main)
    return changed


def sync_all(main):
    if getattr(main, "_addin_hold", 0):
        return
    for h in hosts(main):
        sync(h)


@contextmanager
def hold(main):
    """a change in steps: the events after the last one"""
    main._addin_hold = getattr(main, "_addin_hold", 0) + 1
    try:
        yield
    finally:
        main._addin_hold -= 1


def quiet_active(main):
    """the active project changed without ItemActivated (VB6: ActiveVBProject set by an add-in, a new project)"""
    for h in hosts(main):
        h.quiet_active = True


@contextmanager
def api_change(main):
    main._addin_api = getattr(main, "_addin_api", 0) + 1
    try:
        yield
    finally:
        main._addin_api -= 1


def file_event(main, name, project, *args):
    """a FileControlEvents event for every add-in; True when one of them cancelled it (Cancel / CancelDefault)"""
    backs = fire_all(main, "FileControlEvents", name, project, ("p", project), *args)
    return any(bool(b[-1]) for b in backs if b)


def _file_hook(event, project, m, target):
    """vbmodel.FILE_HOOK: the writing of a file of a project (m None: the .vbp). before: the events before it (a file
    never written: DoGetNewFileName, RequestChangeFileName; RequestWriteFile) - False when an add-in cancelled it;
    after: AfterWriteFile (AfterChangeFileName for a new one)"""
    main = _HOOK_MAIN[0]
    if main is None or not hosts(main):
        return True
    ft = vbide.file_type(m) if m is not None else vbide.FT_PROJECT
    name = m.name if m is not None else project.name
    if event == "before":
        obj = m if m is not None else project     # never written (a new one), even if a file of its name is there
        new = getattr(obj, "never_saved", False) or not os.path.exists(target)
        setattr(m if m is not None else project, "_lb_was_new", new)
        if new:
            if not getattr(obj, "_lb_getnew_done", False):     # (Save File As fired it before its dialog)
                file_event(main, "DoGetNewFileName", project, ft, name, name, False)
            obj._lb_getnew_done = False
            if file_event(main, "RequestChangeFileName", project, ft, target, name, False):
                return False
        return not file_event(main, "RequestWriteFile", project, target, False)
    if event == "after":
        file_event(main, "AfterWriteFile", project, ft, target, 0)
        if getattr(m if m is not None else project, "_lb_was_new", False):
            file_event(main, "AfterChangeFileName", project, ft, target, name)
    return True


_HOOK_MAIN = [None]


def install_file_hook(main):
    from . import vbmodel
    _HOOK_MAIN[0] = main
    vbmodel.FILE_HOOK = _file_hook


def close_events(main):
    """the IDE closes (after OnBeginShutdown and its question): for each project ItemRemoved, its References, its
    files closed (VB6, oracle evlog-vb6-shutdown1)"""
    for h in hosts(main):
        snap = getattr(h, "snap", None) or Snap.of(main)
        for p in snap.projects:
            fire(h, "VBProjectsEvents", "ItemRemoved", None, ("p", p))
            for line in snap.refs.get(id(p), []):
                fire(h, "ReferencesEvents", "ItemRemoved", p, ("r", line))
            for m, _ in reversed(snap.comps.get(id(p), [])):
                if m.path and os.path.exists(m.path):
                    fire(h, "FileControlEvents", "AfterCloseFile", p, ("p", p), vbide.file_type(m), m.path, bool(m.dirty))
            if p.path:
                fire(h, "FileControlEvents", "AfterCloseFile", p, ("p", p), vbide.FT_PROJECT, p.path, bool(p.any_dirty()))
        h.snap = Snap()


# ---------------------------------------------------------------- the program of an add-in
class AddInHost(QObject):
    def __init__(self, main, info):
        super().__init__(main)
        self.main = main
        self.info = info
        self.error = ""
        self.proc = None
        self.sock = None
        self.buf = b""
        self.waiting = 0
        self.vbe = vbide.VBE(main, self)
        self.ids = {}                           # id(object) -> its id for the add-in
        self.objs = {}                          # id -> object
        self.next_id = 1
        self.controls = []                      # the command bar controls the add-in added (gone on unload)
        self.hide_projects = False
        self.snap = None
        self.syncing = False
        self.quiet_active = False

    # ------------------------------------------------------------ the process
    def running(self):
        return self.proc is not None and self.proc.state() != QProcess.ProcessState.NotRunning and self.sock is not None

    def start(self, timeout=15000):
        exe = self.info.get("exe", "")
        if not exe or not os.path.exists(exe):
            self.error = f"The file of the add-in is not there: {exe}"
            return False
        self.server = QTcpServer(self)
        if not self.server.listen(QHostAddress(QHostAddress.SpecialAddress.LocalHost), 0):
            self.error = "can't listen for the add-in"
            return False
        key = secrets.token_hex(16)
        env = QProcessEnvironment.systemEnvironment()
        env.insert("LBADDIN_PORT", str(self.server.serverPort()))
        env.insert("LBADDIN_KEY", key)
        env.insert("VBRT_APP_PATH", os.path.dirname(os.path.abspath(exe)))
        p = QProcess(self)
        p.setProcessEnvironment(env)
        p.setWorkingDirectory(os.path.dirname(os.path.abspath(exe)))
        p.setProcessChannelMode(QProcess.ProcessChannelMode.ForwardedChannels)
        p.finished.connect(self._ended)
        self.proc = p
        p.start(exe, [])
        if not p.waitForStarted(timeout):
            self.error = f"The add-in did not start: {exe}"
            return False
        t = QElapsedTimer()
        t.start()
        while not self.server.hasPendingConnections():
            if p.state() == QProcess.ProcessState.NotRunning or t.elapsed() > timeout:
                self.error = "The add-in did not connect to LinBASIC Studio"
                p.kill()
                return False
            self.server.waitForNewConnection(50)
        self.sock = self.server.nextPendingConnection()
        self.server.close()
        line = self._read_line(timeout)
        if line is None or line != ["K", key]:
            self.error = "The add-in did not give the key"
            self.stop()
            return False
        self.sock.readyRead.connect(self._ready)
        return True

    def stop(self):
        for c in list(self.controls):           # (an add-in that did not delete its commands: they go with it)
            try:
                c.bar.w.removeAction(c.a)
            except RuntimeError:
                pass
        self.controls = []
        if self.sock is not None and self.running():
            self._send(["X", "quit"])
            self.sock.flush()
        p = self.proc
        if p is not None and p.state() != QProcess.ProcessState.NotRunning:
            if not p.waitForFinished(3000):
                p.kill()
                p.waitForFinished(1000)
        if self.sock is not None:
            self.sock.close()
        self.sock = None

    def _ended(self, *_):
        hosts = getattr(self.main, "addin_hosts", {})
        for k, h in list(hosts.items()):
            if h is self:
                hosts.pop(k)
        if self.sock is not None:
            self.sock.close()
            self.sock = None
        for c in list(self.controls):
            try:
                c.bar.w.removeAction(c.a)
            except RuntimeError:
                pass
        self.controls = []

    # ------------------------------------------------------------ the lines
    def _send(self, fields):
        if self.sock is None:
            return
        self.sock.write(("\t".join(_esc(str(f)) for f in fields) + "\n").encode("utf-8"))
        self.sock.flush()

    def _take_line(self):
        if b"\n" not in self.buf:
            return None
        line, self.buf = self.buf.split(b"\n", 1)
        return [_unesc(f) for f in line.decode("utf-8", "replace").rstrip("\r").split("\t")]

    def _read_line(self, timeout=60000):
        t = QElapsedTimer()
        t.start()
        while True:
            ln = self._take_line()
            if ln is not None:
                return ln
            if self.sock is None or (self.proc is not None and self.proc.state() == QProcess.ProcessState.NotRunning
                                     and self.sock.bytesAvailable() == 0):
                return None
            if self.sock.bytesAvailable() == 0:
                self.sock.waitForReadyRead(20)
                QCoreApplication.processEvents(QEventLoop.ProcessEventsFlag.ExcludeUserInputEvents)
            self.buf += bytes(self.sock.readAll())
            if timeout and t.elapsed() > timeout and self.sock.bytesAvailable() == 0 and b"\n" not in self.buf:
                return None

    def _ready(self):
        """what the add-in asks by itself (its forms: a click on its button uses the IDE)"""
        if self.waiting or self.sock is None:
            return
        self.buf += bytes(self.sock.readAll())
        while True:
            ln = self._take_line()
            if ln is None:
                return
            self._request(ln)

    # ------------------------------------------------------------ values
    def ref(self, o):
        k = id(o)
        i = self.ids.get(k)
        if i is None or self.objs.get(i) is not o:
            i = self.next_id
            self.next_id += 1
            self.ids[k] = i
            self.objs[i] = o
        return i

    def enc(self, v):
        if v is EMPTY or v is None:
            return "e"
        if v is NOTHING:
            return "o:0"
        if isinstance(v, VBObject):
            return f"o:{self.ref(v)}:{v.typename()}"
        if isinstance(v, bool):
            return "b:1" if v else "b:0"
        if isinstance(v, int):
            return f"i:{v}" if -2 ** 31 <= v < 2 ** 31 else f"d:{v}"
        if isinstance(v, float):
            return f"d:{v!r}"
        return "s:" + str(v)

    def dec(self, a):
        k, _, v = a.partition(":")
        if k == "e" or k == "m":
            return None if k == "m" else EMPTY
        if k == "n":
            return None
        if k == "s":
            return v
        if k in ("h", "i", "y"):
            return int(v)
        if k in ("f", "d", "c", "t"):
            return float(v)
        if k == "b":
            return v == "1"
        if k == "o":
            i = int(v.split(":")[0] or 0)
            return self.objs.get(i, NOTHING) if i else NOTHING
        return EMPTY

    # ------------------------------------------------------------ the add-in uses the IDE
    def _request(self, f):
        if f[0] == "R" and len(f) > 1:
            i = int(f[1] or 0)
            o = self.objs.pop(i, None)
            if o is not None and self.ids.get(id(o)) == i:
                self.ids.pop(id(o), None)
            return
        if f[0] != "C" or len(f) < 4:
            return
        o = self.objs.get(int(f[1] or 0))
        kind, name = f[2], f[3].lower()
        args = [self.dec(a) for a in f[4:]]
        args = [a for a in args] if args else []
        while args and args[-1] is None:        # (arguments left out at the end)
            args.pop()
        args = [None if a is EMPTY else a for a in args]
        try:
            if o is None:
                raise AddInError(91, "Object variable or With block variable not set")
            r = self.member(o, kind, name or "item", args)
            sync_all(self.main)                 # (the events of the change come before the answer, as in VB6)
            if isinstance(r, Out):              # ByRef arguments back: the fields after the value ("-" unchanged)
                n = max(r.back) + 1 if r.back else 0
                self._send(["OK", self.enc(r.value)] + [self.enc(r.back[i]) if i in r.back else "-" for i in range(n)])
            else:
                self._send(["OK", self.enc(r)])
        except AddInError as e:
            self._send(["ER", e.number, e.description])
        except (TypeError, ValueError) as e:
            self._send(["ER", 5, f"Invalid procedure call or argument ({e})"])
        except Exception as e:                  # noqa: BLE001 - the IDE goes on; the add-in gets an error
            self._send(["ER", 440, f"Automation error ({e})"])

    def member(self, o, kind, name, args):
        if kind in ("L", "S"):
            fn = getattr(o, "let_" + name, None)
        elif kind == "G":
            fn = getattr(o, "get_" + name, None) or (getattr(o, "call_" + name, None) if args else None)
        else:
            fn = getattr(o, "call_" + name, None) or getattr(o, "get_" + name, None)
        if fn is None:
            raise AddInError(438, "Object doesn't support this property or method")
        return fn(*args)

    # ------------------------------------------------------------ the IDE calls the add-in
    def call(self, op, *args, raw=None):
        """X <op> <args>: the add-in does it (its event procedures may use the IDE meanwhile); its arguments after it.
        raw: fields already written (the object and the name of an event)"""
        if self.sock is None:
            raise AddInError(440, "The add-in is not running")
        self._send(["X", op] + list(raw or []) + [self.enc(a) for a in args])
        self.waiting += 1
        try:
            while True:
                ln = self._read_line(0)
                if ln is None:
                    raise AddInError(440, "The add-in ended")
                if ln[0] == "D":
                    return [self.dec(a) for a in ln[1:]]
                if ln[0] == "ER":
                    raise AddInError(int(ln[1]) if len(ln) > 1 and ln[1].lstrip("-").isdigit() else 5,
                                     ln[2] if len(ln) > 2 else "")
                self._request(ln)
        finally:
            self.waiting -= 1

    def connect(self, mode):
        inst = vbide.AddIn(self.vbe, self.main, self.info["progid"])
        self.call("connect", self.vbe, mode, inst)

    def fire_event(self, src, name, *args):
        return self.call("event", *args, raw=[str(self.ref(src)), name])

    def control_clicked(self, ctl):
        """a command bar control of the add-in: the Click of its CommandBarEvents (handled, CancelDefault)"""
        for e in list(self.vbe.get_events().bar_events):
            if e.ctl is ctl:
                try:
                    self.fire_event(e, "Click", ctl, False, False)
                except AddInError as er:
                    self.main.addin_error(self.info, f"{er.description} ({er.number})")

    def control_added(self, ctl):
        self.controls.append(ctl)

    def control_deleted(self, ctl):
        if ctl in self.controls:
            self.controls.remove(ctl)
