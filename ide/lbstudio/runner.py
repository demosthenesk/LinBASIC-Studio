"""Run (F5) and Make EXE: lbvc compiles the project, the program runs with its output in the
Immediate window. The project is compiled as it is in the editors: an unsaved project or
unsaved changes are written to a temporary copy first, like VB6 runs what is in memory."""
import os
import re
import shutil
import sys
import tempfile

from PyQt6.QtCore import QObject, QProcess, QProcessEnvironment, pyqtSignal
from PyQt6.QtNetwork import QLocalServer

from . import paths

ERROR_RE = re.compile(r"^(.*?):(\d+):(\d+): error: (.*)$")
WARNING_RE = re.compile(r"^(.*?):(\d+):(\d+): warning: (.*)$")
RUNTIME_RE = re.compile(r"Run-time error '(\d+)':\s*\n\s*\n?(.*)", re.S)


def lbvc_command():
    exe = sys.executable
    if os.name == "nt" and os.path.basename(exe).lower() == "pythonw.exe":   # the console Python: its output
        exe = os.path.join(os.path.dirname(exe), "python.exe")             # comes to the IDE
    return [exe, os.path.join(paths.COMPILER, "lbvc.py")]


class Runner(QObject):
    output = pyqtSignal(str)                    # text for the Immediate window
    compile_error = pyqtSignal(str, int, int, str)   # file, line (1-based, of the file), column, message
    compile_warning = pyqtSignal(str, int, int, str)
    runtime_error = pyqtSignal(int, str)
    state_changed = pyqtSignal(str)             # "design" | "compiling" | "running" | "break"
    built = pyqtSignal(str)                     # Make EXE: the executable
    broke = pyqtSignal(str, int, str, str)      # module, line of the file, reason, detail (error n<TAB>text)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.proc = None
        self.state = "design"
        self.tmp = None
        self.copy_dir = None
        self.project_dir = None
        self.exe = None
        self.make_target = None
        self.stderr = ""
        self.debug = False
        self.first = "go"
        self.server = None
        self.sock = None
        self.buf = b""
        self.pending = []                       # [(kind, callback, lines)] answers of the program, in order
        self.bp_provider = lambda: []           # [(module, line of the file)] of the IDE
        self.trap_provider = lambda: 1          # Error Trapping: 0 all errors, 1 in class module, 2 unhandled
        self.watch_provider = lambda: []        # the watch commands of the watches that break

    def _set(self, st):
        self.state = st
        self.state_changed.emit(st)

    def _env(self, codepage):
        env = QProcessEnvironment.systemEnvironment()
        env.insert("LBVC_RUNTIME", paths.RUNTIME)
        return env

    def start(self, project, codepage, make_target=None, debug=False, first="go", target=None):
        """compile, then run (make_target None) or only build make_target; debug: with the
        debugger (F5 and F8 of the IDE), first: "go" or "step" (F8 stops at the first line)"""
        if self.state != "design":
            return
        self.debug = debug and make_target is None
        self.codepage = codepage
        self.first = first
        self.cleanup()
        self.tmp = tempfile.mkdtemp(prefix="lbstudio-")
        self.copy_dir = os.path.join(self.tmp, "src")
        group = getattr(project, "group", None)
        vbp = group.write_copy(self.copy_dir, project) if group is not None else project.write_copy(self.copy_dir)
        self.project_dir = os.path.dirname(project.path) if project.path else os.path.expanduser("~")
        target = target or paths.host_target()
        self.exe = make_target or os.path.join(self.tmp, project.name + (".exe" if os.name == "nt" else ""))
        self.make_target = make_target
        # Make of an ActiveX DLL or Control: no program; the registry names the project (its real .vbp)
        self.register = project.path if make_target and getattr(project, "is_component", False) else None
        self.no_exe = make_target is not None and getattr(project, "ptype", "Exe") in ("OleDll", "Control") and \
            not getattr(project, "is_addin", False)       # (an add-in is made as a program, phase 12.5)
        self.stderr = ""
        p = QProcess(self)
        p.setProcessEnvironment(self._env(codepage))
        p.setProcessChannelMode(QProcess.ProcessChannelMode.SeparateChannels)
        p.readyReadStandardError.connect(lambda: self._collect_err(p))
        p.finished.connect(self._compiled)
        self.proc = p
        self._set("compiling")
        p.start(lbvc_command()[0], lbvc_command()[1:] + [vbp, "-o", self.exe, "--codepage", codepage] +
                (["--debug"] if self.debug else []) +
                # Run (F5) is a run of the IDE: the order of P-code and Debug.Print; Make: the order the Compile
                # settings of the project ask for, and no Debug.Print (as VB6)
                (["--ide"] if make_target is None else []) +
                (["--target", target] if target != paths.host_target() else []))

    def _collect_err(self, p):
        self.stderr += bytes(p.readAllStandardError()).decode("utf-8", "replace")

    def _compiled(self, code, status):
        p = self.proc
        self._collect_err(p)
        out = bytes(p.readAllStandardOutput()).decode("utf-8", "replace")
        self.proc = None
        if code != 0 or (not os.path.exists(self.exe) and not self.no_exe):
            self._set("design")
            msg = (self.stderr + out).strip()
            for line in msg.splitlines():
                m = ERROR_RE.match(line.strip())
                if m:
                    f = m.group(1)
                    if self.copy_dir and f.startswith(self.copy_dir):
                        f = f[len(self.copy_dir) + 1:]      # relative to the project
                    self.compile_error.emit(f, int(m.group(2)), int(m.group(3)), m.group(4))
                    return
            self.compile_error.emit("", 0, 0, msg or "the compiler failed")
            return
        for line in (self.stderr + out).splitlines():  # it compiles, but something will not work
            m = WARNING_RE.match(line.strip())
            if m:
                f = m.group(1)
                if self.copy_dir and f.startswith(self.copy_dir):
                    f = f[len(self.copy_dir) + 1:]
                self.compile_warning.emit(f, int(m.group(2)), int(m.group(3)), m.group(4))
        if self.make_target:
            if self.register:
                from lbvc.driver import register_component
                try:
                    register_component(self.register, self.codepage)
                except Exception as e:          # noqa: BLE001 - shown as a compile error
                    self._set("design")
                    self.compile_error.emit("", 0, 0, f"can't register the component: {e}")
                    return
            self._set("design")
            self.built.emit(self.make_target)
            return
        self._run()

    def _run(self):
        p = QProcess(self)
        env = QProcessEnvironment.systemEnvironment()
        env.insert("VBRT_APP_PATH", self.project_dir)    # App.Path: the project folder, as in the VB6 IDE
        env.insert("VBRT_NO_ERRBOX", "1")                 # run-time errors go to the Output tab
        if self.debug:
            if os.name == "nt":                         # a named pipe (QLocalServer on Windows)
                path = "\\\\.\\pipe\\" + os.path.basename(self.tmp) + "-debug"
            else:
                path = os.path.join(self.tmp, "debug.sock")
            self.server = QLocalServer(self)
            QLocalServer.removeServer(path)
            if self.server.listen(path):
                self.server.newConnection.connect(self._connected)
                env.insert("VBRT_DEBUG", path)
        p.setProcessEnvironment(env)
        p.setWorkingDirectory(self.project_dir)
        p.readyReadStandardOutput.connect(lambda: self.output.emit(bytes(p.readAllStandardOutput()).decode("utf-8", "replace")))
        self.stderr = ""
        p.readyReadStandardError.connect(lambda: self._collect_err(p))
        p.finished.connect(self._ended)
        self.proc = p
        self._set("running")
        p.start(self.exe, [])

    # ------------------------------------------------------------ the debugger
    def _connected(self):
        self.sock = self.server.nextPendingConnection()
        self.sock.readyRead.connect(self._read)

    def _read(self):
        if self.sock is None:
            return
        self.buf += bytes(self.sock.readAll())
        while b"\n" in self.buf:
            line, self.buf = self.buf.split(b"\n", 1)
            self._line(line.decode("utf-8", "replace"))

    def _line(self, line):
        parts = line.split("\t")
        if parts[0] == "hello":
            self.send(f"trap {self.trap_provider()}")
            for cmd in self.watch_provider():
                self.send(cmd)
            for mod, ln in self.bp_provider():
                self.send(f"bp+ {mod} {ln}")
            self.send(self.first)
            return
        if parts[0] == "break" and len(parts) >= 4:
            self._set("break")
            self.broke.emit(parts[1], int(parts[2]), parts[3], "\t".join(parts[4:]))
            return
        if not self.pending:
            return
        kind, cb, acc = self.pending[0]
        if kind == "multi":
            if line == "end":
                self.pending.pop(0)
                cb(acc)
            else:
                acc.append(parts)
        else:
            self.pending.pop(0)
            cb(parts[0] == "ok", "\t".join(parts[1:]))

    def send(self, cmd):
        if self.sock is not None:
            self.sock.write((cmd.replace("\n", " ") + "\n").encode("utf-8"))
            self.sock.flush()

    def request(self, cmd, cb, multi=False):
        """a question to the program in break mode; cb(rows) for locals/stack, cb(ok, text) else"""
        if self.sock is None or self.state != "break":
            return False
        self.pending.append(("multi" if multi else "one", cb, []))
        self.send(cmd)
        return True

    def resume(self, cmd):
        """go, step, over, out, runto <module> <line>, errgo, errstep"""
        if self.state == "break":
            self.pending.clear()
            self._set("running")
            self.send(cmd)

    def pause(self):
        if self.state == "running":
            self.send("pause")

    def _ended(self, code, status):
        p = self.proc
        if p is not None:
            self._collect_err(p)
            rest = bytes(p.readAllStandardOutput()).decode("utf-8", "replace")
            if rest:
                self.output.emit(rest)
        self.proc = None
        if self.sock is not None:
            self.sock.close()
            self.sock = None
        if self.server is not None:
            self.server.close()
            self.server = None
        self.pending.clear()
        self.buf = b""
        self._set("design")
        for line in self.stderr.splitlines():       # Compile On Demand: a procedure with a compile error was called
            ce = ERROR_RE.match(line.strip())
            if ce:
                f = ce.group(1)
                if self.copy_dir and f.startswith(self.copy_dir):
                    f = f[len(self.copy_dir) + 1:]
                self.compile_error.emit(f, int(ce.group(2)), int(ce.group(3)), ce.group(4))
                self.cleanup()
                return
        m = RUNTIME_RE.search(self.stderr)
        if m:
            self.runtime_error.emit(int(m.group(1)), m.group(2).strip())
        self.cleanup()

    def stop(self):
        """Run > End"""
        if self.proc is not None:
            self.proc.kill()
            self.proc.waitForFinished(2000)

    def cleanup(self):
        if self.tmp and os.path.isdir(self.tmp) and self.state == "design":
            shutil.rmtree(self.tmp, ignore_errors=True)
            self.tmp = None
