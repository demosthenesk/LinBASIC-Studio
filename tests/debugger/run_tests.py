#!/usr/bin/env python3
"""Tests of the debugger: a program compiled with lbvc --debug talks to this script as it talks
to the IDE (breakpoints, Step Into/Over/Out, Locals, Call Stack, ? expressions, assignments in
the Immediate window, Stop, the break at an unhandled error). d2 (phase 8.2): Set Next Statement, watches that
break (Break When Value Changes / Is True), Error Trapping (Break in Class Module, on All Errors, on Unhandled
Errors: MSDN "Debugging Class Modules"), F5 at an error runs the statement again, Alt+F5 (errgo) goes on with it.

    python3 tests/debugger/run_tests.py
    python.exe tests\\debugger\\run_tests.py     (Windows: a named pipe, as QLocalServer has there)
"""
import os
import socket
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))

# (what the script sends, what the program must answer: a line, or the lines up to "end")
SCRIPT = [
    ("bp+ Module1 21", None),
    ("go", ["break\tModule1\t21\tbp"]),
    ("locals", ["var\ta\tLong\t0", "var\tb\tLong\t1", "var\tAdd\tLong\t0", "var\tr\tLong\t0",
                "mvar\tTotal\tLong\t0", "mvar\tNames\tString\tString(0 to 2)", "end"]),
    ("stack", ["frame\tModule1\tAdd\t21", "frame\tModule1\tMain\t11", "end"]),
    ("eval 0\ta + b * 10", ["ok\t10"]),
    ("eval 1\ts & \"!\"", ["ok\tabc!"]),
    ("eval 0\tLen(\"ab\" & Left$(\"xyz\", 2)) + Total", ["ok\t4"]),
    ("eval 0\tnosuchvar", ["err\tVariable not found: nosuchvar"]),
    ("eval 0\t1 / 0", ["err\tRun-time error '11': Division by zero"]),
    ("exec 1\ts = UCase(s)", ["ok\t"]),
    ("bp- Module1 21", None),
    ("out", ["break\tModule1\t11\tstep"]),
    ("eval 0\ts", ["ok\tABC"]),
    ("over", ["break\tModule1\t11\tstep"]),
    ("go", ["break\tModule1\t15\tstop"]),
    ("eval 0\tv", ["ok\t12"]),
    ("go", ["break\tModule1\t16\terror\t11\tDivision by zero"]),
    ("eval 0\ti", ["ok\t4"]),
    ("errgo", None),
]
OUT1 = "ABC 12 \nafter stopRun-time error '11':\n\nDivision by zero\n"

NEXT_ERR = "err\tCan't set the next statement to this line: it must be a statement of the procedure that is running"
DIV0 = "error\t11\tDivision by zero"
# d2/Debug2.vbp, three runs: the default Error Trapping (Break in Class Module), Break on All Errors, Break on
# Unhandled Errors
SCRIPT_A = [
    ("bp+ Module2 8", None),
    ("watch 1\t2\tModule2\tMain\tCount", None),
    ("go", ["break\tModule2\t8\tbp"]),
    ("next Module2 9", ["ok\t", "break\tModule2\t9\tstep"]),          # x = x + 10 is not run
    ("eval 0\tx", ["ok\t1"]),
    ("next Module2 3", [NEXT_ERR]),
    ("bp- Module2 8", None),
    ("go", ["break\tModule2\t12\twatch\t1"]),                       # Count went from 0 to 1
    ("eval 0\ti", ["ok\t2"]),
    ("wdel 1", None),
    ("watch 2\t1\t\t\tCount >= 4", None),
    ("go", ["break\tModule2\t12\twatch\t2"]),
    ("eval 0\ti", ["ok\t5"]),
    ("go", ["break\tClass1\t14\t" + DIV0]),                         # Count >= 4 stays true: no break again;
                                                                        # On Error of Main: not for the class
    ("go", ["break\tClass1\t14\t" + DIV0]),                         # F5: the statement again
    ("errgo", None),                                                    # Alt+F5: Main does not handle it
]
OUT_A = "x 101 \ncount 5 \nRun-time error '11':\n\nDivision by zero\n"
SCRIPT_B = [
    ("trap 0", None),
    ("go", ["break\tModule2\t15\t" + DIV0]),                        # On Error Resume Next: it stops too
    ("errgo", ["break\tClass1\t14\t" + DIV0]),
    ("errgo", None),
]
OUT_B = "x 111 \ncount 5 \nRun-time error '11':\n\nDivision by zero\n"
SCRIPT_C = [
    ("trap 2", None),
    ("go", ["break\tModule2\t19\t" + DIV0]),                        # at the call of the class
    ("go", ["break\tModule2\t19\t" + DIV0]),                        # the call again
    ("next Module2 20", ["ok\t", "break\tModule2\t20\tstep"]),      # past it: the error is dropped
    ("go", None),
]
OUT_C = "x 111 \ncount 5 \nend\n"
RUNS = [("Module1.bas", SCRIPT, OUT1), (os.path.join("d2", "Debug2.vbp"), SCRIPT_A, OUT_A),
        (os.path.join("d2", "Debug2.vbp"), SCRIPT_B, OUT_B), (os.path.join("d2", "Debug2.vbp"), SCRIPT_C, OUT_C)]


def pipe_server(name):
    """Windows: a byte-mode named pipe; (its name, a function that waits for the program and gives
    the connection as a text file)"""
    import ctypes
    import msvcrt
    from ctypes import wintypes
    k = ctypes.WinDLL("kernel32", use_last_error=True)
    k.CreateNamedPipeW.restype = wintypes.HANDLE
    k.CreateNamedPipeW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.DWORD,
                                   wintypes.DWORD, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p]
    k.ConnectNamedPipe.argtypes = [wintypes.HANDLE, ctypes.c_void_p]
    path = "\\\\.\\pipe\\" + name
    h = k.CreateNamedPipeW(path, 3, 0, 1, 65536, 65536, 0, None)     # PIPE_ACCESS_DUPLEX, PIPE_TYPE_BYTE

    def accept():
        k.ConnectNamedPipe(h, None)
        fd = msvcrt.open_osfhandle(h, 0)
        return os.fdopen(fd, "r+b", buffering=0)
    return path, accept


class LineFile:
    """readline / write / flush over a raw pipe (text, UTF-8, lines end with \\n)"""

    def __init__(self, raw):
        self.raw, self.buf = raw, b""

    def readline(self):
        while b"\n" not in self.buf:
            d = self.raw.read(4096)
            if not d:
                break
            self.buf += d
        line, _, self.buf = self.buf.partition(b"\n")
        return line.decode("utf-8") + "\n"

    def write(self, s):
        self.raw.write(s.encode("utf-8"))

    def flush(self):
        pass


def run(work, n, source, script, want_out):
    """one program: its script; the number of checks that failed"""
    exe = os.path.join(work, f"prog{n}" + (".exe" if os.name == "nt" else ""))
    r = subprocess.run([sys.executable, os.path.join(ROOT, "compiler", "lbvc.py"), os.path.join(HERE, source),
                        "-o", exe, "--debug"], capture_output=True, text=True)
    if r.returncode:
        print(f"FAIL  {source}: compile\n{r.stderr}")
        return len(script) + 1
    if os.name == "nt":
        path, accept = pipe_server(os.path.basename(work) + str(n))
    else:
        path = os.path.join(work, f"sock{n}")
        srv = socket.socket(socket.AF_UNIX)
        srv.bind(path)
        srv.listen(1)
        srv.settimeout(20)
    p = subprocess.Popen([exe], env=dict(os.environ, VBRT_DEBUG=path), stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT, text=True)
    if os.name == "nt":
        f = LineFile(accept())
    else:
        c, _ = srv.accept()
        c.settimeout(20)
        f = c.makefile("rw", newline="\n", encoding="utf-8")
    failed = 0
    if f.readline().rstrip("\n") != "hello":
        print(f"FAIL  {source}: no hello")
        p.kill()
        return len(script) + 1
    for cmd, want in script:
        f.write(cmd + "\n")
        f.flush()
        if want is None:
            continue
        got = [f.readline().rstrip("\n") for _ in want]
        ok = got == want
        failed += not ok
        print(f"{'ok  ' if ok else 'FAIL'}  {cmd!r}")
        if not ok:
            for g, w in zip(got, want):
                if g != w:
                    print(f"      want {w!r}\n      got  {g!r}")
    try:
        out = p.communicate(timeout=20)[0].replace("\r\n", "\n")
    except subprocess.TimeoutExpired:
        p.kill()
        out = "(the program did not end)"
    if out != want_out:
        failed += 1
        print(f"FAIL  {source}: output {out!r}")
    return failed


def main():
    work = tempfile.mkdtemp(prefix="lbdbg-")
    total = failed = 0
    for n, (source, script, want_out) in enumerate(RUNS):
        print(f"---- {source} ({n})")
        failed += run(work, n, source, script, want_out)
        total += sum(1 for _, w in script if w is not None) + 1
    print(f"{total - failed}/{total} passed")
    return 1 if failed else 0


sys.exit(main())
