#!/usr/bin/env python3
"""The test robot over a folder of VB6 projects (the MSDN samples of your own VB6, for example).

    python3 tests/samples/run_samples.py <folder> [<out folder>] [name ...]

Every Standard EXE project under <folder> is copied (with the whole archive it is in, the first folder
under <folder>, since a project may use ..\\ files and projects) to <out folder>/<name>/tree, compiled with
lbvc --debug, and run on a virtual screen (Xvfb) with VBRT_MONKEY=1: the runtime clicks every
button, option, check box, list and menu item once, answers MsgBox and InputBox, cancels the
CommonDialog, and writes each action and each run-time error (with module, procedure and line) to
run.txt instead of stopping; pictures of the forms go to <form>-start.png and <form>-end.png.
One line per project: OK, ERR (run-time errors, and compile errors of procedures that were called: the
robot compiles with --debug, so Compile On Demand as in the VB6 IDE), COMPILE, HANG (still running after 45 s),
FLOOD (an endless loop of errors), EXITn (ended before the robot was done).
Needs: Xvfb, ImageMagick (import) for the screen shot, bubblewrap (bwrap).

The programs are strangers' code that the robot clicks through (and answers Yes to every MsgBox): each runs
inside bubblewrap, where the whole disk is read-only and only its own <out folder>/<name> (and $HOME, when
ROBOT_HOME=1) can be written, without network, without /run (your session's D-Bus) and with no X display but the
robot's (robot 2026-11-10: Dir2HTML answered "Yes" to "delete / ?" and called fso.DeleteFolder "/"). Without bwrap the robot does not run (ROBOT_UNSAFE=1 runs it anyway).
No project files are part of LinBASIC Studio: the folder is yours."""
import glob, os, re, shutil, subprocess, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
LBVC = os.path.join(HERE, '..', '..', 'compiler', 'lbvc.py')
DISPLAY = ':77'


def projects(src):
    for vbp in sorted(glob.glob(src + '/**/*.[vV][bB][pP]', recursive=True)):
        txt = open(vbp, 'rb').read().decode('cp1252', 'replace')
        m = re.search(r'^Type=(\w+)', txt, re.M | re.I)
        st = re.search(r'^Startup="?([^"\r\n]*)', txt, re.M | re.I)
        if m and (m.group(1).lower() == 'exe' or        # an ActiveX EXE with a startup object runs on its own too
                  (m.group(1).lower() == 'oleexe' and st and st.group(1).strip().lower() not in ('(none)', ''))):
            yield vbp


def sandbox(out):
    """the command that runs a program of the robot in bubblewrap: the disk read-only, its folder writable"""
    if os.environ.get('ROBOT_UNSAFE') == '1':
        return []
    bw = shutil.which('bwrap')
    if not bw:
        sys.exit("run_samples: bubblewrap (bwrap) is needed to run the programs safely: sudo apt install bubblewrap "
                 "(ROBOT_UNSAFE=1 runs them without it, on your real files)")
    # /run hidden: the D-Bus of your session (/run/user/<uid>/bus: through it a program could ask your desktop to
    # open, move or delete things - robot 2026-11-10: ShellExecute "mailto:" started your Thunderbird), Wayland, sound
    cmd = [bw, '--ro-bind', '/', '/', '--dev', '/dev', '--proc', '/proc', '--tmpfs', '/tmp', '--tmpfs', '/run']
    x = '/tmp/.X11-unix/X' + DISPLAY.lstrip(':')
    if os.path.exists(x):                                               # only the Xvfb of the robot, not your screen
        cmd += ['--dir', '/tmp/.X11-unix', '--ro-bind', x, x]
    cmd += ['--bind', out, out]
    home = os.environ.get('HOME', '')
    if os.environ.get('ROBOT_HOME') == '1' and home and os.path.isdir(home):
        cmd += ['--bind', home, home]                                   # a HOME of the robot's own (ce.sh)
    xdg = os.environ.get('XDG_RUNTIME_DIR', '')
    if xdg and os.path.isdir(xdg) and os.environ.get('ROBOT_HOME') == '1':
        cmd += ['--bind', xdg, xdg]
    return cmd + ['--unshare-net', '--unshare-pid', '--die-with-parent', '--']


def where(exe):
    """SIGUSR1 to the program (in the sandbox it has another pid): a --debug program of the robot writes where it is"""
    import signal
    real = os.path.realpath(exe)
    for d in os.listdir('/proc'):
        if d.isdigit():
            try:
                if os.readlink('/proc/%s/exe' % d) == real:
                    os.kill(int(d), signal.SIGUSR1)
            except OSError:
                pass
    time.sleep(0.5)


def run(src, outdir, vbp):
    name = os.path.relpath(vbp, src).replace('/', '_').rsplit('.', 1)[0]
    out = os.path.join(outdir, name)
    shutil.rmtree(out, ignore_errors=True)
    os.makedirs(out)
    # a copy of the whole archive (the programs write files next to them; a project may use files and other
    # projects of its archive: Module=..\Common\x.bas, Reference=*\A..\Lib\Lib.vbp)
    rel = os.path.relpath(vbp, src).replace(os.sep, '/')
    root = rel.split('/')[0] if '/' in rel else ''
    tree = os.path.join(out, 'tree')
    if root:
        shutil.copytree(os.path.join(src, root), tree, ignore=shutil.ignore_patterns('*-win64', '*-win32', '*.exe'))
        work = os.path.normpath(os.path.join(tree, os.path.dirname(rel)[len(root):].lstrip('/')))   # (no "/" at the end)
    else:
        work = os.path.join(out, 'proj')
        shutil.copytree(os.path.dirname(vbp), work, ignore=shutil.ignore_patterns('*-win64', '*-win32', '*.exe'))
    exe = os.path.join(out, 'prog')
    b = subprocess.run([sys.executable, LBVC, '--debug', os.path.join(work, os.path.basename(vbp)), '-o', exe],
                       capture_output=True, text=True, timeout=600)
    open(os.path.join(out, 'build.txt'), 'w').write(b.stdout + b.stderr)
    if b.returncode != 0 or not os.path.exists(exe):
        return name, 'COMPILE', (b.stdout + b.stderr).strip().splitlines()[-3:]
    env = dict(os.environ, DISPLAY=DISPLAY, VBRT_MONKEY='1', VBRT_APP_PATH=work, QT_QPA_PLATFORM='xcb',
               VBRT_MONKEY_SHOTS=out)
    log = open(os.path.join(out, 'run.txt'), 'w')
    env.pop('SESSION_MANAGER', None)            # (Qt: "Session management error" from the desktop's session)
    for k in ('WAYLAND_DISPLAY', 'DBUS_SESSION_BUS_ADDRESS', 'DBUS_STARTER_ADDRESS'):
        env.pop(k, None)                        # (the sandbox hides them too: /run)
    # stdin: nothing (a program without forms reads InputBox from the console: with the robot's terminal it waited
    # for the user - robot 2026-11-12: RegisterScriptFiles)
    p = subprocess.Popen(sandbox(out) + [exe], cwd=work, stdin=subprocess.DEVNULL, stdout=log,
                         stderr=subprocess.STDOUT, env=env, start_new_session=True)   # its own process group: stop = all of it

    def stop():
        try:
            os.killpg(p.pid, 9)
        except OSError:
            p.kill()
        p.wait()
    rc, t0 = None, time.time()
    while rc is None and time.time() - t0 < 45:
        rc = p.poll()
        if rc is None:
            if os.path.getsize(os.path.join(out, 'run.txt')) > 2000000:
                stop(); rc = 'FLOOD'
            time.sleep(0.2)
    if rc is None:
        where(exe)                                  # the program writes where it is: "monkey: HANG at ..."
        subprocess.run(['import', '-display', DISPLAY, '-window', 'root', os.path.join(out, 'hang.png')],
                       capture_output=True, timeout=20)
        stop(); rc = 'TIMEOUT'
    log.close()
    text = open(os.path.join(out, 'run.txt'), errors='replace').read()
    errs = []
    for line in text.splitlines():
        if ('monkey: ERROR' in line or 'monkey: COMPILE ERROR' in line) and line not in errs:
            errs.append(line)
    done = 'monkey: done' in text
    if rc == 'TIMEOUT':
        status = 'HANG'
        errs = [line for line in text.splitlines() if line.startswith('monkey: HANG at')][-1:] + errs
    elif rc == 'FLOOD':
        status = 'FLOOD'
    elif errs:
        status = 'ERR'
    elif done and rc == 0:
        status = 'OK'
    else:
        status = 'EXIT%s' % rc
        errs = text.strip().splitlines()[-4:]
    warn = [line for line in b.stderr.splitlines() if ': warning:' in line]
    return name, status, errs[:6] + warn[:4]


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    src = os.path.abspath(sys.argv[1])
    outdir = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(HERE, 'work')
    sandbox(outdir)                             # bwrap is there (or the robot stops here)
    want = [w.lower() for w in sys.argv[3:]]
    xv = subprocess.Popen(['Xvfb', DISPLAY, '-screen', '0', '1280x900x24'], stderr=subprocess.DEVNULL)
    time.sleep(1)
    try:
        for vbp in projects(src):
            if want and not any(w in os.path.relpath(vbp, src).lower() for w in want):
                continue
            name, st, lines = run(src, outdir, vbp)
            print('%-30s %s' % (name, st))
            for line in lines:
                print('      ', line[:200])
            sys.stdout.flush()
    finally:
        xv.kill()
    return 0


sys.exit(main())
