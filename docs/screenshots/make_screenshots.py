#!/usr/bin/env python3
"""make_screenshots.py - the pictures of README.md: the IDE with the example projects, and the examples
themselves, on a virtual screen (Xvfb).

    python3 docs/screenshots/make_screenshots.py            (needs Xvfb, PyQt6, PyQt6-QScintilla)

The IDE pictures are taken by the IDE's own Qt (the screen, so that the IntelliSense list is in them);
the programs are built with lbvc and take their own picture (VBRT_SCREENSHOT=file.png)."""
import os
import shutil
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
EX = os.path.join(ROOT, "examples")
DISPLAY = ":93"
work = tempfile.mkdtemp(prefix="lbshots-")
os.makedirs(os.path.join(work, "xdg"), mode=0o700)
xvfb = subprocess.Popen(["Xvfb", DISPLAY, "-screen", "0", "1600x1100x24"], stderr=subprocess.DEVNULL)
time.sleep(1)
os.environ.update(DISPLAY=DISPLAY, QT_QPA_PLATFORM="xcb", XDG_RUNTIME_DIR=os.path.join(work, "xdg"),
                  LBVC_REGISTRY=os.path.join(work, "components.json"))


def examples():
    """each example: copied, built and run until it takes its picture"""
    for name, form in (("Hello", None), ("Notepad", None), ("Shapes", None), ("Explorer", None),
                       ("Chart", None), ("AddressBook", None)):
        d = os.path.join(work, name)
        shutil.copytree(os.path.join(EX, name), d)
        exe = os.path.join(d, "prog")
        subprocess.run([sys.executable, os.path.join(ROOT, "compiler", "lbvc.py"), os.path.join(d, name + ".vbp"),
                        "-o", exe], check=True)
        out = os.path.join(HERE, "example-" + name.lower() + ".png")
        subprocess.run([exe], cwd=d, timeout=60, env=dict(os.environ, VBRT_SCREENSHOT=out),
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(out)


def ide():
    sys.path.insert(0, os.path.join(ROOT, "ide"))
    from lbstudio import paths, settings
    paths.set_config_dir(os.path.join(work, "config"))
    paths.ensure_user_files()
    settings.init()
    from PyQt6.QtWidgets import QApplication, QMessageBox
    app = QApplication(sys.argv[:1])
    QMessageBox.information = lambda *a, **k: None
    from lbstudio.mainwindow import MainWindow

    def settle(seconds=0.6):
        t0 = time.time()
        while time.time() - t0 < seconds:
            app.processEvents()
            time.sleep(0.02)

    def shot(win, name, extra=None):
        settle()
        g = win.frameGeometry()
        if extra is not None:                    # a popup beside the window (the IntelliSense list)
            g = g.united(extra.frameGeometry())
        pic = app.primaryScreen().grabWindow(0, g.x(), g.y(), g.width(), g.height())
        out = os.path.join(HERE, name)
        pic.save(out)
        print(out)

    def project(name):
        d = os.path.join(work, "ide-" + name)
        shutil.copytree(os.path.join(EX, name), d)
        return os.path.join(d, name + ".vbp")

    w = MainWindow()
    w.move(0, 0)
    w.resize(1360, 860)
    w.show()
    from PyQt6.QtCore import Qt
    w.resizeDocks([w.dock_bottom], [170], Qt.Orientation.Vertical)

    # the form designer, the Toolbox, the Properties window and the Project Explorer
    w.open_project(project("Hello"))
    f = w.open_designer(w.project.module("frmHello"))
    c = next((x for x in f.root.walk() if getattr(x, "name", "") == "cmdSay"), None)
    if c is not None:
        f.select_control(c)
    shot(w, "ide-designer.png")

    # the controls of the .ocx libraries in the designer
    w.open_project(project("Chart"))
    w.open_designer(w.project.module("frmChart"))
    shot(w, "ide-designer-ocx.png")

    # the code window: the Object and Procedure boxes, IntelliSense after a dot
    w.open_project(project("Notepad"))
    cw = w.open_code(w.project.module("frmNotepad"))
    ed = cw.editor
    text = ed.all_text()
    at = text.find("    txtText.FontName = dlgFile.FontName")
    if at >= 0:
        line = text.count("\n", 0, at)
        ed.replace_range(ed.line_start(line), ed.line_start(line), "    txtText.\n")
        ed.goto_pos(ed.line_start(line) + len("    txtText."))
        cw.assist.list_members()
    popup = next((x for x in app.topLevelWidgets() if x.isVisible() and x is not w and x.width() < 400), None)
    shot(w, "ide-code.png", popup)
    ed.SendScintilla(2101)                      # SCI_AUTOCCANCEL

    # the debugger: a breakpoint, the yellow line, Locals
    w.open_project(project("Shapes"))
    cw = w.open_code(w.project.module("frmShapes"))
    ed = cw.editor
    text = ed.all_text()
    line = text[:text.find(".R = 12 + Rnd * 18")].count("\n")
    ed.goto_pos(ed.line_start(line))
    w.toggle_breakpoint()
    w.run_start()
    t0 = time.time()
    while w.runner.state != "break" and time.time() - t0 < 120:
        settle(0.2)
    if w.runner.state == "break":
        w.debug_step("step")                    # F8 once: the yellow line moves on
        settle(1.5)
        w.bottom_tabs.setCurrentWidget(w.locals_win)
        shot(w, "ide-debugger.png")
        w.run_end()
        settle(1)

    # File > Make: Linux, Windows 32-bit, Windows 64-bit
    from lbstudio.dialogs import MakeDialog
    dlg = MakeDialog(w, "/home/user/Projects/Shapes", "Shapes", "linux")
    dlg.move(420, 260)
    dlg.show()
    shot(dlg, "ide-make.png")
    dlg.close()
    for q in [w.project]:
        q.dirty = False
        for m in q.modules:
            m.dirty = False
    w.close()


try:
    examples()
    ide()
finally:
    xvfb.kill()
    shutil.rmtree(work, ignore_errors=True)
