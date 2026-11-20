#!/usr/bin/env python3
"""Tests of the form designer of LinBASIC Studio: forms are built with the designer, as a user
would build them (place controls, a frame with options, a control array by copy and paste),
saved, and compared byte for byte with the same form saved by VB6 (<test>/expected/).

    python3 tests/designer/run_tests.py
"""
import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "ide"))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
xdg = os.path.join(HERE, "work", ".xdg")
os.makedirs(xdg, exist_ok=True)
os.chmod(xdg, 0o700)
os.environ["XDG_RUNTIME_DIR"] = xdg

from lbstudio import paths, settings  # noqa: E402

work = os.path.join(HERE, "work")
paths.set_config_dir(os.path.join(work, "config"))
paths.ensure_user_files()
settings.init()
settings.S["codepage"] = "cp1253"          # the VB6 of the tests runs on a Greek Windows

from PyQt6.QtCore import QRect  # noqa: E402
from PyQt6.QtWidgets import QApplication, QMessageBox, QFileDialog  # noqa: E402

app = QApplication(sys.argv[:1])

from lbstudio.mainwindow import MainWindow  # noqa: E402
from lbstudio.designer import TOOLS  # noqa: E402
from lbstudio import vbprops as V  # noqa: E402
from lbstudio.menueditor import MenuEditor, _entry  # noqa: E402


def setp(w, d, ctl, name, value):
    """a change in the Properties window, as the user makes it"""
    w.props.show_object(d.module, d, ctl)
    p = V.prop_def(ctl.type_name, name)
    assert p is not None, (ctl.type_name, name)
    w.props.commit(p, value)


def menus(d, rows):
    """the Menu Editor: rows of (level, caption, name, {other fields})"""
    dlg = MenuEditor(d, d.root, [])
    dlg.entries = []
    for level, cap, name, extra in rows:
        e = _entry(level)
        e.update(caption=cap, name=name, **extra)
        dlg.entries.append(e)
    dlg.entries.append(_entry(0))
    dlg._ok()
    assert dlg.result_menus is not None, "the Menu Editor refused the menus"
    d.set_menus(dlg.result_menus)


def build_d_all(w):
    """all the controls of the Toolbox, Frame1 with two options, Command2(0) and Command2(1)"""
    d = w.current_designer()
    d.set_client_size(8400, 7200)
    x = y = 8
    for cls, base, dw, dh in TOOLS:
        if cls in ("Frame", "Data"):          # Data: after the VB6 oracle of d_all (d_data has it)
            continue
        wpx, hpx = min(dw, 1215) // 15, (max(dh, 255) // 15 if cls != "Line" else 30)
        if x + wpx > 550:
            x, y = 8, y + 90
        d.place_new(cls, QRect(x, y, wpx, hpx), None)
        x += wpx + 16
    fr = d.place_new("Frame", QRect(8, y + 90, 200, 100), None)
    for i in range(2):
        d.place_new("OptionButton", QRect(16, y + 110 + 30 * i, 120, 20), d.items[fr])
    cmd = d.place_new("CommandButton", QRect(240, y + 100, 90, 32), None)
    QMessageBox.question = lambda *a, **k: QMessageBox.StandardButton.Yes
    d.ask_control_array = lambda name: True     # "... Do you want to create a control array?" Yes (its own box, 8.7)
    d.select_control(cmd)
    d.copy_selection()
    d.paste()
    d.move_item(d.items[d.primary.ctl], 5040, (y + 100) * 15)
    w.project.vbp.set("Name", "Designed", quote=True)
    return ["Form1.frm"]


def build_d_props(w):
    """the Properties window and the Menu Editor: Form1 with properties of every kind, Form2 made
    again as the steps done in VB6 (ODHGIES.txt)"""
    d = w.current_designer()
    f = d.root
    d.set_client_size(9000, 6600)
    put = lambda cls, x, y, wd, h, parent=None: d.place_new(cls, QRect(x, y, wd, h), d.items[parent] if parent else None)  # noqa: E731
    setp(w, d, f, "Caption", "Δοκιμή ιδιοτήτων")
    setp(w, d, f, "BackColor", 0x00FFFFC0)
    setp(w, d, f, "KeyPreview", True)
    setp(w, d, f, "StartUpPosition", 2)
    setp(w, d, f, "Font", {"Name": "Tahoma", "Size": 9, "Charset": None, "Weight": 700, "Underline": False,
                           "Italic": False, "Strikethrough": False})
    lab = put("Label", 8, 8, 120, 32)
    setp(w, d, lab, "Caption", 'He said "hi"')
    setp(w, d, lab, "Alignment", 2)
    setp(w, d, lab, "BorderStyle", 1)
    setp(w, d, lab, "BackStyle", 0)
    setp(w, d, lab, "WordWrap", True)
    setp(w, d, lab, "ForeColor", 0x000000FF)
    txt = put("TextBox", 136, 8, 120, 24)
    setp(w, d, txt, "Text", "abc")
    setp(w, d, txt, "MaxLength", 10)
    setp(w, d, txt, "PasswordChar", "*")
    setp(w, d, txt, "Locked", True)
    setp(w, d, txt, "Alignment", 1)
    setp(w, d, txt, "ToolTipText", "tip")
    setp(w, d, txt, "Font", {"Name": "Courier New", "Size": 10, "Charset": None, "Weight": 400, "Underline": False,
                             "Italic": True, "Strikethrough": False})
    c1 = put("CommandButton", 264, 8, 96, 32)
    setp(w, d, c1, "Caption", "&OK")
    setp(w, d, c1, "Default", True)
    setp(w, d, c1, "Style", 1)
    setp(w, d, c1, "BackColor", 0x0080FF80)
    setp(w, d, c1, "MousePointer", 11)
    c2 = put("CommandButton", 368, 8, 96, 32)
    setp(w, d, c2, "Caption", "Cancel")
    setp(w, d, c2, "Cancel", True)
    setp(w, d, c2, "TabIndex", 0)
    ch1 = put("CheckBox", 8, 48, 120, 24)
    setp(w, d, ch1, "Value", 1)
    setp(w, d, ch1, "Alignment", 1)
    ch2 = put("CheckBox", 136, 48, 120, 24)
    setp(w, d, ch2, "Value", 2)
    op = put("OptionButton", 264, 48, 120, 24)
    setp(w, d, op, "Value", True)
    cb = put("ComboBox", 392, 48, 120, 21)
    setp(w, d, cb, "Style", 2)
    setp(w, d, cb, "Sorted", True)
    lst = put("ListBox", 8, 80, 120, 80)
    setp(w, d, lst, "MultiSelect", 2)
    setp(w, d, lst, "Columns", 2)
    setp(w, d, lst, "IntegralHeight", False)
    setp(w, d, lst, "Height", 1200)
    hs = put("HScrollBar", 136, 80, 160, 17)
    for k, v in (("Min", 1), ("Max", 100), ("Value", 50), ("LargeChange", 10), ("SmallChange", 2)):
        setp(w, d, hs, k, v)
    vs = put("VScrollBar", 304, 80, 17, 80)
    setp(w, d, vs, "TabStop", False)
    tm = put("Timer", 336, 80, 28, 28)
    setp(w, d, tm, "Interval", 500)
    setp(w, d, tm, "Enabled", False)
    pic = put("PictureBox", 8, 176, 120, 96)
    setp(w, d, pic, "Appearance", 0)
    setp(w, d, pic, "ScaleMode", 3)
    setp(w, d, pic, "AutoRedraw", True)
    setp(w, d, pic, "BackColor", 0x80000005)
    shp = put("Shape", 136, 176, 96, 96)
    setp(w, d, shp, "Shape", 3)
    setp(w, d, shp, "FillStyle", 0)
    setp(w, d, shp, "FillColor", 0x000080FF)
    setp(w, d, shp, "BorderWidth", 3)
    setp(w, d, shp, "BorderStyle", 2)
    ln = put("Line", 240, 176, 96, 48)
    setp(w, d, ln, "BorderColor", 0x00FF0000)
    setp(w, d, ln, "BorderWidth", 2)
    setp(w, d, ln, "DrawMode", 7)
    img = put("Image", 344, 176, 64, 64)
    setp(w, d, img, "Stretch", True)
    setp(w, d, img, "BorderStyle", 1)
    fr = put("Frame", 8, 288, 200, 96)
    setp(w, d, fr, "Caption", "Group")
    setp(w, d, fr, "ForeColor", 0x8000000D)
    op2 = put("OptionButton", 16, 312, 160, 24, fr)
    setp(w, d, op2, "Caption", "Inside")
    fl = put("FileListBox", 216, 288, 120, 80)
    setp(w, d, fl, "Pattern", "*.txt")
    setp(w, d, fl, "Hidden", True)
    setp(w, d, fl, "ReadOnly", False)
    t2 = put("TextBox", 344, 288, 120, 24)
    setp(w, d, t2, "(Name)", "txtName")
    setp(w, d, t2, "Index", 0)
    c3 = put("CommandButton", 344, 320, 96, 32)
    setp(w, d, c3, "Visible", False)
    setp(w, d, c3, "Enabled", False)
    menus(d, [(0, "&File", "mnuFile", {}),
              (1, "&New", "mnuNew", {"shortcut": "^N"}),
              (1, "&Open...", "mnuOpen", {"shortcut": "{F12}"}),
              (1, "-", "mnuSep1", {}),
              (1, "&Recent", "mnuRecent", {}),
              (2, "Item 0", "mnuItem", {"index": 0}),
              (2, "Item 1", "mnuItem", {"index": 1}),
              (1, "E&xit", "mnuExit", {"shortcut": "%{BKSP}"}),
              (0, "&Edit", "mnuEdit", {"negotiate": 2}),
              (1, "&Copy", "mnuCopy", {"shortcut": "^{INSERT}", "checked": True}),
              (1, "&Delete", "mnuDelete", {"shortcut": "{DEL}", "enabled": False}),
              (1, "&Hidden", "mnuHidden", {"visible": False, "helpid": 5}),
              (1, "Run", "mnuRun", {"shortcut": "+^{F5}"}),
              (0, "&Window", "mnuWindow", {"windowlist": True})])
    # the menu bar took 20 pixels of the client area (Windows keeps the window): back to 9000 x 6600
    setp(w, d, f, "Top", -405)
    setp(w, d, f, "Height", 7470)
    # Form2: the steps done in VB6
    m2 = w.project.add_module("Form", "Form2")
    w.fill_tree()
    d2 = w.open_designer(m2)
    setp(w, d2, d2.root, "BorderStyle", 3)
    menus(d2, [(0, "&File", "mnuFile", {}), (1, "E&xit", "mnuExit", {"shortcut": "^Q"})])
    setp(w, d2, d2.root, "Width", 6000)
    setp(w, d2, d2.root, "Height", 4500)
    w.project.vbp.set("Name", "Props", quote=True)
    return ["Form1.frm", "Form2.frm"]


def pic(name):
    with open(os.path.join(HERE, "pictures", name), "rb") as f:
        return f.read()


def build_d_frx(w):
    """the .frx: pictures, List/ItemData, strings with line breaks and long strings; Form2 made
    again as the steps done in VB6 (ODHGIES.txt)"""
    d = w.current_designer()
    f = d.root
    d.set_client_size(8400, 6000)
    put = lambda cls, x, y, wd, h: d.place_new(cls, QRect(x, y, wd, h), None)  # noqa: E731
    setp(w, d, f, "Icon", pic("test.ico"))
    setp(w, d, f, "Picture", pic("test.bmp"))
    pb = put("PictureBox", 8, 8, 96, 64)
    setp(w, d, pb, "Picture", pic("test.bmp"))
    im1 = put("Image", 112, 8, 48, 48)
    setp(w, d, im1, "Picture", pic("test.gif"))
    im2 = put("Image", 168, 8, 48, 48)
    setp(w, d, im2, "Stretch", True)
    setp(w, d, im2, "Picture", pic("test.jpg"))
    c1 = put("CommandButton", 224, 8, 96, 48)
    setp(w, d, c1, "Style", 1)
    setp(w, d, c1, "Picture", pic("test.ico"))
    setp(w, d, c1, "Caption", "Pic")
    t0 = put("TextBox", 328, 8, 96, 24)
    setp(w, d, t0, "MousePointer", 99)
    setp(w, d, t0, "MouseIcon", pic("test.ico"))
    l1 = put("ListBox", 8, 80, 120, 80)
    setp(w, d, l1, "List", ["Alpha", "Beta", "Γάμμα"])
    setp(w, d, l1, "ItemData", [10, 20, -1])
    setp(w, d, l1, "DragIcon", pic("test.ico"))
    l2 = put("ListBox", 136, 80, 120, 80)
    setp(w, d, l2, "List", ["one", "two", "three"])
    cb = put("ComboBox", 264, 80, 120, 21)
    setp(w, d, cb, "List", ["x", "y"])
    setp(w, d, cb, "ItemData", [1000, 70000])
    t1 = put("TextBox", 8, 176, 160, 48)
    setp(w, d, t1, "MultiLine", True)
    setp(w, d, t1, "Text", "Line one\r\nΓραμμή δύο")
    t2 = put("TextBox", 176, 176, 200, 48)
    setp(w, d, t2, "MultiLine", True)
    setp(w, d, t2, "Text", "A" * 150 + "B" * 150)
    lb1 = put("Label", 8, 240, 160, 40)
    setp(w, d, lb1, "Caption", "First\r\nSecond")
    lb2 = put("Label", 176, 240, 200, 80)
    setp(w, d, lb2, "Caption", "Long caption. " * 40)
    c2 = put("CommandButton", 384, 176, 96, 32)
    setp(w, d, c2, "Tag", "T" * 260)
    setp(w, d, c2, "ToolTipText", "Short tip")
    # Form2: the steps done in VB6
    m2 = w.project.add_module("Form", "Form2")
    w.fill_tree()
    d2 = w.open_designer(m2)
    pb2 = d2.place_new("PictureBox", QRect(8, 8, 96, 64), None)
    setp(w, d2, pb2, "Picture", pic("test.bmp"))
    ls2 = d2.place_new("ListBox", QRect(112, 8, 120, 80), None)
    setp(w, d2, ls2, "List", ["One", "Two", "Three"])
    setp(w, d2, ls2, "ItemData", [5, -7, 300])
    tx2 = d2.place_new("TextBox", QRect(8, 96, 160, 48), None)
    setp(w, d2, tx2, "MultiLine", True)
    setp(w, d2, tx2, "Text", "")
    w.props.show_object(d2.module, d2, tx2)
    w.props._popup_done(V.prop_def("TextBox", "Text"), "A\nB")     # the drop-down of the Text
    w.project.vbp.set("Name", "Frx", quote=True)
    return ["Form1.frm", "Form1.frx", "Form2.frm", "Form2.frx"]


def build_d_data(w):
    """a Data control and the controls bound to it, set in the Properties window (phase 4.4)"""
    d = w.current_designer()
    d.set_client_size(6000, 3600)
    data = d.place_new("Data", QRect(8, 200, 300, 24), None)
    txt = d.place_new("TextBox", QRect(8, 8, 160, 24), None)
    chk = d.place_new("CheckBox", QRect(8, 40, 160, 24), None)
    lbl = d.place_new("Label", QRect(8, 72, 160, 24), None)
    lst = d.place_new("ListBox", QRect(200, 8, 160, 100), None)
    dc = d.items[data].ctl
    for name, value in (("Caption", "Titles"), ("DatabaseName", "BIBLIO.MDB"), ("RecordSource", "Titles"),
                        ("RecordsetType", 2), ("EOFAction", 2), ("BOFAction", 1), ("ReadOnly", True),
                        ("Exclusive", True), ("Options", 64), ("Connect", "Access 2000;"), ("Align", 2)):
        setp(w, d, dc, name, value)
    for it, field in ((txt, "Title"), (chk, "PubID"), (lbl, "ISBN"), (lst, "Subject")):
        setp(w, d, d.items[it].ctl, "DataSource", "Data1")
        setp(w, d, d.items[it].ctl, "DataField", field)
    setp(w, d, dc, "DefaultType", 1)
    setp(w, d, dc, "DefaultCursorType", 1)
    w.project.vbp.set("Name", "DataForm", quote=True)
    return ["Form1.frm"]


def build_d_dev(w):
    """the Multimedia control and the Comm control, from Project > Components (phase 4.5)"""
    w.project.set_components(["MCI", "MSCommLib"])
    w.refresh_toolbox() if hasattr(w, "refresh_toolbox") else None
    d = w.current_designer()
    d.set_client_size(6000, 3600)
    mm = d.place_new("MCI.MMControl", QRect(8, 8, 2, 2), None)
    mm2 = d.place_new("MCI.MMControl", QRect(8, 60, 2, 2), None)
    com = d.place_new("MSCommLib.MSComm", QRect(300, 8, 2, 2), None)
    com2 = d.place_new("MSCommLib.MSComm", QRect(350, 8, 2, 2), None)
    c2 = d.items[mm2].ctl
    for name, value in (("DeviceType", "WaveAudio"), ("FileName", "C:\\Windows\\Media\\tada.wav"),
                        ("StopVisible", False), ("EjectVisible", False), ("RecordVisible", False),
                        ("PlayEnabled", True), ("AutoEnable", False), ("UpdateInterval", 500),
                        ("Orientation", 0), ("Shareable", True), ("Frames", 5)):
        setp(w, d, c2, name, value)
    cc = d.items[com2].ctl
    for name, value in (("CommPort", 2), ("Settings", "19200,e,7,1"), ("RThreshold", 1), ("SThreshold", 2),
                        ("InputLen", 10), ("InBufferSize", 2048), ("Handshaking", 2), ("RTSEnable", True),
                        ("DTREnable", False), ("EOFEnable", True), ("NullDiscard", True), ("InputMode", 1),
                        ("ParityReplace", "#")):
        setp(w, d, cc, name, value)
    w.project.vbp.set("Name", "DevForm", quote=True)
    return ["Form1.frm"]


def build_d_dev2(w):
    """the other properties of the Multimedia and the Comm control: their order in the .frm"""
    w.project.set_components(["MCI", "MSCommLib"])
    d = w.current_designer()
    d.set_client_size(6000, 3600)
    mm = d.place_new("MCI.MMControl", QRect(8, 8, 2, 2), None)
    c = d.items[mm].ctl
    for name, value in (("Orientation", 1), ("RecordMode", 0), ("BorderStyle", 0), ("Silent", True),
                        ("Enabled", False), ("PrevEnabled", True), ("StepEnabled", True), ("EjectEnabled", True),
                        ("NextVisible", False), ("BackVisible", False), ("PauseVisible", False),
                        ("UpdateInterval", 0), ("Frames", 2), ("MousePointer", 2)):
        setp(w, d, c, name, value)
    for x, settings, extra in ((300, "9600,n,8,2", (("OutBufferSize", 1024), ("ParityReplace", ""), ("SThreshold", 1))),
                               (350, "1200,o,6,1.5", (("CommPort", 16), ("InputLen", 1), ("DTREnable", False))),
                               (400, "115200,m,5,1", (("ParityReplace", "x"), ("Handshaking", 3))),
                               (450, "300,s,4,2", ())):
        k = d.place_new("MSCommLib.MSComm", QRect(x, 8, 2, 2), None)
        cc = d.items[k].ctl
        setp(w, d, cc, "Settings", settings)
        for name, value in extra:
            setp(w, d, cc, name, value)
    w.project.vbp.set("Name", "DevForm2", quote=True)
    return ["Form1.frm"]


TESTS = {"d_all": build_d_all, "d_props": build_d_props, "d_frx": build_d_frx, "d_data": build_d_data,
         "d_dev": build_d_dev, "d_dev2": build_d_dev2}


def run(name, build):
    out = os.path.join(work, name)
    shutil.rmtree(out, ignore_errors=True)
    os.makedirs(out)
    w = MainWindow()
    w.show()
    files = build(w)
    vbp = {"d_all": "Designed.vbp", "d_props": "Props.vbp", "d_frx": "Frx.vbp", "d_data": "DataForm.vbp",
           "d_dev": "DevForm.vbp", "d_dev2": "DevForm2.vbp"}[name]
    names = iter([os.path.join(out, f) for f in files if f.endswith(".frm")] + [os.path.join(out, vbp)])
    QFileDialog.getSaveFileName = lambda *a, **k: (next(names), "")
    w.save_project_as()
    w.project.dirty = False
    for m in w.project.modules:
        m.dirty = False
    w.close()
    ok = True
    for f in files:
        got = open(os.path.join(out, f), "rb").read()
        expf = os.path.join(HERE, name, "expected", f)
        if not os.path.exists(expf):
            print(f"  {f}: no VB6 file yet (made {os.path.join(out, f)})")
            ok = False
            continue
        exp = open(expf, "rb").read()
        if got != exp:
            ok = False
            print(f"  {f} differs from VB6:")
            import difflib
            if f.endswith(".frx"):
                n = next((i for i in range(min(len(got), len(exp))) if got[i] != exp[i]), min(len(got), len(exp)))
                print(f"    sizes {len(exp)} (VB6) / {len(got)}; first difference at {n:#x}")
                continue
            for line in difflib.unified_diff(exp.decode("cp1253").splitlines(), got.decode("cp1253").splitlines(),
                                             "vb6", "designer", lineterm="", n=1):
                print("   ", line)
    return ok


def main():
    failed = 0
    only = sys.argv[1:]
    for name, build in TESTS.items():
        if only and name not in only:
            continue
        ok = run(name, build)
        print(f"{'PASS' if ok else 'FAIL'}  {name}")
        failed += not ok
    n = len([t for t in TESTS if not only or t in only])
    print(f"{n - failed}/{n} passed")
    os._exit(1 if failed else 0)


main()
