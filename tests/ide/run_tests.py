#!/usr/bin/env python3
"""Tests of the IDE of LinBASIC Studio (phase 4.3): what the user does with the menus and the windows,
checked in the files the IDE writes and in the program it runs.

    t_resize   a control placed in the designer is sized with the mouse (one drag, many moves)
    t_types    New Project of each type (Standard EXE, ActiveX EXE / DLL / Control), Instancing, Make
    t_uctl     a user control: its designer, the Toolbox, placed on a form, its properties, F5
    t_group    a project group: Add Project, References, Save Project Group, F5, open again
    t_dsr      Add Data Environment / Add Data Report: a connection and a hierarchy of commands (the
               fields from the database), the report's sections from it (Retrieve Structure), its
               controls, the code window's objects, Save (the .dsr / .dsx as VB6 writes them), F5
    t_refactor Definition, Find References, Rename Symbol (a variable, a procedure, a control with its
               event procedures, a class with its WithEvents variable, a module), Extract Procedure, F5
    t_ondemand Compile On Demand: a procedure with a compile error stops the run only when it is called
    t_dsr2     a command with parameters (from the ? of its SQL), fields dragged from the Data Environment
               onto the report, the events of a control of a component in the code window, F5
    t_dbgrid   a DBGrid and a DBList on two Data controls (phase 7.6): the properties of the grid in its
               OleObjectBlob, its columns (the Columns dialog), DataSource / RowSource in the Bindings of the
               .frx, Save, F5
    t_dfmt     the DataFormat of two TextBoxes (phase 7.7): the Format page (Number, Boolean), BeginProperty
               DataFormat in the .frm, the page opened again on it, F5 (the DataFormat of the controls)
    t_dgcols   a DataGrid and a DataList on an ADO Data control (phase 7.7): the columns of the design of the grid
               with their DataFormat (the Columns dialog), DataSource in the Bindings of the .frx (dispid 42), the
               DataFormat of the DataList (an .ocx: the block with the CLSID), Save, F5
    t_coolbar  a CoolBar (phase 7.4): placed on the form (Align Top), a TextBox in it, its bands (the Bands dialog):
               the child of a band laid out as the runtime does it, the flat properties of the bands in the .frm, F5
    t_rdc      a RemoteData control (phase 7.8) and a TextBox bound to it: its properties in the .frm as the VB6 IDE
               writes them (RecordsetType, RecordSource ...), F5 over the SQLite ODBC driver (when it is installed)
    t_objbrowser  the Object Browser (phase 8.1): libraries, classes, members, the details as VB6 shows them,
               Search, the project's members, View Definition, Definition on MsgBox, Go Back / Forward
    t_debug82  the debugger of phase 8.2: Set Next Statement (Ctrl+F9; a line of another procedure is refused),
               Add Watch with its Context and Watch Type (Break When Value Is True), a watch out of its context,
               Disable All Breakpoints, Error Trapping (Toggle > Break on All Errors), Alt+F5 at an error,
               Edit and Continue (an edit in break mode resets the project)
    t_editor83 the code window of phase 8.3: Comment Block / Uncomment Block and Add Procedure as the VB6 IDE wrote
               them (VB6 oracle DEV-8/ed1, five stages), one undo for each, the Add Procedure dialog and its messages,
               Edit > Bookmarks across the modules, the Edit toolbar
    t_procattr83  Tools > Procedure Attributes (phase 8.3): the dialog on a class (what is on for a Sub and for a
               Property), Procedure ID, Hide this member, User Interface Default, Property Category, Data Binding, the
               Attribute lines as the VB6 IDE wrote them (VB6 oracle DEV-8/ed2), read back, a standard module
    t_print83  File > Print (phase 8.3): the pages as the VB6 IDE printed them (VB6 oracle DEV-8/pr1, through the XPS
               writer): "<module> - <page>" and an empty line on every page, 79 lines on A4, Current Project in the
               order of the project, Form As Text, Form Image, Selection; the Print dialog; a PDF of 5 pages
    t_projprops85  Project > Properties (phase 8.5): the five pages, what is on for each Project Type, the .vbp lines as
               the VB6 IDE wrote them (VB6 oracle DEV-8/pp1: General, Make, Compile, Advanced Optimizations, Component,
               Debugging, Version Compatibility and its message, a new Project Type and "Start Mode has been changed")
    t_formdesign84  the form designer (phase 8.4): Horizontal / Vertical Spacing, Size to Grid, Lock Controls, the arrow
               keys, the tab order, the Form Layout window, the grid of Tools > Options, the .frm as the VB6 IDE wrote it
               at each step (VB6 oracle DEV-8/fd1)
    t_drep     a DataRepeater (phase 7.5): its properties, RepeatedControlName and RepeaterBindings set in the Properties
               window; the .frm as the VB6 IDE wrote the same control (VB6 oracle f_drep3), the .vbp

    python3 tests/ide/run_tests.py [names...]
"""
import faulthandler
import os
import shutil
import sys
import tempfile
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "ide"))
os.environ.setdefault("QT_QPA_PLATFORM", "windows" if os.name == "nt" else "offscreen")   # the programs F5 runs: the Qt of the runtime
work = os.path.join(HERE, "work")
shutil.rmtree(work, ignore_errors=True)
os.makedirs(os.path.join(work, ".xdg"), mode=0o700)
os.environ["XDG_RUNTIME_DIR"] = os.path.join(work, ".xdg")
os.environ["LBVC_REGISTRY"] = os.path.join(work, "components.json")     # not the user's

from lbstudio import paths, settings  # noqa: E402

paths.set_config_dir(os.path.join(work, "config"))
paths.ensure_user_files()
settings.init()
settings.S["codepage"] = "cp1252"

from PyQt6.QtCore import QRect, QPoint, Qt  # noqa: E402
from PyQt6.QtTest import QTest  # noqa: E402
from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox, QToolBar  # noqa: E402
from PyQt6.Qsci import QsciScintilla as S  # noqa: E402

app = QApplication(sys.argv[:1])
QMessageBox.information = lambda *a, **k: None
QMessageBox.warning = lambda *a, **k: None

from lbstudio.mainwindow import MainWindow  # noqa: E402
from lbstudio import vbprops as V  # noqa: E402


def saves(*files):
    """the answers of the Save As dialogs, in the order the IDE asks"""
    it = iter(files)
    QFileDialog.getSaveFileName = lambda *a, **k: (next(it), "")


def wait(w, seconds=60):
    """until F5 or Make is over"""
    t0 = time.time()
    while time.time() - t0 < seconds:
        app.processEvents()
        time.sleep(0.03)
        if w.runner.state == "design" and time.time() - t0 > 0.3:
            return
    raise AssertionError("the program did not end")


def done(w):
    for q in (w.group.projects if w.group else [w.project]):
        q.dirty = False
        for m in q.modules:
            m.dirty = False
    if w.group:
        w.group.dirty = False
    w.close()


def code(w, m, text):
    """the code of a module, typed in its code window"""
    w.open_code(m).editor.setText(text)


def t_resize(d):
    w = MainWindow()
    w.show()
    f = w.current_designer()
    c = f.place_new("CommandButton", QRect(20, 20, 80, 30), None)
    app.processEvents()
    item = f.items[c]
    r0 = item.rect_twips()
    surf = item.parentWidget()
    h = next(x for x in surf.handles if x.pos_name == "e")
    p0 = h.pos() + QPoint(3, 3)
    QTest.mousePress(h, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier, QPoint(3, 3))
    from PyQt6.QtGui import QMouseEvent
    from PyQt6.QtCore import QEvent, QPointF
    for i in range(1, 11):                          # the mouse moves on; the same handle has it
        lp = QPointF(p0 + QPoint(i * 4, 0) - h.pos())
        QApplication.sendEvent(h, QMouseEvent(QEvent.Type.MouseMove, lp, QPointF(h.mapToGlobal(lp.toPoint())),
                                              Qt.MouseButton.NoButton, Qt.MouseButton.LeftButton,
                                              Qt.KeyboardModifier.NoModifier))
        app.processEvents()
    QTest.mouseRelease(h, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier, p0 + QPoint(40, 0) - h.pos())
    app.processEvents()
    r1 = item.rect_twips()
    assert r1.width() == r0.width() + 600, (r0, r1)
    assert h in surf.handles
    done(w)


def t_types(d):
    w = MainWindow()
    w.show()
    want = {"Exe": ("Form1", 1), "OleExe": ("Class1", 5), "OleDll": ("Class1", 5), "Control": ("UserControl1", 5)}
    for ptype, (first, inst) in want.items():
        w.new_project(ptype)
        p = w.project
        assert p.ptype == ptype and p.modules[0].name == first, (ptype, [m.name for m in p.modules])
        assert p.modules[0].instancing() == inst, (ptype, p.modules[0].instancing())
        st = (p.vbp.get("Startup") or "").strip('"')
        assert st == ("Form1" if ptype == "Exe" else "(None)"), (ptype, st)
    # an ActiveX DLL: Instancing in the Properties window, Save, Make (it registers the project)
    w.new_project("OleDll")
    p = w.project
    m = p.modules[0]
    w.select_in_tree(m)
    w.fill_properties()
    inst = next(r for r in w.props.rows if r.name == "Instancing")
    assert [n for n, _ in V.enum_items(inst)] == [1, 2, 5, 6]
    w.props.commit(inst, 6)
    assert m.instancing() == 6
    code(w, m, "Public Function Twice(ByVal n As Long) As Long\n    Twice = n * 2\nEnd Function\n")
    saves(os.path.join(d, "Class1.cls"), os.path.join(d, "MyLib.vbp"))
    assert w.save_project_as()
    cls = open(os.path.join(d, "Class1.cls"), encoding="cp1252").read()
    assert "Attribute VB_GlobalNameSpace = True" in cls and "Attribute VB_Exposed = True" in cls and "MultiUse = -1" in cls, cls
    vbp = open(os.path.join(d, "MyLib.vbp"), encoding="cp1252").read()
    assert vbp.startswith("Type=OleDll") and 'Startup="(None)"' in vbp and "Class=Class1; Class1.cls" in vbp, vbp
    assert w.a_make.text() == "Ma&ke Project1.dll...", w.a_make.text()
    w.make_exe()
    wait(w)
    import json
    reg = json.load(open(os.environ["LBVC_REGISTRY"]))
    assert reg["project1"]["vbp"] == os.path.join(d, "MyLib.vbp") and reg["project1"]["exe"] == "Project1.dll", reg
    # Instancing a Standard EXE allows: Private only
    w.new_project("Exe")
    w.add_module("Class")
    c = w.project.module("Class1")
    w.select_in_tree(c)
    w.fill_properties()
    inst = next(r for r in w.props.rows if r.name == "Instancing")
    assert [n for n, _ in V.enum_items(inst)] == [1] and c.instancing() == 1
    done(w)


def t_uctl(d):
    w = MainWindow()
    w.show()
    p = w.project
    w.add_module("UserControl")
    uc = p.module("UserControl1")
    u = w.current_designer()
    assert u.module is uc and u.frame_w.frame() == (0, 0, 0)     # no window frame around a user control
    u.set_client_size(2400, 900)
    u.place_new("CommandButton", QRect(8, 8, 100, 30), None)
    code(w, uc, "Public Event Pressed(ByVal n As Long)\n\nPrivate mN As Long\n\n"
                "Public Property Get Count() As Long\n    Count = mN\nEnd Property\n\n"
                "Public Property Let Count(ByVal v As Long)\n    mN = v\nEnd Property\n\n"
                "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)\n"
                "    mN = PropBag.ReadProperty(\"Count\", 0)\nEnd Sub\n\n"
                "Private Sub Command1_Click()\n    mN = mN + 1\n    RaiseEvent Pressed(mN)\nEnd Sub\n")
    w.fill_tree()
    assert "Project1.UserControl1" in w.toolbox.buttons, list(w.toolbox.buttons)
    f = w.open_designer(p.module("Form1"))
    c = f.place_new("Project1.UserControl1", QRect(16, 16, 0, 0), None)
    assert f.place_new("Project1.UserControl1", QRect(0, 0, 0, 0), None) is not None
    assert u.place_new("Project1.UserControl1", QRect(0, 0, 0, 0), None) is None   # not inside itself
    w.props.show_object(f.module, f, c)
    names = [r.name for r in w.props.rows]
    assert "Count" in names and "Left" in names and "Tag" in names, names
    w.props.commit(V.prop_def("UserControl1", "Count"), 5)
    cw = w.open_code(p.module("Form1"))
    i = cw.obj_box.findText("UserControl11")
    cw._fill_procs("UserControl11")
    evs = [cw.proc_box.itemText(k) for k in range(cw.proc_box.count())]
    assert "Pressed" in evs and "GotFocus" in evs, evs
    code(w, p.module("Form1"), "Private Sub Form_Load()\n    Dim f As Integer\n    f = FreeFile\n"
                               "    Open App.Path & \"/out.txt\" For Output As #f\n"
                               "    Print #f, UserControl11.Count; TypeName(UserControl11); UserControl11.Left\n"
                               "    Close #f\n    Unload Me\nEnd Sub\n")
    w.add_module("PropertyPage")                # design time only: it opens in its designer, F5 leaves it out
    pg = w.current_designer()
    assert pg is not None and pg.module.kind == "PropertyPage" and pg.frame_w.frame() == (0, 0, 0)
    pg.place_new("CheckBox", QRect(8, 8, 120, 20), None)
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "UserControl1.ctl"), os.path.join(d, "PropertyPage1.pag"),
          os.path.join(d, "UCDemo.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(d, "Form1.frm"), encoding="cp1252").read()
    assert "Begin Project1.UserControl1 UserControl11" in frm and "_ExtentX        =   4233" in frm and \
        "Count           =   5" in frm, frm
    vbp = open(os.path.join(d, "UCDemo.vbp"), encoding="cp1252").read()
    assert "UserControl=UserControl1.ctl" in vbp and "PropertyPage=PropertyPage1.pag" in vbp, vbp
    pag = open(os.path.join(d, "PropertyPage1.pag"), encoding="cp1252", newline="").read()
    assert pag.startswith("VERSION 5.00\r\nBegin VB.PropertyPage PropertyPage1") and "Begin VB.CheckBox Check1" in pag, pag
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "out.txt")).read()
    assert out.split() == ["5", "UserControl1", "240"], out
    done(w)


def t_group(d):
    os.makedirs(os.path.join(d, "App"))
    os.makedirs(os.path.join(d, "Lib"))
    w = MainWindow()
    w.show()
    app_p = w.project
    w.add_module("Module")
    code(w, app_p.module("Module1"),
         "Sub Main()\n    Dim c As New Project2.Class1\n    Dim f As Integer\n    f = FreeFile\n"
         "    Open App.Path & \"/out.txt\" For Output As #f\n    Print #f, c.Hello(\"group\"); \" \"; TypeName(c)\n"
         "    Close #f\nEnd Sub\n")
    app_p.vbp.set("Startup", "Sub Main", quote=True)
    lib = w.add_project("OleDll")
    assert w.group is not None and w.group.startup is app_p and w.project is lib
    cls = lib.module("Class1")
    cw = w.open_code(cls)
    cw.editor.setText(cw.editor.text() + "\nPublic Function Hello(ByVal s As String) As String\n"
                                         "    Hello = \"hello \" & s\nEnd Function\n")
    saves(*[os.path.join(d, x) for x in ("App/Form1.frm", "App/Module1.bas", "App/App.vbp", "Lib/Class1.cls",
                                         "Lib/Lib.vbp", "Both.vbg")])
    assert w.save_group()
    w.set_active(app_p)
    from lbstudio.dialogs import ReferencesDialog
    dlg = ReferencesDialog(w, app_p, w.group)
    mine = [(it, path) for it, path in dlg.items if "project of the group" in it.text()]
    assert len(mine) == 1, [it.text() for it, _ in dlg.items]
    mine[0][0].setCheckState(Qt.CheckState.Checked)
    dlg._ok()
    assert w.save_project()
    vbg = open(os.path.join(d, "Both.vbg")).read().splitlines()
    assert vbg == ["VBGROUP 5.0", "StartupProject=App\\App.vbp", "Project=Lib\\Lib.vbp"], vbg
    vbp = open(os.path.join(d, "App", "App.vbp"), encoding="cp1252").read()
    assert "Reference=*\\A..\\Lib\\Lib.vbp" in vbp, vbp
    w.set_active(lib)                           # F5 runs the startup project, whichever is active
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "App", "out.txt")).read().strip()
    assert out == "hello group Class1", out
    # File > Make Project Group... (DEV-8/me2 pictures 26, 27): both projects chosen, the default options
    w._update_file_actions()
    assert w.a_make_group.isEnabled()
    from PyQt6.QtCore import QTimer as _QT
    seen = {}

    def _build():
        dlg = w._last_make_group
        seen["rows"] = [dlg.list.item(i).text() for i in range(dlg.list.count())]
        seen["chosen"] = dlg.chosen()
        seen["default"] = dlg.default_opts.isChecked()
        seen["title"] = dlg.windowTitle()
        dlg.accept()
    _QT.singleShot(0, _build)
    w.make_group()
    assert seen == {"rows": ["Project1", "Project2"], "chosen": [0, 1], "default": True,
                    "title": "LinBASIC Studio"}, seen
    t0 = time.time()
    while (w._group_jobs is not None or w.runner.state != "design") and time.time() - t0 < 180:
        app.processEvents()
        time.sleep(0.05)
    assert w._group_jobs is None and w.runner.state == "design"
    exe = os.path.join(d, "App", "Project1" + (".exe" if os.name == "nt" else ""))
    assert os.path.exists(exe), os.listdir(os.path.join(d, "App"))
    assert "Project2: making" in w.output.toPlainText() and "Project2 is registered" in w.output.toPlainText()
    w.group.projects[0].dirty = False
    w.close_project()
    w.open_project(os.path.join(d, "Both.vbg"))
    assert [q.name for q in w.group.projects] == ["Project1", "Project2"] and w.group.startup.name == "Project1"
    w.set_startup(w.group.projects[1])
    assert w.group.startup.name == "Project2" and w.group.dirty
    assert w.a_close_group.isEnabled()
    w.group.dirty = False
    w.set_active(w.group.projects[0])
    w.close_project()                           # in a group: only the active project closes
    assert w.group is not None and [q.name for q in w.group.projects] == ["Project2"], w.group.projects
    w.group.dirty = False
    w.close_group()                             # File > Close Project Group: everything
    assert w.project is None and w.group is None and w.tree.topLevelItemCount() == 0
    assert w.dock_project.windowTitle() == "Project Explorer"
    w.open_project(os.path.join(d, "Both.vbg"))
    assert len(w.group.projects) == 2           # the .vbg was not saved: both projects are still in it
    done(w)


def t_ondemand(d):
    """Compile On Demand (as the VB6 IDE, VB-TESTS ide-6.8): F5 runs a program with a compile error in a
    procedure; the program runs until that procedure is called, then the IDE shows the compile error at its line"""
    w = MainWindow()
    w.show()
    p = w.project
    w.add_module("Module")
    code(w, p.module("Module1"),
         "Option Explicit\n\nSub Main()\n    Log \"start\"\n    GoodA\n    Log \"after GoodA\"\n    BadA\n"
         "    Log \"after BadA\"\nEnd Sub\n\nPrivate Sub GoodA()\n    Log \"in GoodA\"\nEnd Sub\n\n"
         "Private Sub BadA()\n    Log \"in BadA\"\n    undefinedA = 1\nEnd Sub\n\n"
         "Private Sub NeverCalled()\n    undefinedB = 2\nEnd Sub\n\n"
         "Public Sub Log(ByVal s As String)\n    Dim f As Integer\n    f = FreeFile\n"
         "    Open App.Path & \"/out.txt\" For Append As #f\n    Print #f, s\n    Close #f\nEnd Sub\n")
    p.vbp.set("Startup", "Sub Main", quote=True)
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Module1.bas"), os.path.join(d, "OnDemand.vbp"))
    assert w.save_project()
    errs = []
    w.runner.compile_error.connect(lambda f, line, col, msg: errs.append((os.path.basename(f), line, msg)))
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "out.txt")).read().split()
    assert out == ["start", "in", "GoodA", "after", "GoodA"], out
    assert len(errs) == 1 and errs[0][0] == "Module1.bas" and "undefineda" in errs[0][2].lower(), errs
    m = p.module("Module1")
    assert errs[0][1] - m.code_offset() == 17, (errs, m.code_offset())   # the line of undefinedA = 1
    done(w)


MAKE_DB = """Sub MakeDb()
    Dim e As Object, db As Object, p As String
    p = App.Path & "\\ide.mdb"
    If Dir(p) <> "" Then Exit Sub
    Set e = CreateObject("DAO.DBEngine.36")
    Set db = e.CreateDatabase(p, ";LANGID=0x0409;CP=1252;COUNTRY=0", 64)
    db.Execute "CREATE TABLE Customers (CustID TEXT(5) CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(30), Country TEXT(15))"
    db.Execute "INSERT INTO Customers VALUES ('ALFKI', 'Alfreds', 'Germany')"
    db.Execute "INSERT INTO Customers VALUES ('BONAP', 'Bon app', 'France')"
    db.Execute "CREATE TABLE Orders (OrderID LONG CONSTRAINT PrimaryKey PRIMARY KEY, CustID TEXT(5), Amount CURRENCY)"
    db.Execute "INSERT INTO Orders VALUES (10, 'ALFKI', 100.5)"
    db.Execute "INSERT INTO Orders VALUES (11, 'ALFKI', 20)"
    db.Execute "INSERT INTO Orders VALUES (12, 'BONAP', 7.25)"
    db.Close
End Sub
"""


def t_dsr(d):
    from lbstudio import dsrmodel as M
    w = MainWindow()
    w.show()
    p = w.project
    w.add_module("Module")
    code(w, p.module("Module1"), MAKE_DB + "\nSub Main()\n    MakeDb\nEnd Sub\n")
    p.vbp.set("Startup", "Sub Main", quote=True)
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Module1.bas"), os.path.join(d, "Rpt.vbp"))
    assert w.save_project_as()
    w.run_start()                               # the database of the designers
    wait(w)
    db = os.path.join(d, "ide.mdb")
    assert os.path.isfile(db)
    # ---- the DataEnvironment
    w.add_module("DataEnvironment")
    dem = p.module("DataEnvironment1")
    assert dem.kind == "Designer" and dem.designer_kind == "DataEnvironment"
    de = w.design_windows[dem].widget()
    conn = de.model.conns[0]
    de.cpage.show_conn(conn)
    de.cpage.cs.setText(f"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={db}")
    de.cpage.apply()
    de.add_command(False)
    cmd = de.model.cmds[0]
    pg = de.mpage
    pg.show_cmd(cmd)
    pg.name.setText("Customers")
    pg.rb_obj.setChecked(True)
    assert [pg.objname.itemText(i) for i in range(pg.objname.count())] == ["Customers", "Orders"]
    pg.objname.setEditText("Customers")
    pg.apply()
    assert cmd.fields == [("CustID", 202, 5, 255), ("Name", 202, 30, 255), ("Country", 202, 15, 255)], cmd.fields
    de.fill_tree(cmd)
    de.add_command(True)                        # a child command of Customers
    ch = de.model.cmds[1]
    assert ch.parent == "Customers"
    pg.show_cmd(ch)
    pg.name.setText("Orders")
    pg.rb_sql.setChecked(True)
    pg.sql.setPlainText("SELECT OrderID, CustID, Amount FROM Orders")
    pg.apply()
    pg.show_cmd(ch)
    pg.pfield.setCurrentText("CustID")
    pg.cfield.setCurrentText("CustID")
    pg._rel_add()
    pg.apply()
    assert ch.relations == [("CustID", "CustID")] and [f[0] for f in ch.fields] == ["OrderID", "CustID", "Amount"]
    pg.show_cmd(cmd)                            # an aggregate of the parent: the sum of the child's amounts
    pg._agg_row(("Total", "Orders", "Amount", 7))
    pg.apply()
    assert cmd.aggregates == [("Total", "Orders", "Amount", 7)], cmd.aggregates
    # ---- the DataReport
    w.add_module("DataReport")
    rm = p.module("DataReport1")
    rd = w.design_windows[rm].widget()
    assert [s.code for s in rd.model.sections] == [2, 4, 7]
    assert rd.data_environments() == ["DataEnvironment1"]
    rd._set_source("DataEnvironment1")
    assert rd.members() == ["Customers"]
    rd.model.datamember = "Customers"
    rd.retrieve_structure()
    assert [(s.code, s.name) for s in rd.model.sections] == [
        (2, "Section2"), (3, "Customers_Header"), (4, "Orders_Detail"), (5, "Customers_Footer"), (7, "Section3")], \
        [(s.code, s.name) for s in rd.model.sections]
    assert rd.members(all_levels=True) == ["Customers", "Orders"] and "Total" in rd.fields_of("Customers")
    secs = {s.name: s for s in rd.model.sections}

    def place(kind, sec, x, y, props):
        rd.tools.button([k for k, _ in __import__("lbstudio.dsrdesign", fromlist=["TOOLS"]).TOOLS].index(kind)).setChecked(True)
        body = next(b for s, _, b in rd.canvas.bands() if s is secs[sec])
        n = len(secs[sec].items)
        rd.create(secs[sec], QRect(body.left() + x, body.top() + y, 0, 0))
        if len(secs[sec].items) == n:
            return None
        it = secs[sec].items[-1]
        it.props.update(props)
        return it
    assert place("RptFunction", "Orders_Detail", 10, 2, {}) is None     # a Function only in a footer
    place("RptLabel", "Section2", 0, 0, {"Caption": "Orders"})
    place("RptTextBox", "Customers_Header", 0, 0, {"DataMember": "Customers", "DataField": "Name"})
    place("RptTextBox", "Orders_Detail", 96, 0, {"DataMember": "Orders", "DataField": "Amount"})
    place("RptTextBox", "Customers_Footer", 0, 0, {"DataMember": "Customers", "DataField": "Total"})
    fn = place("RptFunction", "Customers_Footer", 192, 0, {"DataMember": "Orders", "DataField": "Amount"})
    assert fn is not None and fn.name == "Function1"
    rd.changed()
    # the code windows: the objects of the designers and their events
    cw = w.open_code(dem)
    objs = [cw.obj_box.itemText(i) for i in range(cw.obj_box.count())]
    assert objs == ["(General)", "DataEnvironment", "Connection1", "rsCustomers"], objs
    cw._fill_procs("rsCustomers")
    evs = [cw.proc_box.itemText(k) for k in range(cw.proc_box.count())]
    assert "MoveComplete" in evs and cw.proc_box.itemData(evs.index("MoveComplete")).startswith("ByVal adReason As ADODB.EventReasonEnum"), evs
    cw = w.open_code(rm)
    cw._fill_procs("DataReport")
    evs = [cw.proc_box.itemText(k) for k in range(cw.proc_box.count())]
    assert "QueryClose" in evs and "Initialize" in evs, evs
    code(w, p.module("Module1"), MAKE_DB + "\nSub Main()\n    MakeDb\n"
         "    DataReport1.ExportReport \"key_def_Text\", App.Path & \"\\\\out.txt\", True, False\n"
         "    Unload DataReport1\nEnd Sub\n")
    assert w.save_project()
    vbp = open(os.path.join(d, "Rpt.vbp"), encoding="cp1252").read()
    assert "Designer=DataEnvironment1.Dsr" in vbp and "Designer=DataReport1.Dsr" in vbp and "msderun.dll" in vbp and \
        "MSDBRPTR.DLL" in vbp, vbp
    rpt = open(os.path.join(d, "DataReport1.Dsr"), encoding="cp1252", newline="").read()
    assert rpt.startswith("VERSION 5.00\r\nBegin {78E93846-85FD-11D0-8487-00A0C90DC8A9} DataReport1 \r\n"
                          "   Bindings        =   \"DataReport1.dsx\":0000\r\n"), rpt[:200]
    assert "Attribute VB_PredeclaredId = True" in rpt and "ItemType1       =   13" in rpt, rpt
    assert M.read_dsx(os.path.join(d, "DataReport1.Dsr")) == "DataEnvironment1"
    dsr = open(os.path.join(d, "DataEnvironment1.Dsr"), encoding="cp1252").read()
    assert 'ParentCommandName=   "Customers"' in dsr and 'AggField        =   "Amount"' in dsr, dsr
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "out.txt"), encoding="cp1252").read().split()
    assert out[:2] == ["Orders", "Alfreds"] and "120.5" in out and "Bon" in out and "7.25" in out, out
    # open again: the designers read what was written
    w.close_project()
    w.open_project(os.path.join(d, "Rpt.vbp"))
    m2 = w.project.module("DataReport1")
    r2 = w.open_designer(m2)
    assert w.design_windows[m2].widget().model.datasource == "DataEnvironment1"
    assert [s.name for s in w.design_windows[m2].widget().model.sections][1:4] == ["Customers_Header", "Orders_Detail", "Customers_Footer"]
    del r2
    done(w)


def t_dsr2(d):
    w = MainWindow()
    w.show()
    p = w.project
    w.add_module("Module")
    code(w, p.module("Module1"), MAKE_DB + "\nSub Main()\n    MakeDb\nEnd Sub\n")
    p.vbp.set("Startup", "Sub Main", quote=True)
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Module1.bas"), os.path.join(d, "Par.vbp"))
    assert w.save_project_as()
    w.run_start()
    wait(w)
    db = os.path.join(d, "ide.mdb")
    w.add_module("DataEnvironment")
    dem = p.module("DataEnvironment1")
    de = w.design_windows[dem].widget()
    de.cpage.show_conn(de.model.conns[0])
    de.cpage.cs.setText(f"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={db}")
    de.cpage.apply()
    de.add_command(False)
    cmd = de.model.cmds[0]
    pg = de.mpage
    pg.show_cmd(cmd)
    pg.name.setText("OrdersOf")
    pg.rb_sql.setChecked(True)
    pg.sql.setPlainText("SELECT OrderID, Amount FROM Orders WHERE CustID = ? AND Amount > ? ORDER BY OrderID")
    pg.apply()                                  # the parameters from the ? of the statement
    assert [x.name for x in cmd.params] == ["Param1", "Param2"], [x.name for x in cmd.params]
    assert [f[0] for f in cmd.fields] == ["OrderID", "Amount"], cmd.fields     # its fields without values
    pg.show_cmd(cmd)
    pg.params.item(0, 0).setText("Cust")
    pg.params.cellWidget(1, 2).setCurrentIndex(pg.params.cellWidget(1, 2).findData(6))
    pg._host_follows(1, 6)
    pg.apply()
    assert [(x.name, x.v["DataType"], x.v["HostType"]) for x in cmd.params] == [("Cust", 202, 8), ("Param2", 6, 6)], \
        [(x.name, x.v["DataType"], x.v["HostType"]) for x in cmd.params]
    # the report: fields dragged from the Data Environment
    w.add_module("DataReport")
    rm = p.module("DataReport1")
    rd = w.design_windows[rm].widget()
    it = de.tree.findItems("OrdersOf", Qt.MatchFlag.MatchRecursive)[0]
    md = de.tree.mimeData([it])
    import json
    from lbstudio.dsrdesign import FIELDS_MIME
    info = json.loads(bytes(md.data(FIELDS_MIME)).decode())
    assert info == {"de": "DataEnvironment1", "member": "OrdersOf", "fields": ["OrderID", "Amount"]}, info
    det = next(s for s in rd.model.sections if s.code == 4)
    rd.drop_fields(det, 300, 0, info["de"], info["member"], info["fields"])
    assert rd.model.datasource == "DataEnvironment1" and rd.model.datamember == "OrdersOf"
    got = [(i.kind, i.name, i.props.get("Caption"), i.props.get("DataField"), i.props.get("DataMember")) for i in det.items]
    assert got == [("RptLabel", "Label1", "OrderID:", None, None), ("RptTextBox", "txtOrderID", None, "OrderID", "OrdersOf"),
                   ("RptLabel", "Label2", "Amount:", None, None), ("RptTextBox", "txtAmount", None, "Amount", "OrdersOf")], got
    # the events of a control of a component (SysInfo) in the code window
    p.set_components(["SysInfoLib"])
    f = w.open_designer(p.module("Form1"))
    f.place_new("SysInfoLib.SysInfo", QRect(8, 8, 0, 0), None)
    cw = w.open_code(p.module("Form1"))
    cw._fill_procs("SysInfo1")
    evs = [cw.proc_box.itemText(k) for k in range(cw.proc_box.count())]
    assert "PowerStatusChanged" in evs and "DeviceArrival" in evs, evs
    code(w, p.module("Module1"), MAKE_DB + "\nSub Main()\n    Dim f As Integer\n    MakeDb\n"
         "    DataEnvironment1.OrdersOf \"ALFKI\", 50\n"
         "    f = FreeFile\n    Open App.Path & \"\\\\out.txt\" For Output As #f\n"
         "    Do Until DataEnvironment1.rsOrdersOf.EOF\n"
         "        Print #f, DataEnvironment1.rsOrdersOf!OrderID; DataEnvironment1.rsOrdersOf!Amount\n"
         "        DataEnvironment1.rsOrdersOf.MoveNext\n    Loop\n    Close #f\n"
         "    DataEnvironment1.rsOrdersOf.Close\n"
         "    DataEnvironment1.Commands(\"OrdersOf\").Parameters(0).Value = \"ALFKI\"\n"
         "    DataEnvironment1.Commands(\"OrdersOf\").Parameters(1).Value = 0\n"
         "    DataReport1.ExportReport \"key_def_Text\", App.Path & \"\\\\rpt.txt\", True, False\n"
         "    Unload DataReport1\nEnd Sub\n")
    assert w.save_project()
    dsr = open(os.path.join(d, "DataEnvironment1.Dsr"), encoding="cp1252").read()
    assert "ParamCount      =   2" in dsr and 'UserName        =   "Cust"' in dsr, dsr
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "out.txt")).read().split()
    assert out == ["10", "100.5"], out
    rpt = open(os.path.join(d, "rpt.txt"), encoding="cp1252").read()
    assert "OrderID:" in rpt and "Amount:" in rpt, rpt
    done(w)


def _caret(w, m, text):
    """the caret on the first place of text in the code of m"""
    cw = w.open_code(m)
    ed = cw.editor
    i = cw.text().index(text)
    ed.SendScintilla(2025, len(cw.text()[:i].encode("utf-8")))      # SCI_GOTOPOS
    return cw


def t_refactor(d):
    from lbstudio import refactorui as RF
    w = MainWindow()
    w.show()
    p = w.project
    f = w.open_designer(p.module("Form1"))
    f.place_new("CommandButton", QRect(8, 8, 100, 30), None)
    w.add_module("Module")
    w.add_module("Class")
    code(w, p.module("Class1"), "Option Explicit\nPublic Event Changed(ByVal v As Long)\nPrivate mValue As Long\n\n"
         "Public Property Get Value() As Long\n    Value = mValue\nEnd Property\n\n"
         "Public Property Let Value(ByVal v As Long)\n    mValue = v\n    RaiseEvent Changed(v)\nEnd Property\n")
    code(w, p.module("Module1"), "Option Explicit\nPublic Total As Long\n\nPublic Sub AddUp(ByVal n As Long)\n"
         "    Dim i As Long\n    For i = 1 To n\n        Total = Total + i\n    Next i\nEnd Sub\n")
    code(w, p.module("Form1"), "Option Explicit\nPrivate WithEvents mCounter As Class1\n\n"
         "Private Sub Form_Load()\n    Dim f As Integer, x As Long, y As Long\n    Set mCounter = New Class1\n"
         "    mCounter.Value = 5\n    AddUp 4\n"
         "    x = Total * 2\n    y = x + mCounter.Value\n    Command1.Caption = CStr(y)\n"
         "    f = FreeFile\n    Open App.Path & \"/out.txt\" For Output As #f\n    Print #f, Command1.Caption\n"
         "    Close #f\n    Unload Me\nEnd Sub\n\n"
         "Private Sub Command1_Click()\n    Command1.Caption = \"x\"\nEnd Sub\n\n"
         "Private Sub mCounter_Changed(ByVal v As Long)\n    Debug.Print v\nEnd Sub\n")
    # Find References: Total
    _caret(w, p.module("Form1"), "Total * 2")
    RF.find_references(w)
    rows = [(w.references.topLevelItem(i).text(0), w.references.topLevelItem(i).text(1))
            for i in range(w.references.topLevelItemCount())]
    assert rows == [("Form1", "9"), ("Module1", "2"), ("Module1", "7"), ("Module1", "7")], rows
    # Definition: the declaration of AddUp
    _caret(w, p.module("Form1"), "AddUp 4")
    RF.definition(w)
    cw = w.current_cw()
    assert cw.module.name == "Module1" and cw.editor.sel_text() == "AddUp", (cw.module.name, cw.editor.sel_text())
    RF.last_position(w)
    assert w.current_cw().module.name == "Form1"
    # Rename: a procedure, a variable, a control (its events), a property, a class (the WithEvents variable stays)
    _caret(w, p.module("Form1"), "AddUp 4")
    assert RF.rename(w, "AddAll")
    _caret(w, p.module("Module1"), "Total As Long")
    assert RF.rename(w, "GrandTotal")
    _caret(w, p.module("Form1"), "Command1.Caption = CStr")
    assert RF.rename(w, "cmdGo")
    _caret(w, p.module("Class1"), "Value() As Long")
    assert RF.rename(w, "Amount")
    _caret(w, p.module("Form1"), "mCounter.Amount = 5")
    assert RF.rename(w, "mCnt")
    _caret(w, p.module("Form1"), "As Class1")
    w.current_cw().editor.SendScintilla(2025, w.current_cw().editor.pos() + 3)
    assert RF.rename(w, "Counter")
    assert not RF.rename(w, "Module1")          # a name that is taken
    frm = w.current_cw().text() if w.current_cw().module.name == "Form1" else p.module("Form1").code_text()
    frm = p.module("Form1").code_text()
    for want in ("Private WithEvents mCnt As Counter", "AddAll 4", "x = GrandTotal * 2", "y = x + mCnt.Amount",
                 "cmdGo.Caption = CStr(y)", "Private Sub cmdGo_Click()", "Private Sub mCnt_Changed(", "New Counter"):
        assert want in frm, (want, frm)
    assert "Public Sub AddAll(" in p.module("Module1").code_text() and "GrandTotal = GrandTotal + i" in p.module("Module1").code_text()
    assert p.module("Counter") is not None and "Public Property Let Amount(" in p.module("Counter").code_text()
    assert [c.name for c in p.module("Form1").sf.root.walk()][1] == "cmdGo"
    # Extract Procedure: the two lines that compute x and y
    cw = w.open_code(p.module("Form1"))
    ed = cw.editor
    t = cw.text()
    a = len(t[:t.index("    x = GrandTotal")].encode())
    b = len(t[:t.index("    cmdGo.Caption = CStr")].encode())
    ed.set_sel(a, b)
    assert RF.extract(w, "Compute")
    frm = p.module("Form1").code_text()
    assert "    Compute x, y\n" in frm.replace("\r\n", "\n") and "Private Sub Compute(x As Long, y As Long)" in frm, frm
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Module1.bas"), os.path.join(d, "Counter.cls"),
          os.path.join(d, "Ref.vbp"))
    assert w.save_project_as()
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "out.txt")).read().split()
    assert out == ["25"], out                   # 2 * (1+2+3+4) + 5
    done(w)


def t_dbgrid(d):
    import re
    from lbvc.project import oleblob as OB
    w = MainWindow()
    w.show()
    p = w.project
    f = w.current_designer()
    d1 = f.place_new("Data", QRect(8, 160, 160, 24), None)
    d2 = f.place_new("Data", QRect(176, 160, 160, 24), None)
    g = f.place_new("MSDBGrid.DBGrid", QRect(8, 8, 240, 120), None)
    lst = f.place_new("MSDBCtls.DBList", QRect(256, 8, 100, 80), None)
    assert g is not None and lst is not None and d1.name == "Data1" and d2.name == "Data2", (g, lst)
    w.props.show_object(f.module, f, g)
    names = [r.name for r in w.props.rows]
    assert "(Columns)" in names and "MarqueeStyle" in names and "DataSource" in names, names
    for name, v in (("Caption", "People"), ("AllowAddNew", True), ("MarqueeStyle", 3), ("DataSource", "Data1")):
        w.props.commit(V.prop_def("DBGrid", name), v)
    cols = [{"Caption": "Number", "DataField": "ID", "Width": 1764,
             "Alignment": 2}, {"Caption": "Name", "DataField": "Name", "Width": 3528, "Alignment": 3}]
    w.props.commit(V.prop_def("DBGrid", "(Columns)"), (cols, 4))
    assert w.props.value(V.prop_def("DBGrid", "Caption")) == "People"
    assert w.props.value(V.prop_def("DBGrid", "DataSource")) == "Data1"
    assert V.display(V.prop_def("DBGrid", "(Columns)"), w.props.value(V.prop_def("DBGrid", "(Columns)"))) == "(2 columns)"
    w.props.show_object(f.module, f, lst)
    for name, v in (("DataSource", "Data1"), ("DataField", "GroupID"), ("RowSource", "Data2"), ("ListField", "GName"),
                    ("BoundColumn", "GID")):
        w.props.commit(V.prop_def("DBList", name), v)
    assert w.props.value(V.prop_def("DBList", "RowSource")) == "Data2"
    code(w, p.module("Form1"),
         "Private Sub Form_Load()\n    Dim e As Object, db As Object, f As Integer, p As String\n"
         "    p = App.Path & \"/g.mdb\"\n    If Dir(p) <> \"\" Then Kill p\n"
         "    Set e = CreateObject(\"DAO.DBEngine.36\")\n"
         "    Set db = e.CreateDatabase(p, \";LANGID=0x0409;CP=1252;COUNTRY=0\", 32)\n"
         "    db.Execute \"CREATE TABLE People (ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, Name TEXT(20), GroupID LONG)\"\n"
         "    db.Execute \"CREATE TABLE Groups (GID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, GName TEXT(20))\"\n"
         "    db.Execute \"INSERT INTO Groups (GName) VALUES ('Alpha')\"\n"
         "    db.Execute \"INSERT INTO People (Name, GroupID) VALUES ('Anna', 1)\"\n"
         "    db.Execute \"INSERT INTO People (Name, GroupID) VALUES ('Bill', 1)\"\n    db.Close\n"
         "    Data1.DatabaseName = p\n    Data1.RecordSource = \"People\"\n    Data1.Refresh\n"
         "    Data2.DatabaseName = p\n    Data2.RecordSource = \"Groups\"\n    Data2.Refresh\n"
         "    f = FreeFile\n    Open App.Path & \"/out.txt\" For Output As #f\n"
         "    Print #f, DBGrid1.Caption & \"|\" & DBGrid1.AllowAddNew & \"|\" & DBGrid1.MarqueeStyle & \"|\" & "
         "DBGrid1.Columns.Count & \"|\" & DBGrid1.Columns(0).Caption & \"|\" & DBGrid1.Columns(0).DataField & \"|\" & "
         "CLng(DBGrid1.Columns(0).Width) & \"|\" & DBGrid1.Columns(0).Alignment & \"|\" & DBGrid1.ApproxCount & \"|\" & "
         "DBGrid1.Columns(1).Text & \"|\" & DBList1.Text\n"
         "    Close #f\n    Unload Me\nEnd Sub\n")
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Grid.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(d, "Form1.frm"), encoding="cp1252").read()
    assert 'Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"' in frm, frm
    assert 'Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"' in frm, frm
    grid = frm[frm.index("Begin MSDBGrid.DBGrid DBGrid1"):]
    grid = grid[:grid.index("   End")]
    assert "OleObjectBlob" in grid and "Bindings" in grid and "Caption" not in grid and "DataSource" not in grid, grid
    lblock = frm[frm.index("Begin MSDBCtls.DBList DBList1"):]
    lblock = lblock[:lblock.index("   End")]
    assert 'DataSource      =   "Data1"' in lblock and "RowSource" not in lblock and "Bindings" in lblock, lblock
    frx = open(os.path.join(d, "Form1.frx"), "rb").read()
    offs = sorted(int(x, 16) for x in re.findall(r'"Form1\.frx":([0-9A-Fa-f]+)', frm))
    def rec(prop, block):
        o = int(re.search(prop + r'\s+=\s+"Form1\.frx":([0-9A-Fa-f]+)', block).group(1), 16)
        nxt = [x for x in offs if x > o]
        return frx[o:nxt[0] if nxt else len(frx)]
    props = dict(OB.decode(rec("OleObjectBlob", grid)))
    assert props["Caption"] == "People" and props["AllowAddNew"] == -1 and props["Split0.MarqueeStyle"] == 3 and \
        props["_LayoutType"] == 4 and props["Columns.Count"] == 2 and props["Columns(0).DataField"] == "ID" and \
        props["Columns(0).Alignment"] == 2, props
    assert OB.bindings(rec("Bindings", grid)) == [(9, "Data1")]
    assert OB.bindings(rec("Bindings", lblock)) == [(3, "Data2")]
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "out.txt")).read().strip()
    assert out == "People|True|3|2|Number|ID|1000|2|2|Anna|Alpha", out
    done(w)


def t_dfmt(d):
    from lbstudio.dfmtdialog import DataFormatDialog, CATS
    w = MainWindow()
    w.show()
    p = w.project
    f = w.current_designer()
    t1 = f.place_new("TextBox", QRect(8, 8, 120, 24), None)
    t2 = f.place_new("TextBox", QRect(8, 40, 120, 24), None)
    assert t1 is not None and t2 is not None
    pdf = V.prop_def("TextBox", "DataFormat")
    assert pdf is not None and pdf.kind == "dataformat", pdf
    dlg = DataFormatDialog(w, "Text1", None)        # Number: 2 decimal places, the 1000 separator
    dlg.cats.setCurrentRow(CATS.index("Number"))
    dlg.w["Number"]["dec"].setValue(2)
    dlg.w["Number"]["thou"].setChecked(True)
    v1 = dlg.chosen()
    assert v1["Type"] == 1 and v1["Format"] == "#,##0.00" and v1["SubFormatType"] == 1 and v1["LCID"] == 1033, v1
    w.props.show_object(f.module, f, t1)
    w.props.commit(pdf, v1)
    assert V.display(pdf, w.props.value(pdf)) == "Number (#,##0.00)", V.display(pdf, w.props.value(pdf))
    dlg = DataFormatDialog(w, "Text2", None)        # Boolean: Yes / No / ?
    dlg.cats.setCurrentRow(CATS.index("Boolean"))
    dlg.w["Boolean"]["true"].setText("Yes")
    dlg.w["Boolean"]["false"].setText("No")
    dlg.w["Boolean"]["null"].setText("?")
    v2 = dlg.chosen()
    assert v2["Type"] == 5 and v2["HaveTrueFalseNull"] == 1 and v2["SubFormatType"] == 7, v2
    w.props.show_object(f.module, f, t2)
    w.props.commit(pdf, v2)
    dlg = DataFormatDialog(w, "Text2", w.props.value(pdf))   # the page opened again: what the .frm has
    assert dlg._cat() == "Boolean" and dlg.w["Boolean"]["true"].text() == "Yes" and dlg.w["Boolean"]["null"].text() == "?"
    code(w, p.module("Form1"),
         "Private Sub Form_Load()\n    Dim f As Integer\n    f = FreeFile\n"
         "    Open App.Path & \"/out.txt\" For Output As #f\n"
         "    Print #f, TypeName(Text1.DataFormat) & \"|\" & Text1.DataFormat.Type & \"|\" & Text1.DataFormat.Format & \"|\" & "
         "Text2.DataFormat.Type & \"|\" & Text2.DataFormat.TrueValue & \"|\" & Text2.DataFormat.FalseValue & \"|\" & "
         "Text2.DataFormat.NullValue\n"
         "    Close #f\n    Unload Me\nEnd Sub\n")
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Dfmt.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(d, "Form1.frm"), encoding="cp1252").read()
    b1 = frm[frm.index("Begin VB.TextBox Text1"):]
    b1 = b1[:b1.index("   End")]
    assert "BeginProperty DataFormat \n" in b1 and 'Format          =   "#,##0.00"' in b1 and \
        "Type            =   1" in b1 and "SubFormatType   =   1" in b1 and "LCID            =   1033" in b1 and \
        "TrueValue" not in b1, b1
    b2 = frm[frm.index("Begin VB.TextBox Text2"):]
    b2 = b2[:b2.index("   End")]
    assert "HaveTrueFalseNull=   1" in b2 and 'TrueValue       =   "Yes"' in b2 and 'NullValue       =   "?"' in b2, b2
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "out.txt")).read().strip()
    assert out == "IStdDataFormatDisp|1|#,##0.00|5|Yes|No|?", out
    done(w)


def t_dgcols(d):
    import re
    from lbvc.project import oleblob as OB
    from lbstudio.dgcolumns import GridColumnsDialog
    w = MainWindow()
    w.show()
    p = w.project
    f = w.current_designer()
    a = f.place_new("MSAdodcLib.Adodc", QRect(8, 200, 160, 24), None)
    g = f.place_new("MSDataGridLib.DataGrid", QRect(8, 8, 240, 120), None)
    lst = f.place_new("MSDataListLib.DataList", QRect(256, 8, 100, 80), None)
    assert a is not None and g is not None and lst is not None and a.name == "Adodc1", (a, g, lst)
    w.props.show_object(f.module, f, g)
    pc = V.prop_def("DataGrid", "(Columns)")
    assert pc is not None and "(Columns)" in [r.name for r in w.props.rows]
    w.props.commit(V.prop_def("DataGrid", "DataSource"), "Adodc1")
    dlg = GridColumnsDialog(w, "DataGrid1", [], ["ID", "Name", "Age"])
    dlg._fill([{"Caption": f, "DataField": f} for f in dlg.fields])    # Retrieve Fields
    dlg.table.item(2, 0).setText("Years")
    dlg.table.item(2, 2).setText("1200")
    dlg.table.cellWidget(2, 3).setCurrentIndex(2)
    dlg.table.cellWidget(0, 5).setChecked(False)
    dlg.set_format(2, {"Type": 1, "Format": "000", "HaveTrueFalseNull": 0, "LCID": 1033, "SubFormatType": 1})
    cols = dlg.chosen()
    assert len(cols) == 3 and cols[2]["Caption"] == "Years" and cols[2]["Width"] == 1200 and cols[2]["Alignment"] == 2 \
        and not cols[0]["Visible"] and cols[2]["DataFormat"]["Format"] == "000", cols
    w.props.commit(pc, cols)
    assert V.display(pc, w.props.value(pc)) == "(3 columns)"
    back = w.props.value(pc)
    assert back[2]["DataFormat"]["Format"] == "000" and back[2]["Width"] == 1200 and not back[0]["Visible"], back
    w.props.show_object(f.module, f, lst)
    w.props.commit(V.prop_def("DataList", "DataFormat"), {"Type": 1, "Format": "0.0", "HaveTrueFalseNull": 0,
                                                          "LCID": 1033, "SubFormatType": 1})
    code(w, p.module("Form1"),
         "Private Sub Form_Load()\n    Dim f As Integer, c As Object\n    f = FreeFile\n"
         "    Open App.Path & \"/out.txt\" For Output As #f\n"
         "    Set c = DataGrid1.Columns(2)\n"
         "    Print #f, DataGrid1.Columns.Count & \"|\" & DataGrid1.Columns(0).Visible & \"|\" & c.Caption & \"|\" & "
         "c.DataField & \"|\" & c.DataFormat.Format & \"|\" & c.NumberFormat & \"|\" & c.Alignment & \"|\" & "
         "CLng(c.Width) & \"|\" & DataList1.DataFormat.Format\n"
         "    Close #f\n    Unload Me\nEnd Sub\n")
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Cols.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(d, "Form1.frm"), encoding="cp1252").read()
    grid = frm[frm.index("Begin MSDataGridLib.DataGrid DataGrid1"):]
    grid = grid[:grid.index("\n   End")]
    assert "ColumnCount     =   3" in grid and 'Caption         =   "Years"' in grid and \
        "BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED}" in grid and 'Format          =   "000"' in grid and \
        "SplitCount      =   1" in grid and "ColumnWidth     =   1200" in grid and "Alignment       =   2" in grid and \
        "Object.Visible  =   0" in grid and "Bindings" in grid and "DataSource" not in grid, grid
    lblock = frm[frm.index("Begin MSDataListLib.DataList DataList1"):]
    lblock = lblock[:lblock.index("\n   End")]
    assert "BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED}" in lblock and 'Format          =   "0.0"' in lblock, lblock
    frx = open(os.path.join(d, "Form1.frx"), "rb").read()
    o = int(re.search(r'Bindings\s+=\s+"Form1\.frx":([0-9A-Fa-f]+)', grid).group(1), 16)
    offs = sorted(int(x, 16) for x in re.findall(r'"Form1\.frx":([0-9A-Fa-f]+)', frm))
    nxt = [x for x in offs if x > o]
    assert OB.bindings(frx[o:nxt[0] if nxt else len(frx)]) == [(42, "Adodc1")]
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "out.txt")).read().strip()
    assert out == "3|False|Years|Age|000|000|2|1200|0.0", out
    done(w)


def t_coolbar(d):
    import re
    from lbstudio.coolbands import CoolBandsDialog
    w = MainWindow()
    w.show()
    p = w.project
    f = w.current_designer()
    p.set_components(["ComCtl3"])                       # Project > Components: the Object= line of the .vbp
    cb = f.place_new("ComCtl3.CoolBar", QRect(0, 0, 200, 28), None)
    assert cb is not None and cb.name == "CoolBar1", cb
    t = f.place_new("TextBox", QRect(20, 4, 80, 19), f.items.get(cb))
    assert t is not None and t.parent is cb, (t, t and t.parent)
    w.props.show_object(f.module, f, cb)
    pb = V.prop_def("CoolBar", "(Bands)")
    assert pb is not None and "(Bands)" in [r.name for r in w.props.rows]
    bands = w.props.value(pb)
    assert len(bands) == 3 and bands[0]["MinHeight"] == 360 and not bands[0]["NewRow"], bands
    dlg = CoolBandsDialog(w, "CoolBar1", bands, ["Text1"])
    dlg.table.item(0, 0).setText("One")                 # Caption
    dlg.table.cellWidget(0, 2).setCurrentText("Text1")  # Child
    dlg.table.item(0, 3).setText("2000")                # Width: whole pixels -> 1995
    dlg.table.item(0, 4).setText("1500")                # MinWidth
    dlg.table.selectRow(2)
    dlg.delete()
    dlg.table.cellWidget(1, 6).setChecked(True)         # NewRow
    dlg.table.item(1, 1).setText("two")                 # Key
    dlg.table.cellWidget(1, 8).setCurrentIndex(1)       # cc3BandFixedSize
    got = dlg.chosen()
    assert len(got) == 2 and got[0]["Width"] == 1995 and got[0]["Child"] == "Text1" and got[1]["NewRow"] and \
        got[1]["Style"] == 1 and got[1]["Key"] == "two", got
    w.props.commit(pb, got)
    assert V.display(pb, w.props.value(pb)) == "(2 bands)"
    cw = int(f.root.get("ClientWidth", 4800))
    # the band: no gripper (the other band has a fixed size), 4, "One" 20 + 6, the child, 4; the row: Width - 4
    tw = (cw // 15 - 4 - 30 - 4) * 15
    geo = [int(t.get(k)) for k in ("Left", "Top", "Width", "Height")]
    assert geo == [450, 0, tw, 360] and int(cb.get("Height")) == 810, (geo, cb.get("Height"))
    code(w, p.module("Form1"),
         "Private Sub Form_Activate()\n    Dim f As Integer\n    f = FreeFile\n"
         "    Open App.Path & \"/out.txt\" For Output As #f\n"
         "    Print #f, CoolBar1.Bands.Count & \"|\" & CoolBar1.Bands(1).Child.Name & \"|\" & Text1.Left & \"|\" & "
         "Text1.Width & \"|\" & CoolBar1.Height & \"|\" & CoolBar1.RowCount & \"|\" & CoolBar1.Bands(\"two\").Style & "
         "\"|\" & CoolBar1.Bands(1).Width & \"|\" & (cc3BandFixedSize = CoolBar1.Bands(2).Style)\n"
         "    Close #f\n    Unload Me\nEnd Sub\n")
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Cool.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(d, "Form1.frm"), encoding="cp1252").read()
    assert 'Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"' in frm, frm[:300]
    blk = frm[frm.index("Begin ComCtl3.CoolBar CoolBar1"):]
    blk = blk[:blk.index("Begin VB.TextBox")]
    for pat in (r"BandCount\s+=\s+2\s", r'Caption1\s+=\s+"One"', r'Child1\s+=\s+"Text1"', r"MinWidth1\s+=\s+1500\s",
                r"Width1\s+=\s+1995\s", r"NewRow1\s+=\s+0\s", r"NewRow2\s+=\s+-1\s", r'Key2\s+=\s+"two"',
                r"BandStyle2\s+=\s+1\s", r"Align\s+=\s+1\s", r'_Version\s+=\s+"6\.7\.9839"'):
        assert re.search(pat, blk), (pat, blk)
    assert "MinHeight3" not in blk and "Caption2" not in blk, blk
    vbp = open(os.path.join(d, "Cool.vbp"), encoding="cp1252").read()
    assert "Object={38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0; COMCT332.OCX" in vbp, vbp
    w.run_start()
    wait(w)
    out = open(os.path.join(d, "out.txt")).read().strip()
    assert out == f"2|Text1|450|{tw}|810|2|1|1995|True", out
    done(w)


def t_rdc(d):
    import re
    import subprocess
    w = MainWindow()
    w.show()
    p = w.project
    f = w.current_designer()
    p.set_components(["MSRDC"])
    r = f.place_new("MSRDC.MSRDC", QRect(8, 120, 160, 22), None)
    assert r is not None and r.name == "MSRDC1", r
    t = f.place_new("TextBox", QRect(8, 8, 120, 21), None)
    w.props.show_object(f.module, f, r)
    w.props.commit(V.prop_def("MSRDC", "ResultsetType"), 3)
    w.props.commit(V.prop_def("MSRDC", "LockType"), 4)
    w.props.commit(V.prop_def("MSRDC", "Prompt"), 1)
    w.props.show_object(f.module, f, t)
    w.props.commit(V.prop_def("TextBox", "DataSource"), "MSRDC1")
    w.props.commit(V.prop_def("TextBox", "DataField"), "Name")
    code(w, p.module("Form1"),
         "Private Sub Form_Load()\n"
         "    MSRDC1.Connect = \"Driver={SQLite3 ODBC Driver};Database=\" & App.Path & \"/t.db\"\n"
         "    MSRDC1.SQL = \"SELECT 1 AS X\"\n    MSRDC1.Refresh\n"
         "    MSRDC1.Connection.Execute \"CREATE TABLE P (Name VARCHAR(20))\"\n"
         "    MSRDC1.Connection.Execute \"INSERT INTO P VALUES ('Zoe')\"\n"
         "    MSRDC1.SQL = \"SELECT * FROM P\"\n    MSRDC1.Refresh\nEnd Sub\n"
         "Private Sub Form_Activate()\n    Dim f As Integer\n    f = FreeFile\n"
         "    Open App.Path & \"/out.txt\" For Output As #f\n"
         "    Print #f, Text1.Text & \"|\" & MSRDC1.ResultsetType & \"|\" & MSRDC1.Resultset.Type & \"|\" & "
         "MSRDC1.LockType & \"|\" & MSRDC1.Caption & \"|\" & (rdOpenStatic = MSRDC1.ResultsetType)\n"
         "    Close #f\n    Unload Me\nEnd Sub\n")
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Rdc.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(d, "Form1.frm"), encoding="cp1252").read()
    assert 'Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"' in frm, frm[:300]
    blk = frm[frm.index("Begin MSRDC.MSRDC MSRDC1"):]
    blk = blk[:blk.index("\n   End")]
    names = re.findall(r"^\s+(\w+)\s+=", blk, re.M)
    order = ["Options", "CursorDriver", "BOFAction", "EOFAction", "RecordsetType", "LockType", "QueryType", "Prompt",
             "Appearance", "QueryTimeout", "RowsetSize", "LoginTimeout", "KeysetSize", "MaxRows", "ErrorThreshold",
             "BatchSize", "BackColor", "ForeColor", "Enabled", "ReadOnly", "Appearance", "DataSourceName",
             "RecordSource", "UserName", "Password", "Connect", "LogMessages", "Caption"]
    assert names[names.index("Options"):names.index("Caption") + 1] == order, names
    for pat in (r"RecordsetType\s+=\s+3\s", r"LockType\s+=\s+4\s", r"Prompt\s+=\s+1\s", r'Caption\s+=\s+"MSRDC1"'):
        assert re.search(pat, blk), (pat, blk)
    assert re.search(r'DataSource\s+=\s+"MSRDC1"', frm) and re.search(r'DataField\s+=\s+"Name"', frm), frm
    vbp = open(os.path.join(d, "Rdc.vbp"), encoding="cp1252").read()
    assert "Object={F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0; MSRDC20.OCX" in vbp, vbp
    try:
        drivers = subprocess.run(["odbcinst", "-q", "-d"], capture_output=True, text=True, timeout=20).stdout.lower()
    except OSError:
        drivers = ""
    if "sqlite3" in drivers:                            # F5: the program over the SQLite ODBC driver
        w.run_start()
        wait(w)
        out = open(os.path.join(d, "out.txt")).read().strip()
        assert out == "Zoe|3|3|4|MSRDC1|True", out
    done(w)


# the DataRepeater the VB6 IDE wrote (VB6 oracle f_drep3, VB-TESTS/DEV-7/f_drep3/Form1.frm)
DREP_FRM = '''   Begin MSDataRepeaterLib.DataRepeater DataRepeater1
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3889
      _StreamID       =   -1412567295
      _Version        =   393216
      RowIndicator    =   0   'False
      BackColor       =   255
      ScrollBars      =   3
      RowDividerStyle =   1
      CaptionStyle    =   0
      Caption         =   "Edited"
      IntegralHeight  =   0   'False
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10}
         _StreamID       =   -1412567295
         _Version        =   65536
         Name            =   "RepCtl.RowCtl"
      EndProperty
      RepeaterBindings=   3
      BeginProperty RepeaterBinding(0) {7D21A594-FC9B-11D0-A320-00AA00688B10}
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "PName"
         DataField       =   "Name"
         Key             =   "kName"
      EndProperty
      BeginProperty RepeaterBinding(1) {7D21A594-FC9B-11D0-A320-00AA00688B10}
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "PAge"
         DataField       =   "Age"
         Key             =   "kAge"
      EndProperty
      BeginProperty RepeaterBinding(2) {7D21A594-FC9B-11D0-A320-00AA00688B10}
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "PActive"
         DataField       =   "Active"
      EndProperty
   End'''


def t_drep(d):
    w = MainWindow()
    w.show()
    p = w.project
    f = w.current_designer()
    p.set_components(["MSDataRepeaterLib"])
    r = f.place_new("MSDataRepeaterLib.DataRepeater", QRect(8, 8, 305, 147), None)
    assert r is not None and r.name == "DataRepeater1", r
    w.props.show_object(f.module, f, r)
    for name, v in (("Width", 4575), ("Height", 2205), ("Caption", "Edited"), ("CaptionStyle", 0), ("RowDividerStyle", 1), ("ScrollBars", 3),
                    ("RowIndicator", False), ("IntegralHeight", False), ("BackColor", 255),
                    ("RepeatedControlName", "RepCtl.RowCtl")):
        w.props.commit(V.prop_def("DataRepeater", name), v)
    w.props.commit(V.prop_def("DataRepeater", "(RepeaterBindings)"), [
        {"PropertyName": "PName", "DataField": "Name", "Key": "kName"},
        {"PropertyName": "PAge", "DataField": "Age", "Key": "kAge"},
        {"PropertyName": "PActive", "DataField": "Active", "Key": ""}])
    assert V.drep_name(r) == "RepCtl.RowCtl" and len(V.drep_bindings(r)) == 3, (V.drep_name(r), V.drep_bindings(r))
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Drep.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(d, "Form1.frm"), encoding="cp1252").read()
    assert 'Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDatRep.Ocx"' in frm, frm[:300]
    blk = frm[frm.index("   Begin MSDataRepeaterLib.DataRepeater DataRepeater1"):]
    blk = blk[:blk.index("\n   End") + len("\n   End")]
    got = [x.rstrip() for x in blk.splitlines()]
    exp = DREP_FRM.splitlines()
    assert got == exp, "\n".join(f"{a!r} | {b!r}" for a, b in zip(got + [''] * len(exp), exp + [''] * len(got)) if a != b)
    vbp = open(os.path.join(d, "Drep.vbp"), encoding="cp1252").read()
    assert "Object={5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0; MSDatRep.Ocx" in vbp, vbp
    # the bindings again, one less: the blocks are written anew (the control of the designer after the save)
    f = w.current_designer()
    r = next(c for c in f.root.walk() if c.name == "DataRepeater1")
    w.props.show_object(f.module, f, r)
    w.props.commit(V.prop_def("DataRepeater", "(RepeaterBindings)"), V.drep_bindings(r)[:1])
    assert [b["PropertyName"] for b in V.drep_bindings(r)] == ["PName"] and V.drep_name(r) == "RepCtl.RowCtl", \
        (V.drep_bindings(r), V.drep_name(r))
    done(w)


MAPI_FRM = '''   Begin MSMAPI.MAPISession MAPISession1
      Left            =   240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   0   'False
      NewSession      =   -1  'True
      Password        =   "secret"
      UserName        =   "tester"
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1
      Left            =   1200
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   3
      AddressModifiable=   -1  'True
      AddressResolveUI=   -1  'True
      FetchSorted     =   -1  'True
      FetchUnreadOnly =   -1  'True
      AddressCaption  =   "Pick"
      AddressLabel    =   "To:"
      FetchMsgType    =   "IPM.Note"
   End'''


def t_mapi(d):
    """MAPISession and MAPIMessages as the VB6 IDE writes them (VB6 oracle f_mapi5: Form1.frm)"""
    w = MainWindow()
    w.show()
    p = w.project
    f = w.current_designer()
    p.set_components(["MSMAPI"])
    s = f.place_new("MSMAPI.MAPISession", QRect(16, 16, 38, 38), None)
    m = f.place_new("MSMAPI.MAPIMessages", QRect(80, 16, 38, 38), None)
    assert s is not None and s.name == "MAPISession1" and m is not None and m.name == "MAPIMessages1", (s, m)
    w.props.show_object(f.module, f, s)
    for name, v in (("DownloadMail", False), ("LogonUI", False), ("NewSession", True), ("Password", "secret"),
                    ("UserName", "tester")):
        w.props.commit(V.prop_def("MAPISession", name), v)
    w.props.show_object(f.module, f, m)
    for name, v in (("AddressCaption", "Pick"), ("AddressEditFieldCount", 3), ("AddressLabel", "To:"),
                    ("AddressModifiable", True), ("AddressResolveUI", True), ("FetchMsgType", "IPM.Note"),
                    ("FetchSorted", True), ("FetchUnreadOnly", True)):
        w.props.commit(V.prop_def("MAPIMessages", name), v)
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Mapi.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(d, "Form1.frm"), encoding="cp1252").read()
    assert 'Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.Ocx"' in frm, frm[:300]
    got = []
    for head in ("   Begin MSMAPI.MAPISession MAPISession1", "   Begin MSMAPI.MAPIMessages MAPIMessages1"):
        assert head in frm, frm
        blk = frm[frm.index(head):]
        got += [x.rstrip() for x in blk[:blk.index("\n   End") + len("\n   End")].splitlines()]
    exp = MAPI_FRM.splitlines()
    assert got == exp, "\n".join(f"{a!r} | {b!r}" for a, b in zip(got + [''] * len(exp), exp + [''] * len(got)) if a != b)
    vbp = open(os.path.join(d, "Mapi.vbp"), encoding="cp1252").read()
    assert "Object={20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0; msmapi32.Ocx" in vbp, vbp
    done(w)


WEB_BAG = '''      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End'''


def t_appwiz92(d):
    """the Application Wizard in the IDE (phase 9.2a-3, VB6 oracle DEV-9/aw1): New Project has "VB Application
    Wizard"; the pages (SDI: Window not checked, the 17 buttons of the toolbar), Standard Forms, Form Templates,
    Resources (the String Table in the Resource Editor, not saved), Internet; Finish: the new project in the IDE,
    the profile .rwp; the designer draws the buttons of the toolbar with the pictures of the ImageList; Save, lbvc.
    Add-Ins: the "VB 6 Application Wizard" loaded gives Application Wizard... and Toolbar Wizard... (aw18). The
    Toolbar Wizard (9.2a-4, aw20 / aw21): Toolbar1 + imlToolbarIcons + Toolbar1_ButtonClick; without buttons only
    Toolbar1"""
    import subprocess
    from lbstudio import appwizgen as G
    from lbstudio.appwizard import AppWizard
    from lbstudio.tbwizard import ToolbarWizard
    from lbstudio.dialogs import NewProjectDialog
    from lbstudio.resedit import AddInManagerDialog, APPWIZ
    from lbstudio.designer import list_images
    w = MainWindow()
    w.show()
    w.S["appwiz_no_created"] = True             # no "Application Created" in the test
    npd = NewProjectDialog(w)
    assert [npd.list.item(i).text() for i in range(npd.list.count())][-1] == "VB Application Wizard"
    dlg = AppWizard(w)
    assert dlg.windowTitle() == "Application Wizard - Introduction" and not dlg.b_finish.isEnabled()
    dlg.go(1)
    dlg.type_group.button(G.APP_SDI).click()
    assert dlg.hint.text() == "All windows exist at the same level."
    dlg.name.setText("Wiz1")
    dlg.go(2)
    caps = [(dlg.menu_list.item(i).text(), dlg.menu_list.item(i).checkState() == Qt.CheckState.Checked)
            for i in range(dlg.menu_list.count())]
    assert caps == [("&File", True), ("&Edit", True), ("&View", True), ("&Tools", False), ("&Window", False),
                    ("&Help", True)], caps
    dlg.menu_list.setCurrentRow(0)
    assert dlg.sub_list.count() == 18 and dlg.sub_list.item(17).text() == "E&xit" and \
        dlg.sub_list.item(16).text() == "MRU List", [dlg.sub_list.item(i).text() for i in range(dlg.sub_list.count())]
    dlg.go(3)
    assert dlg.tbed.chosen.count() == 17 and dlg.tbed.avail.item(0).text() == "[Separator]"
    assert len(dlg.tbed.preview.actions()) == 17
    dlg.go(4)
    dlg.res_group.button(1).setChecked(True)
    dlg.go(5)
    dlg.net_group.button(1).setChecked(True)
    dlg.go(6)
    dlg.form_checks["About"].setChecked(True)
    dlg.opts.templates = ["Tip of the Day"]
    dlg.go(8)
    assert not dlg.b_next.isEnabled() and dlg.b_finish.isEnabled()
    prof = os.path.join(d, "wiz.rwp")
    dlg.save_combo.addItem(prof)
    dlg.save_combo.setCurrentIndex(dlg.save_combo.count() - 1)
    dlg.finish()
    p = w.project
    assert [m.name for m in p.modules] == ["Module1", "frmMain", "frmBrowser", "frmAbout", "frmTip"], \
        [m.name for m in p.modules]
    assert p.name == "Wiz1" and (p.vbp.get("Startup") or "").strip('"') == "Sub Main"
    rwp = open(prof, encoding="cp1252").read()
    assert "ApplicationType=1" in rwp and "Resources=Yes" in rwp and "OtherForms3=1" in rwp, rwp
    ed = w.res_editor
    assert ed is not None and ed.modified and ed.res.strings[0][1000] == "&File", ed.res.strings
    d1 = w.current_designer()
    assert d1 is not None and d1.module.name == "frmMain"
    keyed, listed = list_images(d1.module, d1.root, "imlToolbarIcons")
    assert len(listed) == 13 and "bold" in keyed and keyed["new"].width() == 16, (len(listed), list(keyed))
    # F5 before Save: the copy that compiles has the String Table of the Resource Editor (AW29m)
    p.res_unsaved = ed.res.to_bytes()
    try:
        cvbp = p.write_copy(os.path.join(d, "wizcopy"))
    finally:
        p.res_unsaved = None
    assert 'ResFile32="Wiz1.RES"' in open(cvbp, encoding="cp1252").read() and p.vbp.get("ResFile32") is None
    from lbstudio.resmodel import ResFile as _Res
    assert _Res.load(os.path.join(d, "wizcopy", "Wiz1.RES")).strings[0][1000] == "&File"
    pd = os.path.join(d, "wiz")
    os.makedirs(pd)
    ed.res.save(os.path.join(pd, "Wiz1.RES"))
    from lbstudio.resedit import set_res_file
    set_res_file(p, "Wiz1.RES")
    p.save(os.path.join(pd, "Wiz1.vbp"))
    cp = subprocess.run([sys.executable, os.path.join(ROOT, "compiler", "lbvc.py"), os.path.join(pd, "Wiz1.vbp"),
                         "-o", os.path.join(pd, "Wiz1")], capture_output=True, text=True, timeout=900)
    assert cp.returncode == 0, (cp.stdout[-3000:], cp.stderr[-3000:])
    # Add-Ins: the add-in loaded gives its two commands (aw18-01)
    assert not w.a_appwiz.isVisible() and not w.a_tbwiz.isVisible()
    am = AddInManagerDialog(w)
    am.state[APPWIZ] = [True, False, False]
    am._ok()
    w.apply_addins()
    assert w.appwiz_loaded and w.a_appwiz.text() == "Application Wizard..." and w.a_tbwiz.text() == "Toolbar Wizard..."
    # the Toolbar Wizard on a new form: three buttons
    for m in p.modules:
        m.dirty = False
    p.dirty = False
    ed.set_modified(False)
    w.new_project("Exe")
    d2 = w.current_designer()
    tw = ToolbarWizard(w, d2)
    assert tw.windowTitle() == "Toolbar Wizard - Introduction"
    tw.go(1)
    assert tw.tbed.chosen.count() == 0 and tw.tbed.avail.item(0).text() == "[Separator]"
    tw.buttons += [G.TbButton("New"), G.TbButton(""), G.TbButton("Bold")]
    tw.tbed.fill()
    assert tw.tbed.chosen.count() == 3
    tw.finish()
    names = [c.name for c in d2.root.children]
    assert names[:2] == ["Toolbar1", "imlToolbarIcons"], names
    assert "Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)" in d2.module.code_text()
    pd2 = os.path.join(d, "tb")
    os.makedirs(pd2)
    saves(os.path.join(pd2, "Form1.frm"), os.path.join(pd2, "Tb.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(pd2, "Form1.frm"), encoding="cp1252").read()
    for want in ("Begin MSComctlLib.Toolbar Toolbar1", 'ImageKey        =   "Bold"', "NumListImages   =   2",
                 'Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"',
                 "Case \"New\""):
        assert want.lower() in frm.lower(), (want, frm[:3000])
    vbp = open(os.path.join(pd2, "Tb.vbp"), encoding="cp1252").read()
    assert "object={831fdd16-0c5c-11d2-a9fc-0000f8754da1}#2.2#0; mscomctl.ocx" in vbp.lower(), vbp
    cp = subprocess.run([sys.executable, os.path.join(ROOT, "compiler", "lbvc.py"), os.path.join(pd2, "Tb.vbp"),
                         "-o", os.path.join(pd2, "Tb")], capture_output=True, text=True, timeout=900)
    assert cp.returncode == 0, (cp.stdout[-3000:], cp.stderr[-3000:])
    # the Toolbar Wizard without buttons: Toolbar2 only (aw20)
    tw = ToolbarWizard(w, d2)
    tw.finish()
    names = [c.name for c in d2.root.children]
    assert names[0] == "Toolbar2" and "imlToolbarIcons2" not in names, names
    assert "Toolbar2_ButtonClick" not in d2.module.code_text()
    done(w)


def t_dfwiz92(d):
    """the Data Form Wizard (phase 9.2b, VB6 oracle DEV-9/df1, df2; the code of the forms is our own: dfwizgen):
    the Add-In Manager's "VB 6 Data Form Wizard" gives Add-Ins > Data Form Wizard...; the pages (Access, Single
    Record / ADO Code on Customers, MS Chart only with ADO Code); Finish: the form in the project with the ADO
    reference, the profile .rwp; the other layouts and bindings from their settings (a class with the member
    "Primary", Master/Detail with MSDataShape, Grid with the ADO Data Control, MS HFlexGrid, MS Chart); lbvc and the
    program: the first record in the text boxes and "Record 1 of 2" from the start, the class's records, the detail
    grid of the master's record, the rows of the grids and of the chart. The Application Wizard's Data Access Forms:
    the settings of Create New Form..., in its report"""
    import subprocess
    from lbstudio import dfwizgen as G
    from lbstudio.dfwizard import DataFormWizard, add_to_project, page_keys
    from lbstudio.resedit import AddInManagerDialog, DFWIZ
    w = MainWindow()
    w.show()
    w.S["dfwiz_no_created"] = True              # no "Data Form Created" in the test
    p = w.project
    w.add_module("Module")
    code(w, p.module("Module1"), MAKE_DB + "\nSub Main()\n    MakeDb\nEnd Sub\n")
    p.vbp.set("Startup", "Sub Main", quote=True)
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Module1.bas"), os.path.join(d, "Dfw.vbp"))
    assert w.save_project_as()
    w.run_start()                               # the database of the wizard
    wait(w)
    db = os.path.join(d, "ide.mdb")
    assert os.path.isfile(db)
    # ---- Add-Ins
    assert not w.a_dfwiz.isVisible()
    am = AddInManagerDialog(w)
    am.state[DFWIZ] = [True, False, False]
    am._ok()
    w.apply_addins()
    assert w.dfwiz_loaded and w.a_dfwiz.isVisible() and w.a_dfwiz.text() == "Data Form Wizard..."
    # ---- the pages
    dlg = DataFormWizard(w)
    assert dlg.windowTitle() == "Data Form Wizard - Introduction" and not dlg.b_finish.isEnabled()
    dlg.step(1)
    assert dlg.windowTitle() == "Data Form Wizard - Database Type"
    assert [dlg.dbtype_list.item(i).text() for i in range(dlg.dbtype_list.count())] == ["Access", "Remote(ODBC)"]
    dlg.dbtype_list.setCurrentRow(0)
    dlg.step(1)
    assert dlg.windowTitle() == "Data Form Wizard - Database"
    dlg.db_name.setText(db)
    dlg.step(1)
    assert dlg.windowTitle() == "Data Form Wizard - Form", dlg.windowTitle()
    dlg.form_name.setText("frmCust")
    dlg.layout_list.setCurrentRow(G.L_CHART)    # VB6 df1-10: MS Chart only with ADO Code
    assert not dlg.bind_group.button(G.B_ADODC).isEnabled() and not dlg.bind_group.button(G.B_CLASS).isEnabled()
    assert dlg.bind_group.checkedId() == G.B_CODE
    dlg.layout_list.setCurrentRow(G.L_FLEX)     # VB6 df2-10: MS HFlexGrid without Class
    assert dlg.bind_group.button(G.B_ADODC).isEnabled() and not dlg.bind_group.button(G.B_CLASS).isEnabled()
    dlg.layout_list.setCurrentRow(G.L_SINGLE)
    dlg.bind_group.button(G.B_CODE).setChecked(True)
    dlg.step(1)
    assert dlg.windowTitle() == "Data Form Wizard - Record Source"
    assert [dlg.src.source.itemText(i) for i in range(dlg.src.source.count())] == ["Customers", "Orders"]
    dlg.src.source.setCurrentIndex(0)
    dlg.src.source_chosen()
    assert dlg.src.fields() == ["CustID", "Name", "Country"], dlg.src.fields()
    dlg.src.sort.setCurrentIndex(dlg.src.sort.findText("Name"))
    dlg.picker_changed()
    assert dlg.b_finish.isEnabled()
    dlg.step(1)
    assert dlg.windowTitle() == "Data Form Wizard - Control Selection" and not dlg.cb_datacontrol.isEnabled()
    dlg.step(1)
    assert dlg.windowTitle() == "Data Form Wizard - Finished!" and not dlg.b_next.isEnabled()
    prof = os.path.join(d, "dfw.rwp")
    dlg.save_combo.addItem(prof)
    dlg.save_combo.setCurrentIndex(dlg.save_combo.count() - 1)
    dlg.finish()
    m = p.module("frmCust")
    assert m is not None and m.dirty and dlg.created is m
    assert "{00000205-0000-0010-8000-00aa006d2ea4}" in "".join(v or "" for k, v in p.vbp.entries).lower()
    rwp = open(prof, encoding="cp1252").read()
    assert "[VB 6 Data Form Wizard]" in rwp and "BindType=1" in rwp and "FormLayout=0" in rwp and db in rwp, rwp
    back = G.load_profile(prof)
    assert back.binding == G.B_CODE and back.database == db and back.buttons["Delete"]
    assert page_keys(back) == ["intro", "dbtype", "database", "form", "source", "controls", "finished"]
    # ---- the other layouts and bindings
    def form(name, layout, binding, source, fields, sort=""):
        df = G.DataForm()
        df.database, df.name, df.layout, df.binding = db, name, layout, binding
        df.source, df.fields, df.sort = source, fields, sort
        return df
    df = form("frmCls", G.L_SINGLE, G.B_CLASS, "Orders", ["OrderID", "CustID", "Amount"], "OrderID")
    add_to_project(w, df)
    assert p.module("clsOrders") is not None and p.module("clsOrders").kind == "Class"
    df = form("frmMD", G.L_MD, G.B_CODE, "Customers", ["CustID", "Name"], "CustID")
    df.detail_source, df.detail_fields, df.master_key, df.detail_key = "Orders", ["OrderID", "Amount"], "CustID", "CustID"
    assert page_keys(df)[4:7] == ["master", "detail", "relation"]
    add_to_project(w, df)
    add_to_project(w, form("frmGrid", G.L_GRID, G.B_ADODC, "Orders", ["OrderID", "CustID", "Amount"]))
    df = form("frmFlex", G.L_FLEX, G.B_CODE, "Customers", ["CustID", "Name", "Country"], "CustID")
    df.grid_style, df.col_widths = 1, {"Country": 0}
    add_to_project(w, df)
    df = form("frmChart", G.L_CHART, G.B_CODE, "Orders", ["CustID", "Amount"])
    df.chart_x, df.chart_y, df.chart_agg, df.chart_style = "CustID", "Amount", 4, 1
    df.captions = {"Title": "Amounts"}
    assert G.check(df) is None
    df.chart_y = "CustID"
    assert "one axis" in G.check(df)
    df.chart_y = "Amount"
    add_to_project(w, df)
    libs = p.components()
    for k in ("MSAdodcLib", "MSDataGridLib", "MSHierarchicalFlexGridLib", "MSChart20Lib"):
        assert k in libs, libs
    assert p.vbp.get("RemoveUnusedControlInfo") is None   # the chart on the form, with its OleObjectBlob
    # a chart drawn in the designer: its OleObjectBlob at once (no _ExtentX), its new size in it when it is saved
    from lbvc.project.chartblob import is_chart_blob, geometry
    from lbstudio.designer import set_prop as dset
    w.open_designer(p.module("frmChart"))
    f = w.current_designer()
    ch = f.place_new("MSChart20Lib.MSChart", QRect(8, 8, 120, 80), None)
    assert ch is not None and ch.name == "MSChart1" and ch.prop("_ExtentX") is None, ch
    assert is_chart_blob(ch.get("OleObjectBlob").data)
    dset(ch, "Width", 3000)
    dset(ch, "Height", 2000)
    # ---- the program
    code(w, p.module("Module1"), MAKE_DB + """
Sub Main()
    Dim f As Integer
    MakeDb
    f = FreeFile
    Open App.Path & "/out.txt" For Output As #f
    Load frmCust
    Print #f, frmCust.txtField(1).Text & "|" & frmCust.lblStatus.Caption
    Load frmCls
    Print #f, frmCls.txtField(0).Text & "|" & frmCls.lblStatus.Caption
    Load frmMD
    Print #f, frmMD.txtField(0).Text & "|" & frmMD.grdDetail.ApproxCount
    Load frmGrid
    Print #f, CStr(frmGrid.grdData.ApproxCount)
    Load frmFlex
    Print #f, frmFlex.flxData.Rows & "|" & frmFlex.flxData.ColWidth(2)
    Load frmChart
    Print #f, frmChart.chtData.RowCount & "|" & frmChart.chtData.TitleText
    frmChart.MSChart1.Row = 1: frmChart.MSChart1.Column = 1
    Print #f, frmChart.MSChart1.RowCount & "|" & frmChart.MSChart1.Plot.SeriesCollection(1).DataPoints(-1).Brush.FillColor.Red & "|" & frmChart.MSChart1.Data
    Close #f
    End
End Sub
""")
    assert w.save_project()
    for n in ("frmCust.frm", "clsOrders.cls", "frmMD.frm", "frmChart.frm"):
        assert os.path.isfile(os.path.join(d, n)), n
    # the charts: their state in the .frx (VB6 cannot load an MSChart without it), no _ExtentX in the .frm
    import re
    ft = open(os.path.join(d, "frmChart.frm"), encoding="cp1252").read()
    offs = [int(x, 16) for x in re.findall(r'OleObjectBlob   =   "frmChart.frx":([0-9A-F]{4})', ft)]
    assert len(offs) == 2 and "_ExtentX" not in ft, ft[:1500]
    frx = open(os.path.join(d, "frmChart.frx"), "rb").read()
    geos = [geometry(frx[o:]) for o in offs]
    assert all(is_chart_blob(frx[o:]) for o in offs), geos
    assert (60, 60, 6360, 3840) in geos and sorted(g[2:] for g in geos)[0] == (3000, 2000), geos
    cp = subprocess.run([sys.executable, os.path.join(ROOT, "compiler", "lbvc.py"), os.path.join(d, "Dfw.vbp"),
                         "-o", os.path.join(d, "Dfw")], capture_output=True, text=True, timeout=900)
    assert cp.returncode == 0, (cp.stdout[-3000:], cp.stderr[-3000:])
    out = os.path.join(d, "out.txt")
    if os.path.exists(out):
        os.remove(out)
    w.run_start()
    wait(w, 120)
    got = open(out, encoding="cp1252").read().splitlines()
    assert got == ["Alfreds|Record 1 of 2", "10|Record 1 of 3", "ALFKI|2", "3", "3|0", "2|Amounts", "5|255|30"], got
    # ---- the Application Wizard's Data Access Forms: the settings of Create New Form..., in its report
    from lbstudio.appwizard import AppWizard
    from lbstudio import appwizgen as AG
    aw = AppWizard(w)
    df = form("frmOrders", G.L_GRID, G.B_CODE, "Orders", ["OrderID", "Amount"])
    aw.opts.data_forms.append(df)
    aw.fill_data_forms()
    assert aw.data_list.count() == 1 and aw.chosen_data_forms() == [df]
    assert "Data forms: frmOrders (Orders)" in AG.summary(aw.opts)
    aw.reject()
    done(w)


def _cb_select(cb, *path):
    """the class of the tree of the Class Builder at that path from Project1"""
    it = cb.tree.topLevelItem(0)
    for n in path:
        it = next(it.child(i) for i in range(it.childCount()) if it.child(i).text(0) == n)
    cb.tree.setCurrentItem(it)
    return it


def _cb_tree(cb):
    out = []

    def walk(it, depth):
        out.append("  " * depth + it.text(0))
        for i in range(it.childCount()):
            walk(it.child(i), depth + 1)
    walk(cb.tree.topLevelItem(0), 0)
    return out


def _cb_rows(cb, tab=3):
    t = cb.lists[tab]
    return [" | ".join(t.topLevelItem(i).text(k) for k in range(t.columnCount())).rstrip(" |")
            for i in range(t.topLevelItemCount())]


def _cb_member(cb, name, tab=3):
    cb.tabs.setCurrentIndex(tab)
    t = cb.lists[tab]
    it = next(t.topLevelItem(i) for i in range(t.topLevelItemCount()) if t.topLevelItem(i).text(0) == name)
    t.setCurrentItem(it)
    return it.member


def _cb_prop(cb, c, name, type_, decl="public", default=False, desc=""):
    pb = cb.new_property()
    pb.name.setText(name)
    pb.type.setCurrentText(type_)
    {"public": pb.rb_public, "friend": pb.rb_friend, "var": pb.rb_var}[decl].setChecked(True)
    pb.cb_default.setChecked(default)
    pb.desc.setPlainText(desc)
    assert pb.apply(), pb.messages
    return cb.add_member(c, pb.result_member)


def t_axwiz92(d):
    """the ActiveX Control Interface Wizard (phase 9.2d, VB6 oracle DEV-9/ax1, ax2; the code is our own: axwizgen):
    the Add-In Manager's "VB 6 ActiveX Ctrl Interface Wizard" gives Add-Ins > ActiveX Control Interface Wizard...;
    a project without a user control; the pages (Introduction, Select Interface Members with the standard members
    chosen, Create Custom Interface Members: Add Custom Member, "Invalid Name!", a name already there, Delete; Set
    Mapping: a member of a control on the user control; Set Attributes: the unmapped members); Finish: the code
    ('MemberInfo / 'MappingInfo, Const m_def_, Dim m_, the events, Init/Read/WriteProperties); the summary report;
    the wizard again: it reads what it wrote, the changed members commented out and written again (Read Only, Not
    Available at run time, a method with arguments); lbvc and the program: the control on Form1"""
    import subprocess
    from lbstudio import axwizgen as G
    from lbstudio.axwizard import AxWizard, CustomMemberDialog, SummaryDialog
    from lbstudio.resedit import AddInManagerDialog, AXWIZ, ADDINS
    w = MainWindow()
    w.show()
    p = w.project
    # ---- Add-Ins
    assert not w.a_axwiz.isVisible()
    assert any(n == AXWIZ and desc for n, desc in ADDINS)
    am = AddInManagerDialog(w)
    am.state[AXWIZ] = [True, False, False]
    am._ok()
    w.apply_addins()
    assert w.axwiz_loaded and w.a_axwiz.isVisible() and w.a_axwiz.text() == "ActiveX Control Interface Wizard..."
    # ---- a project without a user control
    wz = AxWizard(w, quiet=True)
    assert wz.uc is None
    wz.step(1)
    assert wz.key == "intro" and wz.messages and "user control" in wz.messages[0]
    wz.reject()
    # ---- the user control: Text1, Command1
    w.add_module("UserControl")
    uc = p.module("UserControl1")
    u = w.current_designer()
    u.set_client_size(3000, 1200)
    u.place_new("TextBox", QRect(8, 8, 120, 24), None)
    u.place_new("CommandButton", QRect(8, 40, 100, 30), None)
    ctls = [(n, t) for n, t, _ in uc.controls()]       # VB6 writes the newest control first
    assert ctls == [("Command1", "CommandButton"), ("Text1", "TextBox")], ctls
    # ---- the pages
    wz = AxWizard(w, quiet=True)
    assert wz.windowTitle() == "ActiveX Control Interface Wizard - Introduction"
    assert wz.page_keys() == ["intro", "members", "custom", "mapping", "attributes", "finished"]
    assert not wz.b_back.isEnabled() and wz.b_next.isEnabled() and not wz.b_finish.isEnabled()
    assert wz.cb_skip.text() == "&Skip this screen in the future." and not wz.cb_skip.isChecked()
    wz.step(1)
    assert wz.key == "members" and wz.windowTitle().endswith("Select Interface Members") and wz.b_finish.isEnabled()

    def names(t):
        return [t.topLevelItem(i).text(0) for i in range(t.topLevelItemCount())]
    assert names(wz.chosen) == sorted(G.DEFAULT_SELECTED, key=str.lower), names(wz.chosen)
    assert "Text" in names(wz.avail) and "Text" not in names(wz.chosen)
    kinds = {wz.chosen.topLevelItem(i).text(0): wz.chosen.topLevelItem(i).text(1)
             for i in range(wz.chosen.topLevelItemCount())}
    assert kinds["Click"] == "Event" and kinds["BackColor"] == "Property" and kinds["Refresh"] == "Method", kinds
    wz.move_members(1, False, ["Text"])
    wz.move_members(-1, False, ["DblClick"])
    assert "Text" in names(wz.chosen) and "DblClick" in names(wz.avail) and "Text" not in names(wz.avail)
    wz.move_members(1, False, ["DblClick"])
    wz.step(1)
    assert wz.key == "custom" and wz.windowTitle().endswith("Create Custom Interface Members")
    dl = wz.new_custom()
    assert isinstance(dl, CustomMemberDialog) and dl.windowTitle() == "Add Custom Member"
    assert [rb.text() for rb in dl.kinds.values()] == ["&Property", "&Method", "&Event"] and dl.kind() == "property"
    dl.reject()
    wz.new_custom("Amount", "property")
    wz.new_custom("Restart", "method")
    wz.new_custom("Changed", "event")
    assert wz.new_custom("Go") is None and wz.messages[-1] == "Invalid Name!"
    assert wz.new_custom("Text") is None and wz.messages[-1] == "'Text' is already a member of the control."
    wz.new_custom("Temp", "property")
    assert names(wz.custom_list) == ["Amount", "Temp", "Restart", "Changed"]   # properties, methods, events (VB6)
    wz.custom_list.setCurrentItem(wz.custom_list.topLevelItem(1))
    assert wz.b_delete.isEnabled()
    wz.delete_custom()
    assert names(wz.custom_list) == ["Amount", "Restart", "Changed"]
    assert [wz.custom_list.topLevelItem(i).text(1) for i in range(3)] == ["Property", "Method", "Event"]
    from PyQt6.QtWidgets import QLabel
    labs = [lab for lab in wz.findChildren(QLabel) if "&" in lab.text()]
    assert len(labs) >= 10 and all(lab.buddy() is not None for lab in labs), \
        [lab.text() for lab in labs if lab.buddy() is None]                 # the & keys, not a & on the screen
    wz.step(1)
    assert wz.key == "mapping" and wz.windowTitle().endswith("Set Mapping")
    assert [wz.map_ctl.itemText(i) for i in range(wz.map_ctl.count())] == ["(None)", "Command1", "Text1", "UserControl"]
    assert names(wz.map_names) == wz.model.all_names() and "Amount" in names(wz.map_names)
    it = next(wz.map_names.topLevelItem(i) for i in range(wz.map_names.topLevelItemCount())
              if wz.map_names.topLevelItem(i).text(0) == "Text")
    wz.map_names.clearSelection()
    it.setSelected(True)
    wz.map_ctl.setCurrentText("Text1")
    wz.mapping_control_changed()
    t = wz.model.member("Text")
    assert (t.map_ctl, t.map_member, t.type) == ("Text1", "Text", "String"), (t.map_ctl, t.map_member, t.type)
    wz.set_mapping("Click", "Command1", "Click")
    wz.step(1)
    assert wz.key == "attributes" and wz.windowTitle().endswith("Set Attributes")
    un = names(wz.attr_names)
    assert "Text" not in un and "Click" not in un and "Amount" in un and un == wz.model.unmapped(), un
    it = next(wz.attr_names.topLevelItem(i) for i in range(wz.attr_names.topLevelItemCount())
              if wz.attr_names.topLevelItem(i).text(0) == "Amount")
    wz.attr_names.setCurrentItem(it)
    assert wz.attr_type.currentText() == "Variant" and wz.attr_default.isEnabled() and not wz.attr_args.isEnabled()
    wz.attr_type.setCurrentText("Long")
    wz.attr_default.setText("3")
    wz.attr_desc.setPlainText("The amount")
    wz.step(1)
    a = wz.model.member("Amount")
    assert (a.type, a.default, a.description) == ("Long", "3", "The amount"), (a.type, a.default, a.description)
    assert wz.key == "finished" and wz.windowTitle().endswith("Finished!") and wz.cb_summary.isChecked()
    new = wz.finish()
    ct = uc.code_text()
    assert ct == new.rstrip("\n") and uc.dirty
    for s in ("'MemberInfo=8,0,0,3\nPublic Property Get Amount() As Long", "Const m_def_Amount = 3",
              "Dim m_Amount As Long", "Event Changed()", "Event Click() 'MappingInfo=Command1,Command1,-1,Click",
              "Private Sub Command1_Click()\n    RaiseEvent Click\nEnd Sub",
              "'MappingInfo=Text1,Text1,-1,Text\nPublic Property Get Text() As String",
              "    Text = Text1.Text", "'MemberInfo=14\nPublic Function Restart() As Variant",
              "Private Sub UserControl_InitProperties()", "    m_Amount = PropBag.ReadProperty(\"Amount\", m_def_Amount)",
              "    Text1.Text = PropBag.ReadProperty(\"Text\", \"Text1\")",
              "    Call PropBag.WriteProperty(\"Amount\", m_Amount, m_def_Amount)",
              'Attribute Amount.VB_Description = "The amount"'):
        assert s in ct, (s, ct)
    sd = wz.summary_dialog
    assert isinstance(sd, SummaryDialog) and sd.windowTitle() == "ActiveX Control Interface Wizard Summary"
    rep = os.path.join(d, "ctlwiz.txt")
    sd.save(rep)
    assert open(rep, encoding="utf-8").read() == G.summary(wz.model) and "Amount" in G.summary(wz.model)
    # ---- the wizard again: it reads what it wrote (the code window open)
    w.open_code(uc)
    wz = AxWizard(w, quiet=True)
    wz.step(1)
    m = wz.model
    assert "Text" in m.selected and sorted(m.custom) == ["Amount", "Changed", "Restart"], (m.selected, m.custom)
    assert (m.member("Text").map_ctl, m.member("Click").map_member) == ("Text1", "Click")
    assert (m.member("Amount").type, m.member("Amount").default) == ("Long", "3")
    wz.go_to("custom")
    wz.new_custom("Secret", "property")
    wz.go_to("attributes")
    wz.set_attributes("Amount", runtime=1)
    wz.set_attributes("Restart", type="String", args="ByVal Times As Integer")
    wz.set_attributes("Secret", type="String", runtime=3)
    wz.go_to("finished")
    wz.cb_summary.setChecked(False)
    wz.finish()
    ct = uc.code_text().replace("\r\n", "\n")
    assert ct.rstrip("\n") == w.code_windows[uc].widget().text().replace("\r\n", "\n").rstrip("\n"), ct
    for s in ("''MemberInfo=14\n'Public Function Restart() As Variant",
              "'MemberInfo=13\nPublic Function Restart(ByVal Times As Integer) As String",
              "'MemberInfo=8,1,0,3\nPublic Property Get Amount() As Long",
              "    If Ambient.UserMode Then Err.Raise 382\n    m_Amount = New_Amount",
              "'MemberInfo=13,3,0,0\nPublic Property Get Secret() As String",
              "    If Ambient.UserMode Then Err.Raise 393", "Dim m_Secret As String"):
        assert s in ct, (s, ct)
    assert ct.count("Public Property Get Text() As String") == 1 and ct.count("Private Sub UserControl_ReadProperties") == 1
    # ---- the program: the control on Form1
    w.fill_tree()
    assert "Project1.UserControl1" in w.toolbox.buttons
    # the Properties window: the defaults of ReadProperty (Const m_def_Amount = 3; "Text1")
    assert V.prop_def("UserControl1", "Amount").default == 3 and V.prop_def("UserControl1", "Text").default == "Text1"
    assert V.prop_def("UserControl1", "Enabled").default is False
    f = w.open_designer(p.module("Form1"))
    assert f.place_new("Project1.UserControl1", QRect(16, 16, 0, 0), None) is not None
    code(w, p.module("Form1"), """Private Sub Form_Load()
    Dim f As Integer, s As String
    f = FreeFile
    Open App.Path & "/out.txt" For Output As #f
    Print #f, UserControl11.Text & "|" & UserControl11.Amount
    UserControl11.Text = "abc"
    On Error Resume Next
    UserControl11.Amount = 9
    Print #f, Err.Number & "|" & UserControl11.Text & "|" & UserControl11.Amount & "|" & TypeName(UserControl11.Restart(2))
    Err.Clear
    s = UserControl11.Secret
    Print #f, Err.Number
    On Error GoTo 0
    Close #f
    Unload Me
End Sub
""")
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "UserControl1.ctl"), os.path.join(d, "Ax.vbp"))
    assert w.save_project_as()
    ctl = open(os.path.join(d, "UserControl1.ctl"), encoding="cp1252").read()
    assert "Begin VB.TextBox Text1" in ctl and "'MappingInfo=Text1,Text1,-1,Text" in ctl, ctl
    cp = subprocess.run([sys.executable, os.path.join(ROOT, "compiler", "lbvc.py"), os.path.join(d, "Ax.vbp"),
                         "-o", os.path.join(d, "Ax")], capture_output=True, text=True, timeout=900)
    assert cp.returncode == 0, (cp.stdout[-3000:], cp.stderr[-3000:])
    out = os.path.join(d, "out.txt")
    want = ["Text1|3", "382|abc|3|String", " 393 "]          # a number: Print # puts its spaces
    try:                                        # the program lbvc made, on its own: what it says when it fails
        rp = subprocess.run([os.path.join(d, "Ax.exe" if os.name == "nt" else "Ax")], cwd=d, capture_output=True,
                            text=True, timeout=60)
        why = (rp.returncode, rp.stdout[-2000:], rp.stderr[-2000:])
    except subprocess.TimeoutExpired as ex:
        why = ("timeout", str(ex.stdout)[-2000:], str(ex.stderr)[-2000:])
    got = open(out, encoding="cp1252").read().splitlines() if os.path.isfile(out) else None
    assert got == want, (got, why)
    os.remove(out)
    # F5: the Err.Raise of the user control's Property Let breaks there with Break in Class Module (VB6's default,
    # as the IDE does); Break on Unhandled Errors gives the error to the On Error Resume Next of Form_Load
    w.trap_session = 2                          # Toggle > Break on Unhandled Errors (this session)
    w.run_start()
    try:
        wait(w, 120)
    except AssertionError:
        raise AssertionError(("F5", w.runner.state, w.runner.stderr[-2000:]))
    w.trap_session = None
    got = open(out, encoding="cp1252").read().splitlines()
    assert got == want, got
    done(w)


def t_classbuilder92(d):
    """the Class Builder Utility (phase 9.2c, VB6 oracle DEV-9/cb1 ... cb6; the code is our own: cbmodel): the
    Add-In Manager's "VB 6 Class Builder Utility" gives Add-Ins > Class Builder Utility...; Project > Add Class
    Module offers Class Module, VB Class Builder, Complex Data Consumer, Data Source; the window (menus, toolbar,
    Classes, the Properties / Methods / Events / All lists); the builders of a class (Based On: the members copied),
    a collection (of an existing class; of a new class with Key), a property, a method (Add Argument), an event,
    an enum; "Invalid Name!" for the words of the language (Text, Go), "'X' already exists!"; a class made in
    another one is a member of it; Update Project: the modules, their VB_Ext_KEY attributes; the Class Builder
    opened again: the tree from the attributes, the members in the code cannot be deleted / renamed / changed, a
    class it did not make, new members added to a class without changing its code, a class renamed, View >
    Options (modErrorHandling, modClassIDGenerator), Exit with changes; a tree that goes round (VB6 DEV-9/cb2-10)
    ends; lbvc and the program: properties, the default member, the class in the class, the collections (Add,
    Item, Count, Remove, For Each), Based On, the debug number"""
    import subprocess
    from lbstudio import cbmodel as CM
    from lbstudio.classbuilder import ClassBuilder, AddClassModuleDialog, add_template, EXISTING_MEMBER
    from lbstudio.resedit import AddInManagerDialog, CBUILDER
    w = MainWindow()
    w.show()
    p = w.project
    # ---- Add-Ins, Project > Add Class Module
    assert not w.a_cbuilder.isVisible()
    am = AddInManagerDialog(w)
    am.state[CBUILDER] = [True, False, False]
    am._ok()
    w.apply_addins()
    assert w.cbuilder_loaded and w.a_cbuilder.isVisible() and w.a_cbuilder.text() == "Class Builder Utility..."
    ac = AddClassModuleDialog(w, d)
    assert [ac.items.item(i).text() for i in range(ac.items.count())] == \
        ["Class Module", "VB Class Builder", "Complex Data Consumer", "Data Source"]
    assert [ac.tabs.tabText(i) for i in range(2)] == ["New", "Existing"]
    ac.reject()
    # ---- the window
    cb = ClassBuilder(w, quiet=True)
    assert cb.windowTitle() == "Class Builder"
    assert [a.text() for a in cb.menubar.actions()] == ["&File", "&Edit", "&View", "&Help"]
    assert [a.toolTip() for a in cb.toolbar.actions() if not a.isSeparator()] == \
        ["Add New Class", "Add New Collection", "Add New Property to Current Class", "Add New Method to Current Class",
         "Add New Event to Current Class", "Add New Enum to Current Class", "Show Properties for the current object",
         "Cut", "Copy", "Paste"]
    assert [cb.tabs.tabText(i) for i in range(4)] == ["Properties", "Methods", "Events", "All"]
    assert _cb_tree(cb) == ["Project1"] and not cb.a_new_prop.isEnabled()
    # ---- clsPerson
    dl = cb.new_class()
    assert dl.windowTitle() == "Class Module Builder" and dl.name.text() == "Class1"
    assert [dl.based.itemText(i) for i in range(dl.based.count())] == ["(New Class)"]
    dl.name.setText("clsPerson")
    dl.desc.setPlainText("A person")
    assert dl.apply()
    person = cb.new_class(dl)
    assert _cb_tree(cb) == ["Project1", "  clsPerson"] and cb.current_class() is person and cb.a_new_prop.isEnabled()
    pb = cb.new_property()
    assert pb.windowTitle() == "Property Builder"
    assert [pb.rb_public.text(), pb.rb_friend.text(), pb.rb_var.text(), pb.cb_default.text()] == \
        ["&Public Property (Let, Get, Set)", "&Friend Property (Let, Get, Set)", "Public &Variable", "De&fault Property?"]
    pb.name.setText("Text")                     # a word of the language (VB6 DEV-9/cb1-27)
    assert not pb.apply() and pb.messages == ["Invalid Name!"]
    _cb_prop(cb, person, "FullName", "String", default=True, desc="The full name")
    _cb_prop(cb, person, "Age", "Integer", "friend")
    _cb_prop(cb, person, "Tag", "Variant", "var")
    _cb_prop(cb, person, "Photo", "Object")
    _cb_prop(cb, person, "Extra", "Variant")
    pb = cb.new_property()
    pb.name.setText("Age")
    assert not pb.apply() and pb.messages == ["'Age' already exists!"]
    mb = cb.new_method()
    assert mb.windowTitle() == "Method Builder" and mb.ret.currentText() == "(None)"
    mb.name.setText("Speak")
    ad = mb.add_arg()
    assert ad.windowTitle() == "Add Argument" and ad.type.currentText() == "Variant"
    ad.name.setText("Text")
    ad._ok()
    assert ad.messages == ["Invalid Name!"] and ad.result_arg is None
    ad.name.setText("Phrase")
    ad.byval.setChecked(True)
    assert not ad.array.isEnabled() and not ad.default.isEnabled()
    ad.type.setCurrentText("String")
    ad._ok()
    mb.add_arg(ad.result_arg)
    ad = mb.add_arg()
    ad.name.setText("Repeats")
    ad.optional.setChecked(True)
    ad.type.setCurrentText("Integer")
    ad.default.setText("1")
    ad._ok()
    mb.add_arg(ad.result_arg)
    assert [mb.arg_list.item(i).text() for i in range(mb.arg_list.count())] == \
        ["ByVal Phrase As String", "Optional Repeats As Integer = 1"]
    mb.ret.setCurrentText("Boolean")
    assert mb.apply()
    cb.add_member(person, mb.result_member)
    mb = cb.new_method()
    mb.name.setText("Go")                       # VB6 DEV-9/cb3
    assert not mb.apply() and mb.messages == ["Invalid Name!"]
    mb.name.setText("Reset")
    mb.cb_friend.setChecked(True)
    assert mb.apply()
    cb.add_member(person, mb.result_member)
    eb = cb.new_event()
    eb.name.setText("Changed")
    ad = eb.add_arg()
    ad.name.setText("Reason")
    ad.byval.setChecked(True)
    ad.type.setCurrentText("String")
    ad._ok()
    eb.add_arg(ad.result_arg)
    assert eb.apply()
    cb.add_member(person, eb.result_member)
    en = cb.new_enum()
    assert en.windowTitle() == "Enum Builder" and not en.b_ok.isEnabled()
    en.name.setText("PersonKind")
    en.add_value("pkChild = 1")
    en.add_value("pkAdult=2")
    assert en.b_ok.isEnabled()
    en._ok()
    cb.add_member(person, en.result_member)
    cb.class_changed()
    assert _cb_rows(cb, 3) == [
        "Age | Integer |  | Get/Let", "Changed |  | ByVal Reason As String", "Extra | Variant |  | Get/Let/Set",
        "FullName | String |  | Get/Let", "PersonKind", "Photo | Object |  | Get/Set", "Reset",
        "Speak | Boolean | ByVal Phrase As String, Optional Repeats As Integer = 1", "Tag | Variant |  | Var"], \
        _cb_rows(cb, 3)
    assert _cb_rows(cb, 2) == ["Changed | ByVal Reason As String"]
    # ---- a class in clsPerson; collections; Based On
    dl = cb.new_class()
    dl.name.setText("clsAddress")
    assert dl.apply()
    addr = cb.new_class(dl)
    _cb_prop(cb, addr, "City", "String")
    _cb_select(cb)
    cd = cb.new_collection()
    assert cd.windowTitle() == "Collection Builder" and cd.name.text() == "Collection1"
    assert [cd.existing.item(i).text() for i in range(cd.existing.count())] == ["clsPerson", "clsAddress"]
    cd.name.setText("colPersons")
    cd.rb_existing.setChecked(True)
    assert not cd.apply() and cd.messages
    cd.existing.setCurrentRow(0)
    assert cd.apply()
    cb.new_collection(cd)
    _cb_select(cb)
    cd = cb.new_collection()
    cd.name.setText("colOrders")
    cd.rb_new.setChecked(True)
    cd.new_name.setText("clsOrder")
    assert cd.apply()
    cb.new_collection(cd)
    _cb_select(cb, "colOrders", "clsOrder")
    order = cb.current_class()
    assert order.name == "clsOrder" and _cb_rows(cb, 3) == ["Key | String |  | Var"]
    _cb_prop(cb, order, "Amount", "Currency")
    _cb_select(cb)
    dl = cb.new_class()
    dl.name.setText("clsBase")
    assert dl.apply()
    base = cb.new_class(dl)
    _cb_prop(cb, base, "Value", "Long")
    mb = cb.new_method()
    mb.name.setText("Run")
    assert mb.apply()
    cb.add_member(base, mb.result_member)
    _cb_select(cb)
    dl = cb.new_class()
    dl.name.setText("clsDerived")
    dl.based.setCurrentText("clsBase")
    assert dl.apply()
    cb.new_class(dl)
    assert _cb_rows(cb, 3) == ["Run", "Value | Long |  | Get/Let"]    # Based On: the members (VB6 DEV-9/cb2)
    _cb_member(cb, "Value")
    assert cb.copy(target="member")
    cb.paste()
    assert cb.messages[-1] == "'Value' already exists!"
    _cb_select(cb)
    dl = cb.new_class()
    dl.name.setText("clsList")
    assert dl.apply()
    cb.new_class(dl)
    cb.set_as_collection()
    assert cb.current_class().user_collection
    _cb_select(cb)
    dl = cb.new_class()
    dl.name.setText("clsTmp")
    assert dl.apply()
    cb.new_class(dl)
    assert cb.delete(target="class") and cb.model.cls("clsTmp") is None
    assert _cb_tree(cb) == ["Project1", "  clsBase", "  clsDerived", "  clsList", "  clsPerson", "    clsAddress",
                            "  colOrders", "    clsOrder", "  colPersons", "    clsPerson"], _cb_tree(cb)
    # ---- Update Project
    assert cb.model.dirty
    assert cb.update_project() and not cb.model.dirty
    for n in ("clsPerson", "clsAddress", "colPersons", "colOrders", "clsOrder", "clsBase", "clsDerived", "clsList"):
        m = p.module(n)
        assert m is not None and m.kind == "Class" and m.dirty, n
    keys = CM.ext_keys(p.module("clsPerson").sf.attributes)
    assert keys == [("SavedWithClassBuilder6", "Yes"), ("Member0", "clsAddress"), ("Top_Level", "Yes")], keys
    assert CM.ext_keys(p.module("colPersons").sf.attributes) == [
        ("SavedWithClassBuilder6", "Yes"), ("Collection", "clsPerson"), ("Member0", "clsPerson"), ("Top_Level", "Yes")]
    assert ("Top_Level", "No") in CM.ext_keys(p.module("clsAddress").sf.attributes)
    assert ("Collection", "UserDefinedCollection") in CM.ext_keys(p.module("clsList").sf.attributes)
    pt = p.module("clsPerson").code_text()
    for line in ("Private mvarFullName As String", "Public Tag As Variant", "Public Event Changed(ByVal Reason As String)",
                 "Friend Property Let Age(ByVal vData As Integer)", "Attribute FullName.VB_UserMemId = 0",
                 "Public Function Speak(ByVal Phrase As String, Optional Repeats As Integer = 1) As Boolean",
                 "    Set mvarclsAddress = New clsAddress", "    pkAdult = 2", "Friend Sub Reset()"):
        assert line in pt.split("\n"), line
    assert "Public Function Add(Key As String, Amount As Currency, Optional sKey As String) As clsOrder" in \
        p.module("colOrders").code_text()
    cb.reject()
    # ---- the Class Builder again: the tree of the attributes, a class it did not make
    w.add_module("Class")
    code(w, p.module("Class1"), "Public Nickname As String\nPublic Sub Hello()\nEnd Sub\n")
    w._sync_editors()
    cb = ClassBuilder(w, quiet=True)
    assert cb.foreign_message() is not None
    assert _cb_tree(cb) == ["Project1", "  Class1", "  clsBase", "  clsDerived", "  clsList", "  clsPerson",
                            "    clsAddress", "  colOrders", "    clsOrder", "  colPersons", "    clsPerson"], \
        _cb_tree(cb)
    _cb_select(cb, "clsPerson")
    assert "FullName | String |  | Get/Let" in _cb_rows(cb, 3) and "clsAddress | clsAddress |  | Get/Set" in _cb_rows(cb, 3)
    _cb_member(cb, "FullName")
    assert not cb.delete(target="member")
    assert cb.messages[-1] == EXISTING_MEMBER.format(name="FullName", what="delete", cls="clsPerson")
    assert not cb.rename_member(cb.current_class(), cb.current_member(), "Fullname2")
    cb.properties(target="member")
    assert cb.last_dialog.windowTitle() == "Property Builder" and not cb.last_dialog.name.isEnabled()
    assert not cb.delete(target="class") and "class module of the project already" in cb.messages[-1]
    _cb_prop(cb, cb.current_class(), "Nickname", "String")
    _cb_select(cb, "clsList")
    assert cb.rename_class(cb.current_class(), "clsItems")
    cb.options(True, True)                      # View > Options: debug code, Err.Raise
    _cb_select(cb)
    dl = cb.new_class()
    dl.name.setText("clsDbg")
    assert dl.apply()
    dbg = cb.new_class(dl)
    _cb_prop(cb, dbg, "Value", "Long")
    cb.answer = "yes"                           # Exit: "Update Project with changes?" Yes
    cb.reject()
    assert cb.messages[-1] == "Update Project with changes?" and not cb.model.dirty
    w.S["cb_debug_code"] = w.S["cb_err_raise"] = False
    pt2 = p.module("clsPerson").code_text()
    assert "Private mvarNickname As String" in pt2 and "Public Property Let Nickname(ByVal vData As String)" in pt2
    rest = iter(pt2.split("\n"))                # the old code as it was, in its order
    assert all(any(x == y for y in rest) for x in pt.split("\n") if x.strip()), pt2
    assert p.module("clsItems") is not None and p.module("clsList") is None
    assert ("SavedWithClassBuilder6", "Yes") in CM.ext_keys(p.module("Class1").sf.attributes)
    for n in ("modErrorHandling", "modClassIDGenerator"):
        assert p.module(n) is not None and p.module(n).kind == "Module", n
    dt = p.module("clsDbg").code_text()
    assert "#Const DebugMode = 1" in dt and "    On Error GoTo ValueLetErr" in dt, dt
    # a tree that goes round (VB6 DEV-9/cb2-10: clsAddress in clsAddress) ends
    mm = CM.Model()
    a = CM.CBClass("clsLoop")
    a.children = ["clsLoop"]
    mm.add(a)
    assert [(c.name, len(k)) for c, _, k in mm.tree()] == [("clsLoop", 1)] and mm.tree()[0][2][0][2] == []
    # ---- the templates of Add Class Module
    add_template(p, "source")
    add_template(p, "consumer")
    assert p.module("clsDataSource").sf.class_props is not None
    assert "Private Sub Class_GetDataMember(DataMember As String, Data As Object)" in p.module("clsDataSource").code_text()
    # ---- the program
    w.add_module("Module")
    code(w, p.module("Module1"), """
Sub Main()
    Dim f As Integer, ps As clsPerson, c As colPersons, o As colOrders, x As clsPerson, s As String
    Dim g As New clsDerived, dbg As New clsDbg
    f = FreeFile
    Open App.Path & "/out.txt" For Output As #f
    Set ps = New clsPerson
    ps.FullName = "Ann"
    ps.Age = 30
    ps.Tag = "t"
    ps.Extra = 5
    ps.clsAddress.City = "Athens"
    Print #f, ps & "|" & ps.Age & "|" & ps.Tag & "|" & ps.Extra & "|" & ps.clsAddress.City
    Print #f, ps.Speak("hi") & "|" & pkAdult
    Set c = New colPersons
    c.Add "Bob", 40, "", Nothing, 1, Nothing, "b"
    c.Add "Cy", 50, "", Nothing, 2, Nothing
    For Each x In c
        s = s & x.FullName & ","
    Next
    Print #f, c.Count & "|" & c("b").FullName & "|" & c(2).Age & "|" & s
    c.Remove 1
    Print #f, c.Count
    Set o = New colOrders
    o.Add "k1", 12.5, "k1"
    Print #f, o("k1").Amount & "|" & o(1).Key
    g.Value = 7
    Print #f, g.Value
    dbg.Value = 3
    Print #f, dbg.Value & "|" & dbg.ClassDebugID
    Close #f
    End
End Sub
""")
    p.vbp.set("Startup", "Sub Main", quote=True)
    files = [os.path.join(d, n) for n in ("Form1.frm", "clsPerson.cls", "clsAddress.cls", "colPersons.cls",
                                          "clsOrder.cls", "colOrders.cls", "clsBase.cls", "clsDerived.cls",
                                          "clsItems.cls", "Class1.cls", "clsDbg.cls", "modErrorHandling.bas",
                                          "modClassIDGenerator.bas", "clsDataSource.cls",
                                          "clsComplexDataConsumer.cls", "Module1.bas", "Cb.vbp")]
    saves(*files)
    assert w.save_project_as()
    for f in files:
        assert os.path.isfile(f), f
    cp = subprocess.run([sys.executable, os.path.join(ROOT, "compiler", "lbvc.py"), os.path.join(d, "Cb.vbp"),
                         "-o", os.path.join(d, "Cb")], capture_output=True, text=True, timeout=900)
    assert cp.returncode == 0, (cp.stdout[-3000:], cp.stderr[-3000:])
    out = os.path.join(d, "out.txt")
    w.run_start()
    wait(w, 120)
    got = open(out, encoding="cp1252").read().splitlines()
    assert got == ["Ann|30|t|5|Athens", "False|2", "2|Bob|50|Bob,Cy,", " 1 ", "12.5|k1", " 7 ", "3|1"], got   # a number: Print # puts its spaces
    done(w)


def t_appwizgen92(d):
    """what the Application Wizard writes (phase 9.2a-2, appwizgen): the profile of the default choices as VB6 wrote
    it (VB6 oracle DEV-9/aw6 test1.rwp); four projects of the VB6 oracle aw1 (SDI with Resources, Internet and the
    four form templates: the String Table of its .res from 1000; SDI with the standard forms; MDI; Explorer with
    Options and About): their modules, .vbp, menus, toolbar pictures in the .frx; each compiled by lbvc"""
    import subprocess
    from lbstudio import appwizgen as G, tbicons
    from lbstudio.resedit import set_res_file
    rwp = open(os.path.join(HERE, "aw", "test1.rwp"), encoding="cp1252", newline="").read()
    o = G.AppOptions(G.APP_MDI)
    assert o.to_rwp() == rwp, "\n".join(f"{a!r} | {b!r}" for a, b in zip(o.to_rwp().splitlines(), rwp.splitlines())
                                       if a != b)
    back = G.AppOptions.from_rwp(rwp)
    assert back.to_rwp() == rwp and back.app_type == G.APP_MDI
    sdi = G.AppOptions(G.APP_SDI)
    assert [m.key for m in sdi.menus if m.on] == ["File", "Edit", "View", "Help"], [m.key for m in sdi.menus]
    exp = G.AppOptions(G.APP_EXPLORER)
    assert [m.key for m in exp.menus if m.on] == ["File", "Edit", "View", "Help"]
    assert [b.key for b in exp.toolbar][:3] == ["", "Back", "Forward"] and len(exp.toolbar) == 15

    def make(name, o):
        r = G.generate(o)
        p = G.build_project(r, o, tbicons.bmp, "cp1252")
        pd = os.path.join(d, name)
        os.makedirs(pd)
        res = G.result_res(r)
        if res is not None:
            res.save(os.path.join(pd, "Project1.RES"))
            set_res_file(p, "Project1.RES")
        p.save(os.path.join(pd, "Project1.vbp"))
        vbp = open(os.path.join(pd, "Project1.vbp"), encoding="cp1252").read()
        exe = os.path.join(pd, "Project1")
        cp = subprocess.run([sys.executable, os.path.join(ROOT, "compiler", "lbvc.py"), os.path.join(pd, "Project1.vbp"),
                             "-o", exe], capture_output=True, text=True, timeout=900)
        assert cp.returncode == 0, (name, cp.stdout[-3000:], cp.stderr[-3000:])
        return r, pd, vbp

    def frm(pd, name):
        return open(os.path.join(pd, name + ".frm"), encoding="cp1252").read()

    # 1. SDI, Resources, Internet, the four templates (aw1 project-sdi-all)
    o = G.AppOptions(G.APP_SDI)
    o.resources = o.internet = True
    o.templates = list(G.TEMPLATES)
    r, pd, vbp = make("sdi-all", o)
    assert [m.name for m in r.modules] == ["Module1", "frmMain", "frmBrowser", "Dialog", "frmODBCLogon", "frmTip",
                                           "frmBrowser1"], [m.name for m in r.modules]
    caps = ["&File", "&New", "&Open...", "&Close", "&Save", "Save &As...", "Save A&ll", "Propert&ies",
            "Page Set&up...", "Print Pre&view", "&Print...", "Sen&d...", "E&xit", "&Edit", "&Undo", "Cu&t", "&Copy",
            "&Paste", "Paste &Special...", "&View", "&Toolbar", "Status &Bar", "&Refresh", "&Options...",
            "&Web Browser", "&Help", "&Contents", "&Search For Help On...", "&About ", "New", "Open", "Save", "Print",
            "Cut", "Copy", "Paste", "Bold", "Italic", "Underline", "Align Left", "Center", "Align Right", "Back",
            "Forward", "Stop", "Refresh", "Home", "Search", "&Address:"]
    assert r.strings == {1000 + i: c for i, c in enumerate(caps)}, r.strings
    from lbstudio.resmodel import ResFile
    assert ResFile.load(os.path.join(pd, "Project1.RES")).strings[0][1048] == "&Address:"
    lines = [x for x in vbp.splitlines() if x.split("=")[0] in ("Module", "Form", "Object", "Startup", "ResFile32")]
    assert lines == ["Module=Module1; Module1.bas", "Form=frmMain.frm",
                     "Object={F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0; ComDlg32.OCX",
                     "Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; mscomctl.OCX",
                     "Object={EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0; ieframe.dll", "Form=frmBrowser.frm",
                     "Form=Dialog.frm", "Form=frmODBCLogon.frm", "Form=frmTip.frm", "Form=frmBrowser1.frm",
                     'ResFile32="Project1.RES"', 'Startup="Sub Main"'], lines
    m = frm(pd, "frmMain")
    for want in ('      Caption         =   "1000"', "      Begin VB.Menu mnuFileMRU ", '      Shortcut        =   ^N',
                 '            Object.ToolTipText=   "1029"', "    LoadResStrings Me",
                 '    frmB.StartingAddress = "http://www.microsoft.com"', '            Picture         =   "frmMain.frx":'):
        assert want in m, (want, m[:3000])
    assert os.path.getsize(os.path.join(pd, "frmMain.frx")) > 13 * 800
    # the 13 ListImages of the ImageList: StdPicture with its CLSID, as VB6 writes the pictures of an ActiveX
    # control (AW32)
    from lbvc.project.frx import CLSID_STDPICTURE
    fx = open(os.path.join(pd, "frmMain.frx"), "rb").read()
    assert fx.count(CLSID_STDPICTURE) == 13, fx.count(CLSID_STDPICTURE)
    assert "Sub LoadResStrings(frm As Form)" in open(os.path.join(pd, "Module1.bas"), encoding="cp1252").read()
    b = frm(pd, "frmBrowser")
    assert '      Tag             =   "1048"' in b and "Public StartingAddress As String" in b, b[:2500]
    assert os.path.exists(os.path.join(pd, "frmBrowser1.frx")) and os.path.exists(os.path.join(pd, "frmTip.frx"))
    # 2. SDI, the standard forms (aw1 project-sdi-forms)
    o = G.AppOptions(G.APP_SDI)
    o.forms = list(G.STANDARD_FORMS)
    r, pd, vbp = make("sdi-forms", o)
    assert [m.name for m in r.modules] == ["Module1", "frmMain", "frmSplash", "frmLogin", "frmOptions", "frmAbout"]
    mod = open(os.path.join(pd, "Module1.bas"), encoding="cp1252").read()
    assert mod.index("frmLogin") < mod.index("frmSplash.Show") < mod.index("Unload frmSplash") \
        < mod.index("fMainForm.Show"), mod
    m = frm(pd, "frmMain")
    assert "    frmOptions.Show vbModal, Me" in m and "    frmAbout.Show vbModal, Me" in m
    assert "Begin MSComctlLib.TabStrip tbsOptions" in frm(pd, "frmOptions")
    assert os.path.exists(os.path.join(pd, "frmSplash.frx")) and os.path.exists(os.path.join(pd, "frmAbout.frx"))
    # 3. MDI (aw1 project-mdi)
    r, pd, vbp = make("mdi", G.AppOptions(G.APP_MDI))
    assert [m.name for m in r.modules] == ["Module1", "frmMain", "frmDocument"]
    assert "Object={3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0; richtx32.Ocx" in vbp
    m = frm(pd, "frmMain")
    assert "Begin VB.MDIForm frmMain" in m and "WindowList      =   -1  'True" in m and "    Me.Arrange vbCascade" in m
    assert "    LoadNewDoc" in m and "MDIChild        =   -1  'True" in frm(pd, "frmDocument")
    # 4. Explorer Style, Options and About (aw1 project-explorer)
    o = G.AppOptions(G.APP_EXPLORER)
    o.forms = ["Options", "About"]
    r, pd, vbp = make("explorer", o)
    m = frm(pd, "frmMain")
    for want in ("Begin MSComctlLib.TreeView tvTreeView", "Begin MSComctlLib.ListView lvListView",
                 "Begin VB.Image imgSplitter", "Begin VB.Menu mnuListViewMode", "Private Sub SizeControls(",
                 "    SetViewMode Index"):
        assert want in m, want
    assert 'Public Const LISTVIEW_MODE0 = "View Large Icons"' in open(os.path.join(pd, "Module1.bas"),
                                                                      encoding="cp1252").read()


def t_web92(d):
    """the WebBrowser (Microsoft Internet Controls, phase 9.2) as the VB6 IDE writes it (VB6 oracle DEV-9/aw3
    frmBrowser): Begin SHDocVwCtl.WebBrowser, ExtentX / ExtentY without "_" after Width, no _Version, its properties;
    a new size changes ExtentX / ExtentY; the Properties window has its properties (VB6 oracle DEV-9 wb3-01, 02)"""
    w = MainWindow()
    w.show()
    p = w.project
    f = w.current_designer()
    p.set_components(["SHDocVwCtl"])
    c = f.place_new("SHDocVwCtl.WebBrowser", QRect(8, 8, 385, 249), None)
    assert c is not None and c.name == "WebBrowser1", c
    for n in ("AddressBar", "FullScreen", "MenuBar", "Offline", "RegisterAsBrowser", "RegisterAsDropTarget", "Silent",
              "StatusBar", "TheaterMode", "ToolBar"):
        assert V.prop_def("WebBrowser", n) is not None, n
    f.set_rect(f.items[c], 120, 120, 5775, 3735)
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Web.vbp"))
    assert w.save_project_as()
    frm = open(os.path.join(d, "Form1.frm"), encoding="cp1252").read()
    assert 'Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"' in frm, frm[:300]
    head = "   Begin SHDocVwCtl.WebBrowser WebBrowser1"
    assert head in frm, frm
    blk = frm[frm.index(head):]
    got = [x.rstrip() for x in blk[:blk.index("\n   End") + len("\n   End")].splitlines()]
    exp = [head, "      Height          =   3735", "      Left            =   120", "      TabIndex        =   0",
           "      Top             =   120", "      Width           =   5775", "      ExtentX         =   10186",
           "      ExtentY         =   6588"] + WEB_BAG.splitlines()
    assert got == exp, "\n".join(f"{a!r} | {b!r}" for a, b in zip(got + [''] * len(exp), exp + [''] * len(got)) if a != b)
    assert "_ExtentX" not in blk[:blk.index("\n   End")] and "_Version" not in blk[:blk.index("\n   End")], blk[:600]
    vbp = open(os.path.join(d, "Web.vbp"), encoding="cp1252").read()
    assert "Object={EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0; ieframe.dll" in vbp, vbp
    done(w)


def t_objbrowser(d):
    """the Object Browser (phase 8.1): the libraries, classes and members, the details as VB6 shows them (VB6
    oracle DEV-8/ob1), Show Hidden Members, Search, the members of the project and View Definition, Definition
    (Shift+F2) on MsgBox, Go Back / Forward, Copy"""
    from lbstudio import refactorui as RF
    w = MainWindow()
    w.show()
    p = w.project
    w.add_module("Module")
    w.add_module("Class")
    code(w, p.module("Module1"), "Option Explicit\nPublic Const MAXV As Long = 10\nPublic Total As Long\n\n"
         "Public Enum Colors\n    cRed\n    cGreen = 5\n    cBlue\nEnd Enum\n\n"
         "Public Sub AddUp(ByVal n As Long)\n    Total = Total + n\nEnd Sub\n\n"
         "Private Function Twice(ByVal x As Long) As Long\n    Twice = x * 2\nEnd Function\n")
    code(w, p.module("Class1"), "Option Explicit\nPublic Event Changed(ByVal v As Long)\nPrivate mValue As Long\n\n"
         "Public Property Get Value() As Long\n    Value = mValue\nEnd Property\n\n"
         "Public Property Let Value(ByVal v As Long)\n    mValue = v\nEnd Property\n\n"
         "Public Property Get Count() As Long\n    Count = 1\nEnd Property\n")
    f = w.open_designer(p.module("Form1"))
    f.place_new("CommandButton", QRect(8, 8, 100, 30), None)
    ob = w.object_browser()
    libs = ob.library_names()
    assert libs[0] == "<All Libraries>" and {"Project1", "VB", "VBA", "VBRUN", "stdole"} <= set(libs), libs

    def first(*path):
        lib, cls, mem = path
        assert ob.set_library(lib) and ob.select_class(cls), (lib, cls, ob.class_names()[:20])
        assert ob.select_member(mem), (mem, ob.member_names()[:40])
        return ob.details_text().replace("\xa0", " ").splitlines()

    # the VB6 Object Browser (DEV-8/ob1)
    ob.set_library("VBA")
    cls = ob.class_names()
    assert cls[0] == "<globals>" and {"Strings", "Collection", "VbMsgBoxResult", "Interaction"} <= set(cls) and \
        "_HiddenModule" not in cls and "_Collection" not in cls, cls
    t = first("VBA", "Strings", "Left")
    assert t[0] == "Function Left(String, Length As Long)" and t[1].strip() == "Member of VBA.Strings", t
    assert "Left$" in ob.member_names()
    t = first("VBA", "Interaction", "MsgBox")
    assert t[0] == "Function MsgBox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], [Title], [HelpFile], [Context]) " \
                   "As VbMsgBoxResult", t
    t = first("VBA", "VbMsgBoxResult", "vbNo")
    assert t[0] == "Const vbNo = 7" and t[1].strip() == "Member of VBA.VbMsgBoxResult", t
    t = first("VBA", "Constants", "vbCrLf")
    assert t[0] == 'Const vbCrLf = "' and t[1] == '"', t
    t = first("VBA", "Collection", "Item")
    assert t[0] == "Function Item(Index)" and t[1].strip() == "Default member of VBA.Collection", t
    t = first("VB", "TextBox", "Move")
    assert t[0] == "Sub Move(Left As Single, [Top], [Width], [Height])" and t[1].strip() == "Member of VB.TextBox", t
    assert "Change" in ob.member_names() and "_Default" not in ob.member_names()
    t = first("VB", "Screen", "ActiveForm")
    assert t[0] == "Property ActiveForm As Form" and t[1].strip() == "read-only" and \
        t[2].strip() == "Member of VB.Screen", t
    ob.set_library("VBA")
    ob.select_class("<globals>")
    g = ob.member_names()
    assert {"vbOK", "vbOKCancel", "vbCrLf", "Left", "MsgBox", "Weekday"} <= set(g) and "VarPtr" not in g, g[:30]
    assert ob.details_text().replace("\xa0", " ").splitlines()[0] == "Library VBA"
    ob.set_show_hidden(True)
    assert "VarPtr" in ob.member_names()
    ob.set_show_hidden(False)
    # Search
    rows = ob.search("MsgBox")
    assert ("VBA", "Interaction", "MsgBox") in rows, rows[:5]
    rows = ob.search("vbYes*")
    assert ("VBA", "VbMsgBoxResult", "vbYes") in rows and ("VBA", "VbMsgBoxStyle", "vbYesNo") in rows, rows[:8]
    # the project
    ob.set_library("Project1")
    cl = ob.class_names()
    assert {"Module1", "Class1", "Form1", "Colors"} <= set(cl), cl
    ob.select_class("Module1")
    assert ob.member_names() == ["AddUp", "MAXV", "Total", "Twice"], ob.member_names()
    t = first("Project1", "Module1", "Twice")
    assert t[0] == "Function Twice(ByVal x As Long) As Long" and "Private" in [x.strip() for x in t], t
    t = first("Project1", "Module1", "MAXV")
    assert t[0] == "Const MAXV = 10", t
    t = first("Project1", "Colors", "cBlue")
    assert t[0] == "Const cBlue = 6" and t[1].strip() == "Member of Project1.Colors", t
    t = first("Project1", "Class1", "Count")
    assert t[0] == "Property Count As Long" and t[1].strip() == "read-only", t
    t = first("Project1", "Class1", "Value")
    assert t[0] == "Property Value As Long" and t[1].strip() == "Member of Project1.Class1", t
    t = first("Project1", "Form1", "Command1")
    assert t[0] == "Command1 As CommandButton", t
    assert "Caption" in ob.member_names()
    ob.select_class("<globals>")
    g = ob.member_names()
    assert {"AddUp", "MAXV", "Total", "cRed", "cGreen", "cBlue"} <= set(g) and "Twice" not in g, g
    # View Definition: the declaration in the code window
    first("Project1", "Module1", "AddUp")
    assert ob.view_definition()
    cw = w.current_cw()
    assert cw.module.name == "Module1" and cw.editor.sel_text() == "AddUp", (cw.module.name, cw.editor.sel_text())
    # Definition (Shift+F2) on MsgBox in the code: the Object Browser shows it
    code(w, p.module("Form1"), "Option Explicit\n\nPrivate Sub Form_Load()\n    MsgBox \"x\"\nEnd Sub\n")
    _caret(w, p.module("Form1"), "MsgBox")
    RF.definition(w)
    ob = w._objbrowser.widget()
    assert ob.current_class().name == "Interaction" and ob.current_member().name == "MsgBox", \
        (ob.current_class(), ob.current_member())
    assert ob.copy() == "MsgBox" and QApplication.clipboard().text() == "MsgBox"
    # Go Back / Go Forward
    ob.set_library("VB")
    ob.select_class("Screen")
    ob.back()
    assert ob.lib_combo.currentText() == "VB" and ob.current_class().name == "<globals>", ob.state()
    ob.back()
    assert ob.current_class().name == "Interaction" and ob.current_member().name == "MsgBox", ob.state()
    ob.forward()
    ob.forward()
    assert ob.lib_combo.currentText() == "VB" and ob.current_class().name == "Screen", ob.state()
    done(w)


DBG82 = """Option Explicit
Dim x As Long

Sub Main()
    Dim i As Long
    x = 1
    x = x + 10
    x = x + 100
    For i = 1 To 5
        x = x + 1
    Next
    Log CStr(x)
    On Error Resume Next
    x = 1 / 0
    On Error GoTo 0
    Log "handled"
    x = 1 / 0
    Log "after"
End Sub

Public Sub Log(ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "/out.txt" For Append As #f
    Print #f, s
    Close #f
End Sub
"""


def t_debug82(d):
    """phase 8.2 (the lines below are the lines of the code window, from 0)"""
    from lbstudio import debugwin
    w = MainWindow()
    w.show()
    p = w.project
    w.add_module("Module")
    m = p.module("Module1")
    code(w, m, DBG82)
    p.vbp.set("Startup", "Sub Main", quote=True)
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Module1.bas"), os.path.join(d, "Debug82.vbp"))
    assert w.save_project()
    breaks = []
    w.runner.broke.connect(lambda mod, ln, why, det: breaks.append((mod, ln - m.code_offset() - 1, why)))

    def settle(seconds=30):
        t0 = time.time()
        while time.time() - t0 < seconds:
            app.processEvents()
            time.sleep(0.02)
            if w.runner.state != "running" and not w.runner.pending and time.time() - t0 > 0.2:
                return
        raise AssertionError("no answer from the program")

    def wait_break(n, seconds=60):
        t0 = time.time()
        while time.time() - t0 < seconds:
            app.processEvents()
            time.sleep(0.02)
            if len(breaks) >= n and w.runner.state == "break":
                settle()
                return
            if w.runner.state == "design" and time.time() - t0 > 0.5:
                raise AssertionError(f"the program ended; breaks {breaks}; {w.output.toPlainText()}")
        raise AssertionError(f"no break {n}: {breaks}")

    def caret(line):
        ed = w.open_code(m).editor
        ed.goto_line(line)
        return ed

    caret(5)
    w.toggle_breakpoint()                       # x = 1
    caret(11)
    w.toggle_breakpoint()                       # Log CStr(x)
    w.run_start()
    wait_break(1)
    assert breaks[0] == ("Module1", 5, "bp"), breaks
    # Set Next Statement: a line of another procedure is refused; from x = 1 (not run yet) to x = x + 100
    caret(22)
    w.set_next_statement()
    settle()
    assert "Can't set the next statement" in w.output.toPlainText() and len(breaks) == 1, (breaks, w.output.toPlainText())
    caret(7)
    w.set_next_statement()
    wait_break(2)
    assert breaks[1][1:] == (7, "step") and w.break_at == (m, 7), (breaks, w.break_at)
    # Add Watch: the Context of the caret; Break When Value Is True
    seen = {}

    def fake_exec(dlg):
        seen["ctx"] = dlg.watch()
        dlg.expr.setText("x >= 104")
        dlg.set_kind(debugwin.WATCH_TRUE)
        return 1
    orig = debugwin.AddWatchDialog.exec
    debugwin.AddWatchDialog.exec = fake_exec
    caret(9)
    try:
        w.add_watch()
    finally:
        debugwin.AddWatchDialog.exec = orig
    assert seen["ctx"]["mod"] == "Module1" and seen["ctx"]["proc"] == "Main", seen
    assert [(x["expr"], x["kind"], x["mod"], x["proc"]) for x in w.watches] == [("x >= 104", 1, "Module1", "Main")], w.watches
    w.new_watch("x", debugwin.WATCH_EXPR, "Form1", "")          # out of its context in Main
    w.disable_breakpoints(True)                 # Log CStr(x) has a breakpoint: it does not stop there
    w.run_start()                               # Continue
    wait_break(3)
    assert breaks[2][2] == "watch", breaks
    assert w.watch_win.current() == 0, w.watch_win.current()
    t = w.watch_win.tree
    assert t.topLevelItem(0).text(1) == "True" and t.topLevelItem(1).text(1) == "<Out of context>", \
        [(t.topLevelItem(i).text(0), t.topLevelItem(i).text(1)) for i in range(t.topLevelItemCount())]
    assert t.topLevelItem(0).text(3) == "Module1.Main" and t.topLevelItem(1).text(3) == "Form1"
    assert t.topLevelItem(0).text(2) == "Boolean" and t.topLevelItem(1).text(2) == "Empty", \
        (t.topLevelItem(0).text(2), t.topLevelItem(1).text(2))            # the Type column (VB6 oracle dbg1)
    # Error Trapping, this session: Break on All Errors; the error of On Error Resume Next stops too
    w.set_trap(0)
    w.run_start()
    wait_break(4)
    assert breaks[3][1:] == (13, "error"), breaks          # the watch stays true: it does not stop again
    w.error_resume("errgo")                     # Alt+F5: On Error Resume Next takes the error
    wait_break(5)
    assert breaks[4][1:] == (16, "error"), breaks
    sub = w.toggle_menu(None)
    assert [a.isChecked() for a in sub.actions() if a.isCheckable()] == [True, False, False]
    # Edit and Continue: an edit in break mode resets the project (OK), then the edit is made
    oq = QMessageBox.question
    asked = []
    QMessageBox.question = lambda *a, **k: (asked.append(a[2] if len(a) > 2 else ""), QMessageBox.StandardButton.Ok)[1]
    try:
        ed = w.open_code(m).editor
        ed.SendScintilla(S.SCI_INSERTTEXT, ed.line_start(17), b"' ")
    finally:
        QMessageBox.question = oq
    wait(w)
    assert asked and "reset your project" in asked[0], asked
    assert w.runner.state == "design"
    lines = w.open_code(m).text().split("\n")
    assert lines[17].rstrip("\r") == "' " + '    Log "after"', lines[17]
    out = open(os.path.join(d, "out.txt")).read().split()
    assert out == ["105", "handled"], out           # x = 1 and x = x + 10 did not run: 100 + 5
    done(w)


ED1_SRC = "Sub Main()\n    Dim a As Integer\n    a = 1\n    If a = 1 Then\n        a = 2\n    End If\nEnd Sub\n"
# what the VB6 IDE wrote at the end of every stage of DEV-8/ed1 (the code, without Attribute VB_Name)
ED1_ST1 = "Sub Main()\n'    Dim a As Integer\n'    a = 1\n'    If a = 1 Then\n'        a = 2\n'    End If\nEnd Sub\n"
ED1_ST2 = "Sub Main()\n    b = 3\n'    Dim a As Integer\n'    a = 1\n'    If a = 1 Then\n        a = 2\n    End If\nEnd Sub\n"
ED1_ST3 = ED1_ST2 + ("\nPublic Sub DoIt()\n\nEnd Sub\n\nPrivate Function Calc()\n\nEnd Function\n\n"
                     "Public Property Get Value() As Variant\n\nEnd Property\n\n"
                     "Public Property Let Value(ByVal vNewValue As Variant)\n\nEnd Property\n")
ED1_ST4 = "\nPublic Event Ev()\n" + ED1_ST3 + "\nPublic Static Sub Cnt()\n\nEnd Sub\n"
ED1_CLS = ("Option Explicit\nDim m As Long\n\nPublic Event Changed()\n\nSub Foo()\nEnd Sub\n\n\n"
           "Private Property Get Title() As Variant\n\nEnd Property\n\n"
           "Private Property Let Title(ByVal vNewValue As Variant)\n\nEnd Property\n")


def t_editor83(d):
    from lbstudio.dialogs import AddProcedureDialog
    w = MainWindow()
    w.show()
    p = w.project
    w.add_module("Module")
    m = p.module("Module1")
    cw = w.open_code(m)
    ed = cw.editor
    ed.setText(ED1_SRC)

    def txt():
        return cw.text().replace("\r\n", "\n")

    def lines(a, b):                            # select from the start of line a to the end of line b
        ed.set_sel(ed.line_start(a), ed.line_end(b))

    def caret(line, col=6):
        ed.goto_pos(ed.line_start(line) + col)

    # stage 1: Comment Block on 4 selected lines, then on the caret's line
    lines(2, 5)
    w.comment_block()
    caret(1)
    w.comment_block()
    assert txt() == ED1_ST1, repr(txt())
    # a selection that ends at the start of a line leaves that line out
    ed.set_sel(ed.line_start(0), ed.line_start(1))
    w.comment_block()
    assert txt().startswith("'Sub Main()\n'    Dim"), repr(txt())
    ed.undo()
    assert txt() == ED1_ST1, repr(txt())        # one undo for the block
    # stage 2: Uncomment Block
    caret(3, 10)
    w.comment_block()                           # '' If
    caret(3, 10)
    w.comment_block(True)
    lines(4, 5)
    w.comment_block(True)
    ed.goto_pos(ed.line_end(0))
    ed.SendScintilla(S.SCI_INSERTTEXT, ed.line_end(0), b"\n    'b = 3")
    caret(1, 7)
    w.comment_block(True)
    caret(0, 3)
    w.comment_block(True)                       # no comment: nothing
    assert txt() == ED1_ST2, repr(txt())
    # stage 3: Add Procedure at the end of the module, wherever the caret is
    for name, kind, scope in (("DoIt", "Sub", "Public"), ("Calc", "Function", "Private"),
                              ("Value", "Property", "Public")):
        caret(0, 6)
        assert w.apply_add_procedure(cw, name, kind, scope, False) is None
    assert txt() == ED1_ST3, repr(txt())
    assert ed.line_from_pos(ed.pos()) == 18, ed.line_from_pos(ed.pos())     # the body of Property Get Value
    ed.undo()
    assert txt() == ED1_ST3[:ED1_ST3.index("\nPublic Property Get")], repr(txt())   # one undo for a procedure
    ed.redo()
    assert txt() == ED1_ST3, repr(txt())
    # stage 4: Statics; the messages (through the dialog); an Event in the declarations
    msgs = []
    oc = QMessageBox.critical
    QMessageBox.critical = lambda *a, **k: msgs.append(a[2])
    try:
        dlg = AddProcedureDialog(w, lambda *a: w.apply_add_procedure(cw, *a))
        dlg.set_values("Cnt", "Sub", "Public", True)
        dlg._ok()
        assert dlg.result() == 1 and not msgs, msgs
        dlg = AddProcedureDialog(w, lambda *a: w.apply_add_procedure(cw, *a))
        dlg.show()
        dlg.set_values("1abc")
        dlg._ok()
        assert msgs == ["Invalid procedure name"] and dlg.isVisible(), msgs      # the dialog stays
        dlg.set_values("DoIt")
        dlg._ok()
        assert msgs[-1] == "A procedure of that name already exists" and not dlg.isVisible(), msgs
    finally:
        QMessageBox.critical = oc
    caret(0, 6)
    assert w.apply_add_procedure(cw, "Ev", "Event", "Public", False) is None
    assert txt() == ED1_ST4, repr(txt())
    # stage 5: a class: the Event after the declarations, a Private Property
    w.add_module("Class")
    c1 = p.module("Class1")
    ccw = w.open_code(c1)
    ccw.editor.setText("Option Explicit\nDim m As Long\n\nSub Foo()\nEnd Sub\n\n")
    ccw.editor.goto_pos(ccw.editor.line_start(3) + 5)
    assert w.apply_add_procedure(ccw, "Changed", "Event", "Public", False) is None
    assert w.apply_add_procedure(ccw, "Title", "Property", "Private", False) is None
    assert ccw.text().replace("\r\n", "\n") == ED1_CLS, repr(ccw.text())
    # Bookmarks: Next / Previous through the modules, round
    cw = w.open_code(m)
    ed = cw.editor
    caret(2, 0)
    w.toggle_code_bookmark()
    caret(9, 0)
    w.toggle_code_bookmark()
    ccw = w.open_code(c1)
    ccw.editor.goto_line(1)
    w.toggle_code_bookmark()
    assert w._bookmarks(m) == [2, 9] and w._bookmarks(c1) == [1], (w._bookmarks(m), w._bookmarks(c1))
    cw = w.open_code(m)
    cw.editor.goto_line(0)
    seen = []
    for _ in range(4):
        w.next_bookmark(True)
        c = w.current_cw()
        seen.append((c.module.name, c.editor.line_from_pos(c.editor.pos())))
    assert seen == [("Module1", 2), ("Module1", 9), ("Class1", 1), ("Module1", 2)], seen
    w.next_bookmark(False)
    c = w.current_cw()
    assert (c.module.name, c.editor.line_from_pos(c.editor.pos())) == ("Class1", 1)
    # a code window closes: its bookmarks stay with the module
    w.code_windows[c1].close()
    app.processEvents()
    assert w._bookmarks(c1) == [1]
    ccw = w.open_code(c1)
    assert ccw.editor.SendScintilla(S.SCI_MARKERGET, 1) & (1 << 24)
    w.clear_bookmarks()
    assert not w._bookmarks(m) and not w._bookmarks(c1)
    # the Edit toolbar (View > Toolbars > Edit) has Comment Block and Uncomment Block
    acts = w.tb_edit.actions()
    assert w.a_comment in acts and w.a_uncomment in acts and w.a_bm_next in acts and w.tb_edit.isHidden()
    # Save: the .bas as VB6 wrote it
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Module1.bas"), os.path.join(d, "Class1.cls"),
          os.path.join(d, "Ed83.vbp"))
    assert w.save_project()
    data = open(os.path.join(d, "Module1.bas"), newline="").read()
    assert data == 'Attribute VB_Name = "Module1"\r\n' + ED1_ST4.replace("\n", "\r\n"), repr(data)
    done(w)


ED2_SRC = ("Option Explicit\n\nPublic Sub Run()\nEnd Sub\n\nPublic Sub Hidden()\nEnd Sub\n\n"
           "Public Property Get Caption() As String\nEnd Property\n\nPublic Property Let Caption(ByVal v As String)\n"
           "End Property\n\nPublic Property Get Value() As Long\nEnd Property\n\n"
           "Public Property Let Value(ByVal n As Long)\nEnd Property\n")
# what the VB6 IDE wrote (DEV-8/ed2, stage 4)
ED2_ST4 = ("Option Explicit\n\nPublic Sub Run()\nAttribute Run.VB_UserMemId = 0\nEnd Sub\n\nPublic Sub Hidden()\n"
           "Attribute Hidden.VB_MemberFlags = \"40\"\nEnd Sub\n\nPublic Property Get Caption() As String\n"
           "Attribute Caption.VB_Description = \"Title Text\"\nAttribute Caption.VB_HelpID = 5\n"
           "Attribute Caption.VB_ProcData.VB_Invoke_Property = \";Appearance\"\nAttribute Caption.VB_UserMemId = -518\n"
           "Attribute Caption.VB_MemberFlags = \"200\"\nEnd Property\n\n"
           "Public Property Let Caption(ByVal v As String)\nEnd Property\n\nPublic Property Get Value() As Long\n"
           "Attribute Value.VB_MemberFlags = \"143c\"\nEnd Property\n\nPublic Property Let Value(ByVal n As Long)\n"
           "End Property\n")


def t_procattr83(d):
    w = MainWindow()
    w.show()
    p = w.project
    w.add_module("Class")
    c1 = p.module("Class1")
    cw = w.open_code(c1)
    ed = cw.editor
    ed.setText(ED2_SRC)
    ed.goto_pos(ed.line_start(2) + 12)          # in "Public Sub Run()"
    dlg = w.procedure_attributes_dialog(cw)
    assert dlg.cur == "Run" and [dlg.name.itemText(i) for i in range(dlg.name.count())] == \
        ["Run", "Hidden", "Caption", "Value"]
    dlg.show_advanced()
    assert not dlg.adv_btn.isEnabled()
    # a Sub: Procedure ID (None, Default, AboutBox, Refresh) and Hide this member; the property options are off
    assert [dlg.procid.itemText(i) for i in range(dlg.procid.count())] == ["(None)", "(Default)", "AboutBox", "Refresh"]
    assert not dlg.category.isEnabled() and not dlg.flag_boxes["uidefault"].isEnabled() and \
        dlg.flag_boxes["hide"].isEnabled() and not dlg.binding_group.isEnabled()
    dlg.procid.setCurrentIndex(1)               # (Default)
    assert dlg.apply_btn.isEnabled()
    dlg._apply()
    assert "Attribute Run.VB_UserMemId = 0" in cw.text()
    ed.undo()                                   # one undo for Apply
    assert "Attribute Run" not in cw.text()
    ed.redo()
    dlg.name.setCurrentIndex(1)                 # Hidden
    dlg.flag_boxes["hide"].setChecked(True)
    dlg.name.setCurrentIndex(2)                 # Caption: Hidden is applied on the way
    assert dlg.cur == "Caption" and dlg.category.isEnabled() and dlg.binding_group.isEnabled()
    assert dlg.procid.count() == 22 and dlg.procid.itemText(9) == "Caption"
    assert not dlg.flag_boxes["defaultbind"].isEnabled()      # until Property is data bound
    dlg.desc.setPlainText("Title Text")
    dlg.helpid.setText("5")
    dlg.procid.setCurrentIndex(9)
    dlg.category.setCurrentIndex(dlg.category.findText("Appearance"))
    dlg.flag_boxes["uidefault"].setChecked(True)
    dlg.name.setCurrentIndex(3)                 # Value
    for k in ("nobrowse", "bindable", "defaultbind", "displaybind", "requestedit", "immediatebind"):
        dlg.flag_boxes[k].setChecked(True)
    dlg._ok()
    got = cw.text().replace("\r\n", "\n")
    assert got == ED2_ST4, repr(got)
    # opened again: the values come back from the Attribute lines
    ed.goto_pos(ed.line_start(12) + 5)
    dlg = w.procedure_attributes_dialog(cw)
    assert dlg.cur == "Caption", dlg.cur
    assert dlg.desc.toPlainText() == "Title Text" and dlg.helpid.text() == "5" and \
        dlg.procid.currentText() == "Caption" and dlg.category.currentText() == "Appearance" and \
        dlg.flag_boxes["uidefault"].isChecked() and not dlg.flag_boxes["hide"].isChecked()
    dlg.reject()
    # a standard module: only Procedure ID and Hide this member, as in VB6 (DEV-8/ed2 stage 1)
    w.add_module("Module")
    m1 = p.module("Module1")
    mcw = w.open_code(m1)
    mcw.editor.setText("Option Explicit\n\nPublic Sub Run()\nEnd Sub\n\nPublic Property Get Caption() As String\n"
                       "End Property\n\nPublic Property Let Caption(ByVal v As String)\nEnd Property\n")
    dlg = w.procedure_attributes_dialog(mcw)
    dlg.name.setCurrentIndex(1)                 # a Property in a module
    assert not dlg.category.isEnabled() and not dlg.flag_boxes["nobrowse"].isEnabled()
    dlg.name.setCurrentIndex(0)
    dlg.desc.setPlainText("Runs It")
    dlg.helpid.setText("12")
    dlg._ok()
    assert mcw.text().replace("\r\n", "\n").startswith(
        "Option Explicit\n\nPublic Sub Run()\nAttribute Run.VB_Description = \"Runs It\"\nAttribute Run.VB_HelpID = 12\n"
        "End Sub\n"), repr(mcw.text())
    # Refresh and AboutBox (DEV-8/ed2 stage 6, Class6.cls)
    w.add_module("Class")
    c2 = p.module("Class2")
    c2cw = w.open_code(c2)
    c2cw.editor.setText("Public Sub Run()\nEnd Sub\n\nPublic Sub About()\nEnd Sub\n")
    c2cw.editor.goto_pos(c2cw.editor.length())   # the caret at the end: the dialog opens on About (as VB6)
    dlg = w.procedure_attributes_dialog(c2cw)
    assert dlg.cur == "About", dlg.cur
    dlg.name.setCurrentIndex(0)                 # Run
    assert dlg.cur == "Run"
    dlg.procid.setCurrentIndex(dlg.procid.findText("Refresh"))
    dlg._apply()
    dlg.name.setCurrentIndex(1)
    dlg.procid.setCurrentIndex(dlg.procid.findText("AboutBox"))
    dlg._ok()
    assert c2cw.text().replace("\r\n", "\n") == ("Public Sub Run()\nAttribute Run.VB_UserMemId = -550\nEnd Sub\n\n"
                                                 "Public Sub About()\nAttribute About.VB_UserMemId = -552\nEnd Sub\n"), \
        repr(c2cw.text())
    # Save: the Attribute lines in the .cls as VB6 writes them
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Class1.cls"), os.path.join(d, "Module1.bas"),
          os.path.join(d, "Class2.cls"), os.path.join(d, "Pa83.vbp"))
    assert w.save_project()
    data = open(os.path.join(d, "Class1.cls"), newline="").read()
    assert data.endswith("Attribute VB_Exposed = False\r\n" + ED2_ST4.replace("\n", "\r\n")), repr(data[-400:])
    done(w)


PR1_MODULE = ("Sub Main()\n    MsgBox \"Hello\"\nEnd Sub\n\nSub LongOne()\n    Dim x As Integer\n" +
              "".join(f"    x = {i} ' line {i:03d}\n" for i in range(1, 101)) + "End Sub\n")


def t_print83(d):
    from PyQt6.QtCore import QMarginsF
    from PyQt6.QtGui import QPageLayout, QPageSize
    from PyQt6.QtPrintSupport import QPrinter
    from lbstudio import printing
    from lbstudio.dialogs import PrintDialog
    w = MainWindow()
    w.show()
    p = w.project
    f1 = p.module("Form1")
    w.open_code(f1).editor.setText("Private Sub Command1_Click()\n    MsgBox \"Hi\"\nEnd Sub\n")
    w.add_module("Module")
    m1 = p.module("Module1")
    cw = w.open_code(m1)
    cw.editor.setText(PR1_MODULE)
    # Current Project, Code (DEV-8/pr1 project2.xps): Form1, then Module1 on two pages
    pages = w.print_layout("project", {"code"}, 79)
    assert [k for k, _ in pages] == ["text"] * 3, pages
    assert pages[0][1] == ["Form1 - 1", "", "Private Sub Command1_Click()", '    MsgBox "Hi"', "End Sub"], pages[0]
    p2, p3 = pages[1][1], pages[2][1]
    assert len(p2) == 79 and p2[:5] == ["Module1 - 1", "", "Sub Main()", '    MsgBox "Hello"', "End Sub"] and \
        p2[-1] == "    x = 71 ' line 071", (len(p2), p2[:5], p2[-1])
    assert p3[:3] == ["Module1 - 2", "", "    x = 72 ' line 072"] and p3[-1] == "End Sub" and len(p3) == 32, p3
    # Selection (selection.xps): the header too
    cw.editor.set_sel(cw.editor.line_start(0), cw.editor.line_start(3))
    assert w.print_layout("selection", {"code"}, 79) == \
        [("text", ["Module1 - 1", "", "Sub Main()", '    MsgBox "Hello"', "End Sub"])]
    # a form, Current Module (project.xps): its code, Form As Text, Form Image, each on its own pages
    w.open_designer(f1)
    pages = w.print_layout("module", {"code", "text", "image"}, 79)
    assert [k for k, _ in pages] == ["text", "text", "image"], pages
    ft = pages[1][1]
    assert ft[:4] == ["Form1 - 1", "", "VERSION 5.00", "Begin VB.Form Form1 "] and ft[-1] == "End", ft
    assert not any(x.startswith("Attribute") for x in ft)
    pm = w.form_image(f1)
    assert pm is not None and not pm.isNull() and pm.width() > 50, pm
    # the Print dialog: Selection when there is one; Form Image / Form As Text for a form or Current Project
    dlg = PrintDialog(w, "Test", has_selection=True, form_active=False)
    assert dlg.r_sel.isChecked() and not dlg.c_image.isEnabled() and dlg.c_code.isChecked()
    dlg.r_proj.setChecked(True)
    assert dlg.c_image.isEnabled() and dlg.c_text.isEnabled()
    dlg.c_text.setChecked(True)
    assert dlg.values() == {"range": "project", "what": {"code", "text"}, "quality": "High", "to_file": False}
    dlg = PrintDialog(w, "Test", has_selection=False, form_active=True)
    assert not dlg.r_sel.isEnabled() and dlg.r_mod.isChecked() and dlg.c_image.isEnabled()
    # printed: a PDF on A4 without margins at 600 dpi (as the XPS writer): 79 lines a page, 5 pages
    pr = QPrinter(QPrinter.PrinterMode.HighResolution)
    pr.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
    out = os.path.join(d, "project.pdf")
    pr.setOutputFileName(out)
    pr.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
    pr.setPageMargins(QMarginsF(0, 0, 0, 0), QPageLayout.Unit.Millimeter)
    pr.setResolution(600)
    assert printing.lines_per_page(pr) == 79, printing.lines_per_page(pr)
    assert printing.metrics(600) == (83, 88, 65)
    n = w.do_print("project", {"code", "text", "image"}, printer=pr)
    assert n == 5, n
    import re
    data = open(out, "rb").read()
    assert len(re.findall(rb"/Type\s*/Page[^s]", data)) == 5, len(re.findall(rb"/Type\s*/Page[^s]", data))
    done(w)



# the .vbp files the VB6 IDE wrote in DEV-8/pp1 (Project Properties, phase 8.5)
PP1 = {
    'base': 'Type=Exe\nForm=Form1.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="Form1"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'a': 'Type=Exe\nForm=F6.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="Form1"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'b': 'Type=Exe\nForm=F6.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="Form1"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'c': 'Type=Exe\nForm=F6.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nIconForm="Form1"\nStartup="Form1"\nHelpFile=""\nTitle="b"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'generall': 'Type=Exe\nForm=Form1.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="Form1"\nHelpFile="my.hlp"\nNoControlUpgrade=1\nCommand32=""\nName="MyProj"\nHelpContextID="7"\nDescription="Desc text"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'make': 'Type=Exe\nForm=Form1.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nIconForm="Form1"\nStartup="Form1"\nHelpFile="my.hlp"\nNoControlUpgrade=1\nTitle="My App"\nCommand32="-x 1"\nName="MyProj"\nHelpContextID="7"\nDescription="Desc text"\nCompatibleMode="0"\nMajorVer=2\nMinorVer=3\nRevisionVer=4\nAutoIncrementVer=1\nServerSupportFiles=0\nVersionComments="c1"\nVersionCompanyName="co1"\nVersionFileDescription="fd1"\nVersionLegalCopyright="lc1"\nVersionLegalTrademarks="lt1"\nVersionProductName="pn1"\nCondComp="DBG = 1"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nRemoveUnusedControlInfo=0\n',
    'compile': 'Type=Exe\nForm=Form1.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nIconForm="Form1"\nStartup="Form1"\nHelpFile="my.hlp"\nNoControlUpgrade=1\nTitle="My App"\nCommand32="-x 1"\nName="MyProj"\nHelpContextID="7"\nDescription="Desc text"\nCompatibleMode="0"\nMajorVer=2\nMinorVer=3\nRevisionVer=4\nAutoIncrementVer=1\nServerSupportFiles=0\nVersionComments="c1"\nVersionCompanyName="co1"\nVersionFileDescription="fd1"\nVersionLegalCopyright="lc1"\nVersionLegalTrademarks="lt1"\nVersionProductName="pn1"\nCondComp="DBG = 1"\nCompilationType=0\nOptimizationType=1\nFavorPentiumPro(tm)=-1\nCodeViewDebugInfo=-1\nNoAliasing=-1\nBoundsCheck=-1\nOverflowCheck=-1\nFlPointCheck=-1\nFDIVCheck=-1\nUnroundedFP=-1\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nRemoveUnusedControlInfo=0\n',
    'pcode': 'Type=Exe\nForm=Form1.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nIconForm="Form1"\nStartup="Form1"\nHelpFile="my.hlp"\nNoControlUpgrade=1\nTitle="My App"\nCommand32="-x 1"\nName="MyProj"\nHelpContextID="7"\nDescription="Desc text"\nCompatibleMode="0"\nMajorVer=2\nMinorVer=3\nRevisionVer=4\nAutoIncrementVer=1\nServerSupportFiles=0\nVersionComments="c1"\nVersionCompanyName="co1"\nVersionFileDescription="fd1"\nVersionLegalCopyright="lc1"\nVersionLegalTrademarks="lt1"\nVersionProductName="pn1"\nCondComp="DBG = 1"\nCompilationType=-1\nOptimizationType=1\nFavorPentiumPro(tm)=-1\nCodeViewDebugInfo=-1\nNoAliasing=-1\nBoundsCheck=-1\nOverflowCheck=-1\nFlPointCheck=-1\nFDIVCheck=-1\nUnroundedFP=-1\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nRemoveUnusedControlInfo=0\n',
    'ax1': 'Type=OleExe\nClass=Class1; C7.cls\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'ax2': 'Type=OleExe\nClass=Class1; C7.cls\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=1\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=-1\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=3\n',
    'ax3': 'Type=OleExe\nClass=Class1; C13.cls\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'ax4': 'Type=OleExe\nClass=Class1; C13.cls\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=-1\nMaxNumberOfThreads=1\n',
    'dll1': 'Type=OleDll\nClass=Class1; C8.cls\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=1\n',
    'dll2': 'Type=OleDll\nClass=Class1; C8.cls\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nDllBaseAddress=&H12340000\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=-1\nRetained=1\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=0\n',
    'dbg1': 'Type=OleDll\nClass=Class1; C8.cls\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nDllBaseAddress=&H12340000\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=-1\nRetained=1\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=0\nDebugStartupOption=2\nDebugStartupComponent=C:\\Windows\\notepad.exe\n',
    'dbg2': 'Type=OleDll\nClass=Class1; C8.cls\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nDllBaseAddress=&H12340000\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=-1\nRetained=1\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=0\nDebugStartupOption=3\nDebugStartupComponent=http://localhost/\nUseExistingBrowser=0\n',
    'dbg3': 'Type=OleDll\nClass=Class1; C8.cls\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nDllBaseAddress=&H12340000\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=-1\nRetained=1\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=0\nDebugStartupOption=0\nUseExistingBrowser=0\n',
    'ocx1': 'Type=Control\nUserControl=U10.ctl\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=1\n',
    'ocx2': 'Type=Control\nUserControl=U10.ctl\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nRequireLicenseKey=1\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=1\nDebugStartupOption=1\nDebugStartupComponent=UserControl1\n',
    'ocx3': 'Type=Control\nUserControl=U10.ctl\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nRequireLicenseKey=1\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=1\nDebugStartupOption=1\nDebugStartupComponent=UserControl1\n',
    'ocx4': 'Type=Control\nUserControl=U10.ctl\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nRequireLicenseKey=1\nNoControlUpgrade=1\nTitle="ocx3"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=1\nRemoveUnusedControlInfo=0\nDebugStartupOption=1\nDebugStartupComponent=UserControl1\n',
    'ocx5': 'Type=Control\nUserControl=U10.ctl\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nRequireLicenseKey=1\nNoControlUpgrade=1\nTitle="ocx3"\nExeName32="ocx4.ocx"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nCompatibleEXE32="ocx4.ocx"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=1\nRemoveUnusedControlInfo=0\nDebugStartupOption=1\nDebugStartupComponent=UserControl1\n',
    'ocx6': 'Type=Control\nUserControl=U10.ctl\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nRequireLicenseKey=1\nNoControlUpgrade=1\nTitle="ocx3"\nExeName32="ocx4.ocx"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="2"\nCompatibleEXE32="ocx4.ocx"\nVersionCompatible32="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=1\nRemoveUnusedControlInfo=0\nDebugStartupOption=1\nDebugStartupComponent=UserControl1\n',
    'ocx7': 'Type=Control\nUserControl=U10.ctl\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nRequireLicenseKey=1\nNoControlUpgrade=1\nTitle="ocx3"\nExeName32="ocx4.ocx"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nCompatibleEXE32="ocx4.ocx"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=1\nRemoveUnusedControlInfo=0\nDebugStartupOption=1\nDebugStartupComponent=UserControl1\n',
    'ocx8': 'Type=Control\nUserControl=U10.ctl\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="(None)"\nHelpFile=""\nRequireLicenseKey=1\nNoControlUpgrade=1\nTitle="ocx3"\nExeName32="ocx4.ocx"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nCompatibleEXE32="ocx4.ocx"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nThreadingModel=1\nRemoveUnusedControlInfo=0\nDebugStartupOption=1\nDebugStartupComponent=UserControl1\n',
    'ex1': 'Type=Exe\nForm=F14.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="Form1"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'ex2': 'Type=Exe\nForm=F14.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="Form1"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\nDebugStartupOption=0\n',
    'ex3': 'Type=OleDll\nForm=F14.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nIconForm="Form1"\nStartup="Sub Main"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'ex5': 'Type=Exe\nForm=F17.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="Form1"\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'ex6': 'Type=Exe\nForm=F17.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nStartup="Sub Main"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'ex7': 'Type=Exe\nForm=F14.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nIconForm="Form1"\nStartup="Sub Main"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="0"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=0\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'ex8': 'Type=OleExe\nForm=F17.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nIconForm="Form1"\nStartup="Sub Main"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
    'ex9': 'Type=Control\nForm=F17.frm\nReference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation\nIconForm="Form1"\nStartup="Sub Main"\nHelpFile=""\nCommand32=""\nName="Project1"\nHelpContextID="0"\nCompatibleMode="1"\nMajorVer=1\nMinorVer=0\nRevisionVer=0\nAutoIncrementVer=0\nServerSupportFiles=0\nVersionCompanyName="home"\nCompilationType=0\nOptimizationType=0\nFavorPentiumPro(tm)=0\nCodeViewDebugInfo=0\nNoAliasing=0\nBoundsCheck=0\nOverflowCheck=0\nFlPointCheck=0\nFDIVCheck=0\nUnroundedFP=0\nStartMode=1\nUnattended=0\nRetained=0\nThreadPerObject=0\nMaxNumberOfThreads=1\n',
}


def _pp1_project(d, name, text=None):
    """the project of a .vbp of DEV-8/pp1 in a folder of its own (its form, class, user control made as VB6 makes
    them), opened as the IDE opens it"""
    from lbstudio import vbmodel
    folder = os.path.join(d, name)
    os.makedirs(folder, exist_ok=True)
    text = PP1[name] if text is None else text
    for line in text.split("\n"):
        k, _, v = line.partition("=")
        if k == "Form":
            src = vbmodel.NEW_FORM.format(name="Form1")
        elif k == "Class":
            v = v.split(";")[1].strip()
            src = vbmodel.NEW_CLASS.format(name="Class1", exposed="True")
        elif k == "UserControl":
            src = vbmodel.NEW_UCTL.format(name="UserControl1", exposed="True")
        else:
            continue
        with open(os.path.join(folder, v), "w", encoding="cp1252", newline="\r\n") as f:
            f.write(src)
    path = os.path.join(folder, name + ".vbp")
    with open(path, "w", encoding="cp1252", newline="\r\n") as f:
        f.write(text)
    return vbmodel.VBProject.open(path, "cp1252")


def _pp1_dialog(w, p, *pages):
    from lbstudio.dialogs import ProjectPropertiesDialog
    dlg = ProjectPropertiesDialog(w, p)
    assert dlg.tabs.currentIndex() == 0                 # it opens on General (DEV-8/pp1 stage 14)
    for page in pages:                                  # the pages shown: what VB6 writes on OK follows them
        dlg.tabs.setCurrentIndex(dlg.PAGES.index(page))
    return dlg


def _pp1_check(p, want):
    got = p.vbp.dumps().replace("\r\n", "\n")
    assert got == PP1[want], f"{want}:\n" + "\n".join(
        x for x in __import__("difflib").unified_diff(PP1[want].split("\n"), got.split("\n"), lineterm=""))


def t_projprops85(d):
    from lbstudio import projprops, vbmodel
    w = MainWindow()
    w.show()
    crit = QMessageBox.critical
    QMessageBox.critical = lambda *a, **k: None
    try:
        # General: OK alone writes HelpFile="" (a -> b); the name, the help file, Upgrade ActiveX Controls (generall)
        p = _pp1_project(d, "base")
        dlg = _pp1_dialog(w, p)
        assert [dlg.startup.itemText(i) for i in range(dlg.startup.count())] == ["Sub Main", "Form1"]
        assert not dlg.unattended.isEnabled() and not dlg.license.isEnabled() and not dlg.retained.isEnabled() and \
            not dlg.threading.isEnabled() and dlg.threading.currentIndex() == -1 and not dlg.per_object.isEnabled() \
            and dlg.pool.isChecked() and dlg.upgrade.isChecked()
        assert not dlg.compat_box.isEnabled() and not dlg.standalone.isEnabled() and not dlg.debug_box.isEnabled() \
            and not dlg.dllbase.isEnabled() and dlg.native.isChecked() and dlg.dllbase.text() == "&H11000000"
        assert dlg._ok() and not dlg.messages
        assert p.vbp.dumps().replace("\r\n", "\n") == PP1["base"].replace('Startup="Form1"\n',
                                                                          'Startup="Form1"\nHelpFile=""\n')
        p = _pp1_project(d, "base")
        dlg = _pp1_dialog(w, p)
        dlg.pname.setText("MyProj")
        dlg.helpfile.setText("my.hlp")
        dlg.helpctx.setText("7")
        dlg.desc.setText("Desc text")
        dlg.upgrade.setChecked(False)
        dlg._ok()
        _pp1_check(p, "generall")
        # Make: shown, it writes IconForm and Title (the name of the .vbp): b -> c
        p = _pp1_project(d, "b")
        dlg = _pp1_dialog(w, p, "make")
        assert dlg.title.text() == "b" and dlg.icon.currentText() == "Form1" and not dlg.icon_view.pixmap().isNull()
        dlg._ok()
        _pp1_check(p, "c")
        # Make: the version, the Version Information, the arguments (generall -> make)
        p = _pp1_project(d, "generall")
        dlg = _pp1_dialog(w, p, "make")
        dlg.major.setValue(2)
        dlg.minor.setValue(3)
        dlg.revision.setValue(4)
        dlg.autoinc.setChecked(True)
        dlg.title.setText("My App")
        for row, text in enumerate(("c1", "co1", "fd1", "lc1", "lt1", "pn1")):
            dlg.info_list.setCurrentRow(row)
            dlg.info_value.setPlainText(text)
        dlg.info_list.setCurrentRow(1)
        assert dlg.info_value.toPlainText() == "co1"
        dlg.command.setText("-x 1")
        dlg.condcomp.setText("DBG = 1")
        dlg.remove_unused.setChecked(False)
        dlg._ok()
        _pp1_check(p, "make")
        # Compile: Small Code, Pentium Pro, Symbolic Debug Info, the six Advanced Optimizations; then P-Code
        p = _pp1_project(d, "make")
        dlg = _pp1_dialog(w, p, "compile")
        dlg.small.setChecked(True)
        dlg.pentium.setChecked(True)
        dlg.symdebug.setChecked(True)
        from lbstudio.dialogs import AdvancedOptimizationsDialog
        adv = AdvancedOptimizationsDialog(dlg, dlg.s["advanced"])
        assert [cb.text() for cb in adv.boxes.values()] == [t for t, _ in projprops.ADVANCED]
        for cb in adv.boxes.values():
            cb.setChecked(True)
        dlg.s["advanced"] = adv.values()
        dlg._ok()
        _pp1_check(p, "compile")
        p = _pp1_project(d, "compile")
        dlg = _pp1_dialog(w, p, "compile")
        dlg.pcode.setChecked(True)
        assert not dlg.native_group.isEnabled()
        dlg._ok()
        _pp1_check(p, "pcode")
        # an ActiveX EXE: Unattended Execution, Thread Pool of 3, Remote Server Files, No Compatibility (ax1 -> ax2)
        p = _pp1_project(d, "ax1")
        dlg = _pp1_dialog(w, p, "component")
        assert dlg.unattended.isEnabled() and not dlg.license.isEnabled() and not dlg.retained.isEnabled() and \
            not dlg.threading.isEnabled() and dlg.per_object.isEnabled() and dlg.threads.isEnabled()
        assert [dlg.startup.itemText(i) for i in range(dlg.startup.count())] == ["(None)", "Sub Main"]
        assert dlg.activex.isChecked() and dlg.standalone.isEnabled() and dlg.proj_compat.isChecked()
        assert dlg.debug_box.isEnabled() and not dlg.dbg_comp.isEnabled() and dlg.dbg_wait.isChecked()
        dlg.unattended.setChecked(True)
        dlg.threads.setValue(3)
        dlg.remote.setChecked(True)
        dlg.no_compat.setChecked(True)
        assert not dlg.compat_path.isEnabled()
        dlg._ok()
        _pp1_check(p, "ax2")
        # Thread per Object: the threads are off (ax3 -> ax4)
        p = _pp1_project(d, "ax3")
        dlg = _pp1_dialog(w, p)
        dlg.per_object.setChecked(True)
        assert not dlg.threads.isEnabled()
        dlg._ok()
        _pp1_check(p, "ax4")
        # an ActiveX DLL: Retained In Memory once Unattended Execution is on, Single Threaded, DLL Base Address
        p = _pp1_project(d, "dll1")
        dlg = _pp1_dialog(w, p, "compile")
        assert dlg.threading.isEnabled() and dlg.threading.currentText() == "Apartment Threaded"
        assert not dlg.per_object.isEnabled() and dlg.unattended.isEnabled() and not dlg.retained.isEnabled()
        dlg.unattended.setChecked(True)
        assert dlg.retained.isEnabled()
        dlg.retained.setChecked(True)
        dlg.threading.setCurrentIndex(dlg.threading.findText("Single Threaded"))
        assert dlg.dllbase.isEnabled()
        dlg.dllbase.setText("&H12340000")
        dlg._ok()
        _pp1_check(p, "dll2")
        # Debugging: Start program, Start browser with URL, Wait for components (dll2 -> dbg1 / dbg2 / dbg3)
        p = _pp1_project(d, "dll2")
        dlg = _pp1_dialog(w, p, "debugging")
        assert not dlg.dbg_comp.isEnabled() and dlg.dbg_wait.isChecked() and dlg.use_existing.isChecked()
        dlg.dbg_prog.setChecked(True)
        assert dlg.dbg_prog_path.isEnabled() and not dlg.dbg_url_text.isEnabled()
        dlg.dbg_prog_path.setText("C:\\Windows\\notepad.exe")
        dlg._ok()
        _pp1_check(p, "dbg1")
        p = _pp1_project(d, "dbg1")
        dlg = _pp1_dialog(w, p, "debugging")
        assert dlg.dbg_prog.isChecked() and dlg.dbg_prog_path.text() == "C:\\Windows\\notepad.exe"
        dlg.dbg_url.setChecked(True)
        dlg.dbg_url_text.setText("http://localhost/")
        dlg.use_existing.setChecked(False)
        dlg._ok()
        _pp1_check(p, "dbg2")
        p = _pp1_project(d, "dbg2")
        dlg = _pp1_dialog(w, p, "debugging")
        dlg.dbg_wait.setChecked(True)
        dlg._ok()
        _pp1_check(p, "dbg3")
        # an ActiveX Control: Require License Key, Start component with its user controls (ocx1 -> ocx2)
        p = _pp1_project(d, "ocx1")
        dlg = _pp1_dialog(w, p, "debugging")
        assert not dlg.unattended.isEnabled() and dlg.license.isEnabled() and not dlg.retained.isEnabled() and \
            dlg.threading.isEnabled() and dlg.threading.currentText() == "Apartment Threaded"
        assert dlg.dbg_comp.isEnabled() and dlg.dbg_comp.isChecked() and \
            [dlg.dbg_comp_list.itemText(i) for i in range(dlg.dbg_comp_list.count())] == ["UserControl1"]
        dlg.license.setChecked(True)
        dlg._ok()
        _pp1_check(p, "ocx2")
        # Upgrade ActiveX Controls off, Make shown (no form: no IconForm), Remove information off (ocx3 -> ocx4)
        p = _pp1_project(d, "ocx3")
        dlg = _pp1_dialog(w, p, "make")
        assert not dlg.icon.isEnabled() and dlg.title.text() == "ocx3" and dlg.icon_view.pixmap().isNull()
        dlg.upgrade.setChecked(False)
        dlg.remove_unused.setChecked(False)
        dlg._ok()
        _pp1_check(p, "ocx4")
        # Binary Compatibility: a file that is not there is refused, the dialog stays on Component (stage 11)
        p = _pp1_project(d, "ocx4")
        before = p.vbp.dumps()
        dlg = _pp1_dialog(w, p)
        dlg.bin_compat.setChecked(True)
        dlg.compat_path.setText("C:\\pp1\\ocx1.ocx")
        assert dlg._ok() is False and dlg.result() == 0
        assert dlg.messages == [("error", "Unable to set the version compatible component: C:\\pp1\\ocx1.ocx")]
        assert dlg.tabs.currentIndex() == dlg.PAGES.index("component") and dlg.bin_compat.isChecked() and \
            dlg.compat_path.text() == "C:\\pp1\\ocx1.ocx" and p.vbp.dumps() == before
        dlg.reject()
        # Binary -> Project -> No with the made ocx4.ocx (ocx5 -> ocx6 -> ocx7 -> ocx8; stage 22)
        p = _pp1_project(d, "ocx5")
        open(os.path.join(os.path.dirname(p.path), "ocx4.ocx"), "wb").close()
        dlg = _pp1_dialog(w, p)
        assert dlg.proj_compat.isChecked() and dlg.compat_path.text() == "ocx4.ocx"
        dlg.bin_compat.setChecked(True)
        assert dlg._ok() and not dlg.messages
        _pp1_check(p, "ocx6")
        dlg = _pp1_dialog(w, p)
        dlg.proj_compat.setChecked(True)
        dlg._ok()
        _pp1_check(p, "ocx7")
        dlg = _pp1_dialog(w, p)
        dlg.no_compat.setChecked(True)
        assert not dlg.compat_path.isEnabled() and dlg.compat_path.text() == "ocx4.ocx"
        dlg._ok()
        _pp1_check(p, "ocx8")
        dlg = _pp1_dialog(w, p)                     # Project Compatibility with a file that is not there
        dlg.proj_compat.setChecked(True)
        dlg.compat_path.setText("C:\\pp1\\nofile.ocx")
        assert dlg._ok() is False and dlg.messages[0][1].endswith("C:\\pp1\\nofile.ocx")
        dlg.reject()
        dlg = _pp1_dialog(w, p)                     # Binary with an empty box: refused too (stage 23)
        dlg.bin_compat.setChecked(True)
        dlg.compat_path.setText("")
        assert dlg._ok() is False and dlg.messages == [("error", "Unable to set the version compatible component: ")]
        dlg.reject()
        # a Standard EXE: the Debugging page shown writes DebugStartupOption=0 (ex1 -> ex2); Sub Main (ex5 -> ex6)
        p = _pp1_project(d, "ex1")
        dlg = _pp1_dialog(w, p, "debugging")
        dlg._ok()
        _pp1_check(p, "ex2")
        p = _pp1_project(d, "ex5")
        dlg = _pp1_dialog(w, p)
        dlg.startup.setCurrentText("Sub Main")
        dlg._ok()
        assert not dlg.messages
        _pp1_check(p, "ex6")
        # a new Project Type (stages 15-20): Sub Main and IconForm, StartMode, CompatibleMode, the Debugging lines
        # dropped, no ThreadingModel; on OK "Project 'Start Mode' property has been changed."
        for src, ptype, want in (("ex2", "OleDll", "ex3"), ("ex3", "Exe", "ex7"), ("ex5", "OleExe", "ex8"),
                                 ("ex5", "Control", "ex9")):
            p = _pp1_project(d, src)
            dlg = _pp1_dialog(w, p)
            dlg.ptype.setCurrentIndex(dlg.ptype.findData(ptype))
            assert not dlg.messages                     # not when it is chosen: on OK
            if ptype != "Exe":
                assert [dlg.startup.itemText(i) for i in range(dlg.startup.count())] == ["(None)", "Sub Main"]
                assert dlg.startup.currentText() == "Sub Main" and not dlg.unattended.isEnabled()   # the form
            if ptype in ("OleDll", "Control"):
                assert dlg.threading.isEnabled() and dlg.threading.currentText() == "Single Threaded"
            assert dlg.license.isEnabled() == (ptype == "Control")
            assert dlg._ok()
            assert dlg.messages == [("warning", "Project 'Start Mode' property has been changed.")], dlg.messages
            _pp1_check(p, want)
        # a new project of each type: StartMode as VB6 (ax1, dll1, ocx1: 1), ThreadingModel for a DLL / a control
        for ptype, sm, tm in (("Exe", "0", None), ("OleExe", "1", None), ("OleDll", "1", "1"), ("Control", "1", "1")):
            p = vbmodel.VBProject.new("cp1252", ptype)
            assert p.vbp.get("StartMode") == sm and p.vbp.get("ThreadingModel") == tm, (ptype, p.vbp.dumps())
        # the menu: Project > Project1 Properties... on the project of the window
        dlg = w.project_properties_dialog()
        assert dlg.windowTitle() == "Project1 - Project Properties" and dlg.tabs.count() == 5 and \
            [dlg.tabs.tabText(i) for i in range(5)] == ["General", "Make", "Compile", "Component", "Debugging"]
        dlg.reject()
    finally:
        QMessageBox.critical = crit
    done(w)


# the .frm files the VB6 IDE wrote in DEV-8/fd1 (the form designer, phase 8.4)
FD1 = {
    'F1': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   3600\n      TabIndex        =   2\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   1\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'h1': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   3600\n      TabIndex        =   2\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1820\n      TabIndex        =   1\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'h2': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   3840\n      TabIndex        =   2\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1935\n      TabIndex        =   1\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'h3': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   3600\n      TabIndex        =   2\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1815\n      TabIndex        =   1\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'h4': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   2445\n      TabIndex        =   2\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1245\n      TabIndex        =   1\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'h5': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   2685\n      TabIndex        =   2\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1365\n      TabIndex        =   1\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'h6': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   2685\n      TabIndex        =   2\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1362\n      TabIndex        =   1\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'v1': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   2685\n      TabIndex        =   2\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1362\n      TabIndex        =   1\n      Top             =   1140\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'v2': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   2685\n      TabIndex        =   2\n      Top             =   2160\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1362\n      TabIndex        =   1\n      Top             =   1260\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'v3': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   2685\n      TabIndex        =   2\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1362\n      TabIndex        =   1\n      Top             =   1140\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'v4': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   2685\n      TabIndex        =   2\n      Top             =   1110\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1362\n      TabIndex        =   1\n      Top             =   735\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'm1': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command4 \n      Caption         =   "Command4"\n      Height          =   495\n      Left            =   5000\n      TabIndex        =   3\n      Top             =   240\n      Width           =   700\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   3612\n      TabIndex        =   2\n      Top             =   1110\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1826\n      TabIndex        =   1\n      Top             =   735\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'm2': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command4 \n      Caption         =   "Command4"\n      Height          =   495\n      Left            =   5715\n      TabIndex        =   3\n      Top             =   240\n      Width           =   700\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   4095\n      TabIndex        =   2\n      Top             =   1110\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   2070\n      TabIndex        =   1\n      Top             =   735\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1000\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    's3a': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command4 \n      Caption         =   "Command4"\n      Height          =   500\n      Left            =   5010\n      TabIndex        =   3\n      Top             =   1300\n      Width           =   700\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   4080\n      TabIndex        =   2\n      Top             =   1080\n      Width           =   855\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   2040\n      TabIndex        =   1\n      Top             =   720\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   500\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1095\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    's3': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command4 \n      Caption         =   "Command4"\n      Height          =   500\n      Left            =   5040\n      TabIndex        =   3\n      Top             =   1320\n      Width           =   735\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   4080\n      TabIndex        =   2\n      Top             =   1080\n      Width           =   855\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   2040\n      TabIndex        =   1\n      Top             =   720\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   500\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1095\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    's4a': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.TextBox Text1 \n      Height          =   500\n      Left            =   250\n      TabIndex        =   4\n      Text            =   "Text1"\n      Top             =   2010\n      Width           =   1010\n   End\n   Begin VB.CommandButton Command4 \n      Caption         =   "Command4"\n      Height          =   500\n      Left            =   5040\n      TabIndex        =   3\n      Top             =   1320\n      Width           =   735\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   4080\n      TabIndex        =   2\n      Top             =   1080\n      Width           =   855\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   2040\n      TabIndex        =   1\n      Top             =   720\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   500\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1095\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    's4': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.TextBox Text1 \n      Height          =   500\n      Left            =   240\n      TabIndex        =   4\n      Text            =   "Text1"\n      Top             =   2040\n      Width           =   1095\n   End\n   Begin VB.CommandButton Command4 \n      Caption         =   "Command4"\n      Height          =   500\n      Left            =   5040\n      TabIndex        =   3\n      Top             =   1320\n      Width           =   735\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   4080\n      TabIndex        =   2\n      Top             =   1080\n      Width           =   855\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   2040\n      TabIndex        =   1\n      Top             =   720\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   500\n      Left            =   240\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1095\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    's4-locked': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   LockControls    =   -1  \'True\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.TextBox Text1 \n      Height          =   500\n      Left            =   240\n      TabIndex        =   4\n      Text            =   "Text1"\n      Top             =   2040\n      Width           =   1095\n   End\n   Begin VB.CommandButton Command4 \n      Caption         =   "Command4"\n      Height          =   500\n      Left            =   5040\n      TabIndex        =   3\n      Top             =   1320\n      Width           =   735\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   4080\n      TabIndex        =   2\n      Top             =   1080\n      Width           =   855\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   2040\n      TabIndex        =   1\n      Top             =   720\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   500\n      Left            =   360\n      TabIndex        =   0\n      Top             =   240\n      Width           =   1215\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    't1': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   1800\n      TabIndex        =   0\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   2\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   5400\n      TabIndex        =   1\n      Top             =   2400\n      Width           =   1140\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    't2': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   1800\n      TabIndex        =   0\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   1\n      Top             =   960\n      Width           =   1200\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    't3': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   2\n      Top             =   1680\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   1800\n      TabIndex        =   0\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   1\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.Label Label1 \n      Caption         =   "Label1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   3\n      Top             =   1680\n      Width           =   1215\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    't4': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   3  \'Windows Default\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   3\n      Top             =   1680\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   1800\n      TabIndex        =   1\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   2\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.Label Label1 \n      Caption         =   "Label1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   0\n      Top             =   1680\n      Width           =   1215\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'l1': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   2  \'CenterScreen\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   3\n      Top             =   1680\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   1800\n      TabIndex        =   1\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   2\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.Label Label1 \n      Caption         =   "Label1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   0\n      Top             =   1680\n      Width           =   1215\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'l2': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   120\n   ClientTop       =   450\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   3\n      Top             =   1680\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   1800\n      TabIndex        =   1\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   2\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.Label Label1 \n      Caption         =   "Label1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   0\n      Top             =   1680\n      Width           =   1215\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'l3': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   4575\n   ClientTop       =   1875\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   3\n      Top             =   1680\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   1800\n      TabIndex        =   1\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   2\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.Label Label1 \n      Caption         =   "Label1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   0\n      Top             =   1680\n      Width           =   1215\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'l4': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   4575\n   ClientTop       =   1875\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   StartUpPosition =   1  \'CenterOwner\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   3\n      Top             =   1680\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   1800\n      TabIndex        =   1\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   2\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.Label Label1 \n      Caption         =   "Label1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   0\n      Top             =   1680\n      Width           =   1215\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
    'l5': 'VERSION 5.00\nBegin VB.Form Form1 \n   Caption         =   "Form1"\n   ClientHeight    =   3750\n   ClientLeft      =   9045\n   ClientTop       =   5310\n   ClientWidth     =   7305\n   LinkTopic       =   "Form1"\n   ScaleHeight     =   3750\n   ScaleWidth      =   7305\n   Begin VB.CommandButton Command1 \n      Caption         =   "Command1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   3\n      Top             =   1680\n      Width           =   1215\n   End\n   Begin VB.CommandButton Command3 \n      Caption         =   "Command3"\n      Height          =   615\n      Left            =   1800\n      TabIndex        =   1\n      Top             =   1920\n      Width           =   800\n   End\n   Begin VB.CommandButton Command2 \n      Caption         =   "Command2"\n      Height          =   375\n      Left            =   1560\n      TabIndex        =   2\n      Top             =   960\n      Width           =   1200\n   End\n   Begin VB.Label Label1 \n      Caption         =   "Label1"\n      Height          =   495\n      Left            =   3000\n      TabIndex        =   0\n      Top             =   1680\n      Width           =   1215\n   End\nEnd\nAttribute VB_Name = "Form1"\nAttribute VB_GlobalNameSpace = False\nAttribute VB_Creatable = False\nAttribute VB_PredeclaredId = True\nAttribute VB_Exposed = False\n',
}


def _fd1_designer(w, d, name, text):
    """a project with the form of a .frm of DEV-8/fd1, opened in the designer"""
    folder = os.path.join(d, name)
    os.makedirs(folder, exist_ok=True)
    with open(os.path.join(folder, "F.frm"), "w", encoding="cp1252", newline="\r\n") as f:
        f.write(text)
    vbp = os.path.join(folder, "fd1.vbp")
    with open(vbp, "w", encoding="cp1252", newline="\r\n") as f:
        f.write(PP1["base"].replace("Form=Form1.frm", "Form=F.frm"))
    for q in ([w.project] if w.project else []):
        q.dirty = False
        for m in q.modules:
            m.dirty = False
    assert w.open_project(vbp)
    dz = w.current_designer()
    assert dz is not None and dz.root.name == "Form1"
    return dz


def _fd1_text(dz):
    """the .frm the designer would save"""
    out = os.path.join(os.path.dirname(dz.module.path), "out.frm")
    dz.module.write_form(out, permanent=False)
    return open(out, encoding="cp1252").read().replace("\r\n", "\n")


def _fd1_check(dz, want):
    got = _fd1_text(dz)
    assert got == FD1[want], f"{want}:\n" + "\n".join(
        __import__("difflib").unified_diff(FD1[want].split("\n"), got.split("\n"), lineterm=""))


def _fd1_items(dz, *names):
    byname = {c.name: c for c in dz.root.walk()}
    return [dz.items[byname[n]] for n in names]


def t_formdesign84(d):
    from lbstudio import formops as F
    from lbstudio.dialogs import OptionsDialog
    w = MainWindow()
    w.show()
    K = Qt.Key
    M = Qt.KeyboardModifier
    # Format menu: with the form selected only Lock Controls (DEV-8/fd1 stage 1)
    dz = _fd1_designer(w, d, "a", FD1["F1"])
    dz.select([])
    w.update_format_menu()
    assert not any(a.isEnabled() for acts in w.fmt_actions.values() for a in acts) and w.a_lock.isEnabled()
    titles = [a.text() for a in w.format_menu.actions() if not a.isSeparator()]
    assert titles == ["&Align", "&Make Same Size", "Size to Gr&id", "&Horizontal Spacing", "&Vertical Spacing",
                      "&Center in Form", "&Order", "&Lock Controls"], titles
    # Horizontal Spacing (stage 3): Command3 has the solid handles; the leftmost stays
    c1, c2, c3 = _fd1_items(dz, "Command1", "Command2", "Command3")
    dz.select([c1, c2, c3], primary=c3)
    w.update_format_menu()
    assert all(a.isEnabled() for acts in w.fmt_actions.values() for a in acts)
    for how, want in (("equal", "h1"), ("increase", "h2"), ("decrease", "h3"), ("remove", "h4")):
        dz.spacing(False, how)
        _fd1_check(dz, want)
    dz.select([c2, c1, c3], primary=c3)          # stage 4: the order of the selection does not matter
    dz.spacing(False, "increase")
    _fd1_check(dz, "h5")
    dz.spacing(False, "equal")
    _fd1_check(dz, "h6")
    dz.select([c1, c2, c3], primary=c3)          # stage 5: Vertical Spacing
    for how, want in (("equal", "v1"), ("increase", "v2"), ("decrease", "v3"), ("remove", "v4")):
        dz.spacing(True, how)
        _fd1_check(dz, want)
    # stage 6: a fourth button, Make Equal; Increase with a grid of 240
    dz.add_default("CommandButton")
    c4 = _fd1_items(dz, "Command4")[0]
    dz.set_rect(c4, 5000, 240, 700, 495)
    dz.select([c1, c2, c3, c4], primary=c4)
    dz.spacing(False, "equal")
    _fd1_check(dz, "m1")
    dz.grid_w = dz.grid_h = 240
    dz.spacing(False, "increase")
    _fd1_check(dz, "m2")
    # Size to Grid (stages 8, 9): Left / Top / right edge on the grid, the Height stays
    dz = _fd1_designer(w, d, "b", FD1["s3a"])
    dz.select(_fd1_items(dz, "Command4"))
    dz.size_to_grid()
    _fd1_check(dz, "s3")
    dz = _fd1_designer(w, d, "c", FD1["s4a"])
    dz.select(_fd1_items(dz, "Text1"))
    w.update_format_menu()                      # one control: Align to Grid, Size to Grid; not Spacing
    one = [a.text() for a in w.fmt_actions["one"] if a.isEnabled()]
    assert "Size to Gr&id" in one and "to &Grid" in one and not any(a.isEnabled() for a in w.fmt_actions["two"])
    # the IDE back from another program with a click on the menu: no current window in QMdiArea for a moment
    w.mdi.currentSubWindow = lambda: None
    try:
        for a in w.fmt_actions["one"]:
            a.setEnabled(False)
        w.a_lock.setEnabled(False)
        w.update_format_menu()
        assert w.current_designer() is dz and w.a_lock.isEnabled()
        assert all(a.isEnabled() for a in w.fmt_actions["one"])
    finally:
        del w.mdi.currentSubWindow
    dz.size_to_grid()
    _fd1_check(dz, "s4")
    # Lock Controls (stage 10): the mouse does not move or size them; Ctrl+arrow and Shift+arrow do
    dz.toggle_lock()
    assert dz.locked
    w.update_format_menu()
    assert w.a_lock.isChecked()
    b1 = _fd1_items(dz, "Command1")[0]
    # the designer draws light, whatever the theme of the IDE / the system (the program follows the system)
    from PyQt6.QtGui import QPalette
    assert b1.visual.palette().color(QPalette.ColorRole.Button).name() == "#f0f0f0"
    assert b1.visual.palette().color(QPalette.ColorRole.ButtonText).name() == "#000000"
    assert b1.visual.style().name().lower() == "fusion"
    assert dz.surface.palette().color(QPalette.ColorRole.Base).name() == "#ffffff"
    dz.select([b1])
    r0 = b1.rect_twips()
    QTest.mousePress(b1, Qt.MouseButton.LeftButton, M.NoModifier, QPoint(5, 5))
    QTest.mouseMove(b1, QPoint(60, 40))
    QTest.mouseRelease(b1, Qt.MouseButton.LeftButton, M.NoModifier, QPoint(60, 40))
    app.processEvents()
    assert b1.rect_twips() == r0, (r0, b1.rect_twips())
    hnd = next(x for x in dz.surface.handles if x.pos_name == "e")
    QTest.mousePress(hnd, Qt.MouseButton.LeftButton, M.NoModifier, QPoint(3, 3))
    assert dz.surface.resizing is None
    QTest.mouseRelease(hnd, Qt.MouseButton.LeftButton, M.NoModifier, QPoint(3, 3))
    assert dz.surface.handles and not any(h.solid() for h in dz.surface.handles)   # stage 20: all hollow
    dz.arrow_key(K.Key_Right, M.ControlModifier)
    dz.arrow_key(K.Key_Right, M.ShiftModifier)
    _fd1_check(dz, "s4-locked")
    dz.toggle_lock()
    assert not dz.locked and "LockControls" not in _fd1_text(dz)
    assert all(h.solid() for h in dz.surface.handles)          # one control: its handles solid again
    # the arrow keys (stages 11-15): Ctrl / Shift by the grid or by a pixel; an arrow alone selects
    assert F.nudge(240, 1, 120, True) == 360 and F.nudge(360, 1, 120, False) == 375 and \
        F.nudge(375, 1, 120, True) == 495 and F.nudge(1000, 1, 120, True, size=True) == 1125 and \
        F.nudge(1125, 1, 120, False, size=True) == 1140 and F.nudge(1260, -1, 120, True, size=True) == 1140
    dz = _fd1_designer(w, d, "e", FD1["F1"])
    c1, c2, c3 = _fd1_items(dz, "Command1", "Command2", "Command3")

    def arrow(start, key):
        dz.select([start])
        dz.arrow_key(key, M.NoModifier)
        return dz.primary.ctl.name

    assert arrow(c1, K.Key_Right) == "Command2"                       # stage 11
    dz.set_rect(c1, 5400, 2400, 1140, 495)                            # stage 14
    assert [arrow(c3, K.Key_Right), arrow(c2, K.Key_Left), arrow(c1, K.Key_Up), arrow(c2, K.Key_Down),
            arrow(c1, K.Key_Left)] == ["Command1", "Command2", "Command1", "Command2", "Command3"]
    dz.set_rect(c3, 1800, 1920, 800, 615)                             # stage 15
    assert [arrow(c2, K.Key_Right), arrow(c2, K.Key_Down), arrow(c3, K.Key_Up)] == \
        ["Command1", "Command3", "Command2"]
    # Tab order (stage 16, on the form of stage 15): TabIndex 0 moves the others up; Delete closes the gap; new
    # controls come last
    c1, c3 = _fd1_items(dz, "Command1", "Command3")
    dz.set_tabindex(c3.ctl, 0)
    _fd1_check(dz, "t1")
    dz.select([c1])
    dz.delete_selection()
    _fd1_check(dz, "t2")
    dz.add_default("CommandButton")
    dz.add_default("Label")
    _fd1_check(dz, "t3")
    dz.set_tabindex(_fd1_items(dz, "Label1")[0].ctl, 0)
    _fd1_check(dz, "t4")
    # Form Layout (stages 17-19): Startup Position, dragging, the forms whose designer is open
    fl = w.formlayout
    assert fl.designers() == [dz]
    fl.set_startup(dz, 2)
    _fd1_check(dz, "l1")
    fl.set_startup(dz, 0)
    _fd1_check(dz, "l2")
    b, cap, menu = dz.frame_w.frame()
    fl._place(dz, 4575 - b * 15, 1875 - (b + cap + menu) * 15, final=True)
    _fd1_check(dz, "l3")
    fl.set_startup(dz, 1)
    _fd1_check(dz, "l4")
    fl._place(dz, 9045 - b * 15, 5310 - (b + cap + menu) * 15, final=True)
    _fd1_check(dz, "l5")
    r = fl.to_widget(fl.window_twips(dz))
    menu_ = fl.context_menu(r.center().toPoint())
    texts = [a.text() for a in menu_.actions() if not a.isSeparator()]
    assert texts == ["Resolution &Guides", "Startup &Position", "Doc&kable", "&Hide"], texts
    sp = next(a for a in menu_.actions() if a.text() == "Startup &Position").menu()
    assert [a.text() for a in sp.actions()] == ["&Manual", "Center &Owner", "Center &Screen", "&Windows Default"] and \
        sp.actions()[0].isChecked()
    # stage 20: Resolution Guides on at first; Dockable off puts the window among the code and form windows
    assert settings.DEFAULTS["layout_guides"] is True and settings.DEFAULTS["layout_dockable"] is True
    acts = {a.text(): a for a in menu_.actions()}
    assert acts["Resolution &Guides"].isChecked() == fl.guides and acts["Doc&kable"].isEnabled() and \
        acts["Doc&kable"].isChecked()
    from PyQt6.QtWidgets import QMdiArea
    QDW = Qt.DockWidgetArea
    tabbed = w.mdi.viewMode() == QMdiArea.ViewMode.TabbedView
    if not tabbed:
        w.toggle_tabbed()
    w.set_layout_dockable(False)                # windows as tabs: a free window that does not dock
    app.processEvents()
    assert w.layout_sub is None and w.dock_layout.isFloating() and w.dock_layout.widget() is fl
    assert w.dock_layout.allowedAreas() == QDW.NoDockWidgetArea and not w.dock_layout.isHidden()
    assert w.dock_layout.width() < 400 and fl.designers() == [dz]
    vpr = QRect(w.mdi.viewport().mapToGlobal(QPoint(0, 0)), w.mdi.viewport().size())
    assert vpr.contains(w.dock_layout.pos()), (vpr, w.dock_layout.pos())     # in the MDI area, by the form
    acts = {a.text(): a for a in fl.context_menu(r.center().toPoint()).actions()}
    assert not acts["Doc&kable"].isChecked()
    w.hide_layout()
    assert w.dock_layout.isHidden()
    w.show_layout()
    assert not w.dock_layout.isHidden()
    w.toggle_tabbed()                           # the MDI view: a small window among the others (VB6)
    app.processEvents()
    assert w.layout_sub is not None and w.layout_sub.widget() is fl and w.layout_sub in w.mdi.subWindowList()
    assert w.dock_layout.isHidden() and not w.layout_sub.isMaximized() and w.layout_sub.width() < 400
    assert fl.designers() == [dz]
    w.hide_layout()
    assert w.layout_sub.isHidden()
    w.show_layout()
    assert not w.layout_sub.isHidden()
    w.toggle_tabbed()                           # back to tabs: the free window again
    app.processEvents()
    assert w.layout_sub is None and w.dock_layout.isFloating() and w.dock_layout.widget() is fl
    w.set_layout_dockable(True)
    app.processEvents()
    assert w.layout_sub is None and w.dock_layout.widget() is fl and settings.S["layout_dockable"] is True
    assert not w.dock_layout.isFloating() and w.dock_layout.allowedAreas() == QDW.AllDockWidgetAreas
    if not tabbed:
        w.toggle_tabbed()
    assert fl.designers() == [dz]
    w.add_module("Form")                        # a second form in the designer: two in the layout
    assert len(fl.designers()) == 2
    w.design_windows[w.project.module("Form2")].close()
    app.processEvents()
    assert fl.designers() == [dz]
    # Tools > Options > General: Form Grid Settings
    dlg = OptionsDialog(w)
    assert dlg.grid_w.value() == 120 and dlg.grid_h.value() == 120 and dlg.show_grid.isChecked() and \
        dlg.align_grid.isChecked()
    dlg.grid_w.setValue(240)
    dlg.align_grid.setChecked(False)
    dlg._ok()
    w.apply_grid_settings()
    assert dz.grid_w == 240 and dz.grid_h == 120 and not dz.align_grid
    settings.S["grid_width"] = 120
    settings.S["align_grid"] = True
    w.apply_grid_settings()
    done(w)


RS1 = os.path.join(HERE, "rs1")                 # the files the VB 6 Resource Editor wrote (DEV-8/rs1)


def _rs1(name):
    with open(os.path.join(RS1, name), "rb") as f:
        return f.read()


# the menus of VB6 (DEV-8/me1, Standard EXE with Form1 in the designer): text | shortcut, in order; the menus of
# LinBASIC Studio may have more items between them (its own commands)
VB6_MENUS = {
    "&File": ["&New Project|Ctrl+N", "&Open Project...|Ctrl+O", "A&dd Project...", "&Remove Project", "Sa&ve Project",
              "Sav&e Project As...", "&Save Form1|Ctrl+S", "Save Form1 &As...", "&Print...|Ctrl+P", "Print Set&up...",
              "Ma&ke Project1.exe...", "Make Project &Group...", "E&xit|Alt+Q"],
    "&Edit": ["Can't Undo|Ctrl+Z", "Can't Redo|", "Cu&t|Ctrl+X", "&Copy|Ctrl+C", "&Paste|Ctrl+V", "Paste Lin&k",
              "&Delete|Del", "Select &All|Ctrl+A", "&Find...|Ctrl+F", "Find &Next|F3", "R&eplace...|Ctrl+H",
              "&Indent|Tab", "&Outdent|Shift+Tab", "Insert Fi&le...", "List Properties/Met&hods|Ctrl+J",
              "Li&st Constants|Ctrl+Shift+J", "&Quick Info|Ctrl+I", "Para&meter Info|Ctrl+Shift+I",
              "Complete &Word|Ctrl+Space", "&Bookmarks"],
    "&View": ["&Code", "O&bject|Shift+F7", "&Definition|Shift+F2", "Last Positio&n|Ctrl+Shift+F2",
              "&Object Browser|F2", "&Immediate Window|Ctrl+G", "Local&s Window", "Watc&h Window",
              "Call Stac&k...|Ctrl+L", "&Project Explorer|Ctrl+R", "Properties &Window|F4", "&Form Layout Window",
              "Property Pa&ges|Shift+F4",
              "Toolbo&x", "Color Pa&lette", "&Toolbars"],
    "&Project": ["Add &Form", "Add MD&I Form", "Add &Module", "Add &Class Module", "Add &User Control",
                 "Add &Property Page", "Add User &Document", "Add WebClass", "Add Data Report", "Add DHTML Page",
                 "Add Data Environment", "More ActiveX Desi&gners...", "&Add File...|Ctrl+D", "&Remove Form1",
                 "Refere&nces...", "C&omponents...|Ctrl+T", "Project1 Prop&erties..."],
    "&Debug": ["Step &Into|F8", "Step &Over|Shift+F8", "Step O&ut|Ctrl+Shift+F8", "&Run To Cursor|Ctrl+F8",
               "&Add Watch...", "&Edit Watch...|Ctrl+W", "&Quick Watch...|Shift+F9", "&Toggle Breakpoint|F9",
               "&Clear All Breakpoints|Ctrl+Shift+F9", "Set &Next Statement|Ctrl+F9", "Show Ne&xt Statement"],
    "&Run": ["&Start|F5", "Start With &Full Compile|Ctrl+F5", "Brea&k|Ctrl+Break", "&End", "&Restart|Shift+F5"],
    "&Tools": ["Add &Procedure...", "Procedure &Attributes...", "&Menu Editor...|Ctrl+E", "&Options..."],
    "&Add-Ins": ["&Visual Data Manager...", "&Add-In Manager..."],
    "&Window": ["S&plit", "Tile &Horizontally", "Tile &Vertically", "&Cascade", "&Arrange Icons"],
    "&Help": ["&Contents...", "&Index...", "&Search...", "&About LinBASIC Studio..."],
}
# the commands of VB6's menus (data/vb6/ide_commands.json) that LinBASIC Studio leaves out (user, 2026-11-17):
# Data View, the Enterprise tools, SourceSafe, the Microsoft links, and what the IDE itself does not have
VB6_LEFT_OUT = {"Save Selection", "Save Change Script", "Get", "Check Out", "Check In", "Undo Check Out",
                "Technical Support", "Microsoft on the Web", "About Microsoft Visual Basic...", "Report Designer...",
                "Debug Options...  (Do not localize this string)", "Disable All Watches", "More Windows...",
                "Arrange Buttons", "Tab Order", "Undo", "Redo", "Save <File Name>", "Save <File Name> As...",
                "Remove <Item>", "<Project Name> Properties...", "Make <Project>..."}
# still to come in Phase 8.7
VB6_PENDING_87 = set()


def _menu_items(menu):
    out = []
    for a in menu.actions():
        if a.isSeparator():
            continue
        text, tab, shown = a.text().partition("\t")      # a shortcut Qt names otherwise (Ctrl+Break)
        out.append((text, shown if tab else a.shortcut().toString(), a))
    return out


def t_menus87(d):
    """Phase 8.7: the menus as VB6's (DEV-8/me1)"""
    import json
    w = MainWindow()
    w.show()
    app.processEvents()
    assert (w.tb_pos.text(), w.tb_size.text()) == ("0, 0", "4800 x 3600")   # at once, for the first form
    mb = w.menuBar()
    tops = [a.text() for a in mb.actions()]
    assert tops == ["&File", "&Edit", "&View", "&Project", "F&ormat", "&Debug", "&Run", "&Tools", "&Add-Ins",
                    "&Window", "&Help"], tops
    for a in mb.actions():
        a.menu().aboutToShow.emit()
    for title, want in VB6_MENUS.items():
        menu = next(a.menu() for a in mb.actions() if a.text() == title)
        have = _menu_items(menu)
        k = 0
        for item in want:
            text, _, key = item.partition("|")
            while k < len(have) and have[k][0] != text:
                k += 1
            assert k < len(have), (title, item, [h[0] for h in have])
            if "|" in item:
                assert have[k][1] == key, (title, item, have[k][1])
            k += 1
    # grey as in VB6 with the form in the designer
    edit = {t: a for t, _, a in _menu_items(next(a.menu() for a in mb.actions() if a.text() == "&Edit"))}
    assert not edit["Can't Undo"].isEnabled() and not edit["Paste Lin&k"].isEnabled()
    assert not edit["Insert Fi&le..."].isEnabled() and not edit["&Find..."].isEnabled()
    proj = {t: a for t, _, a in _menu_items(next(a.menu() for a in mb.actions() if a.text() == "&Project"))}
    for t in ("Add User &Document", "Add WebClass", "Add DHTML Page", "More ActiveX Desi&gners..."):
        assert not proj[t].isEnabled(), t
    assert proj["Add &Form"].isEnabled() and proj["&Remove Form1"].isEnabled()
    w._update_file_actions()                    # one project, no group: Make Project Group... grey (me2-23)
    assert not w.a_make_group.isEnabled()
    # Project > Add MDI Form (DEV-8/me2 stage 1): the dialog (New), the .frm as VB6 writes it, one per project
    from PyQt6.QtCore import QTimer as _QT0
    w.S["add_mdi_nodialog"] = False
    w.update_project_menu()
    assert w.a_add_mdi.isEnabled()

    def _new_mdi():
        dlg = w._last_add_mdi
        assert dlg.windowTitle() == "Add MDI Form" and dlg.tabs.tabText(0) == "New" and dlg.tabs.tabText(1) == "Existing"
        assert dlg.items.currentItem().text() == "MDI Form" and not dlg.dont_show.isChecked()
        dlg.b_open.click()
    _QT0.singleShot(0, _new_mdi)
    w.a_add_mdi.trigger()
    mdi = w.project.module("MDIForm1")
    assert mdi is not None and mdi.display_kind == "MDIForm" and w.project.mdi_form() is mdi
    mdi.save(os.path.join(d, "MDIForm1.frm"))
    got = open(os.path.join(d, "MDIForm1.frm"), "rb").read().decode("cp1252")
    want = ("VERSION 5.00\r\nBegin VB.MDIForm MDIForm1 \r\n   BackColor       =   &H8000000C&\r\n"
            "   Caption         =   \"MDIForm1\"\r\n   ClientHeight    =   3030\r\n   ClientLeft      =   120\r\n"
            "   ClientTop       =   450\r\n   ClientWidth     =   4560\r\n   LinkTopic       =   \"MDIForm1\"\r\n"
            "   StartUpPosition =   3  'Windows Default\r\nEnd\r\nAttribute VB_Name = \"MDIForm1\"\r\n"
            "Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n"
            "Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\n")
    assert got == want, got
    ents = [(k.strip(), v) for k, v in w.project.vbp.entries if v is not None]
    i_mdi = next(i for i, (k, v) in enumerate(ents) if k == "Form" and "MDIForm1.frm" in v)
    i_ref = max(i for i, (k, v) in enumerate(ents) if k == "Reference")
    assert i_mdi == i_ref + 1, ents                     # Form=MDIForm1.frm right after the Reference (VB6)
    assert ents[:3] == [("Type", "Exe"), ("Form", "Form1.frm"), ("Reference", ents[2][1])], ents[:3]   # new project
    assert w.designer_title(mdi).startswith("Project1 - MDIForm1 (MDIForm)")
    md = w.design_windows[mdi].widget()
    assert md.surface.palette().window().color().name() == "#ababab"           # AppWorkspace
    w.update_project_menu()
    assert not w.a_add_mdi.isEnabled()                                          # one MDI form per project
    w.design_windows[mdi].close()
    w.project.remove_module(mdi)
    w.fill_tree()
    w.update_project_menu()
    assert w.a_add_mdi.isEnabled()
    # the code window: Undo after typing, Insert File, the window list
    m = w.project.modules[0]
    cw = w.open_code(m)
    app.processEvents()
    from PyQt6.Qsci import QsciScintilla as _S
    cw.editor.SendScintilla(_S.SCI_REPLACESEL, 0, b"' typed")         # an edit that can be undone
    w.update_edit_menu()
    assert w.current_editor() is cw.editor, w.current_sub()
    assert w.a_undo.text() == "&Undo" and w.a_undo.isEnabled(), w.a_undo.text()
    assert w.a_insert_file.isEnabled() and w.a_redo.text() == "Can't Redo"
    from lbstudio import mainwindow as _mw
    assert _mw.INSERT_FILE_FILTER.split(";;") == ["Text Files (*.txt)", "Basic Files (*.bas)",
                                                  "Class Files (*.cls)", "All Files (*)"]
    w.fill_window_menu()
    names = [a.text() for a in w.window_menu.actions() if not a.isSeparator()]
    assert names[:5] == ["S&plit", "Tile &Horizontally", "Tile &Vertically", "&Cascade", "&Arrange Icons"], names
    assert any(t.startswith("&1 ") for t in names) and any(a.isChecked() for a in w.window_menu.actions())
    # Window > Split: two panes of one text (DEV-8/me2 pictures 5, 6)
    assert w.a_split.isEnabled() and not w.a_split.isChecked()
    w.a_split.trigger()
    app.processEvents()
    assert cw.is_split() and len(cw.panes()) == 2 and cw.second.isVisible()
    top, bottom = cw.second, cw.first
    assert cw.splitter.indexOf(top) == 0 and cw.splitter.indexOf(bottom) == 1
    top.SendScintilla(_S.SCI_DOCUMENTEND)
    top.SendScintilla(_S.SCI_REPLACESEL, 0, b"\r\n' second pane")
    assert bottom.all_text() == top.all_text() and "' second pane" in bottom.all_text()
    top.setFocus()
    cw._pane_focused(top)
    assert w.current_editor() is top and w._cw_of(top) is cw
    w.fill_window_menu()
    assert w.a_split.isChecked()
    w.a_split.trigger()                                  # again: one pane, the text stays
    app.processEvents()
    assert not cw.is_split() and cw.panes() == [bottom] and w.current_editor() is bottom
    from PyQt6 import sip as _sip
    assert _sip.isdeleted(top)                           # gone at once: no orphan pane calling a closed window
    assert "' second pane" in bottom.all_text()
    cw.set_split(True)
    cw.splitter.setSizes([0, 300])                       # the bar dragged to the top: one pane again
    cw._splitter_moved(0, 1)
    assert not cw.is_split()
    # View > Color Palette (DEV-8/me2 pictures 7, 8): a colour of the palette is the BackColor of the form
    from lbstudio import colorpal as CP
    from lbstudio import vbprops as _V
    assert len(CP.PALETTE) == 48 and CP.PALETTE[34] == 0xFF0000 and CP.PALETTE[42] == 0x804040
    fd = w.open_designer(w.project.module("Form1"))
    app.processEvents()
    w.show_palette()
    app.processEvents()
    pal = w.palette_win
    assert w.dock_palette.isVisible() and w.dock_palette.isFloating()
    assert pal.designer is fd and pal.ctl is fd.root and pal.sample.target == "BackColor"
    assert not pal.b_define.isEnabled() and pal.grid.current == -1        # &H8000000F&: not a box of the palette
    pal.cell_clicked(34)                                                  # the red box
    root = w.project.module("Form1").sf.root
    back = _V.value_of(root, _V.prop_def("Form", "BackColor"))
    assert back == 0x000000FF, hex(back)                                  # &H000000FF&, as VB6 wrote it
    assert pal.grid.current == 34 and pal.sample.back.name() == "#ff0000" and fd.undo_stack
    pal.sample.target = "ForeColor"
    pal.cell_clicked(18)                                                  # FF8080
    fore = _V.value_of(w.project.module("Form1").sf.root, _V.prop_def("Form", "ForeColor"))
    assert fore == 0x008080FF, hex(fore)
    pal.sample.target = "BackColor"
    pal.set_default()                                                     # Default: &H8000000F& again
    back = _V.value_of(w.project.module("Form1").sf.root, _V.prop_def("Form", "BackColor"))
    assert back == 0x8000000F, hex(back)
    pal.cell_clicked(CP.COLS * 3 + 2)                                     # a colour of your own: Define Colors...
    assert pal.b_define.isEnabled()
    w.dock_palette.hide()
    # View > Toolbars: Debug, Edit, Form Editor, Standard (DEV-8/me2 pictures 13-16)
    assert [a.text() for a in w.m_toolbars.actions() if not a.isSeparator()][:4] == \
        ["&Debug", "&Edit", "&Form Editor", "&Standard"]
    assert w.tb_debug.isHidden() and w.tb_form_editor.isHidden()
    dbg = [a for a in w.tb_debug.actions()]
    assert [a for a in dbg if not a.isSeparator()] == [w.a_start, w.a_break, w.a_end, w.a_toggle_bp, w.a_step_into,
                                                        w.a_step_over, w.a_step_out, w.a_view_locals,
                                                        w.a_view_immediate, w.a_view_watch, w.a_quick_watch,
                                                        w.a_call_stack]
    assert sum(1 for a in dbg if a.isSeparator()) == 2
    fe = [a for a in w.tb_form_editor.actions() if not a.isSeparator()]
    assert fe[0] is w.a_front and fe[1] is w.a_back and fe[-1] is w.a_lock and len(w.fe_buttons) == 3
    assert [b.menu() for b in w.fe_buttons] == [w.m_align, w.m_center, w.m_same_size]
    fd.select([])                                      # the form: only Lock Controls (as the Format menu)
    w.update_format_menu()
    assert not w.a_front.isEnabled() and not w.fe_buttons[0].isEnabled() and w.a_lock.isEnabled()
    for tb in (w.tb_debug, w.tb_form_editor):
        tb.toggleViewAction().trigger()
        assert tb.isVisible()
        tb.toggleViewAction().trigger()
    # the Standard toolbar: where the selected control is and its size, the form's outer size, Ln / Col in code
    # (VB6: "0, 0"  "4800 x 3600" for a new form, "1680, 1320"  "1215 x 495" for a button, "Ln 2, Col 1")
    w.mdi.setActiveSubWindow(w.design_windows[fd.module])
    app.processEvents()
    fd.select([])
    assert w.toolbar_pos_text() == ("0, 0", "4800 x 3600"), w.toolbar_pos_text()
    assert w.tb_pos.text() == "0, 0" and w.tb_size.text() == "4800 x 3600" and not w.tb_size_icon.isHidden()
    btn = fd.place_new("CommandButton", QRect(16, 16, 80, 30), None)
    fd.set_rect(fd.items[btn], 1680, 1320, 1215, 495)              # moved / resized: the toolbar follows
    assert w.tb_pos.text() == "1680, 1320" and w.tb_size.text() == "1215 x 495", (w.tb_pos.text(), w.tb_size.text())
    cwb = w.open_code(fd.module)
    app.processEvents()
    w.update_status()
    assert w.tb_pos.text().startswith("Ln ") and ", Col " in w.tb_pos.text() and w.tb_size.text() == ""
    assert w.tb_pos_icon.isHidden() and w.tb_size_icon.isHidden()
    w.code_windows[fd.module].close()
    w.mdi.setActiveSubWindow(w.design_windows[fd.module])
    app.processEvents()
    fd.select([fd.items[btn]])
    fd.delete_selection()
    del cwb
    # View > Property Pages (DEV-8/me2 pictures 11, 12): the "None" page for the controls of the IDE
    from PyQt6.QtCore import QTimer as _QT
    _QT.singleShot(0, lambda: w._last_ppages.reject())
    w.a_view_ppages.trigger()
    dlg = w._last_ppages
    assert dlg.windowTitle() == "Property Pages" and dlg.tabs.tabText(0) == "None"
    assert dlg.message.text() == "There are no selected objects with property pages." and not dlg.b_apply.isEnabled()
    # every command of VB6's menus is there, left out on purpose, or still to come
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "data", "vb6", "ide_commands.json")
    cmds = json.load(open(path, encoding="utf-8"))["menus"]
    ours = set()
    stack = [a.menu() for a in mb.actions()]
    while stack:
        mm = stack.pop()
        for a in mm.actions():
            ours.add(a.text().partition("\t")[0].replace("&", "").strip())
            if a.menu() is not None:
                stack.append(a.menu())
    main_ids = list(range(19144, 19163)) + list(range(19176, 19196)) + list(range(19208, 19226)) + \
        list(range(19240, 19261)) + list(range(19305, 19318)) + list(range(19336, 19342)) + \
        list(range(19369, 19379)) + [19400, 19401] + list(range(19432, 19475))
    missing = []
    for i in main_ids:
        t = cmds.get(str(i))
        if t is None:
            continue
        t = t.replace("&", "").strip()
        if t in ours or t in VB6_LEFT_OUT or t in VB6_PENDING_87 or t in ("Module", "Class Module", "User Control",
                                                                          "Property Page", "User Document"):
            continue
        missing.append(t)
    assert not missing, missing


def t_theme(d):
    """Tools > Options > Theme: system (as the OS is; LBVC_THEME light in the tests), light, dark; the forms of
    the designer stay light"""
    from lbstudio.dialogs import OptionsDialog
    from PyQt6.QtGui import QPalette as _P
    w = MainWindow()
    w.show()                                    # the form of the new project fills its tab from the start
    app.processEvents()
    from PyQt6.QtWidgets import QMdiArea as _M
    if w.mdi.viewMode() == _M.ViewMode.TabbedView:
        assert w.current_sub() is not None and w.current_sub().isMaximized()
    S = w.S
    old = S["dark_mode"]
    try:
        dlg = OptionsDialog(w)
        assert [dlg.theme.itemText(i) for i in range(dlg.theme.count())] == ["system", "light", "dark"]
        win = lambda: app.palette().color(_P.ColorRole.Window).name()
        for mode, dark, colour in (("dark", True, "#2d2d30"), ("light", False, "#f0f0f0"), ("system", False, "#f0f0f0")):
            S["dark_mode"] = mode
            w.apply_theme()
            assert w.is_dark() is dark and win() == colour, (mode, win())
            assert w.mdi.background().color().name() == ("#1a1a1a" if dark else "#a0a0a0"), mode
            w.styles = w._load_styles()
        S["dark_mode"] = "dark"
        w.apply_theme()
        f = w.current_designer()
        assert f is not None and f.surface.palette().color(_P.ColorRole.Text).name() == "#000000"
        os.environ["LBVC_THEME"] = "dark"
        S["dark_mode"] = "system"
        assert w.is_dark()
        # the theme of the system, asked each time (MATE does not tell Qt): GTK_THEME first
        from lbstudio.mainwindow import system_dark
        gtk = os.environ.get("GTK_THEME")
        try:
            os.environ["GTK_THEME"] = "Yaru-blue-dark"
            assert system_dark()
            os.environ["GTK_THEME"] = "Yaru-blue"
            assert not system_dark()
        finally:
            if gtk is None:
                os.environ.pop("GTK_THEME", None)
            else:
                os.environ["GTK_THEME"] = gtk
        assert w.theme_timer.isActive() and w.theme_timer.interval() == 3000
    finally:
        os.environ["LBVC_THEME"] = "light"
        S["dark_mode"] = old
        w.apply_theme()


def t_resedit86(d):
    """Phase 8.6: the VB 6 Resource Editor (DEV-8/rs1)"""
    from lbstudio import resmodel as RM, lcid as LC, resedit as RE
    # the files of VB6, read and written again: the same bytes
    for n in ("s1-stage3.RES", "s1-stage3b.RES", "s1-stage4.RES", "s1-stage4b.RES"):
        b = _rs1(n)
        assert RM.ResFile.from_bytes(b, 0x0408).to_bytes() == b, n
    # made as VB6 made them (stages 2-4)
    r = RM.ResFile(0x0408)
    r.set_string(0x0408, 101, "Hello")
    r.set_string(0x0408, 102, "World")
    assert len(r.to_bytes()) == 116
    bm = r.add_bitmap(_rs1("input/pic1.bmp"))
    ic = r.add_icon(_rs1("input/icon1.ico"))
    cu = r.add_icon(_rs1("input/cur1.cur"), cursor=True)
    cs = r.add_custom(_rs1("input/data1.txt"))
    assert (bm.name, ic.name, cu.name, cs.name, cs.rtype) == (101, 101, 101, 101, "CUSTOM")
    assert r.to_bytes() == _rs1("s1-stage3.RES")
    cs.rtype, cs.name = RM.norm_name('"MYDATA"'), RM.norm_name('"INTRO"')
    ic.name, ic.lang = RM.norm_name("200"), 0x0409
    assert r.to_bytes() == _rs1("s1-stage3b.RES")
    r.set_string(0x0408, 120, "far")
    r.set_string(0x0409, 101, "Hi")
    assert r.to_bytes() == _rs1("s1-stage4.RES")
    ic2 = r.add_icon(_rs1("input/icon2.ico"))
    assert ic2.name == 201
    assert r.to_bytes() == _rs1("s1-stage4b.RES")
    assert r.duplicate(bm).name == 102
    assert RM.norm_name('"lower"') == "LOWER" and RM.norm_name('"intro"') == "INTRO" and RM.show_name("INTRO") == '"INTRO"'
    tr = dict(RM.tree(r))
    assert list(tr) == ['"MYDATA"', "Bitmap", "Cursor", "Icon", "String Table"], list(tr)
    assert [c for c, _ in tr["Icon"]] == ["200 [English (United States)]", "201"]
    assert [c for c, _ in tr["String Table"]] == ["String Table", "String Table [English (United States)]"]
    assert [c for c, _ in tr["Bitmap"]] == ["101", "102"]
    assert RE.renumber([102, 110], 300, True) == [402, 410] and RE.renumber([101, 103, 110], 200, False) == [200, 201, 202]
    assert LC.name(0x0408) == "Greek (Greece)" and LC.name(1112) == "1112" and LC.from_name("English (United States)") == 0x0409
    assert RE.changed_msg("Project1") == ("The resource file in Project1 has changed.  Click Yes to save the changes, "
                                          "or No to discard.")
    # the IDE: the add-in, the window, the string tables, the project
    w = MainWindow()
    w.show()
    oq, ow = QMessageBox.question, QMessageBox.warning
    asked = []
    QMessageBox.question = lambda *a, **k: QMessageBox.StandardButton.No
    QMessageBox.warning = lambda *a, **k: (asked.append(a[2] if len(a) > 2 else ""), QMessageBox.StandardButton.No)[1]
    try:
        w.res_loaded = False
        w.apply_addins()
        assert not w.a_res_editor.isVisible() and not w.a_new_res.isVisible()
        dlg = RE.AddInManagerDialog(w)
        row = [n for n, _ in RE.ADDINS].index(RE.ADDIN)
        dlg.list.setCurrentCell(row, 0)
        assert dlg.desc.toPlainText() == "Visual Basic 6 Resource Editor"
        dlg.cb_loaded.setChecked(True)
        assert dlg.list.item(row, 1).text() == "Loaded"
        dlg._ok()
        w.apply_addins()
        assert w.res_loaded and w.a_res_editor.isVisible() and w.a_new_res.isVisible()
        assert w.a_res_editor in w.tb_standard.actions()
        path = os.path.join(d, "s1.vbp")
        w.project.save(path)
        # the Project Explorer filled again (a Save, the Related Documents): the form stays selected
        from PyQt6.QtCore import Qt as _Qt
        m0 = w.project.modules[0]
        w.select_in_tree(m0)
        w.fill_tree()
        assert w.tree.currentItem() is not None and w.tree.currentItem().data(0, _Qt.ItemDataRole.UserRole) is m0
        ed = w.show_res_editor()
        assert ed.title() == "VB Resource Editor - " and ed.tree.topLevelItem(0).text(0) == "Resource File"
        assert not ed.acts["props"].isEnabled() or ed.selected() is None
        assert not ed.acts["undo"].isEnabled()
        ed.res.lang = 0x0408
        ed.edit_strings()
        st = ed.strings_win.widget()
        assert st.table.horizontalHeaderItem(1).text() == LC.name(0x0408) and st.table.item(0, 0).text() == "101"
        st.table.setCurrentCell(0, 1)
        st.table.item(0, 1).setText("Hello")
        st.enter_pressed()
        assert st.table.rowCount() == 2 and st.table.item(1, 0).text() == "102"
        st.table.item(1, 1).setText("World")
        assert ed.res.strings[0x0408] == {101: "Hello", 102: "World"} and ed.modified
        # Insert New String Table: a column at once, the Language list in its header, 1112 first (rs1 stage 4)
        assert LC.combo_items()[0] == (1112, "1112") and LC.combo_items()[len(LC.NUMERIC)][1] == LC.names()[0][1]
        st.insert_table()
        assert st.langs == [0x0408, 1112] and st.table.columnCount() == 3 and st.combo is not None
        assert not st.combo.isHidden() and st.combo.currentText() == "1112" and st.combo_lang == 1112
        st.pick_lang(0x0409)
        assert st.langs == [0x0408, 0x0409] and st.table.horizontalHeaderItem(2).text() == LC.name(0x0409)
        assert st.combo.currentData() == 0x0409 and ed.res.strings[0x0408] == {101: "Hello", 102: "World"}
        st.pick_lang(0x0408)                    # the language of the other table: nothing changes
        assert st.langs == [0x0408, 0x0409] and st.combo.currentData() == 0x0409
        st.undo()
        st.undo()
        assert st.langs == [0x0408] and st.combo.isHidden() and st.table.columnCount() == 2
        assert ed.res.strings[0x0408] == {101: "Hello", 102: "World"}
        st.table.setCurrentCell(1, 1)           # the table was filled again: back on the last row
        st.enter_pressed()                      # the empty row 103 is not written
        assert st.table.item(2, 0).text() == "103"
        w.res_sub.close()                       # the X hides the window, the changes stay
        assert w.res_sub.isHidden() and ed.modified
        w.show_res_editor()
        for kind, f in (("bitmap", "pic1.bmp"), ("icon", "icon1.ico"), ("cursor", "cur1.cur"), ("custom", "data1.txt")):
            assert ed.add_file(kind, os.path.join(RS1, "input", f)) is not None
        ed.save_as(os.path.join(d, "s1.res"))
        assert os.path.isfile(os.path.join(d, "s1.RES")) and ed.title().endswith("s1.RES")
        with open(os.path.join(d, "s1.RES"), "rb") as f:
            assert f.read() == _rs1("s1-stage3.RES")
        vbp = w.project.vbp.dumps() if hasattr(w.project.vbp, "dumps") else ""
        lines = [ln.split("=")[0] for ln in vbp.splitlines()]
        assert 'ResFile32="s1.RES"' in vbp and lines.index("ResFile32") < lines.index("Startup"), vbp
        assert lines.index("ResFile32") > max(i for i, k in enumerate(lines) if k in ("Form", "Reference"))
        texts = []
        for i in range(w.tree.topLevelItemCount()):
            root = w.tree.topLevelItem(i)
            for k in range(root.childCount()):
                texts.append((root.child(k).text(0), [root.child(k).child(j).text(0)
                                                      for j in range(root.child(k).childCount())]))
        assert ("Related Documents", ["(s1.RES)"]) in texts, texts
        # its menu, as VB6's: Properties | Add >, Save, Save As..., Remove s1.RES | Print... | Dockable, Hide
        from PyQt6.QtWidgets import QMenu as _QMenu
        mn = _QMenu()
        w.res_doc_menu(mn, w.project, "s1.RES")
        got = [a.text().replace("&", "") or "|" for a in mn.actions()]
        assert got == ["Properties", "|", "Add", "Save", "Save As...", "Remove s1.RES", "|", "Print...", "|",
                       "Dockable", "Hide"], got
        assert mn.actions()[2].menu() is not None and mn.actions()[-2].isChecked()
        sub = [(a.text().replace("&", ""), a.isEnabled()) for a in mn.actions()[2].menu().actions()]
        assert [t for t, _ in sub] == ["Form", "MDI Form", "Module", "Class Module", "User Control", "Property Page",
                                       "User Document", "WebClass", "Data Report", "DHTML Page", "Data Environment",
                                       "More ActiveX Designers...", "Add File...", "Resource File"], sub
        assert [t for t, on in sub if not on] == ["User Document", "WebClass", "DHTML Page",
                                                   "More ActiveX Designers..."], sub     # (MDI Form: 8.7A)
        # Delete and one Undo; Duplicate: the largest id + 1
        item = next(it for it in ed.res.items if it.kind == "bitmap")
        ed.select(item)
        ed.delete()
        assert item not in ed.res.items and ed.acts["undo"].isEnabled()
        ed.undo()
        assert item in ed.res.items and not ed.acts["undo"].isEnabled()
        ed.select(item)
        ed.duplicate()
        assert [it.name for it in ed.res.items if it.kind == "bitmap"] == [101, 102]
        # Edit Properties: names in capitals
        cust = next(it for it in ed.res.items if it.kind == "custom")
        ed.apply_properties(cust, {"type": RM.norm_name('"lower"'), "name": RM.norm_name('"intro"'), "lang": 0x0408})
        assert (cust.rtype, cust.name) == ("LOWER", "INTRO")
        # Renumber Rows in the table
        st = ed.strings_win.widget() if ed.strings_win is not None else None
        if st is None:
            ed.edit_strings()
            st = ed.strings_win.widget()
        st.table.clearSelection()
        st.table.setRangeSelected(__import__("PyQt6.QtWidgets", fromlist=["QTableWidgetSelectionRange"])
                                  .QTableWidgetSelectionRange(0, 0, 1, 0), True)
        assert st.acts["renumber"].isEnabled()
        st.renumber_rows(base=300, rebase=True)
        assert sorted(ed.res.strings[0x0408]) == [401, 402], ed.res.strings
        w.project.save()
        # a new project: the question of the Resource Editor (Yes / No), then an empty resource file
        asked.clear()
        w.new_project("Exe")
        assert asked == [RE.changed_msg("Project1")], asked
        assert ed.project is w.project and not ed.res.items and not ed.modified and ed.title() == "VB Resource Editor - "
        # the project again: its resource file; Remove s1.RES
        assert w.open_project(path)
        assert len(ed.res.items) == 4 and ed.res.strings[0x0408] == {101: "Hello", 102: "World"}, ed.res.strings
        w.remove_res_file(w.project)
        assert w.project.vbp.get("ResFile32") is None and not ed.res.items
        # Project > Add File with a .res
        w.attach_res_file(os.path.join(d, "s1.RES"))
        assert (w.project.vbp.get("ResFile32") or "").strip('"') == "s1.RES"
        # Project > Add New Resource File, a name that is not there (stage 6): the question, an empty file of
        # 32 bytes, the project's resource file without a question
        asked.clear()
        QMessageBox.warning = lambda *a, **k: (asked.append(a[2] if len(a) > 2 else ""), QMessageBox.StandardButton.Yes)[1]
        newp = w.add_new_resource_file(os.path.join(d, "s3.res"))
        assert newp.endswith("s3.RES") and asked == [RE.res_create_msg(newp)], asked
        assert RE.res_create_msg(newp) == "s3.RES\nThis file doesn't exist.\n\nCreate the file?"
        with open(newp, "rb") as f:
            assert f.read() == RM.EMPTY and len(RM.EMPTY) == 32
        assert (w.project.vbp.get("ResFile32") or "").strip('"') == "s3.RES" and ed.res.path == newp
        ed.modified = False
    finally:
        QMessageBox.question, QMessageBox.warning = oq, ow
    done(w)


def t_help87(d):
    """Phase 8.7: Help > Contents / Index / Search as the MSDN Library of VB6 (DEV-8/me3), on a small library made
    here (the .CHM of MSDN are not part of LinBASIC Studio): the Contents of VBCON98 + VBREF98 under Visual Basic
    Documentation, the pages of other .CHMs that the Reference names, the full-text Search (Found, Rank, AND /
    OR / NOT, titles only, previous results), Locate, Previous / Next"""
    import struct
    from lbstudio import msdn
    from lbstudio.helpwin import HelpWindow
    folder = os.path.join(d, "msdn")
    cache = os.path.join(d, "cache")
    os.makedirs(folder)
    old_cache = msdn.CACHE
    msdn.CACHE = cache

    def system(title, home):
        rec = b""
        for code, val in ((2, home), (3, title)):
            b = val.encode() + b"\0"
            rec += struct.pack("<HH", code, len(b)) + b
        return struct.pack("<I", 3) + rec

    def sitemap(entries):
        out = ['<HTML><BODY><OBJECT type="text/site properties"></OBJECT><UL>']
        for depth, name, local in entries:
            out.append(depth)
            if name:
                out.append(f'<LI> <OBJECT type="text/sitemap"><param name="Name" value="{name}">' +
                           (f'<param name="Local" value="{local}">' if local else "") + "</OBJECT>")
        out.append("</UL></BODY></HTML>")
        return "\n".join(out)

    def chm(name, title, home, toc, pages, ivb=None):
        open(os.path.join(folder, name.upper() + ".CHM"), "wb").close()
        pd = os.path.join(cache, name)
        idx = os.path.join(cache, name + ".index")
        os.makedirs(os.path.join(pd, "html"))
        os.makedirs(idx)
        for x in (pd, idx):
            open(os.path.join(x, ".done"), "w").close()
            with open(os.path.join(x, "#SYSTEM"), "wb") as f:
                f.write(system(title, home))
        with open(os.path.join(pd, name + ".hhc"), "w", encoding="cp1252") as f:
            f.write(toc)
        if ivb:                                 # the help context numbers of F1 (#IVB -> #STRINGS)
            strings, table = b"\0", b""
            for cid, local in ivb.items():
                table += struct.pack("<II", cid, len(strings))
                strings += local.encode() + b"\0"
            with open(os.path.join(idx, "#IVB"), "wb") as f:
                f.write(struct.pack("<I", len(table)) + table)
            with open(os.path.join(idx, "#STRINGS"), "wb") as f:
                f.write(strings)
        for local, (ptitle, body) in pages.items():
            with open(os.path.join(pd, local), "w", encoding="cp1252") as f:
                f.write(f"<html><head><title>{ptitle}</title></head><body>{body}</body></html>")
    try:
        chm("vbcon98", "Visual Basic Concepts", "html/vbstartpage.htm",
            sitemap([("", "Visual Basic Start Page", "html/vbstartpage.htm"),
                     ("", "Using Visual Basic", None), ("<UL>", "Programmer's Guide", "html/guide.htm"),
                     ("</UL>", "", None)]),
            {"html/vbstartpage.htm": ("Visual Basic Start Page", "Welcome. MsgBox shows a message."),
             "html/guide.htm": ("Programmer's Guide", "Forms and controls. MsgBox, MsgBox and InputBox.")})
        chm("vbref98", "Visual Basic Reference", "html/ref.htm",
            sitemap([("", "Reference", "html/ref.htm"), ("<UL>", "Language Reference", None),
                     ("<UL>", "MsgBox Function", "vb98.chm::/html/vafctMsgBox.htm"),
                     ("", "InputBox Function", "vb98.chm::/html/vafctInputBox.htm"), ("</UL>", "", None),
                     ("</UL>", "", None)]),
            {"html/ref.htm": ("Reference", "The reference of the language.")})
        chm("vb98", "Visual Basic Reference", "html/vafctMsgBox.htm", sitemap([("", "MsgBox Function",
                                                                                "html/vafctMsgBox.htm")]),
            {"html/vafctMsgBox.htm": ("MsgBox Function", "MsgBox MsgBox MsgBox: displays a message. msgboxes"),
             "html/vafctInputBox.htm": ("InputBox Function", "InputBox displays a prompt <script>MsgBox</script>")},
            ivb={1008978: "html/vafctMsgBox.htm"})
        chm("xtra", "Extra Book", "html/x.htm", sitemap([("", "Extra Control", "html/x.htm")]),
            {"html/x.htm": ("Extra Control", '<meta http-equiv="Content-Type" content="text/html; '
                                             'charset=iso8859-1">We\u2019ve nothing here.')})
        for n, kids in (("defa", ["beta", "alpha"]), ("defb", ["gamma", "alpha"])):   # two books of one name
            chm(n, "Visual Basic Reference", "html/a.htm",
                sitemap([("", "Definitions", None), ("<UL>", kids[0], "html/a.htm")] +
                        [("", k, "html/a.htm") for k in kids[1:]] + [("</UL>", "", None)]),
                {"html/a.htm": ("A definition", "One word.")})
        lib = msdn.Library(folder)
        # Contents: the root, Visual Basic Documentation (VBCON98 + VBREF98), the .CHMs nobody names
        root = lib.contents()[0]
        assert root.name == "MSDN Library Visual Studio 6.0" and lib.resolve(root)[1] == "html/vbstartpage.htm"
        names = [n.name for n in root.children]
        assert names == ["Visual Basic Documentation", "Definitions", "Extra Control"], names
        assert [n.name for n in root.children[1].children] == ["alpha", "beta", "gamma"]       # merged, sorted
        vbdoc = root.children[0]
        assert [n.name for n in vbdoc.children] == ["Visual Basic Start Page", "Using Visual Basic", "Reference"]
        assert [n.name for n in vbdoc.children[1].children] == ["Programmer's Guide"]
        lang = vbdoc.children[2].children[0]
        assert [n.name for n in lang.children] == ["MsgBox Function", "InputBox Function"]
        f, local = lib.resolve(lang.children[0])
        assert f.name == "vb98" and local == "html/vafctMsgBox.htm" and f.title == "Visual Basic Reference"
        # Search: AND, rank by count, titles only, NOT, OR, similar words, previous results
        found = lib.search("MsgBox")
        assert [t for t, _f, _l, _r in found] == ["MsgBox Function", "Programmer's Guide", "Visual Basic Start Page"]
        assert [r for *_x, r in found] == [1, 2, 3] and found[0][1].name == "vb98"
        assert [t for t, *_x in lib.search("MsgBox", similar=False)] == \
            ["MsgBox Function", "Programmer's Guide", "Visual Basic Start Page"]
        assert [t for t, *_x in lib.search("msgbox", titles_only=True)] == ["MsgBox Function"]
        assert [t for t, *_x in lib.search("MsgBox NOT forms")] == ["MsgBox Function", "Visual Basic Start Page"]
        assert {t for t, *_x in lib.search("InputBox OR welcome")} == \
            {"InputBox Function", "Programmer's Guide", "Visual Basic Start Page"}
        assert [t for t, *_x in lib.search("prompt")] == ["InputBox Function"]      # not the words of a <script>
        assert [t for t, *_x in lib.search("message", within=lib.search("welcome"))] == ["Visual Basic Start Page"]
        assert os.path.exists(lib.fts_path())
        lib2 = msdn.Library(folder)                                                   # the index from the cache
        assert [t for t, *_x in lib2.search("MsgBox")][0] == "MsgBox Function"
        # the window: the three tabs, Contents, Locate, Previous / Next, List Topics
        w = MainWindow()
        # Help > Set Help Folder...: a folder with .CHM files only (the same setting as Tools > Options)
        helpm = next(a.menu() for a in w.menuBar().actions() if a.text() == "&Help")
        assert "Set Help &Folder..." in [a.text() for a in helpm.actions()]
        os.makedirs(os.path.join(d, "empty"))
        before = w.S["msdn_path"]
        assert not w.set_help_folder(os.path.join(d, "empty")) and w.S["msdn_path"] == before
        get_dir = QFileDialog.getExistingDirectory
        QFileDialog.getExistingDirectory = lambda *a, **k: folder
        try:
            assert w.set_help_folder() and w.S["msdn_path"] == folder
        finally:
            QFileDialog.getExistingDirectory = get_dir
        hw = HelpWindow(w, lib)
        assert [hw.tabs.tabText(i) for i in range(3)] == ["&Contents", "I&ndex", "&Search"]
        tools = [a.text() for a in hw.findChildren(QToolBar)[0].actions()]
        assert tools == ["Hide", "Locate", "Previous", "Next", "Back", "Forward", "Stop", "Refresh", "Home",
                         "Print"], tools
        w._helpwin = hw
        w.help_window = lambda: hw
        w.help_contents()
        app.processEvents()
        assert hw.tabs.currentIndex() == 0 and hw.tree.topLevelItem(0).text(0) == "MSDN Library Visual Studio 6.0"
        assert hw.current[1] == "html/vbstartpage.htm"
        top = hw.tree.topLevelItem(0)
        assert top.isExpanded() and top.child(0).text(0) == "Visual Basic Documentation"
        hw.show_page(f, local)                                    # the MsgBox page, then Locate
        hw.locate()
        assert hw.tree.currentItem().text(0) == "MsgBox Function"
        assert hw.tree.currentItem().parent().text(0) == "Language Reference"
        hw.step_toc(1)
        assert hw.current[1] == "html/vafctInputBox.htm"
        hw.step_toc(-1)
        assert hw.current[1] == "html/vafctMsgBox.htm"
        w.help_contents(search=True)
        assert hw.tabs.currentIndex() == 2 and hw.similar.isChecked() and not hw.titles_only.isChecked()
        hw.words.setEditText("MsgBox")
        hw.list_topics()
        assert hw.found_label.text() == "Found: 3" and hw.found.topLevelItemCount() == 3
        first = hw.found.topLevelItem(0)
        assert [first.text(i) for i in range(3)] == ["MsgBox Function", "Visual Basic Reference", "1"]
        hw.found.setCurrentItem(hw.found.topLevelItem(1))
        hw._display_found()
        assert hw.current[1] == "html/guide.htm"
        w.help_contents(index=True)
        assert hw.tabs.currentIndex() == 1
        # F1 in the code window: MsgBox -> its help context number (1008978 of VBA) -> the page of the .CHM
        cw = w.open_code(w.project.modules[0])
        cw.editor.setText('Private Sub Form_Load()\r\n    MsgBox "x"\r\nEnd Sub\r\n')
        cw.editor.SendScintilla(S.SCI_GOTOPOS, cw.editor.text().index("MsgBox") + 2)
        hw.show_page(*lib.resolve(vbdoc.children[0]))
        from lbstudio import mainwindow as _mw
        focus = _mw.QApplication.focusWidget
        _mw.QApplication.focusWidget = staticmethod(lambda: cw.editor)
        try:
            w.context_help()
        finally:
            _mw.QApplication.focusWidget = focus
        assert hw.current[0].name == "vb98" and hw.current[1] == "html/vafctMsgBox.htm", hw.current
        assert w.help_context_number(1008978, "VBA") and not w.help_context_number(4242, "VBA")
        hw.show_page(lib.file("xtra.chm"), "html/x.htm")         # charset=iso8859-1 is read as cp1252 (’)
        assert "We\u2019ve nothing here." in hw.view.toPlainText(), hw.view.toPlainText()
        # a dark theme of the IDE: the page stays black on white (the <pre> code of MSDN has no colour of its own)
        from PyQt6.QtGui import QPalette as _P
        mode = w.S["dark_mode"]
        w.S["dark_mode"] = "dark"
        w.apply_theme()
        app.processEvents()
        pal = hw.view.palette()
        assert pal.color(_P.ColorRole.Text).name() == "#000000" and pal.color(_P.ColorRole.Base).name() == "#ffffff"
        w.S["dark_mode"] = mode
        w.apply_theme()
        hw.toggle_side()
        assert hw.tabs.isHidden() and hw.a_hide.text() == "Show"
        hw.close()
        done(w)
    finally:
        msdn.CACHE = old_cache


def t_dock87(d):
    """Phase 8.7C: Dockable (DEV-8/me4): Tools > Options in VB6's tabs with Docking (the nine windows, the Object
    Browser not dockable at first), a window without Dockable floats and no side takes it, Immediate / Locals /
    Watches leave the tabs of the bottom dock and come back, the Object Browser docks when it is dockable, the
    right button of the tool windows (Dockable, Hide; the Immediate window's own commands first)"""
    from PyQt6.QtWidgets import QDockWidget
    from lbstudio.dialogs import OptionsDialog
    w = MainWindow()
    w.show()
    app.processEvents()
    dlg = OptionsDialog(w)
    assert [dlg.tabs.tabText(i) for i in range(dlg.tabs.count())] == ["Editor", "General", "Docking", "Environment"]
    assert [(k, cb.text()) for k, cb in dlg.dock_boxes.items()] == [
        ("immediate", "Immediate Window"), ("locals", "Locals Window"), ("watch", "Watch Window"),
        ("project", "Project Explorer"), ("props", "Properties Window"), ("objbrowser", "Object Browser"),
        ("layout", "Form Layout"), ("toolbox", "Toolbox"), ("palette", "Color Palette")]
    assert [cb.isChecked() for cb in dlg.dock_boxes.values()] == [True] * 5 + [False] + [True] * 3
    dlg.dock_boxes["props"].setChecked(False)
    dlg.dock_boxes["immediate"].setChecked(False)
    dlg._ok()
    w.apply_dockable()
    app.processEvents()
    NoArea = Qt.DockWidgetArea.NoDockWidgetArea
    assert w.dock_props.allowedAreas() == NoArea and w.dock_props.isFloating()
    assert "immediate" in w.free_docks and w.bottom_tabs.indexOf(w.immediate) == -1
    free = w.free_docks["immediate"]
    assert free.widget() is w.immediate and free.isFloating() and free.allowedAreas() == NoArea
    free.hide()
    w._show_dock(w.dock_immediate, w.immediate)                  # View > Immediate Window: its own window
    assert free.isVisible()
    w.set_dockable("immediate", True)                            # back: the second tab, after Output
    assert not w.free_docks and w.bottom_tabs.indexOf(w.immediate) == 1
    assert [w.bottom_tabs.tabText(i) for i in range(4)] == ["Output", "Immediate", "Locals", "Watches"]
    w.set_dockable("watch", False)
    assert w.bottom_tabs.indexOf(w.watch_win) == -1 and w.free_docks["watch"].windowTitle() == "Watches"
    w.set_dockable("watch", True)
    assert w.bottom_tabs.indexOf(w.watch_win) == 3
    w.set_dockable("props", True)
    assert w.dock_props.allowedAreas() == Qt.DockWidgetArea.AllDockWidgetAreas
    # the Object Browser: an MDI window, or a dock when it is dockable
    ob = w.object_browser()
    assert not isinstance(w._objbrowser, QDockWidget) and ob is w._objbrowser.widget()
    w.set_dockable("objbrowser", True)
    assert isinstance(w._objbrowser, QDockWidget) and w._objbrowser.isVisible()
    w.set_dockable("objbrowser", False)
    assert not isinstance(w._objbrowser, QDockWidget)
    w._close_objbrowser()
    # the right button: Dockable / Hide; the Immediate window: Cut, Copy, Paste | Object Browser, Definition |
    # Start, End first (DEV-8/me4 pictures 2, 3)
    for wd in (w.props.grid, w.immediate, w.locals_win, w.watch_win, w.toolbox, w.palette_win):
        assert wd.contextMenuPolicy() == Qt.ContextMenuPolicy.CustomContextMenu, wd
    items = w._immediate_menu()
    assert [e[0] if e else None for e in items] == ["Cu&t", "&Copy", "&Paste", None, "&Object Browser",
                                                     "&Definition", None, "&Start", "&End"]
    assert items[7][2] and not items[8][2]                      # design time: Start, not End
    from PyQt6.QtCore import QTimer as _QT, QPoint as _QP
    seen = {}

    def grab():
        m = w._last_tool_menu
        seen["items"] = [a.text() for a in m.actions() if not a.isSeparator()]
        seen["dockable"] = [a.isChecked() for a in m.actions() if a.text() == "Doc&kable"]
        m.close()
    _QT.singleShot(0, grab)
    w.props.grid.customContextMenuRequested.emit(_QP(5, 5))
    assert seen == {"items": ["&Description", "Doc&kable", "&Hide"], "dockable": [True]}, seen     # (me8-03)
    w.hide_tool("props", w.props.grid)
    assert w.dock_props.isHidden()
    w.dock_props.show()                         # the state of the window is kept for the next tests
    w.S["dockable"] = {}
    w.S["layout_dockable"] = True
    done(w)



def t_pe87(d):
    """the Project Explorer as VB6's (phase 8.7D, VB6 oracle DEV-8/me5): the names of what is not saved yet, the
    buttons, the menus of a form, a module, a folder and the project, Properties of the project, Toggle Folders,
    a project opened with its folders closed"""
    from PyQt6.QtWidgets import QMenu as _QMenu
    w = MainWindow()
    w.show()
    app.processEvents()
    w.res_loaded = False                        # (Add > Resource File: only with the Resource Editor loaded)
    w.new_project("Exe")
    w.add_module("Module")
    w.add_module("Class")
    p = w.project

    def kids(it):
        return [it.child(i) for i in range(it.childCount())]

    def label(it):
        return it.text(0).replace(" *", "")

    def item(text):
        stack = [w.tree.topLevelItem(i) for i in range(w.tree.topLevelItemCount())]
        while stack:
            x = stack.pop()
            if label(x) == text:
                return x
            stack += kids(x)
        raise AssertionError(text)

    def menu(it):
        w.tree.setCurrentItem(it)
        mn = _QMenu()
        w.pe_menu(mn, it)
        return mn, [a.text().replace("&", "") or "|" for a in mn.actions()]

    # me5-01: unsaved, the name in the brackets; the folders; "Project - Project1"
    root = w.tree.topLevelItem(0)
    assert label(root) == "Project1 (Project1)" and root.font(0).bold() and w.dock_project.windowTitle() == "Project - Project1"
    tree = [(label(f), [label(c) for c in kids(f)]) for f in kids(root)]
    assert tree == [("Forms", ["Form1 (Form1)"]), ("Modules", ["Module1 (Module1)"]),
                    ("Class Modules", ["Class1 (Class1)"])], tree
    assert [a.toolTip() for a in (w.a_pe_code, w.a_pe_object, w.a_pe_folders)] == ["View Code", "View Object",
                                                                                     "Toggle Folders"]
    # the buttons: a form both, a class View Code only, the project / a folder none (me5-02..07, 11)
    for text, code_on, obj_on in (("Form1 (Form1)", True, True), ("Class1 (Class1)", True, False),
                                  ("Module1 (Module1)", True, False), ("Project1 (Project1)", False, False),
                                  ("Forms", False, False)):
        w.tree.setCurrentItem(item(text))
        assert (w.a_pe_code.isEnabled(), w.a_pe_object.isEnabled()) == (code_on, obj_on), text
    # the menus (me5-05..09)
    mn, got = menu(item("Form1 (Form1)"))
    assert got == ["View Object", "View Code", "|", "Properties", "|", "Add", "Save Form1", "Save Form1 As...",
                   "Remove Form1", "|", "Print...", "|", "Dockable", "Hide"], got
    raw = [a.text() for a in mn.actions()]
    assert raw[:4] == ["View O&bject", "View C&ode", "", "Prope&rties"] and raw[6:9] == \
        ["&Save Form1", "Save Form1 &As...", "&Remove Form1"], raw
    assert all(not a.icon().isNull() for a in mn.actions() if a.text().replace("&", "") in
               ("View Object", "View Code", "Properties", "Print..."))
    _, got = menu(item("Class1 (Class1)"))
    assert got == ["View Code", "|", "Properties", "|", "Add", "Save Class1", "Save Class1 As...", "Remove Class1",
                   "|", "Print...", "|", "Dockable", "Hide"], got
    mn, got = menu(item("Forms"))
    assert got == ["View Code", "View Object", "|", "Project1 Properties...", "|", "Add", "Save", "Save As...",
                   "Print...", "|", "Dockable", "Hide"], got
    assert [a.text().replace("&", "") for a in mn.actions() if not a.isEnabled()] == ["View Code", "View Object",
                                                                                       "Save", "Save As..."]
    assert mn.actions()[3].text() == "Project1 Prop&erties..."
    mn, got = menu(item("Project1 (Project1)"))
    assert got == ["Set as Start Up", "Project1 Properties...", "|", "Save Project", "Remove Project", "|", "Add", "|",
                   "Print...", "|", "Dockable", "Hide"], got
    assert mn.actions()[0].isEnabled() and mn.actions()[0].text() == "Set as Start &Up"
    assert mn.actions()[3].text() == "Sa&ve Project" and mn.actions()[-2].isChecked()
    add = next(a for a in mn.actions() if a.text() == "&Add").menu()
    sub = [(a.text().replace("&", ""), a.isEnabled()) for a in add.actions()]
    assert [t for t, _ in sub] == ["Form", "MDI Form", "Module", "Class Module", "User Control", "Property Page",
                                   "User Document", "WebClass", "Data Report", "DHTML Page", "Data Environment",
                                   "More ActiveX Designers...", "Add File..."], sub          # (me5-10)
    assert [t for t, on in sub if not on] == ["User Document", "WebClass", "DHTML Page", "More ActiveX Designers..."]
    assert all(not a.icon().isNull() for a in add.actions()[:7])
    # Properties of the project: "Project1 Project", (Name) only; a new name goes to the .vbp (me5-11)
    w.tree.setCurrentItem(item("Project1 (Project1)"))
    assert w.props.project is p and w.props.obj.currentText() == "Project1 Project"
    assert [r.name for r in w.props.rows] == ["(Name)"] and w.dock_props.windowTitle() == "Properties - Project1"
    assert w.rename_project(p, "1x") == "Invalid property value" and p.name == "Project1"
    w.props.commit(w.props.rows[0], "MyProj")
    assert p.name == "MyProj" and 'Name="MyProj"' in p.vbp.dumps().splitlines() and label(w.tree.topLevelItem(0)) == "MyProj (MyProj)"
    assert w.dock_project.windowTitle() == "Project - MyProj" and w.props.obj.currentText() == "MyProj Project"
    # a folder: nothing; a form with no window open: nothing; a module: (Name) (me5-01, 08, 14)
    w.tree.setCurrentItem(item("Forms"))
    assert w.props.project is None and w.props.module is None and w.dock_props.windowTitle() == "Properties"
    w.close_all_code()
    for _ in range(3):
        app.processEvents()
    w.tree.setCurrentItem(item("Forms"))        # (closing the windows may have selected Form1 already)
    w.tree.setCurrentItem(item("Form1 (Form1)"))
    assert w.props.module is None and w.props.rows == [], (w.props.module, w.code_windows, w.design_windows)
    w.tree.setCurrentItem(item("Module1 (Module1)"))
    assert w.props.module is p.module("Module1") and [r.name for r in w.props.rows] == ["(Name)"]
    assert w.dock_props.windowTitle() == "Properties - Module1"
    # Toggle Folders: the modules by name under the project; again: the folders, closed (me5-12, 13)
    w.toggle_folders()
    root = w.tree.topLevelItem(0)
    assert not w.a_pe_folders.isChecked() and [label(c) for c in kids(root)] == \
        ["Class1 (Class1)", "Form1 (Form1)", "Module1 (Module1)"], [label(c) for c in kids(root)]
    w.toggle_folders()
    root = w.tree.topLevelItem(0)
    assert w.a_pe_folders.isChecked() and [label(c) for c in kids(root)] == ["Forms", "Modules", "Class Modules"]
    assert root.isExpanded() and not any(c.isExpanded() for c in kids(root))
    # a folder the user opens stays open when the tree is filled again
    item("Modules").setExpanded(True)
    w.fill_tree()
    assert item("Modules").isExpanded() and not item("Forms").isExpanded()
    # saved: the files in the brackets; opened again: the folders closed, the project selected, Properties
    # "MyProj Project" (rs1 stage 8)
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Module1.bas"), os.path.join(d, "Class1.cls"),
          os.path.join(d, "MyProj.vbp"))
    assert w.save_project_as()
    assert label(item("Form1 (Form1.frm)")) == "Form1 (Form1.frm)" and label(w.tree.topLevelItem(0)) == \
        "MyProj (MyProj.vbp)"
    done(w)
    w.new_project("Exe")
    done(w)
    assert w.open_project(os.path.join(d, "MyProj.vbp"))
    for _ in range(5):
        app.processEvents()
    root = w.tree.topLevelItem(0)
    assert w.tree.currentItem() is root and root.isExpanded() and not any(c.isExpanded() for c in kids(root))
    assert w.props.project is w.project and w.props.obj.currentText() == "MyProj Project"
    assert w.current_designer() is not None and w.current_designer().module.name == "Form1"   # (Startup)
    # a group: the same menu for both projects, Set as Start Up on also for the start-up one (me5-17, 18)
    w.add_project("Exe")
    _, got1 = menu(item("MyProj (MyProj.vbp)"))
    _, got2 = menu(item("Project1 (Project1)"))
    want = ["Set as Start Up", "{} Properties...", "|", "Save Project", "Remove Project", "|", "Add", "|", "Print...",
            "|", "Dockable", "Hide"]
    assert got1 == [x.format("MyProj") for x in want] and got2 == [x.format("Project1") for x in want], (got1, got2)
    assert w.dock_project.windowTitle().startswith("Project Group - ")
    done(w)



def t_tb87(d):
    """the Toolbox as VB6's (phase 8.7E, VB6 oracles DEV-8/me6, me7): the tab General and the user's tabs, their
    menus, a new tab right after the open one, any name (also one there is, "General" or none), a tool dragged
    onto a tab, Delete Tab gives the tools back, the tabs stay (settings), Components puts the new controls into
    the open tab"""
    from PyQt6.QtCore import QTimer as _QT, QPoint as _QP
    from lbstudio.designer import Toolbox
    w = MainWindow()
    w.show()
    app.processEvents()
    tb = w.toolbox
    for t in list(tb.tabs):                     # (a clean start)
        tb.delete_tab(t)
    seen = {}

    def grab():
        m = w._last_tool_menu
        seen["items"] = [a.text() for a in m.actions()]
        m.close()
    # the menu of a tool and of the empty space: Components..., Add Tab... | Dockable, Hide (me6-01, 02)
    _QT.singleShot(0, grab)
    tb.customContextMenuRequested.emit(_QP(5, 5))
    assert seen["items"] == ["C&omponents...", "&Add Tab...", "", "Doc&kable", "&Hide"], seen
    # the menu of General: only Add Tab... (me6-03, 08)
    m = tb.tab_menu(None)
    assert [a.text() for a in m.actions()] == ["&Add Tab...", "&Delete Tab", "&Rename Tab...", "", "Move &Up",
                                               "Move Do&wn"], [a.text() for a in m.actions()]

    def on(t):
        return [a.isEnabled() for a in tb.tab_menu(t).actions() if a.text()]
    assert on(None) == [True, False, False, False, False]

    def names(lay):
        return [lay.itemAt(i).widget().text() for i in range(lay.count())]
    general = [c for c, _ in tb.entries]
    # Add Tab: the tab goes right after the open one, which stays open (me6-05, 06, 09)
    my = tb.add_tab("MyTab")
    assert tb.names() == ["MyTab"] and tb.cur is None and tb.cur_tab == "General"
    assert names(tb.top) == ["General"] and names(tb.bottom) == ["MyTab"]
    tb.set_tab(my)
    assert names(tb.top) == ["General", "MyTab"] and names(tb.bottom) == [] and tb.tab_tools(my) == [None]
    assert w.dock_toolbox.windowTitle() == "MyTab"
    t2 = tb.add_tab("Tab2")
    assert tb.cur is my and tb.names() == ["MyTab", "Tab2"] and names(tb.bottom) == ["Tab2"]
    assert on(my) == [True, True, True, False, True] and on(t2) == [True, True, True, True, False]
    tb.move_tab(t2, -1)
    assert tb.names() == ["Tab2", "MyTab"]
    tb.move_tab(t2, 1)
    assert tb.names() == ["MyTab", "Tab2"]
    tb.rename_tab(t2, "Tools2")                 # (me6-12)
    assert tb.names() == ["MyTab", "Tools2"] and names(tb.bottom) == ["Tools2"]
    tb.delete_tab(t2)                           # empty: no question (me6-13)
    assert tb.names() == ["MyTab"] and tb.cur is None
    # any name: "General" again, none; after the open tab (me7-04, 05, 06)
    g2 = tb.add_tab("General")
    assert tb.names() == ["General", "MyTab"] and tb.tab("General") is None and names(tb.bottom) == ["General", "MyTab"]
    tb.set_tab(my)
    blank = tb.add_tab("")
    assert tb.names() == ["General", "MyTab", ""] and names(tb.bottom) == [""]
    tb.set_tab(None)
    zed = tb.add_tab("Zed")
    assert tb.names() == ["Zed", "General", "MyTab", ""]
    my2 = tb.add_tab("MyTab")                   # a name there is
    assert tb.names() == ["MyTab", "Zed", "General", "MyTab", ""] and my2 is not my
    for t in (g2, blank, zed, my2):
        tb.delete_tab(t)
    assert tb.names() == ["MyTab"]
    # a tool dragged onto MyTab moves there and MyTab opens (me6-14)
    tb.move_tool("Timer", my)
    assert tb.cur is my and tb.tab_tools(my) == [None, "Timer"] and "Timer" not in tb.tab_tools(None)
    assert not tb.buttons["Timer"].isHidden() and tb.buttons["CommandButton"].isHidden()
    # the tabs stay after the IDE is closed: LinBASIC Studio's settings, not the project's (me6-17, 18)
    assert w.S["toolbox_tabs"] == [{"name": "MyTab", "tools": ["Timer"]}], w.S["toolbox_tabs"]
    tb2 = Toolbox(w)
    assert tb2.names() == ["MyTab"] and tb2.tab_of["Timer"] is tb2.tabs[0] and tb2.cur is None
    tb2.deleteLater()
    tb.fill()                                   # (the second Toolbox named the dock)
    # Delete Tab with tools: no question, General opens, the Timer is back in its place (me6-16)
    tb.delete_tab(my)
    assert tb.tab_tools(None) == general and tb.cur is None and w.S["toolbox_tabs"] == []
    # Project > Components: the new controls go into the open tab (me6-19)
    saved = tb.add_tab("Saved")
    tb.set_tab(saved)
    keys = list(tb.components)
    tb.adding = True
    tb.set_components(keys + ["MSComDlg"])
    tb.adding = False
    assert tb.tab_tools(saved) == [None, "MSComDlg.CommonDialog"], tb.tab_tools(saved)
    assert "MSComDlg.CommonDialog" not in tb.tab_tools(None)
    tb.delete_tab(saved)
    tb.set_components(keys)
    assert w.S["toolbox_tabs"] == [] and tb.tab_tools(None) == general
    done(w)


def t_zoom87(d):
    """View > Zoom (user, 2026-11-18; a touchpad has no Ctrl+Mouse Wheel): Zoom In / Zoom Out / Zoom Reset with
    their keys; every code window and both panes of a split one take the same zoom, Ctrl+Mouse Wheel in one pane
    too; the zoom is kept in the settings (code_zoom) and a new code window opens with it"""
    settings.S["code_zoom"] = 0
    w = MainWindow()
    w.show()
    app.processEvents()
    view = {t: a for t, _, a in _menu_items(next(a.menu() for a in w.menuBar().actions() if a.text() == "&View"))}
    items = [(t, k) for t, k, _ in _menu_items(view["&Zoom"].menu())]
    assert items == [("Zoom &In", "Ctrl+="), ("Zoom &Out", "Ctrl+-"), ("Zoom &Reset", "Ctrl+0")], items
    w.view_code()
    cw = w.current_cw()

    def zoom(e):
        return e.SendScintilla(2374)            # SCI_GETZOOM
    w.a_zoom_in.trigger()
    w.a_zoom_in.trigger()
    assert zoom(cw.editor) == 2 and settings.S["code_zoom"] == 2
    w.a_zoom_out.trigger()
    assert zoom(cw.editor) == 1 and settings.S["code_zoom"] == 1
    cw.set_split(True)
    assert [zoom(e) for e in cw.panes()] == [1, 1]
    cw.panes()[1].SendScintilla(2373, 4)        # SCI_SETZOOM, as Ctrl+Mouse Wheel does: the other pane follows
    assert [zoom(e) for e in cw.panes()] == [4, 4] and settings.S["code_zoom"] == 4
    cw.set_split(False)
    w.add_module("Module")                      # a new code window: the same zoom
    cw2 = w.current_cw()
    assert cw2 is not cw and zoom(cw2.editor) == 4
    w.a_zoom_reset.trigger()
    assert zoom(cw.editor) == 0 and zoom(cw2.editor) == 0 and settings.S["code_zoom"] == 0
    for _ in range(40):
        w.a_zoom_out.trigger()
    assert zoom(cw2.editor) == -10 and settings.S["code_zoom"] == -10       # the smallest zoom of Scintilla
    w.a_zoom_reset.trigger()
    done(w)


def t_ctxmenu87(d):
    """the menus of the right button (phase 8.7F, VB6 oracle DEV-8/me8): the form and its controls, a user control,
    Paste of controls / of a picture, Properties > Description, Watches, Locals, the Object Browser, Window > More
    Windows..."""
    from PyQt6.QtCore import QTimer as _QT
    from PyQt6.QtGui import QImage, QColor as _QC
    w = MainWindow()
    w.show()
    app.processEvents()
    p = w.project
    f = w.open_designer(p.module("Form1"))

    def texts(items):
        return [e[1] if e else None for e in items]

    def on(items):
        return {e[1]: bool(e[3]) for e in items if e}
    cb = QApplication.clipboard()
    cb.setText("abc")                           # only text: Paste grey (me8-05)
    f.select([])
    it = w.designer_menu_items(f, None)
    assert texts(it) == ["View C&ode", "&Menu Editor...", None, "&Lock Controls", "&Paste", None, "Prope&rties", None,
                         "Update &UserControls"], texts(it)
    assert on(it) == {"View C&ode": True, "&Menu Editor...": True, "&Lock Controls": True, "&Paste": False,
                      "Prope&rties": True, "Update &UserControls": False}, on(it)
    assert [e[0] for e in it if e][:4] == ["view_code", "menu_editor", "lock", "paste"]
    c1 = f.place_new("CommandButton", QRect(8, 8, 0, 0), None)
    i1 = f.items[c1]
    f.select([i1])
    app.processEvents()                          # the title of Properties: the selected control (me8-02)
    assert w.dock_props.windowTitle() == "Properties - Command1", w.dock_props.windowTitle()
    it = w.designer_menu_items(f, i1)            # one control: Bring / Send grey (me8-02)
    assert texts(it) == ["Cu&t", "&Copy", "&Paste", "&Delete", None, "&Bring to Front", "&Send to Back", None,
                         "View C&ode", "Align to &Grid", None, "Prope&rties"], texts(it)
    o = on(it)
    assert not o["&Paste"] and not o["&Bring to Front"] and not o["&Send to Back"] and o["&Delete"], o
    c2 = f.place_new("TextBox", QRect(150, 120, 0, 0), None)
    i1, i2 = f.items[c1], f.items[c2]
    f.select([i1])
    o = on(w.designer_menu_items(f, i1))         # Text1 came later, in front of Command1 (me8-06)
    assert o["&Bring to Front"] and not o["&Send to Back"], o
    f.select([i1, i2])                           # both: nothing to change (me8-07)
    o = on(w.designer_menu_items(f, i2))
    assert not o["&Bring to Front"] and not o["&Send to Back"], o
    f.select([i1])
    f.copy_selection()                           # controls copied: Paste everywhere (me8-08, 09, 10)
    assert on(w.designer_menu_items(f, i1))["&Paste"]
    c3 = f.place_new("Frame", QRect(150, 8, 0, 0), None)
    i1, i3 = f.items[c1], f.items[c3]
    f.select([i3])
    o = on(w.designer_menu_items(f, i3))
    assert o["&Paste"] and not o["&Bring to Front"] and o["&Send to Back"], o
    f.select([])
    assert on(w.designer_menu_items(f, None))["&Paste"]
    # the menu by the right button on the surface and on a control
    seen = {}

    def grab():
        m = w._last_designer_menu
        seen.setdefault("menus", []).append([a.text() for a in m.actions() if not a.isSeparator()])
        m.close()
    _QT.singleShot(0, grab)
    QTest.mousePress(f.surface, Qt.MouseButton.RightButton, Qt.KeyboardModifier.NoModifier, QPoint(300, 250))
    i1 = f.items[c1]                             # (a new control builds the items again)
    _QT.singleShot(0, grab)
    QTest.mousePress(i1, Qt.MouseButton.RightButton, Qt.KeyboardModifier.NoModifier, QPoint(3, 3))
    assert seen["menus"][0][0] == "View C&ode" and seen["menus"][1][0] == "Cu&t", seen
    assert f.selection == [i1]
    # a picture on the clipboard: Paste on the form only, into its Picture (me8-01, 02, 16)
    img = QImage(24, 16, QImage.Format.Format_RGB32)
    img.fill(_QC(255, 0, 0))
    cb.setImage(img)
    app.processEvents()
    assert not on(w.designer_menu_items(f, i1))["&Paste"]
    f.select([])
    assert on(w.designer_menu_items(f, None))["&Paste"]
    f.paste()
    data = f.module.frx_value(f.root, "Picture")
    assert data and bytes(data)[:2] == b"BM", data[:8] if data else data
    # ... a PictureBox (its size stays) and an Image (the size of the picture) take it too (me9-02..05)
    cp = f.place_new("PictureBox", QRect(20, 150, 0, 0), None)
    f.select([f.items[cp]])
    assert on(w.designer_menu_items(f, f.items[cp]))["&Paste"]
    size = (int(cp.get("Width")), int(cp.get("Height")))
    f.paste()
    assert bytes(f.module.frx_value(cp, "Picture") or b"")[:2] == b"BM"
    assert (int(cp.get("Width")), int(cp.get("Height"))) == size
    ci = f.place_new("Image", QRect(200, 150, 0, 0), None)
    f.select([f.items[ci]])
    assert on(w.designer_menu_items(f, f.items[ci]))["&Paste"]
    f.paste()
    assert bytes(f.module.frx_value(ci, "Picture") or b"")[:2] == b"BM"
    assert (int(ci.get("Width")), int(ci.get("Height"))) == (24 * 15, 16 * 15), (ci.get("Width"), ci.get("Height"))
    # a user control on the form: Update UserControls; its own menu without pictures or Properties (me8-12, 13)
    w.add_module("UserControl")
    w.fill_tree()
    f = w.open_designer(p.module("Form1"))
    cu = f.place_new("Project1.UserControl1", QRect(16, 200, 0, 0), None)
    assert cu is not None
    f.select([])
    assert on(w.designer_menu_items(f, None))["Update &UserControls"]
    f.update_user_controls()
    iu = f.items[[c for c in f.items if c.name == "UserControl11"][0]]
    f.select([iu])
    it = w.designer_menu_items(f, iu)
    assert texts(it) == ["Cu&t", "&Copy", "&Paste", "&Delete", None, "&Bring To Front", "Se&nd To Back", None,
                         "&View Code", "Align to &Grid"], texts(it)
    assert all(e[0] is None for e in it if e)
    # Properties > Description (me8-03, 04)
    assert w._props_menu()[0][0] == "&Description" and w._props_menu()[0][4] is True
    w.set_props_description(False)
    assert not w.props.desc.isVisible() and w.S["props_description"] is False and w._props_menu()[0][4] is False
    w.set_props_description(True)
    assert w.props.desc.isVisible()
    # Watches (me8-17, 18, 22)
    w.watch_win.fill([("x", "<Out of context>", "Empty", "Form1.Form_Load")])
    w.watch_win.select(0)
    it = w._watch_menu()
    assert [e[0] if e else None for e in it] == ["Cu&t", "&Copy", "&Paste", None, "&Edit Watch...", "&Add Watch...",
                                                 "&Delete Watch", None, "Co&llapse Parent", None, "&Start", "&End",
                                                 None, "&Definition"], it
    assert [bool(e[2]) for e in it if e] == [False, True, False, True, True, True, False, True, False, False], it
    w.watch_win.copy()
    assert cb.text() == "x\t<Out of context>\tEmpty\tForm1.Form_Load", cb.text()
    # Locals (me8-19, 20, 23): Collapse Parent on a line inside an other one
    lw = w.locals_win
    lw.fill("Project1.Form1.Form_Load", [("mvar", "a", "Integer", "1"), ("var", "x", "Integer", "5")], "Me")
    me = lw.tree.topLevelItem(0)
    me.setExpanded(True)
    lw.tree.setCurrentItem(me)
    it = w._locals_menu()
    assert [e[0] if e else None for e in it] == ["Cu&t", "&Copy", "&Paste", None, "Co&llapse Parent", None,
                                                 "&Start", "&End"], it
    assert not it[4][2]
    lw.tree.setCurrentItem(me.child(0))
    assert w._locals_menu()[4][2]
    lw.collapse_parent()
    assert lw.tree.currentItem() is me and not me.isExpanded()
    # the Object Browser (me8-21)
    w.object_browser()
    ob = w._objbrowser.widget()
    m = ob.menu_actions()
    assert [a.text() for a in m.actions() if not a.isSeparator()] == [
        "&Copy", "View &Definition", "&Find Whole Word Only", "Group &Members", "Show &Hidden Members",
        "&References...", "&Properties...", "&Help", "Doc&kable", "&Hide"], [a.text() for a in m.actions()]
    w._close_objbrowser()
    # Window: the window in front first, then as activated, at most 8; More Windows... from 8 windows on (me9)
    for _ in range(10):
        w.add_module("Module")
    app.processEvents()
    w.fill_window_menu()
    acts = [a for a in w.window_menu.actions() if not a.isSeparator()]
    texts_ = [a.text() for a in acts]
    assert texts_[-1] == "&More Windows..." and texts_[-2].startswith("&8 "), texts_
    first = [a for a in acts if a.text().startswith("&1 ")][0]
    assert first.isChecked() and first.text().endswith(w.mdi.activeSubWindow().windowTitle()), first.text()
    subs = w.window_list()
    assert subs[0] is w.mdi.activeSubWindow()

    def choose():
        dlg = w._more_windows_dlg
        lst = dlg.findChildren(__import__("PyQt6.QtWidgets", fromlist=["QListWidget"]).QListWidget)[0]
        items = [lst.item(i).text() for i in range(lst.count())]
        assert dlg.windowTitle() == "Windows" and items == sorted(sw.windowTitle() for sw in subs), items
        assert lst.currentRow() == 0
        dlg.accept()
    _QT.singleShot(0, choose)
    w.more_windows(subs)
    assert w.mdi.activeSubWindow().windowTitle() == min(sw.windowTitle() for sw in subs)
    done(w)


def t_ctlarray87(d):
    """control arrays by Copy / Paste and by (Name), as VB6 (phase 8.7, VB6 oracle DEV-8/me10)"""
    from PyQt6.QtCore import QTimer as _QT
    from PyQt6.QtWidgets import QMessageBox as _MB
    w = MainWindow()
    w.show()
    app.processEvents()
    f = w.open_designer(w.project.module("Form1"))

    def idx(name):
        return sorted(int(c.get("Index")) for c in f.all_controls() if c.name == name)
    c1 = f.place_new("CommandButton", QRect(160, 120, 0, 0), None)
    f.select([f.items[c1]])
    f.copy_selection()
    seen = {}

    def answer():
        box = f._array_box
        seen["text"] = box.text()
        seen["buttons"] = sorted(b.text().replace("&", "") for b in box.buttons())
        box.reject()                            # Esc / the close button: nothing (me10-08)
        app.processEvents()
        seen["after_esc"] = box.isVisible()
        box.button(_MB.StandardButton.Yes).click()
    f.select([])
    _QT.singleShot(0, answer)
    f.paste()                                   # the question, Yes: an array (me10-01, ca1 Form1.frm)
    assert seen["text"] == "You already have a control named 'Command1'.  Do you want to create a control array?", seen
    assert seen["buttons"] == ["Help", "No", "Yes"] and seen["after_esc"], seen
    assert idx("Command1") == [0, 1], idx("Command1")
    new = [c for c in f.all_controls() if c.name == "Command1" and int(c.get("Index")) == 1][0]
    assert (int(new.get("Left")), int(new.get("Top"))) == (0, 0)
    assert int(c1.get("Left")) > 0                # the first one stays where it was
    # a Paste into the array: no question, the next Index, at 0, 0 again (me10-03, 07)
    f.ask_control_array = lambda name: seen.setdefault("asked", True)
    f.select([])
    f.paste()
    assert "asked" not in seen and idx("Command1") == [0, 1, 2], idx("Command1")
    # No: a name of its own (me10-04)
    t1 = f.place_new("TextBox", QRect(300, 250, 0, 0), None)
    f.select([f.items[t1]])
    f.copy_selection()
    f.select([])
    f.ask_control_array = lambda name: False
    f.paste()
    assert sorted(c.name for c in f.all_controls() if c.type_name == "TextBox") == ["Text1", "Text2"]
    t2 = [c for c in f.all_controls() if c.name == "Text2"][0]
    # (Name) of an array of another type (me10-09)
    assert f.rename(t2, "Command1") == "Control must be same type as rest of array"
    ck = f.place_new("CheckBox", QRect(20, 250, 0, 0), None)      # of another type, not an array (me10-11)
    assert f.rename(t2, ck.name) == "'Check1' is a control name", f.rename(t2, ck.name)
    # (Name) of a control of the same type, No: the name stays; Yes: an array (me10-05, ca1 Form1.frm)
    assert f.rename(t2, "Text1") is None and t2.name == "Text2"
    f.ask_control_array = lambda name: True
    assert f.rename(t2, "Text1") is None and t2.name == "Text1" and idx("Text1") == [0, 1], idx("Text1")
    # (Name) of an array: no question, the next element (me10-10, ca2 F2.frm)
    t3 = f.place_new("TextBox", QRect(300, 50, 0, 0), None)
    f.ask_control_array = lambda name: seen.setdefault("asked2", True)
    assert f.rename(t3, "Text1") is None and "asked2" not in seen and idx("Text1") == [0, 1, 2], idx("Text1")
    del f.ask_control_array
    done(w)


def t_formname9(d):
    """a control may not take the name of a form, nor a form the name of a control (the user's decision,
    phase 9; VB6 oracle DEV-9/nm-01..04: VB6 accepts it and Form1 in the code is then the control)"""
    w = MainWindow()
    w.show()
    app.processEvents()
    w.add_module("Form")                        # Form2
    f = w.open_designer(w.project.module("Form1"))
    f.place_new("CommandButton", QRect(160, 120, 0, 0), None)
    c1 = [c for c in f.all_controls() if c.name == "Command1"][0]
    assert f.rename(c1, "Form1") == "Ambiguous name detected: Form1" and c1.name == "Command1"
    assert f.rename(c1, "form2") == "Ambiguous name detected: form2" and c1.name == "Command1"
    assert f.rename(c1, "Go") is None and c1.name == "Go"
    # a form into the name of a control (of this form or of an other one)
    m1, m2 = w.project.module("Form1"), w.project.module("Form2")
    assert w.rename_module(m2, "Go") == "Ambiguous name detected: Go" and m2.name == "Form2"
    assert w.rename_module(m1, "Go") == "Ambiguous name detected: Go" and m1.name == "Form1"
    # a new control never gets a form's name: a form called Command1
    assert w.rename_module(m2, "Command1") is None and m2.name == "Command1"
    f.place_new("CommandButton", QRect(20, 20, 0, 0), None)
    assert sorted(c.name for c in f.all_controls()) == ["Command2", "Go"], sorted(c.name for c in f.all_controls())
    # nor a pasted one
    c2 = [c for c in f.all_controls() if c.name == "Command2"][0]
    c2.name = "Command1"                        # as one that came from a VB6 form (the clipboard keeps its name)
    f.build()
    f.select([f.items[c2]])
    f.copy_selection()
    c2.name = "Command2"
    f.build()
    f.select([])
    f.paste()
    names = sorted(c.name for c in f.all_controls())
    assert "Command1" not in names and len(names) == 3, names
    done(w)


def t_pdw95(d):
    """Add-Ins > Package and Deployment Wizard (phase 9.5, VB6 oracle DEV-9/pd1): loaded in the Add-In Manager,
    the question of an unsaved project, the wizard's pages (Package Type, the question to recompile, Package
    Folder made when missing, Included Files with Add... and a checkbox, Package Format, Finished!), the package
    lbvc builds and the Packaging Report"""
    from lbstudio import pdwizard as PW
    from lbstudio.resedit import AddInManagerDialog, PDW
    w = MainWindow()
    w.show()
    app.processEvents()
    assert not w.a_pdw.isVisible() and w.a_pdw.text() == "Package and Deployment Wizard..."
    assert not w.a_pdw.icon().isNull()           # VB6 (DEV-9/E5-01): with its picture, after a separator
    acts = [a for a in w.menuBar().actions() if a.text() == "&Add-Ins"][0].menu().actions()
    i = acts.index(w.a_pdw)
    assert acts[i - 1] is w.sep_pdw and acts[i - 2].text() == "&Add-In Manager..." and acts[i - 3].isSeparator()
    # the Add-In Manager loads it (pd1-03: its Description is its name)
    dlg = AddInManagerDialog(w)
    row = [dlg.list.item(i, 0).text() for i in range(dlg.list.rowCount())].index(PDW)
    dlg.list.setCurrentCell(row, 0)
    assert dlg.desc.toPlainText() == PDW and dlg.cb_loaded.isEnabled()
    dlg.cb_loaded.setChecked(True)
    assert dlg.list.item(row, 1).text() == "Loaded"
    dlg._ok()
    w.apply_addins()
    assert w.a_pdw.isVisible() and w.pdw_loaded
    # an unsaved project: VB6's question first (pd1-01); Cancel: no wizard
    asked = []
    w.pdw_ask = lambda text, buttons: asked.append(text) or "cancel"
    assert w.package_wizard() is None and asked == [PW.SAVE_FIRST], asked
    code(w, w.project.module("Form1"),
         "Private Sub Form_Load()\n    Dim f As Integer\n    f = FreeFile\n"
         "    Open App.Path & \"/out.txt\" For Output As #f\n    Print #f, \"pdw\"\n    Close #f\n    Unload Me\nEnd Sub\n")
    saves(os.path.join(d, "Form1.frm"), os.path.join(d, "Project1.vbp"))
    assert w.save_project()
    vbp = os.path.join(d, "Project1.vbp")
    # the main window (pd1-02)
    wz = PW.Wizard(w)
    assert wz.windowTitle() == "Package and Deployment Wizard" and wz.active.text() == vbp
    assert wz.b_package.text() == "&Package" and wz.b_close.text() == "Close"
    # Package: Package Type
    pw = PW.PackageWizard(w, w.project)
    assert pw.windowTitle() == "Package and Deployment Wizard - Package Type"
    assert [pw.types.item(i).text() for i in range(pw.types.count())] == [
        "Portable Package for Linux", "Portable Package for Windows (32-bit)", "Portable Package for Windows (64-bit)"]
    assert pw.types.currentRow() == 0 and pw.type_desc.toPlainText().startswith("Use to create a folder")
    assert not pw.b_back.isEnabled() and pw.b_next.isEnabled() and not pw.b_finish.isEnabled()
    questions = []
    answers = []
    pw.ask = lambda text, buttons: questions.append(text) or answers.pop(0)
    pw.next()                                   # no program yet: it is compiled, no question
    exe = os.path.join(d, "Project1")
    assert questions == [] and os.path.exists(exe) and pw.stack.currentIndex() == 1
    assert pw.windowTitle() == "Package and Deployment Wizard - Package Folder"
    # a source newer than the program: VB6's question (pd1-05); No: the program as it is
    frm = os.path.join(d, "Form1.frm")
    later = time.time() + 5
    os.utime(frm, (later, later))
    t0 = os.path.getmtime(exe)
    pw.go(0)
    answers[:] = ["no"]
    pw.next()
    assert questions == [PW.recompile_text(exe, [frm])], questions
    assert questions[0].startswith("The following source files for this project are newer than the executable file")
    assert questions[0].endswith("Do you want to recompile? Choose No to use the existing executable.")
    assert pw.stack.currentIndex() == 1 and os.path.getmtime(exe) == t0
    # Package Folder: <project>\Package, made when missing (pd1-08, 09); No: the page stays
    folder = os.path.join(d, "Package")
    assert pw.folder.text() == folder
    answers[:] = ["no"]
    pw.next()
    assert questions[-1] == PW.folder_text(folder) and pw.stack.currentIndex() == 1 and not os.path.isdir(folder)
    answers[:] = ["yes"]
    pw.next()
    assert os.path.isdir(folder) and pw.stack.currentIndex() == 2
    assert pw.windowTitle() == "Package and Deployment Wizard - Included Files"
    # Included Files: the program first (always in), its runtime, the licenses (pd1-10, 11)
    names = [pw.list.topLevelItem(i).text(0) for i in range(pw.list.topLevelItemCount())]
    first = pw.list.topLevelItem(0)
    assert names[0] == "Project1" and not (first.flags() & Qt.ItemFlag.ItemIsUserCheckable), names
    assert first.checkState(0) == Qt.CheckState.Checked and pw.list.columnWidth(0) <= 240
    for n in ("lib/libQt6Core.so.6", "plugins/platforms/libqxcb.so", "qt.conf", "licenses/LGPL-3.0.txt",
              "licenses/GPL-3.0.txt"):
        assert n in names, (n, names)
    gpl = pw.list.topLevelItem(names.index("licenses/GPL-3.0.txt"))
    gpl.setCheckState(0, Qt.CheckState.Unchecked)
    readme = os.path.join(d, "readme.txt")
    open(readme, "w").write("a file of the user\n")
    QFileDialog.getOpenFileName = lambda *a, **k: (readme, "")
    pw.add_file()
    names = [pw.list.topLevelItem(i).text(0) for i in range(pw.list.topLevelItemCount())]
    assert names[-1] == "readme.txt", names
    assert pw.list.topLevelItem(names.index("licenses/GPL-3.0.txt")).checkState(0) == Qt.CheckState.Unchecked
    pw.next()
    # Package Format (in the place of Cab Options): Linux tar.gz / AppImage / folder
    assert pw.windowTitle() == "Package and Deployment Wizard - Package Format"
    shown = [k for k, b in pw.fmt_buttons.items() if not b.isHidden()]
    assert shown == ["tar.gz", "appimage", "folder"] and pw.fmt_buttons["tar.gz"].isChecked(), shown
    assert pw.format() == "tar.gz"
    pw.next()
    assert pw.windowTitle() == "Package and Deployment Wizard - Finished!"
    assert pw.script_name.text() == "Portable Package 1"
    assert not pw.b_next.isEnabled() and pw.b_finish.isEnabled()
    # Finish: the settings, lbvc --package, the report (pd1-20, 21)
    PW.ReportDialog.exec = lambda self: 0
    pw.finish()
    rep = w._pdw_report
    tgz = os.path.join(folder, "Project1-linux.tar.gz")
    assert rep.windowTitle() == "Project1.vbp - Packaging Report"
    assert rep.text.toPlainText().startswith(f"The portable package of your application has been built as '{tgz}'."), \
        rep.text.toPlainText()
    assert os.path.exists(tgz) and os.path.getmtime(exe) == t0
    app_dir = os.path.join(folder, "Project1-linux")
    assert os.path.exists(os.path.join(app_dir, "readme.txt"))
    assert not os.path.exists(os.path.join(app_dir, "licenses", "GPL-3.0.txt"))
    s = open(os.path.join(d, "Project1.lbpkg"), encoding="utf-8").read().splitlines()
    assert s == ["[Root]", "Most Recent Package=Portable Package 1", "", "[Package|Portable Package 1]", "Target=linux",
                 "Format=tar.gz", "PackageFolder=Package", "Added1=readme.txt", "Removed1=licenses/GPL-3.0.txt"], s
    # the settings come back the next time; Windows: zip / folder
    pw2 = PW.PackageWizard(w, w.project)
    assert pw2.added == ["readme.txt"] and pw2.removed == ["licenses/GPL-3.0.txt"] and pw2.folder.text() == folder
    # the next time through the pages: the file left out stays out, the file added stays in (DEV-9/E2, E3)
    pw2.ask = lambda text, buttons: "no"
    pw2.next()
    pw2.next()
    assert pw2.stack.currentIndex() == 2
    items = {pw2.list.topLevelItem(i).text(0): pw2.list.topLevelItem(i) for i in range(pw2.list.topLevelItemCount())}
    assert items["licenses/GPL-3.0.txt"].checkState(0) == Qt.CheckState.Unchecked
    assert items["readme.txt"].checkState(0) == Qt.CheckState.Checked
    pw2.next()
    assert pw2.removed == ["licenses/GPL-3.0.txt"] and pw2.added == ["readme.txt"], (pw2.removed, pw2.added)
    pw2.go(0)
    pw2.types.setCurrentRow(2)
    pw2.fill_formats()
    assert [k for k, b in pw2.fmt_buttons.items() if not b.isHidden()] == ["zip", "folder"]
    assert pw2.format() == "zip"
    # an IDE that cannot make programs for Windows (the AppImage: no mingw-w64, no Windows runtime): the Windows
    # systems shown but disabled, in the wizard and in Make (the user's decision, 2026-11-19)
    from lbstudio.dialogs import MakeDialog
    from lbvc import package as P
    real = MakeDialog.__dict__["available"]
    MakeDialog.available = classmethod(lambda cls, k: (k == "linux", "" if k == "linux" else cls.NOT_HERE))
    try:
        s = P.Settings.load(vbp)
        s.scripts[s.recent]["Target"] = "win64"           # the last package was for Windows
        s.save()
        pw3 = PW.PackageWizard(w, w.project)
        assert pw3.types.currentRow() == 0 and pw3.target() == "linux"
        for i in (1, 2):
            it = pw3.types.item(i)
            assert not (it.flags() & Qt.ItemFlag.ItemIsEnabled) and it.toolTip() == MakeDialog.NOT_HERE
        md = MakeDialog(w, d, "Project1", "win64")
        assert md.buttons["linux"].isChecked() and md.buttons["linux"].isEnabled()
        assert not md.buttons["win64"].isEnabled() and not md.buttons["win32"].isEnabled()
        assert md.buttons["win64"].toolTip() == MakeDialog.NOT_HERE
    finally:
        MakeDialog.available = real
    assert MakeDialog.available("win64")[0] in (True, False)
    done(w)


def t_cwbreak92(d):
    """a code window opened in break mode (the line of a run-time error in a module whose window was closed) has the
    code of the module, read-only, and no "reset your project" question (DEV-9 df3-02, df3-05)"""
    w = MainWindow()
    w.show()
    m = w.project.modules[0]
    m.set_code("Private Sub Form_Load()\r\n    x = 1\r\nEnd Sub\r\n")
    asked = []
    w.edit_while_running = lambda ed: asked.append(ed)
    w.runner.state = "break"
    try:
        cw = w.open_code(m)
        assert "x = 1" in cw.text(), repr(cw.text())
        assert cw.editor.SendScintilla(cw.editor.SCI_GETREADONLY) and not asked
    finally:
        w.runner.state = "design"
    done(w)


TESTS = {"t_axwiz92": t_axwiz92, "t_classbuilder92": t_classbuilder92, "t_cwbreak92": t_cwbreak92, "t_dfwiz92": t_dfwiz92, "t_appwiz92": t_appwiz92, "t_appwizgen92": t_appwizgen92, "t_web92": t_web92, "t_pdw95": t_pdw95, "t_formname9": t_formname9, "t_ctlarray87": t_ctlarray87, "t_ctxmenu87": t_ctxmenu87, "t_zoom87": t_zoom87, "t_tb87": t_tb87, "t_pe87": t_pe87, "t_menus87": t_menus87, "t_help87": t_help87, "t_dock87": t_dock87, "t_theme": t_theme, "t_resedit86": t_resedit86, "t_formdesign84": t_formdesign84, "t_projprops85": t_projprops85, "t_print83": t_print83, "t_procattr83": t_procattr83, "t_editor83": t_editor83, "t_debug82": t_debug82, "t_objbrowser": t_objbrowser, "t_mapi": t_mapi, "t_drep": t_drep, "t_rdc": t_rdc, "t_coolbar": t_coolbar, "t_dgcols": t_dgcols, "t_dfmt": t_dfmt, "t_dbgrid": t_dbgrid, "t_ondemand": t_ondemand, "t_resize": t_resize, "t_refactor": t_refactor, "t_types": t_types, "t_uctl": t_uctl, "t_group": t_group, "t_dsr": t_dsr,
         "t_dsr2": t_dsr2}


def main():
    names = sys.argv[1:] or list(TESTS)
    ok = 0
    faulthandler.enable(all_threads=False)      # a crash of Qt / Scintilla: the Python lines where it happened
    for n in names:
        print(f"---- {n}", flush=True)           # the name of a test that crashes is the last one written
        d = os.path.join(work, n)
        os.makedirs(d)
        try:
            TESTS[n](d)
            print(f"PASS  {n}")
            ok += 1
        except Exception:                       # noqa: BLE001 - the report of the test
            print(f"FAIL  {n}")
            traceback.print_exc()
    print(f"{ok}/{len(names)} passed")
    return 0 if ok == len(names) else 1


if __name__ == "__main__":
    sys.exit(main())
