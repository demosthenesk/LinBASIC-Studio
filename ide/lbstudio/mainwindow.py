"""The main window: the layout of the VB6 IDE (menus, Standard toolbar, Project Explorer,
Properties, Immediate) around the code windows, which use the editor of Linpad++."""
import os

from PyQt6.QtCore import Qt, QTimer, QSize
from PyQt6.QtGui import QAction, QKeySequence, QFont, QColor, QPalette, QIcon
from PyQt6.QtWidgets import (QMainWindow, QMdiArea, QMdiSubWindow, QDockWidget, QTreeWidget, QTreeWidgetItem,
                             QPlainTextEdit, QFileDialog, QMessageBox, QTabWidget,
                             QWidget, QVBoxLayout, QHBoxLayout, QToolBar, QToolButton, QLabel, QApplication, QStyle)
from PyQt6.Qsci import QsciScintilla as S

from . import APP_NAME, VERSION, paths, settings, styles as stl
from .codewin import CodeWindow
from . import refactorui as RF
from . import codeops, edicons, vbmodel
from .designer import FormDesigner, Toolbox
from .propwin import PropertiesWindow
from .dialogs import FindDialog, OptionsDialog, ProjectPropertiesDialog
from .langs import LangDB
from .runner import Runner
from .vbmodel import VBProject, VBGroup, FOLDER_OF_KIND

# markers of the debugger (the ones of Linpad++ are 20-24)
MARK_CUR, MARK_CUR_BACK, MARK_BP, MARK_BP_BACK, MARK_CALL = 19, 18, 17, 16, 15     # the yellow line over a breakpoint


def file_filter(label, exts):
    """a filter of QFileDialog that finds the names of Windows in any case: *.vbp *.VBP *.Vbp"""
    pats = []
    for e in exts:
        for v in (e.lower(), e.upper(), e[:1].upper() + e[1:].lower()):
            if f"*.{v}" not in pats:
                pats.append(f"*.{v}")
    return f"{label} ({' '.join(pats)})"


class CodeSubWindow(QMdiSubWindow):
    def __init__(self, main):
        super().__init__()
        self.main = main

    def closeEvent(self, e):
        self.main.code_closing(self)
        super().closeEvent(e)


class QuerySubWindow(QMdiSubWindow):
    """a window of the Query Designer: the SQL of a command goes back to it when it closes (it asks)"""

    def __init__(self, main):
        super().__init__()
        self.main = main

    def closeEvent(self, e):
        w = self.widget()
        if w is not None and hasattr(w, "can_close") and not w.can_close():
            e.ignore()
            return
        super().closeEvent(e)


class DesignSubWindow(QMdiSubWindow):
    def __init__(self, main):
        super().__init__()
        self.main = main

    def closeEvent(self, e):
        self.main.design_windows.pop(self.widget().module, None)
        super().closeEvent(e)
        fl = getattr(self.main, "formlayout", None)
        if fl is not None:                      # the Form Layout window shows the open forms only
            fl.update()


def _gsetting(schema, key):
    """the value of a gsettings key ('' when there is no such schema / key / gsettings)"""
    import subprocess
    try:
        r = subprocess.run(["gsettings", "get", schema, key], capture_output=True, text=True, timeout=2)
    except (OSError, subprocess.SubprocessError):
        return ""
    return r.stdout.strip().strip("'").lower() if r.returncode == 0 else ""


def system_dark():
    """is the theme of the system dark? Windows: AppsUseLightTheme. Linux: GTK_THEME, then the desktop's own
    setting (MATE: org.mate.interface gtk-theme, "Yaru-blue" / "Yaru-blue-dark"), GNOME's color-scheme and
    gtk-theme, KDE's kdeglobals; Qt's colorScheme where these say nothing. The settings are asked each time:
    some desktops (MATE with the gtk2 platform theme of Qt) do not tell Qt when the theme changes"""
    if os.name == "nt":
        try:
            import winreg
            k = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                               r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
        except OSError:
            return False

    def qt_scheme():
        try:
            from PyQt6.QtCore import Qt as _Qt
            cs = QApplication.instance().styleHints().colorScheme()
            return {_Qt.ColorScheme.Dark: True, _Qt.ColorScheme.Light: False}.get(cs)
        except (AttributeError, RuntimeError):
            return None

    desktop = os.environ.get("XDG_CURRENT_DESKTOP", "").lower()
    gtk = os.environ.get("GTK_THEME", "").lower()
    if gtk:
        return gtk.endswith(":dark") or "dark" in gtk
    if "kde" in desktop:
        q = qt_scheme()
        if q is not None:
            return q
        try:
            with open(os.path.expanduser("~/.config/kdeglobals"), encoding="utf-8", errors="replace") as f:
                for line in f:
                    if line.startswith("ColorScheme="):
                        return "dark" in line.lower()
        except OSError:
            pass
        return False
    if "mate" in desktop:
        t = _gsetting("org.mate.interface", "gtk-theme")
        if t:
            return "dark" in t
    if "cinnamon" in desktop:
        t = _gsetting("org.cinnamon.desktop.interface", "gtk-theme")
        if t:
            return "dark" in t
    cs = _gsetting("org.gnome.desktop.interface", "color-scheme")
    if cs in ("prefer-dark", "prefer-light"):
        return cs == "prefer-dark"
    t = _gsetting("org.gnome.desktop.interface", "gtk-theme")
    if t:
        return "dark" in t
    q = qt_scheme()
    return bool(q)


class ResSubWindow(QMdiSubWindow):
    """the windows of the Resource Editor: its X only hides "VB Resource Editor" (the changes stay, DEV-8/rs1
    stage 4); "Edit String Tables" closes"""

    def __init__(self, main, keep):
        super().__init__()
        self.main, self.keep = main, keep

    def closeEvent(self, e):
        if self.keep:
            e.ignore()
            self.hide()
            return
        ed = getattr(self.main, "res_editor", None)
        super().closeEvent(e)
        if ed is not None and ed.strings_win is self:
            ed.strings_closed()


# Edit > Insert File...: the four types of VB6's dialog (DEV-8/me2 picture 4); "All Files" as everywhere else here
INSERT_FILE_FILTER = "Text Files (*.txt);;Basic Files (*.bas);;Class Files (*.cls);;All Files (*)"



def norm_text(text):
    """the text of a menu command without its & and its keys"""
    return text.split("\t")[0].replace("&&", "\0").replace("&", "").replace("\0", "&").strip()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.S = settings.S
        self.langdb = LangDB()
        self.styles = self._load_styles()
        self.project = None
        self.code_windows = {}                  # module -> QMdiSubWindow
        self.design_windows = {}                # form module -> QMdiSubWindow
        self.find_dlg = None
        self.res_loaded = bool(self.S["addin_reseditor_startup"])   # Add-Ins: the VB 6 Resource Editor
        self.pdw_loaded = bool(self.S["addin_pdw_startup"])         # and the Package and Deployment Wizard
        self.appwiz_loaded = bool(self.S["addin_appwiz_startup"])   # and the VB 6 Application Wizard (phase 9.2)
        self.dfwiz_loaded = bool(self.S["addin_dfwiz_startup"])     # and the VB 6 Data Form Wizard (phase 9.2b)
        self.cbuilder_loaded = bool(self.S["addin_cbuilder_startup"])   # and the VB 6 Class Builder Utility (9.2c)
        self.axwiz_loaded = bool(self.S["addin_axwiz_startup"])   # and the VB 6 ActiveX Ctrl Interface Wizard (9.2d)
        self.ppwiz_loaded = bool(self.S["addin_ppwiz_startup"])   # and the VB 6 Property Page Wizard (9.2e)
        self.dowiz_loaded = bool(self.S["addin_dowiz_startup"])   # and the VB 6 Data Object Wizard (9.2f)
        self.apiviewer_loaded = bool(self.S["addin_apiviewer_startup"])   # and the VB 6 API Viewer (9.4)
        self.res_editor = None
        self.res_sub = None
        self.runner = Runner(self)
        self.runner.output.connect(self.immediate_write)
        self.runner.compile_error.connect(self.on_compile_error)
        self.runner.compile_warning.connect(self.on_compile_warning)
        self.runner.runtime_error.connect(self.on_runtime_error)
        self.runner.state_changed.connect(self.on_run_state)
        self.runner.built.connect(self.on_built)
        self.resize(1200, 800)
        self._build_ui()
        self._build_menus()
        from .cmdbars import CommandBars        # View > Toolbars > Customize... (phase 12.2)
        self.cmdbars = CommandBars(self)
        self.apply_addins()
        self.addin_hosts = {}                   # the add-ins of the user that are loaded (addinhost.py, phase 12.5)
        self.addin_errors = []
        QTimer.singleShot(0, self.load_startup_addins)
        self.apply_theme()
        hints = QApplication.instance().styleHints()
        if hasattr(hints, "colorSchemeChanged"):           # Qt 6.5
            hints.colorSchemeChanged.connect(self.system_theme_changed)
        self.theme_timer = QTimer(self)                     # Theme = system: the desktops that do not tell Qt
        self.theme_timer.setInterval(3000)
        self.theme_timer.timeout.connect(self.system_theme_changed)
        self.theme_timer.start()
        self.group = None                       # a project group (.vbg), or None: one project
        self.new_project("Exe")
        self.update_title()

    # ================================================================ UI
    def _build_ui(self):
        self.mdi = QMdiArea()
        self.mdi.setViewMode(QMdiArea.ViewMode.TabbedView if self.S["mdi_tabbed"] else QMdiArea.ViewMode.SubWindowView)
        self.mdi.setTabsClosable(True)
        self.mdi.setTabsMovable(True)
        self.mdi.setDocumentMode(True)
        self.mdi.subWindowActivated.connect(self._sub_activated)
        self.setCentralWidget(self.mdi)

        # Toolbox
        self.toolbox = Toolbox(self)
        self.dock_toolbox = self._dock("Toolbox", "General", self.toolbox, Qt.DockWidgetArea.LeftDockWidgetArea)

        # Project Explorer
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.itemDoubleClicked.connect(self._tree_open)
        self.tree.currentItemChanged.connect(self._tree_current)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._tree_menu)
        self.tree.itemExpanded.connect(self._tree_expanded)
        self.tree.itemCollapsed.connect(self._tree_expanded)
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        # VB6's buttons (DEV-8/me5 pictures 2-4): View Code, View Object | Toggle Folders (pressed: folders)
        tb = QToolBar()
        tb.setIconSize(QSize(16, 16))
        self.a_pe_code = tb.addAction(edicons.icon("view_code"), "View Code", self.view_code)
        self.a_pe_object = tb.addAction(edicons.icon("view_object"), "View Object", self.view_object)
        tb.addSeparator()
        self.a_pe_folders = tb.addAction(edicons.icon("toggle_folders"), "Toggle Folders", self.toggle_folders)
        self.a_pe_folders.setCheckable(True)
        self.a_pe_folders.setChecked(True)
        self.pe_folders = True                  # the modules in their folders (Forms, Modules ...)
        self.pe_expanded = {}                   # (project, folder) -> open; what the user opened stays open
        lay.addWidget(tb)
        lay.addWidget(self.tree)
        self.dock_project = self._dock("Project", "Project Explorer", w, Qt.DockWidgetArea.RightDockWidgetArea)

        # Properties
        self.props = PropertiesWindow(self)
        pw = self.props
        self.dock_props = self._dock("Properties", "Properties", pw, Qt.DockWidgetArea.RightDockWidgetArea)
        # Form Layout (VB6: under the Properties window)
        from .formlayout import FormLayoutView
        self.formlayout = FormLayoutView(self)
        self.dock_layout = self._dock("FormLayout", "Form Layout", self.formlayout, Qt.DockWidgetArea.RightDockWidgetArea)
        self.layout_sub = None                  # the Form Layout as an MDI window when it is not Dockable
        self.formlayout.hide_requested.connect(self.hide_layout)
        # after the menu of the right button has closed: the window changes its parent
        self.formlayout.dockable_toggled.connect(lambda on: QTimer.singleShot(0, lambda: self.set_layout_dockable(on)))
        self.dock_layout.topLevelChanged.connect(self._layout_top_level)
        self._place_layout()
        # View > Color Palette: a small window of its own, over the IDE (DEV-8/me2 pictures 7, 8)
        from .colorpal import ColorPalette
        self.palette_win = ColorPalette(self)
        self.dock_palette = self._dock("ColorPalette", "Color Palette", self.palette_win,
                                       Qt.DockWidgetArea.TopDockWidgetArea)
        self.dock_palette.setFloating(True)
        self.dock_palette.hide()
        # the Data View (phase 12.3; View > Data View Window): at the bottom, hidden at first
        from .dataview import DataView
        self.dataview = DataView(self)
        self.dock_dataview = self._dock("DataView", "Data View", self.dataview, Qt.DockWidgetArea.BottomDockWidgetArea)
        self.dock_dataview.hide()

        # the bottom: Output, Immediate, Locals, Watches in tabs
        from .debugwin import OutputWindow, ImmediateEdit, LocalsWindow, WatchWindow
        f = QFont(stl.font_family("Courier New"))
        f.setPointSize(10)
        self.output = OutputWindow()
        self.output.setFont(f)
        self.output.goto.connect(self.goto_message)
        self.immediate = ImmediateEdit()
        self.immediate.setFont(f)
        self.immediate.command.connect(self.immediate_command)
        self.locals_win = LocalsWindow()
        self.locals_win.stack_requested.connect(self.call_stack)
        self.watch_win = WatchWindow()
        self.watch_win.edit_requested.connect(self.edit_watch)
        self.watch_win.delete_requested.connect(self.delete_watch)
        self.bottom_tabs = QTabWidget()
        self.bottom_tabs.setDocumentMode(True)
        self.bottom_tabs.setTabPosition(QTabWidget.TabPosition.South)
        from .refactorui import ReferencesWindow
        self.references = ReferencesWindow()
        self.references.goto.connect(lambda m, line, col, n: RF.select_range(self, m, line, col, n))
        for w, name in ((self.output, "Output"), (self.immediate, "Immediate"), (self.locals_win, "Locals"),
                        (self.watch_win, "Watches"), (self.references, "References")):
            self.bottom_tabs.addTab(w, name)
        self.dock_bottom = self._dock("Bottom", "Output", self.bottom_tabs, Qt.DockWidgetArea.BottomDockWidgetArea)
        self.bottom_tabs.currentChanged.connect(lambda i: self.dock_bottom.setWindowTitle(self.bottom_tabs.tabText(i)))
        # the menus and the toolbar name the windows: one dock, a tab each
        self.dock_output = self.dock_immediate = self.dock_locals = self.dock_watch = self.dock_bottom
        self.resizeDocks([self.dock_bottom], [190], Qt.Orientation.Vertical)
        # Dockable (Tools > Options > Docking, the right button of each tool window: DEV-8/me4)
        self.free_docks = {}                    # Immediate / Locals / Watches out of the bottom dock
        for key, w, extras in (("props", self.props.grid, self._props_menu), ("props", self.props.desc, self._props_menu),
                               ("immediate", self.immediate, self._immediate_menu),
                               ("locals", self.locals_win, self._locals_menu), ("watch", self.watch_win, self._watch_menu),
                               ("toolbox", self.toolbox, self._toolbox_menu), ("palette", self.palette_win, None)):
            self.tool_menu(key, w, extras)
        self.props.desc.setVisible(bool(self.S.get("props_description", True)))
        for key, _title in self.TOOL_WINDOWS:
            if key not in ("layout", "objbrowser") and not self.is_dockable(key):
                self._apply_dockable(key)
        self.watches = []                       # the Watches window: {"id", "expr", "kind", "mod", "proc"}
        self._watch_id = 0
        self.frame_depth = 0                    # the frame of the Call Stack that Locals shows
        self.break_at = None                    # (module, editor line) of the yellow line
        self.last_break = None                  # (module, file line, reason, detail) of the last break
        self.trap_session = None                # Toggle > Break on ...: Error Trapping for this session only
        self.bp_disabled = False                # Debug > Disable All Breakpoints
        self.runner.broke.connect(self.on_break)
        self.runner.bp_provider = lambda: [] if self.bp_disabled else self.breakpoint_list()
        self.runner.trap_provider = self.trap_mode
        self.runner.watch_provider = lambda: [self._watch_cmd(w) for w in self.watches if w["kind"]]

        self.pos_label = QLabel()
        self.statusBar().addPermanentWidget(self.pos_label)

    def _dock(self, obj, title, widget, area):
        d = QDockWidget(title, self)
        d.setObjectName(obj)
        d.setWidget(widget)
        self.addDockWidget(area, d)
        return d

    def _act(self, menu, text, slot, key=None, tip=None, icon=None):
        a = QAction(text, self)
        if key:
            a.setShortcut(QKeySequence(key))
        if isinstance(icon, QIcon):              # a drawn picture (edicons)
            a.setIcon(icon)
        elif icon is not None:                  # a picture of the Qt style
            a.setIcon(self.style().standardIcon(icon))
        a.triggered.connect(lambda checked=False: slot())
        menu.addAction(a)
        return a

    def _build_menus(self):
        mb = self.menuBar()
        SP = QStyle.StandardPixmap
        m = mb.addMenu("&File")
        self.a_new = self._act(m, "&New Project", lambda: self.new_project(), "Ctrl+N", icon=SP.SP_FileIcon)
        self.a_open = self._act(m, "&Open Project...", self.open_project_dialog, "Ctrl+O", icon=SP.SP_DialogOpenButton)
        self.a_add_proj = self._act(m, "A&dd Project...", self.add_project)
        self.a_remove_proj = self._act(m, "&Remove Project", lambda: self.remove_project(self.project))
        self.a_close_proj = self._act(m, "&Close Project", self.close_project)
        self.a_close_group = self._act(m, "Close Project &Group", self.close_group)
        m.addSeparator()
        self.a_save_proj = self._act(m, "Sa&ve Project", self.save_project, icon=SP.SP_DialogSaveButton)
        self.a_save_proj_as = self._act(m, "Sav&e Project As...", self.save_project_as)
        m.addSeparator()
        self.a_save_file = self._act(m, "&Save <File>", self.save_file, "Ctrl+S")
        self.a_save_file_as = self._act(m, "Save <File> &As...", self.save_file_as)
        self.a_save_group = self._act(m, "Save Project Gro&up", self.save_group)
        self.a_save_group_as = self._act(m, "Save Project Group As...", lambda: self.save_group(ask=True))
        m.addSeparator()
        self.a_print = self._act(m, "&Print...", self.print_dialog, "Ctrl+P")
        self.a_print_setup = self._act(m, "Print Set&up...", self.print_setup)
        m.addSeparator()
        self.a_make = self._act(m, "Ma&ke <Project>...", self.make_exe)
        self.a_make_group = self._act(m, "Make Project &Group...", self.make_group)
        m.addSeparator()
        self.recent_menu = m.addMenu("&Recent Projects")
        self.fill_recent()
        m.addSeparator()
        self._act(m, "E&xit", self.close, "Alt+Q")
        m.aboutToShow.connect(self._update_file_actions)       # Save Form1, Make Project1.exe ... (VB6)

        m = mb.addMenu("&Edit")                # VB6's Edit menu (DEV-8/me1, pictures 2-5, 27-29)
        self.edit_menu = m
        self.a_undo = self._act(m, "Can't Undo", lambda: self._ed_call("undo"), "Ctrl+Z")
        self.a_redo = self._act(m, "Can't Redo", lambda: self._ed_call("redo"))
        m.addSeparator()
        self.a_cut = self._act(m, "Cu&t", lambda: self._ed_call("cut"), "Ctrl+X")
        self.a_copy = self._act(m, "&Copy", lambda: self._ed_call("copy"), "Ctrl+C")
        self.a_paste = self._act(m, "&Paste", lambda: self._ed_call("paste"), "Ctrl+V")
        self.a_paste_link = self._act(m, "Paste Lin&k", lambda: None)          # OLE links: always grey here
        self.a_delete = self._act(m, "&Delete", lambda: self._ed_call("removeSelectedText"), "Del")
        self.a_select_all = self._act(m, "Select &All", lambda: self._ed_call("selectAll"), "Ctrl+A")
        m.addSeparator()
        self.a_find = self._act(m, "&Find...", self.find, "Ctrl+F")
        self.a_find_next = self._act(m, "Find &Next", self.find_next, "F3")
        self.a_replace = self._act(m, "R&eplace...", self.replace, "Ctrl+H")
        m.addSeparator()
        self.a_indent = self._act(m, "&Indent", lambda: self._indent(True), "Tab", icon=edicons.icon("indent"))
        self.a_indent.setShortcutContext(Qt.ShortcutContext.WidgetShortcut)
        self.a_outdent = self._act(m, "&Outdent", lambda: self._indent(False), "Shift+Tab", icon=edicons.icon("outdent"))
        self.a_outdent.setShortcutContext(Qt.ShortcutContext.WidgetShortcut)
        m.addSeparator()
        self.a_insert_file = self._act(m, "Insert Fi&le...", self.insert_file)
        m.addSeparator()
        self.a_list_members = self._act(m, "List Properties/Met&hods",
                                        lambda: self._assist("list_members") or self._designer_do("z_order", True),
                                        "Ctrl+J", icon=edicons.icon("list_members"))
        self.a_list_consts = self._act(m, "Li&st Constants", lambda: self._assist("list_constants"), "Ctrl+Shift+J",
                                       icon=edicons.icon("list_constants"))
        self.a_quick_info = self._act(m, "&Quick Info", lambda: self._assist("quick_info"), "Ctrl+I",
                                      icon=edicons.icon("quick_info"))
        self.a_param_info = self._act(m, "Para&meter Info", lambda: self._assist("parameter_info"), "Ctrl+Shift+I",
                                      icon=edicons.icon("param_info"))
        self.a_complete = self._act(m, "Complete &Word", lambda: self._assist("complete_word"), "Ctrl+Space",
                                    icon=edicons.icon("complete_word"))
        m.addSeparator()
        # refactoring of LinBASIC Studio (Phase 5; not in VB6): the names of the code as the compiler binds them
        self._act(m, "Find Re&ferences", lambda: RF.find_references(self), "Shift+F12")
        self._act(m, "Rena&me Symbol...", lambda: RF.rename(self), "Ctrl+Shift+R")
        self._act(m, "E&xtract Procedure...", lambda: RF.extract(self), "Ctrl+Shift+M")
        m.addSeparator()
        bm = m.addMenu("&Bookmarks")              # VB6: Edit > Bookmarks
        self.a_bm_toggle = self._act(bm, "&Toggle Bookmark", self.toggle_code_bookmark, icon=edicons.icon("bm_toggle"))
        self.a_bm_next = self._act(bm, "&Next Bookmark", lambda: self.next_bookmark(True), icon=edicons.icon("bm_next"))
        self.a_bm_prev = self._act(bm, "&Previous Bookmark", lambda: self.next_bookmark(False),
                                   icon=edicons.icon("bm_prev"))
        self.a_bm_clear = self._act(bm, "&Clear All Bookmarks", self.clear_bookmarks, icon=edicons.icon("bm_clear"))
        m.aboutToShow.connect(self.update_edit_menu)
        # Comment Block / Uncomment Block: on the Edit toolbar only, as in VB6
        self.a_comment = QAction(edicons.icon("comment"), "Comment Block", self)
        self.a_comment.triggered.connect(lambda: self.comment_block(False))
        self.a_uncomment = QAction(edicons.icon("uncomment"), "Uncomment Block", self)
        self.a_uncomment.triggered.connect(lambda: self.comment_block(True))

        m = mb.addMenu("&View")                # VB6's View menu (DEV-8/me1, pictures 6-9, 30-33)
        self.view_menu = m
        self._act(m, "&Code", self.view_code, "F7")
        self._act(m, "O&bject", self.view_object, "Shift+F7")
        m.addSeparator()
        self.a_definition = self._act(m, "&Definition", lambda: RF.definition(self), "Shift+F2")
        self.a_last_pos = self._act(m, "Last Positio&n", lambda: RF.last_position(self), "Ctrl+Shift+F2")
        m.addSeparator()
        self.a_objbrowser = self._act(m, "&Object Browser", self.object_browser, "F2",
                                      icon=QStyle.StandardPixmap.SP_FileDialogContentsView)
        m.addSeparator()
        self.a_view_output = self._act(m, "O&utput Window", lambda: self._show_dock(self.dock_output, self.output),
                                       "Ctrl+Alt+O")
        self.a_view_immediate = self._act(m, "&Immediate Window",
                                          lambda: self._show_dock(self.dock_immediate, self.immediate), "Ctrl+G")
        self.a_view_locals = self._act(m, "Local&s Window", lambda: self._show_dock(self.dock_locals, self.locals_win.tree))
        self.a_view_watch = self._act(m, "Watc&h Window", lambda: self._show_dock(self.dock_watch, self.watch_win.tree))
        self.a_call_stack = self._act(m, "Call Stac&k...", self.call_stack, "Ctrl+L", icon=edicons.icon("call_stack"))
        m.addSeparator()
        self.a_view_project = self._act(m, "&Project Explorer", lambda: self._show_dock(self.dock_project, self.tree),
                                        "Ctrl+R")
        self.a_view_props = self._act(m, "Properties &Window", lambda: self._show_dock(self.dock_props, self.props.grid),
                                      "F4")
        self.a_view_layout = self._act(m, "&Form Layout Window", self.show_layout)
        self.a_view_ppages = self._act(m, "Property Pa&ges", self.property_pages, "Shift+F4")
        self.sep_view_tools = m.addSeparator()   # (Table, Zoom, Show Panes of the Query Designer before it)
        self.a_view_toolbox = self._act(m, "Toolbo&x", lambda: self._show_dock(self.dock_toolbox, self.toolbox))
        self.a_view_dataview = self._act(m, "Data Vie&w Window", self.show_data_view,
                                         icon=edicons.icon("data_view"))
        self.a_view_palette = self._act(m, "Color Pa&lette", self.show_palette)
        m.addSeparator()
        # View > Zoom: the zoom of the code windows without Ctrl+Mouse Wheel (a laptop touchpad), not in VB6
        zm = m.addMenu("&Zoom")
        self.a_zoom_in = self._act(zm, "Zoom &In", lambda: self.code_zoom_step(1))
        self.a_zoom_in.setShortcuts([QKeySequence("Ctrl+="), QKeySequence("Ctrl++")])
        self.a_zoom_out = self._act(zm, "Zoom &Out", lambda: self.code_zoom_step(-1), "Ctrl+-")
        self.a_zoom_reset = self._act(zm, "Zoom &Reset", lambda: self.code_zoom_step(0), "Ctrl+0")

        m = mb.addMenu("&Project")             # VB6's Project menu (DEV-8/me1, picture 10)
        self.project_menu = m
        self._act(m, "Add &Form", lambda: self.add_module("Form"))
        self.a_add_mdi = self._act(m, "Add MD&I Form", self.add_mdi_form)
        self._act(m, "Add &Module", lambda: self.add_module("Module"))
        self._act(m, "Add &Class Module", self.add_class_module)
        self.a_add_uctl = self._act(m, "Add &User Control", lambda: self.add_module("UserControl"))
        self._act(m, "Add &Property Page", lambda: self.add_module("PropertyPage"))
        # not made by LinBASIC Studio yet (User Document: Phase 10.2, WebClass: 10.3, DHTML Page: never)
        self.a_add_udoc = self._act(m, "Add User &Document", lambda: None)
        self.a_add_webclass = self._act(m, "Add WebClass", lambda: None)
        self._act(m, "Add Data Report", lambda: self.add_module("DataReport"))
        self.a_add_dhtml = self._act(m, "Add DHTML Page", lambda: None)
        self._act(m, "Add Data Environment", lambda: self.add_module("DataEnvironment"))
        self.m_more_designers = m.addMenu("More ActiveX Desi&gners...")
        self._act(m, "&Add File...", self.add_file, "Ctrl+D")
        m.addSeparator()
        self.a_remove = self._act(m, "&Remove <File>", self.remove_module)
        m.addSeparator()
        self._act(m, "Refere&nces...", self.references_dialog)
        self._act(m, "C&omponents...", self.components_dialog, "Ctrl+T")
        m.addSeparator()
        self.a_proj_props = self._act(m, "<Project> Prop&erties...", self.project_properties)
        from .resedit import _icon as res_icon
        self.a_new_res = self._act(m, "Add New Resource File", self.add_new_resource_file, icon=res_icon("bitmap"))
        m.aboutToShow.connect(self.update_project_menu)

        m = mb.addMenu("F&ormat")              # VB6's Format menu (DEV-8/fd1 stage 1)
        self.fmt_actions = {"one": [], "two": [], "three": []}
        al = m.addMenu("&Align")
        for t, h in (("&Lefts", "lefts"), ("&Centers", "centers"), ("&Rights", "rights"), ("&Tops", "tops"),
                     ("&Middles", "middles"), ("&Bottoms", "bottoms")):
            self.fmt_actions["two"].append(self._act(al, t, lambda hh=h: self._designer_do("align", hh)))
        al.addSeparator()
        self.fmt_actions["one"].append(self._act(al, "to &Grid", lambda: self._designer_do("align", "grid")))
        ms = m.addMenu("&Make Same Size")
        for t, h in (("&Width", "width"), ("&Height", "height"), ("&Both", "both")):
            self.fmt_actions["two"].append(self._act(ms, t, lambda hh=h: self._designer_do("same_size", hh)))
        m.addSeparator()
        self.fmt_actions["one"].append(self._act(m, "Size to Gr&id", lambda: self._designer_do("size_to_grid", None)))
        m.addSeparator()
        for title, vertical in (("&Horizontal Spacing", False), ("&Vertical Spacing", True)):
            sm = m.addMenu(title)
            for t, h, need in (("Make &Equal", "equal", "three"), ("&Increase", "increase", "two"),
                               ("&Decrease", "decrease", "two"), ("&Remove", "remove", "two")):
                self.fmt_actions[need].append(
                    self._act(sm, t, lambda vv=vertical, hh=h: self._designer_do("spacing", (vv, hh))))
        m.addSeparator()
        cf = m.addMenu("&Center in Form")
        self.fmt_actions["one"].append(self._act(cf, "&Horizontally", lambda: self._designer_do("center_in_form", True)))
        self.fmt_actions["one"].append(self._act(cf, "&Vertically", lambda: self._designer_do("center_in_form", False)))
        m.addSeparator()
        od = m.addMenu("&Order")
        self.a_front = self._act(od, "&Bring to Front", lambda: self._designer_do("z_order", True), "Ctrl+J",
                                 icon=edicons.icon("front"))
        self.a_back = self._act(od, "&Send to Back", lambda: self._designer_do("z_order", False), "Ctrl+K",
                                icon=edicons.icon("back"))
        self.fmt_actions["one"] += [self.a_front, self.a_back]
        self.m_align, self.m_same_size, self.m_center = al, ms, cf
        m.addSeparator()
        self.a_lock = self._act(m, "&Lock Controls", lambda: self._designer_do("toggle_lock", None),
                                icon=edicons.icon("lock"))
        self.a_lock.setCheckable(True)
        self.format_menu = m
        m.aboutToShow.connect(self.update_format_menu)

        m = mb.addMenu("&Debug")
        self.a_step_into = self._act(m, "Step &Into", lambda: self.debug_step("step"), "F8",
                                     icon=edicons.icon("step_into"))
        self.a_step_over = self._act(m, "Step &Over", lambda: self.debug_step("over"), "Shift+F8",
                                     icon=edicons.icon("step_over"))
        self.a_step_out = self._act(m, "Step O&ut", lambda: self.debug_step("out"), "Ctrl+Shift+F8",
                                    icon=edicons.icon("step_out"))
        self._act(m, "&Run To Cursor", self.run_to_cursor, "Ctrl+F8")
        m.addSeparator()
        self._act(m, "&Add Watch...", self.add_watch)
        self._act(m, "&Edit Watch...", lambda: self.edit_watch(self.watch_win.tree.indexOfTopLevelItem(self.watch_win.tree.currentItem())), "Ctrl+W")
        self.a_quick_watch = self._act(m, "&Quick Watch...", self.quick_watch, "Shift+F9",
                                       icon=edicons.icon("quick_watch"))
        m.addSeparator()
        self.a_toggle_bp = self._act(m, "&Toggle Breakpoint", self.toggle_breakpoint, "F9",
                                     icon=edicons.icon("breakpoint"))
        self._act(m, "&Clear All Breakpoints", self.clear_breakpoints, "Ctrl+Shift+F9")
        self.a_bp_disable = self._act(m, "&Disable All Breakpoints", self.disable_breakpoints)
        self.a_bp_disable.setCheckable(True)
        m.addSeparator()
        self._act(m, "Set &Next Statement", self.set_next_statement, "Ctrl+F9")
        self._act(m, "Show Ne&xt Statement", self.show_next_statement)
        # VB6: after a break at an error, Alt+F5 / Alt+F8 go on with the error (to the error handler, or the
        # error box); F5 / F8 run the statement again. Not in the menus, as in VB6
        for key, cmd in (("Alt+F5", "errgo"), ("Alt+F8", "errstep")):
            a = QAction(self)
            a.setShortcut(QKeySequence(key))
            a.setShortcutContext(Qt.ShortcutContext.ApplicationShortcut)
            a.triggered.connect(lambda checked=False, c=cmd: self.error_resume(c))
            self.addAction(a)

        m = mb.addMenu("&Run")
        self.a_start = self._act(m, "&Start", self.run_start, "F5", icon=SP.SP_MediaPlay)
        self.a_full = self._act(m, "Start With &Full Compile", self.run_start, "Ctrl+F5")
        # Qt calls the Break key Pause: the menu shows VB6's "Ctrl+Break", the key is Ctrl+Pause
        self.a_break = self._act(m, "Brea&k\tCtrl+Break", self.runner.pause, "Ctrl+Pause", icon=SP.SP_MediaPause)
        self.a_break.setIconText("Break")
        self.a_break.setToolTip("Break")
        self.a_break.setEnabled(False)
        self.a_end = self._act(m, "&End", self.run_end, icon=SP.SP_MediaStop)
        self.a_end.setEnabled(False)
        self.a_restart = self._act(m, "&Restart", self.run_restart, "Shift+F5")

        # the Query and Diagram menus of the Query Designer (phase 12.3): grey without one in front (VB6)
        qm = mb.addMenu("&Query")
        dm = mb.addMenu("D&iagram")
        from .querydesign import IdeMenus
        self.query_menus = IdeMenus(self, qm, dm, self.view_menu, self.sep_view_tools)

        m = mb.addMenu("&Tools")
        self.a_add_proc = self._act(m, "Add &Procedure...", self.add_procedure)
        self.a_proc_attr = self._act(m, "Procedure &Attributes...", self.procedure_attributes)
        m.addSeparator()
        self.a_menu_editor = self._act(m, "&Menu Editor...", lambda: self._designer_do("menu_editor", None), "Ctrl+E")
        m.addSeparator()
        self.a_options = self._act(m, "&Options...", self.options)
        self.sep_res = m.addSeparator()
        self.a_res_editor = self._act(m, "Resource Editor", self.show_res_editor, icon=res_icon("bitmap"))

        m = mb.addMenu("&Add-Ins")             # VB6 (DEV-8/me1, picture 23)
        self.m_addins = m                       # (the add-ins of the user add their commands to it, phase 12.5)
        # (phase 9.3: not made, decision 2026-11-26 - what to use instead)
        self.a_visdata = self._act(m, "&Visual Data Manager...", self.visual_data_manager)
        m.addSeparator()
        self._act(m, "&Add-In Manager...", self.addin_manager)
        # a loaded add-in: its command at the end of the menu, after a separator, with its picture (VB6:
        # DEV-9/E5-01)
        self.sep_pdw = m.addSeparator()
        from .pdwizard import package_icon
        self.a_pdw = self._act(m, "Package and Deployment Wizard...", self.package_wizard, icon=package_icon(16))
        # the VB 6 Application Wizard: Application Wizard... and Toolbar Wizard... (DEV-9/aw18-01)
        self.sep_appwiz = m.addSeparator()
        from .appwizard import wizard_icon, toolbar_wizard_icon
        self.a_appwiz = self._act(m, "Application Wizard...", self.application_wizard, icon=wizard_icon(16))
        self.a_tbwiz = self._act(m, "Toolbar Wizard...", self.toolbar_wizard, icon=toolbar_wizard_icon(16))
        # the VB 6 Data Form Wizard: Data Form Wizard... (VB6 df1-01, phase 9.2b)
        self.sep_dfwiz = m.addSeparator()
        from .dfwizard import dfwizard_icon
        self.a_dfwiz = self._act(m, "Data Form Wizard...", self.data_form_wizard, icon=dfwizard_icon(16))
        # the VB 6 Class Builder Utility: Class Builder Utility... (VB6 DEV-9/cb1-03, phase 9.2c)
        self.sep_cbuilder = m.addSeparator()
        from .classbuilder import builder_icon
        self.a_cbuilder = self._act(m, "Class Builder Utility...", self.class_builder, icon=builder_icon(16))
        # the VB 6 ActiveX Ctrl Interface Wizard: ActiveX Control Interface Wizard... (VB6 DEV-9/ax1, phase 9.2d)
        self.sep_axwiz = m.addSeparator()
        from .axwizard import axwizard_icon
        self.a_axwiz = self._act(m, "ActiveX Control Interface Wizard...", self.activex_wizard,
                                 icon=axwizard_icon(16))
        # the VB 6 Property Page Wizard: Property Page Wizard... (VB6 DEV-9/pp1, phase 9.2e)
        self.sep_ppwiz = m.addSeparator()
        from .ppwizard import ppwizard_icon
        self.a_ppwiz = self._act(m, "Property Page Wizard...", self.property_page_wizard, icon=ppwizard_icon(16))
        # the VB 6 Data Object Wizard: its name, without "..." (VB6 DEV-9/do1, phase 9.2f)
        self.sep_dowiz = m.addSeparator()
        from .dowizard import dowizard_icon
        self.a_dowiz = self._act(m, "VB 6 Data Object Wizard", self.data_object_wizard, icon=dowizard_icon(16))
        # the VB 6 API Viewer: API Viewer (VB6 DEV-9 api-w, phase 9.4)
        self.sep_apiview = m.addSeparator()
        from .apiviewer import apiviewer_icon
        self.a_apiview = self._act(m, "API Viewer", self.api_viewer, icon=apiviewer_icon(16))

        m = mb.addMenu("&Window")              # VB6 (DEV-8/me1, pictures 24, 37)
        self.window_menu = m
        m.aboutToShow.connect(self.fill_window_menu)
        self.fill_window_menu()

        m = mb.addMenu("&Help")                # VB6 without its Microsoft links (user, DEV-8/me1)
        self._act(m, "&Contents...", self.help_contents)
        self._act(m, "&Index...", lambda: self.help_contents(index=True))
        self._act(m, "&Search...", lambda: self.help_contents(search=True))
        m.addSeparator()
        self._act(m, "&Help on the Word or Object", self.context_help, "F1")
        self.a_help_folder = self._act(m, "Set Help &Folder...", self.set_help_folder)   # where the user's MSDN is
        m.addSeparator()
        self._act(m, "&About " + APP_NAME + "...", self.about)

        # the commands that need an open project (File > Close Project leaves none)
        self.project_actions = [self.a_close_proj, self.a_close_group, self.a_save_proj, self.a_save_proj_as, self.a_save_file,
                                self.a_save_file_as, self.a_make, self.a_add_proj, self.a_remove_proj,
                                self.a_save_group, self.a_save_group_as]
        for title in ("&View", "F&ormat", "&Project", "&Debug", "&Run", "&Tools"):
            menu = next(a.menu() for a in mb.actions() if a.text() == title)
            stack = [menu]
            while stack:
                mm = stack.pop()
                for a in mm.actions():
                    if a.menu() is not None:
                        stack.append(a.menu())
                    elif not a.isSeparator() and a not in self.project_actions:
                        self.project_actions.append(a)
        keep = {self.a_view_output, self.a_view_immediate, self.a_view_locals, self.a_view_watch, self.a_view_project,
                self.a_view_props, self.a_view_toolbox, self.a_objbrowser, self.a_options}
        self.project_actions = [a for a in self.project_actions if a not in keep]
        # always grey: not made by LinBASIC Studio yet, or not on these systems
        self.never_actions = [self.a_paste_link, self.a_add_udoc, self.a_add_webclass, self.a_add_dhtml,
                              self.m_more_designers.menuAction()]
        self.project_actions = [a for a in self.project_actions if a not in self.never_actions]
        for a in self.never_actions:
            a.setEnabled(False)

        # Standard toolbar, as VB6's (phase 12.2, oracle d4): Add Standard EXE Project and Add Form (each a button
        # and its list; the button does the command chosen last), Menu Editor | Open Project, Save Project | Cut,
        # Copy, Paste, Find | Undo, Redo | Start, Break, End | Project Explorer, Properties Window, Form Layout
        # Window, Object Browser, Toolbox, Data View Window, Visual Component Manager | the position, the size
        tb = self.addToolBar("Standard")
        tb.setObjectName("Standard")
        self.tb_add_project = self._list_button(tb, "add_project", [
            (f"Add {text} Project", lambda k=key: self.add_project(k)) for key, text in vbmodel.PROJECT_TYPES])
        self.tb_add_form = self._list_button(tb, "form", [
            (norm_text(a.text()), a) for a in self.project_menu.actions()
            if not a.isSeparator() and a.menu() is None and norm_text(a.text()).startswith("Add ")
            and norm_text(a.text()) != "Add File..."])
        tb.addAction(self.a_menu_editor)
        tb.addSeparator()
        for a in (self.a_open, self.a_save_proj):
            tb.addAction(a)
        tb.addSeparator()
        for a in (self.a_cut, self.a_copy, self.a_paste, self.a_find):
            tb.addAction(a)
        tb.addSeparator()
        for a in (self.a_undo, self.a_redo):
            tb.addAction(a)
        tb.addSeparator()
        for a in (self.a_start, self.a_break, self.a_end):
            tb.addAction(a)
        tb.addSeparator()
        for a in (self.a_view_project, self.a_view_props, self.a_view_layout, self.a_objbrowser,
                  self.a_view_toolbox):
            tb.addAction(a)
        self.a_data_view = self.a_view_dataview                                              # phase 12.3
        self.a_vcm = QAction(edicons.icon("component_manager"), "Visual Component Manager", self)   # Enterprise:
        self.a_vcm.setEnabled(False)                                                         # not made
        tb.addAction(self.a_data_view)
        tb.addAction(self.a_vcm)
        tb.addAction(self.a_res_editor)         # the button of the VB 6 Resource Editor (when loaded)
        # VB6: where the selected control is and its size ("1680, 1320"  "1215 x 495"); "Ln 2, Col 1" with a
        # code window in front (DEV-8/fd1, me1 stage 2, me2 picture 10)
        tb.addSeparator()
        box = QWidget()
        hl = QHBoxLayout(box)
        hl.setContentsMargins(6, 0, 6, 0)
        hl.setSpacing(4)
        self.tb_pos_icon, self.tb_pos = QLabel(), QLabel()
        self.tb_size_icon, self.tb_size = QLabel(), QLabel()
        self.tb_pos_icon.setPixmap(edicons.icon("position").pixmap(16, 16))
        self.tb_size_icon.setPixmap(edicons.icon("size").pixmap(16, 16))
        self.tb_pos.setMinimumWidth(90)
        self.tb_size.setMinimumWidth(90)
        for wd in (self.tb_pos_icon, self.tb_pos, self.tb_size_icon, self.tb_size):
            hl.addWidget(wd)
        tb.addWidget(box)
        self.tb_standard = tb
        # the Edit toolbar of VB6 (View > Toolbars > Edit; hidden at first, as in VB6)
        tb = self.addToolBar("Edit")
        tb.setObjectName("EditToolbar")
        for name, text, slot in (("list_members", "List Properties/Methods", lambda: self._assist("list_members")),
                                 ("list_constants", "List Constants", lambda: self._assist("list_constants")),
                                 ("quick_info", "Quick Info", lambda: self._assist("quick_info")),
                                 ("param_info", "Parameter Info", lambda: self._assist("parameter_info")),
                                 ("complete_word", "Complete Word", lambda: self._assist("complete_word"))):
            a = tb.addAction(edicons.icon(name), text)
            a.triggered.connect(lambda checked=False, f=slot: f())
        tb.addSeparator()
        a = tb.addAction(edicons.icon("indent"), "Indent")
        a.triggered.connect(lambda: self._indent(True))
        a = tb.addAction(edicons.icon("outdent"), "Outdent")
        a.triggered.connect(lambda: self._indent(False))
        tb.addSeparator()
        a = tb.addAction(edicons.icon("breakpoint"), "Toggle Breakpoint")
        a.triggered.connect(self.toggle_breakpoint)
        tb.addSeparator()
        tb.addAction(self.a_comment)
        tb.addAction(self.a_uncomment)
        tb.addSeparator()
        for a in (self.a_bm_toggle, self.a_bm_next, self.a_bm_prev, self.a_bm_clear):
            tb.addAction(a)
        tb.hide()
        self.tb_edit = tb
        # the Debug toolbar of VB6 (DEV-8/me2 picture 16): Start, Break, End | Toggle Breakpoint, Step Into,
        # Step Over, Step Out | Locals, Immediate, Watch, Quick Watch, Call Stack; hidden at first
        tb = self.addToolBar("Debug")
        tb.setObjectName("DebugToolbar")
        for a in (self.a_start, self.a_break, self.a_end, None, self.a_toggle_bp, self.a_step_into,
                  self.a_step_over, self.a_step_out, None, self.a_view_locals, self.a_view_immediate,
                  self.a_view_watch, self.a_quick_watch, self.a_call_stack):
            tb.addSeparator() if a is None else tb.addAction(a)
        for a, name in ((self.a_view_locals, "locals"), (self.a_view_immediate, "immediate"),
                        (self.a_view_watch, "watch")):
            if a.icon().isNull():
                a.setIcon(edicons.icon(name))
        tb.hide()
        self.tb_debug = tb
        # the Form Editor toolbar of VB6 (DEV-8/me2 picture 16): Bring to Front, Send to Back | Align, Center,
        # Make Same Size (a button and its list) | Lock Controls; hidden at first
        tb = self.addToolBar("Form Editor")
        tb.setObjectName("FormEditorToolbar")
        tb.addAction(self.a_front)
        tb.addAction(self.a_back)
        tb.addSeparator()
        self.fe_buttons = []
        for menu, name, tip in ((self.m_align, "align", "Align"), (self.m_center, "center", "Center in Form"),
                                (self.m_same_size, "same_size", "Make Same Size")):
            first = next(a for a in menu.actions() if not a.isSeparator())
            b = QToolButton()
            b.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
            b.setMenu(menu)
            b.setIcon(edicons.icon(name))
            b.setToolTip(tip)
            b.clicked.connect(first.trigger)  # VB6: the button does the first command of its list
            b.setEnabled(first.isEnabled())
            first.changed.connect(lambda bb=b, f=first: bb.setEnabled(f.isEnabled()))
            b.first_action = first
            tb.addWidget(b)
            self.fe_buttons.append(b)
        tb.addSeparator()
        tb.addAction(self.a_lock)
        tb.hide()
        self.tb_form_editor = tb
        tm = self.view_menu.addMenu("&Toolbars")    # VB6: the last item of View (DEV-8/me1 picture 9)
        names = {"Debug": "&Debug", "Edit": "&Edit", "Form Editor": "&Form Editor", "Standard": "&Standard"}
        for t in (self.tb_debug, self.tb_edit, self.tb_form_editor, self.tb_standard):
            a = t.toggleViewAction()
            a.setText(names[t.windowTitle()])
            tm.addAction(a)
        tm.addSeparator()
        self.a_customize = tm.addAction("&Customize...", lambda: self.cmdbars.customize())
        self.m_toolbars = tm

    def _list_button(self, tb, icon, entries):
        """a button of the Standard toolbar with its list (VB6: Add Standard EXE Project, Add Form): the button does
        the command chosen last from the list, and shows it"""
        from PyQt6.QtWidgets import QMenu
        b = QToolButton(tb)
        b.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
        menu = QMenu(b)
        for text, target in entries:
            a = menu.addAction(text)
            if isinstance(target, QAction):
                a.setIcon(target.icon())
                a.triggered.connect(lambda _c=False, t=target: t.trigger())
                target.changed.connect(lambda a=a, t=target: a.setEnabled(t.isEnabled()))
            else:
                a.triggered.connect(lambda _c=False, f=target: f())
            a.triggered.connect(lambda _c=False, a=a, b=b: (b.setToolTip(a.text()), setattr(b, "last", a)))
        b.setMenu(menu)
        b.setIcon(edicons.icon(icon))
        first = menu.actions()[0] if menu.actions() else None
        b.last = first
        b.setToolTip(first.text() if first else "")
        b.clicked.connect(lambda: b.last.trigger() if b.last is not None else None)
        tb.addWidget(b)
        return b

    def _show_dock(self, dock, widget):
        holder = widget.parentWidget()         # Immediate / Locals / Watches out of the bottom dock: their own
        while holder is not None and not isinstance(holder, QDockWidget):
            holder = holder.parentWidget()
        if holder is not None:
            dock = holder
        dock.show()
        dock.raise_()
        if dock is self.dock_bottom:            # the tab that holds the widget
            for i in range(self.bottom_tabs.count()):
                page = self.bottom_tabs.widget(i)
                if page is widget or page.isAncestorOf(widget):
                    self.bottom_tabs.setCurrentIndex(i)
                    break
        widget.setFocus()

    # ================================================================ Dockable (Tools > Options > Docking)
    # the windows of VB6's Docking tab, in its order (DEV-8/me4 picture 1); the Object Browser is not dockable
    # at first (an MDI window)
    TOOL_WINDOWS = [("immediate", "Immediate Window"), ("locals", "Locals Window"), ("watch", "Watch Window"),
                    ("project", "Project Explorer"), ("props", "Properties Window"), ("objbrowser", "Object Browser"),
                    ("layout", "Form Layout"), ("toolbox", "Toolbox"), ("palette", "Color Palette")]

    def is_dockable(self, key):
        if key == "layout":
            return bool(self.S["layout_dockable"])
        d = self.S["dockable"] or {}
        return bool(d.get(key, key != "objbrowser"))

    def set_dockable(self, key, on):
        """Dockable of one tool window (its right button, or the Docking tab): off, it floats and no side of the
        IDE takes it; on, it can dock again (and goes back to its place)"""
        on = bool(on)
        if key == "layout":
            if on != self.is_dockable("layout"):
                self.set_layout_dockable(on)
            return
        d = dict(self.S["dockable"] or {})
        d[key] = on
        self.S["dockable"] = d
        self.S.save()
        self._apply_dockable(key)

    def apply_dockable(self):
        for key, _title in self.TOOL_WINDOWS:
            if key == "layout":
                if self.is_dockable("layout") != (self.layout_sub is None and
                                                  self.dock_layout.allowedAreas() != Qt.DockWidgetArea.NoDockWidgetArea):
                    self._unplace_layout()
                    self._place_layout()
            else:
                self._apply_dockable(key)

    def _tool_dock(self, key):
        return {"project": getattr(self, "dock_project", None), "props": getattr(self, "dock_props", None),
                "dataview": getattr(self, "dock_dataview", None),
                "toolbox": getattr(self, "dock_toolbox", None), "palette": getattr(self, "dock_palette", None)}.get(key)

    def _tool_widget(self, key):
        return {"immediate": self.immediate, "locals": self.locals_win, "watch": self.watch_win}.get(key)

    def _apply_dockable(self, key):
        on = self.is_dockable(key)
        if key == "objbrowser":
            ob = getattr(self, "_objbrowser", None)
            if ob is not None and isinstance(ob, QDockWidget) != on:     # open the other way: shown again as it is now
                self._close_objbrowser()
                self.object_browser()
            return
        dock = self._tool_dock(key)
        if dock is not None:
            if on:
                dock.setAllowedAreas(Qt.DockWidgetArea.AllDockWidgetAreas)
            elif dock.allowedAreas() != Qt.DockWidgetArea.NoDockWidgetArea:
                dock.setAllowedAreas(Qt.DockWidgetArea.NoDockWidgetArea)
                if not dock.isFloating():
                    dock.setFloating(True)
            return
        w = self._tool_widget(key)
        if w is None:
            return
        free = self.free_docks.get(key)
        if not on and free is None:            # out of the tabs of the bottom dock: a free window of its own
            title = dict(self.TOOL_WINDOWS)[key].replace(" Window", "")
            if key == "watch":
                title = "Watches"
            i = self.bottom_tabs.indexOf(w)
            if i >= 0:
                self.bottom_tabs.removeTab(i)
            d = QDockWidget(title, self)
            d.setObjectName("Free" + key.title())
            d.setAllowedAreas(Qt.DockWidgetArea.NoDockWidgetArea)
            d.setWidget(w)
            w.show()
            self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, d)
            d.setFloating(True)
            d.resize(520, 200)
            g = self.geometry()
            d.move(g.x() + g.width() // 4, g.y() + g.height() // 2)
            d.show()
            self.free_docks[key] = d
        elif on and free is not None:           # back to its tab of the bottom dock
            free.setWidget(None)
            names = {"immediate": "Immediate", "locals": "Locals", "watch": "Watches"}
            order = [self.output, self.immediate, self.locals_win, self.watch_win, self.references]
            pos = sum(1 for i in range(self.bottom_tabs.count())
                      if self.bottom_tabs.widget(i) in order and
                      order.index(self.bottom_tabs.widget(i)) < order.index(w))
            self.bottom_tabs.insertTab(pos, w, names[key])
            self.removeDockWidget(free)
            free.deleteLater()
            del self.free_docks[key]
            self._show_dock(self.dock_bottom, w)

    def _close_objbrowser(self):
        ob = getattr(self, "_objbrowser", None)
        if ob is None:
            return
        if isinstance(ob, QDockWidget):
            self.removeDockWidget(ob)
            ob.deleteLater()
        else:
            ob.close()
        self._objbrowser = None

    def tool_menu(self, key, widget, extras=None):
        """the right button of a tool window: its own commands (extras: [(text, slot, enabled[, icon[, checked]])
        or None]), then Dockable and Hide, as in VB6 (DEV-8/me4 pictures 2, 3; me8)"""
        widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)

        def show(pos):
            from PyQt6.QtWidgets import QMenu
            menu = QMenu(self)
            for e in (extras() if callable(extras) else extras or []):
                if e is None:
                    menu.addSeparator()
                else:
                    a = menu.addAction(e[0])
                    if e[1] is not None:
                        a.triggered.connect(lambda _c=False, f=e[1]: f())
                    a.setEnabled(bool(e[2]))
                    ic = e[3] if len(e) > 3 else None
                    if ic is not None:
                        a.setIcon(ic if isinstance(ic, QIcon) else edicons.icon(ic))
                    if len(e) > 4 and e[4] is not None:
                        a.setCheckable(True)
                        a.setChecked(bool(e[4]))
            if not menu.isEmpty():
                menu.addSeparator()
            a = menu.addAction("Doc&kable", lambda: self.set_dockable(key, not self.is_dockable(key)))
            a.setCheckable(True)
            a.setChecked(self.is_dockable(key))
            menu.addAction("&Hide", lambda: self.hide_tool(key, widget))
            self._last_tool_menu = menu
            self.shortcut_menu(self.TOOL_SHORTCUTS.get(key), menu)
            menu.exec(widget.mapToGlobal(pos))
        widget.customContextMenuRequested.connect(show)

    # the names of the context menus on the Shortcut Menus toolbar of Customize (phase 12.2)
    TOOL_SHORTCUTS = {"props": "Property Browser", "immediate": "Immediate Window", "locals": "Locals Window",
                      "watch": "Watch Window", "toolbox": "Toolbox", "palette": "Color Palette",
                      "layout": "Form Layout Window", "objbrowser": "Object Browser", "project": "Project Window"}

    def shortcut_menu(self, name, menu):
        """a context menu as the user changed it in Customize (its name on the Shortcut Menus toolbar)"""
        cb = getattr(self, "cmdbars", None)
        if cb is not None and name:
            cb.shortcut(name, menu)
        return menu

    def createPopupMenu(self):
        """the right button on the toolbars and the menu bar: the toolbars, Customize... (VB6)"""
        cb = getattr(self, "cmdbars", None)
        return cb.popup_menu() if cb is not None else super().createPopupMenu()

    def hide_tool(self, key, widget):
        holder = widget
        while holder is not None and not isinstance(holder, QDockWidget):
            holder = holder.parentWidget()
        if holder is not None:
            holder.hide()

    def _toolbox_menu(self):
        """VB6's menu of the Toolbox (on a tool or on its empty space): Components..., Add Tab... | Dockable,
        Hide (DEV-8/me6 pictures 1, 2)"""
        return [("C&omponents...", self.components_dialog, self.project is not None),
                ("&Add Tab...", lambda: self.toolbox.add_tab(), True)]

    def _run_items(self):
        """Start (Continue in break) and End, with their pictures, in the menus of Watches / Locals (me8-17..23)"""
        state = self.runner.state
        return [("&Continue" if state == "break" else "&Start", self.run_start, state in ("design", "break"),
                 self.a_start.icon()),
                ("&End", self.run_end, state != "design", self.a_end.icon())]

    def _props_menu(self):
        """the menu of the Properties window (its list and its description): Description | Dockable, Hide
        (DEV-8/me8-03, 04)"""
        return [("&Description", lambda: self.set_props_description(not self.props.desc.isVisible()), True, None,
                 self.props.desc.isVisible())]

    def set_props_description(self, on):
        """Properties > Description: the box of the description under the list, shown or not (kept)"""
        self.props.desc.setVisible(bool(on))
        self.S["props_description"] = bool(on)

    def _watch_menu(self):
        """the menu of Watches (DEV-8/me8-17, 18, 22): Cut, Copy, Paste | Edit Watch..., Add Watch..., Delete Watch |
        Collapse Parent | Start / Continue, End | Definition (then Dockable, Hide)"""
        i = self.watch_win.current()
        return [("Cu&t", None, False, "cut"), ("&Copy", self.watch_win.copy, i >= 0, "copy"),
                ("&Paste", None, False, "paste"), None,
                ("&Edit Watch...", lambda: self.edit_watch(i), i >= 0), ("&Add Watch...", self.add_watch, True),
                ("&Delete Watch", lambda: self.delete_watch(i), i >= 0), None,
                ("Co&llapse Parent", None, False), None] + self._run_items() + [None, ("&Definition", None, False)]

    def _locals_menu(self):
        """the menu of Locals (DEV-8/me8-19, 20, 23): Cut, Copy, Paste | Collapse Parent (a line inside an other
        one) | Start / Continue, End (then Dockable, Hide)"""
        lw = self.locals_win
        return [("Cu&t", None, False, "cut"), ("&Copy", lw.copy, lw.tree.currentItem() is not None, "copy"),
                ("&Paste", None, False, "paste"), None,
                ("Co&llapse Parent", lw.collapse_parent, lw.can_collapse_parent()), None] + self._run_items()

    # ================================================================ the menu of a form designer (DEV-8/me8)
    def designer_menu_items(self, d, item):
        """[(icon name or None, text, slot, enabled, checked or None) or None] of the right button in a designer:
        the form (item None): View Code, Menu Editor... | Lock Controls, Paste | Properties | Update UserControls
        (me8-01, 05, 12); a control: Cut, Copy, Paste, Delete | Bring to Front, Send to Back | View Code, Align to
        Grid | Properties (me8-02, 06..09); an ActiveX control (a user control, an .ocx): the same without the
        pictures and without Properties, \"Bring To Front\", \"Send To Back\" (me8-13)"""
        paste = d.can_paste()

        def after(f):
            def run():
                f()
                self.update_format_menu()
                self.update_toolbar_pos()
            return run
        if item is None:
            return [("view_code", "View C&ode", self.view_code, True, None),
                    ("menu_editor", "&Menu Editor...", after(d.menu_editor), True, None), None,
                    ("lock", "&Lock Controls", after(d.toggle_lock), True, bool(d.locked)),
                    ("paste", "&Paste", after(d.paste), paste, None), None,
                    ("properties", "Prope&rties", self.show_properties, True, None), None,
                    (None, "Update &UserControls", d.update_user_controls, bool(d.user_controls_used()), None)]
        front, back = d.z_order_state()
        cut = after(lambda: (d.copy_selection(), d.delete_selection()))
        if item.ctl.cls.lower().startswith("vb."):
            return [("cut", "Cu&t", cut, True, None), ("copy", "&Copy", d.copy_selection, True, None),
                    ("paste", "&Paste", after(d.paste), paste, None),
                    (None, "&Delete", after(d.delete_selection), True, None), None,
                    ("front", "&Bring to Front", after(lambda: d.z_order(True)), front, None),
                    ("back", "&Send to Back", after(lambda: d.z_order(False)), back, None), None,
                    ("view_code", "View C&ode", self.view_code, True, None),
                    ("align_grid", "Align to &Grid", after(lambda: d.align("grid")), True, None), None,
                    ("properties", "Prope&rties", self.show_properties, True, None)]
        from .ppdesign import uc_pages_of
        from .ocxpages import pages_of
        pages = [None, (None, "P&roperties", self.property_pages, True, None)] \
            if uc_pages_of(self, d, item.ctl) is not None or pages_of(item.ctl.type_name) else []  # pp1-16, phase 11
        return [(None, "Cu&t", cut, True, None), (None, "&Copy", d.copy_selection, True, None),
                (None, "&Paste", after(d.paste), paste, None), (None, "&Delete", after(d.delete_selection), True, None),
                None, (None, "&Bring To Front", after(lambda: d.z_order(True)), front, None),
                (None, "Se&nd To Back", after(lambda: d.z_order(False)), back, None), None,
                (None, "&View Code", self.view_code, True, None),
                (None, "Align to &Grid", after(lambda: d.align("grid")), True, None)] + pages

    def designer_menu(self, d, gpos, item):
        from PyQt6.QtWidgets import QMenu
        menu = QMenu(self)
        for e in self.designer_menu_items(d, item):
            if e is None:
                menu.addSeparator()
                continue
            ic, text, slot, on, checked = e
            a = menu.addAction(text)
            if ic:
                a.setIcon(edicons.icon(ic))
            a.triggered.connect(lambda _c=False, f=slot: f())
            a.setEnabled(bool(on))
            if checked is not None:
                a.setCheckable(True)
                a.setChecked(checked)
        self._last_designer_menu = menu
        self.shortcut_menu("Forms" if item is None else "Controls", menu)
        menu.exec(gpos)

    def show_properties(self):
        """Properties of the menus: the Properties window, in front"""
        self._show_dock(self.dock_props, self.props.grid)

    def _immediate_menu(self):
        """VB6's menu of the Immediate window: Cut, Copy, Paste | Object Browser, Definition | Start, End"""
        im = self.immediate
        sel = im.textCursor().hasSelection()
        state = self.runner.state
        return [("Cu&t", im.cut, sel), ("&Copy", im.copy, sel), ("&Paste", im.paste, im.canPaste()), None,
                ("&Object Browser", self.object_browser, True), ("&Definition", self.browse_word, False), None,
                ("&Start", self.run_start, state in ("design", "break")), ("&End", self.run_end, state != "design")]

    # ================================================================ Form Layout: Dockable
    def set_layout_dockable(self, on):
        """the right button of the Form Layout > Dockable (DEV-8/fd1 stage 20): off, the window does not dock
        when dragged; on, back to its dock"""
        self.S["layout_dockable"] = bool(on)
        self.S.save()
        self._unplace_layout()
        self._place_layout()

    def _unplace_layout(self):
        """the Form Layout back in its dock, as it is when Dockable"""
        fl, dock = self.formlayout, self.dock_layout
        visible = self.layout_visible()
        sub = self.layout_sub
        if sub is not None:
            self.layout_sub = None
            sub.setWidget(None)
            self.mdi.removeSubWindow(sub)
            sub.deleteLater()
            dock.setWidget(fl)
            fl.show()
        dock.setAllowedAreas(Qt.DockWidgetArea.AllDockWidgetAreas)
        if dock.isFloating():
            self._layout_moving = True
            try:
                dock.setFloating(False)
            finally:
                self._layout_moving = False
        dock.setVisible(visible)

    def _place_layout(self):
        """without Dockable: in the MDI view a small window among the code and form windows, as in VB6
        (picture 31); with the windows as tabs a free window, so that it does not take the place of the form"""
        if self.S["layout_dockable"]:
            return
        fl, dock = self.formlayout, self.dock_layout
        visible = self.layout_visible()
        if self.mdi.viewMode() == QMdiArea.ViewMode.TabbedView:
            dock.setAllowedAreas(Qt.DockWidgetArea.NoDockWidgetArea)
            if not dock.isFloating():           # a free window keeps the place it was left at
                dock.setFloating(True)
                dock.resize(260, 200)
                dock.move(self.mdi.viewport().mapToGlobal(self._layout_spot(260, 200)))
            dock.setVisible(visible)
            return
        dock.hide()
        if self.layout_sub is not None:
            return
        spot = self._layout_spot(260, 200)
        sub = QMdiSubWindow()
        sub.setWidget(fl)
        sub.setWindowTitle("Form Layout")
        self.mdi.addSubWindow(sub)
        self.layout_sub = sub
        fl.show()
        sub.showNormal()
        sub.resize(260, 200)
        sub.move(spot)
        sub.setVisible(visible)

    def _layout_spot(self, w, h):
        """where the Form Layout goes out of its dock, in the MDI area: beside the form being designed, as VB6
        puts it (DEV-8/fd1 stage 20, picture 31); the middle of the area without a form"""
        from PyQt6.QtCore import QPoint
        vp = self.mdi.viewport()
        d = self.current_designer()
        if d is not None and getattr(d, "frame_w", None) is not None and d.frame_w.isVisible():
            f = d.frame_w
            p = f.mapTo(vp, QPoint(f.width() + 40, max(0, f.height() // 2 - h // 2)))
        else:
            p = QPoint((vp.width() - w) // 2, (vp.height() - h) // 2)
        x = min(max(0, p.x()), max(0, vp.width() - w))
        y = min(max(0, p.y()), max(0, vp.height() - h))
        return QPoint(x, y)

    def _layout_top_level(self, floating):
        """a double click on the title of the free window would dock it: not without Dockable"""
        if not floating and not self.S["layout_dockable"] and self.layout_sub is None and \
                not getattr(self, "_layout_moving", False):
            QTimer.singleShot(0, lambda: self.dock_layout.setFloating(True))

    def layout_visible(self):
        if self.layout_sub is not None:
            return not self.layout_sub.isHidden()
        return not self.dock_layout.isHidden()

    def show_layout(self):
        """View > Form Layout Window"""
        if self.layout_sub is not None:
            self.layout_sub.show()
            self.mdi.setActiveSubWindow(self.layout_sub)
            self.formlayout.setFocus()
        else:
            self._show_dock(self.dock_layout, self.formlayout)

    def property_pages(self, quiet=False):
        """View > Property Pages (Shift+F4): the pages of a user control on a form (its PropertyPages, 9.2e:
        ppdesign.py) or of an ActiveX control (phase 11: ocxpages.py); the other controls have none, VB6's "None"
        (DEV-8/me2 picture 11)"""
        from .ppdesign import DesignPagesDialog, uc_pages_of
        d = self.current_designer()
        item = getattr(d, "primary", None) if d is not None else None
        ctl = getattr(item, "ctl", None)
        uc = uc_pages_of(self, d, ctl) if ctl is not None else None
        if uc is not None:
            dlg = DesignPagesDialog(self, d, ctl, uc)
            dlg.quiet = quiet
            self._last_ppages = dlg
            if not quiet:
                dlg.exec()
            return dlg
        from .ocxpages import OcxPagesDialog, pages_of
        if ctl is not None and pages_of(ctl.type_name):     # an ActiveX control with property pages (phase 11)
            dlg = OcxPagesDialog(self, d, ctl)
            dlg.quiet = quiet
            self._last_ppages = dlg
            if not quiet:
                dlg.exec()
            return dlg
        from .dialogs import PropertyPagesDialog
        dlg = PropertyPagesDialog(self)
        self._last_ppages = dlg
        dlg.exec()

    def show_palette(self):
        """View > Color Palette: the colours of the selected control (or the form)"""
        d = self.dock_palette
        first = not d.isVisible() and not getattr(self, "_palette_placed", False)
        d.show()
        d.raise_()
        if first and d.isFloating():           # the first time: over the middle of the IDE, as VB6 opens it
            self._palette_placed = True
            d.adjustSize()
            g = self.geometry()
            d.move(g.x() + (g.width() - d.width()) // 2, g.y() + g.height() // 3)
        self.palette_win.refresh()

    def hide_layout(self):
        """the right button of the Form Layout > Hide"""
        if self.layout_sub is not None:
            self.layout_sub.hide()
        else:
            self.dock_layout.hide()

    # ================================================================ Output
    def show_output(self):
        self.dock_bottom.show()
        self.dock_bottom.raise_()
        self.bottom_tabs.setCurrentWidget(self.output)

    def message(self, text, kind="info", where=None, show=False):
        """a line in the Output window (errors there, not in message boxes)"""
        self.output.add(text, kind, where)
        if show:
            self.show_output()

    def _module_of_file(self, rel):
        if not rel or self.project is None:
            return None
        m = self.project.module_by_path(self.project.resolve(rel)) if self.project.path else None
        if m is None:
            base = os.path.basename(rel.replace(chr(92), "/")).lower()
            m = next((x for x in self.project.modules if os.path.basename(x.rel.replace(chr(92), "/")).lower() == base), None)
        return m

    def goto_message(self, rel, line, col):
        """double click on a message of the Output window: the code at its line"""
        m = self.project.module(rel) if self.project is not None else None
        m = m or self._module_of_file(rel)
        if m is None:
            return
        self.open_code(m, max(0, line - 1 - m.code_offset()))

    # ================================================================ editor host (Linpad++ Editor)
    def _load_styles(self):
        dark = self.is_dark()
        path = paths.data("themes", "DarkModeDefault.xml") if dark else paths.user_or_model("stylers.xml", "stylers.model.xml")
        return stl.load_styleset(path)

    def setup_code_editor(self, ed):
        snd = ed.SendScintilla
        lg = self.langdb.get("vb")
        snd(S.SCI_SETLEXERLANGUAGE, 0, lg.lexer.encode())
        for k, v in self.langdb.props("vb").items():
            snd(S.SCI_SETPROPERTY, k.encode(), v.encode())
        for i, words in self.langdb.keyword_lists("vb", self.styles).items():
            snd(S.SCI_SETKEYWORDS, i, words.encode("utf-8"))
        self.styles.apply(ed, self.langdb, "vb", self.S)
        ed.apply_view_settings(self.S)
        snd(S.SCI_SETTABWIDTH, 4)
        snd(S.SCI_SETUSETABS, False)            # VB6 indents with spaces
        snd(S.SCI_SETEOLMODE, S.SC_EOL_CRLF)
        snd(S.SCI_ASSIGNCMDKEY, ord("Y") + (S.SCMOD_CTRL << 16), S.SCI_LINECUT)   # VB6: Ctrl+Y cuts the line
        ed.setAutoIndent(True)
        if not getattr(ed, "_ctx", False):              # setup_code_editor runs again after Options / a theme change
            ed._ctx = True
            ed.context_requested.connect(self.code_menu)
        snd(S.SCI_COLOURISE, 0, -1)
        # the debugger: breakpoints (dark red), the next statement (yellow), a line of the call stack (green)
        self._bp_style(ed)
        ed.markerDefine(edicons.bookmark_pixmap(), stl.MARK_BOOKMARK)     # VB6's cyan bookmark
        snd(S.SCI_MARKERDEFINE, MARK_CUR, S.SC_MARK_SHORTARROW)
        snd(S.SCI_MARKERSETFORE, MARK_CUR, QColor("#000000"))
        snd(S.SCI_MARKERSETBACK, MARK_CUR, QColor("#FFFF00"))
        snd(S.SCI_MARKERDEFINE, MARK_CUR_BACK, S.SC_MARK_BACKGROUND)
        snd(S.SCI_MARKERSETBACK, MARK_CUR_BACK, QColor("#FFFF00"))
        snd(S.SCI_MARKERDEFINE, MARK_CALL, S.SC_MARK_BACKGROUND)
        snd(S.SCI_MARKERSETBACK, MARK_CALL, QColor("#B4E8B4"))
        snd(S.SCI_SETCARETLINEBACKALPHA, 70)        # the yellow and red lines show through the caret line
        mask = snd(S.SCI_GETMARGINMASKN, 1)
        snd(S.SCI_SETMARGINMASKN, 1, mask | (1 << MARK_BP) | (1 << MARK_CUR))
        snd(S.SCI_SETMARGINWIDTHN, 1, max(16, snd(S.SCI_GETMARGINWIDTHN, 1)))
        snd(S.SCI_SETREADONLY, self.runner.state in ("running", "break"))
        z = self._code_zoom()
        if snd(S.SCI_GETZOOM) != z:             # the zoom of the last session (on_zoom keeps it)
            snd(S.SCI_SETZOOM, z)
        if not getattr(ed, "_dwell", False):
            ed._dwell = True
            ed.SCN_MODIFYATTEMPTRO.connect(lambda e=ed: self.edit_while_running(e))
            ed.SCN_DWELLSTART.connect(lambda pos, x, y, e=ed: self.data_tip(e, pos))
            ed.SCN_DWELLEND.connect(lambda *a, e=ed: e.SendScintilla(S.SCI_CALLTIPCANCEL) if self.runner.state == "break" else None)

    def on_savepoint(self, ed, dirty):
        cw = self._cw_of(ed)
        if cw is not None:
            if dirty:
                cw.module.dirty = True
            self.update_tab_title(cw)
            self.update_tree_marks()

    def on_editor_idle(self, ed):
        pass

    def on_editor_ui(self, ed):
        self.update_status()
        cw = self._cw_of(ed)
        if cw is not None:
            cw.assist.caret_moved()

    def on_char_added(self, ed, ch):
        cw = self._cw_of(ed)
        if cw is not None:
            cw.assist.char_added(ch)

    def intellisense(self):
        """the IntelliSense engine of the open project (made again when the project changes)"""
        from .intellisense import IntelliSense
        eng = getattr(self, "_intellisense", None)
        if eng is None or eng.project is not self.project:
            eng = self._intellisense = IntelliSense(self.project)
        return eng

    def _assist(self, what):
        cw = self.current_cw()
        if cw is not None:
            getattr(cw.assist, what)()
            return True
        return False

    def on_text_modified(self, ed, pos, mtype, length, added):
        cw = self._cw_of(ed)
        if cw is not None:
            cw.module.set_code(None)            # read from the editor when needed

    def on_double_click(self, ed, pos, line, mods):
        pass

    def _code_zoom(self):
        try:
            return max(-10, min(60, int(self.S["code_zoom"] or 0)))
        except (TypeError, ValueError):
            return 0

    def on_zoom(self, ed):
        """Ctrl+Mouse Wheel in a code window: every code window (and both panes of a split one) takes
        the same zoom, and it is saved with the settings so the next session starts with it, as
        Notepad++ / Linpad++ keep their zoom"""
        if self._cw_of(ed) is None:
            return
        z = int(ed.SendScintilla(S.SCI_GETZOOM))
        if z == self._code_zoom():
            return
        self.set_code_zoom(z)

    def code_zoom_step(self, step):
        """View > Zoom: Zoom In (+1), Zoom Out (-1), Zoom Reset (0) of the code windows"""
        z = 0 if step == 0 else max(-10, min(20, self._code_zoom() + step))
        self.set_code_zoom(z)

    def set_code_zoom(self, z):
        """every code window (both panes of a split one) at zoom z; kept in the settings for the next session"""
        self.S["code_zoom"] = z
        for sw in list(self.code_windows.values()):
            cw = sw.widget()
            if isinstance(cw, CodeWindow):
                for e in cw.panes():
                    if e.SendScintilla(S.SCI_GETZOOM) != z:
                        e.SendScintilla(S.SCI_SETZOOM, z)

    def on_macro_record(self, ed, msg, wparam, lparam):
        pass

    def on_key_press(self, ed, e):
        return False

    def delimiter_select(self, ed, e):
        return False

    def toggle_bookmark(self, ed, line):
        """a click in the margin: a breakpoint, as on the Margin Indicator Bar of VB6"""
        cw = self._cw_of(ed)
        if cw is None:
            return
        ed.goto_pos(ed.line_start(line))
        self.mdi.setActiveSubWindow(self.code_windows.get(cw.module))
        self.toggle_breakpoint()

    def _cw_of(self, ed):
        p = ed.parent()
        while p is not None and not isinstance(p, CodeWindow):     # the pane sits in the splitter (Split)
            p = p.parent()
        return p

    # ================================================================ code windows
    def code_menu(self, ed, pos):
        """the menu of a code window: the edit commands and the refactoring ones (VB6's code menu)"""
        from PyQt6.QtWidgets import QMenu
        cw = self._cw_of(ed)
        if cw is None:
            return
        self.mdi.setActiveSubWindow(self.code_windows.get(cw.module))
        menu = QMenu(self)
        menu.addAction("Cu&t", ed.cut)
        menu.addAction("&Copy", ed.copy)
        menu.addAction("&Paste", ed.paste)
        menu.addSeparator()
        menu.addAction("&Definition\tShift+F2", lambda: RF.definition(self))
        menu.addAction("&Last Position\tCtrl+Shift+F2", lambda: RF.last_position(self))
        menu.addAction("Find Re&ferences\tShift+F12", lambda: RF.find_references(self))
        menu.addAction("Rena&me Symbol...\tCtrl+Shift+R", lambda: RF.rename(self))
        menu.addAction("E&xtract Procedure...\tCtrl+Shift+M", lambda: RF.extract(self))
        menu.addSeparator()
        menu.addAction("&Quick Info\tCtrl+I", lambda: self._assist("quick_info"))
        menu.addAction("List Properties/&Methods\tCtrl+J", lambda: self._assist("list_members"))
        menu.addSeparator()
        menu.addMenu(self.toggle_menu(menu))
        self.shortcut_menu("Code Window (Break)" if self.runner.state == "break" else "Code Window", menu)
        menu.exec(pos)

    def toggle_menu(self, parent):
        """the Toggle submenu of the code window: a breakpoint, and the Error Trapping of this session (VB6:
        the setting of Tools > Options stays as it is)"""
        from PyQt6.QtWidgets import QMenu
        from PyQt6.QtGui import QActionGroup
        sub = QMenu("To&ggle", parent)
        sub.addAction("&Breakpoint\tF9", self.toggle_breakpoint)
        sub.addAction("Boo&kmark", self.toggle_code_bookmark)
        sub.addSeparator()
        grp = QActionGroup(sub)
        for n, text in enumerate(("Break on &All Errors", "Break in &Class Module", "Break on &Unhandled Errors")):
            a = sub.addAction(text, lambda n=n: self.set_trap(n))
            a.setCheckable(True)
            a.setChecked(self.trap_mode() == n)
            grp.addAction(a)
        return sub

    def current_sub(self):
        """the window in front: when the IDE comes back from another program with a click straight on a menu,
        QMdiArea has no current window yet while the menu opens (DEV-8/fd1 visual check), so the last one
        activated stands for it"""
        sw = self.mdi.currentSubWindow()
        if sw is None:
            order = self.mdi.subWindowList(QMdiArea.WindowOrder.ActivationHistoryOrder)
            sw = order[-1] if order else None
        return sw

    def current_cw(self):
        sw = self.current_sub()
        return sw.widget() if sw is not None and isinstance(sw.widget(), CodeWindow) else None

    def current_editor(self):
        cw = self.current_cw()
        return cw.editor if cw else None

    def _ed_call(self, name):
        d = self.current_designer()
        if d is not None:
            {"undo": d.undo, "cut": lambda: (d.copy_selection(), d.delete_selection()), "copy": d.copy_selection,
             "paste": d.paste, "removeSelectedText": d.delete_selection,
             "selectAll": lambda: d.select([it for it in d.items.values() if it.parentWidget() is d.surface]),
             "redo": lambda: None}[name]()
            return
        w = QApplication.focusWidget()
        if isinstance(w, QPlainTextEdit) and hasattr(w, name):
            getattr(w, name)()
            return
        ed = self.current_editor()
        if ed is not None:
            getattr(ed, name)()

    def _indent(self, more):
        ed = self.current_editor()
        if ed is not None:
            ed.SendScintilla(S.SCI_TAB if more else S.SCI_BACKTAB)

    def open_code(self, m, line=None):
        sw = self.code_windows.get(m)
        if sw is None:
            cw = CodeWindow(self, m)
            sw = CodeSubWindow(self)
            sw.setWidget(cw)
            self.mdi.addSubWindow(sw)
            sw.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
            self.code_windows[m] = sw
            self.restore_breakpoints(cw)
            self.update_tab_title(cw)
            sw.resize(800, 520)
            sw.show()
        self.mdi.setActiveSubWindow(sw)
        self.select_in_tree(m)
        cw = sw.widget()
        if line is not None:
            cw.goto_line(line)
        else:
            cw.editor.setFocus()
        return cw

    def code_closing(self, sw):
        """a code window closes: its text stays in the module"""
        cw = sw.widget()
        if isinstance(cw, CodeWindow):
            cw.module.set_code(cw.text())
            self._bookmarks(cw.module)
            self._breakpoints(cw.module)
            self.code_windows.pop(cw.module, None)

    def _sync_editors(self):
        """the text of the open code windows into the modules"""
        for m, sw in list(self.code_windows.items()):
            cw = sw.widget()
            if isinstance(cw, CodeWindow):
                m.set_code(cw.text())

    def update_tab_title(self, cw):
        sw = self.code_windows.get(cw.module)
        if sw is None:
            return
        m = cw.module
        kind = {"Form": "(Code)", "Module": "(Code)", "Class": "(Code)"}.get(m.kind, "(Code)")
        sw.setWindowTitle(f"{m.project.name} - {m.name} {kind}" + (" *" if m.dirty else ""))

    def update_status(self):
        ed = self.current_editor()
        if ed is None:
            self.pos_label.setText("")
            self.update_toolbar_pos()
            return
        p = ed.pos()
        line = ed.line_from_pos(p)
        col = ed.SendScintilla(S.SCI_GETCOLUMN, p)
        self.pos_label.setText(f"Ln {line + 1}, Col {col + 1}")
        self.update_toolbar_pos()
        self._update_file_actions()

    def toolbar_pos_text(self):
        """what the Standard toolbar shows: ("Ln 2, Col 1", None) for a code window, ("1680, 1320",
        "1215 x 495") for the selected control or the form ("0, 0", "4800 x 3600": its outer size), (None, None)"""
        cw = self.current_cw()
        if cw is not None:
            ed = cw.editor
            p = ed.pos()
            return f"Ln {ed.line_from_pos(p) + 1}, Col {ed.SendScintilla(S.SCI_GETCOLUMN, p) + 1}", None
        d = self.current_designer()
        if d is None or getattr(d, "root", None) is None:
            return None, None
        from .designer import geti
        from . import vbprops as V
        it = d.primary
        if it is None:                          # the form (an MDI form, a user control ...): its outer rectangle
            root = d.root
            if root.type_name in ("Form", "MDIForm"):
                x, y, w, h = V.outer_rect(root)
            else:
                x, y, w, h = 0, 0, geti(root, "ClientWidth"), geti(root, "ClientHeight")
        else:
            r = it.rect_twips()
            x, y, w, h = r.x(), r.y(), r.width(), r.height()
        return f"{int(x)}, {int(y)}", f"{int(w)} x {int(h)}"

    def update_toolbar_pos(self):
        if not hasattr(self, "tb_pos"):
            return
        pos, size = self.toolbar_pos_text()
        self.tb_pos.setText(pos or "")
        self.tb_size.setText(size or "")
        self.tb_pos_icon.setVisible(size is not None)
        self.tb_size_icon.setVisible(size is not None)

    def view_code(self):
        d = self.current_designer()
        m = d.module if d is not None else self.selected_module()
        if m is not None:
            self.open_code(m)

    def view_object(self):
        cw = self.current_cw()
        m = cw.module if cw is not None and self.current_sub() and self.current_sub().widget() is cw else self.selected_module()
        if m is not None and m.has_designer:
            self.open_designer(m)

    # ================================================================ form designers
    def current_query(self):
        """the Query Designer in front (the Query and Diagram menus act on it)"""
        sw = self.current_sub()
        from .querydesign import QueryDesigner
        return sw.widget() if sw is not None and isinstance(sw.widget(), QueryDesigner) else None

    def add_query_window(self, w):
        """a window of the Query Designer among the code and form windows ("Design: Command1", "Run Table: ...")"""
        sw = QuerySubWindow(self)
        sw.setWidget(w)
        sw.setWindowTitle(w.title)
        self.mdi.addSubWindow(sw)
        sw.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        sw.resize(860, 600)
        sw.show()
        self.mdi.setActiveSubWindow(sw)
        self.query_menus.update()
        return w

    def show_data_view(self):
        """View > Data View Window"""
        if not self.dock_dataview.isVisible():
            self.dataview.fill()
        self._show_dock(self.dock_dataview, self.dataview.tree)
        return self.dataview

    def current_designer(self):
        sw = self.current_sub()
        return sw.widget() if sw is not None and isinstance(sw.widget(), FormDesigner) else None

    def designer_title(self, m):
        kind = m.designer_kind or (m.kind if m.kind in ("UserControl", "PropertyPage") else
                                   "MDIForm" if m.display_kind == "MDIForm" else "Form")
        return f"{m.project.name} - {m.name} ({kind})" + (" *" if m.dirty else "")

    def open_designer(self, m):
        if m.sf is None or m.sf.root is None:
            return None
        sw = self.design_windows.get(m)
        if sw is None and m.kind == "Designer":
            from .dsrdesign import designer_for
            d = designer_for(self, m)
            if d is None:
                return self.open_code(m)
            d.modified.connect(self.on_designer_modified)
            sw = DesignSubWindow(self)
            sw.setWidget(d)
            self.mdi.addSubWindow(sw)
            sw.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
            self.design_windows[m] = sw
            sw.setWindowTitle(self.designer_title(m))
            sw.resize(880, 560)
            sw.show()
        if sw is None:
            d = FormDesigner(self, m)
            d.selection_changed.connect(lambda ctl, dd=d: self.on_designer_selection(dd, ctl))
            d.modified.connect(self.on_designer_modified)
            d.geometry_changed.connect(self.update_toolbar_pos)
            sw = DesignSubWindow(self)
            sw.setWidget(d)
            self.mdi.addSubWindow(sw)
            sw.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
            self.design_windows[m] = sw
            sw.setWindowTitle(self.designer_title(m))
            sw.resize(820, 560)
            sw.show()
        self.mdi.setActiveSubWindow(sw)
        self.select_in_tree(m)
        sw.widget().surface.setFocus()
        return sw.widget()

    def _sub_activated(self, sw):
        self.update_status()
        if getattr(self, "query_menus", None) is not None:
            self.query_menus.update()
        if getattr(self, "palette_win", None) is not None:
            self.palette_win.show_for()
        self.update_format_menu()               # the Form Editor toolbar follows the window in front
        if getattr(self, "formlayout", None) is not None:
            self.formlayout.update()
        if sw is not None and isinstance(sw.widget(), FormDesigner):
            d = sw.widget()
            self.select_in_tree(d.module)
            self.on_designer_selection(d, d.primary.ctl if d.primary else d.root)
        elif sw is not None and sw in self.design_windows.values():
            self.select_in_tree(sw.widget().module)

    def _designer_do(self, op, arg):
        d = self.current_designer()
        if d is not None:
            if op == "spacing":
                d.spacing(*arg)
            elif arg is None:
                getattr(d, op)()
            else:
                getattr(d, op)(arg)
        self.update_format_menu()

    def update_format_menu(self):
        """Format: what is on for the selection of the designer (DEV-8/fd1 stages 1, 19)"""
        if not hasattr(self, "fmt_actions"):
            return
        d = self.current_designer()
        st = d.format_state() if d is not None and hasattr(d, "format_state") else None
        for need, acts in self.fmt_actions.items():
            for a in acts:
                a.setEnabled(bool(st and st[need]))
        self.a_lock.setEnabled(st is not None)
        self.a_lock.setChecked(bool(st and st["locked"]))

    def apply_grid_settings(self):
        """Tools > Options > General: Form Grid Settings, for the open designers"""
        for sw in self.design_windows.values():
            w = sw.widget()
            if hasattr(w, "grid_settings"):
                w.grid_settings(self.S)

    def on_designer_selection(self, d, ctl):
        """the Properties window shows the selected control"""
        if self.selected_module() is not d.module:
            self.select_in_tree(d.module)
        self.props.show_object(d.module, d, ctl)
        self.update_format_menu()
        if getattr(self, "palette_win", None) is not None:   # (the first form is made before the palette)
            self.palette_win.show_for(d, ctl)
        self.update_toolbar_pos()
        self.addins_sync()

    def on_designer_modified(self, m):
        if getattr(self, "formlayout", None) is not None:
            self.formlayout.update()
        sw = self.design_windows.get(m)
        if sw is not None:
            sw.setWindowTitle(self.designer_title(m))
        if m.kind == "UserControl":             # the forms that have it draw it again
            from .designer import forget_uctl_look, uctl_of
            forget_uctl_look(m)
            for om, osw in self.design_windows.items():
                od = osw.widget()
                if om is not m and od is not None and om.sf is not None and om.sf.root is not None and any(
                        (uctl_of(c.cls) or (None,) * 5)[4] is m for c in om.sf.root.walk()):
                    od.build()
        cw = self.code_windows.get(m)
        if cw is not None:
            cw.widget().fill_objects()          # new controls in the Object box
        d = sw.widget() if sw is not None else None
        self.fill_properties()
        if isinstance(d, FormDesigner):
            self.on_designer_selection(d, d.primary.ctl if d.primary else d.root)
        self.update_tree_marks()

    def open_event_procedure(self, m, ctl):
        """double click in the designer: the default event of the control in the code window"""
        cw = self.open_code(m)
        name = "MDIForm" if ctl is None and m.display_kind == "MDIForm" else ("Form" if ctl is None else ctl.name)
        i = cw.obj_box.findText(name, Qt.MatchFlag.MatchFixedString)
        if i >= 0:
            cw.obj_box.setCurrentIndex(i)
            cw._object_chosen(i)

    # ================================================================ project explorer, properties
    PROJECT_ROLE = Qt.ItemDataRole.UserRole + 1
    RELDOC_ROLE = Qt.ItemDataRole.UserRole + 2
    KIND_ROLE = Qt.ItemDataRole.UserRole + 3     # project, folder, module, reldoc
    MODULE_ICON = {"Form": "form", "MDIForm": "mdiform", "Module": "module", "Class": "class",
                   "UserControl": "usercontrol", "PropertyPage": "proppage", "UserDocument": "userdoc"}

    def fill_tree(self, fresh=False):
        """the Project Explorer again; the module that was selected stays selected (a Save, a new related
        document ... do not empty the Properties window), the folders the user opened or closed stay so.
        fresh (a project just opened): the folders closed and the project selected, as VB6 (DEV-8/rs1 stage 8)"""
        it = self.tree.currentItem()
        keep = it.data(0, Qt.ItemDataRole.UserRole) if it is not None else None
        keep_proj = it.data(0, self.PROJECT_ROLE) if it is not None and self.pe_kind(it) == "project" else None
        if fresh:
            self.pe_expanded = {}
        else:
            self._pe_remember()
        self.tree.blockSignals(True)
        self.tree.clear()
        self.tree.blockSignals(False)
        p = self.project
        self.refresh_toolbox()
        if p is None:
            self.dock_project.setWindowTitle("Project Explorer")
            self.update_pe_buttons()
            return
        g = self.group
        if g is not None:                       # Project Group - Group1, a branch for each project
            self.dock_project.setWindowTitle(f"Project Group - {g.name}")
            for q in g.projects:
                self._tree_project(q, bold=q is g.startup, fresh=fresh)
        else:
            self.dock_project.setWindowTitle(f"Project - {p.name}")
            self._tree_project(p, bold=True, fresh=fresh)
        self.tree.blockSignals(True)
        self.tree.setAutoScroll(False)          # (selecting a module in a closed folder does not open it)
        if keep is not None:
            self.select_in_tree(keep)
        elif keep_proj is not None:
            self.select_project_in_tree(keep_proj)
        self.tree.setAutoScroll(True)
        self.tree.blockSignals(False)
        self.fill_properties()
        self.update_pe_buttons()
        self.addins_sync()

    def addins_sync(self):
        """the add-ins of the user see what changed (their events: addinhost.sync)"""
        if getattr(self, "addin_hosts", None):
            from . import addinhost as AH
            AH.sync_all(self)

    def _pe_key(self, it):
        """the key of a branch whose open / closed state is kept: (project, folder name or None)"""
        q = it.data(0, self.PROJECT_ROLE)
        kind = self.pe_kind(it)
        if q is None or kind not in ("project", "folder"):
            return None
        return (id(q), it.text(0) if kind == "folder" else None)

    def _pe_remember(self):
        stack = [self.tree.invisibleRootItem()]
        while stack:
            x = stack.pop()
            for i in range(x.childCount()):
                c = x.child(i)
                k = self._pe_key(c)
                if k is not None:
                    self.pe_expanded[k] = c.isExpanded()
                stack.append(c)

    def _pe_expand(self, it, fresh, default=True):
        """open a branch as it was; a new one opens (VB6: a project or folder that appears is open), except in
        a project just opened, where the folders are closed"""
        k = self._pe_key(it)
        if k in self.pe_expanded:
            it.setExpanded(self.pe_expanded[k])
        else:
            it.setExpanded(default and not fresh)

    def _tree_project(self, p, bold, fresh=False):
        """the branch of one project: its name and file, its modules by kind (VB6's folders); Toggle Folders
        off: the modules under the project, by name (DEV-8/me5 picture 12)"""
        file = os.path.basename(p.path) if p.path else p.name     # unsaved: "Project1 (Project1)" (me5-01)
        root = QTreeWidgetItem([f"{p.name} ({file})"])
        root.setData(0, self.PROJECT_ROLE, p)
        root.setData(0, self.KIND_ROLE, "project")
        root.setIcon(0, edicons.icon("project"))
        f = root.font(0)
        f.setBold(bold)
        root.setFont(0, f)
        self.tree.addTopLevelItem(root)
        items = []
        for m in p.modules:
            it = QTreeWidgetItem([self._tree_label(m)])
            it.setData(0, Qt.ItemDataRole.UserRole, m)
            it.setData(0, self.PROJECT_ROLE, p)
            it.setData(0, self.KIND_ROLE, "module")
            it.setIcon(0, edicons.icon(self.MODULE_ICON.get(m.display_kind, "designer")))
            items.append((FOLDER_OF_KIND.get(m.display_kind, m.kind), it))
        docs = []                               # Related Documents: "(s1.RES)" (DEV-8/rs1 picture 12)
        res = (p.vbp.get("ResFile32") or "").strip('"')
        if res:
            docs.append(("res", res))
        docs += [("doc", v.strip().strip('"')) for v in p.vbp.get_all("RelatedDoc")]
        for kind, path in docs:
            it = QTreeWidgetItem([f"({os.path.basename(path.replace(chr(92), '/'))})"])
            it.setData(0, self.RELDOC_ROLE, (kind, path))
            it.setData(0, self.PROJECT_ROLE, p)
            it.setData(0, self.KIND_ROLE, "reldoc")
            it.setIcon(0, edicons.icon("reldoc"))
            items.append(("Related Documents", it))
        if not self.pe_folders:
            for _, it in sorted(items, key=lambda x: x[1].text(0).lower()):
                root.addChild(it)
        else:
            folders = {}
            for fname, it in items:
                if fname not in folders:
                    fo = QTreeWidgetItem([fname])
                    fo.setData(0, self.PROJECT_ROLE, p)
                    fo.setData(0, self.KIND_ROLE, "folder")
                    folders[fname] = fo
                folders[fname].addChild(it)
            order = ["Forms", "Modules", "Class Modules", "User Controls", "Property Pages", "User Documents",
                     "Designers", "Related Documents"]
            for name in sorted(folders, key=lambda n: order.index(n) if n in order else 99):
                fo = folders[name]
                root.addChild(fo)
                fo.sortChildren(0, Qt.SortOrder.AscendingOrder)
                self._pe_expand(fo, fresh)
                self._pe_folder_icon(fo)
        self._pe_expand(root, False)            # the project itself is open, also in a project just opened

    def _pe_folder_icon(self, fo):
        fo.setIcon(0, edicons.icon("folder_open" if fo.isExpanded() else "folder"))

    def pe_kind(self, it):
        """what an item of the Project Explorer is: project, folder, module, reldoc (None: no item)"""
        return it.data(0, self.KIND_ROLE) if it is not None else None

    def select_project_in_tree(self, q):
        for i in range(self.tree.topLevelItemCount()):
            root = self.tree.topLevelItem(i)
            if root.data(0, self.PROJECT_ROLE) is q:
                self.tree.setCurrentItem(root)
                return root
        return None

    def pe_opened(self):
        """a project just opened: its folders closed and the project selected (the designer of the startup
        form opened meanwhile selected the form), Properties "Project1 Project" (DEV-8/rs1 stage 8)"""
        self.pe_expanded = {}
        for i in range(self.tree.topLevelItemCount()):
            root = self.tree.topLevelItem(i)
            for k in range(root.childCount()):
                c = root.child(k)
                if self.pe_kind(c) == "folder":
                    c.setExpanded(False)
                    self._pe_folder_icon(c)
        if self.project is not None:
            self.select_project_in_tree(self.project)
        self.fill_properties()
        self.update_pe_buttons()

    def toggle_folders(self):
        """Toggle Folders: the modules without their folders, by name; back again: the folders closed
        (DEV-8/me5 pictures 12, 13)"""
        self.pe_folders = not self.pe_folders
        self.a_pe_folders.setChecked(self.pe_folders)
        it = self.tree.currentItem()
        if self.pe_kind(it) == "folder":        # a folder that goes away: its project stays selected
            self.tree.blockSignals(True)
            self.tree.setCurrentItem(it.parent())
            self.tree.blockSignals(False)
        self.fill_tree(fresh=self.pe_folders)   # the folders come back closed

    def update_pe_buttons(self):
        """View Code with a module, View Object with a form (grey with a project, a folder, a class or a
        module: DEV-8/me5 pictures 7, 11)"""
        if not hasattr(self, "a_pe_code"):
            return
        it = self.tree.currentItem()
        m = it.data(0, Qt.ItemDataRole.UserRole) if self.pe_kind(it) == "module" else None
        self.a_pe_code.setEnabled(m is not None and m.kind != "Designer")
        self.a_pe_object.setEnabled(m is not None and m.has_designer)

    def tree_project(self, it=None):
        """the project of an item of the Project Explorer"""
        it = it or self.tree.currentItem()
        return it.data(0, self.PROJECT_ROLE) if it is not None else None

    def _tree_current(self, cur, prev):
        """in a group, the project of the chosen item becomes the active one (VB6)"""
        q = self.tree_project(cur)
        if q is not None and q is not self.project:
            self.set_active(q)
        self.fill_properties()
        self.update_pe_buttons()
        self.addins_sync()

    def _tree_expanded(self, it):
        if self.pe_kind(it) == "folder":
            self._pe_folder_icon(it)

    def set_active(self, p):
        """the project the menus work on (Add Form, References, Make ...)"""
        if p is self.project:
            return
        self.project = p
        self.refresh_toolbox()
        self.update_title()
        self.addins_sync()

    def _tree_menu(self, pos):
        """the menus of the Project Explorer, as VB6's (DEV-8/me5 pictures 5-10, 17, 18; rs1 picture 29). The
        item under the mouse becomes the selected one first, as in VB6."""
        it = self.tree.itemAt(pos)
        from PyQt6.QtWidgets import QMenu
        menu = QMenu(self)
        if it is not None:
            self.tree.setCurrentItem(it)
        self.pe_menu(menu, it)
        self._last_tree_menu = menu
        self.shortcut_menu(self.tree_shortcut(it), menu)
        menu.exec(self.tree.viewport().mapToGlobal(pos))

    def tree_shortcut(self, it):
        """the name of the menu of an item of the Project Explorer on the Shortcut Menus toolbar"""
        if it is None:
            return "Project Window (Break)" if self.runner.state == "break" else "Project Window"
        t = it.text(0).lower()
        if it.parent() is None:
            return "Project Window Project"
        if it.childCount():
            if "form" in t or "control" in t or "page" in t or "document" in t or "designer" in t:
                return "Project Window Form Folder"
            if "module" in t or "class" in t:
                return "Project Window Module/Class Folder"
            return "Project Window Related Documents Folder"
        return "Project Window"

    def pe_menu(self, menu, it):
        """the menu of an item of the Project Explorer (and of its empty space: Dockable, Hide)"""
        kind = self.pe_kind(it)
        q = self.tree_project(it)
        if q is not None:
            self.set_active(q)
        if kind == "reldoc":
            doc = it.data(0, self.RELDOC_ROLE)
            if doc[0] == "res":
                self.res_doc_menu(menu, q, os.path.basename(doc[1].replace(chr(92), '/')))
                return menu
            kind = None
        if kind == "module":
            m = it.data(0, Qt.ItemDataRole.UserRole)
            if m.has_designer:
                menu.addAction(edicons.icon("view_object"), "View O&bject", lambda: self.open_designer(m))
            if m.kind != "Designer":
                menu.addAction(edicons.icon("view_code"), "View C&ode", lambda: self.open_code(m))
            menu.addSeparator()
            menu.addAction(edicons.icon("properties"), "Prope&rties",
                           lambda: self._show_dock(self.dock_props, self.props.grid))
            menu.addSeparator()
            self._pe_add_menu(menu)
            menu.addAction(f"&Save {m.name}", self.save_file)
            menu.addAction(f"Save {m.name} &As...", self.save_file_as)
            menu.addAction(f"&Remove {m.name}", self.remove_module)
            menu.addSeparator()
            menu.addAction(edicons.icon("print"), "&Print...", self.print_dialog)
        elif kind == "folder":
            for icon, text in (("view_code", "View C&ode"), ("view_object", "View O&bject")):
                menu.addAction(edicons.icon(icon), text).setEnabled(False)
            menu.addSeparator()
            menu.addAction(f"{q.name} Prop&erties...", self.project_properties)
            menu.addSeparator()
            self._pe_add_menu(menu)
            menu.addAction("Save").setEnabled(False)
            menu.addAction("Save As...").setEnabled(False)
            menu.addAction(edicons.icon("print"), "&Print...", self.print_dialog)
        elif kind == "project":
            # Set as Start Up: on also with one project and for the start-up project itself (me5-09, me5-18)
            menu.addAction("Set as Start &Up", lambda: self.set_startup(q))
            menu.addAction(f"{q.name} Prop&erties...", self.project_properties)
            menu.addSeparator()
            menu.addAction("Sa&ve Project", self.save_project)
            menu.addAction("&Remove Project", lambda: self.remove_project(q))
            menu.addSeparator()
            self._pe_add_menu(menu)
            menu.addSeparator()
            menu.addAction(edicons.icon("print"), "&Print...", self.print_dialog)
        if not menu.isEmpty():
            menu.addSeparator()
        self._pe_dock_items(menu)
        return menu

    def _pe_dock_items(self, menu):
        dockable = self.is_dockable("project")
        a = menu.addAction("Doc&kable", lambda: self.set_dockable("project", not dockable))
        a.setCheckable(True)
        a.setChecked(dockable)
        menu.addAction("&Hide", self.dock_project.hide)

    def _pe_add_menu(self, menu):
        """Add >, as VB6's (DEV-8/me5 picture 10, rs1 picture 30): the kinds LinBASIC Studio does not make are
        grey; Resource File only with the VB 6 Resource Editor loaded"""
        add = menu.addMenu("&Add")
        q = self.project
        mdi = q is not None and q.mdi_form() is not None
        for text, kind, icon in (("&Form", "Form", "form"), ("MD&I Form", "MDIForm", "mdiform"),
                                 ("&Module", "Module", "module"), ("&Class Module", "Class", "class"),
                                 ("&User Control", "UserControl", "usercontrol"),
                                 ("P&roperty Page", "PropertyPage", "proppage"), ("User &Document", None, "userdoc"),
                                 ("WebClass", None, None), ("Data Report", "DataReport", None),
                                 ("DHTML Page", None, None), ("Data Environment", "DataEnvironment", None)):
            slot = self.add_mdi_form if kind == "MDIForm" else (lambda k=kind: self.add_module(k))
            a = add.addAction(text, slot)
            if icon is not None:
                a.setIcon(edicons.icon(icon))
            a.setEnabled(kind is not None and q is not None and not (kind == "MDIForm" and mdi))
        more = add.addMenu("More ActiveX Desi&gners...")
        more.setEnabled(False)
        add.addAction("&Add File...", self.add_file)
        if self.res_loaded:
            add.addAction("Resource File", self.add_new_resource_file)
        return add

    def res_doc_menu(self, menu, q, name):
        """the menu of the resource file in the Project Explorer, as VB6's (DEV-8/rs1 stage 4, stage 8 picture
        29: all of it enabled): Properties | Add >, Save, Save As..., Remove <file> | Print... | Dockable, Hide"""
        menu.addAction(edicons.icon("properties"), "Prope&rties",
                       lambda: self._show_dock(self.dock_props, self.props.grid))
        menu.addSeparator()
        self._pe_add_menu(menu)
        menu.addAction("Save", lambda: self.show_res_editor().save())
        menu.addAction("Save As...", lambda: self.show_res_editor().save_as())
        menu.addAction(f"&Remove {name}", lambda: self.remove_res_file(q))
        menu.addSeparator()
        menu.addAction(edicons.icon("print"), "&Print...", self.print_dialog)
        menu.addSeparator()
        self._pe_dock_items(menu)

    def set_dock_dockable(self, dock, on):
        """Dockable of a tool window: without it the window floats and no side takes it"""
        if on:
            dock.setAllowedAreas(Qt.DockWidgetArea.AllDockWidgetAreas)
        else:
            dock.setAllowedAreas(Qt.DockWidgetArea.NoDockWidgetArea)
            dock.setFloating(True)

    def refresh_toolbox(self):
        """the Toolbox: the controls of the components (.ocx) and the user controls the project can use:
        its own, and the public ones of the control projects of the group"""
        if getattr(self, "toolbox", None) is None:
            return
        from .designer import set_user_controls, geti
        for m, sw in getattr(self, "code_windows", {}).items():   # the code being typed: its Property and Event lines
            if m.kind == "UserControl":
                m.set_code(sw.widget().text())
        p = self.project
        group = getattr(self, "group", None)
        projects = group.projects if group is not None else ([p] if p is not None else [])
        entries = []
        for q in projects:
            for m in q.modules:
                if m.kind != "UserControl" or m.sf is None or m.sf.root is None:
                    continue
                if q is p or (q.ptype == "Control" and m.instancing() != 1):
                    entries.append((q.name, m.name, geti(m.sf.root, "ClientWidth", 4800),
                                    geti(m.sf.root, "ClientHeight", 3600), m))
        set_user_controls(entries)
        keys = (p.components() if p is not None else []) + [f"uc:{e[0]}.{e[1]}" for e in entries]
        self.toolbox.set_components(keys)

    def _tree_label(self, m):
        """"Form1 (Form1.frm)"; a module not saved yet has its name in the brackets: "Form1 (Form1)" (DEV-8/me5)"""
        saved = not m.project.untitled and os.path.exists(m.path)
        file = os.path.basename(m.rel.replace(chr(92), '/')) if saved else m.name
        return f"{m.name} ({file})" + (" *" if m.dirty else "")

    def update_tree_marks(self):
        it = self.tree.invisibleRootItem()
        stack = [it]
        while stack:
            x = stack.pop()
            for i in range(x.childCount()):
                c = x.child(i)
                m = c.data(0, Qt.ItemDataRole.UserRole)
                if m is not None:
                    c.setText(0, self._tree_label(m))
                stack.append(c)
        self.update_title()

    def select_in_tree(self, m):
        stack = [self.tree.invisibleRootItem()]
        while stack:
            x = stack.pop()
            for i in range(x.childCount()):
                c = x.child(i)
                if c.data(0, Qt.ItemDataRole.UserRole) is m:
                    if self.tree.currentItem() is not c:
                        self.tree.setCurrentItem(c)
                    return
                stack.append(c)

    def selected_module(self):
        it = self.tree.currentItem()
        m = it.data(0, Qt.ItemDataRole.UserRole) if it is not None else None
        if m is None:
            cw = self.current_cw()
            m = cw.module if cw else None
        return m

    def _tree_open(self, it, col):
        doc = it.data(0, self.RELDOC_ROLE)
        if doc is not None and doc[0] == "res":     # the resource file opens in the Resource Editor
            q = it.data(0, self.PROJECT_ROLE)
            if q is not None:
                self.set_active(q)
            self.show_res_editor()
            return
        m = it.data(0, Qt.ItemDataRole.UserRole)
        if m is not None:
            if m.has_designer:
                self.open_designer(m)       # VB6: a form opens in its designer
            else:
                self.open_code(m)

    def fill_properties(self):
        """the Properties window: the selected module, with its designer when it is open; the project ("Project1
        Project", (Name)); nothing for a folder, or for a form that has no window open (DEV-8/me5 pictures 8,
        11, 14)"""
        self.refresh_toolbox()                  # the properties of the user controls follow their code
        it = self.tree.currentItem()
        kind = self.pe_kind(it)
        if kind == "project":
            self.props.show_project(self.tree_project(it))
            return
        if kind == "folder":
            self.props.show_object(None, None, None)
            return
        m = self.selected_module()
        sw = self.design_windows.get(m) if m is not None else None
        d = sw.widget() if sw is not None else None
        if m is not None and m.kind == "Designer":
            self.props.show_object(m, None, None)   # the designer has the properties of its objects
        elif d is not None:
            self.props.show_object(m, d, d.primary.ctl if d.primary else d.root)
        elif m is not None and m.sf is not None and m.sf.root is not None:
            if m in self.code_windows:
                self.props.show_object(m, None, m.sf.root)
            else:
                self.props.show_object(None, None, None)
        else:
            self.props.show_object(m, None, None)

    def rename_project(self, q, new):
        """(Name) of a project: Name= of the .vbp"""
        import re
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,39}", new or ""):
            return "Invalid property value"
        others = self.group.projects if self.group is not None else [q]
        if any(x is not q and x.name.lower() == new.lower() for x in others) or \
                any(m.name.lower() == new.lower() for m in q.modules):
            return f"Name '{new}' conflicts with existing module, project, or object library"
        if new == q.name:
            return None
        q.vbp.set("Name", new, quote=True)
        q.dirty = True
        self.fill_tree()
        self.update_title()
        return None

    def rename_module(self, m, new):
        """(Name) of a form, module or class: VB_Name, the Begin line of a form and the .vbp"""
        import re
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,39}", new or ""):
            return "Invalid property value"
        if any(x is not m and x.name.lower() == new.lower() for x in self.project.modules):
            return f"Name '{new}' conflicts with existing module, project, or object library"
        if m.kind == "Form":        # nor the name of a control in a form (the user's decision, Phase 9 / DEV-9 nm)
            for x in self.project.modules:
                r = x.sf.root if x.kind == "Form" and x.sf is not None else None
                if r is not None and any(c is not r and c.name.lower() == new.lower() for c in r.walk()):
                    return f"Ambiguous name detected: {new}"
        old = m.name
        if new == old:
            return None
        sf = m.sf
        sf.attributes = [(n, f'"{new}"', f'Attribute VB_Name = "{new}"') if n == "VB_Name" else (n, v, raw)
                         for n, v, raw in sf.attributes]
        if sf.root is not None:
            sf.root.name = new
            sf.root.raw_begin = None
        m.name = new
        v = self.project.vbp
        if (v.get("Startup") or "").strip('"').lower() == old.lower():
            v.set("Startup", new, quote=True)
        for e in v.entries:
            k, val = e
            if val is not None and k.strip() == m.kind and ";" in val and val.split(";")[0].strip().lower() == old.lower():
                e[1] = f"{new}; {val.split(';', 1)[1].strip()}"
        # a file never written takes the new name (VB6 writes Module1 renamed ModEv as ModEv.bas, oracle evlog)
        rel = m.rel.replace("\\", "/")
        if (getattr(m, "never_saved", False) or not os.path.exists(m.path)) and \
                os.path.splitext(os.path.basename(rel))[0].lower() == old.lower():
            d = os.path.dirname(rel)
            nrel = (d + "/" if d else "") + new + os.path.splitext(rel)[1]
            for e in v.entries:
                if e[1] is not None and e[0].strip() == m.kind and e[1].split(";")[-1].strip() == m.rel:
                    e[1] = e[1][:len(e[1]) - len(m.rel)] + nrel
            m.rel = nrel
        m.dirty = True
        self.project.dirty = True
        sw = self.design_windows.get(m)
        if sw is not None:
            if m.kind == "Designer":
                sw.widget().model.name = new
            sw.setWindowTitle(self.designer_title(m))
        cw = self.code_windows.get(m)
        if cw is not None:
            self.update_tab_title(cw.widget())
        self.fill_tree()
        self.select_in_tree(m)
        return None

    # ================================================================ project files
    def codepage(self):
        from .vbmodel import default_codepage
        return self.S["codepage"] or default_codepage()

    def confirm_discard(self):
        if self.project is None:
            return True
        self._sync_editors()
        g = self.group
        if not (g.any_dirty() if g is not None else self.project.any_dirty()):
            return True
        names = "\n".join([g.name] + [q.name for q in g.projects if q.any_dirty()]) if g is not None else self.project.name
        r = QMessageBox.question(self, APP_NAME, f"Save changes to the following files?\n\n{names}",
                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No |
                                 QMessageBox.StandardButton.Cancel)
        if r == QMessageBox.StandardButton.Cancel:
            return False
        if r == QMessageBox.StandardButton.Yes:
            return self.save_group() if g is not None else self.save_project()
        return True

    def close_all_code(self):
        for sw in list(self.code_windows.values()) + list(self.design_windows.values()):
            sw.close()
        self.code_windows.clear()
        self.design_windows.clear()

    def close_project(self):
        """File > Close Project: the IDE without a project (New or Open makes one); in a group, the
        active project leaves the group and the others stay open"""
        if self.project is None:
            return
        if self.group is not None and len(self.group.projects) > 1:
            if self.runner.state != "design":
                self.runner.stop()
            self.remove_project(self.project)
            return
        self.close_group()

    def close_group(self, ask=True):
        """File > Close Project Group: every project of the group (or the one project) closes; the IDE
        asks first to save what changed"""
        if self.project is None:
            return
        if self.runner.state != "design":
            self.runner.stop()
        if ask and (not self.confirm_discard() or not self.res_confirm()):
            return
        self.close_all_code()
        self.project = None
        self.group = None
        self.res_rebind()
        self._intellisense = None
        self.watches = []
        self.trap_session = None
        self.fill_tree()
        self.props.show_object(None, None, None)
        self.immediate.clear()
        self.locals_win.clear()
        self.watch_win.fill([])
        self.output.clear()
        self.update_title()

    def new_project(self, ptype=None):
        if ptype is None:                       # VB6 asks for the type of the project
            from .dialogs import NewProjectDialog
            dlg = NewProjectDialog(self)
            if not dlg.exec():
                return
            ptype = dlg.ptype()
        if ptype == "AppWizard":                # VB Application Wizard (phase 9.2): the wizard makes the project
            self.application_wizard()
            return
        if not self.confirm_discard() or not self.res_confirm():
            return
        self.close_all_code()
        self.group = None
        if ptype == "Addin":                    # an add-in project (phase 12.5)
            from .vbmodel import new_addin_project
            self.project = new_addin_project(self.codepage())
        else:
            self.project = VBProject.new(self.codepage(), ptype)
        self.res_rebind()
        self.fill_tree()
        self.update_title()
        self.immediate.clear()
        if self.project.modules:
            self.open_designer(self.project.modules[0])

    def open_project_dialog(self):
        if not self.confirm_discard():
            return
        start = self.S["last_dir"] or os.path.expanduser("~")
        path, _ = QFileDialog.getOpenFileName(self, "Open Project", start,
                                              file_filter("Project Files", ["vbp", "mak", "vbg"]) + ";;All Files (*)")
        if path:
            self.open_project(path)

    def open_project(self, path):
        group = None
        try:
            if path.lower().endswith(".vbg"):   # a project group: its startup project is the active one
                group = VBGroup.open(path, self.codepage())
                if not group.projects:
                    raise ValueError("the group has no projects")
                p = group.startup
            else:
                p = VBProject.open(path, self.codepage())
        except Exception as e:                  # noqa: BLE001 - shown to the user
            self.message(f"Can't open {path}: {e}", "error", show=True)
            return False
        if not self.res_confirm():
            return False
        self.close_all_code()
        self.group = group
        self.project = p
        self.res_rebind()
        self.S["last_dir"] = os.path.dirname(path)
        self.add_recent(path)
        self.fill_tree(fresh=True)
        self.update_title()
        self.immediate.clear()
        # VB6 opens the startup object
        st = (p.startup or "").strip('"')
        m = p.module(st) if st and st.lower() != "sub main" else None
        if m is None:
            m = next((x for x in p.modules if x.kind == "Module"), p.modules[0] if p.modules else None)
        if m is not None:
            if m.has_designer:
                self.open_designer(m)
            else:
                self.open_code(m)
        # the Project Explorer: folders closed, the project selected (also after the window activation that
        # the event loop delivers later)
        self.pe_opened()
        QTimer.singleShot(0, self.pe_opened)
        return True

    def save_project(self, ask_new=True):
        p = self.project
        if p is None:
            return False
        if p.untitled:
            return self.save_project_as()
        self._sync_editors()
        for m in p.modules:                     # VB6: Save File As for each file never saved (oracle 12.5/evlog-vb6-gui)
            if ask_new and getattr(m, "never_saved", False) and not self._ask_new_file(m):
                return False
        try:
            p.save()
        except OSError as e:
            self.message(f"Can't save the project: {e}", "error", show=True)
            return False
        for sw in self.code_windows.values():
            sw.widget().mark_saved()
        self.update_tree_marks()
        for sw in self.code_windows.values():
            self.update_tab_title(sw.widget())
        self.statusBar().showMessage("Saved " + p.path, 4000)
        return True

    def save_project_as(self):
        if self.project is None:
            return
        p = self.project
        start = os.path.dirname(p.path) if p.path else (self.S["last_dir"] or os.path.expanduser("~"))
        # VB6 asks for each new file first, then the project
        for m in p.modules:
            if p.untitled or not os.path.exists(m.path):
                ext = os.path.splitext(m.rel)[1]
                path, _ = QFileDialog.getSaveFileName(self, f"Save File As ({m.name})", os.path.join(start, m.name + ext),
                                                      file_filter(f"{m.kind} Files", [ext[1:]]))
                if not path:
                    return False
                start = os.path.dirname(path)
                m.rel = os.path.basename(path) if p.untitled else os.path.relpath(path, os.path.dirname(p.path))
                m.dirty = True
        path, _ = QFileDialog.getSaveFileName(self, "Save Project As", os.path.join(start, p.name + ".vbp"),
                                              "Project Files (*.vbp)")
        if not path:
            return False
        if not path.lower().endswith(".vbp"):
            path += ".vbp"
        # member lines of the .vbp follow the new file names
        p.vbp.entries = [[k, v] for k, v in p.vbp.entries
                         if k.strip() not in ("Form", "Module", "Class", "UserControl", "PropertyPage", "UserDocument")]
        old_untitled = p.untitled
        p.path = os.path.abspath(path)
        p.vbp.path = p.path
        for m in p.modules:
            if old_untitled:
                m.rel = os.path.relpath(os.path.join(start, os.path.basename(m.rel)), os.path.dirname(p.path))
            p._vbp_add(m)
        self.add_recent(p.path)
        ok = self.save_project(ask_new=False)   # (the files were asked for above)
        self.fill_tree()
        self.update_title()
        return ok

    def _ask_new_file(self, m):
        """Save File As for a file never saved in a saved project: its name and folder (the .vbp follows)"""
        p = m.project
        ext = os.path.splitext(m.rel)[1]
        if getattr(self, "addin_hosts", None):  # (the add-ins: DoGetNewFileName before the dialog, as VB6)
            from . import addinhost as AH
            from .vbide import file_type
            AH.file_event(self, "DoGetNewFileName", p, file_type(m), m.name, m.name, False)
            m._lb_getnew_done = True
        path, _ = QFileDialog.getSaveFileName(self, "Save File As", os.path.join(os.path.dirname(p.path),
                                                                                 os.path.basename(m.rel)),
                                              file_filter(f"{m.kind} Files", [ext[1:]]))
        if not path:
            m._lb_getnew_done = False
            return False
        old = m.rel
        m.rel = os.path.relpath(path, os.path.dirname(p.path))
        for e in p.vbp.entries:
            if e[1] is not None and e[0].strip() == m.kind and e[1].split(";")[-1].strip() == old:
                e[1] = e[1][:len(e[1]) - len(old)] + m.rel
        p.dirty = True
        m.dirty = True
        return True

    def save_file(self):
        m = self.selected_module()
        if m is None or self.project is None:
            return
        if self.project.untitled:
            self.save_project_as()
            return
        self._sync_editors()
        if getattr(m, "never_saved", False):
            if not self._ask_new_file(m):
                return
            m.save()                            # (the .vbp has its new name when the project is saved)
            self.fill_tree()
        else:
            m.save()
        sw = self.code_windows.get(m)
        if sw:
            sw.widget().mark_saved()
            self.update_tab_title(sw.widget())
        self.update_tree_marks()

    def save_file_as(self):
        m = self.selected_module()
        if m is None or self.project is None or self.project.untitled:
            self.save_project_as()
            return
        ext = os.path.splitext(m.rel)[1]
        path, _ = QFileDialog.getSaveFileName(self, "Save File As", m.path, file_filter(f"{m.kind} Files", [ext[1:]]))
        if not path:
            return
        self._sync_editors()
        p = self.project
        old = m.rel
        m.rel = os.path.relpath(path, os.path.dirname(p.path))
        for e in p.vbp.entries:
            if e[1] is not None and e[0].strip() == m.kind and old in e[1]:
                e[1] = e[1].replace(old, m.rel)
        m.save()
        p.vbp.save()
        self.fill_tree()

    def add_module(self, kind):
        from . import addinhost as AH
        with AH.hold(self):                     # (the add-ins: ItemAdded and ItemSelected in VB6's order)
            m = self.project.add_module(kind)
            self.fill_tree()
            if m.has_designer:                  # VB6: a new form or user control opens in its designer
                self.open_designer(m)
            else:
                self.open_code(m)
            self.select_in_tree(m)
            self.update_title()
        self.addins_sync()

    def add_file(self):
        start = os.path.dirname(self.project.path) if self.project.path else os.path.expanduser("~")
        path, _ = QFileDialog.getOpenFileName(self, "Add File", start,
                                              file_filter("VB Files", ["frm", "bas", "cls", "ctl", "pag", "dsr"]) + ";;" +
                                              file_filter("Resource Files", ["res"]) + ";;All Files (*)")
        if not path:
            return
        self.add_file_path(path)

    def add_file_path(self, path, events=True):
        """Project > Add File... (and the Existing tab of Add MDI Form): a module of a file into the project
        (the add-ins: ItemAdded, AfterAddFile; events False: VBComponents.AddFile gives them itself)"""
        if path.lower().endswith(".res"):       # the resource file of the project (ResFile32=)
            self.attach_res_file(path)
            return
        from .vbmodel import Module, KIND_OF_EXT, load_source
        kind = KIND_OF_EXT.get(os.path.splitext(path)[1].lower())
        if kind not in ("Form", "Module", "Class", "UserControl", "PropertyPage", "Designer"):
            return
        sf = load_source(path, encoding=self.project.encoding)
        base = os.path.dirname(self.project.path) if self.project.path else os.path.dirname(path)
        m = Module(self.project, kind, sf.name or os.path.splitext(os.path.basename(path))[0], os.path.relpath(path, base), sf)
        self.project.modules.append(m)
        self.project._vbp_add(m)
        self.project.dirty = True
        self.fill_tree()
        if events and getattr(self, "addin_hosts", None):
            from . import addinhost as AH
            from .vbide import file_type
            AH.file_event(self, "AfterAddFile", self.project, file_type(m), os.path.abspath(path))

    def remove_module(self):
        m = self.selected_module()
        if m is None:
            return
        self._sync_editors()
        if m.dirty:                             # VB6 (oracle 12.5/evlog-vb6-gui): Save changes to the following files?
            r = QMessageBox.question(self, APP_NAME, f"Save changes to the following files?\n\n{m.name}",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No |
                                     QMessageBox.StandardButton.Cancel)
            if r == QMessageBox.StandardButton.Cancel:
                return
            if r == QMessageBox.StandardButton.Yes:
                if os.path.exists(m.path) or not self.project.untitled:
                    m.save()
                else:
                    ext = os.path.splitext(m.rel)[1]
                    path, _ = QFileDialog.getSaveFileName(self, f"Save File As ({m.name})", m.name + ext,
                                                          file_filter(f"{m.kind} Files", [ext[1:]]))
                    if not path:
                        return
                    m.save(path)
        from . import addinhost as AH
        with AH.hold(self):
            for sw in (self.code_windows.get(m), self.design_windows.get(m)):
                if sw:
                    sw.close()
            # the Project Explorer follows the window in front (VB6: ItemSelected of it before ItemRemoved; none
            # when no window is left - oracle 12.5/evlog-vb6-gui)
            subs = [x for x in self.mdi.subWindowList(QMdiArea.WindowOrder.ActivationHistoryOrder) if x.isVisible()]
            front = self.mdi.activeSubWindow() or (subs[-1] if subs else None)
            fm = getattr(front.widget(), "module", None) if front is not None else None
            if fm is not None and fm is not m:
                self.select_in_tree(fm)
            else:
                self.tree.setCurrentItem(None)
        self.addins_sync()
        p = self.project
        p.remove_module(m)
        self.fill_tree()
        self.update_title()
        if getattr(self, "addin_hosts", None) and os.path.exists(m.path):
            from . import addinhost as AH
            from .vbide import file_type
            AH.file_event(self, "AfterRemoveFile", p, file_type(m), m.path)

    # ================================================================ the VB 6 Resource Editor (phase 8.6)
    def apply_addins(self):
        """Add-Ins: the VB 6 Resource Editor loaded gives Tools > Resource Editor, Project > Add New Resource
        File and the button of the Standard toolbar (DEV-8/rs1 pictures 3, 4)"""
        on = self.res_loaded
        for a in (self.sep_res, self.a_res_editor, self.a_new_res):
            a.setVisible(on)
        for a in (self.sep_pdw, self.a_pdw):     # the Package and Deployment Wizard (phase 9.5)
            a.setVisible(bool(self.pdw_loaded))
        for a in (self.sep_appwiz, self.a_appwiz, self.a_tbwiz):   # the VB 6 Application Wizard (phase 9.2)
            a.setVisible(bool(self.appwiz_loaded))
        for a in (self.sep_dfwiz, self.a_dfwiz):                    # the VB 6 Data Form Wizard (phase 9.2b)
            a.setVisible(bool(self.dfwiz_loaded))
        for a in (self.sep_cbuilder, self.a_cbuilder):              # the VB 6 Class Builder Utility (9.2c)
            a.setVisible(bool(self.cbuilder_loaded))
        for a in (self.sep_axwiz, self.a_axwiz):                    # the VB 6 ActiveX Ctrl Interface Wizard (9.2d)
            a.setVisible(bool(self.axwiz_loaded))
        for a in (self.sep_ppwiz, self.a_ppwiz):                    # the VB 6 Property Page Wizard (9.2e)
            a.setVisible(bool(self.ppwiz_loaded))
        for a in (self.sep_dowiz, self.a_dowiz):                    # the VB 6 Data Object Wizard (9.2f)
            a.setVisible(bool(self.dowiz_loaded))
        for a in (self.sep_apiview, self.a_apiview):                # the VB 6 API Viewer (9.4)
            a.setVisible(bool(self.apiviewer_loaded))
        v = getattr(self, "_apiviewer", None)
        if not self.apiviewer_loaded and v is not None:
            v.close()                           # (VB6: unloading the add-in closes its window)
        if not on and self.res_sub is not None:
            self.res_sub.hide()
            if self.res_editor.strings_win is not None:
                self.res_editor.close_strings()

    def application_wizard(self):
        """File > New Project > VB Application Wizard, Add-Ins > Application Wizard... (phase 9.2, appwizard.py)"""
        from .appwizard import open_wizard
        return open_wizard(self)

    def data_form_wizard(self):
        """Add-Ins > Data Form Wizard... (phase 9.2b, dfwizard.py)"""
        from .dfwizard import open_wizard
        return open_wizard(self)

    def class_builder(self):
        """Add-Ins > Class Builder Utility..., Project > Add Class Module > VB Class Builder (phase 9.2c)"""
        from .classbuilder import open_builder
        return open_builder(self)

    def activex_wizard(self):
        """Add-Ins > ActiveX Control Interface Wizard... (phase 9.2d, axwizard.py)"""
        from .axwizard import open_wizard
        return open_wizard(self)

    def property_page_wizard(self):
        """Add-Ins > Property Page Wizard... (phase 9.2e, ppwizard.py)"""
        from .ppwizard import open_wizard
        return open_wizard(self)

    def data_object_wizard(self):
        """Add-Ins > VB 6 Data Object Wizard (phase 9.2f, dowizard.py)"""
        from .dowizard import open_wizard
        return open_wizard(self)

    def visual_data_manager(self):
        """Add-Ins > Visual Data Manager... (phase 12.4, visdata/): the databases of the programs through the DAO of
        the runtime"""
        from .visdata import open_visdata
        return open_visdata(self)

    def api_viewer(self):
        """Add-Ins > API Viewer (phase 9.4, apiviewer.py)"""
        from .apiviewer import open_viewer
        return open_viewer(self)

    def add_class_module(self):
        """Project > Add Class Module: the dialog of VB6 (Class Module, VB Class Builder, Complex Data Consumer,
        Data Source; Existing) unless "Don't show this dialog in the future" (DEV-9/cb1-01, cb4)"""
        p = self.project
        if p is None:
            return None
        if self.S["add_class_nodialog"]:
            return self.add_module("Class")
        from .classbuilder import AddClassModuleDialog, add_template
        start = os.path.dirname(p.path) if p.path else (self.S["last_dir"] or os.path.expanduser("~"))
        dlg = AddClassModuleDialog(self, start)
        self._last_add_class = dlg
        if not dlg.exec():
            return None
        if dlg.dont_show.isChecked():
            self.S["add_class_nodialog"] = True
            self.S.save()
        if dlg.path:
            return self.add_file_path(dlg.path)
        if dlg.kind == "builder":       # the Class Builder, loaded or not as an add-in (DEV-9/cb4-01)
            return self.class_builder()
        if dlg.kind in ("source", "consumer"):
            m = add_template(p, dlg.kind)
            self.fill_tree()
            self.open_code(m)
            self.update_title()
            return m
        return self.add_module("Class")

    def toolbar_wizard(self):
        """Add-Ins > Toolbar Wizard... (phase 9.2a-4)"""
        from .tbwizard import open_wizard
        return open_wizard(self)

    def package_wizard(self):
        """Add-Ins > Package and Deployment Wizard... (phase 9.5, pdwizard.py)"""
        from .pdwizard import open_wizard
        return open_wizard(self)

    def pdw_ask(self, text, buttons):
        """a question of the Package and Deployment Wizard before it opens (the tests answer it)"""
        from .pdwizard import TITLE
        B = QMessageBox.StandardButton
        r = QMessageBox.question(self, TITLE, text, buttons)
        return {B.Yes: "yes", B.No: "no", B.Cancel: "cancel"}.get(r, "cancel")

    def addin_manager(self):
        from .resedit import AddInManagerDialog
        dlg = AddInManagerDialog(self)
        if dlg.exec():
            self.apply_addins()

    # ------------------------------------------------------------ the add-ins of the user (phase 12.5)
    def load_startup_addins(self, cmdline=False):
        """Load on Startup (and Command Line, when the IDE starts with a project on its command line): OnConnection
        with ext_cm_Startup / ext_cm_CommandLine, then OnStartupComplete"""
        from . import addinhost as AH
        done = []
        for a in AH.user_addins(self):
            if a.get("startup") or (cmdline and a.get("cmdline")):
                h = AH.load_addin(self, a["progid"], AH.CM_COMMANDLINE if cmdline and not a.get("startup")
                                  else AH.CM_STARTUP)
                if h is not None:
                    done.append(h)
        for h in done:
            try:
                h.call("startup")
            except AH.AddInError as e:
                self.addin_error(h.info, f"{e.description} ({e.number})")

    def unload_user_addins(self, shutdown=False):
        from . import addinhost as AH
        hosts = list(self.addin_hosts.values())
        if shutdown:
            self.addins_begin_shutdown()
        for h in hosts:
            AH.unload_addin(self, h.info["progid"], AH.DM_HOSTSHUTDOWN if shutdown else AH.DM_USERCLOSED)

    def addins_begin_shutdown(self):
        """OnBeginShutdown of each add-in, once"""
        from . import addinhost as AH
        for h in list(getattr(self, "addin_hosts", {}).values()):
            if getattr(h, "shut", False) or not h.running():
                continue
            h.shut = True
            try:
                h.call("shutdown")
            except AH.AddInError as er:
                self.addin_error(h.info, f"{er.description} ({er.number})")

    def addin_error(self, info, text):
        """an add-in that could not be loaded, or an error of its event procedures"""
        msg = f"'{info.get('name', info.get('progid', ''))}' could not be loaded or used: {text}"
        self.addin_errors.append(msg)
        self.message(msg, "error", show=True)

    def resource_editor(self):
        """the Resource Editor (one for the IDE), with the resource file of the active project"""
        if self.res_editor is None:
            from .resedit import ResourceEditor
            ed = ResourceEditor(self)
            ed.bind(self.project)
            self.res_editor = ed
            self.res_sub = self.add_res_window(ed, ed.title(), keep=True, size=(440, 340))
            ed.title_changed.connect(self.res_sub.setWindowTitle)
        return self.res_editor

    def show_res_editor(self):
        ed = self.resource_editor()
        if ed.project is not self.project and not ed.modified:
            ed.bind(self.project)
        self.show_res_window(self.res_sub)
        return ed

    def add_res_window(self, w, title, keep=False, size=(480, 300)):
        """the size is given before show, as for the code and form windows: a resize after it takes a tab of
        the tabbed view out of its place (a window without a frame over the others)"""
        sw = ResSubWindow(self, keep)
        sw.setWidget(w)
        sw.setWindowTitle(title)
        self.mdi.addSubWindow(sw)
        if not keep:
            sw.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        sw.resize(*size)
        sw.show()
        self.mdi.setActiveSubWindow(sw)
        return sw

    def show_res_window(self, sw):
        if self.mdi.viewMode() == QMdiArea.ViewMode.TabbedView:
            sw.showMaximized()                  # a hidden tab comes back in its place
        else:
            sw.show()
        self.mdi.setActiveSubWindow(sw)

    def remove_res_window(self, sw):
        sw.keep = False
        self.mdi.removeSubWindow(sw)
        sw.deleteLater()

    def res_confirm(self):
        """a project that closes: "The resource file in <project> has changed ..." (Yes / No)"""
        ed = self.res_editor
        return ed.ask_save() if ed is not None else True

    def res_rebind(self):
        if self.res_editor is not None:
            self.res_editor.bind(self.project)

    def attach_res_file(self, path):
        from .resedit import set_res_file
        p = self.project
        rel = os.path.relpath(path, os.path.dirname(p.path)) if p.path else path
        if not rel.startswith(".."):
            rel = rel.replace("/", "\\")
        set_res_file(p, rel)
        if self.res_editor is not None and self.res_editor.project is p and not self.res_editor.modified:
            self.res_editor.bind(p)
        self.fill_tree()
        self.update_title()

    def add_new_resource_file(self, path=None):
        """Project > Add New Resource File (DEV-8/rs1 stage 6): "Open A Resource File"; a name that is not
        there: "<file> / This file doesn't exist. / Create the file?" (Yes / No) and an empty file (32 bytes);
        the file takes the place of the project's resource file without a question"""
        p = self.project
        if p is None:
            return None
        if path is None:
            start = os.path.dirname(p.path) if p.path else (self.S["last_dir"] or os.path.expanduser("~"))
            dlg = QFileDialog(self, "Open A Resource File", start, "Resource File (*.RES *.res)")
            dlg.setFileMode(QFileDialog.FileMode.AnyFile)
            dlg.setAcceptMode(QFileDialog.AcceptMode.AcceptOpen)
            # the dialog of Qt: the one of the system (GTK) has no "File name" box when it opens a file, and a
            # new name has to be written there (VB6: File name "*.RES")
            dlg.setOption(QFileDialog.Option.DontUseNativeDialog, True)
            dlg.selectFile("*.RES")
            if not dlg.exec() or not dlg.selectedFiles():
                return None
            path = dlg.selectedFiles()[0]
        base, ext = os.path.splitext(path)
        if ext.lower() in (".res", ""):
            path = base + ".RES"
        if not os.path.exists(path):
            from .resedit import res_create_msg
            r = QMessageBox.warning(self, "Open A Resource File", res_create_msg(path),
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if r != QMessageBox.StandardButton.Yes:
                return None
            from .resmodel import ResFile
            ResFile().save(path)
        if not self.res_confirm():
            return None
        self.attach_res_file(path)
        ed = self.resource_editor()
        ed.bind(p)
        self.show_res_window(self.res_sub)
        return path

    def remove_res_file(self, q):
        """the right button of the resource file in the Project Explorer > Remove <file>"""
        from .resedit import set_res_file
        ed = self.res_editor
        if ed is not None and ed.project is q:
            ed.modified = False
        set_res_file(q, None)
        if ed is not None and ed.project is q:
            ed.bind(q)
        self.fill_tree()
        self.update_title()

    # ================================================================ project groups
    def add_project(self, ptype=None, path=None):
        """File > Add Project: a new project of a type, or an existing .vbp; the project and the one
        open make a group (VB6)"""
        if self.project is None:
            return None
        if ptype is None and path is None:
            from .dialogs import NewProjectDialog
            dlg = NewProjectDialog(self, "Add Project", existing=True)
            if not dlg.exec():
                return None
            path = dlg.existing_path
            ptype = None if path else dlg.ptype()
        if self.group is None:
            g = VBGroup()
            g.add(self.project)
            g.startup = self.project
            self.group = g
        g = self.group
        try:
            if path:
                if any(q.path and os.path.samefile(q.path, path) for q in g.projects):
                    self.message(f"{path} is in the group already", "error", show=True)
                    return None
                q = VBProject.open(path, self.codepage())
            else:
                names = {x.name.lower() for x in g.projects}
                i = 1
                while f"project{i}" in names:
                    i += 1
                if ptype == "Addin":            # (phase 12.5) MyAddIn, as New Project makes it
                    from .vbmodel import new_addin_project
                    q = new_addin_project(self.codepage())
                    k = 1
                    while q.name.lower() in names:
                        k += 1
                        q.vbp.set("Name", f"MyAddIn{k}", quote=True)
                else:
                    q = VBProject.new(self.codepage(), ptype, f"Project{i}")
        except Exception as e:                  # noqa: BLE001 - shown to the user
            self.message(f"Can't add the project: {e}", "error", show=True)
            return None
        g.add(q)
        g.dirty = True
        self.project = q
        self.fill_tree()
        self.update_title()
        first = next((m for m in q.modules if m.has_designer), q.modules[0] if q.modules else None)
        if first is not None:
            self.open_designer(first) if first.has_designer else self.open_code(first)
        return q

    def remove_project(self, q, ask=True):
        if q is None:
            return
        if self.group is None or len(self.group.projects) < 2:
            if ask:
                self.close_project()
            else:
                self.close_group(ask=False)
            return
        if ask and q.any_dirty():
            r = QMessageBox.question(self, APP_NAME, f"Save changes to {q.name}?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No |
                                     QMessageBox.StandardButton.Cancel)
            if r == QMessageBox.StandardButton.Cancel:
                return
            if r == QMessageBox.StandardButton.Yes:
                self.set_active(q)
                if not self.save_project():
                    return
        for m in list(q.modules):
            for sw in (self.code_windows.get(m), self.design_windows.get(m)):
                if sw:
                    sw.close()
        self.group.remove(q)
        self.project = self.group.startup or self.group.projects[0]
        self.fill_tree()
        self.update_title()

    def set_startup(self, q):
        """Set as Start Up: the project F5 runs"""
        if self.group is None:
            return
        self.group.startup = q
        self.group.dirty = True
        self.fill_tree()
        self.update_title()

    def save_group(self, ask=False):
        """Save Project Group (As): every project, then the .vbg"""
        g = self.group
        if g is None:
            return self.save_project_as() if ask else self.save_project()
        active = self.project
        for q in g.projects:
            if q.any_dirty() or q.untitled:
                self.project = q
                if not (self.save_project_as() if q.untitled else self.save_project()):
                    self.project = active
                    return False
        self.project = active
        if g.untitled or ask:
            start = os.path.dirname(g.startup.path) if g.startup is not None and g.startup.path else os.path.expanduser("~")
            path, _ = QFileDialog.getSaveFileName(self, "Save Project Group As", os.path.join(start, g.name + ".vbg"),
                                                  "Project Group Files (*.vbg)")
            if not path:
                return False
            if not path.lower().endswith(".vbg"):
                path += ".vbg"
            g.path = os.path.abspath(path)
        g.save()
        self.add_recent(g.path)
        self.fill_tree()
        self.update_title()
        self.statusBar().showMessage("Saved " + g.path, 4000)
        return True

    def references_dialog(self):
        from .dialogs import ReferencesDialog
        if self.project is None:
            return
        if self.project.untitled:
            QMessageBox.information(self, APP_NAME, "Save the project first: a reference names the .vbp of a "
                                    "project by its place relative to this one.")
            return
        if ReferencesDialog(self, self.project, self.group).exec():
            self.update_title()

    def components_dialog(self):
        from .dialogs import ComponentsDialog
        if self.project is None:
            return
        if ComponentsDialog(self, self.project).exec():
            self.toolbox.adding = True          # the new controls go into the open tab (DEV-8/me6-19)
            try:
                self.refresh_toolbox()
            finally:
                self.toolbox.adding = False
            self.update_title()

    def project_properties(self):
        """Project > <Project> Properties (VB6's five pages, phase 8.5)"""
        if self.project is None:
            return
        d = self.project_properties_dialog()
        if d.exec():
            self.fill_tree()
            self.refresh_toolbox()
            self.update_title()

    def project_properties_dialog(self):
        return ProjectPropertiesDialog(self, self.project)

    # ================================================================ recent projects
    def add_recent(self, path):
        lst = [p for p in self.S["recent_projects"] if p != path]
        lst.insert(0, path)
        self.S["recent_projects"] = lst[:10]
        self.S.save()
        self.fill_recent()

    def fill_recent(self):
        self.recent_menu.clear()
        for i, p in enumerate(self.S["recent_projects"] or []):
            a = self.recent_menu.addAction(f"&{i + 1} {p}")
            a.triggered.connect(lambda checked=False, pp=p: self.confirm_discard() and self.open_project(pp))
        self.recent_menu.setEnabled(bool(self.S["recent_projects"]))

    # ================================================================ find
    def find(self):
        if self.find_dlg is None:
            self.find_dlg = FindDialog(self)
        self.find_dlg.show_find(replace=False)

    def replace(self):
        if self.find_dlg is None:
            self.find_dlg = FindDialog(self)
        self.find_dlg.show_find(replace=True)

    def find_next(self):
        if self.find_dlg is None or not self.find_dlg.what.currentText():
            self.find()
            return
        self.find_dlg.find_next()

    # ================================================================ run
    def run_start(self, first="go"):
        """F5: run with the debugger (Continue in break mode)"""
        if self.project is None:
            return
        if self.runner.state == "break":
            self.runner.resume("go")
            return
        if self.runner.state != "design":
            return
        if self.group is not None:              # a group: its startup project runs, with the others as components
            st = self.group.startup
            if st is None or st.ptype in ("OleDll", "Control"):
                QMessageBox.information(self, APP_NAME, "The startup project of the group is an ActiveX DLL or "
                                        "Control: make a Standard EXE project the startup project "
                                        "(Set as Start Up in the Project Explorer).")
                return
            if any(q.untitled for q in self.group.projects):
                QMessageBox.information(self, APP_NAME, "Save the project group first: the projects of a group "
                                        "find each other by their files.")
                if not self.save_group():
                    return
            self.set_active(st)
        if self.project.ptype in ("OleDll", "Control") and getattr(self.project, "group", None) is None:
            what = "An ActiveX DLL" if self.project.ptype == "OleDll" else "An ActiveX Control"
            QMessageBox.information(self, APP_NAME, f"{what} project does not run on its own: add a Standard EXE "
                                    "project that uses it (File > Add Project) and make that the startup project.")
            return
        self._sync_editors()
        self.immediate.clear()
        self.output.clear()
        self.message(f"------ {self.project.name}: compiling ------", "head")
        self.runner.error_seen = False
        # F5 runs the project as it is in the editors: the resource file too, when the Resource Editor has
        # changes not saved yet (the String Table of the Application Wizard, before the project is saved)
        ed = self.res_editor
        changed = [] if ed is None or not ed.modified or ed.res is None else [ed.project]
        for p in changed:
            p.res_unsaved = ed.res.to_bytes()
        try:
            self.runner.start(self.project, self.codepage(), debug=True, first=first)
        finally:
            for p in changed:
                p.res_unsaved = None

    def run_end(self):
        self.runner.stop()

    def run_restart(self):
        self.runner.stop()
        QTimer.singleShot(100, self.run_start)

    def make_exe(self):
        if self.project is None:
            return
        p = self.project
        if p.is_addin:                          # an add-in (phase 12.5): a program the Add-In Manager loads
            return self.make_addin()
        if p.ptype in ("OleDll", "Control"):    # no program: checked, then registered for the programs that use it
            if p.untitled and not self.save_project_as():
                return
            self._sync_editors()
            self.output.clear()
            what = "ActiveX DLL" if p.ptype == "OleDll" else "ActiveX Control"
            self.message(f"------ {p.name}: making the {what} (registering it) ------", "head", show=True)
            import tempfile
            out = os.path.join(tempfile.gettempdir(), f"lbstudio-make-{os.getpid()}-{p.name}")
            self.runner.start(p, self.codepage(), make_target=out)
            return
        start = os.path.dirname(p.path) if p.path else (self.S["last_dir"] or os.path.expanduser("~"))
        name = os.path.splitext(os.path.basename(p.vbp.exe_name.replace(chr(92), "/")))[0] if p.vbp else p.name
        from .dialogs import MakeDialog
        saved = self.S["make_target"]
        if saved not in dict(MakeDialog.targets()) or not MakeDialog.available(saved)[0]:
            saved = paths.host_target()
        dlg = MakeDialog(self, start, name or p.name, saved)
        if not dlg.exec():
            return
        target, path = dlg.result_values()
        if not path:
            return
        self.S["make_target"] = target
        self._made_name(path)
        self._sync_editors()
        self.output.clear()
        label = dict(MakeDialog.targets()).get(target, target)
        self.message(f"------ {p.name}: making {path} ({label}) ------", "head", show=True)
        if target != "linux":
            os.makedirs(os.path.dirname(path), exist_ok=True)
        self.runner.start(p, self.codepage(), make_target=path, target=target)

    def make_addin(self, path=None):
        """File > Make <name>.dll... of an add-in project: the program of the add-in (for this computer), then it is
        registered in the Add-In Manager with the Display Name, Description and Initial Load Behavior of its Connect
        designer (VB6: oracles 12.5)"""
        p = self.project
        if p.untitled and not self.save_project_as():
            return
        if path is None:
            start = os.path.dirname(p.path)
            name = os.path.basename(p.vbp.exe_name.replace(chr(92), "/")) if p.vbp and p.vbp.exe_name else p.name + ".dll"
            if not name.lower().endswith(".dll"):
                name = os.path.splitext(name)[0] + ".dll"
            path, _ = QFileDialog.getSaveFileName(self, "Make Project", os.path.join(start, name),
                                                  "Add-In (*.dll);;All Files (*)")
            if not path:
                return
        self._made_name(path)
        self._sync_editors()
        self.output.clear()
        self.message(f"------ {p.name}: making the add-in {path} ------", "head", show=True)
        self._making_addin = p
        self.runner.start(p, self.codepage(), make_target=path)

    def register_made_addin(self, p, path):
        from . import addinhost as AH
        from lbvc.designers import addin_info
        d = p.addin_designer()
        info = addin_info(d.sf)
        progid = f"{p.name}.{d.name}"
        lb = info["loadbehavior"]
        AH.register_addin(self, {"progid": progid, "name": info["displayname"], "description": info["description"],
                                 "exe": os.path.abspath(path), "vbp": p.path, "startup": bool(lb & 1),
                                 "cmdline": bool(lb & 4), "cmdsafe": info["cmdlinesupport"]})
        return progid

    def _made_name(self, path, q=None):
        """Make writes the name of the program in the .vbp (ExeName32=, after Title) and, for an ActiveX EXE with
        Project / Binary Compatibility and no file for it yet, takes it as the compatible component
        (CompatibleEXE32=, VB6: DEV-8/pp1 stage 12)"""
        from . import projprops
        p = q or self.project
        v = p.vbp
        name = os.path.basename(path)
        rel = os.path.relpath(path, os.path.dirname(p.path)) if p.path else name
        if (v.get("ExeName32") or "") != rel:
            projprops.put(v, "ExeName32", rel, quote=True)
            p.dirty = True
        if p.ptype == "OleExe" and (v.get("CompatibleMode") or "0").strip('"') != "0" and not v.get("CompatibleEXE32"):
            projprops.put(v, "CompatibleEXE32", rel, quote=True)
            p.dirty = True
        self.update_title()

    def on_built(self, path):
        a = getattr(self, "_making_addin", None)
        if a is not None:                       # an add-in: registered for the Add-In Manager (phase 12.5)
            self._making_addin = None
            progid = self.register_made_addin(a, path)
            self.message(f"Built {path}; the add-in {progid} is in the Add-In Manager", "ok", show=True)
            self.statusBar().showMessage(f"Built {path}", 8000)
            return
        p = getattr(self, "_making", None) or self.project
        if getattr(self, "_group_jobs", None) is not None:   # Make Project Group: the next project
            QTimer.singleShot(0, self._next_group_job)
        if p is not None and p.ptype in ("OleDll", "Control"):
            self.message(f"{p.name} is registered: the programs that use it compile it into themselves", "ok", show=True)
            self.statusBar().showMessage(f"{p.name} registered", 8000)
            return
        if p is not None and p.ptype == "OleExe":
            self.message(f"Built {path}; {p.name} is registered", "ok", show=True)
            self.statusBar().showMessage(f"Built {path}", 8000)
            return
        self.message(f"Built {path}", "ok", show=True)
        self.statusBar().showMessage(f"Built {path}", 8000)

    def on_run_state(self, state):
        prev = getattr(self, "_prev_state", "design")
        self._prev_state = state
        if prev == "compiling" and state == "running":
            self.message("Compiled. The program runs.", "ok")
        elif prev in ("running", "break") and state == "design":
            self.message("The program ended.")
        design = state == "design"
        self._run_state_for_actions = state
        self.a_start.setEnabled(design or state == "break")
        self.a_start.setText("&Continue" if state == "break" else "&Start")
        self.a_full.setEnabled(design)
        self.a_make.setEnabled(design)
        self.a_end.setEnabled(not design)
        self.a_break.setEnabled(state == "running")
        if state != "break":
            self.clear_break_marks()
        if design:
            self.locals_win.clear()
            self.fill_watches_ooc()
        for sw in self.code_windows.values():   # Edit and Continue: see edit_while_running
            if isinstance(sw.widget(), CodeWindow):
                sw.widget().editor.SendScintilla(S.SCI_SETREADONLY, state in ("running", "break"))
        mode = {"design": "design", "compiling": "compiling", "running": "run", "break": "break"}[state]
        self.update_title(mode)

    def immediate_write(self, text):
        self.immediate.moveCursor(self.immediate.textCursor().MoveOperation.End)
        self.immediate.insertPlainText(text)
        self.immediate.ensureCursorVisible()

    def on_compile_error(self, rel, line, col, msg):
        self._group_jobs = None                 # Make Project Group stops at the first error, as VB6
        self._making = None
        m = self._module_of_file(rel)
        if m is not None and line > 0:
            self.message(f"{os.path.basename(m.rel.replace(chr(92), '/'))}({line - m.code_offset()}, {col}): "
                         f"Compile error: {msg}", "error", (m.name, line, col), show=True)
            self.message("Compilation failed.", "error")
            self.open_code(m, line - 1 - m.code_offset())
        else:
            self.message(msg if msg.startswith("lbvc:") else f"Compile error: {msg}", "error", show=True)
            self.message("Compilation failed.", "error")

    def on_compile_warning(self, rel, line, col, msg):
        m = self._module_of_file(rel)
        name = os.path.basename(m.rel.replace(chr(92), '/')) if m is not None else os.path.basename(rel)
        self.message(f"{name}: warning: {msg}", "warning", (m.name, line, col) if m is not None and line > 0 else None,
                     show=True)

    def on_runtime_error(self, number, msg):
        if getattr(self.runner, "error_seen", False):
            return                              # already shown when the debugger stopped at it
        self.message(f"Run-time error '{number}': {msg}", "error", show=True)

    # ================================================================ the debugger
    def _editor_line(self, m, file_line):
        return file_line - 1 - m.code_offset()

    def breakpoint_list(self):
        """[(module, line of the file)] for the program"""
        out = []
        for m in self.project.modules if self.project else []:
            for ln in sorted(self._breakpoints(m)):
                out.append((m.name, ln + 1 + m.code_offset()))
        return out

    def _breakpoints(self, m):
        """the editor lines with a breakpoint (the markers move with the text)"""
        sw = self.code_windows.get(m)
        if sw is not None:
            ed = sw.widget().editor
            lines, ln = set(), ed.SendScintilla(S.SCI_MARKERNEXT, 0, 1 << MARK_BP)
            while ln >= 0:
                lines.add(ln)
                ln = ed.SendScintilla(S.SCI_MARKERNEXT, ln + 1, 1 << MARK_BP)
            m.breakpoints = lines
        return getattr(m, "breakpoints", set())

    def toggle_breakpoint(self):
        cw = self.current_cw()
        if cw is None:
            return
        ed = cw.editor
        line = ed.line_from_pos(ed.pos())
        text = ed.get_range(ed.line_start(line), ed.line_end(line)).strip()
        m = cw.module
        on = ed.SendScintilla(S.SCI_MARKERGET, line) & (1 << MARK_BP)
        if not on and (not text or text.startswith("'") or text.lower().startswith(("dim ", "rem ")) or
                       text.lower() in ("end sub", "end function", "end property")):
            self.message("Can't set breakpoint on this line", "error", show=True)   # VB6 says the same
            self.statusBar().showMessage("Can't set breakpoint on this line", 4000)
            return
        fl = line + 1 + m.code_offset()
        if on:
            ed.SendScintilla(S.SCI_MARKERDELETE, line, MARK_BP)
            ed.SendScintilla(S.SCI_MARKERDELETE, line, MARK_BP_BACK)
            if not self.bp_disabled:
                self.runner.send(f"bp- {m.name} {fl}")
        else:
            ed.SendScintilla(S.SCI_MARKERADD, line, MARK_BP)
            ed.SendScintilla(S.SCI_MARKERADD, line, MARK_BP_BACK)
            if not self.bp_disabled:
                self.runner.send(f"bp+ {m.name} {fl}")
        self._breakpoints(m)

    def _bp_style(self, ed):
        """the breakpoint markers: dark red, or hollow and without the red line while they are disabled"""
        snd = ed.SendScintilla
        snd(S.SCI_MARKERDEFINE, MARK_BP, S.SC_MARK_CIRCLE)
        snd(S.SCI_MARKERSETFORE, MARK_BP, QColor("#800000"))
        snd(S.SCI_MARKERSETBACK, MARK_BP, QColor("#FFFFFF") if self.bp_disabled else QColor("#800000"))
        snd(S.SCI_MARKERDEFINE, MARK_BP_BACK, S.SC_MARK_EMPTY if self.bp_disabled else S.SC_MARK_BACKGROUND)
        snd(S.SCI_MARKERSETBACK, MARK_BP_BACK, QColor("#E8B4B4"))

    def disable_breakpoints(self, on=None):
        """Debug > Disable All Breakpoints: the breakpoints stay in the code, the program does not stop at them"""
        self.bp_disabled = self.a_bp_disable.isChecked() if on is None else bool(on)
        self.a_bp_disable.setChecked(self.bp_disabled)
        for sw in self.code_windows.values():
            if isinstance(sw.widget(), CodeWindow):
                for ed in sw.widget().panes():
                    self._bp_style(ed)
        if self.bp_disabled:
            self.runner.send("bpclear")
        else:
            for mod, ln in self.breakpoint_list():
                self.runner.send(f"bp+ {mod} {ln}")
        self.statusBar().showMessage("Breakpoints disabled" if self.bp_disabled else "Breakpoints enabled", 4000)

    def clear_breakpoints(self):
        for m in self.project.modules if self.project else []:
            sw = self.code_windows.get(m)
            if sw is not None:
                ed = sw.widget().editor
                ed.SendScintilla(S.SCI_MARKERDELETEALL, MARK_BP)
                ed.SendScintilla(S.SCI_MARKERDELETEALL, MARK_BP_BACK)
            m.breakpoints = set()
        self.runner.send("bpclear")

    def restore_breakpoints(self, cw):
        """a code window opens: the breakpoints and the bookmarks of its module again"""
        for ln in getattr(cw.module, "breakpoints", set()):
            cw.editor.SendScintilla(S.SCI_MARKERADD, ln, MARK_BP)
            cw.editor.SendScintilla(S.SCI_MARKERADD, ln, MARK_BP_BACK)
        for ln in getattr(cw.module, "bookmarks", set()):
            cw.editor.SendScintilla(S.SCI_MARKERADD, ln, stl.MARK_BOOKMARK)

    # ---------------------------------------------------------------- 8.3: the code window
    def comment_block(self, uncomment=False):
        """Comment Block / Uncomment Block (the Edit toolbar): the lines of the selection, or the caret's line"""
        cw = self.current_cw()
        if cw is None:
            return
        ed = cw.editor
        a, b = ed.sel_range()
        l0, l1 = ed.line_from_pos(a), ed.line_from_pos(b)
        l0, l1 = codeops.block_lines(l0, l1, a == b, b - ed.line_start(l1))
        caret = ed.pos()
        ed.SendScintilla(S.SCI_BEGINUNDOACTION)
        try:
            for ln in range(l1, l0 - 1, -1):    # from the end: the positions above stay
                p0, p1 = ed.line_start(ln), ed.line_end(ln)
                old = ed.get_range(p0, p1)
                new = codeops.uncomment_line(old) if uncomment else codeops.comment_line(old)
                if new != old:
                    ed.replace_range(p0, p1, new)
        finally:
            ed.SendScintilla(S.SCI_ENDUNDOACTION)
        if a != b:
            ed.set_sel(ed.line_start(l0), ed.line_end(l1))
        else:
            ln = ed.line_from_pos(caret)
            ed.goto_pos(min(ed.line_end(ln), ed.line_start(ln) + max(0, caret - ed.line_start(ln))))

    def _bookmarks(self, m):
        """the editor lines with a bookmark (they move with the text); kept in the module when its window closes"""
        sw = self.code_windows.get(m)
        if sw is not None and isinstance(sw.widget(), CodeWindow):
            ed = sw.widget().editor
            lines, ln = [], ed.SendScintilla(S.SCI_MARKERNEXT, 0, 1 << stl.MARK_BOOKMARK)
            while ln >= 0:
                lines.append(ln)
                ln = ed.SendScintilla(S.SCI_MARKERNEXT, ln + 1, 1 << stl.MARK_BOOKMARK)
            m.bookmarks = set(lines)
        return sorted(getattr(m, "bookmarks", set()))

    def toggle_code_bookmark(self):
        """Edit > Bookmarks > Toggle Bookmark: the caret's line"""
        cw = self.current_cw()
        if cw is None:
            return
        ed = cw.editor
        line = ed.line_from_pos(ed.pos())
        if ed.SendScintilla(S.SCI_MARKERGET, line) & (1 << stl.MARK_BOOKMARK):
            ed.SendScintilla(S.SCI_MARKERDELETE, line, stl.MARK_BOOKMARK)
        else:
            ed.SendScintilla(S.SCI_MARKERADD, line, stl.MARK_BOOKMARK)
        self._bookmarks(cw.module)

    def next_bookmark(self, forward=True):
        """Next / Previous Bookmark: in the code window, then in the next modules of the project (their windows
        open), round to the first"""
        if self.project is None:
            return
        mods = [m for m in self.project.modules if self._bookmarks(m)]
        if not mods:
            return
        cw = self.current_cw()
        cur_m = cw.module if cw is not None else None
        cur_line = cw.editor.line_from_pos(cw.editor.pos()) if cw is not None else -1
        order = self.project.modules
        start = order.index(cur_m) if cur_m in order else (-1 if forward else len(order))
        n = len(order)
        for step in range(n + 1):
            i = (start + (step if forward else -step)) % n
            m = order[i]
            marks = self._bookmarks(m)
            if not marks:
                continue
            if step == 0 and m is cur_m:
                cand = [x for x in marks if x > cur_line] if forward else [x for x in marks if x < cur_line]
            elif step == n and m is cur_m:      # round: the first (last) of the same module
                cand = marks
            else:
                cand = marks if m is not cur_m else []
            if cand:
                line = cand[0] if forward else cand[-1]
                c = self.open_code(m)
                c.editor.goto_line(line)
                return

    def clear_bookmarks(self):
        """Edit > Bookmarks > Clear All Bookmarks: of every module"""
        for m in self.project.modules if self.project else []:
            sw = self.code_windows.get(m)
            if sw is not None and isinstance(sw.widget(), CodeWindow):
                sw.widget().editor.SendScintilla(S.SCI_MARKERDELETEALL, stl.MARK_BOOKMARK)
            m.bookmarks = set()

    def add_procedure(self):
        """Tools > Add Procedure... (a code window must be the active window, as in VB6)"""
        from .dialogs import AddProcedureDialog
        cw = self.current_cw()
        if cw is None:
            self.statusBar().showMessage("Add Procedure: open the code window of a module first", 5000)
            return
        dlg = AddProcedureDialog(self, lambda *a: self.apply_add_procedure(cw, *a))
        dlg.exec()

    def _set_code_text(self, cw, new):
        """the code window gets the text new (\\n line ends) in one edit (one undo): only the part that differs"""
        ed = cw.editor
        raw = cw.text()                         # the line ends of the text (CR LF from VB6 files)
        eol = "\r\n" if "\r\n" in raw else "\n" if "\n" in raw else \
            "\r\n" if ed.SendScintilla(S.SCI_GETEOLMODE) == S.SC_EOL_CRLF else "\n"
        old_b = ed.get_bytes()
        new_b = new.replace("\n", eol).encode("utf-8")
        p = 0
        while p < len(old_b) and p < len(new_b) and old_b[p] == new_b[p]:
            p += 1
        q = 0
        while q < len(old_b) - p and q < len(new_b) - p and old_b[-1 - q] == new_b[-1 - q]:
            q += 1
        if p == len(old_b) == len(new_b):
            return
        ed.SendScintilla(S.SCI_BEGINUNDOACTION)
        ed.replace_range(p, len(old_b) - q, new_b[p:len(new_b) - q])
        ed.SendScintilla(S.SCI_ENDUNDOACTION)

    def procedure_attributes(self):
        """Tools > Procedure Attributes... for the procedures of the active code window (the caret's one first)"""
        from .dialogs import ProcedureAttributesDialog
        cw = self.current_cw()
        if cw is None:
            self.statusBar().showMessage("Procedure Attributes: open the code window of a module first", 5000)
            return None
        dlg = self.procedure_attributes_dialog(cw, ProcedureAttributesDialog)
        if dlg is None:
            return None
        dlg.exec()
        return dlg

    def procedure_attributes_dialog(self, cw, cls=None):
        """the dialog for the code window cw (not shown: the tests use it as it is)"""
        if cls is None:
            from .dialogs import ProcedureAttributesDialog as cls
        text = cw.text().replace("\r\n", "\n")
        procs, seen = [], set()
        for n, kind, _ in codeops.procedures(text):
            if n.lower() not in seen:
                seen.add(n.lower())
                procs.append((n, "Property" if kind.startswith("Property") else kind))
        if not procs:
            self.statusBar().showMessage("Procedure Attributes: the module has no procedures", 5000)
            return None
        line = cw.editor.line_from_pos(cw.editor.pos())
        start = None
        for n, _, l0 in codeops.procedures(text):
            if l0 <= line:
                start = n
        m = cw.module
        pages = [x.name for x in self.project.modules if x.kind == "PropertyPage"] if self.project else []
        help_file = (self.project.vbp.get("HelpFile") or "").strip('"') if self.project else ""

        def read(name):
            return codeops.read_attributes(cw.text().replace("\r\n", "\n"), name)

        def apply(name, attrs):
            self._set_code_text(cw, codeops.write_attributes(cw.text().replace("\r\n", "\n"), name, attrs))
        return cls(self, procs, read, apply, start, help_file, m.kind != "Module", pages)

    # ---------------------------------------------------------------- File > Print (printing.py)
    def printer(self):
        if getattr(self, "_printer", None) is None:
            from PyQt6.QtPrintSupport import QPrinter
            self._printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        return self._printer

    def print_setup(self):
        """File > Print Setup...: the printer, the paper and the orientation"""
        from PyQt6.QtPrintSupport import QPrintDialog
        d = QPrintDialog(self.printer(), self)
        d.setWindowTitle("Print Setup")
        return d.exec()

    def _active_module(self):
        cw = self.current_cw()
        if cw is not None:
            return cw.module
        d = self.current_designer()
        return getattr(d, "module", None) if d is not None else None

    def print_dialog(self):
        """File > Print... (VB6's dialog)"""
        from .dialogs import PrintDialog
        if self.project is None:
            return
        m = self._active_module()
        cw = self.current_cw()
        sel = cw.editor.sel_text() if cw is not None else ""
        pr = self.printer()
        dlg = PrintDialog(self, pr.printerName() or "(no printer)", bool(sel),
                          m is not None and getattr(m, "is_form", False), self.print_setup, self.help_contents)
        if not dlg.exec():
            return
        v = dlg.values()
        target = None
        if v["to_file"]:
            target, _ = QFileDialog.getSaveFileName(self, "Print to File", os.path.join(self.project_dir_or_home(),
                                                    self.project.name + ".pdf"), "PDF (*.pdf)")
            if not target:
                return
        n = self.do_print(v["range"], v["what"], quality=v["quality"], to_file=target)
        self.statusBar().showMessage(f"Printed {n} page(s)", 5000)

    def project_dir_or_home(self):
        return os.path.dirname(self.project.path) if self.project and self.project.path else os.path.expanduser("~")

    def print_layout(self, rng, what, per_page):
        """the pages Print makes (printing.layout): Range module / project / selection"""
        from . import printing
        m = self._active_module()

        def code_of(mod):
            sw = self.code_windows.get(mod)
            if sw is not None and isinstance(sw.widget(), CodeWindow):
                return sw.widget().text()
            return mod.code_text()
        if rng == "selection":
            cw = self.current_cw()
            jobs = printing.jobs([], code_of, what, (cw.module, cw.editor.sel_text())) if cw is not None else []
        elif rng == "project":
            jobs = printing.jobs(list(self.project.modules), code_of, what)
        else:
            jobs = printing.jobs([m] if m is not None else [], code_of, what)
        return printing.layout(jobs, per_page)

    def form_image(self, m):
        """Form Image: the client area of the form as the designer shows it, without the grid and the handles"""
        sw = self.design_windows.get(m)
        active = self.mdi.currentSubWindow()
        d = sw.widget() if sw is not None else self.open_designer(m)
        if active is not None and sw is None:
            self.mdi.setActiveSubWindow(active)
        if not isinstance(d, FormDesigner):
            return None
        old, grid = list(d.selection), d.show_grid
        d.select([])
        d.show_grid = False
        d.surface.update()
        try:
            return d.surface.grab()
        finally:
            d.show_grid = grid
            d.select(old)
            d.surface.update()

    def do_print(self, rng, what, printer=None, quality="High", to_file=None):
        """print (printing.print_pages): the number of pages"""
        from . import printing
        from PyQt6.QtPrintSupport import QPrinter
        pr = printer or self.printer()
        if printer is None:
            pr.setOutputFormat(QPrinter.OutputFormat.PdfFormat if to_file else QPrinter.OutputFormat.NativeFormat)
            pr.setOutputFileName(to_file or "")
            if quality != "High":
                pr.setResolution({"Medium": 300, "Low": 150, "Draft": 75}[quality])
        self._sync_editors()
        pages = self.print_layout(rng, what, printing.lines_per_page(pr))
        return printing.print_pages(pr, pages, self.form_image)

    def apply_add_procedure(self, cw, name, kind, scope, statics):
        """the procedure into the code window: None, or VB6's message (the dialog shows it)"""
        ed = cw.editor
        text = cw.text().replace("\r\n", "\n")
        try:
            new, line = codeops.add_procedure(text, name.strip(), kind, scope, statics)
        except ValueError as e:
            return str(e)
        self._set_code_text(cw, new)
        ed.goto_line(line)
        ed.goto_pos(ed.line_end(line))
        ed.setFocus()
        return None

    def debug_step(self, how):
        if self.runner.state == "break":
            self.runner.resume(how)
        elif self.runner.state == "design":
            self.run_start(first="step")             # F8 at design time: the first line

    def run_to_cursor(self):
        cw = self.current_cw()
        if cw is None:
            return
        ed = cw.editor
        fl = ed.line_from_pos(ed.pos()) + 1 + cw.module.code_offset()
        cmd = f"runto {cw.module.name} {fl}"
        if self.runner.state == "break":
            self.runner.resume(cmd)
        elif self.runner.state == "design":
            self.run_start(first=cmd)

    def clear_break_marks(self):
        for sw in self.code_windows.values():
            ed = sw.widget().editor
            for mk in (MARK_CUR, MARK_CUR_BACK, MARK_CALL):
                ed.SendScintilla(S.SCI_MARKERDELETEALL, mk)
        self.break_at = None

    def on_break(self, module, file_line, reason, detail):
        """the program stopped: the yellow line, Locals, Watches"""
        m = self.project.module(module) if self.project else None
        self.frame_depth = 0
        self.raise_()
        self.activateWindow()
        if m is not None:
            line = self._editor_line(m, file_line)
            cw = self.open_code(m)
            ed = cw.editor
            self.clear_break_marks()
            ed.SendScintilla(S.SCI_MARKERADD, line, MARK_CUR)
            ed.SendScintilla(S.SCI_MARKERADD, line, MARK_CUR_BACK)
            ed.goto_line(line)
            ed.setFocus()
            self.break_at = (m, line)
        self.last_break = (module, file_line, reason, detail)
        self.watch_win.select(-1)
        self.refresh_debug()
        if reason == "error":
            num, _, text = detail.partition("\t")
            self.runner.error_seen = True
            # VB6 asks End / Debug; here the program waits at the line (Debug) and Run > End ends it
            where = f" ({module}, line {file_line - m.code_offset()})" if m is not None else ""
            self.message(f"Run-time error '{num}': {text}{where}", "error",
                         (module, file_line, 0) if m is not None else None, show=True)
            self.message("The program is in break mode at the line of the error: F5 / F8 run the line again, "
                         "Alt+F5 / Alt+F8 go on with the error (to its error handler), Ctrl+F9 sets the next "
                         "statement, Run > End ends it.")
        elif reason == "watch":
            i = next((k for k, w in enumerate(self.watches) if str(w["id"]) == detail.strip()), -1)
            self._show_dock(self.dock_watch, self.watch_win.tree)
            self.watch_win.select(i)
            if i >= 0:
                self.statusBar().showMessage(f"Watch: {self.watches[i]['expr']}", 6000)

    def refresh_debug(self):
        d = self.frame_depth

        def got_stack(rows):
            self.stack = [(r[1], r[2], int(r[3]) if len(r) > 3 and r[3].isdigit() else 0) for r in rows if r[0] == "frame"]
            if d < len(self.stack):
                mod, proc, ln = self.stack[d]
                m = self.project.module(mod)
                label = mod if m is None or m.kind == "Module" else "Me"
                self.runner.request(f"locals {d}", lambda vrows: self.locals_win.fill(f"{self.project.name}.{mod}.{proc}",
                                                                                     vrows, label), multi=True)
            self.refresh_watches()              # the Context of the watches: the frame of the stack
        self.runner.request("stack", got_stack, multi=True)

    def _watch_cmd(self, w):
        return f"watch {w['id']}\t{w['kind']}\t{w['mod']}\t{w['proc']}\t{w['expr']}"

    def _watch_rows(self, values=None):
        """[(expression, value, type, context, kind)] for the Watches window; values: [(value, type)].
        Out of its context a watch shows <Out of context> and the type Empty, as in VB6"""
        from .debugwin import watch_context
        return [(w["expr"], *(values[i] if values else ("<Out of context>", "Empty")), watch_context(w), w["kind"])
                for i, w in enumerate(self.watches)]

    def fill_watches_ooc(self):
        self.watch_win.fill(self._watch_rows())

    def _in_context(self, w):
        if not getattr(self, "stack", None) or self.frame_depth >= len(self.stack):
            return not w["mod"]
        mod, proc, _ = self.stack[self.frame_depth]
        if w["mod"] and w["mod"].lower() != mod.lower():
            return False
        return not w["proc"] or w["proc"].lower() == proc.lower()

    def refresh_watches(self):
        vals = []

        def one(i):
            if i >= len(self.watches):
                self.watch_win.fill(self._watch_rows(vals))
                return
            w = self.watches[i]

            def typed(value, ok, text):
                vals.append((value, text if ok else ""))
                one(i + 1)

            def done(ok, text):
                if not ok:
                    vals.append(("<" + text + ">", ""))
                    one(i + 1)
                elif not self.runner.request(f"eval {self.frame_depth}\tTypeName({w['expr']})",
                                             lambda ok2, t2: typed(text, ok2, t2)):
                    typed(text, False, "")
            if not self._in_context(w) or not self.runner.request(f"eval {self.frame_depth}\t{w['expr']}", done):
                vals.append(("<Out of context>", "Empty"))
                one(i + 1)
        one(0)

    def _watch_changed(self):
        self.watch_win.hit = -1
        self._show_dock(self.dock_watch, self.watch_win.tree)
        if self.runner.state == "break":
            self.refresh_watches()
        else:
            self.fill_watches_ooc()

    def new_watch(self, expr, kind=0, mod="", proc=""):
        """a watch into the Watches window, and into the program when it breaks"""
        self._watch_id += 1
        w = {"id": self._watch_id, "expr": expr, "kind": kind, "mod": mod, "proc": proc}
        self.watches.append(w)
        if kind:
            self.runner.send(self._watch_cmd(w))
        self._watch_changed()
        return w

    def _watch_modules(self):
        """[(module, [procedures])] for the Context of Add Watch"""
        from . import vbmodel
        out = []
        for m in self.project.modules if self.project else []:
            sw = self.code_windows.get(m)
            text = sw.widget().text() if sw is not None and isinstance(sw.widget(), CodeWindow) else m.code_text()
            procs = []
            for name, kind, line in vbmodel.procedures(text):
                if name not in procs:
                    procs.append(name)
            out.append((m.name, procs))
        return out

    def _caret_context(self):
        """the module and the procedure of the caret: the Context that Add Watch proposes"""
        from . import vbmodel
        cw = self.current_cw()
        if cw is None:
            return "", ""
        line = cw.editor.line_from_pos(cw.editor.pos())
        proc = ""
        for name, kind, l0 in vbmodel.procedures(cw.text()):
            if l0 <= line:
                proc = name
        return cw.module.name, proc

    def call_stack(self):
        from .debugwin import CallStackDialog
        if self.runner.state != "break" or not getattr(self, "stack", None):
            self.message("Call Stack: the program is not in break mode.", show=True)
            return
        dlg = CallStackDialog(self, self.stack)
        if dlg.exec():
            d = dlg.chosen()
            self.frame_depth = max(0, d)
            mod, proc, ln = self.stack[self.frame_depth]
            m = self.project.module(mod)
            if m is not None and self.frame_depth > 0:
                cw = self.open_code(m)
                line = self._editor_line(m, ln)
                cw.editor.SendScintilla(S.SCI_MARKERDELETEALL, MARK_CALL)
                cw.editor.SendScintilla(S.SCI_MARKERADD, line, MARK_CALL)
                cw.editor.goto_line(line)
            self.refresh_debug_frame()

    def refresh_debug_frame(self):
        d = self.frame_depth
        mod, proc, ln = self.stack[d]
        m = self.project.module(mod)
        label = mod if m is None or m.kind == "Module" else "Me"
        self.runner.request(f"locals {d}", lambda rows: self.locals_win.fill(f"{self.project.name}.{mod}.{proc}", rows, label),
                            multi=True)
        self.refresh_watches()

    def show_next_statement(self):
        if self.break_at is not None:
            m, line = self.break_at
            cw = self.open_code(m)
            cw.editor.goto_line(line)

    def _word_or_selection(self):
        cw = self.current_cw()
        if cw is None:
            return ""
        ed = cw.editor
        sel = ed.sel_text()
        if sel and "\n" not in sel:
            return sel.strip()
        a, b = ed.word_at(ed.pos())
        # take the object before the dot too: Text1.Text
        line = ed.line_from_pos(a)
        before = ed.get_range(ed.line_start(line), a)
        import re
        mm = re.search(r"((?:[A-Za-z_][A-Za-z0-9_]*(?:\([^()]*\))?\.)*)$", before)
        return (mm.group(1) if mm else "") + ed.get_range(a, b)

    def add_watch(self):
        from .debugwin import AddWatchDialog
        mod, proc = self._caret_context()
        dlg = AddWatchDialog(self, {"expr": self._word_or_selection(), "kind": 0, "mod": mod, "proc": proc},
                             self._watch_modules(), self.project.name if self.project else "")
        if dlg.exec() and dlg.expression():
            w = dlg.watch()
            self.new_watch(w["expr"], w["kind"], w["mod"], w["proc"])

    def edit_watch(self, i):
        from .debugwin import AddWatchDialog
        if not 0 <= i < len(self.watches):
            return
        old = self.watches[i]
        dlg = AddWatchDialog(self, old, self._watch_modules(), self.project.name if self.project else "", "Edit Watch")
        if dlg.exec():
            if dlg.expression():
                w = dict(old, **dlg.watch())
                self.watches[i] = w
                self.runner.send(self._watch_cmd(w) if w["kind"] else f"wdel {w['id']}")
            else:
                self.delete_watch(i)
                return
            self._watch_changed()

    def delete_watch(self, i):
        if 0 <= i < len(self.watches):
            w = self.watches.pop(i)
            if w["kind"]:
                self.runner.send(f"wdel {w['id']}")
            self._watch_changed()

    def quick_watch(self):
        expr = self._word_or_selection()
        if not expr or self.runner.state != "break":
            return

        def done(ok, text):
            box = QMessageBox(self)
            box.setWindowTitle("Quick Watch")
            box.setText(f"Expression:\n{expr}\n\nValue:\n{text if ok else '<' + text + '>'}")
            add = box.addButton("&Add", QMessageBox.ButtonRole.ActionRole)
            box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
            box.exec()
            if box.clickedButton() is add:
                mod, proc = self._caret_context()
                self.new_watch(expr, 0, mod, proc)
        self.runner.request(f"eval {self.frame_depth}\t{expr}", done)

    # ---------------------------------------------------------------- 8.2: the rest of VB6's Debug menu
    def set_next_statement(self):
        """Debug > Set Next Statement (Ctrl+F9): the line of the caret runs next; a line of the procedure that
        stopped (VB6 says the same of another one)"""
        cw = self.current_cw()
        if cw is None or self.runner.state != "break":
            return
        ed = cw.editor
        line = ed.line_from_pos(ed.pos())
        fl = line + 1 + cw.module.code_offset()

        def done(ok, text):
            if not ok:
                self.message(text, "error", show=True)
                self.statusBar().showMessage(text, 6000)
        self.runner.request(f"next {cw.module.name} {fl}", done)

    def error_resume(self, cmd):
        """Alt+F5 (errgo) / Alt+F8 (errstep): on with the error the program stopped at"""
        if self.runner.state == "break":
            self.runner.resume(cmd)

    def trap_mode(self):
        """Error Trapping: 0 Break on All Errors, 1 Break in Class Module, 2 Break on Unhandled Errors"""
        if self.trap_session is not None:
            return self.trap_session
        try:
            n = int(self.S["error_trapping"])
        except (TypeError, ValueError):
            n = 1
        return n if n in (0, 1, 2) else 1

    def set_trap(self, n, session=True):
        """Toggle > Break on ...: for this session (session False: Tools > Options set it)"""
        self.trap_session = n if session else None
        self.runner.send(f"trap {self.trap_mode()}")
        self.statusBar().showMessage(("Break on All Errors", "Break in Class Module",
                                      "Break on Unhandled Errors")[self.trap_mode()], 4000)

    def edit_while_running(self, ed):
        """Edit and Continue: the program is compiled to native code, so a change to the code can't go into the
        running program; as VB6 does for the edits it can't continue with, the IDE asks to reset the project.
        OK ends the program and the edit is made; Cancel leaves the code as it was"""
        if self.runner.state not in ("running", "break") or getattr(self, "_asking_reset", False):
            return
        self._asking_reset = True
        try:
            r = QMessageBox.question(self, APP_NAME, "This action will reset your project, proceed anyway?",
                                     QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)
        finally:
            self._asking_reset = False
        if r != QMessageBox.StandardButton.Ok:
            return
        self.run_end()
        for sw in self.code_windows.values():   # the edit goes on (Scintilla reads the state again)
            if isinstance(sw.widget(), CodeWindow):
                sw.widget().editor.SendScintilla(S.SCI_SETREADONLY, False)

    def immediate_command(self, text):
        """? expression / Print expression, or a statement (x = 5, Form1.Show) in break mode"""
        t = text.strip()
        low = t.lower()
        if self.runner.state != "break":
            self.immediate.answer(None)
            self.statusBar().showMessage("The Immediate window runs statements in break mode (F8, F9, Ctrl+Break)", 6000)
            return
        if low.startswith("?") or low.startswith("print ") or low.startswith("debug.print "):
            expr = t[1:] if low.startswith("?") else t.split(None, 1)[1] if " " in t else ""
            if low.startswith("debug.print "):
                expr = t[12:]
            self.runner.request(f"eval {self.frame_depth}\t{expr.strip()}",
                                lambda ok, v: self.immediate.answer(v if ok else None) if ok else self._imm_error(v))
        else:
            def done(ok, v):
                self.immediate.answer(None)
                if not ok:
                    self._imm_error(v)
                else:
                    self.refresh_debug()
            self.runner.request(f"exec {self.frame_depth}\t{t}", done)

    def _imm_error(self, text):
        """the error of a line of the Immediate window: under that line, and in Output"""
        self.immediate.answer(text)
        self.message(text, "error")

    def data_tip(self, ed, pos):
        """Auto Data Tips: the value of the variable under the mouse in break mode"""
        if self.runner.state != "break" or pos < 0:
            return
        import re
        a, b = ed.word_at(pos)
        word = ed.get_range(a, b)
        if not word or not re.match(r"[A-Za-z_]", word):
            return
        line = ed.line_from_pos(a)
        before = ed.get_range(ed.line_start(line), a)
        mm = re.search(r"((?:[A-Za-z_][A-Za-z0-9_]*\.)*)$", before)
        expr = (mm.group(1) if mm else "") + word

        def done(ok, v):
            if ok and self.runner.state == "break":
                ed.SendScintilla(S.SCI_CALLTIPSHOW, a, f"{expr} = {v}".encode("utf-8"))
        self.runner.request(f"eval {self.frame_depth}\t{expr}", done)

    # ================================================================ tools, help, window
    def options(self):
        d = OptionsDialog(self)
        if d.exec():
            self.S.save()
            self.set_trap(None, session=False)          # the setting of Options, for this session too
            self.apply_grid_settings()
            self.apply_dockable()
            self.styles = self._load_styles()
            self.apply_theme()
            for sw in self.code_windows.values():
                for ed in sw.widget().panes():
                    self.setup_code_editor(ed)

    def is_dark(self):
        """the theme of the IDE (Tools > Options > Theme): light, dark, or system = as Ubuntu / Windows is"""
        mode = self.S["dark_mode"]
        if mode in ("light", "dark"):
            return mode == "dark"
        env = os.environ.get("LBVC_THEME", "").lower()     # as the runtime (the tests: light)
        if env in ("light", "dark"):
            return env == "dark"
        return system_dark()

    def apply_theme(self):
        """Fusion with the light or the dark colours; the forms of the designer stay light (designer.py)"""
        from .designer import light_palette
        app = QApplication.instance()
        app.setStyle("Fusion")
        if self.is_dark():
            pal = QPalette()
            for role, c in ((QPalette.ColorRole.Window, "#2d2d30"), (QPalette.ColorRole.WindowText, "#e0e0e0"),
                            (QPalette.ColorRole.Base, "#1e1e1e"), (QPalette.ColorRole.AlternateBase, "#2d2d30"),
                            (QPalette.ColorRole.Text, "#e0e0e0"), (QPalette.ColorRole.Button, "#3c3c3c"),
                            (QPalette.ColorRole.ButtonText, "#e0e0e0"), (QPalette.ColorRole.Highlight, "#264f78"),
                            (QPalette.ColorRole.HighlightedText, "#ffffff"), (QPalette.ColorRole.ToolTipBase, "#2d2d30"),
                            (QPalette.ColorRole.ToolTipText, "#e0e0e0"), (QPalette.ColorRole.BrightText, "#ffffff"),
                            (QPalette.ColorRole.Light, "#505050"), (QPalette.ColorRole.Midlight, "#454545"),
                            (QPalette.ColorRole.Mid, "#3a3a3a"), (QPalette.ColorRole.Dark, "#1a1a1a"),
                            (QPalette.ColorRole.Shadow, "#000000"), (QPalette.ColorRole.Link, "#4ea1ff"),
                            (QPalette.ColorRole.PlaceholderText, "#909090")):
                pal.setColor(role, QColor(c))
            for role in (QPalette.ColorRole.WindowText, QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText):
                pal.setColor(QPalette.ColorGroup.Disabled, role, QColor("#7a7a7a"))
            app.setPalette(pal)
        else:
            app.setPalette(light_palette())
        self._theme_dark = self.is_dark()
        # QMdiArea takes its background (Dark) once, when it is made: with the theme of the system then
        if getattr(self, "mdi", None) is not None:
            self.mdi.setBackground(app.palette().brush(QPalette.ColorRole.Dark))

    def system_theme_changed(self, *args):
        """Ubuntu / Windows went light or dark while the IDE runs: with Theme = system the IDE follows"""
        if self.S["dark_mode"] in ("light", "dark") or self.is_dark() == getattr(self, "_theme_dark", None):
            return
        self.apply_theme()
        self.styles = self._load_styles()
        for sw in self.code_windows.values():
            for ed in sw.widget().panes():
                self.setup_code_editor(ed)

    def toggle_tabbed(self):
        tabbed = self.mdi.viewMode() != QMdiArea.ViewMode.TabbedView
        if not self.S["layout_dockable"]:       # the Form Layout leaves the MDI area first
            self._unplace_layout()
        self.mdi.setViewMode(QMdiArea.ViewMode.TabbedView if tabbed else QMdiArea.ViewMode.SubWindowView)
        self.S["mdi_tabbed"] = tabbed
        self.S.save()
        self._place_layout()

    @staticmethod
    def chm_files(folder):
        """the .CHM files of a folder (any case of the extension)"""
        try:
            return [n for n in os.listdir(folder) if n.lower().endswith(".chm")]
        except OSError:
            return []

    def set_help_folder(self, path=None):
        """Help > Set Help Folder...: the folder of the user's own MSDN Library (its .CHM / .CHI files, for example
        .../MSDN-VB6); the same setting as Tools > Options. True when a folder with .CHM files was set"""
        if path is None:
            start = self.S["msdn_path"] or os.path.expanduser("~")
            path = QFileDialog.getExistingDirectory(self, "Help Folder - the .CHM files of your MSDN Library", start)
            if not path:
                return False
        if not self.chm_files(path):
            QMessageBox.warning(self, APP_NAME, f"There are no .CHM files in {path}.\n\n"
                                "Choose the folder of the .CHM (and .CHI) files of your MSDN Library.")
            return False
        self.S["msdn_path"] = path
        self.S.save()
        if getattr(self, "_helpwin", None) is not None:      # the window of the old folder goes
            self._helpwin.close()
        self._helpwin = None
        self._msdn = None
        self.statusBar().showMessage(f"Help folder: {path}", 6000)
        return True

    def help_library(self):
        """the user's MSDN Library (Help > Set Help Folder... or Tools > Options), or None with a message that offers
        to set the folder"""
        from . import msdn
        path = self.S["msdn_path"]
        if not path or not os.path.isdir(path) or not self.chm_files(path):
            box = QMessageBox(QMessageBox.Icon.Information, APP_NAME,
                              "The help of VB6 is the MSDN Library, which is not part of LinBASIC Studio.\n\n"
                              "If you own MSDN, set the folder of its .CHM files." +
                              (f"\n\n{path}: no .CHM files there." if path else ""), parent=self)
            b_set = box.addButton("Set Help &Folder...", QMessageBox.ButtonRole.AcceptRole)
            box.addButton(QMessageBox.StandardButton.Cancel)
            self._last_help_box = box
            box.exec()
            if box.clickedButton() is not b_set or not self.set_help_folder():
                return None
            path = self.S["msdn_path"]
        if not msdn.Library.available():
            QMessageBox.information(self, APP_NAME,
                                    "To read the .CHM files of MSDN, install one of:\n\n"
                                    "    sudo apt install libchm-bin     (extract_chmLib)\n"
                                    "    sudo apt install p7zip-full     (7z)")
            return None
        lib = getattr(self, "_msdn", None)
        if lib is None or lib.folder != path:
            lib = self._msdn = msdn.Library(path)
            self._helpwin = None
        return lib

    def help_window(self):
        from .helpwin import HelpWindow
        lib = self.help_library()
        if lib is None:
            return None
        if getattr(self, "_helpwin", None) is None:
            self._helpwin = HelpWindow(self, lib)
        return self._helpwin

    def help_contents(self, index=False, search=False):
        """Help > Contents / Index / Search: the user's MSDN Library on that tab (DEV-8/me3)"""
        hw = self.help_window()
        if hw is None:
            return
        if search:
            hw.show_search()
        elif index:
            hw.show_index("")
        else:
            hw.show_contents()

    # ================================================================ the menus as VB6's (Phase 8.7, DEV-8/me1)
    def make_group(self):
        """File > Make Project Group...: the projects chosen are built one after the other; with "Use Default
        Build Options" each where its Make would put it (ExeName32 in the project folder), without one the Make
        dialog of each project first (DEV-8/me2 picture 27)"""
        g = self.group
        if g is None or len(g.projects) < 2 or self.runner.state != "design":
            return
        from .dialogs import MakeGroupDialog, MakeDialog
        dlg = MakeGroupDialog(self, [q.name for q in g.projects])
        self._last_make_group = dlg
        if not dlg.exec():
            return
        projects = [g.projects[i] for i in dlg.chosen()]
        target = self.S["make_target"]
        if target not in dict(MakeDialog.targets()) or not MakeDialog.available(target)[0]:
            target = paths.host_target()
        jobs = []
        for q in projects:
            if q.ptype in ("OleDll", "Control"):              # registered, no file (as make_exe)
                import tempfile
                jobs.append((q, os.path.join(tempfile.gettempdir(), f"lbstudio-make-{os.getpid()}-{q.name}"), target))
                continue
            path = self.default_make_path(q, target)
            if not dlg.default_opts.isChecked():
                md = MakeDialog(self, os.path.dirname(path), os.path.splitext(os.path.basename(path))[0], target)
                md.setWindowTitle(f"Make Project - {q.name}")
                if not md.exec():
                    return                                     # Cancel: nothing is built
                t, path = md.result_values()
                if not path:
                    return
                jobs.append((q, path, t))
            else:
                jobs.append((q, path, target))
        self._sync_editors()
        self.output.clear()
        self._group_jobs = jobs
        self._next_group_job()

    def default_make_path(self, q, target):
        """where Make puts the program of a project without asking: its ExeName32 (the last Make) in the project
        folder, or what the Make dialog proposes for the system"""
        from .dialogs import MakeDialog
        folder = os.path.dirname(q.path) if q.path else (self.S["last_dir"] or os.path.expanduser("~"))
        exe = (q.vbp.get("ExeName32") or "").strip('"').replace(chr(92), "/") if q.vbp else ""
        if exe and (target == "linux") == (not exe.lower().endswith(".exe")):
            return os.path.normpath(exe if os.path.isabs(exe) else os.path.join(folder, exe))
        name = os.path.splitext(os.path.basename(exe))[0] if exe else q.name
        md = MakeDialog.__new__(MakeDialog)                    # only its rule of the file name
        md.project_dir, md.exe_name = folder, name
        return md.default_path(target)

    def _next_group_job(self):
        jobs = getattr(self, "_group_jobs", None)
        if not jobs:
            self._group_jobs = None
            self._making = None
            return
        q, path, target = jobs.pop(0)
        self._making = q
        if q.ptype not in ("OleDll", "Control"):              # registered, no program: no ExeName32
            self._made_name(path, q)
        from .dialogs import MakeDialog
        label = dict(MakeDialog.targets()).get(target, target)
        self.message(f"------ {q.name}: making {path} ({label}) ------", "head", show=True)
        if target != "linux":
            os.makedirs(os.path.dirname(path), exist_ok=True)
        self.runner.start(q, self.codepage(), make_target=path, target=target)

    def insert_file(self):
        """Edit > Insert File...: the text of a file at the caret of the code window"""
        ed = self.current_editor()
        if ed is None:
            return
        start = self.S["last_dir"] or (os.path.dirname(self.project.path) if self.project and self.project.path
                                       else os.path.expanduser("~"))
        path, _ = QFileDialog.getOpenFileName(self, "Insert File", start, INSERT_FILE_FILTER)
        if not path:
            return
        try:
            with open(path, "rb") as f:
                data = f.read()
        except OSError as e:
            QMessageBox.warning(self, APP_NAME, str(e))
            return
        text = data.decode(self.codepage() or "cp1252", errors="replace").replace("\r\n", "\n")
        ed.SendScintilla(S.SCI_REPLACESEL, 0, text.encode("utf-8"))

    def update_edit_menu(self):
        """VB6: "Can't Undo" / "Can't Redo" when there is nothing to undo; Cut / Copy need a selection; the code
        commands need a code window (DEV-8/me1 pictures 2, 27)"""
        ed = self.current_editor()
        d = self.current_designer() if ed is None else None
        can_undo = bool(ed.SendScintilla(S.SCI_CANUNDO)) if ed is not None else bool(d and d.undo_stack)
        can_redo = bool(ed.SendScintilla(S.SCI_CANREDO)) if ed is not None else False
        self.a_undo.setText("&Undo" if can_undo else "Can't Undo")
        self.a_redo.setText("&Redo" if can_redo else "Can't Redo")
        self.a_undo.setEnabled(can_undo)
        self.a_redo.setEnabled(can_redo)
        sel = ed is not None and ed.SendScintilla(S.SCI_GETSELECTIONSTART) != ed.SendScintilla(S.SCI_GETSELECTIONEND)
        has_ctl = d is not None and bool(getattr(d, "selection", None))     # controls (the form is not one)
        for a in (self.a_cut, self.a_copy):
            a.setEnabled(sel or has_ctl)
        self.a_delete.setEnabled(ed is not None or has_ctl)
        for a in (self.a_find, self.a_find_next, self.a_replace, self.a_indent, self.a_outdent, self.a_insert_file,
                  self.a_list_members, self.a_list_consts, self.a_quick_info, self.a_param_info, self.a_complete):
            a.setEnabled(ed is not None)

    def update_project_menu(self):
        """Project > Remove <name>, <Project> Properties... (VB6: the name of the module); Add MDI Form grey when
        the project has one (DEV-8/me2 picture 2)"""
        self._update_file_actions()
        self.a_add_mdi.setEnabled(self.project is not None and self.project.mdi_form() is None)

    def add_mdi_form(self):
        """Project > Add MDI Form: the dialog of VB6 (New / Existing) unless "Don't show this dialog in the
        future"; one MDI form per project"""
        p = self.project
        if p is None or p.mdi_form() is not None:
            return
        if not self.S["add_mdi_nodialog"]:
            from .dialogs import AddMDIFormDialog
            start = os.path.dirname(p.path) if p.path else (self.S["last_dir"] or os.path.expanduser("~"))
            dlg = AddMDIFormDialog(self, start)
            self._last_add_mdi = dlg
            if not dlg.exec():
                return
            if dlg.dont_show.isChecked():
                self.S["add_mdi_nodialog"] = True
                self.S.save()
            if dlg.path:
                self.add_file_path(dlg.path)
                return
        self.add_module("MDIForm")

    def fill_window_menu(self):
        """Window: Split | Tile Horizontally, Tile Vertically, Cascade, Arrange Icons | the open windows (VB6,
        DEV-8/me1 pictures 24, 37); Tabs / Windows is of LinBASIC Studio"""
        m = self.window_menu
        m.clear()
        cw = self.current_cw()
        a = m.addAction("S&plit")              # the code window in two panes (DEV-8/me2 pictures 5, 6)
        a.setCheckable(True)
        a.setEnabled(cw is not None)
        a.setChecked(cw is not None and cw.is_split())
        a.triggered.connect(lambda on, c=cw: c.set_split(on) if c is not None else None)
        self.a_split = a
        m.addSeparator()
        m.addAction("Tile &Horizontally", lambda: self.tile_windows(True))
        m.addAction("Tile &Vertically", lambda: self.tile_windows(False))
        m.addAction("&Cascade", self.mdi.cascadeSubWindows)
        m.addAction("&Arrange Icons", self.arrange_icons)
        m.addSeparator()
        m.addAction("Tabs / Windows", self.toggle_tabbed)
        subs = self.window_list()
        if subs:
            m.addSeparator()
        cur = self.mdi.activeSubWindow()
        for i, sw in enumerate(subs[:8]):       # VB6: the last activated first, at most 8 (DEV-8/me9-06..13)
            a = m.addAction(f"&{i + 1} {sw.windowTitle().replace('&', '&&')}",
                            lambda w=sw: self.activate_sub(w))
            a.setCheckable(True)
            a.setChecked(sw is cur)
        if len(subs) >= 8:                      # VB6: from 8 windows on, also with all 8 listed (me9-13)
            m.addAction("&More Windows...", lambda: self.more_windows(subs))

    def window_list(self):
        """the open code / designer windows, the one in front first, then as they were activated (VB6)"""
        order = self.mdi.subWindowList(QMdiArea.WindowOrder.ActivationHistoryOrder)
        subs = [sw for sw in reversed(order) if sw.isVisible() or sw.isMinimized()]
        cur = self.mdi.activeSubWindow()
        if cur in subs:
            subs.remove(cur)
            subs.insert(0, cur)
        return subs

    def activate_sub(self, w):
        if w.isMinimized():
            w.showNormal()
        self.mdi.setActiveSubWindow(w)

    def more_windows(self, subs=None):
        """Window > More Windows...: "Windows", all the windows by their titles (A-Z), the first one chosen,
        OK / Cancel (VB6, DEV-8/me9-07)"""
        from PyQt6.QtWidgets import QDialog, QListWidget, QDialogButtonBox, QVBoxLayout
        subs = sorted(subs if subs is not None else self.window_list(), key=lambda sw: sw.windowTitle())
        dlg = QDialog(self)
        dlg.setWindowTitle("Windows")
        lay = QVBoxLayout(dlg)
        lst = QListWidget()
        for sw in subs:
            lst.addItem(sw.windowTitle())
        lst.setCurrentRow(0)
        lst.itemDoubleClicked.connect(lambda _: dlg.accept())
        lay.addWidget(lst)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        lay.addWidget(bb)
        self._more_windows_dlg = dlg
        if dlg.exec() and 0 <= lst.currentRow() < len(subs):
            self.activate_sub(subs[lst.currentRow()])

    def tile_windows(self, horizontally):
        """Tile Horizontally: the windows one above the other, as wide as the area; Tile Vertically: side by side"""
        if self.mdi.viewMode() == QMdiArea.ViewMode.TabbedView:
            return
        subs = [sw for sw in self.mdi.subWindowList() if sw.isVisible() and not sw.isMinimized()]
        if not subs:
            return
        r = self.mdi.viewport().rect()
        n = len(subs)
        for i, sw in enumerate(subs):
            sw.showNormal()
            if horizontally:
                h = r.height() // n
                sw.setGeometry(0, i * h, r.width(), h if i < n - 1 else r.height() - i * h)
            else:
                w = r.width() // n
                sw.setGeometry(i * w, 0, w if i < n - 1 else r.width() - i * w, r.height())

    def arrange_icons(self):
        """Arrange Icons: the minimized windows in a row at the bottom of the area"""
        r = self.mdi.viewport().rect()
        x = 0
        for sw in self.mdi.subWindowList():
            if sw.isMinimized():
                sz = sw.size()
                sw.move(x, r.height() - sz.height())
                x += sz.width()

    def help_context_number(self, number, lib_short=None):
        """F1 by the help context number of a type library (MsgBox -> 1008978)"""
        from . import msdn
        hw = self.help_window()
        if hw is None:
            return False
        hit = hw.lib.by_context(number, msdn.helpfile_of_lib(lib_short) if lib_short else None)
        if hit is None:
            return False
        return hw.show_page(*hit)

    def help_keyword(self, word):
        hw = self.help_window()
        if hw is None:
            return
        topics = hw.lib.by_keyword(word, main_only=True) or hw.lib.by_keyword(word)
        if topics:
            hw.show_topics(topics)
        else:
            hw.show_index(word)

    def object_browser(self):
        """View > Object Browser (F2): the classes and members of the libraries and of the project"""
        from .objbrowser import ObjectBrowser, ObjectBrowserWindow
        sw = getattr(self, "_objbrowser", None)
        if sw is None and self.is_dockable("objbrowser"):    # Options > Docking: Object Browser dockable
            sw = QDockWidget("Object Browser", self)
            sw.setObjectName("ObjectBrowser")
            sw.setWidget(ObjectBrowser(self))
            self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, sw)
            self._objbrowser = sw
        elif sw is None:                       # VB6 at first: a window among the code and form windows
            sw = ObjectBrowserWindow(self)
            sw.setWidget(ObjectBrowser(self))
            self.mdi.addSubWindow(sw)
            sw.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
            sw.setWindowTitle("Object Browser")
            sw.resize(760, 520)
            self._objbrowser = sw
            sw.show()
        else:
            sw.widget().refresh()
        if isinstance(sw, QDockWidget):
            sw.show()
            sw.raise_()
        else:
            self.mdi.setActiveSubWindow(sw)
        return sw.widget()

    def browse_word(self):
        """Definition (Shift+F2) on a name of a library (MsgBox, Text1.Text): the Object Browser shows it, as in
        VB6"""
        cw = self.current_cw()
        if cw is None:
            return False
        from . import intellisense as IS
        ed = cw.editor
        a, b = ed.word_at(ed.pos())
        word = ed.get_range(a, b)
        if not word:
            return False
        eng = self.intellisense()
        line = ed.line_from_pos(a)
        text = ed.all_text()
        got = IS.chain_before(ed.get_range(ed.line_start(line), a) + word)
        chain = (got[0] + [word]) if got else [word]
        if chain[0] == "<with>":
            wo = IS.with_object(text, line)
            chain = (IS.split_chain(wo) if wo else []) + [word]
        mem = eng.callable(chain, cw.module, line, text) or (eng.globals().get(word.lower()) if len(chain) == 1 else None)
        cls = None
        if len(chain) > 1:
            typ = eng.resolve(chain[:-1], cw.module, line, text)
            if typ and typ[0] in ("lib", "control") and isinstance(typ[1], str):
                cls = typ[1].split(".")[-1]
        lib = mem.lib if mem is not None else None
        if lib is None:
            hit = eng.tl.find(word)                 # a class: As TextBox
            if hit is None:
                return False
            lib, cls = hit[0], None
        ob = self.object_browser()
        return ob.reveal(lib, mem.name if mem is not None else word, cls)

    def context_help(self):
        """F1: help on the word at the caret, the property of the Properties window, the selected
        control of the designer"""
        if self.help_window() is None:
            return
        w = QApplication.focusWidget()
        ob = getattr(self, "_objbrowser", None)
        if ob is not None and w is not None and (w is ob.widget() or ob.widget().isAncestorOf(w)):
            ob.widget().help()
            return
        cw = self.current_cw()
        eng = self.intellisense()
        if cw is not None and w is cw.editor:
            ed = cw.editor
            a, b = ed.word_at(ed.pos())
            word = ed.get_range(a, b)
            if not word:
                return self.help_contents(index=True)
            from . import intellisense as IS
            line = ed.line_from_pos(a)
            text = ed.all_text()
            got = IS.chain_before(ed.get_range(ed.line_start(line), a) + word)
            chain = (got[0] + [word]) if got else [word]
            if chain[0] == "<with>":
                wo = IS.with_object(text, line)
                chain = (IS.split_chain(wo) if wo else []) + [word]
            mem = eng.callable(chain, cw.module, line, text) or (eng.globals().get(word.lower()) if len(chain) == 1 else None)
            if mem is None and len(chain) == 1:
                typ = eng.type_of_name(word, cw.module, line, text)
                if typ and typ[0] == "control":
                    hit = eng.tl.find(typ[1])
                    if hit and hit[1].get("helpcontext") and self.help_context_number(hit[1]["helpcontext"], hit[0]):
                        return
            if mem is not None and mem.help and self.help_context_number(mem.help, mem.lib):
                return
            return self.help_keyword(word)
        if w is not None and (w is self.props.grid or self.props.grid.isAncestorOf(w)):
            r = self.props.grid.currentRow()
            p = self.props.rows[r] if 0 <= r < len(self.props.rows) else None
            ctl = self.props.ctl
            if p is not None and not isinstance(p, str) and ctl is not None:
                for mem in eng.tl.members("VB." + ctl.type_name):
                    if mem.name.lower() == p.name.lower() and mem.help:
                        if self.help_context_number(mem.help, mem.lib):
                            return
                return self.help_keyword(p.name.strip("()") + " property")
        # the Toolbox: the control of the button pressed (VB6: F1 on a tool)
        if w is not None and (w is self.toolbox or self.toolbox.isAncestorOf(w)):
            tool = self.toolbox.current()
            if tool and self._help_type(tool):
                return
            return self.help_keyword("Toolbox")
        # the Immediate window: the word at the cursor, as in a code window
        if w is not None and w is self.immediate:
            c = self.immediate.textCursor()
            c.select(c.SelectionType.WordUnderCursor)
            word = c.selectedText().strip()
            if word:
                mem = eng.globals().get(word.lower())
                if mem is not None and mem.help and self.help_context_number(mem.help, mem.lib):
                    return
                return self.help_keyword(word)
            return self.help_keyword("Immediate window")
        # the Project Explorer: its own topic
        if w is not None and (w is self.tree or self.tree.isAncestorOf(w)):
            return self.help_keyword("Project Explorer")
        d = self.current_designer()
        if d is not None:
            ctl = d.primary.ctl if d.primary else d.root
            if self._help_type(getattr(ctl, "cls", "") or ctl.type_name):     # VB.Form, MSComctlLib.ListView ...
                return
            return self.help_keyword(ctl.type_name)
        self.help_contents(index=True)

    def _help_type(self, name):
        """F1 on a control type ("CommandButton", "MSComctlLib.ListView"): its page by the help context number"""
        eng = self.intellisense()
        hit = eng.tl.find(name if "." in name else "VB." + name) or eng.tl.find(name)
        return bool(hit and hit[1].get("helpcontext") and self.help_context_number(hit[1]["helpcontext"], hit[0]))

    def about(self):
        QMessageBox.about(self, "About " + APP_NAME,
                          f"<b>{APP_NAME}</b> {VERSION}<br><br>A development environment for the VB6 dialect.<br>"
                          "Compiler lbvc, runtime libvbrt, editor from Linpad++.<br><br>"
                          "Copyright © 2026 Demosthenes Koptsis<br>GPL-3.0-or-later")

    # ================================================================ title, close
    def _update_file_actions(self):
        has = self.project is not None
        for a in getattr(self, "project_actions", []):
            a.setEnabled(has)
        if has:                                  # Start, Break, End ... as the program is
            state = getattr(self, "_run_state_for_actions", "design")
            design = state == "design"
            self.a_start.setEnabled(design or state == "break")
            self.a_full.setEnabled(design)
            self.a_make.setEnabled(design)
            self.a_end.setEnabled(not design)
            self.a_break.setEnabled(state == "running")
            grp = getattr(self, "group", None) is not None
            self.a_close_group.setEnabled(grp)
            self.a_save_group.setEnabled(grp)
            self.a_remove_proj.setEnabled(grp)
            # VB6: Make Project Group... for a group of two or more projects (DEV-8/me2 pictures 23, 26)
            self.a_make_group.setEnabled(design and grp and len(self.group.projects) > 1)
        m = self.selected_module()
        fname = m.name if m else "<File>"       # VB6: "Save Form1", "Remove Form1" (DEV-8/me1)
        self.a_save_file.setText(f"&Save {fname}")
        self.a_save_file_as.setText(f"Save {fname} &As...")
        self.a_remove.setText(f"&Remove {fname}")
        if self.project:
            ext = {"OleDll": ".dll", "Control": ".ocx"}.get(self.project.ptype, ".exe")
            exe = self.project.vbp.get("ExeName32") or self.project.name + ext
            self.a_make.setText(f"Ma&ke {exe.strip(chr(34))}...")
            self.a_proj_props.setText(f"{self.project.name} Prop&erties...")

    def update_title(self, mode=None):
        mode = mode or {"design": "design", "compiling": "compiling", "running": "run", "break": "break"}.get(self.runner.state, "design")
        name = self.project.name if self.project else ""
        if self.group is not None and self.project is not None:
            name = f"{self.group.name} ({self.project.name})"
        dirty = " *" if self.project and (self.group.any_dirty() if self.group is not None else self.project.any_dirty()) else ""
        self.setWindowTitle(f"{name}{dirty} - {APP_NAME} [{mode}]" if self.project else APP_NAME)
        self._update_file_actions()

    def showEvent(self, e):
        """the first time the IDE shows: the window made before it (the form of the new project) takes its
        whole tab, as the ones opened later do"""
        super().showEvent(e)
        if getattr(self, "_shown_once", False):
            return
        self._shown_once = True
        sw = self.mdi.activeSubWindow() or self.current_sub()
        if sw is not None and self.mdi.viewMode() == QMdiArea.ViewMode.TabbedView:
            sw.showMaximized()
        QTimer.singleShot(0, self.update_toolbar_pos)  # the form of the new project: "0, 0"  "4800 x 3600"

    def closeEvent(self, e):
        if self.runner.state != "design":
            self.runner.stop()
        self.addins_begin_shutdown()            # (VB6: OnBeginShutdown before it asks to save)
        if not self.confirm_discard() or not self.res_confirm():
            e.ignore()
            return
        self.S["geometry"] = bytes(self.saveGeometry().toHex()).decode()
        self.S["window_state"] = bytes(self.saveState().toHex()).decode()
        self.S.save()
        v = getattr(self, "_apiviewer", None)
        if v is not None:                       # the API Viewer of the add-in goes with the IDE
            v.close()
        if getattr(self, "addin_hosts", None):
            from . import addinhost as AH
            AH.close_events(self)               # the projects close: ItemRemoved ... AfterCloseFile
        self.unload_user_addins(shutdown=True)  # OnDisconnection (ext_dm_HostShutdown)
        e.accept()

    def restore_layout(self):
        from PyQt6.QtCore import QByteArray
        try:
            if self.S["geometry"]:
                self.restoreGeometry(QByteArray.fromHex(self.S["geometry"].encode()))
            if self.S["window_state"]:
                self.restoreState(QByteArray.fromHex(self.S["window_state"].encode()))
        except Exception:                       # noqa: BLE001 - a bad saved layout is ignored
            pass
        self._place_layout()                    # the Form Layout without Dockable: not in its dock
