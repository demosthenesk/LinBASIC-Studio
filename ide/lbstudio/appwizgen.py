"""What the Application Wizard (phase 9.2a) writes: the main form (MDI, SDI or Explorer Style) with its menus,
toolbar and status bar, Module1 with Sub Main, frmDocument, frmBrowser, the standard forms and the form templates
(Splash, Login, Options, About, Dialog, ODBC Log In, Tip of the Day, Web Browser), the strings of the .res file
and the .rwp profiles.

The structure follows the VB6 oracle DEV-9/aw1 (the names of the forms and controls, mnu<Menu><Item>, the
.rwp format, the captions of the wizard's own pages and lists), so that a project and a profile of VB6 and
of LinBASIC Studio look alike. The code, the wording and the layout of the forms are our own: the VB6
templates are Microsoft's. No Qt here: the pictures are placeholders ("@pic:<name>:<size>") that the IDE
replaces with the BMPs of tbicons (attach_pictures).
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
import re

APP_MDI, APP_SDI, APP_EXPLORER = 0, 1, 2
TYPE_NAMES = {APP_MDI: "Multiple Document Interface (MDI)", APP_SDI: "Single Document Interface (SDI)",
              APP_EXPLORER: "Explorer Style"}

# the Object= lines (.vbp) and their headers (.frm)
OBJ_COMDLG = ("{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0", "ComDlg32.OCX")
OBJ_COMCTL = ("{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0", "mscomctl.OCX")
OBJ_RICHTX = ("{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0", "richtx32.Ocx")
OBJ_WEB = ("{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0", "ieframe.dll")

G_BUTTONS = "{66833FE8-8583-11D1-B16A-00C0F0283628}"
G_BUTTON = "{66833FEA-8583-11D1-B16A-00C0F0283628}"
G_PANELS = "{8E3867A5-8586-11D1-B16A-00C0F0283628}"
G_PANEL = "{8E3867AB-8586-11D1-B16A-00C0F0283628}"
G_IMAGES = "{2C247F25-8591-11D1-B16A-00C0F0283628}"
G_IMAGE = "{2C247F27-8591-11D1-B16A-00C0F0283628}"
G_TABS = "{1EFB6598-857C-11D1-B16A-00C0F0283628}"
G_TAB = "{1EFB659A-857C-11D1-B16A-00C0F0283628}"
G_FONT = "{0BE35203-8F91-11CE-9DE3-00AA004BB851}"

SEPARATOR = "[Separator]"
MRU = "MRU"

# ------------------------------------------------------------------------------------------------ the menus
# the lists of the Menus page as the .rwp keeps them (VB6 oracle aw6/test1.rwp): "+" checked, "-" not, the menu
# first, then its items, Caption|Key|
RWP_MENUS = [
    "+&File|File|+&New|New|+&Open...|Open|+&Close|Close|+[Separator]|[Separator]|+&Save|Save|+Save &As...|SaveAs|"
    "+Save A&ll|SaveAll|+[Separator]|[Separator]|+Propert&ies|Properties|+[Separator]|[Separator]|"
    "+Page Set&up...|PageSetup|+Print Pre&view|PrintPreview|+&Print...|Print|+[Separator]|[Separator]|"
    "+Sen&d...|Send|+[Separator]|[Separator]|+MRU List|MRU|+E&xit|Exit|",
    "+&Edit|Edit|+&Undo|Undo|+[Separator]|[Separator]|+Cu&t|Cut|+&Copy|Copy|+&Paste|Paste|"
    "+Paste &Special...|PasteSpecial|",
    "+&View|View|+&Toolbar|Toolbar|+Status &Bar|StatusBar|+[Separator]|[Separator]|+&Refresh|Refresh|"
    "+&Options...|Options|+&Web Browser|WebBrowser|",
    "-&Tools|Tools|+&Options...|Options|",
    "+&Window|Window|+&New Window|NewWindow|+[Separator]|[Separator]|+&Cascade|Cascade|"
    "+Tile &Horizontal|TileHorizontal|+Tile &Vertical|TileVertical|+&Arrange Icons|ArrangeIcons|",
    "+&Help|Help|+&Contents|Contents|+&Search For Help On...|SearchForHelpOn|+[Separator]|[Separator]|"
    "+&About |About|",
]
RWP_MENUS_EXP = [
    "+&File|File|+&Open...|Open|+&Find|Find|+[Separator]|[Separator]|+Sen&d to|SendTo|+[Separator]|[Separator]|"
    "+&New|New|+[Separator]|[Separator]|+&Delete|Delete|+Rena&me|Rename|+Propert&ies|Properties|"
    "+[Separator]|[Separator]|+MRU List|MRU|+[Separator]|[Separator]|+&Close|Close|",
    "+&Edit|Edit|+&Undo|Undo|+[Separator]|[Separator]|+Cu&t|Cut|+&Copy|Copy|+&Paste|Paste|"
    "+Paste &Special...|PasteSpecial|+[Separator]|[Separator]|+Select &All|SelectAll|"
    "+&Invert Selection|InvertSelection|",
    "+&View|View|+&Toolbar|Toolbar|+Status &Bar|StatusBar|+[Separator]|[Separator]|+ListView Mode|ListViewMode|"
    "+[Separator]|[Separator]|+Arrange &Icons|ArrangeIcons|+[Separator]|[Separator]|+&Refresh|Refresh|"
    "+&Options...|Options|+&Web Browser|WebBrowser|",
]
SHORTCUTS = {"FileNew": "^N", "FileOpen": "^O", "EditCut": "^X", "EditCopy": "^C", "EditPaste": "^V",
             "EditSelectAll": "^A"}
LISTVIEW_MODES = [("Lar&ge Icons", "View Large Icons"), ("S&mall Icons", "View Small Icons"),
                  ("&List", "View List"), ("&Details", "View Details")]


class MenuItem:
    def __init__(self, caption, key, on=True):
        self.caption, self.key, self.on = caption, key, on

    @property
    def separator(self):
        return self.key == SEPARATOR

    def copy(self):
        return MenuItem(self.caption, self.key, self.on)

    def __repr__(self):
        return f"MenuItem({self.caption!r}, {self.key!r}, {self.on})"


class Menu(MenuItem):
    def __init__(self, caption, key, on=True, items=()):
        super().__init__(caption, key, on)
        self.items = list(items)

    def copy(self):
        return Menu(self.caption, self.key, self.on, [i.copy() for i in self.items])


def parse_menu(s):
    """a menu of a .rwp (MENU_n=) -> Menu"""
    parts = s.split("|")
    pairs = []
    for i in range(0, len(parts) - 1, 2):
        cap, key = parts[i], parts[i + 1]
        if not cap:
            continue
        on = not cap.startswith("-")
        if cap[:1] in "+-":
            cap = cap[1:]
        pairs.append(MenuItem(cap, key, on))
    if not pairs:
        return None
    head = pairs[0]
    return Menu(head.caption, head.key, head.on, pairs[1:])


def format_menu(m):
    """Menu -> its line of a .rwp"""
    return "".join(f"{'+' if it.on else '-'}{it.caption}|{it.key}|" for it in [m] + m.items)


def default_menus(app_type):
    """the lists of the Menus page for an interface type: MDI all but Tools; SDI and Explorer without Tools and
    Window; Explorer its own File, Edit, View (aw1 AW8 - AW10)"""
    src = (RWP_MENUS_EXP + RWP_MENUS[3:]) if app_type == APP_EXPLORER else RWP_MENUS
    menus = [parse_menu(s) for s in src]
    for m in menus:
        if m.key == "Window" and app_type != APP_MDI:
            m.on = False
    return menus


# ------------------------------------------------------------------------------------------------ the toolbar
class TbButton:
    """a button of the toolbar: Key|Caption|ToolTip| in the .rwp; a separator has no key"""
    def __init__(self, key, caption=None, tip=None):
        self.key = key
        self.caption = caption if caption is not None else (key or SEPARATOR)
        self.tip = tip if tip is not None else key

    @property
    def separator(self):
        return not self.key

    def __repr__(self):
        return f"TbButton({self.key!r})"


def _buttons(names):
    return [TbButton("") if n == "|" else TbButton(n) for n in names]


TOOLBAR_MDI = ["New", "Open", "Save", "|", "Print", "|", "Cut", "Copy", "Paste", "|", "Bold", "Italic",
               "Underline", "|", "Align Left", "Center", "Align Right"]
TOOLBAR_EXPLORER = ["|", "Back", "Forward", "|", "Cut", "Copy", "Paste", "Delete", "|", "Properties", "|",
                    "View Large Icons", "View Small Icons", "View List", "View Details"]
GROUP_BUTTONS = {"Align Left", "Center", "Align Right", "View Large Icons", "View Small Icons", "View List",
                 "View Details"}                                   # Style 2 (tbrButtonGroup)
BROWSER_BUTTONS = ["Back", "Forward", "Stop", "Refresh", "Home", "Search"]


def default_toolbar(app_type):
    return _buttons(TOOLBAR_EXPLORER if app_type == APP_EXPLORER else TOOLBAR_MDI)


def format_toolbar(buttons):
    """-> the lines of [AppWizard6_ToolbarWizard]"""
    return [f"Toolbar{i}={b.key}|{b.caption}|{b.tip}|" for i, b in enumerate(buttons, 1)]


def parse_toolbar(values):
    out = []
    for v in values:
        p = (v.split("|") + ["", "", ""])[:3]
        out.append(TbButton("") if not p[0] or p[1] == SEPARATOR else TbButton(p[0], p[1] or p[0], p[2]))
    return out


# ------------------------------------------------------------------------------------------------ the choices
STANDARD_FORMS = ["Splash", "Login", "Options", "About"]          # OtherForms0..3 of the .rwp
TEMPLATES = ["Dialog", "ODBC Log In", "Tip of the Day", "Web Browser"]
DEFAULT_URL = "http://www.microsoft.com"


class AppOptions:
    def __init__(self, app_type=APP_MDI, name="Project1"):
        self.app_type = app_type
        self.name = name
        self.menus = default_menus(app_type)
        self.toolbar = default_toolbar(app_type)
        self.resources = False
        self.internet = False
        self.url = DEFAULT_URL
        self.forms = []            # of STANDARD_FORMS
        self.templates = []        # of TEMPLATES

    def set_type(self, app_type):
        """the Interface Type page: the menus and the toolbar of the new type"""
        if app_type != self.app_type:
            self.app_type = app_type
            self.menus = default_menus(app_type)
            self.toolbar = default_toolbar(app_type)

    # -------------------------------------------------------------- the profile (.rwp)
    def to_rwp(self):
        """the profile as VB6 writes it (aw6/test1.rwp): the lists of both kinds of menus are kept"""
        std = [m.copy() for m in (default_menus(APP_MDI) if self.app_type == APP_EXPLORER else self.menus)]
        exp = [m.copy() for m in (self.menus if self.app_type == APP_EXPLORER
                                  else default_menus(APP_EXPLORER))]
        lines = ["[AppWizard6]", f"ApplicationType={self.app_type}"]
        for i in range(6):
            if i < len(std):
                lines.append(f"MENU_{i}={format_menu(std[i])}")
            if i < 3 and i < len(exp):
                lines.append(f"MENUEXP_{i}={format_menu(exp[i])}")
        lines.append(f"Internet={'Yes' if self.internet else 'No'}")
        if self.internet:
            lines.append(f"URL={self.url}")
        lines.append(f"Resources={'Yes' if self.resources else 'No'}")
        for i, f in enumerate(STANDARD_FORMS):
            lines.append(f"OtherForms{i}={1 if f in self.forms else 0}")
        for i, t in enumerate(TEMPLATES):
            if t in self.templates:
                lines.append(f"Template{i}=1")
        lines.append("[AppWizard6_ToolbarWizard]")
        lines += format_toolbar(self.toolbar)
        return "\r\n".join(lines) + "\r\n"

    @classmethod
    def from_rwp(cls, text):
        sec, vals = None, {}
        for raw in text.splitlines():
            line = raw.strip()
            if line.startswith("[") and line.endswith("]"):
                sec = line[1:-1]
                continue
            if "=" in line and sec:
                k, v = line.split("=", 1)
                vals.setdefault(sec, []).append((k.strip(), v))
        a = dict(vals.get("AppWizard6", []))
        try:
            t = int(a.get("ApplicationType", "0"))
        except ValueError:
            t = APP_MDI
        o = cls(t if t in TYPE_NAMES else APP_MDI)
        pref = "MENUEXP_" if o.app_type == APP_EXPLORER else "MENU_"
        menus = []
        for i in range(6):
            k = (pref if i < 3 else "MENU_") + str(i)
            m = parse_menu(a[k]) if k in a else None
            if m is not None:
                menus.append(m)
        if menus:
            o.menus = menus
        o.internet = a.get("Internet", "No").lower() == "yes"
        o.url = a.get("URL", DEFAULT_URL)
        o.resources = a.get("Resources", "No").lower() == "yes"
        o.forms = [f for i, f in enumerate(STANDARD_FORMS) if a.get(f"OtherForms{i}", "0").strip() == "1"]
        o.templates = [t for i, t in enumerate(TEMPLATES) if a.get(f"Template{i}", "0").strip() == "1"]
        tb = [v for k, v in vals.get("AppWizard6_ToolbarWizard", []) if k.lower().startswith("toolbar")]
        if tb:
            o.toolbar = parse_toolbar(tb)
        return o


# ------------------------------------------------------------------------------------------------ .frm text
def _fmt(v):
    if isinstance(v, bool):
        return "-1  'True" if v else "0   'False"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return repr(v)
    if isinstance(v, Raw):
        return str(v)
    return '"' + str(v).replace('"', '""') + '"'


class Raw(str):
    """a value as it is (Shortcut = ^N, enum values with their comments)"""


def pic(name, size=16):
    """the placeholder of a picture (attach_pictures)"""
    return f"@pic:{name}:{size}"


_PIC_RE = re.compile(r"^@pic:(.+):(\d+)$")


class Frm:
    """writes the controls of a .frm as VB6 lays them out"""
    def __init__(self):
        self.lines = []
        self.depth = 0
        self.objects = []          # (guid, file) of the header

    def obj(self, o):
        if o not in self.objects:
            self.objects.append(o)

    def begin(self, cls, name):
        self.lines.append("   " * self.depth + f"Begin {cls} {name} ")
        self.depth += 1

    def end(self):
        self.depth -= 1
        self.lines.append("   " * self.depth + "End")

    def p(self, name, value):
        if value is None:
            return
        self.lines.append("   " * self.depth + f"{name:<16}=   {_fmt(value)}")

    def props(self, **kw):
        for k, v in kw.items():
            self.p(k.replace("__", "."), v)

    def bprop(self, name, guid=None):
        self.lines.append("   " * self.depth + f"BeginProperty {name} " + (f"{guid} " if guid else ""))
        self.depth += 1

    def eprop(self):
        self.depth -= 1
        self.lines.append("   " * self.depth + "EndProperty")

    def font(self, name="MS Sans Serif", size="8.25", bold=False, italic=False):
        self.bprop("Font", G_FONT)
        self.p("Name", name)
        self.p("Size", Raw(size))
        self.p("Charset", 0)
        self.p("Weight", 700 if bold else 400)
        self.p("Underline", False)
        self.p("Italic", italic)
        self.p("Strikethrough", False)
        self.eprop()

    def text(self, name, code):
        head = ["VERSION 5.00"] + [f'Object = "{g}"; "{f}"' for g, f in self.objects]
        attrs = [f'Attribute VB_Name = "{name}"', "Attribute VB_GlobalNameSpace = False",
                 "Attribute VB_Creatable = False", "Attribute VB_PredeclaredId = True",
                 "Attribute VB_Exposed = False"]
        return "\r\n".join(head + self.lines + attrs) + "\r\n" + _crlf(code)


def _crlf(code):
    code = code.strip("\n")
    return (code.replace("\r\n", "\n").replace("\n", "\r\n") + "\r\n") if code else ""


def _ext(twips):
    return int(round(twips * 2540 / 1440))


def _dialog_form(f, cls_name, caption, w, h, start=1):
    f.begin("VB.Form", cls_name)
    f.p("BorderStyle", Raw("3  'Fixed Dialog"))
    f.p("Caption", caption)
    f.props(ClientHeight=h, ClientLeft=45, ClientTop=330, ClientWidth=w, LinkTopic="Form1", MaxButton=False,
            MinButton=False, ScaleHeight=h, ScaleWidth=w, ShowInTaskbar=False)
    f.p("StartUpPosition", Raw("1  'CenterOwner" if start == 1 else "2  'CenterScreen"))


def _button(f, name, caption, left, top, w=1215, h=375, tab=0, default=False, cancel=False):
    f.begin("VB.CommandButton", name)
    if cancel:
        f.p("Cancel", True)
    f.p("Caption", caption)
    if default:
        f.p("Default", True)
    f.props(Height=h, Left=left, TabIndex=tab, Top=top, Width=w)
    f.end()


def _label(f, name, caption, left, top, w, h, tab=0, index=None, align=None, autosize=False, font=None,
           tag=None, back=None, fore=None, border=False):
    f.begin("VB.Label", name)
    if align:
        f.p("Alignment", Raw(align))
    if autosize:
        f.p("AutoSize", True)
    if back is not None:
        f.p("BackColor", Raw(back))
    if border:
        f.p("BorderStyle", Raw("1  'Fixed Single"))
    f.p("Caption", caption)
    if font:
        f.font(**font)
    if fore is not None:
        f.p("ForeColor", Raw(fore))
    f.p("Height", h)
    if index is not None:
        f.p("Index", index)
    f.props(Left=left, TabIndex=tab)
    f.p("Tag", tag)
    f.props(Top=top, Width=w)
    f.end()


# ------------------------------------------------------------------------------------------------ the result
class GenModule:
    def __init__(self, kind, name, text):
        self.kind, self.name, self.text = kind, name, text       # kind: Form, Module

    @property
    def filename(self):
        return self.name + (".bas" if self.kind == "Module" else ".frm")

    def pictures(self):
        return sorted(set(re.findall(r'"(@pic:[^"]+)"', self.text)))


class GenResult:
    def __init__(self):
        self.modules = []          # GenModule, in the order of the .vbp
        self.objects = []          # (guid, file) for the Object= lines
        self.strings = {}          # id -> text of the String Table (resources)
        self.startup = "Sub Main"
        self.notes = []

    def add(self, kind, name, text, objects=()):
        self.modules.append(GenModule(kind, name, text))
        for o in objects:
            if o not in self.objects:
                self.objects.append(o)

    def module(self, name):
        return next((m for m in self.modules if m.name.lower() == name.lower()), None)

    def object_lines(self):
        return [f"{g}; {f}" for g, f in self.objects]


class _Res:
    """the numbers that the captions get with the Resources option (from 1000, in order: VB6 oracle aw7)"""
    def __init__(self, on, start=1000):
        self.on, self.next, self.strings = on, start, {}

    def __call__(self, text):
        if not self.on:
            return text
        n = self.next
        self.next += 1
        self.strings[n] = text
        return str(n)


def _todo(what, indent="    "):
    return (f"{indent}'TODO: write the code of {what}.\n"
            f"{indent}MsgBox \"There is no code yet for {what}.\", vbInformation, App.Title\n")


def _sub(name, body, args=""):
    return f"Private Sub {name}({args})\n{body}End Sub\n"


# ------------------------------------------------------------------------------------------------ the main form
def _menu_name(menu, item=None):
    return "mnu" + menu.key + ("" if item is None else item.key)


def _write_menus(f, opts, res):
    mdi = opts.app_type == APP_MDI
    for m in opts.menus:
        if not m.on:
            continue
        f.begin("VB.Menu", _menu_name(m))
        f.p("Caption", res(m.caption))
        if m.key == "Window" and mdi:
            f.p("WindowList", True)
        bar = 0
        for it in m.items:
            if not it.on:
                continue
            if it.separator:
                f.begin("VB.Menu", f"{_menu_name(m)}Bar{bar}")
                f.p("Caption", "-")
                f.end()
                bar += 1
            elif it.key == MRU:
                for i in (1, 2, 3):
                    f.begin("VB.Menu", f"{_menu_name(m)}MRU")
                    f.p("Caption", "")
                    f.p("Index", i)
                    f.p("Visible", False)
                    f.end()
                f.begin("VB.Menu", f"{_menu_name(m)}Bar{bar}")      # shown with the first file of the list
                f.p("Caption", "-")
                f.p("Visible", False)
                f.end()
                bar += 1
            elif it.key == "ListViewMode":
                for i, (cap, _k) in enumerate(LISTVIEW_MODES):
                    f.begin("VB.Menu", "mnuListViewMode")
                    f.p("Caption", res(cap))
                    f.p("Index", i)
                    f.end()
            else:
                f.begin("VB.Menu", _menu_name(m, it))
                f.p("Caption", res(it.caption))
                if m.key in ("View",) and it.key in ("Toolbar", "StatusBar"):
                    f.p("Checked", True)
                sc = SHORTCUTS.get(m.key + it.key)
                if sc:
                    f.p("Shortcut", Raw(sc))
                f.end()
        f.end()


def _write_toolbar(f, opts, res, name="tbToolBar", images="imlToolbarIcons", width=4560):
    f.obj(OBJ_COMCTL)
    f.begin("MSComctlLib.Toolbar", name)
    f.p("Align", Raw("1  'Align Top"))
    f.props(Height=420, Left=0, TabIndex=1, Top=0, Width=width, _ExtentX=_ext(width), _ExtentY=_ext(420),
            ButtonWidth=609, Appearance=1, ImageList=images, _Version=393216)
    btns = opts.toolbar
    if btns:
        f.bprop("Buttons", G_BUTTONS)
        f.p("NumButtons", len(btns))
        for i, b in enumerate(btns, 1):
            f.bprop(f"Button{i}", G_BUTTON)
            if b.separator:
                f.p("Style", 3)
            else:
                f.p("Key", b.key)
                f.p("Object.ToolTipText", res(b.tip))
                f.p("ImageKey", b.key)
                if b.key in GROUP_BUTTONS:
                    f.p("Style", 2)
            f.eprop()
        f.eprop()
    f.end()


def _write_imagelist(f, name, keys, size=16, left=1680, top=1275):
    f.obj(OBJ_COMCTL)
    f.begin("MSComctlLib.ImageList", name)
    f.props(Left=left, Top=top, _ExtentX=1005, _ExtentY=1005, BackColor=-2147483643, ImageWidth=size,
            ImageHeight=size, MaskColor=12632256, _Version=393216)
    if keys:
        f.bprop("Images", G_IMAGES)
        f.p("NumListImages", len(keys))
        for i, k in enumerate(keys, 1):
            f.bprop(f"ListImage{i}", G_IMAGE)
            f.p("Picture", pic(k, size))
            f.p("Key", k)
            f.eprop()
        f.eprop()
    f.end()


def _write_statusbar(f, top, width=4560):
    f.obj(OBJ_COMCTL)
    f.begin("MSComctlLib.StatusBar", "sbStatusBar")
    f.p("Align", Raw("2  'Align Bottom"))
    f.props(Height=270, Left=0, TabIndex=0, Top=top, Width=width, _ExtentX=_ext(width), _ExtentY=476,
            _Version=393216)
    f.bprop("Panels", G_PANELS)
    f.p("NumPanels", 3)
    f.bprop("Panel1", G_PANEL)
    f.props(AutoSize=1, Text="Status", TextSave="Status")
    f.eprop()
    f.bprop("Panel2", G_PANEL)
    f.props(Style=6, AutoSize=2)
    f.eprop()
    f.bprop("Panel3", G_PANEL)
    f.props(Style=5, AutoSize=2)
    f.eprop()
    f.eprop()
    f.end()


def _write_comdlg(f):
    f.obj(OBJ_COMDLG)
    f.begin("MSComDlg.CommonDialog", "dlgCommonDialog")
    f.props(Left=1680, Top=1275, _ExtentX=847, _ExtentY=847, _Version=393216)
    f.end()


def _image_keys(buttons):
    out = []
    for b in buttons:
        if not b.separator and b.key not in out:
            out.append(b.key)
    return out


def _has(opts, menu, item):
    for m in opts.menus:
        if m.on and m.key == menu:
            return any(i.on and i.key == item for i in m.items)
    return False


TB_TO_MENU = {"New": ("File", "New"), "Open": ("File", "Open"), "Save": ("File", "Save"),
              "Print": ("File", "Print"), "Cut": ("Edit", "Cut"), "Copy": ("Edit", "Copy"),
              "Paste": ("Edit", "Paste"), "Delete": ("File", "Delete"), "Properties": ("File", "Properties"),
              "Undo": ("Edit", "Undo"), "Find": ("File", "Find"), "Help": ("Help", "Contents")}
LV_VIEWS = {"View Large Icons": 0, "View Small Icons": 1, "View List": 2, "View Details": 3}
LV_CONSTS = ["lvwIcon", "lvwSmallIcon", "lvwList", "lvwReport"]
RTF_FORMAT = {"Bold": "SelBold", "Italic": "SelItalic", "Underline": "SelUnderline"}
RTF_ALIGN = {"Align Left": "rtfLeft", "Center": "rtfCenter", "Align Right": "rtfRight"}


def _toolbar_code(opts, ctl="tbToolBar"):
    mdi, exp = opts.app_type == APP_MDI, opts.app_type == APP_EXPLORER
    cases = []
    for b in opts.toolbar:
        if b.separator:
            continue
        body = None
        mk = TB_TO_MENU.get(b.key)
        if mk and _has(opts, *mk):
            body = f"            mnu{mk[0]}{mk[1]}_Click\n"
        elif mdi and b.key in RTF_FORMAT:
            p = RTF_FORMAT[b.key]
            body = (f"            If Not ActiveForm Is Nothing Then\n"
                    f"                ActiveForm.rtfText.{p} = Not ActiveForm.rtfText.{p}\n"
                    f"            End If\n")
        elif mdi and b.key in RTF_ALIGN:
            body = (f"            If Not ActiveForm Is Nothing Then ActiveForm.rtfText.SelAlignment = "
                    f"{RTF_ALIGN[b.key]}\n")
        elif exp and b.key in LV_VIEWS:
            body = f"            SetViewMode {LV_VIEWS[b.key]}\n"
        if body is None:
            body = _todo(f"the '{b.key}' button", " " * 12)
        cases.append(f'        Case "{b.key}"\n{body}')
    if not cases:
        return ""
    return (f"Private Sub {ctl}_ButtonClick(ByVal Button As MSComctlLib.Button)\n"
            "    On Error Resume Next\n"
            "    Select Case Button.Key\n" + "".join(cases) + "    End Select\nEnd Sub\n")


FILE_FILTER = "All Files (*.*)|*.*"


def _ask_file(title, save, var="sFile", indent="    "):
    lines = ["With dlgCommonDialog",
             f"    .DialogTitle = \"{title}\"",
             "    .CancelError = False",
             "    .FileName = \"\"",
             f"    .Filter = \"{FILE_FILTER}\"    'TODO: the kinds of files of the program",
             f"    .{'ShowSave' if save else 'ShowOpen'}",
             "    If Len(.FileName) = 0 Then Exit Sub        'Cancel",
             f"    {var} = .FileName",
             "End With"]
    return "".join(indent + ln + "\n" for ln in lines)


def _menu_code(opts, m, it):
    """the body of mnu<Menu><Item>_Click (None: a TODO)"""
    t, k = opts.app_type, m.key + it.key
    mdi, exp = t == APP_MDI, t == APP_EXPLORER
    if k == "FileNew" and mdi:
        return "    LoadNewDoc\n"
    if k == "FileOpen":
        if mdi:
            return ("    Dim sFile As String\n" + _ask_file("Open", False) +
                    "    If ActiveForm Is Nothing Then LoadNewDoc\n"
                    "    ActiveForm.rtfText.LoadFile sFile\n"
                    "    ActiveForm.Tag = sFile                     'the file of the document\n"
                    "    ActiveForm.Caption = sFile\n")
        return ("    Dim sFile As String\n" + _ask_file("Open", False) +
                "    'TODO: open the file sFile.\n")
    if k == "FileClose":
        if mdi:
            return "    If Not ActiveForm Is Nothing Then Unload ActiveForm\n"
        if exp:
            return "    Unload Me\n"
    if k in ("FileSave", "FileSaveAs") and mdi:
        head = "    Dim sFile As String\n    If ActiveForm Is Nothing Then Exit Sub\n"
        ask = _ask_file("Save As" if k == "FileSaveAs" else "Save", True)
        if k == "FileSave":
            return (head + "    If ActiveForm.Tag = \"\" Then          'a new document: ask for its name\n"
                    + _ask_file("Save", True, indent=" " * 8) +
                    "        ActiveForm.Tag = sFile\n"
                    "        ActiveForm.Caption = sFile\n"
                    "    End If\n"
                    "    ActiveForm.rtfText.SaveFile ActiveForm.Tag\n")
        return (head + ask + "    ActiveForm.Tag = sFile\n    ActiveForm.Caption = sFile\n"
                "    ActiveForm.rtfText.SaveFile sFile\n")
    if k in ("FileSave", "FileSaveAs"):
        return ("    Dim sFile As String\n" + _ask_file("Save As" if k == "FileSaveAs" else "Save", True) +
                "    'TODO: save the file sFile.\n")
    if k == "FilePageSetup":
        return ("    On Error Resume Next\n    With dlgCommonDialog\n        .DialogTitle = \"Page Setup\"\n"
                "        .CancelError = True\n        .ShowPrinter\n    End With\n")
    if k == "FilePrint":
        body = ("    On Error Resume Next\n    With dlgCommonDialog\n        .DialogTitle = \"Print\"\n"
                "        .CancelError = True\n        .ShowPrinter\n        If Err.Number <> 0 Then Exit Sub"
                "        'Cancel\n    End With\n")
        if mdi:
            return ("    If ActiveForm Is Nothing Then Exit Sub\n" + body +
                    "    Printer.Print ActiveForm.rtfText.Text\n    Printer.EndDoc\n")
        return body + "    'TODO: print the document with Printer.Print ... Printer.EndDoc.\n"
    if k == "FileExit":
        return "    Unload Me\n"
    if k in ("EditCut", "EditCopy", "EditPaste") and mdi:
        out = "    On Error Resume Next\n    If ActiveForm Is Nothing Then Exit Sub\n"
        if k == "EditPaste":
            return out + "    ActiveForm.rtfText.SelText = Clipboard.GetText\n"
        out += "    Clipboard.Clear\n    Clipboard.SetText ActiveForm.rtfText.SelText\n"
        if k == "EditCut":
            out += "    ActiveForm.rtfText.SelText = \"\"\n"
        return out
    if k in ("EditSelectAll", "EditInvertSelection") and exp:
        expr = "True" if k == "EditSelectAll" else "Not lvListView.ListItems(i).Selected"
        return ("    Dim i As Long\n    For i = 1 To lvListView.ListItems.Count\n"
                f"        lvListView.ListItems(i).Selected = {expr}\n    Next\n")
    if k in ("ViewToolbar", "ViewStatusBar"):
        c = "tbToolBar" if k == "ViewToolbar" else "sbStatusBar"
        out = f"    {_menu_name(m, it)}.Checked = Not {_menu_name(m, it)}.Checked\n"
        if opts.toolbar or c == "sbStatusBar":
            out += f"    {c}.Visible = {_menu_name(m, it)}.Checked\n"
        if exp:
            out += "    SizeControls imgSplitter.Left\n"
        return out
    if it.key == "Options" and "Options" in opts.forms:
        return "    frmOptions.Show vbModal, Me\n"
    if k == "ViewWebBrowser" and opts.internet:
        return ("    Dim frmB As New frmBrowser\n"
                f"    frmB.StartingAddress = \"{opts.url}\"\n"
                "    frmB.Show\n")
    if m.key == "Window" and mdi:
        arr = {"Cascade": "vbCascade", "TileHorizontal": "vbTileHorizontal", "TileVertical": "vbTileVertical",
               "ArrangeIcons": "vbArrangeIcons"}
        if it.key in arr:
            return f"    Me.Arrange {arr[it.key]}\n"
        if it.key == "NewWindow":
            return "    LoadNewDoc\n"
    if k in ("HelpContents", "HelpSearchForHelpOn"):
        return "    ShowHelp\n"
    if k == "HelpAbout":
        if "About" in opts.forms:
            return "    frmAbout.Show vbModal, Me\n"
        return ("    MsgBox App.Title & vbCrLf & \"Version \" & App.Major & \".\" & App.Minor & \".\" & "
                "App.Revision, vbInformation, \"About\"\n")
    return None


HELP_CODE = '''Private Sub ShowHelp()
    'the help file of the program: Project > Properties > Help File Name
    If Len(App.HelpFile) = 0 Then
        MsgBox "This program has no help file.", vbInformation, Me.Caption
        Exit Sub
    End If
    On Error Resume Next
    If Environ("OS") = "Windows_NT" Then
        Shell "explorer.exe """ & App.HelpFile & """", vbNormalFocus
    Else
        Shell "xdg-open """ & App.HelpFile & """", vbNormalFocus
    End If
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, Me.Caption
End Sub
'''

PLACE_LOAD = """    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
"""
PLACE_SAVE = """    If Me.WindowState <> vbMinimized Then         'the place of the window for the next time
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
"""


def gen_main(opts, res):
    t = opts.app_type
    mdi, exp = t == APP_MDI, t == APP_EXPLORER
    f = Frm()
    w, h = 4560, 3030
    menus = Frm()                    # first: with Resources the menus take the first numbers (aw7: 1000 &File)
    menus.depth = 1
    _write_menus(menus, opts, res)
    if mdi:
        f.begin("VB.MDIForm", "frmMain")
        f.p("BackColor", Raw("&H8000000C&"))
        f.p("Caption", opts.name)
        f.props(ClientHeight=h, ClientLeft=120, ClientTop=750, ClientWidth=w, LinkTopic="MDIForm1")
    else:
        f.begin("VB.Form", "frmMain")
        f.p("Caption", opts.name)
        f.props(ClientHeight=h, ClientLeft=120, ClientTop=750, ClientWidth=w, LinkTopic="Form1",
                ScaleHeight=h, ScaleWidth=w)
    f.p("StartUpPosition", Raw("3  'Windows Default"))
    if exp:
        f.begin("VB.PictureBox", "picSplitter")
        f.props(BackColor=Raw("&H00808080&"), BorderStyle=Raw("0  'None"), FillColor=Raw("&H00808080&"),
                Height=4800, Left=5400, ScaleHeight=4800, ScaleWidth=72, TabIndex=6, Top=705,
                Visible=False, Width=72)
        f.end()
        f.begin("MSComctlLib.TreeView", "tvTreeView")
        f.props(Height=4800, Left=0, TabIndex=5, Top=705, Width=2016, _ExtentX=_ext(2016), _ExtentY=_ext(4800),
                _Version=393217, LineStyle=1, Style=7, Appearance=1)
        f.end()
        f.begin("MSComctlLib.ListView", "lvListView")
        f.props(Height=4800, Left=2052, TabIndex=4, Top=705, Width=3216, _ExtentX=_ext(3216),
                _ExtentY=_ext(4800), LabelWrap=True, HideSelection=True, _Version=393217,
                ForeColor=-2147483640, BackColor=-2147483643, BorderStyle=1, Appearance=1, NumItems=0)
        f.end()
        f.begin("VB.PictureBox", "picTitles")
        f.p("Align", Raw("1  'Align Top"))
        f.props(Appearance=Raw("0  'Flat"), BorderStyle=Raw("0  'None"), ForeColor=Raw("&H80000008&"),
                Height=300, Left=0, ScaleHeight=300, ScaleWidth=w, TabIndex=2, TabStop=False, Top=420, Width=w)
        _label(f, "lblTitle", " ListView:", 2078, 12, 3216, 270, 3, index=1, border=True,
               tag=res(" ListView:") if res.on else None)
        _label(f, "lblTitle", " TreeView:", 0, 12, 2016, 270, 3, index=0, border=True,
               tag=res(" TreeView:") if res.on else None)
        f.end()
    if opts.toolbar:
        _write_toolbar(f, opts, res, width=w)
    _write_statusbar(f, h - 270, w)
    _write_comdlg(f)
    keys = _image_keys(opts.toolbar)
    if opts.toolbar:
        _write_imagelist(f, "imlToolbarIcons", keys)
    if exp:
        f.begin("VB.Image", "imgSplitter")
        f.props(Height=4788, Left=1965, MousePointer=Raw("9  'Size W E"), Top=705, Width=150)
        f.end()
    f.lines += menus.lines
    f.end()

    code = []
    if exp:
        code.append("Option Explicit\n\nDim mbMoving As Boolean               'the splitter is dragged\n"
                    "Const SPLIT_LIMIT = 500                'the narrowest tree or list\n")
    form = "MDIForm" if mdi else "Form"
    load = ("    LoadResStrings Me\n" if opts.resources else "") + PLACE_LOAD
    if exp:
        load += ("    SetViewMode Val(GetSetting(App.Title, \"Settings\", \"ViewMode\", \"0\"))\n")
    if mdi:
        load += "    LoadNewDoc\n"
    code.append(_sub(f"{form}_Load", load))
    if mdi:
        code.append("Private Sub LoadNewDoc()\n"
                    "    'a new document window, Document 1, Document 2, ...\n"
                    "    Static nDocuments As Long\n"
                    "    Dim frmD As frmDocument\n"
                    "    nDocuments = nDocuments + 1\n"
                    "    Set frmD = New frmDocument\n"
                    "    frmD.Caption = \"Document \" & nDocuments\n"
                    "    frmD.Show\n"
                    "End Sub\n")
    unload = ""
    if not mdi:
        unload += ("    Dim i As Integer\n"
                   "    For i = Forms.Count - 1 To 1 Step -1        'the other forms of the program first\n"
                   "        Unload Forms(i)\n"
                   "    Next\n")
    unload += PLACE_SAVE
    if exp:
        unload += "    SaveSetting App.Title, \"Settings\", \"ViewMode\", lvListView.View\n"
    code.append(_sub(f"{form}_Unload", unload, "Cancel As Integer"))
    if exp:
        code.append(EXPLORER_CODE)
    tb = _toolbar_code(opts)
    if tb:
        code.append(tb)
    help_needed = False
    for m in opts.menus:
        if not m.on:
            continue
        for it in m.items:
            if not it.on or it.separator or it.key == MRU:
                continue
            if it.key == "ListViewMode":
                if exp:
                    code.append("Private Sub mnuListViewMode_Click(Index As Integer)\n"
                                "    SetViewMode Index\nEnd Sub\n")
                continue
            body = _menu_code(opts, m, it)
            if body is None:
                body = _todo(f"'{_menu_name(m, it)}_Click'")
            if body == "    ShowHelp\n":
                help_needed = True
            code.append(_sub(f"{_menu_name(m, it)}_Click", body))
    if help_needed:
        code.append(HELP_CODE)
    if exp:
        code.append(_view_mode_code(opts))
    objs = list(f.objects)
    return f.text("frmMain", "\n".join(code)), objs


EXPLORER_CODE = """Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    SizeControls imgSplitter.Left
End Sub

'the bar between the tree and the list: drag it to share the width
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSplitter.Move imgSplitter.Left, imgSplitter.Top, imgSplitter.Width \\ 2, imgSplitter.Height - 20
    picSplitter.Visible = True
    mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    If Not mbMoving Then Exit Sub
    sglPos = X + imgSplitter.Left
    If sglPos < SPLIT_LIMIT Then sglPos = SPLIT_LIMIT
    If sglPos > Me.ScaleWidth - SPLIT_LIMIT Then sglPos = Me.ScaleWidth - SPLIT_LIMIT
    picSplitter.Left = sglPos
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub SizeControls(ByVal X As Single)
    'the tree on the left of X, the list on the right, between the toolbar and the status bar
    Dim sglTop As Single, sglHeight As Single
    On Error Resume Next
    If X < 1500 Then X = 1500
    If X > Me.ScaleWidth - 1500 Then X = Me.ScaleWidth - 1500
    sglTop = picTitles.Top + picTitles.Height
    sglHeight = Me.ScaleHeight - sglTop
    If sbStatusBar.Visible Then sglHeight = sglHeight - sbStatusBar.Height
    If sglHeight < 0 Then sglHeight = 0
    tvTreeView.Move 0, sglTop, X, sglHeight
    imgSplitter.Move X, sglTop, imgSplitter.Width, sglHeight
    lvListView.Move X + 40, sglTop, Me.ScaleWidth - X - 40, sglHeight
    lblTitle(0).Width = tvTreeView.Width
    lblTitle(1).Left = lvListView.Left + 20
    lblTitle(1).Width = lvListView.Width - 40
End Sub
"""


def _view_mode_code(opts):
    has_tb = any(b.key in LV_VIEWS for b in opts.toolbar)
    has_menu = _has(opts, "View", "ListViewMode")
    lines = ["Private Sub SetViewMode(ByVal nMode As Integer)",
             "    'Large Icons, Small Icons, List, Details: the list, its menu and its buttons",
             "    Dim i As Integer",
             "    On Error Resume Next",
             "    If nMode < 0 Or nMode > 3 Then nMode = 0",
             "    lvListView.View = nMode"]
    if has_menu:
        lines += ["    For i = 0 To 3", "        mnuListViewMode(i).Checked = (i = nMode)", "    Next"]
    if has_tb:
        lines += ["    tbToolBar.Buttons(Choose(nMode + 1, LISTVIEW_MODE0, LISTVIEW_MODE1, LISTVIEW_MODE2, "
                  "LISTVIEW_MODE3)).Value = tbrPressed"]
    lines.append("End Sub")
    return "\n".join(lines) + "\n"


# ------------------------------------------------------------------------------------------------ Module1
LOADRES_CODE = """
Public Sub LoadResStrings(frm As Form)
    'the captions and tips of a form are numbers of the String Table of the .res file: their strings
    Dim ctl As Control
    Dim obj As Object
    On Error Resume Next
    If Val(frm.Tag) > 0 Then frm.Caption = LoadResString(Val(frm.Tag))
    For Each ctl In frm.Controls
        Select Case TypeName(ctl)
            Case "Menu"
                If Val(ctl.Caption) > 0 Then ctl.Caption = LoadResString(Val(ctl.Caption))
            Case "Toolbar"
                For Each obj In ctl.Buttons
                    If Val(obj.ToolTipText) > 0 Then obj.ToolTipText = LoadResString(Val(obj.ToolTipText))
                Next
            Case "TabStrip"
                For Each obj In ctl.Tabs
                    If Val(obj.Tag) > 0 Then obj.Caption = LoadResString(Val(obj.Tag))
                Next
            Case Else
                If Val(ctl.Tag) > 0 Then ctl.Caption = LoadResString(Val(ctl.Tag))
                If Val(ctl.ToolTipText) > 0 Then ctl.ToolTipText = LoadResString(Val(ctl.ToolTipText))
        End Select
    Next
End Sub
"""


def gen_module1(opts):
    lines = ['Attribute VB_Name = "Module1"']
    if opts.app_type == APP_EXPLORER:
        lines += ['Public Const LISTVIEW_MODE0 = "View Large Icons"',
                  'Public Const LISTVIEW_MODE1 = "View Small Icons"',
                  'Public Const LISTVIEW_MODE2 = "View List"',
                  'Public Const LISTVIEW_MODE3 = "View Details"']
    lines += ["Public fMainForm As frmMain", "", "Sub Main()"]
    if "Splash" in opts.forms:
        lines += ["    Dim sglStart As Single"]
    if "Login" in opts.forms:
        lines += ["    Dim fLogin As New frmLogin",
                  "    fLogin.Show vbModal",
                  "    If Not fLogin.OK Then End                    'the user did not log in",
                  "    Unload fLogin"]
    if "Splash" in opts.forms:
        lines += ["    frmSplash.Show",
                  "    frmSplash.Refresh",
                  "    sglStart = Timer"]
    lines += ["    Set fMainForm = New frmMain",
              "    Load fMainForm"]
    if "Splash" in opts.forms:
        lines += ["    'the splash screen stays at least 2 seconds (a main form loads fast: it would only flash)",
                  "    Do While Timer - sglStart < 2 And Timer >= sglStart",
                  "        DoEvents",
                  "    Loop",
                  "    Unload frmSplash"]
    lines += ["    fMainForm.Show", "End Sub"]
    text = "\n".join(lines) + "\n"
    if opts.resources:
        text += LOADRES_CODE
    return _crlf(text)


# ------------------------------------------------------------------------------------------------ frmDocument
def gen_document(toolbar=True):
    f = Frm()
    f.obj(OBJ_RICHTX)
    f.begin("VB.Form", "frmDocument")
    f.p("Caption", "frmDocument")
    f.props(ClientHeight=3030, ClientLeft=120, ClientTop=450, ClientWidth=4560, LinkTopic="Form1",
            MDIChild=True, ScaleHeight=3030, ScaleWidth=4560)
    f.begin("RichTextLib.RichTextBox", "rtfText")
    f.props(Height=2000, Left=100, TabIndex=0, Top=100, Width=3000, _ExtentX=_ext(3000), _ExtentY=_ext(2000),
            _Version=393217, Enabled=True, ScrollBars=3)
    f.end()
    f.end()
    code = """Private Sub Form_Load()
    Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfText.RightMargin = rtfText.Width - 400
End Sub
"""
    if toolbar:
        code += """
Private Sub rtfText_SelChange()
    'the buttons of the toolbar show the format of the selection
    On Error Resume Next
    With fMainForm.tbToolBar.Buttons
        .Item("Bold").Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
        .Item("Italic").Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
        .Item("Underline").Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        .Item("Align Left").Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
        .Item("Center").Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
        .Item("Align Right").Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
    End With
End Sub
"""
    return f.text("frmDocument", code), list(f.objects)


# ------------------------------------------------------------------------------------------------ frmBrowser
def gen_browser(name, res, load_res=False):
    """the Web Browser form: of the Internet Connectivity page (StartingAddress) or the template"""
    f = Frm()
    f.obj(OBJ_COMCTL)
    f.obj(OBJ_WEB)
    w, h = 6420, 4965
    tips = [res(k) for k in BROWSER_BUTTONS]            # the numbers of the tips, then of the label (aw7)
    addr = res("&Address:") if res.on else None
    f.begin("VB.Form", name)
    f.p("Caption", name)
    f.props(ClientHeight=h, ClientLeft=60, ClientTop=345, ClientWidth=w, LinkTopic="Form1", ScaleHeight=h,
            ScaleWidth=w)
    f.p("ShowInTaskbar", False)
    f.begin("VB.Timer", "timTimer")
    f.props(Enabled=False, Interval=5, Left=6180, Top=1500)
    f.end()
    f.begin("VB.PictureBox", "picAddress")
    f.p("Align", Raw("1  'Align Top"))
    f.props(BorderStyle=Raw("0  'None"), Height=675, Left=0, ScaleHeight=675, ScaleWidth=w, TabIndex=1,
            TabStop=False, Top=540, Width=w)
    f.begin("VB.ComboBox", "cboAddress")
    f.props(Height=315, Left=45, TabIndex=2, Top=300, Width=3795)
    f.end()
    _label(f, "lblAddress", "&Address:", 45, 60, 3075, 255, 3, tag=addr)
    f.end()
    f.begin("MSComctlLib.Toolbar", "tbToolBar")
    f.p("Align", Raw("1  'Align Top"))
    f.props(Height=540, Left=0, TabIndex=0, Top=0, Width=w, _ExtentX=_ext(w), _ExtentY=_ext(540), ButtonWidth=820,
            ButtonHeight=794, Appearance=1, ImageList="imlIcons", _Version=393216)
    f.bprop("Buttons", G_BUTTONS)
    f.p("NumButtons", len(BROWSER_BUTTONS))
    for i, k in enumerate(BROWSER_BUTTONS, 1):
        f.bprop(f"Button{i}", G_BUTTON)
        f.p("Key", k)
        f.p("Object.ToolTipText", tips[i - 1])
        f.p("ImageKey", k)
        f.eprop()
    f.eprop()
    f.end()
    _write_imagelist(f, "imlIcons", BROWSER_BUTTONS, 24, 2610, 2235)
    f.begin("SHDocVwCtl.WebBrowser", "brwWebBrowser")
    f.props(Height=3735, Left=50, TabIndex=4, Top=1215, Width=5400, ExtentX=_ext(5400), ExtentY=_ext(3735),
            ViewMode=0, Offline=0, Silent=0, RegisterAsBrowser=0, RegisterAsDropTarget=1, AutoArrange=False,
            NoClientEdge=False, AlignLeft=False, NoWebView=False, HideFileNames=False, SingleClick=False,
            SingleSelection=False, NoFolders=False, Transparent=False,
            ViewID="{0057D0E0-3573-11CF-AE69-08002B2E1262}", Location="")
    f.end()
    f.end()
    code = ["Option Explicit", "",
            "Public StartingAddress As String       'the first page: set it before Show",
            "Dim mbNoNavigate As Boolean           'the list of addresses changes: no navigation", "",
            "Private Sub Form_Load()",
            "    On Error Resume Next"]
    if load_res:
        code.append("    LoadResStrings Me")
    code += ["    Me.Show",
             "    Form_Resize",
             "    If Len(StartingAddress) > 0 Then",
             "        cboAddress.Text = StartingAddress",
             "        cboAddress.AddItem StartingAddress",
             "        timTimer.Enabled = True",
             "        brwWebBrowser.Navigate StartingAddress",
             "    End If",
             "End Sub", ""]
    code.append("""Private Sub Form_Resize()
    On Error Resume Next
    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Move 50, picAddress.Top + picAddress.Height, Me.ScaleWidth - 100, _
        Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub

Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    'the address of the page goes first in the list, once
    Dim i As Integer
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
    mbNoNavigate = True
    For i = cboAddress.ListCount - 1 To 0 Step -1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then cboAddress.RemoveItem i
    Next
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbNoNavigate = False
End Sub

Private Sub cboAddress_Click()
    If mbNoNavigate Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cboAddress_Click
    End If
End Sub

Private Sub timTimer_Timer()
    'the title while the page loads
    If brwWebBrowser.Busy Then
        Me.Caption = "Loading..."
    Else
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    timTimer.Enabled = True
    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select
End Sub
""")
    return f.text(name, "\n".join(code)), list(f.objects)


# ------------------------------------------------------------------------------------------------ the templates
def gen_splash(name, opts):
    f = Frm()
    w, h = 6900, 4200
    f.begin("VB.Form", name)
    f.p("BorderStyle", Raw("3  'Fixed Dialog"))
    f.props(ClientHeight=h, ClientLeft=255, ClientTop=1410, ClientWidth=w, ClipControls=False, ControlBox=False,
            KeyPreview=True, LinkTopic="Form2", MaxButton=False, MinButton=False, ScaleHeight=h, ScaleWidth=w,
            ShowInTaskbar=False)
    f.p("StartUpPosition", Raw("2  'CenterScreen"))
    f.p("Visible", False)
    f.begin("VB.Frame", "fraMainFrame")
    f.props(Height=h - 60, Left=45, TabIndex=0, Top=0, Width=w - 90)
    f.begin("VB.PictureBox", "picLogo")
    f.props(AutoSize=True, BorderStyle=Raw("0  'None"), Height=1440, Left=420, Picture=pic("Logo", 96),
            ScaleHeight=1440, ScaleWidth=1440, TabIndex=1, Top=900, Width=1440)
    f.end()
    big = dict(name="Arial", size="24", bold=True)
    mid = dict(name="Arial", size="12", bold=True)
    _label(f, "lblLicenseTo", "Licensed to: ", 240, 240, w - 600, 255, 2, align="1  'Right Justify",
           tag="LicenseTo")
    _label(f, "lblCompanyProduct", "Company Product", 2160, 840, 4200, 300, 3, font=mid, tag="CompanyProduct")
    _label(f, "lblProductName", "Product", 2160, 1200, 4200, 600, 4, font=big, tag="Product")
    _label(f, "lblPlatform", "Platform", 4200, 2100, 2160, 300, 5, align="1  'Right Justify", font=mid,
           tag="Platform")
    _label(f, "lblVersion", "Version", 4200, 2460, 2160, 255, 6, align="1  'Right Justify", tag="Version")
    _label(f, "lblCopyright", "Copyright", 2160, 3000, 4200, 255, 7, tag="Copyright")
    _label(f, "lblCompany", "Company", 2160, 3240, 4200, 255, 8, tag="Company")
    _label(f, "lblWarning", "This program is protected by copyright law.", 240, 3660, w - 600, 255, 9,
           tag="Warning")
    f.end()
    f.end()
    code = """Option Explicit

Private Sub Form_Load()
    lblProductName.Caption = App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    If Len(App.CompanyName) > 0 Then lblCompany.Caption = App.CompanyName
    If Len(App.LegalCopyright) > 0 Then lblCopyright.Caption = App.LegalCopyright
    If Len(App.ProductName) > 0 Then lblCompanyProduct.Caption = App.ProductName
    If Environ("OS") = "Windows_NT" Then
        lblPlatform.Caption = "for Windows"
    Else
        lblPlatform.Caption = "for Linux"
    End If
End Sub

'a key or a click closes the splash screen
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub fraMainFrame_Click()
    Unload Me
End Sub
"""
    return f.text(name, code), []


def gen_login(name, opts):
    f = Frm()
    w, h = 3750, 1545
    _dialog_form(f, name, "Login", w, h, 2)
    _label(f, "lblLabels", "&User Name:", 105, 150, 1080, 248, 0, index=0)
    f.begin("VB.TextBox", "txtUserName")
    f.props(Height=285, Left=1290, TabIndex=1, Top=135, Width=2325)
    f.end()
    _label(f, "lblLabels", "&Password:", 105, 540, 1080, 248, 2, index=1)
    f.begin("VB.TextBox", "txtPassword")
    f.props(Height=285, IMEMode=Raw("3  'DISABLE"), Left=1290, PasswordChar="*", TabIndex=3, Top=525,
            Width=2325)
    f.end()
    _button(f, "cmdOK", "OK", 495, 1020, 1140, 390, 4, default=True)
    _button(f, "cmdCancel", "Cancel", 2100, 1020, 1140, 390, 5, cancel=True)
    f.end()
    code = """Option Explicit

Public OK As Boolean         'True: the user logged in

Private Sub Form_Load()
    'the name of the user of the computer
    txtUserName.Text = Environ("USER")
    If Len(txtUserName.Text) = 0 Then txtUserName.Text = Environ("USERNAME")
End Sub

Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'TODO: check the user name and the password; this template takes an empty password
    If txtPassword.Text = "" Then
        OK = True
        Me.Hide
    Else
        MsgBox "The password is not correct. Please try again.", vbExclamation, "Login"
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    End If
End Sub
"""
    return f.text(name, code), []


def gen_options(name, opts):
    f = Frm()
    f.obj(OBJ_COMCTL)
    w, h = 6030, 4875
    _dialog_form(f, name, "Options", w, h)
    f.p("KeyPreview", True)
    _button(f, "cmdOK", "OK", 2490, 4455, 1095, 375, 1)
    _button(f, "cmdCancel", "Cancel", 3720, 4455, 1095, 375, 2, cancel=True)
    _button(f, "cmdApply", "&Apply", 4920, 4455, 1095, 375, 3)
    for i in (3, 2, 1, 0):
        f.begin("VB.PictureBox", "picOptions")
        f.props(BorderStyle=Raw("0  'None"), Height=3780, Index=i, Left=210 if i == 0 else -20000,
                ScaleHeight=3780, ScaleWidth=5685, TabIndex=4 + i, TabStop=False, Top=480, Width=5685)
        f.begin("VB.Frame", f"fraSample{i + 1}")
        f.props(Caption=f"Sample {i + 1}", Height=2025, Left=210, TabIndex=8 + i, Top=210, Width=2775)
        f.end()
        f.end()
    f.begin("MSComctlLib.TabStrip", "tbsOptions")
    f.props(Height=4245, Left=105, TabIndex=0, Top=120, Width=5895, _ExtentX=_ext(5895), _ExtentY=_ext(4245),
            _Version=393216)
    f.bprop("Tabs", G_TABS)
    f.p("NumTabs", 4)
    for i in range(1, 5):
        f.bprop(f"Tab{i}", G_TAB)
        f.p("Caption", f"Group {i}")
        f.p("ImageVarType", 2)
        f.eprop()
    f.eprop()
    f.end()
    f.end()
    code = """Option Explicit

Private Sub cmdApply_Click()
    'TODO: apply the options and keep the dialog open.
    MsgBox "There is no code yet for Apply.", vbInformation, Me.Caption
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'TODO: apply the options.
    MsgBox "There is no code yet for OK.", vbInformation, Me.Caption
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Ctrl+Tab: the next tab, Ctrl+Shift+Tab: the previous one
    Dim i As Integer, n As Integer
    If KeyCode <> vbKeyTab Or (Shift And vbCtrlMask) = 0 Then Exit Sub
    i = tbsOptions.SelectedItem.Index
    n = tbsOptions.Tabs.Count
    If (Shift And vbShiftMask) = 0 Then
        i = i Mod n + 1
    Else
        i = (i + n - 2) Mod n + 1
    End If
    Set tbsOptions.SelectedItem = tbsOptions.Tabs(i)
    KeyCode = 0
End Sub

Private Sub tbsOptions_Click()
    'the page of the selected tab is shown, the others are moved away
    Dim i As Integer
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
End Sub
"""
    return f.text(name, code), [OBJ_COMCTL]


def gen_about(name, opts):
    f = Frm()
    w, h = 5745, 3465
    _dialog_form(f, name, "About " + opts.name, w, h)
    f.p("ClipControls", False)
    f.begin("VB.PictureBox", "picIcon")
    f.props(AutoSize=True, BorderStyle=Raw("0  'None"), ClipControls=False, Height=480, Left=240,
            Picture=pic("AppIcon", 32), ScaleHeight=480, ScaleMode=Raw("0  'User"), ScaleWidth=480,
            TabIndex=1, Top=240, Width=480)
    f.end()
    _button(f, "cmdOK", "OK", 4245, 2625, 1467, 345, 0, default=True, cancel=True)
    _button(f, "cmdSysInfo", "&System Info...", 4260, 3075, 1452, 345, 2)
    _label(f, "lblTitle", "Application Title", 1050, 240, 4095, 480, 3,
           font=dict(name="MS Sans Serif", size="12", bold=True))
    _label(f, "lblVersion", "Version", 1050, 780, 4095, 225, 4)
    _label(f, "lblDescription", "Application Description", 1050, 1125, 4095, 1170, 5)
    # an etched line under the description (a gray line and a white one under it), not under a label
    f.begin("VB.Line", "Line1")
    f.props(BorderColor=Raw("&H00FFFFFF&"), Index=0, X1=225, X2=5450, Y1=2445, Y2=2445)
    f.end()
    f.begin("VB.Line", "Line1")
    f.props(BorderColor=Raw("&H00808080&"), BorderStyle=Raw("6  'Inside Solid"), Index=1, X1=225, X2=5450,
            Y1=2430, Y2=2430)
    f.end()
    _label(f, "lblDisclaimer", "Warning: this program is protected by copyright law.", 255, 2625, 3870, 825, 6)
    f.end()
    code = """Option Explicit

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblTitle.Caption = App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    If Len(App.FileDescription) > 0 Then lblDescription.Caption = App.FileDescription
    If Len(App.LegalCopyright) > 0 Then lblDisclaimer.Caption = App.LegalCopyright
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSysInfo_Click()
    'the system information program of the computer: msinfo32 on Windows, the first system monitor found on Linux
    Dim vProgram As Variant
    On Error Resume Next
    If Environ("OS") = "Windows_NT" Then
        Shell "msinfo32.exe", vbNormalFocus
        If Err.Number = 0 Then Exit Sub
    Else
        For Each vProgram In Array("gnome-system-monitor", "mate-system-monitor", "plasma-systemmonitor", _
                                   "ksysguard", "xfce4-taskmanager", "lxtask")
            Err.Clear
            Shell vProgram, vbNormalFocus
            If Err.Number = 0 Then Exit Sub
        Next
    End If
    MsgBox "System information is not available on this computer.", vbInformation, Me.Caption
End Sub
"""
    return f.text(name, code), []


def gen_dialog(name, opts):
    f = Frm()
    _dialog_form(f, name, "Dialog Caption", 6030, 3195)
    _button(f, "OKButton", "OK", 4680, 120, 1215, 375, 0, default=True)
    _button(f, "CancelButton", "Cancel", 4680, 600, 1215, 375, 1, cancel=True)
    f.end()
    code = """Option Explicit

Private Sub OKButton_Click()
    'TODO: what OK does.
    Unload Me
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub
"""
    return f.text(name, code), []


def gen_odbc(name, opts):
    f = Frm()
    w, h = 4470, 3180
    _dialog_form(f, name, "ODBC Logon", w, h)
    _button(f, "cmdOK", "&OK", 915, 2655, 1440, 450, 0, default=True)
    _button(f, "cmdCancel", "Cancel", 2520, 2655, 1440, 450, 1, cancel=True)
    f.begin("VB.Frame", "fraConnection")
    f.props(Caption="Connection Values", Height=2415, Left=120, TabIndex=2, Top=120, Width=4230)
    rows = [("&DSN:", "cboDSNList", "combo"), ("&UID:", "txtUID", "text"), ("&Password:", "txtPWD", "pwd"),
            ("Data&base:", "txtDatabase", "text"), ("Dri&ver:", "cboDrivers", "combo"),
            ("&Server:", "txtServer", "text")]
    tab = 3
    for i, (cap, ctl, kind) in enumerate(rows):
        top = 240 + i * 345
        _label(f, "lblConnection", cap, 135, top + 45, 900, 195, tab, index=i)
        tab += 1
        if kind == "combo":
            f.begin("VB.ComboBox", ctl)
            f.props(Enabled=False if ctl == "cboDrivers" else None, Height=315, Left=1125,
                    Style=Raw("2  'Dropdown List"), TabIndex=tab, Top=top, Width=3015)
        else:
            f.begin("VB.TextBox", ctl)
            f.props(Enabled=False if ctl == "txtServer" else None, Height=300, Left=1125,
                    PasswordChar="*" if kind == "pwd" else None, TabIndex=tab, Top=top, Width=3015)
        tab += 1
        f.end()
    f.end()
    f.end()
    code = '''Option Explicit

'the data sources (DSN) and drivers of ODBC: on Windows from odbc32.dll, on Linux from the files of unixODBC
Private Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv As Long) As Integer
Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal henv As Long, ByVal fDirection As Integer, _
    ByVal szDSN As String, ByVal cbDSNMax As Integer, pcbDSN As Integer, ByVal szDescription As String, _
    ByVal cbDescriptionMax As Integer, pcbDescription As Integer) As Integer
Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv As Long) As Integer
Private Const SQL_SUCCESS = 0
Private Const SQL_FETCH_NEXT = 1

Private Sub Form_Load()
    cboDSNList.AddItem "(None)"
    If Environ("OS") = "Windows_NT" Then
        If Environ("PROCESSOR_ARCHITECTURE") = "x86" Then ListWindowsDSNs     '32-bit programs
    Else
        ListIniSections Environ("HOME") & "/.odbc.ini", cboDSNList
        ListIniSections "/etc/odbc.ini", cboDSNList
        ListIniSections "/etc/odbcinst.ini", cboDrivers
    End If
    cboDSNList.ListIndex = 0
End Sub

Private Sub ListWindowsDSNs()
    Dim henv As Long, sDSN As String, sDriver As String, nDSN As Integer, nDriver As Integer
    On Error Resume Next
    If SQLAllocEnv(henv) <> SQL_SUCCESS Then Exit Sub
    Do
        sDSN = Space$(1024)
        sDriver = Space$(1024)
        If SQLDataSources(henv, SQL_FETCH_NEXT, sDSN, 1024, nDSN, sDriver, 1024, nDriver) <> SQL_SUCCESS Then Exit Do
        cboDSNList.AddItem Left$(sDSN, nDSN)
        AddOnce cboDrivers, Left$(sDriver, nDriver)
    Loop
    SQLFreeEnv henv
End Sub

Private Sub ListIniSections(ByVal sFile As String, cbo As ComboBox)
    'the [sections] of an .ini file of unixODBC: the names of the DSNs, or of the drivers
    Dim nFile As Integer, sLine As String
    On Error GoTo Done
    If Len(Dir$(sFile)) = 0 Then Exit Sub
    nFile = FreeFile
    Open sFile For Input As #nFile
    Do While Not EOF(nFile)
        Line Input #nFile, sLine
        sLine = Trim$(sLine)
        If Left$(sLine, 1) = "[" And Right$(sLine, 1) = "]" Then
            sLine = Mid$(sLine, 2, Len(sLine) - 2)
            If LCase$(sLine) <> "odbc data sources" And LCase$(sLine) <> "odbc" Then AddOnce cbo, sLine
        End If
    Loop
Done:
    If nFile <> 0 Then Close #nFile
End Sub

Private Sub AddOnce(cbo As ComboBox, ByVal s As String)
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = s Then Exit Sub
    Next
    cbo.AddItem s
End Sub

Private Sub cboDSNList_Click()
    'without a DSN the driver and the server are given here
    cboDrivers.Enabled = (cboDSNList.ListIndex <= 0)
    txtServer.Enabled = (cboDSNList.ListIndex <= 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'the connection string of the values; TODO: open the connection with it
    Dim sConnect As String
    If cboDSNList.ListIndex > 0 Then
        sConnect = "DSN=" & cboDSNList.Text & ";"
    Else
        sConnect = "Driver={" & cboDrivers.Text & "};Server=" & txtServer.Text & ";"
    End If
    sConnect = sConnect & "UID=" & txtUID.Text & ";PWD=" & txtPWD.Text & ";"
    If Len(txtDatabase.Text) > 0 Then sConnect = sConnect & "Database=" & txtDatabase.Text & ";"
    MsgBox "ADO:" & vbCrLf & "    cn.Open ""Provider=MSDASQL;" & sConnect & """" & vbCrLf & vbCrLf & _
        "DAO:" & vbCrLf & "    Set db = OpenDatabase("""", False, False, ""ODBC;" & sConnect & """)", _
        vbInformation, "Connection string"
End Sub
'''
    return f.text(name, code), []


def gen_tip(name, opts):
    f = Frm()
    w, h = 5415, 3285
    f.begin("VB.Form", name)
    f.p("Caption", "Tip of the Day")
    f.props(ClientHeight=h, ClientLeft=2370, ClientTop=2400, ClientWidth=w, LinkTopic="Form1", MaxButton=False,
            MinButton=False, ScaleHeight=h, ScaleWidth=w, ShowInTaskbar=False)
    f.p("StartUpPosition", Raw("1  'CenterOwner"))
    f.p("WhatsThisButton", True)
    f.begin("VB.CheckBox", "chkLoadTipsAtStartup")
    f.props(Caption="&Show Tips at Startup", Height=315, Left=120, TabIndex=3, Top=2940, Width=2055)
    f.end()
    _button(f, "cmdNextTip", "&Next Tip", 4080, 600, 1215, 375, 2)
    f.begin("VB.PictureBox", "picTip")
    f.props(BackColor=Raw("&H00FFFFFF&"), Height=2715, Left=120, Picture=pic("Tip", 32), ScaleHeight=2655,
            ScaleWidth=3675, TabIndex=1, Top=120, Width=3735)
    _label(f, "lblDidYouKnow", "Did you know...", 540, 180, 2655, 255, 4, back="&H00FFFFFF&",
           font=dict(name="MS Sans Serif", size="9.75", bold=True), fore="&H00000000&")
    _label(f, "lblTipText", "", 180, 840, 3255, 1635, 5, back="&H00FFFFFF&", fore="&H00000000&")
    f.end()
    _button(f, "cmdOK", "OK", 4080, 120, 1215, 375, 0, default=True, cancel=True)
    f.end()
    code = """Option Explicit

'the tips are the lines of TIPOFDAY.TXT, in the folder of the program
Const TIP_FILE = "TIPOFDAY.TXT"
Dim Tips As New Collection
Dim nTip As Long

Private Sub Form_Load()
    If Val(GetSetting(App.EXEName, "Options", "Show Tips at Startup", "1")) = 0 Then
        Unload Me
        Exit Sub
    End If
    chkLoadTipsAtStartup.Value = vbChecked
    Randomize
    If Not LoadTips(App.Path & IIf(InStr(App.Path, "/") > 0, "/", "\\") & TIP_FILE) Then
        lblTipText.Caption = "...that " & TIP_FILE & " is missing? Write the tips in a text file, " & _
            "one tip on each line, name it " & TIP_FILE & " and put it in the folder of the program."
    End If
End Sub

Private Function LoadTips(ByVal sFile As String) As Boolean
    Dim nFile As Integer, sTip As String
    On Error GoTo Failed
    If Len(Dir$(sFile)) = 0 Then Exit Function
    nFile = FreeFile
    Open sFile For Input As #nFile
    Do While Not EOF(nFile)
        Line Input #nFile, sTip
        If Len(Trim$(sTip)) > 0 Then Tips.Add sTip
    Loop
    Close #nFile
    If Tips.Count = 0 Then Exit Function
    ShowNextTip
    LoadTips = True
    Exit Function
Failed:
    LoadTips = False
End Function

Private Sub ShowNextTip()
    'a random tip, not the one shown now
    Dim n As Long
    If Tips.Count = 0 Then Exit Sub
    Do
        n = Int(Tips.Count * Rnd) + 1
    Loop While n = nTip And Tips.Count > 1
    nTip = n
    lblTipText.Caption = Tips(nTip)
End Sub

Private Sub chkLoadTipsAtStartup_Click()
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdNextTip_Click()
    ShowNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
"""
    return f.text(name, code), []


STANDARD_GEN = {"Splash": ("frmSplash", gen_splash), "Login": ("frmLogin", gen_login),
                "Options": ("frmOptions", gen_options), "About": ("frmAbout", gen_about)}
TEMPLATE_GEN = {"Dialog": ("Dialog", gen_dialog), "ODBC Log In": ("frmODBCLogon", gen_odbc),
                "Tip of the Day": ("frmTip", gen_tip)}


def gen_template(template, name, opts):
    """a form of Form Templates (also Project > Add Form): (text, objects)"""
    if template == "Web Browser":
        return gen_browser(name, _Res(False))
    return dict(STANDARD_GEN, **TEMPLATE_GEN)[template][1](name, opts)


# ------------------------------------------------------------------------------------------------ the project
def generate(opts):
    """everything the wizard adds to the new project -> GenResult"""
    r = GenResult()
    res = _Res(opts.resources)
    used = set()

    def unique(base):
        n, i = base, 1
        while n.lower() in used:
            n = f"{base}{i}"
            i += 1
        used.add(n.lower())
        return n

    used.update(["module1", "frmmain"])
    r.add("Module", "Module1", gen_module1(opts))
    text, objs = gen_main(opts, res)
    r.add("Form", "frmMain", text, objs)
    if opts.app_type == APP_MDI:
        used.add("frmdocument")
        text, objs = gen_document(bool(opts.toolbar))
        r.add("Form", "frmDocument", text, objs)
    if opts.internet:
        used.add("frmbrowser")
        text, objs = gen_browser("frmBrowser", res, load_res=opts.resources)
        r.add("Form", "frmBrowser", text, objs)
    for f in STANDARD_FORMS:
        if f in opts.forms:
            base, fn = STANDARD_GEN[f]
            name = unique(base)
            text, objs = fn(name, opts)
            r.add("Form", name, text, objs)
    for t in TEMPLATES:
        if t in opts.templates:
            if t == "Web Browser":
                name = unique("frmBrowser")
                text, objs = gen_browser(name, _Res(False))
            else:
                base, fn = TEMPLATE_GEN[t]
                name = unique(base)
                text, objs = fn(name, opts)
            r.add("Form", name, text, objs)
    # the .vbp lists the libraries in the order VB6 does: ComDlg32, mscomctl, richtx32, ieframe
    order = [OBJ_COMDLG, OBJ_COMCTL, OBJ_RICHTX, OBJ_WEB]
    r.objects.sort(key=lambda o: order.index(o) if o in order else len(order))
    r.strings = dict(res.strings)
    return r


def attach_pictures(sf, make):
    """the placeholders of the pictures of a parsed .frm -> FrxData of make(name, size) (bytes of a BMP)"""
    from lbvc.project.source import Property, PropertyBlock
    from lbvc.project.frx import FrxData

    def items(seq):
        for it in seq:
            if isinstance(it, PropertyBlock):
                yield from items(it.items)
            elif isinstance(it, Property):
                yield it
    n = 0
    if sf is None or sf.root is None:
        return 0
    for c in sf.root.walk():
        for p in items(c.items):
            if isinstance(p.value, str):
                m = _PIC_RE.match(p.value)
                if m:
                    p.value = FrxData("picture", make(m.group(1), int(m.group(2))))
                    n += 1
    return n


def build_project(result, opts, make_pic, encoding=None):
    """a new, unsaved VBProject with what the wizard generated (the Form1 of a new project goes): Module1,
    frmMain, the Object= lines of its libraries after frmMain (where VB6 writes them, aw1 Project1.vbp), the
    other forms, Startup "Sub Main". make_pic(name, size) -> the bytes of a BMP (tbicons.bmp). The .res of the
    Resources option is result_res(result) (the IDE saves it with the project)."""
    from .vbmodel import VBProject, Module
    from lbvc.project.source import parse_source
    p = VBProject.new(encoding, "Exe", opts.name)
    for m in list(p.modules):
        p.remove_module(m)
    for gm in result.modules:
        sf = parse_source(gm.text, None, gm.kind, True, p.encoding)
        sf.encoding = p.encoding
        attach_pictures(sf, make_pic)
        m = Module(p, gm.kind, gm.name, gm.filename, sf)
        m.dirty = True
        p.modules.append(m)
        p._vbp_add(m)
        if gm.name == "frmMain":
            ents = p.vbp.entries
            i = next(i for i, (k, v) in enumerate(ents) if v is not None and k.strip() == "Form"
                     and v.strip().lower() == "frmmain.frm")
            for j, line in enumerate(result.object_lines()):
                ents.insert(i + 1 + j, ["Object", line])
    p.vbp.set("Startup", result.startup, quote=True)
    p.dirty = True
    return p


def result_res(result, lang=0):
    """the String Table of the Resources option (language neutral, as VB6 writes it: aw7 Project1.RES)"""
    if not result.strings:
        return None
    from .resmodel import ResFile
    r = ResFile()
    for sid, text in sorted(result.strings.items()):
        r.set_string(lang, sid, text)
    return r


def summary(opts):
    """the Summary Report of the Finished! page"""
    lines = [f"Application Wizard - summary for {opts.name}", "",
             f"Interface type: {TYPE_NAMES[opts.app_type]}", "", "Menus:"]
    for m in opts.menus:
        if m.on:
            items = [i.caption.replace("&", "") for i in m.items if i.on and not i.separator]
            lines.append(f"    {m.caption.replace('&', '')}: " + ", ".join(items))
    tb = [b.key for b in opts.toolbar if not b.separator]
    lines += ["", "Toolbar: " + (", ".join(tb) if tb else "(none)"),
              "", f"Resource file: {'yes' if opts.resources else 'no'}",
              f"Web browser: {opts.url if opts.internet else 'no'}",
              "Standard forms: " + (", ".join(opts.forms) if opts.forms else "(none)"),
              "Form templates: " + (", ".join(opts.templates) if opts.templates else "(none)")]
    return "\r\n".join(lines) + "\r\n"


__all__ = ["APP_MDI", "APP_SDI", "APP_EXPLORER", "AppOptions", "Menu", "MenuItem", "TbButton", "generate",
           "gen_template", "attach_pictures", "build_project", "result_res", "summary", "parse_menu", "format_menu", "default_menus",
           "default_toolbar", "STANDARD_FORMS", "TEMPLATES"]
