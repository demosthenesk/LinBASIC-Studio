"""What the Property Page Wizard writes (phase 9.2e, VB6 oracle DEV-9/pp1, pp2); no Qt here.

The property pages of a user control are its PropertyPages property: a record of its .ctx (u32 count, then each
name as u16 length with its 0, the characters, 0) - "StandardFont", "StandardColor", "StandardPicture" are pages of
the system, the others PropertyPage modules of the project. A page the wizard makes has, for each property, a Label
and a TextBox (a CheckBox for a Boolean) and the code that moves the values between them and SelectedControls(0);
the properties of the user control on a page carry Attribute X.VB_ProcData.VB_Invoke_Property = "Page" (the page of
the property in the Properties window) and the user control Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes".

VB6 lays the controls of a page out in the reverse order of its list (DEV-9/pp2: Title, Active, Rating, BackColor
-> BackColor on top) and writes the code in the order of the list; LinBASIC Studio lays them out in the order of
the list and writes the code from the bottom of the page up: the same .pag as VB6's for the same page.
"""
import re
import struct

STANDARD = ["StandardFont", "StandardColor", "StandardPicture"]
# the type of the properties each page of the system shows
STANDARD_TYPES = {"StandardFont": {"font", "stdfont", "stdole.stdfont", "stdole.font"},
                  "StandardColor": {"ole_color", "stdole.ole_color"},
                  "StandardPicture": {"picture", "stdpicture", "stdole.picture", "stdole.stdpicture", "ipicturedisp"}}
BOOLEAN = {"boolean"}
# the layout of a page (twips, DEV-9/pp1 Options.pag)
PAGE_W, PAGE_H = 5925, 3495
LEFT, WIDTH, FIRST, STEP = 90, 2700, 120, 650
LABEL_H, TEXT_H, CHECK_H, TEXT_DY = 240, 330, 285, 250
WIZARD_KEY = "PropPageWizardRun"


# ---------------------------------------------------------------- the record of PropertyPages
def read_pages(data):
    """the names in the PropertyPages record of a .ctx"""
    if not data or len(data) < 4:
        return []
    n = struct.unpack_from("<I", data, 0)[0]
    out, p = [], 4
    for _ in range(min(n, 64)):
        if p + 2 > len(data):
            break
        ln = struct.unpack_from("<H", data, p)[0]
        p += 2
        raw = bytes(data[p:p + ln])
        p += ln
        out.append(raw.split(b"\0", 1)[0].decode("cp1252", "replace"))
    return out


def pages_blob(names, encoding="cp1252"):
    """the PropertyPages record of the names"""
    out = bytearray(struct.pack("<I", len(names)))
    for n in names:
        b = n.encode(encoding, "replace") + b"\0"
        out += struct.pack("<H", len(b)) + b
    return bytes(out)


# ---------------------------------------------------------------- the properties of a user control
_PROP_RE = re.compile(r"^\s*Public\s+Property\s+(Get|Let|Set)\s+([A-Za-z_]\w*)\s*\((.*?)\)\s*(?:As\s+([\w.]+))?",
                      re.IGNORECASE)


def uc_properties(code):
    """[(name, type)] of the properties of a user control the wizard can put on a page: a Public Property Get
    without arguments and a Let or Set, in the order of the code"""
    gets, writable, order = {}, set(), []
    for line in code.replace("\r\n", "\n").split("\n"):
        m = _PROP_RE.match(line)
        if not m:
            continue
        kind, name, args = m.group(1).lower(), m.group(2), m.group(3).strip()
        low = name.lower()
        if kind == "get" and not args:
            if low not in gets:
                order.append(low)
            gets[low] = (name, (m.group(4) or "Variant"))
        elif kind in ("let", "set"):
            writable.add(low)
    return [gets[k] for k in order if k in writable]


def standard_of(type_name):
    t = (type_name or "").lower()
    return next((p for p, ts in STANDARD_TYPES.items() if t in ts), None)


# ---------------------------------------------------------------- what a page the wizard made shows
_BIND_RE = re.compile(r"^\s*(txt|chk)([A-Za-z_]\w*)\.(Text|Value)\s*=.*SelectedControls\(0\)\.([A-Za-z_]\w*)",
                      re.IGNORECASE)


def page_properties(module):
    """the properties of a page (the lines of its SelectionChanged that take them from SelectedControls(0)), from
    the top of the page down (the place of their controls)"""
    code = module.code_text() if module is not None else ""
    found = []
    for line in code.replace("\r\n", "\n").split("\n"):
        m = _BIND_RE.match(line)
        if m and m.group(4).lower() not in {f.lower() for f, _ in found}:
            found.append((m.group(4), (m.group(1) + m.group(2)).lower()))
    tops = {}
    root = module.sf.root if module is not None and module.sf is not None else None
    if root is not None:
        for c in root.walk():
            if c is not root:
                try:
                    tops[c.name.lower()] = int(float(c.get("Top", 0)))
                except (TypeError, ValueError):
                    tops[c.name.lower()] = 0
    order = {name: i for i, (name, _) in enumerate(found)}
    found.sort(key=lambda f: (tops.get(f[1], 10 ** 9), order[f[0]]))
    return [name for name, _ in found]


# ---------------------------------------------------------------- the model of the wizard
class Page:
    def __init__(self, name, checked=True, existing=False, module=None):
        self.name = name
        self.checked = checked
        self.existing = existing                # a PropertyPage module of the project already
        self.module = module
        self.props = []                         # the properties on it (custom pages)
        self.old_props = []

    @property
    def standard(self):
        return self.name in STANDARD


class PPModel:
    """the settings of the wizard for one user control"""

    def __init__(self, uc_name, props, current=(), modules=None):
        """props: [(name, type)] of uc_properties; current: the pages of its PropertyPages; modules: the
        PropertyPage modules of the project by lower name"""
        self.uc = uc_name
        self.props = list(props)
        self.types = {n: t for n, t in self.props}
        modules = modules or {}
        self.pages = []
        for n in current:
            low = n.lower()
            if any(p.name.lower() == low for p in self.pages):
                continue
            m = modules.get(low)
            if n in STANDARD or m is not None:
                pg = Page(m.name if m is not None else n, True, m is not None, m)
                if m is not None:
                    pg.props = [x for x in page_properties(m) if x in self.types or self._find(x)]
                    pg.props = [self._find(x) or x for x in pg.props]
                    pg.old_props = list(pg.props)
                self.pages.append(pg)
        first = not self.pages
        for std in STANDARD:                    # the pages of the system the types of its properties ask for
            if std not in self.page_names() and any(standard_of(t) == std for _, t in self.props):
                self.pages.append(Page(std, checked=first))   # checked the first time (DEV-9/pp1: StandardFont)

    def _find(self, name):
        low = name.lower()
        return next((n for n in self.types if n.lower() == low), None)

    def page_names(self):
        return [p.name for p in self.pages]

    def page(self, name):
        low = name.lower()
        return next((p for p in self.pages if p.name.lower() == low), None)

    def custom_pages(self):
        return [p for p in self.pages if not p.standard]

    # ---- Select the Property Pages
    def new_name(self):
        i = 1
        while self.page(f"PropertyPage{i}") is not None:
            i += 1
        return f"PropertyPage{i}"

    def check_name(self, name, but=None):
        """None, or why the name cannot be a page"""
        if not re.match(r"^[A-Za-z][A-Za-z0-9_]{0,39}$", name or ""):
            return "Invalid page name."
        p = self.page(name)
        if p is not None and p is not but:
            return f"A page named '{name}' already exists."
        return None

    def add_page(self, name):
        pg = Page(name)
        self.pages.append(pg)
        return pg

    def rename_page(self, old, new):
        pg = self.page(old)
        if pg is not None and not pg.standard and not pg.existing:
            pg.name = new
        return pg

    def move_page(self, name, d):
        pg = self.page(name)
        i = self.pages.index(pg)
        j = max(0, min(len(self.pages) - 1, i + d))
        self.pages.insert(j, self.pages.pop(i))
        return j

    # ---- Add Properties
    def available(self):
        """the properties not on a page yet (those of the pages of the system left out), in the order of their
        names"""
        used = {x.lower() for p in self.custom_pages() for x in p.props}
        return sorted((n for n, t in self.props if standard_of(t) is None and n.lower() not in used), key=str.lower)

    def page_props(self, name):
        pg = self.page(name)
        if pg is None:
            return []
        if pg.standard:
            return [n for n, t in self.props if standard_of(t) == pg.name]
        return list(pg.props)

    def add_prop(self, page, prop):
        pg = self.page(page)
        if pg is not None and not pg.standard and prop in self.available():
            pg.props.append(prop)

    def remove_prop(self, page, prop):
        pg = self.page(page)
        if pg is not None and prop in pg.props:
            pg.props.remove(prop)

    def selected_pages(self):
        return [p for p in self.pages if p.checked]


# ---------------------------------------------------------------- a page
def _line(depth, name, value):
    return f"{'   ' * depth}{name:<16}=   {value}"


def _q(s):
    return '"' + s.replace('"', '""') + '"'


def page_text(name, props, types):
    """the .pag of a page the wizard makes (our own text in VB6's format): a Label and a TextBox for a property
    (a CheckBox for a Boolean), from the top down in the order of the list"""
    items = []                                  # (kind, prop, top), the order of the layout
    for k, p in enumerate(props):
        items.append(("check" if (types.get(p) or "").lower() in BOOLEAN else "text", p, FIRST + STEP * k))
    height = max(PAGE_H, FIRST + STEP * len(props) + FIRST)
    tab, ctls = 0, []                           # (windowed, lines) in the order they are made
    for kind, p, top in items:
        if kind == "check":
            ctls.append((True, [f"   Begin VB.CheckBox chk{p} ", _line(2, "Caption", _q(p)),
                                _line(2, "Height", CHECK_H), _line(2, "Left", LEFT), _line(2, "TabIndex", tab),
                                _line(2, "Top", top), _line(2, "Width", WIDTH), "   End"]))
            tab += 1
            continue
        ctls.append((False, [f"   Begin VB.Label lbl{p} ", _line(2, "Caption", _q(p + ":")),
                             _line(2, "Height", LABEL_H), _line(2, "Left", LEFT), _line(2, "TabIndex", tab),
                             _line(2, "Top", top), _line(2, "Width", WIDTH), "   End"]))
        ctls.append((True, [f"   Begin VB.TextBox txt{p} ", _line(2, "Height", TEXT_H), _line(2, "Left", LEFT),
                            _line(2, "TabIndex", tab + 1), _line(2, "Top", top + TEXT_DY),
                            _line(2, "Width", WIDTH), "   End"]))
        tab += 2
    # VB6 writes the newest control first, the windowed controls before the Labels
    body = [ln for w, lines in reversed(ctls) if w for ln in lines] + \
        [ln for w, lines in reversed(ctls) if not w for ln in lines]
    head = ["VERSION 5.00", f"Begin VB.PropertyPage {name} ", _line(1, "Caption", _q(name)),
            _line(1, "ClientHeight", height), _line(1, "ClientLeft", 0), _line(1, "ClientTop", 0),
            _line(1, "ClientWidth", PAGE_W), _line(1, "PaletteMode", "0  'Halftone"),
            _line(1, "ScaleHeight", height), _line(1, "ScaleWidth", PAGE_W)]
    attrs = [f'Attribute VB_Name = "{name}"', "Attribute VB_GlobalNameSpace = False",
             "Attribute VB_Creatable = False", "Attribute VB_PredeclaredId = False", "Attribute VB_Exposed = False"]
    return "\n".join(head + body + ["End"] + attrs + page_code(props, types)) + "\n"


def page_code(props, types):
    """the code of a page: Changed when a box changes, ApplyChanges, SelectionChanged - from the bottom of the
    page up, as VB6 writes them (DEV-9/pp1, pp2: the same file as VB6's for the same page)"""
    out = []
    props = list(reversed(props))
    boolean = {p for p in props if (types.get(p) or "").lower() in BOOLEAN}
    for p in props:
        if p in boolean:
            out += [f"Private Sub chk{p}_Click()", "    Changed = True", "End Sub", "", ""]
        else:
            out += [f"Private Sub txt{p}_Change()", "    Changed = True", "End Sub", "", ""]
    out.append("Private Sub PropertyPage_ApplyChanges()")
    for p in props:
        out.append(f"    SelectedControls(0).{p} = (chk{p}.Value = vbChecked)" if p in boolean else
                   f"    SelectedControls(0).{p} = txt{p}.Text")
    out += ["End Sub", "", "", "Private Sub PropertyPage_SelectionChanged()"]
    for p in props:
        out.append(f"    chk{p}.Value = (SelectedControls(0).{p} And vbChecked)" if p in boolean else
                   f"    txt{p}.Text = SelectedControls(0).{p}")
    out += ["End Sub", "", ""]
    return out


# ---------------------------------------------------------------- the code of the user control
_INVOKE_RE = re.compile(r'^\s*Attribute\s+([A-Za-z_]\w*)\.VB_ProcData\.VB_Invoke_Property\s*=\s*"(.*)"\s*$',
                        re.IGNORECASE)
_ATTR_LINE = re.compile(r"^\s*Attribute\s+", re.IGNORECASE)


def invoke_code(code, page_of):
    """the code of the user control with Attribute P.VB_ProcData.VB_Invoke_Property = "Page" after the Property
    Get of each property on a page (page_of: lower name -> page), and without it for the other properties"""
    lines = code.replace("\r\n", "\n").split("\n")
    out, i = [], 0
    while i < len(lines):
        line = lines[i]
        m = _INVOKE_RE.match(line)
        if m:                                   # written again below, after its Property Get
            i += 1
            continue
        out.append(line)
        g = _PROP_RE.match(line)
        i += 1
        if g and g.group(1).lower() == "get":
            name = g.group(2)
            while i < len(lines) and _ATTR_LINE.match(lines[i]) and not _INVOKE_RE.match(lines[i]):
                out.append(lines[i])
                i += 1
            while i < len(lines) and _INVOKE_RE.match(lines[i]):
                i += 1
            pg = page_of.get(name.lower())
            if pg:
                out.append(f'Attribute {name}.VB_ProcData.VB_Invoke_Property = "{pg}"')
    return "\n".join(out)


def _set_prop(ctl, name, value):
    """a property of a control where VB6 writes it (alphabetical, before the controls)"""
    from lbvc.project.source import Property
    p = ctl.prop(name)
    if p is not None:
        p.value = value
        return
    p = Property(name, value)
    ctl.items.append(p)
    pos = None
    for i, (k, it) in enumerate(ctl.order):
        if k == "c" or it.name.lower() > name.lower():
            pos = i
            break
    ctl.order.insert(len(ctl.order) if pos is None else pos, ("p", p))


def _del_prop(ctl, name):
    p = ctl.prop(name)
    if p is not None:
        ctl.items.remove(p)
        ctl.order = [(k, it) for k, it in ctl.order if it is not p]


def apply(model, project, uc, set_code=None, new_module=None):
    """Finish: the new pages as PropertyPage modules, the pages whose properties changed written again, the
    PropertyPages of the user control, the attributes of its properties and its VB_Ext_KEY. set_code(module,
    text): the new code of a module (an open code window). new_module(kind, name, ext, text) -> the module (the
    IDE's); returns the modules it made or wrote again"""
    from lbvc.project.frx import FrxData
    changed = []
    for pg in model.custom_pages():
        if not pg.checked and not pg.existing:
            continue
        if pg.existing and [p.lower() for p in pg.props] == [p.lower() for p in pg.old_props]:
            continue
        text = page_text(pg.name, pg.props, model.types)
        if pg.existing and pg.module is not None:
            from .vbmodel import parse_source
            sf = parse_source(text.replace("\n", "\r\n"), None, "PropertyPage", True, project.encoding)
            sf.encoding = project.encoding
            pg.module.sf = sf
            pg.module.editor_text = None
            pg.module.dirty = True
            m = pg.module
        else:
            m = new_module("PropertyPage", pg.name, ".pag", text)
            pg.module, pg.existing = m, True
        pg.old_props = list(pg.props)
        changed.append(m)
    names = [p.name for p in model.selected_pages()]
    root = uc.sf.root
    if names:
        _set_prop(root, "PropertyPages", FrxData("raw", pages_blob(names, project.encoding)))
    else:
        _del_prop(root, "PropertyPages")
    page_of = {x.lower(): p.name for p in model.custom_pages() if p.checked for x in p.props}
    new = invoke_code(uc.code_text(), page_of)
    if set_code is not None:
        set_code(uc, new)
    else:
        uc.editor_text = None
        uc.sf.code = new.split("\n")
    keep = [a for a in uc.sf.attributes
            if not (a[0].strip().lower() == "vb_ext_key" and WIZARD_KEY.lower() in (a[1] or "").lower())]
    line = f'Attribute VB_Ext_KEY = "{WIZARD_KEY}" ,"Yes"'
    keep.append(("VB_Ext_KEY", f'"{WIZARD_KEY}" ,"Yes"', line))
    uc.sf.attributes = keep
    uc.dirty = True
    project.dirty = True
    return changed


def summary(model):
    """the Summary Report (our own text)"""
    pages = model.selected_pages()
    out = [f"The Property Page Wizard has made the property pages of {model.uc}.", "",
           "Pages: " + (", ".join(p.name for p in pages) or "(none)")]
    for p in pages:
        if not p.standard:
            out.append(f"  {p.name}: " + (", ".join(p.props) or "(no properties)"))
    out += ["", "What to do next:",
            "1) Property pages work at design time: put the control on a form of a test project, select it and "
            "press Shift+F4 (or right-click it, Properties).",
            "2) Check that the Properties window shows the values you change on a page, with one control and with "
            "several selected.",
            "3) A property with a list of values (an Enum) is better in a ComboBox than in the TextBox the wizard "
            "made: change the page and fill the list in SelectionChanged.",
            "4) Check the values in ApplyChanges; set Changed = True to keep the page open when one is wrong.",
            "5) A color property declared As OLE_COLOR uses the StandardColor page by itself.",
            "", "Run the wizard again to change the pages: a page whose properties change is written again."]
    return "\n".join(out) + "\n"
