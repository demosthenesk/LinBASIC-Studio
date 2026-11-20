"""The properties of the intrinsic VB6 controls as the Properties window shows them, with their
defaults and the way VB6 writes them in a .frm.

The names of the enumerated values ("1 - Fixed Single") are the ones of the VB6 Properties
window; VB6 writes the name after the number as a comment ("BorderStyle = 1  'Fixed Single").
A property that has its default value is not written in the .frm (VB6 leaves it out), except the
ones that VB6 always writes (position, size, TabIndex, the scale of forms and picture boxes).
"""

from lbvc.project.source import Property, PropertyBlock, Raw
from lbvc.project.frx import FrxData, needs_frx, sniff_image

# ---------------------------------------------------------------- enumerations
ENUMS = {
    "appearance": [(0, "Flat"), (1, "3D")],
    "dragmode": [(0, "Manual"), (1, "Automatic")],
    "oledragmode": [(0, "Manual"), (1, "Automatic")],
    "oledropmode": [(0, "None"), (1, "Manual")],
    "mousepointer": [(0, "Default"), (1, "Arrow"), (2, "Cross"), (3, "I-Beam"), (4, "Icon"), (5, "Size"),
                     (6, "Size NE SW"), (7, "Size N S"), (8, "Size NW SE"), (9, "Size W E"), (10, "Up Arrow"),
                     (11, "Hourglass"), (12, "No Drop"), (13, "Arrow and Hourglass"), (14, "Arrow and Question"),
                     (15, "Size All"), (99, "Custom")],
    "linkmode": [(0, "None"), (1, "Automatic"), (2, "Manual"), (3, "Notify")],
    "formlinkmode": [(0, "None"), (1, "Source")],
    "border": [(0, "None"), (1, "Fixed Single")],
    "formborder": [(0, "None"), (1, "Fixed Single"), (2, "Sizable"), (3, "Fixed Dialog"), (4, "Fixed ToolWindow"),
                   (5, "Sizable ToolWindow")],
    "lineborder": [(0, "Transparent"), (1, "Solid"), (2, "Dash"), (3, "Dot"), (4, "Dash-Dot"), (5, "Dash-Dot-Dot"),
                   (6, "Inside Solid")],
    "align3": [(0, "Left Justify"), (1, "Right Justify"), (2, "Center")],
    "align2": [(0, "Left Justify"), (1, "Right Justify")],
    "scrollbars": [(0, "None"), (1, "Horizontal"), (2, "Vertical"), (3, "Both")],
    "windowstate": [(0, "Normal"), (1, "Minimized"), (2, "Maximized")],
    "checkvalue": [(0, "Unchecked"), (1, "Checked"), (2, "Grayed")],
    "fillstyle": [(0, "Solid"), (1, "Transparent"), (2, "Horizontal Line"), (3, "Vertical Line"),
                  (4, "Upward Diagonal"), (5, "Downward Diagonal"), (6, "Cross"), (7, "Diagonal Cross")],
    "drawstyle": [(0, "Solid"), (1, "Dash"), (2, "Dot"), (3, "Dash-Dot"), (4, "Dash-Dot-Dot"), (5, "Transparent"),
                  (6, "Inside Solid")],
    "drawmode": [(1, "Blackness"), (2, "Not Merge Pen"), (3, "Mask Not Pen"), (4, "Not Copy Pen"), (5, "Mask Pen Not"),
                 (6, "Invert"), (7, "Xor Pen"), (8, "Not Mask Pen"), (9, "Mask Pen"), (10, "Not Xor Pen"), (11, "Nop"),
                 (12, "Merge Not Pen"), (13, "Copy Pen"), (14, "Merge Pen Not"), (15, "Merge Pen"), (16, "Whiteness")],
    "scalemode": [(0, "User"), (1, "Twip"), (2, "Point"), (3, "Pixel"), (4, "Character"), (5, "Inch"),
                  (6, "Millimeter"), (7, "Centimeter")],
    "combostyle": [(0, "Dropdown Combo"), (1, "Simple Combo"), (2, "Dropdown List")],
    "multiselect": [(0, "None"), (1, "Simple"), (2, "Extended")],
    "shape": [(0, "Rectangle"), (1, "Square"), (2, "Oval"), (3, "Circle"), (4, "Rounded Rectangle"),
              (5, "Rounded Square")],
    "align": [(0, "None"), (1, "Align Top"), (2, "Align Bottom"), (3, "Align Left"), (4, "Align Right")],
    "backstyle": [(0, "Transparent"), (1, "Opaque")],
    "buttonstyle": [(0, "Standard"), (1, "Graphical")],
    "liststyle": [(0, "Standard"), (1, "Checkbox")],
    "startup": [(0, "Manual"), (1, "CenterOwner"), (2, "CenterScreen"), (3, "Windows Default")],
    "palettemode": [(0, "Halftone"), (1, "UseZOrder"), (2, "Custom")],
    "negotiate": [(0, "None"), (1, "Left"), (2, "Middle"), (3, "Right")],
    "imemode": [(0, "No Operation"), (1, "ON"), (2, "OFF"), (3, "DISABLE"), (4, "DBCS HIRAGANA"),
                (5, "DBCS KATAKANA"), (6, "SBCS KATAKANA"), (7, "DBCS ALPHABET"), (8, "SBCS ALPHABET"),
                (9, "DBCS HANGUL"), (10, "SBCS HANGUL")],
}

# the system colors of the System tab (index n is &H800000nn&)
SYSTEM_COLORS = ["Scroll Bars", "Desktop", "Active Title Bar", "Inactive Title Bar", "Menu Bar",
                 "Window Background", "Window Frame", "Menu Text", "Window Text", "Active Title Bar Text",
                 "Active Border", "Inactive Border", "Application Workspace", "Highlight", "Highlight Text",
                 "Button Face", "Button Shadow", "Disabled Text", "Button Text", "Inactive Title Bar Text",
                 "Button Highlight", "Button Dark Shadow", "Button Light Shadow", "ToolTip Text", "ToolTip"]

# menu shortcuts: the list of the Menu Editor and how the .frm writes them
SHORTCUTS = [("(None)", "")]
SHORTCUTS += [(f"Ctrl+{c}", f"^{c}") for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"]
SHORTCUTS += [(f"F{n}", f"{{F{n}}}") for n in (1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12)]
SHORTCUTS += [(f"Ctrl+F{n}", f"^{{F{n}}}") for n in (1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12)]
SHORTCUTS += [(f"Shift+F{n}", f"+{{F{n}}}") for n in (1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12)]
SHORTCUTS += [(f"Shift+Ctrl+F{n}", f"+^{{F{n}}}") for n in (1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12)]
SHORTCUTS += [("Ctrl+Ins", "^{INSERT}"), ("Shift+Ins", "+{INSERT}"), ("Del", "{DEL}"), ("Shift+Del", "+{DEL}"),
              ("Alt+Bksp", "%{BKSP}")]
SHORTCUT_OF_TEXT = {t: code for t, code in SHORTCUTS}
TEXT_OF_SHORTCUT = {code.upper(): t for t, code in SHORTCUTS}

BTNFACE, WINDOW, WINDOWTEXT, BTNTEXT, APPWORKSPACE = 0x8000000F, 0x80000005, 0x80000008, 0x80000012, 0x8000000C


class VBColor(int):
    """a color as VB6 writes it: &H00BBGGRR& or &H800000nn&"""

    def __new__(cls, v):
        v = int(v) & 0xFFFFFFFF
        return super().__new__(cls, v - (1 << 32) if v >= 0x80000000 else v)

    def __str__(self):
        return f"&H{int(self) & 0xFFFFFFFF:08X}&"

    __repr__ = __str__


# ---------------------------------------------------------------- the property tables
class P:
    """one property: kind is bool, int, single, string, enum, color, font, picture, list, shortcut,
    name, index, pos (Left/Top/Width/Height), formpos (the outer size of a form)"""
    __slots__ = ("name", "kind", "default", "enum", "category", "always", "note")

    def __init__(self, name, kind, default=None, enum=None, category="Misc", always=False, note=""):
        self.name, self.kind, self.default, self.enum = name, kind, default, enum
        self.category, self.always, self.note = category, always, note


def _b(name, d, cat="Behavior"):
    return P(name, "bool", d, category=cat)


def _e(name, enum, d, cat="Appearance"):
    return P(name, "enum", d, enum, cat)


def _c(name, d, cat="Appearance"):
    return P(name, "color", d, category=cat)


def _s(name, d="", cat="Misc"):
    return P(name, "string", d, category=cat)


def _i(name, d=0, cat="Misc"):
    return P(name, "int", d, category=cat)


def _pic(name, cat="Misc"):
    return P(name, "picture", None, category=cat, note="frx")


def _pos(name):
    return P(name, "pos", 0, category="Position", always=True)


NAME = P("(Name)", "name", category="Misc")
INDEX = P("Index", "index", None, category="Misc")
APPEARANCE = _e("Appearance", "appearance", 1)
DRAGICON, DRAGMODE = _pic("DragIcon", "Behavior"), _e("DragMode", "dragmode", 0, "Behavior")
ENABLED, VISIBLE = _b("Enabled", True), _b("Visible", True)
MOUSEICON, MOUSEPOINTER = _pic("MouseIcon"), _e("MousePointer", "mousepointer", 0, "Misc")
TAG, TOOLTIP = _s("Tag"), _s("ToolTipText")
HELPID, WHATSID = _i("HelpContextID"), _i("WhatsThisHelpID")
TABINDEX = P("TabIndex", "int", 0, category="Behavior", always=True)
TABSTOP, CAUSESVAL = _b("TabStop", True), _b("CausesValidation", True)
OLEDRAG, OLEDROP = _e("OLEDragMode", "oledragmode", 0, "Behavior"), _e("OLEDropMode", "oledropmode", 0, "Behavior")
RTL = _b("RightToLeft", False)
FONT = P("Font", "font", None, category="Font")
LEFT, TOP, WIDTH, HEIGHT = _pos("Left"), _pos("Top"), _pos("Width"), _pos("Height")
DATA = [P("DataField", "string", "", category="Data", note="datafield"),
        P("DataFormat", "dataformat", None, category="Data"), _s("DataMember", cat="Data"),
        P("DataSource", "string", "", category="Data", note="datasource")]
LINK = [_s("LinkItem", cat="DDE"), _e("LinkMode", "linkmode", 0, "DDE"), _i("LinkTimeout", 50, "DDE"),
        _s("LinkTopic", cat="DDE")]
CAPTION = _s("Caption", "", "Appearance")
POSSIZE = [LEFT, TOP, WIDTH, HEIGHT]
COMMON = [NAME, INDEX, DRAGICON, DRAGMODE, ENABLED, MOUSEICON, MOUSEPOINTER, TAG, TOOLTIP, VISIBLE,
          WHATSID] + POSSIZE
BUTTONLIKE = [APPEARANCE, _c("BackColor", BTNFACE), CAPTION, CAUSESVAL, _pic("DisabledPicture", "Appearance"),
              _pic("DownPicture", "Appearance"), FONT, HELPID, _c("MaskColor", 0x00C0C0C0), OLEDROP,
              _pic("Picture", "Appearance"), RTL, _e("Style", "buttonstyle", 0), TABINDEX, TABSTOP,
              _b("UseMaskColor", False)]
FILEISH = [APPEARANCE, _c("BackColor", WINDOW), CAUSESVAL, FONT, _c("ForeColor", WINDOWTEXT), HELPID, OLEDROP,
           TABINDEX, TABSTOP]
DRAWING = [_b("AutoRedraw", False), _b("ClipControls", True), _e("DrawMode", "drawmode", 13, "Behavior"),
           _e("DrawStyle", "drawstyle", 0), P("DrawWidth", "int", 1, category="Appearance"), _c("FillColor", 0),
           _e("FillStyle", "fillstyle", 1), _b("FontTransparent", True, "Font"), _b("HasDC", True),
           _pic("Picture", "Appearance"), P("ScaleHeight", "scale", None, category="Scale", always=True),
           P("ScaleLeft", "scale", 0, category="Scale"), _e("ScaleMode", "scalemode", 1, "Scale"),
           P("ScaleTop", "scale", 0, category="Scale"), P("ScaleWidth", "scale", None, category="Scale", always=True)]

CLASSES = {
    "CommandButton": COMMON + BUTTONLIKE + [_b("Cancel", False, "Misc"), _b("Default", False, "Misc")],
    "CheckBox": COMMON + BUTTONLIKE + DATA + [_e("Alignment", "align2", 0), _c("ForeColor", BTNTEXT),
                                              _e("Value", "checkvalue", 0, "Misc")],
    "OptionButton": COMMON + BUTTONLIKE + [_e("Alignment", "align2", 0), _c("ForeColor", BTNTEXT),
                                           _b("Value", False, "Misc")],
    "Label": COMMON + DATA + LINK + [_e("Alignment", "align3", 0), APPEARANCE, _b("AutoSize", False, "Position"),
                                     _c("BackColor", BTNFACE), _e("BackStyle", "backstyle", 1),
                                     _e("BorderStyle", "border", 0), CAPTION, FONT, _c("ForeColor", BTNTEXT),
                                     OLEDROP, RTL, TABINDEX, _b("UseMnemonic", True, "Misc"),
                                     _b("WordWrap", False, "Position")],
    "TextBox": COMMON + DATA + LINK + [_e("Alignment", "align3", 0), APPEARANCE, _c("BackColor", WINDOW),
                                       _e("BorderStyle", "border", 1), CAUSESVAL, FONT,
                                       _c("ForeColor", WINDOWTEXT), HELPID, _b("HideSelection", True, "Misc"),
                                       _e("IMEMode", "imemode", 0, "Misc"),
                                       _b("Locked", False, "Misc"), P("MaxLength", "int", 0, category="Behavior"),
                                       _b("MultiLine", False, "Behavior"), OLEDRAG, OLEDROP,
                                       _s("PasswordChar", "", "Misc"), RTL,
                                       _e("ScrollBars", "scrollbars", 0, "Behavior"), TABINDEX, TABSTOP,
                                       P("Text", "text", "", category="Text")],
    "Frame": COMMON + [APPEARANCE, _c("BackColor", BTNFACE), _e("BorderStyle", "border", 1), CAPTION,
                       _b("ClipControls", True), FONT, _c("ForeColor", BTNTEXT), HELPID, OLEDROP, RTL, TABINDEX],
    "ComboBox": COMMON + DATA + [APPEARANCE, _c("BackColor", WINDOW), CAUSESVAL, FONT, _c("ForeColor", WINDOWTEXT),
                                 HELPID, _b("IntegralHeight", True, "List"),
                                 P("ItemData", "list", None, category="List", note="frx"),
                                 P("List", "list", None, category="List", note="frx"), _b("Locked", False, "Misc"),
                                 OLEDRAG, OLEDROP, RTL, _b("Sorted", False, "List"),
                                 _e("Style", "combostyle", 0), TABINDEX, TABSTOP, P("Text", "text", "", category="Text")],
    "ListBox": COMMON + DATA + [APPEARANCE, _c("BackColor", WINDOW), CAUSESVAL,
                                P("Columns", "int", 0, category="List"), FONT, _c("ForeColor", WINDOWTEXT), HELPID,
                                _b("IntegralHeight", True, "List"),
                                P("ItemData", "list", None, category="List", note="frx"),
                                P("List", "list", None, category="List", note="frx"),
                                _e("MultiSelect", "multiselect", 0, "Behavior"), OLEDRAG, OLEDROP, RTL,
                                _b("Sorted", False, "List"), _e("Style", "liststyle", 0), TABINDEX, TABSTOP],
    "HScrollBar": None, "VScrollBar": None,
    "Timer": [NAME, INDEX, _b("Enabled", True), P("Interval", "int", 0, category="Behavior"),
              P("Left", "pos", 0, category="Position", always=True), TAG,
              P("Top", "pos", 0, category="Position", always=True)],
    "DriveListBox": COMMON + FILEISH,
    "DirListBox": COMMON + FILEISH + [OLEDRAG],
    "FileListBox": COMMON + FILEISH + [OLEDRAG, _b("Archive", True, "Misc"), _b("Hidden", False, "Misc"),
                                       _e("MultiSelect", "multiselect", 0, "Behavior"), _b("Normal", True, "Misc"),
                                       _s("Pattern", "*.*", "Misc"), _b("ReadOnly", True, "Misc"),
                                       _b("System", False, "Misc")],
    "Shape": [NAME, INDEX, _c("BackColor", WINDOW), _e("BackStyle", "backstyle", 0), _c("BorderColor", WINDOWTEXT),
              _e("BorderStyle", "lineborder", 1), P("BorderWidth", "int", 1, category="Appearance"),
              _e("DrawMode", "drawmode", 13), _c("FillColor", 0), _e("FillStyle", "fillstyle", 1),
              _e("Shape", "shape", 0), TAG, VISIBLE] + POSSIZE,
    "Line": [NAME, INDEX, _c("BorderColor", WINDOWTEXT), _e("BorderStyle", "lineborder", 1),
             P("BorderWidth", "int", 1, category="Appearance"), _e("DrawMode", "drawmode", 13), TAG, VISIBLE,
             P("X1", "pos", 0, category="Position", always=True), P("X2", "pos", 0, category="Position", always=True),
             P("Y1", "pos", 0, category="Position", always=True), P("Y2", "pos", 0, category="Position", always=True)],
    "Image": COMMON + DATA + [APPEARANCE, _e("BorderStyle", "border", 0), OLEDRAG, OLEDROP, _pic("Picture", "Appearance"),
                              _b("Stretch", False, "Misc")],
    "PictureBox": COMMON + DATA + LINK + DRAWING + [
        _e("Align", "align", 0, "Position"), APPEARANCE, _b("AutoSize", False, "Position"), _c("BackColor", BTNFACE),
        _e("BorderStyle", "border", 1), CAUSESVAL, FONT, _c("ForeColor", BTNTEXT), HELPID,
        _b("Negotiate", False, "Misc"), OLEDRAG, OLEDROP, RTL, TABINDEX, TABSTOP],
    "Form": [NAME, APPEARANCE, _c("BackColor", BTNFACE), _e("BorderStyle", "formborder", 2), CAPTION,
             _b("ControlBox", True, "Misc"), _b("Enabled", True), FONT, _c("ForeColor", BTNTEXT), HELPID,
             _pic("Icon", "Misc"), _b("KeyPreview", False, "Misc"), _e("LinkMode", "formlinkmode", 0, "DDE"),
             _s("LinkTopic", "", "DDE"), _b("MaxButton", True, "Misc"), _b("MDIChild", False, "Misc"),
             _b("MinButton", True, "Misc"), MOUSEICON, MOUSEPOINTER, _b("Moveable", True, "Position"),
             _b("NegotiateMenus", True, "Misc"), OLEDROP, _pic("Palette", "Misc"),
             _e("PaletteMode", "palettemode", 0, "Misc"), RTL, _b("ShowInTaskbar", True, "Misc"),
             _e("StartUpPosition", "startup", 0, "Position"), TAG, _b("Visible", True),
             _b("WhatsThisButton", False, "Misc"), _b("WhatsThisHelp", False, "Misc"),
             _e("WindowState", "windowstate", 0, "Position"),
             P("Height", "formpos", 0, category="Position"), P("Left", "formpos", 0, category="Position"),
             P("Top", "formpos", 0, category="Position"), P("Width", "formpos", 0, category="Position")] +
            [p for p in DRAWING],
    "MDIForm": [NAME, APPEARANCE, _b("AutoShowChildren", True, "Behavior"), _c("BackColor", APPWORKSPACE), CAPTION,
                _b("Enabled", True), HELPID, _pic("Icon", "Misc"), _e("LinkMode", "formlinkmode", 0, "DDE"),
                _s("LinkTopic", "", "DDE"), MOUSEICON, MOUSEPOINTER, _b("Moveable", True, "Position"),
                _b("NegotiateToolbars", True, "Misc"), OLEDROP, _pic("Picture", "Appearance"), RTL,
                _b("ScrollBars", True), _e("StartUpPosition", "startup", 0, "Position"), TAG, _b("Visible", True),
                _b("WhatsThisHelp", False, "Misc"), _e("WindowState", "windowstate", 0, "Position"),
                P("Height", "formpos", 0, category="Position"), P("Left", "formpos", 0, category="Position"),
                P("Top", "formpos", 0, category="Position"), P("Width", "formpos", 0, category="Position")],
    "Menu": [NAME, INDEX, _s("Caption", "", "Misc"), _b("Checked", False, "Misc"), _b("Enabled", True),
             HELPID, _e("NegotiatePosition", "negotiate", 0, "Misc"), P("Shortcut", "shortcut", "", category="Misc"),
             TAG, _b("Visible", True), _b("WindowList", False, "Misc")],
}
# a user control (.ctl): Public is its VB_Exposed (the module keeps it, not the .ctl's Begin block)
CLASSES["UserControl"] = [NAME, _s("AccessKeys", "", "Misc"), _b("Alignable", False, "Position"), APPEARANCE,
                          _c("BackColor", BTNFACE), _e("BackStyle", "ucbackstyle", 1), _e("BorderStyle", "border", 0),
                          _b("CanGetFocus", True, "Behavior"), _b("ControlContainer", False, "Misc"),
                          _b("DefaultCancel", False, "Misc"), _b("EditAtDesignTime", False, "Misc"), _b("Enabled", True),
                          FONT, _c("ForeColor", BTNTEXT), _b("ForwardFocus", False, "Behavior"),
                          _b("InvisibleAtRuntime", False, "Behavior"), MOUSEICON, MOUSEPOINTER, OLEDROP,
                          _pic("Picture", "Appearance"), P("Public", "enum", 0, "public", "Misc"), RTL,
                          _b("Windowless", False, "Misc")] + [p for p in DRAWING]
ENUMS["ucbackstyle"] = [(0, "Transparent"), (1, "Opaque")]
# a property page (.pag): the dialog of a user control's properties (design time)
CLASSES["PropertyPage"] = [NAME, APPEARANCE, _c("BackColor", BTNFACE), _s("Caption", "", "Appearance"), FONT,
                           _c("ForeColor", BTNTEXT), MOUSEICON, MOUSEPOINTER, OLEDROP, _pic("Picture", "Appearance"),
                           P("Public", "enum", 0, "public", "Misc"), RTL] + [p for p in DRAWING]
_SCROLL = [CAUSESVAL, HELPID, P("LargeChange", "int", 1, category="Behavior"), P("Max", "int", 32767, category="Behavior"),
           P("Min", "int", 0, category="Behavior"), RTL, P("SmallChange", "int", 1, category="Behavior"), TABINDEX,
           TABSTOP, P("Value", "int", 0, category="Misc")]
CLASSES["HScrollBar"] = COMMON + _SCROLL
# the Data control (DAO): VB6 always writes its database properties; after DefaultType VB6 writes the
# name of the first value whatever the value is (VB6 samples: "2  'UseODBC"; VB6 oracle d_data: "1  'UseODBC")
ENUMS.update({
    "bofaction": [(0, "Move First"), (1, "BOF")],
    "eofaction": [(0, "Move Last"), (1, "EOF"), (2, "Add New")],
    "cursortype": [(0, "DefaultCursor"), (1, "ODBCCursor"), (2, "ServerSideCursor")],
    "defaulttype": [(1, "UseODBC"), (2, "UseJet")],
    "rstype": [(0, "Table"), (1, "Dynaset"), (2, "Snapshot")],
})
CLASSES["Data"] = [NAME, INDEX, DRAGICON, DRAGMODE, ENABLED, MOUSEICON, MOUSEPOINTER, TAG, TOOLTIP, VISIBLE,
                   WHATSID] + POSSIZE + [
    _e("Align", "align", 0, "Position"), APPEARANCE, _c("BackColor", WINDOW), _e("BOFAction", "bofaction", 0, "Behavior"),
    CAPTION, P("Connect", "string", "Access", category="Misc", always=True, note="connect"),
    P("DatabaseName", "string", "", category="Misc", always=True, note="dbfile"),
    P("DefaultCursorType", "enum", 0, "cursortype", "Misc", always=True),
    P("DefaultType", "enum", 2, "defaulttype", "Misc", always=True),
    _e("EOFAction", "eofaction", 0, "Behavior"), P("Exclusive", "bool", False, category="Behavior", always=True),
    FONT, _c("ForeColor", WINDOWTEXT), _b("Negotiate", False, "Misc"), OLEDROP,
    P("Options", "int", 0, category="Behavior", always=True), P("ReadOnly", "bool", False, category="Behavior", always=True),
    P("RecordsetType", "enum", 1, "rstype", "Behavior", always=True),
    P("RecordSource", "string", "", category="Misc", always=True, note="recordsource"), RTL]
CLASSES["VScrollBar"] = COMMON + _SCROLL
CATEGORIES = ["Appearance", "Behavior", "DDE", "Data", "Font", "List", "Misc", "Position", "Scale", "Text"]


# ---------------------------------------------------------------- the controls of the .ocx libraries
# (ocxlib.py): the properties the Properties window shows and the .frm keeps, with the defaults of
# the controls (a property at its default is not written)
def _n(name, d=0, cat="Misc"):
    return P(name, "num", d, category=cat)


ENUMS.update({
    "orientation": [(0, "ccOrientationHorizontal"), (1, "ccOrientationVertical")],
    "scrolling": [(0, "ccScrollingStandard"), (1, "ccScrollingSmooth")],
    "tickstyle": [(0, "sldBottomRight"), (1, "sldTopLeft"), (2, "sldBoth"), (3, "sldNoTicks")],
    "sbrstyle": [(0, "sbrNormal"), (1, "sbrSimple")],
    "tbrstyle": [(0, "tbrStandard"), (1, "tbrFlat")],
    "tbrtext": [(0, "tbrTextAlignBottom"), (1, "tbrTextAlignRight")],
    "tabstyle": [(0, "tabTabs"), (1, "tabButtons"), (2, "tabFlatButtons")],
    "tabwidth": [(0, "tabJustified"), (1, "tabNonJustified"), (2, "tabFixed")],
    "placement": [(0, "tabPlacementTop"), (1, "tabPlacementBottom"), (2, "tabPlacementLeft"), (3, "tabPlacementRight")],
    "treestyle": [(0, "tvwTextOnly"), (1, "tvwPictureText"), (2, "tvwPlusMinusText"), (3, "tvwPlusPictureText"),
                  (4, "tvwTreelinesText"), (5, "tvwTreelinesPictureText"), (6, "tvwTreelinesPlusMinusText"),
                  (7, "tvwTreelinesPlusMinusPictureText")],
    "treelines": [(0, "tvwTreeLines"), (1, "tvwRootLines")],
    "labeledit": [(0, "lvwAutomatic"), (1, "lvwManual")],
    "lvwview": [(0, "lvwIcon"), (1, "lvwSmallIcon"), (2, "lvwList"), (3, "lvwReport")],
    "lvwarrange": [(0, "lvwNone"), (1, "lvwAutoLeft"), (2, "lvwAutoTop")],
    "sortorder": [(0, "lvwAscending"), (1, "lvwDescending")],
    "ccborder": [(0, "ccNone"), (1, "ccFixedSingle")],
    "ccappearance": [(0, "ccFlat"), (1, "cc3D")],
    "updownalign": [(0, "cc2AlignmentLeft"), (1, "cc2AlignmentRight")],
    "cc2orientation": [(0, "cc2OrientationVertical"), (1, "cc2OrientationHorizontal")],
    "dtpformat": [(0, "dtpLongDate"), (1, "dtpShortDate"), (2, "dtpTime"), (3, "dtpCustom")],
    "dayofweek": [(0, "mvwUseSystem"), (1, "mvwSunday"), (2, "mvwMonday"), (3, "mvwTuesday"), (4, "mvwWednesday"),
                  (5, "mvwThursday"), (6, "mvwFriday"), (7, "mvwSaturday")],
    "fsbappearance": [(0, "fsb3D"), (1, "fsbFlat"), (2, "fsbTrack3D")],
    "ssstyle": [(0, "ssStyleTabbedDialog"), (1, "ssStylePropertyPage")],
    "sstaborient": [(0, "ssTabOrientationTop"), (1, "ssTabOrientationBottom"), (2, "ssTabOrientationLeft"),
                    (3, "ssTabOrientationRight")],
    "flexresize": [(0, "flexResizeNone"), (1, "flexResizeColumns"), (2, "flexResizeRows"), (3, "flexResizeBoth")],
    "flexgrid": [(0, "flexGridNone"), (1, "flexGridFlat"), (2, "flexGridInset"), (3, "flexGridRaised")],
    "flexfocus": [(0, "flexFocusNone"), (1, "flexFocusLight"), (2, "flexFocusHeavy")],
    "flexhighlight": [(0, "flexHighlightNever"), (1, "flexHighlightAlways"), (2, "flexHighlightWithFocus")],
    "flexfill": [(0, "flexFillSingle"), (1, "flexFillRepeat")],
    "flexmerge": [(0, "flexMergeNever"), (1, "flexMergeFree"), (2, "flexMergeRestrictRows"),
                  (3, "flexMergeRestrictColumns"), (4, "flexMergeRestrictAll")],
    "flexsel": [(0, "flexSelectionFree"), (1, "flexSelectionByRow"), (2, "flexSelectionByColumn")],
    "rtfscroll": [(0, "rtfNone"), (1, "rtfHorizontal"), (2, "rtfVertical"), (3, "rtfBoth")],
    "rtfborder": [(0, "rtfNoBorder"), (1, "rtfFixedSingle")],
    "rtfappearance": [(0, "rtfFlat"), (1, "rtfThreeD")],
    "adobof": [(0, "adDoMoveFirst"), (1, "adStayBOF")],
    "rdcbof": [(0, "rdMoveFirst"), (1, "rdBOF")],
    "rdceof": [(0, "rdMoveLast"), (1, "rdEOF"), (2, "rdAddNew")],
    "rdccursor": [(0, "rdUseIfNeeded"), (1, "rdUseOdbc"), (2, "rdUseServer"), (3, "rdUseClientBatch")],
    "rdclock": [(1, "rdConcurReadOnly"), (2, "rdConcurLock"), (3, "rdConcurRowver"), (4, "rdConcurValues"),
                (5, "rdConcurBatch")],
    "rdctype": [(1, "rdOpenKeyset"), (3, "rdOpenStatic")],
    "rdcprompt": [(0, "rdDriverPrompt"), (1, "rdDriverNoPrompt"), (2, "rdDriverComplete"), (3, "rdDriverCompleteRequired")],
    "rdcappearance": [(0, "rdFlatBevel"), (1, "rd3DBevel")],
    "drpappearance": [(0, "drpFlat"), (1, "drp3D")],
    "drpborder": [(0, "drpNoBorder"), (1, "drpFixedSingle")],
    "drpscroll": [(0, "drpNoScrollBars"), (1, "drpHorizontal"), (2, "drpVertical"), (3, "drpBoth"), (4, "drpAuto")],
    "drpdivider": [(0, "drpNoDivider"), (1, "drpFlatLine"), (2, "drpInset"), (3, "drpRaised")],
    "drpcaption": [(0, "drpNoCaption"), (1, "drpLeftAligned"), (2, "drpRightAligned"), (3, "drpCentered")],
    "adoeof": [(0, "adDoMoveLast"), (1, "adStayEOF"), (2, "adDoAddNew")],
    "adocmd": [(1, "adCmdText"), (2, "adCmdTable"), (4, "adCmdStoredProc"), (8, "adCmdUnknown")],
    "adoloc": [(2, "adUseServer"), (3, "adUseClient")],
    "adocursor": [(1, "adOpenKeyset"), (2, "adOpenDynamic"), (3, "adOpenStatic")],
    "adolock": [(1, "adLockReadOnly"), (2, "adLockPessimistic"), (3, "adLockOptimistic"), (4, "adLockBatchOptimistic")],
    "adoorient": [(0, "adHorizontal"), (1, "adVertical")],
    "matchentry": [(0, "dblBasicMatching"), (1, "dblExtendedMatching")],
    "dbcstyle": [(0, "dbcDropdownCombo"), (1, "dbcSimpleCombo"), (2, "dbcDropdownList")],
    "mciorient": [(0, "mciOrientHorz"), (1, "mciOrientVert")],
    "mciborder": [(0, "mciNone"), (1, "mciFixedSingle")],
    "mcirecmode": [(0, "mciRecordInsert"), (1, "mciRecordOverwrite")],
    "comhandshake": [(0, "comNone"), (1, "comXOnXoff"), (2, "comRTS"), (3, "comRTSXOnXOff")],
    "cominputmode": [(0, "comInputModeText"), (1, "comInputModeBinary")],
    "sckprotocol": [(0, "sckTCPProtocol"), (1, "sckUDPProtocol")],
    "icaccess": [(0, "icUseDefault"), (1, "icDirect"), (2, "icNamedProxy")],
    "mskclipmode": [(0, "mskIncludeLiterals"), (1, "mskExcludeLiterals")],
    "mskappearance": [(0, "mskFlat"), (1, "mskThreeD")],
    "mskborder": [(0, "mskNone"), (1, "mskFixedSingle")],
    "dbgmarquee": [(0, "dbgDottedCellBorder"), (1, "dbgSolidCellBorder"), (2, "dbgHighlightCell"),
                   (3, "dbgHighlightRow"), (4, "dbgHighlightRowRaiseCell"), (5, "dbgNoMarquee"),
                   (6, "dbgFloatingEditor")],
    "dbgscroll": [(0, "dbgNone"), (1, "dbgHorizontal"), (2, "dbgVertical"), (3, "dbgBoth"), (4, "dbgAutomatic")],
    "dbgdivider": [(0, "dbgNoDividers"), (1, "dbgBlackLine"), (2, "dbgDarkGrayLine"), (3, "dbgRaised"),
                   (4, "dbgInset"), (5, "dbgUseForeColor"), (6, "dbgLightGrayLine")],
    "dbgtab": [(0, "dbgControlNavigation"), (1, "dbgColumnNavigation"), (2, "dbgGridNavigation")],
    "dbgmode": [(0, "dbgBound"), (1, "dbgUnbound")],
    "dbgappearance": [(0, "dbgFlat"), (1, "dbg3D")],
    "dbgborder": [(0, "dbgNoBorder"), (1, "dbgFixedSingle")],
    "dblappearance": [(0, "dblFlat"), (1, "dbl3D")],
    "icprotocol": [(0, "icUnknown"), (1, "icDefault"), (2, "icFTP"), (3, "icGopher"), (4, "icHTTP"), (5, "icHTTPS")],
})
_OCX_EXT = [NAME, INDEX, DRAGICON, DRAGMODE, HELPID, WHATSID, TAG, TOOLTIP, VISIBLE, TABINDEX, TABSTOP,
            CAUSESVAL] + POSSIZE
_OCX_ICON = [NAME, INDEX, TAG, P("Left", "pos", 0, category="Position", always=True),
             P("Top", "pos", 0, category="Position", always=True)]
_IL = lambda n="ImageList": _s(n, "", "Misc")   # noqa: E731 - the name of an ImageList on the form
CLASSES.update({
    "CommonDialog": _OCX_ICON + [_b("CancelError", False, "Misc"), _c("Color", 0, "Misc"), P("Copies", "int", 1),
                                 _s("DefaultExt"), _s("DialogTitle"), _s("FileName"), _s("Filter"),
                                 _i("FilterIndex", 0), _i("Flags", 0), _s("FontName"), _n("FontSize", 8),
                                 _i("FromPage", 0), _i("ToPage", 0), _s("HelpFile"), _s("InitDir"), _i("Max", 0),
                                 _i("Min", 0), P("MaxFileSize", "int", 260)],
    "ImageList": _OCX_ICON + [_c("BackColor", WINDOW), _i("ImageHeight", 0), _i("ImageWidth", 0),
                              _c("MaskColor", 0x00C0C0C0), _b("UseMaskColor", True, "Misc")],
    "ProgressBar": _OCX_EXT + [ENABLED, MOUSEPOINTER, OLEDROP, _e("Appearance", "ccappearance", 1),
                               _e("BorderStyle", "ccborder", 0), _n("Max", 100, "Behavior"), _n("Min", 0, "Behavior"),
                               _e("Orientation", "orientation", 0), _e("Scrolling", "scrolling", 0)],
    "Slider": _OCX_EXT + [ENABLED, MOUSEPOINTER, OLEDROP, _e("BorderStyle", "ccborder", 0),
                          P("LargeChange", "int", 5, category="Behavior"), P("Max", "int", 10, category="Behavior"),
                          P("Min", "int", 0, category="Behavior"), _e("Orientation", "orientation", 0),
                          _b("SelectRange", False, "Misc"), _i("SelLength", 0), _i("SelStart", 0),
                          P("SmallChange", "int", 1, category="Behavior"), _i("TickFrequency", 1),
                          _e("TickStyle", "tickstyle", 0), P("Value", "int", 0, category="Misc")],
    "StatusBar": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDROP, _e("Align", "align", 2, "Position"),
                             _b("ShowTips", False, "Misc"), _s("SimpleText"), _e("Style", "sbrstyle", 0)],
    "Toolbar": _OCX_EXT + [ENABLED, MOUSEPOINTER, OLEDROP, _e("Align", "align", 1, "Position"),
                           _b("AllowCustomize", True, "Behavior"), _e("Appearance", "ccappearance", 1),
                           _e("BorderStyle", "ccborder", 0), _n("ButtonHeight", 0), _n("ButtonWidth", 0),
                           _IL("DisabledImageList"), _IL("HotImageList"), _IL(), _b("ShowTips", True, "Misc"),
                           _e("Style", "tbrstyle", 0), _e("TextAlignment", "tbrtext", 0),
                           _b("Wrappable", True, "Behavior")],
    "TabStrip": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDROP, _b("HotTracking", False, "Behavior"), _IL(),
                            _b("MultiRow", False, "Behavior"), _e("Placement", "placement", 0),
                            _b("Separators", False, "Appearance"), _b("ShowTips", True, "Misc"),
                            _e("Style", "tabstyle", 0), _n("TabFixedHeight", 0), _n("TabFixedWidth", 0),
                            _e("TabWidthStyle", "tabwidth", 0)],
    "TreeView": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDRAG, OLEDROP, _e("Appearance", "ccappearance", 1),
                            _e("BorderStyle", "ccborder", 0), _b("Checkboxes", False, "Behavior"),
                            _b("FullRowSelect", False, "Behavior"), _b("HideSelection", True, "Behavior"),
                            _b("HotTracking", False, "Behavior"), _IL(), _n("Indentation", 566.9291),
                            _e("LabelEdit", "labeledit", 0, "Behavior"), _e("LineStyle", "treelines", 0),
                            _s("PathSeparator", "\\"), _b("Scroll", True, "Behavior"),
                            _b("SingleSel", False, "Behavior"), _b("Sorted", False, "Behavior"),
                            _e("Style", "treestyle", 7)],
    "ListView": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDRAG, OLEDROP, _b("AllowColumnReorder", False, "Behavior"),
                            _e("Appearance", "ccappearance", 1), _e("Arrange", "lvwarrange", 0),
                            _c("BackColor", WINDOW), _e("BorderStyle", "ccborder", 1),
                            _b("Checkboxes", False, "Behavior"), _IL("ColumnHeaderIcons"),
                            _b("FlatScrollBar", False, "Appearance"), _c("ForeColor", WINDOWTEXT),
                            _b("FullRowSelect", False, "Behavior"), _b("GridLines", False, "Appearance"),
                            _b("HideColumnHeaders", False, "Appearance"), _b("HideSelection", True, "Behavior"),
                            _b("HotTracking", False, "Behavior"), _b("HoverSelection", False, "Behavior"),
                            _IL("Icons"), _e("LabelEdit", "labeledit", 0, "Behavior"), _b("LabelWrap", True, "Behavior"),
                            _b("MultiSelect", False, "Behavior"), _IL("SmallIcons"), _b("Sorted", False, "Behavior"),
                            _i("SortKey", 0), _e("SortOrder", "sortorder", 0), _e("View", "lvwview", 0)],
    "ImageCombo": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDRAG, OLEDROP, _c("BackColor", WINDOW),
                              _c("ForeColor", WINDOWTEXT), _IL(), _i("Indentation", 0), _b("Locked", False, "Misc"),
                              P("Text", "text", "", category="Text")],
    "Animation": _OCX_EXT + [ENABLED, OLEDROP, _b("AutoPlay", False, "Behavior"), _c("BackColor", BTNFACE),
                             _e("BackStyle", "backstyle", 0), _b("Center", True, "Appearance")],
    "UpDown": _OCX_EXT + [ENABLED, MOUSEPOINTER, OLEDROP, _e("Alignment", "updownalign", 1),
                          _b("AutoBuddy", False, "Behavior"), _s("BuddyControl"), _s("BuddyProperty"),
                          P("Increment", "int", 1, category="Behavior"), P("Max", "int", 10, category="Behavior"),
                          P("Min", "int", 0, category="Behavior"), _e("Orientation", "cc2orientation", 0),
                          _b("SyncBuddy", False, "Behavior"), P("Value", "int", 0, category="Misc"),
                          _b("Wrap", False, "Behavior")],
    "MonthView": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDRAG, OLEDROP, _e("Appearance", "ccappearance", 1),
                             _e("BorderStyle", "ccborder", 0), _e("StartOfWeek", "dayofweek", 0),
                             _i("MaxSelCount", 7), _i("MonthColumns", 1), _i("MonthRows", 1),
                             _b("MultiSelect", False, "Behavior"), _i("ScrollRate", 0), _b("ShowToday", True, "Appearance"),
                             _b("ShowWeekNumbers", False, "Appearance")],
    "DTPicker": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDROP, _b("CheckBox", False, "Behavior"),
                            _s("CustomFormat"), _e("Format", "dtpformat", 1), _b("UpDown", False, "Behavior")],
    "FlatScrollBar": _OCX_EXT + [ENABLED, MOUSEPOINTER, _e("Appearance", "fsbappearance", 1),
                                 P("LargeChange", "int", 1, category="Behavior"), P("Max", "int", 100, category="Behavior"),
                                 P("Min", "int", 0, category="Behavior"), P("SmallChange", "int", 1, category="Behavior"),
                                 P("Value", "int", 0, category="Misc")],
    "SSTab": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDROP, _c("BackColor", BTNFACE), _c("ForeColor", BTNTEXT),
                         _b("ShowFocusRect", True, "Appearance"), _e("Style", "ssstyle", 0), P("Tab", "int", 0),
                         _n("TabHeight", 520), _n("TabMaxWidth", 0), _e("TabOrientation", "sstaborient", 0),
                         P("Tabs", "int", 3), P("TabsPerRow", "int", 3), _b("WordWrap", True, "Appearance")],
    "MSFlexGrid": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDROP, _b("AllowBigSelection", True, "Behavior"),
                              _e("AllowUserResizing", "flexresize", 0, "Behavior"), _c("BackColor", WINDOW),
                              _c("BackColorBkg", APPWORKSPACE), _c("BackColorFixed", BTNFACE),
                              _c("BackColorSel", 0x8000000D), P("Cols", "int", 2, category="Misc"),
                              _e("FillStyle", "flexfill", 0), P("FixedCols", "int", 1, category="Misc"),
                              P("FixedRows", "int", 1, category="Misc"), _e("FocusRect", "flexfocus", 1),
                              _c("ForeColor", WINDOWTEXT), _c("ForeColorFixed", BTNTEXT),
                              _c("ForeColorSel", 0x8000000E), _s("FormatString"), _c("GridColor", 0x00C0C0C0),
                              _c("GridColorFixed", 0), _e("GridLines", "flexgrid", 1),
                              _e("GridLinesFixed", "flexgrid", 2), _e("HighLight", "flexhighlight", 1),
                              _e("MergeCells", "flexmerge", 0), P("Rows", "int", 2, category="Misc"),
                              _e("ScrollBars", "scrollbars", 3, "Behavior"),
                              _e("SelectionMode", "flexsel", 0, "Behavior"), _b("WordWrap", False, "Behavior")],
    "Adodc": _OCX_EXT + [ENABLED, FONT, _e("Align", "align", 0, "Position"), _e("Appearance", "ccappearance", 1),
                         _c("BackColor", WINDOW), _e("BOFAction", "adobof", 0, "Behavior"), _n("CacheSize", 50),
                         _s("Caption", "", "Appearance"), _e("CommandType", "adocmd", 8, "Misc"),
                         _n("CommandTimeout", 30), _s("ConnectionString", "", "Misc"), _n("ConnectionTimeout", 15),
                         _e("CursorLocation", "adoloc", 3, "Misc"), _e("CursorType", "adocursor", 3, "Misc"),
                         _e("EOFAction", "adoeof", 0, "Behavior"), _c("ForeColor", WINDOWTEXT),
                         _e("LockType", "adolock", 3, "Misc"), _n("MaxRecords", 0), _n("Mode", 0),
                         _e("Orientation", "adoorient", 0), _s("Password", "", "Misc"),
                         P("RecordSource", "string", "", category="Misc", note="adosource"),
                         _s("UserName", "", "Misc")],
    "MSRDC": _OCX_EXT + [ENABLED, FONT, _e("Appearance", "rdcappearance", 1), _c("BackColor", WINDOW),
                         _n("BatchSize", 15), _e("BOFAction", "rdcbof", 0, "Behavior"), _s("Caption", "", "Appearance"),
                         _s("Connect", "", "Misc"), _e("CursorDriver", "rdccursor", 0, "Misc"),
                         _s("DataSourceName", "", "Misc"), _e("EOFAction", "rdceof", 0, "Behavior"),
                         _n("ErrorThreshold", -1), _c("ForeColor", WINDOWTEXT), _n("KeysetSize", 0),
                         _e("LockType", "rdclock", 3, "Misc"), _n("LoginTimeout", 15), _s("LogMessages", "", "Misc"),
                         _n("MaxRows", 0), _n("Options", 0), _s("Password", "", "Misc"),
                         _e("Prompt", "rdcprompt", 0, "Misc"), _n("QueryTimeout", 30), _b("ReadOnly", False),
                         _e("ResultsetType", "rdctype", 1, "Misc"), _n("RowsetSize", 100), _s("SQL", "", "Misc"),
                         _s("UserName", "", "Misc")],
    # MSDataRepeaterLib (MSDatRep.Ocx, phase 7.5): RepeatedControlName and RepeaterBindings are blocks of the .frm
    # (drep_name, drep_bindings)
    "DataRepeater": _OCX_EXT + [ENABLED, FONT, _e("Appearance", "drpappearance", 1), _c("BackColor", BTNFACE),
                                _e("BorderStyle", "drpborder", 1), _s("Caption", "", "Appearance"),
                                _e("CaptionStyle", "drpcaption", 1), _c("ForeColor", BTNTEXT),
                                _b("IntegralHeight", True, "Appearance"), _b("RowIndicator", True, "Appearance"),
                                _e("RowDividerStyle", "drpdivider", 2), _e("ScrollBars", "drpscroll", 2, "Behavior"),
                                P("RepeatedControlName", "drepname", "", category="Misc"),
                                P("(RepeaterBindings)", "drepbinds", None, category="Misc")],
    "DataGrid": _OCX_EXT + [ENABLED, FONT, _b("AllowAddNew", False, "Behavior"), _b("AllowArrows", True, "Behavior"),
                            _b("AllowDelete", False, "Behavior"), _b("AllowRowSizing", True, "Behavior"),
                            _b("AllowUpdate", True, "Behavior"), _e("Appearance", "ccappearance", 1),
                            _c("BackColor", WINDOW), _e("BorderStyle", "ccborder", 1), _s("Caption", "", "Appearance"),
                            _b("ColumnHeaders", True, "Appearance"),
                            P("DataMember", "string", "", category="Data"),
                            P("DataSource", "string", "", category="Data", note="datasource"),
                            _c("ForeColor", WINDOWTEXT), _n("HeadLines", 1), _b("RecordSelectors", True, "Appearance"),
                            _n("RowHeight", 15), _b("WrapCellPointer", False, "Behavior"),
                            P("(Columns)", "dgcolumns", None, category="Misc")],
    "DataList": _OCX_EXT + [ENABLED, FONT, _e("Appearance", "ccappearance", 1), _c("BackColor", WINDOW),
                            P("BoundColumn", "string", "", category="Data"),
                            P("DataField", "string", "", category="Data", note="datafield"),
                            P("DataFormat", "dataformat", None, category="Data"),
                            P("DataMember", "string", "", category="Data"),
                            P("DataSource", "string", "", category="Data", note="datasource"),
                            _c("ForeColor", WINDOWTEXT), _b("IntegralHeight", True, "List"),
                            P("ListField", "string", "", category="List"), _b("Locked", False, "Misc"),
                            _e("MatchEntry", "matchentry", 0, "Behavior"), P("RowMember", "string", "", category="Data"),
                            P("RowSource", "string", "", category="Data", note="datasource")],
    "DataCombo": _OCX_EXT + [ENABLED, FONT, _e("Appearance", "ccappearance", 1), _c("BackColor", WINDOW),
                             P("BoundColumn", "string", "", category="Data"),
                             P("DataField", "string", "", category="Data", note="datafield"),
                             P("DataFormat", "dataformat", None, category="Data"),
                             P("DataMember", "string", "", category="Data"),
                             P("DataSource", "string", "", category="Data", note="datasource"),
                             _c("ForeColor", WINDOWTEXT), _b("IntegralHeight", True, "List"),
                             P("ListField", "string", "", category="List"), _b("Locked", False, "Misc"),
                             _e("MatchEntry", "matchentry", 0, "Behavior"), P("RowMember", "string", "", category="Data"),
                             P("RowSource", "string", "", category="Data", note="datasource"),
                             _e("Style", "dbcstyle", 0, "Misc"), P("Text", "text", "", category="Text")],
    # DBGrid: the properties are in its OleObjectBlob (BLOB_PROPS), DataSource in its Bindings
    "DBGrid": _OCX_EXT + [ENABLED, FONT, _b("AllowAddNew", False, "Behavior"), _b("AllowArrows", True, "Behavior"),
                          _b("AllowDelete", False, "Behavior"), _b("AllowRowSizing", True, "Behavior"),
                          _b("AllowUpdate", True, "Behavior"), _e("Appearance", "dbgappearance", 1),
                          _c("BackColor", WINDOW), _e("BorderStyle", "dbgborder", 1), _s("Caption", "", "Appearance"),
                          _b("ColumnHeaders", True, "Appearance"), _e("DataMode", "dbgmode", 0, "Data"),
                          P("DataSource", "string", "", category="Data", note="datasource"),
                          _n("DefColWidth", 0), _c("ForeColor", WINDOWTEXT), _n("HeadLines", 1),
                          _e("MarqueeStyle", "dbgmarquee", 6), _b("MarqueeUnique", True, "Behavior"),
                          _b("RecordSelectors", True, "Appearance"), _e("RowDividerStyle", "dbgdivider", 2),
                          _e("ScrollBars", "dbgscroll", 4, "Behavior"), _b("TabAcrossSplits", False, "Behavior"),
                          _e("TabAction", "dbgtab", 0, "Behavior"), _b("WrapCellPointer", False, "Behavior"),
                          P("(Columns)", "dbgcolumns", None, category="Misc")],
    "DBList": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDRAG, OLEDROP, _e("Appearance", "dblappearance", 1),
                          _c("BackColor", WINDOW), P("BoundColumn", "string", "", category="Data"),
                          P("DataField", "string", "", category="Data", note="datafield"),
                          P("DataSource", "string", "", category="Data", note="datasource"),
                          _c("ForeColor", WINDOWTEXT), _b("IntegralHeight", True, "List"),
                          P("ListField", "string", "", category="List"), _b("Locked", False, "Misc"),
                          _e("MatchEntry", "matchentry", 0, "Behavior"),
                          P("RowSource", "string", "", category="Data", note="datasource")],
    "DBCombo": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDRAG, OLEDROP, _e("Appearance", "dblappearance", 1),
                           _c("BackColor", WINDOW), P("BoundColumn", "string", "", category="Data"),
                           P("DataField", "string", "", category="Data", note="datafield"),
                           P("DataSource", "string", "", category="Data", note="datasource"),
                           _c("ForeColor", WINDOWTEXT), _b("IntegralHeight", True, "List"),
                           P("ListField", "string", "", category="List"), _b("Locked", False, "Misc"),
                           _e("MatchEntry", "matchentry", 0, "Behavior"),
                           P("RowSource", "string", "", category="Data", note="datasource"),
                           _e("Style", "dbcstyle", 0, "Misc"), P("Text", "text", "", category="Text")],
    "RichTextBox": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDRAG, OLEDROP,
                               _e("Appearance", "rtfappearance", 1), _b("AutoVerbMenu", False, "Behavior"),
                               _c("BackColor", WINDOW), _e("BorderStyle", "rtfborder", 1), _n("BulletIndent", 0),
                               _b("DisableNoScroll", False, "Behavior"), _s("FileName"),
                               _b("HideSelection", True, "Behavior"), _b("Locked", False, "Behavior"),
                               P("MaxLength", "int", 0, category="Behavior"), _b("MultiLine", True, "Behavior"),
                               _n("RightMargin", 0), _e("ScrollBars", "rtfscroll", 0, "Behavior")],
    "MSChart": _OCX_EXT + [ENABLED, MOUSEPOINTER, OLEDROP],
    "MMControl": _OCX_EXT + [ENABLED, MOUSEPOINTER, OLEDROP, _b("AutoEnable", True, "Behavior")] + [
        _b(f"{b}{w}", w == "Visible", "Appearance" if w == "Visible" else "Behavior")
        for b in ("Back", "Eject", "Next", "Pause", "Play", "Prev", "Record", "Step", "Stop") for w in ("Enabled", "Visible")] + [
        _e("BorderStyle", "mciborder", 1), P("DeviceType", "string", "", category="Misc", always=True),
        P("FileName", "string", "", category="Misc", always=True), _i("Frames", 1),
        _e("Orientation", "mciorient", 0), _e("RecordMode", "mcirecmode", 1, "Behavior"),
        _b("Shareable", False, "Misc"), _b("Silent", False, "Misc"), _i("UpdateInterval", 1000)],
    "MSComm": _OCX_ICON + [P("CommPort", "int", 1), _b("DTREnable", False, "Misc"), _b("EOFEnable", False, "Misc"),
                           _e("Handshaking", "comhandshake", 0, "Misc"), P("InBufferSize", "int", 1024),
                           P("InputLen", "int", 0), _e("InputMode", "cominputmode", 0, "Misc"),
                           _b("NullDiscard", False, "Misc"), P("OutBufferSize", "int", 512),
                           _s("ParityReplace", "?"), P("RThreshold", "int", 0), _b("RTSEnable", False, "Misc"),
                           _s("Settings", "9600,n,8,1"), P("SThreshold", "int", 0)],
    "Winsock": _OCX_ICON + [P("LocalPort", "int", 0), _e("Protocol", "sckprotocol", 0, "Misc"),
                            _s("RemoteHost", ""), P("RemotePort", "int", 0)],
    "MAPISession": _OCX_ICON + [_b("DownloadMail", True, "Misc"), _b("LogonUI", True, "Misc"),
                                _b("NewSession", False, "Misc"), _s("Password"), _s("UserName")],
    "MAPIMessages": _OCX_ICON + [_s("AddressCaption"), _i("AddressEditFieldCount", 1), _s("AddressLabel"),
                                 _b("AddressModifiable", False, "Misc"), _b("AddressResolveUI", False, "Misc"),
                                 _s("FetchMsgType"), _b("FetchSorted", False, "Misc"),
                                 _b("FetchUnreadOnly", False, "Misc")],
    "WebBrowser": _OCX_EXT + [_b("AddressBar", True, "Misc"), _b("FullScreen", False, "Misc"),
                              _b("MenuBar", True, "Misc"), _b("Offline", False, "Misc"),
                              _b("RegisterAsBrowser", False, "Misc"), _b("RegisterAsDropTarget", True, "Misc"),
                              _b("Silent", False, "Misc"), _b("StatusBar", True, "Misc"),
                              _b("TheaterMode", False, "Misc"), _i("ToolBar", 1)],
    "Inet": _OCX_ICON + [_e("AccessType", "icaccess", 0, "Misc"), _s("Document"), _s("Password"),
                         _e("Protocol", "icprotocol", 1, "Misc"), _s("Proxy"), _s("RemoteHost"), P("RemotePort", "int", 80),
                         P("RequestTimeout", "int", 60), _s("URL"), _s("UserName")],
    "MaskEdBox": _OCX_EXT + [ENABLED, FONT, MOUSEPOINTER, OLEDRAG, OLEDROP, _b("AllowPrompt", False, "Behavior"),
                             _e("Appearance", "mskappearance", 1), _b("AutoTab", False, "Behavior"),
                             _c("BackColor", WINDOW), _e("BorderStyle", "mskborder", 1),
                             _e("ClipMode", "mskclipmode", 0, "Behavior"), _c("ForeColor", WINDOWTEXT),
                             _s("Format", "", "Misc"), _b("HideSelection", True, "Behavior"),
                             _s("Mask", "", "Misc"), P("MaxLength", "int", 64, category="Behavior"),
                             _s("PromptChar", "_", "Misc"), _b("PromptInclude", True, "Behavior")],
    "PictureClip": [NAME, INDEX, TAG] + POSSIZE + [P("Cols", "int", 1), P("Rows", "int", 1),
                                                  _pic("Picture", "Appearance")],
    # ComCtl3 (COMCT332.OCX, phase 7.4): its bands are flat properties (Caption1, Child1 ...: cool_bands)
    "CoolBar": _OCX_EXT + [ENABLED, MOUSEICON, MOUSEPOINTER, OLEDROP, _e("Align", "align", 0, "Position"),
                           _c("BackColor", BTNFACE), _b("BandBorders", True, "Appearance"),
                           _c("EmbossHighlight", 0x80000014), _b("EmbossPicture", False, "Appearance"),
                           _c("EmbossShadow", 0x8000000F), _b("FixedOrder", False, "Behavior"),
                           _c("ForeColor", BTNTEXT), _IL(), _e("Orientation", "cc3orient", 0),
                           _pic("Picture", "Appearance"), _b("VariantHeight", True, "Appearance"),
                           P("(Bands)", "cbbands", None, category="Misc")],
})
ENUMS["cc3orient"] = [(0, "cc3OrientationHorizontal"), (1, "cc3OrientationVertical")]

# the frame of a form window at design time in pixels (left, top, right, bottom), as Windows 7
# draws it (VB6 oracle d_props: a Fixed Dialog with a menu, Width 6000 -> ClientWidth 5910);
# a menu bar adds 20 pixels at the top
FRAMES = {0: (0, 0, 0, 0), 1: (3, 25, 3, 3), 2: (8, 30, 8, 8), 3: (3, 25, 3, 3), 4: (3, 21, 3, 3), 5: (8, 26, 8, 8)}


# the user controls of the project (designer.set_user_controls): lower name -> their properties
DYNAMIC = {}
_UC_KIND = {"boolean": "bool", "string": "string", "long": "num", "integer": "num", "byte": "num", "single": "num",
            "double": "num", "currency": "num", "ole_color": "color", "stdole.ole_color": "color"}


def uctl_props(code):
    """the properties of a user control on a form: its Extender's, and its own public properties that
    can be read and written (Property Get with a Let), as the Properties window shows them"""
    import re
    ext = [NAME, INDEX, CAUSESVAL, _pic("DragIcon", "Behavior"), _e("DragMode", "dragmode", 0, "Behavior"), HELPID,
           TABINDEX, TABSTOP, TAG, TOOLTIP, VISIBLE, WHATSID] + POSSIZE
    gets, lets = {}, set()
    for line in code.splitlines():
        m = re.match(r"^\s*Public\s+Property\s+(Get|Let|Set)\s+([A-Za-z_]\w*)\s*\((.*?)\)\s*(?:As\s+([\w.]+))?",
                     line, re.IGNORECASE)
        if not m:
            continue
        kind, name = m.group(1).lower(), m.group(2)
        if kind == "get" and not m.group(3).strip():
            gets[name.lower()] = (name, (m.group(4) or "Variant").lower())
        elif kind in ("let", "set"):
            lets.add(name.lower())
    have = {p.name.lower() for p in ext}
    for low, (name, t) in gets.items():
        if low in lets and low not in have:
            if t in ("font", "stdfont", "stdole.stdfont"):
                ext.append(P(name, "font", None, category="Font"))
            elif t in ("picture", "stdpicture", "stdole.picture"):
                ext.append(_pic(name, "Misc"))
            else:
                ext.append(P(name, _UC_KIND.get(t, "string"), None, category="Misc"))
    return ext


def props_of(type_name):
    """the properties of a class in the order of the Alphabetic tab: (Name) first"""
    seen, out = set(), []
    for p in CLASSES.get(type_name, DYNAMIC.get((type_name or "").lower(), [NAME])):
        if p.name.lower() in seen:
            continue
        seen.add(p.name.lower())
        out.append(p)
    out.sort(key=lambda p: (p.name != "(Name)", p.name.lower()))
    return out


def prop_def(type_name, name):
    for p in props_of(type_name):
        if p.name.lower() == name.lower():
            return p
    return None


def enum_items(p):
    return ENUMS.get(p.enum, [])


# Instancing of a class module: the values a project type allows (VB6)
INSTANCING_OF_TYPE = {"Exe": [1], "Control": [1, 2], "OleDll": [1, 2, 5, 6], "OleExe": [1, 2, 3, 4, 5, 6]}
ENUMS.update({f"instancing_{t.lower()}": [(n, ["", "Private", "PublicNotCreatable", "SingleUse", "GlobalSingleUse",
                                               "MultiUse", "GlobalMultiUse"][n]) for n in ns]
              for t, ns in INSTANCING_OF_TYPE.items()})
# Public of a user control: True in an ActiveX Control project (False: it can't be)
ENUMS["public"] = [(0, "False"), (1, "True")]


def frm_comment(p, v):
    """the comment VB6 writes after an enumerated value; for DrawMode VB6 writes the name of the
    value before it (VB6 oracle d_props: DrawMode = 7  'Invert)"""
    if p.enum == "drawmode":
        return enum_label(p, v - 1)
    if p.enum == "defaulttype":
        return enum_items(p)[0][1]
    return enum_label(p, v)


def enum_label(p, v):
    for n, t in enum_items(p):
        if n == v:
            return t
    return None


# ---------------------------------------------------------------- reading values
def _num(v, default=0):
    if isinstance(v, bool):
        return -1 if v else 0
    try:
        return int(v) if float(v) == int(float(v)) else float(v)
    except (TypeError, ValueError):
        return default


# a property the .frm keeps under another name: the ADO Data control writes ConnectionString as OLEDBString
FRM_KEYS = {("Adodc", "ConnectionString"): "OLEDBString", ("MSRDC", "ResultsetType"): "RecordsetType",
            ("MSRDC", "SQL"): "RecordSource"}


def frm_prop(ctl, p):
    k = FRM_KEYS.get((getattr(ctl, "type_name", ""), p.name))
    return P(k, p.kind, p.default, p.enum, p.category, p.always, p.note) if k else p


# ---------------------------------------------------------------- DBGrid, DBList, DBCombo (phase 7.6)
# a DBGrid keeps its properties in the OleObjectBlob of the .frx, not in the .frm; the DataSource of a
# DBGrid and the RowSource of a DBList / DBCombo are in their Bindings (.frx) (lbvc/project/oleblob.py)
BLOB_PROPS = {"AllowAddNew", "AllowArrows", "AllowDelete", "AllowRowSizing", "AllowUpdate", "Appearance", "BackColor",
              "BorderStyle", "Caption", "ColumnHeaders", "DataMode", "DefColWidth", "Enabled", "ForeColor", "HeadLines",
              "MarqueeStyle", "MarqueeUnique", "RecordSelectors", "RowDividerStyle", "ScrollBars", "TabAcrossSplits",
              "TabAction", "WrapCellPointer"}


def _frx_bytes(ctl, name):
    v = ctl.get(name)
    return bytes(v.data) if isinstance(v, FrxData) and isinstance(v.data, (bytes, bytearray)) else None


def grid_blob(ctl):
    """the OleObjectBlob of a DBGrid (a new one when it has none)"""
    from lbvc.project.oleblob import new_blob
    b = _frx_bytes(ctl, "OleObjectBlob")
    return b if b else new_blob(ctl.name, int(_num(ctl.get("Width", 6495))), int(_num(ctl.get("Height", 3015))))


def _binding_dispid(ctl, p):
    from lbvc.project.oleblob import DISPID_DATASOURCE, DISPID_ROWSOURCE
    t = getattr(ctl, "type_name", "")
    if p.name == "DataSource":
        return DISPID_DATASOURCE.get(t)
    if p.name == "RowSource":
        return DISPID_ROWSOURCE.get(t)
    return None


DFMT_GUID = "{6D835690-900B-11D0-9484-00A0C91110ED}"   # StdDataFormat: the block of an .ocx control carries it
DFMT_KEYS = ("Type", "Format", "HaveTrueFalseNull", "TrueValue", "FalseValue", "NullValue", "FirstDayOfWeek",
             "FirstWeekOfYear", "LCID", "SubFormatType")


def _dfmt_dict(blk):
    out = {}
    for k in DFMT_KEYS:
        x = blk.get(k)
        if x is not None:
            out[k] = x
    return out


def dataformat_of(ctl):
    """BeginProperty DataFormat of a bound control as a dict, or None (phase 7.7)"""
    blk = ctl.block("DataFormat") if hasattr(ctl, "block") else None
    return None if blk is None else _dfmt_dict(blk)


def _dfmt_items(value):
    """the properties of BeginProperty DataFormat in the order of VB6 (the True / False / Null values only with
    HaveTrueFalseNull)"""
    have = int(value.get("HaveTrueFalseNull", 0) or 0)
    items = [Property("Type", int(value.get("Type", 0) or 0)), Property("Format", str(value.get("Format", "") or "")),
             Property("HaveTrueFalseNull", have)]
    if have:
        items += [Property(k, str(value.get(k, "") if value.get(k) is not None else "")) for k in ("TrueValue", "FalseValue", "NullValue")]
    items += [Property("FirstDayOfWeek", int(value.get("FirstDayOfWeek", 0) or 0)),
              Property("FirstWeekOfYear", int(value.get("FirstWeekOfYear", 0) or 0)),
              Property("LCID", int(value.get("LCID", 1033) or 1033)),
              Property("SubFormatType", int(value.get("SubFormatType", 0) or 0))]
    return items


def set_dataformat(ctl, value):
    """write BeginProperty DataFormat; None takes it away"""
    remove(ctl, "DataFormat", block=True)
    if not value:
        return
    blk = PropertyBlock("DataFormat", None if getattr(ctl, "library", "VB") in ("VB", "") else DFMT_GUID)
    blk.items = _dfmt_items(value)
    _insert_item(ctl, blk, "DataFormat")


def _append(ctl, item):
    """a property or block after the others (before the controls inside)"""
    ctl.items.append(item)
    pos = next((i for i, (k, _) in enumerate(ctl.order) if k == "c"), len(ctl.order))
    ctl.order.insert(pos, ("p", item))


def _sub(blk, name):
    return next((x for x in blk.items if isinstance(x, PropertyBlock) and x.name.lower() == name.lower()), None) \
        if blk is not None else None


def grid_columns(ctl):
    """the columns of the design of a DataGrid (ColumnCount, ColumnNN, Split0): [{Caption, DataField, Width, Alignment,
    Locked, Visible, DataFormat}]"""
    n = ctl.get("ColumnCount")
    if n is None:
        return []
    split = ctl.block("Split0")
    out = []
    for i in range(int(_num(n))):
        b, lay = ctl.block(f"Column{i:02d}"), _sub(split, f"Column{i:02d}")
        c = {"Caption": "", "DataField": "", "Width": None, "Alignment": 0, "Locked": False, "Visible": True,
             "DataFormat": None}
        if b is not None:
            c["Caption"], c["DataField"] = str(b.get("Caption", "") or ""), str(b.get("DataField", "") or "")
            df = _sub(b, "DataFormat")
            c["DataFormat"] = _dfmt_dict(df) if df is not None else None
        if lay is not None:
            w = lay.get("ColumnWidth")
            c["Width"] = float(_num(w)) if w is not None else None
            c["Alignment"] = int(_num(lay.get("Alignment", 0)))
            c["Locked"] = bool(_num(lay.get("Locked", 0)))
            vis = lay.get("Object.Visible", lay.get("Visible"))
            c["Visible"] = True if vis is None else bool(_num(vis))
        out.append(c)
    return out


def set_grid_columns(ctl, cols):
    """write the columns of a DataGrid as VB6 does: ColumnCount, ColumnNN (DataField, Caption, DataFormat), SplitCount,
    Split0 (its own properties stay; ColumnNN: Alignment, ColumnWidth, Locked, Object.Visible); [] takes them away"""
    import re
    split = ctl.block("Split0")
    keep = [it for it in (split.items if split is not None else [])
            if not (isinstance(it, PropertyBlock) and re.fullmatch(r"column\d+", it.name.lower()))]
    remove(ctl, "ColumnCount")
    remove(ctl, "SplitCount")
    for it in list(ctl.items):
        if isinstance(it, PropertyBlock) and (re.fullmatch(r"column\d+", it.name.lower()) or it.name.lower() == "split0"):
            remove(ctl, it.name, block=True)
    if not cols:
        return
    _append(ctl, Property("ColumnCount", len(cols)))
    for i, c in enumerate(cols):
        b = PropertyBlock(f"Column{i:02d}")
        df = PropertyBlock("DataFormat", DFMT_GUID)
        df.items = _dfmt_items(c.get("DataFormat") or {"Type": 0, "Format": "", "LCID": 1033})
        b.items = [Property("DataField", str(c.get("DataField") or "")), Property("Caption", str(c.get("Caption") or "")), df]
        _append(ctl, b)
    _append(ctl, Property("SplitCount", 1))
    sp = PropertyBlock("Split0")
    sp.items = keep
    for i, c in enumerate(cols):
        cb = PropertyBlock(f"Column{i:02d}")
        if int(c.get("Alignment", 0) or 0):
            cb.items.append(Property("Alignment", int(c["Alignment"])))
        if c.get("Width") is not None:
            w = float(c["Width"])
            cb.items.append(Property("ColumnWidth", int(w) if w == int(w) else w))
        if c.get("Locked"):
            cb.items.append(Property("Locked", True))
        if not c.get("Visible", True):
            cb.items.append(Property("Object.Visible", False))
        sp.items.append(cb)
    _append(ctl, sp)


# ---------------------------------------------------------------- the bands of a CoolBar (ComCtl3)
# the properties of band N in the .frm, in the order the VB6 IDE writes them (VB6 oracles f_cool2, f_cool6), with their
# defaults (None: always written); Image is written as a string ("1"), BandPicture goes to the .frx
BAND_PROPS = [("BandForeColor", "ForeColor", "color", 0x80000012), ("BandBackColor", "BackColor", "color", 0x8000000F),
              ("Caption", "Caption", "string", ""), ("Image", "Image", "image", 0), ("Child", "Child", "string", ""),
              ("MinWidth", "MinWidth", "int", 0), ("MinHeight", "MinHeight", "int", None),
              ("Width", "Width", "int", 0), ("Key", "Key", "string", ""), ("BandPicture", "Picture", "picture", None),
              ("FixedBackground", "FixedBackground", "bool", True), ("UseCoolbarColors", "UseCoolbarColors", "bool", True),
              ("UseCoolbarPicture", "UseCoolbarPicture", "bool", True), ("NewRow", "NewRow", "bool", None),
              ("AllowVertical", "AllowVertical", "bool", True), ("Visible", "Visible", "bool", True),
              ("BandStyle", "Style", "int", 0), ("BandTag", "Tag", "string", ""),
              ("BandEmbossHighlight", "EmbossHighlight", "color", 0x80000014),
              ("BandEmbossShadow", "EmbossShadow", "color", 0x8000000F),
              ("BandEmbossPicture", "EmbossPicture", "bool", False)]
BAND_DEFAULT = {key: (360 if key == "MinHeight" else False if key == "NewRow" else d) for _, key, _, d in BAND_PROPS}


def _band_prop_re():
    import re
    names = "|".join(f for f, _, _, _ in BAND_PROPS)
    return re.compile(rf"^({names})(\d+)$", re.I)


def cool_bands(ctl):
    """the bands of a CoolBar: [{Caption, Child, MinWidth, MinHeight, Width, Key, NewRow, Visible, Style, Tag ...}]
    (sizes in twips); BandCount is 3 when the .frm does not have it"""
    n = int(_num(ctl.get("BandCount", 3)))
    out = []
    for i in range(1, n + 1):
        b = dict(BAND_DEFAULT)
        for frm, key, kind, _ in BAND_PROPS:
            if kind == "picture":
                pic = ctl.prop(f"{frm}{i}")
                if pic is not None:
                    b[key] = pic.value
                continue
            v = ctl.get(f"{frm}{i}")
            if v is None:
                continue
            if kind == "image":
                b[key] = _num(v) if str(v).strip().lstrip("-").isdigit() else str(v)
            elif kind == "bool":
                b[key] = bool(_num(v))
            elif kind == "color":
                b[key] = int(_num(v)) & 0xFFFFFFFF
            elif kind == "int":
                b[key] = _num(v)
            else:
                b[key] = str(v)
        out.append(b)
    return out


def set_cool_bands(ctl, bands):
    """write the bands of a CoolBar as the VB6 IDE does: BandCount, then the properties of each band"""
    pat = _band_prop_re()
    for it in list(ctl.items):
        if isinstance(it, Property) and (pat.match(it.name) or it.name.lower() == "bandcount"):
            ctl.items.remove(it)
            ctl.order = [(k, x) for k, x in ctl.order if x is not it]
    if len(bands) != 3:                          # 3 is what a CoolBar has when the .frm does not say (f_cool6)
        _append(ctl, Property("BandCount", len(bands)))
    for i, b in enumerate(bands, 1):
        for frm, key, kind, d in BAND_PROPS:
            v = b.get(key, BAND_DEFAULT[key])
            if kind == "picture":
                if v is not None:
                    _append(ctl, Property(f"{frm}{i}", v))
                continue
            if d is not None and v == d:
                continue
            if kind == "image":
                v = str(int(v)) if isinstance(v, (int, float)) else str(v)
            elif kind == "color":
                v = VBColor(v)
            elif kind == "int":
                v = int(v) if float(v) == int(v) else float(v)
            elif kind == "string":
                v = str(v)
            _append(ctl, Property(f"{frm}{i}", v))


# MS Sans Serif 8, the font of the rebar (runtime/forms/ocx_cool.inc)
_CB_TEXTW = dict(zip(" !\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~",
                     [3, 3, 4, 6, 6, 9, 8, 2, 4, 4, 6, 6, 3, 4, 3, 3] + [6] * 10 + [3, 3, 6, 6, 6, 6, 11] +
                     [7, 7, 7, 7, 6, 6, 8, 7, 3, 5, 7, 6, 9, 7, 8, 7, 8, 7, 7, 7, 7, 7, 11, 7, 7, 7] + [3, 3, 3, 6, 6, 4] +
                     [6, 6, 5, 6, 6, 3, 6, 5, 2, 2, 5, 2, 8, 6, 6, 6, 6, 4, 5, 3, 6, 6, 8, 5, 6, 5] + [4, 3, 4, 7]))


def cool_layout(bands, width_px, fixedorder=False, borders=True, variant=True, has_child=None):
    """the layout of the runtime (ocx_cool.inc) at design time: [(x, y, w, h, header, gripper) or None] for the bands
    (pixels in the client area), and the height of the CoolBar in pixels (None: no band shown)"""
    W = max(0, width_px - 4)
    vb = [i for i, b in enumerate(bands) if b.get("Visible", True)]
    nonfixed = sum(1 for i in vb if not bands[i].get("Style", 0))
    sep = 2 if borders else 0
    head, minb, need, grips = {}, {}, {}, {}
    for k, i in enumerate(vb):
        b = bands[i]
        grip = not b.get("Style", 0) and nonfixed > 1 and not (fixedorder and k == 0)
        cap = str(b.get("Caption", "") or "")
        child = has_child(b) if has_child else bool(b.get("Child"))
        h = 9 if grip else (4 if cap else 0)
        if cap:
            h += sum(_CB_TEXTW.get(ch, 7) for ch in cap) + (6 if child else 5)
        head[i], grips[i] = h, grip
        mw = int(_num(b.get("MinWidth", 0))) // 15
        minb[i] = h + (mw + 4 if mw > 0 else 0)
        need[i] = int(_num(b.get("Width", 0))) // 15 if b.get("Style", 0) else minb[i]
    rows, used = [], 0
    for i in vb:
        if not rows or bands[i].get("NewRow") or (used + need[i] > W and rows[-1]):
            rows.append([])
            used = 0
        rows[-1].append(i)
        used += need[i] + sep
    geo = [None] * len(bands)
    wd = {}
    for r in rows:
        tot = sep * (len(r) - 1)
        for i in r:
            cx = int(_num(bands[i].get("Width", 0))) // 15
            wd[i] = cx if bands[i].get("Style", 0) else max(cx, minb[i])
            tot += wd[i]
        if tot < W:                             # the last band with a Width of its own, else the first that can be sized
            g = next((i for i in reversed(r) if not bands[i].get("Style", 0) and int(_num(bands[i].get("Width", 0))) // 15 > 0),
                     next((i for i in r if not bands[i].get("Style", 0)), r[-1]))
            wd[g] += W - tot
        else:
            ex = tot - W
            for i in reversed(r):
                if ex <= 0:
                    break
                if bands[i].get("Style", 0):
                    continue
                d = min(ex, wd[i] - minb[i])
                if d > 0:
                    wd[i] -= d
                    ex -= d
    rh = [max(int(_num(bands[i].get("MinHeight", 360))) // 15 for i in r) for r in rows]
    if not variant and rh:
        rh = [max(rh)] * len(rh)
    y = 0
    for n, r in enumerate(rows):
        x = 0
        for i in r:
            geo[i] = (x, y, wd[i], rh[n], head[i], grips[i])
            x += wd[i] + sep
        y += rh[n] + (sep if n + 1 < len(rows) else 0)
    return geo, (y + 4 if rows else None)


def cool_children(ctl):
    """the controls inside a CoolBar, by the names a band gives them ("Text1", "tbar (0)")"""
    out = {}
    for k, c in getattr(ctl, "order", []):
        if k != "c":
            continue
        ix = c.get("Index")
        out[(f"{c.name} ({int(_num(ix))})" if ix is not None else c.name).lower()] = c
        if ix is not None:
            out[f"{c.name}({int(_num(ix))})".lower()] = c
    return out


def cool_place(ctl):
    """the CoolBar of the design as the runtime lays it out: its height and the children of its bands"""
    bands = cool_bands(ctl)
    kids = cool_children(ctl)
    width = int(_num(ctl.get("Width", 0))) // 15
    geo, h = cool_layout(bands, width, bool(_num(ctl.get("FixedOrder", 0))), bool(_num(ctl.get("BandBorders", -1))),
                         bool(_num(ctl.get("VariantHeight", -1))),
                         has_child=lambda b: str(b.get("Child", "") or "").lower() in kids)
    for b, g in zip(bands, geo):
        c = kids.get(str(b.get("Child", "") or "").lower())
        if c is None or g is None:
            continue
        x, y, w, rh, hd = g[:5]
        ch = int(_num(c.get("Height", 0))) // 15
        if not (getattr(c, "type_name", "") in ("ComboBox", "DriveListBox") and int(_num(c.get("Style", 0))) != 1):
            ch = int(_num(b.get("MinHeight", 360))) // 15
        for name, v in (("Left", (x + hd) * 15), ("Top", (y + (rh - ch) // 2) * 15),
                        ("Width", max(0, w - hd - 4) * 15), ("Height", ch * 15)):
            put(c, name, v)
    if h is not None:
        put(ctl, "Height", h * 15)
        if ctl.get("_ExtentY") is not None:
            from .ocxlib import extent
            put(ctl, "_ExtentY", extent(h * 15))
        if ctl.get("_CBHeight") is not None:
            put(ctl, "_CBHeight", h * 15)


# ---------------------------------------------------------------- DataRepeater (phase 7.5)
# the .frm as the VB6 IDE writes it (VB6 oracle f_drep3): after the other properties
#     BeginProperty RepeatedControlName {21FC0FC0-...}   _StreamID, _Version 65536, Name = "Proj.Ctl"
#     RepeaterBindings= n
#     BeginProperty RepeaterBinding(i) {7D21A594-...}    _StreamID, _Version 65536, PropertyName, DataField[, Key]
DREP_NAME_GUID = "{21FC0FC0-1E5C-11D1-A327-00AA00688B10}"
DREP_BIND_GUID = "{7D21A594-FC9B-11D0-A320-00AA00688B10}"
DREP_STREAM = -1412567295


def _blocks(ctl, pred):
    return [it for it in ctl.items if isinstance(it, PropertyBlock) and pred(it.name.lower())]


def drep_name(ctl):
    """the RepeatedControlName of a DataRepeater ("" when there is none)"""
    b = _blocks(ctl, lambda n: n == "repeatedcontrolname")
    if not b:
        return ""
    v = next((x.value for x in b[0].items if isinstance(x, Property) and x.name.lower() == "name"), "")
    return str(v or "")


def drep_bindings(ctl):
    """its RepeaterBindings: [{PropertyName, DataField, Key}] in the order of the .frm"""
    import re
    out = []
    blks = _blocks(ctl, lambda n: re.fullmatch(r"repeaterbinding\(\d+\)", n) is not None)
    blks.sort(key=lambda b: int(re.search(r"\d+", b.name).group(0)))
    for b in blks:
        d = {"PropertyName": "", "DataField": "", "Key": ""}
        for x in b.items:
            if isinstance(x, Property) and x.name in d:
                d[x.name] = str(x.value or "")
        out.append(d)
    return out


def set_drep(ctl, name=None, binds=None):
    """write RepeatedControlName and the RepeaterBindings again (None: as they are), after the other properties"""
    import re
    name = drep_name(ctl) if name is None else str(name or "")
    binds = drep_bindings(ctl) if binds is None else list(binds)
    for it in list(ctl.items):
        n = it.name.lower()
        if (isinstance(it, PropertyBlock) and (n == "repeatedcontrolname" or re.fullmatch(r"repeaterbinding\(\d+\)", n))) \
                or (isinstance(it, Property) and n == "repeaterbindings"):
            ctl.items.remove(it)
            ctl.order = [(k, x) for k, x in ctl.order if x is not it]
    if name:
        b = PropertyBlock("RepeatedControlName", DREP_NAME_GUID)
        b.items = [Property("_StreamID", DREP_STREAM), Property("_Version", 65536), Property("Name", name)]
        _append(ctl, b)
    binds = [d for d in binds if str(d.get("PropertyName", "") or "").strip()]
    if binds:
        _append(ctl, Property("RepeaterBindings", len(binds)))
        for i, d in enumerate(binds):
            b = PropertyBlock(f"RepeaterBinding({i})", DREP_BIND_GUID)
            b.items = [Property("_StreamID", DREP_STREAM), Property("_Version", 65536),
                       Property("PropertyName", str(d.get("PropertyName", ""))),
                       Property("DataField", str(d.get("DataField", "") or ""))]
            if str(d.get("Key", "") or ""):
                b.items.append(Property("Key", str(d["Key"])))
            _append(ctl, b)


def _special_get(ctl, p):
    """(True, value) for a property of the blob or of the Bindings, (False, None) otherwise"""
    from lbvc.project.oleblob import get_prop, bindings
    t = getattr(ctl, "type_name", "")
    if p.kind == "drepname":
        return True, drep_name(ctl)
    if p.kind == "drepbinds":
        return True, drep_bindings(ctl)
    if p.kind == "dataformat":
        return True, dataformat_of(ctl)
    if p.kind == "dgcolumns":
        return True, grid_columns(ctl)
    if p.kind == "cbbands":
        return True, cool_bands(ctl)
    k = _binding_dispid(ctl, p)
    if k is not None:
        for d, n in bindings(_frx_bytes(ctl, "Bindings")):
            if d == k:
                return True, n
        v = ctl.get(p.name)                     # (a form that names it in the .frm)
        return True, v if isinstance(v, str) else ""
    if t == "DBGrid" and p.kind == "dbgcolumns":   # (columns, _LayoutType)
        from lbvc.project.oleblob import get_columns
        b = grid_blob(ctl)
        return True, (get_columns(b), get_prop(b, "_LayoutType") or 0)
    if t == "DBGrid" and p.name in BLOB_PROPS:
        v = get_prop(grid_blob(ctl), p.name)
        if v is None:
            return True, p.default
        if p.kind == "bool":
            return True, bool(v)
        if p.kind == "color":
            return True, int(v) & 0xFFFFFFFF
        return True, v
    return False, None


def _special_set(ctl, p, value):
    from lbvc.project.oleblob import set_prop, bindings_set
    t = getattr(ctl, "type_name", "")
    if p.kind == "drepname":
        set_drep(ctl, name=str(value or "").strip())
        return True
    if p.kind == "drepbinds":
        set_drep(ctl, binds=value or [])
        return True
    if p.kind == "dataformat":
        set_dataformat(ctl, value)
        return True
    if p.kind == "dgcolumns":
        set_grid_columns(ctl, value)
        return True
    if p.kind == "cbbands":
        set_cool_bands(ctl, value)
        cool_place(ctl)
        return True
    k = _binding_dispid(ctl, p)
    if k is not None:
        b = bindings_set(_frx_bytes(ctl, "Bindings"), k, str(value or ""))
        if b:
            put(ctl, "Bindings", FrxData("raw", b))
        else:
            remove(ctl, "Bindings")
        remove(ctl, p.name)
        return True
    if t == "DBGrid" and p.kind == "dbgcolumns":
        from lbvc.project.oleblob import set_columns
        cols, layout = value
        put(ctl, "OleObjectBlob", FrxData("raw", set_columns(grid_blob(ctl), cols, layout)))
        return True
    if t == "DBGrid" and p.name in BLOB_PROPS:
        if p.kind == "bool":
            value = -1 if value else 0
        elif p.kind == "color":
            value = int(value) - (1 << 32) if int(value) & 0x80000000 else int(value)
        put(ctl, "OleObjectBlob", FrxData("raw", set_prop(grid_blob(ctl), p.name, value)))
        return True
    return False


def value_of(ctl, p, form_root=None):
    """the value of a property as the Properties window shows it"""
    if getattr(ctl, "type_name", "") == "MSComm" and p.name in ("Settings", "ParityReplace"):
        return _comm_value(ctl, p)
    ok, v = _special_get(ctl, p)
    if ok:
        return v
    p = frm_prop(ctl, p)
    if p.kind == "name":
        return ctl.name
    if p.kind == "formpos":
        return outer_rect(ctl)[{"Left": 0, "Top": 1, "Width": 2, "Height": 3}[p.name]]
    if p.kind == "font":
        return font_of(ctl, form_root)
    v = ctl.get(p.name)
    if p.kind == "index":
        return v if isinstance(v, int) else None
    if v is None:
        if p.kind == "scale":
            return _num(ctl.get("Width" if p.name == "ScaleWidth" else "Height", 0)) if p.default is None else p.default
        return p.default
    if p.kind == "bool":
        return bool(_num(v))
    if p.kind in ("enum", "int", "pos", "scale", "num"):
        return _num(v)
    if p.kind == "color":
        return int(v) & 0xFFFFFFFF if not isinstance(v, str) else p.default
    return v


def frame_of(form, has_menu=None):
    """(left, top, right, bottom) of the frame of a form in twips"""
    bs = _num(form.get("BorderStyle", 2))
    l, t, r, b = FRAMES.get(bs, FRAMES[2])
    if has_menu is None:
        has_menu = any(c.type_name == "Menu" for c in form.children)
    if has_menu and bs != 0:
        t += 20
    return l * 15, t * 15, r * 15, b * 15


def outer_rect(form, frame=None):
    """Left, Top, Width, Height of a form from its Client* properties"""
    l, t, r, b = frame or frame_of(form)
    cl, ct = _num(form.get("ClientLeft", 60)), _num(form.get("ClientTop", 345))
    cw, ch = _num(form.get("ClientWidth", 4680)), _num(form.get("ClientHeight", 3195))
    return cl - l, ct - t, cw + l + r, ch + t + b


def set_outer_rect(form, x, y, w, h, frame=None):
    """the Client* properties for an outer rectangle (what VB6 writes)"""
    l, t, r, b = frame or frame_of(form)
    put(form, "ClientHeight", max(0, int(round(h)) - t - b))
    put(form, "ClientLeft", int(round(x)) + l)
    put(form, "ClientTop", int(round(y)) + t)
    put(form, "ClientWidth", max(0, int(round(w)) - l - r))


def font_of(ctl, form_root=None):
    """the Font of a control: its BeginProperty Font, else the one of the form, else MS Sans Serif"""
    if getattr(ctl, "type_name", "") == "DBGrid":   # in its OleObjectBlob
        from lbvc.project.oleblob import get_prop
        f = get_prop(grid_blob(ctl), "Font") or {}
        return {"Name": f.get("Name", "MS Sans Serif"), "Size": f.get("Size", 8.25), "Charset": f.get("Charset"),
                "Weight": f.get("Weight", 400), "Underline": bool(f.get("Underline")), "Italic": bool(f.get("Italic")),
                "Strikethrough": bool(f.get("Strikethrough"))}
    blk = ctl.block("Font")
    if blk is None and form_root is not None and form_root is not ctl:
        blk = form_root.block("Font")
    f = {"Name": "MS Sans Serif", "Size": 8.25, "Charset": None, "Weight": 400, "Underline": False,
         "Italic": False, "Strikethrough": False}
    if blk is not None:
        for k in f:
            v = blk.get(k)
            if v is not None:
                f[k] = bool(_num(v)) if k in ("Underline", "Italic", "Strikethrough") else v
    return f


# ---------------------------------------------------------------- showing values
def display(p, v):
    """the text of the value column"""
    if v is None and p.kind in ("bool", "num", "color", "string", "enum"):
        return ""                               # a property of a user control not in the .frm: its default is its own
    if p.kind == "bool":
        return "True" if v else "False"
    if p.kind == "enum":
        t = enum_label(p, v)
        return f"{v} - {t}" if t is not None else str(v)
    if p.kind == "color":
        return f"&H{int(v) & 0xFFFFFFFF:08X}&"
    if p.kind == "font":
        return str(v["Name"])
    if p.kind == "picture":
        if not v:
            return "(None)"
        kind = sniff_image(bytes(v))
        return "(Icon)" if kind in ("ico", "cur") else "(Metafile)" if kind in ("wmf", "emf") else "(Bitmap)"
    if p.kind == "list":
        return "(List)"
    if p.kind == "dgcolumns":                    # the columns of the design of a DataGrid
        return f"({len(v)} columns)" if v else "(None)"
    if p.kind == "cbbands":                      # the bands of a CoolBar
        return f"({len(v)} band{'s' if len(v) != 1 else ''})" if v else "(None)"
    if p.kind == "drepbinds":                    # the RepeaterBindings of a DataRepeater
        return f"({len(v)} binding{'s' if len(v) != 1 else ''})" if v else "(None)"
    if p.kind == "drepname":
        return str(v or "")
    if p.kind == "dataformat":                   # the category of the Format page
        if not v:
            return ""
        from .dfmtdialog import category_of
        f = str(v.get("Format", "") or "")
        return category_of(v) + (f" ({f})" if f else "")
    if p.kind == "dbgcolumns":                   # (columns, _LayoutType)
        cols, layout = v if isinstance(v, tuple) else ([], 0)
        return "(Default layout)" if not layout else f"({len(cols)} columns)"
    if p.kind == "shortcut":
        return TEXT_OF_SHORTCUT.get(str(v).upper(), "(None)") if v else "(None)"
    if p.kind == "readonly":
        return ""
    if v is None:
        return ""
    if isinstance(v, float):
        s = repr(v)
        return s[:-2] if s.endswith(".0") else s
    if isinstance(v, str) and ("\r" in v or "\n" in v):
        return v.replace("\r\n", " ").replace("\n", " ")
    return str(v)


def parse_input(p, text):
    """the text the user typed -> the value, or raise ValueError ("Invalid property value")"""
    t = text.strip()
    if p.kind == "bool":
        if t.lower() in ("true", "-1", "1"):
            return True
        if t.lower() in ("false", "0"):
            return False
        raise ValueError
    if p.kind == "enum":
        head = t.split("-")[0].strip() if t else ""
        try:
            n = int(head)
        except ValueError:
            for k, lab in enum_items(p):
                if lab.lower() == t.lower():
                    return k
            raise ValueError
        if n not in [k for k, _ in enum_items(p)]:
            raise ValueError
        return n
    if p.kind in ("int",):
        n = float(t)
        if n != int(n) or not -2147483648 <= n <= 2147483647:
            raise ValueError
        return int(n)
    if p.kind in ("pos", "scale", "formpos", "num"):
        n = float(t)
        return int(n) if n == int(n) else n
    if p.kind == "color":
        u = t.upper()
        if u.startswith("&H"):
            return int(u[2:].rstrip("&"), 16) & 0xFFFFFFFF
        return int(float(t)) & 0xFFFFFFFF
    if p.kind == "index":
        if t == "":
            return None
        n = int(t)
        if not 0 <= n <= 32767:
            raise ValueError
        return n
    if p.kind == "string" and p.name == "PasswordChar":
        return t[:1]
    return text


# ---------------------------------------------------------------- writing values
def _replace(ctl, old, new):
    i = ctl.items.index(old)
    ctl.items[i] = new
    ctl.order = [(k, new if it is old else it) for k, it in ctl.order]


def put(ctl, name, value, comment=None):
    """set a property where VB6 writes it (alphabetical among the properties, before the
    controls); an enum carries the name of its value as a comment"""
    p = ctl.prop(name)
    if p is not None:
        if comment is not None or p.comment is not None and not isinstance(value, bool):
            _replace(ctl, p, Property(name, value, comment=comment))
        else:
            p.value = value
        return
    p = Property(name, value, comment=comment)
    _insert_item(ctl, p, name)


def _bag_of(ctl):
    """the order in which an .ocx control writes its own properties (ocxlib), or None"""
    from .ocxlib import control_of
    lc = control_of(getattr(ctl, "cls", ""))
    return lc[1].bag if lc else None


def _insert_item(ctl, item, name):
    bag = _bag_of(ctl)
    if bag and name in bag:                      # after _Version, in the order of the control
        ctl.items.append(item)
        rank = bag.index(name)
        pos = None
        for i, (k, it) in enumerate(ctl.order):
            if k == "c":
                pos = i if pos is None else pos
                break
            n = getattr(it, "name", "")
            if n == "_Version" or (n in bag and bag.index(n) < rank):
                pos = i + 1
        ctl.order.insert(len(ctl.order) if pos is None else pos, ("p", item))
        return
    ctl.items.append(item)
    key = name.lower()
    pos = None
    for i, (k, it) in enumerate(ctl.order):
        if k == "c" or it.name.lower() > key:
            pos = i
            break
    ctl.order.insert(len(ctl.order) if pos is None else pos, ("p", item))


def remove(ctl, name, block=False):
    for it in list(ctl.items):
        if it.name.lower() == name.lower() and isinstance(it, PropertyBlock if block else Property):
            ctl.items.remove(it)
            ctl.order = [(k, x) for k, x in ctl.order if x is not it]


def _same(p, a, b):
    if p.kind == "color":
        return (int(a) & 0xFFFFFFFF) == (int(b) & 0xFFFFFFFF)
    if p.kind == "bool":
        return bool(a) == bool(b)
    return a == b


COMM_PARITY = "noems"                            # ParitySetting: the DCB numbers


def _comm_settings(ctl, value):
    """MSComm Settings "19200,e,7,1" -> BaudRate, ParitySetting, DataBits, StopBits (VB6 oracle d_dev)"""
    f = [x.strip().lower() for x in str(value).split(",")] + ["", "", "", ""]
    parts = (("BaudRate", f[0], "9600"), ("ParitySetting", str(COMM_PARITY.find(f[1][:1])) if f[1] else "0", "0"),
             ("DataBits", f[2], "8"), ("StopBits", {"1": "0", "1.5": "1", "2": "2"}.get(f[3], "0"), "0"))
    for name, v, default in parts:
        if not v or v == default or v == "-1":
            remove(ctl, name)
        else:
            put(ctl, name, int(v))


def _comm_value(ctl, p):
    if p.name == "Settings":
        par = int(_num(ctl.get("ParitySetting", 0)))
        stop = {0: "1", 1: "1.5", 2: "2"}.get(int(_num(ctl.get("StopBits", 0))), "1")
        return f"{int(_num(ctl.get('BaudRate', 9600)))},{COMM_PARITY[par] if 0 <= par < 5 else 'n'}," \
               f"{int(_num(ctl.get('DataBits', 8)))},{stop}"
    v = ctl.get("ParityReplace")
    return p.default if v is None else (chr(int(_num(v))) if int(_num(v)) else "")


def apply(ctl, p, value):
    """write a value as VB6 would: left out when it is the default, the VB6 layout otherwise"""
    if _special_set(ctl, p, value):
        return
    if getattr(ctl, "type_name", "") == "MSComm" and p.name in ("Settings", "ParityReplace"):
        if p.name == "Settings":
            _comm_settings(ctl, value)
        elif str(value) == "?":
            remove(ctl, "ParityReplace")
        else:
            put(ctl, "ParityReplace", ord(str(value)[:1]) if str(value) else 0)
        return
    p = frm_prop(ctl, p)
    if p.kind == "index":
        if value is None:
            remove(ctl, "Index")
        else:
            put(ctl, "Index", int(value))
        return
    if p.kind == "shortcut":
        if value:
            put(ctl, "Shortcut", Raw(value))
        else:
            remove(ctl, "Shortcut")
        return
    if p.kind == "picture":                      # the bytes of the image file, or None
        if value:
            put(ctl, p.name, FrxData("picture", bytes(value)))
        else:
            remove(ctl, p.name)
        return
    if p.kind == "list":                         # List / ItemData
        if value:
            kind = "itemdata" if p.name == "ItemData" else "list"
            put(ctl, p.name, FrxData(kind, [int(x) for x in value] if kind == "itemdata" else [str(x) for x in value]))
        else:
            remove(ctl, p.name)
        return
    if p.kind in ("string", "text") and needs_frx(value):
        put(ctl, p.name, FrxData("text" if p.name == "Text" else "string", value))
        return
    if not p.always and p.default is not None and _same(p, value, p.default):
        remove(ctl, p.name)
        return
    if p.kind == "bool":
        put(ctl, p.name, bool(value))
    elif p.kind == "enum":
        bag = _bag_of(ctl)
        put(ctl, p.name, int(value), None if bag and p.name in bag else frm_comment(p, int(value)))
    elif p.kind == "color" and getattr(ctl, "type_name", "") == "DataRepeater":
        v = int(value) & 0xFFFFFFFF                 # a number, as the DataRepeater writes it (VB6 oracle f_drep3)
        put(ctl, p.name, v - 0x100000000 if v >= 0x80000000 else v)
    elif p.kind == "color":
        put(ctl, p.name, VBColor(value))
    elif p.kind in ("int", "pos", "scale", "num"):
        put(ctl, p.name, value if isinstance(value, float) else int(value))
    else:
        put(ctl, p.name, str(value))


# ---------------------------------------------------------------- fonts
CHARSET_OF_CODEPAGE = {"cp1252": 0, "cp1253": 161, "cp1250": 238, "cp1251": 204, "cp1254": 162, "cp1255": 177,
                       "cp1256": 178, "cp1257": 186, "cp1258": 163, "cp874": 222, "cp932": 128, "cp936": 134,
                       "cp949": 129, "cp950": 136}


def font_size(points):
    """VB6 keeps the sizes that a whole number of pixels gives at 96 dpi: 8 -> 8.25, 10 -> 9.75"""
    px = max(1, round(float(points) * 4 / 3))
    s = px * 0.75
    return int(s) if s == int(s) else s


def set_font(ctl, f, default_font):
    """write BeginProperty Font, or leave it out when the font is the default one"""
    if getattr(ctl, "type_name", "") == "DBGrid":   # in its OleObjectBlob
        from lbvc.project.oleblob import set_prop
        g = dict(f, Size=font_size(f["Size"]), Charset=int(f["Charset"] or 0))
        put(ctl, "OleObjectBlob", FrxData("raw", set_prop(grid_blob(ctl), "Font", g)))
        return
    remove(ctl, "Font", block=True)
    keys = ("Name", "Size", "Weight", "Underline", "Italic", "Strikethrough")
    if all(_fcmp(f[k], default_font[k]) for k in keys):
        return
    blk = PropertyBlock("Font")
    blk.items = [Property("Name", str(f["Name"])), Property("Size", font_size(f["Size"])),
                 Property("Charset", int(f["Charset"] or 0)), Property("Weight", int(f["Weight"])),
                 Property("Underline", bool(f["Underline"])), Property("Italic", bool(f["Italic"])),
                 Property("Strikethrough", bool(f["Strikethrough"]))]
    _insert_item(ctl, blk, "Font")


def _fcmp(a, b):
    if isinstance(a, (int, float)) or isinstance(b, (int, float)):
        try:
            return float(a) == float(b)
        except (TypeError, ValueError):
            return False
    return str(a).lower() == str(b).lower()


def clear_raw(c):
    """a control written again in the layout of VB6 (after a move to another depth)"""
    c.raw_begin = c.raw_end = None
    for it in c.items:
        if isinstance(it, Property):
            it.raw = None
        elif isinstance(it, PropertyBlock):
            it.raw_begin = it.raw_end = None
            for q in it.items:
                q.raw = None
    for ch in c.children:
        clear_raw(ch)
