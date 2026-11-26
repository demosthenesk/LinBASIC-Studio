"""ocxgrids.py - the property pages of the grids (phase 11, stage 2c): the Columns, Layout, Splits and Format pages
of a DataGrid, the Columns, Layout and Splits pages of a DBGrid, the Format page of a DataRepeater.

The pages of a grid share its columns: a change on one page is on the others when they are shown. A DataGrid keeps
them in the .frm (vbprops.grid_columns / set_grid_columns: ColumnNN, Split0 and its ColumnNN; the properties of the
split in Split0, vbprops.grid_split), a DBGrid in the OleObjectBlob of its .frx (lbvc/project/oleblob.py: the records
of the Layout page from the VB6 oracle DEV-11 s_dbgbag). The Format page (StdFormat of msstkprp.dll) gives the
BeginProperty DataFormat of an item (dfmtdialog.py: the formats, in the notation of Format$); its boxes are where VB6
has them (VB6 oracle DEV-11 s_format).
Copyright (C) 2026 Demosthenes Koptsis. GPL-3.0-or-later."""
import copy

from PyQt6.QtCore import QLocale, QDate, QTime
from PyQt6.QtWidgets import (QLabel, QLineEdit, QComboBox, QCheckBox, QListWidget, QGroupBox, QPushButton)

from . import vbprops as V
from . import dfmtdialog as DF


def _ids(made, *ids):
    from .ocxpages import by_id
    return [by_id(made, i) for i in ids]


def _enum(text):
    from .ocxpages import enum_value
    v = enum_value(text)
    return 0 if v is None else v


def _pick(cb, n):
    """the item of a combo box whose number is n ("2 - dbgDarkGrayLine")"""
    cb.setCurrentIndex(next((i for i in range(cb.count()) if _enum(cb.itemText(i)) == int(n)), 0))


def _num_text(v, dec=3):
    t = f"{float(v):.{dec}f}".rstrip("0").rstrip(".")
    return t.replace(".", QLocale().decimalPoint())


def _float(s):
    return float(str(s).strip().replace(",", "."))


class Grid:
    """the columns and the split of a grid while its pages are open: one for each dialog"""

    def __init__(self, dlg):
        self.dlg, self.t = dlg, dlg.type_name
        self.refresh = []                       # what each page shows again when it comes in front
        if self.t == "DBGrid":
            cols, self.layout = dlg.value("(Columns)") or ([], 0)
            self.cols = copy.deepcopy(list(cols))
            from lbvc.project.oleblob import get_split
            self.split = get_split(V.grid_blob(dlg.ctl))
        else:
            self.cols = copy.deepcopy(list(dlg.value("(Columns)") or []))
            self.split = V.grid_split(dlg.ctl)
        dlg.tabs.currentChanged.connect(lambda _: [f() for f in self.refresh])

    @staticmethod
    def of(dlg):
        g = dlg.__dict__.get("grid")
        if g is None:
            g = dlg.__dict__["grid"] = Grid(dlg)
        return g

    def names(self):
        return [f"Column{i}" if self.t == "DBGrid" else f"Column {i}" for i in range(len(self.cols))] or \
            (["Column0"] if self.t == "DBGrid" else ["Column 0"])

    def changed(self):
        cols = copy.deepcopy(self.cols)
        self.dlg.changed("(Columns)", (cols, 4) if self.t == "DBGrid" else cols)

    def split_changed(self):
        dlg = self.dlg
        if self.t == "DBGrid":
            from lbvc.project.oleblob import set_split
            from lbvc.project.frx import FrxData

            def put(v):
                V.put(dlg.ctl, "OleObjectBlob", FrxData("raw", set_split(V.grid_blob(dlg.ctl), v)))
        else:
            def put(v):
                V.set_grid_split(dlg.ctl, v)
        dlg.raw_apply["(Split0)"] = put
        dlg.changed("(Split0)", dict(self.split))


def _item_combo(cb, g):
    cb.blockSignals(True)
    i = max(0, cb.currentIndex())
    cb.clear()
    cb.addItems(g.names())
    cb.setCurrentIndex(min(i, cb.count() - 1))
    cb.blockSignals(False)


def _fields(dlg):
    """the fields of the data source of the grid: ADO (DataGrid) or a Data control (DBGrid)"""
    if dlg.type_name != "DBGrid":
        from .ocxlists import data_fields
        return data_fields(dlg)
    import os
    from . import dbinfo
    d = dlg.designer
    src = str(dlg.value("DataSource") or "")
    if d is None or not src:
        return []
    ds = next((c for c in d.root.walk() if c.name.lower() == src.lower() and c.type_name == "Data"), None)
    if ds is None:
        return []
    mod = d.module
    path = getattr(mod, "path", None) or getattr(getattr(mod, "project", None), "path", None)
    try:
        db = dbinfo.resolve(str(ds.get("DatabaseName", "") or ""), os.path.dirname(path) if path else None)
        return dbinfo.fields(db, str(ds.get("RecordSource", "") or ""))
    except Exception:                           # noqa: BLE001 - no database here
        return []


NUMBER_FORMATS = ["General Number", "Currency", "Fixed", "Standard", "Percent", "Scientific", "Yes/No", "True/False",
                  "On/Off", "General Date", "Long Date", "Medium Date", "Short Date", "Long Time", "Medium Time",
                  "Short Time"]


# ---------------------------------------------------------------- Columns
def columns_page(dlg, key, w, made):
    """the Columns page: the column, its Caption and DataField (DBGrid: DefaultValue, NumberFormat too)"""
    g = Grid.of(dlg)
    db = g.t == "DBGrid"
    combo, cap, fld = _ids(made, *((8000, 5, 8) if db else (3018, 3019, 3020)))
    extra = dict(zip(("DefaultValue", "NumberFormat"), _ids(made, 9, 17))) if db else {}
    fld.setEditable(True)
    fld.clear()
    fld.addItems(_fields(dlg))
    if "NumberFormat" in extra:
        extra["NumberFormat"].setEditable(True)
        extra["NumberFormat"].clear()
        extra["NumberFormat"].addItems(NUMBER_FORMATS)
    busy = [False]

    def show(*_):
        busy[0] = True
        _item_combo(combo, g)
        i = combo.currentIndex()
        c = g.cols[i] if 0 <= i < len(g.cols) else None
        for wd in (cap, fld, *extra.values()):
            wd.setEnabled(c is not None)
        cap.setText(str(c.get("Caption", "")) if c else "")
        fld.setEditText(str(c.get("DataField", "")) if c else "")
        for k, wd in extra.items():
            if isinstance(wd, QComboBox):
                wd.setEditText(str(c.get(k, "")) if c else "")
            else:
                wd.setText(str(c.get(k, "")) if c else "")
        busy[0] = False

    def edit(k, v):
        i = combo.currentIndex()
        if busy[0] or not 0 <= i < len(g.cols):
            return
        g.cols[i][k] = v
        g.changed()
    combo.currentIndexChanged.connect(show)
    cap.textEdited.connect(lambda s: edit("Caption", s))
    fld.editTextChanged.connect(lambda s: edit("DataField", s))
    for k, wd in extra.items():
        (wd.editTextChanged if isinstance(wd, QComboBox) else wd.textEdited).connect(lambda s, k=k: edit(k, s))
    g.refresh.append(show)
    show()
    dlg.bound[key] = {"(Columns)": combo}
    dlg.unbound[key] = []


# ---------------------------------------------------------------- Layout
DEFAULT_WIDTH = 1514.835                        # twips: what the Layout page of a DataGrid shows for a column of its own


def layout_page(dlg, key, w, made):
    """the Layout page: the split, the column; Locked, AllowSizing, Visible, WrapText, Button, DividerStyle,
    Alignment, Width"""
    g = Grid.of(dlg)
    db = g.t == "DBGrid"
    ids = (8001, 8000, 16, 2, 22, 37, 40, 4, 1, 23) if db else (3023, 3024, 3025, 3026, 3027, 3028, 3029, 3030, 3031,
                                                                3032)
    split, combo, locked, sizing, vis, wrap, button, div, align, width = _ids(made, *ids)
    checks = {"Locked": locked, "AllowSizing": sizing, "Visible": vis, "WrapText": wrap, "Button": button}
    busy = [False]

    def cur():
        i = combo.currentIndex()
        return g.cols[i] if 0 <= i < len(g.cols) else None

    def show(*_):
        busy[0] = True
        split.clear()
        split.addItems(["Split0" if db else "Split 0"])
        _item_combo(combo, g)
        c = cur()
        for wd in (*checks.values(), div, align, width):
            wd.setEnabled(c is not None)
        if c is not None:
            for k, cb in checks.items():
                cb.setChecked(bool(c.get(k, k in ("AllowSizing", "Visible"))))
            _pick(div, c.get("DividerStyle", 2))
            al = c.get("Alignment", 3 if db else 0)
            _pick(align, 0 if db and al == 3 and c.get("AlignFlag") else al)      # (s_dbgal: 0 - Left)
            if db:
                width.setText(_num_text(float(c.get("Width") or 0) * 1440 / 2540))
            else:
                wv = c.get("Width")
                width.setText(_num_text(DEFAULT_WIDTH if wv is None else wv))
        busy[0] = False

    def edit(k, v):
        c = cur()
        if busy[0] or c is None:
            return
        if k == "DividerStyle" and db:          # the width of a DBGrid column counts its divider (s_dbgbag2)
            from lbvc.project.oleblob import DIVIDERS
            old, new = DIVIDERS.get(int(c.get("DividerStyle", 2)), DIVIDERS[2])[2], DIVIDERS.get(int(v), DIVIDERS[2])[2]
            c["Width"] = int(round(float(c.get("Width") or 0) + (new - old) * 2540 / 96))
        c[k] = v
        if k == "Alignment":
            c["AlignFlag"] = False
        g.changed()
        if k == "DividerStyle" and db:
            show()

    def width_typed(s):
        c = cur()
        if busy[0] or c is None:
            return
        try:
            tw = _float(s)
        except ValueError:
            return
        if db:
            from lbvc.project.oleblob import him_of_twips
            c["Width"] = him_of_twips(tw)
        else:
            c["Width"] = tw
        g.changed()
    combo.currentIndexChanged.connect(show)
    for k, cb in checks.items():
        cb.toggled.connect(lambda on, k=k: edit(k, on))
    div.activated.connect(lambda i: edit("DividerStyle", _enum(div.itemText(i))))
    align.activated.connect(lambda i: edit("Alignment", _enum(align.itemText(i))))
    width.textEdited.connect(width_typed)
    g.refresh.append(show)
    show()
    dlg.bound[key] = {"(Columns)": combo}
    dlg.unbound[key] = []


# ---------------------------------------------------------------- Splits
def splits_page(dlg, key, w, made):
    """the Splits page: Split0 - Locked, AllowFocus, AllowSizing, AllowRowSizing, RecordSelectors, ScrollGroup,
    ScrollBars, MarqueeStyle, Size, SizeMode"""
    g = Grid.of(dlg)
    db = g.t == "DBGrid"
    ids = (8001, 1, 5, 4, 15, 16, 3, 17, 31, 6, 7) if db else (3033, 3034, 3035, 3036, 3037, 3038, 3039, 3040, 3041,
                                                                3042, 3043)
    split, *boxes = _ids(made, *ids)
    names = ("Locked", "AllowFocus", "AllowSizing", "AllowRowSizing", "RecordSelectors", "ScrollGroup", "ScrollBars",
             "MarqueeStyle", "Size", "SizeMode")
    bx = dict(zip(names, boxes))
    defaults = dict(V.SPLIT_PROPS)
    split.clear()
    split.addItems(["Split0" if db else "Split 0"])
    mq = bx["MarqueeStyle"]
    if mq.count() < 7:                          # (the list of VB6 has dbgFloatingEditor too)
        mq.addItem("6 - Floating Editor" if db else "6 - dbgFloatingEditor")
    if bx["SizeMode"].count() < 2:
        bx["SizeMode"].addItem("1 - Exact" if db else "1 - dbgExact")
    busy = [False]

    def show(*_):
        busy[0] = True
        if db:                                  # (Properties window: AllowRowSizing ... of the grid are of Split0)
            pend = {k: dlg.pending[k] for k in ("AllowRowSizing", "RecordSelectors", "MarqueeStyle", "ScrollBars")
                    if k in dlg.pending}
            g.split.update(pend)
        for k, wd in bx.items():
            v = g.split.get(k, defaults.get(k))
            if isinstance(wd, QCheckBox):
                wd.setChecked(bool(v))
            elif isinstance(wd, QComboBox):
                _pick(wd, v)
            else:
                wd.setText(_num_text(v) if v is not None else "")
        busy[0] = False

    def edit(k, v):
        if busy[0]:
            return
        g.split[k] = v
        g.split_changed()
    for k, wd in bx.items():
        if isinstance(wd, QCheckBox):
            wd.toggled.connect(lambda on, k=k: edit(k, on))
        elif isinstance(wd, QComboBox):
            wd.activated.connect(lambda i, k=k, wd=wd: edit(k, _enum(wd.itemText(i))))
        else:
            def typed(s, k=k):
                try:
                    edit(k, _float(s) if k == "Size" else int(_float(s)))
                except ValueError:
                    pass
            wd.textEdited.connect(typed)
    g.refresh.append(show)
    show()
    dlg.bound[key] = {"(Split0)": split}
    dlg.unbound[key] = []


# ---------------------------------------------------------------- Format (StdFormat)
SAMPLE_NUMBER, SAMPLE_VALUE = 12345.6789, 35983.9079861     # what VB6 shows (s_format): 12346, 7/7/1998 9:47:30
TEXTS = {"General": "General causes no special formatting to occur.",
         "Checkbox": "This format allows a checkbox to be set properly when bound to a boolean column.",
         "Picture": "This format allows a picture to be read from and written to the database."}
CURRENCIES = ["$ English (United States)", "€ Euro", "£ English (United Kingdom)", "¥ Japanese (Japan)",
              "None"]
SHOWN = {"General": ("general",), "Number": ("sample", "dec", "thou", "neg", "fmt"),
         "Currency": ("sample", "dec", "neg", "fmt", "sym"), "Date": ("sample", "fmt", "list"),
         "Time": ("sample", "fmt", "list"), "Percentage": ("sample", "dec", "fmt"),
         "Scientific": ("sample", "dec", "fmt"), "Boolean": ("bool",), "Checkbox": ("check",),
         "Picture": ("picture",), "Custom": ("sample", "custom")}


def _qt_format(fmt, time=False):
    """a format of Format$ for a date or a time -> the notation of Qt (the sample)"""
    import re
    f = fmt.replace("AMPM", "AP").replace("nn", "mm")
    if not time:
        return f
    return re.sub(r"h+", lambda m: m.group(0), f)


def sample_of(cat, fmt):
    """the Sample of the page (a little of what Format$ does)"""
    import math
    if cat in ("Number", "Currency", "Percentage"):
        v = SAMPLE_NUMBER if cat == "Number" else SAMPLE_VALUE
        if cat == "Percentage":
            v *= 100
        head = fmt.split(";")[0]
        dec = DF._decimals(head)
        s = f"{v:,.{dec}f}" if "," in head else f"{v:.{dec}f}"
        pre = head.split("#")[0].split("0")[0].replace('"', "") if cat == "Currency" else ""
        return pre + s + ("%" if cat == "Percentage" else "")
    if cat == "Scientific":
        dec = DF._decimals(fmt.split("E")[0])
        m, e = f"{SAMPLE_VALUE:.{dec}E}".split("E")
        return f"{m}E{int(e):+03d}"
    if cat == "Date":
        return QDate(1899, 12, 30).addDays(int(SAMPLE_VALUE)).toString(_qt_format(fmt))
    if cat == "Time":
        secs = int(round((SAMPLE_VALUE - math.floor(SAMPLE_VALUE)) * 86400))
        return QTime(0, 0).addSecs(secs).toString(_qt_format(fmt, True))
    if cat == "Custom":
        return str(SAMPLE_VALUE) if not fmt else fmt
    return ""


def format_page(dlg, key, w, made):
    """Format (StdFormat): Format Item (the columns of a DataGrid, the bindings of a DataRepeater), Format Type and
    the boxes of the type; the DataFormat of the item"""
    from .ocxpages import UpDown
    repeater = dlg.type_name == "DataRepeater"
    item, cats = _ids(made, 3018, 3006)
    for it, wd in made:                         # (the text of General: a box of its own here)
        if it["id"] == 3019:
            wd.hide()
    x = {}

    def lab(name, text, r):
        wd = QLabel(text, w)
        wd.setGeometry(*r)
        x[name] = x.get(name, []) + [wd]
        return wd

    def put(name, wd, r):
        wd.setParent(w)
        wd.setGeometry(*r)
        x[name] = x.get(name, []) + [wd]
        return wd
    for name in TEXTS:
        t = lab(name.lower(), TEXTS[name], (173, 41, 150, 46))
        t.setWordWrap(True)
    grp = put("sample", QGroupBox("Sample"), (146, 11, 204, 36))
    sample = put("sample", QLabel(""), (155, 26, 185, 13))
    sample.raise_()
    lab("dec", "&Decimal places:", (146, 60, 77, 13))
    dec = put("dec", QLineEdit("0"), (225, 55, 57, 23))
    x["dec"][0].setBuddy(dec)
    spin = put("dec", UpDown(w), (281, 55, 17, 23))
    spin.buddy = dec
    thou = put("thou", QCheckBox("&Use 1000 Separator (,)"), (146, 83, 131, 16))
    lab("sym", "&Symbol:", (146, 76, 39, 13))
    sym = put("sym", QComboBox(), (146, 93, 204, 21))
    sym.setEditable(True)
    sym.addItems(CURRENCIES)
    lab("neg", "&Negative numbers:", (146, 120, 90, 13))
    neg = put("neg", QListWidget(), (146, 137, 204, 59))
    neg.addItems(["-123.45", "(123.45)"])
    lab("fmt", "Format$:", (146, 205, 42, 13))
    fmt_text = put("fmt", QLabel(""), (191, 205, 159, 13))
    lab("list", "&Format:", (146, 54, 36, 13))
    flist = put("list", QListWidget(), (146, 68, 197, 95))
    lab("bool", "&True value:", (146, 63, 56, 13))
    tv = put("bool", QLineEdit("True"), (212, 59, 90, 23))
    lab("bool", "&False value:", (146, 89, 59, 13))
    fv = put("bool", QLineEdit("False"), (212, 85, 90, 23))
    lab("bool", "&Null value:", (146, 114, 51, 13))
    nv = put("bool", QLineEdit(""), (212, 111, 90, 23))
    lab("custom", "&Format String:", (146, 59, 66, 13))
    ctext = put("custom", QLineEdit(""), (215, 55, 135, 23))
    for wd in [v for vs in x.values() for v in vs]:
        wd.show()
    busy = [False]

    def items():
        if repeater:
            return Grid_binds.of(dlg).binds
        return Grid.of(dlg).cols

    def names():
        if repeater:
            return [str(b.get("PropertyName", "")) for b in items()]
        return Grid.of(dlg).names()

    def cat():
        return DF.CATS[max(0, cats.currentRow())]

    def fmt_of(c):
        d = int(_float(dec.text() or "0")) if dec.text().strip().lstrip("-").replace(",", "").isdigit() else 0
        d = max(0, min(d, 15))
        if c == "Number":
            return DF.number_format(d, thou.isChecked(), neg.currentRow() == 1)
        if c == "Currency":
            s = sym.currentText().split(" ")[0] if sym.currentText() != "None" else ""
            return DF.number_format(d, True, neg.currentRow() == 1, prefix=f'"{s}"' if s else "")
        if c == "Percentage":
            return DF.number_format(d, suffix="%")
        if c == "Scientific":
            return "0" + ("." + "0" * d if d else "") + "E+00"
        if c in ("Date", "Time"):
            it = flist.currentItem()
            return it.data(256) if it is not None else ""
        if c == "Custom":
            return ctext.text()
        return ""

    def layout():
        c = cat()
        on = SHOWN[c]
        for name, ws in x.items():
            for wd in ws:
                wd.setVisible(name in on)
        if c in ("Date", "Time"):
            flist.blockSignals(True)
            flist.clear()
            for f in (DF.DATES if c == "Date" else DF.TIMES):
                from PyQt6.QtWidgets import QListWidgetItem
                li = QListWidgetItem(sample_of(c, f))
                li.setData(256, f)
                flist.addItem(li)
            flist.setCurrentRow(0)
            flist.blockSignals(False)

    def samples():
        c = cat()
        f = fmt_of(c)
        sample.setText(sample_of(c, f))
        fmt_text.setText(f)

    def value():
        c = cat()
        out = DF.default_value()
        out["Type"] = DF.FMTTYPE.get(c, 1)
        out["Format"] = fmt_of(c)
        out["SubFormatType"] = DF.SUBTYPE[c]
        if c == "Boolean":
            out["HaveTrueFalseNull"] = 1
            out["TrueValue"], out["FalseValue"], out["NullValue"] = tv.text(), fv.text(), nv.text()
        return out

    def store(*_):
        if busy[0]:
            return
        samples()
        its = items()
        i = item.currentIndex()
        if not 0 <= i < len(its):
            return
        its[i]["DataFormat"] = value()
        if repeater:
            Grid_binds.of(dlg).changed()
        else:
            Grid.of(dlg).changed()

    def show(*_):
        busy[0] = True
        item.blockSignals(True)
        i = max(0, item.currentIndex())
        item.clear()
        item.addItems(names())
        item.setCurrentIndex(min(i, item.count() - 1))
        item.blockSignals(False)
        its = items()
        i = item.currentIndex()
        v = its[i].get("DataFormat") if 0 <= i < len(its) else None
        for wd in (cats, *[v2 for vs in x.values() for v2 in vs]):
            wd.setEnabled(0 <= i < len(its))
        c = DF.category_of(v)
        cats.setCurrentRow(DF.CATS.index(c))
        layout()
        f = str((v or {}).get("Format") or "")
        dec.setText(str(DF._decimals(f) if f else {"Currency": 2, "Scientific": 3}.get(c, 0)))
        thou.setChecked("," in f.split(";")[0])
        neg.setCurrentRow(1 if ";(" in f else 0)
        if c == "Currency" and f.startswith('"'):
            s = f[1:f.index('"', 1)] if f.count('"') >= 2 else ""
            sym.setCurrentIndex(next((k for k in range(sym.count()) if sym.itemText(k).split(" ")[0] == s), -1))
            if sym.currentIndex() < 0:
                sym.setEditText(s)
        if c in ("Date", "Time"):
            k = next((k for k in range(flist.count()) if flist.item(k).data(256) == f), None)
            if k is None and f:
                from PyQt6.QtWidgets import QListWidgetItem
                li = QListWidgetItem(sample_of(c, f))
                li.setData(256, f)
                flist.addItem(li)
                k = flist.count() - 1
            flist.setCurrentRow(k or 0)
        have = int((v or {}).get("HaveTrueFalseNull", 0) or 0)
        tv.setText(str(v.get("TrueValue", "True")) if have else "True")
        fv.setText(str(v.get("FalseValue", "False")) if have else "False")
        nv.setText(str(v.get("NullValue", "")) if have else "")
        ctext.setText(f if c == "Custom" else "")
        samples()
        busy[0] = False

    def cat_chosen(_):
        if busy[0]:
            return
        busy[0] = True
        layout()
        dec.setText(str({"Currency": 2, "Scientific": 3}.get(cat(), 0)))
        busy[0] = False
        store()
    item.currentIndexChanged.connect(show)
    cats.currentRowChanged.connect(cat_chosen)
    for wd in (dec, tv, fv, nv, ctext):
        wd.textChanged.connect(store)
    thou.toggled.connect(store)
    neg.currentRowChanged.connect(store)
    flist.currentRowChanged.connect(store)
    sym.editTextChanged.connect(store)
    if repeater:
        Grid_binds.of(dlg).refresh.append(show)
    else:
        Grid.of(dlg).refresh.append(show)
    show()
    dlg.bound[key] = {"DataFormat": cats}
    dlg.unbound[key] = []


class Grid_binds:
    """the RepeaterBindings of a DataRepeater while its pages are open (the RepeaterBindings and Format pages)"""

    def __init__(self, dlg):
        self.dlg = dlg
        self.binds = [dict(b) for b in (dlg.value("(RepeaterBindings)") or [])]
        self.refresh = []
        dlg.tabs.currentChanged.connect(lambda _: [f() for f in self.refresh])

    @staticmethod
    def of(dlg):
        g = dlg.__dict__.get("binds")
        if g is None:
            g = dlg.__dict__["binds"] = Grid_binds(dlg)
        return g

    def changed(self):
        self.dlg.changed("(RepeaterBindings)", [dict(b) for b in self.binds])


# ---------------------------------------------------------------- the Bands of an MSHFlexGrid
BAND_PROPS = [("BandIndent", 0, False), ("ColHeader", 0, False), ("GridLinesBand", 1, True), ("TextStyleBand", 0, True),
              ("TextStyleHeader", 0, True)]          # (name, default, always written: the corpus of PSC)


def band0(ctl):
    """the properties of band 0 of an MSHFlexGrid in the .frm (_Band(0).GridLinesBand ...)"""
    out = {}
    for k, d, _ in BAND_PROPS:
        v = ctl.get(f"_Band(0).{k}")
        out[k] = d if v is None else int(V._num(v))
    return out


def set_band0(ctl, values):
    """write band 0 as VB6 does when its page is applied: _NumberOfBands, _Band(0).Cols, then its properties in the
    order of their names (MSHFLXGD.OCX: "_Band(%1!d!)." and the name of the property)"""
    from lbvc.project.source import Property
    cols = ctl.get("_Band(0).Cols")
    for k, _, _ in BAND_PROPS:
        V.remove(ctl, f"_Band(0).{k}")
    if ctl.get("_NumberOfBands") is None:
        V._append(ctl, Property("_NumberOfBands", 1))
    if cols is None:
        V.remove(ctl, "_Band(0).Cols")
        V._append(ctl, Property("_Band(0).Cols", int(V._num(ctl.get("Cols", 2)))))
    for k, d, always in BAND_PROPS:
        v = int(values.get(k, d))
        if always or v != d:
            V._append(ctl, Property(f"_Band(0).{k}", v))


def hbands_page(dlg, key, w, made):
    """the Bands page of an MSHFlexGrid: band 0 (a grid without hierarchical data has only it) - Gridlines,
    BandIndent, TextStyle, TextStyleHeader, ColumnHeaders; BandExpandable and the columns of a band need its data"""
    band, grid, indent, ts, tsh, expand, colh, cols, up, down = _ids(made, 229, 230, 233, 237, 238, 227, 232, 234,
                                                                      235, 236)
    band.clear()
    band.addItem("Band 0 ()")
    for wd in (expand, cols, up, down):
        wd.setEnabled(False)
    cols.setColumnCount(2)
    cols.setHorizontalHeaderLabels(["Column Caption", "Column Name"])
    cols.verticalHeader().hide()
    from PyQt6.QtCore import Qt
    for b, arrow in ((up, Qt.ArrowType.UpArrow), (down, Qt.ArrowType.DownArrow)):
        from PyQt6.QtWidgets import QStyle
        b.setIcon(dlg.style().standardIcon(QStyle.StandardPixmap.SP_ArrowUp if arrow == Qt.ArrowType.UpArrow
                                           else QStyle.StandardPixmap.SP_ArrowDown))
    vals = band0(dlg.ctl)
    _pick(grid, vals["GridLinesBand"])
    _pick(ts, vals["TextStyleBand"])
    _pick(tsh, vals["TextStyleHeader"])
    indent.setText(str(vals["BandIndent"]))
    colh.setChecked(bool(vals["ColHeader"]))

    def edit(k, v):
        vals[k] = v
        dlg.raw_apply["(Band0)"] = lambda v: set_band0(dlg.ctl, v)
        dlg.changed("(Band0)", dict(vals))
    grid.activated.connect(lambda i: edit("GridLinesBand", _enum(grid.itemText(i))))
    ts.activated.connect(lambda i: edit("TextStyleBand", _enum(ts.itemText(i))))
    tsh.activated.connect(lambda i: edit("TextStyleHeader", _enum(tsh.itemText(i))))
    colh.toggled.connect(lambda on: edit("ColHeader", 1 if on else 0))
    indent.textEdited.connect(lambda s: edit("BandIndent", int(s)) if s.strip().isdigit() else None)
    dlg.bound[key] = {"(Band0)": band}
    dlg.unbound[key] = []


PAGES = {("MSHFlexGrid", "Bands"): hbands_page, ("DataGrid", "Columns"): columns_page, ("DataGrid", "Layout"): layout_page,
         ("DataGrid", "Splits"): splits_page, ("DataGrid", "Format"): format_page,
         ("DBGrid", "Columns"): columns_page, ("DBGrid", "Layout"): layout_page, ("DBGrid", "Splits"): splits_page,
         ("DataRepeater", "Format"): format_page}
