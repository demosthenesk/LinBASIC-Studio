"""dfmtdialog.py - the DataFormat of a bound control at design time (phase 7.7): the "Format" property page of VB6
(General, Number, Currency, Date, Time, Percentage, Scientific, Boolean, Checkbox, Picture, Custom). What it gives
is the BeginProperty DataFormat block of the .frm: Type (FormatType of MSSTDFMT), Format, HaveTrueFalseNull and the
True / False / Null values, FirstDayOfWeek, FirstWeekOfYear, LCID and SubFormatType (the category of the page).
The formats it writes are in the notation of Format$ ("." decimal point, "," thousands) with LCID 1033: the runtime
reads a Format in the language of its LCID (runtime/stdfmt.c) and shows it with the regional settings of the user."""

from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QStackedWidget, QWidget, QFormLayout,
                             QSpinBox, QCheckBox, QComboBox, QLineEdit, QLabel, QPushButton, QGroupBox)

CATS = ["General", "Number", "Currency", "Date", "Time", "Percentage", "Scientific", "Boolean", "Checkbox", "Picture",
        "Custom"]
SUBTYPE = {c: i for i, c in enumerate(CATS)}                 # SubFormatType
FMTTYPE = {"General": 0, "Boolean": 5, "Checkbox": 4, "Picture": 2}   # Type; the others are Custom (1)
DATES = ["M/d/yyyy", "M/d/yy", "MM/dd/yy", "MM/dd/yyyy", "dd-MMM-yy", "d-MMM-yyyy", "yyyy-MM-dd",
         "MMMM d, yyyy", "dddd, MMMM dd, yyyy"]
TIMES = ["h:nn:ss AMPM", "hh:nn AMPM", "H:mm:ss", "H:mm"]
DAYS = ["0 - fmtDayUseSystem", "1 - fmtSunday", "2 - fmtMonday", "3 - fmtTuesday", "4 - fmtWednesday",
        "5 - fmtThursday", "6 - fmtFriday", "7 - fmtSaturday"]
WEEKS = ["0 - fmtWeekUseSystem", "1 - fmtFirstJan1", "2 - fmtFirstFourDays", "3 - fmtFirstFullWeek"]
LCID = 1033


def _decimals(fmt):
    """the decimal places of a number format ("#,##0.00" -> 2)"""
    head = fmt.split(";")[0]
    if "." not in head:
        return 0
    return sum(1 for ch in head.split(".", 1)[1] if ch in "0#")


def number_format(dec, thousands=False, parens=False, prefix="", suffix=""):
    base = ("#,##0" if thousands else "0") + ("." + "0" * dec if dec else "")
    base = prefix + base + suffix
    return f"{base};({base})" if parens else base


def default_value():
    return {"Type": 0, "Format": "", "HaveTrueFalseNull": 0, "TrueValue": "True", "FalseValue": "False",
            "NullValue": "", "FirstDayOfWeek": 0, "FirstWeekOfYear": 0, "LCID": LCID, "SubFormatType": 0}


def category_of(v):
    """the category of the page for a DataFormat of the .frm"""
    if not v:
        return "General"
    sub = v.get("SubFormatType")
    try:
        sub = int(sub)
    except (TypeError, ValueError):
        sub = None
    if sub is not None and 0 <= sub < len(CATS):
        cat = CATS[sub]
        if FMTTYPE.get(cat, 1) == int(v.get("Type", 0) or 0):
            return cat
    t = int(v.get("Type", 0) or 0)
    return {0: "General", 2: "Picture", 3: "Picture", 4: "Checkbox", 5: "Boolean", 6: "Picture"}.get(t, "Custom")


class DataFormatDialog(QDialog):
    def __init__(self, parent, name, value):
        super().__init__(parent)
        self.setWindowTitle(f"{name} - DataFormat")
        v = dict(default_value(), **(value or {}))
        fmt = str(v.get("Format") or "")
        top = QHBoxLayout()
        self.cats = QListWidget()
        self.cats.addItems(CATS)
        self.cats.setMaximumWidth(140)
        top.addWidget(self.cats)
        self.pages = QStackedWidget()
        top.addWidget(self.pages, 1)
        self.w = {}
        for c in CATS:
            self.pages.addWidget(self._page(c, fmt, v))
        self.cats.currentRowChanged.connect(self.pages.setCurrentIndex)
        self.cats.currentRowChanged.connect(lambda _: self._sample())
        box = QVBoxLayout(self)
        box.addLayout(top, 1)
        self.sample = QLabel()
        box.addWidget(self.sample)
        row = QHBoxLayout()
        row.addStretch(1)
        ok, cancel = QPushButton("OK"), QPushButton("Cancel")
        ok.setDefault(True)
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        row.addWidget(ok)
        row.addWidget(cancel)
        box.addLayout(row)
        self.cats.setCurrentRow(CATS.index(category_of(value)))
        self.resize(560, 340)
        self._sample()

    # ---------------------------------------------------------------- the pages
    def _page(self, cat, fmt, v):
        page = QWidget()
        f = QFormLayout(page)
        w = {}
        dec = _decimals(fmt)
        if cat == "General":
            f.addRow(QLabel("The value as it is, without a format."))
        if cat in ("Number", "Currency", "Percentage", "Scientific"):
            w["dec"] = QSpinBox()
            w["dec"].setRange(0, 15)
            w["dec"].setValue(dec if fmt else 2)
            f.addRow("Decimal places:", w["dec"])
        if cat in ("Number", "Currency"):
            w["thou"] = QCheckBox("Use 1000 separator")
            w["thou"].setChecked("," in fmt.split(";")[0] if fmt else cat == "Currency")
            f.addRow(w["thou"])
            w["neg"] = QComboBox()
            w["neg"].addItems(["-1234.10", "(1234.10)"])
            w["neg"].setCurrentIndex(1 if ";(" in fmt else 0)
            f.addRow("Negative numbers:", w["neg"])
        if cat == "Currency":
            w["sym"] = QLineEdit("$")
            if fmt.startswith('"') and fmt.count('"') >= 2:
                w["sym"].setText(fmt[1:fmt.index('"', 1)])
            f.addRow("Symbol:", w["sym"])
        if cat in ("Date", "Time"):
            w["pick"] = QComboBox()
            w["pick"].setEditable(True)
            w["pick"].addItems(DATES if cat == "Date" else TIMES)
            if fmt:
                w["pick"].setCurrentText(fmt)
            f.addRow("Format:", w["pick"])
        if cat == "Custom":
            w["text"] = QLineEdit(fmt)
            f.addRow("Format:", w["text"])
            f.addRow(QLabel('Format$ notation: "." is the decimal point, "," the thousands separator.'))
        if cat in ("Date", "Time", "Custom"):
            w["day"] = QComboBox()
            w["day"].addItems(DAYS)
            w["day"].setCurrentIndex(int(v.get("FirstDayOfWeek", 0) or 0) % 8)
            f.addRow("First day of week:", w["day"])
            w["week"] = QComboBox()
            w["week"].addItems(WEEKS)
            w["week"].setCurrentIndex(int(v.get("FirstWeekOfYear", 0) or 0) % 4)
            f.addRow("First week of year:", w["week"])
        if cat in ("Boolean", "Custom"):
            g = QGroupBox("Values" if cat == "Boolean" else "Null value")
            gf = QFormLayout(g)
            have = int(v.get("HaveTrueFalseNull", 0) or 0)
            if cat == "Boolean":
                w["true"] = QLineEdit(str(v.get("TrueValue")) if have else "True")
                w["false"] = QLineEdit(str(v.get("FalseValue")) if have else "False")
                gf.addRow("True value:", w["true"])
                gf.addRow("False value:", w["false"])
            w["null"] = QLineEdit(str(v.get("NullValue") or "") if have else "")
            gf.addRow("Null value:", w["null"])
            f.addRow(g)
        if cat == "Checkbox":
            f.addRow(QLabel("True: 1 (Checked), False: 0 (Unchecked), Null: 2 (Grayed)."))
        if cat == "Picture":
            f.addRow(QLabel("The field holds a picture (an Image or a PictureBox)."))
        for x in w.values():
            sig = getattr(x, "textChanged", None) or getattr(x, "valueChanged", None) or \
                getattr(x, "currentIndexChanged", None) or getattr(x, "toggled", None)
            if isinstance(x, QComboBox):
                x.currentTextChanged.connect(lambda _=None: self._sample())
            elif sig is not None:
                sig.connect(lambda _=None: self._sample())
        self.w[cat] = w
        return page

    def _cat(self):
        return CATS[max(0, self.cats.currentRow())]

    def _format(self, cat):
        w = self.w[cat]
        if cat == "Number":
            return number_format(w["dec"].value(), w["thou"].isChecked(), w["neg"].currentIndex() == 1)
        if cat == "Currency":
            sym = w["sym"].text().replace('"', "")
            return number_format(w["dec"].value(), w["thou"].isChecked(), w["neg"].currentIndex() == 1,
                                 prefix=f'"{sym}"' if sym else "")
        if cat == "Percentage":
            return number_format(w["dec"].value(), suffix="%")
        if cat == "Scientific":
            d = w["dec"].value()
            return "0" + ("." + "0" * d if d else "") + "E+00"
        if cat in ("Date", "Time"):
            return w["pick"].currentText()
        if cat == "Custom":
            return w["text"].text()
        return ""

    def _sample(self):
        if not hasattr(self, "sample"):
            return
        cat = self._cat()
        fmt = self._format(cat)
        self.sample.setText(f"Format: {fmt}" if fmt else "")

    def chosen(self):
        """the DataFormat block as a dict (vbprops: BeginProperty DataFormat)"""
        cat = self._cat()
        w = self.w[cat]
        out = default_value()
        out["Type"] = FMTTYPE.get(cat, 1)
        out["Format"] = self._format(cat)
        out["SubFormatType"] = SUBTYPE[cat]
        if "day" in w:
            out["FirstDayOfWeek"] = w["day"].currentIndex()
            out["FirstWeekOfYear"] = w["week"].currentIndex()
        if cat == "Boolean":
            out["HaveTrueFalseNull"] = 1
            out["TrueValue"], out["FalseValue"], out["NullValue"] = w["true"].text(), w["false"].text(), w["null"].text()
        elif cat == "Custom" and w["null"].text():
            out["HaveTrueFalseNull"] = 1
            out["TrueValue"], out["FalseValue"], out["NullValue"] = "", "", w["null"].text()
        return out
