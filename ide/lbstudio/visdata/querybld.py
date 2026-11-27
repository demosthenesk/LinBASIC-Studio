"""Query Builder of the Visual Data Manager (VB6 QUERY.FRM, JOIN.FRM): a SELECT built from the tables, the fields
to show, criteria (And / Or into Criteria, List Possible Values), table joins, Group By, Order By and Top N;
Run, Show, Copy (into the SQL Statement window), Save (a QueryDef)."""
from PyQt6.QtWidgets import (QWidget, QDialog, QListWidget, QPushButton, QComboBox, QCheckBox, QLineEdit,
                             QPlainTextEdit, QRadioButton, QAbstractItemView)

from ..rtobj import RtError
from . import core as C
from .dialogs import strip_owner

NONE = "(none)"
dbText, dbMemo, dbDate = 10, 12, 8
OPERATORS = ["=", "<>", ">", ">=", "<", "<=", "Like"]


def add_brackets(name):
    return "[" + name + "]" if " " in name and not name.startswith("[") else name


def strip_brackets(name):
    return name[1:-1] if name.startswith("[") else name


def get_token(line, delim):
    """stGetToken: the first token of line (a quoted part kept whole, its quotes taken off) and the rest"""
    iq = line.find('"') + 1
    idl = line.find(delim) + 1
    if iq > 0 and iq < idl:
        ic = line.find('"', iq) + 1
        idl = line.find(delim, ic) + 1
    if idl != 0:
        tok, rest = line[:idl - 1].strip(), line[idl:]
    else:
        tok, rest = line.strip(), ""
    if tok:
        if tok[0] == '"':
            tok = tok[1:]
        if tok and tok[-1] == '"':
            tok = tok[:-1]
    return tok, rest


def stf(name, part):
    """stSTF: the table (0) or the field (1) of 'table.field' (an owner taken off)"""
    first = name.find(".")
    if name.find(".", first + 1) >= 0:
        name = strip_owner(name)
    i = name.find(".")
    return name[:i] if part == 0 else name[i + 1:]


def build_where(criteria):
    """the criteria as VisData puts them into the WHERE: each And part in parentheses"""
    s = "AND " + criteria.lstrip()
    s = "".join(" " if ch == "\r" else ch for ch in s if ch != "\n").rstrip()
    new, paren, rest = "", False, s
    while True:
        tok, rest = get_token(rest, " ")
        tok += " "
        if not paren and tok.upper() == "AND ":
            new += tok + "("
            paren = True
        elif paren and tok.upper() == "AND ":
            new += ") " + tok + "("
        elif tok.upper() in ("OR", "IN ", "LIKE"):
            new += " " + tok
        else:
            new += tok
        if rest == "":
            break
    s = new + ")"
    if s[1:3] == "OR":
        s = s[4:len(s) - 1]
    else:
        _tok, s = get_token(s, " ")
    return " Where " + s if s else ""


def build_sql(tables, show, criteria, joins, group, order, desc, top, percent):
    where = build_where(criteria) if criteria else ""
    if joins:
        where += " Where " if not where else " And "
        where += " And ".join(joins)
    if group != NONE and group:
        where += " Group By " + group
    if order != NONE and order:
        where += " Order By " + order
        if desc:
            where += " Desc "
    fs = ",".join(show) if show else (",".join(add_brackets(t) + ".*" for t in tables) or "*")
    ts = ",".join(add_brackets(t) for t in tables)
    sql = "Select "
    if top:
        sql += " TOP " + top + " "
        if percent:
            sql += " PERCENT "
    return sql + fs + " From " + ts + where


class QueryBuilder(QWidget):
    def __init__(self, vd):
        super().__init__()
        self.vd = vd
        self.setWindowTitle("Query Builder")
        P = C.Place(self, 497, 335)             # (VB6 oracle vb_qb.png)
        self.field, self.operator, self.value = QComboBox(), QComboBox(), QComboBox()
        self.value.setEditable(True)
        self.operator.addItems(OPERATORS)
        P.label("Field Name:", 8, 0, 56)
        P.put(self.field, 8, 16, 193, 21)
        P.label("Operator:", 208, 0, 48)
        P.put(self.operator, 208, 16, 89, 21)
        P.label("Value:", 304, 0, 30)
        P.put(self.value, 304, 16, 185, 21)
        self.b_and = P.put(QPushButton("&And into Criteria"), 8, 40, 144, 21)
        self.b_or = P.put(QPushButton("&Or into Criteria"), 152, 40, 145, 21)
        self.b_values = P.put(QPushButton("List &Possible Values"), 304, 40, 185, 21)
        self.tables, self.show_fields = QListWidget(), QListWidget()
        self.tables.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.show_fields.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        P.label("Tables: ", 8, 64, 38)
        P.put(self.tables, 8, 80, 121, 95)
        P.label("Fields to Show: ", 136, 64, 76)
        P.put(self.show_fields, 136, 80, 153, 95)
        self.group, self.order = QComboBox(), QComboBox()
        P.label("Group By: ", 304, 64, 51)
        P.put(self.group, 304, 80, 185, 21)
        P.label("Order By: ", 304, 104, 50)
        self.asc, self.desc = QRadioButton("Asc"), QRadioButton("Desc")
        self.asc.setChecked(True)
        P.put(self.asc, 384, 104, 41, 15)
        P.put(self.desc, 432, 104, 57, 15)
        P.put(self.order, 304, 120, 185, 21)
        self.b_join = P.put(QPushButton("Set Table &Joins"), 304, 144, 185, 17)
        self.joins = P.put(QListWidget(), 304, 160, 185, 17)
        self.top = QLineEdit()
        P.label("Top N Value:", 96, 194, 98)
        P.put(self.top, 200, 192, 49, 19)
        self.percent = P.put(QCheckBox("Top Percent"), 256, 192, 145, 17)
        self.criteria = QPlainTextEdit()
        P.label("Criteria: ", 8, 200, 42)
        P.put(self.criteria, 8, 216, 481, 81)
        self.b_run, self.b_show, self.b_copy = QPushButton("&Run"), QPushButton("&Show"), QPushButton("Cop&y")
        self.b_save, self.b_clear, self.b_close = QPushButton("Sa&ve"), QPushButton("C&lear"), QPushButton("&Close")
        for k, b in enumerate((self.b_run, self.b_show, self.b_copy, self.b_save, self.b_clear, self.b_close)):
            P.put(b, 8 + 80 * k, 304, 80, 25)
        self.b_and.clicked.connect(lambda: self.into_criteria("And "))
        self.b_or.clicked.connect(lambda: self.into_criteria(" Or "))
        self.b_values.clicked.connect(self.possible_values)
        self.b_join.clicked.connect(self.set_joins)
        self.b_run.clicked.connect(lambda: self.run("run"))
        self.b_show.clicked.connect(lambda: self.run("show"))
        self.b_copy.clicked.connect(lambda: self.run("copy"))
        self.b_save.clicked.connect(lambda: self.run("save"))
        self.b_clear.clicked.connect(self.clear)
        self.b_close.clicked.connect(self.close_window)
        self.field.activated.connect(lambda *_: self.value.clear())
        self.tables.itemClicked.connect(lambda *_: self.tables_changed())
        try:
            self.tables.addItems([td.get("Name") for td in vd.s.tabledefs(False)])
        except RtError as e:
            vd.show_error(e)
        if self.tables.count():
            self.tables.setCurrentRow(0, self.tables.selectionModel().SelectionFlag.NoUpdate)
        self.value.setEditText("")

    def selected_tables(self):
        return [self.tables.item(i).text() for i in range(self.tables.count()) if self.tables.item(i).isSelected()]

    def tables_changed(self):
        """the fields of the tables chosen into the lists"""
        vd = self.vd
        vd.msgbar("Updating Form Fields", True)
        for w in (self.field, self.group, self.order):
            w.clear()
        self.show_fields.clear()
        self.value.clear()
        self.group.addItem(NONE)
        self.order.addItem(NONE)
        try:
            for t in self.selected_tables():
                for f in vd.s.db.get("TableDefs", t).each("Fields"):
                    s = add_brackets(t) + "." + add_brackets(f.get("Name"))
                    self.field.addItem(s)
                    self.show_fields.addItem(s)
                    self.group.addItem(s)
                    self.order.addItem(s)
        except RtError as e:
            vd.show_error(e)
        vd.msgbar()

    def field_type(self, text):
        t, f = stf(text, 0), stf(text, 1)
        return self.vd.s.db.get("TableDefs", strip_brackets(t)).get("Fields", strip_brackets(f)).get("Type")

    def into_criteria(self, joiner):
        text = self.field.currentText()
        if not text:
            return
        try:
            kind = self.field_type(text)
        except RtError as e:
            self.vd.show_error(e)
            return
        cur = self.criteria.toPlainText()
        if cur:
            cur += "\n" + joiner
        q = "'" if kind in (dbText, dbMemo, dbDate) else ""
        self.criteria.setPlainText(cur + text + " " + self.operator.currentText() + " " + q +
                                   self.value.currentText() + q)
        self.field.setFocus()

    def possible_values(self):
        vd = self.vd
        text = self.field.currentText()
        vd.msgbar("Getting Possible Values", True)
        try:
            rs = vd.s.db.call("OpenRecordset", "select Distinct " + strip_owner(text) + " from " + stf(text, 0))
            while not rs.get("EOF"):
                v = rs.get("Fields", 0).get("Value")
                if C.text_of(v).strip():
                    self.value.addItem(C.text_of(v))
                rs.call("MoveNext")
            rs.call("Close")
            self.value.setEditText(self.value.itemText(0))
            self.value.setFocus()
        except RtError:
            self.value.setEditText("")
        vd.msgbar()

    def set_joins(self):
        if len(self.selected_tables()) < 2:
            self.vd.say("You Must Have at Least 2 Tables Selected!", "warn")
            return None
        self.vd.msgbar("Choose Joins")
        d = JoinDialog(self)
        if not self.vd.quiet:
            d.exec()
            self.vd.msgbar()
        return d

    def clear(self):
        self.tables.clearSelection()
        self.criteria.setPlainText("")
        self.top.setText("")

    def sql(self):
        show = [self.show_fields.item(i).text() for i in range(self.show_fields.count())
                if self.show_fields.item(i).isSelected()]
        return build_sql(self.selected_tables(), show, self.criteria.toPlainText().replace("\n", "\r\n"),
                         [self.joins.item(i).text() for i in range(self.joins.count())],
                         self.group.currentText(), self.order.currentText(), self.desc.isChecked(),
                         self.top.text(), self.percent.isChecked())

    def run(self, how):
        vd = self.vd
        if self.show_fields.count() == 0:
            vd.say("No Query Entered!", "warn")
            return None
        vd.msgbar("Building Query", True)
        sql = self.sql()
        if how == "run":
            from .dbwindow import open_query
            vd.msgbar("Running Query", True)
            open_query(vd, sql, True)
        elif how == "show":
            vd.msgbar()
            vd.messages.append(sql)
            if not vd.quiet:
                from PyQt6.QtWidgets import QMessageBox
                QMessageBox.information(self, "SQL Query", sql)
        elif how == "copy":
            if vd.sqlwin is not None:
                vd.sqlwin.text.setPlainText(sql)
        else:
            from .sqlwin import dupe_name
            vd.msgbar()
            name = vd.ask_text("Enter QueryDef Name:")
            if not name:
                return sql
            try:
                if dupe_name(vd, name):
                    return sql
                vd.s.db.call("CreateQueryDef", name, sql)
            except RtError as e:
                vd.show_error(e)
                return sql
            vd.refresh_tables()
        vd.msgbar()
        return sql

    def close_window(self):
        p = self.parentWidget()
        (p if p is not None and p.inherits("QMdiSubWindow") else self).close()


class JoinDialog(QDialog):
    """Join Tables: a pair of the tables and the fields to join on, added to the joins of the query"""

    def __init__(self, qb):
        super().__init__(qb)
        self.qb, self.vd = qb, qb.vd
        self.t1 = self.t2 = ""
        self.setWindowTitle("Join Tables")
        P = C.Place(self, 394, 135)             # (VB6 oracle vb_join.png)
        self.tables, self.f1, self.f2 = QListWidget(), QListWidget(), QListWidget()
        self.tables.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        P.label("Select Table Pair: ", 8, 3, 86)
        P.put(self.tables, 8, 20, 121, 69)
        P.label("Select Fields to Join on: ", 136, 3, 249, center=True)
        P.put(self.f1, 136, 20, 121, 69)
        P.put(self.f2, 264, 20, 121, 69)
        self.b_add = P.put(QPushButton("&Add Join to Query"), 8, 104, 121, 25)
        self.b_clear = P.put(QPushButton("C&lear All Joins"), 136, 104, 121, 25)
        self.b_close = P.put(QPushButton("&Close"), 264, 104, 121, 25)
        self.b_add.setEnabled(False)
        self.tables.addItems(qb.selected_tables())
        self.tables.itemClicked.connect(lambda *_: self.pair_changed())
        self.f1.itemClicked.connect(lambda *_: self.b_add.setEnabled(bool(self.f2.currentItem())))
        self.f2.itemClicked.connect(lambda *_: self.b_add.setEnabled(bool(self.f1.currentItem())))
        self.b_add.clicked.connect(self.add_join)
        self.b_clear.clicked.connect(lambda: qb.joins.clear())
        self.b_close.clicked.connect(self.reject)

    def pair_changed(self):
        sel = [self.tables.item(i).text() for i in range(self.tables.count()) if self.tables.item(i).isSelected()]
        self.t1 = sel[0] if sel else ""
        self.t2 = sel[1] if len(sel) > 1 else ""
        if not self.t2:
            return
        try:
            for t, w in ((self.t1, self.f1), (self.t2, self.f2)):
                w.clear()
                w.addItems([f.get("Name") for f in self.vd.s.db.get("TableDefs", t).each("Fields")])
        except RtError as e:
            self.vd.show_error(e)

    def add_join(self):
        a, b = self.f1.currentItem(), self.f2.currentItem()
        if a is None or b is None:
            return
        self.qb.joins.addItem(add_brackets(self.t1) + "." + add_brackets(a.text()) + "=" + add_brackets(self.t2) +
                              "." + add_brackets(b.text()))
        self.tables.clearSelection()
