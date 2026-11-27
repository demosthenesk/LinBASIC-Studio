"""The SQL Statement window of the Visual Data Manager (VB6 SQL.FRM): Execute, Clear, Save (a QueryDef)."""
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QPlainTextEdit

from ..rtobj import RtError
from . import core as C


class SQLWindow(QWidget):
    def __init__(self, vd):
        super().__init__()
        self.vd = vd
        self.setWindowTitle("SQL Statement")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        row = QHBoxLayout()
        self.b_exec = QPushButton("&Execute")
        self.b_clear = QPushButton("&Clear")
        self.b_save = QPushButton("&Save")
        for b in (self.b_exec, self.b_clear, self.b_save):
            row.addWidget(b)
        lay.addLayout(row)
        self.text = QPlainTextEdit()
        lay.addWidget(self.text, 1)
        self.b_exec.clicked.connect(self.execute)
        self.b_clear.clicked.connect(self.clear)
        self.b_save.clicked.connect(self.save)
        self.text.textChanged.connect(lambda: self.b_exec.setEnabled(bool(self.text.toPlainText())))
        self.text.setPlainText(vd.prefs.get("sql", "") or "")
        self.b_exec.setEnabled(bool(self.text.toPlainText()))

    def clear(self):
        self.text.setPlainText("")
        self.text.setFocus()

    def execute(self):
        from .dbwindow import open_query
        sql = self.text.toPlainText()
        if not sql:
            return None
        self.vd.prefs["sql"] = sql
        return open_query(self.vd, sql, True)

    def save(self):
        """Save: the SQL into the query chosen in the Database Window, or a new QueryDef"""
        vd, db = self.vd, self.vd.s.db
        try:
            dw = vd.dbwin
            it = dw.selected() if dw is not None else None
            if it is not None and dw.tag(it) == "Query":
                if vd.ask(f"Update '{it.text(0)}'?"):
                    db.get("QueryDefs", it.text(0)).let("SQL", self.text.toPlainText())
                    return it.text(0)
            name = vd.ask_text("Enter QueryDef Name:")
            if not name:
                return None
            if dupe_name(vd, name):
                return None
            qd = db.call("CreateQueryDef", name)
            if vd.ask("Is this a SQLPassThrough QueryDef?", default=False):
                cn = vd.ask_text(C.MSG[18])
                if cn:
                    qd.let("Connect", cn)
                    if not vd.ask("Is the Query Row Returning?"):
                        qd.let("ReturnsRecords", False)
            qd.let("SQL", self.text.toPlainText())
            db.get("QueryDefs").call("Refresh")
            vd.refresh_tables()
            return name
        except RtError as e:
            vd.show_error(e)
        return None

    def closeEvent(self, e):
        self.vd.prefs["sql"] = self.text.toPlainText()
        super().closeEvent(e)


def dupe_name(vd, name):
    """DupeTableName: a table or query of that name - deleted when the user says so (True: it stays)"""
    db = vd.s.db
    for coll, msg in (("TableDefs", C.MSG[27]), ("QueryDefs", C.MSG[28])):
        for o in db.each(coll):
            if o.get("Name").upper() == name.upper():
                if vd.ask(msg):
                    db.get(coll).call("Delete", o.get("Name"))
                    return False
                return True
    return False
