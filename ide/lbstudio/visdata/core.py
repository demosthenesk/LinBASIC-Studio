"""The state of the Visual Data Manager: the runtime's DAO (DBEngine, the workspace, the open database), its
preferences and the texts of VisData (VB6 VISDATA.BAS)."""
from ..rtobj import Server, RtError, EMPTY

# the kinds of database (the Connect of OpenDatabase), as VisData names them
MSACCESS = "Microsoft Access"
DBASE3, DBASE4, DBASE5 = "Dbase III;", "Dbase IV;", "Dbase 5.0;"
FOX20, FOX25, FOX26, FOX30 = "FoxPro 2.0;", "FoxPro 2.5;", "FoxPro 2.6;", "FoxPro 3.0;"
PDOX3, PDOX4, PDOX5 = "Paradox 3.X;", "Paradox 4.X;", "Paradox 5.X;"
BTRIEVE, EXCEL50, TEXTFILES, SQLDB = "Btrieve;", "Excel 5.0;", "Text;", "ODBC;"

RS_TABLE, RS_DYNASET, RS_SNAPSHOT, RS_PASSTHRU = 0, 1, 2, 3
FORM_DATACTL, FORM_NODATACTL, FORM_DATAGRID = 0, 1, 2

# DAO constants
dbOpenTable, dbOpenDynaset, dbOpenSnapshot = 1, 2, 4
dbSQLPassThrough = 64
dbSystemObject = -2147483646
dbAttachedTable, dbAttachedODBC = 0x40000000, 0x20000000

MSG = {
    1: "Execute Commit or Rollback First.", 2: "Closing Recordsets", 3: "Table already exists, delete it?",
    4: "Enter New Table Name:", 5: "Ready", 6: ", please wait...", 7: "Refreshing Table List", 8: "Number: ",
    9: "Display the Data Access Errors Collection?",
    10: "Can't Open a Table Object on an Attached Table, Use Dynaset?",
    11: "Opening Attached Table as Dynaset", 12: "Opening Attached Table as Snapshot", 13: "Opening Full Table",
    14: "Opening Single Table Dynaset", 15: "Opening Single Table Snapshot", 16: "Opening PassThru Snapshot",
    17: "Is this a SQLPassThrough Query?", 18: "Enter Connect property value:",
    19: "Can't Open a Table Object from a QueryDef, Use Dynaset?", 20: "Opening Query Snapshot",
    21: "Opening Query Dynaset", 22: "SQL Statement", 23: "Execute ", 24: " Query?", 25: "Executing Query",
    26: "  [Not Updatable]", 27: "Table already exists, Delete it?", 28: "QueryDef already exists, Delete it?",
    29: "Enter Value for Parameter:", 30: "There are no current data access errors!",
    31: "Can't show Errors at this time!", 32: "Data has been changed, Commit it?", 33: "RollBack All changes?",
    34: "Can't Close with Transactions Pending!", 35: "You must Close First!",
    36: "Open Microsoft Access Database", 37: "Open Dbase Database", 38: "Open FoxPro Database",
    39: "Open Paradox Database", 40: "Open Excel File", 41: "Open Btrieve Database", 42: "Open Text Database",
    43: "Opening Database", 44: "NOTE: Use of Attached Tables is the Recommended Method", 45: "Repairing ",
    46: "Attempt to Repair it?", 47: "Enter Directory Name for New ISAM Database:",
    48: "Select Microsoft Access Database to Compact", 49: "Microsoft Access MDBs (*.mdb)",
    51: "Select Microsoft Access Database to Compact to", 52: "Encrypt Compacted Database?", 53: "Compacting ",
    54: "Open Newly Compacted Database?", 55: "Select Microsoft Access Database to Create",
    56: "Exporting Table: ", 57: "Export ", 58: "in ", 59: "Creating Indexes:", 60: "Successfully Exported:",
    61: "Successfully Exported SQL Statement.", 62: "Table already exists - overwrite?", 63: "Importing Table: ",
    64: "Successfully Imported:", 65: "Invalid Directory Name!",
}

DEFAULT_PREFS = {"open_last": False, "allow_sys": False, "query_timeout": 5, "login_timeout": 20,
                 "rstype": RS_DYNASET, "formtype": FORM_NODATACTL, "mru": [], "last": None}


class Session:
    """the runtime's DBEngine, its default workspace and the database VisData has open"""

    def __init__(self):
        self.srv = None
        self.dbe = self.ws = self.db = None
        self.dbname = ""
        self.datatype = ""
        self.odbc_source = ""
        self.readonly = False
        self.trans_pending = False
        self.changed = False

    def start(self):
        if self.srv is None:
            self.srv = Server()
            self.dbe = self.srv.create("DAO.DBEngine.36")
            self.ws = self.dbe.get("Workspaces", 0)
        return self

    def stop(self):
        if self.srv is not None:
            self.srv.close()
        self.srv = self.dbe = self.ws = self.db = None

    def username(self):
        try:
            return self.ws.get("UserName")
        except (RtError, AttributeError):
            return "admin"

    def errors(self):
        """DBEngine.Errors: [(number, source, description)]"""
        try:
            return [(e.get("Number"), e.get("Source"), e.get("Description")) for e in self.dbe.each("Errors")]
        except RtError:
            return []

    def tabledefs(self, include_sys=False):
        out = []
        for td in self.db.each("TableDefs"):
            a = td.get("Attributes")
            if (a & dbSystemObject) == 0 or include_sys:
                out.append(td)
        return out


def attached(td):
    """0, 1 (an attached table) or 2 (attached through ODBC)"""
    a = td.get("Attributes")
    if a & dbAttachedTable == dbAttachedTable:
        return 1
    if a & dbAttachedODBC == dbAttachedODBC:
        return 2
    return 0


def text_of(v):
    """a value as VisData shows it"""
    if v is None:
        return ""
    if v is EMPTY:
        return ""
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return str(int(v))
    if hasattr(v, "strftime"):
        if v.hour == v.minute == v.second == 0:
            return f"{v.month}/{v.day}/{v.year}"
        return v.strftime(f"{v.month}/{v.day}/{v.year} %I:%M:%S %p").lstrip("0").replace(" 0", " ")
    if isinstance(v, (bytes, bytearray)):
        return "(Binary)"
    return str(v)
