"""The query of the Query Designer (SQL Builder, the Data View; phase 12.3): its tables and their joins (the Diagram
pane), its columns (the Grid pane) and the SQL made from them (the SQL pane), as the Visual Database Tools of VB6 make
it for Jet (VB6 oracles Claude Enviroment/oracles/12.3): the names qualified by their table when there are more
tables, the joins in the WHERE clause, the criteria in parentheses, GROUP BY with HAVING, ORDER BY in the sort order.
parse() reads a SQL text back into the panes, or says it cannot be shown in them (VB6: the Diagram and Grid panes go
grey and the SQL is kept as it is)."""
import re

QUERY_TYPES = ["Select", "Insert", "Insert Values", "Update", "Delete", "Make Table"]
SORTS = ["", "Ascending", "Descending"]
GROUP_FUNCS = ["Group By", "Sum", "Avg", "Min", "Max", "Count", "Expression", "Where", "Sum Distinct",
               "Avg Distinct", "Min Distinct", "Max Distinct", "Count Distinct", "StDev", "StDevP", "Var", "VarP"]
TEXT_FUNCS = ["Group By", "Min", "Max", "Count", "Expression", "Where", "Min Distinct", "Max Distinct",
              "Count Distinct"]
ORS = 3                                         # the Or... columns of the grid


class QTable:
    def __init__(self, name, alias="", columns=None, x=20, y=20):
        self.name = name
        self.alias = alias
        self.columns = list(columns or [])      # the names of its columns (the database's order)
        self.x, self.y = x, y
        self.collapsed = False                  # View > Table > Name Only

    @property
    def ref(self):
        """how the SQL names it"""
        return self.alias or self.name


class Join:
    def __init__(self, left, lcol, right, rcol, op="=", kind=""):
        self.left, self.lcol, self.right, self.rcol, self.op = left, lcol, right, rcol, op
        self.kind = kind                        # "" (inner), "left", "right", "full": Select All Rows From <Table>

    def all_from(self, ref):
        """Select All Rows From <ref> is on"""
        side = "left" if ref.lower() == self.left.lower() else "right"
        return self.kind in (side, "full")

    def set_all_from(self, ref, on):
        side = "left" if ref.lower() == self.left.lower() else "right"
        other = "right" if side == "left" else "left"
        has = {k for k in (("left", "right") if self.kind == "full" else (self.kind,)) if k}
        has = (has | {side}) if on else (has - {side})
        self.kind = "full" if has == {"left", "right"} else (has.pop() if has else "")
        del other

    def cond(self):
        return f"{_q(self.left)}.{_q(self.lcol)} {self.op} {_q(self.right)}.{_q(self.rcol)}"


class Row:
    """a line of the Grid pane"""

    def __init__(self, column="", table="", alias="", output=True):
        self.column = column                    # a column of the table, "*", or an expression
        self.table = table                      # the ref of its table ("" for an expression)
        self.alias = alias
        self.output = output
        self.sort = ""                          # "", "Ascending", "Descending"
        self.order = 0                          # the Sort Order (1, 2 ...)
        self.group = "Group By"                 # with Group By: its function
        self.criteria = ""
        self.ors = [""] * ORS
        self.value = ""                         # the New Value of Update / Insert Values
        self.append = ""                        # the Append column of Insert
        self.bare = set()                       # the cells read from SQL without parentheses (0: Criteria, 1..: Or)

    def is_expr(self):
        return not self.table


class QueryModel:
    def __init__(self):
        self.tables = []
        self.joins = []
        self.rows = []
        self.qtype = "Select"
        self.group_by = False
        self.distinct = False
        self.all_columns = False                # Properties: Output all columns
        self.target = ""                        # the table of Insert / Insert Values / Update / Delete / Make Table
        self.name = ""
        self.parsed = True                      # False: the SQL could not be shown in the Diagram and Grid panes
        self.text = ""                          # the SQL kept when it could not be shown

    # ---------------------------------------------------------------- the tables
    def table(self, ref):
        r = (ref or "").lower()
        return next((t for t in self.tables if t.ref.lower() == r or (not t.alias and t.name.lower() == r)), None)

    def add_table(self, name, columns, x=None, y=None, relations=()):
        """a table dropped on the Diagram pane (a second one of the same name gets an alias: Customers_1); the joins
        of its relations with the tables already there"""
        alias = ""
        if self.table(name) is not None:
            n = 1
            while self.table(f"{name}_{n}") is not None:
                n += 1
            alias = f"{name}_{n}"
        if x is None:
            x = 20 + 200 * len(self.tables)
        if y is None:
            y = 20
        t = QTable(name, alias, columns, x, y)
        self.tables.append(t)
        for (pt, pc, ft, fc) in relations:      # (primary table, its column, foreign table, its column)
            for other in self.tables[:-1]:
                if pt.lower() == name.lower() and ft.lower() == other.name.lower():
                    self.joins.append(Join(t.ref, pc, other.ref, fc))
                elif ft.lower() == name.lower() and pt.lower() == other.name.lower():
                    self.joins.append(Join(other.ref, pc, t.ref, fc))
        self.parsed = True
        return t

    def remove_table(self, ref):
        t = self.table(ref)
        if t is None:
            return
        self.tables.remove(t)
        self.joins = [j for j in self.joins if j.left != t.ref and j.right != t.ref]
        self.rows = [r for r in self.rows if r.table != t.ref]

    def many(self):
        return len(self.tables) > 1

    # ---------------------------------------------------------------- the columns
    def row_of(self, table, column):
        return next((r for r in self.rows if r.table.lower() == table.lower() and r.column.lower() == column.lower()),
                    None)

    def checked(self, table, column):
        r = self.row_of(table, column)
        return r is not None and r.output

    def check(self, table, column, on):
        """a column checked or not in the table of the Diagram pane"""
        r = self.row_of(table, column)
        if on:
            if r is None:
                self.rows.append(Row(column, table))
            else:
                r.output = True
        elif r is not None:
            if r.sort or r.criteria or any(r.ors):
                r.output = False                # (it stays for its sort or its criteria)
            else:
                self.rows.remove(r)

    def set_sort(self, row, sort):
        """Sort Type: the Sort Order follows (the next number); Unsorted takes it away and renumbers"""
        if sort and not row.sort:
            row.order = 1 + max([r.order for r in self.rows if r.sort] + [0])
        row.sort = sort
        if not sort:
            row.order = 0
            for i, r in enumerate(sorted([r for r in self.rows if r.sort], key=lambda r: r.order)):
                r.order = i + 1

    def set_order(self, row, n):
        """Sort Order typed: the others move up or down"""
        others = sorted([r for r in self.rows if r.sort and r is not row], key=lambda r: r.order)
        n = max(1, min(n, len(others) + 1))
        others.insert(n - 1, row)
        for i, r in enumerate(others):
            r.order = i + 1

    def set_group_by(self, on):
        self.group_by = on
        for r in self.rows:
            r.group = "Group By"

    # ---------------------------------------------------------------- the SQL
    def name_of(self, row):
        if row.is_expr():
            return row.column
        if row.column == "*":
            return f"{_q(row.table)}.*" if self.many() else "*"
        return f"{_q(row.table)}.{_q(row.column)}" if self.many() else _q(row.column)

    def select_item(self, row, n):
        s = self.name_of(row)
        g = row.group if self.group_by else "Group By"
        if self.group_by and g not in ("Group By", "Expression", "Where"):
            fn, _, dist = g.partition(" ")
            fn = {"StDev": "STDEV", "StDevP": "STDEVP", "Var": "VAR", "VarP": "VARP"}.get(fn, fn.upper())
            s = f"{fn}({'DISTINCT ' if dist else ''}{s})"
            if not row.alias:
                row.alias = f"Expr{n}"
        if row.alias:
            s += f" AS {_q(row.alias)}"
        return s

    def conditions(self, row, text):
        """a criteria cell as SQL: "= 'London'" -> Customers.City = 'London' (a cell that starts with no operator:
        "= ...")"""
        t = text.strip()
        if not t:
            return ""
        name = self.name_of(row) if row.column != "*" else ""
        if re.match(r"(?i)^(=|<>|<=|>=|<|>|like\b|not\b|in\b|between\b|is\b)", t):
            return f"{name} {t}"
        return f"{name} = {t}"

    def where_parts(self, having):
        """the criteria: each Or... column an alternative, the cells of one column joined by AND"""
        alts = []
        for k in range(ORS + 1):
            cells, all_bare = [], True
            for r in self.rows:
                txt = r.criteria if k == 0 else r.ors[k - 1]
                if not txt.strip():
                    continue
                in_having = self.group_by and r.group not in ("Where",)
                if in_having != having:
                    continue
                cond = self.conditions(r, txt)
                cells.append(cond if k in r.bare else f"({cond})")   # (VB6 keeps the SQL's own parentheses)
                all_bare = all_bare and k in r.bare
            if cells:
                alts.append((cells, all_bare))
        if not alts:
            return ""
        alts = [c[0] if len(c) == 1 else " AND ".join(c) if bare or len(alts) == 1 else "(" + " AND ".join(c) + ")"
                for c, bare in alts]
        return " OR\n    ".join(alts)

    def sql(self):
        if not self.parsed:
            return self.text
        if self.qtype == "Select":
            return self.select_sql()
        if self.qtype == "Delete":
            s = f"DELETE FROM {_q(self.target or (self.tables[0].name if self.tables else ''))}"
            w = self.where_parts(False)
            return s + (f"\nWHERE {w}" if w else "")
        if self.qtype == "Update":
            sets = [f"{_q(r.column)} = {r.value}" for r in self.rows if r.value.strip() and not r.is_expr()]
            s = f"UPDATE {_q(self.target or (self.tables[0].name if self.tables else ''))}\nSET {', '.join(sets)}"
            w = self.where_parts(False)
            return s + (f"\nWHERE {w}" if w else "")
        if self.qtype == "Insert Values":
            rows = [r for r in self.rows if not r.is_expr()]
            return (f"INSERT INTO {_q(self.target or (self.tables[0].name if self.tables else ''))}\n"
                    f"    ({', '.join(_q(r.column) for r in rows)})\nVALUES ({', '.join(r.value or 'NULL' for r in rows)})")
        if self.qtype == "Insert":
            rows = [r for r in self.rows if r.output]
            cols = ", ".join(_q(r.append or r.column) for r in rows)
            return f"INSERT INTO {_q(self.target)}\n    ({cols})\n" + self.select_sql()
        if self.qtype == "Make Table":
            s = self.select_sql()
            lines = s.split("\n")
            lines.insert(1, f"INTO {_q(self.target)}")
            return "\n".join(lines)
        return ""

    def select_sql(self):
        items = []
        n = 0
        for r in self.rows:
            if r.output or (self.all_columns and False):
                n += 1
                items.append(self.select_item(r, n))
        if self.all_columns:
            items = ["*"] + items if not self.many() else [f"{_q(t.ref)}.*" for t in self.tables] + items
        s = ("SELECT " + ("DISTINCT " if self.distinct else "") + ", ".join(items)).rstrip()
        if self.tables:
            frm, used = self.from_clause()
            s += "\nFROM " + frm
        else:
            used = []
        parts = [j.cond() for j in self.joins if j not in used]
        w = self.where_parts(False)
        if w:
            parts.append(w if not parts or " OR\n" not in w else f"({w})")
        if parts:
            s += "\nWHERE " + " AND\n    ".join(parts)
        if self.group_by:
            gb = [self.name_of(r) for r in self.rows if r.group == "Group By" and r.column != "*"]
            if gb:
                s += "\nGROUP BY " + ", ".join(gb)
            h = self.where_parts(True)
            if h:
                s += "\nHAVING " + h
        srt = sorted([r for r in self.rows if r.sort], key=lambda r: r.order)
        if srt:
            s += "\nORDER BY " + ", ".join(self.name_of(r) + (" DESC" if r.sort == "Descending" else "")
                                           for r in srt)
        return s

    def from_clause(self):
        """the tables of FROM: separated by commas, the joins in WHERE (Jet, as VB6 writes them); with an outer join
        (Select All Rows From) the joined tables in JOIN ... ON: ([the FROM text], [the joins written there])"""
        def tname(t):
            return _q(t.name) + (f" {_q(t.alias)}" if t.alias else "")
        if not any(j.kind for j in self.joins):
            return ", ".join(tname(t) for t in self.tables), []
        first = next(j for j in self.joins if j.kind)
        start = self.table(first.left)
        expr, inside, used, n = tname(start), {start.ref.lower()}, [], 0
        while True:
            j = next((j for j in self.joins if j not in used and
                      ((j.left.lower() in inside) != (j.right.lower() in inside))), None)
            if j is None:
                break
            new_ref = j.right if j.left.lower() in inside else j.left
            kind = j.kind
            if new_ref == j.left and kind in ("left", "right"):   # the table comes first in the join
                kind = "right" if kind == "left" else "left"
            word = {"": "INNER JOIN", "left": "LEFT OUTER JOIN", "right": "RIGHT OUTER JOIN",
                    "full": "FULL OUTER JOIN"}[kind]
            ons = [j] + [k for k in self.joins if k is not j and k not in used and
                         {k.left.lower(), k.right.lower()} == {j.left.lower(), j.right.lower()}]
            if n:
                expr = f"({expr})"
            expr += f" {word} {tname(self.table(new_ref))} ON " + " AND ".join(k.cond() for k in ons)
            used += ons
            inside.add(new_ref.lower())
            n += 1
        rest = [tname(t) for t in self.tables if t.ref.lower() not in inside]
        return ", ".join([expr] + rest), used

    # ---------------------------------------------------------------- the SQL read back
    def parse(self, text, columns_of=None):
        """the SQL typed in the SQL pane: the tables, joins and rows; False if it cannot be shown in the Diagram and
        Grid panes (it is kept as it is). columns_of(table) gives the columns of a table of the database."""
        self.text = text
        try:
            m = _parse_select(text)
        except ValueError:
            m = None
        if m is None:
            self.parsed = False
            return False
        old = {t.name.lower(): (t.x, t.y) for t in self.tables}
        self.tables, self.joins, self.rows = [], [], []
        self.qtype, self.distinct = "Select", m["distinct"]
        kinds = m.get("kinds", {})
        for name, alias in m["from"]:
            x, y = old.get(name.lower(), (None, None))
            t = QTable(name, alias, columns_of(name) if columns_of else [], 20 + 200 * len(self.tables) if x is None
                       else x, 20 if y is None else y)
            self.tables.append(t)
        self.group_by = bool(m["group"])

        def split_name(expr):
            e = expr.strip()
            mm = re.fullmatch(r"(?:(\[[^\]]+\]|\w+)\.)?(\[[^\]]+\]|\w+|\*)", e)
            if not mm:
                return None, e
            if mm.group(2) == "*" and not mm.group(1):      # SELECT *: no table in the Grid pane (VB6)
                return None, "*"
            tref = _uq(mm.group(1)) if mm.group(1) else (self.tables[0].ref if len(self.tables) == 1 else None)
            if tref is None or self.table(tref) is None:
                return None, e
            t, col = self.table(tref), _uq(mm.group(2))
            if t.columns and col != "*" and col.lower() not in (c.lower() for c in t.columns):
                return None, e                  # not a column of the table: an expression (VB6: no Table)
            return t.ref, next((c for c in t.columns if c.lower() == col.lower()), col)
        for item in m["select"]:
            expr, alias = item
            fm = re.fullmatch(r"(?i)(SUM|AVG|MIN|MAX|COUNT|STDEV|STDEVP|VAR|VARP)\s*\(\s*(DISTINCT\s+)?(.+)\)", expr.strip())
            func = None
            if fm and self.group_by:
                func = {"STDEV": "StDev", "STDEVP": "StDevP", "VAR": "Var", "VARP": "VarP"}.get(
                    fm.group(1).upper(), fm.group(1).capitalize()) + (" Distinct" if fm.group(2) else "")
                expr = fm.group(3)
            t, c = split_name(expr)
            r = Row(c, t or "", alias or "")
            if func:
                r.group = func
            self.rows.append(r)
        pending = []
        for cond in m["where"]:                  # the joins; the criteria after the sort (VB6: its rows after)
            jm = re.fullmatch(r"\s*((?:\[[^\]]+\]|\w+))\.((?:\[[^\]]+\]|\w+))\s*(=|<>|<=|>=|<|>)\s*"
                              r"((?:\[[^\]]+\]|\w+))\.((?:\[[^\]]+\]|\w+))\s*", cond)
            if jm and self.table(_uq(jm.group(1))) is not None and self.table(_uq(jm.group(4))) is not None and \
                    _uq(jm.group(1)).lower() != _uq(jm.group(4)).lower():
                j = Join(self.table(_uq(jm.group(1))).ref, _uq(jm.group(2)), self.table(_uq(jm.group(4))).ref,
                         _uq(jm.group(5)), jm.group(3))
                k = kinds.get(cond)
                if k:                           # (the side as the join named its tables)
                    a, b = k[1].lower(), k[2].lower()
                    kk = k[0]
                    if kk in ("left", "right") and j.left.lower() == b and j.right.lower() == a:
                        kk = "right" if kk == "left" else "left"
                    j.kind = kk
                self.joins.append(j)
                continue
            pending.append(cond)
        for n, (expr, desc) in enumerate(m["order"]):
            t, c = split_name(expr)
            r = self.row_of(t or "", c)
            if r is None:
                r = Row(c, t or "", output=False)
                self.rows.append(r)
            r.sort = "Descending" if desc else "Ascending"
            r.order = n + 1
        for cond in pending:
            if not self._criteria(cond, split_name, False):
                self.parsed = False
                return False
        for cond in m["having"]:
            if not self._criteria(cond, split_name, True):
                self.parsed = False
                return False
        self.parsed = True
        return True

    def _criteria(self, cond, split_name, having):
        """a condition of WHERE / HAVING: one of the criteria of a row (an alternative with OR: the Or... columns)"""
        alts = _split_top(cond, "OR")
        if len(alts) == 1 and len(_split_top(_strip_parens(cond), "OR")) > 1:   # ((a) AND (b)) OR (c) in parentheses
            alts = _split_top(_strip_parens(cond), "OR")
        if len(alts) > ORS + 1:
            return False
        for k, alt in enumerate(alts):
            parts = _split_top(_strip_parens(alt), "AND")
            for part in parts:
                p = _strip_parens(part)
                bare = p == part.strip() and (len(parts) == 1 and _strip_parens(alt) == alt.strip())
                mm = re.fullmatch(r"(?is)\s*((?:(?:\[[^\]]+\]|\w+)\.)?(?:\[[^\]]+\]|\w+))\s*"
                                  r"(=|<>|<=|>=|<|>|LIKE\b|NOT LIKE\b|IN\b|NOT IN\b|BETWEEN\b|IS\b)(.*)", p)
                if not mm:
                    return False
                t, c = split_name(mm.group(1))
                if t is None and not self.group_by:
                    return False
                r = self.row_of(t or "", c)
                if r is None:
                    r = Row(c, t or "", output=False)
                    if self.group_by and not having:
                        r.group = "Where"
                    self.rows.append(r)
                op = mm.group(2).strip().upper() if mm.group(2).strip().isalpha() or " " in mm.group(2).strip() \
                    else mm.group(2).strip()
                cell = f"{op} {mm.group(3).strip()}".strip()
                if k == 0:
                    r.criteria = (r.criteria + " AND " + cell) if r.criteria else cell
                else:
                    r.ors[k - 1] = cell
                if bare:
                    r.bare.add(k)
                else:
                    r.bare.discard(k)
        return True


# ---------------------------------------------------------------- the SQL text
def _q(name):
    """a name in brackets when it needs them (Order Details -> [Order Details])"""
    if not name or re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*|\*", name):
        return name
    return f"[{name}]"


def _uq(name):
    return name[1:-1] if name and name.startswith("[") and name.endswith("]") else name


def _strip_parens(s):
    s = s.strip()
    while s.startswith("(") and s.endswith(")") and _balanced(s[1:-1]):
        s = s[1:-1].strip()
    return s


def _balanced(s):
    depth, quote = 0, None
    for ch in s:
        if quote:
            if ch == quote:
                quote = None
            continue
        if ch in "'\"":
            quote = ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth < 0:
                return False
    return depth == 0


def _split_top(s, word):
    """s split at a keyword (AND / OR) outside parentheses and quotes"""
    out, depth, quote, start, i = [], 0, None, 0, 0
    w = word.upper()
    U = s.upper()
    while i < len(s):
        ch = s[i]
        if quote:
            if ch == quote:
                quote = None
        elif ch in "'\"[":
            quote = "]" if ch == "[" else ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif depth == 0 and U.startswith(w, i) and (i == 0 or not (s[i - 1].isalnum() or s[i - 1] == "_")) and \
                (i + len(w) >= len(s) or not (s[i + len(w)].isalnum() or s[i + len(w)] == "_")):
            if w == "AND" and re.search(r"(?i)\bBETWEEN\s+\S+\s*$", s[start:i]):
                i += len(w)
                continue
            out.append(s[start:i].strip())
            start = i + len(w)
            i += len(w)
            continue
        i += 1
    out.append(s[start:].strip())
    return [x for x in out if x]


def _split_commas(s):
    out, depth, quote, start = [], 0, None, 0
    for i, ch in enumerate(s):
        if quote:
            if ch == quote:
                quote = None
        elif ch in "'\"[":
            quote = "]" if ch == "[" else ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            out.append(s[start:i].strip())
            start = i + 1
    out.append(s[start:].strip())
    return [x for x in out if x]


KEYWORDS = ["SELECT", "FROM", "WHERE", "GROUP BY", "HAVING", "ORDER BY"]


def _clauses(text):
    """the clauses of a SELECT, found outside quotes, brackets and parentheses"""
    t = " ".join(text.replace("\r", " ").replace("\n", " ").split())
    pos, depth, quote, i = [], 0, None, 0
    U = t.upper()
    while i < len(t):
        ch = t[i]
        if quote:
            if ch == quote:
                quote = None
            i += 1
            continue
        if ch in "'\"[":
            quote = "]" if ch == "[" else ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif depth == 0 and (i == 0 or t[i - 1] == " "):
            for k in KEYWORDS:
                if U.startswith(k, i) and (i + len(k) == len(t) or not t[i + len(k)].isalnum()):
                    pos.append((i, k))
                    i += len(k)
                    break
            else:
                i += 1
            continue
        i += 1
    if not pos or pos[0][1] != "SELECT" or pos[0][0] != 0:
        return None
    out = {}
    for n, (i, k) in enumerate(pos):
        end = pos[n + 1][0] if n + 1 < len(pos) else len(t)
        if k in out:
            return None
        out[k] = t[i + len(k):end].strip()
    return out


def _parse_select(text):
    c = _clauses(text)
    if c is None or "FROM" not in c:
        return None
    sel = c["SELECT"]
    distinct = False
    if re.match(r"(?i)DISTINCT\s", sel):
        distinct, sel = True, sel[9:].strip()
    if re.match(r"(?i)TOP\s", sel):
        return None
    items = []
    for it in _split_commas(sel):
        m = re.fullmatch(r"(?is)(.+?)\s+AS\s+(\[[^\]]+\]|\w+)", it)
        items.append((m.group(1).strip(), _uq(m.group(2))) if m else (it, ""))
    frm, joins, kinds = [], [], {}
    for it in _split_commas(c["FROM"]):
        if re.search(r"(?i)\bJOIN\b", it):        # A INNER JOIN B ON A.x = B.y ...: the tables and the conditions
            pieces = re.split(r"(?i)\s+(INNER|LEFT(?:\s+OUTER)?|RIGHT(?:\s+OUTER)?|FULL(?:\s+OUTER)?)?\s*JOIN\s+",
                              " " + it.strip())
            first = pieces[0].strip()
            m = re.fullmatch(r"\(*\s*(\[[^\]]+\]|\w+)(?:\s+(?:AS\s+)?(\[[^\]]+\]|\w+))?", first)
            if not m:
                return None
            frm.append((_uq(m.group(1)), _uq(m.group(2)) if m.group(2) else ""))
            for word, p in zip(pieces[1::2], pieces[2::2]):
                m = re.fullmatch(r"(?is)(\[[^\]]+\]|\w+)(?:\s+(?:AS\s+)?(?!ON\b)(\[[^\]]+\]|\w+))?\s+ON\s+(.+?)\)*",
                                 p.strip())
                if not m:
                    return None
                prev = frm[-1][1] or frm[-1][0]
                frm.append((_uq(m.group(1)), _uq(m.group(2)) if m.group(2) else ""))
                kind = (word or "").split()[0].lower() if word else ""
                kind = "" if kind == "inner" else kind
                new = frm[-1][1] or frm[-1][0]
                for cnd in _split_top(_strip_parens(m.group(3)), "AND"):
                    joins.append(cnd)
                    if kind:
                        kinds[cnd] = (kind, prev, new)
            continue
        m = re.fullmatch(r"(\[[^\]]+\]|\w+)(?:\s+(?:AS\s+)?(\[[^\]]+\]|\w+))?", it.strip())
        if not m:
            return None
        frm.append((_uq(m.group(1)), _uq(m.group(2)) if m.group(2) else ""))
    where = joins + (_split_top(c.get("WHERE", ""), "AND") if c.get("WHERE") else [])
    # (a criteria with OR at the top: one condition, kept whole)
    if c.get("WHERE") and len(_split_top(c["WHERE"], "OR")) > 1:
        where = joins + [c["WHERE"]]
    group = _split_commas(c["GROUP BY"]) if c.get("GROUP BY") else []
    having = [c["HAVING"]] if c.get("HAVING") else []
    order = []
    for it in _split_commas(c.get("ORDER BY", "")):
        m = re.fullmatch(r"(?i)(.+?)(\s+(ASC|DESC))?", it.strip())
        order.append((m.group(1).strip(), bool(m.group(3) and m.group(3).upper() == "DESC")))
    return {"distinct": distinct, "select": items, "from": frm, "where": where, "group": group, "having": having,
            "order": order, "kinds": kinds}
