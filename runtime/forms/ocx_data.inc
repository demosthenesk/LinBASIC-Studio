/* ocx_data.inc - the data-bound controls of VB6 over ADO (phase 4.4): DataGrid (MSDataGridLib),
 * DataList and DataCombo (MSDataListLib). They read the rows of an ADO Recordset (of an ADO Data
 * control, a Recordset Set to DataSource / RowSource, or a class that is a data source) through
 * vb_dbhooks, and follow it by listening to it. Included by vbforms.cpp after data.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

/* a cell as the grid and the lists show it: Null is empty, True is -1 */
static QString cell_text(vbvar v)
{
    if (v.vt == VT_NULL || v.vt == VT_EMPTY) return QString();
    if (v.vt == VT_BOOL) return v.u.b ? QString("-1") : QString("0");
    vb_i32 save = vb_errpend;
    QString s = QString::fromUtf8(vb_stoutf8(vb_v2s(v)));
    if (vb_errpend && !save) { vb_err_clear_pending(); return QString(); }
    return s;
}
/* the rows of a Recordset, ADO or DAO (the Data control: DBGrid, DBList, DBCombo) */
static bool rs_is_dao(vbobj rs) { return rs && vb_dbhooks.is_dao_rs && vb_dbhooks.is_dao_rs(rs); }
static int rsx_shape(vbobj rs, int *nc, int *cur)
{
    if (rs_is_dao(rs)) return vb_dbhooks.dao_shape(rs, nc, cur);
    if (!rs || !vb_dbhooks.ado_shape) { if (nc) *nc = 0; if (cur) *cur = -1; return -1; }
    return vb_dbhooks.ado_shape(rs, nc, cur);
}
static vbvar rsx_cell(vbobj rs, int row, int col)
{
    if (rs_is_dao(rs)) return vb_dbhooks.dao_cell(rs, row, col);
    return rs && vb_dbhooks.ado_cell ? vb_dbhooks.ado_cell(rs, row, col) : vb_v_empty();
}
static const char *rsx_colname(vbobj rs, int col)
{
    if (rs_is_dao(rs)) return vb_dbhooks.dao_colname(rs, col);
    return rs && vb_dbhooks.ado_colname ? vb_dbhooks.ado_colname(rs, col) : "";
}
static int rsx_coltype(vbobj rs, int col)
{
    if (rs_is_dao(rs)) return vb_dbhooks.dao_coltype(rs, col);
    return rs && vb_dbhooks.ado_coltype ? vb_dbhooks.ado_coltype(rs, col) : 0;
}
static int rs_col_index(vbobj rs, const QString &name)
{
    int nc = 0;
    if (rsx_shape(rs, &nc, nullptr) < 0) return -1;
    for (int i = 0; i < nc; i++) if (!name.compare(QString::fromUtf8(rsx_colname(rs, i)), Qt::CaseInsensitive)) return i;
    return -1;
}

/* ---------------------------------------------------------------- DataGrid
 * Its columns (VB6 oracles f_dfmt .. f_dfmt4): those of the .frm (ColumnCount, ColumnNN with their DataFormat, Split0)
 * until the grid is bound to a rowset; then one column for each field, unless HoldFields kept the layout. A
 * column shows its value through its DataFormat (NumberFormat is the Format of it); Column.Text = writes the field
 * (the type of the field with the regional settings; 6160 when it does not go). DataSource of the .frm: the
 * Bindings of the .frx (dispid 42). */
struct OcxDataGrid;
static QString bindings_for(const QByteArray &b, int dispid);    /* ocx_dbgrid.inc */
struct GridColumn {
    int col; QString caption, field; int him = -1; bool visible = true, locked = false; int align = 0;
    vbobj fmt = nullptr;                        /* DataFormat (counted), or none yet */
    vbobj obj = nullptr;
};
/* the width of a column as the grid keeps it: whole pixels, in HIMETRIC; Width gives it back in twips with a
 * multiplication of Singles (VB6 oracles f_dfmt2, f_dfmt3: 1614 -> 915,0237, 2646 -> 1500,095) */
static int grid_px(int px) { return (int)std::lround(px * 2540.0 / 96.0); }
static int grid_width(double tw) { return grid_px((int)std::lround(tw / 15.0)); }
static float grid_twips(int him)
{
    volatile float k = (float)(1440.0 / 2540.0), r = (float)him * k;
    return r;
}
/* the width of a new column of a field, by its type (VB6 oracle f_dfmt2, MS Sans Serif 8) */
static int grid_auto_width(int adotype)
{
    switch (adotype) {
    case 2: case 16: case 17: return grid_px(41);   /* adSmallInt, adTinyInt, adUnsignedTinyInt */
    case 3: return grid_px(61);                      /* adInteger */
    case 11: return grid_px(51);                     /* adBoolean */
    default: return grid_px(116);
    }
}
static void gcol_free(GridColumn &g)
{
    if (g.obj) { g.obj->ext = nullptr; vb_orelease(g.obj); g.obj = nullptr; }
    if (g.fmt) { vb_orelease(g.fmt); g.fmt = nullptr; }
}
static vbvar gcol_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vbvar gcols_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vb_i32 gcols_count(vbobj o);
static vbvar gcols_item(vbobj o, vb_i32 i);
static const vb_class CLS_GCOL = { "Column", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "text", nullptr, nullptr, nullptr, gcol_fallback, nullptr };
static const vb_class CLS_GCOLS = { "Columns", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "item", gcols_count, gcols_item, nullptr, gcols_fallback, nullptr };

struct OcxDataGrid : Ocx {
    vbobj src = nullptr;                        /* DataSource (counted) */
    vbobj rs = nullptr;                         /* its Recordset (counted) */
    QString dmember, caption;
    std::vector<GridColumn> cols;
    vbobj colsobj = nullptr;
    QPointer<QTableWidget> tw;
    bool allowadd = false, allowdelete = false, allowupdate = true, recsel = true;
    int headlines = 1, curcol = 0, lastrow = -1, lastcol = -1;
    double rowheight = 0;
    bool building = false;
    ~OcxDataGrid() override
    {
        listen(false);
        if (rs) vb_orelease(rs);
        if (src) vb_orelease(src);
        for (auto &g : cols) gcol_free(g);
        if (colsobj) { colsobj->ext = nullptr; vb_orelease(colsobj); }
    }
    bool held = false;                          /* HoldFields: the columns stay when the grid binds again */
    QString errtext;
    bool data_aware() override { return true; }
    bool own_binding() override { return true; }
    QString srcname;                            /* DataSource = "Adodc1" in the .frm */
    void source_reset(vbobj s) override
    {
        if (!src && !srcname.isEmpty() && is_ctl_obj(s) && !srcname.compare(C(s)->name, Qt::CaseInsensitive)) { src = vb_oid(s); vb_oaddref(src); }
        if (src && vb_oid(src) == vb_oid(s)) bind();
    }
    QWidget *make(QWidget *parent) override
    {
        tw = new QTableWidget(parent);
        tw->setSelectionMode(QAbstractItemView::SingleSelection);
        tw->setEditTriggers(QAbstractItemView::NoEditTriggers);
        tw->verticalHeader()->setDefaultSectionSize(18);
        vbobj self = c->self;
        QObject::connect(tw, &QTableWidget::currentCellChanged, [self](int r, int col, int, int) {
            Ctl *x = C(self);
            auto *g = x ? dynamic_cast<OcxDataGrid *>(x->ocx) : nullptr;
            if (!g || g->building || r < 0) return;
            g->user_move(r, col);
        });
        QObject::connect(tw, &QTableWidget::cellClicked, [self](int, int) { Ctl *x = C(self); if (x) fire_ctl(x, "click"); });
        QObject::connect(tw, &QTableWidget::cellDoubleClicked, [self](int, int) { Ctl *x = C(self); if (x) fire_ctl(x, "dblclick"); });
        QObject::connect(tw->horizontalHeader(), &QHeaderView::sectionClicked, [self](int i) {
            Ctl *x = C(self);
            if (!x) return;
            vbvar a = vb_v_i((vb_i16)i);
            fire_ctl(x, "headclick", 1, &a);
        });
        return tw;
    }
    void load(const char *const *p) override
    {
        QString v;
        if (prop(p, "Caption", &v)) caption = v;
        if (prop(p, "DataSource", &v)) srcname = v;
        else if (prop(p, "Bindings", &v)) srcname = bindings_for(QByteArray::fromHex(v.toLatin1()), 42);
        if (prop(p, "DataMember", &v)) dmember = v;
        int n = (int)propn(p, "ColumnCount", 0);   /* the columns of the design */
        for (int i = 0; i < n; i++) {
            QString pre = QString("Column%1.").arg(i, 2, 10, QChar('0')), sp = "Split0." + pre;
            GridColumn g; g.col = -1;
            prop(p, (pre + "DataField").toUtf8().constData(), &g.field);
            prop(p, (pre + "Caption").toUtf8().constData(), &g.caption);
            g.fmt = dfmt_props(p, pre);
            double w = propn(p, (sp + "ColumnWidth").toUtf8().constData(), -1);
            g.him = w >= 0 ? grid_width(w) : grid_px(101);
            g.align = (int)propn(p, (sp + "Alignment").toUtf8().constData(), 0);
            g.locked = propn(p, (sp + "Locked").toUtf8().constData(), 0) != 0;
            g.visible = propn(p, (sp + "Object.Visible").toUtf8().constData(), propn(p, (sp + "Visible").toUtf8().constData(), -1)) != 0;
            cols.push_back(g);
        }
        allowadd = propn(p, "AllowAddNew", 0) != 0;
        allowdelete = propn(p, "AllowDelete", 0) != 0;
        allowupdate = propn(p, "AllowUpdate", -1) != 0;
        headlines = (int)propn(p, "HeadLines", 1);
        recsel = propn(p, "RecordSelectors", -1) != 0;
    }
    static void grid_listen(void *ctx, vbobj, const char *ev, int, vbvar *)
    {
        Ctl *x = C((vbobj)ctx);
        auto *g = x ? dynamic_cast<OcxDataGrid *>(x->ocx) : nullptr;
        if (!g) return;
        if (!strcmp(ev, "lb_fill")) g->rebuild();
    }
    void listen(bool on)
    {
        if (rs && !rs_is_dao(rs) && vb_dbhooks.ado_listen) vb_dbhooks.ado_listen(rs, grid_listen, c->self, on ? 1 : 0);
    }
    void bind()
    {
        listen(false);
        if (rs) { vb_orelease(rs); rs = nullptr; }
        vbobj r = source_rs(src, dmember);
        if (r) { rs = r; vb_oaddref(rs); listen(true); }
        make_columns();
        rebuild();
    }
    /* the columns for the rowset: one for each field, or the held ones (HoldFields) on their fields */
    void make_columns()
    {
        if (held) {
            for (auto &g : cols) g.col = rs && !g.field.isEmpty() ? rs_col_index(rs, g.field) : -1;
            return;
        }
        for (auto &g : cols) gcol_free(g);
        cols.clear();
        int nc = 0;
        if (rs && rsx_shape(rs, &nc, nullptr) >= 0)
            for (int i = 0; i < nc; i++) {
                GridColumn g; g.col = i;
                g.caption = g.field = QString::fromUtf8(rsx_colname(rs, i));
                g.him = grid_auto_width(rsx_coltype(rs, i));
                cols.push_back(g);
            }
    }
    void blank_columns()                        /* ClearFields without a rowset: the two columns of a new grid */
    {
        for (auto &g : cols) gcol_free(g);
        cols.clear();
        for (int i = 0; i < 2; i++) { GridColumn g; g.col = -1; g.him = grid_px(101); cols.push_back(g); }
    }
    /* a cell as the column shows it: through its DataFormat */
    QString shown(GridColumn &g, vbvar v)
    {
        if (!g.fmt || vb_dfmt_kind(g.fmt) <= 0) return cell_text(v);
        vb_i32 save = vb_errpend;
        vbvar r = vb_dfmt_format(g.fmt, v, g.obj);
        if (vb_errpend && !save) { vb_err_clear_pending(); return cell_text(v); }
        return r.vt == VT_BOOL ? cell_text(r) : cell_text(r);
    }
    int current_row()
    {
        int cur = -1;
        if (rs) rsx_shape(rs, nullptr, &cur);
        return cur;
    }
    void rebuild()
    {
        if (!tw) return;
        building = true;
        int nc = 0, cur = -1;
        int n = rs ? rsx_shape(rs, &nc, &cur) : -1;
        if (n < 0) n = 0;
        if ((int)cols.size() != nc && rs && !held) make_columns();
        tw->clear();
        tw->setColumnCount((int)cols.size());
        QStringList heads;
        for (auto &g : cols) heads << g.caption;
        tw->setHorizontalHeaderLabels(heads);
        int rows = std::min(n, 5000);
        tw->setRowCount(rows + (cur == n ? 1 : 0));
        vb_i32 m = vb_mark();
        for (int r = 0; r < rows; r++)
            for (int k = 0; k < (int)cols.size(); k++)
                tw->setItem(r, k, new QTableWidgetItem(cols[k].col < 0 ? QString() : shown(cols[k], rsx_cell(rs, r, cols[k].col))));
        if (cur == n)
            for (int k = 0; k < (int)cols.size(); k++)
                tw->setItem(rows, k, new QTableWidgetItem(cols[k].col < 0 ? QString() : shown(cols[k], rsx_cell(rs, -1, cols[k].col))));
        vb_release(m);
        for (int k = 0; k < (int)cols.size(); k++) tw->setColumnHidden(k, !cols[k].visible);
        tw->verticalHeader()->setVisible(recsel);
        if (cur >= 0 && cur < tw->rowCount()) tw->setCurrentCell(cur, std::min(curcol, std::max(0, (int)cols.size() - 1)));
        building = false;
        int row = cur;
        if (row != lastrow || curcol != lastcol) {
            vbvar a[2] = { vb_v_l(lastrow < 0 ? 0 : lastrow), vb_v_i((vb_i16)(lastcol < 0 ? 0 : lastcol)) };
            lastrow = row; lastcol = curcol;
            if (c->form) fire_ctl(c, "rowcolchange", 2, a);
        }
    }
    void user_move(int r, int col)
    {
        curcol = col;
        if (rs && !rs_is_dao(rs) && vb_dbhooks.ado_goto) {
            vb_dbhooks.ado_goto(rs, r);
            if (vb_errpend) vb_unhandled();
        }
    }
    QString col_text(int k)
    {
        if (!rs || k < 0 || k >= (int)cols.size() || cols[k].col < 0) return QString();
        int cur = current_row();
        if (cur < 0) return QString();
        vb_i32 m = vb_mark();
        QString s = shown(cols[k], rsx_cell(rs, cur == rsx_shape(rs, nullptr, nullptr) ? -1 : cur, cols[k].col));
        vb_release(m);
        return s;
    }
    vbvar col_raw(int k)
    {
        if (!rs || k < 0 || k >= (int)cols.size() || cols[k].col < 0) return vb_v_empty();
        int cur = current_row();
        if (cur < 0) return vb_v_empty();
        return rsx_cell(rs, cur == rsx_shape(rs, nullptr, nullptr) ? -1 : cur, cols[k].col);
    }
    /* Column.Value: the text of the field without the format (a String, as VB6 gives it), Null for Null */
    vbvar col_value(int k)
    {
        if (!rs || k < 0 || k >= (int)cols.size()) return vb_v_empty();
        if (cols[k].col < 0) return vstr(QString());
        vbvar v = col_raw(k);
        if (v.vt == VT_NULL) return v;          /* VB6: error 13 (not here) */
        return vstr(cell_text(v));
    }
    /* Column.Text = : the field of the current record (through the DataFormat, then its type) */
    void col_set(int k, vbvar val)
    {
        GridColumn &g = cols[k];
        if (!rs || g.col < 0 || rs_is_dao(rs) || current_row() < 0) return;
        vb_i32 m = vb_mark();
        vbvar v = val;
        bool bad = false;
        if (g.fmt && vb_dfmt_kind(g.fmt) > 0) { vbvar u; if (vb_dfmt_unformat(g.fmt, v, g.obj, &u)) bad = true; else v = u; }
        vbvar w;
        if (!bad && !vb_errpend) { if (vb_db_strto(v, rsx_coltype(rs, g.col), &w)) bad = true; else v = w; }
        if (!bad && !vb_errpend) {
            vbvar a[2] = { vstr(g.field), v };
            vb_late_o(rs, "fields", VBM_LET, 2, a);
            if (vb_errpend) { vb_err_clear_pending(); bad = true; }
        }
        vb_release(m);
        if (bad && !vb_errpend) { vb_raise_full(6160, vs(QString("DataGrid")), vs(QString("Data access error"))); return; }
        rebuild();
    }
    bool no_rowset()                            /* the Text / Value of a column without a rowset: 7005 */
    {
        if (rs) return false;
        vb_raise_full(7005, vs(QString("DataGrid")), vs(QString("Rowset not available")));
        return true;
    }
    vbobj columns()
    {
        if (!colsobj) { colsobj = vb_onew(&CLS_GCOLS); colsobj->ext = this; vb_oaddref(colsobj); }
        return colsobj;
    }
    vbobj column(int k)
    {
        GridColumn &g = cols[k];
        if (!g.obj) { g.obj = vb_onew(&CLS_GCOL); g.obj->ext = this; vb_oaddref(g.obj); }
        return g.obj;
    }
    int col_of_obj(vbobj o)
    {
        for (int k = 0; k < (int)cols.size(); k++) if (cols[k].obj == o) return k;
        return -1;
    }
    int find_col(vbvar key)
    {
        vbvar k = vb_v_letval(key);
        if (k.vt == VT_BSTR) {
            QString s = qs(k.u.s);
            for (int i = 0; i < (int)cols.size(); i++)
                if (!s.compare(cols[i].field, Qt::CaseInsensitive) || !s.compare(cols[i].caption, Qt::CaseInsensitive)) return i;
            return -1;
        }
        vb_i32 i = vb_v2l(k);
        return i >= 0 && i < (vb_i32)cols.size() ? i : -1;
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        if (n == "datasource") {
            if (set) {
                vbobj o = vb_v2obj(val);
                if (vb_errpend) return vb_v_empty();
                if (o) vb_oaddref(o);
                if (src) vb_orelease(src);
                src = o ? vb_oid(o) : nullptr;
                bind();
                return vb_v_empty();
            }
            return vb_vtmp(vb_v_obj(src));
        }
        if (n == "datamember") { if (set) { dmember = data_qs(val); if (src) bind(); return vb_v_empty(); } return vstr(dmember); }
        if (n == "columns") {
            if (argc == 0 || set) return retobj(columns());
            int k = find_col(argv[0]);
            if (k < 0) { vb_raise_full(6147, vs(QString("DataGrid")), vs(QString("Column not found"))); return vb_v_empty(); }
            return retobj(column(k));
        }
        if (n == "row") {
            if (set) { int r = vb_v2l(val); if (rs && !rs_is_dao(rs)) { vb_dbhooks.ado_goto(rs, r); } return vb_v_empty(); }
            return vb_v_i((vb_i16)current_row());
        }
        if (n == "col") {
            if (set) { curcol = vb_v2l(val); if (tw && current_row() >= 0) { building = true; tw->setCurrentCell(current_row(), curcol); building = false; } return vb_v_empty(); }
            return vb_v_i((vb_i16)curcol);
        }
        if (n == "text") { if (set) return vb_v_empty(); return vstr(col_text(curcol)); }
        if (n == "approxcount") { int k = rs ? rsx_shape(rs, nullptr, nullptr) : 0; return vb_v_l(k < 0 ? 0 : k); }
        if (n == "bookmark") {
            if (set) { if (rs && !rs_is_dao(rs)) vb_dbhooks.ado_goto(rs, (int)vb_v2d(val) - 1); return vb_v_empty(); }
            int cur = current_row();
            return cur < 0 ? vb_v_null() : vb_v_d(cur + 1);
        }
        if (n == "firstrow") return vb_v_d(1);
        if (n == "leftcol") return vb_v_l(0);
        if (n == "visiblerows") return vb_v_l(tw ? std::max(0, tw->viewport()->height() / 18) : 0);
        if (n == "visiblecols") return vb_v_l((vb_i32)cols.size());
        if (n == "caption") { if (set) { caption = data_qs(val); return vb_v_empty(); } return vstr(caption); }
        bool *bp = n == "allowaddnew" ? &allowadd : n == "allowdelete" ? &allowdelete : n == "allowupdate" ? &allowupdate :
                   n == "recordselectors" ? &recsel : nullptr;
        if (bp) { if (set) { *bp = vb_v2bool(val) != 0; rebuild(); return vb_v_empty(); } return vbool(*bp); }
        if (n == "headlines") { if (set) { headlines = vb_v2l(val); return vb_v_empty(); } return vb_v_f((float)headlines); }
        if (n == "rowheight") { if (set) { rowheight = vb_v2d(val); return vb_v_empty(); } return vb_v_f((float)(rowheight ? rowheight : 225)); }
        if (n == "refresh") { rebuild(); return vb_v_empty(); }
        if (n == "rebind") {
            bind();
            if (!rs) {                          /* no rowset: Error 4097, the columns of a new grid */
                blank_columns();
                rebuild();
                errtext = "Cannot initialize data bindings";
                vbvar a[2] = { vb_v_i(4097), vb_v_i(1) };
                fire_ctl(c, "error", 2, a);
                vb_raise_full(4097, vs(QString("DataGrid")), vs(errtext));
            }
            return vb_v_empty();
        }
        if (n == "clearfields") {
            held = false;
            if (rs) make_columns(); else blank_columns();
            rebuild();
            return vb_v_empty();
        }
        if (n == "holdfields") { held = true; return vb_v_empty(); }
        if (n == "errortext") { if (set) { errtext = data_qs(val); return vb_v_empty(); } return vstr(errtext); }
        if (n == "scroll" || n == "aboutbox" || n == "clearselcols") return vb_v_empty();
        if (n == "getbookmark") { int cur = current_row(); return cur < 0 ? vb_v_null() : vb_v_d(cur + 1 + arg_l(argc, argv, 0, 0)); }
        if (n == "rowbookmark") return vb_v_d(arg_l(argc, argv, 0, 0) + 1);
        if (n == "editactive" || n == "currentcellmodified" || n == "datachanged") return vbool(false);
        if (n == "selbookmarks" || n == "splits") return retobj(columns());
        *found = 0;
        return vb_v_empty();
    }
};
static OcxDataGrid *grid_of_obj(vbobj o) { o = vb_oid(o); return o ? (OcxDataGrid *)o->ext : nullptr; }
static vb_i32 gcols_count(vbobj o) { auto *g = grid_of_obj(o); return g ? (vb_i32)g->cols.size() : 0; }
static vbvar gcols_item(vbobj o, vb_i32 i) { auto *g = grid_of_obj(o); return g && i >= 0 && i < (vb_i32)g->cols.size() ? retobj(g->column(i)) : vb_v_empty(); }
static vbvar gcols_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    auto *g = grid_of_obj(o);
    std::string n = lower(name);
    *found = 1;
    if (!g) { vb_raise(91); return vb_v_empty(); }
    if (n == "count") return vb_v_l((vb_i32)g->cols.size());
    if (n == "item" || n.empty()) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int k = g->find_col(argv[0]);
        if (k < 0) { vb_raise_full(6147, vs(QString("DataGrid")), vs(QString("Column not found"))); return vb_v_empty(); }
        return retobj(g->column(k));
    }
    if (n == "add") {
        GridColumn c; c.col = -1;
        int pos = (int)arg_l(argc, argv, 0, (vb_i32)g->cols.size());
        if (pos < 0 || pos > (int)g->cols.size()) pos = (int)g->cols.size();
        g->cols.insert(g->cols.begin() + pos, c);
        g->rebuild();
        return retobj(g->column(pos));
    }
    if (n == "remove") {
        int k = argc ? g->find_col(argv[0]) : -1;
        if (k < 0) { vb_raise_full(6147, vs(QString("DataGrid")), vs(QString("Column not found"))); return vb_v_empty(); }
        gcol_free(g->cols[k]);
        g->cols.erase(g->cols.begin() + k);
        g->rebuild();
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
static vbvar gcol_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    auto *g = grid_of_obj(o);
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (!g) { vb_raise(91); return vb_v_empty(); }
    int k = g->col_of_obj(vb_oid(o));
    if (k < 0) { vb_raise(91); return vb_v_empty(); }
    GridColumn &c = g->cols[k];
    if (n == "caption") { if (set) { c.caption = data_qs(val); g->rebuild(); return vb_v_empty(); } return vstr(c.caption); }
    if (n == "datafield") {
        if (set) { c.field = data_qs(val); c.col = g->rs ? rs_col_index(g->rs, c.field) : -1; g->rebuild(); return vb_v_empty(); }
        return vstr(c.field);
    }
    if (n == "text" || n.empty() || n == "value") {
        if (g->no_rowset()) return vb_v_empty();
        if (set) { g->col_set(k, val); return vb_v_empty(); }
        return n == "value" ? g->col_value(k) : vstr(g->col_text(k));
    }
    if (n == "colindex") return vb_v_i((vb_i16)k);
    if (n == "width") {
        if (set) { c.him = grid_width(vb_v2d(val)); return vb_v_empty(); }
        return vb_v_f(grid_twips(c.him >= 0 ? c.him : grid_px(101)));
    }
    if (n == "dataformat") {                    /* NumberFormat is its Format */
        if (set) {
            vbobj f = vb_v2obj(val);
            if (vb_errpend) return vb_v_empty();
            if (f && !vb_is_dfmt(f)) { vb_raise(13); return vb_v_empty(); }
            f = f ? vb_oid(f) : nullptr;
            if (f) vb_oaddref(f);
            if (c.fmt) vb_orelease(c.fmt);
            c.fmt = f;
            g->rebuild();
            return vb_v_empty();
        }
        if (!c.fmt) c.fmt = vb_dfmt_new();
        return retobj(c.fmt);
    }
    if (n == "visible") { if (set) { c.visible = vb_v2bool(val) != 0; g->rebuild(); return vb_v_empty(); } return vbool(c.visible); }
    if (n == "locked") { if (set) { c.locked = vb_v2bool(val) != 0; return vb_v_empty(); } return vbool(c.locked); }
    if (n == "alignment") { if (set) { c.align = vb_v2l(val); return vb_v_empty(); } return vb_v_l(c.align); }
    if (n == "numberformat") {
        if (!c.fmt) c.fmt = vb_dfmt_new();
        if (set) { vb_late_o(c.fmt, "format", VBM_LET, 1, &val); g->rebuild(); return vb_v_empty(); }
        return vb_late_o(c.fmt, "format", VBM_GET, 0, nullptr);
    }
    if (n == "celltext" || n == "cellvalue") {
        int row = argc ? (int)vb_v2d(argv[0]) - 1 : g->current_row();
        if (!g->rs || c.col < 0) return n == "celltext" ? vstr(QString()) : vb_v_empty();
        vbvar v = rsx_cell(g->rs, row, c.col);
        return n == "celltext" ? vstr(g->shown(c, v)) : v;
    }
    if (n == "datachanged" || n == "button" || n == "wraptext" || n == "allowsizing") return vbool(false);
    if (n == "dividerstyle") return vb_v_l(2);
    if (n == "left" || n == "top") return vb_v_f(0);
    *found = 0;
    return vb_v_empty();
}
static Ocx *new_datagrid() { return new OcxDataGrid; }

/* ---------------------------------------------------------------- DataList, DataCombo */
struct OcxDataList : Ocx {
    bool combo;
    vbobj rowsrc = nullptr, rowrs = nullptr;    /* RowSource and its Recordset (counted) */
    QString rowmember, listfield, boundcol, text, bound;
    int sel = -1, style = 0;
    bool matched = false, locked = false, filling_list = false;
    QPointer<QListWidget> lw;
    QPointer<QComboBox> cb;
    std::vector<QString> items, keys;
    explicit OcxDataList(bool c) : combo(c) {}
    ~OcxDataList() override { listen(false); if (rowrs) vb_orelease(rowrs); if (rowsrc) vb_orelease(rowsrc); }
    bool data_aware() override { return true; }
    const char *bound_prop() override { return "boundtext"; }
    QString rowsrcname;                         /* RowSource = "Adodc2" in the .frm */
    void source_reset(vbobj s) override
    {
        if (!rowsrc && !rowsrcname.isEmpty() && is_ctl_obj(s) && !rowsrcname.compare(C(s)->name, Qt::CaseInsensitive)) { rowsrc = vb_oid(s); vb_oaddref(rowsrc); }
        if (rowsrc && vb_oid(rowsrc) == vb_oid(s)) bind_rows();
    }
    QWidget *make(QWidget *parent) override
    {
        vbobj self = c->self;
        if (combo) {
            cb = new QComboBox(parent);
            cb->setEditable(style != 2);
            QObject::connect(cb, QOverload<int>::of(&QComboBox::activated), [self](int i) {
                Ctl *x = C(self);
                auto *d = x ? dynamic_cast<OcxDataList *>(x->ocx) : nullptr;
                if (d && !d->filling_list) { d->pick(i); fire_ctl(x, "click"); }
            });
            return cb;
        }
        lw = new QListWidget(parent);
        QObject::connect(lw, &QListWidget::currentRowChanged, [self](int i) {
            Ctl *x = C(self);
            auto *d = x ? dynamic_cast<OcxDataList *>(x->ocx) : nullptr;
            if (d && !d->filling_list && i >= 0 && i != d->sel) { d->pick(i); fire_ctl(x, "click"); }
        });
        QObject::connect(lw, &QListWidget::itemDoubleClicked, [self](QListWidgetItem *) { Ctl *x = C(self); if (x) fire_ctl(x, "dblclick"); });
        return lw;
    }
    void load(const char *const *p) override
    {
        QString v;
        if (prop(p, "ListField", &v)) listfield = v;
        if (prop(p, "RowSource", &v)) rowsrcname = v;
        if (prop(p, "DataField", &v)) c->datafield = v;
        if (prop(p, "DataSource", &v)) c->datasource = v;
        if (prop(p, "BoundColumn", &v)) boundcol = v;
        if (prop(p, "Text", &v)) text = v;
        style = (int)propn(p, "Style", 0);
        locked = propn(p, "Locked", 0) != 0;
        if (cb) { cb->setEditable(style != 2); cb->setEditText(text); }
    }
    static void list_listen(void *ctx, vbobj, const char *ev, int, vbvar *)
    {
        Ctl *x = C((vbobj)ctx);
        auto *d = x ? dynamic_cast<OcxDataList *>(x->ocx) : nullptr;
        if (d && (!strcmp(ev, "lb_fill") || !strcmp(ev, "recordsetchangecomplete"))) d->refill();
    }
    void listen(bool on) { if (rowrs && !rs_is_dao(rowrs) && vb_dbhooks.ado_listen) vb_dbhooks.ado_listen(rowrs, list_listen, c->self, on ? 1 : 0); }
    void bind_rows()
    {
        listen(false);
        if (rowrs) { vb_orelease(rowrs); rowrs = nullptr; }
        vbobj r = source_rs(rowsrc, rowmember);
        if (r) { rowrs = r; vb_oaddref(rowrs); listen(true); }
        refill();
    }
    void refill()
    {
        items.clear(); keys.clear();
        int nc = 0;
        int n = rowrs ? rsx_shape(rowrs, &nc, nullptr) : -1;
        int lc = rowrs ? rs_col_index(rowrs, listfield) : -1;
        int bc = rowrs ? rs_col_index(rowrs, boundcol.isEmpty() ? listfield : boundcol) : -1;
        vb_i32 m = vb_mark();
        for (int r = 0; r < n; r++) {
            items.push_back(lc >= 0 ? cell_text(rsx_cell(rowrs, r, lc)) : QString());
            keys.push_back(bc >= 0 ? cell_text(rsx_cell(rowrs, r, bc)) : QString());
        }
        vb_release(m);
        filling_list = true;
        if (lw) { lw->clear(); for (auto &s : items) lw->addItem(s); }
        if (cb) { QString t = cb->currentText(); cb->clear(); for (auto &s : items) cb->addItem(s); cb->setCurrentIndex(-1); cb->setEditText(text); (void)t; }
        filling_list = false;
        if (sel >= (int)items.size()) sel = -1;
        show_sel();
    }
    void show_sel()
    {
        filling_list = true;
        if (lw) lw->setCurrentRow(sel);
        if (cb) { cb->setCurrentIndex(sel); cb->setEditText(text); }
        filling_list = false;
    }
    void pick(int i)                            /* the user (or the program) chose item i */
    {
        if (i < 0 || i >= (int)items.size()) return;
        sel = i; text = items[i]; bound = keys[i]; matched = true;
        if (!data_filling && !c->datafield.isEmpty()) c->datachanged = true;
        show_sel();
    }
    /* BoundText: an item of the list, or nothing (VB6 oracle f_dfmt5: a value that is not in the list, also the
     * field through a DataFormat "031" for the key 31, leaves BoundText and Text empty) */
    bool set_bound(const QString &v)
    {
        for (int i = 0; i < (int)keys.size(); i++)
            if (keys[i] == v) { sel = i; text = items[i]; bound = v; matched = true; show_sel(); return true; }
        sel = -1; bound = QString(); matched = false;
        text = QString();
        show_sel();
        return false;
    }
    void set_text(const QString &v)
    {
        text = v;
        for (int i = 0; i < (int)items.size(); i++)
            if (!items[i].compare(v, Qt::CaseInsensitive)) { sel = i; bound = keys[i]; matched = true; show_sel(); return; }
        sel = -1; matched = false;
        bound = v;                              /* VB6 oracle f_ado: BoundText is the text typed */
        show_sel();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        if (n == "rowsource") {
            if (set) {
                vbobj o = vb_v2obj(val);
                if (vb_errpend) return vb_v_empty();
                if (o) vb_oaddref(o);
                if (rowsrc) vb_orelease(rowsrc);
                rowsrc = o ? vb_oid(o) : nullptr;
                bind_rows();
                return vb_v_empty();
            }
            return vb_vtmp(vb_v_obj(rowsrc));
        }
        if (n == "rowmember") { if (set) { rowmember = data_qs(val); if (rowsrc) bind_rows(); return vb_v_empty(); } return vstr(rowmember); }
        if (n == "listfield") { if (set) { listfield = data_qs(val); refill(); return vb_v_empty(); } return vstr(listfield); }
        if (n == "boundcolumn") { if (set) { boundcol = data_qs(val); refill(); return vb_v_empty(); } return vstr(boundcol); }
        if (n == "boundtext") {
            if (set) { if (set_bound(data_qs(val)) && !data_filling && !c->datafield.isEmpty()) c->datachanged = true; return vb_v_empty(); }
            return vstr(bound);
        }
        if (n == "text" || n == "deftext" || n.empty()) {
            if (set) { set_text(data_qs(val)); if (!data_filling && !c->datafield.isEmpty()) c->datachanged = true; return vb_v_empty(); }
            if (cb && cb->isEditable()) text = cb->currentText();
            return vstr(text);
        }
        if (n == "matchedwithlist") return vbool(matched);
        if (n == "selecteditem") return sel < 0 ? vb_v_null() : vb_v_d(sel + 1);
        if (n == "visiblecount") return vb_v_l(lw ? std::max(1, lw->viewport()->height() / std::max(1, lw->sizeHintForRow(0))) : 8);
        if (n == "style") { if (set) { style = vb_v2l(val); if (cb) cb->setEditable(style != 2); return vb_v_empty(); } return vb_v_l(style); }
        if (n == "locked") { if (set) { locked = vb_v2bool(val) != 0; return vb_v_empty(); } return vbool(locked); }
        if (n == "refill" || n == "refresh") { if (n == "refill") bind_rows(); else refill(); return vb_v_empty(); }
        if (n == "matchentry" || n == "integralheight") return vb_v_l(0);
        if (n == "datasource" || n == "datafield" || n == "datachanged" || n == "datamember" || n == "dataformat") {
            vbvar r = bound_member(c, n, set, val, found);
            return r;
        }
        *found = 0;
        return vb_v_empty();
    }
};
static Ocx *new_datalist() { return new OcxDataList(false); }
static Ocx *new_datacombo() { return new OcxDataList(true); }

/* ---------------------------------------------------------------- MSFlexGrid / MSHFlexGrid bound */
OcxFlexGrid::~OcxFlexGrid() { if (dsrc) vb_orelease(dsrc); }
/* the data control the Bindings of the .frm name: its name, length-prefixed, in them */
static bool bindings_name(const QByteArray &b, const QString &name)
{
    QByteArray n = name.toLatin1();
    for (int i = 0; i + n.size() < b.size(); i++)
        if ((unsigned char)b[i] == n.size() && !qstrnicmp(b.constData() + i + 1, n.constData(), (uint)n.size())) return true;
    return false;
}
void OcxFlexGrid::source_reset(vbobj s)
{
    if (!dsrc && is_ctl_obj(s) && !bindings.isEmpty() && bindings_name(bindings, C(s)->name)) { dsrc = vb_oid(s); vb_oaddref(dsrc); }
    if (dsrc && vb_oid(dsrc) == vb_oid(s)) bind_fill();
}
vbvar OcxFlexGrid::datasource_member(bool set, vbvar val)
{
    if (!set) return vb_vtmp(vb_v_obj(dsrc));
    vbobj o = vb_v2obj(val);
    if (vb_errpend) return vb_v_empty();
    if (o) vb_oaddref(o);
    if (dsrc) vb_orelease(dsrc);
    dsrc = o ? vb_oid(o) : nullptr;
    bind_fill();
    return vb_v_empty();
}
/* all the records of the source: the names of the fields in the fixed row, a fixed column before them */
void OcxFlexGrid::bind_fill()
{
    if (!tw || !dsrc) return;
    vbobj rs = nullptr;
    bool dao = false;
    if (is_ctl_obj(dsrc) && C(dsrc)->kind == K_DATA) { rs = C(dsrc)->data ? C(dsrc)->data->rs : nullptr; dao = true; }
    else rs = source_rs(dsrc, QString());
    if (!rs) return;
    std::vector<QString> heads;
    std::vector<std::vector<QString>> data;
    vb_i32 m = vb_mark();
    vb_i32 save = vb_errpend;
    if (!dao) {
        int nc = 0, n = rsx_shape(rs, &nc, nullptr);
        for (int k = 0; k < nc; k++) heads.push_back(QString::fromUtf8(rsx_colname(rs, k)));
        for (int r = 0; r < n; r++) {
            std::vector<QString> row;
            for (int k = 0; k < nc; k++) row.push_back(::cell_text(rsx_cell(rs, r, k)));
            data.push_back(row);
        }
    } else {
        vbvar cl = vb_late_o(rs, "clone", VBM_METHOD, 0, nullptr);
        vbobj cr = vb_errpend ? nullptr : vb_v2obj(cl);
        if (cr) {
            vbvar fl = vb_late_o(cr, "fields", VBM_GET, 0, nullptr);
            int nc = vb_v2l(vb_late(fl, "count", VBM_GET, 0, nullptr));
            for (int k = 0; k < nc; k++) { vbvar ix = vb_v_l(k); vbvar f = vb_late(fl, "item", VBM_GET, 1, &ix); heads.push_back(data_qs(vb_late(f, "name", VBM_GET, 0, nullptr))); }
            bool any = !(vb_v2bool(vb_late_o(cr, "eof", VBM_GET, 0, nullptr)) && vb_v2bool(vb_late_o(cr, "bof", VBM_GET, 0, nullptr)));
            if (any) vb_late_o(cr, "movefirst", VBM_METHOD, 0, nullptr);
            while (any && !vb_errpend && !vb_v2bool(vb_late_o(cr, "eof", VBM_GET, 0, nullptr)) && data.size() < 20000) {
                std::vector<QString> row;
                for (int k = 0; k < nc; k++) { vbvar ix = vb_v_l(k); vbvar f = vb_late(fl, "item", VBM_GET, 1, &ix); row.push_back(::cell_text(vb_late(f, "value", VBM_GET, 0, nullptr))); }
                data.push_back(row);
                vb_late_o(cr, "movenext", VBM_METHOD, 0, nullptr);
            }
            vb_late_o(cr, "close", VBM_METHOD, 0, nullptr);
        }
    }
    if (vb_errpend && !save) vb_err_clear_pending();
    vb_release(m);
    /* MSFlexGrid bound: Rows = records + 1, Cols = fields + FixedCols */
    int fc = fcols;
    rows = (int)data.size() + 1; cols = (int)heads.size() + fc;
    if (frows < 1) frows = 1;
    shape();
    for (int k = 0; k < (int)heads.size(); k++) cell(0, k + fc)->setText(heads[k]);
    for (int r = 0; r < (int)data.size(); r++)
        for (int k = 0; k < (int)heads.size(); k++) cell(r + 1, k + fc)->setText(data[r][k]);
    paint_sel();
}

/* DBGrid, DBList, DBCombo over the Data control (phase 7.6) */
#include "ocx_dbgrid.inc"
