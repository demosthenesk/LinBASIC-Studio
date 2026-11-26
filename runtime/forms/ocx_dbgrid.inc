/* ocx_dbgrid.inc - the data-bound controls of VB5 / VB6 over the Data control (DAO), phase 7.6: DBGrid
 * (DBGRID32.OCX, MSDBGrid) and DBList, DBCombo (DBLIST32.OCX, MSDBCtls). They read the rows of the DAO
 * Recordset of a Data control through vb_dbhooks (dao_shape ...), and follow the Data control: it tells
 * them before its recordset moves (the grid saves the row it changed: BeforeUpdate, AfterUpdate) and after
 * (the grid shows the record, RowColChange). The behaviour is that of the VB6 oracle tests/dev/f_dbgrid;
 * where VB6 has a bug, the control does what VB6 meant (notes/phase-7.md). Included by ocx_data.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

/* the Bindings of an .ocx control in the .frx: records [u32 0x0001FAC6][u32 dispid][u8 n][name][zeros]: the
 * name that a dispid binds (DBGrid DataSource 9, DBList RowSource 3, DBCombo RowSource 16) */
static QString bindings_for(const QByteArray &b, int dispid)
{
    for (int i = 0; i + 9 <= b.size(); i++) {
        const uchar *p = (const uchar *)b.constData() + i;
        if (p[0] != 0xc6 || p[1] != 0xfa || p[2] != 0x01 || p[3] != 0) continue;
        int id = (int)(p[4] | p[5] << 8 | p[6] << 16 | (unsigned)p[7] << 24);
        int n = p[8];
        if (i + 9 + n > b.size()) break;
        if (id == dispid) return QString::fromLatin1(b.constData() + i + 9, n);
    }
    return QString();
}

/* bookmarks: the DAO Recordset has Byte arrays, the DBGrid gives Strings of the same bytes (VB6 oracle
 * f_dbgrid: TypeName(DBGrid1.Bookmark) is String) and takes both */
static QByteArray dbg_bytes(vbvar v)
{
    v = vb_v_letval(v);
    vb_i32 save = vb_errpend;
    if (v.vt == VT_BSTR) {
        QByteArray b;
        vb_i32 n = vb_sbytes(v.u.s);
        for (vb_i32 i = 0; i < n; i++) b.append((char)vb_sbyte(v.u.s, i));
        return b;
    }
    if (v.vt & VT_ARRAY) {
        vbarr *a = vb_v_arrptr(v);
        if (vb_errpend && !save) { vb_err_clear_pending(); return QByteArray(); }
        if (a && a->data) return QByteArray(a->data, (int)(a->n * a->esz));
    }
    return QByteArray();
}
static vbvar dbg_bmvar(const QByteArray &b)     /* as the DAO Recordset has it */
{
    if (b.isEmpty()) return vb_v_null();
    vbarr a; memset(&a, 0, sizeof a);
    vb_adim(&a, 1, VBA_PLAIN, 0, 1, 0, (vb_i32)b.size() - 1);
    memcpy(a.data, b.constData(), (size_t)b.size());
    vbvar r = vb_v_arr(&a, VT_UI1);
    vb_afree(&a);
    return r;
}
static vbvar dbg_bmstr(const QByteArray &b)     /* as the grid gives it */
{
    if (b.isEmpty()) return vb_v_null();
    return vb_v_s(vb_sfrombytes(b.constData(), (vb_i32)b.size()));
}
static QByteArray dao_row_bm(vbobj rs, int row)
{
    if (!rs_is_dao(rs) || row < 0) return QByteArray();
    vb_i32 m = vb_mark();
    QByteArray b = dbg_bytes(vb_dbhooks.dao_bookmark(rs, row));
    vb_release(m);
    return b;
}
static int dao_bm_row(vbobj rs, const QByteArray &b)
{
    if (!rs_is_dao(rs) || b.isEmpty()) return -1;
    vb_i32 m = vb_mark();
    int r = vb_dbhooks.dao_bmrow(rs, dbg_bmvar(b));
    vb_release(m);
    return r;
}

/* the width of a column: VB6 keeps it in HIMETRIC of whole pixels (VB6 oracle f_dbgrid: Width = 2000 gives
 * 1995.024: 133 pixels, 3519 HIMETRIC) */
static int dbg_tw2him(double tw) { long px = std::lround(tw / TPP); return (int)std::lround(px * 2540.0 / 96.0); }
static double dbg_him2tw(int him) { return (float)(him * 1440.0 / 2540.0); }
static int dbg_him2px(int him) { return (int)std::lround(him * 96.0 / 2540.0); }

static bool dao_numeric(int t) { return (t >= 2 && t <= 8) || t == 16 || (t >= 19 && t <= 23); }   /* Byte .. Date, BigInt, Numeric .. */

/* a value as the grid shows it */
static QString dbg_text(vbvar v, const QString &fmt)
{
    if (v.vt == VT_NULL || v.vt == VT_EMPTY) return QString();
    vb_i32 save = vb_errpend;
    vb_i32 m = vb_mark();
    QString s = fmt.isEmpty() ? qs(vb_v2s(v)) : qs(vb_format(v, vs(fmt), 1, 1, 1));
    vb_release(m);
    if (vb_errpend && !save) { vb_err_clear_pending(); return QString(); }
    return s;
}

struct DbgKeys : QObject {
    std::function<bool(QKeyEvent *)> fn;
    bool eventFilter(QObject *o, QEvent *e) override
    {
        if (e->type() == QEvent::KeyPress && fn && fn(static_cast<QKeyEvent *>(e))) return true;
        return QObject::eventFilter(o, e);
    }
};

/* ---------------------------------------------------------------- DBGrid */
struct DbgCol {
    QString caption, field, numfmt, defval;
    int fcol = -1;                              /* the field of the recordset */
    int him = -1;                               /* Width in HIMETRIC; -1: as the grid makes it */
    int align = 3, divider = 2;
    bool visible = true, locked = false, allowsizing = true, wraptext = false, button = false;
    vbobj obj = nullptr;                        /* the Column object, when it is asked for */
};
static vbvar dbcol_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vbvar dbcols_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vbvar dbsplits_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vbvar dbsplit_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vbvar dbselbm_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vb_i32 dbcols_count(vbobj o);
static vbvar dbcols_item(vbobj o, vb_i32 i);
static vb_i32 dbselbm_count(vbobj o);
static vbvar dbselbm_item(vbobj o, vb_i32 i);
static vb_i32 dbsplits_count(vbobj) { return 1; }
static vbvar dbsplits_item(vbobj o, vb_i32 i);
static const vb_class CLS_DBCOL = { "Column", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "value", nullptr, nullptr, nullptr, dbcol_fallback, nullptr };
static const vb_class CLS_DBCOLS = { "Columns", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "item", dbcols_count, dbcols_item, nullptr, dbcols_fallback, nullptr };
static const vb_class CLS_DBSPLITS = { "Splits", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "item", dbsplits_count, dbsplits_item, nullptr, dbsplits_fallback, nullptr };
static const vb_class CLS_DBSPLIT = { "Split", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr, nullptr, nullptr, nullptr, dbsplit_fallback, nullptr };
static const vb_class CLS_DBSELBM = { "SelBookmarks", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "item", dbselbm_count, dbselbm_item, nullptr, dbselbm_fallback, nullptr };

/* the RowBuffer of the Unbound events (DataMode 1): rows x columns of values and a bookmark for each row */
struct RowBuf {
    int rows = 0, cols = 0;
    std::vector<vbvar> val, bm;                 /* owned (vb_vset) */
    std::vector<QString> names;
    RowBuf(int r, int c) : rows(r), cols(c), val((size_t)(r * c)), bm((size_t)r) {}
    ~RowBuf() { for (auto &v : val) vb_vfree(&v); for (auto &v : bm) vb_vfree(&v); }
};
static vbvar rowbuf_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    RowBuf *b = (RowBuf *)vb_oid(o)->ext;
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    int na = set ? argc - 1 : argc;
    *found = 1;
    if (!b) { vb_raise(91); return vb_v_empty(); }   /* the event has ended */
    if (n == "rowcount") {
        if (!set) return vb_v_l(b->rows);
        int v = (int)vb_v2l(val);
        if (v < 0 || v > (int)b->bm.size()) { vb_raise(380); return vb_v_empty(); }
        b->rows = v;
        return vb_v_empty();
    }
    if (n == "columncount") return vb_v_i((vb_i16)b->cols);
    if (n == "columnname") {
        int k = na > 0 ? (int)vb_v2l(argv[0]) : -1;
        if (k < 0 || k >= b->cols) { vb_raise(9); return vb_v_empty(); }
        return vstr(b->names[k]);
    }
    if (n == "value" || n.empty()) {
        int r = na > 0 ? (int)vb_v2l(argv[0]) : -1, k = na > 1 ? (int)vb_v2l(argv[1]) : -1;
        if (r < 0 || r >= (int)b->bm.size() || k < 0 || k >= b->cols) { vb_raise(9); return vb_v_empty(); }
        vbvar &cell = b->val[(size_t)(r * b->cols + k)];
        if (!set) return vb_vtmp(cell);
        vb_vset(&cell, vb_v_letval(val));
        return vb_v_empty();
    }
    if (n == "bookmark") {
        int r = na > 0 ? (int)vb_v2l(argv[0]) : -1;
        if (r < 0 || r >= (int)b->bm.size()) { vb_raise(9); return vb_v_empty(); }
        if (!set) return vb_vtmp(b->bm[r]);
        vb_vset(&b->bm[r], vb_v_letval(val));
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
static const vb_class CLS_ROWBUF = { "RowBuffer", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, "value", nullptr, nullptr, nullptr, rowbuf_fallback, nullptr };
/* a bookmark of an unbound grid as a key (Long 3 and Integer 3 are the same row) */
static QString ukey(vbvar v)
{
    v = vb_v_letval(v);
    if (v.vt == VT_NULL || v.vt == VT_EMPTY) return QString();
    vb_i32 save = vb_errpend;
    vb_i32 m = vb_mark();
    QString s = (v.vt & VT_ARRAY) ? QString::fromLatin1(dbg_bytes(v).toHex()) : qs(vb_v2s(v));
    vb_release(m);
    if (vb_errpend && !save) vb_err_clear_pending();
    return s;
}

static void dbg_raise(vb_i32 n)
{
    const char *t = n == 6145 ? "Invalid column index" : n == 6148 ? "Invalid row number" : n == 6149 ? "Invalid bookmark" :
                    n == 6150 ? "Invalid SelBookmarks index" : n == 6151 ? "Scroll out of range" :
                    n == 6153 ? "Update failed" : n == 6154 ? "Add failed" : n == 6155 ? "Delete failed" :
                    n == 6160 ? "Data access error" : n == 6163 ? "No current record" : n == 6244 ? "Invalid split index" :
                    "Invalid property value";
    vb_i32 m = vb_mark();
    vb_raise_full(n, vs(QString("DBGrid")), vs(QString(t)));
    vb_release(m);
}

struct OcxDBGrid : Ocx {
    vbobj src = nullptr;                        /* the Data control (counted) */
    QString srcname;                            /* its name, from the Bindings of the .frx */
    std::vector<DbgCol> cols, design;           /* the columns; those of the .frx */
    bool held = false;                          /* the columns stay at ReBind: _LayoutType 4 of the .frx, HoldFields */
    /* DataMode 1 (unbound): the rows the program gave through UnboundReadData, its bookmarks (Variants) */
    struct URow { vbvar bm{}; std::vector<vbvar> v; };
    std::vector<URow> urows;
    bool uread_due = false;
    VFont headvf;                               /* HeadFont */
    vbobj headfont = nullptr;
    bool bound_once = false;
    QPointer<QFrame> box;
    QPointer<QLabel> cap;
    QPointer<QTableWidget> tw;
    DbgKeys *keys = nullptr;
    /* properties */
    bool allowadd = false, allowdelete = false, allowupdate = true, allowrowsizing = true, allowarrows = true;
    bool colheaders = true, recsel = true, marqueeunique = true, tabacross = false, wrapcell = false;
    int datamode = 0, marquee = 6, scrollbars = 4, rowdivider = 2, appearance = 1, borderstyle = 1, tabaction = 0;
    int scrollgroup = 1, sizemode = 0, leftcol = 0;
    bool split_locked = false, split_allowfocus = true, split_allowsizing = false;
    double headlines = 1, defcolwidth = 0, rowheight = 0, split_size = 1;
    long back = 0x80000005L, fore = 0x80000008L;
    QString caption, errortext;
    /* the state */
    int nrows = 0, currow = -1, top = 0, curcol = 0;
    bool adding = false;                        /* the recordset is in AddNew (the grid shows the new row) */
    bool add_bu = false, add_done = false;      /* its BeforeUpdate came; its Update was answered */
    bool rcc_pending = false;                   /* the Data control has another recordset: RowColChange after Reposition */
    int rcc_lastcol = 0, want_col = -1;         /* want_col: the user clicked a cell of another row */
    QByteArray lastbm;                          /* the row of the last RowColChange */
    std::map<int, QString> pend;                /* the changed cells of the current row, not saved yet */
    bool pend_add = false, on_addrow = false;   /* the user writes a new row (AllowAddNew) */
    bool building = false;
    std::vector<QByteArray> selbm;
    vbobj colsobj = nullptr, splitsobj = nullptr, splitobj = nullptr, selbmobj = nullptr;

    ~OcxDBGrid() override
    {
        for (auto &g : cols) drop(g);
        for (vbobj *o : { &colsobj, &splitsobj, &splitobj, &selbmobj }) if (*o) { (*o)->ext = nullptr; vb_orelease(*o); }
        if (src) vb_orelease(src);
        if (headfont) vb_orelease(headfont);
        uclear();
        delete keys;
    }
    bool unbound() { return datamode == 1; }
    void uclear() { for (auto &r : urows) { vb_vfree(&r.bm); for (auto &v : r.v) vb_vfree(&v); } urows.clear(); }
    int urow_of(vbvar bm) { QString k = ukey(bm); if (k.isEmpty()) return -1; for (int i = 0; i < (int)urows.size(); i++) if (ukey(urows[i].bm) == k) return i; return -1; }
    /* a bookmark as the program sees it: a String (bound: the bytes of the DAO bookmark), the program's own (unbound) */
    vbvar bm_var(int row)
    {
        if (unbound()) return row >= 0 && row < (int)urows.size() ? vb_vtmp(urows[row].bm) : vb_v_null();
        return row >= 0 && row < nrows ? dbg_bmstr(bm_of(row)) : vb_v_null();
    }
    /* UnboundReadData until the program has no more rows: all of them, in chunks (VB6 oracle f_dbgunb) */
    void uread()
    {
        uread_due = false;
        if (!c->form) return;
        QByteArray keep = ukey(bm_var(currow)).toUtf8();
        uclear();
        int nc = (int)cols.size(), chunk = std::max(10, visible_rows() + 2);
        vbvar start = vb_v_null();
        QString startkey;
        for (int guard = 0; guard < 100000 && nc >= 0; guard += chunk) {
            RowBuf *rb = new RowBuf(chunk, nc);
            for (auto &g : cols) rb->names.push_back(g.field.isEmpty() ? g.caption : g.field);
            vbobj o = vb_onew(&CLS_ROWBUF);
            o->ext = rb;
            vb_oaddref(o);
            vb_i32 m = vb_mark();
            vbvar a[3] = { vb_v_obj(o), vb_vtmp(start), vb_v_bool(0) };
            fire_ctl(c, "unboundreaddata", 3, a);
            vb_release(m);
            o->ext = nullptr;
            vb_orelease(o);
            if (vb_errpend) { delete rb; break; }
            int got = std::min(rb->rows, chunk);
            for (int i = 0; i < got; i++) {
                URow r;
                vb_vset(&r.bm, rb->bm[i]);
                r.v.resize((size_t)nc);
                for (int k = 0; k < nc; k++) vb_vset(&r.v[k], rb->val[(size_t)(i * nc + k)]);
                urows.push_back(r);
            }
            QString lastkey = got ? ukey(rb->bm[got - 1]) : QString();
            if (got) { vb_vfree(&start); vb_vset(&start, rb->bm[got - 1]); }
            delete rb;
            if (got < chunk || lastkey.isEmpty() || lastkey == startkey) break;
            startkey = lastkey;
        }
        vb_vfree(&start);
        nrows = (int)urows.size();
        int r = -1;
        if (!keep.isEmpty()) for (int i = 0; i < nrows; i++) if (ukey(urows[i].bm).toUtf8() == keep) { r = i; break; }
        currow = r >= 0 ? r : nrows ? std::min(std::max(currow, 0), nrows - 1) : -1;
    }
    void fire_rcc_v(vbvar last, int lastcol)
    {
        vb_i32 m = vb_mark();
        vbvar a[2] = { vb_vtmp(last), vb_v_i((vb_i16)lastcol) };
        fire_ctl(c, "rowcolchange", 2, a);
        vb_release(m);
    }
    /* the unbound grid goes to a row: the changed row saved first (BeforeUpdate, UnboundWriteData, AfterUpdate) */
    bool umove(int r, bool center = false)
    {
        if (r < 0 || r >= nrows) return false;
        if ((!pend.empty() || pend_add) && !commit()) return false;
        vb_i32 m = vb_mark();
        vbvar last = vb_vtmp(bm_var(currow));
        currow = r;
        on_addrow = false;
        int vis = std::max(1, visible_rows());
        if (center) top = std::max(0, r - vis / 2);   /* Bookmark = : the row in the middle (VB6 oracle f_dbgunb) */
        else if (r < top) top = r;
        else if (r >= top + vis) top = r - vis + 1;
        paint();
        fire_rcc_v(last, curcol);
        vb_release(m);
        return true;
    }
    static void drop(DbgCol &g) { if (g.obj) { g.obj->ext = nullptr; vb_orelease(g.obj); g.obj = nullptr; } }
    bool data_aware() override { return true; }
    bool own_binding() override { return true; }
    bool focusable() override { return true; }

    vbobj rs()
    {
        Ctl *d = src ? C(src) : nullptr;
        return d && d->kind == K_DATA && d->data ? d->data->rs : nullptr;
    }
    bool mine(vbobj dc)
    {
        if (!src && !srcname.isEmpty() && is_ctl_obj(dc) && !srcname.compare(C(dc)->name, Qt::CaseInsensitive)) { src = vb_oid(dc); vb_oaddref(src); }
        return src && vb_oid(src) == vb_oid(dc);
    }

    /* ------------------------------------------------ the widget */
    QWidget *make(QWidget *parent) override
    {
        box = new QFrame(parent);
        box->setFrameShape(QFrame::Box);
        auto *lay = new QVBoxLayout(box);
        lay->setContentsMargins(1, 1, 1, 1);
        lay->setSpacing(0);
        cap = new QLabel(box);
        cap->setAlignment(Qt::AlignCenter);
        cap->setFrameShape(QFrame::Panel);
        cap->setVisible(false);
        tw = new QTableWidget(box);
        tw->setFrameShape(QFrame::NoFrame);
        tw->setSelectionMode(QAbstractItemView::SingleSelection);
        tw->setVerticalScrollMode(QAbstractItemView::ScrollPerItem);
        tw->verticalHeader()->setSectionsClickable(true);
        lay->addWidget(cap);
        lay->addWidget(tw, 1);
        vbobj self = c->self;
        auto me = [self]() -> OcxDBGrid * { Ctl *x = C(self); return x ? dynamic_cast<OcxDBGrid *>(x->ocx) : nullptr; };
        QObject::connect(tw, &QTableWidget::currentCellChanged, [me](int r, int col, int, int) {
            auto *g = me();
            if (g && !g->building && r >= 0) g->user_move(r, col);
        });
        QObject::connect(tw, &QTableWidget::itemChanged, [me](QTableWidgetItem *it) {
            auto *g = me();
            if (g && !g->building && it) g->user_edit(it->row(), it->column(), it->text());
        });
        QObject::connect(tw, &QTableWidget::cellClicked, [self](int, int) { Ctl *x = C(self); if (x) fire_ctl(x, "click"); });
        QObject::connect(tw, &QTableWidget::cellDoubleClicked, [self](int, int) { Ctl *x = C(self); if (x) fire_ctl(x, "dblclick"); });
        QObject::connect(tw->horizontalHeader(), &QHeaderView::sectionClicked, [self](int i) {
            Ctl *x = C(self);
            if (!x) return;
            vbvar a = vb_v_i((vb_i16)i);
            fire_ctl(x, "headclick", 1, &a);
        });
        QObject::connect(tw->verticalHeader(), &QHeaderView::sectionClicked, [me](int r) {
            auto *g = me();
            if (!g || r < 0 || r >= g->nrows) return;
            vbvar a = vb_v_i(0);
            Ctl *x = g->c;
            fire_ctl(x, "selchange", 1, &a);
            if ((g = me()) && !vb_v2i(a)) { g->selbm.clear(); g->selbm.push_back(dao_row_bm(g->rs(), r)); }
        });
        QObject::connect(tw->horizontalHeader(), &QHeaderView::sectionResized, [me](int k, int, int w) {
            auto *g = me();
            if (g && !g->building && k >= 0 && k < (int)g->cols.size()) g->cols[k].him = (int)std::lround(w * 2540.0 / 96.0);
        });
        keys = new DbgKeys;
        keys->fn = [me](QKeyEvent *e) -> bool {
            auto *g = me();
            if (!g || e->key() != Qt::Key_Delete || !g->allowdelete || !g->tw) return false;
            if (g->editing()) return false;
            if (g->currow < 0 || g->currow >= g->nrows || !g->tw->selectionModel()->isRowSelected(g->currow, QModelIndex())) return false;
            g->delete_row();
            return true;
        };
        tw->installEventFilter(keys);
        look();
        return box;
    }
    void look()
    {
        if (!tw) return;
        if (box) box->setFrameShape(borderstyle ? QFrame::Box : QFrame::NoFrame);
        if (cap) { cap->setText(caption); cap->setVisible(!caption.isEmpty()); }
        tw->verticalHeader()->setVisible(recsel);
        tw->horizontalHeader()->setVisible(colheaders);
        tw->verticalHeader()->setSectionResizeMode(allowrowsizing ? QHeaderView::Interactive : QHeaderView::Fixed);
        tw->setShowGrid(rowdivider != 0);
        QAbstractItemView::EditTriggers et(QAbstractItemView::NoEditTriggers);
        if (allowupdate || allowadd) et = QAbstractItemView::DoubleClicked | QAbstractItemView::EditKeyPressed | QAbstractItemView::AnyKeyPressed;
        tw->setEditTriggers(et);
        tw->setHorizontalScrollBarPolicy(scrollbars == 0 || scrollbars == 2 ? Qt::ScrollBarAlwaysOff : scrollbars == 4 ? Qt::ScrollBarAsNeeded : Qt::ScrollBarAlwaysOn);
        tw->setVerticalScrollBarPolicy(scrollbars == 0 || scrollbars == 1 ? Qt::ScrollBarAlwaysOff : scrollbars == 4 ? Qt::ScrollBarAsNeeded : Qt::ScrollBarAlwaysOn);
        tw->setSelectionBehavior(marquee == 3 || marquee == 4 ? QAbstractItemView::SelectRows : QAbstractItemView::SelectItems);
    }
    bool editing()                              /* a cell editor is open (the state of the view is protected) */
    {
        QWidget *fw = QApplication::focusWidget();
        return tw && fw && fw != tw && tw->isAncestorOf(fw);
    }
    int row_px()                                /* the height of a row */
    {
        if (rowheight > 0) return std::max(8, (int)std::lround(rowheight / TPP));
        QFont f = tw ? tw->font() : qfont(c->font);
        return QFontMetrics(f).height() + 1;
    }
    int visible_rows()                          /* the rows the height of the grid shows, whole or in part */
    {
        int rh = row_px();
        int avail = (int)(c->height / TPP) - 4;
        if (!caption.isEmpty()) avail -= rh;
        if (colheaders) avail -= (int)std::lround(rh * std::max(1.0, headlines));
        return std::max(0, (avail + rh - 1) / rh);
    }

    /* ------------------------------------------------ the .frm, the .frx */
    void load(const char *const *p) override
    {
        QString v;
        if (prop(p, "Bindings", &v)) srcname = bindings_for(QByteArray::fromHex(v.toLatin1()), 9);
        if (prop(p, "DataSource", &v) && !v.isEmpty()) srcname = v;
        if (prop(p, "Caption", &v)) caption = v;
        allowadd = propn(p, "AllowAddNew", 0) != 0;
        allowdelete = propn(p, "AllowDelete", 0) != 0;
        allowupdate = propn(p, "AllowUpdate", -1) != 0;
        allowarrows = propn(p, "AllowArrows", -1) != 0;
        colheaders = propn(p, "ColumnHeaders", -1) != 0;
        marqueeunique = propn(p, "MarqueeUnique", -1) != 0;
        tabacross = propn(p, "TabAcrossSplits", 0) != 0;
        wrapcell = propn(p, "WrapCellPointer", 0) != 0;
        datamode = (int)propn(p, "DataMode", 0);
        rowdivider = (int)propn(p, "RowDividerStyle", 2);
        appearance = (int)propn(p, "Appearance", 1);
        borderstyle = (int)propn(p, "BorderStyle", 1);
        tabaction = (int)propn(p, "TabAction", 0);
        headlines = propn(p, "HeadLines", 1);
        defcolwidth = propn(p, "DefColWidth", 0);
        {                                       /* _RowHeight of the blob: the pixels less 2 (VB6 oracle DEV-11 s_rowh:
                                                   RowHeight 300 -> 18), 0x80000001 none */
            double rh = propn(p, "_RowHeight", -2147483647.0);
            if (rh > -1000000) rowheight = (rh + 2) * TPP;
        }
        back = propc(p, "BackColor", 0x80000005L);
        fore = propc(p, "ForeColor", 0x80000008L);
        marquee = (int)propn(p, "Split0.MarqueeStyle", 6);
        recsel = propn(p, "Split0.RecordSelectors", -1) != 0;
        scrollbars = (int)propn(p, "Split0.ScrollBars", 4);
        allowrowsizing = propn(p, "Split0.AllowRowSizing", -1) != 0;
        split_locked = propn(p, "Split0.Locked", 0) != 0;
        split_allowfocus = propn(p, "Split0.AllowFocus", -1) != 0;
        split_allowsizing = propn(p, "Split0.AllowSizing", 0) != 0;
        scrollgroup = (int)propn(p, "Split0.ScrollGroup", 1);
        sizemode = (int)propn(p, "Split0.SizeMode", 0);
        split_size = propn(p, "Split0.Size", 1);
        int n = (int)propn(p, "Columns.Count", 2);
        bool has_layout = prop(p, "_LayoutType", nullptr);
        if (has_layout) held = propn(p, "_LayoutType", 0) != 0;   /* 4: columns made at design time (PSC-VB6, VB6 oracle f_dbgunb) */
        headvf = c->font;
        if (prop(p, "HeadFont.Name", &v)) headvf.name = v;
        headvf.size = real_size(propn(p, "HeadFont.Size", headvf.size));
        headvf.charset = (int)propn(p, "HeadFont.Charset", headvf.charset);
        headvf.weight = (int)propn(p, "HeadFont.Weight", headvf.weight);
        headvf.bold = headvf.weight >= 600;
        headvf.italic = propn(p, "HeadFont.Italic", 0) != 0;
        headvf.underline = propn(p, "HeadFont.Underline", 0) != 0;
        headvf.strike = propn(p, "HeadFont.Strikethrough", 0) != 0;
        if (tw) tw->horizontalHeader()->setFont(qfont(headvf));
        design.clear();
        for (int i = 0; i < n; i++) {
            DbgCol g;
            QString b = QString("Columns(%1).").arg(i);
            if (prop(p, (b + "Caption").toUtf8().constData(), &v)) g.caption = v;
            if (prop(p, (b + "DataField").toUtf8().constData(), &v)) g.field = v;
            if (prop(p, (b + "NumberFormat").toUtf8().constData(), &v)) g.numfmt = v;
            if (prop(p, (b + "DefaultValue").toUtf8().constData(), &v)) g.defval = v;
            g.him = (int)propn(p, (b + "_Width").toUtf8().constData(), -1);
            g.align = (int)propn(p, (b + "Alignment").toUtf8().constData(), 3);
            if (!has_layout && !g.field.isEmpty()) held = true;   /* (a .frm of the IDE without a blob) */
            design.push_back(g);
        }
        for (auto &g : cols) drop(g);
        cols = design;
        look();
        paint();
        if (unbound() && !uread_due) {          /* the rows when the form shows (after Form_Load) */
            uread_due = true;
            vbobj self = c->self;
            QTimer::singleShot(0, [self]() {
                Ctl *x = C(self);
                auto *g = x ? dynamic_cast<OcxDBGrid *>(x->ocx) : nullptr;
                if (g && g->uread_due && g->unbound()) { g->uread(); g->reload(); }
            });
        }
    }

    /* ------------------------------------------------ the rows */
    void auto_columns()                         /* one column for each field (the default layout) */
    {
        vbobj r = rs();
        int nc = 0;
        if (!r || rsx_shape(r, &nc, nullptr) < 0 || nc == 0) return;
        for (auto &g : cols) drop(g);
        cols.clear();
        for (int i = 0; i < nc; i++) {
            DbgCol g;
            g.caption = g.field = QString::fromUtf8(rsx_colname(r, i));
            g.fcol = i;
            g.align = dao_numeric(rsx_coltype(r, i)) ? 1 : 0;   /* VB6 oracle f_dbgrid: numbers right, text left */
            if (defcolwidth > 0) g.him = dbg_tw2him(defcolwidth);
            cols.push_back(g);
        }
    }
    QByteArray bm_of(int row) { return row >= 0 && row < nrows ? dao_row_bm(rs(), row) : QByteArray(); }
    QByteArray cur_bm() { return bm_of(currow); }
    QString cell(int row, int k)                /* row: nrows is the new row (AddNew) */
    {
        if (k < 0 || k >= (int)cols.size()) return QString();
        if (row == (on_addrow ? nrows : currow) && pend.count(k)) return pend[k];
        if (unbound()) {
            if (row < 0 || row >= (int)urows.size() || k >= (int)urows[row].v.size()) return QString();
            return dbg_text(urows[row].v[k], cols[k].numfmt);
        }
        vbobj r = rs();
        if (!r || cols[k].fcol < 0 || row < 0) return QString();
        if (row >= nrows && !adding) return QString();
        vb_i32 m = vb_mark();
        QString s = dbg_text(rsx_cell(r, row >= nrows ? -1 : row, cols[k].fcol), cols[k].numfmt);
        vb_release(m);
        return s;
    }
    vbvar value(int row, int k)                 /* Value: the text (VB6 oracle f_dbgrid); Empty without a field */
    {
        if (k < 0 || k >= (int)cols.size()) return vb_v_empty();
        if (row == (on_addrow ? nrows : currow) && pend.count(k)) return vstr(pend[k]);
        if (unbound())                          /* the program's own values (VB6 oracle f_dbgunb: Long '10') */
            return row >= 0 && row < (int)urows.size() && k < (int)urows[row].v.size() ? vb_vtmp(urows[row].v[k]) : vb_v_empty();
        if (cols[k].fcol < 0 || row < 0 || !rs()) return vb_v_empty();
        return vstr(cell(row, k));
    }
    void reload()
    {
        if (unbound()) {
            nrows = (int)urows.size();
            adding = false;
            if (currow >= nrows) currow = nrows - 1;
            if (currow < 0 && nrows) currow = 0;
            int vis = std::max(1, visible_rows());
            if (currow >= 0) { if (currow < top) top = currow; else if (currow >= top + vis) top = currow - vis + 1; }
            if (top > std::max(0, nrows - 1)) top = std::max(0, nrows - 1);
            if (curcol >= (int)cols.size()) curcol = std::max(0, (int)cols.size() - 1);
            paint();
            return;
        }
        vbobj r = rs();
        int nc = 0, cur = -1;
        int n = r ? rsx_shape(r, &nc, &cur) : -1;
        nrows = n < 0 ? 0 : n;
        adding = n >= 0 && cur == n;
        currow = cur;
        for (auto &g : cols) g.fcol = g.field.isEmpty() || !r ? -1 : rs_col_index(r, g.field);
        int vis = std::max(1, visible_rows());
        int last = adding ? nrows : nrows - 1;
        if (currow >= 0 && currow <= last) {
            if (currow < top) top = currow;
            else if (currow >= top + vis) top = currow - vis + 1;
        }
        if (top > std::max(0, nrows - 1)) top = std::max(0, nrows - 1);
        if (top < 0) top = 0;
        if (curcol >= (int)cols.size()) curcol = std::max(0, (int)cols.size() - 1);
        paint();
    }
    void paint()
    {
        if (!tw) return;
        building = true;
        int nc = (int)cols.size();
        int rows = std::min(nrows, 20000) + (adding || allowadd ? 1 : 0);
        tw->clear();
        tw->setColumnCount(nc);
        tw->setRowCount(rows);
        QStringList heads;
        for (auto &g : cols) heads << g.caption;
        tw->setHorizontalHeaderLabels(heads);
        tw->verticalHeader()->setDefaultSectionSize(row_px());
        QFontMetrics fm(tw->font());
        for (int k = 0; k < nc; k++) {
            DbgCol &g = cols[k];
            bool num = g.align == 1 || (g.align == 3 && general_num(k));
            Qt::Alignment al = Qt::AlignVCenter | (g.align == 2 ? Qt::AlignHCenter : num ? Qt::AlignRight : Qt::AlignLeft);
            int wmax = fm.horizontalAdvance(g.caption);
            for (int r = 0; r < rows; r++) {
                QString t = cell(r, k);
                auto *it = new QTableWidgetItem(t);
                it->setTextAlignment(al);
                Qt::ItemFlags f = Qt::ItemIsSelectable | Qt::ItemIsEnabled;
                bool newrow = r >= nrows;
                if (!g.locked && !split_locked && (newrow ? allowadd || adding : allowupdate)) f |= Qt::ItemIsEditable;
                it->setFlags(f);
                tw->setItem(r, k, it);
                if (r < 40) wmax = std::max(wmax, fm.horizontalAdvance(t));
            }
            if (g.him < 0) g.him = (int)std::lround(std::max(48, wmax + 12) * 2540.0 / 96.0);
            tw->setColumnWidth(k, dbg_him2px(g.him));
            tw->setColumnHidden(k, !g.visible);
        }
        if (on_addrow && rows > nrows) tw->setCurrentCell(nrows, std::min(curcol, std::max(0, nc - 1)));
        else if (currow >= 0 && currow < rows && nc) tw->setCurrentCell(currow, std::min(curcol, nc - 1));
        tw->verticalScrollBar()->setValue(top);
        building = false;
    }
    bool general_num(int k)                     /* a column of numbers (General: right; VB6 oracle f_dbgrid) */
    {
        DbgCol &g = cols[k];
        if (unbound()) return false;            /* no field types: left (VB6 oracle f_dbgunb: Alignment 0) */
        return g.fcol >= 0 && rs() && dao_numeric(rsx_coltype(rs(), g.fcol));
    }
    void sync_cell()
    {
        if (!tw) return;
        building = true;
        int r = on_addrow ? nrows : currow;
        if (r >= 0 && r < tw->rowCount() && curcol < tw->columnCount()) tw->setCurrentCell(r, curcol);
        building = false;
    }
    void fire_rcc(const QByteArray &last, int lastcol)
    {
        vbvar a[2] = { dbg_bmstr(last), vb_v_i((vb_i16)lastcol) };
        vb_i32 m = vb_mark();
        fire_ctl(c, "rowcolchange", 2, a);
        vb_release(m);
    }
    void error_event(vb_i32 num, const QString &msg)
    {
        errortext = msg;
        vbvar a[2] = { vb_v_i((vb_i16)num), vb_v_i(1) };
        vb_i32 m = vb_mark();
        fire_ctl(c, "error", 2, a);
        int resp = vb_v2i(a[1]);
        vb_release(m);
        if (resp) vb_msgbox(vb_sfromutf8(msg.toUtf8().constData()), 48, vb_app_exename());
    }

    /* ------------------------------------------------ the Data control */
    void source_reset(vbobj s) override         /* its recordset is another (Refresh): the columns again */
    {
        if (!mine(s)) return;
        int old = curcol;
        bool was = bound_once;
        bound_once = true;
        pend.clear(); pend_add = false; on_addrow = false; adding = false; add_bu = false; add_done = false;
        if (!held) auto_columns();
        curcol = 0;
        reload();
        if (was) { rcc_pending = true; rcc_lastcol = old; }   /* VB6 oracle f_dbgrid: Data1.Refresh: Reposition, RowColChange */
        else lastbm = cur_bm();
    }
    void data_before(vbobj dc, int action, int *cancel) override
    {
        if (!mine(dc) || action == 0) return;
        if (action == 6 && adding) {            /* the AddNew of the recordset is saved: the grid's BeforeUpdate first */
            vbvar a = vb_v_i(0);
            fire_ctl(c, "beforeupdate", 1, &a);
            if (vb_v2i(a)) { *cancel = 1; return; }
            add_bu = true;
            return;
        }
        if ((!pend.empty() || pend_add) && !commit()) *cancel = 1;   /* the row the user changed is saved before it moves */
    }
    void data_after(vbobj dc, int phase) override
    {
        if (!mine(dc)) return;
        if (phase == 3) { reload(); lastbm = cur_bm(); return; }
        vbobj r = rs();
        int nc = 0, cur = -1;
        int n = r ? rsx_shape(r, &nc, &cur) : -1;
        bool now_adding = n >= 0 && cur == n;
        if (phase == 1) {
            if (adding && !now_adding) {        /* the AddNew ended: RowColChange, AfterUpdate before Reposition (VB6 oracle f_dbgrid) */
                QByteArray prev = lastbm;
                reload();
                lastbm = cur_bm();
                fire_rcc(prev, curcol);
                if (add_bu) { add_bu = false; fire_ctl(c, "afterupdate"); }
                add_done = true;
            }
            return;
        }
        if (add_done) { add_done = false; return; }
        if (now_adding) { reload(); return; }   /* AddNew: the new row, no RowColChange (VB6 oracle f_dbgrid) */
        reload();
        QByteArray cb = cur_bm();
        int lastcol = curcol;
        if (want_col >= 0) { curcol = std::min(want_col, std::max(0, (int)cols.size() - 1)); want_col = -1; sync_cell(); }
        if (rcc_pending) lastcol = rcc_lastcol;
        bool any = rcc_pending || cb != lastbm || lastcol != curcol;
        QByteArray prev = lastbm;
        lastbm = cb;
        rcc_pending = false;
        if (any) fire_rcc(prev, lastcol);
    }
    /* the changed row into the recordset: BeforeUpdate, Edit (AddNew) and Update, AfterUpdate; false: not saved */
    bool commit()
    {
        if (unbound()) return ucommit();
        vbobj r = rs();
        Ctl *d = src ? C(src) : nullptr;
        if (!r || !d || !d->data || (pend.empty() && !pend_add)) { pend.clear(); pend_add = false; return true; }
        vbvar a = vb_v_i(0);
        fire_ctl(c, "beforeupdate", 1, &a);
        if (vb_v2i(a)) return false;
        bool add = pend_add;
        if (add) {
            vbvar b = vb_v_i(0);
            fire_ctl(c, "beforeinsert", 1, &b);
            if (vb_v2i(b)) { pend.clear(); pend_add = false; on_addrow = false; reload(); return false; }
        }
        d->data->in_validate = true;           /* the Data control does not see this move */
        vb_i32 m = vb_mark();
        vb_late_o(r, add ? "addnew" : "edit", VBM_METHOD, 0, nullptr);
        for (auto &e : pend) {
            if (vb_errpend) break;
            DbgCol &g = cols[e.first];
            if (g.field.isEmpty()) continue;
            int t = g.fcol >= 0 ? rsx_coltype(r, g.fcol) : 10;
            vbvar v = e.second.isEmpty() && t != 10 && t != 12 ? vb_v_null() : vstr(e.second);
            vbvar fa[2] = { vstr(g.field), v };
            vb_late_o(r, "fields", VBM_LET, 2, fa);
        }
        if (!vb_errpend) vb_late_o(r, "update", VBM_METHOD, 0, nullptr);
        d->data->in_validate = false;
        vb_release(m);
        if (vb_errpend) {
            vb_i32 num = vb_err_number_pending();
            QString msg = QString::fromUtf8(vb_err_text_pending());
            vb_err_clear_pending();
            d->data->in_validate = true;
            vb_late_o(r, "cancelupdate", VBM_METHOD, 0, nullptr);
            if (vb_errpend) vb_err_clear_pending();
            d->data->in_validate = false;
            error_event(num, msg);
            return false;
        }
        pend.clear(); pend_add = false; on_addrow = false;
        fire_ctl(c, "afterupdate");
        if (add) fire_ctl(c, "afterinsert");
        reload();
        return true;
    }
    /* the changed row of an unbound grid to the program: UnboundWriteData (UnboundAddData for the new row) with
     * the changed values, Null in the others (VB6 oracle f_dbgunb) */
    bool ucommit()
    {
        if (pend.empty() && !pend_add) return true;
        vbvar a = vb_v_i(0);
        fire_ctl(c, "beforeupdate", 1, &a);
        if (vb_v2i(a)) return false;
        bool add = pend_add;
        if (add) {
            vbvar b = vb_v_i(0);
            fire_ctl(c, "beforeinsert", 1, &b);
            if (vb_v2i(b)) { pend.clear(); pend_add = false; on_addrow = false; reload(); return false; }
        }
        int nc = (int)cols.size();
        RowBuf *rb = new RowBuf(1, nc);
        for (int k = 0; k < nc; k++) {
            rb->names.push_back(cols[k].field.isEmpty() ? cols[k].caption : cols[k].field);
            vb_vset(&rb->val[k], pend.count(k) ? vstr(pend[k]) : vb_v_null());
        }
        int row = currow;
        if (!add && row >= 0 && row < (int)urows.size()) vb_vset(&rb->bm[0], urows[row].bm);
        vbobj o = vb_onew(&CLS_ROWBUF);
        o->ext = rb;
        vb_oaddref(o);
        vb_i32 m = vb_mark();
        vbvar ev[2] = { vb_v_obj(o), add ? vb_v_null() : vb_vtmp(rb->bm[0]) };
        fire_ctl(c, add ? "unboundadddata" : "unboundwritedata", 2, ev);
        o->ext = nullptr;
        vb_orelease(o);
        bool err = vb_errpend != 0;
        if (!err) {
            if (add) {
                URow u;
                vb_vset(&u.bm, ev[1]);
                u.v.resize((size_t)nc);
                for (int k = 0; k < nc; k++) vb_vset(&u.v[k], rb->val[k]);
                urows.push_back(u);
                currow = (int)urows.size() - 1;
            } else if (row >= 0 && row < (int)urows.size()) {
                for (auto &e : pend) if (e.first < (int)urows[row].v.size()) vb_vset(&urows[row].v[e.first], rb->val[e.first]);
            }
        }
        vb_release(m);
        delete rb;
        if (err) {
            vb_i32 num = vb_err_number_pending();
            QString msg = QString::fromUtf8(vb_err_text_pending());
            vb_err_clear_pending();
            error_event(num, msg);
            return false;
        }
        pend.clear(); pend_add = false; on_addrow = false;
        nrows = (int)urows.size();
        fire_ctl(c, "afterupdate");
        if (add) fire_ctl(c, "afterinsert");
        return true;
    }
    void move_to(int row)                       /* the recordset to a row: Bookmark (Validate 9 ...) */
    {
        vbobj r = rs();
        if (!r) return;
        vb_i32 m = vb_mark();
        vbvar b = dbg_bmvar(bm_of(row));
        vb_late_o(r, "bookmark", VBM_LET, 1, &b);
        vb_release(m);
    }
    void user_move(int r, int col)              /* the user went to a cell */
    {
        if (unbound()) {
            if (r == (on_addrow ? nrows : currow)) {
                if (col != curcol) { int old = curcol; curcol = col; fire_rcc_v(bm_var(currow), old); }
                return;
            }
            if (r >= nrows) {                   /* the new row (AllowAddNew) */
                if ((!pend.empty() || pend_add) && !commit()) { sync_cell(); return; }
                int old = curcol;
                vb_i32 m = vb_mark();
                vbvar last = vb_vtmp(bm_var(currow));
                on_addrow = true; curcol = col;
                fire_rcc_v(last, old);
                vb_release(m);
                return;
            }
            int old = curcol;
            if (on_addrow && pend_add && !commit()) { sync_cell(); return; }
            curcol = col;
            if (!umove(r)) { curcol = old; sync_cell(); }
            return;
        }
        if (r == (on_addrow ? nrows : currow)) {
            if (col != curcol) { int old = curcol; curcol = col; fire_rcc(lastbm, old); }
            return;
        }
        if (on_addrow) {                        /* the new row the user wrote is added */
            on_addrow = false;
            if (pend_add && !commit()) { on_addrow = true; sync_cell(); return; }
            pend.clear(); pend_add = false;
            reload();
        }
        if (r >= nrows) {                       /* the empty row at the end (AllowAddNew) */
            if (!pend.empty() && !commit()) { sync_cell(); return; }
            int old = curcol;
            on_addrow = true; curcol = col;
            fire_rcc(lastbm, old);
            return;
        }
        want_col = col;
        move_to(r);
        if (vb_errpend) vb_unhandled();
        if (want_col >= 0) { want_col = -1; sync_cell(); }   /* it did not move (Validate cancelled it) */
    }
    void user_edit(int r, int k, const QString &t)
    {
        if (k < 0 || k >= (int)cols.size()) return;
        bool add = r >= nrows;
        QString old = add ? QString() : cell(r, k);
        if (!add && r != currow) return;
        fire_ctl(c, "change");
        vbvar a[3] = { vb_v_i((vb_i16)k), vstr(old), vb_v_i(0) };
        vb_i32 m = vb_mark();
        fire_ctl(c, "beforecolupdate", 3, a);
        bool cancel = vb_v2i(a[2]) != 0;
        vb_release(m);
        if (cancel) { building = true; if (tw && tw->item(r, k)) tw->item(r, k)->setText(old); building = false; return; }
        pend[k] = t;
        if (add) { pend_add = true; on_addrow = true; }
        vbvar ci = vb_v_i((vb_i16)k);
        fire_ctl(c, "aftercolupdate", 1, &ci);
    }
    void delete_row()
    {
        if (unbound()) {
            if (currow < 0 || currow >= (int)urows.size()) return;
            vbvar a = vb_v_i(0);
            fire_ctl(c, "beforedelete", 1, &a);
            if (vb_v2i(a)) return;
            pend.clear();
            vb_i32 m = vb_mark();
            vbvar bm = vb_vtmp(urows[currow].bm);
            fire_ctl(c, "unbounddeleterow", 1, &bm);
            vb_release(m);
            if (vb_errpend) {
                vb_i32 num = vb_err_number_pending();
                QString msg = QString::fromUtf8(vb_err_text_pending());
                vb_err_clear_pending();
                error_event(num, msg);
                return;
            }
            vb_vfree(&urows[currow].bm);
            for (auto &v : urows[currow].v) vb_vfree(&v);
            urows.erase(urows.begin() + currow);
            nrows = (int)urows.size();
            fire_ctl(c, "afterdelete");
            reload();
            return;
        }
        vbobj r = rs();
        if (!r || currow < 0 || currow >= nrows) return;
        vbvar a = vb_v_i(0);
        fire_ctl(c, "beforedelete", 1, &a);
        if (vb_v2i(a)) return;
        pend.clear();
        vb_late_o(r, "delete", VBM_METHOD, 0, nullptr);
        if (vb_errpend) {
            vb_i32 num = vb_err_number_pending();
            QString msg = QString::fromUtf8(vb_err_text_pending());
            vb_err_clear_pending();
            error_event(num, msg);
            return;
        }
        fire_ctl(c, "afterdelete");
        vb_late_o(r, "movenext", VBM_METHOD, 0, nullptr);
        if (!vb_errpend && vb_v2bool(vb_late_o(r, "eof", VBM_GET, 0, nullptr)) && nrows > 0) vb_late_o(r, "movelast", VBM_METHOD, 0, nullptr);
        if (vb_errpend) vb_err_clear_pending();
    }
    void rebind()                               /* ReBind: the columns again (the default layout: the fields) */
    {
        int old = curcol;
        pend.clear(); pend_add = false; on_addrow = false;
        if (unbound()) {                        /* the rows again; the cell stays (VB6 oracle f_dbgunb) */
            uread();
            reload();
            fire_rcc_v(bm_var(currow), old);
            return;
        }
        if (!held) auto_columns();
        if (curcol >= (int)cols.size()) curcol = 0;   /* the column stays (VB6 oracle f_dbgunb) */
        reload();
        QByteArray prev = lastbm;
        lastbm = cur_bm();
        fire_rcc(prev, old);                    /* VB6 oracle f_dbgrid: ReBind fires RowColChange */
    }

    /* ------------------------------------------------ the objects */
    vbobj columns_obj() { if (!colsobj) { colsobj = vb_onew(&CLS_DBCOLS); colsobj->ext = this; vb_oaddref(colsobj); } return colsobj; }
    vbobj column_obj(int k) { DbgCol &g = cols[k]; if (!g.obj) { g.obj = vb_onew(&CLS_DBCOL); g.obj->ext = this; vb_oaddref(g.obj); } return g.obj; }
    vbobj splits_obj() { if (!splitsobj) { splitsobj = vb_onew(&CLS_DBSPLITS); splitsobj->ext = this; vb_oaddref(splitsobj); } return splitsobj; }
    vbobj split_obj() { if (!splitobj) { splitobj = vb_onew(&CLS_DBSPLIT); splitobj->ext = this; vb_oaddref(splitobj); } return splitobj; }
    vbobj selbm_obj() { if (!selbmobj) { selbmobj = vb_onew(&CLS_DBSELBM); selbmobj->ext = this; vb_oaddref(selbmobj); } return selbmobj; }
    int col_of_obj(vbobj o) { for (int k = 0; k < (int)cols.size(); k++) if (cols[k].obj == o) return k; return -1; }
    /* Columns(i), Columns("Name"): -1 and the error raised */
    int find_col(vbvar key)
    {
        vbvar k = vb_v_letval(key);
        if (vb_errpend) return -1;
        if (k.vt == VT_BSTR) {
            QString s = qs(k.u.s);
            for (int i = 0; i < (int)cols.size(); i++)
                if (!s.compare(cols[i].caption, Qt::CaseInsensitive) || (!cols[i].field.isEmpty() && !s.compare(cols[i].field, Qt::CaseInsensitive))) return i;
            vb_i32 m = vb_mark();               /* VB6 oracle f_dbgrid: "Column not found, Nope" */
            vb_raise_full(6147, vs(QString("DBGrid")), vs(QString("Column not found, ") + s));
            vb_release(m);
            return -1;
        }
        vb_i32 i = vb_v2l(k);
        if (vb_errpend) return -1;
        if (i < 0 || i >= (vb_i32)cols.size()) { vb_raise(9); return -1; }
        return i;
    }
    int row_arg(vbvar v)                        /* a bookmark argument: its row, -1 */
    {
        if (unbound()) return urow_of(v);
        QByteArray b = dbg_bytes(v);
        return dao_bm_row(rs(), b);
    }
    vbvar col_member(int k, const std::string &n, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found);

    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        if (n == "columns") {
            if (argc == 0 || set) return retobj(columns_obj());
            int k = find_col(argv[0]);
            return k < 0 ? vb_v_empty() : retobj(column_obj(k));
        }
        if (n == "splits") {
            if (argc == 0) return retobj(splits_obj());
            if (vb_v2l(argv[0]) != 0) { dbg_raise(6244); return vb_v_empty(); }
            return retobj(split_obj());
        }
        if (n == "selbookmarks") {
            if (argc == 0) return retobj(selbm_obj());
            int i = (int)vb_v2l(argv[0]);
            if (i < 0 || i >= (int)selbm.size()) { dbg_raise(6150); return vb_v_empty(); }
            return dbg_bmstr(selbm[i]);
        }
        if (n == "datasource") {
            if (!set) return vb_vtmp(vb_v_obj(src));
            vbobj o = vb_v2obj(val);
            if (vb_errpend) return vb_v_empty();
            if (o) vb_oaddref(o);
            if (src) vb_orelease(src);
            src = o ? vb_oid(o) : nullptr;
            srcname = src && is_ctl_obj(src) ? C(src)->name : QString();
            bound_once = false;
            if (src) source_reset(src);
            return vb_v_empty();
        }
        if (n == "row") {
            if (!set) return vb_v_i((vb_i16)(on_addrow ? nrows - top : currow < 0 ? -1 : currow - top));
            int r = (int)vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (r < 0 || top + r >= nrows || r >= std::max(1, visible_rows())) { dbg_raise(6148); return vb_v_empty(); }
            if (unbound()) { if (top + r != currow || on_addrow) umove(top + r); return vb_v_empty(); }
            if (top + r != currow || on_addrow) { on_addrow = false; move_to(top + r); }
            return vb_v_empty();
        }
        if (n == "col") {
            if (!set) return vb_v_i((vb_i16)curcol);
            int k = (int)vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (k < 0 || k >= (int)cols.size()) { dbg_raise(6145); return vb_v_empty(); }
            if (k != curcol) {
                int old = curcol; curcol = k; sync_cell();
                if (unbound()) fire_rcc_v(bm_var(currow), old); else fire_rcc(lastbm, old);
            }
            return vb_v_empty();
        }
        if (n == "text") {
            if (!set) return vstr(cell(on_addrow ? nrows : currow, curcol));
            if (curcol >= (int)cols.size() || (currow < 0 && !on_addrow)) return vb_v_empty();
            pend[curcol] = data_qs(val);
            if (on_addrow) pend_add = true;
            int r = on_addrow ? nrows : currow;
            building = true;
            if (tw && tw->item(r, curcol)) tw->item(r, curcol)->setText(pend[curcol]);
            building = false;
            return vb_v_empty();
        }
        if (n == "bookmark" && unbound()) {     /* the rows again, the row in the middle (VB6 oracle f_dbgunb) */
            if (!set) return on_addrow ? vb_v_null() : bm_var(currow);
            int r = urow_of(val);
            if (r < 0) { dbg_raise(6149); return vb_v_empty(); }
            if ((!pend.empty() || pend_add) && !commit()) return vb_v_empty();
            vb_i32 m = vb_mark();
            vbvar want = vb_vtmp(urows[r].bm), last = vb_vtmp(bm_var(currow));
            uread();
            r = urow_of(want);
            if (r >= 0) {
                currow = r; on_addrow = false;
                top = std::max(0, r - std::max(1, visible_rows()) / 2);
            }
            paint();
            fire_rcc_v(last, curcol);
            vb_release(m);
            return vb_v_empty();
        }
        if (n == "bookmark") {
            if (!set) return currow >= 0 && currow < nrows && !on_addrow ? dbg_bmstr(cur_bm()) : vb_v_null();
            QByteArray b = dbg_bytes(val);
            int r = dao_bm_row(rs(), b);
            if (r < 0) { dbg_raise(6149); return vb_v_empty(); }
            on_addrow = false;
            move_to(r);
            return vb_v_empty();
        }
        if (n == "firstrow") {
            if (!set) return nrows ? bm_var(top) : vb_v_null();
            int r = row_arg(val);
            if (r < 0) { dbg_raise(6149); return vb_v_empty(); }
            top = r;
            if (tw) tw->verticalScrollBar()->setValue(top);
            return vb_v_empty();
        }
        if (n == "leftcol") {
            if (!set) return vb_v_i((vb_i16)leftcol);
            int k = (int)vb_v2l(val);
            if (k < 0 || k >= std::max(1, (int)cols.size())) { dbg_raise(6145); return vb_v_empty(); }
            leftcol = k;
            if (tw && k < tw->columnCount()) tw->scrollToItem(tw->item(std::max(0, top), k), QAbstractItemView::PositionAtTop);
            return vb_v_empty();
        }
        if (n == "getbookmark") {
            int r = (on_addrow ? nrows : currow) + (int)arg_l(argc, argv, 0, 0);
            return r >= 0 && r < nrows ? bm_var(r) : vb_v_null();
        }
        if (n == "rowbookmark") {
            int r = top + (int)arg_l(argc, argv, 0, 0);
            if (r < 0 || r >= nrows) { dbg_raise(6148); return vb_v_empty(); }
            return bm_var(r);
        }
        if (n == "approxcount") { if (set) return vb_v_empty(); return vb_v_l(nrows); }   /* (set: a hint for the scroll bar) */
        if (n == "visiblerows") return vb_v_i((vb_i16)visible_rows());
        if (n == "visiblecols") { int k = 0; for (auto &g : cols) if (g.visible) k++; return vb_v_i((vb_i16)k); }
        if (n == "rebind") { rebind(); return vb_v_empty(); }
        if (n == "refresh") {
            if (unbound()) { uread(); reload(); fire_rcc_v(bm_var(currow), curcol); return vb_v_empty(); }   /* VB6 oracle f_dbgunb */
            reload();
            return vb_v_empty();
        }
        if (n == "headfont") {
            if (!headfont) {
                headfont = vb_onew(&CLS_FONT);
                vb_oaddref(headfont);
                ((FontRef *)headfont->ext)->own = headvf;
            }
            FontRef *fr = (FontRef *)headfont->ext;
            if (!set) return retobj(headfont);
            vbobj fo = val.vt == VT_DISPATCH ? vb_oid(val.u.o) : nullptr;
            if (!fo || fo->cls != &CLS_FONT) { vb_raise(13); return vb_v_empty(); }
            FontRef *r2 = (FontRef *)fo->ext;
            VFont *srcf = r2->owner ? font_of(r2->owner) : &r2->own;
            if (srcf) { fr->own = *srcf; headvf = *srcf; }
            if (tw) tw->horizontalHeader()->setFont(qfont(headvf));
            return vb_v_empty();
        }
        if (n == "holdfields") { held = true; return vb_v_empty(); }
        if (n == "clearfields") { held = false; return vb_v_empty(); }   /* the next ReBind makes the columns (VB6 oracle f_dbgrid) */
        if (n == "clearselcols" || n == "aboutbox") return vb_v_empty();
        if (n == "scroll") {
            int dc = (int)arg_l(argc, argv, 0, 0), dr = (int)arg_l(argc, argv, 1, 0);
            int nt = top + dr, nl = leftcol + dc;
            if (nt < 0 || (nrows && nt >= nrows) || nl < 0 || (nl && nl >= (int)cols.size())) { dbg_raise(6151); return vb_v_empty(); }
            if (dr > 0) nt = std::min(nt, std::max(top, nrows - std::max(1, visible_rows())));   /* not past the last row (VB6 oracle f_dbgunb) */
            vbvar a = vb_v_i(0);
            fire_ctl(c, "scroll", 1, &a);
            if (vb_v2i(a)) return vb_v_empty();
            top = nt; leftcol = nl;
            if (tw) tw->verticalScrollBar()->setValue(top);
            return vb_v_empty();
        }
        if (n == "rowtop") {
            int r = (int)arg_l(argc, argv, 0, 0);
            if (!tw || r < 0) return vb_v_f(0);
            int y = tw->rowViewportPosition(top + r) + (tw->horizontalHeader()->isVisible() ? tw->horizontalHeader()->height() : 0);
            return vb_v_f((float)from_twips_ctl(c, y * TPP));
        }
        if (n == "colcontaining" || n == "rowcontaining") {
            double v = argc ? to_twips_ctl(c, vb_v2d(argv[0])) : 0;
            int px = (int)(v / TPP);
            if (!tw) return vb_v_i(-1);
            if (n == "colcontaining") return vb_v_i((vb_i16)tw->columnAt(px - (tw->verticalHeader()->isVisible() ? tw->verticalHeader()->width() : 0)));
            int r = tw->rowAt(px - (tw->horizontalHeader()->isVisible() ? tw->horizontalHeader()->height() : 0));
            return vb_v_i((vb_i16)(r < 0 ? -1 : r - top));
        }
        if (n == "splitcontaining" || n == "split") return vb_v_i(0);
        if (n == "captureimage") return tw ? retobj(new_pic(box ? box->grab().toImage() : tw->grab().toImage(), 1)) : vb_v_empty();
        if (n == "datachanged") {
            if (!set) return vbool(!pend.empty());
            if (!vb_v2bool(val)) { pend.clear(); pend_add = false; reload(); }
            return vb_v_empty();
        }
        if (n == "currentcellmodified") return vbool(pend.count(curcol) != 0);
        if (n == "editactive") return vbool(editing());
        if (n == "addnewmode") return vb_v_l(on_addrow ? (pend_add ? 2 : 1) : adding ? 1 : 0);
        if (n == "currentcellvisible") return vbool(true);
        if (n == "errortext") return vstr(errortext);
        if (n == "selstart" || n == "sellength") return vb_v_l(0);
        if (n == "seltext") return vstr(QString());
        if (n == "selstartcol" || n == "selendcol") return vb_v_i(-1);
        if (n == "hwndeditor") return vb_v_l(0);
        if (n == "caption") { if (set) { caption = data_qs(val); look(); return vb_v_empty(); } return vstr(caption); }
        if (n == "headlines" || n == "defcolwidth" || n == "rowheight") {
            double *p = n == "headlines" ? &headlines : n == "defcolwidth" ? &defcolwidth : &rowheight;
            if (!set) {
                if (n == "rowheight" && rowheight <= 0) return vb_v_f((float)from_twips_ctl(c, row_px() * TPP));
                return vb_v_f((float)(n == "headlines" ? *p : from_twips_ctl(c, *p)));
            }
            *p = n == "headlines" ? vb_v2d(val) : to_twips_ctl(c, vb_v2d(val));
            reload();
            return vb_v_empty();
        }
        struct { const char *name; bool *p; } bools[] = {
            { "allowaddnew", &allowadd }, { "allowdelete", &allowdelete }, { "allowupdate", &allowupdate },
            { "allowrowsizing", &allowrowsizing }, { "allowarrows", &allowarrows }, { "columnheaders", &colheaders },
            { "recordselectors", &recsel }, { "marqueeunique", &marqueeunique }, { "tabacrosssplits", &tabacross },
            { "wrapcellpointer", &wrapcell } };
        for (auto &e : bools)
            if (n == e.name) { if (!set) return vbool(*e.p); *e.p = vb_v2bool(val) != 0; look(); reload(); return vb_v_empty(); }
        struct { const char *name; int *p; int lo, hi; } ints[] = {
            { "datamode", &datamode, 0, 1 }, { "marqueestyle", &marquee, 0, 6 }, { "scrollbars", &scrollbars, 0, 4 },
            { "rowdividerstyle", &rowdivider, 0, 6 }, { "appearance", &appearance, 0, 1 }, { "borderstyle", &borderstyle, 0, 1 },
            { "tabaction", &tabaction, 0, 2 } };
        for (auto &e : ints)
            if (n == e.name) {
                if (!set) return vb_v_l(*e.p);
                int v = (int)vb_v2l(val);
                if (v < e.lo || v > e.hi) { vb_raise(380); return vb_v_empty(); }
                *e.p = v; look();
                return vb_v_empty();
            }
        if (n == "backcolor" || n == "forecolor") {
            long *p = n == "backcolor" ? &back : &fore;
            if (!set) return vb_v_l((vb_i32)*p);
            *p = vb_v2l(val);
            return vb_v_empty();
        }
        *found = 0;
        return vb_v_empty();
    }
};

static OcxDBGrid *dbgrid_of(vbobj o) { o = vb_oid(o); return o ? (OcxDBGrid *)o->ext : nullptr; }

vbvar OcxDBGrid::col_member(int k, const std::string &n, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    DbgCol &g = cols[k];
    int row = on_addrow ? nrows : currow;
    if (n == "caption") { if (!set) return vstr(g.caption); g.caption = data_qs(val); paint(); return vb_v_empty(); }
    if (n == "datafield") {
        if (!set) return vstr(g.field);
        g.field = data_qs(val);
        reload();
        return vb_v_empty();
    }
    if (n == "value" || n == "text" || n.empty()) {
        if (!set) return n == "text" ? vstr(cell(row, k)) : value(row, k);
        if (row < 0) return vb_v_empty();
        pend[k] = data_qs(val);
        if (on_addrow) pend_add = true;
        building = true;
        if (tw && tw->item(row, k)) tw->item(row, k)->setText(pend[k]);
        building = false;
        return vb_v_empty();
    }
    if (n == "celltext" || n == "cellvalue") {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int r = row_arg(argv[0]);
        if (r < 0) { dbg_raise(6149); return vb_v_empty(); }
        if (unbound()) return n == "celltext" ? vstr(cell(r, k)) : value(r, k);   /* CellValue: the program's value */
        if (g.fcol < 0) return n == "celltext" ? vstr(QString()) : vb_v_empty();
        return vstr(cell(r, k));
    }
    if (n == "colindex") return vb_v_i((vb_i16)k);
    if (n == "width") {
        if (!set) {
            if (g.him < 0 && tw && k < tw->columnCount()) g.him = (int)std::lround(tw->columnWidth(k) * 2540.0 / 96.0);
            return vb_v_f((float)from_twips_ctl(c, dbg_him2tw(std::max(0, g.him))));
        }
        double t = to_twips_ctl(c, vb_v2d(val));
        if (vb_errpend) return vb_v_empty();
        if (t < 0) { vb_raise(380); return vb_v_empty(); }
        g.him = dbg_tw2him(t);
        building = true;
        if (tw && k < tw->columnCount()) tw->setColumnWidth(k, dbg_him2px(g.him));
        building = false;
        return vb_v_empty();
    }
    if (n == "visible") { if (!set) return vbool(g.visible); g.visible = vb_v2bool(val) != 0; if (tw && k < tw->columnCount()) tw->setColumnHidden(k, !g.visible); return vb_v_empty(); }
    if (n == "locked") { if (!set) return vbool(g.locked); g.locked = vb_v2bool(val) != 0; paint(); return vb_v_empty(); }
    if (n == "alignment") {
        if (!set) return vb_v_l(g.align == 3 ? (general_num(k) ? 1 : 0) : g.align);   /* General: as it shows (VB6 oracles f_dbgrid, f_dbgw) */
        int v = (int)vb_v2l(val);
        if (v < 0 || v > 3) { vb_raise(380); return vb_v_empty(); }
        g.align = v; paint();
        return vb_v_empty();
    }
    if (n == "numberformat") { if (!set) return vstr(g.numfmt); g.numfmt = data_qs(val); paint(); return vb_v_empty(); }
    if (n == "defaultvalue") { if (!set) return vstr(g.defval); g.defval = data_qs(val); return vb_v_empty(); }
    if (n == "datachanged") return vbool(pend.count(k) != 0);
    if (n == "allowsizing" || n == "wraptext" || n == "button") {
        bool *p = n == "allowsizing" ? &g.allowsizing : n == "wraptext" ? &g.wraptext : &g.button;
        if (!set) return vbool(*p);
        *p = vb_v2bool(val) != 0;
        return vb_v_empty();
    }
    if (n == "dividerstyle") { if (!set) return vb_v_l(g.divider); g.divider = (int)vb_v2l(val); return vb_v_empty(); }
    if (n == "left" || n == "top") {
        if (!tw || k >= tw->columnCount()) return vb_v_f(0);
        int px = n == "left" ? tw->columnViewportPosition(k) + (tw->verticalHeader()->isVisible() ? tw->verticalHeader()->width() : 0) : 0;
        return vb_v_f((float)from_twips_ctl(c, px * TPP));
    }
    *found = 0;
    return vb_v_empty();
}

static vbvar dbcol_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    auto *g = dbgrid_of(o);
    *found = 1;
    if (!g) { vb_raise(91); return vb_v_empty(); }
    int k = g->col_of_obj(vb_oid(o));
    if (k < 0) { vb_raise(91); return vb_v_empty(); }      /* a column taken away (ReBind, Remove) */
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    return g->col_member(k, lower(name), set, val, set ? argc - 1 : argc, argv, found);
}
static vb_i32 dbcols_count(vbobj o) { auto *g = dbgrid_of(o); return g ? (vb_i32)g->cols.size() : 0; }
static vbvar dbcols_item(vbobj o, vb_i32 i) { auto *g = dbgrid_of(o); return g && i >= 0 && i < (vb_i32)g->cols.size() ? retobj(g->column_obj(i)) : vb_v_empty(); }
static vbvar dbcols_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    auto *g = dbgrid_of(o);
    std::string n = lower(name);
    *found = 1;
    if (!g) { vb_raise(91); return vb_v_empty(); }
    if (n == "count") return vb_v_l((vb_i32)g->cols.size());
    if (n == "item" || n.empty()) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int k = g->find_col(argv[0]);
        return k < 0 ? vb_v_empty() : retobj(g->column_obj(k));
    }
    if (n == "add") {
        int pos = (int)arg_l(argc, argv, 0, (long)g->cols.size());
        if (pos < 0 || pos > (int)g->cols.size()) { vb_raise(9); return vb_v_empty(); }
        g->cols.insert(g->cols.begin() + pos, DbgCol());
        if (g->curcol >= pos && g->curcol < (int)g->cols.size() - 1 && pos <= g->curcol) g->curcol++;
        g->paint();
        return retobj(g->column_obj(pos));
    }
    if (n == "remove") {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int k = g->find_col(argv[0]);
        if (k < 0) return vb_v_empty();
        OcxDBGrid::drop(g->cols[k]);
        g->cols.erase(g->cols.begin() + k);
        g->pend.erase(k);
        if (g->curcol > k) g->curcol--;
        if (g->curcol >= (int)g->cols.size()) g->curcol = std::max(0, (int)g->cols.size() - 1);
        g->paint();
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
static vbvar dbsplits_item(vbobj o, vb_i32 i) { auto *g = dbgrid_of(o); return g && i == 0 ? retobj(g->split_obj()) : vb_v_empty(); }
static vbvar dbsplits_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    auto *g = dbgrid_of(o);
    std::string n = lower(name);
    *found = 1;
    if (!g) { vb_raise(91); return vb_v_empty(); }
    if (n == "count") return vb_v_l(1);
    if (n == "item" || n.empty()) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        if (vb_v2l(argv[0]) != 0) { dbg_raise(6244); return vb_v_empty(); }
        return retobj(g->split_obj());
    }
    if (n == "add" || n == "remove") { dbg_raise(6244); return vb_v_empty(); }   /* one split */
    *found = 0;
    return vb_v_empty();
}
/* the one Split: its properties are those of the grid */
static vbvar dbsplit_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    auto *g = dbgrid_of(o);
    std::string n = lower(name);
    *found = 1;
    if (!g) { vb_raise(91); return vb_v_empty(); }
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    if (n == "index") return vb_v_i(0);
    struct { const char *name; bool *p; } bools[] = { { "locked", &g->split_locked }, { "allowfocus", &g->split_allowfocus },
                                                      { "allowsizing", &g->split_allowsizing } };
    for (auto &e : bools) if (n == e.name) { if (!set) return vbool(*e.p); *e.p = vb_v2bool(val) != 0; g->paint(); return vb_v_empty(); }
    if (n == "scrollgroup" || n == "sizemode") {
        int *p = n == "scrollgroup" ? &g->scrollgroup : &g->sizemode;
        if (!set) return vb_v_i((vb_i16)*p);
        *p = (int)vb_v2l(val);
        return vb_v_empty();
    }
    if (n == "size") { if (!set) return vb_v_d(g->split_size); g->split_size = vb_v2d(val); return vb_v_empty(); }
    if (n == "columns" || n == "marqueestyle" || n == "recordselectors" || n == "scrollbars" || n == "allowrowsizing" ||
        n == "firstrow" || n == "leftcol" || n == "selstartcol" || n == "selendcol" || n == "currentcellvisible" || n == "clearselcols")
        return g->member(n, kind, set, val, set ? argc - 1 : argc, argv, found);
    *found = 0;
    return vb_v_empty();
}
static vb_i32 dbselbm_count(vbobj o) { auto *g = dbgrid_of(o); return g ? (vb_i32)g->selbm.size() : 0; }
static vbvar dbselbm_item(vbobj o, vb_i32 i) { auto *g = dbgrid_of(o); return g && i >= 0 && i < (vb_i32)g->selbm.size() ? dbg_bmstr(g->selbm[i]) : vb_v_empty(); }
static vbvar dbselbm_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    auto *g = dbgrid_of(o);
    std::string n = lower(name);
    *found = 1;
    if (!g) { vb_raise(91); return vb_v_empty(); }
    if (n == "count") return vb_v_l((vb_i32)g->selbm.size());
    if (n == "item" || n.empty() || n == "remove") {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = (int)vb_v2l(argv[0]);
        if (i < 0 || i >= (int)g->selbm.size()) { dbg_raise(6150); return vb_v_empty(); }
        if (n == "remove") { g->selbm.erase(g->selbm.begin() + i); return vb_v_empty(); }
        return dbg_bmstr(g->selbm[i]);
    }
    if (n == "add") {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        QByteArray b = dbg_bytes(argv[0]);
        if (dao_bm_row(g->rs(), b) < 0) { dbg_raise(6149); return vb_v_empty(); }
        g->selbm.push_back(b);
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
static Ocx *new_dbgrid() { return new OcxDBGrid; }

/* ---------------------------------------------------------------- DBList, DBCombo */
/* the lists of DataList / DataCombo, with what DBList / DBCombo do otherwise (VB6 oracle f_dbgrid): Click
 * (DBList) and Change (DBCombo) when the program or the Data control changes the choice; a text that the
 * list does not have; SelectedItem and VisibleItems are bookmarks of the RowSource */
struct OcxDBList : OcxDataList {
    int matchentry = 0, selstart = 0, sellen = 0;
    bool integral = true;
    explicit OcxDBList(bool combo) : OcxDataList(combo) {}
    QWidget *make(QWidget *parent) override
    {
        vbobj self = c->self;
        auto me = [self]() -> OcxDBList * { Ctl *x = C(self); return x ? dynamic_cast<OcxDBList *>(x->ocx) : nullptr; };
        if (combo) {
            cb = new QComboBox(parent);
            cb->setEditable(style != 2);
            QObject::connect(cb, QOverload<int>::of(&QComboBox::activated), [me](int i) {
                auto *d = me();
                if (!d || d->filling_list) return;
                QString ot = d->text;
                d->pick(i);
                vbvar a = vb_v_i(2);           /* dbcAreaList */
                fire_ctl(d->c, "click", 1, &a);
                if ((d = me()) && d->text != ot) fire_ctl(d->c, "change");
            });
            if (cb->lineEdit()) QObject::connect(cb->lineEdit(), &QLineEdit::textEdited, [me](const QString &t) {
                auto *d = me();
                if (!d || d->filling_list) return;
                d->text = t;
                int i = d->find_item(t);
                if (i >= 0) { d->sel = i; d->bound = d->keys[i]; d->matched = true; } else d->matched = false;
                fire_ctl(d->c, "change");
            });
            return cb;
        }
        lw = new QListWidget(parent);
        QObject::connect(lw, &QListWidget::currentRowChanged, [me](int i) {
            auto *d = me();
            if (d && !d->filling_list && i >= 0 && i != d->sel) { d->pick(i); fire_ctl(d->c, "click"); }
        });
        QObject::connect(lw, &QListWidget::itemDoubleClicked, [self](QListWidgetItem *) { Ctl *x = C(self); if (x) fire_ctl(x, "dblclick"); });
        return lw;
    }
    void load(const char *const *p) override
    {
        OcxDataList::load(p);
        QString b;
        if (prop(p, "Bindings", &b)) {
            QString n = bindings_for(QByteArray::fromHex(b.toLatin1()), combo ? 16 : 3);
            if (!n.isEmpty()) rowsrcname = n;
        }
        matchentry = (int)propn(p, "MatchEntry", 0);
        integral = propn(p, "IntegralHeight", -1) != 0;
    }
    void source_reset(vbobj s) override
    {
        if (!rowsrc && !rowsrcname.isEmpty() && is_ctl_obj(s) && !rowsrcname.compare(C(s)->name, Qt::CaseInsensitive)) { rowsrc = vb_oid(s); vb_oaddref(rowsrc); }
        if (rowsrc && vb_oid(rowsrc) == vb_oid(s)) { bind_rows(); resync(); }
    }
    int find_key(const QString &v) { for (int i = 0; i < (int)keys.size(); i++) if (keys[i] == v) return i; return -1; }
    int find_item(const QString &v) { for (int i = 0; i < (int)items.size(); i++) if (!items[i].compare(v, Qt::CaseInsensitive)) return i; return -1; }
    /* the list again: the choice of the BoundText it has (the list read later than the Data control filled it) */
    void resync()
    {
        int i = find_key(bound);
        if (i >= 0) { sel = i; text = items[i]; matched = true; }
        else { sel = -1; matched = false; if (!combo) text.clear(); }
        show_sel();
    }
    void notify(int osel, const QString &otext)
    {
        if (!combo) { if (sel != osel || text != otext) fire_ctl(c, "click"); }
        else if (text != otext) fire_ctl(c, "change");
    }
    void set_bound_db(const QString &v)
    {
        int osel = sel; QString ot = text;
        int i = find_key(v);
        if (i >= 0) { sel = i; text = items[i]; bound = keys[i]; matched = true; }
        else { sel = -1; bound = v; matched = false; text.clear(); }
        show_sel();
        notify(osel, ot);
    }
    void set_text_db(const QString &v)
    {
        int osel = sel; QString ot = text;
        int i = find_item(v);
        if (i >= 0) { sel = i; text = items[i]; bound = keys[i]; matched = true; }
        else if (!combo) { sel = -1; text.clear(); bound.clear(); matched = false; }   /* the list has no such text */
        else { text = v; matched = false; }     /* DBCombo: BoundText, SelectedItem stay (VB6 oracle f_dbgrid) */
        show_sel();
        notify(osel, ot);
    }
    int visible_count()
    {
        if (combo) return cb && cb->view() && cb->view()->isVisible() ? std::min(8, std::max(1, cb->count())) : 0;   /* closed: 0 (VB6 oracle f_dbgunb) */
        int ih = QFontMetrics(lw ? lw->font() : qfont(c->font)).height();
        return std::max(1, ((int)(c->height / TPP) - 4) / std::max(1, ih));
    }
    vbvar bookmark(int i)
    {
        if (i < 0) return vstr(QString());     /* VB6 oracle f_dbgrid: no choice, "" */
        if (rs_is_dao(rowrs)) return vb_dbhooks.dao_bookmark(rowrs, i);
        return vb_v_d(i + 1);
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        if (n == "boundtext") {
            if (!set) return vstr(bound);
            vbvar v = vb_v_letval(val);
            set_bound_db(v.vt == VT_NULL || v.vt == VT_EMPTY ? QString() : data_qs(v));
            if (!data_filling && !c->datafield.isEmpty()) c->datachanged = true;
            return vb_v_empty();
        }
        if (n == "text" || n == "deftext" || n.empty()) {
            if (!set) { if (cb && cb->isEditable()) text = cb->currentText(); return vstr(text); }
            vbvar v = vb_v_letval(val);
            set_text_db(v.vt == VT_NULL || v.vt == VT_EMPTY ? QString() : data_qs(v));
            if (!data_filling && !c->datafield.isEmpty()) c->datachanged = true;
            return vb_v_empty();
        }
        if (n == "selecteditem") return bookmark(sel);
        if (n == "rowsource" && !set) return vb_vtmp(vb_v_obj(nullptr));   /* VB6 oracle f_dbgunb: Nothing at run time */
        if (n == "visiblecount") return vb_v_i((vb_i16)visible_count());
        if (n == "visibleitems") {
            if (argc < 1) { vb_raise(450); return vb_v_empty(); }
            int i = (int)vb_v2l(argv[0]);
            int first = lw ? std::max(0, lw->indexAt(QPoint(2, 2)).row()) : 0;
            if (i < 0 || i >= visible_count() || first + i >= (int)items.size()) { vb_raise(380); return vb_v_empty(); }
            return bookmark(first + i);
        }
        if (n == "refill") { bind_rows(); resync(); return vb_v_empty(); }   /* the list again, its choice kept */
        if (n == "refresh") { if (lw) lw->viewport()->update(); return vb_v_empty(); }
        if (n == "matchentry") {
            if (!set) return vb_v_l(matchentry);
            int v = (int)vb_v2l(val);
            if (v < 0 || v > 1) { vb_raise(380); return vb_v_empty(); }
            matchentry = v;
            return vb_v_empty();
        }
        if (n == "integralheight") { if (!set) return vbool(integral); integral = vb_v2bool(val) != 0; return vb_v_empty(); }
        if (combo && (n == "selstart" || n == "sellength" || n == "seltext")) {
            QLineEdit *le = cb ? cb->lineEdit() : nullptr;
            if (!set) {
                if (n == "seltext") return vstr(le ? le->selectedText() : text.mid(selstart, sellen));
                if (le) return vb_v_l(n == "selstart" ? (le->hasSelectedText() ? le->selectionStart() : le->cursorPosition()) : (vb_i32)le->selectedText().size());
                return vb_v_l(n == "selstart" ? selstart : sellen);
            }
            if (n == "seltext") {
                QString s = data_qs(val);
                if (le) { le->insert(s); text = le->text(); }
                else { text = text.left(selstart) + s + text.mid(selstart + sellen); sellen = 0; }
                fire_ctl(c, "change");
                return vb_v_empty();
            }
            int v = (int)vb_v2l(val);
            if (v < 0) { vb_raise(380); return vb_v_empty(); }
            if (n == "selstart") { selstart = std::min(v, (int)text.size()); sellen = 0; }
            else sellen = std::min(v, (int)text.size() - selstart);
            if (le) le->setSelection(selstart, sellen);
            return vb_v_empty();
        }
        if (n == "smoothscroll" || n == "righttoleft") return vbool(false);
        if (n == "cacheup" || n == "cachedown") return vb_v_i(0);
        if (n == "aboutbox") return vb_v_empty();
        return OcxDataList::member(n, kind, set, val, argc, argv, found);
    }
};
static Ocx *new_dblist() { return new OcxDBList(false); }
static Ocx *new_dbcombo() { return new OcxDBList(true); }
