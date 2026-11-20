/* data.inc - the Data control (VB.Data) and the controls bound to it (DataSource, DataField,
 * DataChanged): TextBox, Label, CheckBox, Image, PictureBox, ListBox, ComboBox. The database and
 * the recordset are DAO objects of libvbdb, reached through vb_dbhooks and late binding (the forms
 * runtime does not link the database runtime). The recordset tells the Data control before it moves
 * (Validate: the changed bound values are saved) and after (Reposition: the bound controls show the
 * new record). Included by vbforms.cpp.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

struct DataCtl {
    QString dbname, source, connect = "Access";
    int rstype = 1, options = 0, eofaction = 0, bofaction = 0, deftype = 2, defcursor = 0;
    bool exclusive = false, readonly = false;
    vbobj db = nullptr, rs = nullptr;           /* DAO Database, Recordset (counted) */
    bool in_validate = false;
    QLabel *label = nullptr;
};


static QString data_qs(vbvar v) { return QString::fromUtf8(vb_stoutf8(vb_v2s(v))); }

/* the Data control a bound control names (DataSource = "Data1") */
static Ctl *data_source_of(Ctl *c)
{
    if (!c->form || c->datasource.isEmpty()) return nullptr;
    auto it = c->form->byname.find(lower(c->datasource.toUtf8().constData()));
    if (it == c->form->byname.end() || it->second->cls == &CLS_ARRAY) return nullptr;
    Ctl *d = C(it->second);
    return d && d->kind == K_DATA ? d : nullptr;
}

/* a recordset (DAO or ADO) with a current record (or one being added) */
static bool rs_has_record(vbobj rs)
{
    if (!rs) return false;
    vb_i32 save = vb_errpend;
    bool bof = vb_v2bool(vb_late_o(rs, "bof", VBM_GET, 0, nullptr)) != 0;
    bool eof = vb_v2bool(vb_late_o(rs, "eof", VBM_GET, 0, nullptr)) != 0;
    if (vb_errpend && !save) { vb_err_clear_pending(); return false; }
    int em = vb_v2i(vb_late_o(rs, "editmode", VBM_GET, 0, nullptr));
    if (vb_errpend && !save) { vb_err_clear_pending(); return false; }
    return (!bof && !eof) || em == 2;
}
static bool data_has_record(DataCtl *d) { return rs_has_record(d->rs); }

static vbvar rs_field(vbobj rs, const QString &field, bool *ok)
{
    *ok = false;
    vb_i32 save = vb_errpend;
    vbvar key = vstr(field);
    vbvar f = vb_late_o(rs, "fields", VBM_GET, 1, &key);
    if (vb_errpend && !save) { vb_err_clear_pending(); return vb_v_empty(); }
    vbvar v = vb_late(f, "value", VBM_GET, 0, nullptr);
    if (vb_errpend && !save) { vb_err_clear_pending(); return vb_v_empty(); }
    *ok = true;
    return v;
}

static vbvar ctl_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static void ctl_let(Ctl *c, const char *prop, vbvar v)
{
    int found;
    ctl_member(c->self, prop, VBM_LET, 1, &v, &found);
}
static vbvar ctl_get(Ctl *c, const char *prop)
{
    int found;
    return ctl_member(c->self, prop, VBM_GET, 0, nullptr, &found);
}

/* the DataFormat of the .frm (BeginProperty DataFormat, the props "<pre>DataFormat.Type" ...): a new StdDataFormat
 * (counted), or none */
static vbobj dfmt_props(const char *const *p, const QString &pre)
{
    QString d = pre + "DataFormat.";
    auto key = [&](const char *n) { return (d + n).toUtf8(); };
    if (!prop(p, key("Type").constData(), nullptr)) return nullptr;
    vbobj f = vb_dfmt_new();
    vb_i32 m = vb_mark();
    QString fmt, tv, fv, nv;
    prop(p, key("Format").constData(), &fmt);
    bool tfn = propn(p, key("HaveTrueFalseNull").constData(), 0) != 0;
    prop(p, key("TrueValue").constData(), &tv);
    prop(p, key("FalseValue").constData(), &fv);
    prop(p, key("NullValue").constData(), &nv);
    vb_dfmt_load(f, (vb_i32)propn(p, key("Type").constData(), 0), vs(fmt), tfn ? 1 : 0, vstr(tv), vstr(fv), vstr(nv),
                 (vb_i32)propn(p, key("FirstDayOfWeek").constData(), 0), (vb_i32)propn(p, key("FirstWeekOfYear").constData(), 0),
                 (vb_i32)propn(p, key("LCID").constData(), 0));
    vb_release(m);
    return f;
}
/* a bound control: its DataFormat of the .frm is the one its binding uses */
static void dfmt_from_props(Ctl *c, const char *const *p)
{
    c->dfmt = dfmt_props(p, QString());
    c->dfmt_on = c->dfmt != nullptr;
}

/* a bound control shows the current record of rs (rec: there is one). ADO: through its DataFormat when it has
 * one (the .frm, or Set in code; VB6 oracles f_dfmt, f_dfmt2; the Data control of DAO does not use it). keep: the
 * values a format refused stay changed (DataChanged), as in VB6 */
static void fill_ctl(Ctl *c, vbobj rs, bool rec, bool ado = false, bool keep = false)
{
    vb_i32 m = vb_mark();
    bool ok = false;
    bool was = data_filling;
    data_filling = true;
    vbvar v = rec ? rs_field(rs, c->datafield, &ok) : vb_v_null();
    if (!rec) ok = true;
    bool fmt = ado && c->dfmt_on && c->dfmt;
    if (ok && fmt) {
        vbvar r = vb_dfmt_format(c->dfmt, v, c->self);
        if (vb_errpend) vb_err_clear_pending();
        switch (c->kind) {
        case K_TEXT: case K_COMBO: case K_LIST: case K_LABEL: {
            bool null = r.vt == VT_NULL || r.vt == VT_EMPTY;
            vb_i32 save = vb_errpend;
            vbvar s = null ? vstr(QString()) : vb_v_s(vb_v2s(r));
            if (vb_errpend && !save) { vb_err_clear_pending(); s = vstr(QString()); }
            ctl_let(c, c->kind == K_LABEL ? "caption" : "text", s);
            break; }
        case K_CHECK:                           /* 1 / 0 of a Checkbox format; a text (Boolean format) does not go */
            ctl_let(c, "value", r.vt == VT_NULL || r.vt == VT_EMPTY ? vb_v_i(2) : r);
            break;
        case K_OCX: {
            int found;
            ctl_member(c->self, c->ocx ? c->ocx->bound_prop() : "", VBM_LET, 1, &r, &found);
            break; }
        default:
            ok = true;
            fmt = false;                        /* a picture: the bytes of the field */
        }
        if (vb_errpend) vb_err_clear_pending();
    }
    if (ok && !fmt) {
        bool null = v.vt == VT_NULL || v.vt == VT_EMPTY;
        switch (c->kind) {
        case K_TEXT: case K_COMBO: case K_LIST:
            ctl_let(c, "text", null ? vstr(QString()) : vb_v_s(vb_v2s(v)));
            break;
        case K_LABEL:
            ctl_let(c, "caption", null ? vstr(QString()) : vb_v_s(vb_v2s(v)));
            break;
        case K_CHECK:                           /* Null: Unchecked (DAO), Grayed (ADO; VB6 oracle f_ado) */
            ctl_let(c, "value", vb_v_i(null ? (ado ? 2 : 0) : vb_v2bool(v) ? 1 : 0));
            break;
        case K_IMAGE: case K_PICTURE: {
            QImage img;
            if (!null && (v.vt & VT_ARRAY)) {
                vbarr *a = vb_v_arrptr(v);
                if (a && a->data) img.loadFromData((const uchar *)a->data, a->n);
            }
            if (c->picture) { vb_orelease(c->picture); c->picture = nullptr; }
            vbvar pv = img.isNull() ? vb_v_obj(nullptr) : vb_v_obj(new_pic(img, 1));
            int found;
            ctl_member(c->self, "picture", VBM_SET, 1, &pv, &found);
            break; }
        case K_OCX: {
            int found;
            ctl_member(c->self, c->ocx ? c->ocx->bound_prop() : "", VBM_LET, 1, &v, &found);   /* DataList: BoundText */
            break; }
        default:
            break;
        }
        if (vb_errpend) vb_err_clear_pending();
    }
    if (!keep) c->datachanged = false;
    data_filling = was;
    vb_release(m);
}

/* the bound controls show the current record (or nothing) */
static void data_fill(Ctl *dc)
{
    DataCtl *d = dc->data;
    FormState *st = dc->form;
    if (!st) return;
    bool rec = data_has_record(d);
    for (vbobj co : st->ctls) {
        Ctl *c = C(co);
        if (!c || c->datafield.isEmpty() || data_source_of(c) != dc) continue;
        fill_ctl(c, d->rs, rec);
    }
}

/* the ADO Recordset a bound control shows: its own (Set DataSource = rs) or the one of its ADO Data control */
static bool is_adodc(Ctl *c);
static vbobj adodc_rs(Ctl *c);
static Ctl *named_source(Ctl *c);
static vbobj bound_ado_rs(Ctl *b)
{
    if (b->dsrs) return b->dsrs;
    Ctl *d = b->dsobj ? nullptr : named_source(b);
    return d && is_adodc(d) ? adodc_rs(d) : nullptr;
}
/* a bound control wrote its field: the other controls on that field that the user did not change show it (VB6
 * oracle f_dfmt5: a DataCombo on the field of a DataList, its Format event) */
static void same_field(Ctl *c, vbobj rs)
{
    if (!c->form) return;
    bool rec = rs_has_record(rs);
    std::vector<vbobj> list = c->form->ctls;
    for (vbobj co : list) {
        Ctl *b = C(co);
        if (!b || b == c || b->datachanged || b->datafield.compare(c->datafield, Qt::CaseInsensitive) ||
            (b->kind == K_OCX && b->ocx && b->ocx->own_binding()) || bound_ado_rs(b) != rs) continue;
        fill_ctl(b, rs, rec, true);
    }
}

/* the value of a changed bound control into the current record of rs; -1 on an error (pending), -2 when its
 * DataFormat does not take the text (VB6 oracle f_dfmt: the move is canceled, bound_cancel). ADO: the text goes
 * through the DataFormat (UnFormat), then into the type of the field with the regional settings (vb_db_strto;
 * VB6 read "12,5" without a DataFormat as 125: not here) */
static int store_ctl(Ctl *c, vbobj rs, bool ado = false)
{
    vb_i32 m = vb_mark();
    vbvar v;
    if (ado && c->kind != K_IMAGE && c->kind != K_PICTURE) {
        switch (c->kind) {
        case K_CHECK: { int x = vb_v2i(ctl_get(c, "value")); v = x == 2 ? vb_v_null() : vb_v_bool(x ? -1 : 0); break; }
        case K_LABEL: v = ctl_get(c, "caption"); break;
        case K_OCX: v = ctl_get(c, c->ocx ? c->ocx->bound_prop() : ""); break;
        default: v = ctl_get(c, "text");
        }
        bool fmt = c->dfmt_on && c->dfmt;
        if (fmt) {
            vbvar u;
            if (vb_dfmt_unformat(c->dfmt, v, c->self, &u)) { vb_release(m); return vb_errpend ? -1 : -2; }
            v = u;
        }
        vb_i32 save = vb_errpend;
        vbvar key = vstr(c->datafield);
        vbvar f = vb_late_o(rs, "fields", VBM_GET, 1, &key);
        int t = vb_errpend && !save ? -1 : vb_v2i(vb_late(f, "type", VBM_GET, 0, nullptr));
        if (vb_errpend && !save) { vb_err_clear_pending(); t = -1; }
        vbvar w;
        if (t >= 0 && !vb_db_strto(v, t, &w)) v = w;
        else if (t >= 0 && fmt) { vb_release(m); return -2; }   /* "abc" for a number through a format */
        vbvar a[2] = { vstr(c->datafield), v };
        vb_late_o(rs, "fields", VBM_LET, 2, a);
        vb_release(m);
        if (vb_errpend) return -1;
        c->datachanged = false;
        same_field(c, rs);
        return 0;
    }
    switch (c->kind) {
    case K_CHECK: { int x = vb_v2i(ctl_get(c, "value")); v = x == 2 ? vb_v_null() : vb_v_bool(x ? -1 : 0); break; }
    case K_LABEL: v = ctl_get(c, "caption"); break;
    case K_IMAGE: case K_PICTURE: vb_release(m); return 0;
    case K_OCX: v = ctl_get(c, c->ocx ? c->ocx->bound_prop() : ""); break;
    default: {
        v = ctl_get(c, "text");
        if (v.vt == VT_BSTR && v.u.s.len == 0) {   /* an empty text box: "" in a text field, Null in the others */
            vb_i32 save = vb_errpend;
            vbvar key = vstr(c->datafield);
            vbvar f = vb_late_o(rs, "fields", VBM_GET, 1, &key);
            int t = vb_errpend && !save ? 0 : vb_v2i(vb_late(f, "type", VBM_GET, 0, nullptr));
            if (vb_errpend && !save) vb_err_clear_pending();
            /* DAO dbText 10, dbMemo 12; ADO adVarWChar 202, adLongVarWChar 203, adVarChar 200 ... */
            if (t != 10 && t != 12 && !(t >= 200 && t <= 203) && t != 129 && t != 130) v = vb_v_null();
        }
    }
    }
    vbvar a[2] = { vstr(c->datafield), v };
    vb_late_o(rs, "fields", VBM_LET, 2, a);
    vb_release(m);
    if (vb_errpend) return -1;
    c->datachanged = false;
    return 0;
}

/* the values of the bound controls into the current record (Edit / AddNew in progress) */
static void data_store(Ctl *dc)
{
    DataCtl *d = dc->data;
    for (vbobj co : dc->form->ctls) {
        Ctl *c = C(co);
        if (!c || c->datafield.isEmpty() || data_source_of(c) != dc || !c->datachanged) continue;
        if (store_ctl(c, d->rs)) return;
    }
}

static bool data_changed(Ctl *dc)
{
    for (vbobj co : dc->form->ctls) {
        Ctl *c = C(co);
        if (c && !c->datafield.isEmpty() && data_source_of(c) == dc && c->datachanged) return true;
    }
    return false;
}

/* the changed bound values saved: Edit (unless an AddNew / Edit is on), the values, Update */
static void data_save(Ctl *dc, bool update)
{
    DataCtl *d = dc->data;
    int em = vb_v2i(vb_late_o(d->rs, "editmode", VBM_GET, 0, nullptr));
    if (vb_errpend) return;
    d->in_validate = true;
    if (em == 0) vb_late_o(d->rs, "edit", VBM_METHOD, 0, nullptr);
    if (!vb_errpend) data_store(dc);
    if (!vb_errpend && update) vb_late_o(d->rs, "update", VBM_METHOD, 0, nullptr);
    d->in_validate = false;
    if (!vb_errpend) for (vbobj co : dc->form->ctls) { Ctl *c = C(co); if (c && data_source_of(c) == dc) c->datachanged = false; }
}

/* the controls of .ocx files on the form of a Data control that follow it themselves (DBGrid) */
static void data_ocx_before(Ctl *dc, int action, int *cancel)
{
    if (!dc->form) return;
    std::vector<vbobj> list = dc->form->ctls;
    for (vbobj co : list) {
        Ctl *b = C(co);
        if (b && b != dc && b->kind == K_OCX && b->ocx) b->ocx->data_before(dc->self, action, cancel);
        if (*cancel || vb_errpend) return;
    }
}
static void data_ocx_after(Ctl *dc, int phase)
{
    if (!dc->form) return;
    std::vector<vbobj> list = dc->form->ctls;
    for (vbobj co : list) { Ctl *b = C(co); if (b && b != dc && b->kind == K_OCX && b->ocx) b->ocx->data_after(dc->self, phase); }
}

/* the recordset moves (what 1: before, Validate; what 2: after, Reposition; what 3: its rows changed, no move) */
static void data_watch(void *ctx, int what, int action, int *cancel)
{
    vbobj self = (vbobj)ctx;
    Ctl *dc = C(self);
    if (!dc || !dc->form || !dc->data) return;
    DataCtl *d = dc->data;
    if (d->in_validate) return;
    if (what == 3) { data_ocx_after(dc, 3); return; }
    if (what == 1) {
        data_ocx_before(dc, action, cancel);    /* a DBGrid saves the row it has changed (BeforeUpdate ...) */
        if (*cancel || vb_errpend) return;
        bool changed = data_changed(dc);
        vbvar a[2] = { vb_v_i((vb_i16)action), vb_v_i(changed ? -1 : 0) };
        vb_i32 m = vb_mark();
        d->in_validate = true;
        fire_ctl(dc, "validate", 2, a);
        d->in_validate = false;
        int act = vb_v2i(a[0]);
        bool save = vb_v2i(a[1]) != 0;
        vb_release(m);
        if (act == 0) { *cancel = 1; return; }
        if (save && changed) data_save(dc, action != 6);
        else if (!save) for (vbobj co : dc->form->ctls) { Ctl *c = C(co); if (c && data_source_of(c) == dc) c->datachanged = false; }
        return;
    }
    data_fill(dc);
    data_ocx_after(dc, 1);
    fire_ctl(dc, "reposition");
    data_ocx_after(dc, 2);
}

static void data_close(DataCtl *d)
{
    if (d->rs) {
        if (vb_dbhooks.rs_watch) vb_dbhooks.rs_watch(d->rs, nullptr, nullptr);
        vb_orelease(d->rs); d->rs = nullptr;
    }
    if (d->db) { vb_orelease(d->db); d->db = nullptr; }
}

/* Refresh: the database and the recordset (again); 0 or the error (pending) */
static int data_refresh(Ctl *dc, bool initial = false)
{
    DataCtl *d = dc->data;
    if (d->rs && !d->in_validate) {             /* the open recordset closes: Validate (vbDataActionClose) */
        int cancel = 0;
        data_watch(dc->self, 1, 10, &cancel);
        if (vb_errpend) return -1;
        if (cancel) return 0;
    }
    if (!vb_dbhooks.engine) { vb_raise_full(3343, vb_sfromutf8("DAO.Workspace"), vb_sfromutf8("The database runtime is not in this program.")); return -1; }
    vbobj eng = vb_dbhooks.engine();
    vbvar zero = vb_v_l(0);
    vbobj ws = vb_v2obj(vb_late_o(eng, "workspaces", VBM_GET, 1, &zero));
    if (vb_errpend) return -1;
    vbobj olddb = d->db;
    if (d->rs) { if (vb_dbhooks.rs_watch) vb_dbhooks.rs_watch(d->rs, nullptr, nullptr); vb_orelease(d->rs); d->rs = nullptr; }
    QString name = d->dbname;
    vbvar a[4] = { vstr(name), vb_v_bool(d->exclusive ? -1 : 0), vb_v_bool(d->readonly ? -1 : 0), vstr(d->connect == "Access" ? QString() : d->connect) };
    bool same = olddb && !d->dbname.isEmpty() && data_qs(vb_late_o(olddb, "name", VBM_GET, 0, nullptr)) == name;
    if (vb_errpend) { vb_err_clear_pending(); same = false; }
    if (!same) {
        if (olddb) { vb_orelease(olddb); d->db = nullptr; }
        vbvar dv = vb_late_o(ws, "opendatabase", VBM_METHOD, 4, a);
        if (vb_errpend) return -1;
        d->db = vb_v2obj(dv);
        vb_oaddref(d->db);
    }
    if (d->source.isEmpty()) return 0;
    int type = d->rstype == 0 ? 1 : d->rstype == 2 ? 4 : 2;
    vbvar ra[3] = { vstr(d->source), vb_v_l(type), vb_v_l(d->options) };
    vbvar rv = vb_late_o(d->db, "openrecordset", VBM_METHOD, 3, ra);
    if (vb_errpend) return -1;
    d->rs = vb_v2obj(rv);
    vb_oaddref(d->rs);
    if (vb_dbhooks.rs_watch) vb_dbhooks.rs_watch(d->rs, data_watch, dc->self);
    {                                           /* the grids bound to it read the new recordset */
        std::vector<vbobj> list = dc->form->ctls;
        for (vbobj co : list) { Ctl *b = C(co); if (b && b != dc && b->kind == K_OCX && b->ocx) b->ocx->source_reset(dc->self); }
    }
    if (initial) {                              /* the form loads: Reposition, then the bound controls (VB6 oracle f_data) */
        fire_ctl(dc, "reposition");
        data_fill(dc);
        data_ocx_after(dc, 1);
        data_ocx_after(dc, 2);
    } else {
        data_fill(dc);
        data_ocx_after(dc, 1);
        fire_ctl(dc, "reposition");
        data_ocx_after(dc, 2);
    }
    return 0;
}

/* an error of the Data control itself (not of the program's code): its Error event, then the message */
static void data_error(Ctl *dc)
{
    vb_i32 num = vb_err_number_pending();
    QString msg = QString::fromUtf8(vb_err_text_pending());
    vb_err_clear_pending();
    vbvar a[2] = { vb_v_i((vb_i16)num), vb_v_i(1) };
    vb_i32 m = vb_mark();
    fire_ctl(dc, "error", 2, a);
    int resp = vb_v2i(a[1]);
    vb_release(m);
    if (resp) vb_msgbox(vb_sfromutf8(msg.toUtf8().constData()), 16, vb_app_exename());
}

/* the form has loaded: each Data control with a database and a RecordSource opens them */
static void adodc_after_load(FormState *st);
static void data_after_load(FormState *st)
{
    std::vector<vbobj> list = st->ctls;
    for (vbobj co : list) {
        Ctl *c = C(co);
        if (!c || c->kind != K_DATA || !c->data || !c->form) continue;
        if (c->data->dbname.isEmpty() || c->data->rs) continue;
        if (data_refresh(c, true)) data_error(c);
    }
}

static void data_unload(Ctl *c)
{
    if (!c->data) return;
    data_close(c->data);
}

/* the buttons: |< < > >| and the caption between them */
static void data_button(vbobj self, int which)
{
    Ctl *dc = C(self);
    if (!dc || !dc->form || !dc->data || !dc->data->rs) return;
    DataCtl *d = dc->data;
    static const char *const M[] = { "movefirst", "moveprevious", "movenext", "movelast" };
    vb_late_o(d->rs, M[which], VBM_METHOD, 0, nullptr);
    if (!vb_errpend && (which == 1 || which == 2)) {
        bool eof = vb_v2bool(vb_late_o(d->rs, "eof", VBM_GET, 0, nullptr)) != 0;
        bool bof = vb_v2bool(vb_late_o(d->rs, "bof", VBM_GET, 0, nullptr)) != 0;
        if (which == 2 && eof) {
            if (d->eofaction == 0) vb_late_o(d->rs, "movelast", VBM_METHOD, 0, nullptr);
            else if (d->eofaction == 2) vb_late_o(d->rs, "addnew", VBM_METHOD, 0, nullptr);
        }
        if (which == 1 && bof && d->bofaction == 0) vb_late_o(d->rs, "movefirst", VBM_METHOD, 0, nullptr);
    }
    if (vb_errpend) data_error(dc);
}

static QWidget *data_make(Ctl *c, QWidget *parent)
{
    auto *w = new QFrame(parent);
    w->setFrameShape(QFrame::Box);
    auto *lay = new QHBoxLayout(w);
    lay->setContentsMargins(1, 1, 1, 1);
    lay->setSpacing(0);
    vbobj self = c->self;
    static const char *const T[] = { "|◀", "◀", "▶", "▶|" };
    QToolButton *b[4];
    for (int i = 0; i < 4; i++) {
        b[i] = new QToolButton(w);
        b[i]->setText(QString::fromUtf8(T[i]));
        b[i]->setFocusPolicy(Qt::NoFocus);
        b[i]->setAutoRaise(false);
        QObject::connect(b[i], &QToolButton::clicked, [self, i]() { data_button(self, i); });
    }
    auto *l = new QLabel(c->caption, w);
    l->setAlignment(Qt::AlignVCenter | Qt::AlignLeft);
    l->setContentsMargins(4, 0, 4, 0);
    c->data->label = l;
    lay->addWidget(b[0]); lay->addWidget(b[1]);
    lay->addWidget(l, 1);
    lay->addWidget(b[2]); lay->addWidget(b[3]);
    return w;
}

static void data_load(Ctl *c, const char *const *p)
{
    DataCtl *d = c->data = new DataCtl;
    QString v;
    if (prop(p, "DatabaseName", &v)) d->dbname = v;
    if (prop(p, "RecordSource", &v)) d->source = v;
    if (prop(p, "Connect", &v)) d->connect = v;
    d->rstype = (int)propn(p, "RecordsetType", 1);
    d->options = (int)propn(p, "Options", 0);
    d->eofaction = (int)propn(p, "EOFAction", 0);
    d->bofaction = (int)propn(p, "BOFAction", 0);
    d->deftype = (int)propn(p, "DefaultType", 2);
    d->defcursor = (int)propn(p, "DefaultCursorType", 0);
    d->exclusive = propn(p, "Exclusive", 0) != 0;
    d->readonly = propn(p, "ReadOnly", 0) != 0;
}

/* the members of a Data control */
static vbvar data_member(Ctl *c, const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)argc; (void)argv;
    DataCtl *d = c->data;
    *found = 1;
    if (!d) { *found = 0; return vb_v_empty(); }
    QString *sp = n == "databasename" ? &d->dbname : n == "recordsource" ? &d->source : n == "connect" ? &d->connect : nullptr;
    if (sp) {
        if (set) { *sp = data_qs(val); return vb_v_empty(); }
        return vstr(*sp);
    }
    int *ip = n == "recordsettype" ? &d->rstype : n == "options" ? &d->options : n == "eofaction" ? &d->eofaction :
              n == "bofaction" ? &d->bofaction : n == "defaulttype" ? &d->deftype : n == "defaultcursortype" ? &d->defcursor : nullptr;
    if (ip) {
        if (set) { *ip = vb_v2i(val); return vb_v_empty(); }
        return vb_v_i((vb_i16)*ip);
    }
    bool *bp = n == "exclusive" ? &d->exclusive : n == "readonly" ? &d->readonly : nullptr;
    if (bp) {
        if (set) { *bp = vb_v2bool(val) != 0; return vb_v_empty(); }
        return vb_v_bool(*bp ? -1 : 0);
    }
    if (n == "caption") {
        if (set) { c->caption = data_qs(val); if (d->label) d->label->setText(c->caption); return vb_v_empty(); }
        return vstr(c->caption);
    }
    if (n == "database") {
        if (set) { vbobj o = vb_v2obj(val); if (d->db) vb_orelease(d->db); d->db = o; if (o) vb_oaddref(o); return vb_v_empty(); }
        return vb_vtmp(vb_v_obj(d->db));
    }
    if ((n == "recordset" || n == "database") && argc > (set ? 1 : 0)) {
        /* Data1.Recordset!Name, Data1.Recordset("Name") = v: the default member of the recordset */
        vbobj o = n == "recordset" ? d->rs : d->db;
        if (!o) { vb_raise(91); return vb_v_empty(); }
        return vb_late_o(o, "", kind, argc, argv);
    }
    if (n == "recordset") {
        if (set) {
            vbobj o = vb_v2obj(val);
            if (vb_errpend) return vb_v_empty();
            if (d->rs) { if (vb_dbhooks.rs_watch) vb_dbhooks.rs_watch(d->rs, nullptr, nullptr); vb_orelease(d->rs); }
            d->rs = o;
            if (o) {
                vb_oaddref(o);
                if (vb_dbhooks.rs_watch) vb_dbhooks.rs_watch(o, data_watch, c->self);
                std::vector<vbobj> list = c->form ? c->form->ctls : std::vector<vbobj>();
                for (vbobj co : list) { Ctl *b = C(co); if (b && b != c && b->kind == K_OCX && b->ocx) b->ocx->source_reset(c->self); }
                data_fill(c);
                data_ocx_after(c, 1);
                fire_ctl(c, "reposition");
                data_ocx_after(c, 2);
            }
            return vb_v_empty();
        }
        return vb_vtmp(vb_v_obj(d->rs));       /* Nothing before it opens (VB6) */
    }
    if (n == "refresh") { data_refresh(c); return vb_v_empty(); }
    if (n == "updatecontrols") { if (d->rs) data_fill(c); return vb_v_empty(); }
    if (n == "updaterecord") {
        if (!d->rs) { vb_raise(91); return vb_v_empty(); }
        if (!data_has_record(d)) { vb_raise_full(3021, vb_sfromutf8("DAO.Recordset"), vb_sfromutf8("No current record.")); return vb_v_empty(); }
        bool any = data_changed(c);
        if (any) data_save(c, true);
        return vb_v_empty();
    }
    if (n == "align") { if (set) { c->align = vb_v2i(val); if (c->form) align_layout(c->form); return vb_v_empty(); } return vb_v_i((vb_i16)c->align); }
    if (n == "negotiate") return vb_v_bool(0);
    if (n == "appearance") return vb_v_i(1);
    (void)kind;
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- ADO: the ADO Data control, the data sources */
static bool is_ctl_obj(vbobj o) { o = o ? vb_oid(o) : nullptr; return o && o->ext && o->cls->clear == ctl_clear; }
static bool is_adodc(Ctl *c) { return c && c->kind == K_OCX && (c->type == "Adodc" || c->type == "MSRDC"); }   /* the RemoteData control too */
static vbobj adodc_rs(Ctl *c);

/* the control a bound control names (DataSource = "Adodc1"): a Data control or an ADO Data control */
static Ctl *named_source(Ctl *c)
{
    if (!c->form || c->datasource.isEmpty()) return nullptr;
    auto it = c->form->byname.find(lower(c->datasource.toUtf8().constData()));
    if (it == c->form->byname.end() || it->second->cls == &CLS_ARRAY) return nullptr;
    Ctl *d = C(it->second);
    return d && (d->kind == K_DATA || is_adodc(d)) ? d : nullptr;
}

/* the ADO Recordset of a data source object: a Recordset, an ADO Data control, an object of a class
 * that is a data source (DataSourceBehavior: its GetDataMember gives it) */
static vbobj source_rs(vbobj src, const QString &member)
{
    if (src && is_ctl_obj(src) && C(src)->kind == K_DATA) return C(src)->data ? C(src)->data->rs : nullptr;   /* DAO (DBList) */
    if (!src || !vb_dbhooks.is_ado_rs) return nullptr;
    if (vb_dbhooks.is_ado_rs(src)) return vb_oid(src);
    if (is_ctl_obj(src)) {
        Ctl *d = C(src);
        if (is_adodc(d)) return adodc_rs(d);
        if (d->kind != K_USER) return nullptr;
    }
    vb_i32 save = vb_errpend;
    vbvar a[2] = { vstr(member), vb_v_obj(nullptr) };
    vbobj inner = is_ctl_obj(src) && C(src)->uc ? C(src)->uc : src;
    vb_late_o(inner, "lb_getdatamember", VBM_METHOD, 2, a);
    if (vb_errpend && !save) { vb_err_clear_pending(); return nullptr; }
    vbobj o = a[1].vt == VT_DISPATCH ? a[1].u.o : nullptr;
    return o && vb_dbhooks.is_ado_rs(o) ? vb_oid(o) : nullptr;
}

/* a DataFormat did not take the text of a bound control (store_ctl -2): the controls bound to rs show the record
 * again, the changed ones stay changed, and the move is canceled (VB6 oracle f_dfmt) */
static void bound_cancel(Ctl *c, vbobj rs)
{
    if (c->form) {
        bool rec = rs_has_record(rs);
        std::vector<vbobj> list = c->form->ctls;
        for (vbobj co : list) {
            Ctl *b = C(co);
            if (!b || b->datafield.isEmpty() || (b->kind == K_OCX && b->ocx && b->ocx->own_binding())) continue;
            Ctl *d = b->dsobj ? nullptr : named_source(b);
            if (b->dsrs == rs || (d && is_adodc(d) && adodc_rs(d) == rs)) fill_ctl(b, rs, rec, true, true);
        }
    }
    vb_raise_full(-2147217842, vs(QString("ADODB.Recordset")), vs(QString("Operation was canceled.")));
}

/* a control bound to an object of its own (Set Text1.DataSource = rs) listens to its Recordset */
static void ctl_ado_listen(void *ctx, vbobj rs, const char *ev, int argc, vbvar *argv)
{
    (void)argc; (void)argv;
    Ctl *c = C((vbobj)ctx);
    if (!c || c->datafield.isEmpty()) return;
    if (!strcmp(ev, "lb_save")) { if (c->datachanged && store_ctl(c, rs, true) == -2) bound_cancel(c, rs); }
    else if (!strcmp(ev, "lb_fill")) fill_ctl(c, rs, rs_has_record(rs), true);
}
static void bound_listen(Ctl *c, vbobj rs)
{
    if (c->dsrs == rs) return;
    if (c->dsrs) { if (vb_dbhooks.ado_listen) vb_dbhooks.ado_listen(c->dsrs, ctl_ado_listen, c->self, 0); vb_orelease(c->dsrs); }
    c->dsrs = rs;
    if (rs) { vb_oaddref(rs); if (vb_dbhooks.ado_listen) vb_dbhooks.ado_listen(rs, ctl_ado_listen, c->self, 1); }
}
static void bound_release(Ctl *c)
{
    bound_listen(c, nullptr);
    if (c->dsobj) { vb_orelease(c->dsobj); c->dsobj = nullptr; }
}

/* the ADO Data control */
struct OcxAdodc : Ocx {
    QString cs, source, user, pwd, caption;
    int cmdtype = 8, cursorloc = 3, cursortype = 3, locktype = 3, mode = 0, cache = 50, maxrec = 0;
    int bofaction = 0, eofaction = 0, cntimeout = 15, cmdtimeout = 30, orientation = 0;
    long back = 0x80000005L, fore = 0x80000008L;
    vbobj cn = nullptr, rs = nullptr;           /* counted */
    QPointer<QLabel> label;
    bool own_cn = true;
    ~OcxAdodc() override { close(); }
    bool data_aware() override { return true; }
    bool own_binding() override { return true; }
    QWidget *make(QWidget *parent) override
    {
        auto *w = new QFrame(parent);
        w->setFrameShape(QFrame::Box);
        auto *lay = new QHBoxLayout(w);
        lay->setContentsMargins(1, 1, 1, 1);
        lay->setSpacing(0);
        vbobj self = c->self;
        static const char *const T[] = { "|◀", "◀", "▶", "▶|" };
        QToolButton *b[4];
        for (int i = 0; i < 4; i++) {
            b[i] = new QToolButton(w);
            b[i]->setText(QString::fromUtf8(T[i]));
            b[i]->setFocusPolicy(Qt::NoFocus);
            QObject::connect(b[i], &QToolButton::clicked, [self, i]() {
                Ctl *x = C(self);
                if (auto *a = x ? dynamic_cast<OcxAdodc *>(x->ocx) : nullptr) a->button(i);
            });
        }
        label = new QLabel(caption, w);
        label->setAlignment(Qt::AlignVCenter | Qt::AlignLeft);
        label->setContentsMargins(4, 0, 4, 0);
        lay->addWidget(b[0]); lay->addWidget(b[1]);
        lay->addWidget(label, 1);
        lay->addWidget(b[2]); lay->addWidget(b[3]);
        return w;
    }
    void load(const char *const *p) override
    {
        QString v;
        if (prop(p, "OLEDBString", &v) && !v.isEmpty()) cs = v;
        else if (prop(p, "Connect", &v) && !v.isEmpty()) cs = v;
        if (prop(p, "RecordSource", &v)) source = v;
        if (prop(p, "UserName", &v)) user = v;
        if (prop(p, "Password", &v)) pwd = v;
        if (prop(p, "Caption", &v)) caption = v;
        cmdtype = (int)propn(p, "CommandType", 8);
        cursorloc = (int)propn(p, "CursorLocation", 3);
        cursortype = (int)propn(p, "CursorType", 3);
        locktype = (int)propn(p, "LockType", 3);
        mode = (int)propn(p, "ConnectMode", 0);
        cache = (int)propn(p, "CacheSize", 50);
        maxrec = (int)propn(p, "MaxRecords", 0);
        bofaction = (int)propn(p, "BOFAction", 0);
        eofaction = (int)propn(p, "EOFAction", 0);
        cntimeout = (int)propn(p, "ConnectionTimeout", 15);
        cmdtimeout = (int)propn(p, "CommandTimeout", 30);
        orientation = (int)propn(p, "Orientation", 0);
        back = propc(p, "BackColor", 0x80000005L);
        fore = propc(p, "ForeColor", 0x80000008L);
        if (label) label->setText(caption);
    }
    void listen(bool on)
    {
        if (rs && vb_dbhooks.ado_listen) vb_dbhooks.ado_listen(rs, adodc_listen, c->self, on ? 1 : 0);
    }
    void close()
    {
        if (rs) {
            listen(false);
            vb_i32 save = vb_errpend;
            vb_late_o(rs, "close", VBM_METHOD, 0, nullptr);
            if (vb_errpend && !save) vb_err_clear_pending();
            vb_orelease(rs); rs = nullptr;
        }
        if (cn) {
            if (own_cn) {
                vb_i32 save = vb_errpend;
                vb_late_o(cn, "close", VBM_METHOD, 0, nullptr);
                if (vb_errpend && !save) vb_err_clear_pending();
            }
            vb_orelease(cn); cn = nullptr;
        }
    }
    void fill_bound()
    {
        if (!c->form) return;
        bool rec = rs_has_record(rs);
        for (vbobj co : c->form->ctls) {
            Ctl *b = C(co);
            if (b && !b->datafield.isEmpty() && !b->dsobj && named_source(b) == c && !(b->kind == K_OCX && b->ocx && b->ocx->own_binding()))
                fill_ctl(b, rs, rec, true);
        }
    }
    void store_bound()
    {
        if (!c->form) return;
        for (vbobj co : c->form->ctls) {
            Ctl *b = C(co);
            if (b && !b->datafield.isEmpty() && !b->dsobj && named_source(b) == c && b->datachanged &&
                !(b->kind == K_OCX && b->ocx && b->ocx->own_binding())) {
                int r = store_ctl(b, rs, true);
                if (r == -2) bound_cancel(b, rs);
                if (r) return;
            }
        }
    }
    static void adodc_listen(void *ctx, vbobj rs, const char *ev, int argc, vbvar *argv)
    {
        (void)rs;
        Ctl *x = C((vbobj)ctx);
        auto *a = x ? dynamic_cast<OcxAdodc *>(x->ocx) : nullptr;
        if (!a) return;
        if (!strcmp(ev, "lb_save")) a->store_bound();
        else if (!strcmp(ev, "lb_fill")) a->fill_bound();
        else fire_ctl(x, ev, argc, argv);       /* the events of the recordset are the control's */
    }
    /* Refresh: the connection and the recordset again; 0 or the error (pending) */
    int refresh()
    {
        close();
        if (!vb_dbhooks.createobject) { vb_raise_full(429, vs(QString("ADODC")), vs(QString("ActiveX component can't create object"))); return -1; }
        vbobj ncn = vb_dbhooks.createobject("adodb.connection");
        vbobj nrs = vb_dbhooks.createobject("adodb.recordset");
        if (!ncn || !nrs) { vb_raise(429); return -1; }
        vb_oaddref(ncn); vb_oaddref(nrs);
        cn = ncn; own_cn = true;
        vbvar v = vb_v_l(cursorloc);
        vb_late_o(cn, "cursorlocation", VBM_LET, 1, &v);
        if (mode) { v = vb_v_l(mode); vb_late_o(cn, "mode", VBM_LET, 1, &v); }
        vbvar oa[3] = { vstr(cs), vstr(user), vstr(pwd) };
        vb_late_o(cn, "open", VBM_METHOD, 3, oa);
        if (vb_errpend) { vb_orelease(nrs); return -1; }
        rs = nrs;
        v = vb_v_l(cursorloc); vb_late_o(rs, "cursorlocation", VBM_LET, 1, &v);
        if (maxrec) { v = vb_v_l(maxrec); vb_late_o(rs, "maxrecords", VBM_LET, 1, &v); }
        listen(true);
        vbvar ra[5] = { vstr(source), vb_v_obj(cn), vb_v_l(cursortype), vb_v_l(locktype), vb_v_l(cmdtype) };
        vb_late_o(rs, "open", VBM_METHOD, 5, ra);
        if (vb_errpend) return -1;
        reset_users();
        return 0;
    }
    void reset_users()                          /* the grids and lists of this control read the new Recordset */
    {
        if (!c->form) return;
        std::vector<vbobj> list = c->form->ctls;
        for (vbobj co : list) {
            Ctl *b = C(co);
            if (b && b != c && b->kind == K_OCX && b->ocx) b->ocx->source_reset(c->self);
        }
    }
    void error_event()
    {
        vb_i32 num = vb_err_number_pending();
        QString msg = QString::fromUtf8(vb_err_text_pending());
        vb_err_clear_pending();
        /* a provider's error: ErrorNumber 3639, Scode its HRESULT, Source the Caption (VB6 oracle f_ado) */
        vbvar a[7] = { vb_v_l(num < 0 ? 3639 : num), vstr(msg), vb_v_l(num), vstr(caption), vstr(QString()), vb_v_l(0), vb_v_bool(0) };
        vb_i32 m = vb_mark();
        fire_ctl(c, "error", 7, a);
        bool cancel = vb_v2bool(a[6]) != 0;
        vb_release(m);
        if (!cancel) vb_msgbox(vb_sfromutf8(msg.toUtf8().constData()), 16, vb_app_exename());
    }
    void button(int which)
    {
        if (!rs) return;
        static const char *const M[] = { "movefirst", "moveprevious", "movenext", "movelast" };
        vb_late_o(rs, M[which], VBM_METHOD, 0, nullptr);
        if (!vb_errpend && (which == 1 || which == 2)) {
            bool eof = vb_v2bool(vb_late_o(rs, "eof", VBM_GET, 0, nullptr)) != 0;
            bool bof = vb_v2bool(vb_late_o(rs, "bof", VBM_GET, 0, nullptr)) != 0;
            if (which == 2 && eof) {
                if (eofaction == 0) vb_late_o(rs, "movelast", VBM_METHOD, 0, nullptr);
                else if (eofaction == 2) vb_late_o(rs, "addnew", VBM_METHOD, 0, nullptr);
            }
            if (which == 1 && bof && bofaction == 0) vb_late_o(rs, "movefirst", VBM_METHOD, 0, nullptr);
        }
        if (vb_errpend) error_event();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        QString *sp = n == "connectionstring" || n == "oledbstring" ? &cs : n == "recordsource" ? &source :
                      n == "username" ? &user : n == "password" ? &pwd : n == "caption" ? &caption : nullptr;
        if (sp) {
            if (!set) return vstr(*sp);
            if (n == "recordsource" && vb_isobject(val)) { *sp = QString(); return vb_v_empty(); }
            *sp = data_qs(val);
            if (n == "caption" && label) label->setText(caption);
            return vb_v_empty();
        }
        int *ip = n == "commandtype" ? &cmdtype : n == "cursorlocation" ? &cursorloc : n == "cursortype" ? &cursortype :
                  n == "locktype" ? &locktype : n == "mode" ? &mode : n == "cachesize" ? &cache : n == "maxrecords" ? &maxrec :
                  n == "bofaction" ? &bofaction : n == "eofaction" ? &eofaction : n == "connectiontimeout" ? &cntimeout :
                  n == "commandtimeout" ? &cmdtimeout : n == "orientation" ? &orientation : nullptr;
        if (ip) { if (!set) return vb_v_l(*ip); *ip = vb_v2l(val); return vb_v_empty(); }
        if (n == "recordset") {
            if (set) {
                vbobj o = vb_v2obj(val);
                if (vb_errpend) return vb_v_empty();
                listen(false);
                if (rs) vb_orelease(rs);
                rs = o ? vb_oid(o) : nullptr;
                if (rs) { vb_oaddref(rs); listen(true); }
                fill_bound();
                reset_users();
                return vb_v_empty();
            }
            if (argc > 0 && rs) return vb_late_o(rs, "", kind, argc, argv);   /* Adodc1.Recordset!Name */
            return vb_vtmp(vb_v_obj(rs));
        }
        if (n == "refresh") {
            if (refresh()) {
                vb_i32 num = vb_err_number_pending();
                QString msg = QString::fromUtf8(vb_err_text_pending());
                QString src = QString::fromUtf8(vb_stoutf8(vb_err_source()));
                error_event();
                vb_raise_full(num, vs(src), vs(msg));
            }
            return vb_v_empty();
        }
        if (n == "backcolor" || n == "forecolor") {
            long *cp = n == "backcolor" ? &back : &fore;
            if (!set) return vb_v_l((vb_i32)*cp);
            *cp = vb_v2l(val);
            return vb_v_empty();
        }
        if (n == "connectstringtype" || n == "isolationlevel" || n == "cursoroptions") return vb_v_l(n == "isolationlevel" ? -1 : n == "connectstringtype" ? 1 : 0);
        if (n == "oledbfile" || n == "datasourcename" || n == "otherattributes") return vstr(QString());
        if (n == "aboutbox" || n == "fireerrorinfo") return vb_v_empty();
        *found = 0;
        return vb_v_empty();
    }
};
static Ocx *new_adodc() { return new OcxAdodc; }
static vbobj rdc_inner(Ctl *c);                 /* ocx_rdc.inc: the rows of a RemoteData control */
static void rdc_after_load(FormState *st);
static vbobj adodc_rs(Ctl *c)
{
    if (c && c->type == "MSRDC") return rdc_inner(c);
    auto *a = c ? dynamic_cast<OcxAdodc *>(c->ocx) : nullptr;
    return a ? a->rs : nullptr;
}

/* the ADO Data controls of a form open when it has loaded (as the Data controls) */
static void adodc_after_load(FormState *st)
{
    std::vector<vbobj> list = st->ctls;
    for (vbobj co : list) {
        Ctl *c = C(co);
        auto *a = c ? dynamic_cast<OcxAdodc *>(c->ocx) : nullptr;
        if (!a || a->rs || a->cs.isEmpty() || a->source.isEmpty()) continue;
        if (a->refresh()) a->error_event();
    }
    rdc_after_load(st);                         /* the RemoteData controls (ocx_rdc.inc) */
}

/* DataField, DataSource, DataChanged of a bound control */
static vbvar bound_member(Ctl *c, const std::string &n, bool set, vbvar val, int *found)
{
    *found = 1;
    if (n == "datafield") {
        if (set) {
            c->datafield = data_qs(val);
            if (c->dsrs) fill_ctl(c, c->dsrs, rs_has_record(c->dsrs), true);
            else if (Ctl *d = named_source(c)) {     /* the control shows the field of the current record */
                if (d->kind == K_DATA && d->data && d->data->rs) fill_ctl(c, d->data->rs, data_has_record(d->data));
                else if (is_adodc(d) && adodc_rs(d)) fill_ctl(c, adodc_rs(d), rs_has_record(adodc_rs(d)), true);
            }
            return vb_v_empty();
        }
        return vstr(c->datafield);
    }
    if (n == "datachanged") { if (set) { c->datachanged = vb_v2bool(val) != 0; return vb_v_empty(); } return vb_v_bool(c->datachanged ? -1 : 0); }
    if (n == "datasource") {
        if (set) {
            vbobj o = vb_v2obj(val);
            if (vb_errpend) return vb_v_empty();
            bound_release(c);
            c->datasource.clear();
            if (o && is_ctl_obj(o) && (C(o)->kind == K_DATA || is_adodc(C(o)))) {
                c->datasource = C(o)->name;
                Ctl *d = C(o);
                if (d->kind == K_DATA && d->data && d->data->rs) fill_ctl(c, d->data->rs, data_has_record(d->data));
                else if (is_adodc(d) && adodc_rs(d)) fill_ctl(c, adodc_rs(d), rs_has_record(adodc_rs(d)), true);
            } else if (o) {
                c->dsobj = vb_oid(o); vb_oaddref(c->dsobj);
                vbobj rs = source_rs(c->dsobj, c->datamember);
                bound_listen(c, rs);
                if (rs && !c->datafield.isEmpty()) fill_ctl(c, rs, rs_has_record(rs), true);
            }
            return vb_v_empty();
        }
        if (c->dsobj) return vb_vtmp(vb_v_obj(c->dsobj));
        Ctl *d = named_source(c);
        return vb_vtmp(vb_v_obj(d ? d->self : nullptr));
    }
    if (n == "dataformat") {
        if (set) {
            vbobj o = vb_v2obj(val);
            if (vb_errpend) return vb_v_empty();
            if (!o) {                           /* VB6: E_INVALIDARG (the DataFormat stays) */
                vb_raise_full(-2147024809, vb_app_exename(),
                              vs(QString("System Error &H80070057 (-2147024809).   The parameter is incorrect.")));
                return vb_v_empty();
            }
            if (!vb_is_dfmt(o)) { vb_raise(13); return vb_v_empty(); }
            o = vb_oid(o);
            vb_oaddref(o);
            if (c->dfmt) vb_orelease(c->dfmt);
            c->dfmt = o;                        /* used from the next record on (VB6 oracle f_dfmt) */
            c->dfmt_on = true;
            return vb_v_empty();
        }
        if (!c->dfmt) c->dfmt = vb_dfmt_new();  /* a control without a DataFormat has one, the binding does not use it */
        return vb_vtmp(vb_v_obj(c->dfmt));
    }
    if (n == "datamember") {                    /* the member of a data source class (the Data Form Wizard's Class
                                                 * binding: DataMember = "Primary", then Set DataSource = obj) */
        if (set) {
            c->datamember = data_qs(val);
            if (c->dsobj) {
                vbobj rs = source_rs(c->dsobj, c->datamember);
                bound_listen(c, rs);
                if (rs && !c->datafield.isEmpty()) fill_ctl(c, rs, rs_has_record(rs), true);
            }
            return vb_v_empty();
        }
        return vstr(c->datamember);
    }
    *found = 0;
    return vb_v_empty();
}
