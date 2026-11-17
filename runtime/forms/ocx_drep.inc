/* ocx_drep.inc - the DataRepeater control of VB6 (MSDATREP.OCX, MSDataRepeaterLib.DataRepeater; phase 7.5): a user
 * control of an ActiveX control project (RepeatedControlName "Proj.Ctl", a component compiled into the program,
 * uctl.inc) repeated for the records of an ADO Recordset (DataSource). The properties of the user control that are
 * data bound (Procedure Attributes: VB_MemberFlags with "bindable", the compiler lists them as "lb_bind:<Name>"
 * members) are bound to fields by RepeaterBindings (PropertyName, DataField, DataFormat, Key).
 *
 * As VB6 does it (VB6 oracles f_drep1 .. f_drep3): two objects of the user control, one that is the current record
 * (RepeatedControl: the user edits it) and one that draws the other rows; a change in the current one
 * (PropertyChanged) goes into the record when the recordset leaves it (DataUpdate, then the record is updated);
 * the rows scroll only as much as the current record needs (CurrentRecord = puts it at the top, ActiveRow = n puts
 * it in row n when that can be, VisibleRecords(i) = scrolls so that record is row i); after a Delete the
 * recordset goes to the next record. ActiveRowChanged, CurrentRecordChanged, VisibleRecordsChanged and
 * RepeatedControlLoaded come after the code that caused them (posted): a move without a scroll gives
 * CurrentRecordChanged [ActiveRowChanged], with a scroll VisibleRecordsChanged and that pair twice, ActiveRowChanged
 * only when the active row is another; a field of a record that changes (not of a new one) and the update of a
 * record give VisibleRecordsChanged, as AddNew, ActiveRow = and VisibleRecords( ) = do. DataUpdate and
 * RepeatedControlUnload come at once.
 * Not as VB6 (its bugs, corrected in the expected outputs): VisibleRecords(VisibleRows + 1) gave a record that is
 * not visible (here 381, as for 0); RepeaterBindings.Add of a property that is not bindable was taken (here 35701);
 * a Delete did not reach the database; Filter = adFilterNone left the control without rows.
 * Included by vbforms.cpp after ocx_data.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

struct OcxDataRepeater;
static std::set<OcxDataRepeater *> &drep_all()
{
    static std::set<OcxDataRepeater *> s;
    return s;
}
static void drep_raise(vb_i32 n, const char *src, const char *msg) { vb_raise_full(n, vb_sfromutf8(src), vb_sfromutf8(msg)); }
static void drep_err(vb_i32 n)
{
    static const std::map<vb_i32, const char *> msg = {
        { 32000, "Unable to create RepeatedControl" }, { 32001, "Invalid ProgID or class object not supported" },
        { 32002, "Invalid bookmark" }, { 32003, "No datasource or datasource is empty" },
        { 32004, "There are no visible rows" },
        { 32005, "The active row cannot be set without affecting the current record position" },
        { 32006, "The row index must be between one and VisibleRows" },
        { 32007, "The bookmark is invalid or the corresponding record is no longer present" },
        { 380, "Invalid property value" }, { 381, "Invalid property array index" } };
    auto it = msg.find(n);
    drep_raise(n, "DataRepeater", it != msg.end() ? it->second : "");
}
static void rbinds_err(vb_i32 n)
{
    static const std::map<vb_i32, const char *> msg = {
        { 35600, "Index out of bounds" }, { 35601, "Element not found" }, { 35602, "Key is not unique in collection" },
        { 35701, "Property is not bindable" }, { 35703, "PropertyName already bound" }, { 35704, "Invalid PropertyName" },
        { 35705, "Invalid DataField" } };
    auto it = msg.find(n);
    drep_raise(n, "RepeaterBindings", it != msg.end() ? it->second : "");
}

/* ---------------------------------------------------------------- RepeaterBinding, RepeaterBindings */
struct RBindData {
    OcxDataRepeater *owner = nullptr;           /* weak; 0 once it left the collection */
    QString prop, field, key;
    vbobj fmt = nullptr;                        /* DataFormat (counted) */
    vbvar tag = vb_v_empty();
    bool changed = false;
    ~RBindData() { if (fmt) vb_orelease(fmt); vb_vfree(&tag); }
};
static vbvar rbind_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static void rbind_clear(vbobj o) { delete (RBindData *)o->ext; o->ext = nullptr; }
static const vb_class CLS_RBIND = { "IRepeaterBinding", (vb_i32)sizeof(vbobjh), nullptr, nullptr, rbind_clear, nullptr, 0,
                                    nullptr, nullptr, nullptr, nullptr, rbind_member, nullptr };
static vbvar rbinds_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vb_i32 rbinds_count(vbobj o);
static vbvar rbinds_item(vbobj o, vb_i32 i);
static const vb_class CLS_RBINDS = { "IRepeaterBindings", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0,
                                     "item", rbinds_count, rbinds_item, nullptr, rbinds_member, nullptr };
static RBindData *RB(vbobj o) { o = o ? vb_oid(o) : nullptr; return o && o->cls == &CLS_RBIND ? (RBindData *)o->ext : nullptr; }

/* the rows of the DataRepeater: a widget that draws them (the current one is the RepeatedControl itself) */
struct DrepView : QWidget {
    vbobj self = nullptr;                       /* the control (weak) */
    explicit DrepView(QWidget *p) : QWidget(p) { setAutoFillBackground(false); }
    void paintEvent(QPaintEvent *) override;
    void mousePressEvent(QMouseEvent *e) override;
    void resizeEvent(QResizeEvent *) override;
};

struct OcxDataRepeater : Ocx {
    QString progid;                             /* RepeatedControlName */
    const vb_class *ucls = nullptr;
    vbobj uc[2] = { nullptr, nullptr };         /* the current record, the one that draws the other rows (counted) */
    std::vector<vbobj> binds;                   /* RepeaterBinding objects (counted) */
    vbobj bindsobj = nullptr;
    vbobj src = nullptr, rs = nullptr;          /* DataSource, its Recordset (counted) */
    QString dmember, srcname, caption;
    int captionstyle = 1, scrollbars = 2, divider = 2, border = 1, appearance = 1;
    bool rowind = true, integral = true;
    long back = 0x8000000FL, fore = 0x80000012L;
    /* where it is: the first visible row, the current row, the rows (with the one being added) - ado_shape rows */
    int top = 0, cur = -1, nrows = -1, want_top = -1;
    bool adding = false, del_pending = false, reset_pending = false, filling = false, storing = false;
    std::set<std::string> dirty;                /* the bound properties the user changed in the current record */
    std::map<int, QPixmap> pix;                 /* the rows drawn (by their row) */
    QPointer<DrepView> view;
    QPointer<QScrollBar> vbar;
    bool syncing_bar = false;
    OcxDataRepeater();
    ~OcxDataRepeater() override;
    bool data_aware() override { return true; }
    bool own_binding() override { return true; }
    QWidget *make(QWidget *parent) override;
    void load(const char *const *p) override;
    void resized() override { layout(); render(); }
    void source_reset(vbobj s) override
    {
        if (!src && !srcname.isEmpty() && is_ctl_obj(s) && !srcname.compare(C(s)->name, Qt::CaseInsensitive)) { src = vb_oid(s); vb_oaddref(src); }
        if (src && vb_oid(src) == vb_oid(s)) bind();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;

    /* ---------------------------------------------------------------- the geometry */
    int rowh() const
    {
        if (!uc[0]) return 0;
        FormState *us = fstate(uc[0]);
        return std::max(1, (int)(propn(us->def->props, "ClientHeight", 600) / TPP));
    }
    int caph() const { return captionstyle != 0 ? 17 : 0; }
    int frame() const { return border ? 2 : 0; }
    int indw() const { return rowind ? 14 : 0; }
    int fit() const                             /* the whole rows that fit in its height */
    {
        if (!uc[0]) return 0;
        int h = (int)(c->height / TPP) - 2 * frame() - caph() - ((scrollbars & 1) ? 16 : 0);
        return std::max(0, h / rowh());
    }
    int count() const { return rs ? std::max(0, nrows) : 0; }
    int vrows() const { return rs && uc[0] ? std::max(0, std::min(fit(), count() - top)) : 0; }
    int active() const
    {
        int r = vrows();
        return rs && uc[0] && cur >= 0 && cur >= top && cur < top + r ? cur - top + 1 : 0;
    }
    QRect row_rect(int i) const                 /* row i (0-based) of the visible ones */
    {
        if (!view) return QRect();
        int x = frame() + indw(), y = frame() + caph() + i * rowh();
        int w = view->width() - x - frame() - ((scrollbars >= 2) ? 16 : 0);
        return QRect(x, y, std::max(0, w), rowh());
    }
    /* the rows of the recordset now: their number (with a new one being added), the current one */
    int shape(int *pcur)
    {
        int c0 = -1;
        int n = rsx_shape(rs, nullptr, &c0);
        if (n < 0) { *pcur = -1; return 0; }
        *pcur = c0;
        return c0 == n ? n + 1 : n;            /* AddNew: the new row is the last */
    }

    /* ---------------------------------------------------------------- the repeated control */
    void make_ucs()
    {
        for (int k = 0; k < 2; k++) {
            vbobj o = vb_onew(ucls);            /* UserControl_Initialize */
            if (!o || vb_errpend) break;
            vb_oaddref(o);
            uc[k] = o;
            uctl_load(o);
        }
        if (!uc[0] || !uc[1]) { release_ucs(false); return; }
        FormState *us = fstate(uc[0]);
        if (us->w && view) { us->w->setParent(view); us->w->hide(); }
        post_event("repeatedcontrolloaded");
        layout(); fill_live(); render();
    }
    void release_ucs(bool event)
    {
        if (!uc[0] && !uc[1]) return;
        if (event && uc[0]) fire_ctl(c, "repeatedcontrolunload");
        dirty.clear(); pix.clear();
        for (int k = 0; k < 2; k++) {
            vbobj o = uc[k];
            if (!o) continue;
            uc[k] = nullptr;
            FormState *us = fstate(o);
            for (vbobj co : us->ctls) { Ctl *ic = C(co); if (ic && ic->timer) ic->timer->stop(); }
            if (us->canvas) { us->canvas->on_paint = nullptr; us->canvas->on_resize = nullptr; }
            if (us->w) us->w->hide();
            vb_i32 m = vb_mark();
            vb_orelease(o);                     /* UserControl_Terminate */
            vb_release(m);
        }
        if (view) view->update();
    }
    /* RepeatedControlName = "Proj.Ctl": the user control of that ProgID (32001: none) */
    void set_progid(const QString &name, bool raise_err)
    {
        if (name.isEmpty()) { release_ucs(true); progid.clear(); ucls = nullptr; return; }
        const vb_class *k = uctl_class(name.toUtf8().constData());
        if (!k) { if (raise_err) drep_err(32001); return; }
        release_ucs(true);
        progid = name; ucls = k;
        make_ucs();
        if (!uc[0] && raise_err && !vb_errpend) drep_err(32000);
    }
    /* the properties of the user control that can be bound, as its code names them */
    QStringList bindable() const
    {
        QStringList out;
        if (!ucls) return out;
        for (vb_i32 i = 0; i < ucls->nmembers; i++)
            if (!strncmp(ucls->members[i].name, "lb_bind:", 8)) out << QString::fromUtf8(ucls->members[i].name + 8);
        return out;
    }
    bool has_member(const QString &p) const
    {
        if (!ucls) return false;
        std::string l = lower(p.toUtf8().constData());
        for (vb_i32 i = 0; i < ucls->nmembers; i++) if (ucls->members[i].kind && l == ucls->members[i].name) return true;
        return false;
    }
    /* a property of the user control from field col of row r (Null: Empty) */
    void put(vbobj o, const QString &prop, int r, int col)
    {
        if (col < 0 || r < 0) return;
        vb_i32 m = vb_mark();
        vb_i32 save = vb_errpend;
        vbvar v = rsx_cell(rs, r, col);
        if (v.vt == VT_NULL) v = vb_v_empty();
        QByteArray pn = prop.toLower().toUtf8();
        vb_late_o(o, pn.constData(), VBM_LET, 1, &v);
        if (vb_errpend && !save) vb_err_clear_pending();
        vb_release(m);
    }
    /* the current record into the RepeatedControl */
    void fill_live(const QString &only = QString())
    {
        if (!uc[0] || !rs || cur < 0 || storing) return;
        filling = true;
        for (vbobj b : binds) {
            RBindData *d = RB(b);
            if (!d || (!only.isEmpty() && d->prop.compare(only, Qt::CaseInsensitive))) continue;
            put(uc[0], d->prop, cur, rs_col_index(rs, d->field));
        }
        filling = false;
        if (only.isEmpty()) dirty.clear();
    }
    /* the changes of the RepeatedControl into the record (the recordset leaves it, or saves it) */
    void store()
    {
        if (!uc[0] || !rs || dirty.empty() || storing) return;
        storing = true;
        std::set<std::string> ch = dirty;
        dirty.clear();
        for (vbobj b : binds) {
            RBindData *d = RB(b);
            if (!d || !ch.count(lower(d->prop.toUtf8().constData()))) continue;
            vb_i32 m = vb_mark();
            QByteArray pn = d->prop.toLower().toUtf8();
            vbvar v = vb_late_o(uc[0], pn.constData(), VBM_GET, 0, nullptr);
            if (!vb_errpend) {
                vbvar fs = vb_late_o(rs, "fields", VBM_GET, 0, nullptr);
                vbvar fn = vstr(d->field);
                vbvar f = fs.vt == VT_DISPATCH ? vb_late_o(fs.u.o, "item", VBM_GET, 1, &fn) : vb_v_empty();
                if (f.vt == VT_DISPATCH && !vb_errpend) vb_late_o(f.u.o, "value", VBM_LET, 1, &v);
            }
            vb_release(m);
            if (vb_errpend) break;
        }
        storing = false;
    }

    /* ---------------------------------------------------------------- the data source */
    static void listen_cb(void *ctx, vbobj, const char *ev, int argc, vbvar *argv)
    {
        Ctl *x = C((vbobj)ctx);
        auto *d = x ? dynamic_cast<OcxDataRepeater *>(x->ocx) : nullptr;
        if (d) d->on_event(ev, argc, argv);
    }
    void listen(bool on)
    {
        if (rs && vb_dbhooks.ado_listen) vb_dbhooks.ado_listen(rs, listen_cb, c->self, on ? 1 : 0);
    }
    void bind()
    {
        listen(false);
        if (rs) { vb_orelease(rs); rs = nullptr; }
        vbobj r = source_rs(src, dmember);
        if (r && rs_is_dao(r)) r = nullptr;     /* an ADO rowset only */
        pix.clear();
        top = 0; cur = -1; nrows = -1; adding = del_pending = reset_pending = false; want_top = -1;
        if (!r) { layout(); render(); return; }
        rs = r; vb_oaddref(rs); listen(true);
        nrows = shape(&cur);
        layout(); fill_live(); render();
        if (uc[0]) { post_event("visiblerecordschanged"); post_event("currentrecordchanged"); post_event("activerowchanged"); }
    }
    void on_event(const char *ev, int argc, vbvar *argv)
    {
        int reason = argc > 0 && strcmp(ev, "lb_willchange") ? vb_v2i(argv[0]) : 0;
        if (!strcmp(ev, "lb_save")) store();
        else if (!strcmp(ev, "lb_willchange")) data_update(argc > 0 ? vb_v2i(argv[0]) : -1);
        else if (!strcmp(ev, "willchangerecord")) {
            if (reason == 1) { adding = true; post_vis(); }   /* adRsnAddNew */
        } else if (!strcmp(ev, "recordchangecomplete")) {
            if (reason == 3) {                  /* adRsnUpdate: the record is saved */
                if (adding) adding = false; else post_vis();
                int c0; nrows = shape(&c0);
                if (c0 >= 0) cur = c0;
                fill_live(); render();
            } else if (reason == 2) { del_pending = true; cur = -1; }   /* adRsnDelete: the record is gone */
            else if (reason == 4 || reason == 5) adding = false;   /* adRsnUndoUpdate, adRsnUndoAddNew */
        } else if (!strcmp(ev, "movecomplete")) {
            if (reason == 7) reset_pending = true;             /* adRsnRequery */
        } else if (!strcmp(ev, "lb_fill")) {
            if (del_pending) { del_pending = false; after_delete(); return; }
            sync();
        }
    }
    /* a field of the current record takes a new value: DataUpdate (Cancel, the record, the field) */
    void data_update(int col)
    {
        if (!uc[0]) return;
        int c0 = -1, n = rsx_shape(rs, nullptr, &c0);
        vbvar a[3] = { vb_v_bool(0), c0 >= 0 && c0 < n && vb_dbhooks.ado_bookmark ? vb_dbhooks.ado_bookmark(rs, c0) : vb_v_null(),
                       vstr(QString::fromUtf8(rsx_colname(rs, col))) };
        vb_i32 m = vb_mark();
        void *e = vb_err_stash();
        fire_ctl(c, "dataupdate", 3, a);
        if (vb_errpend) vb_err_clear_pending();
        if (e) vb_err_unstash(e);
        vb_release(m);
        if (!adding) post_vis();
    }
    /* the current record was deleted: the next one (the last when there is none after it) */
    void after_delete()
    {
        vb_i32 m = vb_mark();
        vb_i32 save = vb_errpend;
        vb_late_o(rs, "movenext", VBM_METHOD, 0, nullptr);
        if (!vb_errpend) {
            vbvar eof = vb_late_o(rs, "eof", VBM_GET, 0, nullptr);
            if (!vb_errpend && vb_v2bool(eof)) vb_late_o(rs, "movelast", VBM_METHOD, 0, nullptr);
        }
        if (vb_errpend && !save) vb_err_clear_pending();
        vb_release(m);
        sync();
    }
    /* where the recordset is now: the rows scroll as the current record needs, the events */
    void sync()
    {
        if (!rs) return;
        int c0 = -1, n = shape(&c0);
        bool reset = reset_pending || (nrows >= 0 && n != nrows && !adding);
        reset_pending = false;
        int old_cur = cur, old_act = active(), old_top = top;
        nrows = n; cur = c0;
        int rows = fit(), t = top;
        if (want_top >= 0) { t = want_top; want_top = -1; }
        else if (cur >= 0 && rows > 0) { if (cur < t) t = cur; else if (cur >= t + rows) t = cur - rows + 1; }
        top = std::max(0, std::min(t, std::max(0, n - rows)));
        bool scrolled = top != old_top;
        int act = active();
        fill_live(); render();
        if (!uc[0]) return;
        if (reset) {
            post_event("visiblerecordschanged"); post_event("currentrecordchanged"); post_event("activerowchanged");
            return;
        }
        bool cc = cur != old_cur, ac = act != old_act;
        if (scrolled) {
            post_event("visiblerecordschanged");
            if (cc) {
                for (int k = 0; k < 2; k++) { post_event("currentrecordchanged"); if (ac) post_event("activerowchanged"); }
            } else if (ac) post_event("activerowchanged");
        } else if (cc) {
            post_event("currentrecordchanged");
            if (ac) post_event("activerowchanged");
        } else if (ac) post_event("activerowchanged");
    }
    /* the rows scroll (ActiveRow =, VisibleRecords( ) =, the scroll bar) */
    void scroll_to(int t)
    {
        int old_act = active();
        top = std::max(0, std::min(t, std::max(0, count() - fit())));
        render();
        post_vis();
        if (active() != old_act) post_event("activerowchanged");
    }
    void post_vis() { if (uc[0]) post_event("visiblerecordschanged"); }
    void post_event(const char *ev)
    {
        vbobj self = c->self;
        vb_oaddref(self);
        std::string e = ev;
        post([self, e]() {
            Ctl *x = C(self);
            if (x && x->form && x->ocx) fire_ctl(x, e.c_str());
            vb_orelease(self);
        });
    }

    /* ---------------------------------------------------------------- the look */
    void layout()
    {
        if (!view) return;
        if (vbar) {
            int n = count(), rows = fit();
            vbar->setVisible((scrollbars >= 2) != 0);
            vbar->setGeometry(view->width() - frame() - 16, frame() + caph(), 16,
                              std::max(0, view->height() - 2 * frame() - caph() - ((scrollbars & 1) ? 16 : 0)));
            syncing_bar = true;
            vbar->setRange(0, std::max(0, n - rows));
            vbar->setPageStep(std::max(1, rows));
            vbar->setValue(top);
            syncing_bar = false;
        }
    }
    /* the rows other than the current one, drawn by the second object of the user control */
    void render()
    {
        layout();
        pix.clear();
        if (view) view->update();
        if (!uc[1] || !rs || !view) { place_live(); return; }
        FormState *us = fstate(uc[1]);
        int rows = vrows();
        for (int i = 0; i < rows; i++) {
            int r = top + i;
            if (r == cur) continue;
            filling = true;
            for (vbobj b : binds) { RBindData *d = RB(b); if (d) put(uc[1], d->prop, r, rs_col_index(rs, d->field)); }
            filling = false;
            QRect rc = row_rect(i);
            if (!us->w || rc.width() <= 0) continue;
            us->w->resize(rc.size());
            QPixmap pm(rc.size());
            pm.fill(qcolor(back));
            us->w->render(&pm);
            pix[r] = pm;
        }
        place_live();
    }
    void place_live()
    {
        if (!uc[0]) return;
        FormState *us = fstate(uc[0]);
        if (!us->w || !view) return;
        int a = active();
        if (a <= 0) { us->w->hide(); return; }
        QRect rc = row_rect(a - 1);
        us->w->setGeometry(rc);
        us->w->show();
        us->w->raise();
    }
    void paint(QPainter &p)
    {
        QWidget *w = view;
        p.fillRect(w->rect(), qcolor(back));
        if (border) {
            p.setPen(QColor(128, 128, 128)); p.drawLine(0, 0, w->width() - 1, 0); p.drawLine(0, 0, 0, w->height() - 1);
            p.setPen(QColor(255, 255, 255)); p.drawLine(0, w->height() - 1, w->width() - 1, w->height() - 1);
            p.drawLine(w->width() - 1, 0, w->width() - 1, w->height() - 1);
        }
        if (caph()) {
            QRect cr(frame(), frame(), w->width() - 2 * frame(), caph());
            p.fillRect(cr, qcolor(0x8000000FL));
            p.setPen(qcolor(fore));
            Qt::Alignment al = captionstyle == 2 ? Qt::AlignRight : captionstyle == 3 ? Qt::AlignHCenter : Qt::AlignLeft;
            p.drawText(cr.adjusted(4, 0, -4, 0), Qt::AlignVCenter | al, caption);
        }
        int rows = vrows(), a = active();
        for (int i = 0; i < rows; i++) {
            QRect rc = row_rect(i);
            auto it = pix.find(top + i);
            if (it != pix.end()) p.drawPixmap(rc.topLeft(), it->second);
            if (divider) { p.setPen(QColor(128, 128, 128)); p.drawLine(rc.left() - indw(), rc.bottom(), rc.right(), rc.bottom()); }
            if (rowind) {
                QRect ir(frame(), rc.top(), indw(), rc.height());
                p.fillRect(ir.adjusted(0, 0, 0, -1), qcolor(0x8000000FL));
                if (i + 1 == a) {
                    p.setPen(Qt::NoPen); p.setBrush(Qt::black);
                    QPoint m = ir.center();
                    QPolygon tri; tri << QPoint(m.x() - 2, m.y() - 4) << QPoint(m.x() + 2, m.y()) << QPoint(m.x() - 2, m.y() + 4);
                    p.drawPolygon(tri);
                    p.setBrush(Qt::NoBrush);
                }
            }
        }
    }
    void click_row(int y)                       /* the user picks a row: it becomes the current record */
    {
        if (!rs) return;
        int rows = vrows();
        for (int i = 0; i < rows; i++) {
            QRect rc = row_rect(i);
            if (y < rc.top() || y > rc.bottom()) continue;
            if (top + i != cur && vb_dbhooks.ado_goto) {
                vb_i32 save = vb_errpend;
                vb_dbhooks.ado_goto(rs, top + i);
                if (vb_errpend && !save) vb_err_clear_pending();
            }
            return;
        }
    }

    /* ---------------------------------------------------------------- RepeaterBindings */
    vbvar bindings_obj()
    {
        if (!bindsobj) { bindsobj = vb_onew(&CLS_RBINDS); vb_oaddref(bindsobj); bindsobj->ext = this; }
        return retobj(bindsobj);
    }
    int bind_index(vbvar k)                     /* Item(1-based index or Key): its place, -1 with the error raised */
    {
        k = vb_v_letval(k);
        if (vb_errpend) return -1;
        if (k.vt == VT_BSTR) {
            QString s = data_qs(k);
            for (size_t i = 0; i < binds.size(); i++) { RBindData *d = RB(binds[i]); if (d && !d->key.isEmpty() && d->key == s) return (int)i; }
            rbinds_err(35601);
            return -1;
        }
        int i = vb_v2i(k);
        if (vb_errpend) return -1;
        if (i < 1 || i > (int)binds.size()) { rbinds_err(35600); return -1; }
        return i - 1;
    }
    vbobj add_binding(const QString &prop, const QString &field, vbobj fmt, const QString &key, bool check)
    {
        if (check && ucls) {
            if (!has_member(prop)) { rbinds_err(35704); return nullptr; }
            bool can = false;
            for (const QString &b : bindable()) if (!b.compare(prop, Qt::CaseInsensitive)) can = true;
            if (!can) { rbinds_err(35701); return nullptr; }
            for (vbobj b : binds) { RBindData *d = RB(b); if (d && !d->prop.compare(prop, Qt::CaseInsensitive)) { rbinds_err(35703); return nullptr; } }
            if (rs && rs_col_index(rs, field) < 0) { drep_raise(42000, "Binding", "Invalid field name supplied"); return nullptr; }
        }
        if (check && !key.isEmpty())
            for (vbobj b : binds) { RBindData *d = RB(b); if (d && d->key == key) { rbinds_err(35602); return nullptr; } }
        vbobj o = vb_onew(&CLS_RBIND);
        vb_oaddref(o);
        RBindData *d = new RBindData;
        d->owner = this; d->prop = prop; d->field = field; d->key = key;
        d->fmt = fmt ? vb_oid(fmt) : vb_dfmt_new();
        if (fmt) vb_oaddref(d->fmt);
        o->ext = d;
        binds.push_back(o);
        if (rs && cur >= 0) { fill_live(prop); render(); }
        return o;
    }
    void drop_binding(size_t i)
    {
        RBindData *d = RB(binds[i]);
        if (d) d->owner = nullptr;
        vb_orelease(binds[i]);
        binds.erase(binds.begin() + (long)i);
    }
    vbvar bindings_member(const std::string &n, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
    {
        (void)kind;
        *found = 1;
        if (n == "count") return vb_v_l((vb_i32)binds.size());
        if (n == "item" || n.empty()) {
            if (argc < 1) { vb_raise(450); return vb_v_empty(); }
            int i = bind_index(argv[0]);
            return i < 0 ? vb_v_empty() : retobj(binds[(size_t)i]);
        }
        if (n == "add") {
            if (argc < 2) { vb_raise(450); return vb_v_empty(); }
            QString p = data_qs(argv[0]), f = data_qs(argv[1]);
            vbobj fmt = argc > 2 && !vb_ismissing(argv[2]) && argv[2].vt == VT_DISPATCH ? argv[2].u.o : nullptr;
            QString k = argc > 3 && !vb_ismissing(argv[3]) ? data_qs(argv[3]) : QString();
            if (vb_errpend) return vb_v_empty();
            vbobj o = add_binding(p, f, fmt, k, true);
            return o ? retobj(o) : vb_v_empty();
        }
        if (n == "remove") {
            if (argc < 1) { vb_raise(450); return vb_v_empty(); }
            int i = bind_index(argv[0]);
            if (i >= 0) { drop_binding((size_t)i); render(); }
            return vb_v_empty();
        }
        if (n == "clear") { while (!binds.empty()) drop_binding(binds.size() - 1); render(); return vb_v_empty(); }
        *found = 0;
        return vb_v_empty();
    }
    vbvar str_array(const QStringList &l)
    {
        vbarr a = { 0 };
        vb_aredim(&a, (vb_i32)sizeof(vbs), VBA_STR, 0, 1, 0, (vb_i32)l.size() - 1);
        for (int i = 0; i < l.size(); i++) vb_sset(&((vbs *)a.data)[i], vs(l[i]));
        vbvar r = vb_v_arr(&a, VT_BSTR);
        vb_afree(&a);
        return vb_vtmp(r);
    }
};

static void drep_end()                          /* End: the repeated controls end (their Terminate) */
{
    std::vector<OcxDataRepeater *> all(drep_all().begin(), drep_all().end());
    for (auto *d : all) d->release_ucs(false);
}
OcxDataRepeater::OcxDataRepeater()
{
    drep_all().insert(this);
    vb_end_hook = drep_end;
}
OcxDataRepeater::~OcxDataRepeater()
{
    drep_all().erase(this);
    listen(false);
    release_ucs(false);
    while (!binds.empty()) drop_binding(binds.size() - 1);
    if (bindsobj) { bindsobj->ext = nullptr; vb_orelease(bindsobj); }
    if (rs) vb_orelease(rs);
    if (src) vb_orelease(src);
}
/* PropertyChanged in the user control (uctl.inc): a bound property of the current record changed */
static void drep_prop_changed(vbobj o, vbvar name)
{
    o = vb_oid(o);
    for (auto *d : drep_all()) {
        if (d->uc[0] != o) continue;
        if (d->filling || d->storing) return;
        vb_i32 m = vb_mark();
        d->dirty.insert(lower(vb_stoutf8(vb_v2s(name))));
        vb_release(m);
        return;
    }
}

void DrepView::paintEvent(QPaintEvent *)
{
    Ctl *x = C(self);
    auto *d = x ? dynamic_cast<OcxDataRepeater *>(x->ocx) : nullptr;
    if (!d) return;
    QPainter p(this);
    d->paint(p);
}
void DrepView::mousePressEvent(QMouseEvent *e)
{
    Ctl *x = C(self);
    auto *d = x ? dynamic_cast<OcxDataRepeater *>(x->ocx) : nullptr;
    if (d) d->click_row((int)ev_pos(e).y());   /* (Qt 5 for Windows: vbforms.cpp) */
}
void DrepView::resizeEvent(QResizeEvent *)
{
    Ctl *x = C(self);
    auto *d = x ? dynamic_cast<OcxDataRepeater *>(x->ocx) : nullptr;
    if (d) d->render();
}

QWidget *OcxDataRepeater::make(QWidget *parent)
{
    view = new DrepView(parent);
    view->self = c->self;
    vbar = new QScrollBar(Qt::Vertical, view);
    vbar->setFocusPolicy(Qt::NoFocus);
    vbobj self = c->self;
    QObject::connect(vbar, &QScrollBar::valueChanged, [self](int v) {
        Ctl *x = C(self);
        auto *d = x ? dynamic_cast<OcxDataRepeater *>(x->ocx) : nullptr;
        if (d && !d->syncing_bar && v != d->top) d->scroll_to(v);
    });
    return view;
}
void OcxDataRepeater::load(const char *const *p)
{
    QString v;
    if (prop(p, "Caption", &v)) caption = v;
    if (prop(p, "DataMember", &v)) dmember = v;
    if (prop(p, "DataSource", &v)) srcname = v;
    captionstyle = (int)propn(p, "CaptionStyle", 1);
    scrollbars = (int)propn(p, "ScrollBars", 2);
    divider = (int)propn(p, "RowDividerStyle", 2);
    border = (int)propn(p, "BorderStyle", 1);
    appearance = (int)propn(p, "Appearance", 1);
    rowind = propn(p, "RowIndicator", -1) != 0;
    integral = propn(p, "IntegralHeight", -1) != 0;
    back = propc(p, "BackColor", 0x8000000FL);
    fore = propc(p, "ForeColor", 0x80000012L);
    int nb = (int)propn(p, "RepeaterBindings", 0);
    for (int i = 0; i < nb; i++) {
        QString pre = QString("RepeaterBinding(%1).").arg(i), pn, fl, ky;
        prop(p, (pre + "PropertyName").toUtf8().constData(), &pn);
        prop(p, (pre + "DataField").toUtf8().constData(), &fl);
        prop(p, (pre + "Key").toUtf8().constData(), &ky);
        if (!pn.isEmpty()) add_binding(pn, fl, nullptr, ky, false);
    }
    if (prop(p, "RepeatedControlName.Name", &v) && !v.isEmpty()) set_progid(v, false);
}
vbvar OcxDataRepeater::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    *found = 1;
    if (n == "repeatedcontrolname") {
        if (set) { QString s = data_qs(val); if (!vb_errpend) set_progid(s, true); return vb_v_empty(); }
        return vstr(progid);
    }
    if (n == "repeatedcontrol") { if (set) { vb_raise(383); return vb_v_empty(); } return retobj(uc[0]); }
    if (n == "repeaterbindings") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        if (argc > 0) { int i = bind_index(argv[0]); return i < 0 ? vb_v_empty() : retobj(binds[(size_t)i]); }   /* RepeaterBindings(i) */
        return bindings_obj();
    }
    if (n == "propertynames") { if (set) { vb_raise(383); return vb_v_empty(); } return ucls && uc[0] ? str_array(bindable()) : vb_v_null(); }
    if (n == "datafields") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        if (!rs) return vb_v_null();
        int nc = 0;
        rsx_shape(rs, &nc, nullptr);
        QStringList l;
        for (int i = 0; i < nc; i++) l << QString::fromUtf8(rsx_colname(rs, i));
        return str_array(l);
    }
    if (n == "visiblerows") { if (set) { vb_raise(383); return vb_v_empty(); } return vb_v_l(vrows()); }
    if (n == "recordcount") { if (set) { vb_raise(383); return vb_v_empty(); } return vb_v_l(count()); }
    if (n == "activerow") {
        if (!set) return vb_v_l(active());
        int i = vb_v2i(val);
        if (vb_errpend) return vb_v_empty();
        int rows = vrows();
        if (!rows) { drep_err(32004); return vb_v_empty(); }
        if (i < 1 || i > rows) { drep_err(380); return vb_v_empty(); }
        if (cur < 0) { drep_err(32005); return vb_v_empty(); }
        int t = cur - (i - 1);
        int old_act = active();
        if (t >= 0 && t <= std::max(0, count() - fit())) { top = t; render(); }
        post_vis();
        if (active() != old_act) post_event("activerowchanged");
        return vb_v_empty();
    }
    if (n == "visiblerecords") {
        if (argc < (set ? 2 : 1)) { vb_raise(450); return vb_v_empty(); }
        int i = vb_v2i(argv[0]);
        if (vb_errpend) return vb_v_empty();
        int rows = vrows();
        if (!set) {
            if (i < 1 || i > rows) { drep_err(381); return vb_v_empty(); }
            return vb_dbhooks.ado_bookmark ? vb_dbhooks.ado_bookmark(rs, top + i - 1) : vb_v_null();
        }
        if (!rows) { drep_err(32004); return vb_v_empty(); }
        if (i < 1 || i > rows) { drep_err(32006); return vb_v_empty(); }
        int r = vb_dbhooks.ado_bmrow ? vb_dbhooks.ado_bmrow(rs, argv[argc - 1]) : -1;
        if (r < 0) { drep_err(32002); return vb_v_empty(); }
        scroll_to(r - (i - 1));
        return vb_v_empty();
    }
    if (n == "currentrecord") {
        if (!set) {
            if (!rs || !uc[0] || cur < 0 || cur >= count() || adding) return vb_v_null();
            return vb_dbhooks.ado_bookmark ? vb_dbhooks.ado_bookmark(rs, cur) : vb_v_null();
        }
        if (!vrows()) { drep_err(32004); return vb_v_empty(); }
        int r = vb_dbhooks.ado_bmrow ? vb_dbhooks.ado_bmrow(rs, val) : -1;
        if (r < 0) { drep_err(32002); return vb_v_empty(); }
        want_top = r;
        if (r != cur && vb_dbhooks.ado_goto) vb_dbhooks.ado_goto(rs, r);
        if (want_top >= 0 && !vb_errpend) sync();
        want_top = -1;
        return vb_v_empty();
    }
    if (n == "datasource") {
        if (!set) return retobj(src);
        vbobj o = vb_v2obj(val);
        if (vb_errpend) return vb_v_empty();
        if (o) vb_oaddref(o);
        if (src) vb_orelease(src);
        src = o ? vb_oid(o) : nullptr;
        srcname.clear();
        bind();
        return vb_v_empty();
    }
    if (n == "datamember") { if (set) { dmember = data_qs(val); if (src) bind(); return vb_v_empty(); } return vstr(dmember); }
    if (n == "refresh") { render(); return vb_v_empty(); }
    if (n == "caption") { if (set) { caption = data_qs(val); if (view) view->update(); return vb_v_empty(); } return vstr(caption); }
    /* the numbers and Booleans of the look */
    struct NumProp { const char *name; int *iv; bool *bv; };
    NumProp np[] = { { "captionstyle", &captionstyle, nullptr }, { "scrollbars", &scrollbars, nullptr },
                     { "rowdividerstyle", &divider, nullptr }, { "borderstyle", &border, nullptr },
                     { "appearance", &appearance, nullptr }, { "rowindicator", nullptr, &rowind },
                     { "integralheight", nullptr, &integral } };
    for (auto &q : np) {
        if (n != q.name) continue;
        if (set) {
            if (q.iv) *q.iv = vb_v2i(val); else *q.bv = vb_v2bool(val) != 0;
            if (!vb_errpend) render();
            return vb_v_empty();
        }
        return q.iv ? vb_v_i((vb_i16)*q.iv) : vb_v_bool(*q.bv ? -1 : 0);
    }
    if (n == "backcolor" || n == "forecolor") {
        long &col = n == "backcolor" ? back : fore;
        if (set) { col = norm_color(vb_v2d(val)); if (!vb_errpend && view) render(); return vb_v_empty(); }
        return vb_v_l((vb_i32)col);
    }
    *found = 0;
    return vb_v_empty();
}
static Ocx *new_drep() { return new OcxDataRepeater; }

static vbvar rbinds_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    auto *d = (OcxDataRepeater *)vb_oid(o)->ext;
    if (!d) { *found = 1; vb_raise(91); return vb_v_empty(); }
    return d->bindings_member(lower(name), kind, argc, argv, found);
}
static vb_i32 rbinds_count(vbobj o)
{
    auto *d = (OcxDataRepeater *)vb_oid(o)->ext;
    return d ? (vb_i32)d->binds.size() : 0;
}
static vbvar rbinds_item(vbobj o, vb_i32 i)
{
    auto *d = (OcxDataRepeater *)vb_oid(o)->ext;
    if (!d || i < 0 || i >= (vb_i32)d->binds.size()) return vb_v_empty();
    return retobj(d->binds[(size_t)i]);
}
static vbvar rbind_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    RBindData *d = RB(o);
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (!d) { vb_raise(91); return vb_v_empty(); }
    if (n == "propertyname") { if (set) { d->prop = data_qs(val); return vb_v_empty(); } return vstr(d->prop); }
    if (n == "datafield") { if (set) { d->field = data_qs(val); return vb_v_empty(); } return vstr(d->field); }
    if (n == "key") { if (set) { d->key = data_qs(val); return vb_v_empty(); } return vstr(d->key); }
    if (n == "datachanged") { if (set) { d->changed = vb_v2bool(val) != 0; return vb_v_empty(); } return vb_v_bool(d->changed ? -1 : 0); }
    if (n == "dataformat") {
        if (!set) return retobj(d->fmt);
        vbobj f = vb_v2obj(val);
        if (vb_errpend) return vb_v_empty();
        if (f) vb_oaddref(f);
        if (d->fmt) vb_orelease(d->fmt);
        d->fmt = f ? vb_oid(f) : nullptr;
        return vb_v_empty();
    }
    if (n == "tag") { if (set) { vb_vset(&d->tag, val); return vb_v_empty(); } return vb_vtmp(d->tag); }
    if (n == "index") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        if (d->owner)
            for (size_t i = 0; i < d->owner->binds.size(); i++) if (vb_oid(d->owner->binds[i]) == vb_oid(o)) return vb_v_l((vb_i32)i + 1);
        return vb_v_l(0);
    }
    *found = 0;
    return vb_v_empty();
}
