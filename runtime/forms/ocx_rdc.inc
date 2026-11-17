/* ocx_rdc.inc - the RemoteData control of VB6 (MSRDC20.OCX, MSRDC.MSRDC; phase 7.8) over RDO 2.0 (runtime/db/rdo.inc):
 * Connect, DataSourceName, SQL, the cursor (CursorDriver, ResultsetType, LockType ...), Refresh, Resultset,
 * Connection, Environment, UpdateRow, UpdateControls, the buttons |< < > >| with BOFAction / EOFAction, and the
 * controls bound to it (DataSource = "MSRDC1"). It opens when its form loads (Connect and SQL in the .frm) or on
 * Refresh; the bound controls show the row once the form is shown. Before the resultset moves or deletes a row
 * (rdo.inc: "lb_validate") the changed bound values are saved (Validate rdActionUpdateModified 13, then the row is
 * updated) and Validate comes with the action (MoveFirst 1 ... MoveLast 4, Delete 7); a value that does not go
 * cancels the change (Validate 15) and comes later as an Error event (Data Type Conversion Error 40527); without a
 * current row the move is canceled (40503). After the move: the bound controls, then Reposition. The errors of the
 * control come as its Error event after the code that caused them (posted), as VB6 does (VB6 oracle f_rdo2).
 * Included by vbforms.cpp after data.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

struct OcxRdc : Ocx {
    QString dsn, sql, connect, user, pwd, logmsg, caption;
    int options = 0, cursor = 0, bofaction = 0, eofaction = 0, rstype = 1, locktype = 3, prompt = 0, querytype = 0;
    int qtimeout = 30, rowset = 100, login = 15, keyset = 0, maxrows = 0, errthr = -1, batch = 15, appearance = 1;
    bool readonly = false;
    long back = 0x80000005L, fore = 0x80000008L;
    vbobj env = nullptr, cn = nullptr, rs = nullptr, inner = nullptr;   /* counted */
    bool shown_once = false, storing = false;
    QPointer<QLabel> label;
    ~OcxRdc() override { close(); }
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
                if (auto *a = x ? dynamic_cast<OcxRdc *>(x->ocx) : nullptr) a->button(i);
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
        if (prop(p, "DataSourceName", &v)) dsn = v;
        if (prop(p, "RecordSource", &v)) sql = v;
        if (prop(p, "Connect", &v)) connect = v;
        if (prop(p, "UserName", &v)) user = v;
        if (prop(p, "Password", &v)) pwd = v;
        if (prop(p, "LogMessages", &v)) logmsg = v;
        if (prop(p, "Caption", &v)) caption = v;
        options = (int)propn(p, "Options", 0);
        cursor = (int)propn(p, "CursorDriver", 0);
        bofaction = (int)propn(p, "BOFAction", 0);
        eofaction = (int)propn(p, "EOFAction", 0);
        rstype = (int)propn(p, "RecordsetType", 1);
        locktype = (int)propn(p, "LockType", 3);
        querytype = (int)propn(p, "QueryType", 0);
        prompt = (int)propn(p, "Prompt", 0);
        appearance = (int)propn(p, "Appearance", 1);
        qtimeout = (int)propn(p, "QueryTimeout", 30);
        rowset = (int)propn(p, "RowsetSize", 100);
        login = (int)propn(p, "LoginTimeout", 15);
        keyset = (int)propn(p, "KeysetSize", 0);
        maxrows = (int)propn(p, "MaxRows", 0);
        errthr = (int)propn(p, "ErrorThreshold", -1);
        batch = (int)propn(p, "BatchSize", 15);
        readonly = propn(p, "ReadOnly", 0) != 0;
        back = propc(p, "BackColor", 0x80000005L);
        fore = propc(p, "ForeColor", 0x80000008L);
        if (label) label->setText(caption);
    }
    void shown() override
    {
        if (shown_once) return;
        shown_once = true;
        if (inner) fill_bound();
    }
    /* ---------------------------------------------------------------- the bound controls */
    bool is_bound(Ctl *b)
    {
        return b && b != c && !b->datafield.isEmpty() && !b->dsobj && named_source(b) == c &&
               !(b->kind == K_OCX && b->ocx && b->ocx->own_binding());
    }
    void fill_bound()
    {
        if (!c->form || !inner) return;
        bool rec = rs_has_record(inner);
        std::vector<vbobj> list = c->form->ctls;
        /* without a current row: empty, a CheckBox unchecked (VB6 oracle f_rdo2; ADO: grayed) */
        for (vbobj co : list) { Ctl *b = C(co); if (is_bound(b)) fill_ctl(b, inner, rec, rec); }
    }
    void clear_changed()
    {
        if (!c->form) return;
        for (vbobj co : c->form->ctls) { Ctl *b = C(co); if (is_bound(b)) b->datachanged = false; }
    }
    bool changed()
    {
        if (!c->form) return false;
        for (vbobj co : c->form->ctls) { Ctl *b = C(co); if (is_bound(b) && b->datachanged) return true; }
        return false;
    }
    /* the changed values into the current row, then Update; 0, or -1 with the error pending */
    int store_bound()
    {
        if (!c->form || !inner) return 0;
        storing = true;
        int r = 0;
        std::vector<vbobj> list = c->form->ctls;
        for (vbobj co : list) {
            Ctl *b = C(co);
            if (!is_bound(b) || !b->datachanged) continue;
            r = store_ctl(b, inner, true);
            if (r == -2) { vb_raise(13); r = -1; }
            if (r) break;
        }
        if (!r && rs_editing()) {
            vb_late_o(inner, "update", VBM_METHOD, 0, nullptr);
            if (vb_errpend) r = -1;
        }
        storing = false;
        return r;
    }
    bool rs_editing()
    {
        vb_i32 save = vb_errpend;
        int em = vb_v2i(vb_late_o(inner, "editmode", VBM_GET, 0, nullptr));
        if (vb_errpend && !save) { vb_err_clear_pending(); return false; }
        return em != 0;
    }
    void cancel_edit()
    {
        if (!inner) return;
        vb_i32 save = vb_errpend;
        vb_late_o(inner, "cancelupdate", VBM_METHOD, 0, nullptr);
        if (vb_errpend && !save) vb_err_clear_pending();
    }
    /* ---------------------------------------------------------------- events */
    bool validate(int action)                   /* false: canceled (Action = 0) */
    {
        vbvar a[2] = { vb_v_i((vb_i16)action), vb_v_i(0) };
        vb_i32 m = vb_mark();
        fire_ctl(c, "validate", 2, a);
        int act = vb_v2i(a[0]);
        vb_release(m);
        return act != 0;
    }
    /* the Error event, after the code that is running now (VB6 posts it) */
    void post_error(vb_i32 num, const QString &desc)
    {
        vbobj self = c->self;
        vb_oaddref(self);
        QTimer::singleShot(0, [self, num, desc]() {
            Ctl *x = C(self);
            if (x && x->form) {
                vbvar a[7] = { vb_v_l(num), vstr(desc), vb_v_l((vb_i32)(0x800A0000u | (vb_u32)(num & 0xFFFF))), vstr(x->name),
                               vstr(QString()), vb_v_l(0), vb_v_bool(0) };
                vb_i32 m = vb_mark();
                fire_ctl(x, "error", 7, a);
                bool cancel = vb_v2bool(a[6]) != 0;
                vb_release(m);
                if (!cancel) vb_msgbox(vs(desc), 16, vb_app_exename());
            }
            vb_orelease(self);
        });
    }
    /* the number of the Error event for an error of RDO (40000 + n of MSRDO20.DLL -> 56515 + n of MSRDC20.OCX) */
    static vb_i32 event_number(vb_i32 n) { return n >= 40000 && n <= 40034 ? 56515 + (n - 40000) : n; }
    void post_pending()                         /* the error pending: cleared, as an Error event later */
    {
        vb_i32 num = vb_err_number_pending();
        QString msg = QString::fromUtf8(vb_err_text_pending());
        vb_err_clear_pending();
        post_error(event_number(num), msg);
    }
    /* the resultset moves (action 1 .. 4, 9) or deletes (7): the changes saved, Validate */
    void before(int action, int *cancel)
    {
        bool failed = false;
        if (changed()) {
            if (!rs_has_record(inner)) { *cancel = 1; return; }
            if (!validate(13)) { *cancel = 1; return; }     /* rdActionUpdateModified */
            if (store_bound()) {
                vb_i32 num = vb_err_number_pending();
                QString msg = QString::fromUtf8(vb_err_text_pending());
                vb_err_clear_pending();
                cancel_edit();
                if (num == 13 || num == -2147352571) post_error(40527, QString("Data Type Conversion Error"));
                else post_error(56517, msg);
                failed = true;
            }
            clear_changed();
        }
        if (!validate(action)) { *cancel = 1; return; }
        if (failed) validate(15);                /* rdActionCancelUpdate */
    }
    static void rdc_listen(void *ctx, vbobj, const char *ev, int argc, vbvar *argv)
    {
        Ctl *x = C((vbobj)ctx);
        auto *a = x ? dynamic_cast<OcxRdc *>(x->ocx) : nullptr;
        if (!a) return;
        if (!strcmp(ev, "lb_validate") && argc >= 2) {
            int cancel = 0;
            a->before(vb_v2i(argv[0]), &cancel);
            if (cancel) argv[1] = vb_v_l(1);
        } else if (!strcmp(ev, "lb_fill")) {
            if (a->shown_once && !a->storing) a->fill_bound();
        } else if (!strcmp(ev, "lb_reposition")) {
            fire_ctl(x, "reposition");
        }
    }
    void listen(bool on)
    {
        if (inner && vb_dbhooks.ado_listen) vb_dbhooks.ado_listen(inner, rdc_listen, c->self, on ? 1 : 0);
    }
    /* ---------------------------------------------------------------- open, close */
    void drop_rs()
    {
        listen(false);
        if (inner) { vb_orelease(inner); inner = nullptr; }
        if (rs) {
            vb_i32 save = vb_errpend;
            vb_late_o(rs, "close", VBM_METHOD, 0, nullptr);
            if (vb_errpend && !save) vb_err_clear_pending();
            vb_orelease(rs); rs = nullptr;
        }
    }
    void close()
    {
        drop_rs();
        if (cn) {
            vb_i32 save = vb_errpend;
            vb_late_o(cn, "close", VBM_METHOD, 0, nullptr);
            if (vb_errpend && !save) vb_err_clear_pending();
            vb_orelease(cn); cn = nullptr;
        }
        if (env) { vb_orelease(env); env = nullptr; }
    }
    void set_rs(vbobj r)                        /* r: an rdoResultset (not counted yet) */
    {
        drop_rs();
        if (!r) return;
        rs = vb_oid(r); vb_oaddref(rs);
        vbobj in = vb_dbhooks.rdo_inner ? vb_dbhooks.rdo_inner(rs) : nullptr;
        if (in) { inner = in; vb_oaddref(inner); listen(true); }
    }
    QString full_connect()
    {
        QString s = connect;
        if (!user.isEmpty() && !s.contains("UID=", Qt::CaseInsensitive)) s += (s.isEmpty() ? "" : ";") + QString("UID=") + user;
        if (!pwd.isEmpty() && !s.contains("PWD=", Qt::CaseInsensitive)) s += (s.isEmpty() ? "" : ";") + QString("PWD=") + pwd;
        return s;
    }
    /* Refresh: the connection and the resultset again; 0, or -1 with the error pending */
    int refresh()
    {
        if (inner && changed()) {               /* the changes of the bound controls first */
            if (!rs_has_record(inner)) {
                for (vbobj co : c->form->ctls) { Ctl *b = C(co); if (is_bound(b) && b->datachanged) post_error(56524, QString("No current row.")); }
                vb_raise_full(-2147467259, vs(c->name), vs(QString("Could not refresh controls.")));
                return -1;
            }
            if (store_bound()) { cancel_edit(); return -1; }
            clear_changed();
        }
        close();
        if (!vb_dbhooks.createobject) { vb_raise_full(429, vs(QString("MSRDC")), vs(QString("ActiveX component can't create object"))); return -1; }
        vbobj eng = vb_dbhooks.createobject("rdo.rdoengine");
        if (!eng) { vb_raise(429); return -1; }
        vbvar zero = vb_v_l(0);
        vbvar ev = vb_late_o(eng, "rdoenvironments", VBM_GET, 1, &zero);
        if (vb_errpend) return -1;
        env = vb_v2obj(ev); vb_oaddref(env);
        vbvar oa[4] = { vstr(dsn), vb_v_l(prompt == 0 ? 1 : prompt), vb_v_bool(readonly ? -1 : 0), vstr(full_connect()) };
        vbvar cv = vb_late_o(env, "openconnection", VBM_METHOD, 4, oa);
        if (vb_errpend) return -1;
        cn = vb_v2obj(cv); vb_oaddref(cn);
        vbvar v = vb_v_l(cursor);
        vb_late_o(cn, "cursordriver", VBM_LET, 1, &v);
        v = vb_v_l(qtimeout);
        vb_late_o(cn, "querytimeout", VBM_LET, 1, &v);
        if (vb_errpend) return -1;
        vbvar ra[4] = { vstr(sql), vb_v_l(rstype), vb_v_l(readonly ? 1 : locktype), vb_v_l(options & ~32) };
        vbvar rv = vb_late_o(cn, "openresultset", VBM_METHOD, 4, ra);
        if (vb_errpend) return -1;
        vbobj r = rv.vt == VT_DISPATCH ? rv.u.o : nullptr;
        set_rs(r);
        clear_changed();
        fire_ctl(c, "querycompleted");
        if (shown_once) fill_bound();
        reset_users();
        return 0;
    }
    void reset_users()                          /* the grids and lists of this control read the new rows */
    {
        if (!c->form) return;
        std::vector<vbobj> list = c->form->ctls;
        for (vbobj co : list) {
            Ctl *b = C(co);
            if (b && b != c && b->kind == K_OCX && b->ocx) b->ocx->source_reset(c->self);
        }
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
                else if (eofaction == 2) { vb_late_o(rs, "addnew", VBM_METHOD, 0, nullptr); if (!vb_errpend) fill_bound(); }
            }
            if (which == 1 && bof && bofaction == 0) vb_late_o(rs, "movefirst", VBM_METHOD, 0, nullptr);
        }
        if (vb_errpend) post_pending();
    }
    bool no_rs()                                /* a member of the resultset without one */
    {
        if (rs) return false;
        vb_raise_full(-2147467259, vs(c->name), vs(QString("Resultset not available.")));
        return true;
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        QString *sp = n == "datasourcename" ? &dsn : n == "sql" ? &sql : n == "connect" ? &connect : n == "username" ? &user :
                      n == "password" ? &pwd : n == "logmessages" ? &logmsg : n == "caption" ? &caption : nullptr;
        if (sp) {
            if (!set) return vstr(*sp);
            *sp = data_qs(val);
            if (n == "caption" && label) label->setText(caption);
            return vb_v_empty();
        }
        int *ip = n == "options" ? &options : n == "cursordriver" ? &cursor : n == "bofaction" ? &bofaction :
                  n == "eofaction" ? &eofaction : n == "resultsettype" ? &rstype : n == "locktype" ? &locktype :
                  n == "prompt" ? &prompt : n == "querytimeout" ? &qtimeout : n == "rowsetsize" ? &rowset :
                  n == "logintimeout" ? &login : n == "keysetsize" ? &keyset : n == "maxrows" ? &maxrows :
                  n == "errorthreshold" ? &errthr : n == "batchsize" ? &batch : n == "appearance" ? &appearance : nullptr;
        if (ip) { if (!set) return vb_v_l(*ip); *ip = vb_v2l(val); return vb_v_empty(); }
        if (n == "readonly") { if (set) { readonly = vb_v2bool(val) != 0; return vb_v_empty(); } return vbool(readonly); }
        if (n == "backcolor" || n == "forecolor") {
            long *cp = n == "backcolor" ? &back : &fore;
            if (!set) return vb_v_l((vb_i32)*cp);
            *cp = vb_v2l(val);
            return vb_v_empty();
        }
        if (n == "resultset") {
            if (set) {
                vbobj o = vb_v2obj(val);
                if (vb_errpend) return vb_v_empty();
                set_rs(o);
                if (shown_once) fill_bound();
                reset_users();
                return vb_v_empty();
            }
            if (argc > 0 && rs) return vb_late_o(rs, "", kind, argc, argv);   /* MSRDC1.Resultset!Name */
            return vb_vtmp(vb_v_obj(rs));
        }
        if (n == "connection") return vb_vtmp(vb_v_obj(cn));
        if (n == "environment") return vb_vtmp(vb_v_obj(env));
        if (n == "refresh") {
            if (refresh()) {
                vb_i32 num = vb_err_number_pending();
                QString msg = QString::fromUtf8(vb_err_text_pending());
                QString src = QString::fromUtf8(vb_stoutf8(vb_err_source()));
                if (num != -2147467259) post_error(event_number(num), msg);   /* the resultset is gone (VB6 oracle f_rdo2) */
                vb_err_clear_pending();
                vb_raise_full(num, vs(src), vs(msg));
            }
            return vb_v_empty();
        }
        if (n == "updatecontrols") { clear_changed(); fill_bound(); return vb_v_empty(); }
        if (n == "updaterow") {                 /* the changes saved, without Validate */
            if (!rs || !inner || !changed() || !rs_has_record(inner)) return vb_v_empty();
            if (store_bound()) { cancel_edit(); return vb_v_empty(); }
            clear_changed();
            return vb_v_empty();
        }
        if (n == "begintrans" || n == "committrans" || n == "rollbacktrans") {
            if (cn) return vb_late_o(cn, n.c_str(), VBM_METHOD, 0, nullptr);
            return vb_v_empty();
        }
        if (n == "cancel" || n == "aboutbox") return vb_v_empty();
        if (n == "stillexecuting") return vbool(false);
        if (n == "editmode") { if (no_rs()) return vb_v_empty(); return vb_late_o(rs, "editmode", VBM_GET, 0, nullptr); }
        if (n == "transactions") { if (no_rs()) return vb_v_empty(); return vbool(true); }
        if (n == "version") { if (!cn) { no_rs(); return vb_v_empty(); } return vb_late_o(cn, "version", VBM_GET, 0, nullptr); }
        if (n == "updatecriteria" || n == "updateoperation" || n == "batchcollisioncount" || n == "batchcollisionrows") {
            if (no_rs()) return vb_v_empty();
            return vb_late_o(rs, n.c_str(), kind, set ? 1 : 0, set ? &val : nullptr);
        }
        *found = 0;
        return vb_v_empty();
    }
};
static Ocx *new_rdc() { return new OcxRdc; }
static vbobj rdc_inner(Ctl *c) { auto *a = c ? dynamic_cast<OcxRdc *>(c->ocx) : nullptr; return a ? a->inner : nullptr; }

/* the RemoteData controls of a form open when it has loaded, before Form_Load (VB6 oracle f_rdo3) */
static void rdc_after_load(FormState *st)
{
    std::vector<vbobj> list = st->ctls;
    for (vbobj co : list) {
        Ctl *c = C(co);
        auto *a = c ? dynamic_cast<OcxRdc *>(c->ocx) : nullptr;
        if (!a || a->rs || a->sql.isEmpty() || (a->connect.isEmpty() && a->dsn.isEmpty())) continue;
        if (a->refresh()) a->post_pending();
    }
}
