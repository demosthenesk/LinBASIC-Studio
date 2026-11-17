/* uctl.inc - user controls (.ctl): a form of its own that lives inside a control of another form.
 *
 * The user control is an object of its class (compiled like a form, vbf_formdef.mdi == 2); its
 * window and controls are built when it is made (UserControl_Initialize can use them), then put
 * into the form that has it. The control of that form (Ctl, kind K_USER) is its Extender: Name,
 * Left, Top, Visible, Tag ... are the Extender's, the rest are the user control's own public
 * members. RaiseEvent in the user control reaches the form's handlers (Counter1_Changed).
 * The order of the events is VB6's (t: f_uctl): Initialize, Resize, ReadProperties for each
 * control, Form_Load, the Show of each after the code that showed the form, before Activate.
 * Included by vbforms.cpp.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

static void drep_prop_changed(vbobj uc, vbvar name);   /* ocx_drep.inc: the DataRepeater's repeated control */

/* the user controls of the program and of the control projects it has: "project.name" and "name" */
static std::map<std::string, const vb_class *> &uctl_registry()
{
    static std::map<std::string, const vb_class *> m;
    return m;
}
extern "C" void vbf_uctl_register(const char *progid, const vb_class *cls)
{
    std::string k = lower(progid);
    uctl_registry()[k] = cls;
    size_t dot = k.rfind('.');
    if (dot != std::string::npos) uctl_registry().emplace(k.substr(dot + 1), cls);
}
static const vb_class *uctl_class(const char *cls)
{
    if (!cls) return nullptr;
    auto &m = uctl_registry();
    auto it = m.find(lower(cls));
    return it == m.end() ? nullptr : it->second;
}

/* the user control's own window: built when its object is made (in UserControl_Initialize
 * already), without a form around it until the form that has it takes it */
static void uctl_load(vbobj o)
{
    o = vb_oid(o);
    FormState *st = fstate(o);
    if (st->loaded) return;
    ensure_app();
    const vbf_formdef *d = st->def;
    st->loaded = true; st->visible = false; st->focus = nullptr;
    st->ctls.clear(); st->byname.clear();
    /* a child widget from the start (a window would paint its own background, not the form's):
     * it waits in a hidden holder until a form takes it */
    static QPointer<QWidget> holder;
    if (!holder) { holder = new QWidget(nullptr); holder->setAttribute(Qt::WA_DontShowOnScreen, true); }
    QWidget *w = new QWidget(holder);
    w->setAutoFillBackground(false);          /* the form under it shows where it paints nothing */
    st->w = w;
    st->clientw = propn(d->props, "ClientWidth", 4800); st->clienth = propn(d->props, "ClientHeight", 3600);
    st->font = VFont();
    font_from_props(&st->font, d->props);
    st->g = Gfx();
    st->g.back = propc(d->props, "BackColor", 0x8000000F);
    st->g.fore = propc(d->props, "ForeColor", 0x80000012);
    st->g.font = &st->font;
    gfx_from_props(&st->g, d->props, st->clientw, st->clienth);
    st->g.pic = frx_image(d->props, "Picture");
    if (!st->g.pic.isNull()) { st->picture = new_pic(st->g.pic, 1); vb_oaddref(st->picture); gfx_clear(&st->g); }
    st->mousepointer = (int)propn(d->props, "MousePointer", 0);
    auto *lay = new QVBoxLayout(w);
    lay->setContentsMargins(0, 0, 0, 0);
    lay->setSpacing(0);
    Canvas *cv = new Canvas(&st->g, w);
    cv->setFont(qfont(st->font));
    cv->setFocusPolicy(Qt::NoFocus);
    lay->addWidget(cv);
    st->canvas = cv;
    if (propn(d->props, "BackStyle", 1) == 0) cv->set_transparent();   /* Transparent: the form shows through */
    cv->on_paint = [o]() { fire_form(o, "paint"); };
    cv->on_resize = [o](int pw, int ph) {     /* the size comes from the Extender (uctl_geometry) */
        FormState *s = fstate(o);
        if (s->g.pw != pw || s->g.ph != ph) gfx_resize(&s->g, pw, ph);
    };
    w->resize((int)(st->clientw / TPP), (int)(st->clienth / TPP));
    build_ctls(st);
    set_tab_order(st);
    for (vbobj co : st->ctls) timer_update(C(co));
}

/* the size of the user control is the Extender's: ScaleWidth, ScaleHeight, then its Resize */
static void uctl_size(Ctl *c, bool event)
{
    if (!c->uc) return;
    FormState *us = fstate(c->uc);
    int pw = std::max(0, (int)(c->width / TPP)), ph = std::max(0, (int)(c->height / TPP));
    double cw = pw * TPP, ch = ph * TPP;
    bool changed = cw != us->clientw || ch != us->clienth || !c->uc_sized;
    us->clientw = cw; us->clienth = ch;
    if (us->w) us->w->resize(pw, ph);
    if (us->g.pw != pw || us->g.ph != ph) gfx_resize(&us->g, pw, ph);
    if (us->g.scalemode == 1) { us->g.sw = cw; us->g.sh = ch; }
    else if (us->g.scalemode != 0) gfx_rescale(&us->g);
    if (event && changed) { c->uc_sized = true; fire_form(c->uc, "resize"); }
}

static QWidget *uctl_make(Ctl *c, QWidget *parent)
{
    const vb_class *cls = uctl_class(c->def ? c->def->cls : c->type.c_str());
    if (!cls) return new QWidget(parent);
    vbobj uc = vb_onew(cls);                    /* UserControl_Initialize */
    if (!uc) return new QWidget(parent);
    vb_oaddref(uc);
    c->uc = uc;
    uctl_load(uc);
    FormState *us = fstate(uc);
    us->host = c;
    QWidget *w = us->w;
    w->setParent(parent);
    w->show();
    return w;
}

/* the PropertyBag of ReadProperties: the properties of the control in the .frm */
struct PropBag {
    const char *const *props = nullptr;         /* the .frm */
    VFont font;
    std::map<std::string, vbvar> written;       /* WriteProperties (Load of a control array element) */
    ~PropBag() { for (auto &kv : written) vb_vfree(&kv.second); }
};
/* Contents: the written properties as a Byte array (to keep in a file, and give back to a new bag).
 * The bytes are LinBASIC's own: "LBPB" 1, the number of properties, then each one: its name (UTF-8),
 * its VarType and its value - numbers, dates, Booleans, Strings and Byte arrays. */
static void bag_put32(QByteArray &o, vb_u32 v) { o.append((const char *)&v, 4); }
static vbvar bag_contents(PropBag *b, vb_i32 kind, vb_i32 argc, vbvar *argv)
{
    if (kind == VBM_LET) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        vbarr a = { 0 };
        vb_v2arr(argv[argc - 1], &a, VT_UI1);
        if (vb_errpend) return vb_v_empty();
        QByteArray in(a.data, a.n);
        vb_afree(&a);
        for (auto &kv : b->written) vb_vfree(&kv.second);
        b->written.clear();
        int at = 0;
        auto get32 = [&](vb_u32 *v) { if (at + 4 > in.size()) return false; memcpy(v, in.constData() + at, 4); at += 4; return true; };
        vb_u32 magic, cnt;
        if (!get32(&magic) || magic != 0x4250424C || !get32(&cnt)) { vb_raise(5); return vb_v_empty(); }   /* "LBPB" */
        for (vb_u32 i = 0; i < cnt; i++) {
            vb_u32 kl, vt;
            if (!get32(&kl) || at + (int)kl > in.size()) break;
            std::string key(in.constData() + at, kl); at += kl;
            if (!get32(&vt)) break;
            vbvar v = vb_v_empty();
            if (vt == VT_BSTR || vt == (VT_ARRAY | VT_UI1)) {
                vb_u32 n;
                if (!get32(&n) || at + (int)n > in.size()) break;
                if (vt == VT_BSTR) {
                    vbs t = vb_stmp((vb_i32)(n / 2));
                    memcpy(t.p, in.constData() + at, n);
                    v = vb_v_s(t);
                } else {
                    vbarr r = { 0 };
                    if (n) { vb_aredim(&r, 1, VBA_PLAIN, 0, 1, 0, (vb_i32)n - 1); memcpy(r.data, in.constData() + at, n); }
                    else { r.nd = 1; r.esz = 1; r.kind = VBA_PLAIN; r.ub[0] = -1; r.mul[0] = 1; }
                    v = vb_v_arr(&r, VT_UI1);
                    vb_afree(&r);
                }
                at += n;
            } else if (vt != VT_EMPTY && vt != VT_NULL) {
                if (at + 8 > in.size()) break;
                memcpy(&v.u, in.constData() + at, 8); at += 8;
                v.vt = (vb_i16)vt;
            } else v.vt = (vb_i16)vt;
            vbvar &slot = b->written[key];
            slot = vb_v_empty();
            vb_vset(&slot, v);
        }
        return vb_v_empty();
    }
    QByteArray o;
    bag_put32(o, 0x4250424C); bag_put32(o, 0);
    vb_u32 cnt = 0;
    vb_i32 m = vb_mark();
    for (auto &kv : b->written) {
        const vbvar &v = kv.second;
        vb_u32 vt = (vb_u32)v.vt;
        if (vt == VT_DISPATCH || ((vt & VT_ARRAY) && vt != (VT_ARRAY | VT_UI1))) continue;   /* objects, other arrays: not kept */
        bag_put32(o, (vb_u32)kv.first.size()); o.append(kv.first.data(), (int)kv.first.size());
        bag_put32(o, vt);
        if (vt == VT_BSTR) { bag_put32(o, (vb_u32)v.u.s.len * 2); o.append((const char *)v.u.s.p, v.u.s.len * 2); }
        else if (vt == (VT_ARRAY | VT_UI1)) {
            vbarr a = { 0 };
            vb_v2arr(v, &a, VT_UI1);
            bag_put32(o, (vb_u32)a.n); o.append(a.data, a.n);
            vb_afree(&a);
        } else if (vt != VT_EMPTY && vt != VT_NULL) o.append((const char *)&v.u, 8);
        cnt++;
    }
    vb_release(m);
    memcpy(o.data() + 4, &cnt, 4);
    vbarr r = { 0 };
    vb_aredim(&r, 1, VBA_PLAIN, 0, 1, 0, o.size() - 1);
    memcpy(r.data, o.constData(), (size_t)o.size());
    vbvar res = vb_v_arr(&r, VT_UI1);
    vb_afree(&r);
    return res;
}
static vbvar bag_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    PropBag *b = (PropBag *)vb_oid(o)->ext;
    std::string n = lower(name);
    *found = 1;
    if (n == "writeproperty") {                 /* a value equal to its default is not written */
        if (argc < 2) { vb_raise(450); return vb_v_empty(); }
        std::string key = lower(vb_stoutf8(vb_v2s(argv[0])));
        vbvar v = argv[1];
        if (argc > 2 && !vb_ismissing(argv[2]) && v.vt != VT_DISPATCH && argv[2].vt != VT_DISPATCH) {
            vb_i32 m = vb_mark();
            bool same = vb_stoutf8(vb_v2s(v)) == std::string(vb_stoutf8(vb_v2s(argv[2])));
            vb_release(m);
            if (same) { auto it = b->written.find(key); if (it != b->written.end()) { vb_vfree(&it->second); b->written.erase(it); } return vb_v_empty(); }
        }
        vbvar &slot = b->written[key];
        vb_vset(&slot, v);
        return vb_v_empty();
    }
    if (n == "contents" && b) return bag_contents(b, kind, argc, argv);
    if (n != "readproperty") { *found = 0; return vb_v_empty(); }
    if (argc < 1) { vb_raise(450); return vb_v_empty(); }
    QString key = qs(vb_v2s(argv[0]));
    vbvar dflt = argc > 1 && !vb_ismissing(argv[1]) ? argv[1] : vb_v_empty();
    if (b) {
        auto it = b->written.find(lower(key.toUtf8().constData()));
        if (it != b->written.end()) { vbvar r = vb_v_empty(); vb_vset(&r, it->second); return vb_vtmp(r); }
    }
    if (!b || !b->props) return dflt;
    QByteArray k = key.toUtf8();
    /* a Font: its BeginProperty block (Font.Name ...) */
    QString v;
    if (prop(b->props, (k + ".Name").constData(), nullptr) || prop(b->props, (k + ".Size").constData(), nullptr)) {
        VFont f = b->font;
        const char *const *p = b->props;
        std::vector<const char *> sub;            /* Font.Name -> Name: font_from_props reads "Font." */
        std::vector<std::string> keep;
        for (int i = 0; p[i]; i += 2) {
            if (!strncasecmp(p[i], k.constData(), (size_t)k.size()) && p[i][k.size()] == '.') {
                keep.push_back(std::string("Font.") + (p[i] + k.size() + 1));
            }
        }
        for (size_t i = 0, j = 0; p[i]; i += 2)
            if (!strncasecmp(p[i], k.constData(), (size_t)k.size()) && p[i][k.size()] == '.') { sub.push_back(keep[j++].c_str()); sub.push_back(p[i + 1]); }
        sub.push_back(nullptr);
        font_from_props(&f, sub.data());
        vbobj fo = vb_onew(&CLS_FONT);
        ((FontRef *)fo->ext)->own = f;
        return vb_vtmp(vb_v_obj(fo));
    }
    for (int i = 0; b->props[i]; i += 2) {
        if (strcasecmp(b->props[i], k.constData())) continue;
        const char *raw = b->props[i + 1];
        QString s = QString::fromUtf8(raw + 1);
        switch (raw[0]) {
        case 's': case 'l': return vstr(s);
        case 'x': {                               /* a picture from the .frx */
            QImage img = frx_image(b->props, k.constData());
            return vb_vtmp(vb_v_obj(new_pic(img, img.isNull() ? 0 : 1)));
        }
        case 'n': {
            bool ok; qlonglong l = s.toLongLong(&ok);
            if (ok && l >= INT32_MIN && l <= INT32_MAX) return vb_v_l((vb_i32)l);
            if (ok) return vb_v_d((double)l);
            return vb_v_d(s.toDouble());
        }
        default: {                                /* r: as the .frm has it (-1 'True, &H8000000F& ...) */
            QString t = s.trimmed();
            if (t.startsWith("&H", Qt::CaseInsensitive)) {
                bool ok; qulonglong u = t.mid(2).remove('&').toULongLong(&ok, 16);
                if (ok) return vb_v_l((vb_i32)(vb_u32)u);
            }
            bool ok; double d = t.toDouble(&ok);
            if (ok) return d == (vb_i32)d ? vb_v_l((vb_i32)d) : vb_v_d(d);
            return vstr(s);
        }
        }
    }
    return dflt;
}
static void bag_clear(vbobj o) { delete (PropBag *)o->ext; o->ext = nullptr; }
static const vb_class CLS_BAG = { "PropertyBag", (vb_i32)sizeof(vbobjh), nullptr, nullptr, bag_clear, nullptr, 0, nullptr,
                                  nullptr, nullptr, nullptr, bag_member, nullptr };
/* New PropertyBag: a bag of the program's own (WriteProperty, ReadProperty, Contents) */
static void bag_init(vbobj o) { if (!o->ext) o->ext = new PropBag; }
static const vb_class CLS_NEWBAG = { "PropertyBag", (vb_i32)sizeof(vbobjh), bag_init, nullptr, bag_clear, nullptr, 0, nullptr,
                                     nullptr, nullptr, nullptr, bag_member, nullptr };
extern "C" const vb_class *vbf_cls_propertybag(void) { return &CLS_NEWBAG; }

/* after the Extender has its place: Resize, then ReadProperties with the properties of the .frm
 * (a new control of Controls.Add: InitProperties) */
static bool uctl_invisible(Ctl *c);
static void uctl_after(Ctl *c)
{
    if (!c->uc) return;
    if (uctl_invisible(c) && c->w) c->w->hide();
    uctl_size(c, false);                        /* its size, without the event yet */
    c->uc_loading = true;
    if (c->bag || (c->def && c->def->props && !c->added)) {
        vbobj bag = c->bag;
        c->bag = nullptr;
        if (!bag) {
            bag = vb_onew(&CLS_BAG);
            PropBag *b = new PropBag;
            b->props = c->def->props;
            b->font = c->form ? c->form->font : VFont();
            bag->ext = b;
        } else vb_otmp(bag);                    /* the reference it had: the expression's now */
        vbvar a[1] = { vb_v_obj(bag) };
        vb_i32 m = vb_mark();
        fire_form(c->uc, "readproperties", 1, a);
        vb_release(m);
    } else {
        fire_form(c->uc, "initproperties");
    }
    c->uc_loading = false;
    c->uc_ready = true;
    /* VB6 (oracles f_uctl, f_uctl2): Resize comes after ReadProperties / InitProperties, or in them
     * when their code asks the Ambient (the control gets its place then) */
    if (!vb_errpend) uctl_size(c, true);
}

static void uctl_geometry(Ctl *c) { if (c->uc && c->uc_ready && c->form && c->form->loaded) uctl_size(c, true); }

/* Show / Hide: the Extender becomes visible (on a visible form) or not */
static bool uctl_invisible(Ctl *c)
{
    return c->uc && propn(fstate(c->uc)->def->props, "InvisibleAtRuntime", 0) != 0;
}
static bool uctl_on_screen(Ctl *c)
{
    if (uctl_invisible(c)) return false;        /* like a Timer: never on the screen */
    bool on = c->visible && c->form && c->form->visible;
    for (vbobj p = c->container; on && p; p = C(p)->container) on = C(p)->visible;
    return on;
}
static void uctl_visible(Ctl *c)
{
    if (!c->uc) return;
    bool on = uctl_on_screen(c);
    if (on == c->uc_shown) return;
    if (!on) { c->uc_shown = false; fire_form(c->uc, "hide"); return; }   /* Hide at once */
    vbobj co = c->self;                         /* Show from the message queue (t: f_uctl) */
    vb_oaddref(co);
    post([co]() {
        Ctl *k = C(co);
        if (k && k->uc && !k->uc_shown && uctl_on_screen(k)) { k->uc_shown = true; fire_form(k->uc, "show"); }
        vb_orelease(co);
    });
}

/* Load Counter2(1): the WriteProperties of the element it copies, for the ReadProperties of the new one */
static vbobj uctl_copy_bag(Ctl *src)
{
    if (!src->uc) return nullptr;
    vbobj bag = vb_onew(&CLS_BAG);
    PropBag *b = new PropBag;
    b->font = src->form ? src->form->font : VFont();
    bag->ext = b;
    vb_oaddref(bag);
    vbvar a[1] = { vb_v_obj(bag) };
    vb_i32 m = vb_mark();
    fire_form(src->uc, "writeproperties", 1, a);
    vb_release(m);
    return bag;
}
static void uctl_shown(FormState *st)
{
    std::vector<vbobj> ucs;
    for (vbobj co : st->ctls) { Ctl *c = C(co); if (c && c->uc && !c->uc_shown && uctl_on_screen(c)) { vb_oaddref(co); ucs.push_back(co); } }
    if (ucs.empty()) return;
    post([ucs]() {
        for (vbobj co : ucs) {
            Ctl *c = C(co);
            if (c && c->uc && !c->uc_shown && uctl_on_screen(c)) { c->uc_shown = true; fire_form(c->uc, "show"); }
            vb_orelease(co);
        }
    });
}

/* the control goes (Unload of the form or of the element): Hide, then the object (Terminate) */
static void uctl_release(Ctl *c)
{
    vbobj uc = c->uc;
    if (!uc) return;
    c->uc = nullptr;
    FormState *ust = fstate(uc);
    if (ust->focus && !c->form) {               /* the form goes: one of its controls had the focus once (f_uctl2) */
        ust->focus = nullptr;
        vb_i32 m = vb_mark(); fire_form(uc, "lostfocus"); vb_release(m);
    }
    if (c->uc_shown) { c->uc_shown = false; vb_i32 m = vb_mark(); fire_form(uc, "hide"); vb_release(m); }
    FormState *us = fstate(uc);
    us->host = nullptr;
    /* its form went: the Timers inside stop and the drawing surface sends no more Paint events */
    for (vbobj co : us->ctls) { Ctl *ic = C(co); if (ic && ic->timer) ic->timer->stop(); }
    if (us->canvas) { us->canvas->on_paint = nullptr; us->canvas->on_resize = nullptr; }
    vb_orelease(uc);
}

/* the members of the control on its form: the Extender's first, then the user control's own
 * (its public members; others: 438, as VB6 has them) */
static bool extender_name(const std::string &n)
{
    static const std::set<std::string> ext = {
        "name", "index", "tag", "parent", "left", "top", "width", "height", "visible", "move", "zorder",
        "setfocus", "tabindex", "tabstop", "container", "tooltiptext", "dragicon", "dragmode", "drag",
        "helpcontextid", "whatsthishelpid", "cancel", "default", "causesvalidation", "object",
        "datafield", "datamember", "datasource", "datachanged", "dataformat", "showwhatsthis" };
    return ext.count(n) != 0;
}
static vbvar uctl_member(Ctl *c, const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)set; (void)val;
    vbobj uc = c->uc;
    if (!uc) { *found = 0; return vb_v_empty(); }
    if (n == "object") return vb_vtmp(vb_v_obj(uc));
    /* an Extender property wins over the control's own of the same name (VB6: Extender.ToolTipText inside
     * Property Get ToolTipText is the container's, not an endless recursion) */
    if (extender_name(n)) { *found = 0; return vb_v_empty(); }
    const vb_class *cls = uc->cls;
    std::string name = n.empty() && cls->defmember ? std::string(cls->defmember) : n;
    vb_i32 want = kind == VBM_GET ? (VBM_GET | VBM_METHOD) : kind == VBM_METHOD ? (VBM_METHOD | VBM_GET) : kind;
    for (vb_i32 i = 0; i < cls->nmembers; i++)
        if ((cls->members[i].kind & want) && name == cls->members[i].name)
            return vb_late_o(uc, name.c_str(), kind, argc, argv);
    if (n.empty()) { *found = 0; return vb_v_empty(); }
    /* another member of the same name and the other kind (a read-only property set ...) */
    for (vb_i32 i = 0; i < cls->nmembers; i++)
        if (name == cls->members[i].name) { vb_raise(set ? 383 : 394); return vb_v_empty(); }
    vb_raise(438);
    return vb_v_empty();
}

/* Ambient: what the form that has the user control says about itself */
static vbvar ambient_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)argc; (void)argv;
    vbobj uc = (vbobj)vb_oid(o)->ext;
    std::string n = lower(name);
    *found = 1;
    if (kind == VBM_LET || kind == VBM_SET) { vb_raise(383); return vb_v_empty(); }
    Ctl *c = uc ? fstate(uc)->host : nullptr;
    FormState *hf = c ? c->form : nullptr;
    if (n == "usermode") return vb_v_bool(-1);
    if (n == "displayname") {
        if (!c) return vstr(QString());
        return vstr(c->index >= 0 ? QString("%1(%2)").arg(c->name).arg(c->index) : c->name);
    }
    if (n == "backcolor") return vb_v_l((vb_i32)(c && c->container ? C(c->container)->back : hf ? hf->g.back : 0x8000000F));
    if (n == "forecolor") return vb_v_l((vb_i32)(c && c->container ? C(c->container)->fore : hf ? hf->g.fore : 0x80000012));
    if (n == "font") {
        vbobj fo = vb_onew(&CLS_FONT);
        ((FontRef *)fo->ext)->own = hf ? hf->font : VFont();
        return vb_vtmp(vb_v_obj(fo));
    }
    if (n == "localeid") return vb_v_l(vb_lcid);
    if (n == "textalign") return vb_v_i(0);
    if (n == "scaleunits") return vstr(QString("Twip"));
    if (n == "appearance") return vb_v_i(1);
    if (n == "displayasdefault" || n == "showgrabhandles" || n == "showhatching" || n == "righttoleft")
        return vb_v_bool(0);
    if (n == "messagereflect" || n == "supportsmnemonics") return vb_v_bool(-1);
    if (n == "palette") return vb_vtmp(vb_v_obj(nullptr));
    *found = 0;
    return vb_v_empty();
}
static const vb_class CLS_AMBIENT = { "AmbientProperties", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr,
                                      nullptr, nullptr, nullptr, ambient_member, nullptr };

/* UserControl.Extender, .Ambient, .Parent ... inside the user control's own code; the members of
 * forms that a user control does not have: 438 */
static vbvar uctl_self_member(vbobj o, const std::string &n, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)argc; (void)argv;
    FormState *st = fstate(o);
    Ctl *c = st->host;
    *found = 1;
    if (n == "extender") {
        if (!c) { vb_raise(398); return vb_v_empty(); }   /* Client Site not available */
        return vb_vtmp(vb_v_obj(c->self));
    }
    if (n == "ambient") {
        if (!c) { vb_raise(398); return vb_v_empty(); }
        if (c->uc_loading && !c->uc_sized) uctl_size(c, true);   /* VB6: its Resize comes now */
        vbobj a = vb_onew(&CLS_AMBIENT);
        a->ext = o;                             /* weak: the user control outlives its Ambient object */
        return vb_vtmp(vb_v_obj(a));
    }
    if (n == "parent") {
        if (!c || !c->form) { vb_raise(398); return vb_v_empty(); }
        return vb_vtmp(vb_v_obj(c->form->form));
    }
    if (n == "containedcontrols") {             /* ControlContainer = True: the controls of the form in it */
        if (!c || propn(st->def->props, "ControlContainer", 0) == 0) { vb_raise(438); return vb_v_empty(); }
        std::vector<vbobj> in;
        if (c->form)
            for (vbobj co : c->form->ctls) if (C(co)->container == c->self) in.push_back(co);
        if (argc > 0) return ctl_list_item(in, argc, argv);   /* ContainedControls(i): from 0, as VB6 */
        return vb_vtmp(vb_v_obj(make_coll(in)));
    }
    if (n == "parentcontrols") {
        if (!c || !c->form) { vb_raise(398); return vb_v_empty(); }
        if (argc > 0) return ctl_list_item(c->form->ctls, argc, argv);
        return vb_vtmp(vb_v_obj(make_coll(c->form->ctls)));
    }
    if (n == "propertychanged") { if (argc > 0) drep_prop_changed(o, argv[0]); return vb_v_empty(); }
    if (n == "canpropertychange") return vb_v_bool(-1);
    if (n == "name") return vstr(QString::fromUtf8(st->def->name));
    if (n == "width" || n == "height") {
        double v = c ? (n == "width" ? c->width : c->height) : (n == "width" ? st->clientw : st->clienth);
        if (!set) return vb_v_f((float)v);
        if (c) { (n == "width" ? c->width : c->height) = vb_v2d(val); ctl_geometry(c); }
        return vb_v_empty();
    }
    if (n == "size") {
        if (argc < 2) { vb_raise(450); return vb_v_empty(); }
        if (c) { c->width = vb_v2d(argv[0]); c->height = vb_v2d(argv[1]); ctl_geometry(c); }
        return vb_v_empty();
    }
    if (n == "setfocus") {
        if (c && c->w) c->w->setFocus();
        return vb_v_empty();
    }
    if (n == "backstyle") return set ? vb_v_empty() : vb_v_i((vb_i16)propn(st->def->props, "BackStyle", 1));
    /* UserControl properties kept as VB6 has them (what they change is not drawn: the mask of a transparent
     * control, the access keys, the 3-D look, the border) */
    {
        static const std::map<std::string, std::pair<const char *, int>> keep = {
            { "maskcolor", { "MaskColor", 0xC0C0C0 } }, { "accesskeys", { "AccessKeys", -1 } }, { "appearance", { "Appearance", 1 } },
            { "borderstyle", { "BorderStyle", 0 } }, { "hitbehavior", { "HitBehavior", 1 } }, { "clipbehavior", { "ClipBehavior", 1 } },
            { "forwardfocus", { "ForwardFocus", 0 } }, { "cangetfocus", { "CanGetFocus", 1 } }, { "defaultcancel", { "DefaultCancel", 0 } },
            { "maskpicture", { "MaskPicture", -2 } } };
        auto k = keep.find(n);
        if (k != keep.end()) {
            vbvar &slot = st->extra[n];
            if (set) { vb_vset(&slot, val); return vb_v_empty(); }
            if (slot.vt != VT_EMPTY) return vb_vtmp(slot);
            int d = k->second.second;
            if (d == -1) { QString v; prop(st->def->props, k->second.first, &v); return vstr(v); }
            if (d == -2) return vb_vtmp(vb_v_obj(nullptr));
            if (n == "forwardfocus" || n == "cangetfocus" || n == "defaultcancel")
                return vb_v_bool(propn(st->def->props, k->second.first, d) ? -1 : 0);
            return n == "maskcolor" ? vb_v_l((vb_i32)propn(st->def->props, k->second.first, d))
                                    : vb_v_i((vb_i16)propn(st->def->props, k->second.first, d));
        }
    }
    if (n == "eventsfrozen") return vb_v_bool(0);
    static const std::set<std::string> no = { "show", "hide", "caption", "left", "top", "move", "windowstate", "visible",
                                              "icon", "zorder", "activeform", "arrange", "mdichild", "controlbox",
                                              "minbutton", "maxbutton", "showintaskbar", "moveable", "keypreview",
                                              "tag", "validatecontrols", "popupmenu", "activecontrol" };
    if (no.count(n)) { vb_raise(438); return vb_v_empty(); }
    *found = 0;
    return vb_v_empty();
}

/* RaiseEvent in the user control: the handler of its form (Counter1_Changed); ByRef arguments
 * come back in argv */
extern "C" void vbf_uctl_event(vbobj uc, const char *event, vb_i32 argc, vbvar *argv)
{
    uc = vb_oid(uc);
    if (!uc || !uc->ext) return;
    FormState *st = (FormState *)uc->ext;
    if (!st->host || !st->host->form) return;
    fire_ctl(st->host, event, argc, argv);
}

/* TypeOf Counter1 Is Counter, Set c = Counter1 (As Counter): the user control behind the control */
static vbobj uctl_inner(vbobj o)
{
    o = vb_oid(o);
    if (!o || o->cls->fallback != ctl_member) return nullptr;
    Ctl *c = C(o);
    return c ? c->uc : nullptr;
}

/* UserControl.Enabled, UserControl.Refresh ... in the user control's code: its built-in members,
 * not the public members of its class (which may have the same names) */
extern "C" vbvar vbf_uc_member(vbobj uc, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv)
{
    if (vb_errpend) return vb_v_empty();
    uc = vb_oid(uc);
    if (!uc) { vb_raise(91); return vb_v_empty(); }
    int found = 0;
    vbvar r = vbf_form_member(uc, name, kind, argc, argv, &found);
    if (!found && !vb_errpend) vb_raise(438);
    return r;
}

/* ---------------------------------------------------------------- focus */
static Ctl *uctl_first(FormState *us)
{
    Ctl *best = nullptr;
    for (vbobj o : us->ctls) { Ctl *k = C(o); if (can_focus(k) && (!best || k->tabindex < best->tabindex)) best = k; }
    return best;
}
/* it can have the focus: one of its controls can, or itself (CanGetFocus, no controls that can) */
static bool uctl_can_focus(Ctl *c)
{
    if (!c->uc || uctl_invisible(c)) return false;
    FormState *us = fstate(c->uc);
    if (uctl_first(us)) return true;
    return propn(us->def->props, "CanGetFocus", -1) != 0;
}
/* the control of the user control that gets the focus: the one that had it last, or the first in its
 * tab order; 0: the user control itself */
static Ctl *uctl_inner_focus(Ctl *c)
{
    if (!c->uc) return nullptr;
    FormState *us = fstate(c->uc);
    if (us->focus) { Ctl *k = C(us->focus); if (k && can_focus(k)) return k; }
    return uctl_first(us);
}
static QWidget *uctl_focus_widget(Ctl *c)
{
    Ctl *in = uctl_inner_focus(c);
    if (in) return focus_widget(in);
    FormState *us = c->uc ? fstate(c->uc) : nullptr;
    if (us && us->canvas) { us->canvas->setFocusPolicy(Qt::StrongFocus); return us->canvas; }
    return c->w;
}
/* Alt + a key of its AccessKeys: AccessKeyPress, then the focus goes to it */
static bool uctl_access_key(FormState *st, int ascii)
{
    for (vbobj co : st->ctls) {
        Ctl *c = C(co);
        if (c->kind != K_USER || !c->uc || !c->visible || !c->enabled) continue;
        QString keys;
        if (!prop(fstate(c->uc)->def->props, "AccessKeys", &keys)) continue;
        if (!keys.toLower().contains(QChar(ascii))) continue;
        vbvar a[1] = { vb_v_i((vb_i16)ascii) };
        vb_i32 m = vb_mark();
        fire_form(c->uc, "accesskeypress", 1, a);
        vb_release(m);
        if (can_focus(c) && st->focus != c->self) focus_to(st, c);
        return true;
    }
    return false;
}

