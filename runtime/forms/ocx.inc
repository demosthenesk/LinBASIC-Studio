/* ocx.inc - the ActiveX controls of VB6, written again on Qt (phase 4.2). Included by vbforms.cpp.
 * A control of an .ocx is a Ctl of kind K_OCX with an Ocx object: it makes the widget, reads the
 * .frm, and answers the members of its type; what every control has (Left, Visible, Font ...) stays
 * in ctl_member. The collections (Panels, Buttons, ListImages, Tabs ...) are OcxList objects of
 * OcxItem objects: an item has one VB object, which the list keeps while the item is in it.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

extern "C" {
extern const char *vb_lc_stime, *vb_lc_sdate;
int vb_fmt_pic(char *o, int cap, const char *pic, double date);
}

/* the run-time errors of the common controls */
static void ocx_error(vb_i32 n)
{
    const char *t = n == 35600 ? "Index out of bounds" : n == 35601 ? "Element not found" :
                    n == 35602 ? "Key is not unique in collection" : n == 35603 ? "Invalid key" :
                    n == 35605 ? "Control has been deleted" : n == 35606 ? "Control's collection has been modified" :
                    n == 35607 ? "Required argument is missing" :
                    n == 35617 ? "ImageList cannot be modified while another control is bound to it" :
                    "Invalid property value";
    vb_i32 m = vb_mark();
    vb_raise_full(n, vs(QString("MSComctlLib")), vs(QString(t)));
    vb_release(m);
}
/* sizes of the parts of a control (Panel.Width ...) are in the ScaleMode of the control's container */
static double from_twips_ctl(Ctl *c, double t) { return c ? from_twips(c, t, false, false) : t; }
static double to_twips_ctl(Ctl *c, double v) { return c ? to_twips(c, v, false, false) : v; }
static vbvar vbool(bool b) { return vb_v_bool(b ? -1 : 0); }
static vbvar retobj(vbobj o) { return vb_vtmp(vb_v_obj(o)); }
/* a property of an item in the .frm: "Panels.Panel3.Object.Width" or "Panels.Panel3.Width" */
static bool iprop(const char *const *p, const QString &base, const char *name, QString *out)
{
    QString a = base + "." + name, b = base + ".Object." + name;
    return prop(p, a.toUtf8().constData(), out) || prop(p, b.toUtf8().constData(), out);
}
static double ipropn(const char *const *p, const QString &base, const char *name, double dflt)
{
    QString v;
    if (!iprop(p, base, name, &v)) return dflt;
    bool ok; double d = v.toDouble(&ok);
    return ok ? d : dflt;
}

/* ---------------------------------------------------------------- items and collections */
struct OcxList;
struct Ocx;
struct OcxItem {
    vbobj obj = nullptr;
    OcxList *list = nullptr;                    /* 0: taken out of its collection */
    QString key;
    vbvar tag{};
    virtual ~OcxItem() { vb_vfree(&tag); }
    virtual vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) = 0;
};
struct OcxList {
    vbobj obj = nullptr;                        /* the collection object, made when it is first asked for */
    Ocx *owner = nullptr;                       /* the control it belongs to */
    std::vector<OcxItem *> items;
    const char *clsname = "Object", *itemcls = "Object", *itemdef = nullptr;
    std::function<OcxItem *(vb_i32 argc, vbvar *argv)> add;     /* Add: makes the item and inserts it */
    std::function<void()> changed;              /* the widget follows */
    std::function<void(OcxItem *)> removing;
    std::function<bool(OcxItem *)> remove_all_of;   /* Remove of a TreeView node: the node and its children */
    std::function<void()> on_clear;             /* Clear: the widget is emptied first */
    std::function<bool()> locked;               /* Remove/Clear refused (an ImageList in use: 35617) */
    ~OcxList();
    int index_of(const OcxItem *it) const { for (size_t i = 0; i < items.size(); i++) if (items[i] == it) return (int)i; return -1; }
    int find_key(const QString &k) const { for (size_t i = 0; i < items.size(); i++) if (!items[i]->key.isEmpty() && !QString::compare(items[i]->key, k, Qt::CaseSensitive)) return (int)i; return -1; }
    int find(vbvar ix, bool raise = true) const;
    bool insert(OcxItem *it, int pos, const QString &key);     /* pos: 0-based, -1 at the end */
    void remove(int i);
    void clear();
    vbobj object();
};
static vbvar item_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static void item_clear(vbobj o) { delete (OcxItem *)o->ext; o->ext = nullptr; }
static vbvar list_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static vb_i32 list_each_count(vbobj o) { OcxList *l = (OcxList *)vb_oid(o)->ext; return l ? (vb_i32)l->items.size() : 0; }
static vbvar list_each_item(vbobj o, vb_i32 i)
{
    OcxList *l = (OcxList *)vb_oid(o)->ext;
    if (!l || i < 0 || i >= (vb_i32)l->items.size()) return vb_v_empty();
    return retobj(l->items[i]->obj);
}
/* one vb_class for each name (TypeName): Node, ListItem, Panel ... */
static const vb_class *ocx_class(const char *name, const char *def, bool list)
{
    static std::map<std::string, vb_class *> m;
    std::string k = std::string(list ? "L:" : "I:") + name;
    auto it = m.find(k);
    if (it != m.end()) return it->second;
    vb_class *c = new vb_class();
    c->name = strdup(name); c->size = (vb_i32)sizeof(vbobjh);
    c->clear = list ? nullptr : item_clear;
    c->defmember = def ? strdup(def) : nullptr;
    c->fallback = list ? list_fallback : item_fallback;
    if (list) { c->each_count = list_each_count; c->each_item = list_each_item; }
    m[k] = c;
    return c;
}
vbobj OcxList::object()
{
    if (!obj) { obj = vb_onew(ocx_class(clsname, "item", true)); obj->ext = this; vb_oaddref(obj); }
    return obj;
}
int OcxList::find(vbvar ix, bool raise) const
{
    if (ix.vt == VT_DISPATCH && ix.u.o) {       /* an item object itself */
        OcxItem *it = (OcxItem *)vb_oid(ix.u.o)->ext;
        int i = it ? index_of(it) : -1;
        if (i < 0 && raise) ocx_error(35601);
        return i;
    }
    if (ix.vt == VT_BSTR) {
        int i = find_key(qs(ix.u.s));
        if (i < 0 && raise) ocx_error(35601);
        return i;
    }
    vb_i32 n = vb_v2l(ix);
    if (vb_errpend) return -1;
    if (n < 1 || n > (vb_i32)items.size()) { if (raise) ocx_error(35600); return -1; }
    return n - 1;
}
bool OcxList::insert(OcxItem *it, int pos, const QString &key)
{
    if (!key.isEmpty()) {
        bool num = false; key.toDouble(&num);
        if (num) { ocx_error(35603); delete it; return false; }     /* a key cannot be a number */
        if (find_key(key) >= 0) { ocx_error(35602); delete it; return false; }
    }
    it->key = key;
    it->list = this;
    it->obj = vb_onew(ocx_class(itemcls, itemdef, false));
    it->obj->ext = it;
    vb_oaddref(it->obj);                        /* the list's reference */
    if (pos < 0 || pos > (int)items.size()) pos = (int)items.size();
    items.insert(items.begin() + pos, it);
    return true;
}
void OcxList::remove(int i)
{
    if (i < 0 || i >= (int)items.size()) return;
    OcxItem *it = items[i];
    if (removing) removing(it);
    items.erase(items.begin() + i);
    it->list = nullptr;
    vbobj o = it->obj;
    if (changed) changed();
    vb_orelease(o);                             /* the item goes with its last reference */
}
void OcxList::clear()
{
    if (on_clear) on_clear();
    while (!items.empty()) {
        OcxItem *it = items.back();
        if (removing) removing(it);
        items.pop_back();
        it->list = nullptr;
        vb_orelease(it->obj);
    }
    if (changed) changed();
}
OcxList::~OcxList()
{
    changed = nullptr; removing = nullptr; on_clear = nullptr;
    clear();
    if (obj) { obj->ext = nullptr; vb_orelease(obj); }
}
static vbvar list_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    OcxList *l = (OcxList *)vb_oid(o)->ext;
    std::string n = lower(name);
    (void)kind;
    *found = 1;
    if (!l) { ocx_error(35605); return vb_v_empty(); }
    if (n == "count") return vb_v_l((vb_i32)l->items.size());
    if (n == "item") {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = l->find(argv[0]);
        return i < 0 ? vb_v_empty() : retobj(l->items[i]->obj);
    }
    if (n == "add") {
        if (!l->add) { vb_raise(438); return vb_v_empty(); }
        OcxItem *it = l->add(argc, argv);
        if (!it) return vb_v_empty();
        if (l->changed) l->changed();
        return retobj(it->obj);
    }
    if ((n == "remove" || n == "clear") && l->locked && l->locked()) { ocx_error(35617); return vb_v_empty(); }
    if (n == "remove") {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = l->find(argv[0]);
        if (i >= 0) {
            if (l->remove_all_of) { l->remove_all_of(l->items[i]); if (l->changed) l->changed(); }
            else l->remove(i);
        }
        return vb_v_empty();
    }
    if (n == "clear") { l->clear(); return vb_v_empty(); }
    if (n == "_newenum") return retobj(o);
    *found = 0;
    return vb_v_empty();
}
/* Panels(1): the collection, then its item */
static vbvar list_ret(OcxList &l, vb_i32 argc, vbvar *argv)
{
    if (argc < 1) return retobj(l.object());
    int i = l.find(argv[0]);
    return i < 0 ? vb_v_empty() : retobj(l.items[i]->obj);
}
/* Add(Index, Key, ...): the place and the key of a new item */
static bool add_place(OcxList *l, vb_i32 argc, vbvar *argv, int *pos, QString *key)
{
    *pos = -1; key->clear();
    if (given(argc, argv, 0)) {
        vb_i32 i = vb_v2l(argv[0]);
        if (vb_errpend) return false;
        if (i < 1 || i > (vb_i32)l->items.size() + 1) { ocx_error(35600); return false; }
        *pos = i - 1;
    }
    if (given(argc, argv, 1)) *key = arg_s(argc, argv, 1);
    return !vb_errpend;
}
static vbvar item_fallback(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    OcxItem *it = (OcxItem *)vb_oid(o)->ext;
    std::string n = lower(name);
    bool set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (!it) { ocx_error(35605); return vb_v_empty(); }
    if (n == "index") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        if (!it->list) { ocx_error(35605); return vb_v_empty(); }
        return vb_v_l(it->list->index_of(it) + 1);
    }
    if (n == "key") {
        if (!set || !it->list) { if (set) it->key = qs(vb_v2s(val)); return set ? vb_v_empty() : vstr(it->key); }
        QString k = qs(vb_v2s(val));
        int j = k.isEmpty() ? -1 : it->list->find_key(k);
        if (j >= 0 && it->list->items[j] != it) { ocx_error(35602); return vb_v_empty(); }
        it->key = k;
        return vb_v_empty();
    }
    if (n == "tag") { if (set) { vb_vset(&it->tag, val); return vb_v_empty(); } return vb_vtmp(it->tag); }
    vbvar r = it->member(n, kind, set, val, argc, argv, found);
    if (*found && it->list && it->list->changed && set && !vb_errpend) it->list->changed();
    return r;
}

/* ---------------------------------------------------------------- the control part */
struct Ocx {
    Ctl *c = nullptr;
    virtual ~Ocx() {}
    virtual bool data_aware() { return false; }       /* its own DataSource / DataMember (data.inc) */
    virtual bool own_binding() { return false; }      /* follows its data source itself (DataGrid), not as a bound control */
    virtual const char *bound_prop() { return ""; }   /* the property a DataField binds (DataList: BoundText) */
    virtual bool focusable() { return false; }        /* SetFocus, the tab order, ActiveControl (MaskEdBox) */
    virtual void source_reset(vbobj src) { (void)src; }   /* an ADO Data control has another Recordset */
    /* a Data control (dc) is told that its recordset will move (Validate: action, *cancel stops it) and that
     * it moved (phase 1: the bound controls show the record, 2: after Reposition, 3: its rows changed and it
     * did not move); the DBGrid follows it so (phase 7.6) */
    virtual void data_before(vbobj dc, int action, int *cancel) { (void)dc; (void)action; (void)cancel; }
    virtual void data_after(vbobj dc, int phase) { (void)dc; (void)phase; }
    virtual QWidget *make(QWidget *parent) = 0;       /* 0: no window (ImageList, PictureClip) */
    virtual void load(const char *const *p) { (void)p; }
    virtual void shown() {}                     /* the form is shown (the CoolBar takes its children) */
    virtual void loaded() {}                    /* all the controls of the form exist (before Form_Load): an ImageList
                                                 * written after the control in the .frm can be found now */
    virtual void resized() {}                   /* its size changed (Align, Width ...) */
    virtual bool clicks() { return false; }     /* Click comes from the mouse events of the forms runtime */
    virtual bool inside(QWidget *w, QPoint lp) { return w->rect().contains(lp); }   /* where a release clicks */
    virtual bool eat_mouse(QMouseEvent *me) { (void)me; return false; }   /* its own (a double click on a gripper) */
    virtual vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
    {
        (void)n; (void)kind; (void)set; (void)val; (void)argc; (void)argv;
        *found = 0;
        return vb_v_empty();
    }
};

/* ---------------------------------------------------------------- ImageList */
struct ListImage : OcxItem {
    QImage img;
    vbobj pic = nullptr;
    ~ListImage() { if (pic) vb_orelease(pic); }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
struct OcxImageList : Ocx {
    OcxList images;
    int iw = 0, ih = 0;
    bool bound = false;                         /* a control has used it: images can be added, not removed */
    long mask = 0xC0C0C0, back = 0x80000005L;
    bool usemask = true;
    QWidget *make(QWidget *) override { return nullptr; }
    void load(const char *const *p) override
    {
        iw = (int)propn(p, "ImageWidth", 0); ih = (int)propn(p, "ImageHeight", 0);
        mask = propc(p, "MaskColor", 0xC0C0C0); back = propc(p, "BackColor", 0x80000005L);
        usemask = propn(p, "UseMaskColor", -1) != 0;
        int n = (int)propn(p, "Images.NumListImages", 0);
        for (int i = 1; i <= n; i++) {
            QString base = QString("Images.ListImage%1").arg(i), key, tag;
            auto *li = new ListImage;
            li->img = frx_image(p, (base + ".Picture").toUtf8().constData());
            iprop(p, base, "Key", &key);
            if (iprop(p, base, "Tag", &tag)) { vbvar t = vb_v_s(vs(tag)); vb_vset(&li->tag, t); }
            if (!images.insert(li, -1, key)) break;
            if (!iw && !li->img.isNull()) { iw = li->img.width(); ih = li->img.height(); }
        }
    }
    QPixmap pixmap(const ListImage *li) const
    {
        if (!li || li->img.isNull()) return QPixmap();
        QImage im = li->img.convertToFormat(QImage::Format_ARGB32);
        if (usemask && !li->img.hasAlphaChannel()) {
            QRgb m = qcolor(mask).rgb() & 0xFFFFFF;
            for (int y = 0; y < im.height(); y++) {
                QRgb *row = (QRgb *)im.scanLine(y);
                for (int x = 0; x < im.width(); x++) if ((row[x] & 0xFFFFFF) == m) row[x] = 0;
            }
        }
        return QPixmap::fromImage(im);
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
static OcxImageList *imagelist_of(Ctl *owner, vbobj *ref, const QString &name);
vbvar ListImage::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    if (n == "picture") {
        if (set) {
            const QImage *im = image_of_value(val);
            if (!im) { vb_raise(481); return vb_v_empty(); }
            img = *im;
            if (pic) { vb_orelease(pic); pic = nullptr; }
            return vb_v_empty();
        }
        if (!pic) { pic = new_pic(img, 1); vb_oaddref(pic); }
        return retobj(pic);
    }
    if (n == "extracticon") return retobj(new_pic(img, 3));
    if (n == "draw") return vb_v_empty();       /* Draw hDC, x, y: the controls draw their pictures themselves */
    *found = 0;
    return vb_v_empty();
}
vbvar OcxImageList::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    if (n == "listimages") return list_ret(images, argc, argv);
    if (n == "imagewidth" || n == "imageheight") {
        int *v = n == "imagewidth" ? &iw : &ih;
        if (!set) return vb_v_i((vb_i16)*v);
        vb_i32 nv = vb_v2i(val);
        if (vb_errpend || nv == *v) return vb_v_empty();
        if (!images.items.empty()) { vb_raise_full(35611, vs(QString("ImageList")), vs(QString("Property is read-only if image list contains images"))); return vb_v_empty(); }
        *v = nv;
        return vb_v_empty();
    }
    if (n == "maskcolor") { if (!set) return vb_v_l((vb_i32)mask); mask = vb_v2l(val); return vb_v_empty(); }
    if (n == "usemaskcolor") { if (!set) return vbool(usemask); usemask = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "backcolor") { if (!set) return vb_v_l((vb_i32)back); back = vb_v2l(val); return vb_v_empty(); }
    if (n == "himagelist") return vb_v_l(0);
    if (n == "overlay") {
        int a = images.find(argc > 0 ? argv[0] : vb_v_empty()), b = images.find(argc > 1 ? argv[1] : vb_v_empty());
        if (a < 0 || b < 0) return vb_v_empty();
        QImage r = pixmap((ListImage *)images.items[a]).toImage();
        QPainter pt(&r); pt.drawPixmap(0, 0, pixmap((ListImage *)images.items[b])); pt.end();
        return retobj(new_pic(r, 1));
    }
    *found = 0;
    return vb_v_empty();
}
/* ListImages.Add [Index], [Key], Picture */
static OcxItem *listimage_add(OcxImageList *il, vb_i32 argc, vbvar *argv)
{
    int pos; QString key;
    if (!add_place(&il->images, argc, argv, &pos, &key)) return nullptr;
    const QImage *im = argc > 2 ? image_of_value(argv[2]) : nullptr;
    if (!im) { if (argc > 2 && given(argc, argv, 2)) vb_raise(481); else ocx_error(35607); return nullptr; }
    auto *li = new ListImage;
    li->img = *im;
    if (!il->iw) { il->iw = im->width(); il->ih = im->height(); }
    if (!il->images.insert(li, pos, key)) return nullptr;
    return li;
}
/* ImageList = "ImageList1" in the .frm, or Set x.ImageList = ImageList1 */
static OcxImageList *imagelist_of(Ctl *owner, vbobj *ref, const QString &name)
{
    if (!*ref && !name.isEmpty() && owner->form) {
        auto it = owner->form->byname.find(lower(name.toUtf8().constData()));
        if (it != owner->form->byname.end()) { *ref = it->second; vb_oaddref(*ref); }
    }
    Ctl *k = *ref ? C(*ref) : nullptr;
    OcxImageList *il = k && k->ocx ? dynamic_cast<OcxImageList *>(k->ocx) : nullptr;
    if (il) il->bound = true;
    return il;
}
/* Image = 3 or Image = "key" of an item: the picture from the image list */
static QPixmap image_from(OcxImageList *il, vbvar img)
{
    if (!il || img.vt == VT_EMPTY) return QPixmap();
    if (img.vt != VT_BSTR && vb_v2l(img) == 0) return QPixmap();
    int i = il->images.find(img, false);
    return i < 0 ? QPixmap() : il->pixmap((ListImage *)il->images.items[i]);
}
/* the ImageList property of a control: get gives the control, set takes one */
static vbvar imagelist_prop(Ctl *c, vbobj *ref, QString *name, bool set, vbvar val)
{
    if (!set) { imagelist_of(c, ref, *name); return *ref ? retobj(*ref) : retobj(nullptr); }
    if (*ref) { vb_orelease(*ref); *ref = nullptr; }
    name->clear();
    if (val.vt == VT_DISPATCH && val.u.o) {
        Ctl *k = C(val.u.o);
        if (!k || !k->ocx || !dynamic_cast<OcxImageList *>(k->ocx)) { vb_raise(380); return vb_v_empty(); }
        *ref = vb_oid(val.u.o); vb_oaddref(*ref);
    }
    return vb_v_empty();
}
/* a picture as the controls have it: Image, a key or an index of the image list */
static bool image_ok(OcxImageList *il, vbvar v)
{
    if (v.vt == VT_EMPTY || (v.vt != VT_BSTR && vb_v2l(v) == 0)) return true;
    if (!il) { vb_raise_full(35613, vs(QString("MSComctlLib")), vs(QString("ImageList must be initialized before it can be used"))); return false; }
    return il->images.find(v) >= 0;
}

/* ---------------------------------------------------------------- ProgressBar */
struct OcxProgress : Ocx {
    double min = 0, max = 100, value = 0;
    int orientation = 0, scrolling = 0;
    QPointer<QProgressBar> bar;
    QWidget *make(QWidget *parent) override
    {
        bar = new QProgressBar(parent);
        bar->setTextVisible(false);
        bar->setRange(0, 10000);
        return bar;
    }
    void sync()
    {
        if (!bar) return;
        bar->setOrientation(orientation ? Qt::Vertical : Qt::Horizontal);
        bar->setValue(max > min ? (int)std::lround((value - min) * 10000.0 / (max - min)) : 0);
    }
    void load(const char *const *p) override
    {
        min = propn(p, "Min", 0); max = propn(p, "Max", 100); value = propn(p, "Value", min);
        orientation = (int)propn(p, "Orientation", 0); scrolling = (int)propn(p, "Scrolling", 0);
        sync();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        if (n == "value" || n == "min" || n == "max") {
            double *v = n == "value" ? &value : n == "min" ? &min : &max;
            if (!set) return vb_v_f((float)*v);
            double d = vb_v2d(val);
            if (vb_errpend) return vb_v_empty();
            if (n == "value" && (d < min || d > max)) { vb_raise(380); return vb_v_empty(); }
            if (n == "min" && d >= max) { vb_raise(380); return vb_v_empty(); }
            if (n == "max" && d <= min) { vb_raise(380); return vb_v_empty(); }
            *v = d;
            if (value < min) value = min;
            if (value > max) value = max;
            sync();
            return vb_v_empty();
        }
        if (n == "orientation" || n == "scrolling") {
            int *v = n == "orientation" ? &orientation : &scrolling;
            if (!set) return vb_v_l(*v);
            *v = vb_v2i(val); sync();
            return vb_v_empty();
        }
        *found = 0;
        return vb_v_empty();
    }
};

/* ---------------------------------------------------------------- Slider */
struct OcxSlider : Ocx {
    QPointer<QSlider> s;
    int selstart = 0, sellength = 0, tickstyle = 0, tickfreq = 1;
    bool selrange = false;
    QWidget *make(QWidget *parent) override
    {
        s = new QSlider(Qt::Horizontal, parent);
        vbobj self = c->self;
        QObject::connect(s, &QSlider::sliderMoved, [self](int) { fire_ctl(C(self), "scroll"); });
        QObject::connect(s, &QSlider::valueChanged, [self](int) {
            Ctl *x = C(self);
            auto *o = x ? dynamic_cast<OcxSlider *>(x->ocx) : nullptr;
            if (o && o->s && !o->s->isSliderDown()) fire_ctl(x, "change");
        });
        QObject::connect(s, &QSlider::sliderReleased, [self]() { fire_ctl(C(self), "change"); });
        return s;
    }
    void ticks()
    {
        static const QSlider::TickPosition tp[] = { QSlider::TicksBelow, QSlider::TicksAbove, QSlider::TicksBothSides, QSlider::NoTicks };
        s->setTickPosition(tp[std::min(3, std::max(0, tickstyle))]);
        s->setTickInterval(std::max(1, tickfreq));
    }
    void load(const char *const *p) override
    {
        QSignalBlocker b(s);
        s->setOrientation(propn(p, "Orientation", 0) ? Qt::Vertical : Qt::Horizontal);
        s->setRange((int)propn(p, "Min", 0), (int)propn(p, "Max", 10));
        s->setValue((int)propn(p, "Value", 0));
        s->setSingleStep((int)propn(p, "SmallChange", 1));
        s->setPageStep((int)propn(p, "LargeChange", 5));
        tickstyle = (int)propn(p, "TickStyle", 0); tickfreq = (int)propn(p, "TickFrequency", 1);
        selrange = propn(p, "SelectRange", 0) != 0;
        selstart = (int)propn(p, "SelStart", 0); sellength = (int)propn(p, "SelLength", 0);
        ticks();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        int v = set ? vb_v2l(val) : 0;
        if (vb_errpend) return vb_v_empty();
        if (n == "value") {
            if (!set) return vb_v_l(s->value());
            s->setValue(std::min(s->maximum(), std::max(s->minimum(), v)));     /* VB6 takes the nearest end */
            return vb_v_empty();
        }
        if (n == "min") { if (!set) return vb_v_l(s->minimum()); s->setMinimum(v); return vb_v_empty(); }
        if (n == "max") { if (!set) return vb_v_l(s->maximum()); s->setMaximum(v); return vb_v_empty(); }
        if (n == "smallchange") { if (!set) return vb_v_l(s->singleStep()); s->setSingleStep(v); return vb_v_empty(); }
        if (n == "largechange") { if (!set) return vb_v_l(s->pageStep()); s->setPageStep(v); return vb_v_empty(); }
        if (n == "orientation") { if (!set) return vb_v_l(s->orientation() == Qt::Vertical); s->setOrientation(v ? Qt::Vertical : Qt::Horizontal); return vb_v_empty(); }
        if (n == "tickstyle") { if (!set) return vb_v_l(tickstyle); tickstyle = v; ticks(); return vb_v_empty(); }
        if (n == "tickfrequency") { if (!set) return vb_v_l(tickfreq); tickfreq = v; ticks(); return vb_v_empty(); }
        if (n == "selstart") { if (!set) return vb_v_l(selstart); selstart = v; return vb_v_empty(); }
        if (n == "sellength") { if (!set) return vb_v_l(sellength); sellength = v; return vb_v_empty(); }
        if (n == "selectrange") { if (!set) return vbool(selrange); selrange = v != 0; return vb_v_empty(); }
        if (n == "clearsel") { selstart = sellength = 0; return vb_v_empty(); }
        if (n == "getnumticks") return vb_v_l((s->maximum() - s->minimum()) / std::max(1, tickfreq) + 1);
        *found = 0;
        return vb_v_empty();
    }
};

/* ---------------------------------------------------------------- StatusBar */
struct Panel : OcxItem {
    QString text, tip;
    int style = 0, alignment = 0, autosize = 0, bevel = 1;
    double width = 1440, minwidth = 1440;       /* twips */
    int x = 0, w = 0;                           /* where it is drawn (pixels) */
    bool enabled = true, visible = true;
    QImage pic;
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
struct OcxStatusBar;
class StatusWidget : public QWidget {
public:
    OcxStatusBar *sb;
    StatusWidget(OcxStatusBar *o, QWidget *parent) : QWidget(parent), sb(o) { setMouseTracking(true); }
protected:
    void paintEvent(QPaintEvent *) override;
    void resizeEvent(QResizeEvent *e) override;
    void mouseReleaseEvent(QMouseEvent *e) override;
    void mouseDoubleClickEvent(QMouseEvent *e) override;
    bool event(QEvent *e) override;
};
struct OcxStatusBar : Ocx {
    OcxList panels;
    int style = 0;                              /* 0 sbrNormal, 1 sbrSimple */
    QString simpletext;
    bool showtips = true;
    QPointer<StatusWidget> wd;
    QTimer *clock = nullptr;
    ~OcxStatusBar() { delete clock; }
    QWidget *make(QWidget *parent) override { wd = new StatusWidget(this, parent); return wd; }
    QString shown(const Panel *p) const
    {
        char b[128];
        switch (p->style) {
        case 1: return "CAPS";
        case 2: return "NUM";
        case 3: return "INS";
        case 4: return "SCRL";
        case 5: vb_fmt_pic(b, sizeof b, vb_lc_stime, vb_now()); return QString::fromUtf8(b);
        case 6: vb_fmt_pic(b, sizeof b, vb_lc_sdate, vb_now()); return QString::fromUtf8(b);
        case 7: return "KANA";
        default: return p->text;
        }
    }
    void layout()
    {
        if (!wd) return;
        QFontMetrics fm(wd->font());
        int total = wd->width(), used = 0, springs = 0;
        for (OcxItem *i : panels.items) {
            auto *p = (Panel *)i;
            if (!p->visible) { p->w = 0; continue; }
            int minw = (int)std::lround(p->minwidth / TPP);
            int wpx = (int)std::lround(p->width / TPP);
            if (p->autosize == 2) wpx = std::max(minw, fm.horizontalAdvance(shown(p)) + 12 + (p->pic.isNull() ? 0 : p->pic.width() + 4));
            else if (p->autosize == 1) springs++;
            p->w = std::max(wpx, p->autosize == 1 ? minw : 0);
            used += p->w;
        }
        int extra = total - used - 2;
        if (springs && extra > 0) {
            int each = extra / springs, rest = extra - each * springs;
            for (OcxItem *i : panels.items) { auto *p = (Panel *)i; if (p->visible && p->autosize == 1) { p->w += each + (rest-- > 0 ? 1 : 0); } }
        }
        int x = 0;
        for (OcxItem *i : panels.items) {
            auto *p = (Panel *)i;
            p->x = x; x += p->w;
            if (p->visible) p->width = p->w * TPP;
        }
        bool timed = false;
        for (OcxItem *i : panels.items) if (((Panel *)i)->style == 5 || ((Panel *)i)->style == 6) timed = true;
        if (timed && !clock) {
            clock = new QTimer;
            QPointer<StatusWidget> w = wd;
            QObject::connect(clock, &QTimer::timeout, [w]() { if (w) w->update(); });
            clock->start(1000);
        }
        wd->update();
    }
    Panel *at(int x) const
    {
        if (style == 1) return nullptr;
        for (OcxItem *i : panels.items) { auto *p = (Panel *)i; if (p->visible && x >= p->x && x < p->x + p->w) return p; }
        return nullptr;
    }
    Panel *new_panel()
    {
        auto *p = new Panel;
        return p;
    }
    void load(const char *const *p) override
    {
        style = (int)propn(p, "Style", 0);
        prop(p, "SimpleText", &simpletext);
        showtips = propn(p, "ShowTips", -1) != 0;
        int n = (int)propn(p, "Panels.NumPanels", 1);
        for (int i = 1; i <= n; i++) {
            QString base = QString("Panels.Panel%1").arg(i), key, tag;
            Panel *pn = new_panel();
            iprop(p, base, "Text", &pn->text);
            iprop(p, base, "ToolTipText", &pn->tip);
            iprop(p, base, "Key", &key);
            if (iprop(p, base, "Tag", &tag)) { vbvar t = vb_v_s(vs(tag)); vb_vset(&pn->tag, t); }
            pn->style = (int)ipropn(p, base, "Style", 0);
            pn->alignment = (int)ipropn(p, base, "Alignment", 0);
            pn->autosize = (int)ipropn(p, base, "AutoSize", 0);
            pn->bevel = (int)ipropn(p, base, "Bevel", 1);
            pn->width = ipropn(p, base, "Width", 1440);
            pn->minwidth = ipropn(p, base, "MinWidth", pn->autosize ? 0 : 1440);
            pn->enabled = ipropn(p, base, "Enabled", -1) != 0;
            pn->visible = ipropn(p, base, "Visible", -1) != 0;
            pn->pic = frx_image(p, (base + ".Picture").toUtf8().constData());
            if (!panels.insert(pn, -1, key)) break;
        }
        layout();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        if (n == "panels") return list_ret(panels, argc, argv);
        if (n == "style") { if (!set) return vb_v_l(style); style = vb_v2i(val); layout(); return vb_v_empty(); }
        if (n == "simpletext") { if (!set) return vstr(simpletext); simpletext = qs(vb_v2s(val)); if (wd) wd->update(); return vb_v_empty(); }
        if (n == "showtips") { if (!set) return vbool(showtips); showtips = vb_v2bool(val) != 0; return vb_v_empty(); }
        *found = 0;
        return vb_v_empty();
    }
};
static OcxStatusBar *panel_owner(Panel *p) { return p->list ? dynamic_cast<OcxStatusBar *>(p->list->owner) : nullptr; }
vbvar Panel::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    if (n == "text") { if (!set) return vstr(text); text = qs(vb_v2s(val)); return vb_v_empty(); }
    if (n == "tooltiptext") { if (!set) return vstr(tip); tip = qs(vb_v2s(val)); return vb_v_empty(); }
    if (n == "style" || n == "alignment" || n == "autosize" || n == "bevel") {
        int *v = n == "style" ? &style : n == "alignment" ? &alignment : n == "autosize" ? &autosize : &bevel;
        if (!set) return vb_v_l(*v);
        int x = vb_v2i(val);
        if (x < 0 || x > (n == "style" ? 7 : 2)) { vb_raise(380); return vb_v_empty(); }
        *v = x;
        return vb_v_empty();
    }
    if (n == "width" || n == "minwidth") {
        double *v = n == "width" ? &width : &minwidth;
        if (!set) return vb_v_f((float)from_twips_ctl(panel_owner(this) ? panel_owner(this)->c : nullptr, *v));
        double d = to_twips_ctl(panel_owner(this) ? panel_owner(this)->c : nullptr, vb_v2d(val));
        if (d < 0) { vb_raise(380); return vb_v_empty(); }
        *v = d;
        if (n == "minwidth" && width < minwidth) width = minwidth;
        return vb_v_empty();
    }
    if (n == "left") { if (set) { vb_raise(383); return vb_v_empty(); } return vb_v_f((float)from_twips_ctl(panel_owner(this) ? panel_owner(this)->c : nullptr, x * TPP)); }
    if (n == "enabled") { if (!set) return vbool(enabled); enabled = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "visible") { if (!set) return vbool(visible); visible = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "picture") {
        if (!set) return retobj(new_pic(pic, 1));
        const QImage *im = image_of_value(val);
        pic = im ? *im : QImage();
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
void StatusWidget::paintEvent(QPaintEvent *)
{
    QPainter pt(this);
    pt.fillRect(rect(), palette().window());
    QColor dark(160, 160, 160), light(255, 255, 255);
    auto bevel = [&](QRect r, int b) {
        if (!b) return;
        pt.setPen(b == 1 ? dark : light); pt.drawLine(r.topLeft(), r.topRight()); pt.drawLine(r.topLeft(), r.bottomLeft());
        pt.setPen(b == 1 ? light : dark); pt.drawLine(r.bottomLeft(), r.bottomRight()); pt.drawLine(r.topRight(), r.bottomRight());
    };
    if (sb->style == 1) {
        QRect r = rect().adjusted(0, 2, -1, -1);
        bevel(r, 1);
        pt.setPen(palette().windowText().color());
        pt.drawText(r.adjusted(4, 0, -4, 0), Qt::AlignVCenter | Qt::AlignLeft, sb->simpletext);
        return;
    }
    for (OcxItem *i : sb->panels.items) {
        auto *p = (Panel *)i;
        if (!p->visible || p->w <= 0) continue;
        QRect r(p->x, 2, p->w - 2, height() - 3);
        bevel(r, p->bevel);
        QRect tr = r.adjusted(3, 0, -3, 0);
        if (!p->pic.isNull()) {
            pt.drawImage(tr.left(), r.top() + (r.height() - p->pic.height()) / 2, p->pic);
            tr.setLeft(tr.left() + p->pic.width() + 3);
        }
        bool key_style = p->style >= 1 && p->style <= 4;
        pt.setPen(!p->enabled || key_style ? palette().color(QPalette::Disabled, QPalette::WindowText) : palette().windowText().color());
        int al = p->alignment == 1 ? Qt::AlignHCenter : p->alignment == 2 ? Qt::AlignRight : Qt::AlignLeft;
        pt.drawText(tr, Qt::AlignVCenter | al, fontMetrics().elidedText(sb->shown(p), Qt::ElideRight, tr.width()));
    }
}
void StatusWidget::resizeEvent(QResizeEvent *e) { QWidget::resizeEvent(e); sb->layout(); }
void StatusWidget::mouseReleaseEvent(QMouseEvent *e)
{
    QWidget::mouseReleaseEvent(e);
    if (e->button() != Qt::LeftButton) return;
    Panel *p = sb->at((int)ev_pos(e).x());
    if (p) { vbvar a = vb_v_obj(p->obj); fire_ctl(sb->c, "panelclick", 1, &a); }
}
void StatusWidget::mouseDoubleClickEvent(QMouseEvent *e)
{
    QWidget::mouseDoubleClickEvent(e);
    Panel *p = sb->at((int)ev_pos(e).x());
    if (p) { vbvar a = vb_v_obj(p->obj); fire_ctl(sb->c, "paneldblclick", 1, &a); }
}
bool StatusWidget::event(QEvent *e)
{
    if (e->type() == QEvent::ToolTip && sb->showtips) {
        auto *he = static_cast<QHelpEvent *>(e);
        Panel *p = sb->at(he->pos().x());
        if (p && !p->tip.isEmpty()) QToolTip::showText(he->globalPos(), p->tip, this);
        else QToolTip::hideText();
        return true;
    }
    return QWidget::event(e);
}
/* Panels.Add [Index], [Key], [Text], [Style], [Picture] */
static OcxItem *panel_add(OcxStatusBar *sb, vb_i32 argc, vbvar *argv)
{
    int pos; QString key;
    if (!add_place(&sb->panels, argc, argv, &pos, &key)) return nullptr;
    Panel *p = sb->new_panel();
    p->text = arg_s(argc, argv, 2);
    p->style = (int)arg_l(argc, argv, 3, 0);
    if (argc > 4) { const QImage *im = image_of_value(argv[4]); if (im) p->pic = *im; }
    if (!sb->panels.insert(p, pos, key)) return nullptr;
    return p;
}

/* ---------------------------------------------------------------- Toolbar */
struct TbButton;
struct TbMenu : OcxItem {                       /* a ButtonMenu of a tbrDropdown button */
    QString text;
    bool enabled = true, visible = true;
    TbButton *button = nullptr;
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
struct TbButton : OcxItem {
    QString caption, tip, description;
    int style = 0, value = 0;                   /* tbrDefault 0, tbrCheck 1, tbrButtonGroup 2, tbrSeparator 3, tbrPlaceholder 4, tbrDropdown 5 */
    vbvar image{};
    bool enabled = true, visible = true, mixed = false;
    double width = 0;
    QPointer<QToolButton> btn;
    QPointer<QWidget> space;
    OcxList menus;                              /* ButtonMenus */
    TbButton();
    ~TbButton() { vb_vfree(&image); }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
struct OcxToolbar : Ocx {
    OcxList buttons;
    vbobj il = nullptr; QString ilname;
    double bw = 0, bh = 0;                      /* ButtonWidth, ButtonHeight (twips) */
    int style = 0, textalign = 0;
    bool showtips = true;
    QPointer<QWidget> bar;
    ~OcxToolbar() { if (il) vb_orelease(il); }
    void loaded() override { if (!ilname.isEmpty() && !il) sync(); }     /* the pictures of an ImageList after it */
    QWidget *make(QWidget *parent) override
    {
        bar = new QWidget(parent);
        /* a pressed button (Value = tbrPressed, a group, a check) stands out in a dark theme too (aw25) */
        bar->setStyleSheet("QToolButton:checked { background-color: palette(highlight); border: 1px solid palette(shadow);"
                           " border-radius: 2px; }");
        auto *lay = new QHBoxLayout(bar);
        lay->setContentsMargins(2, 2, 2, 2);
        lay->setSpacing(0);
        lay->addStretch(1);
        return bar;
    }
    void rebuild();
    void menu_sync(TbButton *b);
    void clicked(TbButton *b)
    {
        if (b->style == 1) b->value = b->value ? 0 : 1;
        else if (b->style == 2) {               /* a group: the buttons next to it go up */
            int i = buttons.index_of(b);
            for (int j = i - 1; j >= 0 && ((TbButton *)buttons.items[j])->style == 2; j--) ((TbButton *)buttons.items[j])->value = 0;
            for (int j = i + 1; j < (int)buttons.items.size() && ((TbButton *)buttons.items[j])->style == 2; j++) ((TbButton *)buttons.items[j])->value = 0;
            b->value = 1;
        }
        sync();
        vbvar a = vb_v_obj(b->obj);
        fire_ctl(c, "buttonclick", 1, &a);
    }
    void sync()
    {
        OcxImageList *list = imagelist_of(c, &il, ilname);
        for (OcxItem *i : buttons.items) {
            auto *b = (TbButton *)i;
            if (b->space) {
                b->space->setVisible(b->visible);
                if (b->style == 4) b->space->setFixedWidth(std::max(0, (int)std::lround(b->width / TPP)));
            }
            if (!b->btn) continue;
            QSignalBlocker bl(b->btn);
            b->btn->setVisible(b->visible);
            b->btn->setEnabled(b->enabled);
            b->btn->setText(b->caption);
            b->btn->setToolTip(showtips ? b->tip : QString());
            QPixmap px = image_from(list, b->image);
            b->btn->setIcon(px.isNull() ? QIcon() : QIcon(px));
            if (!px.isNull()) b->btn->setIconSize(px.size());
            b->btn->setToolButtonStyle(b->caption.isEmpty() ? Qt::ToolButtonIconOnly : px.isNull() ? Qt::ToolButtonTextOnly :
                                       textalign ? Qt::ToolButtonTextBesideIcon : Qt::ToolButtonTextUnderIcon);
            /* a plain button with Value = tbrPressed is drawn pressed too (frmDocument of the wizard: Bold) */
            b->btn->setCheckable(b->style == 1 || b->style == 2 || b->value != 0);
            b->btn->setChecked(b->value != 0);
            b->btn->setAutoRaise(style == 1);
            if (b->style == 5) menu_sync(b);
            if (bw > 0) b->btn->setMinimumWidth((int)std::lround(bw / TPP));
            if (bh > 0) b->btn->setMinimumHeight((int)std::lround(bh / TPP) - 0);
        }
    }
    void load(const char *const *p) override
    {
        prop(p, "ImageList", &ilname);
        /* the .frm keeps ButtonWidth / ButtonHeight in HIMETRIC like _ExtentX (VB6: Height 420 with ButtonHeight 582,
         * aw1 frmMain, f_comctl): in twips here */
        bw = propn(p, "ButtonWidth", 0) * 1440.0 / 2540.0; bh = propn(p, "ButtonHeight", 0) * 1440.0 / 2540.0;
        style = (int)propn(p, "Style", 0); textalign = (int)propn(p, "TextAlignment", 0);
        showtips = propn(p, "ShowTips", -1) != 0;
        int n = (int)propn(p, "Buttons.NumButtons", 0);
        for (int i = 1; i <= n; i++) {
            QString base = QString("Buttons.Button%1").arg(i), key, tag, img;
            auto *b = new TbButton;
            iprop(p, base, "Caption", &b->caption);
            iprop(p, base, "ToolTipText", &b->tip);
            iprop(p, base, "Description", &b->description);
            iprop(p, base, "Key", &key);
            if (iprop(p, base, "Tag", &tag)) { vbvar t = vb_v_s(vs(tag)); vb_vset(&b->tag, t); }
            b->style = (int)ipropn(p, base, "Style", 0);
            b->value = (int)ipropn(p, base, "Value", 0);
            b->enabled = ipropn(p, base, "Enabled", -1) != 0;
            b->visible = ipropn(p, base, "Visible", -1) != 0;
            b->mixed = ipropn(p, base, "MixedState", 0) != 0;
            b->width = ipropn(p, base, "Width", 0);
            /* Image as VB6 writes it: ImageIndex = n, or ImageKey = "key" (the Application Wizard: aw1 frmMain) */
            if ((iprop(p, base, "Image", &img) || iprop(p, base, "ImageKey", &img)) && !img.isEmpty()) {
                vbvar t = vb_v_s(vs(img)); vb_vset(&b->image, t);
            }
            int ix = (int)ipropn(p, base, "ImageIndex", 0);
            if (ix) { vbvar t = vb_v_l(ix); vb_vset(&b->image, t); }
            int nm = (int)ipropn(p, base, "ButtonMenus.NumButtonMenus", 0);
            for (int j = 1; j <= nm; j++) {
                QString mb = base + QString(".ButtonMenus.ButtonMenu%1").arg(j), mkey, mtag;
                auto *m = new TbMenu;
                m->button = b;
                iprop(p, mb, "Text", &m->text);
                iprop(p, mb, "Key", &mkey);
                if (iprop(p, mb, "Tag", &mtag)) { vbvar t = vb_v_s(vs(mtag)); vb_vset(&m->tag, t); }
                m->enabled = ipropn(p, mb, "Enabled", -1) != 0;
                m->visible = ipropn(p, mb, "Visible", -1) != 0;
                if (!b->menus.insert(m, -1, mkey)) break;
            }
            if (!buttons.insert(b, -1, key)) break;
        }
        rebuild();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        if (n == "buttons") return list_ret(buttons, argc, argv);
        if (n == "imagelist") { vbvar r = imagelist_prop(c, &il, &ilname, set, val); if (set) sync(); return r; }
        if (n == "buttonwidth" || n == "buttonheight") {
            double *v = n == "buttonwidth" ? &bw : &bh;
            if (!set) return vb_v_f((float)from_twips_ctl(c, *v));
            *v = to_twips_ctl(c, vb_v2d(val)); sync();
            return vb_v_empty();
        }
        if (n == "style") { if (!set) return vb_v_l(style); style = vb_v2i(val); sync(); return vb_v_empty(); }
        if (n == "textalignment") { if (!set) return vb_v_l(textalign); textalign = vb_v2i(val); sync(); return vb_v_empty(); }
        if (n == "showtips") { if (!set) return vbool(showtips); showtips = vb_v2bool(val) != 0; sync(); return vb_v_empty(); }
        if (n == "wrappable" || n == "allowcustomize") return set ? vb_v_empty() : vbool(n[0] == 'w');
        if (n == "customize" || n == "restoretoolbar" || n == "savetoolbar") return vb_v_empty();
        *found = 0;
        return vb_v_empty();
    }
};
static OcxToolbar *button_owner(TbButton *b) { return b->list ? dynamic_cast<OcxToolbar *>(b->list->owner) : nullptr; }
void OcxToolbar::rebuild()
{
    if (!bar) return;
    QLayout *lay = bar->layout();
    while (lay->count() > 1) {                  /* everything but the stretch at the end */
        QLayoutItem *li = lay->takeAt(0);
        if (li->widget()) li->widget()->deleteLater();
        delete li;
    }
    auto *hl = static_cast<QHBoxLayout *>(lay);
    int at = 0;
    for (OcxItem *i : buttons.items) {
        auto *b = (TbButton *)i;
        b->btn = nullptr; b->space = nullptr;
        if (b->style == 3 || b->style == 4) {
            auto *sp = new QWidget(bar);
            sp->setFixedWidth(b->style == 3 ? 8 : std::max(0, (int)std::lround(b->width / TPP)));
            b->space = sp;
            hl->insertWidget(at++, sp);
            continue;
        }
        auto *tb = new QToolButton(bar);
        tb->setFocusPolicy(Qt::NoFocus);
        if (b->style == 5) tb->setPopupMode(QToolButton::MenuButtonPopup);
        vbobj item = b->obj;
        vbobj self = c->self;
        QObject::connect(tb, &QToolButton::clicked, [item, self]() {
            Ctl *x = C(self);
            auto *o = x ? dynamic_cast<OcxToolbar *>(x->ocx) : nullptr;
            auto *bt = (TbButton *)vb_oid(item)->ext;
            if (o && bt && bt->list) o->clicked(bt);
        });
        b->btn = tb;
        hl->insertWidget(at++, tb);
    }
    sync();
}
vbvar TbButton::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    OcxToolbar *tb = button_owner(this);
    if (n == "buttonmenus") return list_ret(menus, argc, argv);
    /* a change shows at once (aw25: Value = tbrPressed from the SelChange of a RichTextBox) */
    auto shown = [tb]() { if (tb) tb->sync(); return vb_v_empty(); };
    if (n == "caption") { if (!set) return vstr(caption); caption = qs(vb_v2s(val)); return shown(); }
    if (n == "tooltiptext") { if (!set) return vstr(tip); tip = qs(vb_v2s(val)); return shown(); }
    if (n == "description") { if (!set) return vstr(description); description = qs(vb_v2s(val)); return vb_v_empty(); }
    if (n == "style") {
        if (!set) return vb_v_l(style);
        int s = vb_v2i(val);
        if (s < 0 || s > 5) { vb_raise(380); return vb_v_empty(); }
        style = s;
        if (tb) tb->rebuild();
        return vb_v_empty();
    }
    if (n == "value") {
        if (!set) return vb_v_l(value);
        value = vb_v2i(val) ? 1 : 0;
        if (value && style == 2 && tb) {        /* a group: the buttons of the group next to it go up */
            int i = tb->buttons.index_of(this);
            for (int j = i - 1; j >= 0 && ((TbButton *)tb->buttons.items[j])->style == 2; j--) ((TbButton *)tb->buttons.items[j])->value = 0;
            for (int j = i + 1; j < (int)tb->buttons.items.size() && ((TbButton *)tb->buttons.items[j])->style == 2; j++) ((TbButton *)tb->buttons.items[j])->value = 0;
        }
        return shown();
    }
    if (n == "image") {
        if (!set) return image.vt == VT_EMPTY ? vb_v_i(0) : vb_vtmp(image);     /* VB6: 0 without a picture (f_aw5) */
        if (tb && !image_ok(imagelist_of(tb->c, &tb->il, tb->ilname), val)) return vb_v_empty();
        vb_vset(&image, val);
        return shown();
    }
    if (n == "enabled") { if (!set) return vbool(enabled); enabled = vb_v2bool(val) != 0; return shown(); }
    if (n == "visible") { if (!set) return vbool(visible); visible = vb_v2bool(val) != 0; return shown(); }
    if (n == "mixedstate") { if (!set) return vbool(mixed); mixed = vb_v2bool(val) != 0; return shown(); }
    if (n == "width") {
        if (!set) {
            double t = btn ? btn->width() * TPP : space ? space->width() * TPP : width;
            return vb_v_f((float)from_twips_ctl(tb ? tb->c : nullptr, t));
        }
        width = to_twips_ctl(tb ? tb->c : nullptr, vb_v2d(val));
        return vb_v_empty();
    }
    if (n == "left" || n == "top" || n == "height") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        QWidget *w = btn ? (QWidget *)btn.data() : space.data();
        double t = !w ? 0 : n == "left" ? w->x() * TPP : n == "top" ? w->y() * TPP : w->height() * TPP;
        return vb_v_f((float)from_twips_ctl(tb ? tb->c : nullptr, t));
    }
    *found = 0;
    return vb_v_empty();
}
TbButton::TbButton()
{
    menus.clsname = "ButtonMenus"; menus.itemcls = "ButtonMenu"; menus.itemdef = "key";
    menus.add = [this](vb_i32 argc, vbvar *argv) -> OcxItem * {    /* ButtonMenus.Add [Index], [Key], [Text] */
        int pos; QString key;
        if (!add_place(&menus, argc, argv, &pos, &key)) return nullptr;
        auto *m = new TbMenu;
        m->button = this;
        m->text = arg_s(argc, argv, 2);
        if (!menus.insert(m, pos, key)) return nullptr;
        return m;
    };
    menus.changed = [this]() { if (OcxToolbar *tb = button_owner(this)) tb->sync(); };
}
vbvar TbMenu::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    if (n == "text") { if (!set) return vstr(text); text = qs(vb_v2s(val)); return vb_v_empty(); }
    if (n == "enabled") { if (!set) return vbool(enabled); enabled = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "visible") { if (!set) return vbool(visible); visible = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "parent") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        return button && list ? retobj(button->obj) : vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
/* the drop-down menu of a tbrDropdown button: its ButtonMenus; a choice fires ButtonMenuClick */
void OcxToolbar::menu_sync(TbButton *b)
{
    if (!b->btn) return;
    QMenu *m = b->btn->menu();
    if (!m) {
        m = new QMenu(b->btn);
        b->btn->setMenu(m);
        vbobj item = b->obj, self = c->self;
        QObject::connect(m, &QMenu::aboutToShow, [item, self]() {    /* ButtonDropDown first */
            Ctl *x = C(self);
            if (!x || !vb_oid(item)->ext) return;
            vbvar a = vb_v_obj(item);
            fire_ctl(x, "buttondropdown", 1, &a);
        });
    }
    m->clear();
    for (OcxItem *i : b->menus.items) {
        auto *bm = (TbMenu *)i;
        QAction *a = m->addAction(bm->text);
        a->setEnabled(bm->enabled);
        a->setVisible(bm->visible);
        vbobj mo = bm->obj, self = c->self;
        vb_oaddref(mo);
        QObject::connect(a, &QAction::destroyed, [mo]() { vb_orelease(mo); });
        QObject::connect(a, &QAction::triggered, [mo, self]() {
            Ctl *x = C(self);
            if (!x || !vb_oid(mo)->ext || !((TbMenu *)vb_oid(mo)->ext)->list) return;
            vbvar v = vb_v_obj(mo);
            fire_ctl(x, "buttonmenuclick", 1, &v);
        });
    }
}
/* Buttons.Add [Index], [Key], [Caption], [Style], [Image] */
static OcxItem *button_add(OcxToolbar *tb, vb_i32 argc, vbvar *argv)
{
    int pos; QString key;
    if (!add_place(&tb->buttons, argc, argv, &pos, &key)) return nullptr;
    auto *b = new TbButton;
    b->caption = arg_s(argc, argv, 2);
    b->style = (int)arg_l(argc, argv, 3, 0);
    if (given(argc, argv, 4)) {
        if (!image_ok(imagelist_of(tb->c, &tb->il, tb->ilname), argv[4])) { delete b; return nullptr; }
        vb_vset(&b->image, argv[4]);
    }
    if (!tb->buttons.insert(b, pos, key)) return nullptr;
    tb->rebuild();
    return b;
}

/* ---------------------------------------------------------------- TabStrip */
struct TabItem : OcxItem {
    QString caption, tip;
    vbvar image{};
    ~TabItem() { vb_vfree(&image); }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
class TabStripWidget : public QWidget {
public:
    QTabBar *bar;
    QFrame *page;
    explicit TabStripWidget(QWidget *parent) : QWidget(parent)
    {
        bar = new QTabBar(this);
        bar->setDrawBase(false);
        bar->setExpanding(false);
        page = new QFrame(this);
        page->setFrameStyle(QFrame::StyledPanel | QFrame::Raised);
        page->lower();
    }
protected:
    void resizeEvent(QResizeEvent *e) override
    {
        QWidget::resizeEvent(e);
        int h = bar->sizeHint().height();
        bar->setGeometry(0, 0, width(), h);
        page->setGeometry(0, h - 1, width(), height() - h + 1);
    }
};
struct OcxTabStrip : Ocx {
    OcxList tabs;
    vbobj il = nullptr; QString ilname;
    int style = 0;
    QPointer<TabStripWidget> wd;
    bool syncing = false;
    ~OcxTabStrip() { if (il) vb_orelease(il); }
    void loaded() override { if (!ilname.isEmpty() && !il) sync(); }
    QWidget *make(QWidget *parent) override
    {
        wd = new TabStripWidget(parent);
        vbobj self = c->self;
        QObject::connect(wd->bar, &QTabBar::currentChanged, [self](int) {
            Ctl *x = C(self);
            auto *o = x ? dynamic_cast<OcxTabStrip *>(x->ocx) : nullptr;
            if (o && !o->syncing) fire_ctl(x, "click");
        });
        return wd;
    }
    void sync()
    {
        if (!wd) return;
        syncing = true;
        QTabBar *b = wd->bar;
        int cur = b->currentIndex();
        while (b->count() > (int)tabs.items.size()) b->removeTab(b->count() - 1);
        while (b->count() < (int)tabs.items.size()) b->addTab(QString());
        OcxImageList *list = imagelist_of(c, &il, ilname);
        for (int i = 0; i < (int)tabs.items.size(); i++) {
            auto *t = (TabItem *)tabs.items[i];
            b->setTabText(i, t->caption);
            b->setTabToolTip(i, t->tip);
            QPixmap px = image_from(list, t->image);
            b->setTabIcon(i, px.isNull() ? QIcon() : QIcon(px));
        }
        if (cur >= b->count()) b->setCurrentIndex(b->count() - 1);
        syncing = false;
        wd->resize(wd->size());
        QResizeEvent re(wd->size(), wd->size());
        QApplication::sendEvent(wd, &re);
    }
    void load(const char *const *p) override
    {
        prop(p, "ImageList", &ilname);
        style = (int)propn(p, "Style", 0);
        int n = (int)propn(p, "Tabs.NumTabs", 1);
        for (int i = 1; i <= n; i++) {
            QString base = QString("Tabs.Tab%1").arg(i), key, tag, img;
            auto *t = new TabItem;
            if (!iprop(p, base, "Caption", &t->caption) && !propn(p, "Tabs.NumTabs", 0)) t->caption = "Tab 1";
            iprop(p, base, "ToolTipText", &t->tip);
            iprop(p, base, "Key", &key);
            if (iprop(p, base, "Tag", &tag)) { vbvar v = vb_v_s(vs(tag)); vb_vset(&t->tag, v); }
            int ix = (int)ipropn(p, base, "ImageIndex", 0);
            if (ix) { vbvar v = vb_v_l(ix); vb_vset(&t->image, v); }
            else if ((iprop(p, base, "Image", &img) || iprop(p, base, "ImageKey", &img)) && !img.isEmpty()) {
                vbvar v = vb_v_s(vs(img)); vb_vset(&t->image, v);
            }
            if (!tabs.insert(t, -1, key)) break;
        }
        sync();
        if (wd) { QSignalBlocker b(wd->bar); wd->bar->setCurrentIndex(0); }
    }
    double client(int what) const               /* ClientLeft, ClientTop, ClientWidth, ClientHeight (twips) */
    {
        int h = wd ? wd->bar->sizeHint().height() : 0;
        double x = c->left + 4 * TPP, y = c->top + (h + 3) * TPP, w = c->width - 8 * TPP, hh = c->height - (h + 7) * TPP;
        return what == 0 ? x : what == 1 ? y : what == 2 ? std::max(0.0, w) : std::max(0.0, hh);
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        if (n == "tabs") return list_ret(tabs, argc, argv);
        if (n == "imagelist") { vbvar r = imagelist_prop(c, &il, &ilname, set, val); if (set) sync(); return r; }
        if (n == "selecteditem") {
            if (!set) { int i = wd ? wd->bar->currentIndex() : -1; return i >= 0 && i < (int)tabs.items.size() ? retobj(tabs.items[i]->obj) : retobj(nullptr); }
            int i = tabs.find(val);
            if (i >= 0 && wd) wd->bar->setCurrentIndex(i);
            return vb_v_empty();
        }
        if (n == "clientleft" || n == "clienttop" || n == "clientwidth" || n == "clientheight") {
            if (set) { vb_raise(383); return vb_v_empty(); }
            int w = n == "clientleft" ? 0 : n == "clienttop" ? 1 : n == "clientwidth" ? 2 : 3;
            double t = client(w);
            return vb_v_f((float)(w < 2 ? from_twips(c, t, w == 1, true) : from_twips(c, t, w == 3, false)));
        }
        if (n == "style") { if (!set) return vb_v_l(style); style = vb_v2i(val); return vb_v_empty(); }
        if (n == "multirow" || n == "hottracking" || n == "separators" || n == "showtips") return set ? vb_v_empty() : vbool(n == "showtips");
        *found = 0;
        return vb_v_empty();
    }
};
static OcxTabStrip *tab_owner(TabItem *t) { return t->list ? dynamic_cast<OcxTabStrip *>(t->list->owner) : nullptr; }
vbvar TabItem::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    OcxTabStrip *ts = tab_owner(this);
    if (n == "caption") { if (!set) return vstr(caption); caption = qs(vb_v2s(val)); return vb_v_empty(); }
    if (n == "tooltiptext") { if (!set) return vstr(tip); tip = qs(vb_v2s(val)); return vb_v_empty(); }
    if (n == "image") { if (!set) return vb_vtmp(image); vb_vset(&image, val); return vb_v_empty(); }
    if (n == "selected") {
        int i = list ? list->index_of(this) : -1;
        if (!set) return vbool(ts && ts->wd && ts->wd->bar->currentIndex() == i);
        if (vb_v2bool(val) && ts && ts->wd) ts->wd->bar->setCurrentIndex(i);
        return vb_v_empty();
    }
    if (n == "left" || n == "top" || n == "width" || n == "height") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        int i = list ? list->index_of(this) : -1;
        QRect r = ts && ts->wd && i >= 0 ? ts->wd->bar->tabRect(i) : QRect();
        double t = (n == "left" ? r.x() : n == "top" ? r.y() : n == "width" ? r.width() : r.height()) * TPP;
        bool pos = n == "left" || n == "top";
        if (pos && ts) t += n == "left" ? ts->c->left : ts->c->top;
        return vb_v_f((float)(ts ? from_twips(ts->c, t, n == "top" || n == "height", pos) : t));
    }
    *found = 0;
    return vb_v_empty();
}
/* Tabs.Add [Index], [Key], [Caption], [Image] */
static OcxItem *tab_add(OcxTabStrip *ts, vb_i32 argc, vbvar *argv)
{
    int pos; QString key;
    if (!add_place(&ts->tabs, argc, argv, &pos, &key)) return nullptr;
    auto *t = new TabItem;
    t->caption = arg_s(argc, argv, 2);
    if (given(argc, argv, 3)) vb_vset(&t->image, argv[3]);
    if (!ts->tabs.insert(t, pos, key)) return nullptr;
    return t;
}

#include "ocx_tree.inc"
#include "ocx_more.inc"
#include "ocx_rich.inc"
#include "ocx_chart.inc"
#include "ocx_anim.inc"
#include "ocx_mci.inc"
#include "ocx_comm.inc"
#include "ocx_winsock.inc"
#include "ocx_inet.inc"
#include "ocx_web.inc"
#include "ocx_mapi.inc"
#include "ocx_mask.inc"
#include "ocx_cool.inc"

/* ---------------------------------------------------------------- PictureClip */
struct OcxPictureClip : Ocx {
    QImage img;
    int rows = 1, cols = 1, cx = 0, cy = 0, cw = 0, ch = 0, sx = 0, sy = 0;
    QWidget *make(QWidget *) override { return nullptr; }
    void load(const char *const *p) override
    {
        img = frx_image(p, "Picture");
        rows = std::max(1, (int)propn(p, "Rows", 1)); cols = std::max(1, (int)propn(p, "Cols", 1));
        cx = (int)propn(p, "ClipX", 0); cy = (int)propn(p, "ClipY", 0);
        cw = (int)propn(p, "ClipWidth", 0); ch = (int)propn(p, "ClipHeight", 0);
        sx = (int)propn(p, "StretchX", 0); sy = (int)propn(p, "StretchY", 0);
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        int v = set ? vb_v2l(val) : 0;
        if (vb_errpend) return vb_v_empty();
        if (n == "picture") {
            if (!set) return retobj(new_pic(img, 1));
            const QImage *im = image_of_value(val);
            img = im ? *im : QImage();
            return vb_v_empty();
        }
        if (n == "rows" || n == "cols") {
            int *p = n == "rows" ? &rows : &cols;
            if (!set) return vb_v_i((vb_i16)*p);
            if (v < 1) { vb_raise(380); return vb_v_empty(); }
            *p = v;
            return vb_v_empty();
        }
        struct { const char *name; int *p; } ints[] = { { "clipx", &cx }, { "clipy", &cy }, { "clipwidth", &cw }, { "clipheight", &ch },
                                                         { "stretchx", &sx }, { "stretchy", &sy } };
        for (auto &e : ints) if (n == e.name) { if (!set) return vb_v_l(*e.p); *e.p = v; return vb_v_empty(); }
        if (n == "cellwidth") return vb_v_l(img.isNull() ? 0 : img.width() / cols);
        if (n == "cellheight") return vb_v_l(img.isNull() ? 0 : img.height() / rows);
        if (n == "width") return vb_v_l(img.width());
        if (n == "height") return vb_v_l(img.height());
        if (n == "graphiccell") {
            if (img.isNull()) { vb_raise(481); return vb_v_empty(); }
            int i = (int)arg_l(argc, argv, 0, 0);
            if (i < 0 || i >= rows * cols) { vb_raise(380); return vb_v_empty(); }
            int w = img.width() / cols, h = img.height() / rows;
            return retobj(new_pic(img.copy((i % cols) * w, (i / cols) * h, w, h), 1));
        }
        if (n == "clip") {
            if (img.isNull()) { vb_raise(481); return vb_v_empty(); }
            QImage part = img.copy(cx, cy, cw > 0 ? cw : img.width() - cx, ch > 0 ? ch : img.height() - cy);
            if (sx > 0 && sy > 0) part = part.scaled(sx, sy);
            return retobj(new_pic(part, 1));
        }
        *found = 0;
        return vb_v_empty();
    }
};

/* the controls of databases (data.inc, ocx_data.inc) */
static Ocx *new_adodc();
static Ocx *new_datagrid();
static Ocx *new_datalist();
static Ocx *new_datacombo();
static Ocx *new_mmcontrol();
static Ocx *new_mscomm();
static Ocx *new_sysinfo();
static Ocx *new_dbgrid();                       /* ocx_dbgrid.inc */
static Ocx *new_dblist();
static Ocx *new_dbcombo();
static Ocx *new_coolbar();                      /* ocx_cool.inc */
static Ocx *new_rdc();                          /* ocx_rdc.inc */
static Ocx *new_drep();                         /* ocx_drep.inc */

/* ---------------------------------------------------------------- the types */
struct OcxType { const char *name, *defmember; Ocx *(*make)(); };
static const OcxType *ocx_type(const std::string &t)
{
    static const OcxType types[] = {
        { "ImageList", nullptr, []() -> Ocx * { return new OcxImageList; } },
        { "ProgressBar", "value", []() -> Ocx * { return new OcxProgress; } },
        { "Slider", "value", []() -> Ocx * { return new OcxSlider; } },
        { "StatusBar", "simpletext", []() -> Ocx * { return new OcxStatusBar; } },
        { "Toolbar", nullptr, []() -> Ocx * { return new OcxToolbar; } },
        { "TabStrip", nullptr, []() -> Ocx * { return new OcxTabStrip; } },
        { "PictureClip", "picture", []() -> Ocx * { return new OcxPictureClip; } },
        { "TreeView", nullptr, []() -> Ocx * { return new OcxTreeView; } },
        { "ListView", nullptr, []() -> Ocx * { return new OcxListView; } },
        { "ImageCombo", "text", []() -> Ocx * { return new OcxImageCombo; } },
        { "UpDown", "value", []() -> Ocx * { return new OcxUpDown; } },
        { "DTPicker", "value", []() -> Ocx * { return new OcxDTPicker; } },
        { "MonthView", "value", []() -> Ocx * { return new OcxMonthView; } },
        { "FlatScrollBar", "value", []() -> Ocx * { return new OcxFlatScroll; } },
        { "SSTab", "tab", []() -> Ocx * { return new OcxSSTab; } },
        { "MSFlexGrid", "text", []() -> Ocx * { return new OcxFlexGrid; } },
        { "MSHFlexGrid", "text", []() -> Ocx * { return new OcxFlexGrid; } },
        { "RichTextBox", "text", []() -> Ocx * { return new OcxRichText; } },
        { "MSChart", nullptr, []() -> Ocx * { return new OcxChart; } },
        { "Animation", nullptr, []() -> Ocx * { return new OcxAnimation; } },
        { "Adodc", nullptr, new_adodc },
        { "DataGrid", "text", new_datagrid },
        { "DataList", "text", new_datalist },
        { "DataCombo", "text", new_datacombo },
        { "DBGrid", "text", new_dbgrid },
        { "DBList", "text", new_dblist },
        { "DBCombo", "text", new_dbcombo },
        { "MMControl", nullptr, new_mmcontrol },
        { "MSComm", nullptr, new_mscomm },
        { "Winsock", "remotehost", new_winsock },
        { "Inet", "url", new_inet },
        { "WebBrowser", nullptr, new_webbrowser },
        { "MAPISession", nullptr, new_mapisession },
        { "MAPIMessages", nullptr, new_mapimessages },
        { "MaskEdBox", "text", new_maskedbox },
        { "SysInfo", nullptr, new_sysinfo },
        { "CoolBar", nullptr, new_coolbar },
        { "MSRDC", nullptr, new_rdc },
        { "DataRepeater", nullptr, new_drep },
    };
    for (auto &e : types) if (t == e.name) return &e;
    return nullptr;
}
/* the control is made: its Ocx, the collections that Add works on */
static Ocx *ocx_new(Ctl *c)
{
    const OcxType *t = ocx_type(c->type);
    if (!t) return nullptr;
    Ocx *x = t->make();
    x->c = c;
    if (auto *il = dynamic_cast<OcxImageList *>(x)) {
        il->images.owner = x;
        il->images.locked = [il]() { return il->bound; };
        il->images.clsname = "ListImages"; il->images.itemcls = "ListImage"; il->images.itemdef = "key";
        il->images.add = [il](vb_i32 argc, vbvar *argv) { return listimage_add(il, argc, argv); };
    } else if (auto *sb = dynamic_cast<OcxStatusBar *>(x)) {
        sb->panels.owner = x;
        sb->panels.clsname = "Panels"; sb->panels.itemcls = "Panel"; sb->panels.itemdef = "text";
        sb->panels.add = [sb](vb_i32 argc, vbvar *argv) { return panel_add(sb, argc, argv); };
        sb->panels.changed = [sb]() { sb->layout(); };
    } else if (auto *tb = dynamic_cast<OcxToolbar *>(x)) {
        tb->buttons.owner = x;
        tb->buttons.clsname = "Buttons"; tb->buttons.itemcls = "Button"; tb->buttons.itemdef = "key";
        tb->buttons.add = [tb](vb_i32 argc, vbvar *argv) { return button_add(tb, argc, argv); };
        tb->buttons.changed = [tb]() { tb->sync(); };
        tb->buttons.removing = [tb](OcxItem *it) {
            auto *b = (TbButton *)it;
            if (b->btn) b->btn->deleteLater();
            if (b->space) b->space->deleteLater();
            b->btn = nullptr; b->space = nullptr;
        };
    } else if (auto *ts = dynamic_cast<OcxTabStrip *>(x)) {
        ts->tabs.owner = x;
        ts->tabs.clsname = "Tabs"; ts->tabs.itemcls = "Tab"; ts->tabs.itemdef = "caption";
        ts->tabs.add = [ts](vb_i32 argc, vbvar *argv) { return tab_add(ts, argc, argv); };
        ts->tabs.changed = [ts]() { ts->sync(); };
    }
    else if (auto *t = dynamic_cast<OcxTreeView *>(x)) {
        t->nodes.owner = x;
        t->nodes.clsname = "Nodes"; t->nodes.itemcls = "Node"; t->nodes.itemdef = "text";
        t->nodes.add = [t](vb_i32 argc, vbvar *argv) { return t->add(argc, argv); };
        t->nodes.remove_all_of = [t](OcxItem *it) { if (((Node *)it)->qi) t->remove_tree((Node *)it); return true; };
        t->nodes.on_clear = [t]() {
            for (OcxItem *i : t->nodes.items) ((Node *)i)->qi = nullptr;
            if (t->tw) t->tw->clear();
        };
    } else if (auto *l = dynamic_cast<OcxListView *>(x)) {
        l->items.owner = x; l->cols.owner = x;
        l->items.clsname = "ListItems"; l->items.itemcls = "ListItem"; l->items.itemdef = "text";
        l->cols.clsname = "ColumnHeaders"; l->cols.itemcls = "ColumnHeader"; l->cols.itemdef = "text";
        l->items.add = [l](vb_i32 argc, vbvar *argv) { return listitem_add(l, argc, argv); };
        l->cols.add = [l](vb_i32 argc, vbvar *argv) { return column_add(l, argc, argv); };
        l->items.removing = [l](OcxItem *it) {
            auto *li = (ListItem *)it;
            if (li->ti) { if (l->report) { int k = l->report->indexOfTopLevelItem(li->ti); if (k >= 0) l->report->takeTopLevelItem(k); } delete li->ti; li->ti = nullptr; }
            if (li->li) { if (l->icons) { int k = l->icons->row(li->li); if (k >= 0) l->icons->takeItem(k); } delete li->li; li->li = nullptr; }
        };
        l->items.on_clear = [l]() {
            for (OcxItem *i : l->items.items) { ((ListItem *)i)->ti = nullptr; ((ListItem *)i)->li = nullptr; }
            if (l->report) l->report->clear();
            if (l->icons) l->icons->clear();
        };
        l->cols.changed = [l]() { l->rebuild(); };
    } else if (auto *ic = dynamic_cast<OcxImageCombo *>(x)) {
        ic->items.owner = x;
        ic->items.clsname = "ComboItems"; ic->items.itemcls = "ComboItem"; ic->items.itemdef = "text";
        ic->items.add = [ic](vb_i32 argc, vbvar *argv) { return comboitem_add(ic, argc, argv); };
        ic->items.changed = [ic]() { ic->sync(); };
    }
    return x;
}
static const char *ocx_default(const std::string &type)
{
    const OcxType *t = ocx_type(type);
    return t ? t->defmember : nullptr;
}

static QWidget *ocx_make(Ctl *c, QWidget *parent) { return c->ocx ? c->ocx->make(parent) : nullptr; }
static void ocx_load(Ctl *c, const char *const *p) { if (c->ocx) c->ocx->load(p); }
static vbvar ocx_member(Ctl *c, const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    *found = 1;
    return c->ocx ? c->ocx->member(n, kind, set, val, argc, argv, found) : vb_v_empty();
}
static void ocx_delete(Ocx *x) { delete x; }
static bool ocx_data_aware(Ocx *x) { return x && x->data_aware(); }
static bool ocx_focusable(Ocx *x) { return x && x->focusable(); }
static bool ocx_clicks(Ocx *x) { return x && x->clicks(); }
static bool ocx_inside(Ocx *x, QWidget *w, QPoint lp) { return x ? x->inside(w, lp) : w->rect().contains(lp); }
static bool ocx_eat_mouse(Ocx *x, QMouseEvent *me) { return x && x->eat_mouse(me); }
static void ocx_resized(Ctl *c) { if (c && c->ocx) c->ocx->resized(); }
static void ocx_loaded(FormState *st)
{
    std::vector<vbobj> v = st->ctls;
    for (vbobj o : v) { Ctl *c = C(o); if (c && c->ocx && c->form == st) c->ocx->loaded(); }
}
static void ocx_shown(FormState *st)
{
    std::vector<vbobj> v = st->ctls;
    for (vbobj o : v) { Ctl *c = C(o); if (c && c->ocx && c->form == st) c->ocx->shown(); }
}
