/* ocx_cool.inc - the CoolBar of ComCtl3 (COMCT332.OCX), phase 7.4. Included by ocx.inc.
 * A CoolBar is a rebar: bands in rows. A band has a gripper, an image of the ImageList, a caption, and
 * one control of the CoolBar as its child. The sizes are those of the VB6 oracles f_cool and f_cool2 (pixels):
 *  - the CoolBar has a border of 2 pixels: its bands are in Width - 4, and Height = the rows + 2 between
 *    two rows (BandBorders) + 4; with no band it keeps its height;
 *  - a row is as high as the highest MinHeight of its bands; VariantHeight = False: as the highest row;
 *  - the header of a band: a gripper of 9 (none for a band of fixed size, for the first band when
 *    FixedOrder, or when one band only can be sized), an image (4 before it when there is no gripper, 2
 *    after it before a caption, else 4), the caption (4 before it when there is no gripper and no image,
 *    6 after it, 5 when the band has no child), then the child, and 4 after the child (none when vertical);
 *    a band is at least as wide as its header + MinWidth + 4, or its header when it has no MinWidth;
 *  - the bands of a row are 2 apart (BandBorders); the last band of a row with a Width of its own takes
 *    the rest of the row (else the last band that can be sized);
 *  - a row is at least as high as the images of its bands + 4; a hidden band hides its child (Visible);
 *  - vertical (Orientation): the bands go down, the rows are columns; the CoolBar keeps its height and
 *    its width follows the columns (HeightChanged gives the width) (f_cool4);
 *  - the child is MinHeight high (a drop-down ComboBox keeps its height) and in the middle of its row;
 *  - the children come to their bands when the form is shown: then Resize (two when aligned), the layout
 *    (HeightChanged and Resize when the height changes).
 *  - a band of fixed size (Style = cc3BandFixedSize) has no gripper and is as wide as its Width; the last
 *    band of the row that can be sized takes the rest of the row (f_cool2).
 * VB6 bugs, not followed: the rebar keeps the gripper a band had before another band was hidden, and AllowVertical
 * = False moved the band to the last Position (f_cool4); Style = cc3BandFixedSize made NewRow False (f_cool2: the .frm the IDE saved), and in
 * f_cool moved the band into the first row, before the band in front of it, with its child 0 wide
 * (f_cool-raw.txt); here the band keeps its row and its place.
 * The user moves a band with the mouse by its gripper: along its row (the band before it grows or shrinks)
 * or to another row (with FixedOrder only a new row after it, or back to the row above).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

/* the width of a caption: MS Sans Serif 8, the font of the rebar (f_cool: "One" 20, "Four" 22, "Zero" 23) */
static int cool_textw(const QString &s, const VFont &f)
{
    static const signed char wd[95] = {
        3, 3, 4, 6, 6, 9, 8, 2, 4, 4, 6, 6, 3, 4, 3, 3,                     /* space ! " # $ % & ' ( ) * + , - . / */
        6, 6, 6, 6, 6, 6, 6, 6, 6, 6,                                         /* 0 - 9 */
        3, 3, 6, 6, 6, 6, 11,                                                 /* : ; < = > ? @ */
        7, 7, 7, 7, 6, 6, 8, 7, 3, 5, 7, 6, 9, 7, 8, 7, 8, 7, 7, 7, 7, 7, 11, 7, 7, 7,   /* A - Z */
        3, 3, 3, 6, 6, 4,                                                     /* [ \ ] ^ _ ` */
        6, 6, 5, 6, 6, 3, 6, 5, 2, 2, 5, 2, 8, 6, 6, 6, 6, 4, 5, 3, 6, 6, 8, 5, 6, 5,     /* a - z (f_cool5: "Three" 28) */
        4, 3, 4, 7 };                                                         /* { | } ~ */
    int w = 0;
    QFontMetrics fm(qfont(f));
    for (QChar ch : s) {
        ushort u = ch.unicode();
        w += u >= 32 && u < 127 ? wd[u - 32] : fm.horizontalAdvance(ch);
    }
    return w;
}
/* a picture in two colours: the light part in the highlight, the dark part in the shadow (EmbossPicture) */
static QImage cool_emboss(const QImage &src, long hi, long sh)
{
    if (src.isNull()) return src;
    QImage im = src.convertToFormat(QImage::Format_ARGB32);
    QRgb h = qcolor(hi).rgb(), s = qcolor(sh).rgb();
    for (int y = 0; y < im.height(); y++) {
        QRgb *row = (QRgb *)im.scanLine(y);
        for (int x = 0; x < im.width(); x++) row[x] = qGray(row[x]) >= 128 ? h : s;
    }
    return im;
}

struct OcxCoolBar;
struct Band : OcxItem {
    OcxCoolBar *cb = nullptr;
    QString caption, childname;
    vbvar image = vb_v_l(0);                    /* an index or a key of the ImageList; 0: none */
    vbobj pic = nullptr;                        /* Picture (counted) */
    vbobj child = nullptr;                      /* a control of the CoolBar (counted) */
    bool newrow = false, visible = true, allowvert = true, fixedbg = true, emboss = false, ucp = true, ucc = true;
    int style = 0;                              /* cc3BandNormal, cc3BandFixedSize */
    long back = 0x8000000F, fore = 0x80000012, ehi = 0x80000014, esh = 0x8000000F;
    int minw = 0, minh = 24, cx = 0;            /* MinWidth, MinHeight, Width: pixels */
    /* the layout: where the band is (pixels, in the client area of the CoolBar) */
    int x = 0, y = 0, w = 0, h = 0, header = 0, row = -1;
    bool grip = false, laid = false;
    int imgh = 0;                               /* the height of its image (the row is at least that + 4) */
    ~Band() { vb_vfree(&image); if (pic) vb_orelease(pic); if (child) vb_orelease(child); }
    bool hid = false;                           /* its child was hidden with the band (Visible False) */
    int post() const;                           /* after the child: 4 (horizontal), 0 (vertical) */
    /* the narrowest it can be: the header, then MinWidth and 4 after the child; a band with no MinWidth can be
     * as narrow as its header (f_cool4: a band of Width 0 at the end of a row is its gripper alone) */
    int minband() const { return header + (minw > 0 ? minw + post() : 0); }
    int need() const { return style == 1 ? cx : minband(); }   /* a band of fixed size is as wide as its Width */
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};

class CoolWidget : public QWidget {
public:
    vbobj self;
    Band *db = nullptr;                         /* the band the user moves */
    QPoint start;
    int grab = 0;                               /* where in the band the user took it (pixels from its left) */
    CoolWidget(vbobj o, QWidget *parent) : QWidget(parent), self(o) { setMouseTracking(true); }
    OcxCoolBar *cb();
protected:
    void paintEvent(QPaintEvent *) override;
    void mousePressEvent(QMouseEvent *e) override;
    void mouseMoveEvent(QMouseEvent *e) override;
    void mouseReleaseEvent(QMouseEvent *e) override;
};

struct OcxCoolBar : Ocx {
    OcxList bands;
    std::vector<Band *> order;                  /* the bands as they are shown (Position) */
    QPointer<CoolWidget> wd;
    bool fixedorder = false, variant = true, borders = true, emboss = false;
    int orient = 0;                             /* cc3OrientationHorizontal, cc3OrientationVertical */
    long back = 0x8000000F, fore = 0x80000012, ehi = 0x80000014, esh = 0x8000000F;
    vbobj pic = nullptr;
    vbobj il = nullptr; QString ilname;
    bool up = false;                            /* the form was shown: the bands have their children */
    bool busy = false, again = false;
    int rows = 0, lastw = -1;
    OcxCoolBar()
    {
        bands.owner = this;
        bands.clsname = "Bands"; bands.itemcls = "Band"; bands.itemdef = nullptr;
        bands.add = [this](vb_i32 argc, vbvar *argv) { return add(argc, argv); };
        bands.changed = [this]() { relayout(); };
        bands.removing = [this](OcxItem *it) {
            auto i = std::find(order.begin(), order.end(), (Band *)it);
            if (i != order.end()) order.erase(i);
            if (wd && wd->db == it) wd->db = nullptr;
        };
        /* Clear takes the bands away from the first: the CoolBar follows each (f_cool: HeightChanged 810, 420) */
        bands.on_clear = [this]() { while (!bands.items.empty()) bands.remove(0); };
    }
    ~OcxCoolBar() override
    {
        bands.changed = nullptr; bands.on_clear = nullptr;
        if (pic) vb_orelease(pic);
        if (il) vb_orelease(il);
    }
    bool vert() const { return orient == 1; }
    int px(double twips) const { return twips <= 0 ? 0 : (int)(twips / TPP); }
    QWidget *make(QWidget *parent) override { wd = new CoolWidget(c->self, parent); return wd; }
    bool clicks() override { return true; }
    /* Click: the button let go over the bands, not over the border of 2 pixels around them (f_cool5) */
    bool inside(QWidget *w, QPoint lp) override { return QRect(0, 0, w->width() - 4, w->height() - 4).contains(lp); }
    bool eat_mouse(QMouseEvent *me) override;
    bool eat_release = false;
    Band *grip_at(QPoint p) const;
    void min_max(Band *b);
    void load(const char *const *p) override;
    void shown() override;
    void resized() override;
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
    OcxItem *add(vb_i32 argc, vbvar *argv);
    Ctl *named(const QString &nm);
    bool is_mine(vbobj o) const;
    int layout(Band *probe, int probeh);
    void place(Band *b, int hh);
    void apply_size();
    bool relayout(Band *probe = nullptr, int probeh = 0);
    std::vector<std::vector<Band *>> row_list() const;
    void set_order(Band *b, int at);
    void leave_row(Band *b, const std::vector<std::vector<Band *>> &rws);
    bool drag_row(Band *b, int a, int d);
    bool drag_along(Band *b, int a, int grab);
    QImage picture_of(vbobj o) const { const QImage *im = image_of_value(vb_v_obj(o)); return im ? *im : QImage(); }
};
int Band::post() const { return cb && cb->vert() ? 0 : 4; }
OcxCoolBar *CoolWidget::cb() { Ctl *x = C(self); return x ? dynamic_cast<OcxCoolBar *>(x->ocx) : nullptr; }

static Ocx *new_coolbar() { return new OcxCoolBar; }

/* "tbMain", "tbar (0)": a control of the form (an element of a control array) */
Ctl *OcxCoolBar::named(const QString &nm)
{
    if (!c->form) return nullptr;
    QString base = nm.trimmed(); int ix = -1;
    int p = base.indexOf('(');
    if (p > 0 && base.endsWith(')')) { ix = base.mid(p + 1, base.size() - p - 2).trimmed().toInt(); base = base.left(p).trimmed(); }
    auto it = c->form->byname.find(lower(base.toUtf8().constData()));
    if (it == c->form->byname.end()) return nullptr;
    vbobj o = it->second;
    if (o->cls == &CLS_ARRAY) { auto *a = (CtlArray *)o->ext; auto e = a->elems.find(ix); o = e == a->elems.end() ? nullptr : e->second; }
    return o ? C(o) : nullptr;
}
/* a control in this CoolBar (only these can be the child of a band: else 380, f_cool2) */
bool OcxCoolBar::is_mine(vbobj o) const
{
    if (!o || !c->form) return false;
    o = vb_oid(o);
    const auto &v = c->form->ctls;
    if (std::find(v.begin(), v.end(), o) == v.end()) return false;
    return C(o)->container == c->self;
}

void OcxCoolBar::load(const char *const *p)
{
    fixedorder = propn(p, "FixedOrder", 0) != 0;
    variant = propn(p, "VariantHeight", -1) != 0;
    borders = propn(p, "BandBorders", -1) != 0;
    emboss = propn(p, "EmbossPicture", 0) != 0;
    orient = (int)propn(p, "Orientation", 0);
    back = propc(p, "BackColor", 0x8000000FL); fore = propc(p, "ForeColor", 0x80000012L);
    ehi = propc(p, "EmbossHighlight", 0x80000014L); esh = propc(p, "EmbossShadow", 0x8000000FL);
    QImage im = frx_image(p, "Picture");
    if (!im.isNull()) { pic = new_pic(im, frx_pictype(p, "Picture")); vb_oaddref(pic); }
    prop(p, "ImageList", &ilname);
    /* the bands: flat properties with the number of the band (VB6 IDE, f_cool2): Caption1, Child1, MinWidth1,
     * MinHeight1, Width1, Key1, NewRow1, AllowVertical1, Visible1, BandTag1, BandStyle1, BandBackColor1 ... */
    int n = (int)propn(p, "BandCount", 3);
    for (int i = 1; i <= n; i++) {
        auto P = [i](const char *name) { return QString("%1%2").arg(name).arg(i).toUtf8(); };
        auto *b = new Band;
        b->cb = this;
        QString v, key;
        prop(p, P("Caption").constData(), &b->caption);
        prop(p, P("Child").constData(), &b->childname);
        prop(p, P("Key").constData(), &key);
        b->minw = px(propn(p, P("MinWidth").constData(), 0));
        b->minh = px(propn(p, P("MinHeight").constData(), 360));
        b->cx = px(propn(p, P("Width").constData(), 0));
        b->newrow = propn(p, P("NewRow").constData(), 0) != 0;
        b->visible = propn(p, P("Visible").constData(), -1) != 0;
        b->allowvert = propn(p, P("AllowVertical").constData(), -1) != 0;
        b->fixedbg = propn(p, P("FixedBackground").constData(), -1) != 0;
        b->ucc = propn(p, P("UseCoolbarColors").constData(), -1) != 0;
        b->ucp = propn(p, P("UseCoolbarPicture").constData(), -1) != 0;
        b->style = (int)propn(p, P("BandStyle").constData(), 0) == 1 ? 1 : 0;
        b->emboss = propn(p, P("BandEmbossPicture").constData(), 0) != 0;
        b->back = propc(p, P("BandBackColor").constData(), 0x8000000FL);
        b->fore = propc(p, P("BandForeColor").constData(), 0x80000012L);
        b->ehi = propc(p, P("BandEmbossHighlight").constData(), 0x80000014L);
        b->esh = propc(p, P("BandEmbossShadow").constData(), 0x8000000FL);
        if (prop(p, P("Image").constData(), &v)) {
            bool ok; int k = v.toInt(&ok);
            vbvar t = ok ? vb_v_l(k) : vb_v_s(vs(v));
            vb_vset(&b->image, t);
        }
        if (prop(p, P("BandTag").constData(), &v)) { vbvar t = vb_v_s(vs(v)); vb_vset(&b->tag, t); }
        QImage bi = frx_image(p, P("BandPicture").constData());
        if (!bi.isNull()) { b->pic = new_pic(bi, frx_pictype(p, P("BandPicture").constData())); vb_oaddref(b->pic); }
        if (!bands.insert(b, -1, key)) { vb_err_clear_pending(); continue; }
        order.push_back(b);
    }
}

/* the form is shown: the bands take their children (VB6: Resize, Resize, then the layout) */
void OcxCoolBar::shown()
{
    if (up || !c->w) return;
    for (Band *b : order) {
        if (b->child || b->childname.isEmpty()) continue;
        Ctl *k = named(b->childname);
        if (k && is_mine(k->self)) { b->child = vb_oid(k->self); vb_oaddref(b->child); }
    }
    lastw = (int)std::lround((vert() ? c->height : c->width) / TPP);
    fire_ctl(c, "resize");                      /* f_cool4; aligned, a second one (f_cool) */
    if (c->align) fire_ctl(c, "resize");
    up = true;
    relayout();
}
/* the CoolBar is wider or narrower (Align, Width): Resize, and its bands follow */
void OcxCoolBar::resized()
{
    if (busy) return;
    int w = (int)std::lround((vert() ? c->height : c->width) / TPP);
    if (w == lastw) return;
    lastw = w;
    if (!up) return;
    fire_ctl(c, "resize");
    relayout();
}

/* the rows of the bands that are shown, from the last layout */
std::vector<std::vector<Band *>> OcxCoolBar::row_list() const
{
    std::vector<std::vector<Band *>> r;
    for (Band *b : order) {
        if (!b->laid) continue;
        if ((int)r.size() <= b->row) r.resize(b->row + 1);
        r[b->row].push_back(b);
    }
    return r;
}

/* where the bands go; the new extent of the CoolBar (pixels, its height) or -1 when it has no band shown.
 * probe: a band that has a new child, as high as that child for a moment (f_cool: HeightChanged 1260, 1335) */
int OcxCoolBar::layout(Band *probe, int probeh)
{
    int W = std::max(0, (int)std::lround((vert() ? c->height : c->width) / TPP) - 4);
    std::vector<Band *> vb;
    for (Band *b : order) {
        b->laid = false; b->row = -1;
        if (b->visible && (!vert() || b->allowvert)) vb.push_back(b);
    }
    int nonfixed = 0;
    for (Band *b : vb) if (b->style == 0) nonfixed++;
    OcxImageList *list = imagelist_of(c, &il, ilname);
    int sep = borders ? 2 : 0;
    for (size_t i = 0; i < vb.size(); i++) {
        Band *b = vb[i];
        /* the gripper: not on a band of fixed size, not on the first band with FixedOrder, not when one band
         * only can be sized (f_cool4; the rebar of VB6 keeps a gripper it drew before another band was hidden:
         * not followed, the gripper is what the bands are now) */
        b->grip = b->style == 0 && nonfixed > 1 && !(fixedorder && i == 0);
        bool img = list && !image_from(list, b->image).isNull();
        /* f_cool4: a gripper 9 (else 4 before an image or a caption); an image, then 2 before a caption or 4;
         * a caption, then 6 before the child (5 when the band has no child) */
        int hd = b->grip ? 9 : 0;
        if ((img || !b->caption.isEmpty()) && !b->grip) hd += 4;
        if (img) hd += list->iw + (b->caption.isEmpty() ? 4 : 2);
        if (!b->caption.isEmpty()) hd += cool_textw(b->caption, c->font) + (b->child ? 6 : 5);
        b->header = hd;
        b->imgh = img ? list->ih : 0;
    }
    /* the rows: NewRow, or no room left in the row */
    std::vector<std::vector<Band *>> rws;
    int used = 0;
    for (Band *b : vb) {
        if (rws.empty() || b->newrow || (used + b->need() > W && !rws.back().empty())) { rws.emplace_back(); used = 0; }
        rws.back().push_back(b);
        used += b->need() + sep;
    }
    for (auto &r : rws) {
        int n = (int)r.size(), tot = sep * (n - 1);
        for (Band *b : r) { b->w = b->style == 1 ? b->cx : std::max(b->cx, b->minband()); tot += b->w; }
        if (tot < W) {                          /* the rest of the row: to the last band that has a Width of its own
                                                 * (f_cool4, f_cool2), else to the first band that can be sized */
            Band *g = nullptr;
            for (int k = n - 1; k >= 0 && !g; k--) if (r[k]->style == 0 && r[k]->cx > 0) g = r[k];
            for (int k = 0; k < n && !g; k++) if (r[k]->style == 0) g = r[k];      /* none: the first (f_cool7) */
            if (!g) g = r.back();
            g->w += W - tot;
        } else for (int k = n - 1, ex = tot - W; k >= 0 && ex > 0; k--) {
            if (r[k]->style == 1) continue;
            int d = std::min(ex, r[k]->w - r[k]->minband());
            if (d > 0) { r[k]->w -= d; ex -= d; }
        }
        int x = 0;
        for (Band *b : r) { b->x = x; x += b->w + sep; }
    }
    std::vector<int> rh(rws.size(), 0);
    int maxh = 0;
    for (size_t i = 0; i < rws.size(); i++) {
        for (Band *b : rws[i]) rh[i] = std::max({ rh[i], b == probe ? probeh : b->minh, b->imgh ? b->imgh + 4 : 0 });
        maxh = std::max(maxh, rh[i]);
    }
    if (!variant) for (int &h : rh) h = maxh;
    int y = 0;
    for (size_t i = 0; i < rws.size(); i++) {
        for (Band *b : rws[i]) { b->y = y; b->h = rh[i]; b->row = (int)i; b->laid = true; }
        y += rh[i] + (i + 1 < rws.size() ? sep : 0);
    }
    rows = (int)rws.size();
    for (Band *b : order) {
        if (!b->child) continue;
        Ctl *k = C(b->child);
        if (!k || !k->form || !k->w) continue;
        if (!b->laid) {                         /* a hidden band hides its child (f_cool4: Visible False) */
            if (k->visible) { k->visible = false; b->hid = true; }
            k->w->setVisible(false);
            continue;
        }
        if (b->hid) { k->visible = true; b->hid = false; }
        k->w->setVisible(k->visible);
        place(b, b == probe ? probeh : b->minh);
    }
    return rws.empty() ? -1 : y + 4;
}
/* the child of a band: after the header, MinHeight high, in the middle of the row */
void OcxCoolBar::place(Band *b, int hh)
{
    Ctl *k = C(b->child);
    int along = b->x + b->header, len = std::max(0, b->w - b->header - b->post());
    bool fixedh = (k->kind == K_COMBO && (!k->def || propn(k->def->props, "Style", 0) != 1)) || k->kind == K_DRIVE;
    if (!vert()) {
        k->left = along * TPP; k->width = len * TPP;
        if (!fixedh) k->height = hh * TPP;
        ctl_geometry(k);
        int ch = (int)std::lround(k->height / TPP);
        k->top = (b->y + (b->h - ch) / 2) * TPP;
    } else {
        k->top = along * TPP; k->height = len * TPP;
        k->width = hh * TPP;
        ctl_geometry(k);
        int cw = (int)std::lround(k->width / TPP);
        k->left = (b->y + (b->h - cw) / 2) * TPP;
    }
    ctl_geometry(k);
}
/* the new height of the CoolBar: on the form (Align: the controls aligned after it move) */
void OcxCoolBar::apply_size()
{
    if (c->align && c->form && aligning) {
        if (c->w) c->w->setGeometry((int)(c->left / TPP), (int)(c->top / TPP), (int)(c->width / TPP), (int)(c->height / TPP));
        FormState *st = c->form;
        QTimer::singleShot(0, [st]() { if (std::find(loaded_forms.begin(), loaded_forms.end(), st->form) != loaded_forms.end()) align_layout(st); });
        return;
    }
    ctl_geometry(c);
}
bool OcxCoolBar::relayout(Band *probe, int probeh)
{
    if (!up || !c->w) return false;
    if (busy) { again = true; return false; }
    busy = true;
    bool fired = false;
    for (int pass = 0; pass < 8; pass++) {
        again = false;
        int ext = layout(probe, probeh);
        double &cur = vert() ? c->width : c->height;
        if (ext >= 0 && (int)std::lround(cur / TPP) != ext) {
            cur = ext * TPP;
            apply_size();
            vbvar a = vb_v_f((float)from_twips(c, cur, !vert(), false));
            fire_ctl(c, "heightchanged", 1, &a);
            fire_ctl(c, "resize");
            fired = true;
            if (!c->w) break;                   /* the form went in an event */
        }
        lastw = (int)std::lround((vert() ? c->height : c->width) / TPP);
        if (probe) { probe = nullptr; again = true; }
        if (!again) break;
    }
    busy = false;
    if (wd) wd->update();
    return fired;
}

/* Bands.Add [Index], [Key], [Caption], [Image], [NewRow] */
OcxItem *OcxCoolBar::add(vb_i32 argc, vbvar *argv)
{
    int pos; QString key;
    if (!add_place(&bands, argc, argv, &pos, &key)) return nullptr;
    if (!key.isEmpty() && bands.find_key(key) >= 0) { vb_raise(457); return nullptr; }   /* f_cool: 457, not 35602 */
    auto *b = new Band;
    b->cb = this;
    if (given(argc, argv, 2)) b->caption = arg_s(argc, argv, 2);
    if (given(argc, argv, 3)) vb_vset(&b->image, argv[3]);
    if (given(argc, argv, 4)) b->newrow = vb_v2bool(argv[4]) != 0;
    if (vb_errpend) { delete b; return nullptr; }
    if (!bands.insert(b, pos, key)) return nullptr;
    order.insert(order.begin() + (pos < 0 || pos > (int)order.size() ? (int)order.size() : pos), b);
    return b;
}

/* ---------------------------------------------------------------- the user moves a band with the mouse
 * (VB6 oracles f_cool5, f_cool7: the rebar of Windows). The left button down on a gripper takes the band; each move of
 * the mouse then does one of two things:
 *  - the cursor is over the band's row: the band's left edge goes to the cursor (where it was taken in the band kept):
 *    to the left, the bands before it get narrower (the nearest first, down to their minimum) and it wider by as
 *    much, or, when they cannot give that much, it changes places with the band before it (not with FixedOrder); to
 *    the right, it and the bands after it get narrower and the band before it wider. The first band of a row does
 *    not move this way. The Width of a band that changed width is its width.
 *  - the cursor is above or below the row: the band goes into the row there (a gap between two rows: the next row
 *    that way), after the bands of that row whose left edge is left of the cursor; below the CoolBar: to the end in a
 *    row of its own (when its row has another band); above it: to the first place. The widths do not change.
 *    With FixedOrder the order stays: down, the band begins a row; up, the first band of a row joins the row above.
 * NewRow follows: a band that leaves the beginning of a row gives NewRow to the band after it; one that comes to the
 * beginning of a row takes the NewRow of the band that was there. */
void OcxCoolBar::set_order(Band *b, int at)
{
    auto i = std::find(order.begin(), order.end(), b);
    if (i == order.end()) return;
    order.erase(i);
    at = std::max(0, std::min(at, (int)order.size()));
    order.insert(order.begin() + at, b);
}
/* the band leaves its row: the band after it in that row begins the row when it did */
void OcxCoolBar::leave_row(Band *b, const std::vector<std::vector<Band *>> &rws)
{
    auto &r = rws[b->row];
    size_t k = std::find(r.begin(), r.end(), b) - r.begin();
    if (k == 0 && r.size() > 1 && b->newrow) { r[1]->newrow = true; }
}
bool OcxCoolBar::drag_row(Band *b, int a, int d)
{
    auto rws = row_list();
    if (b->row < 0 || b->row >= (int)rws.size()) return false;
    int H = (int)std::lround((vert() ? c->width : c->height) / TPP);
    auto &mine = rws[b->row];
    size_t k = std::find(mine.begin(), mine.end(), b) - mine.begin();
    if (fixedorder) {
        bool down = d > b->y + b->h;
        if (down && k > 0) b->newrow = true;
        else if (!down && k == 0 && b->row > 0) b->newrow = false;
        else return false;
        relayout();
        return true;
    }
    if (d < 0) {                                /* above the CoolBar: the first place */
        if (std::find(order.begin(), order.end(), b) == order.begin()) return false;
        leave_row(b, rws);
        set_order(b, 0);
        b->newrow = true;
        relayout();
        return true;
    }
    if (d >= H) {                               /* below it: a row of its own at the end */
        if (mine.size() < 2) return false;
        leave_row(b, rws);
        set_order(b, (int)order.size());
        b->newrow = true;
        relayout();
        return true;
    }
    int row = -1;
    for (size_t i = 0; i < rws.size(); i++)
        if (d >= rws[i].front()->y && d <= rws[i].front()->y + rws[i].front()->h) { row = (int)i; break; }
    if (row < 0) {                              /* the gap between its row and the next row that way */
        row = d < b->y ? b->row - 1 : b->row + 1;
        if (row < 0 || row >= (int)rws.size()) return false;
        Band *f = rws[row].front();
        bool gap = d < b->y ? (d > f->y + f->h && d < b->y) : (d > b->y + b->h && d < f->y);
        if (!gap) return false;                 /* beyond it (the border of the CoolBar): nothing (f_cool7 s1) */
    }
    if (row == b->row) return false;
    auto &to = rws[row];
    int at = 0;
    for (Band *x : to) if (x->x < a) at++;
    leave_row(b, rws);
    Band *before = at > 0 ? to[at - 1] : nullptr;
    if (at == 0) { b->newrow = to.front()->newrow; to.front()->newrow = false; }
    else b->newrow = false;
    int pos;
    if (before) pos = (int)(std::find(order.begin(), order.end(), before) - order.begin()) + 1;
    else pos = (int)(std::find(order.begin(), order.end(), to.front()) - order.begin());
    if (pos > (int)(std::find(order.begin(), order.end(), b) - order.begin())) pos--;
    set_order(b, pos);
    relayout();
    return true;
}
bool OcxCoolBar::drag_along(Band *b, int a, int grab)
{
    auto rws = row_list();
    if (b->row < 0 || b->row >= (int)rws.size()) return false;
    auto &r = rws[b->row];
    size_t k = std::find(r.begin(), r.end(), b) - r.begin();
    if (k == 0) return false;
    int movement = a - (b->x + grab);
    if (!movement) return false;
    auto shrink = [](Band *x, int &left) {
        if (x->style == 1 || left <= 0) return;
        int dd = std::min(left, x->w - x->minband());
        if (dd > 0) { x->w -= dd; left -= dd; x->cx = x->w; }
    };
    int need = std::abs(movement), left = need;
    if (movement < 0) {
        int avail = 0;
        for (size_t j = 0; j < k; j++) if (r[j]->style == 0) avail += std::max(0, r[j]->w - r[j]->minband());
        if (avail < need && !fixedorder) {      /* past the band before it: they change places */
            Band *p = r[k - 1];
            bool first = k - 1 == 0;
            if (first) { b->newrow = p->newrow; p->newrow = false; }
            set_order(b, (int)(std::find(order.begin(), order.end(), p) - order.begin()));
            relayout();
            return true;
        }
        for (int j = (int)k - 1; j >= 0 && left > 0; j--) shrink(r[j], left);
        b->w += need - left; b->cx = b->w;
    } else {
        for (size_t j = k; j < r.size() && left > 0; j++) shrink(r[j], left);
        r[k - 1]->w += need - left; r[k - 1]->cx = r[k - 1]->w;
    }
    if (left == need) return false;
    relayout();
    return true;
}

/* the band whose gripper is under a point of the widget */
Band *OcxCoolBar::grip_at(QPoint p) const
{
    int a = vert() ? p.y() : p.x(), d = vert() ? p.x() : p.y();
    for (Band *b : order)
        if (b->laid && b->grip && a >= b->x && a < b->x + std::max(9, b->header) && d >= b->y && d < b->y + b->h) return b;
    return nullptr;
}
/* a double click on a gripper: the band down to its header (the band before it, or after it, takes the room), or,
 * when it is that already, as wide as the row lets it (the others down to theirs); no DblClick (f_cool5) */
void OcxCoolBar::min_max(Band *b)
{
    auto rws = row_list();
    if (b->row < 0 || b->row >= (int)rws.size()) return;
    auto &r = rws[b->row];
    size_t k = std::find(r.begin(), r.end(), b) - r.begin();
    if (r.size() < 2) return;
    if (b->w > b->minband()) {
        Band *o = k > 0 ? r[k - 1] : r[k + 1];
        o->w += b->w - b->minband();
        b->w = b->minband();
    } else {
        for (Band *x : r) if (x != b && x->style == 0 && x->w > x->minband()) { b->w += x->w - x->minband(); x->w = x->minband(); }
    }
    for (Band *x : r) x->cx = x->w;
    relayout();
}
bool OcxCoolBar::eat_mouse(QMouseEvent *me)
{
    if (me->type() == QEvent::MouseButtonRelease && eat_release) { eat_release = false; return true; }
    if (me->type() != QEvent::MouseButtonDblClick || me->button() != Qt::LeftButton || !wd) return false;
    Band *b = grip_at(wd->mapFromGlobal(ev_gpos(me).toPoint()));
    if (!b) return false;
    eat_release = true;
    if (wd) wd->db = nullptr;
    min_max(b);
    return true;
}

void CoolWidget::mousePressEvent(QMouseEvent *e)
{
    e->accept();
    OcxCoolBar *o = cb();
    db = nullptr;
    if (!o || e->button() != Qt::LeftButton) return;
    QPoint p = e->pos();
    db = o->grip_at(p);
    if (db) { start = p; grab = (o->vert() ? p.y() : p.x()) - db->x; }
}
void CoolWidget::mouseMoveEvent(QMouseEvent *e)
{
    e->accept();
    OcxCoolBar *o = cb();
    if (!o) return;
    QPoint p = e->pos();
    int a = o->vert() ? p.y() : p.x(), d = o->vert() ? p.x() : p.y();
    if (!db) {                                  /* over a gripper: the cursor that moves it */
        if (o->grip_at(p)) setCursor(o->vert() ? Qt::SizeVerCursor : Qt::SizeHorCursor); else unsetCursor();
        return;
    }
    if (std::find(o->order.begin(), o->order.end(), db) == o->order.end() || !db->laid) { db = nullptr; return; }
    if (d < db->y || d > db->y + db->h) o->drag_row(db, a, d);
    else o->drag_along(db, a, grab);
}
void CoolWidget::mouseReleaseEvent(QMouseEvent *e)
{
    e->accept();
    db = nullptr;
}
void CoolWidget::paintEvent(QPaintEvent *)
{
    OcxCoolBar *o = cb();
    if (!o) return;
    QPainter p(this);
    p.fillRect(rect(), qcolor(o->back));
    QImage bg = o->pic ? o->picture_of(o->pic) : QImage();
    if (!bg.isNull() && o->emboss) bg = cool_emboss(bg, o->ehi, o->esh);
    if (!bg.isNull()) p.drawTiledPixmap(rect(), QPixmap::fromImage(bg));
    Ctl *c = o->c;
    QFont f = qfont(c->font);
    p.setFont(f);
    QColor light = qcolor(0x80000014L), dark = qcolor(0x80000010L);
    OcxImageList *list = imagelist_of(c, &o->il, o->ilname);
    for (Band *b : o->order) {
        if (!b->laid) continue;
        QRect r = o->vert() ? QRect(b->y, b->x, b->h, b->w) : QRect(b->x, b->y, b->w, b->h);
        if (!b->ucc) p.fillRect(r, qcolor(b->back));
        if (!b->ucp && b->pic) {
            QImage bi = o->picture_of(b->pic);
            if (b->emboss) bi = cool_emboss(bi, b->ehi, b->esh);
            if (!bi.isNull()) p.drawTiledPixmap(r, QPixmap::fromImage(bi), b->fixedbg ? QPoint(r.x(), r.y()) : QPoint());
        }
        int at = 0;
        if (b->grip) {                          /* the gripper: two raised lines */
            for (int i = 0; i < 2; i++) {
                if (!o->vert()) {
                    int gx = r.x() + 2 + i * 3;
                    p.setPen(light); p.drawLine(gx, r.y() + 2, gx, r.bottom() - 2);
                    p.setPen(dark); p.drawLine(gx + 1, r.y() + 2, gx + 1, r.bottom() - 2);
                } else {
                    int gy = r.y() + 2 + i * 3;
                    p.setPen(light); p.drawLine(r.x() + 2, gy, r.right() - 2, gy);
                    p.setPen(dark); p.drawLine(r.x() + 2, gy + 1, r.right() - 2, gy + 1);
                }
            }
            at = 9;
        } else if (!b->caption.isEmpty() || (list && !image_from(list, b->image).isNull())) at = 4;
        QPixmap im = list ? image_from(list, b->image) : QPixmap();
        if (!im.isNull()) {
            if (!o->vert()) p.drawPixmap(r.x() + at, r.y() + (r.height() - im.height()) / 2, im);
            else p.drawPixmap(r.x() + (r.width() - im.width()) / 2, r.y() + at, im);
            at += im.width() + 2;
        }
        if (!b->caption.isEmpty()) {
            p.setPen(qcolor(b->ucc ? o->fore : b->fore));
            if (!o->vert()) p.drawText(QRect(r.x() + at, r.y(), b->header, r.height()), Qt::AlignLeft | Qt::AlignVCenter, b->caption);
            else p.drawText(QRect(r.x(), r.y() + at, r.width(), b->header), Qt::AlignHCenter | Qt::AlignTop, b->caption);
        }
        if (o->borders) {                       /* the etched lines between the bands and under the rows */
            if (!o->vert()) {
                int ex = r.right() + 1;
                if (b->x + b->w + 2 < width() - 4) { p.setPen(dark); p.drawLine(ex, r.y(), ex, r.bottom()); p.setPen(light); p.drawLine(ex + 1, r.y(), ex + 1, r.bottom()); }
                if (b->row + 1 < o->rows) { p.setPen(dark); p.drawLine(r.x(), r.bottom() + 1, r.right() + 2, r.bottom() + 1); p.setPen(light); p.drawLine(r.x(), r.bottom() + 2, r.right() + 2, r.bottom() + 2); }
            } else {
                int ey = r.bottom() + 1;
                if (b->x + b->w + 2 < height() - 4) { p.setPen(dark); p.drawLine(r.x(), ey, r.right(), ey); p.setPen(light); p.drawLine(r.x(), ey + 1, r.right(), ey + 1); }
                if (b->row + 1 < o->rows) { p.setPen(dark); p.drawLine(r.right() + 1, r.y(), r.right() + 1, r.bottom() + 2); p.setPen(light); p.drawLine(r.right() + 2, r.y(), r.right() + 2, r.bottom() + 2); }
            }
        }
    }
}

/* ---------------------------------------------------------------- the members */
vbvar Band::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    OcxCoolBar *o = cb;
    Ctl *c = o ? o->c : nullptr;
    auto flag = [&](bool *f) { if (!set) return vbool(*f); *f = vb_v2bool(val) != 0; return vb_v_empty(); };
    auto color = [&](long *v) { if (!set) return vb_v_l((vb_i32)*v); *v = vb_v2l(val); if (o && o->wd) o->wd->update(); return vb_v_empty(); };
    /* MinWidth, Width: across; MinHeight, Height: down (the ScaleMode of the CoolBar's container); whole pixels */
    auto size = [&](int *v, bool down) {
        if (!set) return vb_v_f((float)(c ? from_twips(c, *v * TPP, down, false) : *v * TPP));
        double t = c ? to_twips(c, vb_v2d(val), down, false) : vb_v2d(val);
        if (vb_errpend) return vb_v_empty();
        *v = t <= 0 ? 0 : (int)(t / TPP);       /* f_cool: 2000 -> 1995, -5 -> 0 */
        return vb_v_empty();
    };
    if (n == "caption") { if (!set) return vstr(caption); caption = qs(vb_v2s(val)); return vb_v_empty(); }
    if (n == "image") {
        if (!set) return vb_vtmp(image);
        /* an index or a key of the ImageList (f_cool4: 35600 for an index out of it or with no ImageList,
         * 35601 for a key it does not have) */
        bool zero = val.vt != VT_BSTR && val.vt != VT_EMPTY && vb_v2l(val) == 0;
        if (vb_errpend) return vb_v_empty();
        if (!zero && val.vt != VT_EMPTY) {
            OcxImageList *lst = o ? imagelist_of(c, &o->il, o->ilname) : nullptr;
            if (!lst) { ocx_error(val.vt == VT_BSTR ? 35601 : 35600); return vb_v_empty(); }
            if (lst->images.find(val) < 0) return vb_v_empty();
        }
        vb_vset(&image, val);
        return vb_v_empty();
    }
    if (n == "picture") {
        if (!set) return retobj(pic);
        if (val.vt == VT_DISPATCH && val.u.o && !image_of_value(val)) { vb_raise(481); return vb_v_empty(); }
        if (pic) vb_orelease(pic);
        pic = val.vt == VT_DISPATCH && val.u.o ? vb_oid(val.u.o) : nullptr;
        if (pic) vb_oaddref(pic);
        if (o && o->wd) o->wd->update();
        return vb_v_empty();
    }
    if (n == "position") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        if (!list || !o) { ocx_error(35605); return vb_v_empty(); }
        auto i = std::find(o->order.begin(), o->order.end(), this);
        return vb_v_l((vb_i32)(i - o->order.begin()) + 1);
    }
    if (n == "newrow") return flag(&newrow);
    if (n == "visible") return flag(&visible);
    if (n == "allowvertical") return flag(&allowvert);
    if (n == "fixedbackground") return flag(&fixedbg);
    if (n == "embosspicture") return flag(&emboss);
    if (n == "usecoolbarpicture") return flag(&ucp);
    if (n == "usecoolbarcolors") return flag(&ucc);
    if (n == "style") {
        if (!set) return vb_v_l(style);
        vb_i32 v = vb_v2l(val);
        if (vb_errpend) return vb_v_empty();
        if (v != 0 && v != 1) { vb_raise(380); return vb_v_empty(); }
        style = v;
        return vb_v_empty();
    }
    if (n == "minwidth") return size(&minw, false);
    if (n == "minheight") return size(&minh, true);
    if (n == "width") return size(&cx, false);
    if (n == "height") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        int hp = laid ? h : 0;
        return vb_v_f((float)(c ? from_twips(c, hp * TPP, true, false) : hp * TPP));
    }
    if (n == "backcolor") return color(&back);
    if (n == "forecolor") return color(&fore);
    if (n == "embosshighlight") return color(&ehi);
    if (n == "embossshadow") return color(&esh);
    if (n == "child") {
        if (!set) return retobj(child);
        vbobj k = val.vt == VT_DISPATCH && val.u.o ? vb_oid(val.u.o) : nullptr;
        if (k && (!o || !o->is_mine(k))) { vb_raise(380); return vb_v_empty(); }   /* f_cool2 */
        if (o) for (Band *x : o->order) if (x != this && x->child == k && k) { vb_orelease(x->child); x->child = nullptr; }
        if (child) vb_orelease(child);
        child = k;
        if (k) vb_oaddref(k);
        /* the band is as high as its new child for a moment, then MinHeight high again (f_cool: 1260, 1335) */
        if (k && o && o->up && laid) {
            int h0 = (int)std::lround(C(k)->height / TPP);
            if (h0 != minh) o->relayout(this, h0);
        }
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

vbvar OcxCoolBar::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    auto flag = [&](bool *f) {
        if (!set) return vbool(*f);
        *f = vb_v2bool(val) != 0;
        relayout();
        if (wd) wd->update();
        return vb_v_empty();
    };
    auto color = [&](long *v) { if (!set) return vb_v_l((vb_i32)*v); *v = vb_v2l(val); if (wd) wd->update(); return vb_v_empty(); };
    if (n == "bands") return list_ret(bands, argc, argv);
    if (n == "rowcount") return vb_v_l(up ? rows : 0);
    if (n == "fixedorder") return flag(&fixedorder);
    if (n == "variantheight") return flag(&variant);
    if (n == "bandborders") return flag(&borders);
    if (n == "embosspicture") return flag(&emboss);
    if (n == "orientation") {
        if (!set) return vb_v_l(orient);
        vb_i32 v = vb_v2l(val);
        if (vb_errpend) return vb_v_empty();
        if (v != 0 && v != 1) { vb_raise(380); return vb_v_empty(); }
        if (v == orient) return vb_v_empty();
        orient = v;
        lastw = (int)std::lround((vert() ? c->height : c->width) / TPP);
        /* the CoolBar keeps its size; its extent across the bands follows them (f_cool4: Width 6000 -> 690 when
         * vertical, then horizontal again 690 wide); a Resize when the size did not change */
        if (!relayout() && up) fire_ctl(c, "resize");
        return vb_v_empty();
    }
    if (n == "backcolor") return color(&back);
    if (n == "forecolor") return color(&fore);
    if (n == "embosshighlight") return color(&ehi);
    if (n == "embossshadow") return color(&esh);
    if (n == "picture") {
        if (!set) return retobj(pic);
        if (val.vt == VT_DISPATCH && val.u.o && !image_of_value(val)) { vb_raise(481); return vb_v_empty(); }
        if (pic) vb_orelease(pic);
        pic = val.vt == VT_DISPATCH && val.u.o ? vb_oid(val.u.o) : nullptr;
        if (pic) vb_oaddref(pic);
        if (wd) wd->update();
        return vb_v_empty();
    }
    if (n == "imagelist") { vbvar r = imagelist_prop(c, &il, &ilname, set, val); if (set) relayout(); return r; }
    if (n == "refresh") { if (wd) wd->update(); return vb_v_empty(); }
    if (n == "aboutbox") return vb_v_empty();
    *found = 0;
    return vb_v_empty();
}
