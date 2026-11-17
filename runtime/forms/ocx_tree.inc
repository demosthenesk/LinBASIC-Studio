/* ocx_tree.inc - TreeView, ListView and ImageCombo of MSComctlLib (included by ocx.inc).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

/* the Image, SelectedImage ... of an item: a number or a key of the image list */
static QIcon icon_from(OcxImageList *il, vbvar img)
{
    QPixmap px = image_from(il, img);
    return px.isNull() ? QIcon() : QIcon(px);
}
static vbvar image_prop(vbvar *slot, OcxImageList *il, bool set, vbvar val, bool *changed)
{
    if (!set) return vb_vtmp(*slot);
    if (!image_ok(il, val)) return vb_v_empty();
    vb_vset(slot, val);
    *changed = true;
    return vb_v_empty();
}
static long item_color(vbvar v) { return vb_v2l(v); }

/* ---------------------------------------------------------------- TreeView */
struct OcxTreeView;
struct Node : OcxItem {
    QTreeWidgetItem *qi = nullptr;
    vbvar image{}, selimage{}, expimage{};
    bool bold = false, sorted = false;
    long fore = -1, back = -1;
    ~Node() { vb_vfree(&image); vb_vfree(&selimage); vb_vfree(&expimage); }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
static Node *node_of(QTreeWidgetItem *qi) { return qi ? (Node *)qi->data(0, Qt::UserRole).value<void *>() : nullptr; }
class TreeW : public QTreeWidget {
public:
    std::function<bool()> before_edit;          /* BeforeLabelEdit: false = cancel */
    bool manual_only = false, allow = false;
    explicit TreeW(QWidget *parent) : QTreeWidget(parent) {}
protected:
    bool edit(const QModelIndex &index, EditTrigger trigger, QEvent *event) override
    {
        if (!allow) {
            if (!(editTriggers() & trigger)) return QTreeWidget::edit(index, trigger, event);
            if (manual_only) return false;
            if (before_edit && !before_edit()) return false;     /* BeforeLabelEdit: Cancel */
        }
        return QTreeWidget::edit(index, trigger, event);
    }
};
struct OcxTreeView : Ocx {
    OcxList nodes;
    vbobj il = nullptr; QString ilname;
    QPointer<TreeW> tw;
    int style = 7, linestyle = 0, labeledit = 0;
    QString pathsep = "\\";
    bool sorted = false, checkboxes = false, editing = false;
    QTreeWidgetItem *lastcur = nullptr;
    ~OcxTreeView() { if (il) vb_orelease(il); }
    OcxImageList *list() { return imagelist_of(c, &il, ilname); }
    QWidget *make(QWidget *parent) override
    {
        tw = new TreeW(parent);
        tw->setHeaderHidden(true);
        tw->setColumnCount(1);
        tw->setEditTriggers(QAbstractItemView::SelectedClicked | QAbstractItemView::EditKeyPressed);
        vbobj self = c->self;
        auto me = [self]() -> OcxTreeView * { Ctl *x = C(self); return x ? dynamic_cast<OcxTreeView *>(x->ocx) : nullptr; };
        tw->before_edit = [self]() {
            vbvar cancel = vb_v_i(0);
            fire_ctl(C(self), "beforelabeledit", 1, &cancel);
            return vb_v2i(cancel) == 0;
        };
        QObject::connect(tw, &QTreeWidget::itemClicked, [self](QTreeWidgetItem *qi, int) {
            Node *n = node_of(qi);
            if (n && n->obj) { vbvar a = vb_v_obj(n->obj); fire_ctl(C(self), "nodeclick", 1, &a); }
        });
        QObject::connect(tw, &QTreeWidget::currentItemChanged, [me](QTreeWidgetItem *cur, QTreeWidgetItem *prev) {
            OcxTreeView *t = me();
            if (!t) return;
            if (prev && node_of(prev)) t->paint(node_of(prev));
            if (cur && node_of(cur)) t->paint(node_of(cur));
        });
        QObject::connect(tw, &QTreeWidget::itemExpanded, [self, me](QTreeWidgetItem *qi) {
            Node *n = node_of(qi);
            if (OcxTreeView *t = me()) if (n) t->paint(n);
            if (n && n->obj) { vbvar a = vb_v_obj(n->obj); fire_ctl(C(self), "expand", 1, &a); }
        });
        QObject::connect(tw, &QTreeWidget::itemCollapsed, [self, me](QTreeWidgetItem *qi) {
            Node *n = node_of(qi);
            if (OcxTreeView *t = me()) if (n) t->paint(n);
            if (n && n->obj) { vbvar a = vb_v_obj(n->obj); fire_ctl(C(self), "collapse", 1, &a); }
        });
        QObject::connect(tw, &QTreeWidget::itemChanged, [self, me](QTreeWidgetItem *qi, int) {
            OcxTreeView *t = me();
            Node *n = node_of(qi);
            if (!t || !n || t->editing) return;
            if (t->checkboxes && qi->data(0, Qt::UserRole + 1).toInt() != (int)qi->checkState(0)) {   /* NodeCheck */
                qi->setData(0, Qt::UserRole + 1, (int)qi->checkState(0));
                vbvar a = vb_v_obj(n->obj); fire_ctl(C(self), "nodecheck", 1, &a);
                return;
            }
            QString old = qi->data(0, Qt::UserRole + 2).toString(), now = qi->text(0);
            if (old == now) return;
            vbvar a[2] = { vb_v_i(0), vb_v_s(vs(now)) };
            fire_ctl(C(self), "afterlabeledit", 2, a);
            t->editing = true;
            if (vb_v2i(a[0])) qi->setText(0, old);                  /* Cancel = True: the old text stays */
            else if (a[1].vt == VT_BSTR) qi->setText(0, qs(a[1].u.s));
            qi->setData(0, Qt::UserRole + 2, qi->text(0));
            t->editing = false;
        });
        return tw;
    }
    /* the look of one node: text, picture (selected / expanded), font, colors */
    void paint(Node *n)
    {
        if (!n->qi || !tw) return;
        editing = true;
        QTreeWidgetItem *qi = n->qi;
        bool pics = style == 1 || style == 3 || style == 5 || style == 7;
        vbvar img = n->image;
        if (tw->currentItem() == qi && n->selimage.vt != VT_EMPTY) img = n->selimage;
        else if (qi->isExpanded() && n->expimage.vt != VT_EMPTY) img = n->expimage;
        qi->setIcon(0, pics ? icon_from(list(), img) : QIcon());
        QFont f = tw->font(); f.setBold(n->bold); qi->setFont(0, f);
        if (n->fore >= 0) qi->setForeground(0, qcolor(n->fore)); else qi->setData(0, Qt::ForegroundRole, QVariant());
        if (n->back >= 0) qi->setBackground(0, qcolor(n->back)); else qi->setData(0, Qt::BackgroundRole, QVariant());
        qi->setData(0, Qt::UserRole + 2, qi->text(0));
        editing = false;
    }
    void look()
    {
        if (!tw) return;
        tw->setRootIsDecorated(linestyle == 1 || style >= 2);
        tw->setItemsExpandable(true);
        tw->manual_only = labeledit == 1;
        tw->setEditTriggers(labeledit == 1 ? QAbstractItemView::NoEditTriggers : QAbstractItemView::SelectedClicked | QAbstractItemView::EditKeyPressed);
        for (OcxItem *i : nodes.items) paint((Node *)i);
    }
    void load(const char *const *p) override
    {
        prop(p, "ImageList", &ilname);
        style = (int)propn(p, "Style", 7); linestyle = (int)propn(p, "LineStyle", 0);
        labeledit = (int)propn(p, "LabelEdit", 0);
        prop(p, "PathSeparator", &pathsep);
        if (pathsep.isEmpty()) pathsep = "\\";
        sorted = propn(p, "Sorted", 0) != 0;
        checkboxes = propn(p, "Checkboxes", 0) != 0;
        double ind = propn(p, "Indentation", 0);
        if (ind > 0 && tw) tw->setIndentation((int)std::lround(ind / TPP));
        if (tw) tw->setSelectionMode(QAbstractItemView::SingleSelection);
        look();
    }
    void remove_tree(Node *n)                   /* Nodes.Remove: the node with all under it */
    {
        std::vector<Node *> all;
        std::function<void(QTreeWidgetItem *)> walk = [&](QTreeWidgetItem *q) {
            for (int i = 0; i < q->childCount(); i++) walk(q->child(i));
            if (Node *x = node_of(q)) all.push_back(x);
        };
        walk(n->qi);
        QTreeWidgetItem *q = n->qi;
        for (Node *x : all) {
            int i = nodes.index_of(x);
            x->qi = nullptr;
            if (i >= 0) { nodes.items.erase(nodes.items.begin() + i); x->list = nullptr; vb_orelease(x->obj); }
        }
        delete q;
    }
    std::vector<Node *> children(QTreeWidgetItem *q) const
    {
        std::vector<Node *> r;
        int n = q ? q->childCount() : (tw ? tw->topLevelItemCount() : 0);
        for (int i = 0; i < n; i++) r.push_back(node_of(q ? q->child(i) : tw->topLevelItem(i)));
        return r;
    }
    void sort_level(QTreeWidgetItem *q)
    {
        if (!tw) return;
        if (q) q->sortChildren(0, Qt::AscendingOrder); else tw->sortItems(0, Qt::AscendingOrder);
        tw->setSortingEnabled(false);
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        if (n == "nodes") return list_ret(nodes, argc, argv);
        if (n == "imagelist") { vbvar r = imagelist_prop(c, &il, &ilname, set, val); if (set) look(); return r; }
        if (n == "selecteditem") {
            if (!set) { Node *x = tw ? node_of(tw->currentItem()) : nullptr; return retobj(x ? x->obj : nullptr); }
            int i = val.vt == VT_DISPATCH && !val.u.o ? -1 : nodes.find(val);
            if (vb_errpend) return vb_v_empty();
            if (tw) tw->setCurrentItem(i >= 0 ? ((Node *)nodes.items[i])->qi : nullptr);
            return vb_v_empty();
        }
        if (n == "style" || n == "linestyle" || n == "labeledit") {
            int *v = n == "style" ? &style : n == "linestyle" ? &linestyle : &labeledit;
            if (!set) return vb_v_l(*v);
            *v = vb_v2i(val); look();
            return vb_v_empty();
        }
        if (n == "indentation") {
            if (!set) return vb_v_f((float)from_twips_ctl(c, (tw ? tw->indentation() : 19) * TPP));
            if (tw) tw->setIndentation((int)std::lround(to_twips_ctl(c, vb_v2d(val)) / TPP));
            return vb_v_empty();
        }
        if (n == "pathseparator") { if (!set) return vstr(pathsep); pathsep = qs(vb_v2s(val)); return vb_v_empty(); }
        if (n == "sorted") { if (!set) return vbool(sorted); sorted = vb_v2bool(val) != 0; if (sorted) sort_level(nullptr); return vb_v_empty(); }
        if (n == "checkboxes") {
            if (!set) return vbool(checkboxes);
            checkboxes = vb_v2bool(val) != 0;
            editing = true;
            for (OcxItem *i : nodes.items) { auto *x = (Node *)i; if (x->qi) { if (checkboxes) x->qi->setCheckState(0, Qt::Unchecked); else x->qi->setData(0, Qt::CheckStateRole, QVariant()); } }
            editing = false;
            return vb_v_empty();
        }
        if (n == "getvisiblecount") return vb_v_l(tw && tw->topLevelItemCount() ? std::max(1, tw->viewport()->height() / std::max(1, tw->visualItemRect(tw->topLevelItem(0)).height())) : 0);
        if (n == "hittest") {
            if (!tw) return retobj(nullptr);
            double x = arg_d(argc, argv, 0), y = arg_d(argc, argv, 1);
            QPoint pt((int)(to_twips(c, x, false, false) / TPP), (int)(to_twips(c, y, true, false) / TPP));
            Node *nd = node_of(tw->itemAt(tw->viewport()->mapFrom(tw, pt)));
            return retobj(nd ? nd->obj : nullptr);
        }
        if (n == "startlabeledit") {
            if (tw && tw->currentItem()) {
                tw->allow = true;
                editing = true; tw->currentItem()->setFlags(tw->currentItem()->flags() | Qt::ItemIsEditable); editing = false;
                tw->editItem(tw->currentItem(), 0);
                tw->allow = false;
            }
            return vb_v_empty();
        }
        if (n == "dropHighlight" || n == "drophighlight") return set ? vb_v_empty() : retobj(nullptr);
        if (n == "hidesellection" || n == "hideselection" || n == "fullrowselect" || n == "hottracking" || n == "singlesel" || n == "scroll")
            return set ? vb_v_empty() : vbool(n == "hideselection" || n == "scroll");
        *found = 0;
        return vb_v_empty();
    }
    /* Nodes.Add [Relative], [Relationship], [Key], [Text], [Image], [SelectedImage] */
    OcxItem *add(vb_i32 argc, vbvar *argv)
    {
        if (!tw) return nullptr;
        Node *rel = nullptr;
        if (given(argc, argv, 0)) {
            int i = nodes.find(argv[0]);
            if (i < 0) return nullptr;
            rel = (Node *)nodes.items[i];
        }
        int how = (int)arg_l(argc, argv, 1, rel ? 2 : 1);
        if (how < 0 || how > 4) { vb_raise(5); return nullptr; }
        QString key = given(argc, argv, 2) ? arg_s(argc, argv, 2) : QString();
        OcxImageList *imgl = list();
        if (given(argc, argv, 4) && !image_ok(imgl, argv[4])) return nullptr;
        if (given(argc, argv, 5) && !image_ok(imgl, argv[5])) return nullptr;
        auto *nd = new Node;
        if (!nodes.insert(nd, -1, key)) return nullptr;
        editing = true;
        auto *qi = new QTreeWidgetItem;
        qi->setText(0, arg_s(argc, argv, 3));
        qi->setData(0, Qt::UserRole, QVariant::fromValue((void *)nd));
        qi->setFlags(qi->flags() | (labeledit == 0 ? Qt::ItemIsEditable : Qt::NoItemFlags));
        if (checkboxes) { qi->setCheckState(0, Qt::Unchecked); qi->setData(0, Qt::UserRole + 1, (int)Qt::Unchecked); }
        nd->qi = qi;
        QTreeWidgetItem *parent = nullptr;
        int at = -1;
        if (rel) {
            if (how == 4) parent = rel->qi;
            else {
                parent = rel->qi->parent();
                int ri = parent ? parent->indexOfChild(rel->qi) : tw->indexOfTopLevelItem(rel->qi);
                at = how == 0 ? 0 : how == 1 ? -1 : how == 2 ? ri + 1 : ri;
            }
        } else at = how == 0 ? 0 : -1;
        if (parent) { if (at < 0) parent->addChild(qi); else parent->insertChild(at, qi); }
        else { if (at < 0) tw->addTopLevelItem(qi); else tw->insertTopLevelItem(at, qi); }
        if (given(argc, argv, 4)) vb_vset(&nd->image, argv[4]);
        if (given(argc, argv, 5)) vb_vset(&nd->selimage, argv[5]);
        editing = false;
        paint(nd);
        Node *pn = node_of(parent);
        if ((pn && pn->sorted) || (!parent && sorted)) sort_level(parent);
        return nd;
    }
};
static OcxTreeView *node_owner(Node *n) { return n->list ? dynamic_cast<OcxTreeView *>(n->list->owner) : nullptr; }
vbvar Node::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    OcxTreeView *t = node_owner(this);
    if (!t || !qi) { ocx_error(35605); return vb_v_empty(); }
    auto obj_of = [](QTreeWidgetItem *q) { Node *x = node_of(q); return retobj(x ? x->obj : nullptr); };
    QTreeWidgetItem *par = qi->parent();
    int pos = par ? par->indexOfChild(qi) : t->tw->indexOfTopLevelItem(qi);
    int cnt = par ? par->childCount() : t->tw->topLevelItemCount();
    auto sib = [&](int i) { return par ? par->child(i) : t->tw->topLevelItem(i); };
    bool ch = false;
    if (n == "text") {
        if (!set) return vstr(qi->text(0));
        t->editing = true; qi->setText(0, qs(vb_v2s(val))); qi->setData(0, Qt::UserRole + 2, qi->text(0)); t->editing = false;
        return vb_v_empty();
    }
    if (n == "children") { if (set) { vb_raise(383); return vb_v_empty(); } return vb_v_i((vb_i16)qi->childCount()); }
    if (n == "child") return obj_of(qi->childCount() ? qi->child(0) : nullptr);
    if (n == "firstsibling") return obj_of(sib(0));
    if (n == "lastsibling") return obj_of(sib(cnt - 1));
    if (n == "next") return obj_of(pos + 1 < cnt ? sib(pos + 1) : nullptr);
    if (n == "previous") return obj_of(pos > 0 ? sib(pos - 1) : nullptr);
    if (n == "root") {      /* VB6 gives the first node of the top level, not the node's own root */
        QTreeWidgetItem *r = t->tw->topLevelItemCount() ? t->tw->topLevelItem(0) : qi;
        return obj_of(r);
    }
    if (n == "parent") {
        if (!set) return obj_of(par);
        Node *np = nullptr;
        if (val.vt == VT_DISPATCH && val.u.o) {
            int i = t->nodes.find(val);
            if (i < 0) return vb_v_empty();
            np = (Node *)t->nodes.items[i];
            for (QTreeWidgetItem *q = np->qi; q; q = q->parent())
                if (q == qi) { vb_raise_full(35614, vs(QString("TreeView")), vs(QString("Can't set Parent to a descendant"))); return vb_v_empty(); }
        }
        bool exp = qi->isExpanded();
        if (par) par->removeChild(qi); else t->tw->takeTopLevelItem(pos);
        if (np) np->qi->addChild(qi); else t->tw->addTopLevelItem(qi);
        qi->setExpanded(exp);
        return vb_v_empty();
    }
    if (n == "expanded") { if (!set) return vbool(qi->isExpanded()); qi->setExpanded(vb_v2bool(val) != 0); return vb_v_empty(); }
    if (n == "selected") {
        if (!set) return vbool(t->tw->currentItem() == qi);
        if (vb_v2bool(val)) t->tw->setCurrentItem(qi);
        return vb_v_empty();
    }
    if (n == "visible") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        bool vis = true;
        for (QTreeWidgetItem *q = par; q; q = q->parent()) if (!q->isExpanded()) vis = false;
        return vbool(vis);
    }
    if (n == "ensurevisible") {
        for (QTreeWidgetItem *q = par; q; q = q->parent()) q->setExpanded(true);
        t->tw->scrollToItem(qi);
        return vbool(true);
    }
    if (n == "fullpath") {
        QStringList parts;
        for (QTreeWidgetItem *q = qi; q; q = q->parent()) parts.prepend(q->text(0));
        return vstr(parts.join(t->pathsep));
    }
    if (n == "image") { vbvar r = image_prop(&image, t->list(), set, val, &ch); if (ch) t->paint(this); return r; }
    if (n == "selectedimage") { vbvar r = image_prop(&selimage, t->list(), set, val, &ch); if (ch) t->paint(this); return r; }
    if (n == "expandedimage") { vbvar r = image_prop(&expimage, t->list(), set, val, &ch); if (ch) t->paint(this); return r; }
    if (n == "bold") { if (!set) return vbool(bold); bold = vb_v2bool(val) != 0; t->paint(this); return vb_v_empty(); }
    if (n == "forecolor" || n == "backcolor") {
        long *p = n == "forecolor" ? &fore : &back;
        if (!set) return vb_v_l((vb_i32)(*p >= 0 ? *p : n == "forecolor" ? 0x80000008L : 0x80000005L));
        *p = item_color(val); t->paint(this);
        return vb_v_empty();
    }
    if (n == "checked") {
        if (!set) return vbool(qi->checkState(0) == Qt::Checked);
        t->editing = true;
        qi->setCheckState(0, vb_v2bool(val) ? Qt::Checked : Qt::Unchecked);
        qi->setData(0, Qt::UserRole + 1, (int)qi->checkState(0));
        t->editing = false;
        return vb_v_empty();
    }
    if (n == "sorted") {
        if (!set) return vbool(sorted);
        sorted = vb_v2bool(val) != 0;
        if (sorted) qi->sortChildren(0, Qt::AscendingOrder);
        return vb_v_empty();
    }
    if (n == "createdragimage") return retobj(new_pic(QImage(), 0));
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- ListView */
struct OcxListView;
struct SubItem : OcxItem {
    QString text;
    vbvar icon{};
    long fore = -1;
    bool bold = false;
    ~SubItem() { vb_vfree(&icon); }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
struct ListItem : OcxItem {
    QString text, tip;
    vbvar icon{}, smallicon{};
    OcxList subs;                               /* ListSubItems */
    bool selected = false, checked = false, ghosted = false, bold = false;
    long fore = -1;
    QTreeWidgetItem *ti = nullptr;
    QListWidgetItem *li = nullptr;
    ~ListItem() { vb_vfree(&icon); vb_vfree(&smallicon); }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
struct ColHeader : OcxItem {
    QString text;
    double width = 1440;
    int align = 0, position = 0;
    vbvar icon{};
    ~ColHeader() { vb_vfree(&icon); }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
static ListItem *li_of(QTreeWidgetItem *q) { return q ? (ListItem *)q->data(0, Qt::UserRole).value<void *>() : nullptr; }
static ListItem *li_of(QListWidgetItem *q) { return q ? (ListItem *)q->data(Qt::UserRole).value<void *>() : nullptr; }
struct OcxListView : Ocx {
    OcxList items, cols;
    vbobj il_big = nullptr, il_small = nullptr, il_head = nullptr;
    QString big_name, small_name, head_name;
    int view = 0, sortkey = 0, sortorder = 0, labeledit = 0, arrange = 0;
    bool sorted = false, multisel = false, fullrow = false, grid = false, checkboxes = false, hideheaders = false, syncing = false;
    int textbg = 0, picalign = 0;               /* TextBackground, PictureAlignment */
    vbobj bgpic = nullptr;                      /* Picture (the background) */
    QPointer<QStackedWidget> stack;
    QPointer<QTreeWidget> report;
    QPointer<QListWidget> icons;
    ~OcxListView() { if (il_big) vb_orelease(il_big); if (il_small) vb_orelease(il_small); if (il_head) vb_orelease(il_head); if (bgpic) vb_orelease(bgpic); }
    OcxImageList *big() { return imagelist_of(c, &il_big, big_name); }
    OcxImageList *small() { return imagelist_of(c, &il_small, small_name); }
    QWidget *make(QWidget *parent) override
    {
        stack = new QStackedWidget(parent);
        report = new QTreeWidget(stack);
        report->setRootIsDecorated(false);
        report->setUniformRowHeights(true);
        report->header()->setSectionsClickable(true);
        report->header()->setStretchLastSection(false);
        icons = new QListWidget(stack);
        stack->addWidget(icons);
        stack->addWidget(report);
        vbobj self = c->self;
        auto me = [self]() -> OcxListView * { Ctl *x = C(self); return x ? dynamic_cast<OcxListView *>(x->ocx) : nullptr; };
        auto clicked = [self](ListItem *it) { if (it && it->obj) { vbvar a = vb_v_obj(it->obj); fire_ctl(C(self), "itemclick", 1, &a); } };
        QObject::connect(report, &QTreeWidget::itemClicked, [clicked](QTreeWidgetItem *q, int) { clicked(li_of(q)); });
        QObject::connect(icons, &QListWidget::itemClicked, [clicked](QListWidgetItem *q) { clicked(li_of(q)); });
        QObject::connect(report->header(), &QHeaderView::sectionClicked, [self, me](int i) {
            OcxListView *l = me();
            if (!l || i < 0 || i >= (int)l->cols.items.size()) return;
            vbvar a = vb_v_obj(l->cols.items[i]->obj);
            fire_ctl(C(self), "columnclick", 1, &a);
        });
        QObject::connect(report, &QTreeWidget::itemChanged, [self, me](QTreeWidgetItem *q, int) {
            OcxListView *l = me(); ListItem *it = li_of(q);
            if (!l || !it || l->syncing || !l->checkboxes) return;
            bool on = q->checkState(0) == Qt::Checked;
            if (on == it->checked) return;
            it->checked = on;
            vbvar a = vb_v_obj(it->obj); fire_ctl(C(self), "itemcheck", 1, &a);
        });
        auto selsync = [me]() { if (OcxListView *l = me()) l->read_selection(); };
        QObject::connect(report, &QTreeWidget::itemSelectionChanged, selsync);
        QObject::connect(icons, &QListWidget::itemSelectionChanged, selsync);
        return stack;
    }
    void read_selection()
    {
        if (syncing) return;
        for (OcxItem *i : items.items) {
            auto *it = (ListItem *)i;
            it->selected = it->ti ? it->ti->isSelected() : it->li ? it->li->isSelected() : it->selected;
        }
    }
    QString sub_text(ListItem *it, int col) const
    {
        if (col == 0) return it->text;
        return col - 1 < (int)it->subs.items.size() ? ((SubItem *)it->subs.items[col - 1])->text : QString();
    }
    void fill(ListItem *it)                     /* one item in the widget */
    {
        if (view == 3) {
            if (!it->ti) { it->ti = new QTreeWidgetItem; it->ti->setData(0, Qt::UserRole, QVariant::fromValue((void *)it)); }
            QTreeWidgetItem *q = it->ti;
            int nc = std::max(1, (int)cols.items.size());
            for (int col = 0; col < nc; col++) {
                q->setText(col, sub_text(it, col));
                int al = col < (int)cols.items.size() ? ((ColHeader *)cols.items[col])->align : 0;
                q->setData(col, Qt::TextAlignmentRole, (int)((al == 1 ? Qt::AlignRight : al == 2 ? Qt::AlignHCenter : Qt::AlignLeft) | Qt::AlignVCenter));
            }
            q->setIcon(0, icon_from(small(), it->smallicon));
            if (checkboxes) q->setCheckState(0, it->checked ? Qt::Checked : Qt::Unchecked);
            QFont f = report->font(); f.setBold(it->bold); q->setFont(0, f);
            if (it->fore >= 0) q->setForeground(0, qcolor(it->fore));
            q->setSelected(it->selected);
            q->setToolTip(0, it->tip);
        } else {
            if (!it->li) { it->li = new QListWidgetItem; it->li->setData(Qt::UserRole, QVariant::fromValue((void *)it)); }
            QListWidgetItem *q = it->li;
            q->setText(it->text);
            q->setIcon(view == 0 ? icon_from(big(), it->icon) : icon_from(small(), it->smallicon));
            if (checkboxes) q->setCheckState(it->checked ? Qt::Checked : Qt::Unchecked);
            QFont f = icons->font(); f.setBold(it->bold); q->setFont(f);
            if (it->fore >= 0) q->setForeground(qcolor(it->fore));
            q->setSelected(it->selected);
            q->setToolTip(it->tip);
        }
    }
    void rebuild()                              /* the view changed, or the items were sorted */
    {
        if (!stack) return;
        syncing = true;
        read_selection();
        for (OcxItem *i : items.items) {
            auto *it = (ListItem *)i;
            if (it->ti && report) { int k = report->indexOfTopLevelItem(it->ti); if (k >= 0) report->takeTopLevelItem(k); delete it->ti; it->ti = nullptr; }
            if (it->li && icons) { int k = icons->row(it->li); if (k >= 0) icons->takeItem(k); delete it->li; it->li = nullptr; }
        }
        headers();
        stack->setCurrentWidget(view == 3 ? (QWidget *)report.data() : (QWidget *)icons.data());
        if (view != 3) {
            icons->setViewMode(view == 0 ? QListView::IconMode : QListView::ListMode);
            icons->setFlow(view == 2 ? QListView::TopToBottom : view == 1 ? QListView::LeftToRight : QListView::LeftToRight);
            icons->setWrapping(view != 2 || true);
            icons->setResizeMode(QListView::Adjust);
            icons->setMovement(QListView::Static);
            OcxImageList *bl = big();
            icons->setIconSize(view == 0 && bl ? QSize(bl->iw ? bl->iw : 32, bl->ih ? bl->ih : 32) : QSize(16, 16));
            if (view == 0) icons->setGridSize(QSize(std::max(76, icons->iconSize().width() + 40), icons->iconSize().height() + 36));
            else icons->setGridSize(QSize());
            icons->setWordWrap(true);
        }
        for (OcxItem *i : items.items) {
            auto *it = (ListItem *)i;
            fill(it);
            if (view == 3) report->addTopLevelItem(it->ti); else icons->addItem(it->li);
        }
        syncing = false;
    }
    void headers()
    {
        if (!report) return;
        int nc = std::max(1, (int)cols.items.size());
        report->setColumnCount(nc);
        QStringList labels;
        for (int i = 0; i < nc; i++) labels << (i < (int)cols.items.size() ? ((ColHeader *)cols.items[i])->text : QString());
        report->setHeaderLabels(labels);
        report->setHeaderHidden(hideheaders || cols.items.empty());
        OcxImageList *hl = imagelist_of(c, &il_head, head_name);
        for (int i = 0; i < (int)cols.items.size(); i++) {
            auto *h = (ColHeader *)cols.items[i];
            report->setColumnWidth(i, (int)std::lround(h->width / TPP));
            int al = h->align == 1 ? Qt::AlignRight : h->align == 2 ? Qt::AlignHCenter : Qt::AlignLeft;
            report->headerItem()->setData(i, Qt::TextAlignmentRole, (int)(al | Qt::AlignVCenter));
            report->headerItem()->setIcon(i, icon_from(hl, h->icon));
        }
        report->setSelectionMode(multisel ? QAbstractItemView::ExtendedSelection : QAbstractItemView::SingleSelection);
        icons->setSelectionMode(multisel ? QAbstractItemView::ExtendedSelection : QAbstractItemView::SingleSelection);
        report->setSelectionBehavior(fullrow ? QAbstractItemView::SelectRows : QAbstractItemView::SelectItems);
        report->setAlternatingRowColors(false);
        report->setStyleSheet(grid ? "QTreeView::item { border-right: 1px solid #d0d0d0; border-bottom: 1px solid #d0d0d0; }" : "");
    }
    void sort()
    {
        if (!sorted) return;
        int key = sortkey, ord = sortorder;
        std::stable_sort(items.items.begin(), items.items.end(), [this, key, ord](OcxItem *a, OcxItem *b) {
            int r = QString::compare(sub_text((ListItem *)a, key), sub_text((ListItem *)b, key), Qt::CaseInsensitive);
            return ord ? r > 0 : r < 0;
        });
        rebuild();
    }
    void load(const char *const *p) override
    {
        prop(p, "Icons", &big_name); prop(p, "SmallIcons", &small_name); prop(p, "ColumnHeaderIcons", &head_name);
        view = (int)propn(p, "View", 0);
        sortkey = (int)propn(p, "SortKey", 0); sortorder = (int)propn(p, "SortOrder", 0);
        sorted = propn(p, "Sorted", 0) != 0;
        labeledit = (int)propn(p, "LabelEdit", 0);
        multisel = propn(p, "MultiSelect", 0) != 0;
        fullrow = propn(p, "FullRowSelect", 0) != 0;
        grid = propn(p, "GridLines", 0) != 0;
        checkboxes = propn(p, "Checkboxes", 0) != 0;
        hideheaders = propn(p, "HideColumnHeaders", 0) != 0;
        int n = (int)propn(p, "NumItems", 0);
        for (int i = 1; i <= n; i++) {
            QString base = QString("ColumnHeader(%1)").arg(i), key, tag;
            auto *h = new ColHeader;
            iprop(p, base, "Text", &h->text);
            iprop(p, base, "Key", &key);
            if (iprop(p, base, "Tag", &tag)) { vbvar t = vb_v_s(vs(tag)); vb_vset(&h->tag, t); }
            h->width = ipropn(p, base, "Width", 1440);
            h->align = (int)ipropn(p, base, "Alignment", 0);
            h->position = i;
            if (!cols.insert(h, -1, key)) break;
        }
        rebuild();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
static OcxListView *lv_owner(OcxItem *it)
{
    if (!it->list) return nullptr;
    return dynamic_cast<OcxListView *>(it->list->owner);
}
vbvar OcxListView::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    if (n == "listitems") return list_ret(items, argc, argv);
    if (n == "columnheaders") return list_ret(cols, argc, argv);
    if (n == "icons") { vbvar r = imagelist_prop(c, &il_big, &big_name, set, val); if (set) rebuild(); return r; }
    if (n == "smallicons") { vbvar r = imagelist_prop(c, &il_small, &small_name, set, val); if (set) rebuild(); return r; }
    if (n == "columnheadericons") { vbvar r = imagelist_prop(c, &il_head, &head_name, set, val); if (set) headers(); return r; }
    if (n == "view") {
        if (!set) return vb_v_l(view);
        int v = vb_v2i(val);
        if (v < 0 || v > 3) { vb_raise(380); return vb_v_empty(); }
        view = v; rebuild();
        return vb_v_empty();
    }
    if (n == "sorted") { if (!set) return vbool(sorted); sorted = vb_v2bool(val) != 0; sort(); return vb_v_empty(); }
    if (n == "sortkey") { if (!set) return vb_v_i((vb_i16)sortkey); sortkey = vb_v2i(val); sort(); return vb_v_empty(); }
    if (n == "sortorder") { if (!set) return vb_v_l(sortorder); sortorder = vb_v2i(val); sort(); return vb_v_empty(); }
    if (n == "multiselect" || n == "fullrowselect" || n == "gridlines" || n == "checkboxes" || n == "hidecolumnheaders") {
        bool *b = n == "multiselect" ? &multisel : n == "fullrowselect" ? &fullrow : n == "gridlines" ? &grid : n == "checkboxes" ? &checkboxes : &hideheaders;
        if (!set) return vbool(*b);
        *b = vb_v2bool(val) != 0;
        rebuild();
        return vb_v_empty();
    }
    if (n == "labeledit") { if (!set) return vb_v_l(labeledit); labeledit = vb_v2i(val); return vb_v_empty(); }
    if (n == "arrange") { if (!set) return vb_v_l(arrange); arrange = vb_v2i(val); return vb_v_empty(); }
    if (n == "selecteditem") {
        if (!set) {
            read_selection();
            QTreeWidgetItem *rq = view == 3 && report ? report->currentItem() : nullptr;
            QListWidgetItem *iq = view != 3 && icons ? icons->currentItem() : nullptr;
            ListItem *it = rq ? li_of(rq) : iq ? li_of(iq) : nullptr;
            if (!it || !it->selected) for (OcxItem *i : items.items) if (((ListItem *)i)->selected) { it = (ListItem *)i; break; }
            return retobj(it ? it->obj : nullptr);
        }
        if (val.vt == VT_DISPATCH && !val.u.o) { syncing = true; if (report) report->clearSelection(); if (icons) icons->clearSelection(); syncing = false; read_selection(); return vb_v_empty(); }
        int i = items.find(val);
        if (i < 0) return vb_v_empty();
        auto *it = (ListItem *)items.items[i];
        syncing = true;
        if (!multisel) for (OcxItem *o : items.items) { auto *x = (ListItem *)o; x->selected = false; if (x->ti) x->ti->setSelected(false); if (x->li) x->li->setSelected(false); }
        it->selected = true;
        if (it->ti) { report->setCurrentItem(it->ti); it->ti->setSelected(true); }
        if (it->li) { icons->setCurrentItem(it->li); it->li->setSelected(true); }
        syncing = false;
        return vb_v_empty();
    }
    if (n == "finditem") {                      /* FindItem(sz, [where 0 text 1 subitem 2 tag], [index], [match 0 whole 1 partial]) */
        QString sz = arg_s(argc, argv, 0);
        int where = (int)arg_l(argc, argv, 1, 0), start = 0, match = (int)arg_l(argc, argv, 3, 1);
        if (given(argc, argv, 2)) { start = items.find(argv[2]); if (start < 0) return vb_v_empty(); }
        for (int i = start; i < (int)items.items.size(); i++) {
            auto *it = (ListItem *)items.items[i];
            QStringList cands;
            if (where == 0) cands << it->text;
            else if (where == 1) for (OcxItem *s : it->subs.items) cands << ((SubItem *)s)->text;
            else cands << qs(vb_v2s(it->tag));
            for (const QString &cnd : cands)
                if (match == 0 ? !QString::compare(cnd, sz, Qt::CaseInsensitive) : cnd.startsWith(sz, Qt::CaseInsensitive)) return retobj(it->obj);
        }
        return retobj(nullptr);
    }
    if (n == "hittest") {
        QPoint pt((int)(to_twips(c, arg_d(argc, argv, 0), false, false) / TPP), (int)(to_twips(c, arg_d(argc, argv, 1), true, false) / TPP));
        ListItem *it = nullptr;
        if (view == 3 && report) it = li_of(report->itemAt(report->viewport()->mapFrom(stack, pt)));
        else if (icons) it = li_of(icons->itemAt(icons->viewport()->mapFrom(stack, pt)));
        return retobj(it ? it->obj : nullptr);
    }
    if (n == "getfirstvisible") {
        ListItem *it = view == 3 && report ? li_of(report->itemAt(1, 1)) : icons ? li_of(icons->itemAt(1, 1)) : nullptr;
        if (!it && !items.items.empty()) it = (ListItem *)items.items[0];
        return retobj(it ? it->obj : nullptr);
    }
    if (n == "startlabeledit" || n == "refresh") return vb_v_empty();
    if (n == "drophighlight") return set ? vb_v_empty() : retobj(nullptr);
    if (n == "hideselection" || n == "labelwrap" || n == "hottracking" || n == "allowcolumnreorder" || n == "flatscrollbar")
        return set ? vb_v_empty() : vbool(n == "hideselection" || n == "labelwrap");
    if (n == "textbackground") { if (!set) return vb_v_l(textbg); textbg = vb_v2l(val); return vb_v_empty(); }
    if (n == "picturealignment") { if (!set) return vb_v_l(picalign); picalign = vb_v2l(val); return vb_v_empty(); }
    if (n == "picture") {                       /* the background picture: kept (not drawn) */
        if (!set) return retobj(bgpic);
        vbobj o = val.vt == VT_DISPATCH ? val.u.o : nullptr;
        if (o) vb_oaddref(o);
        if (bgpic) vb_orelease(bgpic);
        bgpic = o;
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
vbvar ListItem::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    OcxListView *l = lv_owner(this);
    bool ch = false;
    if (n == "text") { if (!set) return vstr(text); text = qs(vb_v2s(val)); if (l && l->sorted && l->sortkey == 0) l->sort(); else if (l) l->fill(this); return vb_v_empty(); }
    if (n == "subitems") {                      /* SubItems(i) */
        vb_i32 i = (vb_i32)arg_l(argc - (set ? 1 : 0), argv, 0, 0);
        int ncols = l ? (int)l->cols.items.size() : 0;
        if (i < 1 || i >= std::max(ncols, (int)subs.items.size() + 1)) { vb_raise(380); return vb_v_empty(); }
        while ((int)subs.items.size() < i) {
            auto *s = new SubItem;
            if (!subs.insert(s, -1, QString())) return vb_v_empty();
        }
        auto *s = (SubItem *)subs.items[i - 1];
        if (!set) return vstr(s->text);
        s->text = qs(vb_v2s(val));
        if (l) l->fill(this);
        return vb_v_empty();
    }
    if (n == "listsubitems") return list_ret(subs, argc, argv);
    if (n == "icon") { vbvar r = image_prop(&icon, l ? l->big() : nullptr, set, val, &ch); if (ch && l) l->fill(this); return r; }
    if (n == "smallicon") { vbvar r = image_prop(&smallicon, l ? l->small() : nullptr, set, val, &ch); if (ch && l) l->fill(this); return r; }
    if (n == "selected") {
        if (l) l->read_selection();
        if (!set) return vbool(selected);
        selected = vb_v2bool(val) != 0;
        if (l) { l->syncing = true; if (ti) ti->setSelected(selected); if (li) li->setSelected(selected); l->syncing = false; }
        return vb_v_empty();
    }
    if (n == "checked") { if (!set) return vbool(checked); checked = vb_v2bool(val) != 0; if (l) { l->syncing = true; l->fill(this); l->syncing = false; } return vb_v_empty(); }
    if (n == "ghosted") { if (!set) return vbool(ghosted); ghosted = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "bold") { if (!set) return vbool(bold); bold = vb_v2bool(val) != 0; if (l) l->fill(this); return vb_v_empty(); }
    if (n == "forecolor") { if (!set) return vb_v_l((vb_i32)(fore >= 0 ? fore : 0x80000008L)); fore = item_color(val); if (l) l->fill(this); return vb_v_empty(); }
    if (n == "tooltiptext") { if (!set) return vstr(tip); tip = qs(vb_v2s(val)); if (l) l->fill(this); return vb_v_empty(); }
    if (n == "ensurevisible") {
        if (l && ti) l->report->scrollToItem(ti);
        if (l && li) l->icons->scrollToItem(li);
        return vbool(true);
    }
    if (n == "left" || n == "top" || n == "width" || n == "height") {
        if (set) { vb_raise(383); return vb_v_empty(); }
        QRect r;
        if (l && ti) r = l->report->visualItemRect(ti);
        if (l && li) r = l->icons->visualItemRect(li);
        double t = (n == "left" ? r.x() : n == "top" ? r.y() : n == "width" ? r.width() : r.height()) * TPP;
        return vb_v_f((float)(l ? from_twips(l->c, t, n == "top" || n == "height", false) : t));
    }
    if (n == "createdragimage") return retobj(new_pic(QImage(), 0));
    *found = 0;
    return vb_v_empty();
}
vbvar SubItem::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    if (n == "text") { if (!set) return vstr(text); text = qs(vb_v2s(val)); return vb_v_empty(); }
    if (n == "reporticon") { if (!set) return vb_vtmp(icon); vb_vset(&icon, val); return vb_v_empty(); }
    if (n == "forecolor") { if (!set) return vb_v_l((vb_i32)(fore >= 0 ? fore : 0x80000008L)); fore = item_color(val); return vb_v_empty(); }
    if (n == "bold") { if (!set) return vbool(bold); bold = vb_v2bool(val) != 0; return vb_v_empty(); }
    if (n == "tooltiptext") return set ? vb_v_empty() : vstr(QString());
    *found = 0;
    return vb_v_empty();
}
vbvar ColHeader::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    OcxListView *l = lv_owner(this);
    if (n == "text") { if (!set) return vstr(text); text = qs(vb_v2s(val)); if (l) l->headers(); return vb_v_empty(); }
    if (n == "width") {
        if (!set) {
            int i = list ? list->index_of(this) : -1;
            double t = l && l->report && i >= 0 && l->view == 3 ? l->report->columnWidth(i) * TPP : width;
            return vb_v_f((float)(l ? from_twips_ctl(l->c, t) : t));
        }
        width = l ? to_twips_ctl(l->c, vb_v2d(val)) : vb_v2d(val);
        if (l) l->headers();
        return vb_v_empty();
    }
    if (n == "alignment") {
        if (!set) return vb_v_l(align);
        int i = list ? list->index_of(this) : -1;
        if (i == 0) { vb_raise_full(35631, vs(QString("ListView")), vs(QString("The first column must be left aligned"))); return vb_v_empty(); }
        align = vb_v2i(val);
        if (l) l->rebuild();
        return vb_v_empty();
    }
    if (n == "position") { if (!set) return vb_v_i((vb_i16)(list ? list->index_of(this) + 1 : position)); return vb_v_empty(); }
    if (n == "subitemindex") { if (set) { vb_raise(383); return vb_v_empty(); } return vb_v_i((vb_i16)(list ? list->index_of(this) : 0)); }
    if (n == "icon") { if (!set) return vb_vtmp(icon); vb_vset(&icon, val); if (l) l->headers(); return vb_v_empty(); }
    if (n == "left") {
        int i = list ? list->index_of(this) : -1;
        double t = l && l->report && i >= 0 ? l->report->header()->sectionPosition(i) * TPP : 0;
        return vb_v_f((float)(l ? from_twips_ctl(l->c, t) : t));
    }
    *found = 0;
    return vb_v_empty();
}
/* ListItems.Add [Index], [Key], [Text], [Icon], [SmallIcon] */
static OcxItem *listitem_add(OcxListView *l, vb_i32 argc, vbvar *argv)
{
    int pos; QString key;
    if (!add_place(&l->items, argc, argv, &pos, &key)) return nullptr;
    if (given(argc, argv, 3) && !image_ok(l->big(), argv[3])) return nullptr;
    if (given(argc, argv, 4) && !image_ok(l->small(), argv[4])) return nullptr;
    auto *it = new ListItem;
    it->text = arg_s(argc, argv, 2);
    if (given(argc, argv, 3)) vb_vset(&it->icon, argv[3]);
    if (given(argc, argv, 4)) vb_vset(&it->smallicon, argv[4]);
    it->subs.clsname = "ListSubItems"; it->subs.itemcls = "ListSubItem"; it->subs.itemdef = "text";
    it->subs.add = [it, l](vb_i32 ac, vbvar *av) -> OcxItem * {   /* ListSubItems.Add [Index], [Key], [Text], [ReportIcon], [ToolTipText] */
        int p; QString k;
        if (!add_place(&it->subs, ac, av, &p, &k)) return nullptr;
        auto *s = new SubItem;
        s->text = arg_s(ac, av, 2);
        if (given(ac, av, 3)) vb_vset(&s->icon, av[3]);
        if (!it->subs.insert(s, p, k)) return nullptr;
        if (it->list) l->fill(it);
        return s;
    };
    if (!l->items.insert(it, pos, key)) return nullptr;
    l->fill(it);
    if (l->view == 3) {
        int at = pos < 0 ? l->report->topLevelItemCount() : pos;
        l->report->insertTopLevelItem(at, it->ti);
    } else {
        int at = pos < 0 ? l->icons->count() : pos;
        l->icons->insertItem(at, it->li);
    }
    if (l->sorted) l->sort();
    return it;
}
/* ColumnHeaders.Add [Index], [Key], [Text], [Width], [Alignment], [Icon] */
static OcxItem *column_add(OcxListView *l, vb_i32 argc, vbvar *argv)
{
    int pos; QString key;
    if (!add_place(&l->cols, argc, argv, &pos, &key)) return nullptr;
    auto *h = new ColHeader;
    h->text = arg_s(argc, argv, 2);
    if (given(argc, argv, 3)) h->width = to_twips_ctl(l->c, arg_d(argc, argv, 3));
    h->align = (int)arg_l(argc, argv, 4, 0);
    if (given(argc, argv, 5)) vb_vset(&h->icon, argv[5]);
    if (!l->cols.insert(h, pos, key)) return nullptr;
    l->rebuild();
    return h;
}

/* ---------------------------------------------------------------- ImageCombo */
struct ComboItem : OcxItem {
    QString text;
    vbvar image{}, selimage{};
    int indent = 0;
    ~ComboItem() { vb_vfree(&image); vb_vfree(&selimage); }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
};
struct OcxImageCombo : Ocx {
    OcxList items;
    vbobj il = nullptr; QString ilname;
    QPointer<QComboBox> cb;
    bool locked = false;
    ~OcxImageCombo() { if (il) vb_orelease(il); }
    OcxImageList *list() { return imagelist_of(c, &il, ilname); }
    QWidget *make(QWidget *parent) override
    {
        cb = new QComboBox(parent);
        cb->setEditable(true);
        cb->setInsertPolicy(QComboBox::NoInsert);
        vbobj self = c->self;
        QObject::connect(cb, QOverload<int>::of(&QComboBox::activated), [self](int) { fire_ctl(C(self), "click"); });
        QObject::connect(cb->lineEdit(), &QLineEdit::textEdited, [self](const QString &) { fire_ctl(C(self), "change"); });
        return cb;
    }
    void sync()
    {
        if (!cb) return;
        QSignalBlocker b(cb);
        QString t = cb->currentText();
        int cur = cb->currentIndex();
        cb->clear();
        OcxImageList *imgl = list();
        if (imgl && imgl->iw) cb->setIconSize(QSize(imgl->iw, imgl->ih));
        for (OcxItem *i : items.items) {
            auto *x = (ComboItem *)i;
            cb->addItem(icon_from(imgl, x->image), QString(x->indent * 2, ' ') + x->text);
        }
        cb->setCurrentIndex(cur < cb->count() ? cur : -1);
        if (cb->currentIndex() < 0) cb->setEditText(t);
    }
    void load(const char *const *p) override
    {
        prop(p, "ImageList", &ilname);
        locked = propn(p, "Locked", 0) != 0;
        QString t;
        if (prop(p, "Text", &t) && cb) { QSignalBlocker b(cb); cb->setEditText(t); }
        if (cb) cb->lineEdit()->setReadOnly(locked);
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        if (n == "comboitems") return list_ret(items, argc, argv);
        if (n == "imagelist") { vbvar r = imagelist_prop(c, &il, &ilname, set, val); if (set) sync(); return r; }
        if (n == "text") { if (!set) return vstr(cb ? cb->currentText() : QString()); if (cb) { QSignalBlocker b(cb); cb->setCurrentIndex(-1); cb->setEditText(qs(vb_v2s(val))); } fire_ctl(c, "change"); return vb_v_empty(); }
        if (n == "selecteditem") {
            if (!set) { int i = cb ? cb->currentIndex() : -1; return retobj(i >= 0 && i < (int)items.items.size() ? items.items[i]->obj : nullptr); }
            int i = val.vt == VT_DISPATCH && !val.u.o ? -1 : items.find(val);
            if (vb_errpend) return vb_v_empty();
            if (cb) { QSignalBlocker b(cb); cb->setCurrentIndex(i); }
            return vb_v_empty();
        }
        if (n == "locked") { if (!set) return vbool(locked); locked = vb_v2bool(val) != 0; if (cb) cb->lineEdit()->setReadOnly(locked); return vb_v_empty(); }
        if (n == "getfirstvisible") return retobj(items.items.empty() ? nullptr : items.items[0]->obj);
        if (n == "indentation") return set ? vb_v_empty() : vb_v_i(0);
        *found = 0;
        return vb_v_empty();
    }
};
static OcxImageCombo *combo_owner(ComboItem *x) { return x->list ? dynamic_cast<OcxImageCombo *>(x->list->owner) : nullptr; }
vbvar ComboItem::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    OcxImageCombo *o = combo_owner(this);
    bool ch = false;
    if (n == "text") { if (!set) return vstr(text); text = qs(vb_v2s(val)); return vb_v_empty(); }
    if (n == "image") return image_prop(&image, o ? o->list() : nullptr, set, val, &ch);
    if (n == "selimage") return image_prop(&selimage, o ? o->list() : nullptr, set, val, &ch);
    if (n == "indentation") { if (!set) return vb_v_i((vb_i16)indent); indent = vb_v2i(val); return vb_v_empty(); }
    if (n == "selected") {
        int i = list ? list->index_of(this) : -1;
        if (!set) return vbool(o && o->cb && o->cb->currentIndex() == i);
        if (vb_v2bool(val) && o && o->cb) { QSignalBlocker b(o->cb); o->cb->setCurrentIndex(i); }
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
/* ComboItems.Add [Index], [Key], [Text], [Image], [SelImage], [Indentation] */
static OcxItem *comboitem_add(OcxImageCombo *o, vb_i32 argc, vbvar *argv)
{
    int pos; QString key;
    if (!add_place(&o->items, argc, argv, &pos, &key)) return nullptr;
    if (given(argc, argv, 3) && !image_ok(o->list(), argv[3])) return nullptr;
    auto *x = new ComboItem;
    x->text = arg_s(argc, argv, 2);
    if (given(argc, argv, 3)) vb_vset(&x->image, argv[3]);
    if (given(argc, argv, 4)) vb_vset(&x->selimage, argv[4]);
    x->indent = (int)arg_l(argc, argv, 5, 0);
    if (!o->items.insert(x, pos, key)) return nullptr;
    return x;
}
