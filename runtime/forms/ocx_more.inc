/* ocx_more.inc - UpDown, DTPicker, MonthView, FlatScrollBar (MSComCtl2), SSTab (TabDlg) and
 * MSFlexGrid (MSFlexGridLib), written again on Qt. Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

/* VB dates (days since 30 Dec 1899) and Qt's */
static QDateTime qdt(double d)
{
    QDate base(1899, 12, 30);
    double day = std::floor(d), frac = std::fabs(d - day);
    if (d < 0) { day = std::ceil(d); frac = std::fabs(d - day); }
    QDate dt = base.addDays((qint64)day);
    int secs = (int)std::lround(frac * 86400.0);
    if (secs >= 86400) { secs -= 86400; dt = dt.addDays(1); }
    return QDateTime(dt, QTime(0, 0).addSecs(secs));
}
static double vbdate(const QDateTime &t)
{
    QDate base(1899, 12, 30);
    double day = (double)base.daysTo(t.date());
    return day + QTime(0, 0).secsTo(t.time()) / 86400.0;
}

/* ---------------------------------------------------------------- UpDown */
struct OcxUpDown : Ocx {
    int min = 0, max = 10, value = 0, inc = 1, orientation = 0, alignment = 1;
    bool wrap = false, autobuddy = false, syncbuddy = false;
    QString buddy, buddyprop;
    QPointer<QWidget> wd;
    QWidget *make(QWidget *parent) override
    {
        wd = new QWidget(parent);
        auto *up = new QToolButton(wd), *dn = new QToolButton(wd);
        up->setObjectName("up"); dn->setObjectName("down");
        up->setAutoRepeat(true); dn->setAutoRepeat(true);
        vbobj self = c->self;
        auto step = [self](int dir) {
            Ctl *x = C(self);
            auto *o = x ? dynamic_cast<OcxUpDown *>(x->ocx) : nullptr;
            if (!o) return;
            o->bump(dir);
        };
        QObject::connect(up, &QToolButton::clicked, [step]() { step(1); });
        QObject::connect(dn, &QToolButton::clicked, [step]() { step(-1); });
        return wd;
    }
    void arrange()
    {
        if (!wd) return;
        auto *up = wd->findChild<QToolButton *>("up"), *dn = wd->findChild<QToolButton *>("down");
        int w = wd->width(), h = wd->height();
        if (orientation == 0) {
            up->setArrowType(Qt::UpArrow); dn->setArrowType(Qt::DownArrow);
            up->setGeometry(0, 0, w, h / 2); dn->setGeometry(0, h / 2, w, h - h / 2);
        } else {
            up->setArrowType(Qt::RightArrow); dn->setArrowType(Qt::LeftArrow);
            dn->setGeometry(0, 0, w / 2, h); up->setGeometry(w / 2, 0, w - w / 2, h);
        }
    }
    Ctl *buddy_ctl()
    {
        if (buddy.isEmpty() || !c->form) return nullptr;
        auto it = c->form->byname.find(lower(buddy.toUtf8().constData()));
        return it == c->form->byname.end() ? nullptr : C(it->second);
    }
    void push()                                 /* SyncBuddy: the buddy shows the Value */
    {
        Ctl *b = syncbuddy ? buddy_ctl() : nullptr;
        if (!b) return;
        vbvar v = vb_v_l(value);
        std::string p = buddyprop.isEmpty() || !buddyprop.compare("Default", Qt::CaseInsensitive) ? std::string("") : lower(buddyprop.toUtf8().constData());
        vb_late_o(b->self, p.c_str(), VBM_LET, 1, &v);
    }
    void bump(int dir)
    {
        int lo = std::min(min, max), hi = std::max(min, max);
        int sign = max >= min ? 1 : -1;
        int v = value + dir * inc * sign;
        if (v > hi) v = wrap ? lo : hi;
        if (v < lo) v = wrap ? hi : lo;
        fire_ctl(c, dir > 0 ? "upclick" : "downclick");
        if (v != value) { value = v; push(); fire_ctl(c, "change"); }
    }
    void load(const char *const *p) override
    {
        min = (int)propn(p, "Min", 0); max = (int)propn(p, "Max", 10); value = (int)propn(p, "Value", 0);
        inc = (int)propn(p, "Increment", 1); orientation = (int)propn(p, "Orientation", 0);
        alignment = (int)propn(p, "Alignment", 1);
        wrap = propn(p, "Wrap", 0) != 0; autobuddy = propn(p, "AutoBuddy", 0) != 0; syncbuddy = propn(p, "SyncBuddy", 0) != 0;
        prop(p, "BuddyControl", &buddy); prop(p, "BuddyProperty", &buddyprop);
        arrange();
        vbobj self = c->self;
        if (wd) wd->installEventFilter(new ResizeHook(wd, [self]() {
            Ctl *x = C(self); auto *o = x ? dynamic_cast<OcxUpDown *>(x->ocx) : nullptr;
            if (o) o->arrange();
        }));
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        struct { const char *nm; int *p; } ints[] = { { "min", &min }, { "max", &max }, { "increment", &inc }, { "orientation", &orientation }, { "alignment", &alignment } };
        for (auto &e : ints) if (n == e.nm) { if (!set) return vb_v_l(*e.p); *e.p = vb_v2l(val); arrange(); return vb_v_empty(); }
        if (n == "value") {
            if (!set) return vb_v_l(value);
            int v = vb_v2l(val);
            if (v < std::min(min, max) || v > std::max(min, max)) { vb_raise(380); return vb_v_empty(); }
            if (v != value) { value = v; push(); fire_ctl(c, "change"); }
            return vb_v_empty();
        }
        if (n == "wrap") { if (!set) return vbool(wrap); wrap = vb_v2bool(val) != 0; return vb_v_empty(); }
        if (n == "syncbuddy") { if (!set) return vbool(syncbuddy); syncbuddy = vb_v2bool(val) != 0; return vb_v_empty(); }
        if (n == "autobuddy") { if (!set) return vbool(autobuddy); autobuddy = vb_v2bool(val) != 0; return vb_v_empty(); }
        if (n == "buddyproperty") { if (!set) return vstr(buddyprop); buddyprop = qs(vb_v2s(val)); return vb_v_empty(); }
        if (n == "buddycontrol") {
            if (!set) { Ctl *b = buddy_ctl(); return b ? retobj(b->self) : vstr(QString()); }
            if (val.vt == VT_DISPATCH && val.u.o) { Ctl *b = C(val.u.o); buddy = b ? b->name : QString(); }
            else buddy = qs(vb_v2s(val));
            return vb_v_empty();
        }
        *found = 0;
        return vb_v_empty();
    }
    struct ResizeHook : QObject {
        std::function<void()> f;
        ResizeHook(QObject *p, std::function<void()> fn) : QObject(p), f(std::move(fn)) {}
        bool eventFilter(QObject *, QEvent *e) override { if (e->type() == QEvent::Resize) f(); return false; }
    };
};

/* ---------------------------------------------------------------- DTPicker */
struct OcxDTPicker : Ocx {
    QPointer<QDateTimeEdit> ed;
    QPointer<QCheckBox> chk;
    QPointer<QWidget> wd;
    int format = 1;                             /* dtpLongDate 0, dtpShortDate 1, dtpTime 2, dtpCustom 3 */
    QString custom;
    bool checkbox = false, updown = false;
    static QString qt_format(const QString &w)  /* Windows date picture -> Qt */
    {
        QString r = w;
        r.replace("tt", "AP").replace("t", "A");
        return r;
    }
    QWidget *make(QWidget *parent) override
    {
        wd = new QWidget(parent);
        auto *lay = new QHBoxLayout(wd);
        lay->setContentsMargins(0, 0, 0, 0); lay->setSpacing(2);
        chk = new QCheckBox(wd); chk->hide();
        ed = new QDateTimeEdit(wd);
        ed->setCalendarPopup(true);
        lay->addWidget(chk); lay->addWidget(ed, 1);
        vbobj self = c->self;
        QObject::connect(ed, &QDateTimeEdit::dateTimeChanged, [self](const QDateTime &) { fire_ctl(C(self), "change"); });
        QObject::connect(chk, &QCheckBox::toggled, [self](bool on) {
            Ctl *x = C(self); auto *o = x ? dynamic_cast<OcxDTPicker *>(x->ocx) : nullptr;
            if (o && o->ed) o->ed->setEnabled(on);
            fire_ctl(x, "change");
        });
        return wd;
    }
    void look()
    {
        if (!ed) return;
        QLocale loc;
        QString f = format == 0 ? loc.dateFormat(QLocale::LongFormat) : format == 1 ? loc.dateFormat(QLocale::ShortFormat) :
                    format == 2 ? loc.timeFormat(QLocale::LongFormat) : qt_format(custom);
        ed->setDisplayFormat(f);
        ed->setCalendarPopup(!updown && format != 2);
        chk->setVisible(checkbox);
    }
    void load(const char *const *p) override
    {
        long f = (long)propn(p, "Format", 1);
        format = f > 3 ? (int)(f & 0xFF) : (int)f;
        if (format > 3) format = 1;
        prop(p, "CustomFormat", &custom);
        checkbox = propn(p, "CheckBox", 0) != 0;
        updown = propn(p, "UpDown", 0) != 0;
        double cur = propn(p, "CurrentDate", 0);
        if (ed) {
            QSignalBlocker b(ed);
            ed->setDateTime(cur > 0 ? qdt(cur) : QDateTime::currentDateTime());
            double mn = propn(p, "MinDate", 0), mx = propn(p, "MaxDate", 0);
            if (mn > 0) ed->setMinimumDateTime(qdt(mn));
            if (mx > 0) ed->setMaximumDateTime(qdt(mx));
        }
        look();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        if (!ed) { *found = 0; return vb_v_empty(); }
        if (n == "value") {
            if (!set) {
                if (checkbox && !chk->isChecked()) return vb_v_null();
                QDateTime t = ed->dateTime();
                if (format != 2) t.setTime(t.time());
                return vb_v_date(vbdate(t));
            }
            if (val.vt == VT_NULL) { if (checkbox) chk->setChecked(false); return vb_v_empty(); }
            double d = vb_v2d(val);
            if (vb_errpend) return vb_v_empty();
            QDateTime t = qdt(d);
            if (t < ed->minimumDateTime() || t > ed->maximumDateTime()) { vb_raise(380); return vb_v_empty(); }
            if (checkbox) chk->setChecked(true);
            ed->setDateTime(t);
            return vb_v_empty();
        }
        if (n == "format") { if (!set) return vb_v_l(format); format = vb_v2i(val); look(); return vb_v_empty(); }
        if (n == "customformat") { if (!set) return vstr(custom); custom = qs(vb_v2s(val)); look(); return vb_v_empty(); }
        if (n == "checkbox") { if (!set) return vbool(checkbox); checkbox = vb_v2bool(val) != 0; look(); return vb_v_empty(); }
        if (n == "updown") { if (!set) return vbool(updown); updown = vb_v2bool(val) != 0; look(); return vb_v_empty(); }
        if (n == "mindate" || n == "maxdate") {
            bool mn = n == "mindate";
            if (!set) return vb_v_date(vbdate(mn ? ed->minimumDateTime() : ed->maximumDateTime()));
            QDateTime t = qdt(vb_v2d(val));
            if (mn) ed->setMinimumDateTime(t); else ed->setMaximumDateTime(t);
            return vb_v_empty();
        }
        QDateTime t = ed->dateTime();
        struct { const char *nm; int what; } parts[] = { { "year", 0 }, { "month", 1 }, { "day", 2 }, { "hour", 3 }, { "minute", 4 }, { "second", 5 }, { "dayofweek", 6 } };
        for (auto &e : parts) {
            if (n != e.nm) continue;
            if (!set) {
                int v = e.what == 0 ? t.date().year() : e.what == 1 ? t.date().month() : e.what == 2 ? t.date().day() :
                        e.what == 3 ? t.time().hour() : e.what == 4 ? t.time().minute() : e.what == 5 ? t.time().second() :
                        t.date().dayOfWeek() % 7 + 1;
                return vb_v_l(v);
            }
            int v = vb_v2l(val);
            QDate d = t.date(); QTime tm = t.time();
            if (e.what == 0) d.setDate(v, d.month(), std::min(d.day(), QDate(v, d.month(), 1).daysInMonth()));
            else if (e.what == 1) d.setDate(d.year(), v, std::min(d.day(), QDate(d.year(), v, 1).daysInMonth()));
            else if (e.what == 2) d.setDate(d.year(), d.month(), v);
            else if (e.what == 3) tm.setHMS(v, tm.minute(), tm.second());
            else if (e.what == 4) tm.setHMS(tm.hour(), v, tm.second());
            else if (e.what == 5) tm.setHMS(tm.hour(), tm.minute(), v);
            else d = d.addDays(v - (d.dayOfWeek() % 7 + 1));
            if (!d.isValid() || !tm.isValid()) { vb_raise(380); return vb_v_empty(); }
            ed->setDateTime(QDateTime(d, tm));
            return vb_v_empty();
        }
        *found = 0;
        return vb_v_empty();
    }
};

/* ---------------------------------------------------------------- MonthView */
struct OcxMonthView : Ocx {
    QPointer<QCalendarWidget> cal;
    QWidget *make(QWidget *parent) override
    {
        cal = new QCalendarWidget(parent);
        cal->setGridVisible(false);
        cal->setVerticalHeaderFormat(QCalendarWidget::NoVerticalHeader);
        vbobj self = c->self;
        QObject::connect(cal, &QCalendarWidget::clicked, [self](const QDate &d) {
            vbvar a = vb_v_date(vbdate(QDateTime(d, QTime(0, 0))));
            fire_ctl(C(self), "dateclick", 1, &a);
        });
        QObject::connect(cal, &QCalendarWidget::selectionChanged, [self]() {
            Ctl *x = C(self); auto *o = x ? dynamic_cast<OcxMonthView *>(x->ocx) : nullptr;
            if (!o || !o->cal) return;
            double d = vbdate(QDateTime(o->cal->selectedDate(), QTime(0, 0)));
            vbvar a[3] = { vb_v_date(d), vb_v_date(d), vb_v_bool(0) };
            fire_ctl(x, "selchange", 3, a);
        });
        return cal;
    }
    void load(const char *const *p) override
    {
        if (!cal) return;
        QSignalBlocker b(cal);
        double v = propn(p, "CurrentDate", 0);
        if (v > 0) cal->setSelectedDate(qdt(v).date());
        cal->setVerticalHeaderFormat(propn(p, "ShowWeekNumbers", 0) ? QCalendarWidget::ISOWeekNumbers : QCalendarWidget::NoVerticalHeader);
        int sw = (int)propn(p, "StartOfWeek", 0);
        if (sw >= 1 && sw <= 7) cal->setFirstDayOfWeek((Qt::DayOfWeek)(sw == 1 ? 7 : sw - 1));
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        if (!cal) { *found = 0; return vb_v_empty(); }
        QDate d = cal->selectedDate();
        if (n == "value" || n == "selstart" || n == "selend") {
            if (!set) return vb_v_date(vbdate(QDateTime(d, QTime(0, 0))));
            QDate nd = qdt(vb_v2d(val)).date();
            if (nd < cal->minimumDate() || nd > cal->maximumDate()) { vb_raise(380); return vb_v_empty(); }
            cal->setSelectedDate(nd);
            return vb_v_empty();
        }
        if (n == "mindate" || n == "maxdate") {
            bool mn = n == "mindate";
            if (!set) return vb_v_date(vbdate(QDateTime(mn ? cal->minimumDate() : cal->maximumDate(), QTime(0, 0))));
            QDate nd = qdt(vb_v2d(val)).date();
            if (mn) cal->setMinimumDate(nd); else cal->setMaximumDate(nd);
            return vb_v_empty();
        }
        if (n == "year" || n == "month" || n == "day") {
            if (!set) return vb_v_l(n == "year" ? d.year() : n == "month" ? d.month() : d.day());
            int v = vb_v2l(val);
            QDate nd = n == "year" ? QDate(v, d.month(), std::min(d.day(), QDate(v, d.month(), 1).daysInMonth())) :
                       n == "month" ? QDate(d.year(), v, std::min(d.day(), QDate(d.year(), v, 1).daysInMonth())) : QDate(d.year(), d.month(), v);
            if (!nd.isValid()) { vb_raise(380); return vb_v_empty(); }
            cal->setSelectedDate(nd);
            return vb_v_empty();
        }
        if (n == "dayofweek") return vb_v_l(d.dayOfWeek() % 7 + 1);
        if (n == "showweeknumbers") {
            if (!set) return vbool(cal->verticalHeaderFormat() != QCalendarWidget::NoVerticalHeader);
            cal->setVerticalHeaderFormat(vb_v2bool(val) ? QCalendarWidget::ISOWeekNumbers : QCalendarWidget::NoVerticalHeader);
            return vb_v_empty();
        }
        if (n == "multiselect" || n == "showtoday" || n == "maxselcount" || n == "monthrows" || n == "monthcolumns")
            return set ? vb_v_empty() : n == "showtoday" ? vbool(true) : n == "multiselect" ? vbool(false) : vb_v_l(n == "maxselcount" ? 7 : 1);
        *found = 0;
        return vb_v_empty();
    }
};

/* ---------------------------------------------------------------- FlatScrollBar */
struct OcxFlatScroll : Ocx {
    QPointer<QScrollBar> sb;
    QWidget *make(QWidget *parent) override
    {
        sb = new QScrollBar(Qt::Horizontal, parent);
        vbobj self = c->self;
        QObject::connect(sb, &QScrollBar::sliderMoved, [self](int) { fire_ctl(C(self), "scroll"); });
        QObject::connect(sb, &QScrollBar::valueChanged, [self](int) {
            Ctl *x = C(self); auto *o = x ? dynamic_cast<OcxFlatScroll *>(x->ocx) : nullptr;
            if (o && o->sb && !o->sb->isSliderDown()) fire_ctl(x, "change");
        });
        QObject::connect(sb, &QScrollBar::sliderReleased, [self]() { fire_ctl(C(self), "change"); });
        return sb;
    }
    void load(const char *const *p) override
    {
        QSignalBlocker b(sb);
        sb->setOrientation(propn(p, "Orientation", 1) == 0 ? Qt::Vertical : Qt::Horizontal);
        sb->setRange((int)propn(p, "Min", 0), (int)propn(p, "Max", 32767));
        sb->setValue((int)propn(p, "Value", 0));
        sb->setSingleStep((int)propn(p, "SmallChange", 1)); sb->setPageStep((int)propn(p, "LargeChange", 1));
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind; (void)argc; (void)argv;
        int v = set ? vb_v2l(val) : 0;
        if (n == "value") {
            if (!set) return vb_v_l(sb->value());
            if (v < std::min(sb->minimum(), sb->maximum()) || v > std::max(sb->minimum(), sb->maximum())) { vb_raise(380); return vb_v_empty(); }
            sb->setValue(v); return vb_v_empty();
        }
        if (n == "min") { if (!set) return vb_v_l(sb->minimum()); sb->setMinimum(v); return vb_v_empty(); }
        if (n == "max") { if (!set) return vb_v_l(sb->maximum()); sb->setMaximum(v); return vb_v_empty(); }
        if (n == "smallchange") { if (!set) return vb_v_l(sb->singleStep()); sb->setSingleStep(v); return vb_v_empty(); }
        if (n == "largechange") { if (!set) return vb_v_l(sb->pageStep()); sb->setPageStep(v); return vb_v_empty(); }
        if (n == "orientation") { if (!set) return vb_v_l(sb->orientation() == Qt::Horizontal ? 1 : 0); sb->setOrientation(v ? Qt::Horizontal : Qt::Vertical); return vb_v_empty(); }
        *found = 0;
        return vb_v_empty();
    }
};

/* ---------------------------------------------------------------- SSTab
 * A container: the controls of the tabs that are not shown stay 75000 twips to the left, as VB6
 * keeps them; a click on a tab moves them back (Tab(i).Control(j) of the .frm names them). */
class SSTabWidget : public QWidget {
public:
    QTabBar *bar;
    QFrame *page;
    int tabh = 0;
    explicit SSTabWidget(QWidget *parent) : QWidget(parent)
    {
        bar = new QTabBar(this);
        bar->setDrawBase(false);
        bar->setExpanding(true);
        bar->setUsesScrollButtons(false);
        page = new QFrame(this);
        page->setFrameStyle(QFrame::StyledPanel | QFrame::Raised);
        page->lower();
    }
    void place()
    {
        int h = tabh > 0 ? tabh : bar->sizeHint().height();
        bar->setGeometry(0, 0, width(), h);
        page->setGeometry(0, h - 1, width(), height() - h + 1);
        page->lower();
    }
protected:
    void resizeEvent(QResizeEvent *e) override { QWidget::resizeEvent(e); place(); }
};
struct OcxSSTab : Ocx {
    QPointer<SSTabWidget> wd;
    std::vector<QStringList> members;           /* Tab(i).Control(j) */
    std::vector<QString> captions;
    std::vector<bool> enabled, visible;
    int cur = 0, perrow = 3, style = 0;
    bool switching = false;
    QWidget *make(QWidget *parent) override
    {
        wd = new SSTabWidget(parent);
        vbobj self = c->self;
        QObject::connect(wd->bar, &QTabBar::currentChanged, [self](int i) {
            Ctl *x = C(self); auto *o = x ? dynamic_cast<OcxSSTab *>(x->ocx) : nullptr;
            if (o && !o->switching) o->go(i, true);
        });
        return wd;
    }
    Ctl *ctl_named(const QString &nm)
    {
        if (!c->form) return nullptr;
        QString base = nm; int ix = -1;
        int p = nm.indexOf('(');
        if (p > 0 && nm.endsWith(')')) { base = nm.left(p); ix = nm.mid(p + 1, nm.size() - p - 2).toInt(); }
        auto it = c->form->byname.find(lower(base.toUtf8().constData()));
        if (it == c->form->byname.end()) return nullptr;
        vbobj o = it->second;
        if (o->cls == &CLS_ARRAY) { auto *a = (CtlArray *)o->ext; auto e = a->elems.find(ix); o = e == a->elems.end() ? nullptr : e->second; }
        return o ? C(o) : nullptr;
    }
    void shift(int tab, double dx)
    {
        if (tab < 0 || tab >= (int)members.size()) return;
        for (const QString &nm : members[tab]) {
            Ctl *k = ctl_named(nm);
            if (!k) continue;
            k->left += dx;
            ctl_geometry(k);
        }
    }
    void go(int i, bool event)
    {
        if (i < 0 || i >= (int)captions.size() || i == cur) return;
        int prev = cur;
        shift(prev, -75000);
        shift(i, 75000);
        cur = i;
        switching = true;
        if (wd && wd->bar->currentIndex() != i) wd->bar->setCurrentIndex(i);
        switching = false;
        if (event) { vbvar a = vb_v_i((vb_i16)prev); fire_ctl(c, "click", 1, &a); }
    }
    void sync()
    {
        if (!wd) return;
        switching = true;
        QTabBar *b = wd->bar;
        while (b->count() > (int)captions.size()) b->removeTab(b->count() - 1);
        while (b->count() < (int)captions.size()) b->addTab(QString());
        for (int i = 0; i < (int)captions.size(); i++) {
            QString t = captions[i]; t.replace("&&", "\x01").remove('&').replace("\x01", "&");
            b->setTabText(i, t);
            b->setTabEnabled(i, enabled[i]);
            b->setTabVisible(i, visible[i]);
        }
        b->setCurrentIndex(cur);
        switching = false;
        wd->place();
    }
    void resize_tabs(int n)
    {
        members.resize(n); captions.resize(n); enabled.resize(n, true); visible.resize(n, true);
        for (int i = 0; i < n; i++) if (captions[i].isEmpty()) captions[i] = QString("Tab %1").arg(i);
    }
    void load(const char *const *p) override
    {
        int n = (int)propn(p, "Tabs", 3);
        resize_tabs(n);
        perrow = (int)propn(p, "TabsPerRow", 3);
        style = (int)propn(p, "Style", 0);
        cur = (int)propn(p, "Tab", 0);
        for (int i = 0; i < n; i++) {
            QString v;
            if (prop(p, QString("TabCaption(%1)").arg(i).toUtf8().constData(), &v)) captions[i] = v;
            int cc = (int)propn(p, QString("Tab(%1).ControlCount").arg(i).toUtf8().constData(), 0);
            for (int j = 0; j < cc; j++)
                if (prop(p, QString("Tab(%1).Control(%2)").arg(i).arg(j).toUtf8().constData(), &v)) members[i] << v;
            enabled[i] = propn(p, QString("TabEnabled(%1)").arg(i).toUtf8().constData(), -1) != 0;
            visible[i] = propn(p, QString("TabVisible(%1)").arg(i).toUtf8().constData(), -1) != 0;
        }
        double th = propn(p, "TabHeight", 0);
        if (wd && th > 0) wd->tabh = (int)std::lround(th / TPP);
        sync();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        int ix = (int)arg_l(argc - (set ? 1 : 0), argv, 0, cur);
        bool indexed = n == "tabcaption" || n == "tabenabled" || n == "tabvisible" || n == "tabpicture";
        if (indexed && (ix < 0 || ix >= (int)captions.size())) { vb_raise(380); return vb_v_empty(); }
        if (n == "tab") {
            if (!set) return vb_v_i((vb_i16)cur);
            int t = vb_v2i(val);
            if (t < 0 || t >= (int)captions.size()) { vb_raise(380); return vb_v_empty(); }
            go(t, true);
            return vb_v_empty();
        }
        if (n == "tabs") {
            if (!set) return vb_v_i((vb_i16)captions.size());
            int t = vb_v2i(val);
            if (t < 1) { vb_raise(380); return vb_v_empty(); }
            if (cur >= t) go(t - 1, false);
            resize_tabs(t); sync();
            return vb_v_empty();
        }
        if (n == "tabcaption" || n == "caption") {
            if (n == "caption") ix = cur;
            if (!set) return vstr(captions[ix]);
            captions[ix] = qs(vb_v2s(val)); sync();
            return vb_v_empty();
        }
        if (n == "tabenabled" || n == "tabvisible") {
            std::vector<bool> &v = n == "tabenabled" ? enabled : visible;
            if (!set) return vbool(v[ix]);
            v[ix] = vb_v2bool(val) != 0; sync();
            return vb_v_empty();
        }
        if (n == "tabsperrow") { if (!set) return vb_v_i((vb_i16)perrow); perrow = vb_v2i(val); return vb_v_empty(); }
        if (n == "style") { if (!set) return vb_v_l(style); style = vb_v2i(val); return vb_v_empty(); }
        if (n == "tabheight") {
            if (!set) return vb_v_f((float)from_twips(c, (wd ? (wd->tabh > 0 ? wd->tabh : wd->bar->sizeHint().height()) : 0) * TPP, true, false));
            if (wd) { wd->tabh = (int)std::lround(to_twips(c, vb_v2d(val), true, false) / TPP); wd->place(); }
            return vb_v_empty();
        }
        if (n == "tabpicture") return set ? vb_v_empty() : retobj(new_pic(QImage(), 0));
        if (n == "showfocusrect" || n == "wordwrap" || n == "tabmaxwidth" || n == "taborientation") return set ? vb_v_empty() : vb_v_l(0);
        *found = 0;
        return vb_v_empty();
    }
};

/* ---------------------------------------------------------------- MSFlexGrid */
static void flex_error(vb_i32 n)
{
    const char *t = n == 30009 ? "Invalid Row value" : n == 30010 ? "Invalid Col value" : n == 30015 ? "Cannot remove last non-fixed row" : "Invalid value";
    vb_i32 m = vb_mark();
    vb_raise_full(n, vs(QString("MSFlexGrid")), vs(QString(t)));
    vb_release(m);
}
class FlexTable : public QTableWidget {
public:
    std::function<void()> on_click, on_dblclick;
    std::function<void(QMouseEvent *)> on_move;
    explicit FlexTable(QWidget *parent) : QTableWidget(parent) { setMouseTracking(true); viewport()->setMouseTracking(true); }
protected:
    void mouseMoveEvent(QMouseEvent *e) override { if (on_move) on_move(e); QTableWidget::mouseMoveEvent(e); }
    void mouseReleaseEvent(QMouseEvent *e) override { QTableWidget::mouseReleaseEvent(e); if (e->button() == Qt::LeftButton && on_click) on_click(); }
    void mouseDoubleClickEvent(QMouseEvent *e) override { QTableWidget::mouseDoubleClickEvent(e); if (on_dblclick) on_dblclick(); }
};
struct OcxFlexGrid : Ocx {
    QPointer<FlexTable> tw;
    int rows = 2, cols = 2, frows = 1, fcols = 1;
    int row = 1, col = 1, rowsel = 1, colsel = 1;
    int mouserow = 0, mousecol = 0, selmode = 0, highlight = 1, fillstyle = 0, mergecells = 0, gridlinewidth = 1;
    double rowheightmin = 0;                    /* RowHeightMin (twips) */
    long back = 0x80000005L, fore = 0x80000008L, backfixed = 0x8000000FL, forefixed = 0x80000012L;
    long backsel = 0x8000000DL, foresel = 0x8000000EL, backbkg = 0x8000000CL, gridcolor = 0xC0C0C0;
    bool redraw = true, wordwrap = false, inside = false;
    std::vector<int> colalign;                  /* -1: general */
    std::vector<int> fixedalign;
    std::vector<long> rowdata, coldata;
    std::vector<bool> mergerow, mergecol;
    QWidget *make(QWidget *parent) override
    {
        tw = new FlexTable(parent);
        tw->horizontalHeader()->hide();
        tw->verticalHeader()->hide();
        tw->setEditTriggers(QAbstractItemView::NoEditTriggers);
        tw->setSelectionMode(QAbstractItemView::NoSelection);
        tw->verticalHeader()->setDefaultSectionSize(16);
        tw->horizontalHeader()->setDefaultSectionSize(64);
        tw->horizontalHeader()->setMinimumSectionSize(1);
        tw->verticalHeader()->setMinimumSectionSize(1);
        vbobj self = c->self;
        auto me = [self]() -> OcxFlexGrid * { Ctl *x = C(self); return x ? dynamic_cast<OcxFlexGrid *>(x->ocx) : nullptr; };
        QObject::connect(tw, &QTableWidget::currentCellChanged, [me](int r, int cc, int, int) {
            OcxFlexGrid *g = me();
            if (!g || g->inside || r < 0 || cc < 0) return;
            if (r < g->frows || cc < g->fcols) return;
            bool shift = QApplication::keyboardModifiers() & Qt::ShiftModifier;
            fire_ctl(g->c, "leavecell");
            if (shift) { g->rowsel = r; g->colsel = cc; }
            else { g->row = r; g->col = cc; g->rowsel = r; g->colsel = cc; }
            if (g->selmode == 1) { g->colsel = g->cols - 1; g->col = g->fcols; }
            if (g->selmode == 2) { g->rowsel = g->rows - 1; g->row = g->frows; }
            g->paint_sel();
            fire_ctl(g->c, "rowcolchange");
            fire_ctl(g->c, "entercell");
            fire_ctl(g->c, "selchange");
        });
        tw->on_click = [self]() { fire_ctl(C(self), "click"); };
        tw->on_dblclick = [self]() { fire_ctl(C(self), "dblclick"); };
        tw->on_move = [me](QMouseEvent *e) {
            OcxFlexGrid *g = me();
            if (!g) return;
            QPoint p = ev_pos(e).toPoint();
            int r = g->tw->rowAt(p.y()), cc = g->tw->columnAt(p.x());
            if (r >= 0) g->mouserow = r;
            if (cc >= 0) g->mousecol = cc;
        };
        QObject::connect(tw->horizontalScrollBar(), &QScrollBar::valueChanged, [self](int) { fire_ctl(C(self), "scroll"); });
        QObject::connect(tw->verticalScrollBar(), &QScrollBar::valueChanged, [self](int) { fire_ctl(C(self), "scroll"); });
        return tw;
    }
    QTableWidgetItem *cell(int r, int cc)
    {
        QTableWidgetItem *it = tw->item(r, cc);
        if (!it) { it = new QTableWidgetItem; it->setFlags(Qt::ItemIsEnabled); tw->setItem(r, cc, it); style(r, cc); }
        return it;
    }
    bool fixed(int r, int cc) const { return r < frows || cc < fcols; }
    void style(int r, int cc)
    {
        QTableWidgetItem *it = tw->item(r, cc);
        if (!it) return;
        bool fx = fixed(r, cc);
        bool sel = !fx && highlight && in_sel(r, cc) && !(r == row && cc == col && selmode == 0 && r == rowsel && cc == colsel);
        if (!it->data(Qt::UserRole + 1).isValid()) it->setBackground(qcolor(fx ? backfixed : sel ? backsel : back));
        else it->setBackground(sel ? QBrush(qcolor(backsel)) : QBrush(qcolor(it->data(Qt::UserRole + 1).toLongLong())));
        if (!it->data(Qt::UserRole + 2).isValid()) it->setForeground(qcolor(fx ? forefixed : sel ? foresel : fore));
        else it->setForeground(sel ? QBrush(qcolor(foresel)) : QBrush(qcolor(it->data(Qt::UserRole + 2).toLongLong())));
        int al = it->data(Qt::UserRole + 3).isValid() ? it->data(Qt::UserRole + 3).toInt() :
                 fx && cc < (int)fixedalign.size() && fixedalign[cc] >= 0 ? fixedalign[cc] :
                 cc < (int)colalign.size() ? colalign[cc] : 9;
        it->setData(Qt::TextAlignmentRole, (int)qalign(al, it->text()));
    }
    static Qt::Alignment qalign(int al, const QString &text)
    {
        if (al == 9) {                          /* flexAlignGeneral: numbers right, text left */
            bool num = false; text.trimmed().toDouble(&num);
            return (num ? Qt::AlignRight : Qt::AlignLeft) | Qt::AlignVCenter;
        }
        Qt::Alignment h = al <= 2 ? Qt::AlignLeft : al <= 5 ? Qt::AlignHCenter : Qt::AlignRight;
        Qt::Alignment v = al % 3 == 0 ? Qt::AlignTop : al % 3 == 1 ? Qt::AlignVCenter : Qt::AlignBottom;
        return h | v;
    }
    bool in_sel(int r, int cc) const
    {
        return r >= std::min(row, rowsel) && r <= std::max(row, rowsel) && cc >= std::min(col, colsel) && cc <= std::max(col, colsel);
    }
    void paint_sel()
    {
        if (!tw || !redraw) return;
        for (int r = 0; r < rows; r++) for (int cc = 0; cc < cols; cc++) style(r, cc);
    }
    void shape()
    {
        if (!tw) return;
        inside = true;
        tw->setRowCount(rows); tw->setColumnCount(cols);
        colalign.resize(cols, 9); fixedalign.resize(cols, -1);
        rowdata.resize(rows, 0); coldata.resize(cols, 0);
        mergerow.resize(rows, false); mergecol.resize(cols, false);
        for (int r = 0; r < rows; r++) for (int cc = 0; cc < cols; cc++) cell(r, cc);
        if (row >= rows) row = rows - 1;
        if (col >= cols) col = cols - 1;
        if (rowsel >= rows) rowsel = rows - 1;
        if (colsel >= cols) colsel = cols - 1;
        tw->setStyleSheet(QString("QTableWidget { gridline-color: %1; background: %2; }").arg(qcolor(gridcolor).name(), qcolor(backbkg).name()));
        paint_sel();
        inside = false;
    }
    void load(const char *const *p) override
    {
        rows = std::max(0, (int)propn(p, "Rows", 2)); cols = std::max(0, (int)propn(p, "Cols", 2));
        frows = (int)propn(p, "FixedRows", 1); fcols = (int)propn(p, "FixedCols", 1);
        row = rowsel = std::min(frows, rows - 1); col = colsel = std::min(fcols, cols - 1);
        back = propc(p, "BackColor", back); fore = propc(p, "ForeColor", fore);
        backfixed = propc(p, "BackColorFixed", backfixed); forefixed = propc(p, "ForeColorFixed", forefixed);
        backsel = propc(p, "BackColorSel", backsel); foresel = propc(p, "ForeColorSel", foresel);
        backbkg = propc(p, "BackColorBkg", backbkg); gridcolor = propc(p, "GridColor", gridcolor);
        highlight = (int)propn(p, "HighLight", 1); selmode = (int)propn(p, "SelectionMode", 0);
        fillstyle = (int)propn(p, "FillStyle", 0); mergecells = (int)propn(p, "MergeCells", 0);
        wordwrap = propn(p, "WordWrap", 0) != 0;
        if (tw) {
            tw->setShowGrid(propn(p, "GridLines", 1) != 0);
            tw->setWordWrap(wordwrap);
            int sb = (int)propn(p, "ScrollBars", 3);
            tw->setHorizontalScrollBarPolicy(sb & 1 ? Qt::ScrollBarAsNeeded : Qt::ScrollBarAlwaysOff);
            tw->setVerticalScrollBarPolicy(sb & 2 ? Qt::ScrollBarAsNeeded : Qt::ScrollBarAlwaysOff);
            int ar = (int)propn(p, "AllowUserResizing", 0);
            tw->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
            (void)ar;
        }
        shape();
        QString fs;
        if (prop(p, "FormatString", &fs)) format_string(fs);
        QString b;
        if (prop(p, "Bindings", &b)) bindings = QByteArray::fromHex(b.toLatin1());
    }
    /* FormatString: "<Name|>Age|^City;Row1|Row2": the captions and alignments of the fixed row (and column) */
    void format_string(const QString &fs)
    {
        QString colpart = fs.section(';', 0, 0), rowpart = fs.contains(';') ? fs.section(';', 1) : QString();
        QStringList cs = colpart.split('|');
        if (!colpart.isEmpty()) {
            if ((int)cs.size() > cols) { cols = cs.size(); shape(); }
            for (int i = 0; i < cs.size(); i++) {
                QString t = cs[i];
                int al = -1;
                if (t.startsWith('<')) { al = 1; t.remove(0, 1); }
                else if (t.startsWith('^')) { al = 4; t.remove(0, 1); }
                else if (t.startsWith('>')) { al = 7; t.remove(0, 1); }
                if (al >= 0) colalign[i] = al;
                if (frows > 0) cell(0, i)->setText(t);
                QFontMetrics fm(tw->font());
                tw->setColumnWidth(i, std::max(tw->columnWidth(i), fm.horizontalAdvance(t) + 8));
            }
        }
        if (!rowpart.isEmpty()) {
            QStringList rs = rowpart.split('|');
            if ((int)rs.size() > rows) { rows = rs.size(); shape(); }
            for (int i = 0; i < rs.size(); i++) if (fcols > 0 && !rs[i].isEmpty()) cell(i, 0)->setText(rs[i]);
        }
        paint_sel();
    }
    bool rc_ok(int r, int cc) const { return r >= 0 && r < rows && cc >= 0 && cc < cols; }
    void each_sel(std::function<void(QTableWidgetItem *, int, int)> f)
    {
        int r0 = row, r1 = row, c0 = col, c1 = col;
        if (fillstyle == 1) { r0 = std::min(row, rowsel); r1 = std::max(row, rowsel); c0 = std::min(col, colsel); c1 = std::max(col, colsel); }
        for (int r = r0; r <= r1; r++) for (int cc = c0; cc <= c1; cc++) if (rc_ok(r, cc)) f(cell(r, cc), r, cc);
    }
    QString cell_text(int r, int cc) { return rc_ok(r, cc) ? cell(r, cc)->text() : QString(); }
    void sort(int how)
    {
        int c0 = std::min(col, colsel), c1 = std::max(col, colsel);
        std::vector<std::vector<QTableWidgetItem *>> data;
        for (int r = frows; r < rows; r++) {
            std::vector<QTableWidgetItem *> line;
            for (int cc = 0; cc < cols; cc++) line.push_back(tw->takeItem(r, cc));
            data.push_back(line);
        }
        auto key = [&](const std::vector<QTableWidgetItem *> &l, int cc) { return l[cc] ? l[cc]->text() : QString(); };
        std::stable_sort(data.begin(), data.end(), [&](const std::vector<QTableWidgetItem *> &a, const std::vector<QTableWidgetItem *> &b) {
            for (int cc = c0; cc <= c1; cc++) {
                QString x = key(a, cc), y = key(b, cc);
                int r = 0;
                bool numeric = how == 3 || how == 4, generic = how == 1 || how == 2;
                bool nx = false, ny = false;
                double dx = x.toDouble(&nx), dy = y.toDouble(&ny);
                if (numeric || (generic && nx && ny)) r = numeric && !nx ? (ny ? -1 : 0) : (dx < dy ? -1 : dx > dy ? 1 : 0);
                else r = QString::compare(x, y, how == 7 || how == 8 ? Qt::CaseSensitive : Qt::CaseInsensitive);
                if (r) return (how % 2 == 0) ? r > 0 : r < 0;
            }
            return false;
        });
        for (int i = 0; i < (int)data.size(); i++)
            for (int cc = 0; cc < cols; cc++) tw->setItem(frows + i, cc, data[i][cc]);
        paint_sel();
    }
    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override;
    /* bound to a data source (ocx_data.inc): the Bindings of the .frm, or Set DataSource */
    QByteArray bindings;
    vbobj dsrc = nullptr;
    ~OcxFlexGrid() override;
    bool data_aware() override { return true; }
    bool own_binding() override { return true; }
    void source_reset(vbobj s) override;
    vbvar datasource_member(bool set, vbvar val);
    void bind_fill();
};
vbvar OcxFlexGrid::member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    if (!tw) { *found = 0; return vb_v_empty(); }
    if (n == "datasource") return datasource_member(set, val);
    if (n == "datamember" || n == "datafield") return set ? vb_v_empty() : vb_v_s(vb_sfromutf8(""));
    int na = argc - (set ? 1 : 0);
    auto ival = [&]() { return vb_v2l(val); };
    auto inrange = [&](int v, int lo, int hi) { if (v < lo || v > hi) { vb_raise(381); return false; } return true; };   /* Subscript out of range */
    if (n == "rows" || n == "cols") {
        int *p = n == "rows" ? &rows : &cols;
        if (!set) return vb_v_l(*p);
        int v = ival();
        if (v < 0 || v < (n == "rows" ? frows : fcols)) { flex_error(30009); return vb_v_empty(); }
        *p = v; shape();
        return vb_v_empty();
    }
    if (n == "fixedrows" || n == "fixedcols") {
        int *p = n == "fixedrows" ? &frows : &fcols;
        if (!set) return vb_v_l(*p);
        int v = ival();
        if (v < 0 || v >= (n == "fixedrows" ? rows : cols) + (v == 0 ? 1 : 0)) { flex_error(30009); return vb_v_empty(); }
        *p = v;
        if (row < frows) row = rowsel = std::min(frows, rows - 1);
        if (col < fcols) col = colsel = std::min(fcols, cols - 1);
        paint_sel();
        return vb_v_empty();
    }
    if (n == "row" || n == "col") {
        bool r = n == "row";
        if (!set) return vb_v_l(r ? row : col);
        int v = ival();
        if (v < 0 || v >= (r ? rows : cols)) { flex_error(r ? 30009 : 30010); return vb_v_empty(); }
        if (r) row = rowsel = v; else col = colsel = v;
        inside = true; tw->setCurrentCell(row, col); inside = false;
        paint_sel();
        fire_ctl(c, "rowcolchange");
        return vb_v_empty();
    }
    if (n == "rowsel" || n == "colsel") {
        bool r = n == "rowsel";
        if (!set) return vb_v_l(r ? rowsel : colsel);
        int v = ival();
        if (!inrange(v, 0, (r ? rows : cols) - 1)) return vb_v_empty();
        if (r) rowsel = v; else colsel = v;
        paint_sel();
        fire_ctl(c, "selchange");
        return vb_v_empty();
    }
    if (n == "text") {
        if (!set) return vstr(cell_text(row, col));
        QString t = qs(vb_v2s(val));
        each_sel([&](QTableWidgetItem *it, int r, int cc) { it->setText(t); style(r, cc); });
        return vb_v_empty();
    }
    if (n == "textmatrix") {
        int r = (int)arg_l(na, argv, 0, 0), cc = (int)arg_l(na, argv, 1, 0);
        if (!rc_ok(r, cc)) { vb_raise(381); return vb_v_empty(); }
        if (!set) return vstr(cell_text(r, cc));
        cell(r, cc)->setText(qs(vb_v2s(val))); style(r, cc);
        return vb_v_empty();
    }
    if (n == "textarray") {
        int i = (int)arg_l(na, argv, 0, 0), r = cols ? i / cols : 0, cc = cols ? i % cols : 0;
        if (!rc_ok(r, cc)) { vb_raise(381); return vb_v_empty(); }
        if (!set) return vstr(cell_text(r, cc));
        cell(r, cc)->setText(qs(vb_v2s(val))); style(r, cc);
        return vb_v_empty();
    }
    if (n == "rowheightmin") {
        if (!set) return vb_v_l((vb_i32)rowheightmin);
        rowheightmin = std::max(0.0, vb_v2d(val));
        if (tw && rowheightmin > 0) tw->verticalHeader()->setMinimumSectionSize((int)std::lround(rowheightmin / TPP));
        return vb_v_empty();
    }
    if (n == "colwidth" || n == "rowheight") {
        bool w = n == "colwidth";
        int i = (int)arg_l(na, argv, 0, 0);
        if (!inrange(i, -1, (w ? cols : rows) - 1)) return vb_v_empty();
        if (!set) return vb_v_l((vb_i32)std::lround((w ? tw->columnWidth(std::max(0, i)) : tw->rowHeight(std::max(0, i))) * TPP));
        int px = ival() < 0 ? (w ? 64 : 16) : (int)std::lround(ival() / TPP);
        for (int k = (i < 0 ? 0 : i); k <= (i < 0 ? (w ? cols : rows) - 1 : i); k++) { if (w) tw->setColumnWidth(k, px); else tw->setRowHeight(k, px); }
        return vb_v_empty();
    }
    if (n == "colalignment" || n == "fixedalignment") {
        std::vector<int> &v = n == "colalignment" ? colalign : fixedalign;
        int i = (int)arg_l(na, argv, 0, 0);
        if (!inrange(i, -1, cols - 1)) return vb_v_empty();
        if (!set) return vb_v_i((vb_i16)(i >= 0 ? (v[i] < 0 ? colalign[i] : v[i]) : 0));
        for (int k = (i < 0 ? 0 : i); k <= (i < 0 ? cols - 1 : i); k++) v[k] = ival();
        paint_sel();
        return vb_v_empty();
    }
    if (n == "cellalignment") {
        if (!set) { auto d = cell(row, col)->data(Qt::UserRole + 3); return vb_v_i((vb_i16)(d.isValid() ? d.toInt() : colalign[col])); }
        int a = ival();
        each_sel([&](QTableWidgetItem *it, int r, int cc) { it->setData(Qt::UserRole + 3, a); style(r, cc); });
        return vb_v_empty();
    }
    if (n == "cellbackcolor" || n == "cellforecolor") {
        int role = n == "cellbackcolor" ? Qt::UserRole + 1 : Qt::UserRole + 2;
        if (!set) { auto d = cell(row, col)->data(role); return vb_v_l(d.isValid() ? (vb_i32)d.toLongLong() : 0); }
        long v = ival();
        each_sel([&](QTableWidgetItem *it, int r, int cc) { if (v) it->setData(role, (qlonglong)v); else it->setData(role, QVariant()); style(r, cc); });
        return vb_v_empty();
    }
    if (n == "cellfontbold" || n == "cellfontitalic" || n == "cellfontunderline" || n == "cellfontstrikethrough" || n == "cellfontsize" || n == "cellfontname") {
        QFont f = cell(row, col)->font();
        if (!set) {
            if (n == "cellfontbold") return vbool(f.bold());
            if (n == "cellfontitalic") return vbool(f.italic());
            if (n == "cellfontunderline") return vbool(f.underline());
            if (n == "cellfontstrikethrough") return vbool(f.strikeOut());
            if (n == "cellfontsize") return vb_v_f((float)f.pointSizeF());
            return vstr(f.family());
        }
        each_sel([&](QTableWidgetItem *it, int, int) {
            QFont g = it->font();
            if (n == "cellfontbold") g.setBold(vb_v2bool(val) != 0);
            else if (n == "cellfontitalic") g.setItalic(vb_v2bool(val) != 0);
            else if (n == "cellfontunderline") g.setUnderline(vb_v2bool(val) != 0);
            else if (n == "cellfontstrikethrough") g.setStrikeOut(vb_v2bool(val) != 0);
            else if (n == "cellfontsize") g.setPointSizeF(vb_v2d(val));
            else g.setFamily(qs(vb_v2s(val)));
            it->setFont(g);
        });
        return vb_v_empty();
    }
    if (n == "additem") {                       /* AddItem "a" & vbTab & "b", [Index] */
        QString t = arg_s(argc, argv, 0);
        int at = given(argc, argv, 1) ? (int)arg_l(argc, argv, 1, rows) : rows;
        if (at < 0 || at > rows) { vb_raise(381); return vb_v_empty(); }
        rows++;
        inside = true;
        tw->insertRow(at);
        rowdata.insert(rowdata.begin() + at, 0); mergerow.insert(mergerow.begin() + at, false);
        QStringList parts = t.split('\t');
        for (int cc = 0; cc < cols; cc++) cell(at, cc)->setText(cc < parts.size() ? parts[cc] : QString());
        inside = false;
        paint_sel();
        return vb_v_empty();
    }
    if (n == "removeitem") {
        int at = (int)arg_l(argc, argv, 0, 0);
        if (at < 0 || at >= rows) { vb_raise(381); return vb_v_empty(); }
        if (rows - 1 < frows || (rows == frows + 1 && frows > 0 && at >= frows)) { flex_error(30015); return vb_v_empty(); }
        inside = true;
        tw->removeRow(at);
        rowdata.erase(rowdata.begin() + at); mergerow.erase(mergerow.begin() + at);
        rows--;
        if (row >= rows) row = rows - 1;
        if (rowsel >= rows) rowsel = rows - 1;
        inside = false;
        paint_sel();
        return vb_v_empty();
    }
    if (n == "clear") {
        for (int r = 0; r < rows; r++) for (int cc = 0; cc < cols; cc++) { cell(r, cc)->setText(QString()); }
        return vb_v_empty();
    }
    if (n == "clearstructure") { colalign.assign(cols, 9); fixedalign.assign(cols, -1); paint_sel(); return vb_v_empty(); }
    if (n == "clip") {
        int r0 = std::min(row, rowsel), r1 = std::max(row, rowsel), c0 = std::min(col, colsel), c1 = std::max(col, colsel);
        if (!set) {
            QStringList lines;
            for (int r = r0; r <= r1; r++) { QStringList l; for (int cc = c0; cc <= c1; cc++) l << cell_text(r, cc); lines << l.join('\t'); }
            return vstr(lines.join('\r'));
        }
        QStringList lines = qs(vb_v2s(val)).split('\r');
        for (int r = r0, i = 0; r <= r1 && i < lines.size(); r++, i++) {
            QStringList l = lines[i].split('\t');
            for (int cc = c0, j = 0; cc <= c1 && j < l.size(); cc++, j++) { cell(r, cc)->setText(l[j]); style(r, cc); }
        }
        return vb_v_empty();
    }
    if (n == "sort") { if (set) sort(ival()); return vb_v_empty(); }
    if (n == "formatstring") { if (!set) return vstr(QString()); format_string(qs(vb_v2s(val))); return vb_v_empty(); }
    if (n == "toprow" || n == "leftcol") {
        bool r = n == "toprow";
        if (!set) { int v = r ? tw->rowAt(0) : tw->columnAt(0); return vb_v_l(std::max(r ? frows : fcols, v)); }
        int v = ival();
        if (!inrange(v, 0, (r ? rows : cols) - 1)) return vb_v_empty();
        if (r) tw->scrollToItem(cell(v, std::max(0, tw->columnAt(0))), QAbstractItemView::PositionAtTop);
        else tw->horizontalScrollBar()->setValue(tw->columnViewportPosition(v) + tw->horizontalScrollBar()->value());
        return vb_v_empty();
    }
    if (n == "mouserow") return vb_v_l(mouserow);
    if (n == "mousecol") return vb_v_l(mousecol);
    if (n == "rowpos" || n == "colpos") {
        int i = (int)arg_l(na, argv, 0, 0);
        int px = n == "rowpos" ? tw->rowViewportPosition(i) : tw->columnViewportPosition(i);
        return vb_v_l((vb_i32)std::lround(px * TPP));
    }
    if (n == "cellleft" || n == "celltop" || n == "cellwidth" || n == "cellheight") {
        QRect r = tw->visualItemRect(cell(row, col));
        int v = n == "cellleft" ? r.x() : n == "celltop" ? r.y() : n == "cellwidth" ? r.width() : r.height();
        return vb_v_l((vb_i32)std::lround(v * TPP));
    }
    if (n == "rowisvisible" || n == "colisvisible") {
        int i = (int)arg_l(na, argv, 0, 0);
        QRect vp = tw->viewport()->rect();
        int pos = n == "rowisvisible" ? tw->rowViewportPosition(i) : tw->columnViewportPosition(i);
        return vbool(pos >= 0 && pos < (n == "rowisvisible" ? vp.height() : vp.width()));
    }
    if (n == "rowdata" || n == "coldata") {
        std::vector<long> &v = n == "rowdata" ? rowdata : coldata;
        int i = (int)arg_l(na, argv, 0, 0);
        if (!inrange(i, 0, (int)v.size() - 1)) return vb_v_empty();
        if (!set) return vb_v_l((vb_i32)v[i]);
        v[i] = ival();
        return vb_v_empty();
    }
    if (n == "mergerow" || n == "mergecol") {
        std::vector<bool> &v = n == "mergerow" ? mergerow : mergecol;
        int i = (int)arg_l(na, argv, 0, 0);
        if (!inrange(i, 0, (int)v.size() - 1)) return vb_v_empty();
        if (!set) return vbool(v[i]);
        v[i] = vb_v2bool(val) != 0;
        return vb_v_empty();
    }
    if (n == "colposition" || n == "rowposition") {     /* moves a column / row to another place */
        if (!set) { vb_raise(394); return vb_v_empty(); }
        int from = (int)arg_l(na, argv, 0, 0), to = ival();
        bool isrow = n == "rowposition";
        int count = isrow ? rows : cols;
        if (!inrange(from, 0, count - 1) || !inrange(to, 0, count - 1)) return vb_v_empty();
        int other = isrow ? cols : rows;
        std::vector<QTableWidgetItem *> moved;
        for (int k = 0; k < other; k++) moved.push_back(isrow ? tw->takeItem(from, k) : tw->takeItem(k, from));
        int step = to > from ? 1 : -1;
        for (int i = from; i != to; i += step)
            for (int k = 0; k < other; k++) {
                QTableWidgetItem *it = isrow ? tw->takeItem(i + step, k) : tw->takeItem(k, i + step);
                if (isrow) tw->setItem(i, k, it); else tw->setItem(k, i, it);
            }
        for (int k = 0; k < other; k++) { if (isrow) tw->setItem(to, k, moved[k]); else tw->setItem(k, to, moved[k]); }
        if (!isrow) { int a = colalign[from]; colalign.erase(colalign.begin() + from); colalign.insert(colalign.begin() + to, a); }
        paint_sel();
        return vb_v_empty();
    }
    struct { const char *nm; long *p; } colors[] = { { "backcolor", &back }, { "forecolor", &fore }, { "backcolorfixed", &backfixed },
        { "forecolorfixed", &forefixed }, { "backcolorsel", &backsel }, { "forecolorsel", &foresel }, { "backcolorbkg", &backbkg }, { "gridcolor", &gridcolor } };
    for (auto &e : colors) if (n == e.nm) { if (!set) return vb_v_l((vb_i32)*e.p); *e.p = ival(); shape(); return vb_v_empty(); }
    struct { const char *nm; int *p; } ints[] = { { "selectionmode", &selmode }, { "highlight", &highlight }, { "fillstyle", &fillstyle }, { "mergecells", &mergecells }, { "gridlinewidth", &gridlinewidth } };
    for (auto &e : ints) if (n == e.nm) { if (!set) return vb_v_l(*e.p); *e.p = ival(); paint_sel(); return vb_v_empty(); }
    if (n == "redraw") { if (!set) return vbool(redraw); redraw = vb_v2bool(val) != 0; if (redraw) paint_sel(); tw->setUpdatesEnabled(redraw); return vb_v_empty(); }
    if (n == "gridlines") { if (!set) return vb_v_l(tw->showGrid() ? 1 : 0); tw->setShowGrid(ival() != 0); return vb_v_empty(); }
    if (n == "wordwrap") { if (!set) return vbool(wordwrap); wordwrap = vb_v2bool(val) != 0; tw->setWordWrap(wordwrap); return vb_v_empty(); }
    if (n == "scrollbars") { if (!set) return vb_v_l(3); return vb_v_empty(); }
    if (n == "allowuserresizing" || n == "focusrect" || n == "gridlinesfixed" || n == "textstyle" || n == "textstylefixed" ||
        n == "allowbigselection" || n == "scrolltrack" || n == "pictureType" || n == "picturetype")
        return set ? vb_v_empty() : vb_v_l(0);
    if (n == "cellpicture" || n == "picture") return set ? vb_v_empty() : retobj(new_pic(n == "picture" ? tw->grab().toImage() : QImage(), 1));
    *found = 0;
    return vb_v_empty();
}
