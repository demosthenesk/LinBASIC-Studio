/* help.inc - the help of the programs (10.1): App.HelpFile and HelpContextID (F1), What's This
 * (WhatsThisHelp, WhatsThisButton, WhatsThisMode, ShowWhatsThis, WhatsThisHelpID), the Help button of a
 * MsgBox, CommonDialog.ShowHelp. The help files are .chm (HTML Help), read by CHMLib (chm/, LGPL).
 * Linux: a window of our own, as the HTML Help window of Windows (Hide, Back, Print, Options; Contents,
 * Index). Windows: HTML Help itself (hhctrl.ocx), as VB6. The text pop-ups of What's This: ours on both.
 *
 * What VB6 does (VB-TESTS/DEV-10, h1, 2026-11-26):
 *  - F1: the HelpContextID of the control with the focus, else of its container (Frame ...), else of the
 *    form; the HTML Help window opens at that topic, its Contents on the same topic.
 *  - a form with WhatsThisHelp = True: F1 gives the pop-up of the control's WhatsThisHelpID, not the window.
 *  - the pop-ups: the text of cshelp.txt in the .chm (".topic <number>"; LinBASIC also takes a name of a
 *    #define of a .h in the .chm, as the user decided); none: "No help topic is associated with this item.";
 *    the top center of the pop-up at the left edge and the middle of the control (F1, ?, WhatsThisMode),
 *    at the mouse pointer for ShowWhatsThis.
 *  - MsgBox with vbMsgBoxHelpButton: OK and Help; Help opens the help and the MsgBox stays.
 *  - CommonDialog.ShowHelp calls WinHelp, which cannot open a .chm (Windows 7: "Why can't I get Help from
 *    this program?"); LinBASIC opens the .chm (the user's decision).
 *  - chmcmd 3.2.2 (Free Pascal) writes the full path of the project into the context map
 *    ("/home/.../help/main.htm"): such topics are found by their name in the .chm. */

#include <QTextBrowser>
#include <QSplitter>
#include <QTabWidget>
#include <QToolBar>
#include <QMainWindow>
#include <QWhatsThis>
#include <QLibrary>
#include <QDesktopServices>
#include <QHelpEvent>
#if QT_VERSION >= QT_VERSION_CHECK(6, 0, 0)
#include <QStringDecoder>
#else
#include <QTextCodec>
#endif
#include <QInputDialog>
#include <QToolButton>
#include <QTreeWidgetItemIterator>
#include <QCursor>
#include <list>
#include "chm/lbchm.h"

/* ---------------------------------------------------------------- the .chm file */
struct SmNode {                                 /* an entry of the Contents (.hhc) or of the Index (.hhk) */
    QString name;
    QStringList locals, titles;
    std::list<SmNode> kids;
};
struct ChmDoc {
    QString file;
    struct chmFile *h = nullptr;
    QString title, deftopic, tocfile, idxfile;
    quint32 lcid = 0x409;
    std::map<quint32, QString> ctx;             /* #IVB: context number -> its file */
    QSet<QString> paths;                        /* every file of it, lower case, without the first / */
    bool pop_read = false;
    std::map<quint32, QString> pop;             /* cshelp.txt: the text pop-ups */
    bool sm_read = false;
    SmNode toc, idx;
};

static int chm_enum_cb(struct chmFile *, struct chmUnitInfo *ui, void *ctx)
{
    QString p = QString::fromUtf8(ui->path);
    while (p.startsWith('/')) p.remove(0, 1);
    if (!p.isEmpty() && !p.endsWith('/')) static_cast<QSet<QString> *>(ctx)->insert(p.toLower());
    return CHM_ENUMERATOR_CONTINUE;
}
static QByteArray chm_get(ChmDoc *d, QString path)
{
    if (!d || !d->h) return QByteArray();
    if (!path.startsWith('/')) path.prepend('/');
    QByteArray p = path.toUtf8();
    struct chmUnitInfo ui;
    if (chm_resolve_object(d->h, p.constData(), &ui) != CHM_RESOLVE_SUCCESS) return QByteArray();
    if (ui.length > 64u * 1024 * 1024) return QByteArray();
    QByteArray b;
    b.resize((int)ui.length);
    LONGINT64 n = ui.length ? chm_retrieve_object(d->h, &ui, (unsigned char *)b.data(), 0, (LONGINT64)ui.length) : 0;
    b.resize(n > 0 ? (int)n : 0);
    return b;
}
static bool chm_has(ChmDoc *d, const QString &path) { return d->paths.contains(path.toLower()); }
/* the text of the .chm: UTF-8, else the code page of its language */
static QString decode_as(const char *name, const QByteArray &b, bool *ok)
{
#if QT_VERSION >= QT_VERSION_CHECK(6, 0, 0)
    QStringDecoder dec(name);
    if (!dec.isValid()) { *ok = false; return QString(); }
    QString s = dec.decode(b);
    *ok = !dec.hasError();
    return s;
#else                                           /* Qt 5.15 of the Windows programs */
    QTextCodec *c = QTextCodec::codecForName(name);
    if (!c) { *ok = false; return QString(); }
    QTextCodec::ConverterState st;
    QString s = c->toUnicode(b.constData(), b.size(), &st);
    *ok = st.invalidChars == 0;
    return s;
#endif
}
static QString chm_text(ChmDoc *d, const QByteArray &b)
{
    bool ok;
    QString s = decode_as("UTF-8", b, &ok);
    if (ok) return s;
    static const std::map<int, const char *> cp = { { 0x408, "windows-1253" }, { 0x419, "windows-1251" }, { 0x402, "windows-1251" },
        { 0x405, "windows-1250" }, { 0x415, "windows-1250" }, { 0x40e, "windows-1250" }, { 0x41f, "windows-1254" },
        { 0x40d, "windows-1255" }, { 0x401, "windows-1256" }, { 0x425, "windows-1257" }, { 0x426, "windows-1257" },
        { 0x427, "windows-1257" }, { 0x411, "Shift_JIS" }, { 0x804, "GBK" }, { 0x404, "Big5" }, { 0x412, "EUC-KR" } };
    auto it = cp.find((int)(d->lcid & 0x3ff ? d->lcid : 0x409));
    if (it == cp.end()) it = cp.find((int)(d->lcid & 0xffff));
    if (it != cp.end()) {
        s = decode_as(it->second, b, &ok);
        if (!s.isEmpty() || b.isEmpty()) return s;
    }
    s = decode_as("windows-1252", b, &ok);
    if (!s.isEmpty() || b.isEmpty()) return s;
    return QString::fromLatin1(b);
}
static QString cstr_at(const QByteArray &b, int off)
{
    if (off < 0 || off >= b.size()) return QString();
    int e = b.indexOf('\0', off);
    return QString::fromLatin1(b.mid(off, (e < 0 ? b.size() : e) - off));
}
static QString help_full_path(const QString &file)
{
    QString f = file.trimmed();
    int k = f.indexOf('>');                     /* "file.chm>window": the window of HTML Help */
    if (k >= 0) f = f.left(k);
    k = f.indexOf("::");
    if (k > 1) f = f.left(k);                   /* "file.chm::/topic.htm" */
    if (f.isEmpty()) return f;
    QFileInfo fi(native(f));
    if (fi.isRelative()) {                      /* App.HelpFile "h1.chm": next to the program, else the current folder */
        vb_i32 m = vb_mark();
        QString ap = qs(vb_app_path());
        vb_release(m);
        QFileInfo a(native(ap + "/" + f));
        if (a.exists()) return a.absoluteFilePath();
    }
    return fi.absoluteFilePath();
}
static ChmDoc *chm_doc(const QString &file)
{
    static std::map<QString, ChmDoc *> docs;
    QString full = help_full_path(file);
    if (full.isEmpty()) return nullptr;
    auto it = docs.find(full);
    if (it != docs.end()) return it->second;
#ifdef _WIN32
    QByteArray name = full.toUtf8();            /* chm/chm_lib.c: CreateFileW of the UTF-8 name */
#else
    QByteArray name = QFile::encodeName(full);
#endif
    struct chmFile *h = chm_open(name.constData());
    if (!h) return nullptr;
    auto *d = new ChmDoc;
    d->file = full; d->h = h;
    chm_enumerate(h, CHM_ENUMERATE_NORMAL | CHM_ENUMERATE_FILES | CHM_ENUMERATE_SPECIAL | CHM_ENUMERATE_META,
                  chm_enum_cb, &d->paths);
    QByteArray sys = chm_get(d, "#SYSTEM");     /* version, then records: code, length, data */
    QString title;
    for (int p = 4; p + 4 <= sys.size();) {
        int code = qFromLittleEndian<quint16>(sys.constData() + p), len = qFromLittleEndian<quint16>(sys.constData() + p + 2);
        p += 4;
        if (p + len > sys.size()) break;
        QByteArray data = sys.mid(p, len);
        int z = data.indexOf('\0');
        QByteArray s = z >= 0 ? data.left(z) : data;
        switch (code) {
        case 0: d->tocfile = QString::fromLatin1(s); break;
        case 1: d->idxfile = QString::fromLatin1(s); break;
        case 2: d->deftopic = QString::fromLatin1(s); break;
        case 3: title = QString::fromLatin1(s); break;      /* decoded after the language (code 4) */
        case 4: if (len >= 4) d->lcid = qFromLittleEndian<quint32>(data.constData()); break;
        }
        p += len;
    }
    if (!title.isEmpty()) d->title = chm_text(d, title.toLatin1());
    QByteArray ivb = chm_get(d, "#IVB"), strs = chm_get(d, "#STRINGS");
    if (ivb.size() >= 4) {                      /* its size, then pairs: the context number, its string in #STRINGS */
        int n = std::min<int>((int)qFromLittleEndian<quint32>(ivb.constData()), ivb.size() - 4);
        for (int p = 4; p + 8 <= 4 + n; p += 8) {
            quint32 id = qFromLittleEndian<quint32>(ivb.constData() + p);
            int off = (int)qFromLittleEndian<quint32>(ivb.constData() + p + 4);
            QString s = cstr_at(strs, off);
            if (!s.isEmpty() && !d->ctx.count(id)) d->ctx[id] = s;
        }
    }
    if (d->tocfile.isEmpty() || d->idxfile.isEmpty())
        for (const QString &p : d->paths) {
            if (d->tocfile.isEmpty() && p.endsWith(".hhc")) d->tocfile = p;
            if (d->idxfile.isEmpty() && p.endsWith(".hhk")) d->idxfile = p;
        }
    docs[full] = d;
    return d;
}
/* a topic as the .chm has it: "main.htm#part"; empty: not in the file. Full paths (chmcmd 3.2.2): by their
 * last parts */
static QString chm_topic(ChmDoc *d, QString t)
{
    t.replace('\\', '/');
    int k = t.indexOf("::");
    if (k >= 0) t = t.mid(k + 2);
    QString frag;
    int h = t.indexOf('#');
    if (h >= 0) { frag = t.mid(h); t = t.left(h); }
    while (t.startsWith('/')) t.remove(0, 1);
    if (t.isEmpty()) return QString();
    if (chm_has(d, t)) return t + frag;
    for (int i = t.indexOf('/'); i >= 0; i = t.indexOf('/', i + 1))
        if (chm_has(d, t.mid(i + 1))) return t.mid(i + 1) + frag;
    return QString();
}
static QString chm_default(ChmDoc *d)
{
    QString t = chm_topic(d, d->deftopic);
    if (t.isEmpty()) {                          /* no Default topic: the first page of the Contents, or of the file */
        for (const QString &p : d->paths) if ((p.endsWith(".htm") || p.endsWith(".html")) && !p.startsWith("#") && !p.startsWith("$"))
            if (t.isEmpty() || p < t) t = p;
    }
    return t;
}
static QString chm_context(ChmDoc *d, quint32 id)
{
    auto it = d->ctx.find(id);
    return it == d->ctx.end() ? QString() : chm_topic(d, it->second);
}

/* ---------------------------------------------------------------- Contents and Index (sitemaps) */
static QString html_unescape(QString s)
{
    s.replace("&quot;", "\"").replace("&lt;", "<").replace("&gt;", ">").replace("&#39;", "'").replace("&nbsp;", " ");
    static const QRegularExpression num("&#(x?)([0-9a-fA-F]+);");
    for (;;) {
        auto m = num.match(s);
        if (!m.hasMatch()) break;
        bool ok; uint c = m.captured(1).isEmpty() ? m.captured(2).toUInt(&ok) : m.captured(2).toUInt(&ok, 16);
        s.replace(m.capturedStart(), m.capturedLength(), ok ? QString(QChar((char16_t)c)) : QString());
    }
    return s.replace("&amp;", "&");
}
static void sm_parse(const QString &html, SmNode *root)
{
    static const QRegularExpression tag("<\\s*(/?)\\s*(ul|object|param)\\b([^>]*)>", QRegularExpression::CaseInsensitiveOption);
    static const QRegularExpression attr("([A-Za-z]+)\\s*=\\s*(\"([^\"]*)\"|'([^']*)'|([^\\s>]+))");
    std::vector<SmNode *> stack{ root };
    SmNode *last = nullptr, cur;
    bool inobj = false;
    auto it = tag.globalMatch(html);
    while (it.hasNext()) {
        auto m = it.next();
        bool close = !m.captured(1).isEmpty();
        QString t = m.captured(2).toLower();
        std::map<QString, QString> a;
        auto ai = attr.globalMatch(m.captured(3));
        while (ai.hasNext()) {
            auto am = ai.next();
            QString v = !am.captured(3).isNull() ? am.captured(3) : !am.captured(4).isNull() ? am.captured(4) : am.captured(5);
            a[am.captured(1).toLower()] = html_unescape(v);
        }
        if (t == "ul") {
            if (!close) { stack.push_back(last ? last : stack.back()); last = nullptr; }
            else if (stack.size() > 1) { last = stack.back() == root ? nullptr : stack.back(); stack.pop_back(); }
        } else if (t == "object") {
            if (!close) { inobj = a["type"].compare("text/sitemap", Qt::CaseInsensitive) == 0; cur = SmNode(); }
            else if (inobj) {
                inobj = false;
                if (!cur.name.isEmpty() || !cur.locals.isEmpty()) { stack.back()->kids.push_back(cur); last = &stack.back()->kids.back(); }
            }
        } else if (t == "param" && inobj) {
            QString n = a["name"].toLower(), v = a["value"];
            if (n == "name") { if (cur.name.isEmpty()) cur.name = v; else cur.titles << v; }
            else if (n == "local") cur.locals << v;
        }
    }
}
static void chm_sitemaps(ChmDoc *d)
{
    if (d->sm_read) return;
    d->sm_read = true;
    if (!d->tocfile.isEmpty()) sm_parse(chm_text(d, chm_get(d, d->tocfile)), &d->toc);
    if (!d->idxfile.isEmpty()) sm_parse(chm_text(d, chm_get(d, d->idxfile)), &d->idx);
}
/* the topic of a keyword of the Index (HelpKey): one topic; empty: none, or more than one */
static QString chm_key_topic(ChmDoc *d, const SmNode *n, const QString &key)
{
    for (const SmNode &k : n->kids) {
        if (k.name.compare(key, Qt::CaseInsensitive) == 0)
            return k.locals.size() == 1 ? chm_topic(d, k.locals.first()) : QString();
        QString t = chm_key_topic(d, &k, key);
        if (!t.isEmpty()) return t;
    }
    return QString();
}

/* ---------------------------------------------------------------- the text pop-ups (cshelp.txt) */
static void chm_popups(ChmDoc *d)
{
    if (d->pop_read) return;
    d->pop_read = true;
    std::map<QString, quint32> names;           /* the #define of the .h files in the .chm (LinBASIC: names too) */
    static const QRegularExpression def("^\\s*#define\\s+(\\w+)\\s+(0[xX][0-9a-fA-F]+|\\d+)", QRegularExpression::MultilineOption);
    for (const QString &p : d->paths) {
        if (!p.endsWith(".h") && !p.endsWith(".hh") && !p.endsWith(".hm")) continue;
        auto it = def.globalMatch(chm_text(d, chm_get(d, p)));
        while (it.hasNext()) {
            auto m = it.next(); bool ok;
            quint32 v = m.captured(2).startsWith("0x", Qt::CaseInsensitive) ? m.captured(2).mid(2).toUInt(&ok, 16) : m.captured(2).toUInt(&ok);
            if (ok) names[m.captured(1).toUpper()] = v;
        }
    }
    QString txt = chm_text(d, chm_get(d, "cshelp.txt"));
    txt.replace("\r\n", "\n"); txt.replace('\r', '\n');
    bool have = false; quint32 id = 0; QStringList lines;
    auto flush = [&]() {
        if (!have) return;
        while (!lines.isEmpty() && lines.last().trimmed().isEmpty()) lines.removeLast();
        while (!lines.isEmpty() && lines.first().trimmed().isEmpty()) lines.removeFirst();
        QString t;                              /* the lines of a paragraph go on one line; an empty line: a new paragraph */
        bool para = false;
        for (const QString &l : lines) {
            QString s = l.trimmed();
            if (s.isEmpty()) { para = true; continue; }
            if (!t.isEmpty()) t += para ? "\n\n" : " ";
            t += s; para = false;
        }
        if (!d->pop.count(id)) d->pop[id] = t;
        have = false; lines.clear();
    };
    for (const QString &l : txt.split('\n')) {
        QString s = l.trimmed();
        if (s.startsWith(".topic", Qt::CaseInsensitive) && (s.size() == 6 || s[6].isSpace())) {
            flush();
            QString key = s.mid(6).trimmed();
            int sp = key.indexOf(QRegularExpression("\\s")); if (sp > 0) key = key.left(sp);
            bool ok; quint32 v = key.startsWith("0x", Qt::CaseInsensitive) ? key.mid(2).toUInt(&ok, 16) : key.toUInt(&ok);
            if (!ok) { auto n = names.find(key.toUpper()); ok = n != names.end(); if (ok) v = n->second; }
            if (ok) { have = true; id = v; }
            continue;
        }
        if (have) lines << l;
    }
    flush();
}

/* ---------------------------------------------------------------- the log of the tests (VBRT_HELP_LOG) */
static void help_log(const QString &line)
{
    const char *f = getenv("VBRT_HELP_LOG");
    if (!f || !*f) return;
    QFile o(QString::fromUtf8(f));
    if (o.open(QIODevice::Append)) o.write((line + "\n").toUtf8());
}
static bool help_quiet() { return monkey_on || (getenv("VBRT_HELP_LOG") && *getenv("VBRT_HELP_LOG")); }

/* ---------------------------------------------------------------- the pop-up */
class HelpPopup : public QFrame {
public:
    explicit HelpPopup(const QString &text) : QFrame(nullptr, Qt::Popup | Qt::FramelessWindowHint)
    {
        setAttribute(Qt::WA_DeleteOnClose);
        setFrameStyle(QFrame::Box | QFrame::Plain);
        setLineWidth(1);
        QPalette pal = palette();
        pal.setColor(QPalette::Window, QColor(0xFF, 0xFF, 0xE1));
        pal.setColor(QPalette::WindowText, Qt::black);
        setPalette(pal);
        setAutoFillBackground(true);
        auto *l = new QLabel(text, this);
        l->setTextFormat(Qt::PlainText);
        l->setWordWrap(true);
        l->setPalette(pal);
        auto *lay = new QVBoxLayout(this);
        lay->setContentsMargins(6, 4, 6, 5);
        lay->addWidget(l);
        QFontMetrics fm(l->font());             /* wider than high, as HTML Help draws them */
        int tw = 0;
        for (const QString &p : text.split('\n')) tw = std::max(tw, fm.horizontalAdvance(p));
        int w = std::max(100, std::min(tw, (int)std::sqrt((double)tw * fm.height() * 8.0)));
        l->setFixedWidth(std::min(w, 420) + 2);
        adjustSize();
    }
protected:
    void keyPressEvent(QKeyEvent *) override { close(); }
    void mousePressEvent(QMouseEvent *) override { close(); }
};
/* the pop-up of a context number of a help file, its top center at pt (global pixels) */
static void help_popup(const QString &file, quint32 id, QPoint pt)
{
    ChmDoc *d = file.isEmpty() ? nullptr : chm_doc(file);
    QString text;
    if (d) { chm_popups(d); auto it = d->pop.find(id); if (it != d->pop.end()) text = it->second; }
    if (text.isEmpty()) text = QString::fromUtf8(greek() ? "Δεν υπάρχει θέμα βοήθειας για αυτό το στοιχείο." : "No help topic is associated with this item.");
    help_log(QString("popup %1 %2").arg(id).arg(text));
    if (help_quiet()) { if (monkey_on) monkey_log("What's This: " + text.left(80)); return; }
    auto *p = new HelpPopup(text);
    QRect scr = QGuiApplication::screenAt(pt) ? QGuiApplication::screenAt(pt)->availableGeometry() : QRect(0, 0, 1024, 768);
    QPoint at(pt.x() - p->width() / 2, pt.y());
    at.setX(std::max(scr.left(), std::min(at.x(), scr.right() - p->width())));
    if (at.y() + p->height() > scr.bottom()) at.setY(std::max(scr.top(), pt.y() - p->height()));
    p->move(at);
    p->show();
}

/* ---------------------------------------------------------------- the window (Linux) */
class HelpView : public QTextBrowser {
public:
    ChmDoc *d;
    explicit HelpView(ChmDoc *doc, QWidget *parent) : QTextBrowser(parent), d(doc)
    {
        setOpenExternalLinks(false);
        setOpenLinks(false);
        setStyleSheet("QTextBrowser { background: #FFFFFF; color: #000000; }");
    }
    QVariant loadResource(int type, const QUrl &name) override
    {
        if (name.scheme() != "chm") return QTextBrowser::loadResource(type, name);
        QString t = chm_topic(d, name.path());
        if (t.isEmpty()) return QVariant();
        QByteArray b = chm_get(d, t.section('#', 0, 0));
        if (type == QTextDocument::ImageResource) { QImage im; if (im.loadFromData(b)) return im; return QVariant(); }
        if (type == QTextDocument::StyleSheetResource) return chm_text(d, b);
        return b;
    }
};
static QString greek_or(const char *en, const char *el) { return QString::fromUtf8(greek() ? el : en); }
class HelpWin : public QMainWindow {
public:
    ChmDoc *d;
    HelpView *view;
    QTabWidget *tabs;
    QTreeWidget *toc;
    QLineEdit *kw;
    QListWidget *idx;
    QAction *a_hide, *a_back, *a_fwd;
    QString cur;
    explicit HelpWin(ChmDoc *doc) : QMainWindow(nullptr), d(doc)
    {
        setAttribute(Qt::WA_QuitOnClose, false);
        chm_sitemaps(d);
        setWindowTitle(d->title.isEmpty() ? QFileInfo(d->file).completeBaseName() : d->title);
        setWindowIcon(style()->standardIcon(QStyle::SP_DialogHelpButton));
        auto *tb = addToolBar("help");
        tb->setMovable(false);
        tb->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
        tb->setIconSize(QSize(20, 20));
        a_hide = tb->addAction(style()->standardIcon(QStyle::SP_FileDialogListView), greek_or("Hide", "Απόκρυψη"));
        a_back = tb->addAction(style()->standardIcon(QStyle::SP_ArrowBack), greek_or("Back", "Πίσω"));
        auto *a_print = tb->addAction(style()->standardIcon(QStyle::SP_FileDialogContentsView), greek_or("Print", "Εκτύπωση"));
        auto *opt = new QMenu(this);
        auto *tbo = new QToolButton(tb);
        tbo->setText(greek_or("Options", "Επιλογές"));
        tbo->setIcon(style()->standardIcon(QStyle::SP_FileDialogDetailedView));
        tbo->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);
        tbo->setPopupMode(QToolButton::InstantPopup);
        tbo->setMenu(opt);
        tb->addWidget(tbo);
        auto *o_hide = opt->addAction(greek_or("Hide Tabs", "Απόκρυψη καρτελών"));
        auto *o_back = opt->addAction(greek_or("Back", "Πίσω"));
        a_fwd = opt->addAction(greek_or("Forward", "Εμπρός"));
        auto *o_home = opt->addAction(greek_or("Home", "Αρχική"));
        opt->addSeparator();
        auto *o_print = opt->addAction(greek_or("Print...", "Εκτύπωση..."));
        auto *split = new QSplitter(this);
        tabs = new QTabWidget(split);
        toc = new QTreeWidget;
        toc->setHeaderHidden(true);
        tabs->addTab(toc, greek_or("Contents", "Περιεχόμενα"));
        auto *ip = new QWidget;
        auto *il = new QVBoxLayout(ip);
        il->addWidget(new QLabel(greek_or("Type in the keyword to find:", "Πληκτρολογήστε τη λέξη-κλειδί:")));
        kw = new QLineEdit;
        il->addWidget(kw);
        idx = new QListWidget;
        il->addWidget(idx, 1);
        auto *disp = new QPushButton(greek_or("Display", "Εμφάνιση"));
        auto *dr = new QHBoxLayout; dr->addStretch(); dr->addWidget(disp);
        il->addLayout(dr);
        tabs->addTab(ip, greek_or("Index", "Ευρετήριο"));
        view = new HelpView(d, split);
        split->addWidget(tabs);
        split->addWidget(view);
        split->setStretchFactor(0, 0);
        split->setStretchFactor(1, 1);
        split->setSizes({ 210, 380 });
        setCentralWidget(split);
        fill_toc(&d->toc, nullptr);
        toc->expandAll();
        fill_idx(&d->idx, 0);
        QObject::connect(toc, &QTreeWidget::itemClicked, [this](QTreeWidgetItem *it, int) { go(it->data(0, Qt::UserRole).toString()); });
        QObject::connect(toc, &QTreeWidget::itemActivated, [this](QTreeWidgetItem *it, int) { go(it->data(0, Qt::UserRole).toString()); });
        QObject::connect(kw, &QLineEdit::textEdited, [this](const QString &s) { find_key(s); });
        auto show_sel = [this]() { show_entry(idx->currentItem()); };
        QObject::connect(kw, &QLineEdit::returnPressed, show_sel);
        QObject::connect(disp, &QPushButton::clicked, show_sel);
        QObject::connect(idx, &QListWidget::itemActivated, [this](QListWidgetItem *it) { show_entry(it); });
        QObject::connect(view, &QTextBrowser::anchorClicked, [this](const QUrl &u) {
            QString s = u.toString();
            if (u.scheme() == "chm") go(u.path() + (u.hasFragment() ? "#" + u.fragment() : QString()));
            else if ((s.startsWith("ms-its:", Qt::CaseInsensitive) || s.startsWith("mk:@MSITStore:", Qt::CaseInsensitive))
                     && s.contains("::")) go(s.mid(s.indexOf("::") + 2));   /* a topic of this .chm by its full name */
            else if (u.scheme().isEmpty()) go(view->source().resolved(u).path() + (u.hasFragment() ? "#" + u.fragment() : QString()));
            else QDesktopServices::openUrl(u);
        });
        QObject::connect(view, &QTextBrowser::backwardAvailable, a_back, &QAction::setEnabled);
        QObject::connect(view, &QTextBrowser::backwardAvailable, o_back, &QAction::setEnabled);
        QObject::connect(view, &QTextBrowser::forwardAvailable, a_fwd, &QAction::setEnabled);
        QObject::connect(view, &QTextBrowser::sourceChanged, [this](const QUrl &u) { synced(u); });
        a_back->setEnabled(false); o_back->setEnabled(false); a_fwd->setEnabled(false);
        QObject::connect(a_back, &QAction::triggered, view, &QTextBrowser::backward);
        QObject::connect(o_back, &QAction::triggered, view, &QTextBrowser::backward);
        QObject::connect(a_fwd, &QAction::triggered, view, &QTextBrowser::forward);
        QObject::connect(o_home, &QAction::triggered, [this]() { go(chm_default(d)); });
        auto hide = [this, o_hide]() {
            bool vis = !tabs->isVisible();
            tabs->setVisible(vis);
            a_hide->setText(vis ? greek_or("Hide", "Απόκρυψη") : greek_or("Show", "Εμφάνιση"));
            o_hide->setText(vis ? greek_or("Hide Tabs", "Απόκρυψη καρτελών") : greek_or("Show Tabs", "Εμφάνιση καρτελών"));
        };
        QObject::connect(a_hide, &QAction::triggered, hide);
        QObject::connect(o_hide, &QAction::triggered, hide);
        auto print = [this]() {                 /* the printer of the programs is a PDF file (Printer) */
            QString f = QFileDialog::getSaveFileName(this, greek_or("Print", "Εκτύπωση"),
                QFileInfo(d->file).absolutePath() + "/" + QFileInfo(cur.section('#', 0, 0)).completeBaseName() + ".pdf", "PDF (*.pdf)");
            if (f.isEmpty()) return;
            QPdfWriter pdf(f);
            pdf.setPageSize(QPageSize(QPageSize::A4));
            view->document()->print(&pdf);
        };
        QObject::connect(a_print, &QAction::triggered, print);
        QObject::connect(o_print, &QAction::triggered, print);
        resize(600, 480);
        if (QScreen *s = QGuiApplication::primaryScreen()) {   /* HTML Help: on the right of the screen */
            QRect a = s->availableGeometry();
            move(std::max(a.left(), a.right() - width() - a.width() / 16), a.top() + a.height() / 5);
        }
    }
    void fill_toc(const SmNode *n, QTreeWidgetItem *parent)
    {
        for (const SmNode &k : n->kids) {
            auto *it = parent ? new QTreeWidgetItem(parent) : new QTreeWidgetItem(toc);
            it->setText(0, k.name);
            QString t = k.locals.isEmpty() ? QString() : chm_topic(d, k.locals.first());
            it->setData(0, Qt::UserRole, t);
            it->setIcon(0, style()->standardIcon(k.kids.empty() ? QStyle::SP_FileIcon : QStyle::SP_DirIcon));
            fill_toc(&k, it);
        }
    }
    void fill_idx(const SmNode *n, int depth)
    {
        for (const SmNode &k : n->kids) {
            auto *it = new QListWidgetItem(QString(depth * 4, ' ') + k.name, idx);
            QStringList ts;
            for (const QString &l : k.locals) { QString t = chm_topic(d, l); if (!t.isEmpty()) ts << t; }
            it->setData(Qt::UserRole, ts);
            it->setData(Qt::UserRole + 1, k.titles);
            it->setData(Qt::UserRole + 2, k.name);
            fill_idx(&k, depth + 1);
        }
    }
    QListWidgetItem *find_key(const QString &s)
    {
        for (int i = 0; i < idx->count(); i++)
            if (idx->item(i)->data(Qt::UserRole + 2).toString().startsWith(s, Qt::CaseInsensitive)) {
                idx->setCurrentRow(i);
                idx->scrollToItem(idx->item(i), QAbstractItemView::PositionAtTop);
                return idx->item(i);
            }
        return nullptr;
    }
    void show_entry(QListWidgetItem *it)
    {
        if (!it) return;
        QStringList ts = it->data(Qt::UserRole).toStringList(), titles = it->data(Qt::UserRole + 1).toStringList();
        if (ts.isEmpty()) return;
        if (ts.size() == 1) { go(ts.first()); return; }
        QStringList names;                      /* "Topics Found" */
        for (int i = 0; i < ts.size(); i++) names << (i < titles.size() && !titles[i].isEmpty() ? titles[i] : ts[i]);
        bool ok = false;
        QString pick = QInputDialog::getItem(this, greek_or("Topics Found", "Θέματα που βρέθηκαν"),
                                             greek_or("Click a topic, then click Display.", "Επιλέξτε ένα θέμα."), names, 0, false, &ok);
        if (ok) go(ts[names.indexOf(pick)]);
    }
    void go(const QString &topic)
    {
        if (topic.isEmpty()) return;
        QString t = chm_topic(d, topic);
        if (t.isEmpty()) {                      /* HTML Help: the page cannot be shown */
            view->setHtml("<h3>" + greek_or("This page can't be displayed", "Δεν είναι δυνατή η εμφάνιση της σελίδας") +
                          "</h3><p>" + topic.toHtmlEscaped() + "</p>");
            cur = topic;
            return;
        }
        cur = t;
        view->setSource(QUrl("chm:/" + t));
    }
    void synced(const QUrl &u)
    {
        QString p = u.path(); while (p.startsWith('/')) p.remove(0, 1);
        p = p.section('#', 0, 0);
        QTreeWidgetItemIterator it(toc);
        for (; *it; ++it)
            if ((*it)->data(0, Qt::UserRole).toString().section('#', 0, 0).compare(p, Qt::CaseInsensitive) == 0) {
                toc->setCurrentItem(*it);
                return;
            }
    }
};
static std::map<ChmDoc *, QPointer<HelpWin>> help_wins;
static HelpWin *help_win(ChmDoc *d)
{
    QPointer<HelpWin> &w = help_wins[d];
    if (!w) w = new HelpWin(d);
    /* the Help of a MsgBox: the window is usable while the MsgBox waits (a window of the dialog is not
     * blocked by it), and stays when the MsgBox closes (as HTML Help) */
    if (auto *dlg = qobject_cast<QDialog *>(QApplication::activeModalWidget())) {
        if (w->parentWidget() != dlg) {
            QRect g = w->geometry();
            w->setParent(dlg, Qt::Window);
            w->setGeometry(g);
            QPointer<HelpWin> hw = w;
            QObject::connect(dlg, &QDialog::finished, [hw]() {
                if (!hw) return;
                bool vis = hw->isVisible();
                QRect g2 = hw->geometry();
                hw->setParent(nullptr, Qt::Window);
                hw->setGeometry(g2);
                if (vis) hw->show();
            });
        }
    }
    w->show();
    if (w->isMinimized()) w->showNormal();
    w->raise();
    w->activateWindow();
    return w;
}
static void help_close_all()
{
    for (auto &p : help_wins) if (p.second) p.second->close();
}

/* ---------------------------------------------------------------- Windows: HTML Help */
#ifdef _WIN32
typedef void *(__stdcall *HtmlHelpW_t)(void *, const wchar_t *, unsigned, quintptr);
static HtmlHelpW_t html_help()
{
    static QLibrary lib("hhctrl.ocx");
    static HtmlHelpW_t f = (HtmlHelpW_t)lib.resolve("HtmlHelpW");
    return f;
}
static void *help_owner()                      /* the form (not a MsgBox: the help stays when it closes) */
{
    QWidget *w = dialog_parent();
    if (!w) w = QApplication::activeWindow();
    return w ? (void *)w->winId() : nullptr;
}
#endif

/* what = 0 a topic (or the Default topic), 1 the Contents, 2 the Index at a keyword */
static void help_open(const QString &file, int what, const QString &topic, const QString &key, const char *why, quint32 id)
{
    if (file.trimmed().isEmpty()) return;
    QString full = help_full_path(file);
    QString ext = QFileInfo(full).suffix().toLower();
    ChmDoc *d = chm_doc(full);
    QString t = d ? (topic.isEmpty() ? chm_default(d) : topic) : QString();
    if (d && what == 2) { chm_sitemaps(d); t = chm_key_topic(d, &d->idx, key); }
    help_log(QString("%1 %2 %3%4").arg(why, QFileInfo(full).fileName(), what == 2 ? key + " " : QString(), t.isEmpty() ? QString("-") : t));
    if (help_quiet()) { if (monkey_on) monkey_log(QString("Help %1: %2 %3").arg(why, QFileInfo(full).fileName(), t)); return; }
#ifdef _WIN32
    if (HtmlHelpW_t hh = (ext == "chm" ? html_help() : nullptr)) {
        std::wstring f = QDir::toNativeSeparators(full).toStdWString();
        if (what == 2) { std::wstring k = key.toStdWString(); hh(help_owner(), f.c_str(), 0x0002 /* HH_DISPLAY_INDEX */, (quintptr)k.c_str()); return; }
        if (what == 0 && !topic.isEmpty()) {
            std::wstring ft = f + L"::/" + topic.toStdWString();
            hh(help_owner(), ft.c_str(), 0x0000 /* HH_DISPLAY_TOPIC */, 0);
            return;
        }
        if (what == 0 && id && d && topic.isEmpty() && d->ctx.count(id)) { hh(help_owner(), f.c_str(), 0x000F /* HH_HELP_CONTEXT */, id); return; }
        hh(help_owner(), f.c_str(), what == 1 ? 0x0001 /* HH_DISPLAY_TOC */ : 0x0000, 0);
        return;
    }
    if (ext == "hlp") {                         /* WinHelp, as VB6 (Windows 7 has it no more: its message) */
        typedef int (__stdcall *WinHelpW_t)(void *, const wchar_t *, unsigned, quintptr);
        static QLibrary u32("user32");
        if (auto wh = (WinHelpW_t)u32.resolve("WinHelpW")) {
            std::wstring f = QDir::toNativeSeparators(full).toStdWString();
            wh(help_owner(), f.c_str(), what == 1 ? 0x000B /* HELP_FINDER */ : 0x0001 /* HELP_CONTEXT */, id);
        }
        return;
    }
#endif
    if (ext == "hlp" && !d) {                   /* the WinHelp files: not read yet (decision after the .chm) */
        QMessageBox::warning(dialog_parent(), "HTML Help",
            greek_or("The help file %1 is a Windows Help (.hlp) file: LinBASIC shows .chm help files.",
                     "Το αρχείο βοήθειας %1 είναι Windows Help (.hlp): το LinBASIC δείχνει αρχεία βοήθειας .chm.").arg(QDir::toNativeSeparators(full)));
        return;
    }
    if (!d) {                                   /* HTML Help: "Cannot open the file: ..." */
        QMessageBox::warning(dialog_parent(), "HTML Help",
            greek_or("Cannot open the file: %1.", "Δεν είναι δυνατό το άνοιγμα του αρχείου: %1.").arg(QDir::toNativeSeparators(full)));
        return;
    }
    HelpWin *w = help_win(d);
    if (what == 1) w->tabs->setCurrentIndex(0);
    if (what == 2) {
        w->tabs->setCurrentIndex(1);
        w->kw->setText(key);
        w->find_key(key);
        if (!t.isEmpty()) w->go(t);
        else if (w->cur.isEmpty()) w->go(chm_default(d));
        return;
    }
    w->go(t);
}
/* a context number of a help file (F1, the Help of MsgBox, ShowHelp cdlHelpContext) */
static void help_context(const QString &file, quint32 id, const char *why)
{
    if (file.trimmed().isEmpty()) return;
    ChmDoc *d = chm_doc(file);
    QString t = d && id ? chm_context(d, id) : QString();
    if (d && id && t.isEmpty() && d->ctx.count(id)) t = d->ctx[id];   /* in the map, not in the file: the window says so */
    help_open(file, 0, t, QString(), why, id);
}

/* ---------------------------------------------------------------- F1, What's This */
static vb_i32 ctl_helpid(Ctl *c, const char *name)
{
    auto it = c->misc.find(name);
    if (it != c->misc.end()) return it->second;
    return c->def ? (vb_i32)propn(c->def->props, name[0] == 'w' ? "WhatsThisHelpID" : "HelpContextID", 0) : 0;
}
static QString app_helpfile()
{
    vb_i32 m = vb_mark();
    QString f = qs(vb_app_helpfile());
    vb_release(m);
    return f;
}
/* where VB6 puts the pop-up of a control: its left edge, half its height */
static QPoint popup_point(Ctl *c)
{
    QWidget *w = c ? (QWidget *)c->w.data() : nullptr;
    if (!w) return QCursor::pos();
    return w->mapToGlobal(QPoint(0, w->height() / 2));
}
static void whatsthis_ctl(Ctl *c, QPoint pt)
{
    help_popup(app_helpfile(), c ? (quint32)ctl_helpid(c, "whatsthishelpid") : 0, pt);
}
/* F1 on a form: the HelpContextID of the control, its containers, the form; WhatsThisHelp: the pop-up */
static void help_f1(FormState *st, Ctl *c)
{
    if (st->whatsthis) {
        if (c) whatsthis_ctl(c, popup_point(c));
        return;
    }
    vb_i32 id = 0;
    for (Ctl *x = c; x && !id; x = x->container ? C(x->container) : nullptr) id = ctl_helpid(x, "helpcontextid");
    if (!id) id = st->helpctx;
    help_context(app_helpfile(), (quint32)id, "F1");
}
/* F1 in an open menu: the HelpContextID of the item, of the menus above it, of the form */
static bool help_menu_f1(QMenu *m)
{
    QAction *a = m ? m->activeAction() : nullptr;
    QVariant v = a ? a->property("vbmenu") : QVariant();
    if (!v.isValid()) return false;
    Ctl *c = C((vbobj)(quintptr)v.toULongLong());
    vb_i32 id = 0;
    for (QMenu *p = m; !id && c; ) {
        id = ctl_helpid(c, "helpcontextid");
        if (id) break;
        a = p ? p->menuAction() : nullptr;
        v = a ? a->property("vbmenu") : QVariant();
        if (!v.isValid()) break;
        c = C((vbobj)(quintptr)v.toULongLong());
        p = qobject_cast<QMenu *>(p->parentWidget());
    }
    FormState *st = c ? c->form : nullptr;
    if (!id && st) id = st->helpctx;
    if (m) m->close();
    help_context(app_helpfile(), (quint32)id, "F1");
    return true;
}
/* the What's This mode of Qt (the ? of the title bar, WhatsThisMode): a click on a control of a form */
static bool help_whatsthis_event(QObject *obj, QEvent *e)
{
    if (e->type() != QEvent::WhatsThis && e->type() != QEvent::QueryWhatsThis) return false;
    QWidget *w = qobject_cast<QWidget *>(obj);
    FormState *st = w ? form_of_widget(w) : nullptr;
    if (!st || !st->loaded) return false;
    e->accept();
    if (e->type() == QEvent::QueryWhatsThis) return true;
    Ctl *c = ctl_of_widget(w);
    if (c && c->kind == K_USER) { Ctl *in = uctl_inner_focus(c); if (in) c = in; }
    whatsthis_ctl(c, c ? popup_point(c) : static_cast<QHelpEvent *>(e)->globalPos());
    return true;
}
static void help_whatsthis_mode()
{
    if (help_quiet()) { help_log("whatsthismode"); return; }
    QWhatsThis::enterWhatsThisMode();
}

/* ---------------------------------------------------------------- CommonDialog.ShowHelp */
static void dlg_showhelp(Ctl *c)
{
    CDlg *d = c->dlg;
    QString file = d->helpfile;
    switch (d->helpcommand) {
    case 1: help_context(file, (quint32)d->helpcontext, "context"); break;          /* cdlHelpContext */
    case 2: if (!help_quiet()) help_close_all(); help_log("quit"); break;            /* cdlHelpQuit */
    case 8: help_popup(file, (quint32)d->helpcontext, QCursor::pos()); break;       /* cdlHelpContextPopup */
    case 0x101: case 0x105:                                                         /* cdlHelpKey, cdlHelpPartialKey */
        help_open(file, 2, QString(), d->helpkey, "key", 0); break;
    default: help_open(file, 1, QString(), QString(), "contents", 0); break;        /* cdlHelpContents, Index ... */
    }
}

/* ---------------------------------------------------------------- the Help button of MsgBox */
static void msgbox_help()
{
    vb_i32 m = vb_mark();
    vbs f;
    vb_i32 ctx = vb_msgbox_helpinfo(&f);
    QString file = qs(f);
    vb_release(m);
    help_context(file, (quint32)ctx, "msgbox");
}
