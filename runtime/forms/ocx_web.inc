/* ocx_web.inc - the WebBrowser control (Microsoft Internet Controls, SHDocVwCtl.WebBrowser) of LinBASIC Studio,
 * phase 9.2 (the Application Wizard's frmBrowser and the Web Browser form template). A simple browser on
 * QTextBrowser, as the user decided (2026-11-19): HTML and CSS that Qt shows, no JavaScript, no plugins.
 * Navigate / Navigate2, GoBack, GoForward, GoHome, GoSearch, Refresh, Refresh2, Stop; LocationURL, LocationName,
 * Busy, ReadyState, Type; the events in the order of VB6 with Internet Explorer (oracle VB-TESTS/DEV-9/wb1, Windows
 * 7): BeforeNavigate2 (Cancel), DownloadBegin, CommandStateChange, TitleChange, NavigateComplete2, StatusTextChange,
 * ProgressChange, DownloadComplete, DocumentComplete, NavigateError (HTTP).
 * Pages come from files (file:, paths, \\server\share on Windows), http: and https: (runtime/inet.c, in a thread:
 * the events come from the message loop, as in VB6) and about:blank. A link the user clicks navigates (Flags 0).
 * Not as Internet Explorer, on purpose:
 *   - GoHome and GoSearch go to about:blank (IE: its home page, go.microsoft.com ...);
 *   - Document is Nothing (IE: the HTMLDocument of MSHTML);
 *   - the pictures of a page come from files only (not from http:);
 *   - PropertyChange (internal GUIDs of IE) is not raised; DownloadBegin once (IE: twice).
 * Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <QTextBrowser>
#include <QUrl>
#include <QDir>
#include <QFileInfo>
#include <QFile>
#include <QImage>
#include <QRegularExpression>
#include <QTextDocument>

/* the view: the pictures of a page relative to the page, from files */
struct WebView : QTextBrowser {
    QUrl base;
    using QTextBrowser::QTextBrowser;
    QVariant loadResource(int type, const QUrl &name) override
    {
        QUrl u = base.isValid() ? base.resolved(name) : name;
        if (u.isLocalFile()) {
            QFile f(u.toLocalFile());
            if (f.open(QIODevice::ReadOnly)) {
                QByteArray b = f.readAll();
                if (type == QTextDocument::ImageResource) { QImage im; if (im.loadFromData(b)) return im; }
                return b;
            }
        }
        return QVariant();
    }
};

struct OcxWebBrowser : Ocx {
    struct Entry { QString given; QUrl url; };
    QPointer<WebView> v;
    Entry cur;                                  /* the page shown (cur.url empty: none yet) */
    std::vector<Entry> back, fwd;
    QString name, status, html;
    bool busy = false, has_doc = false;
    int rs = 0;                                 /* ReadyState: 0 before the first page (oracle wb1: 1 at Load) */
    bool offline = false, silent = false, regbrowser = false, regdrop = true, addressbar = true, menubar = true,
         statusbar = true, fullscreen = false, theater = false, resizable = true;
    long toolbar = 1;
    std::map<QString, QString> props;           /* PutProperty / GetProperty */
    int gen = 0;                                /* the navigation in progress: a newer one stops it */
    void *job = nullptr;
    QTimer *poll = nullptr;
    std::shared_ptr<bool> alive = std::make_shared<bool>(true);

    ~OcxWebBrowser() override
    {
        *alive = false;
        if (job) lbinet_release(job);
        if (poll) { poll->stop(); poll->deleteLater(); }
    }
    QWidget *make(QWidget *parent) override
    {
        v = new WebView(parent);
        {                                       /* a page as a browser shows it: black on white, whatever the theme */
            QPalette pal = v->palette();
            for (QPalette::ColorGroup g : { QPalette::Active, QPalette::Inactive, QPalette::Disabled }) {
                pal.setColor(g, QPalette::Base, QColor(255, 255, 255));
                pal.setColor(g, QPalette::Text, QColor(0, 0, 0));
                pal.setColor(g, QPalette::Link, QColor(0, 0, 238));
                pal.setColor(g, QPalette::LinkVisited, QColor(85, 26, 139));
            }
            v->setPalette(pal);
            v->setStyleSheet("QTextBrowser { background: #FFFFFF; color: #000000; }");   /* over a dark style sheet */
        }
        v->setOpenLinks(false);
        v->setOpenExternalLinks(false);
        vbobj self = c->self;
        QObject::connect(v, &QTextBrowser::anchorClicked, [self](const QUrl &u) {
            Ctl *x = C(self);
            auto *o = x ? dynamic_cast<OcxWebBrowser *>(x->ocx) : nullptr;
            if (!o) return;
            QUrl t = o->v && o->v->base.isValid() ? o->v->base.resolved(u) : u;
            o->navigate(t.toString(), 0, 0);
        });
        rs = 1;                                 /* VB6: ReadyState 1 at Form_Load (oracle wb1) */
        return v;
    }
    void load(const char *const *p) override
    {
        offline = propn(p, "Offline", 0) != 0;
        silent = propn(p, "Silent", 0) != 0;
        regbrowser = propn(p, "RegisterAsBrowser", 0) != 0;
        regdrop = propn(p, "RegisterAsDropTarget", 1) != 0;
    }

    /* ------------------------------------------------ events */
    bool fire(const char *ev, int argc = 0, vbvar *argv = nullptr)
    {
        std::shared_ptr<bool> a = alive;
        fire_ctl(c, ev, argc, argv);
        return *a;
    }
    bool ev_s(const char *ev, const QString &s) { vbvar a = vstr(s); bool r = fire(ev, 1, &a); if (r) vb_vfree(&a); return r; }
    bool ev_csc(long cmd, bool on) { vbvar a[2] = { vb_v_l((vb_i32)cmd), vbool(on) }; return fire("commandstatechange", 2, a); }
    bool ev_progress(long p, long m) { vbvar a[2] = { vb_v_l((vb_i32)p), vb_v_l((vb_i32)m) }; return fire("progresschange", 2, a); }
    bool ev_url(const char *ev, const QString &given)
    {
        vbvar a[2] = { vb_v_obj(c->self), vstr(given) };
        bool r = fire(ev, 2, a);
        if (r) vb_vfree(&a[1]);
        return r;
    }
    /* BeforeNavigate2: false when the program cancelled it (or the control is gone) */
    bool before(const QString &given, long flags, bool *gone)
    {
        vbvar a[7] = { vb_v_obj(c->self), vstr(given), vb_v_l((vb_i32)flags), vstr(QString()), vb_v_empty(),
                       vb_v_empty(), vbool(false) };
        *gone = !fire("beforenavigate2", 7, a);
        if (*gone) return false;
        bool cancel = vb_v2l(a[6]) != 0;
        for (int i = 1; i < 7; i++) vb_vfree(&a[i]);   /* not pDisp: the control itself, not a reference of its own */
        return !cancel;
    }
    bool states() { return ev_csc(2, !back.empty()) && ev_csc(1, !fwd.empty()); }

    /* ------------------------------------------------ the address */
    static QUrl to_url(const QString &s, const QUrl &base)
    {
        QString t = s.trimmed();
        if (t.isEmpty()) return QUrl("about:blank");
        if (t.startsWith("\\\\")) {             /* \\server\share\file */
#ifdef _WIN32
            return QUrl::fromLocalFile(t);
#else
            QString r = t.mid(2).replace('\\', '/');
            int sl = r.indexOf('/');
            QUrl u;
            u.setScheme("file");
            u.setHost(sl < 0 ? r : r.left(sl));
            u.setPath(sl < 0 ? QString("/") : r.mid(sl));
            return u;
#endif
        }
        if (t.size() > 2 && t[1] == ':' && (t[2] == '\\' || t[2] == '/')) return QUrl::fromLocalFile(t);   /* C:\ */
#ifndef _WIN32
        if (t.startsWith('/')) return QUrl::fromLocalFile(QString(t).replace('\\', '/'));   /* App.Path & "\x.html" */
#else
        if (t.startsWith('/')) return QUrl::fromLocalFile(t);
#endif
        int i = t.indexOf(':');
        if (i > 1 && !t.left(i).contains('/') && !t.left(i).contains('\\')) {     /* http:, about: ... */
            QUrl u(t);
            if ((u.scheme() == "http" || u.scheme() == "https") && u.path().isEmpty()) u.setPath("/");  /* IE: http://host/ */
            return u;
        }
        if (base.isValid() && !base.isEmpty()) return base.resolved(QUrl(QString(t).replace('\\', '/')));
        if (t.startsWith("www.", Qt::CaseInsensitive)) return to_url("http://" + t, QUrl());
        return QUrl::fromLocalFile(QDir::current().absoluteFilePath(QString(t).replace('\\', '/')));
    }
    static QString url_text(const QUrl &u)
    {
        if (u.scheme() == "about") return u.toString();
        return u.toString(QUrl::FullyEncoded);
    }
    /* the <title> of the page itself, not one of an <svg> in the body (QTextBrowser takes the last <title> it meets:
     * microsoft.com gave the title of a privacy icon, aw26) */
    static QString head_title(const QString &page)
    {
        int body = page.indexOf(QRegularExpression("<body[\\s>]", QRegularExpression::CaseInsensitiveOption));
        QString head = body < 0 ? page : page.left(body);
        QRegularExpression re("<title[^>]*>(.*?)</title>",
                              QRegularExpression::CaseInsensitiveOption | QRegularExpression::DotMatchesEverythingOption);
        auto m = re.match(head);
        if (!m.hasMatch()) return QString();
        QTextDocument d;
        d.setHtml(m.captured(1));                 /* the entities: &amp; &#8211; ... */
        return d.toPlainText().simplified();
    }
    static QString name_of(const QUrl &u, const QString &title)
    {
        if (u.scheme() == "about") return u.toString();
        if (u.scheme() == "file") {             /* VB6: the name of the file (oracle wb1) */
            QString n = QFileInfo(u.path()).fileName();
            return n.isEmpty() ? u.toString() : n;
        }
        return title.isEmpty() ? u.toString() : title;
    }
    static QString decode(const QByteArray &b)
    {
        if (b.startsWith("\xEF\xBB\xBF")) return QString::fromUtf8(b.mid(3));
        if (b.startsWith("\xFF\xFE")) return QString::fromUtf16(reinterpret_cast<const char16_t *>(b.constData() + 2), (b.size() - 2) / 2);
        QString u = QString::fromUtf8(b);
        if (!u.contains(QChar(0xFFFD))) return u;
        return QString::fromLatin1(b);
    }
    static QString error_page(const QString &what)
    {
        QString w = what.toHtmlEscaped();
        return QString("<html><head><title>This page can\u2019t be displayed</title></head><body>"
                       "<h2>This page can\u2019t be displayed</h2><p>%1</p></body></html>").arg(w);
    }

    /* ------------------------------------------------ navigating */
    /* mode 0: a new page, 1: back, 2: forward */
    void navigate(const QString &given, long flags, int mode)
    {
        QUrl u = to_url(given, cur.url);
        bool gone = false;
        if (!before(given, flags, &gone)) {
            if (!gone) states();                /* Cancel: VB6 tells the commands again (oracle wb1) */
            return;
        }
        stop_job();
        int g = ++gen;
        rs = 1;                                 /* VB6: Busy False, ReadyState 1 when Navigate returns (oracle f_web) */
        std::shared_ptr<bool> a = alive;
        QTimer::singleShot(0, v ? (QObject *)v.data() : qApp, [this, a, g, given, u, mode]() {
            if (!*a || g != gen) return;
            busy = true;
            if (!fire("downloadbegin") || g != gen) return;
            if (u.scheme() == "http" || u.scheme() == "https") { fetch(g, given, u, mode, 0); return; }
            if (u.scheme() == "about") { show(g, given, u, mode, QString(), true); return; }
            if (u.isLocalFile()) {
                QFile f(u.toLocalFile());
                if (f.open(QIODevice::ReadOnly)) { show(g, given, u, mode, decode(f.readAll()), false); return; }
            }
            show(g, given, u, mode, error_page(given), false);  /* as IE: no NavigateError for a file (oracle wb1) */
        });
    }
    void stop_job()
    {
        if (job) { lbinet_release(job); job = nullptr; }
        if (poll) poll->stop();
    }
    void fetch(int g, const QString &given, const QUrl &u, int mode, int hops)
    {
        int https = u.scheme() == "https";
        QString path = u.path(QUrl::FullyEncoded);
        if (path.isEmpty()) path = "/";
        if (u.hasQuery()) path += "?" + u.query(QUrl::FullyEncoded);
        job = lbinet_http(nullptr, https, 1, "GET", u.host().toUtf8().constData(), u.port(https ? 443 : 80),
                          path.toUtf8().constData(), u.userName().toUtf8().constData(), u.password().toUtf8().constData(),
                          "", nullptr, 0, 60, 0, "");
        if (!poll) poll = new QTimer();
        QObject::disconnect(poll, nullptr, nullptr, nullptr);
        std::shared_ptr<bool> a = alive;
        QObject::connect(poll, &QTimer::timeout, [this, a, g, given, u, mode, hops]() {
            if (!*a) return;
            if (g != gen || !job) { if (poll) poll->stop(); return; }
            for (int e; (e = lbinet_event(job)) >= 0;) {}
            if (!lbinet_done(job)) return;
            poll->stop();
            long n = 0;
            const char *b = lbinet_body(job, &n);
            QByteArray body(b ? b : "", (int)n);
            QByteArray hdr(lbinet_headers(job));
            int code = lbinet_code(job), ok = lbinet_ok(job);
            lbinet_release(job);
            job = nullptr;
            if (ok && code >= 300 && code < 400 && hops < 5) {      /* a redirection: Location */
                for (const QByteArray &l : hdr.split('\n')) {
                    if (l.toLower().startsWith("location:")) {
                        QUrl to = u.resolved(QUrl(QString::fromLatin1(l.mid(9)).trimmed()));
                        if (to.scheme() == "http" || to.scheme() == "https") { fetch(g, given, to, mode, hops + 1); return; }
                    }
                }
            }
            if (!ok || code >= 400) {
                vbvar ea[5] = { vb_v_obj(c->self), vstr(given), vstr(QString()), vb_v_l(ok ? code : 12007), vbool(false) };
                if (!fire("navigateerror", 5, ea) || g != gen) return;
                bool cancel = vb_v2l(ea[4]) != 0;
                for (int i = 1; i < 5; i++) vb_vfree(&ea[i]);
                if (cancel) { busy = false; rs = 4; ev_csc(-1, false); return; }
                show(g, given, u, mode, error_page(given), false);
                return;
            }
            show(g, given, u, mode, decode(body), false);
        });
        poll->start(10);
    }
    /* the page has come: the events of VB6 / IE, in their order (oracle wb1) */
    void show(int g, const QString &given, const QUrl &u, int mode, const QString &page, bool blank)
    {
        if (mode == 1) { fwd.push_back(cur); cur = back.back(); back.pop_back(); }
        else if (mode == 2) { back.push_back(cur); cur = fwd.back(); fwd.pop_back(); }
        else { if (!cur.url.isEmpty()) back.push_back(cur); fwd.clear(); cur = Entry{given, u}; }
        cur.url = u;
        name = name_of(u, QString());
        if (!ev_s("statustextchange", QString()) || g != gen) return;
        if (!states() || g != gen) return;
        if (!ev_csc(-1, false) || g != gen) return;
        if (!ev_s("titlechange", name) || g != gen) return;
        if (!ev_url("navigatecomplete2", given) || g != gen) return;
        status = "Done";
        if (!ev_s("statustextchange", status) || g != gen) return;
        if (!blank && (!ev_progress(10000, 10000) || g != gen)) return;
        if (!ev_progress(-1, blank ? 0 : 10000) || g != gen) return;
        html = page;
        if (v) {
            v->base = u;
            if (blank) v->clear(); else v->setHtml(page);
        }
        QString title = blank ? QString() : head_title(page);
        if (title.isEmpty() && v && !blank) title = v->documentTitle();
        name = name_of(u, title);
        has_doc = true;
        busy = false;
        rs = 3;
        if (!fire("downloadcomplete") || g != gen) return;
        if (!ev_s("titlechange", blank ? u.toString() : title.isEmpty() ? name : title) || g != gen) return;
        if (!blank && (!ev_progress(10000, 10000) || g != gen)) return;
        rs = 4;                                 /* about:blank too (oracle f_web) */
        if (!ev_url("documentcomplete", given) || g != gen) return;
        if (!ev_csc(-1, false) || g != gen) return;
        ev_progress(0, 0);
    }
    vbvar failed(const char *method)
    {
        vb_raise_full((vb_i32)0x80004005, vs(QString("WebBrowser")),
                      vs(QString("Method '%1' of object 'IWebBrowser2' failed").arg(method)));
        return vb_v_empty();
    }

    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        auto flag = [&](bool &f) -> vbvar {
            if (!set) return vbool(f);
            f = vb_v2l(val) != 0;
            return vb_v_empty();
        };
        if (set && (n == "locationurl" || n == "locationname" || n == "busy" || n == "readystate" || n == "type" ||
                    n == "document" || n == "hwnd")) {
            vb_raise(383);                      /* read-only */
            return vb_v_empty();
        }
        if (n == "navigate" || n == "navigate2") {
            if (argc < 1) { vb_raise(449); return vb_v_empty(); }
            QString u = qs(vb_v2s(argv[0]));
            if (vb_errpend) return vb_v_empty();
            navigate(u, given(argc, argv, 1) ? arg_l(argc, argv, 1) : 256, 0);
            return vb_v_empty();
        }
        if (n == "goback" || n == "goforward") {
            bool b = n == "goback";
            if (b ? back.empty() : fwd.empty()) {
                states();                       /* VB6 tells the commands, then fails (oracle wb1 step 4) */
                return failed(b ? "GoBack" : "GoForward");
            }
            navigate((b ? back.back() : fwd.back()).given, 320, b ? 1 : 2);
            return vb_v_empty();
        }
        if (n == "gohome" || n == "gosearch") { navigate("about:blank", 256, 0); return vb_v_empty(); }
        if (n == "refresh" || n == "refresh2") {
            if (v && has_doc && !html.isEmpty()) {
                if (cur.url.isLocalFile()) {    /* a file: as it is now */
                    QFile f(cur.url.toLocalFile());
                    if (f.open(QIODevice::ReadOnly)) html = decode(f.readAll());
                }
                v->setHtml(html);
            }
            ev_csc(-1, false);                  /* VB6: only this (oracle wb1 step 7) */
            return vb_v_empty();
        }
        if (n == "stop") {
            if (job || busy) { ++gen; stop_job(); busy = false; rs = has_doc ? 4 : 1; }
            ev_csc(-1, false);                  /* oracle wb1 step 9 */
            return vb_v_empty();
        }
        if (n == "locationurl") return vstr(cur.url.isEmpty() ? QString() : url_text(cur.url));
        if (n == "locationname") return vstr(cur.url.isEmpty() ? QString() : name);
        if (n == "busy") return vbool(busy);
        if (n == "readystate") return vb_v_l(rs);
        if (n == "type") { if (!has_doc) return failed("Type"); return vstr(QString("HTML Document")); }
        if (n == "document" || n == "application" || n == "container" || n == "parent") return vb_v_obj((vbobj)0);
        if (n == "toplevelcontainer") return vbool(false);
        if (n == "hwnd") return vb_v_l(0);
        if (n == "statustext") {
            if (!set) return vstr(status);
            status = qs(vb_v2s(val));
            return vb_v_empty();
        }
        if (n == "offline") return flag(offline);
        if (n == "silent") return flag(silent);
        if (n == "registerasbrowser") return flag(regbrowser);
        if (n == "registerasdroptarget") return flag(regdrop);
        if (n == "addressbar") return flag(addressbar);
        if (n == "menubar") return flag(menubar);
        if (n == "statusbar") return flag(statusbar);
        if (n == "fullscreen") return flag(fullscreen);
        if (n == "theatermode") return flag(theater);
        if (n == "resizable") return flag(resizable);
        if (n == "toolbar") {
            if (!set) return vb_v_l((vb_i32)toolbar);
            toolbar = vb_v2l(val);
            return vb_v_empty();
        }
        if (n == "putproperty") {
            if (argc < 2) { vb_raise(449); return vb_v_empty(); }
            props[qs(vb_v2s(argv[0])).toLower()] = qs(vb_v2s(argv[1]));
            return vb_v_empty();
        }
        if (n == "getproperty") {
            if (argc < 1) { vb_raise(449); return vb_v_empty(); }
            auto it = props.find(qs(vb_v2s(argv[0])).toLower());
            return it == props.end() ? vb_v_empty() : vstr(it->second);
        }
        if (n == "execwb" || n == "showbrowserbar" || n == "quit" || n == "clienttowindow") return vb_v_empty();
        if (n == "querystatuswb") return vb_v_l(0);
        *found = 0;
        return vb_v_empty();
    }
};
static Ocx *new_webbrowser() { return new OcxWebBrowser; }
