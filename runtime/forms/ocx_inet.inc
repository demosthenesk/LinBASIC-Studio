/* ocx_inet.inc - the Internet Transfer control (MSINET.OCX, InetCtlsObjects.Inet): OpenURL, Execute with
 * HTTP(S) and FTP, GetChunk, GetHeader, StateChanged, the URL and its parts, as VB6 (oracle tests/dev
 * f_inet, f_inetftp). The requests run in threads of runtime/inet.c; the control reads their states with a
 * timer, so StateChanged comes from the message loop, as in VB6.
 * Not as VB6, on purpose (bugs of MSINET.OCX, oracle f_inet / f_inetftp):
 *   - an FTP command on an HTTP URL (Execute , "CD sub") is an error, but the control does not stay busy for
 *     ever (VB6: StillExecuting stays True, even Cancel does not help; every Execute after it is 35764);
 *   - an HTTP verb other than GET, HEAD, POST, PUT (DELETE, OPTIONS ...) is sent as it is (VB6 sends nothing
 *     and stays busy until Cancel);
 *   - GetChunk after a cancelled request returns at once (VB6: it waits for ever);
 *   - OpenURL of an ftp:// URL returns the file (VB6: it waits for RequestTimeout and fails).
 * Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <deque>
#include <memory>
#include <QThread>
#include <QElapsedTimer>
#include <QCoreApplication>
extern "C" {
void *lbinet_conn_new(void);
void lbinet_conn_free(void *c);
void *lbinet_http(void *conn, int https, int openurl, const char *method, const char *host, int port, const char *path,
                  const char *user, const char *pass, const char *hdrs, const char *data, int ndata, int timeout,
                  int access, const char *proxy);
void *lbinet_ftp(void *session, const char *op, const char *a1, const char *a2, int timeout);
void *lbinet_ftpurl(const char *host, int port, const char *path, const char *user, const char *pass, int timeout);
int lbinet_event(void *r);
int lbinet_done(void *r);
int lbinet_stage(void *r);
int lbinet_ok(void *r);
int lbinet_code(void *r);
const char *lbinet_info(void *r);
const char *lbinet_body(void *r, long *n);
const char *lbinet_headers(void *r);
void *lbinet_take_conn(void *r);
void *lbinet_take_ftp(void *r);
void lbinet_release(void *r);
void *lbinet_ftp_new(void);
void lbinet_ftp_free(void *f);
int lbinet_ftp_open(void *f, const char *host, int port, const char *user, const char *pass, int timeout);
int lbinet_ftp_isopen(void *f);
int lbinet_ftp_same(void *f, const char *host, int port, const char *user);
void lbinet_ftp_close(void *f);
}

static const char *inet_errtext(int n)
{
    switch (n) {
    case 5: return "Invalid procedure call or argument";
    case 13: return "Type mismatch";
    case 380: return "Invalid property value";
    case 35750: return "Unable to open internet handle";
    case 35751: return "Unable to open URL";
    case 35752: return "URL is malformed";
    case 35753: return "Protocol not supported for this method";
    case 35754: return "Unable to connect to remote host";
    case 35755: return "No remote computer is specified";
    case 35756: return "Unable to complete request";
    case 35757: return "You must execute an operation before retrieving data";
    case 35758: return "Unable to retrieve data";
    case 35759: return "FTP command failed";
    case 35760: return "Cannot coerce type";
    case 35761: return "Request timed out";
    case 35762: return "Not a valid or supported command";
    case 35763: return "Invalid operation argument";
    case 35764: return "Still executing last request";
    case 35765: return "This call is not valid for an FTP connection";
    }
    return "Unable to complete request";
}

struct OcxInet : Ocx {
    enum { UNKNOWN, DEFAULT, FTP, GOPHER, HTTP, HTTPS };
    int proto = DEFAULT, access = 0;
    long timeout = 60;
    quint16 port = 80;
    QString host, doc, user, pass, proxy, scheme0;
    void *conn = nullptr, *ftp = nullptr, *job = nullptr;
    bool job_openurl = false, job_ftpurl = false;
    bool executing = false, any = false, failed = false, busy = false;
    long code = 0;
    QString info;
    QByteArray data, headers;
    int pos = 0;
    std::deque<int> pending;                    /* states the control itself gives (3 4 of an FTP login, 9 10 of Cancel) */
    QTimer *timer = nullptr;
    std::shared_ptr<bool> alive = std::make_shared<bool>(true);

    ~OcxInet() override
    {
        *alive = false;
        if (job) lbinet_release(job);
        lbinet_conn_free(conn);
        lbinet_ftp_free(ftp);
        if (timer) { timer->stop(); timer->deleteLater(); }
    }
    QWidget *make(QWidget *parent) override { (void)parent; return nullptr; }

    /* ------------------------------------------------ the URL and its parts */
    static quint16 defport(int p) { return p == FTP ? 21 : p == GOPHER ? 0 : p == HTTPS ? 443 : 80; }
    QString scheme() const
    {
        switch (proto) {
        case FTP: return "ftp";
        case GOPHER: return "gopher";
        case HTTPS: return "https";
        case UNKNOWN: return scheme0;
        default: return "http";
        }
    }
    QString url() const
    {
        if (host.isEmpty()) return QString();
        QString u = scheme() + "://";
        if (!user.isEmpty()) u += user + (pass.isEmpty() ? QString() : ":" + pass) + "@";
        u += host;
        if (port != defport(proto)) u += ":" + QString::number(port);
        if (!doc.isEmpty() && !doc.startsWith('/')) u += '/';
        return u + doc;
    }
    /* VB6 (oracle f_inet): a URL without a scheme keeps the Protocol; an empty one keeps Protocol and port */
    void set_url(const QString &u)
    {
        if (u.isEmpty()) { host.clear(); doc.clear(); user.clear(); pass.clear(); return; }
        QString rest = u;
        int i = u.indexOf("://");
        if (i > 0) {
            QString sc = u.left(i).toLower();
            proto = sc == "http" ? HTTP : sc == "https" ? HTTPS : sc == "ftp" ? FTP : sc == "gopher" ? GOPHER : UNKNOWN;
            scheme0 = proto == UNKNOWN ? u.left(i) : QString();
            rest = u.mid(i + 3);
        }
        port = defport(proto);
        int sl = rest.indexOf('/');
        QString auth = sl < 0 ? rest : rest.left(sl);
        doc = sl < 0 ? QString() : rest.mid(sl);
        int at = auth.lastIndexOf('@');
        user.clear(); pass.clear();
        if (at >= 0) {
            QString ui = auth.left(at);
            auth = auth.mid(at + 1);
            int c = ui.indexOf(':');
            user = c < 0 ? ui : ui.left(c);
            pass = c < 0 ? QString() : ui.mid(c + 1);
        }
        int c = auth.lastIndexOf(':');
        if (c >= 0 && auth.indexOf(']', c) < 0) { port = (quint16)auth.mid(c + 1).toInt(); host = auth.left(c); }
        else host = auth;
    }
    void load(const char *const *p) override
    {
        access = (int)propn(p, "AccessType", 0);
        proto = (int)propn(p, "Protocol", DEFAULT);
        port = defport(proto);
        port = (quint16)propn(p, "RemotePort", port);
        QString s;
        if (prop(p, "URL", &s)) set_url(s);
        if (prop(p, "RemoteHost", &s)) host = s;
        if (prop(p, "Document", &s)) doc = s;
        if (prop(p, "UserName", &s)) user = s;
        if (prop(p, "Password", &s)) pass = s;
        prop(p, "Proxy", &proxy);
        timeout = (long)propn(p, "RequestTimeout", 60);
    }

    vbvar err(int n) { vb_raise_full(n, vs(QString("Inet")), vs(QString(inet_errtext(n)))); return vb_v_empty(); }

    /* ------------------------------------------------ the requests */
    void run()
    {
        if (!timer) {
            timer = new QTimer();
            QObject::connect(timer, &QTimer::timeout, [this]() { if (!busy) tick(); });
        }
        if (!timer->isActive()) timer->start(10);
    }
    /* the end of a request: what it gave, its connection back to the control */
    void finish()
    {
        if (!job) return;
        while (!lbinet_done(job)) QThread::msleep(1);       /* its last states are on their way */
        for (int e; (e = lbinet_event(job)) >= 0;) if (e != 11 && e != 12) pending.push_back(e);
        bool ok = lbinet_ok(job) != 0;
        long n = 0;
        const char *b = lbinet_body(job, &n);
        bool http = !job_ftpurl && !ftp_job;
        if (!(job_openurl && http)) {               /* OpenURL of HTTP does not set them (oracle) */
            code = lbinet_code(job);
            info = QString::fromLocal8Bit(lbinet_info(job));
        }
        data = QByteArray(b, (int)n);
        pos = 0;
        headers = http ? QByteArray(lbinet_headers(job)) : QByteArray();
        failed = !ok;
        if (http) { void *c = lbinet_take_conn(job); if (c) { lbinet_conn_free(conn); conn = c; } }
        if (ftp_job) { void *f = lbinet_take_ftp(job); if (f) { lbinet_ftp_free(ftp); ftp = f; } }
        lbinet_release(job);
        job = nullptr;
        ftp_job = false;
        executing = false;
    }
    bool ftp_job = false;
    /* the next state: the control's own first, then the request's; -1 none now */
    int next_state()
    {
        if (!pending.empty()) { int e = pending.front(); pending.pop_front(); return e; }
        if (!job) return -1;
        int e = lbinet_event(job);
        if (e == 11 || e == 12) { finish(); if (job_openurl && e == 12) return -2; }   /* OpenURL: no 12 (oracle) */
        return e;
    }
    /* StateChanged for what has come; false: the control is gone */
    bool pump()
    {
        for (;;) {
            int e = next_state();
            if (e == -2) continue;
            if (e < 0) return true;
            std::shared_ptr<bool> a = alive;
            vbvar arg = vb_v_i((vb_i16)e);
            fire_ctl(c, "statechanged", 1, &arg);
            if (!*a) return false;
        }
    }
    void tick()
    {
        busy = true;
        std::shared_ptr<bool> a = alive;
        bool live = pump();
        if (!live || !*a) return;
        busy = false;
        if (!job && pending.empty() && timer) timer->stop();
    }
    /* the states that have come, now (Cancel, CLOSE) */
    void fire_now()
    {
        if (busy) return;
        busy = true;
        std::shared_ptr<bool> a = alive;
        if (!pump() || !*a) return;
        busy = false;
        if (!job && pending.empty() && timer) timer->stop();
    }
    void start_http(int openurl, const QString &method, const QByteArray &hdrs, const QByteArray &body)
    {
        QByteArray d = doc.isEmpty() ? QByteArray("/") : doc.toUtf8();
        job = lbinet_http(conn, proto == HTTPS, openurl, method.toLatin1().constData(), host.toUtf8().constData(), port,
                          d.constData(), user.toUtf8().constData(), pass.toUtf8().constData(), hdrs.constData(),
                          body.constData(), body.size(), (int)timeout, access, proxy.toUtf8().constData());
        conn = nullptr;
        job_openurl = openurl != 0;
        job_ftpurl = false;
        ftp_job = false;
    }
    static QByteArray ansi(vbvar v)
    {
        QByteArray b;
        if (v.vt == (VT_ARRAY | VT_UI1)) {
            vbarr *a = vb_v_arrptr(v);
            if (a && a->nd) b = QByteArray((const char *)a->data, (int)a->n);
            return b;
        }
        vbs s = vb_v2s(v);
        for (vb_i32 i = 0; i < s.len; i++) { int ch = vb_w2ansi(s.p[i]); b.append((char)(ch < 0 ? '?' : ch)); }
        return b;
    }
    static vbvar as_value(const QByteArray &b, bool bytes)
    {
        if (bytes) {
            vbarr *a = vb_atmp();
            if (b.isEmpty()) vb_aempty(a, 1, VBA_PLAIN);
            else {
                vb_adim(a, 1, VBA_PLAIN, 0, 1, 0, (vb_i32)b.size() - 1);
                if (vb_errpend) return vb_v_empty();
                memcpy(a->data, b.constData(), (size_t)b.size());
            }
            return vb_v_arr(a, VT_UI1);
        }
        vbs t = vb_stmp((vb_i32)b.size());
        for (int i = 0; i < b.size(); i++) t.p[i] = vb_ansi2w((unsigned char)b[i]);
        return vb_v_s(t);
    }

    vbvar openurl(vb_i32 argc, vbvar *argv)
    {
        if (executing) return err(35764);
        QString u = given(argc, argv, 0) ? qs(vb_v2s(argv[0])) : QString();
        if (vb_errpend) return vb_v_empty();
        long dt = given(argc, argv, 1) ? vb_v2l(argv[1]) : 0;
        if (vb_errpend) return vb_v_empty();
        if (dt != 0 && dt != 1) { vb_raise(5); return vb_v_empty(); }
        if (!u.isEmpty()) set_url(u);
        if (host.isEmpty()) return err(35752);
        any = true;
        if (proto == HTTP || proto == HTTPS) start_http(1, "GET", QByteArray(), QByteArray());
        else if (proto == FTP) {
            job = lbinet_ftpurl(host.toUtf8().constData(), port, doc.toUtf8().constData(), user.toUtf8().constData(),
                                pass.toUtf8().constData(), (int)timeout);
            job_openurl = true; job_ftpurl = true; ftp_job = false;
        } else {                                    /* gopher, an unknown scheme: State 11, "" (oracle) */
            pending.push_back(11);
            busy = true; pump(); busy = false;
            return as_value(QByteArray(), dt == 1);
        }
        executing = true;
        bool ftpurl = job_ftpurl;
        QElapsedTimer clock; clock.start();
        busy = true;
        std::shared_ptr<bool> a = alive;
        while (job) {
            if (!pump() || !*a) return vb_v_empty();
            if (!job) break;
            if (timeout > 0 && clock.elapsed() > timeout * 1000) {
                lbinet_release(job); job = nullptr; executing = false;
                if (ftp_job) ftp = nullptr;
                busy = false;
                return err(35761);
            }
            QCoreApplication::processEvents(QEventLoop::AllEvents, 10);
            if (job) QThread::msleep(2);
        }
        if (!pump() || !*a) return vb_v_empty();
        busy = false;
        executing = false;
        if (failed) return ftpurl ? vb_v_empty() : as_value(QByteArray(), dt == 1);
        if (dt == 0 && data.contains('\0')) { vb_raise(13); return vb_v_empty(); }   /* binary data as a String (oracle) */
        return as_value(data, dt == 1);
    }

    vbvar execute(vb_i32 argc, vbvar *argv)
    {
        if (executing) return err(35764);
        QString u = given(argc, argv, 0) ? qs(vb_v2s(argv[0])) : QString();
        QString op = given(argc, argv, 1) ? qs(vb_v2s(argv[1])).trimmed() : QString();
        QByteArray body = given(argc, argv, 2) ? ansi(argv[2]) : QByteArray();
        QByteArray hdrs = given(argc, argv, 3) ? ansi(argv[3]) : QByteArray();
        if (vb_errpend) return vb_v_empty();
        if (!u.isEmpty()) set_url(u);
        if (host.isEmpty()) return err(35752);
        QString verb = op.section(' ', 0, 0, QString::SectionSkipEmpty).toUpper();
        QString rest = op.section(' ', 1, -1, QString::SectionSkipEmpty);
        static const QStringList ftpverbs = { "CD", "CDUP", "CLOSE", "DELETE", "DIR", "LS", "MKDIR", "PWD", "QUIT",
                                              "RECV", "RENAME", "RMDIR", "SEND", "SIZE" };
        any = true;
        if (proto == HTTP || proto == HTTPS) {
            if (verb.isEmpty()) verb = "GET";
            if (ftpverbs.contains(verb) && verb != "DELETE") {   /* an FTP command on HTTP: VB6's error (not busy after it) */
                code = 87; info = "The parameter is incorrect.\r\n";
                return err(35756);
            }
            bool token = true;                      /* GET, HEAD, POST, PUT, and DELETE, OPTIONS ... as they are */
            for (QChar ch : verb) if (!ch.isLetter() && ch != '-' && ch != '_') token = false;
            if (!token) return err(35762);
            code = 0; info.clear(); data.clear(); headers.clear(); pos = 0;
            start_http(0, verb, hdrs, body);
            executing = true;
            /* VB6 has sent the request when Execute returns (a Cancel right after it: 5 6 7 9 10, oracle) */
            QElapsedTimer w; w.start();
            while (lbinet_stage(job) < 7 && w.elapsed() < 500) QThread::msleep(1);
            run();
            return vb_v_empty();
        }
        if (proto != FTP) return err(35753);
        /* FTP: the arguments; local file names as the program gives them */
        QStringList w = rest.split(' ', Qt::SkipEmptyParts);
        QString a1 = w.value(0), a2 = w.value(1);
        if ((verb == "GET" || verb == "RECV" || verb == "PUT" || verb == "SEND" || verb == "RENAME") && w.size() < 2) {
            if (verb == "GET" || verb == "RECV" || verb == "PUT" || verb == "SEND") return vb_v_empty();   /* oracle: nothing */
        }
        if (!ftp) ftp = lbinet_ftp_new();
        QByteArray h8 = host.toUtf8(), u8 = user.toUtf8(), p8 = pass.toUtf8();
        if (!lbinet_ftp_same(ftp, h8.constData(), port, u8.constData())) {
            lbinet_ftp_close(ftp);
            if (lbinet_ftp_open(ftp, h8.constData(), port, u8.constData(), p8.constData(), (int)timeout)) return err(35754);
            pending.push_back(3); pending.push_back(4);
        }
        code = 0; info.clear(); data.clear(); headers.clear(); pos = 0;
        if (verb == "CLOSE" || verb == "QUIT") {
            lbinet_ftp_close(ftp);
            pending.push_back(9); pending.push_back(10);
            fire_now();                             /* VB6: StateChanged 9 10 before Execute returns (oracle) */
            return vb_v_empty();
        }
        QByteArray l1 = a1.toUtf8(), l2 = a2.toUtf8();
        if (verb == "GET" || verb == "RECV") l2 = native(a2).toUtf8();
        if (verb == "PUT" || verb == "SEND") l1 = native(a1).toUtf8();
        job = lbinet_ftp(ftp, verb.toLatin1().constData(), l1.constData(), l2.constData(), (int)timeout);
        ftp = nullptr;
        ftp_job = true; job_openurl = false; job_ftpurl = false;
        executing = true;
        run();
        return vb_v_empty();
    }

    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        auto str = [&](QString &f) -> vbvar {
            if (!set) return vstr(f);
            f = qs(vb_v2s(val));
            return vb_v_empty();
        };
        if (set && (n == "responsecode" || n == "responseinfo" || n == "stillexecuting" || n == "hinternet")) {
            vb_raise(450);
            return vb_v_empty();
        }
        if (n == "url" || n == "_url") {
            if (!set) return vstr(url());
            set_url(qs(vb_v2s(val)));
            return vb_v_empty();
        }
        if (n == "protocol") {
            if (!set) return vb_v_l(proto);
            long v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < 0 || v > 5) return err(380);
            proto = (int)v;
            port = defport(proto);
            return vb_v_empty();
        }
        if (n == "remotehost") return str(host);
        if (n == "remoteport") {
            if (!set) return vb_v_i((vb_i16)port);
            long v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < -32768 || v > 65535) { vb_raise(6); return vb_v_empty(); }
            port = (quint16)v;
            return vb_v_empty();
        }
        if (n == "document") return str(doc);
        if (n == "username") return str(user);
        if (n == "password") return str(pass);
        if (n == "proxy") return str(proxy);
        if (n == "accesstype") {
            if (!set) return vb_v_l(access);
            long v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < 0 || v > 2) return err(380);
            access = (int)v;
            return vb_v_empty();
        }
        if (n == "requesttimeout") {
            if (!set) return vb_v_l((vb_i32)timeout);
            long v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < 0) return err(380);
            timeout = v;
            return vb_v_empty();
        }
        if (n == "responsecode") return vb_v_l((vb_i32)code);
        if (n == "responseinfo") return vstr(info);
        if (n == "stillexecuting") return vbool(executing);
        if (n == "hinternet") return vb_v_l(any ? 0x00CC0004 : 0);
        if (n == "openurl") return openurl(argc, argv);
        if (n == "execute") return execute(argc, argv);
        if (n == "cancel") {
            if (job) {
                for (int e; (e = lbinet_event(job)) >= 0;) if (e != 11 && e != 12) pending.push_back(e);
                lbinet_release(job);                /* its connection goes with it */
                job = nullptr;
                ftp_job = false;
                pending.push_back(9); pending.push_back(10);
            }
            executing = false;
            fire_now();                             /* VB6: the states up to 9 10 before Cancel returns (oracle) */
            return vb_v_empty();
        }
        if (n == "getchunk") {
            if (argc < 1) { vb_raise(449); return vb_v_empty(); }
            long size = vb_v2l(argv[0]);
            long dt = given(argc, argv, 1) ? vb_v2l(argv[1]) : 0;
            if (vb_errpend) return vb_v_empty();
            if (dt != 0 && dt != 1) { vb_raise(5); return vb_v_empty(); }
            long k = executing ? 0 : std::max(0L, std::min(size, (long)(data.size() - pos)));
            QByteArray part = data.mid(pos, (int)k);
            pos += (int)k;
            vb_vset(&argv[0], vb_v_l((vb_i32)k));    /* Size: the bytes it gave (ByRef, oracle) */
            return as_value(part, dt == 1);
        }
        if (n == "getheader") {
            if (!any) return err(35756);
            QString name = given(argc, argv, 0) ? qs(vb_v2s(argv[0])) : QString();
            if (failed) return vstr(name.isEmpty() ? QString("\r\n") : QString());
            if (name.isEmpty()) return vstr(QString::fromLatin1(headers));
            for (const QByteArray &l : headers.split('\n')) {
                int c = l.indexOf(':');
                if (c > 0 && QString::fromLatin1(l.left(c)).trimmed().compare(name, Qt::CaseInsensitive) == 0)
                    return vstr(QString::fromLatin1(l.mid(c + 1)).trimmed());
            }
            return vstr(QString());
        }
        if (n == "aboutbox") return vb_v_empty();
        *found = 0;
        return vb_v_empty();
    }
};
static Ocx *new_inet() { return new OcxInet; }
