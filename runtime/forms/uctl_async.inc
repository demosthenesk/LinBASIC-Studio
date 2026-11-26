/* uctl_async.inc - UserControl.AsyncRead / CancelAsyncRead, the AsyncReadProgress and AsyncReadComplete events,
 * the AsyncProperty object, UserControl.Hyperlink (10.4), as VB6 has them (oracle tests/dev/f_inetasync, DEV-10):
 *  - a local file (a path or file://) is read during the call: the progress events 11 (SendingRequest), 13 (the
 *    MIME type), 14 (the file), 6 (EndDownloadData), then AsyncReadComplete, all before AsyncRead returns; a file
 *    that is not there: 11, then error 53;
 *  - HTTP (runtime/inet.c, in a thread): the events come from the message loop after AsyncRead returns: 1 and 2
 *    (FindingResource, Connecting) the first time a server / its port is used, 11 when the request is sent,
 *    3 and 11 again after a redirect, 13, 14 (the copy in the cache), 6, AsyncReadComplete; no events between
 *    (DownloadingData does not come in VB6 either); an HTTP error (404, 500 ...): 11 and AsyncReadComplete with
 *    StatusCode 0, its Value an error (4xx &H80070006, the others &H80070008); a server that refuses: 2 twice,
 *    AsyncReadComplete with the server in Status and the WinINet error in Value (&H80072EFD);
 *  - vbAsyncReadSynchronousDownload: everything during the call;
 *  - the Value of vbAsyncTypeFile is a copy in the cache, also for a local file; vbAsyncTypePicture a StdPicture;
 *    vbAsyncTypeByteArray a Byte array;
 *  - the same PropertyName while a read of it runs: 689; an empty Target: 693; an AsyncType other than 0-2 and
 *    CancelAsyncRead of a name that is not being read: the number &H8007000E with the text of E_INVALIDARG (a bug
 *    of VB6, kept: programs may test the number); a cancelled read sends no more events;
 *  - Hyperlink: NavigateTo opens the target with the program of the system (VB6 in a form: the default browser),
 *    an empty target E_INVALIDARG; GoBack and GoForward E_NOINTERFACE (a form is no hyperlink container).
 * Included by vbforms.cpp after uctl.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <QStandardPaths>
#include <QThread>
#include <QUrl>
#include <QDesktopServices>
#include <memory>
extern "C" {
void *lbinet_http(void *conn, int https, int openurl, const char *method, const char *host, int port, const char *path,
                  const char *user, const char *pass, const char *hdrs, const char *data, int ndata, int timeout,
                  int access, const char *proxy);
void *lbinet_ftpurl(const char *host, int port, const char *path, const char *user, const char *pass, int timeout);
int lbinet_event(void *r);
int lbinet_done(void *r);
int lbinet_ok(void *r);
int lbinet_code(void *r);
const char *lbinet_body(void *r, long *n);
const char *lbinet_headers(void *r);
const char *lbinet_final(void *r, int *redirects);
void lbinet_release(void *r);
}

/* ---------------------------------------------------------------- the errors with a system text */
static QString async_systext(vb_u32 hr)
{
    struct T { vb_u32 hr; const char *en, *el; };
    static const T t[] = {
        { 0x80070057u, "The parameter is incorrect.", "Η παράμετρος είναι εσφαλμένη." },
        { 0x80070006u, "The handle is invalid.", "Ο δείκτης χειρισμού δεν είναι έγκυρος." },
        { 0x80070008u, "Not enough storage is available to process this command.",
          "Ο διαθέσιμος αποθηκευτικός χώρος δεν επαρκεί για να εκτελεστεί αυτή η εντολή." },
        { 0x80072EFDu, "A connection with the server could not be established",
          "Δεν ήταν δυνατή η δημιουργία μιας σύνδεσης με το διακομιστή" },
        { 0x80072EE7u, "The server name or address could not be resolved",
          "Δεν ήταν δυνατή η επίλυση του ονόματος ή της διεύθυνσης του διακομιστή" },
        { 0x80072EE2u, "The operation timed out", "Λήξη χρονικού ορίου λειτουργίας" },
        { 0x80004002u, "No such interface supported", "Δεν υποστηρίζεται τέτοια διασύνδεση" },
        { 0, nullptr, nullptr } };
    for (int i = 0; t[i].en; i++) if (t[i].hr == hr) return QString::fromUtf8(greek() ? t[i].el : t[i].en);
    return QString();
}
/* "System Error &H80070057 (-2147024809).   The parameter is incorrect. " as VB6 writes it */
static void async_raise(vb_i32 number, vb_u32 hr)
{
    QString d = QString("System Error &H%1 (%2).   %3 ").arg(QString::number(hr, 16).toUpper().rightJustified(8, '0'))
                    .arg((vb_i32)hr).arg(async_systext(hr));
    vb_i32 m = vb_mark();
    vb_raise_full(number, vb_app_exename(), vs(d));
    vb_release(m);
}

/* ---------------------------------------------------------------- AsyncProperty */
struct AProp {
    int type = 0, code = 0;
    QString prop, target, status;
    double read = 0, max = 0;
    vbvar value = {};
    vb_u32 err = 0;                             /* Value: this error */
};
static vbvar aprop_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)argc; (void)argv;
    AProp *p = (AProp *)vb_oid(o)->ext;
    std::string n = lower(name);
    *found = 1;
    if (!p) { *found = 0; return vb_v_empty(); }
    if (kind == VBM_LET || kind == VBM_SET) { vb_raise(383); return vb_v_empty(); }
    if (n == "value" || n.empty()) {
        if (p->err) { async_raise((vb_i32)p->err, p->err); return vb_v_empty(); }
        return vb_vtmp(p->value);
    }
    if (n == "asynctype") return vb_v_l(p->type);
    if (n == "propertyname") return vstr(p->prop);
    if (n == "target") return vstr(p->target);
    if (n == "bytesread") return vb_v_l((vb_i32)p->read);
    if (n == "bytesmax") return vb_v_l((vb_i32)p->max);
    if (n == "status") return vstr(p->status);
    if (n == "statuscode") return vb_v_l(p->code);
    *found = 0;
    return vb_v_empty();
}
static void aprop_clear(vbobj o)
{
    AProp *p = (AProp *)o->ext;
    if (p) { vb_vfree(&p->value); delete p; }
    o->ext = nullptr;
}
static const vb_class CLS_APROP = { "AsyncProperty", (vb_i32)sizeof(vbobjh), nullptr, nullptr, aprop_clear, nullptr, 0, "value",
                                    nullptr, nullptr, nullptr, aprop_member, nullptr };

/* ---------------------------------------------------------------- the reads */
struct ARead {
    vbobj uc = nullptr;
    QString prop, target, url, host, path;
    int type = 0, opts = 0, port = 80;
    bool https = false, ftp = false, newport = false, sent = false;
    void *req = nullptr;
    QTimer *timer = nullptr;
};
static std::list<ARead *> async_reads;
static QSet<QString> async_hosts, async_ports;  /* the names and the ports this process has used (WinINet's memory) */
static bool async_alive(ARead *r) { return std::find(async_reads.begin(), async_reads.end(), r) != async_reads.end(); }
static void async_free(ARead *r)
{
    async_reads.remove(r);
    if (r->timer) { r->timer->stop(); r->timer->deleteLater(); r->timer = nullptr; }
    if (r->req) { lbinet_release(r->req); r->req = nullptr; }
    delete r;
}
/* one event; false: the read is no more (cancelled, its control unloaded) */
static bool async_event(ARead *r, bool complete, int code, const QString &status, double read, double max,
                        vbvar value = vbvar(), vb_u32 err = 0)
{
    vbobj uc = r->uc;
    vb_i32 m = vb_mark();
    vbobj o = vb_onew(&CLS_APROP);
    AProp *p = new AProp;
    o->ext = p;
    p->type = r->type; p->prop = r->prop; p->target = r->target;
    p->code = code; p->status = status; p->read = read; p->max = max; p->err = err;
    if (complete && !err) vb_vset(&p->value, value);
    if (complete) async_free(r);               /* (a new read of the same name may start in the event) */
    vbvar a[1] = { vb_v_obj(o) };
    vb_oaddref(uc);
    fire_form(uc, complete ? "asyncreadcomplete" : "asyncreadprogress", 1, a);
    vb_orelease(uc);
    vb_release(m);
    return !complete && async_alive(r);
}
static QString async_mime(const QString &file)
{
    static const std::map<QString, const char *> t = {
        { "bmp", "image/bmp" }, { "dib", "image/bmp" }, { "gif", "image/gif" }, { "jpg", "image/jpeg" },
        { "jpeg", "image/jpeg" }, { "png", "image/png" }, { "ico", "image/x-icon" }, { "txt", "text/plain" },
        { "htm", "text/html" }, { "html", "text/html" }, { "xml", "text/xml" }, { "wav", "audio/wav" },
        { "mid", "audio/mid" }, { "avi", "video/avi" }, { "zip", "application/x-zip-compressed" } };
    auto it = t.find(QFileInfo(file).suffix().toLower());
    return QString::fromUtf8(it != t.end() ? it->second : "application/octet-stream");
}
/* the copy in the cache ("Temporary Internet Files" of LinBASIC): hello[1].txt, hello[2].txt ... */
static QString async_cache(const QString &name, const QByteArray &data)
{
    QString dir = QDir::tempPath() + "/LinBASIC Internet Files";
    QDir().mkpath(dir);
    QFileInfo fi(name.isEmpty() ? QString("index") : name);
    QString base = fi.completeBaseName(), suf = fi.suffix();
    for (int i = 1; i < 100000; i++) {
        QString f = dir + "/" + base + QString("[%1]").arg(i) + (suf.isEmpty() ? QString() : "." + suf);
        if (QFileInfo::exists(f)) continue;
        QFile o(f);
        if (o.open(QIODevice::WriteOnly)) { o.write(data); o.close(); }
        return QDir::toNativeSeparators(f);
    }
    return QString();
}
/* the Value of a read that came: a StdPicture, the file in the cache, a Byte array */
static vbvar async_value(ARead *r, const QByteArray &data, const QString &cachefile, vb_u32 *err)
{
    if (r->type == 0) {
        QImage img = QImage::fromData(data);
        if (img.isNull()) { *err = 0x80004005u; return vb_v_empty(); }
        bool ico = img.isNull() ? false : QFileInfo(r->path).suffix().toLower() == "ico";
        vbobj pic = new_pic(img.convertToFormat(img.hasAlphaChannel() ? QImage::Format_ARGB32 : QImage::Format_RGB32), ico ? 3 : 1);
        return vb_v_obj(pic);
    }
    if (r->type == 1) return vstr(cachefile);
    vbarr *a = vb_atmp();
    if (data.isEmpty()) vb_aempty(a, 1, VBA_PLAIN);
    else {
        vb_adim(a, 1, VBA_PLAIN, 0, 1, 0, (vb_i32)data.size() - 1);
        if (vb_errpend) return vb_v_empty();
        memcpy(a->data, data.constData(), (size_t)data.size());
    }
    return vb_v_arr(a, VT_UI1);
}

/* a local file: during the call */
static void async_local(ARead *r, const QString &path)
{
    QString shown = QDir::toNativeSeparators(path);
    if (!async_event(r, false, 11, QString(), 0, 0)) return;
    QFile f(path);
    if (!QFileInfo(path).isFile() || !f.open(QIODevice::ReadOnly)) { async_free(r); vb_raise(53); return; }
    QByteArray data = f.readAll();
    double n = (double)data.size();
    if (!async_event(r, false, 13, async_mime(path), 0, 0)) return;
    if (!async_event(r, false, 14, shown, 0, n)) return;
    if (!async_event(r, false, 6, "file://" + shown, n, n)) return;
    QString cache = r->type == 1 ? async_cache(QFileInfo(path).fileName(), data) : QString();
    vb_i32 m = vb_mark();
    vb_u32 err = 0;
    vbvar v = async_value(r, data, cache, &err);
    async_event(r, true, err ? 0 : 6, err ? QString() : "file://" + shown, err ? 0 : n, err ? 0 : n, v, err);
    vb_release(m);
}

/* HTTP / FTP: what the thread of the request did since the last time; false when the read ended */
static bool async_poll(ARead *r)
{
    if (!r->req) return false;
    int e;
    while ((e = lbinet_event(r->req)) >= 0)
        if ((e == 5 || e == 6) && !r->sent) { r->sent = true; if (!async_event(r, false, 11, QString(), 0, 0)) return false; }
    if (!lbinet_done(r->req)) return true;
    long n = 0;
    const char *body = lbinet_body(r->req, &n);
    QByteArray data(body, (int)n);
    QString hdrs = QString::fromLatin1(lbinet_headers(r->req));
    bool ok = lbinet_ok(r->req) != 0;
    int code = lbinet_code(r->req), redir = 0;
    QString final = QString::fromLatin1(lbinet_final(r->req, &redir));
    if (!ok) {
        if (code == 12029 && r->newport && !async_event(r, false, 2, r->host, 0, 0)) return false;   /* VB6: once more */
        vb_i32 m = vb_mark();
        async_event(r, true, 0, r->host, 0, 0, vb_v_empty(), 0x80070000u | (vb_u32)(code & 0xFFFF));
        vb_release(m);
        return false;
    }
    if (!r->sent && !async_event(r, false, 11, QString(), 0, 0)) return false;
    int status = r->ftp ? 200 : hdrs.section(' ', 1, 1).toInt();
    if (redir > 0 && !final.isEmpty()) {
        if (!async_event(r, false, 3, final, 0, 0)) return false;
        if (!async_event(r, false, 11, QString(), 0, 0)) return false;
    }
    if (status >= 400) {                        /* VB6: StatusCode 0, an empty Status, the Value an error */
        vb_i32 m = vb_mark();
        async_event(r, true, 0, QString(), 0, 0, vb_v_empty(), status < 500 ? 0x80070006u : 0x80070008u);
        vb_release(m);
        return false;
    }
    QString mime = "application/octet-stream";
    for (const QString &l : hdrs.split("\r\n"))
        if (l.startsWith("Content-Type:", Qt::CaseInsensitive)) { mime = l.mid(13).section(';', 0, 0).trimmed(); break; }
    if (r->ftp) mime = async_mime(r->path);
    QString shown = redir > 0 && !final.isEmpty() ? final : r->target;
    QString file = QUrl(redir > 0 && !final.isEmpty() ? final : r->url).fileName();
    QString cache = async_cache(file, data);
    double sz = (double)data.size();
    if (!async_event(r, false, 13, mime, 0, 0)) return false;
    if (!async_event(r, false, 14, cache, 0, sz)) return false;
    if (!async_event(r, false, 6, shown, sz, sz)) return false;
    vb_i32 m = vb_mark();
    vb_u32 err = 0;
    vbvar v = async_value(r, data, cache, &err);
    async_event(r, true, err ? 0 : 6, err ? QString() : shown, err ? 0 : sz, err ? 0 : sz, v, err);
    vb_release(m);
    return false;
}
static void async_net(ARead *r, const QUrl &u)
{
    r->host = u.host();
    r->path = u.path(QUrl::FullyEncoded);
    if (r->path.isEmpty()) r->path = "/";
    if (u.hasQuery()) r->path += "?" + u.query(QUrl::FullyEncoded);
    r->port = u.port(r->ftp ? 21 : r->https ? 443 : 80);
    QByteArray host = r->host.toUtf8(), path = r->path.toUtf8(), user = u.userName().toUtf8(), pass = u.password().toUtf8();
    if (r->ftp) {
        QByteArray fp = u.path(QUrl::FullyDecoded).toUtf8();
        r->req = lbinet_ftpurl(host.constData(), r->port, fp.constData(), user.isEmpty() ? "anonymous" : user.constData(),
                               user.isEmpty() ? "anonymous@" : pass.constData(), 60);
    } else
        r->req = lbinet_http(nullptr, r->https, 1, "GET", host.constData(), r->port, path.constData(), user.constData(),
                             pass.constData(), "", nullptr, 0, 60, 0, "");
    bool newhost = !async_hosts.contains(r->host.toLower());
    QString hp = r->host.toLower() + ":" + QString::number(r->port);
    r->newport = !async_ports.contains(hp);
    async_hosts.insert(r->host.toLower());
    async_ports.insert(hp);
    auto begin = [r, newhost]() -> bool {      /* FindingResource and Connecting: a server, a port not used before */
        if (newhost && !async_event(r, false, 1, r->host, 0, 0)) return false;
        if (r->newport && !async_event(r, false, 2, r->host, 0, 0)) return false;
        return true;
    };
    if (r->opts & 1) {                          /* vbAsyncReadSynchronousDownload: all during the call */
        if (!begin()) return;
        while (async_alive(r) && async_poll(r)) QThread::msleep(5);
        return;
    }
    r->timer = new QTimer;
    r->timer->setInterval(10);
    auto first = std::make_shared<bool>(true);
    QObject::connect(r->timer, &QTimer::timeout, [r, first, begin]() {
        if (!async_alive(r)) return;
        if (*first) { *first = false; if (!begin()) return; }
        async_poll(r);
    });
    r->timer->start();
}
/* UserControl.AsyncRead Target, AsyncType [, PropertyName [, AsyncReadOptions]] */
static void async_read(vbobj uc, vb_i32 argc, vbvar *argv)
{
    if (argc < 2) { vb_raise(449); return; }
    vb_i32 m = vb_mark();
    QString target = qs(vb_v2s(argv[0]));
    vb_i32 type = vb_v2l(argv[1]);
    QString prop = argc > 2 && !vb_ismissing(argv[2]) ? qs(vb_v2s(argv[2])) : QString();
    int opts = argc > 3 && !vb_ismissing(argv[3]) ? vb_v2l(argv[3]) : 0;
    vb_release(m);
    if (vb_errpend) return;
    if (type < 0 || type > 2) { async_raise((vb_i32)0x8007000Eu, 0x80070057u); return; }
    for (ARead *x : async_reads)
        if (x->uc == uc && x->prop.compare(prop, Qt::CaseInsensitive) == 0) { vb_raise(689); return; }
    if (target.trimmed().isEmpty()) { vb_raise(693); return; }
    auto *r = new ARead;
    r->uc = uc; r->prop = prop; r->target = target; r->type = type; r->opts = opts;
    async_reads.push_back(r);
    QString t = target.trimmed();
    int sc = t.indexOf("://");
    QString scheme = sc > 1 ? t.left(sc).toLower() : QString();
    if (scheme.isEmpty() || scheme == "file" || t.startsWith("file:", Qt::CaseInsensitive)) {
        QString p = t;
        if (p.startsWith("file:", Qt::CaseInsensitive)) {
            p = QUrl::fromPercentEncoding(p.mid(5).toUtf8());
            while (p.startsWith('/') || p.startsWith('\\')) p.remove(0, 1);
#ifdef _WIN32
            if (!(p.size() > 1 && p[1] == ':')) p = "\\\\" + p;   /* file://server/share/... */
#else
            if (!(p.size() > 1 && p[1] == ':')) p = "/" + p;
#endif
        }
        r->path = native(p);
        async_local(r, r->path);
        return;
    }
    if (scheme != "http" && scheme != "https" && scheme != "ftp") { async_free(r); vb_raise(693); return; }
    r->https = scheme == "https";
    r->ftp = scheme == "ftp";
    r->url = t;
    QUrl u(t);
    if (!u.isValid() || u.host().isEmpty()) { async_free(r); vb_raise(693); return; }
    async_net(r, u);
}
/* UserControl.CancelAsyncRead [Property]: without a name, the read without a name */
static void async_cancel(vbobj uc, vb_i32 argc, vbvar *argv)
{
    vb_i32 m = vb_mark();
    QString prop = argc > 0 && !vb_ismissing(argv[0]) ? qs(vb_v2s(argv[0])) : QString();
    vb_release(m);
    if (vb_errpend) return;
    for (ARead *x : async_reads)
        if (x->uc == uc && x->prop.compare(prop, Qt::CaseInsensitive) == 0) { async_free(x); return; }
    async_raise((vb_i32)0x8007000Eu, 0x80070057u);
}
/* the user control goes: its reads stop */
static void async_cancel_all(vbobj uc)
{
    std::vector<ARead *> mine;
    for (ARead *x : async_reads) if (x->uc == uc) mine.push_back(x);
    for (ARead *x : mine) async_free(x);
}

/* ---------------------------------------------------------------- Hyperlink */
static vbvar hlink_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind;
    std::string n = lower(name);
    *found = 1;
    (void)o;
    if (n == "goback" || n == "goforward") { async_raise((vb_i32)0x80004002u, 0x80004002u); return vb_v_empty(); }
    if (n == "navigateto") {
        if (argc < 1) { vb_raise(449); return vb_v_empty(); }
        vb_i32 m = vb_mark();
        QString t = qs(vb_v2s(argv[0])).trimmed();
        QString loc = argc > 1 && !vb_ismissing(argv[1]) ? qs(vb_v2s(argv[1])) : QString();
        vb_release(m);
        if (vb_errpend) return vb_v_empty();
        if (t.isEmpty()) { async_raise((vb_i32)0x80070057u, 0x80070057u); return vb_v_empty(); }
        QUrl u = t.contains("://") || t.startsWith("mailto:", Qt::CaseInsensitive) ? QUrl(t) : QUrl::fromLocalFile(native(t));
        if (!loc.isEmpty()) u.setFragment(loc.startsWith('#') ? loc.mid(1) : loc);
        if (monkey_on) { monkey_log("NavigateTo " + u.toString()); return vb_v_empty(); }
        const char *nb = getenv("VBRT_NO_BROWSER");   /* the tests: no browser */
        if (nb && *nb) return vb_v_empty();
        QDesktopServices::openUrl(u);
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
static const vb_class CLS_HLINK = { "Hyperlink", (vb_i32)sizeof(vbobjh), nullptr, nullptr, nullptr, nullptr, 0, nullptr,
                                    nullptr, nullptr, nullptr, hlink_member, nullptr };
static vbvar async_hyperlink(vbobj uc)
{
    vbobj h = vb_onew(&CLS_HLINK);
    h->ext = uc;                                /* weak: the user control outlives it */
    return vb_vtmp(vb_v_obj(h));
}
