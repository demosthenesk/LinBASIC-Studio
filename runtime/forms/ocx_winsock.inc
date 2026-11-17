/* ocx_winsock.inc - the Winsock control (MSWINSCK.OCX, MSWinsockLib.Winsock): TCP clients and servers
 * and UDP, with the properties, methods, events and errors of VB6. The sockets are in runtime/net.c;
 * a timer of the control asks them what is ready, so the events come from the message loop, as in
 * VB6 (DoEvents, or the program waiting). Included by ocx.inc.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <map>
#include <deque>
#include <memory>
extern "C" {
int lbnet_resolve(const char *host, char ip[16]);
int lbnet_is_ip(const char *host);
long long lbnet_open(int udp);
void lbnet_close(long long h);
int lbnet_bind(long long h, const char *ip, int port);
int lbnet_listen(long long h);
int lbnet_connect(long long h, const char *ip, int port);
int lbnet_poll(long long h, int *rd, int *wr, int *ex);
int lbnet_sockerr(long long h);
long long lbnet_accept(long long h);
int lbnet_send(long long h, const char *b, int n);
int lbnet_recv(long long h, char *b, int n);
int lbnet_sendto(long long h, const char *ip, int port, const char *b, int n);
int lbnet_recvfrom(long long h, char *b, int n, char ip[16], int *port);
int lbnet_name(long long h, int peer, char ip[16], int *port);
void lbnet_hostname(char *b, int n);
void lbnet_localip(char ip[16]);
}

/* the texts of the errors of the control (ErrorConstants of MSWinsockLib) */
static const char *wsk_text(int n)
{
    switch (n) {
    case 7: return "Out of memory";
    case 380: return "Invalid property value";
    case 383: return "Property is read-only";
    case 394: return "Property is write-only";
    case 40006: return "Wrong protocol or connection state for the requested transaction or request";
    case 40014: return "The argument passed to a function was not in the correct format or in the specified range";
    case 40017: return "Successful";
    case 40018: return "Unsupported variant types";
    case 40020: return "Invalid operation at current state";
    case 40021: return "Argument is out of range";
    case 40026: return "Wrong protocol or connection state for the requested transaction or request";
    case 10004: return "The operation is canceled";
    case 10013: return "The requested address is a broadcast address, but flag is not set";
    case 10014: return "Invalid argument";
    case 10022: return "Invalid argument";
    case 10035: return "Socket is non-blocking and the specified operation will block";
    case 10036: return "A blocking Winsock operation is in progress";
    case 10037: return "The operation is completed. No blocking operation is in progress.";
    case 10038: return "The descriptor is not a socket";
    case 10040: return "The datagram is too large to fit into the buffer and is truncated";
    case 10043: return "The specified port is not supported";
    case 10047: return "Address Family is not supported";
    case 10048: return "Address in use";
    case 10049: return "Address not available from the local machine";
    case 10050: return "Network subsystem failed";
    case 10051: return "The network cannot be reached from this host at this time";
    case 10052: return "Connection has timed out when SO_KEEPALIVE is set";
    case 10053: return "Connection is aborted due to timeout or other failure";
    case 10054: return "Connection is reset by remote side";
    case 10055: return "No buffer space is available";
    case 10056: return "Socket is already connected";
    case 10057: return "Socket is not connected";
    case 10058: return "Socket has been shut down";
    case 10060: return "Connection has timed out";
    case 10061: return "Connection is forcefully rejected";
    case 10065: return "The network cannot be reached from this host at this time";
    case 10093: return "WinsockInit should be called first";
    case 11001: return "Authoritative answer: Host not found";
    case 11002: return "Non-Authoritative answer: Host not found";
    case 11003: return "Non-recoverable errors";
    case 11004: return "Valid name, no data record of requested type";
    }
    return "Unknown error";
}

/* the connections a listening control has been asked for: requestID -> the socket, until a
 * Winsock control accepts it (in ConnectionRequest; after it the connection is closed) */
static std::map<vb_i32, long long> &wsk_pending()
{
    static std::map<vb_i32, long long> m;
    return m;
}

struct OcxWinsock : Ocx {
    enum { CLOSED, OPEN, LISTENING, PENDING, RESOLVING, RESOLVED, CONNECTING, CONNECTED, CLOSING, ERR };
    long long s = -1;
    int proto = 0, state = CLOSED;
    QString rhost, rhostip;
    long rport = 0, lport = 0;
    bool host_set = false;                      /* RemoteHost given when the program runs: a UDP control with
                                                 * the RemoteHost of the .frm only cannot send (10047, VB6) */
    QByteArray in, out;                         /* TCP: what has come and not been read; what waits to go */
    struct Dgram { QByteArray d; QString ip; int port; };
    std::deque<Dgram> dgrams;                   /* UDP: the datagrams not read yet */
    unsigned long dg_taken = 0;                 /* datagrams GetData has taken (or thrown away) */
    long sent_mark = 0;                         /* bytes sent that SendProgress has not told yet */
    long send_total = 0;
    bool complete_due = false, busy = false;
    QTimer *timer = nullptr;
    std::shared_ptr<bool> alive = std::make_shared<bool>(true);

    ~OcxWinsock() override
    {
        *alive = false;
        lbnet_close(s);
        if (timer) { timer->stop(); timer->deleteLater(); }   /* it may be in its own timeout (Unload in an event) */
    }
    QWidget *make(QWidget *parent) override { (void)parent; return nullptr; }
    void load(const char *const *p) override
    {
        proto = (int)propn(p, "Protocol", 0);
        prop(p, "RemoteHost", &rhost);
        rport = propn(p, "RemotePort", 0);
        lport = propn(p, "LocalPort", 0);
    }

    vbvar err(int n) { vb_raise_full(n, vs(QString("Winsock")), vs(QString(wsk_text(n)))); return vb_v_empty(); }
    void run()
    {
        if (!timer) {
            timer = new QTimer();
            QObject::connect(timer, &QTimer::timeout, [this]() { tick(); });
        }
        if (!timer->isActive()) timer->start(10);
    }
    void shut()
    {
        lbnet_close(s);
        s = -1;
        state = CLOSED;
        in.clear(); out.clear(); dgrams.clear();
        sent_mark = send_total = 0;
        complete_due = false;
        if (timer) timer->stop();
    }
    /* an event; false: the control is gone (the form was unloaded in it) */
    bool fire(const char *ev, int argc = 0, vbvar *argv = nullptr)
    {
        std::shared_ptr<bool> a = alive;
        fire_ctl(c, ev, argc, argv);
        return *a;
    }
    bool fire_error(int n)
    {
        /* Source: the file of the control, as VB6 gives it; CancelDisplay comes True */
        vbvar a[7] = { vb_v_i((vb_i16)n), vstr(QString(wsk_text(n))), vb_v_l((vb_i32)(0x800A0000u | (unsigned)n)),
                       vstr(QString("C:\\Windows\\system32\\MSWINSCK.ocx")), vstr(QString()), vb_v_l(0), vb_v_bool(-1) };
        return fire("error", 7, a);
    }
    void failed(int n)
    {
        lbnet_close(s);
        s = -1;
        state = ERR;
        if (timer) timer->stop();
        fire_error(n);
    }
    void names()
    {
        char ip[16]; int port;
        if (!lbnet_name(s, 0, ip, &port)) lport = port;
        if (!lbnet_name(s, 1, ip, &port)) { rhostip = ip; rport = port; }
    }
    /* the connection on its way (an address resolved) */
    void start_connect()
    {
        state = CONNECTING;
        s = lbnet_open(0);
        if (s == -1) { failed(10050); return; }
        if (lport) {
            int e = lbnet_bind(s, nullptr, (int)lport);
            if (e) { failed(e); return; }
        }
        int e = lbnet_connect(s, rhostip.toLatin1().constData(), (int)rport);
        if (e && e != 10035 && e != 10036) { failed(e); return; }   /* told by the timer, as Windows does */
        run();
    }
    void tick()
    {
        if (busy) return;
        busy = true;
        std::shared_ptr<bool> a = alive;
        step();
        if (*a) busy = false;
    }
    void step()
    {
        if (state == RESOLVING) {
            char ip[16];
            if (rhost.trimmed().isEmpty()) { failed(10061); return; }
            if (lbnet_resolve(rhost.toUtf8().constData(), ip)) { failed(11001); return; }
            rhostip = ip;
            state = RESOLVED;
            start_connect();
            return;
        }
        if (s == -1) { if (timer) timer->stop(); return; }
        int rd, wr, ex;
        if (lbnet_poll(s, &rd, &wr, &ex)) return;
        if (state == CONNECTING) {
            if (!wr && !ex) return;
            int e = lbnet_sockerr(s);
            if (e || ex) { failed(e ? e : 10061); return; }
            state = CONNECTED;
            names();
            if (!fire("connect")) return;
            return;
        }
        if (state == LISTENING) {
            if (!rd) return;
            long long ns;
            while (state == LISTENING && (ns = lbnet_accept(s)) != -1) {
                vb_i32 id = (vb_i32)ns;
                wsk_pending()[id] = ns;
                char ip[16]; int port;
                if (!lbnet_name(ns, 1, ip, &port)) { rhostip = ip; rport = port; }
                vbvar arg = vb_v_l(id);
                bool ok = fire("connectionrequest", 1, &arg);
                auto it = wsk_pending().find(id);
                if (it != wsk_pending().end()) { lbnet_close(it->second); wsk_pending().erase(it); }   /* not accepted */
                if (!ok || s == -1) return;
            }
            return;
        }
        if (proto == 1 && (state == OPEN || state == CONNECTED)) {      /* UDP */
            /* VB6 (oracle f_wsock): DataArrival tells the bytes of all the datagrams waiting; GetData reads
             * one, and only then RemoteHostIP / RemotePort become its sender (they stay); a datagram the
             * event has not read is lost; after a GetData in the event the next one comes in its own event */
            static char buf[65536];
            char ip[16]; int port, n;
            while (rd && s != -1 && (n = lbnet_recvfrom(s, buf, sizeof buf, ip, &port)) >= 0)
                dgrams.push_back({ QByteArray(buf, n), QString(ip), port });
            while (!dgrams.empty() && s != -1) {
                unsigned long before = dg_taken;
                vbvar arg = vb_v_l(udp_bytes());
                if (!fire("dataarrival", 1, &arg)) return;
                if (dg_taken == before && !dgrams.empty()) { dgrams.pop_front(); dg_taken++; }
            }
            return;
        }
        if (state == CONNECTED) {
            if (!out.isEmpty() && wr) send_more();
            if (sent_mark > 0) {
                long done = sent_mark;
                sent_mark = 0;
                vbvar args[2] = { vb_v_l((vb_i32)done), vb_v_l((vb_i32)out.size()) };
                if (!fire("sendprogress", 2, args)) return;
            }
            if (complete_due && out.isEmpty()) {
                complete_due = false;
                if (!fire("sendcomplete")) return;
            }
            if (state != CONNECTED || s == -1 || !rd) return;
            static char buf[65536];
            int got = 0, n;
            bool closed = false;
            int e = 0;
            while ((n = lbnet_recv(s, buf, sizeof buf)) > 0) { in.append(buf, n); got += n; }
            if (n == 0) closed = true;
            else if (n != -10035) e = -n;
            if (got) {
                vbvar arg = vb_v_l((vb_i32)in.size());
                if (!fire("dataarrival", 1, &arg)) return;
            }
            if (state != CONNECTED || s == -1) return;
            if (closed) {
                state = CLOSING;
                if (timer) timer->stop();
                fire("close");
                return;
            }
            if (e) failed(e);
        }
    }
    vb_i32 udp_bytes() const
    {
        vb_i32 n = 0;
        for (auto &d : dgrams) n += (vb_i32)d.d.size();
        return n;
    }
    void send_more()
    {
        while (!out.isEmpty()) {
            int n = lbnet_send(s, out.constData(), (int)std::min<qsizetype>(out.size(), 65536));
            if (n <= 0) break;
            out.remove(0, n);
            sent_mark += n;
        }
    }
    /* the bytes of a value SendData sends: a String in ANSI, a Byte array as it is, numbers as
     * their bytes in memory; false: a type it cannot send (40018) */
    static bool bytes_of(vbvar v, QByteArray *b)
    {
        switch (v.vt) {
        case VT_NULL: return true;              /* nothing is sent (VB6) */
        case VT_DISPATCH: return v.u.o == nullptr;
        case VT_BSTR: {
            vbs t = vb_v2s(v);
            for (vb_i32 i = 0; i < t.len; i++) { int ch = vb_w2ansi(t.p[i]); b->append((char)(ch < 0 ? '?' : ch)); }
            return true;
        }
        case VT_UI1: b->append((char)v.u.b); return true;
        case VT_I2: case VT_BOOL: b->append((const char *)&v.u.i, 2); return true;
        case VT_I4: case VT_ERROR: b->append((const char *)&v.u.l, 4); return true;
        case VT_R4: b->append((const char *)&v.u.f, 4); return true;
        case VT_R8: case VT_DATE: b->append((const char *)&v.u.d, 8); return true;
        case VT_CY: b->append((const char *)&v.u.cy, 8); return true;
        case VT_ARRAY | VT_UI1: {
            vbarr *a = vb_v_arrptr(v);
            if (a && a->nd) b->append((const char *)a->data, (int)a->n);
            return true;
        }
        }
        return false;
    }
    /* GetData / PeekData: data as the type asked for (the type of the variable when it is left out) */
    vbvar get_data(vb_i32 argc, vbvar *argv, bool peek)
    {
        if (argc < 1) { vb_raise(449); return vb_v_empty(); }
        QByteArray *src = nullptr;
        QByteArray one;
        if (proto == 1) {
            if (dgrams.empty()) {               /* nothing waiting: the variable stays as it is (VB6) */
                if (state != OPEN && state != CONNECTED) return err(40006);
                return vb_v_empty();
            }
            one = dgrams.front().d;
            rhostip = dgrams.front().ip;
            rport = dgrams.front().port;
            src = &one;
        } else {
            if (state != CONNECTED && in.isEmpty()) return err(40006);
            src = &in;
        }
        int type = given(argc, argv, 1) ? (int)vb_v2l(argv[1]) : 0;
        if (vb_errpend) return vb_v_empty();
        /* the type of the variable; an Empty Variant gets a Byte array (VB6 oracle f_wsock) */
        if (!type) type = argv[0].vt == VT_EMPTY || argv[0].vt == (VT_ARRAY | VT_UI1) ? (VT_ARRAY | VT_UI1) :
                          argv[0].vt == VT_NULL || (argv[0].vt & VT_ARRAY) ? VT_BSTR : argv[0].vt;
        long maxlen = given(argc, argv, 2) ? vb_v2l(argv[2]) : -1;
        if (vb_errpend) return vb_v_empty();
        int fixed = type == VT_UI1 ? 1 : type == VT_I2 || type == VT_BOOL ? 2 : type == VT_I4 || type == VT_R4 || type == VT_ERROR ? 4 :
                    type == VT_R8 || type == VT_CY || type == VT_DATE ? 8 : 0;
        if (!fixed && type != VT_BSTR && type != (VT_ARRAY | VT_UI1)) return err(40018);
        int k = fixed ? fixed : (int)src->size();
        if (proto == 1 && maxlen >= 0 && maxlen < src->size()) {   /* the datagram is lost (VB6) */
            dgrams.pop_front(); dg_taken++;
            return err(10040);
        }
        if (!fixed && maxlen >= 0 && maxlen < k) k = (int)maxlen;
        if (k > src->size()) k = (int)src->size();
        QByteArray part = src->left(k);
        vbvar r = vb_v_empty();
        if (fixed) {
            char raw[8] = { 0 };
            memcpy(raw, part.constData(), (size_t)part.size());
            switch (type) {
            case VT_UI1: r = vb_v_byte((vb_u8)raw[0]); break;
            case VT_I2: { vb_i16 x; memcpy(&x, raw, 2); r = vb_v_i(x); break; }
            case VT_BOOL: { vb_i16 x; memcpy(&x, raw, 2); r = vb_v_bool(x ? -1 : 0); break; }
            case VT_I4: case VT_ERROR: { vb_i32 x; memcpy(&x, raw, 4); r = vb_v_l(x); break; }
            case VT_R4: { float x; memcpy(&x, raw, 4); r = vb_v_f(x); break; }
            case VT_R8: { double x; memcpy(&x, raw, 8); r = vb_v_d(x); break; }
            case VT_DATE: { double x; memcpy(&x, raw, 8); r = vb_v_date(x); break; }
            case VT_CY: { vb_cy x; memcpy(&x, raw, 8); r = vb_v_c(x); break; }
            }
        } else if (type == (VT_ARRAY | VT_UI1)) {
            vbarr *a = vb_atmp();
            if (part.isEmpty()) vb_aempty(a, 1, VBA_PLAIN);
            else {
                vb_adim(a, 1, VBA_PLAIN, 0, 1, 0, (vb_i32)part.size() - 1);
                if (vb_errpend) return vb_v_empty();
                memcpy(a->data, part.constData(), (size_t)part.size());
            }
            r = vb_v_arr(a, VT_UI1);
        } else {
            vbs t = vb_stmp((vb_i32)part.size());
            for (int i = 0; i < part.size(); i++) t.p[i] = vb_ansi2w((unsigned char)part[i]);
            r = vb_v_s(t);
        }
        vb_vset(&argv[0], r);
        if (!peek) {
            if (proto == 1) { if (!dgrams.empty()) { dgrams.pop_front(); dg_taken++; } }
            else in.remove(0, k);
        }
        return vb_v_empty();
    }
    /* the socket of a control that is closed, bound to LocalPort (Bind, Listen, UDP SendData) */
    int open_bound(long port, const QString &ip)
    {
        s = lbnet_open(proto == 1);
        if (s == -1) return 10050;
        int e = lbnet_bind(s, ip.isEmpty() ? nullptr : ip.toLatin1().constData(), (int)port);
        if (e) { lbnet_close(s); s = -1; return e; }
        char a[16]; int p;
        if (!lbnet_name(s, 0, a, &p)) lport = p;
        return 0;
    }

    vbvar member(const std::string &n, vb_i32 kind, bool set, vbvar val, vb_i32 argc, vbvar *argv, int *found) override
    {
        (void)kind;
        auto port_prop = [&](long &f) -> vbvar {
            if (!set) return vb_v_l((vb_i32)f);
            long v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v < 0 || v > 65535) return err(&f == &rport ? 40014 : 380);   /* VB6: RemotePort 40014, LocalPort 380 */
            f = v;
            return vb_v_empty();
        };
        if (n == "protocol") {
            if (!set) return vb_v_l(proto);
            long v = vb_v2l(val);
            if (vb_errpend) return vb_v_empty();
            if (v != 0 && v != 1) { vb_raise(380); return vb_v_empty(); }   /* Err.Source: the program (VB6) */
            if (state != CLOSED) return err(40020);
            proto = (int)v;
            return vb_v_empty();
        }
        if (n == "remotehost") {
            if (!set) return vstr(rhost);
            rhost = qs(vb_v2s(val));
            host_set = true;
            return vb_v_empty();
        }
        if (n == "remoteport") return port_prop(rport);
        if (n == "localport") {
            if (set && state != CLOSED) return err(40020);
            return port_prop(lport);
        }
        /* the read-only properties: 450 when they are assigned (late bound; early bound does not compile) */
        if (set && (n == "remotehostip" || n == "localhostname" || n == "localip" || n == "sockethandle" || n == "state" ||
                    n == "bytesreceived")) { vb_raise(450); return vb_v_empty(); }
        if (n == "remotehostip") return vstr(rhostip);
        if (n == "localhostname") {
            char b[256]; lbnet_hostname(b, sizeof b);
            return vstr(QString::fromLocal8Bit(b));
        }
        if (n == "localip") {
            char ip[16]; lbnet_localip(ip);
            return vstr(QString(ip));
        }
        if (n == "sockethandle") { return vb_v_l(s == -1 ? -1 : (vb_i32)s); }
        if (n == "state") { return vb_v_i((vb_i16)state); }
        if (n == "bytesreceived") {
            if (proto == 1) {                   /* the bytes of the datagrams waiting; no socket: 10038 (VB6) */
                if (s == -1) return err(10038);
                return vb_v_l(udp_bytes());
            }
            return vb_v_l((vb_i32)in.size());
        }
        if (n == "connect") {
            if (state != CLOSED && !(proto == 1 && state == OPEN)) return err(40020);
            if (given(argc, argv, 0)) { rhost = qs(vb_v2s(argv[0])); host_set = true; }
            if (given(argc, argv, 1)) {
                long v = vb_v2l(argv[1]);
                if (vb_errpend) return vb_v_empty();
                if (v < 0 || v > 65535) return err(40021);
                rport = v;
            }
            if (vb_errpend) return vb_v_empty();
            if (proto == 1) {                   /* UDP: the other side of SendData; no error (VB6) */
                host_set = true;
                if (state == CLOSED) {
                    int e = open_bound(lport, QString());
                    if (e) return err(e);
                    state = OPEN;
                    run();
                }
                return vb_v_empty();
            }
            QByteArray h = rhost.trimmed().toUtf8();
            in.clear(); out.clear();
            if (h.isEmpty()) {                  /* VB6: 0.0.0.0, refused (10061) when the timer comes */
                state = RESOLVING;
                run();
            } else if (lbnet_is_ip(h.constData())) {
                char ip[16];
                lbnet_resolve(h.constData(), ip);
                rhostip = ip;
                start_connect();
            } else {
                state = RESOLVING;                  /* the name is resolved by the timer */
                run();
            }
            return vb_v_empty();
        }
        if (n == "listen") {
            if (proto == 1) return err(40026);
            if (state == LISTENING) return err(10037);
            if (state != CLOSED && state != OPEN) return err(40020);
            if (state == CLOSED) {
                int e = open_bound(lport, QString());
                if (e) return err(e);
            }
            int e = lbnet_listen(s);
            if (e) { lbnet_close(s); s = -1; state = CLOSED; return err(e); }
            state = LISTENING;
            run();
            return vb_v_empty();
        }
        if (n == "accept") {
            if (argc < 1) { vb_raise(449); return vb_v_empty(); }
            vb_i32 id = vb_v2l(argv[0]);
            if (vb_errpend) return vb_v_empty();
            if (state != CLOSED) return err(40020);
            auto it = wsk_pending().find(id);
            if (it == wsk_pending().end()) return err(10038);
            s = it->second;
            wsk_pending().erase(it);
            proto = 0;
            state = CONNECTED;
            in.clear(); out.clear();
            names();
            run();
            return vb_v_empty();
        }
        if (n == "bind") {
            if (state != CLOSED) return err(40020);
            long port = lport;
            if (given(argc, argv, 0)) {
                port = vb_v2l(argv[0]);
                if (vb_errpend) return vb_v_empty();
                if (port < 0 || port > 65535) return err(40021);
            }
            QString ip = given(argc, argv, 1) ? qs(vb_v2s(argv[1])) : QString();
            int e = open_bound(port, ip);
            if (e) return err(e);
            state = OPEN;
            run();
            return vb_v_empty();
        }
        if (n == "senddata") {
            if (argc < 1) { vb_raise(449); return vb_v_empty(); }
            QByteArray data;
            if (!bytes_of(argv[0], &data)) return err(40018);
            if (proto == 1) {
                if (state == CLOSED) {
                    int e = open_bound(lport, QString());
                    if (e) return err(e);
                    state = OPEN;
                    run();
                }
                if (state != OPEN && state != CONNECTED) return err(40006);
                if (!host_set) return err(10047);
                char ip[16];
                QByteArray h = rhost.trimmed().toUtf8();
                if (h.isEmpty() || lbnet_resolve(h.constData(), ip)) return err(11001);
                int r = lbnet_sendto(s, ip, (int)rport, data.constData(), data.size());
                if (r < 0) return err(-r);
                return vb_v_empty();
            }
            if (state != CONNECTED) return err(40006);
            out.append(data);
            send_total += data.size();
            send_more();
            complete_due = true;
            run();
            return vb_v_empty();
        }
        if (n == "getdata") return get_data(argc, argv, false);
        if (n == "peekdata") return get_data(argc, argv, true);
        if (n == "close") { shut(); return vb_v_empty(); }
        if (n == "aboutbox") return vb_v_empty();
        *found = 0;
        return vb_v_empty();
    }
};
static Ocx *new_winsock() { return new OcxWinsock; }
