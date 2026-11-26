#define _GNU_SOURCE
/* inet.c - the requests of the Internet Transfer control (forms/ocx_inet.inc), each in a thread of its own:
 *   http://   our own HTTP/1.1 client: the requests of VB6's Inet byte for byte (oracle f_inet: the headers,
 *             their order, keep-alive, redirects, Basic authentication and WinINet's memory of it)
 *   https://  Linux: libcurl, loaded when it is first needed (libcurl.so.4 / libcurl-gnutls.so.4);
 *             Windows: WinINet, the library of MSINET.OCX itself
 *   ftp://    our own FTP client (oracle f_inetftp): a session per control, active mode (PORT) as WinINet,
 *             passive (PASV) when the server is beyond a NAT router
 * The control reads the states (StateChanged) from a queue; the errors are the numbers of WinINet.
 * Apart from vbrt.h (windows.h names VT_EMPTY too).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifdef _WIN32
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0601
#endif
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <wininet.h>
#define strcasecmp _stricmp
#define strncasecmp _strnicmp
#else
#include <errno.h>
#include <fcntl.h>
#include <unistd.h>
#include <netdb.h>
#include <poll.h>
#include <dlfcn.h>
#include <pthread.h>
#include <sys/socket.h>
#include <sys/stat.h>
#include <netinet/in.h>
#include <arpa/inet.h>
#endif
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#define UA "Microsoft URL Control - 6.01.9839"

/* ---------------------------------------------------------------- the platform */
#ifdef _WIN32
typedef SOCKET sock_t;
#define BADS INVALID_SOCKET
static void closes(sock_t s) { if (s != BADS) closesocket(s); }
static int neterr(void) { return WSAGetLastError(); }
#define E_INPROGRESS WSAEWOULDBLOCK
typedef CRITICAL_SECTION mutex_t;
static void mu_init(mutex_t *m) { InitializeCriticalSection(m); }
static void mu_lock(mutex_t *m) { EnterCriticalSection(m); }
static void mu_unlock(mutex_t *m) { LeaveCriticalSection(m); }
static void mu_free(mutex_t *m) { DeleteCriticalSection(m); }
static void net_init(void) { static int done; if (!done) { WSADATA w; WSAStartup(MAKEWORD(2, 2), &w); done = 1; } }
static FILE *fopen8(const char *p, const char *m)
{
    wchar_t wp[2048], wm[8];
    MultiByteToWideChar(CP_UTF8, 0, p, -1, wp, 2048);
    MultiByteToWideChar(CP_UTF8, 0, m, -1, wm, 8);
    return _wfopen(wp, wm);
}
static int exists8(const char *p)
{
    wchar_t wp[2048];
    MultiByteToWideChar(CP_UTF8, 0, p, -1, wp, 2048);
    return GetFileAttributesW(wp) != INVALID_FILE_ATTRIBUTES;
}
#else
typedef int sock_t;
#define BADS (-1)
static void closes(sock_t s) { if (s != BADS) close(s); }
static int neterr(void) { return errno; }
#define E_INPROGRESS EINPROGRESS
typedef pthread_mutex_t mutex_t;
static void mu_init(mutex_t *m) { pthread_mutex_init(m, 0); }
static void mu_lock(mutex_t *m) { pthread_mutex_lock(m); }
static void mu_unlock(mutex_t *m) { pthread_mutex_unlock(m); }
static void mu_free(mutex_t *m) { pthread_mutex_destroy(m); }
static void net_init(void) {}
static FILE *fopen8(const char *p, const char *m) { return fopen(p, m); }
static int exists8(const char *p) { struct stat st; return stat(p, &st) == 0; }
#endif

/* ---------------------------------------------------------------- a growing buffer */
typedef struct { char *p; size_t n, cap; } buf;
static void bput(buf *b, const void *d, size_t n)
{
    if (b->n + n + 1 > b->cap) {
        size_t c = b->cap ? b->cap * 2 : 1024;
        while (c < b->n + n + 1) c *= 2;
        char *q = (char *)realloc(b->p, c);
        if (!q) return;
        b->p = q; b->cap = c;
    }
    if (n) memcpy(b->p + b->n, d, n);
    b->n += n;
    b->p[b->n] = 0;
}
static void bputs(buf *b, const char *s) { bput(b, s, strlen(s)); }
static void bfree(buf *b) { free(b->p); b->p = 0; b->n = b->cap = 0; }
static void bclear(buf *b) { b->n = 0; if (b->p) b->p[0] = 0; }

/* ---------------------------------------------------------------- the request */
enum { K_HTTP = 1, K_FTP, K_FTPURL };
typedef struct lbconn { sock_t s; char host[256]; int port; char rb[16384]; int rn, rpos; } lbconn;
typedef struct lbftp {
    lbconn c;                           /* the control connection */
    char user[256], pass[256];
    char cwd[1024]; int cwd_known;
    int passive;                        /* PASV: the server is beyond a NAT router */
} lbftp;

typedef struct lbinet {
#ifdef _WIN32
    HANDLE th;
#else
    pthread_t th;
#endif
    mutex_t mu;
    volatile int cancel;
    int finished, autofree, noend;
    volatile int stage;                 /* the highest state so far */      /* noend: the last states were sent already (11 9 10) */
    int ev[512], ev_in, ev_out;
    /* what is asked */
    int kind, https, openurl, timeout, access, port;
    char method[32], host[256], path[2048], user[256], pass[256], proxy[512], a1[2048], a2[2048];
    buf hdrs, data;
    lbconn *conn;                       /* HTTP: the kept connection of the control */
    lbftp *ftp;                         /* FTP: the session of the control */
    /* what comes */
    int ok, code, status;
    int cnew, cpre;                     /* https: a new connection was opened; the request is being sent */
    char info[2048];
    buf body, headers;
    char finalurl[2400];                /* HTTP: the URL of the last redirect (AsyncRead: its Status) */
    int redirects;
} lbinet;

static void ev(lbinet *r, int st)
{
    mu_lock(&r->mu);
    if ((r->ev_in + 1) % 512 != r->ev_out) { r->ev[r->ev_in] = st; r->ev_in = (r->ev_in + 1) % 512; }
    if (st > r->stage) r->stage = st;
    mu_unlock(&r->mu);
}

/* the short texts of MSINET.OCX for the errors of WinINet (ResponseInfo) */
static const char *inet_text(int code)
{
    static const char *t1[] = { 0, "Out of handles", "Timeout", "Extended error", "Internal error", "Invalid URL",
        "Unrecognized scheme", "Name not resolved", "Protocol not found", "Invalid option", "Bad option length",
        "Option not settable", "Shutdown", "Incorrect user name", "Incorrect password", "Login failure",
        "Invalid operation", "Operation cancelled", "Incorrect handle type", "Incorrect handle state", "Not a proxy",
        "Registry value not found", "Bad registry parameter", "No direct access", "No context", "No callback",
        "Request pending", "Incorrect format", "Item not found", "Cannot connect", "Connection aborted",
        "Connection reset", "Force retry", "Invalid proxy request", "Would block", 0, "Handle exists",
        "Security certificate date invalid", "Security certificate number invalid", "HTTP to HTTPS on redirect",
        "HTTPS to HTTP on redirect", "Mixed security", "Change post is non-secure", "Post is non-secure",
        "Client authorization certificate needed", "Invalid client authorization", "Client authorization not setup",
        "Async thread failed", "Redirect scheme change", "Dialog pending", "Retry dialog", 0,
        "HTTPS HTTP submit redir", "Insert CD-ROM" };
    /* MSINET's string table (35767...) is in the order of the codes, without their gaps: 12045 INVALID_CA is
     * "Invalid client authorization" (CA read as client authorization), 12046 "Client authorization not setup".
     * A bug of VB6 fixed here: MSINET.OCX takes the text 35766 + (code - 12000), one too far from 12035 on
     * (oracle f_inethttps: 12037 "Security certificate number invalid", 12038 "HTTP to HTTPS on redirect",
     * 12045 "Client authorization not setup"); LinBASIC gives the text of the code itself. */
    if (code > 12000 && code - 12000 < (int)(sizeof t1 / sizeof *t1) && t1[code - 12000]) return t1[code - 12000];
    switch (code) {
    case 12110: return "FTP - Transfer in progress";
    case 12111: return "FTP - Connection dropped";
    case 12112: return "FTP - no passive mode";
    case 12130: return "Gopher - Protocol error";
    case 12131: return "Gopher - Not a file";
    case 12132: return "Gopher -  Data error";
    case 12133: return "Gopher - End of data";
    case 12134: return "Gopher - Invalid locator";
    case 12135: return "Gopher - Incorrect locator type";
    case 12136: return "Gopher - Not Gopher Plus";
    case 12137: return "Gopher - Attribute not found";
    case 12138: return "Gopher - Unknown locator";
    case 12150: return "HTTP - Header not found";
    case 12151: return "HTTP - Downlevel server";
    case 12152: return "HTTP - Invalid server response";
    case 12153: return "HTTP - Invalid Header";
    case 12154: return "HTTP - Invalid query request";
    case 12155: return "HTTP - Header already exists";
    case 12156: return "HTTP - Redirect failed";
    case 12157: return "Security channel error";
    case 12158: return "Unable to cache file";
    case 12161: return "HTTP - cookie needs confirmation";
    case 12162: return "HTTP - cookie declined";
    case 12168: return "HTTP - redirect needs confirmation";
    case 12163: return "Disconnected";
    case 12164: return "Server unreachable";
    case 12165: return "Proxy server unreachable";
    case 12166: return "Bad auto proxy script";
    case 12167: return "Unable to download script";
    case 12169: return "Invalid cert";
    case 12170: return "Cert revoked";
    case 12171: return "Failed due to security check";
    }
    return "Extended error";
}
static void fail(lbinet *r, int code, const char *info)
{
    r->ok = 0;
    r->code = code;
    snprintf(r->info, sizeof r->info, "%s", info ? info : inet_text(code));
}

/* ---------------------------------------------------------------- sockets with a time limit and Cancel */
/* 0: ready; 12002 timeout; 12017 cancelled */
static int wait_io(lbinet *r, sock_t s, int wr, int timeout)
{
    long left = timeout > 0 ? timeout * 1000L : -1;
    for (;;) {
        if (r && r->cancel) return 12017;
        int slice = left < 0 || left > 100 ? 100 : (int)left;
#ifdef _WIN32
        fd_set f, x; FD_ZERO(&f); FD_ZERO(&x); FD_SET(s, &f); FD_SET(s, &x);
        struct timeval tv = { 0, slice * 1000 };
        int n = select(0, wr ? 0 : &f, wr ? &f : 0, &x, &tv);
#else
        struct pollfd p = { s, (short)(wr ? POLLOUT : POLLIN), 0 };
        int n = poll(&p, 1, slice);
#endif
        if (n > 0) return 0;
        if (n < 0) return 12030;
        if (left >= 0) { left -= slice; if (left <= 0) return 12002; }
    }
}
static void set_nonblock(sock_t s, int on)
{
#ifdef _WIN32
    u_long v = (u_long)on; ioctlsocket(s, FIONBIO, &v);
#else
    int f = fcntl(s, F_GETFL); fcntl(s, F_SETFL, on ? f | O_NONBLOCK : f & ~O_NONBLOCK);
    fcntl(s, F_SETFD, FD_CLOEXEC);
#endif
}
/* 0 or 12007 (name not resolved), 12029 (cannot connect), 12002, 12017 */
static void ev(lbinet *r, int st);
static int tcp_connect(lbinet *r, const char *host, int port, int timeout, sock_t *out, int ev3)
{
    net_init();
    char ps[16]; snprintf(ps, sizeof ps, "%d", port);
    struct addrinfo hint, *res = 0, *a;
    memset(&hint, 0, sizeof hint);
    hint.ai_family = AF_UNSPEC;
    hint.ai_socktype = SOCK_STREAM;
    if (getaddrinfo(host, ps, &hint, &res) || !res) return 12007;
    if (ev3) ev(r, 3);
    int err = 12029;
    for (a = res; a; a = a->ai_next) {
        sock_t s = socket(a->ai_family, SOCK_STREAM, IPPROTO_TCP);
        if (s == BADS) continue;
        set_nonblock(s, 1);
        int c = connect(s, a->ai_addr, (int)a->ai_addrlen);
        if (c != 0 && neterr() != E_INPROGRESS
#ifndef _WIN32
            && neterr() != EWOULDBLOCK
#endif
            ) { closes(s); continue; }
        if (c != 0) {
            int w = wait_io(r, s, 1, timeout ? timeout : 60);
            int so = 0; socklen_t sl = sizeof so;
            if (w == 0) getsockopt(s, SOL_SOCKET, SO_ERROR, (char *)&so, &sl);
            if (w || so) { closes(s); if (w == 12017 || w == 12002) { err = w; break; } continue; }
        }
        set_nonblock(s, 0);
        *out = s;
        freeaddrinfo(res);
        return 0;
    }
    freeaddrinfo(res);
    return err;
}
static int send_all(lbinet *r, sock_t s, const char *d, size_t n, int timeout)
{
    while (n) {
        int w = wait_io(r, s, 1, timeout);
        if (w) return w;
        int k = (int)send(s, d, (int)(n > 65536 ? 65536 : n), 0);
        if (k <= 0) return 12031;
        d += k; n -= (size_t)k;
    }
    return 0;
}
/* the bytes that come on a connection: *got 0 when it is closed */
static int recv_some(lbinet *r, lbconn *c, char *d, int n, int timeout, int *got)
{
    if (c->rpos < c->rn) {
        int k = c->rn - c->rpos; if (k > n) k = n;
        memcpy(d, c->rb + c->rpos, (size_t)k); c->rpos += k; *got = k; return 0;
    }
    int w = wait_io(r, c->s, 0, timeout);
    if (w) return w;
    int k = (int)recv(c->s, d, n, 0);
    if (k < 0) return 12030;
    *got = k;
    return 0;
}
static int fill(lbinet *r, lbconn *c, int timeout)
{
    if (c->rpos < c->rn) return 0;
    int got = 0, e = recv_some(r, c, c->rb, (int)sizeof c->rb, timeout, &got);
    if (e) return e;
    if (!got) return 12030;
    c->rn = got; c->rpos = 0;
    return 0;
}
/* one line (without CRLF) */
static int read_line(lbinet *r, lbconn *c, buf *line, int timeout)
{
    bclear(line);
    for (;;) {
        int e = fill(r, c, timeout);
        if (e) return e;
        while (c->rpos < c->rn) {
            char ch = c->rb[c->rpos++];
            if (ch == '\n') { if (line->n && line->p[line->n - 1] == '\r') line->p[--line->n] = 0; return 0; }
            bput(line, &ch, 1);
        }
    }
}
static void conn_close(lbconn *c) { closes(c->s); c->s = BADS; c->rn = c->rpos = 0; c->host[0] = 0; }
/* the connections kept for later requests, one for each server, for the whole process (as WinINet:
 * a request to another server does not close them, oracle f_inet) */
static struct { lbconn c; int used; } pool[16];
static mutex_t poolmu; static int poolmu_ok;
static void pool_lock(void) { if (!poolmu_ok) { mu_init(&poolmu); poolmu_ok = 1; } mu_lock(&poolmu); }
static int pool_take(const char *host, int port, lbconn *out)
{
    int r = 0;
    pool_lock();
    for (int i = 0; i < 16; i++)
        if (pool[i].used && pool[i].c.port == port && !strcmp(pool[i].c.host, host)) {
            *out = pool[i].c; pool[i].used = 0; r = 1; break;
        }
    mu_unlock(&poolmu);
    return r;
}
static void pool_put(lbconn *c)
{
    if (c->s == BADS) return;
    pool_lock();
    int slot = -1;
    for (int i = 0; i < 16; i++) {
        if (pool[i].used && pool[i].c.port == c->port && !strcmp(pool[i].c.host, c->host)) { conn_close(&pool[i].c); pool[i].used = 0; }
        if (!pool[i].used && slot < 0) slot = i;
    }
    if (slot < 0) { slot = 0; conn_close(&pool[0].c); }
    pool[slot].c = *c; pool[slot].used = 1;
    mu_unlock(&poolmu);
    c->s = BADS; c->rn = c->rpos = 0; c->host[0] = 0;
}

/* ---------------------------------------------------------------- Basic authentication */
static const char b64[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
static void basic(const char *user, const char *pass, char *out, size_t n)
{
    char in[600]; snprintf(in, sizeof in, "%s:%s", user, pass);
    size_t l = strlen(in), o = 0;
    o += (size_t)snprintf(out, n, "Basic ");
    for (size_t i = 0; i < l && o + 5 < n; i += 3) {
        unsigned v = (unsigned char)in[i] << 16 | (i + 1 < l ? (unsigned char)in[i + 1] << 8 : 0) | (i + 2 < l ? (unsigned char)in[i + 2] : 0);
        out[o++] = b64[v >> 18 & 63]; out[o++] = b64[v >> 12 & 63];
        out[o++] = i + 1 < l ? b64[v >> 6 & 63] : '=';
        out[o++] = i + 2 < l ? b64[v & 63] : '=';
    }
    out[o] = 0;
}
/* WinINet keeps a user and a password that a server accepted for the rest of the process, and sends them
 * to that server from then on (oracle f_inet) */
static struct { char key[300]; char auth[700]; } authc[32];
static mutex_t authmu; static int authmu_ok;
static void auth_lock(void) { if (!authmu_ok) { mu_init(&authmu); authmu_ok = 1; } mu_lock(&authmu); }
static int auth_get(const char *host, int port, char *out, size_t n)
{
    char k[300]; snprintf(k, sizeof k, "%s:%d", host, port);
    int r = 0;
    auth_lock();
    for (int i = 0; i < 32; i++) if (authc[i].key[0] && !strcmp(authc[i].key, k)) { snprintf(out, n, "%s", authc[i].auth); r = 1; break; }
    mu_unlock(&authmu);
    return r;
}
static void auth_put(const char *host, int port, const char *auth)
{
    char k[300]; snprintf(k, sizeof k, "%s:%d", host, port);
    auth_lock();
    int slot = -1;
    for (int i = 0; i < 32; i++) if (!strcmp(authc[i].key, k) || (slot < 0 && !authc[i].key[0])) { slot = i; if (!strcmp(authc[i].key, k)) break; }
    if (slot < 0) slot = 0;
    snprintf(authc[slot].key, sizeof authc[slot].key, "%s", k);
    snprintf(authc[slot].auth, sizeof authc[slot].auth, "%s", auth);
    mu_unlock(&authmu);
}

/* ---------------------------------------------------------------- HTTP (ours) */
static int hdr_value(const char *hdrs, const char *name, char *out, size_t n)
{
    size_t l = strlen(name);
    const char *p = hdrs;
    while (p && *p) {
        const char *e = strstr(p, "\r\n");
        if (!e) e = p + strlen(p);
        if ((size_t)(e - p) > l && !strncasecmp(p, name, l) && p[l] == ':') {
            const char *v = p + l + 1;
            while (v < e && (*v == ' ' || *v == '\t')) v++;
            size_t k = (size_t)(e - v); if (k >= n) k = n - 1;
            memcpy(out, v, k); out[k] = 0;
            return 1;
        }
        p = *e ? e + 2 : e;
    }
    return 0;
}
/* one request on the kept connection (a new one if it is closed): 0 or the error of WinINet */
static int http_once(lbinet *r, const char *method, const char *host, int port, const char *path,
                     const char *auth, int with_data, int *status, buf *hdrs, buf *body)
{
    lbconn *c = r->conn;
    int tries = 0;
again:
    if (c->s != BADS && (strcmp(c->host, host) || c->port != port)) pool_put(c);
    if (c->s == BADS) pool_take(host, port, c);
    if (c->s == BADS) {
        conn_close(c);
        int e = tcp_connect(r, host, port, r->timeout, &c->s, 1);
        if (e) { c->s = BADS; return e; }
        snprintf(c->host, sizeof c->host, "%s", host);
        c->port = port;
        ev(r, 4);
        tries = 1;
    }
    buf q = { 0 };
    char line[3000];
    snprintf(line, sizeof line, "%s %s HTTP/1.1\r\n", method, *path ? path : "/"); bputs(&q, line);
    bputs(&q, r->openurl ? "Accept: image/gif,image/x-xbitmap,image/jpeg,image/pjpeg,*/*\r\n"
                         : "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*\r\n");
    if (r->hdrs.n) {
        bput(&q, r->hdrs.p, r->hdrs.n);
        if (r->hdrs.n < 2 || memcmp(r->hdrs.p + r->hdrs.n - 2, "\r\n", 2)) bputs(&q, "\r\n");
    }
    bputs(&q, "User-Agent: " UA "\r\n");
    if (port == 80) snprintf(line, sizeof line, "Host: %s\r\n", host);
    else snprintf(line, sizeof line, "Host: %s:%d\r\n", host, port);
    bputs(&q, line);
    int has_body = with_data || !strcasecmp(method, "POST") || !strcasecmp(method, "PUT") || !strcasecmp(method, "HEAD");
    size_t dn = with_data ? r->data.n : 0;
    if (has_body) { snprintf(line, sizeof line, "Content-Length: %u\r\n", (unsigned)dn); bputs(&q, line); }
    if (r->openurl) bputs(&q, "Connection: Keep-Alive\r\n");
    bputs(&q, "Cache-Control: no-cache\r\n");
    if (auth && *auth) { snprintf(line, sizeof line, "Authorization: %s\r\n", auth); bputs(&q, line); }
    bputs(&q, "\r\n");
    if (dn) bput(&q, r->data.p, dn);
    ev(r, 5);
    int e = send_all(r, c->s, q.p, q.n, r->timeout);
    bfree(&q);
    if (e == 12017) return e;
    if (e) { conn_close(c); if (!tries++) goto again; return e; }
    ev(r, 6);
    ev(r, 7);
    /* the answer: the status line and the headers */
    buf l = { 0 };
    bclear(hdrs);
    for (;;) {
        e = read_line(r, c, &l, r->timeout);
        if (e) {
            bfree(&l);
            if (e == 12030 && !hdrs->n) { conn_close(c); if (!tries++) goto again; }
            conn_close(c);
            return e == 12030 && !hdrs->n ? 12152 : e;
        }
        if (!hdrs->n && !l.n) continue;
        bput(hdrs, l.p ? l.p : "", l.n); bputs(hdrs, "\r\n");
        if (!l.n) {
            /* 100 Continue: the real answer follows */
            if (!strncmp(hdrs->p, "HTTP/1.1 100", 12) || !strncmp(hdrs->p, "HTTP/1.0 100", 12)) { bclear(hdrs); continue; }
            break;
        }
    }
    bfree(&l);
    if (strncmp(hdrs->p, "HTTP/", 5)) { conn_close(c); return 12152; }
    *status = atoi(strchr(hdrs->p, ' ') ? strchr(hdrs->p, ' ') + 1 : "0");
    /* the body */
    char v[256];
    int keep = !strncmp(hdrs->p, "HTTP/1.1", 8);
    if (hdr_value(hdrs->p, "Connection", v, sizeof v)) keep = !strcasecmp(v, "keep-alive") ? 1 : strcasecmp(v, "close") ? keep : 0;
    bclear(body);
    char tmp[16384]; int got;
    if (!strcasecmp(method, "HEAD") || *status == 204 || *status == 304 || (*status >= 100 && *status < 200)) {
    } else if (hdr_value(hdrs->p, "Transfer-Encoding", v, sizeof v) && strstr(v, "chunked")) {
        buf cl = { 0 };
        for (;;) {
            if ((e = read_line(r, c, &cl, r->timeout))) { bfree(&cl); conn_close(c); return e; }
            long n = strtol(cl.p ? cl.p : "0", 0, 16);
            if (n <= 0) {                     /* the trailer */
                do { if ((e = read_line(r, c, &cl, r->timeout))) break; } while (cl.n);
                break;
            }
            while (n > 0) {
                if ((e = recv_some(r, c, tmp, n > (long)sizeof tmp ? (int)sizeof tmp : (int)n, r->timeout, &got)) || !got) {
                    bfree(&cl); conn_close(c); return e ? e : 12030;
                }
                bput(body, tmp, (size_t)got); n -= got;
                ev(r, 7);
            }
            read_line(r, c, &cl, r->timeout);
        }
        bfree(&cl);
    } else if (hdr_value(hdrs->p, "Content-Length", v, sizeof v)) {
        long n = atol(v);
        while (n > 0) {
            if ((e = recv_some(r, c, tmp, n > (long)sizeof tmp ? (int)sizeof tmp : (int)n, r->timeout, &got)) || !got) {
                conn_close(c); return e ? e : 12030;
            }
            bput(body, tmp, (size_t)got); n -= got;
        }
    } else {                                  /* until the server closes the connection */
        for (;;) {
            if ((e = recv_some(r, c, tmp, (int)sizeof tmp, r->timeout, &got))) { conn_close(c); return e; }
            if (!got) break;
            bput(body, tmp, (size_t)got);
        }
        keep = 0;
    }
    ev(r, 8);
    if (!keep) conn_close(c);
    return 0;
}
/* a URL of a Location header -> host, port, path (relative to the request) */
static int split_location(const char *loc, char *host, size_t hn, int *port, char *path, size_t pn, int *https)
{
    if (!strncasecmp(loc, "http://", 7) || !strncasecmp(loc, "https://", 8)) {
        int s = loc[4] == 's' || loc[4] == 'S';
        const char *p = loc + (s ? 8 : 7);
        const char *slash = strchr(p, '/');
        size_t al = slash ? (size_t)(slash - p) : strlen(p);
        char auth[512]; if (al >= sizeof auth) al = sizeof auth - 1;
        memcpy(auth, p, al); auth[al] = 0;
        char *at = strrchr(auth, '@');
        char *hp = at ? at + 1 : auth;
        char *colon = strrchr(hp, ':');
        *port = s ? 443 : 80;
        if (colon && !strchr(colon, ']')) { *port = atoi(colon + 1); *colon = 0; }
        snprintf(host, hn, "%.250s", hp);
        snprintf(path, pn, "%s", slash ? slash : "/");
        *https = s;
        return 1;
    }
    if (loc[0] == '/') { snprintf(path, pn, "%s", loc); return 1; }
    /* relative to the folder of the request */
    char base[2048]; snprintf(base, sizeof base, "%s", path);
    char *q = strchr(base, '?'); if (q) *q = 0;
    char *ls = strrchr(base, '/'); if (ls) ls[1] = 0; else snprintf(base, sizeof base, "/");
    char nb[4096]; snprintf(nb, sizeof nb, "%s%s", base, loc);
    snprintf(path, pn, "%.2000s", nb);
    return 1;
}
static int https_request(lbinet *r, const char *method, const char *host, int port, const char *path, int with_data,
                         const char *auth);
static void http_job(lbinet *r)
{
    char host[256], path[2048], method[32], cached[700] = "", uauth[700] = "";
    int port = r->port, https = r->https, with_data = r->data.n > 0 || !strcasecmp(r->method, "POST") || !strcasecmp(r->method, "PUT");
    snprintf(host, sizeof host, "%s", r->host);
    snprintf(path, sizeof path, "%s", r->path);
    snprintf(method, sizeof method, "%s", r->method);
    if (r->user[0]) basic(r->user, r->pass, uauth, sizeof uauth);
    int tried = 0;
    for (int hop = 0; hop < 12; hop++) {
        if (https) {                    /* libcurl: the Basic authentication as WinINet does it (http:// below) */
            const char *auth = uauth[0] ? "" : auth_get(host, port, cached, sizeof cached) ? cached : "";
            if (tried) auth = uauth;
            int e = https_request(r, method, host, port, path, with_data, auth);
            if (e) {
                /* the certificate / TLS refused: VB6 (oracle f_inethttps) has 3 4 9 10 11, the connection closed */
                if (e == 12037 || e == 12038 || e == 12044 || e == 12045 || e == 12055 || e == 12157 || e == 12169 ||
                    e == 12170) { ev(r, 9); ev(r, 10); }
                if (e != 12017) fail(r, e, 0);
                return;
            }
            if (r->status == 401 && uauth[0] && !tried) { tried = 1; continue; }
            if (uauth[0] && r->status != 401 && tried) auth_put(host, port, uauth);
            r->ok = 1;
            return;
        }
        /* a user in the URL: WinINet asks without it first, then with it (not with what it keeps) */
        const char *auth = uauth[0] ? "" : auth_get(host, port, cached, sizeof cached) ? cached : "";
        if (tried) auth = uauth;
        int status = 0;
        int e = http_once(r, method, host, port, path, auth, with_data, &status, &r->headers, &r->body);
        if (e) { if (e != 12017) fail(r, e, 0); return; }
        r->status = status;
        char loc[2048];
        if ((status == 301 || status == 302 || status == 303 || status == 307 || status == 308) &&
            hdr_value(r->headers.p, "Location", loc, sizeof loc)) {
            int np = port, nh = https;
            char nhost[256]; snprintf(nhost, sizeof nhost, "%s", host);
            split_location(loc, nhost, sizeof nhost, &np, path, sizeof path, &nh);
            snprintf(host, sizeof host, "%s", nhost);
            port = np; https = nh;
            r->redirects++;
            if (port == (https ? 443 : 80)) snprintf(r->finalurl, sizeof r->finalurl, "%s://%s%s", https ? "https" : "http", host, path);
            else snprintf(r->finalurl, sizeof r->finalurl, "%s://%s:%d%s", https ? "https" : "http", host, port, path);
            if (status == 303 || ((status == 301 || status == 302) && strcasecmp(method, "GET") && strcasecmp(method, "HEAD"))) {
                snprintf(method, sizeof method, "GET");
                with_data = 0;
            }
            continue;
        }
        if (status == 401 && uauth[0] && !tried) { tried = 1; continue; }
        if (status != 401 && tried) auth_put(host, port, uauth);
        break;
    }
    r->ok = 1;
}

/* ---------------------------------------------------------------- HTTPS: libcurl (Linux), WinINet (Windows) */
#ifndef _WIN32
typedef void CURL;
struct curl_slist;
static struct {
    int tried, ok;
    CURL *(*init)(void);
    int (*setopt)(CURL *, int, ...);
    int (*perform)(CURL *);
    void (*cleanup)(CURL *);
    int (*getinfo)(CURL *, int, ...);
    struct curl_slist *(*append)(struct curl_slist *, const char *);
    void (*slist_free)(struct curl_slist *);
    int (*global_init)(long);
    void *(*share_init)(void);
    int (*share_setopt)(void *, int, ...);
    void *share;                        /* the connections, TLS sessions, names: kept for the whole program */
} cu;
static mutex_t cushmu[8];
static void cu_lock(CURL *h, int data, int access, void *u)
{
    (void)h; (void)access; (void)u;
    mu_lock(&cushmu[data & 7]);
}
static void cu_unlock(CURL *h, int data, void *u)
{
    (void)h; (void)u;
    mu_unlock(&cushmu[data & 7]);
}
static mutex_t cumu; static int cumu_ok;
static int curl_load(void)
{
    if (!cumu_ok) { mu_init(&cumu); cumu_ok = 1; }
    mu_lock(&cumu);
    if (!cu.tried) {
        cu.tried = 1;
        const char *names[] = { "libcurl.so.4", "libcurl-gnutls.so.4", "libcurl-nss.so.4", "libcurl.so" };
        void *h = 0;
        for (unsigned i = 0; i < sizeof names / sizeof *names && !h; i++) h = dlopen(names[i], RTLD_NOW);
        if (h) {
            cu.init = (CURL *(*)(void))dlsym(h, "curl_easy_init");
            cu.setopt = (int (*)(CURL *, int, ...))dlsym(h, "curl_easy_setopt");
            cu.perform = (int (*)(CURL *))dlsym(h, "curl_easy_perform");
            cu.cleanup = (void (*)(CURL *))dlsym(h, "curl_easy_cleanup");
            cu.getinfo = (int (*)(CURL *, int, ...))dlsym(h, "curl_easy_getinfo");
            cu.append = (struct curl_slist *(*)(struct curl_slist *, const char *))dlsym(h, "curl_slist_append");
            cu.slist_free = (void (*)(struct curl_slist *))dlsym(h, "curl_slist_free_all");
            cu.global_init = (int (*)(long))dlsym(h, "curl_global_init");
            cu.ok = cu.init && cu.setopt && cu.perform && cu.cleanup && cu.getinfo && cu.append && cu.slist_free;
            if (cu.ok && cu.global_init) cu.global_init(3);
            cu.share_init = (void *(*)(void))dlsym(h, "curl_share_init");
            cu.share_setopt = (int (*)(void *, int, ...))dlsym(h, "curl_share_setopt");
            if (cu.ok && cu.share_init && cu.share_setopt && (cu.share = cu.share_init()) != 0) {
                for (int i = 0; i < 8; i++) mu_init(&cushmu[i]);
                cu.share_setopt(cu.share, 3, cu_lock);           /* CURLSHOPT_LOCKFUNC */
                cu.share_setopt(cu.share, 4, cu_unlock);         /* CURLSHOPT_UNLOCKFUNC */
                cu.share_setopt(cu.share, 1, 3);                 /* SHARE: DNS */
                cu.share_setopt(cu.share, 1, 4);                 /* SSL_SESSION */
                cu.share_setopt(cu.share, 1, 5);                 /* CONNECT (7.57) */
            }
        }
    }
    mu_unlock(&cumu);
    return cu.ok;
}
static size_t cu_body(char *d, size_t a, size_t b, void *u)
{
    lbinet *r = (lbinet *)u;
    if (r->cancel) return 0;
    bput(&r->body, d, a * b);
    ev(r, 7);
    return a * b;
}
static size_t cu_hdr(char *d, size_t a, size_t b, void *u)
{
    lbinet *r = (lbinet *)u;
    size_t n = a * b;
    if (n >= 5 && !memcmp(d, "HTTP/", 5)) { bclear(&r->headers); bclear(&r->body); }
    bput(&r->headers, d, n);
    return n;
}
/* a socket for a new connection (before connect): icConnecting */
static int cu_sockopt(void *u, int fd, int purpose)
{
    lbinet *r = (lbinet *)u;
    (void)fd; (void)purpose;
    if (!r->cnew) { r->cnew = 1; ev(r, 3); }
    return 0;
}
/* connected (and TLS done), the request goes now (7.80): icConnected when the connection is new, icRequesting */
static int cu_prereq(void *u, char *ip, char *lip, int port, int lport)
{
    lbinet *r = (lbinet *)u;
    (void)ip; (void)lip; (void)port; (void)lport;
    if (r->cnew && !r->cpre) ev(r, 4);
    if (!r->cpre) ev(r, 5);
    r->cpre = 1;
    return 0;
}
static int cu_progress(void *u, long long dt, long long dn, long long ut, long long un)
{
    (void)dt; (void)dn; (void)ut; (void)un;
    return ((lbinet *)u)->cancel ? 1 : 0;
}
static int https_request(lbinet *r, const char *method, const char *host, int port, const char *path, int with_data,
                         const char *auth)
{
    if (!curl_load()) return 12157;          /* no libcurl on this computer: no TLS */
    CURL *c = cu.init();
    if (!c) return 12004;
    char url[4096];
    snprintf(url, sizeof url, "https://%s:%d%s", host, port, *path ? path : "/");
    struct curl_slist *h = 0;
    h = cu.append(h, r->openurl ? "Accept: image/gif,image/x-xbitmap,image/jpeg,image/pjpeg,*/*"
                                : "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*");
    if (r->hdrs.n) {
        char *copy = strdup(r->hdrs.p), *save = 0;
        for (char *l = strtok_r(copy, "\r\n", &save); l; l = strtok_r(0, "\r\n", &save)) if (*l) h = cu.append(h, l);
        free(copy);
    }
    if (r->openurl) h = cu.append(h, "Connection: Keep-Alive");  /* as VB6's requests (oracle f_inethttps) */
    if (!strcasecmp(method, "HEAD")) h = cu.append(h, "Content-Length: 0");
    h = cu.append(h, "Cache-Control: no-cache");
    h = cu.append(h, "Expect:");
    if (auth && *auth) { char a[800]; snprintf(a, sizeof a, "Authorization: %s", auth); h = cu.append(h, a); }
    cu.setopt(c, 10002, url);                            /* CURLOPT_URL */
    cu.setopt(c, 10018, UA);                             /* USERAGENT */
    cu.setopt(c, 10023, h);                              /* HTTPHEADER */
    cu.setopt(c, 20011, cu_body); cu.setopt(c, 10001, r);  /* WRITEFUNCTION, WRITEDATA */
    cu.setopt(c, 20079, cu_hdr); cu.setopt(c, 10029, r);   /* HEADERFUNCTION, HEADERDATA */
    cu.setopt(c, 20219, cu_progress); cu.setopt(c, 10057, r); cu.setopt(c, 43, 0L);   /* XFERINFO, NOPROGRESS 0 */
    cu.setopt(c, 52, 1L); cu.setopt(c, 68, 10L);         /* FOLLOWLOCATION, MAXREDIRS */
    cu.setopt(c, 99, 1L);                                /* NOSIGNAL */
    cu.setopt(c, 84, 2L);                                /* HTTP_VERSION 1.1: as WinINet (no HTTP/2 in GetHeader) */
    char eb[256] = "";
    cu.setopt(c, 10010, eb);                             /* ERRORBUFFER */
    if (r->timeout > 0) { cu.setopt(c, 13, (long)r->timeout); cu.setopt(c, 78, (long)r->timeout); }
    const char *ca = getenv("VBRT_CA_FILE");             /* certificates trusted besides the system's (tests) */
    if (ca && *ca) cu.setopt(c, 10065, ca);              /* CAINFO */
    if (r->access == 1) cu.setopt(c, 10004, "");         /* PROXY: none */
    else if (r->access == 2 && r->proxy[0]) cu.setopt(c, 10004, r->proxy);
    if (!strcasecmp(method, "HEAD")) cu.setopt(c, 44, 1L);   /* NOBODY */
    else if (strcasecmp(method, "GET")) {
        cu.setopt(c, 10036, method);                     /* CUSTOMREQUEST */
        if (with_data) { cu.setopt(c, 10015, r->data.p ? r->data.p : ""); cu.setopt(c, 60, (long)r->data.n); }
    }
    if (cu.share) cu.setopt(c, 10100, cu.share);         /* SHARE */
    r->cnew = r->cpre = 0;
    cu.setopt(c, 20148, cu_sockopt); cu.setopt(c, 10149, r);   /* SOCKOPTFUNCTION, SOCKOPTDATA */
    int pre = cu.setopt(c, 20312, cu_prereq) == 0;       /* PREREQFUNCTION */
    if (pre) cu.setopt(c, 10313, r);                     /* PREREQDATA */
    int e = cu.perform(c);
    long st = 0;
    cu.getinfo(c, 0x200000 + 2, &st);                   /* RESPONSE_CODE */
    long nred = 0;                                      /* libcurl follows the redirects: AsyncRead needs the URL */
    if (e == 0 && cu.getinfo(c, 0x200000 + 20, &nred) == 0 && nred > 0) {   /* REDIRECT_COUNT */
        const char *eu = 0;
        if (cu.getinfo(c, 0x100000 + 1, &eu) == 0 && eu) {                 /* EFFECTIVE_URL */
            r->redirects += (int)nred;
            snprintf(r->finalurl, sizeof r->finalurl, "%s", eu);
        }
    }
    long vr = 0;
    if (e == 51 || e == 60) cu.getinfo(c, 0x200000 + 13, &vr);   /* SSL_VERIFYRESULT (OpenSSL: X509_V_ERR_...) */
    cu.cleanup(c);
    cu.slist_free(h);
    if (r->cancel) return 12017;
    if (e == 0 && !r->cpre) { if (r->cnew) ev(r, 4); ev(r, 5); }   /* a libcurl older than 7.80 */
    if (e && r->cnew && !r->cpre && e != 7 && e != 28) ev(r, 4); /* connected, TLS failed (WinINet: 3 4 11) */
    switch (e) {
    case 0: break;
    case 1: return 12006;                               /* unsupported protocol */
    case 3: return 12005;                               /* bad URL */
    case 5: case 6: return 12007;                       /* proxy / host not resolved */
    case 7: return 12029;
    case 28: return 12002;
    case 35: case 53: case 54: case 58: case 59: case 66: case 80: case 83: return 12157;
    case 51: case 60:                                   /* the certificate: as WinINet tells the three cases */
        if (vr == 9 || vr == 10 || strstr(eb, "expired") || strstr(eb, "not yet valid")) return 12037;  /* DATE_INVALID */
        if (vr == 62 || strstr(eb, "subject name") || strstr(eb, "host name") || strstr(eb, "hostname"))
            return 12038;                               /* SEC_CERT_CN_INVALID */
        return 12045;                                   /* INVALID_CA */
    case 77: return 12157;
    case 52: return 12152;
    case 55: return 12031;
    case 56: return 12030;
    case 67: return 12015;
    default: return 12029;
    }
    r->status = (int)st;
    ev(r, 6); ev(r, 7); ev(r, 8);
    /* the headers as WinINet gives them: CRLF lines, an empty line at the end */
    return 0;
}
#else
/* WinINet trusts the certificates of the store of Windows only (the tests put theirs there: tests/dev/run_tests.py
 * for Wine, the Trusted Root Certification Authorities of the VM); VBRT_CA_FILE is for libcurl (Linux) */
static void CALLBACK wi_status(HINTERNET h, DWORD_PTR ctx, DWORD st, LPVOID info, DWORD n)
{
    lbinet *r = (lbinet *)ctx;
    (void)h; (void)info; (void)n;
    if (!r) return;
    if (st == INTERNET_STATUS_CONNECTING_TO_SERVER) ev(r, 3);
    else if (st == INTERNET_STATUS_CONNECTED_TO_SERVER) ev(r, 4);
    else if (st == INTERNET_STATUS_SENDING_REQUEST && !r->cpre) { r->cpre = 1; ev(r, 5); }
    else if (st == INTERNET_STATUS_REDIRECT && info) {     /* WinINet follows them itself: AsyncRead needs the URL */
        r->redirects++;
        snprintf(r->finalurl, sizeof r->finalurl, "%s", (const char *)info);
    }
}
/* the sessions of WinINet, kept for the whole program (by AccessType and Proxy): their connections stay open */
static HINTERNET wi_session(int access, const char *proxy)
{
    static struct { int access; char proxy[512]; HINTERNET h; } ss[8];
    static int n;
    static mutex_t m; static volatile LONG m_ok;
    if (InterlockedCompareExchange(&m_ok, 1, 0) == 0) { mu_init(&m); m_ok = 2; }
    while (m_ok != 2) Sleep(1);
    const char *px = access == 2 && proxy[0] ? proxy : "";
    HINTERNET h = NULL;
    mu_lock(&m);
    for (int i = 0; i < n && !h; i++) if (ss[i].access == access && !strcmp(ss[i].proxy, px)) h = ss[i].h;
    if (!h) {
        DWORD type = access == 1 ? INTERNET_OPEN_TYPE_DIRECT : access == 2 ? INTERNET_OPEN_TYPE_PROXY : INTERNET_OPEN_TYPE_PRECONFIG;
        h = InternetOpenA(UA, type, *px ? px : NULL, NULL, 0);
        if (h) {
            InternetSetStatusCallbackA(h, wi_status);
            if (n < 8) { ss[n].access = access; snprintf(ss[n].proxy, sizeof ss[n].proxy, "%s", px); ss[n].h = h; n++; }
        }
    }
    mu_unlock(&m);
    return h;
}
static int https_request(lbinet *r, const char *method, const char *host, int port, const char *path, int with_data,
                         const char *auth)
{
    (void)auth;                         /* WinINet remembers the authentication itself */
    HINTERNET hi = wi_session(r->access, r->proxy);
    if (!hi) return (int)GetLastError();
    r->cnew = r->cpre = 0;
    HINTERNET hc = InternetConnectA(hi, host, (INTERNET_PORT)port, r->user[0] ? r->user : NULL, r->user[0] ? r->pass : NULL,
                                    INTERNET_SERVICE_HTTP, 0, (DWORD_PTR)r);
    if (!hc) return (int)GetLastError();
    if (r->timeout > 0) {
        DWORD ms = (DWORD)r->timeout * 1000;
        InternetSetOptionA(hc, INTERNET_OPTION_CONNECT_TIMEOUT, &ms, sizeof ms);
        InternetSetOptionA(hc, INTERNET_OPTION_SEND_TIMEOUT, &ms, sizeof ms);
        InternetSetOptionA(hc, INTERNET_OPTION_RECEIVE_TIMEOUT, &ms, sizeof ms);
    }
    const char *acc[] = { "image/gif", "image/x-xbitmap", "image/jpeg", "image/pjpeg", "*/*", NULL };
    const char *acc1[] = { "image/gif,image/x-xbitmap,image/jpeg,image/pjpeg,*/*", NULL };  /* OpenURL: no spaces */
    HINTERNET hr = HttpOpenRequestA(hc, method, *path ? path : "/", "HTTP/1.1", NULL, r->openurl ? acc1 : acc,
                                    INTERNET_FLAG_SECURE | INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE |
                                    INTERNET_FLAG_KEEP_CONNECTION, (DWORD_PTR)r);
    if (!hr) { int e = (int)GetLastError(); InternetCloseHandle(hc); return e; }
    buf hd = { 0 };
    bputs(&hd, "Cache-Control: no-cache\r\n");
    if (r->hdrs.n) { bput(&hd, r->hdrs.p, r->hdrs.n); if (r->hdrs.n < 2 || memcmp(r->hdrs.p + r->hdrs.n - 2, "\r\n", 2)) bputs(&hd, "\r\n"); }
    BOOL ok = HttpSendRequestA(hr, hd.p, (DWORD)hd.n, with_data ? r->data.p : NULL, with_data ? (DWORD)r->data.n : 0);
    int e = ok ? 0 : (int)GetLastError();
    bfree(&hd);
    if (ok) {
        if (!r->cpre) ev(r, 5);
        ev(r, 6); ev(r, 7);
        char tmp[16384]; DWORD n = sizeof tmp;
        if (HttpQueryInfoA(hr, HTTP_QUERY_RAW_HEADERS_CRLF, tmp, &n, NULL)) { bclear(&r->headers); bput(&r->headers, tmp, n); }
        DWORD sc = 0, sl = sizeof sc;
        if (HttpQueryInfoA(hr, HTTP_QUERY_STATUS_CODE | HTTP_QUERY_FLAG_NUMBER, &sc, &sl, NULL)) r->status = (int)sc;
        for (;;) {
            if (r->cancel) { e = 12017; break; }
            DWORD got = 0;
            if (!InternetReadFile(hr, tmp, sizeof tmp, &got)) { e = (int)GetLastError(); break; }
            if (!got) break;
            bput(&r->body, tmp, got);
        }
        if (!e) ev(r, 8);
    }
    InternetCloseHandle(hr); InternetCloseHandle(hc);
    return e;
}
#endif

/* ---------------------------------------------------------------- FTP (ours) */
static void reply_add(lbinet *r, const char *s)
{
    size_t l = strlen(r->info);
    snprintf(r->info + l, sizeof r->info - l, "%s\r\n", s);
}
/* the reply of the server (the lines of a multi-line one); its code, or -1 when the connection failed */
static int ftp_reply(lbinet *r, lbftp *f, buf *text)
{
    buf l = { 0 };
    bclear(text);
    if (read_line(r, &f->c, &l, r->timeout)) { bfree(&l); return -1; }
    bput(text, l.p ? l.p : "", l.n);
    int code = atoi(l.p ? l.p : "0");
    if (l.n > 3 && l.p[3] == '-') {
        char end[5]; snprintf(end, sizeof end, "%.3s ", l.p);
        for (;;) {
            if (read_line(r, &f->c, &l, r->timeout)) { bfree(&l); return -1; }
            bputs(text, "\r\n"); bput(text, l.p ? l.p : "", l.n);
            if (l.n >= 4 && !memcmp(l.p, end, 4)) break;
        }
    }
    bfree(&l);
    if (r) reply_add(r, text->p ? text->p : "");
    return code;
}
static int ftp_cmd(lbinet *r, lbftp *f, const char *cmd, buf *text)
{
    char line[2100]; snprintf(line, sizeof line, "%s\r\n", cmd);
    if (send_all(r, f->c.s, line, strlen(line), r->timeout)) return -1;
    return ftp_reply(r, f, text);
}
static int is_private(const struct sockaddr_storage *a)
{
    if (a->ss_family != AF_INET) return 1;
    unsigned v = ntohl(((const struct sockaddr_in *)a)->sin_addr.s_addr);
    return (v >> 24) == 10 || (v >> 24) == 127 || (v >> 20) == 0xAC1 || (v >> 16) == 0xC0A8 || (v >> 16) == 0xA9FE;
}
/* connect and log in: 0, or the error of WinINet */
int lbftp_open(lbftp *f, const char *host, int port, const char *user, const char *pass, int timeout)
{
    lbinet tmp; memset(&tmp, 0, sizeof tmp); tmp.timeout = timeout;
    lbinet *r = &tmp;
    conn_close(&f->c);
    int e = tcp_connect(0, host, port, timeout, &f->c.s, 0);
    if (e) { f->c.s = BADS; return e; }
    snprintf(f->c.host, sizeof f->c.host, "%s", host);
    f->c.port = port;
    snprintf(f->user, sizeof f->user, "%s", user && *user ? user : "anonymous");
    snprintf(f->pass, sizeof f->pass, "%s", user && *user ? (pass ? pass : "") : "IEUser@");
    f->cwd_known = 0;
    buf t = { 0 };
    int c = ftp_reply(r, f, &t);
    char cmd[600];
    if (c / 100 == 2) {
        snprintf(cmd, sizeof cmd, "USER %s", f->user);
        c = ftp_cmd(r, f, cmd, &t);
        if (c == 331) { snprintf(cmd, sizeof cmd, "PASS %s", f->pass); c = ftp_cmd(r, f, cmd, &t); }
    }
    bfree(&t);
    if (c / 100 != 2) { conn_close(&f->c); return c < 0 ? 12030 : 12015; }
    struct sockaddr_storage me, peer; socklen_t l1 = sizeof me, l2 = sizeof peer;
    getsockname(f->c.s, (struct sockaddr *)&me, &l1);
    getpeername(f->c.s, (struct sockaddr *)&peer, &l2);
    f->passive = is_private(&me) && !is_private(&peer);
    return 0;
}
int lbftp_isopen(lbftp *f) { return f && f->c.s != BADS; }
int lbftp_same(lbftp *f, const char *host, int port, const char *user)
{
    const char *u = user && *user ? user : "anonymous";
    return lbftp_isopen(f) && !strcmp(f->c.host, host) && f->c.port == port && !strcmp(f->user, u);
}
lbftp *lbftp_new(void) { lbftp *f = (lbftp *)calloc(1, sizeof *f); f->c.s = BADS; return f; }
void lbftp_close(lbftp *f, int quit)
{
    if (!f || f->c.s == BADS) return;
    if (quit) { const char *q = "QUIT\r\n"; send(f->c.s, q, (int)strlen(q), 0); }
    conn_close(&f->c);
}
void lbftp_free(lbftp *f) { if (f) { lbftp_close(f, 1); free(f); } }

/* a data connection: active (PORT, as WinINet) or passive; *ls is the listening socket of PORT */
static int ftp_data_open(lbinet *r, lbftp *f, sock_t *ls, sock_t *ds)
{
    buf t = { 0 };
    *ls = *ds = BADS;
    if (!f->passive) {
        struct sockaddr_in me; socklen_t ml = sizeof me;
        if (getsockname(f->c.s, (struct sockaddr *)&me, &ml) == 0 && me.sin_family == AF_INET) {
            sock_t s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
            me.sin_port = 0;
            if (s != BADS && !bind(s, (struct sockaddr *)&me, sizeof me) && !listen(s, 1)) {
                socklen_t l = sizeof me;
                getsockname(s, (struct sockaddr *)&me, &l);
                unsigned a = ntohl(me.sin_addr.s_addr), p = ntohs(me.sin_port);
                char cmd[100];
                snprintf(cmd, sizeof cmd, "PORT %u,%u,%u,%u,%u,%u", a >> 24, a >> 16 & 255, a >> 8 & 255, a & 255, p >> 8, p & 255);
                int c = ftp_cmd(r, f, cmd, &t);
                if (c / 100 == 2) { *ls = s; bfree(&t); return 0; }
                if (c < 0) { closes(s); bfree(&t); return 12030; }
            }
            closes(s);
        }
    }
    int c = ftp_cmd(r, f, "PASV", &t);
    if (c != 227) { bfree(&t); return c < 0 ? 12030 : 12003; }
    const char *p = strchr(t.p, '(');
    unsigned h1, h2, h3, h4, p1, p2;
    if (!p || sscanf(p + 1, "%u,%u,%u,%u,%u,%u", &h1, &h2, &h3, &h4, &p1, &p2) != 6) { bfree(&t); return 12003; }
    bfree(&t);
    char ip[64]; snprintf(ip, sizeof ip, "%u.%u.%u.%u", h1, h2, h3, h4);
    /* a server that names an address of its own network: the address of the control connection */
    struct sockaddr_storage peer; socklen_t pl = sizeof peer;
    if (getpeername(f->c.s, (struct sockaddr *)&peer, &pl) == 0 && peer.ss_family == AF_INET && !is_private(&peer)) {
        struct sockaddr_in pa; memset(&pa, 0, sizeof pa); pa.sin_family = AF_INET;
        inet_pton(AF_INET, ip, &pa.sin_addr);
        struct sockaddr_storage ps; memcpy(&ps, &pa, sizeof pa);
        if (is_private(&ps)) inet_ntop(AF_INET, &((struct sockaddr_in *)&peer)->sin_addr, ip, sizeof ip);
    }
    return tcp_connect(r, ip, (int)(p1 * 256 + p2), r->timeout, ds, 0);
}
/* after the 1xx reply: the data socket (accepted for PORT) */
static int sock_readable(sock_t s, int ms)
{
#ifdef _WIN32
    fd_set f; FD_ZERO(&f); FD_SET(s, &f);
    struct timeval tv = { 0, ms * 1000 };
    return select(0, &f, 0, 0, &tv) > 0;
#else
    struct pollfd p = { s, POLLIN, 0 };
    return poll(&p, 1, ms) > 0;
#endif
}
/* PORT: the server connects to us, or answers on the control connection that it cannot (425: a firewall) */
static int ftp_data_ready(lbinet *r, lbftp *f, sock_t *ls, sock_t *ds)
{
    if (*ds != BADS) return 0;
    long left = (r->timeout ? r->timeout : 60) * 1000L;
    int e = 0;
    while (!e) {
        if (r->cancel) { e = 12017; break; }
        if (sock_readable(*ls, 100)) break;
        if (f->c.rpos < f->c.rn || sock_readable(f->c.s, 0)) {
            buf t = { 0 };
            int c = ftp_reply(r, f, &t);
            bfree(&t);
            if (c < 0) e = 12030;
            else if (c / 100 >= 4) e = 12003;
        }
        if (!e && (left -= 100) <= 0) e = 12002;
    }
    if (e) { closes(*ls); *ls = BADS; return e; }
    *ds = accept(*ls, 0, 0);
    closes(*ls); *ls = BADS;
    return *ds == BADS ? 12030 : 0;
}
/* a transfer: cmd (LIST, RETR x, STOR x); download into out, or upload up */
static int ftp_transfer(lbinet *r, lbftp *f, const char *type, const char *cmd, buf *out, const buf *up)
{
    buf t = { 0 };
    char tc[16]; snprintf(tc, sizeof tc, "TYPE %s", type);
    int c = ftp_cmd(r, f, tc, &t);
    if (c < 0) { bfree(&t); return 12030; }
    sock_t ls, ds;
    int e = ftp_data_open(r, f, &ls, &ds);
    if (e) { bfree(&t); return e; }
    c = ftp_cmd(r, f, cmd, &t);
    if (c / 100 != 1) { closes(ls); closes(ds); bfree(&t); return c < 0 ? 12030 : 12003; }
    e = ftp_data_ready(r, f, &ls, &ds);
    if (e) { closes(ls); closes(ds); bfree(&t); return e; }
    lbconn dc; memset(&dc, 0, sizeof dc); dc.s = ds;
    if (up) {
        e = send_all(r, ds, up->p ? up->p : "", up->n, r->timeout);
        ev(r, 7);
    } else {
        char tmp[16384]; int got;
        for (;;) {
            if ((e = recv_some(r, &dc, tmp, (int)sizeof tmp, r->timeout, &got))) break;
            if (!got) break;
            bput(out, tmp, (size_t)got);
            ev(r, 7);
        }
    }
    closes(ds);
    if (e) { bfree(&t); return e; }
    c = ftp_reply(r, f, &t);
    bfree(&t);
    return c / 100 == 2 ? 0 : c < 0 ? 12030 : 12003;
}
/* a listing of LIST -> names, folders with "/" (as FtpFindFirstFile gives them); size of name */
static int list_names(const char *text, const char *pattern, buf *out, char *size_of, size_t sn)
{
    int found = 0;
    const char *p = text;
    char line[2048];
    while (p && *p) {
        const char *e = strpbrk(p, "\r\n");
        size_t l = e ? (size_t)(e - p) : strlen(p);
        if (l >= sizeof line) l = sizeof line - 1;
        memcpy(line, p, l); line[l] = 0;
        p = e ? e + 1 : 0;
        if (!*line || !strncmp(line, "total ", 6)) continue;
        char name[1024] = "", size[64] = ""; int dir = 0;
        if (isdigit((unsigned char)line[0])) {         /* DOS: 01-01-24  12:00AM  <DIR>  name */
            char d1[32], d2[32], s[64];
            int off = 0;
            if (sscanf(line, "%31s %31s %63s %n", d1, d2, s, &off) >= 3 && off) {
                dir = !strcmp(s, "<DIR>");
                snprintf(size, sizeof size, "%s", dir ? "0" : s);
                snprintf(name, sizeof name, "%s", line + off);
            }
        } else {                                        /* Unix: drwxr-xr-x 1 u g size Mon dd hh:mm name */
            char *f[9]; int nf = 0;
            char *q = line;
            while (nf < 8) {
                while (*q == ' ') q++;
                if (!*q) break;
                f[nf++] = q;
                while (*q && *q != ' ') q++;
                if (*q) *q++ = 0;
            }
            while (*q == ' ') q++;
            if (nf == 8 && *q) {
                dir = f[0][0] == 'd';
                snprintf(size, sizeof size, "%s", f[4]);
                snprintf(name, sizeof name, "%.1000s", q);
                char *arrow = strstr(name, " -> ");
                if (arrow && f[0][0] == 'l') *arrow = 0;
            }
        }
        if (!*name || !strcmp(name, ".") || !strcmp(name, "..")) continue;
        if (pattern && *pattern) {                      /* * and ? (no case) */
            const char *s = name, *pt = pattern, *star = 0, *ss = 0;
            int ok = 1;
            while (*s) {
                if (*pt == '?' || tolower((unsigned char)*pt) == tolower((unsigned char)*s)) { s++; pt++; }
                else if (*pt == '*') { star = pt++; ss = s; }
                else if (star) { pt = star + 1; s = ++ss; }
                else { ok = 0; break; }
            }
            while (ok && *pt == '*') pt++;
            if (!ok || *pt) continue;
        }
        if (!found && size_of) snprintf(size_of, sn, "%s", size);
        found++;
        if (out) { bputs(out, name); if (dir) bputs(out, "/"); bputs(out, "\r\n"); }
    }
    return found;
}
static void ftp_success(lbinet *r) { r->ok = 1; r->code = 0; snprintf(r->info, sizeof r->info, "The operation completed successfully.\r\n"); }
/* a reply of the server that was an error: WinINet's extended error, the replies in ResponseInfo */
static void ftp_server_error(lbinet *r, int e)
{
    r->ok = 0;
    if (e == 12003) { r->code = 12003; return; }
    r->code = e;
    snprintf(r->info, sizeof r->info, "%s", inet_text(e));
}
static int ftp_pwd(lbinet *r, lbftp *f)
{
    if (f->cwd_known) return 0;
    buf t = { 0 };
    int c = ftp_cmd(r, f, "PWD", &t);
    if (c == 257 && t.p) {
        const char *a = strchr(t.p, '"'), *b = a ? strchr(a + 1, '"') : 0;
        if (a && b) { snprintf(f->cwd, sizeof f->cwd, "%.*s", (int)(b - a - 1), a + 1); f->cwd_known = 1; }
    }
    bfree(&t);
    return c < 0 ? 12030 : c == 257 ? 0 : 12003;
}
static void ftp_job(lbinet *r)
{
    lbftp *f = r->ftp;
    char op[32]; snprintf(op, sizeof op, "%s", r->method);
    for (char *q = op; *q; q++) *q = (char)toupper((unsigned char)*q);
    r->info[0] = 0;
    buf t = { 0 };
    int e = 0;
    /* a local file that GET would replace, or that PUT cannot read: 11 9 10, nothing sent (oracle) */
    int get = !strcmp(op, "GET") || !strcmp(op, "RECV"), put = !strcmp(op, "PUT") || !strcmp(op, "SEND");
    FILE *in = 0;
    if ((get && exists8(r->a2)) || (put && !(in = fopen8(r->a1, "rb")))) {
        r->ok = 0;
        r->code = get ? 80 : 2;
        snprintf(r->info, sizeof r->info, get ? "The file exists.\r\n" : "The system cannot find the file specified.\r\n");
        ev(r, 11); lbftp_close(f, 1); ev(r, 9); ev(r, 10);
        r->noend = 1;
        return;
    }
    ev(r, 5); ev(r, 6); ev(r, 7);
    if (!strcmp(op, "PWD")) {
        if (!(e = ftp_pwd(r, f))) bputs(&r->body, f->cwd);
    } else if (!strcmp(op, "CD") || !strcmp(op, "CDUP")) {
        char cmd[2100]; snprintf(cmd, sizeof cmd, "CWD %s", !strcmp(op, "CDUP") ? ".." : r->a1);
        r->info[0] = 0;
        int c = ftp_cmd(r, f, cmd, &t);
        e = c < 0 ? 12030 : c / 100 == 2 ? 0 : 12003;
        if (!e) { f->cwd_known = 0; ftp_pwd(r, f); }
    } else if (!strcmp(op, "DIR") || !strcmp(op, "LS") || !strcmp(op, "SIZE")) {
        buf list = { 0 };
        const char *slash = strrchr(r->a1, '/');
        char cmd[2100] = "LIST", pat[1024];
        snprintf(pat, sizeof pat, "%.1000s", slash ? slash + 1 : r->a1);
        if (slash) snprintf(cmd, sizeof cmd, "LIST %.*s", (int)(slash - r->a1 + 1), r->a1);
        e = ftp_transfer(r, f, "A", cmd, &list, 0);
        if (!e) {
            char size[64] = "";
            int n = list_names(list.p ? list.p : "", pat, !strcmp(op, "SIZE") ? 0 : &r->body, size, sizeof size);
            if (!strcmp(op, "SIZE")) bputs(&r->body, size);
            if (!n || strcmp(op, "SIZE")) {
                /* FtpFindFirstFile / InternetFindNextFile end with "no more files" (VB6 oracle) */
                if (strcmp(op, "SIZE")) bputs(&r->body, "\r\n");
                r->code = 18; snprintf(r->info, sizeof r->info, "There are no more files.\r\n");
                r->ok = n > 0;
                if (!n) bclear(&r->body);
                bfree(&list); bfree(&t);
                ev(r, 8);
                return;
            }
        }
        bfree(&list);
    } else if (!strcmp(op, "MKDIR") || !strcmp(op, "RMDIR") || !strcmp(op, "DELETE")) {
        char cmd[2100];
        snprintf(cmd, sizeof cmd, "%s %s", !strcmp(op, "MKDIR") ? "MKD" : !strcmp(op, "RMDIR") ? "RMD" : "DELE", r->a1);
        int c = ftp_cmd(r, f, cmd, &t);
        e = c < 0 ? 12030 : c / 100 == 2 ? 0 : 12003;
    } else if (!strcmp(op, "RENAME")) {
        char cmd[2100]; snprintf(cmd, sizeof cmd, "RNFR %s", r->a1);
        int c = ftp_cmd(r, f, cmd, &t);
        if (c == 350) { snprintf(cmd, sizeof cmd, "RNTO %s", r->a2); c = ftp_cmd(r, f, cmd, &t); }
        e = c < 0 ? 12030 : c / 100 == 2 ? 0 : 12003;
    } else if (!strcmp(op, "GET") || !strcmp(op, "RECV")) {
        buf data = { 0 };
        char cmd[2100];
        int c = ftp_cmd(r, f, "TYPE I", &t);
        if (c < 0) e = 12030;
        else {
            snprintf(cmd, sizeof cmd, "SIZE %s", r->a1);
            sock_t ls, ds;
            /* as WinINet: TYPE, PORT, SIZE, RETR */
            e = ftp_data_open(r, f, &ls, &ds);
            if (!e) {
                c = ftp_cmd(r, f, cmd, &t);
                snprintf(cmd, sizeof cmd, "RETR %s", r->a1);
                c = c < 0 ? -1 : ftp_cmd(r, f, cmd, &t);
                if (c / 100 != 1) { closes(ls); closes(ds); e = c < 0 ? 12030 : 12003; }
                else if (!(e = ftp_data_ready(r, f, &ls, &ds))) {
                    lbconn dc; memset(&dc, 0, sizeof dc); dc.s = ds;
                    char tmp[16384]; int got;
                    for (;;) {
                        if ((e = recv_some(r, &dc, tmp, (int)sizeof tmp, r->timeout, &got)) || !got) break;
                        bput(&data, tmp, (size_t)got);
                        ev(r, 7);
                    }
                    closes(ds);
                    if (!e) { c = ftp_reply(r, f, &t); e = c / 100 == 2 ? 0 : c < 0 ? 12030 : 12003; }
                }
            }
        }
        if (!e) {
            FILE *o = fopen8(r->a2, "wb");
            if (!o) { e = 3; r->ok = 0; r->code = 3; snprintf(r->info, sizeof r->info, "The system cannot find the path specified.\r\n"); }
            else { if (data.n) fwrite(data.p, 1, data.n, o); fclose(o); }
        }
        bfree(&data);
        if (e == 3) { bfree(&t); ev(r, 8); return; }
    } else if (!strcmp(op, "PUT") || !strcmp(op, "SEND")) {
        buf up = { 0 }; char tmp[16384]; size_t n;
        while ((n = fread(tmp, 1, sizeof tmp, in)) > 0) bput(&up, tmp, n);
        fclose(in);
        char cmd[2100]; snprintf(cmd, sizeof cmd, "STOR %s", r->a2);
        e = ftp_transfer(r, f, "I", cmd, 0, &up);
        bfree(&up);
    } else {
        /* a command the control does not know: nothing is sent (VB6 oracle: "FOO bar") */
        bfree(&t);
        r->ok = 1; r->code = 0; r->info[0] = 0;
        ev(r, 8);
        return;
    }
    bfree(&t);
    if (r->cancel) return;
    ev(r, 8);
    if (e) { ftp_server_error(r, e); if (e != 12003) lbftp_close(f, 0); }
    else ftp_success(r);
}
/* OpenURL of an ftp:// URL: a session of its own, the file (or the names of a folder) */
static void ftpurl_job(lbinet *r)
{
    lbftp *f = lbftp_new();
    int e = lbftp_open(f, r->host, r->port, r->user, r->pass, r->timeout);
    if (e) { fail(r, e, 0); lbftp_free(f); return; }
    ev(r, 3); ev(r, 4); ev(r, 5); ev(r, 6); ev(r, 7);
    size_t l = strlen(r->path);
    if (!l || r->path[l - 1] == '/') {
        buf list = { 0 };
        char cmd[2100]; snprintf(cmd, sizeof cmd, "LIST %s", l ? r->path : "/");
        e = ftp_transfer(r, f, "A", cmd, &list, 0);
        if (!e) list_names(list.p ? list.p : "", "", &r->body, 0, 0);
        bfree(&list);
    } else {
        char cmd[2100]; snprintf(cmd, sizeof cmd, "RETR %s", r->path);
        buf t = { 0 };
        snprintf(cmd, sizeof cmd, "SIZE %s", r->path);
        int c = ftp_cmd(r, f, "TYPE I", &t);
        if (c >= 0) c = ftp_cmd(r, f, cmd, &t);
        bfree(&t);
        snprintf(cmd, sizeof cmd, "RETR %s", r->path);
        e = c < 0 ? 12030 : ftp_transfer(r, f, "I", cmd, &r->body, 0);
    }
    lbftp_free(f);
    if (r->cancel) return;
    if (e) { r->ok = 0; r->code = e; if (e != 12003) snprintf(r->info, sizeof r->info, "%s", inet_text(e)); return; }
    ev(r, 8);
    r->ok = 1; r->code = 0; r->info[0] = 0;
}

/* ---------------------------------------------------------------- the thread */
static void job_free(lbinet *r)
{
    if (r->conn) { conn_close(r->conn); free(r->conn); }
    if (r->ftp) lbftp_free(r->ftp);
    bfree(&r->hdrs); bfree(&r->data); bfree(&r->body); bfree(&r->headers);
    mu_free(&r->mu);
    free(r);
}
static void job_run(lbinet *r)
{
    if (r->kind == K_HTTP) http_job(r);
    else if (r->kind == K_FTP) ftp_job(r);
    else ftpurl_job(r);
    if (!r->cancel && !r->noend) ev(r, r->ok ? 12 : 11);
    mu_lock(&r->mu);
    r->finished = 1;
    int af = r->autofree;
    mu_unlock(&r->mu);
    if (af) job_free(r);
}
#ifdef _WIN32
static DWORD WINAPI thread_main(LPVOID p) { job_run((lbinet *)p); return 0; }
#else
static void *thread_main(void *p) { job_run((lbinet *)p); return 0; }
#endif
static int job_start(lbinet *r)
{
    mu_init(&r->mu);
#ifdef _WIN32
    r->th = CreateThread(NULL, 0, thread_main, r, 0, NULL);
    if (!r->th) return 0;
    CloseHandle(r->th);
#else
    pthread_attr_t a; pthread_attr_init(&a); pthread_attr_setdetachstate(&a, PTHREAD_CREATE_DETACHED);
    int e = pthread_create(&r->th, &a, thread_main, r);
    pthread_attr_destroy(&a);
    if (e) return 0;
#endif
    return 1;
}

/* ---------------------------------------------------------------- what the control calls */
typedef struct lbinet_conn_box { lbconn *c; } lbinet_conn_box;
void *lbinet_conn_new(void) { lbconn *c = (lbconn *)calloc(1, sizeof *c); c->s = BADS; return c; }
void lbinet_conn_free(void *c) { if (c) { conn_close((lbconn *)c); free(c); } }

/* an HTTP(S) request; conn: the kept connection of the control (the request owns it until it is taken back) */
void *lbinet_http(void *conn, int https, int openurl, const char *method, const char *host, int port, const char *path,
                  const char *user, const char *pass, const char *hdrs, const char *data, int ndata, int timeout,
                  int access, const char *proxy)
{
    lbinet *r = (lbinet *)calloc(1, sizeof *r);
    r->kind = K_HTTP; r->https = https; r->openurl = openurl; r->timeout = timeout; r->access = access; r->port = port;
    snprintf(r->method, sizeof r->method, "%s", method);
    snprintf(r->host, sizeof r->host, "%s", host);
    snprintf(r->path, sizeof r->path, "%s", path);
    snprintf(r->user, sizeof r->user, "%s", user ? user : "");
    snprintf(r->pass, sizeof r->pass, "%s", pass ? pass : "");
    snprintf(r->proxy, sizeof r->proxy, "%s", proxy ? proxy : "");
    if (hdrs && *hdrs) bputs(&r->hdrs, hdrs);
    if (data && ndata > 0) bput(&r->data, data, (size_t)ndata);
    r->conn = (lbconn *)(conn ? conn : lbinet_conn_new());
    if (!job_start(r)) { fail(r, 12004, 0); r->finished = 1; ev(r, 11); }
    return r;
}
/* an FTP command (Execute) on the session of the control */
void *lbinet_ftp(void *session, const char *op, const char *a1, const char *a2, int timeout)
{
    lbinet *r = (lbinet *)calloc(1, sizeof *r);
    r->kind = K_FTP; r->timeout = timeout;
    snprintf(r->method, sizeof r->method, "%s", op);
    snprintf(r->a1, sizeof r->a1, "%s", a1 ? a1 : "");
    snprintf(r->a2, sizeof r->a2, "%s", a2 ? a2 : "");
    r->ftp = (lbftp *)session;
    if (!job_start(r)) { fail(r, 12004, 0); r->finished = 1; ev(r, 11); }
    return r;
}
void *lbinet_ftpurl(const char *host, int port, const char *path, const char *user, const char *pass, int timeout)
{
    lbinet *r = (lbinet *)calloc(1, sizeof *r);
    r->kind = K_FTPURL; r->timeout = timeout; r->port = port;
    snprintf(r->host, sizeof r->host, "%s", host);
    snprintf(r->path, sizeof r->path, "%s", path);
    snprintf(r->user, sizeof r->user, "%s", user ? user : "");
    snprintf(r->pass, sizeof r->pass, "%s", pass ? pass : "");
    if (!job_start(r)) { fail(r, 12004, 0); r->finished = 1; ev(r, 11); }
    return r;
}
/* the next state (StateChanged), -1 none now; 11 / 12 come last */
int lbinet_event(void *p)
{
    lbinet *r = (lbinet *)p;
    int v = -1;
    mu_lock(&r->mu);
    if (r->ev_out != r->ev_in) { v = r->ev[r->ev_out]; r->ev_out = (r->ev_out + 1) % 512; }
    mu_unlock(&r->mu);
    return v;
}
int lbinet_stage(void *p) { lbinet *r = (lbinet *)p; mu_lock(&r->mu); int s = r->finished ? 99 : r->stage; mu_unlock(&r->mu); return s; }
int lbinet_done(void *p) { lbinet *r = (lbinet *)p; mu_lock(&r->mu); int d = r->finished; mu_unlock(&r->mu); return d; }
int lbinet_ok(void *p) { return ((lbinet *)p)->ok; }
int lbinet_code(void *p) { return ((lbinet *)p)->code; }
const char *lbinet_info(void *p) { return ((lbinet *)p)->info; }
const char *lbinet_body(void *p, long *n) { lbinet *r = (lbinet *)p; *n = (long)r->body.n; return r->body.p ? r->body.p : ""; }
const char *lbinet_headers(void *p) { lbinet *r = (lbinet *)p; return r->headers.p ? r->headers.p : ""; }
/* HTTP: the URL after the redirects ("" none) and their number */
const char *lbinet_final(void *p, int *redirects) { lbinet *r = (lbinet *)p; *redirects = r->redirects; return r->finalurl; }
/* the kept connection back to the control (after the request) */
void *lbinet_take_conn(void *p) { lbinet *r = (lbinet *)p; void *c = r->conn; r->conn = 0; return c; }
void *lbinet_take_ftp(void *p) { lbinet *r = (lbinet *)p; void *f = r->ftp; r->ftp = 0; return f; }
/* the request is no longer wanted: cancelled if it still runs, freed when its thread ends */
void lbinet_release(void *p)
{
    lbinet *r = (lbinet *)p;
    if (!r) return;
    mu_lock(&r->mu);
    r->cancel = 1;
    int fin = r->finished;
    if (!fin) r->autofree = 1;
    mu_unlock(&r->mu);
    if (fin) job_free(r);
}
/* the session of the control */
void *lbinet_ftp_new(void) { return lbftp_new(); }
void lbinet_ftp_free(void *f) { lbftp_free((lbftp *)f); }
int lbinet_ftp_open(void *f, const char *host, int port, const char *user, const char *pass, int timeout)
{
    return lbftp_open((lbftp *)f, host, port, user, pass, timeout);
}
int lbinet_ftp_isopen(void *f) { return lbftp_isopen((lbftp *)f); }
int lbinet_ftp_same(void *f, const char *host, int port, const char *user) { return lbftp_same((lbftp *)f, host, port, user); }
void lbinet_ftp_close(void *f) { lbftp_close((lbftp *)f, 1); }
