#define _GNU_SOURCE
/* net.c - the sockets of the Winsock control (forms/ocx_winsock.inc): TCP and UDP over IPv4, all of
 * them non-blocking; the control asks lbnet_poll what is ready (a timer of the forms runtime).
 * The errors are the numbers of Windows Sockets (10061 ...) on every system, as the control gives them.
 * Linux: BSD sockets. Windows: Winsock 2 (ws2_32). Apart from vbrt.h (windows.h names VT_EMPTY too).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifdef _WIN32
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0601
#endif
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#else
#include <errno.h>
#include <fcntl.h>
#include <unistd.h>
#include <netdb.h>
#include <poll.h>
#include <ifaddrs.h>
#include <net/if.h>
#include <sys/socket.h>
#include <sys/ioctl.h>
#include <netinet/in.h>
#include <netinet/tcp.h>
#include <arpa/inet.h>
#endif
#include <stdio.h>
#include <string.h>

#ifdef _WIN32
typedef SOCKET sock_t;
#define BAD_SOCK INVALID_SOCKET
static int ready;
static void init(void)
{
    if (!ready) { WSADATA w; WSAStartup(MAKEWORD(2, 2), &w); ready = 1; }
}
static int lasterr(void) { return WSAGetLastError(); }
#else
typedef int sock_t;
#define BAD_SOCK (-1)
static void init(void) {}
/* errno -> the number Windows Sockets gives for it */
static int wsa(int e)
{
    switch (e) {
    case EINTR: return 10004;
    case EBADF: return 10009;
    case EACCES: return 10013;
    case EFAULT: return 10014;
    case EINVAL: return 10022;
    case EMFILE: return 10024;
    case EWOULDBLOCK: return 10035;
    case EINPROGRESS: return 10036;
    case EALREADY: return 10037;
    case ENOTSOCK: return 10038;
    case EDESTADDRREQ: return 10039;
    case EMSGSIZE: return 10040;
    case EPROTOTYPE: return 10041;
    case ENOPROTOOPT: return 10042;
    case EPROTONOSUPPORT: return 10043;
    case EOPNOTSUPP: return 10045;
    case EAFNOSUPPORT: return 10047;
    case EADDRINUSE: return 10048;
    case EADDRNOTAVAIL: return 10049;
    case ENETDOWN: return 10050;
    case ENETUNREACH: return 10051;
    case ENETRESET: return 10052;
    case ECONNABORTED: return 10053;
    case ECONNRESET: case EPIPE: return 10054;
    case ENOBUFS: case ENOMEM: return 10055;
    case EISCONN: return 10056;
    case ENOTCONN: return 10057;
    case ESHUTDOWN: return 10058;
    case ETIMEDOUT: return 10060;
    case ECONNREFUSED: return 10061;
    case EHOSTDOWN: return 10064;
    case EHOSTUNREACH: return 10065;
    default: return 10050;
    }
}
static int lasterr(void) { return wsa(errno); }
#endif
#define S(h) ((sock_t)(h))

static void nonblock(sock_t s)
{
#ifdef _WIN32
    u_long one = 1;
    ioctlsocket(s, FIONBIO, &one);
#else
    fcntl(s, F_SETFL, fcntl(s, F_GETFL) | O_NONBLOCK);
    fcntl(s, F_SETFD, FD_CLOEXEC);
#endif
}

static int addr_of(const char *ip, int port, struct sockaddr_in *a)
{
    memset(a, 0, sizeof *a);
    a->sin_family = AF_INET;
    a->sin_port = htons((unsigned short)port);
    if (!ip || !*ip) { a->sin_addr.s_addr = htonl(INADDR_ANY); return 0; }
    return inet_pton(AF_INET, ip, &a->sin_addr) == 1 ? 0 : 10049;
}

/* a host name or a dotted address -> "a.b.c.d"; 0 or 11001 (host not found) */
int lbnet_resolve(const char *host, char ip[16])
{
    init();
    struct in_addr x;
    if (inet_pton(AF_INET, host, &x) == 1) { inet_ntop(AF_INET, &x, ip, 16); return 0; }
    struct addrinfo hint, *r = 0;
    memset(&hint, 0, sizeof hint);
    hint.ai_family = AF_INET;
    hint.ai_socktype = SOCK_STREAM;
    if (getaddrinfo(host, 0, &hint, &r) || !r) return 11001;
    inet_ntop(AF_INET, &((struct sockaddr_in *)r->ai_addr)->sin_addr, ip, 16);
    freeaddrinfo(r);
    return 0;
}
/* 1: the text is an address ("127.0.0.1"), not a name to resolve */
int lbnet_is_ip(const char *host) { struct in_addr x; return inet_pton(AF_INET, host, &x) == 1; }

/* a new socket, non-blocking; -1 */
long long lbnet_open(int udp)
{
    init();
    sock_t s = socket(AF_INET, udp ? SOCK_DGRAM : SOCK_STREAM, udp ? IPPROTO_UDP : IPPROTO_TCP);
    if (s == BAD_SOCK) return -1;
    nonblock(s);
#ifndef _WIN32
    /* Windows binds the port of a connection that is closing (TIME_WAIT) again; Linux only with this */
    if (!udp) { int one = 1; setsockopt(s, SOL_SOCKET, SO_REUSEADDR, &one, sizeof one); }
#endif
    return (long long)s;
}
void lbnet_close(long long h)
{
    if (h == -1) return;
#ifdef _WIN32
    closesocket(S(h));
#else
    close(S(h));
#endif
}
int lbnet_bind(long long h, const char *ip, int port)
{
    struct sockaddr_in a;
    int e = addr_of(ip, port, &a);
    if (e) return e;
    return bind(S(h), (struct sockaddr *)&a, sizeof a) ? lasterr() : 0;
}
int lbnet_listen(long long h) { return listen(S(h), SOMAXCONN) ? lasterr() : 0; }
/* 0: connected; 10035 / 10036: on its way (lbnet_poll says when); else the error */
int lbnet_connect(long long h, const char *ip, int port)
{
    struct sockaddr_in a;
    int e = addr_of(ip, port, &a);
    if (e) return e;
    if (!connect(S(h), (struct sockaddr *)&a, sizeof a)) return 0;
    return lasterr();
}
/* what the socket is ready for, without waiting: *rd (data, a connection to accept, the end),
 * *wr (it can send; a connection made), *ex (a connection that failed) */
int lbnet_poll(long long h, int *rd, int *wr, int *ex)
{
    *rd = *wr = *ex = 0;
#ifdef _WIN32
    fd_set r, w, x;
    FD_ZERO(&r); FD_ZERO(&w); FD_ZERO(&x);
    FD_SET(S(h), &r); FD_SET(S(h), &w); FD_SET(S(h), &x);
    struct timeval tv = { 0, 0 };
    int n = select(0, &r, &w, &x, &tv);
    if (n == SOCKET_ERROR) return lasterr();
    *rd = FD_ISSET(S(h), &r) != 0; *wr = FD_ISSET(S(h), &w) != 0; *ex = FD_ISSET(S(h), &x) != 0;
#else
    struct pollfd p = { S(h), POLLIN | POLLOUT, 0 };
    if (poll(&p, 1, 0) < 0) return lasterr();
    *rd = (p.revents & (POLLIN | POLLHUP)) != 0;
    *wr = (p.revents & POLLOUT) != 0;
    *ex = (p.revents & POLLERR) != 0;
#endif
    return 0;
}
/* the error of a connection that failed (0: none) */
int lbnet_sockerr(long long h)
{
    int e = 0;
    socklen_t n = sizeof e;
    if (getsockopt(S(h), SOL_SOCKET, SO_ERROR, (char *)&e, &n)) return lasterr();
#ifdef _WIN32
    return e;
#else
    return e ? wsa(e) : 0;
#endif
}
/* a connection that waits on a listening socket; -1 */
long long lbnet_accept(long long h)
{
    sock_t s = accept(S(h), 0, 0);
    if (s == BAD_SOCK) return -1;
    nonblock(s);
    return (long long)s;
}
/* >= 0: the bytes sent; -10035: none can go now; < 0: -error */
int lbnet_send(long long h, const char *b, int n)
{
#ifdef _WIN32
    int r = send(S(h), b, n, 0);
#else
    int r = (int)send(S(h), b, (size_t)n, MSG_NOSIGNAL);
#endif
    return r < 0 ? -lasterr() : r;
}
/* > 0: the bytes; 0: the other side closed the connection; -10035: nothing now; < 0: -error */
int lbnet_recv(long long h, char *b, int n)
{
    int r = (int)recv(S(h), b, n, 0);
    return r < 0 ? -lasterr() : r;
}
int lbnet_sendto(long long h, const char *ip, int port, const char *b, int n)
{
    struct sockaddr_in a;
    int e = addr_of(ip, port, &a);
    if (e) return -e;
    int r = (int)sendto(S(h), b, n, 0, (struct sockaddr *)&a, sizeof a);
    return r < 0 ? -lasterr() : r;
}
/* one datagram (b: 65536 bytes) and who sent it */
int lbnet_recvfrom(long long h, char *b, int n, char ip[16], int *port)
{
    struct sockaddr_in a;
    socklen_t al = sizeof a;
    int r = (int)recvfrom(S(h), b, n, 0, (struct sockaddr *)&a, &al);
    if (r < 0) return -lasterr();
    inet_ntop(AF_INET, &a.sin_addr, ip, 16);
    *port = ntohs(a.sin_port);
    return r;
}
/* the address and the port of the socket (peer 0) or of the other side (peer 1) */
int lbnet_name(long long h, int peer, char ip[16], int *port)
{
    struct sockaddr_in a;
    socklen_t al = sizeof a;
    int r = peer ? getpeername(S(h), (struct sockaddr *)&a, &al) : getsockname(S(h), (struct sockaddr *)&a, &al);
    if (r) { ip[0] = 0; *port = 0; return lasterr(); }
    inet_ntop(AF_INET, &a.sin_addr, ip, 16);
    *port = ntohs(a.sin_port);
    return 0;
}
void lbnet_hostname(char *b, int n)
{
    init();
    if (gethostname(b, n)) snprintf(b, (size_t)n, "localhost");
    b[n - 1] = 0;
#ifndef _WIN32
    char *dot = strchr(b, '.');         /* Windows gives the name of the computer without its domain */
    if (dot) *dot = 0;
#endif
}
/* the address of the computer on its network (LocalIP): the first IPv4 address that is not 127.x */
void lbnet_localip(char ip[16])
{
    init();
    snprintf(ip, 16, "127.0.0.1");
#ifdef _WIN32
    char host[256];
    struct addrinfo hint, *r = 0, *p;
    if (gethostname(host, sizeof host)) return;
    memset(&hint, 0, sizeof hint);
    hint.ai_family = AF_INET;
    if (getaddrinfo(host, 0, &hint, &r)) return;
    for (p = r; p; p = p->ai_next) {
        struct in_addr x = ((struct sockaddr_in *)p->ai_addr)->sin_addr;
        if ((ntohl(x.s_addr) >> 24) == 127) continue;
        inet_ntop(AF_INET, &x, ip, 16);
        break;
    }
    freeaddrinfo(r);
#else
    struct ifaddrs *l = 0, *p;
    if (getifaddrs(&l)) return;
    for (p = l; p; p = p->ifa_next) {
        if (!p->ifa_addr || p->ifa_addr->sa_family != AF_INET || (p->ifa_flags & IFF_LOOPBACK) || !(p->ifa_flags & IFF_UP))
            continue;
        inet_ntop(AF_INET, &((struct sockaddr_in *)p->ifa_addr)->sin_addr, ip, 16);
        break;
    }
    freeifaddrs(l);
#endif
}
