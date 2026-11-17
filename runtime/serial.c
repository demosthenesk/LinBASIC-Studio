#define _GNU_SOURCE
/* serial.c - the serial ports of the MSComm control (forms/ocx_comm.inc).
 * Linux: termios; CommPort n is $VBRT_COMn when it is set, otherwise the n-th serial port of the
 * computer: the real /dev/ttyS* (not the empty ones Linux always has), then /dev/ttyUSB*, /dev/ttyACM*.
 * Windows: \\.\COMn. Apart from vbrt.h (windows.h names VT_EMPTY too).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdint.h>
#ifdef _WIN32
#include <windows.h>
#else
#include <errno.h>
#include <fcntl.h>
#include <unistd.h>
#include <dirent.h>
#include <termios.h>
#include <sys/ioctl.h>
#include <sys/file.h>
#include <linux/serial.h>
#endif

#ifndef _WIN32
static int numcmp(const void *a, const void *b)
{
    const char *x = *(const char *const *)a, *y = *(const char *const *)b;
    size_t i = 0;
    while (x[i] && x[i] == y[i] && !(x[i] >= '0' && x[i] <= '9')) i++;
    long nx = strtol(x + i, 0, 10), ny = strtol(y + i, 0, 10);
    int c = strncmp(x, y, i);
    if (c) return c;
    return nx < ny ? -1 : nx > ny;
}
/* the serial ports of the computer, in the order of CommPort */
static int list_ports(char out[][64], int max)
{
    DIR *d = opendir("/dev");
    if (!d) return 0;
    char *names[512]; int n = 0;
    struct dirent *e;
    while ((e = readdir(d)) && n < 512)
        if (!strncmp(e->d_name, "ttyS", 4) || !strncmp(e->d_name, "ttyUSB", 6) || !strncmp(e->d_name, "ttyACM", 6))
            names[n++] = strdup(e->d_name);
    closedir(d);
    qsort(names, (size_t)n, sizeof *names, numcmp);
    int k = 0;
    for (int pass = 0; pass < 3; pass++)
        for (int i = 0; i < n && k < max; i++) {
            const char *nm = names[i];
            int is_s = !strncmp(nm, "ttyS", 4), is_u = !strncmp(nm, "ttyUSB", 6);
            if ((pass == 0 && !is_s) || (pass == 1 && !is_u) || (pass == 2 && (is_s || is_u))) continue;
            char p[64]; snprintf(p, sizeof p, "/dev/%s", nm);
            if (is_s) {                         /* a real 8250/16550 port */
                int fd = open(p, O_RDWR | O_NOCTTY | O_NONBLOCK);
                if (fd < 0) continue;
                struct serial_struct ss;
                int real = ioctl(fd, TIOCGSERIAL, &ss) == 0 && ss.type != PORT_UNKNOWN;
                close(fd);
                if (!real) continue;
            }
            snprintf(out[k++], 64, "%s", p);
        }
    for (int i = 0; i < n; i++) free(names[i]);
    return k;
}
#endif

/* 0, or the error of MSComm: 8002 no such port, 8005 in use */
int lbser_open(int n, long long *handle)
{
#ifdef _WIN32
    wchar_t name[32];
    swprintf(name, 32, L"\\\\.\\COM%d", n);
    HANDLE h = CreateFileW(name, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) return GetLastError() == ERROR_ACCESS_DENIED ? 8005 : 8002;
    COMMTIMEOUTS to = { MAXDWORD, 0, 0, 0, 0 };
    SetCommTimeouts(h, &to);
    *handle = (long long)(intptr_t)h;
    return 0;
#else
    char env[32], path[256] = "";
    snprintf(env, sizeof env, "VBRT_COM%d", n);
    const char *e = getenv(env);
    if (e && *e) snprintf(path, sizeof path, "%s", e);
    else {
        char ports[64][64];
        int k = list_ports(ports, 64);
        if (n >= 1 && n <= k) snprintf(path, sizeof path, "%s", ports[n - 1]);
    }
    if (!*path) return 8002;
    int fd = open(path, O_RDWR | O_NOCTTY | O_NONBLOCK);
    if (fd < 0) return errno == EBUSY || errno == EACCES ? 8005 : 8002;
    if (flock(fd, LOCK_EX | LOCK_NB) != 0) { close(fd); return 8005; }
    *handle = fd;
    return 0;
#endif
}
void lbser_close(long long h)
{
#ifdef _WIN32
    CloseHandle((HANDLE)(intptr_t)h);
#else
    flock((int)h, LOCK_UN);
    close((int)h);
#endif
}
#ifndef _WIN32
static speed_t baud_const(int b)
{
    switch (b) {
    case 110: return B110; case 300: return B300; case 600: return B600; case 1200: return B1200; case 2400: return B2400;
    case 4800: return B4800; case 9600: return B9600; case 19200: return B19200; case 38400: return B38400;
    case 57600: return B57600; case 115200: return B115200; case 230400: return B230400;
    }
    return 0;
}
#endif
/* 1 if the port took the settings */
int lbser_config(long long h, int baud, char parity, int data, double stop, int hs)
{
#ifdef _WIN32
    HANDLE p = (HANDLE)(intptr_t)h;
    DCB d; memset(&d, 0, sizeof d); d.DCBlength = sizeof d;
    GetCommState(p, &d);
    d.BaudRate = (DWORD)baud; d.ByteSize = (BYTE)data; d.fBinary = 1;
    d.Parity = parity == 'E' ? EVENPARITY : parity == 'O' ? ODDPARITY : parity == 'M' ? MARKPARITY : parity == 'S' ? SPACEPARITY : NOPARITY;
    d.fParity = d.Parity != NOPARITY;
    d.StopBits = stop >= 2 ? TWOSTOPBITS : stop > 1 ? ONE5STOPBITS : ONESTOPBIT;
    d.fOutX = d.fInX = (hs == 1 || hs == 3);
    d.fOutxCtsFlow = (hs == 2 || hs == 3);
    if (d.fOutxCtsFlow) d.fRtsControl = RTS_CONTROL_HANDSHAKE;
    return SetCommState(p, &d) != 0;
#else
    int fd = (int)h;
    struct termios t;
    if (tcgetattr(fd, &t) != 0) return 1;
    cfmakeraw(&t);
    speed_t sp = baud_const(baud);
    if (!sp) sp = baud < 14400 ? B9600 : baud < 28800 ? B19200 : baud < 56000 ? B38400 : baud < 115200 ? B57600 : B115200;
    cfsetispeed(&t, sp); cfsetospeed(&t, sp);
    t.c_cflag &= ~(tcflag_t)(CSIZE | PARENB | PARODD | CSTOPB | CRTSCTS);
    t.c_cflag |= CLOCAL | CREAD;
    t.c_cflag |= data == 5 ? CS5 : data == 6 ? CS6 : data == 7 ? CS7 : CS8;
    if (parity == 'E') t.c_cflag |= PARENB;
    else if (parity == 'O') t.c_cflag |= PARENB | PARODD;
#ifdef CMSPAR
    else if (parity == 'M') t.c_cflag |= PARENB | PARODD | CMSPAR;
    else if (parity == 'S') t.c_cflag |= PARENB | CMSPAR;
#endif
    if (stop >= 1.5) t.c_cflag |= CSTOPB;
    t.c_iflag &= ~(tcflag_t)(IXON | IXOFF | IXANY);
    if (hs == 1 || hs == 3) t.c_iflag |= IXON | IXOFF;
    if (hs == 2 || hs == 3) t.c_cflag |= CRTSCTS;
    t.c_cc[VMIN] = 0; t.c_cc[VTIME] = 0;
    tcsetattr(fd, TCSANOW, &t);
    return 1;
#endif
}
void lbser_lines(long long h, int dtr, int rts)
{
#ifdef _WIN32
    HANDLE p = (HANDLE)(intptr_t)h;
    EscapeCommFunction(p, dtr ? SETDTR : CLRDTR);
    EscapeCommFunction(p, rts ? SETRTS : CLRRTS);
#else
    int bits = TIOCM_DTR;
    ioctl((int)h, dtr ? TIOCMBIS : TIOCMBIC, &bits);
    bits = TIOCM_RTS;
    ioctl((int)h, rts ? TIOCMBIS : TIOCMBIC, &bits);
#endif
}
void lbser_break(long long h, int on)
{
#ifdef _WIN32
    if (on) SetCommBreak((HANDLE)(intptr_t)h); else ClearCommBreak((HANDLE)(intptr_t)h);
#else
    ioctl((int)h, on ? TIOCSBRK : TIOCCBRK, 0);
#endif
}
int lbser_read(long long h, char *buf, int n)
{
#ifdef _WIN32
    DWORD r = 0;
    if (!ReadFile((HANDLE)(intptr_t)h, buf, (DWORD)n, &r, NULL)) return 0;
    return (int)r;
#else
    ssize_t r = read((int)h, buf, (size_t)n);
    return r > 0 ? (int)r : 0;
#endif
}
int lbser_write(long long h, const char *buf, int n)
{
#ifdef _WIN32
    DWORD w = 0;
    WriteFile((HANDLE)(intptr_t)h, buf, (DWORD)n, &w, NULL);
    return (int)w;
#else
    int done = 0;
    for (int tries = 0; done < n && tries < 2000; tries++) {
        ssize_t r = write((int)h, buf + done, (size_t)(n - done));
        if (r > 0) done += (int)r;
        else if (errno == EAGAIN) usleep(1000);
        else break;
    }
    return done;
#endif
}
int lbser_outq(long long h)
{
#ifdef _WIN32
    COMSTAT s; DWORD e;
    if (!ClearCommError((HANDLE)(intptr_t)h, &e, &s)) return 0;
    return (int)s.cbOutQue;
#else
    int n = 0;
    if (ioctl((int)h, TIOCOUTQ, &n) != 0) n = 0;
    return n;
#endif
}
void lbser_flush(long long h, int in, int out)
{
#ifdef _WIN32
    PurgeComm((HANDLE)(intptr_t)h, (in ? PURGE_RXCLEAR : 0) | (out ? PURGE_TXCLEAR : 0));
#else
    if (in) tcflush((int)h, TCIFLUSH);
    if (out) tcflush((int)h, TCOFLUSH);
#endif
}
/* bits: 1 CTS, 2 DSR, 4 CD, 8 RING */
int lbser_modem(long long h)
{
#ifdef _WIN32
    DWORD s = 0;
    GetCommModemStatus((HANDLE)(intptr_t)h, &s);
    return (s & MS_CTS_ON ? 1 : 0) | (s & MS_DSR_ON ? 2 : 0) | (s & MS_RLSD_ON ? 4 : 0) | (s & MS_RING_ON ? 8 : 0);
#else
    int s = 0, r = 0;
    if (ioctl((int)h, TIOCMGET, &s) != 0) return 0;
    if (s & TIOCM_CTS) r |= 1;
    if (s & TIOCM_DSR) r |= 2;
    if (s & TIOCM_CD) r |= 4;
    if (s & TIOCM_RI) r |= 8;
    return r;
#endif
}
