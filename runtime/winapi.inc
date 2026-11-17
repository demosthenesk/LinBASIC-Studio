/* winapi.inc - more Windows procedures for the Declares of programs on Linux (part of dll.c):
 * kernel32 (files, memory, modules, locale, time), advapi32 (the registry, kept in a file),
 * shell32 (ShellExecute, the special folders, SHFileOperation), user32 (rectangles, MessageBox),
 * comctl32, ole32, winmm, wininet. The procedures of windows and drawing are in the forms runtime
 * (forms/winapi_gui.inc). Seen in the Planet Source Code test corpus: the procedures that VB6
 * programs call most.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef _WIN32
#include <sys/statvfs.h>
#include <fnmatch.h>
#include <dirent.h>
#include <ctype.h>
#include <sys/stat.h>
#include <errno.h>

static vb_i32 *a_long(args *x, int i) { return (vb_i32 *)a_ptr(x, i, 4); }
static void put16(unsigned char *p, vb_u16 v) { memcpy(p, &v, 2); }
/* a result string into a buffer of the program: n = its size; the length written (without the 0) */
static vb_i64 put_text(char *buf, vb_i64 size, const char *v)
{
    size_t n = strlen(v);
    if (!buf || size <= 0) return (vb_i64)n + 1;
    if ((vb_i64)n + 1 > size) { memcpy(buf, v, (size_t)size - 1); buf[size - 1] = 0; return (vb_i64)n + 1; }
    memcpy(buf, v, n + 1);
    return (vb_i64)n;
}
/* a path of the program (C:\x\y.txt, x\y) as a path of this computer */
static void lpath(const char *in, char *out, size_t n)
{
    vb_i32 mk = vb_mark();
    vbs s = vb_sfromansi(in ? in : "", in ? (int)strlen(in) : 0);
    snprintf(out, n, "%s", vb_stoutf8(s));
    vb_release(mk);
    for (char *p = out; *p; p++) if (*p == '\\') *p = '/';
    if (((out[0] | 0x20) >= 'a' && (out[0] | 0x20) <= 'z') && out[1] == ':') memmove(out, out + 2, strlen(out + 2) + 1);
    if (!out[0]) snprintf(out, n, ".");
}

/* ---------------------------------------------------------------- user32: rectangles, MessageBox */
static vb_i64 u_SetRect(args *x)
{
    vb_i32 *r = (vb_i32 *)a_ptr(x, 0, 16);
    if (!r) return 0;
    r[0] = (vb_i32)a_int(x, 1); r[1] = (vb_i32)a_int(x, 2); r[2] = (vb_i32)a_int(x, 3); r[3] = (vb_i32)a_int(x, 4);
    return 1;
}
static vb_i64 u_SetRectEmpty(args *x) { vb_i32 *r = (vb_i32 *)a_ptr(x, 0, 16); if (r) memset(r, 0, 16); return r != 0; }
static vb_i64 u_CopyRect(args *x)
{
    vb_i32 *d = (vb_i32 *)a_ptr(x, 0, 16), *s = (vb_i32 *)a_ptr(x, 1, 16);
    if (!d || !s) return 0;
    memmove(d, s, 16);
    return 1;
}
static vb_i64 u_InflateRect(args *x)
{
    vb_i32 *r = (vb_i32 *)a_ptr(x, 0, 16), dx = (vb_i32)a_int(x, 1), dy = (vb_i32)a_int(x, 2);
    if (!r) return 0;
    r[0] -= dx; r[1] -= dy; r[2] += dx; r[3] += dy;
    return 1;
}
static vb_i64 u_OffsetRect(args *x)
{
    vb_i32 *r = (vb_i32 *)a_ptr(x, 0, 16), dx = (vb_i32)a_int(x, 1), dy = (vb_i32)a_int(x, 2);
    if (!r) return 0;
    r[0] += dx; r[1] += dy; r[2] += dx; r[3] += dy;
    return 1;
}
static int rect_empty(const vb_i32 *r) { return r[2] <= r[0] || r[3] <= r[1]; }
static vb_i64 u_IntersectRect(args *x)
{
    vb_i32 *d = (vb_i32 *)a_ptr(x, 0, 16), *a = (vb_i32 *)a_ptr(x, 1, 16), *b = (vb_i32 *)a_ptr(x, 2, 16);
    if (!d || !a || !b) return 0;
    vb_i32 r[4] = { a[0] > b[0] ? a[0] : b[0], a[1] > b[1] ? a[1] : b[1], a[2] < b[2] ? a[2] : b[2], a[3] < b[3] ? a[3] : b[3] };
    if (rect_empty(r)) { memset(d, 0, 16); return 0; }
    memcpy(d, r, 16);
    return 1;
}
static vb_i64 u_UnionRect(args *x)
{
    vb_i32 *d = (vb_i32 *)a_ptr(x, 0, 16), *a = (vb_i32 *)a_ptr(x, 1, 16), *b = (vb_i32 *)a_ptr(x, 2, 16);
    if (!d || !a || !b) return 0;
    if (rect_empty(a)) { memmove(d, b, 16); return !rect_empty(b); }
    if (rect_empty(b)) { memmove(d, a, 16); return 1; }
    vb_i32 r[4] = { a[0] < b[0] ? a[0] : b[0], a[1] < b[1] ? a[1] : b[1], a[2] > b[2] ? a[2] : b[2], a[3] > b[3] ? a[3] : b[3] };
    memcpy(d, r, 16);
    return 1;
}
static vb_i64 u_PtInRect(args *x)
{
    vb_i32 *r = (vb_i32 *)a_ptr(x, 0, 16);
    vb_i32 px = (vb_i32)a_int(x, 1), py = (vb_i32)a_int(x, 2);
    if (x->n == 2) { vb_i32 *p = (vb_i32 *)a_ptr(x, 1, 8); if (!p) return 0; px = p[0]; py = p[1]; }
    return r && px >= r[0] && px < r[2] && py >= r[1] && py < r[3];
}
static vb_i64 u_EqualRect(args *x)
{
    vb_i32 *a = (vb_i32 *)a_ptr(x, 0, 16), *b = (vb_i32 *)a_ptr(x, 1, 16);
    return a && b && !memcmp(a, b, 16);
}
static vb_i64 u_IsRectEmpty(args *x) { vb_i32 *r = (vb_i32 *)a_ptr(x, 0, 16); return !r || rect_empty(r); }
static vb_i64 u_MessageBoxA(args *x)
{
    vb_i32 mk = vb_mark();
    char *t = a_str(x, 1), *c = a_str(x, 2);
    vb_i32 r = vb_msgbox(vb_sfromansi(t ? t : "", t ? (int)strlen(t) : 0), (vb_i32)(a_int(x, 3) & 0xFFFF),
                         vb_sfromansi(c ? c : "", c ? (int)strlen(c) : 0));
    vb_release(mk);
    return r;
}
static vb_i64 u_CharUpperA(args *x) { char *s = a_str(x, 0); if (s) for (char *p = s; *p; p++) *p = (char)toupper((unsigned char)*p); return x->n ? a_int(x, 0) : 0; }
static vb_i64 u_CharLowerA(args *x) { char *s = a_str(x, 0); if (s) for (char *p = s; *p; p++) *p = (char)tolower((unsigned char)*p); return x->n ? a_int(x, 0) : 0; }
static vb_i64 ret0(args *x) { (void)x; return 0; }
static vb_i64 ret1(args *x) { (void)x; return 1; }

/* ---------------------------------------------------------------- kernel32: names, drives, modules */
static vb_i64 k_GetShortPathNameA(args *x)
{
    char *s = a_str(x, 0);
    return put_text(a_str(x, 1), a_int(x, 2), s ? s : "");
}
static vb_i64 k_GetFullPathNameA(args *x)
{
    char *s = a_str(x, 0), p[2048], cwd[1500] = "";
    if (!s) return 0;
    if (s[0] == '/' || s[0] == '\\' || (s[0] && s[1] == ':')) snprintf(p, sizeof p, "%s", s);
    else { if (!getcwd(cwd, sizeof cwd)) cwd[0] = 0; snprintf(p, sizeof p, "%s/%s", cwd, s); }
    return put_text(a_str(x, 2), a_int(x, 1), p);
}
static vb_u32 volume_serial(void)
{
    FILE *f = fopen("/etc/machine-id", "r");
    char id[64] = "linbasic";
    if (f) { if (!fgets(id, sizeof id, f)) id[0] = 0; fclose(f); }
    vb_u32 h = 2166136261u;
    for (char *p = id; *p; p++) h = (h ^ (unsigned char)*p) * 16777619u;
    return h;
}
static vb_i64 k_GetVolumeInformationA(args *x)
{
    put_text(a_str(x, 1), a_int(x, 2), "");
    vb_i32 *serial = x->n > 3 ? a_long(x, 3) : 0, *maxc = x->n > 4 ? a_long(x, 4) : 0, *fl = x->n > 5 ? a_long(x, 5) : 0;
    if (x->n > 3 && x->a[3].kind == VBD_PTR && serial) *serial = (vb_i32)volume_serial();
    if (x->n > 4 && x->a[4].kind == VBD_PTR && maxc) *maxc = 255;
    if (x->n > 5 && x->a[5].kind == VBD_PTR && fl) *fl = 0x3;
    if (x->n > 6) put_text(a_str(x, 6), a_int(x, 7), "NTFS");
    return 1;
}
static vb_i64 k_GetLogicalDriveStringsA(args *x)
{
    char *b = a_str(x, 1);
    vb_i64 n = a_int(x, 0);
    if (!b || n < 5) return 4;
    memcpy(b, "C:\\\0\0", 5);
    return 4;
}
static vb_i64 k_GetLogicalDrives(args *x) { (void)x; return 4; }      /* C: */
static vb_i64 k_GetDriveTypeA(args *x)
{
    char *r = a_str(x, 0);
    if (r && *r && (r[0] | 0x20) != 'c' && r[1] == ':') return 1;    /* DRIVE_NO_ROOT_DIR */
    return 3;                                   /* DRIVE_FIXED */
}
static void disk_space(const char *root, unsigned long long *freeb, unsigned long long *total)
{
    char p[1024]; struct statvfs sv;
    lpath(root ? root : "/", p, sizeof p);
    if (p[0] == '/' && p[1] == 0) snprintf(p, sizeof p, "%s", getenv("HOME") ? getenv("HOME") : "/");
    if (statvfs(p, &sv) != 0 && statvfs("/", &sv) != 0) { *freeb = *total = 0; return; }
    *freeb = (unsigned long long)sv.f_bavail * sv.f_frsize; *total = (unsigned long long)sv.f_blocks * sv.f_frsize;
}
static vb_i64 k_GetDiskFreeSpaceA(args *x)
{
    unsigned long long fr, tot; disk_space(a_str(x, 0), &fr, &tot);
    vb_i32 *spc = a_long(x, 1), *bps = a_long(x, 2), *nfree = a_long(x, 3), *ntot = a_long(x, 4);
    if (spc) *spc = 8;
    if (bps) *bps = 512;
    unsigned long long c1 = fr / 4096, c2 = tot / 4096;
    if (nfree) *nfree = (vb_i32)(c1 > 0x7FFFFFFF ? 0x7FFFFFFF : c1);
    if (ntot) *ntot = (vb_i32)(c2 > 0x7FFFFFFF ? 0x7FFFFFFF : c2);
    return 1;
}
static vb_i64 k_GetDiskFreeSpaceExA(args *x)
{
    unsigned long long fr, tot; disk_space(a_str(x, 0), &fr, &tot);
    /* the 64-bit counts: as Currency (count / 10000) or as two Longs (a Type) - the same 8 bytes */
    for (int i = 1; i <= 3; i++) {
        unsigned char *p = (unsigned char *)a_ptr(x, i, 8);
        unsigned long long v = i == 2 ? tot : fr;
        if (p && x->a[i].kind == VBD_PTR) memcpy(p, &v, 8);
    }
    return 1;
}
static vb_i64 k_GetModuleHandleA(args *x) { (void)x; return 0x00400000; }
static vb_i64 k_GetModuleFileNameA(args *x)
{
    char p[1024] = "";
    ssize_t n = readlink("/proc/self/exe", p, sizeof p - 1);
    p[n > 0 ? n : 0] = 0;
    vb_i64 r = put_text(a_str(x, 1), a_int(x, 2), p);
    return r > a_int(x, 2) ? a_int(x, 2) : r;
}
static vb_i64 k_GetCurrentProcess(args *x) { (void)x; return -1; }
static vb_i64 k_GetCurrentThread(args *x) { (void)x; return -2; }
static vb_i64 k_GetCurrentThreadId(args *x) { (void)x; return (vb_i64)getpid(); }
static vb_i64 k_LoadLibraryA(args *x)
{
    char key[256], *n = a_str(x, 0);
    if (!n) return 0;
    lib_key(n, key, sizeof key);
    for (int i = 0; win_libs[i]; i++) if (!strcmp(win_libs[i], key)) return 0x10000000 + i * 0x10000;
    return 0;                                   /* another DLL: none on this computer */
}
static vb_i64 k_GetProcAddress(args *x) { (void)x; return 0; }
static vb_i64 k_GetVersion(args *x) { (void)x; return 0x1DB10106; }   /* 6.1 build 7601 */
static vb_i64 k_GetCommandLineA(args *x)
{
    static char *cl;
    (void)x;
    if (!cl) {
        char buf[4096] = "", exe[1024] = "";
        ssize_t n = readlink("/proc/self/exe", exe, sizeof exe - 1);
        exe[n > 0 ? n : 0] = 0;
        vb_i32 mk = vb_mark();
        snprintf(buf, sizeof buf, "\"%s\" %s", exe, vb_stoutf8(vb_command()));
        vb_release(mk);
        cl = strdup(buf);
    }
    return map_addr(cl, (vb_i32)strlen(cl) + 1);
}
static vb_i64 k_lstrcpynA(args *x)
{
    char *d = a_str(x, 0), *s = a_str(x, 1);
    vb_i64 n = a_int(x, 2);
    if (!d || !s || n <= 0) return 0;
    size_t k = strlen(s);
    if ((vb_i64)k > n - 1) k = (size_t)n - 1;
    memmove(d, s, k); d[k] = 0;
    return x->a[0].kind == VBD_INT ? x->a[0].i : 1;
}
static vb_i64 k_GetCurrentDirectoryA(args *x)
{
    char c[1500] = "";
    if (!getcwd(c, sizeof c)) c[0] = 0;
    return put_text(a_str(x, 1), a_int(x, 0), c);
}
static vb_i64 k_SetCurrentDirectoryA(args *x) { char p[1500]; lpath(a_str(x, 0), p, sizeof p); return chdir(p) == 0 ? 1 : seterr(w32err(errno, p), 0); }
static vb_i64 k_ExpandEnvironmentStringsA(args *x)
{
    char *s = a_str(x, 0), out[4096]; size_t k = 0;
    if (!s) return 0;
    for (const char *p = s; *p && k < sizeof out - 1;) {
        const char *e = *p == '%' ? strchr(p + 1, '%') : 0;
        if (e && e > p + 1) {
            char name[256]; size_t l = (size_t)(e - p - 1) < sizeof name - 1 ? (size_t)(e - p - 1) : sizeof name - 1;
            memcpy(name, p + 1, l); name[l] = 0;
            const char *v = getenv(name);
            if (!v && !strcasecmp(name, "windir")) v = win_dir(0);
            if (!v && !strcasecmp(name, "systemroot")) v = win_dir(0);
            if (!v && !strcasecmp(name, "userprofile")) v = getenv("HOME");
            if (!v && !strcasecmp(name, "temp")) v = getenv("TMPDIR") ? getenv("TMPDIR") : "/tmp";
            if (!v && !strcasecmp(name, "tmp")) v = getenv("TMPDIR") ? getenv("TMPDIR") : "/tmp";
            if (v) { for (; *v && k < sizeof out - 1; v++) out[k++] = *v; p = e + 1; continue; }
        }
        out[k++] = *p++;
    }
    out[k] = 0;
    vb_i64 r = put_text(a_str(x, 1), a_int(x, 2), out);
    return r <= a_int(x, 2) ? r + 1 : r;        /* the size with the 0 */
}

/* ---------------------------------------------------------------- kernel32: time, locale */
static void systime(unsigned char *p, int local)
{
    struct timespec ts; clock_gettime(CLOCK_REALTIME, &ts);
    time_t t = ts.tv_sec; struct tm tm;
    if (local) localtime_r(&t, &tm); else gmtime_r(&t, &tm);
    put16(p, (vb_u16)(tm.tm_year + 1900)); put16(p + 2, (vb_u16)(tm.tm_mon + 1)); put16(p + 4, (vb_u16)tm.tm_wday);
    put16(p + 6, (vb_u16)tm.tm_mday); put16(p + 8, (vb_u16)tm.tm_hour); put16(p + 10, (vb_u16)tm.tm_min);
    put16(p + 12, (vb_u16)tm.tm_sec); put16(p + 14, (vb_u16)(ts.tv_nsec / 1000000));
}
static vb_i64 k_GetSystemTime(args *x) { unsigned char *p = (unsigned char *)a_ptr(x, 0, 16); if (p) systime(p, 0); return 0; }
static vb_i64 k_GetLocalTime(args *x) { unsigned char *p = (unsigned char *)a_ptr(x, 0, 16); if (p) systime(p, 1); return 0; }
static vb_i64 k_GetTimeZoneInformation(args *x)
{
    unsigned char *p = (unsigned char *)a_ptr(x, 0, 172);
    if (!p) return 0;
    memset(p, 0, 172);
    time_t t = time(0); struct tm tm; localtime_r(&t, &tm);
    put32(p, (vb_u32)(vb_i32)(-tm.tm_gmtoff / 60));
    return tm.tm_isdst > 0 ? 2 : 1;
}
static vb_i64 k_GetUserDefaultLCID(args *x) { (void)x; return vb_lcid; }
static vb_i64 k_GetUserDefaultLangID(args *x) { (void)x; return vb_lcid & 0xFFFF; }
static vb_i64 k_GetLocaleInfoA(args *x)
{
    extern const char *vb_lc_sdate, *vb_lc_ldate, *vb_lc_ltime, *vb_lc_cursym, *vb_lc_am, *vb_lc_pm;
    vb_i32 t = (vb_i32)a_int(x, 1) & 0xFFFF;
    char v[128] = "";
    switch (t) {
    case 0x0E: v[0] = vb_lc_dec; v[1] = 0; break;                       /* LOCALE_SDECIMAL */
    case 0x0F: v[0] = vb_lc_grp; v[1] = 0; break;                       /* LOCALE_STHOUSAND */
    case 0x10: snprintf(v, sizeof v, "3;0"); break;                    /* LOCALE_SGROUPING */
    case 0x1F: snprintf(v, sizeof v, "%s", vb_lc_sdate); break;        /* LOCALE_SSHORTDATE */
    case 0x20: snprintf(v, sizeof v, "%s", vb_lc_ldate); break;        /* LOCALE_SLONGDATE */
    case 0x1003: snprintf(v, sizeof v, "%s", vb_lc_ltime); break;      /* LOCALE_STIMEFORMAT */
    case 0x14: snprintf(v, sizeof v, "%s", vb_lc_cursym); break;       /* LOCALE_SCURRENCY */
    case 0x28: snprintf(v, sizeof v, "%s", vb_lc_am); break;           /* LOCALE_S1159 */
    case 0x29: snprintf(v, sizeof v, "%s", vb_lc_pm); break;           /* LOCALE_S2359 */
    case 0x1D: snprintf(v, sizeof v, "%c", strchr(vb_lc_sdate, '.') ? '.' : strchr(vb_lc_sdate, '-') ? '-' : '/'); break;
    case 0x1E: snprintf(v, sizeof v, ":"); break;                      /* LOCALE_STIME */
    case 0x0C: snprintf(v, sizeof v, ";"); break;                      /* LOCALE_SLIST */
    case 0x01: snprintf(v, sizeof v, "%04x", vb_lcid); break;          /* LOCALE_ILANGUAGE */
    case 0x02: case 0x1001: snprintf(v, sizeof v, "%s", vb_lcid == 1032 ? "Greek" : vb_lcid == 1031 ? "German" :
                                       vb_lcid == 1036 ? "French" : "English"); break;
    case 0x06: case 0x1002: snprintf(v, sizeof v, "%s", vb_lcid == 1032 ? "Greece" : vb_lcid == 2057 ? "United Kingdom" :
                                       vb_lcid == 1031 ? "Germany" : "United States"); break;
    case 0x03: case 0x59: snprintf(v, sizeof v, "%s", vb_lcid == 1032 ? "ELL" : "ENU"); break;
    case 0x1004: snprintf(v, sizeof v, "%d", vb_lcid == 1032 ? 1253 : 1252); break;   /* LOCALE_IDEFAULTANSICODEPAGE */
    case 0x0B: snprintf(v, sizeof v, "%d", vb_lcid == 1032 ? 737 : 437); break;
    case 0x05: snprintf(v, sizeof v, "%d", vb_lcid == 1032 ? 30 : 1); break;          /* LOCALE_ICOUNTRY */
    default: return 0;
    }
    vb_i64 r = put_text(a_str(x, 2), a_int(x, 3), v);
    return a_int(x, 3) == 0 ? r : (r <= a_int(x, 3) ? r + 1 : 0);
}

/* ---------------------------------------------------------------- kernel32: files */
/* (on failure: the Windows error code for Err.LastDllError) */
static vb_i64 k_CreateDirectoryA(args *x)
{
    char p[1500]; lpath(a_str(x, 0), p, sizeof p);
    return mkdir(p, 0777) == 0 ? 1 : seterr(w32err(errno, p), 0);
}
static vb_i64 k_RemoveDirectoryA(args *x)
{
    char p[1500]; lpath(a_str(x, 0), p, sizeof p);
    return rmdir(p) == 0 ? 1 : seterr(errno == ENOTDIR ? 267 : w32err(errno, p), 0);
}
static vb_i64 k_DeleteFileA(args *x)
{
    char p[1500]; lpath(a_str(x, 0), p, sizeof p);
    return unlink(p) == 0 ? 1 : seterr(w32err(errno, p), 0);
}
static vb_i64 k_MoveFileA(args *x)
{
    char a[1500], b[1500]; lpath(a_str(x, 0), a, sizeof a); lpath(a_str(x, 1), b, sizeof b);
    if (access(b, F_OK) == 0 && access(a, F_OK) == 0) return seterr(183, 0);   /* MoveFile does not replace */
    return rename(a, b) == 0 ? 1 : seterr(w32err(errno, a), 0);
}
static int copy_file(const char *a, const char *b, int fail_if_exists)
{
    if (fail_if_exists && access(b, F_OK) == 0) return (int)seterr(80, 0);    /* ERROR_FILE_EXISTS */
    FILE *f = fopen(a, "rb"); if (!f) return (int)seterr(w32err(errno, a), 0);
    FILE *g = fopen(b, "wb"); if (!g) { int e = errno; fclose(f); return (int)seterr(w32err(e, b), 0); }
    char buf[65536]; size_t n;
    while ((n = fread(buf, 1, sizeof buf, f)) > 0) fwrite(buf, 1, n, g);
    fclose(f); fclose(g);
    return 1;
}
static vb_i64 k_CopyFileA(args *x)
{
    char a[1500], b[1500]; lpath(a_str(x, 0), a, sizeof a); lpath(a_str(x, 1), b, sizeof b);
    return copy_file(a, b, (int)a_int(x, 2));
}
static vb_i64 k_GetFileAttributesA(args *x)
{
    char p[1500]; struct stat st;
    lpath(a_str(x, 0), p, sizeof p);
    if (stat(p, &st) != 0) return seterr(w32err(errno, p), -1);    /* INVALID_FILE_ATTRIBUTES */
    vb_i64 a = S_ISDIR(st.st_mode) ? 0x10 : 0x20;
    if (access(p, W_OK) != 0) a |= 1;
    const char *b = strrchr(p, '/'); b = b ? b + 1 : p;
    if (b[0] == '.' && b[1]) a |= 2;
    return a;
}
static vb_i64 k_SetFileAttributesA(args *x)
{
    char p[1500]; struct stat st;
    lpath(a_str(x, 0), p, sizeof p);
    if (stat(p, &st) != 0) return seterr(w32err(errno, p), 0);
    mode_t m = st.st_mode;
    if (a_int(x, 1) & 1) m &= (mode_t)~0222; else m |= S_IWUSR;
    return chmod(p, m) == 0;
}
static vb_i64 k_GetTempFileNameA(args *x)
{
    char d[1500], out[1700]; static unsigned seq;
    lpath(a_str(x, 0), d, sizeof d);
    char *pre = a_str(x, 1);
    vb_u32 u = (vb_u32)a_int(x, 2);
    if (!u) u = ((vb_u32)getpid() + ++seq) & 0xFFFF;
    snprintf(out, sizeof out, "%s/%.3s%04X.tmp", d, pre ? pre : "tmp", u & 0xFFFF);
    if (!a_int(x, 2)) { FILE *f = fopen(out, "wb"); if (f) fclose(f); }
    put_text(a_str(x, 3), 260, out);
    return u;
}
/* CreateFile / ReadFile / WriteFile: handles of open files */
#define FH_BASE 0x00700000
static FILE *fh[64];
static vb_i64 k_CreateFileA(args *x)
{
    char p[1500]; lpath(a_str(x, 0), p, sizeof p);
    vb_u32 acc = (vb_u32)a_int(x, 1);
    vb_i32 disp = (vb_i32)a_int(x, 4);
    int exists = access(p, F_OK) == 0;
    if (disp == 1 && exists) return seterr(80, -1);                 /* CREATE_NEW: ERROR_FILE_EXISTS */
    if ((disp == 3 || disp == 5) && !exists) return seterr(w32err(ENOENT, p), -1);
    const char *mode = (acc & 0x40000000) ? ((disp == 2 || disp == 1 || disp == 5) ? "w+b" : (exists ? "r+b" : "w+b")) : "rb";
    if (!strcmp(p, "\\\\.\\PhysicalDrive0") || !strncmp(p, "//./", 4)) return -1;
    FILE *f = fopen(p, mode);
    if (!f) return seterr(w32err(errno, p), -1);
    if ((disp == 2 || disp == 4) && exists) emu_lasterr = 183;     /* CREATE_ALWAYS / OPEN_ALWAYS over a file */
    for (int i = 1; i < 64; i++) if (!fh[i]) { fh[i] = f; return FH_BASE + i; }
    fclose(f);
    return -1;
}
/* OpenFile(lpFileName, lpReOpenBuff As OFSTRUCT, wStyle): OF_EXIST, OF_DELETE, OF_READ / WRITE / READWRITE,
 * OF_CREATE; the OFSTRUCT gets the full path (szPathName, at offset 8) */
static vb_i64 k_OpenFile(args *x)
{
    char p[1500]; lpath(a_str(x, 0), p, sizeof p);
    vb_u32 st = (vb_u32)a_int(x, 2);
    unsigned char *of = (unsigned char *)a_ptr(x, 1, 136);
    if (vb_errpend) { vb_errpend = 0; of = 0; }
    if (of) { of[0] = 136; of[1] = 1; of[2] = of[3] = 0; memset(of + 8, 0, 128); snprintf((char *)of + 8, 128, "%s", a_str(x, 0)); }
    int exists = access(p, F_OK) == 0;
    if (st & 0x0200) {                                  /* OF_DELETE */
        if (remove(p) != 0) return seterr(w32err(errno, p), -1);
        return 1;
    }
    if (st & 0x4000) return exists ? 1 : seterr(2, -1); /* OF_EXIST: the file is there (and closed) */
    if (!exists && !(st & 0x1000)) return seterr(2, -1);
    const char *mode = (st & 0x1000) ? "w+b" : (st & 3) ? "r+b" : "rb";   /* OF_CREATE, OF_WRITE / READWRITE */
    FILE *f = fopen(p, mode);
    if (!f) return seterr(w32err(errno, p), -1);
    for (int i = 1; i < 64; i++) if (!fh[i]) { fh[i] = f; return FH_BASE + i; }
    fclose(f);
    return -1;
}
static FILE *fh_of(vb_i64 h) { vb_i64 i = h - FH_BASE; return i > 0 && i < 64 ? fh[i] : 0; }
static vb_i64 k_ReadFile(args *x)
{
    FILE *f = fh_of(a_int(x, 0));
    vb_i32 n = (vb_i32)a_int(x, 2);
    void *b = n > 0 ? a_ptr(x, 1, n) : 0;
    vb_i32 *done = x->n > 3 && x->a[3].kind == VBD_PTR ? a_long(x, 3) : 0;
    if (!f || !b) { if (done) *done = 0; return 0; }
    size_t k = fread(b, 1, (size_t)n, f);
    if (done) *done = (vb_i32)k;
    return 1;
}
static vb_i64 k_WriteFile(args *x)
{
    FILE *f = fh_of(a_int(x, 0));
    vb_i32 n = (vb_i32)a_int(x, 2);
    void *b = n > 0 ? a_ptr(x, 1, n) : 0;
    vb_i32 *done = x->n > 3 && x->a[3].kind == VBD_PTR ? a_long(x, 3) : 0;
    if (!f) { if (done) *done = 0; return 0; }
    size_t k = b ? fwrite(b, 1, (size_t)n, f) : 0;
    if (done) *done = (vb_i32)k;
    return 1;
}
static vb_i64 k_CloseHandleProc(vb_i32 h);
static vb_i64 k_CloseHandle(args *x)
{
    vb_i64 h = a_int(x, 0), i = h - FH_BASE;
    if (i > 0 && i < 64 && fh[i]) { fclose(fh[i]); fh[i] = 0; }
    else k_CloseHandleProc((vb_i32)h);
    return 1;
}
static vb_i64 k_GetFileSize(args *x)
{
    FILE *f = fh_of(a_int(x, 0));
    if (!f) return -1;
    long at = ftell(f); fseek(f, 0, SEEK_END); long n = ftell(f); fseek(f, at, SEEK_SET);
    vb_i32 *hi = x->n > 1 && x->a[1].kind == VBD_PTR ? a_long(x, 1) : 0;
    if (hi) *hi = 0;
    return (vb_i32)n;
}
static vb_i64 k_SetFilePointer(args *x)
{
    FILE *f = fh_of(a_int(x, 0));
    if (!f) return -1;
    int how = (int)a_int(x, 3);
    fseek(f, (long)(vb_i32)a_int(x, 1), how == 2 ? SEEK_END : how == 1 ? SEEK_CUR : SEEK_SET);
    return ftell(f);
}
static vb_i64 k_SetEndOfFile(args *x)
{
    FILE *f = fh_of(a_int(x, 0));
    if (!f) return 0;
    fflush(f);
    return ftruncate(fileno(f), ftell(f)) == 0;
}
static vb_i64 k_FlushFileBuffers(args *x) { FILE *f = fh_of(a_int(x, 0)); if (f) fflush(f); return f != 0; }

/* ---------------------------------------------------------------- kernel32: memory */
typedef struct { vb_i32 h; char *p; vb_i32 n; } gmem;
static gmem gm[256];
static vb_i64 k_GlobalAlloc(args *x)
{
    vb_i32 n = (vb_i32)a_int(x, 1);
    if (n < 0) return 0;
    char *p = (char *)calloc(1, (size_t)n + 1);
    if (!p) return 0;
    vb_i32 h = (vb_i32)map_addr(p, n + 1);
    for (int i = 0; i < 256; i++) if (!gm[i].p) { gm[i].h = h; gm[i].p = p; gm[i].n = n; break; }
    return h;
}
static vb_i64 k_GlobalLock(args *x) { return a_int(x, 0); }
static vb_i64 k_GlobalSize(args *x)
{
    vb_i32 h = (vb_i32)a_int(x, 0);
    for (int i = 0; i < 256; i++) if (gm[i].p && gm[i].h == h) return gm[i].n;
    return 0;
}
static vb_i64 k_GlobalFree(args *x)
{
    vb_i32 h = (vb_i32)a_int(x, 0);
    for (int i = 0; i < 256; i++) if (gm[i].p && gm[i].h == h) { gm[i].p = 0; break; }   /* the map keeps the memory */
    return 0;
}

/* ---------------------------------------------------------------- kernel32: processes, sync */
static vb_i64 k_WinExec(args *x)
{
    char *c = a_str(x, 0);
    if (!c) return 2;
    char cmd[4200];
    vb_i32 mk = vb_mark();
    snprintf(cmd, sizeof cmd, "(%s) >/dev/null 2>&1 &", vb_stoutf8(vb_sfromansi(c, (int)strlen(c))));
    vb_release(mk);
    for (char *p = cmd; *p; p++) if (*p == '\\') *p = '/';
    return system(cmd) == 0 ? 33 : 2;
}
static vb_i64 k_ExitProcess(args *x) { vb_files_flush(); exit((int)a_int(x, 0)); return 0; }
static vb_i64 k_OutputDebugStringA(args *x) { char *s = a_str(x, 0); if (s && getenv("VBRT_TRACE")) fprintf(stderr, "%s\n", s); return 0; }
static vb_i64 k_handle(args *x) { (void)x; static vb_i32 next = 0x00600010; next += 4; return next; }   /* CreateMutex, CreateEvent */
/* processes: CreateProcess / OpenProcess (by the number Shell gives) / WaitForSingleObject /
 * GetExitCodeProcess / TerminateProcess - the classic "Shell and wait" of VB6 programs */
#include <sys/wait.h>
#include <signal.h>
static struct { vb_i32 h; pid_t pid; int done, code; } procs[64];
static vb_i32 proc_handle(pid_t pid)
{
    static vb_i32 next = 0x00700010;
    for (int i = 0; i < 64; i++)
        if (!procs[i].h) { procs[i].h = (next += 4); procs[i].pid = pid; procs[i].done = 0; procs[i].code = 0; return procs[i].h; }
    return 0;
}
static int proc_of(vb_i32 h)
{
    for (int i = 0; i < 64; i++) if (procs[i].h && procs[i].h == h) return i;
    return -1;
}
static int proc_poll(int i)                     /* 1: it has ended */
{
    if (procs[i].done) return 1;
    int st;
    pid_t r = waitpid(procs[i].pid, &st, WNOHANG);
    if (r == procs[i].pid) { procs[i].done = 1; procs[i].code = WIFEXITED(st) ? WEXITSTATUS(st) : 1; return 1; }
    if (r < 0 && kill(procs[i].pid, 0) != 0) { procs[i].done = 1; return 1; }   /* not our child: gone */
    return 0;
}
static vb_i64 k_CreateProcessA(args *x)
{
    char *app = a_str(x, 0), *cmdl = a_str(x, 1), *dir = a_str(x, 7);
    unsigned char *pi = (unsigned char *)a_ptr(x, 9, 16);
    char cmd[4200] = "";
    vb_i32 mk = vb_mark();
    const char *c = cmdl && *cmdl ? cmdl : app;
    if (!c) { vb_release(mk); return 0; }
    snprintf(cmd, sizeof cmd, "%s", vb_stoutf8(vb_sfromansi(c, (int)strlen(c))));
    vb_release(mk);
    for (char *p = cmd; *p; p++) if (*p == '\\') *p = '/';
    char wd[1500] = "";
    if (dir && *dir) lpath(dir, wd, sizeof wd);
    vb_files_flush();
    pid_t pid = fork();
    if (pid < 0) return 0;
    if (pid == 0) {
        if (wd[0] && chdir(wd) != 0) _exit(127);
        execl("/bin/sh", "sh", "-c", cmd, (char *)0);
        _exit(127);
    }
    if (pi) {
        vb_i32 h = proc_handle(pid), t = h + 1, id = (vb_i32)pid;
        memcpy(pi, &h, 4); memcpy(pi + 4, &t, 4); memcpy(pi + 8, &id, 4); memcpy(pi + 12, &id, 4);
    }
    return 1;
}
static vb_i64 k_OpenProcess(args *x)
{
    pid_t pid = (pid_t)a_int(x, 2);
    if (pid <= 0 || kill(pid, 0) != 0) return 0;
    for (int i = 0; i < 64; i++) if (procs[i].h && procs[i].pid == pid) return procs[i].h;
    return proc_handle(pid);
}
static vb_i64 k_WaitForSingleObject(args *x)
{
    int i = proc_of((vb_i32)a_int(x, 0));
    if (i < 0) return 0;                        /* WAIT_OBJECT_0: mutexes, events */
    vb_i64 ms = a_int(x, 1);
    vb_files_flush();
    for (vb_i64 waited = 0; ; waited += 20) {
        if (proc_poll(i)) return 0;
        if (ms >= 0 && ms != 0xFFFFFFFF && waited >= ms) return 0x102;   /* WAIT_TIMEOUT */
        usleep(20000);
    }
}
static vb_i64 k_GetExitCodeProcess(args *x)
{
    int i = proc_of((vb_i32)a_int(x, 0));
    vb_i32 *code = a_long(x, 1);
    if (i < 0 || !code) return 0;
    *code = proc_poll(i) ? procs[i].code : 259;   /* STILL_ACTIVE */
    return 1;
}
static vb_i64 k_TerminateProcess(args *x)
{
    int i = proc_of((vb_i32)a_int(x, 0));
    if (i < 0) return 0;
    kill(procs[i].pid, SIGTERM);
    proc_poll(i);
    procs[i].code = (int)a_int(x, 1);
    return 1;
}
static vb_i64 k_CloseHandleProc(vb_i32 h)
{
    int i = proc_of(h);
    if (i < 0) return 0;
    procs[i].h = 0;
    return 1;
}

/* ---------------------------------------------------------------- FindFirstFileA / FindNextFileA / FindClose
 * WIN32_FIND_DATA (318 bytes in VB6's layout): dwFileAttributes, ftCreationTime, ftLastAccessTime,
 * ftLastWriteTime, nFileSizeHigh, nFileSizeLow, dwReserved0, dwReserved1, cFileName (260), cAlternate (14) */
#define FF_BASE 0x4F000
typedef struct { DIR *d; char dir[1500]; char pat[1500]; } findh;
static findh *ffh[32];
static void ft_put(unsigned char *p, time_t t)
{
    vb_i64 v = ((vb_i64)t + 11644473600LL) * 10000000LL;   /* 100 ns since 1601 */
    memcpy(p, &v, 8);
}
static int ff_next(findh *f, unsigned char *wfd)
{
    struct dirent *e;
    while ((e = readdir(f->d)) != 0) {
        if (fnmatch(f->pat, e->d_name, FNM_CASEFOLD) != 0) {
            if (!(strcmp(f->pat, "*.*") == 0)) continue;         /* "*.*" is every name, with or without a dot */
        }
        char full[1800]; struct stat st;
        snprintf(full, sizeof full, "%s/%s", f->dir, e->d_name);
        if (stat(full, &st) != 0) memset(&st, 0, sizeof st);
        memset(wfd, 0, 318);
        vb_i32 at = S_ISDIR(st.st_mode) ? 0x10 : 0x20;
        if (access(full, W_OK) != 0 && getuid() != 0) at |= 1;
        if (e->d_name[0] == '.' && strcmp(e->d_name, ".") && strcmp(e->d_name, "..")) at |= 2;
        memcpy(wfd, &at, 4);
        ft_put(wfd + 4, st.st_ctime); ft_put(wfd + 12, st.st_atime); ft_put(wfd + 20, st.st_mtime);
        vb_u32 hi = (vb_u32)((unsigned long long)st.st_size >> 32), lo = (vb_u32)st.st_size;
        memcpy(wfd + 28, &hi, 4); memcpy(wfd + 32, &lo, 4);
        put_ansi(e->d_name, (char *)wfd + 44, 260);         /* the name in the ANSI code page of the program */
        return 1;
    }
    return 0;
}
static vb_i64 k_FindFirstFileA(args *x)
{
    char p[1500];
    lpath(a_str(x, 0), p, sizeof p);
    unsigned char *wfd = (unsigned char *)a_ptr(x, 1, 318);
    if (!wfd) return -1;
    int slot = -1;
    for (int i = 1; i < 32; i++) if (!ffh[i]) { slot = i; break; }
    if (slot < 0) return -1;
    findh *f = (findh *)calloc(1, sizeof *f);
    char *b = strrchr(p, '/');
    if (b) { *b = 0; snprintf(f->dir, sizeof f->dir, "%s", *p ? p : "/"); snprintf(f->pat, sizeof f->pat, "%s", b + 1); }
    else { snprintf(f->dir, sizeof f->dir, "."); snprintf(f->pat, sizeof f->pat, "%s", p); }
    if (!f->pat[0]) snprintf(f->pat, sizeof f->pat, "*");
    if (!strpbrk(f->pat, "*?")) {                 /* one name: found by its own name, not by the case of the disk */
        DIR *d = opendir(f->dir);
        if (!d) { free(f); return seterr(3, -1); }
        f->d = d;
    } else if (!(f->d = opendir(f->dir))) { free(f); return seterr(3, -1); }
    if (!ff_next(f, wfd)) { closedir(f->d); free(f); return seterr(2, -1); }
    ffh[slot] = f;
    return FF_BASE + slot;
}
static vb_i64 k_FindNextFileA(args *x)
{
    vb_i64 i = a_int(x, 0) - FF_BASE;
    unsigned char *wfd = (unsigned char *)a_ptr(x, 1, 318);
    if (i < 1 || i >= 32 || !ffh[i] || !wfd) return seterr(6, 0);
    return ff_next(ffh[i], wfd) ? 1 : seterr(18, 0);                /* ERROR_NO_MORE_FILES */
}
static vb_i64 k_FindClose(args *x)
{
    vb_i64 i = a_int(x, 0) - FF_BASE;
    if (i < 1 || i >= 32 || !ffh[i]) return 0;
    closedir(ffh[i]->d); free(ffh[i]); ffh[i] = 0;
    return 1;
}

/* ---------------------------------------------------------------- the processes (Toolhelp32), from /proc
 * PROCESSENTRY32 (296 bytes): dwSize, cntUsage, th32ProcessID, th32DefaultHeapID, th32ModuleID,
 * cntThreads, th32ParentProcessID, pcPriClassBase, dwFlags, szExeFile (260) */
#define SNAP_HANDLE 0x4E001
typedef struct { int pid, ppid, threads; char name[260]; } psnap;
static psnap *snap; static int nsnap, isnap;
static vb_i64 k_CreateToolhelp32Snapshot(args *x)
{
    (void)x;
    free(snap); snap = 0; nsnap = isnap = 0;
    DIR *d = opendir("/proc");
    if (!d) return -1;
    struct dirent *e; int cap = 0;
    while ((e = readdir(d)) != 0) {
        if (!isdigit((unsigned char)e->d_name[0])) continue;
        char fn[300], buf[600]; FILE *f;
        snprintf(fn, sizeof fn, "/proc/%s/stat", e->d_name);
        if (!(f = fopen(fn, "r"))) continue;
        size_t n = fread(buf, 1, sizeof buf - 1, f); fclose(f); buf[n] = 0;
        char *l = strchr(buf, '('), *r = strrchr(buf, ')');
        if (!l || !r) continue;
        if (nsnap == cap) { cap = cap ? cap * 2 : 64; snap = (psnap *)realloc(snap, sizeof(psnap) * (size_t)cap); }
        psnap *p = &snap[nsnap++];
        memset(p, 0, sizeof *p);
        p->pid = atoi(e->d_name);
        snprintf(p->name, sizeof p->name, "%.*s", (int)(r - l - 1), l + 1);
        int ppid = 0, thr = 1; char st;
        long dummy;
        if (sscanf(r + 2, "%c %d %ld %ld %ld %ld %ld %ld %ld %ld %ld %ld %ld %ld %ld %ld %ld %d", &st, &ppid,
                   &dummy, &dummy, &dummy, &dummy, &dummy, &dummy, &dummy, &dummy, &dummy, &dummy, &dummy, &dummy,
                   &dummy, &dummy, &dummy, &thr) >= 2) { p->ppid = ppid; p->threads = thr; }
    }
    closedir(d);
    return SNAP_HANDLE;
}
static vb_i64 k_Process32(args *x, int first)
{
    if (a_int(x, 0) != SNAP_HANDLE) return 0;
    unsigned char *pe = (unsigned char *)a_ptr(x, 1, 296);
    if (!pe) return 0;
    if (first) isnap = 0;
    if (isnap >= nsnap) return 0;
    psnap *p = &snap[isnap++];
    vb_i32 v[9] = { 296, 0, p->pid, 0, 0, p->threads, p->ppid, 8, 0 };
    memcpy(pe, v, sizeof v);
    memset(pe + 36, 0, 260);
    memcpy(pe + 36, p->name, strnlen(p->name, 259));
    return 1;
}
static vb_i64 k_Process32First(args *x) { return k_Process32(x, 1); }
static vb_i64 k_Process32Next(args *x) { return k_Process32(x, 0); }
static vb_i64 k_GetPriorityClass(args *x) { (void)x; return 0x20; }     /* NORMAL_PRIORITY_CLASS */
/* winmm timeGetDevCaps: TIMECAPS { wPeriodMin, wPeriodMax } */
static vb_i64 w_timeGetDevCaps(args *x)
{
    vb_i32 *tc = (vb_i32 *)a_ptr(x, 0, 8);
    if (!tc) return 97;
    tc[0] = 1; tc[1] = 1000000;
    return 0;
}

/* ---------------------------------------------------------------- advapi32: the registry
 * kept in ~/.local/share/linbasic-studio/registry.txt: one line per value,
 *   <key path>\t<value name>\t<type>\t<hex data>   (a key without values: <path>\t\t-1\t) */
typedef struct { char *key, *name; vb_i32 type; unsigned char *data; vb_i32 n; } regval;
static regval *rv; static int nrv, caprv, reg_loaded;
static struct { vb_i32 h; char *path; } rkeys[256];
static const char *reg_file(void)
{
    static char f[700];
    if (!f[0]) {
        const char *x = getenv("XDG_DATA_HOME"), *h = getenv("HOME");
        char d[600];
        if (x && *x) snprintf(d, sizeof d, "%s/linbasic-studio", x); else snprintf(d, sizeof d, "%s/.local/share/linbasic-studio", h ? h : "/tmp");
        mkdir(d, 0777);
        snprintf(f, sizeof f, "%s/registry.txt", d);
    }
    return f;
}
static void reg_add(const char *key, const char *name, vb_i32 type, const unsigned char *d, vb_i32 n)
{
    if (nrv == caprv) { caprv = caprv ? caprv * 2 : 64; rv = (regval *)realloc(rv, sizeof(regval) * (size_t)caprv); }
    rv[nrv].key = strdup(key); rv[nrv].name = strdup(name); rv[nrv].type = type;
    rv[nrv].data = (unsigned char *)malloc((size_t)n + 1); if (n) memcpy(rv[nrv].data, d, (size_t)n); rv[nrv].data[n] = 0;
    rv[nrv].n = n; nrv++;
}
static int reg_find(const char *key, const char *name)
{
    for (int i = 0; i < nrv; i++) if (!strcasecmp(rv[i].key, key) && !strcasecmp(rv[i].name, name) && rv[i].type >= 0) return i;
    return -1;
}
static int reg_key_exists(const char *key)
{
    size_t n = strlen(key);
    for (int i = 0; i < nrv; i++)
        if (!strncasecmp(rv[i].key, key, n) && (rv[i].key[n] == 0 || rv[i].key[n] == '\\')) return 1;
    return 0;
}
static void reg_set_sz(const char *key, const char *name, const char *v) { if (reg_find(key, name) < 0) reg_add(key, name, 1, (const unsigned char *)v, (vb_i32)strlen(v) + 1); }
static void reg_load(void)
{
    if (reg_loaded) return;
    reg_loaded = 1;
    FILE *f = fopen(reg_file(), "r");
    if (f) {
        char line[66000];
        while (fgets(line, sizeof line, f)) {
            char *k = line, *t1 = strchr(k, '\t'); if (!t1) continue; *t1 = 0;
            char *nm = t1 + 1, *t2 = strchr(nm, '\t'); if (!t2) continue; *t2 = 0;
            char *ty = t2 + 1, *t3 = strchr(ty, '\t'); if (!t3) continue; *t3 = 0;
            char *hx = t3 + 1; hx[strcspn(hx, "\r\n")] = 0;
            vb_i32 n = (vb_i32)strlen(hx) / 2; unsigned char *d = (unsigned char *)malloc((size_t)n + 1);
            for (vb_i32 i = 0; i < n; i++) { unsigned v; sscanf(hx + 2 * i, "%2x", &v); d[i] = (unsigned char)v; }
            reg_add(k, nm, atoi(ty), d, n);
            free(d);
        }
        fclose(f);
    }
    /* what programs read of Windows */
    const char *u = getenv("USER");
    reg_set_sz("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", "ProductName", "Windows 7 Professional");
    reg_set_sz("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", "CurrentVersion", "6.1");
    reg_set_sz("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", "CurrentBuildNumber", "7601");
    reg_set_sz("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", "RegisteredOwner", u ? u : "user");
    reg_set_sz("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", "RegisteredOrganization", "");
    reg_set_sz("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion", "ProgramFilesDir", "C:\\Program Files");
    reg_set_sz("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion", "CommonFilesDir", "C:\\Program Files\\Common Files");
    reg_set_sz("HKEY_LOCAL_MACHINE\\HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0", "ProcessorNameString", "x86 Processor");
    reg_set_sz("HKEY_LOCAL_MACHINE\\HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0", "VendorIdentifier", "GenuineIntel");
    if (reg_find("HKEY_LOCAL_MACHINE\\HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0", "~MHz") < 0) {
        vb_i32 mhz = 2000;
        FILE *c = fopen("/proc/cpuinfo", "r");
        if (c) { char l[512]; while (fgets(l, sizeof l, c)) if (!strncmp(l, "cpu MHz", 7)) { char *p = strchr(l, ':'); if (p) mhz = (vb_i32)atof(p + 1); break; } fclose(c); }
        reg_add("HKEY_LOCAL_MACHINE\\HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0", "~MHz", 4, (unsigned char *)&mhz, 4);
    }
    reg_set_sz("HKEY_CURRENT_USER\\Control Panel\\Desktop", "Wallpaper", "");
}
static void reg_save(void)
{
    FILE *f = fopen(reg_file(), "w");
    if (!f) return;
    for (int i = 0; i < nrv; i++) {
        if (!strncasecmp(rv[i].key, "HKEY_LOCAL_MACHINE\\HARDWARE", 27)) continue;
        fprintf(f, "%s\t%s\t%d\t", rv[i].key, rv[i].name, rv[i].type);
        for (vb_i32 k = 0; k < rv[i].n; k++) fprintf(f, "%02x", rv[i].data[k]);
        fputc('\n', f);
    }
    fclose(f);
}
static const char *reg_root(vb_i64 h)
{
    switch ((vb_u32)h) {
    case 0x80000000u: return "HKEY_CLASSES_ROOT";
    case 0x80000001u: return "HKEY_CURRENT_USER";
    case 0x80000002u: return "HKEY_LOCAL_MACHINE";
    case 0x80000003u: return "HKEY_USERS";
    case 0x80000005u: return "HKEY_CURRENT_CONFIG";
    }
    for (int i = 0; i < 256; i++) if (rkeys[i].path && rkeys[i].h == (vb_i32)h) return rkeys[i].path;
    return 0;
}
static void reg_join(char *out, size_t n, const char *base, const char *sub)
{
    snprintf(out, n, "%s%s%s", base, sub && *sub ? "\\" : "", sub ? sub : "");
    size_t l = strlen(out);
    while (l && out[l - 1] == '\\') out[--l] = 0;
}
static vb_i32 reg_handle(const char *path)
{
    static vb_i32 next = 0x00500010;
    for (int i = 0; i < 256; i++) if (!rkeys[i].path) { rkeys[i].path = strdup(path); rkeys[i].h = next; next += 4; return rkeys[i].h; }
    return 0;
}
static vb_i64 reg_open(args *x, int sub, int out, int create)
{
    reg_load();
    const char *base = reg_root(a_int(x, 0));
    vb_i32 *ph = a_long(x, out);
    if (!base) return 6;                          /* ERROR_INVALID_HANDLE */
    char p[2048];
    reg_join(p, sizeof p, base, a_str(x, sub));
    if (!reg_key_exists(p) && !strchr(p, '\\')) {}         /* a root key: always there */
    else if (!reg_key_exists(p)) {
        if (!create) { if (ph) *ph = 0; return 2; } /* ERROR_FILE_NOT_FOUND */
        reg_add(p, "", -1, (const unsigned char *)"", 0);
        reg_save();
    }
    if (ph) *ph = reg_handle(p);
    return 0;
}
static vb_i64 a_RegOpenKeyA(args *x) { return reg_open(x, 1, 2, 0); }
static vb_i64 a_RegOpenKeyExA(args *x) { return reg_open(x, 1, 4, 0); }
static vb_i64 a_RegCreateKeyA(args *x) { return reg_open(x, 1, 2, 1); }
static vb_i64 a_RegCreateKeyExA(args *x)
{
    vb_i64 r = reg_open(x, 1, 7, 1);
    vb_i32 *disp = x->n > 8 && x->a[8].kind == VBD_PTR ? a_long(x, 8) : 0;
    if (disp) *disp = 1;
    return r;
}
static vb_i64 a_RegCloseKey(args *x)
{
    vb_i32 h = (vb_i32)a_int(x, 0);
    for (int i = 0; i < 256; i++) if (rkeys[i].path && rkeys[i].h == h) { free(rkeys[i].path); rkeys[i].path = 0; }
    return 0;
}
static vb_i64 a_RegSetValueExA(args *x)
{
    reg_load();
    const char *k = reg_root(a_int(x, 0));
    if (!k) return 6;
    char *name = a_str(x, 1);
    vb_i32 type = (vb_i32)a_int(x, 3), n = (vb_i32)a_int(x, 5);
    const unsigned char *d;
    if (x->a[4].kind == VBD_STR) { d = (const unsigned char *)x->ansi[4]; if (type == 1 || type == 2) n = d ? (vb_i32)strlen((const char *)d) + 1 : 1; }
    else d = (const unsigned char *)a_ptr(x, 4, n > 0 ? n : 1);
    if (!d) { d = (const unsigned char *)""; n = 0; }
    int i = reg_find(k, name ? name : "");
    if (i >= 0) { free(rv[i].data); rv[i].data = (unsigned char *)malloc((size_t)n + 1); memcpy(rv[i].data, d, (size_t)n); rv[i].data[n] = 0; rv[i].n = n; rv[i].type = type; }
    else reg_add(k, name ? name : "", type, d, n);
    reg_save();
    return 0;
}
static vb_i64 a_RegSetValueA(args *x)       /* RegSetValue(hKey, subkey, REG_SZ, data, n): the default value */
{
    reg_load();
    const char *base = reg_root(a_int(x, 0));
    if (!base) return 6;
    char p[2048]; reg_join(p, sizeof p, base, a_str(x, 1));
    char *v = a_str(x, 3);
    int i = reg_find(p, "");
    if (i >= 0) { free(rv[i].data); rv[i].data = (unsigned char *)strdup(v ? v : ""); rv[i].n = (vb_i32)strlen(v ? v : "") + 1; rv[i].type = 1; }
    else reg_add(p, "", 1, (const unsigned char *)(v ? v : ""), (vb_i32)strlen(v ? v : "") + 1);
    reg_save();
    return 0;
}
static vb_i64 a_RegQueryValueExA(args *x)
{
    reg_load();
    const char *k = reg_root(a_int(x, 0));
    if (!k) return 6;
    char *name = a_str(x, 1);
    int i = reg_find(k, name ? name : "");
    vb_i32 *type = x->n > 3 && x->a[3].kind == VBD_PTR ? a_long(x, 3) : 0;
    vb_i32 *size = x->n > 5 ? a_long(x, 5) : 0;
    if (i < 0) return 2;
    if (type) *type = rv[i].type;
    vb_i32 room = size ? *size : 0;
    unsigned char *d = x->n > 4 && !(x->a[4].kind == VBD_INT && x->a[4].i == 0) ? (unsigned char *)a_ptr(x, 4, room > 0 ? room : 1) : 0;
    if (x->n > 4 && x->a[4].kind == VBD_STR) d = (unsigned char *)x->ansi[4];
    if (size) *size = rv[i].n;
    if (!d) return 0;
    if (room < rv[i].n) return 234;               /* ERROR_MORE_DATA */
    memcpy(d, rv[i].data, (size_t)rv[i].n);
    return 0;
}
static vb_i64 a_RegQueryValueA(args *x)
{
    reg_load();
    const char *base = reg_root(a_int(x, 0));
    if (!base) return 6;
    char p[2048]; reg_join(p, sizeof p, base, a_str(x, 1));
    int i = reg_find(p, "");
    vb_i32 *size = a_long(x, 3);
    if (i < 0) return 2;
    char *b = a_str(x, 2);
    if (b && size && *size >= rv[i].n) memcpy(b, rv[i].data, (size_t)rv[i].n);
    if (size) *size = rv[i].n;
    return 0;
}
static vb_i64 a_RegDeleteValueA(args *x)
{
    reg_load();
    const char *k = reg_root(a_int(x, 0));
    char *name = a_str(x, 1);
    int i = k ? reg_find(k, name ? name : "") : -1;
    if (i < 0) return 2;
    rv[i] = rv[--nrv];
    reg_save();
    return 0;
}
static vb_i64 a_RegDeleteKeyA(args *x)
{
    reg_load();
    const char *base = reg_root(a_int(x, 0));
    if (!base) return 6;
    char p[2048]; reg_join(p, sizeof p, base, a_str(x, 1));
    size_t n = strlen(p); int found = 0;
    for (int i = 0; i < nrv; i++)
        if (!strncasecmp(rv[i].key, p, n) && (rv[i].key[n] == 0 || rv[i].key[n] == '\\')) { rv[i] = rv[--nrv]; i--; found = 1; }
    if (found) reg_save();
    return found ? 0 : 2;
}
/* the sub-keys of a key, in order: RegEnumKey(Ex) */
static vb_i64 a_RegEnumKeyA(args *x)
{
    reg_load();
    const char *k = reg_root(a_int(x, 0));
    if (!k) return 6;
    vb_i32 want = (vb_i32)a_int(x, 1), seen = 0;
    size_t n = strlen(k);
    char names[256][260]; int nn = 0;
    for (int i = 0; i < nrv && nn < 256; i++) {
        if (strncasecmp(rv[i].key, k, n) || rv[i].key[n] != '\\') continue;
        const char *rest = rv[i].key + n + 1;
        size_t l = strcspn(rest, "\\"); if (l >= 260) l = 259;
        int dup = 0;
        for (int j = 0; j < nn; j++) if (strlen(names[j]) == l && !strncasecmp(names[j], rest, l)) dup = 1;
        if (!dup) { memcpy(names[nn], rest, l); names[nn][l] = 0; nn++; }
    }
    (void)seen;
    if (want < 0 || want >= nn) return 259;      /* ERROR_NO_MORE_ITEMS */
    char *b = a_str(x, 2);
    vb_i32 *psize = x->a[3].kind == VBD_PTR ? a_long(x, 3) : 0;
    vb_i64 room = psize ? *psize : a_int(x, 3);
    put_text(b, room, names[want]);
    if (psize) *psize = (vb_i32)strlen(names[want]);
    return 0;
}
static vb_i64 a_RegEnumValueA(args *x)
{
    reg_load();
    const char *k = reg_root(a_int(x, 0));
    if (!k) return 6;
    vb_i32 want = (vb_i32)a_int(x, 1), seen = 0;
    for (int i = 0; i < nrv; i++) {
        if (strcasecmp(rv[i].key, k) || rv[i].type < 0) continue;
        if (seen++ != want) continue;
        vb_i32 *nsize = a_long(x, 3);
        put_text(a_str(x, 2), nsize ? *nsize : 0, rv[i].name);
        if (nsize) *nsize = (vb_i32)strlen(rv[i].name);
        vb_i32 *type = x->n > 5 && x->a[5].kind == VBD_PTR ? a_long(x, 5) : 0;
        if (type) *type = rv[i].type;
        vb_i32 *dsize = x->n > 7 && x->a[7].kind == VBD_PTR ? a_long(x, 7) : 0;
        unsigned char *d = x->n > 6 && !(x->a[6].kind == VBD_INT && x->a[6].i == 0) ? (unsigned char *)a_ptr(x, 6, dsize && *dsize > 0 ? *dsize : 1) : 0;
        if (x->n > 6 && x->a[6].kind == VBD_STR) d = (unsigned char *)x->ansi[6];
        if (d && dsize && *dsize >= rv[i].n) memcpy(d, rv[i].data, (size_t)rv[i].n);
        if (dsize) *dsize = rv[i].n;
        return 0;
    }
    return 259;
}
static vb_i64 a_RegQueryInfoKeyA(args *x) { (void)x; return 0; }

/* ---------------------------------------------------------------- shell32 */
static vb_i64 s_ShellExecuteA(args *x)
{
    char *op = a_str(x, 1), *file = a_str(x, 2), *par = a_str(x, 3);
    if (!file || !*file) return 2;
    char f[1500], cmd[3500];
    int url = strstr(file, "://") != 0 || !strncasecmp(file, "mailto:", 7) || !strncasecmp(file, "www.", 4);
    if (url) snprintf(f, sizeof f, "%s%s", !strncasecmp(file, "www.", 4) ? "http://" : "", file);
    else lpath(file, f, sizeof f);
    const char *ext = strrchr(f, '.');
    int exe = !url && ext && (!strcasecmp(ext, ".exe") || !strcasecmp(ext, ".bat") || !strcasecmp(ext, ".com"));
    if (op && !strcasecmp(op, "explore")) exe = 0;
    if (!url && access(f, F_OK) != 0) return 2;  /* SE_ERR_FNF */
    if (exe) snprintf(cmd, sizeof cmd, "('%s' %s) >/dev/null 2>&1 &", f, par ? par : "");
    else snprintf(cmd, sizeof cmd, "xdg-open '%s' >/dev/null 2>&1 &", f);
    return system(cmd) == 0 ? 42 : 31;
}
/* the special folders (CSIDL) and their "item lists": numbers that stand for a path */
static char *pidls[256];
vb_i32 vb_pidl_make(const char *path)
{
    for (int i = 1; i < 256; i++) if (!pidls[i]) { pidls[i] = strdup(path); return 0x00800000 + i * 16; }
    free(pidls[1]); pidls[1] = strdup(path);
    return 0x00800000 + 16;
}
static const char *pidl_path(vb_i64 p) { vb_i64 i = (p - 0x00800000) / 16; return i > 0 && i < 256 ? pidls[i] : 0; }
static void csidl_path(int c, char *out, size_t n)
{
    const char *h = getenv("HOME"); if (!h) h = "/tmp";
    switch (c & 0xFF) {
    case 0x00: case 0x10: case 0x19: snprintf(out, n, "%s/Desktop", h); break;
    case 0x05: case 0x2E: snprintf(out, n, "%s/Documents", h); break;
    case 0x0D: snprintf(out, n, "%s/Music", h); break;
    case 0x27: snprintf(out, n, "%s/Pictures", h); break;
    case 0x0E: snprintf(out, n, "%s/Videos", h); break;
    case 0x1A: case 0x23: snprintf(out, n, "%s/.config", h); break;
    case 0x1C: snprintf(out, n, "%s/.local/share", h); break;
    case 0x24: snprintf(out, n, "%s", win_dir(0)); break;
    case 0x25: case 0x29: snprintf(out, n, "%s", win_dir(1)); break;
    case 0x14: snprintf(out, n, "/usr/share/fonts"); break;
    case 0x26: case 0x2A: snprintf(out, n, "/usr"); break;
    case 0x20: snprintf(out, n, "%s/.cache", h); break;
    default: snprintf(out, n, "%s", h);
    }
}
static vb_i64 s_SHGetSpecialFolderLocation(args *x)
{
    char p[1024]; csidl_path((int)a_int(x, 1), p, sizeof p);
    vb_i32 *out = a_long(x, 2);
    if (out) *out = vb_pidl_make(p);
    return 0;
}
static vb_i64 s_SHGetPathFromIDListA(args *x)
{
    const char *p = pidl_path(a_int(x, 0));
    char *b = a_str(x, 1);
    if (!p || !b) { if (b) b[0] = 0; return 0; }
    put_text(b, 260, p);
    return 1;
}
static vb_i64 s_SHGetSpecialFolderPathA(args *x)
{
    char p[1024]; csidl_path((int)a_int(x, 2), p, sizeof p);
    if (a_int(x, 3)) mkdir(p, 0777);
    put_text(a_str(x, 1), 260, p);
    return 1;
}
static vb_i64 s_SHGetFolderPathA(args *x)
{
    char p[1024]; csidl_path((int)a_int(x, 1), p, sizeof p);
    put_text(a_str(x, 4), 260, p);
    return 0;
}
/* SHFileOperation: SHFILEOPSTRUCT (wFunc 1 move, 2 copy, 3 delete, 4 rename; pFrom/pTo lists) */
static vb_i64 s_SHFileOperationA(args *x)
{
    unsigned char *s = (unsigned char *)a_ptr(x, 0, 30);
    if (!s) return 1;
    vb_i32 func, from, to;
    memcpy(&func, s + 4, 4); memcpy(&from, s + 8, 4); memcpy(&to, s + 12, 4);
    const char *pf = from ? (const char *)vb_ptr_resolve(from, 1) : 0, *pt = to ? (const char *)vb_ptr_resolve(to, 1) : 0;
    if (!pf) return 1;
    int bad = 0;
    for (const char *f = pf; *f; f += strlen(f) + 1) {
        char a[1500], b[1500]; lpath(f, a, sizeof a);
        if (func == 3) { if (remove(a) != 0) bad = 1; continue; }
        if (!pt) { bad = 1; continue; }
        lpath(pt, b, sizeof b);
        struct stat st;
        if (stat(b, &st) == 0 && S_ISDIR(st.st_mode)) {
            const char *base = strrchr(a, '/'); base = base ? base + 1 : a;
            size_t l = strlen(b), k = strlen(base); if (l + 1 + k < sizeof b) { b[l] = '/'; memcpy(b + l + 1, base, k + 1); }
        }
        if (func == 2) { if (!copy_file(a, b, 0)) bad = 1; }
        else if (rename(a, b) != 0) bad = 1;
    }
    return bad;
}
static vb_i64 s_ShellAboutA(args *x)
{
    vb_i32 mk = vb_mark();
    char *app = a_str(x, 1), *other = a_str(x, 2), t[1024];
    snprintf(t, sizeof t, "%s\n%s", app ? app : "", other ? other : "");
    vb_msgbox(vb_sfromansi(t, (int)strlen(t)), 64, vb_sfromansi(app ? app : "", app ? (int)strlen(app) : 0));
    vb_release(mk);
    return 1;
}
static vb_i64 o_CoCreateGuid(args *x)
{
    unsigned char *g = (unsigned char *)a_ptr(x, 0, 16);
    if (!g) return 0x80004003;
    FILE *f = fopen("/dev/urandom", "rb");
    if (!f || fread(g, 1, 16, f) != 16) for (int i = 0; i < 16; i++) g[i] = (unsigned char)rand();
    if (f) fclose(f);
    g[7] = (unsigned char)((g[7] & 0x0F) | 0x40); g[8] = (unsigned char)((g[8] & 0x3F) | 0x80);
    return 0;
}
static vb_i64 i_InternetGetConnectedState(args *x)
{
    vb_i32 *fl = a_long(x, 0);
    if (fl && x->a[0].kind == VBD_PTR) *fl = 0x02;      /* INTERNET_CONNECTION_LAN */
    return 1;
}
static vb_i64 w_waveOutGetVolume(args *x) { vb_i32 *v = a_long(x, 1); if (v && x->a[1].kind == VBD_PTR) *v = -1; return 0; }


/* ---------------------------------------------------------------- more of kernel32, shlwapi, msvbvm60, winmm */
static vb_i64 k_SleepEx(args *x) { vb_files_flush(); usleep((useconds_t)(a_int(x, 0) * 1000)); return 0; }
static vb_i64 k_GetStdHandle(args *x) { vb_i32 n = (vb_i32)a_int(x, 0); return n == -10 ? 0x00600001 : n == -11 ? 0x00600002 : 0x00600003; }
static vb_i64 k_GlobalAddAtomA(args *x) { (void)x; static vb_i32 next = 0xC000; return next++; }
static vb_i64 k_GetModuleHandleW(args *x) { (void)x; return 0x00400000; }
static vb_i64 k_IsProcessorFeaturePresent(args *x) { vb_i32 f = (vb_i32)a_int(x, 0); return f == 6 || f == 10 || f == 3; }   /* MMX, SSE2, ... */
static vb_i64 k_GetSystemPowerStatus(args *x)
{
    unsigned char *p = (unsigned char *)a_ptr(x, 0, 12);
    if (!p) return 0;
    memset(p, 0, 12);
    p[0] = 1; p[1] = 255; p[2] = 255;           /* on AC power, battery unknown */
    vb_i32 m1 = -1; memcpy(p + 4, &m1, 4); memcpy(p + 8, &m1, 4);
    return 1;
}
static vb_i64 k_GetComputerNameExA(args *x)
{
    char host[256] = "LINUX";
    gethostname(host, sizeof host - 1);
    char *buf = a_str(x, 1);
    vb_i32 *n = (vb_i32 *)a_ptr(x, 2, 4);
    if (!buf || !n) return 0;
    size_t l = strlen(host);
    if ((vb_i32)l >= *n) { *n = (vb_i32)l + 1; return 0; }
    memcpy(buf, host, l + 1); *n = (vb_i32)l;
    return 1;
}
static vb_i64 sh_PathFileExistsA(args *x)
{
    char *s = a_str(x, 0), p[1500];
    if (!s) return 0;
    lpath(s, p, sizeof p);
    struct stat st;
    return stat(p, &st) == 0;
}
static vb_i64 sh_PathIsDirectoryA(args *x)
{
    char *s = a_str(x, 0), p[1500];
    if (!s) return 0;
    lpath(s, p, sizeof p);
    struct stat st;
    return stat(p, &st) == 0 && S_ISDIR(st.st_mode) ? 16 : 0;
}
static vb_i64 sh_DllGetVersion(args *x)
{
    unsigned char *p = (unsigned char *)a_ptr(x, 0, 20);
    if (!p) return 0x80004005;
    vb_i32 v[4] = { 6, 0, 2900, 0 };            /* the shell of Windows XP and later */
    memcpy(p + 4, v, 12);
    return 0;
}
/* VarPtr / VarPtrArray of msvbvm60: the address of the variable, as the VarPtr of VB */
static vb_i64 v_VarPtr(args *x) { void *p = a_ptr(x, 0, 4); return p ? vb_varptr(p, 16) : 0; }
static vb_i64 v_vbaCopyBytes(args *x)
{
    vb_i32 n = (vb_i32)a_int(x, 0);
    if (bstr_copy(x, 1, 2, n)) return 0;
    void *d = n > 0 ? a_ptr(x, 1, n) : 0, *s = n > 0 ? a_ptr(x, 2, n) : 0;
    if (d && s) memmove(d, s, (size_t)n);
    return 0;
}
static vb_i64 v_vbaCopyBytesZero(args *x)
{
    vb_i32 n = (vb_i32)a_int(x, 0);
    void *d = n > 0 ? a_ptr(x, 1, n) : 0, *s = n > 0 ? a_ptr(x, 2, n) : 0;
    if (d && s) { memmove(d, s, (size_t)n); memset(s, 0, (size_t)n); }
    return 0;
}
static vb_i64 v_GetMem4(args *x) { if (bstr_copy(x, 1, 0, 4)) return 0; void *s = a_ptr(x, 0, 4), *d = a_ptr(x, 1, 4); if (s && d) memmove(d, s, 4); return 0; }
static vb_i64 v_GetMem2(args *x) { void *s = a_ptr(x, 0, 2), *d = a_ptr(x, 1, 2); if (s && d) memmove(d, s, 2); return 0; }
static vb_i64 v_GetMem1(args *x) { void *s = a_ptr(x, 0, 1), *d = a_ptr(x, 1, 1); if (s && d) memmove(d, s, 1); return 0; }
static vb_i64 v_PutMem4(args *x) { vbs *ds = str_slot(x, 0); if (ds) { str_put_ptr(ds, (vb_i32)a_int(x, 1)); return 0; } vb_i32 *d = (vb_i32 *)a_ptr(x, 0, 4); if (d) *d = (vb_i32)a_int(x, 1); return 0; }
static vb_i64 v_PutMem2(args *x) { vb_i16 *d = (vb_i16 *)a_ptr(x, 0, 2); if (d) *d = (vb_i16)a_int(x, 1); return 0; }
static vb_i64 v_PutMem1(args *x) { vb_u8 *d = (vb_u8 *)a_ptr(x, 0, 1); if (d) *d = (vb_u8)a_int(x, 1); return 0; }

static const struct { const char *lib, *name; apifn fn; } apis2[] = {
    { "user32", "SetRect", u_SetRect }, { "user32", "SetRectEmpty", u_SetRectEmpty }, { "user32", "CopyRect", u_CopyRect },
    { "user32", "InflateRect", u_InflateRect }, { "user32", "OffsetRect", u_OffsetRect },
    { "user32", "IntersectRect", u_IntersectRect }, { "user32", "UnionRect", u_UnionRect },
    { "user32", "PtInRect", u_PtInRect }, { "user32", "EqualRect", u_EqualRect }, { "user32", "IsRectEmpty", u_IsRectEmpty },
    { "user32", "MessageBoxA", u_MessageBoxA }, { "user32", "MessageBox", u_MessageBoxA },
    { "user32", "CharUpperA", u_CharUpperA }, { "user32", "CharLowerA", u_CharLowerA },
    { "user32", "ExitWindowsEx", ret0 }, { "user32", "SwapMouseButton", ret0 }, { "user32", "ClipCursor", ret1 },
    { "user32", "RegisterHotKey", ret1 }, { "user32", "UnregisterHotKey", ret1 }, { "user32", "WinHelpA", ret0 },
    { "user32", "SetWindowsHookExA", k_handle }, { "user32", "SetWindowsHookEx", k_handle },
    { "user32", "UnhookWindowsHookEx", ret1 }, { "user32", "CallNextHookEx", ret0 },
    { "comctl32", "InitCommonControls", ret1 }, { "comctl32", "InitCommonControlsEx", ret1 },
    { "kernel32", "GetShortPathNameA", k_GetShortPathNameA }, { "kernel32", "GetLongPathNameA", k_GetShortPathNameA },
    { "kernel32", "GetFullPathNameA", k_GetFullPathNameA },
    { "kernel32", "GetVolumeInformationA", k_GetVolumeInformationA }, { "kernel32", "GetVolumeInformation", k_GetVolumeInformationA },
    { "kernel32", "GetLogicalDriveStringsA", k_GetLogicalDriveStringsA }, { "kernel32", "GetLogicalDrives", k_GetLogicalDrives },
    { "kernel32", "GetDriveTypeA", k_GetDriveTypeA }, { "kernel32", "GetDriveType", k_GetDriveTypeA },
    { "kernel32", "GetDiskFreeSpaceA", k_GetDiskFreeSpaceA }, { "kernel32", "GetDiskFreeSpaceExA", k_GetDiskFreeSpaceExA },
    { "kernel32", "GetModuleHandleA", k_GetModuleHandleA }, { "kernel32", "GetModuleFileNameA", k_GetModuleFileNameA },
    { "kernel32", "GetCurrentProcess", k_GetCurrentProcess }, { "kernel32", "GetCurrentThread", k_GetCurrentThread },
    { "kernel32", "GetCurrentThreadId", k_GetCurrentThreadId },
    { "kernel32", "LoadLibraryA", k_LoadLibraryA }, { "kernel32", "LoadLibrary", k_LoadLibraryA },
    { "kernel32", "FreeLibrary", ret1 }, { "kernel32", "GetProcAddress", k_GetProcAddress },
    { "kernel32", "GetVersion", k_GetVersion }, { "kernel32", "GetCommandLineA", k_GetCommandLineA },
    { "kernel32", "lstrcpynA", k_lstrcpynA }, { "kernel32", "lstrcpyn", k_lstrcpynA },
    { "kernel32", "GetCurrentDirectoryA", k_GetCurrentDirectoryA }, { "kernel32", "SetCurrentDirectoryA", k_SetCurrentDirectoryA },
    { "kernel32", "ExpandEnvironmentStringsA", k_ExpandEnvironmentStringsA },
    { "kernel32", "GetSystemTime", k_GetSystemTime }, { "kernel32", "GetLocalTime", k_GetLocalTime },
    { "kernel32", "GetTimeZoneInformation", k_GetTimeZoneInformation },
    { "kernel32", "GetUserDefaultLCID", k_GetUserDefaultLCID }, { "kernel32", "GetSystemDefaultLCID", k_GetUserDefaultLCID },
    { "kernel32", "GetThreadLocale", k_GetUserDefaultLCID }, { "kernel32", "GetUserDefaultLangID", k_GetUserDefaultLangID },
    { "kernel32", "GetSystemDefaultLangID", k_GetUserDefaultLangID },
    { "kernel32", "GetLocaleInfoA", k_GetLocaleInfoA }, { "kernel32", "GetLocaleInfo", k_GetLocaleInfoA },
    { "kernel32", "CreateDirectoryA", k_CreateDirectoryA }, { "kernel32", "RemoveDirectoryA", k_RemoveDirectoryA },
    { "kernel32", "DeleteFileA", k_DeleteFileA }, { "kernel32", "MoveFileA", k_MoveFileA }, { "kernel32", "CopyFileA", k_CopyFileA },
    { "kernel32", "GetFileAttributesA", k_GetFileAttributesA }, { "kernel32", "SetFileAttributesA", k_SetFileAttributesA },
    { "kernel32", "GetTempFileNameA", k_GetTempFileNameA },
    { "kernel32", "CreateFileA", k_CreateFileA }, { "kernel32", "CreateFile", k_CreateFileA },
    { "kernel32", "OpenFile", k_OpenFile },
    { "kernel32", "ReadFile", k_ReadFile }, { "kernel32", "WriteFile", k_WriteFile },
    { "kernel32", "FindFirstFileA", k_FindFirstFileA }, { "kernel32", "FindNextFileA", k_FindNextFileA },
    { "kernel32", "FindFirstFile", k_FindFirstFileA }, { "kernel32", "FindNextFile", k_FindNextFileA },
    { "kernel32", "FindClose", k_FindClose }, { "kernel32", "CreateToolhelp32Snapshot", k_CreateToolhelp32Snapshot },
    { "kernel32", "Process32First", k_Process32First }, { "kernel32", "Process32Next", k_Process32Next },
    { "kernel32", "GetPriorityClass", k_GetPriorityClass },
    { "kernel32", "IsBadCodePtr", ret0 }, { "kernel32", "LoadLibraryExA", k_LoadLibraryA },
    { "winmm", "timeGetDevCaps", w_timeGetDevCaps },
    { "kernel32", "CloseHandle", k_CloseHandle }, { "kernel32", "GetFileSize", k_GetFileSize },
    { "kernel32", "SetFilePointer", k_SetFilePointer }, { "kernel32", "SetEndOfFile", k_SetEndOfFile },
    { "kernel32", "FlushFileBuffers", k_FlushFileBuffers },
    { "kernel32", "GlobalAlloc", k_GlobalAlloc }, { "kernel32", "GlobalLock", k_GlobalLock },
    { "kernel32", "GlobalUnlock", ret1 }, { "kernel32", "GlobalFree", k_GlobalFree }, { "kernel32", "GlobalSize", k_GlobalSize },
    { "kernel32", "LocalAlloc", k_GlobalAlloc }, { "kernel32", "LocalFree", k_GlobalFree },
    { "kernel32", "WinExec", k_WinExec }, { "kernel32", "ExitProcess", k_ExitProcess },
    { "kernel32", "SetErrorMode", ret0 }, { "kernel32", "OutputDebugStringA", k_OutputDebugStringA },
    { "kernel32", "IsDebuggerPresent", ret0 }, { "kernel32", "FormatMessageA", ret0 },
    { "kernel32", "CreateMutexA", k_handle }, { "kernel32", "OpenMutexA", ret0 }, { "kernel32", "ReleaseMutex", ret1 },
    { "kernel32", "CreateEventA", k_handle }, { "kernel32", "SetEvent", ret1 }, { "kernel32", "ResetEvent", ret1 },
    { "kernel32", "WaitForSingleObject", k_WaitForSingleObject }, { "kernel32", "OpenProcess", k_OpenProcess },
    { "kernel32", "TerminateProcess", k_TerminateProcess }, { "kernel32", "GetExitCodeProcess", k_GetExitCodeProcess },
    { "kernel32", "CreateProcessA", k_CreateProcessA }, { "kernel32", "CreateProcess", k_CreateProcessA },
    { "kernel32", "SetConsoleTextAttribute", ret1 }, { "kernel32", "WaitForInputIdle", ret0 },
    { "kernel32", "SetPriorityClass", ret1 }, { "kernel32", "SetThreadPriority", ret1 },
    { "kernel32", "AllocConsole", ret1 }, { "kernel32", "FreeConsole", ret1 },
    { "advapi32", "RegOpenKeyA", a_RegOpenKeyA }, { "advapi32", "RegOpenKey", a_RegOpenKeyA },
    { "advapi32", "RegOpenKeyExA", a_RegOpenKeyExA }, { "advapi32", "RegOpenKeyEx", a_RegOpenKeyExA },
    { "advapi32", "RegCreateKeyA", a_RegCreateKeyA }, { "advapi32", "RegCreateKey", a_RegCreateKeyA },
    { "advapi32", "RegCreateKeyExA", a_RegCreateKeyExA }, { "advapi32", "RegCreateKeyEx", a_RegCreateKeyExA },
    { "advapi32", "RegCloseKey", a_RegCloseKey }, { "advapi32", "RegFlushKey", ret0 },
    { "advapi32", "RegSetValueExA", a_RegSetValueExA }, { "advapi32", "RegSetValueEx", a_RegSetValueExA },
    { "advapi32", "RegSetValueExString", a_RegSetValueExA }, { "advapi32", "RegSetValueExLong", a_RegSetValueExA },
    { "advapi32", "RegSetValueA", a_RegSetValueA },
    { "advapi32", "RegQueryValueExA", a_RegQueryValueExA }, { "advapi32", "RegQueryValueEx", a_RegQueryValueExA },
    { "advapi32", "RegQueryValueA", a_RegQueryValueA },
    { "advapi32", "RegDeleteValueA", a_RegDeleteValueA }, { "advapi32", "RegDeleteKeyA", a_RegDeleteKeyA },
    { "advapi32", "RegEnumKeyA", a_RegEnumKeyA }, { "advapi32", "RegEnumKeyExA", a_RegEnumKeyA },
    { "advapi32", "RegEnumValueA", a_RegEnumValueA }, { "advapi32", "RegQueryInfoKeyA", a_RegQueryInfoKeyA },
    { "shell32", "ShellExecuteA", s_ShellExecuteA }, { "shell32", "ShellExecute", s_ShellExecuteA },
    { "shell32", "SHGetSpecialFolderLocation", s_SHGetSpecialFolderLocation },
    { "shell32", "SHGetPathFromIDListA", s_SHGetPathFromIDListA }, { "shell32", "SHGetPathFromIDList", s_SHGetPathFromIDListA },
    { "shell32", "SHGetSpecialFolderPathA", s_SHGetSpecialFolderPathA }, { "shell32", "SHGetFolderPathA", s_SHGetFolderPathA },
    { "shfolder", "SHGetFolderPathA", s_SHGetFolderPathA },
    { "shell32", "SHFileOperationA", s_SHFileOperationA }, { "shell32", "SHFileOperation", s_SHFileOperationA },
    { "shell32", "ShellAboutA", s_ShellAboutA }, { "shell32", "SHAddToRecentDocs", ret0 },
    { "shell32", "SHAppBarMessage", ret0 }, { "shell32", "ExtractIconA", ret0 }, { "shell32", "DragAcceptFiles", ret0 },
    { "shell32", "SHEmptyRecycleBinA", ret0 }, { "shell32", "SHChangeNotify", ret0 },
    { "ole32", "CoTaskMemFree", ret0 }, { "ole32", "CoCreateGuid", o_CoCreateGuid },
    { "ole32", "CoInitialize", ret0 }, { "ole32", "CoUninitialize", ret0 },
    { "winmm", "timeBeginPeriod", ret0 }, { "winmm", "timeEndPeriod", ret0 }, { "winmm", "midiOutGetNumDevs", ret0 },
    { "winmm", "waveOutSetVolume", ret0 }, { "winmm", "waveOutGetVolume", w_waveOutGetVolume },
    { "wininet", "InternetGetConnectedState", i_InternetGetConnectedState }, { "wininet", "InternetAttemptConnect", ret0 },
    { "version", "GetFileVersionInfoSizeA", ret0 }, { "version", "GetFileVersionInfoA", ret0 },
    { "kernel32", "SleepEx", k_SleepEx }, { "kernel32", "GetStdHandle", k_GetStdHandle },
    { "kernel32", "GlobalAddAtomA", k_GlobalAddAtomA }, { "kernel32", "GlobalDeleteAtom", ret0 },
    { "kernel32", "GetModuleHandleW", k_GetModuleHandleW }, { "kernel32", "IsProcessorFeaturePresent", k_IsProcessorFeaturePresent },
    { "kernel32", "GetSystemPowerStatus", k_GetSystemPowerStatus }, { "kernel32", "IsSystemResumeAutomatic", ret0 },
    { "kernel32", "GetComputerNameExA", k_GetComputerNameExA }, { "kernel32", "RegisterServiceProcess", ret1 },
    { "kernel32", "SetSystemTime", ret0 }, { "kernel32", "SetLocalTime", ret0 }, { "kernel32", "Beep", ret1 },
    { "kernel32", "FlushInstructionCache", ret1 },
    { "kernel32", "SetConsoleTitleA", ret1 }, { "kernel32", "SetProcessWorkingSetSize", ret1 },
    { "shlwapi", "PathFileExistsA", sh_PathFileExistsA }, { "shlwapi", "PathFileExists", sh_PathFileExistsA },
    { "shlwapi", "PathIsDirectoryA", sh_PathIsDirectoryA }, { "shlwapi", "DllGetVersion", sh_DllGetVersion },
    { "shlwapi", "SHAutoComplete", ret0 }, { "comctl32", "DllGetVersion", sh_DllGetVersion },
    { "shell32", "DllGetVersion", sh_DllGetVersion }, { "shell32", "IsUserAnAdmin", ret0 },
    { "msvbvm60", "VarPtr", v_VarPtr }, { "msvbvm60", "__vbaCopyBytes", v_vbaCopyBytes },
    { "msvbvm60", "__vbaCopyBytesZero", v_vbaCopyBytesZero },
    { "msvbvm60", "GetMem4", v_GetMem4 }, { "msvbvm60", "GetMem2", v_GetMem2 }, { "msvbvm60", "GetMem1", v_GetMem1 },
    { "msvbvm60", "PutMem4", v_PutMem4 }, { "msvbvm60", "PutMem2", v_PutMem2 }, { "msvbvm60", "PutMem1", v_PutMem1 },
    { "winmm", "midiOutOpen", ret0 }, { "winmm", "midiOutClose", ret0 }, { "winmm", "midiOutShortMsg", ret0 },
    { "winmm", "midiOutReset", ret0 }, { "winmm", "midiOutSetVolume", ret0 },
    { "advapi32", "CloseEventLog", ret1 }, { "advapi32", "OpenEventLogA", k_handle }, { "advapi32", "ReportEventA", ret1 },
    { "advapi32", "LookupPrivilegeValueA", ret0 }, { "advapi32", "OpenProcessToken", ret0 },
    { "advapi32", "AdjustTokenPrivileges", ret0 },
    { 0, 0, 0 }
};
#endif
