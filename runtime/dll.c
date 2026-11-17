#define _GNU_SOURCE
/* dll.c - Declare: calls of DLL procedures.
 *
 * VB6 programs call the Windows API through Declare. On Linux there is no kernel32, so
 * the runtime provides the procedures that programs use most (same names, same results);
 * a Declare of any other library loads it as a shared object (dlopen) and calls the
 * procedure with integer / pointer arguments.
 *
 * VarPtr, StrPtr and ObjPtr return 32-bit numbers that stand for real addresses (a map),
 * so that they fit in a Long as in VB6; the procedures here accept them wherever the
 * Windows procedure takes an address (CopyMemory dest, ByVal StrPtr(s), ...).
 *
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"
#include <unistd.h>
#include <time.h>
#include <ctype.h>
#include <stdint.h>
#include <errno.h>
#include <sys/stat.h>
#include "vbos.h"
#ifndef _WIN32
#include <dlfcn.h>
#include <sys/time.h>
#include <sys/wait.h>
#include <dirent.h>
#include <strings.h>
#include <signal.h>
#include <fcntl.h>
#endif

/* ---------------------------------------------------------------- address map
 * One entry per (address, kind): kind 0 a variable, 1 a String variable (its vbs: VarPtr(s)), 2 the characters of a
 * String (StrPtr). The fake numbers grow, so a number is found by a binary search, and an address by a hash
 * (robot 2026-11-10: TriQuickSort of 170,000 Strings with CopyMemory ByVal VarPtr(...): the old linear search) */
typedef struct { char *real; vb_i32 size; vb_i32 kind; vb_i64 fake; } amap;
static amap *maps; static int nmaps, capmaps;
static int *mhash; static int hcap;             /* index + 1 of the newest entry of an address, 0 = empty */
static vb_i64 next_fake = 0x00100000;

static unsigned map_hkey(const void *p, int kind)
{
    uint64_t v = (uint64_t)(uintptr_t)p;
    v ^= v >> 29; v *= 0x9E3779B97F4A7C15ull;
    return (unsigned)(v >> 32) ^ (unsigned)kind * 0x85EBCA6Bu;
}
static void map_hput(int i)
{
    unsigned h = map_hkey(maps[i].real, maps[i].kind) & (unsigned)(hcap - 1);
    while (mhash[h]) {
        amap *m = &maps[mhash[h] - 1];
        if (m->real == maps[i].real && m->kind == maps[i].kind) break;
        h = (h + 1) & (unsigned)(hcap - 1);
    }
    mhash[h] = i + 1;
}
static int map_hget(const void *p, int kind)
{
    if (!hcap) return -1;
    unsigned h = map_hkey(p, kind) & (unsigned)(hcap - 1);
    while (mhash[h]) {
        amap *m = &maps[mhash[h] - 1];
        if (m->real == (const char *)p && m->kind == kind) return mhash[h] - 1;
        h = (h + 1) & (unsigned)(hcap - 1);
    }
    return -1;
}
static void map_rehash(int cap)
{
    free(mhash);
    mhash = (int *)calloc((size_t)cap, sizeof(int));
    hcap = mhash ? cap : 0;
    for (int i = 0; i < nmaps && hcap; i++) map_hput(i);
}

static vb_i32 map_addr_k(const void *p, vb_i32 size, int kind)
{
    if (!p) return 0;
    int j = map_hget(p, kind);
    if (j >= 0 && maps[j].size >= size) return (vb_i32)maps[j].fake;
    int grow = j;                               /* mapped before, smaller: a new entry takes its place */
    if (sizeof(void *) != 4 && next_fake + size + 32 > 0x7FF00000) {   /* wrap (very long programs): a new map */
        nmaps = 0; next_fake = 0x00100000; grow = -1;
        if (hcap) memset(mhash, 0, sizeof(int) * (size_t)hcap);
    }
    if (nmaps == capmaps) {
        int nc = capmaps ? capmaps * 2 : 64;
        amap *n = (amap *)realloc(maps, sizeof(amap) * nc);
        if (!n) { vb_raise(7); return 0; }
        maps = n; capmaps = nc;
    }
    amap *m = &maps[nmaps];
    m->real = (char *)p; m->size = size; m->kind = kind;
    if (sizeof(void *) == 4) {                  /* 32-bit: the real address fits in a Long, as in VB6 */
        m->fake = (vb_i64)(uintptr_t)p;
    } else {
        m->fake = next_fake;
        next_fake += ((size + 16 + 15) / 16) * 16;
    }
    nmaps++;
    if (grow >= 0 && hcap && (nmaps + 1) * 2 <= hcap) {
        for (unsigned h = map_hkey(p, kind) & (unsigned)(hcap - 1); mhash[h]; h = (h + 1) & (unsigned)(hcap - 1))
            if (mhash[h] == grow + 1) { mhash[h] = nmaps; return (vb_i32)m->fake; }
    }
    if ((nmaps + 1) * 2 > hcap) map_rehash(hcap ? hcap * 2 : 1024); else map_hput(nmaps - 1);
    return (vb_i32)m->fake;
}
static vb_i32 map_addr(const void *p, vb_i32 size) { return map_addr_k(p, size, 0); }

/* the entry whose range holds [addr, addr + size) */
static amap *map_find(vb_i64 addr, vb_i32 size)
{
    if (sizeof(void *) == 4) {                  /* 32-bit: fake = real, not in order */
        for (int k = 0; k < 3; k++) { int j = map_hget((void *)(uintptr_t)addr, k); if (j >= 0 && maps[j].size >= size) return &maps[j]; }
        for (int i = nmaps - 1; i >= 0; i--)
            if (addr >= maps[i].fake && addr + size <= maps[i].fake + maps[i].size) return &maps[i];
        return 0;
    }
    int lo = 0, hi = nmaps - 1, at = -1;
    while (lo <= hi) {
        int mid = (lo + hi) / 2;
        if (maps[mid].fake <= addr) { at = mid; lo = mid + 1; } else hi = mid - 1;
    }
    if (at >= 0 && addr + size <= maps[at].fake + maps[at].size) return &maps[at];
    return 0;
}

vb_i32 vb_varptr(const void *p, vb_i32 size) { return map_addr(p, size); }
/* VarPtr(a(0)): the element and those after it (GradientFill hDC, VarPtr(Vert(0)), 5, ...) */
vb_i32 vb_varptr_elem(const void *p, vb_i32 size, const vbarr *a)
{
    if (p && a && a->data && a->n > 0) {       /* the whole array: VarPtr(a(1)) - VarPtr(a(0)) = 4, as in VB6 */
        const char *e = a->data + (vb_i64)a->n * a->esz;
        if ((const char *)p >= a->data && (const char *)p < e && e - a->data < 0x40000000) {
            vb_i32 base = map_addr(a->data, (vb_i32)(e - a->data));
            return base ? base + (vb_i32)((const char *)p - a->data) : 0;
        }
    }
    return map_addr(p, size);
}
vb_i32 vb_varptr_str(const vbs *p) { return map_addr_k(p, (vb_i32)sizeof(vbs), 1); }
vb_i32 vb_strptr(vbs s) { return s.p ? map_addr_k(s.p, (s.len + 1) * 2, 2) : 0; }

void *vb_ptr_resolve(vb_i64 addr, vb_i32 size)
{
    amap *m = map_find(addr, size);
    if (m) return m->real + (addr - m->fake);
    vb_raise(5);
    return 0;
}

/* ---------------------------------------------------------------- arguments */
typedef struct { vb_dllarg *a; int n; char **ansi; } args;

static vb_i64 a_int(args *x, int i)
{
    if (i >= x->n) return 0;
    vb_dllarg *a = &x->a[i];
    return a->kind == VBD_DBL ? (vb_i64)a->d : a->i;
}

/* an address argument: a ByRef variable, a ByVal String (its ANSI copy) or a VarPtr number */
static void *a_ptr(args *x, int i, vb_i32 size)
{
    if (i >= x->n) return 0;
    vb_dllarg *a = &x->a[i];
    switch (a->kind) {
    case VBD_PTR: return a->p;
    case VBD_STR: return x->ansi[i];
    case VBD_INT: return a->i ? vb_ptr_resolve(a->i, size) : 0;
    }
    return 0;
}

static char *a_str(args *x, int i)
{
    if (i >= x->n) return 0;
    vb_dllarg *a = &x->a[i];
    if (a->kind == VBD_STR) return x->ansi[i];
    if (a->kind == VBD_INT && a->i) return (char *)vb_ptr_resolve(a->i, 1);
    if (a->kind == VBD_PTR) return (char *)a->p;
    return 0;
}

/* ---------------------------------------------------------------- Err.LastDllError */
vb_i32 vb_lastdllerr;                           /* GetLastError right after the last Declare call */
static vb_i32 emu_lasterr;                      /* the last error of the procedures of this runtime */
#ifdef _WIN32
__declspec(dllimport) unsigned long __stdcall GetLastError(void);
__declspec(dllimport) void __stdcall SetLastError(unsigned long);
#endif
static vb_i64 seterr(vb_i32 e, vb_i64 r) { emu_lasterr = e; return r; }
/* errno -> the Windows error code (path: to tell ERROR_FILE_NOT_FOUND from ERROR_PATH_NOT_FOUND) */
static __attribute__((unused)) vb_i32 w32err(int e, const char *path)
{
    switch (e) {
    case ENOENT: case ENOTDIR:
        if (path) {
            char d[1500]; snprintf(d, sizeof d, "%s", path);
            char *s = strrchr(d, '/');
            if (s && s != d) { *s = 0; struct stat st; if (stat(d, &st) != 0 || !S_ISDIR(st.st_mode)) return 3; }
        }
        return e == ENOTDIR ? 3 : 2;
    case EEXIST: return 183;                    /* ERROR_ALREADY_EXISTS */
    case EACCES: case EPERM: case EISDIR: return 5;
    case ENOTEMPTY: return 145;                 /* ERROR_DIR_NOT_EMPTY */
    case EBUSY: case ETXTBSY: return 32;        /* ERROR_SHARING_VIOLATION */
    case EMFILE: case ENFILE: return 4;
    case ENOSPC: return 112;
    case ENAMETOOLONG: return 206;
    case EROFS: return 19;
    case EXDEV: return 17;
    case EINVAL: return 87;
    case ENOMEM: return 8;
    }
    return 31;                                  /* ERROR_GEN_FAILURE */
}

/* ---------------------------------------------------------------- BSTR
 * VB6 gives a String to a Declare as an ANSI BSTR: the length in bytes in the 4 bytes before the text,
 * two zero bytes after it. On Windows the real ones of oleaut32 (the procedure may free or reallocate
 * them); here a buffer with the same layout, marked so that SysFreeString knows its own. */
#define BSTR_MAGIC 0x52545342u                  /* "BSTR" */
#ifdef _WIN32
typedef char *(__stdcall *bs_alloc_fn)(const char *, unsigned);
typedef void (__stdcall *bs_free_fn)(char *);
static bs_alloc_fn bs_alloc; static bs_free_fn bs_free;
static int bs_loaded;
static void bs_load(void)
{
    if (bs_loaded) return;
    bs_loaded = 1;
    void *h = vbos_dlopen("oleaut32.dll");
    if (h) { bs_alloc = (bs_alloc_fn)vbos_dlsym(h, "SysAllocStringByteLen"); bs_free = (bs_free_fn)vbos_dlsym(h, "SysFreeString"); }
}
#endif
static char *bstr_new(const void *b, vb_i32 n)
{
    if (n < 0) n = 0;
#ifdef _WIN32
    bs_load();
    if (bs_alloc && bs_free) {
        char *p = bs_alloc(0, (unsigned)n);
        if (p && n) { if (b) memcpy(p, b, (size_t)n); else memset(p, 0, (size_t)n); }
        return p;
    }
#endif
    /* room for at least 1024 bytes: a procedure told a bigger size than the string has (GetWindowsDirectory tmp, 255
     * with tmp = "": robot, malloc() corrupted) writes into it, not past it (VB6: into the slack of its heap) */
    char *m = (char *)malloc((size_t)(n < 1024 ? 1024 : n) + 10);
    if (!m) return 0;
    vb_u32 h[2] = { BSTR_MAGIC, (vb_u32)n };
    memcpy(m, h, 8);
    if (n) { if (b) memcpy(m + 8, b, (size_t)n); else memset(m + 8, 0, (size_t)n); }
    m[8 + n] = m[9 + n] = 0;
    return m + 8;
}
static vb_i32 bstr_len(const char *p) { vb_u32 n = 0; if (p) memcpy(&n, p - 4, 4); return (vb_i32)n; }
static int bstr_ours(const char *p)
{
#ifdef _WIN32
    bs_load();
    if (bs_alloc && bs_free) return p != 0;
#endif
    vb_u32 m = 0;
    if (p) memcpy(&m, p - 8, 4);
    return m == BSTR_MAGIC;
}
static void bstr_free(char *p)
{
    if (!p) return;
#ifdef _WIN32
    bs_load();
    if (bs_alloc && bs_free) { bs_free(p); return; }
#endif
    if (!bstr_ours(p)) return;
    /* freed a little later: a program may still read it through the address it was given, which
     * works in VB6 (OLE keeps freed BSTRs in a cache) */
    static char *ring[32]; static int k;
    memset(p - 8, 0, 4);
    if (ring[k]) free(ring[k] - 8);
    ring[k] = p; k = (k + 1) % 32;
}
/* an ANSI BSTR of a String (NULL for vbNullString) */
static char *bstr_ansi(const vbs *s)
{
    if (!s || !s->p) return 0;
    int n; char *a = vb_stoansi(*s, &n);
    return bstr_new(a, n);
}
/* the String of an ANSI BSTR: one character for each byte (VB6 converts the whole length) */
static vbs bstr_str(const char *p, vb_i32 n)
{
    if (!p) { vbs z = { 0, 0, 0 }; return z; }
    vbs r = vb_sfromansi(p, n);
    if (!r.p) { r = vb_stmp(0); }               /* "" (not vbNullString) */
    return r;
}

/* ---------------------------------------------------------------- the Windows procedures */
typedef vb_i64 (*apifn)(args *x);

static vb_i64 k_GetTickCount(args *x)
{
    (void)x;
#ifndef _WIN32
    struct timespec ts; clock_gettime(CLOCK_MONOTONIC, &ts);
    return (vb_i32)(vb_u32)((vb_i64)ts.tv_sec * 1000 + ts.tv_nsec / 1000000);
#else
    return 0;
#endif
}
static vb_i64 k_Sleep(args *x) { vb_files_flush(); usleep((useconds_t)(a_int(x, 0) * 1000)); return 0; }
static vb_i64 k_lstrlenA(args *x) { char *s = a_str(x, 0); return s ? (vb_i64)strlen(s) : 0; }
static vb_i64 k_lstrlenW(args *x)
{
    vb_dllarg *a = x->n ? &x->a[0] : 0;
    if (!a) return 0;
    if (a->kind == VBD_STR) return a->s ? a->s->len : 0;
    const vb_wchar *w = a->kind == VBD_PTR ? (const vb_wchar *)a->p : a->i ? (const vb_wchar *)vb_ptr_resolve(a->i, 2) : 0;
    if (!w) return 0;
    vb_i64 n = 0;
    while (w[n]) n++;
    return n;
}
static vb_i64 k_lstrcpyA(args *x)
{
    char *d = a_str(x, 0), *s = a_str(x, 1);
    if (!d || !s) return 0;
    memmove(d, s, strlen(s) + 1);
    return x->a[0].kind == VBD_INT ? x->a[0].i : 1;   /* a non-zero "pointer" */
}
static vb_i64 k_lstrcatA(args *x)
{
    char *d = a_str(x, 0), *s = a_str(x, 1);
    if (!d || !s) return 0;
    memmove(d + strlen(d), s, strlen(s) + 1);
    return 1;
}
static vb_i64 k_lstrcmpA(args *x) { char *a = a_str(x, 0), *b = a_str(x, 1); int r = strcoll(a ? a : "", b ? b : ""); return r < 0 ? -1 : r > 0; }
static vb_i64 k_lstrcmpiA(args *x) { char *a = a_str(x, 0), *b = a_str(x, 1); int r = strcasecmp(a ? a : "", b ? b : ""); return r < 0 ? -1 : r > 0; }

static vb_i64 k_GetEnvironmentVariableA(args *x)
{
    char *name = a_str(x, 0), *buf = a_str(x, 1);
    vb_i64 size = a_int(x, 2);
    const char *v = name ? getenv(name) : 0;
    if (!v) return seterr(203, 0);              /* ERROR_ENVVAR_NOT_FOUND */
    size_t n = strlen(v);
    if ((vb_i64)n + 1 > size || !buf) return (vb_i64)n + 1;      /* buffer too small: size needed */
    memcpy(buf, v, n + 1);
    return (vb_i64)n;
}
static vb_i64 k_SetEnvironmentVariableA(args *x)
{
    char *name = a_str(x, 0), *v = a_str(x, 1);
    if (!name) return 0;
#ifndef _WIN32
    if (v) setenv(name, v, 1); else unsetenv(name);
#endif
    return 1;
}
/* VB6 programs copy the 4-byte BSTR pointer of a String variable: CopyMemory ByVal VarPtr(s1), ByVal VarPtr(s2), 4
 * swaps / moves Strings without copying them (TriQuickSort, InsertionSort of PSC), and a 0 empties the variable
 * without freeing (the other variable still has the characters). A String of LinBASIC is more than a pointer: the
 * copy of the 4 bytes is the copy of the whole String (its buffer shared, as VB6 shares the BSTR); the 4 bytes of a
 * String read into a Long are its StrPtr, and a StrPtr written into a String makes it point at those characters */
static vbs *str_slot(args *x, int i)
{
    if (i >= x->n || x->a[i].kind != VBD_INT || !x->a[i].i) return 0;
    amap *m = map_find(x->a[i].i, 1);
    return m && m->kind == 1 && m->fake == x->a[i].i ? (vbs *)m->real : 0;
}
static void str_put_ptr(vbs *d, vb_i32 v)       /* the BSTR pointer of d = v (0, a StrPtr, a VarPtr of a String) */
{
    amap *m = v ? map_find(v, 1) : 0;
    if (!v) { d->p = 0; d->len = 0; d->flags = 0; }
    else if (m && m->kind == 2 && m->fake == v) { d->p = (vb_wchar *)m->real; d->len = m->size / 2 - 1; d->flags = 1; }
    else if (m && m->kind == 1 && m->fake == v) *d = *(vbs *)m->real;
    else vb_raise(5);
}
static int bstr_copy(args *x, int di, int si, vb_i32 n)   /* 1: done */
{
    if (n != 4) return 0;
    vbs *ds = str_slot(x, di), *ss = str_slot(x, si);
    if (ds && ss) { *ds = *ss; return 1; }
    if (ds) {
        void *s = a_ptr(x, si, 4);
        vb_i32 v = 0;
        if (s) memcpy(&v, s, 4);
        str_put_ptr(ds, v);
        return 1;
    }
    if (ss) {
        void *d = a_ptr(x, di, 4);
        vb_i32 v = vb_strptr(*ss);
        if (d) memcpy(d, &v, 4);
        return 1;
    }
    return 0;
}
/* Windows: CopyMemory / GetMem4 / PutMem4 / __vbaCopyBytes of the real DLLs, when they would move the 4-byte BSTR
 * pointer of a String variable, are done here (the real procedure would copy only a part of LinBASIC's String) */
static int name_is(const char *a, const char *b)           /* the same name, without case */
{
    for (; *a && *b; a++, b++) if (tolower((unsigned char)*a) != tolower((unsigned char)*b)) return 0;
    return *a == *b;
}
static __attribute__((unused)) int bstr_intercept(const char *entry, args *x)
{
    if (!entry) return 0;
    if ((name_is(entry, "RtlMoveMemory") || name_is(entry, "CopyMemory") || name_is(entry, "MoveMemory") ||
         name_is(entry, "RtlCopyMemory")) && x->n == 3)
        return bstr_copy(x, 0, 1, (vb_i32)a_int(x, 2));
    if (name_is(entry, "__vbaCopyBytes") && x->n == 3) return bstr_copy(x, 1, 2, (vb_i32)a_int(x, 0));
    if (name_is(entry, "GetMem4") && x->n == 2) return bstr_copy(x, 1, 0, 4);
    if (name_is(entry, "PutMem4") && x->n == 2) {
        vbs *ds = str_slot(x, 0);
        if (ds) { str_put_ptr(ds, (vb_i32)a_int(x, 1)); return 1; }
    }
    return 0;
}
static vb_i64 k_RtlMoveMemory(args *x)
{
    vb_i32 n = (vb_i32)a_int(x, 2);
    if (n <= 0) return 0;
    if (bstr_copy(x, 0, 1, n)) return 0;
    void *d = a_ptr(x, 0, n), *s = a_ptr(x, 1, n);
    if (d && s) memmove(d, s, (size_t)n);
    return 0;
}
static vb_i64 k_RtlZeroMemory(args *x)
{
    vb_i32 n = (vb_i32)a_int(x, 1);
    void *d = n > 0 ? a_ptr(x, 0, n) : 0;
    if (d) memset(d, 0, (size_t)n);
    return 0;
}
static vb_i64 k_GetCurrentProcessId(args *x) { (void)x; return (vb_i64)getpid(); }
static vb_i64 k_MulDiv(args *x)
{
    vb_i64 a = (vb_i32)a_int(x, 0), b = (vb_i32)a_int(x, 1), c = (vb_i32)a_int(x, 2);
    if (c == 0) return -1;
    vb_i64 p = a * b;
    int neg = (p < 0) != (c < 0);
    vb_i64 ap = p < 0 ? -p : p, ac = c < 0 ? -c : c;
    vb_i64 q = (ap * 2 + ac) / (ac * 2);       /* round half away from zero */
    if (q > 0x7FFFFFFF) return -1;
    return neg ? -q : q;
}
static vb_i64 k_GetLastError(args *x) { (void)x; return emu_lasterr; }
static vb_i64 k_SetLastError(args *x) { emu_lasterr = (vb_i32)a_int(x, 0); return 0; }

/* ---------------------------------------------------------------- oleaut32: BSTR */
static vb_i64 bstr_ret(char *p) { return p ? (vb_i64)vb_varptr(p, bstr_len(p) + 2) : 0; }
/* a BSTR argument: a ByVal String (its ANSI BSTR) or the number of one */
static char *a_bstr(args *x, int i)
{
    if (i >= x->n) return 0;
    vb_dllarg *a = &x->a[i];
    if (a->kind == VBD_STR) return x->ansi[i];
    if (a->kind == VBD_INT && a->i) return (char *)vb_ptr_resolve(a->i, 1);
    if (a->kind == VBD_PTR) return (char *)a->p;
    return 0;
}
static const vb_wchar *a_wstr(args *x, int i, vb_i32 n)
{
    if (i >= x->n) return 0;
    vb_dllarg *a = &x->a[i];
    if (a->kind == VBD_INT) return a->i ? (const vb_wchar *)vb_ptr_resolve(a->i, n * 2 > 2 ? n * 2 : 2) : 0;
    if (a->kind == VBD_PTR) return (const vb_wchar *)a->p;
    if (a->kind == VBD_STR) return (const vb_wchar *)x->ansi[i];
    return 0;
}
static vb_i32 wlen(const vb_wchar *w) { vb_i32 n = 0; if (w) while (w[n]) n++; return n; }
static vb_i64 o_SysAllocString(args *x)
{
    const vb_wchar *w = a_wstr(x, 0, 1);
    return w ? bstr_ret(bstr_new(w, wlen(w) * 2)) : 0;
}
static vb_i64 o_SysAllocStringLen(args *x)
{
    vb_i32 n = (vb_i32)a_int(x, 1);
    if (n < 0) return 0;
    return bstr_ret(bstr_new(a_wstr(x, 0, n), n * 2));
}
static vb_i64 o_SysAllocStringByteLen(args *x)
{
    vb_i32 n = (vb_i32)a_int(x, 1);
    if (n < 0) return 0;
    const char *s = a_str(x, 0);
    if (s && x->a[0].kind == VBD_STR && bstr_len(s) + 1 < n) {     /* (no reading past the copy) */
        char *p = bstr_new(0, n);
        if (p) memcpy(p, s, (size_t)bstr_len(s) + 1);
        return bstr_ret(p);
    }
    return bstr_ret(bstr_new(s, n));
}
static vb_i64 o_SysStringLen(args *x) { char *p = a_bstr(x, 0); return p ? bstr_len(p) / 2 : 0; }
static vb_i64 o_SysStringByteLen(args *x) { char *p = a_bstr(x, 0); return p ? bstr_len(p) : 0; }
static vb_i64 o_SysFreeString(args *x)
{
    if (x->n && x->a[0].kind != VBD_STR) { char *p = a_bstr(x, 0); if (p && bstr_ours(p)) bstr_free(p); }
    return 0;
}
/* a BSTR* (a ByRef String): the number of the BSTR in a Long (VarPtr-like) */
static void bstr_replace(vb_i32 *pp, char *n)
{
    char *o = *pp ? (char *)vb_ptr_resolve(*pp, 1) : 0;
    if (o && bstr_ours(o)) bstr_free(o);
    *pp = (vb_i32)bstr_ret(n);
}
static vb_i64 o_SysReAllocString(args *x)
{
    vb_i32 *pp = (vb_i32 *)a_ptr(x, 0, 4);
    if (!pp) return 0;
    const vb_wchar *w = a_wstr(x, 1, 1);
    bstr_replace(pp, w ? bstr_new(w, wlen(w) * 2) : 0);
    return 1;
}
static vb_i64 o_SysReAllocStringLen(args *x)
{
    vb_i32 *pp = (vb_i32 *)a_ptr(x, 0, 4);
    vb_i32 n = (vb_i32)a_int(x, 2);
    if (!pp || n < 0) return 0;
    bstr_replace(pp, bstr_new(a_wstr(x, 1, n), n * 2));
    return 1;
}
static vb_i64 k_Beep(args *x) { (void)x; vb_beep(); return 1; }
static vb_i64 k_QueryPerformanceFrequency(args *x)
{
    vb_cy *p = (vb_cy *)a_ptr(x, 0, 8);
    if (p) *p = 1000000;                        /* (Currency holds the 64-bit count / 10000) */
    return 1;
}
static vb_i64 k_QueryPerformanceCounter(args *x)
{
    vb_cy *p = (vb_cy *)a_ptr(x, 0, 8);
#ifndef _WIN32
    struct timespec ts; clock_gettime(CLOCK_MONOTONIC, &ts);
    if (p) *p = (vb_i64)ts.tv_sec * 1000000 + ts.tv_nsec / 1000;
#endif
    return 1;
}
static vb_i64 put_name(args *x, const char *v)
{
    char *buf = a_str(x, 0);
    vb_i32 *size = (vb_i32 *)a_ptr(x, 1, 4);
    size_t n = strlen(v);
    if (!buf || !size || (vb_i32)n + 1 > *size) { if (size) *size = (vb_i32)n + 1; return 0; }
    memcpy(buf, v, n + 1);
    *size = (vb_i32)n;
    return 1;
}
/* GetCurrencyFormatA(Locale, dwFlags, lpValue, lpFormat, lpCurrencyStr, cchCurrency): the number in lpValue
 * ("1234.5", "." as the separator) as FormatCurrency writes it in the program's locale */
static vb_i64 k_GetCurrencyFormatA(args *x)
{
    const char *v = a_str(x, 2);
    if (!v) return seterr(87, 0);
    vb_i32 m = vb_mark();
    vbs r = vb_formatnumber(vb_v_d(vb_val(vb_sfromutf8(v))), -1, -2, -2, -2, 1);
    int tn = 0;
    char *t = vb_stoansi(r, &tn);
    vb_i32 n = (vb_i32)tn, cch = (vb_i32)a_int(x, 5);
    vb_i64 ret = n + 1;
    if (cch > 0) {
        char *buf = a_str(x, 4);
        if (!buf || cch < n + 1) ret = seterr(122, 0);   /* ERROR_INSUFFICIENT_BUFFER */
        else memcpy(buf, t, (size_t)n + 1);
    }
    vb_release(m);
    return ret;
}
static vb_i64 k_GetComputerNameA(args *x)
{
    char h[256] = "localhost";
#ifndef _WIN32
    gethostname(h, sizeof h - 1);
#endif
    return put_name(x, h);
}
static vb_i64 a_GetUserNameA(args *x)
{
    const char *u = getenv("USER");
    return put_name(x, u ? u : "user");
}
static vb_i64 k_GetTempPathA(args *x)
{
    vb_i64 size = a_int(x, 0);
    char *buf = a_str(x, 1);
    const char *t = getenv("TMPDIR");
    char v[512];
    snprintf(v, sizeof v, "%s/", t && *t ? t : "/tmp");
    size_t n = strlen(v);
    if ((vb_i64)n + 1 > size || !buf) return (vb_i64)n + 1;
    memcpy(buf, v, n + 1);
    return (vb_i64)n;
}

#ifndef _WIN32
/* sndPlaySound / PlaySound (winmm): a .wav file, or one in memory (SND_MEMORY: LoadResData), played
 * by the sound player of the desktop (paplay of PulseAudio/PipeWire, pw-play, aplay of ALSA) */
enum { SND_ASYNC = 1, SND_NODEFAULT = 2, SND_MEMORY = 4, SND_LOOP = 8, SND_NOSTOP = 0x10, SND_PURGE = 0x40,
       SND_ALIAS = 0x10000, SND_FILENAME = 0x20000, SND_RESOURCE = 0x40004 };
static pid_t snd_pid;
static char snd_tmp[512];

static int snd_busy(void)
{
    if (snd_pid > 0 && waitpid(snd_pid, 0, WNOHANG) == snd_pid) snd_pid = 0;
    return snd_pid > 0;
}
static void snd_stop(void)
{
    if (snd_busy()) { kill(snd_pid, SIGTERM); waitpid(snd_pid, 0, 0); }
    snd_pid = 0;
}
static int snd_play(const char *file, int async)
{
    pid_t pid = fork();
    if (pid < 0) return 0;
    if (pid == 0) {
        int nul = open("/dev/null", O_RDWR);
        if (nul >= 0) { dup2(nul, 1); dup2(nul, 2); }
        execlp("paplay", "paplay", file, (char *)0);
        execlp("pw-play", "pw-play", file, (char *)0);
        execlp("aplay", "aplay", "-q", file, (char *)0);
        _exit(127);
    }
    snd_pid = pid;
    if (!async) { int st = 0; waitpid(pid, &st, 0); snd_pid = 0; return WIFEXITED(st) && WEXITSTATUS(st) == 0; }
    return 1;
}
static vb_i64 snd_common(args *x, int flags)
{
    if (flags & SND_NOSTOP && snd_busy()) return 0;
    snd_stop();
    if (flags & SND_PURGE) return 1;
    const char *file = 0;
    if (flags & SND_MEMORY) {
        const unsigned char *h = (const unsigned char *)a_ptr(x, 0, 12);
        if (!h || memcmp(h, "RIFF", 4)) return 0;
        vb_i32 n = (vb_i32)(h[4] | h[5] << 8 | h[6] << 16 | (vb_u32)h[7] << 24) + 8;
        const unsigned char *d = (const unsigned char *)a_ptr(x, 0, n);
        if (!d) return 0;
        const char *t = getenv("TMPDIR");
        snprintf(snd_tmp, sizeof snd_tmp, "%s/vbrt-sound-%d.wav", t && *t ? t : "/tmp", (int)getpid());
        FILE *f = fopen(snd_tmp, "wb");
        if (!f) return 0;
        fwrite(d, 1, (size_t)n, f);
        fclose(f);
        file = snd_tmp;
    } else {
        char *name = a_str(x, 0);
        if (!name || !*name) return 1;          /* sndPlaySound vbNullString: stop */
        static char path[1024];
        snprintf(path, sizeof path, "%s", name);
        for (char *q = path; *q; q++) if (*q == '\\') *q = '/';
        if (access(path, R_OK) != 0) {         /* a system sound (SystemStart ...): the beep */
            if (!(flags & SND_NODEFAULT)) vb_beep();
            return 0;
        }
        file = path;
    }
    return snd_play(file, flags & SND_ASYNC);
}
static vb_i64 w_sndPlaySoundA(args *x) { return snd_common(x, (int)a_int(x, 1)); }
static vb_i64 w_PlaySoundA(args *x) { return snd_common(x, (int)a_int(x, 2)); }
#endif

#ifndef _WIN32
/* ---------------------------------------------------------------- the Windows folder and .ini files
 * GetWindowsDirectory is ~/.local/share/linbasic-studio/windows (system32 under it): the place of
 * WIN.INI and of the .ini files that programs name without a folder, as on Windows */
#include <sys/sysinfo.h>
#include <sys/stat.h>
static const char *win_dir(int system)
{
    static char d[600], sd[640];
    if (!d[0]) {
        const char *x = getenv("XDG_DATA_HOME"), *h = getenv("HOME");
        if (x && *x) snprintf(d, sizeof d, "%s/linbasic-studio/windows", x);
        else snprintf(d, sizeof d, "%s/.local/share/linbasic-studio/windows", h ? h : "/tmp");
        snprintf(sd, sizeof sd, "%s/system32", d);
        for (char *p = sd + 1; *p; p++) if (*p == '/') { *p = 0; mkdir(sd, 0777); *p = '/'; }
        mkdir(sd, 0777);
    }
    return system ? sd : d;
}
static vb_i64 put_dir(args *x, const char *v)
{
    char *buf = a_str(x, 0);
    vb_i64 size = a_int(x, 1);
    size_t n = strlen(v);
    if ((vb_i64)n + 1 > size || !buf) return (vb_i64)n + 1;
    memcpy(buf, v, n + 1);
    return (vb_i64)n;
}
static vb_i64 k_GetWindowsDirectoryA(args *x) { return put_dir(x, win_dir(0)); }
static vb_i64 k_GetSystemDirectoryA(args *x) { return put_dir(x, win_dir(1)); }

/* the file of a profile call: a name alone is in the Windows folder; C:\x\y.ini is y.ini there too */
static void ini_path(const char *name, char *out, size_t cap)
{
    if (!name || !*name) name = "win.ini";
    char n[1024]; snprintf(n, sizeof n, "%s", name);
    for (char *p = n; *p; p++) if (*p == '\\') *p = '/';
    if (n[0] == '/') { snprintf(out, cap, "%s", n); return; }
    const char *b = (isalpha((unsigned char)n[0]) && n[1] == ':') ? (strrchr(n, '/') ? strrchr(n, '/') + 1 : n + 2) : n;
    if (b == n && strchr(n, '/')) { snprintf(out, cap, "%s", n); return; }   /* relative to the current folder */
    char low[1024]; snprintf(low, sizeof low, "%s", b);
    snprintf(out, cap, "%s/%s", win_dir(0), low);
    /* WIN.INI and win.ini are the same file */
    char dir[700]; snprintf(dir, sizeof dir, "%s", win_dir(0));
    DIR *dh = opendir(dir);
    if (dh) {
        struct dirent *e;
        while ((e = readdir(dh))) if (!strcasecmp(e->d_name, low)) { snprintf(out, cap, "%s/%s", dir, e->d_name); break; }
        closedir(dh);
    }
}
static vb_i64 profile_string(args *x, int first, const char *file)
{
    const char *sec = a_str(x, first), *key = a_str(x, first + 1), *def = a_str(x, first + 2);
    char *buf = a_str(x, first + 3);
    vb_i64 size = a_int(x, first + 4);
    if (!buf || size <= 0) return 0;
    char path[1200]; ini_path(file, path, sizeof path);
    int n = vb_ini_read(path, sec, key, buf, (int)size);
    if (n < 0) {                                /* the default, without its blanks at the end */
        snprintf(buf, (size_t)size, "%s", def ? def : "");
        n = (int)strlen(buf);
        while (n && buf[n - 1] == ' ') buf[--n] = 0;
    }
    return n;
}
static vb_i64 k_GetPrivateProfileStringA(args *x) { return profile_string(x, 0, a_str(x, 5)); }
static vb_i64 k_GetProfileStringA(args *x) { return profile_string(x, 0, "win.ini"); }
static vb_i64 profile_int(args *x, const char *file)
{
    char path[1200], v[64]; ini_path(file, path, sizeof path);
    const char *sec = a_str(x, 0), *key = a_str(x, 1);
    if (!sec || !key || vb_ini_read(path, sec, key, v, sizeof v) < 0) return a_int(x, 2);
    return (vb_i64)(vb_i32)strtol(v, 0, 10);
}
static vb_i64 k_GetPrivateProfileIntA(args *x) { return profile_int(x, a_str(x, 3)); }
static vb_i64 k_GetProfileIntA(args *x) { return profile_int(x, "win.ini"); }
static vb_i64 profile_write(args *x, const char *file)
{
    char path[1200]; ini_path(file, path, sizeof path);
    const char *sec = a_str(x, 0);
    if (!sec) return 0;
    return vb_ini_write(path, sec, a_str(x, 1), a_str(x, 2));
}
static vb_i64 k_WritePrivateProfileStringA(args *x) { return profile_write(x, a_str(x, 3)); }
static vb_i64 k_WriteProfileStringA(args *x) { return profile_write(x, "win.ini"); }

/* GetSystemInfo, GlobalMemoryStatus, GetVersionEx: the structures of Windows, filled from Linux */
static void put32(unsigned char *p, vb_u32 v) { memcpy(p, &v, 4); }
static vb_i64 k_GetSystemInfo(args *x)
{
    unsigned char *p = (unsigned char *)a_ptr(x, 0, 36);
    if (!p) return 0;
    memset(p, 0, 36);
    long cpus = sysconf(_SC_NPROCESSORS_ONLN);
    put32(p + 0, 0);                            /* PROCESSOR_ARCHITECTURE_INTEL */
    put32(p + 4, (vb_u32)sysconf(_SC_PAGESIZE));
    put32(p + 8, 0x10000); put32(p + 12, 0x7FFEFFFF);
    put32(p + 16, cpus >= 32 ? 0xFFFFFFFFu : (1u << cpus) - 1);
    put32(p + 20, (vb_u32)cpus);
    put32(p + 24, 586);                         /* PROCESSOR_INTEL_PENTIUM */
    put32(p + 28, 0x10000);
    p[32] = 6; p[34] = 0;                       /* wProcessorLevel 6 */
    return 0;
}
static vb_i64 k_GlobalMemoryStatus(args *x)
{
    unsigned char *p = (unsigned char *)a_ptr(x, 0, 32);
    struct sysinfo si;
    if (!p || sysinfo(&si) != 0) return 0;
    double u = si.mem_unit;
    double total = si.totalram * u, avail = (si.freeram + si.bufferram) * u;
    double swt = si.totalswap * u, swf = si.freeswap * u;
    vb_u32 cap = 0x7FFFFFFFu;                   /* a 32-bit program (not large-address aware) sees 2 GB at most */
    put32(p, 32);
    put32(p + 4, (vb_u32)(total > 0 ? 100 - 100 * avail / total : 0));
    put32(p + 8, total > cap ? cap : (vb_u32)total); put32(p + 12, avail > cap ? cap : (vb_u32)avail);
    put32(p + 16, total + swt > cap ? cap : (vb_u32)(total + swt)); put32(p + 20, avail + swf > cap ? cap : (vb_u32)(avail + swf));
    put32(p + 24, 0x7FFE0000u); put32(p + 28, 0x7FF00000u);
    return 0;
}
static vb_i64 k_GetVersionExA(args *x)
{
    unsigned char *p = (unsigned char *)a_ptr(x, 0, 148);
    if (!p) return 0;
    put32(p + 4, 6); put32(p + 8, 1); put32(p + 12, 7601); put32(p + 16, 2);   /* Windows NT 6.1 (7) */
    memset(p + 20, 0, 128);
    strcpy((char *)p + 20, "Service Pack 1");
    return 1;
}
static vb_i64 w_waveOutGetNumDevs(args *x) { (void)x; return 1; }

/* MCI (mci.c): the commands and the strings are ANSI in the program, UTF-8 in the MCI */
static char *ansi_utf8(const char *a)
{
    vbs s = vb_sfromansi(a ? a : "", a ? (int)strlen(a) : 0);
    return strdup(vb_stoutf8(s));
}
static vb_i64 put_ansi(const char *u, char *buf, vb_i64 n)
{
    int an = 0;
    char *a = vb_stoansi(vb_sfromutf8(u), &an);
    if (!buf || n <= 0) return 0;
    if (an >= n) { memcpy(buf, a, (size_t)n - 1); buf[n - 1] = 0; return 268; }   /* MCIERR_PARAM_OVERFLOW */
    memcpy(buf, a, (size_t)an + 1);
    return 0;
}
static vb_i64 m_mciSendStringA(args *x)
{
    vb_i32 mk = vb_mark();
    char *cmd = ansi_utf8(a_str(x, 0)), r[2048];
    vb_i64 e = vb_mci_send_hwnd(cmd, r, sizeof r, 0, 0, a_int(x, 3));
    free(cmd);
    vb_i64 o = put_ansi(r, a_str(x, 1), a_int(x, 2));
    vb_release(mk);
    return e ? e : o;
}
static vb_i64 m_mciGetErrorStringA(args *x)
{
    vb_i32 mk = vb_mark();
    char t[1024];
    int ok = vb_mci_error_text((vb_i32)a_int(x, 0), t, sizeof t);
    if (ok) put_ansi(t, a_str(x, 1), a_int(x, 2));
    else { char *b = a_str(x, 1); if (b && a_int(x, 2) > 0) b[0] = 0; }
    vb_release(mk);
    return ok;
}
static vb_i64 m_mciGetDeviceIDA(args *x)
{
    vb_i32 mk = vb_mark();
    char *n = ansi_utf8(a_str(x, 0));
    vb_i64 r = vb_mci_device_id(n);
    free(n);
    vb_release(mk);
    return r;
}
static vb_i64 m_mciExecute(args *x)
{
    vb_i32 mk = vb_mark();
    char *cmd = ansi_utf8(a_str(x, 0)), r[2048];
    vb_i32 e = vb_mci_send(cmd, r, sizeof r, 0, 0);
    free(cmd);
    if (e) { char t[1024]; vb_mci_error_text(e, t, sizeof t); vb_msgbox(vb_sfromutf8(t), 16, vb_sfromutf8("MCI")); }
    vb_release(mk);
    return e == 0;
}
#endif

static const struct { const char *lib, *name; apifn fn; } apis[] __attribute__((unused)) = {
#ifndef _WIN32
    { "winmm", "sndPlaySoundA", w_sndPlaySoundA }, { "winmm", "sndPlaySound", w_sndPlaySoundA },
    { "winmm", "PlaySoundA", w_PlaySoundA }, { "winmm", "PlaySound", w_PlaySoundA },
    { "winmm", "waveOutGetNumDevs", w_waveOutGetNumDevs },
    { "winmm", "mciSendStringA", m_mciSendStringA }, { "winmm", "mciSendString", m_mciSendStringA },
    { "winmm", "mciGetErrorStringA", m_mciGetErrorStringA }, { "winmm", "mciGetErrorString", m_mciGetErrorStringA },
    { "winmm", "mciGetDeviceIDA", m_mciGetDeviceIDA }, { "winmm", "mciGetDeviceID", m_mciGetDeviceIDA },
    { "winmm", "mciExecute", m_mciExecute },
    { "kernel32", "GetWindowsDirectoryA", k_GetWindowsDirectoryA }, { "kernel32", "GetSystemDirectoryA", k_GetSystemDirectoryA },
    { "kernel32", "GetPrivateProfileStringA", k_GetPrivateProfileStringA }, { "kernel32", "GetProfileStringA", k_GetProfileStringA },
    { "kernel32", "GetPrivateProfileIntA", k_GetPrivateProfileIntA }, { "kernel32", "GetProfileIntA", k_GetProfileIntA },
    { "kernel32", "WritePrivateProfileStringA", k_WritePrivateProfileStringA },
    { "kernel32", "WriteProfileStringA", k_WriteProfileStringA },
    { "kernel32", "GetSystemInfo", k_GetSystemInfo }, { "kernel32", "GlobalMemoryStatus", k_GlobalMemoryStatus },
    { "kernel32", "GetVersionExA", k_GetVersionExA }, { "kernel32", "GetCurrencyFormatA", k_GetCurrencyFormatA },
#endif
    { "kernel32", "GetTickCount", k_GetTickCount }, { "kernel32", "Sleep", k_Sleep },
    { "kernel32", "lstrlenA", k_lstrlenA }, { "kernel32", "lstrlen", k_lstrlenA }, { "kernel32", "lstrlenW", k_lstrlenW },
    { "kernel32", "lstrcpyA", k_lstrcpyA }, { "kernel32", "lstrcpy", k_lstrcpyA },
    { "kernel32", "lstrcatA", k_lstrcatA }, { "kernel32", "lstrcat", k_lstrcatA },
    { "kernel32", "lstrcmpA", k_lstrcmpA }, { "kernel32", "lstrcmp", k_lstrcmpA },
    { "kernel32", "lstrcmpiA", k_lstrcmpiA }, { "kernel32", "lstrcmpi", k_lstrcmpiA },
    { "kernel32", "GetEnvironmentVariableA", k_GetEnvironmentVariableA },
    { "kernel32", "SetEnvironmentVariableA", k_SetEnvironmentVariableA },
    { "kernel32", "RtlMoveMemory", k_RtlMoveMemory }, { "kernel32", "RtlZeroMemory", k_RtlZeroMemory },
    { "kernel32", "GetCurrentProcessId", k_GetCurrentProcessId }, { "kernel32", "MulDiv", k_MulDiv },
    { "kernel32", "GetLastError", k_GetLastError }, { "kernel32", "SetLastError", k_SetLastError },
    { "kernel32", "SetLastErrorEx", k_SetLastError }, { "kernel32", "Beep", k_Beep },
    { "oleaut32", "SysAllocString", o_SysAllocString }, { "oleaut32", "SysAllocStringLen", o_SysAllocStringLen },
    { "oleaut32", "SysAllocStringByteLen", o_SysAllocStringByteLen }, { "oleaut32", "SysStringLen", o_SysStringLen },
    { "oleaut32", "SysStringByteLen", o_SysStringByteLen }, { "oleaut32", "SysFreeString", o_SysFreeString },
    { "oleaut32", "SysReAllocString", o_SysReAllocString }, { "oleaut32", "SysReAllocStringLen", o_SysReAllocStringLen },
    { "kernel32", "QueryPerformanceCounter", k_QueryPerformanceCounter },
    { "kernel32", "QueryPerformanceFrequency", k_QueryPerformanceFrequency },
    { "kernel32", "GetComputerNameA", k_GetComputerNameA }, { "kernel32", "GetTempPathA", k_GetTempPathA },
    { "advapi32", "GetUserNameA", a_GetUserNameA },
    { "winmm", "timeGetTime", k_GetTickCount },
    { 0, 0, 0 }
};
static const char *win_libs[] __attribute__((unused)) = { "kernel32", "user32", "gdi32", "advapi32", "shell32", "winmm", "ole32",
                                  "oleaut32", "comdlg32", "comctl32", "shlwapi", "version", "msvbvm60", "msimg32",
                                  "wininet", "shfolder", "olepro32", "mpr", "ntdll",
                                  /* GDI+: Linux may have Mono's libgdiplus.so, which takes the handles of Windows
                                   * and crashes (robot: GdipCreateBitmapFromHBITMAP of a Picture.Handle) */
                                  "gdiplus", 0 };

vb_apiext (*vb_api_hook)(const char *lib, const char *name);

/* procedures that call a procedure of the program by its AddressOf number (on every system; on
 * Windows a real procedure address goes on to the real one) */
static vb_i64 u_CallWindowProc(vb_i32 argc, vb_dllarg *a, const char *real)
{
    vb_i64 h = argc > 0 ? (a[0].kind == VBD_DBL ? (vb_i64)a[0].d : a[0].i) : 0;
    vb_i32 v[4] = { 0, 0, 0, 0 };
    for (int i = 0; i < 4 && i + 1 < argc; i++) v[i] = (vb_i32)(a[i + 1].kind == VBD_DBL ? (vb_i64)a[i + 1].d : a[i + 1].i);
    if (vb_callback_fn(h)) return vb_callback(h, 4, v);
#ifdef _WIN32
    void *u = vbos_dlopen("user32.dll");
    typedef intptr_t (__stdcall *cwp)(void *, void *, unsigned, uintptr_t, intptr_t);
    cwp f = u ? (cwp)vbos_dlsym(u, real) : 0;
    if (f) return f((void *)(intptr_t)h, (void *)(intptr_t)v[0], (unsigned)v[1], (uintptr_t)(unsigned)v[2], (intptr_t)v[3]);
#else
    (void)real;
#endif
    return 0;
}
static vb_i64 u_CallWindowProcA(vb_i32 argc, vb_dllarg *a) { return u_CallWindowProc(argc, a, "CallWindowProcA"); }
static vb_i64 u_CallWindowProcW(vb_i32 argc, vb_dllarg *a) { return u_CallWindowProc(argc, a, "CallWindowProcW"); }
static const struct { const char *lib, *name; vb_apiext fn; } cb_apis[] = {
    { "user32", "CallWindowProcA", u_CallWindowProcA }, { "user32", "CallWindowProcW", u_CallWindowProcW },
    { 0, 0, 0 }
};

void *vb_dllarg_ptr(vb_dllarg *a, vb_i32 size)
{
    switch (a->kind) {
    case VBD_PTR: return a->p;
    case VBD_INT: return a->i ? vb_ptr_resolve(a->i, size) : 0;
    }
    return 0;
}

static void lib_key(const char *lib, char *out, size_t n)
{
    const char *b = lib;
    for (const char *p = lib; *p; p++) if (*p == '\\' || *p == '/') b = p + 1;
    size_t k = 0;
    for (; *b && k < n - 1; b++) out[k++] = (char)tolower((unsigned char)*b);
    out[k] = 0;
    if (k > 4 && !strcmp(out + k - 4, ".dll")) out[k - 4] = 0;
}

static void entry_error(vb_declare *d, int number)
{
    char msg[512];
    if (number == 453) snprintf(msg, sizeof msg, "Can't find DLL entry point %s in %s", d->entry, d->lib);
    else snprintf(msg, sizeof msg, "File not found: %s", d->lib);
    if (number != 453) vb_lastdllerr = 0;       /* (VB6: the DLL not found leaves 0) */
    vb_i32 m = vb_mark();
    vb_raise_full(number, vb_stmp(0), vb_sfromutf8(msg));
    vb_release(m);
}

#include "winapi.inc"

/* find the procedure once: state 1 = runtime procedure, 2 = shared object, -1 = error */
static int resolve(vb_declare *d)
{
    if (d->state) return d->state;
    char key[256];
    lib_key(d->lib, key, sizeof key);
    for (int i = 0; cb_apis[i].lib; i++)
        if (!strcmp(cb_apis[i].lib, key) && !strcmp(cb_apis[i].name, d->entry)) { d->fn = (void *)cb_apis[i].fn; return d->state = 3; }
#ifndef _WIN32                                  /* on Windows the real procedures of the DLLs */
    for (int pass = 0; pass < 2; pass++) {      /* the name, then with "A" as VB6 tries (MessageBox -> MessageBoxA) */
        char nm[300];
        snprintf(nm, sizeof nm, pass ? "%sA" : "%s", d->entry);
        for (int i = 0; apis[i].lib; i++)
            if (!strcmp(apis[i].lib, key) && !strcmp(apis[i].name, nm)) { d->fn = (void *)apis[i].fn; return d->state = 1; }
        for (int i = 0; apis2[i].lib; i++)
            if (!strcmp(apis2[i].lib, key) && !strcmp(apis2[i].name, nm)) { d->fn = (void *)apis2[i].fn; return d->state = 1; }
        if (vb_api_hook) { vb_apiext f = vb_api_hook(key, nm); if (f) { d->fn = (void *)f; return d->state = 3; } }
    }
#endif
    if (vb_api_hook) { vb_apiext f = vb_api_hook(key, d->entry); if (f) { d->fn = (void *)f; return d->state = 3; } }
#ifndef _WIN32
    for (int i = 0; win_libs[i]; i++)
        if (!strcmp(win_libs[i], key)) { entry_error(d, 453); return 0; }
    char path[600];
    const char *tries[] = { "%s", "%s.so", "lib%s.so", 0 };
    void *h = 0;
    for (int i = 0; tries[i] && !h; i++) { snprintf(path, sizeof path, tries[i], i ? key : d->lib); h = dlopen(path, RTLD_NOW); }
    if (!h) { entry_error(d, 53); return 0; }
    d->fn = dlsym(h, d->entry);
    if (!d->fn) { entry_error(d, 453); return 0; }
    return d->state = 2;
#else
    /* Windows: the DLL itself (kernel32, user32, the user's own DLLs) */
    void *h = vbos_dlopen(d->lib);
    if (!h) { char k2[300]; snprintf(k2, sizeof k2, "%s.dll", key); h = vbos_dlopen(k2); }
    if (!h) { entry_error(d, 53); return 0; }
    d->fn = vbos_dlsym(h, d->entry);
    if (!d->fn) {                               /* VB6 tries the ANSI name too: MessageBox -> MessageBoxA */
        char a[300]; snprintf(a, sizeof a, "%sA", d->entry);
        d->fn = vbos_dlsym(h, a);
    }
    if (!d->fn) { entry_error(d, 453); return 0; }
    return d->state = 2;
#endif
}

#ifdef _WIN32
/* an argument that is a VarPtr/StrPtr number: the real address for the DLL (64-bit Windows) */
static __attribute__((unused)) vb_i64 real_int(vb_i64 v)
{
    if (sizeof(void *) == 4 || v < 0x00100000 || v > 0x7FF00000) return v;
    amap *m = map_find(v, 1);
    if (m) return (vb_i64)(intptr_t)(m->real + (v - m->fake));
    return v;
}
#if defined(__i386__)
/* Declare is __stdcall: the procedure takes its arguments off the stack, so it is called with
 * exactly as many 32-bit words as it expects (a ByVal Double is two) */
typedef unsigned W;
static vb_i64 call_std(void *fn, W *w, int n)
{
#define F(...) ((vb_i64 (__stdcall *)(__VA_ARGS__))fn)
    switch (n) {
    case 0: return F(void)();
    case 1: return F(W)(w[0]);
    case 2: return F(W, W)(w[0], w[1]);
    case 3: return F(W, W, W)(w[0], w[1], w[2]);
    case 4: return F(W, W, W, W)(w[0], w[1], w[2], w[3]);
    case 5: return F(W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4]);
    case 6: return F(W, W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4], w[5]);
    case 7: return F(W, W, W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4], w[5], w[6]);
    case 8: return F(W, W, W, W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4], w[5], w[6], w[7]);
    case 9: return F(W, W, W, W, W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4], w[5], w[6], w[7], w[8]);
    case 10: return F(W, W, W, W, W, W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4], w[5], w[6], w[7], w[8], w[9]);
    case 11: return F(W, W, W, W, W, W, W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4], w[5], w[6], w[7], w[8], w[9], w[10]);
    case 12: return F(W, W, W, W, W, W, W, W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4], w[5], w[6], w[7], w[8], w[9], w[10], w[11]);
    case 13: return F(W, W, W, W, W, W, W, W, W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4], w[5], w[6], w[7], w[8], w[9], w[10], w[11], w[12]);
    case 14: return F(W, W, W, W, W, W, W, W, W, W, W, W, W, W)(w[0], w[1], w[2], w[3], w[4], w[5], w[6], w[7], w[8], w[9], w[10], w[11], w[12], w[13]);
    }
#undef F
    vb_raise(450);
    return 0;
}
#endif
#endif

/* ---------------------------------------------------------------- Type arguments
 * A user-defined type passed to a Declare (ByRef, or As Any) goes as VB6 passes it: in the layout
 * of 32-bit Windows (each member aligned to its size, at most 4), its String members as pointers to
 * ANSI copies, its String * n members as n ANSI bytes; after the call the values come back. */
typedef struct { char *buf; vbs *str; vb_i32 off; } mstr;
typedef struct { mstr *s; int n, cap; } mstrs;

static vb_i32 kind_size(vb_i32 kind, vb_i32 fixlen, const vb_udt *u, vb_i32 *align);
static vb_i32 udt_layout(const vb_udt *u, vb_i32 *align)
{
    vb_i32 off = 0, al = 1;
    for (int i = 0; i < u->nfields; i++) {
        const vb_ufield *f = &u->fields[i];
        vb_i32 a, sz = f->arr == 2 ? (a = 4, 4) : kind_size(f->kind, f->fixlen, f->udt, &a);
        if (f->arr == 1) {
            /* the count of a fixed array is only known from the array itself: see udt_field_size */
        }
        if (a > al) al = a;
        off = (off + a - 1) / a * a + sz;
    }
    if (align) *align = al;
    return (off + al - 1) / al * al;
}
static vb_i32 kind_size(vb_i32 kind, vb_i32 fixlen, const vb_udt *u, vb_i32 *align)
{
    switch (kind) {
    case VBK_I2: case VBK_BOOL: *align = 2; return 2;
    case VBK_I4: case VBK_R4: case VBK_STR: case VBK_OBJ: *align = 4; return 4;
    case VBK_R8: case VBK_CY: case VBK_DATE: *align = 4; return 8;
    case VBK_U8: *align = 1; return 1;
    case VBK_FSTR: *align = 1; return fixlen;
    case VBK_VAR: *align = 4; return 16;
    case VBK_UDT: return u ? udt_layout(u, align) : (*align = 4, 0);
    }
    *align = 4;
    return 4;
}
/* the size of one field in the VB6 layout, with the elements of a fixed array (from the array) */
static vb_i32 field_size(const vb_ufield *f, const char *src, vb_i32 *align, vb_i32 *count)
{
    *count = 1;
    if (f->arr == 2) { *align = 4; return 4; }
    vb_i32 one = kind_size(f->kind, f->fixlen, f->udt, align);
    if (f->arr == 1) {
        const vbarr *a = (const vbarr *)(src + f->offset);
        *count = a->nd ? a->n : 0;
    }
    return one * *count;
}
static vb_i32 udt_size(const vb_udt *u, const char *src, vb_i32 *align)
{
    vb_i32 off = 0, al = 1;
    for (int i = 0; i < u->nfields; i++) {
        vb_i32 a, cnt, sz = field_size(&u->fields[i], src, &a, &cnt);
        if (u->fields[i].kind == VBK_UDT && u->fields[i].arr == 0) sz = udt_size(u->fields[i].udt, src + u->fields[i].offset, &a);
        if (a > al) al = a;
        off = (off + a - 1) / a * a + sz;
    }
    if (align) *align = al;
    return (off + al - 1) / al * al;
}
static void mstr_add(mstrs *m, char *buf, vbs *str, vb_i32 off)
{
    if (m->n == m->cap) { m->cap = m->cap ? m->cap * 2 : 8; m->s = (mstr *)realloc(m->s, sizeof(mstr) * (size_t)m->cap); }
    m->s[m->n].buf = buf; m->s[m->n].str = str; m->s[m->n].off = off; m->n++;
}
static void put_value(vb_i32 kind, vb_i32 fixlen, const vb_udt *u, const char *src, char *dst, vb_i32 doff, mstrs *m);
static void udt_put(const vb_udt *u, const char *src, char *dst, vb_i32 base, mstrs *m)
{
    vb_i32 off = 0;
    for (int i = 0; i < u->nfields; i++) {
        const vb_ufield *f = &u->fields[i];
        vb_i32 a, cnt, sz = field_size(f, src, &a, &cnt);
        if (f->kind == VBK_UDT && f->arr == 0) sz = udt_size(f->udt, src + f->offset, &a);
        off = (off + a - 1) / a * a;
        if (f->arr == 1 && cnt) {
            const vbarr *ar = (const vbarr *)(src + f->offset);
            vb_i32 one = sz / cnt;
            for (vb_i32 k = 0; k < cnt; k++)
                put_value(f->kind, f->fixlen, f->udt, ar->data + (size_t)k * ar->esz, dst, base + off + k * one, m);
        } else if (f->arr == 0)
            put_value(f->kind, f->fixlen, f->udt, src + f->offset, dst, base + off, m);
        off += sz;
    }
}
static void put_value(vb_i32 kind, vb_i32 fixlen, const vb_udt *u, const char *src, char *dst, vb_i32 doff, mstrs *m)
{
    char *d = dst + doff;
    switch (kind) {
    case VBK_I2: case VBK_BOOL: memcpy(d, src, 2); break;
    case VBK_I4: case VBK_R4: memcpy(d, src, 4); break;
    case VBK_R8: case VBK_CY: case VBK_DATE: memcpy(d, src, 8); break;
    case VBK_U8: *d = *src; break;
    case VBK_OBJ: { vb_i32 z = 0; memcpy(d, &z, 4); break; }
    case VBK_FSTR: {
        const vbs *v = (const vbs *)src;
        for (vb_i32 k = 0; k < fixlen; k++) {
            int c = k < v->len && v->p ? vb_w2ansi(v->p[k]) : 0;
            d[k] = (char)(c < 0 ? '?' : c);
        }
        break;
    }
    case VBK_STR: {
        const vbs *v = (const vbs *)src;
        vb_i32 addr = 0;
        if (!m) break;                          /* (back to a callback's caller: its pointer stays)*/
        if (v->p) {                             /* vbNullString stays a null pointer */
            char *b = bstr_ansi(v);             /* an ANSI BSTR */
            if (!b) break;
            addr = map_addr(b, v->len + 1);
            mstr_add(m, b, (vbs *)src, doff);
        }
        memcpy(d, &addr, 4);
        break;
    }
    case VBK_UDT: if (u) udt_put(u, src, dst, doff, m); break;
    default: break;
    }
}
static void get_value(vb_i32 kind, vb_i32 fixlen, const vb_udt *u, char *dstv, const char *buf, vb_i32 boff);
static void udt_get(const vb_udt *u, char *dstv, const char *buf, vb_i32 base)
{
    vb_i32 off = 0;
    for (int i = 0; i < u->nfields; i++) {
        const vb_ufield *f = &u->fields[i];
        vb_i32 a, cnt, sz = field_size(f, dstv, &a, &cnt);
        if (f->kind == VBK_UDT && f->arr == 0) sz = udt_size(f->udt, dstv + f->offset, &a);
        off = (off + a - 1) / a * a;
        if (f->arr == 1 && cnt) {
            vbarr *ar = (vbarr *)(dstv + f->offset);
            vb_i32 one = sz / cnt;
            for (vb_i32 k = 0; k < cnt; k++)
                get_value(f->kind, f->fixlen, f->udt, ar->data + (size_t)k * ar->esz, buf, base + off + k * one);
        } else if (f->arr == 0)
            get_value(f->kind, f->fixlen, f->udt, dstv + f->offset, buf, base + off);
        off += sz;
    }
}
static void get_value(vb_i32 kind, vb_i32 fixlen, const vb_udt *u, char *dstv, const char *buf, vb_i32 boff)
{
    const char *s = buf + boff;
    switch (kind) {
    case VBK_I2: case VBK_BOOL: memcpy(dstv, s, 2); break;
    case VBK_I4: case VBK_R4: memcpy(dstv, s, 4); break;
    case VBK_R8: case VBK_CY: case VBK_DATE: memcpy(dstv, s, 8); break;
    case VBK_U8: *dstv = *s; break;
    case VBK_FSTR: {
        vbs *v = (vbs *)dstv;
        if (v->p && v->len >= fixlen) {
            vbs t = vb_sfromansi(s, fixlen);
            memcpy(v->p, t.p, sizeof(vb_wchar) * (size_t)fixlen);
        }
        break;
    }
    case VBK_UDT: if (u) udt_get(u, dstv, buf, boff); break;
    default: break;                             /* Strings: after, from their buffers */
    }
}
/* ---------------------------------------------------------------- callbacks (AddressOf)
 * A ByRef Type parameter of a callback: the caller's memory (in the VB6 layout) is copied into the
 * program's Type before the call and back after it (its String members stay as they are) */
static char *cb_addr(vb_i64 addr, vb_i32 size)
{
    if (!addr) return 0;
    amap *m = map_find(addr, size);
    if (m) return m->real + (addr - m->fake);
    if (sizeof(void *) == 4) return (char *)(uintptr_t)(vb_u32)addr;   /* 32-bit Windows: the address itself */
    return 0;
}
void vb_cb_udt_in(const vb_udt *u, void *dst, vb_i32 addr)
{
    vb_i32 al, sz = udt_size(u, (const char *)dst, &al);
    char *p = cb_addr(addr, sz);
    if (p) udt_get(u, (char *)dst, p, 0);
}
void vb_cb_udt_out(const vb_udt *u, const void *src, vb_i32 addr)
{
    vb_i32 al, sz = udt_size(u, (const char *)src, &al);
    char *p = cb_addr(addr, sz);
    if (p) udt_put(u, (const char *)src, p, 0, 0);
}
vbobj vb_cb_obj(vb_i32 addr)                   /* ByVal o As Object: the ObjPtr the caller was given */
{
    return (vbobj)cb_addr(addr, 1);
}

/* a Type of numbers of 1, 2 and 4 bytes whose fields lie where VB6 puts them: its own memory is the
 * VB6 layout, and the procedure gets it as it is - so that pts(0) passes the whole array (Polygon) */
static int udt_flat(const vb_udt *u)
{
    vb_i32 off = 0;
    for (int i = 0; i < u->nfields; i++) {
        const vb_ufield *f = &u->fields[i];
        vb_i32 sz;
        if (f->arr) return 0;
        switch (f->kind) {
        case VBK_U8: sz = 1; break;
        case VBK_I2: case VBK_BOOL: sz = 2; break;
        case VBK_I4: case VBK_R4: sz = 4; break;
        default: return 0;
        }
        off = (off + sz - 1) / sz * sz;
        if (f->offset != off) return 0;
        off += sz;
    }
    return u->nfields > 0;
}
/* the arguments of Type: a buffer in the VB6 layout for each, given to the procedure as a pointer */
typedef struct { int i; char *buf; const vb_udt *u; void *src; mstrs strs; } marg;
static int marshal_in(vb_i32 argc, vb_dllarg *argv, marg *ms)
{
    int n = 0;
    for (int i = 0; i < argc; i++) {
        if (argv[i].kind != VBD_UDT) continue;
        const vb_udt *u = (const vb_udt *)argv[i].s;
        marg *m = &ms[n++];
        memset(m, 0, sizeof *m);
        m->i = i; m->u = u; m->src = argv[i].p;
        vb_i32 al, size = udt_size(u, (const char *)m->src, &al);
        if (udt_flat(u) && size == u->size) {   /* its own memory (and the elements after it) */
            argv[i].kind = VBD_PTR;
            argv[i].i = size;
            continue;
        }
        m->buf = (char *)calloc(1, (size_t)size + 16);
        udt_put(u, (const char *)m->src, m->buf, 0, &m->strs);
        argv[i].kind = VBD_PTR;
        argv[i].p = m->buf;
        argv[i].i = size;                       /* the procedures here may want to know it */
    }
    return n;
}
static void marshal_out(vb_dllarg *argv, marg *ms, int n)
{
    for (int k = 0; k < n; k++) {
        marg *m = &ms[k];
        if (!m->buf) { argv[m->i].kind = VBD_UDT; continue; }    /* passed as it is */
        udt_get(m->u, (char *)m->src, m->buf, 0);
        for (int j = 0; j < m->strs.n; j++) {   /* the String members: what the procedure wrote */
            mstr *x = &m->strs.s[j];
            vbs *v = x->str;
            vb_i32 addr; memcpy(&addr, m->buf + x->off, 4);
            const char *b = addr ? (const char *)vb_ptr_resolve(addr, 1) : 0;
            if (b && v->p) {
                vbs t = vb_sfromansi(b, v->len);
                if (t.len == v->len) memcpy(v->p, t.p, sizeof(vb_wchar) * (size_t)v->len);
            }
            bstr_free(x->buf);
        }
        free(m->strs.s);
        free(m->buf);
        argv[m->i].kind = VBD_UDT;
        argv[m->i].p = m->src;
    }
}

/* LSet a = b of two Types: the bytes of b (VB6 layout) over those of a */
void vb_udt_lset(void *d, const vb_udt *du, const void *s, const vb_udt *su)
{
    if (du == su) {
        /* the same Type: VB6 copies the bytes, a String member its BSTR pointer, so that t = a(j): a(j) = a(i):
         * a(i) = t swaps the Strings too (a sort of PSC: will-barden CombSort). Here the copy of the Type, its
         * Strings as copies (the bytes of the layout have no Strings, so they were left as they were) */
        if (d == s) return;
        if (du->copy) du->copy(d, s); else memcpy(d, s, (size_t)du->size);
        return;
    }
    vb_i32 al, ns = udt_size(su, (const char *)s, &al), nd = udt_size(du, (const char *)d, &al);
    char *bs = (char *)calloc(1, (size_t)ns + 8), *bd = (char *)calloc(1, (size_t)nd + 8);
    mstrs m1 = { 0, 0, 0 }, m2 = { 0, 0, 0 };
    udt_put(su, (const char *)s, bs, 0, &m1);
    udt_put(du, (const char *)d, bd, 0, &m2);
    memcpy(bd, bs, (size_t)(ns < nd ? ns : nd));
    udt_get(du, (char *)d, bd, 0);
    for (int j = 0; j < m1.n; j++) bstr_free(m1.s[j].buf);
    for (int j = 0; j < m2.n; j++) bstr_free(m2.s[j].buf);
    free(m1.s); free(m2.s); free(bs); free(bd);
}

/* the ANSI copies of the ByVal String arguments of the call that runs (the procedures of the forms
 * runtime read them with vb_dllarg_str) */
static vb_dllarg *cur_argv; static char **cur_ansi; static int cur_argc;
char *vb_dllarg_str(vb_dllarg *a)
{
    if (!a) return 0;
    if (a->kind == VBD_STR) {
        if (cur_argv && a >= cur_argv && a < cur_argv + cur_argc) return cur_ansi[a - cur_argv];
        return 0;
    }
    if (a->kind == VBD_INT) return a->i ? (char *)vb_ptr_resolve(a->i, 1) : 0;
    if (a->kind == VBD_PTR) return (char *)a->p;
    return 0;
}

/* a Variant ByRef (As Variant, As Any): a VARIANT of 16 bytes - vt, then the value at offset 8; a String
 * as a BSTR of its Unicode text */
typedef struct { int i; vbvar *v; unsigned char buf[16], was[16]; char *bstr; } mvar;
static void var_put(mvar *m)
{
    vbvar *v = m->v;
    unsigned char *b = m->buf;
    memset(b, 0, 16);
    vb_u16 vt = v->vt;
    m->bstr = 0;
    switch (vt) {
    case VT_I2: case VT_BOOL: memcpy(b + 8, &v->u.i, 2); break;
    case VT_I4: case VT_ERROR: memcpy(b + 8, &v->u.l, 4); break;
    case VT_R4: memcpy(b + 8, &v->u.f, 4); break;
    case VT_R8: case VT_DATE: memcpy(b + 8, &v->u.d, 8); break;
    case VT_CY: memcpy(b + 8, &v->u.cy, 8); break;
    case VT_UI1: b[8] = v->u.b; break;
    case VT_DECIMAL:
        b[2] = (unsigned char)(v->pad & 0xFF); b[3] = (v->pad & 0x8000) ? 0x80 : 0;
        memcpy(b + 4, &v->u.dec.hi, 4); memcpy(b + 8, &v->u.dec.lo, 4); memcpy(b + 12, &v->u.dec.mid, 4);
        break;
    case VT_BSTR:
        if (v->u.s.p) {
            m->bstr = bstr_new(v->u.s.p, v->u.s.len * 2);
#ifdef _WIN32
            void *pp = m->bstr; memcpy(b + 8, &pp, sizeof pp);
#else
            vb_i32 a = vb_varptr(m->bstr, v->u.s.len * 2 + 2); memcpy(b + 8, &a, 4);
#endif
        }
        break;
    case VT_EMPTY: case VT_NULL: break;
    default: vt = v->vt; break;                 /* objects, arrays: the type only */
    }
    memcpy(b, &vt, 2);
    memcpy(m->was, b, 16);
}
static void var_get(mvar *m)
{
    unsigned char *b = m->buf;
    if (memcmp(b, m->was, 16)) {                /* the procedure wrote into it */
        vb_u16 vt; memcpy(&vt, b, 2);
        vbvar n; memset(&n, 0, sizeof n);
        n.vt = vt;
        int ok = 1;
        switch (vt) {
        case VT_EMPTY: case VT_NULL: break;
        case VT_I2: case VT_BOOL: memcpy(&n.u.i, b + 8, 2); break;
        case VT_I4: case VT_ERROR: memcpy(&n.u.l, b + 8, 4); break;
        case VT_R4: memcpy(&n.u.f, b + 8, 4); break;
        case VT_R8: case VT_DATE: memcpy(&n.u.d, b + 8, 8); break;
        case VT_CY: memcpy(&n.u.cy, b + 8, 8); break;
        case VT_UI1: n.u.b = b[8]; break;
        default: ok = 0;
        }
        if (ok) { vb_vset(m->v, n); }
    }
    bstr_free(m->bstr);
}

static vb_i64 call(vb_declare *d, vb_i32 argc, vb_dllarg *argv, int *ok)
{
    *ok = 0;
    if (vb_errpend || !resolve(d)) return 0;
    if (argc > 16) { vb_raise(450); return 0; }
    for (int i = 0; i < argc; i++)            /* a Variant ByVal to As Any: VB6 pushes its 16 bytes */
        if (argv[i].kind == VBD_BADCC) { vb_raise(49); return 0; }
    marg mbuf[16];
    int nm = marshal_in(argc, argv, mbuf);
    /* ByVal String: an ANSI BSTR, copied back afterwards (the procedure may write into it);
     * ByRef String: the address of one, and what the procedure leaves there comes back */
    char *ansi[16] = { 0 };
    vb_i64 slot[16] = { 0 };            /* the BSTR of a ByRef String: its number here, its address for a DLL */
    int real = d->state == 2;
    mvar mv[16]; int nv = 0;
    for (int i = 0; i < argc; i++) {
        if (argv[i].kind == VBD_STR) ansi[i] = bstr_ansi(argv[i].s);
        else if (argv[i].kind == VBD_SREF) {
            char *b = bstr_ansi(argv[i].s);
            slot[i] = real ? (vb_i64)(intptr_t)b : bstr_ret(b);
            argv[i].kind = VBD_PTR;
            argv[i].p = &slot[i];
            argv[i].i = -1;                     /* (marks it, for after the call) */
        } else if (argv[i].kind == VBD_OREF) {   /* an object ByRef to As Any: its ObjPtr (CopyMemory o, lPtr, 4) */
            vbobj *ov = (vbobj *)argv[i].p;
            slot[i] = *ov ? (vb_i64)vb_varptr(vb_oid(*ov), (vb_i32)sizeof(vbobjh)) : 0;
            argv[i].i = slot[i];               /* (the value before the call) */
            argv[i].s = (vbs *)ov;
            argv[i].kind = VBD_PTR;
            argv[i].p = &slot[i];
        } else if (argv[i].kind == VBD_VAR) {
            mvar *m = &mv[nv++];
            m->i = i; m->v = (vbvar *)argv[i].p;
            var_put(m);
            argv[i].kind = VBD_PTR;
            argv[i].p = m->buf;
        }
    }
    for (int i = 0; i < argc; i++)            /* VB6: a call with a String argument (not vbNullString) starts from 0 */
        if (ansi[i] || (argv[i].kind == VBD_PTR && argv[i].p == &slot[i] && slot[i])) {   /* (not vbNullString) */
#ifdef _WIN32
            SetLastError(0);
#else
            emu_lasterr = 0;
#endif
            break;
        }
    args x = { argv, argc, ansi };
    vb_i64 r = 0;
#ifdef _WIN32
#define LASTERR() (vb_lastdllerr = (vb_i32)GetLastError())
#else
#define LASTERR() (vb_lastdllerr = emu_lasterr)
#endif
    if (d->state == 1) { r = ((apifn)d->fn)(&x); LASTERR(); }
    else if (d->state == 3) {
        vb_dllarg *sa = cur_argv; char **sn = cur_ansi; int sc = cur_argc;
        cur_argv = argv; cur_ansi = ansi; cur_argc = argc;
        r = ((vb_apiext)d->fn)(argc, argv);
        LASTERR();
        cur_argv = sa; cur_ansi = sn; cur_argc = sc;
    }
    else {
#ifdef _WIN32
        if (bstr_intercept(d->entry, &x)) { r = 0; goto bstr_done; }
#if defined(__i386__)
        unsigned w[16]; int nw = 0;
        for (int i = 0; i < argc && nw < 15; i++) {
            if (argv[i].kind == VBD_DBL) { union { double d; unsigned u[2]; } c; c.d = argv[i].d; w[nw++] = c.u[0]; w[nw++] = c.u[1]; }
            else w[nw++] = argv[i].kind == VBD_PTR ? (unsigned)(uintptr_t)argv[i].p : argv[i].kind == VBD_STR ? (unsigned)(uintptr_t)ansi[i]
                         : (unsigned)argv[i].i;
        }
        r = (vb_i32)call_std(d->fn, w, nw);
        LASTERR();
        goto copied;
#else
        for (int i = 0; i < argc; i++) if (argv[i].kind == VBD_INT) argv[i].i = real_int(argv[i].i);
#endif
#endif
        vb_i64 v[8] = { 0 };
        for (int i = 0; i < argc && i < 8; i++)
            v[i] = argv[i].kind == VBD_PTR ? (vb_i64)(intptr_t)argv[i].p : argv[i].kind == VBD_STR ? (vb_i64)(intptr_t)ansi[i]
                 : argv[i].kind == VBD_DBL ? (vb_i64)argv[i].d : argv[i].i;
        r = ((vb_i64 (*)(vb_i64, vb_i64, vb_i64, vb_i64, vb_i64, vb_i64, vb_i64, vb_i64))d->fn)
            (v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7]);
        LASTERR();
    }
#undef LASTERR
#ifdef _WIN32
bstr_done:
#endif
#if defined(_WIN32) && defined(__i386__)
copied:
#endif
    for (int i = 0; i < argc; i++) {
        if (ansi[i]) {
            vbs *s = argv[i].s;
            vbs t = vb_sfromansi(ansi[i], s->len);     /* same length as before the call */
            if (t.len == s->len) memcpy(s->p, t.p, sizeof(vb_wchar) * (size_t)s->len);
            bstr_free(ansi[i]);
        }
        if (argv[i].kind == VBD_PTR && argv[i].p == &slot[i] && argv[i].s && argv[i].i != -1) {
            vbobj *ov = (vbobj *)argv[i].s;     /* the object variable: what the procedure wrote, as VB6 (no AddRef) */
            vb_i32 nv = (vb_i32)slot[i];
            if (nv != (vb_i32)argv[i].i) *ov = nv ? (vbobj)vb_ptr_resolve(nv, 1) : 0;
            argv[i].kind = VBD_OREF; argv[i].p = ov; argv[i].s = 0;
            continue;
        }
        if (argv[i].kind == VBD_PTR && argv[i].p == &slot[i] && argv[i].i == -1) {
            char *b = real ? (char *)(intptr_t)slot[i] : slot[i] ? (char *)vb_ptr_resolve(slot[i], 1) : 0;
            vb_sset(argv[i].s, bstr_str(b, bstr_len(b)));
            bstr_free(b);
            argv[i].kind = VBD_SREF;
        }
    }
    for (int k = 0; k < nv; k++) { var_get(&mv[k]); argv[mv[k].i].kind = VBD_VAR; argv[mv[k].i].p = mv[k].v; }
    if (nm) marshal_out(argv, mbuf, nm);
    *ok = 1;
    return r;
}

vb_i64 vb_dllcall(vb_declare *d, vb_i32 argc, vb_dllarg *argv) { int ok; return call(d, argc, argv, &ok); }
double vb_dllcall_d(vb_declare *d, vb_i32 argc, vb_dllarg *argv) { int ok; return (double)call(d, argc, argv, &ok); }
/* As String: the procedure returns a BSTR of ANSI bytes; VB6 makes them a String and frees the BSTR.
 * (A shared object of Linux returns a C string, which stays its own.) */
vbs vb_dllcall_s(vb_declare *d, vb_i32 argc, vb_dllarg *argv)
{
    int ok;
    vb_i64 r = call(d, argc, argv, &ok);
    vbs z = { 0, 0, 0 };
    if (!ok || !r) return z;
    vb_lastdllerr = 0;                          /* (VB6: making the String of the result leaves 0) */
    char *p;
    if (d->state == 2) {
#ifdef _WIN32
        p = sizeof(void *) == 4 ? (char *)(uintptr_t)(vb_u32)r : (char *)(intptr_t)r;
        vbs s = bstr_str(p, bstr_len(p));
        bstr_free(p);
        return s;
#else
        p = (char *)(intptr_t)r;
        return bstr_str(p, (vb_i32)strlen(p));
#endif
    }
    p = (char *)vb_ptr_resolve(r, 1);
    if (!p) return z;
    if (bstr_ours(p)) { vbs s = bstr_str(p, bstr_len(p)); bstr_free(p); return s; }
    return bstr_str(p, (vb_i32)strlen(p));
}
