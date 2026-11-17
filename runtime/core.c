/* core.c - strings, errors, numbers and their conversion to text.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"
#include "vbos.h"
#include <locale.h>
#include <stdint.h>
#ifndef _WIN32
#include <sys/resource.h>
#endif

/* ---------------------------------------------------------------- code page */
int vb_codepage = 1252;
vb_wchar vb_ansi2w(unsigned char c)
{
    if (c < 0x80) return c;
    vb_wchar w; vb_cp_decode(vb_codepage, &c, 1, &w);
    return w;
}

/* the single byte of a character in the code page (-1: none; no best fit: Option Compare Binary uses it) */
int vb_w2ansi(vb_wchar w)
{
    if (w < 0x80) return w;
    int c = vb_cp_code(vb_codepage, w, 0);
    return c >= 0 && c <= 0xFF ? c : -1;
}

/* ---------------------------------------------------------------- strings */
/* Temporaries of the current statement: string buffers (kind 0) and object
 * references (kind 1). vb_release(mark) frees / releases them in reverse order;
 * releasing an object may run Class_Terminate, which uses the same stack above
 * the current top, so the top is lowered before each release. */
/* (a stack that grows: an outer statement's mark must stay valid while events and procedures run inside it) */
static void **tmpbuf;
static char *tmpkind;
static vb_i32 ntmp, captmp;

vb_i32 vb_mark(void) { return ntmp; }
void vb_release(vb_i32 mark)
{
    while (ntmp > mark) {
        --ntmp;
        if (tmpkind[ntmp] == 1) vb_orelease((vbobj)tmpbuf[ntmp]);
        else if (tmpkind[ntmp] == 2) { vb_afree((vbarr *)tmpbuf[ntmp]); free(tmpbuf[ntmp]); }
        else if (tmpkind[ntmp] == 3) vb_rec_free((struct vbrec *)tmpbuf[ntmp]);
        else free(tmpbuf[ntmp]);
    }
}

static void tmp_push(void *p, int kind)
{
    if (ntmp == captmp) {
        vb_i32 nc = captmp ? captmp * 2 : 4096;
        void **nb = (void **)realloc(tmpbuf, sizeof(void *) * (size_t)nc);
        char *nk = nb ? (char *)realloc(tmpkind, (size_t)nc) : 0;
        if (!nb || !nk) {                       /* no memory: Out of memory (the buffer is left alone) */
            if (nb) tmpbuf = nb;
            vb_raise(7);
            return;
        }
        tmpbuf = nb; tmpkind = nk; captmp = nc;
    }
    tmpbuf[ntmp] = p; tmpkind[ntmp++] = (char)kind;
}

vbobj vb_otmp(vbobj o) { if (o) tmp_push(o, 1); return o; }
/* ---- a Type in a Variant (VT_RECORD): its descriptor, then its data (aligned to 16) */
struct vbrec { const vb_udt *udt; vb_i64 tmp; };   /* tmp: one of the temporaries (vb_release frees it) */
#define REC_DATA(r) ((char *)(r) + sizeof(struct vbrec))
const vb_udt *vb_rec_udt(const struct vbrec *r) { return r ? r->udt : 0; }
void *vb_rec_data(struct vbrec *r) { return r ? REC_DATA(r) : 0; }
struct vbrec *vb_rec_new(const vb_udt *u, const void *src)
{
    struct vbrec *r = (struct vbrec *)calloc(1, sizeof(struct vbrec) + (size_t)(u->size > 0 ? u->size : 1));
    if (!r) { vb_raise(7); return 0; }
    r->udt = u;
    if (u->init) u->init(REC_DATA(r));
    if (src) {
        if (u->copy) u->copy(REC_DATA(r), src);
        else memcpy(REC_DATA(r), src, (size_t)u->size);
    }
    return r;
}
void vb_rec_free(struct vbrec *r)
{
    if (!r) return;
    if (r->udt && r->udt->free) r->udt->free(REC_DATA(r));
    free(r);
}
struct vbrec *vb_rec_tmp(const vb_udt *u, const void *src)
{
    struct vbrec *r = vb_rec_new(u, src);
    if (r) { r->tmp = 1; tmp_push(r, 3); }
    return r;
}
int vb_rec_istmp(const struct vbrec *r) { return r && r->tmp; }
vbvar vb_v_rec(const void *p, const vb_udt *u)
{
    vbvar v; memset(&v, 0, sizeof v);
    struct vbrec *r = vb_rec_tmp(u, p);
    if (!r) return v;
    v.vt = VT_RECORD; v.u.rec = r;
    return v;
}
/* the same Type: one descriptor, or two of the same Public Type (a component and the program that uses it: the
 * same declaration, so the same layout) */
static int ascii_ieq(const char *a, const char *b)
{
    for (; *a && *b; a++, b++) {
        int x = (unsigned char)*a, y = (unsigned char)*b;
        if (x >= 'A' && x <= 'Z') x += 32;
        if (y >= 'A' && y <= 'Z') y += 32;
        if (x != y) return 0;
    }
    return *a == *b;
}
static int rec_same(const vb_udt *a, const vb_udt *b)
{
    if (a == b) return 1;
    return a && b && a->name && b->name && a->size == b->size && a->nfields == b->nfields &&
           ascii_ieq(a->name, b->name);
}
void vb_v2rec(vbvar v, const vb_udt *u, void *dst)
{
    if (v.vt != VT_RECORD || !v.u.rec || !rec_same(v.u.rec->udt, u)) { vb_raise(13); return; }
    if (u->copy) u->copy(dst, REC_DATA(v.u.rec));
    else memcpy(dst, REC_DATA(v.u.rec), (size_t)u->size);
}

/* v.Qty, v.Tags(1), v.Qty = 5 of a Variant holding a Type (late binding: by the names of the fields) */
static vb_i32 vbk_vt(vb_i32 k)
{
    switch (k) {
    case VBK_I2: return VT_I2; case VBK_I4: return VT_I4; case VBK_R4: return VT_R4; case VBK_R8: return VT_R8;
    case VBK_CY: return VT_CY; case VBK_DATE: return VT_DATE; case VBK_BOOL: return VT_BOOL; case VBK_U8: return VT_UI1;
    case VBK_STR: case VBK_FSTR: return VT_BSTR; case VBK_VAR: return VT_VARIANT; case VBK_OBJ: return VT_DISPATCH;
    default: return -1;
    }
}
static vbvar ufield_get(const vb_ufield *f, const char *p)
{
    switch (f->kind) {
    case VBK_I2: return vb_v_i(*(const vb_i16 *)p);
    case VBK_I4: return vb_v_l(*(const vb_i32 *)p);
    case VBK_R4: return vb_v_f(*(const float *)p);
    case VBK_R8: return vb_v_d(*(const double *)p);
    case VBK_CY: return vb_v_c(*(const vb_cy *)p);
    case VBK_DATE: return vb_v_date(*(const vb_date *)p);
    case VBK_BOOL: return vb_v_bool(*(const vb_bool *)p);
    case VBK_U8: return vb_v_byte(*(const vb_u8 *)p);
    case VBK_STR: case VBK_FSTR: return vb_vtmp(vb_v_s(*(const vbs *)p));
    case VBK_VAR: return vb_vtmp(*(const vbvar *)p);
    case VBK_OBJ: return vb_vtmp(vb_v_obj(*(vbobj const *)p));
    case VBK_UDT: if (f->udt && f->udt->name) return vb_v_rec(p, f->udt);
    }
    vb_raise(13);
    return vb_v_empty();
}
static void ufield_let(const vb_ufield *f, char *p, vbvar x, vb_i32 kind)
{
    if (kind == VBM_SET) {
        if (f->kind == VBK_OBJ) vb_oset((vbobj *)p, vb_v2obj(x));
        else if (f->kind == VBK_VAR) vb_vset((vbvar *)p, x);
        else vb_raise(13);
        return;
    }
    switch (f->kind) {
    case VBK_I2: { vb_i16 t = vb_v2i(x); if (!vb_errpend) *(vb_i16 *)p = t; break; }
    case VBK_I4: { vb_i32 t = vb_v2l(x); if (!vb_errpend) *(vb_i32 *)p = t; break; }
    case VBK_R4: { float t = vb_v2f(x); if (!vb_errpend) *(float *)p = t; break; }
    case VBK_R8: { double t = vb_v2d(x); if (!vb_errpend) *(double *)p = t; break; }
    case VBK_CY: { vb_cy t = vb_v2c(x); if (!vb_errpend) *(vb_cy *)p = t; break; }
    case VBK_DATE: { vb_date t = vb_v2date(x); if (!vb_errpend) *(vb_date *)p = t; break; }
    case VBK_BOOL: { vb_bool t = vb_v2bool(x); if (!vb_errpend) *(vb_bool *)p = t; break; }
    case VBK_U8: { vb_u8 t = vb_v2byte(x); if (!vb_errpend) *(vb_u8 *)p = t; break; }
    case VBK_STR: { vbs t = vb_v2s(x); if (!vb_errpend) vb_sset((vbs *)p, t); break; }
    case VBK_VAR: vb_vset((vbvar *)p, x); break;
    case VBK_UDT: if (f->udt && f->udt->name) { vb_v2rec(x, f->udt, p); break; } vb_raise(13); break;
    default: vb_raise(13);
    }
}
vbvar vb_rec_member(vbvar v, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv)
{
    if (vb_errpend) return vb_v_empty();
    const vb_udt *u = v.u.rec ? v.u.rec->udt : 0;
    vb_i32 i = 0;
    if (u && u->fnames && *name)
        for (; i < u->nfields; i++) if (ascii_ieq(u->fnames[i], name)) break;
    if (!u || !u->fnames || !*name || i == u->nfields) { vb_raise(461); return vb_v_empty(); }   /* (VB6: 461) */
    const vb_ufield *f = &u->fields[i];
    char *p = (char *)REC_DATA(v.u.rec) + f->offset;
    vb_i32 nargs = (kind == VBM_LET || kind == VBM_SET) ? argc - 1 : argc;
    if (f->arr) {                                  /* v.Tags(1) */
        vb_i32 vt = vbk_vt(f->kind);
        if (vt < 0) { vb_raise(13); return vb_v_empty(); }
        if (kind == VBM_LET || kind == VBM_SET) {
            vb_i32 ix[3];                          /* v.Tags(1) = "x": the element */
            if (nargs <= 0 || nargs > 3) { vb_raise(nargs <= 0 ? 13 : 9); return vb_v_empty(); }
            for (vb_i32 j = 0; j < nargs; j++) ix[j] = vb_v2l(argv[j]);
            if (vb_errpend) return vb_v_empty();
            vbarr *a = (vbarr *)p;
            void *e = nargs == 1 ? vb_aelem(a, 1, ix[0]) : nargs == 2 ? vb_aelem(a, 2, ix[0], ix[1])
                                                                       : vb_aelem(a, 3, ix[0], ix[1], ix[2]);
            if (!e || vb_errpend) return vb_v_empty();
            vb_ufield ef = *f; ef.arr = 0;
            ufield_let(&ef, (char *)e, argv[argc - 1], kind);
            return vb_v_empty();
        }
        if (nargs <= 0) { vb_raise(13); return vb_v_empty(); }   /* v.Tags without an index: VB6's error 13 */
        vbvar a = vb_v_arr((const vbarr *)p, vt);
        return vb_v_index(a, nargs, argv);
    }
    if (nargs > 0) { vb_raise(13); return vb_v_empty(); }
    if (kind == VBM_LET || kind == VBM_SET) { ufield_let(f, p, argv[argc - 1], kind); return vb_v_empty(); }
    return ufield_get(f, p);
}

vbarr *vb_atmp(void) { vbarr *a = (vbarr *)calloc(1, sizeof(vbarr)); if (!a) { vb_raise(7); return 0; } tmp_push(a, 2); return a; }

vbs vb_stmp(vb_i32 len)
{
    vbs s;
    if (len < 0) len = 0;
    if (len > 0x20000000) {                     /* more than VB6 can ever have (1 GB) */
        vb_raise(14); s.len = 0; s.flags = 0; static vb_wchar z0; s.p = &z0; return s;
    }
    s.p = (vb_wchar *)malloc(sizeof(vb_wchar) * (len + 1));
    if (!s.p) { vb_raise(7); s.len = 0; s.flags = 0; static vb_wchar z; s.p = &z; return s; }
    s.len = len; s.flags = 0;
    s.p[len] = 0;
    tmp_push(s.p, 0);
    return s;
}

vbs vb_slit(const vb_wchar *p, vb_i32 len) { vbs s; s.p = (vb_wchar *)p; s.len = len; s.flags = 0; return s; }

void vb_sset(vbs *dst, vbs src)
{
    int odd = (src.flags & VBS_ODD) != 0;
    if (!src.p) { vb_sfree(dst); return; }      /* vbNullString stays a null string (StrPtr = 0) */
    vb_wchar *p = (vb_wchar *)malloc(sizeof(vb_wchar) * (src.len + 1 + odd));
    if (!p) { vb_raise(7); return; }
    if (src.len + odd) memcpy(p, src.p, sizeof(vb_wchar) * (src.len + odd));
    p[src.len + odd] = 0;
    if (dst->flags & 1) free(dst->p);
    dst->p = p; dst->len = src.len; dst->flags = 1 | (odd ? VBS_ODD : 0);
}

void vb_sfree(vbs *v)
{
    if (v->flags & 1) free(v->p);
    v->p = 0; v->len = 0; v->flags = 0;
}

vbs vb_scopy(vbs s)
{
    if (s.flags & VBS_ODD) return vb_sfrombytes2(s, 0, vb_sbytes(s));
    vbs r = vb_stmp(s.len); if (s.len) memcpy(r.p, s.p, sizeof(vb_wchar) * s.len); return r;
}

/* ---------------------------------------------------------------- byte strings (LenB, ChrB, MidB ...) */
vb_i32 vb_sbytes(vbs s) { return s.len * 2 + ((s.flags & VBS_ODD) ? 1 : 0); }
int vb_sbyte(vbs s, vb_i32 i)                   /* byte i (from 0) */
{
    return i / 2 < s.len ? (s.p[i / 2] >> (8 * (i & 1))) & 0xFF : s.p[s.len] & 0xFF;
}
/* a string of n bytes (an odd number: VBS_ODD) */
vbs vb_sfrombytes(const char *b, vb_i32 n)
{
    if (n < 0) n = 0;
    vbs r = vb_stmp(n / 2 + (n & 1));
    for (vb_i32 i = 0; i + 1 < n; i += 2) r.p[i / 2] = (vb_wchar)((unsigned char)b[i] | ((unsigned char)b[i + 1] << 8));
    if (n & 1) { r.len = n / 2; r.p[r.len] = (unsigned char)b[n - 1]; r.flags |= VBS_ODD; }
    return r;
}
/* the bytes [from, from + n) of s as a string */
vbs vb_sfrombytes2(vbs s, vb_i32 from, vb_i32 n)
{
    if (n < 0) n = 0;
    char *b = (char *)malloc((size_t)(n > 0 ? n : 1));
    for (vb_i32 i = 0; i < n; i++) b[i] = (char)vb_sbyte(s, from + i);
    vbs r = vb_sfrombytes(b, n);
    free(b);
    return r;
}

vbs vb_scat(vbs a, vbs b)
{
    if ((a.flags | b.flags) & VBS_ODD) {        /* byte strings: the bytes follow each other */
        vb_i32 na = vb_sbytes(a), nb = vb_sbytes(b);
        char *t = (char *)malloc((size_t)(na + nb + 1));
        for (vb_i32 i = 0; i < na; i++) t[i] = (char)vb_sbyte(a, i);
        for (vb_i32 i = 0; i < nb; i++) t[na + i] = (char)vb_sbyte(b, i);
        vbs r = vb_sfrombytes(t, na + nb);
        free(t);
        return r;
    }
    vbs r = vb_stmp(a.len + b.len);
    if (a.len) memcpy(r.p, a.p, sizeof(vb_wchar) * a.len);
    if (b.len) memcpy(r.p + a.len, b.p, sizeof(vb_wchar) * b.len);
    return r;
}

#define wlower vb_wlower
#define wupper vb_wupper

vb_i32 vb_scmp(vbs a, vbs b, vb_i32 text)
{
    vb_i32 n = a.len < b.len ? a.len : b.len;
    if ((a.flags | b.flags) & VBS_ODD) {        /* byte strings: the characters, then the bytes left over */
        for (vb_i32 i = 0; i < n; i++) {
            vb_wchar x = a.p[i], y = b.p[i];
            if (text) { x = wlower(x); y = wlower(y); }
            if (x != y) return x < y ? -1 : 1;
        }
        vb_i32 na = vb_sbytes(a), nb = vb_sbytes(b);
        for (vb_i32 k = 2 * n; k < na && k < nb; k++) {
            int x = vb_sbyte(a, k), y = vb_sbyte(b, k);
            if (x != y) return x < y ? -1 : 1;
        }
        return na < nb ? -1 : na > nb ? 1 : 0;
    }
    if (text) return vb_collate(a, b);          /* Option Compare Text: CompareString (intl.c) */
    for (vb_i32 i = 0; i < n; i++) {
        vb_wchar x = a.p[i], y = b.p[i];
        {   /* Option Compare Binary compares the ANSI codes, as VB6 does */
            int ax = vb_w2ansi(x), ay = vb_w2ansi(y);
            if (ax >= 0 && ay >= 0) { x = (vb_wchar)ax; y = (vb_wchar)ay; }
        }
        if (x != y) return x < y ? -1 : 1;
    }
    return a.len < b.len ? -1 : a.len > b.len ? 1 : 0;
}

vbs vb_sascii(const char *s)
{
    vb_i32 n = (vb_i32)strlen(s);
    vbs r = vb_stmp(n);
    for (vb_i32 i = 0; i < n; i++) r.p[i] = (unsigned char)s[i];
    return r;
}

vbs vb_sfromansi(const char *s, int n)
{
    vbs r = vb_stmp(n);
    r.len = vb_cp_decode(vb_codepage, (const unsigned char *)s, n, r.p);
    r.p[r.len] = 0;
    return r;
}

/* the text in the ANSI code page (best fit, '?' for the rest); a temporary buffer */
char *vb_stoansi(vbs s, int *n)
{
    vbs t = vb_stmp(s.len + 1);                 /* buffer owned by the temporaries: 2 bytes a character at most */
    char *o = (char *)t.p;
    int k = vb_cp_encode(vb_codepage, s.p, s.len, (unsigned char *)o);
    o[k] = 0;
    if (n) *n = k;
    return o;
}

vbs vb_sfromutf8(const char *s)
{
    vb_i32 n = (vb_i32)strlen(s), k = 0;
    vbs r = vb_stmp(n);
    for (vb_i32 i = 0; i < n;) {
        unsigned c = (unsigned char)s[i];
        unsigned cp; int l;
        if (c < 0x80) { cp = c; l = 1; }
        else if ((c & 0xE0) == 0xC0 && i + 1 < n) { cp = ((c & 0x1F) << 6) | (s[i + 1] & 0x3F); l = 2; }
        else if ((c & 0xF0) == 0xE0 && i + 2 < n) { cp = ((c & 0x0F) << 12) | ((s[i + 1] & 0x3F) << 6) | (s[i + 2] & 0x3F); l = 3; }
        else { cp = '?'; l = 1; }
        r.p[k++] = (vb_wchar)cp; i += l;
    }
    r.len = k; r.p[k] = 0;
    return r;
}

char *vb_stoutf8(vbs s)
{
    vbs t = vb_stmp(s.len * 2 + 1);             /* 3 bytes per char at most */
    unsigned char *o = (unsigned char *)t.p; int k = 0;
    for (vb_i32 i = 0; i < s.len; i++) {
        unsigned c = s.p[i];
        if (c < 0x80) o[k++] = c;
        else if (c < 0x800) { o[k++] = 0xC0 | (c >> 6); o[k++] = 0x80 | (c & 0x3F); }
        else { o[k++] = 0xE0 | (c >> 12); o[k++] = 0x80 | ((c >> 6) & 0x3F); o[k++] = 0x80 | (c & 0x3F); }
    }
    o[k] = 0;
    return (char *)o;
}

vbs vb_lcase(vbs s) { vbs r = vb_stmp(s.len); for (vb_i32 i = 0; i < s.len; i++) r.p[i] = wlower(s.p[i]); return r; }
vbs vb_ucase(vbs s) { vbs r = vb_stmp(s.len); for (vb_i32 i = 0; i < s.len; i++) r.p[i] = wupper(s.p[i]); return r; }

/* ---------------------------------------------------------------- errors */
vb_i32 vb_errpend;
vb_i32 vb_fpchecks = 1;       /* 0: Remove Floating Point Error Checks (native code of VB6) */
vb_i32 vb_fpstore;            /* the pending error comes from a floating-point operation: VB6 stores its result
                                 (Inf, NaN) and then raises the error */
static vb_i32 pend_num, err_num, err_erl;
static vbs pend_src, pend_desc, err_src, err_desc;
static int pend_has_desc;
static vb_i32 pend_ctx, err_ctx; static int pend_has_ctx;

#include "errors.inc"
const char *vb_errtext(int n)
{
    for (int i = 0; vb_errtab[i].n; i++) if (vb_errtab[i].n == n) return vb_errtab[i].s;
    return n >= 1 && n <= 65535 ? "Application-defined or object-defined error" : "Invalid procedure call or argument";
}
/* the text of Error$(n) / Err.Raise n: VB6 knows only those of VBA there (87); the others (the forms
 * errors ...) have their text only when the runtime itself raises them */
static int errtab_pub(int n)
{
    for (int i = 0; vb_errtab[i].n; i++) if (vb_errtab[i].n == n) return vb_errtab[i].pub;
    return 0;
}
static const char *errtext_pub(int n)
{
    return errtab_pub(n) ? vb_errtext(n) : "Application-defined or object-defined error";
}
/* Err.HelpContext of an error: 1000000 + n in the help of VBA (1000095: "Application-defined ...") */
static vb_i32 help_context(vb_i32 n)
{
    if (n < 1 || n > 65535) return 0;
    for (int i = 0; vb_errtab[i].n; i++) if (vb_errtab[i].n == n) return 1000000 + n;
    return 1000095;
}

extern const char *vb_project_title;
static vbs persist(vbs *slot, vbs v) { vb_sset(slot, v); return *slot; }

void vb_raise(vb_i32 number)
{
    if (vb_errpend) return;                     /* the first error of a statement wins */
    vb_errpend = 1; pend_num = number; pend_has_desc = 0; pend_has_ctx = 0; vb_fpstore = 0;
}

void vb_fpraise(vb_i32 number)
{
    if (!vb_fpchecks || vb_errpend) return;
    vb_raise(number);
    vb_fpstore = 1;
}

void vb_raise_full(vb_i32 number, vbs source, vbs description)
{
    if (vb_errpend) return;
    vb_errpend = 1; pend_num = number; vb_fpstore = 0; pend_has_ctx = 0;
    persist(&pend_src, source);
    persist(&pend_desc, description);
    pend_has_desc = 1;
}

/* an HRESULT given to Err.Raise: the VB error it stands for (VB6 maps the common ones) */
static vb_i32 hresult_vb(vb_u32 h)
{
    if ((h & 0xFFFF0000u) == 0x800A0000u && (h & 0xFFFF)) return (vb_i32)(h & 0xFFFF);   /* FACILITY_CONTROL */
    switch (h) {
    case 0x80070005u: return 70;                /* E_ACCESSDENIED */
    case 0x8007000Eu: return 7;                 /* E_OUTOFMEMORY */
    case 0x80070057u: return 5;                 /* E_INVALIDARG */
    case 0x80004002u: return 430;               /* E_NOINTERFACE */
    case 0x80020003u: return 461;               /* DISP_E_MEMBERNOTFOUND */
    case 0x80020005u: return 13;                /* DISP_E_TYPEMISMATCH */
    case 0x80020006u: return 438;               /* DISP_E_UNKNOWNNAME */
    case 0x80020007u: return 446;               /* DISP_E_NONAMEDARGS */
    case 0x8002000Au: return 6;                 /* DISP_E_OVERFLOW */
    case 0x8002000Bu: return 9;                 /* DISP_E_BADINDEX */
    case 0x8002000Eu: return 450;               /* DISP_E_BADPARAMCOUNT */
    case 0x8002000Fu: return 449;               /* DISP_E_PARAMNOTOPTIONAL */
    case 0x80020012u: return 11;                /* DISP_E_DIVBYZERO */
    case 0x800401F3u: return 429;               /* CO_E_CLASSSTRING */
    case 0x80040154u: return 429;               /* REGDB_E_CLASSNOTREG */
    case 0x80070002u: return 53;                /* ERROR_FILE_NOT_FOUND */
    case 0x80070003u: return 76;                /* ERROR_PATH_NOT_FOUND */
    }
    return 0;
}
/* the system text of an HRESULT (FormatMessage on Windows; the common ones here) */
static const char *hresult_text(vb_u32 h)
{
    static const struct { vb_u32 h; const char *s; } t[] = {
        { 0x80004001u, "Not implemented" }, { 0x80004003u, "Invalid pointer" },
        { 0x80004004u, "Operation aborted" }, { 0x80004005u, "Unspecified error" },
        { 0x8000FFFFu, "Catastrophic failure" }, { 0x80040001u, "Invalid advise flags" },
        { 0x80040002u, "Can't enumerate any more, because the associated data is missing" },
        { 0x80040003u, "This implementation doesn't take advises" },
        { 0x80040004u, "There is no connection for this connection ID" },
        { 0x80040005u, "Need to run the object to perform this operation" },
        { 0x80040006u, "There is no cache to operate on" },
        { 0x80040007u, "Uninitialized object" }, { 0x80040008u, "Linked object's source class has changed" },
        { 0x80040009u, "Not able to get the moniker of the object" },
        { 0x8004000Au, "Not able to bind to the source" }, { 0x8004000Bu, "Object is static; operation not allowed" },
        { 0x8004000Cu, "User canceled out of save dialog" }, { 0x8004000Du, "Invalid rectangle" },
        { 0x8004000Eu, "compobj.dll is too old for the ole2.dll initialized" },
        { 0x8004000Fu, "Invalid window handle" }, { 0x80040010u, "Object is not in any of the inplace active states" },
        { 0x80040011u, "Not able to convert object" },
        { 0x80040012u, "Not able to perform the operation because object is not given storage yet" },
        { 0, 0 } };
    for (int i = 0; t[i].s; i++) if (t[i].h == h) return t[i].s;
    return 0;
}

/* Err.Raise / the Error statement. flags: 1 Source given, 2 Description given, 4 HelpContext given */
void vb_err_raise(vb_i32 number, vbs source, vbs description, vb_i32 ctx, vb_i32 flags)
{
    if (vb_errpend) return;
    vb_i32 m = vb_mark();
    vbs desc = description;
    if (number == 0 || number > 65535) {        /* not an error number: "Invalid procedure call" */
        number = 5; flags &= ~2;
    } else if (number < 0) {
        vb_i32 v = hresult_vb((vb_u32)number);
        if (v) number = v;
        else if (!(flags & 2)) {                /* "Automation error" and the system's text */
            const char *t = hresult_text((vb_u32)number);
            char buf[512];
            snprintf(buf, sizeof buf, t ? "Automation error\n%s " : "Automation error", t ? t : "");
            desc = vb_sfromutf8(buf); flags |= 2;
        }
    }
    vb_errpend = 1; pend_num = number; vb_fpstore = 0;
    vb_sset(&pend_src, (flags & 1) ? source : vb_stmp(0));
    vb_sset(&pend_desc, (flags & 2) ? desc : vb_sascii(errtext_pub(number)));
    pend_has_desc = 1;
    pend_ctx = (flags & 4) ? ctx : number < 0 ? 0 : errtab_pub(number) ? 1000000 + number : 1000095;
    pend_has_ctx = 1;
    vb_release(m);
}

/* an error whose text has a |1 in it (the name or the index VB6 puts there) */
void vb_raise_arg(vb_i32 number, const char *arg)
{
    if (vb_errpend) return;
    const char *t = vb_errtext(number), *p = strstr(t, "|1");
    char buf[1024];
    if (p) snprintf(buf, sizeof buf, "%.*s%s%s", (int)(p - t), t, arg ? arg : "", p + 2);
    else snprintf(buf, sizeof buf, "%s", t);
    vb_i32 m = vb_mark();
    vb_sset(&pend_desc, vb_sfromutf8(buf));
    vb_sfree(&pend_src);
    vb_release(m);
    vb_errpend = 1; pend_num = number; pend_has_desc = 1; vb_fpstore = 0;
}

void vb_set_error_text(vbs *dst, int n)
{
    const char *t = vb_errtext(n);
    vb_i32 m = vb_mark(); vb_sset(dst, vb_sascii(t)); vb_release(m);
}

vbs vb_error_str(vb_i32 n)
{
    if (n < -32768 || n > 65535) { vb_raise(6); return vb_sfromutf8(""); }    /* (an unsigned 16-bit number) */
    if (n == 0) return vb_sfromutf8("");
    return vb_sascii(errtext_pub(n));
}

/* the debugger: the error that is pending (not in Err yet) */
vb_i32 vb_err_number_pending(void) { return vb_errpend ? pend_num : 0; }

/* the pending error put aside while events run (ExecuteComplete of ADO), then back */
typedef struct { vb_i32 num; vbs src, desc; int has; } errstash;
void *vb_err_stash(void)
{
    if (!vb_errpend) return 0;
    errstash *e = (errstash *)calloc(1, sizeof *e);
    e->num = pend_num; e->has = pend_has_desc;
    vb_sset(&e->src, pend_src); vb_sset(&e->desc, pend_desc);
    vb_errpend = 0;
    return e;
}
void vb_err_unstash(void *p)
{
    errstash *e = (errstash *)p;
    if (!e) return;
    vb_errpend = 1; pend_num = e->num; pend_has_desc = e->has;
    vb_sset(&pend_src, e->src); vb_sset(&pend_desc, e->desc);
    vb_sfree(&e->src); vb_sfree(&e->desc);
    free(e);
}
void vb_err_clear_pending(void) { vb_errpend = 0; }
const char *vb_err_text_pending(void)
{
    if (!vb_errpend) return "";
    if (pend_has_desc && pend_desc.len) return vb_stoutf8(pend_desc);
    return vb_errtext(pend_num);
}

/* an error leaving a component for its client: its Source is the component (VB6: the project the
 * error happened in), unless the code gave one */
void vb_err_source_default(const char *project)
{
    if (!vb_errpend || (pend_has_desc && pend_src.len)) return;
    vb_i32 m = vb_mark();
    if (!pend_has_desc) vb_set_error_text(&pend_desc, pend_num);
    vb_sset(&pend_src, vb_sfromutf8(project));
    vb_release(m);
    pend_has_desc = 1;
}

void vb_err_take(vb_i32 erl)
{
    vb_errpend = 0;
    err_num = pend_num; err_erl = erl;
    if (pend_has_desc && pend_desc.len) vb_sset(&err_desc, pend_desc);
    else vb_set_error_text(&err_desc, err_num);
    if (pend_has_desc && pend_src.len) vb_sset(&err_src, pend_src);
    else if (vb_project_title) { vb_i32 m = vb_mark(); vb_sset(&err_src, vb_sfromutf8(vb_project_title)); vb_release(m); }
    else { vb_i32 m = vb_mark(); vb_sset(&err_src, vb_app_exename()); vb_release(m); }
    err_ctx = pend_has_ctx ? pend_ctx : help_context(err_num);
    pend_has_desc = 0; pend_has_ctx = 0;
}
vb_i32 vb_err_helpcontext(void) { return err_ctx; }
void vb_err_set_helpcontext(vb_i32 n) { err_ctx = n; }

/* the project of the program (Err.Source of its own errors: App.Title, as VB6) */
const char *vb_project_title;
void vb_set_project(const char *title) { vb_project_title = title; }

void vb_err_clear(void)
{
    err_num = 0; err_erl = 0; err_ctx = 0;
    vb_sfree(&err_desc); vb_sfree(&err_src);
}

vb_i32 vb_err_number(void) { return err_num; }
vbs vb_err_description(void) { vbs r = vb_stmp(err_desc.len); if (err_desc.len) memcpy(r.p, err_desc.p, 2 * err_desc.len); return r; }
vbs vb_err_source(void) { vbs r = vb_stmp(err_src.len); if (err_src.len) memcpy(r.p, err_src.p, 2 * err_src.len); return r; }
void vb_err_set_number(vb_i32 n) { err_num = n; }
void vb_err_set_description(vbs s) { vb_sset(&err_desc, s); }
void vb_err_set_source(vbs s) { vb_sset(&err_src, s); }
vb_i32 vb_erl(void) { return err_erl; }

/* the forms runtime: an error box like VB6 (a program started on its own), or the test robot,
 * which writes the error and goes on (returns 1) */
int (*vb_unhandled_hook)(vb_i32 number, const char *description);

/* Compile On Demand: a procedure with a compile error, called (VB6 compiles a procedure in the IDE when it is
 * first called; the error box "Compile error:" and the end of the run) */
int (*vb_compile_hook)(const char *message);

void vb_compile_error(const char *file, vb_i32 line, const char *message)
{
    vb_files_flush();
    fflush(stdout);
    fprintf(stderr, "%s:%d:1: error: %s\n", file ? file : "", (int)line, message ? message : "");
    fprintf(stderr, "Compile error:\n\n%s\n", message ? message : "");
    fflush(stderr);
    if (vb_compile_hook && vb_compile_hook(message ? message : "")) return;   /* the robot goes on */
    exit(1);
}

/* where the program is: Module.Procedure, line N (debug.c sets it in a --debug build) */
const char *(*vb_where_hook)(void);
const char *vb_where(void) { return vb_where_hook ? vb_where_hook() : ""; }

void vb_unhandled(void)
{
    vb_i32 m = vb_mark();
    const char *d = pend_has_desc && pend_desc.len ? vb_stoutf8(pend_desc) : vb_errtext(pend_num);
    vb_files_flush();
    fflush(stdout);
    const char *w = vb_where();
    fprintf(stderr, "Run-time error '%d':\n\n%s\n", pend_num, d);
    if (*w) fprintf(stderr, "at %s\n", w);
    if (vb_unhandled_hook && vb_unhandled_hook(pend_num, d)) {
        vb_release(m);
        vb_errpend = 0;
        return;
    }
    vb_release(m);
    exit(1);
}

/* ---------------------------------------------------------------- numbers */
static vb_i16 ck16(vb_i64 v) { if (v < -32768 || v > 32767) { vb_raise(6); return 0; } return (vb_i16)v; }
static vb_i32 ck32(vb_i64 v) { if (v < -2147483647LL - 1 || v > 2147483647LL) { vb_raise(6); return 0; } return (vb_i32)v; }
static vb_u8 ck8(vb_i64 v) { if (v < 0 || v > 255) { vb_raise(6); return 0; } return (vb_u8)v; }
vb_i16 vb_add_i(vb_i32 a, vb_i32 b) { return ck16((vb_i64)a + b); }
vb_i16 vb_sub_i(vb_i32 a, vb_i32 b) { return ck16((vb_i64)a - b); }
vb_i16 vb_mul_i(vb_i32 a, vb_i32 b) { return ck16((vb_i64)a * b); }
vb_i32 vb_add_l(vb_i64 a, vb_i64 b) { return ck32(a + b); }
vb_i32 vb_sub_l(vb_i64 a, vb_i64 b) { return ck32(a - b); }
vb_i32 vb_mul_l(vb_i64 a, vb_i64 b) { return ck32(a * b); }
vb_u8 vb_add_b(vb_i32 a, vb_i32 b) { return ck8((vb_i64)a + b); }
vb_u8 vb_sub_b(vb_i32 a, vb_i32 b) { return ck8((vb_i64)a - b); }
vb_u8 vb_mul_b(vb_i32 a, vb_i32 b) { return ck8((vb_i64)a * b); }
#define CYMAX 922337203685477.5807
vb_cy vb_add_c(vb_cy a, vb_cy b) { vb_cy r; if (__builtin_add_overflow(a, b, &r)) { vb_raise(6); return 0; } return r; }
vb_cy vb_sub_c(vb_cy a, vb_cy b) { vb_cy r; if (__builtin_sub_overflow(a, b, &r)) { vb_raise(6); return 0; } return r; }
vb_cy vb_mul_c(vb_cy a, vb_cy b)
{
    long double r = (long double)a * (long double)b / 10000.0L;
    if (r > 9.2233720368547758e18L || r < -9.2233720368547758e18L) { vb_raise(6); return 0; }
    return (vb_cy)vb_round_even((double)r);
}
double vb_chkd(double v) { if (isinf(v) || isnan(v)) { vb_raise(6); return 0; } return v; }
float vb_chkf(double v) { if (isinf(v) || isnan(v) || fabs(v) > 3.402823466e38) { vb_raise(6); return 0; } return (float)v; }
double vb_div(double a, double b)
{
    if (b == 0) { vb_raise(a == 0 ? 6 : 11); return 0; }
    return vb_chkd(a / b);
}
float vb_div_f(float a, float b)
{
    if (b == 0) { vb_raise(a == 0 ? 6 : 11); return 0; }
    return vb_chkf((double)a / b);
}
vb_i32 vb_idiv(vb_i32 a, vb_i32 b) { if (b == 0) { vb_raise(11); return 0; } if (a == (-2147483647 - 1) && b == -1) { vb_raise(6); return 0; } return a / b; }
vb_i32 vb_mod(vb_i32 a, vb_i32 b) { if (b == 0) { vb_raise(11); return 0; } if (b == -1) return 0; return a % b; }
double vb_pow(double a, double b)
{
    if (a == 0 && b < 0) { vb_raise(5); return 0; }
    if (a < 0 && b != floor(b)) { vb_raise(5); return 0; }
    return vb_chkd(pow(a, b));
}

double vb_round_even(double v)
{
    double f = floor(v), d = v - f;
    if (d > 0.5) return f + 1;
    if (d < 0.5) return f;
    return fmod(f, 2) == 0 ? f : f + 1;
}
vb_i16 vb_cint(double v) { if (isnan(v)) { vb_raise(6); return 0; } double r = vb_round_even(v); if (r < -32768 || r > 32767) { vb_raise(6); return 0; } return (vb_i16)r; }
vb_i32 vb_clng(double v) { if (isnan(v)) { vb_raise(6); return 0; } double r = vb_round_even(v); if (r < -2147483648.0 || r > 2147483647.0) { vb_raise(6); return 0; } return (vb_i32)r; }
vb_u8 vb_cbyte(double v) { double r = vb_round_even(v); if (!(r >= 0 && r <= 255)) { vb_raise(6); return 0; } return (vb_u8)r; }
vb_cy vb_ccur(double v)
{
    double r = vb_round_even(v * 10000.0);
    if (!(r >= -9.2233720368547758e18 && r < 9.2233720368547758e18)) { vb_raise(6); return 0; }
    return (vb_cy)r;
}
float vb_csng(double v) { return vb_chkf(v); }
double vb_cy2d(vb_cy c) { return (double)c / 10000.0; }
vb_bool vb_cbool(double v) { return v != 0 ? -1 : 0; }
double vb_fix(double v) { return v < 0 ? ceil(v) : floor(v); }
double vb_int(double v) { return floor(v); }
vb_i16 vb_abs_i(vb_i32 v) { return ck16(v < 0 ? -(vb_i64)v : v); }
vb_i32 vb_abs_l(vb_i64 v) { return ck32(v < 0 ? -v : v); }
double vb_round(double v, vb_i32 digits)
{
    if (digits < 0) { vb_raise(5); return 0; }
    double p = pow(10.0, digits);
    double x = v * p;
    /* round half to even on the decimal value (VB6 Round uses banker's rounding) */
    double r = vb_round_even(x);
    if (fabs(x - floor(x) - 0.5) > 1e-9) r = floor(x + 0.5);
    return r / p;
}

/* ---------------------------------------------------------------- regional settings
 * VB6 takes the decimal separator, the digit-grouping symbol and the leading-zero
 * choice from the user's regional settings (oleaut32, LOCALE_USER_DEFAULT).
 * Confirmed with VB6 on a Greek Windows 7: Print 0.5 -> "0,5", CStr(1.5) -> "1,5",
 * CDbl("3.25") -> 325 (the "." is the grouping symbol there), while Str, Val and
 * Write # always use ".".
 * VBRT_LCID picks the locale (1032 Greek, 1033 English US ...); without it the
 * POSIX locale (LC_ALL / LC_NUMERIC / LANG) decides. */
int vb_lcid = 1033;
char vb_lc_dec = '.', vb_lc_grp = ',';
int vb_lc_lzero = 1;

/* Short date, long time and AM/PM symbols as Windows 7 ships them. 1032 (Greek) is
 * confirmed by VB6 itself (t_dates); the others are the documented Windows defaults.
 * Long date, short time and currency (n = the number, $ = the symbol) for Format$. */
const char *vb_lc_sdate = "M/d/yyyy", *vb_lc_ltime = "h:mm:ss tt", *vb_lc_am = "AM", *vb_lc_pm = "PM";
const char *vb_lc_ldate = "dddd, MMMM d, yyyy", *vb_lc_stime = "h:mm tt";
const char *vb_lc_cursym = "$", *vb_lc_curpos = "$n", *vb_lc_curneg = "($n)";
int vb_lc_names = 0;                    /* 0 English, 1 Greek month and day names */
int vb_lc_firstday = 1;                 /* first day of the week: 1 Sunday, 2 Monday */

static const struct {
    int lcid; const char *tag; char dec, grp; int lzero;
    const char *sdate, *ltime, *am, *pm, *ldate, *stime, *cursym, *curpos, *curneg; int names;
} lc_table[] = {
    { 1033, "en_US", '.', ',', 1, "M/d/yyyy",   "h:mm:ss tt", "AM", "PM", "dddd, MMMM d, yyyy", "h:mm tt", "$", "$n", "($n)", 0 },
    { 2057, "en_GB", '.', ',', 1, "dd/MM/yyyy", "HH:mm:ss",   "AM", "PM", "dd MMMM yyyy", "HH:mm", "£", "$n", "-$n", 0 },
    { 1032, "el_GR", ',', '.', 1, "d/M/yyyy",   "h:mm:ss tt", "πμ", "μμ", "dddd, d MMMM yyyy", "h:mm tt", "€", "n $", "-n $", 1 },
    { 1031, "de_DE", ',', '.', 1, "dd.MM.yyyy", "HH:mm:ss",   "", "", "dddd, d. MMMM yyyy", "HH:mm", "€", "n $", "-n $", 0 },
    { 1036, "fr_FR", ',', ' ', 1, "dd/MM/yyyy", "HH:mm:ss",   "", "", "dddd d MMMM yyyy", "HH:mm", "€", "n $", "-n $", 0 },
    { 1040, "it_IT", ',', '.', 1, "dd/MM/yyyy", "HH:mm:ss",   "", "", "dddd d MMMM yyyy", "HH:mm", "€", "$ n", "-$ n", 0 },
    { 3082, "es_ES", ',', '.', 1, "dd/MM/yyyy", "H:mm:ss",    "", "", "dddd, d' de 'MMMM' de 'yyyy", "H:mm", "€", "n $", "-n $", 0 },
    { 1043, "nl_NL", ',', '.', 1, "d-M-yyyy",   "H:mm:ss",    "", "", "dddd d MMMM yyyy", "H:mm", "€", "$ n", "$ -n", 0 },
    { 2070, "pt_PT", ',', ' ', 1, "dd-MM-yyyy", "HH:mm:ss",   "", "", "dddd, d' de 'MMMM' de 'yyyy", "HH:mm", "€", "n $", "-n $", 0 },
    { 1046, "pt_BR", ',', '.', 1, "dd/MM/yyyy", "HH:mm:ss",   "", "", "dddd, d' de 'MMMM' de 'yyyy", "HH:mm", "R$", "$ n", "-$ n", 0 },
    { 1049, "ru_RU", ',', ' ', 1, "dd.MM.yyyy", "H:mm:ss",    "", "", "d MMMM yyyy 'г.'", "H:mm", "р.", "n$", "-n$", 0 },
    { 1045, "pl_PL", ',', ' ', 1, "yyyy-MM-dd", "HH:mm:ss",   "", "", "d MMMM yyyy", "HH:mm", "zł", "n $", "-n $", 0 },
    { 1055, "tr_TR", ',', '.', 1, "dd.MM.yyyy", "HH:mm:ss",   "ÖÖ", "ÖS", "dd MMMM yyyy dddd", "HH:mm", "TL", "n $", "-n $", 0 },
    { 1053, "sv_SE", ',', ' ', 1, "yyyy-MM-dd", "HH:mm:ss",   "", "", "'den 'd MMMM yyyy", "HH:mm", "kr", "n $", "-n $", 0 },
    { 1041, "ja_JP", '.', ',', 1, "yyyy/MM/dd", "H:mm:ss",    "午前", "午後", "yyyy'年'M'月'd'日'", "H:mm", "¥", "$n", "-$n", 0 },
    { 2052, "zh_CN", '.', ',', 1, "yyyy/M/d",   "H:mm:ss",    "上午", "下午", "yyyy'年'M'月'd'日'", "H:mm", "¥", "$n", "$-n", 0 },
    { 1042, "ko_KR", '.', ',', 1, "yyyy-MM-dd", "tt h:mm:ss", "오전", "오후", "yyyy'년' M'월' d'일' dddd", "tt h:mm", "₩", "$n", "-$n", 0 },
    { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 }
};

static int lc_apply(int lcid)
{
    for (int i = 0; lc_table[i].lcid; i++)
        if (lc_table[i].lcid == lcid) {
            vb_lcid = lcid; vb_lc_dec = lc_table[i].dec; vb_lc_grp = lc_table[i].grp; vb_lc_lzero = lc_table[i].lzero;
            vb_lc_sdate = lc_table[i].sdate; vb_lc_ltime = lc_table[i].ltime; vb_lc_am = lc_table[i].am; vb_lc_pm = lc_table[i].pm;
            vb_lc_ldate = lc_table[i].ldate; vb_lc_stime = lc_table[i].stime; vb_lc_cursym = lc_table[i].cursym;
            vb_lc_curpos = lc_table[i].curpos; vb_lc_curneg = lc_table[i].curneg; vb_lc_names = lc_table[i].names;
            vb_lc_firstday = (lcid == 1033 || lcid == 1041 || lcid == 2052 || lcid == 1042 || lcid == 1046) ? 1 : 2;
            return 1;
        }
    return 0;
}

/* the decimal point and the thousands separator of a language (StdDataFormat: the LCID of its Format); 0 when
 * the table does not have it */
int vb_lc_seps(int lcid, char *dec, char *grp)
{
    for (int i = 0; lc_table[i].lcid; i++)
        if (lc_table[i].lcid == lcid) { *dec = lc_table[i].dec; *grp = lc_table[i].grp; return 1; }
    return 0;
}

static void lc_init(void)
{
    const char *e = getenv("VBRT_LCID");
    if (e && *e && lc_apply(atoi(e))) return;
#ifdef _WIN32
    if (lc_apply(vbos_user_lcid())) return;      /* the regional settings of Windows */
#endif
    const char *l = getenv("LC_ALL"); if (!l || !*l) l = getenv("LC_NUMERIC"); if (!l || !*l) l = getenv("LANG");
    if (!l || !*l || !strcmp(l, "C") || !strcmp(l, "POSIX")) return;
    for (int i = 0; lc_table[i].lcid; i++)                      /* exact language_TERRITORY */
        if (!strncmp(l, lc_table[i].tag, 5)) { lc_apply(lc_table[i].lcid); return; }
    for (int i = 0; lc_table[i].lcid; i++)                      /* language only */
        if (!strncmp(l, lc_table[i].tag, 2)) { lc_apply(lc_table[i].lcid); return; }
}

/* ---------------------------------------------------------------- number -> text
 * digits: 7 for a Single, 15 for a Double (oleaut32 VarBstrFromR4/R8).
 * mode VB_FMT_STR: Str$, Write # - "." and no leading zero ("0.5" -> ".5").
 * mode VB_FMT_LOCALE: CStr, Print, & - regional separator and leading zero.
 * Exponent form: when the integer part needs more than `digits` digits, or when the
 * fixed form would need more than `digits` decimals (VB6: 2.5E-10 -> 0,00000000025,
 * 1E-20 -> 1E-20). The second boundary is still being checked with VB6 (t_numfmt). */
void vb_fmt_real(double v, int digits, int mode, char *out)
{
    if (v == 0) { strcpy(out, "0"); return; }
    char dec = mode == VB_FMT_LOCALE ? vb_lc_dec : '.';
    if (isnan(v) || isinf(v)) {                 /* 1,#INF with the regional separator (VB6, CStr) */
        strcpy(out, isnan(v) ? "-1.#IND" : v > 0 ? "1.#INF" : "-1.#INF");
        *strchr(out, '.') = dec;
        return;
    }
    int lead0 = mode == VB_FMT_LOCALE ? vb_lc_lzero : 0;
    char buf[64];
    int neg = v < 0; if (neg) v = -v;
    snprintf(buf, sizeof buf, "%.*e", digits - 1, v);
    char mant[40]; int nm = 0; char *p = buf;
    for (; *p && *p != 'e'; p++) if (*p >= '0' && *p <= '9') mant[nm++] = *p;
    int e = atoi(p + 1);
    while (nm > 1 && mant[nm - 1] == '0') nm--;
    mant[nm] = 0;
    char *o = out;
    if (neg) *o++ = '-';
    if (e < digits && (e >= 0 || nm - 1 - e <= digits)) {
        if (e < 0) {
            if (lead0) *o++ = '0';
            *o++ = dec;
            for (int i = 0; i < -e - 1; i++) *o++ = '0';
            for (int i = 0; i < nm; i++) *o++ = mant[i];
        } else {
            for (int i = 0; i <= e; i++) *o++ = i < nm ? mant[i] : '0';
            if (nm > e + 1) { *o++ = dec; for (int i = e + 1; i < nm; i++) *o++ = mant[i]; }
        }
        *o = 0;
    } else {
        *o++ = mant[0];
        if (nm > 1) { *o++ = dec; for (int i = 1; i < nm; i++) *o++ = mant[i]; }
        *o++ = 'E'; *o++ = e < 0 ? '-' : '+';
        int ae = e < 0 ? -e : e;
        o += sprintf(o, "%02d", ae);
        *o = 0;
    }
}

void vb_fmt_cy(vb_cy c, int mode, char *out)
{
    char *o = out;
    char dec = mode == VB_FMT_LOCALE ? vb_lc_dec : '.';
    int lead0 = mode == VB_FMT_LOCALE ? vb_lc_lzero : 0;
    unsigned long long u = c < 0 ? (unsigned long long)(-(c + 1)) + 1 : (unsigned long long)c;
    if (c < 0) *o++ = '-';
    unsigned long long ip = u / 10000, fp = u % 10000;
    if (ip || lead0 || !fp) o += sprintf(o, "%llu", ip);
    if (fp) { sprintf(o, "%c%04llu", dec, fp); size_t n = strlen(o); while (o[n - 1] == '0') o[--n] = 0; }
}

vbs vb_cstr_i(vb_i32 v) { char b[16]; snprintf(b, sizeof b, "%d", v); return vb_sascii(b); }
vbs vb_cstr_f(float v) { char b[48]; vb_fmt_real(v, 7, VB_FMT_LOCALE, b); return vb_sascii(b); }
vbs vb_cstr_d(double v) { char b[48]; vb_fmt_real(v, 15, VB_FMT_LOCALE, b); return vb_sascii(b); }
vbs vb_cstr_c(vb_cy v) { char b[48]; vb_fmt_cy(v, VB_FMT_LOCALE, b); return vb_sascii(b); }
vbs vb_cstr_b(vb_bool v) { return vb_sascii(v ? "True" : "False"); }
vbs vb_str_i(vb_i32 v) { char b[16]; snprintf(b, sizeof b, v < 0 ? "%d" : " %d", v); return vb_sascii(b); }
vbs vb_str_f(float v) { char b[48]; b[0] = ' '; vb_fmt_real(v, 7, VB_FMT_STR, b + 1); return vb_sascii(b[1] == '-' ? b + 1 : b); }
vbs vb_str_d(double v) { char b[48]; b[0] = ' '; vb_fmt_real(v, 15, VB_FMT_STR, b + 1); return vb_sascii(b[1] == '-' ? b + 1 : b); }
vbs vb_str_c(vb_cy v) { char b[48]; b[0] = ' '; vb_fmt_cy(v, VB_FMT_STR, b + 1); return vb_sascii(b[1] == '-' ? b + 1 : b); }

/* ---------------------------------------------------------------- text -> number */
static int parse_num(vbs s, double *out, int strict)
{
    vb_i32 i = 0, n = s.len + ((s.flags & VBS_ODD) ? 1 : 0);   /* (Val(byte string): the last byte too) */
    char buf[400]; int k = 0;
    while (i < n && (s.p[i] == ' ' || s.p[i] == '\t')) i++;
    if (i + 1 < n && s.p[i] == '&' && (s.p[i + 1] == 'H' || s.p[i + 1] == 'h' || s.p[i + 1] == 'O' || s.p[i + 1] == 'o')) {
        int hex = s.p[i + 1] == 'H' || s.p[i + 1] == 'h'; i += 2;
        long long v = 0; int any = 0;
        for (; i < n; i++) {
            int c = s.p[i], d;
            if (c >= '0' && c <= '9') d = c - '0';
            else if (hex && c >= 'a' && c <= 'f') d = c - 'a' + 10;
            else if (hex && c >= 'A' && c <= 'F') d = c - 'A' + 10;
            else break;
            if (!hex && d > 7) break;
            v = v * (hex ? 16 : 8) + d; any = 1;
        }
        if (v > 0xFFFF && v <= 0xFFFFFFFFLL) v = (vb_i32)(vb_u32)v; else if (v >= 0x8000 && v <= 0xFFFF) v = (vb_i16)(vb_u16)v;
        *out = (double)v;
        if (strict) { while (i < n && s.p[i] == ' ') i++; return any && i == n; }
        return 1;
    }
    for (; i < n && k < 390; i++) {
        vb_wchar c = s.p[i];
        if (c == ' ' || c == '\t') { if (strict) { if (k) break; else continue; } else continue; }  /* Val ignores blanks */
        if (strict && c == (vb_wchar)(unsigned char)vb_lc_dec) { buf[k++] = '.'; continue; }   /* CDbl & co: regional settings */
        if (strict && c == (vb_wchar)(unsigned char)vb_lc_grp && k) continue;
        if (strict && vb_lc_grp == ' ' && c == 0xA0 && k) continue;
        if ((c >= '0' && c <= '9') || (!strict && c == '.') || c == '+' || c == '-' || c == 'e' || c == 'E' || c == 'd' || c == 'D')
            buf[k++] = (c == 'd' || c == 'D') ? 'E' : (char)c;
        else break;
    }
    buf[k] = 0;
    char *end;
    double v = k ? strtod(buf, &end) : 0;
    if (!k) end = buf;
    *out = v;
    if (strict) {
        if (end == buf || *end) return 0;
        while (i < n && (s.p[i] == ' ' || s.p[i] == '\t')) i++;
        return i == n;
    }
    return 1;
}

double vb_val(vbs s) { double v; parse_num(s, &v, 0); return v; }

double vb_s2d(vbs s)
{
    double v;
    if (!parse_num(s, &v, 1)) {
        vb_i32 m = vb_mark();
        vbs t = vb_lcase(vb_trim(s));
        vb_i32 ok = 0;
        if (t.len == 4 && t.p[0] == 't' && t.p[1] == 'r' && t.p[2] == 'u' && t.p[3] == 'e') { v = -1; ok = 1; }
        if (t.len == 5 && t.p[0] == 'f' && t.p[1] == 'a' && t.p[2] == 'l' && t.p[3] == 's' && t.p[4] == 'e') { v = 0; ok = 1; }
        vb_release(m);
        if (!ok) { vb_raise(13); return 0; }
    }
    return v;
}

vb_bool vb_s2bool(vbs s) { double v = vb_s2d(s); return v != 0 ? -1 : 0; }

/* ---------------------------------------------------------------- program */
static int g_argc; static char **g_argv;
/* Out of stack space (28): each procedure compares its frame with this limit, set below the stack of
 * vb_init with room left for the runtime (VB6 runs out at 1 MB too; Linux has 8 MB) */
char *vb_stack_lim;
static void stack_init(void)
{
    char here;
    size_t size = 1024 * 1024;                  /* Windows: the 1 MB of an .exe */
#ifndef _WIN32
    struct rlimit rl;
    if (getrlimit(RLIMIT_STACK, &rl) == 0 && rl.rlim_cur != RLIM_INFINITY) size = (size_t)rl.rlim_cur;
    else size = 8 * 1024 * 1024;
    if (size > 64 * 1024 * 1024) size = 64 * 1024 * 1024;
#endif
    size_t room = size / 4 > 256 * 1024 ? 256 * 1024 : size / 4;
    uintptr_t top = (uintptr_t)&here;
    vb_stack_lim = (char *)(top - (size - room - 64 * 1024));
}
void vb_init(int argc, char **argv)
{
    g_argc = argc; g_argv = argv;
    stack_init();
    lc_init();
    { int c = vb_lcid_cp(vb_lcid); if (c > 0) vb_codepage = c; }   /* the code page of the locale */
    const char *cp = getenv("VBRT_CODEPAGE");
    if (cp) vb_codepage = atoi(cp);
#ifdef _WIN32
    else {
        vb_codepage = vbos_ansi_codepage();     /* the ANSI code page of Windows (1253 on a Greek system) */
        vbos_console_utf8();
    }
#else
    else {
        const char *l = getenv("LC_ALL"); if (!l || !*l) l = getenv("LC_CTYPE"); if (!l || !*l) l = getenv("LANG");
        if (l && strncmp(l, "el", 2) == 0) vb_codepage = 1253;
    }
#endif
    setvbuf(stdout, 0, _IOLBF, 0);
}
const char *vb_argv0(void) { return g_argc > 0 ? g_argv[0] : "program"; }
vbs vb_command(void)
{
    char buf[4096]; size_t k = 0; buf[0] = 0;
    for (int i = 1; i < g_argc && k < sizeof buf - 2; i++) {
        if (i > 1) buf[k++] = ' ';
        size_t n = strlen(g_argv[i]); if (k + n >= sizeof buf - 1) n = sizeof buf - 1 - k;
        memcpy(buf + k, g_argv[i], n); k += n;
    }
    buf[k] = 0;
    return vb_sfromutf8(buf);
}
void (*vb_end_hook)(void) = 0;
void vb_end(void)
{
    if (vb_end_hook) { void (*h)(void) = vb_end_hook; vb_end_hook = 0; h(); }
    vb_files_flush(); fflush(stdout); exit(0);
}
void vb_stop(void) { vb_files_flush(); fflush(stdout); fprintf(stderr, "Stop statement encountered\n"); exit(0); }

/* ---------------------------------------------------------------- String * n
 * A fixed-length string always holds n characters; a new one is n Chr(0). */
void vb_sfix_init(vbs *v, vb_i32 n)
{
    vbs t = vb_stmp(n);
    for (vb_i32 i = 0; i < n; i++) t.p[i] = 0;
    vb_sset(v, t);
}

void vb_afix_init(vbarr *a, vb_i32 n)
{
    for (vb_i32 i = 0; i < a->n; i++) {
        vbs *e = (vbs *)(a->data + (size_t)i * (size_t)a->esz);
        if (e->len == n) continue;
        if (e->len == 0) vb_sfix_init(e, n);
        else vb_sset_fix(e, *e, n);
    }
}

void vb_sset_fix(vbs *dst, vbs src, vb_i32 n)
{
    vbs t = vb_stmp(n);
    for (vb_i32 i = 0; i < n; i++) t.p[i] = i < src.len ? src.p[i] : ' ';
    vb_sset(dst, t);
}

/* App.Title and App.HelpFile changed by the program */
static vbs app_strs[3];
static int app_set[3];
vbs vb_app_str(int which, vbs dflt)
{
    if (which < 1 || which > 2 || !app_set[which]) return dflt;
    return vb_scopy(app_strs[which]);
}
void vb_app_set_str(int which, vbs v)
{
    if (which < 1 || which > 2) return;
    vb_sset(&app_strs[which], v);
    app_set[which] = 1;
}

/* App as an object (ShowAbout App, Me): what the .vbp says, set by the program at its start */
static struct { const char *key; vbs val; } app_info[24];
static int app_ninfo;
void vb_app_info(const char *key, vbs v)
{
    for (int i = 0; i < app_ninfo; i++)
        if (!strcmp(app_info[i].key, key)) { vb_sset(&app_info[i].val, v); return; }
    if (app_ninfo >= 24) return;
    app_info[app_ninfo].key = key;
    app_info[app_ninfo].val = (vbs){0, 0, 0};
    vb_sset(&app_info[app_ninfo].val, v);
    app_ninfo++;
}
static vbs app_get(const char *key)
{
    for (int i = 0; i < app_ninfo; i++)
        if (!strcmp(app_info[i].key, key)) return vb_scopy(app_info[i].val);
    return vb_stmp(0);
}
static vbvar app_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    char n[40]; int i = 0;
    (void)o;
    for (; name[i] && i < 39; i++) n[i] = (char)tolower((unsigned char)name[i]);
    n[i] = 0;
    int set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (!strcmp(n, "title") || !strcmp(n, "helpfile")) {
        int w = n[0] == 't' ? 1 : 2;
        if (set) { vb_app_set_str(w, vb_v2s(val)); return vb_v_empty(); }
        vbs d = app_get(n);
        if (w == 1 && !d.len) d = vb_app_exename();
        return vb_v_s(vb_app_str(w, d));
    }
    if (set && (!strcmp(n, "taskvisible") || !strcmp(n, "oleserverbusytimeout") || !strcmp(n, "olerequestpendingtimeout")
                || !strcmp(n, "oleserverbusyraiseerror") || !strcmp(n, "olerequestpendingmsgtext")
                || !strcmp(n, "olerequestpendingmsgtitle") || !strcmp(n, "oleserverbusymsgtext")
                || !strcmp(n, "oleserverbusymsgtitle"))) return vb_v_empty();
    if (set) { vb_raise(!strcmp(n, "path") || !strcmp(n, "exename") ? 383 : 438); return vb_v_empty(); }
    if (!strcmp(n, "path")) return vb_v_s(vb_app_path());
    if (!strcmp(n, "exename")) return vb_v_s(vb_app_exename());
    if (!strcmp(n, "major") || !strcmp(n, "minor") || !strcmp(n, "revision")) {
        vbs v = app_get(n);
        return vb_v_i((vb_i16)(v.len ? vb_val(v) : n[1] == 'a' ? 1 : 0));
    }
    if (!strcmp(n, "comments") || !strcmp(n, "companyname") || !strcmp(n, "filedescription") || !strcmp(n, "legalcopyright")
        || !strcmp(n, "legaltrademarks") || !strcmp(n, "productname")) return vb_v_s(app_get(n));
    if (!strcmp(n, "previnstance") || !strcmp(n, "unattendedapp")) return vb_v_bool(0);
    if (!strcmp(n, "nonmodalallowed") || !strcmp(n, "taskvisible")) return vb_v_bool(-1);
    if (!strcmp(n, "startmode")) return vb_v_i(0);
    if (!strcmp(n, "threadid")) return vb_v_l(1);
    if (!strcmp(n, "hinstance")) return vb_v_l(0x400000);
    if (!strcmp(n, "logmode")) return vb_v_l(0);
    if (!strcmp(n, "logpath")) return vb_v_s(vb_stmp(0));
    if (!strcmp(n, "logevent") || !strcmp(n, "startlogging")) return vb_v_empty();
    *found = 0;
    return vb_v_empty();
}
static const vb_class vb_cls_app = { "App", sizeof(vbobjh), 0, 0, 0, 0, 0, 0, 0, 0, 0, app_member };
vbobj vb_app_obj(void)
{
    static vbobj app;
    if (!app) { app = vb_onew(&vb_cls_app); vb_oaddref(app); }
    return app;
}

/* Err as an object (ShowErrorMsg Err, with errError As ErrObject) */
static int arg_given(vb_i32 argc, vbvar *argv, int i) { return i < argc && argv[i].vt != VT_EMPTY && !vb_ismissing(argv[i]); }
static vbvar errobj_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    char n[40]; int i = 0;
    (void)o;
    for (; name[i] && i < 39; i++) n[i] = (char)tolower((unsigned char)name[i]);
    n[i] = 0;
    int set = kind == VBM_LET || kind == VBM_SET;
    vbvar val = set && argc ? argv[argc - 1] : vb_v_empty();
    *found = 1;
    if (!*n || !strcmp(n, "number")) {
        if (set) { vb_err_set_number(vb_v2l(val)); return vb_v_empty(); }
        return vb_v_l(vb_err_number());
    }
    if (!strcmp(n, "description")) {
        if (set) { vb_err_set_description(vb_v2s(val)); return vb_v_empty(); }
        return vb_v_s(vb_err_description());
    }
    if (!strcmp(n, "source")) {
        if (set) { vb_err_set_source(vb_v2s(val)); return vb_v_empty(); }
        return vb_v_s(vb_err_source());
    }
    if (!strcmp(n, "helpcontext")) {
        if (set) { vb_err_set_helpcontext(vb_v2l(val)); return vb_v_empty(); }
        return vb_v_l(vb_err_helpcontext());
    }
    if (!strcmp(n, "helpfile")) return set ? vb_v_empty() : vb_v_s(vb_stmp(0));
    if (!strcmp(n, "lastdllerror")) return vb_v_l(vb_lastdllerr);
    if (!strcmp(n, "clear")) { vb_err_clear(); vb_lastdllerr = 0; return vb_v_empty(); }
    if (!strcmp(n, "raise")) {
        if (argc < 1) { vb_raise(449); return vb_v_empty(); }
        int fl = (arg_given(argc, argv, 1) ? 1 : 0) | (arg_given(argc, argv, 2) ? 2 : 0) | (arg_given(argc, argv, 4) ? 4 : 0);
        vb_err_raise(vb_v2l(argv[0]), arg_given(argc, argv, 1) ? vb_v2s(argv[1]) : vb_stmp(0),
                     arg_given(argc, argv, 2) ? vb_v2s(argv[2]) : vb_stmp(0),
                     arg_given(argc, argv, 4) ? vb_v2l(argv[4]) : 0, fl);
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}
static const vb_class vb_cls_errobj = { "ErrObject", sizeof(vbobjh), 0, 0, 0, 0, 0, "number", 0, 0, 0, errobj_member };
vbobj vb_err_obj(void)
{
    static vbobj e;
    if (!e) { e = vb_onew(&vb_cls_errobj); vb_oaddref(e); }
    return e;
}

/* VB.Global as an object (Set g = VB.Global: g.LoadResData ...): the functions of the VB library */
static vbvar global_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    char n[40]; int i = 0;
    (void)o; (void)kind;
    for (; name[i] && i < 39; i++) n[i] = (char)tolower((unsigned char)name[i]);
    n[i] = 0;
    *found = 1;
    if (!strcmp(n, "app")) return vb_v_obj(vb_app_obj());
    if (!strcmp(n, "loadresstring")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        return vb_v_s(vb_loadresstring(vb_v2l(argv[0])));
    }
    if (!strcmp(n, "loadresdata")) {
        if (argc < 2) { vb_raise(450); return vb_v_empty(); }
        return vb_loadresdata(argv[0], argv[1]);
    }
    if (!strcmp(n, "loadrespicture")) {
        if (argc < 2) { vb_raise(450); return vb_v_empty(); }
        return vb_v_obj(vb_loadrespicture(argv[0], vb_v2l(argv[1])));
    }
    if (!strcmp(n, "loadpicture")) {
        if (argc < 1 || argv[0].vt == VT_EMPTY || vb_ismissing(argv[0])) return vb_v_obj(vb_loadpicture(vb_stmp(0), 0));
        return vb_v_obj(vb_loadpicture(vb_v2s(argv[0]), 1));
    }
    *found = 0;
    return vb_v_empty();
}
static const vb_class vb_cls_global = { "Unknown", sizeof(vbobjh), 0, 0, 0, 0, 0, 0, 0, 0, 0, global_member };
vbobj vb_global_obj(void)
{
    static vbobj g;
    if (!g) { g = vb_onew(&vb_cls_global); vb_oaddref(g); }
    return g;
}

/* the database runtime (libvbdb): filled by its vb_db_init() */
vb_db_hooks vb_dbhooks;
/* CallByName obj, "Name", vbMethod (1) / vbGet (2) / vbLet (4) / vbSet (8), args */
vbvar vb_callbyname(vbobj target, vbs name, vb_i32 calltype, vb_i32 argc, vbvar *argv)
{
    char *n = vb_stoutf8(name);
    char low[256]; int i = 0;
    for (; n[i] && i < 255; i++) low[i] = (char)((n[i] >= 'A' && n[i] <= 'Z') ? n[i] + 32 : n[i]);
    low[i] = 0;
    vb_i32 kind = calltype == 2 ? VBM_GET : calltype == 4 ? VBM_LET : calltype == 8 ? VBM_SET : VBM_METHOD;
    if (calltype != 1 && calltype != 2 && calltype != 4 && calltype != 8) { vb_raise(5); return vb_v_empty(); }
    return vb_late_o(target, low, kind, argc, argv);
}

void vb_prop_changed(vbobj o, vbs name)
{
    if (vb_dbhooks.prop_changed) vb_dbhooks.prop_changed(o, vb_stoutf8(name));
}
