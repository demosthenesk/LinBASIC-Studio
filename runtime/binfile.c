/* binfile.c - Get and Put for files opened For Binary or For Random.
 *
 * The layout is the one of VB6 (little endian, no alignment):
 *   Integer, Boolean 2 bytes; Long, Single 4; Double, Currency, Date 8; Byte 1;
 *   String * n: n bytes (ANSI); a String: its bytes in Binary, but with a 2-byte length
 *   before them in Random files and inside a Type; a Variant: its 2-byte VarType, then
 *   the value (a String with its 2-byte length); a Type: its fields one after the other;
 *   a fixed array: its elements; a dynamic array inside a Type or in a Random file: a
 *   descriptor (2-byte number of dimensions, then per dimension 4-byte count and 4-byte
 *   lower bound) before the elements.
 * Random files: rec is a record number, a record is reclen bytes, a Put longer than
 * reclen is error 59, and a Put writes only its data (the file can end inside a record). Binary files: rec is a byte position (1 = first byte).
 *
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"

/* ---------------------------------------------------------------- byte buffer */
typedef struct { unsigned char *b; long n, cap, pos; int past_end; } buf;

static void put_bytes(buf *w, const void *p, long n)
{
    if (w->n + n > w->cap) {
        long nc = (w->cap ? w->cap * 2 : 64);
        while (nc < w->n + n) nc *= 2;
        unsigned char *nb = (unsigned char *)realloc(w->b, (size_t)nc);
        if (!nb) { vb_raise(7); return; }
        w->b = nb; w->cap = nc;
    }
    memcpy(w->b + w->n, p, (size_t)n);
    w->n += n;
}
static void put_u16(buf *w, unsigned v) { unsigned char b[2] = { (unsigned char)v, (unsigned char)(v >> 8) }; put_bytes(w, b, 2); }
static void put_u32(buf *w, unsigned long v) { unsigned char b[4] = { (unsigned char)v, (unsigned char)(v >> 8), (unsigned char)(v >> 16), (unsigned char)(v >> 24) }; put_bytes(w, b, 4); }

static void get_bytes(buf *r, void *p, long n)
{
    long have = r->n - r->pos;
    if (have < 0) have = 0;
    long k = n < have ? n : have;
    if (k > 0) memcpy(p, r->b + r->pos, (size_t)k);
    if (k < n) { memset((char *)p + k, 0, (size_t)(n - k)); r->past_end = 1; }
    r->pos += n;
}
static unsigned get_u16(buf *r) { unsigned char b[2]; get_bytes(r, b, 2); return b[0] | (b[1] << 8); }
static unsigned long get_u32(buf *r) { unsigned char b[4]; get_bytes(r, b, 4); return b[0] | (b[1] << 8) | ((unsigned long)b[2] << 16) | ((unsigned long)b[3] << 24); }

/* ---------------------------------------------------------------- values */
static vb_i32 scalar_size(vb_i32 kind)
{
    switch (kind) {
    case VBK_I2: case VBK_BOOL: return 2;
    case VBK_I4: case VBK_R4: return 4;
    case VBK_R8: case VBK_CY: case VBK_DATE: return 8;
    case VBK_U8: return 1;
    }
    return 0;
}

static void w_str(buf *w, vbs s, int prefix)
{
    int n; char *a = vb_stoansi(s, &n);
    if (prefix) put_u16(w, (unsigned)n);
    put_bytes(w, a, n);
}

static void w_value(buf *w, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, const void *p, int str_prefix);

static void w_variant(buf *w, const vbvar *v)
{
    put_u16(w, v->vt);
    switch (v->vt) {
    case VT_EMPTY: case VT_NULL: break;
    case VT_I2: case VT_BOOL: put_bytes(w, &v->u.i, 2); break;
    case VT_I4: case VT_ERROR: put_bytes(w, &v->u.l, 4); break;
    case VT_R4: put_bytes(w, &v->u.f, 4); break;
    case VT_R8: case VT_DATE: put_bytes(w, &v->u.d, 8); break;
    case VT_CY: put_bytes(w, &v->u.cy, 8); break;
    case VT_DECIMAL: { unsigned char d[16]; vb_dec_bytes(*v, d); put_bytes(w, d, 16); break; }   /* the whole DECIMAL */
    case VT_UI1: put_bytes(w, &v->u.b, 1); break;
    case VT_BSTR: w_str(w, v->u.s, 1); break;
    default: vb_raise(13);                      /* objects cannot be written */
    }
}

static void w_array(buf *w, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, const vbarr *a, int descriptor)
{
    if (descriptor) {
        put_u16(w, (unsigned)a->nd);
        for (int d = 0; d < a->nd; d++) { put_u32(w, (unsigned long)(a->ub[d] - a->lb[d] + 1)); put_u32(w, (unsigned long)a->lb[d]); }
    }
    for (vb_i32 i = 0; i < a->n && !vb_errpend; i++)
        w_value(w, kind, fixlen, udt, a->data + (vb_i64)i * a->esz, 1);
}

static void w_value(buf *w, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, const void *p, int str_prefix)
{
    switch (kind) {
    case VBK_STR: w_str(w, *(const vbs *)p, str_prefix); break;
    case VBK_FSTR: {
        vbs s = *(const vbs *)p;
        int n; char *a = vb_stoansi(s, &n);
        for (vb_i32 i = 0; i < fixlen; i++) { char c = i < n ? a[i] : ' '; put_bytes(w, &c, 1); }
        break;
    }
    case VBK_VAR: w_variant(w, (const vbvar *)p); break;
    case VBK_UDT:
        for (vb_i32 i = 0; i < udt->nfields && !vb_errpend; i++) {
            const vb_ufield *f = &udt->fields[i];
            const char *fp = (const char *)p + f->offset;
            if (f->arr) w_array(w, f->kind, f->fixlen, f->udt, (const vbarr *)fp, f->arr == 2);
            else w_value(w, f->kind, f->fixlen, f->udt, fp, 1);
        }
        break;
    case VBK_OBJ: vb_raise(13); break;
    default: put_bytes(w, p, scalar_size(kind));
    }
}

static vbs r_str(buf *r, long n)
{
    char *a = (char *)malloc((size_t)(n > 0 ? n : 1));
    if (!a) { vb_raise(7); return vb_stmp(0); }
    get_bytes(r, a, n);
    vbs s = vb_sfromansi(a, (int)n);
    free(a);
    return s;
}

static void r_value(buf *r, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, void *p, int str_prefix);

static void r_variant(buf *r, vbvar *v)
{
    vbvar t = vb_v_empty();
    t.vt = (vb_u16)get_u16(r);
    switch (t.vt) {
    case VT_EMPTY: case VT_NULL: break;
    case VT_I2: case VT_BOOL: get_bytes(r, &t.u.i, 2); break;
    case VT_I4: case VT_ERROR: get_bytes(r, &t.u.l, 4); break;
    case VT_R4: get_bytes(r, &t.u.f, 4); break;
    case VT_R8: case VT_DATE: get_bytes(r, &t.u.d, 8); break;
    case VT_CY: get_bytes(r, &t.u.cy, 8); break;
    case VT_DECIMAL: { unsigned char d[16]; get_bytes(r, d, 16); t = vb_dec_from_bytes(d); break; }
    case VT_UI1: get_bytes(r, &t.u.b, 1); break;
    case VT_BSTR: t.u.s = r_str(r, (long)get_u16(r)); break;
    default: if (!r->past_end) vb_raise(458); return;   /* Variant uses a type not supported */
    }
    vb_vset(v, t);
}

static void r_array(buf *r, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, vbarr *a, int descriptor)
{
    if (descriptor) {
        vb_i32 nd = (vb_i32)get_u16(r), lb[8], ub[8];
        if (nd > 8) { vb_raise(9); return; }
        for (int d = 0; d < nd; d++) {
            vb_i32 cnt = (vb_i32)get_u32(r); lb[d] = (vb_i32)get_u32(r); ub[d] = lb[d] + cnt - 1;
        }
        if (nd == 0) { vb_afree(a); return; }
        vb_i32 esz = a->esz ? a->esz : (kind == VBK_UDT ? udt->size : kind == VBK_STR || kind == VBK_FSTR ? (vb_i32)sizeof(vbs)
                                         : kind == VBK_VAR ? (vb_i32)sizeof(vbvar) : scalar_size(kind));
        vb_i32 ak = kind == VBK_STR || kind == VBK_FSTR ? VBA_STR : kind == VBK_VAR ? VBA_VAR : kind == VBK_UDT ? VBA_UDT : VBA_PLAIN;
        if (kind == VBK_UDT) a->udt = udt;
        switch (nd) {                           /* re-dimension to what the file says */
        case 1: vb_aredim(a, esz, ak, 0, 1, lb[0], ub[0]); break;
        case 2: vb_aredim(a, esz, ak, 0, 2, lb[0], ub[0], lb[1], ub[1]); break;
        case 3: vb_aredim(a, esz, ak, 0, 3, lb[0], ub[0], lb[1], ub[1], lb[2], ub[2]); break;
        default: vb_raise(9); return;
        }
    }
    for (vb_i32 i = 0; i < a->n && !vb_errpend; i++)
        r_value(r, kind, fixlen, udt, a->data + (vb_i64)i * a->esz, 1);
}

static void r_value(buf *r, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, void *p, int str_prefix)
{
    switch (kind) {
    case VBK_STR:
        if (str_prefix) vb_sset((vbs *)p, r_str(r, (long)get_u16(r)));
        else vb_sset((vbs *)p, r_str(r, ((vbs *)p)->len));    /* Binary: as many bytes as the string has */
        break;
    case VBK_FSTR: vb_sset((vbs *)p, r_str(r, fixlen)); break;
    case VBK_VAR: r_variant(r, (vbvar *)p); break;
    case VBK_UDT:
        for (vb_i32 i = 0; i < udt->nfields && !vb_errpend; i++) {
            const vb_ufield *f = &udt->fields[i];
            char *fp = (char *)p + f->offset;
            if (f->arr) r_array(r, f->kind, f->fixlen, f->udt, (vbarr *)fp, f->arr == 2);
            else r_value(r, f->kind, f->fixlen, f->udt, fp, 1);
        }
        break;
    case VBK_OBJ: vb_raise(13); break;
    default: get_bytes(r, p, scalar_size(kind));
    }
}

/* ---------------------------------------------------------------- file positions */
/* the record number written in Put/Get (0 there is "Bad record number"; an omitted one is 0) */
vb_i32 vb_recno(vb_i32 rec) { return rec == 0 ? -1 : rec; }
static vbfile *binfile(vb_i32 fileno)
{
    vbfile *f = vb_getfile(fileno, 0);
    if (f && !(f->mode & (VBF_BINARY | VBF_RANDOM))) { vb_raise(54); return 0; }   /* Bad file mode */
    return f;
}

static int position(vbfile *f, vb_i32 rec)
{
    if (rec > 0) {
        long off = f->mode & VBF_RANDOM ? (long)(rec - 1) * f->reclen : (long)rec - 1;
        if (fseek(f->f, off, SEEK_SET) != 0) { vb_raise(63); return 0; }
    } else if (rec < 0 || (rec == 0 && 0)) {
        vb_raise(63); return 0;                 /* Bad record number */
    }
    return 1;
}

static void after_record(vbfile *f, long start)
{
    if (f->mode & VBF_RANDOM) {
        f->lastrec = (vb_i32)(start / f->reclen) + 1;
        fseek(f->f, (long)f->lastrec * f->reclen, SEEK_SET);
    }
}

static void write_buf(vbfile *f, buf *w)
{
    long start = ftell(f->f);
    if (f->mode & VBF_RANDOM) {
        if (w->n > f->reclen) { vb_raise(59); return; }       /* Bad record length */
        /* only the data is written, not the rest of the record (VB6: LOF 38, not 48) */
        start -= start % f->reclen;
        fseek(f->f, start, SEEK_SET);
    }
    if (w->n && fwrite(w->b, 1, (size_t)w->n, f->f) != (size_t)w->n) { vb_raise(61); return; }
    fflush(f->f);
    f->eofflag = 0;
    after_record(f, start);
}

void vb_put(vb_i32 fileno, vb_i32 rec, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, const void *p)
{
    vbfile *f = binfile(fileno);
    if (!f || !position(f, rec)) return;
    buf w = { 0, 0, 0, 0, 0 };
    w_value(&w, kind, fixlen, udt, p, (f->mode & VBF_RANDOM) != 0);
    if (!vb_errpend) write_buf(f, &w);
    free(w.b);
}

void vb_put_arr(vb_i32 fileno, vb_i32 rec, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, vbarr *a, vb_i32 dynamic)
{
    vbfile *f = binfile(fileno);
    if (!f || !position(f, rec)) return;
    buf w = { 0, 0, 0, 0, 0 };
    w_array(&w, kind, fixlen, udt, a, dynamic && (f->mode & VBF_RANDOM));
    if (!vb_errpend) write_buf(f, &w);
    free(w.b);
}

/* read what the item needs: a whole record (Random), or the rest of the file (Binary) */
static int read_buf(vbfile *f, buf *r, long *start)
{
    *start = ftell(f->f);
    long n;
    if (f->mode & VBF_RANDOM) { *start -= *start % f->reclen; fseek(f->f, *start, SEEK_SET); n = f->reclen; }
    else { long end; fseek(f->f, 0, SEEK_END); end = ftell(f->f); fseek(f->f, *start, SEEK_SET); n = end - *start; }
    if (n < 0) n = 0;
    r->b = (unsigned char *)malloc((size_t)(n ? n : 1));
    if (!r->b) { vb_raise(7); return 0; }
    r->n = (long)fread(r->b, 1, (size_t)n, f->f);
    r->pos = 0; r->past_end = 0;
    return 1;
}

static void done_read(vbfile *f, buf *r, long start)
{
    if (f->mode & VBF_RANDOM) {
        f->eofflag = r->n < f->reclen;          /* no whole record there */
        after_record(f, start);
    } else {
        f->eofflag = r->past_end;
        fseek(f->f, start + r->pos, SEEK_SET);
    }
    free(r->b);
}

void vb_get(vb_i32 fileno, vb_i32 rec, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, void *p)
{
    vbfile *f = binfile(fileno);
    if (!f || !position(f, rec)) return;
    buf r; long start;
    if (!read_buf(f, &r, &start)) return;
    r_value(&r, kind, fixlen, udt, p, (f->mode & VBF_RANDOM) != 0);
    done_read(f, &r, start);
}

void vb_get_arr(vb_i32 fileno, vb_i32 rec, vb_i32 kind, vb_i32 fixlen, const vb_udt *udt, vbarr *a, vb_i32 dynamic)
{
    vbfile *f = binfile(fileno);
    if (!f || !position(f, rec)) return;
    buf r; long start;
    if (!read_buf(f, &r, &start)) return;
    r_array(&r, kind, fixlen, udt, a, dynamic && (f->mode & VBF_RANDOM));
    done_read(f, &r, start);
}

void vb_seek(vb_i32 fileno, vb_i32 pos)
{
    if (pos < 1) { vb_raise(63); return; }     /* (VB6 looks at the position before the file) */
    vbfile *f = vb_getfile(fileno, 0);
    if (!f) return;
    long off = f->mode & VBF_RANDOM ? (long)(pos - 1) * f->reclen : (long)pos - 1;
    fseek(f->f, off, SEEK_SET);
}

vb_i32 vb_seekpos(vb_i32 fileno)
{
    vbfile *f = vb_getfile(fileno, 0);
    if (!f) return 0;
    long p = ftell(f->f);
    if (f->mode & VBF_RANDOM) return (vb_i32)(p / f->reclen) + 1;
    return (vb_i32)p + 1;
}

vb_i32 vb_loc(vb_i32 fileno)
{
    vbfile *f = vb_getfile(fileno, 0);
    if (!f) return 0;
    if (f->mode & VBF_RANDOM) return f->lastrec;
    long p = ftell(f->f);
    if (f->mode & VBF_BINARY) return (vb_i32)p;
    return (vb_i32)((p + 127) / 128);           /* sequential files: position / 128 */
}
