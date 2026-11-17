/* array.c - arrays: Dim, ReDim [Preserve], Erase, bounds, element access.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"

static void free_elems(vbarr *a)
{
    if (!a->data) return;
    for (vb_i32 i = 0; i < a->n; i++) {
        char *e = a->data + (vb_i64)i * a->esz;
        switch (a->kind) {
        case VBA_STR: vb_sfree((vbs *)e); break;
        case VBA_VAR: vb_vfree((vbvar *)e); break;
        case VBA_OBJ: vb_ofree((vbobj *)e); break;
        case VBA_UDT: if (a->udt && a->udt->free) a->udt->free(e); break;
        }
    }
}

/* new elements of a Type with fixed arrays inside need their set-up */
static void init_elems(vbarr *a, char *data, vb_i32 from, vb_i32 to)
{
    if (a->kind != VBA_UDT || !a->udt || !a->udt->init) return;
    for (vb_i32 i = from; i < to; i++) a->udt->init(data + (vb_i64)i * a->esz);
}

static int layout(vbarr *a, vb_i32 nd, va_list ap, vb_i32 *lb, vb_i32 *ub)
{
    vb_i64 n = 1;
    for (int d = 0; d < nd; d++) {
        lb[d] = va_arg(ap, vb_i32); ub[d] = va_arg(ap, vb_i32);
        if (ub[d] < lb[d]) { vb_raise(9); return -1; }
        n *= (vb_i64)(ub[d] - lb[d] + 1);
        if (n > 0x7FFFFFFF) { vb_raise(7); return -1; }
    }
    (void)a;
    return (int)n;
}

static void setup(vbarr *a, vb_i32 nd, vb_i32 *lb, vb_i32 *ub, vb_i32 n)
{
    a->nd = nd; a->n = n;
    vb_i32 mul = 1;
    for (int d = 0; d < nd; d++) { a->lb[d] = lb[d]; a->ub[d] = ub[d]; a->mul[d] = mul; mul *= ub[d] - lb[d] + 1; }
}

void vb_adim(vbarr *a, vb_i32 esz, vb_i32 kind, vb_i32 fixed, vb_i32 nd, ...)
{
    vb_i32 lb[8], ub[8];
    va_list ap; va_start(ap, nd);
    int n = layout(a, nd, ap, lb, ub);
    va_end(ap);
    if (n < 0) return;
    a->esz = esz; a->kind = kind; a->fixed = fixed;
    a->data = (char *)calloc((size_t)n ? (size_t)n : 1, (size_t)esz);
    if (!a->data) { vb_raise(7); a->nd = 0; return; }
    setup(a, nd, lb, ub, n);
    init_elems(a, a->data, 0, n);
}

/* an array whose element is passed ByRef to a procedure is locked while it runs: ReDim / Erase of it
 * is "This array is fixed or temporarily locked" (10), as in VB6 (its memory must not move) */
static const vbarr *alocks[256]; static int nalocks;
void vb_alock(const vbarr *a) { if (nalocks < 256) alocks[nalocks++] = a; }
void vb_aunlock(const vbarr *a)
{
    for (int i = nalocks - 1; i >= 0; i--) if (alocks[i] == a) { alocks[i] = alocks[--nalocks]; return; }
}
int vb_alocked(const vbarr *a)
{
    for (int i = 0; i < nalocks; i++) if (alocks[i] == a) return 1;
    return 0;
}

void vb_aredim(vbarr *a, vb_i32 esz, vb_i32 kind, vb_i32 preserve, vb_i32 nd, ...)
{
    if (a->fixed || vb_alocked(a)) { vb_raise(10); return; }
    vb_i32 lb[8] = { 0 }, ub[8] = { 0 };
    va_list ap; va_start(ap, nd);
    int n = layout(a, nd, ap, lb, ub);
    va_end(ap);
    if (n < 0) return;
    if (preserve && a->nd == nd && a->data && a->esz == esz && a->lb[nd - 1] == lb[nd - 1]) {
        /* ReDim Preserve with the same lower bound (the common ReDim Preserve a(UBound(a) + 1) in a loop): the
         * elements keep their place, so the memory grows or shrinks in place (realloc), as VB6's SafeArrayRedim;
         * a new block and a copy each time made such a loop quadratic (robot: BrionWordSearch, 170,000 words) */
        int same = 1;
        for (int d = 0; d < nd - 1; d++) if (a->lb[d] != lb[d] || a->ub[d] != ub[d]) same = 0;
        if (!same) { vb_raise(9); return; }
        vb_i32 old = a->n;
        if (n < old) { vbarr tmp = *a; tmp.data = a->data + (vb_i64)n * esz; tmp.n = old - n; free_elems(&tmp); }
        char *nd2 = (char *)realloc(a->data, (size_t)(n ? n : 1) * (size_t)esz);
        if (!nd2) { vb_raise(7); return; }
        a->data = nd2; a->kind = kind;
        if (n > old) {
            memset(a->data + (vb_i64)old * esz, 0, (size_t)(n - old) * esz);
            init_elems(a, a->data, old, n);
        }
        setup(a, nd, lb, ub, n);
        return;
    }
    char *nw = (char *)calloc((size_t)n ? (size_t)n : 1, (size_t)esz);
    if (!nw) { vb_raise(7); return; }
    if (preserve && a->nd) {
        /* only the last dimension may change, and the number of dimensions must not */
        if (a->nd != nd) { free(nw); vb_raise(9); return; }
        for (int d = 0; d < nd - 1; d++)
            if (a->lb[d] != lb[d] || a->ub[d] != ub[d]) { free(nw); vb_raise(9); return; }
        vb_i32 inner = 1;                       /* elements per last-dim slice (an empty Array() too) */
        for (int d = 0; d < nd - 1; d++) inner *= a->ub[d] - a->lb[d] + 1;
        a->kind = kind;
        for (vb_i32 k = lb[nd - 1]; k <= ub[nd - 1]; k++)              /* new slices: set-up first */
            if (k < a->lb[nd - 1] || k > a->ub[nd - 1])
                init_elems(a, nw + (vb_i64)(k - lb[nd - 1]) * inner * esz, 0, inner);
        for (vb_i32 k = a->lb[nd - 1]; k <= a->ub[nd - 1]; k++) {
            char *src = a->data + (vb_i64)(k - a->lb[nd - 1]) * inner * esz;
            if (k >= lb[nd - 1] && k <= ub[nd - 1]) {
                memcpy(nw + (vb_i64)(k - lb[nd - 1]) * inner * esz, src, (size_t)inner * esz);
            } else {                              /* dropped slice: free its strings, objects ... */
                vbarr tmp = *a; tmp.data = src; tmp.n = inner; free_elems(&tmp);
            }
        }
        free(a->data);
        a->data = nw; a->esz = esz;
        setup(a, nd, lb, ub, n);
        return;
    } else {
        free_elems(a); free(a->data);
    }
    a->data = nw; a->esz = esz; a->kind = kind;
    setup(a, nd, lb, ub, n);
    init_elems(a, a->data, 0, n);
}

void vb_aerase(vbarr *a)
{
    if (vb_alocked(a)) { vb_raise(10); return; }
    if (a->fixed) {                               /* fixed arrays are reset, not freed */
        free_elems(a);
        memset(a->data, 0, (size_t)a->n * a->esz);
        init_elems(a, a->data, 0, a->n);
        return;
    }
    vb_afree(a);
}

void vb_afree(vbarr *a)
{
    free_elems(a);
    free(a->data);
    a->data = 0; a->nd = 0; a->n = 0;
}

void *vb_aelem_err(vbarr *a)
{
    static char dummy[64];
    (void)a;
    vb_raise(9);
    memset(dummy, 0, sizeof dummy);
    return dummy;
}

/* Remove Array Bounds Checks (native code of VB6): no error 9; the indices give an offset in the memory of the
   array (m(4, 1) of m(1 To 3, 1 To 3) is m(1, 2)); outside it VB6 reads and writes whatever is there: here a
   scratch element of zeros. An array without dimensions: error 91 (VB6 reads a null descriptor) */
void *vb_aelem_nc(vbarr *a, vb_i32 nd, ...)
{
    static char scratch[65536];
    if (!a->nd || !a->data) { vb_raise(91); memset(scratch, 0, sizeof scratch); return scratch; }
    va_list ap; va_start(ap, nd);
    vb_i64 off = 0;
    for (int d = 0; d < nd; d++) {
        vb_i32 i = va_arg(ap, vb_i32);
        if (d < a->nd) off += (vb_i64)(i - a->lb[d]) * a->mul[d];
    }
    va_end(ap);
    if (off < 0 || off >= a->n || a->esz > (vb_i32)sizeof scratch) {
        memset(scratch, 0, sizeof scratch);
        return scratch;
    }
    return a->data + off * a->esz;
}

void *vb_aelem(vbarr *a, vb_i32 nd, ...)
{
    if (a->nd != nd) return vb_aelem_err(a);
    va_list ap; va_start(ap, nd);
    vb_i64 off = 0;
    for (int d = 0; d < nd; d++) {
        vb_i32 i = va_arg(ap, vb_i32);
        if (i < a->lb[d] || i > a->ub[d]) { va_end(ap); return vb_aelem_err(a); }
        off += (vb_i64)(i - a->lb[d]) * a->mul[d];
    }
    va_end(ap);
    return a->data + off * a->esz;
}

vb_i32 vb_lbound(vbarr *a, vb_i32 dim)
{
    if (dim < 1 || dim > a->nd) { vb_raise(9); return 0; }
    return a->lb[dim - 1];
}
vb_i32 vb_ubound(vbarr *a, vb_i32 dim)
{
    if (dim < 1 || dim > a->nd) { vb_raise(9); return 0; }
    return a->ub[dim - 1];
}

/* dst = a deep copy of src (used when a Type with an array inside is copied) */
void vb_acopy(vbarr *dst, const vbarr *src)
{
    if (dst == src) return;
    vbarr nw = *src;
    nw.data = 0;
    if (src->data && src->nd) {
        nw.data = (char *)calloc((size_t)src->n ? (size_t)src->n : 1, (size_t)src->esz);
        if (!nw.data) { vb_raise(7); return; }
        for (vb_i32 i = 0; i < src->n; i++) {
            const char *s = src->data + (vb_i64)i * src->esz;
            char *d = nw.data + (vb_i64)i * src->esz;
            switch (src->kind) {
            case VBA_STR: vb_sset((vbs *)d, *(const vbs *)s); break;
            case VBA_VAR: vb_vset((vbvar *)d, *(const vbvar *)s); break;
            case VBA_OBJ: vb_oset((vbobj *)d, *(const vbobj *)s); break;
            case VBA_UDT: if (src->udt && src->udt->copy) { src->udt->copy(d, s); break; } /* fall through */
            default: memcpy(d, s, (size_t)src->esz);
            }
        }
    }
    free_elems(dst); free(dst->data);
    *dst = nw;
}

void *vb_aelem_lin(vbarr *a, vb_i32 i)
{
    if (i < 0 || i >= a->n) return vb_aelem_err(a);
    return a->data + (vb_i64)i * a->esz;
}

/* Erase v: a Variant holding an array keeps an erased (undimensioned) array of the same type - VB6
 * (compile-6.8 c28): IsArray True, TypeName "Variant()", UBound and v(0) error 9; any other Variant: Empty */
void vb_v_erase(vbvar *v)
{
    if ((v->vt & VT_ARRAY) && v->u.a) { vb_aerase(v->u.a); return; }
    vb_vset(v, vb_v_empty());
}
