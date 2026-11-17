/* varray.c - arrays as values: arrays inside Variants, Array(), ParamArray, Split, Join,
 * Filter, and the arrays that functions return.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"

void vb_aempty(vbarr *a, vb_i32 esz, vb_i32 kind)
{
    vb_afree(a);
    a->data = (char *)calloc(1, 1);
    a->esz = esz; a->kind = kind; a->fixed = 0;
    a->nd = 1; a->n = 0; a->lb[0] = 0; a->ub[0] = -1; a->mul[0] = 1;
}

void vb_amove(vbarr *dst, vbarr *src)
{
    if (dst == src) return;
    if (dst->fixed || vb_alocked(dst)) { vb_raise(10); vb_afree(src); return; }   /* the array is fixed or temporarily locked */
    vb_afree(dst);
    *dst = *src;
    memset(src, 0, sizeof *src);
}

void vb_aparam(vbarr *dst, vb_i32 n, vbvar *items)
{
    if (n == 0) { vb_aempty(dst, sizeof(vbvar), VBA_VAR); return; }
    vb_adim(dst, sizeof(vbvar), VBA_VAR, 0, 1, 0, n - 1);
    for (vb_i32 i = 0; i < n && dst->data; i++) vb_vset((vbvar *)dst->data + i, items[i]);
}

vbvar vb_array(vb_i32 n, vbvar *items)
{
    vbarr *a = vb_atmp();
    vbvar v = vb_v_empty();
    if (!a) return v;
    vb_aparam(a, n, items);
    v.vt = VT_ARRAY | VT_VARIANT; v.u.a = a;
    return v;
}

/* ---------------------------------------------------------------- Variant <-> array */
vbvar vb_v_arr(const vbarr *a, vb_i32 elemvt)
{
    vbvar v = vb_v_empty();
    vbarr *t = vb_atmp();
    if (!t) return v;
    vb_acopy(t, a);
    v.vt = (vb_u16)(VT_ARRAY | elemvt); v.u.a = t;
    return v;
}

vbarr *vb_v_arrptr(vbvar v)
{
    static vbarr none;
    if (!(v.vt & VT_ARRAY) || !v.u.a) { vb_raise(13); memset(&none, 0, sizeof none); return &none; }
    return v.u.a;
}

void vb_v2arr(vbvar v, vbarr *dst, vb_i32 elemvt)
{
    if (v.vt == VT_BSTR && elemvt == VT_UI1) { vb_s2barr(v.u.s, dst); return; }   /* b = StrConv(s, vbFromUnicode) */
    if (!(v.vt & VT_ARRAY) || (v.vt & 0xFFF) != elemvt) { vb_raise(13); return; }
    if (dst->fixed || vb_alocked(dst)) { vb_raise(10); return; }
    vb_acopy(dst, v.u.a);
}

vb_bool vb_isarray(vbvar v) { return (v.vt & VT_ARRAY) ? -1 : 0; }

static vbvar elem_get(vbarr *a, vb_i32 vt, const char *e)
{
    switch (vt) {
    case VT_I2: return vb_v_i(*(const vb_i16 *)e);
    case VT_BOOL: return vb_v_bool(*(const vb_bool *)e);
    case VT_I4: return vb_v_l(*(const vb_i32 *)e);
    case VT_R4: return vb_v_f(*(const float *)e);
    case VT_R8: return vb_v_d(*(const double *)e);
    case VT_CY: return vb_v_c(*(const vb_cy *)e);
    case VT_DATE: return vb_v_date(*(const double *)e);
    case VT_UI1: return vb_v_byte(*(const vb_u8 *)e);
    case VT_BSTR: return vb_v_s(vb_scopy(*(const vbs *)e));
    case VT_DISPATCH: return vb_vtmp(vb_v_obj(*(const vbobj *)e));
    default: (void)a; return vb_vtmp(*(const vbvar *)e);
    }
}

static void elem_set(vb_i32 vt, char *e, vbvar v, int set)
{
    switch (vt) {
    case VT_I2: *(vb_i16 *)e = vb_v2i(v); break;
    case VT_BOOL: *(vb_bool *)e = vb_v2bool(v); break;
    case VT_I4: *(vb_i32 *)e = vb_v2l(v); break;
    case VT_R4: *(float *)e = vb_v2f(v); break;
    case VT_R8: *(double *)e = vb_v2d(v); break;
    case VT_CY: *(vb_cy *)e = vb_v2c(v); break;
    case VT_DATE: *(double *)e = vb_v2date(v); break;
    case VT_UI1: *(vb_u8 *)e = vb_v2byte(v); break;
    case VT_BSTR: vb_sset((vbs *)e, vb_v2s(v)); break;
    case VT_DISPATCH: vb_oset((vbobj *)e, vb_v2obj(v)); break;
    default:
        if (v.vt == VT_DISPATCH && !set) { vbvar d = vb_late(v, "", VBM_GET, 0, 0); if (!vb_errpend) vb_vset((vbvar *)e, d); }
        else vb_vset((vbvar *)e, v);
    }
}

static char *elem_at(vbarr *a, vb_i32 argc, vbvar *argv)
{
    if (a->nd != argc || !a->data) { vb_raise(9); return 0; }
    vb_i64 off = 0;
    for (int d = 0; d < argc; d++) {
        vb_i32 i = vb_v2l(argv[d]);
        if (vb_errpend) return 0;
        if (i < a->lb[d] || i > a->ub[d]) { vb_raise(9); return 0; }
        off += (vb_i64)(i - a->lb[d]) * a->mul[d];
    }
    return a->data + off * a->esz;
}

vbvar vb_v_index(vbvar v, vb_i32 argc, vbvar *argv)
{
    if (vb_errpend) return vb_v_empty();
    if (v.vt == VT_DISPATCH) return vb_late(v, "", VBM_GET, argc, argv);   /* default member */
    if (!(v.vt & VT_ARRAY) || !v.u.a) { vb_raise(13); return vb_v_empty(); }
    char *e = elem_at(v.u.a, argc, argv);
    return e ? elem_get(v.u.a, v.vt & 0xFFF, e) : vb_v_empty();
}

void vb_v_setindex(vbvar *v, vbvar value, vb_i32 argc, vbvar *argv, vb_i32 set)
{
    if (vb_errpend) return;
    if (v->vt == VT_DISPATCH) {
        vbvar *all = (vbvar *)malloc(sizeof(vbvar) * (size_t)(argc + 1));
        if (!all) { vb_raise(7); return; }
        memcpy(all, argv, sizeof(vbvar) * (size_t)argc);
        all[argc] = value;
        vb_late(*v, "", set ? VBM_SET : VBM_LET, argc + 1, all);
        free(all);
        return;
    }
    if (!(v->vt & VT_ARRAY) || !v->u.a) { vb_raise(13); return; }
    char *e = elem_at(v->u.a, argc, argv);
    if (e) elem_set(v->vt & 0xFFF, e, value, set);
}

/* ---------------------------------------------------------------- Split, Join, Filter */
static int match_at(vbs s, vb_i32 i, vbs d, vb_i32 text)
{
    if (i + d.len > s.len) return 0;
    vbs a = { s.p + i, d.len, 0 };
    return vb_scmp(a, d, text) == 0;
}

static void push_str(vbarr *a, vb_i32 *cap, vbs s)
{
    if (a->n == *cap) {
        vb_i32 nc = *cap ? *cap * 2 : 8;
        char *nd = (char *)realloc(a->data, sizeof(vbs) * (size_t)nc);
        if (!nd) { vb_raise(7); return; }
        memset(nd + sizeof(vbs) * (size_t)*cap, 0, sizeof(vbs) * (size_t)(nc - *cap));
        a->data = nd; *cap = nc;
    }
    vb_sset((vbs *)a->data + a->n, s);
    a->n++;
}

static vbarr finish(vbarr a)
{
    if (!a.data) a.data = (char *)calloc(1, 1);
    a.nd = 1; a.lb[0] = 0; a.ub[0] = a.n - 1; a.mul[0] = 1;
    return a;
}

vbarr vb_split(vbs s, vbs delim, vb_i32 has_delim, vb_i32 limit, vb_i32 text)
{
    vbarr a; memset(&a, 0, sizeof a);
    a.esz = sizeof(vbs); a.kind = VBA_STR;
    vb_i32 cap = 0;
    static const vb_wchar sp = ' ';
    vbs d = has_delim ? delim : vb_slit(&sp, 1);
    if (limit < -1) { vb_raise(5); return finish(a); }
    if (limit == 0 || s.len == 0) return finish(a);     /* no elements: UBound = -1 */
    if (d.len == 0 || limit == 1) { push_str(&a, &cap, s); return finish(a); }
    vb_i32 start = 0;
    for (vb_i32 i = 0; i <= s.len - d.len;) {
        if ((limit < 0 || a.n < limit - 1) && match_at(s, i, d, text)) {
            vbs part = { s.p + start, i - start, 0 };
            push_str(&a, &cap, part);
            i += d.len; start = i;
        } else i++;
    }
    vbs last = { s.p + start, s.len - start, 0 };
    push_str(&a, &cap, last);
    return finish(a);
}

static vbs elem_str(vbarr *a, vb_i32 vt, vb_i32 i)
{
    char *e = a->data + (vb_i64)i * a->esz;
    if (vt == VT_BSTR) return *(vbs *)e;
    return vb_v2s(elem_get(a, vt, e));
}

vbs vb_join(vbarr *a, vb_i32 elemvt, vbs delim, vb_i32 has_delim)
{
    static const vb_wchar sp = ' ';
    vbs d = has_delim ? delim : vb_slit(&sp, 1);
    if (a->nd > 1) { vb_raise(5); return vb_stmp(0); }
    vb_i64 total = 0;
    for (vb_i32 i = 0; i < a->n; i++) total += elem_str(a, elemvt, i).len + (i ? d.len : 0);
    if (vb_errpend) return vb_stmp(0);
    vbs r = vb_stmp((vb_i32)total);
    vb_i32 k = 0;
    for (vb_i32 i = 0; i < a->n; i++) {
        if (i) { memcpy(r.p + k, d.p, sizeof(vb_wchar) * (size_t)d.len); k += d.len; }
        vbs e = elem_str(a, elemvt, i);
        if (e.len) memcpy(r.p + k, e.p, sizeof(vb_wchar) * (size_t)e.len);
        k += e.len;
    }
    return r;
}

vbarr vb_filter(vbarr *a, vb_i32 elemvt, vbs match, vb_i32 include, vb_i32 text)
{
    vbarr r; memset(&r, 0, sizeof r);
    r.esz = sizeof(vbs); r.kind = VBA_STR;
    vb_i32 cap = 0;
    if (a->nd != 1) { vb_raise(a->nd ? 5 : 13); return finish(r); }
    for (vb_i32 i = 0; i < a->n && !vb_errpend; i++) {
        vbs e = elem_str(a, elemvt, i);
        int found = vb_instr(1, e, match, text) > 0;
        if (found == (include != 0)) push_str(&r, &cap, e);
    }
    return finish(r);
}

/* For Each over an array in a Variant (used by vb_each_count / vb_each_item) */
vb_i32 vb_varr_count(vbvar v) { return v.u.a ? v.u.a->n : 0; }
vbvar vb_varr_item(vbvar v, vb_i32 i) { return elem_get(v.u.a, v.vt & 0xFFF, v.u.a->data + (vb_i64)i * v.u.a->esz); }

/* Byte arrays and strings: b = s takes the bytes of the string (two in each character, as VB6
 * keeps strings), s = b makes a string of the bytes (StrConv(b, vbUnicode) reads them as text) */
void vb_s2barr(vbs s, vbarr *dst)
{
    if (dst->fixed || vb_alocked(dst)) { vb_raise(10); return; }
    vb_i32 n = vb_sbytes(s);                   /* (a byte string: an odd number too) */
    vb_aerase(dst);
    if (n == 0) {                               /* an empty array: LBound 0, UBound -1 */
        dst->nd = 1; dst->n = 0; dst->esz = 1; dst->kind = VBA_PLAIN;
        dst->lb[0] = 0; dst->ub[0] = -1; dst->mul[0] = 1;
        return;
    }
    vb_aredim(dst, 1, VBA_PLAIN, 0, 1, 0, n - 1);
    if (vb_errpend) return;
    for (vb_i32 i = 0; i < n; i++) dst->data[i] = (char)vb_sbyte(s, i);
}
vbs vb_barr2s(const vbarr *a)
{
    if (!a || !a->nd || !a->data) return vb_stmp(0);
    return vb_sfrombytes(a->data, a->n * a->esz);
}
