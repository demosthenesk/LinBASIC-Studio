/* variant.c - the Variant type: storage, conversions, operators.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"

vbvar vb_v_empty(void) { vbvar v; memset(&v, 0, sizeof v); return v; }
vbvar vb_v_null(void) { vbvar v = vb_v_empty(); v.vt = VT_NULL; return v; }
vbvar vb_v_missing(void) { vbvar v = vb_v_empty(); v.vt = VT_ERROR; v.u.l = (vb_i32)0x80020004; return v; }
vbvar vb_cverr(vb_i32 n) { vbvar v = vb_v_empty(); v.vt = VT_ERROR; v.u.l = (vb_i32)(0x800A0000u | (vb_u32)(n & 0xFFFF)); return v; }
vbvar vb_v_i(vb_i16 x) { vbvar v = vb_v_empty(); v.vt = VT_I2; v.u.i = x; return v; }
vbvar vb_v_l(vb_i32 x) { vbvar v = vb_v_empty(); v.vt = VT_I4; v.u.l = x; return v; }
vbvar vb_v_f(float x) { vbvar v = vb_v_empty(); v.vt = VT_R4; v.u.f = x; return v; }
vbvar vb_v_d(double x) { vbvar v = vb_v_empty(); v.vt = VT_R8; v.u.d = x; return v; }
vbvar vb_v_c(vb_cy x) { vbvar v = vb_v_empty(); v.vt = VT_CY; v.u.cy = x; return v; }
vbvar vb_v_date(vb_date x) { vbvar v = vb_v_empty(); v.vt = VT_DATE; v.u.d = x; return v; }
vbvar vb_v_bool(vb_bool x) { vbvar v = vb_v_empty(); v.vt = VT_BOOL; v.u.i = x ? -1 : 0; return v; }
vbvar vb_v_byte(vb_u8 x) { vbvar v = vb_v_empty(); v.vt = VT_UI1; v.u.b = x; return v; }
vbvar vb_v_s(vbs x) { vbvar v = vb_v_empty(); v.vt = VT_BSTR; v.u.s = x; v.u.s.flags &= VBS_ODD; return v; }

void vb_vfree(vbvar *v)
{
    vbvar old = *v;
    memset(v, 0, sizeof *v);
    if (old.vt == VT_BSTR) vb_sfree(&old.u.s);
    else if (old.vt == VT_DISPATCH && old.u.o) vb_orelease(old.u.o);   /* may run Class_Terminate */
    else if ((old.vt & VT_ARRAY) && old.u.a) { vb_afree(old.u.a); free(old.u.a); }
    else if (old.vt == VT_RECORD && !vb_rec_istmp(old.u.rec)) vb_rec_free(old.u.rec);   /* (a temporary: vb_release) */
}

vbvar vb_vtmp(vbvar v)
{
    if (v.vt == VT_BSTR) v.u.s = vb_scopy(v.u.s);
    else if (v.vt == VT_DISPATCH && v.u.o) { vb_oaddref(v.u.o); vb_otmp(v.u.o); }
    else if ((v.vt & VT_ARRAY) && v.u.a) { vbarr *t = vb_atmp(); if (t) vb_acopy(t, v.u.a); v.u.a = t; }
    else if (v.vt == VT_RECORD && v.u.rec) v.u.rec = vb_rec_tmp(vb_rec_udt(v.u.rec), vb_rec_data(v.u.rec));
    return v;
}

void vb_vset(vbvar *dst, vbvar src)
{
    if (src.vt & VT_ARRAY) {                    /* arrays are copied, like in VB */
        vbarr *a = (vbarr *)calloc(1, sizeof(vbarr));
        if (!a) { vb_raise(7); return; }
        if (src.u.a) vb_acopy(a, src.u.a);
        vb_vfree(dst);
        dst->vt = src.vt; dst->u.a = a;
        return;
    }
    if (src.vt == VT_RECORD) {                  /* a Type: copied, like in VB */
        struct vbrec *r = src.u.rec ? vb_rec_new(vb_rec_udt(src.u.rec), vb_rec_data(src.u.rec)) : 0;
        vb_vfree(dst);
        dst->vt = VT_RECORD; dst->u.rec = r;
        return;
    }
    if (src.vt == VT_DISPATCH) {
        if (src.u.o) vb_oaddref(src.u.o);
        vb_vfree(dst);
        *dst = src;
        return;
    }
    if (src.vt == VT_BSTR) {
        vbs keep = { 0, 0, 0 };
        vb_sset(&keep, src.u.s);                /* copy first: src may be dst itself */
        vb_vfree(dst);
        dst->vt = VT_BSTR; dst->u.s = keep;
        return;
    }
    vb_vfree(dst);
    *dst = src;
}

/* ---------------------------------------------------------------- conversions */
/* an object used as a value: its default member (error 438 if it has none) */
static vbvar deref(vbvar v)
{
    while (v.vt == VT_DISPATCH && !vb_errpend) v = vb_late(v, "", VBM_GET, 0, 0);
    return v;
}

static double num(vbvar v, int *ok)
{
    *ok = 1;
    if (v.vt == VT_DISPATCH) { v = deref(v); if (vb_errpend) { *ok = 0; return 0; } }
    switch (v.vt) {
    case VT_EMPTY: return 0;
    case VT_I2: case VT_BOOL: return v.u.i;
    case VT_I4: return v.u.l;
    case VT_R4: return v.u.f;
    case VT_R8: case VT_DATE: return v.u.d;
    case VT_CY: return vb_cy2d(v.u.cy);
    case VT_DECIMAL: return vb_dec2d(v);
    case VT_UI1: return v.u.b;
    case VT_BSTR: return vb_s2d(v.u.s);
    case VT_NULL: vb_raise(94); *ok = 0; return 0;
    default: vb_raise(13); *ok = 0; return 0;
    }
}
vb_i16 vb_v2i(vbvar v) { int ok; if (v.vt == VT_I2 || v.vt == VT_BOOL) return v.u.i; if (v.vt == VT_DECIMAL) return (vb_i16)vb_dec2i64(v, -32768, 32767); double d = num(v, &ok); return ok ? vb_cint(d) : 0; }
vb_i32 vb_v2l(vbvar v) { int ok; if (v.vt == VT_I4) return v.u.l; if (v.vt == VT_DECIMAL) return (vb_i32)vb_dec2i64(v, -2147483647 - 1, 2147483647); double d = num(v, &ok); return ok ? vb_clng(d) : 0; }
float vb_v2f(vbvar v) { int ok; if (v.vt == VT_R4) return v.u.f; double d = num(v, &ok); return ok ? vb_csng(d) : 0; }
double vb_v2d(vbvar v) { int ok; return num(v, &ok); }
vb_cy vb_v2c(vbvar v) { int ok; if (v.vt == VT_CY) return v.u.cy; if (v.vt == VT_DECIMAL) return vb_dec2cy(v); double d = num(v, &ok); return ok ? vb_ccur(d) : 0; }
vb_bool vb_v2bool(vbvar v) { int ok; if (v.vt == VT_DISPATCH) v = deref(v); if (v.vt == VT_BSTR) return vb_s2bool(v.u.s); double d = num(v, &ok); return ok && d != 0 ? -1 : 0; }
vb_u8 vb_v2byte(vbvar v) { int ok; if (v.vt == VT_DECIMAL) return (vb_u8)vb_dec2i64(v, 0, 255); double d = num(v, &ok); return ok ? vb_cbyte(d) : 0; }
vb_date vb_v2date(vbvar v) { int ok; if (v.vt == VT_DISPATCH) v = deref(v); if (v.vt == VT_BSTR) return vb_s2date(v.u.s); return num(v, &ok); }

vbs vb_v2s(vbvar v)
{
    if (v.vt == VT_DISPATCH) { v = deref(v); if (vb_errpend) return vb_stmp(0); }
    switch (v.vt) {
    case VT_EMPTY: return vb_stmp(0);
    case VT_I2: return vb_cstr_i(v.u.i);
    case VT_I4: return vb_cstr_i(v.u.l);
    case VT_UI1: return vb_cstr_i(v.u.b);
    case VT_R4: return vb_cstr_f(v.u.f);
    case VT_R8: return vb_cstr_d(v.u.d);
    case VT_CY: return vb_cstr_c(v.u.cy);
    case VT_DECIMAL: return vb_dec2s(v);
    case VT_BOOL: return vb_cstr_b(v.u.i);
    case VT_DATE: return vb_cstr_date(v.u.d);
    case VT_BSTR: return v.u.s;
    case VT_NULL: vb_raise(94); return vb_stmp(0);
    case VT_ARRAY | VT_UI1: return vb_barr2s(v.u.a);  /* s = a Variant with a Byte array */
    default: vb_raise(13); return vb_stmp(0);
    }
}

/* ---------------------------------------------------------------- operators */
/* numeric rank: 1 Byte, 2 Integer, 3 Long, 4 Single, 5 Currency, 6 Double, 7 Date */
static int rank(vbvar v)
{
    switch (v.vt) {
    case VT_UI1: return 1;
    case VT_EMPTY: case VT_I2: case VT_BOOL: return 2;
    case VT_I4: return 3;
    case VT_R4: return 4;
    case VT_CY: return 5;
    case VT_R8: case VT_BSTR: return 6;
    case VT_DATE: return 7;
    case VT_DECIMAL: return 8;
    default: return 0;
    }
}

static vbvar mk(int r, double d)
{
    /* integer results grow to the next type on overflow, like VB6 */
    if (r == 1) { if (d >= 0 && d <= 255) return vb_v_byte((vb_u8)d); r = 2; }
    if (r == 2) { if (d >= -32768 && d <= 32767) return vb_v_i((vb_i16)d); r = 3; }
    if (r == 3) { if (d >= -2147483648.0 && d <= 2147483647.0) return vb_v_l((vb_i32)d); return vb_v_d(vb_chkd(d)); }
    if (r == 4) { if (fabs(d) <= 3.402823466e38) return vb_v_f((float)d); return vb_v_d(vb_chkd(d)); }
    if (r == 5) return vb_v_c(vb_ccur(d));
    if (r == 7) return vb_v_date(vb_chkd(d));
    return vb_v_d(vb_chkd(d));
}

static int both_null(vbvar a, vbvar b) { return a.vt == VT_NULL || b.vt == VT_NULL; }

/* an operand of an operator with a Decimal: a Single goes through a Double (VB6: 0.1! <> CDec("0.1")) */
static vbvar dec_of(vbvar v)
{
    if (v.vt == VT_R4) return vb_dec_from_double(v.u.f, 15);
    return vb_cdec(v);
}

static vbvar arith(vbvar a, vbvar b, int op)
{
    a = deref(a); b = deref(b);
    if (both_null(a, b)) return vb_v_null();
    if (a.vt == VT_EMPTY && b.vt == VT_EMPTY) return vb_v_i(0);
    int ra = rank(a), rb = rank(b);
    if (!ra || !rb) { vb_raise(13); return vb_v_empty(); }
    int r = ra > rb ? ra : rb;
    if ((ra == 4 && rb == 3) || (ra == 3 && rb == 4)) r = 6;          /* Single with Long gives Double */
    if (r == 7 && op == 2) r = 6;                                       /* Date * x is a Double */
    if (a.vt == VT_DATE && b.vt == VT_DATE && op == 1) r = 6;           /* Date - Date is a Double */
    if (r == 8) {                               /* a Decimal: Decimal arithmetic (decimal.c) */
        vbvar x = dec_of(a); if (vb_errpend) return vb_v_empty();
        vbvar y = dec_of(b); if (vb_errpend) return vb_v_empty();
        return op == 0 ? vb_dec_add(x, y) : op == 1 ? vb_dec_sub(x, y) : vb_dec_mul(x, y);
    }
    if (r == 5) {
        vb_cy x = vb_v2c(a), y = vb_v2c(b);
        return vb_v_c(op == 0 ? vb_add_c(x, y) : op == 1 ? vb_sub_c(x, y) : vb_mul_c(x, y));
    }
    int ok;
    double x = num(a, &ok), y; if (!ok) return vb_v_empty();
    y = num(b, &ok); if (!ok) return vb_v_empty();
    double z = op == 0 ? x + y : op == 1 ? x - y : x * y;
    if (r == 4) return mk(4, (float)z);
    return mk(r, z);
}

vbvar vb_v_add(vbvar a, vbvar b)
{
    if (a.vt == VT_BSTR && b.vt == VT_BSTR) return vb_v_s(vb_scat(a.u.s, b.u.s));
    if (a.vt == VT_BSTR && b.vt == VT_EMPTY) return a;
    if (b.vt == VT_BSTR && a.vt == VT_EMPTY) return b;
    return arith(a, b, 0);
}
vbvar vb_v_sub(vbvar a, vbvar b) { return arith(a, b, 1); }
vbvar vb_v_mul(vbvar a, vbvar b) { return arith(a, b, 2); }

vbvar vb_v_div(vbvar a, vbvar b)
{
    if (both_null(a, b)) return vb_v_null();
    a = deref(a); b = deref(b);
    if (a.vt == VT_DECIMAL || b.vt == VT_DECIMAL) {
        vbvar x = dec_of(a); if (vb_errpend) return vb_v_empty();
        vbvar y = dec_of(b); if (vb_errpend) return vb_v_empty();
        return vb_dec_div(x, y);
    }
    int ok; double x = num(a, &ok); if (!ok) return vb_v_empty();
    double y = num(b, &ok); if (!ok) return vb_v_empty();
    int ra = rank(a), rb = rank(b);
    if (ra <= 4 && rb <= 4 && ra != 3 && rb != 3 && (ra == 4 || rb == 4)) return vb_v_f(vb_div_f((float)x, (float)y));
    return vb_v_d(vb_div(x, y));
}

static vbvar intop(vbvar a, vbvar b, int op)
{
    if (both_null(a, b)) return vb_v_null();
    int ok; double x = num(a, &ok); if (!ok) return vb_v_empty();
    double y = num(b, &ok); if (!ok) return vb_v_empty();
    vb_i32 xi = vb_clng(x), yi = vb_clng(y);
    vb_i32 r = op == 0 ? vb_idiv(xi, yi) : op == 1 ? vb_mod(xi, yi) : 0;
    int rk = rank(a) > rank(b) ? rank(a) : rank(b);
    if (rank(a) == 1 && rank(b) == 1) return vb_v_byte((vb_u8)r);
    if (rk <= 2) return vb_v_i((vb_i16)r);
    return vb_v_l(r);
}
vbvar vb_v_idiv(vbvar a, vbvar b) { return intop(a, b, 0); }
vbvar vb_v_mod(vbvar a, vbvar b) { return intop(a, b, 1); }

vbvar vb_v_pow(vbvar a, vbvar b)
{
    if (both_null(a, b)) return vb_v_null();
    int ok; double x = num(a, &ok); if (!ok) return vb_v_empty();
    double y = num(b, &ok); if (!ok) return vb_v_empty();
    return vb_v_d(vb_pow(x, y));
}

vbvar vb_v_neg(vbvar a)
{
    switch (a.vt) {
    case VT_NULL: return a;
    case VT_EMPTY: return vb_v_i(0);
    case VT_UI1: return mk(2, -(double)a.u.b);
    case VT_I2: case VT_BOOL: return mk(2, -(double)a.u.i);
    case VT_I4: return mk(3, -(double)a.u.l);
    case VT_R4: return vb_v_f(-a.u.f);
    case VT_CY: return vb_v_c(vb_sub_c(0, a.u.cy));
    case VT_DECIMAL: return vb_dec_neg(a);
    case VT_DATE: return vb_v_date(-a.u.d);
    default: { int ok; double d = num(a, &ok); return ok ? vb_v_d(-d) : vb_v_empty(); }
    }
}

vbvar vb_v_cat(vbvar a, vbvar b)
{
    if (a.vt == VT_NULL && b.vt == VT_NULL) return vb_v_null();
    vbs x = a.vt == VT_NULL ? vb_stmp(0) : vb_v2s(a);
    vbs y = b.vt == VT_NULL ? vb_stmp(0) : vb_v2s(b);
    return vb_v_s(vb_scat(x, y));
}

static vbvar logic(vbvar a, vbvar b, int op)
{
    if (a.vt == VT_NULL || b.vt == VT_NULL) {
        /* Null propagation with the usual exceptions: False And Null = False, True Or Null = True */
        vbvar o = a.vt == VT_NULL ? b : a;
        if (o.vt != VT_NULL) {
            vb_i32 x = vb_v2l(o);
            if (op == 0 && x == 0) return vb_v_i(0);
            if (op == 1 && x == -1) return vb_v_i(-1);
        }
        return vb_v_null();
    }
    int bools = a.vt == VT_BOOL && b.vt == VT_BOOL;
    vb_i32 x = vb_v2l(a), y = vb_v2l(b), r;
    switch (op) {
    case 0: r = x & y; break;
    case 1: r = x | y; break;
    case 2: r = x ^ y; break;
    case 3: r = ~(x ^ y); break;
    default: r = (~x) | y; break;
    }
    if (bools) return vb_v_bool((vb_bool)r);
    int rk = rank(a) > rank(b) ? rank(a) : rank(b);
    if (rk <= 2) return vb_v_i((vb_i16)r);
    return vb_v_l(r);
}
vbvar vb_v_and(vbvar a, vbvar b) { return logic(a, b, 0); }
vbvar vb_v_or(vbvar a, vbvar b) { return logic(a, b, 1); }
vbvar vb_v_xor(vbvar a, vbvar b) { return logic(a, b, 2); }
vbvar vb_v_eqv(vbvar a, vbvar b) { return logic(a, b, 3); }
vbvar vb_v_imp(vbvar a, vbvar b) { return logic(a, b, 4); }
vbvar vb_v_not(vbvar a)
{
    if (a.vt == VT_NULL) return a;
    if (a.vt == VT_BOOL) return vb_v_bool(~a.u.i);
    vb_i32 x = vb_v2l(a);
    return rank(a) <= 2 ? vb_v_i((vb_i16)~x) : vb_v_l(~x);
}

/* flags: 1 = Option Compare Text, 2 = a had a static type, 4 = b had a static type */
vbvar vb_v_cmp(vbvar a, vbvar b, vb_i32 op, vb_i32 flags)
{
    a = deref(a); b = deref(b);
    if ((a.vt | b.vt) & VT_ARRAY) { vb_raise(13); return vb_v_bool(0); }   /* an array compared (VB6 oracle f_dbgrid) */
    if (a.vt == VT_NULL || b.vt == VT_NULL) return vb_v_null();
    int c, text = flags & 1;
    int an = a.vt != VT_BSTR && a.vt != VT_EMPTY, bn = b.vt != VT_BSTR && b.vt != VT_EMPTY;
    if ((an && b.vt == VT_BSTR) || (bn && a.vt == VT_BSTR)) {
        int typed_num = an ? (flags & 2) : (flags & 4);
        int typed_str = an ? (flags & 4) : (flags & 2);
        if (typed_str) {                           /* Variant number with a String: string comparison */
            vbs x = vb_v2s(a), y = vb_v2s(b);
            c = vb_scmp(x, y, text);
        } else if (typed_num) {                    /* typed number with a Variant string: numeric comparison */
            int ok; double x = num(a, &ok), y;
            if (!ok) return vb_v_bool(0);
            y = num(b, &ok);
            if (!ok) return vb_v_bool(0);
            c = x < y ? -1 : x > y ? 1 : 0;
        } else c = an ? -1 : 1;                    /* two Variants: a number is less than a string */
    } else if (a.vt == VT_BSTR && b.vt == VT_BSTR) c = vb_scmp(a.u.s, b.u.s, text);
    else if ((a.vt == VT_BSTR && b.vt == VT_EMPTY) || (a.vt == VT_EMPTY && b.vt == VT_BSTR)) {
        vbs e = vb_stmp(0);
        c = vb_scmp(a.vt == VT_BSTR ? a.u.s : e, b.vt == VT_BSTR ? b.u.s : e, text);
    } else if (a.vt == VT_DECIMAL || b.vt == VT_DECIMAL) {
        vbvar x = dec_of(a), y;
        if (vb_errpend) return vb_v_bool(0);
        y = dec_of(b);
        if (vb_errpend) return vb_v_bool(0);
        c = vb_dec_cmp(x, y);
    } else {
        int ok; double x = num(a, &ok), y;
        if (!ok) return vb_v_bool(0);
        y = num(b, &ok);
        if (!ok) return vb_v_bool(0);
        c = x < y ? -1 : x > y ? 1 : 0;
    }
    vb_bool r;
    switch (op) {
    case 0: r = c == 0; break;
    case 1: r = c != 0; break;
    case 2: r = c < 0; break;
    case 3: r = c > 0; break;
    case 4: r = c <= 0; break;
    default: r = c >= 0; break;
    }
    return vb_v_bool(r ? -1 : 0);
}

vb_bool vb_v_truth(vbvar v)
{
    if (v.vt == VT_NULL) return 0;              /* If Null ... takes the Else branch */
    return vb_v2bool(v);
}

vb_i32 vb_vartype(vbvar v) { return v.vt; }   /* VT_DISPATCH = vbObject (9) */
/* IsNull, IsEmpty: an object is its default value (VB6: IsNull(rs!Name) of a Field, an rdoColumn; the VB6 oracle
 * f_rdo1 shows the default value read for each of them); one without a default value is an object */
vb_i32 vb_vartype_def(vbvar v)
{
    if (v.vt != VT_DISPATCH || !v.u.o) return v.vt;
    void *keep = vb_err_stash();
    vb_i32 mk = vb_mark();
    vbvar d = vb_late_o(v.u.o, "", VBM_GET, 0, 0);
    vb_i32 t = vb_errpend ? VT_DISPATCH : d.vt == VT_DISPATCH ? VT_DISPATCH : d.vt;
    if (vb_errpend) vb_err_clear_pending();
    vb_release(mk);
    vb_err_unstash(keep);
    return t;
}

vbs vb_typename(vbvar v)
{
    const char *n;
    switch (v.vt) {
    case VT_EMPTY: n = "Empty"; break; case VT_NULL: n = "Null"; break; case VT_I2: n = "Integer"; break;
    case VT_I4: n = "Long"; break; case VT_R4: n = "Single"; break; case VT_R8: n = "Double"; break;
    case VT_CY: n = "Currency"; break; case VT_DATE: n = "Date"; break; case VT_BSTR: n = "String"; break;
    case VT_BOOL: n = "Boolean"; break; case VT_UI1: n = "Byte"; break; case VT_ERROR: n = "Error"; break;
    case VT_DECIMAL: n = "Decimal"; break;
    case VT_DISPATCH: return vb_otypename(v.u.o);
    case VT_RECORD: {                           /* the name of the Type, without its project */
        const vb_udt *u = vb_rec_udt(v.u.rec);
        const char *nm = u && u->name ? u->name : "Unknown", *dot = strrchr(nm, '.');
        return vb_sascii(dot ? dot + 1 : nm);
    }
    default:
        if (v.vt & VT_ARRAY) {
            vbvar e = vb_v_empty(); e.vt = (vb_u16)(v.vt & 0xFFF);
            vbs t = e.vt == VT_VARIANT ? vb_sascii("Variant") : e.vt == VT_DISPATCH ? vb_sascii("Object") : vb_typename(e);
            return vb_scat(t, vb_sascii("()"));
        }
        n = "Unknown";
    }
    return vb_sascii(n);
}

/* IsDate: a Date, or a string that CDate would take */
vb_bool vb_isdate(vbvar v)
{
    if (v.vt == VT_DISPATCH && v.u.o) { v = vb_v_letval(v); if (vb_errpend) { vb_errpend = 0; return 0; } }   /* IsDate(Text1) */
    if (v.vt == VT_DATE) return -1;
    if (v.vt != VT_BSTR) return 0;
    if (vb_errpend) return 0;
    vb_s2date(v.u.s);
    if (vb_errpend) { vb_errpend = 0; return 0; }
    return -1;
}

vb_bool vb_isnumeric(vbvar v)
{
    if (v.vt == VT_DISPATCH && v.u.o) {        /* IsNumeric(Text1): the default property */
        v = vb_v_letval(v);
        if (vb_errpend) { vb_errpend = 0; return 0; }
    }
    switch (v.vt) {
    case VT_EMPTY: case VT_I2: case VT_I4: case VT_R4: case VT_R8: case VT_CY: case VT_BOOL: case VT_UI1:
    case VT_DECIMAL: return -1;
    case VT_BSTR: {
        vb_i32 save = vb_errpend; double d;
        extern vb_i32 vb_errpend;
        d = vb_s2d(v.u.s); (void)d;
        if (vb_errpend && !save) { vb_errpend = 0; return 0; }
        return -1;
    }
    default: return 0;
    }
}

/* Int, Fix, Abs of a Variant: the same subtype (VB6), a Decimal stays exact, Null gives Null.
   kind: 0 Int, 1 Fix, 2 Abs */
vbvar vb_v_intfix(vbvar v, vb_i32 kind)
{
    v = deref(v);
    switch (v.vt) {
    case VT_NULL: return v;
    case VT_EMPTY: return vb_v_i(0);
    case VT_DECIMAL: return kind == 2 ? vb_dec_abs(v) : vb_dec_int(v, kind);
    case VT_I2: case VT_BOOL: return kind == 2 ? mk(2, fabs((double)v.u.i)) : vb_v_i(v.u.i);
    case VT_I4: return kind == 2 ? mk(3, fabs((double)v.u.l)) : v;
    case VT_UI1: return v;
    case VT_CY: {
        double d = vb_cy2d(v.u.cy);
        return vb_v_c(vb_ccur(kind == 2 ? fabs(d) : kind == 1 ? vb_fix(d) : vb_int(d)));
    }
    default: {
        int ok; double d = num(v, &ok);
        if (!ok) return vb_v_empty();
        d = kind == 2 ? fabs(d) : kind == 1 ? vb_fix(d) : vb_int(d);
        if (v.vt == VT_R4) return vb_v_f((float)d);
        if (v.vt == VT_DATE) return vb_v_date(d);
        return vb_v_d(d);
    }
    }
}

vb_i16 vb_v_sgn(vbvar v)
{
    v = deref(v);
    if (v.vt == VT_DECIMAL) return (vb_i16)vb_dec_sgn(v);
    int ok; double d = num(v, &ok);
    return !ok ? 0 : d > 0 ? 1 : d < 0 ? -1 : 0;
}

/* Round(v, n): a Decimal rounds exactly (half to even); other values as Double */
vbvar vb_v_round(vbvar v, vb_i32 n)
{
    v = deref(v);
    if (v.vt == VT_NULL) return v;
    if (v.vt == VT_DECIMAL) return vb_dec_round(v, n);
    int ok; double d = num(v, &ok);
    return ok ? vb_v_d(vb_round(d, n)) : vb_v_empty();
}

/* Str$ of a Variant */
vbs vb_str_v(vbvar v)
{
    v = deref(v);
    switch (v.vt) {
    case VT_I2: case VT_BOOL: return vb_str_i(v.u.i);
    case VT_I4: return vb_str_i(v.u.l);
    case VT_UI1: return vb_str_i(v.u.b);
    case VT_R4: return vb_str_f(v.u.f);
    case VT_CY: return vb_str_c(v.u.cy);
    case VT_DECIMAL: { char b[100]; b[0] = ' '; vb_dec_text(v, 0, b + 1); return vb_sascii(b[1] == '-' ? b + 1 : b); }
    default: { int ok; double d = num(v, &ok); return ok ? vb_str_d(d) : vb_stmp(0); }
    }
}
