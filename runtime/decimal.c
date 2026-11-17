/* decimal.c - the Decimal subtype of the Variant (VT_DECIMAL) of LinBASIC Studio's runtime.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION).
 *
 * A Decimal is a 96-bit unsigned integer, a sign and a scale 0..28 (value = integer / 10^scale), as the
 * DECIMAL of OLE Automation that VB6 uses. In a vbvar: u.dec holds the three 32-bit words, the pad field
 * the scale (low byte) and the sign (0x8000). The arithmetic is exact; when a result does not fit, its
 * last digits are rounded half to even (banker's rounding), and error 6 when even the integer part does
 * not fit (|x| > 79228162514264337593543950335). */
#include "vbrt.h"
#include "vbrt_int.h"
#include <string.h>
#include <stdio.h>
#include <stdlib.h>
#include <math.h>

#define NL 8                                    /* limbs of the work numbers: 256 bits */
typedef struct { vb_u32 w[NL]; int scale; int neg; } big;

/* ---------------------------------------------------------------- 256-bit unsigned helpers */
static int b_zero(const vb_u32 *a) { for (int i = 0; i < NL; i++) if (a[i]) return 0; return 1; }
static int b_fits96(const vb_u32 *a) { for (int i = 3; i < NL; i++) if (a[i]) return 0; return 1; }
static int b_cmp(const vb_u32 *a, const vb_u32 *b)
{
    for (int i = NL - 1; i >= 0; i--) if (a[i] != b[i]) return a[i] < b[i] ? -1 : 1;
    return 0;
}
static vb_u32 b_mul_small(vb_u32 *a, vb_u32 m)      /* a *= m, returns the carry out of 256 bits */
{
    vb_u64 c = 0;
    for (int i = 0; i < NL; i++) { c += (vb_u64)a[i] * m; a[i] = (vb_u32)c; c >>= 32; }
    return (vb_u32)c;
}
static vb_u32 b_add_small(vb_u32 *a, vb_u32 m)
{
    vb_u64 c = m;
    for (int i = 0; i < NL && c; i++) { c += a[i]; a[i] = (vb_u32)c; c >>= 32; }
    return (vb_u32)c;
}
static vb_u32 b_div_small(vb_u32 *a, vb_u32 d)      /* a /= d, returns the remainder */
{
    vb_u64 r = 0;
    for (int i = NL - 1; i >= 0; i--) { r = (r << 32) | a[i]; a[i] = (vb_u32)(r / d); r %= d; }
    return (vb_u32)r;
}
static void b_add(vb_u32 *a, const vb_u32 *b)
{
    vb_u64 c = 0;
    for (int i = 0; i < NL; i++) { c += (vb_u64)a[i] + b[i]; a[i] = (vb_u32)c; c >>= 32; }
}
static void b_sub(vb_u32 *a, const vb_u32 *b)       /* a -= b, a >= b */
{
    vb_i64 c = 0;
    for (int i = 0; i < NL; i++) { c += (vb_i64)a[i] - b[i]; a[i] = (vb_u32)c; c = c < 0 ? -1 : 0; }
}
static void b_mul(vb_u32 *r, const vb_u32 *a, const vb_u32 *b)  /* r = a * b (both < 2^128) */
{
    vb_u32 t[NL]; memset(t, 0, sizeof t);
    for (int i = 0; i < 4; i++) {
        vb_u64 c = 0;
        for (int j = 0; j < 4 && i + j < NL; j++) {
            c += (vb_u64)a[i] * b[j] + t[i + j];
            t[i + j] = (vb_u32)c; c >>= 32;
        }
        if (i + 4 < NL) t[i + 4] = (vb_u32)c;
    }
    memcpy(r, t, sizeof t);
}
static void b_shl1(vb_u32 *a) { for (int i = NL - 1; i > 0; i--) a[i] = (a[i] << 1) | (a[i - 1] >> 31); a[0] <<= 1; }
/* q = n / d, r = n % d (d != 0), by shifting and subtracting */
static void b_divmod(const vb_u32 *n, const vb_u32 *d, vb_u32 *q, vb_u32 *r)
{
    vb_u32 qq[NL], rr[NL];
    memset(qq, 0, sizeof qq); memset(rr, 0, sizeof rr);
    for (int bit = NL * 32 - 1; bit >= 0; bit--) {
        b_shl1(rr);
        rr[0] |= (n[bit / 32] >> (bit % 32)) & 1;
        if (b_cmp(rr, d) >= 0) { b_sub(rr, d); qq[bit / 32] |= 1u << (bit % 32); }
    }
    if (q) memcpy(q, qq, sizeof qq);
    if (r) memcpy(r, rr, sizeof rr);
}

/* ---------------------------------------------------------------- big <-> vbvar */
static big from_var(vbvar v)
{
    big b; memset(&b, 0, sizeof b);
    b.w[0] = v.u.dec.lo; b.w[1] = v.u.dec.mid; b.w[2] = v.u.dec.hi;
    b.scale = v.pad & 0xFF; b.neg = (v.pad & 0x8000) != 0;
    return b;
}

/* rounds b to fit 96 bits and a scale <= maxscale (half to even); 0: it does not fit (overflow) */
static int fit(big *b, int maxscale)
{
    vb_u32 rem = 0; int sticky = 0, cut = 0;
    while ((!b_fits96(b->w) || b->scale > maxscale) && b->scale > 0) {
        if (cut) sticky |= rem != 0;
        rem = b_div_small(b->w, 10);
        b->scale--; cut = 1;
    }
    if (cut && (rem > 5 || (rem == 5 && (sticky || (b->w[0] & 1))))) {
        b_add_small(b->w, 1);
        if (!b_fits96(b->w)) {                   /* 999...9 rounded up: one digit more */
            if (b->scale == 0) return 0;
            b_div_small(b->w, 10); b->scale--;
        }
    }
    return b_fits96(b->w);
}

static vbvar to_var(big b)
{
    vbvar v = vb_v_empty();
    if (!fit(&b, 28)) { vb_raise(6); return v; }
    v.vt = VT_DECIMAL;
    v.u.dec.lo = b.w[0]; v.u.dec.mid = b.w[1]; v.u.dec.hi = b.w[2];
    v.pad = (vb_u16)(b.scale | (b.neg && !b_zero(b.w) ? 0x8000 : 0));
    return v;
}

/* ---------------------------------------------------------------- text */
/* the digits of b, with a '.' (or dec) before the last `scale` digits; a leading 0 before the point */
static void dec_text(big b, char dec, char *out, int lead0)
{
    char d[100]; int n = 0;
    vb_u32 w[NL]; memcpy(w, b.w, sizeof w);
    int neg = b.neg && !b_zero(w);
    while (b.scale > 0) {                        /* VB6 shows no zeros at the end: 1.10 -> 1.1 */
        vb_u32 t[NL]; memcpy(t, w, sizeof t);
        if (b_div_small(t, 10)) break;
        memcpy(w, t, sizeof t); b.scale--;
    }
    do { d[n++] = (char)('0' + b_div_small(w, 10)); } while (!b_zero(w));
    while (n <= b.scale) d[n++] = '0';
    char *p = out;
    if (neg) *p++ = '-';
    for (int i = n - 1; i >= 0; i--) {
        if (i == n - 1 && i == b.scale && i > 0 && d[i] == '0' && !lead0) { *p++ = dec; continue; }  /* Str: .5 */
        *p++ = d[i];
        if (i == b.scale && i > 0) *p++ = dec;
    }
    *p = 0;
}

/* VarBstrFromDec: the regional decimal separator and a 0 before it (CStr, Print); else Str$ / Write #: '.'
   and no 0 before it (.5) */
void vb_dec_text(vbvar v, int locale, char *out)
{
    dec_text(from_var(v), locale ? vb_lc_dec : '.', out, locale);
}
vbs vb_dec2s(vbvar v) { char b[100]; vb_dec_text(v, 1, b); return vb_sascii(b); }

/* text -> Decimal: digits, the regional decimal separator, group separators, a sign, an exponent */
static int parse(vbs s, big *b, char dch, char gch)
{
    memset(b, 0, sizeof *b);
    vb_i32 i = 0, n = s.len;
    while (i < n && (s.p[i] == ' ' || s.p[i] == '\t')) i++;
    if (i < n && (s.p[i] == '-' || s.p[i] == '+')) { b->neg = s.p[i] == '-'; i++; }
    int digits = 0, frac = 0, point = 0, dropped = 0, sticky = 0, any = 0;
    vb_u32 last = 0;
    for (; i < n; i++) {
        vb_wchar c = s.p[i];
        if (c >= '0' && c <= '9') {
            any = 1;
            if (!digits && c == '0') { if (point) frac++; continue; }     /* leading zeros */
            if (digits < 40) {
                b_mul_small(b->w, 10); b_add_small(b->w, (vb_u32)(c - '0'));
                digits++; if (point) frac++;
            } else {                              /* beyond 40 digits: only their rounding counts */
                if (dropped == 0) last = (vb_u32)(c - '0'); else sticky |= c != '0';
                dropped++; if (!point) frac--;
            }
            continue;
        }
        if (c == (vb_wchar)(unsigned char)dch && !point) { point = 1; continue; }
        if (gch && c == (vb_wchar)(unsigned char)gch && !point && any) continue;
        break;
    }
    if (!any) return 0;
    int exp = 0;
    if (i < n && (s.p[i] == 'e' || s.p[i] == 'E' || s.p[i] == 'd' || s.p[i] == 'D')) {
        int j = i + 1, en = 0, eneg = 0, ed = 0;
        if (j < n && (s.p[j] == '-' || s.p[j] == '+')) { eneg = s.p[j] == '-'; j++; }
        while (j < n && s.p[j] >= '0' && s.p[j] <= '9') { if (en < 10000) en = en * 10 + (s.p[j] - '0'); j++; ed = 1; }
        if (!ed) return 0;
        exp = eneg ? -en : en; i = j;
    }
    while (i < n && (s.p[i] == ' ' || s.p[i] == '\t')) i++;
    if (i != n) return 0;
    if (dropped) {                                /* the first dropped digit and the rest decide the rounding */
        if (last > 5 || (last == 5 && (sticky || (b->w[0] & 1)))) b_add_small(b->w, 1);
    }
    int scale = frac - exp;
    while (scale < 0) {
        if (b_zero(b->w)) { scale = 0; break; }
        if (b_mul_small(b->w, 10) || !b_fits96(b->w)) return -1;
        scale++;
    }
    b->scale = scale;
    if (b->scale > 200) {                         /* far below the smallest Decimal: 0 */
        memset(b->w, 0, sizeof b->w); b->scale = 0;
    }
    return 1;
}

/* ---------------------------------------------------------------- conversions */
vbvar vb_dec_from_i64(vb_i64 x)
{
    big b; memset(&b, 0, sizeof b);
    vb_u64 m = x < 0 ? (vb_u64)0 - (vb_u64)x : (vb_u64)x;
    b.w[0] = (vb_u32)m; b.w[1] = (vb_u32)(m >> 32); b.neg = x < 0;
    return to_var(b);
}

vbvar vb_dec_from_cy(vb_cy c)
{
    vbvar v = vb_dec_from_i64(c);
    v.pad = (vb_u16)((v.pad & 0x8000) | 4);
    return v;
}

/* VarDecFromR8 / R4: the value with 15 (a Single: 7) significant digits, without the zeros at the end */
vbvar vb_dec_from_double(double d, int digits)
{
    if (d != d || fabs(d) > 7.9228162514264338e28 || (d - d) != (d - d)) { vb_raise(6); return vb_v_empty(); }
    char buf[64], m[40]; int k = 0, e;
    snprintf(buf, sizeof buf, "%.*e", digits - 1, d);
    char *p = buf; int neg = 0;
    if (*p == '-') { neg = 1; p++; }
    for (; *p && *p != 'e'; p++) if (*p >= '0' && *p <= '9') m[k++] = *p;
    m[k] = 0;
    e = atoi(p + 1);
    big b; memset(&b, 0, sizeof b);
    for (int i = 0; i < k; i++) { b_mul_small(b.w, 10); b_add_small(b.w, (vb_u32)(m[i] - '0')); }
    int scale = (k - 1) - e;
    while (scale < 0) { b_mul_small(b.w, 10); scale++; }
    b.scale = scale; b.neg = neg;
    if (!fit(&b, 28)) { vb_raise(6); return vb_v_empty(); }
    while (b.scale > 0 && !b_zero(b.w)) {        /* 1.10000000000000 -> 1.1 */
        vb_u32 t[NL]; memcpy(t, b.w, sizeof t);
        if (b_div_small(t, 10)) break;
        memcpy(b.w, t, sizeof t); b.scale--;
    }
    if (b_zero(b.w)) b.scale = 0;
    return to_var(b);
}

vbvar vb_dec_from_str(vbs s)
{
    big b;
    int r = parse(s, &b, vb_lc_dec, vb_lc_grp);
    if (r < 0) { vb_raise(6); return vb_v_empty(); }
    if (r == 0) {                                 /* &H10, True ...: as CDbl takes them */
        vb_i32 save = vb_errpend;
        double d = vb_s2d(s);
        if (vb_errpend && !save) return vb_v_empty();
        return vb_dec_from_double(d, 15);
    }
    return to_var(b);
}

double vb_dec2d(vbvar v)
{
    char b[100];
    dec_text(from_var(v), '.', b, 1);
    return strtod(b, 0);
}

/* rounds to `scale` decimals (half to even) */
static big round_to(big b, int scale)
{
    if (b.scale <= scale) return b;
    vb_u32 rem = 0; int sticky = 0, cut = 0;
    while (b.scale > scale) {
        if (cut) sticky |= rem != 0;
        rem = b_div_small(b.w, 10); b.scale--; cut = 1;
    }
    if (rem > 5 || (rem == 5 && (sticky || (b.w[0] & 1)))) b_add_small(b.w, 1);
    return b;
}

/* the integer of a Decimal rounded half to even (CInt, CLng, CByte, CCur * 10000); 0: too big */
static int to_i64(big b, vb_i64 *out)
{
    b = round_to(b, 0);
    if (b.w[2] || b.w[3] || (b.w[1] & 0x80000000u)) {
        if (b.neg && !b.w[2] && !b.w[3] && b.w[1] == 0x80000000u && !b.w[0]) { *out = (vb_i64)((vb_u64)1 << 63); return 1; }
        return 0;
    }
    vb_u64 m = ((vb_u64)b.w[1] << 32) | b.w[0];
    *out = b.neg ? -(vb_i64)m : (vb_i64)m;
    return 1;
}

vb_i64 vb_dec2i64(vbvar v, vb_i64 lo, vb_i64 hi)
{
    vb_i64 r;
    if (!to_i64(from_var(v), &r) || r < lo || r > hi) { vb_raise(6); return 0; }
    return r;
}

vb_cy vb_dec2cy(vbvar v)
{
    big b = round_to(from_var(v), 4);
    while (b.scale < 4) { b_mul_small(b.w, 10); b.scale++; }
    b.scale = 0;
    vb_i64 r;
    if (!to_i64(b, &r)) { vb_raise(6); return 0; }
    return r;
}

/* any Variant -> its Decimal (CDec) */
vbvar vb_cdec(vbvar v)
{
    if (v.vt == VT_DISPATCH) v = vb_v_letval(v);
    switch (v.vt) {
    case VT_DECIMAL: return v;
    case VT_EMPTY: return vb_dec_from_i64(0);
    case VT_NULL: vb_raise(94); return vb_v_empty();
    case VT_I2: case VT_BOOL: return vb_dec_from_i64(v.u.i);
    case VT_I4: return vb_dec_from_i64(v.u.l);
    case VT_UI1: return vb_dec_from_i64(v.u.b);
    case VT_CY: return vb_dec_from_cy(v.u.cy);
    case VT_R4: return vb_dec_from_double(v.u.f, 7);
    case VT_R8: case VT_DATE: return vb_dec_from_double(v.u.d, 15);
    case VT_BSTR: return vb_dec_from_str(v.u.s);
    default: vb_raise(13); return vb_v_empty();
    }
}

/* ---------------------------------------------------------------- arithmetic */
static void align(big *a, big *b)                 /* the same scale (the larger one) */
{
    while (a->scale < b->scale) { b_mul_small(a->w, 10); a->scale++; }
    while (b->scale < a->scale) { b_mul_small(b->w, 10); b->scale++; }
}

static vbvar addsub(vbvar x, vbvar y, int sub)
{
    big a = from_var(x), b = from_var(y);
    if (sub) b.neg = !b.neg;
    align(&a, &b);
    big r; memset(&r, 0, sizeof r); r.scale = a.scale;
    if (a.neg == b.neg) { memcpy(r.w, a.w, sizeof r.w); b_add(r.w, b.w); r.neg = a.neg; }
    else if (b_cmp(a.w, b.w) >= 0) { memcpy(r.w, a.w, sizeof r.w); b_sub(r.w, b.w); r.neg = a.neg; }
    else { memcpy(r.w, b.w, sizeof r.w); b_sub(r.w, a.w); r.neg = b.neg; }
    return to_var(r);
}
vbvar vb_dec_add(vbvar a, vbvar b) { return addsub(a, b, 0); }
vbvar vb_dec_sub(vbvar a, vbvar b) { return addsub(a, b, 1); }

vbvar vb_dec_mul(vbvar x, vbvar y)
{
    big a = from_var(x), b = from_var(y), r; memset(&r, 0, sizeof r);
    b_mul(r.w, a.w, b.w);
    r.scale = a.scale + b.scale; r.neg = a.neg != b.neg;
    return to_var(r);
}

vbvar vb_dec_div(vbvar x, vbvar y)
{
    big a = from_var(x), b = from_var(y);
    if (b_zero(b.w)) { vb_raise(11); return vb_v_empty(); }      /* 0 / 0 too (VB6) */
    big r; memset(&r, 0, sizeof r); r.neg = a.neg != b.neg;
    int sc = a.scale - b.scale;
    vb_u32 n[NL], q[NL], rem[NL]; memcpy(n, a.w, sizeof n);
    while (sc < 0) { b_mul_small(n, 10); sc++; }
    b_divmod(n, b.w, q, rem);
    if (!b_fits96(q)) { vb_raise(6); return vb_v_empty(); }
    while (!b_zero(rem) && sc < 28) {             /* more digits while they fit */
        vb_u32 q10[NL], r10[NL], d[NL], m[NL];
        memcpy(q10, q, sizeof q10); b_mul_small(q10, 10);
        vb_u32 hi[NL]; memcpy(hi, q10, sizeof hi); b_add_small(hi, 9);
        if (!b_fits96(hi)) break;                 /* (VarDecDiv: a digit more only if any digit would fit) */
        memcpy(r10, rem, sizeof r10); b_mul_small(r10, 10);
        b_divmod(r10, b.w, d, m);
        b_add(q10, d);
        memcpy(q, q10, sizeof q); memcpy(rem, m, sizeof rem); sc++;
    }
    if (!b_zero(rem)) {                           /* the rest: rounded half to even */
        vb_u32 twice[NL]; memcpy(twice, rem, sizeof twice); b_mul_small(twice, 2);
        int c = b_cmp(twice, b.w);
        if (c > 0 || (c == 0 && (q[0] & 1))) b_add_small(q, 1);
    }
    memcpy(r.w, q, sizeof r.w); r.scale = sc;
    return to_var(r);
}

int vb_dec_cmp(vbvar x, vbvar y)
{
    big a = from_var(x), b = from_var(y);
    int za = b_zero(a.w), zb = b_zero(b.w);
    if (za && zb) return 0;
    if (a.neg != b.neg && !(za || zb)) return a.neg ? -1 : 1;
    if (za) return b.neg ? 1 : -1;
    if (zb) return a.neg ? -1 : 1;
    align(&a, &b);
    int c = b_cmp(a.w, b.w);
    return a.neg ? -c : c;
}

vbvar vb_dec_neg(vbvar v) { if (v.u.dec.lo || v.u.dec.mid || v.u.dec.hi) v.pad ^= 0x8000; return v; }
vbvar vb_dec_abs(vbvar v) { v.pad &= 0xFF; return v; }
int vb_dec_sgn(vbvar v) { if (!(v.u.dec.lo || v.u.dec.mid || v.u.dec.hi)) return 0; return v.pad & 0x8000 ? -1 : 1; }

/* Int (fix = 0: down) and Fix (toward 0) */
vbvar vb_dec_int(vbvar v, int fix)
{
    big b = from_var(v);
    int frac = 0;
    while (b.scale > 0) { frac |= b_div_small(b.w, 10) != 0; b.scale--; }
    if (frac && b.neg && !fix) b_add_small(b.w, 1);
    return to_var(b);
}

/* Round(x, n): half to even */
vbvar vb_dec_round(vbvar v, vb_i32 n)
{
    if (n < 0) { vb_raise(5); return vb_v_empty(); }
    return to_var(round_to(from_var(v), n));
}

/* the 16 bytes of a DECIMAL (Put # of a Variant): wReserved (here VT_DECIMAL), scale, sign, Hi32, Lo64 */
void vb_dec_bytes(vbvar v, unsigned char *out)
{
    memset(out, 0, 16);
    out[0] = VT_DECIMAL;
    out[2] = (unsigned char)(v.pad & 0xFF); out[3] = (v.pad & 0x8000) ? 0x80 : 0;
    memcpy(out + 4, &v.u.dec.hi, 4); memcpy(out + 8, &v.u.dec.lo, 4); memcpy(out + 12, &v.u.dec.mid, 4);
}
vbvar vb_dec_from_bytes(const unsigned char *in)
{
    vbvar v = vb_v_empty();
    v.vt = VT_DECIMAL;
    v.pad = (vb_u16)(in[2] | (in[3] & 0x80 ? 0x8000 : 0));
    memcpy(&v.u.dec.hi, in + 4, 4); memcpy(&v.u.dec.lo, in + 8, 4); memcpy(&v.u.dec.mid, in + 12, 4);
    return v;
}

vbvar vb_dec_from_ascii(const char *s)
{
    vb_i32 m = vb_mark();
    big b;
    int r = parse(vb_sascii(s), &b, '.', 0);
    vb_release(m);
    if (r <= 0 || !fit(&b, 28)) return vb_v_empty();
    return to_var(b);
}
