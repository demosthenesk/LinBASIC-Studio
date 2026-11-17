/* intl.c - international text: the ANSI code pages of Windows (single and double byte) with the "best fit"
 * of WideCharToMultiByte, the case of letters (UCase / LCase), StrConv with a LocaleID (vbWide, vbNarrow,
 * vbKatakana, vbHiragana) and the comparison of Option Compare Text (StrComp). Measured on VB6 SP6, Greek
 * Windows 7 (tests/lang/t_intl.bas). The tables: intl_tables.inc (tools/mkintl.py).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"

typedef struct { int cp; int dbcs; const vb_u16 *single; const unsigned char *lead; const vb_u16 *rows; vb_u16 defchar;
                 const vb_u16 *pref; int npref; } cpdef;
#include "intl_tables.inc"

#define NTRAIL 191                              /* trail bytes 0x40 .. 0xFE */

static const cpdef *cp_find(int cp)
{
    for (int i = 0; cpdefs[i].cp; i++) if (cpdefs[i].cp == cp) return &cpdefs[i];
    return 0;
}

/* ---------------------------------------------------------------- bytes -> Unicode */
int vb_cp_decode(int cp, const unsigned char *b, int n, vb_wchar *out)
{
    const cpdef *d = cp_find(cp);
    int k = 0;
    for (int i = 0; i < n; i++) {
        unsigned c = b[i];
        vb_wchar w;
        if (c < 0x80 || !d) w = (vb_wchar)c;
        else if (!d->dbcs || !d->lead[c - 0x80]) w = d->single[c - 0x80];
        else {                                  /* a lead byte: the pair (any trail byte goes with it) */
            if (i + 1 >= n) w = d->defchar;
            else {
                unsigned t = b[++i];
                w = 0;
                if (t >= 0x40 && t <= 0xFE) w = d->rows[(d->lead[c - 0x80] - 1) * NTRAIL + (t - 0x40)];
                if (!w) w = d->defchar;
            }
        }
        if (out) out[k] = w;
        k++;
    }
    return k;
}

/* ---------------------------------------------------------------- Unicode -> bytes */
typedef struct { vb_u16 u, code; } rev;
static rev *revs[32];
static int nrevs[32];
static int rev_cmp(const void *a, const void *b)
{
    const rev *x = (const rev *)a, *y = (const rev *)b;
    return x->u != y->u ? (int)x->u - (int)y->u : (int)x->code - (int)y->code;
}
static int cp_index(const cpdef *d) { return (int)(d - cpdefs); }
static const rev *rev_table(const cpdef *d, int *n)
{
    int k = cp_index(d);
    if (!revs[k]) {
        int cap = 128 + (d->dbcs ? 126 * NTRAIL : 0), m = 0;
        rev *r = (rev *)malloc(sizeof(rev) * (size_t)cap);
        if (!r) { *n = 0; return 0; }
        for (int c = 0x80; c < 0x100; c++) {
            if (d->dbcs && d->lead[c - 0x80]) continue;
            vb_u16 u = d->single[c - 0x80];
            if (u >= 0x80 && !(u >= 0xF8F0 && u <= 0xF8FF && d->dbcs)) { r[m].u = u; r[m].code = (vb_u16)c; m++; }
            else if (u == c) { r[m].u = u; r[m].code = (vb_u16)c; m++; }
        }
        if (d->dbcs)
            for (int c = 0x80; c < 0x100; c++) {
                int li = d->lead[c - 0x80];
                if (!li) continue;
                for (int t = 0; t < NTRAIL; t++) {
                    vb_u16 u = d->rows[(li - 1) * NTRAIL + t];
                    if (u) { r[m].u = u; r[m].code = (vb_u16)((c << 8) | (t + 0x40)); m++; }
                }
            }
        /* a character with several codes: the lowest, or the one Windows prefers */
        qsort(r, (size_t)m, sizeof(rev), rev_cmp);
        int w = 0;
        for (int i = 0; i < m; i++) if (!w || r[w - 1].u != r[i].u) r[w++] = r[i];
        for (int p = 0; p < d->npref; p++)
            for (int i = 0; i < w; i++) if (r[i].u == d->pref[2 * p]) { r[i].code = d->pref[2 * p + 1]; break; }
        revs[k] = r; nrevs[k] = w;
    }
    *n = nrevs[k];
    return revs[k];
}
static int cp_lookup(const cpdef *d, vb_wchar u)
{
    int n; const rev *r = rev_table(d, &n);
    int lo = 0, hi = n - 1;
    while (lo <= hi) {
        int mid = (lo + hi) / 2;
        if (r[mid].u == u) return r[mid].code;
        if (r[mid].u < u) lo = mid + 1; else hi = mid - 1;
    }
    return -1;
}

/* the best fit of Windows for a character the code page does not have (0: none) */
static vb_wchar best_fit(int cp, vb_wchar u)
{
    int latin1 = cp == 1250 || cp == 1251 || cp == 1252 || cp == 1253 || cp == 1254 || cp == 932 || cp == 949 || cp == 950;
    int latinx = cp == 1250 || cp == 1251 || cp == 1252 || cp == 1253 || cp == 1254;
    if (u >= 0xC0 && u < 0x100 && latin1 && latin_base[u - 0xC0]) return latin_base[u - 0xC0];
    if (u >= 0x100 && u < 0x250 && latinx && latin_base[u - 0xC0]) return latin_base[u - 0xC0];
    if (u == 0x131 && (cp == 1252 || cp == 1250)) return 'i';          /* dotless i */
    if (u == 0xDF && (cp == 932 || cp == 950)) return 's';             /* sharp s */
    if (u == 0x30FB && (cp == 1252 || cp == 1250 || cp == 1254)) return 0xB7;   /* katakana middle dot */
    if (u >= 0xFF01 && u <= 0xFF5E && latinx) return (vb_wchar)(u - 0xFEE0);    /* fullwidth ASCII */
    if (cp == 1252 || cp == 1250 || cp == 1254) {                       /* Greek letters of the old PC fonts */
        if (u == 0x3B2) return 0xDF;            /* β -> ß */
        if (u == 0x3BC) return 0xB5;            /* μ -> µ */
    }
    if (cp == 1252) {
        static const vb_u16 g[][2] = { {0x393,'G'}, {0x398,'T'}, {0x3A3,'S'}, {0x3A6,'F'}, {0x3A9,'O'}, {0x3B1,'a'},
                                       {0x3B4,'d'}, {0x3B5,'e'}, {0x3C0,'p'}, {0x3C3,'s'}, {0x3C4,'t'}, {0x3C6,'f'}, {0,0} };
        for (int i = 0; g[i][0]; i++) if (g[i][0] == u) return g[i][1];
    }
    if (cp == 1256) {                                                   /* Persian letters -> the Arabic ones */
        if (u == 0x6CC) return 0x64A;
        if (u == 0x6A9) return 0x643;
    }
    return 0;
}

/* the code (1 or 2 bytes, as a number) of a character; -1 if the code page cannot write it */
int vb_cp_code(int cp, vb_wchar u, int bestfit)
{
    if (u < 0x80) return u;
    const cpdef *d = cp_find(cp);
    if (!d) return u < 0x100 ? u : -1;
    int c = cp_lookup(d, u);
    if (c >= 0 || !bestfit) return c;
    vb_wchar f = best_fit(cp, u);
    if (f) { c = f < 0x80 ? f : cp_lookup(d, f); if (c >= 0) return c; }
    return -1;
}

int vb_cp_encode(int cp, const vb_wchar *w, int n, unsigned char *out)
{
    int k = 0;
    for (int i = 0; i < n; i++) {
        int c = vb_cp_code(cp, w[i], 1);
        if (c < 0) c = '?';
        if (c > 0xFF) { if (out) { out[k] = (unsigned char)(c >> 8); out[k + 1] = (unsigned char)c; } k += 2; }
        else { if (out) out[k] = (unsigned char)c; k++; }
    }
    return k;
}

int vb_cp_known(int cp) { return cp_find(cp) != 0; }
int vb_cp_isdbcs(int cp) { const cpdef *d = cp_find(cp); return d && d->dbcs; }
int vb_cp_islead(int cp, unsigned char c)
{
    const cpdef *d = cp_find(cp);
    return d && d->dbcs && c >= 0x80 && d->lead[c - 0x80];
}

/* ---------------------------------------------------------------- locales */
/* the ANSI code page of a locale; 0: the system's; -1: no such locale (error 5) */
int vb_lcid_cp(vb_i32 lcid)
{
    if (lcid < 0 || lcid > 0xFFFFF) return -1;
    int lang = lcid & 0x3FF, sub = (lcid >> 10) & 0x3F;
    if (lang == 0 || sub == 0) return 0;        /* neutral, system default, user default (VB6: 0, 1, 1024, 2048) */
    switch (lang) {
    case 0x01: case 0x29: case 0x20: case 0x8C: case 0x92: return 1256;     /* Arabic, Persian, Urdu ... */
    case 0x02: case 0x19: case 0x22: case 0x23: case 0x2F: case 0x3F: case 0x40: case 0x44: case 0x50: case 0x6D:
    case 0x85: return 1251;                     /* Bulgarian, Russian, Ukrainian, Belarusian, Macedonian ... */
    case 0x04: return (sub == 1 || sub == 3 || sub == 5) ? 950 : 936;   /* Chinese: Taiwan, Hong Kong, Macao */
    case 0x05: case 0x0E: case 0x15: case 0x18: case 0x1B: case 0x1C: case 0x24: return 1250;
    case 0x1A: return sub == 3 || sub == 7 ? 1251 : 1250;              /* Croatian / Serbian (Cyrillic) */
    case 0x2C: case 0x43: return sub == 2 ? 1251 : 1254;               /* Azeri, Uzbek */
    case 0x08: return 1253;
    case 0x0D: return 1255;
    case 0x11: return 932;
    case 0x12: return 949;
    case 0x1E: return 874;
    case 0x1F: return 1254;
    case 0x25: case 0x26: case 0x27: return 1257;
    case 0x2A: return 1258;
    case 0x03: case 0x06: case 0x07: case 0x09: case 0x0A: case 0x0B: case 0x0C: case 0x0F: case 0x10: case 0x13:
    case 0x14: case 0x16: case 0x17: case 0x1D: case 0x21: case 0x2D: case 0x2E: case 0x36: case 0x38: case 0x3B:
    case 0x3C: case 0x3E: case 0x41: case 0x52: case 0x56: case 0x62: case 0x6E: case 0x83: case 0x88:
        return 1252;
    }
    return -1;
}

/* ---------------------------------------------------------------- case
 * Windows changes a letter only when the pair goes both ways: ς, ı, ǅ, İ keep their case */
static int bsearch16(const vb_u16 *t, int n, vb_wchar c)
{
    int lo = 0, hi = n - 1;
    while (lo <= hi) {
        int mid = (lo + hi) / 2;
        if (t[mid] == c) return mid;
        if (t[mid] < c) lo = mid + 1; else hi = mid - 1;
    }
    return -1;
}
vb_wchar vb_wupper(vb_wchar c)
{
    if (c < 0x80) return c >= 'a' && c <= 'z' ? (vb_wchar)(c - 32) : c;
    int i = bsearch16(case_lower, NCASE, c);
    return i >= 0 ? case_upper[i] : c;
}
vb_wchar vb_wlower(vb_wchar c)
{
    if (c < 0x80) return c >= 'A' && c <= 'Z' ? (vb_wchar)(c + 32) : c;
    int i = bsearch16(case_u2, NCASE, c);
    return i >= 0 ? case_l2[i] : c;
}

/* ---------------------------------------------------------------- StrConv */
/* the text through the code page and back (what LCMapStringA sees) */
static vbs round_trip(vbs s, int cp)
{
    unsigned char *b = (unsigned char *)malloc((size_t)s.len * 2 + 2);
    if (!b) { vb_raise(7); return vb_stmp(0); }
    int n = vb_cp_encode(cp, s.p, s.len, b);
    vbs r = vb_stmp(n);
    r.len = vb_cp_decode(cp, b, n, r.p);
    free(b);
    return r;
}

/* halfwidth katakana FF61..FF9F -> fullwidth */
static const vb_u16 half_kana[63] = {
    0x3002,0x300C,0x300D,0x3001,0x30FB,0x30F2,0x30A1,0x30A3,0x30A5,0x30A7,0x30A9,0x30E3,0x30E5,0x30E7,0x30C3,0x30FC,
    0x30A2,0x30A4,0x30A6,0x30A8,0x30AA,0x30AB,0x30AD,0x30AF,0x30B1,0x30B3,0x30B5,0x30B7,0x30B9,0x30BB,0x30BD,0x30BF,
    0x30C1,0x30C4,0x30C6,0x30C8,0x30CA,0x30CB,0x30CC,0x30CD,0x30CE,0x30CF,0x30D2,0x30D5,0x30D8,0x30DB,0x30DE,0x30DF,
    0x30E0,0x30E1,0x30E2,0x30E4,0x30E6,0x30E8,0x30E9,0x30EA,0x30EB,0x30EC,0x30ED,0x30EF,0x30F3,0x309B,0x309C };
static const struct { vb_u16 half, full; } wide_misc[] = {
    { 0xA2, 0xFFE0 }, { 0xA3, 0xFFE1 }, { 0xAC, 0xFFE2 }, { 0xAF, 0xFFE3 }, { 0xA6, 0xFFE4 }, { 0xA5, 0xFFE5 },
    { 0x20A9, 0xFFE6 }, { 0, 0 } };

/* katakana that take the voiced mark (カ -> ガ), and the semi-voiced one (ハ -> パ) */
static int voiceable(vb_wchar k)
{
    static const vb_u16 v[] = { 0x30AB,0x30AD,0x30AF,0x30B1,0x30B3,0x30B5,0x30B7,0x30B9,0x30BB,0x30BD,0x30BF,0x30C1,0x30C4,
                                0x30C6,0x30C8,0x30CF,0x30D2,0x30D5,0x30D8,0x30DB, 0 };
    for (int i = 0; v[i]; i++) if (v[i] == k) return 1;
    return 0;
}
static int semivoiceable(vb_wchar k) { return k == 0x30CF || k == 0x30D2 || k == 0x30D5 || k == 0x30D8 || k == 0x30DB; }

static vbs to_wide(vbs s)
{
    vbs r = vb_stmp(s.len);
    int k = 0;
    for (vb_i32 i = 0; i < s.len; i++) {
        vb_wchar c = s.p[i], w = c;
        if (c == 0x20) w = 0x3000;
        else if (c > 0x20 && c < 0x7F) w = (vb_wchar)(c + 0xFEE0);
        else if (c >= 0xFF61 && c <= 0xFF9F) {
            w = half_kana[c - 0xFF61];
            vb_wchar nx = i + 1 < s.len ? s.p[i + 1] : 0;
            if (nx == 0xFF9E && (w == 0x30A6 || voiceable(w))) { w = w == 0x30A6 ? 0x30F4 : (vb_wchar)(w + 1); i++; }
            else if (nx == 0xFF9F && semivoiceable(w)) { w = (vb_wchar)(w + 2); i++; }
        } else for (int j = 0; wide_misc[j].half; j++) if (wide_misc[j].half == c) w = wide_misc[j].full;
        r.p[k++] = w;
    }
    r.len = k; r.p[k] = 0;
    return r;
}

static vbs to_narrow(vbs s)
{
    vbs r = vb_stmp(s.len * 2);
    int k = 0;
    for (vb_i32 i = 0; i < s.len; i++) {
        vb_wchar c = s.p[i], w = c, w2 = 0;
        if (c == 0x3000) w = 0x20;
        else if (c >= 0xFF01 && c <= 0xFF5E) w = (vb_wchar)(c - 0xFEE0);
        else if (c >= 0x3001 && c <= 0x30FE) {
            vb_wchar base = c, mark = 0;
            int direct = 0;
            for (int j = 0; j < 63; j++) if (half_kana[j] == c) direct = 1;
            if (!direct) {                      /* voiced / semi-voiced: the plain one + the mark */
                if (c == 0x30F4) { base = 0x30A6; mark = 0xFF9E; }
                else if (voiceable((vb_wchar)(c - 1))) { base = (vb_wchar)(c - 1); mark = 0xFF9E; }
                else if (semivoiceable((vb_wchar)(c - 2))) { base = (vb_wchar)(c - 2); mark = 0xFF9F; }
            }
            for (int j = 0; j < 63; j++) if (half_kana[j] == base) { w = (vb_wchar)(0xFF61 + j); w2 = mark; break; }
        } else for (int j = 0; wide_misc[j].half; j++) if (wide_misc[j].full == c) w = wide_misc[j].half;
        r.p[k++] = w;
        if (w2) r.p[k++] = w2;
    }
    r.len = k; r.p[k] = 0;
    return r;
}

static vbs kana(vbs s, int to_kata)
{
    vbs r = vb_stmp(s.len);
    for (vb_i32 i = 0; i < s.len; i++) {
        vb_wchar c = s.p[i];
        if (to_kata && ((c >= 0x3041 && c <= 0x3096) || c == 0x309D || c == 0x309E)) c = (vb_wchar)(c + 0x60);
        else if (!to_kata && ((c >= 0x30A1 && c <= 0x30F6) || c == 0x30FD || c == 0x30FE)) c = (vb_wchar)(c - 0x60);
        r.p[i] = c;
    }
    return r;
}

static vbs case_map(vbs s, int how)
{
    vbs r = vb_stmp(s.len);
    int start = 1;
    for (vb_i32 i = 0; i < s.len; i++) {
        vb_wchar c = s.p[i];
        if (how == 1) r.p[i] = vb_wupper(c);
        else if (how == 2) r.p[i] = vb_wlower(c);
        else {                                  /* proper: the first letter of each word */
            int ws = c == ' ' || c == '\t' || c == '\n' || c == '\r' || c == '\v' || c == '\f' || c == 0;
            r.p[i] = start && !ws ? vb_wupper(c) : vb_wlower(c);
            start = ws;
        }
    }
    return r;
}

/* StrConv(s, conv[, lcid]) */
vbs vb_strconv_l(vbs s, vb_i32 conv, vb_i32 lcid, vb_i32 has_lcid)
{
    int far = conv & 0xF00;
    if ((conv & ~0xFFF) || (conv & 0xC0) == 0xC0 || (conv & 0x0C) == 0x0C || (conv & 0x30) == 0x30 || (far & (far - 1)) ||
        ((conv & 0xC0) && (conv & 0x3F))) { vb_raise(5); return vb_stmp(0); }
    int cp = vb_codepage, lang = vb_lcid & 0x3FF;
    if (has_lcid) {
        int c = vb_lcid_cp(lcid);
        if (c < 0) { vb_raise(5); return vb_stmp(0); }
        if (c) { cp = c; lang = lcid & 0x3FF; }
    }
    if ((conv & 0x0C) && lang != 0x11 && lang != 0x04 && lang != 0x12) { vb_raise(5); return vb_stmp(0); }  /* wide/narrow: CJK */
    if ((conv & 0x30) && lang != 0x11) { vb_raise(5); return vb_stmp(0); }                               /* kana: Japanese */
    if (conv & 64) {                            /* vbUnicode: the bytes in the code page to characters */
        vb_i32 n = vb_sbytes(s);
        unsigned char *b = (unsigned char *)malloc((size_t)n + 1);
        if (!b) { vb_raise(7); return vb_stmp(0); }
        for (vb_i32 i = 0; i < n; i++) b[i] = (unsigned char)vb_sbyte(s, i);
        vbs r = vb_stmp(n);
        r.len = vb_cp_decode(cp, b, n, r.p); r.p[r.len] = 0;
        free(b);
        return r;
    }
    if (conv & 128) {                           /* vbFromUnicode: the characters as bytes of the code page */
        unsigned char *b = (unsigned char *)malloc((size_t)s.len * 2 + 2);
        if (!b) { vb_raise(7); return vb_stmp(0); }
        int n = vb_cp_encode(cp, s.p, s.len, b);
        vbs r = vb_sfrombytes((const char *)b, n);
        free(b);
        return r;
    }
    if (!(conv & 0x3F)) return vb_scopy(s);
    /* case, width, kana: through the code page (as LCMapStringA), changed, and through it again */
    vbs r = round_trip(s, cp);
    if (conv & 3) r = case_map(r, conv & 3);
    if (conv & 16) r = kana(r, 1);              /* (the kana before the width: ア -> あ stays wide) */
    if (conv & 32) r = kana(r, 0);
    if (conv & 4) r = to_wide(r);
    if (conv & 8) r = to_narrow(r);
    return round_trip(r, cp);
}

/* ---------------------------------------------------------------- Option Compare Text (StrComp)
 * CompareString with the case, the kana and the width ignored: letters by their base letter (accents count
 * only when the letters are the same), hyphens and apostrophes only when all else is equal; punctuation
 * before digits before letters; ß = ss, æ = ae, ς = σ. */
typedef struct { vb_wchar prim[4]; int np; vb_wchar marks[4]; int nm; } celem;

static int decomp_of(vb_wchar c, vb_wchar *base, int *nb, vb_wchar *marks, int *nm)
{
    static int built; static int idx[NDECOMP + 1];
    if (!built) { int k = 0; for (int i = 0; i < NDECOMP; i++) { idx[i] = k; int h = decomp[k + 1]; k += 2 + (h & 15) + (h >> 4); } idx[NDECOMP] = k; built = 1; }
    int lo = 0, hi = NDECOMP - 1;
    while (lo <= hi) {
        int mid = (lo + hi) / 2, k = idx[mid];
        if (decomp[k] == c) {
            int h = decomp[k + 1];
            *nb = h & 15; *nm = h >> 4;
            for (int j = 0; j < *nb; j++) base[j] = decomp[k + 2 + j];
            for (int j = 0; j < *nm; j++) marks[j] = decomp[k + 2 + *nb + j];
            return 1;
        }
        if (decomp[k] < c) lo = mid + 1; else hi = mid - 1;
    }
    return 0;
}

static vb_wchar fold(vb_wchar c)
{
    c = vb_wlower(c);
    if (c == 0x3C2) return 0x3C3;               /* final sigma */
    if (c >= 0x30A1 && c <= 0x30F6) return (vb_wchar)(c - 0x60);   /* katakana = hiragana */
    return c;
}

static int is_mark(vb_wchar c) { return (c >= 0x300 && c <= 0x36F) || c == 0x3099 || c == 0x309A; }

/* the weight of a base character: punctuation, digits, letters */
static unsigned weight(vb_wchar c)
{
    int cls;
    if (c >= '0' && c <= '9') cls = 1;
    else if ((c >= 'a' && c <= 'z') || (c >= 0xC0 && c != 0xD7 && c != 0xF7 && !(c >= 0x2000 && c <= 0x2BFF) &&
             !(c >= 0x3000 && c <= 0x303F) && !(c >= 0xFE30 && c <= 0xFE6F) && !(c >= 0xE000 && c <= 0xF8FF))) cls = 2;
    else cls = 0;
    return ((unsigned)cls << 16) | c;
}

static int elems(vbs s, celem *e, int cap, vb_wchar *ign, int *nign)
{
    int n = 0; *nign = 0;
    for (vb_i32 i = 0; i < s.len && n < cap; i++) {
        vb_wchar c = s.p[i];
        if (c == '-' || c == '\'' || c == 0xAD) { if (*nign < cap) ign[(*nign)++] = c; continue; }
        if (is_mark(c) && n) { celem *p = &e[n - 1]; if (p->nm < 4) p->marks[p->nm++] = c; continue; }
        vb_wchar base[4], marks[4]; int nb = 1, nm = 0;
        base[0] = c;
        decomp_of(c, base, &nb, marks, &nm);
        for (int j = 0; j < nb && n < cap; j++) {
            vb_wchar b = fold(base[j]);
            celem *x = &e[n++];
            x->np = 1; x->nm = 0;
            if (b == 0xDF) { x->prim[0] = 's'; if (n < cap) { e[n].np = 1; e[n].nm = 0; e[n].prim[0] = 's'; n++; } }
            else if (b == 0xE6) { x->prim[0] = 'a'; if (n < cap) { e[n].np = 1; e[n].nm = 0; e[n].prim[0] = 'e'; n++; } }
            else if (b == 0x153) { x->prim[0] = 'o'; if (n < cap) { e[n].np = 1; e[n].nm = 0; e[n].prim[0] = 'e'; n++; } }
            else x->prim[0] = b;
            if (j == nb - 1) for (int m = 0; m < nm && x->nm < 4; m++) x->marks[x->nm++] = marks[m];
        }
    }
    return n;
}

/* a hash of s from what vb_collate compares (the letters, their accents, the hyphens and apostrophes), so that
 * the Strings it finds equal have the same hash: the keys of a Collection (object.c) */
unsigned vb_collate_hash(vbs s)
{
    enum { CAP = 512 };
    celem *e = (celem *)malloc(sizeof(celem) * CAP);
    vb_wchar ig[CAP];
    int nig;
    if (!e) return 0;
    int n = elems(s, e, CAP, ig, &nig);
    unsigned h = 2166136261u;
#define HX(v) (h = (h ^ (unsigned)(v)) * 16777619u)
    HX(n);
    for (int i = 0; i < n; i++) { HX(e[i].prim[0]); HX(e[i].nm); for (int k = 0; k < e[i].nm; k++) HX(e[i].marks[k]); }
    HX(nig);
    for (int i = 0; i < nig; i++) HX(ig[i]);
#undef HX
    free(e);
    return h;
}

vb_i32 vb_collate(vbs a, vbs b)
{
    enum { CAP = 512 };
    celem *ea = (celem *)malloc(sizeof(celem) * CAP * 2);
    vb_wchar ia[CAP], ib[CAP];
    if (!ea) return 0;
    celem *eb = ea + CAP;
    int nia, nib;
    int na = elems(a, ea, CAP, ia, &nia), nb = elems(b, eb, CAP, ib, &nib);
    int r = 0;
    for (int i = 0; i < na && i < nb && !r; i++) {             /* the letters */
        unsigned x = weight(ea[i].prim[0]), y = weight(eb[i].prim[0]);
        if (x != y) r = x < y ? -1 : 1;
    }
    if (!r && na != nb) r = na < nb ? -1 : 1;
    for (int i = 0; i < na && !r; i++) {                       /* the accents */
        int m = ea[i].nm < eb[i].nm ? ea[i].nm : eb[i].nm;
        for (int k = 0; k < m && !r; k++) if (ea[i].marks[k] != eb[i].marks[k]) r = ea[i].marks[k] < eb[i].marks[k] ? -1 : 1;
        if (!r && ea[i].nm != eb[i].nm) r = ea[i].nm < eb[i].nm ? -1 : 1;
    }
    if (!r && nia != nib) r = nia < nib ? -1 : 1;             /* hyphens, apostrophes */
    for (int i = 0; i < nia && !r; i++) if (ia[i] != ib[i]) r = ia[i] < ib[i] ? -1 : 1;
    free(ea);
    return r;
}
