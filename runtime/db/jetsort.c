/* jetsort.c - the keys Jet writes in its indexes (Jet 3 and Jet 4, General sort order): text, numbers, dates,
 * Yes/No, GUID, binary; ascending and descending; several columns; the long keys cut with a CRC. Measured on the
 * oracles of DEV-14 (jetkeys, jetnums, jetgen2): see tools/jet/jetsort.py and notes/phase-14.md.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "jetfile.h"
#include <stdlib.h>
#include <string.h>

#include "jetsort_tab.inc"

static void kput(jkey *k, const void *p, int n)
{
    if (k->n + n > k->cap) {
        k->cap = (k->n + n) * 2 + 32;
        k->p = realloc(k->p, (size_t)k->cap);
    }
    if (n) memcpy(k->p + k->n, p, (size_t)n);
    k->n += n;
}
static void kbyte(jkey *k, int b) { unsigned char c = (unsigned char)b; kput(k, &c, 1); }

/* ---------------------------------------------------------------- Jet 4 text */
/* the sort order of the keys (Jet 4: the LCID of the database / column; 0x409 General) */
static int cur_lcid = 0x409;
void jet_sort_lcid(int lcid) { cur_lcid = lcid ? lcid : 0x409; }

static const jk4run *find_run(const jk4run *t, int n, unsigned c)
{
    int lo = 0, hi = n - 1;
    while (lo <= hi) {
        int m = (lo + hi) / 2;
        if (c < t[m].cp) hi = m - 1;
        else if (c >= (unsigned)t[m].cp + t[m].n) lo = m + 1;
        else return &t[m];
    }
    return 0;
}
static int ord_of(int lcid)
{
    for (int i = 0; i < (int)(sizeof jk4_orders / sizeof jk4_orders[0]); i++) if (jk4_orders[i].lcid == lcid) return i;
    return -1;
}
/* a character in the order: its flags, primary weight (p, *pl), diacritic (d, *dl); 0 when it is not weighed */
static int weights4(int ord, unsigned c, unsigned char *p, int *pl, unsigned char *d, int *dl, int *flags)
{
    *pl = *dl = 0; *flags = 0;
    if (c == 0x20) { p[0] = 7; *pl = 1; return 1; }      /* a space inside the text */
    const jk4run *r = 0;
    if (ord >= 0 && jk4_orders[ord].nruns) r = find_run(jk4_orders[ord].runs, jk4_orders[ord].nruns, c);
    if (!r) r = find_run(jk4_runs, (int)(sizeof jk4_runs / sizeof jk4_runs[0]), c);
    if (!r) return 0;
    *flags = r->f;
    *pl = r->pl;
    memcpy(p, r->p, (size_t)*pl);
    if (*pl) p[*pl - 1] = (unsigned char)(p[*pl - 1] + (c - r->cp));
    *dl = r->dl;
    memcpy(d, r->d, (size_t)*dl);
    return *pl || *dl || (r->f & 4);
}
static int units4(int ord, unsigned c)
{
    if (ord >= 0)
        for (int i = 0; i < jk4_orders[ord].nu; i++) if (jk4_lunits[jk4_orders[ord].u0 + i].cp == c) return jk4_lunits[jk4_orders[ord].u0 + i].n;
    int lo = 0, hi = (int)(sizeof jk4_units / sizeof jk4_units[0]) - 1;
    while (lo <= hi) {
        int m = (lo + hi) / 2;
        if (c < jk4_units[m].cp) hi = m - 1; else if (c > jk4_units[m].cp) lo = m + 1; else return jk4_units[m].n;
    }
    return 1;
}
static unsigned vowel_of(unsigned c)
{
    int lo = 0, hi = (int)(sizeof jk4_vowel / sizeof jk4_vowel[0]) - 1;
    while (lo <= hi) {
        int m = (lo + hi) / 2;
        if (c < jk4_vowel[m].cp) hi = m - 1; else if (c > jk4_vowel[m].cp) lo = m + 1; else return jk4_vowel[m].vowel;
    }
    return 0;
}
/* the small letter of a capital (Basic Latin, Latin-1, Latin Extended-A, Ǆ ... Ǌ: the letters of the contractions) */
static int ext_a_upper(unsigned c)
{
    if (c >= 0x100 && c <= 0x137) return !(c & 1);
    if (c >= 0x139 && c <= 0x148) return c & 1;
    if (c >= 0x14A && c <= 0x177) return !(c & 1);
    if (c >= 0x179 && c <= 0x17E) return c & 1;
    return 0;
}
static unsigned lower16(unsigned c)
{
    if ((c >= 'A' && c <= 'Z') || (c >= 0xC0 && c <= 0xDE && c != 0xD7)) return c + 32;
    if (ext_a_upper(c)) return c + 1;
    if (c == 0x178) return 0xFF;
    if (c == 0x1C4 || c == 0x1C5 || c == 0x1C7 || c == 0x1C8 || c == 0x1CA || c == 0x1CB) return c + (c == 0x1C4 || c == 0x1C7 || c == 0x1CA ? 2 : 1);
    return c;
}
static int is_upper16(unsigned c) { return lower16(c) != c; }
static int is_lower16(unsigned c)
{
    if ((c >= 'a' && c <= 'z') || (c >= 0xDF && c <= 0xFF && c != 0xF7)) return 1;
    if (c > 0x100 && c <= 0x17E && ext_a_upper(c - 1) && !ext_a_upper(c)) return 1;
    if (c == 0x1C6 || c == 0x1C9 || c == 0x1CC) return 1;
    return 0;
}
/* a contraction of the order at t[i] (as written, in capitals or a capital and small letters; a doubled one by its
 * last two; letters without case always): its index, *len characters */
static int contraction(int ord, const unsigned *t, int n, int i, int *len)
{
    if (ord < 0) return -1;
    for (int want = 4; want >= 2; want--) {
        if (i + want > n) continue;
        for (int k = 0; k < jk4_orders[ord].nc; k++) {
            int ci = jk4_orders[ord].c0 + k;
            if (jk4_contr[ci].n != want) continue;
            int ok = 1;
            for (int j = 0; j < want && ok; j++) if (lower16(t[i + j]) != jk4_contr[ci].s[j]) ok = 0;
            if (!ok) continue;
            int s = want >= 3 && lower16(t[i]) == lower16(t[i + 1]) ? 1 : 0, cased = 0, up = 0, lo = 0;
            for (int j = s; j < want; j++) { if (is_upper16(t[i + j])) { up++; cased++; } else if (is_lower16(t[i + j])) { lo++; cased++; } }
            int m = want - s;
            if (cased && !(lo == m || up == m || (is_upper16(t[i + s]) && lo == m - 1))) continue;
            *len = want;
            return ci;
        }
    }
    return -1;
}

static void put_unp(jkey *k, const int *upos, const int *uw, int nu)
{
    for (int i = 0; i < nu; i++) {
        int v = 7 + 4 * upos[i];
        if (v < 256) { kbyte(k, 0x80); kbyte(k, v); }
        else { kbyte(k, 0x80 | (v >> 8)); kbyte(k, v & 255); }
        kbyte(k, 6);
        kbyte(k, uw[i]);
    }
}
/* a kana field: bytes of '10' + three codes of two bits (the last one filled with 0), the defaults at the end out */
static void kana_bits(jkey *k, const unsigned char *codes, int n, int def)
{
    while (n && codes[n - 1] == def) n--;
    for (int i = 0; i < n; i += 3) {
        unsigned v = 2;
        for (int j = 0; j < 3; j++) v = (v << 2) | (i + j < n ? codes[i + j] : 0);
        kbyte(k, (int)v);
    }
}

/* 7f, the primary weights, 01, the diacritics (02 none, trailing 02 / 00 / 01 cut), then: with kana 01 01, their
 * small-kana field ff the ー / ゝ field 02, 80 a seven kana ff 80 a seven kana [ff 01 unprintables]; without,
 * [01 01 01 and for each unprintable character 80 (7 + 4 * its place) 06 weight]; 00 (DEV-14/jetloc) */
static void text_key4(jkey *k, const unsigned char *tb, int nbytes)
{
    int n = nbytes / 2;
    while (n > 0 && tb[2 * (n - 1)] == 0x20 && tb[2 * (n - 1) + 1] == 0) n--;
    unsigned *t = malloc(sizeof(unsigned) * (size_t)(n + 1));
    for (int i = 0; i < n; i++) t[i] = (unsigned)tb[2 * i] | (unsigned)tb[2 * i + 1] << 8;
    int ord = cur_lcid == 0x409 ? -1 : ord_of(cur_lcid);
    jkey prim = { 0 }, dia = { 0 };
    int cap = n * 4 + 4;
    int *dstart = malloc(sizeof(int) * (size_t)cap), *dlen = malloc(sizeof(int) * (size_t)cap), nd = 0;
    int *upos = malloc(sizeof(int) * (size_t)(n + 1)), *uw = malloc(sizeof(int) * (size_t)(n + 1)), nu = 0, pos = 0;
    unsigned char *a1 = malloc((size_t)n + 1), *a2 = malloc((size_t)n + 1);
    int na = 0, anykana = 0, letter = 0;
    /* the character before: its weight, kana?, small?, the weight of its vowel */
    int have_prev = 0, prev_kana = 0, prev_small = 0, ppl = 0, pvl = 0;
    unsigned char pp[4], pv[4];
#define DPUT(b, l) do { dstart[nd] = dia.n; dlen[nd] = (l); nd++; kput(&dia, (b), (l)); } while (0)
    for (int i = 0; i < n; i++) {
        unsigned c = t[i];
        int clen;
        int ci = contraction(ord, t, n, i, &clen);
        if (ci >= 0) {
            kput(&prim, jk4_contr[ci].p, jk4_contr[ci].pl);
            if (jk4_contr[ci].dl) DPUT(jk4_contr[ci].d, jk4_contr[ci].dl); else DPUT("\2", 1);
            for (int j = 1; j < jk4_contr[ci].units; j++) DPUT("\2", 1);
            pos += jk4_contr[ci].units;
            i += clen - 1;
            have_prev = 1; prev_kana = 0; ppl = jk4_contr[ci].pl; memcpy(pp, jk4_contr[ci].p, (size_t)ppl);
            letter = 1;
            continue;
        }
        unsigned char p[4], d[2];
        int pl, dl, f;
        int weighed = weights4(ord, c, p, &pl, d, &dl, &f);
        /* the marks that repeat / lengthen the character before them (ー ゝ 々 shadda ...: weight ffff alone) */
        int rep = (c == 0x3005 || c == 0x3031 || c == 0x309D || c == 0x30FD || c == 0x0651 || c == 0xFE7C || c == 0xFE7D) ? 1 :
                  (c == 0x3032 || c == 0x309E || c == 0x30FE) ? 2 : (c == 0x30FC || c == 0xFF70 || c == 0x2015) ? 3 : 0;
        if (rep && pl == 2 && p[0] == 0xFF && p[1] == 0xFF) {
            if (!have_prev) { kput(&prim, p, 2); continue; }
            if (rep == 3 && prev_kana && pvl) kput(&prim, pv, pvl); else kput(&prim, pp, ppl);
            DPUT(rep == 2 ? "\3" : c == 0x3005 ? "\5" : "\2", 1);
            pos++;
            if (prev_kana) { a1[na] = prev_small ? 2 : 3; a2[na] = rep == 3 ? 3 : 2; na++; }
            continue;
        }
        if (!weighed) continue;                 /* not weighed (DEV-14/jetloc y_0409) */
        if (f & 4) { upos[nu] = pos; uw[nu++] = d[0]; continue; }
        if (f & 1) {                            /* kana */
            kput(&prim, p, pl);
            if (dl) DPUT(d, dl); else DPUT("\2", 1);
            pos++;
            anykana = 1;
            a1[na] = (f & 2) ? 2 : 3; a2[na] = 1; na++;
            have_prev = 1; prev_kana = 1; prev_small = (f & 2) != 0; ppl = pl; memcpy(pp, p, (size_t)pl);
            unsigned v = vowel_of(c);
            pvl = 0;
            if (v) { unsigned char vd[2]; int vdl, vf; weights4(ord, v, pv, &pvl, vd, &vdl, &vf); }
            else { pvl = pl; memcpy(pv, p, (size_t)pl); }
            continue;
        }
        if (f & 8) {                            /* a combining mark: what it adds to the letter's diacritic */
            unsigned char add = dl ? d[0] : 2;
            if (c == 0x0E4C && cur_lcid == 0x41E && nd && letter && dlen[nd - 1] == 1 && dia.p[dstart[nd - 1]] == 3)
                dia.p[dstart[nd - 1]] = 2;      /* Thai ์ on a consonant */
            else if (nd && dlen[nd - 1] == 1 && dl <= 1) dia.p[dstart[nd - 1]] = (unsigned char)(dia.p[dstart[nd - 1]] + add);
            else if (nd) { dia.n = dstart[nd - 1]; dlen[nd - 1] = dl; kput(&dia, d, dl); }
            else if (dl == 2) DPUT(d, 2);
            else if (add || (c == 0x0E4C && cur_lcid == 0x41E)) DPUT(&add, 1);
            letter = 0;
            continue;
        }
        kput(&prim, p, pl);
        int units = units4(ord, c);
        if (dl) DPUT(d, dl); else DPUT("\2", 1);
        for (int j = 1; j < units; j++) DPUT("\2", 1);
        pos += units;
        have_prev = 1; prev_kana = 0; ppl = pl; memcpy(pp, p, (size_t)pl);
        letter = 1;
    }
#undef DPUT
    while (nd && dlen[nd - 1] == 1 && (dia.p[dstart[nd - 1]] == 2 || dia.p[dstart[nd - 1]] == 0 || dia.p[dstart[nd - 1]] == 1)) {
        dia.n = dstart[nd - 1]; nd--;
    }
    kbyte(k, 0x7f);
    kput(k, prim.p, prim.n);
    kbyte(k, 1);
    kput(k, dia.p, dia.n);
    if (anykana) {
        kput(k, "\1\1", 2);
        kana_bits(k, a1, na, 3);
        kbyte(k, 0xFF);
        kana_bits(k, a2, na, 1);
        kbyte(k, 2);
        for (int i = 0; i < (na + 6) / 7; i++) kbyte(k, 0x80);
        kbyte(k, 0xFF);
        for (int i = 0; i < (na + 6) / 7; i++) kbyte(k, 0x80);
        if (nu) { kput(k, "\xff\1", 2); put_unp(k, upos, uw, nu); }
    } else if (nu) {
        kput(k, "\1\1\1", 3);
        put_unp(k, upos, uw, nu);
    }
    kbyte(k, 0);
    free(prim.p); free(dia.p); free(dstart); free(dlen); free(upos); free(uw); free(a1); free(a2); free(t);
}

/* ---------------------------------------------------------------- Jet 3 text */
static unsigned char p3[256][4], p3n[256], c3[256], w3[256];
static int init3;
static void init_jet3(void)
{
    for (int b = 0; b < 256; b++) {
        const unsigned char *key = jk3_keys[b];
        int len = key[0];
        const unsigned char *body = key + 2;    /* after the 7f */
        int bl = len - 1;
        if (bl >= 2 && body[bl - 1] == 0 && body[bl - 2] > 0 && body[bl - 2] < 0x10) {
            p3n[b] = (unsigned char)(bl - 2); c3[b] = body[bl - 2];
        } else {
            p3n[b] = (unsigned char)(bl > 0 ? bl - 1 : 0); c3[b] = 0;
        }
        memcpy(p3[b], body, p3n[b]);
    }
    p3[0x20][0] = 0x11; p3n[0x20] = 1; c3[0x20] = 0;
    for (int b = 0; b < 256; b++) if (c3[b] && p3n[b] == 1) w3[p3[b][0]] = 1;
    w3[0x76] = 1;                               /* s (an accented form in the cp1252 table: s with caron) */
    init3 = 1;
}
/* 7f, the weights of each byte, a run of 4 bit codes starting with 0 (one per weight of a letter with accented
 * forms: the accent or 2), trailing 2s cut, two to a byte; an even run ends with 00 */
static void text_key3(jkey *k, const unsigned char *t, int n)
{
    if (!init3) init_jet3();
    while (n > 0 && t[n - 1] == ' ') n--;
    unsigned char *codes = malloc((size_t)n * 4 + 2);
    int nc = 0;
    kbyte(k, 0x7f);
    for (int i = 0; i < n; i++) {
        int b = t[i], first = 1;
        kput(k, p3[b], p3n[b]);
        for (int j = 0; j < p3n[b]; j++)
            if (w3[p3[b][j]]) { codes[nc++] = (first && c3[b]) ? c3[b] : 2; first = 0; }
    }
    while (nc && codes[nc - 1] == 2) nc--;
    int total = nc + 1;                         /* the leading 0 */
    for (int i = 0; i < total; i += 2) {
        int hi = i == 0 ? 0 : codes[i - 1];
        int lo = i + 1 < total ? codes[i] : 0;
        kbyte(k, hi << 4 | lo);
    }
    if (total % 2 == 0) kbyte(k, 0);
    free(codes);
}

void jet_text_key(jkey *k, const unsigned char *text, int n, int jet4)
{
    if (jet4) text_key4(k, text, n); else text_key3(k, text, n);
}

/* ---------------------------------------------------------------- columns */
static void groups8(jkey *k, const unsigned char *d, int n)
{
    if (n == 0) { static const unsigned char z[9]; kput(k, z, 9); return; }
    for (int i = 0; i < n; i += 8) {
        unsigned char g[8] = { 0 };
        int m = n - i < 8 ? n - i : 8;
        memcpy(g, d + i, (size_t)m);
        kput(k, g, 8);
        kbyte(k, i + 8 < n ? 9 : m);
    }
}

/* the ascending body of a fixed value: big-endian, the sign bit flipped (negative floating point: every bit) */
static void fixed_body(jkey *k, int type, const unsigned char *r, int n)
{
    unsigned char b[16];
    int len;
    switch (type) {
    case JT_BYTE: kput(k, r, 1); return;
    case JT_INT: case JT_LONG: case JT_MONEY:
        len = type == JT_INT ? 2 : type == JT_LONG ? 4 : 8;
        for (int i = 0; i < len; i++) b[i] = r[len - 1 - i];
        b[0] ^= 0x80;
        kput(k, b, len);
        return;
    case JT_FLOAT: case JT_DOUBLE: case JT_DATETIME:
        len = type == JT_FLOAT ? 4 : 8;
        for (int i = 0; i < len; i++) b[i] = r[len - 1 - i];
        if (b[0] & 0x80) for (int i = 0; i < len; i++) b[i] ^= 0xFF;
        else b[0] ^= 0x80;
        kput(k, b, len);
        return;
    case JT_GUID:                               /* in the order of its text {01234567-89AB-...} */
        b[0] = r[3]; b[1] = r[2]; b[2] = r[1]; b[3] = r[0]; b[4] = r[5]; b[5] = r[4]; b[6] = r[7]; b[7] = r[6];
        memcpy(b + 8, r + 8, 8);
        groups8(k, b, 16);
        return;
    case JT_BINARY:
        groups8(k, r, n);
        return;
    case JT_NUMERIC: {                          /* Decimal: ff + the unscaled value big-endian; negative: 00 + it
                                                 * inverted (VB6 oracle DEV-14/jetdec) */
        unsigned char d[17];
        if (n < 17) break;
        int neg = (r[0] & 0x80) != 0;
        d[0] = neg ? 0x00 : 0xFF;
        for (int w = 0; w < 4; w++)
            for (int i = 0; i < 4; i++) d[1 + 4 * w + i] = r[1 + 4 * w + 3 - i];
        if (neg) for (int i = 1; i < 17; i++) d[i] ^= 0xFF;
        kput(k, d, 17);
        return; }
    }
    kput(k, r, n);
}

void jet_col_key(jkey *k, int type, const jv *v, int desc, int jet4)
{
    if (type == JT_BOOL) {                      /* never Null: True 00, False ff */
        kbyte(k, desc ? 0x80 : 0x7f);
        int b = v->n ? 0x00 : 0xFF;
        kbyte(k, desc ? b ^ 0xFF : b);
        return;
    }
    if (v->null) { kbyte(k, desc ? 0xFF : 0x00); return; }
    int start = k->n;
    if (type == JT_TEXT || type == JT_MEMO) {
        int n = v->n;
        if (type == JT_MEMO && n > (jet4 ? 510 : 255)) n = jet4 ? 510 : 255;   /* its first 255 characters */
        jet_text_key(k, v->p, n, jet4);
        if (desc) {
            k->p[start] = 0x80;
            for (int i = start + 1; i < k->n; i++) k->p[i] ^= 0xFF;
            if (jet4) kbyte(k, 0);
        }
        return;
    }
    kbyte(k, 0x7f);
    fixed_body(k, type, v->p, v->n);
    if (desc) {
        k->p[start] = 0x80;
        for (int i = start + 1; i < k->n; i++) k->p[i] ^= 0xFF;
        if (type == JT_GUID || type == JT_BINARY)   /* the 09 between the groups stays */
            for (int i = start + 1 + 8; i < k->n - 1; i += 9) k->p[i] = 9;
    }
}

/* a key longer than 255 (Jet 3) / 510 (Jet 4) bytes: all but the last 2 of those, then the CRC-16 (8005, the bits
 * shifted in at the bottom) of the bytes it drops, little-endian */
void jet_key_truncate(jkey *k, int jet4)
{
    int lim = jet4 ? 510 : 255;
    if (k->n <= lim) return;
    unsigned r = 0;
    for (int i = lim - 2; i < k->n; i++)
        for (int bit = 7; bit >= 0; bit--) {
            unsigned top = r & 0x8000;
            r = ((r << 1) | ((k->p[i] >> bit) & 1)) & 0xFFFF;
            if (top) r ^= 0x8005;
        }
    k->n = lim - 2;
    kbyte(k, r & 0xFF);
    kbyte(k, r >> 8);
}
