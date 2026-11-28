/* jetread.c - a Jet 3 / Jet 4 file read whole into the model of jetfile.h: every table found in MSysObjects (the
 * system tables too), its TDEF (columns, indexes, the raw bits it does not understand) and its rows, long values
 * resolved. The Python prototype: tools/jet/jetdb.py (Reader).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "jetfile.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#define U16(b, o) ((unsigned)(b)[o] | (unsigned)(b)[(o) + 1] << 8)
#define U32(b, o) ((uint32_t)(b)[o] | (uint32_t)(b)[(o) + 1] << 8 | (uint32_t)(b)[(o) + 2] << 16 | (uint32_t)(b)[(o) + 3] << 24)

static const int FIXLEN[17] = { 0, 0, 1, 2, 4, 8, 4, 8, 8, 0, 0, 0, 0, 0, 0, 16, 17 };

void jet_rc4(uint32_t key, unsigned char *data, int n)
{
    unsigned char s[256], kb[4] = { key & 255, key >> 8 & 255, key >> 16 & 255, key >> 24 & 255 };
    for (int i = 0; i < 256; i++) s[i] = (unsigned char)i;
    for (int i = 0, j = 0; i < 256; i++) {
        j = (j + s[i] + kb[i % 4]) & 255;
        unsigned char t = s[i]; s[i] = s[j]; s[j] = t;
    }
    for (int k = 0, i = 0, j = 0; k < n; k++) {
        i = (i + 1) & 255;
        j = (j + s[i]) & 255;
        unsigned char t = s[i]; s[i] = s[j]; s[j] = t;
        data[k] ^= s[(s[i] + s[j]) & 255];
    }
}
#define HDR_KEY 0x6B39DAC7u

/* ---------------------------------------------------------------- the model */
jtable *jet_table_new(const char *name)
{
    jtable *t = calloc(1, sizeof *t);
    t->name = strdup(name ? name : "");
    t->type = 0x4E;
    t->autonum_flag = 1;
    t->maxcols = t->nvar = -1;
    return t;
}
void jet_table_free(jtable *t)
{
    if (!t) return;
    for (int i = 0; i < t->ncols; i++) free(t->cols[i].name);
    for (int i = 0; i < t->nlidx; i++) free(t->lidx[i].name);
    if (t->vals) for (int i = 0; i < t->nrows * t->width; i++) free(t->vals[i].p);   /* (0: freed when written) */
    free(t->cols); free(t->ridx); free(t->lidx); free(t->vals); free(t->lvorder); free(t->name);
    free(t);
}
void jet_free(jdb *db)
{
    if (!db) return;
    for (int i = 0; i < db->nt; i++) jet_table_free(db->t[i]);
    free(db->t); free(db->page0); free(db);
}
jtable *jet_table(jdb *db, const char *name)
{
    for (int i = 0; i < db->nt; i++) if (!jet_strcasecmp8(db->t[i]->name, name)) return db->t[i];
    return 0;
}
jetcol *jet_col(jtable *t, const char *name)
{
    for (int i = 0; i < t->ncols; i++) if (!jet_strcasecmp8(t->cols[i].name, name)) return &t->cols[i];
    return 0;
}
jv *jet_val(jtable *t, int row, int colnum) { return &t->vals[(size_t)row * (size_t)t->width + (size_t)colnum]; }
void jet_val_set(jv *v, const void *p, int n)
{
    free(v->p);
    v->p = 0; v->n = 0; v->null = !p;
    if (p) { v->p = malloc((size_t)n + 1); memcpy(v->p, p, (size_t)n); v->n = n; }
}
int jet_row_add(jtable *t)
{
    if (t->width <= 0) {
        int w = 0;
        for (int i = 0; i < t->ncols; i++) if (t->cols[i].num + 1 > w) w = t->cols[i].num + 1;
        if (t->maxcols > w) w = t->maxcols;
        t->width = w ? w : 1;
    }
    if (t->nrows == t->caprows) {
        t->caprows = t->caprows ? t->caprows * 2 : 16;
        t->vals = realloc(t->vals, sizeof(jv) * (size_t)t->caprows * (size_t)t->width);
    }
    jv *r = &t->vals[(size_t)t->nrows * (size_t)t->width];
    for (int c = 0; c < t->width; c++) { r[c].p = 0; r[c].n = 0; r[c].null = 1; }
    for (int i = 0; i < t->ncols; i++) if (t->cols[i].type == JT_BOOL) { r[t->cols[i].num].null = 0; r[t->cols[i].num].n = 0; }
    return t->nrows++;
}

/* ---------------------------------------------------------------- text */
int jet_strcasecmp8(const char *a, const char *b)
{
    uint16_t *x, *y;
    int nx = jet_utf8_to_utf16(a, &x), ny = jet_utf8_to_utf16(b, &y), r = 0;
    for (int i = 0; ; i++) {
        if (i == nx || i == ny) { r = (i == nx) - (i == ny); r = -r; break; }
        unsigned p = x[i], q = y[i];
        if ((p >= 'a' && p <= 'z') || (p >= 0xE0 && p <= 0xFE && p != 0xF7)) p -= 32;
        if ((q >= 'a' && q <= 'z') || (q >= 0xE0 && q <= 0xFE && q != 0xF7)) q -= 32;
        if (p >= 0x3B1 && p <= 0x3C9 && p != 0x3C2) p -= 32;
        if (q >= 0x3B1 && q <= 0x3C9 && q != 0x3C2) q -= 32;
        if (p != q) { r = p < q ? -1 : 1; break; }
    }
    free(x); free(y);
    return r;
}
char *jet_utf16_to_utf8(const uint16_t *w, int n)
{
    char *s = malloc((size_t)n * 3 + 1), *o = s;
    for (int i = 0; i < n; i++) {
        unsigned c = w[i];
        if (c >= 0xD800 && c < 0xDC00 && i + 1 < n && w[i + 1] >= 0xDC00 && w[i + 1] < 0xE000) {
            c = 0x10000 + ((c - 0xD800) << 10) + (w[i + 1] - 0xDC00); i++;
            *o++ = (char)(0xF0 | c >> 18); *o++ = (char)(0x80 | (c >> 12 & 63));
            *o++ = (char)(0x80 | (c >> 6 & 63)); *o++ = (char)(0x80 | (c & 63));
        } else if (c < 0x80) *o++ = (char)c;
        else if (c < 0x800) { *o++ = (char)(0xC0 | c >> 6); *o++ = (char)(0x80 | (c & 63)); }
        else { *o++ = (char)(0xE0 | c >> 12); *o++ = (char)(0x80 | (c >> 6 & 63)); *o++ = (char)(0x80 | (c & 63)); }
    }
    *o = 0;
    return s;
}
int jet_utf8_to_utf16(const char *s, uint16_t **out)
{
    int n = (int)strlen(s), k = 0;
    uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(n + 1));
    for (int i = 0; i < n; ) {
        unsigned c = (unsigned char)s[i];
        if (c < 0x80) i++;
        else if ((c & 0xE0) == 0xC0 && i + 1 < n) { c = (c & 31) << 6 | (s[i + 1] & 63); i += 2; }
        else if ((c & 0xF0) == 0xE0 && i + 2 < n) { c = (c & 15) << 12 | (s[i + 1] & 63) << 6 | (s[i + 2] & 63); i += 3; }
        else if ((c & 0xF8) == 0xF0 && i + 3 < n) {
            c = (c & 7) << 18 | (s[i + 1] & 63) << 12 | (s[i + 2] & 63) << 6 | (s[i + 3] & 63); i += 4;
            c -= 0x10000;
            w[k++] = (uint16_t)(0xD800 + (c >> 10));
            c = 0xDC00 + (c & 0x3FF);
        } else i++;
        w[k++] = (uint16_t)c;
    }
    w[k] = 0;
    *out = w;
    return k;
}

/* ---------------------------------------------------------------- reading */
/* the file is never all in memory: a page is read when it is needed (and dropped after the data page whose rows it
 * served); the type and owner of every page from one pass over the file (a table's rows: its pages only) */
typedef struct rd {
    FILE *f; size_t size;
    int jet4, ps, dhdr, npages, cp;
    uint32_t dbkey;
    unsigned char **pg; uint32_t *loaded; int nloaded;   /* the pages in memory */
    unsigned char *ptype; uint32_t *powner;              /* byte 0 and bytes 4..7 of each page */
    char *err; int cap;
} rd;

static const unsigned char *page(rd *r, uint32_t n)
{
    if ((size_t)(n + 1) * (size_t)r->ps > r->size) return 0;
    if (!r->pg[n]) {
        unsigned char *b = malloc((size_t)r->ps);
        if (fseek(r->f, (long)n * r->ps, SEEK_SET) || fread(b, 1, (size_t)r->ps, r->f) != (size_t)r->ps) memset(b, 0, (size_t)r->ps);
        /* an encrypted database (Encrypt, dbEncrypt): each page but the first RC4 with the database key xor its number */
        if (r->dbkey && n) jet_rc4(r->dbkey ^ n, b, r->ps);
        r->pg[n] = b;
        r->loaded[r->nloaded++] = n;
    }
    return r->pg[n];
}
static void cache_drop(rd *r)
{
    for (int i = 0; i < r->nloaded; i++) { free(r->pg[r->loaded[i]]); r->pg[r->loaded[i]] = 0; }
    r->nloaded = 0;
}
static void rd_close(rd *r)
{
    if (r->pg) cache_drop(r);
    free(r->pg); free(r->loaded); free(r->ptype); free(r->powner);
    if (r->f) fclose(r->f);
}

/* a name in a TDEF */
static char *tname(rd *r, const unsigned char *b, int len, int *o)
{
    if (r->jet4) {
        int n = (int)U16(b, *o);
        if (*o + 2 + n > len) n = 0;
        uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(n / 2 + 1));
        for (int i = 0; i < n / 2; i++) w[i] = (uint16_t)U16(b, *o + 2 + 2 * i);
        char *s = jet_utf16_to_utf8(w, n / 2);
        free(w);
        *o += 2 + n;
        return s;
    }
    int n = b[*o];
    uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(n + 1));
    int k = jet_cp_to_utf16(r->cp, b + *o + 1, n, w);
    char *s = jet_utf16_to_utf8(w, k);
    free(w);
    *o += 1 + n;
    return s;
}

static jtable *read_tdef(rd *r, uint32_t pg, const char *name)
{
    unsigned char *b = 0;
    size_t len = 0;
    uint32_t p = pg;
    int first = 1, guard = 0;
    while (p && guard++ < 100000) {
        const unsigned char *blk = page(r, p);
        if (!blk || blk[0] != 2) { free(b); snprintf(r->err, (size_t)r->cap, "page %u is not a TDEF", p); return 0; }
        int from = first ? 0 : 8;
        b = realloc(b, len + (size_t)(r->ps - from));
        memcpy(b + len, blk + from, (size_t)(r->ps - from));
        len += (size_t)(r->ps - from);
        first = 0;
        p = U32(blk, 4);
    }
    jtable *t = jet_table_new(name);
    t->page = pg;
    t->keep_layout = 1;
    int ncol, nidx, nridx, o;
    if (r->jet4) {
        t->autonum = U32(b, 20); t->autonum_flag = b[24]; t->ct_autonum = U32(b, 28);
        t->type = b[0x28]; t->maxcols = (int)U16(b, 0x29); t->nvar = (int)U16(b, 0x2B); ncol = (int)U16(b, 0x2D);
        nidx = (int)U32(b, 0x2F); nridx = (int)U32(b, 0x33);
        t->usedmap = U32(b, 0x37);
        o = 0x3F + nridx * 12;
    } else {
        t->autonum = U32(b, 16);
        t->type = b[0x14]; t->maxcols = (int)U16(b, 0x15); t->nvar = (int)U16(b, 0x17); ncol = (int)U16(b, 0x19);
        nidx = (int)U32(b, 0x1B); nridx = (int)U32(b, 0x1F);
        t->usedmap = U32(b, 0x23);
        o = 0x2B + nridx * 8;
    }
    int csz = r->jet4 ? 25 : 18;
    if (ncol < 0 || ncol > 4096 || nidx < 0 || nidx > 4096 || nridx < 0 || nridx > 4096 ||
        (size_t)(o + ncol * csz) > len) {
        free(b); jet_table_free(t);
        snprintf(r->err, (size_t)r->cap, "a damaged TDEF (page %u)", pg);
        return 0;
    }
    t->cols = calloc((size_t)(ncol ? ncol : 1), sizeof(jetcol));
    t->ncols = ncol;
    for (int i = 0; i < ncol; i++, o += csz) {
        const unsigned char *c = b + o;
        jetcol *k = &t->cols[i];
        memcpy(k->raw, c, (size_t)csz); k->rawlen = csz;
        k->type = c[0];
        if (r->jet4) {
            k->num = (int)U16(c, 5); k->offv = (int)U16(c, 7); k->id = (int)U16(c, 9); k->misc = (int)U16(c, 11);
            k->misc2 = (int)U16(c, 13); k->flags = c[15]; k->xflags = c[16]; k->offf = (int)U16(c, 21); k->len = (int)U16(c, 23);
        } else {
            k->num = (int)U16(c, 1); k->offv = (int)U16(c, 3); k->id = (int)U16(c, 5); k->fileflag = (int)U16(c, 7);
            k->misc = (int)U16(c, 9); k->misc2 = (int)U16(c, 11); k->flags = c[13]; k->offf = (int)U16(c, 14); k->len = (int)U16(c, 16);
        }
    }
    for (int i = 0; i < ncol; i++) t->cols[i].name = tname(r, b, (int)len, &o);
    t->ridx = calloc((size_t)(nridx ? nridx : 1), sizeof(jridx));
    t->nridx = nridx;
    for (int i = 0; i < nridx; i++) {
        const unsigned char *blk = b + o;
        int q = r->jet4 ? 4 : 0;
        jridx *x = &t->ridx[i];
        for (int k = 0; k < 10; k++, q += 3) {
            unsigned cn = U16(blk, q);
            if (cn != 0xFFFF && x->n < 10) { x->col[x->n] = (int)cn; x->order[x->n] = blk[q + 2]; x->n++; }
        }
        q += 8;
        if (r->jet4) q += 4;
        x->flags = blk[q];
        x->complex = r->jet4 ? blk[q + 1] : 0;
        o += r->jet4 ? 52 : 39;
    }
    t->lidx = calloc((size_t)(nidx ? nidx : 1), sizeof(jlidx));
    t->nlidx = nidx;
    for (int i = 0; i < nidx; i++) {
        int q = o + (r->jet4 ? 4 : 0);
        jlidx *l = &t->lidx[i];
        l->num = (int)U32(b, q); l->real = (int)U32(b, q + 4); l->reltype = b[q + 8]; l->relidx = U32(b, q + 9);
        l->reltable = U32(b, q + 13); l->cup = b[q + 17]; l->cdel = b[q + 18]; l->type = b[q + 19];
        o += r->jet4 ? 28 : 20;
    }
    for (int i = 0; i < nidx; i++) t->lidx[i].name = tname(r, b, (int)len, &o);
    t->lvorder = malloc(sizeof(int) * 64);
    int caplv = 64;
    while ((size_t)o + 2 <= len) {
        unsigned cn = U16(b, o);
        if (cn == 0xFFFF) break;
        if (t->nlv == caplv) { caplv *= 2; t->lvorder = realloc(t->lvorder, sizeof(int) * (size_t)caplv); }
        t->lvorder[t->nlv++] = (int)cn;
        o += 10;
    }
    int w = t->maxcols;
    for (int i = 0; i < ncol; i++) if (t->cols[i].num + 1 > w) w = t->cols[i].num + 1;
    t->width = w ? w : 1;
    free(b);
    return t;
}

/* the rows of a data page: (offset field, start, end) */
static int row_slice(rd *r, const unsigned char *b, int i, unsigned *off, int *start, int *end)
{
    int h = r->dhdr, nr = (int)U16(b, h - 2);
    if (i >= nr) return 0;
    int prev = r->ps;
    for (int k = 0; k <= i; k++) {
        unsigned of = U16(b, h + 2 * k);
        int st = (int)(of & 0x1FFF);
        if (k == i) { *off = of; *start = st; *end = prev; return st <= prev; }
        prev = st;
    }
    return 0;
}

static unsigned char *lvrow(rd *r, uint32_t dp, int *n)
{
    const unsigned char *b = page(r, dp >> 8);
    unsigned off; int st, en;
    if (!b || !row_slice(r, b, (int)(dp & 0xFF), &off, &st, &en)) return 0;
    *n = en - st;
    return (unsigned char *)b + st;
}

static unsigned char *longvalue(rd *r, const unsigned char *h, int hn, int *outn)
{
    if (hn < 12) { *outn = 0; return calloc(1, 1); }
    uint32_t ln = U32(h, 0), kind = ln >> 30;
    ln &= 0x3FFFFFFF;
    unsigned char *out = malloc((size_t)ln + 1);
    int got = 0;
    if (kind == 2) {
        got = hn - 12 < (int)ln ? hn - 12 : (int)ln;
        memcpy(out, h + 12, (size_t)got);
    } else if (kind == 1) {
        int n;
        unsigned char *p = lvrow(r, U32(h, 4), &n);
        if (p) { got = n < (int)ln ? n : (int)ln; memcpy(out, p, (size_t)got); }
    } else {
        uint32_t dp = U32(h, 4);
        int guard = 0;
        while (dp && got < (int)ln && guard++ < 1000000) {
            int n;
            unsigned char *p = lvrow(r, dp, &n);
            if (!p || n < 4) break;
            dp = U32(p, 0);
            int m = n - 4;
            if (got + m > (int)ln) m = (int)ln - got;
            memcpy(out + got, p + 4, (size_t)m);
            got += m;
        }
    }
    *outn = got;
    return out;
}

/* Jet 4 text: UTF-16LE, or the compressed form (ff fe, then single bytes and pairs, 00 switching) */
static unsigned char *text4(const unsigned char *v, int n, int *outn)
{
    if (n >= 2 && v[0] == 0xFF && v[1] == 0xFE) {
        unsigned char *o = malloc((size_t)n * 2 + 2);
        int k = 0, comp = 1;
        for (int i = 2; i < n; ) {
            if (v[i] == 0) { comp = !comp; i++; continue; }
            if (comp) { o[k++] = v[i]; o[k++] = 0; i++; }
            else { if (i + 1 >= n) break; o[k++] = v[i]; o[k++] = v[i + 1]; i += 2; }
        }
        *outn = k;
        return o;
    }
    unsigned char *o = malloc((size_t)n + 1);
    memcpy(o, v, (size_t)n);
    *outn = n & ~1;
    return o;
}

static void crack(rd *r, jtable *t, const unsigned char *row, int len, jv *out)
{
    int nc, nb, nvar = 0, fixed0;
    int *tab = 0;
    const unsigned char *mask;
    if (r->jet4) {
        nc = (int)U16(row, 0); nb = (nc + 7) / 8;
        mask = row + len - nb;
        if (t->nvar) {
            nvar = (int)U16(row, len - nb - 2);
            if (len - nb - 4 - 2 * nvar < 0) nvar = 0;
            tab = malloc(sizeof(int) * (size_t)(nvar + 1));
            for (int k = 0; k <= nvar; k++) tab[k] = (int)U16(row, len - nb - 4 - 2 * k);
        }
        fixed0 = 2;
    } else {
        nc = row[0]; nb = (nc + 7) / 8;
        mask = row + len - nb;
        if (t->nvar) {
            nvar = row[len - nb - 1];
            int njump = (len - 1) / 256, pos = len - nb - 1 - njump;
            tab = malloc(sizeof(int) * (size_t)(nvar + 1));
            for (int k = 0; k <= nvar; k++) {
                int hi = 0;
                for (int j = 0; j < njump; j++) {
                    int jv_ = row[len - nb - 2 - j];
                    if (jv_ != 0xFF && k >= jv_) hi++;
                }
                tab[k] = (pos - 1 - k >= 0 ? row[pos - 1 - k] : 0) + 256 * hi;
            }
        }
        fixed0 = 1;
    }
    for (int i = 0; i < t->ncols; i++) {
        jetcol *c = &t->cols[i];
        jv *v = &out[c->num];
        int n = c->num;
        int bit = n < nc && n / 8 < nb ? (mask[n / 8] >> (n % 8)) & 1 : 0;
        free(v->p); v->p = 0; v->n = 0;
        if (c->type == JT_BOOL) { v->null = 0; v->n = bit; continue; }
        v->null = 1;
        if (!bit) continue;
        const unsigned char *vp; int vn;
        if (c->flags & 1) {
            int fl = c->type < 17 && FIXLEN[c->type] ? FIXLEN[c->type] : c->len;
            int s = fixed0 + c->offf;
            if (s + fl > len) continue;
            vp = row + s; vn = fl;
        } else {
            int k = c->offv;
            if (k >= nvar) continue;
            int s = tab[k], e = tab[k + 1];
            if (s < 0 || e > len || e < s) continue;
            vp = row + s; vn = e - s;
        }
        v->null = 0;
        if (c->type == JT_OLE || c->type == JT_MEMO) {
            int ln;
            unsigned char *lv = longvalue(r, vp, vn, &ln);
            if (c->type == JT_MEMO && r->jet4) { v->p = text4(lv, ln, &v->n); free(lv); }
            else { v->p = lv; v->n = ln; }
        } else if (c->type == JT_TEXT && r->jet4) {
            v->p = text4(vp, vn, &v->n);
        } else {
            v->p = malloc((size_t)vn + 1);
            memcpy(v->p, vp, (size_t)vn);
            v->n = vn;
        }
    }
    free(tab);
}

/* the pages a usage map marks (malloc, one byte per page of the file), 0 if it cannot be read */
static unsigned char *map_pages(rd *r, uint32_t ptr)
{
    const unsigned char *b = page(r, ptr >> 8);
    unsigned off; int st, en;
    if (!b || b[0] != 1 || !row_slice(r, b, (int)(ptr & 0xFF), &off, &st, &en) || en - st < 5) return 0;
    const unsigned char *m = b + st;
    int n = en - st;
    unsigned char *set = calloc((size_t)r->npages + 1, 1);
    if (m[0] == 0) {
        uint32_t start = U32(m, 1);
        for (int i = 0; i < (n - 5) * 8; i++)
            if (m[5 + i / 8] >> (i % 8) & 1) { uint32_t p = start + (uint32_t)i; if (p < (uint32_t)r->npages) set[p] = 1; }
    } else if (m[0] == 1) {
        uint32_t per = (uint32_t)(r->ps - 4) * 8;
        for (int k = 0; 1 + 4 * k + 4 <= n; k++) {
            uint32_t mp = U32(m, 1 + 4 * k);
            const unsigned char *pg = mp ? page(r, mp) : 0;
            if (!pg) continue;
            for (uint32_t i = 0; i < per; i++)
                if (pg[4 + i / 8] >> (i % 8) & 1) { uint32_t p = (uint32_t)k * per + i; if (p < (uint32_t)r->npages) set[p] = 1; }
        }
    } else { free(set); return 0; }
    return set;
}

static int read_rows(rd *r, jtable *t)
{
    /* the pages its usage map marks (as Jet reads a table: a page of it the map does not have, of a damaged file, is
     * not part of it; PSC file of DEV-14/q1) */
    unsigned char *used = t->usedmap ? map_pages(r, t->usedmap) : 0;
    const uint32_t lval = U32((const unsigned char *)"LVAL", 0);
    for (int p = 0; p < r->npages; p++) {
        if (r->ptype[p] != 1 || r->powner[p] != t->page || r->powner[p] == lval) continue;
        if (used && !used[p]) continue;
        cache_drop(r);                          /* (the pages of the rows before: copied) */
        const unsigned char *b = page(r, (uint32_t)p);
        if (!b) continue;
        int nr = (int)U16(b, r->dhdr - 2);
        for (int i = 0; i < nr; i++) {
            unsigned off; int st, en;
            if (!row_slice(r, b, i, &off, &st, &en)) continue;
            const unsigned char *row = b + st;
            int len = en - st;
            if (off & 0xC000) {
                if ((off & 0x4000) && !(off & 0x8000) && len >= 4) {     /* the row lives on another page */
                    int n;
                    unsigned char *q = lvrow(r, U32(row, 0), &n);
                    if (!q) continue;
                    row = q; len = n;
                } else continue;
            }
            if (len < 1) continue;
            int k = jet_row_add(t);
            crack(r, t, row, len, jet_val(t, k, 0));
        }
    }
    cache_drop(r);
    free(used);
    return 0;
}

jdb *jet_read(const char *path, int codepage, char *err, int cap) { return jet_read_opt(path, codepage, err, cap, 0); }
jdb *jet_read_opt(const char *path, int codepage, char *err, int cap, int opts)
{
    char dummy[8];
    if (!err) { err = dummy; cap = sizeof dummy; }
    *err = 0;
    FILE *f = fopen(path, "rb");
    if (!f) { snprintf(err, (size_t)cap, "cannot open %s", path); return 0; }
    fseek(f, 0, SEEK_END);
    long size = ftell(f);
    fseek(f, 0, SEEK_SET);
    rd r = { 0 };
    r.f = f;
    r.size = size > 0 ? (size_t)size : 0;
    r.err = err; r.cap = cap; r.cp = codepage ? codepage : 1252;
    unsigned char head[4096];
    size_t hn = fread(head, 1, sizeof head, f);
    if (r.size < 0x800 || hn < 0x800 || memcmp(head + 4, "Standard Jet DB", 15)) {
        rd_close(&r); snprintf(err, (size_t)cap, "not a Jet database"); return 0;
    }
    r.jet4 = head[0x14] >= 1;
    r.ps = r.jet4 ? 4096 : 2048;
    r.dhdr = r.jet4 ? 14 : 10;
    r.npages = (int)(r.size / (size_t)r.ps);
    if (r.npages < 3 || hn < (size_t)r.ps) { rd_close(&r); snprintf(err, (size_t)cap, "not a Jet database"); return 0; }
    jdb *db = calloc(1, sizeof *db);
    db->jet4 = r.jet4; db->ps = r.ps; db->codepage = r.cp;
    db->page0 = malloc((size_t)r.ps);
    memcpy(db->page0, head, (size_t)r.ps);
    jet_rc4(HDR_KEY, db->page0 + 0x18, r.jet4 ? 128 : 126);
    r.dbkey = U32(db->page0, 0x3E);
    r.pg = calloc((size_t)r.npages, sizeof(unsigned char *));
    r.loaded = malloc(sizeof(uint32_t) * (size_t)r.npages);
    r.ptype = malloc((size_t)r.npages);
    r.powner = malloc(sizeof(uint32_t) * (size_t)r.npages);
    {                                           /* the type and owner of each page: one pass, 256 pages at a time */
        unsigned char *chunk = malloc((size_t)r.ps * 256);
        fseek(f, 0, SEEK_SET);
        for (int p = 0; p < r.npages; p += 256) {
            int k = r.npages - p < 256 ? r.npages - p : 256;
            size_t got = fread(chunk, (size_t)r.ps, (size_t)k, f);
            for (int i = 0; i < k; i++) {
                unsigned char h8[8] = { 0 };
                if ((size_t)i < got) memcpy(h8, chunk + (size_t)i * (size_t)r.ps, 8);
                if (r.dbkey && p + i) jet_rc4(r.dbkey ^ (uint32_t)(p + i), h8, 8);   /* (RC4: its first 8 bytes) */
                r.ptype[p + i] = h8[0];
                r.powner[p + i] = U32(h8, 4);
            }
        }
        free(chunk);
    }
    jtable *cat = read_tdef(&r, 2, "MSysObjects");
    if (!cat) { jet_free(db); rd_close(&r); return 0; }
    read_rows(&r, cat);
    db->t = malloc(sizeof(jtable *) * 16);
    int capt = 16;
    db->t[db->nt++] = cat;
    jetcol *cid = jet_col(cat, "Id"), *cname = jet_col(cat, "Name"), *ctype = jet_col(cat, "Type"), *cflags = jet_col(cat, "Flags");
    if (!cid || !cname || !ctype) { jet_free(db); rd_close(&r); snprintf(err, (size_t)cap, "no catalog"); return 0; }
    /* the tables in the order of their TDEF pages */
    int nf = 0;
    uint32_t *pages = malloc(sizeof(uint32_t) * (size_t)(cat->nrows + 1)), *flags = malloc(sizeof(uint32_t) * (size_t)(cat->nrows + 1));
    char **names = malloc(sizeof(char *) * (size_t)(cat->nrows + 1));
    for (int i = 0; i < cat->nrows; i++) {
        jv *ty = jet_val(cat, i, ctype->num), *id = jet_val(cat, i, cid->num), *nm = jet_val(cat, i, cname->num);
        if (ty->null || ty->n < 2 || (short)U16(ty->p, 0) != 1 || id->null || id->n < 4 || nm->null) continue;
        uint32_t pg = U32(id->p, 0) & 0x00FFFFFF;
        if (pg == 2 || (int)pg >= r.npages) continue;
        char *s;
        if (r.jet4) {
            uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(nm->n / 2 + 1));
            for (int k = 0; k < nm->n / 2; k++) w[k] = (uint16_t)U16(nm->p, 2 * k);
            s = jet_utf16_to_utf8(w, nm->n / 2);
            free(w);
        } else {
            uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(nm->n + 1));
            int k = jet_cp_to_utf16(r.cp, nm->p, nm->n, w);
            s = jet_utf16_to_utf8(w, k);
            free(w);
        }
        jv *fl = cflags ? jet_val(cat, i, cflags->num) : 0;
        uint32_t f = fl && !fl->null && fl->n >= 4 ? U32(fl->p, 0) : 0;
        int at = nf;
        while (at > 0 && pages[at - 1] > pg) { pages[at] = pages[at - 1]; names[at] = names[at - 1]; flags[at] = flags[at - 1]; at--; }
        pages[at] = pg; names[at] = s; flags[at] = f; nf++;
    }
    for (int i = 0; i < nf; i++) {
        jtable *t = read_tdef(&r, pages[i], names[i]);
        free(names[i]);
        if (!t) { *err = 0; continue; }
        /* a table of the user (as the export tells them: not system, not MSys..., no system / hidden flags) */
        int user = t->type != 0x53 && strncasecmp(t->name, "MSys", 4) && !(flags[i] & 0x80000002u);
        if (!(opts & JET_NO_USER_ROWS) || !user) read_rows(&r, t);
        if (db->nt == capt) { capt *= 2; db->t = realloc(db->t, sizeof(jtable *) * (size_t)capt); }
        db->t[db->nt++] = t;
    }
    free(pages); free(names); free(flags);
    rd_close(&r);
    return db;
}
