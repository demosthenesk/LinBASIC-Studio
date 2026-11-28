/* jetwrite.c - the model of jetfile.h written as a new Jet 3 / Jet 4 file: page 0 (the original's, masked again),
 * the global usage map (page 1), the TDEFs (MSysObjects at page 2), the usage maps (type 0, or type 1 with pages of
 * type 5), the rows (Jet 3 with its jump table), the long values (in the row, one LVAL row or a chain), the indexes
 * (leaves and node levels, prefix compression, the bitmask of the entries). The TDEF pages are new: MSysObjects.Id,
 * MSysACEs.ObjectId and the foreign keys follow them. Written to a temporary file, then renamed.
 * The Python prototype: tools/jet/jetdb.py (Writer); the same file byte for byte (tools/jet/jetcmp.sh).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "jetfile.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#define U16(b, o) ((unsigned)(b)[o] | (unsigned)(b)[(o) + 1] << 8)
#define U32(b, o) ((uint32_t)(b)[o] | (uint32_t)(b)[(o) + 1] << 8 | (uint32_t)(b)[(o) + 2] << 16 | (uint32_t)(b)[(o) + 3] << 24)
static void P16(unsigned char *b, unsigned v) { b[0] = (unsigned char)v; b[1] = (unsigned char)(v >> 8); }
static void P32(unsigned char *b, uint32_t v) { b[0] = (unsigned char)v; b[1] = (unsigned char)(v >> 8); b[2] = (unsigned char)(v >> 16); b[3] = (unsigned char)(v >> 24); }

void jet_rc4(uint32_t key, unsigned char *data, int n);
#define HDR_KEY 0x6B39DAC7u
/* a long value is kept in its row up to 32 bytes (Jet 3) / 34 (Jet 4), as Jet writes them; Jet 3 reads a longer
 * one there as a damaged record (3049; PSC files of DEV-14/q1) */
#define INLINE_MAX(w) ((w)->jet4 ? 34 : 32)

static const int FIXLEN[17] = { 0, 0, 1, 2, 4, 8, 4, 8, 8, 0, 0, 0, 0, 0, 0, 16, 17 };
static int fixlen(const jetcol *c) { return c->type < 17 && FIXLEN[c->type] ? FIXLEN[c->type] : c->len; }
static int is_lv(int type) { return type == JT_OLE || type == JT_MEMO; }

/* ---------------------------------------------------------------- growing bytes */
typedef struct bb { unsigned char *p; int n, cap; } bb;
static void bput(bb *b, const void *p, int n)
{
    if (b->n + n > b->cap) { b->cap = (b->n + n) * 2 + 64; b->p = realloc(b->p, (size_t)b->cap); }
    if (n) memcpy(b->p + b->n, p, (size_t)n);
    b->n += n;
}
static void b8(bb *b, int v) { unsigned char c = (unsigned char)v; bput(b, &c, 1); }
static void b16(bb *b, unsigned v) { unsigned char c[2]; P16(c, v); bput(b, c, 2); }
static void b32(bb *b, uint32_t v) { unsigned char c[4]; P32(c, v); bput(b, c, 4); }
static void bzero_(bb *b, int n) { static const unsigned char z[64]; while (n > 0) { int m = n > 64 ? 64 : n; bput(b, z, m); n -= m; } }

/* ---------------------------------------------------------------- the writer */
typedef struct lvpage { uint32_t page; bb *rows; int nrows, caprows, free, done; } lvpage;
typedef struct lvcol { int num; lvpage *pages; int npages, cap; int cur; } lvcol;
typedef struct maprow { uint32_t *p; int n; } maprow;

typedef struct wr {
    jdb *db;
    int jet4, ps, dhdr, maplen, ixmask, ixfirst;
    unsigned char **pages; int cappages;
    uint32_t next;
    uint32_t *mappages; int nmappages, mapcap;
    maprow *maps; int nmaps, capmaps;
    lvcol *lv; int nlv;                         /* of the table being written */
    char *err; int errcap;
    FILE *f; uint32_t dbkey; int ioerr;         /* the file being written: the pages done go to it (spill) */
} wr;
static unsigned char spilled_mark;
#define SPILLED (&spilled_mark)                 /* a page already in the file */

static uint32_t walloc(wr *w)
{
    uint32_t p = w->next++;
    if ((int)w->next > w->cappages) {
        int nc = w->cappages ? w->cappages * 2 : 256;
        while (nc < (int)w->next) nc *= 2;
        w->pages = realloc(w->pages, sizeof(unsigned char *) * (size_t)nc);
        for (int i = w->cappages; i < nc; i++) w->pages[i] = 0;
        w->cappages = nc;
    }
    return p;
}
static unsigned char *wpage(wr *w, uint32_t p)
{
    if (w->pages[p] == SPILLED) { w->ioerr = 1; w->pages[p] = 0; }    /* (never: a page is done when spilled) */
    if (!w->pages[p]) w->pages[p] = calloc(1, (size_t)w->ps);
    return w->pages[p];
}
/* a page written for good (its data, LVAL and index pages): into the file at its place, out of memory -- a database of
 * hundreds of MB is never all in memory (encrypted: the page's key, as below) */
static void put_page(wr *w, uint32_t p, const unsigned char *src)
{
    static unsigned char *enc; static int encn;
    if (w->dbkey && p) {
        if (encn < w->ps) { free(enc); enc = malloc((size_t)w->ps); encn = w->ps; }
        memcpy(enc, src, (size_t)w->ps); jet_rc4(w->dbkey ^ p, enc, w->ps); src = enc;
    }
    if (fseek(w->f, (long)p * w->ps, SEEK_SET) || fwrite(src, 1, (size_t)w->ps, w->f) != (size_t)w->ps) w->ioerr = 1;
}
static void spill(wr *w, uint32_t p)
{
    if (!w->f || !w->pages[p] || w->pages[p] == SPILLED) return;
    put_page(w, p, w->pages[p]);
    free(w->pages[p]);
    w->pages[p] = SPILLED;
}

/* ---------------------------------------------------------------- text */
/* Jet 4 compressed unicode: ff fe, single bytes for characters below 0x100, pairs for the rest, 00 switching;
 * 0 when it does not make the text shorter (or holds a character 0) */
static int compress_text(const unsigned char *t, int n, bb *out)
{
    int units = n / 2;
    for (int i = 0; i < units; i++) if (!t[2 * i] && !t[2 * i + 1]) return 0;
    bb o = { 0 };
    bput(&o, "\xff\xfe", 2);
    int comp = 1;
    for (int i = 0; i < units; i++) {
        int small = t[2 * i + 1] == 0;
        if (small != comp) { b8(&o, 0); comp = small; }
        if (small) b8(&o, t[2 * i]); else bput(&o, t + 2 * i, 2);
    }
    if (o.n < units * 2) { *out = o; return 1; }
    free(o.p);
    return 0;
}
static void enc_text(wr *w, const jetcol *c, const jv *v, bb *out)
{
    out->n = 0;
    if (w->jet4 && (c->xflags & 1) && (c->type == JT_TEXT || v->n / 2 <= 1024)) {
        bb z = { 0 };
        if (compress_text(v->p, v->n, &z)) { free(out->p); *out = z; return; }
    }
    bput(out, v->p, v->n);
}

/* ---------------------------------------------------------------- long values */
static lvcol *lv_of(wr *w, int num)
{
    for (int i = 0; i < w->nlv; i++) if (w->lv[i].num == num) return &w->lv[i];
    return 0;
}
static int data_page(wr *w, uint32_t pg, const unsigned char owner[4], bb *rows, int n);
/* an LVAL page no longer filled: written, its rows freed (memory: a database of large memos) */
static void lv_done(wr *w, lvpage *pg)
{
    if (pg->done) return;
    data_page(w, pg->page, (const unsigned char *)"LVAL", pg->rows, pg->nrows);
    spill(w, pg->page);
    for (int q = 0; q < pg->nrows; q++) free(pg->rows[q].p);
    free(pg->rows);
    pg->rows = 0; pg->caprows = 0; pg->done = 1;
}
/* a row on an LVAL page of the column (a new page when it does not fit, or for each part of a chain) */
static uint32_t lv_put(wr *w, lvcol *lv, const unsigned char *row, int n, int whole)
{
    int need = n + 2;
    lvpage *cur = lv->cur >= 0 ? &lv->pages[lv->cur] : 0;
    if (whole || !cur || cur->free < need || cur->nrows >= 255) {
        if (cur) lv_done(w, cur);               /* (the page filled until now: never again) */
        if (lv->npages == lv->cap) { lv->cap = lv->cap ? lv->cap * 2 : 8; lv->pages = realloc(lv->pages, sizeof(lvpage) * (size_t)lv->cap); }
        cur = &lv->pages[lv->npages];
        memset(cur, 0, sizeof *cur);
        cur->page = walloc(w);
        cur->free = w->ps - w->dhdr;
        lv->cur = whole ? -1 : lv->npages;
        lv->npages++;
    }
    if (cur->nrows == cur->caprows) { cur->caprows = cur->caprows ? cur->caprows * 2 : 8; cur->rows = realloc(cur->rows, sizeof(bb) * (size_t)cur->caprows); }
    bb *r = &cur->rows[cur->nrows++];
    memset(r, 0, sizeof *r);
    bput(r, row, n);
    cur->free -= need;
    uint32_t at = cur->page << 8 | (uint32_t)(cur->nrows - 1);
    if (whole) lv_done(w, cur);                 /* a part of a chain: a page of its own */
    return at;
}
static void lv_header(wr *w, const jetcol *c, const unsigned char *data, int n, bb *out)
{
    lvcol *lv = lv_of(w, c->num);
    int one = w->ps - w->dhdr - 2 - (w->jet4 ? 4 : 0);   /* the longest LVAL row (measured: 4076 / 2036) */
    out->n = 0;
    if (n <= one) {
        uint32_t dp = lv_put(w, lv, data, n, 0);
        b32(out, (uint32_t)n | 0x40000000u); b32(out, dp); b32(out, 0);
        return;
    }
    int seg = one - 4, nparts = (n + seg - 1) / seg;
    uint32_t nxt = 0;
    unsigned char *tmp = malloc((size_t)seg + 4);
    for (int k = nparts - 1; k >= 0; k--) {     /* the last part first: each knows the pointer of the next */
        int m = k == nparts - 1 ? n - k * seg : seg;
        P32(tmp, nxt);
        memcpy(tmp + 4, data + (size_t)k * (size_t)seg, (size_t)m);
        nxt = lv_put(w, lv, tmp, m + 4, 1);
    }
    free(tmp);
    b32(out, (uint32_t)n); b32(out, nxt); b32(out, 0);
}

/* ---------------------------------------------------------------- rows */
static void layout(jtable *t)
{
    if (t->keep_layout) return;
    /* by number: a fixed column takes the next place of the fixed part (a Yes/No none: 0) and as offv the count of
     * the variable columns before it; a variable column its place among them (as Jet writes them) */
    int o = 0, k = 0;
    for (int num = 0; num < t->width; num++)
        for (int i = 0; i < t->ncols; i++) {
            jetcol *c = &t->cols[i];
            if (c->num != num) continue;
            c->offv = k;
            if (c->flags & 1) { c->offf = c->type == JT_BOOL ? 0 : o; if (c->type != JT_BOOL) o += fixlen(c); }
            else k++;
        }
}

static void make_row(wr *w, jtable *t, int nb, const bb *fixed, bb *var, const unsigned char *mask, bb *out)
{
    out->n = 0;
    if (w->jet4) b16(out, (unsigned)t->maxcols); else b8(out, t->maxcols);
    bput(out, fixed->p, fixed->n);
    if (!t->nvar) { bput(out, mask, nb); return; }
    int *offs = malloc(sizeof(int) * (size_t)(t->nvar + 1));
    int bodystart = out->n;
    (void)bodystart;
    for (int k = 0; k < t->nvar; k++) { offs[k] = out->n; bput(out, var[k].p, var[k].n); }
    offs[t->nvar] = out->n;
    if (w->jet4) {
        for (int k = t->nvar; k >= 0; k--) b16(out, (unsigned)offs[k]);
        b16(out, (unsigned)t->nvar);
        bput(out, mask, nb);
    } else {
        int bodylen = out->n, nofs = t->nvar + 1, njump = 0;
        for (;;) {
            int rowlen = bodylen + nofs + njump + 1 + nb, need = (rowlen - 1) / 256;
            if (need == njump) break;
            njump = need;
        }
        for (int k = t->nvar; k >= 0; k--) b8(out, offs[k] & 0xFF);
        for (int j = njump - 1; j >= 0; j--) {
            int lim = 256 * (j + 1), jv_ = 0xFF;
            for (int k = 0; k < nofs; k++) if (offs[k] >= lim) { jv_ = k; break; }
            b8(out, jv_);
        }
        b8(out, t->nvar);
        bput(out, mask, nb);
    }
    free(offs);
}

static void encode_row(wr *w, jtable *t, int r, bb *out)
{
    int nb = (t->maxcols + 7) / 8;
    unsigned char *mask = calloc((size_t)nb + 1, 1);
    int fsize = 0;
    for (int i = 0; i < t->ncols; i++) {
        jetcol *c = &t->cols[i];
        if ((c->flags & 1) && c->type != JT_BOOL && c->offf + fixlen(c) > fsize) fsize = c->offf + fixlen(c);
    }
    bb fixed = { 0 };
    bzero_(&fixed, fsize);
    bb *var = calloc((size_t)(t->nvar + 1), sizeof(bb));
    int *small = malloc(sizeof(int) * (size_t)(t->ncols + 1)), nsmall = 0;
    for (int i = 0; i < t->ncols; i++) {
        jetcol *c = &t->cols[i];
        jv *v = jet_val(t, r, c->num);
        if (c->type == JT_BOOL) { if (v->n) mask[c->num / 8] |= (unsigned char)(1 << (c->num % 8)); continue; }
        if (v->null) continue;
        mask[c->num / 8] |= (unsigned char)(1 << (c->num % 8));
        if (c->flags & 1) {
            int n = fixlen(c), m = v->n < n ? v->n : n;
            memcpy(fixed.p + c->offf, v->p, (size_t)m);
            if (m < n) memset(fixed.p + c->offf + m, 0, (size_t)(n - m));
        } else if (c->offv < t->nvar) {
            bb *dst = &var[c->offv];
            if (c->type == JT_TEXT) enc_text(w, c, v, dst);
            else if (is_lv(c->type)) {
                bb data = { 0 };
                if (c->type == JT_MEMO) enc_text(w, c, v, &data); else bput(&data, v->p, v->n);
                if (data.n <= INLINE_MAX(w) && t->type != 0x53) { small[nsmall++] = i; free(data.p); }
                else lv_header(w, c, data.p, data.n, dst), free(data.p);
            } else { dst->n = 0; bput(dst, v->p, v->n); }
        }
    }
    for (int s = 0; s < nsmall; s++) {
        jetcol *c = &t->cols[small[s]];
        bb data = { 0 };
        if (c->type == JT_MEMO) enc_text(w, c, jet_val(t, r, c->num), &data); else bput(&data, jet_val(t, r, c->num)->p, jet_val(t, r, c->num)->n);
        bb *dst = &var[c->offv];
        dst->n = 0;
        b32(dst, (uint32_t)data.n | 0x80000000u); b32(dst, 0); b32(dst, 0);
        bput(dst, data.p, data.n);
        free(data.p);
    }
    make_row(w, t, nb, &fixed, var, mask, out);
    if (nsmall && !w->jet4 && out->n > 255) {
        /* Jet 3 reads a long value kept in its row only in a row shorter than 256 bytes */
        for (int s = 0; s < nsmall; s++) {
            jetcol *c = &t->cols[small[s]];
            bb data = { 0 };
            if (c->type == JT_MEMO) enc_text(w, c, jet_val(t, r, c->num), &data); else bput(&data, jet_val(t, r, c->num)->p, jet_val(t, r, c->num)->n);
            lv_header(w, c, data.p, data.n, &var[c->offv]);
            free(data.p);
        }
        make_row(w, t, nb, &fixed, var, mask, out);
    }
    for (int k = 0; k <= t->nvar; k++) free(var[k].p);
    free(var); free(fixed.p); free(mask); free(small);
}

/* ---------------------------------------------------------------- pages */
static int data_page(wr *w, uint32_t pg, const unsigned char owner[4], bb *rows, int n)
{
    unsigned char *b = wpage(w, pg);
    int h = w->dhdr;
    b[0] = 1; b[1] = 1;
    memcpy(b + 4, owner, 4);
    P16(b + h - 2, (unsigned)n);
    int end = w->ps;
    for (int i = 0; i < n; i++) {
        int start = end - rows[i].n;
        memcpy(b + start, rows[i].p, (size_t)rows[i].n);
        P16(b + h + 2 * i, (unsigned)start);
        end = start;
    }
    P16(b + 2, (unsigned)(end - (h + 2 * n)));
    return end - (h + 2 * n);
}

/* ---------------------------------------------------------------- indexes */
typedef struct ent { unsigned char *p; int n; } ent;
static int ent_cmp(const void *a, const void *b)
{
    const ent *x = a, *y = b;
    int m = x->n < y->n ? x->n : y->n, c = memcmp(x->p, y->p, (size_t)m);
    return c ? c : x->n - y->n;
}
static int common(const ent *a, const ent *b)
{
    int m = a->n < b->n ? a->n : b->n, i = 0;
    while (i < m && a->p[i] == b->p[i]) i++;
    return i;
}
static int prefix_of(const ent *first, const ent *e, int tail)
{
    int c = common(first, e);
    if (first->n - tail < c) c = first->n - tail;
    if (e->n - tail < c) c = e->n - tail;
    return c;
}
/* split entries into pages: [start, count] pairs */
static int group(const ent *items, int n, int cap, int tail, int **starts)
{
    int ng = 0, capg = 16;
    *starts = malloc(sizeof(int) * (size_t)capg * 2);
    int cs = -1, cn = 0, sum = 0, pref = 0;
    for (int i = 0; i < n; i++) {
        const ent *x = &items[i];
        if (cn) {
            int np = prefix_of(&items[cs], x, tail);
            if (pref < np) np = pref;
            if (sum + x->n - np * cn > cap) {
                if (ng == capg) { capg *= 2; *starts = realloc(*starts, sizeof(int) * (size_t)capg * 2); }
                (*starts)[2 * ng] = cs; (*starts)[2 * ng + 1] = cn; ng++;
                cn = 0;
            } else { pref = np; cn++; sum += x->n; continue; }
        }
        cs = i; cn = 1; sum = x->n; pref = x->n;
    }
    if (cn) {
        if (ng == capg) { capg *= 2; *starts = realloc(*starts, sizeof(int) * (size_t)capg * 2); }
        (*starts)[2 * ng] = cs; (*starts)[2 * ng + 1] = cn; ng++;
    }
    return ng;
}
static void ix_page(wr *w, uint32_t pg, uint32_t tdef, int typ, const ent *ents, int n, uint32_t prev, uint32_t nxt, uint32_t tail, int level)
{
    unsigned char *b = wpage(w, pg);
    b[0] = (unsigned char)typ; b[1] = 1;
    int pref = 0;
    if (n > 1) {
        pref = 1 << 30;
        for (int i = 1; i < n; i++) { int p = prefix_of(&ents[0], &ents[i], typ == 4 ? 4 : 8); if (p < pref) pref = p; }
    }
    int o = w->ixfirst;
    int mbits = (w->ixfirst - w->ixmask) * 8;
    for (int i = 0; i < n; i++) {
        const unsigned char *part = i ? ents[i].p + pref : ents[i].p;
        int len = i ? ents[i].n - pref : ents[i].n;
        if (i) { int bit = o - w->ixfirst; if (bit < mbits) b[w->ixmask + bit / 8] |= (unsigned char)(1 << (bit % 8)); }
        if (o + len <= w->ps) memcpy(b + o, part, (size_t)len);
        o += len;
    }
    if (n) { int bit = o - w->ixfirst; if (bit < mbits) b[w->ixmask + bit / 8] |= (unsigned char)(1 << (bit % 8)); }
    P16(b + 2, (unsigned)(w->ps - o));
    P32(b + 4, tdef);
    if (w->jet4) { P32(b + 12, prev); P32(b + 16, nxt); P32(b + 20, tail); P16(b + 24, (unsigned)pref); b[26] = (unsigned char)level; }
    else { P32(b + 8, prev); P32(b + 12, nxt); P32(b + 16, tail); b[20] = (unsigned char)pref; b[21] = (unsigned char)level; }
}

/* the pages of one index from its sorted entries; its root; all its pages in *pages */
static uint32_t index_pages(wr *w, jtable *t, ent *entries, int n, uint32_t **pages, int *npages)
{
    int cap = w->ps - w->ixfirst;
    *npages = 0;
    *pages = 0;
    int capp = 0;
    /* the levels: [pages, entries of each page (start, count) into the level's items] */
    typedef struct lvl { ent *items; int nitems; int *grp; int ng; uint32_t *pg; } lvl;
    lvl levels[16];
    int nl = 0;
    levels[0].items = entries; levels[0].nitems = n;
    if (n) levels[0].ng = group(entries, n, cap, 4, &levels[0].grp);
    else { levels[0].grp = calloc(2, sizeof(int)); levels[0].ng = 1; }
    levels[0].pg = malloc(sizeof(uint32_t) * (size_t)levels[0].ng);
    for (int g = 0; g < levels[0].ng; g++) levels[0].pg[g] = walloc(w);
    nl = 1;
    while (levels[nl - 1].ng > 1 && nl < 16) {
        lvl *lo = &levels[nl - 1], *up = &levels[nl];
        up->nitems = lo->ng;
        up->items = malloc(sizeof(ent) * (size_t)lo->ng);
        for (int g = 0; g < lo->ng; g++) {
            const ent *last = &lo->items[lo->grp[2 * g] + lo->grp[2 * g + 1] - 1];
            up->items[g].n = last->n + 4;
            up->items[g].p = malloc((size_t)last->n + 4);
            memcpy(up->items[g].p, last->p, (size_t)last->n);
            uint32_t c = lo->pg[g];
            up->items[g].p[last->n] = (unsigned char)(c >> 24); up->items[g].p[last->n + 1] = (unsigned char)(c >> 16);
            up->items[g].p[last->n + 2] = (unsigned char)(c >> 8); up->items[g].p[last->n + 3] = (unsigned char)c;
        }
        up->ng = group(up->items, up->nitems, cap, 8, &up->grp);
        up->pg = malloc(sizeof(uint32_t) * (size_t)up->ng);
        for (int g = 0; g < up->ng; g++) up->pg[g] = walloc(w);
        nl++;
    }
    for (int d = 0; d < nl; d++) {
        lvl *L = &levels[d];
        uint32_t tail = d ? levels[d - 1].pg[levels[d - 1].ng - 1] : 0;
        for (int g = 0; g < L->ng; g++) {
            uint32_t prev = g ? L->pg[g - 1] : 0, nxt = g + 1 < L->ng ? L->pg[g + 1] : 0;
            int start = L->grp[2 * g], cnt = L->grp[2 * g + 1];
            ix_page(w, L->pg[g], t->page, d == 0 ? 4 : 3, L->nitems ? L->items + start : 0, L->nitems ? cnt : 0, prev, nxt, tail, d);
            spill(w, L->pg[g]);
            if (*npages == capp) { capp = capp ? capp * 2 : 16; *pages = realloc(*pages, sizeof(uint32_t) * (size_t)capp); }
            (*pages)[(*npages)++] = L->pg[g];
        }
    }
    uint32_t root = levels[nl - 1].pg[0];
    for (int d = 0; d < nl; d++) {
        if (d) { for (int i = 0; i < levels[d].nitems; i++) free(levels[d].items[i].p); free(levels[d].items); }
        free(levels[d].grp); free(levels[d].pg);
    }
    return root;
}

/* ---------------------------------------------------------------- usage maps */
static int u32cmp(const void *a, const void *b) { uint32_t x = *(const uint32_t *)a, y = *(const uint32_t *)b; return x < y ? -1 : x > y; }
static uint32_t add_map(wr *w, const uint32_t *pages, int n)
{
    if (w->nmaps == w->capmaps) { w->capmaps = w->capmaps ? w->capmaps * 2 : 64; w->maps = realloc(w->maps, sizeof(maprow) * (size_t)w->capmaps); }
    maprow *m = &w->maps[w->nmaps];
    m->p = malloc(sizeof(uint32_t) * (size_t)(n + 1));
    if (n) memcpy(m->p, pages, sizeof(uint32_t) * (size_t)n);
    m->n = n;
    qsort(m->p, (size_t)n, sizeof(uint32_t), u32cmp);
    int k = w->nmaps++;
    return w->mappages[k / w->mapcap] << 8 | (uint32_t)(k % w->mapcap);
}
static int map_row(wr *w, const maprow *m, bb *out)
{
    out->n = 0;
    if (!m->n) { b8(out, 0); b32(out, 0); bzero_(out, w->maplen); return 0; }
    uint32_t lo = m->p[0], hi = m->p[m->n - 1], start = lo & ~7u;
    if (hi - start < (uint32_t)w->maplen * 8) {
        b8(out, 0); b32(out, start);
        int at = out->n;
        bzero_(out, w->maplen);
        for (int i = 0; i < m->n; i++) { uint32_t q = m->p[i] - start; out->p[at + q / 8] |= (unsigned char)(1 << (q % 8)); }
        return 0;
    }
    /* type 1: pointers to pages of type 5, each mapping (page size - 4) * 8 pages */
    uint32_t per = (uint32_t)(w->ps - 4) * 8;
    int nptr = (w->maplen + 4) / 4;
    uint32_t *ptrs = calloc((size_t)nptr, sizeof(uint32_t));
    for (int i = 0; i < m->n; i++) {
        uint32_t blk = m->p[i] / per;
        if ((int)blk >= nptr) { free(ptrs); snprintf(w->err, (size_t)w->errcap, "a usage map past what the format holds"); return -1; }
        if (!ptrs[blk]) {                       /* in the order the pages are met (Python: by block, first seen) */
            ptrs[blk] = walloc(w);
            unsigned char *b = wpage(w, ptrs[blk]);
            b[0] = 5; b[1] = 1;
        }
        uint32_t q = m->p[i] - blk * per;
        wpage(w, ptrs[blk])[4 + q / 8] |= (unsigned char)(1 << (q % 8));
    }
    b8(out, 1);
    for (int i = 0; i < nptr; i++) b32(out, ptrs[i]);
    free(ptrs);
    return 0;
}

/* ---------------------------------------------------------------- TDEF */
static int lidx_cmp_db;
static jdb *lidx_db;
static int lidx_order_cmp(const void *a, const void *b)
{
    const jlidx *x = *(const jlidx *const *)a, *y = *(const jlidx *const *)b;
    jkey ka = { 0 }, kb = { 0 };
    uint16_t *wa, *wb;
    int na = jet_utf8_to_utf16(x->name, &wa), nb = jet_utf8_to_utf16(y->name, &wb);
    if (lidx_db->jet4) { jet_text_key(&ka, (unsigned char *)wa, na * 2, 1); jet_text_key(&kb, (unsigned char *)wb, nb * 2, 1); }
    else {
        unsigned char *ba = malloc((size_t)na * 2 + 1), *bbb = malloc((size_t)nb * 2 + 1);
        int la = jet_utf16_to_cp(lidx_db->codepage, wa, na, ba), lb = jet_utf16_to_cp(lidx_db->codepage, wb, nb, bbb);
        jet_text_key(&ka, ba, la, 0); jet_text_key(&kb, bbb, lb, 0);
        free(ba); free(bbb);
    }
    int m = ka.n < kb.n ? ka.n : kb.n, c = memcmp(ka.p, kb.p, (size_t)m);
    if (!c) c = ka.n - kb.n;
    free(ka.p); free(kb.p); free(wa); free(wb);
    return c;
}
static jlidx **lidx_order(jdb *db, jtable *t)
{
    jlidx **o = malloc(sizeof(jlidx *) * (size_t)(t->nlidx + 1));
    for (int i = 0; i < t->nlidx; i++) o[i] = &t->lidx[i];
    lidx_db = db; (void)lidx_cmp_db;
    /* a stable sort: insertion (few indexes) */
    for (int i = 1; i < t->nlidx; i++) {
        jlidx *x = o[i];
        int j = i;
        while (j > 0 && lidx_order_cmp(&o[j - 1], &x) > 0) { o[j] = o[j - 1]; j--; }
        o[j] = x;
    }
    return o;
}

static void put_name(wr *w, bb *b, const char *name)
{
    uint16_t *u;
    int n = jet_utf8_to_utf16(name, &u);
    if (w->jet4) {
        b16(b, (unsigned)n * 2);
        for (int i = 0; i < n; i++) b16(b, u[i]);
    } else {
        unsigned char *c = malloc((size_t)n * 2 + 1);
        int m = jet_utf16_to_cp(w->db->codepage, u, n, c);
        if (m > 255) m = 255;
        b8(b, m);
        bput(b, c, m);
        free(c);
    }
    free(u);
}

static void tdef_bytes(wr *w, jtable *t, uint32_t used, uint32_t freem, const uint32_t *ixu, const uint32_t *ixf, const uint32_t *lvu, const uint32_t *lvf, bb *b)
{
    b->n = 0;
    bzero_(b, 8);
    uint32_t nrows = (uint32_t)t->nrows;
    if (w->jet4) {
        b32(b, 0); b32(b, 1625); b32(b, nrows); b32(b, t->autonum);
        b8(b, t->autonum_flag); b8(b, 0); b8(b, 0); b8(b, 0);
        b32(b, t->ct_autonum); bzero_(b, 8);
        b8(b, t->type); b16(b, (unsigned)t->maxcols); b16(b, (unsigned)t->nvar); b16(b, (unsigned)t->ncols);
        b32(b, (uint32_t)t->nlidx); b32(b, (uint32_t)t->nridx); b32(b, used); b32(b, freem);
        for (int i = 0; i < t->nridx; i++) { b32(b, 0); b32(b, nrows); b32(b, 0); }
        for (int i = 0; i < t->ncols; i++) {
            jetcol *c = &t->cols[i];
            unsigned char raw[25] = { 0 };
            if (c->rawlen == 25) memcpy(raw, c->raw, 25); else P32(raw + 1, 1625);
            raw[0] = (unsigned char)c->type;
            P16(raw + 5, (unsigned)c->num); P16(raw + 7, (unsigned)c->offv); P16(raw + 9, (unsigned)c->id);
            P16(raw + 11, (unsigned)c->misc); P16(raw + 13, (unsigned)c->misc2);
            raw[15] = (unsigned char)c->flags; raw[16] = (unsigned char)c->xflags;
            P16(raw + 21, (unsigned)c->offf); P16(raw + 23, (unsigned)c->len);
            bput(b, raw, 25);
        }
        for (int i = 0; i < t->ncols; i++) put_name(w, b, t->cols[i].name);
        for (int i = 0; i < t->nridx; i++) {
            jridx *r = &t->ridx[i];
            b32(b, 1923);
            for (int k = 0; k < 10; k++) {
                if (k < r->n) { b16(b, (unsigned)r->col[k]); b8(b, r->order[k]); }
                else { b16(b, 0xFFFF); b8(b, 0); }
            }
            b32(b, ixu[i]); b32(b, ixf[i]); b32(b, 0); b8(b, r->flags); b8(b, r->complex); b32(b, 0);
        }
        jlidx **o = lidx_order(w->db, t);
        for (int i = 0; i < t->nlidx; i++) {
            jlidx *l = o[i];
            b32(b, 1625); b32(b, (uint32_t)l->num); b32(b, (uint32_t)l->real); b8(b, l->reltype); b32(b, l->relidx);
            b32(b, l->reltable); b8(b, l->cup); b8(b, l->cdel); b8(b, l->type); b32(b, 0);
        }
        for (int i = 0; i < t->nlidx; i++) put_name(w, b, o[i]->name);
        free(o);
    } else {
        b32(b, 0); b32(b, nrows); b32(b, t->autonum);
        b8(b, t->type); b16(b, (unsigned)t->maxcols); b16(b, (unsigned)t->nvar); b16(b, (unsigned)t->ncols);
        b32(b, (uint32_t)t->nlidx); b32(b, (uint32_t)t->nridx); b32(b, used); b32(b, freem);
        for (int i = 0; i < t->nridx; i++) { b32(b, 0); b32(b, nrows); }
        for (int i = 0; i < t->ncols; i++) {
            jetcol *c = &t->cols[i];
            unsigned char raw[18] = { 0 };
            if (c->rawlen == 18) memcpy(raw, c->raw, 18);
            raw[0] = (unsigned char)c->type;
            P16(raw + 1, (unsigned)c->num); P16(raw + 3, (unsigned)c->offv); P16(raw + 5, (unsigned)c->id); P16(raw + 7, (unsigned)c->fileflag);
            P16(raw + 9, (unsigned)c->misc); P16(raw + 11, (unsigned)c->misc2);
            raw[13] = (unsigned char)c->flags;
            P16(raw + 14, (unsigned)c->offf); P16(raw + 16, (unsigned)c->len);
            bput(b, raw, 18);
        }
        for (int i = 0; i < t->ncols; i++) put_name(w, b, t->cols[i].name);
        for (int i = 0; i < t->nridx; i++) {
            jridx *r = &t->ridx[i];
            for (int k = 0; k < 10; k++) {
                if (k < r->n) { b16(b, (unsigned)r->col[k]); b8(b, r->order[k]); }
                else { b16(b, 0xFFFF); b8(b, 0); }
            }
            b32(b, ixu[i]); b32(b, ixf[i]); b8(b, r->flags);
        }
        jlidx **o = lidx_order(w->db, t);
        for (int i = 0; i < t->nlidx; i++) {
            jlidx *l = o[i];
            b32(b, (uint32_t)l->num); b32(b, (uint32_t)l->real); b8(b, l->reltype); b32(b, l->relidx);
            b32(b, l->reltable); b8(b, l->cup); b8(b, l->cdel); b8(b, l->type);
        }
        for (int i = 0; i < t->nlidx; i++) put_name(w, b, o[i]->name);
        free(o);
    }
    for (int i = 0; i < t->nlv; i++) { b16(b, (unsigned)t->lvorder[i]); b32(b, lvu ? lvu[i] : 0); b32(b, lvf ? lvf[i] : 0); }
    b16(b, 0xFFFF); bzero_(b, 8);
    P32(b->p + 8, (uint32_t)(b->n - 8));
}
static int tdef_npages(wr *w, jtable *t)
{
    bb b = { 0 };
    uint32_t *z = calloc((size_t)(t->nridx + t->nlv + 1), sizeof(uint32_t));
    tdef_bytes(w, t, 0, 0, z, z, z, z, &b);
    int n = b.n - 8, room = w->ps - 8, k = (n + room - 1) / room;
    free(b.p); free(z);
    return k < 1 ? 1 : k;
}
static void tdef_pages(wr *w, const bb *data, const uint32_t *pages, int np)
{
    int room = w->ps - 8, body = data->n - 8;
    for (int k = 0; k < np; k++) {
        unsigned char *b = wpage(w, pages[k]);
        int at = k * room, len = body - at;
        if (len > room) len = room;
        if (len < 0) len = 0;
        b[0] = 2; b[1] = 1;
        if (w->jet4) P16(b + 2, (unsigned)(w->ps - 8 - len)); else { b[2] = 'V'; b[3] = 'C'; }
        P32(b + 4, k + 1 < np ? pages[k + 1] : 0);
        memcpy(b + 8, data->p + 8 + at, (size_t)len);
    }
}

/* ---------------------------------------------------------------- the catalog follows the TDEFs */
static void remap_catalog(wr *w, const uint32_t *oldp, const uint32_t *newp, int n)
{
    jtable *cat = w->db->t[0];
    jetcol *cid = jet_col(cat, "Id"), *cty = jet_col(cat, "Type");
    uint32_t *from = malloc(sizeof(uint32_t) * (size_t)(n + 1)), *to = malloc(sizeof(uint32_t) * (size_t)(n + 1));
    int nm = 0;
    for (int r = 0; r < cat->nrows; r++) {
        jv *ty = jet_val(cat, r, cty->num), *id = jet_val(cat, r, cid->num);
        if (ty->null || ty->n < 2 || (short)U16(ty->p, 0) != 1 || id->null || id->n < 4) continue;
        uint32_t old = U32(id->p, 0);
        for (int i = 0; i < n; i++)
            if (oldp[i] == old) { P32(id->p, newp[i]); from[nm] = old; to[nm] = newp[i]; nm++; break; }
    }
    jtable *aces = jet_table(w->db, "MSysACEs");
    jetcol *oc = aces ? jet_col(aces, "ObjectId") : 0;
    if (oc)
        for (int r = 0; r < aces->nrows; r++) {
            jv *v = jet_val(aces, r, oc->num);
            if (v->null || v->n < 4) continue;
            uint32_t o = U32(v->p, 0);
            for (int i = 0; i < nm; i++) if (from[i] == o) { P32(v->p, to[i]); break; }
        }
    free(from); free(to);
}

/* ---------------------------------------------------------------- the whole file */
static int build(wr *w)
{
    jdb *db = w->db;
    int nt = db->nt;
    walloc(w); walloc(w);                       /* page 0: definition, page 1: the global usage map */
    for (int i = 0; i < nt; i++) {
        jtable *t = db->t[i];
        int mx = 0, nv = 0;
        for (int k = 0; k < t->ncols; k++) { if (t->cols[k].num + 1 > mx) mx = t->cols[k].num + 1; if (!(t->cols[k].flags & 1)) nv++; }
        if (t->maxcols < 0) t->maxcols = mx;
        if (t->nvar < 0) t->nvar = nv;
        if (!t->lvorder) {
            t->lvorder = malloc(sizeof(int) * (size_t)(t->ncols + 1));
            t->nlv = 0;
            for (int k = 0; k < t->ncols; k++) if (is_lv(t->cols[k].type)) t->lvorder[t->nlv++] = t->cols[k].num;
        }
        if (t->width < mx) {                    /* the rows: wide enough for every column number */
            int nw = mx;
            jv *nvals = calloc((size_t)(t->caprows ? t->caprows : 1) * (size_t)nw, sizeof(jv));
            for (int r = 0; r < t->nrows; r++) {
                for (int c = 0; c < nw; c++) nvals[(size_t)r * nw + c].null = 1;
                for (int c = 0; c < t->width; c++) nvals[(size_t)r * nw + c] = t->vals[(size_t)r * t->width + c];
            }
            free(t->vals);
            t->vals = nvals; t->width = nw;
        }
        layout(t);
    }
    uint32_t **tp = calloc((size_t)nt, sizeof(uint32_t *));
    int *ntp = calloc((size_t)nt, sizeof(int));
    for (int i = 0; i < nt && i < 4; i++) {
        ntp[i] = tdef_npages(w, db->t[i]);
        tp[i] = calloc((size_t)ntp[i], sizeof(uint32_t));
        tp[i][0] = walloc(w);
    }
    int nmaps = 0;
    for (int i = 0; i < nt; i++) nmaps += 2 + db->t[i]->nridx + 2 * db->t[i]->nlv;
    w->mapcap = (w->ps - w->dhdr) / (5 + w->maplen + 2);
    w->nmappages = (nmaps + w->mapcap - 1) / w->mapcap;
    w->mappages = malloc(sizeof(uint32_t) * (size_t)(w->nmappages + 1));
    for (int i = 0; i < w->nmappages; i++) w->mappages[i] = walloc(w);
    for (int i = 0; i < nt && i < 4; i++) for (int k = 1; k < ntp[i]; k++) tp[i][k] = walloc(w);
    for (int i = 4; i < nt; i++) {
        ntp[i] = tdef_npages(w, db->t[i]);
        tp[i] = calloc((size_t)ntp[i], sizeof(uint32_t));
        for (int k = 0; k < ntp[i]; k++) tp[i][k] = walloc(w);
    }
    uint32_t *oldp = malloc(sizeof(uint32_t) * (size_t)nt), *newp = malloc(sizeof(uint32_t) * (size_t)nt);
    for (int i = 0; i < nt; i++) { oldp[i] = db->t[i]->page; newp[i] = tp[i][0]; }
    remap_catalog(w, oldp, newp, nt);
    for (int i = 0; i < nt; i++) {
        jtable *t = db->t[i];
        for (int k = 0; k < t->nlidx; k++) {
            jlidx *l = &t->lidx[k];
            if (!l->reltable) continue;
            for (int j = 0; j < nt; j++) if (oldp[j] == l->reltable) { l->reltable = newp[j]; break; }
        }
        t->page = tp[i][0];
    }
    /* the content of each table */
    uint32_t *tused = calloc((size_t)nt, sizeof(uint32_t)), *tfree = calloc((size_t)nt, sizeof(uint32_t));
    uint32_t **ixu = calloc((size_t)nt, sizeof(uint32_t *)), **ixf = calloc((size_t)nt, sizeof(uint32_t *));
    uint32_t **lvu = calloc((size_t)nt, sizeof(uint32_t *)), **lvf = calloc((size_t)nt, sizeof(uint32_t *));
    int rc = 0;
    for (int i = 0; i < nt && !rc; i++) {
        jtable *t = db->t[i];
        if (t->load) {                          /* its rows now (the export: one table in memory at a time) */
            rc = t->load(t, t->loadctx, w->err, w->errcap);
            t->load = 0;
            if (rc) break;
        }
        w->nlv = t->nlv;
        w->lv = calloc((size_t)(t->nlv + 1), sizeof(lvcol));
        for (int k = 0; k < t->nlv; k++) { w->lv[k].num = t->lvorder[k]; w->lv[k].cur = -1; }
        bb *enc = calloc((size_t)(t->nrows + 1), sizeof(bb));
        for (int r = 0; r < t->nrows; r++) encode_row(w, t, r, &enc[r]);
        /* the data pages */
        uint32_t *locp = malloc(sizeof(uint32_t) * (size_t)(t->nrows + 1));
        int *locr = malloc(sizeof(int) * (size_t)(t->nrows + 1));
        uint32_t *dpages = malloc(sizeof(uint32_t) * (size_t)(t->nrows + 1));
        int *dfirst = malloc(sizeof(int) * (size_t)(t->nrows + 2)), nd = 0, freeb = 0, cnt = 0;
        for (int r = 0; r < t->nrows; r++) {
            int need = enc[r].n + 2;
            if (!nd || freeb < need || cnt >= 255) {
                dpages[nd] = walloc(w); dfirst[nd] = r; nd++;
                freeb = w->ps - w->dhdr; cnt = 0;
            }
            freeb -= need;
            locp[r] = dpages[nd - 1]; locr[r] = cnt++;
        }
        dfirst[nd] = t->nrows;
        unsigned char owner[4];
        P32(owner, t->page);
        int *dfree = malloc(sizeof(int) * (size_t)(nd + 1));
        for (int d = 0; d < nd; d++) {
            dfree[d] = data_page(w, dpages[d], owner, enc + dfirst[d], dfirst[d + 1] - dfirst[d]);
            spill(w, dpages[d]);
            for (int r = dfirst[d]; r < dfirst[d + 1]; r++) { free(enc[r].p); enc[r].p = 0; }
        }
        /* the indexes */
        ixu[i] = calloc((size_t)(t->nridx + 1), sizeof(uint32_t));
        ixf[i] = calloc((size_t)(t->nridx + 1), sizeof(uint32_t));
        for (int x = 0; x < t->nridx; x++) {
            jridx *rx = &t->ridx[x];
            ent *ents = malloc(sizeof(ent) * (size_t)(t->nrows + 1));
            int ne = 0;
            for (int r = 0; r < t->nrows; r++) {
                int allnull = 1;
                jkey k = { 0 };
                for (int c = 0; c < rx->n; c++) {
                    jetcol *col = 0;
                    for (int q = 0; q < t->ncols; q++) if (t->cols[q].num == rx->col[c]) { col = &t->cols[q]; break; }
                    if (!col) continue;
                    jv *v = jet_val(t, r, col->num);
                    if (col->type == JT_BOOL || !v->null) allnull = 0;
                    if (w->jet4)                /* the keys of the column's sort order (the database's) */
                        jet_sort_lcid(col->type == JT_TEXT && col->misc ? col->misc : (int)(w->db->page0[0x6E] | w->db->page0[0x6F] << 8));
                    jet_col_key(&k, col->type, v, rx->order[c] != 1, w->jet4);
                    jet_sort_lcid(0x409);
                }
                if ((rx->flags & 2) && allnull) { free(k.p); continue; }   /* IgnoreNulls */
                jet_key_truncate(&k, w->jet4);
                unsigned char ptr[4] = { (unsigned char)(locp[r] >> 16), (unsigned char)(locp[r] >> 8), (unsigned char)locp[r], (unsigned char)locr[r] };
                if (k.n + 4 > k.cap) { k.cap = k.n + 4; k.p = realloc(k.p, (size_t)k.cap); }
                memcpy(k.p + k.n, ptr, 4);
                ents[ne].p = k.p; ents[ne].n = k.n + 4; ne++;
            }
            qsort(ents, (size_t)ne, sizeof(ent), ent_cmp);
            uint32_t *ip; int nip;
            uint32_t root = index_pages(w, t, ents, ne, &ip, &nip);
            ixu[i][x] = add_map(w, ip, nip);
            ixf[i][x] = root;
            free(ip);
            for (int e = 0; e < ne; e++) free(ents[e].p);
            free(ents);
        }
        /* the long values */
        lvu[i] = calloc((size_t)(t->nlv + 1), sizeof(uint32_t));
        lvf[i] = calloc((size_t)(t->nlv + 1), sizeof(uint32_t));
        for (int k = 0; k < t->nlv; k++) {
            lvcol *lv = &w->lv[k];
            uint32_t *lp = malloc(sizeof(uint32_t) * (size_t)(lv->npages + 1)), *fp = malloc(sizeof(uint32_t) * (size_t)(lv->npages + 1));
            int nf = 0;
            for (int p = 0; p < lv->npages; p++) {
                lv_done(w, &lv->pages[p]);
                lp[p] = lv->pages[p].page;
                if (lv->pages[p].free >= 64) fp[nf++] = lv->pages[p].page;
            }
            lvu[i][k] = add_map(w, lp, lv->npages);
            lvf[i][k] = add_map(w, fp, nf);
            free(lv->pages); free(lp); free(fp);
        }
        uint32_t *fp = malloc(sizeof(uint32_t) * (size_t)(nd + 1));
        int nf = 0;
        for (int d = 0; d < nd; d++) if (dfree[d] >= 64) fp[nf++] = dpages[d];
        free(dfree);
        tused[i] = add_map(w, dpages, nd);
        tfree[i] = add_map(w, fp, nf);
        free(fp);
        for (int r = 0; r < t->nrows; r++) free(enc[r].p);
        free(enc); free(locp); free(locr); free(dpages); free(dfirst); free(w->lv);
        w->lv = 0;
        /* its values: written (the TDEF needs only the count; the catalog was mapped above) -- memory */
        for (int v = 0; v < t->nrows * t->width; v++) free(t->vals[v].p);
        free(t->vals);
        t->vals = 0; t->caprows = 0;
    }
    for (int i = 0; i < nt && !rc; i++) {
        bb b = { 0 };
        tdef_bytes(w, db->t[i], tused[i], tfree[i], ixu[i], ixf[i], lvu[i], lvf[i], &b);
        tdef_pages(w, &b, tp[i], ntp[i]);
        free(b.p);
    }
    /* the usage maps */
    if (!rc) {
        bb *rows = calloc((size_t)(w->nmaps + 1), sizeof(bb));
        for (int k = 0; k < w->nmaps && !rc; k++) rc = map_row(w, &w->maps[k], &rows[k]);
        for (int m = 0; m < w->nmappages && !rc; m++) {
            int a = m * w->mapcap, n = w->nmaps - a;
            if (n > w->mapcap) n = w->mapcap;
            if (n < 0) n = 0;
            unsigned char owner[4] = { 0 };
            data_page(w, w->mappages[m], owner, rows + a, n);
        }
        for (int k = 0; k < w->nmaps; k++) free(rows[k].p);
        free(rows);
    }
    /* page 0 and the global usage map: 1 = a free page (those past the end of the file) */
    if (!rc) {
        unsigned char *p0 = wpage(w, 0);
        memcpy(p0, db->page0, (size_t)w->ps);
        jet_rc4(HDR_KEY, p0 + 0x18, w->jet4 ? 128 : 126);
        uint32_t n = w->next;
        int nbytes = (int)((n + 7) / 8);
        if (nbytes < w->maplen) nbytes = w->maplen;
        if (5 + nbytes + 2 + (5 + w->maplen) + 2 > w->ps - w->dhdr) {
            /* type 1 (a database larger than one row maps: Jet 3 > 32 MB, Jet 4 > 128 MB): pointers to pages of
             * type 5, each mapping (page size - 4) * 8 pages; the pages past the end of the file in the last one are
             * free, the blocks after it have none (VB6 oracle DEV-14/jetbig) */
            uint32_t per = (uint32_t)(w->ps - 4) * 8;
            int nptr = (w->maplen + 4) / 4, nblk = 0;
            uint32_t ptrs[64] = { 0 };
            while ((uint32_t)nblk * per < w->next) {     /* (a map page is a page too: the end moves) */
                if (nblk >= nptr) { snprintf(w->err, (size_t)w->errcap, "a database larger than the format holds"); rc = -1; break; }
                ptrs[nblk++] = walloc(w);
            }
            if (!rc) {
                uint32_t end = w->next;
                for (int k = 0; k < nblk; k++) {
                    unsigned char *b = wpage(w, ptrs[k]);
                    b[0] = 5; b[1] = 1;
                    for (uint32_t q = 0; q < per; q++)
                        if ((uint32_t)k * per + q >= end) b[4 + q / 8] |= (unsigned char)(1 << (q % 8));
                }
                bb r[2] = { { 0 }, { 0 } };
                b8(&r[0], 1);
                for (int k = 0; k < nptr; k++) b32(&r[0], ptrs[k]);
                b8(&r[1], 0); b32(&r[1], 0); bzero_(&r[1], w->maplen);
                unsigned char owner[4] = { 1, 0, 0, 0 };
                data_page(w, 1, owner, r, 2);
                free(r[0].p); free(r[1].p);
            }
        } else {
            bb r[2] = { { 0 }, { 0 } };
            b8(&r[0], 0); b32(&r[0], 0);
            int at = r[0].n;
            bzero_(&r[0], nbytes);
            for (uint32_t p = n; p < (uint32_t)nbytes * 8; p++) r[0].p[at + p / 8] |= (unsigned char)(1 << (p % 8));
            b8(&r[1], 0); b32(&r[1], 0); bzero_(&r[1], w->maplen);
            unsigned char owner[4] = { 1, 0, 0, 0 };
            data_page(w, 1, owner, r, 2);
            free(r[0].p); free(r[1].p);
        }
    }
    for (int i = 0; i < nt; i++) { free(tp[i]); free(ixu[i]); free(ixf[i]); free(lvu[i]); free(lvf[i]); }
    free(tp); free(ntp); free(oldp); free(newp); free(tused); free(tfree); free(ixu); free(ixf); free(lvu); free(lvf);
    return rc;
}

/* the keys of text: Jet 4, those of the sort order of the column (else of the database: page 0 0x6E), every order
 * measured (jetsort.c, tools/jet/jetloc*.py); its sort order stays, as the key of its SIDs follows from it (changed:
 * 3033). Jet 3: a database of another order (Swedish/Finnish ...) is written as one of General, as CompactDatabase
 * with dbLangGeneral makes it (its texts unchanged; its keys of other orders differ even for plain letters) */
static void general_sort(jdb *db)
{
    unsigned char *p = db->page0;
    int off = db->jet4 ? 0x6E : 0x3A;
    unsigned s = (unsigned)p[off] | (unsigned)p[off + 1] << 8;
    if (s == 0x409 || db->jet4) return;
    /* as Jet's CompactDatabase with dbLangGeneral: sort order 1033 and code page 1252, in the header and in each text
     * column, the bytes of the texts as they were (DEV-14/j025: Greek 1032/1253; with 1253 left, Jet: 3011) */
    p[off] = 0x09; p[off + 1] = 0x04;
    p[0x3C] = 0xE4; p[0x3D] = 0x04;
    for (int i = 0; i < db->nt; i++)
        for (int k = 0; k < db->t[i]->ncols; k++) {
            jetcol *c = &db->t[i]->cols[k];
            if (c->type == JT_TEXT || c->type == JT_MEMO) { c->misc = 0x409; if (c->misc2) c->misc2 = 1252; }
        }
}

int jet_write(jdb *db, const char *path, char *err, int cap)
{
    char dummy[8];
    if (!err) { err = dummy; cap = sizeof dummy; }
    *err = 0;
    wr w = { 0 };
    w.db = db; w.jet4 = db->jet4; w.ps = db->ps;
    w.dhdr = w.jet4 ? 14 : 10;
    w.maplen = w.jet4 ? 64 : 128;
    w.ixmask = w.jet4 ? 0x1B : 0x16;
    w.ixfirst = w.jet4 ? 0x1E0 : 0xF8;
    w.err = err; w.errcap = cap;
    general_sort(db);
    size_t pl = strlen(path);
    char *tmp = malloc(pl + 8);
    snprintf(tmp, pl + 8, "%s.tmp", path);
    /* an encrypted database stays encrypted: the key of page 0 (see jetread.c) */
    w.dbkey = (uint32_t)db->page0[0x3E] | (uint32_t)db->page0[0x3F] << 8 | (uint32_t)db->page0[0x40] << 16 | (uint32_t)db->page0[0x41] << 24;
    w.f = fopen(tmp, "wb");
    int rc = !w.f ? (snprintf(err, (size_t)cap, "cannot write %s", tmp), -1) : db->nt >= 1 ? build(&w) : -1;
    if (w.f) {
        if (!rc) {
            uint32_t total = w.next;
            if (!w.jet4 && total < 20) total = 20;   /* a Jet 3 file shorter than 20 pages is refused */
            unsigned char *zero = calloc(1, (size_t)w.ps);
            for (uint32_t p = 0; p < total; p++) {
                const unsigned char *src = p < w.next ? w.pages[p] : 0;
                if (src == SPILLED) continue;
                put_page(&w, p, src ? src : zero);
            }
            free(zero);
        }
        if (fclose(w.f) || w.ioerr) { if (!rc) snprintf(err, (size_t)cap, "cannot write %s", tmp); rc = -1; }
        if (rc) remove(tmp);
        else {
            remove(path);
            if (rename(tmp, path)) { snprintf(err, (size_t)cap, "cannot write %s", path); rc = -1; }
        }
    }
    free(tmp);
    for (int i = 0; i < w.cappages; i++) if (w.pages[i] != SPILLED) free(w.pages[i]);
    free(w.pages); free(w.mappages);
    for (int i = 0; i < w.nmaps; i++) free(w.maps[i].p);
    free(w.maps);
    return rc;
}
