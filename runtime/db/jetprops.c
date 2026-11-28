/* jetprops.c - the properties Jet keeps in MSysObjects.LvProp: of the object itself (a table, a query, the
 * database), of its columns and of its indexes (Description, Caption, DefaultValue, ValidationRule, Format ...).
 * 'MR2\0' (Jet 4) / 'KKD\0' (Jet 3), then chunks: the names (type 0x80) and the values (0 the object, 1 a column,
 * 2 an index). Read and written back byte for byte (tools/jet/jetprops.py).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "jetprops.h"
#include <stdlib.h>
#include <string.h>

#define U16(b, o) ((unsigned)(b)[o] | (unsigned)(b)[(o) + 1] << 8)
#define U32(b, o) ((uint32_t)(b)[o] | (uint32_t)(b)[(o) + 1] << 8 | (uint32_t)(b)[(o) + 2] << 16 | (uint32_t)(b)[(o) + 3] << 24)

/* a text ending in a NUL (Jet keeps one after a field's ValidationRule set by DAO when the field is in the table:
 * VB6 oracle d_jetrule, DAO shows it) is said with JET_NUL at its end */
static char *nul_end(char *s, int nul)
{
    if (!nul || !s) return s;
    size_t l = strlen(s);
    s = realloc(s, l + sizeof JET_NUL);
    memcpy(s + l, JET_NUL, sizeof JET_NUL);
    return s;
}
static char *text_of(const jprops *p, const unsigned char *b, int n)
{
    if (p->jet4) {
        int m = n / 2, nul = m > 0 && U16(b, 2 * (m - 1)) == 0;
        uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(m + 1));
        for (int i = 0; i < m; i++) w[i] = (uint16_t)U16(b, 2 * i);
        char *s = jet_utf16_to_utf8(w, nul ? m - 1 : m);
        free(w);
        return nul_end(s, nul);
    }
    int nul = n > 0 && b[n - 1] == 0;
    if (nul) n--;
    uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(n + 1));
    int k = jet_cp_to_utf16(p->codepage, b, n, w);
    char *s = jet_utf16_to_utf8(w, k);
    free(w);
    return nul_end(s, nul);
}
/* the bytes of a name / a text value */
static unsigned char *bytes_of(const jprops *p, const char *s, int *n)
{
    uint16_t *w;
    size_t l = strlen(s), ml = sizeof JET_NUL - 1;
    int nul = l >= ml && !memcmp(s + l - ml, JET_NUL, ml);
    char *c = nul ? strndup(s, l - ml) : 0;
    int k = jet_utf8_to_utf16(c ? c : s, &w);
    free(c);
    unsigned char *b;
    if (p->jet4) {
        b = malloc((size_t)k * 2 + 3);
        for (int i = 0; i < k; i++) { b[2 * i] = (unsigned char)w[i]; b[2 * i + 1] = (unsigned char)(w[i] >> 8); }
        *n = k * 2;
        if (nul) { b[*n] = 0; b[*n + 1] = 0; *n += 2; }
    } else {
        b = malloc((size_t)k * 2 + 2);
        *n = jet_utf16_to_cp(p->codepage, w, k, b);
        if (nul) b[(*n)++] = 0;
    }
    free(w);
    return b;
}

jprops *jprops_parse(const unsigned char *b, int n, int jet4, int codepage)
{
    jprops *p = calloc(1, sizeof *p);
    p->jet4 = jet4;
    p->codepage = codepage;
    if (!b || n < 4 || (memcmp(b, "MR2", 4) && memcmp(b, "KKD", 4))) return p;
    int o = 4;
    while (o + 6 <= n) {
        uint32_t ln = U32(b, o);
        int typ = (int)U16(b, o + 4);
        if (ln < 6 || o + (int)ln > n) break;
        const unsigned char *c = b + o + 6;
        int cl = (int)ln - 6;
        if (typ == 0x80) {
            int q = 0;
            while (q + 2 <= cl) {
                int nl = (int)U16(c, q);
                if (q + 2 + nl > cl) break;
                char *s = text_of(p, c + q + 2, nl);
                p->names = realloc(p->names, sizeof(char *) * (size_t)(p->nnames + 1));
                p->names[p->nnames++] = s;
                q += 2 + nl;
            }
        } else if (cl >= 6) {
            int hl = (int)U32(c, 0), nl = (int)U16(c, 4);
            jpblock *bl = jprops_block(p, typ, 0, 1);
            free(bl->name);
            bl->name = text_of(p, c + 6, nl < cl - 6 ? nl : cl - 6);
            int q = hl;
            while (q + 8 <= cl) {
                int el = (int)U16(c, q), vl = (int)U16(c, q + 6), ni = (int)U16(c, q + 4);
                if (el < 8 || q + el > cl) break;
                jprop *pr = jprops_add(bl);
                pr->flags = c[q + 2];
                pr->dtype = c[q + 3];
                pr->name = strdup(ni < p->nnames ? p->names[ni] : "");
                if (vl > el - 8) vl = el - 8;
                pr->n = vl;
                pr->v = malloc((size_t)vl + 1);
                memcpy(pr->v, c + q + 8, (size_t)vl);
                q += el;
            }
        }
        o += (int)ln;
    }
    return p;
}

jpblock *jprops_block(jprops *p, int kind, const char *name, int make)
{
    if (name)
        for (int i = 0; i < p->nblocks; i++)
            if (p->blocks[i].kind == kind && !jet_strcasecmp8(p->blocks[i].name, name)) return &p->blocks[i];
    if (!make) return 0;
    p->blocks = realloc(p->blocks, sizeof(jpblock) * (size_t)(p->nblocks + 1));
    jpblock *b = &p->blocks[p->nblocks++];
    memset(b, 0, sizeof *b);
    b->kind = kind;
    b->name = strdup(name ? name : "");
    return b;
}
jprop *jprops_add(jpblock *b)
{
    b->props = realloc(b->props, sizeof(jprop) * (size_t)(b->nprops + 1));
    jprop *pr = &b->props[b->nprops++];
    memset(pr, 0, sizeof *pr);
    return pr;
}
jprop *jprops_get(jpblock *b, const char *name)
{
    for (int i = 0; i < b->nprops; i++) if (!jet_strcasecmp8(b->props[i].name, name)) return &b->props[i];
    return 0;
}
void jprops_remove(jpblock *b, const char *name)
{
    for (int i = 0; i < b->nprops; i++)
        if (!jet_strcasecmp8(b->props[i].name, name)) {
            free(b->props[i].name); free(b->props[i].v);
            memmove(&b->props[i], &b->props[i + 1], sizeof(jprop) * (size_t)(b->nprops - i - 1));
            b->nprops--;
            return;
        }
}
void jprops_set(jprops *p, jpblock *b, const char *name, int dtype, const void *v, int n, int flags)
{
    jprop *pr = jprops_get(b, name);
    if (!pr) { pr = jprops_add(b); pr->name = strdup(name); pr->flags = flags; }
    pr->dtype = dtype;
    free(pr->v);
    pr->v = malloc((size_t)n + 1);
    memcpy(pr->v, v, (size_t)n);
    pr->n = n;
    (void)p;
}
void jprops_set_text(jprops *p, jpblock *b, const char *name, const char *s, int flags)
{
    int n;
    unsigned char *v = bytes_of(p, s, &n);
    jprops_set(p, b, name, 10, v, n, flags);
    free(v);
}
void jprops_set_bool(jprops *p, jpblock *b, const char *name, int v, int flags)
{
    unsigned char c = v ? 0xFF : 0;
    jprops_set(p, b, name, 1, &c, 1, flags);
}
char *jprops_text(jprops *p, jprop *pr) { return text_of(p, pr->v, pr->n); }

static void put(unsigned char **o, int *n, int *cap, const void *s, int m)
{
    if (*n + m > *cap) { *cap = (*n + m) * 2 + 64; *o = realloc(*o, (size_t)*cap); }
    memcpy(*o + *n, s, (size_t)m);
    *n += m;
}
static void put16(unsigned char **o, int *n, int *cap, unsigned v) { unsigned char c[2] = { (unsigned char)v, (unsigned char)(v >> 8) }; put(o, n, cap, c, 2); }
static void put32(unsigned char **o, int *n, int *cap, uint32_t v) { unsigned char c[4] = { (unsigned char)v, (unsigned char)(v >> 8), (unsigned char)(v >> 16), (unsigned char)(v >> 24) }; put(o, n, cap, c, 4); }

/* the bytes: the name table (the names read, then new ones in order of use; names no longer used drop out) */
unsigned char *jprops_build(jprops *p, int *outn)
{
    char **names = malloc(sizeof(char *) * (size_t)(p->nnames + 1));
    int nn = 0, capn = p->nnames + 1;
    for (int i = 0; i < p->nnames; i++) names[nn++] = p->names[i];
    for (int b = 0; b < p->nblocks; b++)
        for (int i = 0; i < p->blocks[b].nprops; i++) {
            const char *s = p->blocks[b].props[i].name;
            int seen = 0;
            for (int k = 0; k < nn; k++) if (!strcmp(names[k], s)) { seen = 1; break; }
            if (!seen) {
                if (nn == capn) { capn *= 2; names = realloc(names, sizeof(char *) * (size_t)capn); }
                names[nn++] = (char *)s;
            }
        }
    int nu = 0;
    for (int k = 0; k < nn; k++) {
        int used = 0;
        for (int b = 0; b < p->nblocks && !used; b++)
            for (int i = 0; i < p->blocks[b].nprops; i++) if (!strcmp(p->blocks[b].props[i].name, names[k])) { used = 1; break; }
        if (used) names[nu++] = names[k];
    }
    nn = nu;
    unsigned char *o = 0, *nb = 0;
    int n = 0, cap = 0, nbn = 0, nbcap = 0;
    put(&o, &n, &cap, p->jet4 ? "MR2\0" : "KKD\0", 4);
    for (int k = 0; k < nn; k++) {
        int l;
        unsigned char *e = bytes_of(p, names[k], &l);
        put16(&nb, &nbn, &nbcap, (unsigned)l);
        put(&nb, &nbn, &nbcap, e, l);
        free(e);
    }
    put32(&o, &n, &cap, (uint32_t)(6 + nbn));
    put16(&o, &n, &cap, 0x80);
    if (nbn) put(&o, &n, &cap, nb, nbn);
    free(nb);
    for (int b = 0; b < p->nblocks; b++) {
        jpblock *bl = &p->blocks[b];
        if (!bl->nprops) continue;
        unsigned char *body = 0;
        int bn = 0, bcap = 0, l;
        unsigned char *e = bytes_of(p, bl->name, &l);
        put32(&body, &bn, &bcap, (uint32_t)(6 + l));
        put16(&body, &bn, &bcap, (unsigned)l);
        put(&body, &bn, &bcap, e, l);
        free(e);
        for (int i = 0; i < bl->nprops; i++) {
            jprop *pr = &bl->props[i];
            int ix = 0;
            for (int k = 0; k < nn; k++) if (!strcmp(names[k], pr->name)) { ix = k; break; }
            put16(&body, &bn, &bcap, (unsigned)(8 + pr->n));
            unsigned char fl[2] = { (unsigned char)pr->flags, (unsigned char)pr->dtype };
            put(&body, &bn, &bcap, fl, 2);
            put16(&body, &bn, &bcap, (unsigned)ix);
            put16(&body, &bn, &bcap, (unsigned)pr->n);
            if (pr->n) put(&body, &bn, &bcap, pr->v, pr->n);
        }
        put32(&o, &n, &cap, (uint32_t)(6 + bn));
        put16(&o, &n, &cap, (unsigned)bl->kind);
        put(&o, &n, &cap, body, bn);
        free(body);
    }
    free(names);
    *outn = n;
    return o;
}

void jprops_free(jprops *p)
{
    if (!p) return;
    for (int i = 0; i < p->nnames; i++) free(p->names[i]);
    for (int b = 0; b < p->nblocks; b++) {
        for (int i = 0; i < p->blocks[b].nprops; i++) { free(p->blocks[b].props[i].name); free(p->blocks[b].props[i].v); }
        free(p->blocks[b].props); free(p->blocks[b].name);
    }
    free(p->names); free(p->blocks); free(p);
}
