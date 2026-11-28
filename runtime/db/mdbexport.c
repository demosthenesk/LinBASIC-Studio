/* mdbexport.c - a LinBASIC database (the SQLite copy of an .mdb) written as the Jet file again: what the copy keeps
 * (the tables with their fields and rows, the indexes, the relations, the linked tables, the queries, the
 * properties of MSysLB_Properties) comes from it; everything else - the system tables, the objects of Access
 * (forms, reports, macros, modules), the permissions, the properties the copy does not show - comes unchanged from
 * the file as it is. A database the program made (CreateDatabase) is written as a new file (jetnew.c). The Python
 * prototype: tools/jet/jetexport.py.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "mdbjet.h"
#include <math.h>
#include <time.h>
#include <ctype.h>

#define U16(b, o) ((unsigned)(b)[o] | (unsigned)(b)[(o) + 1] << 8)
#define U32(b, o) ((uint32_t)(b)[o] | (uint32_t)(b)[(o) + 1] << 8 | (uint32_t)(b)[(o) + 2] << 16 | (uint32_t)(b)[(o) + 3] << 24)
#define ID_TABLES 0x0F000001u
#define ID_RELATIONSHIPS 0x0F000003u
enum { REL_UNIQUE = 1, REL_DONTENFORCE = 2, REL_UPDATECASCADE = 0x100, REL_DELETECASCADE = 0x1000 };

static void P16(unsigned char *b, unsigned v) { b[0] = (unsigned char)v; b[1] = (unsigned char)(v >> 8); }
static void P32(unsigned char *b, uint32_t v) { b[0] = (unsigned char)v; b[1] = (unsigned char)(v >> 8); b[2] = (unsigned char)(v >> 16); b[3] = (unsigned char)(v >> 24); }
static char *dupn(const unsigned char *s) { return strdup(s ? (const char *)s : ""); }

/* ---------------------------------------------------------------- the copy */
typedef struct cfield { char *name; int ord, type, size, attrs, required, allowzero, numps; char *defval, *vrule, *vtext; } cfield;
typedef struct cindex { char *name, *fields; int primary, unique, required, ignorenulls, foreign; } cindex;
typedef struct ctable { char *name, *vrule, *vtext; cfield *f; int nf; cindex *x; int nx; int64_t seq; } ctable;
typedef struct clink { char *name, *src, *connect; int attrs; } clink;
typedef struct crel { char *name, *tbl, *ftbl, *fields; int attrs; } crel;
typedef struct cquery { char *name, *sql, *imported; } cquery;
typedef struct copy {
    sqlite3 *s;
    ctable *t; int nt;
    clink *l; int nl;
    crel *r; int nr;
    cquery *q; int nq;
} copy;

static sqlite3_stmt *prep(sqlite3 *s, const char *sql)
{
    sqlite3_stmt *st = 0;
    if (sqlite3_prepare_v2(s, sql, -1, &st, 0) != SQLITE_OK) return 0;
    return st;
}
static int has_table(sqlite3 *s, const char *name)
{
    sqlite3_stmt *st = prep(s, "SELECT 1 FROM main.sqlite_master WHERE type = 'table' AND name = ?1");
    if (!st) return 0;
    sqlite3_bind_text(st, 1, name, -1, SQLITE_STATIC);
    int r = sqlite3_step(st) == SQLITE_ROW;
    sqlite3_finalize(st);
    return r;
}

static void copy_read(copy *c, sqlite3 *s)
{
    memset(c, 0, sizeof *c);
    c->s = s;
    sqlite3_stmt *st = prep(s, "SELECT name, IFNULL(vrule, ''), IFNULL(vtext, ''), attrs, IFNULL(srcname, ''), IFNULL(connect, '') FROM main.MSysLB_Tables ORDER BY rowid");
    while (st && sqlite3_step(st) == SQLITE_ROW) {
        const unsigned char *con = sqlite3_column_text(st, 5);
        if (con && *con) {
            c->l = realloc(c->l, sizeof(clink) * (size_t)(c->nl + 1));
            clink *l = &c->l[c->nl++];
            l->name = dupn(sqlite3_column_text(st, 0)); l->src = dupn(sqlite3_column_text(st, 4)); l->connect = dupn(con);
            l->attrs = sqlite3_column_int(st, 3);
            continue;
        }
        c->t = realloc(c->t, sizeof(ctable) * (size_t)(c->nt + 1));
        ctable *t = &c->t[c->nt++];
        memset(t, 0, sizeof *t);
        t->name = dupn(sqlite3_column_text(st, 0)); t->vrule = dupn(sqlite3_column_text(st, 1)); t->vtext = dupn(sqlite3_column_text(st, 2));
    }
    sqlite3_finalize(st);
    int seq = has_table(s, "sqlite_sequence");
    for (int i = 0; i < c->nt; i++) {
        ctable *t = &c->t[i];
        st = prep(s, "SELECT name, type, size, attrs, required, allowzero, IFNULL(defval, ''), IFNULL(vrule, ''), IFNULL(vtext, ''), ord, IFNULL(misc, 0) FROM main.MSysLB_Fields WHERE tbl = ?1 ORDER BY ord, rowid");
        if (!st)                                /* (a copy made before misc) */
            st = prep(s, "SELECT name, type, size, attrs, required, allowzero, IFNULL(defval, ''), IFNULL(vrule, ''), IFNULL(vtext, ''), ord, 0 FROM main.MSysLB_Fields WHERE tbl = ?1 ORDER BY ord, rowid");
        sqlite3_bind_text(st, 1, t->name, -1, SQLITE_STATIC);
        while (sqlite3_step(st) == SQLITE_ROW) {
            t->f = realloc(t->f, sizeof(cfield) * (size_t)(t->nf + 1));
            cfield *f = &t->f[t->nf++];
            f->name = dupn(sqlite3_column_text(st, 0));
            f->type = sqlite3_column_int(st, 1); f->size = sqlite3_column_int(st, 2); f->attrs = sqlite3_column_int(st, 3);
            f->required = sqlite3_column_int(st, 4); f->allowzero = sqlite3_column_int(st, 5);
            f->defval = dupn(sqlite3_column_text(st, 6)); f->vrule = dupn(sqlite3_column_text(st, 7)); f->vtext = dupn(sqlite3_column_text(st, 8));
            f->ord = sqlite3_column_int(st, 9); f->numps = sqlite3_column_int(st, 10);
        }
        sqlite3_finalize(st);
        st = prep(s, "SELECT name, IFNULL(fields, ''), isprimary, isunique, required, ignorenulls, isforeign FROM main.MSysLB_Indexes WHERE tbl = ?1 ORDER BY ord, rowid");
        sqlite3_bind_text(st, 1, t->name, -1, SQLITE_STATIC);
        while (sqlite3_step(st) == SQLITE_ROW) {
            t->x = realloc(t->x, sizeof(cindex) * (size_t)(t->nx + 1));
            cindex *x = &t->x[t->nx++];
            x->name = dupn(sqlite3_column_text(st, 0)); x->fields = dupn(sqlite3_column_text(st, 1));
            x->primary = sqlite3_column_int(st, 2); x->unique = sqlite3_column_int(st, 3); x->required = sqlite3_column_int(st, 4);
            x->ignorenulls = sqlite3_column_int(st, 5); x->foreign = sqlite3_column_int(st, 6);
        }
        sqlite3_finalize(st);
        if (seq) {
            st = prep(s, "SELECT seq FROM main.sqlite_sequence WHERE name = ?1");
            sqlite3_bind_text(st, 1, t->name, -1, SQLITE_STATIC);
            if (sqlite3_step(st) == SQLITE_ROW) t->seq = sqlite3_column_int64(st, 0);
            sqlite3_finalize(st);
        }
    }
    st = prep(s, "SELECT name, IFNULL(tbl, ''), IFNULL(ftbl, ''), IFNULL(fields, ''), attrs FROM main.MSysLB_Relations ORDER BY rowid");
    while (st && sqlite3_step(st) == SQLITE_ROW) {
        c->r = realloc(c->r, sizeof(crel) * (size_t)(c->nr + 1));
        crel *r = &c->r[c->nr++];
        r->name = dupn(sqlite3_column_text(st, 0)); r->tbl = dupn(sqlite3_column_text(st, 1)); r->ftbl = dupn(sqlite3_column_text(st, 2));
        r->fields = dupn(sqlite3_column_text(st, 3)); r->attrs = sqlite3_column_int(st, 4);
    }
    sqlite3_finalize(st);
    int jq = has_table(s, "MSysLB_JetQueries");
    st = prep(s, jq ? "SELECT q.name, IFNULL(q.sql, ''), j.sql FROM main.MSysLB_Queries q LEFT JOIN main.MSysLB_JetQueries j ON j.name = q.name ORDER BY q.rowid"
                    : "SELECT name, IFNULL(sql, ''), sql FROM main.MSysLB_Queries ORDER BY rowid");   /* (as if none changed) */
    while (st && sqlite3_step(st) == SQLITE_ROW) {
        c->q = realloc(c->q, sizeof(cquery) * (size_t)(c->nq + 1));
        cquery *q = &c->q[c->nq++];
        q->name = dupn(sqlite3_column_text(st, 0)); q->sql = dupn(sqlite3_column_text(st, 1));
        q->imported = sqlite3_column_type(st, 2) == SQLITE_NULL ? 0 : dupn(sqlite3_column_text(st, 2));
    }
    sqlite3_finalize(st);
}

static void copy_free(copy *c)
{
    for (int i = 0; i < c->nt; i++) {
        ctable *t = &c->t[i];
        for (int k = 0; k < t->nf; k++) { free(t->f[k].name); free(t->f[k].defval); free(t->f[k].vrule); free(t->f[k].vtext); }
        for (int k = 0; k < t->nx; k++) { free(t->x[k].name); free(t->x[k].fields); }
        free(t->f); free(t->x); free(t->name); free(t->vrule); free(t->vtext);
    }
    for (int i = 0; i < c->nl; i++) { free(c->l[i].name); free(c->l[i].src); free(c->l[i].connect); }
    for (int i = 0; i < c->nr; i++) { free(c->r[i].name); free(c->r[i].tbl); free(c->r[i].ftbl); free(c->r[i].fields); }
    for (int i = 0; i < c->nq; i++) { free(c->q[i].name); free(c->q[i].sql); free(c->q[i].imported); }
    free(c->t); free(c->l); free(c->r); free(c->q);
}

/* ---------------------------------------------------------------- the model */
typedef struct ex {
    jdb *db;
    copy *cp;
    jtable *cat;
    int cId, cParent, cName, cType, cDateCreate, cDateUpdate, cOwner, cFlags, cDatabase, cConnect, cForeign, cLvProp;
    unsigned char now[8];
    unsigned char owner[64]; int nowner;
    unsigned char creator[2]; int has_creator;
} ex;

static double now_date(void)
{
    time_t t = time(0);
    struct tm tm;
#ifdef _WIN32
    localtime_s(&tm, &t);
#else
    localtime_r(&t, &tm);
#endif
    long y = tm.tm_year + 1900, m = tm.tm_mon + 1, d = tm.tm_mday;
    long a = (14 - m) / 12, yy = y + 4800 - a, mm = m + 12 * a - 3;
    long jdn = d + (153 * mm + 2) / 5 + 365 * yy + yy / 4 - yy / 100 + yy / 400 - 32045;
    return (double)(jdn - 2415019) + (tm.tm_hour * 3600 + tm.tm_min * 60 + tm.tm_sec) / 86400.0;
}

/* a text of the copy (UTF-8) as the file keeps it: Jet 4 UTF-16LE, Jet 3 its code page (malloc, *n bytes) */
static unsigned char *enc_text(jdb *db, const char *s, int maxchars, int *n)
{
    uint16_t *w;
    int k = jet_utf8_to_utf16(s ? s : "", &w);
    if (maxchars > 0 && k > maxchars) k = maxchars;
    unsigned char *b = malloc((size_t)k * 2 + 1);
    if (db->jet4) {
        for (int i = 0; i < k; i++) { b[2 * i] = (unsigned char)w[i]; b[2 * i + 1] = (unsigned char)(w[i] >> 8); }
        *n = 2 * k;
    } else *n = jet_utf16_to_cp(db->codepage, w, k, b);
    free(w);
    return b;
}
static void set_text(jdb *db, jv *v, const char *s)
{
    int n;
    unsigned char *b = enc_text(db, s, 0, &n);
    jet_val_set(v, b, n);
    free(b);
}
static void set_u32(jv *v, uint32_t x) { unsigned char b[4]; P32(b, x); jet_val_set(v, b, 4); }
static void set_i16(jv *v, int x) { unsigned char b[2]; P16(b, (unsigned)x); jet_val_set(v, b, 2); }
static uint32_t get_u32(jv *v) { return v->null || v->n < 4 ? 0 : U32(v->p, 0); }
static int get_i16(jv *v) { return v->null || v->n < 2 ? -32768 : (short)U16(v->p, 0); }

static void row_del(jtable *t, int r)
{
    jv *row = &t->vals[(size_t)r * (size_t)t->width];
    for (int c = 0; c < t->width; c++) free(row[c].p);
    memmove(row, row + t->width, sizeof(jv) * (size_t)(t->nrows - r - 1) * (size_t)t->width);
    t->nrows--;
}
static void rows_clear(jtable *t)
{
    for (int i = 0; i < t->nrows * t->width; i++) free(t->vals[i].p);
    t->nrows = 0;
}
static int col_num(jtable *t, const char *name) { jetcol *c = jet_col(t, name); return c ? c->num : -1; }

static char *name_of(ex *e, int r) { jv *v = jet_val(e->cat, r, e->cName); return v->null ? strdup("") : jet_text_utf8(e->db, v->p, v->n); }
static int obj_row(ex *e, const char *name, int t1, int t2)
{
    for (int r = 0; r < e->cat->nrows; r++) {
        int ty = get_i16(jet_val(e->cat, r, e->cType));
        if (ty != t1 && ty != t2) continue;
        char *n = name_of(e, r);
        int hit = !jet_strcasecmp8(n, name);
        free(n);
        if (hit) return r;
    }
    return -1;
}

static uint32_t new_id(ex *e)
{
    uint32_t hi = 0x7FFFFFFFu;
    for (int r = 0; r < e->cat->nrows; r++) { uint32_t id = get_u32(jet_val(e->cat, r, e->cId)); if (id >= 0x80000000u && id > hi && id < 0xFFFFFFF0u) hi = id; }
    return hi + 1;
}

/* the permissions of a new object: those its container passes on (inheritable), the creator as the owner */
static void add_aces(ex *e, uint32_t oid, uint32_t parent)
{
    jtable *a = jet_table(e->db, "MSysACEs");
    if (!a) return;
    int co = col_num(a, "ObjectId"), cs = col_num(a, "SID"), ci = col_num(a, "FInheritable");
    if (co < 0 || cs < 0 || ci < 0) return;
    int n = a->nrows;
    for (int r = 0; r < n; r++) {
        if (get_u32(jet_val(a, r, co)) != parent || !jet_val(a, r, ci)->n) continue;
        int k = jet_row_add(a);
        for (int c = 0; c < a->width; c++) {
            jv *src = jet_val(a, r, c), *dst = jet_val(a, k, c);
            if (src->p) jet_val_set(dst, src->p, src->n); else { dst->null = src->null; dst->n = src->n; }
        }
        set_u32(jet_val(a, k, co), oid);
        jet_val(a, k, ci)->n = 0;
        jv *sid = jet_val(a, k, cs);
        if (e->has_creator && e->nowner && sid->n == 2 && !memcmp(sid->p, e->creator, 2)) jet_val_set(sid, e->owner, e->nowner);
    }
}

static int new_object(ex *e, const char *name, int type, uint32_t parent, uint32_t oid, int flags)
{
    jtable *t = e->cat;
    int r = jet_row_add(t);
    set_u32(jet_val(t, r, e->cId), oid);
    set_u32(jet_val(t, r, e->cParent), parent);
    set_text(e->db, jet_val(t, r, e->cName), name);
    set_i16(jet_val(t, r, e->cType), type);
    jet_val_set(jet_val(t, r, e->cDateCreate), e->now, 8);
    jet_val_set(jet_val(t, r, e->cDateUpdate), e->now, 8);
    if (e->nowner) jet_val_set(jet_val(t, r, e->cOwner), e->owner, e->nowner);
    set_u32(jet_val(t, r, e->cFlags), (uint32_t)flags);
    add_aces(e, oid, parent);
    return r;
}

static uint32_t drop_object(ex *e, int r)
{
    uint32_t oid = get_u32(jet_val(e->cat, r, e->cId));
    row_del(e->cat, r);
    jtable *a = jet_table(e->db, "MSysACEs");
    int co = a ? col_num(a, "ObjectId") : -1;
    for (int k = 0; co >= 0 && k < a->nrows; ) { if (get_u32(jet_val(a, k, co)) == oid) row_del(a, k); else k++; }
    return oid;
}

static int is_user_table(ex *e, jtable *t)
{
    if (t->type == 0x53 || !strncasecmp(t->name, "MSys", 4)) return 0;
    int r = obj_row(e, t->name, 1, 1);
    if (r < 0) return 0;
    return !(get_u32(jet_val(e->cat, r, e->cFlags)) & 0x80000002u);
}

/* ---------------------------------------------------------------- tables */
static int dao_jet(int dt)
{
    switch (dt) {
    case dbBoolean: return JT_BOOL; case dbByte: return JT_BYTE; case dbInteger: return JT_INT; case dbLong: return JT_LONG;
    case dbCurrency: return JT_MONEY; case dbSingle: return JT_FLOAT; case dbDouble: return JT_DOUBLE; case dbDate: return JT_DATETIME;
    case dbBinary: return JT_BINARY; case dbText: return JT_TEXT; case dbLongBinary: return JT_OLE; case dbMemo: return JT_MEMO;
    case dbGUID: return JT_GUID; case dbDecimal: case dbNumeric: return JT_NUMERIC;
    }
    return JT_MEMO;
}
static int fixlen(int jt) { static const int L[17] = { 0, 1, 1, 2, 4, 8, 4, 8, 8, 0, 0, 0, 0, 0, 0, 16, 17 }; return jt < 17 ? L[jt] : 0; }
static int jet_len(ex *e, int jt, int size)
{
    if (jt == JT_TEXT) return e->db->jet4 ? size * 2 : size;
    if (jt == JT_MEMO || jt == JT_OLE) return 0;
    if (jt == JT_BINARY) return size;
    return jt == JT_BOOL ? 1 : fixlen(jt) ? fixlen(jt) : size;
}

static jetcol *col_by_num(jtable *t, int num) { for (int i = 0; i < t->ncols; i++) if (t->cols[i].num == num) return &t->cols[i]; return 0; }
static jetcol *col_by_name(jtable *t, const char *n) { return jet_col(t, n); }

static int same_columns(ex *e, jtable *t, ctable *ct)
{
    if (t->ncols != ct->nf) return 0;
    for (int k = 0; k < ct->nf; k++) {
        jetcol *c = col_by_num(t, k);
        cfield *f = &ct->f[k];
        if (!c || jet_strcasecmp8(c->name, f->name) || c->type != dao_jet(f->type)) return 0;
        if ((c->type == JT_TEXT || c->type == JT_BINARY) && c->len != jet_len(e, c->type, f->size)) return 0;
        if (!(c->flags & 4) != !(f->attrs & dbAutoIncrField)) return 0;
    }
    return 1;
}

/* new column descriptors; a column the original had (same name and type) keeps its own */
static jtable *build_table(ex *e, ctable *ct, jtable *orig)
{
    jtable *t = jet_table_new(ct->name);
    t->cols = calloc((size_t)(ct->nf ? ct->nf : 1), sizeof(jetcol));
    t->ncols = ct->nf;
    for (int k = 0; k < ct->nf; k++) {
        cfield *f = &ct->f[k];
        int jt = dao_jet(f->type), fixed = fixlen(jt) > 0;
        jetcol *c = &t->cols[k], *old = 0;
        if (orig) for (int i = 0; i < orig->ncols; i++) if (!jet_strcasecmp8(orig->cols[i].name, f->name) && orig->cols[i].type == jt) { old = &orig->cols[i]; break; }
        int flags = (fixed ? 1 : 0) | 2 | (f->attrs & dbAutoIncrField ? 4 : 0) | (jt == JT_MEMO && (f->attrs & dbHyperlinkField) ? 0x80 : 0);
        c->name = strdup(f->name);
        c->type = jt; c->num = k; c->id = f->ord; c->len = jet_len(e, jt, f->size); c->flags = flags;   /* id: OrdinalPosition */
        if (old) {
            memcpy(c->raw, old->raw, sizeof c->raw); c->rawlen = old->rawlen;
            c->misc = old->misc; c->misc2 = old->misc2; c->xflags = old->xflags; c->fileflag = old->fileflag;
            c->flags = (old->flags & ~0x87) | flags;
        } else if (e->db->jet4) {
            int sort = e->db->page0[0x6E] | e->db->page0[0x6F] << 8;     /* a new text column: the database's order */
            c->misc = jt == JT_TEXT ? (sort ? sort : 0x409) : 0;
            if (jt == JT_NUMERIC) c->misc = f->numps ? f->numps : 18 | 0 << 8;   /* DECIMAL(p, s) of ADO; else 18, 0 */
        } else {
            c->misc = 0x409; c->misc2 = 0x4E4; c->fileflag = 1;
        }
    }
    t->width = ct->nf ? ct->nf : 1;
    if (orig) t->page = orig->page;
    return t;
}

/* a Decimal of the copy kept as text, exactly (see jet_numeric_text); 0 if it is not a plain decimal */
static int numeric_from_text(const jetcol *c, const char *s, unsigned char b[17])
{
    int scale = (c->misc >> 8) & 0xFF, neg = 0, frac = -1, nd = 0;
    uint32_t w[4] = { 0, 0, 0, 0 };
    while (*s == ' ') s++;
    if (*s == '-' || *s == '+') neg = *s++ == '-';
    for (; *s; s++) {
        if (*s == '.' && frac < 0) { frac = 0; continue; }
        if (!isdigit((unsigned char)*s)) return 0;
        if (frac >= 0) { if (frac >= scale) continue; frac++; }   /* (digits past the scale: cut) */
        uint64_t carry = (uint64_t)(*s - '0');
        for (int k = 3; k >= 0; k--) { uint64_t x = (uint64_t)w[k] * 10 + carry; w[k] = (uint32_t)x; carry = x >> 32; }
        if (carry) return 0;
        nd++;
    }
    if (!nd) return 0;
    for (int f = frac < 0 ? 0 : frac; f < scale; f++) {
        uint64_t carry = 0;
        for (int k = 3; k >= 0; k--) { uint64_t x = (uint64_t)w[k] * 10 + carry; w[k] = (uint32_t)x; carry = x >> 32; }
        if (carry) return 0;
    }
    memset(b, 0, 17);
    b[0] = neg && (w[0] | w[1] | w[2] | w[3]) ? 0x80 : 0;
    for (int k = 0; k < 4; k++) P32(b + 1 + 4 * k, w[k]);
    return 1;
}

static void put_value(ex *e, jetcol *c, jv *v, sqlite3_stmt *st, int i)
{
    int ty = sqlite3_column_type(st, i);
    unsigned char b[17];
    if (c->type == JT_BOOL) { v->null = 0; v->n = ty != SQLITE_NULL && sqlite3_column_double(st, i) != 0; return; }
    if (ty == SQLITE_NULL) { jet_val_set(v, 0, 0); return; }
    switch (c->type) {
    case JT_BYTE: b[0] = (unsigned char)sqlite3_column_int(st, i); jet_val_set(v, b, 1); return;
    case JT_INT: P16(b, (unsigned)sqlite3_column_int(st, i)); jet_val_set(v, b, 2); return;
    case JT_LONG: P32(b, (uint32_t)sqlite3_column_int64(st, i)); jet_val_set(v, b, 4); return;
    case JT_MONEY: {
        double d = sqlite3_column_double(st, i) * 10000.0;
        int64_t m = d >= 9.2233720368547758e18 ? INT64_MAX : d <= -9.2233720368547758e18 ? INT64_MIN : (int64_t)llround(d);
        P32(b, (uint32_t)m); P32(b + 4, (uint32_t)((uint64_t)m >> 32));
        jet_val_set(v, b, 8);
        return; }
    case JT_FLOAT: { float f = (float)sqlite3_column_double(st, i); memcpy(b, &f, 4); jet_val_set(v, b, 4); return; }
    case JT_DOUBLE: { double d = sqlite3_column_double(st, i); memcpy(b, &d, 8); jet_val_set(v, b, 8); return; }
    case JT_DATETIME: {
        double d = 0;
        if (ty == SQLITE_TEXT) { vb_date x; if (jet_text_date((const char *)sqlite3_column_text(st, i), sqlite3_column_bytes(st, i), &x)) d = x; else d = atof((const char *)sqlite3_column_text(st, i)); }
        else d = sqlite3_column_double(st, i);
        memcpy(b, &d, 8); jet_val_set(v, b, 8);
        return; }
    case JT_GUID: {
        if (ty == SQLITE_BLOB && sqlite3_column_bytes(st, i) == 16) { jet_val_set(v, sqlite3_column_blob(st, i), 16); return; }
        const char *s = (const char *)sqlite3_column_text(st, i);
        if (!strncasecmp(s, "{guid", 5)) s += 5;       /* (as DAO shows it) */
        unsigned char h[16]; int k = 0;
        for (; *s && k < 32; s++) {
            int x = isdigit((unsigned char)*s) ? *s - '0' : (*s >= 'a' && *s <= 'f') ? *s - 'a' + 10 : (*s >= 'A' && *s <= 'F') ? *s - 'A' + 10 : -1;
            if (x < 0) continue;
            if (k & 1) h[k / 2] = (unsigned char)(h[k / 2] << 4 | x); else h[k / 2] = (unsigned char)x;
            k++;
        }
        if (k != 32) { jet_val_set(v, 0, 0); return; }
        unsigned char g[16] = { h[3], h[2], h[1], h[0], h[5], h[4], h[7], h[6], h[8], h[9], h[10], h[11], h[12], h[13], h[14], h[15] };
        jet_val_set(v, g, 16);
        return; }
    case JT_BINARY: case JT_OLE: {
        int n = sqlite3_column_bytes(st, i);
        const void *p = ty == SQLITE_BLOB ? sqlite3_column_blob(st, i) : (const void *)sqlite3_column_text(st, i);
        if (c->type == JT_BINARY && c->len > 0 && n > c->len) n = c->len;
        jet_val_set(v, p ? p : "", n);
        return; }
    case JT_NUMERIC: {
        if ((ty == SQLITE_TEXT || ty == SQLITE_BLOB) && numeric_from_text(c, (const char *)sqlite3_column_text(st, i), b)) { jet_val_set(v, b, 17); return; }
        double d = sqlite3_column_double(st, i);
        int scale = (c->misc >> 8) & 0xFF;
        long double n = roundl(fabsl((long double)d) * powl(10.0L, scale));
        memset(b, 0, sizeof b);
        b[0] = d < 0 ? 0x80 : 0;
        for (int k = 3; k >= 0; k--) {                  /* the words, the least significant last */
            long double q = floorl(n / 4294967296.0L);
            uint32_t w = (uint32_t)(n - q * 4294967296.0L);
            P32(b + 1 + 4 * k, w);
            n = q;
        }
        jet_val_set(v, b, 17);
        return; }
    default: {                                          /* text, memo */
        int n;
        unsigned char *t = enc_text(e->db, (const char *)sqlite3_column_text(st, i), c->type == JT_TEXT ? (e->db->jet4 ? c->len / 2 : c->len) : 0, &n);
        jet_val_set(v, t, n);
        free(t);
        return; }
    }
}

static int load_rows(ex *e, jtable *t, char *err, int cap)
{
    rows_clear(t);
    jbuf q = { 0 };
    jb_s(&q, "SELECT ");
    int nc = 0;
    jetcol **cols = malloc(sizeof(jetcol *) * (size_t)(t->ncols + 1));
    for (int num = 0; num < t->width; num++) {
        jetcol *c = col_by_num(t, num);
        if (!c) continue;
        if (nc) jb_s(&q, ", ");
        jb_qid(&q, c->name, -1);
        cols[nc++] = c;
    }
    if (!nc) jb_s(&q, "1");
    jb_s(&q, " FROM main."); jb_qid(&q, t->name, -1); jb_s(&q, " ORDER BY rowid");
    sqlite3_stmt *st = prep(e->cp->s, q.p);
    jb_free(&q);
    if (!st) { snprintf(err, (size_t)cap, "%s: %s", t->name, sqlite3_errmsg(e->cp->s)); free(cols); return -1; }
    while (sqlite3_step(st) == SQLITE_ROW) {
        int r = jet_row_add(t);
        for (int i = 0; i < nc; i++) put_value(e, cols[i], jet_val(t, r, cols[i]->num), st, i);
    }
    sqlite3_finalize(st);
    jetcol *au = 0;
    for (int i = 0; i < t->ncols; i++) if (t->cols[i].flags & 4) au = &t->cols[i];
    if (au) {
        int64_t hi = (int64_t)t->autonum;
        for (int r = 0; r < t->nrows; r++) { jv *v = jet_val(t, r, au->num); if (!v->null && v->n == 4 && (int32_t)U32(v->p, 0) > hi) hi = (int32_t)U32(v->p, 0); }
        for (int i = 0; i < e->cp->nt; i++) if (!jet_strcasecmp8(e->cp->t[i].name, t->name) && e->cp->t[i].seq > hi) hi = e->cp->t[i].seq;
        t->autonum = (uint32_t)hi;
    }
    free(cols);
    return 0;
}

static int load_lazy(jtable *t, void *ctx, char *err, int cap) { return load_rows((ex *)ctx, t, err, cap); }

/* ---------------------------------------------------------------- indexes */
static int ridx_find(jtable *t, const int *col, const unsigned char *ord, int n, int flags, int any_flags)
{
    for (int k = 0; k < t->nridx; k++) {
        jridx *r = &t->ridx[k];
        if (r->n != n || (!any_flags && r->flags != flags)) continue;
        int same = 1;
        for (int i = 0; i < n; i++) if (r->col[i] != col[i] || r->order[i] != ord[i]) same = 0;
        if (same) return k;
    }
    return -1;
}
static int ridx_add(jtable *t, const int *col, const unsigned char *ord, int n, int flags)
{
    int k = ridx_find(t, col, ord, n, flags, 0);
    if (k >= 0) return k;
    t->ridx = realloc(t->ridx, sizeof(jridx) * (size_t)(t->nridx + 1));
    jridx *r = &t->ridx[t->nridx];
    memset(r, 0, sizeof *r);
    r->n = n; r->flags = flags;
    for (int i = 0; i < n; i++) { r->col[i] = col[i]; r->order[i] = ord[i]; }
    return t->nridx++;
}
static jlidx *lidx_add(jtable *t, const char *name, int real, int type)
{
    t->lidx = realloc(t->lidx, sizeof(jlidx) * (size_t)(t->nlidx + 1));
    jlidx *l = &t->lidx[t->nlidx];
    memset(l, 0, sizeof *l);
    l->name = strdup(name); l->num = t->nlidx; l->real = real; l->type = type; l->relidx = 0xFFFFFFFFu; l->cup = l->cdel = 4;
    t->nlidx++;
    return l;
}

/* the fields of an index of the copy ("+a;-b") as columns; -1 when one is not a column of the table */
static int parse_fields(jtable *t, const char *spec, int *col, unsigned char *ord)
{
    int n = 0;
    const char *p = spec;
    while (*p && n < 10) {
        const char *q = strchr(p, ';');
        int len = q ? (int)(q - p) : (int)strlen(p);
        if (len > 0) {
            int desc = *p == '-';
            const char *nm = (*p == '+' || *p == '-') ? p + 1 : p;
            int nl = len - (int)(nm - p);
            char *name = strndup(nm, (size_t)nl);
            jetcol *c = col_by_name(t, name);
            free(name);
            if (!c) return -1;
            col[n] = c->num; ord[n] = desc ? 0 : 1; n++;
        }
        if (!q) break;
        p = q + 1;
    }
    return n;
}

static int same_indexes(jtable *t, ctable *ct)
{
    int nwant = 0, nhave = 0;
    for (int i = 0; i < ct->nx; i++) if (!ct->x[i].foreign) nwant++;
    for (int i = 0; i < t->nlidx; i++) if (t->lidx[i].type != 2) nhave++;
    if (nwant != nhave) return 0;
    for (int i = 0; i < ct->nx; i++) {
        cindex *x = &ct->x[i];
        if (x->foreign) continue;
        jlidx *l = 0;
        for (int k = 0; k < t->nlidx; k++) if (t->lidx[k].type != 2 && !jet_strcasecmp8(t->lidx[k].name, x->name)) l = &t->lidx[k];
        if (!l || l->real < 0 || l->real >= t->nridx) return 0;
        jridx *r = &t->ridx[l->real];
        int col[10]; unsigned char ord[10];
        int n = parse_fields(t, x->fields, col, ord);
        if (n != r->n) return 0;
        for (int k = 0; k < n; k++) if (col[k] != r->col[k] || ord[k] != r->order[k]) return 0;
        if ((l->type == 1) != !!x->primary || !!(r->flags & 1) != !!(x->unique || x->primary) || !!(r->flags & 2) != !!x->ignorenulls) return 0;
    }
    return 1;
}

static void build_indexes(ex *e, jtable *t, ctable *ct)
{
    for (int i = 0; i < t->nlidx; i++) free(t->lidx[i].name);
    free(t->lidx); free(t->ridx);
    t->lidx = 0; t->ridx = 0; t->nlidx = t->nridx = 0;
    for (int i = 0; i < ct->nx; i++) {
        cindex *x = &ct->x[i];
        if (x->foreign) continue;
        int col[10]; unsigned char ord[10];
        int n = parse_fields(t, x->fields, col, ord);
        if (n <= 0) continue;
        int flags = (x->unique || x->primary ? 1 : 0) | (x->ignorenulls ? 2 : 0) | (x->required || x->primary ? 8 : 0) | (e->db->jet4 ? 0x80 : 0);
        lidx_add(t, x->name, ridx_add(t, col, ord, n, flags), x->primary ? 1 : 0);
    }
}

/* logical indexes numbered in order (the relations' dropped first); a real index no logical one uses is dropped */
static void drop_relation_indexes(jtable *t)
{
    for (int i = 1; i < t->nlidx; i++) {         /* in the order of their numbers (the file keeps them by name) */
        jlidx x = t->lidx[i];
        int j = i;
        while (j > 0 && t->lidx[j - 1].num > x.num) { t->lidx[j] = t->lidx[j - 1]; j--; }
        t->lidx[j] = x;
    }
    int k = 0;
    for (int i = 0; i < t->nlidx; i++) {
        if (t->lidx[i].type == 2) { free(t->lidx[i].name); continue; }
        t->lidx[k++] = t->lidx[i];
    }
    t->nlidx = k;
    int *map = malloc(sizeof(int) * (size_t)(t->nridx + 1)), nr = 0;
    for (int r = 0; r < t->nridx; r++) {
        int used = 0;
        for (int i = 0; i < t->nlidx; i++) if (t->lidx[i].real == r) used = 1;
        map[r] = used ? nr : -1;
        if (used) t->ridx[nr++] = t->ridx[r];
    }
    t->nridx = nr;
    for (int i = 0; i < t->nlidx; i++) { t->lidx[i].num = i; t->lidx[i].real = map[t->lidx[i].real]; }
    free(map);
}

/* ---------------------------------------------------------------- properties */
static void block_drop(jprops *p, int b)
{
    jpblock *bl = &p->blocks[b];
    for (int i = 0; i < bl->nprops; i++) { free(bl->props[i].name); free(bl->props[i].v); }
    free(bl->props); free(bl->name);
    memmove(bl, bl + 1, sizeof(jpblock) * (size_t)(p->nblocks - b - 1));
    p->nblocks--;
}
static void set_text_prop(jprops *p, jpblock *b, const char *name, const char *val, int dtype, int flags)
{
    jprop *old = jprops_get(b, name);
    if (old && (old->dtype == 10 || old->dtype == 12)) {
        char *t = jprops_text(p, old);
        int same = !strcmp(t, val);
        free(t);
        if (same) return;                       /* the bytes as they are */
    }
    int dt = old ? old->dtype : dtype;
    jprops_set_text(p, b, name, val, flags);
    jprops_get(b, name)->dtype = dt == 12 ? 12 : 10;
}
static void set_bool_prop(jprops *p, jpblock *b, const char *name, int v, int flags)
{
    jprop *old = jprops_get(b, name);
    if (old && old->dtype == 1 && old->n >= 1 && !old->v[0] == !v) return;
    jprops_set_bool(p, b, name, v, flags);
}
/* DefaultValue, ValidationRule, ValidationText: as Jet writes them (memo / text, flag 1; VB6 oracle d_jet) */
static void set_rule_prop(jprops *p, jpblock *b, const char *name, const char *val, int dtype)
{
    if (val && *val) set_text_prop(p, b, name, val, dtype, 1); else jprops_remove(b, name);
}

/* a value of MSysLB_Properties (text) in the type of the property */
static void set_typed_prop(jprops *p, jpblock *b, const char *name, int type, const char *text)
{
    jprop *old = jprops_get(b, name);
    if (old && old->dtype == type) {
        char *t = mdb_prop_value_text(p, old);
        int same = t && text && !strcmp(t, text);
        free(t);
        if (same) return;
    }
    unsigned char v[8];
    int flags = old ? old->flags : 0;
    if (!text) text = "";
    switch (type) {
    case 1: jprops_set_bool(p, b, name, atol(text) != 0, flags); return;
    case 2: v[0] = (unsigned char)atoi(text); jprops_set(p, b, name, 2, v, 1, flags); return;
    case 3: P16(v, (unsigned)atoi(text)); jprops_set(p, b, name, 3, v, 2, flags); return;
    case 4: P32(v, (uint32_t)atol(text)); jprops_set(p, b, name, 4, v, 4, flags); return;
    case 5: { int64_t m = llround(atof(text) * 10000.0); P32(v, (uint32_t)m); P32(v + 4, (uint32_t)((uint64_t)m >> 32)); jprops_set(p, b, name, 5, v, 8, flags); return; }
    case 6: { float f = (float)atof(text); memcpy(v, &f, 4); jprops_set(p, b, name, 6, v, 4, flags); return; }
    case 7: case 8: { double d = atof(text); memcpy(v, &d, 8); jprops_set(p, b, name, type, v, 8, flags); return; }
    default: jprops_set_text(p, b, name, text, flags); jprops_get(b, name)->dtype = type == 12 ? 12 : 10; return;
    }
}

/* the properties of the program (MSysLB_Properties) of one block: those it shows are what the copy has */
static void user_props(ex *e, jprops *p, jpblock *b, char kind, const char *obj, const char *sub)
{
    if (!has_table(e->cp->s, "MSysLB_Properties")) return;
    sqlite3_stmt *st = prep(e->cp->s, "SELECT name, type, value FROM main.MSysLB_Properties WHERE kind = ?1 AND obj = ?2 AND sub = ?3 ORDER BY rowid");
    if (!st) return;
    char ks[2] = { kind, 0 };
    sqlite3_bind_text(st, 1, ks, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 2, obj, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 3, sub, -1, SQLITE_STATIC);
    char **names = 0; int nn = 0;
    while (sqlite3_step(st) == SQLITE_ROW) {
        const char *nm = (const char *)sqlite3_column_text(st, 0);
        if (!nm || mdb_builtin_prop(kind, nm)) continue;
        int ty = sqlite3_column_int(st, 1);
        set_typed_prop(p, b, nm, ty, sqlite3_column_type(st, 2) == SQLITE_NULL ? 0 : (const char *)sqlite3_column_text(st, 2));
        names = realloc(names, sizeof(char *) * (size_t)(nn + 1));
        names[nn++] = strdup(nm);
    }
    sqlite3_finalize(st);
    for (int i = 0; i < b->nprops; ) {            /* one the program deleted */
        jprop *pr = &b->props[i];
        int keep = mdb_builtin_prop(kind, pr->name) || !mdb_prop_shown(pr->dtype);
        for (int k = 0; !keep && k < nn; k++) if (!jet_strcasecmp8(names[k], pr->name)) keep = 1;
        if (keep) i++; else jprops_remove(b, pr->name);
    }
    for (int k = 0; k < nn; k++) free(names[k]);
    free(names);
}

/* a new block: a column's after the last column's (the order of the blocks as the file has it stays: Access writes the
 * object's first, DAO last; PSC file of DEV-14/q1), the others at the end */
static jpblock *new_block(jprops *p, int kind, const char *name)
{
    jpblock *b = jprops_block(p, kind, name, 1);         /* (appended) */
    if (kind != 1) return b;
    int last = -1, at = (int)(b - p->blocks);
    for (int i = 0; i < at; i++) if (p->blocks[i].kind == 1) last = i;
    if (last < 0 || last == at - 1) return b;
    jpblock x = *b;
    memmove(&p->blocks[last + 2], &p->blocks[last + 1], sizeof(jpblock) * (size_t)(at - last - 1));
    p->blocks[last + 1] = x;
    return &p->blocks[last + 1];
}

/* the LvProp of a table: the original blocks, with what the copy keeps set again */
static unsigned char *table_props(ex *e, jtable *t, ctable *ct, jv *old, int *n)
{
    jprops *p = jprops_parse(old && !old->null ? old->p : 0, old && !old->null ? old->n : 0, e->db->jet4, e->db->codepage);
    for (int b = 0; b < p->nblocks; ) {
        jpblock *bl = &p->blocks[b];
        if (bl->kind == 1 && !col_by_name(t, bl->name)) block_drop(p, b); else b++;
    }
    for (int b = 0; b < p->nblocks; ) {         /* an index that is no more */
        jpblock *bl = &p->blocks[b];
        int keep = bl->kind != 2;
        for (int i = 0; !keep && i < t->nlidx; i++) if (!jet_strcasecmp8(t->lidx[i].name, bl->name)) keep = 1;
        if (keep) b++; else block_drop(p, b);
    }
    for (int k = 0; k < ct->nf; k++) {
        cfield *f = &ct->f[k];
        jetcol *c = col_by_name(t, f->name);
        if (!c) continue;
        jpblock *b = jprops_block(p, 1, c->name, 0);
        int had = b != 0;
        if (!b) b = new_block(p, 1, c->name);
        if (strcmp(b->name, c->name)) { free(b->name); b->name = strdup(c->name); }
        /* a field made by DAO has AllowZeroLength (text, memo) and Required (not a counter); one made by SQL DDL
         * (allowzero 2) none of them (VB6 oracle d_jet) */
        int ddl = f->allowzero == 2;
        if ((c->type == JT_TEXT || c->type == JT_MEMO) && (!ddl || jprops_get(b, "AllowZeroLength")))
            set_bool_prop(p, b, "AllowZeroLength", f->allowzero == 1, 1);
        if (f->required || (!ddl && !(c->flags & 4)) || jprops_get(b, "Required")) set_bool_prop(p, b, "Required", f->required, 1);
        set_rule_prop(p, b, "DefaultValue", f->defval, 12);
        set_rule_prop(p, b, "ValidationRule", f->vrule, 12);
        set_rule_prop(p, b, "ValidationText", f->vtext, 10);
        user_props(e, p, b, 'F', t->name, c->name);
        if (!had && !b->nprops) block_drop(p, (int)(b - p->blocks));
    }
    for (int i = 0; i < t->nlidx; i++) {
        if (t->lidx[i].type == 2) continue;
        jpblock *b = jprops_block(p, 2, t->lidx[i].name, 0);
        int had = b != 0;
        if (!b) b = jprops_block(p, 2, t->lidx[i].name, 1);
        user_props(e, p, b, 'I', t->name, t->lidx[i].name);
        if (!had && !b->nprops) block_drop(p, (int)(b - p->blocks));
    }
    jpblock *b0 = jprops_block(p, 0, "", 0);
    int had0 = b0 != 0;
    if (!b0) b0 = jprops_block(p, 0, "", 1);
    set_rule_prop(p, b0, "ValidationRule", ct->vrule, 10);
    set_rule_prop(p, b0, "ValidationText", ct->vtext, 10);
    user_props(e, p, b0, 'T', t->name, "");
    if (!b0->nprops && !had0) block_drop(p, (int)(b0 - p->blocks));
    unsigned char *out = p->nblocks ? jprops_build(p, n) : 0;
    if (!p->nblocks) *n = 0;
    jprops_free(p);
    return out;
}

/* ---------------------------------------------------------------- relations */
static char *hidden_name(jtable *t)
{
    char nm[16];
    for (char ch = 'A'; ch <= 'Z'; ch++) {
        snprintf(nm, sizeof nm, ".r%c", ch);
        int used = 0;
        for (int i = 0; i < t->nlidx; i++) if (!strcmp(t->lidx[i].name, nm)) used = 1;
        if (!used) return strdup(nm);
    }
    for (int k = 0;; k++) {
        snprintf(nm, sizeof nm, ".r%d", k);
        int used = 0;
        for (int i = 0; i < t->nlidx; i++) if (!strcmp(t->lidx[i].name, nm)) used = 1;
        if (!used) return strdup(nm);
    }
}

static void relations(ex *e, jtable **tabs, int nt)
{
    jtable *rel = jet_table(e->db, "MSysRelationships");
    if (!rel) return;
    int cc = col_num(rel, "ccolumn"), cg = col_num(rel, "grbit"), ci = col_num(rel, "icolumn"), cs = col_num(rel, "szColumn");
    int co = col_num(rel, "szObject"), crc = col_num(rel, "szReferencedColumn"), cro = col_num(rel, "szReferencedObject"), cr = col_num(rel, "szRelationship");
    if (cc < 0 || cg < 0 || ci < 0 || cs < 0 || co < 0 || crc < 0 || cro < 0 || cr < 0) return;
    /* the rows of relations of system tables (Access's MSysAccessStorageSTGREL) stay; the others come from the copy */
    char **sysrel = 0; int nsys = 0;
    for (int r = 0; r < rel->nrows; ) {
        jv *a = jet_val(rel, r, co), *b = jet_val(rel, r, cro);
        char *ta = a->null ? strdup("") : jet_text_utf8(e->db, a->p, a->n), *tb = b->null ? strdup("") : jet_text_utf8(e->db, b->p, b->n);
        int sys = !strncasecmp(ta, "MSys", 4) || !strncasecmp(tb, "MSys", 4);
        free(ta); free(tb);
        if (!sys) { row_del(rel, r); continue; }
        jv *nm = jet_val(rel, r, cr);
        sysrel = realloc(sysrel, sizeof(char *) * (size_t)(nsys + 1));
        sysrel[nsys++] = nm->null ? strdup("") : jet_text_utf8(e->db, nm->p, nm->n);
        r++;
    }
    /* the objects of relations the copy no longer has */
    for (int r = 0; r < e->cat->nrows; ) {
        if (get_i16(jet_val(e->cat, r, e->cType)) != 8) { r++; continue; }
        char *n = name_of(e, r);
        int keep = 0;
        for (int k = 0; k < e->cp->nr; k++) if (!jet_strcasecmp8(e->cp->r[k].name, n)) keep = 1;
        for (int k = 0; k < nsys; k++) if (!jet_strcasecmp8(sysrel[k], n)) keep = 1;
        free(n);
        if (keep) r++; else drop_object(e, r);
    }
    for (int k = 0; k < nsys; k++) free(sysrel[k]);
    free(sysrel);
    for (int k = 0; k < e->cp->nr; k++) {
        crel *R = &e->cp->r[k];
        char *pc[10], *fc[10];
        int np = 0;
        for (const char *p = R->fields; *p && np < 10; ) {
            const char *q = strchr(p, ';');
            int len = q ? (int)(q - p) : (int)strlen(p);
            const char *eq = memchr(p, '=', (size_t)len);
            if (eq) { pc[np] = strndup(p, (size_t)(eq - p)); fc[np] = strndup(eq + 1, (size_t)(len - (eq - p) - 1)); np++; }
            if (!q) break;
            p = q + 1;
        }
        for (int i = 0; i < np; i++) {
            int r = jet_row_add(rel);
            set_u32(jet_val(rel, r, cc), (uint32_t)np);
            set_u32(jet_val(rel, r, cg), (uint32_t)R->attrs);
            set_u32(jet_val(rel, r, ci), (uint32_t)i);
            set_text(e->db, jet_val(rel, r, cs), fc[i]);
            set_text(e->db, jet_val(rel, r, co), R->ftbl);
            set_text(e->db, jet_val(rel, r, crc), pc[i]);
            set_text(e->db, jet_val(rel, r, cro), R->tbl);
            set_text(e->db, jet_val(rel, r, cr), R->name);
        }
        if (obj_row(e, R->name, 8, 8) < 0) new_object(e, R->name, 8, ID_RELATIONSHIPS, new_id(e), 0);
        jtable *P = 0, *F = 0;
        for (int i = 0; i < nt; i++) { if (!jet_strcasecmp8(tabs[i]->name, R->tbl)) P = tabs[i]; if (!jet_strcasecmp8(tabs[i]->name, R->ftbl)) F = tabs[i]; }
        int ok = np > 0 && P && F && !(R->attrs & REL_DONTENFORCE);
        int pcol[10], fcol[10]; unsigned char asc[10];
        for (int i = 0; ok && i < np; i++) {
            jetcol *a = col_by_name(P, pc[i]), *b = col_by_name(F, fc[i]);
            if (!a || !b) { ok = 0; break; }
            pcol[i] = a->num; fcol[i] = b->num; asc[i] = 1;
        }
        if (ok) {
            int preal = ridx_find(P, pcol, asc, np, 0, 1);
            if (preal < 0) preal = ridx_add(P, pcol, asc, np, e->db->jet4 ? 0x80 : 0);
            /* the foreign side: DAO shows the flags of its real index, so one that is unique only for a one-to-one
             * relation, never required (Jet does not share the primary key's: o030) */
            int want = R->attrs & REL_UNIQUE ? 1 : 0, freal = -1;
            for (int k = 0; k < F->nridx && freal < 0; k++)
                if ((F->ridx[k].flags & 0x0B) == want && ridx_find(F, fcol, asc, np, F->ridx[k].flags, 0) == k) freal = k;
            if (freal < 0) freal = ridx_add(F, fcol, asc, np, want | (e->db->jet4 ? 0x80 : 0));
            int cup = R->attrs & REL_UPDATECASCADE ? 1 : 0, cdel = R->attrs & REL_DELETECASCADE ? 1 : 0;
            char *hn = hidden_name(P);
            jlidx *lp = lidx_add(P, hn, preal, 2);
            free(hn);
            lp->reltype = 1; lp->relidx = 0; lp->cup = cup; lp->cdel = cdel;
            int lpnum = lp->num;
            jlidx *lf = lidx_add(F, R->name, freal, 2);
            lf->reltype = 2; lf->relidx = (uint32_t)lpnum; lf->cup = cup; lf->cdel = cdel; lf->rel_to = P;
            P->lidx[lpnum].relidx = (uint32_t)lf->num;
            P->lidx[lpnum].rel_to = F;
        }
        for (int i = 0; i < np; i++) { free(pc[i]); free(fc[i]); }
    }
}

/* ---------------------------------------------------------------- linked tables */
static void link_parts(const char *con, char **dbpath, char **rest)
{
    jbuf r = { 0 };
    *dbpath = strdup("");
    const char *p = con;
    while (*p) {
        const char *q = strchr(p, ';');
        int len = q ? (int)(q - p) : (int)strlen(p);
        if (len >= 9 && !strncasecmp(p, "DATABASE=", 9)) { free(*dbpath); *dbpath = strndup(p + 9, (size_t)(len - 9)); }
        else if (len > 0) { jb_s(&r, ";"); jb_put(&r, p, len); }      /* (";PWD=x"; none: Null, as Jet keeps it) */
        if (!q) break;
        p = q + 1;
    }
    *rest = r.p ? r.p : strdup("");
}

static void links(ex *e)
{
    for (int r = 0; r < e->cat->nrows; ) {      /* the links the copy no longer has */
        int ty = get_i16(jet_val(e->cat, r, e->cType));
        if (ty != 4 && ty != 6) { r++; continue; }
        char *n = name_of(e, r);
        int keep = 0;
        for (int k = 0; k < e->cp->nl; k++) if (!jet_strcasecmp8(e->cp->l[k].name, n)) keep = 1;
        free(n);
        if (keep) r++; else drop_object(e, r);
    }
    for (int k = 0; k < e->cp->nl; k++) {
        clink *L = &e->cp->l[k];
        int odbc = !strncasecmp(L->connect, "ODBC", 4);
        int r = obj_row(e, L->name, 4, 6);
        if (r < 0) r = new_object(e, L->name, odbc ? 4 : 6, ID_TABLES, new_id(e), (odbc ? 0 : 0x2000) | (L->attrs & 0x30000));   /* (0x2000: as Jet makes a link, d_jetlink) */
        else {
            /* as the import read it: unchanged, the row as it is */
            jv *dv = e->cDatabase >= 0 ? jet_val(e->cat, r, e->cDatabase) : 0, *cv = e->cConnect >= 0 ? jet_val(e->cat, r, e->cConnect) : 0;
            jv *fv = e->cForeign >= 0 ? jet_val(e->cat, r, e->cForeign) : 0;
            char *dbs = dv && !dv->null ? jet_text_utf8(e->db, dv->p, dv->n) : strdup(""), *con = cv && !cv->null ? jet_text_utf8(e->db, cv->p, cv->n) : strdup("");
            char *src = fv && !fv->null ? jet_text_utf8(e->db, fv->p, fv->n) : strdup("");
            jbuf was = { 0 };
            if (get_i16(jet_val(e->cat, r, e->cType)) == 6) { jb_s(&was, con); if (!strstr(con, "DATABASE=")) { jb_s(&was, ";DATABASE="); jb_s(&was, dbs); } }
            else jb_s(&was, con);
            int same = !strcmp(was.p ? was.p : "", L->connect) && !jet_strcasecmp8(src, L->src);
            jb_free(&was); free(dbs); free(con); free(src);
            if (same) continue;
            set_i16(jet_val(e->cat, r, e->cType), odbc ? 4 : 6);
            jet_val_set(jet_val(e->cat, r, e->cDateUpdate), e->now, 8);
        }
        if (e->cForeign >= 0) set_text(e->db, jet_val(e->cat, r, e->cForeign), L->src);
        if (odbc) {
            if (e->cConnect >= 0) set_text(e->db, jet_val(e->cat, r, e->cConnect), L->connect);
            if (e->cDatabase >= 0) jet_val_set(jet_val(e->cat, r, e->cDatabase), 0, 0);
        } else {
            char *dbp, *rest;
            link_parts(L->connect, &dbp, &rest);
            if (e->cDatabase >= 0) set_text(e->db, jet_val(e->cat, r, e->cDatabase), dbp);
            if (e->cConnect >= 0) { if (*rest) set_text(e->db, jet_val(e->cat, r, e->cConnect), rest); else jet_val_set(jet_val(e->cat, r, e->cConnect), 0, 0); }
            free(dbp); free(rest);
        }
    }
}

/* ---------------------------------------------------------------- queries */
static void query_props(ex *e, int r, const char *name, int fresh)
{
    jv *old = jet_val(e->cat, r, e->cLvProp);
    jprops *p = jprops_parse(old->null ? 0 : old->p, old->null ? 0 : old->n, e->db->jet4, e->db->codepage);
    jpblock *b = jprops_block(p, 0, "", 0);
    int had = b != 0;
    if (!b) b = jprops_block(p, 0, "", 1);
    if (fresh) {
        unsigned char v[4];
        P16(v, 60); jprops_set(p, b, "ODBCTimeout", 3, v, 2, 1);
        P32(v, 0); jprops_set(p, b, "MaxRecords", 4, v, 4, 1);
    }
    user_props(e, p, b, 'Q', name, "");
    if (!had && !b->nprops) block_drop(p, (int)(b - p->blocks));
    if (!p->nblocks) { if (!old->null) jet_val_set(old, 0, 0); jprops_free(p); return; }
    int n;
    unsigned char *out = jprops_build(p, &n);
    if (old->null || n != old->n || memcmp(out, old->p, (size_t)n)) jet_val_set(old, out, n);
    free(out);
    jprops_free(p);
}

static void queries(ex *e, char *err, int cap)
{
    jtable *q = jet_table(e->db, "MSysQueries");
    if (!q) return;
    int qo = col_num(q, "ObjectId"), qa = col_num(q, "Attribute"), qr = col_num(q, "Order"), qf = col_num(q, "Flag");
    int q1 = col_num(q, "Name1"), q2 = col_num(q, "Name2"), qe = col_num(q, "Expression");
    for (int r = 0; r < e->cat->nrows; ) {      /* the queries the copy no longer has */
        if ((get_i16(jet_val(e->cat, r, e->cType)) & 0x7F) != 5) { r++; continue; }
        char *n = name_of(e, r);
        int keep = n[0] == '~';
        for (int k = 0; !keep && k < e->cp->nq; k++) if (!jet_strcasecmp8(e->cp->q[k].name, n)) keep = 1;
        free(n);
        if (keep) { r++; continue; }
        uint32_t oid = drop_object(e, r);
        for (int k = 0; k < q->nrows; ) { if (get_u32(jet_val(q, k, qo)) == oid) row_del(q, k); else k++; }
    }
    for (int k = 0; k < e->cp->nq; k++) {
        cquery *Q = &e->cp->q[k];
        int r = obj_row(e, Q->name, 5, 5);
        if (r < 0) for (int i = 0; i < e->cat->nrows && r < 0; i++) {
            if ((get_i16(jet_val(e->cat, i, e->cType)) & 0x7F) != 5) continue;
            char *n = name_of(e, i);
            if (!jet_strcasecmp8(n, Q->name)) r = i;
            free(n);
        }
        int changed = r < 0 || !Q->imported || strcmp(Q->imported, Q->sql);
        if (!changed) { query_props(e, r, Q->name, 0); continue; }
        int n;
        char qerr[300];
        jqrow *rows = jq_compile(Q->sql, &n, qerr, sizeof qerr);
        if (!rows) {                            /* kept as it is in the file (a new one left out); the caller sees why */
            snprintf(err, (size_t)cap, "query %s: %s", Q->name, qerr);
            continue;
        }
        int fresh = r < 0;
        if (fresh) r = new_object(e, Q->name, 5, ID_TABLES, new_id(e), jq_flags(rows, n));
        else {
            set_u32(jet_val(e->cat, r, e->cFlags), (uint32_t)jq_flags(rows, n));
            jet_val_set(jet_val(e->cat, r, e->cDateUpdate), e->now, 8);
            uint32_t oid = get_u32(jet_val(e->cat, r, e->cId));
            for (int i = 0; i < q->nrows; ) { if (get_u32(jet_val(q, i, qo)) == oid) row_del(q, i); else i++; }
        }
        uint32_t oid = get_u32(jet_val(e->cat, r, e->cId));
        for (int i = 0; i < n; i++) {
            int x = jet_row_add(q);
            unsigned char b[4];
            set_u32(jet_val(q, x, qo), oid);
            b[0] = (unsigned char)rows[i].attr; jet_val_set(jet_val(q, x, qa), b, 1);
            b[0] = (unsigned char)(rows[i].order >> 24); b[1] = (unsigned char)(rows[i].order >> 16);
            b[2] = (unsigned char)(rows[i].order >> 8); b[3] = (unsigned char)rows[i].order;
            jet_val_set(jet_val(q, x, qr), b, 4);
            if (rows[i].flag != JQ_NOFLAG) set_i16(jet_val(q, x, qf), rows[i].flag);
            if (rows[i].name1) set_text(e->db, jet_val(q, x, q1), rows[i].name1);
            if (rows[i].name2) set_text(e->db, jet_val(q, x, q2), rows[i].name2);
            if (rows[i].expr) set_text(e->db, jet_val(q, x, qe), rows[i].expr);
        }
        jq_free(rows, n);
        query_props(e, r, Q->name, fresh);
    }
}

/* ---------------------------------------------------------------- all */
static int file_exists(const char *p) { FILE *f = fopen(p, "rb"); if (f) fclose(f); return f != 0; }

static int version_of(sqlite3 *s)
{
    int v = 4;
    sqlite3_stmt *st = prep(s, "SELECT value FROM main.MSysLB_Props WHERE name = 'Version'");
    if (st && sqlite3_step(st) == SQLITE_ROW) { const char *t = (const char *)sqlite3_column_text(st, 0); if (t && t[0] == '3') v = 3; }
    sqlite3_finalize(st);
    return v;
}

/* the sort order a new database is made in (CreateDatabase's ;LANGID=), 0 = General */
static int lcid_of(sqlite3 *s)
{
    int v = 0;
    sqlite3_stmt *st = prep(s, "SELECT value FROM main.MSysLB_Props WHERE name = 'SortLCID'");
    if (st && sqlite3_step(st) == SQLITE_ROW) v = sqlite3_column_int(st, 0);
    sqlite3_finalize(st);
    return v;
}

int mdb_export(sqlite3 *s, const char *path, char *err, int cap)
{
    if (cap) *err = 0;
    jdb *db = file_exists(path) ? jet_read_opt(path, mdb_codepage(), err, cap, JET_NO_USER_ROWS) : jet_new_order(version_of(s) == 4, mdb_codepage(), lcid_of(s));
    if (!db) { if (!*err) snprintf(err, (size_t)cap, "not a Jet database"); return -1; }
    copy cp;
    copy_read(&cp, s);
    ex e;
    memset(&e, 0, sizeof e);
    e.db = db; e.cp = &cp; e.cat = db->t[0];
    e.cId = col_num(e.cat, "Id"); e.cParent = col_num(e.cat, "ParentId"); e.cName = col_num(e.cat, "Name");
    e.cType = col_num(e.cat, "Type"); e.cDateCreate = col_num(e.cat, "DateCreate"); e.cDateUpdate = col_num(e.cat, "DateUpdate");
    e.cOwner = col_num(e.cat, "Owner"); e.cFlags = col_num(e.cat, "Flags"); e.cDatabase = col_num(e.cat, "Database");
    e.cConnect = col_num(e.cat, "Connect"); e.cForeign = col_num(e.cat, "ForeignName"); e.cLvProp = col_num(e.cat, "LvProp");
    double now = now_date();
    memcpy(e.now, &now, 8);
    /* the owner of the user's objects: that of a table or query of the user, else of MSysDb; the key of the SIDs of
     * Jet 4 from the owner of MSysObjects (always the system's, 02 03) */
    for (int pass = 0; pass < 2 && !e.nowner; pass++)
        for (int r = 0; r < e.cat->nrows && !e.nowner; r++) {
            int ty = get_i16(jet_val(e.cat, r, e.cType));
            char *n = name_of(&e, r);
            int pick = pass == 0 ? (ty == 1 || ty == 5) && strncasecmp(n, "MSys", 4) : !strcmp(n, "MSysDb");
            free(n);
            jv *o = jet_val(e.cat, r, e.cOwner);
            if (pick && !o->null && o->n && o->n <= (int)sizeof e.owner) { memcpy(e.owner, o->p, (size_t)o->n); e.nowner = o->n; }
        }
    int sysrow = obj_row(&e, "MSysObjects", 1, 1);
    if (sysrow >= 0) {
        jv *o = jet_val(e.cat, sysrow, e.cOwner);
        if (!o->null && o->n == 2) {
            unsigned char sysid[2], cre[2];
            jet_sid(0, 0x0302, sysid); jet_sid(0, 0x0402, cre);
            e.creator[0] = (unsigned char)(cre[0] ^ sysid[0] ^ o->p[0]);
            e.creator[1] = (unsigned char)(cre[1] ^ sysid[1] ^ o->p[1]);
            e.has_creator = 1;
        }
    }
    int rc = 0;
    /* the tables of the user in the file, the system's ones kept as they are */
    jtable **user = calloc((size_t)db->nt + 1, sizeof(jtable *)), **sys = calloc((size_t)db->nt + 1, sizeof(jtable *));
    int nu = 0, ns = 0;
    for (int i = 4; i < db->nt; i++) { if (is_user_table(&e, db->t[i])) user[nu++] = db->t[i]; else sys[ns++] = db->t[i]; }
    for (int i = 0; i < nu; i++) {              /* the dropped ones */
        int keep = 0;
        for (int k = 0; k < cp.nt; k++) if (!jet_strcasecmp8(cp.t[k].name, user[i]->name)) keep = 1;
        if (keep) continue;
        int r = obj_row(&e, user[i]->name, 1, 1);
        if (r >= 0) drop_object(&e, r);
        jet_table_free(user[i]);
        user[i] = 0;
    }
    jtable **out = calloc((size_t)cp.nt + 1, sizeof(jtable *));
    int *catrow_id = calloc((size_t)cp.nt + 1, sizeof(int));
    for (int k = 0; k < cp.nt && !rc; k++) {
        ctable *ct = &cp.t[k];
        jtable *orig = 0;
        for (int i = 0; i < nu; i++) if (user[i] && !jet_strcasecmp8(user[i]->name, ct->name)) { orig = user[i]; user[i] = 0; break; }
        jtable *t;
        int keep_ix;
        if (orig && same_columns(&e, orig, ct)) {
            t = orig;
            keep_ix = same_indexes(t, ct);
            for (int i = 0; i < ct->nf; i++) col_by_num(t, i)->id = ct->f[i].ord;   /* OrdinalPosition */
        } else {
            t = build_table(&e, ct, orig);
            keep_ix = 0;
            if (orig) jet_table_free(orig);
            orig = 0;
        }
        free(t->name); t->name = strdup(ct->name);
        rows_clear(t);
        t->load = load_lazy; t->loadctx = &e;   /* its rows: when the writer writes it (memory) */
        if (!keep_ix) build_indexes(&e, t, ct);
        drop_relation_indexes(t);
        out[k] = t;
        int r = obj_row(&e, ct->name, 1, 1);
        /* the TDEF page of a new table is known only when the file is laid out: a key the writer maps (the catalog
         * and the permissions) */
        if (!t->page) t->page = 0x7FFF0000u + (uint32_t)k;
        if (r < 0) r = new_object(&e, ct->name, 1, ID_TABLES, t->page, 0);
        else {
            char *n = name_of(&e, r);
            if (strcmp(n, ct->name)) set_text(db, jet_val(e.cat, r, e.cName), ct->name);
            free(n);
        }
        if (!orig || !keep_ix) jet_val_set(jet_val(e.cat, r, e.cDateUpdate), e.now, 8);
    }
    for (int i = 0; i < nu; i++) if (user[i]) jet_table_free(user[i]);   /* (not in the copy: already dropped) */
    if (!rc) {
        relations(&e, out, cp.nt);
        /* LvProp after the relations (the logical indexes are known), then the list of tables */
        for (int k = 0; k < cp.nt; k++) {
            int r = obj_row(&e, out[k]->name, 1, 1);
            if (r < 0) continue;
            int n;
            jv *lp = jet_val(e.cat, r, e.cLvProp);
            unsigned char *pb = table_props(&e, out[k], &cp.t[k], lp, &n);
            if (!pb) { if (!lp->null) jet_val_set(lp, 0, 0); }
            else if (lp->null || n != lp->n || memcmp(pb, lp->p, (size_t)n)) jet_val_set(lp, pb, n);
            free(pb);
            catrow_id[k] = r;
        }
        jtable **all = malloc(sizeof(jtable *) * (size_t)(4 + ns + cp.nt + 1));
        int na = 0;
        for (int i = 0; i < 4; i++) all[na++] = db->t[i];
        for (int i = 0; i < ns; i++) all[na++] = sys[i];
        for (int k = 0; k < cp.nt; k++) all[na++] = out[k];
        free(db->t);
        db->t = all; db->nt = na;
        for (int k = 0; k < cp.nt; k++)
            for (int i = 0; i < out[k]->nlidx; i++) if (out[k]->lidx[i].rel_to) out[k]->lidx[i].reltable = out[k]->lidx[i].rel_to->page;
        links(&e);
        queries(&e, err, cap);
    } else {
        for (int k = 0; k < cp.nt; k++) if (out[k]) jet_table_free(out[k]);
        db->nt = 4;                             /* (the others freed above) */
        for (int i = 0; i < ns; i++) jet_table_free(sys[i]);
    }
    /* a password the program set (NewPassword) */
    sqlite3_stmt *pw = prep(s, "SELECT value FROM main.MSysLB_Props WHERE name = 'Password'");
    if (pw && sqlite3_step(pw) == SQLITE_ROW) jet_password_set(db->page0, db->jet4, db->codepage, (const char *)sqlite3_column_text(pw, 0));
    sqlite3_finalize(pw);
    char qwarn[400];
    snprintf(qwarn, sizeof qwarn, "%s", rc ? "" : err);
    if (!rc) rc = jet_write(db, path, err, cap);
    if (!rc && cap) snprintf(err, (size_t)cap, "%s", qwarn);   /* written: a query that could not be compiled */
    if (!rc) {
        /* the queries as they are now in the file */
        sqlite3_exec(s, "CREATE TABLE IF NOT EXISTS main.MSysLB_JetQueries (name TEXT PRIMARY KEY COLLATE NOCASE, sql TEXT);"
                        "DELETE FROM main.MSysLB_JetQueries;"
                        "INSERT INTO main.MSysLB_JetQueries (name, sql) SELECT name, sql FROM main.MSysLB_Queries", 0, 0, 0);
    }
    free(user); free(sys); free(out); free(catrow_id);
    copy_free(&cp);
    jet_free(db);
    return rc;
}
