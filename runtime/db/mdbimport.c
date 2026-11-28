/* mdbimport.c - a Jet .mdb file (Access 97 / 2000, Jet 3 and 4) copied into a LinBASIC (SQLite) database: the
 * tables with their fields, sizes and properties, the rows, the indexes, the relations (MSysRelationships), the
 * linked tables, the saved queries (MSysQueries, written as the SQL DAO shows) and the properties Jet keeps in
 * LvProp (Description, Caption, Format ... -> MSysLB_Properties). The file is read whole by jetread.c; the program
 * works on the copy and mdbexport.c writes the .mdb again from it.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "mdbjet.h"
#include <ctype.h>
#include <math.h>
#include <sys/stat.h>

#define U16(b, o) ((unsigned)(b)[o] | (unsigned)(b)[(o) + 1] << 8)
#define U32(b, o) ((uint32_t)(b)[o] | (uint32_t)(b)[(o) + 1] << 8 | (uint32_t)(b)[(o) + 2] << 16 | (uint32_t)(b)[(o) + 3] << 24)

/* 3 or 4 for a Jet 3 / Jet 4 file, 0 otherwise */
int mdb_is_jet(const char *path)
{
    unsigned char h[0x20];
    FILE *f = fopen(path, "rb");
    if (!f) return 0;
    size_t n = fread(h, 1, sizeof h, f);
    fclose(f);
    if (n < sizeof h) return 0;
    if (memcmp(h + 4, "Standard Jet DB", 15) && memcmp(h + 4, "Standard ACE DB", 15)) return 0;
    return h[0x14] == 0 ? 3 : 4;
}

/* the code page of the texts of a Jet 3 file: the system's (the program's, as Jet writes them) */
int mdb_codepage(void) { return vb_codepage > 0 ? vb_codepage : 1252; }

int jet_dao_type(int jt)
{
    switch (jt) {
    case JT_BOOL: return dbBoolean; case JT_BYTE: return dbByte; case JT_INT: return dbInteger;
    case JT_LONG: return dbLong; case JT_MONEY: return dbCurrency; case JT_FLOAT: return dbSingle;
    case JT_DOUBLE: return dbDouble; case JT_DATETIME: return dbDate; case JT_BINARY: return dbBinary;
    case JT_TEXT: return dbText; case JT_OLE: return dbLongBinary; case JT_MEMO: return dbMemo;
    case JT_GUID: return dbGUID; case JT_NUMERIC: return dbDecimal;
    }
    return dbLongBinary;
}

static int exec(sqlite3 *db, const char *sql, char *err, int cap)
{
    char *em = 0;
    if (sqlite3_exec(db, sql, 0, 0, &em) != SQLITE_OK) {
        snprintf(err, (size_t)cap, "%s", em ? em : "sqlite");
        sqlite3_free(em);
        return -1;
    }
    return 0;
}

/* ---------------------------------------------------------------- values */
/* a text of the file (Jet 4: UTF-16LE; Jet 3: the code page) as UTF-8 (malloc) */
char *jet_text_utf8(jdb *db, const unsigned char *p, int n)
{
    if (db->jet4) {
        uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(n / 2 + 1));
        for (int i = 0; i < n / 2; i++) w[i] = (uint16_t)U16(p, 2 * i);
        char *s = jet_utf16_to_utf8(w, n / 2);
        free(w);
        return s;
    }
    uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(n + 1));
    int k = jet_cp_to_utf16(db->codepage, p, n, w);
    char *s = jet_utf16_to_utf8(w, k);
    free(w);
    return s;
}
static char *val_text(jdb *db, jv *v) { return v->null ? 0 : jet_text_utf8(db, v->p, v->n); }
static int val_int(jv *v, int n) { if (v->null || v->n < n) return 0; return n == 1 ? v->p[0] : n == 2 ? (short)U16(v->p, 0) : (int)U32(v->p, 0); }

void jet_guid_text(const unsigned char *g, char out[40])
{
    snprintf(out, 40, "{%02X%02X%02X%02X-%02X%02X-%02X%02X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
             g[3], g[2], g[1], g[0], g[5], g[4], g[7], g[6], g[8], g[9], g[10], g[11], g[12], g[13], g[14], g[15]);
}
/* Decimal: the sign (0x80) and the unscaled value in four 32 bit words, most significant first */
double jet_numeric_value(const jetcol *c, const unsigned char *b)
{
    double v = 0;
    for (int k = 0; k < 4; k++) v = v * 4294967296.0 + (double)U32(b, 1 + 4 * k);
    int scale = (c->misc >> 8) & 0xFF;
    v /= pow(10.0, scale);
    return (b[0] & 0x80) ? -v : v;
}

/* Decimal exactly, as text ("-1234.5678"): the unscaled value of four 32 bit words (the most significant first, each
 * little-endian) divided down by 10 (VB6 oracle DEV-14/jetdec). SQLite keeps it as a number when that is exact */
void jet_numeric_text(const jetcol *c, const unsigned char *b, char out[48])
{
    uint32_t w[4];
    for (int k = 0; k < 4; k++) w[k] = U32(b, 1 + 4 * k);
    char dig[48];
    int n = 0;
    do {
        uint64_t r = 0;
        for (int k = 0; k < 4; k++) { uint64_t x = r << 32 | w[k]; w[k] = (uint32_t)(x / 10); r = x % 10; }
        dig[n++] = (char)('0' + r);
    } while ((w[0] | w[1] | w[2] | w[3]) && n < 45);
    int scale = (c->misc >> 8) & 0xFF;
    while (n <= scale) dig[n++] = '0';
    int o = 0;
    if ((b[0] & 0x80) && !(n == 1 && dig[0] == '0')) out[o++] = '-';
    for (int i = n - 1; i >= 0; i--) {
        out[o++] = dig[i];
        if (i == scale && scale > 0) out[o++] = '.';
    }
    out[o] = 0;
    if (scale > 0) {                            /* trailing zeros of the fraction dropped */
        while (o > 0 && out[o - 1] == '0') out[--o] = 0;
        if (o > 0 && out[o - 1] == '.') out[--o] = 0;
    }
}

static void bind_value(sqlite3_stmt *st, int k, jdb *db, jetcol *c, jv *v, int required)
{
    if (c->type == JT_BOOL) { sqlite3_bind_int(st, k, v->n ? -1 : 0); return; }
    if (v->null) {
        if (required && (c->type == JT_TEXT || c->type == JT_MEMO)) sqlite3_bind_text(st, k, "", 0, SQLITE_STATIC);
        else sqlite3_bind_null(st, k);
        return;
    }
    const unsigned char *p = v->p;
    switch (c->type) {
    case JT_BYTE: sqlite3_bind_int(st, k, p[0]); return;
    case JT_INT: sqlite3_bind_int(st, k, (short)U16(p, 0)); return;
    case JT_LONG: sqlite3_bind_int(st, k, (int)U32(p, 0)); return;
    case JT_MONEY: {
        int64_t m = (int64_t)((uint64_t)U32(p, 0) | (uint64_t)U32(p, 4) << 32);
        sqlite3_bind_double(st, k, (double)m / 10000.0);
        return; }
    case JT_FLOAT: { float f; memcpy(&f, p, 4); sqlite3_bind_double(st, k, (double)f); return; }
    case JT_DOUBLE: { double d; memcpy(&d, p, 8); sqlite3_bind_double(st, k, d); return; }
    case JT_DATETIME: {
        double d; memcpy(&d, p, 8);
        char tx[24];
        jet_date_text(d, tx);
        sqlite3_bind_text(st, k, tx, -1, SQLITE_TRANSIENT);
        return; }
    case JT_GUID: { char g[40]; jet_guid_text(p, g); sqlite3_bind_text(st, k, g, -1, SQLITE_TRANSIENT); return; }
    case JT_NUMERIC: {
        char t[48] = "0";
        if (v->n >= 17) jet_numeric_text(c, p, t);
        /* up to 15 digits a number of SQLite (exact); more: the text as a blob, which SQLite does not round */
        int nd = 0;
        for (const char *q = t; *q; q++) if (isdigit((unsigned char)*q)) nd++;
        for (const char *q = t; *q == '-' || *q == '0' || *q == '.'; q++) if (*q == '0') nd--;
        if (nd <= 15) sqlite3_bind_text(st, k, t, -1, SQLITE_TRANSIENT);
        else sqlite3_bind_blob(st, k, t, (int)strlen(t), SQLITE_TRANSIENT);
        return; }
    case JT_BINARY: case JT_OLE: sqlite3_bind_blob(st, k, p, v->n, SQLITE_TRANSIENT); return;
    default: {
        char *s = jet_text_utf8(db, p, v->n);
        sqlite3_bind_text(st, k, s, -1, free);
        return; }
    }
}

/* ---------------------------------------------------------------- the catalog */
typedef struct cat {
    jdb *db;
    jtable *t;
    int cid, cparent, cname, ctype, cflags, cprop, cdb, cconn, cforeign;
} cat;

static jv *cv(cat *c, int r, int col) { return jet_val(c->t, r, col); }
static int cat_row(cat *c, const char *name, int type)
{
    for (int r = 0; r < c->t->nrows; r++) {
        if (val_int(cv(c, r, c->ctype), 2) != type) continue;
        char *s = val_text(c->db, cv(c, r, c->cname));
        int hit = s && !jet_strcasecmp8(s, name);
        free(s);
        if (hit) return r;
    }
    return -1;
}

/* the properties a program sees as its own (CreateProperty): in MSysLB_Properties */
int mdb_builtin_prop(char kind, const char *name)
{
    static const char *const F[] = { "DefaultValue", "ValidationRule", "ValidationText", "Required", "AllowZeroLength", 0 };
    static const char *const T[] = { "ValidationRule", "ValidationText", 0 };
    static const char *const Q[] = { "ODBCTimeout", "MaxRecords", "ReturnsRecords", "Connect", 0 };
    const char *const *l = kind == 'F' ? F : kind == 'T' ? T : kind == 'Q' ? Q : 0;
    for (int i = 0; l && l[i]; i++) if (!strcasecmp(l[i], name)) return 1;
    return 0;
}
int mdb_prop_shown(int dtype) { return dtype != 0 && dtype != 9 && dtype != 11 && dtype != 15; }
char *mdb_prop_value_text(jprops *p, jprop *pr)
{
    char b[64];
    const unsigned char *v = pr->v;
    switch (pr->dtype) {
    case 1: return strdup(pr->n && v[0] ? "-1" : "0");
    case 2: snprintf(b, sizeof b, "%d", pr->n ? v[0] : 0); return strdup(b);
    case 3: snprintf(b, sizeof b, "%d", pr->n >= 2 ? (short)U16(v, 0) : 0); return strdup(b);
    case 4: snprintf(b, sizeof b, "%ld", pr->n >= 4 ? (long)(int32_t)U32(v, 0) : 0L); return strdup(b);
    case 5: { int64_t m = pr->n >= 8 ? (int64_t)((uint64_t)U32(v, 0) | (uint64_t)U32(v, 4) << 32) : 0; snprintf(b, sizeof b, "%.17g", (double)m / 10000.0); return strdup(b); }
    case 6: { float f = 0; if (pr->n >= 4) memcpy(&f, v, 4); snprintf(b, sizeof b, "%.17g", (double)f); return strdup(b); }
    case 7: case 8: { double d = 0; if (pr->n >= 8) memcpy(&d, v, 8); snprintf(b, sizeof b, "%.17g", d); return strdup(b); }
    case 10: case 12: return jprops_text(p, pr);
    }
    return 0;                                   /* binary values: kept in the file, not shown */
}
static void save_props(sqlite3 *s, jprops *p, char kind, const char *obj)
{
    sqlite3_exec(s, "CREATE TABLE IF NOT EXISTS MSysLB_Properties (kind TEXT, obj TEXT COLLATE NOCASE, "
                 "sub TEXT COLLATE NOCASE, name TEXT COLLATE NOCASE, type INT, value TEXT)", 0, 0, 0);
    sqlite3_stmt *st = 0;
    sqlite3_prepare_v2(s, "INSERT INTO MSysLB_Properties (kind, obj, sub, name, type, value) VALUES (?1, ?2, ?3, ?4, ?5, ?6)", -1, &st, 0);
    for (int b = 0; b < p->nblocks; b++) {
        jpblock *bl = &p->blocks[b];
        if (bl->kind && kind != 'T') continue;  /* the columns of a query, of the database: kept in the file */
        char k = bl->kind == 1 ? 'F' : bl->kind == 2 ? 'I' : kind;
        for (int i = 0; i < bl->nprops; i++) {
            jprop *pr = &bl->props[i];
            if (mdb_builtin_prop(k, pr->name) || !mdb_prop_shown(pr->dtype)) continue;
            char *txt = mdb_prop_value_text(p, pr), ks[2] = { k, 0 };
            sqlite3_bind_text(st, 1, ks, -1, SQLITE_TRANSIENT);
            sqlite3_bind_text(st, 2, obj, -1, SQLITE_TRANSIENT);
            sqlite3_bind_text(st, 3, bl->kind ? bl->name : "", -1, SQLITE_TRANSIENT);
            sqlite3_bind_text(st, 4, pr->name, -1, SQLITE_TRANSIENT);
            sqlite3_bind_int(st, 5, pr->dtype);
            if (txt) sqlite3_bind_text(st, 6, txt, -1, free); else sqlite3_bind_null(st, 6);
            sqlite3_step(st);
            sqlite3_reset(st);
        }
    }
    sqlite3_finalize(st);
}
static int prop_bool(jpblock *b, const char *name)
{
    jprop *p = b ? jprops_get(b, name) : 0;
    return p && p->n && p->v[0];
}
static char *prop_str(jprops *ps, jpblock *b, const char *name)
{
    jprop *p = b ? jprops_get(b, name) : 0;
    return p && (p->dtype == 10 || p->dtype == 12) ? jprops_text(ps, p) : strdup("");
}

/* ---------------------------------------------------------------- tables */
static int is_system(cat *c, int r, jtable *t)
{
    if (t && t->type == 0x53) return 1;
    uint32_t fl = (uint32_t)val_int(cv(c, r, c->cflags), 4);
    if (fl & 0x80000002u) return 1;
    char *s = val_text(c->db, cv(c, r, c->cname));
    int sys = s && !strncasecmp(s, "MSys", 4);
    free(s);
    return sys;
}

static int import_schema(cat *c, jtable *t, jprops *props, sqlite3 *s, char *err, int cap)
{
    jdb *db = c->db;
    sqlite3_stmt *st;
    jpblock *b0 = jprops_block(props, 0, "", 0);
    char *vr = prop_str(props, b0, "ValidationRule"), *vt = prop_str(props, b0, "ValidationText");
    sqlite3_prepare_v2(s, "INSERT INTO MSysLB_Tables (name, attrs, seq, vrule, vtext) VALUES (?1, 0, 0, ?2, ?3)", -1, &st, 0);
    sqlite3_bind_text(st, 1, t->name, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 2, vr, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 3, vt, -1, SQLITE_STATIC);
    sqlite3_step(st);
    sqlite3_finalize(st);
    free(vr); free(vt);
    sqlite3_prepare_v2(s, "INSERT INTO MSysLB_Fields (tbl, name, ord, type, size, attrs, required, allowzero, defval, vrule, vtext, misc) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12)", -1, &st, 0);
    /* the fields in the order of their numbers; OrdinalPosition: the id of the column (VB6 oracle d_dao3, NWIND) */
    for (int num = 0; num < t->width; num++)
        for (int i = 0; i < t->ncols; i++) {
            jetcol *k = &t->cols[i];
            if (k->num != num) continue;
            int type = jet_dao_type(k->type), size, attrs;
            switch (type) {
            case dbText: size = db->jet4 ? k->len / 2 : k->len; attrs = dbVariableField; break;
            case dbMemo: case dbLongBinary: size = 0; attrs = dbVariableField; break;
            case dbBinary: size = k->len; attrs = dbVariableField; break;
            case dbBoolean: case dbByte: size = 1; attrs = dbFixedField; break;
            case dbInteger: size = 2; attrs = dbFixedField; break;
            case dbLong: case dbSingle: size = 4; attrs = dbFixedField; break;
            case dbCurrency: case dbDouble: case dbDate: size = 8; attrs = dbFixedField; break;
            case dbGUID: case dbDecimal: size = 16; attrs = dbFixedField; break;
            default: size = k->len; attrs = dbFixedField;
            }
            if (k->flags & 4) attrs |= dbAutoIncrField;
            if (type == dbMemo && (k->flags & 0x80)) attrs |= dbHyperlinkField;
            jpblock *fb = jprops_block(props, 1, k->name, 0);
            char *dv = prop_str(props, fb, "DefaultValue"), *fr = prop_str(props, fb, "ValidationRule"), *ft = prop_str(props, fb, "ValidationText");
            sqlite3_bind_text(st, 1, t->name, -1, SQLITE_STATIC);
            sqlite3_bind_text(st, 2, k->name, -1, SQLITE_STATIC);
            sqlite3_bind_int(st, 3, k->id);
            sqlite3_bind_int(st, 4, type);
            sqlite3_bind_int(st, 5, size);
            sqlite3_bind_int(st, 6, attrs);
            sqlite3_bind_int(st, 7, prop_bool(fb, "Required"));
            /* AllowZeroLength: a field made by SQL DDL has neither it nor Required (2: reads False, "" allowed) */
            jprop *az = fb ? jprops_get(fb, "AllowZeroLength") : 0, *rq = fb ? jprops_get(fb, "Required") : 0;
            sqlite3_bind_int(st, 8, az ? (az->n && az->v[0] ? 1 : 0) : rq ? 0 : 2);
            sqlite3_bind_text(st, 9, dv, -1, free);
            sqlite3_bind_text(st, 10, fr, -1, free);
            sqlite3_bind_text(st, 11, ft, -1, free);
            if (k->type == JT_NUMERIC) sqlite3_bind_int(st, 12, k->misc & 0xFFFF); else sqlite3_bind_null(st, 12);
            sqlite3_step(st);
            sqlite3_reset(st);
        }
    sqlite3_finalize(st);
    sqlite3_prepare_v2(s, "INSERT OR IGNORE INTO MSysLB_Indexes (tbl, name, fields, isprimary, isunique, required, ignorenulls, isforeign, ord) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9)", -1, &st, 0);
    for (int i = 0; i < t->nlidx; i++) {
        jlidx *l = &t->lidx[i];
        if (l->name[0] == '.') continue;        /* the hidden side of a relation (.rB ...) */
        if (l->real < 0 || l->real >= t->nridx) continue;
        jridx *x = &t->ridx[l->real];
        jbuf fl = { 0 };
        for (int k = 0; k < x->n; k++) {
            jetcol *col = 0;
            for (int q = 0; q < t->ncols; q++) if (t->cols[q].num == x->col[k]) { col = &t->cols[q]; break; }
            if (!col) continue;
            if (fl.n) jb_s(&fl, ";");
            jb_s(&fl, x->order[k] == 1 ? "+" : "-");
            jb_s(&fl, col->name);
        }
        if (!fl.n) { jb_free(&fl); continue; }
        int primary = l->type == 1, foreign = l->type == 2;
        sqlite3_bind_text(st, 1, t->name, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, l->name, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 3, fl.p, -1, SQLITE_TRANSIENT);
        sqlite3_bind_int(st, 4, primary);
        sqlite3_bind_int(st, 5, primary || (x->flags & 1) ? 1 : 0);   /* (a foreign one: unique for a one-to-one relation) */
        sqlite3_bind_int(st, 6, primary || (!foreign && (x->flags & 8)) ? 1 : 0);
        sqlite3_bind_int(st, 7, !foreign && (x->flags & 2) ? 1 : 0);
        sqlite3_bind_int(st, 8, foreign);
        sqlite3_bind_int(st, 9, l->num);
        sqlite3_step(st);
        sqlite3_reset(st);
        jb_free(&fl);
    }
    sqlite3_finalize(st);
    save_props(s, props, 'T', t->name);
    (void)err; (void)cap;
    return 0;
}

static int import_rows(cat *c, jtable *t, jprops *props, sqlite3 *s, char *err, int cap)
{
    int nc = t->ncols;
    jetcol **cols = malloc(sizeof(jetcol *) * (size_t)(nc + 1));
    int k = 0;
    for (int num = 0; num < t->width; num++)
        for (int i = 0; i < nc; i++) if (t->cols[i].num == num) cols[k++] = &t->cols[i];
    jbuf sql = { 0 };
    jb_s(&sql, "INSERT INTO "); jb_qid(&sql, t->name, -1); jb_s(&sql, " (");
    for (int i = 0; i < k; i++) { if (i) jb_s(&sql, ", "); jb_qid(&sql, cols[i]->name, -1); }
    jb_s(&sql, ") VALUES (");
    for (int i = 0; i < k; i++) jb_f(&sql, "%s?%d", i ? ", " : "", i + 1);
    jb_s(&sql, ")");
    sqlite3_stmt *st = 0;
    int rc = sqlite3_prepare_v2(s, sql.p, -1, &st, 0);
    jb_free(&sql);
    if (rc != SQLITE_OK) { snprintf(err, (size_t)cap, "%s", sqlite3_errmsg(s)); free(cols); return -1; }
    int *req = calloc((size_t)(k + 1), sizeof(int));
    for (int i = 0; i < k; i++) req[i] = prop_bool(jprops_block(props, 1, cols[i]->name, 0), "Required");
    int r = 0, lost = 0;
    for (int row = 0; row < t->nrows && !r; row++) {
        for (int i = 0; i < k; i++) bind_value(st, i + 1, c->db, cols[i], jet_val(t, row, cols[i]->num), req[i]);
        rc = sqlite3_step(st);
        sqlite3_reset(st);
        if (rc != SQLITE_DONE && (rc & 0xff) != SQLITE_CONSTRAINT) { snprintf(err, (size_t)cap, "%s: %s", t->name, sqlite3_errmsg(s)); r = -1; }
        else if (rc != SQLITE_DONE && !lost++) snprintf(err, (size_t)cap, "%s: a row not imported: %s", t->name, sqlite3_errmsg(s));
    }
    sqlite3_finalize(st);
    free(cols); free(req);
    return r;
}

/* ---------------------------------------------------------------- linked tables */
static void import_links(cat *c, sqlite3 *s)
{
    sqlite3_stmt *st = 0;
    sqlite3_prepare_v2(s, "INSERT OR IGNORE INTO MSysLB_Tables (name, attrs, seq, srcname, connect) VALUES (?1, ?2, 0, ?3, ?4)", -1, &st, 0);
    for (int r = 0; r < c->t->nrows; r++) {
        int ty = val_int(cv(c, r, c->ctype), 2);
        if (ty != 6 && ty != 4) continue;
        char *name = val_text(c->db, cv(c, r, c->cname)), *dbs = c->cdb >= 0 ? val_text(c->db, cv(c, r, c->cdb)) : 0;
        char *con = c->cconn >= 0 ? val_text(c->db, cv(c, r, c->cconn)) : 0, *src = c->cforeign >= 0 ? val_text(c->db, cv(c, r, c->cforeign)) : 0;
        jbuf cs = { 0 };
        if (ty == 6) { jb_s(&cs, con && *con ? con : ""); if (!(con && strstr(con, "DATABASE="))) { jb_s(&cs, ";DATABASE="); jb_s(&cs, dbs ? dbs : ""); } }
        else jb_s(&cs, con ? con : "");
        sqlite3_bind_text(st, 1, name ? name : "", -1, SQLITE_TRANSIENT);
        sqlite3_bind_int(st, 2, ty == 6 ? dbAttachedTable : dbAttachedODBC);
        sqlite3_bind_text(st, 3, src ? src : (name ? name : ""), -1, SQLITE_TRANSIENT);
        sqlite3_bind_text(st, 4, cs.p ? cs.p : "", -1, SQLITE_TRANSIENT);
        sqlite3_step(st);
        sqlite3_reset(st);
        jb_free(&cs);
        free(name); free(dbs); free(con); free(src);
    }
    sqlite3_finalize(st);
}

/* ---------------------------------------------------------------- relations */
static void import_relations(cat *c, sqlite3 *s)
{
    jtable *rel = jet_table(c->db, "MSysRelationships");
    if (!rel) return;
    jetcol *cg = jet_col(rel, "grbit"), *ci = jet_col(rel, "icolumn"), *cc = jet_col(rel, "szColumn"), *co = jet_col(rel, "szObject");
    jetcol *crc = jet_col(rel, "szReferencedColumn"), *cro = jet_col(rel, "szReferencedObject"), *cr = jet_col(rel, "szRelationship");
    if (!cg || !ci || !cc || !co || !crc || !cro || !cr) return;
    sqlite3_stmt *st = 0;
    sqlite3_prepare_v2(s, "INSERT OR IGNORE INTO MSysLB_Relations (name, tbl, ftbl, fields, attrs) VALUES (?1, ?2, ?3, ?4, ?5)", -1, &st, 0);
    int n = rel->nrows;
    char **rname = calloc((size_t)(n + 1), sizeof(char *));
    for (int i = 0; i < n; i++) rname[i] = val_text(c->db, jet_val(rel, i, cr->num));
    for (int i = 0; i < n; i++) {
        if (!rname[i]) continue;
        int seen = 0;
        for (int k = 0; k < i; k++) if (rname[k] && !strcmp(rname[k], rname[i])) { seen = 1; break; }
        if (seen) continue;
        jbuf fl = { 0 }, ix = { 0 }, cols = { 0 };
        for (int pos = 0; pos < 32; pos++)
            for (int j = 0; j < n; j++) {
                if (!rname[j] || strcmp(rname[j], rname[i]) || val_int(jet_val(rel, j, ci->num), 4) != pos) continue;
                char *a = val_text(c->db, jet_val(rel, j, crc->num)), *b = val_text(c->db, jet_val(rel, j, cc->num));
                if (fl.n) { jb_s(&fl, ";"); jb_s(&ix, ";"); jb_s(&cols, ", "); }
                jb_s(&fl, a ? a : ""); jb_s(&fl, "="); jb_s(&fl, b ? b : "");
                jb_s(&ix, "+"); jb_s(&ix, b ? b : "");
                jb_qid(&cols, b ? b : "", -1);
                free(a); free(b);
            }
        char *tbl = val_text(c->db, jet_val(rel, i, cro->num)), *ftbl = val_text(c->db, jet_val(rel, i, co->num));
        int attrs = val_int(jet_val(rel, i, cg->num), 4);
        /* a relation of system tables (MSysAccessStorageSTGREL of Access): kept in the file, not in the copy */
        int tr = tbl ? cat_row(c, tbl, 1) : -1, fr = ftbl ? cat_row(c, ftbl, 1) : -1;
        if ((tr >= 0 && is_system(c, tr, jet_table(c->db, tbl))) || (fr >= 0 && is_system(c, fr, jet_table(c->db, ftbl)))) {   /* (a query's: kept) */
            free(tbl); free(ftbl); jb_free(&fl); jb_free(&ix); jb_free(&cols);
            continue;
        }
        sqlite3_bind_text(st, 1, rname[i], -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, tbl ? tbl : "", -1, SQLITE_TRANSIENT);
        sqlite3_bind_text(st, 3, ftbl ? ftbl : "", -1, SQLITE_TRANSIENT);
        sqlite3_bind_text(st, 4, fl.p ? fl.p : "", -1, SQLITE_TRANSIENT);
        sqlite3_bind_int(st, 5, attrs);
        sqlite3_step(st);
        sqlite3_reset(st);
        if (!(attrs & 2))    {     /* a relation not enforced has no index */
            sqlite3_stmt *x = 0;
            sqlite3_prepare_v2(s, "INSERT OR IGNORE INTO MSysLB_Indexes (tbl, name, fields, isprimary, isunique, required, ignorenulls, isforeign, ord) VALUES (?1, ?2, ?3, 0, ?4, 0, 0, 1, 99)", -1, &x, 0);   /* (one-to-one: unique) */
            sqlite3_bind_text(x, 1, ftbl ? ftbl : "", -1, SQLITE_TRANSIENT);
            sqlite3_bind_text(x, 2, rname[i], -1, SQLITE_STATIC);
            sqlite3_bind_text(x, 3, ix.p ? ix.p : "", -1, SQLITE_TRANSIENT);
            sqlite3_bind_int(x, 4, attrs & 1);
            sqlite3_step(x);
            int added = sqlite3_changes(s) > 0;
            sqlite3_finalize(x);
            if (added && cols.n) {
                jbuf q = { 0 }, nm = { 0 };
                jb_f(&nm, "%s|%s", ftbl ? ftbl : "", rname[i]);
                jb_s(&q, "CREATE INDEX "); jb_qid(&q, nm.p, -1); jb_s(&q, " ON "); jb_qid(&q, ftbl ? ftbl : "", -1);
                jb_s(&q, " ("); jb_put(&q, cols.p, cols.n); jb_s(&q, ")");
                sqlite3_exec(s, q.p, 0, 0, 0);
                jb_free(&q); jb_free(&nm);
            }
        }
        free(tbl); free(ftbl);
        jb_free(&fl); jb_free(&ix); jb_free(&cols);
    }
    for (int i = 0; i < n; i++) free(rname[i]);
    free(rname);
    sqlite3_finalize(st);
}

/* ---------------------------------------------------------------- queries */
/* the rows of a query in MSysQueries (malloc; *n of them) */
jqrow *jet_query_rows(jdb *db, uint32_t oid, int *n)
{
    jtable *q = jet_table(db, "MSysQueries");
    *n = 0;
    if (!q) return 0;
    jetcol *co = jet_col(q, "ObjectId"), *ca = jet_col(q, "Attribute"), *cor = jet_col(q, "Order"), *cf = jet_col(q, "Flag");
    jetcol *c1 = jet_col(q, "Name1"), *c2 = jet_col(q, "Name2"), *ce = jet_col(q, "Expression");
    if (!co || !ca || !cor || !cf || !c1 || !c2 || !ce) return 0;
    jqrow *rows = calloc((size_t)(q->nrows + 1), sizeof(jqrow));
    for (int r = 0; r < q->nrows; r++) {
        jv *o = jet_val(q, r, co->num);
        if (o->null || o->n < 4 || U32(o->p, 0) != oid) continue;
        jqrow *x = &rows[(*n)++];
        x->attr = val_int(jet_val(q, r, ca->num), 1);
        jv *ord = jet_val(q, r, cor->num);
        x->order = 0;
        for (int i = 0; !ord->null && i < ord->n && i < 4; i++) x->order = x->order << 8 | ord->p[i];
        jv *f = jet_val(q, r, cf->num);
        x->flag = f->null ? JQ_NOFLAG : (short)U16(f->p, 0);
        x->name1 = val_text(db, jet_val(q, r, c1->num));
        x->name2 = val_text(db, jet_val(q, r, c2->num));
        x->expr = val_text(db, jet_val(q, r, ce->num));
    }
    return rows;
}

static void import_queries(cat *c, sqlite3 *s)
{
    sqlite3_stmt *st = 0, *jq = 0;
    sqlite3_prepare_v2(s, "INSERT OR IGNORE INTO MSysLB_Queries (name, sql, type) VALUES (?1, ?2, ?3)", -1, &st, 0);
    sqlite3_exec(s, "CREATE TABLE IF NOT EXISTS MSysLB_JetQueries (name TEXT PRIMARY KEY COLLATE NOCASE, sql TEXT)", 0, 0, 0);
    sqlite3_prepare_v2(s, "INSERT OR REPLACE INTO MSysLB_JetQueries (name, sql) VALUES (?1, ?2)", -1, &jq, 0);
    for (int r = 0; r < c->t->nrows; r++) {
        if ((val_int(cv(c, r, c->ctype), 2) & 0x7F) != 5) continue;
        char *name = val_text(c->db, cv(c, r, c->cname));
        if (!name || name[0] == '~') { free(name); continue; }   /* the hidden queries of forms and reports */
        int n, qt;
        jqrow *rows = jet_query_rows(c->db, (uint32_t)val_int(cv(c, r, c->cid), 4), &n);
        char *sql = jq_text(rows, n, &qt);
        jq_free(rows, n);
        sqlite3_bind_text(st, 1, name, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, sql, -1, SQLITE_STATIC);
        sqlite3_bind_int(st, 3, qt);
        sqlite3_step(st);
        sqlite3_reset(st);
        /* the text as read: a query whose text the program does not change keeps its rows */
        sqlite3_bind_text(jq, 1, name, -1, SQLITE_STATIC);
        sqlite3_bind_text(jq, 2, sql, -1, SQLITE_STATIC);
        sqlite3_step(jq);
        sqlite3_reset(jq);
        jv *lp = cv(c, r, c->cprop);
        if (!lp->null && lp->n) {
            jprops *p = jprops_parse(lp->p, lp->n, c->db->jet4, c->db->codepage);
            save_props(s, p, 'Q', name);
            jprops_free(p);
        }
        free(sql); free(name);
    }
    sqlite3_finalize(st);
    sqlite3_finalize(jq);
}

/* ---------------------------------------------------------------- the import */
int mdb_import(const char *path, const char *sqlite_path, char *err, int cap)
{
    if (cap) *err = 0;
    jdb *db = jet_read(path, mdb_codepage(), err, cap);
    if (!db) { if (!*err) snprintf(err, (size_t)cap, "not a Jet database"); return -1; }
    cat c = { db, db->t[0], -1, -1, -1, -1, -1, -1, -1, -1, -1 };
    jetcol *k;
    if ((k = jet_col(c.t, "Id"))) c.cid = k->num;
    if ((k = jet_col(c.t, "ParentId"))) c.cparent = k->num;
    if ((k = jet_col(c.t, "Name"))) c.cname = k->num;
    if ((k = jet_col(c.t, "Type"))) c.ctype = k->num;
    if ((k = jet_col(c.t, "Flags"))) c.cflags = k->num;
    if ((k = jet_col(c.t, "LvProp"))) c.cprop = k->num;
    if ((k = jet_col(c.t, "Database"))) c.cdb = k->num;
    if ((k = jet_col(c.t, "Connect"))) c.cconn = k->num;
    if ((k = jet_col(c.t, "ForeignName"))) c.cforeign = k->num;
    if (c.cid < 0 || c.cname < 0 || c.ctype < 0 || c.cflags < 0 || c.cprop < 0) { jet_free(db); snprintf(err, (size_t)cap, "not a Jet database"); return -1; }
    char *tmp = malloc(strlen(sqlite_path) + 8);
    sprintf(tmp, "%s.part", sqlite_path);
    remove(tmp);
    sqlite3 *s = 0;
    if (sqlite3_open(tmp, &s) != SQLITE_OK) { jet_free(db); snprintf(err, (size_t)cap, "cannot write %s", tmp); free(tmp); return -1; }
    jet_register(s);
    {   /* the order of texts of the copy: the file's sort order (its indexes made in it) */
        int lcid = db->jet4 ? (int)(db->page0[0x6E] | db->page0[0x6F] << 8) : 0x409;
        jet_collation(s, lcid, db->jet4, mdb_codepage());
    }
    exec(s, "PRAGMA legacy_alter_table = ON; PRAGMA journal_mode = OFF; PRAGMA synchronous = OFF", err, cap);
    ddb *d = (ddb *)vb_onew(&vb_cls_dao_database);
    vb_oaddref((vbobj)d);
    d->db = s;
    d->closed = 1;                              /* not a database of the program: db_clear leaves it */
    int r = schema_init(d);
    if (!r) r = exec(s, "BEGIN", err, cap);
    /* the user's tables, as MSysObjects names them */
    int nt = 0;
    jtable **tabs = calloc((size_t)(db->nt + 1), sizeof(jtable *));
    jprops **props = calloc((size_t)(db->nt + 1), sizeof(jprops *));
    for (int i = 4; i < db->nt; i++) {
        jtable *t = db->t[i];
        int row = cat_row(&c, t->name, 1);
        if (row < 0 || is_system(&c, row, t)) continue;
        jv *lp = cv(&c, row, c.cprop);
        props[nt] = jprops_parse(lp->null ? 0 : lp->p, lp->n, db->jet4, db->codepage);
        tabs[nt++] = t;
    }
    for (int i = 0; !r && i < nt; i++) r = import_schema(&c, tabs[i], props[i], s, err, cap);
    if (!r) r = schema_build_all(d);
    /* the rows as the file has them: Jet checks a validation rule, AllowZeroLength, when a value is written, not the
     * rows already there (DEV-14/q1 039: ">Now()" of 1997's rows; 053: "" in a field without AllowZeroLength) */
    if (!r) r = schema_drop_triggers(d);
    if (!r) r = exec(s, "PRAGMA ignore_check_constraints = ON", err, cap);
    for (int i = 0; !r && i < nt; i++) r = import_rows(&c, tabs[i], props[i], s, err, cap);
    exec(s, "PRAGMA ignore_check_constraints = OFF", err, cap);
    if (!r) {
        import_links(&c, s);
        import_relations(&c, s);
        import_queries(&c, s);
        int row = cat_row(&c, "MSysDb", 2);
        if (row >= 0 && !cv(&c, row, c.cprop)->null) {
            jv *lp = cv(&c, row, c.cprop);
            jprops *p = jprops_parse(lp->p, lp->n, db->jet4, db->codepage);
            save_props(s, p, 'D', "");
            jprops_free(p);
        }
        char sql[300];
        snprintf(sql, sizeof sql, "INSERT OR REPLACE INTO MSysLB_Props VALUES ('Version', '%s')", db->jet4 ? "4.0" : "3.0");
        exec(s, sql, err, cap);
        exec(s, "INSERT OR REPLACE INTO MSysLB_Props VALUES ('Collation', 'keys1')", err, cap);
        int lcid = db->jet4 ? (int)(db->page0[0x6E] | db->page0[0x6F] << 8) : (int)(db->page0[0x3A] | db->page0[0x3B] << 8);
        if (lcid && lcid != 0x409) {            /* its sort order (CollatingOrder; a Jet 4 file keeps its keys) */
            snprintf(sql, sizeof sql, "INSERT OR REPLACE INTO MSysLB_Props VALUES ('SortLCID', '%d')", lcid);
            exec(s, sql, err, cap);
        }
        r = schema_after_import(d);
    }
    if (!r) r = exec(s, "COMMIT", err, cap);
    if (r && !*err) {                           /* a DAO error of the schema (its text), or SQLite's */
        if (vb_err_number_pending()) snprintf(err, (size_t)cap, "%d %s", (int)vb_err_number_pending(), vb_stoutf8(vb_err_description()));
        else snprintf(err, (size_t)cap, "%s", sqlite3_errmsg(s));
    }
    vb_err_clear_pending();
    d->db = 0;
    vb_orelease((vbobj)d);
    sqlite3_close(s);
    for (int i = 0; i < nt; i++) jprops_free(props[i]);
    free(tabs); free(props);
    jet_free(db);
    if (!r) {
        remove(sqlite_path);
        if (rename(tmp, sqlite_path)) { snprintf(err, (size_t)cap, "cannot write %s", sqlite_path); r = -1; }
    } else remove(tmp);
    free(tmp);
    return r;
}

/* ---------------------------------------------------------------- the password of a database
 * At 0x42 of page 0 (unmasked): Jet 4 20 characters UTF-16LE xor the creation date (0x72) as a 32 bit integer
 * repeated; Jet 3 20 bytes of the code page (VB6 oracle tests/db/d_jetpw) */
static void pw_mask(const unsigned char *p0, unsigned char m[4])
{
    double d;
    memcpy(&d, p0 + 0x72, 8);
    uint32_t k = (uint32_t)(int32_t)d;
    for (int i = 0; i < 4; i++) m[i] = (unsigned char)(k >> (8 * i));
}
void jet_password_get(const unsigned char *p0, int jet4, int codepage, char *out, int cap)
{
    unsigned char b[40];
    memcpy(b, p0 + 0x42, 40);
    uint16_t w[21];
    int n = 0;
    if (jet4) {
        unsigned char m[4];
        pw_mask(p0, m);
        for (int i = 0; i < 40; i++) b[i] ^= m[i % 4];
        while (n < 20 && (b[2 * n] | b[2 * n + 1])) { w[n] = (uint16_t)(b[2 * n] | b[2 * n + 1] << 8); n++; }
    } else {
        int k = 0;
        while (k < 20 && b[k]) k++;
        n = jet_cp_to_utf16(codepage, b, k, w);
    }
    char *s = jet_utf16_to_utf8(w, n);
    snprintf(out, (size_t)cap, "%s", s);
    free(s);
}
int jet_password_set(unsigned char *p0, int jet4, int codepage, const char *pw)
{
    uint16_t *w;
    int n = jet_utf8_to_utf16(pw ? pw : "", &w);
    unsigned char b[40] = { 0 };
    int r = 0;
    if (n > 20) r = -1;
    else if (jet4) {
        unsigned char m[4];
        pw_mask(p0, m);
        for (int i = 0; i < n; i++) { b[2 * i] = (unsigned char)w[i]; b[2 * i + 1] = (unsigned char)(w[i] >> 8); }
        for (int i = 0; i < 40; i++) b[i] ^= m[i % 4];
        memcpy(p0 + 0x42, b, 40);
    } else {
        unsigned char c[48];
        int k = jet_utf16_to_cp(codepage, w, n, c);
        if (k > 20) r = -1;
        else { memcpy(b, c, (size_t)k); memcpy(p0 + 0x42, b, 20); }
    }
    free(w);
    return r;
}
/* the password of a Jet file ("" none); -1 if it cannot be read */
int mdb_password(const char *path, char *out, int cap)
{
    unsigned char p0[4096];
    FILE *f = fopen(path, "rb");
    if (!f) return -1;
    size_t n = fread(p0, 1, 2048, f);
    fclose(f);
    if (n < 2048 || memcmp(p0 + 4, "Standard Jet DB", 15)) return -1;
    int jet4 = p0[0x14] >= 1;
    jet_rc4(0x6B39DAC7u, p0 + 0x18, jet4 ? 128 : 126);
    jet_password_get(p0, jet4, mdb_codepage(), out, cap);
    return 0;
}
