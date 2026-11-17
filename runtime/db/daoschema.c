/* daoschema.c - the schema of a LinBASIC database: what Jet keeps in its system tables (field types,
 * sizes, Required, AllowZeroLength, indexes, relations, saved queries) is kept in the MSysLB_* tables;
 * the SQLite table, its indexes and its triggers are made from them, so that SQLite enforces what Jet
 * enforces, with Jet's error numbers (the constraint names and the trigger messages carry them).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "daoint.h"
#include <ctype.h>

static const char *const META =
    "CREATE TABLE IF NOT EXISTS MSysLB_Tables (name TEXT PRIMARY KEY COLLATE NOCASE, attrs INT DEFAULT 0, seq INT DEFAULT 0,"
    " created TEXT, updated TEXT, vrule TEXT, vtext TEXT, srcname TEXT, connect TEXT);"
    "CREATE TABLE IF NOT EXISTS MSysLB_Fields (tbl TEXT COLLATE NOCASE, name TEXT COLLATE NOCASE, ord INT, type INT,"
    " size INT, attrs INT, required INT, allowzero INT, defval TEXT, vrule TEXT, vtext TEXT, PRIMARY KEY (tbl, name));"
    "CREATE TABLE IF NOT EXISTS MSysLB_Indexes (tbl TEXT COLLATE NOCASE, name TEXT COLLATE NOCASE, fields TEXT,"
    " isprimary INT, isunique INT, required INT, ignorenulls INT, isforeign INT, ord INT, PRIMARY KEY (tbl, name));"
    "CREATE TABLE IF NOT EXISTS MSysLB_Relations (name TEXT PRIMARY KEY COLLATE NOCASE, tbl TEXT, ftbl TEXT, fields TEXT, attrs INT);"
    "CREATE TABLE IF NOT EXISTS MSysLB_Queries (name TEXT PRIMARY KEY COLLATE NOCASE, sql TEXT, type INT, created TEXT, updated TEXT);"
    "CREATE TABLE IF NOT EXISTS MSysLB_Props (name TEXT PRIMARY KEY COLLATE NOCASE, value TEXT);";

int schema_init(ddb *d) { return dao_exec(d, META); }
static sqlite3_int64 seq_of(ddb *d, const char *table);
static void seq_set(ddb *d, const char *table, sqlite3_int64 n);

static char *sdup(const char *s) { return s ? strdup(s) : 0; }
static const char *col_text(sqlite3_stmt *st, int i) { const char *s = (const char *)sqlite3_column_text(st, i); return s ? s : ""; }

/* ---------------------------------------------------------------- queries on the schema */
static sqlite3_stmt *prep(ddb *d, const char *sql)
{
    sqlite3_stmt *st = 0;
    if (sqlite3_prepare_v2(d->db, sql, -1, &st, 0) != SQLITE_OK) return 0;
    return st;
}
static int sqlite_has(ddb *d, const char *type, const char *name, char *real, int cap)
{
    sqlite3_stmt *st = prep(d, "SELECT name FROM sqlite_master WHERE type = ?1 AND name = ?2 COLLATE NOCASE");
    if (!st) return 0;
    sqlite3_bind_text(st, 1, type, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 2, name, -1, SQLITE_STATIC);
    int r = sqlite3_step(st) == SQLITE_ROW;
    if (r && real) snprintf(real, (size_t)cap, "%s", col_text(st, 0));
    sqlite3_finalize(st);
    return r;
}
static int hidden_table(const char *n)
{
    return !strncasecmp(n, "MSysLB_", 7) || !strncasecmp(n, "sqlite_", 7);
}
int schema_has_table(ddb *d, const char *name, char *real, int cap)
{
    if (hidden_table(name)) return 0;
    if (sqlite_has(d, "table", name, real, cap)) return 1;
    sqlite3_stmt *st = prep(d, "SELECT name FROM MSysLB_Queries WHERE name = ?1");
    if (st) {
        sqlite3_bind_text(st, 1, name, -1, SQLITE_STATIC);
        int r = sqlite3_step(st) == SQLITE_ROW;
        if (r && real) snprintf(real, (size_t)cap, "%s", col_text(st, 0));
        sqlite3_finalize(st);
        if (r) return 2;
    }
    if (sqlite_has(d, "view", name, real, cap)) return 2;
    return 0;
}

static vb_date now_date(void) { return vb_now(); }
static void date_text(vb_date d, char out[24]) { jet_date_text(d, out); }

/* ---------------------------------------------------------------- objects */

void dao_field_init(vbobj o)
{
    dfld *f = (dfld *)o;
    f->name = strdup("");
    f->defval = strdup(""); f->vrule = strdup(""); f->vtext = strdup(""); f->foreign = strdup("");
    f->col = -1;
}
static void field_type(dfld *f, int type, int size)
{
    f->type = type; f->size = size;
    if (!type) return;
    switch (type) {
    case dbText: if (!size) f->size = 50; f->attrs = (f->attrs & ~dbFixedField) | dbVariableField; break;
    case dbMemo: case dbLongBinary: f->attrs = (f->attrs & ~dbFixedField) | dbVariableField; f->size = 0; break;
    case dbBinary: f->attrs = (f->attrs & ~dbFixedField) | dbVariableField; if (!size) f->size = 255; break;
    case dbBoolean: case dbByte: f->size = 1; f->attrs = (f->attrs & ~dbVariableField) | dbFixedField; break;
    case dbInteger: f->size = 2; f->attrs = (f->attrs & ~dbVariableField) | dbFixedField; break;
    case dbLong: case dbSingle: f->size = 4; f->attrs = (f->attrs & ~dbVariableField) | dbFixedField; break;
    case dbCurrency: case dbDouble: case dbDate: f->size = 8; f->attrs = (f->attrs & ~dbVariableField) | dbFixedField; break;
    case dbGUID: case dbDecimal: f->size = 16; f->attrs = (f->attrs & ~dbVariableField) | dbFixedField; break;
    default: f->attrs = (f->attrs & ~dbVariableField) | dbFixedField;
    }
}
void dao_field_settype(dfld *f, int type) { int a = f->attrs & dbAutoIncrField; f->attrs = a; field_type(f, type, f->type == dbText || f->type == dbBinary ? f->size : 0); }
dfld *field_new(int fk, vbobj owner, const char *name, int type, int size)
{
    dfld *f = (dfld *)vb_onew(&vb_cls_dao_field);
    vb_oaddref((vbobj)f);
    f->fk = fk;
    if (owner) { f->owner = owner; vb_oaddref(owner); }
    free(f->name); f->name = sdup(name ? name : "");
    field_type(f, type, size);
    return f;
}
void dao_index_init(vbobj o) { dix *ix = (dix *)o; ix->name = strdup(""); ix->fields = coll_new(C_IXFIELDS, o); }
dix *index_new(vbobj td, const char *name)
{
    dix *ix = (dix *)vb_onew(&vb_cls_dao_index);
    vb_oaddref((vbobj)ix);
    if (td) { ix->td = td; vb_oaddref(td); }
    free(ix->name); ix->name = sdup(name ? name : "");
    return ix;
}
void dao_tabledef_init(vbobj o)
{
    dtd *t = (dtd *)o;
    t->name = strdup("");
    t->fields = coll_new(C_FIELDS, o);
    t->indexes = coll_new(C_INDEXES, o);
    t->fields->loaded = t->indexes->loaded = 1;
    t->connect = strdup(""); t->srcname = strdup(""); t->vrule = strdup(""); t->vtext = strdup("");
    t->created = t->updated = now_date();
}
dtd *tabledef_new(const char *name)
{
    dtd *t = (dtd *)vb_onew(&vb_cls_dao_tabledef);
    vb_oaddref((vbobj)t);
    free(t->name); t->name = sdup(name ? name : "");
    return t;
}
void dao_relation_init(vbobj o)
{
    drel *r = (drel *)o;
    r->name = strdup(""); r->table = strdup(""); r->ftable = strdup("");
    r->fields = coll_new(C_RELFIELDS, o);
    r->fields->loaded = 1;
}
drel *relation_new(const char *name, const char *table, const char *ftable, int attrs)
{
    drel *r = (drel *)vb_onew(&vb_cls_dao_relation);
    vb_oaddref((vbobj)r);
    free(r->name); r->name = sdup(name ? name : "");
    free(r->table); r->table = sdup(table ? table : "");
    free(r->ftable); r->ftable = sdup(ftable ? ftable : "");
    r->attrs = attrs;
    return r;
}
dpar *param_new(const char *name, int type)
{
    dpar *p = (dpar *)vb_onew(&vb_cls_dao_parameter);
    vb_oaddref((vbobj)p);
    p->name = sdup(name);
    p->type = type;
    p->dir = 1;
    p->value = vb_v_empty();
    return p;
}

/* ---------------------------------------------------------------- reading the schema */
static void fill_field_from_meta(dfld *f, sqlite3_stmt *st)
{
    /* name, ord, type, size, attrs, required, allowzero, defval, vrule, vtext */
    f->ord = sqlite3_column_int(st, 1);
    f->type = sqlite3_column_int(st, 2);
    f->size = sqlite3_column_int(st, 3);
    f->attrs = sqlite3_column_int(st, 4);
    f->required = sqlite3_column_int(st, 5);
    f->allowzero = sqlite3_column_int(st, 6);
    free(f->defval); f->defval = strdup(col_text(st, 7));
    free(f->vrule); f->vrule = strdup(col_text(st, 8));
    free(f->vtext); f->vtext = strdup(col_text(st, 9));
}

static int has_meta(ddb *d, const char *table)
{
    sqlite3_stmt *st = prep(d, "SELECT 1 FROM MSysLB_Tables WHERE name = ?1");
    if (!st) return 0;
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    int r = sqlite3_step(st) == SQLITE_ROW;
    sqlite3_finalize(st);
    return r;
}

/* a table of a SQLite database made by another program: its fields from SQLite itself */
static void load_foreign_fields(ddb *d, dtd *t)
{
    jbuf q = { 0 };
    jb_s(&q, "PRAGMA table_info("); jb_qid(&q, t->name, -1); jb_s(&q, ")");
    sqlite3_stmt *st = prep(d, q.p);
    jb_free(&q);
    if (!st) return;
    int ord = 0;
    while (sqlite3_step(st) == SQLITE_ROW) {
        int size, attrs;
        int type = jet_decl_type(col_text(st, 2), &size, &attrs);
        if (!type) type = dbText, size = 255, attrs = dbVariableField;
        dfld *f = field_new(FK_TABLE, (vbobj)t, col_text(st, 1), type, size);
        f->attrs = attrs | dbUpdatableField;
        f->required = sqlite3_column_int(st, 3);
        f->ord = ord++;
        f->appended = 1;
        coll_add(t->fields, (vbobj)f);
    }
    sqlite3_finalize(st);
}

dtd *schema_load_table(ddb *d, const char *name)
{
    char real[256];
    if (!sqlite_has(d, "table", name, real, sizeof real)) return 0;
    dtd *t = tabledef_new(real);
    t->db = (vbobj)d; vb_oaddref((vbobj)d);
    t->appended = 1;
    if (!has_meta(d, real)) { load_foreign_fields(d, t); return t; }
    sqlite3_stmt *st = prep(d, "SELECT attrs, created, updated, vrule, vtext, srcname, connect FROM MSysLB_Tables WHERE name = ?1");
    if (st) {
        sqlite3_bind_text(st, 1, real, -1, SQLITE_STATIC);
        if (sqlite3_step(st) == SQLITE_ROW) {
            t->attrs = sqlite3_column_int(st, 0);
            jet_text_date(col_text(st, 1), (int)strlen(col_text(st, 1)), &t->created);
            jet_text_date(col_text(st, 2), (int)strlen(col_text(st, 2)), &t->updated);
            free(t->vrule); t->vrule = strdup(col_text(st, 3));
            free(t->vtext); t->vtext = strdup(col_text(st, 4));
            free(t->srcname); t->srcname = strdup(col_text(st, 5));
            free(t->connect); t->connect = strdup(col_text(st, 6));
        }
        sqlite3_finalize(st);
    }
    st = prep(d, "SELECT name, ord, type, size, attrs, required, allowzero, defval, vrule, vtext FROM MSysLB_Fields WHERE tbl = ?1 ORDER BY ord");
    if (st) {
        sqlite3_bind_text(st, 1, real, -1, SQLITE_STATIC);
        while (sqlite3_step(st) == SQLITE_ROW) {
            dfld *f = field_new(FK_TABLE, (vbobj)t, col_text(st, 0), 0, 0);
            fill_field_from_meta(f, st);
            f->appended = 1;
            coll_add(t->fields, (vbobj)f);
        }
        sqlite3_finalize(st);
    }
    /* Jet lists the indexes of a table it reads by name */
    st = prep(d, "SELECT name, fields, isprimary, isunique, required, ignorenulls, isforeign FROM MSysLB_Indexes WHERE tbl = ?1 ORDER BY name COLLATE NOCASE");
    if (st) {
        sqlite3_bind_text(st, 1, real, -1, SQLITE_STATIC);
        while (sqlite3_step(st) == SQLITE_ROW) {
            dix *ix = index_new((vbobj)t, col_text(st, 0));
            ix->primary = sqlite3_column_int(st, 2);
            ix->unique = sqlite3_column_int(st, 3);
            ix->required = sqlite3_column_int(st, 4);
            ix->ignorenulls = sqlite3_column_int(st, 5);
            ix->foreign = sqlite3_column_int(st, 6);
            ix->appended = 1;
            /* fields: "+A;-B" */
            char *fl = strdup(col_text(st, 1)), *save = 0;
            for (char *p = strtok_r(fl, ";", &save); p; p = strtok_r(0, ";", &save)) {
                int desc = *p == '-';
                if (*p == '+' || *p == '-') p++;
                dfld *f = field_new(FK_INDEX, (vbobj)ix, p, 0, 0);
                f->attrs = desc ? dbDescending : 0;
                f->appended = 1;
                coll_add(ix->fields, (vbobj)f);
            }
            free(fl);
            coll_add(t->indexes, (vbobj)ix);
        }
        sqlite3_finalize(st);
    }
    return t;
}

int schema_field_info(ddb *d, const char *table, const char *field, rcol *out)
{
    sqlite3_stmt *st = prep(d, "SELECT name, ord, type, size, attrs, required, allowzero, defval FROM MSysLB_Fields WHERE tbl = ?1 AND name = ?2");
    if (!st) return 0;
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 2, field, -1, SQLITE_STATIC);
    int r = 0;
    if (sqlite3_step(st) == SQLITE_ROW) {
        out->type = sqlite3_column_int(st, 2);
        out->size = sqlite3_column_int(st, 3);
        out->attrs = sqlite3_column_int(st, 4);
        out->required = sqlite3_column_int(st, 5);
        out->allowzero = sqlite3_column_int(st, 6) != 0;
        free(out->defval);
        out->defval = strdup(col_text(st, 7));
        out->autoinc = (out->attrs & dbAutoIncrField) != 0;
        r = 1;
    }
    sqlite3_finalize(st);
    return r;
}

int schema_text_size(ddb *d, const char *table, const char *field)
{
    rcol c; memset(&c, 0, sizeof c);
    int r = schema_field_info(d, table, field, &c) && c.type == dbText ? c.size : 0;
    free(c.defval);
    return r;
}

/* ---------------------------------------------------------------- the SQLite table from the schema */
typedef struct mfield { char *name; int type, size, attrs, required, allowzero; char *defval, *vrule; } mfield;
typedef struct mindex { char *name; char *fields; int primary, unique, required, ignorenulls, foreign; } mindex;

static int load_meta(ddb *d, const char *table, mfield **fs, int *nf, mindex **is, int *ni)
{
    *fs = 0; *nf = 0; *is = 0; *ni = 0;
    sqlite3_stmt *st = prep(d, "SELECT name, type, size, attrs, required, allowzero, defval, vrule FROM MSysLB_Fields WHERE tbl = ?1 ORDER BY ord");
    if (!st) return -1;
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    while (sqlite3_step(st) == SQLITE_ROW) {
        *fs = realloc(*fs, sizeof(mfield) * (size_t)(*nf + 1));
        mfield *f = &(*fs)[(*nf)++];
        f->name = strdup(col_text(st, 0));
        f->type = sqlite3_column_int(st, 1); f->size = sqlite3_column_int(st, 2); f->attrs = sqlite3_column_int(st, 3);
        f->required = sqlite3_column_int(st, 4); f->allowzero = sqlite3_column_int(st, 5);
        f->defval = strdup(col_text(st, 6)); f->vrule = strdup(col_text(st, 7));
    }
    sqlite3_finalize(st);
    st = prep(d, "SELECT name, fields, isprimary, isunique, required, ignorenulls, isforeign FROM MSysLB_Indexes WHERE tbl = ?1 ORDER BY ord");
    if (!st) return -1;
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    while (sqlite3_step(st) == SQLITE_ROW) {
        *is = realloc(*is, sizeof(mindex) * (size_t)(*ni + 1));
        mindex *x = &(*is)[(*ni)++];
        x->name = strdup(col_text(st, 0)); x->fields = strdup(col_text(st, 1));
        x->primary = sqlite3_column_int(st, 2); x->unique = sqlite3_column_int(st, 3); x->required = sqlite3_column_int(st, 4);
        x->ignorenulls = sqlite3_column_int(st, 5); x->foreign = sqlite3_column_int(st, 6);
    }
    sqlite3_finalize(st);
    return 0;
}
static void free_meta(mfield *fs, int nf, mindex *is, int ni)
{
    for (int i = 0; i < nf; i++) { free(fs[i].name); free(fs[i].defval); free(fs[i].vrule); }
    for (int i = 0; i < ni; i++) { free(is[i].name); free(is[i].fields); }
    free(fs); free(is);
}

/* a constant Jet DefaultValue as SQLite text, or 0 */
static int default_sql(const char *dv, jbuf *out)
{
    if (!dv || !*dv) return 0;
    jbuf e = { 0 };
    char err[200];
    if (jet_expr(dv, &e, 0, err, sizeof err)) { jb_free(&e); return 0; }
    if (strstr(e.p, "jet_now") || strstr(e.p, "jet_date") || strstr(e.p, "jet_time") || strchr(e.p, '"')) {
        jb_free(&e); return 0;                  /* not a constant: the recordset sets it */
    }
    jb_s(out, " DEFAULT ("); jb_put(out, e.p, e.n); jb_s(out, ")");
    jb_free(&e);
    return 1;
}

static void split_fields(const char *s, char ***names, int **desc, int *n)
{
    *names = 0; *desc = 0; *n = 0;
    char *fl = strdup(s), *save = 0;
    for (char *p = strtok_r(fl, ";", &save); p; p = strtok_r(0, ";", &save)) {
        *names = realloc(*names, sizeof(char *) * (size_t)(*n + 1));
        *desc = realloc(*desc, sizeof(int) * (size_t)(*n + 1));
        (*desc)[*n] = *p == '-';
        if (*p == '+' || *p == '-') p++;
        (*names)[(*n)++] = strdup(p);
    }
    free(fl);
}
static void free_names(char **n, int *d, int k) { for (int i = 0; i < k; i++) free(n[i]); free(n); free(d); }

static int make_table_sql(ddb *d, const char *table, const char *as_name, jbuf *out)
{
    mfield *fs; int nf; mindex *is; int ni;
    if (load_meta(d, table, &fs, &nf, &is, &ni) < 0) return -1;
    jb_s(out, "CREATE TABLE "); jb_qid(out, as_name, -1); jb_s(out, " (");
    int counter_done = 0;
    for (int i = 0; i < nf; i++) {
        mfield *f = &fs[i];
        char buf[48];
        if (i) jb_s(out, ", ");
        jb_qid(out, f->name, -1);
        int ai = (f->attrs & dbAutoIncrField) && !counter_done;
        if (ai) { jb_s(out, " INTEGER PRIMARY KEY AUTOINCREMENT"); counter_done = 1; continue; }
        if (f->type == dbText) jb_f(out, " TEXT(%d) COLLATE NOCASE", f->size > 0 ? f->size : 255);
        else if (f->type == dbMemo || f->type == dbGUID) jb_f(out, " %s COLLATE NOCASE", jet_type_decl(f->type, 0, buf, sizeof buf));
        else if (f->type == dbBinary) jb_f(out, " BINARY_BLOB(%d)", f->size > 0 ? f->size : 255);
        else jb_f(out, " %s", jet_type_decl(f->type, 0, buf, sizeof buf));
        if (f->required) jb_s(out, " NOT NULL");
        if (!default_sql(f->defval, out) && f->type == dbBoolean) jb_s(out, " DEFAULT 0");
        if (!f->allowzero && (f->type == dbText || f->type == dbMemo)) {
            jb_s(out, " CONSTRAINT "); jbuf cn = { 0 }; jb_f(&cn, "LB3315 %s.%s", table, f->name);
            jb_qid(out, cn.p, -1); jb_free(&cn);
            jb_s(out, " CHECK ("); jb_qid(out, f->name, -1); jb_s(out, " <> '')");
        }
    }
    for (int i = 0; i < ni; i++) {
        if (!is[i].primary && !is[i].required) continue;
        char **names; int *desc, n;
        split_fields(is[i].fields, &names, &desc, &n);
        jb_s(out, ", CONSTRAINT ");
        jbuf cn = { 0 }; jb_f(&cn, "LB%d %s", is[i].primary ? 3058 : 3316, is[i].name);
        jb_qid(out, cn.p, -1); jb_free(&cn);
        jb_s(out, " CHECK (");
        for (int k = 0; k < n; k++) { if (k) jb_s(out, " AND "); jb_qid(out, names[k], -1); jb_s(out, " IS NOT NULL"); }
        jb_s(out, ")");
        free_names(names, desc, n);
    }
    jb_s(out, ")");
    free_meta(fs, nf, is, ni);
    return 0;
}

static void index_sql(const char *table, mindex *x, jbuf *out)
{
    char **names; int *desc, n;
    split_fields(x->fields, &names, &desc, &n);
    jb_s(out, x->primary || x->unique ? "CREATE UNIQUE INDEX " : "CREATE INDEX ");
    jbuf nm = { 0 }; jb_f(&nm, "%s|%s", table, x->name);
    jb_qid(out, nm.p, -1); jb_free(&nm);
    jb_s(out, " ON "); jb_qid(out, table, -1); jb_s(out, " (");
    for (int k = 0; k < n; k++) { if (k) jb_s(out, ", "); jb_qid(out, names[k], -1); if (desc[k]) jb_s(out, " DESC"); }
    jb_s(out, ")");
    free_names(names, desc, n);
}

/* the triggers of a table: text cut to its size (Jet SQL does), the counter that is never used again */
static void table_triggers(ddb *d, const char *table, jbuf *out)
{
    mfield *fs; int nf; mindex *is; int ni;
    if (load_meta(d, table, &fs, &nf, &is, &ni) < 0) return;
    jbuf cut = { 0 };
    for (int i = 0; i < nf; i++) {
        if (fs[i].type != dbText || fs[i].size <= 0) continue;
        jb_s(&cut, " UPDATE "); jb_qid(&cut, table, -1); jb_s(&cut, " SET "); jb_qid(&cut, fs[i].name, -1);
        jb_f(&cut, " = substr("); jb_qid(&cut, fs[i].name, -1); jb_f(&cut, ", 1, %d) WHERE rowid = NEW.rowid AND length(NEW.", fs[i].size);
        jb_qid(&cut, fs[i].name, -1); jb_f(&cut, ") > %d;", fs[i].size);
    }
    for (int w = 0; w < 2 && cut.n; w++) {
        jb_s(out, "CREATE TRIGGER ");
        jbuf nm = { 0 }; jb_f(&nm, "lbt:%s:%s", table, w ? "upd" : "ins");
        jb_qid(out, nm.p, -1); jb_free(&nm);
        jb_s(out, w ? " AFTER UPDATE ON " : " AFTER INSERT ON "); jb_qid(out, table, -1);
        jb_s(out, " BEGIN"); jb_put(out, cut.p, cut.n); jb_s(out, " END;");
    }
    jb_free(&cut);
    free_meta(fs, nf, is, ni);
}

/* the triggers of every relation: Jet's referential integrity, with its error numbers */
static void relation_triggers(ddb *d, jbuf *out)
{
    sqlite3_stmt *st = prep(d, "SELECT name, tbl, ftbl, fields, attrs FROM MSysLB_Relations");
    if (!st) return;
    while (sqlite3_step(st) == SQLITE_ROW) {
        const char *name = col_text(st, 0), *pt = col_text(st, 1), *ct = col_text(st, 2), *fl = col_text(st, 3);
        int attrs = sqlite3_column_int(st, 4);
        if (attrs & dbRelationDontEnforce) continue;
        /* fields: "pk=fk;pk2=fk2" */
        char *pk[16], *fk[16]; int n = 0;
        char *copy = strdup(fl), *save = 0;
        for (char *p = strtok_r(copy, ";", &save); p && n < 16; p = strtok_r(0, ";", &save)) {
            char *eq = strchr(p, '=');
            if (!eq) continue;
            *eq = 0; pk[n] = p; fk[n] = eq + 1; n++;
        }
        if (!n) { free(copy); continue; }
        jbuf match_new_child = { 0 }, child_nonnull = { 0 }, match_old_parent = { 0 }, key_changed = { 0 };
        for (int i = 0; i < n; i++) {
            if (i) { jb_s(&match_new_child, " AND "); jb_s(&child_nonnull, " AND "); jb_s(&match_old_parent, " AND "); jb_s(&key_changed, " OR "); }
            jb_qid(&match_new_child, pk[i], -1); jb_s(&match_new_child, " = NEW."); jb_qid(&match_new_child, fk[i], -1);
            jb_s(&child_nonnull, "NEW."); jb_qid(&child_nonnull, fk[i], -1); jb_s(&child_nonnull, " IS NOT NULL");
            jb_qid(&match_old_parent, fk[i], -1); jb_s(&match_old_parent, " = OLD."); jb_qid(&match_old_parent, pk[i], -1);
            jb_s(&key_changed, "NEW."); jb_qid(&key_changed, pk[i], -1); jb_s(&key_changed, " IS NOT OLD."); jb_qid(&key_changed, pk[i], -1);
        }
        for (int w = 0; w < 2; w++) {           /* the child: a new or changed row needs its parent */
            jb_s(out, "CREATE TRIGGER ");
            jbuf nm = { 0 }; jb_f(&nm, "lbr:%s:c%d", name, w); jb_qid(out, nm.p, -1); jb_free(&nm);
            jb_s(out, w ? " BEFORE UPDATE ON " : " BEFORE INSERT ON "); jb_qid(out, ct, -1);
            jb_s(out, " WHEN "); jb_put(out, child_nonnull.p, child_nonnull.n);
            jb_s(out, " AND NOT EXISTS (SELECT 1 FROM "); jb_qid(out, pt, -1); jb_s(out, " WHERE ");
            jb_put(out, match_new_child.p, match_new_child.n);
            jb_s(out, ") BEGIN SELECT RAISE(ABORT, "); jbuf m = { 0 }; jb_f(&m, "LB3201 %s", pt); jb_qstr(out, m.p, -1); jb_free(&m);
            jb_s(out, "); END;");
        }
        /* the parent: deleting or changing the key of a row with related rows */
        jb_s(out, "CREATE TRIGGER ");
        { jbuf nm = { 0 }; jb_f(&nm, "lbr:%s:pd", name); jb_qid(out, nm.p, -1); jb_free(&nm); }
        if (attrs & dbRelationDeleteCascade) {
            jb_s(out, " AFTER DELETE ON "); jb_qid(out, pt, -1); jb_s(out, " BEGIN DELETE FROM "); jb_qid(out, ct, -1);
            jb_s(out, " WHERE "); jb_put(out, match_old_parent.p, match_old_parent.n); jb_s(out, "; END;");
        } else {
            jb_s(out, " BEFORE DELETE ON "); jb_qid(out, pt, -1); jb_s(out, " WHEN EXISTS (SELECT 1 FROM "); jb_qid(out, ct, -1);
            jb_s(out, " WHERE "); jb_put(out, match_old_parent.p, match_old_parent.n);
            jb_s(out, ") BEGIN SELECT RAISE(ABORT, "); jbuf m = { 0 }; jb_f(&m, "LB3200 %s", ct); jb_qstr(out, m.p, -1); jb_free(&m);
            jb_s(out, "); END;");
        }
        jb_s(out, "CREATE TRIGGER ");
        { jbuf nm = { 0 }; jb_f(&nm, "lbr:%s:pu", name); jb_qid(out, nm.p, -1); jb_free(&nm); }
        if (attrs & dbRelationUpdateCascade) {
            jb_s(out, " AFTER UPDATE ON "); jb_qid(out, pt, -1); jb_s(out, " WHEN "); jb_put(out, key_changed.p, key_changed.n);
            jb_s(out, " BEGIN UPDATE "); jb_qid(out, ct, -1); jb_s(out, " SET ");
            for (int i = 0; i < n; i++) { if (i) jb_s(out, ", "); jb_qid(out, fk[i], -1); jb_s(out, " = NEW."); jb_qid(out, pk[i], -1); }
            jb_s(out, " WHERE "); jb_put(out, match_old_parent.p, match_old_parent.n); jb_s(out, "; END;");
        } else {
            jb_s(out, " BEFORE UPDATE ON "); jb_qid(out, pt, -1); jb_s(out, " WHEN ("); jb_put(out, key_changed.p, key_changed.n);
            jb_s(out, ") AND EXISTS (SELECT 1 FROM "); jb_qid(out, ct, -1); jb_s(out, " WHERE ");
            jb_put(out, match_old_parent.p, match_old_parent.n);
            jb_s(out, ") BEGIN SELECT RAISE(ABORT, "); jbuf m = { 0 }; jb_f(&m, "LB3200 %s", ct); jb_qstr(out, m.p, -1); jb_free(&m);
            jb_s(out, "); END;");
        }
        jb_free(&match_new_child); jb_free(&child_nonnull); jb_free(&match_old_parent); jb_free(&key_changed);
        free(copy);
    }
    sqlite3_finalize(st);
}

/* every trigger of the database made again from the schema */
static int regen_triggers(ddb *d)
{
    jbuf drop = { 0 };
    sqlite3_stmt *st = prep(d, "SELECT name FROM sqlite_master WHERE type = 'trigger' AND (name LIKE 'lbt:%' OR name LIKE 'lbr:%')");
    if (st) {
        while (sqlite3_step(st) == SQLITE_ROW) { jb_s(&drop, "DROP TRIGGER "); jb_qid(&drop, col_text(st, 0), -1); jb_s(&drop, ";"); }
        sqlite3_finalize(st);
    }
    if (drop.n && dao_exec(d, drop.p)) { jb_free(&drop); return -1; }
    jb_free(&drop);
    jbuf mk = { 0 };
    st = prep(d, "SELECT name FROM MSysLB_Tables");
    if (st) {
        while (sqlite3_step(st) == SQLITE_ROW) table_triggers(d, col_text(st, 0), &mk);
        sqlite3_finalize(st);
    }
    relation_triggers(d, &mk);
    int r = mk.n ? dao_exec(d, mk.p) : 0;
    jb_free(&mk);
    return r;
}

/* the SQLite table made (again) from its schema; the rows kept */
static int build_table(ddb *d, const char *table, int rebuild)
{
    jbuf sql = { 0 };
    int r = 0;
    sqlite3_int64 seq = rebuild ? seq_of(d, table) : 0;
    if (rebuild) {
        /* the views and triggers keep naming the table (legacy_alter_table) */
        jb_s(&sql, "ALTER TABLE "); jb_qid(&sql, table, -1); jb_s(&sql, " RENAME TO \"MSysLB_old\";");
        if ((r = dao_exec(d, sql.p))) { jb_free(&sql); return r; }
        sql.n = 0;
    }
    if (make_table_sql(d, table, table, &sql) < 0) { jb_free(&sql); return -1; }
    jb_s(&sql, ";");
    if ((r = dao_exec(d, sql.p))) { jb_free(&sql); return r; }
    sql.n = 0;
    if (rebuild) {
        /* the columns both have */
        jbuf cols = { 0 };
        sqlite3_stmt *st = prep(d, "SELECT name FROM pragma_table_info('MSysLB_old') WHERE name IN (SELECT name FROM MSysLB_Fields WHERE tbl = ?1)");
        if (st) {
            sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
            while (sqlite3_step(st) == SQLITE_ROW) { if (cols.n) jb_s(&cols, ", "); jb_qid(&cols, col_text(st, 0), -1); }
            sqlite3_finalize(st);
        }
        if (cols.n) {
            jb_s(&sql, "INSERT INTO "); jb_qid(&sql, table, -1); jb_s(&sql, " ("); jb_put(&sql, cols.p, cols.n);
            jb_s(&sql, ") SELECT "); jb_put(&sql, cols.p, cols.n); jb_s(&sql, " FROM \"MSysLB_old\";");
        }
        jb_s(&sql, "DROP TABLE \"MSysLB_old\";");
        jb_free(&cols);
        if ((r = dao_exec(d, sql.p))) { jb_free(&sql); return r; }
        sql.n = 0;
        if (seq) seq_set(d, table, seq);
    }
    mfield *fs; int nf; mindex *is; int ni;
    if (load_meta(d, table, &fs, &nf, &is, &ni) == 0) {
        for (int i = 0; i < ni; i++) { index_sql(table, &is[i], &sql); jb_s(&sql, ";"); }
        free_meta(fs, nf, is, ni);
    }
    if (sql.n && (r = dao_exec(d, sql.p))) { jb_free(&sql); return r; }
    jb_free(&sql);
    return regen_triggers(d);
}

static int meta_insert_field(ddb *d, const char *table, dfld *f, int ord)
{
    sqlite3_stmt *st = prep(d, "INSERT INTO MSysLB_Fields (tbl, name, ord, type, size, attrs, required, allowzero, defval, vrule, vtext) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11)");
    if (!st) return dao_raise_sqlite(d, sqlite3_errcode(d->db), "DAO.Fields");
    int attrs = f->attrs;
    if (f->type == dbText || f->type == dbMemo || f->type == dbBinary || f->type == dbLongBinary) attrs = (attrs & ~dbFixedField) | dbVariableField;
    else attrs = (attrs & ~dbVariableField) | dbFixedField;
    attrs &= ~dbUpdatableField;
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 2, f->name, -1, SQLITE_STATIC);
    sqlite3_bind_int(st, 3, ord);
    sqlite3_bind_int(st, 4, f->type);
    sqlite3_bind_int(st, 5, f->size);
    sqlite3_bind_int(st, 6, attrs);
    sqlite3_bind_int(st, 7, f->required);
    sqlite3_bind_int(st, 8, f->allowzero);
    sqlite3_bind_text(st, 9, f->defval ? f->defval : "", -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 10, f->vrule ? f->vrule : "", -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 11, f->vtext ? f->vtext : "", -1, SQLITE_STATIC);
    int rc = sqlite3_step(st);
    sqlite3_finalize(st);
    if (rc != SQLITE_DONE) return dao_raise_sqlite(d, rc, "DAO.Fields");
    f->attrs = attrs;
    return 0;
}
static int meta_insert_index(ddb *d, const char *table, dix *ix, int ord)
{
    jbuf fl = { 0 };
    for (int i = 0; i < ix->fields->n; i++) {
        dfld *f = (dfld *)ix->fields->items[i];
        if (i) jb_s(&fl, ";");
        jb_s(&fl, (f->attrs & dbDescending) ? "-" : "+");
        jb_s(&fl, f->name);
    }
    sqlite3_stmt *st = prep(d, "INSERT INTO MSysLB_Indexes (tbl, name, fields, isprimary, isunique, required, ignorenulls, isforeign, ord) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9)");
    if (!st) { jb_free(&fl); return dao_raise_sqlite(d, sqlite3_errcode(d->db), "DAO.Indexes"); }
    if (ix->primary) { ix->unique = 1; ix->required = 1; }
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 2, ix->name, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 3, fl.p ? fl.p : "", -1, SQLITE_STATIC);
    sqlite3_bind_int(st, 4, ix->primary);
    sqlite3_bind_int(st, 5, ix->unique);
    sqlite3_bind_int(st, 6, ix->required);
    sqlite3_bind_int(st, 7, ix->ignorenulls);
    sqlite3_bind_int(st, 8, ix->foreign);
    sqlite3_bind_int(st, 9, ord);
    int rc = sqlite3_step(st);
    sqlite3_finalize(st);
    jb_free(&fl);
    if (rc == SQLITE_CONSTRAINT) { dao_raise(3284, "DAO.Indexes", "Index already exists."); return 3284; }
    if (rc != SQLITE_DONE) return dao_raise_sqlite(d, rc, "DAO.Indexes");
    return 0;
}
static int meta_insert_table(ddb *d, const char *name, int attrs)
{
    sqlite3_stmt *st = prep(d, "INSERT INTO MSysLB_Tables (name, attrs, seq, created, updated) VALUES (?1, ?2, 0, ?3, ?3)");
    if (!st) return dao_raise_sqlite(d, sqlite3_errcode(d->db), "DAO.TableDefs");
    char t[24]; date_text(now_date(), t);
    sqlite3_bind_text(st, 1, name, -1, SQLITE_STATIC);
    sqlite3_bind_int(st, 2, attrs);
    sqlite3_bind_text(st, 3, t, -1, SQLITE_STATIC);
    int rc = sqlite3_step(st);
    sqlite3_finalize(st);
    if (rc != SQLITE_DONE) return dao_raise_sqlite(d, rc, "DAO.TableDefs");
    return 0;
}

static int begin(ddb *d) { return dao_exec(d, "SAVEPOINT lbschema"); }
static int end(ddb *d, int err)
{
    if (err) { sqlite3_exec(d->db, "ROLLBACK TO lbschema; RELEASE lbschema", 0, 0, 0); return err; }
    return dao_exec(d, "RELEASE lbschema");
}

int schema_create_table(ddb *d, dtd *t)
{
    if (!t->name || !*t->name) { dao_raise(3264, "DAO.TableDefs", "No field defined--cannot append TableDef or Index."); return 3264; }
    if (schema_has_table(d, t->name, 0, 0) || hidden_table(t->name)) {
        dao_raise(3010, "DAO.TableDefs", "Table '%s' already exists.", t->name);
        return 3010;
    }
    if (t->fields->n == 0) { dao_raise(3264, "DAO.TableDefs", "No field defined--cannot append TableDef or Index."); return 3264; }
    int r = begin(d);
    if (r) return r;
    r = meta_insert_table(d, t->name, t->attrs);
    for (int i = 0; !r && i < t->fields->n; i++) r = meta_insert_field(d, t->name, (dfld *)t->fields->items[i], i);
    for (int i = 0; !r && i < t->indexes->n; i++) {
        dix *ix = (dix *)t->indexes->items[i];
        if (ix->fields->n == 0) { dao_raise(3264, "DAO.TableDefs", "No field defined--cannot append TableDef or Index."); r = 3264; break; }
        r = meta_insert_index(d, t->name, ix, i);
    }
    if (!r) r = build_table(d, t->name, 0);
    r = end(d, r);
    if (r) return r;
    t->appended = 1;
    for (int i = 0; i < t->fields->n; i++) ((dfld *)t->fields->items[i])->appended = 1;
    for (int i = 0; i < t->indexes->n; i++) ((dix *)t->indexes->items[i])->appended = 1;
    for (int i = 0; i < t->fields->n; i++) ((dfld *)t->fields->items[i])->ord = i;
    return 0;
}

int schema_drop_table(ddb *d, const char *name)
{
    char real[256];
    if (schema_has_table(d, name, real, sizeof real) != 1) {
        dao_raise(3265, "DAO.TableDefs", "Item not found in this collection.");
        return 3265;
    }
    int r = begin(d);
    if (r) return r;
    jbuf s = { 0 };
    jb_s(&s, "DROP TABLE "); jb_qid(&s, real, -1); jb_s(&s, ";");
    jb_s(&s, "DELETE FROM MSysLB_Fields WHERE tbl = "); jb_qstr(&s, real, -1); jb_s(&s, ";");
    jb_s(&s, "DELETE FROM MSysLB_Indexes WHERE tbl = "); jb_qstr(&s, real, -1); jb_s(&s, ";");
    jb_s(&s, "DELETE FROM MSysLB_Tables WHERE name = "); jb_qstr(&s, real, -1); jb_s(&s, ";");
    jb_s(&s, "DELETE FROM MSysLB_Relations WHERE tbl = "); jb_qstr(&s, real, -1);
    jb_s(&s, " OR ftbl = "); jb_qstr(&s, real, -1); jb_s(&s, ";");
    r = dao_exec(d, s.p);
    jb_free(&s);
    if (!r) r = regen_triggers(d);
    return end(d, r);
}

int schema_add_field(ddb *d, const char *table, dfld *f)
{
    sqlite3_stmt *st = prep(d, "SELECT count(*), (SELECT count(*) FROM MSysLB_Fields WHERE tbl = ?1 AND name = ?2) FROM MSysLB_Fields WHERE tbl = ?1");
    int n = 0, dup = 0;
    if (st) {
        sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, f->name, -1, SQLITE_STATIC);
        if (sqlite3_step(st) == SQLITE_ROW) { n = sqlite3_column_int(st, 0); dup = sqlite3_column_int(st, 1); }
        sqlite3_finalize(st);
    }
    if (dup) { dao_raise(3191, "DAO.Fields", "Cannot define field more than once."); return 3191; }
    int r = begin(d);
    if (r) return r;
    r = meta_insert_field(d, table, f, n);
    if (!r) r = build_table(d, table, 1);
    r = end(d, r);
    if (!r) { f->appended = 1; f->ord = n; }
    return r;
}

int schema_drop_field(ddb *d, const char *table, const char *field)
{
    sqlite3_stmt *st = prep(d, "SELECT 1 FROM MSysLB_Fields WHERE tbl = ?1 AND name = ?2");
    int there = 0;
    if (st) {
        sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, field, -1, SQLITE_STATIC);
        there = sqlite3_step(st) == SQLITE_ROW;
        sqlite3_finalize(st);
    }
    if (!there) { dao_raise(3265, "DAO.Fields", "Item not found in this collection."); return 3265; }
    int r = begin(d);
    if (r) return r;
    jbuf s = { 0 };
    jb_s(&s, "DELETE FROM MSysLB_Fields WHERE tbl = "); jb_qstr(&s, table, -1); jb_s(&s, " AND name = "); jb_qstr(&s, field, -1);
    jb_s(&s, "; DELETE FROM MSysLB_Indexes WHERE tbl = "); jb_qstr(&s, table, -1);
    jb_s(&s, " AND (';' || fields || ';') LIKE "); { jbuf p = { 0 }; jb_f(&p, "%%;_%s;%%", field); jb_qstr(&s, p.p, -1); jb_free(&p); }
    jb_s(&s, ";");
    r = dao_exec(d, s.p);
    jb_free(&s);
    /* the order again 0 .. n-1 */
    if (!r) {
        st = prep(d, "SELECT name FROM MSysLB_Fields WHERE tbl = ?1 ORDER BY ord");
        if (st) {
            sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
            int k = 0;
            jbuf u = { 0 };
            while (sqlite3_step(st) == SQLITE_ROW) {
                jb_f(&u, "UPDATE MSysLB_Fields SET ord = %d WHERE tbl = ", k++); jb_qstr(&u, table, -1);
                jb_s(&u, " AND name = "); jb_qstr(&u, col_text(st, 0), -1); jb_s(&u, ";");
            }
            sqlite3_finalize(st);
            if (u.n) r = dao_exec(d, u.p);
            jb_free(&u);
        }
    }
    if (!r) r = build_table(d, table, 1);
    return end(d, r);
}

int schema_create_index(ddb *d, const char *table, dix *ix)
{
    if (ix->fields->n == 0) { dao_raise(3264, "DAO.Indexes", "No field defined--cannot append TableDef or Index."); return 3264; }
    if (ix->primary) {
        sqlite3_stmt *st = prep(d, "SELECT 1 FROM MSysLB_Indexes WHERE tbl = ?1 AND isprimary = 1");
        if (st) {
            sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
            int has = sqlite3_step(st) == SQLITE_ROW;
            sqlite3_finalize(st);
            if (has) { dao_raise(3283, "DAO.Indexes", "Primary key already exists."); return 3283; }
        }
    }
    int n = 0;
    sqlite3_stmt *st = prep(d, "SELECT count(*) FROM MSysLB_Indexes WHERE tbl = ?1");
    if (st) { sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC); if (sqlite3_step(st) == SQLITE_ROW) n = sqlite3_column_int(st, 0); sqlite3_finalize(st); }
    int r = begin(d);
    if (r) return r;
    r = meta_insert_index(d, table, ix, n);
    if (!r && (ix->primary || ix->required)) r = build_table(d, table, 1);
    else if (!r) {
        mfield *fs; int nf; mindex *is; int ni;
        if (load_meta(d, table, &fs, &nf, &is, &ni) == 0) {
            jbuf s = { 0 };
            for (int i = 0; i < ni; i++) if (!strcasecmp(is[i].name, ix->name)) index_sql(table, &is[i], &s);
            if (s.n) r = dao_exec(d, s.p);
            jb_free(&s);
            free_meta(fs, nf, is, ni);
        }
    }
    r = end(d, r);
    if (!r) ix->appended = 1;
    return r;
}

int schema_drop_index(ddb *d, const char *table, const char *index)
{
    sqlite3_stmt *st = prep(d, "SELECT isprimary FROM MSysLB_Indexes WHERE tbl = ?1 AND name = ?2");
    int there = 0;
    if (st) {
        sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, index, -1, SQLITE_STATIC);
        there = sqlite3_step(st) == SQLITE_ROW;
        sqlite3_finalize(st);
    }
    if (!there) { dao_raise(3265, "DAO.Indexes", "Item not found in this collection."); return 3265; }
    int r = begin(d);
    if (r) return r;
    jbuf s = { 0 };
    jb_s(&s, "DELETE FROM MSysLB_Indexes WHERE tbl = "); jb_qstr(&s, table, -1); jb_s(&s, " AND name = "); jb_qstr(&s, index, -1);
    r = dao_exec(d, s.p);
    jb_free(&s);
    if (!r) r = build_table(d, table, 1);
    return end(d, r);
}

int schema_create_relation(ddb *d, drel *rl)
{
    char pt[256], ct[256];
    if (schema_has_table(d, rl->table, pt, sizeof pt) != 1 || schema_has_table(d, rl->ftable, ct, sizeof ct) != 1) {
        dao_raise(3265, "DAO.Relations", "Item not found in this collection.");
        return 3265;
    }
    if (rl->fields->n == 0) { dao_raise(3367, "DAO.Relations", "Cannot append. An object with that name already exists in the collection."); return 3367; }
    jbuf fl = { 0 };
    for (int i = 0; i < rl->fields->n; i++) {
        dfld *f = (dfld *)rl->fields->items[i];
        if (i) jb_s(&fl, ";");
        jb_s(&fl, f->name); jb_s(&fl, "="); jb_s(&fl, f->foreign && *f->foreign ? f->foreign : f->name);
    }
    int r = begin(d);
    if (r) { jb_free(&fl); return r; }
    sqlite3_stmt *st = prep(d, "INSERT INTO MSysLB_Relations (name, tbl, ftbl, fields, attrs) VALUES (?1, ?2, ?3, ?4, ?5)");
    if (st) {
        sqlite3_bind_text(st, 1, rl->name, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, pt, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 3, ct, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 4, fl.p, -1, SQLITE_STATIC);
        sqlite3_bind_int(st, 5, rl->attrs);
        int rc = sqlite3_step(st);
        sqlite3_finalize(st);
        if (rc == SQLITE_CONSTRAINT) { dao_raise(3012, "DAO.Relations", "Relationship '%s' already exists.", rl->name); r = 3012; }
        else if (rc != SQLITE_DONE) r = dao_raise_sqlite(d, rc, "DAO.Relations");
    }
    /* Jet makes an index on the foreign fields, named as the relation */
    if (!r) {
        dix *ix = index_new(0, rl->name);
        ix->foreign = 1;
        for (int i = 0; i < rl->fields->n; i++) {
            dfld *f = (dfld *)rl->fields->items[i];
            dfld *g = field_new(FK_INDEX, (vbobj)ix, f->foreign && *f->foreign ? f->foreign : f->name, 0, 0);
            coll_add(ix->fields, (vbobj)g);
            vb_orelease((vbobj)g);
        }
        int n = 0;
        st = prep(d, "SELECT count(*) FROM MSysLB_Indexes WHERE tbl = ?1");
        if (st) { sqlite3_bind_text(st, 1, ct, -1, SQLITE_STATIC); if (sqlite3_step(st) == SQLITE_ROW) n = sqlite3_column_int(st, 0); sqlite3_finalize(st); }
        r = meta_insert_index(d, ct, ix, n);
        vb_orelease((vbobj)ix);
        if (!r) {
            mindex mx = { (char *)rl->name, 0, 0, 0, 0, 0, 1 };
            jbuf f2 = { 0 };
            for (int i = 0; i < rl->fields->n; i++) {
                dfld *f = (dfld *)rl->fields->items[i];
                if (i) jb_s(&f2, ";");
                jb_s(&f2, "+"); jb_s(&f2, f->foreign && *f->foreign ? f->foreign : f->name);
            }
            mx.fields = f2.p;
            jbuf s = { 0 };
            index_sql(ct, &mx, &s);
            r = dao_exec(d, s.p);
            jb_free(&s); jb_free(&f2);
        }
    }
    if (!r) r = regen_triggers(d);
    jb_free(&fl);
    r = end(d, r);
    if (!r) rl->appended = 1;
    return r;
}

/* CREATE TABLE / INDEX, DROP, ALTER from SQL */
int schema_ddl(ddb *d, jsql *q)
{
    char real[256];
    switch (q->kind) {
    case J_CREATE_TABLE: {
        dtd *t = tabledef_new(q->table);
        for (int i = 0; i < q->ncols; i++) {
            jcol *c = &q->cols[i];
            dfld *f = field_new(FK_TABLE, (vbobj)t, c->name, c->type, c->size);
            f->attrs |= c->attrs;
            f->required = c->notnull;
            f->allowzero = 2;               /* SQL DDL: AllowZeroLength reads False, "" is allowed (no property) */
            coll_add(t->fields, (vbobj)f);
            vb_orelease((vbobj)f);
            if (c->pk || c->unique) {
                dix *ix = index_new((vbobj)t, c->cname ? c->cname : c->pk ? "PrimaryKey" : c->name);
                ix->primary = c->pk; ix->unique = 1;
                dfld *g = field_new(FK_INDEX, (vbobj)ix, c->name, 0, 0);
                coll_add(ix->fields, (vbobj)g); vb_orelease((vbobj)g);
                coll_add(t->indexes, (vbobj)ix); vb_orelease((vbobj)ix);
            }
        }
        for (int i = 0; i < q->ncons; i++) {
            jcons *c = &q->cons[i];
            if (c->kind == 3) continue;
            dix *ix = index_new((vbobj)t, c->name ? c->name : c->kind == 1 ? "PrimaryKey" : c->cols[0]);
            ix->primary = c->kind == 1; ix->unique = 1;
            for (int k = 0; k < c->ncols; k++) {
                dfld *g = field_new(FK_INDEX, (vbobj)ix, c->cols[k], 0, 0);
                g->attrs = c->desc[k] ? dbDescending : 0;
                coll_add(ix->fields, (vbobj)g); vb_orelease((vbobj)g);
            }
            coll_add(t->indexes, (vbobj)ix); vb_orelease((vbobj)ix);
        }
        int r = schema_create_table(d, t);
        if (r == 3010) { vb_err_clear_pending(); dao_raise(3010, "DAO.Database", "Table '%s' already exists.", q->table); }
        /* REFERENCES and FOREIGN KEY: relations */
        for (int i = 0; !r && i < q->ncols; i++) {
            jcol *c = &q->cols[i];
            if (!c->ref_table) continue;
            char nm[300]; snprintf(nm, sizeof nm, "%s%s", c->ref_table, q->table);
            drel *rl = relation_new(c->cname ? c->cname : nm, c->ref_table, q->table, 0);
            dfld *g = field_new(FK_REL, (vbobj)rl, c->ref_col ? c->ref_col : c->name, 0, 0);
            free(g->foreign); g->foreign = strdup(c->name);
            coll_add(rl->fields, (vbobj)g); vb_orelease((vbobj)g);
            r = schema_create_relation(d, rl);
            vb_orelease((vbobj)rl);
        }
        for (int i = 0; !r && i < q->ncons; i++) {
            jcons *c = &q->cons[i];
            if (c->kind != 3) continue;
            char nm[300]; snprintf(nm, sizeof nm, "%s%s", c->ref_table, q->table);
            drel *rl = relation_new(c->name ? c->name : nm, c->ref_table, q->table, 0);
            for (int k = 0; k < c->ncols; k++) {
                dfld *g = field_new(FK_REL, (vbobj)rl, k < c->nref ? c->ref_cols[k] : c->cols[k], 0, 0);
                free(g->foreign); g->foreign = strdup(c->cols[k]);
                coll_add(rl->fields, (vbobj)g); vb_orelease((vbobj)g);
            }
            r = schema_create_relation(d, rl);
            vb_orelease((vbobj)rl);
        }
        vb_orelease((vbobj)t);
        return r;
    }
    case J_CREATE_INDEX: {
        if (schema_has_table(d, q->table, real, sizeof real) != 1) { dao_raise(3371, "DAO.Database", "Cannot find table or constraint."); return 3371; }
        dix *ix = index_new(0, q->index);
        ix->unique = q->unique; ix->primary = q->primary; ix->required = q->disallow_null; ix->ignorenulls = q->ignore_null;
        for (int i = 0; i < q->ncols; i++) {
            dfld *g = field_new(FK_INDEX, (vbobj)ix, q->cols[i].name, 0, 0);
            g->attrs = q->cols[i].attrs;
            coll_add(ix->fields, (vbobj)g); vb_orelease((vbobj)g);
        }
        int r = schema_create_index(d, real, ix);
        vb_orelease((vbobj)ix);
        return r;
    }
    case J_DROP_TABLE:
        if (schema_has_table(d, q->table, real, sizeof real) != 1) {
            dao_raise(3376, "DAO.Database", "Table '%s' does not exist.", q->table);
            return 3376;
        }
        return schema_drop_table(d, real);
    case J_DROP_INDEX:
        if (schema_has_table(d, q->table, real, sizeof real) != 1) { dao_raise(3371, "DAO.Database", "Cannot find table or constraint."); return 3371; }
        return schema_drop_index(d, real, q->index);
    case J_ALTER_ADD: {
        if (schema_has_table(d, q->table, real, sizeof real) != 1) { dao_raise(3371, "DAO.Database", "Cannot find table or constraint."); return 3371; }
        jcol *c = &q->cols[0];
        dfld *f = field_new(FK_TABLE, 0, c->name, c->type, c->size);
        f->attrs |= c->attrs; f->required = c->notnull; f->allowzero = 2;
        int r;
        if (q->unique == -1) {                  /* ALTER COLUMN */
            jbuf s = { 0 };
            jb_s(&s, "UPDATE MSysLB_Fields SET type = "); jb_f(&s, "%d, size = %d, required = %d WHERE tbl = ", c->type, f->size, c->notnull);
            jb_qstr(&s, real, -1); jb_s(&s, " AND name = "); jb_qstr(&s, c->name, -1);
            r = begin(d);
            if (!r) r = dao_exec(d, s.p);
            if (!r) r = build_table(d, real, 1);
            r = end(d, r);
            jb_free(&s);
        } else r = schema_add_field(d, real, f);
        vb_orelease((vbobj)f);
        return r;
    }
    case J_ALTER_DROP:
        if (schema_has_table(d, q->table, real, sizeof real) != 1) { dao_raise(3371, "DAO.Database", "Cannot find table or constraint."); return 3371; }
        return schema_drop_field(d, real, q->cols[0].name);
    case J_ALTER_CONSTRAINT: {
        if (schema_has_table(d, q->table, real, sizeof real) != 1) { dao_raise(3371, "DAO.Database", "Cannot find table or constraint."); return 3371; }
        jcons *c = &q->cons[0];
        if (c->kind == 3) {
            char nm[300]; snprintf(nm, sizeof nm, "%s%s", c->ref_table, q->table);
            drel *rl = relation_new(c->name ? c->name : nm, c->ref_table, real, 0);
            for (int k = 0; k < c->ncols; k++) {
                dfld *g = field_new(FK_REL, (vbobj)rl, k < c->nref ? c->ref_cols[k] : c->cols[k], 0, 0);
                free(g->foreign); g->foreign = strdup(c->cols[k]);
                coll_add(rl->fields, (vbobj)g); vb_orelease((vbobj)g);
            }
            int r = schema_create_relation(d, rl);
            vb_orelease((vbobj)rl);
            return r;
        }
        dix *ix = index_new(0, c->name ? c->name : "PrimaryKey");
        ix->primary = c->kind == 1; ix->unique = 1;
        for (int k = 0; k < c->ncols; k++) {
            dfld *g = field_new(FK_INDEX, (vbobj)ix, c->cols[k], 0, 0);
            coll_add(ix->fields, (vbobj)g); vb_orelease((vbobj)g);
        }
        int r = schema_create_index(d, real, ix);
        vb_orelease((vbobj)ix);
        return r;
    }
    case J_DROP_CONSTRAINT: {
        if (schema_has_table(d, q->table, real, sizeof real) != 1) { dao_raise(3371, "DAO.Database", "Cannot find table or constraint."); return 3371; }
        sqlite3_stmt *st = prep(d, "SELECT 1 FROM MSysLB_Relations WHERE name = ?1");
        int rel = 0;
        if (st) { sqlite3_bind_text(st, 1, q->index, -1, SQLITE_STATIC); rel = sqlite3_step(st) == SQLITE_ROW; sqlite3_finalize(st); }
        if (rel) {
            jbuf s = { 0 };
            jb_s(&s, "DELETE FROM MSysLB_Relations WHERE name = "); jb_qstr(&s, q->index, -1);
            int r = begin(d);
            if (!r) r = dao_exec(d, s.p);
            if (!r) r = schema_drop_index(d, real, q->index);
            if (!r) r = regen_triggers(d);
            jb_free(&s);
            return end(d, r);
        }
        return schema_drop_index(d, real, q->index);
    }
    }
    return 0;
}

/* ---------------------------------------------------------------- the lists */
static int cmp_name(const void *a, const void *b)
{
    return strcasecmp(obj_name(*(vbobj const *)a), obj_name(*(vbobj const *)b));
}
void schema_list_tables(ddb *d, dcoll *c)
{
    static const char *const SYS[] = { "MSysACEs", "MSysObjects", "MSysQueries", "MSysRelationships" };
    for (int i = 0; i < 4; i++) {
        dtd *t = tabledef_new(SYS[i]);
        t->db = (vbobj)d; vb_oaddref((vbobj)d);
        t->attrs = (int)0x80000000;
        t->appended = 1;
        coll_add(c, (vbobj)t);
        vb_orelease((vbobj)t);
    }
    sqlite3_stmt *st = prep(d, "SELECT name FROM sqlite_master WHERE type = 'table'");
    if (st) {
        while (sqlite3_step(st) == SQLITE_ROW) {
            const char *n = col_text(st, 0);
            if (hidden_table(n)) continue;
            dtd *t = schema_load_table(d, n);
            if (t) { coll_add(c, (vbobj)t); vb_orelease((vbobj)t); }
        }
        sqlite3_finalize(st);
    }
    qsort(c->items, (size_t)c->n, sizeof(vbobj), cmp_name);
}

static void implicit_params(dqd *q);
static void declared_params(dqd *q);
void querydef_parse_params(dqd *q)
{
    coll_clear_items(q->params);
    declared_params(q);
    implicit_params(q);
}
/* Forms!Dialog!Date in a query without PARAMETERS: Jet makes it a parameter too */
static void implicit_params(dqd *q)
{
    const char *s = q->sql;
    char qc = 0;
    for (const char *p = s; *p; p++) {
        if (qc) { if (*p == qc) qc = 0; continue; }
        if (*p == '\'' || *p == '"') { qc = *p; continue; }
        if (*p != '!') continue;
        /* the chain around the ! */
        const char *b = p, *e = p;
        for (;;) {                              /* back over [x] or x */
            const char *x = b;
            if (x > s && x[-1] == ']') { x--; while (x > s && *x != '[') x--; }
            else while (x > s && (isalnum((unsigned char)x[-1]) || x[-1] == '_')) x--;
            if (x == b) break;
            b = x;
            if (b > s && b[-1] == '!') { b--; continue; }
            break;
        }
        for (;;) {                              /* forward */
            const char *x = e + 1;
            if (*x == '[') { while (*x && *x != ']') x++; if (*x) x++; }
            else while (isalnum((unsigned char)*x) || *x == '_') x++;
            if (x == e + 1) break;
            e = x;
            if (*e == '!') continue;
            break;
        }
        char nm[256]; snprintf(nm, sizeof nm, "%.*s", (int)(e - b), b);
        p = e - 1;
        if (*b == '!' || !*nm) continue;
        char plain[256]; dao_param_plain(nm, plain, sizeof plain);
        if (coll_find(q->params, plain) >= 0) continue;
        dpar *pr = param_new(nm, dbText);
        coll_add(q->params, (vbobj)pr);
        vb_orelease((vbobj)pr);
    }
}
static void declared_params(dqd *q)
{
    const char *s = q->sql;
    while (*s && isspace((unsigned char)*s)) s++;
    if (strncasecmp(s, "parameters", 10) || !isspace((unsigned char)s[10])) return;
    jsql j; memset(&j, 0, sizeof j);
    /* the declarations only: the text up to the first ; */
    const char *semi = strchr(s, ';');
    if (!semi) return;
    jbuf t = { 0 };
    jb_put(&t, s, (int)(semi - s + 1));
    jb_s(&t, " SELECT 1");
    if (jet_translate(t.p, &j, 0) == 0) {
        /* the names as written: [Min ID] keeps its brackets */
        const char *p = s + 10;
        for (int i = 0; i < j.nparams; i++) {
            while (*p && isspace((unsigned char)*p)) p++;
            const char *b = p;
            for (;;) {                          /* [Forms]![Dialog]![Date]: one name */
                if (*p == '[') { while (*p && *p != ']') p++; if (*p) p++; }
                else while (*p && !isspace((unsigned char)*p) && *p != ',' && *p != ';' && *p != '!' && *p != '[') p++;
                if (*p == '!' || *p == '.') { p++; continue; }
                break;
            }
            char nm[256]; snprintf(nm, sizeof nm, "%.*s", (int)(p - b), b);
            dpar *pr = param_new(nm, j.params[i].type);
            coll_add(q->params, (vbobj)pr);
            vb_orelease((vbobj)pr);
            while (*p && *p != ',' && *p != ';') p++;
            if (*p == ',') p++;
        }
    }
    jsql_free(&j);
    jb_free(&t);
}

void dao_querydef_init(vbobj o)
{
    dqd *q = (dqd *)o;
    q->name = strdup(""); q->sql = strdup("");
    q->params = coll_new(C_PARAMS, o);
    q->params->loaded = 1;
    q->connect = strdup("");
    q->created = q->updated = now_date();
    q->odbctimeout = 60;
}
dqd *querydef_new(vbobj db, const char *name, const char *sql)
{
    dqd *q = (dqd *)vb_onew(&vb_cls_dao_querydef);
    vb_oaddref((vbobj)q);
    if (db) { q->db = db; vb_oaddref(db); }
    free(q->name); q->name = strdup(name ? name : "");
    free(q->sql); q->sql = strdup(sql ? sql : "");
    querydef_parse_params(q);
    return q;
}

void schema_list_queries(ddb *d, dcoll *c)
{
    sqlite3_stmt *st = prep(d, "SELECT name, sql, type FROM MSysLB_Queries ORDER BY name COLLATE NOCASE");
    if (!st) return;
    while (sqlite3_step(st) == SQLITE_ROW) {
        dqd *q = querydef_new((vbobj)d, col_text(st, 0), col_text(st, 1));
        q->type = sqlite3_column_int(st, 2);
        q->appended = 1;
        coll_add(c, (vbobj)q);
        vb_orelease((vbobj)q);
    }
    sqlite3_finalize(st);
}

void schema_list_relations(ddb *d, dcoll *c)
{
    sqlite3_stmt *st = prep(d, "SELECT name, tbl, ftbl, fields, attrs FROM MSysLB_Relations ORDER BY name COLLATE NOCASE");
    if (!st) return;
    while (sqlite3_step(st) == SQLITE_ROW) {
        drel *r = relation_new(col_text(st, 0), col_text(st, 1), col_text(st, 2), sqlite3_column_int(st, 4));
        r->db = (vbobj)d; vb_oaddref((vbobj)d);
        r->appended = 1;
        char *fl = strdup(col_text(st, 3)), *save = 0;
        for (char *p = strtok_r(fl, ";", &save); p; p = strtok_r(0, ";", &save)) {
            char *eq = strchr(p, '=');
            if (eq) *eq = 0;
            dfld *f = field_new(FK_REL, (vbobj)r, p, 0, 0);
            free(f->foreign); f->foreign = strdup(eq ? eq + 1 : p);
            f->appended = 1;
            coll_add(r->fields, (vbobj)f);
            vb_orelease((vbobj)f);
        }
        free(fl);
        coll_add(c, (vbobj)r);
        vb_orelease((vbobj)r);
    }
    sqlite3_finalize(st);
}

/* ---------------------------------------------------------------- counters */
static sqlite3_int64 seq_of(ddb *d, const char *table)
{
    sqlite3_int64 n = 0;
    sqlite3_stmt *st = prep(d, "SELECT seq FROM sqlite_sequence WHERE name = ?1");
    if (st) { sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC); if (sqlite3_step(st) == SQLITE_ROW) n = sqlite3_column_int64(st, 0); sqlite3_finalize(st); }
    return n;
}
static void seq_set(ddb *d, const char *table, sqlite3_int64 n)
{
    sqlite3_stmt *st = prep(d, "UPDATE sqlite_sequence SET seq = ?2 WHERE name = ?1");
    if (!st) return;
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    sqlite3_bind_int64(st, 2, n);
    sqlite3_step(st);
    sqlite3_finalize(st);
    if (sqlite3_changes(d->db)) return;
    st = prep(d, "INSERT INTO sqlite_sequence (name, seq) VALUES (?1, ?2)");
    if (!st) return;
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    sqlite3_bind_int64(st, 2, n);
    sqlite3_step(st);
    sqlite3_finalize(st);
}
/* the next counter value of a table, taken now (Jet: AddNew and a failed INSERT use one up) */
int schema_counter_next(ddb *d, const char *table, sqlite3_int64 *next)
{
    sqlite3_stmt *st = prep(d, "SELECT 1 FROM MSysLB_Fields WHERE tbl = ?1 AND (attrs & 16) <> 0");
    if (!st) return 0;
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    int has = sqlite3_step(st) == SQLITE_ROW;
    sqlite3_finalize(st);
    if (!has) return 0;
    sqlite3_int64 n = seq_of(d, table), mx = 0;
    jbuf s = { 0 };
    jb_s(&s, "SELECT coalesce(max(rowid), 0) FROM "); jb_qid(&s, table, -1);
    st = prep(d, s.p);
    jb_free(&s);
    if (st) { if (sqlite3_step(st) == SQLITE_ROW) mx = sqlite3_column_int64(st, 0); sqlite3_finalize(st); }
    *next = (n > mx ? n : mx) + 1;
    seq_set(d, table, *next);
    return 1;
}
void schema_counter_bump(ddb *d, const char *table)
{
    sqlite3_int64 n;
    schema_counter_next(d, table, &n);
}

/* ---------------------------------------------------------------- saved queries */
/* Jet keeps the text of a query with a line break before each clause and a ; at the end */
char *jet_pretty_sql(const char *sql)
{
    static const char *const KW[] = { "FROM", "WHERE", "GROUP BY", "HAVING", "ORDER BY", "UNION", 0 };
    jbuf b = { 0 };
    const char *s = sql;
    while (*s && isspace((unsigned char)*s)) s++;
    int depth = 0;
    char q = 0;
    const char *start = s;
    /* PARAMETERS ...; then the statement */
    if (!strncasecmp(s, "parameters", 10)) {
        const char *semi = strchr(s, ';');
        if (semi) {
            jb_put(&b, s, (int)(semi - s + 1));
            jb_s(&b, "\r\n");
            s = semi + 1;
            while (*s && isspace((unsigned char)*s)) s++;
            start = s;
        }
    }
    for (const char *p = s; *p; p++) {
        char c = *p;
        if (q) { if (c == q) q = 0; continue; }
        if (c == '\'' || c == '"') { q = c; continue; }
        if (c == '[') { while (*p && *p != ']') p++; if (!*p) break; continue; }
        if (c == '(') depth++;
        else if (c == ')') depth--;
        if (depth || p == start || !isspace((unsigned char)c)) continue;
        for (int k = 0; KW[k]; k++) {
            size_t n = strlen(KW[k]);
            const char *w = p + 1;
            while (*w == ' ') w++;
            if (!strncasecmp(w, KW[k], n) && !isalnum((unsigned char)w[n]) && w[n] != '_') {
                /* the text so far, without its trailing spaces */
                const char *e = p;
                while (e > s && isspace((unsigned char)e[-1])) e--;
                jb_put(&b, s, (int)(e - s));
                jb_s(&b, "\r\n");
                s = w;
                p = w - 1;
                break;
            }
        }
    }
    const char *e = s + strlen(s);
    while (e > s && (isspace((unsigned char)e[-1]) || e[-1] == ';')) e--;
    jb_put(&b, s, (int)(e - s));
    jb_s(&b, ";\r\n");
    return b.p;
}

/* a SELECT query is also a SQLite view, so that a query can use it by name */
static void query_view(ddb *d, const char *name, const char *sql)
{
    jbuf s = { 0 };
    jb_s(&s, "DROP VIEW IF EXISTS "); jb_qid(&s, name, -1);
    sqlite3_exec(d->db, s.p, 0, 0, 0);
    jb_free(&s);
    jsql j;
    if (jet_translate(sql, &j, 0) == 0 && j.kind == J_SELECT && j.returns && j.nbinds == 0 && j.nparams == 0) {
        jb_s(&s, "CREATE VIEW "); jb_qid(&s, name, -1); jb_s(&s, " AS "); jb_s(&s, j.sql);
        sqlite3_exec(d->db, s.p, 0, 0, 0);
        jb_free(&s);
    }
    jsql_free(&j);
}

int schema_save_query(ddb *d, dqd *q, int replace)
{
    sqlite3_stmt *st = prep(d, replace ? "UPDATE MSysLB_Queries SET sql = ?2, updated = ?4 WHERE name = ?1"
                                       : "INSERT INTO MSysLB_Queries (name, sql, type, created, updated) VALUES (?1, ?2, ?3, ?4, ?4)");
    if (!st) return dao_raise_sqlite(d, sqlite3_errcode(d->db), "DAO.QueryDefs");
    char t[24]; date_text(now_date(), t);
    sqlite3_bind_text(st, 1, q->name, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 2, q->sql, -1, SQLITE_STATIC);
    sqlite3_bind_int(st, 3, q->type);
    sqlite3_bind_text(st, 4, t, -1, SQLITE_STATIC);
    int rc = sqlite3_step(st);
    sqlite3_finalize(st);
    if (rc == SQLITE_CONSTRAINT) { dao_raise(3012, "DAO.QueryDefs", "Object '%s' already exists.", q->name); return 3012; }
    if (rc != SQLITE_DONE) return dao_raise_sqlite(d, rc, "DAO.QueryDefs");
    query_view(d, q->name, q->sql);
    return 0;
}
int schema_drop_query(ddb *d, const char *name)
{
    sqlite3_stmt *st = prep(d, "DELETE FROM MSysLB_Queries WHERE name = ?1");
    if (!st) return dao_raise_sqlite(d, sqlite3_errcode(d->db), "DAO.QueryDefs");
    sqlite3_bind_text(st, 1, name, -1, SQLITE_STATIC);
    sqlite3_step(st);
    sqlite3_finalize(st);
    if (!sqlite3_changes(d->db)) { dao_raise(3265, "DAO.QueryDefs", "Item not found in this collection."); return 3265; }
    jbuf s = { 0 };
    jb_s(&s, "DROP VIEW IF EXISTS "); jb_qid(&s, name, -1);
    sqlite3_exec(d->db, s.p, 0, 0, 0);
    jb_free(&s);
    return 0;
}

void schema_after_change(ddb *d) { (void)d; }

/* after an import: every table of the schema made, then the triggers and the views */
int schema_build_all(ddb *d)
{
    sqlite3_stmt *st = prep(d, "SELECT name FROM MSysLB_Tables");
    if (!st) return -1;
    char **names = 0; int n = 0;
    while (sqlite3_step(st) == SQLITE_ROW) { names = realloc(names, sizeof(char *) * (size_t)(n + 1)); names[n++] = strdup(col_text(st, 0)); }
    sqlite3_finalize(st);
    int r = 0;
    for (int i = 0; i < n; i++) {
        jbuf sql = { 0 };
        if (!r && make_table_sql(d, names[i], names[i], &sql) == 0) { jb_s(&sql, ";"); r = dao_exec(d, sql.p); }
        sql.n = 0;
        mfield *fs; int nf; mindex *is; int ni;
        if (!r && load_meta(d, names[i], &fs, &nf, &is, &ni) == 0) {
            for (int k = 0; k < ni; k++) { index_sql(names[i], &is[k], &sql); jb_s(&sql, ";"); }
            free_meta(fs, nf, is, ni);
            if (sql.n) r = dao_exec(d, sql.p);
        }
        jb_free(&sql);
        free(names[i]);
    }
    free(names);
    return r;
}
int schema_after_import(ddb *d)
{
    int r = regen_triggers(d);
    sqlite3_stmt *st = prep(d, "SELECT name, sql FROM MSysLB_Queries");
    if (!st) return r;
    while (sqlite3_step(st) == SQLITE_ROW) query_view(d, col_text(st, 0), col_text(st, 1));
    sqlite3_finalize(st);
    return r;
}
