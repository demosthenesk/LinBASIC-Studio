/* mdbimport.c - a Jet .mdb file (Access 97 / 2000, Jet 3 and 4) copied into a LinBASIC (SQLite)
 * database: the tables with their field types, sizes and properties, the rows, the indexes, the
 * relations (MSysRelationships) and the saved queries (MSysQueries, written again as Jet SQL text as
 * DAO shows it). The file is read with libmdb of mdbtools (LGPL, runtime/db/mdb); the .mdb itself is
 * never written: the program's changes go to the SQLite copy.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "daoint.h"
#include "mdb/mdbtools.h"
#include <ctype.h>
#include <sys/stat.h>

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

static int dao_type(int mdbtype)
{
    switch (mdbtype) {
    case MDB_BOOL: return dbBoolean;
    case MDB_BYTE: return dbByte;
    case MDB_INT: return dbInteger;
    case MDB_LONGINT: return dbLong;
    case MDB_MONEY: return dbCurrency;
    case MDB_FLOAT: return dbSingle;
    case MDB_DOUBLE: return dbDouble;
    case MDB_DATETIME: return dbDate;
    case MDB_BINARY: return dbBinary;
    case MDB_TEXT: return dbText;
    case MDB_OLE: return dbLongBinary;
    case MDB_MEMO: return dbMemo;
    case MDB_REPID: return dbGUID;
    case MDB_NUMERIC: return dbDecimal;
    }
    return dbLongBinary;
}

static int prop_yes(MdbColumn *c, const char *name)
{
    const char *v = mdb_col_get_prop(c, name);
    return v && (!strcmp(v, "yes") || !strcmp(v, "1") || !strcasecmp(v, "true"));
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

/* ---------------------------------------------------------------- tables */
static int import_schema(MdbHandle *mdb, MdbCatalogEntry *e, sqlite3 *db, char *err, int cap)
{
    MdbTableDef *t = mdb_read_table(e);
    if (!t) return 0;
    mdb_read_columns(t);
    mdb_read_indices(t);
    sqlite3_stmt *st;
    sqlite3_prepare_v2(db, "INSERT INTO MSysLB_Tables (name, attrs, seq) VALUES (?1, 0, 0)", -1, &st, 0);
    sqlite3_bind_text(st, 1, e->object_name, -1, SQLITE_STATIC);
    sqlite3_step(st);
    sqlite3_finalize(st);
    sqlite3_prepare_v2(db, "INSERT INTO MSysLB_Fields (tbl, name, ord, type, size, attrs, required, allowzero, defval, vrule, vtext) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11)", -1, &st, 0);
    for (unsigned i = 0; i < t->num_cols; i++) {
        MdbColumn *c = g_ptr_array_index(t->columns, i);
        int type = dao_type(c->col_type), size, attrs;
        switch (type) {
        case dbText: size = IS_JET3(mdb) ? c->col_size : c->col_size / 2; attrs = dbVariableField; break;
        case dbMemo: case dbLongBinary: size = 0; attrs = dbVariableField; break;
        case dbBinary: size = c->col_size; attrs = dbVariableField; break;
        case dbBoolean: case dbByte: size = 1; attrs = dbFixedField; break;
        case dbInteger: size = 2; attrs = dbFixedField; break;
        case dbLong: case dbSingle: size = 4; attrs = dbFixedField; break;
        case dbCurrency: case dbDouble: case dbDate: size = 8; attrs = dbFixedField; break;
        case dbGUID: case dbDecimal: size = 16; attrs = dbFixedField; break;
        default: size = c->col_size; attrs = dbFixedField;
        }
        if (c->is_long_auto) attrs |= dbAutoIncrField;
        if (type == dbMemo && (c->col_flags & 0x80)) attrs |= dbHyperlinkField;
        const char *dv = mdb_col_get_prop(c, "DefaultValue"), *vr = mdb_col_get_prop(c, "ValidationRule"), *vt = mdb_col_get_prop(c, "ValidationText");
        sqlite3_bind_text(st, 1, e->object_name, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, c->name, -1, SQLITE_STATIC);
        sqlite3_bind_int(st, 3, c->row_col_num);
        sqlite3_bind_int(st, 4, type);
        sqlite3_bind_int(st, 5, size);
        sqlite3_bind_int(st, 6, attrs);
        sqlite3_bind_int(st, 7, prop_yes(c, "Required"));
        sqlite3_bind_int(st, 8, prop_yes(c, "AllowZeroLength"));
        sqlite3_bind_text(st, 9, dv ? dv : "", -1, SQLITE_TRANSIENT);
        sqlite3_bind_text(st, 10, vr ? vr : "", -1, SQLITE_TRANSIENT);
        sqlite3_bind_text(st, 11, vt ? vt : "", -1, SQLITE_TRANSIENT);
        sqlite3_step(st);
        sqlite3_reset(st);
    }
    sqlite3_finalize(st);
    sqlite3_prepare_v2(db, "INSERT OR IGNORE INTO MSysLB_Indexes (tbl, name, fields, isprimary, isunique, required, ignorenulls, isforeign, ord) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9)", -1, &st, 0);
    for (unsigned i = 0; t->indices && i < t->num_idxs; i++) {
        MdbIndex *x = g_ptr_array_index(t->indices, i);
        if (x->name[0] == '.') continue;        /* the hidden side of a relation (.rB ...) */
        MdbIndex *keys = x;
        if (!x->num_keys)                       /* a foreign key: the columns of the real index it uses */
            for (unsigned j = 0; j < t->num_idxs; j++) {
                MdbIndex *y = g_ptr_array_index(t->indices, j);
                if (y != x && y->index_num == x->index_num && y->num_keys) { keys = y; break; }
            }
        jbuf fl = { 0 };
        for (unsigned k = 0; k < keys->num_keys; k++) {
            int cn = keys->key_col_num[k] - 1;
            if (cn < 0 || cn >= (int)t->num_cols) continue;
            MdbColumn *c = g_ptr_array_index(t->columns, cn);
            if (fl.n) jb_s(&fl, ";");
            jb_s(&fl, keys->key_col_order[k] == MDB_DESC ? "-" : "+");
            jb_s(&fl, c->name);
        }
        if (!fl.n) { jb_free(&fl); continue; }
        int primary = x->index_type == 1, foreign = x->index_type == 2;
        sqlite3_bind_text(st, 1, e->object_name, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, x->name, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 3, fl.p ? fl.p : "", -1, SQLITE_TRANSIENT);
        sqlite3_bind_int(st, 4, primary);
        sqlite3_bind_int(st, 5, primary || (!foreign && (keys->flags & MDB_IDX_UNIQUE)) ? 1 : 0);
        sqlite3_bind_int(st, 6, primary || (!foreign && (keys->flags & MDB_IDX_REQUIRED)) ? 1 : 0);
        sqlite3_bind_int(st, 7, !foreign && (keys->flags & MDB_IDX_IGNORENULLS) ? 1 : 0);
        sqlite3_bind_int(st, 8, foreign);
        sqlite3_bind_int(st, 9, (int)i);
        sqlite3_step(st);
        sqlite3_reset(st);
        jb_free(&fl);
    }
    sqlite3_finalize(st);
    mdb_free_tabledef(t);
    (void)err; (void)cap;
    return 0;
}

static void money_text(const unsigned char *b, double *out)
{
    long long v = 0;
    for (int i = 7; i >= 0; i--) v = (v << 8) | b[i];
    *out = (double)v / 10000.0;
}

static int import_rows(MdbHandle *mdb, MdbCatalogEntry *e, sqlite3 *db, char *err, int cap)
{
    MdbTableDef *t = mdb_read_table(e);
    if (!t) return 0;
    mdb_read_columns(t);
    int nc = (int)t->num_cols;
    char **buf = calloc((size_t)(nc ? nc : 1), sizeof(char *));
    int *len = calloc((size_t)(nc ? nc : 1), sizeof(int));
    int *req = calloc((size_t)(nc ? nc : 1), sizeof(int));
    jbuf s = { 0 };
    jb_s(&s, "INSERT INTO "); jb_qid(&s, e->object_name, -1); jb_s(&s, " (");
    for (int i = 0; i < nc; i++) {
        MdbColumn *c = g_ptr_array_index(t->columns, i);
        buf[i] = malloc(mdb->bind_size + 1);
        mdb_bind_column(t, i + 1, buf[i], &len[i]);
        req[i] = prop_yes(c, "Required");
        if (i) jb_s(&s, ", ");
        jb_qid(&s, c->name, -1);
    }
    jb_s(&s, ") VALUES (");
    for (int i = 0; i < nc; i++) jb_f(&s, "%s?%d", i ? ", " : "", i + 1);
    jb_s(&s, ")");
    sqlite3_stmt *st = 0;
    int rc = sqlite3_prepare_v2(db, s.p, -1, &st, 0);
    jb_free(&s);
    if (rc != SQLITE_OK) {
        snprintf(err, (size_t)cap, "%s", sqlite3_errmsg(db));
        for (int i = 0; i < nc; i++) free(buf[i]);
        free(buf); free(len); free(req);
        mdb_free_tabledef(t);
        return -1;
    }
    mdb_rewind_table(t);
    int r = 0;
    while (mdb_fetch_row(t)) {
        for (int i = 0; i < nc; i++) {
            MdbColumn *c = g_ptr_array_index(t->columns, i);
            int k = i + 1;
            if (c->col_type == MDB_BOOL) { sqlite3_bind_int(st, k, !strcmp(buf[i], "1") ? -1 : 0); continue; }
            if (c->cur_value_is_null || (!c->cur_value_len && c->col_type != MDB_TEXT && c->col_type != MDB_MEMO)) {
                if (req[i] && (c->col_type == MDB_TEXT || c->col_type == MDB_MEMO)) sqlite3_bind_text(st, k, "", 0, SQLITE_STATIC);
                else sqlite3_bind_null(st, k);
                continue;
            }
            int start = c->cur_value_start;
            unsigned char *pg = (unsigned char *)mdb->pg_buf;
            switch (c->col_type) {
            case MDB_BYTE: sqlite3_bind_int(st, k, pg[start]); break;
            case MDB_INT: sqlite3_bind_int(st, k, (short)mdb_get_int16(pg, start)); break;
            case MDB_LONGINT: sqlite3_bind_int(st, k, (int)mdb_get_int32(pg, start)); break;
            case MDB_FLOAT: sqlite3_bind_double(st, k, (double)mdb_get_single(pg, start)); break;
            case MDB_DOUBLE: sqlite3_bind_double(st, k, mdb_get_double(pg, start)); break;
            case MDB_MONEY: { double d; money_text(pg + start, &d); sqlite3_bind_double(st, k, d); break; }
            case MDB_DATETIME: {
                char tx[24];
                jet_date_text(mdb_get_double(pg, start), tx);
                sqlite3_bind_text(st, k, tx, -1, SQLITE_TRANSIENT);
                break; }
            case MDB_OLE: {
                size_t size = 0;
                void *p = mdb_ole_read_full(mdb, c, &size);
                if (p) { sqlite3_bind_blob(st, k, p, (int)size, SQLITE_TRANSIENT); free(p); }
                else sqlite3_bind_null(st, k);
                break; }
            case MDB_NUMERIC: {
                char *q = buf[i];
                for (char *z = q; *z; z++) if (*z == ',') *z = '.';
                sqlite3_bind_double(st, k, atof(q));
                break; }
            default:
                if (!len[i] && !req[i] && c->cur_value_is_null) sqlite3_bind_null(st, k);
                else sqlite3_bind_text(st, k, buf[i], len[i], SQLITE_TRANSIENT);
            }
        }
        rc = sqlite3_step(st);
        sqlite3_reset(st);
        if (rc != SQLITE_DONE && (rc & 0xff) != SQLITE_CONSTRAINT) {
            snprintf(err, (size_t)cap, "%s: %s", e->object_name, sqlite3_errmsg(db));
            r = -1;
            break;
        }
    }
    sqlite3_finalize(st);
    for (int i = 0; i < nc; i++) free(buf[i]);
    free(buf); free(len); free(req);
    mdb_free_tabledef(t);
    return r;
}

/* ---------------------------------------------------------------- a system table as rows of text */
typedef struct strows { int ncols; char **names; char ***rows; int n; } strows;

static MdbCatalogEntry *catalog_entry(MdbHandle *mdb, const char *name)
{
    for (unsigned i = 0; i < mdb->num_catalog; i++) {
        MdbCatalogEntry *e = g_ptr_array_index(mdb->catalog, i);
        if (!strcmp(e->object_name, name)) return e;
    }
    return 0;
}
static int read_rows(MdbHandle *mdb, const char *table, const char *const *cols, strows *out)
{
    memset(out, 0, sizeof *out);
    MdbCatalogEntry *e = catalog_entry(mdb, table);
    if (!e) return -1;
    MdbTableDef *t = mdb_read_table(e);
    if (!t) return -1;
    mdb_read_columns(t);
    int nc = 0;
    while (cols[nc]) nc++;
    char **buf = calloc((size_t)nc, sizeof(char *));
    int *isnull = calloc((size_t)nc, sizeof(int));
    int *found = calloc((size_t)nc, sizeof(int));
    for (int i = 0; i < nc; i++) {
        buf[i] = calloc(1, mdb->bind_size + 1);
        found[i] = mdb_bind_column_by_name(t, (gchar *)cols[i], buf[i], 0) > 0;
    }
    mdb_rewind_table(t);
    out->ncols = nc;
    while (mdb_fetch_row(t)) {
        out->rows = realloc(out->rows, sizeof(char **) * (size_t)(out->n + 1));
        char **row = calloc((size_t)nc, sizeof(char *));
        for (int i = 0; i < nc; i++) row[i] = strdup(found[i] ? buf[i] : "");
        out->rows[out->n++] = row;
    }
    for (int i = 0; i < nc; i++) free(buf[i]);
    free(buf); free(isnull); free(found);
    mdb_free_tabledef(t);
    return 0;
}
static void free_rows(strows *r)
{
    for (int i = 0; i < r->n; i++) { for (int k = 0; k < r->ncols; k++) free(r->rows[i][k]); free(r->rows[i]); }
    free(r->rows);
}

/* ---------------------------------------------------------------- relations */
static void import_relations(MdbHandle *mdb, sqlite3 *db)
{
    static const char *const C[] = { "ccolumn", "grbit", "icolumn", "szColumn", "szObject", "szReferencedColumn", "szReferencedObject", "szRelationship", 0 };
    strows r;
    if (read_rows(mdb, "MSysRelationships", C, &r)) return;
    sqlite3_stmt *st = 0;
    sqlite3_prepare_v2(db, "INSERT OR IGNORE INTO MSysLB_Relations (name, tbl, ftbl, fields, attrs) VALUES (?1, ?2, ?3, ?4, ?5)", -1, &st, 0);
    char **done = 0; int nd = 0;
    for (int i = 0; i < r.n; i++) {
        const char *name = r.rows[i][7];
        int seen = 0;
        for (int k = 0; k < nd; k++) if (!strcmp(done[k], name)) seen = 1;
        if (seen) continue;
        done = realloc(done, sizeof(char *) * (size_t)(nd + 1));
        done[nd++] = (char *)name;
        /* the fields in their order (icolumn) */
        jbuf fl = { 0 };
        for (int pos = 0; pos < 32; pos++)
            for (int j = 0; j < r.n; j++)
                if (!strcmp(r.rows[j][7], name) && atoi(r.rows[j][2]) == pos) {
                    if (fl.n) jb_s(&fl, ";");
                    jb_s(&fl, r.rows[j][5]); jb_s(&fl, "="); jb_s(&fl, r.rows[j][3]);
                }
        sqlite3_bind_text(st, 1, name, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, r.rows[i][6], -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 3, r.rows[i][4], -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 4, fl.p ? fl.p : "", -1, SQLITE_TRANSIENT);
        sqlite3_bind_int(st, 5, atoi(r.rows[i][1]));
        sqlite3_step(st);
        sqlite3_reset(st);
        jb_free(&fl);
        {
            jbuf ix = { 0 }, cols = { 0 };
            for (int pos = 0; pos < 32; pos++)
                for (int j = 0; j < r.n; j++)
                    if (!strcmp(r.rows[j][7], name) && atoi(r.rows[j][2]) == pos) {
                        if (ix.n) { jb_s(&ix, ";"); jb_s(&cols, ", "); }
                        jb_s(&ix, "+"); jb_s(&ix, r.rows[j][3]);
                        jb_qid(&cols, r.rows[j][3], -1);
                    }
            sqlite3_stmt *x = 0;
            sqlite3_prepare_v2(db, "INSERT OR IGNORE INTO MSysLB_Indexes (tbl, name, fields, isprimary, isunique, required, ignorenulls, isforeign, ord) VALUES (?1, ?2, ?3, 0, 0, 0, 0, 1, 99)", -1, &x, 0);
            sqlite3_bind_text(x, 1, r.rows[i][4], -1, SQLITE_STATIC);
            sqlite3_bind_text(x, 2, name, -1, SQLITE_STATIC);
            sqlite3_bind_text(x, 3, ix.p ? ix.p : "", -1, SQLITE_STATIC);
            if (!(atoi(r.rows[i][1]) & 2)) sqlite3_step(x);   /* a relation not enforced has no index */
            int added = sqlite3_changes(db) > 0;
            sqlite3_finalize(x);
            if (added && cols.n) {
                jbuf s = { 0 };
                jbuf nm = { 0 }; jb_f(&nm, "%s|%s", r.rows[i][4], name);
                jb_s(&s, "CREATE INDEX "); jb_qid(&s, nm.p, -1); jb_s(&s, " ON "); jb_qid(&s, r.rows[i][4], -1);
                jb_s(&s, " ("); jb_put(&s, cols.p, cols.n); jb_s(&s, ")");
                sqlite3_exec(db, s.p, 0, 0, 0);
                jb_free(&s); jb_free(&nm);
            }
            jb_free(&ix); jb_free(&cols);
        }
    }
    free(done);
    sqlite3_finalize(st);
    free_rows(&r);
}

/* ---------------------------------------------------------------- queries
 * MSysQueries keeps a query as rows by attribute: 1 the kind (Flag: 1 select, 2 make-table, 3 append,
 * 4 update, 5 delete, 6 crosstab, 7 DDL, 8 pass-through, 9 union; Name1 the target table), 2 a
 * parameter, 3 DISTINCT / TOP (Flag bits, Name1 the number), 5 a table (Name2 its alias) or a part of a
 * union (Expression), 6 a column (Name1 its alias, Name2 the target column), 7 a join (Name1, Name2,
 * Expression the condition, Flag 1 inner, 2 left, 3 right), 8 WHERE, 9 GROUP BY, 10 HAVING, 11 ORDER BY
 * (Name1 "D": descending). Written again as Jet shows the SQL of a QueryDef. */
static int needs_brackets(const char *n)
{
    if (!*n || isdigit((unsigned char)*n)) return 1;
    for (const char *p = n; *p; p++) if (!isalnum((unsigned char)*p) && *p != '_' && (unsigned char)*p < 0x80) return 1;
    return 0;
}
static void put_name(jbuf *b, const char *n)
{
    if (needs_brackets(n)) { jb_s(b, "["); jb_s(b, n); jb_s(b, "]"); }
    else jb_s(b, n);
}

typedef struct qsrc { char *name; char *sql; struct qsrc *l, *r; int jt; jbuf on; int non; } qsrc;
static qsrc *src_table(const char *t, const char *alias)
{
    qsrc *s = calloc(1, sizeof *s);
    jbuf b = { 0 };
    put_name(&b, t);
    if (alias && *alias) { jb_s(&b, " AS "); put_name(&b, alias); s->name = strdup(alias); }
    else s->name = strdup(t);
    s->sql = b.p;
    return s;
}
static int src_has(qsrc *s, const char *n)
{
    if (s->name) return !strcasecmp(s->name, n);
    return src_has(s->l, n) || src_has(s->r, n);
}
static void src_sql(qsrc *s, jbuf *b, int nested)
{
    if (s->name) { jb_s(b, s->sql); return; }
    if (nested) jb_s(b, "(");
    src_sql(s->l, b, 1);
    jb_s(b, s->jt == 2 ? " LEFT JOIN " : s->jt == 3 ? " RIGHT JOIN " : " INNER JOIN ");
    src_sql(s->r, b, 1);
    jb_s(b, " ON ");
    jb_put(b, s->on.p, s->on.n);
    if (nested) jb_s(b, ")");
}
static void src_free(qsrc *s)
{
    if (!s) return;
    src_free(s->l); src_free(s->r);
    free(s->name); free(s->sql); jb_free(&s->on);
    free(s);
}
static void add_on(qsrc *s, const char *e)
{
    if (s->non == 0) jb_s(&s->on, e);
    else {
        jbuf n = { 0 };
        if (s->non == 1) { jb_s(&n, "("); jb_put(&n, s->on.p, s->on.n); jb_s(&n, ") AND ("); jb_s(&n, e); jb_s(&n, ")"); }
        else { jb_put(&n, s->on.p, s->on.n); jb_s(&n, " AND ("); jb_s(&n, e); jb_s(&n, ")"); }
        jb_free(&s->on);
        s->on = n;
    }
    s->non++;
}

static const char *param_type_name(int t)
{
    switch (t) {
    case MDB_BOOL: return "Bit";
    case MDB_BYTE: return "Byte";
    case MDB_INT: return "Short";
    case MDB_LONGINT: return "Long";
    case MDB_MONEY: return "Currency";
    case MDB_FLOAT: return "IEEESingle";
    case MDB_DOUBLE: return "IEEEDouble";
    case MDB_DATETIME: return "DateTime";
    case MDB_BINARY: return "Binary";
    case MDB_TEXT: return "Text ( 255 )";
    case MDB_OLE: return "LongBinary";
    case MDB_MEMO: return "LongText";
    case MDB_REPID: return "Guid";
    }
    return "Value";
}

static char *query_text(strows *q, const char *id, int *qtype)
{
    /* columns: Attribute, Expression, Flag, Name1, Name2, ObjectId, Order */
    int kind = 1, flags = 0;
    char target[256] = "", top[64] = "";
    jbuf params = { 0 }, cols = { 0 }, tcols = { 0 }, where = { 0 }, group = { 0 }, having = { 0 }, order = { 0 };
    jbuf xvalue = { 0 }, xrows = { 0 }, xpivot = { 0 }, xgroup = { 0 };   /* a crosstab: TRANSFORM ... PIVOT */
    qsrc **srcs = 0; int ns = 0;
    char **unions = 0; int nu = 0;
    struct { char *a, *b, *e; int f; } *joins = 0; int nj = 0;
    /* the rows of this query in their Order */
    int *idx = 0, n = 0;
    for (int i = 0; i < q->n; i++) if (!strcmp(q->rows[i][5], id)) { idx = realloc(idx, sizeof(int) * (size_t)(n + 1)); idx[n++] = i; }
    for (int i = 0; i < n; i++)
        for (int j = i + 1; j < n; j++)
            if (strcmp(q->rows[idx[j]][6], q->rows[idx[i]][6]) < 0) { int t = idx[i]; idx[i] = idx[j]; idx[j] = t; }
    for (int k = 0; k < n; k++) {
        char **r = q->rows[idx[k]];
        int attr = atoi(r[0]), flag = atoi(r[2]);
        const char *expr = r[1], *n1 = r[3], *n2 = r[4];
        switch (attr) {
        case 1: kind = flag; snprintf(target, sizeof target, "%s", n1); break;
        case 2:
            if (params.n) jb_s(&params, ", ");
            if (strchr(n1, '[')) jb_s(&params, n1);
            else {                              /* Forms!Dialog!Date -> [Forms]![Dialog]![Date] */
                const char *pp = n1;
                int many = strchr(n1, '!') != 0;
                while (*pp) {
                    const char *e = strchr(pp, '!');
                    int l = e ? (int)(e - pp) : (int)strlen(pp);
                    char part[256]; snprintf(part, sizeof part, "%.*s", l, pp);
                    (void)many;
                    if (needs_brackets(part)) { jb_s(&params, "["); jb_s(&params, part); jb_s(&params, "]"); }
                    else jb_s(&params, part);
                    if (!e) break;
                    jb_s(&params, "!");
                    pp = e + 1;
                }
            }
            jb_s(&params, " "); jb_s(&params, param_type_name(flag));
            break;
        case 3: flags = flag; snprintf(top, sizeof top, "%s", n1); break;
        case 5:
            if (!*n1 && *expr) { unions = realloc(unions, sizeof(char *) * (size_t)(nu + 1)); unions[nu++] = strdup(expr); break; }
            srcs = realloc(srcs, sizeof(qsrc *) * (size_t)(ns + 1));
            srcs[ns++] = src_table(n1, n2);
            break;
        case 6:
            if (cols.n) jb_s(&cols, ", ");
            jb_s(&cols, expr);
            if (*n1 && strcmp(n1, expr)) {
                jb_s(&cols, " AS ");
                if (jet_is_function(n1)) { jb_s(&cols, "["); jb_s(&cols, n1); jb_s(&cols, "]"); }
                else put_name(&cols, n1);
            }

            if (*n2) { if (tcols.n) jb_s(&tcols, ", "); put_name(&tcols, n2); }
            {                                   /* its part in a crosstab: 0 the value, 1 the pivot, 2 a row */
                jbuf *xb = flag == 1 ? &xpivot : flag == 2 ? &xrows : &xvalue;
                if (xb->n) jb_s(xb, ", ");
                jb_s(xb, expr);
                if (flag != 1 && *n1 && strcmp(n1, expr)) {
                    jb_s(xb, " AS ");
                    if (jet_is_function(n1)) { jb_s(xb, "["); jb_s(xb, n1); jb_s(xb, "]"); }
                    else put_name(xb, n1);
                }
            }
            break;
        case 7:
            joins = realloc(joins, sizeof *joins * (size_t)(nj + 1));
            joins[nj].a = strdup(n1); joins[nj].b = strdup(n2); joins[nj].e = strdup(expr); joins[nj].f = flag;
            nj++;
            break;
        case 8: jb_s(&where, expr); break;
        case 9:
            if (group.n) jb_s(&group, ", ");
            jb_s(&group, expr);
            if (flag != 1) { if (xgroup.n) jb_s(&xgroup, ", "); jb_s(&xgroup, expr); }
            break;
        case 10: jb_s(&having, expr); break;
        case 11: if (order.n) jb_s(&order, ", "); jb_s(&order, expr); if (!strcmp(n1, "D")) jb_s(&order, " DESC"); break;
        }
    }
    /* the joins folded into the tables */
    for (int j = 0; j < nj; j++) {
        qsrc *from = 0, *to = 0;
        int fi = -1, ti = -1;
        for (int i = 0; i < ns; i++) {
            if (fi < 0 && src_has(srcs[i], joins[j].a)) fi = i;
            if (ti < 0 && src_has(srcs[i], joins[j].b)) ti = i;
        }
        if (fi >= 0 && fi == ti) { if (!srcs[fi]->name) add_on(srcs[fi], joins[j].e); continue; }
        from = fi >= 0 ? srcs[fi] : src_table(joins[j].a, 0);
        to = ti >= 0 ? srcs[ti] : src_table(joins[j].b, 0);
        /* the two out of the list (the higher index first) */
        int hi = fi > ti ? fi : ti, lo = fi > ti ? ti : fi;
        if (hi >= 0) { memmove(srcs + hi, srcs + hi + 1, sizeof(qsrc *) * (size_t)(ns - hi - 1)); ns--; }
        if (lo >= 0) { memmove(srcs + lo, srcs + lo + 1, sizeof(qsrc *) * (size_t)(ns - lo - 1)); ns--; }
        qsrc *jn = calloc(1, sizeof *jn);
        jn->l = from; jn->r = to; jn->jt = joins[j].f;
        add_on(jn, joins[j].e);
        srcs = realloc(srcs, sizeof(qsrc *) * (size_t)(ns + 1));
        srcs[ns++] = jn;
    }
    jbuf from = { 0 };
    for (int i = 0; i < ns; i++) { if (i) jb_s(&from, ", "); src_sql(srcs[i], &from, 0); }

    jbuf out = { 0 };
    if (params.n) { jb_s(&out, "PARAMETERS "); jb_put(&out, params.p, params.n); jb_s(&out, ";\r\n"); }
    const char *pred = (flags & 0x08) ? "DISTINCTROW " : (flags & 0x02) && kind != 9 ? "DISTINCT " : "";
    jbuf topb = { 0 };
    if (flags & 0x30) { jb_s(&topb, "TOP "); jb_s(&topb, top); jb_s(&topb, (flags & 0x20) ? " PERCENT " : " "); }
    jbuf sel = { 0 };
    if (flags & 0x01) { jb_s(&sel, "*"); if (cols.n) jb_s(&sel, ", "); }
    jb_put(&sel, cols.p, cols.n);
    *qtype = 0;
    switch (kind) {
    case 9:                                     /* union */
        for (int i = 0; i < nu; i++) {
            char *u = unions[i];
            size_t l = strlen(u);
            while (l && strchr(" \t\r\n;", u[l - 1])) u[--l] = 0;
            if (i) jb_s(&out, (flags & 0x02) ? "\r\nUNION " : "\r\nUNION ALL ");
            jb_s(&out, u);
        }
        if (order.n) { jb_s(&out, "\r\nORDER BY "); jb_put(&out, order.p, order.n); }
        *qtype = 128;
        break;
    case 4:                                     /* update */
        jb_s(&out, "UPDATE "); jb_put(&out, from.p, from.n);
        jb_s(&out, " SET "); jb_put(&out, sel.p, sel.n);
        *qtype = 48;
        break;
    case 5:
        jb_s(&out, "DELETE "); jb_put(&out, sel.p, sel.n); if (sel.n) jb_s(&out, " ");
        jb_s(&out, "\r\nFROM "); jb_put(&out, from.p, from.n);
        *qtype = 32;
        break;
    case 6:                                     /* crosstab */
        jb_s(&out, "TRANSFORM "); jb_put(&out, xvalue.p, xvalue.n);
        jb_s(&out, "\r\nSELECT "); jb_put(&out, topb.p, topb.n); jb_put(&out, xrows.p, xrows.n);
        if (from.n) { jb_s(&out, "\r\nFROM "); jb_put(&out, from.p, from.n); }
        if (where.n) { jb_s(&out, "\r\nWHERE "); jb_put(&out, where.p, where.n); }
        if (xgroup.n) { jb_s(&out, "\r\nGROUP BY "); jb_put(&out, xgroup.p, xgroup.n); }
        if (having.n) { jb_s(&out, "\r\nHAVING "); jb_put(&out, having.p, having.n); }
        if (order.n) { jb_s(&out, "\r\nORDER BY "); jb_put(&out, order.p, order.n); }
        jb_s(&out, "\r\nPIVOT "); jb_put(&out, xpivot.p, xpivot.n);
        *qtype = 16;
        break;
    default:
        if (kind == 3) {
            jb_s(&out, "INSERT INTO "); put_name(&out, target);
            if (tcols.n) { jb_s(&out, " ( "); jb_put(&out, tcols.p, tcols.n); jb_s(&out, " )"); }
            jb_s(&out, "\r\n");
            *qtype = 64;
        }
        jb_s(&out, "SELECT "); jb_s(&out, pred); jb_put(&out, topb.p, topb.n); jb_put(&out, sel.p, sel.n);
        if (kind == 2) { jb_s(&out, " INTO "); put_name(&out, target); *qtype = 80; }
        if (kind == 6) *qtype = 16;
        if (from.n) { jb_s(&out, "\r\nFROM "); jb_put(&out, from.p, from.n); }
    }
    if (kind != 9 && kind != 6) {
        if (where.n) { jb_s(&out, "\r\nWHERE "); jb_put(&out, where.p, where.n); }
        if (group.n) { jb_s(&out, "\r\nGROUP BY "); jb_put(&out, group.p, group.n); }
        if (having.n) { jb_s(&out, "\r\nHAVING "); jb_put(&out, having.p, having.n); }
        if (order.n) { jb_s(&out, "\r\nORDER BY "); jb_put(&out, order.p, order.n); }
    }
    if (flags & 0x04) jb_s(&out, "\r\nWITH OWNERACCESS OPTION");
    jb_s(&out, ";\r\n");
    for (int i = 0; i < ns; i++) src_free(srcs[i]);
    free(srcs);
    for (int i = 0; i < nu; i++) free(unions[i]);
    free(unions);
    for (int j = 0; j < nj; j++) { free(joins[j].a); free(joins[j].b); free(joins[j].e); }
    free(joins);
    free(idx);
    jb_free(&params); jb_free(&cols); jb_free(&tcols); jb_free(&where); jb_free(&group); jb_free(&having);
    jb_free(&order); jb_free(&from); jb_free(&topb); jb_free(&sel);
    jb_free(&xvalue); jb_free(&xrows); jb_free(&xpivot); jb_free(&xgroup);
    return out.p;
}

static void import_queries(MdbHandle *mdb, sqlite3 *db)
{
    static const char *const O[] = { "Id", "Name", "Type", 0 };
    static const char *const Q[] = { "Attribute", "Expression", "Flag", "Name1", "Name2", "ObjectId", "Order", 0 };
    strows objs, qs;
    if (read_rows(mdb, "MSysObjects", O, &objs)) return;
    if (read_rows(mdb, "MSysQueries", Q, &qs)) { free_rows(&objs); return; }
    sqlite3_stmt *st = 0;
    sqlite3_prepare_v2(db, "INSERT OR IGNORE INTO MSysLB_Queries (name, sql, type) VALUES (?1, ?2, ?3)", -1, &st, 0);
    for (int i = 0; i < objs.n; i++) {
        if ((atoi(objs.rows[i][2]) & 0x7F) != MDB_QUERY) continue;
        const char *name = objs.rows[i][1];
        if (name[0] == '~') continue;           /* the hidden queries of forms and reports */
        int qt;
        char *sql = query_text(&qs, objs.rows[i][0], &qt);
        sqlite3_bind_text(st, 1, name, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, sql, -1, SQLITE_TRANSIENT);
        sqlite3_bind_int(st, 3, qt);
        sqlite3_step(st);
        sqlite3_reset(st);
        free(sql);
    }
    sqlite3_finalize(st);
    free_rows(&objs);
    free_rows(&qs);
}

/* ---------------------------------------------------------------- the import */
int mdb_import(const char *path, const char *sqlite_path, char *err, int cap)
{
    if (cap) *err = 0;
    MdbHandle *mdb = mdb_open(path, MDB_NOFLAGS);
    if (!mdb) { snprintf(err, (size_t)cap, "cannot open %s", path); return -1; }
    mdb_set_bind_size(mdb, 1 << 20);
    mdb_set_boolean_fmt_numbers(mdb);
    if (!mdb_read_catalog(mdb, MDB_ANY)) { mdb_close(mdb); snprintf(err, (size_t)cap, "not a Jet database"); return -1; }
    char *tmp = malloc(strlen(sqlite_path) + 8);
    sprintf(tmp, "%s.part", sqlite_path);
    remove(tmp);
    sqlite3 *db = 0;
    if (sqlite3_open(tmp, &db) != SQLITE_OK) { mdb_close(mdb); snprintf(err, (size_t)cap, "cannot write %s", tmp); free(tmp); return -1; }
    jet_register(db);
    exec(db, "PRAGMA legacy_alter_table = ON; PRAGMA journal_mode = OFF; PRAGMA synchronous = OFF", err, cap);
    ddb *d = (ddb *)vb_onew(&vb_cls_dao_database);
    vb_oaddref((vbobj)d);
    d->db = db;
    d->closed = 1;                              /* not a database of the program: db_clear leaves it */
    int r = schema_init(d);
    if (!r) r = exec(db, "BEGIN", err, cap);
    for (unsigned i = 0; !r && i < mdb->num_catalog; i++) {
        MdbCatalogEntry *e = g_ptr_array_index(mdb->catalog, i);
        if (e->object_type == MDB_TABLE && mdb_is_user_table(e)) r = import_schema(mdb, e, db, err, cap);
    }
    if (!r) r = schema_build_all(d);
    for (unsigned i = 0; !r && i < mdb->num_catalog; i++) {
        MdbCatalogEntry *e = g_ptr_array_index(mdb->catalog, i);
        if (e->object_type == MDB_TABLE && mdb_is_user_table(e)) r = import_rows(mdb, e, db, err, cap);
    }
    if (!r) {
        import_relations(mdb, db);
        import_queries(mdb, db);
        char sql[300];
        snprintf(sql, sizeof sql, "INSERT OR REPLACE INTO MSysLB_Props VALUES ('Version', '%s')", IS_JET3(mdb) ? "3.0" : "4.0");
        exec(db, sql, err, cap);
        r = schema_after_import(d);
    }
    if (!r) r = exec(db, "COMMIT", err, cap);
    if (r && !*err) snprintf(err, (size_t)cap, "%s", sqlite3_errmsg(db));
    vb_err_clear_pending();
    d->db = 0;
    vb_orelease((vbobj)d);
    sqlite3_close(db);
    mdb_close(mdb);
    if (!r) {
        remove(sqlite_path);
        if (rename(tmp, sqlite_path)) { snprintf(err, (size_t)cap, "cannot write %s", sqlite_path); r = -1; }
    } else remove(tmp);
    free(tmp);
    return r;
}
