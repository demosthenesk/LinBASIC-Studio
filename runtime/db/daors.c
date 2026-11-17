/* daors.c - the DAO Recordset: the rows of a table, a query or a SELECT, read into memory, with
 * Jet's rules for moving (BOF/EOF, RecordCount of a dynaset), editing (Edit/AddNew/Update, the
 * errors of the Jet engine), Find (criteria in Jet SQL), Seek (table-type, by an index), Bookmark,
 * Clone, Sort/Filter and Requery. A change goes to the base table by its rowid.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "daoint.h"
#include <ctype.h>

#define RS(o) ((drs *)vb_oid(o))
static const char *SRC = "DAO.Recordset";

static char *sdup(const char *s) { return s ? strdup(s) : 0; }

/* ---------------------------------------------------------------- rows */
static int alive(drs *r, int i) { return i >= 0 && i < r->d->n && !r->d->r[i].del; }
static int n_alive(drs *r) { int k = 0; for (int i = 0; i < r->d->n; i++) if (!r->d->r[i].del) k++; return k; }
static int first_alive(drs *r) { for (int i = 0; i < r->d->n; i++) if (!r->d->r[i].del) return i; return -1; }
static int last_alive(drs *r) { for (int i = r->d->n - 1; i >= 0; i--) if (!r->d->r[i].del) return i; return -1; }
static int is_bof(drs *r) { return r->cur < 0 || n_alive(r) == 0; }
static int is_eof(drs *r) { return r->cur >= r->d->n || n_alive(r) == 0; }
/* a sorted query that no index gives in order: Jet reads all its rows before the first, so RecordCount
 * counts them all (VB6 oracles: f_data ORDER BY a field without an index: all; d_dao1 ORDER BY an
 * indexed field: the rows reached) */
static int sorted_sql(ddb *d, const char *base, const char *sql)
{
    if (!sql) return 0;
    int depth = 0;
    const char *ob = 0;
    for (const char *p = sql; *p; p++) {
        if (*p == '\'' || *p == '"') { char q = *p; for (p++; *p && *p != q; p++) ; if (!*p) break; continue; }
        if (*p == '(') depth++;
        else if (*p == ')') depth--;
        else if (!depth && !strncasecmp(p, "ORDER BY", 8)) ob = p + 8;
    }
    if (!ob) return 0;
    if (!base || !d) return 1;
    /* the first sort column: "T"."c", "c", c */
    char col[256]; int n = 0;
    const char *p = ob;
    while (*p == ' ') p++;
    for (;;) {
        n = 0;
        if (*p == '"') { for (p++; *p && *p != '"' && n < 255; p++) col[n++] = *p; if (*p) p++; }
        else for (; *p && (isalnum((unsigned char)*p) || *p == '_') && n < 255; p++) col[n++] = *p;
        col[n] = 0;
        if (*p == '.') { p++; continue; }
        break;
    }
    if (!n) return 1;
    sqlite3_stmt *st;
    int indexed = 0;
    if (sqlite3_prepare_v2(d->db, "SELECT fields FROM MSysLB_Indexes WHERE tbl = ?1", -1, &st, 0) == SQLITE_OK) {
        sqlite3_bind_text(st, 1, base, -1, SQLITE_STATIC);
        while (!indexed && sqlite3_step(st) == SQLITE_ROW) {
            const char *f = (const char *)sqlite3_column_text(st, 0);
            if (!f) continue;
            if (*f == '+' || *f == '-') f++;
            size_t k = strcspn(f, ";");
            if (k == (size_t)n && !strncasecmp(f, col, k)) indexed = 1;
        }
        sqlite3_finalize(st);
    }
    return !indexed;
}
static void reach(drs *r) { if (r->cur >= 0 && r->cur < r->d->n && r->cur + 1 > r->reached) r->reached = r->cur + 1; }

static void row_free(rrow *w, int ncols)
{
    if (!w->v) return;
    for (int i = 0; i < ncols; i++) vb_vfree(&w->v[i]);
    free(w->v);
    w->v = 0;
}
static void data_release(rdata *d, int ncols)
{
    if (!d || --d->refs > 0) return;
    for (int i = 0; i < d->n; i++) row_free(&d->r[i], ncols);
    free(d->r);
    free(d);
}
static rrow *data_add(rdata *d)
{
    if (d->n == d->cap) { d->cap = d->cap ? d->cap * 2 : 64; d->r = realloc(d->r, sizeof(rrow) * (size_t)d->cap); }
    rrow *w = &d->r[d->n++];
    memset(w, 0, sizeof *w);
    return w;
}

static void cols_free(rcol *c, int n)
{
    for (int i = 0; i < n; i++) { free(c[i].name); free(c[i].stable); free(c[i].sfield); free(c[i].defval); }
    free(c);
}

static void edit_clear(drs *r)
{
    if (r->ebuf) { for (int i = 0; i < r->ncols; i++) vb_vfree(&r->ebuf[i]); free(r->ebuf); free(r->edirty); }
    r->ebuf = 0; r->edirty = 0;
    r->editmode = 0;
}

static void mirror_drop(drs *r)
{
    if (r->mirror) sqlite3_close(r->mirror);
    r->mirror = 0;
    r->mirror_ok = 0;
}

void rs_clear(vbobj o)
{
    drs *r = (drs *)o;
    edit_clear(r);
    mirror_drop(r);
    data_release(r->d, r->ncols);
    r->d = 0;
    cols_free(r->cols, r->ncols);
    r->cols = 0;
    free(r->source); free(r->sqlite); free(r->base); free(r->index); free(r->ixcols); free(r->ixdesc);
    r->source = r->sqlite = r->base = r->index = 0;
    vb_sfree(&r->filter); vb_sfree(&r->sort);
    if (r->fields) { vbobj f = (vbobj)r->fields; r->fields = 0; vb_orelease(f); }
    if (r->db) { vbobj d = r->db; r->db = 0; vb_orelease(d); }
    if (r->parent_qd) { vbobj q = r->parent_qd; r->parent_qd = 0; vb_orelease(q); }
}

void rs_close(drs *r)
{
    if (r->closed) return;
    if (r->db) {
        ddb *d = DB(r->db);
        if (d->rss) for (int i = 0; i < d->rss->n; i++) if (vb_oid(d->rss->items[i]) == (vbobj)r) { coll_remove(d->rss, i); break; }
    }
    r->closed = 1;
    edit_clear(r);
    mirror_drop(r);
}

/* ---------------------------------------------------------------- opening */
static int col_index(drs *r, const char *name)
{
    for (int i = 0; i < r->ncols; i++) if (!strcasecmp(r->cols[i].name, name)) return i;
    /* T.Name of a join, by the field name alone */
    for (int i = 0; i < r->ncols; i++) {
        const char *dot = strchr(r->cols[i].name, '.');
        if (dot && !strcasecmp(dot + 1, name)) return i;
    }
    return -1;
}

/* the type of a column with no declared type: from its values */
static void infer_type(drs *r, int c)
{
    int t = 0;
    for (int i = 0; i < r->d->n; i++) {
        vbvar v = r->d->r[i].v[c];
        int k = 0;
        switch (v.vt) {
        case VT_I2: k = dbInteger; break;
        case VT_I4: k = dbLong; break;
        case VT_R4: case VT_R8: k = dbDouble; break;
        case VT_CY: k = dbCurrency; break;
        case VT_DATE: k = dbDate; break;
        case VT_BOOL: k = dbBoolean; break;
        case VT_UI1: k = dbByte; break;
        case VT_BSTR: k = dbText; break;
        default: if (v.vt & VT_ARRAY) k = dbLongBinary;
        }
        if (!k) continue;
        if (!t) t = k;
        else if (t != k) {
            if ((t == dbLong && k == dbDouble) || (t == dbDouble && k == dbLong)) t = dbDouble;
            else t = dbText;
        }
    }
    if (!t) t = dbBinary;                       /* only Nulls: Jet says dbBinary */
    r->cols[c].type = t;
    switch (t) {
    case dbLong: case dbSingle: r->cols[c].size = 4; break;
    case dbInteger: r->cols[c].size = 2; break;
    case dbDouble: case dbCurrency: case dbDate: r->cols[c].size = 8; break;
    case dbText: r->cols[c].size = 255; break;
    case dbBoolean: case dbByte: r->cols[c].size = 1; break;
    case dbBinary: r->cols[c].size = 510; break;
    default: r->cols[c].size = 0;
    }
    r->cols[c].attrs = (t == dbText || t == dbMemo || t == dbBinary || t == dbLongBinary) ? dbVariableField : dbFixedField;
    /* the values as that type */
    for (int i = 0; i < r->d->n; i++) {
        vbvar *v = &r->d->r[i].v[c];
        if (v->vt == VT_NULL || v->vt == VT_EMPTY) continue;
        vb_i32 m = vb_mark();
        vbvar nv = t == dbText && v->vt != VT_BSTR ? vb_v_s(vb_v2s(*v)) : t == dbDouble && v->vt == VT_I4 ? vb_v_d(v->u.l) : *v;
        if (nv.vt != v->vt) vb_vset(v, nv);
        vb_release(m);
    }
}

static int load_rows(drs *r, sqlite3_stmt *st, int ridcol)
{
    int nc = sqlite3_column_count(st);
    for (;;) {
        int rc = sqlite3_step(st);
        if (rc == SQLITE_DONE) break;
        if (rc != SQLITE_ROW) return rc;
        rrow *w = data_add(r->d);
        w->v = calloc((size_t)(r->ncols ? r->ncols : 1), sizeof(vbvar));
        vb_i32 m = vb_mark();
        for (int i = 0, k = 0; i < nc; i++) {
            if (i == ridcol) { w->rid = sqlite3_column_int64(st, i); continue; }
            vbvar v = jet_column(st, i, r->cols[k].type);
            vb_vset(&w->v[k], v);
            k++;
        }
        vb_release(m);
    }
    return SQLITE_OK;
}

/* the types of the computed columns of the statement being opened (from jet_translate) */
static jctype *pend_ct; static int pend_nct;

static int coltype(void *ctx, const char *qual, const char *col)
{
    ddb *d = ctx;
    sqlite3_stmt *st = 0;
    int t = 0;
    if (qual && sqlite3_prepare_v2(d->db, "SELECT type FROM MSysLB_Fields WHERE tbl = ?1 AND name = ?2", -1, &st, 0) == SQLITE_OK) {
        sqlite3_bind_text(st, 1, qual, -1, SQLITE_STATIC);
        sqlite3_bind_text(st, 2, col, -1, SQLITE_STATIC);
        if (sqlite3_step(st) == SQLITE_ROW) t = sqlite3_column_int(st, 0);
        sqlite3_finalize(st);
        if (t) return t;
    }
    if (sqlite3_prepare_v2(d->db, "SELECT type FROM MSysLB_Fields WHERE name = ?1 LIMIT 1", -1, &st, 0) == SQLITE_OK) {
        sqlite3_bind_text(st, 1, col, -1, SQLITE_STATIC);
        if (sqlite3_step(st) == SQLITE_ROW) t = sqlite3_column_int(st, 0);
        sqlite3_finalize(st);
    }
    return t;
}
int dao_translate(ddb *d, const char *sql, jsql *j, const char *const *params)
{
    jet_coltype_fn = coltype; jet_coltype_ctx = d;
    int r = jet_translate(sql, j, params);
    jet_coltype_fn = 0; jet_coltype_ctx = 0;
    return r;
}

/* the columns of the statement; names as Jet gives them (T.ID for a name two tables have) */
static void make_cols(drs *r, ddb *d, sqlite3_stmt *st, int ridcol)
{
    int nc = sqlite3_column_count(st);
    r->ncols = nc - (ridcol >= 0 ? 1 : 0);
    r->cols = calloc((size_t)(r->ncols ? r->ncols : 1), sizeof(rcol));
    for (int i = 0, k = 0; i < nc; i++) {
        if (i == ridcol) continue;
        rcol *c = &r->cols[k++];
        c->name = sdup(sqlite3_column_name(st, i));
        int size = 0, attrs = 0;
        c->type = jet_decl_type(sqlite3_column_decltype(st, i), &size, &attrs);
        c->size = size; c->attrs = attrs;
        if (!c->type)
            for (int k2 = 0; k2 < pend_nct; k2++)
                if (!strcasecmp(pend_ct[k2].name, c->name)) {
                    c->type = pend_ct[k2].type;
                    switch (c->type) {
                    case dbText: c->size = 255; c->attrs = dbVariableField; break;
                    case dbMemo: case dbLongBinary: c->size = 0; c->attrs = dbVariableField; break;
                    case dbBinary: c->size = 510; c->attrs = dbVariableField; break;
                    case dbBoolean: case dbByte: c->size = 1; c->attrs = dbFixedField; break;
                    case dbInteger: c->size = 2; c->attrs = dbFixedField; break;
                    case dbLong: case dbSingle: c->size = 4; c->attrs = dbFixedField; break;
                    case dbDecimal: c->size = 16; c->attrs = dbFixedField; break;
                    default: c->size = 8; c->attrs = dbFixedField;
                    }
                }
        const char *tn = sqlite3_column_table_name(st, i), *on = sqlite3_column_origin_name(st, i);
        if (tn && on) {
            c->stable = sdup(tn); c->sfield = sdup(on);
            if (schema_field_info(d, tn, on, c)) {
                if (c->autoinc) c->attrs |= dbAutoIncrField;
            }
            c->updatable = r->base && !strcasecmp(r->base, tn);
        }
    }
    /* SQLite names a second "ID" "ID:1": the name itself */
    for (int k = 0; k < r->ncols; k++) {
        char *colon = strrchr(r->cols[k].name, ':');
        if (colon && isdigit((unsigned char)colon[1]) && r->cols[k].sfield && !strncasecmp(r->cols[k].name, r->cols[k].sfield, (size_t)(colon - r->cols[k].name)))
            *colon = 0;
    }
    /* a name two tables have: Table.Field */
    int *dup = calloc(r->ncols > 0 ? (size_t)r->ncols : 1, sizeof(int));
    for (int k = 0; k < r->ncols; k++)
        for (int j = 0; j < r->ncols; j++) if (j != k && !strcasecmp(r->cols[j].name, r->cols[k].name)) dup[k] = 1;
    for (int k = 0; k < r->ncols; k++) {
        rcol *c = &r->cols[k];
        if (dup[k] && c->stable) {
            char *nm = malloc(strlen(c->stable) + strlen(c->name) + 2);
            sprintf(nm, "%s.%s", c->stable, c->name);
            free(c->name); c->name = nm;
        }
    }
    free(dup);
}

static void make_fields(drs *r)
{
    r->fields = coll_new(C_RSFIELDS, (vbobj)r);
    for (int i = 0; i < r->ncols; i++) {
        dfld *f = field_new(FK_RS, 0, r->cols[i].name, 0, 0);
        f->owner = (vbobj)r;                    /* not counted: the recordset holds its fields */
        f->col = i;
        f->appended = 1;
        coll_add(r->fields, (vbobj)f);
        vb_orelease((vbobj)f);
    }
}

static int fill(drs *r, ddb *d, const char *sql, int nbinds, vbvar *binds)
{
    sqlite3_stmt *st = 0;
    int rc = sqlite3_prepare_v2(d->db, sql, -1, &st, 0);
    if (rc != SQLITE_OK) return rc;
    for (int i = 0; i < nbinds; i++) jet_bind(st, i + 1, binds[i], 0);
    int ridcol = -1;
    for (int i = 0; i < sqlite3_column_count(st); i++) if (!strcmp(sqlite3_column_name(st, i), "_lbrid")) ridcol = i;
    if (!r->cols) make_cols(r, d, st, ridcol);
    rc = load_rows(r, st, ridcol);
    if (rc != SQLITE_OK) { sqlite3_reset(st); }
    sqlite3_finalize(st);
    if (rc != SQLITE_OK) return rc;
    for (int i = 0; i < r->ncols; i++) if (!r->cols[i].type) infer_type(r, i);
    return SQLITE_OK;
}

static int rs_error_from_sqlite(ddb *d, int rc)
{
    return dao_raise_sqlite(d, rc, "DAO.Database");
}

vbobj rs_open_sqlite(ddb *d, const char *sqlite_sql, const char *base, int type, const char *source, int nbinds, vbvar *binds)
{
    drs *r = (drs *)vb_onew(&vb_cls_dao_recordset);
    r->db = (vbobj)d; vb_oaddref((vbobj)d);
    r->type = type;
    r->source = sdup(source);
    r->base = sdup(base);
    r->sqlite = sdup(sqlite_sql);
    r->d = calloc(1, sizeof(rdata));
    r->d->refs = 1;
    r->lastmod = -1;
    vb_sset(&r->filter, vb_sfromutf8("")); vb_sset(&r->sort, vb_sfromutf8(""));
    int rc = fill(r, d, sqlite_sql, nbinds, binds);
    if (rc != SQLITE_OK) {
        rs_error_from_sqlite(d, rc);
        return 0;
    }
    make_fields(r);
    r->cur = first_alive(r) >= 0 ? 0 : -1;
    if (r->cur >= 0) reach(r);
    if (r->type == dbOpenTable || sorted_sql(d, r->base, r->sqlite)) r->reached = r->d->n;
    coll_add(d->rss, (vbobj)r);
    return (vbobj)r;
}

/* the query text of a saved query */
static char *query_sql(ddb *d, const char *name)
{
    sqlite3_stmt *st = 0;
    if (sqlite3_prepare_v2(d->db, "SELECT sql FROM MSysLB_Queries WHERE name = ?1", -1, &st, 0) != SQLITE_OK) return 0;
    sqlite3_bind_text(st, 1, name, -1, SQLITE_STATIC);
    char *s = 0;
    if (sqlite3_step(st) == SQLITE_ROW) s = sdup((const char *)sqlite3_column_text(st, 0));
    sqlite3_finalize(st);
    return s;
}

static int starts_sql(const char *s)
{
    while (*s && isspace((unsigned char)*s)) s++;
    static const char *const K[] = { "select", "transform", "parameters", "insert", "update", "delete", "create", "drop", "alter", 0 };
    for (int i = 0; K[i]; i++) {
        size_t n = strlen(K[i]);
        if (!strncasecmp(s, K[i], n) && (isspace((unsigned char)s[n]) || s[n] == '(' || s[n] == '[' || s[n] == '*' || !s[n])) return 1;
    }
    return *s == '(';
}

/* the parameters of a query: their values by name; 0 or Jet's error (Too few parameters) */
static int param_values(jsql *j, vbobj qd, vbvar **vals)
{
    *vals = calloc((size_t)(j->nbinds ? j->nbinds : 1), sizeof(vbvar));
    int missing = 0;
    for (int i = 0; i < j->nbinds; i++) {
        int found = 0;
        if (qd) {
            dqd *q = (dqd *)vb_oid(qd);
            for (int k = 0; k < q->params->n; k++) {
                dpar *p = (dpar *)q->params->items[k];
                char pn[256];
                dao_param_plain(p->name, pn, sizeof pn);
                if (!strcasecmp(pn, j->binds[i])) {
                    if (p->value.vt != VT_EMPTY) { (*vals)[i] = p->value; found = 1; }
                    break;
                }
            }
        }
        if (!found) missing++;
    }
    return missing;
}

vbobj rs_open(ddb *d, const char *source, int type, int options, int lockedits, vbobj qd, vbvar *pvals, int npvals)
{
    (void)pvals; (void)npvals;
    char real[256];
    int what = starts_sql(source) ? 0 : schema_has_table(d, source, real, sizeof real);
    char *sql = 0;
    vbobj qd_used = qd;
    if (!starts_sql(source) && !what) {
        dao_raise(3078, "DAO.Database", "The Microsoft Jet database engine cannot find the input table or query '%s'.  Make sure it exists and that its name is spelled correctly.", source);
        return 0;
    }
    if (what == 1) {                            /* a table */
        if (!type) type = dbOpenTable;
        jbuf s = { 0 };
        jb_s(&s, "SELECT rowid AS _lbrid, * FROM "); jb_qid(&s, real, -1);
        vbobj o = rs_open_sqlite(d, s.p, d->ro ? 0 : real, type, real, 0, 0);
        jb_free(&s);
        if (o) {
            drs *r = (drs *)o;
            r->options = options; r->lockedits = lockedits;
            for (int i = 0; i < r->ncols; i++) r->cols[i].updatable = !d->ro;
        }
        return o;
    }
    if (what == 2) {                            /* a saved query */
        sql = query_sql(d, real);
        if (!sql) {                             /* a view of another program */
            jbuf s = { 0 };
            jb_s(&s, "SELECT * FROM "); jb_qid(&s, real, -1);
            sql = s.p;
        }
        if (!qd_used) {
            /* opened by name: the query's parameters have no values */
            dqd *tmp = querydef_new((vbobj)d, real, sql);
            qd_used = (vbobj)tmp;
            vbobj o = rs_open(d, sql, type, options, lockedits, qd_used, 0, 0);
            if (o) { free(RS(o)->source); RS(o)->source = sdup(real); }
            vb_orelease(qd_used);
            free(sql);
            return o;
        }
        source = sql;
    }
    if (type == dbOpenTable) {
        free(sql);
        dao_raise(3219, SRC, "Invalid operation.");
        return 0;
    }
    if (!type) type = dbOpenDynaset;
    const char *params[64]; int np = 0;
    if (qd_used) {
        dqd *q = (dqd *)vb_oid(qd_used);
        for (int k = 0; k < q->params->n && np < 63; k++) {
            const char *pn = ((dpar *)q->params->items[k])->name;
            static char buf[64][256];
            size_t pl = strlen(pn);
            (void)pl; dao_param_plain(pn, buf[np], 256); pn = buf[np];
            params[np++] = pn;
        }
    }
    params[np] = 0;
    jsql j;
    if (dao_translate(d, source, &j, np ? params : 0) < 0) {
        dao_raise(j.err, "DAO.Database", "%s", j.errmsg);
        jsql_free(&j); free(sql);
        return 0;
    }
    if (j.kind != J_SELECT || !j.returns) {
        jsql_free(&j); free(sql);
        dao_raise(3219, "DAO.Database", "Invalid operation.");
        return 0;
    }
    vbvar *vals;
    int missing = param_values(&j, qd_used, &vals);
    if (missing) {
        dao_raise(3061, "DAO.Database", "Too few parameters. Expected %d.", missing);
        free(vals); jsql_free(&j); free(sql);
        return 0;
    }
    const char *base = j.base;
    char basebuf[256];
    if (base && schema_has_table(d, base, basebuf, sizeof basebuf) != 1) base = 0;
    if (base) base = basebuf;
    jbuf s = { 0 };
    if (base && type != dbOpenSnapshot && !d->ro) {
        /* the rowid of each row: SELECT <alias>.rowid AS _lbrid, ... */
        jb_s(&s, "SELECT ");
        jb_qid(&s, j.base_alias ? j.base_alias : j.base, -1);
        jb_s(&s, ".rowid AS _lbrid, ");
        jb_s(&s, j.sql + 7);
    } else {
        jb_s(&s, j.sql);
        base = 0;
    }
    pend_ct = j.ctypes; pend_nct = j.nctypes;
    vbobj o = rs_open_sqlite(d, s.p, base, type, qd_used && what == 2 ? source : source, j.nbinds, vals);
    pend_ct = 0; pend_nct = 0;
    jb_free(&s);
    if (o) {
        drs *r = (drs *)o;
        r->options = options; r->lockedits = lockedits;
        if (qd_used) { r->parent_qd = qd_used; vb_oaddref(qd_used); }
    }
    free(vals);
    jsql_free(&j);
    free(sql);
    return o;
}

/* ---------------------------------------------------------------- values */
static int need_current(drs *r, const char *src)
{
    if (r->closed) { dao_raise(3420, SRC, "Object invalid or no longer set."); return -1; }
    if (!alive(r, r->cur)) { dao_raise(3021, src, "No current record."); return -1; }
    return 0;
}

vbvar rs_field_value(drs *r, int col, int *err)
{
    *err = 0;
    if (r->closed) { dao_raise(3420, "DAO.Field", "Object invalid or no longer set."); *err = 1; return vb_v_empty(); }
    if (col < 0 || col >= r->ncols) { dao_raise(3265, "DAO.Fields", "Item not found in this collection."); *err = 1; return vb_v_empty(); }
    if (r->editmode && r->ebuf) return vb_vtmp(r->ebuf[col]);
    if (!alive(r, r->cur)) { dao_raise(3021, "DAO.Field", "No current record."); *err = 1; return vb_v_empty(); }
    return vb_vtmp(r->d->r[r->cur].v[col]);
}

/* a value as the type of the column; 0 or the Jet error */
static int convert(rcol *c, vbvar v, vbvar *out)
{
    *out = v;
    if (v.vt == VT_NULL || v.vt == VT_EMPTY) { *out = vb_v_null(); return 0; }
    vb_i32 save = vb_errpend;
    switch (c->type) {
    case dbBoolean: *out = vb_v_bool(vb_v2bool(v) ? -1 : 0); break;
    case dbByte: case dbInteger: case dbLong: {
        double d = vb_v2d(v);
        if (vb_errpend && !save) break;
        double rr = vb_round_even(d);
        if ((c->type == dbByte && (rr < 0 || rr > 255)) || (c->type == dbInteger && (rr < -32768 || rr > 32767)) ||
            (c->type == dbLong && (rr < -2147483648.0 || rr > 2147483647.0))) {
            return 3349;
        }
        *out = c->type == dbByte ? vb_v_byte((vb_u8)rr) : c->type == dbInteger ? vb_v_i((vb_i16)rr) : vb_v_l((vb_i32)rr);
        break; }
    case dbCurrency: *out = vb_v_c(vb_v2c(v)); break;
    case dbSingle: *out = vb_v_f(vb_v2f(v)); break;
    case dbDouble: case dbDecimal: case dbFloat: case dbNumeric: *out = vb_v_d(vb_v2d(v)); break;
    case dbDate: *out = vb_v_date(vb_v2date(v)); break;
    case dbText: case dbMemo: case dbGUID: *out = vb_v_s(vb_v2s(v)); break;
    }
    if (vb_errpend && !save) { vb_err_clear_pending(); return 3421; }
    if (c->type == dbText && c->size > 0 && out->vt == VT_BSTR && out->u.s.len > c->size) return 3163;
    return 0;
}

int rs_field_let(drs *r, int col, vbvar v)
{
    if (r->closed) { dao_raise(3420, "DAO.Field", "Object invalid or no longer set."); return -1; }
    if (col < 0 || col >= r->ncols) { dao_raise(3265, "DAO.Fields", "Item not found in this collection."); return -1; }
    if (!r->editmode) { dao_raise(3020, "DAO.Field", "Update or CancelUpdate without AddNew or Edit."); return -1; }
    rcol *c = &r->cols[col];
    if (!c->updatable || (c->attrs & dbAutoIncrField)) { dao_raise(3164, "DAO.Field", "Field cannot be updated."); return -1; }
    vbvar cv;
    int e = convert(c, v, &cv);
    if (e == 3421) { dao_raise(3421, "DAO.Field", "Data type conversion error."); return -1; }
    if (e == 3349) { dao_raise(3349, "DAO.Field", "Numeric field overflow."); return -1; }
    if (e == 3163) { dao_raise(3163, "DAO.Field", "The field is too small to accept the amount of data you attempted to add.  Try inserting or pasting less data."); return -1; }
    vb_vset(&r->ebuf[col], cv);
    r->edirty[col] = 1;
    return 0;
}

/* ---------------------------------------------------------------- editing */
static int check_updatable(drs *r)
{
    if (r->closed) { dao_raise(3420, SRC, "Object invalid or no longer set."); return -1; }
    if (r->type == dbOpenSnapshot || r->type == dbOpenForwardOnly) { dao_raise(3251, SRC, "Operation is not supported for this type of object."); return -1; }
    if (!r->base || DB(r->db)->ro) { dao_raise(3027, SRC, "Cannot update.  Database or object is read-only."); return -1; }
    return 0;
}

static void edit_begin(drs *r, int add)
{
    edit_clear(r);
    r->ebuf = calloc((size_t)(r->ncols ? r->ncols : 1), sizeof(vbvar));
    r->edirty = calloc((size_t)(r->ncols ? r->ncols : 1), sizeof(int));
    r->editmode = add ? dbEditAdd : dbEditInProgress;
    if (!add) for (int i = 0; i < r->ncols; i++) vb_vset(&r->ebuf[i], r->d->r[r->cur].v[i]);
    else for (int i = 0; i < r->ncols; i++) r->ebuf[i] = vb_v_null();
}

/* the default values of a new record (Jet fills them in at AddNew) */
static void defaults(drs *r)
{
    ddb *d = DB(r->db);
    for (int i = 0; i < r->ncols; i++) {
        rcol *c = &r->cols[i];
        if (c->attrs & dbAutoIncrField) {
            sqlite3_int64 n;
            if (schema_counter_next(d, r->base, &n)) { r->ebuf[i] = vb_v_l((vb_i32)n); r->new_counter = n; }
            continue;
        }
        if (c->type == dbBoolean && c->sfield) r->ebuf[i] = vb_v_bool(0);
        if (!c->defval || !*c->defval) continue;
        jbuf e = { 0 };
        char err[200];
        if (jet_expr(c->defval, &e, 0, err, sizeof err) == 0) {
            jbuf s = { 0 };
            jb_s(&s, "SELECT "); jb_put(&s, e.p, e.n);
            sqlite3_stmt *st = 0;
            if (sqlite3_prepare_v2(d->db, s.p, -1, &st, 0) == SQLITE_OK && sqlite3_step(st) == SQLITE_ROW) {
                vb_i32 m = vb_mark();
                vbvar v = jet_column(st, 0, c->type), cv;
                if (convert(c, v, &cv) == 0) vb_vset(&r->ebuf[i], cv);
                vb_release(m);
            }
            sqlite3_finalize(st);
            jb_free(&s);
        }
        jb_free(&e);
    }
}

/* the row again from its table */
static void reload_row(drs *r, int idx)
{
    ddb *d = DB(r->db);
    rrow *w = &r->d->r[idx];
    jbuf s = { 0 };
    jb_s(&s, "SELECT ");
    int any = 0;
    for (int i = 0; i < r->ncols; i++) {
        if (!r->cols[i].sfield || !r->cols[i].stable || strcasecmp(r->cols[i].stable, r->base)) continue;
        if (any++) jb_s(&s, ", ");
        jb_qid(&s, r->cols[i].sfield, -1);
    }
    if (!any) { jb_free(&s); return; }
    jb_s(&s, " FROM "); jb_qid(&s, r->base, -1); jb_s(&s, " WHERE rowid = ?1");
    sqlite3_stmt *st = 0;
    if (sqlite3_prepare_v2(d->db, s.p, -1, &st, 0) == SQLITE_OK) {
        sqlite3_bind_int64(st, 1, w->rid);
        if (sqlite3_step(st) == SQLITE_ROW) {
            vb_i32 m = vb_mark();
            for (int i = 0, k = 0; i < r->ncols; i++) {
                if (!r->cols[i].sfield || !r->cols[i].stable || strcasecmp(r->cols[i].stable, r->base)) continue;
                vb_vset(&w->v[i], jet_column(st, k++, r->cols[i].type));
            }
            vb_release(m);
        }
        sqlite3_finalize(st);
    }
    jb_free(&s);
}

static int vcmp(vbvar a, vbvar b);
static void sort_by_index(drs *r);

static int do_update(drs *r)
{
    ddb *d = DB(r->db);
    if (!r->editmode) { dao_raise(3020, SRC, "Update or CancelUpdate without AddNew or Edit."); return -1; }
    if (d->remote && d->remote(d->remote_ctx, r, r->editmode == dbEditAdd ? 1 : 2)) return -1;
    jbuf s = { 0 };
    int n = 0;
    int *map = calloc((size_t)(r->ncols ? r->ncols : 1), sizeof(int));
    if (r->editmode == dbEditAdd) {
        jb_s(&s, "INSERT INTO "); jb_qid(&s, r->base, -1); jb_s(&s, " (");
        for (int i = 0; i < r->ncols; i++) {
            rcol *c = &r->cols[i];
            if (!c->sfield || (!c->updatable && !(c->attrs & dbAutoIncrField))) continue;
            if (!r->edirty[i] && !(c->attrs & dbAutoIncrField) && (r->ebuf[i].vt == VT_NULL || r->ebuf[i].vt == VT_EMPTY)) continue;
            if (n) jb_s(&s, ", ");
            jb_qid(&s, c->sfield, -1);
            map[n++] = i;
        }
        if (!n) { jb_s(&s, ") DEFAULT VALUES"); s.p[s.n - 17] = ' '; }
        else {
            jb_s(&s, ") VALUES (");
            for (int k = 0; k < n; k++) jb_f(&s, "%s?%d", k ? ", " : "", k + 1);
            jb_s(&s, ")");
        }
        if (!n) { s.n = 0; jb_s(&s, "INSERT INTO "); jb_qid(&s, r->base, -1); jb_s(&s, " DEFAULT VALUES"); }
    } else {
        jb_s(&s, "UPDATE "); jb_qid(&s, r->base, -1); jb_s(&s, " SET ");
        for (int i = 0; i < r->ncols; i++) {
            if (!r->edirty[i]) continue;
            if (n) jb_s(&s, ", ");
            jb_qid(&s, r->cols[i].sfield, -1); jb_f(&s, " = ?%d", n + 1);
            map[n++] = i;
        }
        jb_f(&s, " WHERE rowid = ?%d", n + 1);
        if (!n) { jb_free(&s); free(map); edit_clear(r); r->lastmod = r->cur; return 0; }
    }
    sqlite3_stmt *st = 0;
    int rc = sqlite3_prepare_v2(d->db, s.p, -1, &st, 0);
    jb_free(&s);
    if (rc != SQLITE_OK) { free(map); dao_raise_sqlite(d, rc, SRC); return -1; }
    for (int k = 0; k < n; k++) jet_bind(st, k + 1, r->ebuf[map[k]], r->cols[map[k]].type);
    if (r->editmode == dbEditInProgress) sqlite3_bind_int64(st, n + 1, r->d->r[r->cur].rid);
    rc = sqlite3_step(st);
    char emsg[1024]; int enr = 0;
    if (rc != SQLITE_DONE) enr = jet_sqlite_error(d->db, rc, emsg, sizeof emsg);
    sqlite3_finalize(st);
    free(map);
    if (rc != SQLITE_DONE) { dao_raise(enr, SRC, "%s", emsg); return -1; }
    r->mirror_ok = 0;
    if (r->editmode == dbEditAdd) {
        sqlite3_int64 rid = sqlite3_last_insert_rowid(d->db);
        int had_all = r->reached >= r->d->n;
        rrow *w = data_add(r->d);
        w->v = calloc((size_t)(r->ncols ? r->ncols : 1), sizeof(vbvar));
        w->rid = rid;
        for (int i = 0; i < r->ncols; i++) vb_vset(&w->v[i], r->ebuf[i]);
        int idx = r->d->n - 1;
        reload_row(r, idx);
        if (had_all) r->reached = r->d->n; else r->added++;
        r->lastmod = idx;
        if (r->type == dbOpenTable && r->nix) {
            sqlite3_int64 cur_rid = alive(r, r->cur) ? r->d->r[r->cur].rid : 0;
            sort_by_index(r);
            for (int i = 0; i < r->d->n; i++) {
                if (r->d->r[i].rid == rid) r->lastmod = i;
                if (cur_rid && r->d->r[i].rid == cur_rid) r->cur = i;
            }
        }
    } else {
        reload_row(r, r->cur);
        r->lastmod = r->cur;
    }
    edit_clear(r);
    return 0;
}

static int do_delete(drs *r)
{
    if (check_updatable(r) || need_current(r, SRC)) return -1;
    ddb *d = DB(r->db);
    if (d->remote && d->remote(d->remote_ctx, r, 3)) return -1;
    jbuf s = { 0 };
    jb_s(&s, "DELETE FROM "); jb_qid(&s, r->base, -1); jb_s(&s, " WHERE rowid = ?1");
    sqlite3_stmt *st = 0;
    int rc = sqlite3_prepare_v2(d->db, s.p, -1, &st, 0);
    jb_free(&s);
    if (rc != SQLITE_OK) { dao_raise_sqlite(d, rc, SRC); return -1; }
    sqlite3_bind_int64(st, 1, r->d->r[r->cur].rid);
    rc = sqlite3_step(st);
    char emsg[1024]; int enr = 0;
    if (rc != SQLITE_DONE) enr = jet_sqlite_error(d->db, rc, emsg, sizeof emsg);
    sqlite3_finalize(st);
    if (rc != SQLITE_DONE) { dao_raise(enr, SRC, "%s", emsg); return -1; }
    r->d->r[r->cur].del = 1;
    edit_clear(r);
    r->mirror_ok = 0;
    return 0;
}

/* ---------------------------------------------------------------- moving */
static int count(drs *r)
{
    if (r->type == dbOpenTable) return n_alive(r);
    int k = 0;
    for (int i = 0; i < r->reached && i < r->d->n; i++) if (!r->d->r[i].del) k++;
    return k + r->added;
}
static int abs_pos(drs *r)
{
    if (!alive(r, r->cur)) return -1;
    int k = 0;
    for (int i = 0; i < r->cur; i++) if (!r->d->r[i].del) k++;
    return k;
}
static void move_to(drs *r, int i)
{
    if (r->editmode) edit_clear(r);
    r->cur = i;
    reach(r);
    /* a dynaset reads the record again when it comes to it: what other connections changed (Jet; VB6 oracle
     * f_dfmt: the Data control sees a change of the ADO Data control) */
    if (r->type == dbOpenDynaset && r->base && alive(r, i)) reload_row(r, i);
}
static int do_move(drs *r, int n)
{
    if (r->closed) { dao_raise(3420, SRC, "Object invalid or no longer set."); return -1; }
    if (n_alive(r) == 0) { dao_raise(3021, SRC, "No current record."); return -1; }
    if (n == 0) { if (!alive(r, r->cur)) { dao_raise(3021, SRC, "No current record."); return -1; } return 0; }
    if (n > 0 && is_eof(r)) { dao_raise(3021, SRC, "No current record."); return -1; }
    if (n < 0 && is_bof(r)) { dao_raise(3021, SRC, "No current record."); return -1; }
    int i = r->cur;
    if (n > 0) {
        while (n > 0) { i++; if (i >= r->d->n) break; if (!r->d->r[i].del) n--; }
        if (i >= r->d->n) { if (r->editmode) edit_clear(r); r->cur = r->d->n; r->reached = r->d->n; return 0; }
    } else {
        while (n < 0) { i--; if (i < 0) break; if (!r->d->r[i].del) n++; }
        if (i < 0) { if (r->editmode) edit_clear(r); r->cur = -1; return 0; }
    }
    move_to(r, i);
    return 0;
}

static vbvar bookmark_of(int idx)
{
    vbarr a; memset(&a, 0, sizeof a);
    vb_adim(&a, 1, VBA_PLAIN, 0, 1, 0, 3);
    vb_u32 v = (vb_u32)(idx + 1);
    memcpy(a.data, &v, 4);
    vbvar r = vb_v_arr(&a, VT_UI1);
    vb_afree(&a);
    return r;
}
static int bookmark_row(drs *r, vbvar v)
{
    if (!(v.vt & VT_ARRAY)) return -1;
    vb_i32 save = vb_errpend;
    vbarr *a = vb_v_arrptr(v);
    if (vb_errpend && !save) { vb_err_clear_pending(); return -1; }
    if (!a || !a->data || a->esz != 1 || a->n < 4) return -1;
    vb_u32 x; memcpy(&x, a->data, 4);
    int i = (int)x - 1;
    return i >= 0 && i < r->d->n ? i : -1;
}

/* ---------------------------------------------------------------- values compared as Jet does */
static int numeric_vt(int vt) { return vt == VT_I2 || vt == VT_I4 || vt == VT_R4 || vt == VT_R8 || vt == VT_CY || vt == VT_BOOL || vt == VT_UI1 || vt == VT_DATE; }
static int vcmp(vbvar a, vbvar b)
{
    int an = a.vt == VT_NULL || a.vt == VT_EMPTY, bn = b.vt == VT_NULL || b.vt == VT_EMPTY;
    if (an || bn) return an && bn ? 0 : an ? -1 : 1;
    vb_i32 m = vb_mark();
    int r;
    if (numeric_vt(a.vt) && numeric_vt(b.vt)) {
        double x = vb_v2d(a), y = vb_v2d(b);
        r = x < y ? -1 : x > y;
    } else if (numeric_vt(a.vt) && b.vt == VT_BSTR) {
        vb_i32 save = vb_errpend;
        double y = vb_v2d(b);
        if (vb_errpend && !save) { vb_err_clear_pending(); r = -1; }
        else { double x = vb_v2d(a); r = x < y ? -1 : x > y; }
    } else if (a.vt == VT_BSTR && numeric_vt(b.vt)) {
        r = -vcmp(b, a);
    } else {
        vbs x = vb_v2s(a), y = vb_v2s(b);
        r = vb_scmp(x, y, 1);
        r = r < 0 ? -1 : r > 0;
    }
    vb_release(m);
    return r;
}

static drs *sort_rs;
static int sort_cmp(const void *pa, const void *pb)
{
    const rrow *a = pa, *b = pb;
    drs *r = sort_rs;
    for (int k = 0; k < r->nix; k++) {
        int c = vcmp(a->v[r->ixcols[k]], b->v[r->ixcols[k]]);
        if (c) return r->ixdesc[k] ? -c : c;
    }
    return a->rid < b->rid ? -1 : a->rid > b->rid;
}
static void sort_by_index(drs *r)
{
    sort_rs = r;
    qsort(r->d->r, (size_t)r->d->n, sizeof(rrow), sort_cmp);
    r->mirror_ok = 0;
}

static int set_index(drs *r, const char *name)
{
    if (r->type != dbOpenTable) { dao_raise(3251, SRC, "Operation is not supported for this type of object."); return -1; }
    ddb *d = DB(r->db);
    if (!*name) {
        free(r->index); r->index = 0; r->nix = 0;
        return 0;
    }
    sqlite3_stmt *st = 0;
    if (sqlite3_prepare_v2(d->db, "SELECT name, fields FROM MSysLB_Indexes WHERE tbl = ?1 AND name = ?2", -1, &st, 0) != SQLITE_OK) return -1;
    sqlite3_bind_text(st, 1, r->base ? r->base : r->source, -1, SQLITE_STATIC);
    sqlite3_bind_text(st, 2, name, -1, SQLITE_STATIC);
    if (sqlite3_step(st) != SQLITE_ROW) {
        sqlite3_finalize(st);
        dao_raise(3800, SRC, "'%s' is not an index in this table.", name);
        return -1;
    }
    free(r->index); r->index = strdup((const char *)sqlite3_column_text(st, 0));
    char *fl = strdup((const char *)sqlite3_column_text(st, 1)), *save = 0;
    sqlite3_finalize(st);
    free(r->ixcols); free(r->ixdesc);
    r->ixcols = 0; r->ixdesc = 0; r->nix = 0;
    for (char *p = strtok_r(fl, ";", &save); p; p = strtok_r(0, ";", &save)) {
        int desc = *p == '-';
        if (*p == '+' || *p == '-') p++;
        int c = col_index(r, p);
        if (c < 0) continue;
        r->ixcols = realloc(r->ixcols, sizeof(int) * (size_t)(r->nix + 1));
        r->ixdesc = realloc(r->ixdesc, sizeof(int) * (size_t)(r->nix + 1));
        r->ixcols[r->nix] = c; r->ixdesc[r->nix] = desc; r->nix++;
    }
    free(fl);
    if (r->editmode) edit_clear(r);
    sort_by_index(r);
    r->cur = first_alive(r);
    return 0;
}

static int do_seek(drs *r, vb_i32 argc, vbvar *argv)
{
    if (r->closed) { dao_raise(3420, SRC, "Object invalid or no longer set."); return -1; }
    if (r->type != dbOpenTable) { dao_raise(3251, SRC, "Operation is not supported for this type of object."); return -1; }
    if (!r->index) { dao_raise(3019, SRC, "Operation invalid without a current index."); return -1; }
    if (argc < 2) { vb_raise(450); return -1; }
    char *op = dao_str(argv[0]);
    while (*op == ' ') op++;
    int kind;
    if (!strcmp(op, "=")) kind = 0;
    else if (!strcmp(op, ">=")) kind = 1;
    else if (!strcmp(op, ">")) kind = 2;
    else if (!strcmp(op, "<=")) kind = 3;
    else if (!strcmp(op, "<")) kind = 4;
    else { dao_raise(3001, SRC, "Invalid argument."); return -1; }
    int nk = argc - 1 < r->nix ? argc - 1 : r->nix;
    if (r->editmode) edit_clear(r);
    int found = -1;
    if (kind <= 2) {
        for (int i = 0; i < r->d->n && found < 0; i++) {
            if (r->d->r[i].del) continue;
            int c = 0;
            for (int k = 0; k < nk && !c; k++) { c = vcmp(r->d->r[i].v[r->ixcols[k]], argv[k + 1]); if (r->ixdesc[k]) c = -c; }
            if ((kind == 0 && c == 0) || (kind == 1 && c >= 0) || (kind == 2 && c > 0)) found = i;
        }
    } else {
        for (int i = r->d->n - 1; i >= 0 && found < 0; i--) {
            if (r->d->r[i].del) continue;
            int c = 0;
            for (int k = 0; k < nk && !c; k++) { c = vcmp(r->d->r[i].v[r->ixcols[k]], argv[k + 1]); if (r->ixdesc[k]) c = -c; }
            if ((kind == 3 && c <= 0) || (kind == 4 && c < 0)) found = i;
        }
    }
    r->nomatch = found < 0;
    if (found >= 0) move_to(r, found);
    return 0;
}

/* ---------------------------------------------------------------- Find: the rows in a memory database */
static int build_mirror(drs *r)
{
    if (r->mirror && r->mirror_ok) return 0;
    mirror_drop(r);
    if (sqlite3_open(":memory:", &r->mirror) != SQLITE_OK) return -1;
    jet_register(r->mirror);
    jbuf s = { 0 };
    jb_s(&s, "CREATE TABLE rs (_pos INTEGER PRIMARY KEY");
    for (int i = 0; i < r->ncols; i++) {
        jb_s(&s, ", "); jb_qid(&s, r->cols[i].name, -1);
        if (r->cols[i].type == dbText || r->cols[i].type == dbMemo) jb_s(&s, " TEXT COLLATE NOCASE");
        if (r->cols[i].sfield && strcasecmp(r->cols[i].sfield, r->cols[i].name)) {
            /* the field's own name works too when it is shown as T.Name */
        }
    }
    jb_s(&s, ")");
    int rc = sqlite3_exec(r->mirror, s.p, 0, 0, 0);
    s.n = 0;
    if (rc == SQLITE_OK) {
        jb_s(&s, "INSERT INTO rs VALUES (?1");
        for (int i = 0; i < r->ncols; i++) jb_f(&s, ", ?%d", i + 2);
        jb_s(&s, ")");
        sqlite3_stmt *st = 0;
        rc = sqlite3_prepare_v2(r->mirror, s.p, -1, &st, 0);
        if (rc == SQLITE_OK) {
            sqlite3_exec(r->mirror, "BEGIN", 0, 0, 0);
            for (int i = 0; i < r->d->n; i++) {
                if (r->d->r[i].del) continue;
                vb_i32 m = vb_mark();
                sqlite3_bind_int(st, 1, i);
                for (int k = 0; k < r->ncols; k++) jet_bind(st, k + 2, r->d->r[i].v[k], 0);
                sqlite3_step(st);
                sqlite3_reset(st);
                vb_release(m);
            }
            sqlite3_exec(r->mirror, "COMMIT", 0, 0, 0);
            sqlite3_finalize(st);
        }
    }
    jb_free(&s);
    r->mirror_ok = rc == SQLITE_OK;
    return rc == SQLITE_OK ? 0 : -1;
}

static int criteria_error(sqlite3 *db, int rc, const char *crit)
{
    const char *m = sqlite3_errmsg(db);
    if (rc == SQLITE_ERROR && !strncmp(m, "no such column: ", 16)) {
        char nm[256];
        const char *n = m + 16;
        if (*n == '"') { const char *e = strchr(n + 1, '"'); snprintf(nm, sizeof nm, "%.*s", e ? (int)(e - n - 1) : 0, n + 1); }
        else { const char *dot = strrchr(n, '.'); snprintf(nm, sizeof nm, "%s", dot ? dot + 1 : n); }
        dao_raise(3070, SRC, "The Microsoft Jet database engine does not recognize '%s' as a valid field name or expression.", nm);
    } else if (m[0] == '#') {
        int num = atoi(m + 1);
        const char *t = strchr(m, ' ');
        dao_raise(num, SRC, "%s", t ? t + 1 : "");
    } else dao_raise(3077, SRC, "Syntax error in expression.");
    (void)crit;
    return -1;
}

static int do_find(drs *r, int how, vbvar crit)
{
    if (r->closed) { dao_raise(3420, SRC, "Object invalid or no longer set."); return -1; }
    if (r->type == dbOpenTable || r->type == dbOpenForwardOnly) { dao_raise(3251, SRC, "Operation is not supported for this type of object."); return -1; }
    char *c = dao_str(crit);
    jbuf e = { 0 };
    char err[400];
    int en = jet_expr(c, &e, 1, err, sizeof err);
    if (en) { jb_free(&e); dao_raise(en == 3075 ? 3077 : en, SRC, "%s", en == 3075 ? "Syntax error in expression." : err); return -1; }
    if (build_mirror(r)) { jb_free(&e); dao_raise(3001, SRC, "Invalid argument."); return -1; }
    jbuf s = { 0 };
    jb_s(&s, "SELECT _pos FROM rs WHERE (");
    jb_put(&s, e.p, e.n);
    jb_s(&s, ")");
    int start = r->cur;
    if (how == 2) jb_f(&s, " AND _pos > %d ORDER BY _pos LIMIT 1", start);
    else if (how == 3) jb_f(&s, " AND _pos < %d ORDER BY _pos DESC LIMIT 1", start);
    else if (how == 0) jb_s(&s, " ORDER BY _pos LIMIT 1");
    else jb_s(&s, " ORDER BY _pos DESC LIMIT 1");
    jb_free(&e);
    sqlite3_stmt *st = 0;
    int rc = sqlite3_prepare_v2(r->mirror, s.p, -1, &st, 0);
    jb_free(&s);
    if (rc != SQLITE_OK) return criteria_error(r->mirror, rc, c);
    rc = sqlite3_step(st);
    int found = -1;
    if (rc == SQLITE_ROW) found = sqlite3_column_int(st, 0);
    else if (rc != SQLITE_DONE) { sqlite3_finalize(st); return criteria_error(r->mirror, rc, c); }
    sqlite3_finalize(st);
    if (r->editmode) edit_clear(r);
    r->nomatch = found < 0;
    if (found >= 0) move_to(r, found);
    else if (how == 0) { int f = first_alive(r); if (f >= 0) move_to(r, f); }
    else if (how == 1) { int l = last_alive(r); if (l >= 0) move_to(r, l); }
    return 0;
}

/* ---------------------------------------------------------------- Clone, Requery, OpenRecordset */
static vbobj do_clone(drs *r)
{
    drs *c = (drs *)vb_onew(&vb_cls_dao_recordset);
    c->db = r->db; vb_oaddref(r->db);
    c->type = r->type; c->options = r->options; c->lockedits = r->lockedits;
    c->source = sdup(r->source); c->sqlite = sdup(r->sqlite); c->base = sdup(r->base);
    c->ncols = r->ncols;
    c->cols = calloc((size_t)(r->ncols ? r->ncols : 1), sizeof(rcol));
    for (int i = 0; i < r->ncols; i++) {
        c->cols[i] = r->cols[i];
        c->cols[i].name = sdup(r->cols[i].name); c->cols[i].stable = sdup(r->cols[i].stable);
        c->cols[i].sfield = sdup(r->cols[i].sfield); c->cols[i].defval = sdup(r->cols[i].defval);
    }
    c->d = r->d; r->d->refs++;
    c->cur = first_alive(r);
    c->reached = r->d->n;
    c->lastmod = -1;
    vb_sset(&c->filter, vb_sfromutf8("")); vb_sset(&c->sort, vb_sfromutf8(""));
    if (r->index) {
        c->index = strdup(r->index);
        c->nix = r->nix;
        c->ixcols = malloc(sizeof(int) * (size_t)(r->nix ? r->nix : 1)); memcpy(c->ixcols, r->ixcols, sizeof(int) * (size_t)r->nix);
        c->ixdesc = malloc(sizeof(int) * (size_t)(r->nix ? r->nix : 1)); memcpy(c->ixdesc, r->ixdesc, sizeof(int) * (size_t)r->nix);
    }
    make_fields(c);
    coll_add(DB(r->db)->rss, (vbobj)c);
    return (vbobj)c;
}

static int do_requery(drs *r)
{
    if (r->closed) { dao_raise(3420, SRC, "Object invalid or no longer set."); return -1; }
    ddb *d = DB(r->db);
    edit_clear(r);
    mirror_drop(r);
    data_release(r->d, r->ncols);
    r->d = calloc(1, sizeof(rdata));
    r->d->refs = 1;
    /* the parameters again from the QueryDef */
    vbvar *vals = 0; int nb = 0;
    if (r->parent_qd) {
        dqd *q = (dqd *)vb_oid(r->parent_qd);
        const char *params[64]; int np = 0;
        static char buf[64][256];
        for (int k = 0; k < q->params->n && np < 63; k++) {
            const char *pn = ((dpar *)q->params->items[k])->name;
            size_t pl = strlen(pn);
            (void)pl; dao_param_plain(pn, buf[np], 256); pn = buf[np];
            params[np++] = pn;
        }
        params[np] = 0;
        jsql j;
        if (jet_translate(q->sql, &j, np ? params : 0) == 0) { param_values(&j, r->parent_qd, &vals); nb = j.nbinds; }
        jsql_free(&j);
    }
    int rc = fill(r, d, r->sqlite, nb, vals);
    free(vals);
    if (rc != SQLITE_OK) { dao_raise_sqlite(d, rc, SRC); return -1; }
    if (r->nix) sort_by_index(r);
    r->cur = first_alive(r);
    r->reached = 0; r->added = 0;
    if (r->cur >= 0) reach(r);
    if (r->type == dbOpenTable || sorted_sql(d, r->base, r->sqlite)) r->reached = r->d->n;
    return 0;
}

/* Sort "Born DESC, Name" as SQLite ORDER BY */
static int sort_sql(const char *sort, jbuf *out, char *err, int cap)
{
    const char *p = sort;
    int first = 1;
    while (*p) {
        const char *b = p;
        int depth = 0; char q = 0;
        while (*p && (depth || q || *p != ',')) {
            if (q) { if (*p == q) q = 0; }
            else if (*p == '\'' || *p == '"') q = *p;
            else if (*p == '[') { while (*p && *p != ']') p++; if (!*p) break; }
            else if (*p == '(') depth++;
            else if (*p == ')') depth--;
            p++;
        }
        char part[512]; snprintf(part, sizeof part, "%.*s", (int)(p - b), b);
        if (*p == ',') p++;
        int n = (int)strlen(part);
        while (n && isspace((unsigned char)part[n - 1])) part[--n] = 0;
        int desc = 0;
        if (n > 5 && !strcasecmp(part + n - 5, " desc")) { desc = 1; part[n - 5] = 0; }
        else if (n > 4 && !strcasecmp(part + n - 4, " asc")) part[n - 4] = 0;
        jbuf e = { 0 };
        int en = jet_expr(part, &e, 0, err, cap);
        if (en) { jb_free(&e); return en; }
        if (!first) jb_s(out, ", ");
        jb_put(out, e.p, e.n);
        if (desc) jb_s(out, " DESC");
        jb_free(&e);
        first = 0;
    }
    return 0;
}

static vbobj do_openrs(drs *r, int type)
{
    if (r->closed) { dao_raise(3420, SRC, "Object invalid or no longer set."); return 0; }
    ddb *d = DB(r->db);
    jbuf s = { 0 };
    char err[400];
    jb_s(&s, "SELECT * FROM ("); jb_s(&s, r->sqlite); jb_s(&s, ")");
    if (r->filter.len) {
        jbuf e = { 0 };
        int en = jet_expr(vb_stoutf8(r->filter), &e, 1, err, sizeof err);
        if (en) { jb_free(&e); jb_free(&s); dao_raise(en, SRC, "%s", err); return 0; }
        jb_s(&s, " WHERE "); jb_put(&s, e.p, e.n);
        jb_free(&e);
    }
    if (r->sort.len) {
        jbuf o = { 0 };
        int en = sort_sql(vb_stoutf8(r->sort), &o, err, sizeof err);
        if (en) { jb_free(&o); jb_free(&s); dao_raise(en, SRC, "%s", err); return 0; }
        jb_s(&s, " ORDER BY "); jb_put(&s, o.p, o.n);
        jb_free(&o);
    }
    if (!type) type = r->type == dbOpenTable ? dbOpenDynaset : r->type;
    vbvar *vals = 0; int nb = 0;
    if (r->parent_qd) {
        dqd *q = (dqd *)vb_oid(r->parent_qd);
        const char *params[64]; int np = 0;
        static char buf[64][256];
        for (int k = 0; k < q->params->n && np < 63; k++) {
            const char *pn = ((dpar *)q->params->items[k])->name;
            size_t pl = strlen(pn);
            (void)pl; dao_param_plain(pn, buf[np], 256); pn = buf[np];
            params[np++] = pn;
        }
        params[np] = 0;
        jsql j;
        if (jet_translate(q->sql, &j, np ? params : 0) == 0) { param_values(&j, r->parent_qd, &vals); nb = j.nbinds; }
        jsql_free(&j);
    }
    jctype *ct = calloc((size_t)(r->ncols ? r->ncols : 1), sizeof(jctype));
    for (int i = 0; i < r->ncols; i++) { snprintf(ct[i].name, sizeof ct[i].name, "%s", r->cols[i].name); ct[i].type = r->cols[i].type; }
    pend_ct = ct; pend_nct = r->ncols;
    vbobj o = rs_open_sqlite(d, s.p, type == dbOpenSnapshot ? 0 : r->base, type, r->source, nb, vals);
    pend_ct = 0; pend_nct = 0;
    free(ct);
    free(vals);
    jb_free(&s);
    if (o) {
        drs *n = (drs *)o;
        for (int i = 0; i < n->ncols && i < r->ncols; i++) n->cols[i].updatable = r->cols[i].updatable && type != dbOpenSnapshot;
    }
    return o;
}

/* GetRows: a Variant array (field, row) from the current record on */
static vbvar do_getrows(drs *r, int n)
{
    if (r->closed) { dao_raise(3420, SRC, "Object invalid or no longer set."); return vb_v_empty(); }
    int rows = 0;
    int idx[4096];
    int i = r->cur;
    if (i < 0) i = first_alive(r);
    for (; i >= 0 && i < r->d->n && rows < n && rows < 4096; i++) if (!r->d->r[i].del) idx[rows++] = i;
    vbarr a; memset(&a, 0, sizeof a);
    if (!rows) { vb_aempty(&a, sizeof(vbvar), VBA_VAR); }
    else {
        vb_adim(&a, sizeof(vbvar), VBA_VAR, 0, 2, 0, r->ncols - 1, 0, rows - 1);
        for (int k = 0; k < rows; k++)
            for (int c = 0; c < r->ncols; c++) vb_vset((vbvar *)vb_aelem(&a, 2, c, k), r->d->r[idx[k]].v[c]);
    }
    vbvar v = vb_v_arr(&a, VT_VARIANT);
    vb_afree(&a);
    if (rows) {
        int last = idx[rows - 1];
        int nx = last + 1;
        while (nx < r->d->n && r->d->r[nx].del) nx++;
        r->cur = nx;
        if (nx < r->d->n) reach(r); else r->reached = r->d->n;
    }
    return v;
}

/* ---------------------------------------------------------------- a Data control watching */
enum { ACT_CANCEL = 0, ACT_FIRST = 1, ACT_PREV, ACT_NEXT, ACT_LAST, ACT_ADDNEW, ACT_UPDATE, ACT_DELETE, ACT_FIND,
       ACT_BOOKMARK, ACT_CLOSE, ACT_UNLOAD };
/* Validate: 0 to go on, 1 if the program cancelled the action */
static int validate(drs *r, int action)
{
    if (!r->watch) return 0;
    int cancel = 0;
    vbobj keep = (vbobj)r;
    vb_oaddref(keep);
    r->watch(r->watch_ctx, 1, action, &cancel);
    vb_orelease(keep);
    return cancel || vb_errpend;
}
static void reposition(drs *r)
{
    if (!r->watch) return;
    vbobj keep = (vbobj)r;
    vb_oaddref(keep);
    r->watch(r->watch_ctx, 2, 0, 0);
    vb_orelease(keep);
}
/* the rows changed and the recordset did not move (Delete): the grids bound to the Data control read them */
static void rows_changed(drs *r, int action)
{
    if (!r->watch) return;
    vbobj keep = (vbobj)r;
    vb_oaddref(keep);
    r->watch(r->watch_ctx, 3, action, 0);
    vb_orelease(keep);
}

/* ---------------------------------------------------------------- members */
static int name_is(const char *a, const char *b) { return !strcmp(a, b); }

vbvar rs_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    drs *r = RS(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    dao_src = SRC;
    /* rs("Name"), rs!Name, rs(0): the default member is Fields */
    if (name_is(name, "fields") || name_is(name, "")) {
        if (r->closed) { dao_raise(3420, SRC, "Object invalid or no longer set."); return vb_v_empty(); }
        if (argc == 0 || (set && argc == 1)) {
            if (set) { vb_raise(450); return vb_v_empty(); }
            return dao_retobj((vbobj)r->fields);
        }
        vbvar key = argv[0];
        int col;
        if (key.vt == VT_BSTR) col = col_index(r, dao_str(key));
        else { col = vb_v2l(key); if (col >= r->ncols) col = -1; }
        if (col < 0) { dao_raise(3265, "DAO.Fields", "Item not found in this collection."); return vb_v_empty(); }
        if (set) { rs_field_let(r, col, vb_v_letval(argv[argc - 1])); return vb_v_empty(); }
        if (argc == 1) return dao_retobj(r->fields->items[col]);
        vb_raise(450);
        return vb_v_empty();
    }
    if (name_is(name, "close")) { rs_close(r); return vb_v_empty(); }
    if (r->closed && !name_is(name, "name")) { dao_raise(3420, SRC, "Object invalid or no longer set."); return vb_v_empty(); }
    if (name_is(name, "bof")) return vb_v_bool(is_bof(r) ? -1 : 0);
    if (name_is(name, "eof")) return vb_v_bool(is_eof(r) ? -1 : 0);
    if (name_is(name, "recordcount")) return vb_v_l(count(r));
    if (name_is(name, "type")) return vb_v_i((vb_i16)r->type);
    if (name_is(name, "name")) {
        const char *n = r->source ? r->source : "";
        char buf[257]; snprintf(buf, sizeof buf, "%s", n);
        return dao_vstr(buf);
    }
    if (name_is(name, "updatable")) return vb_v_bool(r->base && r->type != dbOpenSnapshot && r->type != dbOpenForwardOnly && !DB(r->db)->ro ? -1 : 0);
    if (name_is(name, "editmode")) return vb_v_i((vb_i16)r->editmode);
    if (name_is(name, "nomatch")) return vb_v_bool(r->nomatch ? -1 : 0);
    if (name_is(name, "bookmarkable")) return vb_v_bool(-1);
    if (name_is(name, "transactions")) return vb_v_bool(r->type != dbOpenSnapshot ? -1 : 0);
    if (name_is(name, "restartable")) return vb_v_bool(-1);
    if (name_is(name, "lockedits")) { if (set) { r->lockedits = vb_v2bool(argv[argc - 1]) != 0; return vb_v_empty(); } return vb_v_bool(r->lockedits ? -1 : 0); }
    if (name_is(name, "validationrule") || name_is(name, "validationtext")) return dao_vstr("");
    if (name_is(name, "cachesize")) return vb_v_l(0);
    if (name_is(name, "stillexecuting")) return vb_v_bool(0);
    if (name_is(name, "absoluteposition")) {
        if (set) {
            vb_i32 p = vb_v2l(argv[argc - 1]);
            int k = -1;
            if (p >= 0) for (int i = 0; i < r->d->n; i++) if (!r->d->r[i].del && ++k == p) { move_to(r, i); return vb_v_empty(); }
            dao_raise(3001, SRC, "Invalid argument.");
            return vb_v_empty();
        }
        return vb_v_l(abs_pos(r));
    }
    if (name_is(name, "percentposition")) {
        int n = count(r);
        if (set) {
            double pc = vb_v2d(argv[argc - 1]);
            int tot = n_alive(r), target = (int)(pc * tot / 100.0);
            if (target >= tot) target = tot - 1;
            int k = -1;
            for (int i = 0; i < r->d->n; i++) if (!r->d->r[i].del && ++k == target) { move_to(r, i); break; }
            return vb_v_empty();
        }
        int p = abs_pos(r);
        return vb_v_f(p < 0 || n == 0 ? 0.0f : (float)((p + 1) * 100.0 / n));
    }
    if (name_is(name, "bookmark")) {
        if (set) {
            int i = bookmark_row(r, argv[argc - 1]);
            if (i < 0 || r->d->r[i].del) { dao_raise(3159, SRC, "Not a valid bookmark."); return vb_v_empty(); }
            if (validate(r, ACT_BOOKMARK)) return vb_v_empty();
            move_to(r, i);
            reposition(r);
            return vb_v_empty();
        }
        if (!alive(r, r->cur)) { dao_raise(3021, SRC, "No current record."); return vb_v_empty(); }
        return bookmark_of(r->cur);
    }
    if (name_is(name, "lastmodified")) {
        if (r->lastmod < 0) { dao_raise(3021, SRC, "No current record."); return vb_v_empty(); }
        return bookmark_of(r->lastmod);
    }
    if (name_is(name, "index")) {
        if (set) { set_index(r, dao_str(argv[argc - 1])); return vb_v_empty(); }
        if (r->type != dbOpenTable) { dao_raise(3251, SRC, "Operation is not supported for this type of object."); return vb_v_empty(); }
        return dao_vstr(r->index ? r->index : "");
    }
    if (name_is(name, "filter") || name_is(name, "sort")) {
        vbs *t = name[0] == 'f' ? &r->filter : &r->sort;
        if (set) { vbs s = vb_v2s(argv[argc - 1]); if (!vb_errpend) vb_sset(t, s); return vb_v_empty(); }
        return vb_v_s(vb_scopy(*t));
    }
    if (name_is(name, "datecreated") || name_is(name, "lastupdated")) return vb_v_date(vb_now());
    /* methods */
    if (name_is(name, "movefirst") || name_is(name, "movelast")) {
        if (n_alive(r) == 0) { dao_raise(3021, SRC, "No current record."); return vb_v_empty(); }
        if (validate(r, name[4] == 'f' ? ACT_FIRST : ACT_LAST)) return vb_v_empty();
        move_to(r, name[4] == 'f' ? first_alive(r) : last_alive(r));
        if (name[4] == 'l') r->reached = r->d->n;
        reposition(r);
        return vb_v_empty();
    }
    if (name_is(name, "movenext") || name_is(name, "moveprevious")) {
        int fwd = name[4] == 'n';
        if (r->watch && ((fwd && is_eof(r)) || (!fwd && is_bof(r)) || n_alive(r) == 0)) { do_move(r, fwd ? 1 : -1); return vb_v_empty(); }
        if (validate(r, fwd ? ACT_NEXT : ACT_PREV)) return vb_v_empty();
        if (do_move(r, fwd ? 1 : -1) == 0) reposition(r);
        return vb_v_empty();
    }
    if (name_is(name, "move")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        vb_i32 n = vb_v2l(argv[0]);
        if (argc > 1 && !dao_missing(argc, argv, 1)) {
            int i = bookmark_row(r, argv[1]);
            if (i < 0) { dao_raise(3159, SRC, "Not a valid bookmark."); return vb_v_empty(); }
            r->cur = i;
        }
        if (validate(r, n > 0 ? ACT_NEXT : ACT_PREV)) return vb_v_empty();
        if (do_move(r, n) == 0) reposition(r);
        return vb_v_empty();
    }
    if (name_is(name, "addnew")) {
        if (check_updatable(r)) return vb_v_empty();
        if (validate(r, ACT_ADDNEW)) return vb_v_empty();
        edit_begin(r, 1);
        defaults(r);
        reposition(r);
        return vb_v_empty();
    }
    if (name_is(name, "edit")) {
        if (check_updatable(r) || need_current(r, SRC)) return vb_v_empty();
        edit_begin(r, 0);
        return vb_v_empty();
    }
    if (name_is(name, "update")) {
        int em = r->editmode;
        if (em && validate(r, ACT_UPDATE)) return vb_v_empty();
        do_update(r);
        /* AddNew: back to the record before it (VB6 oracle f_data); Edit: the same record again, its new values
         * (VB6 oracle f_dbgrid: Validate 6, then Reposition) */
        if (em && !vb_errpend) reposition(r);
        return vb_v_empty();
    }
    if (name_is(name, "cancelupdate")) { edit_clear(r); return vb_v_empty(); }
    if (name_is(name, "delete")) {
        if (alive(r, r->cur) && validate(r, ACT_DELETE)) return vb_v_empty();
        if (do_delete(r) == 0) rows_changed(r, ACT_DELETE);
        return vb_v_empty();
    }
    if (name_is(name, "findfirst") || name_is(name, "findlast") || name_is(name, "findnext") || name_is(name, "findprevious")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int how = name[4] == 'f' ? 0 : name[4] == 'l' ? 1 : name[4] == 'n' ? 2 : 3;
        if (validate(r, ACT_FIND)) return vb_v_empty();
        int before = r->cur;
        if (do_find(r, how, argv[0]) == 0 && (r->cur != before || !r->nomatch)) reposition(r);
        return vb_v_empty();
    }
    if (name_is(name, "seek")) { do_seek(r, argc, argv); return vb_v_empty(); }
    if (name_is(name, "clone")) return dao_retobj(do_clone(r));
    if (name_is(name, "requery")) { if (do_requery(r) == 0) reposition(r); return vb_v_empty(); }
    if (name_is(name, "openrecordset")) {
        vbobj n = do_openrs(r, argc > 0 && !dao_missing(argc, argv, 0) ? vb_v2l(argv[0]) : 0);
        return n ? dao_retobj(n) : vb_v_empty();
    }
    if (name_is(name, "getrows")) return do_getrows(r, argc > 0 && !dao_missing(argc, argv, 0) ? vb_v2l(argv[0]) : 1);
    if (name_is(name, "fillcache") || name_is(name, "cancel")) return vb_v_empty();
    if (name_is(name, "properties")) { dao_raise(3251, SRC, "Operation is not supported for this type of object."); return vb_v_empty(); }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- the rows for the grids and the lists (phase 7.6) */
/* the k-th record that is not deleted: its index in the rows, -1 */
static int row_index(drs *r, int k)
{
    if (k < 0) return -1;
    for (int i = 0; i < r->d->n; i++) if (!r->d->r[i].del && k-- == 0) return i;
    return -1;
}
int dao_rs_shape(vbobj o, int *ncols, int *cur)
{
    drs *r = RS(o);
    if (r->closed) { if (ncols) *ncols = 0; if (cur) *cur = -1; return -1; }
    int n = n_alive(r);
    if (ncols) *ncols = r->ncols;
    if (cur) *cur = r->editmode == dbEditAdd ? n : alive(r, r->cur) ? abs_pos(r) : -1;
    return n;
}
vbvar dao_rs_cell(vbobj o, int row, int col)
{
    drs *r = RS(o);
    if (r->closed || col < 0 || col >= r->ncols) return vb_v_empty();
    if (row < 0) return r->editmode && r->ebuf ? vb_vtmp(r->ebuf[col]) : vb_v_empty();
    int i = row_index(r, row);
    return i < 0 ? vb_v_empty() : vb_vtmp(r->d->r[i].v[col]);
}
const char *dao_rs_colname(vbobj o, int col)
{
    drs *r = RS(o);
    return col >= 0 && col < r->ncols && r->cols[col].name ? r->cols[col].name : "";
}
int dao_rs_coltype(vbobj o, int col)
{
    drs *r = RS(o);
    return col >= 0 && col < r->ncols ? r->cols[col].type : 0;
}
vbvar dao_rs_bookmark(vbobj o, int row)
{
    drs *r = RS(o);
    int i = r->closed ? -1 : row_index(r, row);
    return i < 0 ? vb_v_null() : bookmark_of(i);
}
int dao_rs_bmrow(vbobj o, vbvar bm)
{
    drs *r = RS(o);
    if (r->closed) return -1;
    int i = bookmark_row(r, bm);
    if (i < 0 || r->d->r[i].del) return -1;
    int k = 0;
    for (int j = 0; j < i; j++) if (!r->d->r[j].del) k++;
    return k;
}
