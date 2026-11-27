/* daoodbc.inc - DAO over ODBC (phase 12.4): a database of an ODBC data source opened by Jet (OpenDatabase with
 * "ODBC;..."), a table of one linked to a Jet database (dbAttachedODBC), [ODBC;...].table in Jet SQL.
 * Included by dao.c.
 *
 * As the runtime's Jet SQL runs on SQLite, each data source has a SQLite file of its own (a temporary one) with a
 * copy of each of its tables that the program uses: their fields and indexes as Jet gives them, the rows read
 * again from the server before each statement that names the table (and before a recordset is opened on it). A
 * change goes to the server: a row of a dynaset (Edit / AddNew / Delete: the row found by the values of all of its
 * columns that are not long, as Jet does it, so a row someone changed meanwhile is a write conflict, 3197); the
 * rows an action query changes in the copy (an INSERT / UPDATE / DELETE of Jet SQL: each row found by its key).
 * Pass-through (dbSQLPassThrough, a QueryDef with an ODBC Connect) goes to the server as it is.
 * Errors as VB6 gives them (VB6 oracle tests/db/d_dao6, the SQLite ODBC driver): DBEngine.Errors has the errors of
 * ODBC first (their native number, Source ODBC.<object>), then Jet's ("ODBC--call failed." 3146 ...).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "odbc.h"
#ifdef _WIN32
#include <process.h>
#else
#include <unistd.h>
#endif

typedef struct otab {
    char *name;                         /* the table on the server (and its copy) */
    char **cols; int *types; int ncols; /* its columns and their DAO types */
    int autocol;                        /* its identity column, or -1 */
    int *iskey;                         /* the primary key */
    int loaded;                         /* the statement (stamp) that read its rows last */
} otab;
typedef struct osrc {
    char *key;                          /* the Connect the program gave */
    char *connect;                      /* "ODBC;" and the connection string the driver completed */
    char *dbname, *dbms;
    char quote;                         /* SQL_IDENTIFIER_QUOTE_CHAR */
    lbo_conn *k;
    char *file;                         /* the SQLite file of the copies */
    ddb *mx;                            /* a connection of its own to it (the schema of the copies) */
    otab *tabs; int ntabs;
    int trans;                          /* BeginTrans depth */
    int stale;                          /* its data source registered again: a new connection for the next open */
} osrc;
static osrc **g_os; static int g_nos;
static int g_stamp;                     /* a statement */
static int g_ptno;                      /* temp.lbpt<n>: the rows of a pass-through query */

static void os_cleanup(void)
{
    for (int i = 0; i < g_nos; i++) {
        osrc *o = g_os[i];
        if (o->k) lbo_disconnect(o->k);
        if (o->mx && o->mx->db) { sqlite3_close_v2(o->mx->db); o->mx->db = 0; }
        if (o->file) vbos_remove(o->file);
    }
}

/* ---------------------------------------------------------------- errors */
static void err_append(int number, const char *desc, const char *src)
{
    derr *e = (derr *)vb_onew(&vb_cls_dao_error);
    e->number = number;
    e->desc = strdup(desc);
    e->src = strdup(src);
    coll_add(err_list(), (vbobj)e);
}
/* the diagnostics of ODBC (Source ODBC.<object>), then Jet's error, raised */
static void odbc_fail(int number, const char *src, const char *fmt, ...)
{
    char msg[1024];
    va_list ap; va_start(ap, fmt); vsnprintf(msg, sizeof msg, fmt, ap); va_end(ap);
    char osrcname[64];
    snprintf(osrcname, sizeof osrcname, "ODBC.%s", !strncmp(src, "DAO.", 4) ? src + 4 : src);
    coll_clear_items(err_list());
    int n = lbo_nerr();
    const lbo_err *e = lbo_errs();
    for (int i = 0; i < n; i++) {
        char m[1100];
        lb_odbc_msg(e[i].msg, m, sizeof m);
        err_append(e[i].native, m, osrcname);
    }
    err_append(number, msg, src);
    vb_raise_full(number, vb_sfromutf8(src), vb_sfromutf8(msg));
}

/* ---------------------------------------------------------------- values */
static vbvar ov_val(lbo_val *v)
{
    switch (v->kind) {
    case LBO_NULL: return vb_v_null();
    case LBO_INT: return v->iv >= -2147483647LL - 1 && v->iv <= 2147483647LL ? vb_v_l((vb_i32)v->iv) : vb_v_d((double)v->iv);
    case LBO_BOOL: return vb_v_bool(v->iv ? -1 : 0);
    case LBO_DBL: return vb_v_d(v->dv);
    case LBO_DEC: return vb_v_d(strtod(v->s ? v->s : "0", 0));
    case LBO_DATE: {
        int y = 0, mo = 0, d = 0, h = 0, mi = 0, s = 0;
        sscanf(v->s, "%d-%d-%d %d:%d:%d", &y, &mo, &d, &h, &mi, &s);
        vb_date dt = (y || mo || d) ? vb_dateserial(y, mo, d) : 0;
        vb_date tm = vb_timeserial(h, mi, s);
        return vb_v_date(dt >= 0 ? dt + tm : dt - tm);
    }
    case LBO_BIN: {
        vbarr *a = vb_atmp();
        if (v->len) { vb_adim(a, 1, VBA_PLAIN, 0, 1, 0, v->len - 1); memcpy(a->data, v->s, (size_t)v->len); }
        else vb_aempty(a, 1, VBA_PLAIN);
        return vb_v_arr(a, VT_UI1);
    }
    default: return vb_v_s(vb_sfromutf8(v->s ? v->s : ""));
    }
}
/* a value for a parameter of ODBC, as Jet sends it (a date with its time) */
static int ov_bind(lbo_stmt *st, int i, vbvar x)
{
    vbvar v = vb_v_letval(x);
    char b[64];
    switch (v.vt) {
    case VT_EMPTY: case VT_NULL: return lbo_bind(st, i, LBO_NULL, 0, 0, 0, 0, 0);
    case VT_I2: case VT_I4: case VT_UI1: return lbo_bind(st, i, LBO_INT, vb_v2l(v), 0, 0, 0, 0);
    case VT_BOOL: return lbo_bind(st, i, LBO_INT, vb_v2bool(v) ? 1 : 0, 0, 0, 0, 0);
    case VT_R4: case VT_R8: return lbo_bind(st, i, LBO_DBL, 0, vb_v2d(v), 0, 0, 0);
    case VT_CY: snprintf(b, sizeof b, "%.4f", vb_v2d(v)); return lbo_bind(st, i, LBO_DEC, 0, 0, b, 0, 0);
    case VT_DATE: {
        vb_date d = vb_v2date(v);
        snprintf(b, sizeof b, "%04d-%02d-%02d %02d:%02d:%02d", vb_year(d), vb_month(d), vb_day(d), vb_hour(d), vb_minute(d), vb_second(d));
        return lbo_bind(st, i, LBO_DATE, 0, 0, b, 0, 0);
    }
    default:
        if (v.vt == (VT_ARRAY | VT_UI1)) {
            vbarr *a = vb_v_arrptr(v);
            return lbo_bind(st, i, LBO_BIN, 0, 0, a && a->nd ? a->data : "", a && a->nd ? a->n : 0, 0);
        }
        return lbo_bind(st, i, LBO_TEXT, 0, 0, vb_stoutf8(vb_v2s(v)), 0, 0);
    }
}

/* the DAO type of a column of ODBC, as Jet gives it for a linked table (VB6 oracle d_dao6) */
static void jet_of_sql(const lbo_colinfo *c, int *type, int *size)
{
    long n = (long)c->size;
    switch (c->sqltype) {
    case LBS_CHAR: case LBS_VARCHAR: case LBS_WCHAR: case LBS_WVARCHAR:
        if (n > 0 && n <= 255) { *type = dbText; *size = (int)n; } else { *type = dbMemo; *size = 0; }
        return;
    case LBS_LONGVARCHAR: case LBS_WLONGVARCHAR: *type = dbMemo; *size = 0; return;
    case LBS_INTEGER: *type = dbLong; *size = 4; return;
    case LBS_SMALLINT: *type = dbInteger; *size = 2; return;
    case LBS_TINYINT: *type = dbByte; *size = 1; return;
    case LBS_BIT: *type = dbBoolean; *size = 1; return;
    case LBS_BIGINT: *type = dbText; *size = 19; return;
    case LBS_REAL: *type = dbSingle; *size = 4; return;
    case LBS_FLOAT: case LBS_DOUBLE: case LBS_NUMERIC: case LBS_DECIMAL: *type = dbDouble; *size = 8; return;
    case LBS_DATE: case LBS_TIME: case LBS_TIMESTAMP: case LBS_TYPE_DATE: case LBS_TYPE_TIME: case LBS_TYPE_TIMESTAMP:
        *type = dbDate; *size = 8; return;
    case LBS_BINARY: case LBS_VARBINARY:
        if (n > 0 && n <= 255) { *type = dbBinary; *size = (int)n; } else { *type = dbLongBinary; *size = 0; }
        return;
    case LBS_LONGVARBINARY: *type = dbLongBinary; *size = 0; return;
    case LBS_GUID: *type = dbGUID; *size = 16; return;
    default: *type = dbText; *size = 255;
    }
}

static void put_qname(jbuf *b, osrc *o, const char *name)
{
    char q[2] = { o->quote, 0 };
    if (o->quote && o->quote != ' ') { jb_s(b, q); jb_s(b, name); jb_s(b, q); }
    else jb_s(b, name);
}

/* ---------------------------------------------------------------- the data sources */
static char *temp_file(void)
{
#ifdef _WIN32
    static int n;
    char dir[3000], path[3200];
    vbos_special_folder(2, dir, sizeof dir);
    for (int k = 0; k < 1000; k++) {
        snprintf(path, sizeof path, "%s\\lbodbc%d_%d.db", dir, (int)_getpid(), ++n);
        FILE *f = vbos_fopen(path, "rb");
        if (f) { fclose(f); continue; }
        f = vbos_fopen(path, "wb");
        if (!f) return 0;
        fclose(f);
        return strdup(path);
    }
    return 0;
#else
    const char *t = getenv("TMPDIR");
    char path[4096];
    snprintf(path, sizeof path, "%s/lbodbcXXXXXX", t && *t ? t : "/tmp");
    int fd = mkstemp(path);
    if (fd < 0) return 0;
    close(fd);
    return strdup(path);
#endif
}

/* the value of a key of a connection string (malloc), or 0 */
static char *cs_value(const char *cs, const char *key)
{
    size_t kl = strlen(key);
    for (const char *p = cs; p && *p;) {
        while (*p == ';' || *p == ' ') p++;
        const char *eq = strchr(p, '=');
        if (!eq) break;
        const char *v = eq + 1, *e;
        if (*v == '{') { e = strchr(v, '}'); e = e ? e + 1 : v + strlen(v); }
        else e = v + strcspn(v, ";");
        size_t n = (size_t)(eq - p);
        while (n && p[n - 1] == ' ') n--;
        if (n == kl && !strncasecmp(p, key, kl)) { char *r = malloc((size_t)(e - v) + 1); memcpy(r, v, (size_t)(e - v)); r[e - v] = 0; return r; }
        p = *e ? e + 1 : e;
    }
    return 0;
}
/* the connection string the driver completed, the files in it as the program named them (Linux: its "\") */
static void put_keys_as_given(jbuf *b, const char *out, const char *given)
{
    static const char *const K[] = { "Database", "DBQ", 0 };
    const char *p = out;
    while (*p) {
        const char *eq = strchr(p, '=');
        const char *v = eq ? eq + 1 : p + strlen(p), *e;
        if (*v == '{') { e = strchr(v, '}'); e = e ? e + 1 : v + strlen(v); e += strcspn(e, ";"); }
        else e = p + strcspn(p, ";");
        char *mine = 0;
        if (eq) for (int k = 0; K[k] && !mine; k++)
            if ((size_t)(eq - p) == strlen(K[k]) && !strncasecmp(p, K[k], strlen(K[k]))) mine = cs_value(given, K[k]);
        if (mine) { jb_put(b, p, (int)(v - p)); jb_s(b, mine); free(mine); }
        else jb_put(b, p, (int)(e - p));
        if (*e == ';') { jb_s(b, ";"); e++; }
        p = e;
    }
}

/* a data source by its Connect: connected once, its copies in a file of their own. 0 with the error raised */
static osrc *osrc_open(const char *connect, const char *src)
{
    for (int i = 0; i < g_nos; i++)
        if (!g_os[i]->stale && (!strcmp(g_os[i]->key, connect) || !strcmp(g_os[i]->connect, connect))) return g_os[i];
    coll_clear_items(err_list());
    const char *cs = connect;
    if (!strncasecmp(cs, "ODBC;", 5)) cs += 5;
    else if (!strcasecmp(cs, "ODBC")) cs = "";
    if (!lbo_available()) {
        err_append(0, "[Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified", "ODBC.Workspace");
        err_append(3146, "ODBC--call failed.", src);
        vb_raise_full(3146, vb_sfromutf8(src), vb_sfromutf8("ODBC--call failed."));
        return 0;
    }
    char *ocs = lb_odbc_cs(cs);
    char out[4096] = "";
    lbo_conn *k = lbo_connect(ocs ? ocs : cs, out, sizeof out);
    free(ocs);
    if (!k) { odbc_fail(3146, src, "ODBC--call failed."); return 0; }
    char *file = temp_file();
    ddb *mx = file ? db_new((dws *)vb_oid(dao_workspace0()), file) : 0;
    if (!mx || db_connect(mx, file, 0, 1) != SQLITE_OK) {
        lbo_disconnect(k);
        if (file) remove(file);
        free(file);
        if (mx) vb_orelease((vbobj)mx);
        dao_raise(3051, src, "The Microsoft Jet database engine cannot open the file '%s'.  It is already opened exclusively by another user, or you need permission to view its data.", "");
        return 0;
    }
    mx->ws = 0;
    /* (the copies are a cache: their changes are journaled in memory - Windows cannot delete the journal of a file
     * that another connection has open, and two connections have it here) */
    sqlite3_exec(mx->db, "PRAGMA journal_mode = MEMORY", 0, 0, 0);
    schema_init(mx);
    osrc *o = calloc(1, sizeof *o);
    o->key = strdup(connect);
    jbuf c = { 0 };
    jb_s(&c, "ODBC;"); put_keys_as_given(&c, out, cs);
    o->connect = c.p;
    char b[1024];
    char *given = cs_value(cs, "Database");
    o->dbname = given ? given : strdup(lbo_info(k, LBI_DATABASE_NAME, b, sizeof b) ? "" : b);
    o->dbms = strdup(lbo_info(k, LBI_DBMS_NAME, b, sizeof b) ? "" : b);
    o->quote = lbo_info(k, LBI_IDENTIFIER_QUOTE, b, sizeof b) ? '"' : b[0];
    o->k = k;
    o->file = file;
    o->mx = mx;
    if (!g_nos) atexit(os_cleanup);
    g_os = realloc(g_os, sizeof(osrc *) * (size_t)(g_nos + 1));
    g_os[g_nos++] = o;
    return o;
}

/* the copy of a table of the server (its schema; the rows come with osrc_reload): 0 when the server has none */
static otab *osrc_table(osrc *o, const char *name)
{
    for (int i = 0; i < o->ntabs; i++) if (!strcasecmp(o->tabs[i].name, name)) return &o->tabs[i];
    if (!strncasecmp(name, "sqlite_", 7) || !strncasecmp(name, "MSysLB_", 7)) return 0;   /* (no copy can have that name) */
    lbo_stmt *st = lbo_stmt_new(o->k);
    if (!st) return 0;
    jbuf q = { 0 };
    jb_s(&q, "SELECT * FROM "); put_qname(&q, o, name); jb_s(&q, " WHERE 1 = 0");
    int e = lbo_exec_direct(st, q.p);
    jb_free(&q);
    if (e) { lbo_stmt_free(st); return 0; }
    int nc = lbo_ncols(st);
    dtd *t = tabledef_new(name);
    otab ot; memset(&ot, 0, sizeof ot);
    ot.name = strdup(name);
    ot.ncols = nc;
    ot.cols = calloc((size_t)(nc ? nc : 1), sizeof(char *));
    ot.types = calloc((size_t)(nc ? nc : 1), sizeof(int));
    ot.iskey = calloc((size_t)(nc ? nc : 1), sizeof(int));
    ot.autocol = -1;
    int *req = calloc((size_t)(nc ? nc : 1), sizeof(int));
    for (int i = 0; i < nc; i++) {
        lbo_colinfo ci; memset(&ci, 0, sizeof ci);
        lbo_colinfo_get(st, i + 1, &ci);
        int type, size;
        jet_of_sql(&ci, &type, &size);
        dfld *f = field_new(FK_TABLE, (vbobj)t, ci.name, type, size);
        f->required = 0;                    /* (the server checks: no constraints of its own in the copy) */
        f->allowzero = 1;
        coll_add(t->fields, (vbobj)f);
        vb_orelease((vbobj)f);
        ot.cols[i] = strdup(ci.name);
        ot.types[i] = type;
        if (ci.autoinc && ot.autocol < 0) ot.autocol = i;
        req[i] = ci.nullable == 0;
    }
    lbo_stmt_free(st);
    /* Required: as SQLColumns gives it (NULLABLE 0) */
    st = lbo_stmt_new(o->k);
    if (st && !lbo_columns(st, 0, 0, name, 0)) {
        for (int i = 0; i < nc; i++) req[i] = 0;
        while (lbo_fetch(st) == 1) {
            lbo_val vc, vn; memset(&vc, 0, sizeof vc); memset(&vn, 0, sizeof vn);
            if (lbo_get(st, 4, LBS_VARCHAR, &vc)) memset(&vc, 0, sizeof vc);
            if (lbo_get(st, 11, LBS_SMALLINT, &vn)) memset(&vn, 0, sizeof vn);
            int nullable = vn.kind == LBO_INT ? (int)vn.iv : vn.s ? atoi(vn.s) : 1;
            for (int i = 0; i < nc && vc.s; i++) if (!strcasecmp(ot.cols[i], vc.s)) req[i] = nullable == 0;
            lbo_val_free(&vc); lbo_val_free(&vn);
        }
    }
    lbo_stmt_free(st);
    int r = schema_create_table(o->mx, t);
    vb_orelease((vbobj)t);
    if (r) { if (vb_errpend) vb_err_clear_pending(); free(req); free(ot.name); for (int i = 0; i < nc; i++) free(ot.cols[i]); free(ot.cols); free(ot.types); free(ot.iskey); return 0; }
    ddb *mx = o->mx;
    sqlite3_stmt *u = 0;
    /* as Jet gives them: positions from 1, Required, the copy marked */
    if (sqlite3_prepare_v2(mx->db, "UPDATE MSysLB_Fields SET ord = ord + 1 WHERE tbl = ?1", -1, &u, 0) == SQLITE_OK) {
        sqlite3_bind_text(u, 1, name, -1, SQLITE_STATIC); sqlite3_step(u); sqlite3_finalize(u);
    }
    if (sqlite3_prepare_v2(mx->db, "UPDATE MSysLB_Fields SET required = 1 WHERE tbl = ?1 AND name = ?2", -1, &u, 0) == SQLITE_OK) {
        for (int i = 0; i < nc; i++) if (req[i]) {
            sqlite3_reset(u);
            sqlite3_bind_text(u, 1, name, -1, SQLITE_STATIC); sqlite3_bind_text(u, 2, ot.cols[i], -1, SQLITE_STATIC);
            sqlite3_step(u);
        }
        sqlite3_finalize(u);
    }
    free(req);
    if (sqlite3_prepare_v2(mx->db, "UPDATE MSysLB_Tables SET attrs = attrs | ?2 WHERE name = ?1", -1, &u, 0) == SQLITE_OK) {
        sqlite3_bind_text(u, 1, name, -1, SQLITE_STATIC); sqlite3_bind_int(u, 2, LB_REMOTE_COPY); sqlite3_step(u); sqlite3_finalize(u);
    }
    /* the primary key */
    char **pk = 0; int npk = 0;
    st = lbo_stmt_new(o->k);
    if (st && !lbo_primary_keys(st, 0, 0, name)) {
        while (lbo_fetch(st) == 1) {
            lbo_val v; memset(&v, 0, sizeof v);
            if (!lbo_get(st, 4, LBS_VARCHAR, &v) && v.kind != LBO_NULL && v.s) {
                pk = realloc(pk, sizeof(char *) * (size_t)(npk + 1));
                pk[npk++] = strdup(v.s);
                for (int i = 0; i < nc; i++) if (!strcasecmp(ot.cols[i], v.s)) ot.iskey[i] = 1;
            }
            lbo_val_free(&v);
        }
    }
    lbo_stmt_free(st);
    /* the indexes (SQLStatistics): their names, fields, unique, the primary key */
    st = lbo_stmt_new(o->k);
    if (st && !lbo_statistics(st, 0, 0, name, 0)) {
        char cur[256] = ""; jbuf fl = { 0 }; int uniq = 0, ord = 0, nfl = 0, pkhit = 0;
        sqlite3_stmt *ins = 0;
        sqlite3_prepare_v2(mx->db, "INSERT OR REPLACE INTO MSysLB_Indexes (tbl, name, fields, isprimary, isunique, required, ignorenulls, isforeign, ord) VALUES (?1, ?2, ?3, ?4, ?5, ?4, 0, 0, ?6)", -1, &ins, 0);
#define FLUSH() do { if (*cur && ins) { int prim = npk > 0 && nfl == npk && pkhit == npk; \
            sqlite3_reset(ins); sqlite3_bind_text(ins, 1, name, -1, SQLITE_STATIC); sqlite3_bind_text(ins, 2, cur, -1, SQLITE_STATIC); \
            sqlite3_bind_text(ins, 3, fl.p ? fl.p : "", -1, SQLITE_STATIC); sqlite3_bind_int(ins, 4, prim); sqlite3_bind_int(ins, 5, uniq); \
            sqlite3_bind_int(ins, 6, ord++); sqlite3_step(ins); } fl.n = 0; if (fl.p) fl.p[0] = 0; nfl = 0; pkhit = 0; } while (0)
        while (lbo_fetch(st) == 1) {
            lbo_val vt, vn, vu, vc, vd;
            memset(&vt, 0, sizeof vt); memset(&vn, 0, sizeof vn); memset(&vu, 0, sizeof vu); memset(&vc, 0, sizeof vc); memset(&vd, 0, sizeof vd);
            if (lbo_get(st, 7, LBS_SMALLINT, &vt)) memset(&vt, 0, sizeof vt);
            int typ = vt.kind == LBO_INT ? (int)vt.iv : vt.kind == LBO_TEXT && vt.s ? atoi(vt.s) : 0;
            lbo_val_free(&vt);
            if (typ == 0) continue;             /* SQL_TABLE_STAT */
            if (lbo_get(st, 6, LBS_VARCHAR, &vn)) memset(&vn, 0, sizeof vn);
            if (lbo_get(st, 4, LBS_SMALLINT, &vu)) memset(&vu, 0, sizeof vu);
            if (lbo_get(st, 9, LBS_VARCHAR, &vc)) memset(&vc, 0, sizeof vc);
            if (lbo_get(st, 10, LBS_VARCHAR, &vd)) memset(&vd, 0, sizeof vd);
            const char *iname = vn.s ? vn.s : "";
            if (strcmp(cur, iname)) {
                FLUSH();
                snprintf(cur, sizeof cur, "%s", iname);
                uniq = !(vu.kind == LBO_INT ? vu.iv : vu.s ? atoi(vu.s) : 1);
            }
            if (vc.s) {
                jb_s(&fl, nfl ? ";" : ""); jb_s(&fl, vd.s && (vd.s[0] == 'D' || vd.s[0] == 'd') ? "-" : "+"); jb_s(&fl, vc.s);
                for (int i = 0; i < npk; i++) if (!strcasecmp(pk[i], vc.s)) pkhit++;
                nfl++;
            }
            lbo_val_free(&vn); lbo_val_free(&vu); lbo_val_free(&vc); lbo_val_free(&vd);
        }
        FLUSH();
#undef FLUSH
        if (ins) sqlite3_finalize(ins);
        jb_free(&fl);
    }
    lbo_stmt_free(st);
    for (int i = 0; i < npk; i++) free(pk[i]);
    free(pk);
    o->tabs = realloc(o->tabs, sizeof(otab) * (size_t)(o->ntabs + 1));
    o->tabs[o->ntabs] = ot;
    return &o->tabs[o->ntabs++];
}

/* the rows of a table again from the server into its copy (schema: "main" or "lbextN" of the connection of d);
 * once a statement. 0, or -1 with the error raised */
static int osrc_reload(ddb *d, const char *schema, osrc *o, otab *t, const char *src)
{
    if (t->loaded == g_stamp) return 0;
    t->loaded = g_stamp;
    jbuf s = { 0 };
    jb_f(&s, "DELETE FROM \"%s\".", schema); jb_qid(&s, t->name, -1);
    int drc = sqlite3_exec(d->db, s.p, 0, 0, 0);
    if (drc != SQLITE_OK && getenv("VBRT_DEBUG_ODBC"))
        fprintf(stderr, "lbodbc: %s: %d %s\n", s.p, sqlite3_extended_errcode(d->db), sqlite3_errmsg(d->db));
    s.n = 0;
    lbo_stmt *st = lbo_stmt_new(o->k);
    jbuf q = { 0 };
    jb_s(&q, "SELECT * FROM "); put_qname(&q, o, t->name);
    if (!st || lbo_exec_direct(st, q.p)) { jb_free(&q); jb_free(&s); odbc_fail(3146, src, "ODBC--call failed."); lbo_stmt_free(st); return -1; }
    jb_free(&q);
    int nc = lbo_ncols(st);
    lbo_colinfo *ci = calloc((size_t)(nc ? nc : 1), sizeof *ci);
    for (int i = 0; i < nc; i++) lbo_colinfo_get(st, i + 1, &ci[i]);
    jb_f(&s, "INSERT INTO \"%s\".", schema); jb_qid(&s, t->name, -1); jb_s(&s, " (");
    for (int i = 0; i < nc; i++) { if (i) jb_s(&s, ", "); jb_qid(&s, ci[i].name, -1); }
    jb_s(&s, ") VALUES (");
    for (int i = 0; i < nc; i++) jb_f(&s, i ? ", ?%d" : "?%d", i + 1);
    jb_s(&s, ")");
    sqlite3_stmt *ins = 0;
    sqlite3_prepare_v2(d->db, s.p, -1, &ins, 0);
    jb_free(&s);
    int *ty = calloc((size_t)(nc ? nc : 1), sizeof(int));
    for (int i = 0; i < nc; i++) for (int k = 0; k < t->ncols; k++) if (!strcasecmp(t->cols[k], ci[i].name)) ty[i] = t->types[k];
    while (ins && lbo_fetch(st) == 1) {
        vb_i32 mk = vb_mark();
        sqlite3_reset(ins);
        for (int i = 0; i < nc; i++) {
            lbo_val v;
            if (lbo_get(st, i + 1, ci[i].sqltype, &v)) { v.kind = LBO_NULL; v.s = 0; }
            jet_bind(ins, i + 1, ov_val(&v), ty[i]);
            lbo_val_free(&v);
        }
        sqlite3_step(ins);
        vb_release(mk);
    }
    if (ins) sqlite3_finalize(ins);
    free(ty); free(ci);
    lbo_stmt_free(st);
    return 0;
}

static osrc *ext_osrc(ddb *d, int ext) { return ext > 0 && ext <= d->next ? (osrc *)d->ext[ext - 1].odbc : ext == 0 ? (osrc *)d->odbc : 0; }
static void schema_name(int ext, char out[32]) { if (ext > 0) snprintf(out, 32, "lbext%d", ext); else snprintf(out, 32, "main"); }

/* before a statement: the rows of the tables of ODBC it names, read again. 0, or -1 with the error raised */
static int odbc_before(ddb *d, jsql *j)
{
    g_stamp++;
    for (int i = 0; i < j->ntabs; i++) {
        osrc *o = ext_osrc(d, j->tabext[i]);
        if (!o) continue;
        otab *t = osrc_table(o, j->tabs[i]);
        if (!t) continue;
        char sch[32]; schema_name(j->tabext[i], sch);
        if (osrc_reload(d, sch, o, t, "DAO.Database")) return -1;
    }
    return 0;
}
/* a table opened by name */
static int odbc_before_table(ddb *d, int ext, const char *name)
{
    osrc *o = ext_osrc(d, ext);
    if (!o) return 0;
    g_stamp++;
    otab *t = osrc_table(o, name);
    if (!t) return 0;
    char sch[32]; schema_name(ext, sch);
    return osrc_reload(d, sch, o, t, "DAO.Database");
}

/* ---------------------------------------------------------------- changes to the server */
static int is_long(int t) { return t == dbMemo || t == dbLongBinary; }
static void server_tran(osrc *o, int what)      /* 0 begin (when the program has none), 1 commit, 2 rollback */
{
    if (o->trans) return;
    if (what == 0) lbo_autocommit(o->k, 0);
    else { lbo_endtran(o->k, what == 1); lbo_autocommit(o->k, 1); }
}
/* run a statement with its values on the server: the rows it affected, -1 when it failed (the diagnostics kept) */
static long server_dml(osrc *o, const char *sql, vbvar *vals, int nv)
{
    lbo_stmt *st = lbo_stmt_new(o->k);
    int e = !st || lbo_prepare(st, sql);
    for (int i = 0; !e && i < nv; i++) e = ov_bind(st, i + 1, vals[i]);
    if (!e) e = lbo_execute(st);
    long n = e ? -1 : lbo_rowcount(st);
    lbo_stmt_free(st);
    return n < -1 ? 0 : n;
}
/* the identity a new row got */
static vbvar server_identity(osrc *o)
{
    const char *q = 0, *db = o->dbms ? o->dbms : "";
    if (strcasestr(db, "sqlite")) q = "SELECT last_insert_rowid()";
    else if (strcasestr(db, "mysql") || strcasestr(db, "mariadb")) q = "SELECT LAST_INSERT_ID()";
    else if (strcasestr(db, "postgres")) q = "SELECT lastval()";
    else if (strcasestr(db, "sql server") || strcasestr(db, "sybase")) q = "SELECT @@IDENTITY";
    vbvar r = vb_v_null();
    lbo_stmt *qs = q ? lbo_stmt_new(o->k) : 0;
    if (qs && !lbo_exec_direct(qs, q) && lbo_fetch(qs) == 1) {
        lbo_val v; memset(&v, 0, sizeof v);
        if (!lbo_get(qs, 1, LBS_BIGINT, &v)) r = ov_val(&v);
        lbo_val_free(&v);
    }
    lbo_stmt_free(qs);
    return r;
}

/* a change of a row of a dynaset (ddb.remote): 1 insert, 2 update, 3 delete. 0, or -1 with the error raised */
static int odbc_remote(void *ctx, drs *r, int op)
{
    ddb *d = ctx;
    osrc *o = ext_osrc(d, r->bext);
    if (!o || !r->base) return 0;
    otab *t = osrc_table(o, r->base);
    if (!t) return 0;
    jbuf s = { 0 };
    vbvar *vals = calloc((size_t)(r->ncols * 2 + 2), sizeof(vbvar));
    int nv = 0;
    vbvar *old = op != 1 && r->cur >= 0 && r->cur < r->d->n ? r->d->r[r->cur].v : 0;
#define MINE(i_) (r->cols[i_].sfield && r->cols[i_].stable && !strcasecmp(r->cols[i_].stable, r->base))
    if (op == 1) {
        jb_s(&s, "INSERT INTO "); put_qname(&s, o, t->name); jb_s(&s, " (");
        int k = 0;
        for (int i = 0; i < r->ncols; i++) {
            if (!MINE(i) || !r->edirty[i]) continue;
            jb_s(&s, k ? ", " : ""); put_qname(&s, o, r->cols[i].sfield);
            vals[nv++] = r->ebuf[i]; k++;
        }
        jb_s(&s, ") VALUES (");
        for (int i = 0; i < k; i++) jb_s(&s, i ? ", ?" : "?");
        jb_s(&s, ")");
        if (!k) { s.n = 0; jb_s(&s, "INSERT INTO "); put_qname(&s, o, t->name); jb_s(&s, " DEFAULT VALUES"); }
    } else {
        if (op == 2) {
            jb_s(&s, "UPDATE "); put_qname(&s, o, t->name); jb_s(&s, " SET ");
            int k = 0;
            for (int i = 0; i < r->ncols; i++) {
                if (!MINE(i) || !r->edirty[i]) continue;
                jb_s(&s, k ? ", " : ""); put_qname(&s, o, r->cols[i].sfield); jb_s(&s, " = ?");
                vals[nv++] = r->ebuf[i]; k++;
            }
            if (!k) { free(vals); jb_free(&s); return 0; }
        } else { jb_s(&s, "DELETE FROM "); put_qname(&s, o, t->name); }
        jb_s(&s, " WHERE ");
        int k = 0;
        for (int i = 0; i < r->ncols && old; i++) {
            if (!MINE(i) || is_long(r->cols[i].type)) continue;
            jb_s(&s, k ? " AND " : ""); put_qname(&s, o, r->cols[i].sfield);
            if (old[i].vt == VT_NULL || old[i].vt == VT_EMPTY) jb_s(&s, " IS NULL");
            else { jb_s(&s, " = ?"); vals[nv++] = old[i]; }
            k++;
        }
        if (!k) jb_s(&s, "1 = 0");
    }
    long n = server_dml(o, s.p, vals, nv);
    jb_free(&s); free(vals);
    if (n < 0) { odbc_fail(3146, "DAO.Recordset", "ODBC--call failed."); return -1; }
    if (op != 1 && n == 0) {
        dao_raise(3197, "DAO.Recordset", "The Microsoft Jet database engine stopped the process because you and another user are attempting to change the same data at the same time.");
        return -1;
    }
    if (op == 1 && t->autocol >= 0)             /* the number the server gave the new row */
        for (int i = 0; i < r->ncols; i++)
            if (MINE(i) && !strcasecmp(r->cols[i].sfield, t->cols[t->autocol]) && !r->edirty[i]) { vbvar id = server_identity(o); if (id.vt != VT_NULL) vb_vset(&r->ebuf[i], id); }
#undef MINE
    return 0;
}

/* ---------------------------------------------------------------- action queries of Jet SQL */
static int same_val(vbvar a, vbvar b)
{
    if (a.vt != b.vt) return 0;
    switch (a.vt) {
    case VT_I2: case VT_I4: case VT_UI1: case VT_BOOL: return vb_v2l(a) == vb_v2l(b);
    case VT_R4: case VT_R8: case VT_DATE: case VT_CY: return vb_v2d(a) == vb_v2d(b);
    case VT_BSTR: return !strcmp(vb_stoutf8(a.u.s), vb_stoutf8(b.u.s));
    default: return 0;
    }
}
typedef struct ocap { osrc *o; otab *t; char sch[32]; int on; } ocap;

/* the changes the statement makes to the copy, logged (temp.lbo_log: op, rowid, the old values) */
static int ocap_begin(ddb *d, jsql *j, ocap *c)
{
    memset(c, 0, sizeof *c);
    if (!(j->kind == J_INSERT || j->kind == J_UPDATE || j->kind == J_DELETE) || !j->table) return 0;
    osrc *o = ext_osrc(d, j->table_ext);
    if (!o) return 0;
    otab *t = osrc_table(o, j->table_src ? j->table_src : j->table);
    if (!t) return 0;
    c->o = o; c->t = t; schema_name(j->table_ext, c->sch);
    jbuf s = { 0 }, oldv = { 0 }, newv = { 0 };
    for (int i = 0; i < t->ncols; i++) {
        jb_s(&oldv, ", OLD."); jb_qid(&oldv, t->cols[i], -1);
        jb_s(&newv, ", NEW."); jb_qid(&newv, t->cols[i], -1);
    }
    jb_s(&s, "DROP TABLE IF EXISTS temp.lbo_log; CREATE TEMP TABLE lbo_log AS SELECT 0 AS lbop, 0 AS lbrid, * FROM \"");
    jb_s(&s, c->sch); jb_s(&s, "\"."); jb_qid(&s, t->name, -1); jb_s(&s, " WHERE 0;");
    jb_s(&s, "CREATE TEMP TRIGGER lbo_ti AFTER INSERT ON \""); jb_s(&s, c->sch); jb_s(&s, "\"."); jb_qid(&s, t->name, -1);
    jb_s(&s, " BEGIN INSERT INTO lbo_log VALUES (1, NEW.rowid"); jb_s(&s, newv.p ? newv.p : ""); jb_s(&s, "); END;");
    jb_s(&s, "CREATE TEMP TRIGGER lbo_tu AFTER UPDATE ON \""); jb_s(&s, c->sch); jb_s(&s, "\"."); jb_qid(&s, t->name, -1);
    jb_s(&s, " BEGIN INSERT INTO lbo_log VALUES (2, OLD.rowid"); jb_s(&s, oldv.p ? oldv.p : ""); jb_s(&s, "); END;");
    jb_s(&s, "CREATE TEMP TRIGGER lbo_td BEFORE DELETE ON \""); jb_s(&s, c->sch); jb_s(&s, "\"."); jb_qid(&s, t->name, -1);
    jb_s(&s, " BEGIN INSERT INTO lbo_log VALUES (3, OLD.rowid"); jb_s(&s, oldv.p ? oldv.p : ""); jb_s(&s, "); END;");
    jb_s(&s, "SAVEPOINT lbodbc;");
    int rc = sqlite3_exec(d->db, s.p, 0, 0, 0);
    jb_free(&s); jb_free(&oldv); jb_free(&newv);
    if (rc != SQLITE_OK) { sqlite3_exec(d->db, "DROP TRIGGER IF EXISTS temp.lbo_ti; DROP TRIGGER IF EXISTS temp.lbo_tu; DROP TRIGGER IF EXISTS temp.lbo_td; DROP TABLE IF EXISTS temp.lbo_log", 0, 0, 0); return 0; }
    c->on = 1;
    return 0;
}

static void ocap_drop(ddb *d)
{
    sqlite3_exec(d->db, "DROP TRIGGER IF EXISTS temp.lbo_ti; DROP TRIGGER IF EXISTS temp.lbo_tu; DROP TRIGGER IF EXISTS temp.lbo_td; DROP TABLE IF EXISTS temp.lbo_log", 0, 0, 0);
}

/* the statement ran on the copy (ok) or failed: its changes to the server, each row by its key (by its values that
 * are not long, without one). 0, or -1 with the error raised (the copy as it was) */
static int ocap_end(ddb *d, ocap *c, int ok, const char *shown)
{
    if (!c->on) return 0;
    if (!ok) { sqlite3_exec(d->db, "ROLLBACK TO lbodbc; RELEASE lbodbc", 0, 0, 0); ocap_drop(d); return 0; }
    otab *t = c->t; osrc *o = c->o;
    int nkey = 0;
    for (int i = 0; i < t->ncols; i++) if (t->iskey[i]) nkey++;
    server_tran(o, 0);
    int fail = 0, failop = 0;
    sqlite3_stmt *lg = 0, *now = 0;
    sqlite3_prepare_v2(d->db, "SELECT * FROM temp.lbo_log ORDER BY rowid", -1, &lg, 0);
    jbuf q = { 0 };
    jb_s(&q, "SELECT * FROM \""); jb_s(&q, c->sch); jb_s(&q, "\"."); jb_qid(&q, t->name, -1); jb_s(&q, " WHERE rowid = ?1");
    sqlite3_prepare_v2(d->db, q.p, -1, &now, 0);
    jb_free(&q);
    while (lg && !fail && sqlite3_step(lg) == SQLITE_ROW) {
        vb_i32 mk = vb_mark();
        int op = sqlite3_column_int(lg, 0);
        vbvar *ov = calloc((size_t)t->ncols + 1, sizeof(vbvar)), *nw = calloc((size_t)t->ncols + 1, sizeof(vbvar));
        for (int i = 0; i < t->ncols; i++) ov[i] = jet_column(lg, i + 2, t->types[i]);
        if (op == 2 && now) {
            sqlite3_reset(now);
            sqlite3_bind_int64(now, 1, sqlite3_column_int64(lg, 1));
            if (sqlite3_step(now) == SQLITE_ROW) for (int i = 0; i < t->ncols; i++) nw[i] = jet_column(now, i, t->types[i]);
        }
        jbuf s = { 0 };
        vbvar *vals = calloc((size_t)t->ncols * 2 + 2, sizeof(vbvar));
        int nv = 0;
        if (op == 1) {
            jb_s(&s, "INSERT INTO "); put_qname(&s, o, t->name); jb_s(&s, " (");
            int k = 0;
            for (int i = 0; i < t->ncols; i++) {
                if (ov[i].vt == VT_NULL || ov[i].vt == VT_EMPTY) continue;
                jb_s(&s, k ? ", " : ""); put_qname(&s, o, t->cols[i]); vals[nv++] = ov[i]; k++;
            }
            jb_s(&s, ") VALUES (");
            for (int i = 0; i < k; i++) jb_s(&s, i ? ", ?" : "?");
            jb_s(&s, ")");
            if (!k) { s.n = 0; jb_s(&s, "INSERT INTO "); put_qname(&s, o, t->name); jb_s(&s, " DEFAULT VALUES"); }
        } else {
            if (op == 2) {
                jb_s(&s, "UPDATE "); put_qname(&s, o, t->name); jb_s(&s, " SET ");
                int k = 0;
                for (int i = 0; i < t->ncols; i++) {
                    vbvar a = ov[i], b = nw[i];
                    int an = a.vt == VT_NULL || a.vt == VT_EMPTY, bn = b.vt == VT_NULL || b.vt == VT_EMPTY;
                    if (an && bn) continue;
                    if (!an && !bn && same_val(a, b)) continue;
                    jb_s(&s, k ? ", " : ""); put_qname(&s, o, t->cols[i]); jb_s(&s, " = ?"); vals[nv++] = b; k++;
                }
                if (!k) { jb_free(&s); free(vals); free(ov); free(nw); vb_release(mk); continue; }
            } else { jb_s(&s, "DELETE FROM "); put_qname(&s, o, t->name); }
            jb_s(&s, " WHERE ");
            int k = 0;
            for (int i = 0; i < t->ncols; i++) {
                if (nkey ? !t->iskey[i] : is_long(t->types[i])) continue;
                jb_s(&s, k ? " AND " : ""); put_qname(&s, o, t->cols[i]);
                if (ov[i].vt == VT_NULL || ov[i].vt == VT_EMPTY) jb_s(&s, " IS NULL");
                else { jb_s(&s, " = ?"); vals[nv++] = ov[i]; }
                k++;
            }
            if (!k) jb_s(&s, "1 = 0");
        }
        if (server_dml(o, s.p, vals, nv) < 0) { fail = 1; failop = op; }
        jb_free(&s); free(vals); free(ov); free(nw);
        vb_release(mk);
    }
    if (lg) sqlite3_finalize(lg);
    if (now) sqlite3_finalize(now);
    if (fail) {
        /* the errors of ODBC are kept by the call that failed: read before anything else runs */
        int n = lbo_nerr();
        lbo_err keep[8]; if (n > 8) n = 8;
        memcpy(keep, lbo_errs(), sizeof(lbo_err) * (size_t)n);
        server_tran(o, 2);
        sqlite3_exec(d->db, "ROLLBACK TO lbodbc; RELEASE lbodbc", 0, 0, 0);
        ocap_drop(d);
        const char *what = failop == 1 ? "insert" : failop == 2 ? "update" : "delete";
        int num = failop == 1 ? 3155 : failop == 2 ? 3157 : 3156;
        char msg[600];
        snprintf(msg, sizeof msg, "ODBC--%s on a linked table '%s' failed.", what, shown);
        coll_clear_items(err_list());
        for (int i = 0; i < n; i++) {
            char m[1100];
            lb_odbc_msg(keep[i].msg, m, sizeof m);
            err_append(keep[i].native, m, "ODBC.Database");
        }
        err_append(num, msg, "DAO.Database");
        vb_raise_full(num, vb_sfromutf8("DAO.Database"), vb_sfromutf8(msg));
        return -1;
    }
    server_tran(o, 1);
    sqlite3_exec(d->db, "RELEASE lbodbc", 0, 0, 0);
    ocap_drop(d);
    return 0;
}

/* ---------------------------------------------------------------- pass-through */
static int odbc_pt_exec(osrc *o, const char *sql, const char *src)
{
    lbo_stmt *st = lbo_stmt_new(o->k);
    if (!st || lbo_exec_direct(st, sql)) { odbc_fail(3146, src, "ODBC--call failed."); lbo_stmt_free(st); return -1; }
    lbo_stmt_free(st);
    return 0;
}
/* the rows of a pass-through query: a recordset that cannot be changed */
static vbobj odbc_pt_open(ddb *d, osrc *o, const char *sql, int type, const char *src)
{
    lbo_stmt *st = lbo_stmt_new(o->k);
    if (!st || lbo_exec_direct(st, sql)) { odbc_fail(3146, src, "ODBC--call failed."); lbo_stmt_free(st); return 0; }
    int nc = lbo_ncols(st);
    if (nc <= 0) { lbo_stmt_free(st); dao_raise(3219, src, "Invalid operation."); return 0; }
    lbo_colinfo *ci = calloc((size_t)nc, sizeof *ci);
    int *ty = calloc((size_t)nc, sizeof(int));
    char tname[32]; snprintf(tname, sizeof tname, "lbpt%d", ++g_ptno);
    jbuf s = { 0 };
    jb_f(&s, "CREATE TEMP TABLE %s (", tname);
    for (int i = 0; i < nc; i++) {
        lbo_colinfo_get(st, i + 1, &ci[i]);
        int size; jet_of_sql(&ci[i], &ty[i], &size);
        char nm[300], decl[64];
        if (ci[i].name[0]) snprintf(nm, sizeof nm, "%s", ci[i].name); else snprintf(nm, sizeof nm, "Expr%d", 1000 + i);
        if (i) jb_s(&s, ", ");
        jb_qid(&s, nm, -1); jb_s(&s, " "); jb_s(&s, jet_type_decl(ty[i], 0, decl, sizeof decl));
        if (ty[i] == dbText) jb_f(&s, "(%d)", size);
    }
    jb_s(&s, ")");
    sqlite3_exec(d->db, s.p, 0, 0, 0);
    s.n = 0;
    jb_f(&s, "INSERT INTO temp.%s VALUES (", tname);
    for (int i = 0; i < nc; i++) jb_f(&s, i ? ", ?%d" : "?%d", i + 1);
    jb_s(&s, ")");
    sqlite3_stmt *ins = 0;
    sqlite3_prepare_v2(d->db, s.p, -1, &ins, 0);
    while (ins && lbo_fetch(st) == 1) {
        vb_i32 mk = vb_mark();
        sqlite3_reset(ins);
        for (int i = 0; i < nc; i++) {
            lbo_val v;
            if (lbo_get(st, i + 1, ci[i].sqltype, &v)) { v.kind = LBO_NULL; v.s = 0; }
            jet_bind(ins, i + 1, ov_val(&v), ty[i]);
            lbo_val_free(&v);
        }
        sqlite3_step(ins);
        vb_release(mk);
    }
    if (ins) sqlite3_finalize(ins);
    lbo_stmt_free(st);
    free(ci); free(ty);
    s.n = 0;
    jb_f(&s, "SELECT * FROM temp.%s", tname);
    vbobj r = rs_open_sqlite(d, s.p, 0, type ? type : dbOpenSnapshot, sql, 0, 0);
    jb_free(&s);
    return r;
}
static osrc *qd_osrc(vbobj qd)
{
    if (!qd) return 0;
    dqd *q = (dqd *)vb_oid(qd);
    if (!q->connect || strncasecmp(q->connect, "ODBC", 4)) return 0;
    return osrc_open(q->connect, "DAO.QueryDef");
}

/* ---------------------------------------------------------------- transactions */
static void odbc_trans(ddb *d, int what)        /* 0 begin, 1 commit, 2 rollback */
{
    osrc *seen[DAO_MAXEXT + 1]; int ns = 0;
    for (int i = -1; i < d->next; i++) {
        osrc *o = i < 0 ? (osrc *)d->odbc : (osrc *)d->ext[i].odbc;
        if (!o) continue;
        int dup = 0;
        for (int k = 0; k < ns; k++) if (seen[k] == o) dup = 1;
        if (dup) continue;
        seen[ns++] = o;
        if (what == 0) { if (o->trans++ == 0) lbo_autocommit(o->k, 0); }
        else if (o->trans > 0 && --o->trans == 0) { lbo_endtran(o->k, what == 1); lbo_autocommit(o->k, 1); }
    }
}

/* ---------------------------------------------------------------- OpenDatabase("", , , "ODBC;...") */
/* DBEngine.RegisterDatabase: the sources of the data source named are not used again (VB6 d_dao7: the next
 * OpenDatabase connects to what was registered) */
static void odbc_forget_dsn(const char *dsn)
{
    size_t n = strlen(dsn);
    for (int i = 0; i < g_nos; i++) {
        const char *ks[2] = { g_os[i]->key, g_os[i]->connect };
        for (int j = 0; j < 2; j++)
            for (const char *p = ks[j]; p && (p = strcasestr(p, "DSN=")) != 0; p += 4)
                if ((p == ks[j] || p[-1] == ';' || p[-1] == ' ') && !strncasecmp(p + 4, dsn, n) &&
                    (p[4 + n] == ';' || p[4 + n] == 0))
                    g_os[i]->stale = 1;
    }
}

static vbobj odbc_db_open(dws *w, const char *connect, int ro)
{
    osrc *o = osrc_open(connect, "DAO.Workspace");
    if (!o) return 0;
    ddb *d = db_new(w, o->dbname);
    free(d->connect); d->connect = strdup(o->connect);
    if (db_connect(d, o->file, ro, 0) != SQLITE_OK) {
        vb_orelease((vbobj)d);
        dao_raise(3051, "DAO.Workspace", "The Microsoft Jet database engine cannot open the file '%s'.  It is already opened exclusively by another user, or you need permission to view its data.", o->file);
        return 0;
    }
    sqlite3_exec(d->db, "PRAGMA journal_mode = MEMORY", 0, 0, 0);
    d->odbc = o;
    d->remote = odbc_remote;
    d->remote_ctx = d;
    coll_add(w->dbs, (vbobj)d);
    vb_orelease((vbobj)d);
    return (vbobj)d;
}
/* the tables of the data source (TableDefs): a copy of each made */
void dao_odbc_sync(ddb *d)
{
    osrc *o = d->odbc;
    if (!o) return;
    lbo_stmt *st = lbo_stmt_new(o->k);
    if (!st || lbo_tables(st, 0, 0, "%", "TABLE,VIEW")) { lbo_stmt_free(st); return; }
    char **names = 0; int n = 0;
    while (lbo_fetch(st) == 1) {
        lbo_val v; memset(&v, 0, sizeof v);
        if (!lbo_get(st, 3, LBS_VARCHAR, &v) && v.s) { names = realloc(names, sizeof(char *) * (size_t)(n + 1)); names[n++] = strdup(v.s); }
        lbo_val_free(&v);
    }
    lbo_stmt_free(st);
    for (int i = 0; i < n; i++) { osrc_table(o, names[i]); free(names[i]); }
    free(names);
}

/* ---------------------------------------------------------------- for daors.c */
vbobj dao_odbc_passthrough(ddb *d, const char *sql, int type, int options, vbobj qd, int *handled)
{
    *handled = 0;
    dqd *q = qd ? (dqd *)vb_oid(qd) : 0;
    osrc *o = 0;
    if (q && q->connect && !strncasecmp(q->connect, "ODBC", 4)) {
        *handled = 1;
        if (!(o = qd_osrc(qd))) return 0;
    }
    if (!o && (options & dbSQLPassThrough) && d->odbc) o = (osrc *)d->odbc;
    if (!o) return 0;
    *handled = 1;
    return odbc_pt_open(d, o, sql, type, "DAO.Database");
}
void dao_odbc_table(ddb *d, const char *name) { if (d->odbc) osrc_table((osrc *)d->odbc, name); }
int dao_odbc_before_table(ddb *d, int ext, const char *name) { return odbc_before_table(d, ext, name); }
int dao_odbc_before(ddb *d, jsql *j) { return odbc_before(d, j); }

/* ---------------------------------------------------------------- the schema of the data source (DDL of DAO) */
static int family(osrc *o)                      /* 1 SQL Server / Sybase, 2 MySQL / MariaDB, 3 PostgreSQL, 4 Oracle, 0 others */
{
    const char *db = o->dbms ? o->dbms : "";
    if (strcasestr(db, "sql server") || strcasestr(db, "sybase")) return 1;
    if (strcasestr(db, "mysql") || strcasestr(db, "mariadb")) return 2;
    if (strcasestr(db, "postgres")) return 3;
    if (strcasestr(db, "oracle")) return 4;
    return 0;
}
/* the server's name of a DAO type */
static const char *server_type(osrc *o, int type, int size, char *buf, int cap)
{
    int f = family(o);
    switch (type) {
    case dbBoolean: return f == 3 ? "BOOLEAN" : f == 4 ? "NUMBER(1)" : "BIT";
    case dbByte: return f == 4 ? "NUMBER(3)" : "SMALLINT";
    case dbInteger: return f == 4 ? "NUMBER(5)" : "SMALLINT";
    case dbLong: return f == 4 ? "NUMBER(10)" : "INTEGER";
    case dbCurrency: return f == 1 ? "MONEY" : "NUMERIC(19,4)";
    case dbSingle: return "REAL";
    case dbDouble: return f == 3 ? "DOUBLE PRECISION" : f == 2 ? "DOUBLE" : "FLOAT";
    case dbDate: return f == 3 || f == 4 ? "TIMESTAMP" : "DATETIME";
    case dbText: snprintf(buf, (size_t)cap, f == 4 ? "VARCHAR2(%d)" : "VARCHAR(%d)", size > 0 ? size : 255); return buf;
    case dbMemo: return f == 4 ? "CLOB" : f == 2 ? "LONGTEXT" : "TEXT";
    case dbBinary: snprintf(buf, (size_t)cap, f == 4 ? "RAW(%d)" : f == 3 ? "BYTEA" : "VARBINARY(%d)", size > 0 ? size : 255); return buf;
    case dbLongBinary: return f == 1 ? "IMAGE" : f == 3 ? "BYTEA" : f == 2 ? "LONGBLOB" : "BLOB";
    case dbGUID: return f == 1 ? "UNIQUEIDENTIFIER" : "CHAR(38)";
    case dbDecimal: return "DECIMAL(18,0)";
    default: snprintf(buf, (size_t)cap, "VARCHAR(%d)", 255); return buf;
    }
}
/* the copy of a table forgotten (its table changed on the server): made again when it is named */
static void osrc_forget(osrc *o, const char *name)
{
    for (int i = 0; i < o->ntabs; i++) {
        if (strcasecmp(o->tabs[i].name, name)) continue;
        otab *t = &o->tabs[i];
        for (int k = 0; k < t->ncols; k++) free(t->cols[k]);
        free(t->cols); free(t->types); free(t->iskey); free(t->name);
        memmove(&o->tabs[i], &o->tabs[i + 1], sizeof(otab) * (size_t)(o->ntabs - i - 1));
        o->ntabs--;
        break;
    }
    char real[256];
    if (schema_has_table(o->mx, name, real, sizeof real) == 1) {
        schema_drop_table(o->mx, real);
        if (vb_errpend) vb_err_clear_pending();
    }
}
static int server_ddl(osrc *o, const char *sql, const char *src)
{
    lbo_stmt *st = lbo_stmt_new(o->k);
    if (!st || lbo_exec_direct(st, sql)) { odbc_fail(3146, src, "ODBC--call failed."); lbo_stmt_free(st); return -1; }
    lbo_stmt_free(st);
    return 0;
}
static void index_cols(osrc *o, dix *ix, jbuf *s)
{
    for (int k = 0; k < ix->fields->n; k++) {
        dfld *f = (dfld *)ix->fields->items[k];
        if (k) jb_s(s, ", ");
        put_qname(s, o, f->name);
        if (f->attrs & dbDescending) jb_s(s, " DESC");
    }
}
static int server_index(osrc *o, const char *table, dix *ix, const char *src)
{
    jbuf s = { 0 };
    jb_s(&s, ix->unique || ix->primary ? "CREATE UNIQUE INDEX " : "CREATE INDEX ");
    put_qname(&s, o, ix->name); jb_s(&s, " ON "); put_qname(&s, o, table); jb_s(&s, " (");
    index_cols(o, ix, &s);
    jb_s(&s, ")");
    int r = server_ddl(o, s.p, src);
    jb_free(&s);
    return r;
}
/* TableDefs.Append in a database of ODBC: CREATE TABLE on the server, its indexes */
int dao_odbc_create_table(ddb *d, dtd *t)
{
    osrc *o = d->odbc;
    if (!t->name || !*t->name || t->fields->n == 0) { dao_raise(3264, "DAO.TableDefs", "No field defined--cannot append TableDef or Index."); return -1; }
    jbuf s = { 0 };
    jb_s(&s, "CREATE TABLE "); put_qname(&s, o, t->name); jb_s(&s, " (");
    for (int i = 0; i < t->fields->n; i++) {
        dfld *f = (dfld *)t->fields->items[i];
        char b[64];
        if (i) jb_s(&s, ", ");
        put_qname(&s, o, f->name); jb_s(&s, " "); jb_s(&s, server_type(o, f->type, f->size, b, sizeof b));
        if (f->required) jb_s(&s, " NOT NULL");
    }
    jb_s(&s, ")");
    int r = server_ddl(o, s.p, "DAO.TableDefs");
    jb_free(&s);
    for (int i = 0; !r && i < t->indexes->n; i++) r = server_index(o, t->name, (dix *)t->indexes->items[i], "DAO.TableDefs");
    osrc_forget(o, t->name);
    if (r) return -1;
    t->appended = 1;
    for (int i = 0; i < t->fields->n; i++) ((dfld *)t->fields->items[i])->appended = 1;
    for (int i = 0; i < t->indexes->n; i++) { ((dix *)t->indexes->items[i])->appended = 1; ((dix *)t->indexes->items[i])->remote = 1; }
    return osrc_table(o, t->name) ? 0 : -1;
}
int dao_odbc_drop_table(ddb *d, const char *name)
{
    osrc *o = d->odbc;
    jbuf s = { 0 };
    jb_s(&s, "DROP TABLE "); put_qname(&s, o, name);
    int r = server_ddl(o, s.p, "DAO.TableDefs");
    jb_free(&s);
    osrc_forget(o, name);
    return r;
}
int dao_odbc_create_index(ddb *d, const char *table, dix *ix)
{
    osrc *o = d->odbc;
    int r = server_index(o, table, ix, "DAO.Indexes");
    if (!r) ix->remote = 1;
    osrc_forget(o, table);
    osrc_table(o, table);
    return r;
}
int dao_odbc_drop_index(ddb *d, const char *table, const char *index)
{
    osrc *o = d->odbc;
    jbuf s = { 0 };
    int f = family(o);
    jb_s(&s, "DROP INDEX ");
    if (f == 1) { put_qname(&s, o, table); jb_s(&s, "."); }
    put_qname(&s, o, index);
    if (f == 2) { jb_s(&s, " ON "); put_qname(&s, o, table); }
    int r = server_ddl(o, s.p, "DAO.Indexes");
    jb_free(&s);
    osrc_forget(o, table);
    osrc_table(o, table);
    return r;
}
int dao_odbc_add_field(ddb *d, const char *table, dfld *fl)
{
    osrc *o = d->odbc;
    char b[64];
    jbuf s = { 0 };
    jb_s(&s, "ALTER TABLE "); put_qname(&s, o, table); jb_s(&s, " ADD ");
    put_qname(&s, o, fl->name); jb_s(&s, " "); jb_s(&s, server_type(o, fl->type, fl->size, b, sizeof b));
    int r = server_ddl(o, s.p, "DAO.Fields");
    jb_free(&s);
    osrc_forget(o, table);
    osrc_table(o, table);
    if (!r) fl->appended = 1;
    return r;
}
int dao_odbc_drop_field(ddb *d, const char *table, const char *field)
{
    osrc *o = d->odbc;
    jbuf s = { 0 };
    jb_s(&s, "ALTER TABLE "); put_qname(&s, o, table); jb_s(&s, " DROP COLUMN "); put_qname(&s, o, field);
    int r = server_ddl(o, s.p, "DAO.Fields");
    jb_free(&s);
    osrc_forget(o, table);
    osrc_table(o, table);
    return r;
}

/* SELECT ... INTO a table of a data source of ODBC ([ODBC;...].t, IN '' [ODBC;...]): made on the server, the rows
 * sent there. The rows affected, or -1 with the error raised */
long dao_odbc_select_into(ddb *d, jsql *j, vbvar *vals)
{
    osrc *o = ext_osrc(d, j->into_ext);
    if (osrc_table(o, j->into)) { dao_raise(3010, "DAO.Database", "Table '%s' already exists.", j->into); return -1; }
    sqlite3_stmt *st = 0;
    int rc = sqlite3_prepare_v2(d->db, j->sql, -1, &st, 0);
    if (rc != SQLITE_OK) { dao_raise_sqlite(d, rc, "DAO.Database"); return -1; }
    for (int i = 0; i < j->nbinds; i++) jet_bind(st, i + 1, vals[i], 0);
    int nc = sqlite3_column_count(st);
    int *ty = calloc((size_t)(nc ? nc : 1), sizeof(int));
    jbuf s = { 0 }, ins = { 0 };
    jb_s(&s, "CREATE TABLE "); put_qname(&s, o, j->into); jb_s(&s, " (");
    jb_s(&ins, "INSERT INTO "); put_qname(&ins, o, j->into); jb_s(&ins, " (");
    for (int i = 0; i < nc; i++) {
        int size, attrs;
        int type = jet_decl_type(sqlite3_column_decltype(st, i), &size, &attrs);
        rcol c; memset(&c, 0, sizeof c);
        const char *tn = sqlite3_column_table_name(st, i), *on = sqlite3_column_origin_name(st, i);
        const char *dn = sqlite3_column_database_name(st, i);
        schema_qualify(dn && !strncmp(dn, "lbext", 5) ? dn : 0);
        if (tn && on && schema_field_info(d, tn, on, &c)) { type = c.type; size = c.size; }
        schema_qualify(0);
        free(c.defval);
        if (!type) type = dbText, size = 255;
        ty[i] = type;
        char b[64];
        if (i) { jb_s(&s, ", "); jb_s(&ins, ", "); }
        put_qname(&s, o, sqlite3_column_name(st, i)); jb_s(&s, " "); jb_s(&s, server_type(o, type, size, b, sizeof b));
        put_qname(&ins, o, sqlite3_column_name(st, i));
    }
    jb_s(&s, ")");
    jb_s(&ins, ") VALUES (");
    for (int i = 0; i < nc; i++) jb_s(&ins, i ? ", ?" : "?");
    jb_s(&ins, ")");
    long n = 0;
    if (server_ddl(o, s.p, "DAO.Database")) n = -1;
    lbo_stmt *ps = n < 0 ? 0 : lbo_stmt_new(o->k);
    if (ps && lbo_prepare(ps, ins.p)) { odbc_fail(3146, "DAO.Database", "ODBC--call failed."); n = -1; }
    if (n >= 0) server_tran(o, 0);
    while (n >= 0 && sqlite3_step(st) == SQLITE_ROW) {
        vb_i32 mk = vb_mark();
        int e = 0;
        for (int i = 0; !e && i < nc; i++) e = ov_bind(ps, i + 1, jet_column(st, i, ty[i]));
        if (!e) e = lbo_execute(ps);
        vb_release(mk);
        if (e) { int k = lbo_nerr(); lbo_err keep[8]; if (k > 8) k = 8; memcpy(keep, lbo_errs(), sizeof(lbo_err) * (size_t)k);
                 server_tran(o, 2);
                 coll_clear_items(err_list());
                 for (int i = 0; i < k; i++) { char m[1100]; lb_odbc_msg(keep[i].msg, m, sizeof m); err_append(keep[i].native, m, "ODBC.Database"); }
                 char msg[600]; snprintf(msg, sizeof msg, "ODBC--insert on a linked table '%s' failed.", j->into);
                 err_append(3155, msg, "DAO.Database");
                 vb_raise_full(3155, vb_sfromutf8("DAO.Database"), vb_sfromutf8(msg));
                 n = -1; break; }
        n++;
    }
    if (n >= 0) server_tran(o, 1);
    lbo_stmt_free(ps);
    sqlite3_finalize(st);
    jb_free(&s); jb_free(&ins); free(ty);
    osrc_forget(o, j->into);
    return n;
}
int dao_odbc_into(ddb *d, jsql *j) { return j->into && ext_osrc(d, j->into_ext); }
