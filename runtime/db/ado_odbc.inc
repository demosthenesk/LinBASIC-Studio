/* ado_odbc.inc - the provider MSDASQL of ADO: ODBC data sources (SQL Server, MySQL / MariaDB,
 * PostgreSQL, SQLite, Oracle ... whatever has an ODBC driver), through odbc.c. Included by ado.c.
 *
 * The rows of a statement go into a SQLite database in memory, a recordset of its own, so that all
 * of the Recordset of ado.c works on them (moves, Find, Filter, Sort, bookmarks, events); a change
 * goes to the server first (ddb.remote), as the cursor engine of ADO does it: UPDATE / INSERT /
 * DELETE of the base table, the row found by its key (a client cursor) or by all of its values (a
 * server cursor: MSDASQL's query-based update). Types, sizes, attributes, errors and their texts as
 * VB6 gives them with MSDASQL (VB6 oracle tests/db/d_odbc1, with the SQLite ODBC driver).
 *
 * Connection strings: "DSN=...", "Driver={...};...", "Provider=MSDASQL;...", "Provider=SQLOLEDB;Data
 * Source=server;Initial Catalog=db" (SQL Server through its ODBC driver). On Linux a driver named as
 * on Windows ({SQL Server}, {MySQL ODBC 3.51 Driver} ...) is the one installed for that database.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "odbc.h"
#ifdef _WIN32
#include <direct.h>
#define getcwd _getcwd
#else
#include <unistd.h>
#endif
vbvar alist_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);

#define SRC_ODBC "Microsoft OLE DB Provider for ODBC Drivers"

typedef struct otarget {
    char *table;                        /* the base table on the server, or 0: not updatable */
    char **cols;                        /* the server column of each column of the recordset, or 0 */
    int ncols;
    int *iskey;                         /* the primary key */
    int autocol;                        /* an identity column, or -1 */
} otarget;

static void ot_free(otarget *t)
{
    if (!t) return;
    for (int i = 0; i < t->ncols; i++) free(t->cols[i]);
    free(t->cols); free(t->iskey); free(t->table); free(t);
}

/* ---------------------------------------------------------------- errors */
/* the texts of the driver manager as Windows writes them */
static void dm_text(const char *in, char *out, size_t n)
{
    const char *p = in;
    if (!strncmp(p, "[unixODBC][Driver Manager]", 26)) {
        p += 26;
        if (strstr(p, "Can't open lib"))       /* a driver that is not installed: as Windows says it */
            p = "Data source name not found and no default driver specified";
        snprintf(out, n, "[Microsoft][ODBC Driver Manager] %s", p);
        return;
    }
    if (!strncmp(p, "[unixODBC]", 10)) p += 10;
    if (!strncmp(p, "[SQLite]", 8)) p += 8;     /* the SQLite driver of Linux; that of Windows has no prefix */
    snprintf(out, n, "%s", p);
}
static vb_i32 odbc_hr(const char *state)
{
    if (!strcmp(state, "37000") || !strncmp(state, "42", 2)) return -2147217900;     /* DB_E_ERRORSINCOMMAND */
    if (!strcmp(state, "S0002") || !strcmp(state, "42S02")) return -2147217865;       /* DB_E_NOTABLE */
    if (!strncmp(state, "23", 2)) return -2147217873;                                 /* DB_E_INTEGRITYVIOLATION */
    if (!strcmp(state, "07001") || !strcmp(state, "07002")) return -2147217904;       /* DB_E_PARAMNOTOPTIONAL */
    if (!strncmp(state, "22", 2)) return -2147217913;                                 /* DB_E_CANTCONVERTVALUE */
    return -2147467259;                                                               /* E_FAIL */
}
/* the diagnostics of ODBC into Errors, the first one raised */
static void odbc_raise(acn *c)
{
    int n = lbo_nerr();
    const lbo_err *e = lbo_errs();
    if (c) alist_clear(c->errs);
    char first[1100] = "Unspecified error";
    vb_i32 hr = -2147467259;
    for (int i = 0; i < n; i++) {
        char msg[1100];
        dm_text(e[i].msg, msg, sizeof msg);
        const char *state = strstr(e[i].msg, "Can't open lib") ? "IM002" : e[i].state;
        vb_i32 h = odbc_hr(state);
        if (!i) { snprintf(first, sizeof first, "%s", msg); hr = h; }
        if (c) {
            aerr *x = (aerr *)vb_onew(&vb_cls_ado_error);
            x->number = h; x->native = e[i].native;
            x->desc = strdup(msg); x->src = strdup(SRC_ODBC); x->state = strdup(state);
            alist_add(c->errs, (vbobj)x);
        }
    }
    raise(hr, SRC_ODBC, first);
}

/* ---------------------------------------------------------------- the connection string */
static int has_key(const char *cs, const char *k) { char *v = cs_get(cs, k); int r = v != 0; free(v); return r; }

/* cs without the keys of ADO (Provider, Extended Properties ...) */
static void drop_keys(const char *cs, jbuf *out)
{
    const char *p = cs;
    while (p && *p) {
        while (*p == ';' || isspace((unsigned char)*p)) p++;
        if (!*p) break;
        const char *start = p, *eq = strchr(p, '=');
        if (!eq) break;
        const char *v = eq + 1;
        while (isspace((unsigned char)*v)) v++;
        const char *end;
        if (*v == '"' || *v == '\'') { char q = *v; end = strchr(v + 1, q); end = end ? end + 1 : v + strlen(v); }
        else if (*v == '{') { end = strchr(v, '}'); end = end ? end + 1 : v + strlen(v); }
        else end = v + strcspn(v, ";");
        size_t kl = (size_t)(eq - start);
        while (kl && isspace((unsigned char)start[kl - 1])) kl--;
        int skip = (kl == 8 && !strncasecmp(start, "Provider", 8)) || (kl == 19 && !strncasecmp(start, "Extended Properties", 19)) ||
                   (kl == 23 && !strncasecmp(start, "Persist Security Info", 21));
        if (!skip) { if (out->n) jb_s(out, ";"); jb_put(out, start, (int)(end - start)); }
        p = end;
    }
}

/* Linux: a driver named as on Windows -> one installed for the same database */
static char *map_driver(const char *want)
{
#ifdef _WIN32
    (void)want; return 0;
#else
    char list[8192];
    if (lbo_drivers(list, sizeof list)) return 0;
    for (char *l = list; *l;) {                 /* installed as it is */
        char *e = strchr(l, '\n'); if (!e) break;
        if ((size_t)(e - l) == strlen(want) && !strncasecmp(l, want, (size_t)(e - l))) return 0;
        l = e + 1;
    }
    static const char *const fam[][2] = {
        { "sql server", "for sql server" }, { "sql server", "freetds" }, { "sql native client", "for sql server" },
        { "sql native client", "freetds" }, { "sql server", "sql server" },
        { "mysql", "mysql" }, { "mysql", "mariadb" }, { "mariadb", "mariadb" }, { "mariadb", "mysql" },
        { "postgresql", "postgresql unicode" }, { "postgresql", "postgresql" },
        { "sqlite3", "sqlite3" }, { "sqlite", "sqlite" }, { "oracle", "oracle" }, { "firebird", "firebird" },
        { "db2", "db2" }, { "informix", "informix" }, { "sybase", "freetds" } };
    char low[256]; snprintf(low, sizeof low, "%s", want);
    for (char *q = low; *q; q++) *q = (char)tolower((unsigned char)*q);
    for (size_t f = 0; f < sizeof fam / sizeof fam[0]; f++) {
        if (!strstr(low, fam[f][0])) continue;
        for (char *l = list; *l;) {
            char *e = strchr(l, '\n'); if (!e) break;
            char name[256]; int n = (int)(e - l < 255 ? e - l : 255);
            memcpy(name, l, (size_t)n); name[n] = 0;
            char lo[256]; snprintf(lo, sizeof lo, "%s", name);
            for (char *q = lo; *q; q++) *q = (char)tolower((unsigned char)*q);
            if (strstr(lo, fam[f][1])) return strdup(name);
            l = e + 1;
        }
    }
    return 0;
#endif
}

/* Linux: the files a connection string names (Database=, DBQ= ...) as App.Path & "\x.db" gives them */
static char *unix_paths(char *cs)
{
#ifndef _WIN32
    static const char *const keys[] = { "Database=", "DBQ=", "DBNAME=", "FILEDSN=", "SAVEFILE=" };
    for (char *p = cs; p && *p; p++) {
        if (p != cs && p[-1] != ';') continue;
        while (*p == ' ') p++;
        for (size_t k = 0; k < sizeof keys / sizeof keys[0]; k++) {
            size_t n = strlen(keys[k]);
            if (strncasecmp(p, keys[k], n)) continue;
            char *v = p + n, *e = v + strcspn(v, ";");
            if (memchr(v, '/', (size_t)(e - v)) || (e - v > 2 && v[1] == ':'))
                for (char *q = v; q < e; q++) if (*q == '\\') *q = '/';
        }
    }
#endif
    return cs;
}

/* a file database (the SQLite driver) named without a folder: in the current folder of the program, as Windows
 * gives it to the driver (in Wine the driver runs where the folder of the program is not the current one) */
static char *full_db_path(char *cs)
{
    char *drv = cs_get(cs, "Driver");
    int file = drv && strcasestr(drv, "sqlite");
    free(drv);
    char *db = file ? cs_get(cs, "Database") : 0;
    if (!db || !*db || db[0] == '/' || db[0] == '\\' || (db[0] && db[1] == ':') || !strcmp(db, ":memory:")) { free(db); return cs; }
    char cwd[4096];
    if (!getcwd(cwd, sizeof cwd)) { free(db); return cs; }
#ifdef _WIN32
    const char *sep = "\\";
#else
    const char *sep = "/";
#endif
    jbuf n = { 0 };
    const char *p = cs;
    while (*p) {                                /* the keys as they are, Database with its folder */
        while (*p == ';') p++;
        if (!*p) break;
        const char *eq = strchr(p, '=');
        const char *v = eq ? eq + 1 : p + strlen(p);
        const char *end;
        if (*v == '{') { end = strchr(v, '}'); end = end ? end + 1 : v + strlen(v); end += strcspn(end, ";"); }
        else end = p + strcspn(p, ";");
        if (n.n) jb_s(&n, ";");
        if (eq && (size_t)(eq - p) == 8 && !strncasecmp(p, "Database", 8)) {
            jb_put(&n, p, (int)(v - p)); jb_s(&n, cwd); jb_s(&n, sep); jb_s(&n, db);
        } else jb_put(&n, p, (int)(end - p));
        p = end;
    }
    free(db);
    free(cs);
    return n.p ? n.p : strdup("");
}

/* the ODBC connection string of an ADO one, or 0 when it is not for ODBC */
static char *odbc_cs(const char *cs)
{
    char *prov = cs_get(cs, "Provider");
    jbuf o = { 0 };
    if (!prov || !strncasecmp(prov, "MSDASQL", 7)) {
        char *ext = cs_get(cs, "Extended Properties");
        if (ext && *ext) { jb_s(&o, ext); jbuf rest = { 0 }; drop_keys(cs, &rest); if (rest.n) { jb_s(&o, ";"); jb_s(&o, rest.p); } jb_free(&rest); }
        else drop_keys(cs, &o);
        free(ext);
        if (!o.n || (!has_key(o.p, "DSN") && !has_key(o.p, "Driver") && !has_key(o.p, "FILEDSN"))) { free(prov); jb_free(&o); return 0; }
    } else if (!strncasecmp(prov, "SQLOLEDB", 8) || !strncasecmp(prov, "SQLNCLI", 7) || !strncasecmp(prov, "MSOLEDBSQL", 10)) {
        /* SQL Server through its ODBC driver */
        char *srv = cs_get(cs, "Data Source"), *db = cs_get(cs, "Initial Catalog"), *uid = cs_get(cs, "User ID"),
             *pwd = cs_get(cs, "Password"), *sec = cs_get(cs, "Integrated Security");
        if (!srv) srv = cs_get(cs, "Server");
        if (!db) db = cs_get(cs, "Database");
        if (!uid) uid = cs_get(cs, "UID");
        if (!pwd) pwd = cs_get(cs, "PWD");
        jb_s(&o, "Driver={SQL Server}");
        if (srv) { jb_s(&o, ";Server="); jb_s(&o, srv); }
        if (db) { jb_s(&o, ";Database="); jb_s(&o, db); }
        if (uid) { jb_s(&o, ";UID="); jb_s(&o, uid); }
        if (pwd) { jb_s(&o, ";PWD="); jb_s(&o, pwd); }
        if (sec && !strcasecmp(sec, "SSPI")) jb_s(&o, ";Trusted_Connection=Yes");
        free(srv); free(db); free(uid); free(pwd); free(sec);
    } else if (!strncasecmp(prov, "MySQLProv", 9)) {
        char *srv = cs_get(cs, "Data Source"), *db = cs_get(cs, "Location"), *uid = cs_get(cs, "User ID"), *pwd = cs_get(cs, "Password");
        if (!db) db = cs_get(cs, "Initial Catalog");
        jb_s(&o, "Driver={MySQL ODBC 3.51 Driver}");
        if (srv) { jb_s(&o, ";Server="); jb_s(&o, srv); }
        if (db) { jb_s(&o, ";Database="); jb_s(&o, db); }
        if (uid) { jb_s(&o, ";UID="); jb_s(&o, uid); }
        if (pwd) { jb_s(&o, ";PWD="); jb_s(&o, pwd); }
        free(srv); free(db); free(uid); free(pwd);
    } else { free(prov); return 0; }
    free(prov);
    /* Access through ODBC: the Jet of LinBASIC Studio (DBQ) */
    char *drv = cs_get(o.p, "Driver");
    if (drv) {
        const char *d = drv; char nm[256]; size_t k = 0;
        if (*d == '{') d++;
        while (*d && *d != '}' && k < 255) nm[k++] = *d++;
        nm[k] = 0;
        if (strcasestr(nm, "Microsoft Access Driver")) { free(drv); jb_free(&o); return 0; }
        char *real = map_driver(nm);
        if (real) {                             /* the Driver key again, with the driver installed */
            jbuf n = { 0 };
            jb_s(&n, "Driver={"); jb_s(&n, real); jb_s(&n, "}");
            jbuf rest = { 0 };
            const char *p = o.p;
            while (p && *p) {
                while (*p == ';') p++;
                if (!*p) break;
                const char *end;
                const char *eq = strchr(p, '=');
                const char *v = eq ? eq + 1 : p + strlen(p);
                if (*v == '{') { end = strchr(v, '}'); end = end ? end + 1 : v + strlen(v); }
                else end = p + strcspn(p, ";");
                if (!(eq && (size_t)(eq - p) == 6 && !strncasecmp(p, "Driver", 6))) { jb_s(&rest, ";"); jb_put(&rest, p, (int)(end - p)); }
                p = end;
            }
            if (rest.n) jb_s(&n, rest.p);
            jb_free(&rest); jb_free(&o);
            o = n;
            free(real);
        }
        free(drv);
    }
    return full_db_path(unix_paths(o.p));
}

/* ---------------------------------------------------------------- Open, Close */
static int odbc_open(acn *c)
{
    char *ocs = odbc_cs(c->cs ? c->cs : "");
    if (!ocs) return 0;
    alist_clear(c->errs);
    char out[2048] = "";
    lbo_conn *k = lbo_connect(ocs, out, sizeof out);
    free(ocs);
    if (!k) { odbc_raise(c); return 1; }
    vbobj mdb = dao_memory_db();
    if (!mdb) { lbo_disconnect(k); raise(-2147467259, SRC_ODBC, "Unspecified error"); return 1; }
    c->db = mdb;
    c->odbc = k;
    free(c->provider); c->provider = strdup("MSDASQL.1");
    char b[512];
    free(c->dbms); c->dbms = strdup(lbo_info(k, LBI_DBMS_NAME, b, sizeof b) ? "" : b);
    c->odbc_ro = !lbo_info(k, LBI_DATA_SOURCE_READ_ONLY, b, sizeof b) && !strcasecmp(b, "Y");
    long conf = 0;
    lbo_info_int(k, LBI_ODBC_SQL_CONFORMANCE, &conf);
    c->odbc_sql = (conf >= 2 ? 7 : conf == 1 ? 3 : 1) | 0x100;   /* DBPROPVAL_SQL_ODBC_* | ESCAPECLAUSES */
    if (!lbo_info(k, LBI_DATABASE_NAME, b, sizeof b)) { free(c->defdb); c->defdb = strdup(b); }
    jbuf ncs = { 0 };
    jb_s(&ncs, "Provider=MSDASQL.1;Extended Properties=\""); jb_s(&ncs, out); jb_s(&ncs, "\"");
    free(c->cs); c->cs = ncs.p;
    c->state = 1;
    c->trans = 0;
    return 1;
}
static void odbc_close(acn *c)
{
    if (!c->odbc) return;
    if (c->trans) { lbo_endtran(c->odbc, 0); lbo_autocommit(c->odbc, 1); }
    lbo_disconnect(c->odbc);
    c->odbc = 0;
}

/* the Properties of the connection that programs read */
static void odbc_props(acn *c, alist *l);

/* ---------------------------------------------------------------- values */
static vbvar val_of(lbo_val *v)
{
    switch (v->kind) {
    case LBO_NULL: return vb_v_null();
    case LBO_INT: return v->iv >= -2147483647LL - 1 && v->iv <= 2147483647LL ? vb_v_l((vb_i32)v->iv) : vb_v_d((double)v->iv);
    case LBO_BOOL: return vb_v_bool(v->iv ? -1 : 0);
    case LBO_DBL: return vb_v_d(v->dv);
    case LBO_DEC: {                             /* NUMERIC / DECIMAL: a Variant Decimal, as ADO gives it */
        vbvar d = vb_dec_from_ascii(v->s ? v->s : "0");
        return d.vt == VT_DECIMAL ? d : vb_v_d(strtod(v->s ? v->s : "0", 0));
    }
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
/* a value of the program for a parameter of ODBC; adotype: of the column it goes to, or 0 */
static int bind_val(lbo_stmt *st, int i, vbvar x, int adotype)
{
    vbvar v = vb_v_letval(x);
    char b[64];
    switch (v.vt) {
    case VT_EMPTY: case VT_NULL: return lbo_bind(st, i, LBO_NULL, 0, 0, 0, 0, 0);
    case VT_I2: case VT_I4: case VT_UI1: return lbo_bind(st, i, LBO_INT, vb_v2l(v), 0, 0, 0, 0);
    case VT_BOOL: return lbo_bind(st, i, LBO_BOOL, vb_v2bool(v) ? 1 : 0, 0, 0, 0, 0);
    case VT_R4: case VT_R8: return lbo_bind(st, i, LBO_DBL, 0, vb_v2d(v), 0, 0, 0);
    case VT_CY: snprintf(b, sizeof b, "%.4f", vb_v2d(v)); return lbo_bind(st, i, LBO_DEC, 0, 0, b, 0, 0);
    case VT_DATE: {
        vb_date d = vb_v2date(v);
        if (adotype == adDBDate) {
            snprintf(b, sizeof b, "%04d-%02d-%02d", vb_year(d), vb_month(d), vb_day(d));
            return lbo_bind(st, i, LBO_DATE, 0, 0, b, 0, LBS_DATE);
        }
        if (adotype == adDBTime) {
            snprintf(b, sizeof b, "%02d:%02d:%02d", vb_hour(d), vb_minute(d), vb_second(d));
            return lbo_bind(st, i, LBO_DATE, 0, 0, b, 0, LBS_TIME);
        }
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

/* ---------------------------------------------------------------- the fields: types as MSDASQL gives them */
/* the SQLite driver of Windows gives its text as Unicode (SQL_WVARCHAR ...), that of Linux as SQL_VARCHAR */
static int odbc_wide;
static void set_wide(acn *c)
{
#ifdef _WIN32
    (void)c; odbc_wide = 0;
#else
    odbc_wide = c && c->dbms && !strncasecmp(c->dbms, "SQLite", 6);
#endif
}
static int ado_of_sql(const lbo_colinfo *c)
{
    switch (c->sqltype) {
    case LBS_CHAR: return odbc_wide ? adWChar : adChar;
    case LBS_VARCHAR: return odbc_wide ? adVarWChar : adVarChar;
    case LBS_LONGVARCHAR: return odbc_wide ? adLongVarWChar : adLongVarChar;
    case LBS_WCHAR: return adWChar;
    case LBS_WVARCHAR: return adVarWChar;
    case LBS_WLONGVARCHAR: return adLongVarWChar;
    case LBS_INTEGER: return c->unsignd ? adUnsignedInt : adInteger;
    case LBS_SMALLINT: return c->unsignd ? adUnsignedSmallInt : adSmallInt;
    case LBS_TINYINT: return c->unsignd ? adUnsignedTinyInt : adTinyInt;
    case LBS_BIGINT: return c->unsignd ? 21 : adBigInt;
    case LBS_REAL: return adSingle;
    case LBS_FLOAT: case LBS_DOUBLE: return adDouble;
    case LBS_NUMERIC: case LBS_DECIMAL: return adNumeric;
    case LBS_BIT: return adBoolean;
    case LBS_DATE: case LBS_TYPE_DATE: return adDBDate;
    case LBS_TIME: case LBS_TYPE_TIME: return adDBTime;
    case LBS_TIMESTAMP: case LBS_TYPE_TIMESTAMP: return adDBTimeStamp;
    case LBS_BINARY: return adBinary;
    case LBS_VARBINARY: return adVarBinary;
    case LBS_LONGVARBINARY: return adLongVarBinary;
    case LBS_GUID: return adGUID;
    }
    return adVarChar;
}
static int is_text_type(int t) { return t == adChar || t == adVarChar || t == adLongVarChar || t == adWChar || t == adVarWChar || t == adLongVarWChar; }
static int is_bin_type(int t) { return t == adBinary || t == adVarBinary || t == adLongVarBinary; }
static int is_long_type(int t) { return t == adLongVarChar || t == adLongVarWChar || t == adLongVarBinary; }
static void def_of(const lbo_colinfo *c, adef *d)
{
    int t = ado_of_sql(c);
    d->type = t;
    switch (t) {
    case adInteger: case adUnsignedInt: case adSingle: d->size = 4; break;
    case adSmallInt: case adUnsignedSmallInt: case adBoolean: d->size = 2; break;
    case adTinyInt: case adUnsignedTinyInt: d->size = 1; break;
    case adBigInt: case 21: case adDouble: case adCurrency: d->size = 8; break;
    case adDBDate: case adDBTime: d->size = 6; break;
    case adDBTimeStamp: d->size = 16; break;
    case adNumeric: d->size = 19; break;
    case adGUID: d->size = 16; break;
    default: d->size = (int)c->size; break;
    }
    int numeric = t == adInteger || t == adUnsignedInt || t == adSmallInt || t == adUnsignedSmallInt || t == adTinyInt ||
                  t == adUnsignedTinyInt || t == adBigInt || t == 21 || t == adSingle || t == adDouble || t == adNumeric;
    d->prec = numeric ? (int)c->size : 255;
    d->scale = t == adNumeric ? c->decimals : 255;
    int a = c->updatable ? adFldUpdatable : 0;
    if (c->nullable) a |= adFldIsNullable | adFldMayBeNull;
    if (is_long_type(t)) a |= adFldLong | adFldMayDefer;
    else if (!is_text_type(t) && !is_bin_type(t)) a |= adFldFixed;
    else if (t == adChar || t == adWChar || t == adBinary) a |= adFldFixed;
    d->attrs = a;
}

/* ---------------------------------------------------------------- where the changes go */
static int cursor_write(void *ctx, drs *d, int op);

/* the base table of a SELECT: the one table its columns come from, or the first of FROM */
static char *base_of(const char *sql, lbo_colinfo *ci, int n)
{
    const char *t = 0;
    for (int i = 0; i < n; i++) {
        if (!ci[i].basetable[0]) continue;
        if (!t) t = ci[i].basetable;
        else if (strcasecmp(t, ci[i].basetable)) return 0;      /* a join: not updatable */
    }
    if (t) return strdup(t);
    const char *f = sql;
    for (const char *p = sql; *p; p++)
        if (!strncasecmp(p, " from ", 6)) { f = p + 6; break; }
    if (f == sql) return 0;
    while (isspace((unsigned char)*f)) f++;
    char nm[256]; int k = 0;
    if (*f == '[' || *f == '"' || *f == '`') {
        char q = *f == '[' ? ']' : *f; f++;
        while (*f && *f != q && k < 255) nm[k++] = *f++;
    } else while (*f && (isalnum((unsigned char)*f) || *f == '_' || *f == '.' || *f == '$') && k < 255) nm[k++] = *f++;
    nm[k] = 0;
    if (!k) return 0;
    while (*f && *f != ' ' && *f != ',') f++;
    const char *rest = f;
    while (isspace((unsigned char)*rest)) rest++;
    if (*rest == ',' || !strncasecmp(rest, "join", 4) || strcasestr(rest, " join ")) return 0;
    return strdup(nm);
}
static void target_of(ars *r, acn *c, const char *sql, lbo_colinfo *ci, int n)
{
    ot_free(r->ot); r->ot = 0;
    char *tab = base_of(sql, ci, n);
    if (!tab) return;
    otarget *t = calloc(1, sizeof *t);
    t->table = tab; t->ncols = n; t->autocol = -1;
    t->cols = calloc((size_t)(n ? n : 1), sizeof(char *));
    t->iskey = calloc((size_t)(n ? n : 1), sizeof(int));
    int known = 0;                              /* the driver says where the columns come from */
    for (int i = 0; i < n; i++) if (ci[i].basetable[0]) known = 1;
    for (int i = 0; i < n; i++) {
        const char *bc = ci[i].basecol[0] ? ci[i].basecol : ci[i].name;
        if (known && (!ci[i].basetable[0] || strcasecmp(ci[i].basetable, tab))) continue;
        t->cols[i] = strdup(bc);
        if (ci[i].autoinc && t->autocol < 0) t->autocol = i;
    }
    /* the primary key */
    lbo_stmt *st = lbo_stmt_new(c->odbc);
    if (st && !lbo_primary_keys(st, 0, 0, tab)) {
        while (lbo_fetch(st) == 1) {
            lbo_val v;
            if (lbo_get(st, 4, LBS_VARCHAR, &v) == 0 && v.s)
                for (int i = 0; i < n; i++) if (t->cols[i] && !strcasecmp(t->cols[i], v.s)) t->iskey[i] = 1;
            lbo_val_free(&v);
        }
    }
    lbo_stmt_free(st);
    r->ot = t;
}

/* the rows of a statement that returned them: into r, a recordset of its own in memory */
static int odbc_fill(ars *r, acn *c, lbo_stmt *st, int nc, const char *sql)
{
    set_wide(c);
    lbo_colinfo *ci = calloc((size_t)nc, sizeof *ci);
    adef *defs = calloc((size_t)nc, sizeof *defs);
    for (int i = 0; i < nc; i++) {
        if (lbo_colinfo_get(st, i + 1, &ci[i])) { free(ci); free(defs); odbc_raise(c); return -1; }
        def_of(&ci[i], &defs[i]);
        char nm[300];
        if (ci[i].name[0]) snprintf(nm, sizeof nm, "%s", ci[i].name); else snprintf(nm, sizeof nm, "Expr%d", 1000 + i);
        defs[i].name = strdup(nm);
    }
    vbobj keepcn = r->cn; if (keepcn) vb_oaddref(keepcn);
    int e = make_memory(r, defs, nc);
    if (keepcn) vb_orelease(keepcn);
    if (!e) {
        for (int i = 0; i < nc; i++) { r->defs[i].prec = defs[i].prec; r->defs[i].scale = defs[i].scale; }
        drs *m = (drs *)vb_oid(r->base);
        ddb *d = DB(r->db);
        jbuf s = { 0 };
        jb_s(&s, "INSERT INTO lbrs VALUES (");
        for (int k = 0; k < nc; k++) jb_f(&s, k ? ", ?%d" : "?%d", k + 1);
        jb_s(&s, ")");
        sqlite3_stmt *ins = 0;
        sqlite3_prepare_v2(d->db, s.p, -1, &ins, 0);
        jb_free(&s);
        int f, max = r->maxrecords > 0 ? r->maxrecords : -1, got = 0;
        while (ins && (max < 0 || got < max) && (f = lbo_fetch(st)) == 1) {
            vb_i32 mk = vb_mark();
            sqlite3_reset(ins);
            for (int k = 0; k < nc; k++) {
                lbo_val v;
                if (lbo_get(st, k + 1, ci[k].sqltype, &v)) { v.kind = LBO_NULL; v.s = 0; }
                vbvar x = val_of(&v);
                jet_bind(ins, k + 1, x, m->cols[k].type);
                lbo_val_free(&v);
            }
            sqlite3_step(ins);
            vb_release(mk);
            got++;
        }
        if (ins) sqlite3_finalize(ins);
        int f2;
        rs_member(r->base, "requery", VBM_METHOD, 0, 0, &f2);
        if (vb_errpend) vb_err_clear_pending();
        rebuild_view(r);
        r->odbc = 1;
        r->ocount = got;
        r->pdel = 0;
        target_of(r, c, sql, ci, nc);
        d->remote = cursor_write;
        d->remote_ctx = r;
    }
    for (int i = 0; i < nc; i++) free(defs[i].name);
    free(defs); free(ci);
    return e;
}

/* run a statement: its rows into r (when it has some and r is given), else the rows affected.
 * *rows: it returned rows. 0, or -1 with the error raised */
static int odbc_run(acn *c, const char *sql, vbvar *vals, int nvals, ars *r, long *affected, int *rows)
{
    if (affected) *affected = -1;
    if (rows) *rows = 0;
    alist_clear(c->errs);
    lbo_stmt *st = lbo_stmt_new(c->odbc);
    if (!st) { odbc_raise(c); return -1; }
    int e;
    if (nvals > 0) {
        e = lbo_prepare(st, sql);
        for (int i = 0; !e && i < nvals; i++) e = bind_val(st, i + 1, vals[i], 0);
        if (!e) e = lbo_execute(st);
    } else e = lbo_exec_direct(st, sql);
    if (e) { odbc_raise(c); lbo_stmt_free(st); return -1; }
    int nc = lbo_ncols(st);
    if (nc <= 0) {
        if (affected) { long n = lbo_rowcount(st); *affected = n < 0 ? 0 : n; }
        lbo_stmt_free(st);
        return 0;
    }
    if (rows) *rows = 1;
    e = r ? odbc_fill(r, c, st, nc, sql) : 0;
    lbo_stmt_free(st);
    return e;
}

/* the SQL of a CommandText */
static char *odbc_sql(const char *text, int type, int nparams, char **src)
{
    int t = type & ~(128 | 16 | 32 | 64);
    jbuf s = { 0 };
    if (t == adCmdTable || t == adCmdTableDirect) {
        jb_s(&s, "select * from "); jb_s(&s, text);
        *src = strdup(s.p);
        return s.p;
    }
    if (t == adCmdStoredProc) {
        jb_s(&s, "{call "); jb_s(&s, text);
        if (nparams > 0) { jb_s(&s, "("); for (int i = 0; i < nparams; i++) jb_s(&s, i ? ",?" : "?"); jb_s(&s, ")"); }
        jb_s(&s, "}");
        *src = strdup(text);
        return s.p;
    }
    /* unknown: SQL, or a bare name of a table */
    const char *p = text;
    while (isspace((unsigned char)*p)) p++;
    int bare = *p && *p != '{' && *p != '(';
    for (const char *q = p; *q && bare; q++) if (isspace((unsigned char)*q)) bare = 0;
    if (bare && t == adCmdUnknown && !looks_sql(p)) { jb_s(&s, "select * from "); jb_s(&s, p); *src = strdup(s.p); return s.p; }
    *src = strdup(text);
    return strdup(text);
}

/* ---------------------------------------------------------------- a change of a row, to the server */
static int cursor_write(void *ctx, drs *d, int op)
{
    /* As the VB6 oracles (tests/db/d_odbc1, d_odbc2) show it:
     * a client cursor (the cursor engine, Update Criteria adCriteriaUpdCols): the row by its key and
     *   the old values of the columns changed (Null as IS NULL); a delete by the key; the counter of
     *   a new row is not read back (Null until Requery);
     * a server cursor (MSDASQL's query-based changes): the row by the values of all its columns that
     *   are not long, so that a Null matches nothing; a delete compares the long columns too, which
     *   never matches: a table with a long column (TEXT, MEMO, IMAGE ...) cannot delete through it. */
    ars *r = (ars *)ctx;
    acn *c = cn_of(r);
    otarget *t = r->ot;
    if (!c || !c->state || (!c->odbc && !c->db)) { raise(3709, SRC_RS, MSG_NOCONN); remote_err = 1; return -1; }
    if (!t || !t->table) { raise(-2147467259, "Microsoft Cursor Engine", "Insufficient base table information for updating or refreshing."); remote_err = 1; return -1; }
    int jet = !c->odbc;                         /* the rows of Jet under MSDataShape: Jet SQL, [names] */
#define QN(n_) do { if (jet) jb_s(&s, "["); jb_s(&s, n_); if (jet) jb_s(&s, "]"); } while (0)
    int client = r->cursorloc == adUseClient;
    int nkey = 0, haslong = 0;
    for (int i = 0; i < t->ncols; i++) { if (t->iskey[i] && t->cols[i]) nkey++; if (t->cols[i] && is_long_type(r->defs[i].type)) haslong = 1; }
    jbuf s = { 0 };
    vbvar *vals = calloc((size_t)(t->ncols * 2 + 2), sizeof(vbvar));
    int *vty = calloc((size_t)(t->ncols * 2 + 2), sizeof(int));
    int nv = 0;
    vbvar *old = op != 1 && d->cur >= 0 && d->cur < d->d->n ? d->d->r[d->cur].v : 0;
    if (op == 1) {
        jb_s(&s, "INSERT INTO "); QN(t->table); jb_s(&s, " (");
        int k = 0;
        for (int i = 0; i < t->ncols; i++) {
            if (!t->cols[i] || i == t->autocol || !d->edirty[i]) continue;
            jb_s(&s, k ? ", " : ""); QN(t->cols[i]);
            vty[nv] = r->defs[i].type; vals[nv++] = d->ebuf[i]; k++;
        }
        jb_s(&s, ") VALUES (");
        for (int i = 0; i < k; i++) jb_s(&s, i ? ", ?" : "?");
        jb_s(&s, ")");
        if (!k) { s.n = 0; jb_s(&s, "INSERT INTO "); QN(t->table); jb_s(&s, " DEFAULT VALUES"); }
    } else {
        if (op == 2) {
            jb_s(&s, "UPDATE "); QN(t->table); jb_s(&s, " SET ");
            int k = 0;
            for (int i = 0; i < t->ncols; i++) {
                if (!d->edirty[i] || !t->cols[i]) continue;     /* an expression: not sent */
                jb_s(&s, k ? ", " : ""); QN(t->cols[i]); jb_s(&s, " = ?");
                vty[nv] = r->defs[i].type; vals[nv++] = d->ebuf[i]; k++;
            }
            if (!k) { free(vals); free(vty); jb_free(&s); return 0; }
        } else { jb_s(&s, "DELETE FROM "); QN(t->table); }
        jb_s(&s, " WHERE ");
        int k = 0;
        for (int i = 0; i < t->ncols && old; i++) {
            if (!t->cols[i]) continue;
            int use;
            if (client) use = nkey ? t->iskey[i] || (op == 2 && d->edirty[i]) : !is_long_type(r->defs[i].type);
            else use = !is_long_type(r->defs[i].type);
            if (!use) continue;
            jb_s(&s, k ? " AND " : ""); QN(t->cols[i]);
            if (client && (old[i].vt == VT_NULL || old[i].vt == VT_EMPTY)) jb_s(&s, " IS NULL");
            else { jb_s(&s, " = ?"); vty[nv] = r->defs[i].type; vals[nv++] = old[i]; }
            k++;
        }
        if (!k) jb_s(&s, "1 = 0");
        else if (!client && op == 3 && haslong) jb_s(&s, " AND 1 = 0");
    }
#undef QN
    alist_clear(c->errs);
    long n;
    if (jet) {
        vbobj o = 0;
        int e = run_sql(c, DB(c->db), s.p, vals, nv, dbOpenSnapshot, &o, &n);
        jb_free(&s); free(vals); free(vty);
        if (o) vb_orelease(o);
        if (e) { remote_err = 1; return -1; }
    } else {
        lbo_stmt *st = lbo_stmt_new(c->odbc);
        int e = !st || lbo_prepare(st, s.p);
        for (int i = 0; !e && i < nv; i++) e = bind_val(st, i + 1, vals[i], vty[i]);
        if (!e) e = lbo_execute(st);
        jb_free(&s); free(vals); free(vty);
        if (e) { odbc_raise(c); lbo_stmt_free(st); remote_err = 1; return -1; }
        n = lbo_rowcount(st);
        lbo_stmt_free(st);
    }
    if (op != 1 && n == 0) {
        remote_err = 1;
        if (client) raise(-2147217864, "Microsoft Cursor Engine", "Row cannot be located for updating. Some values may have been changed since it was last read.");
        else if (op == 2) raise(-2147467259, SRC_ODBC, "Query-based update failed because the row to update could not be found.");
        else { raise(-2147467259, SRC_ODBC, "Query-based delete failed because the row to delete could not be found."); r->pdel = 1; }
        return -1;
    }
    /* a new row of a server cursor: the number its identity column got */
    if (op == 1 && !client && t->autocol >= 0) {
        const char *q = 0, *db = c->dbms ? c->dbms : "";
        if (strcasestr(db, "sqlite")) q = "SELECT last_insert_rowid()";
        else if (strcasestr(db, "mysql") || strcasestr(db, "mariadb")) q = "SELECT LAST_INSERT_ID()";
        else if (strcasestr(db, "postgres")) q = "SELECT lastval()";
        else if (strcasestr(db, "sql server") || strcasestr(db, "sybase")) q = "SELECT @@IDENTITY";
        lbo_stmt *qs = q ? lbo_stmt_new(c->odbc) : 0;
        if (qs && !lbo_exec_direct(qs, q) && lbo_fetch(qs) == 1) {
            lbo_val v;
            if (!lbo_get(qs, 1, LBS_BIGINT, &v)) vb_vset(&d->ebuf[t->autocol], val_of(&v));
            lbo_val_free(&v);
        }
        lbo_stmt_free(qs);
    }
    return 0;
}

/* a delete of a server cursor that failed stays pending: what moves on or changes the recordset
 * tries it again first (VB6 oracle d_odbc2). 0, or -1 with the error raised */
static int odbc_pending(ars *r)
{
    if (!r->pdel) return 0;
    drs *v = VW(r);
    if (cursor_write(r, v, 3)) { remote_err = 0; return -1; }
    r->pdel = 0;
    ddb *d = DB(r->db);
    int (*keep)(void *, drs *, int) = d->remote;
    d->remote = 0;
    int f;
    rs_member(r->view, "delete", VBM_METHOD, 0, 0, &f);
    d->remote = keep;
    return vb_errpend ? -1 : 0;
}

/* ---------------------------------------------------------------- transactions */
static vbvar odbc_trans(acn *c, int what)       /* 0 begin, 1 commit, 2 rollback */
{
    alist_clear(c->errs);
    if (what == 0) {
        if (c->trans) { raise(-2147168227, SRC_ODBC, "Only one transaction can be active on this session."); return vb_v_empty(); }
        if (lbo_autocommit(c->odbc, 0)) { odbc_raise(c); return vb_v_empty(); }
        c->trans = 1;
        return vb_v_l(1);
    }
    if (!c->trans) { raise(-2147168242, SRC_ODBC, "No transaction is active."); return vb_v_empty(); }
    int e = lbo_endtran(c->odbc, what == 1);
    lbo_autocommit(c->odbc, 1);
    c->trans = 0;
    if (e) odbc_raise(c);
    return vb_v_empty();
}

/* ---------------------------------------------------------------- OpenSchema */
static int schema_type(int sqltype)             /* DATA_TYPE of the schema rowsets: the OLE DB type */
{
    lbo_colinfo ci; memset(&ci, 0, sizeof ci);
    ci.sqltype = sqltype;
    int t = ado_of_sql(&ci);
    if (t == adChar || t == adVarChar || t == adLongVarChar) return adChar;
    if (t == adWChar || t == adVarWChar || t == adLongVarWChar) return adWChar;
    if (t == adVarBinary || t == adLongVarBinary) return adBinary;
    return t;
}
static vbvar odbc_openschema(acn *c, int what, const char *rtable)
{
    set_wide(c);
    ars *r = (ars *)rs_new_for(0);
    r->cursortype = adOpenStatic; r->locktype = adLockReadOnly; r->cursorloc = adUseClient;
    lbo_stmt *st = lbo_stmt_new(c->odbc);
    int e = !st;
    if (what == 20) {
        adef defs[] = { { "TABLE_CATALOG", adVarWChar, 255, 0 }, { "TABLE_SCHEMA", adVarWChar, 255, 0 },
                        { "TABLE_NAME", adVarWChar, 255, 0 }, { "TABLE_TYPE", adVarWChar, 255, 0 },
                        { "TABLE_GUID", adGUID, 16, 0 }, { "DESCRIPTION", adVarWChar, 255, 0 },
                        { "TABLE_PROPID", adInteger, 4, 0 }, { "DATE_CREATED", adDate, 8, 0 },
                        { "DATE_MODIFIED", adDate, 8, 0 } };
        if (!e) e = lbo_tables(st, 0, 0, rtable, 0);
        if (e) { odbc_raise(c); lbo_stmt_free(st); vb_orelease((vbobj)r); return vb_v_empty(); }
        if (make_memory(r, defs, 9)) { lbo_stmt_free(st); vb_orelease((vbobj)r); return vb_v_empty(); }
        sqlite3_stmt *ins = 0;
        sqlite3_prepare_v2(DB(r->db)->db, "INSERT INTO lbrs (TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE, DESCRIPTION) VALUES (?1, ?2, ?3, ?4, ?5)", -1, &ins, 0);
        while (ins && lbo_fetch(st) == 1) {
            sqlite3_reset(ins);
            int map[5] = { 1, 2, 3, 4, 5 };
            for (int k = 0; k < 5; k++) {
                lbo_val v;
                if (!lbo_get(st, map[k], LBS_VARCHAR, &v) && v.kind != LBO_NULL) sqlite3_bind_text(ins, k + 1, v.s, -1, SQLITE_TRANSIENT);
                else sqlite3_bind_null(ins, k + 1);
                lbo_val_free(&v);
            }
            sqlite3_step(ins);
        }
        if (ins) sqlite3_finalize(ins);
    } else if (what == 4) {
        adef defs[] = { { "TABLE_CATALOG", adVarWChar, 255, 0 }, { "TABLE_SCHEMA", adVarWChar, 255, 0 },
                        { "TABLE_NAME", adVarWChar, 255, 0 }, { "COLUMN_NAME", adVarWChar, 255, 0 },
                        { "COLUMN_GUID", adGUID, 16, 0 }, { "COLUMN_PROPID", adInteger, 4, 0 },
                        { "ORDINAL_POSITION", adInteger, 4, 0 }, { "COLUMN_HASDEFAULT", adBoolean, 2, 0 },
                        { "COLUMN_DEFAULT", adVarWChar, 255, 0 }, { "COLUMN_FLAGS", adInteger, 4, 0 },
                        { "IS_NULLABLE", adBoolean, 2, 0 }, { "DATA_TYPE", adSmallInt, 2, 0 },
                        { "TYPE_GUID", adGUID, 16, 0 }, { "CHARACTER_MAXIMUM_LENGTH", adInteger, 4, 0 },
                        { "CHARACTER_OCTET_LENGTH", adInteger, 4, 0 }, { "NUMERIC_PRECISION", adSmallInt, 2, 0 },
                        { "NUMERIC_SCALE", adSmallInt, 2, 0 }, { "DATETIME_PRECISION", adInteger, 4, 0 },
                        { "CHARACTER_SET_CATALOG", adVarWChar, 255, 0 }, { "CHARACTER_SET_SCHEMA", adVarWChar, 255, 0 },
                        { "CHARACTER_SET_NAME", adVarWChar, 255, 0 }, { "COLLATION_CATALOG", adVarWChar, 255, 0 },
                        { "COLLATION_SCHEMA", adVarWChar, 255, 0 }, { "COLLATION_NAME", adVarWChar, 255, 0 },
                        { "DOMAIN_CATALOG", adVarWChar, 255, 0 }, { "DOMAIN_SCHEMA", adVarWChar, 255, 0 },
                        { "DOMAIN_NAME", adVarWChar, 255, 0 }, { "DESCRIPTION", adVarWChar, 255, 0 } };
        if (!e) e = lbo_columns(st, 0, 0, rtable, 0);
        if (e) { odbc_raise(c); lbo_stmt_free(st); vb_orelease((vbobj)r); return vb_v_empty(); }
        if (make_memory(r, defs, 28)) { lbo_stmt_free(st); vb_orelease((vbobj)r); return vb_v_empty(); }
        sqlite3_stmt *ins = 0;
        sqlite3_prepare_v2(DB(r->db)->db, "INSERT INTO lbrs (TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, ORDINAL_POSITION, "
                           "COLUMN_HASDEFAULT, COLUMN_DEFAULT, IS_NULLABLE, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION, "
                           "NUMERIC_SCALE, DESCRIPTION) VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13)", -1, &ins, 0);
        char lastt[256] = ""; int ord = 0;
        while (ins && lbo_fetch(st) == 1) {
            sqlite3_reset(ins);
            lbo_val v[18];
            for (int k = 0; k < 18; k++) {
                int ty = (k == 4 || k == 6 || k == 7 || k == 8 || k == 9 || k == 10 || k == 15 || k == 16) ? LBS_INTEGER : LBS_VARCHAR;
                if (lbo_get(st, k + 1, ty, &v[k])) { v[k].kind = LBO_NULL; v[k].s = 0; }
            }
            /* SQLColumns: 1 CAT 2 SCHEM 3 TABLE 4 COLUMN 5 DATA_TYPE 7 COLUMN_SIZE 9 DECIMAL_DIGITS 11 NULLABLE 12 REMARKS
               13 COLUMN_DEF 17 ORDINAL_POSITION */
            for (int k = 0; k < 4; k++) { if (v[k].kind != LBO_NULL && v[k].s) sqlite3_bind_text(ins, k + 1, v[k].s, -1, SQLITE_TRANSIENT); else sqlite3_bind_null(ins, k + 1); }
            /* ORDINAL_POSITION (ODBC 3), or counted */
            if (!v[2].s || strcmp(lastt, v[2].s)) { snprintf(lastt, sizeof lastt, "%s", v[2].s ? v[2].s : ""); ord = 0; }
            ord++;
            sqlite3_bind_int64(ins, 5, v[16].kind != LBO_NULL && v[16].iv > 0 ? v[16].iv : ord);
            sqlite3_bind_int(ins, 6, v[12].kind != LBO_NULL ? -1 : 0);
            if (v[12].kind != LBO_NULL && v[12].s) sqlite3_bind_text(ins, 7, v[12].s, -1, SQLITE_TRANSIENT); else sqlite3_bind_null(ins, 7);
            sqlite3_bind_int(ins, 8, v[10].kind != LBO_NULL && v[10].iv != 0 ? -1 : 0);
            int sty = v[4].kind == LBO_NULL ? 0 : (int)v[4].iv;
            int ty = schema_type(sty);
            sqlite3_bind_int(ins, 9, ty);
            if (ty == adChar || ty == adWChar || ty == adBinary) sqlite3_bind_int64(ins, 10, v[6].iv); else sqlite3_bind_null(ins, 10);
            if (ty != adChar && ty != adWChar && ty != adBinary && ty != adDBDate && ty != adDBTime && ty != adDBTimeStamp && ty != adBoolean)
                sqlite3_bind_int64(ins, 11, v[6].iv); else sqlite3_bind_null(ins, 11);
            if (ty == adNumeric) sqlite3_bind_int64(ins, 12, v[8].iv); else sqlite3_bind_null(ins, 12);
            if (v[11].kind != LBO_NULL && v[11].s && *v[11].s) sqlite3_bind_text(ins, 13, v[11].s, -1, SQLITE_TRANSIENT); else sqlite3_bind_null(ins, 13);
            sqlite3_step(ins);
            for (int k = 0; k < 18; k++) lbo_val_free(&v[k]);
        }
        if (ins) sqlite3_finalize(ins);
    } else {
        lbo_stmt_free(st);
        vb_orelease((vbobj)r);
        raise(3251, SRC_CN, "Object or provider is not capable of performing requested operation.");
        return vb_v_empty();
    }
    lbo_stmt_free(st);
    int f2;
    rs_member(r->base, "requery", VBM_METHOD, 0, 0, &f2);
    if (vb_errpend) vb_err_clear_pending();
    rebuild_view(r);
    r->state = 1;
    vbvar v = retobj((vbobj)r);
    vb_orelease((vbobj)r);
    return v;
}

/* ---------------------------------------------------------------- Parameters.Refresh */
static void odbc_refresh_params(acmd *cmd)
{
    acn *c = cmd->cn ? ACN(cmd->cn) : 0;
    if (!c || !c->odbc || !cmd->text) return;
    set_wide(c);
    char *src = 0;
    char *sql = odbc_sql(cmd->text, cmd->type, 0, &src);
    lbo_stmt *st = lbo_stmt_new(c->odbc);
    if (st && !lbo_prepare(st, sql)) {
        int n = lbo_nparams(st);
        alist_clear(cmd->params);
        for (int i = 1; i <= n; i++) {
            apar *p = (apar *)vb_onew(&vb_cls_ado_parameter);
            char nm[32]; snprintf(nm, sizeof nm, "Param%d", i);
            free(p->name); p->name = strdup(nm);
            int ty = 0, dec = 0, nl = 0; unsigned long sz = 0;
            if (!lbo_describe_param(st, i, &ty, &sz, &dec, &nl)) {
                lbo_colinfo ci; memset(&ci, 0, sizeof ci); ci.sqltype = ty; ci.size = sz;
                p->type = ado_of_sql(&ci); p->size = (int)sz; p->prec = (int)sz; p->scale = dec;
            } else p->type = adVarChar;
            p->dir = 1;
            alist_add(cmd->params, (vbobj)p);
        }
    }
    lbo_stmt_free(st);
    free(sql); free(src);
}

/* ---------------------------------------------------------------- Properties */
typedef struct aprop { vbobjh h; char *name; vbvar value; int type, attrs; } aprop;
static void prop_clear(vbobj o) { aprop *p = (aprop *)o; free(p->name); vb_vfree(&p->value); }
static vbvar prop_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    aprop *p = (aprop *)vb_oid(o);
    *found = 1;
    if (!strcmp(name, "value") || !strcmp(name, "")) {
        if (is_set(kind)) { vb_vset(&p->value, vb_v_letval(argv[argc - 1])); return vb_v_empty(); }
        return vb_vtmp(p->value);
    }
    if (!strcmp(name, "name")) return vstr(p->name);
    if (!strcmp(name, "type")) return vb_v_l(p->type);
    if (!strcmp(name, "attributes")) return vb_v_l(p->attrs);
    *found = 0;
    return vb_v_empty();
}
const vb_class vb_cls_ado_property = { "Property", sizeof(aprop), 0, 0, prop_clear, 0, 0, "value", 0, 0, 0, prop_member };
static void prop_add(alist *l, const char *name, vbvar v, int type)
{
    aprop *p = (aprop *)vb_onew(&vb_cls_ado_property);
    p->name = strdup(name); vb_vset(&p->value, v); p->type = type; p->attrs = 1;
    alist_add(l, (vbobj)p);
}
static void odbc_props(acn *c, alist *l)
{
    char b[512];
    prop_add(l, "Current Catalog", vstr(c->defdb), adBSTR);
    prop_add(l, "Data Source Name", vstr(lbo_info(c->odbc, LBI_DATA_SOURCE_NAME, b, sizeof b) ? "" : b), adBSTR);
    prop_add(l, "Read-Only Data Source", vb_v_bool(c->odbc_ro ? -1 : 0), adBoolean);
    prop_add(l, "DBMS Name", vstr(c->dbms), adBSTR);
    prop_add(l, "DBMS Version", vstr(lbo_info(c->odbc, LBI_DBMS_VER, b, sizeof b) ? "" : b), adBSTR);
    prop_add(l, "Driver Name", vstr(lbo_info(c->odbc, LBI_DRIVER_NAME, b, sizeof b) ? "" : b), adBSTR);
    prop_add(l, "Driver Version", vstr(lbo_info(c->odbc, LBI_DRIVER_VER, b, sizeof b) ? "" : b), adBSTR);
    prop_add(l, "Provider Name", vstr("MSDASQL.DLL"), adBSTR);
    prop_add(l, "Provider Friendly Name", vstr(SRC_ODBC), adBSTR);
    prop_add(l, "Server Name", vstr(lbo_info(c->odbc, LBI_SERVER_NAME, b, sizeof b) ? "" : b), adBSTR);
    prop_add(l, "User Name", vstr(lbo_info(c->odbc, LBI_USER_NAME, b, sizeof b) ? "" : b), adBSTR);
    prop_add(l, "SQL Support", vb_v_l(c->odbc_sql), adInteger);
    prop_add(l, "Extended Properties", vstr(c->cs), adBSTR);
}

static void odbc_rs_props(ars *r, alist *l)
{
    if (r->cursorloc == adUseClient) prop_add(l, "Update Criteria", vb_v_l(2), adInteger);   /* adCriteriaUpdCols */
}
static const char *prop_name(vbobj o) { return ((aprop *)vb_oid(o))->name; }

/* ---------------------------------------------------------------- Recordset.Open, Command.Execute */
static vbvar par_coerce(int type, vbvar v);

static int is_bare_name(const char *text)
{
    const char *p = text;
    while (isspace((unsigned char)*p)) p++;
    if (!*p || *p == '{' || *p == '(' || looks_sql(p)) return 0;
    for (; *p; p++) if (isspace((unsigned char)*p)) {
        while (isspace((unsigned char)*p)) p++;
        return !*p;
    }
    return 1;
}

static vbvar odbc_rs_open(ars *r, acn *c, acmd *cmd, vbvar src, int opts, vb_i32 argc, vbvar *argv)
{
    int req = !missing(argc, argv, 2) ? vb_v2l(argv[2]) : r->cursortype;
    char *text = cmd ? strdup(cmd->text ? cmd->text : "") : strdup(vb_stoutf8(vb_v2s(src)));
    if (cmd && cmd->type != adCmdUnknown && opts == adCmdUnknown) opts = cmd->type;
    int t = opts & ~(128 | 16 | 32 | 64);
    /* a bare name without adCmdTable: MSDASQL refuses it (VB6 oracle d_odbc2) */
    if ((t == adCmdUnknown || t == adCmdText) && is_bare_name(text)) {
        free(r->source); r->source = text;
        alist_clear(c->errs);
        raise(-2147217887, SRC_ODBC, "ODBC driver does not support the requested properties.");
        return vb_v_empty();
    }
    int nv = cmd ? cmd->params->n : 0;
    char *srcname;
    char *sql = odbc_sql(text, opts, nv, &srcname);
    free(text);
    vbvar *vals = calloc((size_t)(nv ? nv : 1), sizeof(vbvar));
    for (int i = 0; i < nv; i++) vb_vset(&vals[i], ((apar *)vb_oid(cmd->params->items[i]))->value);
    int ct = r->cursortype, lt = r->locktype, cl = r->cursorloc;
    int rows = 0;
    int e = odbc_run(c, sql, vals, nv, r, 0, &rows);
    if (e || !rows) {                           /* an error, or a statement without rows: closed */
        for (int i = 0; i < nv; i++) vb_vfree(&vals[i]);
        free(vals); free(sql);
        free(r->source); r->source = srcname;
        return vb_v_empty();
    }
    r->cursorloc = cl;
    r->cursortype = cl == adUseClient ? adOpenStatic : ct == adOpenForwardOnly ? adOpenForwardOnly : adOpenStatic;
    r->locktype = lt;
    if (cl != adUseClient && req == adOpenDynamic) r->ocount = -1;
    free(r->source); r->source = srcname;
    free(r->osql); r->osql = sql;
    for (int i = 0; i < r->novals; i++) vb_vfree(&r->ovals[i]);
    free(r->ovals); r->ovals = vals; r->novals = nv;
    r->state = 1;
    vb_vset(&r->filter, vb_v_l(0));
    free(r->sort); r->sort = strdup("");
    ev_opened(r);
    note_fill(r);
    return vb_v_empty();
}

static vbvar odbc_cmd_execute(acmd *c, acn *cn, vb_i32 argc, vbvar *argv)
{
    if (!missing(argc, argv, 1)) {              /* the values: the array given */
        vbvar a = argv[1];
        int n = 0; vbarr *arr = 0;
        if (a.vt & VT_ARRAY) { arr = vb_v_arrptr(a); n = arr->ub[0] - arr->lb[0] + 1; }
        else n = 1;
        for (int i = 0; i < n; i++) {
            vbvar x = arr ? *(vbvar *)vb_aelem1(arr, arr->lb[0] + i) : a;
            if (i >= c->params->n) {
                apar *p = (apar *)vb_onew(&vb_cls_ado_parameter);
                alist_add(c->params, (vbobj)p);
            }
            apar *p = (apar *)vb_oid(c->params->items[i]);
            vb_vset(&p->value, par_coerce(p->type, x));
        }
    }
    int opts = along(argc, argv, 2, -1);
    if (opts == -1) opts = c->type;
    int nv = c->params->n;
    char *src;
    char *sql = odbc_sql(c->text ? c->text : "", opts, nv, &src);
    vbvar *vals = calloc((size_t)(nv ? nv : 1), sizeof(vbvar));
    for (int i = 0; i < nv; i++) vb_vset(&vals[i], ((apar *)vb_oid(c->params->items[i]))->value);
    ars *r = (ars *)rs_new_for(cn);
    long n = 0; int rows = 0;
    int e = odbc_run(cn, sql, vals, nv, r, &n, &rows);
    if (e) {
        for (int i = 0; i < nv; i++) vb_vfree(&vals[i]);
        free(vals); free(sql); free(src); vb_orelease((vbobj)r);
        return vb_v_empty();
    }
    if (argc > 0 && !vb_ismissing(argv[0])) vb_vset(&argv[0], vb_v_l(rows ? -1 : (vb_i32)n));
    if (rows) {
        r->cursortype = cn->cursorloc == adUseClient ? adOpenStatic : adOpenForwardOnly;
        r->locktype = adLockReadOnly;
        r->cursorloc = cn->cursorloc;
        r->state = 1;
        free(r->source); r->source = src; src = 0;
        r->osql = sql; sql = 0;
        r->ovals = vals; r->novals = nv; vals = 0; nv = 0;
    }
    for (int i = 0; i < nv; i++) vb_vfree(&vals[i]);
    free(vals); free(sql); free(src);
    vbvar v = retobj((vbobj)r);
    vb_orelease((vbobj)r);
    return v;
}
