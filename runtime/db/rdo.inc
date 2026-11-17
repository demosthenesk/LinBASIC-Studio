/* rdo.inc - Microsoft Remote Data Objects 2.0 (MSRDO20.DLL), written again over the ODBC provider of ado.c
 * (ado_odbc.inc): rdoEngine, rdoEnvironment, rdoConnection, rdoQuery, rdoResultset, rdoColumn, rdoParameter,
 * rdoTable, rdoError and their collections, with the members, the defaults, the error numbers and texts (the
 * string table of MSRDO20.DLL: 40000 + n) and the rdoErrors of the original (VB6 oracle f_rdo1, VB-TESTS/DEV-7).
 * Included by ado.c.
 *
 * An rdoConnection keeps an ADO Connection (MSDASQL) and an rdoResultset an ADO Recordset with the rows of the
 * statement (a recordset of its own in memory, its changes going to the server by the key of the row): the
 * moves, the edits, the bookmarks of ado.c, with the rules of RDO on top: Edit before a change, AddNew / Update
 * leaving the current row where it was, a change dropped by a move, RowCount known after MoveLast, BOF / EOF
 * errors 40025 / 40024 ...
 * The cursor drivers give the cursors that MSRDO20 gives with the SQLite ODBC driver (rdUseIfNeeded, rdUseServer:
 * a keyset of the driver; rdUseOdbc: the ODBC cursor library; rdUseClientBatch: as asked; rdUseNone: forward-only
 * read-only only). What the oracle showed of the defects of that driver and of the cursor library (an Update that
 * fails with "out of memory", the names of the columns "??" with rdUseClientBatch, 40054 for the properties of a
 * column) is not done: the expected output of f_rdo1 is the one of VB6 without them (the output of VB6 as it is:
 * VB-TESTS/DEV-7/f_rdo1/expected/f_rdo1-raw.txt).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */

extern const vb_class vb_cls_rdo_engine, vb_cls_rdo_environment, vb_cls_rdo_connection, vb_cls_rdo_query,
    vb_cls_rdo_resultset, vb_cls_rdo_column, vb_cls_rdo_parameter, vb_cls_rdo_table, vb_cls_rdo_error, vb_cls_rdo_coll;

#define RDO_DLL "MSRDO20.DLL"
#define RDO_SCN "MicrosoftRDO.RdoConnection"
#define RDO_SQRY "MicrosoftRDO.RdoQuery"

/* ---------------------------------------------------------------- the texts of the errors (MSRDO20.DLL) */
static const struct { int n; const char *t; } rdo_texts[] = {
    { 40000, "An error occurred configuring the DSN.  Please check the parameters and try again." },
    { 40001, "SQL returned No Data Found." },
    { 40002, "An internal ODBC error was encountered." },
    { 40003, "An invalid value for the Cursor Driver was passed.\nPlease use one of the following values:\nrdUseIfNeeded, rdUseODBC, rdUseServer\nrdUseClientBatch, rdUseNone." },
    { 40004, "An invalid ODBC handle was encountered." },
    { 40005, "Invalid connection string." },
    { 40006, "An unexpected error occurred." },
    { 40007, "Out of memory." },
    { 40008, "Invalid operation for forward-only cursor." },
    { 40009, "No current row." },
    { 40010, "Invalid row for AddNew or Edit." },
    { 40011, "Object is invalid or not set." },
    { 40012, "Invalid seek flag." },
    { 40013, "Partial equality requires string column." },
    { 40014, "Incompatible data types for compare." },
    { 40015, "Can't create prepared statement for invalid database connection." },
    { 40017, "Can't execute unprepared rdoPreparedStatement." },
    { 40018, "Can't execute empty rdoPreparedStatement." },
    { 40019, "An invalid value for the Concurrency option was passed.  Please use one of the following values:\nrdConcurReadOnly, rdConcurLock, rdConcurRowver\nrdConcurValues, rdConcurBatch." },
    { 40020, "Can't open result set for unnamed table." },
    { 40021, "Object Collection: This collection doesn't support location by text tag." },
    { 40022, "Resultset is empty." },
    { 40023, "Invalid state for Move." },
    { 40024, "Already beyond the end of the result set." },
    { 40025, "BOF already set." },
    { 40026, "Invalid result set state for Update." },
    { 40027, "Invalid Bookmark." },
    { 40028, "Invalid bookmark argument supplied to Move." },
    { 40029, "Can't move relative to current row when EOF/BOF are set." },
    { 40030, "Already at BOF." },
    { 40031, "Already at EOF." },
    { 40033, "An invalid value for the prompt option was passed.  Please use one of the following values:\nrdDriverPrompt\nrdDriverNoPrompt\nrdDriverComplete\n\nrdDriverCompleteRequired." },
    { 40034, "An invalid value for the cursor type parameter was passed.  Please use one of the following values:\nrdOpenKeyset\nrdOpenDynamic\nrdOpenStatic\n\nrdOpenForwardOnly." },
    { 40035, "Field not bound correctly." },
    { 40036, "Unbound Field - Use GetChunk." },
    { 40037, "Can't assign value to unbound field." },
    { 40038, "Can't assign value to non-updatable field." },
    { 40039, "Can't assign value to field unless in Edit mode." },
    { 40040, "Incorrect type for parameter." },
    { 40041, "Object Collection: Couldn't find item indicated by text." },
    { 40042, "Can't assign or read value to or from unbound parameter." },
    { 40043, "Can't assign value to output only parameter." },
    { 40044, "Incorrect type found when assigning rdoParameter Value property." },
    { 40045, "This operation is not allowed at this time." },
    { 40046, "The object has already been closed." },
    { 40047, "You must specify a valid name for the environment." },
    { 40048, "This environment name already exists in the collection." },
    { 40049, "Object Collection: Illegal modification - collection is read-only." },
    { 40052, "You must use AppendChunk to set data in a TEXT or IMAGE(BLOB) field." },
    { 40054, "An invalid parameter was passed." },
    { 40055, "Invalid Operation." },
    { 40056, "You attempted to reposition to a deleted row." },
    { 40057, "An attempt was made to issue a SELECT statement using the Execute method." },
    { 40058, "The Resultset is read-only." },
    { 40060, "This column does not have the ChunkRequired flag set.  Use the Value property to retrieve it's contents." },
    { 40069, "General Client Cursor error." },
    { 40070, "This property or method is not supported when using the currently selected cursor library." },
    { 40071, "The rdoConnection object is not connected to a data source." },
    { 40072, "The rdoConnection object is already connected to a data source." },
    { 40074, "The rdoQuery or rdoResultset has no active connection." },
    { 40076, "This property is read-only, now." },
    { 40079, "Can't find table to update." },
    { 40080, "Cursor type should be: rdOpenForwardOnly.\n LockType should be: rdConcurReadOnly.\n RowsetSize should be: 1." },
    { 40081, "Only the rdUseClientBatch cursor driver can support\nrdConcurBatch lock type." },
    { 40082, "This function is not supported." },
    { 40085, "Invalid Option flags." },
    { 40086, "Empty SQL string not allowed." },
    { 40088, "No open cursor or cursor closed." },
    { 40503, "A control bound to a data control using this Resultset\ncanceled the operation or was unable to synchronize itself with the Resultset." },
};
static const char *rdo_text(int n)
{
    for (size_t i = 0; i < sizeof rdo_texts / sizeof rdo_texts[0]; i++) if (rdo_texts[i].n == n) return rdo_texts[i].t;
    return "An unexpected error occurred.";
}

/* ---------------------------------------------------------------- collections */
enum { RL_ENVS, RL_ERRS, RL_CNS, RL_QRYS, RL_RSS, RL_PARS, RL_COLS, RL_TABS };
static const char *const rl_src[] = { "MSRDO.rdoEnvironments", "MSRDO.rdoErrors", "MSRDO.rdoConnections", "MSRDO.rdoQuerys",
                                      "MSRDO.rdoResultsets", "MSRDO.rdoParameters", "MSRDO.rdoColumns", "MSRDO.rdoTables" };
/* owner: the rdoResultset (counted) of its rdoColumns, whose items are made when asked for */
typedef struct rlist { vbobjh h; int kind; vbobj *items; int n, cap; vbobj owner; void *back; } rlist;   /* back: the rcn of rdoTables (not counted) */

typedef struct rerr { vbobjh h; vb_i32 number; int retcode; char *desc, *src, *state; } rerr;
typedef struct reng {
    vbobjh h;
    rlist *envs, *errs;
    int cursor, login, errthr, locale;
    char *user, *pwd;
} reng;
typedef struct renv { vbobjh h; char *name, *user, *pwd; int cursor, login, closed; rlist *cns; } renv;
typedef struct rcn {
    vbobjh h;
    renv *env;                          /* not counted (the engine keeps it) */
    acn *a;                             /* the ADO connection (counted) while connected */
    char *name, *connect, *logmsg;
    int cursor, login, qtimeout, ro, connected, closed, inenv;
    long affected;
    rlist *qrys, *rss, *tabs;
    int tabs_loaded;
} rcn;
typedef struct rqry {
    vbobjh h;
    vbobj cn;                           /* rcn (counted) or 0 */
    char *name, *sql;
    rlist *pars;
    int type, prepared, maxrows, rowset, keyset, cursor, lock, timeout, bindthr, errthr, closed;
    long affected;
} rqry;
typedef struct rrs {
    vbobjh h;
    vbobj cn;                           /* rcn (counted) */
    ars *a;                             /* the rows (counted) */
    char *name;
    int type, lock, drv, upd, bm, closed, known, rcdel, lastmod, addfrom;
} rrs;
typedef struct rdocol {                 /* of a resultset (rs, col), or of a table (its own description) */
    vbobjh h;
    vbobj rs;
    int col;
    char *name, *stab, *scol;
    int type, size, attrs, req, zl, ord;
} rdcol;
typedef struct rpar { vbobjh h; char *name; int dir, type, size; vbvar value; } rpar;
typedef struct rtab { vbobjh h; vbobj cn; char *name, *type; rlist *cols; int ncols; } rtab;

#define RRS(o) ((rrs *)vb_oid(o))
#define RCN(o) ((rcn *)vb_oid(o))

static rlist *rlist_new(int kind)
{
    rlist *l = (rlist *)vb_onew(&vb_cls_rdo_coll);
    vb_oaddref((vbobj)l);
    l->kind = kind;
    return l;
}
static void rlist_add(rlist *l, vbobj o, int front)
{
    if (l->n == l->cap) { l->cap = l->cap ? l->cap * 2 : 8; l->items = realloc(l->items, sizeof(vbobj) * (size_t)l->cap); }
    vb_oaddref(o);
    if (front) { memmove(&l->items[1], &l->items[0], sizeof(vbobj) * (size_t)l->n); l->items[0] = o; }
    else l->items[l->n] = o;
    l->n++;
}
static void rlist_remove(rlist *l, vbobj o)
{
    for (int i = 0; i < l->n; i++)
        if (l->items[i] == o) {
            memmove(&l->items[i], &l->items[i + 1], sizeof(vbobj) * (size_t)(l->n - i - 1));
            l->n--;
            vb_orelease(o);
            return;
        }
}
static void rlist_clear(rlist *l) { while (l->n) { vbobj o = l->items[--l->n]; vb_orelease(o); } }
static void rlist_free(vbobj o) { rlist *l = (rlist *)o; rlist_clear(l); free(l->items); if (l->owner) vb_orelease(l->owner); }

/* ---------------------------------------------------------------- the engine, the errors */
static reng *g_rdo;
static renv *env_new(const char *name, const char *user, const char *pwd);
static reng *rdo_eng(void)
{
    if (!g_rdo) {
        g_rdo = (reng *)vb_onew(&vb_cls_rdo_engine);
        vb_oaddref((vbobj)g_rdo);               /* for the whole run */
        g_rdo->envs = rlist_new(RL_ENVS);
        g_rdo->errs = rlist_new(RL_ERRS);
        g_rdo->login = 15; g_rdo->errthr = -1; g_rdo->locale = 1;
        g_rdo->user = strdup(""); g_rdo->pwd = strdup("");
        renv *e = env_new("Default_Environment", "", "");
        rlist_add(g_rdo->envs, (vbobj)e, 0);
        vb_orelease((vbobj)e);
    }
    return g_rdo;
}
vbobj vb_rdo_engine(void) { return (vbobj)rdo_eng(); }

static void rdo_errs_clear(void) { rlist_clear(rdo_eng()->errs); }
static void rdo_err_add(vb_i32 number, const char *desc, const char *src, const char *state, int ret, int front)
{
    rerr *e = (rerr *)vb_onew(&vb_cls_rdo_error);
    e->number = number; e->retcode = ret;
    e->desc = strdup(desc ? desc : ""); e->src = strdup(src ? src : ""); e->state = strdup(state ? state : "");
    rlist_add(rdo_eng()->errs, (vbobj)e, front);
}
/* an error of RDO: first in rdoErrors (objsrc: its Source there, "MSRDO.rdoResultset"), then raised (errsrc:
 * Err.Source, MSRDO20.DLL when 0) */
static void rdo_fail(int n, const char *objsrc, const char *errsrc)
{
    const char *t = rdo_text(n);
    rdo_err_add(n, t, objsrc, "", 0, 1);
    vb_raise_full(n, vb_sfromutf8(errsrc ? errsrc : RDO_DLL), vb_sfromutf8(t));
}
/* an object closed (40011): not in rdoErrors */
static void rdo_invalid(const char *errsrc)
{
    vb_raise_full(40011, vb_sfromutf8(errsrc ? errsrc : RDO_DLL), vb_sfromutf8(rdo_text(40011)));
}
/* the diagnostics of ODBC (Errors of the ADO connection) as rdoErrors, 40002 raised with the first */
static void rdo_odbc_raise(alist *errs, const char *errsrc)
{
    rdo_errs_clear();
    char first[1200] = "";
    for (int i = 0; errs && i < errs->n; i++) {
        aerr *x = (aerr *)vb_oid(errs->items[i]);
        char d[1200];
        snprintf(d, sizeof d, "%s: %s", x->state ? x->state : "S1000", x->desc ? x->desc : "");
        if (!i) snprintf(first, sizeof first, "%s", d);
        rdo_err_add(x->native, d, "ODBC", x->state, -1, 0);
    }
    vb_raise_full(40002, vb_sfromutf8(errsrc ? errsrc : RDO_DLL), vb_sfromutf8(first[0] ? first : rdo_text(40002)));
}
/* an ODBC error of one state, as the driver manager or the cursor library gives it */
static void rdo_odbc_state(const char *state, vb_i32 native, const char *msg, const char *errsrc)
{
    rdo_errs_clear();
    char d[1200];
    snprintf(d, sizeof d, "%s: %s", state, msg);
    rdo_err_add(native, d, "ODBC", state, -1, 0);
    vb_raise_full(40002, vb_sfromutf8(errsrc ? errsrc : RDO_DLL), vb_sfromutf8(d));
}
/* an error pending from ado.c, raised again as RDO's; 1 if there was one */
static int rdo_failed(acn *a, const char *errsrc)
{
    if (!vb_errpend) return 0;
    if (a && a->errs && a->errs->n) {
        vb_err_clear_pending();
        rdo_odbc_raise(a->errs, errsrc);
        return 1;
    }
    int n = vb_err_number_pending();
    char *msg = strdup(vb_err_text_pending());
    vb_err_clear_pending();
    if (n == 3021) rdo_fail(40009, "MSRDO.rdoResultset", errsrc);
    else if (n == 13 || n == 3421) vb_raise(13);
    else rdo_odbc_state("S1000", 0, msg, errsrc);
    free(msg);
    return 1;
}

/* ---------------------------------------------------------------- values */
static char *rdo_sdup(vbvar v) { return strdup(vb_stoutf8(vb_v2s(v))); }
static vbvar rdo_value(vbvar v)
{
    if (v.vt == VT_UI1) return vb_v_i((vb_i16)v.u.b);
    if (v.vt == VT_DECIMAL) return vb_v_d(vb_dec2d(v));     /* NUMERIC / DECIMAL: a Double in RDO */
    return v;
}
/* the SQL type (rdTypeXXX) of an ADO type of the provider */
static int rdo_sqltype(int t)
{
    switch (t) {
    case adChar: case adWChar: return 1;
    case adNumeric: case adDecimal: return 2;
    case adInteger: case adUnsignedInt: return 4;
    case adSmallInt: case adUnsignedSmallInt: return 5;
    case adSingle: return 7;
    case adDouble: case adCurrency: return 8;
    case adDBDate: return 9;
    case adDBTime: return 10;
    case adDBTimeStamp: case adDate: return 11;
    case adVarChar: case adVarWChar: case adBSTR: return 12;
    case adLongVarChar: case adLongVarWChar: return -1;
    case adBinary: return -2;
    case adVarBinary: return -3;
    case adLongVarBinary: return -4;
    case adBigInt: case 21: return -5;
    case adTinyInt: case adUnsignedTinyInt: return -6;
    case adBoolean: return -7;
    case adGUID: return -11;
    }
    return 12;
}
/* ODBC's text types as the SQLite driver of Windows gives them (wide) */
static int rdo_wide_sql(acn *a, int t)
{
    if (a && a->dbms && !strncasecmp(a->dbms, "SQLite", 6)) {  /* (in Wine too: the driver of Linux under it) */
        if (t == LBS_CHAR) return LBS_WCHAR;
        if (t == LBS_VARCHAR) return LBS_WVARCHAR;
        if (t == LBS_LONGVARCHAR) return LBS_WLONGVARCHAR;
    }
    return t;
}
static int rdo_size_of(int sqltype, int size)
{
    switch (sqltype) {
    case 4: return 10;
    case 5: return 5;
    case -6: return 3;
    case -5: return 19;
    case 7: return 7;
    case 6: case 8: return 15;
    case 9: return 10;
    case 10: return 8;
    case 11: return 19;
    case -7: return 1;
    case -11: return 36;
    }
    return size;
}

/* ---------------------------------------------------------------- rdoEnvironment */
static renv *env_new(const char *name, const char *user, const char *pwd)
{
    renv *e = (renv *)vb_onew(&vb_cls_rdo_environment);
    vb_oaddref((vbobj)e);
    e->name = strdup(name); e->user = strdup(user); e->pwd = strdup(pwd);
    e->cursor = g_rdo ? g_rdo->cursor : 0;
    e->login = g_rdo ? g_rdo->login : 15;
    e->cns = rlist_new(RL_CNS);
    return e;
}
static void env_clear(vbobj o)
{
    renv *e = (renv *)o;
    free(e->name); free(e->user); free(e->pwd);
    if (e->cns) vb_orelease((vbobj)e->cns);
}

static void cn_disconnect(rcn *k);

/* the ODBC connection string of an rdoConnection: DSN, the connect string, the user of the environment */
static char *rdo_cs(const char *dsn, const char *connect, renv *env)
{
    jbuf s = { 0 };
    char *c = strdup(connect ? connect : "");
    if (!strncasecmp(c, "ODBC;", 5)) memmove(c, c + 5, strlen(c + 5) + 1);
    if (dsn && *dsn && !has_key(c, "DSN") && !has_key(c, "Driver")) { jb_s(&s, "DSN="); jb_s(&s, dsn); }
    if (*c) { if (s.n) jb_s(&s, ";"); jb_s(&s, c); }
    if (env && env->user && *env->user && !has_key(c, "UID")) { jb_s(&s, ";UID="); jb_s(&s, env->user); }
    if (env && env->pwd && *env->pwd && !has_key(c, "PWD")) { jb_s(&s, ";PWD="); jb_s(&s, env->pwd); }
    free(c);
    return s.p ? s.p : strdup("");
}

/* connect k (its Connect, the DSN given); 0, or -1 with the error raised */
static int cn_connect(rcn *k, const char *dsn, int ro)
{
    rdo_errs_clear();
    char *cs = rdo_cs(dsn, k->connect, k->env);
    acn *a = (acn *)vb_onew(&vb_cls_ado_connection);
    vb_oaddref((vbobj)a);
    free(a->cs); a->cs = cs;
    int r = odbc_open(a);
    if (!r) {                                   /* not for ODBC: no DSN, no Driver */
        vb_orelease((vbobj)a);
        rdo_odbc_state("IM002", 0, "[Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified", RDO_SCN);
        return -1;
    }
    if (!a->state) {
        if (!vb_errpend) raise(-2147467259, SRC_ODBC, "Unspecified error");
        vb_err_clear_pending();
        rdo_odbc_raise(a->errs, RDO_SCN);
        vb_orelease((vbobj)a);
        return -1;
    }
    a->cursorloc = adUseClient;
    k->a = a;
    k->connected = 1;
    k->ro = ro;
    free(k->name);
    if (dsn && *dsn) k->name = strdup(dsn);
    else {
        char *db = cs_get(k->connect ? k->connect : "", "Database");
        if (!db) db = cs_get(k->connect ? k->connect : "", "DBQ");
        k->name = db ? db : strdup("");
    }
    return 0;
}

static vbvar env_open_cn(renv *e, vb_i32 argc, vbvar *argv)
{
    char *dsn = missing(argc, argv, 0) ? strdup("") : rdo_sdup(argv[0]);
    int prompt = along(argc, argv, 1, 0);
    int ro = missing(argc, argv, 2) ? 0 : vb_v2bool(argv[2]) != 0;
    if (prompt < 0 || prompt > 3) { free(dsn); rdo_fail(40033, "MSRDO.rdoEnvironment", 0); return vb_v_empty(); }
    rcn *k = (rcn *)vb_onew(&vb_cls_rdo_connection);
    k->env = e;
    k->cursor = e->cursor; k->login = e->login;
    free(k->connect); k->connect = missing(argc, argv, 3) ? strdup("") : rdo_sdup(argv[3]);
    if (cn_connect(k, dsn, ro)) { free(dsn); return vb_v_empty(); }
    free(dsn);
    k->inenv = 1;
    rlist_add(e->cns, (vbobj)k, 0);
    return retobj((vbobj)k);
}

static vbvar rdo_trans_all(renv *e, int what)
{
    for (int i = 0; i < e->cns->n; i++) {
        rcn *k = RCN(e->cns->items[i]);
        if (!k->connected || !k->a) continue;
        if (what == 0 && !k->a->trans) odbc_trans(k->a, 0);
        else if (what && k->a->trans) odbc_trans(k->a, what);
        if (rdo_failed(k->a, 0)) return vb_v_empty();
    }
    return vb_v_empty();
}

static vbvar rlist_ret(rlist *l, vb_i32 kind, vb_i32 argc, vbvar *argv);

static vbvar env_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    renv *e = (renv *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (e->closed) { rdo_invalid(0); return vb_v_empty(); }
    if (!strcmp(name, "name")) return vstr(e->name);
    if (!strcmp(name, "username")) { if (set) { free(e->user); e->user = rdo_sdup(argv[argc - 1]); return vb_v_empty(); } return vstr(e->user); }
    if (!strcmp(name, "password")) { if (set) { free(e->pwd); e->pwd = rdo_sdup(argv[argc - 1]); return vb_v_empty(); } return vstr(e->pwd); }
    if (!strcmp(name, "cursordriver")) {
        if (set) {
            int d = vb_v2l(argv[argc - 1]);
            if (d < 0 || d > 4) { rdo_fail(40003, "MSRDO.rdoEnvironment", 0); return vb_v_empty(); }
            e->cursor = d; return vb_v_empty();
        }
        return vb_v_l(e->cursor);
    }
    if (!strcmp(name, "logintimeout")) { if (set) { e->login = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(e->login); }
    if (!strcmp(name, "rdoconnections") || !strcmp(name, "")) return rlist_ret(e->cns, kind, argc, argv);
    if (!strcmp(name, "openconnection")) return env_open_cn(e, argc, argv);
    if (!strcmp(name, "begintrans")) return rdo_trans_all(e, 0);
    if (!strcmp(name, "committrans")) return rdo_trans_all(e, 1);
    if (!strcmp(name, "rollbacktrans")) return rdo_trans_all(e, 2);
    if (!strcmp(name, "henv")) return vb_v_l(0);
    if (!strcmp(name, "close")) {
        while (e->cns->n) {
            rcn *k = RCN(e->cns->items[e->cns->n - 1]);
            vb_oaddref((vbobj)k);
            cn_disconnect(k);
            k->closed = 1;
            rlist_remove(e->cns, (vbobj)k);
            vb_orelease((vbobj)k);
        }
        reng *g = rdo_eng();
        if (g->envs->n && g->envs->items[0] != o) {     /* the default one stays */
            e->closed = 1;
            rlist_remove(g->envs, vb_oid(o));
        }
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- rdoEngine */
static void eng_clear(vbobj o)
{
    reng *g = (reng *)o;
    if (g->envs) vb_orelease((vbobj)g->envs);
    if (g->errs) vb_orelease((vbobj)g->errs);
    free(g->user); free(g->pwd);
}
static vbvar eng_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    reng *g = (reng *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (!strcmp(name, "rdoversion")) return vstr("6.01.9839");
    if (!strcmp(name, "rdodefaultcursordriver")) {
        if (set) {
            int d = vb_v2l(argv[argc - 1]);
            if (d < 0 || d > 4) { rdo_fail(40003, "MSRDO.rdoEngine", 0); return vb_v_empty(); }
            g->cursor = d; return vb_v_empty();
        }
        return vb_v_l(g->cursor);
    }
    if (!strcmp(name, "rdodefaultlogintimeout")) { if (set) { g->login = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(g->login); }
    if (!strcmp(name, "rdodefaulterrorthreshold")) { if (set) { g->errthr = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(g->errthr); }
    if (!strcmp(name, "rdolocaleid")) { if (set) { g->locale = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(g->locale); }
    if (!strcmp(name, "rdodefaultuser")) { if (set) { free(g->user); g->user = rdo_sdup(argv[argc - 1]); return vb_v_empty(); } return vstr(g->user); }
    if (!strcmp(name, "rdodefaultpassword")) { if (set) { free(g->pwd); g->pwd = rdo_sdup(argv[argc - 1]); return vb_v_empty(); } return vstr(g->pwd); }
    if (!strcmp(name, "rdoenvironments") || !strcmp(name, "")) return rlist_ret(g->envs, kind, argc, argv);
    if (!strcmp(name, "rdoerrors")) return rlist_ret(g->errs, kind, argc, argv);
    if (!strcmp(name, "rdocreateenvironment")) {
        char *nm = missing(argc, argv, 0) ? strdup("") : rdo_sdup(argv[0]);
        if (!*nm) { free(nm); rdo_fail(40047, "MSRDO.rdoEngine", 0); return vb_v_empty(); }
        for (int i = 0; i < g->envs->n; i++)
            if (!strcasecmp(((renv *)vb_oid(g->envs->items[i]))->name, nm)) { free(nm); rdo_fail(40048, "MSRDO.rdoEngine", 0); return vb_v_empty(); }
        char *u = missing(argc, argv, 1) ? strdup(g->user) : rdo_sdup(argv[1]);
        char *p = missing(argc, argv, 2) ? strdup(g->pwd) : rdo_sdup(argv[2]);
        renv *e = env_new(nm, u, p);
        free(nm); free(u); free(p);
        rlist_add(g->envs, (vbobj)e, 0);
        vbvar v = retobj((vbobj)e);
        vb_orelease((vbobj)e);
        return v;
    }
    if (!strcmp(name, "rdoregisterdatasource")) return vb_v_empty();
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- rdoConnection */
static void rcn_init(vbobj o)
{
    rcn *k = (rcn *)o;
    reng *g = rdo_eng();
    renv *e = (renv *)vb_oid(g->envs->items[0]);
    k->env = e;
    k->cursor = e->cursor; k->login = e->login; k->qtimeout = 30;
    k->name = strdup(""); k->connect = strdup(""); k->logmsg = strdup("");
    k->qrys = rlist_new(RL_QRYS); k->rss = rlist_new(RL_RSS); k->tabs = rlist_new(RL_TABS);
    k->tabs->back = k;
}
static void rs_invalidate(rrs *s);
static void cn_disconnect(rcn *k)
{
    while (k->rss->n) {
        rrs *s = RRS(k->rss->items[k->rss->n - 1]);
        vb_oaddref((vbobj)s);
        rs_invalidate(s);
        rlist_remove(k->rss, (vbobj)s);
        vb_orelease((vbobj)s);
    }
    for (int i = 0; i < k->qrys->n; i++) ((rqry *)vb_oid(k->qrys->items[i]))->closed = 1;
    rlist_clear(k->qrys);
    rlist_clear(k->tabs); k->tabs_loaded = 0;
    if (k->a) { cn_close(k->a); vb_orelease((vbobj)k->a); k->a = 0; }
    k->connected = 0;
}
static void rcn_clear(vbobj o)
{
    rcn *k = (rcn *)o;
    cn_disconnect(k);
    free(k->name); free(k->connect); free(k->logmsg);
    vb_orelease((vbobj)k->qrys); vb_orelease((vbobj)k->rss); vb_orelease((vbobj)k->tabs);
}

static vbvar rs_open_sql(rcn *k, const char *sql, vbvar *vals, int nv, int type, int lock, int maxrows, int drv);

/* rdoConnection.Execute, rdoQuery.Execute: the rows affected; -1 with the error raised */
static long rdo_execute(rcn *k, const char *sql, vbvar *vals, int nv)
{
    if (!sql || !*sql) { rdo_fail(40086, "MSRDO.rdoQuery", RDO_SQRY); return -1; }
    long n = 0; int rows = 0;
    int e = odbc_run(k->a, sql, vals, nv, 0, &n, &rows);
    if (e) { rdo_failed(k->a, RDO_SQRY); return -1; }
    if (rows) { rdo_fail(40057, "MSRDO.rdoQuery", RDO_SQRY); return -1; }
    return n < 0 ? 0 : n;
}

static void tabs_load(rcn *k);
static vbvar qry_new_on(rcn *k, const char *name, const char *sql);

static vbvar rcn_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    rcn *k = (rcn *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (k->closed) { rdo_invalid(RDO_SCN); return vb_v_empty(); }
    if (!strcmp(name, "name")) return vstr(k->name);
    if (!strcmp(name, "connect")) {
        if (set) {
            if (k->connected) { rdo_fail(40072, "MSRDO.rdoConnection", 0); return vb_v_empty(); }
            free(k->connect); k->connect = rdo_sdup(argv[argc - 1]); return vb_v_empty();
        }
        return vstr(k->connect);
    }
    if (!strcmp(name, "cursordriver")) {
        if (set) {
            int d = vb_v2l(argv[argc - 1]);
            if (d < 0 || d > 4) { rdo_fail(40003, "MSRDO.rdoConnection", 0); return vb_v_empty(); }
            k->cursor = d; return vb_v_empty();
        }
        return vb_v_l(k->cursor);
    }
    if (!strcmp(name, "logintimeout")) { if (set) { k->login = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(k->login); }
    if (!strcmp(name, "querytimeout")) { if (set) { k->qtimeout = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(k->qtimeout); }
    if (!strcmp(name, "logmessages")) { if (set) { free(k->logmsg); k->logmsg = rdo_sdup(argv[argc - 1]); return vb_v_empty(); } return vstr(k->logmsg); }
    if (!strcmp(name, "stillexecuting") || !strcmp(name, "stillconnecting")) return vb_v_bool(0);
    if (!strcmp(name, "asynccheckinterval")) return vb_v_l(1000);
    if (!strcmp(name, "establishconnection")) {
        if (k->connected) { rdo_fail(40072, "MSRDO.rdoConnection", RDO_SCN); return vb_v_empty(); }
        int prompt = along(argc, argv, 0, 0);
        if (prompt < 0 || prompt > 3) { rdo_fail(40033, "MSRDO.rdoConnection", RDO_SCN); return vb_v_empty(); }
        int ro = missing(argc, argv, 1) ? 0 : vb_v2bool(argv[1]) != 0;
        cn_connect(k, "", ro);
        return vb_v_empty();
    }
    if (!strcmp(name, "rdoqueries") || !strcmp(name, "rdopreparedstatements")) return rlist_ret(k->qrys, kind, argc, argv);
    if (!strcmp(name, "rdoresultsets")) return rlist_ret(k->rss, kind, argc, argv);
    if (!strcmp(name, "close")) {
        cn_disconnect(k);
        if (k->inenv) { k->closed = 1; k->inenv = 0; rlist_remove(k->env->cns, vb_oid(o)); }
        return vb_v_empty();
    }
    if (!strcmp(name, "cancel")) return vb_v_empty();
    /* the members of a connection that is connected */
    if (!k->connected) { rdo_fail(40071, "MSRDO.rdoConnection", RDO_SCN); return vb_v_empty(); }
    if (!strcmp(name, "updatable")) return vb_v_bool(k->ro ? 0 : -1);
    if (!strcmp(name, "transactions")) return vb_v_bool(-1);
    if (!strcmp(name, "version")) { char b[256]; return vstr(lbo_info(k->a->odbc, LBI_DRIVER_VER, b, sizeof b) ? "" : b); }
    if (!strcmp(name, "rowsaffected")) return vb_v_l((vb_i32)k->affected);
    if (!strcmp(name, "hdbc")) return vb_v_l(0);
    if (!strcmp(name, "lastqueryresults")) return retobj(0);
    if (!strcmp(name, "execute")) {
        if (missing(argc, argv, 0)) { vb_raise(450); return vb_v_empty(); }
        char *sql = rdo_sdup(argv[0]);
        long n = rdo_execute(k, sql, 0, 0);
        free(sql);
        if (n >= 0) k->affected = n;
        return vb_v_empty();
    }
    if (!strcmp(name, "openresultset")) {
        if (missing(argc, argv, 0)) { vb_raise(450); return vb_v_empty(); }
        char *sql = rdo_sdup(argv[0]);
        /* a name of a table or of a query of the connection */
        const char *p = sql; while (isspace((unsigned char)*p)) p++;
        for (int i = 0; i < k->qrys->n && *p; i++) {
            rqry *q = (rqry *)vb_oid(k->qrys->items[i]);
            if (q->name && !strcasecmp(q->name, p)) { free(sql); sql = strdup(q->sql ? q->sql : ""); break; }
        }
        if (is_bare_name(sql)) { jbuf b = { 0 }; jb_s(&b, "SELECT * FROM "); jb_s(&b, sql); free(sql); sql = b.p; }
        vbvar v = rs_open_sql(k, sql, 0, 0, along(argc, argv, 1, 0), along(argc, argv, 2, 1), 0, k->cursor);
        free(sql);
        return v;
    }
    if (!strcmp(name, "createquery") || !strcmp(name, "createpreparedstatement")) {
        char *nm = missing(argc, argv, 0) ? strdup("") : rdo_sdup(argv[0]);
        char *sql = missing(argc, argv, 1) ? strdup("") : rdo_sdup(argv[1]);
        vbvar v = qry_new_on(k, nm, sql);
        free(nm); free(sql);
        return v;
    }
    if (!strcmp(name, "rdotables")) { if (!k->tabs_loaded) tabs_load(k); return rlist_ret(k->tabs, kind, argc, argv); }
    if (!strcmp(name, "begintrans")) {
        rdo_errs_clear();
        if (!k->a->trans) { odbc_trans(k->a, 0); rdo_failed(k->a, 0); }
        return vb_v_empty();
    }
    if (!strcmp(name, "committrans") || !strcmp(name, "rollbacktrans")) {
        rdo_errs_clear();
        if (k->a->trans) {                      /* without BeginTrans: nothing (VB6 oracle f_rdo1) */
            for (int i = 0; i < k->rss->n; i++) {               /* the changes pending of its resultsets */
                rrs *s = RRS(k->rss->items[i]);
                drs *v = s->a ? VW(s->a) : 0;
                if (v && v->editmode) { int f; rs_member(s->a->view, "cancelupdate", VBM_METHOD, 0, 0, &f); if (vb_errpend) vb_err_clear_pending(); }
            }
            odbc_trans(k->a, name[0] == 'c' ? 1 : 2);
            rdo_failed(k->a, 0);
        }
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- rdoResultset */
static void rrs_clear(vbobj o)
{
    rrs *s = (rrs *)o;
    if (s->a) { if (s->a->state) rs_close_m(s->a); vb_orelease((vbobj)s->a); }
    if (s->cn) vb_orelease(s->cn);
    free(s->name);
}
static void rs_invalidate(rrs *s)
{
    if (s->a) {
        drs *v = VW(s->a);
        if (v && v->editmode) { int f; rs_member(s->a->view, "cancelupdate", VBM_METHOD, 0, 0, &f); if (vb_errpend) vb_err_clear_pending(); }
        if (s->a->state) rs_close_m(s->a);
    }
    s->closed = 1;
}

/* the cursor RDO gives for a cursor driver, a type and a lock asked for (MSRDO20 with the SQLite ODBC driver,
 * VB6 oracle f_rdo1); -1 with the error raised */
static int rs_cursor(int drv, int type, int lock, int *rtype, int *rlock, int *upd, int *bm)
{
    if (type < 0 || type > 3) { rdo_fail(40034, "MSRDO.rdoResultset", 0); return -1; }
    if (lock < 1 || lock > 5) { rdo_fail(40019, "MSRDO.rdoResultset", 0); return -1; }
    if (drv == 4 && (type != 0 || lock != 1)) { rdo_fail(40080, "MSRDO.rdoResultset", 0); return -1; }   /* rdUseNone */
    if (lock == 5 && drv != 3) { rdo_fail(40081, "MSRDO.rdoResultset", 0); return -1; }
    switch (drv) {
    case 4:                                     /* rdUseNone: forward-only, read-only: the driver's cursor */
    case 0: case 2:                             /* rdUseIfNeeded, rdUseServer: the keyset of the driver */
        *rtype = 3; *rlock = 2; *upd = 1; *bm = 1;
        return 0;
    case 1:                                     /* rdUseOdbc: the ODBC cursor library */
        if (type == 0) {
            if (lock != 1) { rdo_odbc_state("S1C00", 0, "[Microsoft][ODBC Cursor Library] Cursor library not capable", 0); return -1; }
            *rtype = 0; *rlock = 1; *upd = 0; *bm = 0;
            return 0;
        }
        if (lock == 2) { rdo_odbc_state("S1C00", 0, "[Microsoft][ODBC Cursor Library] Cursor library not capable", 0); return -1; }
        *rtype = 3; *rlock = lock == 1 ? 1 : 4; *upd = lock != 1; *bm = 1;
        return 0;
    default:                                    /* rdUseClientBatch: as asked */
        *rtype = type; *rlock = lock; *upd = lock != 1; *bm = type == 1 || type == 3;
        return 0;
    }
}

static vbvar rs_open_sql(rcn *k, const char *sql, vbvar *vals, int nv, int type, int lock, int maxrows, int drv)
{
    int rtype, rlock, upd, bm;
    if (rs_cursor(drv, type, lock, &rtype, &rlock, &upd, &bm)) return vb_v_empty();
    if (!sql || !*sql) { rdo_fail(40086, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
    ars *r = (ars *)rs_new_for(k->a);
    r->maxrecords = maxrows > 0 ? maxrows : 0;
    long n = 0; int rows = 0;
    int e = odbc_run(k->a, sql, vals, nv, r, &n, &rows);
    if (e) { vb_orelease((vbobj)r); rdo_failed(k->a, 0); return vb_v_empty(); }
    if (!rows) { vb_orelease((vbobj)r); k->affected = n; return retobj(0); }
    r->cursorloc = adUseClient;
    r->cursortype = adOpenStatic;
    r->locktype = upd && !k->ro ? adLockOptimistic : adLockReadOnly;
    r->state = 1;
    free(r->source); r->source = strdup(sql);
    free(r->osql); r->osql = strdup(sql);
    if (nv > 0) {
        r->ovals = calloc((size_t)nv, sizeof(vbvar));
        for (int i = 0; i < nv; i++) vb_vset(&r->ovals[i], vals[i]);
        r->novals = nv;
    }
    vb_vset(&r->filter, vb_v_l(0));
    free(r->sort); r->sort = strdup("");
    rrs *s = (rrs *)vb_onew(&vb_cls_rdo_resultset);
    s->cn = (vbobj)k; vb_oaddref((vbobj)k);
    s->a = r;
    s->name = strdup(sql);
    s->type = rtype; s->lock = rlock; s->drv = drv; s->upd = upd && !k->ro; s->bm = bm;
    s->known = drv == 3;
    s->lastmod = -1; s->addfrom = -1;
    rlist_add(k->rss, (vbobj)s, 0);
    return retobj((vbobj)s);
}

static int rs_rowcount(rrs *s)
{
    if (s->type == 0 || s->rcdel) return -1;
    return s->known ? nalive(VW(s->a)) : 0;
}
static int rs_abspos(rrs *s)
{
    drs *v = VW(s->a);
    if (is_bof(v) || is_eof(v) || v->cur < 0 || v->cur >= v->d->n) return -1;
    return v->cur + 1;                          /* in the keyset: a row deleted keeps its place (VB6 oracle f_rdo2) */
}
/* a change not saved is lost when the resultset moves (as DAO's) */
static void rs_drop_edit(rrs *s)
{
    drs *v = VW(s->a);
    if (v && v->editmode) {
        int add = v->editmode == 2, f;
        rs_member(s->a->view, "cancelupdate", VBM_METHOD, 0, 0, &f);
        if (vb_errpend) vb_err_clear_pending();
        if (add && s->addfrom >= -1 && s->addfrom < v->d->n) v->cur = s->addfrom;
    }
}
/* the RemoteData control that binds the resultset (a listener of its rows) is told before it moves or deletes:
 * "lb_validate" (Action, *cancel): the changed bound values are saved, Validate; then "lb_reposition" after a
 * move; -1 when it canceled (40503) */
static int rdo_validate(rrs *s, int action)
{
    if (!s->a->nhooks) return 0;
    vbvar a[2] = { vb_v_i((vb_i16)action), vb_v_l(0) };
    fire((vbobj)s->a, "lb_validate", 2, a);
    if (vb_errpend) return -1;
    if (vb_v2l(a[1])) { rdo_fail(40503, "MSRDO.rdoResultset", 0); return -1; }
    return 0;
}
/* how: 1 first, 2 previous, 3 next, 4 last, 5 relative n (from the row of a bookmark when bm >= 0) */
static void rs_move(rrs *s, int how, int n, int bm)
{
    drs *v = VW(s->a);
    rdo_errs_clear();
    if (s->type == 0 && (how == 1 || how == 2 || (how == 5 && n < 0))) { rdo_fail(40008, "MSRDO.rdoResultset", 0); return; }
    if (rdo_validate(s, how <= 4 ? how : 9)) return;   /* rdActionMoveFirst ... MoveLast, rdActionBookmark */
    v = VW(s->a);
    int deleted = v->cur >= 0 && v->cur < v->d->n && v->d->r[v->cur].del;
    if (!v->editmode || v->editmode == 1) {
        if (how == 3 && is_eof(v) && nalive(v)) { rdo_fail(40024, "MSRDO.rdoResultset", 0); return; }
        if (how == 2 && is_bof(v) && nalive(v)) { rdo_fail(40025, "MSRDO.rdoResultset", 0); return; }
        if (how == 5 && bm < 0 && !deleted && (is_bof(v) || is_eof(v)) && nalive(v)) { rdo_fail(40029, "MSRDO.rdoResultset", 0); return; }
    }
    if (how == 3 && !nalive(v)) { rdo_fail(40024, "MSRDO.rdoResultset", 0); return; }
    if (how == 2 && !nalive(v)) { rdo_fail(40025, "MSRDO.rdoResultset", 0); return; }
    if (bm >= 0) {
        if (bm >= v->d->n || v->d->r[bm].del) { rdo_fail(40028, "MSRDO.rdoResultset", 0); return; }
    }
    rs_drop_edit(s);
    if (bm >= 0) v->cur = bm;
    do_move(s->a, how, n);
    if (rdo_failed(s->a->cn ? ACN(s->a->cn) : 0, 0)) return;
    if (how == 4) s->known = 1;
    if (s->a->nhooks) fire((vbobj)s->a, "lb_reposition", 0, 0);
}
/* the column of a resultset for a key (a name or an index); -1 with 40041 raised */
static int rs_col(rrs *s, vbvar key)
{
    drs *v = VW(s->a);
    vbvar k = vb_v_letval(key);
    if (vb_errpend) return -1;
    if (k.vt == VT_BSTR) {
        char *nm = vb_stoutf8(k.u.s);
        for (int i = 0; i < v->ncols; i++) if (!strcasecmp(s->a->defs && i < s->a->ndefs ? s->a->defs[i].name : v->cols[i].name, nm)) return i;
    } else if (k.vt != VT_EMPTY && k.vt != VT_NULL) {
        vb_i32 save = vb_errpend;
        vb_i32 i = vb_v2l(k);
        if (vb_errpend && !save) { vb_err_clear_pending(); i = -1; }
        if (i >= 0 && i < v->ncols) return i;
    }
    rdo_fail(40041, "MSRDO.rdoColumns", 0);
    return -1;
}
static vbvar col_obj(rrs *s, int col)
{
    rdcol *c = (rdcol *)vb_onew(&vb_cls_rdo_column);
    c->rs = (vbobj)s; vb_oaddref((vbobj)s);
    c->col = col;
    return retobj((vbobj)c);
}
/* the value of a column of the current row; -1 with the error raised */
static int col_value(rrs *s, int col, vbvar *out)
{
    drs *v = VW(s->a);
    if (v->editmode && v->ebuf) { *out = rdo_value(v->ebuf[col]); return 0; }
    if (!alive(v, v->cur)) { rdo_fail(40009, "MSRDO.rdoColumn", 0); return -1; }
    *out = rdo_value(v->d->r[v->cur].v[col]);
    return 0;
}
static int col_updatable(rrs *s, int col)
{
    otarget *t = s->a->ot;
    return s->upd && t && col < t->ncols && t->cols[col] != 0;
}
static int col_let(rrs *s, int col, vbvar val)
{
    drs *v = VW(s->a);
    if (!v->editmode) { rdo_fail(40039, "MSRDO.rdoColumn", 0); return -1; }
    if (!col_updatable(s, col)) { rdo_fail(40038, "MSRDO.rdoColumn", 0); return -1; }
    vbvar x = vb_v_letval(val);
    if (vb_errpend) return -1;
    if (rs_field_let(v, col, x)) {
        int jet = vb_err_number_pending();
        vb_err_clear_pending();
        if (jet == 3421 || jet == 13) vb_raise(13);
        else rdo_fail(40054, "MSRDO.rdoColumn", 0);
        return -1;
    }
    return 0;
}

/* Edit, AddNew, Update, CancelUpdate, Delete */
static void rs_edit(rrs *s)
{
    drs *v = VW(s->a);
    rdo_errs_clear();
    alist_clear(cn_of(s->a)->errs);
    if (!s->upd) { rdo_fail(40058, "MSRDO.rdoResultset", 0); return; }
    if (v->editmode == 2) rs_drop_edit(s);
    if (v->editmode) return;
    if (!alive(v, v->cur)) { rdo_fail(40009, "MSRDO.rdoResultset", 0); return; }
    dao_call(s->a, s->a->view, "edit", VBM_METHOD, 0, 0, 1);
    rdo_failed(cn_of(s->a), 0);
}
static void rs_addnew(rrs *s)
{
    drs *v = VW(s->a);
    rdo_errs_clear();
    alist_clear(cn_of(s->a)->errs);
    if (!s->upd) { rdo_fail(40058, "MSRDO.rdoResultset", 0); return; }
    if (v->editmode) rs_drop_edit(s);
    s->addfrom = v->cur;
    dao_call(s->a, s->a->view, "addnew", VBM_METHOD, 0, 0, 1);
    rdo_failed(cn_of(s->a), 0);
}
static int rs_update(rrs *s)
{
    drs *v = VW(s->a);
    rdo_errs_clear();
    if (!v->editmode) { rdo_fail(40026, "MSRDO.rdoResultset", 0); return -1; }
    int add = v->editmode == 2, from = s->addfrom;
    alist_clear(cn_of(s->a)->errs);
    dao_call(s->a, s->a->view, "update", VBM_METHOD, 0, 0, 1);
    if (rdo_failed(cn_of(s->a), 0)) return -1;
    s->lastmod = v->lastmod;
    if (add) {                                  /* the current row: the one before AddNew */
        v->cur = from >= -1 && from < v->d->n ? from : v->cur;
        s->addfrom = -1;
    }
    return 0;
}
static void rs_cancel(rrs *s)
{
    rdo_errs_clear();
    rs_drop_edit(s);
    s->addfrom = -1;
}
static void rs_delete(rrs *s)
{
    drs *v = VW(s->a);
    rdo_errs_clear();
    if (!s->upd) { rdo_fail(40058, "MSRDO.rdoResultset", 0); return; }
    if (v->editmode) rs_drop_edit(s);
    if (!alive(v, v->cur)) { rdo_fail(40009, "MSRDO.rdoResultset", 0); return; }
    if (rdo_validate(s, 7)) return;             /* rdActionDelete */
    v = VW(s->a);
    if (v->editmode) rs_drop_edit(s);
    int cur = v->cur;
    alist_clear(cn_of(s->a)->errs);
    dao_call(s->a, s->a->view, "delete", VBM_METHOD, 0, 0, 1);
    if (rdo_failed(cn_of(s->a), 0)) return;
    v->cur = cur;
    s->rcdel = 1;                               /* the bound controls keep showing it (VB6 oracle f_rdo2) */
}

static vbvar rs_getrows(rrs *s, vb_i32 argc, vbvar *argv)
{
    drs *v = VW(s->a);
    int n = along(argc, argv, 0, -1);
    rdo_errs_clear();
    if (!alive(v, v->cur)) { rdo_fail(40009, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
    int *idx = malloc(sizeof(int) * (size_t)(v->d->n + 1));
    int rows = 0;
    for (int i = v->cur; i < v->d->n && (n < 0 || rows < n); i++) if (!v->d->r[i].del) idx[rows++] = i;
    vbarr a; memset(&a, 0, sizeof a);
    vb_adim(&a, sizeof(vbvar), VBA_VAR, 0, 2, 0, v->ncols - 1, 0, rows - 1);
    for (int k = 0; k < rows; k++)
        for (int c = 0; c < v->ncols; c++) vb_vset((vbvar *)vb_aelem(&a, 2, c, k), rdo_value(v->d->r[idx[k]].v[c]));
    vbvar out = vb_v_arr(&a, VT_VARIANT);
    vb_afree(&a);
    int nx = idx[rows - 1] + 1;
    while (nx < v->d->n && v->d->r[nx].del) nx++;
    v->cur = nx;
    if (nx >= v->d->n) s->known = 1;
    free(idx);
    note_fill(s->a);
    return out;
}

static vbvar rcols_ret(rrs *s, vb_i32 kind, vb_i32 argc, vbvar *argv)
{
    if (argc == 0) {
        rlist *l = rlist_new(RL_COLS);
        l->owner = (vbobj)s; vb_oaddref((vbobj)s);
        vbvar v = retobj((vbobj)l);
        vb_orelease((vbobj)l);
        return v;
    }
    int c = rs_col(s, argv[0]);
    if (c < 0) return vb_v_empty();
    if (is_set(kind)) { col_let(s, c, argv[argc - 1]); return vb_v_empty(); }
    return col_obj(s, c);
}

static vbvar rrs_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    rrs *s = (rrs *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (s->closed || !s->a || !s->a->state) { rdo_invalid(0); return vb_v_empty(); }
    drs *v = VW(s->a);
    if (!strcmp(name, "rdocolumns") || !strcmp(name, "")) return rcols_ret(s, kind, argc, argv);
    if (!strcmp(name, "name")) return vstr(s->name);
    if (!strcmp(name, "type")) return vb_v_l(s->type);
    if (!strcmp(name, "locktype")) return vb_v_l(s->lock);
    if (!strcmp(name, "updatable")) return vb_v_bool(s->upd ? -1 : 0);
    if (!strcmp(name, "bookmarkable")) return vb_v_bool(s->bm ? -1 : 0);
    if (!strcmp(name, "restartable") || !strcmp(name, "transactions")) return vb_v_bool(-1);
    if (!strcmp(name, "stillexecuting")) return vb_v_bool(0);
    if (!strcmp(name, "lockedits")) return vb_v_bool(s->lock == 2 ? -1 : 0);
    if (!strcmp(name, "status")) return vb_v_l(0);
    if (!strcmp(name, "bof")) return vb_v_bool(is_bof(v) ? -1 : 0);
    if (!strcmp(name, "eof")) return vb_v_bool(is_eof(v) ? -1 : 0);
    if (!strcmp(name, "editmode")) return vb_v_l(v->editmode == 2 ? 2 : v->editmode ? 1 : 0);
    if (!strcmp(name, "rowcount")) return vb_v_l(rs_rowcount(s));
    if (!strcmp(name, "activeconnection")) {
        if (set) return vb_v_empty();
        return retobj(s->cn);
    }
    if (!strcmp(name, "absoluteposition")) {
        if (set) {
            int n = vb_v2l(argv[argc - 1]);
            if (s->type == 0) { rdo_fail(40008, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
            if (n >= 1 && n <= v->d->n && !v->d->r[n - 1].del) { rs_move(s, 5, 0, n - 1); return vb_v_empty(); }
            rdo_fail(40054, "MSRDO.rdoResultset", 0);
            return vb_v_empty();
        }
        return vb_v_l(s->type == 0 ? -1 : rs_abspos(s));
    }
    if (!strcmp(name, "percentposition")) {
        int n = v->d->n, p = rs_abspos(s);
        if (set) {
            double f = vb_v2d(argv[argc - 1]);
            int want = n ? (int)(f * n / 100.0 + 0.5) : 0, k = 0;
            if (want < 1) want = 1;
            for (int i = 0; i < v->d->n; i++) if (!v->d->r[i].del && ++k == want) { rs_move(s, 5, 0, i); break; }
            return vb_v_empty();
        }
        return vb_v_f(p > 0 && n ? (float)(p * 100.0 / n) : 0.0f);
    }
    if (!strcmp(name, "bookmark")) {
        if (!s->bm) { rdo_fail(40027, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
        if (set) {
            vbvar b = vb_v_letval(argv[argc - 1]);
            if (vb_errpend) return vb_v_empty();
            if (b.vt == VT_EMPTY || b.vt == VT_NULL || b.vt == VT_BSTR) { rdo_fail(40027, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
            int i = (int)vb_v2d(b) - 1;
            if (i < 0 || i >= v->d->n || v->d->r[i].del) { rdo_fail(40027, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
            rs_move(s, 5, 0, i);
            return vb_v_empty();
        }
        if (!alive(v, v->cur)) { rdo_fail(40009, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
        return vb_v_d(v->cur + 1);
    }
    if (!strcmp(name, "lastmodified")) {
        if (s->lastmod < 0 || s->lastmod >= v->d->n) { rdo_fail(40027, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
        return vb_v_d(s->lastmod + 1);
    }
    if (!strcmp(name, "movefirst")) { rs_move(s, 1, 0, -1); return vb_v_empty(); }
    if (!strcmp(name, "movelast")) { rs_move(s, 4, 0, -1); return vb_v_empty(); }
    if (!strcmp(name, "movenext")) { rs_move(s, 3, 0, -1); return vb_v_empty(); }
    if (!strcmp(name, "moveprevious")) { rs_move(s, 2, 0, -1); return vb_v_empty(); }
    if (!strcmp(name, "move")) {
        int n = along(argc, argv, 0, 0), bm = -1;
        if (!missing(argc, argv, 1)) {
            vbvar b = vb_v_letval(argv[1]);
            bm = (int)vb_v2d(b) - 1;
            if (bm < 0) { rdo_fail(40028, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
        }
        rs_move(s, 5, n, bm);
        return vb_v_empty();
    }
    if (!strcmp(name, "edit")) { rs_edit(s); return vb_v_empty(); }
    if (!strcmp(name, "addnew")) { rs_addnew(s); return vb_v_empty(); }
    if (!strcmp(name, "update")) { rs_update(s); return vb_v_empty(); }
    if (!strcmp(name, "cancelupdate")) { rs_cancel(s); return vb_v_empty(); }
    if (!strcmp(name, "delete")) { rs_delete(s); return vb_v_empty(); }
    if (!strcmp(name, "requery")) {
        rdo_errs_clear();
        rs_drop_edit(s);
        int ct = s->a->cursortype, lt = s->a->locktype;
        int f;
        rs_member2((vbobj)s->a, "requery", VBM_METHOD, 0, 0, &f);
        s->a->cursortype = ct; s->a->locktype = lt;
        if (rdo_failed(cn_of(s->a), 0)) return vb_v_empty();
        s->known = s->drv == 3; s->rcdel = 0; s->lastmod = -1;
        return vb_v_empty();
    }
    if (!strcmp(name, "close")) {
        rcn *k = s->cn ? RCN(s->cn) : 0;
        vb_oaddref((vbobj)s);
        rs_invalidate(s);
        if (k) rlist_remove(k->rss, (vbobj)s);
        vb_orelease((vbobj)s);
        return vb_v_empty();
    }
    if (!strcmp(name, "getrows")) return rs_getrows(s, argc, argv);
    if (!strcmp(name, "getclipstring")) {
        rdo_errs_clear();
        if (!alive(v, v->cur)) { rdo_fail(40009, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
        vbvar a[5];
        a[0] = vb_v_missing();
        for (int i = 0; i < 4; i++) a[i + 1] = missing(argc, argv, i) ? vb_v_missing() : argv[i];
        vbvar r = getstring_m(s->a, 5, a);
        if (VW(s->a)->cur >= VW(s->a)->d->n) s->known = 1;
        note_fill(s->a);
        return r;
    }
    if (!strcmp(name, "moreresults")) {
        rdo_errs_clear();
        rs_invalidate(s);                       /* the one result set of the statement: no more */
        rcn *k = s->cn ? RCN(s->cn) : 0;
        if (k) { vb_oaddref((vbobj)s); rlist_remove(k->rss, (vbobj)s); vb_orelease((vbobj)s); }
        return vb_v_bool(0);
    }
    if (!strcmp(name, "cancel") || !strcmp(name, "resync") || !strcmp(name, "cancelbatch")) return vb_v_empty();
    if (!strcmp(name, "batchupdate")) { if (s->drv != 3) { rdo_fail(40070, "MSRDO.rdoResultset", 0); } return vb_v_empty(); }
    if (!strcmp(name, "batchsize") || !strcmp(name, "updatecriteria") || !strcmp(name, "updateoperation") ||
        !strcmp(name, "batchcollisioncount") || !strcmp(name, "batchcollisionrows")) {
        if (s->drv != 3) { rdo_fail(40070, "MSRDO.rdoResultset", 0); return vb_v_empty(); }
        if (set) return vb_v_empty();
        if (!strcmp(name, "batchsize")) return vb_v_l(15);
        if (!strcmp(name, "updatecriteria")) return vb_v_l(1);
        return vb_v_l(0);
    }
    *found = 0;
    return vb_v_empty();
}

/* the inner ADO recordset of an rdoResultset (the RemoteData control binds through it), or 0 */
vbobj rdo_rs_inner(vbobj o)
{
    o = o ? vb_oid(o) : 0;
    if (!o || o->cls != &vb_cls_rdo_resultset) return 0;
    rrs *s = (rrs *)o;
    return s->closed ? 0 : (vbobj)s->a;
}
/* ---------------------------------------------------------------- rdoColumn */
static void rcol_clear(vbobj o)
{
    rdcol *c = (rdcol *)o;
    if (c->rs) vb_orelease(c->rs);
    free(c->name); free(c->stab); free(c->scol);
}
static vbvar rcol_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    rdcol *c = (rdcol *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (!c->rs) {                               /* a column of an rdoTable */
        if (!strcmp(name, "name")) return vstr(c->name);
        if (!strcmp(name, "type")) return vb_v_l(c->type);
        if (!strcmp(name, "size")) return vb_v_l(c->size);
        if (!strcmp(name, "attributes")) return vb_v_l(c->attrs);
        if (!strcmp(name, "required")) return vb_v_bool(c->req ? -1 : 0);
        if (!strcmp(name, "allowzerolength")) return vb_v_bool(c->zl ? -1 : 0);
        if (!strcmp(name, "ordinalposition")) return vb_v_l(c->ord);
        if (!strcmp(name, "sourcetable")) return vstr(c->stab);
        if (!strcmp(name, "sourcecolumn")) return vstr(c->scol);
        if (!strcmp(name, "updatable") || !strcmp(name, "keycolumn") || !strcmp(name, "chunkrequired")) return vb_v_bool(0);
        if (!strcmp(name, "value") || !strcmp(name, "") || !strcmp(name, "originalvalue")) { rdo_fail(40009, "MSRDO.rdoColumn", 0); return vb_v_empty(); }
        *found = 0;
        return vb_v_empty();
    }
    rrs *s = RRS(c->rs);
    if (s->closed || !s->a || !s->a->state) { rdo_invalid(0); return vb_v_empty(); }
    drs *v = VW(s->a);
    int col = c->col;
    adef *d = s->a->defs && col < s->a->ndefs ? &s->a->defs[col] : 0;
    int t = rdo_sqltype(d ? d->type : adVarChar);
    otarget *ot = s->a->ot;
    if (!strcmp(name, "value") || !strcmp(name, "")) {
        if (set) { col_let(s, col, argv[argc - 1]); return vb_v_empty(); }
        vbvar x;
        if (col_value(s, col, &x)) return vb_v_empty();
        return vb_vtmp(x);
    }
    if (strcmp(name, "originalvalue") && strcmp(name, "columnsize") && strcmp(name, "getchunk") && strcmp(name, "appendchunk") &&
        strcmp(name, "status"))
        rdo_errs_clear();                       /* the attributes of the column, asked of ODBC (VB6 oracle f_rdo1) */
    if (!strcmp(name, "name")) return vstr(d ? d->name : v->cols[col].name);
    if (!strcmp(name, "type")) return vb_v_l(t);
    if (!strcmp(name, "size")) return vb_v_l(rdo_size_of(t, d ? (t == 2 || t == 3 ? d->prec : d->size) : 0));
    if (!strcmp(name, "attributes")) {
        int a = (t == 12 || t == -1 || t == -3 || t == -4 || t == -9 || t == -10) ? 2 : 1;
        if (col_updatable(s, col)) a |= 32;
        if (ot && ot->autocol == col) a |= 16;
        return vb_v_l(a);
    }
    if (!strcmp(name, "required")) return vb_v_bool(d && !(d->attrs & adFldIsNullable) ? -1 : 0);
    if (!strcmp(name, "allowzerolength")) return vb_v_bool(t == 1 || t == 12 || t == -1 || t == -8 || t == -9 || t == -10 ? -1 : 0);
    if (!strcmp(name, "ordinalposition")) return vb_v_l(col + 1);
    if (!strcmp(name, "sourcetable")) return vstr(ot && ot->table && col < ot->ncols && ot->cols[col] ? ot->table : "");
    if (!strcmp(name, "sourcecolumn")) return vstr(ot && col < ot->ncols && ot->cols[col] ? ot->cols[col] : "");
    if (!strcmp(name, "updatable")) return vb_v_bool(col_updatable(s, col) ? -1 : 0);
    if (!strcmp(name, "chunkrequired")) return vb_v_bool(t == -1 || t == -4 || t == -10 ? -1 : 0);
    if (!strcmp(name, "keycolumn")) return vb_v_bool(ot && col < ot->ncols && ot->iskey[col] ? -1 : 0);
    if (!strcmp(name, "status")) return vb_v_l(v->editmode && v->edirty && v->edirty[col] ? 1 : 0);
    if (!strcmp(name, "originalvalue")) {
        if (!alive(v, v->cur)) { rdo_fail(40009, "MSRDO.rdoColumn", 0); return vb_v_empty(); }
        return vb_vtmp(rdo_value(v->d->r[v->cur].v[col]));
    }
    if (!strcmp(name, "columnsize")) {
        vbvar x;
        if (col_value(s, col, &x)) return vb_v_empty();
        if (x.vt == VT_NULL || x.vt == VT_EMPTY) return vb_v_l(-1);
        if (x.vt & VT_ARRAY) { vbarr *a = vb_v_arrptr(x); return vb_v_l(a && a->nd ? a->n : 0); }
        return vb_v_l(vb_v2s(x).len);
    }
    if (!strcmp(name, "getchunk")) {
        vbvar x;
        if (col_value(s, col, &x)) return vb_v_empty();
        if (x.vt == VT_NULL || x.vt == VT_EMPTY || (x.vt & VT_ARRAY)) return vb_vtmp(x);
        int n = along(argc, argv, 0, -1);
        vbs str = vb_v2s(x);
        if (n >= 0 && n < str.len) return vb_v_s(vb_left(str, n));
        return vb_v_s(str);
    }
    if (!strcmp(name, "appendchunk")) {
        if (!v->editmode) { rdo_fail(40039, "MSRDO.rdoColumn", 0); return vb_v_empty(); }
        vbvar cur = v->ebuf ? v->ebuf[col] : vb_v_null();
        vbvar add = argc ? vb_v_letval(argv[0]) : vb_v_empty();
        if (cur.vt == VT_NULL || cur.vt == VT_EMPTY) col_let(s, col, add);
        else col_let(s, col, vb_v_s(vb_scat(vb_v2s(cur), vb_v2s(add))));
        return vb_v_empty();
    }
    if (!strcmp(name, "batchconflictvalue")) return vb_v_null();
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- rdoParameter, rdoQuery */
static void rpar_init(vbobj o) { rpar *p = (rpar *)o; p->name = strdup(""); p->type = 12; }
static void rpar_clear(vbobj o) { rpar *p = (rpar *)o; free(p->name); vb_vfree(&p->value); }
static vbvar rpar_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    rpar *p = (rpar *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (!strcmp(name, "value") || !strcmp(name, "")) {
        if (set) { vb_vset(&p->value, vb_v_letval(argv[argc - 1])); return vb_v_empty(); }
        return vb_vtmp(p->value);
    }
    if (!strcmp(name, "name")) return vstr(p->name);
    if (!strcmp(name, "direction")) { if (set) { p->dir = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(p->dir); }
    if (!strcmp(name, "type")) { if (set) { p->type = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(p->type); }
    if (!strcmp(name, "size")) { if (set) { p->size = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(p->size); }
    if (!strcmp(name, "appendchunk")) {
        vbvar add = argc ? vb_v_letval(argv[0]) : vb_v_empty();
        if (p->value.vt == VT_NULL || p->value.vt == VT_EMPTY) vb_vset(&p->value, add);
        else vb_vset(&p->value, vb_v_s(vb_scat(vb_v2s(p->value), vb_v2s(add))));
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* the ? markers of a statement, outside its quotes */
static int count_markers(const char *s)
{
    int n = 0; char q = 0;
    for (; s && *s; s++) {
        if (q) { if (*s == q) q = 0; continue; }
        if (*s == '\'' || *s == '"' || *s == '[') { q = *s == '[' ? ']' : *s; continue; }
        if (*s == '?') n++;
    }
    return n;
}
/* the rdoParameters of the SQL of a query: one for each ?, their types as the driver describes them */
static void qry_params(rqry *q)
{
    int n = count_markers(q->sql);
    if (q->pars->n == n) return;
    rlist_clear(q->pars);
    rcn *k = q->cn ? RCN(q->cn) : 0;
    lbo_stmt *st = 0;
    if (k && k->connected && k->a && n > 0) {
        st = lbo_stmt_new(k->a->odbc);
        if (st && lbo_prepare(st, q->sql)) { lbo_stmt_free(st); st = 0; }
    }
    for (int i = 1; i <= n; i++) {
        rpar *p = (rpar *)vb_onew(&vb_cls_rdo_parameter);
        char nm[32]; snprintf(nm, sizeof nm, "Parameter%d", i);
        free(p->name); p->name = strdup(nm);
        int ty = 0, dec = 0, nl = 0; unsigned long sz = 0;
        if (st && !lbo_describe_param(st, i, &ty, &sz, &dec, &nl)) { p->type = rdo_wide_sql(k->a, ty); p->size = (int)sz; }
        rlist_add(q->pars, (vbobj)p, 0);
    }
    if (st) lbo_stmt_free(st);
    if (vb_errpend) vb_err_clear_pending();
}
static void rqry_init(vbobj o)
{
    rqry *q = (rqry *)o;
    q->name = strdup(""); q->sql = strdup("");
    q->pars = rlist_new(RL_PARS);
    q->prepared = 1; q->rowset = 100; q->lock = 1; q->timeout = 30; q->bindthr = 1024; q->errthr = -1;
}
static void rqry_clear(vbobj o)
{
    rqry *q = (rqry *)o;
    if (q->cn) vb_orelease(q->cn);
    free(q->name); free(q->sql);
    vb_orelease((vbobj)q->pars);
}
static void qry_set_cn(rqry *q, rcn *k)
{
    if (q->cn) { rlist_remove(RCN(q->cn)->qrys, (vbobj)q); vb_orelease(q->cn); q->cn = 0; }
    if (k) { q->cn = (vbobj)k; vb_oaddref((vbobj)k); rlist_add(k->qrys, (vbobj)q, 0); }
}
static vbvar qry_new_on(rcn *k, const char *name, const char *sql)
{
    rqry *q = (rqry *)vb_onew(&vb_cls_rdo_query);
    free(q->name); q->name = strdup(name);
    free(q->sql); q->sql = strdup(sql);
    q->cursor = 0;
    qry_set_cn(q, k);
    qry_params(q);
    return retobj((vbobj)q);
}
/* the values of the parameters, for ODBC */
static vbvar *qry_vals(rqry *q, int *nv)
{
    qry_params(q);
    *nv = q->pars->n;
    vbvar *vals = calloc((size_t)(*nv ? *nv : 1), sizeof(vbvar));
    for (int i = 0; i < *nv; i++) vb_vset(&vals[i], ((rpar *)vb_oid(q->pars->items[i]))->value);
    return vals;
}
static void free_vals(vbvar *vals, int nv) { for (int i = 0; i < nv; i++) vb_vfree(&vals[i]); free(vals); }

static vbvar rqry_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    rqry *q = (rqry *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (q->closed) { rdo_invalid(0); return vb_v_empty(); }
    if (!strcmp(name, "rdoparameters") || !strcmp(name, "")) { qry_params(q); return rlist_ret(q->pars, kind, argc, argv); }
    if (!strcmp(name, "name")) { if (set) { free(q->name); q->name = rdo_sdup(argv[argc - 1]); return vb_v_empty(); } return vstr(q->name); }
    if (!strcmp(name, "sql")) {
        if (set) { free(q->sql); q->sql = rdo_sdup(argv[argc - 1]); qry_params(q); return vb_v_empty(); }
        return vstr(q->sql);
    }
    if (!strcmp(name, "activeconnection")) {
        if (set) {
            vbvar a = argv[argc - 1];
            vbobj co = vb_isobject(a) ? vb_v2obj(a) : 0;
            co = co ? vb_oid(co) : 0;
            if (co && co->cls != &vb_cls_rdo_connection) { vb_raise(13); return vb_v_empty(); }
            vb_oaddref(o);
            qry_set_cn(q, (rcn *)co);
            vb_orelease(o);
            int n = q->pars->n; (void)n;
            rlist_clear(q->pars);
            qry_params(q);
            return vb_v_empty();
        }
        return retobj(q->cn);
    }
    if (!strcmp(name, "type")) { if (set) { q->type = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->type); }
    if (!strcmp(name, "prepared")) { if (set) { q->prepared = vb_v2bool(argv[argc - 1]) != 0; return vb_v_empty(); } return vb_v_bool(q->prepared ? -1 : 0); }
    if (!strcmp(name, "maxrows")) { if (set) { q->maxrows = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->maxrows); }
    if (!strcmp(name, "rowsetsize")) { if (set) { q->rowset = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->rowset); }
    if (!strcmp(name, "keysetsize")) { if (set) { q->keyset = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->keyset); }
    if (!strcmp(name, "cursortype")) { if (set) { q->cursor = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->cursor); }
    if (!strcmp(name, "locktype")) { if (set) { q->lock = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->lock); }
    if (!strcmp(name, "querytimeout")) { if (set) { q->timeout = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->timeout); }
    if (!strcmp(name, "bindthreshold")) { if (set) { q->bindthr = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->bindthr); }
    if (!strcmp(name, "errorthreshold")) { if (set) { q->errthr = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->errthr); }
    if (!strcmp(name, "rowsaffected")) return vb_v_l((vb_i32)q->affected);
    if (!strcmp(name, "stillexecuting")) return vb_v_bool(0);
    if (!strcmp(name, "updatable")) return vb_v_bool(-1);
    if (!strcmp(name, "hstmt")) return vb_v_l(0);
    if (!strcmp(name, "rdocolumns")) { rlist *l = rlist_new(RL_COLS); vbvar v = retobj((vbobj)l); vb_orelease((vbobj)l); return v; }
    if (!strcmp(name, "cancel")) return vb_v_empty();
    if (!strcmp(name, "close")) {
        vb_oaddref(o);
        qry_set_cn(q, 0);
        q->closed = 1;
        vb_orelease(o);
        return vb_v_empty();
    }
    rcn *k = q->cn ? RCN(q->cn) : 0;
    if (!strcmp(name, "execute") || !strcmp(name, "openresultset")) {
        if (!k || !k->connected) { rdo_fail(40074, "MSRDO.rdoQuery", name[0] == 'e' ? RDO_SQRY : 0); return vb_v_empty(); }
        int nv;
        vbvar *vals = qry_vals(q, &nv);
        vbvar r = vb_v_empty();
        if (name[0] == 'e') {
            long n = rdo_execute(k, q->sql, vals, nv);
            if (n >= 0) q->affected = k->affected = n;
        } else r = rs_open_sql(k, q->sql, vals, nv, along(argc, argv, 0, q->cursor), along(argc, argv, 1, q->lock), q->maxrows, k->cursor);
        free_vals(vals, nv);
        return r;
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- rdoTable */
static void rtab_clear(vbobj o)
{
    rtab *t = (rtab *)o;
    if (t->cn) vb_orelease(t->cn);
    free(t->name); free(t->type);
    if (t->cols) vb_orelease((vbobj)t->cols);
}
static void tab_cols(rtab *t)
{
    if (t->cols) return;
    t->cols = rlist_new(RL_COLS);
    rcn *k = t->cn ? RCN(t->cn) : 0;
    if (!k || !k->connected) return;
    lbo_stmt *st = lbo_stmt_new(k->a->odbc);
    if (st && !lbo_columns(st, 0, 0, t->name, 0)) {
        int ord = 0;
        while (lbo_fetch(st) == 1) {
            lbo_val v; memset(&v, 0, sizeof v);
            rdcol *c = (rdcol *)vb_onew(&vb_cls_rdo_column);
            c->name = strdup(""); c->stab = strdup(t->name); c->scol = strdup("");
            memset(&v, 0, sizeof v);
            if (!lbo_get(st, 4, LBS_VARCHAR, &v) && v.s) { free(c->name); c->name = strdup(v.s); free(c->scol); c->scol = strdup(v.s); }
            lbo_val_free(&v);
            memset(&v, 0, sizeof v);
            if (!lbo_get(st, 5, LBS_INTEGER, &v)) c->type = (int)v.iv;
            lbo_val_free(&v);
            memset(&v, 0, sizeof v);
            if (!lbo_get(st, 7, LBS_INTEGER, &v)) c->size = (int)v.iv;
            lbo_val_free(&v);
            memset(&v, 0, sizeof v);
            if (!lbo_get(st, 11, LBS_INTEGER, &v)) c->req = v.kind != LBO_NULL && v.iv == 0;
            lbo_val_free(&v);
            c->size = rdo_size_of(c->type, c->size);
            c->attrs = (c->type == 12 || c->type == -1 || c->type == -3 || c->type == -4 || c->type == -9 || c->type == -10) ? 2 : 1;
            c->zl = c->type == 1 || c->type == 12 || c->type == -1 || c->type == -8 || c->type == -9 || c->type == -10;
            c->ord = ++ord;
            rlist_add(t->cols, (vbobj)c, 0);
        }
    }
    lbo_stmt_free(st);
    if (vb_errpend) vb_err_clear_pending();
}
static vbvar rtab_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    rtab *t = (rtab *)vb_oid(o);
    *found = 1;
    rcn *k = t->cn ? RCN(t->cn) : 0;
    if (!k || k->closed) { rdo_invalid(0); return vb_v_empty(); }
    if (!strcmp(name, "name") || !strcmp(name, "")) return vstr(t->name);
    if (!strcmp(name, "type")) return vstr(t->type);
    if (!strcmp(name, "updatable")) return vb_v_bool(0);
    if (!strcmp(name, "rowcount")) return vb_v_l(0);
    if (!strcmp(name, "rdocolumns")) { tab_cols(t); return rlist_ret(t->cols, kind, argc, argv); }
    if (!strcmp(name, "openresultset")) {
        if (!k->connected) { rdo_fail(40071, "MSRDO.rdoTable", 0); return vb_v_empty(); }
        jbuf b = { 0 };
        jb_s(&b, "SELECT * FROM "); jb_s(&b, t->name);
        vbvar v = rs_open_sql(k, b.p, 0, 0, along(argc, argv, 0, 0), along(argc, argv, 1, 1), 0, k->cursor);
        jb_free(&b);
        return v;
    }
    *found = 0;
    return vb_v_empty();
}
static void tabs_load(rcn *k)
{
    rlist_clear(k->tabs);
    k->tabs_loaded = 1;
    if (!k->connected) return;
    rdo_errs_clear();
    lbo_stmt *st = lbo_stmt_new(k->a->odbc);
    if (!st || lbo_tables(st, 0, 0, 0, 0)) {
        if (st) lbo_stmt_free(st);
        alist_clear(k->a->errs);
        odbc_raise(k->a);
        rdo_failed(k->a, 0);
        return;
    }
    while (lbo_fetch(st) == 1) {
        lbo_val v, w;
        memset(&v, 0, sizeof v); memset(&w, 0, sizeof w);
        if (lbo_get(st, 3, LBS_VARCHAR, &v) || !v.s) { lbo_val_free(&v); continue; }
        if (lbo_get(st, 4, LBS_VARCHAR, &w)) { lbo_val_free(&w); memset(&w, 0, sizeof w); }
        const char *ty = w.s ? w.s : "TABLE";
        if (!strcasecmp(ty, "SYSTEM TABLE") || !strncasecmp(v.s, "sqlite_", 7)) { lbo_val_free(&v); lbo_val_free(&w); continue; }
        rtab *t = (rtab *)vb_onew(&vb_cls_rdo_table);
        t->cn = (vbobj)k; vb_oaddref((vbobj)k);
        t->name = strdup(v.s); t->type = strdup(ty);
        rlist_add(k->tabs, (vbobj)t, 0);
        lbo_val_free(&v); lbo_val_free(&w);
    }
    lbo_stmt_free(st);
    if (vb_errpend) vb_err_clear_pending();
}

/* ---------------------------------------------------------------- rdoError */
static void rerr_clear(vbobj o) { rerr *e = (rerr *)o; free(e->desc); free(e->src); free(e->state); }
static vbvar rerr_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    rerr *e = (rerr *)vb_oid(o);
    *found = 1;
    if (!strcmp(name, "number")) return vb_v_l(e->number);
    if (!strcmp(name, "description") || !strcmp(name, "")) return vstr(e->desc);
    if (!strcmp(name, "source")) return vstr(e->src);
    if (!strcmp(name, "sqlstate")) return vstr(e->state);
    if (!strcmp(name, "sqlretcode")) return vb_v_l(e->retcode);
    if (!strcmp(name, "helpfile")) return vstr("");
    if (!strcmp(name, "helpcontext")) return vb_v_l(0);
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- the collections */
static const char *rl_name(vbobj o)
{
    const vb_class *c = o->cls;
    if (c == &vb_cls_rdo_environment) return ((renv *)o)->name;
    if (c == &vb_cls_rdo_connection) return ((rcn *)o)->name;
    if (c == &vb_cls_rdo_query) return ((rqry *)o)->name;
    if (c == &vb_cls_rdo_resultset) return ((rrs *)o)->name;
    if (c == &vb_cls_rdo_parameter) return ((rpar *)o)->name;
    if (c == &vb_cls_rdo_table) return ((rtab *)o)->name;
    if (c == &vb_cls_rdo_column) return ((rdcol *)o)->name;
    return 0;
}
static vb_i32 rl_count(rlist *l)
{
    if (l->kind == RL_COLS && l->owner) {
        rrs *s = RRS(l->owner);
        return s->closed || !s->a || !s->a->state ? 0 : VW(s->a)->ncols;
    }
    return l->n;
}
static vb_i32 rl_each_count(vbobj o) { return rl_count((rlist *)vb_oid(o)); }
static vbvar rl_each_item(vbobj o, vb_i32 i)
{
    rlist *l = (rlist *)vb_oid(o);
    if (l->kind == RL_COLS && l->owner) return i >= 0 && i < rl_count(l) ? col_obj(RRS(l->owner), i) : vb_v_empty();
    return i >= 0 && i < l->n ? retobj(l->items[i]) : vb_v_empty();
}
/* Item(key): the index, or -1 with 40041 raised */
static int rl_index(rlist *l, vbvar key)
{
    vbvar k = vb_v_letval(key);
    if (vb_errpend) return -1;
    if (k.vt == VT_BSTR) {
        char *nm = vb_stoutf8(k.u.s);
        for (int i = 0; i < l->n; i++) {
            const char *n = rl_name(vb_oid(l->items[i]));
            if (n && !strcasecmp(n, nm)) return i;
        }
    } else if (k.vt != VT_EMPTY && k.vt != VT_NULL) {
        vb_i32 save = vb_errpend;
        vb_i32 i = vb_v2l(k);
        if (vb_errpend && !save) { vb_err_clear_pending(); i = -1; }
        if (i >= 0 && i < l->n) return i;
    }
    rdo_fail(40041, rl_src[l->kind], 0);
    return -1;
}
static vbvar rl_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    rlist *l = (rlist *)vb_oid(o);
    *found = 1;
    if (l->kind == RL_COLS && l->owner) {       /* the rdoColumns of a resultset */
        rrs *s = RRS(l->owner);
        if (s->closed || !s->a || !s->a->state) { rdo_invalid(0); return vb_v_empty(); }
        if (!strcmp(name, "count")) return vb_v_l(rl_count(l));
        if (!strcmp(name, "item") || !strcmp(name, "")) {
            if (argc < 1) { vb_raise(450); return vb_v_empty(); }
            return rcols_ret(s, kind, argc, argv);
        }
        if (!strcmp(name, "refresh")) return vb_v_empty();
        *found = 0;
        return vb_v_empty();
    }
    if (!strcmp(name, "count")) return vb_v_l(l->n);
    if (!strcmp(name, "item") || !strcmp(name, "")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = rl_index(l, argv[0]);
        if (i < 0) return vb_v_empty();
        if (is_set(kind) && argc > 1) {        /* q.rdoParameters(0) = x: the value of the item */
            vb_late_o(l->items[i], "", kind, 1, &argv[argc - 1]);
            return vb_v_empty();
        }
        return retobj(l->items[i]);
    }
    if (!strcmp(name, "refresh")) {
        if (l->kind == RL_TABS && l->back) tabs_load((rcn *)l->back);
        return vb_v_empty();
    }
    if (!strcmp(name, "clear") && l->kind == RL_ERRS) { rlist_clear(l); return vb_v_empty(); }
    *found = 0;
    return vb_v_empty();
}
/* a collection as a member gives it (no arguments) or its item */
static vbvar rlist_ret(rlist *l, vb_i32 kind, vb_i32 argc, vbvar *argv)
{
    if (argc == 0) return retobj((vbobj)l);
    int f;
    return rl_member((vbobj)l, "item", kind, argc, argv, &f);
}

/* ---------------------------------------------------------------- the classes */
/* TypeName: "Object" (MSRDO20 gives no type information for them), except the two that can be made with New */
const vb_class vb_cls_rdo_engine = { "Object", sizeof(reng), 0, 0, eng_clear, 0, 0, "rdoenvironments", FB(eng_member) };
const vb_class vb_cls_rdo_environment = { "Object", sizeof(renv), 0, 0, env_clear, 0, 0, "rdoconnections", FB(env_member) };
const vb_class vb_cls_rdo_connection = { "rdoConnection", sizeof(rcn), rcn_init, 0, rcn_clear, 0, 0, "name", FB(rcn_member) };
const vb_class vb_cls_rdo_query = { "rdoQuery", sizeof(rqry), rqry_init, 0, rqry_clear, 0, 0, "rdoparameters", FB(rqry_member) };
const vb_class vb_cls_rdo_resultset = { "Object", sizeof(rrs), 0, 0, rrs_clear, 0, 0, "rdocolumns", FB(rrs_member) };
const vb_class vb_cls_rdo_column = { "Object", sizeof(rdcol), 0, 0, rcol_clear, 0, 0, "value", FB(rcol_member) };
const vb_class vb_cls_rdo_parameter = { "Object", sizeof(rpar), rpar_init, 0, rpar_clear, 0, 0, "value", FB(rpar_member) };
const vb_class vb_cls_rdo_table = { "Object", sizeof(rtab), 0, 0, rtab_clear, 0, 0, "name", FB(rtab_member) };
const vb_class vb_cls_rdo_error = { "Object", sizeof(rerr), 0, 0, rerr_clear, 0, 0, "description", FB(rerr_member) };
const vb_class vb_cls_rdo_coll = { "Object", sizeof(rlist), 0, 0, rlist_free, 0, 0, "item", rl_each_count, rl_each_item, 0, rl_member };

/* CreateObject("RDO.rdoConnection") ... */
vbobj rdo_createobject(const char *progid)
{
    if (strncmp(progid, "rdo.", 4) && strncmp(progid, "msrdo20.", 8)) return 0;
    const char *c = strchr(progid, '.') + 1;
    size_t n = strcspn(c, ".");
    if (n == 13 && !strncmp(c, "rdoconnection", 13)) return vb_onew(&vb_cls_rdo_connection);
    if (n == 8 && !strncmp(c, "rdoquery", 8)) return vb_onew(&vb_cls_rdo_query);
    if (n == 9 && !strncmp(c, "rdoengine", 9)) return vb_rdo_engine();
    return 0;
}
