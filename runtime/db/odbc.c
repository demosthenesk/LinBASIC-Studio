/* odbc.c - ODBC for the ADO provider MSDASQL (ado_odbc.inc): the driver manager of the system,
 * loaded when a program first connects (Linux: unixODBC, libodbc.so.2; Windows: odbc32.dll), so
 * that programs without ODBC need neither. The wide-character (W) functions are used on both
 * systems; strings are UTF-8 here. Apart from vbrt.h (windows.h names VT_EMPTY too).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdint.h>
#ifdef _WIN32
#include <windows.h>
#else
#include <dlfcn.h>
#endif
#if defined(_WIN32) || (defined(__has_include) && __has_include(<sqlext.h>))
#include <sql.h>
#include <sqlext.h>
#include <sqlucode.h>
#else
#include "odbc_types.h"                         /* a computer without the ODBC headers */
#endif
#include "odbc.h"

/* ---------------------------------------------------------------- the driver manager */
static struct {
    int tried, ok;
    SQLRETURN (SQL_API *AllocHandle)(SQLSMALLINT, SQLHANDLE, SQLHANDLE *);
    SQLRETURN (SQL_API *FreeHandle)(SQLSMALLINT, SQLHANDLE);
    SQLRETURN (SQL_API *SetEnvAttr)(SQLHENV, SQLINTEGER, SQLPOINTER, SQLINTEGER);
    SQLRETURN (SQL_API *DriverConnectW)(SQLHDBC, SQLHWND, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLSMALLINT *, SQLUSMALLINT);
    SQLRETURN (SQL_API *Disconnect)(SQLHDBC);
    SQLRETURN (SQL_API *GetInfoW)(SQLHDBC, SQLUSMALLINT, SQLPOINTER, SQLSMALLINT, SQLSMALLINT *);
    SQLRETURN (SQL_API *SetConnectAttrW)(SQLHDBC, SQLINTEGER, SQLPOINTER, SQLINTEGER);
    SQLRETURN (SQL_API *EndTran)(SQLSMALLINT, SQLHANDLE, SQLSMALLINT);
    SQLRETURN (SQL_API *PrepareW)(SQLHSTMT, SQLWCHAR *, SQLINTEGER);
    SQLRETURN (SQL_API *ExecDirectW)(SQLHSTMT, SQLWCHAR *, SQLINTEGER);
    SQLRETURN (SQL_API *Execute)(SQLHSTMT);
    SQLRETURN (SQL_API *NumParams)(SQLHSTMT, SQLSMALLINT *);
    SQLRETURN (SQL_API *DescribeParam)(SQLHSTMT, SQLUSMALLINT, SQLSMALLINT *, SQLULEN *, SQLSMALLINT *, SQLSMALLINT *);
    SQLRETURN (SQL_API *BindParameter)(SQLHSTMT, SQLUSMALLINT, SQLSMALLINT, SQLSMALLINT, SQLSMALLINT, SQLULEN, SQLSMALLINT, SQLPOINTER, SQLLEN, SQLLEN *);
    SQLRETURN (SQL_API *NumResultCols)(SQLHSTMT, SQLSMALLINT *);
    SQLRETURN (SQL_API *DescribeColW)(SQLHSTMT, SQLUSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLSMALLINT *, SQLSMALLINT *, SQLULEN *, SQLSMALLINT *, SQLSMALLINT *);
    SQLRETURN (SQL_API *ColAttributeW)(SQLHSTMT, SQLUSMALLINT, SQLUSMALLINT, SQLPOINTER, SQLSMALLINT, SQLSMALLINT *, SQLLEN *);
    SQLRETURN (SQL_API *RowCount)(SQLHSTMT, SQLLEN *);
    SQLRETURN (SQL_API *Fetch)(SQLHSTMT);
    SQLRETURN (SQL_API *GetData)(SQLHSTMT, SQLUSMALLINT, SQLSMALLINT, SQLPOINTER, SQLLEN, SQLLEN *);
    SQLRETURN (SQL_API *MoreResults)(SQLHSTMT);
    SQLRETURN (SQL_API *CloseCursor)(SQLHSTMT);
    SQLRETURN (SQL_API *GetDiagRecW)(SQLSMALLINT, SQLHANDLE, SQLSMALLINT, SQLWCHAR *, SQLINTEGER *, SQLWCHAR *, SQLSMALLINT, SQLSMALLINT *);
    SQLRETURN (SQL_API *TablesW)(SQLHSTMT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT);
    SQLRETURN (SQL_API *ColumnsW)(SQLHSTMT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT);
    SQLRETURN (SQL_API *PrimaryKeysW)(SQLHSTMT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT);
    SQLRETURN (SQL_API *StatisticsW)(SQLHSTMT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLUSMALLINT, SQLUSMALLINT);
    SQLRETURN (SQL_API *DriversW)(SQLHENV, SQLUSMALLINT, SQLWCHAR *, SQLSMALLINT, SQLSMALLINT *, SQLWCHAR *, SQLSMALLINT, SQLSMALLINT *);
    SQLHENV env;
} O;

static void *sym(void *h, const char *n)
{
#ifdef _WIN32
    return (void *)GetProcAddress((HMODULE)h, n);
#else
    return dlsym(h, n);
#endif
}

int lbo_available(void)
{
    if (O.tried) return O.ok;
    O.tried = 1;
    void *h = 0;
#ifdef _WIN32
    h = (void *)LoadLibraryW(L"odbc32.dll");
#else
    const char *e = getenv("VBRT_ODBC_LIB");
    if (e && *e) h = dlopen(e, RTLD_NOW | RTLD_GLOBAL);
    if (!h) h = dlopen("libodbc.so.2", RTLD_NOW | RTLD_GLOBAL);
    if (!h) h = dlopen("libodbc.so", RTLD_NOW | RTLD_GLOBAL);
    if (!h) h = dlopen("libiodbc.so.2", RTLD_NOW | RTLD_GLOBAL);
#endif
    if (!h) return 0;
#define S(f, n) if (!(*(void **)&O.f = sym(h, n))) return 0;
    S(AllocHandle, "SQLAllocHandle") S(FreeHandle, "SQLFreeHandle") S(SetEnvAttr, "SQLSetEnvAttr")
    S(DriverConnectW, "SQLDriverConnectW") S(Disconnect, "SQLDisconnect") S(GetInfoW, "SQLGetInfoW")
    S(SetConnectAttrW, "SQLSetConnectAttrW") S(EndTran, "SQLEndTran") S(PrepareW, "SQLPrepareW")
    S(ExecDirectW, "SQLExecDirectW") S(Execute, "SQLExecute") S(NumParams, "SQLNumParams")
    S(DescribeParam, "SQLDescribeParam") S(BindParameter, "SQLBindParameter") S(NumResultCols, "SQLNumResultCols")
    S(DescribeColW, "SQLDescribeColW") S(ColAttributeW, "SQLColAttributeW") S(RowCount, "SQLRowCount")
    S(Fetch, "SQLFetch") S(GetData, "SQLGetData") S(MoreResults, "SQLMoreResults") S(CloseCursor, "SQLCloseCursor")
    S(GetDiagRecW, "SQLGetDiagRecW") S(TablesW, "SQLTablesW") S(ColumnsW, "SQLColumnsW")
    S(PrimaryKeysW, "SQLPrimaryKeysW") S(DriversW, "SQLDriversW") S(StatisticsW, "SQLStatisticsW")
#undef S
    if (!SQL_SUCCEEDED(O.AllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &O.env))) return 0;
    /* ODBC 2, as MSDASQL (VB6 oracle d_odbc1: the states are S1000 ...) */
    O.SetEnvAttr(O.env, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)(intptr_t)2, 0);
    return O.ok = 1;
}

/* ---------------------------------------------------------------- UTF-8 <-> SQLWCHAR (UTF-16) */
static SQLWCHAR *wz(const char *s)
{
    if (!s) s = "";
    size_t n = strlen(s);
    SQLWCHAR *w = (SQLWCHAR *)malloc(sizeof(SQLWCHAR) * (n * 2 + 2));
    size_t k = 0;
    const unsigned char *p = (const unsigned char *)s;
    while (*p) {
        unsigned c = *p++;
        if (c >= 0xC0) {
            int more = c >= 0xF0 ? 3 : c >= 0xE0 ? 2 : 1;
            c &= 0x3F >> more;
            while (more-- && (*p & 0xC0) == 0x80) c = (c << 6) | (*p++ & 0x3F);
        }
        if (c >= 0x10000) { c -= 0x10000; w[k++] = (SQLWCHAR)(0xD800 + (c >> 10)); w[k++] = (SQLWCHAR)(0xDC00 + (c & 0x3FF)); }
        else w[k++] = (SQLWCHAR)c;
    }
    w[k] = 0;
    return w;
}
/* n: SQLWCHARs, -1: up to the 0 */
static char *uz(const SQLWCHAR *w, long n)
{
    if (n < 0) { n = 0; while (w[n]) n++; }
    char *s = (char *)malloc((size_t)n * 3 + 1);
    size_t k = 0;
    for (long i = 0; i < n; i++) {
        unsigned c = w[i];
        if (c >= 0xD800 && c < 0xDC00 && i + 1 < n) { c = 0x10000 + ((c - 0xD800) << 10) + (w[i + 1] - 0xDC00); i++; }
        if (c < 0x80) s[k++] = (char)c;
        else if (c < 0x800) { s[k++] = (char)(0xC0 | c >> 6); s[k++] = (char)(0x80 | (c & 0x3F)); }
        else if (c < 0x10000) { s[k++] = (char)(0xE0 | c >> 12); s[k++] = (char)(0x80 | (c >> 6 & 0x3F)); s[k++] = (char)(0x80 | (c & 0x3F)); }
        else { s[k++] = (char)(0xF0 | c >> 18); s[k++] = (char)(0x80 | (c >> 12 & 0x3F)); s[k++] = (char)(0x80 | (c >> 6 & 0x3F)); s[k++] = (char)(0x80 | (c & 0x3F)); }
    }
    s[k] = 0;
    return s;
}

/* ---------------------------------------------------------------- errors */
static lbo_err errs[8];
static int nerrs;
int lbo_nerr(void) { return nerrs; }
const lbo_err *lbo_errs(void) { return errs; }

static void diag(SQLSMALLINT type, SQLHANDLE h)
{
    nerrs = 0;
    for (SQLSMALLINT i = 1; i <= 8; i++) {
        SQLWCHAR st[8], msg[1024];
        SQLINTEGER native = 0;
        SQLSMALLINT len = 0;
        if (!SQL_SUCCEEDED(O.GetDiagRecW(type, h, i, st, &native, msg, 1024, &len))) break;
        lbo_err *e = &errs[nerrs++];
        char *s = uz(st, 5), *m = uz(msg, len < 1024 ? len : 1023);
        snprintf(e->state, sizeof e->state, "%s", s);
        snprintf(e->msg, sizeof e->msg, "%s", m);
        e->native = (int)native;
        free(s); free(m);
    }
}
static void fail(const char *state, const char *msg)
{
    nerrs = 1;
    snprintf(errs[0].state, sizeof errs[0].state, "%s", state);
    snprintf(errs[0].msg, sizeof errs[0].msg, "%s", msg);
    errs[0].native = 0;
}

/* ---------------------------------------------------------------- connections */
struct lbo_conn { SQLHDBC dbc; };
struct lbo_stmt {
    SQLHSTMT st;
    lbo_conn *c;
    struct pbuf { SQLLEN ind; void *data; } **pb; int npb;  /* the values bound to the parameters: each
                                                                where the driver keeps its address */
};

lbo_conn *lbo_connect(const char *cs, char *outcs, int outlen)
{
    nerrs = 0;
    if (!lbo_available()) {
#ifdef _WIN32
        fail("IM002", "[Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified");
#else
        fail("IM002", "ODBC is not installed (unixODBC: sudo apt install unixodbc)");
#endif
        return 0;
    }
    SQLHDBC dbc;
    if (!SQL_SUCCEEDED(O.AllocHandle(SQL_HANDLE_DBC, O.env, &dbc))) { diag(SQL_HANDLE_ENV, O.env); return 0; }
    SQLWCHAR *w = wz(cs), out[2048];
    SQLSMALLINT olen = 0;
    SQLRETURN r = O.DriverConnectW(dbc, 0, w, SQL_NTS, out, 2048, &olen, SQL_DRIVER_NOPROMPT);
    free(w);
    if (!SQL_SUCCEEDED(r)) { diag(SQL_HANDLE_DBC, dbc); O.FreeHandle(SQL_HANDLE_DBC, dbc); return 0; }
    if (outcs && outlen > 0) { char *o = uz(out, olen < 2048 ? olen : 2047); snprintf(outcs, (size_t)outlen, "%s", o); free(o); }
    lbo_conn *c = (lbo_conn *)calloc(1, sizeof *c);
    c->dbc = dbc;
    return c;
}
void lbo_disconnect(lbo_conn *c)
{
    if (!c) return;
    O.Disconnect(c->dbc);
    O.FreeHandle(SQL_HANDLE_DBC, c->dbc);
    free(c);
}
int lbo_info(lbo_conn *c, int what, char *buf, int len)
{
    SQLWCHAR w[512]; SQLSMALLINT n = 0;
    buf[0] = 0;
    if (!SQL_SUCCEEDED(O.GetInfoW(c->dbc, (SQLUSMALLINT)what, w, sizeof w, &n))) return -1;
    char *s = uz(w, n / (SQLSMALLINT)sizeof(SQLWCHAR));
    snprintf(buf, (size_t)len, "%s", s);
    free(s);
    return 0;
}
int lbo_info_int(lbo_conn *c, int what, long *out)
{
    union { SQLUINTEGER u; SQLUSMALLINT s; } v; SQLSMALLINT n = 0;
    v.u = 0;
    if (!SQL_SUCCEEDED(O.GetInfoW(c->dbc, (SQLUSMALLINT)what, &v, sizeof v, &n))) return -1;
    *out = n == 2 ? (long)v.s : (long)v.u;
    return 0;
}
int lbo_autocommit(lbo_conn *c, int on)
{
    SQLRETURN r = O.SetConnectAttrW(c->dbc, SQL_ATTR_AUTOCOMMIT, (SQLPOINTER)(intptr_t)(on ? SQL_AUTOCOMMIT_ON : SQL_AUTOCOMMIT_OFF), 0);
    if (!SQL_SUCCEEDED(r)) { diag(SQL_HANDLE_DBC, c->dbc); return -1; }
    return 0;
}
int lbo_endtran(lbo_conn *c, int commit)
{
    SQLRETURN r = O.EndTran(SQL_HANDLE_DBC, c->dbc, commit ? SQL_COMMIT : SQL_ROLLBACK);
    if (!SQL_SUCCEEDED(r)) { diag(SQL_HANDLE_DBC, c->dbc); return -1; }
    return 0;
}

/* the names of the drivers installed (one per line) */
int lbo_drivers(char *buf, int len)
{
    buf[0] = 0;
    if (!lbo_available()) return -1;
    SQLWCHAR d[256], a[1024];
    SQLSMALLINT dl, al;
    SQLUSMALLINT dir = SQL_FETCH_FIRST;
    size_t k = 0;
    while (SQL_SUCCEEDED(O.DriversW(O.env, dir, d, 256, &dl, a, 1024, &al))) {
        char *s = uz(d, -1);
        size_t n = strlen(s);
        if (k + n + 2 < (size_t)len) { memcpy(buf + k, s, n); k += n; buf[k++] = '\n'; buf[k] = 0; }
        free(s);
        dir = SQL_FETCH_NEXT;
    }
    return 0;
}

/* ---------------------------------------------------------------- statements */
lbo_stmt *lbo_stmt_new(lbo_conn *c)
{
    SQLHSTMT st;
    if (!SQL_SUCCEEDED(O.AllocHandle(SQL_HANDLE_STMT, c->dbc, &st))) { diag(SQL_HANDLE_DBC, c->dbc); return 0; }
    lbo_stmt *s = (lbo_stmt *)calloc(1, sizeof *s);
    s->st = st; s->c = c;
    return s;
}
static void free_params(lbo_stmt *s)
{
    for (int i = 0; i < s->npb; i++) { free(s->pb[i]->data); free(s->pb[i]); }
    free(s->pb); s->pb = 0; s->npb = 0;
}
void lbo_stmt_free(lbo_stmt *s)
{
    if (!s) return;
    O.FreeHandle(SQL_HANDLE_STMT, s->st);
    free_params(s);
    free(s);
}
static int check(lbo_stmt *s, SQLRETURN r)
{
    if (r == SQL_NO_DATA) return 0;
    if (!SQL_SUCCEEDED(r)) { diag(SQL_HANDLE_STMT, s->st); return -1; }
    if (r == SQL_SUCCESS_WITH_INFO) nerrs = 0;
    return 0;
}
int lbo_prepare(lbo_stmt *s, const char *sql)
{
    nerrs = 0;
    SQLWCHAR *w = wz(sql);
    SQLRETURN r = O.PrepareW(s->st, w, SQL_NTS);
    free(w);
    return check(s, r);
}
int lbo_exec_direct(lbo_stmt *s, const char *sql)
{
    nerrs = 0;
    SQLWCHAR *w = wz(sql);
    SQLRETURN r = O.ExecDirectW(s->st, w, SQL_NTS);
    free(w);
    return check(s, r);
}
int lbo_execute(lbo_stmt *s) { nerrs = 0; return check(s, O.Execute(s->st)); }
int lbo_nparams(lbo_stmt *s) { SQLSMALLINT n = 0; if (!SQL_SUCCEEDED(O.NumParams(s->st, &n))) return -1; return n; }
int lbo_describe_param(lbo_stmt *s, int i, int *sqltype, unsigned long *size, int *dec, int *nullable)
{
    SQLSMALLINT t = 0, d = 0, nl = 0; SQLULEN z = 0;
    if (!SQL_SUCCEEDED(O.DescribeParam(s->st, (SQLUSMALLINT)i, &t, &z, &d, &nl))) return -1;
    *sqltype = t; *size = (unsigned long)z; *dec = d; *nullable = nl;
    return 0;
}

/* a value for parameter i (1-based) */
int lbo_bind(lbo_stmt *s, int i, int kind, long long iv, double dv, const char *text, int len, int sqltype)
{
    if (i > s->npb) {
        s->pb = (struct pbuf **)realloc(s->pb, sizeof *s->pb * (size_t)i);
        for (int k = s->npb; k < i; k++) s->pb[k] = (struct pbuf *)calloc(1, sizeof **s->pb);
        s->npb = i;
    }
    struct pbuf *b = s->pb[i - 1];
    free(b->data); b->data = 0;
    SQLSMALLINT ctype, stype;
    SQLULEN col = 0;
    switch (kind) {
    case LBO_NULL:
        b->ind = SQL_NULL_DATA;
        return check(s, O.BindParameter(s->st, (SQLUSMALLINT)i, SQL_PARAM_INPUT, SQL_C_CHAR, sqltype ? (SQLSMALLINT)sqltype : SQL_VARCHAR, 1, 0, 0, 0, &b->ind));
    case LBO_INT: case LBO_BOOL:
        b->data = malloc(sizeof(long long)); *(long long *)b->data = iv; b->ind = 0;
        ctype = SQL_C_SBIGINT; stype = kind == LBO_BOOL ? SQL_BIT : sqltype ? (SQLSMALLINT)sqltype : SQL_INTEGER;
        if (stype == SQL_INTEGER && (iv > 2147483647LL || iv < -2147483648LL)) stype = SQL_BIGINT;
        break;
    case LBO_DBL:
        b->data = malloc(sizeof(double)); *(double *)b->data = dv; b->ind = 0;
        ctype = SQL_C_DOUBLE; stype = SQL_DOUBLE;
        break;
    case LBO_BIN:
        b->data = malloc((size_t)len + 1); memcpy(b->data, text, (size_t)len); b->ind = len;
        ctype = SQL_C_BINARY; stype = SQL_LONGVARBINARY; col = (SQLULEN)(len ? len : 1);
        break;
    default: {                                  /* text, a date as text, a decimal as text */
        SQLWCHAR *w = wz(text);
        long n = 0; while (w[n]) n++;
        b->data = w; b->ind = (SQLLEN)(n * (long)sizeof(SQLWCHAR));
        ctype = SQL_C_WCHAR;
        stype = kind == LBO_DATE ? 11 /* SQL_TIMESTAMP */ : kind == LBO_DEC ? SQL_NUMERIC : SQL_WVARCHAR;
        col = (SQLULEN)(n ? n : 1);
        if (kind == LBO_DATE) {                 /* a date, a time or both: as the column has it */
            int dt = sqltype == 9 || sqltype == 91, tm = sqltype == 10 || sqltype == 92;
            stype = dt ? 9 : tm ? 10 : 11;
            col = dt ? 10 : tm ? 8 : 19;
        }
        if (kind == LBO_DEC) col = 38;
        if (kind == LBO_TEXT && sqltype && sqltype != SQL_WVARCHAR) stype = (SQLSMALLINT)sqltype;
        break;
    }
    }
    return check(s, O.BindParameter(s->st, (SQLUSMALLINT)i, SQL_PARAM_INPUT, ctype, stype, col,
                                    (SQLSMALLINT)(kind == LBO_DEC ? 10 : 0), b->data, b->ind > 0 ? b->ind : 0, &b->ind));
}

int lbo_ncols(lbo_stmt *s) { SQLSMALLINT n = 0; if (!SQL_SUCCEEDED(O.NumResultCols(s->st, &n))) { diag(SQL_HANDLE_STMT, s->st); return -1; } return n; }
long lbo_rowcount(lbo_stmt *s) { SQLLEN n = -1; if (!SQL_SUCCEEDED(O.RowCount(s->st, &n))) return -1; return (long)n; }

static void attr_str(lbo_stmt *s, int i, int what, char *out, int len)
{
    SQLWCHAR w[512]; SQLSMALLINT n = 0;
    out[0] = 0;
    if (!SQL_SUCCEEDED(O.ColAttributeW(s->st, (SQLUSMALLINT)i, (SQLUSMALLINT)what, w, sizeof w, &n, 0))) return;
    char *u = uz(w, n / (SQLSMALLINT)sizeof(SQLWCHAR));
    snprintf(out, (size_t)len, "%s", u);
    free(u);
}
static long attr_num(lbo_stmt *s, int i, int what)
{
    SQLLEN v = 0;
    if (!SQL_SUCCEEDED(O.ColAttributeW(s->st, (SQLUSMALLINT)i, (SQLUSMALLINT)what, 0, 0, 0, &v))) return -1;
    return (long)v;
}
int lbo_colinfo_get(lbo_stmt *s, int i, lbo_colinfo *ci)
{
    memset(ci, 0, sizeof *ci);
    SQLWCHAR w[512]; SQLSMALLINT nl = 0, t = 0, d = 0, nu = 0; SQLULEN z = 0;
    if (!SQL_SUCCEEDED(O.DescribeColW(s->st, (SQLUSMALLINT)i, w, 512, &nl, &t, &z, &d, &nu))) { diag(SQL_HANDLE_STMT, s->st); return -1; }
    char *u = uz(w, nl < 512 ? nl : 511);
    snprintf(ci->name, sizeof ci->name, "%s", u);
    free(u);
    ci->sqltype = t; ci->size = (unsigned long)z; ci->decimals = d; ci->nullable = nu;
    attr_str(s, i, SQL_DESC_BASE_TABLE_NAME, ci->basetable, sizeof ci->basetable);
    if (!ci->basetable[0]) attr_str(s, i, SQL_DESC_TABLE_NAME, ci->basetable, sizeof ci->basetable);
    attr_str(s, i, SQL_DESC_BASE_COLUMN_NAME, ci->basecol, sizeof ci->basecol);
    ci->autoinc = attr_num(s, i, SQL_DESC_AUTO_UNIQUE_VALUE) == SQL_TRUE;
    ci->unsignd = attr_num(s, i, SQL_DESC_UNSIGNED) == SQL_TRUE;
    long up = attr_num(s, i, SQL_DESC_UPDATABLE);
    ci->updatable = up < 0 ? 1 : up != SQL_ATTR_READONLY;
    long oct = attr_num(s, i, SQL_DESC_OCTET_LENGTH);
    ci->octets = oct > 0 ? (unsigned long)oct : 0;
    return 0;
}

int lbo_fetch(lbo_stmt *s)
{
    SQLRETURN r = O.Fetch(s->st);
    if (r == SQL_NO_DATA) return 0;
    if (!SQL_SUCCEEDED(r)) { diag(SQL_HANDLE_STMT, s->st); return -1; }
    return 1;
}
int lbo_more(lbo_stmt *s) { return O.MoreResults(s->st) == SQL_SUCCESS; }
void lbo_close_cursor(lbo_stmt *s) { O.CloseCursor(s->st); }

static int get_long(lbo_stmt *s, int i, SQLSMALLINT ctype, lbo_val *v)
{
    size_t cap = 1024, n = 0;
    char *buf = (char *)malloc(cap);
    int unit = ctype == SQL_C_WCHAR ? (int)sizeof(SQLWCHAR) : 1;
    for (;;) {
        SQLLEN ind = 0;
        SQLLEN room = (SQLLEN)(cap - n);
        SQLRETURN r = O.GetData(s->st, (SQLUSMALLINT)i, ctype, buf + n, room, &ind);
        if (r == SQL_NO_DATA) break;
        if (!SQL_SUCCEEDED(r)) { diag(SQL_HANDLE_STMT, s->st); free(buf); return -1; }
        if (ind == SQL_NULL_DATA) { free(buf); v->kind = LBO_NULL; return 0; }
        size_t got;
        if (r == SQL_SUCCESS_WITH_INFO && (ind == SQL_NO_TOTAL || ind >= room - (ctype == SQL_C_WCHAR ? unit : 0))) {
            got = (size_t)room - (ctype == SQL_C_WCHAR ? (size_t)unit : 0);
            got -= got % (size_t)unit;
            n += got;
            cap = cap * 2 + (ind > 0 ? (size_t)ind : 0);
            buf = (char *)realloc(buf, cap);
            continue;
        }
        n += (size_t)ind;
        break;
    }
    if (ctype == SQL_C_WCHAR) {
        v->s = uz((SQLWCHAR *)buf, (long)(n / (size_t)unit));
        v->len = (int)strlen(v->s);
        free(buf);
    } else { v->s = buf; v->len = (int)n; }
    return 0;
}

/* the value of column i (1-based) of the current row, by its SQL type */
int lbo_get(lbo_stmt *s, int i, int sqltype, lbo_val *v)
{
    memset(v, 0, sizeof *v);
    SQLLEN ind = 0;
    switch (sqltype) {
    case SQL_TINYINT: case SQL_SMALLINT: case SQL_INTEGER: case SQL_BIGINT: case SQL_BIT: {
        long long x = 0;
        SQLRETURN r = O.GetData(s->st, (SQLUSMALLINT)i, SQL_C_SBIGINT, &x, 0, &ind);
        if (!SQL_SUCCEEDED(r)) { diag(SQL_HANDLE_STMT, s->st); return -1; }
        if (ind == SQL_NULL_DATA) { v->kind = LBO_NULL; return 0; }
        v->kind = sqltype == SQL_BIT ? LBO_BOOL : LBO_INT; v->iv = x;
        return 0;
    }
    case SQL_REAL: case SQL_FLOAT: case SQL_DOUBLE: {
        double x = 0;
        SQLRETURN r = O.GetData(s->st, (SQLUSMALLINT)i, SQL_C_DOUBLE, &x, 0, &ind);
        if (!SQL_SUCCEEDED(r)) { diag(SQL_HANDLE_STMT, s->st); return -1; }
        if (ind == SQL_NULL_DATA) { v->kind = LBO_NULL; return 0; }
        v->kind = LBO_DBL; v->dv = x;
        return 0;
    }
    case SQL_TYPE_DATE: case SQL_TYPE_TIME: case SQL_TYPE_TIMESTAMP: case SQL_DATE: case SQL_TIME: case SQL_TIMESTAMP: {
        SQL_TIMESTAMP_STRUCT ts; memset(&ts, 0, sizeof ts);
        SQLRETURN r = O.GetData(s->st, (SQLUSMALLINT)i, 11 /* SQL_C_TIMESTAMP (ODBC 2) */, &ts, sizeof ts, &ind);
        if (!SQL_SUCCEEDED(r)) { diag(SQL_HANDLE_STMT, s->st); return -1; }
        if (ind == SQL_NULL_DATA) { v->kind = LBO_NULL; return 0; }
        char b[64];
        snprintf(b, sizeof b, "%04d-%02u-%02u %02u:%02u:%02u", ts.year, ts.month, ts.day, ts.hour, ts.minute, ts.second);
        v->kind = LBO_DATE; v->s = strdup(b); v->len = (int)strlen(b);
        return 0;
    }
    case SQL_BINARY: case SQL_VARBINARY: case SQL_LONGVARBINARY:
        v->kind = LBO_BIN;
        return get_long(s, i, SQL_C_BINARY, v);
    case SQL_NUMERIC: case SQL_DECIMAL:
        v->kind = LBO_DEC;
        return get_long(s, i, SQL_C_WCHAR, v);
    default:
        v->kind = LBO_TEXT;
        return get_long(s, i, SQL_C_WCHAR, v);
    }
}
void lbo_val_free(lbo_val *v) { free(v->s); v->s = 0; }

/* ---------------------------------------------------------------- the catalog */
static SQLWCHAR *wn(const char *s) { return s ? wz(s) : 0; }
int lbo_tables(lbo_stmt *s, const char *cat, const char *sch, const char *tab, const char *types)
{
    SQLWCHAR *a = wn(cat), *b = wn(sch), *c = wn(tab), *d = wn(types);
    SQLRETURN r = O.TablesW(s->st, a, a ? SQL_NTS : 0, b, b ? SQL_NTS : 0, c, c ? SQL_NTS : 0, d, d ? SQL_NTS : 0);
    free(a); free(b); free(c); free(d);
    return check(s, r);
}
int lbo_columns(lbo_stmt *s, const char *cat, const char *sch, const char *tab, const char *col)
{
    SQLWCHAR *a = wn(cat), *b = wn(sch), *c = wn(tab), *d = wn(col);
    SQLRETURN r = O.ColumnsW(s->st, a, a ? SQL_NTS : 0, b, b ? SQL_NTS : 0, c, c ? SQL_NTS : 0, d, d ? SQL_NTS : 0);
    free(a); free(b); free(c); free(d);
    return check(s, r);
}
int lbo_primary_keys(lbo_stmt *s, const char *cat, const char *sch, const char *tab)
{
    SQLWCHAR *a = wn(cat), *b = wn(sch), *c = wn(tab);
    SQLRETURN r = O.PrimaryKeysW(s->st, a, a ? SQL_NTS : 0, b, b ? SQL_NTS : 0, c, c ? SQL_NTS : 0);
    free(a); free(b); free(c);
    return check(s, r);
}
/* the indexes of a table (SQLStatistics; unique_only: SQL_INDEX_UNIQUE) */
int lbo_statistics(lbo_stmt *s, const char *cat, const char *sch, const char *tab, int unique_only)
{
    SQLWCHAR *a = wn(cat), *b = wn(sch), *c = wn(tab);
    SQLRETURN r = O.StatisticsW(s->st, a, a ? SQL_NTS : 0, b, b ? SQL_NTS : 0, c, c ? SQL_NTS : 0,
                                (SQLUSMALLINT)(unique_only ? 0 : 1), (SQLUSMALLINT)1);   /* SQL_INDEX_UNIQUE / ALL, SQL_ENSURE */
    free(a); free(b); free(c);
    return check(s, r);
}
