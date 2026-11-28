/* jetsql.c - Jet SQL (the SQL of Access and DAO, ANSI-89 mode) written as SQLite SQL, and the VBA
 * functions, the LIKE and the text order Jet SQL has, on a SQLite connection.
 *
 * Jet has its own rules where SQLite differs: True is -1, a comparison gives -1/0, & joins text and
 * takes Null as "", / always divides as Double, \ Mod ^ are operators, #m/d/y# is a date, "text" is
 * a string, [name] a name, LIKE uses * ? # [..] and ignores case, text compares without case. So the
 * statement is parsed (expressions with Jet's precedence) and written again for SQLite; the VBA
 * functions are SQLite functions named jet_<name> that run the runtime's own VBA code.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbdb.h"
#include "jetfile.h"
#include <ctype.h>
#include <wctype.h>
#include <stdint.h>

/* ---------------------------------------------------------------- jbuf */
void jb_put(jbuf *b, const char *s, int n)
{
    if (n < 0) n = (int)strlen(s);
    if (b->n + n + 1 > b->cap) {
        b->cap = (b->n + n + 1) * 2 + 32;
        b->p = realloc(b->p, (size_t)b->cap);
    }
    memcpy(b->p + b->n, s, (size_t)n);
    b->n += n;
    b->p[b->n] = 0;
}
void jb_s(jbuf *b, const char *s) { jb_put(b, s, -1); }
void jb_f(jbuf *b, const char *fmt, ...)
{
    char tmp[512];
    va_list ap; va_start(ap, fmt);
    int n = vsnprintf(tmp, sizeof tmp, fmt, ap);
    va_end(ap);
    if (n < (int)sizeof tmp) { jb_put(b, tmp, n); return; }
    char *big = malloc((size_t)n + 1);
    va_start(ap, fmt); vsnprintf(big, (size_t)n + 1, fmt, ap); va_end(ap);
    jb_put(b, big, n); free(big);
}
void jb_qid(jbuf *b, const char *s, int n)
{
    if (n < 0) n = (int)strlen(s);
    jb_put(b, "\"", 1);
    for (int i = 0; i < n; i++) { if (s[i] == '"') jb_put(b, "\"", 1); jb_put(b, s + i, 1); }
    jb_put(b, "\"", 1);
}
void jb_qstr(jbuf *b, const char *s, int n)
{
    if (n < 0) n = (int)strlen(s);
    jb_put(b, "'", 1);
    for (int i = 0; i < n; i++) { if (s[i] == '\'') jb_put(b, "'", 1); jb_put(b, s + i, 1); }
    jb_put(b, "'", 1);
}
void jb_free(jbuf *b) { free(b->p); b->p = 0; b->n = b->cap = 0; }
static char *jb_take(jbuf *b) { char *p = b->p ? b->p : strdup(""); b->p = 0; b->n = b->cap = 0; return p; }
static char *dupn(const char *s, int n) { char *p = malloc((size_t)n + 1); memcpy(p, s, (size_t)n); p[n] = 0; return p; }

/* ---------------------------------------------------------------- dates */
void jet_date_text(vb_date d, char out[24])
{
    double days = floor(d), t = fabs(d - (double)(long long)d);
    if (d < 0) days = ceil(d);                  /* VB dates before 1899: the day counts down, the time up */
    vb_i32 y = vb_year(d), m = vb_month(d), dd = vb_day(d);
    long secs = (long)floor(t * 86400.0 + 0.5);
    if (secs >= 86400) secs = 86399;
    (void)days;
    snprintf(out, 24, "%c%04d-%02d-%02d %02ld:%02ld:%02ld", JDATE_MARK, (int)y, (int)m, (int)dd,
             secs / 3600, secs / 60 % 60, secs % 60);
}
int jet_text_date(const char *s, int n, vb_date *d)
{
    if (n < JDATE_LEN || s[0] != JDATE_MARK) return 0;
    int y, mo, dd, h, mi, se;
    if (sscanf(s + 1, "%d-%d-%d %d:%d:%d", &y, &mo, &dd, &h, &mi, &se) != 6) return 0;
    vb_date day = vb_dateserial(y, mo, dd);
    double t = (h * 3600 + mi * 60 + se) / 86400.0;
    *d = day >= 0 ? day + t : day - t;
    return 1;
}

/* ---------------------------------------------------------------- tokens */
enum { T_END, T_ID, T_QID, T_NUM, T_STR, T_DATE, T_OP, T_QMARK };
typedef struct tok { int k; const char *s; int n; } tok;
typedef struct ps {
    tok *t; int nt, i;
    jsql *q;
    const char *const *params;
    int err; char msg[400];
    const char *src;
    int exprno;                         /* Expr1000, Expr1001 ... */
    int ansi_like;                      /* ADO: % and _ */
    int nqmark;                         /* the ?s so far */
    int from_ext;                       /* the FROM being parsed has IN 'file': its schema number */
} ps;

static int from_tables;                 /* the tables of the FROM being parsed (for DISTINCTROW) */
static int is_idc(int c, int first)
{
    return isalpha(c) || c == '_' || c >= 0x80 || (!first && isdigit(c));
}
static int tokenize(const char *s, tok **out, char *err, int cap)
{
    int n = 0, cap2 = 64;
    tok *t = malloc(sizeof(tok) * (size_t)cap2);
    const char *p = s;
    for (;;) {
        while (*p && isspace((unsigned char)*p)) p++;
        if (n + 1 >= cap2) { cap2 *= 2; t = realloc(t, sizeof(tok) * (size_t)cap2); }
        tok *k = &t[n];
        if (!*p) { k->k = T_END; k->s = p; k->n = 0; break; }
        unsigned char c = (unsigned char)*p;
        if (is_idc(c, 1)) {
            const char *b = p;
            while (*p && is_idc((unsigned char)*p, 0)) p++;
            k->k = T_ID; k->s = b; k->n = (int)(p - b);
        } else if (c == '[') {
            const char *b = ++p;
            while (*p && *p != ']') p++;
            if (!*p) { snprintf(err, (size_t)cap, "["); free(t); return -1; }
            k->k = T_QID; k->s = b; k->n = (int)(p - b); p++;
        } else if (c == '`') {
            const char *b = ++p;
            while (*p && *p != '`') p++;
            if (!*p) { snprintf(err, (size_t)cap, "`"); free(t); return -1; }
            k->k = T_QID; k->s = b; k->n = (int)(p - b); p++;
        } else if (isdigit(c) || (c == '.' && isdigit((unsigned char)p[1]))) {
            const char *b = p;
            while (isdigit((unsigned char)*p)) p++;
            if (*p == '.') { p++; while (isdigit((unsigned char)*p)) p++; }
            if ((*p == 'e' || *p == 'E') && (isdigit((unsigned char)p[1]) || ((p[1] == '+' || p[1] == '-') && isdigit((unsigned char)p[2])))) {
                p += 2; while (isdigit((unsigned char)*p)) p++;
            }
            k->k = T_NUM; k->s = b; k->n = (int)(p - b);
        } else if (c == '\'' || c == '"') {
            const char *b = p++;
            for (;;) {
                if (!*p) { snprintf(err, (size_t)cap, "string"); free(t); return -1; }
                if (*p == (char)c) { if (p[1] == (char)c) { p += 2; continue; } p++; break; }
                p++;
            }
            k->k = T_STR; k->s = b; k->n = (int)(p - b);   /* with the quotes */
        } else if (c == '#') {
            const char *b = ++p;
            while (*p && *p != '#') p++;
            if (!*p) { snprintf(err, (size_t)cap, "#"); free(t); return -1; }
            k->k = T_DATE; k->s = b; k->n = (int)(p - b); p++;
        } else if (c == '?') {
            k->k = T_QMARK; k->s = p++; k->n = 1;
        } else {
            k->k = T_OP; k->s = p;
            if ((c == '<' && (p[1] == '>' || p[1] == '=')) || (c == '>' && p[1] == '=') || (c == '!' && p[1] == '=')) k->n = 2;
            else k->n = 1;
            p += k->n;
        }
        n++;
    }
    *out = t;
    return n;
}

static tok *cur(ps *p) { return &p->t[p->i]; }
static int kw(tok *k, const char *w)
{
    return k->k == T_ID && (int)strlen(w) == k->n && !strncasecmp(k->s, w, (size_t)k->n);
}
static int iskw(ps *p, const char *w) { return kw(cur(p), w); }
static int iskw2(ps *p, const char *w) { return p->i + 1 < p->nt && kw(&p->t[p->i + 1], w); }
static int accept(ps *p, const char *w) { if (iskw(p, w)) { p->i++; return 1; } return 0; }
static int isop(ps *p, const char *o) { tok *k = cur(p); return k->k == T_OP && (int)strlen(o) == k->n && !strncmp(k->s, o, (size_t)k->n); }
static int acceptop(ps *p, const char *o) { if (isop(p, o)) { p->i++; return 1; } return 0; }

static const char *const RESERVED[] = {
    "select", "from", "where", "group", "by", "having", "order", "union", "and", "or", "not", "xor", "eqv", "imp",
    "like", "is", "between", "in", "as", "on", "join", "inner", "left", "right", "outer", "into", "values",
    "set", "asc", "desc", "mod", "distinct", "distinctrow", "top", "all", "exists", "with", 0 };
static int reserved(tok *k)
{
    if (k->k != T_ID) return 0;
    for (int i = 0; RESERVED[i]; i++) if (kw(k, RESERVED[i])) return 1;
    return 0;
}

static void fail(ps *p, int err, const char *fmt, ...)
{
    if (p->err) return;
    p->err = err;
    va_list ap; va_start(ap, fmt);
    vsnprintf(p->msg, sizeof p->msg, fmt, ap);
    va_end(ap);
}
/* the text of the expression that has the syntax error, as Jet quotes it */
static void fail_expr_kind(ps *p, int from, int missing);
static void fail_expr(ps *p, int from) { fail_expr_kind(p, from, 0); }
static void fail_expr_kind(ps *p, int from, int missing)
{
    const char *b = p->t[from].s, *e = b;
    int i = from;
    while (i < p->nt && p->t[i].k != T_END && !(p->t[i].k == T_OP && (p->t[i].s[0] == ',' || p->t[i].s[0] == ';'))) i++;
    e = i < p->nt ? p->t[i].s : b + strlen(b);
    if (i > from) { tok *l = &p->t[i - 1]; e = l->s + l->n + (l->k == T_QID || l->k == T_DATE ? 1 : 0); }
    while (e > b && isspace((unsigned char)e[-1])) e--;
    if (missing) fail(p, 3075, "Missing ), ], or Item in query expression '%.*s'.", (int)(e - b), b);
    else fail(p, 3075, "Syntax error (missing operator) in query expression '%.*s'.", (int)(e - b), b);
}

/* ---------------------------------------------------------------- expressions
 * Each part is written to its own buffer; boolean parts (comparisons, And/Or/Not) are kept as
 * SQLite conditions and become -1/0 values only where a value is needed. */
typedef struct ex { jbuf b; int boolean; int simple_col; int type; } ex;

/* the type of a column in an expression: the caller's catalog (dao.c) */
int (*jet_coltype_fn)(void *ctx, const char *qual, const char *col);
void *jet_coltype_ctx;
int (*jet_ext_fn)(void *ctx, const char *path, const char *connect, int *err, char *msg, int cap);
int (*jet_link_fn)(void *ctx, const char *name, char *src, int cap, int *err, char *msg, int mcap);
void *jet_ext_ctx;

/* Jet's type of an arithmetic result */
static int rank(int t)
{
    switch (t) {
    case dbBoolean: case dbByte: return 1;
    case dbInteger: return 2;
    case dbLong: return 3;
    case dbSingle: return 5;
    case dbDouble: return 6;
    case dbCurrency: return 4;
    case dbDecimal: return 8;
    }
    return 0;
}
static int arith_type(int a, int b, int op)
{
    if (a == dbBinary || b == dbBinary) return dbBinary;      /* Null */
    if (op == '/') return a == dbDecimal || b == dbDecimal ? dbDecimal : dbDouble;
    if (op == '^') return dbDouble;
    if (op == '\\' || op == '%') return dbLong;
    if (op == '+' && (a == dbText || a == dbMemo) && (b == dbText || b == dbMemo)) return dbText;
    if ((op == '+' || op == '-') && (a == dbDate || b == dbDate)) return a == dbDate && b == dbDate && op == '-' ? dbDouble : dbDate;
    if (!a || !b) return 0;
    if (a == dbDecimal || b == dbDecimal) return dbDecimal;
    if (a == dbCurrency || b == dbCurrency) return rank(a) && rank(b) ? dbCurrency : 0;
    int ra = rank(a), rb = rank(b);
    if (!ra || !rb) return 0;
    int t = ra >= rb ? a : b;
    if (t == dbBoolean || t == dbByte) t = dbInteger;
    return t;
}
/* the result type of a VBA function of Jet SQL; arg: the type of its first argument */
static int func_type(const char *f, int arg)
{
    static const struct { const char *n; int t; } T[] = {
        { "ucase", dbText }, { "lcase", dbText }, { "left", dbText }, { "right", dbText }, { "mid", dbText },
        { "trim", dbText }, { "ltrim", dbText }, { "rtrim", dbText }, { "space", dbText }, { "string", dbText },
        { "chr", dbText }, { "str", dbText }, { "cstr", dbText }, { "format", dbText }, { "hex", dbText }, { "oct", dbText },
        { "replace", dbText }, { "strreverse", dbText }, { "strconv", dbText }, { "monthname", dbText },
        { "weekdayname", dbText }, { "formatnumber", dbText }, { "formatcurrency", dbText }, { "formatpercent", dbText },
        { "formatdatetime", dbText }, { "error", dbText },
        { "len", dbLong }, { "instr", dbLong }, { "instrrev", dbLong }, { "datediff", dbLong }, { "strcomp", dbLong },
        { "clng", dbLong },
        { "year", dbInteger }, { "month", dbInteger }, { "day", dbInteger }, { "hour", dbInteger },
        { "minute", dbInteger }, { "second", dbInteger }, { "weekday", dbInteger }, { "datepart", dbInteger },
        { "asc", dbInteger }, { "cint", dbInteger }, { "sgn", dbInteger },
        { "dateadd", dbDate }, { "dateserial", dbDate }, { "timeserial", dbDate }, { "datevalue", dbDate },
        { "timevalue", dbDate }, { "now", dbDate }, { "date", dbDate }, { "time", dbDate }, { "cdate", dbDate }, { "cvdate", dbDate },
        { "sqr", dbDouble }, { "exp", dbDouble }, { "log", dbDouble }, { "sin", dbDouble }, { "cos", dbDouble },
        { "tan", dbDouble }, { "atn", dbDouble }, { "val", dbDouble }, { "cdbl", dbDouble },
        { "csng", dbSingle }, { "rnd", dbSingle }, { "timer", dbSingle }, { "ccur", dbCurrency }, { "cbyte", dbByte },
        { "cbool", dbBoolean }, { "isnull", dbInteger }, { "isnumeric", dbInteger }, { "isdate", dbInteger },
        { "isempty", dbInteger },
    };
    for (size_t i = 0; i < sizeof T / sizeof *T; i++) if (!strcmp(T[i].n, f)) return T[i].t;
    if (!strcmp(f, "int") || !strcmp(f, "fix") || !strcmp(f, "abs") || !strcmp(f, "round") || !strcmp(f, "cvar"))
        return arg == dbBoolean || arg == dbByte ? dbInteger : arg;
    return 0;
}

static void ex_value(ex *e, jbuf *out)        /* as a value: a condition gives -1/0/Null */
{
    if (e->boolean) { jb_s(out, "(-("); jb_put(out, e->b.p, e->b.n); jb_s(out, "))"); }
    else jb_put(out, e->b.p, e->b.n);
}
static void ex_cond(ex *e, jbuf *out) { jb_put(out, e->b.p, e->b.n); }
static void ex_free(ex *e) { jb_free(&e->b); }

static int parse_expr(ps *p, int minbp, ex *out);
static int parse_select(ps *p, jbuf *out, int top);

static int param_index(ps *p, const char *name, int n)
{
    jsql *q = p->q;
    for (int i = 0; i < q->nbinds; i++)
        if ((int)strlen(q->binds[i]) == n && !strncasecmp(q->binds[i], name, (size_t)n)) return i + 1;
    q->binds = realloc(q->binds, sizeof(char *) * (size_t)(q->nbinds + 1));
    q->binds[q->nbinds++] = dupn(name, n);
    return q->nbinds;
}
static int is_param(ps *p, const char *name, int n)
{
    for (int i = 0; i < p->q->nparams; i++)
        if ((int)strlen(p->q->params[i].name) == n && !strncasecmp(p->q->params[i].name, name, (size_t)n)) return 1;
    if (p->params)
        for (int i = 0; p->params[i]; i++)
            if ((int)strlen(p->params[i]) == n && !strncasecmp(p->params[i], name, (size_t)n)) return 1;
    return 0;
}

static void put_str_literal(jbuf *b, tok *k)
{
    char q = k->s[0];
    jb_put(b, "'", 1);
    for (int i = 1; i < k->n - 1; i++) {
        if (k->s[i] == q && k->s[i + 1] == q) { if (q == '\'') jb_put(b, "''", 2); else jb_put(b, "\"", 1); i++; continue; }
        if (k->s[i] == '\'') jb_put(b, "''", 2);
        else jb_put(b, k->s + i, 1);
    }
    jb_put(b, "'", 1);
}
static int put_date_literal(ps *p, jbuf *b, tok *k)
{
    char tmp[128];
    snprintf(tmp, sizeof tmp, "%.*s", k->n, k->s);
    vb_i32 m = vb_mark();
    /* Jet reads #...# as month/day/year (US), like a VB date literal */
    int a = 0, bb = 0, c = 0, h = 0, mi = 0, se = 0; char ap[4] = "";
    vb_date d = 0; int ok = 0;
    const char *s = tmp;
    while (*s == ' ') s++;
    int n1 = sscanf(s, "%d/%d/%d", &a, &bb, &c), n2 = 0;
    if (n1 != 3) n1 = sscanf(s, "%d-%d-%d", &a, &bb, &c);
    if (n1 == 3) {
        int y, mo, da;
        if (a > 31) { y = a; mo = bb; da = c; } else { mo = a; da = bb; y = c; }
        if (y < 30) y += 2000; else if (y < 100) y += 1900;
        d = vb_dateserial(y, mo, da);
        ok = 1;
        const char *t = strchr(s, ' ');
        if (t) s = t; else s = "";
    }
    while (*s == ' ') s++;
    if (*s) {
        n2 = sscanf(s, "%d:%d:%d %3s", &h, &mi, &se, ap);
        if (n2 < 3) { se = 0; n2 = sscanf(s, "%d:%d %3s", &h, &mi, ap); if (n2 >= 2) n2 += 1; }
        if (n2 >= 3) {
            if ((ap[0] == 'p' || ap[0] == 'P') && h < 12) h += 12;
            if ((ap[0] == 'a' || ap[0] == 'A') && h == 12) h = 0;
            double t = (h * 3600 + mi * 60 + se) / 86400.0;
            d = d >= 0 ? d + t : d - t;
            ok = 1;
        } else ok = 0;
    }
    vb_release(m);
    if (!ok) { fail(p, 3075, "Syntax error in date in query expression '#%.*s#'.", k->n, k->s); return -1; }
    char txt[24];
    jet_date_text(d, txt);
    jb_qstr(b, txt, -1);
    return 0;
}

/* the VBA functions Jet SQL knows: name -> the SQLite function (registered by jet_register) */
static const char *const FUNCS[] = {
    "abs", "asc", "atn", "cbool", "cbyte", "ccur", "cdate", "cdbl", "chr", "chr$", "cint", "clng", "cos", "csng", "cstr",
    "cvar", "cvdate", "date", "date$", "dateadd", "datediff", "datepart", "dateserial", "datevalue", "day", "exp",
    "fix", "format", "format$", "hex", "hex$", "hour", "instr", "instrrev", "int", "isdate", "isempty", "isnull",
    "isnumeric", "lcase", "lcase$", "left", "left$", "len", "log", "ltrim", "ltrim$", "mid", "mid$", "minute",
    "month", "monthname", "now", "oct", "replace", "right", "right$", "rnd", "round", "rtrim", "rtrim$",
    "second", "sgn", "sin", "space", "space$", "sqr", "str", "str$", "strcomp", "strconv", "string", "string$",
    "strreverse", "tan", "time", "time$", "timeserial", "timevalue", "trim", "trim$", "ucase", "ucase$", "val",
    "weekday", "weekdayname", "year", "formatnumber", "formatcurrency", "formatpercent", "formatdatetime",
    "timer", "error", 0 };
static const char *const AGGS[] = { "count", "sum", "avg", "min", "max", "first", "last", "stdev", "stdevp", "var", "varp", 0 };

static int find_name(const char *const *list, const char *s, int n)
{
    for (int i = 0; list[i]; i++) if ((int)strlen(list[i]) == n && !strncasecmp(list[i], s, (size_t)n)) return i;
    return -1;
}

static int parse_args(ps *p, ex **args, int *nargs)
{
    *args = 0; *nargs = 0;
    if (!acceptop(p, "(")) { fail(p, 3075, "("); return -1; }
    if (acceptop(p, ")")) return 0;
    for (;;) {
        *args = realloc(*args, sizeof(ex) * (size_t)(*nargs + 1));
        ex *a = &(*args)[*nargs];
        memset(a, 0, sizeof *a);
        if (parse_expr(p, 0, a) < 0) { (*nargs)++; return -1; }
        (*nargs)++;
        if (acceptop(p, ",")) continue;
        if (acceptop(p, ")")) return 0;
        fail(p, 3075, ")");
        return -1;
    }
}
static void free_args(ex *a, int n) { for (int i = 0; i < n; i++) ex_free(&a[i]); free(a); }

static int parse_primary(ps *p, ex *e)
{
    tok *k = cur(p);
    jbuf *b = &e->b;
    if (k->k == T_NUM) {
        jb_put(b, k->s, k->n); p->i++;
        if (memchr(k->s, '.', (size_t)k->n)) e->type = dbDecimal;
        else if (memchr(k->s, 'e', (size_t)k->n) || memchr(k->s, 'E', (size_t)k->n)) {
            double v = atof(k->s);
            e->type = v == floor(v) && fabs(v) <= 2147483647.0 ? dbLong : dbDouble;
        }
        else e->type = atof(k->s) > 2147483647.0 ? dbDecimal : dbLong;
        return 0;
    }
    if (k->k == T_STR) { put_str_literal(b, k); p->i++; e->type = dbText; return 0; }
    if (k->k == T_DATE) { p->i++; e->type = dbDate; return put_date_literal(p, b, k); }
    if (k->k == T_QMARK) {                      /* ADO: each ? is a parameter of its own, "?1", "?2" ... */
        char nm[16]; snprintf(nm, sizeof nm, "?%d", ++p->nqmark);
        p->i++; jb_f(b, "?%d", param_index(p, nm, (int)strlen(nm))); return 0;
    }
    if (k->k == T_OP && k->s[0] == '(') {
        p->i++;
        if (iskw(p, "select")) {
            jb_s(b, "(");
            if (parse_select(p, b, 0) < 0) return -1;
            if (!acceptop(p, ")")) { fail(p, 3075, ")"); return -1; }
            jb_s(b, ")");
            return 0;
        }
        ex in; memset(&in, 0, sizeof in);
        if (parse_expr(p, 0, &in) < 0) { ex_free(&in); return -1; }
        if (!acceptop(p, ")")) { ex_free(&in); fail(p, 3075, ")"); return -1; }
        jb_s(b, "("); jb_put(b, in.b.p, in.b.n); jb_s(b, ")");
        e->boolean = in.boolean;
        e->type = in.type;
        ex_free(&in);
        return 0;
    }
    if (k->k == T_OP && k->s[0] == '*') { jb_s(b, "*"); p->i++; return 0; }
    int is_call = k->k == T_ID && p->i + 1 < p->nt && p->t[p->i + 1].k == T_OP && p->t[p->i + 1].s[0] == '(' &&
                  (find_name(FUNCS, k->s, k->n) >= 0 || find_name(AGGS, k->s, k->n) >= 0 || !reserved(k) ||
                   kw(k, "iif") || kw(k, "switch") || kw(k, "choose"));
    if (k->k == T_ID && (!reserved(k) || is_call)) {
        if (kw(k, "null")) { jb_s(b, "NULL"); p->i++; e->type = dbBinary; return 0; }
        if ((kw(k, "true") || kw(k, "false") || kw(k, "yes") || kw(k, "no") || kw(k, "on") || kw(k, "off")) &&
            !(p->i + 1 < p->nt && p->t[p->i + 1].k == T_OP && strchr("(.!", p->t[p->i + 1].s[0]))) {
            jb_s(b, kw(k, "true") || kw(k, "yes") || kw(k, "on") ? "-1" : "0"); p->i++; e->type = dbInteger; return 0;
        }
        /* a function */
        if (is_call) {
            char name[40]; snprintf(name, sizeof name, "%.*s", k->n < 39 ? k->n : 39, k->s);
            for (char *c = name; *c; c++) *c = (char)tolower((unsigned char)*c);
            p->i++;
            int agg = find_name(AGGS, k->s, k->n);
            if (agg >= 0) {
                /* Count(*), Count(DISTINCT x) is not Jet; aggregates over a value */
                if (!acceptop(p, "(")) return -1;
                if (isop(p, "*") && agg == 0) { p->i++; if (!acceptop(p, ")")) { fail(p, 3075, ")"); return -1; } jb_s(b, "count(*)"); e->type = dbLong; return 0; }
                ex a; memset(&a, 0, sizeof a);
                if (parse_expr(p, 0, &a) < 0) { ex_free(&a); return -1; }
                if (!acceptop(p, ")")) { ex_free(&a); fail(p, 3075, ")"); return -1; }
                static const char *const sq[] = { "count", "sum", "avg", "min", "max", "jet_first", "jet_last",
                                                  "jet_stdev", "jet_stdevp", "jet_var", "jet_varp" };
                jb_f(b, "%s(", sq[agg]);
                if (agg == 1 || agg == 2) { jb_s(b, "jet_num("); ex_value(&a, b); jb_s(b, ")"); }
                else ex_value(&a, b);
                jb_s(b, ")");
                /* Count: Long; Sum of whole numbers: Double; Sum/Avg of Currency: Currency */
                switch (agg) {
                case 0: e->type = dbLong; break;
                case 1: case 2: e->type = a.type == dbCurrency ? dbCurrency : a.type == dbDecimal ? dbDecimal : a.type ? dbDouble : 0; break;
                case 3: case 4: case 5: case 6: e->type = a.type; break;
                default: e->type = dbDouble;
                }
                ex_free(&a);
                return 0;
            }
            ex *args; int n;
            if (parse_args(p, &args, &n) < 0) { free_args(args, n); return -1; }
            if (!strcmp(name, "iif")) {
                if (n != 3) { free_args(args, n); fail(p, 3075, "iif"); return -1; }
                jb_s(b, "(CASE WHEN "); ex_cond(&args[0], b); jb_s(b, " THEN "); ex_value(&args[1], b);
                jb_s(b, " ELSE "); ex_value(&args[2], b); jb_s(b, " END)");
                e->type = args[1].type == args[2].type ? args[1].type : 0;
            } else if (!strcmp(name, "switch")) {
                jb_s(b, "(CASE");
                for (int i = 0; i + 1 < n; i += 2) {
                    jb_s(b, " WHEN "); ex_cond(&args[i], b); jb_s(b, " THEN "); ex_value(&args[i + 1], b);
                }
                jb_s(b, " ELSE NULL END)");
            } else if (!strcmp(name, "choose")) {
                jb_s(b, "(CASE CAST(");
                if (n) ex_value(&args[0], b);
                jb_s(b, " AS INTEGER)");
                for (int i = 1; i < n; i++) { jb_f(b, " WHEN %d THEN ", i); ex_value(&args[i], b); }
                jb_s(b, " ELSE NULL END)");
            } else {
                if (find_name(FUNCS, k->s, k->n) < 0) {
                    free_args(args, n);
                    fail(p, 3085, "Undefined function '%.*s' in expression.", k->n, k->s);
                    return -1;
                }
                char *dollar = strchr(name, '$');
                if (dollar) *dollar = 0;
                jb_f(b, "jet_%s(", name);
                for (int i = 0; i < n; i++) { if (i) jb_s(b, ", "); ex_value(&args[i], b); }
                jb_s(b, ")");
                e->type = func_type(name, n ? (args[0].boolean ? dbInteger : args[0].type) : 0);
            }
            free_args(args, n);
            return 0;
        }
    }
    if (k->k == T_ID || k->k == T_QID) {
        if (k->k == T_ID && reserved(k)) { fail(p, 3075, "reserved"); return -1; }
        /* a name, maybe qualified: a.b.c, a.*; Forms!Dialog!Date may be a parameter */
        int start = p->i;
        {
            char full[512] = "";
            int i = p->i, nparts = 0;
            for (;;) {
                tok *t = &p->t[i];
                if (t->k != T_ID && t->k != T_QID) break;
                if (nparts) strncat(full, "!", sizeof full - strlen(full) - 1);
                strncat(full, t->s, (size_t)t->n < sizeof full - strlen(full) - 1 ? (size_t)t->n : sizeof full - strlen(full) - 1);
                nparts++;
                i++;
                if (!(p->t[i].k == T_OP && (p->t[i].s[0] == '!' || p->t[i].s[0] == '.'))) break;
                i++;
            }
            if (nparts > 1 && is_param(p, full, (int)strlen(full))) {
                p->i = i;
                jb_f(b, "?%d", param_index(p, full, (int)strlen(full)));
                for (int k2 = 0; k2 < p->q->nparams; k2++)
                    if (!strcasecmp(p->q->params[k2].name, full)) e->type = p->q->params[k2].type;
                return 0;
            }
        }
        int parts = 0;
        for (;;) {
            tok *t = cur(p);
            if (t->k != T_ID && t->k != T_QID) {
                if (parts && t->k == T_OP && t->s[0] == '*') { jb_s(b, "*"); p->i++; parts++; break; }
                fail(p, 3075, "name"); return -1;
            }
            p->i++;
            parts++;
            if (!(isop(p, ".") || isop(p, "!"))) {
                if (parts == 1 && is_param(p, t->s, t->n)) {
                    b->n = 0; if (b->p) b->p[0] = 0;
                    jb_f(b, "?%d", param_index(p, t->s, t->n));
                    for (int i = 0; i < p->q->nparams; i++)
                        if ((int)strlen(p->q->params[i].name) == t->n && !strncasecmp(p->q->params[i].name, t->s, (size_t)t->n)) e->type = p->q->params[i].type;
                    return 0;
                }
                jb_qid(b, t->s, t->n);
                if (jet_coltype_fn) {
                    char col[256], qual[256] = "";
                    snprintf(col, sizeof col, "%.*s", t->n, t->s);
                    if (parts == 2) { tok *qt = &p->t[p->i - 3]; snprintf(qual, sizeof qual, "%.*s", qt->n, qt->s); }
                    e->type = jet_coltype_fn(jet_coltype_ctx, parts == 2 ? qual : 0, col);
                }
                break;
            }
            jb_qid(b, t->s, t->n);
            jb_s(b, ".");
            p->i++;
        }
        (void)start;
        e->simple_col = 1;
        return 0;
    }
    fail(p, 3075, "?");
    return -1;
}

/* binding powers (Jet's order, low to high) */
enum { BP_IMP = 1, BP_EQV, BP_XOR, BP_OR, BP_AND, BP_NOT, BP_CMP, BP_CAT, BP_ADD, BP_MOD, BP_IDIV, BP_MUL, BP_NEG, BP_POW };

static int infix_bp(ps *p, int *op)
{
    tok *k = cur(p);
    if (k->k == T_ID) {
        if (kw(k, "imp")) { *op = 'I'; return BP_IMP; }
        if (kw(k, "eqv")) { *op = 'E'; return BP_EQV; }
        if (kw(k, "xor")) { *op = 'X'; return BP_XOR; }
        if (kw(k, "or")) { *op = 'O'; return BP_OR; }
        if (kw(k, "and")) { *op = 'A'; return BP_AND; }
        if (kw(k, "like")) { *op = 'L'; return BP_CMP; }
        if (kw(k, "is")) { *op = 'S'; return BP_CMP; }
        if (kw(k, "between")) { *op = 'B'; return BP_CMP; }
        if (kw(k, "in")) { *op = 'N'; return BP_CMP; }
        if (kw(k, "not") && (iskw2(p, "like") || iskw2(p, "between") || iskw2(p, "in"))) { *op = '!'; return BP_CMP; }
        if (kw(k, "mod")) { *op = '%'; return BP_MOD; }
        return 0;
    }
    if (k->k != T_OP) return 0;
    if (k->n == 2) {
        if (!strncmp(k->s, "<>", 2) || !strncmp(k->s, "!=", 2)) { *op = 'n'; return BP_CMP; }
        if (!strncmp(k->s, "<=", 2)) { *op = 'l'; return BP_CMP; }
        if (!strncmp(k->s, ">=", 2)) { *op = 'g'; return BP_CMP; }
        return 0;
    }
    switch (k->s[0]) {
    case '=': *op = '='; return BP_CMP;
    case '<': *op = '<'; return BP_CMP;
    case '>': *op = '>'; return BP_CMP;
    case '&': *op = '&'; return BP_CAT;
    case '+': *op = '+'; return BP_ADD;
    case '-': *op = '-'; return BP_ADD;
    case '\\': *op = '\\'; return BP_IDIV;
    case '*': *op = '*'; return BP_MUL;
    case '/': *op = '/'; return BP_MUL;
    case '^': *op = '^'; return BP_POW;
    }
    return 0;
}

static int parse_in_list(ps *p, jbuf *b)
{
    if (!acceptop(p, "(")) { fail(p, 3075, "in"); return -1; }
    jb_s(b, "(");
    if (iskw(p, "select")) {
        if (parse_select(p, b, 0) < 0) return -1;
    } else {
        for (int first = 1;; first = 0) {
            ex a; memset(&a, 0, sizeof a);
            if (parse_expr(p, BP_CAT, &a) < 0) { ex_free(&a); return -1; }
            if (!first) jb_s(b, ", ");
            ex_value(&a, b);
            ex_free(&a);
            if (!acceptop(p, ",")) break;
        }
    }
    if (!acceptop(p, ")")) { fail(p, 3075, ")"); return -1; }
    jb_s(b, ")");
    return 0;
}

static int parse_expr(ps *p, int minbp, ex *out)
{
    int from = p->i;
    ex l; memset(&l, 0, sizeof l);
    /* prefix */
    if (iskw(p, "not")) {
        p->i++;
        ex a; memset(&a, 0, sizeof a);
        if (parse_expr(p, BP_NOT, &a) < 0) { ex_free(&a); ex_free(&l); if (!p->err) fail_expr(p, from); return -1; }
        jb_s(&l.b, "(NOT "); ex_cond(&a, &l.b); jb_s(&l.b, ")");
        l.boolean = 1;
        ex_free(&a);
    } else if (iskw(p, "exists")) {
        p->i++;
        if (!acceptop(p, "(") || !iskw(p, "select")) { ex_free(&l); fail_expr(p, from); return -1; }
        jb_s(&l.b, "(EXISTS (");
        if (parse_select(p, &l.b, 0) < 0 || !acceptop(p, ")")) { ex_free(&l); if (!p->err) fail_expr(p, from); return -1; }
        jb_s(&l.b, "))");
        l.boolean = 1;
    } else if (isop(p, "-") || isop(p, "+")) {
        int neg = cur(p)->s[0] == '-';
        p->i++;
        ex a; memset(&a, 0, sizeof a);
        if (parse_expr(p, BP_NEG, &a) < 0) { ex_free(&a); ex_free(&l); if (!p->err) fail_expr(p, from); return -1; }
        if (neg) { jb_s(&l.b, "jet_neg("); ex_value(&a, &l.b); jb_s(&l.b, ")"); }
        else ex_value(&a, &l.b);
        l.type = a.boolean ? dbInteger : a.type;
        ex_free(&a);
    } else {
        if (parse_primary(p, &l) < 0) {
            ex_free(&l);
            if (p->err == 3075 && strlen(p->msg) < 12) { int miss = !strcmp(p->msg, ")"); p->err = 0; fail_expr_kind(p, from, miss); }
            return -1;
        }
    }
    for (;;) {
        int op = 0, bp = infix_bp(p, &op);
        if (!bp || bp < minbp || (bp == minbp && bp != BP_POW && minbp)) {
            if (bp && bp == minbp && bp != BP_POW) {}       /* left associative: the caller takes it */
            break;
        }
        if (bp <= minbp && minbp) break;
        p->i++;
        ex r; memset(&r, 0, sizeof r);
        jbuf res = { 0 };
        int boolean = 0, ltype = 0;
        if (op == 'S') {                                /* Is [Not] Null */
            int nt = accept(p, "not");
            if (!accept(p, "null")) { ex_free(&l); fail_expr(p, from); return -1; }
            jb_s(&res, "("); ex_value(&l, &res); jb_s(&res, nt ? " IS NOT NULL)" : " IS NULL)");
            boolean = 1;
        } else if (op == '!' || op == 'L' || op == 'B' || op == 'N') {
            int nt = 0;
            if (op == '!') {
                nt = 1;
                if (accept(p, "like")) op = 'L'; else if (accept(p, "between")) op = 'B'; else if (accept(p, "in")) op = 'N';
            }
            if (op == 'L') {
                if (parse_expr(p, BP_CMP, &r) < 0) { ex_free(&r); ex_free(&l); if (!p->err) fail_expr(p, from); return -1; }
                jb_s(&res, nt ? "(NOT jet_like(" : "(jet_like(");
                ex_value(&l, &res); jb_s(&res, ", "); ex_value(&r, &res);
                jb_s(&res, p->ansi_like ? ", 1))" : ", 0))");
            } else if (op == 'B') {
                ex lo; memset(&lo, 0, sizeof lo);
                if (parse_expr(p, BP_CAT, &lo) < 0 || !accept(p, "and") || parse_expr(p, BP_CAT, &r) < 0) {
                    ex_free(&lo); ex_free(&r); ex_free(&l); if (!p->err) fail_expr(p, from); return -1;
                }
                jb_s(&res, "("); ex_value(&l, &res); jb_s(&res, nt ? " NOT BETWEEN " : " BETWEEN ");
                ex_value(&lo, &res); jb_s(&res, " AND "); ex_value(&r, &res); jb_s(&res, ")");
                ex_free(&lo);
            } else {
                jb_s(&res, "("); ex_value(&l, &res); jb_s(&res, nt ? " NOT IN " : " IN ");
                if (parse_in_list(p, &res) < 0) { jb_free(&res); ex_free(&l); if (!p->err) fail_expr(p, from); return -1; }
                jb_s(&res, ")");
            }
            boolean = 1;
        } else {
            if (parse_expr(p, op == '^' ? bp - 1 : bp, &r) < 0) {
                ex_free(&r); ex_free(&l); jb_free(&res);
                if (!p->err) fail_expr(p, from);
                return -1;
            }
            switch (op) {
            case 'O': case 'A':
                jb_s(&res, "("); ex_cond(&l, &res); jb_s(&res, op == 'O' ? " OR " : " AND "); ex_cond(&r, &res); jb_s(&res, ")");
                boolean = 1; break;
            case 'X': case 'E': case 'I': {
                const char *f = op == 'X' ? "jet_xor" : op == 'E' ? "jet_eqv" : "jet_imp";
                jb_f(&res, "%s(", f); ex_value(&l, &res); jb_s(&res, ", "); ex_value(&r, &res); jb_s(&res, ")");
                break; }
            case '=': case '<': case '>': case 'n': case 'l': case 'g': {
                const char *o = op == '=' ? " = " : op == '<' ? " < " : op == '>' ? " > " : op == 'n' ? " <> " : op == 'l' ? " <= " : " >= ";
                jb_s(&res, "("); ex_value(&l, &res); jb_s(&res, o); ex_value(&r, &res); jb_s(&res, ")");
                boolean = 1; break; }
            default: {
                const char *f = op == '&' ? "jet_cat" : op == '+' ? "jet_add" : op == '-' ? "jet_sub" : op == '*' ? "jet_mul" :
                                op == '/' ? "jet_div" : op == '\\' ? "jet_idiv" : op == '%' ? "jet_mod" : "jet_pow";
                jb_f(&res, "%s(", f); ex_value(&l, &res); jb_s(&res, ", "); ex_value(&r, &res); jb_s(&res, ")");
                ltype = op == '&' ? dbText : arith_type(l.boolean ? dbInteger : l.type, r.boolean ? dbInteger : r.type, op);
            }
            }
        }
        ex_free(&r);
        ex_free(&l);
        l.b = res; l.boolean = boolean; l.simple_col = 0; l.type = boolean ? dbInteger : ltype;
    }
    *out = l;
    return 0;
}


/* ---------------------------------------------------------------- other databases, linked tables */
/* the text of a string token without its quotes */
static char *str_text(tok *k)
{
    char q = k->s[0];
    char *o = malloc((size_t)k->n + 1);
    int n = 0;
    for (int i = 1; i < k->n - 1; i++) {
        if (k->s[i] == q && k->s[i + 1] == q) i++;
        o[n++] = k->s[i];
    }
    o[n] = 0;
    return o;
}
/* a connect string: "[kind];database=path;..." -> path (malloc, "" if none) */
static char *connect_database(const char *c)
{
    const char *p = c;
    while (p && *p) {
        while (*p == ';' || *p == ' ') p++;
        if (!strncasecmp(p, "database", 8)) {
            const char *q = p + 8;
            while (*q == ' ') q++;
            if (*q == '=') {
                q++;
                while (*q == ' ') q++;
                const char *e = strchr(q, ';');
                int n = e ? (int)(e - q) : (int)strlen(q);
                while (n > 0 && q[n - 1] == ' ') n--;
                return dupn(q, n);
            }
        }
        p = strchr(p, ';');
    }
    return strdup("");
}
/* the schema of another database: its number (0: this database) or -1 with the error */
static int ext_schema(ps *p, const char *path, const char *connect)
{
    if (!jet_ext_fn) { fail(p, 3131, "Syntax error in FROM clause."); return -1; }
    int err = 0; char msg[400] = "";
    int n = jet_ext_fn(jet_ext_ctx, path, connect, &err, msg, sizeof msg);
    if (n < 0) { fail(p, err ? err : 3044, "%s", msg); return -1; }
    return n;
}
/* "[connect]." before a table name: its schema; 0 when the name is not one (-1: error) */
static int parse_connect_prefix(ps *p, int *ext)
{
    tok *k = cur(p);
    *ext = 0;
    if (k->k != T_QID || !memchr(k->s, ';', (size_t)k->n) || p->i + 1 >= p->nt) return 0;
    tok *d = &p->t[p->i + 1];
    if (!(d->k == T_OP && d->n == 1 && d->s[0] == '.')) return 0;
    char *c = dupn(k->s, k->n), *path = connect_database(c);
    p->i += 2;
    int n = ext_schema(p, path, c);
    free(c); free(path);
    if (n < 0) return -1;
    *ext = n;
    return 1;
}
/* IN 'path' ['connect' | [connect]] after a table name: its schema (-1: error) */
static int parse_in_db(ps *p)
{
    tok *k = cur(p);
    if (k->k != T_STR) { fail(p, 3131, "Syntax error in FROM clause."); return -1; }
    char *path = str_text(k), *c = 0;
    p->i++;
    k = cur(p);
    if (k->k == T_STR) { c = str_text(k); p->i++; }
    else if (k->k == T_QID) { c = dupn(k->s, k->n); p->i++; }
    if (!*path && c) { free(path); path = connect_database(c); }
    int n = ext_schema(p, path, c ? c : "");
    free(path); free(c);
    return n;
}
/* a table of this database that is linked to one of another: its schema, the name there (0: not linked) */
static int link_schema(ps *p, const char *name, int len, char **src)
{
    *src = 0;
    if (!jet_link_fn) return 0;
    char nm[256], s[256] = "", msg[400] = "";
    int err = 0;
    snprintf(nm, sizeof nm, "%.*s", len, name);
    int n = jet_link_fn(jet_ext_ctx, nm, s, sizeof s, &err, msg, sizeof msg);
    if (n < 0) { fail(p, err ? err : 3044, "%s", msg); return -1; }
    if (n > 0) *src = strdup(s);
    return n;
}
static void add_tab(ps *p, const char *name, int n, int ext)
{
    jsql *q = p->q;
    q->tabs = realloc(q->tabs, sizeof(char *) * (size_t)(q->ntabs + 1));
    q->tabext = realloc(q->tabext, sizeof(int) * (size_t)(q->ntabs + 1));
    q->tabs[q->ntabs] = n < 0 ? strdup(name) : dupn(name, n);
    q->tabext[q->ntabs++] = ext;
}
static void put_table(jbuf *b, int ext, const char *name, int n)
{
    if (ext > 0) jb_f(b, "\"lbext%d\".", ext);
    jb_qid(b, name, n);
}
/* the IN 'file' of a FROM clause (it names the database of all its tables): looked for ahead, its schema */
static int from_in_ahead(ps *p)
{
    int depth = 0;
    for (int i = p->i; i < p->nt && p->t[i].k != T_END; i++) {
        tok *k = &p->t[i];
        if (k->k == T_OP && k->n == 1 && k->s[0] == '(') depth++;
        else if (k->k == T_OP && k->n == 1 && k->s[0] == ')') { if (--depth < 0) break; }
        else if (depth == 0 && (kw(k, "where") || kw(k, "group") || kw(k, "order") || kw(k, "having") || kw(k, "union") || kw(k, "on")))
            break;
        else if (depth == 0 && kw(k, "in") && i + 1 < p->nt && p->t[i + 1].k == T_STR) {
            int save = p->i;
            p->i = i + 1;
            int n = parse_in_db(p);
            p->i = save;
            return n;
        }
    }
    return 0;
}

/* ---------------------------------------------------------------- SELECT */
static int parse_name(ps *p, char **name)
{
    tok *k = cur(p);
    if ((k->k != T_ID && k->k != T_QID) || (k->k == T_ID && reserved(k))) return -1;
    *name = dupn(k->s, k->n);
    p->i++;
    return 0;
}

static int parse_from_item(ps *p, jbuf *b, int *simple, char **table, char **alias, int *text, char **tsrc);

static int parse_from_primary(ps *p, jbuf *b, int *simple, char **table, char **alias, int *text, char **tsrc)
{
    if (isop(p, "{") && p->i + 1 < p->nt && kw(&p->t[p->i + 1], "oj")) {
        /* ODBC's outer join escape, as the Query Designer of VB6 writes it: { oj A LEFT OUTER JOIN B ON ... } */
        p->i += 2;
        if (parse_from_item(p, b, simple, table, alias, text, tsrc) < 0) return -1;
        *simple = 0;
        if (!acceptop(p, "}")) { fail(p, 3131, "Syntax error in FROM clause."); return -1; }
        return 0;
    }
    if (acceptop(p, "(")) {
        if (iskw(p, "select")) {
            *simple = 0;
            jb_s(b, "(");
            if (parse_select(p, b, 0) < 0) return -1;
            if (!acceptop(p, ")")) { fail(p, 3131, "Syntax error in FROM clause."); return -1; }
            jb_s(b, ")");
        } else {
            *simple = 0;
            jb_s(b, "(");
            if (parse_from_item(p, b, simple, table, alias, text, tsrc) < 0) return -1;
            *simple = 0;
            if (!acceptop(p, ")")) { fail(p, 3131, "Syntax error in FROM clause."); return -1; }
            jb_s(b, ")");
        }
    } else {
        int ext = 0;
        int pre = parse_connect_prefix(p, &ext);  /* [;database=other.mdb].table */
        if (pre < 0) return -1;
        tok *k = cur(p);
        if ((k->k != T_ID && k->k != T_QID) || (k->k == T_ID && reserved(k))) { fail(p, 3131, "Syntax error in FROM clause."); return -1; }
        p->i++;
        char *src = 0;
        if (!pre) ext = p->from_ext;             /* the IN 'file' of the FROM clause */
        if (!pre && !ext) {                      /* a linked table */
            ext = link_schema(p, k->s, k->n, &src);
            if (ext < 0) return -1;
        }
        /* owner.table is not Jet; a name only */
        if (*table == 0) { *table = dupn(k->s, k->n); *text = ext; *tsrc = src ? strdup(src) : 0; } else *simple = 0;
        from_tables++;
        if (src) add_tab(p, src, -1, ext); else add_tab(p, k->s, k->n, ext);
        if (src) {
            put_table(b, ext, src, -1);
            if (!(iskw(p, "as") || (cur(p)->k == T_ID && !reserved(cur(p))) || cur(p)->k == T_QID)) { jb_s(b, " AS "); jb_qid(b, k->s, k->n); }
            free(src);
        } else put_table(b, ext, k->s, k->n);
        if (accept(p, "in")) {                  /* IN 'other.mdb': its schema was found ahead */
            if (pre || parse_in_db(p) < 0) { if (!p->err) fail(p, 3131, "Syntax error in FROM clause."); return -1; }
        }
    }
    int has_as = accept(p, "as");
    tok *k = cur(p);
    if ((k->k == T_ID && !reserved(k)) || k->k == T_QID) {
        jb_s(b, " AS "); jb_qid(b, k->s, k->n);
        if (*alias == 0) *alias = dupn(k->s, k->n);
        p->i++;
    } else if (has_as) { fail(p, 3131, "Syntax error in FROM clause."); return -1; }
    return 0;
}
static int parse_from_item(ps *p, jbuf *b, int *simple, char **table, char **alias, int *text, char **tsrc)
{
    if (parse_from_primary(p, b, simple, table, alias, text, tsrc) < 0) return -1;
    for (;;) {
        const char *j = 0;
        if (iskw(p, "inner") && iskw2(p, "join")) { p->i += 2; j = " INNER JOIN "; }
        else if (iskw(p, "left")) { p->i++; accept(p, "outer"); if (!accept(p, "join")) { fail(p, 3131, "Syntax error in FROM clause."); return -1; } j = " LEFT JOIN "; }
        else if (iskw(p, "right")) { p->i++; accept(p, "outer"); if (!accept(p, "join")) { fail(p, 3131, "Syntax error in FROM clause."); return -1; } j = " RIGHT JOIN "; }
        else if (iskw(p, "join")) { fail(p, 3131, "Syntax error in FROM clause."); return -1; }
        if (!j) return 0;
        *simple = 0;
        jb_s(b, j);
        if (parse_from_primary(p, b, simple, table, alias, text, tsrc) < 0) return -1;
        if (!accept(p, "on")) { fail(p, 3131, "Syntax error in JOIN operation."); return -1; }
        ex c; memset(&c, 0, sizeof c);
        if (parse_expr(p, 0, &c) < 0) { ex_free(&c); if (!p->err) fail(p, 3131, "Syntax error in JOIN operation."); return -1; }
        jb_s(b, " ON "); ex_cond(&c, b);
        ex_free(&c);
    }
}

static int simple_from_one(ps *p) { (void)p; return from_tables <= 1; }

/* one SELECT (top: the statement itself, which may have UNION, ORDER BY, INTO) */
static int parse_select_core(ps *p, jbuf *b, int top, int *topn, int *pct, int *simple_out)
{
    if (!accept(p, "select")) { fail(p, 3141, "The SELECT statement includes a reserved word or an argument name that is misspelled or missing, or the punctuation is incorrect."); return -1; }
    jb_s(b, "SELECT ");
    int simple = 1;
    int distinctrow = 0;
    size_t dpos = (size_t)b->n;
    if (accept(p, "distinct")) { jb_s(b, "DISTINCT "); simple = 0; }
    else if (accept(p, "distinctrow")) { simple = 0; distinctrow = 1; }
    else accept(p, "all");
    if (accept(p, "top")) {
        tok *k = cur(p);
        if (k->k != T_NUM) { fail(p, 3141, "The SELECT statement includes a reserved word or an argument name that is misspelled or missing, or the punctuation is incorrect."); return -1; }
        *topn = atoi(k->s);
        p->i++;
        if (accept(p, "percent")) *pct = 1;
        simple = 0;
    }
    /* the select list */
    jbuf list = { 0 };
    int nitems = 0;
    for (;;) {
        int from = p->i;
        ex e; memset(&e, 0, sizeof e);
        if (parse_expr(p, 0, &e) < 0) { ex_free(&e); jb_free(&list); if (!p->err) fail_expr(p, from); return -1; }
        if (nitems++) jb_s(&list, ", ");
        ex_value(&e, &list);
        int star = e.b.n && e.b.p[e.b.n - 1] == '*';
        char iname[64] = "";
        int etype = e.boolean ? dbInteger : e.type;
        if (accept(p, "as")) {
            char *a;
            if (parse_name(p, &a) < 0) { ex_free(&e); jb_free(&list); fail(p, 3141, "The SELECT statement includes a reserved word or an argument name that is misspelled or missing, or the punctuation is incorrect."); return -1; }
            jb_s(&list, " AS "); jb_qid(&list, a, -1);
            snprintf(iname, sizeof iname, "%s", a);
            free(a);
            simple = simple && e.simple_col;
        } else if (!e.simple_col && !star) {
            snprintf(iname, sizeof iname, "Expr%d", 1000 + nitems - 1);   /* Jet: 1000 + the column's place (VB6 oracle d_ado1) */
            p->exprno++;
            jb_f(&list, " AS \"%s\"", iname);
            simple = 0;
        }
        if (top && *iname && !e.simple_col && etype) {
            jsql *q = p->q;
            q->ctypes = realloc(q->ctypes, sizeof(jctype) * (size_t)(q->nctypes + 1));
            snprintf(q->ctypes[q->nctypes].name, sizeof q->ctypes[q->nctypes].name, "%s", iname);
            q->ctypes[q->nctypes].type = etype;
            q->nctypes++;
        }
        ex_free(&e);
        if (!acceptop(p, ",")) break;
    }
    if (top && accept(p, "into")) {
        char *t;
        int ext = 0, pre = parse_connect_prefix(p, &ext);   /* INTO [;database=other.mdb].t */
        if (pre < 0) { jb_free(&list); return -1; }
        if (parse_name(p, &t) < 0) { jb_free(&list); fail(p, 3141, "The SELECT statement includes a reserved word or an argument name that is misspelled or missing, or the punctuation is incorrect."); return -1; }
        if (!pre && accept(p, "in")) {           /* INTO t IN 'other.mdb' */
            ext = parse_in_db(p);
            if (ext < 0) { free(t); jb_free(&list); return -1; }
        }
        free(p->q->into);
        p->q->into = t;
        p->q->into_ext = ext;
        add_tab(p, t, -1, ext);
    }
    char *table = 0, *alias = 0, *tsrc = 0;
    int text = 0;
    int saved_from_ext = p->from_ext;
    jb_put(b, list.p ? list.p : "", list.n);
    jb_free(&list);
    if (accept(p, "from")) {
        from_tables = 0;
        p->from_ext = from_in_ahead(p);
        if (p->from_ext < 0) { p->from_ext = saved_from_ext; return -1; }
        jb_s(b, " FROM ");
        for (int n = 0;; n++) {
            if (n) { jb_s(b, ", "); simple = 0; }
            if (parse_from_item(p, b, &simple, &table, &alias, &text, &tsrc) < 0) { free(table); free(alias); free(tsrc); p->from_ext = saved_from_ext; return -1; }
            if (!acceptop(p, ",")) break;
        }
    } else simple = 0;
    p->from_ext = saved_from_ext;
    /* DISTINCTROW: the rows of the tables are distinct; over a join, the rows of the output are */
    if (distinctrow && table && !(simple_from_one(p))) {
        jbuf t = { 0 };
        jb_put(&t, b->p, (int)dpos); jb_s(&t, "DISTINCT "); jb_put(&t, b->p + dpos, b->n - (int)dpos);
        jb_free(b); *b = t;
    }
    if (accept(p, "where")) {
        int from = p->i;
        ex c; memset(&c, 0, sizeof c);
        if (parse_expr(p, 0, &c) < 0) { ex_free(&c); free(table); free(alias); free(tsrc); if (!p->err) fail_expr(p, from); return -1; }
        jb_s(b, " WHERE "); ex_cond(&c, b);
        ex_free(&c);
    }
    if (iskw(p, "group") && iskw2(p, "by")) {
        p->i += 2;
        simple = 0;
        jb_s(b, " GROUP BY ");
        for (int n = 0;; n++) {
            int from = p->i;
            ex c; memset(&c, 0, sizeof c);
            if (parse_expr(p, 0, &c) < 0) { ex_free(&c); free(table); free(alias); free(tsrc); if (!p->err) fail_expr(p, from); return -1; }
            if (n) jb_s(b, ", ");
            ex_value(&c, b);
            ex_free(&c);
            if (!acceptop(p, ",")) break;
        }
    }
    if (accept(p, "having")) {
        int from = p->i;
        ex c; memset(&c, 0, sizeof c);
        if (parse_expr(p, 0, &c) < 0) { ex_free(&c); free(table); free(alias); free(tsrc); if (!p->err) fail_expr(p, from); return -1; }
        jb_s(b, " HAVING "); ex_cond(&c, b);
        ex_free(&c);
    }
    if (simple && table && top) {
        free(p->q->base); free(p->q->base_alias); free(p->q->base_src);
        p->q->base = table; p->q->base_alias = alias;
        p->q->base_ext = text; p->q->base_src = tsrc;
    } else { free(table); free(alias); free(tsrc); }
    *simple_out = simple;
    return 0;
}

static int parse_order(ps *p, jbuf *b)
{
    if (!(iskw(p, "order") && iskw2(p, "by"))) return 0;
    p->i += 2;
    jb_s(b, " ORDER BY ");
    for (int n = 0;; n++) {
        int from = p->i;
        ex c; memset(&c, 0, sizeof c);
        if (parse_expr(p, 0, &c) < 0) { ex_free(&c); if (!p->err) fail_expr(p, from); return -1; }
        if (n) jb_s(b, ", ");
        ex_value(&c, b);
        ex_free(&c);
        if (accept(p, "desc")) jb_s(b, " DESC");
        else accept(p, "asc");
        if (!acceptop(p, ",")) break;
    }
    return 0;
}

static int parse_select(ps *p, jbuf *out, int top)
{
    int topn = 0, pct = 0, simple = 0;
    jbuf b = { 0 };
    int saved_exprno = p->exprno;
    if (!top) p->exprno = 0;
    if (parse_select_core(p, &b, top, &topn, &pct, &simple) < 0) { jb_free(&b); return -1; }
    while (accept(p, "union")) {
        int all = accept(p, "all");
        jb_s(&b, all ? " UNION ALL " : " UNION ");
        int t2 = 0, p2 = 0, s2 = 0;
        if (iskw(p, "select")) {
            if (parse_select_core(p, &b, 0, &t2, &p2, &s2) < 0) { jb_free(&b); return -1; }
        } else if (iskw(p, "table")) {
            p->i++;
            char *t; if (parse_name(p, &t) < 0) { jb_free(&b); fail(p, 3141, "union"); return -1; }
            jb_s(&b, "SELECT * FROM "); jb_qid(&b, t, -1); free(t);
        } else { jb_free(&b); fail(p, 3141, "The SELECT statement includes a reserved word or an argument name that is misspelled or missing, or the punctuation is incorrect."); return -1; }
        if (top) { free(p->q->base); p->q->base = 0; free(p->q->base_alias); p->q->base_alias = 0; free(p->q->base_src); p->q->base_src = 0; p->q->base_ext = 0; }
    }
    jbuf order = { 0 };
    if (parse_order(p, &order) < 0) { jb_free(&b); jb_free(&order); return -1; }
    if (topn > 0 && pct) {
        jb_s(out, "SELECT * FROM (");
        jb_put(out, b.p, b.n); jb_put(out, order.p ? order.p : "", order.n);
        jb_f(out, ") LIMIT (SELECT (count(*) * %d + 99) / 100 FROM (", topn);
        jb_put(out, b.p, b.n);
        jb_s(out, "))");
    } else {
        jb_put(out, b.p, b.n);
        jb_put(out, order.p ? order.p : "", order.n);
        if (topn > 0) jb_f(out, " LIMIT %d", topn);
    }
    jb_free(&b); jb_free(&order);
    if (!top) p->exprno = saved_exprno;
    return 0;
}

/* ---------------------------------------------------------------- DDL */
int jet_type_name(const char *s, int n, int *autoinc)
{
    static const struct { const char *n; int t; int ai; } T[] = {
        { "text", dbText }, { "char", dbText }, { "varchar", dbText }, { "alphanumeric", dbText }, { "string", dbText },
        { "character", dbText }, { "nchar", dbText }, { "nvarchar", dbText },
        { "memo", dbMemo }, { "longtext", dbMemo }, { "longchar", dbMemo }, { "note", dbMemo }, { "ntext", dbMemo },
        { "byte", dbByte }, { "integer1", dbByte }, { "tinyint", dbByte },
        { "short", dbInteger }, { "smallint", dbInteger }, { "integer2", dbInteger },
        { "long", dbLong }, { "int", dbLong }, { "integer", dbLong }, { "integer4", dbLong },
        { "counter", dbLong, 1 }, { "autoincrement", dbLong, 1 }, { "identity", dbLong, 1 },
        { "single", dbSingle }, { "real", dbSingle }, { "float4", dbSingle }, { "ieeesingle", dbSingle },
        { "double", dbDouble }, { "float", dbDouble }, { "float8", dbDouble }, { "ieeedouble", dbDouble }, { "number", dbDouble },
        { "currency", dbCurrency }, { "money", dbCurrency },
        { "datetime", dbDate }, { "date", dbDate }, { "time", dbDate }, { "timestamp", dbDate },
        { "bit", dbBoolean }, { "logical", dbBoolean }, { "logical1", dbBoolean }, { "yesno", dbBoolean }, { "boolean", dbBoolean },
        { "binary", dbBinary }, { "varbinary", dbBinary },
        { "longbinary", dbLongBinary }, { "oleobject", dbLongBinary }, { "general", dbLongBinary }, { "image", dbLongBinary },
        { "guid", dbGUID }, { "uniqueidentifier", dbGUID },
        { "decimal", dbDecimal }, { "numeric", dbDecimal }, { "dec", dbDecimal },
    };
    *autoinc = 0;
    for (size_t i = 0; i < sizeof T / sizeof *T; i++)
        if ((int)strlen(T[i].n) == n && !strncasecmp(T[i].n, s, (size_t)n)) { *autoinc = T[i].ai; return T[i].t; }
    return 0;
}

static int parse_name_list(ps *p, char ***names, int **desc, int *n)
{
    *names = 0; *n = 0; if (desc) *desc = 0;
    if (!acceptop(p, "(")) return -1;
    for (;;) {
        char *c;
        if (parse_name(p, &c) < 0) return -1;
        *names = realloc(*names, sizeof(char *) * (size_t)(*n + 1));
        if (desc) *desc = realloc(*desc, sizeof(int) * (size_t)(*n + 1));
        (*names)[*n] = c;
        int d = 0;
        if (accept(p, "desc")) d = 1; else accept(p, "asc");
        if (desc) (*desc)[*n] = d;
        (*n)++;
        if (acceptop(p, ",")) continue;
        if (acceptop(p, ")")) return 0;
        return -1;
    }
}

static int parse_constraint(ps *p, jsql *q, char *name)
{
    jcons c; memset(&c, 0, sizeof c);
    c.name = name;
    if (accept(p, "primary")) { if (!accept(p, "key")) goto bad; c.kind = 1; }
    else if (accept(p, "unique")) c.kind = 2;
    else if (accept(p, "foreign")) { if (!accept(p, "key")) goto bad; c.kind = 3; }
    else goto bad;
    if (parse_name_list(p, &c.cols, &c.desc, &c.ncols) < 0) goto bad;
    if (c.kind == 3) {
        if (!accept(p, "references") || parse_name(p, &c.ref_table) < 0) goto bad;
        if (isop(p, "(") && parse_name_list(p, &c.ref_cols, 0, &c.nref) < 0) goto bad;
    }
    q->cons = realloc(q->cons, sizeof(jcons) * (size_t)(q->ncons + 1));
    q->cons[q->ncons++] = c;
    return 0;
bad:
    free(name);
    return -1;
}

static int parse_column_def(ps *p, jsql *q)
{
    jcol c; memset(&c, 0, sizeof c);
    if (parse_name(p, &c.name) < 0) return -1;
    tok *k = cur(p);
    if (k->k != T_ID) { free(c.name); return -1; }
    int ai;
    c.type = jet_type_name(k->s, k->n, &ai);
    if (!c.type) { free(c.name); return -1; }
    /* DAO's SQL: no DECIMAL (3292), NUMERIC a Double (VB6 oracle d_jetdec); ADO's (ANSI): DECIMAL(p, s) */
    int numeric89 = 0;
    if (c.type == dbDecimal && !jet_ansi) {
        if (k->n == 7 && !strncasecmp(k->s, "numeric", 7)) { c.type = dbDouble; numeric89 = 1; }
        else { fail(p, 3292, "Syntax error in field definition."); free(c.name); return -1; }
    }
    p->i++;
    if (ai) c.attrs |= dbAutoIncrField;
    if (c.type == dbText) c.size = 255;
    if (c.type == dbBinary) c.size = 255;
    if (acceptop(p, "(")) {
        tok *nk = cur(p);
        if (nk->k != T_NUM) { free(c.name); return -1; }
        int v = atoi(nk->s), v2 = 0;
        p->i++;
        if (acceptop(p, ",")) { if (cur(p)->k == T_NUM) { v2 = atoi(cur(p)->s); p->i++; } }
        if (!acceptop(p, ")")) { free(c.name); return -1; }
        if (c.type == dbText || c.type == dbBinary) c.size = v;
        if (c.type == dbDecimal) c.numps = (v & 0xFF) | (v2 & 0xFF) << 8;
        else if (numeric89) { free(c.name); return -1; }             /* NUMERIC(10, 2) of DAO: 3290 */
        if (ai) {}                              /* COUNTER(seed, step): ignored */
    }
    for (;;) {
        if (accept(p, "not")) { if (!accept(p, "null")) { free(c.name); return -1; } c.notnull = 1; continue; }
        if (accept(p, "null")) continue;
        if (iskw(p, "constraint") || iskw(p, "primary") || iskw(p, "unique") || iskw(p, "references")) {
            char *cn = 0;
            if (accept(p, "constraint")) { if (parse_name(p, &cn) < 0) { free(c.name); return -1; } }
            if (accept(p, "primary")) { if (!accept(p, "key")) { free(cn); free(c.name); return -1; } c.pk = 1; }
            else if (accept(p, "unique")) c.unique = 1;
            else if (accept(p, "references")) {
                if (parse_name(p, &c.ref_table) < 0) { free(cn); free(c.name); return -1; }
                if (acceptop(p, "(")) {
                    if (parse_name(p, &c.ref_col) < 0 || !acceptop(p, ")")) { free(cn); free(c.name); return -1; }
                }
            } else { free(cn); free(c.name); return -1; }
            free(c.cname); c.cname = cn;
            continue;
        }
        break;
    }
    q->cols = realloc(q->cols, sizeof(jcol) * (size_t)(q->ncols + 1));
    q->cols[q->ncols++] = c;
    return 0;
}

static int parse_ddl(ps *p, jsql *q)
{
    if (accept(p, "create")) {
        int uniq = accept(p, "unique");
        if (!uniq && accept(p, "table")) {
            q->kind = J_CREATE_TABLE;
            if (parse_name(p, &q->table) < 0 || !acceptop(p, "(")) goto bad_table;
            for (;;) {
                if (iskw(p, "constraint")) {
                    p->i++;
                    char *cn; if (parse_name(p, &cn) < 0 || parse_constraint(p, q, cn) < 0) goto bad_table;
                } else if (iskw(p, "primary") || iskw(p, "unique") || iskw(p, "foreign")) {
                    if (parse_constraint(p, q, 0) < 0) goto bad_table;
                } else if (parse_column_def(p, q) < 0) goto bad_table;
                if (acceptop(p, ",")) continue;
                if (acceptop(p, ")")) break;
                goto bad_table;
            }
            return 0;
        bad_table:
            fail(p, 3290, "Syntax error in CREATE TABLE statement.");
            return -1;
        }
        if (accept(p, "index")) {
            q->kind = J_CREATE_INDEX;
            q->unique = uniq;
            char **names; int *desc, n;
            /* the index's name may be a word of SQL (CREATE INDEX IS ON T: VB6 oracle d_jetsort) */
            tok *ik = cur(p);
            if ((ik->k == T_ID && !kw(ik, "on")) || ik->k == T_QID) { q->index = dupn(ik->s, ik->n); p->i++; }
            if (!q->index || !accept(p, "on") || parse_name(p, &q->table) < 0 ||
                parse_name_list(p, &names, &desc, &n) < 0) goto bad_index;
            q->cols = calloc((size_t)n, sizeof(jcol));
            for (int i = 0; i < n; i++) { q->cols[i].name = names[i]; q->cols[i].attrs = desc[i] ? dbDescending : 0; }
            q->ncols = n;
            free(names); free(desc);
            if (accept(p, "with")) {
                for (;;) {
                    if (accept(p, "primary")) q->primary = 1;
                    else if (accept(p, "disallow")) { if (!accept(p, "null")) goto bad_index; q->disallow_null = 1; }
                    else if (accept(p, "ignore")) { if (!accept(p, "null")) goto bad_index; q->ignore_null = 1; }
                    else break;
                }
            }
            return 0;
        bad_index:
            fail(p, 3289, "Syntax error in CREATE INDEX statement.");
            return -1;
        }
        fail(p, 3129, "Invalid SQL statement; expected 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT', or 'UPDATE'.");
        return -1;
    }
    if (accept(p, "drop")) {
        if (accept(p, "table")) {
            q->kind = J_DROP_TABLE;
            if (parse_name(p, &q->table) < 0) { fail(p, 3295, "Syntax error in DROP TABLE or DROP INDEX."); return -1; }
            return 0;
        }
        if (accept(p, "index")) {
            q->kind = J_DROP_INDEX;
            if (parse_name(p, &q->index) < 0 || !accept(p, "on") || parse_name(p, &q->table) < 0) {
                fail(p, 3295, "Syntax error in DROP TABLE or DROP INDEX."); return -1;
            }
            return 0;
        }
        fail(p, 3295, "Syntax error in DROP TABLE or DROP INDEX.");
        return -1;
    }
    if (accept(p, "alter")) {
        if (!accept(p, "table") || parse_name(p, &q->table) < 0) goto bad_alter;
        if (accept(p, "add")) {
            if (accept(p, "constraint")) {
                char *cn; if (parse_name(p, &cn) < 0 || parse_constraint(p, q, cn) < 0) goto bad_alter;
                q->kind = J_ALTER_CONSTRAINT;
                return 0;
            }
            accept(p, "column");
            q->kind = J_ALTER_ADD;
            if (parse_column_def(p, q) < 0) goto bad_alter;
            return 0;
        }
        if (accept(p, "drop")) {
            if (accept(p, "constraint")) {
                q->kind = J_DROP_CONSTRAINT;
                if (parse_name(p, &q->index) < 0) goto bad_alter;
                return 0;
            }
            accept(p, "column");
            q->kind = J_ALTER_DROP;
            q->cols = calloc(1, sizeof(jcol));
            if (parse_name(p, &q->cols[0].name) < 0) goto bad_alter;
            q->ncols = 1;
            return 0;
        }
        if (accept(p, "alter")) {
            accept(p, "column");
            q->kind = J_ALTER_ADD;
            q->unique = -1;                     /* ALTER COLUMN: a new definition of the column */
            if (parse_column_def(p, q) < 0) goto bad_alter;
            return 0;
        }
    bad_alter:
        fail(p, 3293, "Syntax error in ALTER TABLE statement.");
        return -1;
    }
    return 1;
}

/* ---------------------------------------------------------------- statements */
static int parse_params(ps *p, jsql *q)
{
    if (!accept(p, "parameters")) return 0;
    for (;;) {
        char *name;
        if (parse_name(p, &name) < 0) goto bad;
        while (isop(p, "!") || isop(p, ".")) {    /* [Forms]![Dialog]![Date] */
            p->i++;
            char *more;
            if (parse_name(p, &more) < 0) { free(name); goto bad; }
            char *j = malloc(strlen(name) + strlen(more) + 2);
            sprintf(j, "%s!%s", name, more);
            free(name); free(more);
            name = j;
        }
        tok *k = cur(p);
        int type = dbText, ai;
        if (k->k == T_ID) {
            int t = jet_type_name(k->s, k->n, &ai);
            if (kw(k, "value")) t = dbText;
            if (!t) { free(name); goto bad; }
            type = t;
            p->i++;
            if (acceptop(p, "(")) { while (cur(p)->k != T_END && !isop(p, ")")) p->i++; if (!acceptop(p, ")")) { free(name); goto bad; } }
        }
        q->params = realloc(q->params, sizeof(jparam) * (size_t)(q->nparams + 1));
        q->params[q->nparams].name = name;
        q->params[q->nparams].type = type;
        q->nparams++;
        if (acceptop(p, ",")) continue;
        if (acceptop(p, ";")) return 0;
        goto bad;
    }
bad:
    fail(p, 3125, "Syntax error in PARAMETER clause.");
    return -1;
}

static int parse_statement(ps *p, jsql *q)
{
    jbuf b = { 0 };
    if (parse_params(p, q) < 0) return -1;
    if (iskw(p, "select")) {
        q->kind = J_SELECT;
        q->returns = 1;
        if (parse_select(p, &b, 1) < 0) { jb_free(&b); return -1; }
        if (q->into) q->returns = 0;
    } else if (iskw(p, "transform")) {
        q->kind = J_TRANSFORM;
        fail(p, 3001, "Crosstab queries (TRANSFORM ... PIVOT) are not supported by LinBASIC Studio.");
        return -1;
    } else if (accept(p, "insert")) {
        q->kind = J_INSERT;
        if (!accept(p, "into")) goto bad_insert;
        int ext = 0, pre = parse_connect_prefix(p, &ext), viain = 0;
        if (pre < 0) { jb_free(&b); return -1; }
        if (parse_name(p, &q->table) < 0) goto bad_insert;
        if (!pre && accept(p, "in")) {
            if ((ext = parse_in_db(p)) < 0) { jb_free(&b); return -1; }
            viain = 1;
        }
        if (!pre && !viain && (ext = link_schema(p, q->table, -1, &q->table_src)) < 0) { jb_free(&b); return -1; }
        q->table_ext = ext;
        add_tab(p, q->table_src ? q->table_src : q->table, -1, ext);
        jb_s(&b, "INSERT INTO "); put_table(&b, ext, q->table_src ? q->table_src : q->table, -1);
        if (isop(p, "(") && !iskw2(p, "select")) {
            char **names; int n;
            if (parse_name_list(p, &names, 0, &n) < 0) goto bad_insert;
            jb_s(&b, " (");
            for (int i = 0; i < n; i++) { if (i) jb_s(&b, ", "); jb_qid(&b, names[i], -1); free(names[i]); }
            jb_s(&b, ")");
            free(names);
        }
        if (accept(p, "values")) {
            if (viain) goto bad_insert;           /* (Jet: INSERT INTO t IN 'file' ... VALUES is a syntax error) */
            if (!acceptop(p, "(")) goto bad_insert;
            jb_s(&b, " VALUES (");
            for (int n = 0;; n++) {
                int from = p->i;
                ex e; memset(&e, 0, sizeof e);
                if (parse_expr(p, 0, &e) < 0) { ex_free(&e); if (!p->err) fail_expr(p, from); jb_free(&b); return -1; }
                if (n) jb_s(&b, ", ");
                ex_value(&e, &b);
                ex_free(&e);
                if (acceptop(p, ",")) continue;
                if (acceptop(p, ")")) break;
                goto bad_insert;
            }
            jb_s(&b, ")");
        } else if (iskw(p, "select") || (isop(p, "(") && iskw2(p, "select"))) {
            int paren = acceptop(p, "(");
            jb_s(&b, " ");
            jsql *save = p->q;
            if (parse_select(p, &b, 0) < 0) { jb_free(&b); return -1; }
            p->q = save;
            if (paren && !acceptop(p, ")")) goto bad_insert;
        } else goto bad_insert;
        goto done;
    bad_insert:
        jb_free(&b);
        fail(p, 3134, "Syntax error in INSERT INTO statement.");
        return -1;
    } else if (accept(p, "update")) {
        q->kind = J_UPDATE;
        int ext = 0, pre = parse_connect_prefix(p, &ext);
        if (pre < 0) { jb_free(&b); return -1; }
        if (parse_name(p, &q->table) < 0) goto bad_update;
        if (!pre && accept(p, "in")) { if ((ext = parse_in_db(p)) < 0) { jb_free(&b); return -1; } }
        else if (!pre && (ext = link_schema(p, q->table, -1, &q->table_src)) < 0) { jb_free(&b); return -1; }
        q->table_ext = ext;
        add_tab(p, q->table_src ? q->table_src : q->table, -1, ext);
        jb_s(&b, "UPDATE "); put_table(&b, ext, q->table_src ? q->table_src : q->table, -1);
        if (q->table_src) { jb_s(&b, " AS "); jb_qid(&b, q->table, -1); }
        if (iskw(p, "as") || (cur(p)->k == T_ID && !iskw(p, "set"))) goto bad_update;
        if (!accept(p, "set")) goto bad_update;
        jb_s(&b, " SET ");
        for (int n = 0;; n++) {
            char *col;
            tok *k = cur(p);
            if (k->k != T_ID && k->k != T_QID) goto bad_update;
            /* t.col = ... : the table name is not needed */
            if (parse_name(p, &col) < 0) goto bad_update;
            if (acceptop(p, ".")) { free(col); if (parse_name(p, &col) < 0) goto bad_update; }
            if (!acceptop(p, "=")) { free(col); goto bad_update; }
            int from = p->i;
            ex e; memset(&e, 0, sizeof e);
            if (parse_expr(p, 0, &e) < 0) { free(col); ex_free(&e); if (!p->err) fail_expr(p, from); jb_free(&b); return -1; }
            if (n) jb_s(&b, ", ");
            jb_qid(&b, col, -1); jb_s(&b, " = "); ex_value(&e, &b);
            free(col); ex_free(&e);
            if (!acceptop(p, ",")) break;
        }
        if (accept(p, "where")) {
            int from = p->i;
            ex c; memset(&c, 0, sizeof c);
            if (parse_expr(p, 0, &c) < 0) { ex_free(&c); if (!p->err) fail_expr(p, from); jb_free(&b); return -1; }
            jb_s(&b, " WHERE "); ex_cond(&c, &b);
            ex_free(&c);
        }
        goto done;
    bad_update:
        jb_free(&b);
        fail(p, 3144, "Syntax error in UPDATE statement.");
        return -1;
    } else if (accept(p, "delete")) {
        q->kind = J_DELETE;
        /* DELETE [*|t.*] FROM t */
        if (!iskw(p, "from")) {
            while (cur(p)->k != T_END && !iskw(p, "from")) p->i++;
        }
        int ext = 0, pre = 0, hasfrom = accept(p, "from");
        if (hasfrom && (pre = parse_connect_prefix(p, &ext)) < 0) { jb_free(&b); return -1; }
        if (!hasfrom || parse_name(p, &q->table) < 0) {
            jb_free(&b);
            fail(p, 3128, "Specify the table containing the records you want to delete.");
            return -1;
        }
        if (!pre && accept(p, "in")) { if ((ext = parse_in_db(p)) < 0) { jb_free(&b); return -1; } }
        else if (!pre && (ext = link_schema(p, q->table, -1, &q->table_src)) < 0) { jb_free(&b); return -1; }
        q->table_ext = ext;
        add_tab(p, q->table_src ? q->table_src : q->table, -1, ext);
        jb_s(&b, "DELETE FROM "); put_table(&b, ext, q->table_src ? q->table_src : q->table, -1);
        if (q->table_src) { jb_s(&b, " AS "); jb_qid(&b, q->table, -1); }
        if (accept(p, "where")) {
            int from = p->i;
            ex c; memset(&c, 0, sizeof c);
            if (parse_expr(p, 0, &c) < 0) { ex_free(&c); if (!p->err) fail_expr(p, from); jb_free(&b); return -1; }
            jb_s(&b, " WHERE "); ex_cond(&c, &b);
            ex_free(&c);
        }
    } else {
        int r = parse_ddl(p, q);
        if (r <= 0) return r;
        fail(p, 3129, "Invalid SQL statement; expected 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT', or 'UPDATE'.");
        return -1;
    }
done:
    acceptop(p, ";");
    if (cur(p)->k != T_END) {
        jb_free(&b);
        if (q->kind == J_SELECT) fail(p, 3142, "Characters found after end of SQL statement.");
        else fail(p, 3142, "Characters found after end of SQL statement.");
        return -1;
    }
    q->sql = jb_take(&b);
    return 0;
}

/* the text Jet quotes for a string not closed: its expression, from after the last "," or "(" to the end (the text
   ends at a Chr(0): INSERT INTO T VALUES (11, 'ae<0>') -> 'ae, d_jetsort) */
static void open_string_expr(const char *s, char *out, int cap)
{
    int start = 0, open = -1;
    char q = 0;
    for (int i = 0; s[i]; i++) {
        char c = s[i];
        if (q) {
            if (c == q) { if (s[i + 1] == q) i++; else q = 0; }
        } else if (c == '\'' || c == '"') { q = c; open = i; }
        else if (c == '[') { while (s[i + 1] && s[i + 1] != ']') i++; }
        else if (c == ',' || c == '(') start = i + 1;
    }
    if (!q) { snprintf(out, (size_t)cap, "%s", s); return; }
    while (start < open && isspace((unsigned char)s[start])) start++;
    snprintf(out, (size_t)cap, "%s", s + start);
}

int jet_ansi;                                   /* ADO (Jet OLE DB): LIKE with % and _ */
int jet_translate(const char *jet, jsql *q, const char *const *params)
{
    memset(q, 0, sizeof *q);
    ps p; memset(&p, 0, sizeof p);
    char err[64];
    p.nt = tokenize(jet, &p.t, err, sizeof err);
    if (p.nt < 0) {
        q->err = 3075;
        char ex[256]; open_string_expr(jet, ex, sizeof ex);
        snprintf(q->errmsg, sizeof q->errmsg, "Syntax error in string in query expression '%s'.", ex);
        return -1;
    }
    p.nt++;                                     /* the T_END */
    p.q = q; p.params = params; p.src = jet;
    p.ansi_like = params == (const char *const *)1 || jet_ansi;
    if (params == (const char *const *)1) p.params = 0;
    int r = parse_statement(&p, q);
    if (r < 0 || p.err) {
        q->err = p.err ? p.err : 3129;
        snprintf(q->errmsg, sizeof q->errmsg, "%s", p.err ? p.msg : "Invalid SQL statement; expected 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT', or 'UPDATE'.");
        free(p.t);
        return -1;
    }
    free(p.t);
    return 0;
}

int jet_expr(const char *jet, jbuf *out, int boolean, char *errmsg, int errcap)
{
    jsql q; memset(&q, 0, sizeof q);
    ps p; memset(&p, 0, sizeof p);
    char err[64];
    p.nt = tokenize(jet, &p.t, err, sizeof err);
    if (p.nt < 0) { snprintf(errmsg, (size_t)errcap, "Syntax error in string in query expression '%s'.", jet); return 3075; }
    p.nt++;
    p.q = &q;
    ex e; memset(&e, 0, sizeof e);
    int r = parse_expr(&p, 0, &e);
    if (r == 0 && cur(&p)->k != T_END) { fail_expr(&p, 0); r = -1; }
    if (r < 0 || p.err) {
        if (!p.err) fail_expr(&p, 0);
        snprintf(errmsg, (size_t)errcap, "%s", p.msg);
        ex_free(&e); free(p.t); jsql_free(&q);
        return p.err;
    }
    if (boolean) ex_cond(&e, out); else ex_value(&e, out);
    ex_free(&e); free(p.t); jsql_free(&q);
    return 0;
}

void jsql_free(jsql *q)
{
    free(q->sql); free(q->into); free(q->table); free(q->index); free(q->base); free(q->base_alias); free(q->base_src); free(q->table_src);
    for (int i = 0; i < q->ntabs; i++) free(q->tabs[i]);
    free(q->tabs); free(q->tabext); q->tabs = 0; q->tabext = 0; q->ntabs = 0;
    for (int i = 0; i < q->ncols; i++) { free(q->cols[i].name); free(q->cols[i].cname); free(q->cols[i].ref_table); free(q->cols[i].ref_col); }
    free(q->cols);
    for (int i = 0; i < q->ncons; i++) {
        jcons *c = &q->cons[i];
        free(c->name);
        for (int j = 0; j < c->ncols; j++) free(c->cols[j]);
        free(c->cols); free(c->desc); free(c->ref_table);
        for (int j = 0; j < c->nref; j++) free(c->ref_cols[j]);
        free(c->ref_cols);
    }
    free(q->cons);
    for (int i = 0; i < q->nparams; i++) free(q->params[i].name);
    free(q->params);
    free(q->ctypes);
    for (int i = 0; i < q->nbinds; i++) free(q->binds[i]);
    free(q->binds);
    memset(q, 0, sizeof *q);
}

/* ---------------------------------------------------------------- types */
int jet_decl_type(const char *decl, int *size, int *attrs)
{
    *size = 0; *attrs = 0;
    if (!decl) return 0;
    char d[64]; int k = 0;
    for (; decl[k] && k < 63; k++) d[k] = (char)toupper((unsigned char)decl[k]);
    d[k] = 0;
    if (!strcmp(d, "INTEGER")) { *size = 4; *attrs = dbAutoIncrField | dbFixedField; return dbLong; }
    static const struct { const char *p; int t, sz; } T[] = {
        { "BIT_INT", dbBoolean, 1 }, { "BYTE_INT", dbByte, 1 }, { "SHORT_INT", dbInteger, 2 }, { "LONG_INT", dbLong, 4 },
        { "CURRENCY_REAL", dbCurrency, 8 }, { "SINGLE_REAL", dbSingle, 4 }, { "DOUBLE", dbDouble, 8 },
        { "DATETIME_TEXT", dbDate, 8 }, { "MEMO_TEXT", dbMemo, 0 }, { "BINARY_BLOB", dbBinary, 255 },
        { "LONGBINARY_BLOB", dbLongBinary, 0 }, { "GUID_TEXT", dbGUID, 16 }, { "DECIMAL_NUMERIC", dbDecimal, 16 },
        { "TEXT", dbText, 255 },
    };
    for (size_t i = 0; i < sizeof T / sizeof *T; i++) {
        size_t n = strlen(T[i].p);
        if (!strncmp(d, T[i].p, n) && (d[n] == 0 || d[n] == '(' || d[n] == ' ')) {
            *size = T[i].sz;
            if (d[n] == '(') *size = atoi(d + n + 1);
            if (T[i].t == dbText || T[i].t == dbMemo || T[i].t == dbBinary || T[i].t == dbLongBinary) *attrs = dbVariableField;
            else *attrs = dbFixedField;
            return T[i].t;
        }
    }
    /* a SQLite database made elsewhere */
    if (strstr(d, "INT")) { *size = 4; *attrs = dbFixedField; return dbLong; }
    if (strstr(d, "CHAR") || strstr(d, "CLOB") || strstr(d, "TEXT")) { *size = 255; *attrs = dbVariableField; return dbText; }
    if (strstr(d, "BLOB")) { *attrs = dbVariableField; return dbLongBinary; }
    if (strstr(d, "REAL") || strstr(d, "FLOA") || strstr(d, "DOUB")) { *size = 8; *attrs = dbFixedField; return dbDouble; }
    if (strstr(d, "DATE") || strstr(d, "TIME")) { *size = 8; *attrs = dbFixedField; return dbDate; }
    if (strstr(d, "BOOL")) { *size = 1; *attrs = dbFixedField; return dbBoolean; }
    return 0;
}
const char *jet_type_decl(int type, int autoinc, char *buf, int cap)
{
    if (autoinc) { snprintf(buf, (size_t)cap, "INTEGER"); return buf; }
    switch (type) {
    case dbBoolean: return "BIT_INT";
    case dbByte: return "BYTE_INT";
    case dbInteger: return "SHORT_INT";
    case dbLong: return "LONG_INT";
    case dbCurrency: return "CURRENCY_REAL";
    case dbSingle: return "SINGLE_REAL";
    case dbDouble: case dbFloat: return "DOUBLE";
    case dbDate: case dbTime: case dbTimeStamp: return "DATETIME_TEXT";
    case dbMemo: return "MEMO_TEXT";
    case dbBinary: case dbVarBinary: return "BINARY_BLOB";
    case dbLongBinary: return "LONGBINARY_BLOB";
    case dbGUID: return "GUID_TEXT";
    case dbDecimal: case dbNumeric: case dbBigInt: return "DECIMAL_NUMERIC";
    default: return "TEXT";
    }
}

/* ---------------------------------------------------------------- values */
static vbvar text_value(const unsigned char *t, int n)
{
    vb_date d;
    if (jet_text_date((const char *)t, n, &d)) return vb_v_date(d);
    char *tmp = malloc((size_t)n + 1);
    memcpy(tmp, t, (size_t)n); tmp[n] = 0;
    vbvar v = vb_v_s(vb_sfromutf8(tmp));
    free(tmp);
    return v;
}
static vbvar blob_value(const void *p, int n)
{
    vbarr a; memset(&a, 0, sizeof a);
    if (n > 0) {
        vb_adim(&a, 1, 0, 0, 1, 0, n - 1);
        memcpy(vb_aelem(&a, 1, 0), p, (size_t)n);
    } else vb_aempty(&a, 1, 0);
    vbvar v = vb_v_arr(&a, VT_UI1);
    vb_afree(&a);
    return v;
}
/* a value as the type of its field */
static vbvar as_type(vbvar v, int dtype)
{
    if (v.vt == VT_NULL || v.vt == VT_EMPTY) return dtype == dbBoolean ? vb_v_bool(0) : v;
    switch (dtype) {
    case dbBoolean: return vb_v_bool(vb_v2d(v) != 0 ? -1 : 0);
    case dbByte: return vb_v_byte((vb_u8)vb_v2l(v));
    case dbInteger: return vb_v_i((vb_i16)vb_v2l(v));
    case dbLong: return vb_v_l(vb_v2l(v));
    case dbCurrency: return vb_v_c(vb_v2c(v));
    case dbSingle: return vb_v_f((float)vb_v2d(v));
    case dbDouble: case dbFloat: return vb_v_d(vb_v2d(v));
    case dbDecimal: case dbNumeric: return vb_cdec(v);          /* a Decimal field: a Variant Decimal (Jet 4) */
    case dbDate: return v.vt == VT_DATE ? v : vb_v_date(vb_v2date(v));
    case dbText: case dbMemo: case dbGUID: return v.vt == VT_BSTR ? v : vb_v_s(vb_v2s(v));
    }
    return v;
}
vbvar jet_value(sqlite3_value *sv, int dtype)
{
    vbvar v;
    switch (sqlite3_value_type(sv)) {
    case SQLITE_INTEGER: {
        sqlite3_int64 i = sqlite3_value_int64(sv);
        if (dtype == dbDecimal || dtype == dbNumeric) { v = vb_dec_from_i64(i); break; }
        v = (i >= -2147483647LL - 1 && i <= 2147483647LL) ? vb_v_l((vb_i32)i) : vb_v_d((double)i);
        break; }
    case SQLITE_FLOAT: v = vb_v_d(sqlite3_value_double(sv)); break;
    case SQLITE_TEXT: v = text_value(sqlite3_value_text(sv), sqlite3_value_bytes(sv)); break;
    case SQLITE_BLOB:
        /* a Decimal of a Jet file with more digits than a Double holds: its text, kept as a blob (mdbimport.c) */
        if (dtype == dbDecimal || dtype == dbNumeric) v = text_value(sqlite3_value_blob(sv), sqlite3_value_bytes(sv));
        else v = blob_value(sqlite3_value_blob(sv), sqlite3_value_bytes(sv));
        break;
    default: v = vb_v_null();
    }
    vb_i32 save = vb_errpend;
    vbvar r = dtype ? as_type(v, dtype) : v;
    if (vb_errpend && !save) { vb_err_clear_pending(); r = v; }
    return r;
}
vbvar jet_column(sqlite3_stmt *st, int i, int dtype)
{
    return jet_value(sqlite3_column_value(st, i), dtype);
}
vbvar jet_arg(sqlite3_value *v) { return jet_value(v, 0); }

static int bind_text(sqlite3_stmt *st, int i, vbs s)
{
    char *u = vb_stoutf8(s);
    return sqlite3_bind_text(st, i, u, -1, SQLITE_TRANSIENT);
}
int jet_bind(sqlite3_stmt *st, int i, vbvar v, int dtype)
{
    if (v.vt == (VT_VARIANT | 0x4000)) v = *(vbvar *)v.u.o;   /* defensive: by reference */
    if (v.vt == VT_NULL || v.vt == VT_EMPTY) { sqlite3_bind_null(st, i); return 0; }
    if (v.vt & VT_ARRAY) {
        vbarr *a = vb_v_arrptr(v);
        if (vb_errpend) return 3421;
        int n = a && a->data ? (int)(a->esz * (size_t)vb_ubound(a, 1) - a->esz * (size_t)vb_lbound(a, 1) + a->esz) : 0;
        sqlite3_bind_blob(st, i, a ? a->data : 0, n, SQLITE_TRANSIENT);
        return 0;
    }
    vb_i32 save = vb_errpend;
    switch (dtype) {
    case dbBoolean: sqlite3_bind_int(st, i, vb_v2bool(v) ? -1 : 0); break;
    case dbByte: case dbInteger: case dbLong: {
        double d = vb_v2d(v);
        if (vb_errpend && !save) { vb_err_clear_pending(); return 3421; }
        double r = vb_round_even(d);
        if ((dtype == dbByte && (r < 0 || r > 255)) || (dtype == dbInteger && (r < -32768 || r > 32767)) ||
            (dtype == dbLong && (r < -2147483648.0 || r > 2147483647.0))) return 3349;
        sqlite3_bind_int64(st, i, (sqlite3_int64)r);
        break; }
    case dbDecimal: case dbNumeric: {           /* exact, as the import keeps it: up to 15 digits a number, more the text */
        vbvar dv = vb_cdec(v);
        if (vb_errpend && !save) { vb_err_clear_pending(); return 3421; }
        char t[100];
        vb_dec_text(dv, 0, t);
        int nd = 0;
        for (const char *q = t; *q; q++) if (isdigit((unsigned char)*q)) nd++;
        for (const char *q = t; *q == '-' || *q == '0' || *q == '.'; q++) if (*q == '0') nd--;
        if (nd <= 15) sqlite3_bind_text(st, i, t, -1, SQLITE_TRANSIENT);
        else sqlite3_bind_blob(st, i, t, (int)strlen(t), SQLITE_TRANSIENT);
        break; }
    case dbCurrency: case dbSingle: case dbDouble: case dbFloat: {
        double d = vb_v2d(v);
        if (vb_errpend && !save) { vb_err_clear_pending(); return 3421; }
        if (dtype == dbCurrency) d = vb_cy2d(vb_ccur(d));
        if (dtype == dbSingle) d = (double)(float)d;
        sqlite3_bind_double(st, i, d);
        break; }
    case dbDate: case dbTime: case dbTimeStamp: {
        vb_date d = vb_v2date(v);
        if (vb_errpend && !save) { vb_err_clear_pending(); return 3421; }
        char t[24]; jet_date_text(d, t);
        sqlite3_bind_text(st, i, t, -1, SQLITE_TRANSIENT);
        break; }
    case dbText: case dbMemo: case dbGUID:
        bind_text(st, i, vb_v2s(v));
        break;
    default:
        switch (v.vt) {
        case VT_I2: sqlite3_bind_int(st, i, v.u.i); break;
        case VT_I4: sqlite3_bind_int(st, i, v.u.l); break;
        case VT_UI1: sqlite3_bind_int(st, i, v.u.b); break;
        case VT_BOOL: sqlite3_bind_int(st, i, v.u.i ? -1 : 0); break;
        case VT_R4: sqlite3_bind_double(st, i, v.u.f); break;
        case VT_R8: sqlite3_bind_double(st, i, v.u.d); break;
        case VT_CY: sqlite3_bind_double(st, i, vb_cy2d(v.u.cy)); break;
        case VT_DATE: { char t[24]; jet_date_text(v.u.d, t); sqlite3_bind_text(st, i, t, -1, SQLITE_TRANSIENT); break; }
        default: bind_text(st, i, vb_v2s(v));
        }
    }
    if (vb_errpend && !save) { vb_err_clear_pending(); return 3421; }
    return 0;
}

void jet_result(sqlite3_context *cx, vbvar v)
{
    switch (v.vt) {
    case VT_EMPTY: case VT_NULL: sqlite3_result_null(cx); break;
    case VT_I2: sqlite3_result_int(cx, v.u.i); break;
    case VT_I4: sqlite3_result_int(cx, v.u.l); break;
    case VT_UI1: sqlite3_result_int(cx, v.u.b); break;
    case VT_BOOL: sqlite3_result_int(cx, v.u.i ? -1 : 0); break;
    case VT_R4: sqlite3_result_double(cx, v.u.f); break;
    case VT_R8: sqlite3_result_double(cx, v.u.d); break;
    case VT_CY: sqlite3_result_double(cx, vb_cy2d(v.u.cy)); break;
    case VT_DATE: { char t[24]; jet_date_text(v.u.d, t); sqlite3_result_text(cx, t, -1, SQLITE_TRANSIENT); break; }
    case VT_BSTR: sqlite3_result_text(cx, vb_stoutf8(v.u.s), -1, SQLITE_TRANSIENT); break;
    default: sqlite3_result_null(cx);
    }
}

/* ---------------------------------------------------------------- the SQL functions */
static int any_null(int argc, sqlite3_value **argv)
{
    for (int i = 0; i < argc; i++) if (sqlite3_value_type(argv[i]) == SQLITE_NULL) return 1;
    return 0;
}
static vbvar A(sqlite3_value **argv, int i) { return jet_arg(argv[i]); }

/* run the runtime's VBA code; a run-time error becomes the error of the statement */
#define BEGIN_FN vb_i32 mk_ = vb_mark(); vb_i32 pend_ = vb_errpend; vbvar r_ = vb_v_null(); (void)r_
#define END_FN  if (vb_errpend && !pend_) { \
        char m_[300]; snprintf(m_, sizeof m_, "#%d %s", vb_err_number_pending(), vb_stoutf8(vb_err_description())); \
        vb_err_clear_pending(); sqlite3_result_error(cx, m_, -1); } \
    else jet_result(cx, r_); \
    vb_release(mk_)

static int fn_id(sqlite3_context *cx) { return (int)(intptr_t)sqlite3_user_data(cx); }

enum { F_UCASE = 1, F_LCASE, F_LEFT, F_RIGHT, F_MID, F_LEN, F_TRIM, F_LTRIM, F_RTRIM, F_INSTR, F_INSTRREV, F_SPACE,
       F_STRING, F_ASC, F_CHR, F_STRREVERSE, F_REPLACE, F_STRCOMP, F_STRCONV, F_VAL, F_STR, F_CSTR, F_CINT, F_CLNG,
       F_CDBL, F_CSNG, F_CCUR, F_CBYTE, F_CBOOL, F_CDATE, F_CVAR, F_INT, F_FIX, F_ABS, F_SGN, F_SQR, F_EXP, F_LOG,
       F_SIN, F_COS, F_TAN, F_ATN, F_RND, F_ROUND, F_HEX, F_OCT, F_YEAR, F_MONTH, F_DAY, F_HOUR, F_MINUTE,
       F_SECOND, F_WEEKDAY, F_DATE, F_TIME, F_NOW, F_DATEADD, F_DATEDIFF, F_DATEPART, F_DATESERIAL, F_TIMESERIAL,
       F_DATEVALUE, F_TIMEVALUE, F_FORMAT, F_NZ, F_ISNULL, F_ISNUMERIC, F_ISDATE, F_ISEMPTY, F_MONTHNAME,
       F_WEEKDAYNAME, F_FORMATNUMBER, F_FORMATCURRENCY, F_FORMATPERCENT, F_FORMATDATETIME, F_TIMER, F_ERROR,
       F_CAT, F_ADD, F_SUB, F_MUL, F_DIV, F_IDIV, F_MOD, F_POW, F_NEG, F_XOR, F_EQV, F_IMP, F_NUM };

static vb_i32 opt_l(int argc, sqlite3_value **argv, int i, vb_i32 d) { return i < argc && sqlite3_value_type(argv[i]) != SQLITE_NULL ? vb_v2l(A(argv, i)) : d; }

static void fn_vba(sqlite3_context *cx, int argc, sqlite3_value **argv)
{
    int id = fn_id(cx);
    BEGIN_FN;
    /* the functions that give Null for a Null argument */
    switch (id) {
    case F_NZ: case F_ISNULL: case F_ISNUMERIC: case F_ISDATE: case F_ISEMPTY: case F_CAT: case F_DATE: case F_TIME:
    case F_NOW: case F_TIMER: case F_RND: case F_FORMAT: case F_STRCOMP: case F_INSTR: case F_ADD: case F_ERROR:
    case F_XOR: case F_EQV: case F_IMP: case F_CVAR:
        break;
    default:
        if (any_null(argc, argv)) { r_ = vb_v_null(); END_FN; return; }
    }
    switch (id) {
    case F_UCASE: r_ = vb_v_s(vb_ucase(vb_v2s(A(argv, 0)))); break;
    case F_LCASE: r_ = vb_v_s(vb_lcase(vb_v2s(A(argv, 0)))); break;
    case F_LEFT: r_ = vb_v_s(vb_left(vb_v2s(A(argv, 0)), vb_v2l(A(argv, 1)))); break;
    case F_RIGHT: r_ = vb_v_s(vb_right(vb_v2s(A(argv, 0)), vb_v2l(A(argv, 1)))); break;
    case F_MID: r_ = vb_v_s(vb_mid(vb_v2s(A(argv, 0)), vb_v2l(A(argv, 1)), argc > 2 ? vb_v2l(A(argv, 2)) : -1)); break;
    case F_LEN: r_ = vb_v_l(vb_len(vb_v2s(A(argv, 0)))); break;
    case F_TRIM: r_ = vb_v_s(vb_trim(vb_v2s(A(argv, 0)))); break;
    case F_LTRIM: r_ = vb_v_s(vb_ltrim(vb_v2s(A(argv, 0)))); break;
    case F_RTRIM: r_ = vb_v_s(vb_rtrim(vb_v2s(A(argv, 0)))); break;
    case F_INSTR:
        if (any_null(argc, argv)) { r_ = vb_v_null(); break; }
        if (argc >= 3 && sqlite3_value_type(argv[0]) != SQLITE_TEXT)
            r_ = vb_v_l(vb_instr(vb_v2l(A(argv, 0)), vb_v2s(A(argv, 1)), vb_v2s(A(argv, 2)), argc > 3 ? vb_v2l(A(argv, 3)) : 1));
        else r_ = vb_v_l(vb_instr(1, vb_v2s(A(argv, 0)), vb_v2s(A(argv, 1)), 1));
        break;
    case F_INSTRREV: r_ = vb_v_l(vb_instrrev(vb_v2s(A(argv, 0)), vb_v2s(A(argv, 1)), opt_l(argc, argv, 2, -1), 1)); break;
    case F_SPACE: r_ = vb_v_s(vb_space(vb_v2l(A(argv, 0)))); break;
    case F_STRING:
        if (sqlite3_value_type(argv[1]) == SQLITE_TEXT) r_ = vb_v_s(vb_string_s(vb_v2l(A(argv, 0)), vb_v2s(A(argv, 1))));
        else r_ = vb_v_s(vb_string_n(vb_v2l(A(argv, 0)), vb_v2l(A(argv, 1))));
        break;
    case F_ASC: r_ = vb_v_i((vb_i16)vb_asc(vb_v2s(A(argv, 0)))); break;
    case F_CHR: r_ = vb_v_s(vb_chr(vb_v2l(A(argv, 0)))); break;
    case F_STRREVERSE: r_ = vb_v_s(vb_strreverse(vb_v2s(A(argv, 0)))); break;
    case F_REPLACE: r_ = vb_v_s(vb_replace(vb_v2s(A(argv, 0)), vb_v2s(A(argv, 1)), vb_v2s(A(argv, 2)), opt_l(argc, argv, 3, 1), opt_l(argc, argv, 4, -1), 1)); break;
    case F_STRCOMP:
        if (sqlite3_value_type(argv[0]) == SQLITE_NULL || sqlite3_value_type(argv[1]) == SQLITE_NULL) { r_ = vb_v_null(); break; }
        r_ = vb_v_l(vb_strcomp(vb_v2s(A(argv, 0)), vb_v2s(A(argv, 1)), opt_l(argc, argv, 2, 1))); break;
    case F_STRCONV: r_ = vb_v_s(vb_strconv(vb_v2s(A(argv, 0)), vb_v2l(A(argv, 1)))); break;
    case F_VAL: r_ = vb_v_d(vb_val(vb_v2s(A(argv, 0)))); break;
    case F_STR: {
        vbvar a = A(argv, 0);
        r_ = vb_v_s(a.vt == VT_I4 ? vb_str_i(a.u.l) : vb_str_d(vb_v2d(a)));
        break; }
    case F_CSTR: r_ = vb_v_s(vb_v2s(A(argv, 0))); break;
    case F_CINT: r_ = vb_v_i(vb_v2i(A(argv, 0))); break;
    case F_CLNG: r_ = vb_v_l(vb_v2l(A(argv, 0))); break;
    case F_CDBL: r_ = vb_v_d(vb_v2d(A(argv, 0))); break;
    case F_CSNG: r_ = vb_v_f(vb_v2f(A(argv, 0))); break;
    case F_CCUR: r_ = vb_v_c(vb_v2c(A(argv, 0))); break;
    case F_CBYTE: r_ = vb_v_byte(vb_v2byte(A(argv, 0))); break;
    case F_CBOOL: r_ = vb_v_bool(vb_v2bool(A(argv, 0))); break;
    case F_CDATE: r_ = vb_v_date(vb_v2date(A(argv, 0))); break;
    case F_CVAR: r_ = A(argv, 0); break;
    case F_INT: r_ = vb_v_d(vb_int(vb_v2d(A(argv, 0)))); if (sqlite3_value_type(argv[0]) == SQLITE_INTEGER) r_ = A(argv, 0); break;
    case F_FIX: r_ = vb_v_d(vb_fix(vb_v2d(A(argv, 0)))); if (sqlite3_value_type(argv[0]) == SQLITE_INTEGER) r_ = A(argv, 0); break;
    case F_ABS: r_ = sqlite3_value_type(argv[0]) == SQLITE_INTEGER ? vb_v_d(fabs((double)sqlite3_value_int64(argv[0]))) : vb_v_d(fabs(vb_v2d(A(argv, 0))));
        if (sqlite3_value_type(argv[0]) == SQLITE_INTEGER) r_ = vb_v_l(abs((int)sqlite3_value_int64(argv[0])));
        break;
    case F_SGN: r_ = vb_v_i((vb_i16)vb_sgn(vb_v2d(A(argv, 0)))); break;
    case F_SQR: r_ = vb_v_d(vb_sqr(vb_v2d(A(argv, 0)))); break;
    case F_EXP: r_ = vb_v_d(vb_exp(vb_v2d(A(argv, 0)))); break;
    case F_LOG: r_ = vb_v_d(vb_log(vb_v2d(A(argv, 0)))); break;
    case F_SIN: r_ = vb_v_d(vb_sin(vb_v2d(A(argv, 0)))); break;
    case F_COS: r_ = vb_v_d(vb_cos(vb_v2d(A(argv, 0)))); break;
    case F_TAN: r_ = vb_v_d(vb_tan(vb_v2d(A(argv, 0)))); break;
    case F_ATN: r_ = vb_v_d(vb_atn(vb_v2d(A(argv, 0)))); break;
    case F_RND: r_ = vb_v_f((float)vb_rnd(argc ? vb_v2d(A(argv, 0)) : 1, argc > 0)); break;
    case F_ROUND: r_ = vb_v_d(vb_round(vb_v2d(A(argv, 0)), argc > 1 ? vb_v2l(A(argv, 1)) : 0)); break;
    case F_HEX: r_ = vb_v_s(vb_hex(vb_v2l(A(argv, 0)), 4)); break;
    case F_OCT: r_ = vb_v_s(vb_oct(vb_v2l(A(argv, 0)), 4)); break;
    case F_YEAR: r_ = vb_v_i((vb_i16)vb_year(vb_v2date(A(argv, 0)))); break;
    case F_MONTH: r_ = vb_v_i((vb_i16)vb_month(vb_v2date(A(argv, 0)))); break;
    case F_DAY: r_ = vb_v_i((vb_i16)vb_day(vb_v2date(A(argv, 0)))); break;
    case F_HOUR: r_ = vb_v_i((vb_i16)vb_hour(vb_v2date(A(argv, 0)))); break;
    case F_MINUTE: r_ = vb_v_i((vb_i16)vb_minute(vb_v2date(A(argv, 0)))); break;
    case F_SECOND: r_ = vb_v_i((vb_i16)vb_second(vb_v2date(A(argv, 0)))); break;
    case F_WEEKDAY: r_ = vb_v_i((vb_i16)vb_weekday(vb_v2date(A(argv, 0)), opt_l(argc, argv, 1, 1))); break;
    case F_DATE: r_ = vb_v_date(floor(vb_now())); break;
    case F_TIME: { vb_date n = vb_now(); r_ = vb_v_date(n - floor(n)); break; }
    case F_NOW: r_ = vb_v_date(vb_now()); break;
    case F_TIMER: r_ = vb_v_f((float)vb_timer()); break;
    case F_DATEADD: r_ = vb_v_date(vb_dateadd(vb_v2s(A(argv, 0)), vb_v2d(A(argv, 1)), vb_v2date(A(argv, 2)))); break;
    case F_DATEDIFF: r_ = vb_v_l(vb_datediff(vb_v2s(A(argv, 0)), vb_v2date(A(argv, 1)), vb_v2date(A(argv, 2)), opt_l(argc, argv, 3, 1))); break;
    case F_DATEPART: r_ = vb_v_l(vb_datepart(vb_v2s(A(argv, 0)), vb_v2date(A(argv, 1)), opt_l(argc, argv, 2, 1), opt_l(argc, argv, 3, 1))); break;
    case F_DATESERIAL: r_ = vb_v_date(vb_dateserial(vb_v2l(A(argv, 0)), vb_v2l(A(argv, 1)), vb_v2l(A(argv, 2)))); break;
    case F_TIMESERIAL: r_ = vb_v_date(vb_timeserial(vb_v2l(A(argv, 0)), vb_v2l(A(argv, 1)), vb_v2l(A(argv, 2)))); break;
    case F_DATEVALUE: r_ = vb_v_date(vb_datevalue(vb_v2s(A(argv, 0)))); break;
    case F_TIMEVALUE: r_ = vb_v_date(vb_timevalue(vb_v2s(A(argv, 0)))); break;
    case F_FORMAT: {
        vbvar a = A(argv, 0);
        if (a.vt == VT_NULL) { r_ = vb_v_null(); break; }
        if (argc > 1 && sqlite3_value_type(argv[1]) != SQLITE_NULL)
            r_ = vb_v_s(vb_format(a, vb_v2s(A(argv, 1)), 1, opt_l(argc, argv, 2, 1), opt_l(argc, argv, 3, 1)));
        else r_ = vb_v_s(vb_format(a, vb_sfromutf8(""), 0, 1, 1));
        break; }
    case F_FORMATNUMBER: case F_FORMATCURRENCY: case F_FORMATPERCENT:
        r_ = vb_v_s(vb_formatnumber(A(argv, 0), opt_l(argc, argv, 1, -1), opt_l(argc, argv, 2, -2), opt_l(argc, argv, 3, -2),
                                    opt_l(argc, argv, 4, -2), id == F_FORMATNUMBER ? 0 : id == F_FORMATCURRENCY ? 1 : 2));
        break;
    case F_FORMATDATETIME: r_ = vb_v_s(vb_formatdatetime(vb_v2date(A(argv, 0)), opt_l(argc, argv, 1, 0))); break;
    case F_MONTHNAME: r_ = vb_v_s(vb_monthname(vb_v2l(A(argv, 0)), opt_l(argc, argv, 1, 0))); break;
    case F_WEEKDAYNAME: r_ = vb_v_s(vb_weekdayname(vb_v2l(A(argv, 0)), opt_l(argc, argv, 1, 0), opt_l(argc, argv, 2, 1))); break;
    case F_NZ:
        r_ = sqlite3_value_type(argv[0]) != SQLITE_NULL ? A(argv, 0) : argc > 1 ? A(argv, 1) : vb_v_l(0);
        if (sqlite3_value_type(argv[0]) == SQLITE_NULL && argc == 1) r_ = vb_v_s(vb_sfromutf8(""));
        break;
    case F_ISNULL: r_ = vb_v_bool(sqlite3_value_type(argv[0]) == SQLITE_NULL ? -1 : 0); break;
    case F_ISEMPTY: r_ = vb_v_bool(0); break;
    case F_ISNUMERIC: {
        int t = sqlite3_value_type(argv[0]);
        if (t == SQLITE_INTEGER || t == SQLITE_FLOAT) { r_ = vb_v_bool(-1); break; }
        if (t != SQLITE_TEXT) { r_ = vb_v_bool(0); break; }
        vbvar a = A(argv, 0);
        if (a.vt == VT_DATE) { r_ = vb_v_bool(0); break; }
        vb_i32 save = vb_errpend;
        vb_s2d(a.u.s);
        int ok = !(vb_errpend && !save);
        if (!ok) vb_err_clear_pending();
        r_ = vb_v_bool(ok ? -1 : 0);
        break; }
    case F_ISDATE: {
        vbvar a = A(argv, 0);
        if (a.vt == VT_DATE) { r_ = vb_v_bool(-1); break; }
        if (a.vt != VT_BSTR) { r_ = vb_v_bool(0); break; }
        vb_i32 save = vb_errpend;
        vb_s2date(a.u.s);
        int ok = !(vb_errpend && !save);
        if (!ok) vb_err_clear_pending();
        r_ = vb_v_bool(ok ? -1 : 0);
        break; }
    case F_ERROR: r_ = vb_v_s(vb_error_str(argc ? vb_v2l(A(argv, 0)) : 0)); break;
    /* the operators */
    case F_CAT: {
        int n0 = sqlite3_value_type(argv[0]) == SQLITE_NULL, n1 = sqlite3_value_type(argv[1]) == SQLITE_NULL;
        if (n0 && n1) { r_ = vb_v_null(); break; }
        vbs a = n0 ? vb_sfromutf8("") : vb_v2s(A(argv, 0)), b = n1 ? vb_sfromutf8("") : vb_v2s(A(argv, 1));
        r_ = vb_v_s(vb_scat(a, b));
        break; }
    case F_ADD:
        if (any_null(argc, argv)) { r_ = vb_v_null(); break; }
        r_ = vb_v_add(A(argv, 0), A(argv, 1)); break;
    case F_SUB: r_ = vb_v_sub(A(argv, 0), A(argv, 1)); break;
    case F_MUL: r_ = vb_v_mul(A(argv, 0), A(argv, 1)); break;
    case F_DIV: r_ = vb_v_div(A(argv, 0), A(argv, 1)); break;
    case F_IDIV: r_ = vb_v_idiv(A(argv, 0), A(argv, 1)); break;
    case F_MOD: r_ = vb_v_mod(A(argv, 0), A(argv, 1)); break;
    case F_POW: r_ = vb_v_pow(A(argv, 0), A(argv, 1)); break;
    case F_NEG: r_ = vb_v_neg(A(argv, 0)); break;
    case F_XOR: r_ = vb_v_xor(A(argv, 0), A(argv, 1)); break;
    case F_EQV: r_ = vb_v_eqv(A(argv, 0), A(argv, 1)); break;
    case F_IMP: r_ = vb_v_imp(A(argv, 0), A(argv, 1)); break;
    case F_NUM: {                               /* Sum/Avg of text that holds numbers, of dates */
        vbvar a = A(argv, 0);
        r_ = a.vt == VT_BSTR ? vb_v_d(vb_v2d(a)) : a.vt == VT_DATE ? vb_v_d(a.u.d) : a;
        break; }
    }
    END_FN;
}

/* Jet LIKE: * ? # [abc] [a-z] [!x], without case; ansi: % _ */
static int wlower(int c) { return (int)towlower((wint_t)c); }
static int like_match(const vb_wchar *p, int pn, const vb_wchar *s, int sn, int ansi)
{
    int pi = 0, si = 0;
    while (pi < pn) {
        vb_wchar c = p[pi];
        if ((!ansi && c == '*') || (ansi && c == '%')) {
            while (pi < pn && ((!ansi && p[pi] == '*') || (ansi && p[pi] == '%'))) pi++;
            if (pi == pn) return 1;
            for (int k = si; k <= sn; k++) if (like_match(p + pi, pn - pi, s + k, sn - k, ansi)) return 1;
            return 0;
        }
        if (si >= sn) return 0;
        if ((!ansi && c == '?') || (ansi && c == '_')) { pi++; si++; continue; }
        if (!ansi && c == '#') { if (s[si] < '0' || s[si] > '9') return 0; pi++; si++; continue; }
        if (c == '[') {
            int e = pi + 1;
            while (e < pn && p[e] != ']') e++;
            if (e >= pn) return 0;
            int neg = 0, k = pi + 1, ok = 0;
            if (k < e && (p[k] == '!' || (ansi && p[k] == '^'))) { neg = 1; k++; }
            if (e == k) { if (s[si] != '[') return 0; pi = e + 1; si++; continue; }   /* [] */
            int sc = wlower(s[si]);
            for (; k < e; k++) {
                if (k + 2 < e && p[k + 1] == '-') {
                    int a = wlower(p[k]), b = wlower(p[k + 2]);
                    if (sc >= a && sc <= b) ok = 1;
                    k += 2;
                } else if (wlower(p[k]) == sc) ok = 1;
            }
            if (ok == neg) return 0;
            pi = e + 1; si++;
            continue;
        }
        if (wlower(c) != wlower(s[si])) return 0;
        pi++; si++;
    }
    return si == sn;
}
static void fn_like(sqlite3_context *cx, int argc, sqlite3_value **argv)
{
    if (any_null(2, argv)) { sqlite3_result_null(cx); return; }
    vb_i32 m = vb_mark();
    vbs s = vb_v2s(A(argv, 0)), p = vb_v2s(A(argv, 1));
    int ansi = argc > 2 && sqlite3_value_int(argv[2]);
    sqlite3_result_int(cx, like_match(p.p, p.len, s.p, s.len, ansi));
    vb_release(m);
}

/* the weight of a character in Jet's order (Windows word sort): space, punctuation, digits, letters */
static int jet_weight(int c)
{
    if (c == ' ') return 1;
    if (c < 0x80) {
        if (isdigit(c)) return 0x200 + c;
        if (isalpha(c)) return 0x300 + tolower(c);
        return 0x100 + c;                       /* punctuation and symbols */
    }
    return 0x400 + wlower(c);
}

/* the text order of Jet ("General"): without case; the hyphen and the apostrophe do not count */
/* the comparison of texts as Jet orders them (ORDER BY, =, <, its indexes): by the keys of the sort order of the
 * database (jetsort.c; Jet 4 in its LCID, Jet 3 General); plain letters, digits and spaces in General directly
 * (VB6 oracle tests/db/d_jetsort) */
typedef struct jcoll { int lcid, jet4, cp; } jcoll;
static jcoll coll_general = { 0x409, 1, 1252 };
static int plain_text(const unsigned char *x, int n)
{
    for (int i = 0; i < n; i++) if (!isalnum(x[i]) && x[i] != ' ') return 0;
    return 1;
}
static int plain_weight(int c) { return c == ' ' ? 0 : isdigit(c) ? 1 + (c - '0') : 20 + (tolower(c) - 'a'); }
static void coll_key(jcoll *co, jkey *k, const unsigned char *x, int n)
{
    uint16_t *w = malloc(sizeof(uint16_t) * (size_t)(n + 1));
    int m = 0;
    for (int i = 0; i < n; ) {                  /* UTF-8 -> UTF-16 */
        unsigned c = x[i];
        int l = c < 0x80 ? 1 : c < 0xE0 ? 2 : c < 0xF0 ? 3 : 4;
        if (i + l > n) l = n - i;
        if (l == 2) c = (c & 0x1F) << 6 | (x[i + 1] & 0x3F);
        else if (l == 3) c = (c & 0x0F) << 12 | (x[i + 1] & 0x3F) << 6 | (x[i + 2] & 0x3F);
        else if (l == 4) { c = (c & 0x07) << 18 | (x[i + 1] & 0x3F) << 12 | (x[i + 2] & 0x3F) << 6 | (x[i + 3] & 0x3F); c = 0xFFFD; }
        w[m++] = (uint16_t)c;
        i += l;
    }
    if (co->jet4) {
        jet_sort_lcid(co->lcid);
        jet_text_key(k, (const unsigned char *)w, m * 2, 1);
        jet_sort_lcid(0x409);
    } else {
        unsigned char *b = malloc((size_t)m * 2 + 1);
        int bn = jet_utf16_to_cp(co->cp, w, m, b);
        jet_text_key(k, b, bn, 0);
        free(b);
    }
    free(w);
}
static int jet_collate_old(const unsigned char *x, int an, const unsigned char *y, int bn);
static int jet_collate(void *u, int an, const void *a, int bn, const void *b)
{
    jcoll *co = u ? u : &coll_general;
    const unsigned char *x = a, *y = b;
    if (an && bn && x[0] == JDATE_MARK && y[0] == JDATE_MARK) return jet_collate_old(x, an, y, bn);
    if (co->jet4 && co->lcid == 0x409 && plain_text(x, an) && plain_text(y, bn)) {
        while (an > 0 && x[an - 1] == ' ') an--;
        while (bn > 0 && y[bn - 1] == ' ') bn--;
        for (int i = 0; i < an && i < bn; i++) {
            int wx = plain_weight(x[i]), wy = plain_weight(y[i]);
            if (wx != wy) return wx < wy ? -1 : 1;
        }
        return an == bn ? 0 : an < bn ? -1 : 1;
    }
    jkey ka = { 0 }, kb = { 0 };
    coll_key(co, &ka, x, an);
    coll_key(co, &kb, y, bn);
    int n = ka.n < kb.n ? ka.n : kb.n, r = memcmp(ka.p, kb.p, (size_t)n);
    if (!r) r = ka.n == kb.n ? 0 : ka.n < kb.n ? -1 : 1;
    free(ka.p); free(kb.p);
    return r < 0 ? -1 : r > 0;
}
/* two texts (UTF-8) compared as Jet orders them in a sort order (a table's index: Seek, MoveNext) */
int jet_compare_text(int lcid, int jet4, const char *a, const char *b)
{
    jcoll co = { lcid ? lcid : 0x409, jet4, 1252 };
    return jet_collate(&co, (int)strlen(a), a, (int)strlen(b), b);
}
/* the sort order of a database's connection: its LCID (Jet 4), Jet 3 */
void jet_collation(sqlite3 *db, int lcid, int jet4, int cp)
{
    jcoll *co = malloc(sizeof *co);
    co->lcid = lcid ? lcid : 0x409; co->jet4 = jet4; co->cp = cp ? cp : 1252;
    sqlite3_create_collation_v2(db, "NOCASE", SQLITE_UTF8, co, jet_collate, free);
}
static int jet_collate_old(const unsigned char *x, int an, const unsigned char *y, int bn)
{
    int i = 0, j = 0;
    while (an > 0 && x[an - 1] == ' ') an--;    /* Jet: trailing spaces do not count ('a ' = 'a') */
    while (bn > 0 && y[bn - 1] == ' ') bn--;
    for (;;) {
        while (i < an && (x[i] == '-' || x[i] == '\'')) i++;
        while (j < bn && (y[j] == '-' || y[j] == '\'')) j++;
        if (i >= an || j >= bn) break;
        /* UTF-8 code points */
        int cx = x[i], cy = y[j], lx = 1, ly = 1;
        if (cx >= 0xC0 && i + 1 < an) { if (cx < 0xE0) { cx = ((cx & 0x1F) << 6) | (x[i + 1] & 0x3F); lx = 2; } else if (i + 2 < an) { cx = ((cx & 0x0F) << 12) | ((x[i + 1] & 0x3F) << 6) | (x[i + 2] & 0x3F); lx = 3; } }
        if (cy >= 0xC0 && j + 1 < bn) { if (cy < 0xE0) { cy = ((cy & 0x1F) << 6) | (y[j + 1] & 0x3F); ly = 2; } else if (j + 2 < bn) { cy = ((cy & 0x0F) << 12) | ((y[j + 1] & 0x3F) << 6) | (y[j + 2] & 0x3F); ly = 3; } }
        if (cx == JDATE_MARK && cy == JDATE_MARK) { i++; j++; continue; }
        int lx2 = jet_weight(cx), ly2 = jet_weight(cy);
        if (lx2 != ly2) return lx2 < ly2 ? -1 : 1;
        i += lx; j += ly;
    }
    if (i < an) return 1;
    if (j < bn) return -1;
    /* equal without case (Jet: 'apple' = 'APPLE'); with a hyphen or an apostrophe more: after */
    if (an != bn) return an < bn ? -1 : 1;
    return 0;
}

/* First, Last, StDev, Var */
typedef struct aggst { int n; double s, s2; sqlite3_value *first, *last; } aggst;
static void agg_step(sqlite3_context *cx, int argc, sqlite3_value **argv)
{
    (void)argc;
    aggst *a = sqlite3_aggregate_context(cx, sizeof(aggst));
    if (!a) return;
    int id = fn_id(cx);
    if (id <= 1) {                              /* First / Last: Nulls count */
        if (id == 0 && !a->first) a->first = sqlite3_value_dup(argv[0]);
        if (id == 1) { sqlite3_value_free(a->last); a->last = sqlite3_value_dup(argv[0]); }
        a->n++;
        return;
    }
    if (sqlite3_value_type(argv[0]) == SQLITE_NULL) return;
    vb_i32 m = vb_mark();
    double d = vb_v2d(A(argv, 0));
    vb_release(m);
    a->n++; a->s += d; a->s2 += d * d;
}
static void agg_final(sqlite3_context *cx)
{
    aggst *a = sqlite3_aggregate_context(cx, 0);
    int id = fn_id(cx);
    if (!a || !a->n) { sqlite3_result_null(cx); return; }
    if (id == 0) { sqlite3_result_value(cx, a->first); sqlite3_value_free(a->first); return; }
    if (id == 1) { sqlite3_result_value(cx, a->last); sqlite3_value_free(a->last); return; }
    int pop = id == 3 || id == 5;
    if (!pop && a->n < 2) { sqlite3_result_null(cx); return; }
    double mean = a->s / a->n, var = (a->s2 - a->n * mean * mean) / (pop ? a->n : a->n - 1);
    if (var < 0) var = 0;
    sqlite3_result_double(cx, id >= 4 ? var : sqrt(var));
}

void jet_register(sqlite3 *db)
{
    static const struct { const char *n; int id, argc; } F[] = {
        { "ucase", F_UCASE, 1 }, { "lcase", F_LCASE, 1 }, { "left", F_LEFT, 2 }, { "right", F_RIGHT, 2 }, { "mid", F_MID, -1 },
        { "len", F_LEN, 1 }, { "trim", F_TRIM, 1 }, { "ltrim", F_LTRIM, 1 }, { "rtrim", F_RTRIM, 1 }, { "instr", F_INSTR, -1 },
        { "instrrev", F_INSTRREV, -1 }, { "space", F_SPACE, 1 }, { "string", F_STRING, 2 }, { "asc", F_ASC, 1 },
        { "chr", F_CHR, 1 }, { "strreverse", F_STRREVERSE, 1 }, { "replace", F_REPLACE, -1 }, { "strcomp", F_STRCOMP, -1 },
        { "strconv", F_STRCONV, -1 }, { "val", F_VAL, 1 }, { "str", F_STR, 1 }, { "cstr", F_CSTR, 1 }, { "cint", F_CINT, 1 },
        { "clng", F_CLNG, 1 }, { "cdbl", F_CDBL, 1 }, { "csng", F_CSNG, 1 }, { "ccur", F_CCUR, 1 }, { "cbyte", F_CBYTE, 1 },
        { "cbool", F_CBOOL, 1 }, { "cdate", F_CDATE, 1 }, { "cvdate", F_CDATE, 1 }, { "cvar", F_CVAR, 1 }, { "int", F_INT, 1 },
        { "fix", F_FIX, 1 }, { "abs", F_ABS, 1 }, { "sgn", F_SGN, 1 }, { "sqr", F_SQR, 1 }, { "exp", F_EXP, 1 },
        { "log", F_LOG, 1 }, { "sin", F_SIN, 1 }, { "cos", F_COS, 1 }, { "tan", F_TAN, 1 }, { "atn", F_ATN, 1 },
        { "rnd", F_RND, -1 }, { "round", F_ROUND, -1 }, { "hex", F_HEX, 1 }, { "oct", F_OCT, 1 }, { "year", F_YEAR, 1 },
        { "month", F_MONTH, 1 }, { "day", F_DAY, 1 }, { "hour", F_HOUR, 1 }, { "minute", F_MINUTE, 1 },
        { "second", F_SECOND, 1 }, { "weekday", F_WEEKDAY, -1 }, { "date", F_DATE, 0 }, { "time", F_TIME, 0 },
        { "now", F_NOW, 0 }, { "timer", F_TIMER, 0 }, { "dateadd", F_DATEADD, 3 }, { "datediff", F_DATEDIFF, -1 },
        { "datepart", F_DATEPART, -1 }, { "dateserial", F_DATESERIAL, 3 }, { "timeserial", F_TIMESERIAL, 3 },
        { "datevalue", F_DATEVALUE, 1 }, { "timevalue", F_TIMEVALUE, 1 }, { "format", F_FORMAT, -1 }, { "nz", F_NZ, -1 },
        { "isnull", F_ISNULL, 1 }, { "isnumeric", F_ISNUMERIC, 1 }, { "isdate", F_ISDATE, 1 }, { "isempty", F_ISEMPTY, 1 },
        { "monthname", F_MONTHNAME, -1 }, { "weekdayname", F_WEEKDAYNAME, -1 }, { "formatnumber", F_FORMATNUMBER, -1 },
        { "formatcurrency", F_FORMATCURRENCY, -1 }, { "formatpercent", F_FORMATPERCENT, -1 },
        { "formatdatetime", F_FORMATDATETIME, -1 }, { "error", F_ERROR, -1 },
        { "cat", F_CAT, 2 }, { "add", F_ADD, 2 }, { "sub", F_SUB, 2 }, { "mul", F_MUL, 2 }, { "div", F_DIV, 2 },
        { "idiv", F_IDIV, 2 }, { "mod", F_MOD, 2 }, { "pow", F_POW, 2 }, { "neg", F_NEG, 1 }, { "xor", F_XOR, 2 },
        { "eqv", F_EQV, 2 }, { "imp", F_IMP, 2 }, { "num", F_NUM, 1 },
    };
    char name[48];
    for (size_t i = 0; i < sizeof F / sizeof *F; i++) {
        snprintf(name, sizeof name, "jet_%s", F[i].n);
        int flags = SQLITE_UTF8 | (F[i].id == F_DATE || F[i].id == F_TIME || F[i].id == F_NOW || F[i].id == F_RND || F[i].id == F_TIMER ? 0 : SQLITE_DETERMINISTIC);
        sqlite3_create_function(db, name, F[i].argc, flags, (void *)(intptr_t)F[i].id, fn_vba, 0, 0);
    }
    sqlite3_create_function(db, "jet_like", 3, SQLITE_UTF8 | SQLITE_DETERMINISTIC, 0, fn_like, 0, 0);
    static const char *const AG[] = { "jet_first", "jet_last", "jet_stdev", "jet_stdevp", "jet_var", "jet_varp" };
    for (int i = 0; i < 6; i++) sqlite3_create_function(db, AG[i], 1, SQLITE_UTF8, (void *)(intptr_t)i, 0, agg_step, agg_final);
    sqlite3_create_collation(db, "NOCASE", SQLITE_UTF8, 0, jet_collate);
}

/* a VBA function of Jet SQL (Jet writes an alias with such a name in brackets: AS [Year]) */
int jet_is_function(const char *name)
{
    return find_name(FUNCS, name, (int)strlen(name)) >= 0 || find_name(AGGS, name, (int)strlen(name)) >= 0;
}
