/* jetquery.c - a saved query of Jet both ways. MSysQueries keeps a query as rows by attribute: 0 the start, 1 the
 * kind (Flag: 1 select, 2 make-table, 3 append, 4 update, 5 delete, 6 crosstab, 7 DDL, 8 pass-through, 9 union;
 * Name1 the target table; Expression the text of a DDL), 2 a parameter (Name1, Flag its type), 3 the predicate
 * (Flag: 1 *, 2 DISTINCT, 4 OWNERACCESS, 8 DISTINCTROW, 0x10 TOP, 0x20 PERCENT; Name1 the number), 5 a table
 * (Name1, Name2 its alias) or a part of a union / a subquery (Expression), 6 a column (Expression, Name1 its alias,
 * Name2 the target column; Flag in a crosstab: 0 the value, 1 the pivot, 2 a row heading), 7 a join (Name1, Name2,
 * Expression the condition, Flag 1 inner, 2 left, 3 right), 8 WHERE, 9 GROUP BY, 10 HAVING, 11 ORDER BY (Name1
 * "D": descending), 255 the end; Order counts the rows of each attribute (4 bytes, big-endian).
 * jq_text writes the SQL DAO shows for them; jq_compile turns Jet SQL into them. The Python prototype:
 * tools/jet/jetquery.py (checked on the 413 queries of the PSC files and by Jet itself, JetDump).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "jetquery.h"
#include <ctype.h>
#include <stdarg.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <strings.h>

/* ---------------------------------------------------------------- text */
typedef struct sb { char *p; int n, cap; } sb;
static void sput(sb *b, const char *s, int n)
{
    if (n < 0) n = (int)strlen(s);
    if (b->n + n + 1 > b->cap) { b->cap = (b->n + n + 1) * 2 + 64; b->p = realloc(b->p, (size_t)b->cap); }
    memcpy(b->p + b->n, s, (size_t)n);
    b->n += n;
    b->p[b->n] = 0;
}
static void ss(sb *b, const char *s) { sput(b, s ? s : "", -1); }
static char *stake(sb *b) { if (!b->p) return strdup(""); char *p = b->p; b->p = 0; b->n = b->cap = 0; return p; }

static int needs_brackets(const char *n)
{
    if (!*n || isdigit((unsigned char)*n)) return 1;
    for (const char *p = n; *p; p++) if (!isalnum((unsigned char)*p) && *p != '_' && (unsigned char)*p < 0x80) return 1;
    return 0;
}
static void put_name(sb *b, const char *n)
{
    if (needs_brackets(n)) { ss(b, "["); ss(b, n); ss(b, "]"); } else ss(b, n);
}
static int is_function(const char *n)
{
    static const char *const F[] = { "count", "sum", "avg", "min", "max", "first", "last", "stdev", "stdevp", "var",
        "varp", "year", "month", "day", "date", "time", "now", "format", "left", "right", "mid", "len", "val", "str",
        "int", "fix", "abs", "iif", "nz", "trim", "ucase", "lcase", "weekday", "hour", "minute", "second", "datepart",
        "datediff", "dateadd", "dateserial", "timeserial", "datevalue", "timevalue", "instr", "space", "string", "chr",
        "asc", "sgn", "sqr", "round", "ccur", "cdbl", "clng", "cint", "csng", "cstr", "cdate", "cbool", "cbyte",
        "cvar", 0 };
    for (int i = 0; F[i]; i++) if (!strcasecmp(F[i], n)) return 1;
    return 0;
}
static void put_alias(sb *b, const char *n)
{
    if (is_function(n)) { ss(b, "["); ss(b, n); ss(b, "]"); } else put_name(b, n);
}
static const char *type_name(int t)
{
    switch (t) {
    case 1: return "Bit"; case 2: return "Byte"; case 3: return "Short"; case 4: return "Long";
    case 5: return "Currency"; case 6: return "IEEESingle"; case 7: return "IEEEDouble"; case 8: return "DateTime";
    case 9: return "Binary"; case 10: return "Text ( 255 )"; case 11: return "LongBinary"; case 12: return "LongText";
    case 15: return "Guid";
    }
    return "Value";
}

/* ---------------------------------------------------------------- rows -> SQL */
typedef struct src { char *name, *sql; struct src *l, *r; int jt; sb on; int non; } src;
static src *src_table(const char *t, const char *alias)
{
    src *s = calloc(1, sizeof *s);
    sb b = { 0 };
    put_name(&b, t);
    if (alias && *alias) { ss(&b, " AS "); put_name(&b, alias); s->name = strdup(alias); }
    else s->name = strdup(t);
    s->sql = stake(&b);
    return s;
}
static int src_has(src *s, const char *n)
{
    if (s->name) return !strcasecmp(s->name, n);
    return src_has(s->l, n) || src_has(s->r, n);
}
static void src_sql(src *s, sb *b, int nested)
{
    if (s->name) { ss(b, s->sql); return; }
    if (nested) ss(b, "(");
    src_sql(s->l, b, 1);
    ss(b, s->jt == 2 ? " LEFT JOIN " : s->jt == 3 ? " RIGHT JOIN " : " INNER JOIN ");
    src_sql(s->r, b, 1);
    ss(b, " ON ");
    ss(b, s->on.p);
    if (nested) ss(b, ")");
}
static void src_free(src *s)
{
    if (!s) return;
    src_free(s->l); src_free(s->r);
    free(s->name); free(s->sql); free(s->on.p);
    free(s);
}
/* the conditions of one join, kept as rows of their own: Jet writes the later one first, "(b) AND (a)" (DEV-14/q1) */
static void add_on(src *s, const char *e)
{
    if (s->non == 0) ss(&s->on, e);
    else {
        sb n = { 0 };
        ss(&n, "("); ss(&n, e); ss(&n, ") AND ");
        if (s->non == 1) { ss(&n, "("); ss(&n, s->on.p); ss(&n, ")"); } else ss(&n, s->on.p);
        free(s->on.p);
        s->on = n;
    }
    s->non++;
}
static void list_add(sb *b, const char *s) { if (b->n) ss(b, ", "); ss(b, s); }
static int order_cmp(const void *a, const void *b)
{
    const jqrow *x = *(const jqrow *const *)a, *y = *(const jqrow *const *)b;
    return x->order < y->order ? -1 : x->order > y->order;
}
#define S(x) ((x) ? (x) : "")

char *jq_text(const jqrow *rows, int n, int *qtype)
{
    int kind = 1, flags = 0;
    const char *target = "", *top = "";
    sb params = { 0 }, cols = { 0 }, tcols = { 0 }, where = { 0 }, group = { 0 }, having = { 0 }, order = { 0 };
    sb xvalue = { 0 }, xrows = { 0 }, xpivot = { 0 }, xgroup = { 0 }, sets = { 0 };
    src **srcs = calloc((size_t)n + 1, sizeof(src *));
    const char **unions = calloc((size_t)n + 1, sizeof(char *));
    const jqrow **joins = calloc((size_t)n + 1, sizeof(jqrow *));
    int ns = 0, nu = 0, nj = 0;
    const jqrow **ord = malloc(sizeof(jqrow *) * (size_t)(n + 1));
    for (int i = 0; i < n; i++) ord[i] = &rows[i];
    /* by Order (the rows of each attribute in their order; a stable sort) */
    for (int i = 1; i < n; i++) { const jqrow *x = ord[i]; int j = i; while (j > 0 && order_cmp(&ord[j - 1], &x) > 0) { ord[j] = ord[j - 1]; j--; } ord[j] = x; }
    *qtype = 0;
    for (int k = 0; k < n; k++) {
        const jqrow *r = ord[k];
        int flag = r->flag == JQ_NOFLAG ? 0 : r->flag;
        const char *e = S(r->expr), *n1 = S(r->name1), *n2 = S(r->name2);
        switch (r->attr) {
        case 1:
            kind = flag; target = n1;
            if (kind == 7) {
                char *s = strdup(e);
                free(srcs); free(unions); free(joins); free(ord);
                *qtype = 96;
                return s;
            }
            break;
        case 2: {
            if (params.n) ss(&params, ", ");
            if (strchr(n1, '[')) ss(&params, n1);
            else {
                const char *p = n1;
                for (;;) {
                    const char *bang = strchr(p, '!');
                    int l = bang ? (int)(bang - p) : (int)strlen(p);
                    char *part = strndup(p, (size_t)l);
                    put_name(&params, part);
                    free(part);
                    if (!bang) break;
                    ss(&params, "!");
                    p = bang + 1;
                }
            }
            ss(&params, " "); ss(&params, type_name(flag));
            break; }
        case 3: flags = flag; top = n1; break;
        case 5:
            if (!*n1 && *e) unions[nu++] = e;
            else srcs[ns++] = src_table(n1, n2);
            break;
        case 6: {
            sb c = { 0 };
            ss(&c, e);
            if (*n1 && strcmp(n1, e)) { ss(&c, " AS "); put_alias(&c, n1); }
            list_add(&cols, c.p);
            free(c.p);
            if (*n2) { if (tcols.n) ss(&tcols, ", "); put_name(&tcols, n2); }
            sb *xb = flag == 1 ? &xpivot : flag == 2 ? &xrows : &xvalue;
            sb x = { 0 };
            ss(&x, e);
            if (flag != 1 && *n1 && strcmp(n1, e)) { ss(&x, " AS "); put_alias(&x, n1); }
            list_add(xb, x.p);
            free(x.p);
            sb st = { 0 };
            if (*n2) { put_name(&st, n2); ss(&st, " = "); }
            ss(&st, e);
            list_add(&sets, st.p);
            free(st.p);
            break; }
        case 7: joins[nj++] = r; break;
        case 8: where.n = 0; ss(&where, e); break;
        case 9: list_add(&group, e); if (flag != 1) list_add(&xgroup, e); break;
        case 10: having.n = 0; ss(&having, e); break;
        case 11: {                              /* Jet: "a DESC , b" after a descending one (VB6 oracle d_jet) */
            if (order.n) ss(&order, order.n > 5 && !strcmp(order.p + order.n - 5, " DESC") ? " , " : ", ");
            ss(&order, e);
            if (!strcasecmp(n1, "D")) ss(&order, " DESC");       /* ("d": Access, o029) */
            break; }
        }
    }
    for (int j = 0; j < nj; j++) {
        const char *a = S(joins[j]->name1), *b = S(joins[j]->name2), *e = S(joins[j]->expr);
        int fi = -1, ti = -1;
        for (int i = 0; i < ns; i++) {
            if (fi < 0 && src_has(srcs[i], a)) fi = i;
            if (ti < 0 && src_has(srcs[i], b)) ti = i;
        }
        if (fi >= 0 && fi == ti) { if (!srcs[fi]->name) add_on(srcs[fi], e); continue; }
        src *from = fi >= 0 ? srcs[fi] : src_table(a, 0), *to = ti >= 0 ? srcs[ti] : src_table(b, 0);
        int hi = fi > ti ? fi : ti, lo = fi > ti ? ti : fi;
        if (hi >= 0) { memmove(srcs + hi, srcs + hi + 1, sizeof(src *) * (size_t)(ns - hi - 1)); ns--; }
        if (lo >= 0) { memmove(srcs + lo, srcs + lo + 1, sizeof(src *) * (size_t)(ns - lo - 1)); ns--; }
        src *jn = calloc(1, sizeof *jn);
        jn->l = from; jn->r = to;
        jn->jt = joins[j]->flag == JQ_NOFLAG ? 1 : joins[j]->flag;
        add_on(jn, e);
        srcs[ns++] = jn;
    }
    sb from = { 0 };
    for (int i = 0; i < ns; i++) { if (i) ss(&from, ", "); src_sql(srcs[i], &from, 0); }
    sb out = { 0 };
    if (params.n) { ss(&out, "PARAMETERS "); ss(&out, params.p); ss(&out, ";\r\n"); }
    const char *pred = (flags & 8) ? "DISTINCTROW " : ((flags & 2) && kind != 9) ? "DISTINCT " : "";
    sb topb = { 0 };
    if (flags & 0x30) { ss(&topb, "TOP "); ss(&topb, top); ss(&topb, (flags & 0x20) ? " PERCENT " : " "); }
    sb sel = { 0 };
    if (flags & 1) { ss(&sel, "*"); if (cols.n) ss(&sel, ", "); }
    ss(&sel, cols.p);
    switch (kind) {
    case 9:
        for (int i = 0; i < nu; i++) {
            const char *u = unions[i];
            int l = (int)strlen(u);
            if (i == nu - 1) while (l && strchr(" \t\r\n;", u[l - 1])) l--;
            /* the parts joined by "UNION ": a part keeps what ends it, its line break (VB6 oracle d_jetrel, o029) */
            if (i) ss(&out, (flags & 2) ? "UNION " : "UNION ALL ");
            sput(&out, u, l);
        }
        if (order.n) { ss(&out, "\r\nORDER BY "); ss(&out, order.p); }
        *qtype = 128;
        break;
    case 4:
        ss(&out, "UPDATE "); ss(&out, from.p); ss(&out, " SET "); ss(&out, sets.p);
        *qtype = 48;
        break;
    case 5:
        ss(&out, "DELETE "); ss(&out, sel.p); if (sel.n) ss(&out, " ");
        ss(&out, "\r\nFROM "); ss(&out, from.p);
        *qtype = 32;
        break;
    case 6:
        ss(&out, "TRANSFORM "); ss(&out, xvalue.p);
        ss(&out, "\r\nSELECT "); ss(&out, topb.p); ss(&out, xrows.p);
        if (from.n) { ss(&out, "\r\nFROM "); ss(&out, from.p); }
        if (where.n) { ss(&out, "\r\nWHERE "); ss(&out, where.p); }
        if (xgroup.n) { ss(&out, "\r\nGROUP BY "); ss(&out, xgroup.p); }
        if (having.n) { ss(&out, "\r\nHAVING "); ss(&out, having.p); }
        if (order.n) { ss(&out, "\r\nORDER BY "); ss(&out, order.p); }
        ss(&out, "\r\nPIVOT "); ss(&out, xpivot.p);
        *qtype = 16;
        break;
    default:
        if (kind == 3) {
            ss(&out, "INSERT INTO "); put_name(&out, target);
            if (tcols.n) { ss(&out, " ( "); ss(&out, tcols.p); ss(&out, " )"); }
            ss(&out, "\r\n");
            *qtype = 64;
        }
        ss(&out, *pred || topb.n || sel.n ? "SELECT " : "SELECT"); ss(&out, pred); ss(&out, topb.p); ss(&out, sel.p);
        if (kind == 2) { ss(&out, " INTO "); put_name(&out, target); *qtype = 80; }
        if (from.n) { ss(&out, "\r\nFROM "); ss(&out, from.p); }
    }
    if (kind != 9 && kind != 6) {
        if (where.n) { ss(&out, "\r\nWHERE "); ss(&out, where.p); }
        if (group.n) { ss(&out, "\r\nGROUP BY "); ss(&out, group.p); }
        if (having.n) { ss(&out, "\r\nHAVING "); ss(&out, having.p); }
        if (order.n) { ss(&out, "\r\nORDER BY "); ss(&out, order.p); }
    }
    if (flags & 4) ss(&out, "\r\nWITH OWNERACCESS OPTION");
    ss(&out, ";\r\n");
    for (int i = 0; i < ns; i++) src_free(srcs[i]);
    free(srcs); free(unions); free(joins); free(ord);
    free(params.p); free(cols.p); free(tcols.p); free(where.p); free(group.p); free(having.p); free(order.p);
    free(xvalue.p); free(xrows.p); free(xpivot.p); free(xgroup.p); free(sets.p); free(from.p); free(topb.p); free(sel.p);
    return stake(&out);
}

/* ---------------------------------------------------------------- SQL -> rows */
enum { K_END, K_ID, K_BR, K_STR, K_DATE, K_NUM, K_OP };
typedef struct tok { int k, a, b; } tok;
typedef struct cp {
    const char *sql;
    tok *t; int nt, i;
    jqrow *rows; int n, cap;
    unsigned counter[256];
    char *err; int ecap; int failed;
} cp;

static int idc(int c, int first) { return c >= 0x80 || isalpha(c) || c == '_' || (!first && isdigit(c)); }
static void tokenize(cp *c)
{
    const char *s = c->sql;
    int n = (int)strlen(s), cap = 64;
    c->t = malloc(sizeof(tok) * (size_t)cap);
    for (int i = 0; i < n; ) {
        unsigned char ch = (unsigned char)s[i];
        if (isspace(ch)) { i++; continue; }
        tok t = { K_OP, i, i + 1 };
        if (ch == '[') { int j = i + 1; while (j < n && s[j] != ']') j++; t.k = K_BR; t.b = j < n ? j + 1 : n; }
        else if (ch == '"' || ch == '\'') {
            int j = i + 1;
            for (;;) { while (j < n && (unsigned char)s[j] != ch) j++; if (j + 1 < n && (unsigned char)s[j + 1] == ch) { j += 2; continue; } break; }
            t.k = K_STR; t.b = j < n ? j + 1 : n;
        } else if (ch == '#') { int j = i + 1; while (j < n && s[j] != '#') j++; t.k = K_DATE; t.b = j < n ? j + 1 : n; }
        else if (isdigit(ch) || (ch == '.' && i + 1 < n && isdigit((unsigned char)s[i + 1]))) {
            int j = i;
            while (j < n && (isdigit((unsigned char)s[j]) || s[j] == '.')) j++;
            if (j < n && (s[j] == 'e' || s[j] == 'E') && j + 1 < n && (isdigit((unsigned char)s[j + 1]) || ((s[j + 1] == '+' || s[j + 1] == '-') && j + 2 < n && isdigit((unsigned char)s[j + 2])))) {
                j += 2;
                while (j < n && isdigit((unsigned char)s[j])) j++;
            }
            t.k = K_NUM; t.b = j;
        } else if (idc(ch, 1)) { int j = i + 1; while (j < n && idc((unsigned char)s[j], 0)) j++; t.k = K_ID; t.b = j; }
        else if (i + 1 < n && ((ch == '<' && (s[i + 1] == '>' || s[i + 1] == '=')) || (ch == '>' && s[i + 1] == '='))) t.b = i + 2;
        if (c->nt == cap) { cap *= 2; c->t = realloc(c->t, sizeof(tok) * (size_t)cap); }
        c->t[c->nt++] = t;
        i = t.b;
    }
}
static tok cur(cp *c) { if (c->i < c->nt) return c->t[c->i]; tok e = { K_END, (int)strlen(c->sql), (int)strlen(c->sql) }; return e; }
static int tkw(cp *c, tok t, const char *w) { return t.k == K_ID && (int)strlen(w) == t.b - t.a && !strncasecmp(c->sql + t.a, w, (size_t)(t.b - t.a)); }
static int iskw(cp *c, const char *w) { return tkw(c, cur(c), w); }
static int accept(cp *c, const char *w) { if (iskw(c, w)) { c->i++; return 1; } return 0; }
static int isop(cp *c, const char *o) { tok t = cur(c); return t.k == K_OP && (int)strlen(o) == t.b - t.a && !strncmp(c->sql + t.a, o, (size_t)(t.b - t.a)); }
static int top_is(cp *c, tok t, const char *o) { return t.k == K_OP && (int)strlen(o) == t.b - t.a && !strncmp(c->sql + t.a, o, (size_t)(t.b - t.a)); }
static void fail(cp *c, const char *fmt, ...)
{
    if (c->failed) return;
    c->failed = 1;
    va_list ap; va_start(ap, fmt);
    vsnprintf(c->err, (size_t)c->ecap, fmt, ap);
    va_end(ap);
}
static void expect(cp *c, const char *w)
{
    if (!accept(c, w)) { tok t = cur(c); fail(c, "expected %s at '%.30s'", w, c->sql + t.a); }
}
static char *text(cp *c, int a, int b)
{
    while (a < b && isspace((unsigned char)c->sql[a])) a++;
    while (b > a && isspace((unsigned char)c->sql[b - 1])) b--;
    return strndup(c->sql + a, (size_t)(b - a));
}
/* stop: a list of words / operators, 0-terminated */
static int in_stop(cp *c, tok t, const char *const *stop)
{
    for (int i = 0; stop[i]; i++) {
        if (isalpha((unsigned char)stop[i][0])) { if (tkw(c, t, stop[i])) return 1; }
        else if (top_is(c, t, stop[i])) return 1;
    }
    return 0;
}
/* an expression: its text up to a word of stop (or a closing parenthesis) at depth 0 */
static char *span(cp *c, const char *const *stop)
{
    int start = cur(c).a, end = start, depth = 0;
    while (c->i < c->nt) {
        tok t = cur(c);
        if (top_is(c, t, "(")) depth++;
        else if (top_is(c, t, ")")) { if (!depth) break; depth--; }
        else if (!depth && in_stop(c, t, stop)) break;
        end = t.b;
        c->i++;
    }
    return text(c, start, end);
}
static char *name(cp *c)
{
    tok t = cur(c);
    if (t.k == K_BR) { c->i++; return strndup(c->sql + t.a + 1, (size_t)(t.b - t.a - 2)); }
    if (t.k == K_ID) { c->i++; return strndup(c->sql + t.a, (size_t)(t.b - t.a)); }
    if (t.k == K_STR) { c->i++; return strndup(c->sql + t.a + 1, (size_t)(t.b - t.a - 2)); }
    fail(c, "a name expected at '%.30s'", c->sql + t.a);
    c->i++;
    return strdup("");
}
static void row(cp *c, int attr, int flag, char *n1, char *n2, char *e)   /* takes the strings */
{
    if (c->n == c->cap) { c->cap = c->cap ? c->cap * 2 : 16; c->rows = realloc(c->rows, sizeof(jqrow) * (size_t)c->cap); }
    jqrow *r = &c->rows[c->n++];
    r->attr = attr; r->order = ++c->counter[attr & 255]; r->flag = flag;
    r->name1 = n1; r->name2 = n2; r->expr = e;
}

static const char *const ST_SEL[] = { "FROM", "INTO", "WHERE", "GROUP", "HAVING", "ORDER", "UNION", "PIVOT", "WITH", ";", ",", 0 };
static const char *const ST_SELD[] = { "FROM", ",", 0 };

static int params(cp *c)
{
    if (!accept(c, "PARAMETERS")) return 0;
    for (;;) {
        int start = c->i;
        while (c->i < c->nt && !isop(c, ",") && !isop(c, ";")) c->i++;
        int end = c->i, j = end - start;
        const tok *T = c->t + start;
        char tname[64] = "";
        if (j >= 4 && top_is(c, T[j - 1], ")") && top_is(c, T[j - 3], "(")) { snprintf(tname, sizeof tname, "%.*s", T[j - 4].b - T[j - 4].a, c->sql + T[j - 4].a); j -= 4; }
        else if (j >= 1) { snprintf(tname, sizeof tname, "%.*s", T[j - 1].b - T[j - 1].a, c->sql + T[j - 1].a); j -= 1; }
        static const struct { const char *n; int t; } PT[] = { { "bit", 1 }, { "yesno", 1 }, { "logical", 1 }, { "boolean", 1 },
            { "byte", 2 }, { "short", 3 }, { "smallint", 3 }, { "integer2", 3 }, { "long", 4 }, { "integer", 4 }, { "int", 4 },
            { "integer4", 4 }, { "counter", 4 }, { "currency", 5 }, { "money", 5 }, { "ieeesingle", 6 }, { "single", 6 },
            { "real", 6 }, { "float4", 6 }, { "ieeedouble", 7 }, { "double", 7 }, { "float", 7 }, { "float8", 7 },
            { "number", 7 }, { "numeric", 7 }, { "datetime", 8 }, { "date", 8 }, { "time", 8 }, { "timestamp", 8 },
            { "binary", 9 }, { "varbinary", 9 }, { "text", 10 }, { "char", 10 }, { "varchar", 10 }, { "string", 10 },
            { "alphanumeric", 10 }, { "longbinary", 11 }, { "oleobject", 11 }, { "general", 11 }, { "image", 11 },
            { "longtext", 12 }, { "memo", 12 }, { "note", 12 }, { "guid", 15 }, { "uniqueidentifier", 15 }, { 0, 0 } };
        int ty = 0;
        for (int k = 0; PT[k].n; k++) if (!strcasecmp(PT[k].n, tname)) { ty = PT[k].t; break; }
        char *nm = j > 0 ? text(c, T[0].a, T[j - 1].b) : strdup("");
        row(c, 2, ty, nm, 0, 0);
        if (isop(c, ",")) { c->i++; continue; }
        break;
    }
    if (isop(c, ";")) c->i++;
    return 0;
}

static int predicate(cp *c, char **top)
{
    int flags = 0;
    *top = 0;
    if (accept(c, "DISTINCTROW")) flags |= 8;
    else if (accept(c, "DISTINCT")) flags |= 2;
    else accept(c, "ALL");
    if (accept(c, "TOP")) {
        tok t = cur(c);
        *top = strndup(c->sql + t.a, (size_t)(t.b - t.a));
        c->i++;
        flags |= 0x10;
        if (accept(c, "PERCENT")) flags |= 0x20;
    }
    return flags;
}

typedef struct item { char *e, *alias; } item;
/* a select list: * and the items (expression [AS alias]) */
static int select_list(cp *c, const char *const *stop, item **items, int *n)
{
    int star = 0, cap = 8;
    *items = malloc(sizeof(item) * (size_t)cap);
    *n = 0;
    for (;;) {
        tok nx = c->i + 1 < c->nt ? c->t[c->i + 1] : (tok){ K_END, 0, 0 };
        if (isop(c, "*") && (nx.k == K_END || top_is(c, nx, ",") || top_is(c, nx, ";") || in_stop(c, nx, stop))) {
            star = 1; c->i++;
        } else {
            int start = c->i;
            const char *st2[16]; int k = 0;
            for (; stop[k] && k < 14; k++) st2[k] = stop[k];
            st2[k++] = ","; st2[k] = 0;
            char *t = span(c, st2);
            int depth = 0, at = -1;
            for (int q = start; q < c->i; q++) {
                if (top_is(c, c->t[q], "(")) depth++;
                else if (top_is(c, c->t[q], ")")) depth--;
                else if (!depth && tkw(c, c->t[q], "AS")) at = q;
            }
            char *alias = 0;
            if (at >= 0 && at + 1 < c->i && at > start) {
                tok al = c->t[at + 1];
                if (al.k == K_BR || al.k == K_STR) alias = strndup(c->sql + al.a + 1, (size_t)(al.b - al.a - 2));
                else alias = strndup(c->sql + al.a, (size_t)(al.b - al.a));
                free(t);
                t = text(c, c->t[start].a, c->t[at - 1].b);
            }
            if (!*t && !alias && c->i == start) { free(t); break; }   /* "SELECT;": no column (Jet keeps none) */
            if (*n == cap) { cap *= 2; *items = realloc(*items, sizeof(item) * (size_t)cap); }
            (*items)[*n].e = t; (*items)[(*n)++].alias = alias;
        }
        if (isop(c, ",")) { c->i++; continue; }
        break;
    }
    return star;
}

static char **join_tree(cp *c, int *nn);
/* the table of one side of a join its condition names */
static const char *pick(char **names, int n, const char *cond, int last)
{
    int hit = -1;
    for (int i = 0; i < n; i++) {
        size_t l = strlen(names[i]);
        for (const char *p = cond; *p; p++) {
            const char *q = p;
            int br = 0;
            if (*q == '[') { br = 1; q++; }
            if (strncasecmp(q, names[i], l)) continue;
            if (p > cond && (isalnum((unsigned char)p[-1]) || p[-1] == '_' || p[-1] == ']' || (unsigned char)p[-1] >= 0x80)) continue;
            q += l;
            if (br) { if (*q != ']') continue; q++; }
            else if (isalnum((unsigned char)*q) || *q == '_' || (unsigned char)*q >= 0x80) continue;
            while (*q == ' ') q++;
            if (*q != '.') continue;
            if (hit < 0 || last) hit = i;
            if (!last) return names[hit];
            break;
        }
    }
    if (hit >= 0) return names[hit];
    return last ? names[n - 1] : names[0];
}
static char **source(cp *c, int *nn)
{
    char **names = malloc(sizeof(char *) * 2);
    *nn = 1;
    if (isop(c, "(")) {
        tok nx = c->i + 1 < c->nt ? c->t[c->i + 1] : (tok){ K_END, 0, 0 };
        if (tkw(c, nx, "SELECT")) {
            c->i++;
            int start = cur(c).a, depth = 0;
            while (c->i < c->nt) {
                tok t = cur(c);
                if (top_is(c, t, "(")) depth++;
                else if (top_is(c, t, ")")) { if (!depth) break; depth--; }
                c->i++;
            }
            char *sub = text(c, start, cur(c).a);
            c->i++;
            accept(c, "AS");
            char *alias = name(c);
            names[0] = strdup(alias);
            row(c, 5, JQ_NOFLAG, 0, alias, sub);
            return names;
        }
        c->i++;
        free(names);
        char **in = join_tree(c, nn);
        if (isop(c, ")")) c->i++;
        return in;
    }
    char *t = name(c);
    while (isop(c, ".")) { c->i++; free(t); t = name(c); }
    char *alias = 0;
    static const char *const NOT_ALIAS[] = { "INNER", "LEFT", "RIGHT", "JOIN", "ON", "WHERE", "GROUP", "HAVING", "ORDER",
        "UNION", "PIVOT", "WITH", "SET", "IN", "OUTER", "FULL", 0 };
    if (accept(c, "AS")) alias = name(c);
    else if ((cur(c).k == K_ID || cur(c).k == K_BR) && !in_stop(c, cur(c), NOT_ALIAS)) alias = name(c);
    names[0] = strdup(alias ? alias : t);
    row(c, 5, JQ_NOFLAG, t, alias, 0);
    return names;
}
/* "(a) AND (b) ..." of a join's ON: its parts (malloc'd, the parentheses taken off), 0 if it is not of that form */
static int on_parts(const char *s, char ***parts)
{
    int n = 0;
    *parts = 0;
    const char *p = s;
    for (;;) {
        while (isspace((unsigned char)*p)) p++;
        if (*p != '(') break;
        int depth = 0;
        const char *q = p;
        char quote = 0;
        for (; *q; q++) {
            if (quote) { if (*q == quote) quote = 0; continue; }
            if (*q == '"' || *q == '\'' || *q == '#') { quote = *q; continue; }
            if (*q == '[') { quote = ']'; continue; }
            if (*q == '(') depth++;
            else if (*q == ')' && --depth == 0) break;
        }
        if (!*q) break;
        *parts = realloc(*parts, sizeof(char *) * (size_t)(n + 1));
        const char *a = p + 1, *b = q;
        while (a < b && isspace((unsigned char)*a)) a++;
        while (b > a && isspace((unsigned char)b[-1])) b--;
        (*parts)[n++] = strndup(a, (size_t)(b - a));
        p = q + 1;
        while (isspace((unsigned char)*p)) p++;
        if (!*p) {
            if (n >= 2) return n;
            break;
        }
        if (strncasecmp(p, "AND", 3) || isalnum((unsigned char)p[3])) break;
        p += 3;
    }
    for (int i = 0; i < n; i++) free((*parts)[i]);
    free(*parts);
    *parts = 0;
    return 0;
}
static char **join_tree(cp *c, int *nn)
{
    char **names = source(c, nn);
    static const char *const ST_ON[] = { "INNER", "LEFT", "RIGHT", "JOIN", "WHERE", "GROUP", "HAVING", "ORDER", "UNION",
        "PIVOT", "WITH", "SET", ",", ";", ")", 0 };
    while (!c->failed) {
        int jt = 0;
        if (accept(c, "INNER")) jt = 1;
        else if (accept(c, "LEFT")) { jt = 2; accept(c, "OUTER"); }
        else if (accept(c, "RIGHT")) { jt = 3; accept(c, "OUTER"); }
        if (!jt) { if (iskw(c, "JOIN")) jt = 1; else return names; }
        expect(c, "JOIN");
        int nr;
        char **right = source(c, &nr);
        expect(c, "ON");
        char *cond = span(c, ST_ON);
        char *a = strdup(pick(names, *nn, cond, 1)), *b = strdup(pick(right, nr, cond, 0));
        char **parts;
        int np = on_parts(cond, &parts);
        if (np) {                               /* one row each, the last first (see add_on) */
            for (int k = np - 1; k >= 0; k--) row(c, 7, jt, strdup(a), strdup(b), parts[k]);
            free(parts); free(a); free(b); free(cond);
        } else row(c, 7, jt, a, b, cond);
        names = realloc(names, sizeof(char *) * (size_t)(*nn + nr + 1));
        memcpy(names + *nn, right, sizeof(char *) * (size_t)nr);
        *nn += nr;
        free(right);
    }
    return names;
}
static void free_names(char **n, int k) { for (int i = 0; i < k; i++) free(n[i]); free(n); }
static void from_clause(cp *c)
{
    if (!accept(c, "FROM")) return;
    for (;;) {
        int nn = 0;
        char **names = join_tree(c, &nn);
        free_names(names, nn);
        if (isop(c, ",") && !c->failed) { c->i++; continue; }
        break;
    }
}
static void order_by(cp *c)
{
    static const char *const ST[] = { "PIVOT", "WITH", ";", ",", "UNION", 0 };
    if (!accept(c, "ORDER")) return;
    expect(c, "BY");
    while (!c->failed) {
        char *e = span(c, ST), *d = 0;
        int l = (int)strlen(e);
        if (l > 4 && !strcasecmp(e + l - 4, "DESC") && isspace((unsigned char)e[l - 5])) {
            e[l - 4] = 0; d = strdup("D");
        } else if (l > 3 && !strcasecmp(e + l - 3, "ASC") && isspace((unsigned char)e[l - 4])) e[l - 3] = 0;
        l = (int)strlen(e);
        while (l && isspace((unsigned char)e[l - 1])) e[--l] = 0;
        row(c, 11, JQ_NOFLAG, d, 0, e);
        if (isop(c, ",")) { c->i++; continue; }
        break;
    }
}
static void tail(cp *c, int xtab)
{
    static const char *const W[] = { "GROUP", "HAVING", "ORDER", "UNION", "PIVOT", "WITH", ";", 0 };
    static const char *const G[] = { "HAVING", "ORDER", "UNION", "PIVOT", "WITH", ";", ",", 0 };
    static const char *const H[] = { "ORDER", "UNION", "PIVOT", "WITH", ";", 0 };
    if (accept(c, "WHERE")) row(c, 8, JQ_NOFLAG, 0, 0, span(c, W));
    if (accept(c, "GROUP")) {
        expect(c, "BY");
        while (!c->failed) {
            row(c, 9, xtab ? 2 : 0, 0, 0, span(c, G));
            if (isop(c, ",")) { c->i++; continue; }
            break;
        }
    }
    if (accept(c, "HAVING")) row(c, 10, JQ_NOFLAG, 0, 0, span(c, H));
    order_by(c);
}
static int owneraccess(cp *c)
{
    if (!accept(c, "WITH")) return 0;
    expect(c, "OWNERACCESS");
    expect(c, "OPTION");
    return 4;
}
static int union_ahead(cp *c)
{
    int depth = 0;
    for (int i = c->i; i < c->nt; i++) {
        if (top_is(c, c->t[i], "(")) depth++;
        else if (top_is(c, c->t[i], ")")) depth--;
        else if (!depth && tkw(c, c->t[i], "UNION")) return 1;
    }
    return 0;
}
static int last_part(cp *c)
{
    int depth = 0;
    for (int i = c->i; i < c->nt; i++) {
        if (top_is(c, c->t[i], "(")) depth++;
        else if (top_is(c, c->t[i], ")")) depth--;
        else if (!depth && tkw(c, c->t[i], "UNION")) return 0;
    }
    return 1;
}
/* a union: two parts, the query before its last UNION as written (with the UNIONs in it) and the one after it, as
 * Jet keeps "a union b UNION c" (VB6 oracle DEV-14/q1, PSC SuppExist) */
static void do_union(cp *c)
{
    row(c, 1, 9, 0, 0, 0);
    int depth = 0, lastu = -1, end = c->nt;
    for (int i = c->i; i < c->nt; i++) {
        tok t = c->t[i];
        if (top_is(c, t, "(")) depth++;
        else if (top_is(c, t, ")")) depth--;
        else if (!depth && tkw(c, t, "UNION")) lastu = i;
    }
    depth = 0;
    for (int i = lastu + 1; i < c->nt; i++) {    /* the end of the last part: its ORDER BY or ";" */
        tok t = c->t[i];
        if (top_is(c, t, "(")) depth++;
        else if (top_is(c, t, ")")) depth--;
        else if (!depth && (tkw(c, t, "ORDER") || top_is(c, t, ";"))) { end = i; break; }
    }
    int start = cur(c).a, e1 = c->t[lastu].a;
    /* the left part as written, with the line break before UNION: Jet shows the parts joined by "UNION " (o029, o046) */
    while (start < e1 && isspace((unsigned char)c->sql[start])) start++;
    row(c, 5, JQ_NOFLAG, 0, strdup("X7YZ_____1"), strndup(c->sql + start, (size_t)(e1 - start)));
    c->i = lastu + 1;
    int distinct = !accept(c, "ALL");
    int s2 = cur(c).a;
    c->i = end;
    row(c, 5, JQ_NOFLAG, 0, strdup("X7YZ_____2"), text(c, s2, c->i < c->nt ? cur(c).a : (int)strlen(c->sql)));
    order_by(c);
    row(c, 3, distinct ? 3 : 1, 0, 0, 0);       /* (DAO's and Access's: 1 always, d_jetrel) */
}
static void items_rows(cp *c, item *it, int n, int flag, char **tcols, int ntc)
{
    for (int k = 0; k < n; k++)
        row(c, 6, flag, it[k].alias, k < ntc ? strdup(tcols[k]) : 0, it[k].e);
    free(it);
}
static void select_or_union(cp *c)
{
    if (union_ahead(c)) { do_union(c); return; }
    expect(c, "SELECT");
    char *top;
    int flags = predicate(c, &top), n;
    item *it;
    int star = select_list(c, ST_SEL, &it, &n);
    if (accept(c, "INTO")) row(c, 1, 2, name(c), 0, 0);
    from_clause(c);
    items_rows(c, it, n, 0, 0, 0);
    tail(c, 0);
    flags |= owneraccess(c);
    if (star) flags |= 1;
    row(c, 3, flags, top, 0, 0);
}
static void insert(cp *c)
{
    expect(c, "INSERT");
    expect(c, "INTO");
    char *target = name(c);
    char **tcols = 0;
    int ntc = 0;
    if (isop(c, "(")) {
        c->i++;
        while (!isop(c, ")") && c->i < c->nt) {
            tcols = realloc(tcols, sizeof(char *) * (size_t)(ntc + 1));
            tcols[ntc++] = name(c);
            if (isop(c, ",")) c->i++;
        }
        c->i++;
    }
    row(c, 1, 3, target, 0, 0);
    if (accept(c, "VALUES")) {
        static const char *const ST[] = { ",", ";", 0 };
        if (isop(c, "(")) c->i++;
        for (int k = 0; !c->failed; k++) {
            row(c, 6, 0, 0, k < ntc ? strdup(tcols[k]) : 0, span(c, ST));
            if (isop(c, ",")) { c->i++; continue; }
            break;
        }
        if (isop(c, ")")) c->i++;
        row(c, 3, 0, 0, 0, 0);
    } else {
        expect(c, "SELECT");
        char *top;
        int flags = predicate(c, &top), n;
        item *it;
        int star = select_list(c, ST_SEL, &it, &n);
        from_clause(c);
        items_rows(c, it, n, 0, tcols, ntc);
        tail(c, 0);
        flags |= owneraccess(c);
        if (star) flags |= 1;
        row(c, 3, flags, top, 0, 0);
    }
    for (int k = 0; k < ntc; k++) free(tcols[k]);
    free(tcols);
}
static void update(cp *c)
{
    static const char *const ST[] = { "WHERE", ",", ";", 0 };
    static const char *const SW[] = { ";", 0 };
    expect(c, "UPDATE");
    row(c, 1, 4, 0, 0, 0);
    for (;;) {
        int nn = 0;
        char **names = join_tree(c, &nn);
        free_names(names, nn);
        if (isop(c, ",") && !c->failed) { c->i++; continue; }
        break;
    }
    expect(c, "SET");
    while (!c->failed) {
        int start = c->i;
        char *e = span(c, ST);
        int depth = 0, at = -1;
        for (int q = start; q < c->i; q++) {
            if (top_is(c, c->t[q], "(")) depth++;
            else if (top_is(c, c->t[q], ")")) depth--;
            else if (!depth && top_is(c, c->t[q], "=")) { at = q; break; }
        }
        if (at > start && at + 1 < c->i) {
            char *tgt = text(c, c->t[start].a, c->t[at - 1].b), *val = text(c, c->t[at + 1].a, c->t[c->i - 1].b);
            int l = (int)strlen(tgt);
            if (l >= 2 && tgt[0] == '[' && tgt[l - 1] == ']' && !strchr(tgt + 1, '[')) { memmove(tgt, tgt + 1, (size_t)l - 2); tgt[l - 2] = 0; }
            row(c, 6, 0, 0, tgt, val);
            free(e);
        } else row(c, 6, 0, 0, 0, e);
        if (isop(c, ",")) { c->i++; continue; }
        break;
    }
    if (accept(c, "WHERE")) row(c, 8, JQ_NOFLAG, 0, 0, span(c, SW));
}
static void delete_(cp *c)
{
    static const char *const SW[] = { ";", 0 };
    expect(c, "DELETE");
    row(c, 1, 5, 0, 0, 0);
    int flags = 0, n = 0;
    item *it = 0;
    if (accept(c, "DISTINCTROW")) flags = 8;
    if (!iskw(c, "FROM")) { if (select_list(c, ST_SELD, &it, &n)) flags |= 1; }
    from_clause(c);
    if (it) items_rows(c, it, n, 0, 0, 0);
    if (accept(c, "WHERE")) row(c, 8, JQ_NOFLAG, 0, 0, span(c, SW));
    row(c, 3, flags, 0, 0, 0);
}
static void crosstab(cp *c)
{
    static const char *const SS[] = { "SELECT", 0 };
    static const char *const W[] = { "GROUP", "HAVING", "ORDER", "PIVOT", ";", 0 };
    static const char *const G[] = { "HAVING", "ORDER", "PIVOT", ";", ",", 0 };
    static const char *const H[] = { "ORDER", "PIVOT", ";", 0 };
    expect(c, "TRANSFORM");
    int vstart = c->i;
    char *value = span(c, SS), *alias = 0;
    /* "value AS alias" */
    for (int q = c->i - 2; q > vstart; q--)
        if (tkw(c, c->t[q], "AS") && q == c->i - 2) {
            tok al = c->t[q + 1];
            alias = al.k == K_BR ? strndup(c->sql + al.a + 1, (size_t)(al.b - al.a - 2)) : strndup(c->sql + al.a, (size_t)(al.b - al.a));
            free(value);
            value = text(c, c->t[vstart].a, c->t[q - 1].b);
            break;
        }
    row(c, 1, 6, 0, 0, 0);
    expect(c, "SELECT");
    char *top;
    int flags = predicate(c, &top), n;
    item *it;
    select_list(c, ST_SEL, &it, &n);
    from_clause(c);
    row(c, 6, 0, alias, 0, value);
    items_rows(c, it, n, 2, 0, 0);
    if (accept(c, "WHERE")) row(c, 8, JQ_NOFLAG, 0, 0, span(c, W));
    if (accept(c, "GROUP")) {
        expect(c, "BY");
        while (!c->failed) {
            row(c, 9, 2, 0, 0, span(c, G));
            if (isop(c, ",")) { c->i++; continue; }
            break;
        }
    }
    if (accept(c, "HAVING")) row(c, 10, JQ_NOFLAG, 0, 0, span(c, H));
    order_by(c);
    expect(c, "PIVOT");
    int start = cur(c).a;
    while (c->i < c->nt && !isop(c, ";")) c->i++;
    char *pivot = text(c, start, cur(c).a), *base = strdup(pivot);
    /* the pivot's own expression: without its IN (...) */
    int depth = 0;
    for (int q = 0; base[q]; q++) {
        if (base[q] == '(') depth++;
        else if (base[q] == ')') depth--;
        else if (!depth && (q == 0 || isspace((unsigned char)base[q - 1])) && !strncasecmp(base + q, "IN", 2) &&
                 (isspace((unsigned char)base[q + 2]) || base[q + 2] == '(')) {
            int e = q;
            while (e > 0 && isspace((unsigned char)base[e - 1])) e--;
            base[e] = 0;
            break;
        }
    }
    row(c, 6, 1, 0, 0, pivot);
    row(c, 9, 1, 0, 0, base);
    if (flags || top) row(c, 3, flags, top, 0, 0);
}

jqrow *jq_compile(const char *sql, int *n, char *err, int cap)
{
    cp c = { 0 };
    char dummy[8];
    c.sql = sql;
    c.err = err ? err : dummy; c.ecap = err ? cap : (int)sizeof dummy;
    *c.err = 0;
    tokenize(&c);
    row(&c, 0, 0, 0, 0, 0);
    params(&c);
    if (iskw(&c, "TRANSFORM")) crosstab(&c);
    else if (iskw(&c, "SELECT") || isop(&c, "(")) select_or_union(&c);
    else if (iskw(&c, "INSERT")) insert(&c);
    else if (iskw(&c, "UPDATE")) update(&c);
    else if (iskw(&c, "DELETE")) delete_(&c);
    else if (iskw(&c, "CREATE") || iskw(&c, "ALTER") || iskw(&c, "DROP")) {
        jq_free(c.rows, c.n);
        c.rows = 0; c.n = c.cap = 0;
        memset(c.counter, 0, sizeof c.counter);
        row(&c, 0, 0, 0, 0, 0);
        char *d = strdup(sql);
        int l = (int)strlen(d);
        while (l && (isspace((unsigned char)d[l - 1]) || d[l - 1] == ';')) d[--l] = 0;
        char *s = d;
        while (isspace((unsigned char)*s)) s++;
        row(&c, 1, 7, 0, 0, strdup(s));
        free(d);
    } else fail(&c, "not a query: '%.40s'", sql);
    row(&c, 255, JQ_NOFLAG, 0, 0, 0);
    free(c.t);
    if (c.failed) { jq_free(c.rows, c.n); *n = 0; return 0; }
    *n = c.n;
    return c.rows;
}

void jq_free(jqrow *rows, int n)
{
    for (int i = 0; i < n; i++) { free(rows[i].name1); free(rows[i].name2); free(rows[i].expr); }
    free(rows);
}

int jq_flags(const jqrow *rows, int n)
{
    int kind = 1;
    for (int i = 0; i < n; i++) if (rows[i].attr == 1) kind = rows[i].flag;
    switch (kind) {
    case 2: return 80; case 3: return 64; case 4: return 48; case 5: return 32; case 6: return 16;
    case 7: return 96; case 8: return 112; case 9: return 128;
    }
    return 0;
}
