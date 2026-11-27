/* lbdbinfo.c - what the IDE shows of a database at design time (the RecordSource and DataField
 * lists of the Properties window):
 *   lbdbinfo FILE            the tables ("T name") and the queries ("Q name")
 *   lbdbinfo FILE SOURCE     the fields of a table, a query or a Jet SQL statement ("F name")
 *   lbdbinfo -ado CONNECT    the tables ("T name") and views ("Q name") of an ADO connection
 *   lbdbinfo -ado CONNECT CMDTYPE SOURCE
 *                            the fields of an ADO command, as the programs' ADO opens it
 *                            ("F name<TAB>type<TAB>defined size<TAB>precision"): the DataEnvironment
 *                            designer's field lists
 * A Jet .mdb is read through its SQLite copy FILE.lbdb, made as the programs make it.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbdb.h"
#include <stdio.h>
#include <ctype.h>
#include <math.h>
#include <string.h>
#include <strings.h>
#include <sys/stat.h>

static int has_meta(sqlite3 *db)
{
    sqlite3_stmt *st;
    int r = 0;
    if (sqlite3_prepare_v2(db, "SELECT 1 FROM sqlite_master WHERE name = 'MSysLB_Tables'", -1, &st, 0) == SQLITE_OK) {
        r = sqlite3_step(st) == SQLITE_ROW;
        sqlite3_finalize(st);
    }
    return r;
}

static void list_rows(sqlite3 *db, const char *sql, const char *tag, const char *arg)
{
    sqlite3_stmt *st;
    if (sqlite3_prepare_v2(db, sql, -1, &st, 0) != SQLITE_OK) return;
    if (arg) sqlite3_bind_text(st, 1, arg, -1, SQLITE_STATIC);
    while (sqlite3_step(st) == SQLITE_ROW)
        printf("%s\t%s\n", tag, (const char *)sqlite3_column_text(st, 0));
    sqlite3_finalize(st);
}

/* the columns of a Jet statement */
static int sql_fields(sqlite3 *db, const char *jet)
{
    jsql q;
    if (jet_translate(jet, &q, 0) || !q.returns || !q.sql) {
        if (q.err) fprintf(stderr, "%s\n", q.errmsg);
        jsql_free(&q);
        return 1;
    }
    sqlite3_stmt *st;
    int rc = sqlite3_prepare_v2(db, q.sql, -1, &st, 0);
    jsql_free(&q);
    if (rc != SQLITE_OK) {
        fprintf(stderr, "%s\n", sqlite3_errmsg(db));
        return 1;
    }
    for (int i = 0; i < sqlite3_column_count(st); i++)
        printf("F\t%s\n", sqlite3_column_name(st, i));
    sqlite3_finalize(st);
    return 0;
}

static int fields(sqlite3 *db, const char *src, int meta)
{
    sqlite3_stmt *st;
    if (meta) {
        /* a table: its fields in their order */
        if (sqlite3_prepare_v2(db, "SELECT name FROM MSysLB_Fields WHERE tbl = ?1 ORDER BY ord", -1, &st, 0) == SQLITE_OK) {
            sqlite3_bind_text(st, 1, src, -1, SQLITE_STATIC);
            int n = 0;
            while (sqlite3_step(st) == SQLITE_ROW) {
                printf("F\t%s\n", (const char *)sqlite3_column_text(st, 0));
                n++;
            }
            sqlite3_finalize(st);
            if (n) return 0;
        }
        /* a query: the columns of its SQL */
        if (sqlite3_prepare_v2(db, "SELECT sql FROM MSysLB_Queries WHERE name = ?1", -1, &st, 0) == SQLITE_OK) {
            sqlite3_bind_text(st, 1, src, -1, SQLITE_STATIC);
            if (sqlite3_step(st) == SQLITE_ROW) {
                char *sql = strdup((const char *)sqlite3_column_text(st, 0));
                sqlite3_finalize(st);
                int r = sql_fields(db, sql);
                free(sql);
                return r;
            }
            sqlite3_finalize(st);
        }
    } else if (sqlite3_prepare_v2(db, "SELECT name FROM pragma_table_info(?1)", -1, &st, 0) == SQLITE_OK) {
        sqlite3_bind_text(st, 1, src, -1, SQLITE_STATIC);
        int n = 0;
        while (sqlite3_step(st) == SQLITE_ROW) {
            printf("F\t%s\n", (const char *)sqlite3_column_text(st, 0));
            n++;
        }
        sqlite3_finalize(st);
        if (n) return 0;
    }
    return sql_fields(db, src);                 /* a SELECT statement */
}

/* ---------------------------------------------------------------- ADO (the DataEnvironment) */
vbobj ado_createobject(const char *progid);

static int ado_failed(void)
{
    if (!vb_errpend) return 0;
    vb_err_take(0);
    fprintf(stderr, "%s\n", vb_stoutf8(vb_err_description()));
    return 1;
}

static vbvar S(const char *s) { return vb_v_s(vb_sfromutf8(s)); }

static int ado_info(int argc, char **argv)
{
    vb_db_init();
    if (argc < 4) {
        vbobj cn = ado_createobject("adodb.connection");
        vbvar a[1] = { S(argv[2]) };
        vb_late_o(cn, "open", VBM_METHOD, 1, a);
        if (ado_failed()) return 1;
        vbvar sa[1] = { vb_v_l(20) };                   /* adSchemaTables */
        vbobj rs = vb_v2obj(vb_late_o(cn, "openschema", VBM_METHOD, 1, sa));
        if (ado_failed()) return 1;
        for (int pass = 0; pass < 2; pass++) {
            vb_late_o(rs, "movefirst", VBM_METHOD, 0, 0);
            vb_err_clear_pending();
            while (!vb_v2l(vb_late_o(rs, "eof", VBM_GET, 0, 0))) {
                vbvar k[1] = { S("TABLE_TYPE") }, n[1] = { S("TABLE_NAME") };
                char *type = vb_stoutf8(vb_v2s(vb_late_o(vb_v2obj(vb_late_o(rs, "fields", VBM_GET, 1, k)), "value", VBM_GET, 0, 0)));
                int view = !strcasecmp(type, "VIEW"), sys = !strcasecmp(type, "SYSTEM TABLE");
                if (((!strcasecmp(type, "TABLE") || sys) && !pass) || (view && pass)) {
                    char *name = vb_stoutf8(vb_v2s(vb_late_o(vb_v2obj(vb_late_o(rs, "fields", VBM_GET, 1, n)), "value", VBM_GET, 0, 0)));
                    printf("%s\t%s\n", view ? "Q" : sys ? "S" : "T", name);   /* S: a system table (the Data View) */
                }
                vb_late_o(rs, "movenext", VBM_METHOD, 0, 0);
                if (ado_failed()) return 1;
            }
        }
        return 0;
    }
    vbobj rs = ado_createobject("adodb.recordset");
    vbvar a[5] = { S(argv[4]), S(argv[2]), vb_v_l(0), vb_v_l(1), vb_v_l(atoi(argv[3])) };
    vb_late_o(rs, "open", VBM_METHOD, 5, a);
    if (ado_failed()) return 1;
    vbobj flds = vb_v2obj(vb_late_o(rs, "fields", VBM_GET, 0, 0));
    int n = vb_v2l(vb_late_o(flds, "count", VBM_GET, 0, 0));
    for (int i = 0; i < n; i++) {
        vbvar ix[1] = { vb_v_l(i) };
        vbobj f = vb_v2obj(vb_late_o(flds, "item", VBM_GET, 1, ix));
        printf("F\t%s\t%d\t%d\t%d\n", vb_stoutf8(vb_v2s(vb_late_o(f, "name", VBM_GET, 0, 0))),
               (int)vb_v2l(vb_late_o(f, "type", VBM_GET, 0, 0)), (int)vb_v2l(vb_late_o(f, "definedsize", VBM_GET, 0, 0)),
               (int)vb_v2l(vb_late_o(f, "precision", VBM_GET, 0, 0)));
    }
    return ado_failed();
}

/* the Data View and the Query Designer of the IDE (phase 12.3): a query through ADO, its columns and rows
 *   lbdbinfo -ado-sql CONNECT SQL [MAX]   C<tab>name<tab>ADO type ...  then R<tab>value ... (\N: Null; \t \n \\)
 *   lbdbinfo -ado-exec CONNECT SQL        the statement run (an action query), "OK<tab>records affected" */
static void put_value(vbvar v)
{
    if (v.vt == VT_NULL) { fputs("\\N", stdout); return; }
    vb_i32 m = vb_mark();
    const char *s = vb_stoutf8(vb_v2s(v));
    if (vb_errpend) { vb_err_clear_pending(); s = "#Error"; }
    for (; *s; s++) {
        if (*s == '\t') fputs("\\t", stdout);
        else if (*s == '\n') fputs("\\n", stdout);
        else if (*s == '\r') fputs("\\r", stdout);
        else if (*s == '\\') fputs("\\\\", stdout);
        else putchar(*s);
    }
    vb_release(m);
}

static int ado_sql(int argc, char **argv, int exec)
{
    vb_db_init();
    vbobj cn = ado_createobject("adodb.connection");
    vbvar a[1] = { S(argv[2]) };
    vb_late_o(cn, "open", VBM_METHOD, 1, a);
    if (ado_failed()) return 1;
    if (exec) {
        vbvar e[2] = { S(argv[3]), vb_v_l(0) };       /* RecordsAffected */
        vb_late_o(cn, "execute", VBM_METHOD, 2, e);
        if (ado_failed()) return 1;
        printf("OK\t%ld\n", (long)vb_v2l(e[1]));
        return 0;
    }
    long max = argc > 4 ? atol(argv[4]) : 1000;
    vbobj rs = ado_createobject("adodb.recordset");
    vbvar o[4] = { S(argv[3]), vb_v_obj(cn), vb_v_l(3), vb_v_l(1) };     /* adOpenStatic, adLockReadOnly */
    vb_late_o(rs, "open", VBM_METHOD, 4, o);
    if (ado_failed()) return 1;
    vbobj flds = vb_v2obj(vb_late_o(rs, "fields", VBM_GET, 0, 0));
    int n = vb_v2l(vb_late_o(flds, "count", VBM_GET, 0, 0));
    vbobj f[256];
    if (n > 256) n = 256;
    for (int i = 0; i < n; i++) {
        vbvar ix[1] = { vb_v_l(i) };
        f[i] = vb_v2obj(vb_late_o(flds, "item", VBM_GET, 1, ix));
        printf("C\t%s\t%d\n", vb_stoutf8(vb_v2s(vb_late_o(f[i], "name", VBM_GET, 0, 0))),
               (int)vb_v2l(vb_late_o(f[i], "type", VBM_GET, 0, 0)));
    }
    long rows = 0;
    while (!vb_v2l(vb_late_o(rs, "eof", VBM_GET, 0, 0)) && (max <= 0 || rows < max)) {
        putchar('R');
        for (int i = 0; i < n; i++) {
            putchar('\t');
            put_value(vb_late_o(f[i], "value", VBM_GET, 0, 0));
        }
        putchar('\n');
        rows++;
        vb_late_o(rs, "movenext", VBM_METHOD, 0, 0);
        if (ado_failed()) return 1;
    }
    if (!vb_v2l(vb_late_o(rs, "eof", VBM_GET, 0, 0)))
        printf("M\tmore\n");                   /* more rows than MAX */
    return ado_failed();
}


/* ---------------------------------------------------------------- the object server (phase 12.4)
 * lbdbinfo -server: the objects of the runtime (DAO, ADO ...) for the IDE - the Visual Data Manager works with the
 * databases through the same DAO as the programs. One request a line, one answer a line (fields by tabs, \t \n \r \\
 * escaped):
 *   N <progid>                         CreateObject                    -> OK o:<id>
 *   C <id> <M|G|L|S> <member> <arg>... a member (Method, Get, Let, Set) -> OK <value>
 *   R <id>                             the object is not needed any more -> OK e
 *   T <id>                             TypeName                        -> OK s:<name>
 * values: e (Empty), m (Missing), n (Null), s:<text>, h:<Integer>, i:<Long>, f:<Single>, d:<Double>, c:<Currency>,
 * t:<Date as a Double>, b:<0|1>, y:<Byte>, o:<id> (o:0 Nothing), a:<bytes> (an array: its size only), x:<Decimal>
 * an error: ER <number> <description> */
#include <stdlib.h>
static vbobj *srv_obj;
static int srv_n, srv_cap;

static int srv_put(vbobj o)
{
    if (!o) return 0;
    for (int i = 1; i < srv_n; i++) if (srv_obj[i] == o) return i;      /* (the same object: the same id) */
    if (srv_n == 0) srv_n = 1;
    if (srv_n >= srv_cap) { srv_cap = srv_cap ? srv_cap * 2 : 64; srv_obj = realloc(srv_obj, sizeof(vbobj) * (size_t)srv_cap); }
    vb_oaddref(o);
    srv_obj[srv_n] = o;
    return srv_n++;
}

static vbobj srv_get(const char *s) { long i = atol(s); return i > 0 && i < srv_n ? srv_obj[i] : 0; }

static void srv_text(const char *s)
{
    for (; *s; s++) {
        if (*s == '\t') fputs("\\t", stdout);
        else if (*s == '\n') fputs("\\n", stdout);
        else if (*s == '\r') fputs("\\r", stdout);
        else if (*s == '\\') fputs("\\\\", stdout);
        else putchar(*s);
    }
}

static void srv_unescape(char *s)
{
    char *d = s;
    for (; *s; s++) {
        if (*s == '\\' && s[1]) {
            s++;
            *d++ = *s == 't' ? '\t' : *s == 'n' ? '\n' : *s == 'r' ? '\r' : *s;
        } else *d++ = *s;
    }
    *d = 0;
}

static vbvar srv_arg(char *a)
{
    char k = a[0];
    const char *v = a[0] && a[1] == ':' ? a + 2 : "";
    switch (k) {
    case 'm': return vb_v_missing();
    case 'n': return vb_v_null();
    case 's': return vb_v_s(vb_sfromutf8(v));
    case 'h': return vb_v_i((vb_i16)atoi(v));
    case 'i': return vb_v_l((vb_i32)atol(v));
    case 'f': return vb_v_f((float)atof(v));
    case 'd': return vb_v_d(atof(v));
    case 'c': return vb_v_c((vb_cy)llround(atof(v) * 10000.0));
    case 't': return vb_v_date(atof(v));
    case 'b': return vb_v_bool(atoi(v) ? -1 : 0);
    case 'y': return vb_v_byte((vb_u8)atoi(v));
    case 'o': { vbobj o = srv_get(v); return vb_v_obj(o); }
    case 'a': {                                     /* a Byte array: its bytes in hex */
        vb_i32 n = (vb_i32)(strlen(v) / 2);
        vbarr arr; memset(&arr, 0, sizeof arr);
        if (n == 0) return vb_v_arr(&arr, VT_UI1);
        vb_adim(&arr, 1, VBA_PLAIN, 0, 1, 0, n - 1);
        for (vb_i32 i = 0; i < n; i++) { unsigned x = 0; sscanf(v + 2 * i, "%2x", &x); ((unsigned char *)arr.data)[i] = (unsigned char)x; }
        vbvar r = vb_v_arr(&arr, VT_UI1);
        vb_afree(&arr);
        return r;
    }
    default: return vb_v_empty();
    }
}

static void srv_value(vbvar v)
{
    char b[64];
    switch (v.vt) {
    case VT_EMPTY: fputs("e", stdout); break;
    case VT_NULL: fputs("n", stdout); break;
    case VT_I2: printf("h:%d", v.u.i); break;
    case VT_I4: printf("i:%ld", (long)v.u.l); break;
    case VT_R4: snprintf(b, sizeof b, "f:%.9g", v.u.f); fputs(b, stdout); break;
    case VT_R8: snprintf(b, sizeof b, "d:%.17g", v.u.d); fputs(b, stdout); break;
    case VT_CY: printf("c:%s", vb_stoutf8(vb_cstr_c(v.u.cy))); break;
    case VT_DATE: snprintf(b, sizeof b, "t:%.17g", v.u.d); fputs(b, stdout); break;
    case VT_BOOL: printf("b:%d", v.u.i ? 1 : 0); break;
    case VT_UI1: printf("y:%d", v.u.b); break;
    case VT_DISPATCH: printf("o:%d", srv_put(v.u.o)); break;
    case VT_DECIMAL: printf("x:%s", vb_stoutf8(vb_v2s(v))); break;
    case VT_BSTR: fputs("s:", stdout); srv_text(vb_stoutf8(v.u.s)); break;
    default:
        if (v.vt == (VT_ARRAY | VT_UI1)) {                           /* binary data: its bytes in hex */
            vbarr *arr = v.u.a;
            fputs("a:", stdout);
            if (arr && arr->nd) for (vb_i32 i = 0; i < arr->n; i++) printf("%02x", ((unsigned char *)arr->data)[i]);
            break;
        }
        if (v.vt & VT_ARRAY) { fputs("a", stdout); break; }         /* (other arrays: not shown) */
        fputs("s:", stdout); srv_text(vb_stoutf8(vb_v2s(v)));
        if (vb_errpend) vb_err_clear_pending();
    }
}

static int srv_error(void)
{
    if (!vb_errpend) return 0;
    vb_err_take(0);
    printf("ER\t%ld\t", (long)vb_err_number());
    srv_text(vb_stoutf8(vb_err_description()));
    putchar('\n');
    vb_err_clear();
    return 1;
}

static int server(void)
{
    vb_db_init();
    size_t cap = 4096;
    char *line = malloc(cap);
    for (;;) {                                  /* (a line of any length; no getline on Windows) */
        size_t n = 0;
        int c;
        while ((c = getchar()) != EOF && c != '\n') {
            if (n + 2 >= cap) { cap *= 2; line = realloc(line, cap); }
            line[n++] = (char)c;
        }
        if (c == EOF && n == 0) break;
        line[n] = 0;
        while (n && line[n - 1] == '\r') line[--n] = 0;
        char *f[64];
        int nf = 0;
        char *p = line;
        f[nf++] = p;
        for (; *p && nf < 64; p++) if (*p == '\t') { *p = 0; f[nf++] = p + 1; }
        for (int i = 0; i < nf; i++) srv_unescape(f[i]);
        vb_i32 m = vb_mark();
        if (!strcmp(f[0], "N") && nf >= 2) {
            vbobj o = vb_createobject(vb_sfromutf8(f[1]), vb_v_missing());
            if (!srv_error()) printf("OK\to:%d\n", srv_put(o));
        } else if (!strcmp(f[0], "C") && nf >= 4) {
            vbobj o = srv_get(f[1]);
            for (char *q = f[3]; *q; q++) *q = (char)tolower((unsigned char)*q);   /* (names as the compiler gives them) */
            int kind = f[2][0] == 'M' ? VBM_METHOD : f[2][0] == 'L' ? VBM_LET : f[2][0] == 'S' ? VBM_SET : VBM_GET;
            vbvar a[60];
            int na = nf - 4;
            for (int i = 0; i < na; i++) a[i] = srv_arg(f[4 + i]);
            if (!o) { printf("ER\t91\tObject variable or With block variable not set\n"); }
            else {
                vbvar r = vb_late_o(o, f[3], kind, na, a);
                if (!srv_error()) { fputs("OK\t", stdout); srv_value(r); putchar('\n'); }
            }
        } else if (!strcmp(f[0], "R") && nf >= 2) {
            long i = atol(f[1]);
            if (i > 0 && i < srv_n && srv_obj[i]) { vbobj o = srv_obj[i]; srv_obj[i] = 0; vb_orelease(o); }
            printf("OK\te\n");
        } else if (!strcmp(f[0], "T") && nf >= 2) {
            vbobj o = srv_get(f[1]);
            fputs("OK\ts:", stdout);
            srv_text(o ? vb_stoutf8(vb_otypename(o)) : "Nothing");
            putchar('\n');
        } else if (!strcmp(f[0], "Q")) {
            break;
        } else printf("ER\t5\tInvalid procedure call or argument\n");
        if (vb_errpend) vb_err_clear_pending();
        vb_release(m);
        fflush(stdout);
    }
    free(line);
    return 0;
}

int main(int argc, char **argv)
{
    if (argc >= 2 && !strcmp(argv[1], "-server")) {
        vb_init(1, argv);
        return server();
    }
    if (argc >= 4 && (!strcmp(argv[1], "-ado-sql") || !strcmp(argv[1], "-ado-exec"))) {
        vb_init(1, argv);
        return ado_sql(argc, argv, !strcmp(argv[1], "-ado-exec"));
    }
    if (argc >= 3 && !strcmp(argv[1], "-ado")) {
        vb_init(1, argv);
        return ado_info(argc, argv);
    }
    if (argc < 2) {
        fprintf(stderr, "usage: lbdbinfo FILE [SOURCE]\n");
        return 2;
    }
    const char *file = argv[1];
    char *use = strdup(file);
    if (mdb_is_jet(file)) {
        free(use);
        use = malloc(strlen(file) + 8);
        sprintf(use, "%s.lbdb", file);
        struct stat a, b;
        int fresh = stat(use, &b) == 0 && stat(file, &a) == 0 && b.st_mtime >= a.st_mtime;
        char err[512];
        if (!fresh && mdb_import(file, use, err, sizeof err)) {
            fprintf(stderr, "%s\n", err);
            return 1;
        }
    }
    sqlite3 *db;
    if (sqlite3_open_v2(use, &db, SQLITE_OPEN_READONLY, 0) != SQLITE_OK) {
        fprintf(stderr, "%s\n", sqlite3_errmsg(db));
        return 1;
    }
    jet_register(db);
    int meta = has_meta(db), r = 0;
    if (argc < 3) {
        if (meta) {
            list_rows(db, "SELECT name FROM MSysLB_Tables WHERE name NOT LIKE 'MSys%' ORDER BY name", "T", 0);
            list_rows(db, "SELECT name FROM MSysLB_Queries ORDER BY name", "Q", 0);
        } else {
            list_rows(db, "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name", "T", 0);
            list_rows(db, "SELECT name FROM sqlite_master WHERE type = 'view' ORDER BY name", "Q", 0);
        }
    } else
        r = fields(db, argv[2], meta);
    sqlite3_close(db);
    free(use);
    return r;
}
