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
                int view = !strcasecmp(type, "VIEW");
                if ((!strcasecmp(type, "TABLE") && !pass) || (view && pass)) {
                    char *name = vb_stoutf8(vb_v2s(vb_late_o(vb_v2obj(vb_late_o(rs, "fields", VBM_GET, 1, n)), "value", VBM_GET, 0, 0)));
                    printf("%s\t%s\n", view ? "Q" : "T", name);
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

int main(int argc, char **argv)
{
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
