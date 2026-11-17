/* odbc.h - ODBC for the ADO provider MSDASQL (odbc.c): plain C, UTF-8 strings.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef LB_ODBC_H
#define LB_ODBC_H

enum { LBO_NULL, LBO_INT, LBO_DBL, LBO_TEXT, LBO_BIN, LBO_DATE, LBO_BOOL, LBO_DEC };

typedef struct lbo_conn lbo_conn;
typedef struct lbo_stmt lbo_stmt;
typedef struct { char state[8]; int native; char msg[1024]; } lbo_err;
typedef struct {
    char name[256], basetable[256], basecol[256];
    int sqltype, decimals, nullable, autoinc, unsignd, updatable;
    unsigned long size, octets;
} lbo_colinfo;
typedef struct { int kind; long long iv; double dv; char *s; int len; } lbo_val;

int lbo_available(void);                        /* the driver manager is there */
int lbo_nerr(void);                             /* the diagnostics of the last call that failed */
const lbo_err *lbo_errs(void);

lbo_conn *lbo_connect(const char *cs, char *outcs, int outlen);
void lbo_disconnect(lbo_conn *c);
int lbo_info(lbo_conn *c, int what, char *buf, int len);          /* SQLGetInfo, a string */
int lbo_info_int(lbo_conn *c, int what, long *out);
int lbo_autocommit(lbo_conn *c, int on);
int lbo_endtran(lbo_conn *c, int commit);
int lbo_drivers(char *buf, int len);

lbo_stmt *lbo_stmt_new(lbo_conn *c);
void lbo_stmt_free(lbo_stmt *s);
int lbo_prepare(lbo_stmt *s, const char *sql);
int lbo_exec_direct(lbo_stmt *s, const char *sql);
int lbo_execute(lbo_stmt *s);
int lbo_nparams(lbo_stmt *s);
int lbo_describe_param(lbo_stmt *s, int i, int *sqltype, unsigned long *size, int *dec, int *nullable);
int lbo_bind(lbo_stmt *s, int i, int kind, long long iv, double dv, const char *text, int len, int sqltype);
int lbo_ncols(lbo_stmt *s);
long lbo_rowcount(lbo_stmt *s);
int lbo_colinfo_get(lbo_stmt *s, int i, lbo_colinfo *ci);
int lbo_fetch(lbo_stmt *s);                     /* 1 a row, 0 no more, -1 error */
int lbo_get(lbo_stmt *s, int i, int sqltype, lbo_val *v);
void lbo_val_free(lbo_val *v);
int lbo_more(lbo_stmt *s);
void lbo_close_cursor(lbo_stmt *s);
int lbo_tables(lbo_stmt *s, const char *cat, const char *sch, const char *tab, const char *types);
int lbo_columns(lbo_stmt *s, const char *cat, const char *sch, const char *tab, const char *col);
int lbo_primary_keys(lbo_stmt *s, const char *cat, const char *sch, const char *tab);

/* SQL types (sql.h) the ADO layer needs without sql.h */
enum { LBS_CHAR = 1, LBS_NUMERIC = 2, LBS_DECIMAL = 3, LBS_INTEGER = 4, LBS_SMALLINT = 5, LBS_FLOAT = 6, LBS_REAL = 7,
       LBS_DOUBLE = 8, LBS_DATE = 9, LBS_TIME = 10, LBS_TIMESTAMP = 11, LBS_VARCHAR = 12, LBS_TYPE_DATE = 91,
       LBS_TYPE_TIME = 92, LBS_TYPE_TIMESTAMP = 93, LBS_LONGVARCHAR = -1, LBS_BINARY = -2, LBS_VARBINARY = -3,
       LBS_LONGVARBINARY = -4, LBS_BIGINT = -5, LBS_TINYINT = -6, LBS_BIT = -7, LBS_WCHAR = -8, LBS_WVARCHAR = -9,
       LBS_WLONGVARCHAR = -10, LBS_GUID = -11 };
/* SQLGetInfo */
enum { LBI_DBMS_NAME = 17, LBI_DBMS_VER = 18, LBI_DRIVER_NAME = 6, LBI_DRIVER_VER = 7, LBI_DATABASE_NAME = 16,
       LBI_DATA_SOURCE_NAME = 2, LBI_SERVER_NAME = 13, LBI_USER_NAME = 47, LBI_IDENTIFIER_QUOTE = 29,
       LBI_DATA_SOURCE_READ_ONLY = 25, LBI_ODBC_SQL_CONFORMANCE = 15 };
#endif
