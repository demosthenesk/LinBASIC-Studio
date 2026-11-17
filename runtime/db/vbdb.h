/* vbdb.h - the databases of LinBASIC Studio (phase 4.4): DAO and ADO over SQLite.
 * Shared by jetsql.c (Jet SQL -> SQLite, the VBA functions of the SQL), dao.c (the DAO objects)
 * and mdbimport.c (Jet .mdb files copied into SQLite).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef VBDB_H
#define VBDB_H
#include "../vbrt_int.h"
#include "sqlite/sqlite3.h"

/* the DAO field types (DataTypeEnum) */
enum { dbBoolean = 1, dbByte = 2, dbInteger = 3, dbLong = 4, dbCurrency = 5, dbSingle = 6, dbDouble = 7,
       dbDate = 8, dbBinary = 9, dbText = 10, dbLongBinary = 11, dbMemo = 12, dbGUID = 15, dbBigInt = 16,
       dbVarBinary = 17, dbChar = 18, dbNumeric = 19, dbDecimal = 20, dbFloat = 21, dbTime = 22, dbTimeStamp = 23 };
/* FieldAttributeEnum */
enum { dbFixedField = 1, dbVariableField = 2, dbAutoIncrField = 16, dbUpdatableField = 32, dbSystemField = 8192,
       dbHyperlinkField = 32768, dbDescending = 1 };

/* A date in SQLite: text "\001yyyy-mm-dd hh:mm:ss". The mark keeps the type through expressions
 * (Max(D), IIf(.., D, ..)); the text sorts and compares in date order. */
#define JDATE_MARK '\001'
#define JDATE_LEN 20
void jet_date_text(vb_date d, char out[24]);          /* the stored text of a date */
int jet_text_date(const char *s, int n, vb_date *d);  /* stored text -> date (1 if it is one) */

/* a growing C string */
typedef struct jbuf { char *p; int n, cap; } jbuf;
void jb_put(jbuf *b, const char *s, int n);
void jb_s(jbuf *b, const char *s);
void jb_f(jbuf *b, const char *fmt, ...);
void jb_qid(jbuf *b, const char *s, int n);           /* "identifier" */
void jb_qstr(jbuf *b, const char *s, int n);          /* 'string' */
void jb_free(jbuf *b);

/* ---------------------------------------------------------------- Jet SQL */
enum { J_SELECT = 1, J_INSERT, J_UPDATE, J_DELETE, J_CREATE_TABLE, J_CREATE_INDEX, J_DROP_TABLE, J_DROP_INDEX,
       J_ALTER_ADD, J_ALTER_DROP, J_ALTER_CONSTRAINT, J_DROP_CONSTRAINT, J_TRANSFORM };

typedef struct jcol {                   /* a column of CREATE TABLE / ALTER TABLE ADD */
    char *name; int type, size, attrs, notnull;
    int pk, unique;                     /* a column constraint */
    char *cname;                        /* its name */
    char *ref_table, *ref_col;          /* REFERENCES */
} jcol;
typedef struct jcons {                  /* CONSTRAINT name PRIMARY KEY / UNIQUE / FOREIGN KEY (cols) */
    char *name; int kind;               /* 1 primary, 2 unique, 3 foreign */
    char **cols; int *desc; int ncols;
    char *ref_table; char **ref_cols; int nref;
} jcons;
typedef struct jparam { char *name; int type; } jparam;
typedef struct jctype { char name[64]; int type; } jctype;   /* the type of a computed column */

typedef struct jsql {
    int kind;
    char *sql;                          /* SQLite text (select / insert / update / delete) */
    int returns;                        /* a SELECT */
    char *into;                         /* SELECT ... INTO into */
    char *table;                        /* the table of the statement (insert, update, delete, DDL) */
    char *index;                        /* CREATE / DROP INDEX */
    jcol *cols; int ncols;              /* CREATE TABLE, ALTER ADD; CREATE INDEX: name + attrs (1 desc) */
    jcons *cons; int ncons;
    int unique, primary, disallow_null, ignore_null;   /* CREATE INDEX */
    jparam *params; int nparams;        /* PARAMETERS */
    char **binds; int nbinds;           /* ?1 .. ?n: the parameter of each, by name */
    jctype *ctypes; int nctypes;        /* the types of the computed columns (Expr1000, AS x) */
    char *base;                         /* a single-table SELECT: the table (updatable) */
    char *base_alias;
    int err; char errmsg[512];          /* a Jet error number and its text */
} jsql;

/* params: names that are parameters (a QueryDef's), 0-terminated, or 0 */
int jet_translate(const char *jet, jsql *q, const char *const *params);
extern int jet_ansi;                    /* set by ADO: LIKE with % and _ (ANSI-92 wildcards) */
void jsql_free(jsql *q);
/* a criteria (Filter, FindFirst) or an expression: SQLite text, in a boolean context or not */
int jet_expr(const char *jet, jbuf *out, int boolean, char *errmsg, int errcap);
/* the type of a Jet DDL type name (TEXT, LONG ...): dbText ..., 0 if unknown; *autoinc for COUNTER */
int jet_type_name(const char *s, int n, int *autoinc);
int jet_is_function(const char *name);

/* the VBA functions of Jet SQL, the Jet LIKE and the text sort order on a connection */
void jet_register(sqlite3 *db);
/* the DAO type of a column of a table (for the types of expressions), set by the caller */
extern int (*jet_coltype_fn)(void *ctx, const char *qual, const char *col);
extern void *jet_coltype_ctx;

/* SQLite value <-> Variant (dtype: the DAO type it should have, 0 = by the value) */
vbvar jet_value(sqlite3_value *v, int dtype);
vbvar jet_column(sqlite3_stmt *st, int i, int dtype);
int jet_bind(sqlite3_stmt *st, int i, vbvar v, int dtype);   /* 0 or a Jet error */
void jet_result(sqlite3_context *cx, vbvar v);
vbvar jet_arg(sqlite3_value *v);
/* the DAO type of a declared SQLite column type ("LONG_INT" ...), its size */
int jet_decl_type(const char *decl, int *size, int *attrs);
const char *jet_type_decl(int type, int autoinc, char *buf, int cap);

/* ---------------------------------------------------------------- errors */
void dao_raise(int number, const char *source, const char *fmt, ...);
int jet_sqlite_error(sqlite3 *db, int rc, char *msg, int cap);  /* SQLite error -> Jet number + text */
extern const char *dao_src;             /* "DAO.Database" ... the Source of the next error */

/* ---------------------------------------------------------------- .mdb */
int mdb_is_jet(const char *path);       /* 3 or 4 for a Jet file, 0 otherwise */
int mdb_import(const char *mdb, const char *sqlite_path, char *err, int cap);
/* the schema tables of a LinBASIC database (dao.c) */
int dao_schema_create(sqlite3 *db);

#endif
