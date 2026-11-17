/* daoint.h - the DAO objects (dao.c, daoschema.c, daors.c).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#ifndef DAOINT_H
#define DAOINT_H
#include "vbdb.h"
#ifdef _WIN32
#include <string.h>
#include <ctype.h>
static inline char *lb_strcasestr(const char *h, const char *n)   /* mingw has no strcasestr */
{
    size_t k = strlen(n);
    for (; *h; h++)
        if (!strncasecmp(h, n, k)) return (char *)h;
    return k ? 0 : (char *)h;
}
#define strcasestr lb_strcasestr
#endif

extern const vb_class vb_cls_dao_dbengine, vb_cls_dao_workspace, vb_cls_dao_database, vb_cls_dao_tabledef,
    vb_cls_dao_field, vb_cls_dao_index, vb_cls_dao_relation, vb_cls_dao_querydef, vb_cls_dao_parameter,
    vb_cls_dao_recordset, vb_cls_dao_error, vb_cls_dao_coll, vb_cls_dao_property;

/* RecordsetTypeEnum, options */
enum { dbOpenTable = 1, dbOpenDynaset = 2, dbOpenSnapshot = 4, dbOpenForwardOnly = 8, dbOpenDynamic = 16 };
enum { dbDenyWrite = 1, dbDenyRead = 2, dbReadOnly = 4, dbAppendOnly = 8, dbInconsistent = 16, dbConsistent = 32,
       dbSQLPassThrough = 64, dbFailOnError = 128, dbForwardOnly = 256, dbSeeChanges = 512 };
enum { dbEditNone = 0, dbEditInProgress = 1, dbEditAdd = 2 };
enum { dbSystemObject = (int)0x80000002, dbHiddenObject = 1, dbAttachedTable = 0x40000000 };
enum { dbRelationUnique = 1, dbRelationDontEnforce = 2, dbRelationInherited = 4, dbRelationUpdateCascade = 256,
       dbRelationDeleteCascade = 4096, dbRelationLeft = 0x1000000, dbRelationRight = 0x2000000 };

/* ---------------------------------------------------------------- collections */
enum { C_WORKSPACES = 1, C_DATABASES, C_TABLEDEFS, C_QUERYDEFS, C_RELATIONS, C_RECORDSETS, C_FIELDS, C_INDEXES,
       C_IXFIELDS, C_RELFIELDS, C_RSFIELDS, C_PARAMS, C_ERRORS, C_PROPERTIES, C_CONTAINERS };
typedef struct dcoll {
    vbobjh h;
    int kind;
    vbobj owner;                        /* not counted: the owner holds the collection */
    vbobj *items; int n, cap;
    int loaded;
} dcoll;
dcoll *coll_new(int kind, vbobj owner);
void coll_add(dcoll *c, vbobj item);    /* takes a reference */
void coll_remove(dcoll *c, int i);
void coll_clear_items(dcoll *c);
int coll_find(dcoll *c, const char *name);
void dao_param_plain(const char *in, char *out, int cap);
const char *obj_name(vbobj o);

/* ---------------------------------------------------------------- objects */
typedef struct dws {
    vbobjh h;
    char *name, *user;
    dcoll *dbs;
    int trans;                          /* BeginTrans depth */
} dws;

struct drs;
typedef struct ddb {
    vbobjh h;
    sqlite3 *db;
    char *name;                         /* as the program gave it */
    char *file;                         /* the SQLite file */
    int ro, closed, version;            /* version: 30, 40 */
    vbobj ws;                           /* not counted */
    dcoll *tds, *qds, *rels, *rss;
    long affected;
    char *connect;
    int trans;
    /* a database in memory that stands for a server (ADO over ODBC): a change of a row goes to the
     * server first (op 1 insert, 2 update, 3 delete); -1 with the error raised stops it */
    int (*remote)(void *ctx, struct drs *r, int op);
    void *remote_ctx;
} ddb;

typedef struct dfld dfld;
typedef struct dtd {
    vbobjh h;
    vbobj db;                           /* ddb (counted), 0 before Append */
    char *name;
    int appended, attrs;
    dcoll *fields, *indexes;
    char *connect, *srcname, *vrule, *vtext;
    vb_date created, updated;
} dtd;

enum { FK_TABLE = 1, FK_INDEX, FK_REL, FK_RS };
struct dfld {
    vbobjh h;
    int fk;                             /* the kind of field */
    vbobj owner;                        /* counted: tabledef / index / relation / recordset */
    char *name;
    int type, size, attrs, required, allowzero, ord, appended;
    char *defval, *vrule, *vtext, *foreign, *stable, *sfield;
    int col;                            /* FK_RS: the column */
};

typedef struct dix {
    vbobjh h;
    vbobj td;                           /* counted */
    char *name;
    dcoll *fields;
    int primary, unique, required, ignorenulls, foreign, clustered, appended;
} dix;

typedef struct drel {
    vbobjh h;
    vbobj db;                           /* counted, 0 before Append */
    char *name, *table, *ftable;
    int attrs, appended;
    dcoll *fields;
} drel;

typedef struct dpar {
    vbobjh h;
    char *name;                         /* as declared: "[Min ID]" */
    int type, dir;
    vbvar value;
} dpar;

typedef struct dqd {
    vbobjh h;
    vbobj db;                           /* counted */
    char *name, *sql;                   /* the Jet text, as Jet shows it */
    int appended, temp, type;
    dcoll *params;
    long affected;
    int maxrecords, odbctimeout;
    char *connect;
    vb_date created, updated;
} dqd;

/* a column of a recordset */
typedef struct rcol {
    char *name;
    int type, size, attrs;
    char *stable, *sfield;              /* the base column, or 0 */
    int updatable, required, allowzero;
    char *defval;
    int autoinc;
} rcol;
typedef struct rrow { vbvar *v; sqlite3_int64 rid; int del; } rrow;
typedef struct rdata { int refs; rrow *r; int n, cap; } rdata;

typedef struct drs {
    vbobjh h;
    vbobj db;                           /* counted */
    int type, options, lockedits;
    char *source;                       /* what was opened: a table, a query, Jet SQL */
    char *sqlite;                       /* the SQLite text of the rows */
    char *base;                         /* the table a change goes to, or 0 */
    int ncols; rcol *cols;
    rdata *d;
    int cur;                            /* -1 before the first (BOF), d->n after the last (EOF) */
    int reached;                        /* rows read so far (RecordCount of a dynaset) */
    int added;                          /* rows added since */
    int editmode;
    vbvar *ebuf; int *edirty;
    int ebuf_row;                       /* AddNew: the row index the new row gets, Edit: cur */
    int lastmod;                        /* the row of LastModified, -1 */
    int nomatch;
    int closed;
    char *index;                        /* a table-type recordset: its Index */
    int *ixcols, *ixdesc, nix;
    vbs filter, sort;
    dcoll *fields;
    sqlite3 *mirror;                    /* the rows in memory, for Find (built when needed) */
    int mirror_ok;
    vbobj parent_qd;
    sqlite3_int64 new_counter;
    void (*watch)(void *ctx, int what, int action, int *arg);   /* a Data control (Validate, Reposition) */
    void *watch_ctx;
} drs;

typedef struct derr { vbobjh h; vb_i32 number; char *desc, *src; } derr;

typedef struct dprop { vbobjh h; char *name; vbvar value; int type, inherited; } dprop;

/* ---------------------------------------------------------------- shared */
#define DB(o) ((ddb *)vb_oid(o))
vbobj dao_engine(void);
vbobj dao_workspace0(void);
char *dao_str(vbvar v);                 /* a temporary UTF-8 copy */
char *dao_strdup_v(vbvar v);            /* malloc */
vbvar dao_vstr(const char *s);          /* a temporary Variant string */
vbvar dao_retobj(vbobj o);
int dao_missing(vb_i32 argc, vbvar *argv, int i);
int dao_check_db(ddb *d, const char *src);   /* 0 if open; raises 3420 otherwise */
int dao_exec(ddb *d, const char *sql);       /* SQLite statements without results; 0 or the error raised */
int dao_raise_sqlite(ddb *d, int rc, const char *src);
void dao_err_set(int number, const char *desc, const char *src);

/* daoschema.c */
int schema_init(ddb *d);
int schema_has_table(ddb *d, const char *name, char *real, int cap);   /* 1 table, 2 query, 0 none */
dtd *schema_load_table(ddb *d, const char *name);
int schema_create_table(ddb *d, dtd *t);      /* a TableDef appended */
int schema_drop_table(ddb *d, const char *name);
int schema_add_field(ddb *d, const char *table, dfld *f);
int schema_drop_field(ddb *d, const char *table, const char *field);
int schema_create_index(ddb *d, const char *table, dix *ix);
int schema_drop_index(ddb *d, const char *table, const char *index);
int schema_create_relation(ddb *d, drel *r);
int schema_ddl(ddb *d, jsql *q);              /* CREATE TABLE ... from SQL */
void schema_list_tables(ddb *d, dcoll *c);
void schema_list_queries(ddb *d, dcoll *c);
void schema_list_relations(ddb *d, dcoll *c);
int schema_field_info(ddb *d, const char *table, const char *field, rcol *out);
int schema_counter_next(ddb *d, const char *table, sqlite3_int64 *next);   /* 1 if the table has a counter */
void schema_counter_bump(ddb *d, const char *table);
int schema_save_query(ddb *d, dqd *q, int replace);
int schema_drop_query(ddb *d, const char *name);
char *jet_pretty_sql(const char *sql);        /* the SQL text as Jet keeps it */
int schema_text_size(ddb *d, const char *table, const char *field);
void schema_after_change(ddb *d);
int schema_build_all(ddb *d);
int schema_after_import(ddb *d);
dfld *field_new(int fk, vbobj owner, const char *name, int type, int size);
dix *index_new(vbobj td, const char *name);
dtd *tabledef_new(const char *name);
drel *relation_new(const char *name, const char *table, const char *ftable, int attrs);
dqd *querydef_new(vbobj db, const char *name, const char *sql);
dpar *param_new(const char *name, int type);
void querydef_parse_params(dqd *q);
void dao_field_init(vbobj o);
void dao_index_init(vbobj o);
void dao_tabledef_init(vbobj o);
void dao_relation_init(vbobj o);
void dao_querydef_init(vbobj o);
void dao_field_settype(dfld *f, int type);

/* dao.c, for ADO */
vbobj dao_open_path(const char *path, int ro);
vbobj dao_memory_db(void);
int dao_execute_qd(ddb *d, const char *sql, vbobj qd, long *affected);
void dao_db_close(vbobj o);
int dao_translate(ddb *d, const char *sql, jsql *j, const char *const *params);

/* daors.c */
vbobj rs_open(ddb *d, const char *source, int type, int options, int lockedits, vbobj qd, vbvar *pvals, int npvals);
vbobj rs_open_sqlite(ddb *d, const char *sqlite_sql, const char *base, int type, const char *source, int nbinds, vbvar *binds);
vbvar rs_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
vbvar rs_field_value(drs *r, int col, int *err);
int rs_field_let(drs *r, int col, vbvar v);
void rs_close(drs *r);
void rs_clear(vbobj o);
#endif
