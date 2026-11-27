/* dao.c - Microsoft DAO 3.6 written again over SQLite: DBEngine, Workspace, Database, TableDef,
 * Field, Index, Relation, QueryDef, Parameter, Error and their collections, with the members, the
 * defaults, the errors (numbers, texts, Source) of the original. The collections are filled when the
 * program first uses them and stay as they are until Refresh, Append or Delete, as Jet's do (a table
 * made by db.Execute "CREATE TABLE" is not in a TableDefs read before it).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "daoint.h"
#include "../vbos.h"
#include <ctype.h>
#include <sys/stat.h>
#ifndef _WIN32
#include <dirent.h>
#endif

const char *dao_src = "DAO.Database";

/* ---------------------------------------------------------------- helpers */
int dao_missing(vb_i32 argc, vbvar *argv, int i) { return i >= argc || vb_ismissing(argv[i]); }
char *dao_str(vbvar v) { vbs s = vb_v2s(v); return vb_errpend ? (char *)"" : vb_stoutf8(s); }
char *dao_strdup_v(vbvar v) { return strdup(dao_str(v)); }
vbvar dao_vstr(const char *s) { return vb_v_s(vb_sfromutf8(s ? s : "")); }
vbvar dao_retobj(vbobj o) { return vb_vtmp(vb_v_obj(o)); }
static vb_i32 along(vb_i32 argc, vbvar *argv, int i, vb_i32 d) { return dao_missing(argc, argv, i) ? d : vb_v2l(argv[i]); }
static char *sdup(const char *s) { return s ? strdup(s) : 0; }

/* ---------------------------------------------------------------- errors */
static vbobj engine;
static dcoll *err_list(void);

void dao_err_set(int number, const char *desc, const char *src)
{
    dcoll *c = err_list();
    coll_clear_items(c);
    derr *e = (derr *)vb_onew(&vb_cls_dao_error);
    e->number = number;
    e->desc = strdup(desc);
    e->src = strdup(src);
    coll_add(c, (vbobj)e);
}

void dao_raise(int number, const char *source, const char *fmt, ...)
{
    char msg[1024];
    va_list ap; va_start(ap, fmt);
    vsnprintf(msg, sizeof msg, fmt, ap);
    va_end(ap);
    if (!source) source = dao_src;
    dao_err_set(number, msg, source);
    vb_raise_full(number, vb_sfromutf8(source), vb_sfromutf8(msg));
}

/* a SQLite error as the Jet error it stands for */
int jet_sqlite_error(sqlite3 *db, int rc, char *msg, int cap)
{
    const char *m = sqlite3_errmsg(db);
    int ext = sqlite3_extended_errcode(db);
    if ((rc & 0xff) == SQLITE_CONSTRAINT || (ext & 0xff) == SQLITE_CONSTRAINT) {
        const char *lb = strstr(m, "LB");
        if (lb && isdigit((unsigned char)lb[2])) {
            int n = atoi(lb + 2);
            const char *arg = strchr(lb, ' ');
            arg = arg ? arg + 1 : "";
            switch (n) {
            case 3315: snprintf(msg, (size_t)cap, "Field '%s' cannot be a zero-length string.", arg); break;
            case 3058: snprintf(msg, (size_t)cap, "Index or primary key cannot contain a Null value."); break;
            case 3316: snprintf(msg, (size_t)cap, "Index or primary key cannot contain a Null value."); n = 3058; break;
            case 3201: snprintf(msg, (size_t)cap, "You cannot add or change a record because a related record is required in table '%s'.", arg); break;
            case 3200: snprintf(msg, (size_t)cap, "The record cannot be deleted or changed because table '%s' includes related records.", arg); break;
            default: snprintf(msg, (size_t)cap, "%s", arg);
            }
            return n;
        }
        if (!strncmp(m, "NOT NULL constraint failed: ", 28)) {
            snprintf(msg, (size_t)cap, "The field '%s' cannot contain a Null value because the Required property for this field is set to True.  Enter a value in this field.", m + 28);
            return 3314;
        }
        if (!strncmp(m, "UNIQUE constraint failed", 24) || ext == SQLITE_CONSTRAINT_PRIMARYKEY) {
            snprintf(msg, (size_t)cap, "The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again.");
            return 3022;
        }
    }
    if (!strncmp(m, "no such table: ", 15)) {
        const char *t = m + 15;
        if (!strncmp(t, "lbext", 5)) { const char *dot = strchr(t, '.'); if (dot) t = dot + 1; }   /* another database's */
        snprintf(msg, (size_t)cap, "The Microsoft Jet database engine cannot find the input table or query '%s'.  Make sure it exists and that its name is spelled correctly.", t);
        return 3078;
    }
    if (!strncmp(m, "no such column: ", 16)) {
        snprintf(msg, (size_t)cap, "Too few parameters. Expected 1.");
        return 3061;
    }
    if (!strncmp(m, "no such function: jet_", 22)) {
        snprintf(msg, (size_t)cap, "Undefined function '%s' in expression.", m + 22);
        return 3085;
    }
    if (m[0] == '#') {                           /* a run-time error of a VBA function in the SQL */
        int n = atoi(m + 1);
        const char *t = strchr(m, ' ');
        snprintf(msg, (size_t)cap, "%s", t ? t + 1 : "");
        if (n == 13) { snprintf(msg, (size_t)cap, "Data type mismatch in criteria expression."); return 3464; }
        if (n == 11) { snprintf(msg, (size_t)cap, "Division by zero."); return 3071; }
        return n;
    }
    if (rc == SQLITE_READONLY) { snprintf(msg, (size_t)cap, "Cannot update.  Database or object is read-only."); return 3027; }
    if (rc == SQLITE_BUSY || rc == SQLITE_LOCKED) { snprintf(msg, (size_t)cap, "Could not update; currently locked."); return 3218; }
    if (rc == SQLITE_NOTADB || rc == SQLITE_CORRUPT) { snprintf(msg, (size_t)cap, "Unrecognized database format."); return 3343; }
    snprintf(msg, (size_t)cap, "%s", m);
    return 3146;
}

int dao_raise_sqlite(ddb *d, int rc, const char *src)
{
    char msg[1024];
    int n = jet_sqlite_error(d->db, rc, msg, sizeof msg);
    dao_raise(n, src, "%s", msg);
    return n;
}

int dao_exec(ddb *d, const char *sql)
{
    char *em = 0;
    int rc = sqlite3_exec(d->db, sql, 0, 0, &em);
    if (rc != SQLITE_OK) {
        sqlite3_free(em);
        return dao_raise_sqlite(d, rc, dao_src);
    }
    return 0;
}

int dao_check_db(ddb *d, const char *src)
{
    if (!d->closed) return 0;
    dao_raise(3420, src, "Object invalid or no longer set.");
    return 3420;
}

/* ---------------------------------------------------------------- collections */
static const char *const COLL_NAMES[] = { "", "Workspaces", "Databases", "TableDefs", "QueryDefs", "Relations", "Recordsets",
    "Fields", "Indexes", "Fields", "Fields", "Fields", "Parameters", "Errors", "Properties", "Containers" };

dcoll *coll_new(int kind, vbobj owner)
{
    dcoll *c = (dcoll *)vb_onew(&vb_cls_dao_coll);
    vb_oaddref((vbobj)c);
    c->kind = kind;
    c->owner = owner;
    return c;
}
void coll_add(dcoll *c, vbobj item)
{
    if (c->n == c->cap) { c->cap = c->cap ? c->cap * 2 : 8; c->items = realloc(c->items, sizeof(vbobj) * (size_t)c->cap); }
    c->items[c->n++] = item;
    vb_oaddref(item);
}
void coll_remove(dcoll *c, int i)
{
    vbobj o = c->items[i];
    memmove(c->items + i, c->items + i + 1, sizeof(vbobj) * (size_t)(c->n - i - 1));
    c->n--;
    vb_orelease(o);
}
void coll_clear_items(dcoll *c)
{
    while (c->n) coll_remove(c, c->n - 1);
}

const char *obj_name(vbobj o)
{
    o = vb_oid(o);
    const vb_class *k = o->cls;
    if (k == &vb_cls_dao_workspace) return ((dws *)o)->name;
    if (k == &vb_cls_dao_database) return ((ddb *)o)->name;
    if (k == &vb_cls_dao_tabledef) return ((dtd *)o)->name;
    if (k == &vb_cls_dao_field) return ((dfld *)o)->name;
    if (k == &vb_cls_dao_index) return ((dix *)o)->name;
    if (k == &vb_cls_dao_relation) return ((drel *)o)->name;
    if (k == &vb_cls_dao_querydef) return ((dqd *)o)->name;
    if (k == &vb_cls_dao_parameter) return ((dpar *)o)->name;
    if (k == &vb_cls_dao_recordset) return ((drs *)o)->source ? ((drs *)o)->source : "";
    if (k == &vb_cls_dao_property) return ((dprop *)o)->name;
    return "";
}
/* a parameter's name without its brackets: [Forms]![Dialog]![Date] -> Forms!Dialog!Date */
void dao_param_plain(const char *in, char *out, int cap)
{
    int k = 0;
    for (; *in && k < cap - 1; in++) if (*in != '[' && *in != ']') out[k++] = *in;
    out[k] = 0;
}
int coll_find(dcoll *c, const char *name)
{
    for (int i = 0; i < c->n; i++) {
        const char *n = obj_name(c->items[i]);
        if (!strcasecmp(n, name)) return i;
        /* a parameter declared [Min ID] is "Min ID" too */
        if (strchr(n, '[')) {
            char a[256], b[256];
            dao_param_plain(n, a, sizeof a); dao_param_plain(name, b, sizeof b);
            if (!strcasecmp(a, b)) return i;
        }
    }
    return -1;
}

static void coll_load(dcoll *c);
static int props_append(dcoll *c, vbobj it);     /* daoprops.inc */
static int props_delete(dcoll *c, int i);
static vbvar prop_create(vbobj owner, vb_i32 argc, vbvar *argv);
static vbvar props_get(vbobj owner, vb_i32 kind, vb_i32 argc, vbvar *argv);
static vbvar coll_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);

static void coll_clear(vbobj o)
{
    dcoll *c = (dcoll *)o;
    for (int i = 0; i < c->n; i++) vb_orelease(c->items[i]);
    free(c->items);
    c->items = 0; c->n = c->cap = 0;
}

static vb_i32 coll_each_count(vbobj o) { dcoll *c = (dcoll *)vb_oid(o); coll_load(c); return c->n; }
static vbvar coll_each_item(vbobj o, vb_i32 i)
{
    dcoll *c = (dcoll *)vb_oid(o);
    return i >= 0 && i < c->n ? dao_retobj(c->items[i]) : vb_v_empty();
}

static const char *coll_src(dcoll *c)
{
    static char buf[40];
    snprintf(buf, sizeof buf, "DAO.%s", COLL_NAMES[c->kind]);
    return buf;
}

static int coll_index(dcoll *c, vbvar key)
{
    if (key.vt == VT_BSTR) return coll_find(c, dao_str(key));
    vb_i32 save = vb_errpend;
    vb_i32 i = vb_v2l(key);
    if (vb_errpend && !save) return -1;
    return i >= 0 && i < c->n ? i : -1;
}

static ddb *coll_db(dcoll *c)
{
    vbobj o = c->owner;
    if (!o) return 0;
    if (o->cls == &vb_cls_dao_database) return (ddb *)o;
    if (o->cls == &vb_cls_dao_tabledef) return ((dtd *)o)->db ? DB(((dtd *)o)->db) : 0;
    return 0;
}

/* a collection read from the database the first time it is used */
static void coll_load(dcoll *c)
{
    if (c->loaded) return;
    c->loaded = 1;
    ddb *d = c->owner && c->owner->cls == &vb_cls_dao_database ? (ddb *)c->owner : 0;
    if (!d || d->closed) return;
    switch (c->kind) {
    case C_TABLEDEFS: schema_list_tables(d, c); break;
    case C_QUERYDEFS: schema_list_queries(d, c); break;
    case C_RELATIONS: schema_list_relations(d, c); break;
    }
}

static int tabledef_append_field(dtd *t, dfld *f);
static int tabledef_append_index(dtd *t, dix *ix);
static int db_append_tabledef(ddb *d, dtd *t);
static int db_append_querydef(ddb *d, dqd *q);
static int db_append_relation(ddb *d, drel *r);

static vbvar coll_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dcoll *c = (dcoll *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    ddb *d = coll_db(c);
    if (d && d->closed) { dao_raise(3420, coll_src(c), "Object invalid or no longer set."); return vb_v_empty(); }
    coll_load(c);
    if (!strcmp(name, "count")) { int n = c->n + c->extra; c->extra = 0; return vb_v_i((vb_i16)n); }
    if (!strcmp(name, "item") || !strcmp(name, "")) {
        if (argc < 1 && !set && c->kind == C_IXFIELDS) {
            /* the value of Index.Fields: "+A;-B" (VisData: idxTo.Fields = idxFrom.Fields) */
            jbuf b = { 0 };
            for (int i = 0; i < c->n; i++) {
                dfld *f = (dfld *)c->items[i];
                jb_s(&b, i ? ";" : ""); jb_s(&b, (f->attrs & dbDescending) ? "-" : "+"); jb_s(&b, f->name);
            }
            vbvar r = dao_vstr(b.p ? b.p : "");
            jb_free(&b);
            return r;
        }
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = coll_index(c, argv[0]);
        if (i < 0 && c->kind == C_PROPERTIES) { dao_raise(3270, coll_src(c), "Property not found."); return vb_v_empty(); }
        if (i < 0) { dao_raise(3265, coll_src(c), "Item not found in this collection."); return vb_v_empty(); }
        vbobj it = c->items[i];
        if (set || argc > 1) {                  /* Fields("x") = v, Parameters("p") = v: the default Value */
            return vb_late_o(it, "value", kind, argc - 1, argv + 1);
        }
        return dao_retobj(it);
    }
    if (!strcmp(name, "_newenum")) return dao_retobj(o);
    if (!strcmp(name, "refresh")) {
        if (c->kind == C_TABLEDEFS || c->kind == C_QUERYDEFS || c->kind == C_RELATIONS) {
            /* the objects already read stay; new ones are added, removed ones go */
            dcoll *fresh = coll_new(c->kind, c->owner);
            fresh->loaded = 0;
            coll_load(fresh);
            dcoll *keep = coll_new(c->kind, 0);
            keep->loaded = 1;
            for (int i = 0; i < fresh->n; i++) {
                int j = coll_find(c, obj_name(fresh->items[i]));
                coll_add(keep, j >= 0 ? c->items[j] : fresh->items[i]);
            }
            coll_clear_items(c);
            for (int i = 0; i < keep->n; i++) coll_add(c, keep->items[i]);
            vb_orelease((vbobj)keep);
            vb_orelease((vbobj)fresh);
        }
        return vb_v_empty();
    }
    if (!strcmp(name, "append")) {
        if (argc != 1) { vb_raise(450); return vb_v_empty(); }
        vbobj it = vb_v2obj(argv[0]);
        if (!it) return vb_v_empty();
        it = vb_oid(it);
        if (c->kind == C_PROPERTIES) {          /* (daoprops.inc) */
            if (!props_append(c, it)) coll_add(c, it);
            return vb_v_empty();
        }
        if (coll_find(c, obj_name(it)) >= 0 && c->kind != C_ERRORS) {
            const char *cs = coll_src(c);
            if (c->kind == C_TABLEDEFS) dao_raise(3010, cs, "Table '%s' already exists.", obj_name(it));
            else if (c->kind == C_FIELDS || c->kind == C_IXFIELDS || c->kind == C_RELFIELDS) dao_raise(3191, cs, "Cannot define field more than once.");
            else dao_raise(3367, cs, "Cannot append. An object with that name already exists in the collection.");
            return vb_v_empty();
        }
        int r = 0;
        switch (c->kind) {
        case C_TABLEDEFS:
            if (it->cls != &vb_cls_dao_tabledef) { vb_raise(13); return vb_v_empty(); }
            r = db_append_tabledef((ddb *)c->owner, (dtd *)it);
            break;
        case C_QUERYDEFS:
            if (it->cls != &vb_cls_dao_querydef) { vb_raise(13); return vb_v_empty(); }
            r = db_append_querydef((ddb *)c->owner, (dqd *)it);
            break;
        case C_RELATIONS:
            if (it->cls != &vb_cls_dao_relation) { vb_raise(13); return vb_v_empty(); }
            r = db_append_relation((ddb *)c->owner, (drel *)it);
            break;
        case C_FIELDS:
            if (it->cls != &vb_cls_dao_field) { vb_raise(13); return vb_v_empty(); }
            r = tabledef_append_field((dtd *)c->owner, (dfld *)it);
            break;
        case C_INDEXES:
            if (it->cls != &vb_cls_dao_index) { vb_raise(13); return vb_v_empty(); }
            r = tabledef_append_index((dtd *)c->owner, (dix *)it);
            break;
        case C_IXFIELDS: case C_RELFIELDS: {
            if (it->cls != &vb_cls_dao_field) { vb_raise(13); return vb_v_empty(); }
            int appended = c->owner->cls == &vb_cls_dao_index ? ((dix *)c->owner)->appended : ((drel *)c->owner)->appended;
            if (appended) { dao_raise(3219, coll_src(c), "Invalid operation."); return vb_v_empty(); }
            break; }
        default:
            dao_raise(3219, coll_src(c), "Invalid operation.");
            return vb_v_empty();
        }
        if (!r) coll_add(c, it);
        return vb_v_empty();
    }
    if (!strcmp(name, "delete")) {
        if (argc != 1) { vb_raise(450); return vb_v_empty(); }
        char *nm = dao_str(argv[0]);
        int i = coll_find(c, nm);
        if (i < 0 && c->kind == C_PROPERTIES) { dao_raise(3270, coll_src(c), "Property not found."); return vb_v_empty(); }
        if (i < 0) { dao_raise(3265, coll_src(c), "Item not found in this collection."); return vb_v_empty(); }
        int r = 0;
        switch (c->kind) {
        case C_PROPERTIES: r = props_delete(c, i); break;
        case C_TABLEDEFS: r = ((ddb *)c->owner)->odbc ? dao_odbc_drop_table((ddb *)c->owner, obj_name(c->items[i]))
                                                       : schema_drop_table((ddb *)c->owner, obj_name(c->items[i])); break;
        case C_QUERYDEFS: r = schema_drop_query((ddb *)c->owner, obj_name(c->items[i])); break;
        case C_RELATIONS: {
            ddb *dd = (ddb *)c->owner;
            jbuf s = { 0 };
            jb_s(&s, "DELETE FROM MSysLB_Relations WHERE name = "); jb_qstr(&s, obj_name(c->items[i]), -1);
            r = dao_exec(dd, s.p);
            jb_free(&s);
            break; }
        case C_FIELDS: {
            dtd *t = (dtd *)c->owner;
            if (t->appended) r = DB(t->db)->odbc ? dao_odbc_drop_field(DB(t->db), t->name, obj_name(c->items[i])) : schema_drop_field(DB(t->db), t->name, obj_name(c->items[i]));
            break; }
        case C_INDEXES: {
            dtd *t = (dtd *)c->owner;
            if (t->appended) r = DB(t->db)->odbc ? dao_odbc_drop_index(DB(t->db), t->name, obj_name(c->items[i])) : schema_drop_index(DB(t->db), t->name, obj_name(c->items[i]));
            break; }
        case C_IXFIELDS: case C_RELFIELDS: break;
        default: dao_raise(3219, coll_src(c), "Invalid operation."); return vb_v_empty();
        }
        if (!r) coll_remove(c, i);
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- DBEngine */
typedef struct deng { vbobjh h; dcoll *wss, *errs; char *inipath, *systemdb, *defuser, *defpwd; int logintimeout, deftype; } deng;
static dws *ws_new(const char *name, const char *user);

static dcoll *err_list(void) { return ((deng *)vb_oid(dao_engine()))->errs; }

vbobj dao_engine(void)
{
    if (!engine) {
        deng *e = (deng *)vb_onew(&vb_cls_dao_dbengine);
        vb_oaddref((vbobj)e);
        e->wss = coll_new(C_WORKSPACES, (vbobj)e);
        e->wss->loaded = 1;
        e->errs = coll_new(C_ERRORS, (vbobj)e);
        e->errs->loaded = 1;
        e->inipath = strdup(""); e->systemdb = strdup(""); e->defuser = strdup("admin"); e->defpwd = strdup("");
        e->logintimeout = 20; e->deftype = 2;
        engine = (vbobj)e;
        dws *w = ws_new("#Default Workspace#", "admin");
        coll_add(e->wss, (vbobj)w);
        vb_orelease((vbobj)w);
    }
    return engine;
}
vbobj dao_workspace0(void) { return ((deng *)vb_oid(dao_engine()))->wss->items[0]; }

/* the members of the DBEngine a program uses without DBEngine. (OpenDatabase, BeginTrans ...) */
vbvar vb_dao_engine_member(const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv)
{
    return vb_late_o(dao_engine(), name, kind, argc, argv);
}
vbobj vb_dao_dbengine(void) { return dao_engine(); }

static vbvar ws_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);

static vbvar engine_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    deng *e = (deng *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    dao_src = "DAO.DBEngine";
    if (!strcmp(name, "workspaces") || !strcmp(name, "")) {
        if (argc == 0) return dao_retobj((vbobj)e->wss);
        int found2;
        return coll_member((vbobj)e->wss, "item", kind, argc, argv, &found2);
    }
    if (!strcmp(name, "properties")) return props_get(o, kind, argc, argv);
    if (!strcmp(name, "errors")) {
        if (argc == 0) return dao_retobj((vbobj)e->errs);
        int found2;
        return coll_member((vbobj)e->errs, "item", kind, argc, argv, &found2);
    }
    if (!strcmp(name, "version")) return dao_vstr("3.6");
    if (!strcmp(name, "inipath") || !strcmp(name, "systemdb") || !strcmp(name, "defaultuser") || !strcmp(name, "defaultpassword")) {
        char **p = name[0] == 'i' ? &e->inipath : name[0] == 's' ? &e->systemdb : name[7] == 'u' ? &e->defuser : &e->defpwd;
        if (set) { free(*p); *p = dao_strdup_v(argv[argc - 1]); return vb_v_empty(); }
        if (p == &e->systemdb && (!*p || !**p)) return dao_vstr("system.mdb");   /* (Jet's, VB6 d_dao4) */
        return dao_vstr(*p);
    }
    if (!strcmp(name, "logintimeout")) { if (set) { e->logintimeout = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_i((vb_i16)e->logintimeout); }
    if (!strcmp(name, "defaulttype")) { if (set) { e->deftype = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(e->deftype); }
    if (!strcmp(name, "createworkspace")) {
        char *n = argc > 0 ? dao_str(argv[0]) : "", *u = argc > 1 ? dao_str(argv[1]) : "admin";
        dws *w = ws_new(n, u);
        return dao_retobj((vbobj)w);
    }
    if (!strcmp(name, "idle") || !strcmp(name, "setoption") || !strcmp(name, "registerdatabase")) return vb_v_empty();
    if (!strcmp(name, "compactdatabase") || !strcmp(name, "repairdatabase")) {
        /* a SQLite file: VACUUM INTO the new name, or nothing */
        char *src = argc > 0 ? dao_str(argv[0]) : "";
        if (name[0] == 'c' && argc > 1) {
            char a[1024], b[1024];
            snprintf(a, sizeof a, "%s", src); snprintf(b, sizeof b, "%s", dao_str(argv[1]));
#ifndef _WIN32
            for (char *p = a; *p; p++) if (*p == '\\') *p = '/';
            for (char *p = b; *p; p++) if (*p == '\\') *p = '/';
#endif
            sqlite3 *db;
            if (sqlite3_open_v2(a, &db, SQLITE_OPEN_READONLY, 0) == SQLITE_OK) {
                jbuf s = { 0 };
                jb_s(&s, "VACUUM INTO "); jb_qstr(&s, b, -1);
                sqlite3_exec(db, s.p, 0, 0, 0);
                jb_free(&s);
            }
            sqlite3_close(db);
        }
        return vb_v_empty();
    }
    /* the members of the default workspace */
    return ws_member(dao_workspace0(), name, kind, argc, argv, found);
}

/* ---------------------------------------------------------------- Workspace */
static dws *ws_new(const char *name, const char *user)
{
    dws *w = (dws *)vb_onew(&vb_cls_dao_workspace);
    vb_oaddref((vbobj)w);
    w->name = strdup(name);
    w->user = strdup(user);
    w->dbs = coll_new(C_DATABASES, (vbobj)w);
    w->dbs->loaded = 1;
    return w;
}

static void ws_clear(vbobj o)
{
    dws *w = (dws *)o;
    free(w->name); free(w->user);
    if (w->dbs) vb_orelease((vbobj)w->dbs);
}

static void db_close(ddb *d);
static vbobj db_open(dws *w, const char *path, int exclusive, int ro, const char *connect);
static vbobj db_create(dws *w, const char *path, const char *locale, int options);

static void odbc_trans(ddb *d, int what);       /* daoodbc.inc */
static int ws_trans(dws *w, int what)
{
    /* 0 begin, 1 commit, 2 rollback: every open database of the workspace */
    if (what && w->trans == 0) {
        dao_raise(3034, "DAO.Workspace", "You tried to commit or rollback a transaction without first beginning a transaction.");
        return -1;
    }
    char sp[32];
    snprintf(sp, sizeof sp, "lbtrans%d", what ? w->trans : w->trans + 1);
    for (int i = 0; i < w->dbs->n; i++) {
        ddb *d = (ddb *)vb_oid(w->dbs->items[i]);
        if (d->closed) continue;
        char sql[80];
        if (what == 0) snprintf(sql, sizeof sql, "SAVEPOINT %s", sp);
        else if (what == 1) snprintf(sql, sizeof sql, "RELEASE %s", sp);
        else snprintf(sql, sizeof sql, "ROLLBACK TO %s; RELEASE %s", sp, sp);
        sqlite3_exec(d->db, sql, 0, 0, 0);
        odbc_trans(d, what);                    /* and the data sources of ODBC it uses */
    }
    w->trans += what ? -1 : 1;
    return 0;
}

static vbvar ws_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dws *w = (dws *)vb_oid(o);
    *found = 1;
    dao_src = "DAO.Workspace";
    if (!strcmp(name, "databases") || !strcmp(name, "")) {
        if (argc == 0) return dao_retobj((vbobj)w->dbs);
        int f2;
        return coll_member((vbobj)w->dbs, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "name")) return dao_vstr(w->name);
    if (!strcmp(name, "username")) return dao_vstr(w->user);
    if (!strcmp(name, "type")) return vb_v_l(2);                /* dbUseJet */
    if (!strcmp(name, "isolateodbctrans")) return vb_v_i(0);
    if (!strcmp(name, "logintimeout")) return vb_v_l(20);
    if (!strcmp(name, "defaultcursordriver")) return vb_v_l(0);
    if (!strcmp(name, "opendatabase")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        char *p = dao_str(argv[0]);
        int excl = 0, ro = 0;
        if (!dao_missing(argc, argv, 1)) excl = vb_v2bool(argv[1]) != 0 && argv[1].vt != VT_I4 ? 1 : 0;
        if (!dao_missing(argc, argv, 2)) ro = vb_v2bool(argv[2]) != 0;
        char *cn = dao_missing(argc, argv, 3) ? "" : dao_str(argv[3]);
        vbobj d = db_open(w, p, excl, ro, cn);
        return d ? dao_retobj(d) : vb_v_empty();
    }
    if (!strcmp(name, "createdatabase")) {
        if (argc < 2) { vb_raise(450); return vb_v_empty(); }
        vbobj d = db_create(w, dao_str(argv[0]), dao_str(argv[1]), along(argc, argv, 2, 0));
        return d ? dao_retobj(d) : vb_v_empty();
    }
    if (!strcmp(name, "begintrans")) { ws_trans(w, 0); return vb_v_empty(); }
    if (!strcmp(name, "committrans")) { ws_trans(w, 1); return vb_v_empty(); }
    if (!strcmp(name, "rollback")) { ws_trans(w, 2); return vb_v_empty(); }
    if (!strcmp(name, "close")) {
        while (w->dbs->n) db_close((ddb *)vb_oid(w->dbs->items[0]));
        return vb_v_empty();
    }
    if (!strcmp(name, "properties")) return props_get(o, kind, argc, argv);
    if (!strcmp(name, "users") || !strcmp(name, "groups") || !strcmp(name, "connections")) {
        dcoll *c = coll_new(C_PROPERTIES, 0);
        c->loaded = 1;
        return dao_retobj((vbobj)c);
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Database */
static char *native_path(const char *p)
{
    char *s = strdup(p);
#ifndef _WIN32
    for (char *q = s; *q; q++) if (*q == '\\') *q = '/';
    struct stat st;
    if (stat(s, &st) != 0) {
        /* a program of Windows names its files in any case ("biblio.mdb" for BIBLIO.MDB): each part
           of the path is looked for without regard to case */
        char out[4096] = "";
        const char *q = s;
        if (*q == '/') { strcpy(out, "/"); q++; }
        int ok = 1;
        while (*q && ok) {
            const char *e = strchr(q, '/');
            size_t n = e ? (size_t)(e - q) : strlen(q);
            char part[1024]; snprintf(part, sizeof part, "%.*s", (int)n, q);
            char trial[4096]; snprintf(trial, sizeof trial, "%s%s", out, part);
            if (stat(trial, &st) != 0) {
                DIR *dh = opendir(*out ? out : ".");
                struct dirent *de; int hit = 0;
                while (dh && (de = readdir(dh))) if (!strcasecmp(de->d_name, part)) { snprintf(part, sizeof part, "%s", de->d_name); hit = 1; break; }
                if (dh) closedir(dh);
                if (!hit) ok = 0;
            }
            size_t L = strlen(out);
            snprintf(out + L, sizeof out - L, "%s%s", part, e ? "/" : "");
            q = e ? e + 1 : q + n;
        }
        if (ok) { free(s); s = strdup(out); }
        else if (isalpha((unsigned char)p[0]) && p[1] == ':') {
            /* a Windows path ("C:\Program Files\...\NWind.mdb") on Linux: the file of that name next
               to the program, or in the current folder */
            const char *base = strrchr(s, '/');
            base = base ? base + 1 : s + 2;
            char *dir = vb_stoutf8(vb_app_path());
            const char *dirs[2] = { dir, "." };
            for (int k = 0; k < 2; k++) {
                char trial[4096]; snprintf(trial, sizeof trial, "%s/%s", dirs[k], base);
                char *r = native_path(trial);
                if (stat(r, &st) == 0) { free(s); return r; }
                free(r);
            }
        }
    }
#endif
    return s;
}
static int file_exists(const char *p) { struct stat st; return stat(p, &st) == 0 && !S_ISDIR(st.st_mode); }

static ddb *db_new(dws *w, const char *name)
{
    ddb *d = (ddb *)vb_onew(&vb_cls_dao_database);
    vb_oaddref((vbobj)d);
    d->name = strdup(name);
    d->ws = (vbobj)w;
    d->tds = coll_new(C_TABLEDEFS, (vbobj)d);
    d->qds = coll_new(C_QUERYDEFS, (vbobj)d);
    d->rels = coll_new(C_RELATIONS, (vbobj)d);
    d->rss = coll_new(C_RECORDSETS, (vbobj)d);
    d->rss->loaded = 1;
    d->connect = strdup("");
    d->version = 40;
    return d;
}

static int db_connect(ddb *d, const char *file, int ro, int create)
{
    int flags = ro ? SQLITE_OPEN_READONLY : SQLITE_OPEN_READWRITE;
    if (create) flags |= SQLITE_OPEN_CREATE;
    int rc = sqlite3_open_v2(file, &d->db, flags, 0);
    if (rc != SQLITE_OK) return rc;
    sqlite3_busy_timeout(d->db, 5000);
    jet_register(d->db);
    sqlite3_exec(d->db, "PRAGMA legacy_alter_table = ON; PRAGMA foreign_keys = OFF; PRAGMA recursive_triggers = OFF", 0, 0, 0);
    /* writes as Jet makes them: not forced to the disk one by one (Jet writes lazily), each still all or nothing
     * (the journal). Was: every Update of a Recordset synced to the disk, a few dozen a second (robot 2026-11-12:
     * a spell checker that adds 500,000 words with AddNew / Update). A power cut can damage the file, as an .mdb. */
    if (!ro) sqlite3_exec(d->db, "PRAGMA synchronous = OFF; PRAGMA journal_mode = TRUNCATE", 0, 0, 0);
    d->file = strdup(file);
    d->ro = ro;
    return SQLITE_OK;
}

static void read_version(ddb *d)
{
    sqlite3_stmt *st = 0;
    if (sqlite3_prepare_v2(d->db, "SELECT value FROM MSysLB_Props WHERE name = 'Version'", -1, &st, 0) == SQLITE_OK) {
        if (sqlite3_step(st) == SQLITE_ROW) d->version = (int)(atof((const char *)sqlite3_column_text(st, 0)) * 10 + 0.5);
        sqlite3_finalize(st);
    }
}

static vbobj db_create(dws *w, const char *path, const char *locale, int options)
{
    (void)locale;
    char *file = native_path(path);
    if (file_exists(file)) {
        free(file);
        dao_raise(3204, "DAO.Workspace", "Database already exists.");
        return 0;
    }
    ddb *d = db_new(w, path);
    int rc = db_connect(d, file, 0, 1);
    free(file);
    if (rc != SQLITE_OK) {
        vb_orelease((vbobj)d);
        dao_raise(3044, "DAO.Workspace", "'%s' isn't a valid path.  Make sure that the path name is spelled correctly and that you are connected to the server on which the file resides.", path);
        return 0;
    }
    d->version = (options & 32) ? 30 : (options & 16) ? 30 : 40;      /* dbVersion30 = 32, dbVersion40 = 64 */
    schema_init(d);
    char sql[200];
    snprintf(sql, sizeof sql, "INSERT OR REPLACE INTO MSysLB_Props VALUES ('Version', '%d.0')", d->version / 10);
    sqlite3_exec(d->db, sql, 0, 0, 0);
    d->tds->loaded = 0;
    coll_add(w->dbs, (vbobj)d);
    vb_orelease((vbobj)d);
    return (vbobj)d;
}

/* the SQLite copy of a Jet file: next to it, or in the user's cache when that folder is read-only */
static char *jet_copy_path(const char *file)
{
    size_t n = strlen(file);
    char *p = malloc(n + 8);
    sprintf(p, "%s.lbdb", file);
    return p;
}

/* the SQLite copy of a Jet file, made again when the .mdb is newer (malloc), 0 if it cannot be read */
static char *jet_sqlite_copy(const char *file)
{
    char *copy = jet_copy_path(file);
    struct stat a, b;
    int fresh = stat(copy, &b) == 0 && stat(file, &a) == 0 && b.st_mtime >= a.st_mtime;
    if (!fresh) {
        char err[512];
        if (mdb_import(file, copy, err, sizeof err)) { free(copy); return 0; }
    }
    return copy;
}

#include "daoodbc.inc"

/* ---------------------------------------------------------------- other databases (IN 'file', links) */
/* the kind of a Connect: "" (Jet's own, Microsoft Access), "ODBC", or an installable ISAM (dBASE ...) */
static void connect_kind(const char *c, char *kind, int cap)
{
    const char *e = c ? strchr(c, ';') : 0;
    int n = c ? (e ? (int)(e - c) : (int)strlen(c)) : 0;
    while (n > 0 && c[0] == ' ') { c++; n--; }
    while (n > 0 && c[n - 1] == ' ') n--;
    if (n > 0 && strchr(c, '=') && (!e || strchr(c, '=') < e)) n = 0;     /* "DATABASE=..." first: Jet's own */
    snprintf(kind, (size_t)cap, "%.*s", n, c ? c : "");
    if (!strcasecmp(kind, "MS Access") || !strcasecmp(kind, "Jet")) kind[0] = 0;
}
static char *connect_db(const char *c)
{
    const char *p = c;
    while (p && *p) {
        while (*p == ';' || *p == ' ') p++;
        if (!strncasecmp(p, "database", 8)) {
            const char *q = p + 8;
            while (*q == ' ') q++;
            if (*q == '=') {
                q++;
                const char *e = strchr(q, ';');
                int n = e ? (int)(e - q) : (int)strlen(q);
                char *r = malloc((size_t)n + 1);
                memcpy(r, q, (size_t)n); r[n] = 0;
                return r;
            }
        }
        p = strchr(p, ';');
    }
    return strdup("");
}
static int full_path(const char *p, char out[4096])
{
#ifdef _WIN32
    return _fullpath(out, p, 4096) != 0;
#else
    return realpath(p, out) != 0;
#endif
}
static int same_file(const char *a, const char *b)
{
#ifdef _WIN32
    return !strcasecmp(a, b);
#else
    return !strcmp(a, b);
#endif
}
int dao_ext_schema(ddb *d, const char *path, const char *connect, int *err, char *msg, int cap)
{
    char kind[64];
    connect_kind(connect, kind, sizeof kind);
    if (!strcasecmp(kind, "ODBC")) {
        /* a data source of ODBC: the file of the copies of its tables (daoodbc.inc) */
        osrc *o = osrc_open(connect, "DAO.Database");
        if (!o) {
            *err = vb_err_number_pending();
            snprintf(msg, (size_t)cap, "%s", vb_stoutf8(vb_err_description()));
            vb_err_clear_pending();
            if (!*err) *err = 3146;
            return -1;
        }
        for (int i = 0; i < d->next; i++) if (d->ext[i].odbc == o) return i + 1;
        if (d->next >= DAO_MAXEXT) { *err = 3048; snprintf(msg, (size_t)cap, "Can't open any more databases."); return -1; }
        int n = d->next + 1;
        char sql[64]; snprintf(sql, sizeof sql, "ATTACH DATABASE ?1 AS \"lbext%d\"", n);
        sqlite3_stmt *st = 0;
        int rc = sqlite3_prepare_v2(d->db, sql, -1, &st, 0);
        if (rc == SQLITE_OK) { sqlite3_bind_text(st, 1, o->file, -1, SQLITE_STATIC); rc = sqlite3_step(st); sqlite3_finalize(st); }
        if (rc != SQLITE_DONE && rc != SQLITE_OK) { *err = 3146; snprintf(msg, (size_t)cap, "ODBC--call failed."); return -1; }
        snprintf(sql, sizeof sql, "PRAGMA \"lbext%d\".journal_mode = MEMORY", n);   /* (see daoodbc.inc) */
        sqlite3_exec(d->db, sql, 0, 0, 0);
        d->ext[d->next].file = strdup(o->file);
        d->ext[d->next++].odbc = o;
        if (!d->remote) { d->remote = odbc_remote; d->remote_ctx = d; }
        return n;
    }
    if (*kind) {
        /* dBASE, FoxPro, Paradox, Btrieve, Excel, Text, Lotus, HTML, Exchange: the installable ISAMs of Jet are not
         * part of LinBASIC Studio (decision of 2026-11-27) */
        *err = 3170; snprintf(msg, (size_t)cap, "Couldn't find installable ISAM.");
        return -1;
    }
    if (!path || !*path) { *err = 3024; snprintf(msg, (size_t)cap, "Could not find file '%s'.", path ? path : ""); return -1; }
    char *file = native_path(path);
    if (!file_exists(file)) {
        free(file);
        *err = 3024; snprintf(msg, (size_t)cap, "Could not find file '%s'.", path);
        return -1;
    }
    /* the database itself */
    char *mine = native_path(d->name);
    char ra[4096], rb[4096];
    if (full_path(file, ra) && full_path(mine, rb) && same_file(ra, rb)) { free(file); free(mine); return 0; }
    free(mine);
    char *use = file;
    if (mdb_is_jet(file)) {
        use = jet_sqlite_copy(file);
        if (!use) { free(file); *err = 3343; snprintf(msg, (size_t)cap, "Unrecognized database format '%s'.", path); return -1; }
        free(file);
    } else {
        char head[16] = { 0 };
        FILE *f = fopen(file, "rb");
        if (f) { if (fread(head, 1, 16, f) < 16) head[0] = 0; fclose(f); }
        if (memcmp(head, "SQLite format 3", 15)) {
            free(file);
            *err = 3343; snprintf(msg, (size_t)cap, "Unrecognized database format '%s'.", path);
            return -1;
        }
    }
    char real[4096];
    if (!full_path(use, real)) snprintf(real, sizeof real, "%s", use);
    for (int i = 0; i < d->next; i++)
        if (same_file(d->ext[i].file, real)) { free(use); return i + 1; }
    if (d->next >= DAO_MAXEXT) {
        free(use);
        *err = 3048; snprintf(msg, (size_t)cap, "Can't open any more databases.");
        return -1;
    }
    int n = d->next + 1;
    char sql[64]; snprintf(sql, sizeof sql, "ATTACH DATABASE ?1 AS \"lbext%d\"", n);
    sqlite3_stmt *st = 0;
    int rc = sqlite3_prepare_v2(d->db, sql, -1, &st, 0);
    if (rc == SQLITE_OK) {
        sqlite3_bind_text(st, 1, real, -1, SQLITE_STATIC);
        rc = sqlite3_step(st);
        sqlite3_finalize(st);
    }
    free(use);
    if (rc != SQLITE_DONE && rc != SQLITE_OK) {
        *err = 3051;
        snprintf(msg, (size_t)cap, "The Microsoft Jet database engine cannot open the file '%s'.  It is already opened exclusively by another user, or you need permission to view its data.", path);
        return -1;
    }
    d->ext[d->next++].file = strdup(real);
    return n;
}
int dao_ext_link(ddb *d, const char *connect, int *err, char *msg, int cap)
{
    char kind[64];
    connect_kind(connect, kind, sizeof kind);
    if (*kind) return dao_ext_schema(d, "", connect, err, msg, cap);
    char *path = connect_db(connect);
    int n = dao_ext_schema(d, path, connect, err, msg, cap);
    free(path);
    return n;
}
int dao_ext_table(ddb *d, int n, const char *table)
{
    osrc *o = ext_osrc(d, n);
    return o && n > 0 ? osrc_table(o, table) != 0 : 0;
}
const char *dao_ext_connect(ddb *d, int n)
{
    osrc *o = n > 0 ? ext_osrc(d, n) : 0;
    return o ? o->connect : 0;
}
/* the callbacks of jet_translate */
static int ext_cb(void *ctx, const char *path, const char *connect, int *err, char *msg, int cap)
{
    return dao_ext_schema((ddb *)ctx, path, connect, err, msg, cap);
}
static int link_cb(void *ctx, const char *name, char *src, int cap, int *err, char *msg, int mcap)
{
    ddb *d = ctx;
    char con[1024];
    if (!schema_link(d, name, 0, 0, src, cap, con, sizeof con)) return 0;
    int n = dao_ext_link(d, con, err, msg, mcap);
    if (n == 0) return 0;                       /* (a link to a table of the database itself) */
    return n;
}
void dao_translate_hooks(ddb *d, int on)
{
    jet_ext_fn = on ? ext_cb : 0;
    jet_link_fn = on ? link_cb : 0;
    jet_ext_ctx = on ? d : 0;
}

static vbobj db_open(dws *w, const char *path, int exclusive, int ro, const char *connect)
{
    (void)exclusive;
    char kind[64];
    connect_kind(connect, kind, sizeof kind);
    if (!strcasecmp(kind, "ODBC")) return odbc_db_open(w, connect, ro);   /* a data source of ODBC (daoodbc.inc) */
    if (*kind) {
        /* dBASE, FoxPro, Paradox, Btrieve, Excel, Text ...: the installable ISAMs of Jet are not part of LinBASIC
         * Studio (decision of 2026-11-27) */
        dao_raise(3170, "DAO.Workspace", "Couldn't find installable ISAM.");
        return 0;
    }
    char *file = native_path(path);
    if (!file_exists(file)) {
        free(file);
        dao_raise(3024, "DAO.Workspace", "Could not find file '%s'.", path);
        return 0;
    }
    ddb *d = db_new(w, path);
    free(d->connect); d->connect = strdup(connect ? connect : "");
    char head[16] = { 0 };
    FILE *f = fopen(file, "rb");
    if (f) { if (fread(head, 1, 16, f) < 16) head[0] = 0; fclose(f); }
    char *use = file;
    int jet = mdb_is_jet(file);
    if (jet) {
        /* a Jet file: read through its SQLite copy (made again when the .mdb is newer) */
        char *copy = jet_sqlite_copy(file);
        if (!copy) {
            free(file);
            vb_orelease((vbobj)d);
            dao_raise(3343, "DAO.Workspace", "Unrecognized database format '%s'.", path);
            return 0;
        }
        use = copy;
        d->version = jet == 3 ? 30 : 40;
    } else if (memcmp(head, "SQLite format 3", 15)) {
        free(file);
        vb_orelease((vbobj)d);
        dao_raise(3343, "DAO.Workspace", "Unrecognized database format '%s'.", path);
        return 0;
    }
    int rc = db_connect(d, use, ro, 0);
    if (use != file) free(use);
    free(file);
    if (rc != SQLITE_OK) {
        vb_orelease((vbobj)d);
        dao_raise(3051, "DAO.Workspace", "The Microsoft Jet database engine cannot open the file '%s'.  It is already opened exclusively by another user, or you need permission to view its data.", path);
        return 0;
    }
    if (!ro) schema_init(d);
    if (!jet) read_version(d);
    coll_add(w->dbs, (vbobj)d);
    vb_orelease((vbobj)d);
    return (vbobj)d;
}

static void db_close(ddb *d)
{
    if (d->closed) return;
    while (d->rss->n) rs_close((drs *)vb_oid(d->rss->items[0]));
    if (d->trans) sqlite3_exec(d->db, "ROLLBACK", 0, 0, 0);
    sqlite3_close_v2(d->db);
    d->db = 0;
    d->closed = 1;
    dws *w = (dws *)d->ws;
    if (w) for (int i = 0; i < w->dbs->n; i++) if (vb_oid(w->dbs->items[i]) == (vbobj)d) { vb_oaddref((vbobj)d); coll_remove(w->dbs, i); vb_orelease((vbobj)d); break; }
}

static void db_clear(vbobj o)
{
    ddb *d = (ddb *)o;
    if (!d->closed && d->db) { sqlite3_close_v2(d->db); d->db = 0; }
    free(d->name); free(d->file); free(d->connect);
    for (int i = 0; i < d->next; i++) free(d->ext[i].file);
    d->next = 0;
    if (d->tds) vb_orelease((vbobj)d->tds);
    if (d->qds) vb_orelease((vbobj)d->qds);
    if (d->rels) vb_orelease((vbobj)d->rels);
    if (d->rss) vb_orelease((vbobj)d->rss);
}

static int is_select(const char *s)
{
    while (*s && (isspace((unsigned char)*s) || *s == '(')) s++;
    return !strncasecmp(s, "select", 6) || !strncasecmp(s, "parameters", 10);
}

/* an INSERT that names the counter of the table in its columns (VB6 oracle d_ado1) */
static int insert_names_counter(ddb *d, const char *table, const char *sql)
{
    sqlite3_stmt *st = 0;
    int r = 0;
    if (sqlite3_prepare_v2(d->db, "SELECT name FROM MSysLB_Fields WHERE tbl = ?1 AND (attrs & 16) <> 0", -1, &st, 0) != SQLITE_OK) return 0;
    sqlite3_bind_text(st, 1, table, -1, SQLITE_STATIC);
    if (sqlite3_step(st) == SQLITE_ROW) {
        char q[300];
        snprintf(q, sizeof q, "\"%s\"", (const char *)sqlite3_column_text(st, 0));
        const char *open = strchr(sql, '('), *vals = strstr(sql, ") VALUES");
        if (!vals) vals = strstr(sql, ") SELECT");
        const char *hit = open ? strcasestr(open, q) : 0;
        r = hit && vals && hit < vals;
    }
    sqlite3_finalize(st);
    return r;
}

/* Execute: an action query (or DDL) */
static int db_execute(ddb *d, const char *sql, int options, vbobj qd, long *affected)
{
    *affected = 0;
    {   /* pass-through: the SQL to the data source of ODBC as it is (RecordsAffected 0, as Jet gives it) */
        dqd *pq = qd ? (dqd *)vb_oid(qd) : 0;
        osrc *po = 0;
        if (pq && pq->connect && !strncasecmp(pq->connect, "ODBC", 4) && !(po = qd_osrc(qd))) return -1;
        if (!po && (options & dbSQLPassThrough)) po = (osrc *)d->odbc;
        if (po) return odbc_pt_exec(po, sql, "DAO.QueryDef");
    }
    char real[256];
    const char *text = sql;
    char *saved = 0;
    if (!is_select(sql) && schema_has_table(d, sql, real, sizeof real) == 2) {
        /* the name of a saved action query */
        sqlite3_stmt *st = 0;
        if (sqlite3_prepare_v2(d->db, "SELECT sql FROM MSysLB_Queries WHERE name = ?1", -1, &st, 0) == SQLITE_OK) {
            sqlite3_bind_text(st, 1, real, -1, SQLITE_STATIC);
            if (sqlite3_step(st) == SQLITE_ROW) saved = strdup((const char *)sqlite3_column_text(st, 0));
            sqlite3_finalize(st);
        }
        if (saved) text = saved;
    }
    const char *params[64]; int np = 0;
    static char pbuf[64][256];
    if (qd) {
        dqd *q = (dqd *)vb_oid(qd);
        for (int k = 0; k < q->params->n && np < 63; k++) {
            const char *pn = ((dpar *)q->params->items[k])->name;
            size_t pl = strlen(pn);
            (void)pl; dao_param_plain(pn, pbuf[np], 256); pn = pbuf[np];
            params[np++] = pn;
        }
    }
    params[np] = 0;
    jsql j;
    if (dao_translate(d, text, &j, np ? params : 0) < 0) {
        dao_raise(j.err, "DAO.Database", "%s", j.errmsg);
        jsql_free(&j); free(saved);
        return -1;
    }
    free(saved);
    if (odbc_before(d, &j)) { jsql_free(&j); return -1; }   /* the tables of ODBC it names: their rows again */
    char extq[32] = "";
    if (j.table_ext > 0 && !j.table_src) snprintf(extq, sizeof extq, "lbext%d", j.table_ext);
    int r = 0;
    if (j.kind >= J_CREATE_TABLE && j.kind <= J_DROP_CONSTRAINT) {
        if (d->ro) { dao_raise(3027, "DAO.Database", "Cannot update.  Database or object is read-only."); jsql_free(&j); return -1; }
        r = schema_ddl(d, &j);
        jsql_free(&j);
        return r ? -1 : 0;
    }
    if (j.kind == J_SELECT && j.returns) {
        jsql_free(&j);
        dao_raise(3065, "DAO.Database", "Cannot execute a select query.");
        return -1;
    }
    if (d->ro) { dao_raise(3027, "DAO.Database", "Cannot update.  Database or object is read-only."); jsql_free(&j); return -1; }
    /* the table of an INSERT / UPDATE / DELETE must exist */
    int has_table = 1;
    if (j.table) {
        schema_qualify(extq);                   /* a table of another database: there */
        has_table = schema_has_table(d, j.table, real, sizeof real) == 1;
        schema_qualify(0);
    }
    if (!has_table) {
        if (j.kind == J_INSERT && !d->odbc) dao_raise(3192, "DAO.Database", "Could not find output table '%s'.", j.table);
        else dao_raise(3078, "DAO.Database", "The Microsoft Jet database engine cannot find the input table or query '%s'.  Make sure it exists and that its name is spelled correctly.", j.table);
        jsql_free(&j);
        return -1;
    }
    /* the parameter values */
    vbvar *vals = calloc((size_t)(j.nbinds ? j.nbinds : 1), sizeof(vbvar));
    int missing = 0;
    for (int i = 0; i < j.nbinds; i++) {
        int ok = 0;
        if (qd) {
            dqd *q = (dqd *)vb_oid(qd);
            int k = coll_find(q->params, j.binds[i]);
            if (k >= 0 && ((dpar *)q->params->items[k])->value.vt != VT_EMPTY) { vals[i] = ((dpar *)q->params->items[k])->value; ok = 1; }
        }
        if (!ok) missing++;
    }
    if (missing) {
        dao_raise(3061, "DAO.Database", "Too few parameters. Expected %d.", missing);
        free(vals); jsql_free(&j);
        return -1;
    }
    jbuf s = { 0 };
    if (dao_odbc_into(d, &j)) {                 /* SELECT ... INTO a table of a data source of ODBC */
        long n = dao_odbc_select_into(d, &j, vals);
        free(vals); jsql_free(&j);
        if (n < 0) return -1;
        *affected = n;
        return 0;
    }
    if (j.into) {
        /* SELECT ... INTO t: the new table has the types of the columns */
        char intoq[32] = "";
        if (j.into_ext > 0) snprintf(intoq, sizeof intoq, "lbext%d", j.into_ext);
        schema_qualify(intoq);
        int exists = schema_has_table(d, j.into, real, sizeof real);
        schema_qualify(0);
        if (exists) {
            dao_raise(3010, "DAO.Database", "Table '%s' already exists.", j.into);
            free(vals); jsql_free(&j);
            return -1;
        }
        sqlite3_stmt *st = 0;
        int rc = sqlite3_prepare_v2(d->db, j.sql, -1, &st, 0);
        if (rc != SQLITE_OK) { dao_raise_sqlite(d, rc, "DAO.Database"); free(vals); jsql_free(&j); return -1; }
        dtd *t = tabledef_new(j.into);
        for (int i = 0; i < sqlite3_column_count(st); i++) {
            int size, attrs;
            int type = jet_decl_type(sqlite3_column_decltype(st, i), &size, &attrs);
            rcol c; memset(&c, 0, sizeof c);
            const char *tn = sqlite3_column_table_name(st, i), *on = sqlite3_column_origin_name(st, i);
            const char *dn = sqlite3_column_database_name(st, i);
            schema_qualify(dn && !strncmp(dn, "lbext", 5) ? dn : 0);
            if (tn && on && schema_field_info(d, tn, on, &c)) { type = c.type; size = c.size; }
            schema_qualify(0);
            free(c.defval);
            if (!type) type = dbText, size = 255;
            dfld *f = field_new(FK_TABLE, (vbobj)t, sqlite3_column_name(st, i), type, size);
            if (c.autoinc) {}                   /* a counter becomes a Long */
            f->allowzero = 2;
            coll_add(t->fields, (vbobj)f); vb_orelease((vbobj)f);
        }
        sqlite3_finalize(st);
        if (j.into_ext > 0) {
            /* the table in the other database: made there (its own connection), filled through this one */
            ddb *x = db_new((dws *)d->ws, d->ext[j.into_ext - 1].file);
            r = db_connect(x, d->ext[j.into_ext - 1].file, 0, 0) != SQLITE_OK;
            if (r) dao_raise(3051, "DAO.Database", "The Microsoft Jet database engine cannot open the file '%s'.  It is already opened exclusively by another user, or you need permission to view its data.", d->ext[j.into_ext - 1].file);
            else { schema_init(x); r = schema_create_table(x, t); }
            if (x->db) sqlite3_close_v2(x->db);
            x->db = 0; x->closed = 1;
            vb_orelease((vbobj)x);
        } else r = schema_create_table(d, t);
        vb_orelease((vbobj)t);
        if (r) { free(vals); jsql_free(&j); return -1; }
        jb_s(&s, "INSERT INTO ");
        if (j.into_ext > 0) jb_f(&s, "\"lbext%d\".", j.into_ext);
        jb_qid(&s, j.into, -1); jb_s(&s, " "); jb_s(&s, j.sql);
    } else jb_s(&s, j.sql);
    ocap cap;
    ocap_begin(d, &j, &cap);                    /* a table of ODBC: its changes logged, to go to the server */
    sqlite3_stmt *st = 0;
    int rc = sqlite3_prepare_v2(d->db, s.p, -1, &st, 0);
    jb_free(&s);
    if (rc != SQLITE_OK) { ocap_end(d, &cap, 0, ""); dao_raise_sqlite(d, rc, "DAO.Database"); free(vals); jsql_free(&j); return -1; }
    for (int i = 0; i < j.nbinds; i++) jet_bind(st, i + 1, vals[i], 0);
    free(vals);
    rc = sqlite3_step(st);
    if (rc == SQLITE_DONE) *affected = sqlite3_changes(d->db);
    char emsg[1024]; int enr = 0;
    if (rc != SQLITE_DONE) enr = jet_sqlite_error(d->db, rc, emsg, sizeof emsg);
    sqlite3_finalize(st);
    if (cap.on && ocap_end(d, &cap, rc == SQLITE_DONE, d->odbc && !j.table_src ? "???" : j.table)) { *affected = 0; jsql_free(&j); return -1; }
    if (rc != SQLITE_DONE) {
        if (j.kind == J_INSERT && !j.table_ext && !insert_names_counter(d, real, j.sql))
            schema_counter_bump(d, real);       /* Jet used up the counter value (not when the INSERT gave it) */
        int fail = (options & dbFailOnError) || !((rc & 0xff) == SQLITE_CONSTRAINT);
        if (fail) { dao_raise(enr, "DAO.Database", "%s", emsg); jsql_free(&j); return -1; }
    }
    jsql_free(&j);
    return 0;
}

/* ---------------------------------------------------------------- for ADO (ado.c) */
/* a database file opened as DAO opens it: the Database (borrowed: Workspaces(0).Databases keeps it) or 0 */
vbobj dao_open_path(const char *path, int ro)
{
    return db_open((dws *)vb_oid(dao_workspace0()), path, 0, ro, "");
}
/* a database in memory (a recordset of ADO without a connection): an owned reference */
vbobj dao_memory_db(void)
{
    ddb *d = db_new((dws *)vb_oid(dao_workspace0()), "");
    if (db_connect(d, ":memory:", 0, 1) != SQLITE_OK) { vb_orelease((vbobj)d); return 0; }
    d->ws = 0;
    schema_init(d);
    return (vbobj)d;
}
int dao_execute_qd(ddb *d, const char *sql, vbobj qd, long *affected)
{
    return db_execute(d, sql, dbFailOnError, qd, affected);
}
void dao_db_close(vbobj o) { db_close((ddb *)vb_oid(o)); }

static vbvar db_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    ddb *d = (ddb *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    dao_src = "DAO.Database";
    if (d->closed && strcmp(name, "close")) { dao_raise(3420, "DAO.Database", "Object invalid or no longer set."); return vb_v_empty(); }
    if (!strcmp(name, "tabledefs") || !strcmp(name, "")) {
        if (argc == 0) return dao_retobj((vbobj)d->tds);
        int f2;
        return coll_member((vbobj)d->tds, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "querydefs")) {
        if (argc == 0) return dao_retobj((vbobj)d->qds);
        int f2; return coll_member((vbobj)d->qds, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "relations")) {
        if (argc == 0) return dao_retobj((vbobj)d->rels);
        int f2; return coll_member((vbobj)d->rels, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "recordsets")) {
        if (argc == 0) return dao_retobj((vbobj)d->rss);
        int f2; return coll_member((vbobj)d->rss, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "properties")) return props_get(o, kind, argc, argv);
    if (!strcmp(name, "containers")) {
        dcoll *c = coll_new(C_CONTAINERS, 0);
        c->loaded = 1;
        vbvar v = dao_retobj((vbobj)c);
        vb_orelease((vbobj)c);
        return v;
    }
    if (!strcmp(name, "name")) return dao_vstr(d->name);
    if (!strcmp(name, "connect")) { if (set) { free(d->connect); d->connect = dao_strdup_v(argv[argc - 1]); return vb_v_empty(); } return dao_vstr(d->connect); }
    if (!strcmp(name, "updatable")) return vb_v_bool(d->ro ? 0 : -1);
    if (!strcmp(name, "version")) { char v[8]; snprintf(v, sizeof v, "%d.0", d->version / 10); return dao_vstr(v); }
    if (!strcmp(name, "recordsaffected")) return vb_v_l((vb_i32)d->affected);
    if (!strcmp(name, "transactions")) return vb_v_bool(-1);
    if (!strcmp(name, "collatingorder")) return vb_v_l(1033);
    if (!strcmp(name, "querytimeout")) return vb_v_i(60);
    if (!strcmp(name, "replicaid") || !strcmp(name, "designmasterid")) return dao_vstr("");
    if (!strcmp(name, "close")) { db_close(d); return vb_v_empty(); }
    if (!strcmp(name, "execute")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        long n;
        if (db_execute(d, dao_str(argv[0]), along(argc, argv, 1, 0), 0, &n) == 0) d->affected = n;
        else d->affected = 0;
        return vb_v_empty();
    }
    if (!strcmp(name, "openrecordset")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        vbobj r = rs_open(d, dao_str(argv[0]), along(argc, argv, 1, 0), along(argc, argv, 2, 0), along(argc, argv, 3, 0), 0, 0, 0);
        return r ? dao_retobj(r) : vb_v_empty();
    }
    if (!strcmp(name, "createtabledef")) {
        dtd *t = tabledef_new(dao_missing(argc, argv, 0) ? "" : dao_str(argv[0]));
        t->db = o; vb_oaddref(o);
        t->attrs = along(argc, argv, 1, 0);
        if (!dao_missing(argc, argv, 2)) { free(t->srcname); t->srcname = dao_strdup_v(argv[2]); }
        if (!dao_missing(argc, argv, 3)) { free(t->connect); t->connect = dao_strdup_v(argv[3]); }
        vbvar v = dao_retobj((vbobj)t);
        vb_orelease((vbobj)t);
        return v;
    }
    if (!strcmp(name, "createquerydef")) {
        const char *nm = dao_missing(argc, argv, 0) ? "" : dao_str(argv[0]);
        const char *sql = dao_missing(argc, argv, 1) ? "" : dao_str(argv[1]);
        int temp = !*nm;
        if (!temp && d->odbc) {                 /* a database of ODBC keeps no queries */
            dao_raise(3251, "DAO.QueryDefs", "Operation is not supported for this type of object.");
            return vb_v_empty();
        }
        if (!temp) {
            coll_load(d->qds);
            if (coll_find(d->qds, nm) >= 0 || schema_has_table(d, nm, 0, 0)) {
                dao_raise(3012, "DAO.QueryDefs", "Object '%s' already exists.", nm);
                return vb_v_empty();
            }
        }
        char *pretty = *sql ? jet_pretty_sql(sql) : strdup("");
        dqd *q = querydef_new(o, temp ? "#Temporary QueryDef#" : nm, pretty);
        free(pretty);
        q->temp = temp;
        if (*sql) {
            jsql j;
            if (jet_translate(sql, &j, 0) < 0 && j.err != 3061) {
                dao_raise(j.err, "DAO.Database", "%s", j.errmsg);
                jsql_free(&j);
                vb_orelease((vbobj)q);
                return vb_v_empty();
            }
            jsql_free(&j);
        }
        if (!temp) {                            /* (a QueryDef with a name is appended at once, its SQL or not) */
            if (schema_save_query(d, q, 0)) { vb_orelease((vbobj)q); return vb_v_empty(); }
            q->appended = 1;
            coll_load(d->qds);
            coll_add(d->qds, (vbobj)q);
        }
        vbvar v = dao_retobj((vbobj)q);
        vb_orelease((vbobj)q);
        return v;
    }
    if (!strcmp(name, "createrelation")) {
        drel *r = relation_new(dao_missing(argc, argv, 0) ? "" : dao_str(argv[0]), dao_missing(argc, argv, 1) ? "" : dao_str(argv[1]),
                               dao_missing(argc, argv, 2) ? "" : dao_str(argv[2]), along(argc, argv, 3, 0));
        r->db = o; vb_oaddref(o);
        vbvar v = dao_retobj((vbobj)r);
        vb_orelease((vbobj)r);
        return v;
    }
    if (!strcmp(name, "createproperty")) return prop_create(o, argc, argv);
    if (!strcmp(name, "newpassword") || !strcmp(name, "makereplica") || !strcmp(name, "synchronize") || !strcmp(name, "populatepartial")) return vb_v_empty();
    *found = 0;
    return vb_v_empty();
}

static int db_append_tabledef(ddb *d, dtd *t)
{
    if (dao_check_db(d, "DAO.TableDefs")) return -1;
    if (t->appended) { dao_raise(3367, "DAO.TableDefs", "Cannot append. An object with that name already exists in the collection."); return -1; }
    if (!t->db) { t->db = (vbobj)d; vb_oaddref((vbobj)d); }
    if ((t->connect && *t->connect) || (t->srcname && *t->srcname)) return schema_create_link(d, t) ? -1 : 0;
    if (d->odbc) return dao_odbc_create_table(d, t) ? -1 : 0;     /* a database of ODBC: on the server */
    return schema_create_table(d, t) ? -1 : 0;
}
static int db_append_querydef(ddb *d, dqd *q)
{
    if (dao_check_db(d, "DAO.QueryDefs")) return -1;
    if (q->appended) { dao_raise(3367, "DAO.QueryDefs", "Cannot append. An object with that name already exists in the collection."); return -1; }
    if (!*q->name || q->temp) { dao_raise(3219, "DAO.QueryDefs", "Invalid operation."); return -1; }
    if (schema_save_query(d, q, 0)) return -1;
    q->appended = 1;
    return 0;
}
static int db_append_relation(ddb *d, drel *r)
{
    if (dao_check_db(d, "DAO.Relations")) return -1;
    if (schema_create_relation(d, r)) return -1;
    if (!r->db) { r->db = (vbobj)d; vb_oaddref((vbobj)d); }
    return 0;
}

/* ---------------------------------------------------------------- TableDef */
static void tabledef_clear(vbobj o)
{
    dtd *t = (dtd *)o;
    free(t->name); free(t->connect); free(t->srcname); free(t->vrule); free(t->vtext);
    if (t->fields) vb_orelease((vbobj)t->fields);
    if (t->indexes) vb_orelease((vbobj)t->indexes);
    if (t->db) vb_orelease(t->db);
}

static int tabledef_append_field(dtd *t, dfld *f)
{
    if (f->fk != FK_TABLE && f->fk != 0) { dao_raise(3219, "DAO.Fields", "Invalid operation."); return -1; }
    if (!f->type) { dao_raise(3259, "DAO.Fields", "Invalid field data type."); return -1; }
    if (!*f->name) { dao_raise(3264, "DAO.Fields", "No field defined--cannot append TableDef or Index."); return -1; }
    if (t->appended) {
        if (DB(t->db)->odbc ? dao_odbc_add_field(DB(t->db), t->name, f) : schema_add_field(DB(t->db), t->name, f)) return -1;
    }
    if (f->owner != (vbobj)t) { if (f->owner) vb_orelease(f->owner); f->owner = (vbobj)t; vb_oaddref((vbobj)t); }
    f->ord = t->fields->n;
    return 0;
}
static int tabledef_append_index(dtd *t, dix *ix)
{
    if (t->appended) {
        if (DB(t->db)->odbc ? dao_odbc_create_index(DB(t->db), t->name, ix) : schema_create_index(DB(t->db), t->name, ix)) return -1;
    }
    if (!ix->td) { ix->td = (vbobj)t; vb_oaddref((vbobj)t); }
    if (ix->primary) { ix->unique = 1; ix->required = 1; }
    return 0;
}

static vbvar tabledef_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dtd *t = (dtd *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    dao_src = "DAO.TableDef";
    if (t->db && DB(t->db)->closed) { dao_raise(3420, "DAO.TableDef", "Object invalid or no longer set."); return vb_v_empty(); }
    if (!strcmp(name, "fields") || !strcmp(name, "")) {
        t->fields->loaded = 1;
        if (argc == 0) return dao_retobj((vbobj)t->fields);
        int f2; return coll_member((vbobj)t->fields, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "indexes")) {
        t->indexes->loaded = 1;
        if (argc == 0) return dao_retobj((vbobj)t->indexes);
        int f2; return coll_member((vbobj)t->indexes, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "name")) {
        if (set) {
            char *nn = dao_str(argv[argc - 1]);
            if (t->appended) {
                ddb *d = DB(t->db);
                jbuf s = { 0 };
                jb_s(&s, "ALTER TABLE "); jb_qid(&s, t->name, -1); jb_s(&s, " RENAME TO "); jb_qid(&s, nn, -1); jb_s(&s, ";");
                jb_s(&s, "UPDATE MSysLB_Tables SET name = "); jb_qstr(&s, nn, -1); jb_s(&s, " WHERE name = "); jb_qstr(&s, t->name, -1); jb_s(&s, ";");
                jb_s(&s, "UPDATE MSysLB_Fields SET tbl = "); jb_qstr(&s, nn, -1); jb_s(&s, " WHERE tbl = "); jb_qstr(&s, t->name, -1); jb_s(&s, ";");
                jb_s(&s, "UPDATE MSysLB_Indexes SET tbl = "); jb_qstr(&s, nn, -1); jb_s(&s, " WHERE tbl = "); jb_qstr(&s, t->name, -1); jb_s(&s, ";");
                int r = dao_exec(d, s.p);
                jb_free(&s);
                if (r) return vb_v_empty();
            }
            free(t->name); t->name = strdup(nn);
            return vb_v_empty();
        }
        return dao_vstr(t->name);
    }
    if (!strcmp(name, "attributes")) { if (set) { t->attrs = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(t->attrs); }
    if (!strcmp(name, "updatable")) return vb_v_bool(t->attrs & (dbSystemObject | dbAttachedTable | dbAttachedODBC) ? 0 : -1);
    if (!strcmp(name, "recordcount")) {
        if (!t->appended || (t->attrs & (dbSystemObject | dbAttachedTable | dbAttachedODBC))) return vb_v_l(t->appended ? 0 : -1);
        ddb *d = DB(t->db);
        jbuf s = { 0 };
        jb_s(&s, "SELECT count(*) FROM "); jb_qid(&s, t->name, -1);
        sqlite3_stmt *st = 0;
        vb_i32 n = 0;
        if (sqlite3_prepare_v2(d->db, s.p, -1, &st, 0) == SQLITE_OK && sqlite3_step(st) == SQLITE_ROW) n = sqlite3_column_int(st, 0);
        sqlite3_finalize(st);
        jb_free(&s);
        return vb_v_l(n);
    }
    if (!strcmp(name, "connect") || !strcmp(name, "sourcetablename") || !strcmp(name, "validationrule") || !strcmp(name, "validationtext")) {
        char **p = name[0] == 'c' ? &t->connect : name[0] == 's' ? &t->srcname : name[10] == 'r' ? &t->vrule : &t->vtext;
        if (set) { free(*p); *p = dao_strdup_v(argv[argc - 1]); return vb_v_empty(); }
        return dao_vstr(*p);
    }
    if (!strcmp(name, "datecreated")) return vb_v_date(t->created);
    if (!strcmp(name, "lastupdated")) return vb_v_date(t->updated);
    if (!strcmp(name, "conflicttable")) return dao_vstr("");
    if (!strcmp(name, "createfield")) {
        int type = along(argc, argv, 1, 0);
        dfld *f = field_new(FK_TABLE, 0, dao_missing(argc, argv, 0) ? "" : dao_str(argv[0]), type, along(argc, argv, 2, 0));
        vbvar v = dao_retobj((vbobj)f);
        vb_orelease((vbobj)f);
        return v;
    }
    if (!strcmp(name, "createindex")) {
        dix *ix = index_new(0, dao_missing(argc, argv, 0) ? "" : dao_str(argv[0]));
        vbvar v = dao_retobj((vbobj)ix);
        vb_orelease((vbobj)ix);
        return v;
    }
    if (!strcmp(name, "openrecordset")) {
        if (!t->appended) { dao_raise(3219, "DAO.TableDef", "Invalid operation."); return vb_v_empty(); }
        vbobj r = rs_open(DB(t->db), t->name, along(argc, argv, 0, 0), along(argc, argv, 1, 0), 0, 0, 0, 0);
        return r ? dao_retobj(r) : vb_v_empty();
    }
    if (!strcmp(name, "refreshlink")) {
        /* the database it links to again (another file perhaps: its Connect) */
        if (!(t->attrs & (dbAttachedTable | dbAttachedODBC)) || !t->db) return vb_v_empty();
        ddb *d = DB(t->db);
        int err = 0; char msg[400];
        if (dao_ext_link(d, t->connect, &err, msg, sizeof msg) < 0) { dao_raise(err, "DAO.TableDef", "%s", msg); return vb_v_empty(); }
        sqlite3_stmt *st = 0;
        if (sqlite3_prepare_v2(d->db, "UPDATE MSysLB_Tables SET connect = ?2, srcname = ?3 WHERE name = ?1", -1, &st, 0) == SQLITE_OK) {
            sqlite3_bind_text(st, 1, t->name, -1, SQLITE_STATIC);
            sqlite3_bind_text(st, 2, t->connect, -1, SQLITE_STATIC);
            sqlite3_bind_text(st, 3, t->srcname, -1, SQLITE_STATIC);
            sqlite3_step(st);
            sqlite3_finalize(st);
        }
        return vb_v_empty();
    }
    if (!strcmp(name, "createproperty")) return prop_create(o, argc, argv);
    if (!strcmp(name, "properties")) return props_get(o, kind, argc, argv);
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Field */
static void field_clear(vbobj o)
{
    dfld *f = (dfld *)o;
    free(f->name); free(f->defval); free(f->vrule); free(f->vtext); free(f->foreign); free(f->stable); free(f->sfield);
    if (f->owner && f->fk != FK_RS) vb_orelease(f->owner);
}

static vbvar field_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dfld *f = (dfld *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    dao_src = "DAO.Field";
    if (f->fk == FK_RS) {
        drs *r = (drs *)f->owner;
        rcol *c = f->col >= 0 && f->col < r->ncols ? &r->cols[f->col] : 0;
        if (!strcmp(name, "value") || !strcmp(name, "")) {
            if (set) { rs_field_let(r, f->col, vb_v_letval(argv[argc - 1])); return vb_v_empty(); }
            int err;
            return rs_field_value(r, f->col, &err);
        }
        if (r->closed) { dao_raise(3420, "DAO.Field", "Object invalid or no longer set."); return vb_v_empty(); }
        if (!c) { *found = 0; return vb_v_empty(); }
        if (!strcmp(name, "name")) return dao_vstr(c->name);
        if (!strcmp(name, "type")) return vb_v_i((vb_i16)c->type);
        if (!strcmp(name, "size")) return vb_v_l(c->size);
        if (!strcmp(name, "attributes")) return vb_v_l(c->attrs | (c->updatable && r->base && r->type != dbOpenSnapshot ? dbUpdatableField : 0));
        if (!strcmp(name, "sourcetable")) return dao_vstr(c->stable ? c->stable : "");
        if (!strcmp(name, "sourcefield")) return dao_vstr(c->sfield ? c->sfield : "");
        if (!strcmp(name, "dataupdatable")) return vb_v_bool(c->updatable && r->base && r->type != dbOpenSnapshot ? -1 : 0);
        if (!strcmp(name, "required")) return vb_v_bool(c->required ? -1 : 0);
        if (!strcmp(name, "allowzerolength")) return vb_v_bool(c->allowzero ? -1 : 0);
        if (!strcmp(name, "ordinalposition")) return vb_v_i((vb_i16)f->col);
        if (!strcmp(name, "defaultvalue")) return dao_vstr(c->defval ? c->defval : "");
        if (!strcmp(name, "validationrule") || !strcmp(name, "validationtext") || !strcmp(name, "foreignname")) return dao_vstr("");
        if (!strcmp(name, "collatingorder")) return vb_v_l(1033);
        if (!strcmp(name, "fieldsize")) {
            int err; vbvar v = rs_field_value(r, f->col, &err);
            if (err) return vb_v_empty();
            if (v.vt == VT_BSTR) return vb_v_l(v.u.s.len);
            if (v.vt & VT_ARRAY) { vbarr *a = vb_v_arrptr(v); return vb_v_l(a ? a->n : 0); }
            return vb_v_l(0);
        }
        if (!strcmp(name, "getchunk")) {
            int err; vbvar v = rs_field_value(r, f->col, &err);
            if (err) return vb_v_empty();
            vb_i32 off = along(argc, argv, 0, 0), n = along(argc, argv, 1, 0);
            if (v.vt == VT_BSTR) return vb_v_s(vb_mid(v.u.s, off + 1, n));
            return v;
        }
        if (!strcmp(name, "appendchunk")) {
            int err; vbvar v = rs_field_value(r, f->col, &err);
            if (err) return vb_v_empty();
            vbvar nv = v.vt == VT_NULL ? argv[0] : vb_v_cat(v, argv[0]);
            rs_field_let(r, f->col, nv);
            return vb_v_empty();
        }
        if (!strcmp(name, "properties")) return props_get(o, kind, argc, argv);
        *found = 0;
        return vb_v_empty();
    }
    /* a field of a TableDef, an Index or a Relation */
    int appended = f->appended && f->fk == FK_TABLE;
    if (!strcmp(name, "name")) {
        if (set) { if (appended) { dao_raise(3219, "DAO.Field", "Invalid operation."); return vb_v_empty(); } free(f->name); f->name = dao_strdup_v(argv[argc - 1]); return vb_v_empty(); }
        return dao_vstr(f->name);
    }
    if (!strcmp(name, "type")) {
        if (set) { if (appended) { dao_raise(3219, "DAO.Field", "Invalid operation."); return vb_v_empty(); }
            dao_field_settype(f, vb_v2l(argv[argc - 1])); return vb_v_empty(); }
        return vb_v_i((vb_i16)f->type);
    }
    if (!strcmp(name, "size")) {
        if (set) { if (appended) { dao_raise(3219, "DAO.Field", "Invalid operation."); return vb_v_empty(); } f->size = vb_v2l(argv[argc - 1]); return vb_v_empty(); }
        return vb_v_l(f->fk == FK_TABLE || f->fk == FK_QD ? f->size : 0);
    }
    if (!strcmp(name, "attributes")) {
        if (set) { f->attrs = vb_v2l(argv[argc - 1]); return vb_v_empty(); }
        return vb_v_l(f->attrs);
    }
    if (!strcmp(name, "required")) { if (set) { f->required = vb_v2bool(argv[argc - 1]) != 0; return vb_v_empty(); } return vb_v_bool(f->required ? -1 : 0); }
    if (!strcmp(name, "allowzerolength")) { if (set) { f->allowzero = vb_v2bool(argv[argc - 1]) != 0; return vb_v_empty(); } return vb_v_bool(f->allowzero == 1 ? -1 : 0); }
    if (!strcmp(name, "ordinalposition")) { if (set) { f->ord = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_i((vb_i16)f->ord); }
    if (!strcmp(name, "defaultvalue") || !strcmp(name, "validationrule") || !strcmp(name, "validationtext") || !strcmp(name, "foreignname")) {
        char **p = name[0] == 'd' ? &f->defval : name[0] == 'f' ? &f->foreign : name[10] == 'r' ? &f->vrule : &f->vtext;
        if (set) { free(*p); *p = dao_strdup_v(argv[argc - 1]); return vb_v_empty(); }
        return dao_vstr(*p);
    }
    if (f->fk == FK_QD && !strcmp(name, "sourcetable")) return dao_vstr(f->stable);
    if (f->fk == FK_QD && !strcmp(name, "sourcefield")) return dao_vstr(f->sfield);
    if (!strcmp(name, "sourcetable")) {
        dtd *ot = f->fk == FK_TABLE && f->owner ? (dtd *)f->owner : 0;
        return dao_vstr(!ot ? "" : (ot->attrs & (dbAttachedTable | dbAttachedODBC)) && ot->srcname && *ot->srcname ? ot->srcname : ot->name);
    }
    if (!strcmp(name, "sourcefield")) return dao_vstr(f->fk == FK_TABLE ? f->name : "");
    if (!strcmp(name, "dataupdatable")) return vb_v_bool(f->fk == FK_QD && f->dataupd ? -1 : 0);
    if (!strcmp(name, "collatingorder")) return vb_v_l(1033);
    if (!strcmp(name, "value")) { dao_raise(3219, "DAO.Field", "Invalid operation."); return vb_v_empty(); }
    if (!strcmp(name, "createproperty")) return prop_create(o, argc, argv);
    if (!strcmp(name, "properties")) return props_get(o, kind, argc, argv);
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Index */
static void index_clear(vbobj o)
{
    dix *ix = (dix *)o;
    free(ix->name);
    if (ix->fields) vb_orelease((vbobj)ix->fields);
    if (ix->td) vb_orelease(ix->td);
}
/* DistinctCount: the different values of the index's fields in its table */
static long index_distinct(dix *ix)
{
    if (!ix->appended || !ix->td) return 0;
    dtd *t = (dtd *)vb_oid(ix->td);
    if (!t->appended || !t->db || !ix->fields || !ix->fields->n) return 0;
    ddb *d = DB(t->db);
    jbuf s = { 0 };
    jb_s(&s, "SELECT COUNT(*) FROM (SELECT DISTINCT ");
    for (int i = 0; i < ix->fields->n; i++) {
        if (i) jb_s(&s, ", ");
        jb_qid(&s, obj_name(ix->fields->items[i]), -1);
    }
    jb_s(&s, " FROM "); jb_qid(&s, t->name, -1); jb_s(&s, ")");
    long n = 0;
    sqlite3_stmt *st = 0;
    if (sqlite3_prepare_v2(d->db, s.p, -1, &st, 0) == SQLITE_OK && sqlite3_step(st) == SQLITE_ROW) n = sqlite3_column_int64(st, 0);
    sqlite3_finalize(st);
    jb_free(&s);
    return n;
}

static vbvar index_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dix *ix = (dix *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    dao_src = "DAO.Index";
    if (!strcmp(name, "fields")) {
        ix->fields->loaded = 1;
        if (set) {                              /* ix.Fields = "+A;-B" */
            if (ix->appended) { dao_raise(3219, "DAO.Index", "Invalid operation."); return vb_v_empty(); }
            coll_clear_items(ix->fields);
            char *s = dao_strdup_v(argv[argc - 1]), *save = 0;
            for (char *p = strtok_r(s, ";", &save); p; p = strtok_r(0, ";", &save)) {
                while (*p == ' ') p++;
                int desc = *p == '-';
                if (*p == '+' || *p == '-') p++;
                dfld *g = field_new(FK_INDEX, (vbobj)ix, p, 0, 0);
                g->attrs = desc ? dbDescending : 0;
                coll_add(ix->fields, (vbobj)g); vb_orelease((vbobj)g);
            }
            free(s);
            return vb_v_empty();
        }
        if (argc == 0) return dao_retobj((vbobj)ix->fields);
        int f2; return coll_member((vbobj)ix->fields, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "name")) { if (set) { free(ix->name); ix->name = dao_strdup_v(argv[argc - 1]); return vb_v_empty(); } return dao_vstr(ix->name); }
    if (!strcmp(name, "primary") && !set && ix->remote) {
        dao_raise(3365, "DAO.Index", "Property value not valid for REMOTE objects.");
        return vb_v_empty();
    }
    int *flag = !strcmp(name, "primary") ? &ix->primary : !strcmp(name, "unique") ? &ix->unique : !strcmp(name, "required") ? &ix->required :
                !strcmp(name, "ignorenulls") ? &ix->ignorenulls : !strcmp(name, "clustered") ? &ix->clustered : 0;
    if (flag) {
        if (set) {
            if (ix->appended) { dao_raise(3219, "DAO.Index", "Invalid operation."); return vb_v_empty(); }
            *flag = vb_v2bool(argv[argc - 1]) != 0;
            if (flag == &ix->primary && *flag) { ix->unique = 1; ix->required = 1; }
            return vb_v_empty();
        }
        return vb_v_bool(*flag ? -1 : 0);
    }
    if (!strcmp(name, "foreign")) return vb_v_bool(ix->foreign ? -1 : 0);
    if (!strcmp(name, "distinctcount")) return vb_v_l(index_distinct(ix));
    if (!strcmp(name, "createfield")) {
        dfld *f = field_new(FK_INDEX, 0, dao_missing(argc, argv, 0) ? "" : dao_str(argv[0]), 0, 0);
        vbvar v = dao_retobj((vbobj)f);
        vb_orelease((vbobj)f);
        return v;
    }
    if (!strcmp(name, "createproperty")) return prop_create(o, argc, argv);
    if (!strcmp(name, "properties")) return props_get(o, kind, argc, argv);
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Relation */
static void relation_clear(vbobj o)
{
    drel *r = (drel *)o;
    free(r->name); free(r->table); free(r->ftable);
    if (r->fields) vb_orelease((vbobj)r->fields);
    if (r->db) vb_orelease(r->db);
}
static vbvar relation_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    drel *r = (drel *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    dao_src = "DAO.Relation";
    if (!strcmp(name, "properties")) return props_get(o, kind, argc, argv);
    if (!strcmp(name, "fields")) {
        r->fields->loaded = 1;
        if (argc == 0) return dao_retobj((vbobj)r->fields);
        int f2; return coll_member((vbobj)r->fields, "item", kind, argc, argv, &f2);
    }
    char **p = !strcmp(name, "name") ? &r->name : !strcmp(name, "table") ? &r->table : !strcmp(name, "foreigntable") ? &r->ftable : 0;
    if (p) {
        if (set) { if (r->appended) { dao_raise(3219, "DAO.Relation", "Invalid operation."); return vb_v_empty(); } free(*p); *p = dao_strdup_v(argv[argc - 1]); return vb_v_empty(); }
        return dao_vstr(*p);
    }
    if (!strcmp(name, "attributes")) { if (set) { r->attrs = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(r->attrs); }
    if (!strcmp(name, "partialreplica")) return vb_v_bool(0);
    if (!strcmp(name, "createfield")) {
        dfld *f = field_new(FK_REL, 0, dao_missing(argc, argv, 0) ? "" : dao_str(argv[0]), 0, 0);
        vbvar v = dao_retobj((vbobj)f);
        vb_orelease((vbobj)f);
        return v;
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- QueryDef, Parameter */
static void querydef_clear(vbobj o)
{
    dqd *q = (dqd *)o;
    free(q->name); free(q->sql); free(q->connect);
    if (q->params) vb_orelease((vbobj)q->params);
    if (q->db) vb_orelease(q->db);
}
/* the fields of a QueryDef: the columns of its result, as Fields of a TableDef read them (Value: Invalid operation) */
static void qd_fields(ddb *d, dqd *q, dcoll *c)
{
    vb_i32 save = vb_errpend;
    vbobj r0 = rs_open(d, q->sql, dbOpenDynaset, 0, 0, 0, 0, 0);
    if (!r0) { if (vb_errpend && !save) vb_err_clear_pending(); return; }
    drs *r = (drs *)vb_oid(r0);
    for (int i = 0; i < r->ncols; i++) {
        rcol *rc = &r->cols[i];
        dfld *f = field_new(FK_QD, 0, rc->name, rc->type, rc->size);
        f->type = rc->type; f->size = rc->size;
        int upd = rc->updatable && r->base;
        f->attrs = rc->attrs | (upd ? dbUpdatableField : 0);
        f->dataupd = upd;
        f->required = rc->required; f->allowzero = rc->allowzero ? 1 : 0;
        f->ord = i;
        free(f->stable); f->stable = strdup(rc->stable ? rc->stable : "");
        free(f->sfield); f->sfield = strdup(rc->sfield ? rc->sfield : "");
        f->appended = 1;
        coll_add(c, (vbobj)f);
        vb_orelease((vbobj)f);
    }
    vb_late_o(r0, "close", VBM_METHOD, 0, 0);    /* (rs_open gives a temporary: the end of the statement frees it) */
    if (vb_errpend && !save) vb_err_clear_pending();
}

static vbvar querydef_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dqd *q = (dqd *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    dao_src = "DAO.QueryDef";
    ddb *d = q->db ? DB(q->db) : 0;
    if (d && d->closed) { dao_raise(3420, "DAO.QueryDef", "Object invalid or no longer set."); return vb_v_empty(); }
    if (!strcmp(name, "parameters") || !strcmp(name, "")) {
        q->params->loaded = 1;
        if (argc == 0) return dao_retobj((vbobj)q->params);
        int f2; return coll_member((vbobj)q->params, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "name")) {
        if (set) { free(q->name); q->name = dao_strdup_v(argv[argc - 1]); return vb_v_empty(); }
        return dao_vstr(q->name);
    }
    if (!strcmp(name, "sql")) {
        if (set) {
            char *s = dao_str(argv[argc - 1]);
            char *pretty = jet_pretty_sql(s);
            free(q->sql); q->sql = pretty;
            querydef_parse_params(q);
            if (q->appended && d) schema_save_query(d, q, 1);
            return vb_v_empty();
        }
        return dao_vstr(q->sql);
    }
    if (!strcmp(name, "type")) return vb_v_i((vb_i16)q->type);
    if (!strcmp(name, "returnsrecords")) { if (set) return vb_v_empty(); return vb_v_bool(-1); }
    if (!strcmp(name, "updatable")) return vb_v_bool(-1);
    if (!strcmp(name, "connect")) { if (set) { free(q->connect); q->connect = dao_strdup_v(argv[argc - 1]); return vb_v_empty(); } return dao_vstr(q->connect); }
    if (!strcmp(name, "recordsaffected")) return vb_v_l((vb_i32)q->affected);
    if (!strcmp(name, "odbctimeout")) { if (set) { q->odbctimeout = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_i((vb_i16)q->odbctimeout); }
    if (!strcmp(name, "maxrecords")) { if (set) { q->maxrecords = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(q->maxrecords); }
    if (!strcmp(name, "datecreated")) return vb_v_date(q->created);
    if (!strcmp(name, "lastupdated")) return vb_v_date(q->updated);
    if (!strcmp(name, "fields")) {                  /* the columns of its result (phase 12.4) */
        dcoll *c = coll_new(C_FIELDS, 0); c->loaded = 1;
        if (d && q->sql && !q->params->n) qd_fields(d, q, c);
        vbvar v;
        if (argc > 0) { int f2; v = coll_member((vbobj)c, "item", kind, argc, argv, &f2); }
        else v = dao_retobj((vbobj)c);
        vb_orelease((vbobj)c);
        return v;
    }
    if (!strcmp(name, "openrecordset")) {
        if (!d) { dao_raise(3219, "DAO.QueryDef", "Invalid operation."); return vb_v_empty(); }
        vbobj r = rs_open(d, q->sql, along(argc, argv, 0, 0), along(argc, argv, 1, 0), along(argc, argv, 2, 0), o, 0, 0);
        return r ? dao_retobj(r) : vb_v_empty();
    }
    if (!strcmp(name, "execute")) {
        if (!d) { dao_raise(3219, "DAO.QueryDef", "Invalid operation."); return vb_v_empty(); }
        long n;
        if (db_execute(d, q->sql, along(argc, argv, 0, 0), o, &n) == 0) { q->affected = n; d->affected = n; }
        return vb_v_empty();
    }
    if (!strcmp(name, "close")) return vb_v_empty();
    if (!strcmp(name, "cancel")) return vb_v_empty();
    if (!strcmp(name, "createproperty")) return prop_create(o, argc, argv);
    if (!strcmp(name, "properties")) return props_get(o, kind, argc, argv);
    *found = 0;
    return vb_v_empty();
}

static void param_clear(vbobj o) { dpar *p = (dpar *)o; free(p->name); vb_vfree(&p->value); }
static vbvar param_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dpar *p = (dpar *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    if (!strcmp(name, "value") || !strcmp(name, "")) {
        if (set) { vb_vset(&p->value, vb_v_letval(argv[argc - 1])); return vb_v_empty(); }
        return vb_vtmp(p->value);
    }
    if (!strcmp(name, "name")) return dao_vstr(p->name);
    if (!strcmp(name, "type")) { if (set) { p->type = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_i((vb_i16)p->type); }
    if (!strcmp(name, "direction")) { if (set) { p->dir = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_i((vb_i16)p->dir); }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Error */
static void error_clear(vbobj o) { derr *e = (derr *)o; free(e->desc); free(e->src); }
static vbvar error_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    derr *e = (derr *)vb_oid(o);
    *found = 1;
    if (!strcmp(name, "number")) return vb_v_l(e->number);
    if (!strcmp(name, "description") || !strcmp(name, "")) return dao_vstr(e->desc);
    if (!strcmp(name, "source")) return dao_vstr(e->src);
    if (!strcmp(name, "helpcontext")) return vb_v_l(5000000 + e->number);
    if (!strcmp(name, "helpfile")) return dao_vstr("jeterr40.chm");
    *found = 0;
    return vb_v_empty();
}

static void engine_clear(vbobj o)
{
    deng *e = (deng *)o;
    if (e->wss) vb_orelease((vbobj)e->wss);
    if (e->errs) vb_orelease((vbobj)e->errs);
    free(e->inipath); free(e->systemdb); free(e->defuser); free(e->defpwd);
}

static void rsclear(vbobj o) { rs_clear(o); }

#include "daoprops.inc"

/* ---------------------------------------------------------------- the classes */
#define FB(f) 0, 0, 0, f
const vb_class vb_cls_dao_dbengine = { "DBEngine", sizeof(deng), 0, 0, engine_clear, 0, 0, "workspaces", FB(engine_member) };
const vb_class vb_cls_dao_workspace = { "Workspace", sizeof(dws), 0, 0, ws_clear, 0, 0, "databases", FB(ws_member) };
const vb_class vb_cls_dao_database = { "Database", sizeof(ddb), 0, 0, db_clear, 0, 0, "tabledefs", FB(db_member) };
const vb_class vb_cls_dao_tabledef = { "TableDef", sizeof(dtd), dao_tabledef_init, 0, tabledef_clear, 0, 0, "fields", FB(tabledef_member) };
const vb_class vb_cls_dao_field = { "Field", sizeof(dfld), dao_field_init, 0, field_clear, 0, 0, "value", FB(field_member) };
const vb_class vb_cls_dao_index = { "Index", sizeof(dix), dao_index_init, 0, index_clear, 0, 0, "name", FB(index_member) };
const vb_class vb_cls_dao_relation = { "Relation", sizeof(drel), dao_relation_init, 0, relation_clear, 0, 0, "fields", FB(relation_member) };
const vb_class vb_cls_dao_querydef = { "QueryDef", sizeof(dqd), dao_querydef_init, 0, querydef_clear, 0, 0, "parameters", FB(querydef_member) };
const vb_class vb_cls_dao_parameter = { "Parameter", sizeof(dpar), 0, 0, param_clear, 0, 0, "value", FB(param_member) };
const vb_class vb_cls_dao_recordset = { "Recordset", sizeof(drs), 0, 0, rsclear, 0, 0, "fields", FB(rs_member) };
const vb_class vb_cls_dao_error = { "Error", sizeof(derr), 0, 0, error_clear, 0, 0, "description", FB(error_member) };
const vb_class vb_cls_dao_coll = { "Collection", sizeof(dcoll), 0, 0, coll_clear, 0, 0, "item", coll_each_count, coll_each_item, 0, coll_member };
const vb_class vb_cls_dao_property = { "Property", sizeof(dprop), 0, 0, prop_clear, 0, 0, "value", FB(prop_member) };
vbvar dao_props_of(vbobj owner) { return props_of(owner); }
vbvar dao_props_get(vbobj owner, vb_i32 kind, vb_i32 argc, vbvar *argv) { return props_get(owner, kind, argc, argv); }

/* ---------------------------------------------------------------- the hooks of the rest of the runtime */
static void rs_watch(vbobj o, void (*cb)(void *, int, int, int *), void *ctx)
{
    o = vb_oid(o);
    if (!o || o->cls != &vb_cls_dao_recordset) return;
    drs *r = (drs *)o;
    r->watch = cb; r->watch_ctx = ctx;
}
static vbobj createobject(const char *progid)
{
    if (!strcmp(progid, "dao.dbengine") || !strcmp(progid, "dao.privatedbengine")) return dao_engine();
    extern vbobj ado_createobject(const char *progid), bind_createobject(const char *progid), rdo_createobject(const char *progid);
    vbobj o = ado_createobject(progid);
    if (!o) o = rdo_createobject(progid);
    return o ? o : bind_createobject(progid);
}
extern void ado_listen(vbobj o, void (*fn)(void *, vbobj, const char *, int, vbvar *), void *ctx, int on);
extern int ado_is_rs(vbobj o);
extern void bind_prop_changed(vbobj o, const char *name);
extern int ado_shape(vbobj o, int *ncols, int *cur);
extern vbvar ado_cell(vbobj o, int row, int col);
extern const char *ado_colname(vbobj o, int col);
extern int ado_coltype(vbobj o, int col);
extern int ado_goto(vbobj o, int row);
extern void ado_consumer(vbobj o);
extern vbvar ado_bookmark(vbobj o, int row);
extern int ado_bmrow(vbobj o, vbvar bm);
extern int dao_rs_shape(vbobj o, int *ncols, int *cur);
extern vbvar dao_rs_cell(vbobj o, int row, int col);
extern const char *dao_rs_colname(vbobj o, int col);
extern int dao_rs_coltype(vbobj o, int col);
extern vbvar dao_rs_bookmark(vbobj o, int row);
extern int dao_rs_bmrow(vbobj o, vbvar bm);
static int is_dao_rs(vbobj o) { o = o ? vb_oid(o) : 0; return o && o->cls == &vb_cls_dao_recordset; }
void vb_db_init(void)
{
    vb_dbhooks.is_dao_rs = is_dao_rs;
    vb_dbhooks.dao_shape = dao_rs_shape;
    vb_dbhooks.dao_cell = dao_rs_cell;
    vb_dbhooks.dao_colname = dao_rs_colname;
    vb_dbhooks.dao_coltype = dao_rs_coltype;
    vb_dbhooks.dao_bookmark = dao_rs_bookmark;
    vb_dbhooks.dao_bmrow = dao_rs_bmrow;
    vb_dbhooks.ado_shape = ado_shape;
    vb_dbhooks.ado_cell = ado_cell;
    vb_dbhooks.ado_colname = ado_colname;
    vb_dbhooks.ado_coltype = ado_coltype;
    vb_dbhooks.ado_goto = ado_goto;
    vb_dbhooks.ado_consumer = ado_consumer;
    vb_dbhooks.ado_bookmark = ado_bookmark;
    vb_dbhooks.ado_bmrow = ado_bmrow;
    vb_dbhooks.ado_listen = ado_listen;
    vb_dbhooks.is_ado_rs = ado_is_rs;
    vb_dbhooks.prop_changed = bind_prop_changed;
    vb_dbhooks.engine = dao_engine;
    vb_dbhooks.rs_watch = rs_watch;
    vb_dbhooks.createobject = createobject;
    extern vbobj rdo_rs_inner(vbobj o);
    vb_dbhooks.rdo_inner = rdo_rs_inner;
}
