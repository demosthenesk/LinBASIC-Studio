/* daoprops.inc - the Properties collections of DAO (phase 12.4): the built-in properties of each object in Jet's
 * order with their types (Type, Inherited, Value: the object's own member, or Jet's error where Jet gives one), and
 * the properties a program adds (CreateProperty, Append, Delete), kept in MSysLB_Properties of the database (VB6
 * oracle tests/db/d_dao4). Included by dao.c.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <math.h>

/* a built-in property: its name, its DataTypeEnum (-1: the type of the field), and how its value is read */
enum { PV_GET = 0, PV_EMPTYSTR, PV_EMPTY, PV_FALSE, PV_L1033, PV_I0, PV_I60 };
typedef struct pspec { const char *name; int type; int err; int how; } pspec;

static const pspec PS_ENGINE[] = { { "Version", 12, 0, 0 }, { "LoginTimeout", 3, 0, 0 }, { "IniPath", 12, 0, 0 },
    { "DefaultUser", 12, 3251, 0 }, { "DefaultPassword", 12, 3251, 0 }, { "DefaultType", 4, 0, 0 },
    { "SystemDB", 12, 0, 0 }, { 0 } };
static const pspec PS_WS[] = { { "Name", 12, 0, 0 }, { "UserName", 12, 0, 0 }, { "IsolateODBCTrans", 1, 0, PV_I0 },
    { "LoginTimeout", 4, 3251, 0 }, { "DefaultCursorDriver", 4, 3251, 0 }, { "Type", 4, 0, 0 }, { 0 } };
static const pspec PS_DB[] = { { "Name", 12, 0, 0 }, { "Connect", 12, 0, 0 }, { "Transactions", 1, 0, 0 },
    { "Updatable", 1, 0, 0 }, { "CollatingOrder", 3, 0, PV_L1033 }, { "QueryTimeout", 3, 0, 0 }, { "Version", 12, 0, 0 },
    { "RecordsAffected", 4, 0, 0 }, { "ReplicaID", 15, 0, PV_EMPTYSTR }, { "DesignMasterID", 15, 0, PV_EMPTYSTR },
    { "Connection", 0, 3251, 0 }, { 0 } };
static const pspec PS_TD[] = { { "Name", 12, 0, 0 }, { "Updatable", 1, 0, 0 }, { "DateCreated", 8, 0, 0 },
    { "LastUpdated", 8, 0, 0 }, { "Connect", 12, 0, 0 }, { "Attributes", 4, 0, 0 }, { "SourceTableName", 12, 0, 0 },
    { "RecordCount", 4, 0, 0 }, { "ValidationRule", 12, 0, 0 }, { "ValidationText", 12, 0, 0 },
    { "ConflictTable", 0, 0, PV_EMPTYSTR }, { "ReplicaFilter", 0, 0, PV_EMPTY }, { 0 } };
/* a field of a TableDef, an Index, a Relation or a QueryDef; a field of a Recordset differs in four (below) */
static const pspec PS_FLD[] = { { "Value", -1, 3219, 0 }, { "Attributes", 4, 0, 0 }, { "CollatingOrder", 3, 0, PV_L1033 },
    { "Type", 3, 0, 0 }, { "Name", 12, 0, 0 }, { "OrdinalPosition", 3, 0, 0 }, { "Size", 4, 0, 0 },
    { "SourceField", 12, 0, 0 }, { "SourceTable", 12, 0, 0 }, { "ValidateOnSet", 1, 3219, 0 },
    { "DataUpdatable", 1, 0, 0 }, { "ForeignName", 12, 3219, 0 }, { "DefaultValue", 12, 0, 0 },
    { "ValidationRule", 12, 0, 0 }, { "ValidationText", 12, 0, 0 }, { "Required", 1, 0, 0 },
    { "AllowZeroLength", 1, 0, 0 }, { "FieldSize", 4, 3267, 0 }, { "OriginalValue", -1, 3251, 0 },
    { "VisibleValue", -1, 3251, 0 }, { 0 } };
static const pspec PS_RSFLD[] = { { "Value", -1, 0, 0 }, { "Attributes", 4, 0, 0 }, { "CollatingOrder", 3, 0, PV_L1033 },
    { "Type", 3, 0, 0 }, { "Name", 12, 0, 0 }, { "OrdinalPosition", 3, 0, 0 }, { "Size", 4, 0, 0 },
    { "SourceField", 12, 0, 0 }, { "SourceTable", 12, 0, 0 }, { "ValidateOnSet", 1, 0, PV_FALSE },
    { "DataUpdatable", 1, 0, 0 }, { "ForeignName", 12, 3219, 0 }, { "DefaultValue", 12, 0, 0 },
    { "ValidationRule", 12, 0, 0 }, { "ValidationText", 12, 0, 0 }, { "Required", 1, 0, 0 },
    { "AllowZeroLength", 1, 0, 0 }, { "FieldSize", 4, 3259, 0 }, { "OriginalValue", -1, 3251, 0 },
    { "VisibleValue", -1, 3251, 0 }, { 0 } };
static const pspec PS_IX[] = { { "Name", 12, 0, 0 }, { "Unique", 1, 0, 0 }, { "Clustered", 1, 0, 0 },
    { "Primary", 1, 0, 0 }, { "Foreign", 1, 0, 0 }, { "Required", 1, 0, 0 }, { "IgnoreNulls", 1, 0, 0 },
    { "DistinctCount", 4, 0, 0 }, { 0 } };
static const pspec PS_QD[] = { { "Name", 12, 0, 0 }, { "DateCreated", 8, 0, 0 }, { "LastUpdated", 8, 0, 0 },
    { "Type", 3, 0, 0 }, { "SQL", 12, 0, 0 }, { "Updatable", 1, 0, 0 }, { "Connect", 12, 0, 0 },
    { "ReturnsRecords", 1, 0, 0 }, { "ODBCTimeout", 3, 0, 0 }, { "RecordsAffected", 4, 0, 0 }, { "MaxRecords", 4, 0, 0 },
    { "StillExecuting", 1, 3251, 0 }, { "CacheSize", 4, 3251, 0 }, { "Prepare", 1, 3251, 0 }, { 0 } };
static const pspec PS_REL[] = { { "Name", 12, 0, 0 }, { "Table", 12, 0, 0 }, { "ForeignTable", 12, 0, 0 },
    { "Attributes", 4, 0, 0 }, { "PartialReplica", 0, 0, PV_FALSE }, { 0 } };
static const pspec PS_RS[] = { { "Name", 12, 0, 0 }, { "BOF", 1, 0, 0 }, { "BookMark", 12, 0, 0 },
    { "Bookmarkable", 1, 0, 0 }, { "DateCreated", 8, 3251, 0 }, { "EOF", 1, 0, 0 }, { "Filter", 12, 0, 0 },
    { "Index", 12, 3251, 0 }, { "LastUpdated", 8, 3251, 0 }, { "LockEdits", 1, 0, 0 }, { "NoMatch", 1, 0, 0 },
    { "RecordCount", 4, 0, 0 }, { "Sort", 12, 0, 0 }, { "Transactions", 1, 0, 0 }, { "Type", 3, 0, 0 },
    { "Updatable", 1, 0, 0 }, { "Restartable", 1, 0, 0 }, { "CacheSize", 4, 0, 0 }, { "EditMode", 3, 0, 0 },
    { "PercentPosition", 6, 0, 0 }, { "ValidationRule", 12, 0, 0 }, { "ValidationText", 12, 0, 0 },
    { "AbsolutePosition", 4, 0, 0 }, { "StillExecuting", 1, 3251, 0 }, { "Connection", 0, 3251, 0 },
    { "BatchCollisionCount", 4, 3251, 0 }, { "BatchCollisions", 0, 3251, 0 }, { "BatchSize", 4, 3251, 0 },
    { "RecordStatus", 3, 3251, 0 }, { "UpdateOptions", 4, 3251, 0 }, { 0 } };

/* where the properties of a program are kept: (kind, object, sub-object) of an appended object of a database */
static ddb *prop_home(vbobj o, char *kind, const char **obj, const char **sub)
{
    o = vb_oid(o);
    *obj = ""; *sub = "";
    if (o->cls == &vb_cls_dao_database) { *kind = 'D'; return (ddb *)o; }
    if (o->cls == &vb_cls_dao_tabledef) {
        dtd *t = (dtd *)o;
        *kind = 'T'; *obj = t->name;
        return t->appended && t->db ? DB(t->db) : 0;
    }
    if (o->cls == &vb_cls_dao_querydef) {
        dqd *q = (dqd *)o;
        *kind = 'Q'; *obj = q->name;
        return q->appended && !q->temp && q->db ? DB(q->db) : 0;
    }
    if (o->cls == &vb_cls_dao_index) {
        dix *x = (dix *)o;
        if (!x->appended || !x->td) return 0;
        dtd *t = (dtd *)vb_oid(x->td);
        *kind = 'I'; *obj = t->name; *sub = x->name;
        return t->appended && t->db ? DB(t->db) : 0;
    }
    if (o->cls == &vb_cls_dao_field) {
        dfld *f = (dfld *)o;
        if (f->fk != FK_TABLE || !f->appended || !f->owner || vb_oid(f->owner)->cls != &vb_cls_dao_tabledef) return 0;
        dtd *t = (dtd *)vb_oid(f->owner);
        *kind = 'F'; *obj = t->name; *sub = f->name;
        return t->appended && t->db ? DB(t->db) : 0;
    }
    return 0;
}

static void prop_table(ddb *d)
{
    sqlite3_exec(d->db, "CREATE TABLE IF NOT EXISTS MSysLB_Properties (kind TEXT, obj TEXT COLLATE NOCASE, "
                 "sub TEXT COLLATE NOCASE, name TEXT COLLATE NOCASE, type INT, value TEXT)", 0, 0, 0);
}

/* a value of a property of a program: as its type gives it */
static vbvar prop_typed(int type, const char *text)
{
    if (!text) return vb_v_null();
    switch (type) {
    case 1: return vb_v_bool(atol(text) ? -1 : 0);
    case 2: return vb_v_byte((vb_u8)atoi(text));
    case 3: return vb_v_i((vb_i16)atoi(text));
    case 4: return vb_v_l((vb_i32)atol(text));
    case 5: return vb_v_c((vb_cy)llround(atof(text) * 10000.0));
    case 6: return vb_v_f((float)atof(text));
    case 7: return vb_v_d(atof(text));
    case 8: return vb_v_date(atof(text));
    default: return dao_vstr(text);
    }
}

static char *prop_text(int type, vbvar v)
{
    char b[64];
    if (v.vt == VT_NULL || v.vt == VT_EMPTY) return 0;
    switch (type) {
    case 1: case 2: case 3: case 4: snprintf(b, sizeof b, "%ld", (long)vb_v2l(v)); return strdup(b);
    case 5: case 6: case 7: case 8: snprintf(b, sizeof b, "%.17g", vb_v2d(v)); return strdup(b);
    default: return strdup(vb_stoutf8(vb_v2s(v)));
    }
}

static dprop *prop_new(vbobj owner, const char *name, int type, int err, int how, int user)
{
    dprop *p = (dprop *)vb_onew(&vb_cls_dao_property);
    vb_oaddref((vbobj)p);
    p->name = strdup(name);
    p->type = type;
    p->err = err;
    p->how = how;
    p->user = user;
    p->value = vb_v_empty();
    if (owner) { p->owner = vb_oid(owner); vb_oaddref(p->owner); }
    return p;
}

static int field_type_of(vbobj o)
{
    vb_i32 save = vb_errpend;
    int t = (int)vb_v2l(vb_late_o(o, "type", VBM_GET, 0, 0));
    if (vb_errpend && !save) { vb_err_clear_pending(); t = 0; }
    return t;
}

/* Properties of an object: the built-in ones, then those of the program */
static vbvar props_of(vbobj owner)
{
    owner = vb_oid(owner);
    const vb_class *k = owner->cls;
    const pspec *ps = k == &vb_cls_dao_dbengine ? PS_ENGINE : k == &vb_cls_dao_workspace ? PS_WS :
                      k == &vb_cls_dao_database ? PS_DB : k == &vb_cls_dao_tabledef ? PS_TD :
                      k == &vb_cls_dao_index ? PS_IX : k == &vb_cls_dao_querydef ? PS_QD :
                      k == &vb_cls_dao_relation ? PS_REL : k == &vb_cls_dao_recordset ? PS_RS :
                      k == &vb_cls_dao_field ? (((dfld *)owner)->fk == FK_RS ? PS_RSFLD : PS_FLD) : 0;
    dcoll *c = coll_new(C_PROPERTIES, owner);
    c->loaded = 1;
    for (; ps && ps->name; ps++) {
        dprop *p = prop_new(owner, ps->name, ps->type < 0 ? field_type_of(owner) : ps->type, ps->err, ps->how, 0);
        coll_add(c, (vbobj)p);
        vb_orelease((vbobj)p);
    }
    if (k == &vb_cls_dao_recordset) c->extra = 2;      /* Jet's Count of a Recordset's Properties: 2 more the first
                                                           time it is read (VB6 oracle d_dao4) */
    char kind; const char *obj, *sub;
    ddb *d = prop_home(owner, &kind, &obj, &sub);
    if (d && !d->closed) {
        prop_table(d);
        sqlite3_stmt *st = 0;
        if (sqlite3_prepare_v2(d->db, "SELECT name, type, value FROM MSysLB_Properties WHERE kind = ?1 AND obj = ?2 "
                               "AND sub = ?3 ORDER BY rowid", -1, &st, 0) == SQLITE_OK) {
            char ks[2] = { kind, 0 };
            sqlite3_bind_text(st, 1, ks, -1, SQLITE_TRANSIENT);
            sqlite3_bind_text(st, 2, obj, -1, SQLITE_TRANSIENT);
            sqlite3_bind_text(st, 3, sub, -1, SQLITE_TRANSIENT);
            while (sqlite3_step(st) == SQLITE_ROW) {
                int t = sqlite3_column_int(st, 1);
                dprop *p = prop_new(owner, (const char *)sqlite3_column_text(st, 0), t, 0, 0, 1);
                p->appended = 1;
                vb_vset(&p->value, prop_typed(t, (const char *)sqlite3_column_text(st, 2)));   /* (its own copy) */
                coll_add(c, (vbobj)p);
                vb_orelease((vbobj)p);
            }
            sqlite3_finalize(st);
        }
    }
    vbvar v = dao_retobj((vbobj)c);
    vb_orelease((vbobj)c);
    return v;
}

/* CreateProperty(Name, Type, Value): a property of the program, not yet in a collection */
static vbvar prop_create(vbobj owner, vb_i32 argc, vbvar *argv)
{
    const char *nm = dao_missing(argc, argv, 0) ? "" : dao_str(argv[0]);
    int type = dao_missing(argc, argv, 1) ? 0 : (int)vb_v2l(argv[1]);
    dprop *p = prop_new(owner, nm, type, 0, 0, 1);
    if (!dao_missing(argc, argv, 2)) vb_vset(&p->value, argv[2]);
    vbvar v = dao_retobj((vbobj)p);
    vb_orelease((vbobj)p);
    return v;
}

static int prop_save(dprop *p, int insert)
{
    char kind; const char *obj, *sub;
    ddb *d = p->owner ? prop_home(p->owner, &kind, &obj, &sub) : 0;
    if (!d) return 0;
    prop_table(d);
    sqlite3_stmt *st = 0;
    const char *sql = insert ? "INSERT INTO MSysLB_Properties (kind, obj, sub, name, type, value) VALUES (?1, ?2, ?3, ?4, ?5, ?6)"
                             : "UPDATE MSysLB_Properties SET value = ?6, type = ?5 WHERE kind = ?1 AND obj = ?2 AND sub = ?3 AND name = ?4";
    if (sqlite3_prepare_v2(d->db, sql, -1, &st, 0) != SQLITE_OK) return 0;
    char ks[2] = { kind, 0 };
    char *txt = prop_text(p->type, p->value);
    sqlite3_bind_text(st, 1, ks, -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(st, 2, obj, -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(st, 3, sub, -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(st, 4, p->name, -1, SQLITE_TRANSIENT);
    sqlite3_bind_int(st, 5, p->type);
    if (txt) sqlite3_bind_text(st, 6, txt, -1, SQLITE_TRANSIENT); else sqlite3_bind_null(st, 6);
    sqlite3_step(st);
    sqlite3_finalize(st);
    free(txt);
    return 0;
}

static void prop_unsave(dprop *p)
{
    char kind; const char *obj, *sub;
    ddb *d = p->owner ? prop_home(p->owner, &kind, &obj, &sub) : 0;
    if (!d) return;
    prop_table(d);
    sqlite3_stmt *st = 0;
    if (sqlite3_prepare_v2(d->db, "DELETE FROM MSysLB_Properties WHERE kind = ?1 AND obj = ?2 AND sub = ?3 AND name = ?4",
                           -1, &st, 0) != SQLITE_OK) return;
    char ks[2] = { kind, 0 };
    sqlite3_bind_text(st, 1, ks, -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(st, 2, obj, -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(st, 3, sub, -1, SQLITE_TRANSIENT);
    sqlite3_bind_text(st, 4, p->name, -1, SQLITE_TRANSIENT);
    sqlite3_step(st);
    sqlite3_finalize(st);
}

static const char *prop_errtext(int e)
{
    switch (e) {
    case 3219: return "Invalid operation.";
    case 3251: return "Operation is not supported for this type of object.";
    case 3259: return "Invalid field data type.";
    case 3267: return "Property value must be set before using this method.";
    default: return "Invalid operation.";
    }
}

static vbvar prop_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    dprop *p = (dprop *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    if (!strcmp(name, "name")) {
        if (set) {
            if (p->appended || !p->user) { dao_raise(3219, "DAO.Property", "Invalid operation."); return vb_v_empty(); }
            free(p->name); p->name = dao_strdup_v(argv[argc - 1]);
            return vb_v_empty();
        }
        return dao_vstr(p->name);
    }
    if (!strcmp(name, "type")) {
        if (set) {
            if (p->appended || !p->user) { dao_raise(3219, "DAO.Property", "Invalid operation."); return vb_v_empty(); }
            p->type = (int)vb_v2l(argv[argc - 1]);
            return vb_v_empty();
        }
        return vb_v_i((vb_i16)p->type);
    }
    if (!strcmp(name, "inherited")) return vb_v_bool(0);
    if (!strcmp(name, "properties")) return props_of(o);
    if (!strcmp(name, "value") || !strcmp(name, "")) {
        if (p->user) {
            if (set) {
                vb_vset(&p->value, argv[argc - 1]);
                if (p->appended) prop_save(p, 0);
                return vb_v_empty();
            }
            vbvar r = p->value;                 /* (a temporary copy: the property keeps its own) */
            if (r.vt == VT_BSTR) r = vb_v_s(vb_scopy(r.u.s));
            return r;
        }
        if (p->err) { dao_raise(p->err, "DAO.Property", "%s", prop_errtext(p->err)); return vb_v_empty(); }
        char low[64];
        int i = 0;
        for (; p->name[i] && i < 63; i++) low[i] = (char)tolower((unsigned char)p->name[i]);
        low[i] = 0;
        if (set) return vb_late_o(p->owner, low, kind, argc, argv);
        switch (p->how) {
        case PV_EMPTYSTR: return dao_vstr("");
        case PV_EMPTY: return vb_v_empty();
        case PV_FALSE: return vb_v_bool(0);
        case PV_L1033: return vb_v_l(1033);
        case PV_I0: return vb_v_i(0);
        case PV_I60: return vb_v_i(60);
        }
        return vb_late_o(p->owner, low, VBM_GET, 0, 0);
    }
    *found = 0;
    return vb_v_empty();
}

static void prop_clear(vbobj o)
{
    dprop *p = (dprop *)o;
    free(p->name);
    vb_vfree(&p->value);
    if (p->owner) vb_orelease(p->owner);
}

/* Append / Delete / Item of a Properties collection (coll_member) */
static int props_append(dcoll *c, vbobj it)
{
    if (it->cls != &vb_cls_dao_property) { vb_raise(13); return 1; }
    dprop *p = (dprop *)it;
    if (coll_find(c, p->name) >= 0) {
        dao_raise(3367, "DAO.Property", "Cannot append.  An object with that name already exists in the collection.");
        return 1;
    }
    if (c->owner && p->owner != c->owner) {
        if (p->owner) vb_orelease(p->owner);
        p->owner = c->owner;
        vb_oaddref(p->owner);
    }
    p->appended = 1;
    prop_save(p, 1);
    return 0;
}

static int props_delete(dcoll *c, int i)
{
    dprop *p = (dprop *)vb_oid(c->items[i]);
    if (!p->user) { dao_raise(3384, "DAO.Properties", "Cannot delete a built-in property."); return 1; }
    prop_unsave(p);
    return 0;
}

/* .Properties, .Properties(x), .Properties(x) = v */
static vbvar props_get(vbobj owner, vb_i32 kind, vb_i32 argc, vbvar *argv)
{
    vbvar c = props_of(owner);
    if (argc == 0) return c;
    int f2;
    return coll_member(vb_v2obj(c), "item", kind, argc, argv, &f2);
}
