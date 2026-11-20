/* ado.c - Microsoft ActiveX Data Objects 2.x with the Jet OLE DB provider, written again over the
 * DAO of LinBASIC Studio: Connection, Recordset, Fields, Field, Command, Parameters, Parameter,
 * Errors, Error, with the members, the defaults, the errors (numbers, texts, Source) and the events
 * of the original (VB6 oracles tests/db/d_ado1, d_ado2); and with MSDASQL, the provider for ODBC
 * data sources (ado_odbc.inc).
 * A recordset of ADO keeps a DAO recordset with all its rows (the "base") and, when Filter or Sort
 * is set, a second one opened from it (the "view"); a recordset without a connection (Fields.Append,
 * a client cursor set apart from its connection) has its rows in a SQLite database in memory.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "daoint.h"
#include <ctype.h>
#include <math.h>

extern const vb_class vb_cls_ado_connection, vb_cls_ado_recordset, vb_cls_ado_fields, vb_cls_ado_field,
    vb_cls_ado_command, vb_cls_ado_parameters, vb_cls_ado_parameter, vb_cls_ado_errors, vb_cls_ado_error,
    vb_cls_ado_properties;

/* DataTypeEnum */
enum { adEmpty = 0, adSmallInt = 2, adInteger = 3, adSingle = 4, adDouble = 5, adCurrency = 6, adDate = 7, adBSTR = 8,
       adError = 10, adBoolean = 11, adVariant = 12, adDecimal = 14, adTinyInt = 16, adUnsignedTinyInt = 17,
       adUnsignedSmallInt = 18, adUnsignedInt = 19, adBigInt = 20, adGUID = 72, adBinary = 128, adChar = 129,
       adWChar = 130, adNumeric = 131, adDBDate = 133, adDBTime = 134, adDBTimeStamp = 135, adVarChar = 200,
       adLongVarChar = 201, adVarWChar = 202, adLongVarWChar = 203, adVarBinary = 204, adLongVarBinary = 205 };
enum { adOpenForwardOnly = 0, adOpenKeyset = 1, adOpenDynamic = 2, adOpenStatic = 3 };
enum { adLockReadOnly = 1, adLockPessimistic = 2, adLockOptimistic = 3, adLockBatchOptimistic = 4 };
enum { adUseNone = 1, adUseServer = 2, adUseClient = 3 };
enum { adCmdText = 1, adCmdTable = 2, adCmdStoredProc = 4, adCmdUnknown = 8, adCmdFile = 256, adCmdTableDirect = 512 };
enum { adEditNone = 0, adEditInProgress = 1, adEditAdd = 2 };
enum { adRsnAddNew = 1, adRsnDelete = 2, adRsnUpdate = 3, adRsnUndoUpdate = 4, adRsnUndoAddNew = 5, adRsnRequery = 7,
       adRsnClose = 9, adRsnMove = 10, adRsnFirstChange = 11, adRsnMoveFirst = 12, adRsnMoveNext = 13,
       adRsnMovePrevious = 14, adRsnMoveLast = 15 };
enum { adStatusOK = 1, adStatusCantDeny = 3, adStatusCancel = 4 };
enum { adFldMayDefer = 2, adFldUpdatable = 4, adFldUnknownUpdatable = 8, adFldFixed = 16, adFldIsNullable = 32,
       adFldMayBeNull = 64, adFldLong = 128 };

#define SRC_CN "ADODB.Connection"
#define SRC_RS "ADODB.Recordset"
#define SRC_FLD "ADODB.Field"
#define SRC_FLDS "ADODB.Fields"
#define SRC_CMD "ADODB.Command"
#define SRC_JET "Microsoft JET Database Engine"

static const char *MSG_CLOSED = "Operation is not allowed when the object is closed.";
static const char *MSG_OPEN = "Operation is not allowed when the object is open.";
static const char *MSG_BOFEOF = "Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record.";
static const char *MSG_CONTEXT = "Operation is not allowed in this context.";
static const char *MSG_NOUPDATE = "Current Recordset does not support updating. This may be a limitation of the provider, or of the selected locktype.";
static const char *MSG_NOTFOUND = "Item cannot be found in the collection corresponding to the requested name or ordinal.";
static const char *MSG_ARGS = "Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another.";
static const char *MSG_NOCONN = "The connection cannot be used to perform this operation. It is either closed or invalid in this context.";
static const char *MSG_MULTISTEP = "Multiple-step operation generated errors. Check each status value.";
static const char *MSG_BADSQL = "Invalid SQL statement; expected 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT', or 'UPDATE'.";

static char *sdup(const char *s) { return s ? strdup(s) : 0; }
static int missing(vb_i32 argc, vbvar *argv, int i) { return i >= argc || vb_ismissing(argv[i]); }
static vb_i32 along(vb_i32 argc, vbvar *argv, int i, vb_i32 d) { return missing(argc, argv, i) ? d : vb_v2l(argv[i]); }
static vbvar vstr(const char *s) { return vb_v_s(vb_sfromutf8(s ? s : "")); }
static vbvar retobj(vbobj o) { return vb_vtmp(vb_v_obj(o)); }
static int is_set(vb_i32 kind) { return kind == VBM_LET || kind == VBM_SET; }

static void raise(vb_i32 n, const char *src, const char *msg)
{
    vb_raise_full(n, vb_sfromutf8(src), vb_sfromutf8(msg));
}

/* ---------------------------------------------------------------- lists: Errors, Parameters, Properties */
typedef struct alist { vbobjh h; int kind; vbobj *items; int n, cap; vbobj owner; } alist;   /* owner: the Command of Parameters (not counted) */
enum { L_ERRORS = 1, L_PARAMS, L_PROPS };
static alist *alist_new(int kind)
{
    const vb_class *c = kind == L_ERRORS ? &vb_cls_ado_errors : kind == L_PARAMS ? &vb_cls_ado_parameters : &vb_cls_ado_properties;
    alist *l = (alist *)vb_onew(c);
    vb_oaddref((vbobj)l);
    l->kind = kind;
    return l;
}
static void alist_add(alist *l, vbobj o)
{
    if (l->n == l->cap) { l->cap = l->cap ? l->cap * 2 : 8; l->items = realloc(l->items, sizeof(vbobj) * (size_t)l->cap); }
    vb_oaddref(o);
    l->items[l->n++] = o;
}
static void alist_clear(alist *l) { for (int i = 0; i < l->n; i++) vb_orelease(l->items[i]); l->n = 0; }
static void alist_free(vbobj o) { alist *l = (alist *)o; alist_clear(l); free(l->items); }

/* ---------------------------------------------------------------- objects */
typedef struct aerr { vbobjh h; vb_i32 number, native; char *desc, *src, *state; } aerr;

typedef struct acn {
    vbobjh h;
    vbobj db;                           /* ddb (counted) while open */
    char *cs, *provider, *defdb;
    int state, cursorloc, mode, timeout, cmdtimeout, isolation, attrs, trans;
    alist *errs;
    struct lbo_conn *odbc;              /* MSDASQL: the ODBC connection (db is then a database in memory) */
    char *dbms;
    int odbc_ro, odbc_sql;
    int shape;                          /* MSDataShape over the provider (ado_shape.inc) */
} acn;

typedef struct adef { char *name; int type, size, attrs, prec, scale; } adef;

typedef struct ars {
    vbobjh h;
    vbobj cn;                           /* acn (counted) or 0 */
    vbobj db;                           /* ddb of the rows (counted) */
    vbobj base, view;                   /* drs (counted); view == base without Filter and Sort */
    int memdb;                          /* db is this recordset's own, in memory */
    int state, cursortype, locktype, cursorloc, maxrecords, cachesize, pagesize, options, status;
    int fixed_count_unknown;            /* RecordCount -1 */
    char *source;
    vbvar filter;
    char *sort;
    adef *defs; int ndefs;              /* Fields.Append; and the ADO types of a recordset in memory */
    int *orig_valid;
    struct ahook { void (*fn)(void *, vbobj, const char *, int, vbvar *); void *ctx; } *hooks;
    int nhooks;                         /* the listeners of the runtime (vb_dbhooks.ado_listen) */
    int storing;                        /* a listener is putting bound values in: no lb_fill */
    int odbc;                           /* the rows of an ODBC statement (defs have the provider's metadata) */
    struct otarget *ot;                 /* where their changes go on the server */
    int ocount, pdel;                   /* RecordCount of a server cursor; a delete that failed, pending */
    char *osql; vbvar *ovals; int novals;   /* the statement, for Requery */
    int exact;                          /* defs: the attributes as they are (MSDASQL, MSDataShape) */
    int shared;                         /* a chapter: db is its parent's (not closed) */
    struct achap **chaps;               /* MSDataShape: the chapter columns (ndefs of them) */
    vbobj chparent; int chcol, chrow;   /* a chapter: its parent (counted), column, the parent row */
    char *chcrit;                       /* a chapter: the criteria of its rows */
    int pendmove;                       /* MSDataShape: the second MoveComplete of Open comes with the first look at a row */
} ars;

typedef struct aflds { vbobjh h; vbobj rs; } aflds;
typedef struct afld { vbobjh h; vbobj rs; int col; } afld;
typedef struct acmd { vbobjh h; vbobj cn; char *text, *name; int type, timeout, prepared; alist *params; } acmd;
typedef struct apar { vbobjh h; char *name; int type, dir, size, attrs, prec, scale; vbvar value; } apar;

#define ARS(o) ((ars *)vb_oid(o))
#define ACN(o) ((acn *)vb_oid(o))

/* the provider MSDASQL (ado_odbc.inc) */
static int remote_err;                  /* the change on the server failed: its error is raised already */
static int odbc_open(acn *c);
static void odbc_close(acn *c);
static void odbc_props(acn *c, alist *l);
static vbvar odbc_trans(acn *c, int what);
static int odbc_run(acn *c, const char *sql, vbvar *vals, int nvals, ars *r, long *affected, int *rows);
static char *odbc_sql(const char *text, int type, int nparams, char **src);
static vbvar odbc_openschema(acn *c, int what, const char *rtable);
static void odbc_refresh_params(acmd *cmd);
static int odbc_pending(ars *r);
static int cursor_write(void *ctx, struct drs *d, int op);
static int shape_open(ars *r, acn *c, const char *text, int locktype);
static vbvar chap_value(ars *p, int col);
static void chap_sync(ars *ch);
static void chap_free(struct achap *a);
static int starts_word(const char *s, const char *w);
static void odbc_rs_props(ars *r, alist *l);
static void ot_free(struct otarget *t);
vbvar alist_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);

/* ---------------------------------------------------------------- errors of the provider */
/* a Jet error as the HRESULT the Jet OLE DB provider gives (ctx 1: Recordset.Update / a field) */
static vb_i32 hr_of(int jet, int ctx)
{
    switch (jet) {
    case 3078: case 3192: return -2147217865;               /* DB_E_NOTABLE */
    case 3061: return -2147217904;                          /* DB_E_PARAMNOTOPTIONAL */
    case 3034: return -2147168242;                          /* XACT_E_NOTRANSACTION */
    case 3075: case 3076: case 3077: case 3079: case 3085: case 3125: case 3129: case 3131: case 3133: case 3134:
    case 3135: case 3136: case 3138: case 3139: case 3141: case 3142: case 3143: case 3144: case 3145: case 3146:
    case 3306: case 3346: case 3065:
        return -2147217900;                                 /* DB_E_ERRORSINCOMMAND */
    case 3464: return -2147217913;                          /* DB_E_CANTCONVERTVALUE */
    case 3163: case 3349: case 3421: return -2147217887;    /* DB_E_ERRORSOCCURRED */
    }
    if (ctx == 1) return -2147217887;
    return -2147467259;                                     /* E_FAIL */
}
static vb_i32 native_of(int jet)
{
    if (jet == 3078) return -524420377;
    return -524420377 - (jet - 3078) * 65536;
}
static void provider_raise(acn *cn, vb_i32 hr, int jet, const char *msg)
{
    if (cn) {
        alist_clear(cn->errs);
        aerr *e = (aerr *)vb_onew(&vb_cls_ado_error);
        e->number = hr; e->native = native_of(jet);
        e->desc = strdup(msg); e->src = strdup(SRC_JET);
        char st[16]; snprintf(st, sizeof st, "%d", jet);
        e->state = strdup(st);
        alist_add(cn->errs, (vbobj)e);
    }
    raise(hr, SRC_JET, msg);
}
/* a DAO error pending: raised again as the provider's; 1 if there was one */
static int dao_failed(acn *cn, int ctx)
{
    if (!vb_errpend) return 0;
    if (remote_err) { remote_err = 0; return 1; }
    int jet = vb_err_number_pending();
    char *msg = strdup(vb_err_text_pending());
    vb_err_clear_pending();
    provider_raise(cn, hr_of(jet, ctx), jet, msg);
    free(msg);
    return 1;
}

/* ---------------------------------------------------------------- events */
static void fire(vbobj src, const char *ev, int argc, vbvar *argv)
{
    if (!src) return;
    ars *ar = src->cls == &vb_cls_ado_recordset ? (ars *)src : 0;
    if (ar && ar->nhooks) {                     /* the listeners of the runtime first */
        int n = ar->nhooks;
        struct ahook *hs = malloc(sizeof *hs * (size_t)n);
        memcpy(hs, ar->hooks, sizeof *hs * (size_t)n);
        vb_oaddref(src);
        for (int i = 0; i < n && !vb_errpend; i++) hs[i].fn(hs[i].ctx, src, ev, argc, argv);
        vb_orelease(src);
        free(hs);
    }
    if (!src->sinks) return;
    vbobj *ls; const vb_fn **ts;
    vb_i32 n = vb_sinks_begin(src, &ls, &ts);
    for (vb_i32 k = 0; k < n && !vb_errpend; k++)
        for (const vb_member *mm = (const vb_member *)ts[k]; mm && mm->name; mm++)
            if (!strcmp(mm->name, ev)) { mm->fn(ls[k], argc, argv); break; }
    vb_sinks_end(ls, ts, n);
}
static int listening(ars *r) { return r->h.sinks || r->nhooks; }
/* the runtime's own notes: the bound values into the record / the record into the bound controls */
static void note_save(ars *r)
{
    if (!r->nhooks || r->storing) return;
    r->storing = 1;
    fire((vbobj)r, "lb_save", 0, 0);
    r->storing = 0;
}
static void note_fill(ars *r) { if (r->nhooks && !r->storing) fire((vbobj)r, "lb_fill", 0, 0); }

static int ev_willmove(ars *r, int reason, int status)
{
    if (!listening(r)) return adStatusOK;
    vbvar a[3] = { vb_v_l(reason), vb_v_l(status), vb_v_obj((vbobj)r) };
    fire((vbobj)r, "willmove", 3, a);
    return vb_v2l(a[1]);
}
static void ev_movecomplete(ars *r, int reason)
{
    if (!listening(r)) return;
    vbvar a[4] = { vb_v_l(reason), vb_v_obj(0), vb_v_l(adStatusOK), vb_v_obj((vbobj)r) };
    fire((vbobj)r, "movecomplete", 4, a);
}
static void ev_willchangerecord(ars *r, int reason)
{
    if (!listening(r)) return;
    vbvar a[4] = { vb_v_l(reason), vb_v_l(1), vb_v_l(adStatusOK), vb_v_obj((vbobj)r) };
    fire((vbobj)r, "willchangerecord", 4, a);
}
static void ev_recordchangecomplete(ars *r, int reason)
{
    if (!listening(r)) return;
    vbvar a[5] = { vb_v_l(reason), vb_v_l(1), vb_v_obj(0), vb_v_l(adStatusOK), vb_v_obj((vbobj)r) };
    fire((vbobj)r, "recordchangecomplete", 5, a);
}
static void ev_willchangefield(ars *r)
{
    if (!listening(r)) return;
    vbvar a[4] = { vb_v_l(1), vb_v_empty(), vb_v_l(adStatusOK), vb_v_obj((vbobj)r) };
    fire((vbobj)r, "willchangefield", 4, a);
}
static void ev_fieldchangecomplete(ars *r)
{
    if (!listening(r)) return;
    vbvar a[5] = { vb_v_l(1), vb_v_empty(), vb_v_obj(0), vb_v_l(adStatusOK), vb_v_obj((vbobj)r) };
    fire((vbobj)r, "fieldchangecomplete", 5, a);
}
static void ev_recordsetchangecomplete(ars *r, int reason)
{
    if (!listening(r)) return;
    vbvar a[4] = { vb_v_l(reason), vb_v_obj(0), vb_v_l(adStatusOK), vb_v_obj((vbobj)r) };
    fire((vbobj)r, "recordsetchangecomplete", 4, a);
}
static void ev_endofrecordset(ars *r)
{
    if (!listening(r)) return;
    vbvar a[3] = { vb_v_bool(0), vb_v_l(adStatusOK), vb_v_obj((vbobj)r) };
    fire((vbobj)r, "endofrecordset", 3, a);
}
static acn *cn_of(ars *r);
static void ev_opened(ars *r)
{
    if (!listening(r)) return;
    if (cn_of(r) && cn_of(r)->shape) {          /* VB6 oracle d_de1 */
        ev_willmove(r, adRsnMove, adStatusCantDeny);
        ev_willmove(r, adRsnMove, adStatusOK);
        ev_movecomplete(r, adRsnMove);
        r->pendmove = 1;
        return;
    }
    ev_willmove(r, adRsnMove, adStatusCantDeny);
    ev_willmove(r, adRsnMove, adStatusOK);
    ev_movecomplete(r, adRsnMove);
    ev_movecomplete(r, adRsnMove);
}

/* the events of a Connection (WithEvents) */
static void cn_fire(acn *c, const char *ev, int argc, vbvar *argv)
{
    if (!c->h.sinks) return;
    void *e = vb_err_stash();                   /* an error of the operation waits until they ran */
    fire((vbobj)c, ev, argc, argv);
    if (e) { if (vb_errpend) vb_err_clear_pending(); vb_err_unstash(e); }
}
static void ev_willconnect(acn *c)
{
    if (!c->h.sinks) return;
    vbvar a[6] = { vstr(c->cs), vstr(""), vstr(""), vb_v_l(-1), vb_v_l(adStatusOK), vb_v_obj((vbobj)c) };
    cn_fire(c, "willconnect", 6, a);
}
static void ev_connectcomplete(acn *c, int ok)
{
    if (!c->h.sinks) return;
    vbvar a[3] = { vb_v_obj(0), vb_v_l(ok ? adStatusOK : 2), vb_v_obj((vbobj)c) };
    cn_fire(c, "connectcomplete", 3, a);
}
static void ev_willexecute(acn *c, const char *src, int ct, int lt, int opts, vbobj cmd, vbobj rs)
{
    if (!c || !c->h.sinks) return;
    vbvar a[8] = { vstr(src), vb_v_l(ct), vb_v_l(lt), vb_v_l(opts), vb_v_l(adStatusOK), vb_v_obj(cmd), vb_v_obj(rs), vb_v_obj((vbobj)c) };
    cn_fire(c, "willexecute", 8, a);
}
static void ev_executecomplete(acn *c, long affected, int ok, vbobj cmd, vbobj rs)
{
    if (!c || !c->h.sinks) return;
    vbvar a[6] = { vb_v_l((vb_i32)affected), vb_v_obj(0), vb_v_l(ok ? adStatusOK : 2), vb_v_obj(cmd), vb_v_obj(rs), vb_v_obj((vbobj)c) };
    cn_fire(c, "executecomplete", 6, a);
}

/* ---------------------------------------------------------------- rows */
static drs *VW(ars *r) { return r->view ? (drs *)vb_oid(r->view) : 0; }
static int alive(drs *d, int i) { return i >= 0 && i < d->d->n && !d->d->r[i].del; }
static int nalive(drs *d) { int k = 0; for (int i = 0; i < d->d->n; i++) if (!d->d->r[i].del) k++; return k; }
static int first_alive(drs *d) { for (int i = 0; i < d->d->n; i++) if (!d->d->r[i].del) return i; return -1; }
static int last_alive(drs *d) { for (int i = d->d->n - 1; i >= 0; i--) if (!d->d->r[i].del) return i; return -1; }
static int is_bof(drs *d) { return d->editmode != 2 && (d->cur < 0 || nalive(d) == 0); }
static int is_eof(drs *d) { return d->editmode != 2 && (d->cur >= d->d->n || nalive(d) == 0); }
static void go_first(drs *d) { int i = first_alive(d); d->cur = i >= 0 ? i : d->d->n; }

static int is_open(ars *r)
{
    if (r->state != 1) return 0;
    drs *v = VW(r);
    if (!v || v->closed) { r->state = 0; return 0; }
    return 1;
}
static int need_open(ars *r, const char *src)
{
    if (is_open(r)) return 0;
    raise(3704, src, MSG_CLOSED);
    return -1;
}
static acn *cn_of(ars *r) { return r->cn ? ACN(r->cn) : 0; }

/* the ADO type of a column */
static int ado_type(ars *r, int col)
{
    drs *v = VW(r);
    if (r->memdb && r->defs && col < r->ndefs) return r->defs[col].type;
    switch (v->cols[col].type) {
    case dbBoolean: return adBoolean;
    case dbByte: return adUnsignedTinyInt;
    case dbInteger: return adSmallInt;
    case dbLong: return adInteger;
    case dbCurrency: return adCurrency;
    case dbSingle: return adSingle;
    case dbDouble: return adDouble;
    case dbDate: return adDate;
    case dbBinary: case dbVarBinary: return adVarBinary;
    case dbLongBinary: return adLongVarBinary;
    case dbMemo: return adLongVarWChar;
    case dbGUID: return adGUID;
    case dbDecimal: case dbNumeric: return adNumeric;
    case dbBigInt: return adBigInt;
    default: return adVarWChar;
    }
}
static int defined_size(ars *r, int col)
{
    drs *v = VW(r);
    if (r->memdb && r->defs && col < r->ndefs && r->defs[col].size) return r->defs[col].size;
    switch (ado_type(r, col)) {
    case adBoolean: case adSmallInt: return 2;
    case adUnsignedTinyInt: return 1;
    case adInteger: case adSingle: return 4;
    case adCurrency: case adDouble: case adDate: case adBigInt: return 8;
    case adGUID: return 16;
    case adNumeric: return 19;
    case adLongVarWChar: return 536870910;
    case adLongVarBinary: return 1073741823;
    default: return v->cols[col].size > 0 ? v->cols[col].size : 255;
    }
}
static int precision_of(int t)
{
    switch (t) {
    case adInteger: return 10;
    case adSmallInt: return 5;
    case adUnsignedTinyInt: return 3;
    case adDouble: return 15;
    case adSingle: return 7;
    case adCurrency: return 19;
    case adBigInt: return 19;
    case adNumeric: return 18;
    default: return 255;
    }
}
static int field_attrs(ars *r, int col)
{
    drs *v = VW(r);
    int t = ado_type(r, col);
    if ((r->odbc || r->exact) && r->defs && col < r->ndefs) return r->defs[col].attrs;
    if (r->memdb && r->defs && col < r->ndefs)
        return (r->defs[col].attrs | adFldUpdatable | adFldUnknownUpdatable) & ~adFldUpdatable;
    rcol *c = &v->cols[col];
    int a = adFldMayDefer | adFldUnknownUpdatable | adFldMayBeNull;
    if (t != adVarWChar && t != adLongVarWChar && t != adVarBinary && t != adLongVarBinary) a |= adFldFixed;
    if (t == adLongVarWChar || t == adLongVarBinary) a |= adFldLong;
    if (!(c->attrs & dbAutoIncrField) && t != adBoolean) a |= adFldIsNullable;
    if (!c->stable && t == adDouble) a &= ~adFldUnknownUpdatable;   /* Avg(...) (VB6 oracle d_ado1) */
    return a;
}
/* a value as ADO gives it: a Byte as an Integer */
static vbvar ado_value(vbvar v)
{
    if (v.vt == VT_UI1) return vb_v_i((vb_i16)v.u.b);
    return v;
}

static int col_of(ars *r, vbvar key, const char *src)
{
    drs *v = VW(r);
    vbvar k = vb_v_letval(key);
    if (k.vt == VT_BSTR) {
        char *nm = vb_stoutf8(k.u.s);
        if (r->memdb && r->defs && r->ndefs == v->ncols) {
            for (int i = 0; i < r->ndefs; i++) if (!strcasecmp(r->defs[i].name, nm)) return i;
        }
        for (int i = 0; i < v->ncols; i++) if (!strcasecmp(v->cols[i].name, nm)) return i;
    } else {
        vb_i32 save = vb_errpend;
        vb_i32 i = vb_v2l(k);
        if (vb_errpend && !save) { vb_err_clear_pending(); i = -1; }
        if (i >= 0 && i < v->ncols) return i;
    }
    raise(3265, src, MSG_NOTFOUND);
    return -1;
}

/* the value of a column of the current record, or an error raised */
static int cur_value(ars *r, int col, vbvar *out)
{
    drs *v = VW(r);
    if (r->chaps && col < r->ndefs && r->chaps[col]) {      /* a chapter: the recordset of its rows */
        if (!alive(v, v->cur)) { raise(3021, SRC_FLD, MSG_BOFEOF); return -1; }
        *out = chap_value(r, col);
        return vb_errpend ? -1 : 0;
    }
    if (v->editmode && v->ebuf) {
        /* a client cursor knows the counter of a new record only after Update (VB6 oracle f_ado) */
        if (v->editmode == 2 && r->cursorloc == adUseClient && !r->memdb && ((v->cols[col].attrs & dbAutoIncrField) || !v->edirty[col])) {
            *out = vb_v_null();                 /* nor the defaults of the fields it has not been given */
            return 0;
        }
        *out = ado_value(v->ebuf[col]);
        return 0;
    }
    if (v->cur >= 0 && v->cur < v->d->n && v->d->r[v->cur].del) {
        provider_raise(cn_of(r), -2147217885, 0, "Row handle referred to a deleted row or a row marked for deletion.");
        return -1;
    }
    if (!alive(v, v->cur)) { raise(3021, SRC_FLD, MSG_BOFEOF); return -1; }
    *out = ado_value(v->d->r[v->cur].v[col]);
    return 0;
}

/* a member of the DAO recordset, its error raised as the provider's */
static vbvar dao_call(ars *r, vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int ctx)
{
    int f;
    vbvar v = rs_member(o, name, kind, argc, argv, &f);
    if (dao_failed(cn_of(r), ctx)) return vb_v_empty();
    return v;
}

/* ---------------------------------------------------------------- Jet SQL with parameters */
/* names Jet does not know become parameters, as with the Jet OLE DB provider; the jsql of the statement
 * (the binds: the parameters in the order Jet takes them) or -1 with the error raised */
static int discover(acn *cn, ddb *d, const char *sql, jsql *j, char ***pnames, int *pn)
{
    char **names = 0; int nn = 0;
    for (int tries = 0; tries < 64; tries++) {
        const char *list[72];
        for (int i = 0; i < nn && i < 70; i++) list[i] = names[i];
        list[nn < 70 ? nn : 70] = 0;
        jet_ansi = 1;
        int rc = dao_translate(d, sql, j, nn ? list : 0);
        jet_ansi = 0;
        if (rc < 0) {
            char msg[600]; snprintf(msg, sizeof msg, "%s", j->errmsg);
            int e = j->err;
            jsql_free(j);
            for (int i = 0; i < nn; i++) free(names[i]);
            free(names);
            provider_raise(cn, hr_of(e, 0), e, msg);
            return -1;
        }
        if (j->kind >= J_CREATE_TABLE && j->kind <= J_DROP_CONSTRAINT) break;
        if (!j->sql) break;
        sqlite3_stmt *st = 0;
        int prc = sqlite3_prepare_v2(d->db, j->sql, -1, &st, 0);
        if (prc == SQLITE_OK) { sqlite3_finalize(st); break; }
        const char *m = sqlite3_errmsg(d->db);
        if (!strncmp(m, "no such column: ", 16) && nn < 64) {
            const char *nm = m + 16;             /* "Nope" - should this be ...? / T.x */
            char buf[256]; int k = 0;
            if (*nm == '"') { nm++; while (*nm && *nm != '"' && k < 255) buf[k++] = *nm++; }
            else while (*nm && !(nm[0] == ' ' && nm[1] == '-') && k < 255) buf[k++] = *nm++;
            buf[k] = 0;
            names = realloc(names, sizeof(char *) * (size_t)(nn + 1));
            names[nn++] = strdup(buf);
            jsql_free(j);
            continue;
        }
        char msg[1024];
        int e = jet_sqlite_error(d->db, prc, msg, sizeof msg);
        jsql_free(j);
        for (int i = 0; i < nn; i++) free(names[i]);
        free(names);
        provider_raise(cn, hr_of(e, 0), e, msg);
        return -1;
    }
    *pnames = names; *pn = nn;
    return 0;
}

static int starts_word(const char *s, const char *w)
{
    while (*s && (isspace((unsigned char)*s) || *s == '(')) s++;
    size_t n = strlen(w);
    return !strncasecmp(s, w, n) && (!s[n] || isspace((unsigned char)s[n]) || s[n] == '(' || s[n] == '[' || s[n] == '*');
}
static int looks_sql(const char *s)
{
    static const char *const K[] = { "select", "insert", "update", "delete", "parameters", "transform", "create",
                                     "drop", "alter", "grant", "revoke", "procedure", 0 };
    for (int i = 0; K[i]; i++) if (starts_word(s, K[i])) return 1;
    return 0;
}

/* run a statement: a recordset (*out) for a SELECT, the rows affected otherwise; vals: the values of
 * the parameters in order (the declared ones in their order). 0 or -1 (error raised) */
static int run_sql(acn *cn, ddb *d, const char *sql, vbvar *vals, int nvals, int daotype, vbobj *out, long *affected)
{
    *out = 0; if (affected) *affected = -1;
    jsql j;
    char **names; int nn;
    if (discover(cn, d, sql, &j, &names, &nn)) return -1;
    int nb = j.nbinds;
    if (nvals < nb) {
        char msg[200];
        if (j.nparams) snprintf(msg, sizeof msg, "Too few parameters. Expected %d.", nb - nvals);
        else snprintf(msg, sizeof msg, "No value given for one or more required parameters.");
        jsql_free(&j);
        for (int i = 0; i < nn; i++) free(names[i]);
        free(names);
        provider_raise(cn, -2147217904, 3061, msg);
        return -1;
    }
    /* a QueryDef of the statement, its parameters with the values */
    dqd *q = querydef_new((vbobj)d, "", sql);
    for (int i = 0; i < nb; i++) {
        char plain[256]; dao_param_plain(j.binds[i], plain, sizeof plain);
        int k = coll_find(q->params, plain);
        dpar *p;
        if (k < 0) { p = param_new(plain, dbText); coll_add(q->params, (vbobj)p); vb_orelease((vbobj)p); }
        else p = (dpar *)q->params->items[k];
        int vi = i;
        if (j.nparams) {                        /* declared: in the order of PARAMETERS */
            for (int m = 0; m < j.nparams; m++) {
                char pm[256]; dao_param_plain(j.params[m].name, pm, sizeof pm);
                if (!strcasecmp(pm, plain)) { vi = m; break; }
            }
        }
        vbvar v = vi < nvals ? vb_v_letval(vals[vi]) : vb_v_null();
        if (v.vt == VT_EMPTY) v = vb_v_null();
        vb_vset(&p->value, v);
    }
    int returns = j.returns && (j.kind == J_SELECT || j.kind == J_TRANSFORM) && !j.into;
    jsql_free(&j);
    for (int i = 0; i < nn; i++) free(names[i]);
    free(names);
    int r = 0;
    jet_ansi = 1;
    if (returns) {
        vbobj o = rs_open(d, sql, daotype, 0, 0, (vbobj)q, 0, 0);
        jet_ansi = 0;
        if (dao_failed(cn, 0) || !o) r = -1;
        else { *out = o; vb_oaddref(o); }
    } else {
        long n = 0;
        int e = dao_execute_qd(d, sql, (vbobj)q, &n);
        jet_ansi = 0;
        if (dao_failed(cn, 0) || e) r = -1;
        else if (affected) *affected = n;
    }
    vb_orelease((vbobj)q);
    return r;
}

/* what a CommandText is: a table, a saved query, SQL; *sql: the statement, *src: the Source */
static int resolve_cmd(acn *cn, ddb *d, const char *text, int type, char **sql, char **src)
{
    char real[256];
    *sql = 0; *src = 0;
    int t = type & ~(128 | 16 | 32 | 64);       /* adExecuteNoRecords, async flags */
    if (t == adCmdTable || t == adCmdTableDirect) {
        jbuf s = { 0 };
        jb_s(&s, "SELECT * FROM "); jb_s(&s, "["); jb_s(&s, text); jb_s(&s, "]");
        *sql = s.p;
        jbuf k = { 0 };
        jb_s(&k, "select * from "); jb_s(&k, text);
        *src = k.p;
        return 0;
    }
    if (t == adCmdStoredProc) {
        int what = schema_has_table(d, text, real, sizeof real);
        if (!what) {
            char msg[600];
            snprintf(msg, sizeof msg, "The Microsoft Jet database engine cannot find the input table or query '%s'.  Make sure it exists and that its name is spelled correctly.", text);
            provider_raise(cn, -2147217865, 3078, msg);
            return -1;
        }
        jbuf s = { 0 };
        jb_s(&s, "SELECT * FROM ["); jb_s(&s, real); jb_s(&s, "]");
        if (what == 2) {
            sqlite3_stmt *st = 0;
            if (sqlite3_prepare_v2(d->db, "SELECT sql FROM MSysLB_Queries WHERE name = ?1", -1, &st, 0) == SQLITE_OK) {
                sqlite3_bind_text(st, 1, real, -1, SQLITE_STATIC);
                if (sqlite3_step(st) == SQLITE_ROW) { jb_free(&s); jb_s(&s, (const char *)sqlite3_column_text(st, 0)); }
                sqlite3_finalize(st);
            }
        }
        *sql = s.p; *src = strdup(text);
        return 0;
    }
    if (looks_sql(text)) { *sql = strdup(text); *src = strdup(text); return 0; }
    /* adCmdUnknown / adCmdText: a table, or a query without parameters, by its name */
    int what = schema_has_table(d, text, real, sizeof real);
    if (what == 1) {
        jbuf s = { 0 }; jb_s(&s, "SELECT * FROM ["); jb_s(&s, real); jb_s(&s, "]");
        *sql = s.p; *src = strdup(text);
        return 0;
    }
    if (what == 2) {
        char *qs = 0;
        sqlite3_stmt *st = 0;
        if (sqlite3_prepare_v2(d->db, "SELECT sql FROM MSysLB_Queries WHERE name = ?1", -1, &st, 0) == SQLITE_OK) {
            sqlite3_bind_text(st, 1, real, -1, SQLITE_STATIC);
            if (sqlite3_step(st) == SQLITE_ROW) qs = strdup((const char *)sqlite3_column_text(st, 0));
            sqlite3_finalize(st);
        }
        if (qs && !starts_word(qs, "parameters")) { *sql = qs; *src = strdup(text); return 0; }
        free(qs);
    }
    provider_raise(cn, -2147217900, 3129, MSG_BADSQL);
    return -1;
}

/* ---------------------------------------------------------------- Connection */
static char *cs_get(const char *cs, const char *key)
{
    size_t kl = strlen(key);
    const char *p = cs;
    while (p && *p) {
        while (*p == ';' || isspace((unsigned char)*p)) p++;
        const char *eq = strchr(p, '=');
        if (!eq) break;
        const char *ke = eq;
        while (ke > p && isspace((unsigned char)ke[-1])) ke--;
        const char *v = eq + 1;
        while (isspace((unsigned char)*v)) v++;
        const char *ve;
        if (*v == '"' || *v == '\'') { char q = *v++; ve = strchr(v, q); if (!ve) ve = v + strlen(v); }
        else { ve = strchr(v, ';'); if (!ve) ve = v + strlen(v); }
        if ((size_t)(ke - p) == kl && !strncasecmp(p, key, kl)) {
            const char *e2 = ve;
            while (e2 > v && isspace((unsigned char)e2[-1])) e2--;
            char *r = malloc((size_t)(e2 - v) + 1);
            memcpy(r, v, (size_t)(e2 - v)); r[e2 - v] = 0;
            return r;
        }
        p = *ve ? ve + 1 : ve;
        if (*p == '"' || *p == '\'') p++;
    }
    return 0;
}

static int cn_open(acn *c)
{
    if (c->state) { raise(3705, SRC_CN, MSG_OPEN); return -1; }
    {   /* MSDataShape: the provider under it opens the data, the shape commands are ours */
        char *pv = cs_get(c->cs ? c->cs : "", "Provider");
        int shape = pv && !strncasecmp(pv, "MSDataShape", 11);
        free(pv);
        if (shape) {
            char *dp = cs_get(c->cs, "Data Provider");
            jbuf n = { 0 };
            if (dp && *dp && strcasecmp(dp, "NONE")) { jb_s(&n, "Provider="); jb_s(&n, dp); }
            const char *p = c->cs;
            while (p && *p) {                   /* the other keys as they are */
                while (*p == ';' || isspace((unsigned char)*p)) p++;
                if (!*p) break;
                const char *st = p, *eq = strchr(p, '=');
                if (!eq) break;
                const char *v = eq + 1;
                while (isspace((unsigned char)*v)) v++;
                const char *e;
                if (*v == '"' || *v == '\'') { char q = *v; e = strchr(v + 1, q); e = e ? e + 1 : v + strlen(v); }
                else e = v + strcspn(v, ";");
                size_t kl = (size_t)(eq - st);
                while (kl && isspace((unsigned char)st[kl - 1])) kl--;
                int skip = (kl == 8 && !strncasecmp(st, "Provider", 8)) || (kl == 13 && !strncasecmp(st, "Data Provider", 13));
                if (!skip) { if (n.n) jb_s(&n, ";"); jb_put(&n, st, (int)(e - st)); }
                p = e;
            }
            free(dp);
            char *keep = c->cs;
            c->cs = n.p ? n.p : strdup("");
            int e = cn_open(c);
            free(c->cs); c->cs = keep;
            if (e) return -1;
            c->shape = 1;
            free(c->provider); c->provider = strdup("MSDataShape.1");
            c->cursorloc = adUseClient;
            return 0;
        }
    }
    if (odbc_open(c)) return c->state ? 0 : -1;
    const char *cs = c->cs ? c->cs : "";
    char *prov = cs_get(cs, "Provider");
    char *path = cs_get(cs, "Data Source");
    if (!path) path = cs_get(cs, "DBQ");
    if (!path) {
        char *dsn = cs_get(cs, "DSN");
        free(prov);
        if (dsn) { free(dsn); raise(-2147467259, "Microsoft OLE DB Provider for ODBC Drivers", "[Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified"); }
        else raise(-2147467259, "Microsoft OLE DB Provider for ODBC Drivers", "[Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified");
        return -1;
    }
    char *mode = cs_get(cs, "Mode");
    int ro = (c->mode == 1) || (mode && !strncasecmp(mode, "read", 4) && !strcasestr(mode, "write"));
    free(mode);
    vbobj d = dao_open_path(path, ro);
    if (dao_failed(c, 0) || !d) { free(path); free(prov); return -1; }
    vb_oaddref(d);
    c->db = d;
    free(c->provider);
    c->provider = prov ? prov : strdup("MSDASQL.1");
    if (prov && !strncasecmp(prov, "Microsoft.Jet.OLEDB", 19) && !strchr(prov + 19, '.')) {}
    free(path);
    c->state = 1;
    c->trans = 0;
    return 0;
}

static void cn_close(acn *c)
{
    c->shape = 0;
    if (c->odbc) { odbc_close(c); c->trans = 0; }
    if (c->db) {
        ddb *d = DB(c->db);
        if (c->trans && d->db) sqlite3_exec(d->db, "ROLLBACK", 0, 0, 0);
        dao_db_close(c->db);
        vb_orelease(c->db);
        c->db = 0;
    }
    c->state = 0;
    c->trans = 0;
}

static void cn_init(vbobj o)
{
    acn *c = (acn *)o;
    c->cursorloc = adUseServer; c->timeout = 15; c->cmdtimeout = 30; c->isolation = 4096;
    c->errs = alist_new(L_ERRORS);
    c->provider = strdup("MSDASQL");
    c->cs = strdup("");
    c->defdb = strdup("");
}
static void cn_clear(vbobj o)
{
    acn *c = (acn *)o;
    cn_close(c);
    free(c->cs); free(c->provider); free(c->defdb); free(c->dbms);
    if (c->errs) vb_orelease((vbobj)c->errs);
}

static vbobj rs_new_for(acn *c);
static int rs_from_drs(ars *r, vbobj o, int fwd);
static vbvar cn_openschema(acn *c, vb_i32 argc, vbvar *argv);

static vbvar cn_execute(acn *c, vb_i32 argc, vbvar *argv)
{
    if (!c->state) { raise(3704, SRC_CN, MSG_CLOSED); return vb_v_empty(); }
    if (argc < 1) { vb_raise(450); return vb_v_empty(); }
    char *text = strdup(vb_stoutf8(vb_v2s(argv[0])));
    int opts = along(argc, argv, 2, -1);
    if (opts == -1) opts = adCmdUnknown;
    if (c->shape) {
        const char *q = text; while (isspace((unsigned char)*q)) q++;
        if (starts_word(q, "shape")) {
            ars *r = (ars *)rs_new_for(c);
            int e = shape_open(r, c, text, adLockReadOnly);
            free(text);
            if (e) { vb_orelease((vbobj)r); return vb_v_empty(); }
            if (argc > 1 && !vb_ismissing(argv[1])) vb_vset(&argv[1], vb_v_l(-1));
            vbvar v = retobj((vbobj)r);
            vb_orelease((vbobj)r);
            return v;
        }
    }
    if (c->odbc) {
        char *src;
        char *sql = odbc_sql(text, opts, 0, &src);
        free(text);
        ars *r = (ars *)rs_new_for(c);
        long n = 0; int rows = 0;
        int e = odbc_run(c, sql, 0, 0, r, &n, &rows);
        if (e) { free(sql); free(src); vb_orelease((vbobj)r); return vb_v_empty(); }
        if (argc > 1 && !vb_ismissing(argv[1])) vb_vset(&argv[1], vb_v_l(rows ? -1 : (vb_i32)n));
        if (rows) {
            r->cursortype = c->cursorloc == adUseClient ? adOpenStatic : adOpenForwardOnly;
            r->locktype = adLockReadOnly;
            r->cursorloc = c->cursorloc;
            r->state = 1;
            free(r->source); r->source = src; src = 0;
            r->osql = sql; sql = 0;
        }
        free(sql); free(src);
        vbvar v = retobj((vbobj)r);
        vb_orelease((vbobj)r);
        return v;
    }
    ddb *d = DB(c->db);
    char *sql, *src;
    if (resolve_cmd(c, d, text, opts, &sql, &src)) { free(text); return vb_v_empty(); }
    free(text);
    vbobj o = 0; long n = 0;
    int e = run_sql(c, d, sql, 0, 0, dbOpenSnapshot, &o, &n);
    free(sql);
    if (e) { free(src); return vb_v_empty(); }
    if (argc > 1 && !vb_ismissing(argv[1])) vb_vset(&argv[1], vb_v_l(o ? -1 : (vb_i32)n));
    ars *r = (ars *)rs_new_for(c);
    if (o) {
        r->cursortype = c->cursorloc == adUseClient ? adOpenStatic : adOpenForwardOnly;
        r->locktype = adLockReadOnly;
        r->cursorloc = c->cursorloc;
        rs_from_drs(r, o, r->cursortype == adOpenForwardOnly);
        vb_orelease(o);
        free(r->source); r->source = src; src = 0;
    }
    free(src);
    vbvar v = retobj((vbobj)r);
    vb_orelease((vbobj)r);
    return v;
}

static vbvar cn_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    acn *c = (acn *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (!strcmp(name, "state")) return vb_v_l(c->state);
    if (!strcmp(name, "connectionstring") || !strcmp(name, "")) {
        if (set) {
            if (c->state) { raise(3705, SRC_CN, MSG_OPEN); return vb_v_empty(); }
            free(c->cs); c->cs = strdup(vb_stoutf8(vb_v2s(argv[argc - 1])));
            char *p = cs_get(c->cs, "Provider");
            if (p) { free(c->provider); c->provider = p; }
            return vb_v_empty();
        }
        return vstr(c->cs);
    }
    if (!strcmp(name, "provider")) {
        if (set) { free(c->provider); c->provider = strdup(vb_stoutf8(vb_v2s(argv[argc - 1]))); return vb_v_empty(); }
        return vstr(c->provider);
    }
    if (!strcmp(name, "cursorlocation")) { if (set) { c->cursorloc = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(c->cursorloc); }
    if (!strcmp(name, "mode")) { if (set) { c->mode = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(c->mode); }
    if (!strcmp(name, "connectiontimeout")) { if (set) { c->timeout = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(c->timeout); }
    if (!strcmp(name, "commandtimeout")) { if (set) { c->cmdtimeout = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(c->cmdtimeout); }
    if (!strcmp(name, "isolationlevel")) { if (set) { c->isolation = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(c->isolation); }
    if (!strcmp(name, "attributes")) { if (set) { c->attrs = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(c->attrs); }
    if (!strcmp(name, "defaultdatabase")) {
        if (set) { free(c->defdb); c->defdb = strdup(vb_stoutf8(vb_v2s(argv[argc - 1]))); return vb_v_empty(); }
        return vstr(c->defdb);
    }
    if (!strcmp(name, "version")) return vstr("2.8");
    if (!strcmp(name, "errors")) {
        if (argc == 0) return retobj((vbobj)c->errs);
        int f2; extern vbvar alist_member(vbobj, const char *, vb_i32, vb_i32, vbvar *, int *);
        return alist_member((vbobj)c->errs, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "properties")) {
        alist *l = alist_new(L_PROPS);
        if (c->odbc && c->state) odbc_props(c, l);
        vbvar v;
        if (argc == 0) v = retobj((vbobj)l);
        else { int f2; v = alist_member((vbobj)l, "item", kind, argc, argv, &f2); }
        vb_orelease((vbobj)l);
        return v;
    }
    if (!strcmp(name, "open")) {
        if (c->state) { raise(3705, SRC_CN, MSG_OPEN); return vb_v_empty(); }
        if (!missing(argc, argv, 0)) {
            free(c->cs); c->cs = strdup(vb_stoutf8(vb_v2s(argv[0])));
        }
        ev_willconnect(c);
        if (vb_errpend) return vb_v_empty();
        cn_open(c);
        ev_connectcomplete(c, !vb_errpend);
        return vb_v_empty();
    }
    if (!strcmp(name, "close")) {
        if (!c->state) { raise(3704, SRC_CN, MSG_CLOSED); return vb_v_empty(); }
        cn_close(c);
        return vb_v_empty();
    }
    if (!strcmp(name, "execute")) return cn_execute(c, argc, argv);
    if (!strcmp(name, "begintrans")) {
        if (!c->state) { raise(3704, SRC_CN, MSG_CLOSED); return vb_v_empty(); }
        if (c->odbc) return odbc_trans(c, 0);
        char sp[40]; snprintf(sp, sizeof sp, "SAVEPOINT adotrans%d", c->trans + 1);
        sqlite3_exec(DB(c->db)->db, sp, 0, 0, 0);
        c->trans++;
        return vb_v_l(c->trans);
    }
    if (!strcmp(name, "committrans") || !strcmp(name, "rollbacktrans")) {
        if (!c->state) { raise(3704, SRC_CN, MSG_CLOSED); return vb_v_empty(); }
        if (c->odbc) return odbc_trans(c, name[0] == 'c' ? 1 : 2);
        if (!c->trans) {
            provider_raise(c, -2147168242, 3034, "You tried to commit or rollback a transaction without first beginning a transaction.");
            return vb_v_empty();
        }
        char sp[80];
        if (name[0] == 'c') snprintf(sp, sizeof sp, "RELEASE adotrans%d", c->trans);
        else snprintf(sp, sizeof sp, "ROLLBACK TO adotrans%d; RELEASE adotrans%d", c->trans, c->trans);
        sqlite3_exec(DB(c->db)->db, sp, 0, 0, 0);
        c->trans--;
        return vb_v_empty();
    }
    if (!strcmp(name, "openschema")) return cn_openschema(c, argc, argv);
    if (!strcmp(name, "lb_willexecute")) {      /* the DataEnvironment: a command whose recordset is open */
        ev_willexecute(c, argc ? vb_stoutf8(vb_v2s(argv[0])) : "", -1, -1, -1, 0, 0);
        return vb_v_empty();
    }
    if (!strcmp(name, "cancel")) return vb_v_empty();
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Recordset */
static void rs_drop_rows(ars *r)
{
    if (r->view && r->view != r->base) { rs_close((drs *)vb_oid(r->view)); vb_orelease(r->view); }
    else if (r->view) vb_orelease(r->view);
    if (r->base) { rs_close((drs *)vb_oid(r->base)); vb_orelease(r->base); }
    r->view = r->base = 0;
    if (r->db) {
        if (r->memdb && !r->shared) dao_db_close(r->db);
        vb_orelease(r->db);
        r->db = 0;
    }
    r->memdb = 0; r->shared = 0; r->pendmove = 0;
    free(r->orig_valid); r->orig_valid = 0;
    ot_free(r->ot); r->ot = 0;
    if (r->chaps) { for (int i = 0; i < r->ndefs; i++) chap_free(r->chaps[i]); free(r->chaps); r->chaps = 0; }
    r->exact = 0;
    r->odbc = 0; r->pdel = 0;
}

static void rs_init(vbobj o)
{
    ars *r = (ars *)o;
    r->cursortype = adOpenForwardOnly; r->locktype = adLockReadOnly; r->cursorloc = adUseServer;
    r->cachesize = 1; r->options = -1;
    r->source = strdup("");
    r->sort = strdup("");
    r->filter = vb_v_l(0);
}
static void rs_clear2(vbobj o)
{
    ars *r = (ars *)o;
    rs_drop_rows(r);
    if (r->cn) vb_orelease(r->cn);
    free(r->source); free(r->sort);
    vb_vfree(&r->filter);
    for (int i = 0; i < r->ndefs; i++) free(r->defs[i].name);
    free(r->defs);
    free(r->hooks);
    free(r->osql);
    for (int i = 0; i < r->novals; i++) vb_vfree(&r->ovals[i]);
    free(r->ovals);
    if (r->chparent) vb_orelease(r->chparent);
    free(r->chcrit);
}

static vbobj rs_new_for(acn *c)
{
    ars *r = (ars *)vb_onew(&vb_cls_ado_recordset);
    vb_oaddref((vbobj)r);
    if (c) { r->cn = (vbobj)c; vb_oaddref((vbobj)c); r->cursorloc = c->cursorloc; }
    return (vbobj)r;
}

/* the rows of a DAO recordset become the recordset's; positioned at the first */
static int rs_from_drs(ars *r, vbobj o, int fwd)
{
    (void)fwd;
    rs_drop_rows(r);
    drs *d = (drs *)vb_oid(o);
    r->base = o; vb_oaddref(o);
    r->view = o; vb_oaddref(o);
    r->db = d->db; vb_oaddref(r->db);
    go_first(d);
    r->state = 1;
    return 0;
}

static int updatable(ars *r) { return r->locktype != adLockReadOnly; }

/* a pending change saved before the recordset moves; -1 if it failed */
static int flush_edit(ars *r)
{
    drs *v = VW(r);
    if (!v || !v->editmode) return 0;
    int add = v->editmode == 2;
    ev_willchangerecord(r, adRsnUpdate);
    dao_call(r, r->view, "update", VBM_METHOD, 0, 0, 1);
    if (vb_errpend) return -1;
    if (add && v->lastmod >= 0) v->cur = v->lastmod;
    ev_recordchangecomplete(r, adRsnUpdate);
    return 0;
}

/* the view again: Filter and Sort over the base */
static int rebuild_view(ars *r)
{
    drs *b = (drs *)vb_oid(r->base);
    vbvar f = r->filter;
    int has_filter = (f.vt == VT_BSTR && f.u.s.len > 0) || (r->chcrit && *r->chcrit);
    int has_sort = r->sort && *r->sort;
    if (r->view && r->view != r->base) { rs_close((drs *)vb_oid(r->view)); vb_orelease(r->view); }
    else if (r->view) vb_orelease(r->view);
    r->view = 0;
    if (!has_filter && !has_sort) {
        r->view = r->base; vb_oaddref(r->base);
        go_first(b);
        return 0;
    }
    /* the rows of the base, filtered and sorted by SQLite */
    char *crit;
    if (r->chcrit && *r->chcrit && f.vt == VT_BSTR && f.u.s.len > 0) {     /* a chapter: its rows, then the Filter */
        jbuf cb = { 0 };
        jb_s(&cb, "("); jb_s(&cb, r->chcrit); jb_s(&cb, ") AND ("); jb_s(&cb, vb_stoutf8(f.u.s)); jb_s(&cb, ")");
        crit = cb.p;
    } else if (r->chcrit && *r->chcrit) crit = strdup(r->chcrit);
    else crit = has_filter ? strdup(vb_stoutf8(f.u.s)) : strdup("");
    for (char *p = crit, q = 0; *p; p++) {      /* LIKE 'A%' as Jet's 'A*' */
        if (q) { if (*p == q) q = 0; else if (*p == '%') *p = '*'; else if (*p == '_') *p = '?'; }
        else if (*p == '\'' || *p == '"') q = *p;
    }
    vb_sset(&b->filter, vb_sfromutf8(crit));
    vb_sset(&b->sort, vb_sfromutf8(has_sort ? r->sort : ""));
    free(crit);
    int fnd;
    vbvar ty = vb_v_l(b->type == dbOpenTable ? dbOpenDynaset : b->type);
    vbvar v = rs_member(r->base, "openrecordset", VBM_METHOD, 1, &ty, &fnd);
    vb_sset(&b->filter, vb_sfromutf8("")); vb_sset(&b->sort, vb_sfromutf8(""));
    if (vb_errpend) {
        vb_err_clear_pending();
        r->view = r->base; vb_oaddref(r->base);
        raise(3001, SRC_RS, MSG_ARGS);
        return -1;
    }
    r->view = vb_v2obj(v); vb_oaddref(r->view);
    go_first(VW(r));
    return 0;
}

/* the rows into a database in memory: the recordset no longer needs its connection */
static const char *jet_ddl_type(int t, int size, char *buf, int cap)
{
    switch (t) {
    case adBoolean: return "BIT";
    case adUnsignedTinyInt: case adTinyInt: return "BYTE";
    case adSmallInt: case adUnsignedSmallInt: return "SHORT";
    case adInteger: case adUnsignedInt: case adBigInt: return "LONG";
    case adCurrency: return "CURRENCY";
    case adSingle: return "SINGLE";
    case adDouble: case adNumeric: case adDecimal: return "DOUBLE";
    case adDate: case adDBDate: case adDBTime: case adDBTimeStamp: return "DATETIME";
    case adLongVarChar: case adLongVarWChar: return "MEMO";
    case adBinary: case adVarBinary: return "BINARY";
    case adLongVarBinary: return "LONGBINARY";
    case adGUID: return "GUID";
    default:
        if (size <= 0 || size > 255) return "MEMO";
        snprintf(buf, (size_t)cap, "TEXT(%d)", size);
        return buf;
    }
}
static int make_memory(ars *r, adef *defs, int n)
{
    vbobj mdb = dao_memory_db();
    if (!mdb) { raise(-2147467259, "Microsoft Cursor Engine", "Unspecified error"); return -1; }
    ddb *d = DB(mdb);
    jbuf s = { 0 };
    jb_s(&s, "CREATE TABLE [lbrs] (");
    for (int i = 0; i < n; i++) {
        char buf[32];
        if (i) jb_s(&s, ", ");
        int dup = 0;                            /* SELECT T.Name, U.Name: two fields of one name */
        for (int k = 0; k < i; k++) if (!strcasecmp(defs[k].name, defs[i].name)) dup++;
        jb_s(&s, "["); jb_s(&s, defs[i].name);
        if (dup) jb_f(&s, "~%d", dup + 1);
        jb_s(&s, "] ");
        jb_s(&s, jet_ddl_type(defs[i].type, defs[i].size, buf, sizeof buf));
    }
    jb_s(&s, ")");
    long aff;
    int e = dao_execute_qd(d, s.p, 0, &aff);
    jb_free(&s);
    if (e) { vb_err_clear_pending(); vb_orelease(mdb); raise(-2147467259, "Microsoft Cursor Engine", "Unspecified error"); return -1; }
    vbobj o = rs_open(d, "lbrs", dbOpenDynaset, 0, 0, 0, 0, 0);
    if (!o) { vb_err_clear_pending(); vb_orelease(mdb); raise(-2147467259, "Microsoft Cursor Engine", "Unspecified error"); return -1; }
    rs_drop_rows(r);
    r->db = mdb; r->memdb = 1;
    r->base = o; vb_oaddref(o);
    r->view = o; vb_oaddref(o);
    if (defs != r->defs) {
        for (int i = 0; i < r->ndefs; i++) free(r->defs[i].name);
        free(r->defs);
        r->defs = malloc(sizeof(adef) * (size_t)(n ? n : 1));
        for (int i = 0; i < n; i++) { r->defs[i] = defs[i]; r->defs[i].name = strdup(defs[i].name); }
        r->ndefs = n;
    }
    return 0;
}
/* a client recordset set apart from its connection: its rows into memory */
static int disconnect(ars *r)
{
    drs *b = (drs *)vb_oid(r->base);
    int n = b->ncols;
    adef *defs = calloc((size_t)(n ? n : 1), sizeof(adef));
    for (int i = 0; i < n; i++) {
        defs[i].name = strdup(b->cols[i].name);
        r->view = r->base;                      /* ado_type of the base's columns */
        defs[i].type = ado_type(r, i);
        defs[i].size = b->cols[i].size;
        defs[i].attrs = field_attrs(r, i);
    }
    /* the rows */
    int nrows = b->d->n;
    vbvar **rows = malloc(sizeof(vbvar *) * (size_t)(nrows ? nrows : 1));
    int k = 0;
    for (int i = 0; i < nrows; i++) {
        if (b->d->r[i].del) continue;
        rows[k] = calloc((size_t)(n ? n : 1), sizeof(vbvar));
        for (int c = 0; c < n; c++) vb_vset(&rows[k][c], b->d->r[i].v[c]);
        k++;
    }
    vbobj keepbase = r->base; vb_oaddref(keepbase);
    int e = make_memory(r, defs, n);
    vb_orelease(keepbase);
    if (!e) {
        drs *m = (drs *)vb_oid(r->base);
        ddb *d = DB(r->db);
        jbuf s = { 0 };
        jb_s(&s, "INSERT INTO lbrs VALUES (");
        for (int c = 0; c < n; c++) jb_f(&s, c ? ", ?%d" : "?%d", c + 1);
        jb_s(&s, ")");
        sqlite3_stmt *st = 0;
        if (sqlite3_prepare_v2(d->db, s.p, -1, &st, 0) == SQLITE_OK) {
            for (int i = 0; i < k; i++) {
                sqlite3_reset(st);
                for (int c = 0; c < n; c++) jet_bind(st, c + 1, rows[i][c], m->cols[c].type);
                sqlite3_step(st);
            }
            sqlite3_finalize(st);
        }
        jb_free(&s);
        int f2;
        rs_member(r->base, "requery", VBM_METHOD, 0, 0, &f2);
        if (vb_errpend) vb_err_clear_pending();
        rebuild_view(r);
    }
    for (int i = 0; i < k; i++) { for (int c = 0; c < n; c++) vb_vfree(&rows[i][c]); free(rows[i]); }
    free(rows);
    for (int i = 0; i < n; i++) free(defs[i].name);
    free(defs);
    if (r->cn) { vb_orelease(r->cn); r->cn = 0; }
    return e;
}

#include "ado_odbc.inc"
#include "ado_shape.inc"

/* Open */
static vbvar rs_open_core(ars *r, vb_i32 argc, vbvar *argv, acn **execd, acmd **execcmd);
static vbvar rs_open_m(ars *r, vb_i32 argc, vbvar *argv)
{
    acn *c = 0; acmd *cmd = 0;
    vbvar v = rs_open_core(r, argc, argv, &c, &cmd);
    if (c) {                                    /* WillExecute ran: ExecuteComplete */
        ev_executecomplete(c, -1, !vb_errpend, cmd ? (vbobj)cmd : 0, (vbobj)r);
        vb_orelease((vbobj)c);
    }
    return v;
}
static vbvar rs_open_core(ars *r, vb_i32 argc, vbvar *argv, acn **execd, acmd **execcmd)
{
    if (r->state == 1 && is_open(r)) { raise(3705, SRC_RS, MSG_OPEN); return vb_v_empty(); }
    vbvar src = missing(argc, argv, 0) ? vb_v_missing() : argv[0];
    acmd *cmd = 0;
    if (!vb_ismissing(src) && vb_isobject(src)) {
        vbobj so = vb_oid(vb_v2obj(src));
        if (so && so->cls == &vb_cls_ado_command) cmd = (acmd *)so;
    }
    /* the connection */
    if (!missing(argc, argv, 1)) {
        vbvar a = argv[1];
        if (vb_isobject(a)) {
            vbobj co = vb_oid(vb_v2obj(a));
            if (co && co->cls == &vb_cls_ado_connection) { vb_oaddref(co); if (r->cn) vb_orelease(r->cn); r->cn = co; }
        } else {
            acn *c = (acn *)vb_onew(&vb_cls_ado_connection);
            vb_oaddref((vbobj)c);
            free(c->cs); c->cs = strdup(vb_stoutf8(vb_v2s(a)));
            if (cn_open(c)) { vb_orelease((vbobj)c); return vb_v_empty(); }
            if (r->cn) vb_orelease(r->cn);
            r->cn = (vbobj)c;
        }
    } else if (cmd && cmd->cn && !r->cn) { r->cn = cmd->cn; vb_oaddref(r->cn); }
    if (!missing(argc, argv, 2)) r->cursortype = vb_v2l(argv[2]);
    if (!missing(argc, argv, 3)) r->locktype = vb_v2l(argv[3]);
    int opts = along(argc, argv, 4, r->options);
    if (opts == -1) opts = adCmdUnknown;
    if (r->cursorloc == adUseClient) r->cursortype = adOpenStatic;
    else if (r->cursortype == adOpenDynamic) r->cursortype = adOpenKeyset;

    /* without a source: the fields appended (a recordset of the program's own) */
    if (vb_ismissing(src) && !cmd) {
        if (!r->ndefs) { raise(3709, SRC_RS, MSG_NOCONN); return vb_v_empty(); }
        r->cursorloc = adUseClient; r->cursortype = adOpenStatic;
        if (r->locktype == adLockReadOnly && missing(argc, argv, 3)) r->locktype = adLockBatchOptimistic;
        if (make_memory(r, r->defs, r->ndefs)) return vb_v_empty();
        r->state = 1;
        go_first(VW(r));
        ev_opened(r);
        return vb_v_empty();
    }
    acn *c = cn_of(r);
    if (!c || !c->state) { raise(3709, SRC_RS, MSG_NOCONN); return vb_v_empty(); }
    if (c->h.sinks) {
        const char *t = cmd ? (cmd->text ? cmd->text : "") : vb_stoutf8(vb_v2s(src));
        char *keep = strdup(t);
        ev_willexecute(c, keep, missing(argc, argv, 2) ? -1 : vb_v2l(argv[2]), missing(argc, argv, 3) ? -1 : vb_v2l(argv[3]),
                       missing(argc, argv, 4) ? -1 : vb_v2l(argv[4]), cmd ? (vbobj)cmd : 0, (vbobj)r);
        free(keep);
        *execd = c; vb_oaddref((vbobj)c); *execcmd = cmd;
        if (vb_errpend) return vb_v_empty();
    }
    if (c->shape) {
        const char *t = cmd ? (cmd->text ? cmd->text : "") : vb_stoutf8(vb_v2s(src));
        const char *q = t; while (isspace((unsigned char)*q)) q++;
        if (starts_word(q, "shape")) {
            char *text = strdup(t);             /* Source keeps the text as it is (" SHAPE ..." of a DataEnvironment) */
            int e = shape_open(r, c, text, r->locktype);
            free(text);
            if (e) return vb_v_empty();
            vb_vset(&r->filter, vb_v_l(0));
            free(r->sort); r->sort = strdup("");
            ev_opened(r);
            note_fill(r);
            return vb_v_empty();
        }
    }
    if (c->odbc) return odbc_rs_open(r, c, cmd, src, opts, argc, argv);
    ddb *d = DB(c->db);
    char *text = cmd ? strdup(cmd->text ? cmd->text : "") : strdup(vb_stoutf8(vb_v2s(src)));
    if (cmd && cmd->type != adCmdUnknown && opts == adCmdUnknown) opts = cmd->type;
    char *sql, *srcname;
    if (resolve_cmd(c, d, text, opts, &sql, &srcname)) { free(text); return vb_v_empty(); }
    free(text);
    vbvar *vals = 0; int nv = 0;
    if (cmd) {
        nv = cmd->params->n;
        vals = calloc((size_t)(nv ? nv : 1), sizeof(vbvar));
        for (int i = 0; i < nv; i++) vals[i] = ((apar *)cmd->params->items[i])->value;
    }
    int daotype = updatable(r) ? dbOpenDynaset : dbOpenSnapshot;
    vbobj o = 0; long n;
    int e = run_sql(c, d, sql, vals, nv, daotype, &o, &n);
    free(vals); free(sql);
    if (e) { free(srcname); return vb_v_empty(); }
    free(r->source); r->source = srcname;
    if (!o) return vb_v_empty();                /* an action query: the recordset stays closed */
    rs_from_drs(r, o, r->cursortype == adOpenForwardOnly);
    vb_orelease(o);
    vb_vset(&r->filter, vb_v_l(0));
    free(r->sort); r->sort = strdup("");
    ev_opened(r);
    note_fill(r);
    return vb_v_empty();
}

static void rs_close_m(ars *r)
{
    drs *v = VW(r);
    if (v && v->editmode) { int f; rs_member(r->view, "cancelupdate", VBM_METHOD, 0, 0, &f); }
    rs_drop_rows(r);
    r->state = 0;
    note_fill(r);
    ev_recordsetchangecomplete(r, adRsnClose);
}

static vb_i32 abs_pos(ars *r)
{
    drs *v = VW(r);
    if (r->cursortype == adOpenForwardOnly) return -1;
    if (r->odbc && r->cursorloc != adUseClient) return -1;     /* MSDASQL: not known */
    if (is_bof(v)) return -2;
    if (is_eof(v)) return -3;
    int k = 0;
    for (int i = 0; i <= v->cur && i < v->d->n; i++) if (!v->d->r[i].del) k++;
    return k ? k : 1;
}
static void set_abs_pos(ars *r, int n)
{
    drs *v = VW(r);
    int k = 0;
    for (int i = 0; i < v->d->n; i++) if (!v->d->r[i].del && ++k == n) { v->cur = i; return; }
    raise(3001, SRC_RS, MSG_ARGS);
}
static vb_i32 rcount(ars *r)
{
    if (r->cursortype == adOpenForwardOnly) return -1;
    if (r->odbc && r->cursorloc != adUseClient) return r->ocount;
    if (r->cursorloc != adUseClient && !r->memdb && r->locktype == adLockPessimistic) return -1;
    return nalive(VW(r));
}

/* a move: Validate-like events, a pending change saved; reason, then the move itself */
static int do_move(ars *r, int how, int n)
{
    /* how: 1 first, 2 previous, 3 next, 4 last, 5 relative n */
    drs *v = VW(r);
    if (how == 2 && r->cursortype == adOpenForwardOnly) { raise(3219, SRC_RS, MSG_CONTEXT); return -1; }
    if (how == 5 && n < 0 && r->cursortype == adOpenForwardOnly) { raise(3219, SRC_RS, MSG_CONTEXT); return -1; }
    if (how == 3 && is_eof(v) && !v->editmode) { raise(3021, SRC_RS, MSG_BOFEOF); return -1; }
    if (how == 2 && is_bof(v) && !v->editmode) { raise(3021, SRC_RS, MSG_BOFEOF); return -1; }
    if (how == 5 && n != 0 && (is_bof(v) || is_eof(v)) && !v->editmode) { raise(3021, SRC_RS, MSG_BOFEOF); return -1; }
    int reason = how == 1 ? adRsnMoveFirst : how == 2 ? adRsnMovePrevious : how == 3 ? adRsnMoveNext : how == 4 ? adRsnMoveLast : adRsnMove;
    if (ev_willmove(r, reason, adStatusOK) == adStatusCancel) {
        raise(-2147217842, SRC_RS, "Operation was canceled.");
        return -1;
    }
    note_save(r);
    if (vb_errpend || flush_edit(r)) return -1;
    v = VW(r);
    int from_deleted = v->cur >= 0 && v->cur < v->d->n && v->d->r[v->cur].del;
    switch (how) {
    case 1: go_first(v); break;
    case 4: { int i = last_alive(v); v->cur = i >= 0 ? i : v->d->n; break; }
    case 3: { int i = v->cur + 1; while (i < v->d->n && v->d->r[i].del) i++; v->cur = i; break; }
    case 2: { int i = v->cur - 1; while (i >= 0 && v->d->r[i].del) i--; v->cur = i; break; }
    case 5: {
        int i = v->cur, k = n;
        while (k > 0 && i < v->d->n) { i++; while (i < v->d->n && v->d->r[i].del) i++; k--; }
        while (k < 0 && i >= 0) { i--; while (i >= 0 && v->d->r[i].del) i--; k++; }
        if (i >= v->d->n) i = v->d->n;
        if (i < 0) i = -1;
        v->cur = i;
        break; }
    }
    if (is_bof(v) || is_eof(v)) ev_endofrecordset(r);
    ev_movecomplete(r, reason);
    note_fill(r);                               /* the bound controls after MoveComplete (VB6 oracle f_ado) */
    if (from_deleted && r->cursorloc == adUseClient) {
        /* the Cursor Engine leaving a deleted record tells the move once more (VB6 oracle f_ado) */
        ev_willmove(r, reason, adStatusOK);
        ev_movecomplete(r, reason);
    }
    return 0;
}

/* Find "Field op value" */
static int like_match(const char *s, const char *p)
{
    for (; *p; p++, s++) {
        if (*p == '*' || *p == '%') {
            while (p[1] == '*' || p[1] == '%') p++;
            if (!p[1]) return 1;
            for (; *s; s++) if (like_match(s, p + 1)) return 1;
            return 0;
        }
        if (!*s) return 0;
        if (*p == '?' || *p == '_') continue;
        if (tolower((unsigned char)*s) != tolower((unsigned char)*p)) return 0;
    }
    return !*s;
}
static int find_m(ars *r, vb_i32 argc, vbvar *argv)
{
    drs *v = VW(r);
    if (argc < 1) { vb_raise(450); return -1; }
    char *crit = strdup(vb_stoutf8(vb_v2s(argv[0])));
    int skip = along(argc, argv, 1, 0), dir = along(argc, argv, 2, 1);
    /* field */
    char *p = crit;
    while (isspace((unsigned char)*p)) p++;
    char fname[256]; int fn = 0;
    if (*p == '[') { p++; while (*p && *p != ']' && fn < 255) fname[fn++] = *p++; if (*p) p++; }
    else while (*p && (isalnum((unsigned char)*p) || *p == '_' || (unsigned char)*p >= 0x80) && fn < 255) fname[fn++] = *p++;
    fname[fn] = 0;
    while (isspace((unsigned char)*p)) p++;
    char op[8] = { 0 };
    if (!strncasecmp(p, "like", 4) && isspace((unsigned char)p[4])) { strcpy(op, "like"); p += 4; }
    else { int k = 0; while (*p && strchr("<>=!", *p) && k < 2) op[k++] = *p++; }
    while (isspace((unsigned char)*p)) p++;
    /* the value */
    vbvar val; char sval[512]; int isstr = 0, isdate = 0;
    if (*p == '\'' || *p == '"') {
        char q = *p++; int k = 0;
        while (*p && k < 511) { if (*p == q) { if (p[1] == q) { sval[k++] = q; p += 2; continue; } p++; break; } sval[k++] = *p++; }
        sval[k] = 0; isstr = 1;
    } else if (*p == '#') {
        p++; int k = 0;
        while (*p && *p != '#' && k < 511) sval[k++] = *p++;
        sval[k] = 0; if (*p) p++; isdate = 1;
    } else {
        int k = 0;
        while (*p && !isspace((unsigned char)*p) && k < 511) sval[k++] = *p++;
        sval[k] = 0;
    }
    while (isspace((unsigned char)*p)) p++;
    if (!fn || !op[0] || *p) {                  /* AND, OR, a second criterion */
        free(crit);
        raise(3001, SRC_RS, MSG_ARGS);
        return -1;
    }
    int col = -1;
    for (int i = 0; i < v->ncols; i++) if (!strcasecmp(v->cols[i].name, fname)) { col = i; break; }
    free(crit);
    if (col < 0) { raise(3265, SRC_RS, MSG_NOTFOUND); return -1; }
    if ((is_bof(v) || is_eof(v)) && missing(argc, argv, 3)) { raise(3021, SRC_RS, MSG_BOFEOF); return -1; }
    if (isstr) val = vb_v_s(vb_sfromutf8(sval));
    else if (isdate) { val = vb_v_s(vb_sfromutf8(sval)); val = vb_v_date(vb_v2date(val)); }
    else if (!strcasecmp(sval, "null")) val = vb_v_null();
    else val = vb_v_d(atof(sval));
    int i = v->cur;
    if (!missing(argc, argv, 3)) {
        double bm = vb_v2d(argv[3]);
        i = (int)bm - 1;
    }
    int step = dir < 0 ? -1 : 1;
    i += skip * step;
    for (; i >= 0 && i < v->d->n; i += step) {
        if (v->d->r[i].del) continue;
        vbvar x = v->d->r[i].v[col];
        int ok = 0;
        if (!strcmp(op, "like")) {
            if (x.vt != VT_NULL && x.vt != VT_EMPTY) ok = like_match(vb_stoutf8(vb_v2s(x)), sval);
        } else if (x.vt == VT_NULL || val.vt == VT_NULL) {
            ok = 0;
        } else {
            int c;
            if (isstr) c = strcasecmp(vb_stoutf8(vb_v2s(x)), sval);
            else { double a = vb_v2d(x), b = vb_v2d(val); c = a < b ? -1 : a > b ? 1 : 0; }
            if (!strcmp(op, "=")) ok = c == 0;
            else if (!strcmp(op, "<>") || !strcmp(op, "!=")) ok = c != 0;
            else if (!strcmp(op, "<")) ok = c < 0;
            else if (!strcmp(op, ">")) ok = c > 0;
            else if (!strcmp(op, "<=")) ok = c <= 0;
            else if (!strcmp(op, ">=")) ok = c >= 0;
        }
        if (vb_errpend) vb_err_clear_pending();
        if (ok) { v->cur = i; return 0; }
    }
    v->cur = step > 0 ? v->d->n : -1;
    return 0;
}

static int set_field(ars *r, int col, vbvar val)
{
    drs *v = VW(r);
    if (!updatable(r)) { raise(3251, SRC_RS, MSG_NOUPDATE); return -1; }
    if (!v->editmode && (is_bof(v) || is_eof(v) || !alive(v, v->cur))) { raise(3021, SRC_FLD, MSG_BOFEOF); return -1; }
    int first = !v->editmode;
    if (first) {
        ev_willchangerecord(r, adRsnFirstChange);
        dao_call(r, r->view, "edit", VBM_METHOD, 0, 0, 1);
        if (vb_errpend) return -1;
    }
    if (r->nhooks) {                            /* the runtime's listeners: the column (the DataRepeater's DataUpdate) */
        vbvar a[1] = { vb_v_l(col) };
        fire((vbobj)r, "lb_willchange", 1, a);
        if (vb_errpend) return -1;
    }
    ev_willchangefield(r);
    vbvar x = vb_v_letval(val);
    if (vb_errpend) return -1;
    if (rs_field_let(v, col, x)) {
        int jet = vb_err_number_pending();
        char *msg = strdup(vb_err_text_pending());
        vb_err_clear_pending();
        if (jet == 3421) raise(-2147352571, "Provider", "Type mismatch.");
        else if (r->memdb) raise(-2147217887, "Microsoft Cursor Engine", MSG_MULTISTEP);
        else if (jet == 3163) provider_raise(cn_of(r), -2147217887, jet, msg);
        else provider_raise(cn_of(r), -2147217887, jet, MSG_MULTISTEP);
        free(msg);
        return -1;
    }
    ev_fieldchangecomplete(r);
    if (first) ev_recordchangecomplete(r, adRsnFirstChange);
    if (first) note_fill(r);                    /* the bound controls: at the first change of the record (VB6 oracle f_dfmt2) */
    return 0;
}

static int update_m(ars *r, vb_i32 argc, vbvar *argv)
{
    drs *v = VW(r);
    if (!updatable(r)) { raise(3251, SRC_RS, MSG_NOUPDATE); return -1; }
    note_save(r);
    if (vb_errpend) return -1;
    if (argc >= 1 && !vb_ismissing(argv[0])) {
        vbvar f = argv[0];
        if (f.vt & VT_ARRAY) {
            vbarr *fa = vb_v_arrptr(f), *va = vb_v_arrptr(argv[1]);
            for (int i = fa->lb[0]; i <= fa->ub[0]; i++) {
                int c = col_of(r, *(vbvar *)vb_aelem1(fa, i), SRC_FLDS);
                if (c < 0 || set_field(r, c, *(vbvar *)vb_aelem1(va, i - fa->lb[0] + va->lb[0]))) return -1;
            }
        } else {
            int c = col_of(r, f, SRC_FLDS);
            if (c < 0 || set_field(r, c, argc > 1 ? argv[1] : vb_v_null())) return -1;
        }
    }
    if (!v->editmode) return 0;
    return flush_edit(r);
}

static int addnew_m(ars *r, vb_i32 argc, vbvar *argv)
{
    if (!updatable(r)) { raise(3251, SRC_RS, MSG_NOUPDATE); return -1; }
    note_save(r);
    if (vb_errpend || flush_edit(r)) return -1;
    ev_willmove(r, adRsnMove, adStatusOK);
    ev_willchangerecord(r, adRsnAddNew);
    dao_call(r, r->view, "addnew", VBM_METHOD, 0, 0, 1);
    if (vb_errpend) return -1;
    ev_recordchangecomplete(r, adRsnAddNew);
    ev_movecomplete(r, adRsnMove);
    note_fill(r);
    if (argc >= 1 && !vb_ismissing(argv[0])) return update_m(r, argc, argv);
    return 0;
}

static vbvar getrows_m(ars *r, vb_i32 argc, vbvar *argv)
{
    drs *v = VW(r);
    int n = along(argc, argv, 0, -1);
    if (!missing(argc, argv, 1)) {
        double bm = vb_v2d(argv[1]);
        v->cur = (int)bm - 1;
    }
    int cols[256], nc = 0;
    if (!missing(argc, argv, 2)) {
        vbvar f = argv[2];
        if (f.vt & VT_ARRAY) {
            vbarr *fa = vb_v_arrptr(f);
            for (int i = fa->lb[0]; i <= fa->ub[0] && nc < 256; i++) {
                int c = col_of(r, *(vbvar *)vb_aelem1(fa, i), SRC_RS);
                if (c < 0) return vb_v_empty();
                cols[nc++] = c;
            }
        } else {
            int c = col_of(r, f, SRC_RS);
            if (c < 0) return vb_v_empty();
            cols[nc++] = c;
        }
    } else for (int i = 0; i < v->ncols && nc < 256; i++) cols[nc++] = i;
    if (is_bof(v) || is_eof(v)) { raise(3021, SRC_RS, MSG_BOFEOF); return vb_v_empty(); }
    int *idx = malloc(sizeof(int) * (size_t)(v->d->n + 1));
    int rows = 0;
    for (int i = v->cur; i < v->d->n && (n < 0 || rows < n); i++) if (!v->d->r[i].del) idx[rows++] = i;
    vbarr a; memset(&a, 0, sizeof a);
    vb_adim(&a, sizeof(vbvar), VBA_VAR, 0, 2, 0, nc - 1, 0, rows - 1);
    for (int k = 0; k < rows; k++)
        for (int c = 0; c < nc; c++) vb_vset((vbvar *)vb_aelem(&a, 2, c, k), ado_value(v->d->r[idx[k]].v[cols[c]]));
    vbvar out = vb_v_arr(&a, VT_VARIANT);
    vb_afree(&a);
    int nx = idx[rows - 1] + 1;
    while (nx < v->d->n && v->d->r[nx].del) nx++;
    v->cur = nx;
    free(idx);
    return out;
}

static vbvar getstring_m(ars *r, vb_i32 argc, vbvar *argv)
{
    drs *v = VW(r);
    int n = along(argc, argv, 1, -1);
    const char *cd = missing(argc, argv, 2) ? "\t" : vb_stoutf8(vb_v2s(argv[2]));
    char *coldel = strdup(cd);
    const char *rd = missing(argc, argv, 3) ? "\r" : vb_stoutf8(vb_v2s(argv[3]));
    char *rowdel = strdup(rd);
    char *nullx = strdup(missing(argc, argv, 4) ? "" : vb_stoutf8(vb_v2s(argv[4])));
    if (is_bof(v) || is_eof(v)) { free(coldel); free(rowdel); free(nullx); raise(3021, SRC_RS, MSG_BOFEOF); return vb_v_empty(); }
    jbuf s = { 0 };
    int rows = 0, i = v->cur;
    for (; i < v->d->n && (n < 0 || rows < n); i++) {
        if (v->d->r[i].del) continue;
        for (int c = 0; c < v->ncols; c++) {
            vbvar x = v->d->r[i].v[c];
            if (x.vt == VT_NULL || x.vt == VT_EMPTY) jb_s(&s, nullx);
            else if (x.vt == VT_BOOL) jb_s(&s, x.u.b ? "-1" : "0");
            else jb_s(&s, vb_stoutf8(vb_v2s(x)));
            jb_s(&s, c < v->ncols - 1 ? coldel : rowdel);
        }
        rows++;
    }
    while (i < v->d->n && v->d->r[i].del) i++;
    v->cur = i;
    vbvar out = vstr(s.p ? s.p : "");
    jb_free(&s);
    free(coldel); free(rowdel); free(nullx);
    return out;
}

static int supports(ars *r, vb_i32 opt)
{
    int scroll = r->cursortype != adOpenForwardOnly, upd = updatable(r);
    if (opt & (0x100 | 0x200 | 0x2000 | 0x4000)) if (!scroll) return 0;
    if (opt & (0x1000400 & ~0x1000000)) if (!upd) return 0;
    if (opt & (0x1000800 & ~0x1000000)) if (!upd) return 0;
    if (opt & (0x1008000 & ~0x1000000)) if (!upd) return 0;
    if (opt & 0x10000) if (r->locktype != adLockBatchOptimistic) return 0;
    if (opt & (0x20000 | 0x400000 | 0x800000)) return 0;
    return 1;
}

static vbvar fields_obj(ars *r)
{
    aflds *f = (aflds *)vb_onew(&vb_cls_ado_fields);
    f->rs = (vbobj)r; vb_oaddref((vbobj)r);
    return retobj((vbobj)f);
}
static vbvar field_obj(ars *r, int col)
{
    afld *f = (afld *)vb_onew(&vb_cls_ado_field);
    f->rs = (vbobj)r; vb_oaddref((vbobj)r);
    f->col = col;
    return retobj((vbobj)f);
}

static vbvar rs_member2(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    ars *r = (ars *)vb_oid(o);
    if (r->chparent) chap_sync(r);
    if (r->pendmove && (!strcmp(name, "bof") || !strcmp(name, "eof") || !strncmp(name, "move", 4) || !strcmp(name, "collect") ||
                        ((!strcmp(name, "fields") || !strcmp(name, "")) && argc > 0))) {
        r->pendmove = 0;
        ev_movecomplete(r, adRsnMove);
    }
    int set = is_set(kind);
    *found = 1;
    /* the members of a closed recordset too */
    if (!strcmp(name, "state")) { is_open(r); return vb_v_l(r->state); }
    if (!strcmp(name, "cursortype")) { if (set) { r->cursortype = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(r->cursortype); }
    if (!strcmp(name, "locktype")) { if (set) { r->locktype = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(r->locktype); }
    if (!strcmp(name, "cursorlocation")) { if (set) { r->cursorloc = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(r->cursorloc); }
    if (!strcmp(name, "maxrecords")) { if (set) { r->maxrecords = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(r->maxrecords); }
    if (!strcmp(name, "cachesize")) { if (set) { r->cachesize = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(r->cachesize); }
    if (!strcmp(name, "pagesize")) { if (set) { r->pagesize = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(r->pagesize ? r->pagesize : 10); }
    if (!strcmp(name, "source")) {
        if (set) {
            vbvar a = argv[argc - 1];
            if (vb_isobject(a)) { acmd *c = (acmd *)vb_oid(vb_v2obj(a)); free(r->source); r->source = sdup(c->text ? c->text : ""); }
            else { free(r->source); r->source = strdup(vb_stoutf8(vb_v2s(a))); }
            return vb_v_empty();
        }
        return vstr(r->source);
    }
    if (!strcmp(name, "activeconnection")) {
        if (set) {
            vbvar a = argv[argc - 1];
            if (vb_isobject(a)) {
                vbobj co = vb_v2obj(a);
                co = co ? vb_oid(co) : 0;
                if (!co) {
                    if (is_open(r) && r->cn) disconnect(r);
                    else if (r->cn) { vb_orelease(r->cn); r->cn = 0; }
                    return vb_v_empty();
                }
                if (co->cls == &vb_cls_ado_connection) { vb_oaddref(co); if (r->cn) vb_orelease(r->cn); r->cn = co; }
            } else {
                acn *c = (acn *)vb_onew(&vb_cls_ado_connection);
                vb_oaddref((vbobj)c);
                free(c->cs); c->cs = strdup(vb_stoutf8(vb_v2s(a)));
                if (cn_open(c)) { vb_orelease((vbobj)c); return vb_v_empty(); }
                if (r->cn) vb_orelease(r->cn);
                r->cn = (vbobj)c;
            }
            return vb_v_empty();
        }
        return retobj(r->cn);
    }
    if (!strcmp(name, "open")) return rs_open_m(r, argc, argv);
    if (!strcmp(name, "fields") || !strcmp(name, "")) {
        if (argc == 0 || (kind == VBM_GET && !strcmp(name, "fields") && argc == 0)) return fields_obj(r);
        if (r->state != 1 && !r->ndefs) { raise(3265, SRC_FLDS, MSG_NOTFOUND); return vb_v_empty(); }
        if (!is_open(r)) { raise(3265, SRC_FLDS, MSG_NOTFOUND); return vb_v_empty(); }
        int c = col_of(r, argv[0], SRC_FLDS);
        if (c < 0) return vb_v_empty();
        if (set) { set_field(r, c, argv[argc - 1]); return vb_v_empty(); }
        return field_obj(r, c);
    }
    if (!strcmp(name, "properties")) {
        alist *l = alist_new(L_PROPS);
        if (r->odbc) odbc_rs_props(r, l);
        vbvar v;
        if (argc == 0) v = retobj((vbobj)l);
        else { int f2; v = alist_member((vbobj)l, "item", kind, argc, argv, &f2); }
        vb_orelease((vbobj)l);
        return v;
    }
    if (!strcmp(name, "supports")) {
        if (need_open(r, SRC_RS)) return vb_v_empty();
        return vb_v_bool(supports(r, along(argc, argv, 0, 0)) ? -1 : 0);
    }
    if (!strcmp(name, "status")) { if (need_open(r, SRC_RS)) return vb_v_empty(); return vb_v_l(r->status); }
    if (!strcmp(name, "filter")) {
        if (set) {
            if (need_open(r, SRC_RS)) return vb_v_empty();
            note_save(r);
            if (vb_errpend || flush_edit(r)) return vb_v_empty();
            vbvar a = vb_v_letval(argv[argc - 1]);
            if (a.vt != VT_BSTR) {
                vb_i32 k = vb_v2l(a);
                vb_vset(&r->filter, vb_v_l(k));
            } else vb_vset(&r->filter, a);
            rebuild_view(r);
            note_fill(r);
            return vb_v_empty();
        }
        return vb_vtmp(r->filter);
    }
    if (!strcmp(name, "sort")) {
        if (set) {
            if (need_open(r, SRC_RS)) return vb_v_empty();
            if (r->cursorloc != adUseClient && !r->memdb) {
                raise(3251, SRC_RS, "Current provider does not support the necessary interfaces for sorting or filtering.");
                return vb_v_empty();
            }
            if (flush_edit(r)) return vb_v_empty();
            free(r->sort); r->sort = strdup(vb_stoutf8(vb_v2s(argv[argc - 1])));
            rebuild_view(r);
            note_fill(r);
            return vb_v_empty();
        }
        return vstr(r->sort);
    }
    if (!strcmp(name, "close")) {
        if (need_open(r, SRC_RS)) return vb_v_empty();
        int pend = r->pdel;
        rs_close_m(r);
        if (pend) raise(3219, SRC_RS, MSG_CONTEXT);  /* a delete that failed was pending (VB6 oracle d_odbc1) */
        return vb_v_empty();
    }
    if (!strcmp(name, "clone")) {
        if (need_open(r, SRC_RS)) return vb_v_empty();
        ars *c = (ars *)rs_new_for(cn_of(r));
        int f;
        vbvar cv = rs_member(r->view, "clone", VBM_METHOD, 0, 0, &f);
        if (dao_failed(cn_of(r), 0)) { vb_orelease((vbobj)c); return vb_v_empty(); }
        vbobj co = vb_v2obj(cv);
        c->cursortype = r->cursortype; c->locktype = r->locktype; c->cursorloc = r->cursorloc;
        free(c->source); c->source = strdup(r->source);
        rs_from_drs(c, co, 0);
        if (r->memdb) { c->memdb = 0; }
        if (r->defs) {
            c->defs = malloc(sizeof(adef) * (size_t)r->ndefs); c->ndefs = r->ndefs;
            for (int i = 0; i < r->ndefs; i++) { c->defs[i] = r->defs[i]; c->defs[i].name = strdup(r->defs[i].name); }
        }
        vbvar v = retobj((vbobj)c);
        vb_orelease((vbobj)c);
        return v;
    }
    /* the members of an open recordset */
    if (need_open(r, SRC_RS)) return vb_v_empty();
    drs *v = VW(r);
    if (!strcmp(name, "bof")) return vb_v_bool(is_bof(v) ? -1 : 0);
    if (!strcmp(name, "eof")) return vb_v_bool(is_eof(v) ? -1 : 0);
    if (!strcmp(name, "recordcount")) return vb_v_l(rcount(r));
    if (!strcmp(name, "editmode")) return vb_v_l(r->pdel ? 4 : v->editmode == 2 ? adEditAdd : v->editmode ? adEditInProgress : adEditNone);
    if (!strcmp(name, "absoluteposition")) {
        if (set) { note_save(r); if (!vb_errpend && !flush_edit(r)) { set_abs_pos(r, vb_v2l(argv[argc - 1])); note_fill(r); } return vb_v_empty(); }
        return vb_v_l(abs_pos(r));
    }
    if (!strcmp(name, "bookmark")) {
        if (set) {
            if (r->cursortype == adOpenForwardOnly) { raise(3251, SRC_RS, "Current Recordset does not support bookmarks. This may be a limitation of the provider or of the selected cursortype."); return vb_v_empty(); }
            int i = (int)vb_v2d(argv[argc - 1]) - 1;
            if (i < 0 || i >= v->d->n) { raise(3001, SRC_RS, MSG_ARGS); return vb_v_empty(); }
            note_save(r);
            if (vb_errpend || flush_edit(r)) return vb_v_empty();
            v->cur = i;
            note_fill(r);
            return vb_v_empty();
        }
        if (r->cursortype == adOpenForwardOnly) { raise(3251, SRC_RS, "Current Recordset does not support bookmarks. This may be a limitation of the provider or of the selected cursortype."); return vb_v_empty(); }
        if (is_bof(v) || is_eof(v)) { raise(3021, SRC_RS, MSG_BOFEOF); return vb_v_empty(); }
        return vb_v_d(v->cur + 1);
    }
    if (!strcmp(name, "pagecount")) {
        int ps = r->pagesize ? r->pagesize : 10, n = rcount(r);
        return vb_v_l(n < 0 ? -1 : (n + ps - 1) / ps);
    }
    if (!strcmp(name, "absolutepage")) {
        int ps = r->pagesize ? r->pagesize : 10;
        if (set) { set_abs_pos(r, (vb_v2l(argv[argc - 1]) - 1) * ps + 1); return vb_v_empty(); }
        int p = abs_pos(r);
        return vb_v_l(p < 0 ? p : (p - 1) / ps + 1);
    }
    if (!strcmp(name, "movenext")) { if (!odbc_pending(r)) do_move(r, 3, 0); return vb_v_empty(); }
    if (!strcmp(name, "moveprevious")) { if (!odbc_pending(r)) do_move(r, 2, 0); return vb_v_empty(); }
    if (!strcmp(name, "movefirst")) { r->pdel = 0; do_move(r, 1, 0); return vb_v_empty(); }
    if (!strcmp(name, "movelast")) { r->pdel = 0; do_move(r, 4, 0); return vb_v_empty(); }
    if (!strcmp(name, "move")) {
        if (odbc_pending(r)) return vb_v_empty();
        int n = along(argc, argv, 0, 0);
        if (!missing(argc, argv, 1)) { double bm = vb_v2d(argv[1]); v->cur = (int)bm - 1; }
        do_move(r, 5, n);
        return vb_v_empty();
    }
    if (!strcmp(name, "find")) {
        note_save(r);
        if (vb_errpend || flush_edit(r)) return vb_v_empty();
        if (!find_m(r, argc, argv)) note_fill(r);
        return vb_v_empty();
    }
    if (!strcmp(name, "addnew")) { if (!odbc_pending(r)) addnew_m(r, argc, argv); return vb_v_empty(); }
    if (!strcmp(name, "update")) { if (!odbc_pending(r)) update_m(r, argc, argv); return vb_v_empty(); }
    if (!strcmp(name, "updatebatch")) { flush_edit(r); return vb_v_empty(); }
    if (!strcmp(name, "cancelupdate") || !strcmp(name, "cancelbatch")) {
        if (v->editmode) { int f; rs_member(r->view, "cancelupdate", VBM_METHOD, 0, 0, &f); if (vb_errpend) vb_err_clear_pending(); }
        note_fill(r);
        return vb_v_empty();
    }
    if (!strcmp(name, "delete")) {
        if (!updatable(r)) { raise(3251, SRC_RS, MSG_NOUPDATE); return vb_v_empty(); }
        if (r->pdel) { odbc_pending(r); return vb_v_empty(); }
        if (v->editmode) { int f; rs_member(r->view, "cancelupdate", VBM_METHOD, 0, 0, &f); if (vb_errpend) vb_err_clear_pending(); }
        if (!alive(v, v->cur)) { raise(3021, SRC_RS, MSG_BOFEOF); return vb_v_empty(); }
        int cur = v->cur;
        if (r->cursorloc == adUseClient) ev_willchangerecord(r, adRsnDelete);
        dao_call(r, r->view, "delete", VBM_METHOD, 0, 0, 1);
        if (vb_errpend) return vb_v_empty();
        v->cur = cur;
        if (r->cursorloc == adUseClient) {      /* the Cursor Engine: Delete, then its Update (VB6 oracle f_ado) */
            ev_recordchangecomplete(r, adRsnDelete);
            ev_willchangerecord(r, adRsnUpdate);
            ev_recordchangecomplete(r, adRsnUpdate);
        } else {
            ev_recordchangecomplete(r, adRsnDelete);
            ev_recordchangecomplete(r, adRsnDelete);
        }
        note_fill(r);
        return vb_v_empty();
    }
    if (!strcmp(name, "requery")) {
        note_save(r);
        if (vb_errpend || flush_edit(r)) return vb_v_empty();
        ev_willmove(r, adRsnRequery, adStatusOK);
        ev_movecomplete(r, adRsnRequery);
        ev_recordsetchangecomplete(r, adRsnClose);
        if (r->odbc && r->osql) {               /* the statement again, on the server */
            acn *c = cn_of(r);
            if (!c || !c->odbc) { raise(3709, SRC_RS, MSG_NOCONN); return vb_v_empty(); }
            int ct = r->cursortype, lt = r->locktype, cl = r->cursorloc, oc = r->ocount;
            char *keep = strdup(r->osql);
            int rows;
            int e = odbc_run(c, keep, r->ovals, r->novals, r, 0, &rows);
            free(keep);
            r->cursortype = ct; r->locktype = lt; r->cursorloc = cl;
            if (oc < 0) r->ocount = -1;
            if (e) return vb_v_empty();
        } else {
            dao_call(r, r->base, "requery", VBM_METHOD, 0, 0, 0);
            if (vb_errpend) return vb_v_empty();
        }
        rebuild_view(r);
        go_first(VW(r));
        ev_opened(r);
        note_fill(r);
        return vb_v_empty();
    }
    if (!strcmp(name, "resync")) return vb_v_empty();
    if (!strcmp(name, "getrows")) return getrows_m(r, argc, argv);
    if (!strcmp(name, "getstring")) return getstring_m(r, argc, argv);
    if (!strcmp(name, "collect")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int c = col_of(r, argv[0], SRC_FLDS);
        if (c < 0) return vb_v_empty();
        if (set) { set_field(r, c, argv[argc - 1]); return vb_v_empty(); }
        vbvar x; if (cur_value(r, c, &x)) return vb_v_empty();
        return vb_vtmp(x);
    }
    if (!strcmp(name, "nextrecordset")) return retobj(0);
    if (!strcmp(name, "cancel")) return vb_v_empty();
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Fields, Field */
static void flds_clear(vbobj o) { aflds *f = (aflds *)o; if (f->rs) vb_orelease(f->rs); }
static vb_i32 flds_count(ars *r)
{
    if (is_open(r)) return VW(r)->ncols;
    return r->ndefs;
}
static vb_i32 flds_each_count(vbobj o) { aflds *f = (aflds *)vb_oid(o); return flds_count(ARS(f->rs)); }
static vbvar flds_each_item(vbobj o, vb_i32 i)
{
    aflds *f = (aflds *)vb_oid(o);
    ars *r = ARS(f->rs);
    if (!is_open(r)) return vb_v_empty();
    return field_obj(r, i);
}
static vbvar flds_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    aflds *f = (aflds *)vb_oid(o);
    ars *r = ARS(f->rs);
    *found = 1;
    if (r->chparent) chap_sync(r);
    if (!strcmp(name, "count")) return vb_v_l(flds_count(r));
    if (!strcmp(name, "item") || !strcmp(name, "")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        if (!is_open(r)) { raise(3265, SRC_FLDS, MSG_NOTFOUND); return vb_v_empty(); }
        int c = col_of(r, argv[0], SRC_FLDS);
        if (c < 0) return vb_v_empty();
        if (is_set(kind)) { set_field(r, c, argv[argc - 1]); return vb_v_empty(); }
        return field_obj(r, c);
    }
    if (!strcmp(name, "append")) {
        if (is_open(r)) { raise(3219, SRC_FLDS, MSG_CONTEXT); return vb_v_empty(); }
        if (argc < 2) { vb_raise(450); return vb_v_empty(); }
        char *nm = vb_stoutf8(vb_v2s(argv[0]));
        for (int i = 0; i < r->ndefs; i++)
            if (!strcasecmp(r->defs[i].name, nm)) { raise(3367, SRC_FLDS, "Cannot append, object already in collection."); return vb_v_empty(); }
        r->defs = realloc(r->defs, sizeof(adef) * (size_t)(r->ndefs + 1));
        adef *d = &r->defs[r->ndefs++];
        d->name = strdup(nm);
        d->type = vb_v2l(argv[1]);
        d->size = along(argc, argv, 2, 0);
        d->attrs = along(argc, argv, 3, 0);
        return vb_v_empty();
    }
    if (!strcmp(name, "delete")) {
        if (is_open(r)) { raise(3219, SRC_FLDS, MSG_CONTEXT); return vb_v_empty(); }
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        for (int i = 0; i < r->ndefs; i++) {
            vbvar k = vb_v_letval(argv[0]);
            int hit = k.vt == VT_BSTR ? !strcasecmp(r->defs[i].name, vb_stoutf8(k.u.s)) : vb_v2l(k) == i;
            if (hit) { free(r->defs[i].name); memmove(&r->defs[i], &r->defs[i + 1], sizeof(adef) * (size_t)(r->ndefs - i - 1)); r->ndefs--; return vb_v_empty(); }
        }
        raise(3265, SRC_FLDS, MSG_NOTFOUND);
        return vb_v_empty();
    }
    if (!strcmp(name, "refresh") || !strcmp(name, "update") || !strcmp(name, "resync") || !strcmp(name, "cancelupdate")) return vb_v_empty();
    *found = 0;
    return vb_v_empty();
}

static void fld_clear(vbobj o) { afld *f = (afld *)o; if (f->rs) vb_orelease(f->rs); }
static vbvar fld_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    afld *f = (afld *)vb_oid(o);
    ars *r = ARS(f->rs);
    *found = 1;
    if (r->chparent) chap_sync(r);
    if (r->pendmove && (!strcmp(name, "value") || !strcmp(name, ""))) { r->pendmove = 0; ev_movecomplete(r, adRsnMove); }
    if (!is_open(r)) { raise(3704, SRC_FLD, MSG_CLOSED); return vb_v_empty(); }
    drs *v = VW(r);
    int c = f->col;
    if (c >= v->ncols) { raise(3265, SRC_FLDS, MSG_NOTFOUND); return vb_v_empty(); }
    if (!strcmp(name, "value") || !strcmp(name, "")) {
        if (is_set(kind)) { set_field(r, c, argv[argc - 1]); return vb_v_empty(); }
        vbvar x;
        if (cur_value(r, c, &x)) return vb_v_empty();
        return vb_vtmp(x);
    }
    if (!strcmp(name, "name")) return vstr(r->memdb && r->defs && c < r->ndefs ? r->defs[c].name : v->cols[c].name);
    if (!strcmp(name, "type")) return vb_v_l(ado_type(r, c));
    if (!strcmp(name, "definedsize")) return vb_v_l(defined_size(r, c));
    if (!strcmp(name, "attributes")) return vb_v_l(field_attrs(r, c));
    if (!strcmp(name, "precision")) return vb_v_byte((vb_u8)((r->odbc || r->exact) && c < r->ndefs ? r->defs[c].prec : precision_of(ado_type(r, c))));
    if (!strcmp(name, "numericscale")) return vb_v_byte((vb_u8)((r->odbc || r->exact) && c < r->ndefs ? r->defs[c].scale : 255));
    if (!strcmp(name, "actualsize")) {
        vbvar x;
        if (cur_value(r, c, &x)) return vb_v_empty();
        if (x.vt == VT_NULL || x.vt == VT_EMPTY) return vb_v_l(-1);
        int t = ado_type(r, c);
        if (t == adVarWChar || t == adLongVarWChar || t == adWChar) return vb_v_l(vb_v2s(x).len * 2);
        if (t == adVarChar || t == adLongVarChar || t == adChar) return vb_v_l(vb_v2s(x).len);
        if (x.vt & VT_ARRAY) { vbarr *a = vb_v_arrptr(x); return vb_v_l(a ? a->n : 0); }
        return vb_v_l(defined_size(r, c));
    }
    if (!strcmp(name, "originalvalue")) {
        if (v->editmode == 2) return vb_v_null();
        if (!alive(v, v->cur)) { raise(3021, SRC_FLD, MSG_BOFEOF); return vb_v_empty(); }
        return vb_vtmp(ado_value(v->d->r[v->cur].v[c]));
    }
    if (!strcmp(name, "underlyingvalue")) {
        if (r->chaps && c < r->ndefs && r->chaps[c]) {     /* a chapter: its recordset, as Value (the Data Form
                                                              * Wizard's Master/Detail: ("ChildCMD").UnderlyingValue) */
            vbvar x;
            if (cur_value(r, c, &x)) return vb_v_empty();
            return vb_vtmp(x);
        }
        if (r->cursorloc != adUseClient && !r->memdb) {
            raise(3251, SRC_FLD, "Current provider does not support refreshing underlying values.");
            return vb_v_empty();
        }
        if (!alive(v, v->cur)) { raise(3021, SRC_FLD, MSG_BOFEOF); return vb_v_empty(); }
        return vb_vtmp(ado_value(v->d->r[v->cur].v[c]));
    }
    if (!strcmp(name, "getchunk")) {
        vbvar x;
        if (cur_value(r, c, &x)) return vb_v_empty();
        if (x.vt == VT_NULL) return vb_v_null();
        vbs s = vb_v2s(x);
        int n = along(argc, argv, 0, s.len);
        return vb_v_s(vb_left(s, n));
    }
    if (!strcmp(name, "appendchunk")) {
        vbvar x;
        if (cur_value(r, c, &x)) return vb_v_empty();
        vbvar add = argc ? argv[0] : vb_v_empty();
        vbvar nv = (x.vt == VT_NULL || x.vt == VT_EMPTY) ? add : vb_v_s(vb_scat(vb_v2s(x), vb_v2s(add)));
        set_field(r, c, nv);
        return vb_v_empty();
    }
    if (!strcmp(name, "properties")) { alist *l = alist_new(L_PROPS); vbvar w = retobj((vbobj)l); vb_orelease((vbobj)l); return w; }
    if (!strcmp(name, "status")) return vb_v_l(0);
    if (!strcmp(name, "datafomat") || !strcmp(name, "dataformat")) return retobj(0);
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Errors, Parameters, Properties */
static vb_i32 alist_each_count(vbobj o) { return ((alist *)vb_oid(o))->n; }
static vbvar alist_each_item(vbobj o, vb_i32 i)
{
    alist *l = (alist *)vb_oid(o);
    return i >= 0 && i < l->n ? retobj(l->items[i]) : vb_v_empty();
}
static const char *list_src(alist *l)
{
    return l->kind == L_ERRORS ? "ADODB.Errors" : l->kind == L_PARAMS ? "ADODB.Parameters" : "ADODB.Properties";
}
static int alist_index(alist *l, vbvar key)
{
    vbvar k = vb_v_letval(key);
    if (k.vt == VT_BSTR) {
        char *nm = vb_stoutf8(k.u.s);
        for (int i = 0; i < l->n; i++) {
            const char *n = l->kind == L_PARAMS ? ((apar *)vb_oid(l->items[i]))->name :
                            l->kind == L_PROPS ? prop_name(l->items[i]) : "";
            if (n && !strcasecmp(n, nm)) return i;
        }
        return -1;
    }
    vb_i32 i = vb_v2l(k);
    return i >= 0 && i < l->n ? i : -1;
}
vbvar alist_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    alist *l = (alist *)vb_oid(o);
    (void)kind;
    *found = 1;
    if (!strcmp(name, "count")) return vb_v_l(l->n);
    if (!strcmp(name, "item") || !strcmp(name, "")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = alist_index(l, argv[0]);
        if (i < 0) { raise(3265, list_src(l), MSG_NOTFOUND); return vb_v_empty(); }
        if (is_set(kind) && l->kind == L_PARAMS) {
            apar *p = (apar *)vb_oid(l->items[i]);
            vb_vset(&p->value, vb_v_letval(argv[argc - 1]));
            return vb_v_empty();
        }
        return retobj(l->items[i]);
    }
    if (!strcmp(name, "clear")) { alist_clear(l); return vb_v_empty(); }
    if (!strcmp(name, "refresh")) {
        if (l->kind == L_PARAMS && l->owner) odbc_refresh_params((acmd *)vb_oid(l->owner));
        return vb_v_empty();
    }
    if (!strcmp(name, "append")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        vbobj p = vb_v2obj(argv[0]);
        if (!p || vb_oid(p)->cls != &vb_cls_ado_parameter) { raise(3001, list_src(l), MSG_ARGS); return vb_v_empty(); }
        alist_add(l, vb_oid(p));
        return vb_v_empty();
    }
    if (!strcmp(name, "delete")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = alist_index(l, argv[0]);
        if (i < 0) { raise(3265, list_src(l), MSG_NOTFOUND); return vb_v_empty(); }
        vb_orelease(l->items[i]);
        memmove(&l->items[i], &l->items[i + 1], sizeof(vbobj) * (size_t)(l->n - i - 1));
        l->n--;
        return vb_v_empty();
    }
    *found = 0;
    return vb_v_empty();
}

static void err_clear(vbobj o) { aerr *e = (aerr *)o; free(e->desc); free(e->src); free(e->state); }
static vbvar err_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    (void)kind; (void)argc; (void)argv;
    aerr *e = (aerr *)vb_oid(o);
    *found = 1;
    if (!strcmp(name, "number")) return vb_v_l(e->number);
    if (!strcmp(name, "description") || !strcmp(name, "")) return vstr(e->desc);
    if (!strcmp(name, "source")) return vstr(e->src);
    if (!strcmp(name, "sqlstate")) return vstr(e->state);
    if (!strcmp(name, "nativeerror")) return vb_v_l(e->native);
    if (!strcmp(name, "helpfile")) return vstr("");
    if (!strcmp(name, "helpcontext")) return vb_v_l(0);
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- Command, Parameter */
/* a value as the type of a Parameter */
static vbvar par_coerce(int type, vbvar v)
{
    v = vb_v_letval(v);
    if (v.vt == VT_NULL || v.vt == VT_EMPTY || vb_errpend) return v;
    vb_i32 save = vb_errpend;
    vbvar r = v;
    switch (type) {
    case adInteger: case adUnsignedInt: r = vb_v_l(vb_v2l(v)); break;
    case adSmallInt: case adUnsignedSmallInt: r = vb_v_i(vb_v2i(v)); break;
    case adTinyInt: case adUnsignedTinyInt: r = vb_v_byte(vb_v2byte(v)); break;
    case adDouble: case adNumeric: case adDecimal: r = vb_v_d(vb_v2d(v)); break;
    case adSingle: r = vb_v_f(vb_v2f(v)); break;
    case adCurrency: r = vb_v_c(vb_v2c(v)); break;
    case adDate: case adDBDate: case adDBTime: case adDBTimeStamp: r = vb_v_date(vb_v2date(v)); break;
    case adBoolean: r = vb_v_bool(vb_v2bool(v) ? -1 : 0); break;
    case adBSTR: case adChar: case adWChar: case adVarChar: case adVarWChar: case adLongVarChar: case adLongVarWChar:
        r = vb_v_s(vb_v2s(v)); break;
    }
    if (vb_errpend && !save) { vb_err_clear_pending(); return v; }
    return r;
}

static void par_clear(vbobj o) { apar *p = (apar *)o; free(p->name); vb_vfree(&p->value); }
static void par_init(vbobj o) { apar *p = (apar *)o; p->name = strdup(""); p->dir = 1; p->type = adEmpty; }
static vbvar par_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    apar *p = (apar *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (!strcmp(name, "value") || !strcmp(name, "")) {
        if (set) { vb_vset(&p->value, par_coerce(p->type, argv[argc - 1])); return vb_v_empty(); }
        return vb_vtmp(p->value);
    }
    if (!strcmp(name, "name")) { if (set) { free(p->name); p->name = strdup(vb_stoutf8(vb_v2s(argv[argc - 1]))); return vb_v_empty(); } return vstr(p->name); }
    if (!strcmp(name, "type")) { if (set) { p->type = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(p->type); }
    if (!strcmp(name, "direction")) { if (set) { p->dir = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(p->dir); }
    if (!strcmp(name, "size")) { if (set) { p->size = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(p->size); }
    if (!strcmp(name, "attributes")) { if (set) { p->attrs = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(p->attrs); }
    if (!strcmp(name, "precision")) { if (set) { p->prec = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_byte((vb_u8)p->prec); }
    if (!strcmp(name, "numericscale")) { if (set) { p->scale = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_byte((vb_u8)p->scale); }
    if (!strcmp(name, "appendchunk")) {
        vbvar add = argc ? argv[0] : vb_v_empty();
        if (p->value.vt == VT_NULL || p->value.vt == VT_EMPTY) vb_vset(&p->value, add);
        else vb_vset(&p->value, vb_v_s(vb_scat(vb_v2s(p->value), vb_v2s(add))));
        return vb_v_empty();
    }
    if (!strcmp(name, "properties")) { alist *l = alist_new(L_PROPS); vbvar w = retobj((vbobj)l); vb_orelease((vbobj)l); return w; }
    *found = 0;
    return vb_v_empty();
}

static void cmd_init(vbobj o)
{
    acmd *c = (acmd *)o;
    c->type = adCmdUnknown; c->timeout = 30;
    c->text = strdup(""); c->name = strdup("");
    c->params = alist_new(L_PARAMS);
    c->params->owner = o;
}
static void cmd_clear(vbobj o)
{
    acmd *c = (acmd *)o;
    if (c->cn) vb_orelease(c->cn);
    free(c->text); free(c->name);
    if (c->params) vb_orelease((vbobj)c->params);
}
static vbvar cmd_execute_run(acmd *c, acn *cn, vb_i32 argc, vbvar *argv);
static vbvar cmd_execute(acmd *c, vb_i32 argc, vbvar *argv)
{
    acn *cn = c->cn ? ACN(c->cn) : 0;
    if (!cn || !cn->state) { raise(3709, SRC_CMD, MSG_NOCONN); return vb_v_empty(); }
    /* Command.Execute tells the connection (VB6 oracle d_de2: the WillExecute of a DataEnvironment) */
    if (cn->h.sinks) {
        char *keep = strdup(c->text ? c->text : "");
        vbobj hold = (vbobj)cn; vb_oaddref(hold);
        ev_willexecute(cn, keep, -1, -1, along(argc, argv, 2, -1), (vbobj)c, 0);
        free(keep);
        vbvar r = vb_errpend ? vb_v_empty() : cmd_execute_run(c, cn, argc, argv);
        ev_executecomplete(cn, -1, !vb_errpend, (vbobj)c, r.vt == VT_DISPATCH ? r.u.o : 0);
        vb_orelease(hold);
        return r;
    }
    return cmd_execute_run(c, cn, argc, argv);
}
static vbvar cmd_execute_run(acmd *c, acn *cn, vb_i32 argc, vbvar *argv)
{
    if (cn->odbc) return odbc_cmd_execute(c, cn, argc, argv);
    /* the values: the array given, or the Parameters */
    if (!missing(argc, argv, 1)) {
        vbvar a = argv[1];
        int n = 0; vbarr *arr = 0;
        if (a.vt & VT_ARRAY) { arr = vb_v_arrptr(a); n = arr->ub[0] - arr->lb[0] + 1; }
        else n = 1;
        for (int i = 0; i < n; i++) {
            vbvar x = arr ? *(vbvar *)vb_aelem1(arr, arr->lb[0] + i) : a;
            if (x.vt == VT_R8 || x.vt == VT_R4 || x.vt == VT_CY) {
                /* the values of the array reach Jet as text, read in the regional settings
                   (VB6 oracle d_ado2: 3.5 is 35 where the decimal point is a comma) */
                char b[64]; snprintf(b, sizeof b, "%.15g", vb_v2d(x));
                vb_i32 save = vb_errpend;
                double dv = vb_v2d(vstr(b));
                if (vb_errpend && !save) vb_err_clear_pending(); else x = vb_v_d(dv);
            }
            if (i >= c->params->n) {
                apar *p = (apar *)vb_onew(&vb_cls_ado_parameter);
                alist_add(c->params, (vbobj)p);
            }
            apar *p = (apar *)vb_oid(c->params->items[i]);
            vb_vset(&p->value, par_coerce(p->type, x));
        }
    }
    int opts = along(argc, argv, 2, -1);
    if (opts == -1) opts = c->type;
    ddb *d = DB(cn->db);
    char *sql, *src;
    if (resolve_cmd(cn, d, c->text, opts, &sql, &src)) return vb_v_empty();
    int nv = c->params->n;
    vbvar *vals = calloc((size_t)(nv ? nv : 1), sizeof(vbvar));
    for (int i = 0; i < nv; i++) vals[i] = ((apar *)vb_oid(c->params->items[i]))->value;
    vbobj o = 0; long n = 0;
    int e = run_sql(cn, d, sql, vals, nv, dbOpenSnapshot, &o, &n);
    free(vals); free(sql);
    if (e) { free(src); return vb_v_empty(); }
    if (argc > 0 && !vb_ismissing(argv[0])) vb_vset(&argv[0], vb_v_l(o ? -1 : (vb_i32)n));
    ars *r = (ars *)rs_new_for(cn);
    if (o) {
        r->cursortype = cn->cursorloc == adUseClient ? adOpenStatic : adOpenForwardOnly;
        r->locktype = adLockReadOnly;
        r->cursorloc = cn->cursorloc;
        rs_from_drs(r, o, 1);
        vb_orelease(o);
        free(r->source); r->source = src; src = 0;
    }
    free(src);
    vbvar v = retobj((vbobj)r);
    vb_orelease((vbobj)r);
    return v;
}
static vbvar cmd_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    acmd *c = (acmd *)vb_oid(o);
    int set = is_set(kind);
    *found = 1;
    if (!strcmp(name, "commandtext") || (!strcmp(name, "") && argc == (set ? 1 : 0))) {
        if (set) { free(c->text); c->text = strdup(vb_stoutf8(vb_v2s(argv[argc - 1]))); return vb_v_empty(); }
        return vstr(c->text);
    }
    if (!strcmp(name, "commandtype")) { if (set) { c->type = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(c->type); }
    if (!strcmp(name, "commandtimeout")) { if (set) { c->timeout = vb_v2l(argv[argc - 1]); return vb_v_empty(); } return vb_v_l(c->timeout); }
    if (!strcmp(name, "prepared")) { if (set) { c->prepared = vb_v2bool(argv[argc - 1]) != 0; return vb_v_empty(); } return vb_v_bool(c->prepared ? -1 : 0); }
    if (!strcmp(name, "name")) { if (set) { free(c->name); c->name = strdup(vb_stoutf8(vb_v2s(argv[argc - 1]))); return vb_v_empty(); } return vstr(c->name); }
    if (!strcmp(name, "state")) return vb_v_l(0);
    if (!strcmp(name, "activeconnection")) {
        if (set) {
            vbvar a = argv[argc - 1];
            if (vb_isobject(a)) {
                vbobj co = vb_v2obj(a);
                co = co ? vb_oid(co) : 0;
                if (c->cn) vb_orelease(c->cn);
                c->cn = co;
                if (co) vb_oaddref(co);
            } else {
                acn *n = (acn *)vb_onew(&vb_cls_ado_connection);
                vb_oaddref((vbobj)n);
                free(n->cs); n->cs = strdup(vb_stoutf8(vb_v2s(a)));
                if (cn_open(n)) { vb_orelease((vbobj)n); return vb_v_empty(); }
                if (c->cn) vb_orelease(c->cn);
                c->cn = (vbobj)n;
            }
            return vb_v_empty();
        }
        return retobj(c->cn);
    }
    if (!strcmp(name, "parameters") || !strcmp(name, "")) {
        if (argc == 0 || (set && argc == 1)) return retobj((vbobj)c->params);
        int f2;
        return alist_member((vbobj)c->params, "item", kind, argc, argv, &f2);
    }
    if (!strcmp(name, "createparameter")) {
        apar *p = (apar *)vb_onew(&vb_cls_ado_parameter);
        if (!missing(argc, argv, 0)) { free(p->name); p->name = strdup(vb_stoutf8(vb_v2s(argv[0]))); }
        p->type = along(argc, argv, 1, adEmpty);
        p->dir = along(argc, argv, 2, 1);
        p->size = along(argc, argv, 3, 0);
        if (!missing(argc, argv, 4)) vb_vset(&p->value, par_coerce(p->type, argv[4]));
        return retobj((vbobj)p);
    }
    if (!strcmp(name, "execute")) return cmd_execute(c, argc, argv);
    if (!strcmp(name, "cancel")) return vb_v_empty();
    if (!strcmp(name, "properties")) { alist *l = alist_new(L_PROPS); vbvar w = retobj((vbobj)l); vb_orelease((vbobj)l); return w; }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- OpenSchema */
static vbvar cn_openschema(acn *c, vb_i32 argc, vbvar *argv)
{
    if (!c->state) { raise(3704, SRC_CN, MSG_CLOSED); return vb_v_empty(); }
    int what = along(argc, argv, 0, -1);
    const char *rtable = 0;
    char rbuf[256] = { 0 };
    if (!missing(argc, argv, 1) && (argv[1].vt & VT_ARRAY)) {
        vbarr *a = vb_v_arrptr(argv[1]);
        int i = a->lb[0] + 2;
        if (i <= a->ub[0]) {
            vbvar x = *(vbvar *)vb_aelem1(a, i);
            if (x.vt == VT_BSTR) { snprintf(rbuf, sizeof rbuf, "%s", vb_stoutf8(x.u.s)); rtable = rbuf; }
        }
    }
    if (c->odbc) return odbc_openschema(c, what, rtable);
    ddb *d = DB(c->db);
    ars *r = (ars *)rs_new_for(0);
    r->cursortype = adOpenStatic; r->locktype = adLockReadOnly; r->cursorloc = adUseClient;
    sqlite3_stmt *st = 0;
    if (what == 20) {                           /* adSchemaTables */
        adef defs[] = { { "TABLE_CATALOG", adVarWChar, 255, 0 }, { "TABLE_SCHEMA", adVarWChar, 255, 0 },
                        { "TABLE_NAME", adVarWChar, 255, 0 }, { "TABLE_TYPE", adVarWChar, 255, 0 },
                        { "TABLE_GUID", adGUID, 16, 0 }, { "DESCRIPTION", adVarWChar, 255, 0 },
                        { "TABLE_PROPID", adInteger, 4, 0 }, { "DATE_CREATED", adDate, 8, 0 },
                        { "DATE_MODIFIED", adDate, 8, 0 } };
        if (make_memory(r, defs, 9)) { vb_orelease((vbobj)r); return vb_v_empty(); }
        ddb *m = DB(r->db);
        sqlite3_stmt *ins = 0;
        sqlite3_prepare_v2(m->db, "INSERT INTO lbrs (TABLE_NAME, TABLE_TYPE) VALUES (?1, ?2)", -1, &ins, 0);
        static const char *const SYS[] = { "MSysACEs", "MSysObjects", "MSysQueries", "MSysRelationships" };
        for (int i = 0; i < 4 && ins; i++) {
            sqlite3_reset(ins);
            sqlite3_bind_text(ins, 1, SYS[i], -1, SQLITE_STATIC);
            sqlite3_bind_text(ins, 2, "SYSTEM TABLE", -1, SQLITE_STATIC);
            sqlite3_step(ins);
        }
        if (sqlite3_prepare_v2(d->db, "SELECT name, 'TABLE' FROM MSysLB_Tables UNION ALL "
                               "SELECT name, CASE WHEN type = 0 AND upper(substr(ltrim(sql), 1, 10)) <> 'PARAMETERS' THEN 'VIEW' ELSE '' END FROM MSysLB_Queries", -1, &st, 0) == SQLITE_OK) {
            while (sqlite3_step(st) == SQLITE_ROW) {
                const char *ty = (const char *)sqlite3_column_text(st, 1);
                if (!ty || !*ty) continue;
                sqlite3_reset(ins);
                sqlite3_bind_text(ins, 1, (const char *)sqlite3_column_text(st, 0), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(ins, 2, ty, -1, SQLITE_TRANSIENT);
                sqlite3_step(ins);
            }
            sqlite3_finalize(st);
        }
        sqlite3_finalize(ins);
        free(r->sort); r->sort = strdup("TABLE_NAME");
    } else if (what == 4) {                     /* adSchemaColumns */
        adef defs[] = { { "TABLE_CATALOG", adVarWChar, 255, 0 }, { "TABLE_SCHEMA", adVarWChar, 255, 0 },
                        { "TABLE_NAME", adVarWChar, 255, 0 }, { "COLUMN_NAME", adVarWChar, 255, 0 },
                        { "COLUMN_GUID", adGUID, 16, 0 }, { "COLUMN_PROPID", adInteger, 4, 0 },
                        { "ORDINAL_POSITION", adInteger, 4, 0 }, { "COLUMN_HASDEFAULT", adBoolean, 2, 0 },
                        { "COLUMN_DEFAULT", adLongVarWChar, 0, 0 }, { "COLUMN_FLAGS", adInteger, 4, 0 },
                        { "IS_NULLABLE", adBoolean, 2, 0 }, { "DATA_TYPE", adSmallInt, 2, 0 },
                        { "TYPE_GUID", adGUID, 16, 0 }, { "CHARACTER_MAXIMUM_LENGTH", adInteger, 4, 0 },
                        { "CHARACTER_OCTET_LENGTH", adInteger, 4, 0 }, { "NUMERIC_PRECISION", adSmallInt, 2, 0 },
                        { "NUMERIC_SCALE", adSmallInt, 2, 0 }, { "DATETIME_PRECISION", adInteger, 4, 0 },
                        { "DESCRIPTION", adVarWChar, 255, 0 } };
        if (make_memory(r, defs, 19)) { vb_orelease((vbobj)r); return vb_v_empty(); }
        ddb *m = DB(r->db);
        sqlite3_stmt *ins = 0;
        sqlite3_prepare_v2(m->db, "INSERT INTO lbrs (TABLE_NAME, COLUMN_NAME, ORDINAL_POSITION, COLUMN_HASDEFAULT, COLUMN_DEFAULT, "
                           "IS_NULLABLE, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, CHARACTER_OCTET_LENGTH, NUMERIC_PRECISION) "
                           "VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10)", -1, &ins, 0);
        if (sqlite3_prepare_v2(d->db, "SELECT tbl, name, ord, type, size, required, defval, attrs FROM MSysLB_Fields "
                               "WHERE ?1 IS NULL OR tbl = ?1", -1, &st, 0) == SQLITE_OK) {
            if (rtable) sqlite3_bind_text(st, 1, rtable, -1, SQLITE_STATIC);
            while (sqlite3_step(st) == SQLITE_ROW) {
                int type = sqlite3_column_int(st, 3), size = sqlite3_column_int(st, 4);
                int dt = type == dbText || type == dbMemo ? adWChar : type == dbBoolean ? adBoolean : type == dbByte ? adUnsignedTinyInt :
                         type == dbInteger ? adSmallInt : type == dbLong ? adInteger : type == dbCurrency ? adCurrency :
                         type == dbSingle ? adSingle : type == dbDouble ? adDouble : type == dbDate ? adDate :
                         type == dbGUID ? adGUID : type == dbDecimal ? adNumeric : adBinary;
                sqlite3_reset(ins);
                sqlite3_bind_text(ins, 1, (const char *)sqlite3_column_text(st, 0), -1, SQLITE_TRANSIENT);
                sqlite3_bind_text(ins, 2, (const char *)sqlite3_column_text(st, 1), -1, SQLITE_TRANSIENT);
                sqlite3_bind_int(ins, 3, sqlite3_column_int(st, 2) + 1);
                const char *dv = (const char *)sqlite3_column_text(st, 6);
                sqlite3_bind_int(ins, 4, dv && *dv ? -1 : 0);
                if (dv && *dv) sqlite3_bind_text(ins, 5, dv, -1, SQLITE_TRANSIENT);
                sqlite3_bind_int(ins, 6, sqlite3_column_int(st, 5) || type == dbBoolean ? 0 : -1);
                sqlite3_bind_int(ins, 7, dt);
                if (type == dbText) { sqlite3_bind_int(ins, 8, size); sqlite3_bind_int(ins, 9, size * 2); }
                if (type == dbMemo) { sqlite3_bind_int(ins, 8, 0); sqlite3_bind_int(ins, 9, 0); }
                int pr = precision_of(dt);
                if (pr != 255) sqlite3_bind_int(ins, 10, pr);
                sqlite3_step(ins);
            }
            sqlite3_finalize(st);
        }
        sqlite3_finalize(ins);
        free(r->sort); r->sort = strdup("TABLE_NAME, COLUMN_NAME");
    } else {
        vb_orelease((vbobj)r);
        raise(3251, SRC_CN, "Object or provider is not capable of performing requested operation.");
        return vb_v_empty();
    }
    int f2;
    rs_member(r->base, "requery", VBM_METHOD, 0, 0, &f2);
    if (vb_errpend) vb_err_clear_pending();
    r->state = 1;
    rebuild_view(r);
    vbvar v = retobj((vbobj)r);
    vb_orelease((vbobj)r);
    return v;
}

/* ---------------------------------------------------------------- the classes */
#define FB(f) 0, 0, 0, f
const vb_class vb_cls_ado_connection = { "Connection", sizeof(acn), cn_init, 0, cn_clear, 0, 0, "connectionstring", FB(cn_member) };
const vb_class vb_cls_ado_recordset = { "Recordset", sizeof(ars), rs_init, 0, rs_clear2, 0, 0, "fields", FB(rs_member2) };
const vb_class vb_cls_ado_fields = { "Fields", sizeof(aflds), 0, 0, flds_clear, 0, 0, "item", flds_each_count, flds_each_item, 0, flds_member };
const vb_class vb_cls_ado_field = { "Field", sizeof(afld), 0, 0, fld_clear, 0, 0, "value", FB(fld_member) };
const vb_class vb_cls_ado_command = { "Command", sizeof(acmd), cmd_init, 0, cmd_clear, 0, 0, "parameters", FB(cmd_member) };
const vb_class vb_cls_ado_parameters = { "Parameters", sizeof(alist), 0, 0, alist_free, 0, 0, "item", alist_each_count, alist_each_item, 0, alist_member };
const vb_class vb_cls_ado_parameter = { "Parameter", sizeof(apar), par_init, 0, par_clear, 0, 0, "value", FB(par_member) };
const vb_class vb_cls_ado_errors = { "Errors", sizeof(alist), 0, 0, alist_free, 0, 0, "item", alist_each_count, alist_each_item, 0, alist_member };
const vb_class vb_cls_ado_error = { "Error", sizeof(aerr), 0, 0, err_clear, 0, 0, "description", FB(err_member) };
const vb_class vb_cls_ado_properties = { "Properties", sizeof(alist), 0, 0, alist_free, 0, 0, "item", alist_each_count, alist_each_item, 0, alist_member };

/* CreateObject("ADODB.Connection") ... */
vbobj ado_createobject(const char *progid)
{
    if (strncmp(progid, "adodb.", 6)) return 0;
    const char *c = progid + 6;
    size_t n = strcspn(c, ".");
    if (n == 10 && !strncmp(c, "connection", 10)) return vb_onew(&vb_cls_ado_connection);
    if (n == 9 && !strncmp(c, "recordset", 9)) return vb_onew(&vb_cls_ado_recordset);
    if (n == 7 && !strncmp(c, "command", 7)) return vb_onew(&vb_cls_ado_command);
    if (n == 9 && !strncmp(c, "parameter", 9)) return vb_onew(&vb_cls_ado_parameter);
    return 0;
}

/* the listeners of the runtime (vb_dbhooks.ado_listen) */
void ado_listen(vbobj o, void (*fn)(void *, vbobj, const char *, int, vbvar *), void *ctx, int on)
{
    o = o ? vb_oid(o) : 0;
    if (!o || o->cls != &vb_cls_ado_recordset) return;
    ars *r = (ars *)o;
    for (int i = 0; i < r->nhooks; i++)
        if (r->hooks[i].fn == fn && r->hooks[i].ctx == ctx) {
            if (!on) { memmove(&r->hooks[i], &r->hooks[i + 1], sizeof *r->hooks * (size_t)(r->nhooks - i - 1)); r->nhooks--; }
            return;
        }
    if (!on) return;
    r->hooks = realloc(r->hooks, sizeof *r->hooks * (size_t)(r->nhooks + 1));
    r->hooks[r->nhooks].fn = fn; r->hooks[r->nhooks].ctx = ctx;
    r->nhooks++;
}
int ado_is_rs(vbobj o) { o = o ? vb_oid(o) : 0; return o && o->cls == &vb_cls_ado_recordset; }

/* ---------------------------------------------------------------- for the bound grids and lists (vb_dbhooks) */
static ars *grid_rs(vbobj o)
{
    o = o ? vb_oid(o) : 0;
    if (!o || o->cls != &vb_cls_ado_recordset) return 0;
    ars *r = (ars *)o;
    return is_open(r) ? r : 0;
}
/* the rows as the recordset has them now (Filter, Sort): their number, the column count, the current one */
int ado_shape(vbobj o, int *ncols, int *cur)
{
    ars *r = grid_rs(o);
    if (!r) { if (ncols) *ncols = 0; if (cur) *cur = -1; return -1; }
    drs *v = VW(r);
    if (ncols) *ncols = v->ncols;
    int k = 0, c = -1;
    for (int i = 0; i < v->d->n; i++) {
        if (v->d->r[i].del) continue;
        if (i == v->cur) c = k;
        k++;
    }
    if (cur) *cur = v->editmode == 2 ? k : c;
    return k;
}
static int nth_row(drs *v, int n)
{
    for (int i = 0, k = 0; i < v->d->n; i++) if (!v->d->r[i].del && k++ == n) return i;
    return -1;
}
vbvar ado_cell(vbobj o, int row, int col)
{
    ars *r = grid_rs(o);
    if (!r) return vb_v_empty();
    drs *v = VW(r);
    if (col < 0 || col >= v->ncols) return vb_v_empty();
    if (v->editmode && v->ebuf && (row == -1 || (v->editmode == 1 && nth_row(v, row) == v->cur) || (v->editmode == 2 && row == nalive(v))))
        return vb_vtmp(ado_value(v->ebuf[col]));
    int i = nth_row(v, row);
    return i < 0 ? vb_v_empty() : vb_vtmp(ado_value(v->d->r[i].v[col]));
}
const char *ado_colname(vbobj o, int col)
{
    ars *r = grid_rs(o);
    if (!r) return "";
    drs *v = VW(r);
    return col >= 0 && col < v->ncols ? v->cols[col].name : "";
}
int ado_coltype(vbobj o, int col)
{
    ars *r = grid_rs(o);
    return r && col >= 0 && col < VW(r)->ncols ? ado_type(r, col) : 0;
}
/* the Bookmark of a row (as ado_shape counts them), Null for none; the row of a Bookmark, -1 for none (the DataRepeater) */
vbvar ado_bookmark(vbobj o, int row)
{
    ars *r = grid_rs(o);
    if (!r || r->cursortype == adOpenForwardOnly) return vb_v_null();
    int i = nth_row(VW(r), row);
    return i < 0 ? vb_v_null() : vb_v_d(i + 1);
}
int ado_bmrow(vbobj o, vbvar bm)
{
    ars *r = grid_rs(o);
    if (!r || bm.vt == VT_NULL || bm.vt == VT_EMPTY || bm.vt == VT_DISPATCH) return -1;
    vb_i32 save = vb_errpend;
    double d = vb_v2d(bm);
    if (vb_errpend && !save) { vb_err_clear_pending(); return -1; }
    drs *v = VW(r);
    int i = (int)d - 1;
    if (d != (double)(i + 1) || i < 0 || i >= v->d->n || v->d->r[i].del) return -1;
    int k = 0;
    for (int j = 0; j < i; j++) if (!v->d->r[j].del) k++;
    return k;
}
/* the user picks a row in a grid or list: a move with its events */
int ado_goto(vbobj o, int row)
{
    ars *r = grid_rs(o);
    if (!r) return -1;
    drs *v = VW(r);
    int i = nth_row(v, row);
    if (i < 0 || (i == v->cur && !v->editmode)) return 0;
    if (ev_willmove(r, adRsnMove, adStatusOK) == adStatusCancel) return -1;
    note_save(r);
    if (vb_errpend || flush_edit(r)) return -1;
    v = VW(r);
    v->cur = i;
    ev_movecomplete(r, adRsnMove);
    note_fill(r);
    return 0;
}

/* ---------------------------------------------------------------- RDO 2.0 (phase 7.8) */
#include "rdo.inc"
