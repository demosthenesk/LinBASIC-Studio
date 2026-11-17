/* bind.c - the Microsoft Data Binding Collection (MSBind): BindingCollection and Binding, written again
 * over the ADO of LinBASIC Studio. A BindingCollection binds properties of objects (controls, objects of
 * classes) to the fields of a data source: an ADO Recordset, an ADO Data control, or an object of a class
 * whose DataSourceBehavior is vbDataSource (its Class_GetDataMember gives the Recordset).
 * The values go into the objects when the current record changes; a property the object marks with
 * PropertyChanged goes back into the field before the record is left (or at once, UpdateMode 1).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "daoint.h"
#include <ctype.h>

extern const vb_class vb_cls_bindingcollection, vb_cls_binding;
extern void ado_listen(vbobj o, void (*fn)(void *, vbobj, const char *, int, vbvar *), void *ctx, int on);
extern int ado_is_rs(vbobj o);

enum { vbUsePropertyAttributes = 0, vbUpdateWhenPropertyChanges = 1, vbUpdateWhenRowChanges = 2 };

typedef struct abind {
    vbobjh h;
    vbobj obj;                          /* the bound object (counted) */
    char *prop, *field, *key;
    vbobj fmt;                          /* DataFormat: a StdDataFormat (counted), or none */
    int changed;
    vbobj bc;                           /* the collection it is in (not counted), 0 when removed */
} abind;

typedef struct abc {
    vbobjh h;
    vbobj source;                       /* what DataSource was set to (counted) */
    vbobj rs;                           /* the ADO Recordset it reads (counted) */
    char *member;
    int mode;
    abind **items; int n, cap;
    int filling;
    struct abc *next, *prev;            /* every collection: PropertyChanged finds its bindings */
} abc;

static abc *all_bcs;
static const char *SRC_BC = "BindingCollection";

/* a member name as late binding takes it: lower case */
static const char *lc(const char *s)
{
    static char b[4][256]; static int k;
    char *o = b[k++ & 3]; int i = 0;
    for (; s[i] && i < 255; i++) o[i] = (char)tolower((unsigned char)s[i]);
    o[i] = 0;
    return o;
}
static vbvar retobj(vbobj o) { return vb_vtmp(vb_v_obj(o)); }
static vbvar vstr(const char *s) { return vb_v_s(vb_sfromutf8(s ? s : "")); }
static void raise(vb_i32 n, const char *msg) { vb_raise_full(n, vb_sfromutf8(SRC_BC), vb_sfromutf8(msg)); }

/* ---------------------------------------------------------------- values */
static int rs_has_record(vbobj rs)
{
    vb_i32 save = vb_errpend;
    int st = vb_v2l(vb_late_o(rs, "state", VBM_GET, 0, 0));
    if (vb_errpend && !save) { vb_err_clear_pending(); return 0; }
    if (st != 1) return 0;
    int em = vb_v2l(vb_late_o(rs, "editmode", VBM_GET, 0, 0));
    int bof = vb_v2bool(vb_late_o(rs, "bof", VBM_GET, 0, 0)) != 0;
    int eof = vb_v2bool(vb_late_o(rs, "eof", VBM_GET, 0, 0)) != 0;
    if (vb_errpend && !save) { vb_err_clear_pending(); return 0; }
    return em == 2 || (!bof && !eof);
}
static int field_get(vbobj rs, const char *field, vbvar *out)
{
    vb_i32 save = vb_errpend;
    vbvar k = vstr(field);
    vbvar f = vb_late_o(rs, "fields", VBM_GET, 1, &k);
    if (vb_errpend && !save) { vb_err_clear_pending(); return -1; }
    *out = vb_late(f, "value", VBM_GET, 0, 0);
    if (vb_errpend && !save) { vb_err_clear_pending(); return -1; }
    return 0;
}

static void fill_one(abc *b, abind *x)
{
    if (!b->rs || !x->obj) return;
    vb_i32 m = vb_mark();
    vbvar v = vb_v_null();
    if (rs_has_record(b->rs) && field_get(b->rs, x->field, &v)) v = vb_v_null();
    b->filling++;
    vb_i32 save = vb_errpend;
    if (x->fmt) {                               /* through the DataFormat (VB6 oracle f_dfmt: "0.000" -> 3,250) */
        v = vb_dfmt_format(x->fmt, v, x->obj);
        if (vb_errpend && !save) vb_err_clear_pending();
    }
    vb_late_o(x->obj, lc(x->prop), VBM_LET, 1, &v);
    if (vb_errpend && !save) {                  /* Null into a String property: "" */
        vb_err_clear_pending();
        if (v.vt == VT_NULL) { vbvar e = vstr(""); vb_late_o(x->obj, lc(x->prop), VBM_LET, 1, &e); if (vb_errpend && !save) vb_err_clear_pending(); }
    }
    b->filling--;
    x->changed = 0;
    vb_release(m);
}
static void fill_all(abc *b)
{
    for (int i = 0; i < b->n; i++) fill_one(b, b->items[i]);
}
/* the changed properties into the current record */
static void save_changed(abc *b, int all)
{
    if (!b->rs || !rs_has_record(b->rs)) return;
    for (int i = 0; i < b->n; i++) {
        abind *x = b->items[i];
        if (!x->changed && !all) continue;
        vb_i32 m = vb_mark();
        vb_i32 save = vb_errpend;
        vbvar v = vb_late_o(x->obj, lc(x->prop), VBM_GET, 0, 0);
        if (!(vb_errpend && !save)) {
            int ok = 1;
            if (x->fmt) {                       /* UnFormat, then the type of the field (regional settings) */
                vbvar u;
                if (vb_dfmt_unformat(x->fmt, v, x->obj, &u)) ok = 0; else v = u;
            }
            if (ok && v.vt == VT_BSTR) {
                vb_i32 s2 = vb_errpend;
                vbvar k = vstr(x->field);
                vbvar f = vb_late_o(b->rs, "fields", VBM_GET, 1, &k);
                int t = vb_errpend && !s2 ? -1 : vb_v2l(vb_late(f, "type", VBM_GET, 0, 0));
                if (vb_errpend && !s2) { vb_err_clear_pending(); t = -1; }
                vbvar w;
                if (t >= 0 && !vb_db_strto(v, t, &w)) v = w;
                else if (t >= 0 && x->fmt) ok = 0;
            }
            if (!ok && !vb_errpend) {           /* the format does not take it: the record is not left */
                vb_release(m);
                fill_all(b);
                vb_raise_full(-2147217842, vb_sfromutf8("ADODB.Recordset"), vb_sfromutf8("Operation was canceled."));
                return;
            }
            vbvar a[2] = { vstr(x->field), v };
            if (!vb_errpend) vb_late_o(b->rs, "fields", VBM_LET, 2, a);
        }
        vb_release(m);
        if (vb_errpend) return;
        x->changed = 0;
    }
}

static void listen_fn(void *ctx, vbobj rs, const char *ev, int argc, vbvar *argv)
{
    (void)rs; (void)argc; (void)argv;
    abc *b = (abc *)ctx;
    if (b->filling) return;                     /* its own change of a field */
    if (!strcmp(ev, "lb_save")) save_changed(b, 0);
    else if (!strcmp(ev, "lb_fill")) fill_all(b);
}

static void set_rs(abc *b, vbobj rs)
{
    if (b->rs) { ado_listen(b->rs, listen_fn, b, 0); vb_orelease(b->rs); }
    b->rs = rs;
    if (rs) { vb_oaddref(rs); ado_listen(rs, listen_fn, b, 1); }
}

/* the Recordset of the data source: a Recordset, a control with one (ADO Data control), an object of a
 * class that is a data source (its GetDataMember) */
static vbobj source_rs(vbobj src, const char *member)
{
    if (!src) return 0;
    if (ado_is_rs(src)) return vb_oid(src);
    vb_i32 save = vb_errpend;
    vbvar a[2] = { vstr(member ? member : ""), vb_v_obj(0) };
    vb_late_o(src, "lb_getdatamember", VBM_METHOD, 2, a);
    if (!(vb_errpend && !save)) {
        vbobj o = a[1].vt == VT_DISPATCH ? a[1].u.o : 0;
        if (o && ado_is_rs(o)) return vb_oid(o);
        return 0;
    }
    vb_err_clear_pending();
    vbvar r = vb_late_o(src, "recordset", VBM_GET, 0, 0);
    if (vb_errpend && !save) { vb_err_clear_pending(); return 0; }
    vbobj o = r.vt == VT_DISPATCH ? r.u.o : 0;
    return o && ado_is_rs(o) ? vb_oid(o) : 0;
}

/* ---------------------------------------------------------------- PropertyChanged */
void bind_prop_changed(vbobj o, const char *name)
{
    o = o ? vb_oid(o) : 0;
    for (abc *b = all_bcs; b; b = b->next) {
        if (b->filling) continue;
        for (int i = 0; i < b->n; i++) {
            abind *x = b->items[i];
            if (vb_oid(x->obj) != o || strcasecmp(x->prop, name)) continue;
            x->changed = 1;
            if (0) {                            /* VB6 writes the value when the record is left, whatever UpdateMode is (oracle d_bind) */
                vb_i32 m = vb_mark();
                vb_i32 save = vb_errpend;
                vbvar v = vb_late_o(x->obj, lc(x->prop), VBM_GET, 0, 0);
                if (!(vb_errpend && !save)) {
                    vbvar a[2] = { vstr(x->field), v };
                    b->filling++;
                    vb_late_o(b->rs, "fields", VBM_LET, 2, a);
                    b->filling--;
                    if (vb_errpend && !save) vb_err_clear_pending();
                    else x->changed = 0;
                } else vb_err_clear_pending();
                vb_release(m);
            }
        }
    }
}

/* ---------------------------------------------------------------- Binding */
static void bind_clear(vbobj o)
{
    abind *x = (abind *)o;
    if (x->obj) vb_orelease(x->obj);
    if (x->fmt) vb_orelease(x->fmt);
    free(x->prop); free(x->field); free(x->key);
}
static vbvar bind_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    abind *x = (abind *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    if (!strcmp(name, "object")) {
        if (set) { vbobj n = vb_v2obj(argv[argc - 1]); if (n) vb_oaddref(n); if (x->obj) vb_orelease(x->obj); x->obj = n; return vb_v_empty(); }
        return retobj(x->obj);
    }
    if (!strcmp(name, "datafield") || !strcmp(name, "")) {
        if (set) { free(x->field); x->field = strdup(vb_stoutf8(vb_v2s(argv[argc - 1]))); if (x->bc) fill_one((abc *)x->bc, x); return vb_v_empty(); }
        return vstr(x->field);
    }
    if (!strcmp(name, "propertyname")) {
        if (set) { free(x->prop); x->prop = strdup(vb_stoutf8(vb_v2s(argv[argc - 1]))); return vb_v_empty(); }
        return vstr(x->prop);
    }
    if (!strcmp(name, "datachanged")) {
        if (set) { x->changed = vb_v2bool(argv[argc - 1]) != 0; return vb_v_empty(); }
        return vb_v_bool(x->changed ? -1 : 0);
    }
    if (!strcmp(name, "key")) {
        if (!x->key || !*x->key) { vb_raise(5); return vb_v_empty(); }   /* a Binding added without a key */
        return vstr(x->key);
    }
    if (!strcmp(name, "dataformat")) {        /* used from the next record on (VB6 oracle f_dfmt) */
        if (set) {
            vbobj n = vb_v2obj(argv[argc - 1]);
            if (vb_errpend) return vb_v_empty();
            if (n && !vb_is_dfmt(n)) { vb_raise(13); return vb_v_empty(); }
            n = n ? vb_oid(n) : 0;
            if (n) vb_oaddref(n);
            if (x->fmt) vb_orelease(x->fmt);
            x->fmt = n;
            return vb_v_empty();
        }
        return retobj(x->fmt);
    }
    *found = 0;
    return vb_v_empty();
}

/* ---------------------------------------------------------------- BindingCollection */
static void bc_init(vbobj o)
{
    abc *b = (abc *)o;
    b->member = strdup("");
    b->mode = vbUpdateWhenPropertyChanges;      /* VB6 oracle d_bind */
    b->next = all_bcs;
    if (all_bcs) all_bcs->prev = b;
    all_bcs = b;
}
static void bc_clear(vbobj o)
{
    abc *b = (abc *)o;
    set_rs(b, 0);
    if (b->source) vb_orelease(b->source);
    for (int i = 0; i < b->n; i++) { b->items[i]->bc = 0; vb_orelease((vbobj)b->items[i]); }
    free(b->items); free(b->member);
    if (b->prev) b->prev->next = b->next; else all_bcs = b->next;
    if (b->next) b->next->prev = b->prev;
}
static int bc_index(abc *b, vbvar key)
{
    vbvar k = vb_v_letval(key);
    if (k.vt == VT_BSTR) {
        const char *s = vb_stoutf8(k.u.s);
        for (int i = 0; i < b->n; i++) if (b->items[i]->key && !strcasecmp(b->items[i]->key, s)) return i;
        return -1;
    }
    vb_i32 i = vb_v2l(k) - 1;                   /* 1-based (VB6 oracle d_bind) */
    return i >= 0 && i < b->n ? i : -1;
}
static void bc_resolve(abc *b)
{
    set_rs(b, source_rs(b->source, b->member));
    fill_all(b);
}
static vb_i32 bc_each_count(vbobj o) { return ((abc *)vb_oid(o))->n; }
static vbvar bc_each_item(vbobj o, vb_i32 i)
{
    abc *b = (abc *)vb_oid(o);
    return i >= 0 && i < b->n ? retobj((vbobj)b->items[i]) : vb_v_empty();
}
static vbvar bc_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    abc *b = (abc *)vb_oid(o);
    int set = kind == VBM_LET || kind == VBM_SET;
    *found = 1;
    if (!strcmp(name, "datasource")) {
        if (set) {
            vbobj n = vb_v2obj(argv[argc - 1]);
            if (vb_errpend) return vb_v_empty();
            if (n) vb_oaddref(n);
            if (b->source) vb_orelease(b->source);
            b->source = n;
            bc_resolve(b);
            return vb_v_empty();
        }
        return retobj(b->source);
    }
    if (!strcmp(name, "datamember")) {
        if (set) { free(b->member); b->member = strdup(vb_stoutf8(vb_v2s(argv[argc - 1]))); if (b->source) bc_resolve(b); return vb_v_empty(); }
        return vstr(b->member);
    }
    if (!strcmp(name, "updatemode")) {
        if (set) { b->mode = vb_v2l(argv[argc - 1]); return vb_v_empty(); }
        return vb_v_l(b->mode);
    }
    if (!strcmp(name, "count")) return vb_v_l(b->n);
    if (!strcmp(name, "item") || !strcmp(name, "")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = bc_index(b, argv[0]);
        if (i < 0) { vb_raise(5); return vb_v_empty(); }
        return retobj((vbobj)b->items[i]);
    }
    if (!strcmp(name, "add")) {
        if (argc < 3) { vb_raise(450); return vb_v_empty(); }
        vbobj obj = vb_v2obj(argv[0]);
        if (vb_errpend) return vb_v_empty();
        if (!obj) { vb_raise(91); return vb_v_empty(); }
        char *prop = strdup(vb_stoutf8(vb_v2s(argv[1])));
        char *field = strdup(vb_stoutf8(vb_v2s(argv[2])));
        char *key = argc > 4 && !vb_ismissing(argv[4]) ? strdup(vb_stoutf8(vb_v2s(argv[4]))) : 0;
        vbobj fmt = 0;                          /* Add(Object, PropertyName, DataField, DataFormat, Key) */
        if (argc > 3 && !vb_ismissing(argv[3])) {
            fmt = vb_v2obj(argv[3]);
            if (vb_errpend) { free(prop); free(field); free(key); return vb_v_empty(); }
            if (fmt && !vb_is_dfmt(fmt)) { free(prop); free(field); free(key); vb_raise(13); return vb_v_empty(); }
        }
        if (key) for (int i = 0; i < b->n; i++)
            if (b->items[i]->key && !strcasecmp(b->items[i]->key, key)) { free(prop); free(field); free(key); vb_raise(457); return vb_v_empty(); }
        {                                       /* the property must exist (438), the field too (42000) */
            vb_i32 save = vb_errpend;
            vb_i32 m = vb_mark();
            vb_late_o(obj, lc(prop), VBM_GET, 0, 0);
            vb_release(m);
            if (vb_errpend && !save) { free(prop); free(field); free(key); return vb_v_empty(); }
            if (b->rs) {
                vbvar fv;
                vb_i32 m2 = vb_mark();
                int bad = 0;
                vbvar k = vstr(field);
                vb_late_o(b->rs, "fields", VBM_GET, 1, &k);
                if (vb_errpend && !save) { vb_err_clear_pending(); bad = 1; }
                (void)fv;
                vb_release(m2);
                if (bad) {
                    free(prop); free(field); free(key);
                    vb_raise_full(42000, vb_sfromutf8("BindingCollection"), vb_sfromutf8("Invalid field name supplied"));
                    return vb_v_empty();
                }
            }
        }
        abind *x = (abind *)vb_onew(&vb_cls_binding);
        vb_oaddref((vbobj)x);
        x->obj = vb_oid(obj); vb_oaddref(x->obj);
        x->prop = prop; x->field = field; x->key = key ? key : strdup("");
        if (fmt) { x->fmt = vb_oid(fmt); vb_oaddref(x->fmt); }
        x->bc = (vbobj)b;
        if (b->n == b->cap) { b->cap = b->cap ? b->cap * 2 : 8; b->items = realloc(b->items, sizeof(abind *) * (size_t)b->cap); }
        b->items[b->n++] = x;
        fill_one(b, x);
        return retobj((vbobj)x);
    }
    if (!strcmp(name, "remove")) {
        if (argc < 1) { vb_raise(450); return vb_v_empty(); }
        int i = bc_index(b, argv[0]);
        if (i < 0) { vb_raise(5); return vb_v_empty(); }
        abind *x = b->items[i];
        memmove(&b->items[i], &b->items[i + 1], sizeof(abind *) * (size_t)(b->n - i - 1));
        b->n--;
        x->bc = 0;
        vb_orelease((vbobj)x);
        return vb_v_empty();
    }
    if (!strcmp(name, "clear")) {
        for (int i = 0; i < b->n; i++) { b->items[i]->bc = 0; vb_orelease((vbobj)b->items[i]); }
        b->n = 0;
        return vb_v_empty();
    }
    if (!strcmp(name, "updatecontrols")) { fill_all(b); return vb_v_empty(); }
    *found = 0;
    return vb_v_empty();
}

#define FB(f) 0, 0, 0, f
const vb_class vb_cls_bindingcollection = { "DBindingCollection", sizeof(abc), bc_init, 0, bc_clear, 0, 0, "item", bc_each_count, bc_each_item, 0, bc_member };
const vb_class vb_cls_binding = { "DBinding", sizeof(abind), 0, 0, bind_clear, 0, 0, "datafield", FB(bind_member) };

vbobj bind_createobject(const char *progid)
{
    if (!strncmp(progid, "msbind.bindingcollection", 24)) return vb_onew(&vb_cls_bindingcollection);
    return 0;
}
