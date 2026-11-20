/* object.c - objects of class modules: reference counting, Class_Initialize /
 * Class_Terminate, late binding (Object and Variant), and the VBA Collection.
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"
#include <strings.h>

extern vb_i32 vb_errpend;
struct vb_sink { vbobj listener; vbobj *field; const vb_fn *handlers; vb_sink *next; };

/* ---------------------------------------------------------------- lifetime */
vbobj vb_onew(const vb_class *c)
{
    vbobj o = (vbobj)calloc(1, (size_t)c->size);
    if (!o) { vb_raise(7); return 0; }
    o->cls = c; o->refs = 1;
    vb_otmp(o);                                 /* the expression's reference */
    if (c->init) c->init(o);
    return o;
}

vbobj vb_oid(vbobj o) { return o && o->main ? o->main : o; }

/* the default instance of a class with VB_PredeclaredId (a DataEnvironment): made when first used */
vbobj vb_odefault(vbobj *slot, const vb_class *cls)
{
    if (!*slot) {                               /* the slot has it before its Class_Initialize runs: the */
        vbobj o = (vbobj)calloc(1, (size_t)cls->size);   /* code there can name the default instance */
        if (!o) { vb_raise(7); return 0; }
        o->cls = cls; o->refs = 1;
        vb_otmp(o);
        vb_oset(slot, o);
        if (cls->init) cls->init(o);
    }
    return *slot;
}

vbobj vb_oview(vbobj o, const vb_class *iface)
{
    o = vb_oid(o);
    if (!o || !iface || o->cls == iface) return o;
    for (vbobj v = o->views; v; v = v->views)
        if (v->cls == iface) return v;
    vbobj v = (vbobj)calloc(1, sizeof(vbobjh));
    if (!v) { vb_raise(7); return o; }
    v->cls = iface; v->main = o;
    v->views = o->views; o->views = v;
    return v;
}

void vb_oaddref(vbobj o) { o = vb_oid(o); if (o) o->refs++; }

void vb_orelease(vbobj o)
{
    o = vb_oid(o);
    if (!o || --o->refs > 0) return;
    /* last reference: Class_Terminate, then the fields. The object is kept alive while
     * Terminate runs (it may pass Me around); an error pending in the caller is put
     * aside so that Terminate runs normally, as in VB6. */
    o->refs = 1; o->state = 1;
    if (o->cls->term) {
        vb_i32 save = vb_errpend;
        vb_errpend = 0;
        o->cls->term(o);
        vb_errpend = save;
    }
    if (o->cls->clear) o->cls->clear(o);
    while (o->sinks) { vb_sink *k = o->sinks; o->sinks = k->next; free(k); }
    while (o->views) { vbobj v = o->views; o->views = v->views; free(v); }
    free(o);
}

void vb_oset(vbobj *dst, vbobj src)
{
    vbobj old = *dst;
    vb_oaddref(src);
    *dst = src;
    vb_orelease(old);                           /* last: may run Class_Terminate */
}

void vb_ofree(vbobj *v) { vbobj old = *v; *v = 0; vb_orelease(old); }

vbobj vb_oauto(vbobj *v, const vb_class *c)
{
    if (!*v && c) { vbobj o = vb_onew(c); if (o) vb_oset(v, o); }
    return *v;
}

/* the same class: one pointer, or the copies of one exposed class in two units (a component's
 * interface that the program implements too) */
static int same_class(const vb_class *a, const vb_class *b)
{
    return a == b || (a && b && a->progid && b->progid && !strcmp(a->progid, b->progid));
}

vb_bool vb_implements(const vb_class *cls, const vb_class *iface)
{
    if (same_class(cls, iface)) return -1;
    if (cls->ifaces)
        for (const vb_class *const *p = cls->ifaces; *p; p++) if (same_class(*p, iface)) return -1;
    return 0;
}

/* Set x = o for x As c (c = 0: As Object keeps the reference as it is) */
vbobj (*vb_obj_inner)(vbobj o);                 /* a control of a user control: its object (forms runtime) */

vbobj vb_ocast(vbobj o, const vb_class *c)
{
    if (!o || !c) return o;
    vbobj m = vb_oid(o);
    if (same_class(m->cls, c)) return m;
    if (vb_implements(m->cls, c)) return vb_oview(m, c);
    if (vb_obj_inner) { vbobj in = vb_obj_inner(m); if (in && vb_implements(in->cls, c)) return vb_ocast(in, c); }
    vb_raise(13);
    return 0;
}

void *vb_ofields(vbobj o, void *dummy)
{
    if (!o) { vb_raise(91); return dummy; }
    return vb_oid(o);
}

vb_bool vb_typeof_name(vbobj o, const char *name)
{
    o = vb_oid(o);
    if (!o) return 0;
    if (!strcmp(name, "form")) return o->cls->form ? -1 : 0;
    if (!strcmp(name, "mdiform")) return o->cls->form && ((const vbf_formdef *)o->cls->form)->mdi ? -1 : 0;
    if (!strcmp(name, "control")) return !o->cls->form && o->cls->fallback && strcmp(o->cls->name, "Object") && strcmp(o->cls->name, "StdFont") ? -1 : 0;
    if (!strcmp(name, "picture")) name = "stdpicture";
    if (!strcmp(name, "font")) name = "stdfont";
    return strcasecmp(o->cls->name, name) == 0 ? -1 : 0;
}
vb_bool vb_typeof(vbobj o, const vb_class *c)
{
    o = vb_oid(o);
    if (!o) return 0;
    if (!c || vb_implements(o->cls, c)) return -1;
    vbobj in = vb_obj_inner ? vb_obj_inner(o) : 0;
    return in && vb_implements(in->cls, c) ? -1 : 0;
}

/* ---------------------------------------------------------------- events */

void vb_wev_set(vbobj listener, vbobj *field, vbobj src, const vb_fn *handlers)
{
    src = vb_oid(src);
    vbobj old = *field;
    vb_oaddref(src);
    if (old) {                                  /* disconnect from the old source */
        for (vb_sink **pp = &old->sinks; *pp; pp = &(*pp)->next)
            if ((*pp)->listener == listener && (*pp)->field == field) {
                vb_sink *k = *pp; *pp = k->next; free(k); break;
            }
    }
    *field = src;
    if (src && handlers) {                      /* connect: new listeners go last */
        vb_sink *k = (vb_sink *)calloc(1, sizeof *k);
        if (!k) vb_raise(7);
        else {
            k->listener = listener; k->field = field; k->handlers = handlers;
            vb_sink **pp = &src->sinks;
            while (*pp) pp = &(*pp)->next;
            *pp = k;
        }
    }
    vb_orelease(old);                           /* last: may run Class_Terminate */
}

vb_i32 vb_sinks_begin(vbobj src, vbobj **listeners, const vb_fn ***tables)
{
    vb_i32 n = 0;
    for (vb_sink *k = src ? src->sinks : 0; k; k = k->next) n++;
    *listeners = 0; *tables = 0;
    if (!n) return 0;
    *listeners = (vbobj *)malloc(sizeof(vbobj) * n);
    *tables = (const vb_fn **)malloc(sizeof(vb_fn *) * n);
    if (!*listeners || !*tables) { free(*listeners); free((void *)*tables); vb_raise(7); return 0; }
    vb_i32 i = 0;
    for (vb_sink *k = src->sinks; k; k = k->next, i++) {
        (*listeners)[i] = k->listener; (*tables)[i] = k->handlers;
        vb_oaddref(k->listener);                /* alive while its handler runs */
    }
    return n;
}

void vb_sinks_end(vbobj *listeners, const vb_fn **tables, vb_i32 n)
{
    for (vb_i32 i = 0; i < n; i++) vb_orelease(listeners[i]);
    free(listeners); free((void *)tables);
}

/* RaiseEvent of a form: its listeners (WithEvents f As Form1) have tables of handlers by name ("loaded"), as for
 * the events of controls; an error of a handler stops it and goes back to RaiseEvent */
void vb_fire_named(vbobj src, const char *event, vb_i32 argc, vbvar *argv)
{
    vbobj *ls; const vb_fn **ts;
    vb_i32 n = vb_sinks_begin(vb_oid(src), &ls, &ts);
    for (vb_i32 k = 0; k < n && !vb_errpend; k++)
        for (const vb_member *m = (const vb_member *)ts[k]; m->name; m++)
            if (!strcmp(m->name, event)) {
                vbvar r = m->fn(ls[k], argc, argv);
                vb_vfree(&r);
                break;
            }
    vb_sinks_end(ls, ts, n);
}

vbs vb_otypename(vbobj o) { o = vb_oid(o); return vb_sascii(o ? o->cls->name : "Nothing"); }

/* ---------------------------------------------------------------- Variant */
vbvar vb_v_obj(vbobj o) { vbvar v = vb_v_empty(); v.vt = VT_DISPATCH; v.u.o = o; return v; }

vbobj vb_v2obj(vbvar v)
{
    if (v.vt == VT_DISPATCH) return v.u.o;
    vb_raise(424);                              /* Object required */
    return 0;
}

vb_bool vb_isobject(vbvar v) { return v.vt == VT_DISPATCH ? -1 : 0; }
/* v = w without Set: an object in w gives the value of its default property (VB6) */
vbvar vb_v_letval(vbvar v)
{
    if (v.vt != VT_DISPATCH) return v;
    if (!v.u.o) { vb_raise(91); return vb_v_empty(); }
    return vb_late_o(v.u.o, "", VBM_GET, 0, 0);
}
vb_bool vb_ismissing(vbvar v) { return v.vt == VT_ERROR && v.u.l == (vb_i32)0x80020004 ? -1 : 0; }

/* ---------------------------------------------------------------- late binding */
vbvar vb_late_o(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv)
{
    if (vb_errpend) return vb_v_empty();
    if (!o) { vb_raise(91); return vb_v_empty(); }
    const vb_class *c = o->cls;
    if (!*name && c->defmember) name = c->defmember;
    else if (!*name && !c->fallback) { vb_raise(438); return vb_v_empty(); }
    /* else: the default member of a form or control (Form!Text1 = Form.Controls("Text1")) is the runtime's */
    vb_i32 want = kind == VBM_GET ? (VBM_GET | VBM_METHOD) : kind == VBM_METHOD ? (VBM_METHOD | VBM_GET) : kind;
    for (vb_i32 i = 0; i < c->nmembers; i++) {
        const vb_member *m = &c->members[i];
        if ((m->kind & want) && strcmp(m->name, name) == 0) {
            vb_oaddref(o); vb_otmp(o);          /* keep the object alive during the call */
            return m->fn(o, argc, argv);
        }
    }
    if (c->fallback) {                          /* forms and controls: members of the runtime */
        int found = 0;
        vb_oaddref(o); vb_otmp(o);
        vbvar r = c->fallback(o, name, kind, argc, argv, &found);
        if (found || vb_errpend) return r;
    }
    static int trace = -1;
    if (trace < 0) trace = getenv("VBRT_MONKEY") || getenv("VBRT_TRACE");
    if (trace) fprintf(stderr, "monkey: 438 member %s.%s (%s)\n", c->name ? c->name : "?", name,
                       kind == VBM_LET ? "let" : kind == VBM_SET ? "set" : kind == VBM_METHOD ? "call" : "get");
    vb_raise(438);                              /* Object doesn't support this property or method */
    return vb_v_empty();
}

vbvar vb_late(vbvar target, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv)
{
    if (vb_errpend) return vb_v_empty();
    if (target.vt == VT_RECORD) return vb_rec_member(target, name, kind, argc, argv);   /* v.Qty: a Type */
    if (target.vt != VT_DISPATCH) { vb_raise(424); return vb_v_empty(); }
    return vb_late_o(target.u.o, name, kind, argc, argv);
}

/* For Each over an object (a Collection) */
vb_i32 vb_varr_count(vbvar v);
vbvar vb_varr_item(vbvar v, vb_i32 i);

/* For Each over an object of a class with NewEnum (VB_UserMemId = -4): the loop goes over what
 * NewEnum returns (mCol.[_NewEnum]: the Collection itself) */
void vb_each_begin(vbvar *v)
{
    for (int depth = 0; depth < 8 && !vb_errpend; depth++) {
        if (v->vt != VT_DISPATCH || !v->u.o) return;
        const vb_class *c = v->u.o->cls;
        if (c->each_count && !c->enummember) return;    /* (Form.Controls: a NewEnum of its own, a snapshot) */
        /* For Each c In Form1 / UserControl: its Controls (a form has no NewEnum of its own) */
        const char *en = c->enummember ? c->enummember : c->form ? "controls" : 0;
        if (!en) return;
        vbvar e = vb_late_o(v->u.o, en, VBM_GET, 0, 0);
        if (vb_errpend) return;
        vb_vset(v, e);
    }
}

vb_i32 vb_each_count(vbvar v)
{
    if (v.vt & VT_ARRAY) return vb_varr_count(v);
    if (v.vt != VT_DISPATCH) { vb_raise(424); return 0; }
    if (!v.u.o) { vb_raise(424); return 0; }   /* For Each x In Nothing: VB6 424 (t_ph68) */
    if (!v.u.o->cls->each_count) { vb_raise(438); return 0; }
    return v.u.o->cls->each_count(v.u.o);
}
vbvar vb_each_item(vbvar v, vb_i32 i) { return (v.vt & VT_ARRAY) ? vb_varr_item(v, i) : v.u.o->cls->each_item(v.u.o, i); }

/* ---------------------------------------------------------------- Collection
 * Items in order; keys are compared without regard to case (Option Compare Text). Errors as VB6:
 * 5 key not found, 9 index out of range, 457 duplicate key.
 * As in VB6 the items are a linked list with a hash of the keys, and Item(i) walks from the item reached last:
 * Add with a key, Item(key), Remove(1) and For Each take the same time however many items there are
 * (robot 6.8: 170,000 Add/Remove(1) took minutes when the items were an array searched for the key). */
typedef struct cnode { struct cnode *prev, *next, *hnext; vbvar val; vbs key; unsigned hash; int haskey; } cnode;
typedef struct { vbobjh h; vb_i32 n; cnode *head, *tail; cnode **tab; vb_i32 tcap, nkeys; cnode *cur; vb_i32 curi; } coll;

static unsigned key_hash(vbs k)
{
    if (k.flags & VBS_ODD) {                    /* a byte string: equal only to one with the same bytes, without case */
        unsigned h = 2166136261u;
        for (vb_i32 i = 0; i < k.len; i++) h = (h ^ vb_wlower(k.p[i])) * 16777619u;
        return (h ^ (unsigned)vb_sbyte(k, 2 * k.len) ^ 0x5bd1e995u) * 16777619u;
    }
    return vb_collate_hash(k);
}

static void coll_clear(vbobj o)
{
    coll *c = (coll *)o;
    cnode *x = c->head;
    c->head = c->tail = c->cur = 0; c->n = 0;
    free(c->tab); c->tab = 0; c->tcap = c->nkeys = 0;
    while (x) {
        cnode *nx = x->next;
        vb_vfree(&x->val);                      /* may run Class_Terminate */
        if (x->haskey) vb_sfree(&x->key);
        free(x);
        x = nx;
    }
}

static void tab_put(coll *c, cnode *x)
{
    unsigned i = x->hash & (unsigned)(c->tcap - 1);
    x->hnext = c->tab[i]; c->tab[i] = x;
}
static int tab_grow(coll *c)
{
    vb_i32 nc = c->tcap ? c->tcap * 2 : 64;
    cnode **t = (cnode **)calloc((size_t)nc, sizeof *t);
    if (!t) return -1;
    free(c->tab); c->tab = t; c->tcap = nc;
    for (cnode *x = c->head; x; x = x->next) if (x->haskey) tab_put(c, x);
    return 0;
}
static void tab_del(coll *c, cnode *x)
{
    if (!c->tcap) return;
    cnode **pp = &c->tab[x->hash & (unsigned)(c->tcap - 1)];
    while (*pp && *pp != x) pp = &(*pp)->hnext;
    if (*pp) *pp = x->hnext;
}

static cnode *coll_find(coll *c, vbs key)
{
    if (!c->tcap) return 0;
    unsigned h = key_hash(key);
    for (cnode *x = c->tab[h & (unsigned)(c->tcap - 1)]; x; x = x->hnext)
        if (x->hash == h && vb_scmp(x->key, key, 1) == 0) return x;
    return 0;
}

/* the item at 0-based position i (0 <= i < n): from the item reached last, the first or the last */
static cnode *node_at(coll *c, vb_i32 i)
{
    cnode *x; vb_i32 k;
    vb_i32 dc = c->cur ? (i > c->curi ? i - c->curi : c->curi - i) : 0x7FFFFFFF;
    if (dc <= i && dc <= c->n - 1 - i) { x = c->cur; k = c->curi; }
    else if (i <= c->n - 1 - i) { x = c->head; k = 0; }
    else { x = c->tail; k = c->n - 1; }
    while (k < i) { x = x->next; k++; }
    while (k > i) { x = x->prev; k--; }
    c->cur = x; c->curi = i;
    return x;
}

static int missing(vb_i32 argc, vbvar *argv, int i)
{
    return i >= argc || argv[i].vt == VT_EMPTY || vb_ismissing(argv[i]);
}

/* the item of an Index argument (a number 1..n or a key), or 0 after an error; *pos: its 0-based position,
 * -1 when it was found by its key */
static cnode *coll_node(coll *c, vbvar ix, vb_i32 *pos)
{
    *pos = -1;
    if (ix.vt == VT_DISPATCH) { ix = vb_late(ix, "", VBM_GET, 0, 0); if (vb_errpend) return 0; }
    if (ix.vt == VT_BSTR) {
        cnode *x = coll_find(c, ix.u.s);
        if (!x) vb_raise(5);
        return x;
    }
    vb_i32 i = vb_v2l(ix);
    if (vb_errpend) return 0;
    if (i < 1 || i > c->n) { vb_raise(9); return 0; }
    *pos = i - 1;
    return node_at(c, i - 1);
}

static vbvar coll_add(vbobj o, vb_i32 argc, vbvar *argv)
{
    coll *c = (coll *)o;
    if (argc < 1 || argc > 4) { vb_raise(450); return vb_v_empty(); }
    vbs key = { 0, 0, 0 };
    int haskey = 0;
    if (!missing(argc, argv, 1)) {
        switch (argv[1].vt) {                   /* a key is a String: a number is "Type mismatch" */
        case VT_I2: case VT_I4: case VT_R4: case VT_R8: case VT_CY: case VT_DATE: case VT_BOOL: case VT_UI1:
        case VT_DECIMAL: vb_raise(13); return vb_v_empty();
        }
        key = vb_v2s(argv[1]);
        if (vb_errpend) return vb_v_empty();
        if (coll_find(c, key)) { vb_raise(457); return vb_v_empty(); }
        haskey = 1;                             /* ("" is a key too) */
    }
    int has_before = !missing(argc, argv, 2), has_after = !missing(argc, argv, 3);
    if (has_before && has_after) { vb_raise(5); return vb_v_empty(); }
    cnode *ref = 0;
    vb_i32 refpos = -1;
    if (has_before) { ref = coll_node(c, argv[2], &refpos); if (!ref) return vb_v_empty(); }
    if (has_after) { ref = coll_node(c, argv[3], &refpos); if (!ref) return vb_v_empty(); }
    if (haskey && (c->nkeys + 1) * 2 > c->tcap && tab_grow(c) != 0) { vb_raise(7); return vb_v_empty(); }
    cnode *x = (cnode *)calloc(1, sizeof *x);
    if (!x) { vb_raise(7); return vb_v_empty(); }
    vbvar item = argv[0];
    if (item.vt == VT_ERROR && vb_ismissing(item)) item = vb_v_empty();
    vb_vset(&x->val, item);
    if (haskey) {
        vb_sset(&x->key, key);
        x->haskey = 1; x->hash = key_hash(key);
        tab_put(c, x); c->nkeys++;
    }
    if (!ref || (has_after && ref == c->tail)) {        /* at the end */
        x->prev = c->tail;
        if (c->tail) c->tail->next = x; else c->head = x;
        c->tail = x;
    } else if (has_before) {
        x->next = ref; x->prev = ref->prev;
        if (ref->prev) ref->prev->next = x; else c->head = x;
        ref->prev = x;
    } else {
        x->prev = ref; x->next = ref->next;
        ref->next->prev = x;
        ref->next = x;
    }
    if (ref && c->cur) {                         /* the items after it moved one place on */
        vb_i32 p = refpos < 0 ? -1 : has_before ? refpos : refpos + 1;
        if (p < 0) c->cur = 0;
        else if (c->curi >= p) c->curi++;
    }
    c->n++;
    return vb_v_empty();
}

static vbvar coll_item(vbobj o, vb_i32 argc, vbvar *argv)
{
    coll *c = (coll *)o;
    if (argc != 1) { vb_raise(450); return vb_v_empty(); }
    vb_i32 pos;
    cnode *x = coll_node(c, argv[0], &pos);
    if (!x) return vb_v_empty();
    return vb_vtmp(x->val);
}

static vbvar coll_count(vbobj o, vb_i32 argc, vbvar *argv)
{
    (void)argv;
    if (argc) { vb_raise(451); return vb_v_empty(); }   /* Count(1): a Long has no default member */
    return vb_v_l(((coll *)o)->n);
}

static vbvar coll_remove(vbobj o, vb_i32 argc, vbvar *argv)
{
    coll *c = (coll *)o;
    if (argc != 1) { vb_raise(450); return vb_v_empty(); }
    vb_i32 pos;
    cnode *x = coll_node(c, argv[0], &pos);
    if (!x) return vb_v_empty();
    if (x->prev) x->prev->next = x->next; else c->head = x->next;
    if (x->next) x->next->prev = x->prev; else c->tail = x->prev;
    if (x->haskey) { tab_del(c, x); c->nkeys--; }
    c->n--;
    if (c->cur) {
        if (c->cur == x || pos < 0) c->cur = 0;
        else if (c->curi > pos) c->curi--;
    }
    vbvar old = x->val;
    vbs oldk = x->key;
    int hk = x->haskey;
    free(x);
    vb_vfree(&old);                             /* last: may run Class_Terminate */
    if (hk) vb_sfree(&oldk);
    return vb_v_empty();
}

static vb_i32 coll_each_count(vbobj o) { return ((coll *)o)->n; }
static vbvar coll_each_item(vbobj o, vb_i32 i)
{
    coll *c = (coll *)o;
    if (i < 0 || i >= c->n) return vb_v_empty();
    return vb_vtmp(node_at(c, i)->val);
}

/* [_NewEnum]: the enumerator of a Collection is the Collection (For Each goes over it) */
static vbvar coll_newenum(vbobj o, vb_i32 argc, vbvar *argv)
{
    (void)argv;
    if (argc) { vb_raise(450); return vb_v_empty(); }
    return vb_vtmp(vb_v_obj(o));
}

static const vb_member coll_members[] = {
    { "add", VBM_METHOD, coll_add }, { "count", VBM_GET, coll_count },
    { "item", VBM_METHOD, coll_item }, { "remove", VBM_METHOD, coll_remove },
    { "_newenum", VBM_GET | VBM_METHOD, coll_newenum },
};
const vb_class vb_cls_collection = {
    "Collection", sizeof(coll), 0, 0, coll_clear, coll_members, 5, "item", coll_each_count, coll_each_item, 0
};

/* ---------------------------------------------------------------- DataMembers, DataMemberChanged
 * of a class that is a data source (the Data Form Wizard's class: DataMembers.Add "Primary" in Class_Initialize,
 * DataMemberChanged "Primary" after a Requery; DEV-9 df2 clsAuthors) */
void (*vb_ds_changed_hook)(vbobj source, const char *member);
static struct dsm { vbobj owner, coll; } *dsm_tab;
static int dsm_n, dsm_cap;

vbvar vb_ds_datamembers(vbobj self, vb_i32 argc, vbvar *argv)
{
    (void)argc; (void)argv;
    self = vb_oid(self);
    for (int i = 0; i < dsm_n; i++)
        if (dsm_tab[i].owner == self) return vb_vtmp(vb_v_obj(dsm_tab[i].coll));
    if (dsm_n == dsm_cap) {
        dsm_cap = dsm_cap ? dsm_cap * 2 : 8;
        dsm_tab = (struct dsm *)realloc(dsm_tab, (size_t)dsm_cap * sizeof *dsm_tab);
    }
    vbobj c = vb_onew(&vb_cls_collection);      /* kept for the object (its names stay while it lives) */
    dsm_tab[dsm_n].owner = self;
    dsm_tab[dsm_n].coll = c;
    dsm_n++;
    return vb_vtmp(vb_v_obj(c));
}

vbvar vb_ds_datamemberchanged(vbobj self, vb_i32 argc, vbvar *argv)
{
    char name[256] = "";
    if (argc > 0) {
        const char *s = vb_stoutf8(vb_v2s(argv[0]));
        snprintf(name, sizeof name, "%s", s ? s : "");
    }
    if (vb_ds_changed_hook) vb_ds_changed_hook(vb_oid(self), name);
    return vb_v_empty();
}
