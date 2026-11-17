/* component.c - the ActiveX components (ActiveX DLL / EXE projects) compiled into a program:
 * each registers its public classes; CreateObject("Proj.Class") and New Proj.Class find them here.
 * The component's Sub Main runs once, before its first object is made (as in VB6).
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include "vbrt_int.h"
#include <ctype.h>
#include <string.h>
#include <stdlib.h>

typedef struct comp {
    char *name;                         /* the project name, lower case */
    const vb_progclass *classes;
    vb_i32 n;
    void (*init)(void);
    int started;
    vbobj *globals;                     /* GlobalMultiUse: the one instance of each class */
} comp;
static comp *comps;
static int ncomps;

void vb_register_component(const char *proj, const vb_progclass *classes, vb_i32 n, void (*init)(void))
{
    comps = realloc(comps, (size_t)(ncomps + 1) * sizeof *comps);
    comp *c = &comps[ncomps++];
    size_t k = strlen(proj);
    c->name = malloc(k + 1);
    for (size_t i = 0; i <= k; i++) c->name[i] = (char)tolower((unsigned char)proj[i]);
    c->classes = classes; c->n = n; c->init = init; c->started = 0;
    c->globals = calloc((size_t)(n ? n : 1), sizeof(vbobj));
}

/* "proj.class" (lower case) -> the component and the index of the class; a bare "class" too */
static int find(const char *low, comp **out)
{
    const char *dot = strchr(low, '.');
    for (int i = 0; i < ncomps; i++) {
        comp *c = &comps[i];
        size_t pl = strlen(c->name);
        const char *cls = low;
        if (dot) {
            if ((size_t)(dot - low) != pl || strncmp(low, c->name, pl)) continue;
            cls = dot + 1;
        }
        for (int j = 0; j < c->n; j++)
            if (!strcmp(c->classes[j].name, cls)) { *out = c; return j; }
    }
    return -1;
}
static void start(comp *c)
{
    if (c->started) return;
    c->started = 1;
    if (c->init) c->init();                     /* Sub Main of the component */
}

/* CreateObject: a creatable class of a component, or 0 if no component has the name */
vbobj vb_component_create(const char *low, int *found)
{
    comp *c; int j = find(low, &c);
    *found = j >= 0;
    if (j < 0) return 0;
    if (!(c->classes[j].flags & VBC_CREATABLE)) { vb_raise(429); return 0; }
    if (!(c->classes[j].flags & VBC_SINGLE)) start(c);   /* SingleUse: a copy of the server of its own */
    if (vb_errpend) return 0;
    return vb_onew(c->classes[j].cls);
}
/* New Proj.Class in a program that uses the component */
vbobj vb_component_new(const char *progid)
{
    char low[256]; int k = 0;
    for (; progid[k] && k < 255; k++) low[k] = (char)tolower((unsigned char)progid[k]);
    low[k] = 0;
    int found;
    vbobj o = vb_component_create(low, &found);
    if (!found && !vb_errpend) vb_raise(429);
    return o;
}
/* the members of a GlobalMultiUse class used without an object: its one instance */
vbobj vb_component_global(const char *progid)
{
    char low[256]; int k = 0;
    for (; progid[k] && k < 255; k++) low[k] = (char)tolower((unsigned char)progid[k]);
    low[k] = 0;
    comp *c; int j = find(low, &c);
    if (j < 0) { vb_raise(429); return 0; }
    if (!c->globals[j]) {
        if (!(c->classes[j].flags & VBC_SINGLE)) start(c);
        if (vb_errpend) return 0;
        c->globals[j] = vb_onew(c->classes[j].cls);
        vb_oaddref(c->globals[j]);              /* its own reference: it lives to the end */
    }
    return c->globals[j];
}
/* Dim x As New Proj.Class: the class, made on first use */
const vb_class *vb_component_class(const char *progid)
{
    char low[256]; int k = 0;
    for (; progid[k] && k < 255; k++) low[k] = (char)tolower((unsigned char)progid[k]);
    low[k] = 0;
    comp *c; int j = find(low, &c);
    if (j < 0) { vb_raise(429); return 0; }
    if (!(c->classes[j].flags & VBC_SINGLE)) start(c);
    return vb_errpend ? 0 : c->classes[j].cls;
}

/* ---------------------------------------------------------------- the copies of an ActiveX EXE
 * A SingleUse object that a client makes starts a copy of the server: the module variables of the
 * server are swapped with a buffer of its own while its code runs. The objects made inside a copy
 * belong to it (o->ext, 0 for the first copy); the first copy is that of the MultiUse objects. */
/* VB6 (oracle c_single): every running copy of the server registers its classes; the class of a
 * SingleUse object serves one object in each copy, so a client's request goes to the first copy
 * whose class has served none yet, and a new copy starts only when every copy has served one.
 * MultiUse objects come from the first copy. A copy's Sub Main runs when it starts. */
typedef struct ctxcopy { void *buf; int started; const vb_class **used; int nused; } ctxcopy;
typedef struct ctxlist { ctxcopy *v; int n; } ctxlist;

static ctxlist *copies(vb_ctxdef *d)
{
    if (!d->priv) {
        ctxlist *l = calloc(1, sizeof *l);
        l->v = calloc(1, sizeof *l->v);         /* the first copy: the globals in place (buf 0) */
        l->n = 1;
        d->priv = l;
    }
    return (ctxlist *)d->priv;
}
static void ctx_switch(vb_ctxdef *d, void *to)
{
    if (d->cur == to) return;
    if (d->cur) d->swap(d->cur);                /* the first copy's globals back in place */
    if (to) d->swap(to);                        /* then those of the other copy */
    d->cur = to;
}
static int used_in(ctxcopy *c, const vb_class *cls)
{
    for (int i = 0; i < c->nused; i++) if (c->used[i] == cls) return 1;
    return 0;
}
void *vb_ctx_begin(vbobj o, vb_ctxdef *d, int single)
{
    void *prev = d->cur;
    if (d->depth > 0) {                         /* New inside the server: the object is of the same copy */
        o->ext = d->cur;
        d->depth++;
        return prev;
    }
    ctxlist *l = copies(d);
    int k = 0;
    if (single) {
        while (k < l->n && used_in(&l->v[k], o->cls)) k++;
        if (k == l->n) {                        /* every copy served one: a new copy of the server */
            l->v = realloc(l->v, (size_t)(l->n + 1) * sizeof *l->v);
            memset(&l->v[l->n], 0, sizeof *l->v);
            l->v[l->n].buf = calloc(1, d->size() ? d->size() : 1);
            l->n++;
        }
        ctxcopy *c = &l->v[k];
        c->used = realloc(c->used, (size_t)(c->nused + 1) * sizeof *c->used);
        c->used[c->nused++] = o->cls;
    }
    ctxcopy *c = &l->v[k];
    o->ext = c->buf;
    ctx_switch(d, c->buf);
    d->depth++;
    if (!c->started) {                          /* the copy starts: its globals, then its Sub Main */
        c->started = 1;
        if (k > 0 && d->init) d->init();        /* (the first copy's were made when the program started) */
        if (d->main && !vb_errpend) d->main();
    }
    return prev;
}
void *vb_ctx_enter(vbobj o, vb_ctxdef *d)
{
    void *prev = d->cur;
    o = vb_oid(o);
    ctx_switch(d, o ? o->ext : 0);
    d->depth++;
    return prev;
}
void vb_ctx_end(vb_ctxdef *d, void *prev)
{
    if (d->depth > 0) d->depth--;
    ctx_switch(d, prev);
}

/* AddressOf: a procedure stands for a number (it fits in a Long); SetTimer and the others call it back */
#define AO_BASE 0x7E000000
static vb_cbfn *aofns;
static int naofns;
vb_i32 vb_addressof(vb_cbfn f)
{
    for (int i = 0; i < naofns; i++) if (aofns[i] == f) return AO_BASE + i;
    aofns = realloc(aofns, (size_t)(naofns + 1) * sizeof *aofns);
    aofns[naofns] = f;
    return AO_BASE + naofns++;
}
vb_cbfn vb_callback_fn(vb_i64 h)
{
    return h >= AO_BASE && h < (vb_i64)AO_BASE + naofns ? aofns[h - AO_BASE] : 0;
}
vb_i32 vb_callback(vb_i64 h, vb_i32 argc, vb_i32 *argv)
{
    vb_cbfn f = vb_callback_fn(h);
    return f ? f(argc, argv) : 0;
}
