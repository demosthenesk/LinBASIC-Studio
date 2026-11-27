/* addin.c - the host of an add-in of LinBASIC Studio (phase 12.5): an add-in project (an ActiveX DLL with the
 * Connect designer) made as a program, which the IDE starts when the Add-In Manager loads it. VB6 loads the
 * add-in as a COM DLL into its own process; here there is no COM: the program connects to the IDE (TCP on
 * 127.0.0.1, the port and a key in LBADDIN_PORT / LBADDIN_KEY), and the objects of the IDE (VBE, its projects, code
 * modules, command bars ...) are objects of this runtime whose members go to the IDE (late binding), their events
 * coming back to the WithEvents variables by name.
 *
 * One message a line, fields by tabs (\t \n \r \\ escaped), the values as lbdbinfo -server writes them (e, m, n,
 * s:, h:, i:, f:, d:, c:, t:, b:, y:, o:<id>[:<TypeName>] an object of the IDE, o:0 Nothing):
 *   the add-in to the IDE:  K <key>                                the first line
 *                           C <id> <M|G|L|S> <member> <arg>...     a member -> OK <value> [<arg>|-]... (ByRef
 *                                                                  arguments back) | ER <number> <description>
 *                           R <id>                                 an object not needed any more (no answer)
 *   the IDE to the add-in:  X connect <application> <mode> <addininst> | X disconnect <mode> | X startup |
 *                           X update | X shutdown | X event <id> <event> <arg>... | X quit
 *                           -> D <arg>... (the arguments after the event: ByRef handled ...) | ER <number> <description>
 * Copyright (C) 2026 Demosthenes Koptsis.
 * GPL-3.0-or-later WITH the LinBASIC Studio Runtime Library Exception (RUNTIME-EXCEPTION). */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <ctype.h>
#include "vbrt.h"

/* net.c (apart from vbrt.h) */
long long lbnet_open(int udp);
void lbnet_close(long long h);
int lbnet_connect(long long h, const char *ip, int port);
int lbnet_poll(long long h, int *rd, int *wr, int *ex);
int lbnet_send(long long h, const char *b, int n);
int lbnet_recv(long long h, char *b, int n);
int lbnet_wait(long long h, int ms);

static long long g_sock = -1;
static char *g_in; static size_t g_inn, g_incap;           /* bytes read, not yet a line */

/* ---------------------------------------------------------------- the connection */
static void put(const char *s, size_t n)
{
    while (n > 0 && g_sock != -1) {
        int r = lbnet_send(g_sock, s, (int)n);
        if (r == -10035) { lbnet_wait(g_sock, 10); continue; }
        if (r <= 0) { lbnet_close(g_sock); g_sock = -1; return; }
        s += r; n -= (size_t)r;
    }
}

typedef struct { char *p; size_t n, cap; } sbuf;
static void sb_put(sbuf *b, const char *s, size_t n)
{
    if (b->n + n + 1 > b->cap) { b->cap = (b->n + n + 1) * 2; b->p = (char *)realloc(b->p, b->cap); }
    memcpy(b->p + b->n, s, n); b->n += n; b->p[b->n] = 0;
}
static void sb_s(sbuf *b, const char *s) { sb_put(b, s, strlen(s)); }
static void sb_text(sbuf *b, const char *s)
{
    for (; *s; s++) {
        if (*s == '\t') sb_s(b, "\\t");
        else if (*s == '\n') sb_s(b, "\\n");
        else if (*s == '\r') sb_s(b, "\\r");
        else if (*s == '\\') sb_s(b, "\\\\");
        else sb_put(b, s, 1);
    }
}
static void sb_send(sbuf *b) { sb_s(b, "\n"); put(b->p, b->n); free(b->p); b->p = 0; b->n = b->cap = 0; }

/* a line from the IDE (waits for it): the fields, unescaped; 0 at the end of the connection */
static char *g_line;
static int read_line(char **f, int maxf)
{
    for (;;) {
        char *nl = g_in ? memchr(g_in, '\n', g_inn) : 0;
        if (nl) {
            size_t n = (size_t)(nl - g_in);
            free(g_line);
            g_line = (char *)malloc(n + 1);
            memcpy(g_line, g_in, n); g_line[n] = 0;
            memmove(g_in, nl + 1, g_inn - n - 1); g_inn -= n + 1;
            if (n && g_line[n - 1] == '\r') g_line[n - 1] = 0;
            int nf = 0;
            char *p = g_line;
            f[nf++] = p;
            for (; *p && nf < maxf; p++) if (*p == '\t') { *p = 0; f[nf++] = p + 1; }
            for (int i = 0; i < nf; i++) {      /* unescape */
                char *s = f[i], *d = f[i];
                for (; *s; s++) {
                    if (*s == '\\' && s[1]) { s++; *d++ = *s == 't' ? '\t' : *s == 'n' ? '\n' : *s == 'r' ? '\r' : *s; }
                    else *d++ = *s;
                }
                *d = 0;
            }
            return nf;
        }
        if (g_sock == -1) return 0;
        char b[4096];
        int r = lbnet_recv(g_sock, b, sizeof b);
        if (r == -10035) { lbnet_wait(g_sock, 50); continue; }
        if (r <= 0) { lbnet_close(g_sock); g_sock = -1; return 0; }
        if (g_inn + (size_t)r > g_incap) { g_incap = (g_inn + (size_t)r) * 2; g_in = (char *)realloc(g_in, g_incap); }
        memcpy(g_in + g_inn, b, (size_t)r); g_inn += (size_t)r;
    }
}

/* ---------------------------------------------------------------- the objects of the IDE
 * one object of this runtime for each object of the IDE (the same object: the same reference, as Is wants it);
 * a class for each TypeName the IDE gives (TypeName(VBInstance) = "VBE") */
typedef struct { vbobjh h; long id; } remote;
static remote **g_objs; static int g_nobjs;

static vbvar remote_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found);
static void remote_clear(vbobj o);
static vb_i32 remote_count(vbobj o);
static vbvar remote_item(vbobj o, vb_i32 i);

typedef struct { vb_class c; char name[64]; } rclass;
static rclass **g_cls; static int g_ncls;

static const vb_class *class_of(const char *tname)
{
    if (!tname || !*tname) tname = "Object";
    for (int i = 0; i < g_ncls; i++) if (!strcmp(g_cls[i]->name, tname)) return &g_cls[i]->c;
    rclass *c = (rclass *)calloc(1, sizeof *c);
    snprintf(c->name, sizeof c->name, "%s", tname);
    c->c.name = c->name;
    c->c.size = sizeof(remote);
    c->c.clear = remote_clear;
    c->c.each_count = remote_count;
    c->c.each_item = remote_item;
    c->c.fallback = remote_member;
    g_cls = (rclass **)realloc(g_cls, sizeof(rclass *) * (size_t)(g_ncls + 1));
    g_cls[g_ncls++] = c;
    return &c->c;
}

static vbobj remote_get(long id, const char *tname)
{
    if (id <= 0) return 0;
    for (int i = 0; i < g_nobjs; i++) if (g_objs[i] && g_objs[i]->id == id) return (vbobj)g_objs[i];
    remote *r = (remote *)vb_onew(class_of(tname));    /* (a temporary reference: the caller keeps it) */
    r->id = id;
    int k = 0;
    while (k < g_nobjs && g_objs[k]) k++;
    if (k == g_nobjs) { g_objs = (remote **)realloc(g_objs, sizeof(remote *) * (size_t)(g_nobjs + 1)); g_nobjs++; }
    g_objs[k] = r;
    return (vbobj)r;
}

static void remote_clear(vbobj o)
{
    remote *r = (remote *)o;
    for (int i = 0; i < g_nobjs; i++) if (g_objs[i] == r) g_objs[i] = 0;
    if (g_sock != -1 && r->id > 0) {
        char b[48];
        snprintf(b, sizeof b, "R\t%ld\n", r->id);
        put(b, strlen(b));
    }
}

/* ---------------------------------------------------------------- values */
static void sb_value(sbuf *b, vbvar v)
{
    char t[64];
    if (v.vt & 0x4000) v = vb_vtmp(v);         /* (a ByRef Variant: its value) */
    switch (v.vt) {
    case VT_EMPTY: sb_s(b, "e"); break;
    case VT_NULL: sb_s(b, "n"); break;
    case VT_ERROR: sb_s(b, v.u.l == (vb_i32)0x80020004 ? "m" : "e"); break;
    case VT_I2: snprintf(t, sizeof t, "h:%d", v.u.i); sb_s(b, t); break;
    case VT_I4: snprintf(t, sizeof t, "i:%ld", (long)v.u.l); sb_s(b, t); break;
    case VT_R4: snprintf(t, sizeof t, "f:%.9g", v.u.f); sb_s(b, t); break;
    case VT_R8: snprintf(t, sizeof t, "d:%.17g", v.u.d); sb_s(b, t); break;
    case VT_DATE: snprintf(t, sizeof t, "t:%.17g", v.u.d); sb_s(b, t); break;
    case VT_CY: sb_s(b, "c:"); sb_s(b, vb_stoutf8(vb_cstr_c(v.u.cy))); break;
    case VT_BOOL: sb_s(b, v.u.i ? "b:1" : "b:0"); break;
    case VT_UI1: snprintf(t, sizeof t, "y:%d", v.u.b); sb_s(b, t); break;
    case VT_DISPATCH: {
        vbobj o = vb_oid(v.u.o);
        long id = 0;
        if (o && o->cls->fallback == remote_member) id = ((remote *)o)->id;
        snprintf(t, sizeof t, "o:%ld", id);   /* (an object of the add-in itself: the IDE gets Nothing) */
        sb_s(b, t);
        break;
    }
    case VT_BSTR: sb_s(b, "s:"); sb_text(b, vb_stoutf8(v.u.s)); break;
    default:
        sb_s(b, "s:"); sb_text(b, vb_stoutf8(vb_v2s(v)));
        if (vb_errpend) vb_err_clear_pending();
    }
}

static vbvar value_of(const char *a)
{
    char k = a[0];
    const char *v = k && a[1] == ':' ? a + 2 : "";
    switch (k) {
    case 'm': return vb_v_missing();
    case 'n': return vb_v_null();
    case 's': return vb_v_s(vb_sfromutf8(v));
    case 'h': return vb_v_i((vb_i16)atoi(v));
    case 'i': return vb_v_l((vb_i32)atol(v));
    case 'f': return vb_v_f((float)atof(v));
    case 'd': return vb_v_d(atof(v));
    case 'c': return vb_v_c((vb_cy)llround(atof(v) * 10000.0));
    case 't': return vb_v_date(atof(v));
    case 'b': return vb_v_bool(atoi(v) ? -1 : 0);
    case 'y': return vb_v_byte((vb_u8)atoi(v));
    case 'o': {                                 /* o:<id>:<TypeName> */
        long id = atol(v);
        const char *tn = strchr(v, ':');
        return vb_v_obj(remote_get(id, tn ? tn + 1 : "Object"));
    }
    default: return vb_v_empty();
    }
}

static void handle_ide(char **f, int nf);

/* the answer of the IDE to a member: its value, or its error raised here (messages of the IDE that come before it are
 * done first) */
static vbvar wait_answer(vb_i32 argc, vbvar *argv)
{
    char *f[64];
    for (;;) {
        int nf = read_line(f, 64);
        if (nf == 0) { vb_raise_full(440, vb_sfromutf8("LinBASIC Studio"), vb_sfromutf8("The IDE is not there any more")); return vb_v_empty(); }
        if (!strcmp(f[0], "OK")) {
            /* the ByRef arguments the IDE gives back (Find: StartLine ..., ProcOfLine: ProcKind); "-" unchanged */
            for (int i = 2; i < nf && i - 2 < argc; i++)
                if (strcmp(f[i], "-")) argv[i - 2] = value_of(f[i]);
            return nf > 1 ? value_of(f[1]) : vb_v_empty();
        }
        if (!strcmp(f[0], "ER")) {
            vb_raise_full(nf > 1 ? atol(f[1]) : 5, vb_sfromutf8("LinBASIC Studio"), vb_sfromutf8(nf > 2 ? f[2] : ""));
            return vb_v_empty();
        }
        if (!strcmp(f[0], "X")) {               /* (the IDE calls the add-in while it answers it) */
            char *keep = g_line; g_line = 0;
            handle_ide(f, nf);
            free(g_line); g_line = keep;
        }
    }
}

static vbvar call_ide(long id, const char *kind, const char *name, vb_i32 argc, vbvar *argv)
{
    if (g_sock == -1) { vb_raise_full(440, vb_sfromutf8("LinBASIC Studio"), vb_sfromutf8("The IDE is not there any more")); return vb_v_empty(); }
    sbuf b = { 0 };
    char t[48];
    snprintf(t, sizeof t, "C\t%ld\t%s\t", id, kind);
    sb_s(&b, t);
    sb_text(&b, name);
    for (vb_i32 i = 0; i < argc; i++) { sb_s(&b, "\t"); sb_value(&b, argv[i]); }
    sb_send(&b);
    return wait_answer(argc, argv);
}

static vbvar remote_member(vbobj o, const char *name, vb_i32 kind, vb_i32 argc, vbvar *argv, int *found)
{
    *found = 1;
    const char *k = kind == VBM_LET ? "L" : kind == VBM_SET ? "S" : kind == VBM_GET ? "G" : "M";
    return call_ide(((remote *)o)->id, k, name, argc, argv);
}

static vb_i32 remote_count(vbobj o)
{
    vbvar n = call_ide(((remote *)o)->id, "G", "count", 0, 0);
    return vb_errpend ? 0 : vb_v2l(n);
}

static vbvar remote_item(vbobj o, vb_i32 i)
{
    vbvar a = vb_v_l(i + 1);                    /* (the collections of the IDE begin at 1) */
    return call_ide(((remote *)o)->id, "M", "item", 1, &a);
}

/* ---------------------------------------------------------------- what the IDE asks */
static vbobj g_connect;
static int g_quit;

static void answer_error(void)
{
    vb_err_take(0);
    sbuf b = { 0 };
    char t[48];
    snprintf(t, sizeof t, "ER\t%ld\t", (long)vb_err_number());
    sb_s(&b, t);
    sb_text(&b, vb_stoutf8(vb_err_description()));
    sb_send(&b);
    vb_err_clear();
}

static void handle_ide(char **f, int nf)
{
    vb_i32 m = vb_mark();
    vbvar a[60];
    int na = 0;
    const char *op = nf > 1 ? f[1] : "";
    if (!strcmp(op, "connect")) {
        for (int i = 2; i < nf && na < 3; i++) a[na++] = value_of(f[i]);
        while (na < 3) a[na++] = vb_v_empty();
        a[1] = vb_v_l(vb_v2l(a[1]));
        vb_late_o(g_connect, "lbaddin_onconnection", VBM_METHOD, 3, a);
    } else if (!strcmp(op, "disconnect")) {
        a[na++] = vb_v_l(nf > 2 ? vb_v2l(value_of(f[2])) : 1);
        vb_late_o(g_connect, "lbaddin_ondisconnection", VBM_METHOD, 1, a);
    } else if (!strcmp(op, "startup")) {
        vb_late_o(g_connect, "lbaddin_onstartupcomplete", VBM_METHOD, 0, a);
    } else if (!strcmp(op, "update")) {
        vb_late_o(g_connect, "lbaddin_onaddinsupdate", VBM_METHOD, 0, a);
    } else if (!strcmp(op, "shutdown")) {
        vb_late_o(g_connect, "lbaddin_onbeginshutdown", VBM_METHOD, 0, a);
    } else if (!strcmp(op, "event") && nf > 3) {
        vbobj src = remote_get(atol(f[2]), "Object");
        char ev[128];
        snprintf(ev, sizeof ev, "%s", f[3]);
        for (char *q = ev; *q; q++) *q = (char)tolower((unsigned char)*q);
        for (int i = 4; i < nf && na < 60; i++) a[na++] = value_of(f[i]);
        if (src) {
            vb_oaddref(src);
            vb_fire_named(src, ev, na, a);       /* the WithEvents variables of the add-in: by name */
            vb_orelease(src);
        }
    } else if (!strcmp(op, "quit")) {
        g_quit = 1;
    }
    if (vb_errpend) { vb_release(m); answer_error(); return; }
    /* the answer is written, then the temporaries go (their R lines reach the IDE before the answer, so the IDE does
     * not give the id of an object this program let go in its next answer) */
    sbuf b = { 0 };
    sb_s(&b, "D");
    for (int i = 0; i < na; i++) { sb_s(&b, "\t"); sb_value(&b, a[i]); }
    vb_release(m);
    sb_send(&b);
}

/* LbAddIn_HostRun (the start of an add-in made as a program): connect to the IDE, then what it asks, the forms of
 * the add-in working in between, until it unloads the add-in */
void vb_addin_run(vbobj connect)
{
    const char *port = getenv("LBADDIN_PORT"), *key = getenv("LBADDIN_KEY");
    if (!connect || vb_errpend) return;
    if (!port || !*port) {
        vb_raise_full(5, vb_sfromutf8("LinBASIC Studio"),
                      vb_sfromutf8("This program is an add-in of LinBASIC Studio: load it with Add-Ins > Add-In Manager"));
        return;
    }
    g_sock = lbnet_open(0);
    int e = g_sock == -1 ? -1 : lbnet_connect(g_sock, "127.0.0.1", atoi(port));
    for (int i = 0; e && (e == 10035 || e == 10036) && i < 500; i++) {   /* (the connection on its way) */
        int rd, wr, ex;
        lbnet_wait(g_sock, 10);
        lbnet_poll(g_sock, &rd, &wr, &ex);
        if (wr) e = 0; else if (ex) break;
    }
    if (e) { lbnet_close(g_sock); g_sock = -1; return; }
    g_connect = connect;
    vb_oaddref(connect);
    sbuf b = { 0 };
    sb_s(&b, "K\t");
    sb_text(&b, key ? key : "");
    sb_send(&b);
    char *f[64];
    while (!g_quit && g_sock != -1) {
        int rd = 0, wr, ex;
        if (g_inn && g_in && memchr(g_in, '\n', g_inn)) rd = 1;
        else { lbnet_wait(g_sock, 15); lbnet_poll(g_sock, &rd, &wr, &ex); }
        if (rd) {
            int nf = read_line(f, 64);
            if (nf == 0) break;
            if (!strcmp(f[0], "X")) handle_ide(f, nf);
        }
        vb_doevents();                          /* the forms of the add-in */
        if (vb_errpend) vb_err_clear_pending();
    }
    if (g_sock != -1) { lbnet_close(g_sock); g_sock = -1; }
    vb_orelease(connect);
    g_connect = 0;
    vb_end();                                   /* (its forms go with it) */
}
